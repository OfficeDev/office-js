import os
import json
import requests
from azure.identity import DefaultAzureCredential
from langchain_openai import AzureChatOpenAI
from langchain_core.messages import SystemMessage, HumanMessage
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import JsonOutputParser

class AzureChatOpenAIError(Exception):
    """Exception raised for errors in Azure OpenAI chat operations."""
    pass

def get_azure_ad_token():
    """Returns a function that gets an Azure AD token."""
    credential = DefaultAzureCredential()
    # This returns the token when called by the LangChain internals
    return lambda: credential.get_token("https://cognitiveservices.azure.com/.default").token

def get_azure_chat_model(model_id="gpt-4o"):
    """Get a configured Azure OpenAI chat model through LangChain."""
    try:
        # Create Azure OpenAI chat model with Azure AD authentication
        chat_model = AzureChatOpenAI(
            deployment_name=model_id,
            api_version="2025-01-01-preview",
            azure_endpoint="https://officegithubcopilotextsubdomain.openai.azure.com/",
            azure_ad_token_provider=get_azure_ad_token(),
            temperature=0,  # Use deterministic output for analysis
            model_kwargs={"response_format": {"type": "json_object"}}  # Ensure JSON output
        )
        return chat_model
    except Exception as e:
        raise AzureChatOpenAIError(e) from e

def ensure_label_exists(repo_owner, repo_name, github_token, label_name, color, description=None):
    """
    Ensure a label exists with the specified color and description.
    If the label doesn't exist, create it. If it exists but with different settings, update it.
    """
    # Check if label exists
    url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/labels/{label_name}"
    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json"
    }
    
    response = requests.get(url, headers=headers)
    
    # If label exists
    if response.status_code == 200:
        current_label = response.json()
        # Check if we need to update
        if current_label["color"] != color or (description and current_label.get("description") != description):
            print(f"Updating label '{label_name}' with new color: {color}")
            update_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/labels/{label_name}"
            data = {"color": color}
            if description:
                data["description"] = description
            
            update_response = requests.patch(update_url, headers=headers, json=data)
            return update_response.status_code == 200
        return True
    
    # If label doesn't exist, create it
    elif response.status_code == 404:
        print(f"Creating label '{label_name}' with color: {color}")
        create_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/labels"
        data = {
            "name": label_name,
            "color": color
        }
        if description:
            data["description"] = description
            
        create_response = requests.post(create_url, headers=headers, json=data)
        return create_response.status_code == 201
    
    # Some other error
    else:
        print(f"Error checking label: {response.status_code}")
        print(response.text)
        return False

def is_outlook_exclusive_issue(issue_title, issue_body, chat_model):
    """
    Use LLM to determine if an issue is specifically related to Outlook and exclusive to it.
    
    Args:
        issue_title: The title of the issue
        issue_body: The body/description of the issue
        chat_model: LangChain model for analysis
        
    Returns:
        bool: True if the issue is Outlook-exclusive, False otherwise
    """
    # Prepare issue content
    issue_content = f"Issue Title: {issue_title}\n\nIssue Description: {issue_body}"
    
    # Create a specific prompt for Outlook exclusivity detection
    outlook_prompt = ChatPromptTemplate.from_messages([
        ("system", """You are an expert at analyzing software issues. 
        Your task is to determine if an issue is exclusively related to Microsoft Outlook.
        An 'Outlook-exclusive issue' is a problem that:
        1. Only occurs in Microsoft Outlook (desktop, web, or mobile)
        2. Is specifically related to Outlook functionality or features
        3. Not related to other Office 365 applications or services, like Word, Excel, etc.
        
        Provide a JSON response with your analysis."""),
        ("human", """Analyze the following issue and determine if it's exclusively an Outlook-related issue.
        Response must be JSON with:
        - 'is_outlook_exclusive': boolean
        - 'confidence': number between 0 and 1
        - 'reason': string explaining your decision
        
        Issue content:
        {issue}""")
    ])
    
    try:
        # Define parser for the expected output format
        parser = JsonOutputParser()
        
        # Execute the analysis
        result = outlook_prompt.pipe(chat_model).pipe(parser).invoke({"issue": issue_content})
        
        # Print result for debugging
        print(f"Outlook exclusivity analysis result:")
        print(json.dumps(result, ensure_ascii=False, indent=2))
        
        # Return true if the model identified this as an Outlook-exclusive issue with reasonable confidence
        return result.get("is_outlook_exclusive", False) and result.get("confidence", 0) > 0.7
        
    except Exception as e:
        print(f"Error in Outlook exclusivity detection: {str(e)}")
        # Fall back to false in case of errors
        return False

def get_issue_comments(repo_owner, repo_name, issue_number, github_token):
    """Get all comments for a specific issue."""
    comments_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/issues/{issue_number}/comments"
    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json"
    }
    
    response = requests.get(comments_url, headers=headers)
    
    if response.status_code == 200:
        comments = response.json()
        print(f"Retrieved {len(comments)} comments for issue #{issue_number}")
        return comments
    else:
        print(f"Error fetching comments: {response.status_code}")
        print(response.text)
        return []

def analyze_issue_for_regression(issue_content, issue_number, chat_model):
    """
    Analyze an issue for regression with dual-LLM verification.
    
    This function performs two independent analyses:
    1. Initial analysis to determine if the issue is a regression
    2. Verification analysis to confirm or challenge the initial result
    
    Args:
        issue_content: The issue content (title, body, and optionally comments)
        issue_number: Issue number for logging
        chat_model: The primary LLM model instance
        
    Returns:
        dict: Analysis result with final decision and reasoning
    """
    print(f"Starting dual-LLM regression analysis for issue #{issue_number}")
    
    try:
        # Step 1: Initial Analysis
        initial_prompt = ChatPromptTemplate.from_messages([
            ("system", """You are an expert at analyzing software issues. 
            Your task is to determine if an issue is a regression. 
            A regression is a bug where functionality that previously worked no longer works due to a recent change.
            Analyze the issue carefully and provide a detailed explanation.
            Provide a JSON response."""),
            ("human", """Analyze the following issue and determine if it's a regression issue.
            Response must be JSON with:
            - 'is_regression': boolean
            - 'confidence': number between 0 and 1
            - 'reason': string explaining your decision
            
            Issue content:
            {issue}""")
        ])
        
        parser = JsonOutputParser()
        initial_result = initial_prompt.pipe(chat_model).pipe(parser).invoke({"issue": issue_content})
        
        print(f"Initial analysis result for issue #{issue_number}:")
        print(json.dumps(initial_result, ensure_ascii=False, indent=2))
        
        # Step 2: Verification with a second independent LLM instance
        # Create a new model instance to ensure no context sharing
        verification_model = get_azure_chat_model(model_id="gpt-4o")
        
        verification_prompt = ChatPromptTemplate.from_messages([
            ("system", """You are an expert at verifying software regression issues.
            Your task is to provide a second opinion on whether an issue is a regression or not.
            
            A regression is a bug where functionality that previously worked properly no longer works due to a recent change.
            
            Your analysis must be completely independent. Be skeptical and thorough.
            Classify the issue into one of three categories:
            1. Confirmed regression: You're confident this is a regression bug
            2. Confirmed not regression: You're confident this is NOT a regression bug
            3. Uncertain: You cannot determine with confidence
            """),
            ("human", """Review this issue and determine if it's a regression issue.
            
            Response must be JSON with:
            - 'verification_result': string, one of ["confirmed_regression", "confirmed_not_regression", "uncertain"]
            - 'confidence': number between 0 and 1
            - 'explanation': string explaining your reasoning
            
            Issue content:
            {issue}""")
        ])
        
        verification_result = verification_prompt.pipe(verification_model).pipe(parser).invoke({"issue": issue_content})
        
        print(f"Verification result for issue #{issue_number}:")
        print(json.dumps(verification_result, ensure_ascii=False, indent=2))
        
        # Step 3: Determine final result based on both analyses
        is_regression = initial_result.get("is_regression", False)
        initial_confidence = initial_result.get("confidence", 0.5)
        verification_status = verification_result.get("verification_result", "uncertain")
        verification_confidence = verification_result.get("confidence", 0.5)
        
        # Final decision logic
        if verification_status == "confirmed_regression" and is_regression:
            # Both agree it's a regression
            final_decision = "confirmed_regression"
            confidence_level = (initial_confidence + verification_confidence) / 2
            final_reason = f"Initial analysis: {initial_result.get('reason', 'N/A')}\n\nVerification: {verification_result.get('explanation', 'N/A')}"
        elif verification_status == "confirmed_not_regression" and not is_regression:
            # Both agree it's not a regression
            final_decision = "confirmed_not_regression"
            confidence_level = (initial_confidence + verification_confidence) / 2
            final_reason = f"Initial analysis: {initial_result.get('reason', 'N/A')}\n\nVerification: {verification_result.get('explanation', 'N/A')}"
        elif verification_status == "confirmed_regression" and not is_regression:
            # Disagreement, but verification is confident it's a regression
            final_decision = "confirmed_regression" if verification_confidence > 0.8 else "uncertain"
            confidence_level = verification_confidence
            final_reason = f"Verification confirmed this is a regression.\n\nVerification: {verification_result.get('explanation', 'N/A')}"
        elif verification_status == "confirmed_not_regression" and is_regression:
            # Disagreement, but verification is confident it's not a regression
            final_decision = "confirmed_not_regression" if verification_confidence > 0.8 else "uncertain"
            confidence_level = verification_confidence
            final_reason = f"Verification confirmed this is not a regression.\n\nVerification: {verification_result.get('explanation', 'N/A')}"
        else:
            # Uncertain cases
            final_decision = "uncertain"
            confidence_level = min(initial_confidence, verification_confidence)
            final_reason = "Analysis was inconclusive. The AI systems could not reach a confident agreement."
        
        return {
            "decision": final_decision,
            "confidence": confidence_level,
            "reason": final_reason,
            "initial_analysis": initial_result,
            "verification": verification_result
        }
        
    except Exception as e:
        print(f"Error in regression analysis: {str(e)}")
        # Return uncertain in case of errors
        return {
            "decision": "uncertain",
            "confidence": 0,
            "reason": f"Error during analysis: {str(e)}",
            "initial_analysis": {},
            "verification": {}
        }

def analyze_single_issue(issue_data, repo_owner, repo_name, github_token, chat_model, parser, prompt, include_comments=False):
    """
    Analyze a single issue and apply labels if needed.
    
    Args:
        issue_data: The issue data from GitHub API
        repo_owner: Repository owner
        repo_name: Repository name
        github_token: GitHub API token
        chat_model: LangChain model for analysis
        parser: Output parser
        prompt: Analysis prompt template (no longer used with new implementation)
        include_comments: Whether to include comments in the analysis
        
    Returns:
        bool: True if the issue is a regression, False otherwise
    """
    issue_number = issue_data["number"]
    issue_title = issue_data["title"]
    issue_body = issue_data["body"] or ""
    
    # Check if issue already has regression label
    has_regression_label = False
    for label in issue_data["labels"]:
        if isinstance(label, dict) and label.get("name") == "regression":
            has_regression_label = True
            break
        elif isinstance(label, str) and label == "regression":
            has_regression_label = True
            break
    
    if has_regression_label:
        print(f"Skipping issue #{issue_number} as it already has regression label.")
        return False
    
    # Check if this is an Outlook-exclusive issue
    if is_outlook_exclusive_issue(issue_title, issue_body, chat_model):
        print(f"Skipping issue #{issue_number} as it's identified as an Outlook-exclusive problem")
        return False
        
    print(f"\nProcessing issue #{issue_number}: {issue_title}")
    print(f"Body length: {len(issue_body)} characters")
    
    # Prepare issue content for analysis
    issue_content = f"Issue Title: {issue_title}\n\nIssue Description: {issue_body}"
    
    # Include comments if requested
    if include_comments:
        print(f"Including comments in analysis for issue #{issue_number}")
        comments = get_issue_comments(repo_owner, repo_name, issue_number, github_token)
        if comments:
            comments_text = "\n\n".join([f"Comment by {comment['user']['login']}: {comment['body']}" for comment in comments])
            issue_content += f"\n\nComments:\n{comments_text}"
            print(f"Added {len(comments)} comments to analysis")
    
    try:
        # Use the new integrated analysis module
        analysis_result = analyze_issue_for_regression(issue_content, issue_number, chat_model)
        final_decision = analysis_result["decision"]
        final_reason = analysis_result["reason"]
        
        # Common headers for GitHub API calls
        headers = {
            "Authorization": f"token {github_token}",
            "Accept": "application/vnd.github.v3+json"
        }
        
        # Take action based on final decision
        if final_decision == "confirmed_regression":
            # Add regression label
            url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/issues/{issue_number}/labels"
            data = {"labels": ["regression"]}
            response = requests.post(url, headers=headers, json=data)
            
            if response.status_code == 200:
                print(f"Successfully added 'regression' label to issue #{issue_number}")
                # Add a comment explaining the label
                comment_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/issues/{issue_number}/comments"
                source = "issue content and comments" if include_comments else "issue content"
                comment_data = {"body": f"This issue was automatically labeled as a regression based on dual-LLM analysis of the {source}.\n\nReason: {final_reason}"}
                requests.post(comment_url, headers=headers, json=comment_data)
                return True
            else:
                print(f"Failed to add label. Status code: {response.status_code}")
                return False
        
        elif final_decision == "confirmed_not_regression":
            print(f"Issue #{issue_number} is confirmed not to be a regression.")
            return False
            
        else:  # final_decision == "uncertain"
            # Add a comment asking for user confirmation with checklist
            print(f"Issue #{issue_number} has uncertain regression status. Asking user for confirmation.")
            
            comment_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/issues/{issue_number}/comments"
            
            # Create a comment with interactive checklist for user input
            comment_body = f"""## Regression Analysis Needs Your Input

Our automated system analyzed this issue but couldn't determine with confidence if this is a regression bug.

A regression bug is when **functionality that previously worked properly no longer works** after a recent change.

**If you believe this is a regression issue, please confirm below:**

- [ ] Yes, confirm this is a regression issue

<!-- comment_id:{issue_number}:{repo_owner}:{repo_name} -->

> Note: Once confirmed, the issue will be labeled as regression.
> After selecting the option, your choice will be automatically processed.

---
*Technical details:*
{final_reason}
"""
            
            comment_data = {"body": comment_body}
            response = requests.post(comment_url, headers=headers, json=comment_data)
            
            if response.status_code == 201:
                print(f"Successfully added user confirmation request to issue #{issue_number}")
            else:
                print(f"Failed to add comment. Status code: {response.status_code}")
                
            return False
            
    except Exception as e:
        print(f"Error processing issue #{issue_number}: {str(e)}")
        return False

def analyze_issues():
    """Process all issues in the repository."""
    try:
        # Get environment variables
        github_token = os.environ["GITHUB_TOKEN"]
        repo_full_name = os.environ["GITHUB_REPOSITORY"]
        repo_owner, repo_name = repo_full_name.split("/")
        
        ensure_label_exists(
            repo_owner, 
            repo_name, 
            github_token, 
            "regression", 
            "d73a4a",  #Red color for regression
            "Functionality that previously worked no longer works"
        )
        
        # Check for specific issue number
        specific_issue = os.environ.get("ISSUE_NUMBER")
        
        # Initialize model - we only need one model now
        model_id = "gpt-4o"
        chat_model = get_azure_chat_model(model_id)
        
        # These are no longer used directly in analyze_single_issue
        # but keeping them for backward compatibility
        parser = JsonOutputParser()
        prompt = None
        
        # Process a specific issue if specified
        if specific_issue:
            issue_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/issues/{specific_issue}"
            headers = {
                "Authorization": f"token {github_token}",
                "Accept": "application/vnd.github.v3+json"
            }
            response = requests.get(issue_url, headers=headers)
            
            if response.status_code == 200:
                issue_data = response.json()
                analyze_single_issue(
                    issue_data, 
                    repo_owner, 
                    repo_name, 
                    github_token, 
                    chat_model, 
                    parser, 
                    prompt,
                    include_comments=True
                )
            else:
                print(f"Error fetching specific issue: {response.status_code}")
                print(response.text)
            
            return
        
        # Process all open issues
        print(f"Fetching all open issues from {repo_owner}/{repo_name}")
        page = 1
        per_page = 30
        processed = 0
        
        while True:
            # Fetch issues with pagination
            issues_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/issues?state=open&per_page={per_page}&page={page}"
            headers = {
                "Authorization": f"token {github_token}",
                "Accept": "application/vnd.github.v3+json"
            }
            
            response = requests.get(issues_url, headers=headers)
            
            if response.status_code != 200:
                print(f"Error fetching issues: {response.status_code}")
                print(response.text)
                break
                
            issues = response.json()
            
            # Break if no more issues
            if not issues:
                break
                
            print(f"Processing page {page} with {len(issues)} issues")
            
            # Process each issue
            for issue in issues:
                # Skip pull requests (they appear in issues API)
                if "pull_request" in issue:
                    continue
                    
                analyze_single_issue(issue, repo_owner, repo_name, github_token, chat_model, parser, prompt)
                processed += 1
                
            # Move to next page
            page += 1
                
        print(f"\nCompleted processing {processed} issues")
        
    except Exception as e:
        print(f"Error occurred: {str(e)}")
        import traceback
        traceback.print_exc()
        raise

if __name__ == "__main__":
    analyze_issues()