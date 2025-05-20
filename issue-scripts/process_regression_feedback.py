import os
import re
import requests

def process_regression_feedback():
    """Process user feedback on regression analysis."""
    # Get environment variables
    github_token = os.environ["GITHUB_TOKEN"]
    comment_id = os.environ["COMMENT_ID"]
    comment_body = os.environ["COMMENT_BODY"]
    issue_number = os.environ["ISSUE_NUMBER"]
    
    print(f"Processing feedback for issue #{issue_number}")
    
    # Extract repository information from the comment body using regex
    match = re.search(r'<!-- comment_id:(\d+):([^:]+):([^:]+) -->', comment_body)
    if not match:
        repo_full_name = os.environ["GITHUB_REPOSITORY"]
        repo_owner, repo_name = repo_full_name.split("/")
        print("Using environment variables for repository information")
    else:
        comment_issue_number = match.group(1)
        repo_owner = match.group(2)
        repo_name = match.group(3)
    
    print(f"Processing feedback for {repo_owner}/{repo_name} issue #{issue_number}")
    
    # API headers
    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json"
    }
    
    # Check if the option is selected
    is_confirmed = "- [x] Yes, confirm this is a regression issue" in comment_body
    
    # If the option is not selected, do nothing
    if not is_confirmed:
        print("No confirmation yet.")
        return
    
    # The option is selected, process it
    print(f"User confirmed issue #{issue_number} is a regression")
    
    # Add regression label
    url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/issues/{issue_number}/labels"
    data = {"labels": ["regression"]}
    label_response = requests.post(url, headers=headers, json=data)
    
    if label_response.status_code == 200:
        print(f"Successfully added 'regression' label to issue #{issue_number}")
    else:
        print(f"Failed to add label. Status code: {label_response.status_code}")
    
    # Update the original comment to lock the selection by replacing the checkbox with a green checkmark
    comment_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/issues/comments/{comment_id}"
    
    # Replace the checkbox with a green checkmark and lock the selection
    updated_body = comment_body.replace(
        "- [x] Yes, confirm this is a regression issue", 
        "✅ **Confirmed as regression issue** _(selection locked)_"
    )
    
    # Add a confirmation message to the note
    updated_body = updated_body.replace(
        "> Note: Once confirmed, the issue will be labeled as regression.", 
        "> ✅ **Confirmed**: This issue has been labeled as regression.\n> Your selection has been processed and locked."
    )
    
    update_data = {"body": updated_body}
    update_response = requests.patch(comment_url, headers=headers, json=update_data)
    
    if update_response.status_code == 200:
        print("Successfully updated comment to lock the selection")
    else:
        print(f"Failed to update comment. Status code: {update_response.status_code}")
    
    # Add a thank you comment
    thank_you_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/issues/{issue_number}/comments"
    thank_you_data = {
        "body": "Thank you for confirming this is a regression issue. The issue has been labeled accordingly."
    }
    requests.post(thank_you_url, headers=headers, json=thank_you_data)

if __name__ == "__main__":
    process_regression_feedback()