"""
Unit tests for analyze_issue.py.

These tests use mocks to avoid real network calls and LLM API calls.
"""

import json
import sys
import os
import unittest
from unittest.mock import MagicMock, patch, call

# Add the parent directory to the path so we can import the module under test
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import analyze_issue


class TestDetectPlatformBehavioralDifference(unittest.TestCase):
    """Tests for the detect_platform_behavioral_difference function."""

    def test_web_desktop_behavioral_difference_detected(self):
        """An issue that clearly describes a web/desktop difference should be detected."""
        llm_response = {
            "is_platform_difference": True,
            "confidence": 0.95,
            "web_only": True,
            "reason": "The issue explicitly states different behavior on web vs desktop.",
        }

        title = "Initial 0 value on cell before displaying actual calculated formula value on Excel Online"
        body = (
            "On the web, when calculating the formula, it would show these values on cell\n"
            "#BUSY > 0 > actual calculated value\n\n"
            "On desktop, value 0 is never shown after #BUSY."
        )

        with patch("analyze_issue.ChatPromptTemplate") as mock_cpt, \
             patch("analyze_issue.JsonOutputParser") as mock_parser_cls:

            mock_parser = MagicMock()
            mock_parser_cls.return_value = mock_parser

            # Build the chain: prompt_instance | chat_model | parser
            mock_prompt_instance = MagicMock()
            mock_cpt.from_messages.return_value = mock_prompt_instance

            mock_chain_after_model = MagicMock()
            mock_chain_after_model.invoke.return_value = llm_response
            mock_chain_after_parser = MagicMock()
            mock_chain_after_parser.invoke.return_value = llm_response

            mock_prompt_instance.pipe.return_value = mock_chain_after_model
            mock_chain_after_model.pipe.return_value = mock_chain_after_parser
            mock_chain_after_parser.invoke.return_value = llm_response

            chat_model = MagicMock()
            result = analyze_issue.detect_platform_behavioral_difference(title, body, chat_model)

        self.assertTrue(result["is_platform_difference"])
        self.assertGreater(result["confidence"], 0.7)
        self.assertTrue(result["web_only"])
        self.assertIsInstance(result["reason"], str)

    def test_non_platform_issue_not_detected(self):
        """An issue that does not describe a web/desktop difference should not be flagged."""
        llm_response = {
            "is_platform_difference": False,
            "confidence": 0.9,
            "web_only": False,
            "reason": "The issue does not mention any difference between web and desktop.",
        }

        title = "Office.js API crashes when calling range.load()"
        body = "Calling range.load() causes the add-in to crash on all platforms."

        with patch("analyze_issue.ChatPromptTemplate") as mock_cpt, \
             patch("analyze_issue.JsonOutputParser") as mock_parser_cls:

            mock_parser = MagicMock()
            mock_parser_cls.return_value = mock_parser

            mock_prompt_instance = MagicMock()
            mock_cpt.from_messages.return_value = mock_prompt_instance

            mock_chain_after_parser = MagicMock()
            mock_chain_after_parser.invoke.return_value = llm_response

            mock_chain_after_model = MagicMock()
            mock_chain_after_model.pipe.return_value = mock_chain_after_parser

            mock_prompt_instance.pipe.return_value = mock_chain_after_model

            chat_model = MagicMock()
            result = analyze_issue.detect_platform_behavioral_difference(title, body, chat_model)

        self.assertFalse(result["is_platform_difference"])
        self.assertFalse(result["web_only"])

    def test_error_during_llm_call_returns_safe_defaults(self):
        """If the LLM call raises an exception the function should return safe False defaults."""
        title = "Some issue title"
        body = "Some issue body"

        with patch("analyze_issue.ChatPromptTemplate") as mock_cpt, \
             patch("analyze_issue.JsonOutputParser") as mock_parser_cls:

            mock_parser = MagicMock()
            mock_parser_cls.return_value = mock_parser

            mock_prompt_instance = MagicMock()
            mock_cpt.from_messages.return_value = mock_prompt_instance

            mock_chain_after_parser = MagicMock()
            mock_chain_after_parser.invoke.side_effect = RuntimeError("LLM API error")

            mock_chain_after_model = MagicMock()
            mock_chain_after_model.pipe.return_value = mock_chain_after_parser

            mock_prompt_instance.pipe.return_value = mock_chain_after_model

            chat_model = MagicMock()
            result = analyze_issue.detect_platform_behavioral_difference(title, body, chat_model)

        self.assertFalse(result["is_platform_difference"])
        self.assertEqual(result["confidence"], 0)
        self.assertFalse(result["web_only"])
        self.assertIn("Error", result["reason"])

    def test_low_confidence_result_returns_correct_values(self):
        """A result with low confidence should still be returned as-is."""
        llm_response = {
            "is_platform_difference": True,
            "confidence": 0.4,
            "web_only": True,
            "reason": "Possibly a web/desktop difference but not certain.",
        }

        title = "Possible web issue"
        body = "This might be different on web."

        with patch("analyze_issue.ChatPromptTemplate") as mock_cpt, \
             patch("analyze_issue.JsonOutputParser") as mock_parser_cls:

            mock_parser = MagicMock()
            mock_parser_cls.return_value = mock_parser

            mock_prompt_instance = MagicMock()
            mock_cpt.from_messages.return_value = mock_prompt_instance

            mock_chain_after_parser = MagicMock()
            mock_chain_after_parser.invoke.return_value = llm_response

            mock_chain_after_model = MagicMock()
            mock_chain_after_model.pipe.return_value = mock_chain_after_parser

            mock_prompt_instance.pipe.return_value = mock_chain_after_model

            chat_model = MagicMock()
            result = analyze_issue.detect_platform_behavioral_difference(title, body, chat_model)

        # The function returns the raw values; caller is responsible for threshold checks
        self.assertTrue(result["is_platform_difference"])
        self.assertAlmostEqual(result["confidence"], 0.4)


class TestEnsureLabelExists(unittest.TestCase):
    """Tests for the ensure_label_exists helper."""

    @patch("analyze_issue.requests.get")
    @patch("analyze_issue.requests.post")
    def test_creates_label_when_not_found(self, mock_post, mock_get):
        """ensure_label_exists should create the label when it doesn't exist (404)."""
        mock_get.return_value = MagicMock(status_code=404)
        mock_post.return_value = MagicMock(status_code=201)

        result = analyze_issue.ensure_label_exists(
            "owner", "repo", "token", "platform: web", "0075ca",
            "Issue behavior is specific to Office on the web"
        )

        self.assertTrue(result)
        mock_post.assert_called_once()
        call_kwargs = mock_post.call_args
        posted_data = call_kwargs[1]["json"]
        self.assertEqual(posted_data["name"], "platform: web")
        self.assertEqual(posted_data["color"], "0075ca")

    @patch("analyze_issue.requests.get")
    @patch("analyze_issue.requests.patch")
    def test_updates_label_when_color_changed(self, mock_patch, mock_get):
        """ensure_label_exists should update the label when the color differs."""
        mock_get.return_value = MagicMock(
            status_code=200,
            json=MagicMock(return_value={"color": "ffffff", "description": "old description"})
        )
        mock_patch.return_value = MagicMock(status_code=200)

        result = analyze_issue.ensure_label_exists(
            "owner", "repo", "token", "platform: web", "0075ca"
        )

        self.assertTrue(result)
        mock_patch.assert_called_once()

    @patch("analyze_issue.requests.get")
    def test_returns_true_when_label_already_correct(self, mock_get):
        """ensure_label_exists should return True without changes when label is correct."""
        mock_get.return_value = MagicMock(
            status_code=200,
            json=MagicMock(return_value={"color": "0075ca", "description": "Issue behavior is specific to Office on the web"})
        )

        result = analyze_issue.ensure_label_exists(
            "owner", "repo", "token", "platform: web", "0075ca",
            "Issue behavior is specific to Office on the web"
        )

        self.assertTrue(result)


if __name__ == "__main__":
    unittest.main()

