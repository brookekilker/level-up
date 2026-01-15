"""
Jira API Helper Module
======================
Shared Jira API utilities for use across Streamlit apps.
Credentials are loaded from .env file - no UI prompts needed.
"""

import os
import re
import requests
from pathlib import Path
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
from typing import Optional, Dict

# Load environment variables from .env file
load_dotenv(Path(__file__).parent / ".env")


# =============================================================================
# CONFIGURATION
# =============================================================================

class JiraConfig:
    """Jira configuration loaded from environment variables."""
    BASE_URL = os.environ.get("JIRA_BASE_URL", "").strip().strip('"')
    EMAIL = os.environ.get("JIRA_EMAIL", "").strip().strip('"')
    API_TOKEN = os.environ.get("JIRA_API_TOKEN", "").strip().strip('"')

    @classmethod
    def is_configured(cls) -> bool:
        """Check if all Jira credentials are configured."""
        return all([cls.BASE_URL, cls.EMAIL, cls.API_TOKEN])

    @classmethod
    def get_missing(cls) -> list:
        """Return list of missing configuration items."""
        missing = []
        if not cls.BASE_URL:
            missing.append("JIRA_BASE_URL")
        if not cls.EMAIL:
            missing.append("JIRA_EMAIL")
        if not cls.API_TOKEN:
            missing.append("JIRA_API_TOKEN")
        return missing


# =============================================================================
# JIRA API CLIENT
# =============================================================================

class JiraClient:
    """
    Jira API client for fetching ticket information.

    Usage:
        client = JiraClient()
        if client.is_available():
            ticket = client.fetch_ticket("PROJ-123")
    """

    def __init__(self):
        self.base_url = JiraConfig.BASE_URL.rstrip("/") if JiraConfig.BASE_URL else ""
        self.auth = HTTPBasicAuth(JiraConfig.EMAIL, JiraConfig.API_TOKEN) if JiraConfig.is_configured() else None
        self.headers = {"Accept": "application/json"}

    def is_available(self) -> bool:
        """Check if Jira client is properly configured."""
        return JiraConfig.is_configured()

    def extract_ticket_key(self, url_or_key: str) -> Optional[str]:
        """
        Extract Jira ticket key from URL or return key if already a key.

        Supports:
        - https://company.atlassian.net/browse/PROJ-123
        - https://company.atlassian.net/jira/software/projects/PROJ/boards/1?selectedIssue=PROJ-123
        - PROJ-123 (direct key)
        """
        if not url_or_key:
            return None

        url_or_key = url_or_key.strip()

        # Match ticket key pattern anywhere in the string
        match = re.search(r"([A-Z][A-Z0-9]+-\d+)", url_or_key)
        if match:
            return match.group(1)
        return None

    def fetch_ticket(self, ticket_key: str) -> Optional[Dict]:
        """
        Fetch ticket information from Jira API.

        Returns dict with: key, summary, status, assignee, project_key, url
        Returns None on error.
        """
        if not self.is_available():
            return None

        api_url = f"{self.base_url}/rest/api/3/issue/{ticket_key}"

        try:
            response = requests.get(
                api_url,
                headers=self.headers,
                auth=self.auth,
                timeout=10
            )

            if response.status_code == 200:
                data = response.json()

                # Extract project key from the ticket key
                project_key = ticket_key.split("-")[0] if "-" in ticket_key else ""

                return {
                    "key": ticket_key,
                    "summary": data["fields"]["summary"],
                    "status": data["fields"]["status"]["name"],
                    "assignee": data["fields"].get("assignee", {}).get("displayName") if data["fields"].get("assignee") else "Unassigned",
                    "project_key": project_key,
                    "url": f"{self.base_url}/browse/{ticket_key}"
                }
            elif response.status_code == 401:
                raise PermissionError("Jira authentication failed. Check JIRA_EMAIL and JIRA_API_TOKEN in .env")
            elif response.status_code == 404:
                raise ValueError(f"Ticket {ticket_key} not found")
            else:
                raise RuntimeError(f"Jira API error: {response.status_code}")

        except requests.exceptions.RequestException as e:
            raise ConnectionError(f"Network error connecting to Jira: {e}")

    def get_ticket_url(self, ticket_key: str) -> str:
        """Generate the browse URL for a ticket."""
        return f"{self.base_url}/browse/{ticket_key}"


# =============================================================================
# CONVENIENCE FUNCTIONS
# =============================================================================

def get_jira_client() -> JiraClient:
    """Get a configured Jira client instance."""
    return JiraClient()


def fetch_ticket_from_url(url: str) -> Optional[Dict]:
    """
    Convenience function to fetch ticket info from a Jira URL.

    Args:
        url: Jira ticket URL or ticket key

    Returns:
        Dict with ticket info or None on error

    Raises:
        ValueError: If ticket key cannot be extracted or ticket not found
        PermissionError: If authentication fails
        ConnectionError: If network error occurs
    """
    client = get_jira_client()

    if not client.is_available():
        raise EnvironmentError(
            f"Jira not configured. Missing: {', '.join(JiraConfig.get_missing())}"
        )

    ticket_key = client.extract_ticket_key(url)
    if not ticket_key:
        raise ValueError(f"Could not extract ticket key from: {url}")

    return client.fetch_ticket(ticket_key)
