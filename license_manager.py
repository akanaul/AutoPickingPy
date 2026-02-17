"""
AutoPickingPy License and User Authorization Module
SECURITY MODEL: All authorization requires live GitHub verification - NO LOCAL CACHE
"""

import os
import sys
import json
import hashlib
import requests
from datetime import datetime, timedelta
from pathlib import Path


class LicenseManager:
    """Manages software licensing with live GitHub gist verification (no local files)."""
    
    # GitHub gist configuration (update with your gist URL)
    # Create a GitHub gist with AUTHORIZED_USERS.json content and paste the raw URL here
    GITHUB_GIST_URL = "https://gist.githubusercontent.com/akanaul/YOUR_GIST_ID/raw/AUTHORIZED_USERS.json"
    
    # Audit logging (only records attempts, never stores auth data)
    AUDIT_DIR = Path(os.path.expanduser("~/.autopickingpy"))
    AUDIT_LOG = AUDIT_DIR / "authorization_audit.log"
    
    # Network settings
    NETWORK_TIMEOUT = 10  # seconds
    RETRY_ATTEMPTS = 3
    
    def __init__(self):
        """Initialize the license manager."""
        self.audit_dir.mkdir(parents=True, exist_ok=True)
        self.authorized_users = None
        self.fetch_timestamp = None
        
    @property
    def audit_dir(self):
        return self.AUDIT_DIR
    
    def load_license_header(self):
        """Display license notice to user."""
        header = """
================================================================================
                         AutoPickingPy - Licensed Software
================================================================================
This software is proprietary and distributed under a "All Rights Reserved" license.
Authorization requires live verification via GitHub - offline use is not permitted.

For licensing and authorization, contact: Clebson Luan Alves da Silva

See LICENSE file for complete terms.
================================================================================
"""
        print(header)
    
    def get_system_identifier(self) -> str:
        """Generate a unique system identifier for this computer."""
        try:
            import uuid
            mac = uuid.UUID(int=uuid.getnode()).hex[-12:]
            hostname = os.environ.get('COMPUTERNAME', 'UNKNOWN')
            system_id = f"{hostname}_{mac}".lower()
            return hashlib.sha256(system_id.encode()).hexdigest()[:16]
        except Exception:
            return "UNKNOWN"
    
    def log_authorization_attempt(self, username: str, status: str, reason: str = ""):
        """Log authorization attempt for audit trail (not auth data)."""
        try:
            timestamp = datetime.now().isoformat()
            system_id = self.get_system_identifier()
            
            log_entry = {
                "timestamp": timestamp,
                "username": username,
                "status": status,
                "reason": reason,
                "system_id": system_id,
                "github_verified": False  # Will be updated based on result
            }
            
            with open(self.AUDIT_LOG, 'a') as f:
                f.write(json.dumps(log_entry) + "\n")
        except Exception as e:
            print(f"[WARNING] Could not write audit log: {e}")
    
    def fetch_authorized_users_from_github(self, retry=0) -> dict:
        """
        Fetch the LIVE list of authorized users from GitHub gist.
        This CANNOT be cached - authorization is only valid if GitHub confirms it.
        
        Returns dict with user info, or None if GitHub cannot be reached.
        """
        try:
            # Add timestamp to prevent any caching by intermediaries
            cache_bust = datetime.now().timestamp()
            url = f"{self.GITHUB_GIST_URL}?t={cache_bust}"
            
            response = requests.get(url, timeout=self.NETWORK_TIMEOUT)
            
            if response.status_code == 200:
                data = response.json()
                self.fetch_timestamp = datetime.now()
                return data
            else:
                print(f"[ERROR] GitHub gist returned status {response.status_code}")
                return None
                
        except requests.exceptions.Timeout:
            if retry < self.RETRY_ATTEMPTS:
                print(f"[INFO] Timeout, retrying... (attempt {retry + 1}/{self.RETRY_ATTEMPTS})")
                return self.fetch_authorized_users_from_github(retry + 1)
            print("[ERROR] Network timeout - GitHub is unreachable")
            return None
        except requests.exceptions.ConnectionError:
            if retry < self.RETRY_ATTEMPTS:
                print(f"[INFO] Connection failed, retrying... (attempt {retry + 1}/{self.RETRY_ATTEMPTS})")
                return self.fetch_authorized_users_from_github(retry + 1)
            print("[ERROR] Cannot connect to GitHub - authorization requires internet connection")
            return None
        except Exception as e:
            print(f"[ERROR] Failed to fetch authorization: {e}")
            return None
    
    def is_user_authorized(self, username: str, machine_id: str = None) -> tuple[bool, str]:
        """
        Check if a user is authorized (LIVE GitHub gist check only).
        
        Args:
            username: GitHub username (from gist keys)
            machine_id: Optional machine identifier for hardware-locked licenses
        
        Returns:
            (authorized: bool, reason: str)
        """
        if machine_id is None:
            machine_id = self.get_system_identifier()
        
        print(f"\n[INFO] Verifying authorization with GitHub...")
        
        # ALWAYS fetch fresh data from GitHub - no cache allowed
        users = self.fetch_authorized_users_from_github()
        
        if users is None:
            reason = "GitHub unreachable - cannot verify authorization"
            print(f"[CRITICAL] {reason}")
            return False, reason
        
        if username not in users:
            reason = f"User '{username}' not found in authorization database"
            print(f"[ERROR] {reason}")
            return False, reason
        
        user_data = users[username]
        
        # Check if revoked (highest priority)
        if user_data.get('revoked', False):
            reason = "License has been REVOKED"
            print(f"[ERROR] {reason}")
            return False, reason
        
        # Check expiration
        if 'expires' in user_data:
            try:
                expiry = datetime.fromisoformat(user_data['expires'])
                if datetime.now() > expiry:
                    reason = f"Authorization expired on {user_data['expires']}"
                    print(f"[ERROR] {reason}")
                    return False, reason
            except ValueError:
                reason = "Invalid expiration date in authorization data"
                print(f"[ERROR] {reason}")
                return False, reason
        
        # Check machine ID if hardware-locked
        if 'machine_id' in user_data and user_data['machine_id']:
            if user_data['machine_id'] != machine_id:
                reason = f"License is locked to different machine (current: {machine_id[:8]}...)"
                print(f"[ERROR] {reason}")
                return False, reason
        
        # All checks passed
        reason = "GitHub verified authorization confirmed"
        return True, reason
    
    def authorize(self, username: str, machine_id: str = None) -> bool:
        """
        Authorize a user (LIVE GitHub verification required).
        
        Args:
            username: GitHub username
            machine_id: Optional machine ID for hardware lock
            
        Returns:
            True if authorized successfully, False otherwise
        """
        print(f"\n[INFO] Authorizing user: {username}")
        print(f"[INFO] System ID: {self.get_system_identifier()[:8]}...")
        
        authorized, reason = self.is_user_authorized(username, machine_id)
        
        if authorized:
            print(f"\n[SUCCESS] ✓ Authorization GRANTED")
            print(f"[INFO] {reason}")
            self.log_authorization_attempt(username, "GRANTED", reason)
            return True
        else:
            print(f"\n[ERROR] ✗ Authorization DENIED")
            print(f"[ERROR] Reason: {reason}")
            print(f"\n[INFO] Contact Clebson Luan Alves da Silva for licensing")
            self.log_authorization_attempt(username, "DENIED", reason)
            return False
    
    def verify_login(self, username: str, password: str = None) -> bool:
        """
        Verify user login credentials with GitHub.
        Requires user to authenticate with valid GitHub credentials.
        
        Args:
            username: GitHub username
            password: GitHub personal access token or password
            
        Returns:
            True if credentials verified, False otherwise
        """
        print(f"\n[INFO] Verifying GitHub login credentials...")
        
        # If no password provided, get it from prompt
        if password is None:
            import getpass
            password = getpass.getpass("GitHub Personal Access Token (or password): ")
        
        if not password:
            print("[ERROR] No credentials provided")
            return False
        
        # Verify credentials via GitHub API
        try:
            auth = (username, password)
            response = requests.get(
                "https://api.github.com/user",
                auth=auth,
                timeout=self.NETWORK_TIMEOUT
            )
            
            if response.status_code == 200:
                user_data = response.json()
                github_username = user_data.get('login', '').lower()
                
                if github_username == username.lower():
                    print(f"[SUCCESS] ✓ GitHub login verified for @{github_username}")
                    return True
                else:
                    print(f"[ERROR] GitHub username mismatch")
                    return False
            elif response.status_code == 401:
                print("[ERROR] Invalid GitHub credentials")
                return False
            else:
                print(f"[ERROR] GitHub API returned status {response.status_code}")
                return False
                
        except requests.exceptions.RequestException as e:
            print(f"[ERROR] Failed to verify credentials: {e}")
            return False
    
    def get_authorization_info(self, username: str) -> dict:
        """
        Get authorization details for a user (requires live GitHub check).
        """
        users = self.fetch_authorized_users_from_github()
        if users:
            return users.get(username, {})
        return {}


class LicenseDisplay:
    """Display license and authorization information."""
    
    @staticmethod
    def show_license():
        """Display the full license to the user."""
        license_path = Path(__file__).parent / "LICENSE"
        if license_path.exists():
            with open(license_path, 'r') as f:
                print(f.read())
        else:
            print("LICENSE file not found")
    
    @staticmethod
    def show_user_info(license_manager: LicenseManager, username: str):
        """Display user's authorization details."""
        info = license_manager.get_authorization_info(username)
        if not info:
            print(f"No authorization found for {username}")
            return
        
        print(f"\n=== Authorization Details for {username} ===")
        
        # Display all info except sensitive fields
        for key, value in info.items():
            if key == 'machine_id':
                if value:
                    print(f"  Hardware Lock: {value[:8]}...")
                else:
                    print(f"  Hardware Lock: Not locked (any machine)")
            elif key not in ['revoked']:  # Hide internal fields
                display_key = key.replace('_', ' ').title()
                print(f"  {display_key}: {value}")
        
        if info.get('revoked'):
            print(f"  STATUS: ⚠ REVOKED")


def get_github_username_from_git() -> str:
    """
    Auto-detect GitHub username from local Git configuration.
    Tries multiple sources:
    1. Git user.name (if contains github username)
    2. Git user.email (extracts username from email)
    3. GitHub CLI auth (gh api user -q .login)
    
    Returns empty string if not found.
    """
    import subprocess
    
    try:
        # Try git user.name
        result = subprocess.run(
            ['git', 'config', '--get', 'user.name'],
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0:
            name = result.stdout.strip().lower()
            # Check if it looks like a GitHub username (no spaces, short)
            if name and ' ' not in name and len(name) < 40:
                return name
    except Exception:
        pass
    
    try:
        # Try git user.email - extract username from email
        result = subprocess.run(
            ['git', 'config', '--get', 'user.email'],
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0:
            email = result.stdout.strip()
            # Extract part before @ if it's a GitHub email
            if '@' in email:
                username = email.split('@')[0].lower()
                if username and ' ' not in username:
                    return username
    except Exception:
        pass
    
    try:
        # Try GitHub CLI if installed
        result = subprocess.run(
            ['gh', 'api', 'user', '-q', '.login'],
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0:
            username = result.stdout.strip().lower()
            if username:
                return username
    except Exception:
        pass
    
    return ""


def get_authorized_username() -> str:
    """
    Get GitHub username from multiple sources in priority order:
    1. AUTOPICKING_USER environment variable
    2. Auto-detect from Git config
    3. Prompt user if auto-detection fails
    """
    # Try environment variable first (for automated/headless execution)
    username = os.environ.get('AUTOPICKING_USER')
    if username:
        print(f"[INFO] Using username from AUTOPICKING_USER environment variable")
        return username
    
    # Try auto-detect from Git configuration
    print("[INFO] Attempting to auto-detect GitHub account from Git config...")
    username = get_github_username_from_git()
    if username:
        print(f"[SUCCESS] ✓ Detected GitHub account: @{username}")
        return username
    
    # Fall back to manual prompt
    print("\n" + "="*80)
    print("Enter your GitHub username for authorization verification")
    print("(Must be registered in the GitHub gist authorization database)")
    print("="*80)
    username = input("\nGitHub Username: ").strip()
    return username


def check_license_and_authorize() -> bool:
    """
    Main function to check license and authorize user.
    Fully automated with auto-detect based account detection.
    
    Steps:
    1. Get GitHub username (auto-detect or prompt)
    2. Verify GitHub login credentials (password/PAT)
    3. Check if user is in authorization database
    4. Grant or deny access
    
    Call this at the start of your main application.
    
    Returns:
        True if user is authorized, False otherwise
    """
    try:
        manager = LicenseManager()
        manager.load_license_header()
        
        # Step 1: Get username (with auto-detect)
        print("\n" + "="*80)
        print("Initializing Authorization System")
        print("="*80)
        username = get_authorized_username()
        if not username:
            print("[ERROR] No username provided")
            manager.log_authorization_attempt("UNKNOWN", "DENIED", "No username provided")
            return False
        
        # Step 2: Verify GitHub login credentials
        print("\n" + "="*80)
        print(f"GitHub Login Verification for @{username}")
        print("="*80)
        print("Provide your GitHub credentials to verify account ownership.")
        print("Use your GitHub password or a Personal Access Token (PAT).")
        print("For security, credentials are NOT saved or logged.")
        
        if not manager.verify_login(username):
            print("\n[ERROR] GitHub login verification failed")
            manager.log_authorization_attempt(username, "DENIED", "Login verification failed")
            return False
        
        # Step 3: Check authorization
        print("\n[INFO] Checking authorization database...")
        
        # Authorize (LIVE GitHub check)
        return manager.authorize(username)
        
    except KeyboardInterrupt:
        print("\n[INFO] Authorization cancelled by user")
        return False
    except Exception as e:
        print(f"\n[CRITICAL] Unexpected error during authorization: {e}")
        return False


if __name__ == "__main__":
    # Test the license manager
    print("AutoPickingPy License Manager - Test Mode")
    print("(All authorization requires live GitHub verification)\n")
    
    manager = LicenseManager()
    manager.load_license_header()
    
    # Test with sample username
    test_user = input("\nEnter username to check: ").strip() or "demo"
    
    print(f"\nChecking authorization for: {test_user}")
    authorized, reason = manager.is_user_authorized(test_user)
    
    if authorized:
        info = manager.get_authorization_info(test_user)
        print(f"\n✓ Authorization info for {test_user}:")
        print(json.dumps(info, indent=2))
    else:
        print(f"\n✗ Authorization check failed: {reason}")
    
    # Show audit log
    print(f"\nAudit log location: {manager.AUDIT_LOG}")
    if manager.AUDIT_LOG.exists():
        print(f"Recent audit entries:")
        with open(manager.AUDIT_LOG, 'r') as f:
            entries = f.readlines()[-5:]  # Last 5 entries
            for entry in entries:
                print(f"  {entry.rstrip()}")
