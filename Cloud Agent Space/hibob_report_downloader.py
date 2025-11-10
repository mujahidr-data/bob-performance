#!/usr/bin/env python3
"""
HiBob Performance Report Downloader
Automates login to HiBob via JumpCloud SSO and downloads performance cycle reports.
"""

import json
import os
import sys
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import requests


class HiBobReportDownloader:
    def __init__(self, config_path="config.json"):
        """Initialize the downloader with configuration."""
        self.config = self.load_config(config_path)
        self.download_dir = Path("downloads")
        self.download_dir.mkdir(exist_ok=True)
        self.browser = None
        self.context = None
        self.page = None
        
    def load_config(self, config_path):
        """Load configuration from JSON file, optionally fetching credentials from Apps Script."""
        if not os.path.exists(config_path):
            print(f"❌ Error: Configuration file '{config_path}' not found!")
            print(f"Please create it from the template:")
            print(f"  cp config.template.json config.json")
            print(f"  # Then edit config.json with your credentials")
            sys.exit(1)
            
        with open(config_path, 'r') as f:
            config = json.load(f)
        
        # Try to fetch credentials from Apps Script if not in config.json
        if not config.get('email') or config.get('email', '').startswith('your-'):
            print("📡 Attempting to fetch credentials from Apps Script...")
            creds = self.fetch_credentials_from_apps_script(config.get('apps_script_url'))
            if creds:
                config['email'] = creds['email']
                config['password'] = creds['password']
                print("✅ Credentials loaded from Apps Script")
            else:
                print("⚠️  Could not fetch from Apps Script, using config.json")
            
        # Validate required fields
        required_fields = ['email', 'password', 'apps_script_url']
        missing_fields = [field for field in required_fields if not config.get(field) or config[field].startswith('your-')]
        
        if missing_fields:
            print(f"❌ Error: Please configure the following fields in config.json:")
            for field in missing_fields:
                print(f"  - {field}")
            print("\n💡 Tip: You can set credentials in Google Sheets:")
            print("   Bob Salary Data > Performance Reports > Set HiBob Credentials")
            sys.exit(1)
            
        return config
    
    def fetch_credentials_from_apps_script(self, apps_script_url):
        """Fetch credentials from Apps Script API endpoint."""
        if not apps_script_url:
            return None
        
        try:
            # Convert POST URL to GET URL for credentials endpoint
            # Replace /exec with /exec?action=getCredentials
            base_url = apps_script_url.replace('/exec', '')
            creds_url = f"{base_url}/exec?action=getCredentials"
            
            response = requests.get(creds_url, timeout=10)
            if response.status_code == 200:
                data = response.json()
                if data.get('success') and data.get('email') and data.get('password'):
                    return {
                        'email': data['email'],
                        'password': data['password']
                    }
        except Exception as e:
            # Silently fail - will use config.json instead
            pass
        
        return None
    
    def start_browser(self):
        """Start Playwright browser instance."""
        print("🌐 Starting browser...")
        playwright = sync_playwright().start()
        self.browser = playwright.chromium.launch(headless=False)  # Set to True for headless mode
        self.context = self.browser.new_context(
            accept_downloads=True,
            viewport={'width': 1920, 'height': 1080}
        )
        self.page = self.context.new_page()
        
    def close_browser(self):
        """Close browser and cleanup."""
        if self.browser:
            self.browser.close()
            
    def login_to_hibob(self):
        """Login to HiBob via JumpCloud SSO."""
        try:
            print("🔐 Navigating to HiBob login page...")
            self.page.goto("https://app.hibob.com/login/home", wait_until="networkidle", timeout=30000)
            
            print("📧 Entering email address...")
            # Wait for email input field
            email_input = self.page.wait_for_selector('input[type="email"], input[name="email"], input[id="email"]', timeout=10000)
            email_input.fill(self.config['email'])
            
            # Click continue/next button
            print("⏭️  Clicking continue...")
            continue_button = self.page.wait_for_selector('button[type="submit"], button:has-text("Continue"), button:has-text("Next")', timeout=5000)
            continue_button.click()
            
            # Wait for JumpCloud SSO redirect
            print("🔄 Waiting for JumpCloud SSO redirect...")
            time.sleep(3)  # Give time for redirect
            
            # On JumpCloud page - enter email again if needed
            print("🔑 Entering JumpCloud credentials...")
            try:
                # Try to find email field on JumpCloud
                jumpcloud_email = self.page.wait_for_selector('input[type="email"], input[name="email"], input[id="email"]', timeout=5000)
                jumpcloud_email.fill(self.config['email'])
            except PlaywrightTimeoutError:
                print("   (Email field not required, already filled)")
            
            # Enter password
            password_input = self.page.wait_for_selector('input[type="password"], input[name="password"], input[id="password"]', timeout=10000)
            password_input.fill(self.config['password'])
            
            # Click sign in button
            print("✅ Signing in...")
            signin_button = self.page.wait_for_selector('button[type="submit"], button:has-text("Sign In"), button:has-text("Log In")', timeout=5000)
            signin_button.click()
            
            # Wait for successful login - check for HiBob dashboard
            print("⏳ Waiting for login to complete...")
            self.page.wait_for_url("**/app.hibob.com/**", timeout=30000)
            time.sleep(3)  # Additional wait for page to fully load
            
            print("✅ Successfully logged in to HiBob!")
            return True
            
        except PlaywrightTimeoutError as e:
            print(f"❌ Login failed: Timeout waiting for element")
            self.page.screenshot(path="error_login.png")
            print(f"📸 Screenshot saved to error_login.png")
            return False
        except Exception as e:
            print(f"❌ Login failed: {str(e)}")
            self.page.screenshot(path="error_login.png")
            print(f"📸 Screenshot saved to error_login.png")
            return False
    
    def navigate_to_performance_cycles(self):
        """Navigate to performance cycles page."""
        try:
            print("📊 Navigating to Performance Cycles page...")
            self.page.goto("https://app.hibob.com/performance/manage-cycles/main/cycles", wait_until="networkidle", timeout=30000)
            time.sleep(2)  # Wait for page to fully load
            print("✅ On Performance Cycles page")
            return True
        except Exception as e:
            print(f"❌ Failed to navigate to performance cycles: {str(e)}")
            self.page.screenshot(path="error_navigation.png")
            print(f"📸 Screenshot saved to error_navigation.png")
            return False
    
    def search_for_report(self, report_name):
        """Search for a report by name (case-insensitive, partial match)."""
        try:
            print(f"🔍 Searching for report: '{report_name}'")
            
            # Wait for page to load
            time.sleep(2)
            
            # Get page content and search for report links/cards
            # This will need to be customized based on the actual page structure
            page_content = self.page.content()
            
            # Try multiple selectors to find report elements
            possible_selectors = [
                'a[href*="cycle"]',
                'div[role="row"]',
                'tr',
                '[data-test-id*="cycle"]',
                '.cycle-name',
                '[class*="cycle"]'
            ]
            
            report_elements = []
            for selector in possible_selectors:
                try:
                    elements = self.page.query_selector_all(selector)
                    if elements:
                        report_elements.extend(elements)
                except:
                    continue
            
            # Filter elements that contain the report name (case-insensitive)
            matching_reports = []
            report_name_lower = report_name.lower()
            
            for element in report_elements:
                try:
                    text = element.inner_text()
                    if report_name_lower in text.lower():
                        matching_reports.append({
                            'element': element,
                            'text': text.strip()
                        })
                except:
                    continue
            
            # Remove duplicates based on text
            seen_texts = set()
            unique_reports = []
            for report in matching_reports:
                if report['text'] not in seen_texts:
                    seen_texts.add(report['text'])
                    unique_reports.append(report)
            
            if not unique_reports:
                print(f"❌ No reports found matching '{report_name}'")
                return None
            
            if len(unique_reports) == 1:
                print(f"✅ Found 1 matching report: {unique_reports[0]['text']}")
                return unique_reports[0]['element']
            
            # Multiple matches - ask user to select
            print(f"\n📋 Found {len(unique_reports)} matching reports:")
            for i, report in enumerate(unique_reports, 1):
                print(f"  {i}. {report['text']}")
            
            while True:
                try:
                    choice = input(f"\nEnter the number of the report to download (1-{len(unique_reports)}): ").strip()
                    choice_num = int(choice)
                    if 1 <= choice_num <= len(unique_reports):
                        selected_report = unique_reports[choice_num - 1]
                        print(f"✅ Selected: {selected_report['text']}")
                        return selected_report['element']
                    else:
                        print(f"❌ Please enter a number between 1 and {len(unique_reports)}")
                except ValueError:
                    print("❌ Please enter a valid number")
                except KeyboardInterrupt:
                    print("\n❌ Cancelled by user")
                    return None
            
        except Exception as e:
            print(f"❌ Error searching for report: {str(e)}")
            self.page.screenshot(path="error_search.png")
            print(f"📸 Screenshot saved to error_search.png")
            return None
    
    def download_report(self, report_element):
        """Click report and download it."""
        try:
            print("📄 Opening report...")
            report_element.click()
            time.sleep(3)  # Wait for report page to load
            
            print("🔍 Looking for Actions button...")
            # Try to find the Actions button (top right corner)
            actions_selectors = [
                'button:has-text("Actions")',
                'button:has-text("Action")',
                '[data-test-id*="action"]',
                'button[aria-label*="action" i]',
                '.actions-button'
            ]
            
            actions_button = None
            for selector in actions_selectors:
                try:
                    actions_button = self.page.wait_for_selector(selector, timeout=5000)
                    if actions_button:
                        break
                except:
                    continue
            
            if not actions_button:
                print("❌ Could not find Actions button")
                self.page.screenshot(path="error_actions_button.png")
                print(f"📸 Screenshot saved to error_actions_button.png")
                return None
            
            print("📥 Clicking Actions button...")
            actions_button.click()
            time.sleep(1)
            
            # Look for "Download cycle report" option
            print("🔍 Looking for Download option...")
            download_selectors = [
                'button:has-text("Download cycle report")',
                'a:has-text("Download cycle report")',
                '[role="menuitem"]:has-text("Download")',
                'button:has-text("Download")',
                'a:has-text("Download")'
            ]
            
            download_button = None
            for selector in download_selectors:
                try:
                    download_button = self.page.wait_for_selector(selector, timeout=5000)
                    if download_button:
                        break
                except:
                    continue
            
            if not download_button:
                print("❌ Could not find Download option")
                self.page.screenshot(path="error_download_button.png")
                print(f"📸 Screenshot saved to error_download_button.png")
                return None
            
            print("⬇️  Clicking Download...")
            
            # Set up download handler
            with self.page.expect_download(timeout=60000) as download_info:
                download_button.click()
            
            download = download_info.value
            
            # Save the download
            filename = download.suggested_filename
            filepath = self.download_dir / filename
            download.save_as(filepath)
            
            print(f"✅ Report downloaded: {filepath}")
            return filepath
            
        except PlaywrightTimeoutError as e:
            print(f"❌ Timeout while downloading report")
            self.page.screenshot(path="error_download.png")
            print(f"📸 Screenshot saved to error_download.png")
            return None
        except Exception as e:
            print(f"❌ Error downloading report: {str(e)}")
            self.page.screenshot(path="error_download.png")
            print(f"📸 Screenshot saved to error_download.png")
            return None
    
    def upload_to_google_sheets(self, filepath):
        """Upload the downloaded file to Google Sheets via Apps Script."""
        try:
            print("☁️  Uploading to Google Sheets...")
            
            with open(filepath, 'rb') as f:
                files = {'file': (filepath.name, f, 'application/octet-stream')}
                response = requests.post(
                    self.config['apps_script_url'],
                    files=files,
                    timeout=60
                )
            
            if response.status_code == 200:
                result = response.json() if response.headers.get('content-type', '').startswith('application/json') else {'message': response.text}
                print(f"✅ Successfully uploaded to Google Sheets!")
                if 'message' in result:
                    print(f"   {result['message']}")
                return True
            else:
                print(f"❌ Upload failed with status code: {response.status_code}")
                print(f"   Response: {response.text}")
                return False
                
        except requests.exceptions.RequestException as e:
            print(f"❌ Error uploading to Google Sheets: {str(e)}")
            return False
        except Exception as e:
            print(f"❌ Unexpected error during upload: {str(e)}")
            return False
    
    def run(self, report_name):
        """Main execution flow."""
        try:
            self.start_browser()
            
            # Login
            if not self.login_to_hibob():
                return False
            
            # Navigate to performance cycles
            if not self.navigate_to_performance_cycles():
                return False
            
            # Search for report
            report_element = self.search_for_report(report_name)
            if not report_element:
                return False
            
            # Download report
            filepath = self.download_report(report_element)
            if not filepath:
                return False
            
            # Upload to Google Sheets
            if not self.upload_to_google_sheets(filepath):
                return False
            
            print("\n🎉 Process completed successfully!")
            return True
            
        except KeyboardInterrupt:
            print("\n❌ Process interrupted by user")
            return False
        except Exception as e:
            print(f"\n❌ Unexpected error: {str(e)}")
            return False
        finally:
            print("\n🧹 Cleaning up...")
            self.close_browser()


def main():
    """Main entry point."""
    print("=" * 60)
    print("  HiBob Performance Report Downloader")
    print("=" * 60)
    print()
    
    # Get report name from user
    report_name = input("📝 Enter the report name (or partial name) to search for: ").strip()
    
    if not report_name:
        print("❌ Report name cannot be empty")
        sys.exit(1)
    
    print()
    
    # Run the downloader
    downloader = HiBobReportDownloader()
    success = downloader.run(report_name)
    
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()

