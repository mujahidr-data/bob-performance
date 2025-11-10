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
    def __init__(self, config_path=None):
        """Initialize the downloader with configuration."""
        if config_path is None:
            # Try to find config.json in config/ folder or root
            project_root = Path(__file__).parent.parent.parent
            config_path = project_root / "config" / "config.json"
            if not config_path.exists():
                config_path = project_root / "config.json"
            config_path = str(config_path)
        self.config = self.load_config(config_path)
        # Downloads folder relative to project root
        project_root = Path(__file__).parent.parent.parent
        self.download_dir = project_root / "downloads"
        self.download_dir.mkdir(exist_ok=True)
        self.playwright = None
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
        try:
            self.playwright = sync_playwright().start()
            
            # Try with minimal args first (macOS compatibility)
            try:
                self.browser = self.playwright.chromium.launch(
                    headless=False,
                    args=[
                        '--disable-dev-shm-usage',
                        '--no-sandbox'
                    ]
                )
            except Exception as e:
                print(f"⚠️  Standard launch failed, trying headless mode: {str(e)}")
                # Fallback to headless if regular launch fails
                self.browser = self.playwright.chromium.launch(headless=True)
                print("✅ Browser started in headless mode")
            
            self.context = self.browser.new_context(
                accept_downloads=True,
                viewport={'width': 1920, 'height': 1080},
                # Add user agent to avoid detection
                user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            )
            
            # Add longer timeouts for page operations
            self.context.set_default_timeout(60000)  # 60 seconds
            self.context.set_default_navigation_timeout(60000)  # 60 seconds
            
            self.page = self.context.new_page()
            
            # Verify browser is actually running with a test navigation
            try:
                print("   Verifying browser connection...")
                self.page.goto("about:blank", timeout=10000)
                
                # Test if we can get page info
                title = self.page.title()
                url = self.page.url
                print(f"   ✓ Browser responsive (title: {title}, url: {url})")
                print("✅ Browser started successfully")
            except Exception as e:
                print(f"⚠️  Browser verification warning: {str(e)}")
                print("   Continuing anyway - browser might still work...")
                
        except Exception as e:
            print(f"❌ Failed to start browser: {str(e)}")
            raise
        
    def close_browser(self):
        """Close browser and cleanup."""
        try:
            if self.browser:
                self.browser.close()
        except:
            pass
        try:
            if self.playwright:
                self.playwright.stop()
        except:
            pass
            
    def login_to_hibob(self):
        """Login to HiBob via JumpCloud SSO."""
        try:
            print("🔐 Navigating to HiBob login page...")
            # Try multiple wait strategies for better page loading
            try:
                # First try networkidle (waits for network to be idle)
                self.page.goto("https://app.hibob.com/login/home", wait_until="networkidle", timeout=60000)
            except:
                # Fallback to domcontentloaded if networkidle times out
                print("   ⚠️  Network idle timeout, trying domcontentloaded...")
                try:
                    self.page.goto("https://app.hibob.com/login/home", wait_until="domcontentloaded", timeout=60000)
                except Exception as e:
                    print(f"   ⚠️  Navigation warning: {str(e)}")
                    # Try one more time with load event
                    self.page.goto("https://app.hibob.com/login/home", wait_until="load", timeout=60000)
            
            time.sleep(2)  # Give page time to fully load
            
            # Verify page loaded
            try:
                page_title = self.page.title()
                current_url = self.page.url
                print(f"   Page loaded: {page_title[:50]}...")
                print(f"   URL: {current_url[:80]}...")
            except:
                print("   ⚠️  Could not verify page load")
            
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
            
            # Wait for URL to change to JumpCloud domain or wait for page load
            try:
                # Wait for either URL change or page load
                self.page.wait_for_load_state("domcontentloaded", timeout=15000)
                time.sleep(2)  # Additional wait for dynamic content
                
                # Check if we're on JumpCloud page
                current_url = self.page.url
                print(f"   Current URL: {current_url[:80]}...")
                
                if "jumpcloud" in current_url.lower() or "sso" in current_url.lower():
                    print("   ✅ Detected JumpCloud SSO page")
                else:
                    print("   ⚠️  May not be on JumpCloud page yet, continuing...")
            except:
                print("   ⚠️  Timeout waiting for redirect, continuing anyway...")
                time.sleep(3)
            
            # On JumpCloud page - enter email again if needed
            print("🔑 Entering JumpCloud credentials...")
            try:
                # Try multiple selectors for email field on JumpCloud
                email_selectors = [
                    'input[type="email"]',
                    'input[name="email"]',
                    'input[id*="email" i]',
                    'input[placeholder*="email" i]',
                    'input[type="text"][name*="email" i]'
                ]
                email_found = False
                for selector in email_selectors:
                    try:
                        jumpcloud_email = self.page.wait_for_selector(selector, timeout=3000, state="visible")
                        if jumpcloud_email:
                            current_value = jumpcloud_email.input_value()
                            if not current_value or current_value != self.config['email']:
                                jumpcloud_email.fill(self.config['email'])
                                print("   ✅ Email entered on JumpCloud page")
                            email_found = True
                            break
                    except:
                        continue
                if not email_found:
                    print("   (Email field not found or already filled)")
            except Exception as e:
                print(f"   (Email field handling: {str(e)})")
            
            # Click Continue button on JumpCloud page (if email was entered)
            if email_found:
                print("⏭️  Clicking Continue on JumpCloud page...")
                try:
                    continue_selectors = [
                        'button:has-text("Continue")',
                        'button:has-text("CONTINUE")',
                        'button[type="submit"]',
                        'button:has-text("Next")',
                        'input[type="submit"][value*="Continue" i]',
                        'button[class*="continue" i]',
                        'button[class*="submit" i]'
                    ]
                    
                    continue_clicked = False
                    for selector in continue_selectors:
                        try:
                            continue_btn = self.page.wait_for_selector(selector, timeout=3000, state="visible")
                            if continue_btn:
                                continue_btn.click()
                                print("   ✅ Continue button clicked")
                                time.sleep(2)  # Wait for password field to appear
                                continue_clicked = True
                                break
                        except:
                            continue
                    
                    if not continue_clicked:
                        # Try pressing Enter on email field as fallback
                        print("   ⚠️  Continue button not found, trying Enter key...")
                        try:
                            jumpcloud_email.press("Enter")
                            time.sleep(2)
                        except:
                            pass
                except Exception as e:
                    print(f"   ⚠️  Error clicking Continue: {str(e)}")
                    time.sleep(2)  # Wait anyway in case it worked
            
            # Enter password - try multiple selectors and wait strategies
            print("   Looking for password field...")
            password_selectors = [
                'input[type="password"]',
                'input[name="password"]',
                'input[id*="password" i]',
                'input[placeholder*="password" i]',
                'input[type="password"][name*="pass" i]',
                '#password',
                '[data-testid*="password" i]'
            ]
            
            password_input = None
            for selector in password_selectors:
                try:
                    print(f"   Trying selector: {selector}")
                    password_input = self.page.wait_for_selector(selector, timeout=5000, state="visible")
                    if password_input:
                        print(f"   ✅ Found password field with: {selector}")
                        break
                except:
                    continue
            
            if not password_input:
                # Last resort: try to find any password input
                print("   ⚠️  Standard selectors failed, trying broader search...")
                try:
                    # Wait a bit more for page to fully render
                    time.sleep(2)
                    password_input = self.page.query_selector('input[type="password"]')
                    if not password_input:
                        # Try waiting for any input that might be password
                        self.page.wait_for_load_state("networkidle", timeout=10000)
                        password_input = self.page.query_selector('input[type="password"]')
                except:
                    pass
            
            if not password_input:
                # Take screenshot for debugging
                self.page.screenshot(path="error_password_field.png")
                print("   ❌ Could not find password field")
                print("   📸 Screenshot saved to error_password_field.png")
                print(f"   Current URL: {self.page.url}")
                print(f"   Page title: {self.page.title()}")
                raise PlaywrightTimeoutError("Password field not found")
            
            password_input.fill(self.config['password'])
            print("   ✅ Password entered")
            
            # Click sign in button - try multiple selectors
            print("✅ Signing in...")
            signin_selectors = [
                'button[type="submit"]',
                'button:has-text("Sign In")',
                'button:has-text("Log In")',
                'button:has-text("Sign in")',
                'button:has-text("Login")',
                'input[type="submit"]',
                '[type="submit"]',
                'button[class*="submit" i]',
                'button[class*="signin" i]'
            ]
            
            signin_button = None
            for selector in signin_selectors:
                try:
                    signin_button = self.page.wait_for_selector(selector, timeout=3000, state="visible")
                    if signin_button:
                        print(f"   ✅ Found sign in button with: {selector}")
                        break
                except:
                    continue
            
            if not signin_button:
                # Try pressing Enter on password field as fallback
                print("   ⚠️  Sign in button not found, trying Enter key...")
                password_input.press("Enter")
                time.sleep(2)
            else:
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
            
            # Navigate with multiple wait strategies
            try:
                # First try networkidle (waits for network to be idle)
                self.page.goto("https://app.hibob.com/performance/manage-cycles/main/cycles", 
                              wait_until="networkidle", timeout=60000)
            except:
                # Fallback to domcontentloaded if networkidle times out
                print("   ⚠️  Network idle timeout, trying domcontentloaded...")
                self.page.goto("https://app.hibob.com/performance/manage-cycles/main/cycles", 
                              wait_until="domcontentloaded", timeout=60000)
            
            # Wait for page to be interactive
            time.sleep(2)
            
            # Programmatic refresh to ensure page fully loads (like manual refresh)
            print("   🔄 Refreshing page to ensure full load...")
            try:
                self.page.reload(wait_until="networkidle", timeout=60000)
                print("   ✓ Page refreshed")
            except:
                # If networkidle times out, try domcontentloaded
                try:
                    self.page.reload(wait_until="domcontentloaded", timeout=60000)
                    print("   ✓ Page refreshed (domcontentloaded)")
                except:
                    print("   ⚠️  Refresh timeout, continuing anyway...")
            
            # Wait for page to be interactive after refresh
            time.sleep(3)
            
            # Wait for specific content to appear (table or cycles list)
            print("   Waiting for page content to load...")
            try:
                # Wait for either a table or any visible content indicating the page loaded
                self.page.wait_for_selector('table, [role="table"], [role="row"], .cycle-item, .report-item', 
                                           timeout=30000, state='visible')
                print("   ✓ Page content detected")
            except:
                print("   ⚠️  Content selector not found, but continuing...")
            
            # Extra wait for JavaScript to finish rendering
            time.sleep(3)
            
            # Verify we're on the right page
            current_url = self.page.url
            if 'manage-cycles' not in current_url:
                print(f"   ⚠️  Warning: URL doesn't contain 'manage-cycles': {current_url}")
            
            print("✅ On Performance Cycles page")
            return True
        except Exception as e:
            print(f"❌ Failed to navigate to performance cycles: {str(e)}")
            try:
                self.page.screenshot(path="error_navigation.png")
                print(f"📸 Screenshot saved to error_navigation.png")
            except:
                pass
            return False
    
    def normalize_search_term(self, text):
        """Normalize text for fuzzy matching (handles / vs &, case, etc.)"""
        # Normalize separators
        normalized = text.lower().replace('/', ' ').replace('&', ' ').replace('-', ' ')
        # Remove extra spaces
        normalized = ' '.join(normalized.split())
        return normalized
    
    def search_for_report(self, report_name):
        """Search for a report by name (case-insensitive, partial match with fuzzy matching)."""
        try:
            print(f"🔍 Searching for report: '{report_name}'")
            
            # Wait for page to load
            time.sleep(2)
            
            # Skip the search box entirely - just extract all reports from the table
            # This is more reliable than trying to find and use the search box
            print("   Extracting all reports from table (skipping search box)...")
            
            # Normalize search term for fuzzy matching
            report_name_normalized = self.normalize_search_term(report_name)
            
            # Debug: Let's see what we can find
            print("   Debugging: Looking for table structure...")
            
            # Try to find any table-like elements
            debug_selectors = {
                'table': 'table',
                'tbody': 'tbody',
                'all tr': 'tr',
                'tr with td': 'tr td',
                'role=row': '[role="row"]',
                'role=table': '[role="table"]',
            }
            
            for name, selector in debug_selectors.items():
                try:
                    elements = self.page.query_selector_all(selector)
                    print(f"      {name} ({selector}): {len(elements)} elements")
                except:
                    print(f"      {name} ({selector}): error")
            
            # Try multiple selectors to find report elements (table rows)
            possible_selectors = [
                'tbody tr',
                'tr',
                '[role="row"]',
                'table tr',
                'div[role="row"]',
            ]
            
            report_elements = []
            for selector in possible_selectors:
                try:
                    elements = self.page.query_selector_all(selector)
                    # Filter out header rows
                    filtered = []
                    for elem in elements:
                        if elem.is_visible():
                            text = elem.inner_text()
                            # Skip if it's clearly a header
                            if text and ('Status' not in text or 'Name' not in text or len(text) > 50):
                                filtered.append(elem)
                    
                    if filtered and len(filtered) > 0:
                        report_elements = filtered
                        print(f"   ✓ Found {len(filtered)} visible rows with: {selector}")
                        break
                except Exception as e:
                    print(f"   ✗ Selector '{selector}' failed: {str(e)}")
                    continue
            
            if not report_elements:
                print("   ⚠️  No table rows found with any selector")
            
            # Extract all reports - get just the name from each row
            all_reports = []
            for element in report_elements:
                try:
                    # Try to find the name cell/column specifically
                    # Look for the second column (usually the name column)
                    name_cell = None
                    name_selectors = [
                        'td:nth-child(2)',  # Second column
                        'td[class*="name" i]',
                        'td:has-text',
                        'div[class*="name" i]',
                        'span[class*="name" i]'
                    ]
                    
                    for name_sel in name_selectors:
                        try:
                            name_cell = element.query_selector(name_sel)
                            if name_cell:
                                break
                        except:
                            continue
                    
                    # If no specific name cell, get all cells and use the second one
                    if not name_cell:
                        cells = element.query_selector_all('td, div[role="cell"]')
                        if len(cells) >= 2:
                            name_cell = cells[1]  # Second cell is usually the name
                    
                    # Get text from name cell or fallback to element
                    if name_cell:
                        text = name_cell.inner_text().strip()
                    else:
                        # Fallback: get all text and try to extract just the name part
                        full_text = element.inner_text().strip()
                        # Split by newlines and take the second line (usually the name)
                        lines = [line.strip() for line in full_text.split('\n') if line.strip()]
                        if len(lines) >= 2:
                            text = lines[1]  # Second line is usually the name
                        else:
                            text = full_text
                    
                    # Skip empty or very short text
                    if len(text) < 3:
                        continue
                    # Skip header rows and common table headers
                    text_lower = text.lower()
                    if text_lower in ['status', 'name', 'participants', 'reviews completion', 'by condition', 'by name', 'draft', 'pending launch', 'running', 'stopped', 'ended']:
                        continue
                    # Skip if it's just a number (like "0/136")
                    if text.replace('/', '').replace(' ', '').isdigit():
                        continue
                    # Skip status indicators
                    if text in ['Draft', 'Pending launch', 'Running', 'Stopped', 'Ended']:
                        continue
                    
                    all_reports.append({
                        'element': element,
                        'text': text,
                        'normalized': self.normalize_search_term(text)
                    })
                except Exception as e:
                    continue
            
            # Remove duplicates based on normalized text
            seen_texts = set()
            unique_reports = []
            for report in all_reports:
                if report['normalized'] not in seen_texts and len(report['text']) > 3:
                    seen_texts.add(report['normalized'])
                    unique_reports.append(report)
            
            # If no reports found, try alternative approach
            if not unique_reports:
                print("   ⚠️  No reports found with standard method, trying alternative...")
                # Try getting all clickable elements that might be reports
                clickable_elements = self.page.query_selector_all('a, button, [role="button"], [onclick]')
                for element in clickable_elements:
                    try:
                        text = element.inner_text().strip()
                        if len(text) > 3 and text not in ['Status', 'Name', 'Participants', 'Reviews completion']:
                            unique_reports.append({
                                'element': element,
                                'text': text,
                                'normalized': self.normalize_search_term(text)
                            })
                    except:
                        continue
            
            if not unique_reports:
                print("❌ No reports found on the page")
                self.page.screenshot(path="error_no_reports.png")
                print("📸 Screenshot saved to error_no_reports.png")
                return None
            
            # Score matches (exact match > partial match > fuzzy match)
            matching_reports = []
            for report in unique_reports:
                score = 0
                report_text_lower = report['text'].lower()
                report_normalized = report['normalized']
                
                # Exact match (case-insensitive)
                if report_name.lower() in report_text_lower:
                    score = 100
                # Normalized match (handles / vs &)
                elif report_name_normalized in report_normalized:
                    score = 80
                # Partial word match
                elif any(word in report_normalized for word in report_name_normalized.split() if len(word) > 1):
                    score = 50
                
                if score > 0:
                    matching_reports.append({
                        'element': report['element'],
                        'text': report['text'],
                        'score': score
                    })
            
            # Sort by score (highest first)
            matching_reports.sort(key=lambda x: x['score'], reverse=True)
            
            # If no matches, show all available reports
            if not matching_reports:
                print(f"⚠️  No exact matches found for '{report_name}'")
                print(f"\n📋 Showing all available reports ({len(unique_reports)} total):")
                print()
                for i, report in enumerate(unique_reports, 1):
                    print(f"  [{i}] {report['text']}")
                
                while True:
                    try:
                        choice = input(f"\n📝 Enter the number of the report to download (1-{len(unique_reports)}) or 'q' to quit: ").strip()
                        if choice.lower() == 'q':
                            print("❌ Cancelled by user")
                            return None
                        choice_num = int(choice)
                        if 1 <= choice_num <= len(unique_reports):
                            selected_report = unique_reports[choice_num - 1]
                            print(f"✅ Selected: {selected_report['text']}")
                            return selected_report['element']
                        else:
                            print(f"❌ Please enter a number between 1 and {len(unique_reports)}")
                    except ValueError:
                        print("❌ Please enter a valid number or 'q' to quit")
                    except KeyboardInterrupt:
                        print("\n❌ Cancelled by user")
                        return None
            
            # Show matches (always prompt for confirmation)
            if len(matching_reports) == 1:
                print(f"✅ Found 1 matching report: {matching_reports[0]['text']}")
                print(f"\n📝 Confirm download this report?")
                confirm = input("   Enter 'y' to confirm, 'n' to see all reports, or a number to select a different one: ").strip().lower()
                if confirm == 'y':
                    return matching_reports[0]['element']
                elif confirm == 'n':
                    # Show all reports
                    print(f"\n📋 All available reports ({len(unique_reports)} total):")
                    print()
                    for i, report in enumerate(unique_reports, 1):
                        marker = " ← MATCH" if any(m['text'] == report['text'] for m in matching_reports) else ""
                        print(f"  [{i}] {report['text']}{marker}")
                    matching_reports = unique_reports
                else:
                    try:
                        choice_num = int(confirm)
                        if 1 <= choice_num <= len(unique_reports):
                            selected_report = unique_reports[choice_num - 1]
                            print(f"✅ Selected: {selected_report['text']}")
                            return selected_report['element']
                    except:
                        pass
            
            # Multiple matches - ask user to select
            print(f"\n📋 Found {len(matching_reports)} matching reports:")
            print()
            for i, report in enumerate(matching_reports, 1):
                print(f"  [{i}] {report['text']}")
            
            while True:
                try:
                    choice = input(f"\n📝 Enter the number of the report to download (1-{len(matching_reports)}) or 'a' to see all reports: ").strip()
                    if choice.lower() == 'a':
                        print(f"\n📋 All available reports ({len(unique_reports)} total):")
                        print()
                        for i, report in enumerate(unique_reports, 1):
                            marker = " ← MATCH" if any(m['text'] == report['text'] for m in matching_reports) else ""
                            print(f"  [{i}] {report['text']}{marker}")
                        choice = input(f"\n📝 Enter the number of the report to download (1-{len(unique_reports)}): ").strip()
                        choice_num = int(choice)
                        if 1 <= choice_num <= len(unique_reports):
                            selected_report = unique_reports[choice_num - 1]
                            print(f"✅ Selected: {selected_report['text']}")
                            return selected_report['element']
                    else:
                        choice_num = int(choice)
                        if 1 <= choice_num <= len(matching_reports):
                            selected_report = matching_reports[choice_num - 1]
                            print(f"✅ Selected: {selected_report['text']}")
                            return selected_report['element']
                        else:
                            print(f"❌ Please enter a number between 1 and {len(matching_reports)}")
                except ValueError:
                    print("❌ Please enter a valid number or 'a' to see all reports")
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
            
            # Click the download button first
            download_button.click()
            
            # Wait a moment for modal to appear (if it does)
            time.sleep(2)
            
            # Check if "Download Report" modal appeared
            modal_selectors = [
                '[role="dialog"]:has-text("Download Report")',
                '.modal:has-text("Download Report")',
                '[class*="modal"]:has-text("Download Report")',
                'div:has-text("Download Report")',
                'div:has-text("Select employee fields")'
            ]
            
            modal = None
            for selector in modal_selectors:
                try:
                    modal = self.page.query_selector(selector)
                    if modal and modal.is_visible():
                        print("   📋 Download Report modal detected")
                        break
                except:
                    continue
            
            download = None
            
            if modal:
                # Modal appeared - need to click Download button in modal
                print("   🔍 Looking for Download button in modal...")
                modal_download_selectors = [
                    'button:has-text("Download")',
                    'button[type="submit"]:has-text("Download")',
                    '[role="button"]:has-text("Download")',
                    'button:has-text("DOWNLOAD")',
                    'button[class*="download" i]'
                ]
                
                modal_download_button = None
                for selector in modal_download_selectors:
                    try:
                        # Look within the modal first
                        modal_download_button = modal.query_selector(selector)
                        if not modal_download_button:
                            # Try in the whole page
                            modal_download_button = self.page.query_selector(selector)
                        
                        if modal_download_button and modal_download_button.is_visible():
                            print(f"   ✓ Found Download button in modal")
                            break
                    except:
                        continue
                
                if modal_download_button:
                    # Set up download handler and click the modal download button
                    print("   ⬇️  Clicking Download button in modal...")
                    with self.page.expect_download(timeout=60000) as download_info:
                        modal_download_button.click()
                    download = download_info.value
                else:
                    print("   ⚠️  Could not find Download button in modal")
                    # Take screenshot for debugging
                    self.page.screenshot(path="error_modal_download_button.png")
                    print("   📸 Screenshot saved to error_modal_download_button.png")
            else:
                # No modal - download should have started from first click
                print("   ✓ No modal detected, waiting for download from first click...")
                try:
                    with self.page.expect_download(timeout=60000) as download_info:
                        # Download might already be in progress, wait a bit
                        time.sleep(2)
                    download = download_info.value
                except:
                    # Download might have already started, check downloads folder
                    print("   ⚠️  Download handler timeout, checking downloads folder...")
                    pass
            
            if not download:
                # If we don't have a download object, wait a bit and check downloads folder
                print("   ⚠️  Download object not available, checking downloads folder...")
                time.sleep(5)
                # List files in download directory
                import os
                files = list(self.download_dir.glob("*.xlsx")) + list(self.download_dir.glob("*.csv"))
                if files:
                    # Get the most recent file
                    latest_file = max(files, key=os.path.getctime)
                    print(f"   ✓ Found downloaded file: {latest_file}")
                    return str(latest_file)
                else:
                    print("   ❌ No download file found")
                    return None
            
            # Save the download
            filename = download.suggested_filename
            filepath = self.download_dir / filename
            download.save_as(filepath)
            
            print(f"✅ Report downloaded: {filepath}")
            return str(filepath)
            
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
        """Read Excel file and upload its content to Google Sheets using Google Sheets API."""
        try:
            print("☁️  Reading file and uploading to Google Sheets...")
            
            # Verify file exists
            if not os.path.exists(filepath):
                print(f"❌ File not found: {filepath}")
                return False
            
            # Get file size for logging
            file_size = os.path.getsize(filepath)
            filename = os.path.basename(filepath)
            print(f"   File: {filename} ({file_size:,} bytes)")
            
            # Read Excel file and convert to data
            print("   📖 Reading Excel file...")
            try:
                import pandas as pd
            except ImportError:
                print("   ⚠️  pandas not installed, trying openpyxl...")
                try:
                    import openpyxl
                except ImportError:
                    print("❌ Need pandas or openpyxl to read Excel files")
                    print("   Install with: pip install pandas openpyxl")
                    return False
            
            # Read the Excel file
            try:
                df = pd.read_excel(filepath, engine='openpyxl')
                print(f"   ✓ Read {len(df)} rows, {len(df.columns)} columns")
            except Exception as e:
                print(f"❌ Error reading Excel file: {str(e)}")
                return False
            
            # Convert DataFrame to list of lists
            df = df.fillna('')
            data = [df.columns.tolist()] + df.values.tolist()
            
            print(f"   📊 Prepared {len(data)} rows, {len(data[0]) if data else 0} columns for upload")
            
            # Use Google Sheets API
            print("   🔐 Authenticating with Google Sheets API...")
            try:
                from google.oauth2 import service_account
                from googleapiclient.discovery import build
            except ImportError:
                print("❌ Google API libraries not installed")
                print("   Install with: pip install google-auth google-api-python-client")
                return False
            
            # Google Sheets API scope
            SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
            
            # Sheet ID from the URL
            SPREADSHEET_ID = '1rnpUlOcqTpny2Pve2L82qWGI9WplOetx-1Wba7ONoeA'
            SHEET_NAME = 'Bob Perf Report'
            
            # Try service account first (better for automation)
            service_account_file = Path('service_account.json')
            creds = None
            
            if service_account_file.exists():
                print("   🔑 Using service account credentials...")
                try:
                    creds = service_account.Credentials.from_service_account_file(
                        str(service_account_file),
                        scopes=SCOPES
                    )
                    print("   ✓ Service account credentials loaded")
                except Exception as e:
                    print(f"   ⚠️  Error loading service account: {str(e)}")
                    print("   💡 Make sure service_account.json is valid and the sheet is shared with the service account email")
                    return False
            else:
                # Fallback to OAuth2 (for user credentials)
                print("   ⚠️  service_account.json not found, trying OAuth2...")
                try:
                    from google.oauth2.credentials import Credentials
                    from google_auth_oauthlib.flow import InstalledAppFlow
                    from google.auth.transport.requests import Request
                    import pickle
                    
                    token_file = Path('token.pickle')
                    
                    # Load existing token
                    if token_file.exists():
                        try:
                            with open(token_file, 'rb') as token:
                                creds = pickle.load(token)
                        except:
                            pass
                    
                    # If no valid credentials, get new ones
                    if not creds or not creds.valid:
                        if creds and creds.expired and creds.refresh_token:
                            print("   🔄 Refreshing OAuth credentials...")
                            creds.refresh(Request())
                        else:
                            print("   🔑 Requesting new OAuth credentials...")
                            print("   📋 You'll need to authorize the app in your browser")
                            
                            # Check for credentials file
                            creds_file = Path('credentials.json')
                            if not creds_file.exists():
                                print("   ❌ credentials.json not found!")
                                print("   💡 To set up Google Sheets API:")
                                print("      Option 1 (Recommended): Use Service Account")
                                print("      1. Go to: https://console.cloud.google.com/")
                                print("      2. Create a project or select existing")
                                print("      3. Enable Google Sheets API")
                                print("      4. Create Service Account")
                                print("      5. Download JSON key as service_account.json")
                                print("      6. Share the Google Sheet with the service account email")
                                print("      Option 2: Use OAuth2 (requires browser)")
                                print("      1. Create OAuth 2.0 credentials (Desktop app)")
                                print("      2. Download as credentials.json")
                                return False
                            
                            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                            creds = flow.run_local_server(port=0)
                        
                        # Save credentials for next time
                        with open(token_file, 'wb') as token:
                            pickle.dump(creds, token)
                except ImportError:
                    print("   ❌ OAuth2 libraries not available")
                    print("   💡 Install with: pip install google-auth-oauthlib")
                    return False
            
            # Build the service
            print("   🔗 Connecting to Google Sheets...")
            service = build('sheets', 'v4', credentials=creds)
            
            # Get or create the sheet
            print(f"   📋 Accessing sheet: {SHEET_NAME}")
            spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
            sheet_names = [sheet['properties']['title'] for sheet in spreadsheet.get('sheets', [])]
            
            if SHEET_NAME not in sheet_names:
                print(f"   ➕ Creating new sheet: {SHEET_NAME}")
                # Create the sheet
                request_body = {
                    'requests': [{
                        'addSheet': {
                            'properties': {
                                'title': SHEET_NAME
                            }
                        }
                    }]
                }
                service.spreadsheets().batchUpdate(
                    spreadsheetId=SPREADSHEET_ID,
                    body=request_body
                ).execute()
            else:
                print(f"   ✓ Found existing sheet: {SHEET_NAME}")
                # Clear existing data
                range_name = f"{SHEET_NAME}!A1:Z10000"
                service.spreadsheets().values().clear(
                    spreadsheetId=SPREADSHEET_ID,
                    range=range_name
                ).execute()
            
            # Write data to sheet
            print(f"   ✍️  Writing {len(data)} rows to sheet...")
            range_name = f"{SHEET_NAME}!A1"
            
            body = {
                'values': data
            }
            
            result = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()
            
            updated_cells = result.get('updatedCells', 0)
            updated_rows = result.get('updatedRows', 0)
            
            # Format header row
            print("   🎨 Formatting header row...")
            sheet_id = None
            for sheet in spreadsheet.get('sheets', []):
                if sheet['properties']['title'] == SHEET_NAME:
                    sheet_id = sheet['properties']['sheetId']
                    break
            
            if sheet_id is not None:
                requests = [
                    {
                        'repeatCell': {
                            'range': {
                                'sheetId': sheet_id,
                                'startRowIndex': 0,
                                'endRowIndex': 1
                            },
                            'cell': {
                                'userEnteredFormat': {
                                    'backgroundColor': {'red': 0.26, 'green': 0.52, 'blue': 0.96},
                                    'textFormat': {
                                        'foregroundColor': {'red': 1.0, 'green': 1.0, 'blue': 1.0},
                                        'bold': True
                                    }
                                }
                            },
                            'fields': 'userEnteredFormat(backgroundColor,textFormat)'
                        }
                    },
                    {
                        'updateSheetProperties': {
                            'properties': {
                                'sheetId': sheet_id,
                                'gridProperties': {
                                    'frozenRowCount': 1
                                }
                            },
                            'fields': 'gridProperties.frozenRowCount'
                        }
                    }
                ]
                
                service.spreadsheets().batchUpdate(
                    spreadsheetId=SPREADSHEET_ID,
                    body={'requests': requests}
                ).execute()
            
            print(f"✅ Successfully uploaded to Google Sheets!")
            print(f"   📊 Updated {updated_rows} rows, {updated_cells} cells")
            print(f"   🔗 Sheet: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid={sheet_id if sheet_id else 0}")
            
            return True
                
        except FileNotFoundError:
            print(f"❌ File not found: {filepath}")
            return False
        except PermissionError:
            print(f"❌ Permission denied reading file: {filepath}")
            return False
        except ImportError as e:
            print(f"❌ Missing required library: {str(e)}")
            print("   Install with: pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client")
            return False
        except Exception as e:
            print(f"❌ Error uploading to Google Sheets: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def run(self, report_name):
        """Main execution flow."""
        try:
            self.start_browser()
            
            # Verify browser is still running
            if not self.browser:
                print("❌ Browser failed to start")
                return False
            
            # Check if browser is connected (if method exists)
            try:
                if hasattr(self.browser, 'is_connected') and not self.browser.is_connected():
                    print("❌ Browser disconnected")
                    return False
            except:
                pass  # Method might not exist, continue anyway
            
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
            import traceback
            print("\n📋 Full error details:")
            traceback.print_exc()
            return False
        finally:
            print("\n🧹 Cleaning up...")
            try:
                self.close_browser()
            except:
                pass  # Ignore errors during cleanup


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

