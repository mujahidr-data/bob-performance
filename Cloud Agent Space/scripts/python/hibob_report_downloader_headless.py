#!/usr/bin/env python3
"""
HiBob Performance Report Downloader - Headless Mode Version
Same as hibob_report_downloader.py but runs in headless mode by default
"""

import sys
import os

# Import the main downloader
sys.path.insert(0, os.path.dirname(__file__))
from hibob_report_downloader import HiBobReportDownloader

# Override to use headless mode
class HeadlessHiBobDownloader(HiBobReportDownloader):
    def start_browser(self):
        """Start Playwright browser instance in headless mode."""
        print("🌐 Starting browser (headless mode)...")
        try:
            self.playwright = __import__('playwright').sync_api.sync_playwright().start()
            
            # Launch in headless mode
            self.browser = self.playwright.chromium.launch(
                headless=True,
                args=['--disable-dev-shm-usage', '--no-sandbox']
            )
            
            self.context = self.browser.new_context(
                accept_downloads=True,
                viewport={'width': 1920, 'height': 1080}
            )
            
            self.context.set_default_timeout(30000)
            self.context.set_default_navigation_timeout(30000)
            
            self.page = self.context.new_page()
            print("✅ Browser started successfully (headless)")
                
        except Exception as e:
            print(f"❌ Failed to start browser: {str(e)}")
            raise

if __name__ == "__main__":
    print("=" * 60)
    print("  HiBob Performance Report Downloader (Headless Mode)")
    print("=" * 60)
    print()
    
    report_name = input("📝 Enter the report name (or partial name) to search for: ").strip()
    
    if not report_name:
        print("❌ Report name cannot be empty")
        sys.exit(1)
    
    print()
    
    downloader = HeadlessHiBobDownloader()
    success = downloader.run(report_name)
    
    sys.exit(0 if success else 1)

