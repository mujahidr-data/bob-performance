#!/usr/bin/env python3
"""
HiBob Performance Report Downloader - Web Interface
Run the automation from a browser instead of terminal
"""

from flask import Flask, render_template, request, jsonify, session
import threading
import time
import json
import os
import sys
from pathlib import Path

# Add scripts/python to path for imports
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "scripts" / "python"))

from hibob_report_downloader import HiBobReportDownloader
from werkzeug.serving import make_server

# Set template and static folders
app = Flask(__name__, 
            template_folder=str(PROJECT_ROOT / "web" / "templates"),
            static_folder=str(PROJECT_ROOT / "web"))
app.secret_key = os.urandom(24)  # For session management

# Store automation status
automation_status = {
    'running': False,
    'status': 'idle',
    'message': '',
    'reports': [],
    'unique_reports_data': [],  # Store serializable report data (text, normalized) only
    'selected_report': None,
    'progress': [],
    'error': None,
    'result': None
}

# Store non-serializable objects separately (not in status dict)
automation_objects = {
    'downloader': None,
    'unique_reports': []  # Store actual report elements here
}

# Lock for thread safety
status_lock = threading.Lock()


def update_status(status, message='', reports=None, unique_reports=None, error=None, result=None, downloader=None):
    """Update automation status thread-safely"""
    with status_lock:
        automation_status['status'] = status
        automation_status['message'] = message
        if reports is not None:
            automation_status['reports'] = reports
        if unique_reports is not None:
            # Store serializable data only
            automation_status['unique_reports_data'] = [
                {'text': r['text'], 'normalized': r['normalized']} 
                for r in unique_reports
            ]
            # Store actual objects separately
            automation_objects['unique_reports'] = unique_reports
        if error is not None:
            automation_status['error'] = error
        if result is not None:
            automation_status['result'] = result
        if downloader is not None:
            automation_objects['downloader'] = downloader
        automation_status['progress'].append({
            'time': time.strftime('%H:%M:%S'),
            'status': status,
            'message': message
        })


class WebHiBobDownloader(HiBobReportDownloader):
    """Extended downloader that reports status to web interface"""
    
    def __init__(self, config_path=None):
        if config_path is None:
            config_path = str(PROJECT_ROOT / "config" / "config.json")
        super().__init__(config_path)
        self.web_callback = None
    
    def log(self, message, status='info'):
        """Override to send logs to web interface"""
        update_status(status, message)
        print(f"[{status.upper()}] {message}")
    
    def start_browser(self):
        """Start browser in headless mode (background) for web interface"""
        self.log("🌐 Starting browser in background...", 'info')
        try:
            from playwright.sync_api import sync_playwright
            
            self.playwright = sync_playwright().start()
            
            # Launch in headless mode (no visible window)
            try:
                self.browser = self.playwright.chromium.launch(
                    headless=True,  # Run in background
                    args=[
                        '--disable-dev-shm-usage',
                        '--no-sandbox'
                    ]
                )
            except Exception as e:
                self.log(f"⚠️  Headless launch failed, trying with minimal args: {str(e)}", 'warning')
                # Fallback
                self.browser = self.playwright.chromium.launch(headless=True)
            
            self.context = self.browser.new_context(
                accept_downloads=True,
                viewport={'width': 1920, 'height': 1080},
                user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            )
            
            # Set longer timeouts
            self.context.set_default_timeout(60000)
            self.context.set_default_navigation_timeout(60000)
            
            self.page = self.context.new_page()
            
            # Verify browser is running
            try:
                self.page.goto("about:blank", timeout=10000)
                self.log("✅ Browser started successfully (running in background)", 'success')
            except Exception as e:
                self.log(f"⚠️  Browser verification warning: {str(e)}", 'warning')
                
        except Exception as e:
            self.log(f"❌ Failed to start browser: {str(e)}", 'error')
            raise
    
    def login_to_hibob(self):
        """Login with web logging"""
        self.log("🔐 Logging in to HiBob...", 'info')
        result = super().login_to_hibob()
        if result:
            self.log("✅ Successfully logged in", 'success')
        else:
            self.log("❌ Login failed", 'error')
        return result
    
    def navigate_to_performance_cycles(self):
        """Navigate with web logging"""
        self.log("📊 Navigating to Performance Cycles...", 'info')
        result = super().navigate_to_performance_cycles()
        if result:
            self.log("✅ On Performance Cycles page", 'success')
        return result
    
    def search_for_report_web(self, report_name):
        """Search for reports and return list for web selection"""
        self.log(f"🔍 Searching for report: '{report_name}'", 'info')
        
        try:
            # Verify page is loaded and browser is connected
            if not self.page or self.page.is_closed():
                self.log("❌ Browser page is closed", 'error')
                return [], []
            
            # Check current URL
            current_url = self.page.url
            self.log(f"   Current URL: {current_url}", 'info')
            
            # Verify we're on the right page
            if 'manage-cycles' not in current_url:
                self.log("❌ Not on Performance Cycles page", 'error')
                return [], []
            
            # Wait for page to fully load
            self.log("   Waiting for page to load...", 'info')
            try:
                self.page.wait_for_load_state('networkidle', timeout=10000)
            except:
                self.log("   ⚠️  Page still loading, continuing anyway...", 'warning')
            
            # Programmatic refresh to ensure page fully loads (like manual refresh)
            self.log("   🔄 Refreshing page to ensure full load...", 'info')
            try:
                self.page.reload(wait_until="networkidle", timeout=60000)
                self.log("   ✓ Page refreshed", 'success')
            except:
                # If networkidle times out, try domcontentloaded
                try:
                    self.page.reload(wait_until="domcontentloaded", timeout=60000)
                    self.log("   ✓ Page refreshed (domcontentloaded)", 'success')
                except:
                    self.log("   ⚠️  Refresh timeout, continuing anyway...", 'warning')
            
            time.sleep(3)  # Extra wait for dynamic content after refresh
            
            # Skip the search box entirely - just extract all reports from the table
            self.log("   Extracting all reports from table (skipping search box)...", 'info')
            
            # Get all reports (same logic as parent class)
            report_name_normalized = self.normalize_search_term(report_name)
            
            # Debug: Let's see what we can find
            self.log("   Debugging: Looking for table structure...", 'info')
            
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
                    self.log(f"      {name} ({selector}): {len(elements)} elements", 'info')
                except:
                    self.log(f"      {name} ({selector}): error", 'warning')
            
            # Now try to get actual report rows
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
                            if text and 'Status' not in text and 'Name' not in text or len(text) > 50:
                                filtered.append(elem)
                    
                    if filtered and len(filtered) > 0:
                        report_elements = filtered
                        self.log(f"   ✓ Found {len(filtered)} visible rows with: {selector}", 'success')
                        break
                except Exception as e:
                    self.log(f"   ✗ Selector '{selector}' failed: {str(e)}", 'warning')
                    continue
            
            if not report_elements:
                self.log("   ⚠️ No table rows found with any selector", 'warning')
            
            all_reports = []
            for element in report_elements:
                try:
                    name_cell = None
                    cells = element.query_selector_all('td, div[role="cell"]')
                    if len(cells) >= 2:
                        name_cell = cells[1]
                    
                    if name_cell:
                        text = name_cell.inner_text().strip()
                    else:
                        full_text = element.inner_text().strip()
                        lines = [line.strip() for line in full_text.split('\n') if line.strip()]
                        text = lines[1] if len(lines) >= 2 else full_text
                    
                    if len(text) < 3:
                        continue
                    text_lower = text.lower()
                    if text_lower in ['status', 'name', 'participants', 'reviews completion', 'by condition', 'by name']:
                        continue
                    if text.replace('/', '').replace(' ', '').isdigit():
                        continue
                    if text in ['Draft', 'Pending launch', 'Running', 'Stopped', 'Ended']:
                        continue
                    
                    all_reports.append({
                        'element': element,
                        'text': text,
                        'normalized': self.normalize_search_term(text)
                    })
                except:
                    continue
            
            # Remove duplicates
            seen_texts = set()
            unique_reports = []
            for report in all_reports:
                if report['normalized'] not in seen_texts and len(report['text']) > 3:
                    seen_texts.add(report['normalized'])
                    unique_reports.append(report)
            
            # Score matches
            matching_reports = []
            for report in unique_reports:
                score = 0
                report_text_lower = report['text'].lower()
                report_normalized = report['normalized']
                
                if report_name.lower() in report_text_lower:
                    score = 100
                elif report_name_normalized in report_normalized:
                    score = 80
                elif any(word in report_normalized for word in report_name_normalized.split() if len(word) > 1):
                    score = 50
                
                if score > 0:
                    matching_reports.append({
                        'element': report['element'],
                        'text': report['text'],
                        'score': score
                    })
            
            matching_reports.sort(key=lambda x: x['score'], reverse=True)
            
            # If no matches, use all unique reports
            if not matching_reports:
                matching_reports = [{'element': r['element'], 'text': r['text'], 'score': 0} 
                                   for r in unique_reports]
            
            # Return reports for web selection
            reports_list = [{'id': i+1, 'name': r['text'], 'score': r['score']} 
                          for i, r in enumerate(matching_reports)]
            
            # Build unique_reports with elements for downloading
            unique_reports_with_elements = [
                {'element': r['element'], 'text': r['text'], 'normalized': self.normalize_search_term(r['text'])}
                for r in matching_reports
            ]
            
            self.log(f"✅ Found {len(reports_list)} reports", 'success')
            return reports_list, unique_reports_with_elements
            
        except Exception as e:
            self.log(f"❌ Error searching: {str(e)}", 'error')
            return [], []


def run_automation(report_name, selected_index=None):
    """Run automation in background thread"""
    global automation_status
    
    with status_lock:
        automation_status['running'] = True
        automation_status['status'] = 'running'
        automation_status['message'] = 'Initializing...'
        automation_status['reports'] = []
        automation_status['error'] = None
        automation_status['result'] = None
        automation_status['progress'] = []
    
    # Update status immediately to show we're starting
    update_status('running', 'Starting automation...')
    
    downloader = None
    
    try:
        config_path = str(PROJECT_ROOT / "config" / "config.json")
        downloader = WebHiBobDownloader(config_path)
        
        # Start browser with error handling
        update_status('running', 'Starting browser...')
        try:
            downloader.start_browser()
        except Exception as e:
            update_status('error', 'Browser failed to start', error=f'Browser error: {str(e)}')
            return
        
        # Verify browser is running
        if not downloader.page or downloader.page.is_closed():
            update_status('error', 'Browser not connected', error='Browser page is not available')
            return
        
        # Login with retry
        update_status('running', 'Logging in...')
        try:
            if not downloader.login_to_hibob():
                update_status('error', 'Login failed', error='Login to HiBob failed')
                return
        except Exception as e:
            update_status('error', 'Login error', error=f'Login exception: {str(e)}')
            return
        
        # Navigate with verification
        update_status('running', 'Navigating to Performance Cycles...')
        try:
            if not downloader.navigate_to_performance_cycles():
                update_status('error', 'Navigation failed', error='Failed to navigate to Performance Cycles')
                return
        except Exception as e:
            update_status('error', 'Navigation error', error=f'Navigation exception: {str(e)}')
            return
        
        # Search for reports with error handling
        update_status('running', 'Searching for reports...')
        try:
            reports_list, unique_reports = downloader.search_for_report_web(report_name)
            
            if not reports_list:
                update_status('error', 'No reports found', error='Could not find any reports on the page')
                return
            
            # Store unique_reports and downloader for later use
            update_status('selecting', 'Please select a report', reports=reports_list, 
                         unique_reports=unique_reports, downloader=downloader)
        except Exception as e:
            update_status('error', 'Search error', error=f'Search exception: {str(e)}')
            return
        
        # Wait for user selection
        if selected_index is None:
            # Wait for selection via API
            max_wait = 300  # 5 minutes
            waited = 0
            while waited < max_wait:
                time.sleep(1)
                waited += 1
                with status_lock:
                    if automation_status['selected_report'] is not None:
                        selected_index = automation_status['selected_report']
                        break
                # Check if still running
                with status_lock:
                    if not automation_status['running']:
                        return
            
            if selected_index is None:
                update_status('error', 'No report selected', error='Timeout waiting for report selection')
                return
        
        # Get unique_reports from stored objects
        with status_lock:
            unique_reports = automation_objects.get('unique_reports', unique_reports)
        
        # Download selected report
        update_status('downloading', f'Downloading report {selected_index}...')
        if 0 <= selected_index - 1 < len(unique_reports):
            report_element = unique_reports[selected_index - 1]['element']
            filepath = downloader.download_report(report_element)
        else:
            update_status('error', 'Invalid report index', error=f'Report index {selected_index} is out of range')
            return
        
        if not filepath:
            update_status('error', 'Download failed', error='Failed to download report')
            return
        
        # Upload to Google Sheets
        update_status('uploading', 'Uploading to Google Sheets...')
        success = downloader.upload_to_google_sheets(filepath)
        
        if success:
            update_status('success', '✅ Process completed successfully!', 
                        result=f'Report uploaded to Google Sheets: {filepath}')
        else:
            update_status('error', '❌ Upload failed to Google Sheets', error='Failed to upload to Google Sheets')
        
    except Exception as e:
        update_status('error', f'❌ Error: {str(e)}', error=str(e))
    finally:
        # Check if error occurred before cleanup
        error_occurred = False
        success_occurred = False
        with status_lock:
            current_status = automation_status.get('status', 'idle')
            if current_status == 'error':
                error_occurred = True
            elif current_status == 'success':
                success_occurred = True
        
        # Clean up browser - ensure it's fully closed
        if downloader:
            try:
                # Close all pages first
                if hasattr(downloader, 'page') and downloader.page and not downloader.page.is_closed():
                    downloader.page.close()
                # Close browser context
                if hasattr(downloader, 'context') and downloader.context:
                    downloader.context.close()
                # Close browser
                if hasattr(downloader, 'close_browser'):
                    downloader.close_browser()
                # Also try to close playwright instance
                if hasattr(downloader, 'playwright') and downloader.playwright:
                    try:
                        downloader.playwright.stop()
                    except:
                        pass
            except Exception as e:
                # Log cleanup error but don't overwrite status
                print(f"Browser cleanup error: {e}")
        
        # Update status lock - preserve error/success status
        with status_lock:
            automation_status['running'] = False
            automation_objects['downloader'] = None
            automation_objects['unique_reports'] = []
            # Preserve error or success status - don't change to idle
            if not error_occurred and not success_occurred:
                # Only change to idle if we're not in error or success state
                if automation_status.get('status') not in ['error', 'success']:
                    automation_status['status'] = 'idle'
                    automation_status['message'] = 'Ready to start. Enter a report name and click "Start Automation".'


@app.route('/')
def index():
    """Main page"""
    return render_template('index.html')


@app.route('/api/status', methods=['GET'])
def get_status():
    """Get current automation status"""
    with status_lock:
        return jsonify(automation_status)


@app.route('/api/start', methods=['POST'])
def start_automation():
    """Start automation"""
    data = request.json
    report_name = data.get('report_name', '').strip()
    
    if not report_name:
        return jsonify({'error': 'Report name is required'}), 400
    
    if automation_status['running']:
        return jsonify({'error': 'Automation already running'}), 400
    
    # Immediately update status to 'running' before starting thread
    with status_lock:
        automation_status['running'] = True
        automation_status['status'] = 'running'
        automation_status['message'] = 'Starting automation...'
        automation_status['error'] = None
        automation_status['progress'] = []
    
    # Start in background thread
    thread = threading.Thread(target=run_automation, args=(report_name,))
    thread.daemon = True
    thread.start()
    
    return jsonify({'status': 'running', 'message': 'Starting automation...'})


@app.route('/api/select', methods=['POST'])
def select_report():
    """Select a report to download"""
    data = request.json
    report_index = data.get('report_index')
    
    if report_index is None:
        return jsonify({'error': 'Report index is required'}), 400
    
    with status_lock:
        automation_status['selected_report'] = int(report_index)
    
    return jsonify({'status': 'selected', 'message': f'Report {report_index} selected'})


@app.route('/api/stop', methods=['POST'])
def stop_automation():
    """Stop automation"""
    # Clean up browser if running
    downloader = automation_objects.get('downloader')
    if downloader:
        try:
            # Close all pages first
            if downloader.page and not downloader.page.is_closed():
                downloader.page.close()
            # Close browser context
            if downloader.context:
                downloader.context.close()
            # Close browser
            downloader.close_browser()
        except:
            pass
    
    with status_lock:
        automation_status['running'] = False
        automation_status['status'] = 'stopped'
        automation_status['message'] = 'Automation stopped by user'
        automation_objects['downloader'] = None
        automation_objects['unique_reports'] = []
    
    return jsonify({'status': 'stopped'})


if __name__ == '__main__':
    import sys
    
    # Get port from command line or use default
    port = 5001
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except:
            pass
    
    print("=" * 60)
    print("  HiBob Performance Report Downloader - Web Interface")
    print("=" * 60)
    print()
    print("🌐 Starting web server...")
    print(f"📱 Open your browser and go to: http://127.0.0.1:{port}")
    print(f"   Or try: http://localhost:{port}")
    print("🛑 Press Ctrl+C to stop the server")
    print()
    
    try:
        app.run(host='127.0.0.1', port=port, debug=True, threaded=True, use_reloader=False)
    except OSError as e:
        if "Address already in use" in str(e):
            print(f"❌ Port {port} is already in use!")
            print(f"💡 Try running with a different port:")
            print(f"   python3 web_app.py {port + 1}")
        else:
            print(f"❌ Error starting server: {e}")
            raise

