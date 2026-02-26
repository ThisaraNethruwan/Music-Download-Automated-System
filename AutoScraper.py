import os
import os.path
import base64
import re
import uuid
import time
from datetime import datetime

# Google API Imports
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from bs4 import BeautifulSoup

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException

# SQL Server Import
import pyodbc

#     CONFIGURATION    
# Defines the Google API access scopes required for Gmail and Google Sheets
SCOPES = [
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/spreadsheets'
]

# Google Sheets Configuration
SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'
SHEET_NAME = 'Sheet_Name'
SIGNATURE_SHEET = 'AppSignature'

# Web Portal Login Configuration
LOGIN_USERNAME = 'YOUR_LOGIN_USERNAME'
LOGIN_PASSWORD = 'YOUR_LOGIN_PASSWORD'
SENDER_EMAIL = 'YOUR_SENDER_EMAIL'

# SQL Server Database Configuration
SQL_SERVER = r'YOUR_SERVER_NAME'
SQL_DATABASE = 'YOUR_DATABASE_NAME'
SQL_TRUSTED_CONNECTION = True 

# Local Download Configuration
BASE_DOWNLOAD_DIR = os.path.join(os.getcwd(), 'downloads')

#     DATABASE FUNCTIONS    

def get_db_connection():
    # Establishes and returns a connection to the SQL Server database.
    try:
        if SQL_TRUSTED_CONNECTION:
            conn_str = (
                f'DRIVER={{ODBC Driver 18 for SQL Server}};'
                f'SERVER={SQL_SERVER};'
                f'DATABASE={SQL_DATABASE};'
                f'Trusted_Connection=yes;'
                f'TrustServerCertificate=yes;'
                f'Encrypt=no;'
            )
        else:
            conn_str = (
                f'DRIVER={{ODBC Driver 18 for SQL Server}};'
                f'SERVER={SQL_SERVER};'
                f'DATABASE={SQL_DATABASE};'
                f'TrustServerCertificate=yes;'
                f'Encrypt=no;'
            )
        
        connection = pyodbc.connect(conn_str)
        print('Database connection established')
        return connection
    except pyodbc.Error as e:
        print(f' Database connection error: {e}')
        return None

def check_track_exists_in_db(connection, title, artist):
    # Queries the database to check if a specific track by an artist is already logged.
    if not connection: return 'Error'
    try:
        cursor = connection.cursor()
        query = "SELECT COUNT(*) FROM Tracks WHERE LOWER(TrackTitle) = LOWER(?) AND LOWER(Artist) = LOWER(?)"
        cursor.execute(query, (title.strip(), artist.strip()))
        count = cursor.fetchone()[0]
        cursor.close()
        return 'Yes' if count > 0 else 'No'
    except pyodbc.Error as e:
        print(f'Database query error: {e}')
        return 'Error'

def close_db_connection(connection):
    # Safely closes the database connection if it is currently open.
    if connection:
        try:
            connection.close()
            print('Database connection closed')
        except: pass

#     AUTHENTICATION FUNCTIONS    

def authenticate():
    # Handles Google OAuth2 authentication flow and returns initialized Gmail and Sheets service objects.
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds), build('sheets', 'v4', credentials=creds)

#     GMAIL FUNCTIONS    

def get_unread_emails_from_sender(gmail_service, sender_email):
    # Fetches all unread emails from a specific sender using the Gmail API.
    try:
        query = f'from:{sender_email} is:unread'
        results = gmail_service.users().messages().list(userId='me', q=query).execute()
        messages = results.get('messages', [])
        while 'nextPageToken' in results:
            page_token = results['nextPageToken']
            results = gmail_service.users().messages().list(userId='me', q=query, pageToken=page_token).execute()
            messages.extend(results.get('messages', []))
        return messages
    except HttpError: return []

def mark_email_as_read(gmail_service, msg_id):
    # Removes the 'UNREAD' label from a specific email to prevent processing it again.
    try:
        gmail_service.users().messages().modify(userId='me', id=msg_id, body={'removeLabelIds': ['UNREAD']}).execute()
        print(f'Marked email {msg_id[:8]}... as read')
        return True
    except HttpError: return False

def get_email_body(gmail_service, msg_id):
    # Decodes and extracts the raw HTML body content from a specific email message.
    try:
        message = gmail_service.users().messages().get(userId='me', id=msg_id, format='full').execute()
        payload = message['payload']
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/html':
                    data = part['body'].get('data')
                    if data: return base64.urlsafe_b64decode(data).decode('utf 8')
                if 'parts' in part:
                    for subpart in part['parts']:
                        if subpart['mimeType'] == 'text/html':
                            data = subpart['body'].get('data')
                            if data: return base64.urlsafe_b64decode(data).decode('utf 8')
        else:
            if payload['mimeType'] == 'text/html':
                data = payload['body'].get('data')
                if data: return base64.urlsafe_b64decode(data).decode('utf 8')
        return None
    except HttpError: return None

def extract_press_play_url(html_content):
    # Parses the email HTML to find and extract the destination URL hidden behind a 'Get Now' button/link.
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        press_play_link = soup.find('a', string=re.compile(r'Get Now', re.IGNORECASE))
        if press_play_link and press_play_link.get('href'): return press_play_link['href']
        for link in soup.find_all('a'):
            if link.get_text() and 'Get Now' in link.get_text().lower():
                if link.get('href'): return link['href']
        return None
    except Exception: return None

#     SELENIUM SETUP    

def setup_selenium_driver(download_folder=None):
    # Configures and launches a headless compatible Chrome WebDriver with automatic download preferences.
    chrome_options = Options()
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    
    if download_folder:
        if not os.path.exists(download_folder):
            os.makedirs(download_folder)
            
        # Sets Chrome preferences to download files directly to the target folder without prompting.
        prefs = {
            "download.default_directory": os.path.abspath(download_folder),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_settings.popups": 0,
            "profile.default_content_setting_values.automatic_downloads": 1 
        }
        chrome_options.add_experimental_option("prefs", prefs)

    try:
        driver = webdriver.Chrome(options=chrome_options)
        return driver
    except Exception as e:
        print(f'✗ Error setting up Chrome driver: {e}')
        return None

#  HELPER: WAIT FOR DOWNLOADS 

def wait_for_downloads_to_finish(download_folder, timeout=300):
    # Pauses script execution until all temporary Chrome download files (.crdownload) disappear from the folder.
    print("  Waiting for downloads to finalize...")
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < timeout:
        time.sleep(1)
        dl_wait = False
        if not os.path.exists(download_folder):
            return 
            
        files = os.listdir(download_folder)
        if not files:
            dl_wait = True # Wait if folder is empty initially
            
        for fname in files:
            if fname.endswith('.crdownload') or fname.endswith('.tmp'):
                dl_wait = True
                
        if dl_wait:
            seconds += 1
            if seconds % 10 == 0:
                print(f"    ...still downloading ({seconds}s)")
    
    if seconds >= timeout:
        print("   Timed out waiting for downloads.")
    else:
        print("  All downloads finished.")

#  STEP 1: SCRAPE 

def scrape_and_check_tracks(press_play_url, db_connection):
    # Navigates to the extracted URL, logs in if required, and scrapes the track names and artists from the DOM.
    driver = None
    tracks_data = []
    try:
        print(f'  → Setting up browser for scraping...')
        driver = setup_selenium_driver()
        if not driver: return []

        print(f'  → Loading page: {press_play_url}')
        driver.get(press_play_url)
        time.sleep(5) 

        # Attempts to find login fields and authenticate using the configured credentials.
        try:
            inputs = driver.find_elements(By.TAG_NAME, "input")
            user_in = None
            pass_in = None
            for inp in inputs:
                ph = inp.get_attribute("placeholder") or ""
                nm = inp.get_attribute("name") or ""
                if "user" in ph.lower() or "user" in nm.lower(): user_in = inp
                if "pass" in ph.lower() or "pass" in nm.lower(): pass_in = inp
            
            if user_in and pass_in:
                print("  → Logging in...")
                user_in.clear()
                user_in.send_keys(LOGIN_USERNAME)
                pass_in.clear()
                pass_in.send_keys(LOGIN_PASSWORD)
                time.sleep(1)
                
                btns = driver.find_elements(By.TAG_NAME, "button")
                clicked = False
                for btn in btns:
                    if btn.get_attribute("type") == "submit" and btn.is_displayed():
                        btn.click()
                        clicked = True
                        break
                if not clicked: pass_in.submit()
                
                time.sleep(5)
                if "dashboard" in driver.current_url and driver.current_url != press_play_url:
                    driver.get(press_play_url)
                    time.sleep(5)
        except Exception: pass

        # Scrapes the primary track list from the webpage layout elements.
        print("  → Scanning track list...")
        rows = driver.find_elements(By.CSS_SELECTOR, "div[class*='row'], div[class*='track'], li, tr")
        found_tracks = False
        for row in rows:
            text = row.text.strip()
            if re.search(r'\d{1,2}:\d{2}', text):
                lines = text.split('\n')
                clean = [l.strip() for l in lines if not re.match(r'^\d+$', l.strip()) and not re.match(r'^\d{1,2}:\d{2}$', l.strip()) and l.strip()]
                if len(clean) >= 2:
                    title = clean[0]
                    artist = clean[1]
                    db_status = check_track_exists_in_db(db_connection, title, artist)
                    tracks_data.append({'title': title, 'artist': artist, 'db_status': db_status})
                    print(f"    • {title} - {artist} [DB: {db_status}]")
                    found_tracks = True

        if found_tracks: return tracks_data

        # Fallback method: Extracts track names and artists via regex from the raw body text if DOM scraping fails.
        print("   DOM Scan failed. Falling back to text scrape.")
        body = driver.find_element(By.TAG_NAME, "body").text
        pattern = re.compile(r'(.+?)\n(.+?)\n.*?(\d{1,2}:\d{2})', re.MULTILINE)
        matches = pattern.findall(body)
        for m in matches:
            t = m[0].strip()
            a = m[1].strip()
            if re.match(r'^\d+$', t) or re.match(r'^\d{1,2}:\d{2}$', t): continue
            stat = check_track_exists_in_db(db_connection, t, a)
            tracks_data.append({'title': t, 'artist': a, 'db_status': stat})
            print(f"    • {t} - {a} [DB: {stat}]")
        return tracks_data
    except Exception as e:
        print(f'  ✗ Scraper Error: {e}')
        return []
    finally:
        if driver: driver.quit()

# STEP 2: DOWNLOAD

def trigger_download_wav(driver, row_element):
    # Finds the track context menu and triggers the 'Download WAV' option specifically.
    wait = WebDriverWait(driver, 10)
    try:
        potential_buttons = row_element.find_elements(By.CSS_SELECTOR, "button, svg, [role='button'], .cursor-pointer, i")
        menu_btn = potential_buttons[-1] if potential_buttons else None

        if not menu_btn:
            print("      Could not find menu button (dots).")
            return False

        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", menu_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", menu_btn)
        
        xpath_wav = "//div[contains(., 'Download WAV')] | //li[contains(., 'Download WAV')] | //button[contains(., 'Download WAV')]"
        download_option = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_wav)))
        
        print("       Found 'Download WAV', clicking...")
        driver.execute_script("arguments[0].click();", download_option)
        time.sleep(3) 
        return True

    except Exception as e:
        print(f"       Error triggering download: {e}")
        try: driver.find_element(By.TAG_NAME, "body").click()
        except: pass
        return False

def download_tracks_from_sheet(sheets_service, press_play_url, unique_id):
    # Reads the Google Sheet to see which scraped tracks aren't in the DB, then orchestrates downloading them.
    driver = None
    downloaded_count = 0
    download_path = os.path.join(BASE_DOWNLOAD_DIR, unique_id)
    
    try:
        # Reads the previously scraped data from Google Sheets based on the email's unique ID.
        print(f"\n  → Reading sheet (Unique ID: {unique_id})...")
        result = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f'{SHEET_NAME}!A:E').execute()
        rows = result.get('values', [])
        
        tracks_to_download = []
        for row in rows[1:]:
            if len(row) >= 5 and row[0] == unique_id and row[4] == 'No':
                tracks_to_download.append({'title': row[2], 'artist': row[3]})
        
        if not tracks_to_download: return 0
        
        print(f"  → Found {len(tracks_to_download)} tracks to download")
        
        driver = setup_selenium_driver(download_folder=download_path)
        if not driver: return 0

        driver.get(press_play_url)
        time.sleep(8) 

        # Re-authenticates to the portal prior to initiating downloads.
        try:
            inputs = driver.find_elements(By.TAG_NAME, "input")
            user_in = None
            pass_in = None
            for inp in inputs:
                ph = inp.get_attribute("placeholder") or ""
                nm = inp.get_attribute("name") or ""
                if "user" in ph.lower() or "user" in nm.lower(): user_in = inp
                if "pass" in ph.lower() or "pass" in nm.lower(): pass_in = inp
            
            if user_in and pass_in and user_in.is_displayed():
                user_in.clear()
                user_in.send_keys(LOGIN_USERNAME)
                pass_in.clear()
                pass_in.send_keys(LOGIN_PASSWORD)
                time.sleep(1)
                btns = driver.find_elements(By.TAG_NAME, "button")
                clicked = False
                for btn in btns:
                    if btn.get_attribute("type") == "submit" and btn.is_displayed():
                        btn.click()
                        clicked = True
                        break
                if not clicked: pass_in.submit()
                time.sleep(8)
                if "dashboard" in driver.current_url and driver.current_url != press_play_url:
                    driver.get(press_play_url)
                    time.sleep(8)
        except Exception: pass

        # Iterates through the needed tracks and clicks the download buttons on the webpage.
        print("  → Starting downloads...")
        rows = driver.find_elements(By.CSS_SELECTOR, "div[class*='row'], div[class*='track'], li, tr")
        
        for track in tracks_to_download:
            target_title = track['title']
            target_artist = track['artist']
            print(f"   Processing: {target_title}")
            
            found_row = False
            for row in rows:
                try:
                    text = row.text.strip()
                    if target_title in text and target_artist in text:
                        if trigger_download_wav(driver, row):
                             downloaded_count += 1
                             time.sleep(2) 
                        found_row = True
                        break
                except StaleElementReferenceException:
                    rows = driver.find_elements(By.CSS_SELECTOR, "div[class*='row'], div[class*='track'], li, tr")
                    break
            
            if not found_row:
                print(f"    Could not find row visible for this track")
        
        if downloaded_count > 0:
            wait_for_downloads_to_finish(download_path)
            
        return downloaded_count

    except Exception as e:
        print(f'  Download Process Error: {e}')
        return downloaded_count
    finally:
        if driver: 
            driver.quit()
            print("   Browser closed")

# SHEET FUNCTIONS

def ensure_signature_sheet_exists(sheets_service):
    # Verifies the logging tab exists in Google Sheets, creating it with headers if missing.
    try:
        sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f'{SIGNATURE_SHEET}!A1').execute()
    except HttpError:
        try:
            req = {'requests': [{'addSheet': {'properties': {'title': SIGNATURE_SHEET}}}]}
            sheets_service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=req).execute()
            headers = [['Run Number', 'Timestamp', 'Emails Processed', 'URLs Added', 'Tracks Extracted', 'Downloads']]
            sheets_service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=f'{SIGNATURE_SHEET}!A1', valueInputOption='RAW', body={'values': headers}).execute()
        except: pass

def ensure_main_sheet_has_headers(sheets_service):
    # Checks the primary sheet for the correct header row and injects it if it's blank or incorrect.
    try:
        res = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f'{SHEET_NAME}!A1:F1').execute()
        expected = ['Unique ID', 'URL', 'Title', 'Artist', 'In DB', 'Path']
        if not res.get('values') or res['values'][0] != expected:
            body = {'values': [expected]}
            sheets_service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=f'{SHEET_NAME}!A1', valueInputOption='RAW', body=body).execute()
    except: pass

def get_next_row_number(sheets_service):
    # Finds the first empty row in the main Google Sheet to append new data to.
    try:
        res = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f'{SHEET_NAME}!A:A').execute()
        return len(res.get('values', [])) + 1
    except: return 2

def append_tracks_to_sheet(sheets_service, row_number, unique_id, url, tracks):
    # Writes the scraped track list and database status directly into the Google Sheet.
    try:
        values = []
        if not tracks:
            values = [[unique_id, url, '', '', '', '']]
        else:
            for t in tracks:
                path_val = 'No' if t['db_status'] == 'Yes' else ''
                values.append([
                    unique_id, 
                    url, 
                    t['title'], 
                    t['artist'], 
                    t['db_status'],
                    path_val
                ])
        
        body = {'values': values}
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, 
            range=f'{SHEET_NAME}!A{row_number}', 
            valueInputOption='RAW', 
            body=body
        ).execute()
        print(f"  Added {len(values)} rows to sheet starting at row {row_number}")
        return len(values)
    except Exception as e:
        print(f"  Error updating sheet: {e}")
        return 0

def update_sheet_with_paths(sheets_service, unique_id, download_folder):
    # Scans the local download folder to match downloaded files to rows and updates the 'Path' column.
    try:
        print(f"\n  Updating file paths in sheet for ID: {unique_id}...")
        
        if not os.path.exists(download_folder):
            print("     Download folder not found.")
            return

        # Get absolute paths of completed files.
        files = [f for f in os.listdir(download_folder) if not f.endswith('.crdownload') and not f.endswith('.tmp')]
        
        if not files:
            print("     No files found in folder.")
            return

        result = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f'{SHEET_NAME}!A:F').execute()
        rows = result.get('values', [])
        
        for i, row in enumerate(rows):
            if i == 0: continue
            
            if len(row) >= 5 and row[0] == unique_id and row[4] == 'No':
                title = row[2]
                matched_file = None
                
                # Loose matching to find the associated file name.
                clean_title = re.sub(r'[\\/*?:"<>|]', "", title).strip()
                for fname in files:
                    if clean_title.lower() in fname.lower():
                        matched_file = fname
                        break
                
                if matched_file:
                    # Constructs the absolute path and updates Google Sheets.
                    full_path = os.path.join(os.path.abspath(download_folder), matched_file)
                    
                    sheet_row = i + 1
                    body = {'values': [[full_path]]}
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=SPREADSHEET_ID,
                        range=f'{SHEET_NAME}!F{sheet_row}',
                        valueInputOption='RAW',
                        body=body
                    ).execute()
                    print(f"    Updated path: {matched_file}")
                else:
                    print(f"    Could not match file for track: {title}")

    except Exception as e:
        print(f"  Error updating paths: {e}")

def get_existing_urls(sheets_service):
    # Fetches all previously processed URLs from the Google Sheet to prevent redundant scrapes.
    try:
        res = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f'{SHEET_NAME}!B:B').execute()
        vals = res.get('values', [])
        return [row[0] for row in vals[1:] if row]
    except: return []

def log_app_run(sheets_service, run_num, emails, urls, tracks, downloads):
    # Appends a summary timestamp and execution stats to the signature log sheet.
    try:
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        values = [[run_num, ts, emails, urls, tracks, downloads]]
        sheets_service.spreadsheets().values().append(spreadsheetId=SPREADSHEET_ID, range=f'{SIGNATURE_SHEET}!A:F', valueInputOption='RAW', body={'values': values}).execute()
    except: pass

def get_last_run_number(sheets_service):
    # Retrieves the last execution run number from the log sheet to increment it.
    try:
        res = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f'{SIGNATURE_SHEET}!A:A').execute()
        vals = res.get('values', [])
        return int(vals[-1][0]) if len(vals) > 1 else 0
    except: return 0

# MAIN 

def main():
    # Orchestrates the full lifecycle: Fetch emails -> Scrape links -> Verify DB -> Write to Sheets -> Download -> Update Paths.
    print('=' * 70)
    print('Email to Music Downloader - V3.0')
    print('=' * 70)
    
    gmail, sheets = authenticate()
    db_conn = get_db_connection()
    if not db_conn: return
    
    ensure_signature_sheet_exists(sheets)
    ensure_main_sheet_has_headers(sheets)
    
    current_run = get_last_run_number(sheets) + 1
    messages = get_unread_emails_from_sender(gmail, SENDER_EMAIL)
    existing_urls = get_existing_urls(sheets)
    next_row = get_next_row_number(sheets)
    stats = {'processed': 0, 'tracks': 0, 'downloaded': 0}
    
    for msg in messages:
        msg_id = msg['id']
        print(f"\n Processing Email ID: {msg_id[:12]}...")
        
        body = get_email_body(gmail, msg_id)
        if not body: continue
            
        url = extract_press_play_url(body)
        
        if url:
            if url in existing_urls:
                print("   Skipping duplicate URL.")
                mark_email_as_read(gmail, msg_id)
                continue
            
            print(f"   Found URL: {url}")
            unique_id = str(uuid.uuid4())[:8]
            
            # STEP 1: Scrape track info to see what we actually need
            tracks = scrape_and_check_tracks(url, db_conn)
            if not tracks:
                mark_email_as_read(gmail, msg_id)
                continue
            
            # STEP 2: Write tracks and DB statuses to Google Sheets
            rows_added = append_tracks_to_sheet(sheets, next_row, unique_id, url, tracks)
            if rows_added == 0: continue
            
            next_row += rows_added
            existing_urls.append(url)
            stats['processed'] += 1
            stats['tracks'] += len(tracks)
            
            # STEP 3: Read back from Google Sheets to trigger the actual Downloads
            downloaded = download_tracks_from_sheet(sheets, url, unique_id)
            stats['downloaded'] += downloaded
            
            # STEP 4: Once downloads finish, tie the local paths back to the Sheet
            if downloaded > 0:
                final_dir = os.path.join(BASE_DOWNLOAD_DIR, unique_id)
                if os.path.exists(final_dir):
                    update_sheet_with_paths(sheets, unique_id, final_dir)
                    files = [f for f in os.listdir(final_dir) if not f.endswith('.crdownload')]
                    print(f"   Complete! Verified {len(files)} valid file(s).")
            
            mark_email_as_read(gmail, msg_id)
            
        else:
            print("   No 'Get Now' link found.")
            mark_email_as_read(gmail, msg_id)

    print(f"\nFINAL SUMMARY: {len(messages)} Emails, {stats['downloaded']} Downloads.")
    log_app_run(sheets, current_run, len(messages), stats['processed'], stats['tracks'], stats['downloaded'])
    close_db_connection(db_conn)

if __name__ == '__main__':
    main()