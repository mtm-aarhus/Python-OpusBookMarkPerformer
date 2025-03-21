import json
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import os
from urllib.parse import urlparse, parse_qs, unquote
from OpenOrchestrator.database.queues import QueueElement
from datetime import datetime
import calendar
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
import time
from urllib.parse import urlparse, parse_qs, unquote
import win32com.client as win32
import gc
import subprocess
import random 
import string
from pebble import concurrent
from concurrent.futures import TimeoutError

# Global variables for ensuring single execution
conversion_in_progress = set()

@concurrent.process(timeout=300)
def convert_xls_to_xlsx( path: str) -> None:
    """
    Converts an .xls file to .xlsx format. Times out if the process exceeds the given duration.
    
    Args:
        path (str): Path to the .xls file.
        timeout (int): Maximum time allowed for conversion (in seconds).
    """
    absolute_path = os.path.abspath(path)
    if absolute_path in conversion_in_progress:
        print(f"Conversion already in progress for {absolute_path}. Skipping.")
        return
    
    conversion_in_progress.add(absolute_path)
    try:
        print(f'Absolute path {absolute_path} found')
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(absolute_path)
        wb.Sheets(1).Name = "YKMD_STD"

        # FileFormat=51 is for .xlsx extension
        new_path = os.path.splitext(absolute_path)[0] + ".xlsx"
        wb.SaveAs(new_path, FileFormat=51)
        wb.Close()
        excel.Application.Quit()
        del wb
        del excel
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        raise e
    finally:
        conversion_in_progress.remove(absolute_path)

def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
   
    orchestrator_connection.log_info("Started process")

    # Opus bruger
    OpusLogin = orchestrator_connection.get_credential("OpusBruger")
    OpusUser = OpusLogin.username
    OpusPassword = OpusLogin.password 

    # Robotpassword
    RobotCredential = orchestrator_connection.get_credential("Robot365User") 
    RobotUsername = RobotCredential.username
    RobotPassword = RobotCredential.password


    specific_content = json.loads(queue_element.data)

    orchestrator_connection.log_info("Assigning variables")

    # Assign variables from SpecificContent
    BookmarkID = specific_content.get("Bookmark")
    OpusBookmark = orchestrator_connection.get_constant("OpusBookMarkUrl").value + str(BookmarkID)
    SharePointURL = specific_content.get("SharePointMappeLink", None)
    FileName = specific_content.get("Filnavn", None)
    Daily = specific_content.get("Dagligt (Ja/Nej)", None)
    MonthEnd = specific_content.get("MånedsSlut (Ja/Nej)", None)
    MonthStart = specific_content.get("MånedsStart (Ja/Nej)", None)
    Yearly = specific_content.get("Årligt (Ja/Nej)", None)
    orchestrator_connection.log_info(f'Processing {FileName}')

    # Ensure that at least one of the options is not None
    if all(option is None for option in [Daily, MonthEnd, MonthStart, Yearly]):
        Run = False

    Run = False
    xlsx_file_path_check = False

    current_date = datetime.now()
    year, month, day = current_date.year, current_date.month, current_date.day
    last_day_of_month = calendar.monthrange(year, month)[1]  

    # Testing if it should run
    if Daily and Daily.lower() == "ja":
        Run = True
    elif MonthEnd and MonthEnd.lower() == "ja" and day == last_day_of_month:
        Run = True
    elif MonthStart and MonthStart.lower() == "ja" and day == 1:
        Run = True
    elif Yearly and Yearly.lower() == "ja" and day == 31 and month == 12:
        Run = True

    if not BookmarkID:
        Run = False

    if Run:
        orchestrator_connection.log_info("Connecting to sharepoint")

        # Parse the base URL
        parsed_url = urlparse(SharePointURL)
        base_url = f"{parsed_url.scheme}://{parsed_url.netloc}"

        # **Automatically Detect if it's a Teams or Sites URL**
        if "/Teams/" in SharePointURL:
            teamsite = SharePointURL.split('Teams/')[1].split('/')[0]
            base_url = f"{base_url}/Teams/{teamsite}"
        elif "/Sites/" in SharePointURL:
            sitename = SharePointURL.split('Sites/')[1].split('/')[0]
            base_url = f"{base_url}/Sites/{sitename}"
        else:
            print("WARNING: Could not determine if this is a Teams or Sites URL. Using default base_url.")


        # Authenticate with SharePoint
        credentials = UserCredential(RobotUsername,RobotPassword)
        ctx = ClientContext(base_url).with_credentials(credentials)
        ctx.load(ctx.web)
        ctx.execute_query()

        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        file_path = os.path.join(downloads_folder, FileName + ".xls")

        if os.path.exists(file_path):
            os.remove(file_path)
            print('File removed')

        # Configure Chrome options
        chrome_options = Options()
        chrome_options.add_argument('--remote-debugging-pipe')
        chrome_options.add_argument("--headless=new")  # More stable headless mode
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": downloads_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
        })

        chrome_service = Service()
        
        max_retries = 3

        for attempt in range(1, max_retries +1):
            try:
                time.sleep(1)
                driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
                break
            except Exception as e:
                orchestrator_connection.log_error(f'Forsøg {attempt} fejlede: {e}')
                if attempt == max_retries:
                    raise
                time.sleep(1)
        
        try:
            orchestrator_connection.log_info("Navigating to Opus login page")
            driver.get(orchestrator_connection.get_constant("OpusAdgangUrl").value)
            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "logonuidfield")))
            
            driver.find_element(By.ID, "logonuidfield").send_keys(OpusUser)
            driver.find_element(By.ID, "logonpassfield").send_keys(OpusPassword)
            driver.find_element(By.ID, "buttonLogon").click()
            
            orchestrator_connection.log_info("Logged in to Opus portal successfully")
            try: 
                driver.get(OpusBookmark)
                WebDriverWait(driver, timeout = 60*10).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe[id^='iframe_Roundtrip']")))
            except Exception as e:
                orchestrator_connection.log_info(f'Exception {e}')
                try:
                    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "changeButton")))
                    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, "changeButton")))
                except TimeoutException:
                    orchestrator_connection.log_info("Change button not found or not clickable")
                else:
                    lower = string.ascii_lowercase
                    upper = string.ascii_uppercase
                    digits = string.digits
                    special = "!@#&%"

                    password_chars = []
                    password_chars += random.choices(lower, k=2)
                    password_chars += random.choices(upper, k=2)
                    password_chars += random.choices(digits, k=4)
                    password_chars += random.choices(special, k=2)

                    random.shuffle(password_chars)
                    password = ''.join(password_chars)

                    driver.find_element(By.ID, "inputUsername").send_keys(OpusPassword)
                    driver.find_element(By.NAME, "j_sap_password").send_keys(password)
                    driver.find_element(By.NAME, "j_sap_again").send_keys(password)
                    driver.find_element(By.ID, "changeButton").click()
                    orchestrator_connection.update_credential('OpusBruger', OpusUser, password)
                    orchestrator_connection.log_info('Password changed and credential updated')
                    time.sleep(2)

            WebDriverWait(driver, timeout = 60*15).until(EC.presence_of_element_located((By.ID, "BUTTON_EXPORT_btn1_acButton")))
            driver.find_element(By.ID, "BUTTON_EXPORT_btn1_acButton").click()
            orchestrator_connection.log_info("Export button clicked")
            initial_file_count = len(os.listdir(downloads_folder))

            orchestrator_connection.log_info("Waiting for file download to complete")

            start_time = time.time()
            while True:
                files = os.listdir(downloads_folder)
                if len(files) > initial_file_count:
                    latest_file = max(
                        [os.path.join(downloads_folder, f) for f in files], key=os.path.getctime
                    )
                    if latest_file.endswith(".xls"):
                        orchestrator_connection.log_info('Found xls file')
                        new_file_path = os.path.join(downloads_folder, f"{FileName}.xls")
                        os.rename(latest_file, new_file_path)
                        orchestrator_connection.log_info(f"File downloaded and renamed to {new_file_path}")
                        xlsx_file_path_check = True
                        break
                    
                if time.time() - start_time > 3600:
                    orchestrator_connection.log_info("Mail sent due to timeout")
                    raise TimeoutError("File download did not complete within 60 minutes.")
                time.sleep(1)

            if xlsx_file_path_check:
                xlsx_file_path = os.path.join(downloads_folder, FileName + ".xlsx")
                try:
                    orchestrator_connection.log_info(f'Converting {new_file_path}')
                    future = convert_xls_to_xlsx( new_file_path)
                    try:
                        future.result()
                        orchestrator_connection.log_info("File converted successfully")
                    except TimeoutError:
                        orchestrator_connection.log_error(f'Conversion of {new_file_path} timed out')
            
                except Exception as e:
                    gc.collect()
                    subprocess.call("taskkill /im excel.exe /f >nul 2>&1", shell=True)
                    time.sleep(2)
                    if os.path.exists(xlsx_file_path):
                        os.remove(xlsx_file_path)
                    orchestrator_connection.log_error(f'An error happened {str(e)}')
                    raise e
            driver.quit()
        except Exception as e:
            orchestrator_connection.log_error(f"An error occurred: {e}")
            print(f"An error occurred: {e}")
            driver.quit()
            raise e

        if xlsx_file_path_check:
            file_name = os.path.basename(xlsx_file_path)
            orchestrator_connection.log_info("Uploading file to sharepoint")

            # Extract path correctly
            query_params = parse_qs(parsed_url.query)
            id_param = query_params.get("id", [None])[0]

            if id_param:
                # If it's a sharing link with an ID, extract the correct path
                decoded_path = unquote(id_param).rstrip('/')
            else:
                # Normal URL or sharing link without ID
                if "/r/" in SharePointURL:
                    decoded_path = SharePointURL.split('/r/', 1)[1].split('?', 1)[0]
                else:
                    decoded_path = parsed_url.path.lstrip('/')
            orchestrator_connection.log_info('Path extracted')

            # **Replace %20 with spaces to match SharePoint folder structure**
            decoded_path = decoded_path.replace("%20", " ")

            # Ensure the correct format
            if not decoded_path.startswith("/"):
                decoded_path = "/" + decoded_path

            folder_relative_url = decoded_path
            target_folder = ctx.web.get_folder_by_server_relative_path(folder_relative_url)
            ctx.load(target_folder)
            ctx.execute_query()

            # Upload file
            file_name = os.path.basename(xlsx_file_path)
            orchestrator_connection.log_info(xlsx_file_path)
            with open(xlsx_file_path, "rb") as local_file:
                target_folder.upload_file(file_name, local_file.read()).execute_query()
                orchestrator_connection.log_info(f"File '{file_name}' uploaded successfully to {SharePointURL}")
                
            if os.path.exists(xlsx_file_path):
                os.remove(xlsx_file_path)
        else:
            print("An error occured - file was not processed correctly")
            orchestrator_connection.log_info("An error occured - file was not processed correctly")
            
        #Deleting potential leftover files from downloads folder
        orchestrator_connection.log_info('Deleting local files')
        if os.path.exists(downloads_folder + '\\' + FileName + ".xls"):
            os.remove(downloads_folder + '\\' + FileName + ".xls")
        if os.path.exists(downloads_folder + '\\' + "YKMD_STD.xls"):
            os.remove(downloads_folder + '\\' + "YKMD_STD.xls")
