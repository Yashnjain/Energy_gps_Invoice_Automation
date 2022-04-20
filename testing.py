import sharepy
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import os
import time
import time,datetime
import winreg
import bu_alerts
import logging
from datetime import date
from bu_config import get_config
import sys

today_date=date.today()
# log progress --
logfile = os.getcwd() + "\\Logs\\" +'Energy_GPS_AUTOMATION_Logfile'+str(today_date)+'.txt'

logging.basicConfig(filename=logfile, filemode='w',
                    format='%(asctime)s %(message)s')
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s [%(levelname)s] - %(message)s',
    filename=logfile)

logger = logging.getLogger()
logger.setLevel(logging.INFO)

mydate = datetime.datetime.now()
month = mydate.strftime("%b")
year = datetime.date.today().year
path = os.getcwd() + "\\Download"
REG_PATH = r'Software\CutePDF Writer'
site = 'https://biourja.sharepoint.com'
path1 = "/BiourjaPower/_api/web/GetFolderByServerRelativeUrl"
path2= "Shared%20Documents/Power%20Reference/Power_Invoices/Energy_GPS/"
# path2= "Shared Documents/Vendor Research/Enverus(PRT)/ERCOT"
temp_path='https://biourja.sharepoint.com/BiourjaPower/Shared%20Documents/Power%20Reference/Power_Invoices/Energy_GPS'
locations_list=[]
body = ''


def remove_existing_files(files_location):
    logger.info("Inside remove_existing_files function")
    try:
        files = os.listdir(files_location)
        if len(files) > 0:
            for file in files:
                os.remove(files_location + "\\" + file)
            logger.info("Existing files removed successfully")
        else:
            print("No existing files available to reomve")
        print("Pause")
    except Exception as e:
        logger.info(e)
        raise e


def set_reg(name, value):
    try:
        winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_PATH)
        registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, 
                                       winreg.KEY_WRITE)
        winreg.SetValueEx(registry_key, name, 0, winreg.REG_SZ, value)
        winreg.CloseKey(registry_key)
        return True
    except WindowsError:
        return False
fp=webdriver.FirefoxProfile()
mime_types=['application/pdf'
            ,'text/plain',
            'application/vnd.ms-excel',
            'test/csv',
            'application/csv',
            'text/comma-separated-values','application/download','application/octet-stream'
            ,'binary/octet-stream'
            ,'application/binary'
            ,'application/x-unknown']
fp.set_preference('browser.download.folderList',2)
fp.set_preference('browser.download.manager.showWhenStarting',False)
fp.set_preference('browser.download.dir',path)
fp.set_preference('browser.helperApps.neverAsk.saveToDisk',','.join(mime_types))
# fp.set_preference('pdfjs.disabled',True)
# fp.set_preference('print.always_print_silent', True)
fp.set_preference('print_printer', 'CutePDF Writer')
fp.set_preference("print.always_print_silent", True)
fp.set_preference("print.show_print_progress", True)
fp.set_preference('print.save_as_pdf.links.enabled', True)
fp.set_preference("pdjs.disabled", True)
fp.set_preference('print.printer_CutePDF.print_to_file', True)
fp.set_preference('print.printer_CutePDF.print_to_file.print_to_filename',
                   "testprint.pdf")

driver=webdriver.Firefox(executable_path=GeckoDriverManager().install(),firefox_profile=fp)
# driver = webdriver.Chrome(executable_path=executable_path,options=chromeOptions)

def login_and_download():  
    '''This function downloads file from the website'''
    try:
            logging.info('Accesing website')
            driver.get("https://www.energygps.com/Account/LogOn")
            time.sleep(10)
            logging.info('Accept Cookies')
            WebDriverWait(driver, 5, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div/a"))).click()
            logging.info('providing id and passwords')
            time.sleep(3)
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//input[@id='UserName']"))).send_keys(username)
            time.sleep(1)
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//input[@id='Password']"))).send_keys(password)        
            time.sleep(1)
            logging.info('click on Log In Button')
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Login']"))).click()        
            time.sleep(5)
            logging.info('click on Order History')
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.LINK_TEXT, "Order History"))).click()        
            time.sleep(5)
            if month == WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html[1]/body[1]/div[1]/div[1]/section[1]/table[1]"))).text.split("Manage Subscription")[0].split()[-5]:
                time.sleep(2)
                receipt_no=WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html[1]/body[1]/div[1]/div[1]/section[1]/table[1]"))).text.split("Manage Subscription")[0].split()[-10]
                time.sleep(5)
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.LINK_TEXT, f"{receipt_no}"))).click()        
                time.sleep(5)
                filename=f"EnergyGPSReceipt - {month}{year}"+'.pdf' 
                time.sleep(10)
                set_reg('BypassSaveAs', '1')
                time.sleep(4)
                set_reg('OutputFile', f'S:\IT Dev\Testing_Environment\ENERGY_GPS_INVOICE_AUTOMATION\Download\{filename}')
                logging.info('downloading file')
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.LINK_TEXT,"Print this page"))).click()
            else:
                try:
                    driver.quit()
                except Exception as e: 
                    logging.info('driver quit failed')
                    print("driver quit failed") 
                bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name},File for the {month} not found',mail_body = f'{job_name} completed successfully, Attached PDF and Logs',attachment_location=logfile)
                sys.exit(0)
            try:
                driver.close()
            except Exception as e: 
                logging.info('driver not closed')
                print("driver not closed") 
                try:
                    driver.quit()
                except Exception as e: 
                    logging.info('driver quit failed')
                    print("driver quit failed") 
    except Exception as e:
            print(f"{e}")
            logging.exception(str(e))
def connect_to_sharepoint():
    logging.info('Connecting to sharepoint')
    try:
        site='https://biourja.sharepoint.com'
        username = os.getenv("user") if os.getenv("user") else sp_username
        password = os.getenv("password") if os.getenv("password") else sp_password
        # Connecting to Sharepoint and downloading the file with sync params
        s = sharepy.connect(site, username, password)
        return s
    except Exception as e:
        raise e

def shp_file_upload(s):
    logging.info('Uploading files to sharepoint')
    try:
        global body
        body = ''
        filesToUpload = os.listdir(os.getcwd() + "\\Download")
        for fileToUpload in filesToUpload:
            z=path+'\\'+fileToUpload
            locations_list.append(z)     
            headers = {"accept": "application/json;odata=verbose",
            "content-type": "application/pdf"}

            with open(os.path.join(os.getcwd() + "\\Download", f'{fileToUpload}'), 'rb') as read_file:
                    content = read_file.read()
            # fileToUpload=fileToUpload.replace("'","_")     
            p = s.post(f"{site}{path1}('{path2}')/Files/add(url='{fileToUpload}',overwrite=true)", data=content, headers=headers)
            nl = '<br>'
            body += (f'{fileToUpload} successfully uploaded, {nl} Attached link for the same:-{nl}{temp_path}{nl}')

        print(f'{fileToUpload} uploaded successfully')
    
        print(f'{job_name} executed succesfully')
        return p   
        
    except Exception as e:
        raise e

def main():
    try:
        remove_existing_files(files_location)
        login_and_download()
        s=connect_to_sharepoint()
        shp_file_upload(s)
        locations_list.append(logfile)
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{body}{job_name} completed successfully, Attached PDF and Logs',attachment_location = logfile)
    except Exception as e:
        logging.exception(str(e))
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed, Attached logs',attachment_location = logfile)
                
if __name__ == "__main__":
    logging.info("Execution Started")
    time_start=time.time()
    directories_created=["Download","Logs"]
    for directory in directories_created:
        path3 = os.path.join(os.getcwd(),directory)  
        try:
            os.makedirs(path3, exist_ok = True)
            print("Directory '%s' created successfully" % directory)
        except OSError as error:
            print("Directory '%s' can not be created" % directory)

    files_location=os.getcwd() + "\\Download"
    Project_name="ENERGY_GPS_INVOICE_AUTOMATION"
    Table_name="ENERGY_GPS_INVOICE_AUTOMATION"
    credential_dict = get_config('ENERGY_GPS_INVOICE_AUTOMATION','ENERGY_GPS_INVOICE_AUTOMATION')
    username = credential_dict['USERNAME'].split(';')[0]
    password = credential_dict['PASSWORD'].split(';')[0]
    sp_username = credential_dict['USERNAME'].split(';')[1]
    sp_password =  credential_dict['PASSWORD'].split(';')[1]
    receiver_email = credential_dict['EMAIL_LIST']
    # receiver_email ='yashn.jain@biourja.com'
    job_name='ENERGY_GPS_INVOICE_AUTOMATION'
   
    main()
    time_end=time.time()
    logging.info(f'It takes {time_start-time_end} seconds to run')      
