from selenium import webdriver
import time
import os
import requests

# Define your download folder
download_folder = os.path.abspath("xlsx")  # Change this to your preferred directory

# Create folder if it doesn’t exist
os.makedirs(download_folder, exist_ok=True)

#browser = webdriver.Chrome()
options = webdriver.ChromeOptions() 
#options.add_experimental_option("useAutomationExtension", False);
options.add_argument("user-data-dir=C:\\TableAu") #Path to your chrome profile
#options.add_argument("--disable-blink-features=AutomationControlled");
#options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
options.add_argument("window-size=1920,1080")
options.add_argument("--headless")
options.add_argument("--log-level=3")  # Suppresses warnings and errors
options.add_experimental_option("excludeSwitches", ["enable-logging"])  # Suppresses DevTools messages
options.add_experimental_option("excludeSwitches", ["enable-automation"])  # Removes "Chrome is controlled" message
options.add_experimental_option("prefs", {
    "download.default_directory": download_folder,  # Set default download location
    "download.prompt_for_download": False,  # Auto-download without prompt
    "directory_upgrade": True
})
browser = webdriver.Chrome(options=options)

def OpenUrl(url):
    browser.get(url)

def ClickButton(xpath):
    elem = browser.find_element("xpath", xpath)
    elem.click()

def DownloadSheet(url):
    url = url + "?%3Adisplay_static_image=n&%3Aembed=true&%3Aembed=y&%3Alanguage=en-US&publish=yes&%3AshowVizHome=n&%3AapiID=host0#navType=0&navSrc=Parse"
    OpenUrl(url)
    time.sleep(5)
    download1_button_xpath = '//*[@id="download"]'
    crosstab_button_xpath = '//*[@id="viz-viewer-toolbar-download-menu"]/div[3]/div/div/span[2]'
    dowload2_button_xpath = '//*[@id="export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id"]/div/div[3]/button'
    ClickButton(download1_button_xpath)
    time.sleep(3)
    ClickButton(crosstab_button_xpath)
    time.sleep(2)
    ClickButton(dowload2_button_xpath)
    time.sleep(5)

if __name__ == "__main__":
    f = open("urls.txt", "r")
    list_urls = f.readlines()
    for url in list_urls:
        print("Downloading file from: " + url.strip())
        DownloadSheet(url.strip())
    
