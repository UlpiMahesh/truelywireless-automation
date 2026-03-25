from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os

def make_driver(headless=True):
    options = Options()

    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-extensions')

    if headless:
        options.add_argument('--headless=new')

    # 🔥 ADD THIS FOR RENDER
    if os.environ.get("RENDER"):
        options.binary_location = "/usr/bin/chromium"
        return webdriver.Chrome(options=options)

    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)