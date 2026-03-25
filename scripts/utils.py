from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import os

def make_driver(headless=True):
    options = Options()

    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    if headless:
        options.add_argument("--headless=new")

    # 🔥 CRITICAL FOR RENDER
    options.binary_location = "/usr/bin/chromium"

    return webdriver.Chrome(options=options)