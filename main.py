from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import time

driver = webdriver.Firefox(executable_path=r'C:\Users\yklsh\Downloads\geckodriver-v0.31.0-win64\geckodriver.exe')
browser = webdriver.Firefox()  # Get local session of firefox
driver = webdriver.Firefox()
driver.get("https://yandex.ru/")