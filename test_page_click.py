from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import InvalidSelectorException
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service
from message import message_box
import os
import time


os.system('cls')
script_dir = os.path.abspath(os.path.dirname( __file__ ))

print(f'Starting Edge...')
options = EdgeOptions()
options.add_argument("start-maximized")
options.add_experimental_option('excludeSwitches', ['enable-logging'])
service_path = script_dir + "\\msedgedriver.exe"
edgeBrowser = webdriver.Edge(service=Service(service_path), options=options)

try:
    page_link = f'https://sww-llsak.sakhalinenergy.ru/glasseic/livelink.exe/Open/140084035'
    edgeBrowser.get(page_link)
except Exception as e:
    print(f"{str(e)}, aborting the program")
    edgeBrowser.quit()
    quit()

# message_box('Message', 'Starting page', 0)
time.sleep(3)

for p in range(2, 4):
    element_xpath = f'//td[@title="Page {p}"]'
    try:
        elem = edgeBrowser.find_element(By.XPATH, element_xpath)
    except NoSuchElementException:
        message_box('Error', 'NoSuchElementException', 0)

    elem.click()
    # message_box('Message', f'Page {p}', 0)
    time.sleep(3)

message_box('Message', 'End of script', 0)
edgeBrowser.quit()