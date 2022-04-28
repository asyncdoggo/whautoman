import os
import subprocess
import time
from typing import Union
import pandas
import requests
import selenium.common.exceptions
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# xpath definitions
send_btn = "/html/body/div[1]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/div[2]/div[2]/div/div"
imgvid_btn = "/html/body/div[1]/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div[" \
             "1]/div/ul/li[1]/button "
pinbutton = "/html/body/div[1]/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div/span[2]/div/div[1]/div[2]"
qr_code = "/html/body/div[1]/div[1]/div/div[2]/div[1]/div/div[2]/div"
ok_button = "/html/body/div[1]/div[1]/span[2]/div[1]/span/div[1]/div/div/div/div/div[2]"
text_box = "/html/body/div[1]/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]"
starting_chat = "_1bpDE"
document_button = "/html/body/div[1]/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div/span[2]/div/div[1]/div[" \
                  "2]/div/span/div[1]/div/ul/li[4]/button"
cmsg = "Please wait...."


class Automate:

    def __init__(self, numbers: list):
        self.numbers = numbers

        # install driver
        os.environ['WDM_LOCAL'] = '1'
        path = ChromeDriverManager(cache_valid_range=30).install()

        path = path.replace("chromedriver.exe", "")
        # add driver to path
        os.environ["PATH"] += os.pathsep + path

        # start driver
        self.driver = webdriver.Chrome()
        self.wait = WebDriverWait(self.driver, 600)
        self.driver.get("https://web.whatsapp.com/")

        # initial qr login
        self.wait.until_not(EC.presence_of_element_located((By.XPATH, qr_code)))

    def send_files(self, documents: list = None, images: list = None):
        # Access the pin button
        pin = self.wait.until(EC.presence_of_element_located((By.XPATH, pinbutton)))
        pin.click()

        # Access the images and video button
        if images:
            image_icon = self.wait.until(
                EC.presence_of_element_located((By.XPATH, imgvid_btn)))
            image_icon.click()

        elif documents:
            docico = self.wait.until(
                EC.presence_of_element_located((By.XPATH, document_button)))
            docico.click()

        # Convert list of path to a single string
        p = os.path.abspath("au.exe")
        args = [p]
        if images:
            for i in images:
                args.append(i)

        elif documents:
            for i in documents:
                args.append(i)

        time.sleep(1)  # SHORT SLEEP TO LET OPEN PROMPT START
        # Use autoit script

        subprocess.Popen(" ".join(args), shell=True)
        time.sleep(1)

        # Access the send button
        send_button = self.wait.until(
            EC.presence_of_element_located((By.XPATH, send_btn)))
        send_button.click()
        time.sleep(1)

    def send(self, data_type: str, data: Union[str, list]) -> None:
        global cmsg
        try:
            for i in self.numbers:
                time.sleep(1)
                phone = str(i)
                self.driver.get(f"https://web.whatsapp.com/send?phone=91{phone}")

                # check for wrong number (not used EC here because uncertainty)
                flag = 0

                self.wait.until_not(EC.presence_of_element_located((By.CLASS_NAME, starting_chat)))
                while True:
                    try:
                        elem = self.driver.find_element(By.XPATH, ok_button)
                        if elem.text == "OK":
                            flag = -1
                            cmsg = f"Number not found {phone}"
                            break
                    except (NoSuchElementException, StaleElementReferenceException):
                        pass
                    try:
                        self.driver.find_element(By.XPATH, text_box)
                        break
                    except NoSuchElementException:
                        pass
                if flag == -1:
                    continue

                text = self.wait.until(EC.presence_of_element_located((By.XPATH, text_box)))

                if data_type == 'TEXT':
                    text.send_keys(data)
                    text.send_keys(Keys.ENTER)
                    cmsg = f"Message sent to {phone}"

                elif data_type == 'IMAGE':
                    self.send_files(images=data)  # image/video files
                    cmsg = f"Image sent to {phone}"

                elif data_type == 'DOCUMENT':
                    self.send_files(documents=data)  # documents
                    cmsg = f"Document sent to {phone}"

                elif data_type == "TEXT+IMAGE":
                    text.send_keys(data[0])
                    text.send_keys(Keys.ENTER)
                    self.send_files(images=data[1])
                    cmsg = f"Data sent to {phone}"

                elif data_type == "TEXT+DOCS":
                    text.send_keys(data[0])
                    text.send_keys(Keys.ENTER)
                    self.send_files(documents=data[1])
                    cmsg = f"Data sent to {phone}"

            cmsg = 'END'
        except selenium.common.exceptions.WebDriverException as e:
            print(repr(e))

    def close(self):
        self.driver.quit()
        exit(0)


if __name__ == "__main__":
    excdata = pandas.read_excel("data.xlsx", sheet_name="Sheet1")
    number = excdata["Numbers"].to_list()
    try:
        test = Automate(number)
        test.send("TEXT", "hello")
    except requests.exceptions.ConnectionError:
        print("Connection error")
