import os
import subprocess
import time
import pandas
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# xpath definitions
send_btn = "/html/body/div[1]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/div[2]/div[2]/div/div"
imgvid_btn = "/html/body/div[1]/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div[1]/div/ul/li[1]/button"
pinbutton = "/html/body/div[1]/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div/span[2]/div/div[1]/div[2]"
qr_code = "/html/body/div[1]/div[1]/div/div[2]/div[1]/div/div[2]/div"
ok_button = "/html/body/div[1]/div[1]/span[2]/div[1]/span/div[1]/div/div/div/div/div[2]"
text_box = "/html/body/div[1]/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]"
loading_circle = "/html/body/div[1]/div[1]/div[1]/div[4]/div[1]/div[3]/div/div[2]/div[2]/div/svg"
starting_chat = "/html/body/div[1]/div[1]/span[2]/div[1]/span/div[1]/div/div/div/div/div[1]"

cmsg = ("Sending", "")


class Automate:

    def __init__(self, numbers):
        self.numbers = numbers

        # install driver
        path = ChromeDriverManager().install()
        path = path.replace("chromedriver.exe", "")
        # add driver to path
        os.environ["PATH"] += os.pathsep + path

        # start driver
        self.driver = webdriver.Chrome()
        self.wait = WebDriverWait(self.driver, 600)
        self.driver.get("https://web.whatsapp.com/")

        # initial qr login
        self.wait.until_not(EC.presence_of_element_located((By.XPATH, qr_code)))

    def send_img_vid(self, file_path):
        # local functions
        def get_file_paths(path):
            file_name = []

            for files in os.listdir(path):
                root, ext = os.path.splitext(files)
                if ext == ".png" or ext == ".jpg" or ext == ".mp3" or ext == ".gif" or ext == "jpeg":
                    file_name.append(path + str("\\") + files)

            return file_name

        def parse_paths(paths):
            ans = ""
            for i in paths:
                ans = ans + str(i)
                ans = ans + str(" ")
            return ans

        # Access the pin button
        pin = self.wait.until(EC.presence_of_element_located((By.XPATH, pinbutton)))
        pin.click()

        # Access the images and video button
        image_icon = self.wait.until(
            EC.presence_of_element_located((By.XPATH, imgvid_btn)))
        image_icon.click()

        # Get all the images and videos path from the given directory
        img_list = get_file_paths(file_path)

        # Convert list of path to a single string (because it is required by subprocess)
        parse_path = parse_paths(img_list)

        # Use custom autoit script to insert doc path
        time.sleep(1)  # SHORT SLEEP TO LET OPEN PROMPT START
        subprocess.run(['AutoIt_Script.exe', parse_path], shell=True)  # runs .exe autoit file
        time.sleep(1)

        # Access the send button
        send_button = self.wait.until(
            EC.presence_of_element_located((By.XPATH, send_btn)))
        send_button.click()
        time.sleep(1)

    def send(self, data_type, data):
        global cmsg

        for i in self.numbers:
            time.sleep(2)
            phone = str(i)
            self.driver.get(f"https://web.whatsapp.com/send?phone=91{phone}")

            # check for wrong number (not used EC here because uncertainty)
            flag = 0
            while True:
                try:
                    self.driver.find_element(By.XPATH, starting_chat)
                    try:
                        elem = self.driver.find_element(By.XPATH, ok_button)
                        if elem.text == "OK":
                            flag = -1
                            cmsg = ("Number not found ", phone)
                            break
                    except (NoSuchElementException, StaleElementReferenceException):
                        pass

                except NoSuchElementException:
                    try:
                        self.driver.find_element(By.XPATH, text_box)
                        break
                    except NoSuchElementException:
                        pass
            if flag == -1:
                continue

            text = self.wait.until(EC.visibility_of_element_located((By.XPATH, text_box)))

            if data_type == 'TEXT':
                text.send_keys(data)
                text.send_keys(Keys.ENTER)
                cmsg = ("Message sent to ", phone)

            if data_type == 'IMAGE':
                self.send_img_vid(data)
                cmsg = ("Message sent to ", phone)

        cmsg = ('END', "")


if __name__ == "__main__":
    excdata = pandas.read_excel("data.xlsx", sheet_name="Sheet1")
    number = excdata["Numbers"].to_list()
    test = Automate(number)
    test.send("TEXTDATA", "hello")
