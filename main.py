import os
import time
import pandas
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

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

        # setup
        self.driver = webdriver.Chrome()

        self.driver.get("https://web.whatsapp.com/")

        # initial qr login
        while True:
            try:
                self.driver.find_element(By.CLASS_NAME, "_2UwZ_")
            except NoSuchElementException:
                break

    def send(self, datatype, data):
        global cmsg

        for i in self.numbers:
            time.sleep(2)
            phone = str(i)
            self.driver.get(f"https://web.whatsapp.com/send?phone=91{phone}")

            # check for wrong number and continue loop if found
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

            while True:
                try:
                    self.driver.find_element(By.XPATH, loading_circle)
                except NoSuchElementException:

                    # get text box and send (text box is reassigned to avoid StaleElementReferenceException)
                    try:
                        text = self.driver.find_element(By.XPATH, text_box)
                        text.send_keys(data)
                        text.send_keys(Keys.ENTER)
                        cmsg = ("Message sent to ", phone)
                        break
                    except NoSuchElementException:
                        continue

        cmsg = ('END',"")

if __name__ == "__main__":
    data = pandas.read_excel("data.xlsx", sheet_name="Sheet1")
    numbers = data["Numbers"].to_list()
    test = Automate(numbers)
    test.send("TEXTDATA", "hello")
