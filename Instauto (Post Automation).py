from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, StaleElementReferenceException
from os import listdir, system, path
from time import sleep
from datetime import datetime
from pathlib import Path
import tkinter as tk
import pyautogui


class INSTAGRAM():
    def __init__(self):
        self.browser = webdriver.Chrome()
        self.browser.get("https://www.instagram.com/")
        print("WEB PAGE of Instagram Loaded")

    def signin(self, username, password):
        while True:  # checks if login form is opened
            try:
                self.browser.find_element(By.ID, "loginForm")
                break
            except NoSuchElementException:
                pass
        self.browser.find_element(By.NAME, "username").send_keys(username)  # enter username
        self.browser.find_element(By.NAME, "password").send_keys(password + '\n')  # enter password

    def createpost(self, image, description):
        print("LOGIN Successful")

        while True:
            try:
                if self.browser.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div[1]/div/div/div/div[1]/div[1]/div[2]/section/main/div/div/div/section/div/div[2]').text in ['Save your login information?', 'Save Your Login Info?']:
                    self.browser.find_element(
                        By.XPATH, '/html/body/div[2]/div/div/div/div[1]/div/div/div/div[1]/div[1]/div[2]/section/main/div/div/div/div/button').click()
            except NoSuchElementException:
                pass
            try:
                if self.browser.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div[2]/h2').text in ['Turn on notifications', 'Turn on Notifications']:
                    self.browser.find_element(
                        By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div[3]/button[2]').click()
                    break
            except NoSuchElementException:
                pass
        while True:  # create post click
            try:
                self.browser.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/div/div[1]/div/div/div/div[1]/div[1]/div[1]/div/div/div/div/div[2]/div[6]/div/div/a/div').click()
                break
            except NoSuchElementException:
                pass
            try:
                self.browser.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/div/div[1]/div/div/div/div[1]/div[1]/div[1]/div/div/div/div/div[2]/div[7]/div/div/a/div/div[1]').click()
                break
            except NoSuchElementException:
                pass
        while True:  # uploads the file
            try:
                self.browser.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[2]/div[1]/div/div/div[2]/div/button').click()
                sleep(1)
                # path of File
                pyautogui.write(image)
                pyautogui.press('enter')
                break
            except NoSuchElementException:
                pass
        while True:
            try:
                # click on crop
                self.browser.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[2]/div[1]/div/div/div/div[1]/div/div[2]/div/button').click()
                # click on original
                self.browser.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[2]/div[1]/div/div/div/div[1]/div/div[1]/div/button[1]').click()
                # click on Next
                self.browser.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div[3]/div/button').click()
                sleep(1)
                self.browser.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div[3]/div/button').click()
                break
            except (NoSuchElementException, ElementNotInteractableException):
                pass

        while True:  # types the description
            try:
                sleep(1)
                xyz = self.browser.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/div/div[2]/div[1]')
                xyz.click()
                pyautogui.write(description)
                break
            except (NoSuchElementException, ElementNotInteractableException, StaleElementReferenceException):
                pass
        # click on share
        try:
            self.browser.find_element(
                By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div[3]/div/button').click()
        except StaleElementReferenceException:
            pass

        while True:
            try:
                post_shared = self.browser.find_element(
                    By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/div[2]/div/div/div/div[2]/div[1]/div/div[2]/div/span').text
                if post_shared == "Your post has been shared.":
                    print("POST SHARED SUCCESSFULLY")
                    break
            except NoSuchElementException:
                pass

    def quit(self):
        self.browser.quit()


def main():
    try:  # try to read username and password from login.txt
        with open(Path('login.txt'), 'r') as f:
            l = f.readlines()
            username, password = l[0], l[1]
    except (FileNotFoundError, IndexError):  # asks user to enter username and password if file not found
        username = input("Enter the username: ")
        password = input("Enter the password: ")
        if input("Save Username & Password into 'login.txt' (Y): ") in ('y', 'Y'):
            with open(Path('login.txt'), 'w') as f:  # create a login.txt
                f.write(username+'\n' + password)  # enter the username & password into file
            print("Details Saved into file!!")

    timing = input("Enter the timing (HH:MM): ")
    if len(timing) == 5:
        try:
            now = datetime.now()
            current_time = now.strftime("%H:%M")
            seconds = (((int(timing[:2]) * 60 + int(timing[3:])) * 60) - ((int(current_time[:2]) * 60 + int(current_time[3:])) * 60))
            if seconds > 30:
                print(f"Sleeping for {seconds} seconds.\nWill Wake up at {timing} - 30 seconds.")
                sleep(seconds)
            print("\nWOKE")
            while True:
                now = datetime.now()
                current_time = now.strftime("%H:%M")
                if current_time == timing:
                    break
        except:
            pass


    driver = INSTAGRAM()  # create a web object
    driver.signin(username, password)  # signin the user
    today_date = str(datetime.now())[:10]

    try:
        with open(Path(today_date + '.txt'), 'r') as d:  # reads the description
            description = d.read()
        post = path.abspath(today_date + '.jpg')
        if not path.exists(post):
            raise FileNotFoundError
    except FileNotFoundError:
        message("Error: Post is not available in specified folder!!")
        exit()

    # create and share the post
    driver.createpost(post, description)
    driver.quit()  # quits the web driver

    message("Instagram Post Shared Successfully!!")


def message(msg):
    root = tk.Tk()
    frm = tk.Frame(root, pady=10)
    frm.grid()
    tk.Label(frm, text=msg).grid(
        column=0, row=0)
    tk.Button(frm, text="Quit", command=root.destroy).grid(column=0, row=1)
    root.mainloop()


main()
