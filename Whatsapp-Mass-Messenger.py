from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, StaleElementReferenceException
import openpyxl as xl
import pyautogui as pag
from time import sleep
import tkinter as tk


def message(msg):
    root = tk.Tk()
    frm = tk.Frame(root, pady=10)
    frm.grid()
    tk.Label(frm, text=msg).grid(column=0, row=0)
    tk.Button(frm, text="Quit", command=root.destroy).grid(column=0, row=1)
    root.mainloop()


details = xl.load_workbook("details.xlsx")
sheet = details.active

driver = webdriver.Chrome()
driver.get("https://web.whatsapp.com/")

while True:
    try:
        flag = driver.find_element(
            By.XPATH, '/html/body/div[1]/div/div/div[4]/div/div/div[2]/div[1]').text
        if flag == "WhatsApp Web":
            break
    except (NoSuchElementException, StaleElementReferenceException):
        pass

for i in range(2, sheet.max_row+1):
    number = str(sheet[f"B{i}"].value)
    search = driver.find_element(
        By.XPATH, '/html/body/div[1]/div/div/div[3]/div/div[1]/div/div')
    while True:
        try:
            search.click()
            pag.write(number)
            sleep(2)
            break
        except (NoSuchElementException, ElementNotInteractableException, StaleElementReferenceException):
            pass

    driver.find_elements(By.CSS_SELECTOR, 'div._21S-L')[-1].click()
    sleep(1)
    driver.find_element(
        By.XPATH, '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[1]/div').click()
    pag.write(sheet[f"C{i}"].value)
    sleep(1)
    pag.press('enter')

message("Messages Sent successfully!")
details.close()
driver.quit()
