from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, ElementClickInterceptedException
from os import listdir, system
from pathlib import Path
from time import sleep
from datetime import datetime
import openpyxl as xl


column_filter_flag = True
default_removals = "inc:inc.:,inc:llc:llc.:,llc:llp:llp.:,llp:co:co.:plan:and:&:trust:ltd:,ltd:ltd.:the:401:401k:401(k):k:(k):assetmark:db"


class BROWSER():
    def __init__(self, headless):  # Opening Automated Webpage in Browser
        self.options = webdriver.EdgeOptions()
        if headless:
            self.options.add_argument("--headless")
        self.options.add_experimental_option(
            "excludeSwitches", ["enable-logging"])
        self.browser = webdriver.Edge(options=self.options)
        self.browser.get(
            "https://cap.brightscope.com/search/beacon/#/plan")

    def sign_in(self, email, password):  # Sign In into Beacon
        self.browser.find_element(By.NAME, "email").click()
        action = ActionChains(self.browser)
        action.key_down(Keys.CONTROL).send_keys(
            'A').key_up(Keys.CONTROL).perform()
        self.browser.find_element(By.NAME, "email").send_keys(email)
        self.browser.find_element(By.NAME, "password").send_keys(password+'\n')
        # self.browser.find_element(By.TAG_NAME, "button").click()
        while True:
            try:
                self.browser.find_element(
                    By.CSS_SELECTOR, "button.main-search-button").click()
                return None
            except (NoSuchElementException, ElementNotInteractableException):
                try:
                    self.browser.find_element(By.CSS_SELECTOR, "ul.errors")
                    email, password = email_n_password()
                    self.sign_in(email, password)
                    return None
                except NoSuchElementException:
                    pass

    def search_filters(self):  # Activate Plan Name in Search Filter Navigation
        flag = False
        plan_name_checkbox = self.browser.find_element(
            By.XPATH, '//*[@id="filter-nav"]/div/div[2]/v-accordion/v-pane[2]/v-pane-content/div/div[14]/bale-filter/div/div[1]/md-checkbox/div[2]')
        while not flag:
            flag = self.browser.find_element(
                By.ID, "filter-nav").is_displayed()
            if flag:
                self.browser.find_elements(
                    By.TAG_NAME, "v-pane")[1].click()
                flag = False
                while not flag:
                    flag = plan_name_checkbox.is_displayed()
        plan_name_checkbox.click()

    def column_filters(self):  # Activate EIN & Company Name in Column Filter Navigation
        self.browser.find_element(By.CSS_SELECTOR, "div.col-button").click()
        flag = False
        while not flag:
            flag = self.browser.find_element(
                By.XPATH, '//*[@id="body-content"]/ui-view/beacon-search/bale-search/bale-column-menu/md-sidenav/div/div[2]/v-accordion/v-pane[2]/v-pane-header/div/column-category/div/div[1]').is_displayed()
            if flag:
                self.browser.find_element(
                    By.XPATH, '//*[@id="body-content"]/ui-view/beacon-search/bale-search/bale-column-menu/md-sidenav/div/div[2]/v-accordion/v-pane[2]/v-pane-header/div/column-category/div/div[1]').click()
                flag = False
                while not flag:
                    flag = self.browser.find_element(
                        By.XPATH, '//*[@id="body-content"]/ui-view/beacon-search/bale-search/bale-column-menu/md-sidenav/div/div[2]/v-accordion/v-pane[2]/v-pane-content/div/div[11]/div/md-checkbox').is_displayed()
        self.browser.find_element(
            By.XPATH, '//*[@id="body-content"]/ui-view/beacon-search/bale-search/bale-column-menu/md-sidenav/div/div[2]/v-accordion/v-pane[2]/v-pane-content/div/div[11]/div/md-checkbox/div[2]').click()
        while True:
            try:
                self.browser.find_element(
                    By.XPATH, '//*[@id="body-content"]/ui-view/beacon-search/bale-search/bale-column-menu/md-sidenav/div/div[2]/v-accordion/v-pane[2]/v-pane-content/div/div[13]/div/md-checkbox/div[2]').click()
                break
            except:
                pass
        try:  # Removes "Send Us Email" button from the bottom
            send_us_email = self.browser.find_element(
                By.CSS_SELECTOR, "div.olark-text-button")
            if self.browser.find_element(By.CSS_SELECTOR, "div.olark-text-button").is_displayed():
                self.browser.execute_script("""
                    var element = arguments[0];
                    element.parentNode.removeChild(element);
                    """, send_us_email)
        except:
            pass

        self.browser.find_element(
            By.XPATH, '//*[@id="body-content"]/ui-view/beacon-search/bale-search/bale-column-menu/md-sidenav/div/div[3]/button/div[1]').click()
        flag = True
        while flag:
            flag = self.browser.find_element(
                By.CSS_SELECTOR, "md-sidenav.md-sidenav-right").is_displayed()

    def plan_name_search(self, pname):  # Type Plan Name in Field and Run Search
        flag = self.browser.find_element(
            By.ID, "filter-nav").is_displayed()
        if not flag:
            try:
                self.browser.find_element(
                    By.CSS_SELECTOR, "div.company-information").click()
            except (NoSuchElementException, ElementNotInteractableException, ElementClickInterceptedException):
                self.browser.find_element(
                    By.CSS_SELECTOR, "button.main-search-button").click()
        while True:
            try:
                self.browser.find_element(By.ID, "fl-input-125").click()
                break
            except (NoSuchElementException, ElementNotInteractableException, ElementClickInterceptedException):
                pass
        plan_name_input_field = self.browser.find_element(
            By.ID, "fl-input-125")
        action = ActionChains(self.browser)
        action.key_down(Keys.CONTROL).send_keys(
            'A').key_up(Keys.CONTROL).perform()
        plan_name_input_field.send_keys(pname)
        flag = False
        plan_name_field_list = self.browser.find_element(
            By.XPATH, '//*[@id="ul-125"]')
        sleep(1)
        time = datetime.timestamp(datetime.now())
        while True:
            try:
                sleep(1)
                flag = plan_name_field_list.is_displayed()
                if flag:
                    self.browser.find_element(
                        By.XPATH, '//*[@id="ul-125"]/li[1]/md-autocomplete-parent-scope').click()
                    break
                else:
                    plan_name_input_field.click()
            except (NoSuchElementException, ElementNotInteractableException):
                if datetime.timestamp(datetime.now()) - time > 60:
                    break
                else:
                    pass

        self.browser.find_element(
            By.CSS_SELECTOR, "button.submit-button").click()

    def results(self):  # Search Results
        flag = True
        while flag:
            flag = self.browser.find_element(
                By.ID, "filter-nav").is_displayed()
        flag = False
        while not flag:
            try:
                flag = self.browser.find_element(
                    By.CSS_SELECTOR, "span.small-header").is_displayed()
                global column_filter_flag
                if column_filter_flag:
                    self.column_filters()
                    column_filter_flag = False
            except NoSuchElementException:
                try:
                    self.browser.find_element(
                        By.CSS_SELECTOR, "button.main-search-button").click()
                    return 0, self.browser.find_elements(By.CSS_SELECTOR, "div.search-apology")
                except (NoSuchElementException, ElementClickInterceptedException):
                    pass
        search_results = self.browser.find_element(
            By.CSS_SELECTOR, "span.small-header").text
        results_plan = self.browser.find_elements(
            By.CSS_SELECTOR, "div.result-cell")
        search_results = search_results[16:-1]
        if len(search_results) > 2:
            search_results = 30
        return int(search_results), results_plan

    def quit(self):  # Close the Automated Webpage
        self.browser.quit()


def main():  # Main Function
    # Email ID and Password for LogIn into Beacon
    global removals
    PARSED_ORG, DATE_COL = "Y", "AP"
    PROGRAM_CHECKED, PROGRAM_SEARCHED, PROGRAM_RESULT = "BA", "BB", "BC"
    headless_mode = True
    try:
        with open(Path("login.txt"), 'r') as file:
            lines = file.readlines()
            removals = lines[0].split(':')
            PARSED_ORG, DATE_COL = lines[10][18:-1], lines[13][14:-1]
            PROGRAM_CHECKED = lines[16][12:-1]
            PROGRAM_SEARCHED = lines[17][12:-1]
            PROGRAM_RESULT = lines[18][12:-1]
        del lines
    except FileNotFoundError:
        removals = default_removals.split(':')
        default_file_format()

    file_name = excel_file_name_input()
    print(f'File - {file_name}\nOpening File...')
    data_excel = xl.load_workbook(Path(file_name))
    data_sheet = data_excel.active
    print("File Opened Successfully!!")

    email, password = email_n_password()
    row_range = input(
        "\n*Range of the records to be checked (Example: '100-200')\n or press 'Enter' key for whole file:\n")
    if row_range == "":
        r1, r2 = 2, data_sheet.max_row
    else:
        try:
            r1, r2 = [int(i) for i in row_range.split('-')]
            if r1 > r2:
                r1, r2 = r2, r1
            elif r1 == r2:
                input("Range can not be Equal!! Closing...")
                exit()
            else:
                pass
        except ValueError:
            input("Wrong input range typed!! Closing...")
            exit()

    if input("\n*By Default, Program will run in Headless Mode.\n To run in Non-Headless mode, type 'F' and press 'Enter' key: ") in ('F', 'f'):
        headless_mode = False

    # Creating the Web Automation Object
    count = save_counter = 0
    start = str(datetime.now())[11:19]
    date = str(datetime.now())[:10]
    previous_plan = previous_cleaned_plan = 'a'

    try:
        driver = BROWSER(headless_mode)
        driver.sign_in(email, password)
        driver.search_filters()
        del email, password, row_range, headless_mode

        data_sheet[f"{PROGRAM_CHECKED}1"].value = "P_Check"
        data_sheet[f"{PROGRAM_SEARCHED}1"].value = "P_Search"
        data_sheet[f"{PROGRAM_RESULT}1"].value = "P_Result"

        for row in range(r1, r2 + 1):
            # checking if already checked
            if (data_sheet[f"{PROGRAM_CHECKED}{row}"].value != None or data_sheet[f"{DATE_COL}{row}"].value != None):
                continue
            system('cls')
            print(
                f"***Beacon: Automation***\n\nStart Time : {start}\nCount      : {count}\nCurrent Row: {row}\n\nSave Counter: {save_counter}")
            data_sheet[f"{PROGRAM_CHECKED}{row}"].value = date
            plan = str(data_sheet[f"{PARSED_ORG}{row}"].value)
            cleaned_plan = clean_plan(plan)
            if (plan == previous_plan) or (cleaned_plan == previous_cleaned_plan):
                data_sheet[f"{PROGRAM_RESULT}{row}"].value = data_sheet[f"{PROGRAM_RESULT}{row - 1}"].value
                continue
            elif len(cleaned_plan) < 5:
                data_sheet[f"{PROGRAM_RESULT}{row}"].value = 'Not Searched'
                continue

            save_counter += 1
            if save_counter % 100 == 0:
                print("\n\n*** Saving File...")
                data_excel.save(f"./{file_name[:-5]} Auto.xlsx")
                print("\n*** File Saved !!! ***")

            previous_plan = plan
            data_sheet[f"{PROGRAM_SEARCHED}{row}"].value = previous_cleaned_plan = cleaned_plan
            driver.plan_name_search(cleaned_plan)

            search_results, results_plan = driver.results()
            if search_results == 0:
                data_sheet[f"{PROGRAM_RESULT}{row}"].value = 'Not Found'
                continue
            elif search_results > 29:
                data_sheet[f"{PROGRAM_RESULT}{row}"].value = 'Multiple Webpages'
                continue

            Results = {'Plans': [], 'Dates': [],
                       'Company': [], 'EIN': [], 'FA': []}
            for i in range(search_results):
                Results['Plans'].append(results_plan[5+i].text)
                Results['Dates'].append(results_plan[5+i+search_results].text)
                Results['Company'].append(
                    results_plan[5+i+search_results*2].text)
                Results['EIN'].append(results_plan[5+i+search_results*3].text)
                Results['FA'].append(results_plan[5+i+search_results*4].text)

            data_sheet[f"{PROGRAM_RESULT}{row}"].value = decision(
                search_results, Results, cleaned_plan)
            count += 1
    except:
        print(
            f"\n\n**Something Went Wrong !! Please Re-Run the program with newly created file.\n{file_name[:-5]} Auto.xlsx")
    finally:
        print("\n\n*** Saving File...")
        data_excel.save(f"./{file_name[:-5]} Auto.xlsx")
        input("\n*** File Saved !!! ***")
        driver.quit()


def decision(search_results, Results, cleaned_plan):  # Takes Final Decision
    if search_results == 1:
        return (Results['EIN'][0] if (Results['EIN'][0] != "") else "Beacon: Found without EIN")

    count = [0, [], []]
    for i in range(search_results):
        cleaned_company = clean_plan(Results['Company'][i])
        if cleaned_company == cleaned_plan:
            count[0] += 1
            count[1].append(i)
            count[2].append(Results['EIN'][i])

    ein = list(set(count[2]))  # Remove Duplicate EIN's
    for i in range(len(ein)):
        l_ein = len(ein[i])
        if 6 < l_ein < 9:
            ein[i] = '0' * (9 - l_ein) + str(ein[i])

    if len(ein) == 1:
        return (ein[0] if (ein[0] != "") else "Beacon: Found without EIN")
    elif len(ein) == 2:
        return ((ein[1] if (ein[0] == "") else ein[0]) if "" in ein else "Mannual Check is Required")
    return "Mannual Check is Required"


def email_n_password():  # Asking User Email & Password if Unable to Login
    try:
        with open(Path("login.txt"), 'r') as file:
            lines = file.readlines()
            email = lines[0]
            password = lines[1]
    except:
        email = input("Enter Login Email    : ")
        password = input("Enter Login Password : ")
    return email, password


def clean_plan(plan=''):  # Remove the words & special characters to search
    temp = plan.lower().split(' ')
    plan = ''
    for word in temp:
        if word in removals:
            continue
        elif 1 < len(word) < 3:
            word = ' '.join(list(word))
        else:
            r = ["'s", ',', '&', '-', '.']
            for i in r:
                word = word.replace(i, " ")
            plan += word + ' '
        if word == ' ':
            pass
        elif word[-1] == 's':
            word = word[:-1]
    return plan.strip()


def excel_file_name_input():  # Select Excel File to be worked upon
    print("Excel files in same directory:")
    files = listdir('.')
    xl_files = []
    counter = 0
    for file in files:
        if file[-4:] == 'xlsx':
            counter += 1
            xl_files.append(file)
            print(f"{counter}. {file}")
    if counter == 1:
        return xl_files[0]
    elif counter == 0:
        input("\n**No Excel Files Found.\n  Closing Program !!")
        exit()
    option = int(
        input(f"\nEnter the file number corresponding to file name: ")) - 1
    if len(xl_files) > option >= 0:
        return xl_files[option]
    else:
        input("\n\n**Wrong Input !!!\n  Closing Program !!")
        exit()


def default_file_format():
    with open("format.txt", 'w') as file:
        file.write(default_removals+'\n'+"Matchtype = A"+"\n"+"Helper_ID = B"+'\n'+"Source = C"+'\n'+"Sponsor_Name = F"+'\n'+"Plan_Name = G"+'\n'+"Reg1 = S"+'\n'+"Reg2 = T"+'\n'+"Reg3 = U"+'\n'+"Reg4 = V"+'\n' +
                   "PARSED_ORG_NAME = Y"+'\n'+"PARTY_ID = BE"+'\n'+"PLAN_ID = BF"+'\n'+"Date_Column = AP"+'\n'+"Initial = AQ"+'\n'+"Comment = BD"+'\n'+"P.Checked = BA"+'\n'+"P.Search  = BB"+'\n'+"P.Results = BC"+'\n'+"EOF")


main()
