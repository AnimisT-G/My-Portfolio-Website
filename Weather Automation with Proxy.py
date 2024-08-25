from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException
from time import sleep
from ctypes import *
import threading
import os
import zipfile


def main():
    # Read the proxies from proxies.txt file present at same location where the program is located.
    # In proxies.txt the ip:port:username:password is kept.
    with open("proxies.txt", 'r') as p:
        available_proxies = p.readlines()

    # Read the locations from locations.txt file present at same location where the program is located.
    # In locations.txt the URL of the weather news of particular place is kept.
    print(available_proxies)
    with open("locations.txt", 'r') as l:
        locations = l.readlines()

    # threads will store all the threads to be runned for later reference
    threads = []

    use_proxies = input("With Proxies ? (Y) - ")

    if use_proxies in ('y', "Y"):
        # Loop to run through all the proxies available to do the testing
        for p in range(len(locations)):
            # unpacking the ip:port:username:password of the proxy
            PROXY_IP, PROXY_PORT, PROXY_USER, PROXY_PASS = available_proxies[p].split(
                ':')
            PROXY_PASS = PROXY_PASS[:-1]
            PROXY_PORT = int(PROXY_PORT)

            # json file for the pluginfile zip to be created later
            manifest_json = """
    {
        "version": "1.0.0",
        "manifest_version": 2,
        "name": "Chrome Proxy",
        "permissions": [
            "proxy",
            "tabs",
            "unlimitedStorage",
            "storage",
            "<all_urls>",
            "webRequest",
            "webRequestBlocking"
        ],
        "background": {
            "scripts": ["background.js"]
        },
        "minimum_chrome_version":"22.0.0"
    }
    """

            # js file for the pluginfile zip to be created later
            background_js = """
    var config = {
            mode: "fixed_servers",
            rules: {
            singleProxy: {
                scheme: "http",
                host: %s,
                port: parseInt(%s)
            },
            bypassList: ["localhost"]
            }
        };

    chrome.proxy.settings.set({value: config, scope: "regular"}, function() {});

    function callbackFn(details) {
        return {
            authCredentials: {
                username: %s,
                password: %s
            }
        };
    }

    chrome.webRequest.onAuthRequired.addListener(
                callbackFn,
                {urls: ["<all_urls>"]},
                ['blocking']
    );
    """ % (PROXY_IP, PROXY_PORT, PROXY_USER, PROXY_PASS)

            # creating and storing the thread in threads list and running it.
            threads.append(threading.Thread(target=thread_func, args=(
                locations[p][:-1], p, use_proxies, manifest_json, background_js)))
            threads[p].start()
    else:
        for p in range(len(locations)):
            threads.append(threading.Thread(target=thread_func,
                           args=(locations[p][:-1], p)))
            threads[p].start()

    # clearing the console to make sure nothing is printed on the console screen
    os.system('cls')
    # weather details that will be printed on the screen
    print_at(0, 0, "%-30s%-7s%-15s%-15s%-15s%-15s%-5s" % ("Location", "Time", "Weather", "Temperature", "Real Feel", "Air Quality", "AQI"))


def thread_func(location, thread_no, proxies=None, manifest_json=None, background_js=None):  # Thread function
    if proxies == None:
        driver = get_chromedriver()
    else:
        driver = get_chromedriver(
            use_proxy=True, manifest_json=manifest_json, background_js=background_js)
    # opening the website
    driver.get(location)

    # coordinates of console to print the data on screen in a format
    x = 2 + thread_no
    y = [30,7,15,15,15,15,5]
    # getting the data from the website which has to be displayed in interval of x seconds in sleep(x).
    while True:
        try:
            current_weather = {
                "Location": driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div[1]/div/a[2]/h1").text,
                "Time": driver.find_element(By.XPATH, "/html/body/div[1]/div[5]/div[1]/div[1]/a[1]/div[1]/div[1]/p").text,
                "Weather": driver.find_element(By.XPATH, "/html/body/div[1]/div[5]/div[1]/div[1]/a[1]/div[2]/span[1]").text,
                "Temperature": driver.find_element(By.XPATH, "/html/body/div[1]/div[5]/div[1]/div[1]/a[1]/div[1]/div[1]/div/div/div[1]").text,
                "Real_Feel": driver.find_element(By.XPATH, "/html/body/div[1]/div[5]/div[1]/div[1]/a[1]/div[1]/div[1]/div/div/div[2]").text,
                "Air_Quality": driver.find_element(By.XPATH, "/html/body/div[1]/div[5]/div[1]/div[1]/div[1]/div/div[2]/div[2]/h3/p[1]").text,
                "AQI": driver.find_element(By.XPATH, "/html/body/div[1]/div[5]/div[1]/div[1]/div[1]/div/div[2]/div[1]/div/div/div/div[1]").text
            }

            print_at(x, 0, "%-30s" % current_weather['Location'])
            print_at(x, sum(y[:1]), "%-7s" % current_weather['Time'])
            print_at(x, sum(y[:2]), "%-15s" % current_weather['Weather'])
            print_at(x, sum(y[:3]), "%-15s" % current_weather['Temperature'])
            print_at(x, sum(y[:4]), "%-15s" %(current_weather['Real_Feel'][-3:] + "C"))
            print_at(x, sum(y[:5]), "%-15s" % current_weather['Air_Quality'])
            print_at(x, sum(y[:6]), "%-5s" % current_weather['AQI'])

            # sleeping this thread
            sleep(15)
            # refreshing the page
            driver.refresh()
        except (StaleElementReferenceException, NoSuchElementException):
            print("Exception")


# get_chromedriver function to setup the chrome webdriver with proxy
def get_chromedriver(use_proxy=False, user_agent=None, manifest_json=None, background_js=None):
    path = os.path.dirname(os.path.abspath(__file__))
    chrome_options = webdriver.ChromeOptions()
    if use_proxy:
        pluginfile = 'proxy_auth_plugin.zip'

        # creating zipfile for the plugin to be used in chrome webdriver
        with zipfile.ZipFile(pluginfile, 'w') as zp:
            zp.writestr("manifest.json", manifest_json)
            zp.writestr("background.js", background_js)
        chrome_options.add_extension(pluginfile)
        # chrome_options.headless = True

    chrome_options.add_experimental_option(
        'excludeSwitches', ['enable-logging'])

    driver = webdriver.Chrome(options=chrome_options)
    return driver


STD_OUTPUT_HANDLE = -11


class COORD(Structure):
    pass


COORD._fields_ = [("X", c_short), ("Y", c_short)]


def print_at(r, c, s):  # printing function to print at coordinates x , y on the console
    h = windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
    windll.kernel32.SetConsoleCursorPosition(h, COORD(c, r))

    c = s.encode("windows-1252")
    windll.kernel32.WriteConsoleA(h, c_char_p(c), len(c), None, None)

main()
