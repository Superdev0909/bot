import pyautogui
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support import ui
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import *
from time import sleep
import json
import random
import string
from datetime import datetime
import sys
import os

def pressTab(count : int):
    for i in range(count):
        pyautogui.hotkey('tab')
        sleep(0.1)                    

def pressShiftTab(count : int):
    for i in range(count):
        pyautogui.hotkey('shiftleft', 'tab')
        sleep(0.1)     

def pressSpace():
    pyautogui.press('space')
    sleep(0.1)

def pressEnter():
    pyautogui.press('enter')
    sleep(1)

def pressDown(count : int):
    for i in range(count):
        pyautogui.press('down')
        sleep(0.1)

def cut():
    pyautogui.hotkey('ctrl', 'x')
    sleep(0.1)

def copy():
    pyautogui.hotkey('ctrl', 'c')
    sleep(0.1)

def paste():
    pyautogui.hotkey('ctrl', 'v')
    sleep(0.1)

def selectAll():
    pyautogui.hotkey('ctrl', 'a')
    sleep(0.1)

def typing(value : str):
    pyautogui.typewrite(value)

def wait_url(driver : webdriver.Chrome, url : str):
    print(url)
    while True:
        cur_url = driver.current_url
        if cur_url == url:
            break
        sleep(0.1)  


def next(driver : webdriver.Chrome):
    find_element(driver, By.CLASS_NAME, 'air3-btn-primary').click()

def find_element(driver : webdriver.Chrome, whichBy, unique : str) -> WebElement:
    while True:
        try:
            element = driver.find_element(whichBy, unique)
            break
        except:
            pass
        sleep(1)
    return element

def find_elements(driver : webdriver.Chrome, whichBy, unique : str) -> list[WebElement]:
    while True:
        try:
            elements = driver.find_elements(whichBy, unique)
            break
        except:
            pass
        sleep(0.1)
    return elements


def waitInfinite(callback, debug=False, callNum=20):
    sleep(1)
    for i in range(callNum):
        try:
            callback()
            break
        except NoSuchElementException as e:
            print("{} on line {}".format(str(e).split('\n')[0], sys.exc_info()[-1].tb_lineno))
            sleep(1)
            pass
        except JavascriptException as e:
            print("{} on line {}".format(str(e).split('\n')[0], sys.exc_info()[-1].tb_lineno))
            sleep(1)
            pass
        except StaleElementReferenceException as e:
            print("{} on line {}".format(str(e).split('\n')[0], sys.exc_info()[-1].tb_lineno))
            sleep(1)
            pass
        except ElementClickInterceptedException as e:
            print("{} on line {}".format(str(e).split('\n')[0], sys.exc_info()[-1].tb_lineno))
            sleep(1)
            pass
        except ElementNotInteractableException as e:
            print("{} on line {}".format(str(e).split('\n')[0], sys.exc_info()[-1].tb_lineno))
            sleep(1)
            pass
        except Exception as e:
            print("{} on line {}".format(str(e).split('\n')[0], sys.exc_info()[-1].tb_lineno))

def waitUntil(callback, driver, selector):
    sleep(1)
    yet = True
    print('wait until')
    while yet:
        try:
            callback(driver.execute_script("x=document.querySelectorAll('{}').length;return document.querySelectorAll('{}')[x-1]".format(selector, selector)))
            yet = False
        except Exception as e:
            print(str(e).split('\n')[0])
            sleep(1)
            pass

def clickByMouse(driver, element):
    ActionChains(driver).click(element)\
                        .perform()

def selectDropDown(driver, itemSelector, country):
    nations = driver.find_elements(By.CSS_SELECTOR, itemSelector)

    if str(type(country)) == "<class 'int'>":
        driver.execute_script(f'document.querySelectorAll("{itemSelector}")[{str(country)}].click()')
    else:
        for i in range(len(nations)):
            try:
                if nations[i].text.find(country) >= 0:
                    driver.execute_script(f'document.querySelectorAll("{itemSelector}")[{str(i)}].click()')
                    break
            except:
                pass

def selectDateDropDown(driver, dropdownId, itemSelector, country):
    tmp = dropdownId.split('##')
    if len(tmp) == 2:
        dropdownId = tmp[0]
        driver.execute_script(f'document.querySelectorAll(\'div[aria-labelledby^="{dropdownId}"]\')[{tmp[1]}].click()')
    else:
        driver.execute_script(f'document.querySelector(\'div[aria-labelledby^="{dropdownId}"]\').click()')
    sleep(1)
    nations = driver.find_elements(By.CSS_SELECTOR, itemSelector)

    if str(type(country)) == "<class 'int'>":
        driver.execute_script(f'document.querySelectorAll("{itemSelector}")[{str(country)}].click()')
    else:
        for i in range(len(nations)):
            try:
                if nations[i].text.find(country) >= 0:
                    driver.execute_script(f'document.querySelectorAll("{itemSelector}")[{str(i)}].click()')
                    break
            except:
                pass


def get_email(count_to_be_created : int) -> list:
    emails = []
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    driver = webdriver.Chrome(options=chrome_options)

    driver.get('https://yopmail.com/email-generator')

    username = driver.find_element(By.CLASS_NAME, 'segen').text;
    driver.switch_to.frame('ifdoms')
    domains = driver.find_elements(By.TAG_NAME, 'option')
    count = 0
    for index, item in enumerate(domains):
        domain = item.text
        if '@' in domain:
            count += 1
            emails.append({"name" : username + domain, "verified" : False})
            if count > count_to_be_created + 50:
                break
    driver.quit()
    print('emails was created.')
    return emails

def verify_email(emails : list, profiles : list, count_to_be_created : int):
    accounts = []

    for profile_name in profiles:
        with open(f'profiles/{profile_name}.json', 'r') as file:
            profile = json.load(file)
        country = profile['photo_others']['country']
        first_names = profile['first_name']
        last_names = profile['last_name']
        count = 0
        while count < count_to_be_created:
            NotVerifedEmailExist = False
            for index, item in enumerate(emails):
                if item['verified'] == False:
                    email = item['name']
                    selected_email_index = index
                    NotVerifedEmailExist = True
                    break
            if NotVerifedEmailExist:
                driver = webdriver.Chrome()
                driver.get('https://www.upwork.com/nx/signup/?dest=home')

                wait_url(driver, 'https://www.upwork.com/nx/signup/?dest=home')
                driver.find_element(By.ID, "button-box-4").click()     
                sleep(0.3)
                driver.find_element(By.CSS_SELECTOR, "button[data-qa='btn-apply']").click()    
                sleep(1)

                first_name = first_names[random.randint(0, len(first_names)-1)]
                last_name = last_names[random.randint(0, len(last_names)-1)]
                driver.find_element(By.ID, "first-name-input").send_keys(first_name)
                sleep(0.3)
                driver.find_element(By.ID, "last-name-input").send_keys(last_name)
                sleep(0.3)
                driver.find_element(By.ID, 'redesigned-input-email').send_keys(email)
                sleep(0.3)
                all_chars = string.ascii_letters + string.digits
                password = random.choice(string.ascii_letters) + random.choice(string.digits)
                for _ in range(10):
                    password += random.choice(all_chars)
                password = ''.join(random.sample(password, len(password)))
                driver.find_element(By.ID, "password-input").send_keys("superstar_0909")
                sleep(0.3)
                span = driver.find_element(By.CLASS_NAME, 'up-dropdown-toggle-title').find_element(By.TAG_NAME, 'span')
                driver.execute_script("arguments[0].innerText = arguments[1];", span, country)
                sleep(0.3)
                driver.execute_script('document.querySelectorAll("span.up-checkbox-replacement-helper")[1].click()')
                sleep(0.3)
                find_element(driver, By.ID, 'button-submit-form').click()
                sleep(3)

                try:
                    driver.find_element(By.CLASS_NAME, 'up-alert')
                    print('rejected', email)
                except:
                    wait_url(driver, 'https://www.upwork.com/nx/create-profile/')
                    accounts.append({"profile" : profile_name, "email" : email, "password" : password, "name" : first_name + ' ' + last_name, "applied"  : False})
                    with open('verified_emails.json', 'w') as file:
                        json.dump(accounts, file)
                    count += 1            
                    print('created', email)
                emails[selected_email_index]['verified'] = True
                driver.quit()
            else:
                print('emails was verified')
    print('emails was verified.')

def login(driver, email, password):
    driver.get('https://www.upwork.com/ab/account-security/login')
    wait_url(driver, 'https://www.upwork.com/ab/account-security/login')
    find_element(driver, By.ID, 'login_username').send_keys(email)
    sleep(0.3)
    find_element(driver, By.ID, 'login_password_continue').click()
    sleep(1)
    find_element(driver, By.ID, 'login_password').send_keys(password)
    sleep(0.3)
    find_element(driver, By.ID, 'login_control_continue').click()
    
def get_started(driver):
    wait_url(driver, 'https://www.upwork.com/nx/create-profile/')
    try:
        driver.execute_script('document.querySelector(\'button[data-qa="get-started-btn"]\').click()')
    except:
        pass

def select_experience(driver, index=2):
    if index < 0:
        index = 0
    if index > 2:
        index = 2
    types = ['NEW_TO_ME', 'NEEDS_TIP', 'FREELANCED_BEFORE']
    wait_url(driver, 'https://www.upwork.com/nx/create-profile/experience')
    find_element(driver, By.CSS_SELECTOR, f"input[value=\"{types[index]}\"]").click()
    next(driver)

def select_what_is_my_goal(driver, index=0):
    if index < 0:
        index = 0
    if index > 3:
        index = 3
    types = ['MAIN_INCOME', 'MONEY_ON_SIDE', 'GET_EXPERIENCE', 'EXPLORING']
    wait_url(driver, 'https://www.upwork.com/nx/create-profile/goal')
    find_element(driver, By.CSS_SELECTOR, f"input[value=\"{types[index]}\"]").click()
    next(driver)

def select_work_preference(driver):
    wait_url(driver, 'https://www.upwork.com/nx/create-profile/work-preference')
    try: find_element(driver, By.CSS_SELECTOR, "input[type='checkbox']").click()
    except: find_element(driver, By.CSS_SELECTOR, "input[type='checkbox']").click()
    next(driver)

def select_manualmode(driver):
    wait_url(driver, 'https://www.upwork.com/nx/create-profile/resume-import')
    find_element(driver, By.CLASS_NAME, 'flex-column').find_elements(By.TAG_NAME, 'button')[3].click()

def add_professional(driver, title):
    wait_url(driver, 'https://www.upwork.com/nx/create-profile/title')
    driver.find_element(By.TAG_NAME, 'input').send_keys(title)
    next(driver)

def add_experience(driver, experiences):
    curr_year = datetime.now().year
    expFlag = True
    for experience in experiences:
        if expFlag:
            waitInfinite(lambda: driver.execute_script('document.querySelectorAll("button.air3-btn.air3-btn-circle")[0].click()'))
        else:
            waitInfinite(lambda: driver.execute_script('document.querySelector("button.air3-btn.air3-btn-secondary.air3-btn-circle").click()'))

        sleep(1)

        if expFlag:
            waitInfinite(lambda: driver.find_elements(By.CSS_SELECTOR, 'input[aria-labelledby="title-label"]')[1].send_keys(""))
        waitInfinite(lambda: driver.find_elements(By.CSS_SELECTOR, 'input[aria-labelledby="title-label"]')[1].send_keys(experience['role']))
        waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="company-label"]').send_keys(experience['company']))
        waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="location-label"]').send_keys(experience['location']))
        selectDateDropDown(driver, "location-label", "span.air3-menu-item-text", experience['country'])

        start_year = eval(experience['start'].split('.')[0])
        start_month = eval(experience['start'].split('.')[1])
        selectDateDropDown(driver, "start-date-month", "span.air3-menu-item-text", start_month - 1)
        selectDateDropDown(driver, "start-date-year", "span.air3-menu-item-text", curr_year - start_year + 1)
        
        if experience['end'] == 'current':
            driver.execute_script('document.querySelector(\'[data-qa="currently-working"]\').querySelector("label").click()')
        else:
            end_year = eval(experience['end'].split('.')[0])
            end_month = eval(experience['end'].split('.')[1])
            selectDateDropDown(driver, "end-date-month", "span.air3-menu-item-text", end_month - 1)
            selectDateDropDown(driver, "end-date-year", "span.air3-menu-item-text", curr_year - end_year + 1)
        
        for text in experience['description']:
            waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'textarea[aria-labelledby="description-label"]').send_keys("• " + text + "\n"))

        driver.execute_script('document.querySelector(\'button[data-qa="btn-save"]\').click()')
        expFlag = False
    next(driver)

def add_education(driver, educations, action):
    for education in educations:
        start = eval(education['start'])
        end = eval(education['end'])
        curr_year = datetime.now().year
 
        waitInfinite(lambda: driver.execute_script('document.querySelector(\'button[data-qa="education-add-btn"]\').click()'))

        waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="school-label"]').click())
        sleep(1)
        waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="school-label"]').send_keys(education['university']))
        sleep(1)
        try:
            driver.find_element(By.CLASS_NAME, "air3-menu-item-text").click()
        except:
            pass

        waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="degree-label"]').click())
        sleep(1)
        waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="degree-label"]').send_keys(education['degree']))
        sleep(1)

        action.move_to_element_with_offset(driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="degree-label"]'), 50, 70)
        action.click()
        action.perform()

        waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="area-of-study-label"]').click())
        sleep(1)
        waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="area-of-study-label"]').send_keys(education['field']))
        sleep(1)
        action.move_to_element_with_offset(driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="area-of-study-label"]'), 50, 70)
        action.click()
        action.perform()

        waitInfinite(lambda: selectDateDropDown(driver, "dates-attended-label##0", "span.air3-menu-item-text", curr_year - start + 4))
        waitInfinite(lambda: selectDateDropDown(driver, "dates-attended-label##1", "span.air3-menu-item-text", 2030 - end + 4))
        waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'textarea[aria-labelledby="description-label"]').send_keys(education['description']))
        driver.execute_script('document.querySelector(\'button[data-qa="btn-save"]\').click()')
    next(driver)

def add_language(driver, languages):
    waitInfinite(lambda: selectDateDropDown(driver, "dropdown-label-english", "span.air3-menu-item-text", 2))
    count = 0
    for i in languages:
        language = i['language']
        pro = eval(i['level'])
        driver.execute_script('document.querySelector("button.air3-btn.air3-btn-secondary.air3-btn-sm").click()')
        sleep(2)
        waitInfinite(lambda: selectDateDropDown(driver, f"dropdown-label-language-{count}", "span.air3-menu-item-text", language))
        waitInfinite(lambda: selectDateDropDown(driver, f"dropdown-label-proficiency-{count}", "span.air3-menu-item-text", pro))
        count += 1
    next(driver)

def add_skills(driver, skills):
    for skill in skills:
        waitUntil(lambda x: x.click(), driver, f'input[aria-labelledby="skills-input"]')
        for i in skill:
            driver.find_element(By.CSS_SELECTOR, f'input[aria-labelledby="skills-input"]').send_keys(i)
            sleep(0.1)

        flag = True
        while flag:
            nations = driver.find_elements(By.CSS_SELECTOR, "span.air3-menu-item-text")
            flag = len(nations) == 0

        for i in range(len(nations)):
            try:
                if nations[i].text.find(skill) >= 0:
                    driver.execute_script(f'document.querySelectorAll("span.air3-menu-item-text")[{str(i)}].click()')
                    break
            except:
                pass
        waitUntil(lambda x: x.clear(), driver, f'input[aria-labelledby="skills-input"]')
    next(driver)

def add_overview(driver, content):
    wait_url(driver, 'https://www.upwork.com/nx/create-profile/overview')
    for item in content:
        find_element(driver, By.TAG_NAME, 'textarea').send_keys(item + '\n')
    next(driver)

def add_service(driver, services):
    waitUntil(lambda x: x.click(), driver, 'div[data-test="dropdown-toggle"]')
    sleep(1)

    for service in services:
        driver.execute_script(f'''
            // document.querySelectorAll(\'div[data-test="dropdown-toggle"]\')[3].click()
            var services = document.querySelectorAll('span.air3-menu-checkbox-labels');
            var toselect;
            for (let i = 0; i < services.length; i++) {{
                console.log(services[i], '{service}');
                if (services[i].textContent.indexOf('{service}') >= 0) {{
                    toselect = services[i].parentNode.parentNode;
                    break;
                }}
            }}
            if (toselect) {{
                if (toselect.getAttribute("aria-selected") == 'false') {{
                    toselect.parentNode.parentNode.parentNode.click();
                    setTimeout(() => toselect.click(), 300);
                }}
            }}
        ''')
        sleep(0.1)
    next(driver)

def add_rate(driver, rate):
    wait_url(driver, 'https://www.upwork.com/nx/create-profile/rate')
    find_element(driver, By.CSS_SELECTOR, 'input[data-test="currency-input"]').send_keys(rate)
    next(driver)

def add_photo_others(driver, others, action):
    waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="date-of-birth-label"]').send_keys(others['birthday']))
    waitInfinite(lambda: selectDateDropDown(driver, "country-label", "span.air3-menu-item-text", others['country']))
    waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="street-label"]').send_keys(others['street']))

    waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="city-label"]').click())
    sleep(1)
    waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="city-label"]').send_keys(others['city']))
    sleep(1)
    action.move_to_element_with_offset(driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="city-label"]'), 50, 70)
    action.click()
    action.perform()

    waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[aria-labelledby="postal-code-label"]').send_keys(others['zipcode']))
    # waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[data-ev-label="phonenumber_input"]').send_keys(others['phone']))
    waitInfinite(lambda: driver.execute_script("document.querySelector('button[data-qa=\"open-loader\"]').click()"))
    sleep(1)
    avatar_count = len(others['avatar'])
    waitInfinite(lambda: driver.find_element(By.CSS_SELECTOR, 'input[type="file"]').send_keys(os.path.join(os.getcwd(), 'avatars', '{}'.format(others['avatar'][random.randint(0, avatar_count - 1)]))))
    waitInfinite(lambda: driver.execute_script("document.querySelectorAll('button.air3-btn.air3-btn-primary')[document.querySelectorAll('button.air3-btn.air3-btn-primary').length - 1].click()"))
    while True:
        sleep(1)
        try:
            driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
        except:
            break
    # next(driver)

def submit_profile(driver):
    wait_url(driver, 'https://www.upwork.com/nx/create-profile/submit')
    