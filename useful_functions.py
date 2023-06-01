from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

def extract_value(gl_dict, num): # извлечение МНН из global_dict

    mnn_value = gl_dict.get(num)[0] # эта функция извлекает из dict = global_dict (скрипт mnn_dict_res) значение num
   
    input_value = list(mnn_value) ; input_value[0] = input_value[0].upper() ; input_value = "".join(mnn_value)

    table_value = mnn_value.split(" ") ; table_value = "_".join(table_value) 

    return [input_value, table_value] # возвращает МНН препарата для вставки в строку input и для вставку в таблицу бд type = string

def captcha_freeze(driver, tag_url):

    try:
        WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, f"""{tag_url}""")))
        time.sleep(2)
        driver_find = driver.find_element(By.XPATH, f"""{tag_url}""")
        time.sleep(2)
        driver_find = driver_find.text.lower()
    
        if "код" in driver_find:
            print_key = input("""Captcha has been spotted. Once you solved it
            print any key to go on with the script.\n""")

            time.sleep(10)
            return "Captcha has been passed."
    except TimeoutException:
        return "No captcha found"

def clear_request(input_field): # очистка поля ввода МНН
    #принимает в себя тип данных Web_element 
    input_field.send_keys(Keys.CONTROL+"a")
    input_field.send_keys(Keys.DELETE)      

def delete_quotes(text): 

    if '\"' in text: # эта часть кода отвечает за обработку и замену символа двойных кавычек "" на 
            # одинарные '' во избежание ошибок компилляции 
        text = list(text)
        for i in range(len(text)):
            if text[i] == '\"':
                text[i] = "\'"
        text = "".join(text)
    
    return text
    
def after_click_prevention(link_text, page_object, tag_statement):
    match page_object:
            case "page_element":
                try:
                    driver.find_element(By.XPATH, tag_statement)
                    return True
                except NoSuchElementException:
                    print("Page recurssion")
                    click_element_state(link_text, page_object, tag_statement)
            case "other_element":
                attr_tg = driver.find_element(By.XPATH, tag_statement).get_attribute("class")
                if attr_tg == "active":
                    # print("Tag is active")
                    return True
                else:
                    print("Instruction or production recurssion")
                    click_element_state(link_text, page_object, tag_statement)
            case None:
                pass

def click_element_state(link_text, page_object, tag_statement): 

    for i in range(15):        
            time.sleep(2)
            try:
                text_find = driver.find_element(By.LINK_TEXT, link_text) 
                #print(text_find).text
                if text_find.is_displayed() == True and text_find.is_enabled() == True:
                    #print((text_find.is_displayed(), text_find.is_enabled()))
                    text_find.click()
                    time.sleep(7)
                    return after_click_prevention(link_text, page_object, tag_statement)
                else:
                    try:
                        print("Element not in expected state.")
                        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.LINK_TEXT, link_text)))
                        continue
                    except TimeoutException:
                        print("Element vanished.\nTrying to find it...")
                        continue
            except (ElementNotInteractableException, NoSuchElementException):
               print("Failed to click-on the element...")
               continue
            except ElementClickInterceptedException:
                driver.execute_script("arguments[0].click();", text_find)
                time.sleep(7)
                print("Trying to click-on the element...")
                return after_click_prevention(link_text, page_object, tag_statement)