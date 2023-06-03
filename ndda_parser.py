# АВТОР СКРИПТА: Александр Самохин. 
# СОАВТОРЫ СКРИПТА: нету

#_________________________________________________________________________________________________________________
# Блок импорта необходимых библиотек
from selenium import webdriver
import chromedriver_autoinstaller
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException, NoSuchElementException, ElementClickInterceptedException
import time 
import sqlite3
import sys 
import tempfile
import zipfile
import numpy as np 
import os 
import docx 
import beepy 
sys.path.append("C:\\Users\\user\\Desktop\\IT-Project\\ProjectsPC\\SPR_Project\\Mnn_dict_files")
sys.path.append("C:\\Users\\user\\Desktop\\IT-Project\\ProjectsPC\\SPR_Project\\Spr-scripts")
from mnn_dict_res import global_dict 
from useful_functions import extract_value, delete_quotes

#_________________________________________________________________________________________________________________
# Блок подключения к сайту-реестру grls и задания основных значений 

URL = r"http://register.ndda.kz/category/search_prep"
chromedriver_autoinstaller.install()
opt = webdriver.ChromeOptions()
# opt.add_argument('headless')
driver = webdriver.Chrome(options=opt)
opt.add_experimental_option("excludeSwitches", ['enable-logging'])
driver.get(URL)
driver.switch_to.frame("iframe1")
time_val = 2
sec = 5
count = 0 
start_value = 0 
element_count = 0
cash_list_one = [] ; cash_list_two = [] # списки с пропущенными\обработанными элементами
hash_table_missing_data = {"missing_mnn_value":[]} # словарь пропущенных элементов и элементов с нулевыми значениями 
hash_table_existing_data = {"existing_mnn_value":[]} # словарь существующих элементов 
# cash_list_one\cash_list_two контролируют найденные и отсуствующие значения
character = "" ; buffer = "" ; num_buffer = 0 
opt.headless = True

#_________________________________________________________________________________________________________________
# Этот блок работает с БД

def table_create(): 

# принимает МНН значение name взятое из global_dict 
    try:
        with sqlite3.connect(r"C:\Users\user\Desktop\IT-Project\ProjectsPC\SPR_Project\Tables\ndda\ndda.db") as db:
            sql = db.cursor()
            sql.execute("""CREATE TABLE IF NOT EXISTS ndda_drugs ( 
                    "Trade_name_rus" TEXT, 
                    "Registrator_tran" TEXT,
                    "Registrator_country" TEXT,
                    "Producer_tran" TEXT,
                    "Producer_country" TEXT,
                    "Dosage_form_full_name" TEXT,
                    "Dose" TEXT,
                    "Sc_name" TEXT,
                    "Recipe_status" TEXT,
                    "As_name_rus" TEXT)""") # в таблицу добавляются основные столбцы-фильтры: уникальный id;
            # торговое название; название компании; страна производителя;
            # форма принятия препарата; дозировка препарата (в будущем возможны дополнения по столбцам)
    except:
        print("Error occured while creating the table.")
        k = input("Print any key to exit...")
    print(f"Table ndda_drugs has been successfully created.\n\n")

def table_add(list_of_information): 

    with sqlite3.connect(r"C:\Users\user\Desktop\IT-Project\ProjectsPC\SPR_Project\Tables\ndda\ndda.db") as db:
            sql = db.cursor()

            # try:
            sql.execute(f'''INSERT INTO ndda_drugs ("Trade_name_rus", "Registrator_tran", "Registrator_country", 
                    "Producer_tran", "Producer_country", "Dosage_form_full_name", "Dose", 
                    "Sc_name", "Recipe_status", "As_name_rus") VALUES 
                    ("{list_of_information[0]}", "{list_of_information[1]}", "{list_of_information[2]}", "{list_of_information[3]}", 
                    "{list_of_information[4]}", "{list_of_information[5]}", "{list_of_information[6]}", "{list_of_information[7]}", 
                    "{list_of_information[8]}", "{list_of_information[9]}")''') 
            
#_________________________________________________________________________________________________________________
# Начало парсинг-блока

# Почти все функции ниже рекурсивны, поскольку элементы, указанные в качестве названия 
# функций должны ОБЯЗАТЕЛЬНО присуствовать во фрейме страницы.
# Паузы добавлены для того, чтобы контролировать процесс рекурсии.
# В случае с информацией о производителях, 

def trade_name(): 

    try:
        WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, """//*[@id="reestr-form-reestrForm-form"]/table/tbody/tr[4]/td[2]""")))
        text_find = driver.find_element(By.XPATH, """//*[@id="reestr-form-reestrForm-form"]/table/tbody/tr[4]/td[2]""")
    except TimeoutException:
        k = input("Failed to find the information about trade_name...")
        trade_name()
        
    text_find = delete_quotes(text_find.text)
    
    return text_find

def company_name_reg(inform):  
    
    if len(inform) == 0:
        return "Не указано"

    try:
        number_of_string = inform.index("Держатель регистрационного удостоверения") + 1

        WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, f"""//*[@id="yw0"]/table/tbody/tr[{number_of_string}]/td[2]""")))
        text_find = driver.find_element(By.XPATH, f"""//*[@id="yw0"]/table/tbody/tr[{number_of_string}]/td[2]""")
    except TimeoutException:
        k = input("Failed to find the information about company_name_reg...")
        company_name_reg(inform)
    except ValueError:
        return "Не указано"

    text_find = delete_quotes(text_find.text)

    return text_find

def reg_country(inform): 

    if len(inform) == 0:
        return "Не указано"

    try:
        number_of_string = inform.index("Держатель регистрационного удостоверения") + 1
    
        WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, f"""//*[@id="yw0"]/table/tbody/tr[{number_of_string}]/td[4]""")))
        text_find = driver.find_element(By.XPATH, f"""//*[@id="yw0"]/table/tbody/tr[{number_of_string}]/td[4]""")
        text_find = text_find.text
    except TimeoutException:
        k = input("Failed to find the information about reg_country...")
        reg_country(inform)
    except ValueError:
        return "Не указано"

    return text_find

def company_name_prod(inform): 

    if len(inform) == 0:
        return "Не указано"

    try:
        number_of_string = inform.index("Производитель") + 1
        
        WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, f"""//*[@id="yw0"]/table/tbody/tr[{number_of_string}]/td[2]""")))
        text_find = driver.find_element(By.XPATH, f"""//*[@id="yw0"]/table/tbody/tr[{number_of_string}]/td[2]""")
        # //*[@id="yw0"]/table/tbody/tr[1]/td[2]
    except TimeoutException:
        k = input("Failed to find the information about company_name_prod...")
        company_name_prod(inform)
    except ValueError:
        return "Не указано"

    text_find = delete_quotes(text_find.text)
    
    return text_find

def prod_country(inform): 

    if len(inform) == 0:
        return "Не указано"

    try:
        number_of_string = inform.index("Производитель") + 1
        
        WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, f"""//*[@id="yw0"]/table/tbody/tr[{number_of_string}]/td[4]""")))
        text_find = driver.find_element(By.XPATH, f"""//*[@id="yw0"]/table/tbody/tr[{number_of_string}]/td[4]""")
        text_find = text_find.text
    except TimeoutException:
        k = input("Failed to find the information about company prod_country...")
        prod_country(inform)
    except ValueError:
        return "Не указано"

    return text_find

def dosage_form(): 

    try:
        WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, """//*[@id="reestr-form-reestrForm-form"]/table/tbody/tr[8]/td[2]""")))
        text_find = driver.find_element(By.XPATH, """//*[@id="reestr-form-reestrForm-form"]/table/tbody/tr[8]/td[2]""")
    except TimeoutException:
        k = input("Failed to find the information about company dosage_form...")
        dosage_form()
    
    text_find = delete_quotes(text_find.text)

    return text_find

def dosage(): 

    try:
        WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, """//*[@id="reestr-form-reestrForm-form"]/table/tbody/tr[10]/td[2]""")))
        text_find = driver.find_element(By.XPATH, """//*[@id="reestr-form-reestrForm-form"]/table/tbody/tr[10]/td[2]""")
        text_find = text_find.text
    except TimeoutException:                       # //*[@id="reestr-form-reestrForm-form"]/table/tbody/tr[10]/td[2]/text()
        k = input("Failed to find the information about dosage...")
        dosage()

    return text_find

def containment_condition(): 

    try:

        WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, """//*[@id="yw3"]/table/tbody/tr/td[3]/a""")))
        
        click_element_state("Скачать", None, None)

        text_find = driver.find_element(By.XPATH, """//*[@id="yw3"]/table/tbody/tr/td[3]/a""")
        text_find = text_find.get_attribute("href") ; text_find = text_find.split("/") 
        doc_index = text_find[len(text_find) - 1]
        time.sleep(sec)

        return word_processing(doc_index) # здесь обрабатывается документ ворда 
        
    except TimeoutException:
        k = input("Failed to find the information about containment_condition...")
        containment_condition()

def recipe(): 
    # в тех препаратах, что выдаются с рецептом, есть checked="checked". У безрецептурных препов такого нету.
    WebDriverWait(driver, time_val).until(EC.presence_of_element_located(((By.XPATH, """//*[@id="recipe_sign"]"""))))

    try:
        WebDriverWait(driver, time_val).until(EC.presence_of_element_located(((By.XPATH, """//*[@id="recipe_sign"]"""))))
        text_find = driver.find_element(By.XPATH, """//*[@id="recipe_sign"]""").get_attribute("checked")
        if text_find == "true":
            return "По рецепту"
    except TimeoutException:
        return "Без рецепта"
    
    # //*[@id="recipe_sign"]

def active_substance(): 

    try:
        WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, """//*[@id="reestr-form-reestrForm-form"]/table/tbody/tr[6]/td[2]""")))
        text_find = driver.find_element(By.XPATH, """//*[@id="reestr-form-reestrForm-form"]/table/tbody/tr[6]/td[2]""")
        text_find = text_find.text
    except TimeoutException:
        k = input("Failed to find the information about company active_substance...")
        active_substance()
    
    return text_find

# Конец парсинг-блока
#_________________________________________________________________________________________________________________


#_________________________________________________________________________________________________________________
# Начало функционального блока

def word_processing(index): 

    flag = False
    path = r"C:\Users\user\Downloads\doc_" + f"{index}.zip"
    if os.path.exists(path):
        with tempfile.TemporaryDirectory() as tmpdir: # создаётся временная директория, в которую будет извлечен zip-файл
            with zipfile.ZipFile(path, 'r') as zip_ref: # далее с помощью данной конструкции открывается zip-файл
                # посредством пути к данному zip-файлу - path
                zip_ref.extractall(tmpdir) # Извлекает все файлы из архива во временную директорию, 
                # используя метод extractall() объекта ZipFile.
            file = os.listdir(tmpdir)[0] # далее мы помещаем все имена файлов в список и извлекаем оттуда только самый
            # первый файл

            try:
                doc = docx.Document(os.path.join(tmpdir, file))
                
                text = []

                for para in doc.paragraphs:

                    if "Условия отпуска из аптек" in para.text:
                        break

                    if "Условия хранения" in para.text:
                        flag = True
                        continue
                        
                    if flag == True:
                        if para.text == "Условия отпуска из аптек":
                            break
                        text.append(para.text)
                        
                text = '\n'.join(text)
                return text

            except ValueError:
                return "Нет доступа к файлу инструкции."
            
    else:
        print("Directory does not exist.")

def value_check(): # проверка на состояние найденного лекарства
    
    WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, """//*[@id="register_pager_right"]/div""")))
    state_find = driver.find_element(By.XPATH, """//*[@id="register_pager_right"]/div""")
    state_find = state_find.text
    return state_find
    
def value_click(tag_elem): # алгоритм кликает на элемент в табличке 
    
    WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, """//*[@id="ReestrTableForNdda_ls_mnn"]""")))
    click_find = driver.find_element(By.XPATH, """//*[@id="ReestrTableForNdda_ls_mnn"]""")
    click_find.click()
    time.sleep(2)
    tag_elem.click()

    WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, """//*[@id="reestr-form-reestrForm-form"]/input[1]""")))
    search_find = driver.find_element(By.XPATH, """//*[@id="reestr-form-reestrForm-form"]/input[1]""")
    search_find.click()
    time.sleep(sec)

def elements_number(): 

    elements_count = driver.execute_script('return document.querySelectorAll("#yw0 tbody tr.odd, #yw0 tbody tr.even").length;')
    # подсчет элементов таблицы по производителям

    return elements_count

def character_assignment(char): # ускорение алгоритма засчет смены цикла в зависимости от буквы

    match char:
        case "а":       
            return 3, 688 
        case "б":
            return 689, 1196
        case "в":
            return 1197, 1346
        case "г":
            return 1347, 1653
        case "д":
            return 1654, 2254
        case "ж":
            return 2255, 2262
        case "з":
            return 2263, 2352
        case "и":
            return 2353, 2663
        case "й":
            return 2664, 2736
        case "к":
            return 2737, 3226
        case "л":
            return 3227, 3578
        case "м":
            return 3579, 4169
        case "н":
            return 4170, 4564
        case "о":
            return 4565, 4823
        case "п":
            return 4824, 5521
        case "р":
            return 5522, 5800
        case "с":
            return 5801, 6278
        case "т":
            return 6279, 7068
        case "у":
            return 7069, 7104
        case "ф":
            return 7105, 7613
        case "х":
            return 7106, 7717
        case "ц":
            return 7718, 8036
        case "э":
            return 8037, 8503

def value_compare(mnn_value, character): # сравнивание препарата с предложенным в табличке
    global num_buffer
    global count
    global start_value

    if num_buffer == 0:
        start, end = character_assignment(character)
    else:
        start = num_buffer
        end = character_assignment(character)[1]
    
    for i in range(start, end + 1):

        WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, f"""//*[@id="ReestrTableForNdda_ls_mnn"]/option[{i}]""")))
        value_find_tag = driver.find_element(By.XPATH, f"""//*[@id="ReestrTableForNdda_ls_mnn"]/option[{i}]""")
        value_find = value_find_tag.text.lower()
        print("Value to analyse is: ", value_find)

        if value_find == mnn_value:
            print()
            print(f"Key value was founded: {value_find}.")
            equality = True
            num_buffer = i

            return equality, value_find_tag, i
        elif i != end and value_find != mnn_value:
            equality = False
            num_buffer = i

            if mnn_value[0:2].lower() == value_find[0:2].lower():
                count += 1
                
                if count == 1:
                    start_value = num_buffer

                try:
                    WebDriverWait(driver, time_val).until(EC.presence_of_element_located((By.XPATH, f"""//*[@id="ReestrTableForNdda_ls_mnn"]/option[{i + 1}]""")))
                    next_tag = driver.find_element(By.XPATH, f"""//*[@id="ReestrTableForNdda_ls_mnn"]/option[{i + 1}]""")
                    next_tag = next_tag.text

                    if value_find[0:2].lower() != next_tag[0:2].lower(): # или до тех пор пока первый символ mnn_value 
                        # равен первому символу next_tag
                        num_buffer = start_value
                        return equality, None, i
                except IndexError: # или до тех пор пока не кончатся тэги
                    num_buffer = start_value
                    return equality, None, i
            else:
                count = 0 
                num_buffer = start_value

            continue
        elif i == end and equality == False:
            num_buffer = start_value
            return equality, None, i 

def inner_find(page_state, element_number): # пробежка по всем элементам паспорта препарата
    global element_count

    try:
        element_count += 1
    # Главная страница паспорта
        trade_name_inf = trade_name() 
        dosage_form_inf = dosage_form() 
        dosage_inf = dosage()
        active_substance_inf = active_substance()
        recipe_inf = recipe()

    # Страница паспорта с производителями и странами производства 
        while page_state == False:
            page_state = click_element_state('Производитель', "other_element", """//*[@id="yw4"]/ul/li[3]""")
        
        page_state = False
        num = elements_number() 
        inform = [""]

        while "" in inform:
            inform = list(map(lambda x: driver.find_element(By.XPATH, f"""//*[@id="yw0"]/table/tbody/tr[{x}]/td[5]""").text, range(1, num + 1)))

        company_name_reg_inf = company_name_reg(inform)
        reg_country_inf = reg_country(inform)
        company_name_prod_inf = company_name_prod(inform)
        prod_country_inf = prod_country(inform)

    # Страница паспорта с инструкцией
        while page_state == False:     
            page_state = click_element_state('Инструкции', "other_element", """//*[@id="yw4"]/ul/li[6]""")

        containment_condition_inf = containment_condition()
    except:
        l = input(f"Problem occured with {input_value}, №{element_number}. Stopped on the element {trade_name_inf}. Restart needed from {element_count}.")

    print(f"Finished with {trade_name_inf}.")

    table_add([trade_name_inf, company_name_reg_inf, reg_country_inf, company_name_prod_inf, prod_country_inf, 
               dosage_form_inf, dosage_inf, containment_condition_inf, recipe_inf, active_substance_inf])

def passport_click(number_of_element): # действия внутри фрейма с элементами лекарства
    global start_fix_element, end_fix_element

    tag_tr = driver.find_elements(By.TAG_NAME, "tr") ; tag_tr = np.array(list(map(lambda x: x.get_attribute("id") if "jqgrow ui-row-ltr" in x.get_attribute("class") else None, tag_tr))) 
    tag_id_res = tag_tr[tag_tr != None] ; count_row = len(tag_id_res) 
    iter_id = list(tag_id_res) ; iter_id.reverse()
    namings = list(map(lambda x: driver.find_element(By.XPATH, f"""//*[@id="{x}"]/td[1]/a""").text, iter_id)) ; namings.reverse()
    state = False

    if interval_parse == False:
        start_fix_element = 1
        end_fix_element = len(namings)

    if start_fix_element > 1 or end_fix_element < len(namings):
        start_fix_element -= 1 ; end_fix_element -= 1
        namings = namings[start_fix_element:end_fix_element + 1]
        count_row = len(namings)
    else:
        pass

    for i in range(count_row):
        popped_el = namings.pop()

        try:
            while state == False:
                state = click_element_state(popped_el, "page_element", """//div[contains(@style, 'block')]""")
            else:
                state = False
                inner_find(state, number_of_element)
        except: # //tagname[ends-with(@attribute, 'value')] /html/body/div[2][ends-with(@style, 'block;')]
            k = input(f"""Problem occured while clicking to drug-pass of {input_value}. Stopped on №{i + 1} element. Value of the 
                  element is {number_of_element}.""")
            
        click_element_state("""Закрыть""", None, None)
  
        time.sleep(sec)

    # чтобы понять какой индекс имеет строка нужно делать перебор по тэгам tr role = "row". Нужно посчитать сколько всего row 
    # на фрейме и пробежаться по каждому, сохранив id строки в стэк и потом, из стека pop-ать каждый элемент так, что стек самостоятельно обнуляется.

def after_click_prevention(link_text, page_object, tag_statement): # проверка на состояние после клика

    match page_object:
            case "page_element":
                try:
                    driver.find_element(By.XPATH, tag_statement)
                    return True
                except NoSuchElementException:
                    click_element_state(link_text, page_object, tag_statement)
            case "other_element":
                attr_tg = driver.find_element(By.XPATH, tag_statement).get_attribute("class")
                if attr_tg == "active":
                    return True
                else:
                    click_element_state(link_text, page_object, tag_statement)
            case None:
                pass

def click_element_state(link_text, page_object, tag_statement): # комплексный алгоритм клика на ссылки

    for i in range(15):        
            time.sleep(sec)
            try:
                text_find = driver.find_element(By.LINK_TEXT, link_text) 
                if text_find.is_displayed() == True and text_find.is_enabled() == True:
                    text_find.click()
                    time.sleep(sec)
                    return after_click_prevention(link_text, page_object, tag_statement)
                else:
                    try:
                        print("Element not in expected state...")
                        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.LINK_TEXT, link_text)))
                        continue
                    except TimeoutException:
                        print("Element vanished.\nTrying to find it...")
                        continue
            except (ElementNotInteractableException, NoSuchElementException):
               print("Failed to click-on the element...")
               continue
            except ElementClickInterceptedException:
                time.sleep(7)
                driver.execute_script("arguments[0].click();", text_find)
                print("Trying to click-on the element...")
                return after_click_prevention(link_text, page_object, tag_statement)
    
    for i in range(4): # если страница заморозится, то будет проигран 4 раза 
        # определенный звук
        beepy.beep(0)

    k = input(f"Something went wrong with {link_text}, try to reload frame and enter any key...")
    click_element_state(link_text, page_object, tag_statement)

# Конец функционального-блока
#_________________________________________________________________________________________________________________

#_________________________________________________________________________________________________________________
# Тело алгоритма. Сюда мы помещаем все функции так, чтобы они корректно работали 

#_________________________________________________________________________________________________________________
# Ниже идёт заключительный алгоритм 

start = int(input("Print down the start-value.\n")) ; end = int(input("Print down the end-value.\n"))

interval_parse = input("""Do you want to parse an interval of values?
                       Print 'Yes' if you want, else print 'No'.\n""")

if interval_parse == "Yes":
    start_fix_element = int(input("Print down an element number to start with.\n"))
    end_fix_element = int(input("Print down an element number to end with.\n"))
else:
    interval_parse = False
    pass

table_create()

for i in range(start, end + 1):

    if i != start:
        interval_parse = False

    input_value = extract_value(global_dict, i)[0] ; table_value = extract_value(global_dict, i)[1]

    print("Input value is", input_value)
    print("Table value is", table_value)
    print()

    character = list(input_value)[0] ; character = character.lower() 

    print("Cash table one", cash_list_one, end="\n\n") ; print("Cash table two", cash_list_two, end="\n\n")   

    if buffer != character:
        cash_list_one.clear() ; cash_list_two.clear()

    buffer = character
    equality, tag_object, cash_num = value_compare(input_value, character)

    if equality == True:  
        value_click(tag_object)
        state = value_check()
        
        if state != "Нет записей для просмотра": # тут блок работы с паспортом препарата
           hash_table_existing_data.get("existing_mnn_value").append(input_value)
           print("Elements exist.")
           
           passport_click(i)

           print(f"Element {input_value} was successfully parsed.")
           element_count = 0
           driver.back()
           cash_list_one.append(cash_num)
           driver.refresh()
           time.sleep(2)
           driver.switch_to.frame("iframe1")

           continue
        else: # этот блок отсылает обратно, при отсуствии значения 
            print("Elements do not exist.")
            driver.back()
            time.sleep(2)
        
            cash_list_one.append(cash_num) ; hash_table_missing_data.update({cash_num:input_value})

            driver.refresh()
            time.sleep(2)
            driver.switch_to.frame("iframe1")
            continue
    else:
        cash_list_two.append(input_value) ; hash_table_missing_data.get("missing_mnn_value").append(input_value) 
        continue

print("Cash table one", cash_list_one, end="\n\n") ; print("Cash table two", cash_list_two, end="\n\n")   
print("Hash table of missing values =", hash_table_missing_data)
print("Hash table of existing values =", hash_table_existing_data)

l = input("End of op. Press any key to continue...")
driver.close()
