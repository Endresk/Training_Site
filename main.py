import json
import os
import re
import os.path
import glob
import sys
import time
import asyncio
import aiohttp
import pandas as pd
import openpyxl as ox
import xlwings as xw
import requests
import win32com
from selenium.common.exceptions import NoSuchElementException
from win32com.client import Dispatch
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Alignment
# from fake_useragent import UserAgent
from openpyxl.utils.dataframe import dataframe_to_rows
from selenium.webdriver import ActionChains
from openpyxl.utils import get_column_letter
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from tabulate import tabulate
from selenium.webdriver.common.by import By
from selenium import webdriver
from win32com.universal import com_error
from xlwings.quickstart_fastapi.app import app
#
# useragent = UserAgent()


class Learning:

    def __init__(self):

        self.branches = 'branches'
        self.specialties = 'specialties'

        self.headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                          '(KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36',
            'Accept': 'text/plain, */*; q=0.01'
        }

        options = webdriver.ChromeOptions()
        # options.add_argument('--headless')
        # options.headless = True
        options.add_argument("start-maximized")
        options.add_argument("disable-infobars")
        options.add_argument("--no-sandbox")
        options.add_argument('--disable-gpu')

        # options.add_argument('--blink-settings=imagesEnabled=false')
        # options.add_argument("--disable-javascript")
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                             "(KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36")

        options.add_argument("--disable-blink-features-AutomationControlled")
        options.add_argument(f"user-data-dir={os.getcwd()}\\Learning")
        options.add_argument('--enable-aggressive-domstorage-flushing')
        options.add_argument('--enable-profile-shortcut-manager')
        options.add_argument('--profile-directory=Profile 1')
        options.add_argument('--profiling-flush=n')
        options.add_argument('--allow-profiles-outside-user-dir')

        self.driver = webdriver.Chrome(executable_path="chromedriver", options=options)

        # self.driver.execute_script("document.body.style.zoom='60%'")
        self.driver.set_window_size(1920, 1080)

        self.education, self.num = self.region()

    def stop(self):
        self.driver.close()
        self.driver.quit()
        sys.exit()

    def region(self):

        try:
            education = str(input('Введите "1" Высшее или (enter) СПО: '))

            if education == '1':
                print("Выбран раздел 'Высшее'")
                education = 'higher'
                url_regions = 'https://postupi.online/ajax_cons.php?mode=load_modal&ckmod=modal_cities&is_spo=0'

                default = '      68: Алтайский край\n' \
                          '      71: Кемеровская область\n' \
                          '      72: Новосибирская область\n' \
                          '      73: Омская область\n' \
                          '      74: Томская область\n'

                num = {'68': 'Алтайский край',
                       '71': 'Кемеровская область',
                       '72': 'Новосибирская область',
                       '73': 'Омская область',
                       '74': 'Томская область'
                       }
            else:
                print("Выбран раздел 'СПО'")
                education = 'secondary'
                url_regions = 'https://postupi.online/ajax_cons.php?mode=load_modal&ckmod=modal_cities&is_spo=1'

                default = '      68: Алтайский край\n' \
                          '      71: Кемеровская область\n' \
                          '      72: Новосибирская область\n' \
                          '      73: Омская область\n'

                num = {'68': 'Алтайский край',
                       '71': 'Кемеровская область',
                       '72': 'Новосибирская область',
                       '73': 'Омская область',
                       }

            response_r = requests.get(url=url_regions, headers=self.headers)
            soup = BeautifulSoup(response_r.text, 'lxml')

            block_region = soup.find('ul', class_='list-unstyled m-choice-region')
            block_city = soup.find('ul', class_='list-unstyled m-choice-city')

            num_region = {}
            for i in block_region:
                try:
                    data_obl = i.get('data-obl')
                    if data_obl is not None:
                        name_region = i.text
                        num_region[data_obl] = name_region
                except:
                    pass

            list_regions = [f'{k}: {v}' for k, v in num_region.items()]
            n = 4
            list_regions_n = [list_regions[i:i + n] for i in range(0, len(list_regions), n)]
            print(tabulate(list_regions_n, tablefmt="github"))

            input_text = 'Введите числа регионов согласно перечню через пробел'

            while True:

                numbers_region_input = str(input('  \nЕсли ничего не вводить нажать enter то, \n'
                                                 '  По умолчанию регионы:\n'
                                                 f'{default}'
                                                 f'{input_text}: '))

                if numbers_region_input == "":
                    break
                else:
                    numbers_region = numbers_region_input.split()
                    num = {}

                    for i in numbers_region:
                        i = i.replace(',', '').replace('.', '').replace('/', '').replace(':', '').replace(';', '')
                        i = i.replace(';', '').replace('+', '').replace('-', '').replace('/', '').replace('*', '')
                        if i.isdigit():
                            if i in num_region:

                                num[i] = num_region[i]
                            else:
                                num[i] = '-1'
                                print(f"Значение {i} нет в перечне")
                        else:
                            num[i] = '-1'

                    if '-1' in num.values():
                        input_text = "Введите заново числа регионов через пробел"
                    else:
                        break

            area = []
            for i in block_city:
                try:
                    data_obl = i.attrs['data-obl']

                    if data_obl in num:
                        code_name = f'{data_obl} - {num_region[data_obl]}'

                        data_chpu = i.span['data-chpu']
                        area.append(
                            {
                                "city": {
                                    "code_name": f"{code_name}",
                                    "name": f"{i.text}",
                                    "name_": f"{data_chpu}",
                                    "code": f"{data_obl}"
                                }
                            }
                        )
                    else:
                        continue
                except:
                    pass

            json.dump(area, open(f"city_{education}.txt", "w"))

            return education, num

        except Exception as f:
            print("--- Ошибка в регионе ---", f)
            self.stop()

    def Branches(self):

        try:
            os.makedirs('files_city')
        except FileExistsError:
            pass

        driver = self.driver

        driver.get('https://postupi.online/')

        try:
            driver.find_element(By.ID, 'cabinet')
        except:
            driver.get('https://postupi.online/158/')

            while True:
                if driver.execute_script("return document.readyState") == 'complete':
                    break

            driver.find_element(By.CSS_SELECTOR, '#regent-form > div > div.reg-inner > small > span').click()
            driver.find_element(By.CSS_SELECTOR, '#regent-form > div > div.enter-inner > div > button > span').click()

            username_input = driver.find_element(By.ID, 'user_emailNew')
            username_input.send_keys('andredjlee@gmail.com')
            time.sleep(1)
            password_input = driver.find_element(By.ID, 'user_pswrdNew')
            password_input.send_keys('GZRvaxv7')
            time.sleep(25)

        print("--- Этап 'Отрасль'  ---")

        try:
            with open(f"city_{self.education}.txt", "r") as read_file:
                file_city = json.load(read_file)

                for city in file_city:

                    for k, v in dict(city).items():
                        name = v['name']
                        name_ = v['name_'].replace("-", "_")
                        name_url = v['name_']
                        code_name = v['code_name']
                    print('-', name)

                    file = f'{self.education}/{code_name}/{name_}/{self.branches}/{name_}_{self.branches}.txt'

                    if not os.path.exists(file) or os.path.getsize(file) < 15:

                        print(f"--- Словарь отраслей ({name}) пуст или не существует, добавляем... ---")

                        if self.education == 'higher':
                            url_branches = f'https://postupi.online/'
                            mode = 'vo_main_upload_nd'
                        else:
                            url_branches = f'https://postupi.online/spo/'
                            mode = 'spo_main_upload_nd'

                        response_branches = requests.get(url=url_branches, headers=self.headers)
                        soup_branches = BeautifulSoup(response_branches.text, 'lxml')

                        html_div_branches = soup_branches.find(
                            'div', attrs={'class': "direction-wrap"})

                        html_a_branches = html_div_branches.find_all('a', href=True)
                        html_branches_all = [item['href'] for item in html_a_branches]

                        html_urls_branches = []

                        try:

                            url_branches_others = f'https://postupi.online/ajax.php?mode={mode}'
                            response_branches_others = requests.get(url=url_branches_others, headers=self.headers)
                            data_branches_others = response_branches_others.text

                            if data_branches_others != "":

                                soup_branches_others = BeautifulSoup(data_branches_others, 'lxml')
                                html_branches_others = soup_branches_others.find_all('a', href=True)
                                [html_branches_all.append(item['href']) for item in html_branches_others]

                                for item in html_branches_all:
                                    text_razdel = item.split("//")[1]
                                    url_razdel = f'https://{name_url}.{text_razdel}'

                                    response_spec = requests.get(url=url_razdel, headers=self.headers)
                                    soup_spec = BeautifulSoup(response_spec.text, 'lxml')
                                    selector_spec = '#main_form > div.content-wrap > div.content > div.list-cover'
                                    class_list_cover = soup_spec.select_one(selector_spec)

                                    if class_list_cover is not None:
                                        html_urls_branches.append(url_razdel)

                            else:
                                pass
                        except:
                            print('Ошибка цикла отраслей')

                        try:
                            os.makedirs(f'{self.education}/{code_name}/{name_}')
                        except FileExistsError:
                            pass

                        try:
                            os.makedirs(f'{self.education}/{code_name}/{name_}/{self.branches}')
                        except FileExistsError:
                            pass

                        json.dump(html_urls_branches, open(
                            f"{self.education}/{code_name}/{name_}/{self.branches}/{name_}_{self.branches}.txt", "w"))
                    else:
                        continue
            print("--- Ссылки на Отрасли все получены --- \n")

        except Exception as f:
            print("--- Ошибка в функции 'Отрасль' (блок общий) --- \n", f)
            self.stop()

    def Specialties(self):
        print("--- Этап 'Разделы отраслей с получением их ссылок на специальности' ---")
        list_code_center = ['01', '02', '03', '04', '05']

        try:
            regions = sorted(os.listdir(f'{self.education}/'))
            list_regions = [all_regions for code_region, name_region in self.num.items()
                            for all_regions in regions
                            if re.search(f'{code_region}', all_regions) is not None]

            for region in list_regions:
                cities = sorted(os.listdir(f'{self.education}/{region}'))

                for city in cities:

                    dir_city = city.find('.')
                    if dir_city == -1:

                        print('-', city)
                        file_branch = glob.glob(f'{self.education}/{region}/{city}/{self.branches}/{city}_branches.txt')

                        for filename_branch in file_branch:
                            with open(filename_branch, "r") as read_file:
                                file_urls_specialties = json.load(read_file)

                        try:
                            os.makedirs(f'{self.education}/{region}/{city}/{self.branches}/{self.specialties}')
                        except FileExistsError:
                            pass

                        for urls_speciality in file_urls_specialties:

                            chapter_speciality = urls_speciality.replace("-", "_").split('/')[-2]

                            file = f'{self.education}/{region}/{city}/{self.branches}/{self.specialties}/{chapter_speciality}.txt'

                            if not os.path.exists(file) or os.path.getsize(file) == 0:

                                response_speciality = requests.get(url=urls_speciality, headers=self.headers)
                                soup_speciality_pag = BeautifulSoup(response_speciality.text, 'lxml')

                                try:
                                    html_pag = soup_speciality_pag.find('div', attrs={'class': "invite fetcher"})
                                    html_pags = html_pag.find_all('a')
                                    html_pags = [item_pag['href'] for item_pag in html_pags][-2]
                                    pag = int(str(html_pags).split("=")[-1])
                                except:
                                    pag = 1

                                list_speciality = {}

                                for num in range(0, pag):
                                    response_speciality = requests.get(url=f'{urls_speciality}?page_num={num + 1}',
                                                                       headers=self.headers)
                                    soup_speciality = BeautifulSoup(response_speciality.text, 'lxml')
                                    html_codes = soup_speciality.select('.list div.list__info')

                                    for html_code in html_codes:
                                        ''' Ссылки с одним учреждением '''
                                        one_code = html_code.select_one('.flex-nd.list__info-inner div:nth-child(1)')
                                        text_code = one_code.select_one('.list__pre span:nth-child(3) a').text

                                        ''' Цифра которая в середине кода специальности '''
                                        center_code = str(re.findall(r"\.([^.]+)\.", text_code)).replace("'", '')[1:-1]

                                        if center_code in list_code_center:
                                            html_a_code = one_code.select_one('.list__h a')
                                            url_speciality = html_a_code['href']
                                            text_speciality = html_a_code.text

                                            if self.education == 'higher':
                                                z = '/vuz/'
                                                programma = 'programma'
                                                one_speciality_html = '.list-var__info div h2 a'
                                            else:
                                                z = '/ssuz/'
                                                programma = 'programma-spo'
                                                one_speciality_html = \
                                                    '.list__info div.flex-nd.list__info-inner div h2 a'

                                            ''' Проверка ссылки имеется ли несколько учреждений '''
                                            if url_speciality.find(z) == -1:
                                                ''' Ссылки с несколькими учреждениями '''
                                                two_code = html_code.select_one('.list__btn.list__btn_extra ')

                                                ''' Цифры специальности '''
                                                program_code = url_speciality.split('/')[-2]

                                                ''' Получаем ссылку на варинаты учреждений по специальности '''
                                                html_variants_speciality = two_code.select_one(
                                                    '.btn-violet-nd')
                                                url_variants_speciality = html_variants_speciality['href']

                                                ''' Получаем варианты специальности '''
                                                response_variants_speciality = requests.get(
                                                    url=url_variants_speciality, headers=self.headers)
                                                soap_variants_speciality = BeautifulSoup(
                                                    response_variants_speciality.text, 'lxml')
                                                html_variants_speciality = soap_variants_speciality.select(
                                                    '.content div.list-cover ul li.list')

                                                list_dupl = []

                                                for one_speciality in html_variants_speciality:

                                                    one_speciality_text = one_speciality.select_one(
                                                        one_speciality_html)
                                                    one_speciality_href = one_speciality_text['href']
                                                    dupl_u = one_speciality_href.split('/')[-2]

                                                    if dupl_u in list_dupl:
                                                        pass
                                                    else:
                                                        list_speciality[
                                                            f'{one_speciality_href}{programma}/{program_code}/'] = \
                                                            text_speciality

                                                        list_dupl.append(dupl_u)
                                                        # print(f'{one_speciality_href}{programma}/{program_code}/')

                                            else:
                                                # print(url_speciality)
                                                list_speciality[url_speciality] = text_speciality
                                        else:
                                            print("--- Код не подходит --- ", text_code, center_code)

                                json.dump(list_speciality, open(
                                    f"{self.education}/{region}/{city}/{self.branches}/{self.specialties}/{chapter_speciality}.txt",
                                    "w"))
            print("--- Ссылки на специальности все получены ---")

        except Exception as f:
            print("--- Ошибка в функции 'В получении ссылок на специальности' (блок общий)  ---", f)
            self.stop()

    async def programs(self):
        print("--- Этап 'Сбор данных по специальностям' ---")

        try:

            driver = self.driver

            headers = [
                "Город",
                "Общежитие",
                "Вуз/Ссуз",
                'Наименование учреждения',
                'Полное наименование учреждения',
                'Контакты',
                'Адрес',
                'База обучения',
                'Форма обучения',
                'Отрасль',
                'Код',
                'Направление',
                'Специальность',
                'Предметы для поступления',
                'Профессии',
                'Проходной балл (Бюджет)',
                'Бюджетных мест',
                'Проходной балл (Платно)',
                'Платных мест',
                'Количество лет обучения',
                'Стоимость обучения'
                #, 'Дисциплины'
            ]

            column_list = {
                0: 'A1', 1: 'B1', 2: 'C1', 3: 'D1', 4: 'E1', 5: 'F1', 6: 'G1', 7: 'H1', 8: 'I1', 9: 'J1', 10: 'K1',
                11: 'L1', 12: 'M1', 13: 'N1', 14: 'O1', 15: 'P1', 16: 'Q1', 17: 'R1', 18: 'S1', 19: 'T1', 20: 'U1',
                21: 'V1', 22: 'W1', 23: 'X1', 24: 'Y1'
            }

            ''' Создаем ссесию '''
            async with aiohttp.ClientSession(headers=self.headers) as session:

                regions = sorted(os.listdir(f'{self.education}/'))
                list_regions = [all_regions for code_region, name_region in self.num.items()
                                for all_regions in regions
                                if re.search(f'{code_region}', all_regions) is not None]

                for region in list_regions:
                    cities = sorted(os.listdir(f'{self.education}/{region}/'))

                    for city in cities:
                        print(city)
                        list_all = []

                        city_dir = city.find('.')
                        if city_dir == -1:

                            file_city = f'files_city/{city}.xlsx'

                            list_all.append(headers)

                            file_parts = glob.glob(
                                f'{self.education}/{region}/{city}/{self.branches}/{self.specialties}/razdel_*.txt')

                            count = 0

                            ''' Перебор всех файлов (раздел) в папке города '''
                            for file_part in file_parts:
                                with open(file_part, "r") as read_file:
                                    urls_speciality_part = json.load(read_file)

                                ''' Перебор всех ссылок в файле соответствующего раздела '''
                                for url_speciality_part, text_speciality in urls_speciality_part.items():

                                    print(text_speciality, url_speciality_part)

                                    ''' Переходим по url раздела (специальности) '''
                                    response_part = await session.get(
                                        url=url_speciality_part)
                                    soup_part = BeautifulSoup(await response_part.text(), 'lxml')

                                    # ''' Получение специальности '''
                                    #
                                    # if re.search(r"[.]", text_speciality):
                                    #     text_speciality = soup_part.select_one('#prTitle').text.split(':')[0]
                                    #     print(text_speciality)

                                    ''' Получение наименование города '''
                                    selector_city = \
                                        '#topRghtMenu > div > div.dropdown.dropdown_city.ddown-choice > a > span'
                                    text_city = soup_part.select_one(selector_city).text

                                    ''' Получение полного наименования учреждения '''
                                    selector_about_university = '#main_form > div.bg-nd > ' \
                                                               'div.bg-nd__main > ol > li:nth-child(1) > a'
                                    url_about_university = soup_part.select_one(selector_about_university)['href']

                                    response_about_university = await session.get(
                                        url=url_about_university)
                                    soup_about_university = BeautifulSoup(await response_about_university.text(), 'lxml')

                                    text_full_university = soup_about_university.select_one('#prTitle').text

                                    selector_dormitory = '#main_form > div.content-wrap > div.content > ' \
                                                         'section.section-box.hideshow-wrap.section-box-flex > ' \
                                                         'div.card-nd-pre-wrap > div.card-nd-pre'
                                    text_all_dormitory = soup_about_university.select(selector_dormitory)

                                    text_dormitory = ["Да" if re.search(r'Общежитие', str(i))
                                                      else "Нет" for i in text_all_dormitory][0]

                                    ''' Определяем тэг с ссылкой на "Контакты" '''
                                    try:
                                        html_contact = soup_part.find('a', class_='menu-internal__link contacts-icon')
                                        url_contact = html_contact['href']

                                        ''' Переходим по адресу "Контакты" (открывается новая ссылка) '''
                                        response_contact = await session.get(url=url_contact)
                                        soup_contact = BeautifulSoup(await response_contact.text(), 'lxml')

                                        ''' Получение url сайта, почты, телефона и адрес вуза '''
                                        html_url_site_contact_ = soup_contact.find(
                                            'span', class_='contact-icon contact-icon_sm site')
                                        html_mail_contact = soup_contact.find(
                                            'span', class_='contact-icon contact-icon_sm mail')
                                        text_phone_contact = soup_contact.find(
                                            'span', class_='contact-icon contact-icon_sm phone').text
                                        text_address_contact = soup_contact.find(
                                            'span', class_='contact-icon contact-icon_sm address').text

                                        text_contact = (f'Сайт: {html_url_site_contact_.a.text} \n'
                                                        f'Почта: {html_mail_contact.a.text} \n'
                                                        f'Телефон: {text_phone_contact} ')
                                    except:
                                        text_contact = 'Список контактов пуст'

                                    ''' Определяем тэг с ссылкой на "Профессии" '''
                                    try:
                                        html_professions = soup_part.find('a',
                                                                          class_='menu-internal__link profession-icon')
                                        url_professions = html_professions['href']

                                        ''' Переходим по адресу "Профессии" (открывается новая ссылка) '''
                                        response_professions = await session.get(url=url_professions)
                                        soup_professions = BeautifulSoup(await response_professions.text(), 'lxml')
                                        html_professions = soup_professions.select(
                                            '#main_form > div.content-wrap > div.content > '
                                            'div.list-cover > ul > li > div.list-col__info > h2')
                                        list_professions = ''
                                        ''' Получаем список профессий '''
                                        for hp in html_professions:
                                            list_professions += f'{hp.text} \n'
                                        list_professions = list_professions.rstrip()
                                    except:
                                        list_professions = 'Список профессий пуст'

                                    # ''' Получение "Профессиональных дисциплин" '''
                                    # selector_prof_discip = '#main_form > div.content-wrap > div.content > ' \
                                    #                        'section.section-box.hideshow-wrap.section-box-flex > ' \
                                    #                        'div.descr-max'
                                    # html_prof_discip = soup_part.select_one(selector_prof_discip)
                                    #
                                    # prof_discip = []
                                    # div = []
                                    #
                                    # number = 0
                                    # for i in html_prof_discip:
                                    #
                                    #     if i.name is None:
                                    #         continue
                                    #     else:
                                    #         if i.name == 'p' or 'h' in i.name:
                                    #             p_tag = i.text
                                    #
                                    #         elif i.name == 'div':
                                    #             div.append(i.get_text())
                                    #
                                    #             for i in div:
                                    #                 i_rep = i.replace('\n\n\n\n', '\n') \
                                    #                     .replace('\n\n\n', '\n') \
                                    #                     .replace('\n\r\n', '\n') \
                                    #                     .replace('\r\n', '\n')
                                    #                 r_st = i_rep.lstrip()
                                    #
                                    #             prof_discip.append(f'{r_st}')
                                    #
                                    #         elif i.name == 'ul':
                                    #             ul_tag = i.text.replace(";", "")
                                    #
                                    #             try:
                                    #                 if len(prof_discip[number].split('\n')) <= 26:
                                    #
                                    #                     if len(f'{prof_discip[number]} {p_tag} {ul_tag}'.split(
                                    #                             '\n')) > 26:
                                    #                         number += 1
                                    #                         prof_discip.append(
                                    #                             f'{p_tag} {ul_tag}\n')
                                    #                     else:
                                    #                         prof_discip[number] = \
                                    #                             f'{prof_discip[number]} {p_tag} {ul_tag}\n'
                                    #                 else:
                                    #                     number += 1
                                    #                     prof_discip.append(
                                    #                         f'{p_tag} {ul_tag}\n')
                                    #             except:
                                    #                 prof_discip.append(
                                    #                     f'{p_tag} {ul_tag}\n')
                                    #
                                    #                 if len(prof_discip[number].split('\n')) > 26:
                                    #                     number += 1

                                    ''' Получаем отрасль программы '''
                                    selector_branch = '#main_form > div.bg-nd > div.bg-nd__main > ol > ' \
                                                      'li:nth-child(3) > a > span'
                                    text_branch = soup_part.select_one(selector_branch).text

                                    ''' Получаем код и направление специальности '''
                                    text_code_and_text = soup_part.select_one(
                                        '#main_form > div.bg-nd > div.bg-nd__main > p > a:nth-child(2)').text
                                    text_direction = text_code_and_text.split('(')[0]
                                    text_code = text_code_and_text.split('(')[-1].replace(')', '')

                                    ''' Получаем ссылки на все варианты обучения по программе '''
                                    selector_variants = '.section-box.carousel-nd.overflow-wrap ' \
                                                        'div div.swiper-wrapper div.swiper-slide'
                                    html_variants = soup_part.select(selector_variants)

                                    list_all_url_variants = []

                                    for variant in html_variants:
                                        class_a_variant = variant.find('a', class_='swiper-slide__h')
                                        url_variant = class_a_variant['href']

                                        code_url_variant = url_variant.split('#')[0]

                                        if code_url_variant in list_all_url_variants:
                                            pass
                                        else:
                                            list_all_url_variants.append(code_url_variant)

                                    ''' Проходим по всем ссылкам вариантов специальности '''
                                    for url_variant in list_all_url_variants:

                                        # print(url_variant)
                                        driver.get(url_variant)

                                        while True:
                                            if driver.execute_script("return document.readyState") == 'complete':
                                                break

                                        ''' Проверка, что есть БЮДЖЕТ и ПЛАТНО '''
                                        try:
                                            get_Free_Exam = driver.find_element(
                                                By.CSS_SELECTOR, "span[onclick='getFreeExam($(this));']")
                                        except NoSuchElementException:
                                            get_Free_Exam = 'Нет бюджета'
                                        try:
                                            get_Pay_Exam = driver.find_element(
                                                By.CSS_SELECTOR, "span[onclick='getPayExam($(this));']")
                                        except NoSuchElementException:
                                            get_Pay_Exam = 'Нет платного'

                                        selector_detail = '#main_form > div.content-wrap > div.content > ' \
                                                          'section.section-box.hideshow-wrap > div'

                                        ''' Выгружаем через selenium, парсим через bs4 '''

                                        html_detail = driver.find_element(By.CSS_SELECTOR,
                                                                          selector_detail).get_attribute('innerHTML')

                                        soup_detail = BeautifulSoup(html_detail, 'lxml')

                                        ''' Получаем Наименование Вуза/СПО '''
                                        text_university = soup_detail.select_one(
                                            'div:nth-child(1) > div > span').text

                                        ''' Получаем количество бюджетных мест '''
                                        col_free_places = soup_detail.select_one(
                                            'div > div:nth-child(3) > div > span').text

                                        ''' Получаем количество платных мест '''
                                        col_pay_places = soup_detail.select_one(
                                            'div > div:nth-child(4) > div > span').text

                                        ''' Получаем период обучения '''
                                        text_period_study = soup_detail.select_one(
                                            'div > div:nth-child(5) > div > span').text

                                        ''' Получаем стоимость обучения '''
                                        text_price_study = soup_detail.select_one(
                                            'div > div:nth-child(6) > div > span').text

                                        if self.education == 'higher':
                                            sheet_name = 'Высшее'
                                            education = 'Вуз'
                                            selector_score = 'div > div:nth-child(2) > span.score-box__score'

                                            ''' Получаем Базу обучения '''
                                            text_training_of_education = 'После 11 класса'

                                            ''' Получаем Форму обучения '''
                                            text_form_of_education = soup_detail.select_one(
                                                'div > div:nth-child(1) > div > span').text

                                            ''' Получаем Предметы для поступления '''
                                            box_items = driver.find_elements(
                                                By.CSS_SELECTOR, '#main_form > div.content-wrap > div.content > '
                                                                 'section:nth-child(3) > '
                                                                 'div.score-box-wrap.swiper-container > '
                                                                 'div.swiper-wrapper > '
                                                                 'div.score-box.swiper-slide.swiper-slide-next > '
                                                                 'div.score-box__inner  > div.score-box__item')
                                            list_form_of_education = ''

                                            for i in box_items:
                                                try:
                                                    one_box_inner = i.find_elements(
                                                        By.CSS_SELECTOR, 'div > p')
                                                    # list_form_of_education.append(f'{one_box_inner} \n')

                                                    for j in one_box_inner:

                                                        try:
                                                            one_box_inner1 = j.find_element(
                                                                By.CSS_SELECTOR, 'span')
                                                            list_form_of_education += f'{one_box_inner1.text} \n'

                                                        except NoSuchElementException:
                                                            list_form_of_education += f'{j.text} \n'

                                                    two_box = i.find_element(
                                                        By.CSS_SELECTOR, 'div > div.score-box__extra')
                                                    driver.execute_script("arguments[0].style.display = 'block';",
                                                                          two_box)
                                                    p = two_box.find_elements(By.CSS_SELECTOR, 'p')

                                                    for j in p:
                                                        two_box_inner = j.find_element(
                                                            By.CSS_SELECTOR, 'p > span').text
                                                        # list_form_of_education.append(f'{two_box_inner} \n')
                                                        list_form_of_education += f'{two_box_inner} \n'
                                                except NoSuchElementException:
                                                    continue

                                            list_form_of_education = list_form_of_education.rstrip()

                                        else:
                                            sheet_name = 'СПО'
                                            education = 'Ссуз'
                                            selector_score = 'div > div > span.score-box__score'

                                            ''' Получаем Базу обучения '''
                                            text_training_of_education = soup_detail.select_one(
                                                'div:nth-child(5) > div > span').text

                                            ''' Получаем Форму обучения '''
                                            text_form_of_education = soup_detail.select_one(
                                                'div:nth-child(6) > div > span').text

                                            ''' Получаем Предметы для поступления '''
                                            list_form_of_education = driver.find_element(
                                                By.CSS_SELECTOR, '#main_form > div.content-wrap > div.content > '
                                                                 'section:nth-child(3) > '
                                                                 'div.score-box-wrap.swiper-container> '
                                                                 'div.swiper-wrapper > '
                                                                 'div.score-box.swiper-slide.swiper-slide-next > '
                                                                 'div > div > span').text

                                        scoreFree = 'div.score-box-wrap.swiper-container.scoreFree > ' \
                                                    'div.swiper-wrapper > ' \
                                                    'div.score-box.swiper-slide > ' \
                                                    f'{selector_score}'

                                        scorePay = 'div.score-box-wrap.swiper-container.scorePay > ' \
                                                   'div.swiper-wrapper > ' \
                                                   'div.score-box.swiper-slide > ' \
                                                   f'{selector_score}'

                                        ''' Если есть бюджет и платно '''
                                        if get_Free_Exam != 'Нет бюджета' and get_Pay_Exam != 'Нет платного':
                                            # print("Бюджет и платно")
                                            try:
                                                free_passing_score = driver.find_element(
                                                    By.CSS_SELECTOR, scoreFree).text
                                            except NoSuchElementException:
                                                free_passing_score = 'Отсутствует'

                                            driver.execute_script("arguments[0].click();", get_Pay_Exam)

                                            try:
                                                pay_passing_score = driver.find_element(
                                                    By.CSS_SELECTOR, scorePay).text
                                            except NoSuchElementException:
                                                pay_passing_score = 'Отсутствует'

                                            ''' Если есть только бюджет '''
                                        elif get_Pay_Exam != 'Нет бюджета' and get_Pay_Exam == 'Нет платного':
                                            # print("Бюджет")
                                            try:
                                                free_passing_score = driver.find_element(
                                                    By.CSS_SELECTOR, scoreFree).text
                                            except NoSuchElementException:
                                                free_passing_score = 'Отсутствует'
                                            pay_passing_score = 'Нет платного'

                                            ''' Если есть только платно '''
                                        else:
                                            # print("Платно")
                                            free_passing_score = 'Нет бюджета'
                                            try:
                                                pay_passing_score = driver.find_element(
                                                    By.CSS_SELECTOR, scorePay).text
                                            except NoSuchElementException:
                                                pay_passing_score = 'Отсутствует'

                                        block = [f'{text_city}', f'{text_dormitory}', f'{education}',
                                                 f'{text_university}', f'{text_full_university}\n', f'{text_contact}',
                                                 f'{text_address_contact}', f'{text_training_of_education}',
                                                 f'{text_form_of_education}', f'{text_branch}', f'{text_code}',
                                                 f'{text_direction}', f'{text_speciality}',
                                                 f'{list_form_of_education}\n', f'{list_professions}\n',
                                                 f'{free_passing_score}', f'{col_free_places}', f'{pay_passing_score}',
                                                 f'{col_pay_places}', f'{text_period_study}', f'{text_price_study}']

                                        # for i in prof_discip:
                                        #     block.append(i)

                                        list_all.append(block)

                                        # count += 1
                                        #
                                        # if count == 8:

                        pd.set_option('display.max_rows', None)
                        pd.set_option('display.max_columns', None)
                        pd.set_option('display.max_colwidth', None)

                        df = pd.DataFrame(list_all)

                        with xw.App(visible=False) as ap:

                            if not os.path.exists(file_city):
                                wb = ap.books.add()
                                ws = wb.sheets[0]
                                ws.name = sheet_name

                            else:
                                wb = ap.books.open(file_city)

                                ws = wb.sheets[sheet_name] \
                                    if sheet_name in [i.name for i in wb.sheets] \
                                    else wb.sheets.add(sheet_name)

                            ws.range('A1').options(header=False, index=False, na_rep='').value = df

                            last_element = ''

                            # for num in range(2, df.shape[0] + 1, 2):

                            for column in df:
                                max_ = 0

                                max_list = max(
                                    [
                                        max(
                                            [len(x) for x in i.split('\n') if len(x) > max_],
                                            default=0
                                        )
                                        for i in list(
                                        map(str, df[column].astype(str))
                                    )
                                    ]
                                )

                                col = column_list[df.columns.get_loc(column)]

                                last_element = col.split('1')[0]

                                # ws.range(f'{last_element}{num}:{last_element}{num + 1}').merge()
                                ws.range(col).column_width = 64 if max_list > 60 \
                                    else max_list + 4

                            ws.range(
                                f'$A1:${last_element}1').api.HorizontalAlignment = -4108

                            ws.range(
                                f'$A2:${last_element}{df.shape[0]}').api.HorizontalAlignment = -4131

                            ws.range(
                                f'$A2:${last_element}{df.shape[0]}').api.VerticalAlignment = -4160

                            wb.save(path=file_city)

                            print(f"Город {file_city} получен")

                            self.stop()
                            # await asyncio.sleep(44)

        except Exception as f:
            print("--- Ошибка в функции 'Специальности' (блок общий) ---", f)
            self.stop()

    def main(self):
        self.Branches()
        self.Specialties()
        asyncio.run(self.programs())


if __name__ == "__main__":
    learning = Learning()
    learning.main()
