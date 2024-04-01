import requests
import xlsxwriter
import sys
import os
import datetime
import base64
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import concurrent.futures
import threading
from queue import Queue
import pandas as pd
import tkinter as tk
from tkinter import scrolledtext, filedialog, PhotoImage, Spinbox, Label, Button, Frame
import json
import shutil
import tkinter.messagebox as messagebox
from contextlib import contextmanager

stop_threads = False

global_results = Queue()

def update_log(message):
    log_text.config(state=tk.NORMAL)
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)
    log_text.config(state=tk.DISABLED)

def get_region(driver):
    driver.implicitly_wait(10)
    driver.get(f"https://yandex.ru/tune/geo")
    input_element = driver.find_element(By.ID, 'city__front-input')
    current_value = input_element.get_attribute('value')

    if current_value != "Москва":
        input_element.clear()
        input_element.send_keys("Москва")

        dropdown = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "popup__content"))
        )

        moscow_option = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//li[.//div[text()='Москва']]"))
        )
        moscow_option.click()

        save_button = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "form__save"))
        )
        save_button.click()
        print("Установлен регион Москва")
        update_log("Установлен регион Москва")

    if current_value == "Москва":
        print("Регион Москва уже установлен")
        update_log("Регион Москва уже установлен")


def click(k, driver):
    if k == 0:
        url1 = driver.current_url
        if 'Нажмите, чтобы продолжить' in driver.page_source:
            checkbox = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "js-button"))
            )
            checkbox.click()

        if 'Потяните вправо' in driver.page_source:
            print(driver.current_url)
            print("Слайдер")

            driver.refresh()
            slider = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'Thumb')))
            action = ActionChains(driver)
            action.click_and_hold(slider).move_by_offset(300, 0).release().perform()


def send_click_captcha_request(image_src, task_image_src):
    click_response = requests.get(image_src).content
    task_response = requests.get(task_image_src).content
    click_base_64 = base64.b64encode(click_response).decode('utf-8')
    task_base_64 = base64.b64encode(task_response).decode('utf-8')

    headers = {
        'Content-type': 'application/json',
        'X-API-Key': 'clickkeyclickkeyclickkey'
    }

    create_data = {
        'type': 'SmartCaptcha',
        'click': click_base_64,
        'task': task_base_64,
    }
    create_response = requests.post(url='https://api.capsola.cloud/create', json=create_data, headers=headers)
    create_response = create_response.json()

    if create_response['status'] == 1:
        task_id = create_response['response']
        start_timer = time.time()
        while True:
            time.sleep(3)
            result_data = {'id': task_id}
            result_response = requests.post(url='https://api.capsola.cloud/result', json=result_data, headers=headers)
            result_response = result_response.json()

            if result_response['response'] == 'CAPCHA_NOT_READY':
                continue

            if result_response['status'] == 1:
                return result_response['response']

            if time.time() - start_timer > 6:
                break

    return "CAPCHA_NOT_AVAILABLE"

def send_text_captcha_request(image_src):
    captcha_image_path = 'C:/save/captcha.png'
    response = requests.get(image_src)

    if response.status_code == 200:
        with open(captcha_image_path, 'wb') as f:
            f.write(response.content)

    url = 'http://api.captcha.guru/in.php'
    key = 'textkeytextkey'
    files = {'file': open(captcha_image_path, 'rb')}
    data = {'key': key, 'method': 'post'}

    r = requests.post(url, files=files, data=data)

    if r.ok and 'OK' in r.text:
        reqid = r.text.split('|')[1]

        for timeout in range(40):
            r = requests.get(f'http://api.captcha.guru/res.php?key={key}&action=get&id={reqid}')

            if 'CAPCHA_NOT_READY' in r.text:
                time.sleep(1)

            elif 'ERROR' in r.text:
                return "CAPCHA_NOT_AVAILABLE"

            elif 'OK' in r.text:
                captcha_solution = r.text.split('|')[1]
                return captcha_solution
    return "CAPCHA_NOT_AVAILABLE"

def click_captcha_solution(driver, solution):

    if solution == "CAPCHA_NOT_AVAILABLE":
        submit_button = driver.find_element(By.CSS_SELECTOR, "button[data-testid='submit']")
        submit_button.click()
        return False

    if solution:
        if 'coordinates:' in solution:

            coordinates_str = solution.replace('coordinates:', '')
            coordinates = coordinates_str.split(';')
            actions = ActionChains(driver)
            captcha_image = driver.find_element(By.CSS_SELECTOR, "img[src*='captchaimage']")
            size = captcha_image.size
            width = size['width']
            height = size['height']

            for coord in coordinates:
                x, y = coord.split(',')
                x = float(x.split('=')[1])
                y = float(y.split('=')[1])
                actions.move_to_element_with_offset(captcha_image, x - (width / 2), y - (height / 2)).click()

            actions.perform()
            submit_button = driver.find_element(By.CSS_SELECTOR, "button[data-testid='submit']")
            submit_button.click()
            return True

        else:
            input_field = driver.find_element(By.NAME, "rep")
            input_field.send_keys(solution)

            submit_button = driver.find_element(By.CSS_SELECTOR, "button[data-testid='submit']")
            submit_button.click()
            return True

    return False

def process_click_captcha(handle, image_src, task_image_src, task_ids):
    task_id = send_click_captcha_request(image_src, task_image_src)
    task_ids[handle] = task_id

def process_text_captcha(handle, image_src, task_ids):
    task_id = send_text_captcha_request(image_src)
    task_ids[handle] = task_id

def check_and_process(driver, window_query_map):

    is_sleep = False
    for handle in window_query_map:
        driver.switch_to.window(handle)
        if 'showcaptcha' in driver.current_url:
            click(0, driver)
            is_sleep = True

    if is_sleep:
        time.sleep(4)

    kap_s = 0
    while kap_s == 0:
        task_ids = {}
        threads = []
        for handle in window_query_map:
            driver.switch_to.window(handle)

            if 'showcaptcha' in driver.current_url:
                page_html = driver.page_source

                if 'Введите текст с картинки' in page_html:

                    ssr_data = driver.execute_script("return window.__SSR_DATA__;")
                    captcha_url = ssr_data['imageSrc']

                    thread = threading.Thread(target=process_text_captcha, args=(handle, captcha_url, task_ids))
                    threads.append(thread)
                    thread.start()

                else:

                    ssr_data = driver.execute_script("return window.__SSR_DATA__;")
                    image_src = ssr_data['imageSrc']
                    task_image_src = ssr_data['taskImageSrc']

                    thread = threading.Thread(target=process_click_captcha, args=(handle, image_src, task_image_src, task_ids))
                    threads.append(thread)
                    thread.start()

        for thread in threads:
            thread.join()

        for handle in window_query_map:
            driver.switch_to.window(handle)

            if handle in task_ids:
                task_id = task_ids[handle]
                if task_id:
                    click_captcha_solution(driver, task_id)

        kap_s = 1
        for handle in window_query_map:
            driver.switch_to.window(handle)
            if 'showcaptcha' in driver.current_url:
                kap_s = 0
                break

def process_start_search(queries, headless, search_page, sites):
    global stop_threads
    if not stop_threads:
        options = Options()
        options.add_argument('--ignore-ssl-errors=yes')
        options.add_argument('--ignore-certificate-errors')
        options.headless = True

        crx1 = resource_path("ChromeMouse.crx")
        crx2 = resource_path("CPULowCaptcha.crx")

        options.add_extension(crx1)
        options.add_extension(crx2)

        if headless:
            options.add_argument('--headless=new')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument("window-size=1200,800")

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

        try:
            window_query_map = {}
            while not stop_threads:
                get_region(driver)
                time.sleep(1)
                if stop_threads:
                    break
                driver.get("https://gzmland.ru/")
                time.sleep(1)

                for query in queries:
                    if stop_threads:
                        break
                    driver.execute_script(f"window.open('https://yandex.ru/search/?text={query}&p=0');")
                    window_handle = driver.window_handles[-1]
                    window_query_map[window_handle] = query
                print("Pages: 0")
                update_log("Pages: 0")
                if stop_threads:
                    break
                check_and_process(driver, window_query_map)
                if stop_threads:
                    break
                check_positions(driver, window_query_map, 0, sites, global_results)
                if stop_threads:
                    break

                for i in range(1, search_page + 1):
                    if stop_threads:
                        break
                    print(f"Pages: {i}")
                    update_log(f"Pages: {i}")

                    new_window_query_map = {}
                    for query in queries:
                        if stop_threads:
                            break
                        driver.execute_script(f"window.open('https://yandex.ru/search/?text={query}&p={i}');")
                        new_handle = driver.window_handles[-1]
                        new_window_query_map[new_handle] = query

                    for old_handle in window_query_map.keys():
                        if stop_threads:
                            break
                        if old_handle in driver.window_handles:
                            driver.switch_to.window(old_handle)
                            driver.close()

                    window_query_map = new_window_query_map
                    if stop_threads:
                        break
                    check_and_process(driver, window_query_map)
                    if stop_threads:
                        break
                    check_positions(driver, window_query_map, i, sites, global_results)
                    if stop_threads:
                        break
                break
        except Exception as e:
            print(f"Произошла ошибка: {e}")
            update_log("Произошла ошибка в потоке. Перезапуск потока.")
            return 52

        finally:
            driver.quit()

def check_positions(driver, window_query_map, page, sites, results_queue):
    for handle in window_query_map:
        driver.switch_to.window(handle)
        dps = driver.page_source
        query = window_query_map[handle]
        for site in sites:
            ss = "<b>" + site + "</b>"
            if ss in dps:
                results = driver.find_elements(By.XPATH, '//a[@accesskey]')
                for result in results:
                    if site in result.get_attribute('href'):
                        res2 = driver.find_elements(By.XPATH, '//li[@data-cid]')
                        for res in res2:
                            link = res.find_element(By.TAG_NAME, 'a')
                            if site in link.get_attribute('href'):
                                count = int(res.get_attribute('data-cid'))
                                position = (page * 10) + count
                                result_data = {'site': site, 'query': query, 'position': position}
                                update_log(f"Найден сайт: {site}, запрос: {query}, позиция: {position}")
                                results_queue.put(result_data)
                                print(result_data)
                                break

def global_start(queries, headless, max_threads, max_queries, search_page, sites, output_path):
    query_groups = [queries[i:i + max_queries] for i in range(0, len(queries), max_queries)]
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_threads) as executor:
        futures = {}
        for group in query_groups:
            future = executor.submit(process_start_search, group, headless, search_page, sites)
            futures[future] = group

        while futures:
            for future in concurrent.futures.as_completed(futures):
                group = futures.pop(future)
                try:
                    result = future.result()
                    if result is not None:
                        new_future = executor.submit(process_start_search, group, headless, search_page, sites)
                        futures[new_future] = group
                except Exception as e:
                    print(f"Ошибка в потоке: {e}")
                    new_future = executor.submit(process_start_search, group, headless, search_page, sites)
                    futures[new_future] = group

    process_results(sites, queries, output_path)

def process_results(sites, queries, output_path):
    result_data = {'site': "empty site", 'query': "empty query", 'position': 777}
    global_results.put(result_data)

    results = []
    while not global_results.empty():
        result = global_results.get()
        if 'position' in result:
            results.append(result)
        else:
            print("Ошибка: в результате отсутствует 'position':", result)
    df = pd.DataFrame(results)
    print(df)
    output_file = os.path.join(output_path, f'sorted_data_{datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.xlsx')
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for site in sites:
            data_for_site = []
            for query in queries:
                filtered_df = df[(df['site'] == site) & (df['query'] == query)]
                if not filtered_df.empty:
                    data_for_site.extend(filtered_df.to_dict('records'))
                else:
                    data_for_site.append({'site': site, 'query': query, 'position': 'Сайт не найден'})
            df_site = pd.DataFrame(data_for_site)
            df_site['position'] = pd.to_numeric(df_site['position'], errors='coerce')
            df_site.sort_values(by='position', inplace=True)
            df_site['position'] = df_site['position'].fillna('Сайт не найден')
            df_site.to_excel(writer, sheet_name=site, index=False)
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for i, col in enumerate(df_site.columns):
                    column_len = df_site[col].astype(str).map(len).max()
                    column_len = max(column_len, len(str(col))) + 2  # Добавляем немного места
                    worksheet.set_column(i, i, column_len)
        print("Данные успешно отсортированы и сохранены в", output_file)
        update_log(f"Данные успешно отсортированы и сохранены в {output_file}")

        os.startfile(output_file)

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def check_and_copy_settings():
    external_settings_path = "C:\\MMCheckPositions\\settings.json"
    if not os.path.exists(external_settings_path):
        if not os.path.exists(os.path.dirname(external_settings_path)):
            os.makedirs(os.path.dirname(external_settings_path))
        internal_settings_path = resource_path("settings.json")
        shutil.copy(internal_settings_path, external_settings_path)

def save_settings():
    settings = {
        "headless": headless_var.get(),
        "max_threads": max_threads_entry.get(),
        "max_queries": max_queries_entry.get(),
        "search_page": search_page_entry.get(),
        "output_dir": output_dir_entry.get(),
        "queries": queries_text.get('1.0', tk.END).strip(),
        "sites": sites_text.get('1.0', tk.END).strip()
    }
    settings_path = "C:\\MMCheckPositions\\settings.json"
    with open(settings_path, "w") as file:
        json.dump(settings, file)

def load_settings():
    settings_path = "C:\\MMCheckPositions\\settings.json"
    check_and_copy_settings()
    try:
        with open(settings_path, "r") as file:
            settings = json.load(file)
            headless_var.set(settings["headless"])
            max_threads_entry.delete(0, tk.END)
            max_threads_entry.insert(0, settings["max_threads"])
            max_queries_entry.delete(0, tk.END)
            max_queries_entry.insert(0, settings["max_queries"])
            search_page_entry.delete(0, tk.END)
            search_page_entry.insert(0, settings["search_page"])
            output_dir_entry.delete(0, tk.END)
            output_dir_entry.insert(0, settings["output_dir"])
            queries_text.delete('1.0', tk.END)
            queries_text.insert(tk.END, settings["queries"])
            sites_text.delete('1.0', tk.END)
            sites_text.insert(tk.END, settings["sites"])
    except FileNotFoundError:
        pass

def select_output_directory():
    directory = filedialog.askdirectory()
    if directory:
        output_dir_entry.delete(0, tk.END)
        output_dir_entry.insert(0, directory)


def validate_input():
    try:
        max_threads = int(max_threads_entry.get())
        if max_threads < 1:
            raise ValueError("Максимальное количество потоков должно быть не меньше 1.")
    except ValueError as e:
        messagebox.showerror("Ошибка ввода", str(e))
        return False

    try:
        max_queries = int(max_queries_entry.get())
        if max_queries < 1:
            raise ValueError("Максимальное количество запросов в потоке должно быть не меньше 1.")
    except ValueError as e:
        messagebox.showerror("Ошибка ввода", str(e))
        return False

    try:
        search_page = int(search_page_entry.get())
        if not (1 <= search_page <= 24):
            raise ValueError("Страница поиска должна быть в диапазоне от 1 до 24.")
    except ValueError as e:
        messagebox.showerror("Ошибка ввода", str(e))
        return False

    output_path = output_dir_entry.get()
    if not os.path.exists(output_path):
        messagebox.showerror("Ошибка ввода", "Указанный путь для сохранения результатов не существует.")
        return False

    return True

def on_run_clicked():
    if validate_input():
        stop_button.config(state="normal")

        global stop_threads
        stop_threads = False

        log_text.config(state=tk.NORMAL)
        log_text.delete('1.0', tk.END)
        log_text.config(state=tk.DISABLED)
        run_button.config(state="disabled")
        threading.Thread(target=run_search).start()

def run_search():
    output_path = output_dir_entry.get()
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    headless = headless_var.get() == 1
    max_threads = int(max_threads_entry.get())
    max_queries = int(max_queries_entry.get())
    search_page = int(search_page_entry.get())

    queries = queries_text.get('1.0', tk.END).split('\n')
    queries = [query for query in queries if query.strip()]

    sites = sites_text.get('1.0', tk.END).split('\n')
    sites = [site for site in sites if site.strip()]

    date_n = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    update_log(date_n)
    update_log("Поиск начался.")

    update_log(f"Количество запросов: {len(queries)}")
    update_log(f"Количество сайтов: {len(sites)}")

    global_start(queries, headless, max_threads, max_queries, search_page, sites, output_path)

    update_log("Поиск завершен.")
    date_n = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    update_log(date_n)

    run_button.config(state="normal")
    stop_button.config(state="normal")

def on_closing():
    save_settings()
    window.destroy()

def on_stop_clicked():
    output_path = output_dir_entry.get()
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    queries = queries_text.get('1.0', tk.END).split('\n')
    queries = [query for query in queries if query.strip()]

    sites = sites_text.get('1.0', tk.END).split('\n')
    sites = [site for site in sites if site.strip()]

    process_results(sites, queries, output_path)

    stop_button.config(state="disabled")

    global stop_threads
    stop_threads = True

    update_log("Идет остановка процессов, ожидайте.")

font_large = ('Arial', 12, 'bold')

window = tk.Tk()
window.title("Поисковый интерфейс")
window.protocol("WM_DELETE_WINDOW", on_closing)

headless_var = tk.IntVar()
tk.Checkbutton(window, text="Скрывать окна браузера", variable=headless_var, font=font_large).pack()

frame = Frame(window)
frame.pack(padx=15, pady=15)

Label(frame, text="Максимальное количество потоков:", font=font_large).pack()
max_threads_entry = Spinbox(frame, from_=1, to=100, font=font_large, width=5)
max_threads_entry.pack()

Label(frame, text="Максимальное количество запросов в потоке:", font=font_large).pack()
max_queries_entry = Spinbox(frame, from_=1, to=100, font=font_large, width=5)
max_queries_entry.pack()

Label(frame, text="До какой страницы идет поиск:", font=font_large).pack()
search_page_entry = Spinbox(frame, from_=1, to=24, font=font_large, width=5)
search_page_entry.pack()

Label(frame, text="Путь для сохранения результатов:", font=font_large).pack()
output_dir_entry = tk.Entry(frame, font=font_large, width=50)
output_dir_entry.pack(side=tk.LEFT)

folder_icon_path = resource_path("folder_icon.png")
folder_icon = PhotoImage(file=folder_icon_path)
folder_icon = folder_icon.subsample(15, 15)
select_dir_button = Button(frame, image=folder_icon, command=select_output_directory)
select_dir_button.pack(side=tk.RIGHT)

tk.Label(window, text="Запросы:").pack()
queries_text = scrolledtext.ScrolledText(window, height=10)
queries_text.pack()

tk.Label(window, text="Сайты:").pack()
sites_text = scrolledtext.ScrolledText(window, height=10)
sites_text.pack()

run_button = tk.Button(window, text="Запустить", command=on_run_clicked)
run_button.pack()

stop_button = tk.Button(window, text="Остановить", command=on_stop_clicked, state="disabled")
stop_button.pack()

log_text = scrolledtext.ScrolledText(window, height=10)
log_text.config(state=tk.DISABLED)
log_text.pack()

load_settings()

window.mainloop()
