import time
# Импортируем библиотеку для excel
import xlrd
# Импортируем библиотеку веб драйвера
from selenium import webdriver
# Импортируем библиотеку с кодами клавиш
from selenium.webdriver.common.keys import Keys
# Импортируем библиотеку с цепочками событий
from selenium.webdriver import ActionChains

# Открываем xls файл
workbook = xlrd.open_workbook('messages.xls')
# Выбираем активный лист
sheet = workbook.sheet_by_index(0)

# Создаем переменную драйвера
driver = webdriver.Chrome('C:/Users/richuser/ChromeDriver/chromedriver.exe')  # <- путь к драйверу

# Открываем на весь экран
driver.maximize_window()
# Задаём максимальное время ожидания
driver.implicitly_wait(10)
# Заходим на сайт гугла
driver.get('http://www.google.com/xhtml')

##### Главная страница
# Ищем элемент кнопку "Вход"
button = driver.find_element_by_id('gb_70')
# Кликаем "Вход"
button.click()
# Попадаем в окно ввода почты
# Находим поле ввода почты
input = driver.find_element_by_id('identifierId')
# Вводим почту
input.send_keys('xxx@gmail.com')
# Ищем элемент кнопку "Далее"
button = driver.find_element_by_id('identifierNext')
# Кликаем "Далее"
button.click()

# Попадаем в окно ввода пароля
# Находим поле ввода пароля
input = driver.find_element_by_name('password')
# Вводим почту
input.send_keys('xxx') # <- пароль от почты
# Ищем элемент кнопку "Далее"
button = driver.find_element_by_id('passwordNext')
# Кликаем "Далее"
button.click()

# Попадаем в главное окно гугла
# Ищем элемент кнопку "Почта"
button = driver.find_element_by_class_name('gb_d')
# Кликаем "Почта"
button.click()

# Попадаем в почту
for i in range(1, sheet.nrows):
    time.sleep(5)
    # Ищем кнопку "Написать"
    button = driver.find_element_by_xpath('//div[@gh="cm"]')
    # Кликаем "Написать"
    button.click()

    ##### Открывается окно с письмом
    # Находим поле ввода почты
    inputReceiver = driver.find_element_by_name('to')
    inputReceiver.send_keys(sheet.row_values(i)[0])

    # Находим поле ввода названия письма
    inputReceiver = driver.find_element_by_name('subjectbox')
    inputReceiver.send_keys(sheet.row_values(i)[1])

    # Находим поле ввода текста письма
    inputReceiver = driver.find_element_by_xpath('//*[@id="undefined"]/tbody/tr/td[2]/div[2]/div')
    driver.execute_script("arguments[0].innerHTML='" + sheet.row_values(i)[2] + "'", inputReceiver)

    # Зажимаем Control нажимаем Enter(Return) и отжимаем Control, это приводит к отправке сообщения
    ActionChains(driver).key_down(Keys.CONTROL).send_keys(Keys.RETURN).key_up(Keys.CONTROL).perform()

