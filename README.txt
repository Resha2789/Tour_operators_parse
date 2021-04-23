
Python 3.8.8
64 bit

Подключаемые библиотеки:
altgraph	0.17	0.17
certifi	2020.12.5	2020.12.5
chardet	4.0.0	4.0.0
colorama	0.4.4	0.4.4
configparser	5.0.2	5.0.2
crayons	0.4.0	0.4.0
future	0.18.2	0.18.2
idna	2.10	3.1
keyboard	0.13.5	0.13.5
pefile	2019.4.18	2019.4.18
pip	21.0.1	21.0.1
pyinstaller	4.2	4.3
pyinstaller-hooks-contrib	2021.1	2021.1
pywin32	300	300
pywin32-ctypes	0.2.0	0.2.0
requests	2.25.1	2.25.1
selenium	3.141.0	3.141.0
setuptools	56.0.0	56.0.0
soupsieve	2.2.1	2.2.1
urllib3	1.26.4	1.26.4
webdriver-manager	3.3.0	3.4.0

Импортируемые библиотеки:
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import time
import win32com.client
import os
import keyboard

Команда для компиляции файла в exe:
pyinstaller --onefile --distpath D:\Фриланс\Авито\Дария\ --hidden-import win32timezone -n Parsing Main.py