1) Установить библиотеки:
	pip3 install pandas
	pip3 install openpyxl
	pip3 install google-api-python-client
	pip3 install oauth2client
2) Создать проeкт, подключить API и скачать json файл для работы с Google таблицей
3) Добавить доступ к гугл таблице, вписав данные из поля "client_email" файла json в настройках доступа к таблице
4) В файле goog_exl.py вписать свои данные:
	1. CREDENTIALS_FILE = "путь к json файлу на компьютере"
	2. spread_id = "id гугл документа, с которым нужно работать"

5) Запустить скрипт из файла script_xl.py