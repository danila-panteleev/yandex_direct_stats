# Отправка отчетов Яндекс Директ
### Установка:
1. Создайте папку yandex_direct_stats
2. Внутри папки ```git clone https://github.com/danila-panteleev/yandex_direct_stats.git```
3. ```pip install -r requirements.txt``` 
4. Введите нужные настройки в [config_example.yml](config_example.yml)
5. Переименуйте config_example.yml в config.yml
6. Запустите yandex_direct_stats.py

В итоге на ящик, указанный в параметре receiver файла config.yml должно придти письмо с данными в теле письма.

### Предупреждение
Отправка писем работает только из-под аккаунта Gmail,
в настройках безопасности которого [разрешен доступ ненадежных приложений](https://support.google.com/accounts/answer/6010255?hl=ru)

