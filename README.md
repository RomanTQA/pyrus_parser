Демка парсера для pyrus.com.
Запускается через консоль,

```
node main.js
```

скрипт использует файл requests.js как библиотеку методов.
Используемые модули:
fs
path
axios
exceljs
нужно установить их, через

```
npm install <модуль>
```

скрипт ищет в папке, где он находится,xlsx файл с тасками от импорта pyrus. примерный вид такого файла "TaskList_20240226093812.xlsx".
чтобы его сформировать, нужно в pyrus экспортировать список задач, см скриншот: https://skr.sh/sOT68HktMeg

чтобы парсить, нужны будут данные: логин (email) в pyrus, а также security_key: их можно найти в настройках профиля. https://pyrus.com/t#authorize, https://skr.sh/sOTe42ZVVHC

при первом запуске, если нет файла настроек pyrus_auth.json в папке со скриптом, он будет создан.
нужно заполнить файл настроек pyrus_auth.json валидными данными - login, security_key.

при повторном запуске скрипт проверит данные, и попытается получить access_token. Токен имеет короткий срок жизни - около 24 ч. Возможно , снова потребуется перезапуск.

когда скрипт будет готов парсить, он уведомит юзера об этом, отдать разрешение на парсинг придется вручную ("y").

скрипт "прозвонит" все id тасков, которые он достал из файла xlsx, возьмет первый коммент (он же - тело таска), создаст копию xlsx файла с приставкой "edited" в конце,
и запишет спарсенные комменты в свободный столбец "I" или следующий после него, если столбец свободен.

в настройках pyrus_auth.json доступны дебаг-методы для подключения, а также бета- когда парсятся все комменты "в строчку" в excel файл.

в конце скрипта будет распечатан блок статистики, для того, чтобы отслеживать, как прошел парсинг.
