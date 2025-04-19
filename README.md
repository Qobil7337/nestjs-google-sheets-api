### 📊 Google Sheets Sync Service (NestJS)
#### Этот сервис на NestJS получает структурированные данные из внутреннего API и записывает их в указанные диапазоны Google Sheets, группируя по имени листа.

### 🚀 Начало работы

#### 1. Убедитесь, что у вас установлена версия Node.js 22

Для этого выполните команду:

```bash
node -v
```
#### 2.Клонируйте репозиторий
```link
git clone https://github.com/Qobil7337/nestjs-google-sheets-api
cd nestjs-google-sheets-api
```

#### 3.Установите зависимости
```bash
npm install
```
#### 4.Создайте файл конфигурации .env
Для корректной работы приложения создайте файл .env в корневой директории проекта и добавьте в него необходимые параметры:
```bash
INTERNAL_API_URL=https://script.google.com/macros/s/AKfycbxBLv4QkfvO18qyD52fmkt34tJ29YTp1aMUifIwEHVzmiYZciEazIVKI1Q0VkA5Jiu9/exec?apikey=640fb8c3-cc56-4ade-a447-8313d65657ee&action=get_data
SPREADSHEET_ID=1lo43rcTStWWO4ZoanL_1mqDJEp9fZjlGrW4rfoFQQHY
```
#### 5.Запуск приложения
После того, как все зависимости установлены, а конфигурация настроена, можно запустить приложение:
```bash
npm start
```
Это запустит приложение на порту 3000 в режиме разработки.

#### 6.Синхронизация данных с Google Sheets
Когда приложение запустится, скопируйте следующую ссылку:
```bash
http://localhost:3000/google-sheets/write
```
Откройте эту ссылку в вашем браузере. После этого приложение начнёт процесс записи данных в Google Sheets
