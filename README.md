# Конвертер PPTX

Сервис разбирает презентации PPTX и собирает из них `out.zip` для редактора Presentonika. Текст и картинки остаются редактируемыми, а сложные элементы сохраняются в фоне слайда.

## Что нужно для работы

- Node.js 18 или новее;
- LibreOffice;
- `pdftoppm` из пакета Poppler.

На Ubuntu зависимости можно поставить так:

```bash
sudo apt-get install libreoffice poppler-utils
```

## Установка и сборка

```bash
npm ci
npm run build
```

## Запуск из консоли

```bash
node dist/importer.js --input ./presentation.pptx --out ./out
```

В папке результата появятся `doc.json`, фоны слайдов и извлечённые изображения.

## HTTP-сервис

```bash
npm run start:server
```

Пример запроса:

```bash
curl -X POST http://localhost:3001/convert \
  -H "Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation" \
  --data-binary @./presentation.pptx \
  --output out.zip
```

Сервис ограничивает размер входного файла, количество файлов внутри PPTX и время работы внешних программ. Параллельные запросы проходят через небольшую очередь.

## Docker

```bash
docker compose up -d --build
```

Внутри общей Docker-сети редактор обращается к конвертеру по адресу `http://converter:3001`.

## Проверки

```bash
npm test
```

Тесты собирают небольшие презентации прямо во время запуска, поэтому хранить отдельные бинарные примеры в репозитории не нужно.
