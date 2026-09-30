# Документооборот

## API для внешней CRM

Приложение предоставляет API для чтения и учёта задач, планов и служебных записок во внешней CRM.

### Подготовка

1. Войдите под администратором или суперадминистратором.
2. Откройте `Администрирование -> Настройки`.
3. В блоке **Интеграция с внешней CRM** нажмите `Создать токен`.
4. Передайте токен администратору CRM защищённым способом.

При перевыпуске токена старый ключ сразу перестаёт работать.

### Авторизация

Каждый запрос должен содержать один из заголовков:

```http
Authorization: Bearer YOUR_API_TOKEN
```

или:

```http
X-API-Key: YOUR_API_TOKEN
```

При неверном или отсутствующем токене API возвращает `401`:

```json
{"error":"unauthorized","message":"Передайте действующий API-токен."}
```

### Адреса

Замените `http://server:5001` на адрес вашего сервера.

| Метод | Адрес | Назначение |
| --- | --- | --- |
| `GET` | `/api/crm/v1/health` | Проверка подключения |
| `GET` | `/api/crm/v1/tasks` | Задачи |
| `POST` | `/api/crm/v1/tasks` | Создать задачу |
| `PATCH` | `/api/crm/v1/tasks/<id>` | Изменить задачу |
| `GET` | `/api/crm/v1/plans` | Планы и строки планов |
| `POST` | `/api/crm/v1/plans` | Создать план с пунктами |
| `PATCH` | `/api/crm/v1/plans/<id>` | Изменить план |
| `POST` | `/api/crm/v1/plans/<id>/items` | Добавить пункт плана |
| `PATCH` | `/api/crm/v1/plans/<id>/items/<item_id>` | Изменить пункт плана |
| `GET` | `/api/crm/v1/memos` | Служебные записки |

Все списочные ресурсы возвращают объект:

```json
{
  "data": [],
  "pagination": {"total": 0, "limit": 100, "offset": 0}
}
```

Общие параметры пагинации:

| Параметр | Значение |
| --- | --- |
| `limit` | Число записей, от `1` до `500`, по умолчанию `100` |
| `offset` | Смещение, по умолчанию `0` |

Общие параметры периода:

| Параметр | Формат |
| --- | --- |
| `date_from` | `YYYY-MM-DD` |
| `date_to` | `YYYY-MM-DD` |

### Задачи

```bash
curl -H "Authorization: Bearer YOUR_API_TOKEN" \
  "http://server:5001/api/crm/v1/tasks?date_from=2026-09-01&date_to=2026-09-30&limit=100"
```

Дополнительный параметр: `status` (`in_progress`, `done`, `not_done`, `postponed`, `cancelled`).

Запись задачи содержит идентификатор, название, описание, тип, статус, даты, дату переноса, комментарий статуса, тему, исходный документ, исполнителя, отдел, связанный план и реквизиты исходного/закрывающего документа.

```json
{
  "id": 53,
  "title": "Подготовить отчет",
  "status": "in_progress",
  "status_label": "В работе",
  "start_date": "2026-09-16T00:00:00",
  "end_date": null,
  "owner": {"id": 1, "full_name": "Иванов И.И.", "department": "ОТД"}
}
```

### Запись задач

Все операции записи принимают только JSON (`Content-Type: application/json`). Создание задачи требует
`title`, `type_id`, `owner_id` и `start_date`. Исполнитель должен быть активным и подтверждённым.
Даты задач передаются в ISO 8601: `2026-09-30` или `2026-09-30T09:00:00`.

```bash
curl -X POST "http://server:5001/api/crm/v1/tasks" \
  -H "Authorization: Bearer YOUR_API_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{
    "title": "Подготовить данные для CRM",
    "description": "Передать сведения по плану",
    "type_id": 1,
    "owner_id": 12,
    "start_date": "2026-10-01",
    "end_date": "2026-10-05",
    "priority": 4,
    "complexity_level": 3,
    "status": "in_progress"
  }'
```

Для изменения передайте только нужные поля методом `PATCH`, например:

```bash
curl -X PATCH "http://server:5001/api/crm/v1/tasks/53" \
  -H "X-API-Key: YOUR_API_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"status":"postponed","postponed_to":"2026-10-10","status_comment":"Ожидаются входные данные"}'
```

Каждое изменение задачи через API заносится в журнал изменений с пометкой внешней CRM.

### Планы

```bash
curl -H "X-API-Key: YOUR_API_TOKEN" \
  "http://server:5001/api/crm/v1/plans?date_from=2026-09-01&date_to=2026-09-30"
```

Каждый план содержит период, текст, статус утверждения, автора и массив `items`. Строка плана содержит исполнителя, срок (`deadline_kind`, `deadline_date`), статус выполнения, комментарий и дату выполнения.

### Запись планов

При создании плана обязательны `text`, `start_date`, `end_date` и непустой массив `items`.
У каждого пункта обязательны `text` и `executor_id`; для `deadline_kind: "date"` также требуется
`deadline_date`. Допустимые сроки: `month`, `q1`, `q2`, `q3`, `q4`, `date`.

```bash
curl -X POST "http://server:5001/api/crm/v1/plans" \
  -H "Authorization: Bearer YOUR_API_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{
    "text": "План отдела на октябрь",
    "start_date": "2026-10-01",
    "end_date": "2026-10-31",
    "created_by_id": 12,
    "items": [{
      "text": "Подготовить ежемесячный отчет",
      "executor_id": 12,
      "deadline_kind": "date",
      "deadline_date": "2026-10-25"
    }]
  }'
```

API не удаляет задачи и планы. Удаление остаётся доступным только через веб-интерфейс администратора.

### Служебные записки

```bash
curl -H "Authorization: Bearer YOUR_API_TOKEN" \
  "http://server:5001/api/crm/v1/memos?approval_status=approved"
```

Дополнительный параметр: `approval_status` (`pending` или `approved`). Запись содержит номер, дату, тему, текст, автора, исполнителя, отделы и статус утверждения.

### Проверка подключения

```bash
curl -H "Authorization: Bearer YOUR_API_TOKEN" \
  "http://server:5001/api/crm/v1/health"
```

Успешный ответ:

```json
{"status":"ok","api_version":"v1","server_time":"2026-09-17T12:00:00Z"}
```
# Распознавание документов Kreuzberg

Для извлечения текста из PDF, изображений и офисных документов приложение использует
локальный HTTP-сервис Kreuzberg. При загрузке файла отправляется запрос `POST /extract`
с русским OCR и параметром `force_ocr=true`.

Из PowerShell, находясь в папке проекта, запустите контейнер:

```powershell
docker rm -f kreuzberg_service
docker run -d --name kreuzberg_service -p 8000:8000 ghcr.io/kreuzberg-dev/kreuzberg-full:latest serve -H 0.0.0.0 -p 8000
```

Проверка готовности:

```powershell
Invoke-WebRequest http://127.0.0.1:8000/health
```

По умолчанию приложение обращается к `http://127.0.0.1:8000`. Для другого адреса
задайте переменную окружения `KREUZBERG_URL`, например `http://192.168.0.187:8000`.
Если сервис временно недоступен, загрузка документа продолжится с резервным локальным
распознаванием.
