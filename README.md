# Newsletter-агент: еженедельные акции

Excel с акциями → Claude пишет тексты на немецком → черновик приходит вам на почту → вы согласуете → в понедельник в 10:00 (Берлин) письмо уходит подписчикам через Brevo.

Цены, скидки и сроки берутся **только из Excel** и подставляются шаблоном. Claude пишет лишь тему, вступление и короткие описания. Код проверяет, что в этих текстах нет цен и процентов, а если модель ошиблась, просит её исправиться. После второй неудачи черновик получает статус `needs_review`.

## Структура

| Файл | Назначение |
|---|---|
| `agent.mjs` | Агент: команды `draft`, `approve`, `send`, `status` |
| `prompts/system_prompt.md` | Системный промпт для Claude |
| `templates/newsletter.html` | HTML-шаблон письма (Handlebars) |
| `config.json` | Бренд, Impressum, расписание, пути к файлам |
| `data/aktionen.xlsx` | Акции (пример: `npm run sample-data`) |
| `data/subscribers.csv` | Подписчики |
| `data/abgemeldet.csv` | Отписавшиеся: им письма не отправляются |
| `state/<дата>/` | Черновики, превью и журнал отправки |

## Настройка (один раз)

1. **Зависимости** (уже установлены): `npm install`
2. **Brevo**: зарегистрируйтесь на brevo.com, подтвердите домен `onlinetabakshop.de` (SPF и DKIM), возьмите SMTP-логин и ключ в *Settings → SMTP & API*. Перед стартом проверьте в правилах Brevo (Acceptable Use Policy), что табачный контент разрешён, или уточните это в поддержке.
3. **Claude API**: получите ключ на console.anthropic.com.
4. Скопируйте `.env.example` в `.env` и заполните.
5. В `config.json` укажите реальные данные Impressum, логотип (`logo_url`), фирменный цвет и ссылку на политику конфиденциальности.

## Подписчики из Gmail

Google Контакты → выберите контакты → *Экспорт* → **Google CSV** → сохраните как `data/subscribers.csv`.
- Если подписчики отмечены ярлыком (например, «Newsletter»), укажите его в `config.json → subscribers.require_label`. Тогда письма получат только контакты с этим ярлыком.
- Контакты с ярлыком `B2B` получат B2B-версию письма, остальные — B2C.

Можно вести и простой CSV: `email,vorname,nachname,segment` — образец в `data/subscribers.example.csv`.
Сам `data/subscribers.csv` в git не попадает, как и `.env`: после скачивания создайте их заново.

## Установка на новом компьютере

Нужен Node.js 20+ (nodejs.org).

```bash
npm install
```

Затем скопируйте `.env.example` в `.env`, заполните ключи и положите `data/subscribers.csv`.

## Формат Excel

Первая строка — заголовки (немецкие или русские):

`Produkt` · `Marke` · `Alter Preis` · `Neuer Preis` · `Gültig von` · `Gültig bis` · `Shop` · `Bedingungen` · `Link` · `Bild`

Обязательные колонки: Produkt, Alter Preis, Neuer Preis, Gültig bis. В письмо попадают акции, которые действуют на неделе отправки. Строки с ошибками (новая цена не ниже старой, нет даты) пропускаются и показываются в черновике как предупреждения.

> «Alter Preis» по PAngV §11 должен быть **самой низкой ценой за последние 30 дней**: в письме он так и подписан.

## Еженедельный цикл

```bash
node agent.mjs draft
```
Обычно в четверг. Агент создаёт черновики для каждого сегмента в `state/<дата>/` и присылает превью на `ADMIN_EMAIL`.

```bash
node agent.mjs approve
```
После проверки превью. Без этой команды ничего не отправится.

```bash
node agent.mjs send
```
Отправляет согласованную рассылку, если в Берлине уже понедельник, 10:00 или позже. Повторный запуск безопасен: уже отправленные адреса пропускаются.

```bash
node agent.mjs status
```

Полезные флаги: `draft --no-llm` (шаблонные тексты без API), `draft --no-mail` (без превью на почту), `draft --date 2026-11-09`, `send --dry-run` (только показать, кому уйдёт), `send --now` (отправить сразу).

## Автозапуск (Планировщик задач Windows)

```bash
schtasks /Create /TN "Newsletter Draft" /SC WEEKLY /D THU /ST 12:00 /TR "cmd /c cd /d \"C:\Users\Microsoft\Desktop\Claude Code\newsletter-agent\" && node agent.mjs draft >> state\log.txt 2>&1"
```

```bash
schtasks /Create /TN "Newsletter Send" /SC WEEKLY /D MON /ST 08:00 /RI 60 /DU 08:00 /TR "cmd /c cd /d \"C:\Users\Microsoft\Desktop\Claude Code\newsletter-agent\" && node agent.mjs send >> state\log.txt 2>&1"
```

Задача отправки запускается по понедельникам каждый час с 08:00 и сама ждёт 10:00 по Берлину, поэтому часовой пояс компьютера не важен. Компьютер должен быть включён. Для работы 24/7 этот же код можно перенести на сервер или VPS с cron.

## Метрика: открытия

Открытия отслеживает Brevo: *Transactional → Statistics / Logs*, фильтр по тегу `newsletter-<дата>`. Учтите, что Apple Mail Privacy Protection завышает число открытий, поэтому смотрите на динамику от недели к неделе, а не на абсолютное значение.

## Отписка

В каждом письме есть ссылка и заголовок `List-Unsubscribe`. Отписка открывает письмо на `UNSUBSCRIBE_EMAIL`. Добавляйте такие адреса в `data/abgemeldet.csv`: агент их пропускает. Отписку нужно обработать сразу, это требование UWG §7 и DSGVO.
