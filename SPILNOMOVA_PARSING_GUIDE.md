# SPILNOMOVA: Підсумкова структура даних

## Що в підсумку

Ми залишили **простий формат**:
- один основний файл з даними: `dataset.json`
- один технічний файл перевірки: `report.json`

CSV/JSONL більше не генеруються, щоб не дублювати дані і не ускладнювати структуру.

## Де що лежить

### Вхідні дані

```text
<PROJECT_ROOT>/
  01_INPUT FILES/
    SPILNOMOVA_DATABASE SAMPLE_TEST_11122025.xlsx
    01_AUDIO/
      01_Zhytomyr/
      02_Chernigiv/
      ...
    02_TRANSCRIPTS/
      01_Zhytomyr/
      02_Chernigiv/
      ...
```

### Скрипт

```text
<PROJECT_ROOT>/scripts/parse_spilnomova_dataset.py
```

### Вихідні дані

```text
<PROJECT_ROOT>/02_OUTPUT/SPILNOMOVA_PARSED/
  dataset.json
  report.json
```

## Що містить `dataset.json`

`dataset.json` має просту ієрархію:
- `meta`: службова інформація про запуск
- `sessions`: масив сесій (1 сесія = 1 дитина/1 аудіо)

Кожна сесія містить:
- метадані (`city`, `school`, `group`, `child_name`)
- `audio` (`filename`, `relpath`)
- `transcript` (`filename`, `sheet`, `relpath`)
- `stats`:
  - `turns_total`
  - `turns_child`
  - `word_count_annotated_total`
  - `word_count_annotated_child_total`
- `dialogue`: масив реплік у правильному порядку

Кожна репліка в `dialogue`:
- `speaker` (оригінальне значення з таблиці, напр. `Дитина` / `Модератор`)
- `role` (`child` або `moderator`)
- `text` (поле `Стенограма`)
- `word_count` (поле `Кількість слів`, якщо заповнене)

## Що містить `report.json`

`report.json` потрібен для контролю якості:
- які шляхи використані
- скільки записів прочитано
- скільки сесій/реплік зібрано
- попередження (`warnings`), якщо щось не знайдено

## Важливо про переносимість

Формат зроблено універсальним:
- у даних зберігаються **відносні шляхи** (`relpath`), а не шляхи конкретного комп’ютера;
- проєкт можна передати іншій людині (ZIP), і все працюватиме, якщо структура папок збережена;
- скрипт можна запускати з будь-якої директорії.

## Запуск

Базовий запуск:

```bash
python3 /path/to/<PROJECT_ROOT>/scripts/parse_spilnomova_dataset.py
```

Або з явними параметрами:

```bash
python3 scripts/parse_spilnomova_dataset.py \
  --project-root /path/to/<PROJECT_ROOT> \
  --db-xlsx "01_INPUT FILES/SPILNOMOVA_DATABASE SAMPLE_TEST_11122025.xlsx" \
  --audio-root "01_INPUT FILES/01_AUDIO" \
  --transcripts-root "01_INPUT FILES/02_TRANSCRIPTS" \
  --output-dir "02_OUTPUT/SPILNOMOVA_PARSED"
```

## Які колонки реально використовуються з транскриптів

Основні:
- `Хто говорить`
- `Стенограма`
- `Кількість слів`

Інші колонки з оригінальних таблиць не втрачаються у джерелі, але в навчальний датасет не включаються, щоб тримати структуру простою.
