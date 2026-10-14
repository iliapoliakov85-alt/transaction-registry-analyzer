# Анализатор реестра транзакций

Локальное приложение для анализа реестров транзакций АСОП, аналитики парка, контроля ФРОД и сверки с отчётами ВТБ.

## Запуск

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python analyzer.py
```

## Обновления через GitHub

Программа проверяет файл `update_manifest.json` по адресу, указанному в локальном `update_config.json`.

1. Скопируйте `update_config.json.example` в `update_config.json`.
2. Укажите в нём прямую GitHub Raw-ссылку на `update_manifest.json`.
3. При новом релизе измените `version`, `notes`, `script_url` и, при необходимости, `script_sha256`.

Локальные БД, исходные реестры, архивы и резервные копии исключены из репозитория.
