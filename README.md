# planificator

This repository is migrated to a Symfony + daisyUI web app.

## App location

- Symfony application: `symfony/`

## Run locally

```bash
cd symfony
composer install
symfony server:start
```

If Symfony CLI is not installed, use:

```bash
cd symfony
php -S 127.0.0.1:8000 -t public
```

## Functionality

- Schedule generation (`/generate_schedule`)
- Schedule Excel export (`/export_schedule`)
- XML formatter (`/format-xml`)
- Word schedule matching (`/convert_word`)

All core workflows are now implemented in Symfony/PHP services.
