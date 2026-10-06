# GDMexico_SepomexImport 1.1.0

Compatibilidad objetivo: Magento Open Source / Adobe Commerce 2.4.2 en adelante y PHP 7.4 en adelante (sujeto a los requisitos PHP de cada versión de Magento).

## Flujo
1. Admin > System > SEPOMEX > Importar SEPOMEX.
2. Acepta XLS/XLSX.
3. Convierte todas las hojas válidas a `var/sepomex_import/sepomex_import.csv` delimitado por `|`.
4. Carga `sepomex_import` y construye `sepomex_new`.
5. Si existen `store_url` y `exclusion_url`, preserva routing: CP+asentamiento, CP único y estado normalizado+municipio para CP nuevos.
6. Valida que no queden filas sin routing y que coincidencias exactas no cambien routing.
7. Sólo después elimina el backup anterior y ejecuta swap atómico: `sepomex -> sepomex_backup`, `sepomex_new -> sepomex`.
8. No modifica `sepomex_redirects` ni `sepomex_missing`.

## Dependencia Excel
El módulo reutiliza `PhpOffice\\PhpSpreadsheet` si está disponible en el proyecto. Antes de usar:
`php -r "require 'vendor/autoload.php'; echo class_exists('PhpOffice\\\\PhpSpreadsheet\\\\IOFactory') ? 'OK' : 'NO';"`

No se fija una versión de PhpSpreadsheet en composer.json para no forzar conflictos entre Magento/PHP antiguos y nuevos. Si un proyecto no la incluye, debe instalarse una versión compatible con su PHP/Magento.

## Instalación
Copiar el contenido en `app/code/GDMexico/SepomexImport` y ejecutar:
`php bin/magento module:enable GDMexico_SepomexImport`
`php bin/magento setup:upgrade`
`php bin/magento cache:clean`

## Nota de seguridad
Probar primero en TEST. El módulo detecta si la tabla `sepomex` tiene las columnas de routing; si no existen, importa únicamente las 10 columnas base.
