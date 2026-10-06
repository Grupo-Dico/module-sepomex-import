<?php

namespace GDMexico\SepomexImport\Model;

use Magento\Framework\App\ResourceConnection;

class Importer
{
    const DB_BATCH_SIZE = 5000;

    private $resource;
    private $converter;

    private $baseColumns = array(
        'id_estado',
        'estado',
        'id_municipio',
        'municipio',
        'ciudad',
        'zona',
        'cp',
        'asentamiento',
        'tipo',
        'id'
    );

    public function __construct(
        ResourceConnection $resource,
        ExcelConverter $converter
    ) {
        $this->resource = $resource;
        $this->converter = $converter;
    }

    public function execute(array $file)
    {
        $ext = strtolower(
            pathinfo(
                isset($file['name']) ? (string) $file['name'] : '',
                PATHINFO_EXTENSION
            )
        );

        if (!in_array($ext, array('xls', 'xlsx', 'csv'), true)) {
            throw new \RuntimeException(
                'Formato no permitido. Use XLS, XLSX o CSV.'
            );
        }

        if (!isset($file['tmp_name'])
            || !is_file($file['tmp_name'])
        ) {
            throw new \RuntimeException(
                'El archivo temporal no existe.'
            );
        }

        $connection = $this->resource->getConnection();

        $active = $this->resource->getTableName('sepomex');
        $import = $this->resource->getTableName('sepomex_import');
        $new = $this->resource->getTableName('sepomex_new');
        $backup = $this->resource->getTableName('sepomex_backup');

        $cpHelper = $this->resource->getTableName(
            'sepomex_cp_routing_tmp'
        );

        $munHelper = $this->resource->getTableName(
            'sepomex_municipio_routing_tmp'
        );

        if (!$connection->isTableExists($active)) {
            throw new \RuntimeException(
                'No existe la tabla sepomex.'
            );
        }

        $lock = (int) $connection->fetchOne(
            "SELECT GET_LOCK('gdmexico_sepomex_import',0)"
        );

        if ($lock !== 1) {
            throw new \RuntimeException(
                'Ya existe otra importación SEPOMEX en ejecución.'
            );
        }

        $csv = null;
        $deleteCsvAfterImport = false;

        try {
            /*
            * 1. Preparar CSV de importación.
            *
            * XLS/XLSX:
            *   se convierten al formato CSV canónico.
            *
            * CSV:
            *   se utiliza directamente, sin pasar por
            *   PhpSpreadsheet.
            */
            if ($ext === 'csv') {
                $csv = $file['tmp_name'];
            } else {
                $converted = $this->converter->convert(
                    $file['tmp_name']
                );

                if (!isset($converted['path'])
                    || !is_file($converted['path'])
                ) {
                    throw new \RuntimeException(
                        'No se generó correctamente el CSV temporal.'
                    );
                }

                $csv = $converted['path'];
                $deleteCsvAfterImport = true;
            }

            /*
             * 2. Validar estructura de la tabla activa.
             */
            $description = $connection->describeTable($active);
            $columns = array_keys($description);

            foreach ($this->baseColumns as $column) {
                if (!in_array($column, $columns, true)) {
                    throw new \RuntimeException(
                        'La tabla sepomex no contiene la columna requerida: '
                        . $column
                    );
                }
            }

            $hasRouting = in_array(
                'store_url',
                $columns,
                true
            ) && in_array(
                'exclusion_url',
                $columns,
                true
            );

            /*
             * Guardamos estadísticas de la tabla actual para comparar
             * contra el nuevo catálogo.
             */
            $oldRecords = (int) $connection->fetchOne(
                "SELECT COUNT(*) FROM {$active}"
            );

            $oldCps = (int) $connection->fetchOne(
                "SELECT COUNT(DISTINCT cp) FROM {$active}"
            );

            /*
             * 3. Crear staging heredando EXACTAMENTE estructura,
             * charset y collation de sepomex.
             */
            $connection->query(
                "DROP TABLE IF EXISTS {$import}"
            );

            $connection->query(
                "CREATE TABLE {$import} LIKE {$active}"
            );

            /*
             * 4. CSV -> staging.
             */
            $this->loadCsv(
                $connection,
                $import,
                $csv,
                $hasRouting
            );

            /*
             * 5. Validaciones del staging.
             */
            $records = (int) $connection->fetchOne(
                "SELECT COUNT(*) FROM {$import}"
            );

            $cps = (int) $connection->fetchOne(
                "SELECT COUNT(DISTINCT cp) FROM {$import}"
            );

            if ($records < 1 || $cps < 1) {
                throw new \RuntimeException(
                    'El CSV convertido no contiene datos válidos.'
                );
            }

            $blankCp = (int) $connection->fetchOne(
                "SELECT COUNT(*)
                 FROM {$import}
                 WHERE cp IS NULL
                    OR TRIM(cp) = ''"
            );

            if ($blankCp > 0) {
                throw new \RuntimeException(
                    'Se encontraron '
                    . $blankCp
                    . ' registros con CP vacío.'
                );
            }

            /*
             * Protección contra archivos claramente incompletos.
             *
             * No usamos 159326 como número fijo porque SEPOMEX cambia.
             * Comparamos contra el catálogo actualmente operativo.
             */
            if ($oldRecords > 0) {
                $minimumRecords = (int) floor(
                    $oldRecords * 0.80
                );

                if ($records < $minimumRecords) {
                    throw new \RuntimeException(
                        'Validación detenida: el nuevo archivo contiene '
                        . $records
                        . ' registros y el catálogo actual contiene '
                        . $oldRecords
                        . '. La reducción supera el 20%.'
                    );
                }
            }

            if ($oldCps > 0) {
                $minimumCps = (int) floor(
                    $oldCps * 0.80
                );

                if ($cps < $minimumCps) {
                    throw new \RuntimeException(
                        'Validación detenida: el nuevo archivo contiene '
                        . $cps
                        . ' CP distintos y el catálogo actual contiene '
                        . $oldCps
                        . '. La reducción supera el 20%.'
                    );
                }
            }

            /*
             * 6. Crear candidato.
             */
            $connection->query(
                "DROP TABLE IF EXISTS {$new}"
            );

            $connection->query(
                "CREATE TABLE {$new} LIKE {$active}"
            );

            $quotedColumns = array();

            foreach ($this->baseColumns as $column) {
                $quotedColumns[] = $connection->quoteIdentifier(
                    $column
                );
            }

            $columnsSql = implode(',', $quotedColumns);

            $connection->query(
                "INSERT INTO {$new} ({$columnsSql})
                 SELECT {$columnsSql}
                 FROM {$import}"
            );

            $candidateRecords = (int) $connection->fetchOne(
                "SELECT COUNT(*) FROM {$new}"
            );

            if ($candidateRecords !== $records) {
                throw new \RuntimeException(
                    'Validación detenida: staging tiene '
                    . $records
                    . ' registros pero el candidato tiene '
                    . $candidateRecords
                    . '.'
                );
            }

            /*
             * 7. Preservar routing.
             */
            if ($hasRouting) {
                $this->preserveRouting(
                    $connection,
                    $active,
                    $new,
                    $cpHelper,
                    $munHelper
                );
            }

            /*
             * 8. Últimas validaciones antes del backup/swap.
             */
            $finalRecords = (int) $connection->fetchOne(
                "SELECT COUNT(*) FROM {$new}"
            );

            $finalCps = (int) $connection->fetchOne(
                "SELECT COUNT(DISTINCT cp) FROM {$new}"
            );

            if ($finalRecords !== $records) {
                throw new \RuntimeException(
                    'Validación final detenida: cambió la cantidad '
                    . 'de registros del candidato.'
                );
            }

            if ($finalCps !== $cps) {
                throw new \RuntimeException(
                    'Validación final detenida: cambió la cantidad '
                    . 'de CP del candidato.'
                );
            }

            /*
             * 9. Limpiar helpers antes del swap.
             */
            $this->dropHelperTables(
                $connection,
                $cpHelper,
                $munHelper
            );

            /*
             * 10. Reemplazo atómico.
             *
             * Sólo llegamos aquí si todas las validaciones anteriores
             * terminaron correctamente.
             *
             * Se conserva UNA sola copia anterior.
             */
            $backupOld = $this->resource->getTableName(
                'sepomex_backup_old'
            );

            $connection->query(
                "DROP TABLE IF EXISTS {$backupOld}"
            );

            if ($connection->isTableExists($backup)) {
                /*
                 * Un solo RENAME TABLE: si alguna parte falla, MySQL
                 * no deja el catálogo a medias. El backup anterior se
                 * elimina únicamente DESPUÉS del swap exitoso.
                 */
                $connection->query(
                    "RENAME TABLE
                        {$backup} TO {$backupOld},
                        {$active} TO {$backup},
                        {$new} TO {$active}"
                );

                $connection->query(
                    "DROP TABLE {$backupOld}"
                );
            } else {
                $connection->query(
                    "RENAME TABLE
                        {$active} TO {$backup},
                        {$new} TO {$active}"
                );
            }

            return array(
                'records' => $records,
                'cps' => $cps,
                'csv' => $csv,
                'routing' => $hasRouting
            );
        } finally {
            /*
             * El CSV sólo es temporal.
             */
            if ($deleteCsvAfterImport
                && $csv !== null
                && is_file($csv)
            ) {
                @unlink($csv);
            }

            /*
             * Las tablas helper tampoco deben quedar acumuladas.
             */
            try {
                $this->dropHelperTables(
                    $connection,
                    $cpHelper,
                    $munHelper
                );
            } catch (\Exception $e) {
                // No ocultar el error principal.
            }

            try {
                $connection->query(
                    "SELECT RELEASE_LOCK(
                        'gdmexico_sepomex_import'
                    )"
                );
            } catch (\Exception $e) {
                // No ocultar el error principal.
            }
        }
    }

    private function loadCsv(
        $connection,
        $table,
        $csv,
        $hasRouting
    ) {
        $handle = fopen($csv, 'rb');

        if ($handle === false) {
            throw new \RuntimeException(
                'No fue posible abrir el CSV convertido.'
            );
        }

        /*
         * Consumimos encabezado.
         */
        $header = fgetcsv($handle, 0, '|');

        if (!is_array($header)) {
            fclose($handle);

            throw new \RuntimeException(
                'El CSV no contiene un encabezado válido.'
            );
        }

        /*
        * El CSV debe utilizar exactamente el formato canónico
        * empleado internamente por el importador.
        */
        $normalizedHeader = array();

        foreach ($header as $value) {
            $normalizedHeader[] = trim(
                (string) $value
            );
        }

        /*
        * El primer campo puede contener BOM UTF-8.
        */
        if (isset($normalizedHeader[0])) {
            $normalizedHeader[0] = preg_replace(
                '/^\xEF\xBB\xBF/',
                '',
                $normalizedHeader[0]
            );
        }

        if ($normalizedHeader !== $this->baseColumns) {
            fclose($handle);

            throw new \RuntimeException(
                'Encabezado CSV inválido. Se esperaba: '
                . implode('|', $this->baseColumns)
            );
        }

        $batch = array();
        $count = 0;

        try {
            while (
                ($row = fgetcsv($handle, 0, '|')) !== false
            ) {
                if (count($row) !== count($this->baseColumns)) {
                    throw new \RuntimeException(
                        'CSV inválido cerca del registro '
                        . ($count + 2)
                        . ': se esperaban '
                        . count($this->baseColumns)
                        . ' columnas y se encontraron '
                        . count($row)
                        . '.'
                    );
                }

                $data = array();

                foreach (
                    $this->baseColumns as $index => $column
                ) {
                    $data[$column] = isset($row[$index])
                        ? $row[$index]
                        : '';
                }

                if ($hasRouting) {
                    $data['store_url'] = null;
                    $data['exclusion_url'] = null;
                }

                $batch[] = $data;
                $count++;

                if (count($batch) >= self::DB_BATCH_SIZE) {
                    $connection->insertMultiple(
                        $table,
                        $batch
                    );

                    $batch = array();
                }
            }

            if (!empty($batch)) {
                $connection->insertMultiple(
                    $table,
                    $batch
                );
            }
        } finally {
            fclose($handle);
        }
    }

private function preserveRouting(
    $connection,
    $active,
    $new,
    $cpHelper,
    $munHelper
) {
    /*
     * Los helpers se conservan en la firma por compatibilidad con
     * versiones anteriores del módulo, pero ya no son necesarios.
     */
    $this->dropHelperTables(
        $connection,
        $cpHelper,
        $munHelper
    );

    /*
     * Ambos lados del JOIN exacto necesitan CP + asentamiento.
     *
     * ensureIndex() ya detecta si el índice existe, por lo que no
     * vuelve a crearlo innecesariamente.
     */
    $this->ensureIndex(
        $connection,
        $active,
        'IDX_SEPOMEX_CP_ASENTAMIENTO',
        'cp, asentamiento'
    );

    $this->ensureIndex(
        $connection,
        $new,
        'IDX_SEPOMEX_CP_ASENTAMIENTO',
        'cp, asentamiento'
    );

    /*
     * 1. VALIDAR PRIMERO que ninguna pareja histórica
     * CP + asentamiento tenga más de un routing.
     *
     * Si existe ambigüedad abortamos ANTES de copiar routing.
     */
    $ambiguousExact = (int) $connection->fetchOne(
        "SELECT COUNT(*)
         FROM (
            SELECT
                cp,
                asentamiento
            FROM {$active}
            GROUP BY cp, asentamiento
            HAVING COUNT(
                DISTINCT CONCAT(
                    COALESCE(store_url, '<NULL>'),
                    '|',
                    COALESCE(exclusion_url, '<NULL>')
                )
            ) > 1
         ) ambiguous"
    );

    if ($ambiguousExact > 0) {
        throw new \RuntimeException(
            'Validación detenida: '
            . $ambiguousExact
            . ' combinaciones CP + asentamiento tienen '
            . 'routing histórico ambiguo. '
            . 'La tabla activa no fue modificada.'
        );
    }

    /*
     * 2. Preservar routing histórico exacto.
     *
     * Como acabamos de demostrar que CP + asentamiento no tiene
     * routings distintos, podemos usar JOIN directo.
     *
     * Aunque existan filas históricas duplicadas, todas contienen
     * exactamente el mismo routing para esa pareja.
     *
     * Esta operación pasó de >13 minutos a ~2 segundos durante
     * las pruebas con 159326 registros.
     */
    $connection->query(
        "UPDATE {$new} n
         INNER JOIN {$active} o
            ON o.cp = n.cp
           AND o.asentamiento = n.asentamiento
         SET
            n.store_url = o.store_url,
            n.exclusion_url = o.exclusion_url"
    );
/*
 * Después de preservar el routing exacto, comprobar si realmente
 * existen filas pendientes.
 *
 * En una reimportación sin nuevos CP/asentamientos esto evita
 * ejecutar los fallbacks y sus GROUP BY innecesariamente.
 */
$pendingAfterExact = (int) $connection->fetchOne(
    "SELECT COUNT(*)
     FROM {$new}
     WHERE store_url IS NULL
       AND exclusion_url IS NULL"
);

if ($pendingAfterExact === 0) {
    /*
     * No hay nada que resolver mediante fallback.
     *
     * Sólo queda verificar que el routing histórico compartido
     * permanezca exactamente igual.
     */
    $changed = (int) $connection->fetchOne(
        "SELECT COUNT(DISTINCT n.entity_id)
         FROM {$new} n
         INNER JOIN {$active} o
            ON o.cp = n.cp
           AND o.asentamiento = n.asentamiento
         WHERE NOT (
            n.store_url <=> o.store_url
            AND n.exclusion_url <=> o.exclusion_url
         )"
    );

    if ($changed > 0) {
        throw new \RuntimeException(
            'Validación detenida: '
            . $changed
            . ' filas compartidas cambiarían de routing. '
            . 'La tabla activa no fue modificada.'
        );
    }

    return;
}
    /*
     * 3. Para asentamientos NUEVOS pertenecientes a un CP histórico,
     * heredar el routing únicamente cuando TODO el CP tiene un
     * único routing.
     *
     * Sólo toca filas que continúan sin routing.
     */
    $connection->query(
        "UPDATE {$new} n
         INNER JOIN (
            SELECT
                cp,
                MAX(store_url) AS store_url,
                MAX(exclusion_url) AS exclusion_url
            FROM {$active}
            GROUP BY cp
            HAVING COUNT(
                DISTINCT CONCAT(
                    COALESCE(store_url, '<NULL>'),
                    '|',
                    COALESCE(exclusion_url, '<NULL>')
                )
            ) = 1
         ) r
            ON r.cp = n.cp
         SET
            n.store_url = r.store_url,
            n.exclusion_url = r.exclusion_url
         WHERE n.store_url IS NULL
           AND n.exclusion_url IS NULL"
    );

    $normalState = "
        CASE estado
            WHEN 'Coahuila'
                THEN 'Coahuila de Zaragoza'
            WHEN 'Michoacán'
                THEN 'Michoacán de Ocampo'
            WHEN 'Veracruz'
                THEN 'Veracruz de Ignacio de la Llave'
            ELSE estado
        END
    ";

    /*
     * 4. Sólo para CP completamente nuevos:
     *
     * usar estado + municipio cuando esa combinación histórica
     * tenga un único routing.
     *
     * Nunca utilizamos id_estado porque sabemos que algunos códigos
     * históricos están reutilizados.
     */
    $connection->query(
        "UPDATE {$new} n
         LEFT JOIN (
            SELECT DISTINCT cp
            FROM {$active}
         ) oldcp
            ON oldcp.cp = n.cp
         INNER JOIN (
            SELECT
                {$normalState} AS estado_normalizado,
                municipio,
                MAX(store_url) AS store_url,
                MAX(exclusion_url) AS exclusion_url
            FROM {$active}
            GROUP BY {$normalState}, municipio
            HAVING COUNT(
                DISTINCT CONCAT(
                    COALESCE(store_url, '<NULL>'),
                    '|',
                    COALESCE(exclusion_url, '<NULL>')
                )
            ) = 1
         ) r
            ON r.estado_normalizado = (
                CASE n.estado
                    WHEN 'Coahuila'
                        THEN 'Coahuila de Zaragoza'
                    WHEN 'Michoacán'
                        THEN 'Michoacán de Ocampo'
                    WHEN 'Veracruz'
                        THEN 'Veracruz de Ignacio de la Llave'
                    ELSE n.estado
                END
            )
           AND r.municipio = n.municipio
         SET
            n.store_url = r.store_url,
            n.exclusion_url = r.exclusion_url
         WHERE oldcp.cp IS NULL
           AND n.store_url IS NULL
           AND n.exclusion_url IS NULL"
    );

    /*
     * 5. No permitir ninguna fila sin routing.
     *
     * store_url NULL es válido cuando existe exclusion_url.
     * Solamente consideramos pendiente cuando AMBOS son NULL.
     */
    $pending = (int) $connection->fetchOne(
        "SELECT COUNT(*)
         FROM {$new}
         WHERE store_url IS NULL
           AND exclusion_url IS NULL"
    );

    if ($pending > 0) {
        throw new \RuntimeException(
            'Validación detenida: '
            . $pending
            . ' filas quedaron sin routing. '
            . 'La tabla activa no fue modificada.'
        );
    }

    /*
     * 6. VALIDACIÓN CRÍTICA FINAL.
     *
     * Ya comprobamos al inicio que no existen routings ambiguos.
     * Por ello podemos comparar directamente contra la tabla
     * histórica sin repetir GROUP BY + COUNT(DISTINCT ...).
     *
     * Cualquier CP + asentamiento existente debe conservar
     * exactamente store_url y exclusion_url.
     */
    $changed = (int) $connection->fetchOne(
        "SELECT COUNT(DISTINCT n.entity_id)
         FROM {$new} n
         INNER JOIN {$active} o
            ON o.cp = n.cp
           AND o.asentamiento = n.asentamiento
         WHERE NOT (
            n.store_url <=> o.store_url
            AND n.exclusion_url <=> o.exclusion_url
         )"
    );

    if ($changed > 0) {
        throw new \RuntimeException(
            'Validación detenida: '
            . $changed
            . ' filas compartidas cambiarían de routing. '
            . 'La tabla activa no fue modificada.'
        );
    }
}

    private function ensureIndex(
        $connection,
        $table,
        $indexName,
        $columns
    ) {
        $requestedColumns = array_map(
            'trim',
            explode(',', $columns)
        );

        $indexes = $connection->fetchAll(
            "SHOW INDEX FROM {$table}"
        );

        $existing = array();

        foreach ($indexes as $index) {
            $name = $index['Key_name'];

            if (!isset($existing[$name])) {
                $existing[$name] = array();
            }

            $existing[$name][(int) $index['Seq_in_index']] =
                $index['Column_name'];
        }

        foreach ($existing as $existingColumns) {
            ksort($existingColumns);

            $existingColumns = array_values(
                $existingColumns
            );

            /*
             * Un índice compuesto también cubre su prefijo izquierdo.
             *
             * Ejemplo:
             * (cp, asentamiento) cubre una búsqueda por cp.
             */
            $prefix = array_slice(
                $existingColumns,
                0,
                count($requestedColumns)
            );

            if ($prefix === $requestedColumns) {
                return;
            }
        }

        $connection->query(
            "ALTER TABLE {$table}
             ADD INDEX {$indexName} ({$columns})"
        );
    }

    private function dropHelperTables(
        $connection,
        $cpHelper,
        $munHelper
    ) {
        $connection->query(
            "DROP TABLE IF EXISTS {$cpHelper}"
        );

        $connection->query(
            "DROP TABLE IF EXISTS {$munHelper}"
        );
    }
}
