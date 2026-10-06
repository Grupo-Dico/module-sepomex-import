<?php

namespace GDMexico\SepomexImport\Model;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

class ExcelConverter
{
    /*
     * Si una hoja futura supera este tamaño utilizamos chunks.
     *
     * El archivo oficial actual tiene como hoja más grande
     * aproximadamente 11,558 filas, por lo que normalmente se
     * procesará una hoja completa.
     */
    const MAX_FULL_SHEET_ROWS = 20000;

    /*
     * Tamaño utilizado solamente como fallback para hojas
     * excepcionalmente grandes.
     */
    const FALLBACK_CHUNK_SIZE = 5000;

    private $header = array(
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

    /**
     * Conversión de c_estado SEPOMEX al código utilizado
     * actualmente por la tabla sepomex.
     */
    private $stateCodes = array(
        '01' => 'AG',
        '02' => 'BC',
        '03' => 'BS',
        '04' => 'CC',
        '05' => 'CL',
        '06' => 'MC',
        '07' => 'CS',
        '08' => 'DL',
        '09' => 'CDMX',
        '10' => 'DG',
        '11' => 'GT',
        '12' => 'GR',
        '13' => 'HG',
        '14' => 'JC',
        '15' => 'MC',
        '16' => 'MN',
        '17' => 'MS',
        '18' => 'NT',
        '19' => 'NL',
        '20' => 'OC',
        '21' => 'PL',
        '22' => 'QT',
        '23' => 'QR',
        '24' => 'SP',
        '25' => 'SL',
        '26' => 'SR',
        '27' => 'TC',
        '28' => 'TS',
        '29' => 'TL',
        '30' => 'VZ',
        '31' => 'YN',
        '32' => 'ZS'
    );

    private $expectedHeader = array(
        'd_codigo',
        'd_asenta',
        'd_tipo_asenta',
        'D_mnpio',
        'd_estado',
        'd_ciudad',
        'd_CP',
        'c_estado',
        'c_oficina',
        'c_CP',
        'c_tipo_asenta',
        'c_mnpio',
        'id_asenta_cpcons',
        'd_zona',
        'c_cve_ciudad'
    );

    public function convert($sourcePath, $destinationPath = null)
    {
        if (!class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory')) {
            throw new \RuntimeException(
                'PhpSpreadsheet no está disponible en esta instalación.'
            );
        }

        if (!is_file($sourcePath) || !is_readable($sourcePath)) {
            throw new \RuntimeException(
                'El archivo Excel de origen no existe o no puede leerse.'
            );
        }

        $destinationPath = $this->prepareDestination(
            $destinationPath
        );

        $handle = fopen($destinationPath, 'wb');

        if ($handle === false) {
            throw new \RuntimeException(
                'No fue posible crear el CSV temporal.'
            );
        }

        $recordsWritten = 0;

        try {
            $inputType = IOFactory::identify($sourcePath);

            if (!in_array(
                $inputType,
                array('Xls', 'Xlsx'),
                true
            )) {
                throw new \RuntimeException(
                    'Formato Excel no soportado: '
                    . $inputType
                );
            }

            if (fputcsv(
                $handle,
                $this->header,
                '|'
            ) === false) {
                throw new \RuntimeException(
                    'No fue posible escribir '
                    . 'el encabezado del CSV.'
                );
            }

            /*
             * Esta lectura obtiene solamente metadatos:
             * nombres de hojas y dimensiones.
             */
            $infoReader = IOFactory::createReader(
                $inputType
            );

            $infoReader->setReadDataOnly(true);

            $worksheetInfo =
                $infoReader->listWorksheetInfo(
                    $sourcePath
                );

            unset($infoReader);

            foreach ($worksheetInfo as $sheetInfo) {
                if (empty(
                    $sheetInfo['worksheetName']
                )) {
                    continue;
                }

                $sheetName =
                    $sheetInfo['worksheetName'];

                /*
                 * La hoja Nota contiene metadatos
                 * del archivo SEPOMEX.
                 */
                if (strcasecmp(
                    $sheetName,
                    'Nota'
                ) === 0) {
                    continue;
                }

                $totalRows =
                    isset($sheetInfo['totalRows'])
                    ? (int) $sheetInfo['totalRows']
                    : 0;

                if ($totalRows < 2) {
                    continue;
                }

                /*
                 * En archivos oficiales normales procesamos
                 * toda la hoja de una sola vez.
                 *
                 * Si alguna versión futura contiene una hoja
                 * demasiado grande utilizamos chunks.
                 */
                if (
                    $totalRows
                    <= self::MAX_FULL_SHEET_ROWS
                ) {
                    $recordsWritten +=
                        $this->processFullSheet(
                            $sourcePath,
                            $inputType,
                            $sheetName,
                            $totalRows,
                            $handle
                        );
                } else {
                    $recordsWritten +=
                        $this->processSheetInChunks(
                            $sourcePath,
                            $inputType,
                            $sheetName,
                            $totalRows,
                            $handle
                        );
                }

                /*
                 * PhpSpreadsheet crea bastantes objetos.
                 * Forzamos GC entre estados para mantener
                 * estable la memoria durante archivos grandes.
                 */
                gc_collect_cycles();
            }

            if ($recordsWritten < 1) {
                throw new \RuntimeException(
                    'No se encontraron registros '
                    . 'SEPOMEX válidos.'
                );
            }
        } catch (\Throwable $e) {
            fclose($handle);

            if (is_file($destinationPath)) {
                @unlink($destinationPath);
            }

            throw new \RuntimeException(
                'Error convirtiendo Excel a CSV: '
                . $e->getMessage(),
                0,
                $e
            );
        }

        fclose($handle);

        return array(
            'path' => $destinationPath,
            'records' => $recordsWritten
        );
    }

    /**
     * Procesa una hoja completa.
     *
     * Esta es la ruta rápida utilizada normalmente.
     */
    private function processFullSheet(
        $sourcePath,
        $inputType,
        $sheetName,
        $totalRows,
        $handle
    ) {
        $reader = IOFactory::createReader(
            $inputType
        );

        $reader->setReadDataOnly(true);
        $reader->setLoadSheetsOnly($sheetName);

        /*
         * Sólo necesitamos A:O.
         *
         * Se incluye desde fila 1 para poder validar
         * el encabezado utilizando esta misma carga.
         */
        $filter = new SepomexSheetReadFilter(
            1,
            $totalRows
        );

        $reader->setReadFilter($filter);

        $spreadsheet = $reader->load(
            $sourcePath
        );

        $worksheet =
            $spreadsheet->getSheetByName(
                $sheetName
            );

        if ($worksheet === null) {
            $spreadsheet->disconnectWorksheets();

            unset(
                $spreadsheet,
                $reader,
                $filter
            );

            throw new \RuntimeException(
                'No fue posible leer la hoja: '
                . $sheetName
            );
        }

        try {
            /*
             * Validamos el encabezado dentro de la misma
             * carga. Ya no volvemos a abrir el XLS.
             */
            $this->validateLoadedSheetHeader(
                $worksheet,
                $sheetName
            );

            /*
             * Una sola extracción para toda la hoja.
             *
             * Evita ejecutar rangeToArray()
             * una vez por cada registro.
             */
            $rows = $worksheet->rangeToArray(
                'A2:O' . $totalRows,
                null,
                false,
                false,
                false
            );

            $written = 0;
            $excelRow = 2;

            foreach ($rows as $row) {
                if ($this->writeRow(
                    $handle,
                    $row,
                    $sheetName,
                    $excelRow
                )) {
                    $written++;
                }

                $excelRow++;
            }

            unset($rows);
        } catch (\Throwable $e) {
            $spreadsheet->disconnectWorksheets();

            unset(
                $worksheet,
                $spreadsheet,
                $reader,
                $filter
            );

            gc_collect_cycles();

            throw $e;
        }

        $spreadsheet->disconnectWorksheets();

        unset(
            $worksheet,
            $spreadsheet,
            $reader,
            $filter
        );

        gc_collect_cycles();

        return $written;
    }

    /**
     * Fallback para hojas excepcionalmente grandes.
     */
    private function processSheetInChunks(
        $sourcePath,
        $inputType,
        $sheetName,
        $totalRows,
        $handle
    ) {
        $written = 0;

        /*
         * Primero cargamos únicamente el header.
         */
        $reader = IOFactory::createReader(
            $inputType
        );

        $reader->setReadDataOnly(true);
        $reader->setLoadSheetsOnly($sheetName);

        $headerFilter =
            new SepomexSheetReadFilter(
                1,
                1
            );

        $reader->setReadFilter(
            $headerFilter
        );

        $spreadsheet = $reader->load(
            $sourcePath
        );

        $worksheet =
            $spreadsheet->getSheetByName(
                $sheetName
            );

        if ($worksheet === null) {
            $spreadsheet->disconnectWorksheets();

            unset(
                $spreadsheet,
                $reader,
                $headerFilter
            );

            throw new \RuntimeException(
                'No fue posible validar la hoja: '
                . $sheetName
            );
        }

        $this->validateLoadedSheetHeader(
            $worksheet,
            $sheetName
        );

        $spreadsheet->disconnectWorksheets();

        unset(
            $worksheet,
            $spreadsheet,
            $reader,
            $headerFilter
        );

        gc_collect_cycles();

        /*
         * Datos desde fila 2.
         */
        for (
            $startRow = 2;
            $startRow <= $totalRows;
            $startRow += self::FALLBACK_CHUNK_SIZE
        ) {
            $endRow = min(
                $startRow
                + self::FALLBACK_CHUNK_SIZE
                - 1,
                $totalRows
            );

            $filter =
                new SepomexSheetReadFilter(
                    $startRow,
                    $endRow
                );

            $reader = IOFactory::createReader(
                $inputType
            );

            $reader->setReadDataOnly(true);
            $reader->setLoadSheetsOnly(
                $sheetName
            );
            $reader->setReadFilter($filter);

            $spreadsheet = $reader->load(
                $sourcePath
            );

            $worksheet =
                $spreadsheet->getSheetByName(
                    $sheetName
                );

            if ($worksheet === null) {
                $spreadsheet
                    ->disconnectWorksheets();

                unset(
                    $spreadsheet,
                    $reader,
                    $filter
                );

                throw new \RuntimeException(
                    'No fue posible leer la hoja: '
                    . $sheetName
                );
            }

            try {
                /*
                 * Extraemos el chunk completo.
                 */
                $rows = $worksheet->rangeToArray(
                    'A'
                    . $startRow
                    . ':O'
                    . $endRow,
                    null,
                    false,
                    false,
                    false
                );

                $excelRow = $startRow;

                foreach ($rows as $row) {
                    if ($this->writeRow(
                        $handle,
                        $row,
                        $sheetName,
                        $excelRow
                    )) {
                        $written++;
                    }

                    $excelRow++;
                }

                unset($rows);
            } catch (\Throwable $e) {
                $spreadsheet
                    ->disconnectWorksheets();

                unset(
                    $worksheet,
                    $spreadsheet,
                    $reader,
                    $filter
                );

                gc_collect_cycles();

                throw $e;
            }

            $spreadsheet->disconnectWorksheets();

            unset(
                $worksheet,
                $spreadsheet,
                $reader,
                $filter
            );

            gc_collect_cycles();
        }

        return $written;
    }

    /**
     * Valida el encabezado oficial SEPOMEX.
     */
    private function validateLoadedSheetHeader(
        $worksheet,
        $sheetName
    ) {
        $headerData = $worksheet->rangeToArray(
            'A1:O1',
            null,
            false,
            false,
            false
        );

        if (
            empty($headerData)
            || !isset($headerData[0])
        ) {
            throw new \RuntimeException(
                'No se pudo leer el encabezado '
                . 'de la hoja: '
                . $sheetName
            );
        }

        $actual = $this->normalizeRow(
            $headerData[0]
        );

        foreach (
            $this->expectedHeader
            as $index => $expectedValue
        ) {
            $actualValue =
                isset($actual[$index])
                ? $actual[$index]
                : '';

            if (strcasecmp(
                $actualValue,
                $expectedValue
            ) !== 0) {
                throw new \RuntimeException(
                    sprintf(
                        'Formato SEPOMEX inesperado '
                        . 'en hoja "%s". '
                        . 'Columna %d: esperado "%s", '
                        . 'recibido "%s".',
                        $sheetName,
                        $index + 1,
                        $expectedValue,
                        $actualValue
                    )
                );
            }
        }
    }

    /**
     * Convierte y escribe una fila SEPOMEX.
     *
     * Devuelve true si la fila fue escrita.
     */
    private function writeRow(
        $handle,
        array $row,
        $sheetName,
        $excelRow
    ) {
        $raw = $this->normalizeRow($row);

        if ($this->isEmptyRow($raw)) {
            return false;
        }

        if (count($raw) < 15) {
            return false;
        }

        /*
         * XLS oficial SEPOMEX:
         *
         * A  d_codigo
         * B  d_asenta
         * C  d_tipo_asenta
         * D  D_mnpio
         * E  d_estado
         * F  d_ciudad
         * G  d_CP
         * H  c_estado
         * I  c_oficina
         * J  c_CP
         * K  c_tipo_asenta
         * L  c_mnpio
         * M  id_asenta_cpcons
         * N  d_zona
         * O  c_cve_ciudad
         */

        /*
         * CP REAL = d_codigo = columna A.
         *
         * NO utilizar d_CP (columna G).
         */
        $cp = $this->normalizeCp($raw[0]);

        if ($cp === null) {
            return false;
        }

        /*
         * c_estado = columna H.
         */
        $numericStateCode =
            $this->normalizeStateCode(
                $raw[7]
            );

        if (!isset(
            $this->stateCodes[
                $numericStateCode
            ]
        )) {
            throw new \RuntimeException(
                sprintf(
                    'Código de estado inválido "%s" '
                    . 'en hoja "%s", fila %d.',
                    $raw[7],
                    $sheetName,
                    $excelRow
                )
            );
        }

        /*
         * Mapeo final hacia tabla sepomex:
         *
         * id_estado    <- c_estado convertido
         * estado       <- d_estado
         * id_municipio <- c_mnpio
         * municipio    <- D_mnpio
         * ciudad       <- d_ciudad
         * zona         <- d_zona
         * cp           <- d_codigo
         * asentamiento <- d_asenta
         * tipo         <- d_tipo_asenta
         * id           <- id_asenta_cpcons
         */
        $values = array(
            $this->stateCodes[
                $numericStateCode
            ],
            $raw[4],
            $this->normalizeIntegerString(
                $raw[11]
            ),
            $raw[3],
            $raw[5],
            $raw[13],
            $cp,
            $raw[1],
            $raw[2],
            $this->normalizeIntegerString(
                $raw[12]
            )
        );

        if (
            $values[1] === ''
            || $values[3] === ''
            || $values[7] === ''
        ) {
            return false;
        }

        if (fputcsv(
            $handle,
            $values,
            '|'
        ) === false) {
            throw new \RuntimeException(
                'No fue posible escribir '
                . 'una fila del CSV.'
            );
        }

        return true;
    }

    private function prepareDestination(
        $destinationPath
    ) {
        if (
            $destinationPath === null
            || $destinationPath === ''
        ) {
            $tempFile = tempnam(
                sys_get_temp_dir(),
                'sepomex_'
            );

            if ($tempFile === false) {
                throw new \RuntimeException(
                    'No fue posible crear '
                    . 'el archivo CSV temporal.'
                );
            }

            $destinationPath =
                $tempFile . '.csv';

            if (!@rename(
                $tempFile,
                $destinationPath
            )) {
                @unlink($tempFile);

                throw new \RuntimeException(
                    'No fue posible preparar '
                    . 'el CSV temporal.'
                );
            }
        }

        $directory = dirname(
            $destinationPath
        );

        if (!is_dir($directory)) {
            if (
                !mkdir(
                    $directory,
                    0775,
                    true
                )
                && !is_dir($directory)
            ) {
                throw new \RuntimeException(
                    'No fue posible crear '
                    . 'el directorio del CSV.'
                );
            }
        }

        return $destinationPath;
    }

    private function normalizeRow(array $row)
    {
        $result = array();

        foreach ($row as $value) {
            if ($value === null) {
                $result[] = '';
                continue;
            }

            if (is_bool($value)) {
                $result[] =
                    $value ? '1' : '0';
                continue;
            }

            if (
                is_float($value)
                && floor($value) == $value
            ) {
                $value =
                    (string) (int) $value;
            }

            $result[] =
                trim((string) $value);
        }

        return $result;
    }

    private function isEmptyRow(array $row)
    {
        foreach ($row as $value) {
            if ($value !== '') {
                return false;
            }
        }

        return true;
    }

    private function normalizeCp($value)
    {
        $value = trim(
            (string) $value
        );

        if ($value === '') {
            return null;
        }

        if (preg_match(
            '/^([0-9]+)(?:\.0+)?$/',
            $value,
            $matches
        )) {
            $value = $matches[1];
        } else {
            return null;
        }

        if (strlen($value) > 5) {
            return null;
        }

        return str_pad(
            $value,
            5,
            '0',
            STR_PAD_LEFT
        );
    }

    private function normalizeStateCode($value)
    {
        $value = trim(
            (string) $value
        );

        if (preg_match(
            '/^([0-9]+)(?:\.0+)?$/',
            $value,
            $matches
        )) {
            $value = $matches[1];
        }

        return str_pad(
            $value,
            2,
            '0',
            STR_PAD_LEFT
        );
    }

    private function normalizeIntegerString(
        $value
    ) {
        $value = trim(
            (string) $value
        );

        if ($value === '') {
            return '';
        }

        if (preg_match(
            '/^([0-9]+)(?:\.0+)?$/',
            $value,
            $matches
        )) {
            return (string) (
                (int) $matches[1]
            );
        }

        return $value;
    }
}

/**
 * Filtro de lectura SEPOMEX.
 *
 * Solamente permite:
 * - columnas A:O
 * - rango de filas solicitado
 */
class SepomexSheetReadFilter implements IReadFilter
{
    private $startRow;
    private $endRow;

    public function __construct(
        $startRow,
        $endRow
    ) {
        $this->startRow =
            (int) $startRow;

        $this->endRow =
            (int) $endRow;
    }

    public function readCell(
        $columnAddress,
        $row,
        $worksheetName = ''
    ) {
        if (
            strlen($columnAddress) > 1
            || $columnAddress < 'A'
            || $columnAddress > 'O'
        ) {
            return false;
        }

        return (int) $row
            >= $this->startRow
            && (int) $row
            <= $this->endRow;
    }
}
