<?php
// index.php
declare(strict_types=1);
ini_set('display_errors', '1');
error_reporting(E_ALL);

require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Shared\Date as XlDate;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function id_long_date_no_weekday(?DateTimeInterface $dt): string {
    if (!$dt) return '';
    static $bulan = [1=>'Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
    $d = (int)$dt->format('j');
    $m = (int)$dt->format('n');
    $y = (int)$dt->format('Y');
    return $d . ' ' . ($bulan[$m] ?? $dt->format('F')) . ' ' . $y;
}

function try_parse_excel_date($val): ?DateTimeInterface {
    if ($val === null || $val === '') return null;

    // Excel serial number
    if (is_numeric($val)) {
        try { return XlDate::excelToDateTimeObject((float)$val); } catch (\Throwable $e) { /* fallthrough */ }
    }

    // PhpSpreadsheet may already give DateTime
    if ($val instanceof DateTimeInterface) return $val;

    // Strings: yyyy-mm-dd, dd/mm/yyyy, dd-mm-yyyy
    if (is_string($val)) {
        $s = trim($val);

        if (preg_match('/^\s*(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})\s*$/', $s, $m)) {
            [$all,$y,$mo,$d] = $m;
            if (checkdate((int)$mo,(int)$d,(int)$y)) return new DateTime("$y-$mo-$d");
        }
        if (preg_match('/^\s*(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})\s*$/', $s, $m)) {
            [$all,$d,$mo,$y] = $m;
            $y = (int)$y;
            if ($y < 100) $y += ($y >= 70 ? 1900 : 2000);
            if (checkdate((int)$mo,(int)$d,(int)$y)) return new DateTime("$y-$mo-$d");
        }
    }
    return null;
}

function normalize_city(?string $raw): ?string {
    if ($raw === null) return null;
    $norm = strtoupper(trim(preg_replace('/\s+/',' ',(string)$raw)));
    $map = [
        'ONLINE'=>'Jakarta',
        'KOTA ADM JAKARTA PUSAT'=>'Jakarta',
        'KOTA ADM JAKARTA TIMUR'=>'Jakarta',
        'KOTA ADM JAKARTA BARAT'=>'Jakarta',
        'KOTA ADM JAKARTA UTARA'=>'Jakarta',
        'KOTA ADM JAKARTA SELATAN'=>'Jakarta',
        'KOTA BANDAR LAMPUNG'=>'Bandar Lampung',
        'KOTA PADANG'=>'Padang',
        'KABUPATEN BOGOR'=>'Bogor',
        'KABUPATEN WONOSOBO'=>'Wonosobo',
        'KOTA SEMARANG'=>'Semarang',
        'KOTA JAMBI'=>'Jambi',
        'KOTA MAKASSAR'=>'Makassar',
        'KOTA MANADO'=>'Manado',
        'KOTA MEDAN'=>'Medan',
        'KOTA PALANGKA RAYA'=>'Palangka Raya',
        'KOTA YOGYAKARTA'=>'Yogyakarta',
        'KOTA SURABAYA'=>'Surabaya',
        'KOTA AMBON'=>'Ambon',
        'KOTA BANJARMASIN'=>'Banjarmasin',
        'KOTA DENPASAR'=>'Denpasar',
        'KOTA PALEMBANG'=>'Palembang',
        'KOTA SAMARINDA'=>'Samarinda',
    ];
    return $map[$norm] ?? null;
}

function norm(string $s): string {
    return strtolower(trim(preg_replace('/\s+/', ' ', $s)));
}

// Header aliases
const HDR = [
    'DATE'        => ['tanggal uji (hh/bb/yyyy)'],
    'CITY'        => ['kota ujian'],
    'NO'          => ['no','no.','nourut','no urut','no_urut'],
    'ID_BSMR'     => ['id bsmr','id_bsmr','id bsmr.'],
    'NAMA'        => ['nama asesi','nama','nama peserta','nama_asesi'],
    'INSTANSI'    => ['instansi','perusahaan','institusi'],
    'KUALIFIKASI' => ['kualifikasi','tingkat','tingkat (kualifikasi)','tingkat/kualifikasi'],
    'KBK'         => ['k/bk','kbk','k\bk','k-bk'],
];

function findHeaderCol(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $ws, array $aliases, int $scanRows = 6): array {
    $dim = $ws->calculateWorksheetDimension();
    [$start, $end] = explode(':', $dim);
    [$sCol, $sRow] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::indexesFromString($start);
    [$eCol, $eRow] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::indexesFromString($end);

    $target = array_map('norm', $aliases);
    for ($r = $sRow; $r <= min($sRow + $scanRows, $eRow); $r++) {
        for ($c = $sCol; $c <= $eCol; $c++) {
            $val = $ws->getCellByColumnAndRow($c, $r)->getValue();
            if (in_array(norm((string)$val), $target, true)) {
                return ['row' => $r, 'col' => $c];
            }
        }
    }
    return ['row' => -1, 'col' => -1];
}

function streamXlsx(Spreadsheet $wb, string $filename): void {
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="' . $filename . '"');
    header('Cache-Control: max-age=0');
    $writer = new Xlsx($wb);
    $writer->save('php://output');
    exit;
}

$action = $_POST['action'] ?? '';
?>
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel Auto Converter (PHP) — Long Date + Normalisasi Kota + Cetak Sertifikat</title>
<style>
  :root{--bg:#0b1220;--card:#131b2e;--muted:#8ea1c0;--acc:#3aa6ff;--good:#1db954}
  body{font-family:system-ui,Segoe UI,Roboto,Arial;background:var(--bg);color:#e8eeff;margin:0;padding:20px}
  .card{background:var(--card);border-radius:14px;padding:20px;max-width:900px;margin:auto;box-shadow:0 10px 30px rgba(0,0,0,.3);border:1px solid #1f2942}
  h1{margin:0 0 10px;color:var(--acc);font-size:clamp(20px,3vw,28px)}
  p{color:var(--muted);margin:0 0 14px}
  .row{display:flex;gap:10px;flex-wrap:wrap}
  input[type=file]{display:block;margin:14px 0 12px;padding:10px;border:1px solid #2a3b66;border-radius:10px;width:100%;max-width:520px;background:#0f1730;color:#e8eeff}
  button{background:var(--good);color:#00180d;font-weight:700;border:none;border-radius:10px;padding:12px 16px;cursor:pointer}
  button.secondary{background:#3aa6ff;color:#001227}
  .note{font-size:13px;color:#a9b9dc;margin-top:8px}
</style>
</head>
<body>
<div class="card">
  <h1>Excel Auto Converter (PHP)</h1>
  <p>Upload Excel: otomatis konversi <b>TANGGAL UJI (hh/bb/yyyy)</b> → tanggal panjang (tanpa hari), normalisasi <b>KOTA UJIAN</b>. Pilih salah satu unduhan:</p>
  <form method="post" enctype="multipart/form-data">
    <div class="row">
      <input type="file" name="excel" accept=".xlsx,.xls,.xlsm,.xlsb,.csv" required>
      <button type="submit" name="action" value="full">Unduh Hasil (semua kolom)</button>
      <button type="submit" class="secondary" name="action" value="cert">Download Untuk Cetak Sertifikat</button>
    </div>
    <div class="note">
      • “Cetak Sertifikat” hanya baris dengan <b>K/BK = 'K'</b> dan kolom: <i>No, ID BSMR, NAMA ASESI, INSTANSI, KUALIFIKASI, TANGGAL, KOTA, K/BK</i>.
    </div>
  </form>
</div>
</body>
</html>
<?php
if (!$action) exit;

if (!isset($_FILES['excel']) || $_FILES['excel']['error'] !== UPLOAD_ERR_OK) {
    http_response_code(400);
    echo "<script>alert('Upload gagal.');</script>";
    exit;
}

$origName = $_FILES['excel']['name'] ?? 'Workbook.xlsx';
$baseName = preg_replace('/\.(xlsx|xls|xlsm|xlsb|csv)$/i', '', $origName);
$tmpPath  = $_FILES['excel']['tmp_name'];

try {
    $reader = IOFactory::createReaderForFile($tmpPath);
    $reader->setReadDataOnly(false);
    $wb = $reader->load($tmpPath);

    // Process: date conversion & city normalization IN-PLACE
    $totalDate = 0; $totalCity = 0;

    foreach ($wb->getAllSheets() as $ws) {
        $posDate = findHeaderCol($ws, HDR['DATE']);
        $posCity = findHeaderCol($ws, HDR['CITY']);
        $headerRow = max($posDate['row'], $posCity['row']);

        if ($posDate['col'] > -1 || $posCity['col'] > -1) {
            $dim = $ws->calculateWorksheetDimension();
            [$start, $end] = explode(':', $dim);
            [$sCol, $sRow] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::indexesFromString($start);
            [$eCol, $eRow] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::indexesFromString($end);

            for ($r = max($headerRow + 1, $sRow); $r <= $eRow; $r++) {
                if ($posDate['col'] > -1) {
                    $cell = $ws->getCellByColumnAndRow($posDate['col'], $r);
                    $val  = $cell->getValue();
                    $dt   = try_parse_excel_date($val);
                    $txt  = id_long_date_no_weekday($dt);
                    if ($txt !== '') { $cell->setValueExplicit($txt, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING); $totalDate++; }
                }
                if ($posCity['col'] > -1) {
                    $cell = $ws->getCellByColumnAndRow($posCity['col'], $r);
                    $mapped = normalize_city((string)$cell->getValue());
                    if ($mapped !== null) { $cell->setValueExplicit($mapped, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING); $totalCity++; }
                }
            }
        }
    }

    if ($action === 'full') {
        // Download full processed workbook
        $filename = $baseName . '-processed.xlsx';
        streamXlsx($wb, $filename);
    }

    if ($action === 'cert') {
        // Build certificate sheet with selected fields, only K/BK = 'K'
        $out = new Spreadsheet();
        $outWS = $out->getActiveSheet();
        $outWS->setTitle('SERTIFIKAT');

        // Header row
        $headers = ['No','ID BSMR','NAMA ASESI','INSTANSI','KUALIFIKASI','TANGGAL','KOTA','K/BK'];
        $outWS->fromArray($headers, null, 'A1');

        $rowOut = 2;

        foreach ($wb->getAllSheets() as $ws) {
            $dim = $ws->calculateWorksheetDimension();
            [$start, $end] = explode(':', $dim);
            [$sCol, $sRow] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::indexesFromString($start);
            [$eCol, $eRow] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::indexesFromString($end);

            $pNO   = findHeaderCol($ws, HDR['NO']);
            $pID   = findHeaderCol($ws, HDR['ID_BSMR']);
            $pNAMA = findHeaderCol($ws, HDR['NAMA']);
            $pINST = findHeaderCol($ws, HDR['INSTANSI']);
            $pKUAL = findHeaderCol($ws, HDR['KUALIFIKASI']);
            $pTGL  = findHeaderCol($ws, HDR['DATE']);
            $pKOTA = findHeaderCol($ws, HDR['CITY']);
            $pKBK  = findHeaderCol($ws, HDR['KBK']);

            $haveAny = max($pNO['row'],$pID['row'],$pNAMA['row'],$pINST['row'],$pKUAL['row'],$pTGL['row'],$pKOTA['row'],$pKBK['row']);
            if ($haveAny < 0) continue;

            $headerRow = max($pNO['row'],$pID['row'],$pNAMA['row'],$pINST['row'],$pKUAL['row'],$pTGL['row'],$pKOTA['row'],$pKBK['row']);

            for ($r = max($headerRow + 1, $sRow); $r <= $eRow; $r++) {
                $vKBK = ($pKBK['col'] > -1) ? (string)$ws->getCellByColumnAndRow($pKBK['col'], $r)->getValue() : '';
                if (strtoupper(trim($vKBK)) !== 'K') continue;

                // Helper to get string value
                $get = function(array $p) use ($ws, $r): string {
                    if ($p['col'] <= -1) return '';
                    return (string)$ws->getCellByColumnAndRow($p['col'], $r)->getValue();
                };

                $no   = $get($pNO);
                $id   = $get($pID);
                $nama = $get($pNAMA);
                $inst = $get($pINST);
                $kual = $get($pKUAL);

                // Ensure TANGGAL is long Indonesian text
                $tglRaw = $get($pTGL);
                $tglTxt = $tglRaw;
                if ($tglTxt === '' || preg_match('/^\d{1,2}\s+\p{L}+\s+\d{4}$/u', $tglTxt) !== 1) {
                    $dt    = try_parse_excel_date($ws->getCellByColumnAndRow($pTGL['col'], $r)->getValue());
                    $tglTxt = id_long_date_no_weekday($dt);
                }

                // Ensure KOTA normalized
                $kotaRaw = $get($pKOTA);
                $kotaMap = normalize_city($kotaRaw);
                $kota = $kotaMap ?? $kotaRaw;

                $outWS->fromArray([$no,$id,$nama,$inst,$kual,$tglTxt,$kota,$vKBK], null, 'A'.$rowOut);
                $rowOut++;
            }
        }

        // Autosize columns (optional)
        foreach (range('A','H') as $col) { $outWS->getColumnDimension($col)->setAutoSize(true); }

        $filename = $baseName . '-sertifikat.xlsx';
        streamXlsx($out, $filename);
    }

} catch (\Throwable $e) {
    http_response_code(500);
    echo "<pre>Processing error:\n" . htmlspecialchars($e->getMessage(), ENT_QUOTES, 'UTF-8') . "</pre>";
    exit;
}
