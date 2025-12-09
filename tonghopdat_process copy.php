<?php
// KHÔNG được echo / var_dump trước khi header tải file!

require 'vendor/autoload.php';
require 'db_connect.php';

use PhpOffice\PhpWord\TemplateProcessor;
use ZipArchive;

// ====================== HÀM LỌC KÝ TỰ XML ======================
function cleanText($text) {
    if ($text === null) return '';
    $text = preg_replace('/[\x00-\x1F\x7F]/u', '', $text);
    return htmlspecialchars($text, ENT_QUOTES | ENT_XML1, 'UTF-8');
}

// ====================== CẤU HÌNH ======================
$templatePath = __DIR__ . '/baocao_template.docx';
if (!file_exists($templatePath)) {
    die('Không tìm thấy file template Word!');
}

$exportDir = __DIR__ . '/export_output/';
if (!is_dir($exportDir)) mkdir($exportDir, 0777, true);

// ====================== LẤY DS HỌC VIÊN ======================
$sqlHV = "
    SELECT DISTINCT LTRIM(RTRIM(MaHocVien)) AS MaHocVien
    FROM DatPhienChay
    WHERE MaHocVien IS NOT NULL AND MaHocVien <> ''
";

$stmtHV = sqlsrv_query($conn, $sqlHV);
if (!$stmtHV) die(print_r(sqlsrv_errors(), true));

$hasStudent = false;

while ($rowHV = sqlsrv_fetch_array($stmtHV, SQLSRV_FETCH_ASSOC)) {
    $hasStudent = true;
    $maHV = trim($rowHV['MaHocVien']);
    if ($maHV === '') continue;

    // ====================== LẤY PHIÊN CHẠY ======================
    $sqlPhien = "
        SELECT PhienChay, Ngay, ThoiGian, QuangDuong
        FROM DatPhienChay
        WHERE LTRIM(RTRIM(MaHocVien)) = ?
        ORDER BY Ngay
    ";
    $stmtPhien = sqlsrv_query($conn, $sqlPhien, [$maHV]);
    if (!$stmtPhien) continue;

    $dataRows = [];
    $stt = 1;
    $tongTg = 0;
    $tongQd = 0;

    while ($p = sqlsrv_fetch_array($stmtPhien, SQLSRV_FETCH_ASSOC)) {
        // Ngày
        $ngay = '';
        if (!empty($p['Ngay'])) {
            if ($p['Ngay'] instanceof DateTime) {
                $ngay = $p['Ngay']->format('d/m/Y');
            } else {
                $ngay = date('d/m/Y', strtotime($p['Ngay']));
            }
        }

        $tg = floatval($p['ThoiGian']);
        $qd = floatval($p['QuangDuong']);
        $tongTg += $tg;
        $tongQd += $qd;

        $dataRows[] = [
            'STT'        => $stt++,
            'PhienChay'  => cleanText($p['PhienChay']),
            'NgayDaoTao' => cleanText($ngay),
            'ThoiGian'   => cleanText(number_format($tg, 2, ',', '')),
            'QuangDuong' => cleanText(number_format($qd, 2, ',', ''))
        ];
    }

    if (empty($dataRows)) continue;

    // ====================== ĐIỀN WORD ======================
    $template = new TemplateProcessor($templatePath);

    // Phần I (giống nhau)
    $template->setValue('NgayBaoCao', cleanText('08/12/2025'));
    $template->setValue('MaKhoaHoc', cleanText('70001K25B0103'));
    $template->setValue('HangDaoTao', cleanText('B.01'));
    $template->setValue('NgayKhaiGiang', cleanText('08/10/2025'));
    $template->setValue('NgayBeGiang', cleanText('20/11/2025'));
    $template->setValue('CoSoDaoTao', cleanText('Trung Tâm GDNN & SHLX Tư Thục Bình Phước'));

    // Clone dòng bảng
    $template->cloneRowAndSetValues('STT', $dataRows);

    // Dòng tổng
    $template->setValue('TongThoiGian', cleanText(number_format($tongTg, 2, ',', '')));
    $template->setValue('TongQuangDuong', cleanText(number_format($tongQd, 2, ',', '')));

    // Lưu file theo Mã HV
    $safeHV = preg_replace('/[^A-Za-z0-9_\-]/', '_', $maHV);
    $outputWord = $exportDir . $safeHV . '.docx';
    $template->saveAs($outputWord);
}

// ====================== NẾU KHÔNG CÓ DỮ LIỆU ======================
if (!$hasStudent) {
    @rmdir($exportDir);
    die('Không có dữ liệu học viên.');
}

// ====================== TẠO ZIP ======================
$zipName = 'BaoCaoTongHop_' . date('Ymd_His') . '.zip';
$zipPath = __DIR__ . '/' . $zipName;

$zip = new ZipArchive();
if ($zip->open($zipPath, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
    die('Không tạo được file ZIP.');
}

foreach (glob($exportDir . '*.docx') as $file) {
    if (file_exists($file)) {
        $zip->addFile($file, basename($file));
    }
}
$zip->close(); // 🔥 đóng ZIP trước khi header

// ====================== TRẢ FILE ZIP ======================
header('Content-Type: application/zip');
header('Content-Disposition: attachment; filename="' . basename($zipPath) . '"');
header('Content-Length: ' . filesize($zipPath));
header('Pragma: public');
header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
header('Expires: 0');

// Không được echo gì ở trên
flush();
readfile($zipPath);

// ====================== DỌN RÁC SAU KHI TẢI XONG ======================
foreach (glob($exportDir . '*.docx') as $file) @unlink($file);
@rmdir($exportDir);
@unlink($zipPath);

exit;
