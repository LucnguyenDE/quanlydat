<?php
// tonghopdat_process.php
// KHÔNG được echo / var_dump / HTML gì ngoài file này

declare(strict_types=1);

// Tắt hiển thị lỗi ra trình duyệt (tránh làm hỏng ZIP)
ini_set('display_errors', '1');
error_reporting(E_ALL);

// Xóa toàn bộ buffer cũ (nếu có) rồi tạo buffer mới
while (ob_get_level()) {
    ob_end_clean();
}
ob_start();

require 'vendor/autoload.php';
require 'db_connect.php';

use PhpOffice\PhpWord\TemplateProcessor;
use ZipArchive;

// ===== HÀM LỌC KÝ TỰ CHO WORD XML =====
function cleanText($text) {
    if ($text === null) return '';
    // Bỏ control characters
    $text = preg_replace('/[\x00-\x1F\x7F]/u', '', $text);
    // Escape XML
    return htmlspecialchars($text, ENT_QUOTES | ENT_XML1, 'UTF-8');
}
function vn_to_ascii($str) {
    $str = iconv('UTF-8', 'ASCII//TRANSLIT//IGNORE', $str);
    return preg_replace('/[^A-Za-z0-9_\-]/', '_', $str);
}
// ===== CẤU HÌNH =====
$templatePath = __DIR__ . '/baocao_template.docx';
if (!file_exists($templatePath)) {
    // Không echo HTML, chỉ text đơn giản (vẫn sẽ làm ZIP lỗi, nhưng giúp debug)
    die('Không tìm thấy file template Word: ' . $templatePath);
}

$exportDir = __DIR__ . '/export_dat_phien/';
if (!is_dir($exportDir)) {
    mkdir($exportDir, 0777, true);
}

// ===== LẤY DANH SÁCH MÃ HỌC VIÊN =====
$sqlHV = "
    SELECT DISTINCT LTRIM(RTRIM(MaHocVien)) AS MaHocVien
    FROM DatPhienChay
    WHERE MaHocVien IS NOT NULL AND MaHocVien <> ''";
$stmtHV = sqlsrv_query($conn, $sqlHV);
if ($stmtHV === false) {
    // Nếu tới đây thì thôi, chưa tạo ZIP được
    die('Lỗi truy vấn học viên.');
}

$hasStudent = false;

while ($rowHV = sqlsrv_fetch_array($stmtHV, SQLSRV_FETCH_ASSOC)) {
    $maHV = trim($rowHV['MaHocVien'] ?? '');
    // Lấy tên học viên theo Mã học viên
    $tenHV = $maHV; // fallback nếu không tìm thấy
    $sqlTen = "SELECT HoVaTen FROM HocVien WHERE LTRIM(RTRIM(MaHocVien)) = ?";
    $stmtTen = sqlsrv_query($conn, $sqlTen, [$maHV]);
    if ($stmtTen && $rowTen = sqlsrv_fetch_array($stmtTen, SQLSRV_FETCH_ASSOC)) {
        $tenHV = trim($rowTen['HoVaTen']);
    }

    if ($maHV === '') continue;

    $hasStudent = true;

    // ===== LẤY CÁC PHIÊN CHẠY CỦA HỌC VIÊN NÀY =====
    $sqlPhien = "
        SELECT PhienChay, Ngay, ThoiGian, QuangDuong
        FROM DatPhienChay
        WHERE LTRIM(RTRIM(MaHocVien)) = ?";
    $stmtPhien = sqlsrv_query($conn, $sqlPhien, [$maHV]);
    if ($stmtPhien === false) {
        // Bỏ qua học viên lỗi, xử lý tiếp học viên khác
        continue;
    }

    $dataRows = [];
    $stt = 1;
    $tongTg = 0;
    $tongQd = 0;

    while ($p = sqlsrv_fetch_array($stmtPhien, SQLSRV_FETCH_ASSOC)) {
        // Ngày đào tạo
        $ngay = $p['Ngay'] ?? 0;
        $tg = floatval($p['ThoiGian'] ?? 0);
        $qd = floatval($p['QuangDuong'] ?? 0);
        $tongTg += $tg;
        $tongQd += $qd;

        $dataRows[] = [
            'STT'        => $stt++,
            'PhienChay'  => cleanText($p['PhienChay'] ?? ''),
            'NgayDaoTao' => $ngay,
            'ThoiGian'   => cleanText(number_format($tg, 2, ',', '')),
            'QuangDuong' => cleanText(number_format($qd, 2, ',', '')),
        ];
    }

    if (empty($dataRows)) {
        continue;
    }

    // ===== TẠO WORD TỪ TEMPLATE =====
    $template = new TemplateProcessor($templatePath);

    // Phần I – giống nhau cho tất cả học viên
    $template->setValue('NgayBaoCao',   cleanText('08/12/2025'));
    $template->setValue('HoVaTen',      cleanText($tenHV));
    $template->setValue('HangDaoTao',   cleanText('B.01'));
    $template->setValue('NgayKhaiGiang',cleanText('08/10/2025'));
    $template->setValue('NgayBeGiang',  cleanText('20/11/2025'));
    $template->setValue('CoSoDaoTao',   cleanText('Trung Tâm GDNN & SHLX Tư Thục Bình Phước'));

    // Nhân dòng bảng
    // Trong Word phải có đúng 1 dòng mẫu với:
    // ${STT} ${PhienChay} ${NgayDaoTao} ${ThoiGian} ${QuangDuong}
    $template->cloneRowAndSetValues('STT', $dataRows);
    $template->setValue('TongTG', cleanText(number_format($tongTg, 2, ',', '')));
    $template->setValue('TongQD', cleanText(number_format($tongQd, 2, ',', '')));

    // Lưu file Word theo mã học viên
    $safeName = preg_replace('/[\\\\\/:*?"<>|]/', '_', $tenHV);
    $wordPath = $exportDir . $safeName . '.docx';
    $template->saveAs($wordPath);
}

// Không có học viên nào
if (!$hasStudent) {
    @rmdir($exportDir);
    die('Không có dữ liệu học viên.');
}

// ===== TẠO FILE ZIP =====
$zipName = 'BaoCaoTongHop_' . date('Ymd_His') . '.zip';
$zipPath = __DIR__ . '/' . $zipName;

$zip = new ZipArchive();
if ($zip->open($zipPath, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
    die('Không tạo được file ZIP.');
}

$wordFiles = glob($exportDir . '*.docx');
if ($wordFiles) {
    foreach ($wordFiles as $file) {
        if (file_exists($file)) {
            $zip->addFile($file, basename($file));
        }
    }
}
$zip->close();

// Kiểm tra ZIP có dữ liệu không
if (!file_exists($zipPath) || filesize($zipPath) <= 0) {
    die('File ZIP rỗng hoặc không tồn tại.');
}

// ===== GỬI ZIP VỀ TRÌNH DUYỆT =====
// Xóa buffer HTML nếu lỡ có
if (ob_get_length()) {
    ob_clean();
}

header('Content-Type: application/zip');
header('Content-Disposition: attachment; filename="' . $zipName . '"');
header('Content-Length: ' . filesize($zipPath));
header('Pragma: public');
header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
header('Expires: 0');

// Đọc file theo từng khúc để tránh lỗi bộ nhớ
$fp = fopen($zipPath, 'rb');
while (!feof($fp)) {
    echo fread($fp, 8192);
}
fclose($fp);
flush();

// ===== DỌN RÁC SAU KHI GỬI XONG =====
if ($wordFiles) {
    foreach ($wordFiles as $file) {
        @unlink($file);
    }
}
@rmdir($exportDir);
@unlink($zipPath);

exit;
