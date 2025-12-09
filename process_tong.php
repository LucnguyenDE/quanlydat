<?php
require "vendor/autoload.php";
require "db_connect.php";

use PhpOffice\PhpSpreadsheet\IOFactory;

if (!isset($_FILES['excel_file']) || $_FILES['excel_file']['error'] !== 0) {
    die("Vui lòng chọn file Excel!");
}

$fileTmp = $_FILES['excel_file']['tmp_name'];
$spreadsheet = IOFactory::load($fileTmp);
$sheet = $spreadsheet->getActiveSheet()->toArray();

$highestRow = count($sheet);

$currentMaHV = null;
$isReadingSession = false;

for ($i = 0; $i < $highestRow; $i++) {

    if (empty($sheet[$i][0])) continue;

    $label = trim($sheet[$i][0]);

    // ==== Bắt đầu 1 học viên mới ====
    if (mb_strtolower($label) == "họ và tên") {
        $currentMaHV = trim($sheet[$i+1][1]); // Lấy mã học viên ở dòng tiếp theo
        continue;
    }

    // ==== Khi gặp dòng "Tổng" thì lấy giá trị ====
    if (mb_strtolower($label) == "tổng") {

        if ($currentMaHV == null) continue;

        $tongThoiGian = floatval($sheet[$i][3]);  // Cột thời gian
        $tongQuangDuong = floatval($sheet[$i][4]); // Cột quãng đường

        $sql = "INSERT INTO QuanLyDAT (MaHocVien, TongThoiGian, TongQuangDuong)
                VALUES (?, ?, ?)";
        $params = [$currentMaHV, $tongThoiGian, $tongQuangDuong];

        $stmt = sqlsrv_prepare($conn, $sql, $params);
        if (!sqlsrv_execute($stmt)) {
            die(print_r(sqlsrv_errors(), true));
        }
    }
}

echo "✔ Đã import dữ liệu Tổng DAT thành công!";
?>
