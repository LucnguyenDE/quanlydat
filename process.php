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

$currentHocVien = null;
$isReadingSession = false;

// Biến cộng dồn cho từng học viên
$tongTG = $tongQD = $tongDem = $tongTD = 0;

for ($i = 0; $i < $highestRow; $i++) {

    // Nếu cell cột 0 không có dữ liệu -> bỏ qua
    if (empty($sheet[$i][0])) continue;

    $label = trim($sheet[$i][0]);

    // ========== GẶP "Họ và tên" -> HỌC VIÊN MỚI ==========
    if (mb_strtolower($label) == "họ và tên") {

        // Nếu là học viên tiếp theo -> reset tổng
        if ($currentHocVien !== null) {
            $tongTG = $tongQD = $tongDem = $tongTD = 0;
        }

        $hoTen      = trim($sheet[$i][1]);
        $maHocVien  = trim($sheet[$i+1][1]);
        $ngaySinh   = trim($sheet[$i+2][1]);
        $maKhoa     = trim($sheet[$i+3][1]);
        $hangDaoTao = trim($sheet[$i+4][1]);
        $coSo       = trim($sheet[$i+5][1]);

        $currentHocVien = $maHocVien;
        $isReadingSession = false;

        $sqlHV = "INSERT INTO HocVien (HoVaTen, MaHocVien, NgaySinh, MaKhoaHoc, HangDaoTao, CoSoDaoTao)
                  VALUES (?, ?, ?, ?, ?, ?)";

        $paramsHV = [$hoTen, $maHocVien, $ngaySinh, $maKhoa, $hangDaoTao, $coSo];
        $stmtHV = sqlsrv_prepare($conn, $sqlHV, $paramsHV);

        if (!sqlsrv_execute($stmtHV)) {
            die(print_r(sqlsrv_errors(), true));
        }

        continue;
    }

    // ========== GẶP "Phiên" thì bắt đầu đọc phiên ==========
    if (mb_strtolower($label) == "phiên") {
        $isReadingSession = true;
        continue;
    }

    // ========== ĐỌC PHIÊN CHẠY ==========
    if ($isReadingSession && $currentHocVien !== null) {

        if (empty($sheet[$i][1])) continue;

        $phien      = trim($sheet[$i][0]);
        $bienSo     = trim($sheet[$i][1]);
        $ngay       = (string)$sheet[$i][2];
        $thoiGian   = floatval($sheet[$i][3]);
        $quangDuong = floatval($sheet[$i][4]);
        $gioDem     = floatval($sheet[$i][5]);
        $gioTuDong  = floatval($sheet[$i][6]);

        // Cộng dồn theo từng học viên
        $tongTG += $thoiGian;
        $tongQD += $quangDuong;
        $tongDem += $gioDem;
        $tongTD += $gioTuDong;

        $sqlPhien = "INSERT INTO DatPhienChay
                     (MaHocVien, PhienChay, BienSo, Ngay, ThoiGian, QuangDuong, GioDem, GioTuDong,
                     TongThoiGian, TongQuangDuong, TongGioDem, TongGioTuDong)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

        $paramsPhien = [
            $currentHocVien, $phien, $bienSo, $ngay,
            $thoiGian, $quangDuong, $gioDem, $gioTuDong,
            $tongTG, $tongQD, $tongDem, $tongTD
        ];

        $stmtPhien = sqlsrv_prepare($conn, $sqlPhien, $paramsPhien);

        if (!sqlsrv_execute($stmtPhien)) {
            die(print_r(sqlsrv_errors(), true));
        }
    }
}

echo "✔ Đã import thành công toàn bộ học viên!";
?>
