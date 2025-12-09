<?php
require "vendor/autoload.php";
require "db_connect.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;

// Lấy dữ liệu
$sql = "SELECT MaHocVien, PhienChay, Ngay, ThoiGian, QuangDuong 
        FROM DatPhienChay";
$stmt = sqlsrv_query($conn, $sql);
if (!$stmt) die(print_r(sqlsrv_errors(), true));

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Cài font mặc định
$spreadsheet->getDefaultStyle()->getFont()->setName('Arial')->setSize(11);

// Tiêu đề cột
$headers = ['A1'=>'STT', 'B1'=>'Mã học viên', 'C1'=>'Mã phiên', 'D1'=>'Ngày đào tạo', 'E1'=>'Thời gian (giờ)', 'F1'=>'Quãng đường (km)'];
foreach ($headers as $cell => $value) {
    $sheet->setCellValue($cell, $value);
}

$sheet->getStyle('A1:F1')->applyFromArray([
    'font'=>['bold'=>true],
    'fill'=>['fillType'=>Fill::FILL_SOLID,'startColor'=>['argb'=>'FFDDEBF7']],
    'alignment'=>['horizontal'=>Alignment::HORIZONTAL_CENTER],
    'borders'=>['allBorders'=>['borderStyle'=>Border::BORDER_THIN]]
]);

$rowIndex = 2;
$hocVienSTT = 0;
$prevMaHV = null;
$isFirstRow = true;

// Tổng của từng học viên
$tongTG = $tongQD = 0;

while ($row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)) {

    // Khi gặp học viên mới
    if ($prevMaHV !== $row['MaHocVien']) {

        // Nếu không phải học viên đầu → ghi dòng TỔNG của học viên trước
        if ($prevMaHV !== null) {
            $sheet->setCellValue('D'.$rowIndex, "TỔNG");
            $sheet->setCellValue('E'.$rowIndex, $tongTG);
            $sheet->setCellValue('F'.$rowIndex, $tongQD);

            $sheet->getStyle('D'.$rowIndex.':F'.$rowIndex)->applyFromArray([
                'font'=>['bold'=>true],
                'fill'=>['fillType'=>Fill::FILL_SOLID, 'startColor'=>['argb'=>'FFFFFF99']],
                'borders'=>['allBorders'=>['borderStyle'=>Border::BORDER_THIN]]
            ]);

            $rowIndex += 2;
            $tongTG = $tongQD = 0;
        }

        $hocVienSTT++;
        $prevMaHV = $row['MaHocVien'];
        $isFirstRow = true;
    }

    // Ghi phiên chạy
    if ($isFirstRow) {
        $sheet->setCellValue('A'.$rowIndex, $hocVienSTT);
        $sheet->setCellValue('B'.$rowIndex, $row['MaHocVien']);
        $isFirstRow = false;
    }

    $sheet->setCellValue('C'.$rowIndex, $row['PhienChay']);
    $sheet->setCellValue('D'.$rowIndex, $row['Ngay']);
    $sheet->setCellValue('E'.$rowIndex, $row['ThoiGian']);
    $sheet->setCellValue('F'.$rowIndex, $row['QuangDuong']);

    // Style từng dòng
    $sheet->getStyle('A'.$rowIndex.':F'.$rowIndex)->applyFromArray([
        'borders'=>['allBorders'=>['borderStyle'=>Border::BORDER_THIN]],
        'alignment'=>['vertical'=>Alignment::VERTICAL_CENTER]
    ]);

    // Căn giữa cột text
    $sheet->getStyle('A'.$rowIndex.':C'.$rowIndex)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
    // Căn phải cột số
    $sheet->getStyle('E'.$rowIndex.':F'.$rowIndex)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);

    // Cộng dồn tổng
    $tongTG += $row['ThoiGian'];
    $tongQD += $row['QuangDuong'];

    $rowIndex++;
}

// Ghi tổng cho học viên cuối cùng
$sheet->setCellValue('D'.$rowIndex, "TỔNG");
$sheet->setCellValue('E'.$rowIndex, $tongTG);
$sheet->setCellValue('F'.$rowIndex, $tongQD);

$sheet->getStyle('D'.$rowIndex.':F'.$rowIndex)->applyFromArray([
    'font'=>['bold'=>true],
    'fill'=>['fillType'=>Fill::FILL_SOLID, 'startColor'=>['argb'=>'FFFFFF99']],
    'borders'=>['allBorders'=>['borderStyle'=>Border::BORDER_THIN]]
]);

// Tự căn độ rộng cột
foreach(range('A','F') as $col){
    $sheet->getColumnDimension($col)->setAutoSize(true);
}

// Xuất file
$fileName = "BaoCaoDAT.xlsx";
header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
header("Content-Disposition: attachment; filename=\"$fileName\"");
header("Cache-Control: max-age=0");

$writer = new Xlsx($spreadsheet);
$writer->save("php://output");
exit;
?>
