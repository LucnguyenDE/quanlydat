<!DOCTYPE html>
<html>
<head>
<title>Upload DAT Excel</title>
<style>
    body {
        font-family: Arial, sans-serif;
        padding: 30px;
    }
    h2 {
        color: #0066cc;
    }
    .form-box {
        margin-bottom: 30px;
        padding: 20px;
        border: 1px solid #ccc;
        width: 350px;
        background: #f8f8f8;
    }
</style>
</head>
<body>
<!-- Nút tải báo cáo Dat Phien -->
<a href="tonghopdat_process.php" class="btn btn-success" style="margin: 10px;">
    📥 Tải báo cáo tổng hợp DAT
</a>
<h1>IMPORT DỮ LIỆU DAT</h1>
<a href="export_phien.php">
    <button>Tải Excel Phiên Chạy DAT</button>
</a>

<div class="form-box">
    <h2>📥 Nhập dữ liệu PHIÊN CHẠY</h2>
    <form action="process.php" method="POST" enctype="multipart/form-data">
        <input type="file" name="excel_file" accept=".xlsx,.xls" required>
        <br><br>
        <button type="submit">Nhập phiên chạy</button>
    </form>
</div>

<div class="form-box">
    <h2>📊 Nhập dữ liệu TỔNG DAT</h2>
    <form action="process_tong.php" method="POST" enctype="multipart/form-data">
        <input type="file" name="excel_file" accept=".xlsx,.xls" required>
        <br><br>
        <button type="submit">Nhập tổng DAT</button>
    </form>
</div>

</body>
</html>
