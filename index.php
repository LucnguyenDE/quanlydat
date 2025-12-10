<!DOCTYPE html>
<html>
<head>
<title>Upload DAT Excel</title>
<style>
    body {
        font-family: Arial, sans-serif;
        padding: 30px;
        display: flex;
        flex-direction: column;
        align-items: center;
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
 
<div class="form-box">
    <h2>📥 Nạp excel tổng hợp DAT</h2>
    <form action="nap_excel.php" method="POST" enctype="multipart/form-data">
        <input type="file" name="excel_file" accept=".xlsx,.xls" required>
        <br><br>
        <button type="submit">Xác nhận</button>
    </form>
</div>
<!-- Nút tải báo cáo Dat Phien -->
<a href="tai_file_word.php" class="btn btn-success" style="margin: 10px;">
    📥 Tải file Word cho toàn bộ học viên 
</a>   
</body>
</html>
