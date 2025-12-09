<?php
session_start();
// --- Cấu hình kết nối SQL Server ---
$serverName = "localhost";
$connectionOptions = [
    "Database" => "GPLX_CSDT",
    "Uid" => "",
    "PWD" => "",
    "CharacterSet" => "UTF-8"
];

// Kết nối
$conn = sqlsrv_connect($serverName, $connectionOptions);
if ($conn === false) {
    die(print_r(sqlsrv_errors(), true));
}
?>
