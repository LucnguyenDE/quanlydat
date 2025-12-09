CREATE TABLE HocVien (
    MaHocVien NVARCHAR(50) PRIMARY KEY,
    HoVaTen NVARCHAR(100),
    NgaySinh DATE,
    MaKhoaHoc NVARCHAR(50),
    HangDaoTao NVARCHAR(20),
    CoSoDaoTao NVARCHAR(200)
);

CREATE TABLE DatPhienChay (
    ID INT IDENTITY(1,1) PRIMARY KEY,
    MaHocVien NVARCHAR(50),

    PhienChay NVARCHAR(200),
    BienSo NVARCHAR(20),
    Ngay DATETIME,
    ThoiGian FLOAT,
    QuangDuong FLOAT,
    GioDem FLOAT,
    GioTuDong FLOAT,

    TongThoiGian FLOAT,
    TongQuangDuong FLOAT,
    TongGioDem FLOAT,
    TongGioTuDong FLOAT,

    FOREIGN KEY (MaHocVien) REFERENCES HocVien(MaHocVien)
);

Mở C:\xampp\php\php.ini
extension=zip