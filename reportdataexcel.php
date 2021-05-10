<?php
//membuat koneksi
include('koneksi.php');
require 'vendor/autoload.php'; // require file autoload.php dari folder vendor
//menggunakan phpOffice phpSpreadsheet
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//membuat Spreadsheet baru
$spreadsheet = new Spreadsheet();
//membuat variable sheet
$sheet = $spreadsheet->getActiveSheet();
// mengisi judul kolom pada spreadsheet
$sheet->setCellValue('A1', 'No');
$sheet->setCellValue('B1', 'Nama');
$sheet->setCellValue('C1', 'Kelas');
$sheet->setCellValue('D1', 'Alamat');

//query sql untuk memasukan data
$query = mysqli_query($koneksi, "SELECT * FROM tb_siswa");
$i = 2;
$no = 1;
while($row = mysqli_fetch_array($query)){
    $sheet->setCellValue('A'.$i, $no++);
    $sheet->setCellValue('B'.$i, $row['nama']);
    $sheet->setCellValue('C'.$i, $row['kelas']);
    $sheet->setCellValue('D'.$i, $row['alamat']);
    $i++;
}

// mengatur style pada border spreadsheet
$styleArray = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];

//menampilkan border
$i = $i - 1;
$sheet->getStyle('A1:D'.$i)->applyFromArray($styleArray);

//membuat file Xlsx dengan nama Report Data Siswa.xlsx
$writer = new Xlsx($spreadsheet);
$writer->save('Report Data Siswa.xlsx');
?>