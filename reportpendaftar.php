<?php
//membuat koneksi
include('conn.php');
require 'vendor/autoload.php'; // require file autoload.php dari folder vendor
//menggunakan phpOffice phpSpreadsheet
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//membuat Spreadsheet baru
$spreadsheet = new Spreadsheet();
//membuat variable sheet
$sheet = $spreadsheet->getActiveSheet();
// mengisi judul kolom pada spreadsheet
$sheet -> setCellValue('A1', 'No');
$sheet -> setCellValue('B1', 'Jenis Pendaftaran');
$sheet -> setCellValue('C1', 'Tanggal Masuk Sekolah');
$sheet -> setCellValue('D1', 'NIS');
$sheet -> setCellValue('E1', 'Nomor Peserta');
$sheet -> setCellValue('F1', 'Pernah Paud');
$sheet -> setCellValue('G1', 'Pernah TK');
$sheet -> setCellValue('H1', 'No. Seri SKHUN ');
$sheet -> setCellValue('I1', 'No. Seri Ijazah ');
$sheet -> setCellValue('J1', 'Hobi');
$sheet -> setCellValue('K1', 'Cita - cita');
$sheet -> setCellValue('L1', 'Nama');
$sheet -> setCellValue('M1', 'Jenis Kelamin');
$sheet -> setCellValue('N1', 'NISN');
$sheet -> setCellValue('O1', 'NIK');
$sheet -> setCellValue('P1', 'Tempat Lahir');
$sheet -> setCellValue('Q1', 'Tanggal Lahir');
$sheet -> setCellValue('R1', 'Agama');
$sheet -> setCellValue('S1', 'Berkebutuhan Khusus');
$sheet -> setCellValue('T1', 'Alamat');
$sheet -> setCellValue('U1', 'RT');
$sheet -> setCellValue('V1', 'RW');
$sheet -> setCellValue('W1', 'Nama Dusun');
$sheet -> setCellValue('X1', 'Kelurahan');
$sheet -> setCellValue('Y1', 'Kecamatan');
$sheet -> setCellValue('Z1', 'Kode Pos');
$sheet -> setCellValue('AA1', 'Tempat Tinggal');
$sheet -> setCellValue('AB1', 'Moda Transportasi');
$sheet -> setCellValue('AC1', 'No HP');
$sheet -> setCellValue('AD1', 'No Telepon');
$sheet -> setCellValue('AE1', 'E-mail Pribadi');
$sheet -> setCellValue('AF1', 'Penerima KPS/PKH/KIP');
$sheet -> setCellValue('AG1', 'No. KPS/PKH/KIP');
$sheet -> setCellValue('AH1', 'Kewarganegaraan');
$sheet -> setCellValue('AI1', 'Negara');

//query sql untuk memasukan data
$query = mysqli_query($koneksi, "SELECT * FROM pendaftar_pdidik");
$i = 2;
$no = 1;
while($row = mysqli_fetch_array($query)){
    $sheet -> setCellValue('A'.$i, $no++);
    $sheet -> setCellValue('B'.$i, $row['jenis_pendaftaran']);
    $sheet -> setCellValue('C'.$i, $row['tanggal_masuk']);
    $sheet -> setCellValue('D'.$i, $row['nis']);
    $sheet -> setCellValue('E'.$i, $row['no_peserta']);
    $sheet -> setCellValue('F'.$i, $row['paud']);
    $sheet -> setCellValue('G'.$i, $row['tk']);
    $sheet -> setCellValue('H'.$i, $row['no_skhun']);
    $sheet -> setCellValue('I'.$i, $row['no_ijazah']);
    $sheet -> setCellValue('J'.$i, $row['hobi']);
    $sheet -> setCellValue('K'.$i, $row['cita_cita']);
    $sheet -> setCellValue('L'.$i, $row['nama']);
    $sheet -> setCellValue('M'.$i, $row['jenis_kelamin']);
    $sheet -> setCellValue('N'.$i, $row['nisn']);
    $sheet -> setCellValue('O'.$i, $row['nik']);
    $sheet -> setCellValue('P'.$i, $row['tempat_lahir']);
    $sheet -> setCellValue('Q'.$i, $row['tanggal_lahir']);
    $sheet -> setCellValue('R'.$i, $row['agama']);
    $sheet -> setCellValue('S'.$i, $row['berkebutuhan_khusus']);
    $sheet -> setCellValue('T'.$i, $row['alamat']);
    $sheet -> setCellValue('U'.$i, $row['rt']);
    $sheet -> setCellValue('V'.$i, $row['rw']);
    $sheet -> setCellValue('W'.$i, $row['nama_dusun']);
    $sheet -> setCellValue('X'.$i, $row['kelurahan']);
    $sheet -> setCellValue('Y'.$i, $row['kecamatan']);
    $sheet -> setCellValue('Z'.$i, $row['kodepos']);
    $sheet -> setCellValue('AA'.$i, $row['tempat_tinggal']);
    $sheet -> setCellValue('AB'.$i, $row['transportasi']);
    $sheet -> setCellValue('AC'.$i, $row['hp']);
    $sheet -> setCellValue('AD'.$i, $row['telp']);
    $sheet -> setCellValue('AE'.$i, $row['email']);
    $sheet -> setCellValue('AF'.$i, $row['penerima_kps']);
    $sheet -> setCellValue('AG'.$i, $row['no_kps']);
    $sheet -> setCellValue('AH'.$i, $row['kewarganegaraan']);
    $sheet -> setCellValue('AI'.$i, $row['negara']);
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
$sheet->getStyle('A1:AI'.$i)->applyFromArray($styleArray);

//membuat file Xlsx dengan nama Report Pendaftaran.xlsx
$writer = new Xlsx($spreadsheet);
$writer->save('Report Pendaftaran.xlsx');
?>