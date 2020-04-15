<?php
//memasukkan file koneksi.php yang dibutuhkan untuk koneksi database
include('koneksi.php');

//perintah import default aplikasi composer
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
 
//membuat instance spreadsheet baru
$spreadsheet = new Spreadsheet();
//memilih sheet yang sedang aktif
$sheet = $spreadsheet->getActiveSheet();
//membuat header kolom A B C dan D
$sheet->setCellValue('A1', 'No');
$sheet->setCellValue('B1', 'Nama');
$sheet->setCellValue('C1', 'Kelas');
$sheet->setCellValue('D1', 'Alamat');
 
//pemanggilan query select 
$query = mysqli_query($koneksi,"select * from tb_siswa");
//agar tabel yang ditulis oleh pemanggilan query dimulai dari baris kedua
$i = 2;
//variabel bantuan nomor urut dimulai pada kolom A2
$no = 1;
//looping 
while($row = mysqli_fetch_array($query))
{
	$sheet->setCellValue('A'.$i, $no++);
	$sheet->setCellValue('B'.$i, $row['nama']);
	$sheet->setCellValue('C'.$i, $row['kelas']);
	$sheet->setCellValue('D'.$i, $row['alamat']);	
	$i++;
}

//mengatur border tabel pada excel menjadi border tipis 
$styleArray = [
			'borders' => [
				'allBorders' => [
					'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				],
			],
		];
//membuat variabel bantuan i untuk menentukan kolom dan baris terakhir yang akan diberi style border.
$i = $i - 1;
//mengaplikasikan style border pada kolom dan baris yang dipilih atau dari A1 sampai D pada kolom ke $i
$sheet->getStyle('A1:D'.$i)->applyFromArray($styleArray);
 
//membuat instance baru yang berguna untuk menyimpan ke dalam file excel 
$writer = new Xlsx($spreadsheet);
//menyimpan ke dalam file excel sesuai nama
$writer->save('Report Data Siswa.xlsx');
?>