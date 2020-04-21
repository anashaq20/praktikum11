<!DOCTYPE html>
<html>
<head>
	<title></title>
<!-- Untuk referensi Bootstrap 4 sebagai style css -->
<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css" integrity="sha384-Vkoo8x4CGsO3+Hhxv8T/Q5PaXtkKtu6ug5TOeNV6gBiFeWPGFN9MuhOf23Q9Ifjh" crossorigin="anonymous">
		<?php
			//perintah import default aplikasi composer
			require 'vendor/autoload.php';
			use PhpOffice\PhpSpreadsheet\Spreadsheet;
			use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
		/* Digunakan untuk menginisiasi setiap variabel error sebagai variabel kosong*/
		$error_nama = $error_jk = $error_nisn = $error_nik = $error_tpt_lahir = $error_tgl_lhr = $error_no_akta = $error_agama = $error_kwn = $error_butuh_khs = $error_alamat = $error_rt = $error_rw = $error_dusun = $error_kelurahan = $error_kec = $error_kdpos = $error_lintang = $error_bujur = $error_tpt_tinggal = $error_moda_trans = $error_no_kks = $error_anak = $error_kps = $error_no_kps = "";
	/*Digunakan sebagai isiasi variabel field simpanan form awal sebagai variabel kosong*/
	$nama = $jk = $nisn = $nik = $tpt_lahir = $tgl_lhr = $no_akta = $agama = $kwn = $butuh_khs = $alamat = $rt = $rw = $dusun = $kelurahan = $kec = $kdpos = $lintang = $bujur = $tpt_tinggal = $moda_trans = $no_kks = $anak_ke = $ada_kps = $no_kps = "";
				/*Cek Jika perintah post adalah method*/
				if ($_SERVER["REQUEST_METHOD"] == "POST") {
						/*Cek kekosongan variabel nama*/
						if (empty($_POST["nama"])) {
							/*Variabel error nama dibuat dan disimpan sesuai format*/
							$error_nama = "Nama tidak boleh kosong";
						}
						else {
							/*cek_input(nama) untuk menghapus kesalahan spasial maupun lainnya pada field dan disimpan dalam variabel $nama*/
							$nama = cek_input($_POST["nama"]);
							/*cek apakah format nama sudah sesuai hanya huruf dan spasi*/
							if (!preg_match("/^[a-zA-Z ]*$/", $nama)) 
							{
								/*variabel error nama disimpan*/
								$error_nama = "Inputan hanya boleh huruf dan spasi";
							}
						}
						/*cek apakah jenis kelamin sudah dipilih*/
						if ($_POST["jk"] != "Laki-Laki" && ($_POST["jk"] != "Perempuan")) {
							/*variabel bantuan error akan tersimpan sesuai format*/
							$error_jk = "Jenis kelamin harus dipilih";
						}
						else {
							/*jika format sesuai maka jenis kelamin akan dirapikan penulisannya dan disimpan dalam variabel $jk*/
							$jk = cek_input($_POST["jk"]);
						}
						/*cek variabel nisn apakah kosong*/
						if (empty($_POST["nisn"])) {
							$error_nisn = "NISN tidak boleh kosong";
						}
						else
						{
							$nisn = cek_input($_POST["nisn"]);
							/*cek apakah variabel nisn hanya terisi oleh angka*/
							if (!is_numeric($nisn)) 
							{
								$error_nisn = "NISN hanya boleh angka";
							}
						}
						if (empty($_POST["nik"])) {
							$error_nik = "NIK tidak boleh kosong";
						}
						else
						{
							$nik = cek_input($_POST["nik"]);
							if (!is_numeric($nik)) 
							{
								$error_nik = "NIK hanya boleh angka";
							}
						}
						if (empty($_POST["tpt_lahir"])) {
							$error_tpt_lahir = "Tempat lahir tidak boleh kosong";
						}
						else
						{
							$tpt_lahir = cek_input($_POST["tpt_lahir"]);
						}
						if (!isset($_POST["tgl_lhr"])) {
							$error_telp = "Tanggal lahir tidak boleh kosong";
						}
						else {
							$tgl_lhr = $_POST["tgl_lhr"];
							/*mengubah format penulisan tanggal eropa menjadi sesuai dengan format database Y-m-d*/
							$tanggal = date('Y-m-d', strtotime($tgl_lhr));
						}
						if (empty($_POST["no_akta"])) {
							$error_no_akta = "Nomor Registrasi Akta tidak boleh kosong";
						}		
						else {
							$no_akta = cek_input($_POST["no_akta"]);
						}
						if (!isset($_POST["agama"])) {
							$error_agama = "Agama harus dipilih";
						}
						else
						{
							$agama = cek_input($_POST["agama"]);
						}
						if (!isset($_POST["kwn"])) {
							$error_kwn = "Kewarganegaraan harus dipilih";
						}
						else
						{
							$kwn = cek_input($_POST["kwn"]);
							if ($_POST["kwn"] != "Indonesia (WNI)" && ($_POST["kwn"] != "Asing (WNA)")) {
								$error_kwn = "Pilih salah satu";
							}
						}
						if (empty($_POST["butuh_khs"])) {
							$error_butuh_khs = "Kebutuhan khusus harus dipilih";
						}
						else
						{
							$butuh_khs = cek_input($_POST["butuh_khs"]);
						}
						if (empty($_POST["alamat"])) {
							$error_alamat = "Alamat tidak boleh kosong";
						}
						else
						{
							$alamat = cek_input($_POST["alamat"]);
						}
						if (empty($_POST["rt"])) {
							$error_rt = "RT tidak boleh kosong";
						}
						else
						{
							$rt = cek_input($_POST["rt"]);
							if (!is_numeric($rt)) {
								$error_rt = "RT harus berupa angka";
							}
						}
						if (empty($_POST["rw"])) {
							$error_rw = "RW tidak boleh kosong";
						}
						else
						{
							$rw = cek_input($_POST["rw"]);
							if (!is_numeric($rw)) {
								$error_rw = "RW harus berupa angka";
							}
						}
						if (empty($_POST["dusun"])) {
							$error_dusun = "Dusun tidak boleh kosong";
						}
						else
						{
							$dusun = cek_input($_POST["dusun"]);
						}
						if (empty($_POST["kelurahan"])) {
							$error_kelurahan = "Kelurahan tidak boleh kosong";
						}
						else
						{
							$kelurahan = cek_input($_POST["kelurahan"]);
						}
						if (empty($_POST["kec"])) {
							$error_kec = "Kecamatan tidak boleh kosong";
						}
						else
						{
							$kec = cek_input($_POST["kec"]);
						}
						if (empty($_POST["kdpos"])) {
							$error_kdpos = "Kode pos tidak boleh kosong";
						}
						else
						{
							$kdpos = cek_input($_POST["kdpos"]);
						}
						if (empty($_POST["lintang"])) {
							$error_lintang = "Garis Lintang Harus Diisi";
						}
						else
						{
							$lintang = cek_input($_POST["lintang"]);
						}
						if (empty($_POST["bujur"])) {
							$error_bujur = "Garis Bujur Harus Diisi";
						}
						else
						{
							$bujur = cek_input($_POST["bujur"]);
						}
						if (!isset($_POST["tpt_tinggal"])) {
							$error_tpt_tinggal = "Tempat tinggal Harus dipilih";
						}
						else
						{
							$tpt_tinggal = cek_input($_POST["tpt_tinggal"]);
						}
						if (!isset($_POST["moda_trans"])) {
							$error_moda_trans = "Moda Transportasi harus dipilih";
						}
						else
						{
							$moda_trans = cek_input($_POST["moda_trans"]);
						}
						if (empty($_POST["no_kks"])) {
							$error_no_kks = "No. KKS harus diisi";
						}
						else
						{
							$no_kks = cek_input($_POST["no_kks"]);
							if (!is_numeric($no_kks)) {
								$error_no_kks = "No. KKS hanya berisi angka"; 
							}
						}
						if (!isset($_POST["anak_ke"])) {
							$error_anak = "Anak keberapa anda harus diisi";
						}
						else
						{
							$anak_ke = cek_input($_POST["anak_ke"]);
							if (!is_numeric($anak_ke)) {
								$error_anak = "Hanya dapat berisi angka"; 
							}
						}
						if (!isset($_POST["ada_kps"])) {
							$error_kps = "Status KPS Harus dipilih";
						}
						else
						{
							$ada_kps = cek_input($_POST["ada_kps"]);
							$no_kps = cek_input($_POST["no_kps"]);
							if (($_POST["ada_kps"]) == "Ya" && (!isset($_POST["no_kps"]))) {
								$error_no_kps = "No. KPS harus diisi";	
								/*cek apakah no kps hanya terdiri dari angka dan huruf*/
								if (!ctype_alnum($no_kps)) {
									$error_no_kps = "No. KPS Hanya dapat berisi angka dan huruf"; 
								}
							 else {
								$no_kps = cek_input($_POST["no_kps"]);	
							}
						}
						}
						/*cek variabel simpan apakah sudah ditekan/terisi*/
						if (isset($_POST['simpan'])) {
								    // ambil data dari formulir
								    $id_pendaftar = NULL;
								    include('koneksi.php');
								    // buat query simpan
								    $sql = "INSERT INTO pendaftaran_siswa VALUES ('$id_pendaftar', '$nama', '$jk', '$nisn', '$nik', '$tpt_lahir', '$tanggal', '$no_akta', '$agama', '$kwn', '$butuh_khs', '$alamat', '$rt', '$rw', '$dusun', '$kelurahan', '$kec', '$kdpos', '$lintang', '$bujur', '$tpt_tinggal', '$moda_trans', '$no_kks', '$anak_ke', '$ada_kps', '$no_kps')";
								    $query = mysqli_query($koneksi, $sql);
								    // apakah query simpan berhasil?
								    if( $query ) {
								        // kalau berhasil reload dengan message alert!!
								        echo "<script>window.alert('Selamat, Anda telah berhasil mendaftar!!');</script>";
								    } else {
								        // kalau gagal reload dengan message alert!!
								        echo "<script>window.alert('Maaf, Anda gagal mendaftar!!');</script>";
										    }
							
							 
							//membuat instance spreadsheet baru
							$spreadsheet = new Spreadsheet();
							//memilih sheet yang sedang aktif
							$sheet = $spreadsheet->getActiveSheet();
							//membuat header kolom A B C dan D
							$sheet->setCellValue('A1', 'No');
							$sheet->setCellValue('B1', 'Nomer Pendaftaran');
							$sheet->setCellValue('C1', 'Nama');
							$sheet->setCellValue('D1', 'Jenis Kelamin');
							$sheet->setCellValue('E1', 'NISN');
							$sheet->setCellValue('F1', 'NIK');
							$sheet->setCellValue('G1', 'Tempat Kelahiran');
							$sheet->setCellValue('H1', 'Tanggal Lahir');
							$sheet->setCellValue('I1', 'No. Akta Kelahiran');
							$sheet->setCellValue('J1', 'Agama');
							$sheet->setCellValue('K1', 'Kewarganegaraan');
							$sheet->setCellValue('L1', 'Kebutuhan Khusus');
							$sheet->setCellValue('M1', 'Alamat');
							$sheet->setCellValue('N1', 'RT');
							$sheet->setCellValue('O1', 'RW');
							$sheet->setCellValue('P1', 'Dusun');
							$sheet->setCellValue('Q1', 'Kelurahan');
							$sheet->setCellValue('R1', 'Kecamatan');
							$sheet->setCellValue('S1', 'Kode Pos');
							$sheet->setCellValue('T1', 'Lintang');
							$sheet->setCellValue('U1', 'Bujur');
							$sheet->setCellValue('V1', 'Tempat Tinggal');
							$sheet->setCellValue('W1', 'Moda Transportasi');
							$sheet->setCellValue('X1', 'No. KKS');
							$sheet->setCellValue('Y1', 'Anak ke');
							$sheet->setCellValue('Z1', 'Ada KPS ?');
							$sheet->setCellValue('AA1', 'No. KPS/PKH');
							//pemanggilan query select 
							$query = mysqli_query($koneksi,"select * from pendaftaran_siswa");
							//agar tabel yang ditulis oleh pemanggilan query dimulai dari baris kedua
							$i = 2;
							//variabel bantuan nomor urut dimulai pada kolom A2
							$no = 1;
							$bulan  = array (1 => 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember');
							
							//looping 
							while($row = mysqli_fetch_array($query))
							{
								//mengatur agar tanggal dapat tertulis sesuai standar indonesia dipisahkan dengan tanda pisah per indeks.
								$tgl  = explode('-', $row['tgl_lhr']);
								//memasukkan isi dari database ke dalam file excel sesuai tabel yang ditentukan database
								$sheet->setCellValue('A'.$i, $no++);
								$sheet->setCellValue('B'.$i, $row[0]);
								$sheet->setCellValue('C'.$i, $row[1]);
								$sheet->setCellValue('D'.$i, $row[2]);
								$sheet->setCellValue('E'.$i, $row[3]);
								$sheet->setCellValue('F'.$i, $row[4]);
								$sheet->setCellValue('G'.$i, $row[5]);
								$sheet->setCellValue('H'.$i, $tgl[2] . ' ' . $bulan[(int)$tgl[1]] . ' '.$tgl[0]);
								$sheet->setCellValue('I'.$i, $row[7]);
								$sheet->setCellValue('J'.$i, $row[8]);
								$sheet->setCellValue('K'.$i, $row[9]);
								$sheet->setCellValue('L'.$i, $row[10]);
								$sheet->setCellValue('M'.$i, $row[11]);
								$sheet->setCellValue('N'.$i, $row[12]);
								$sheet->setCellValue('O'.$i, $row[13]);
								$sheet->setCellValue('P'.$i, $row[14]);
								$sheet->setCellValue('Q'.$i, $row[15]);
								$sheet->setCellValue('R'.$i, $row[16]);
								$sheet->setCellValue('S'.$i, $row[17]);
								$sheet->setCellValue('T'.$i, $row[18]);
								$sheet->setCellValue('U'.$i, $row[19]);
								$sheet->setCellValue('V'.$i, $row[20]);
								$sheet->setCellValue('W'.$i, $row[21]);
								$sheet->setCellValue('X'.$i, $row[22]);
								$sheet->setCellValue('Y'.$i, $row[23]);
								$sheet->setCellValue('Z'.$i, $row[24]);
								$sheet->setCellValue('AA'.$i, $row[25]);

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
							$sheet->getStyle('A1:AA'.$i)->applyFromArray($styleArray);
							 
							//membuat instance baru yang berguna untuk menyimpan ke dalam file excel 
							$writer = new Xlsx($spreadsheet);
							//menyimpan ke dalam file excel sesuai nama
							$writer->save('Report Data Pendaftaran Siswa.xlsx');
							//memunculkan window alert berupa pesan berhasil export di lokasi direktori
							echo "<script>window.alert('Data berhasil di export ke C:/xampp/htdocs/reportexcel/tugasprak11/Report Data Pendaftaran Siswa.xlsx');</script>"; 
													
							}
						}
	function cek_input($data){
		$data = trim($data);
		$data = stripcslashes($data);
		$data = htmlspecialchars($data);
		return $data;
	}
	?>
<style>
	.warning { color: #FF0000; }
</style>
</head>
<body>
	<div class="row">
		<div class="col-md-11">
			<div class="card-header">
				Pendaftaran Siswa Baru SMAN 1 Sampang
			</div>
			<div class="card-body">
				<form action="<?php echo htmlspecialchars($_SERVER["PHP_SELF"]);?>" method="POST">
					<div class="form-group row">
						<label for="nama" class="col-sm-2 col-form-label">Nama Lengkap</label>
						<div class="col-sm-10">
							<input type="text" name="nama" class="form-control <?php echo ($error_nama !="" ? "is-invalid" : "");?>" id="nama" maxlength="52" placeholder="Nama Lengkap" value="<?php echo $nama; ?>"><span class="warning"><?php echo $error_nama; ?></span>
						</div>
						</div>
					<div class="form-group row">
						<label for="nama" class="col-sm-2 col-form-label">Jenis Kelamin</label>
						<div class="col-sm-2 form-check-label">
							<input type="radio" name="jk" id="jk" <?php if (isset($jk) && $jk == "Laki-Laki") echo "checked";?> value="Laki-Laki"> Laki-Laki
						</div>
						<div class="col-sm-2 form-check-label">
							<input type="radio" name="jk" id="jk" <?php if (isset($jk) && $jk == "Perempuan") echo "checked";?> value="Perempuan"> Perempuan
						</div>
					</div>
					<div class="form-group row">
						<label for="nisn" class="col-sm-2 col-form-label">NISN</label>
						<div class="col-sm-10">
							<input type="text" name="nisn" maxlength="10" class="form-control <?php echo ($error_nisn !="" ? "is-invalid" : ""); ?>" id="nisn" placeholder="Silahkan isi NISN Anda dengan sesuai" value="<?php echo $nisn; ?>"><span class="warning"><?php echo $error_nisn; ?></span>
						</div>
					</div>
					
					<div class="form-group row">
						<label for="nik" class="col-sm-2 col-form-label">NIK</label>
						<div class="col-sm-10">
							<input type="text" name="nik" maxlength="16" class="form-control <?php echo ($error_nik !="" ? "is-invalid" : ""); ?>" id="nik" placeholder="Silahkan isi NIK Anda dengan sesuai" value="<?php echo $nik; ?>"><span class="warning"><?php echo $error_nik; ?></span>
						</div>
					</div>

					<div class="form-group row">
						<label for="tpt_lahir" class="col-sm-2 col-form-label">Tempat Lahir</label>
						<div class="col-sm-10">
							<input type="text" name="tpt_lahir" maxlength="50" class="form-control <?php echo ($error_tpt_lahir !="" ? "is-invalid" : ""); ?>" id="tpt_lahir" placeholder="Silahkan isi Tempat Lahir Anda sesuai dengan Akta Kelahiran Anda" value="<?php echo $tpt_lahir; ?>"><span class="warning"><?php echo $error_tpt_lahir; ?></span>
						</div>
					</div>
					
					<div class="form-group row">
						<label for="tgl_lhr" class="col-sm-2 col-form-label">Tanggal Lahir</label>
						<div class="col-sm-10">
							<input type="date" name="tgl_lhr" class="form-control <?php echo ($error_tgl_lhr != "" ? "is-invalid" : ""); ?>" id="tgl_lhr" value="<?php echo $tgl_lhr; ?>"><span class="warning"><?php echo $error_tgl_lhr; ?></span>
						</div>
					</div>
					
					<div class="form-group row">
						<label for="no_akta" class="col-sm-2 col-form-label">No. Registrasi Akta</label>
						<div class="col-sm-10">
							<input type="text" name="no_akta" maxlength="50" class="form-control <?php echo ($error_no_akta !="" ? "is-invalid" : ""); ?>" id="no_akta" placeholder="Silahkan isi No. Registrasi Akta Kelahiran Anda dengan sesuai" value="<?php echo $no_akta; ?>"><span class="warning"><?php echo $error_no_akta; ?></span>
						</div>
					</div>

					<div class="form-group row">
						<label for="agama" class="col-form-label col-sm-2">Agama</label>
						<select name="agama" class="col-sm-8 form-control" >
								<option value="Islam">Islam</option>
				                <option value="Kristen Katolik">Kristen Katolik</option>
				                <option value="Kristen Protestan">Kristen Protestan</option>
				                <option value="Hindu">Hindu</option>
				                <option value="Budha">Budha</option>
				                <option value="Konghuchu">Konghuchu</option>
				                <option value="Kepercayaan terhadap Tuhan YME">Kepercayaan terhadap Tuhan YME</option>
				                <option value="Lainnya">Lainnya</option>
				            </select>    
					</div>
					<div class="form-group row">
						<label for="kwn" class="col-form-label col-sm-2">Kewarganegaraan</label>
						<div class="col-sm-2 form-check-label">
							<input type="radio" name="kwn" id="kwn" <?php if (isset($kwn) && $kwn == "Indonesia (WNI)") echo "checked";?>  value="Indonesia (WNI)">Indonesia (WNI)
						</div>
						<div class="col-sm-2 form-check-label">
							<input type="radio" name="kwn" id="kwn" <?php if (isset($kwn) && $kwn == "Asing (WNA)") echo "checked";?> value="Asing (WNA)">Asing (WNA)
						</div> <span class="warning"><?php echo $error_kwn ?></span>
					</div>

					<div class="form-group row">
						<label for="butuh_khs" class="col-form-label col-sm-2">Kebutuhan Khusus</label>
						<select name="butuh_khs" class="col-sm-8 form-control" >
								<option value="Tidak">Tidak</option>
				                <option value="Netra">Netra</option>
				                <option value="Rungu">Rungu</option>
				                <option value="Grahita ringan">Grahita ringan</option>
				                <option value="Grahita sedang">Grahita sedang</option>
				                <option value="Daksa ringan">Daksa ringan</option>
				                <option value="Daksa sedang">Daksa sedang</option>
				                <option value="Laras">Laras</option>
				                <option value="Wicara">Wicara</option>
				                <option value="Tuna ganda">Tuna ganda</option>
				                <option value="Hiperaktif">Hiperaktif</option>
				                <option value="Cerdas Istimewa">Cerdas Istimewa</option>
				                <option value="Bakat Istimewa">Bakat Istimewa</option>
				                <option value="Kesulitan Belajar">Kesulitan Belajar</option>
				                <option value="Narkoba">Narkoba</option>
				                <option value="Indigo">Indigo</option>
				                <option value="Down Sindrome">Down Sindrome</option>
				                <option value="Autis">Autis</option>
				            </select>    
					</div>

					<div class="form-group row">
						<label for="alamat" class="col-sm-2 col-form-label">Alamat</label>
						<div class="col-sm-10">
							<input type="text" name="alamat" maxlength="60" class="form-control <?php echo ($error_alamat !="" ? "is-invalid" : ""); ?>" id="alamat" placeholder="Silahkan isi alamat Anda dengan sesuai" value="<?php echo $no_akta; ?>"><span>Contoh : Jln. Selong Permai II/C-20</span><span class="warning"><?php echo $error_alamat; ?></span>
						</div>
					</div>

					<div class="form-group row">
						<label for="rt" class="col-sm-2 col-form-label">RT</label>
						<div class="col-sm-10">
							<input type="text" name="rt" maxlength="3" class="form-control <?php echo ($error_rt !="" ? "is-invalid" : ""); ?>" id="rt" placeholder="Silahkan isi RT Anda dengan sesuai" value="<?php echo $rt; ?>"><span class="warning"><?php echo $error_rt; ?></span>
						</div>
					</div>


					<div class="form-group row">
						<label for="rw" class="col-sm-2 col-form-label">RW</label>
						<div class="col-sm-10">
							<input type="text" name="rw" maxlength="3" class="form-control <?php echo ($error_rw !="" ? "is-invalid" : ""); ?>" id="rw" placeholder="Silahkan isi RW Anda dengan sesuai" value="<?php echo $rw; ?>"><span class="warning"><?php echo $error_rw; ?></span>
						</div>
					</div>

					<div class="form-group row">
						<label for="dusun" class="col-sm-2 col-form-label">Dusun</label>
						<div class="col-sm-10">
							<input type="text" name="dusun" maxlength="31" class="form-control <?php echo ($error_dusun !="" ? "is-invalid" : ""); ?>" id="dusun" placeholder="Silahkan isi Dusun Anda dengan sesuai" value="<?php echo $dusun; ?>"><span class="warning"><?php echo $error_dusun; ?></span>
						</div>
					</div>


					<div class="form-group row">
						<label for="kelurahan" class="col-sm-2 col-form-label">Kelurahan</label>
						<div class="col-sm-10">
							<input type="text" name="kelurahan" maxlength="31" class="form-control <?php echo ($error_kelurahan !="" ? "is-invalid" : ""); ?>" id="kelurahan" placeholder="Silahkan isi Kelurahan Anda dengan sesuai" value="<?php echo $kelurahan; ?>"><span class="warning"><?php echo $error_kelurahan; ?></span>
						</div>
					</div>

					<div class="form-group row">
						<label for="kec" class="col-sm-2 col-form-label">Kecamatan</label>
						<div class="col-sm-10">
							<input type="text" name="kec" maxlength="31" class="form-control <?php echo ($error_kec !="" ? "is-invalid" : ""); ?>" id="kec" placeholder="Silahkan isi Kecamatan Anda dengan sesuai" value="<?php echo $kec; ?>"><span class="warning"><?php echo $error_kec; ?></span>
						</div>
					</div>

					<div class="form-group row">
						<label for="kdpos" class="col-sm-2 col-form-label">Kode Pos</label>
						<div class="col-sm-10">
							<input type="text" name="kdpos" maxlength="5" class="form-control <?php echo ($error_kdpos !="" ? "is-invalid" : ""); ?>" id="kdpos" placeholder="Silahkan isi Kode Pos Anda dengan sesuai" value="<?php echo $kdpos; ?>"><span class="warning"><?php echo $error_kdpos; ?></span>
						</div>
					</div>

					<div class="form-group row">
						<label for="lintang" class="col-sm-2 col-form-label">Lintang</label>
						<div class="col-sm-10">
							<input type="text" name="lintang" maxlength="31" class="form-control <?php echo ($error_lintang !="" ? "is-invalid" : ""); ?>" id="lintang" placeholder="Silahkan isi Koordinat Lintang Anda dengan sesuai" value="<?php echo $lintang; ?>"><span class="warning"><?php echo $error_lintang; ?></span>
						</div>
					</div>

					<div class="form-group row">
						<label for="bujur" class="col-sm-2 col-form-label">Bujur</label>
						<div class="col-sm-10">
							<input type="text" name="bujur" maxlength="31" class="form-control <?php echo ($error_bujur !="" ? "is-invalid" : ""); ?>" id="bujur" placeholder="Silahkan isi Koordinat Bujur Anda dengan sesuai" value="<?php echo $bujur; ?>"><span class="warning"><?php echo $error_bujur; ?></span>
						</div>
					</div>

					<div class="form-group row">
						<label for="tpt_tinggal" class="col-sm-2 col-form-label">Tempat Tinggal</label>
						<select name="tpt_tinggal" class="col-sm-8 form-control" >
								<option value="Bersama Orang Tua">Bersama Orang Tua</option>
				                <option value="Wali">Wali</option>
				                <option value="Kos">Kos</option>
				                <option value="Asrama">Asrama</option>
				                <option value="Panti Asuhan">Panti Asuhan</option>
				                <option value="Lainnya">Lainnya</option>
				        </select>    
					</div>

					<div class="form-group row">
						<label for="moda_trans" class="col-form-label col-sm-2">Moda Transportasi</label>
						<select name="moda_trans" class="col-sm-8 form-control">
								<option value=Jalan Kaki">Jalan Kaki</option>
				                <option value="Kendaraan Pribadi">Kendaraan Pribadi</option>
				                <option value="Kendaraab Umum">Kendaraan Umum</option>
				                <option value="Jemputan Sekolah">Jemputan Sekolah</option>
				                <option value="Kereta Api">Kereta Api</option>
				                <option value="Ojek">Ojek</option>
				                <option value="Andong/Bendi/Sado/Dokar">Andong/Bendi/Sado/Dokar</option>
				                <option value="Perahu Penyebrangan">Perahu Penyebrangan</option>
				                <option value="Lainnya">Lainnya</option>
				        </select>    
					</div>



					<div class="form-group row">
						<label for="no_kks" class="col-sm-2 col-form-label">No. KKS</label>
						<div class="col-sm-10">
							<input type="text" name="no_kks" maxlength="8" class="form-control <?php echo ($error_no_kks !="" ? "is-invalid" : ""); ?>" id="no_kks" placeholder="Silahkan isi No. KKS Anda dengan sesuai dengan kartu" value="<?php echo $no_kks; ?>"><span class="warning"><?php echo $error_no_kks; ?></span>
						</div>
					</div>

					<div class="form-group row">
						<label for="anak_ke" class="col-sm-2 col-form-label">Anak ke</label>
						<div class="col-sm-10">
							<input type="text" name="anak_ke" maxlength="2" class="form-control <?php echo ($error_anak !="" ? "is-invalid" : ""); ?>" id="anak_ke" placeholder="Anda anak ke berapa dari bersaudara" value="<?php echo $anak_ke; ?>"><span class="warning"><?php echo $error_anak; ?></span>
						</div>
					</div>

					<div class="form-group row">
						<label for="ada_kps" class="col-sm-2 col-form-label">Anda Penerima KPS/PKH?</label>
						<div class="col-sm-2 form-check-label">
							<input type="radio" name="ada_kps" id="ada_kps" <?php if (isset($ada_kps) && $ada_kps == "Ya") echo "checked";?>  value="Ya"> Ya
						</div>
						<div class="col-sm-2 form-check-label">
							<input type="radio" name="ada_kps" id="ada_kps" <?php if (isset($ada_kps) && $ada_kps == "Tidak") echo "checked";?> value="Tidak"> Tidak
						</div><span class="warning"><?php echo $error_kps ?></span>
					</div>

					<div class="form-group row">
						<label for="no_kps" class="col-sm-2 col-form-label">No. KPS/PKH</label>
						<div class="col-sm-10">
							<input type="text" name="no_kps" maxlength="14" class="form-control <?php echo ($error_no_kps !="" ? "is-invalid" : ""); ?>" id="no_kps" placeholder="Silahkan masukkan no. KPS yang tertera pada kartu Anda" value="<?php echo $no_kps; ?>"><span class="warning"><?php echo $error_no_kps; ?></span>
						</div>
					</div>

					<div class="form-group row">
						<div class="col-sm-10">
							<button type="submit" class="btn btn-primary" name="simpan">Daftarkan</button>
						</div>
					</div>

					</form>
					</div>							
				</div>					
			</div>
	</body>
</html>