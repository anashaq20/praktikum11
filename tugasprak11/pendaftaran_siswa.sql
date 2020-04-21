-- phpMyAdmin SQL Dump
-- version 5.0.2
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Waktu pembuatan: 21 Apr 2020 pada 10.20
-- Versi server: 10.4.11-MariaDB
-- Versi PHP: 7.4.4

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `db_siswa`
--

-- --------------------------------------------------------

--
-- Struktur dari tabel `pendaftaran_siswa`
--

CREATE TABLE `pendaftaran_siswa` (
  `id_pendaftar` int(11) NOT NULL,
  `nama` varchar(52) NOT NULL,
  `jk` varchar(9) NOT NULL,
  `nisn` varchar(10) NOT NULL,
  `nik` varchar(16) NOT NULL,
  `tpt_lahir` varchar(50) NOT NULL,
  `tgl_lhr` date NOT NULL,
  `no_akta` varchar(50) NOT NULL,
  `agama` varchar(30) NOT NULL,
  `kwn` varchar(50) NOT NULL,
  `butuh_khs` varchar(40) NOT NULL,
  `alamat` varchar(60) NOT NULL,
  `rt` varchar(3) NOT NULL,
  `rw` varchar(3) NOT NULL,
  `dusun` varchar(31) NOT NULL,
  `kelurahan` varchar(31) NOT NULL,
  `kecamatan` varchar(31) NOT NULL,
  `kdpos` varchar(5) NOT NULL,
  `lintang` varchar(31) NOT NULL,
  `bujur` varchar(31) NOT NULL,
  `tpt_tinggal` varchar(30) NOT NULL,
  `moda_trans` varchar(30) NOT NULL,
  `no_kks` varchar(8) NOT NULL,
  `anak_ke` varchar(2) NOT NULL,
  `ada_kps` varchar(5) NOT NULL,
  `no_kps_pkh` varchar(14) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `pendaftaran_siswa`
--

INSERT INTO `pendaftaran_siswa` (`id_pendaftar`, `nama`, `jk`, `nisn`, `nik`, `tpt_lahir`, `tgl_lhr`, `no_akta`, `agama`, `kwn`, `butuh_khs`, `alamat`, `rt`, `rw`, `dusun`, `kelurahan`, `kecamatan`, `kdpos`, `lintang`, `bujur`, `tpt_tinggal`, `moda_trans`, `no_kks`, `anak_ke`, `ada_kps`, `no_kps_pkh`) VALUES
(1, 'Ahmad Nashirul Haq', 'Laki-Laki', '0004096817', '3527031908000007', 'Sampang', '2000-08-19', '1231241212312512312123123124', 'Islam', 'Indonesia (WNI)', 'Tidak', 'Jalan Jaksa Agung Suprapto No 95', '002', '008', 'Gunung Sekar', 'Gunung Sekar', 'Sampang', '69213', '144 LU 114 LS', '100 BB 100 BT', 'Bersama Orang Tua', 'Jalan', '90180293', '1', 'Ya', '12312412312123'),
(2, 'Ahmad Nashirul Haq', 'Laki-Laki', '0004096817', '3527031908000007', 'Sampang', '2000-08-19', '1231241212312512312123123124', 'Islam', 'Indonesia (WNI)', 'Tidak', '1231241212312512312123123124', '002', '008', 'Gunung Sekar', 'Gunung Sekar', 'Sampang', '69213', '144 LU 114 LS', '100 BB 100 BT', 'Bersama Orang Tua', 'Jalan', '90180293', '1', 'Ya', '12312412312123'),
(3, 'Ahmad Nashirul Haq', 'Laki-Laki', '0004096817', '3527031908000007', 'Sampang', '2000-08-19', '1231241212312512312123123124', 'Islam', 'Indonesia (WNI)', 'Tidak', '1231241212312512312123123124', '002', '008', 'Gunung Sekar', 'Gunung Sekar', 'Sampang', '69213', '144 LU 114 LS', '100 BB 100 BT', 'Bersama Orang Tua', 'Jalan', '90180293', '1', 'Ya', '12312412312123');

--
-- Indexes for dumped tables
--

--
-- Indeks untuk tabel `pendaftaran_siswa`
--
ALTER TABLE `pendaftaran_siswa`
  ADD PRIMARY KEY (`id_pendaftar`);

--
-- AUTO_INCREMENT untuk tabel yang dibuang
--

--
-- AUTO_INCREMENT untuk tabel `pendaftaran_siswa`
--
ALTER TABLE `pendaftaran_siswa`
  MODIFY `id_pendaftar` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=4;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
