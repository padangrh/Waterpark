# SQL Manager Lite for MySQL 5.5.1.45563
# ---------------------------------------
# Host     : localhost
# Port     : 3306
# Database : waterpark


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES latin1 */;

SET FOREIGN_KEY_CHECKS=0;

CREATE DATABASE `waterpark`
    CHARACTER SET 'latin1'
    COLLATE 'latin1_swedish_ci';

USE `waterpark`;

#
# Structure for the `bill` table : 
#

CREATE TABLE `bill` (
  `nobukti` CHAR(20) COLLATE utf8_general_ci NOT NULL,
  `kasir` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL,
  `tanggal` DATE DEFAULT NULL,
  `jam` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL,
  `jumlah` INTEGER(20) DEFAULT NULL,
  `total` INTEGER(20) DEFAULT NULL,
  `bayar` INTEGER(20) DEFAULT NULL,
  `cash` TINYINT(2) DEFAULT NULL,
  `diskon` INTEGER(20) DEFAULT 0,
  PRIMARY KEY (`nobukti`) USING BTREE
) ENGINE=InnoDB
CHARACTER SET 'utf8' COLLATE 'utf8_general_ci'
;

#
# Structure for the `bill_beli` table : 
#

CREATE TABLE `bill_beli` (
  `nobukti` CHAR(20) COLLATE utf8_general_ci NOT NULL,
  `staff` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL,
  `tanggal` DATE DEFAULT NULL,
  `jam` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL,
  `total` INTEGER(11) DEFAULT NULL,
  `kode_supplier` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL,
  `pembayaran` TINYINT(4) DEFAULT NULL,
  `lunas` TINYINT(2) DEFAULT NULL,
  `settled` TINYINT(2) DEFAULT NULL,
  `tanggal_lunas` DATE DEFAULT NULL,
  PRIMARY KEY (`nobukti`) USING BTREE
) ENGINE=InnoDB
CHARACTER SET 'utf8' COLLATE 'utf8_general_ci'
;

#
# Structure for the `tbaktif` table : 
#

CREATE TABLE `tbaktif` (
  `rfid` CHAR(20) COLLATE latin1_swedish_ci NOT NULL,
  `tanggal` DATE DEFAULT NULL,
  `jam` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL,
  `status` TINYINT(4) DEFAULT NULL,
  `keterangan` CHAR(100) COLLATE latin1_swedish_ci DEFAULT NULL,
  PRIMARY KEY (`rfid`) USING BTREE
) ENGINE=InnoDB
CHARACTER SET 'latin1' COLLATE 'latin1_swedish_ci'
;

#
# Structure for the `tbbarang` table : 
#

CREATE TABLE `tbbarang` (
  `kode` CHAR(20) COLLATE utf8mb3_general_ci NOT NULL,
  `nama` VARCHAR(50) COLLATE utf8mb3_general_ci NOT NULL,
  `kategori` VARCHAR(20) COLLATE utf8mb3_general_ci NOT NULL,
  `harga_jual` DOUBLE(15,5) NOT NULL,
  `kdsuplier` CHAR(20) COLLATE utf8mb3_general_ci NOT NULL,
  PRIMARY KEY (`kode`, `kategori`, `kdsuplier`) USING BTREE,
  UNIQUE KEY `kode_test` (`kode`) USING BTREE
) ENGINE=MyISAM
CHARACTER SET 'utf8mb3' COLLATE 'utf8mb3_general_ci'
;

#
# Structure for the `tbbeli` table : 
#

CREATE TABLE `tbbeli` (
  `nobukti` CHAR(20) COLLATE latin1_swedish_ci NOT NULL,
  `tglbukti` DATE DEFAULT NULL,
  `kode` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL,
  `nama_barang` CHAR(70) COLLATE latin1_swedish_ci DEFAULT '',
  `harga` DOUBLE(15,2) DEFAULT 0.00,
  `jumlah` INTEGER(5) DEFAULT 0,
  `return` INTEGER(11) DEFAULT 0,
  KEY `nobukti` (`nobukti`, `tglbukti`, `kode`) USING BTREE
) ENGINE=MyISAM
ROW_FORMAT=FIXED CHARACTER SET 'latin1' COLLATE 'latin1_swedish_ci'
;

#
# Structure for the `tbdeposit` table : 
#

CREATE TABLE `tbdeposit` (
  `nodeposit` CHAR(20) COLLATE latin1_swedish_ci NOT NULL,
  `kasir` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL,
  `tanggal` DATE DEFAULT NULL,
  `jam` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL,
  `deposit` INTEGER(11) DEFAULT NULL,
  PRIMARY KEY (`nodeposit`) USING BTREE
) ENGINE=InnoDB
CHARACTER SET 'latin1' COLLATE 'latin1_swedish_ci'
;

#
# Structure for the `tbdiskon` table : 
#

CREATE TABLE `tbdiskon` (
  `nobukti` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL,
  `supervisor` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL,
  `status` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL,
  `customer` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL,
  `nilai` INTEGER(11) DEFAULT NULL
) ENGINE=InnoDB
CHARACTER SET 'utf8' COLLATE 'utf8_general_ci'
;

#
# Structure for the `tbjual` table : 
#

CREATE TABLE `tbjual` (
  `nobukti` CHAR(20) COLLATE latin1_swedish_ci NOT NULL,
  `tglbukti` DATE DEFAULT '0000-00-00',
  `kode` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL,
  `nama_barang` CHAR(50) COLLATE latin1_swedish_ci DEFAULT '',
  `harga_jual` DOUBLE(15,2) DEFAULT 0.00,
  `jumlah_jual` INTEGER(11) DEFAULT 0,
  `kdsuplier` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL,
  KEY `NOBUKTI` (`nobukti`, `tglbukti`, `kode`) USING BTREE,
  KEY `nobukti_2` (`nobukti`) USING BTREE
) ENGINE=MyISAM
ROW_FORMAT=FIXED CHARACTER SET 'latin1' COLLATE 'latin1_swedish_ci'
;

#
# Structure for the `tbkategori` table : 
#

CREATE TABLE `tbkategori` (
  `kode` VARCHAR(20) COLLATE utf8mb3_general_ci NOT NULL,
  PRIMARY KEY (`kode`) USING BTREE
) ENGINE=MyISAM
CHARACTER SET 'utf8mb3' COLLATE 'utf8mb3_general_ci'
;

#
# Structure for the `tblogin` table : 
#

CREATE TABLE `tblogin` (
  `userid` VARCHAR(25) COLLATE utf8mb3_general_ci NOT NULL,
  `pass` VARCHAR(25) COLLATE utf8mb3_general_ci NOT NULL,
  `posisi` VARCHAR(30) COLLATE utf8mb3_general_ci DEFAULT NULL,
  `hak1` VARCHAR(1) COLLATE utf8mb3_general_ci DEFAULT NULL,
  `hak2` VARCHAR(1) COLLATE utf8mb3_general_ci DEFAULT NULL,
  `hak3` VARCHAR(1) COLLATE utf8mb3_general_ci DEFAULT NULL,
  `hak4` VARCHAR(1) COLLATE utf8mb3_general_ci DEFAULT NULL,
  PRIMARY KEY (`userid`) USING BTREE
) ENGINE=MyISAM
CHARACTER SET 'utf8mb3' COLLATE 'utf8mb3_general_ci'
;

#
# Structure for the `tbnonaktif` table : 
#

CREATE TABLE `tbnonaktif` (
  `rfid` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL,
  `tanggal` DATE DEFAULT NULL,
  `jam` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL,
  `status` TINYINT(4) DEFAULT NULL,
  `keterangan` CHAR(200) COLLATE latin1_swedish_ci DEFAULT NULL,
  `userid` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL
) ENGINE=InnoDB
CHARACTER SET 'latin1' COLLATE 'latin1_swedish_ci'
;

#
# Structure for the `tbreader` table : 
#

CREATE TABLE `tbreader` (
  `id` INTEGER(11) NOT NULL AUTO_INCREMENT,
  `rfid` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL,
  PRIMARY KEY (`id`) USING BTREE
) ENGINE=InnoDB
AUTO_INCREMENT=11 CHARACTER SET 'latin1' COLLATE 'latin1_swedish_ci'
;

#
# Structure for the `tbrfid` table : 
#

CREATE TABLE `tbrfid` (
  `nobukti` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL,
  `rfid` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL
) ENGINE=InnoDB
CHARACTER SET 'latin1' COLLATE 'latin1_swedish_ci'
;

#
# Structure for the `tbrfiddeposit` table : 
#

CREATE TABLE `tbrfiddeposit` (
  `nodeposit` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL,
  `rfid` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL,
  `hargarfid` INTEGER(11) DEFAULT NULL
) ENGINE=InnoDB
CHARACTER SET 'latin1' COLLATE 'latin1_swedish_ci'
;

#
# Structure for the `tbsuplier` table : 
#

CREATE TABLE `tbsuplier` (
  `kdsuplier` CHAR(20) COLLATE latin1_swedish_ci NOT NULL,
  `nmsuplier` VARCHAR(50) COLLATE latin1_swedish_ci DEFAULT NULL,
  `alamat` VARCHAR(250) COLLATE latin1_swedish_ci DEFAULT NULL,
  `telp` VARCHAR(40) COLLATE latin1_swedish_ci DEFAULT NULL,
  `tgl_gabung` DATE NOT NULL,
  `nama_rek` CHAR(50) COLLATE latin1_swedish_ci DEFAULT NULL,
  `no_rek` CHAR(30) COLLATE latin1_swedish_ci DEFAULT NULL,
  `bank` CHAR(20) COLLATE latin1_swedish_ci DEFAULT NULL,
  PRIMARY KEY (`kdsuplier`) USING BTREE
) ENGINE=MyISAM
CHARACTER SET 'latin1' COLLATE 'latin1_swedish_ci'
;

#
# Data for the `bill` table  (LIMIT 0,500)
#

INSERT INTO `bill` (`nobukti`, `kasir`, `tanggal`, `jam`, `jumlah`, `total`, `bayar`, `cash`, `diskon`) VALUES
  ('W15202','admin','2017-06-17','10:23:36',150000,195000,200000,1,0),
  ('W15203','anton','2017-06-17','11:01:13',300000,390000,400000,1,0),
  ('W15204','admin','2017-06-20','13:10:35',500000,650000,700000,1,0),
  ('W15205','admin','2017-06-20','13:10:59',400000,520000,0,0,0),
  ('W15206','admin','2017-06-20','13:11:23',400000,520000,0,0,0),
  ('W15207','admin','2017-06-21','11:46:06',150000,195000,0,0,0),
  ('W15208','admin','2017-06-21','11:46:50',150000,195000,0,0,0),
  ('W15209','admin','2017-06-21','11:58:18',150000,195000,0,0,0),
  ('W15210','admin','2017-06-21','11:58:36',150000,195000,0,0,0),
  ('W15211','admin','2017-06-21','11:59:07',200000,260000,0,0,0),
  ('W15212','admin','2017-06-21','11:59:19',50000,65000,0,0,0),
  ('W15213','admin','2017-06-21','12:00:58',150000,195000,0,0,0),
  ('W15214','admin','2017-06-21','13:13:13',400000,520000,0,0,0),
  ('W15215','admin','2017-06-21','13:18:08',400000,520000,0,0,0),
  ('W15216','admin','2017-06-21','13:25:40',100000,130000,0,0,0),
  ('W15217','admin','2017-06-21','13:38:03',150000,195000,0,0,0),
  ('W15218','admin','2017-06-23','13:11:19',100000,130000,0,0,0),
  ('W15219','admin','2017-06-23','13:13:51',100000,130000,0,0,0),
  ('W15220','admin','2017-06-23','13:21:07',150000,195000,200000,1,0),
  ('W15221','admin','2017-06-23','13:34:04',100000,130000,150000,1,0),
  ('W15222','admin','2017-06-23','13:34:34',100000,130000,0,0,0),
  ('W15223','admin','2017-06-23','13:52:17',50000,65000,0,0,0),
  ('W15224','admin','2017-06-23','13:52:56',50000,65000,0,0,0),
  ('W15225','admin','2017-06-23','13:55:31',50000,65000,0,0,0),
  ('W15226','admin','2017-06-23','13:59:20',50000,65000,0,0,0),
  ('W15227','admin','2017-06-23','14:01:58',150000,195000,0,0,0),
  ('W15228','admin','2017-06-23','14:05:13',50000,65000,0,0,0),
  ('W15229','admin','2017-06-23','16:12:14',100000,130000,0,0,0),
  ('W15230','admin','2017-06-23','16:16:03',200000,260000,0,0,0),
  ('W15231','admin','2017-06-23','16:17:50',200000,260000,0,0,0),
  ('W15232','admin','2017-06-23','16:23:02',200000,260000,0,0,0),
  ('W15233','admin','2017-06-23','16:24:41',200000,260000,0,0,0),
  ('W15234','admin','2017-06-23','16:25:33',200000,260000,0,0,0),
  ('W15235','admin','2017-06-24','11:21:39',50000,65000,0,0,0),
  ('W15236','admin','2017-06-24','12:11:05',150000,195000,0,0,0),
  ('W15237','admin','2017-06-24','12:14:05',100000,130000,0,0,0),
  ('W15238','admin','2017-06-24','12:15:22',100000,130000,0,0,0),
  ('W15239','admin','2017-06-24','12:17:32',50000,65000,0,0,0),
  ('W15240','admin','2017-06-24','12:19:15',50000,65000,0,0,0),
  ('W15241','admin','2017-06-24','12:19:43',150000,195000,0,0,0);
COMMIT;

#
# Data for the `tbaktif` table  (LIMIT 0,500)
#

INSERT INTO `tbaktif` (`rfid`, `tanggal`, `jam`, `status`, `keterangan`) VALUES
  ('0011145357','2017-06-23','13:34:33',0,'W15222'),
  ('0011145364','2017-06-23','13:34:33',0,'W15222'),
  ('0011145385','2017-06-23','13:34:04',0,'W15221'),
  ('0011145481','2017-06-23','13:31:35',0,'W15221'),
  ('0011145504','2017-06-23','13:21:06',0,'W15220'),
  ('0011146024','2017-06-23','13:10:31',0,'W15218'),
  ('0011146199','2017-06-23','13:11:19',0,'W15218'),
  ('0011221671','2017-06-23','13:13:51',0,'W15219'),
  ('0011221725','2017-06-23','13:13:51',0,'W15219'),
  ('0011221726','2017-06-23','13:21:07',0,'W15220'),
  ('0011366260','2017-06-23','13:21:07',0,'W15220'),
  ('0011366322','2017-06-24','12:19:42',1,'W15241'),
  ('0011559569','2017-06-24','12:19:43',1,'W15241'),
  ('0011559695','2017-06-24','12:19:43',1,'W15241'),
  ('0011559878','2017-06-24','12:11:04',1,'W15236'),
  ('0011626187','2017-06-24','12:14:04',1,'W15237'),
  ('0011651486','2017-06-24','12:14:05',1,'W15237'),
  ('0011739633','2017-06-24','12:15:22',1,'W15238'),
  ('0011771269','2017-06-24','12:15:22',1,'W15238'),
  ('0011787421','2017-06-24','12:17:32',1,'W15239'),
  ('0011899088','2017-06-24','12:19:15',1,'W15240');
COMMIT;

#
# Data for the `tbbarang` table  (LIMIT 0,500)
#

INSERT INTO `tbbarang` (`kode`, `nama`, `kategori`, `harga_jual`, `kdsuplier`) VALUES
  ('1','Tiket Waterpark','Tiket',50000.00000,'1'),
  ('2','Deposit','Deposit',15000.00000,'1'),
  ('3','Voucher Diskon 50%','Voucher',-25000.00000,'1');
COMMIT;

#
# Data for the `tbdeposit` table  (LIMIT 0,500)
#

INSERT INTO `tbdeposit` (`nodeposit`, `kasir`, `tanggal`, `jam`, `deposit`) VALUES
  ('D17033','admin','2017-06-17','10:25:29',45000),
  ('D17034','admin','2017-06-20','13:20:45',15000),
  ('D17035','admin','2017-06-21','12:00:32',120000);
COMMIT;

#
# Data for the `tbjual` table  (LIMIT 0,500)
#

INSERT INTO `tbjual` (`nobukti`, `tglbukti`, `kode`, `nama_barang`, `harga_jual`, `jumlah_jual`, `kdsuplier`) VALUES
  ('W15202','2017-06-17','1','Tiket Waterpark',50000.00,3,'1'),
  ('W15202','2017-06-17','2','Deposit',15000.00,3,'2'),
  ('W15203','2017-06-17','1','Tiket Waterpark',50000.00,6,'1'),
  ('W15203','2017-06-17','2','Deposit',15000.00,6,'2'),
  ('W15204','2017-06-20','1','Tiket Waterpark',50000.00,10,'1'),
  ('W15204','2017-06-20','2','Deposit',15000.00,10,'2'),
  ('W15205','2017-06-20','1','Tiket Waterpark',50000.00,8,'1'),
  ('W15205','2017-06-20','2','Deposit',15000.00,8,'2'),
  ('W15206','2017-06-20','1','Tiket Waterpark',50000.00,8,'1'),
  ('W15206','2017-06-20','2','Deposit',15000.00,8,'2'),
  ('W15207','2017-06-21','1','Tiket Waterpark',50000.00,3,'1'),
  ('W15207','2017-06-21','2','Deposit',15000.00,3,'2'),
  ('W15208','2017-06-21','1','Tiket Waterpark',50000.00,3,'1'),
  ('W15208','2017-06-21','2','Deposit',15000.00,3,'2'),
  ('W15209','2017-06-21','1','Tiket Waterpark',50000.00,3,'1'),
  ('W15209','2017-06-21','2','Deposit',15000.00,3,'2'),
  ('W15210','2017-06-21','1','Tiket Waterpark',50000.00,3,'1'),
  ('W15210','2017-06-21','2','Deposit',15000.00,3,'2'),
  ('W15211','2017-06-21','1','Tiket Waterpark',50000.00,4,'1'),
  ('W15211','2017-06-21','2','Deposit',15000.00,4,'2'),
  ('W15212','2017-06-21','1','Tiket Waterpark',50000.00,1,'1'),
  ('W15212','2017-06-21','2','Deposit',15000.00,1,'2'),
  ('W15213','2017-06-21','1','Tiket Waterpark',50000.00,3,'1'),
  ('W15213','2017-06-21','2','Deposit',15000.00,3,'2'),
  ('W15214','2017-06-21','1','Tiket Waterpark',50000.00,8,'1'),
  ('W15214','2017-06-21','2','Deposit',15000.00,8,'2'),
  ('W15215','2017-06-21','1','Tiket Waterpark',50000.00,8,'1'),
  ('W15215','2017-06-21','1','Tiket Waterpark',50000.00,8,'1'),
  ('W15215','2017-06-21','2','Deposit',15000.00,8,'2'),
  ('W15216','2017-06-21','1','Tiket Waterpark',50000.00,2,'1'),
  ('W15216','2017-06-21','2','Deposit',15000.00,2,'2'),
  ('W15217','2017-06-21','1','Tiket Waterpark',50000.00,3,'1'),
  ('W15217','2017-06-21','2','Deposit',15000.00,3,'2'),
  ('W15218','2017-06-23','1','Tiket Waterpark',50000.00,2,'1'),
  ('W15218','2017-06-23','2','Deposit',15000.00,2,'2'),
  ('W15219','2017-06-23','1','Tiket Waterpark',50000.00,2,'1'),
  ('W15219','2017-06-23','2','Deposit',15000.00,2,'2'),
  ('W15220','2017-06-23','1','Tiket Waterpark',50000.00,3,'1'),
  ('W15220','2017-06-23','2','Deposit',15000.00,3,'2'),
  ('W15221','2017-06-23','1','Tiket Waterpark',50000.00,2,'1'),
  ('W15221','2017-06-23','2','Deposit',15000.00,2,'2'),
  ('W15222','2017-06-23','1','Tiket Waterpark',50000.00,2,'1'),
  ('W15222','2017-06-23','2','Deposit',15000.00,2,'2'),
  ('W15223','2017-06-23','1','Tiket Waterpark',50000.00,1,'1'),
  ('W15223','2017-06-23','2','Deposit',15000.00,1,'2'),
  ('W15224','2017-06-23','1','Tiket Waterpark',50000.00,1,'1'),
  ('W15224','2017-06-23','2','Deposit',15000.00,1,'2'),
  ('W15225','2017-06-23','1','Tiket Waterpark',50000.00,1,'1'),
  ('W15225','2017-06-23','2','Deposit',15000.00,1,'2'),
  ('W15226','2017-06-23','1','Tiket Waterpark',50000.00,1,'1'),
  ('W15226','2017-06-23','2','Deposit',15000.00,1,'2'),
  ('W15227','2017-06-23','1','Tiket Waterpark',50000.00,3,'1'),
  ('W15227','2017-06-23','2','Deposit',15000.00,3,'2'),
  ('W15228','2017-06-23','1','Tiket Waterpark',50000.00,1,'1'),
  ('W15228','2017-06-23','2','Deposit',15000.00,1,'2'),
  ('W15229','2017-06-23','1','Tiket Waterpark',50000.00,2,'1'),
  ('W15229','2017-06-23','2','Deposit',15000.00,2,'2'),
  ('W15230','2017-06-23','1','Tiket Waterpark',50000.00,4,'1'),
  ('W15230','2017-06-23','2','Deposit',15000.00,4,'2'),
  ('W15231','2017-06-23','1','Tiket Waterpark',50000.00,4,'1'),
  ('W15231','2017-06-23','2','Deposit',15000.00,4,'2'),
  ('W15232','2017-06-23','1','Tiket Waterpark',50000.00,4,'1'),
  ('W15232','2017-06-23','2','Deposit',15000.00,4,'2'),
  ('W15233','2017-06-23','1','Tiket Waterpark',50000.00,4,'1'),
  ('W15233','2017-06-23','2','Deposit',15000.00,4,'2'),
  ('W15234','2017-06-23','1','Tiket Waterpark',50000.00,4,'1'),
  ('W15234','2017-06-23','2','Deposit',15000.00,4,'2'),
  ('W15235','2017-06-24','1','Tiket Waterpark',50000.00,1,'1'),
  ('W15235','2017-06-24','2','Deposit',15000.00,1,'2'),
  ('W15236','2017-06-24','1','Tiket Waterpark',50000.00,3,'1'),
  ('W15236','2017-06-24','2','Deposit',15000.00,3,'2'),
  ('W15237','2017-06-24','1','Tiket Waterpark',50000.00,2,'1'),
  ('W15237','2017-06-24','2','Deposit',15000.00,2,'2'),
  ('W15238','2017-06-24','1','Tiket Waterpark',50000.00,2,'1'),
  ('W15238','2017-06-24','2','Deposit',15000.00,2,'2'),
  ('W15239','2017-06-24','1','Tiket Waterpark',50000.00,1,'1'),
  ('W15239','2017-06-24','2','Deposit',15000.00,1,'2'),
  ('W15240','2017-06-24','1','Tiket Waterpark',50000.00,1,'1'),
  ('W15240','2017-06-24','2','Deposit',15000.00,1,'2'),
  ('W15241','2017-06-24','1','Tiket Waterpark',50000.00,3,'1'),
  ('W15241','2017-06-24','2','Deposit',15000.00,3,'2');
COMMIT;

#
# Data for the `tbkategori` table  (LIMIT 0,500)
#

INSERT INTO `tbkategori` (`kode`) VALUES
  ('Deposit'),
  ('Tiket'),
  ('Voucher');
COMMIT;

#
# Data for the `tblogin` table  (LIMIT 0,500)
#

INSERT INTO `tblogin` (`userid`, `pass`, `posisi`, `hak1`, `hak2`, `hak3`, `hak4`) VALUES
  ('admin','admin','Master','1','1','1','1'),
  ('anton','anton','Karyawan','1','0','0','0');
COMMIT;

#
# Data for the `tbnonaktif` table  (LIMIT 0,500)
#

INSERT INTO `tbnonaktif` (`rfid`, `tanggal`, `jam`, `status`, `keterangan`, `userid`) VALUES
  ('0011899088','2017-06-17','10:22:57',1,'W15202 - D17033','admin'),
  ('0011787421','2017-06-17','10:22:57',1,'W15202 - D17033','admin'),
  ('0011771269','2017-06-17','10:22:57',1,'W15202 - D17033','admin'),
  ('0011651486','2017-06-17','11:01:13',1,'W15203 - W15204 - Print','admin'),
  ('0011626187','2017-06-17','11:01:13',1,'W15203 - W15204 - Print','admin'),
  ('0011559878','2017-06-17','11:01:13',1,'W15203 - W15204 - Print','admin'),
  ('0011559695','2017-06-17','11:01:13',1,'W15203 - W15204 - Print','admin'),
  ('0011559569','2017-06-17','11:01:13',1,'W15203 - W15204 - Print','admin'),
  ('0011366322','2017-06-17','11:01:13',1,'W15203 - W15204 - Print','admin'),
  ('0011899088','2017-06-20','13:09:59',1,'W15204 - D17034','admin'),
  ('0011144960','2017-06-21','11:46:06',1,'W15207 - W15208 - Print','admin'),
  ('0011144962','2017-06-21','11:46:06',1,'W15207 - W15208 - Print','admin'),
  ('0011145060','2017-06-21','11:46:06',1,'W15207 - W15208 - Print','admin'),
  ('0011144960','2017-06-21','11:58:18',1,'W15209 - W15210 - Print','admin'),
  ('0011144962','2017-06-21','11:58:18',1,'W15209 - W15210 - Print','admin'),
  ('0011145060','2017-06-21','11:58:18',1,'W15209 - W15210 - Print','admin'),
  ('0011144960','2017-06-21','11:58:36',1,'W15210 - D17035','admin'),
  ('0011144962','2017-06-21','11:58:36',1,'W15210 - D17035','admin'),
  ('0011145060','2017-06-21','11:58:36',1,'W15210 - D17035','admin'),
  ('0011145357','2017-06-21','11:59:07',1,'W15211 - D17035','admin'),
  ('0011145364','2017-06-21','11:59:07',1,'W15211 - D17035','admin'),
  ('0011145385','2017-06-21','11:59:07',1,'W15211 - D17035','admin'),
  ('0011145481','2017-06-21','11:59:19',1,'W15212 - D17035','admin'),
  ('0011145281','2017-06-21','11:59:07',1,'W15211 - D17035','admin'),
  ('0011145060','2017-06-21','12:00:58',1,'W15213 - Deaktivasi - list RFID','admin'),
  ('0011145060','2017-06-21','12:00:58',0,'W15213 - Aktivasi - list RFID','admin'),
  ('0011145060','2017-06-21','12:01:27',1,'admin - perubahan - EditRFID','admin'),
  ('0011145060','2017-06-21','12:01:27',0,'admin - Aktivasi - list RFID','admin'),
  ('0011145060','2017-06-21','12:04:29',1,'admin - perubahan kartu - RFID','admin'),
  ('0011145481','2017-06-21','12:00:58',1,'W15213 - Deaktivasi - list RFID','admin'),
  ('0011145481','2017-06-21','12:00:58',0,'W15213 - perubahan - EditRFID','admin'),
  ('0011144960','2017-06-21','12:00:58',1,'W15213 - W15214 - Print','admin'),
  ('0011144962','2017-06-21','12:00:58',1,'W15213 - W15214 - Print','admin'),
  ('0011145481','2017-06-21','12:00:58',1,'admin - W15214 - Print','admin'),
  ('0011899088','2017-06-21','13:25:40',1,'W15216 - Deaktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:25:40',0,'W15216 - Aktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:26:54',1,'admin - Deaktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:26:54',0,'admin - Aktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:28:26',1,'admin - Deaktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:28:26',0,'admin - Aktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:28:46',1,'admin - Deaktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:28:46',0,'admin - Aktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:32:01',1,'admin - Deaktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:32:01',0,'admin - Aktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:32:08',1,'admin - Deaktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:32:08',0,'admin - Aktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:32:15',1,'dihapus-list rfid','admin'),
  ('0011145481','2017-06-21','13:18:07',1,'W15215 - Deaktivasi - list RFID','admin'),
  ('0011145481','2017-06-21','13:18:07',0,'W15215 - Aktivasi - list RFID','admin'),
  ('0011145481','2017-06-21','13:33:04',1,'admin - Deaktivasi - list RFID','admin'),
  ('0011145481','2017-06-21','13:33:04',0,'admin - Aktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:38:03',1,'W15217 - Deaktivasi - list RFID','admin'),
  ('0011899088','2017-06-21','13:38:03',0,'W15217 - W15223 - Print','admin'),
  ('0011787421','2017-06-21','13:38:03',1,'W15217 - W15224 - Print','admin'),
  ('0011771269','2017-06-21','13:38:03',1,'W15217 - W15225 - Print','admin'),
  ('0011366322','2017-06-23','12:12:11',0,'W15229 - W15230 - Print','admin'),
  ('0011559569','2017-06-23','12:12:14',0,'W15229 - W15230 - Print','admin'),
  ('0011559695','2017-06-23','14:04:07',0,'W15228 - W15230 - Print','admin'),
  ('0011559878','2017-06-23','14:01:57',0,'W15227 - W15230 - Print','admin'),
  ('0011366322','2017-06-23','16:16:02',1,'W15230 - W15231 - Print','admin'),
  ('0011559569','2017-06-23','16:16:02',1,'W15230 - W15231 - Print','admin'),
  ('0011559695','2017-06-23','16:16:03',1,'W15230 - W15231 - Print','admin'),
  ('0011559878','2017-06-23','16:16:03',1,'W15230 - W15231 - Print','admin'),
  ('0011366322','2017-06-23','16:17:50',1,'W15231 - W15232 - Print','admin'),
  ('0011559569','2017-06-23','16:17:50',1,'W15231 - W15232 - Print','admin'),
  ('0011559695','2017-06-23','16:17:50',1,'W15231 - W15232 - Print','admin'),
  ('0011559878','2017-06-23','16:17:50',1,'W15231 - W15232 - Print','admin'),
  ('0011366322','2017-06-23','16:23:02',1,'W15232 - W15233 - Print','admin'),
  ('0011559569','2017-06-23','16:23:02',1,'W15232 - W15233 - Print','admin'),
  ('0011559695','2017-06-23','16:23:02',1,'W15232 - W15233 - Print','admin'),
  ('0011559878','2017-06-23','16:23:02',1,'W15232 - W15233 - Print','admin'),
  ('0011366322','2017-06-23','16:24:41',1,'W15233 - W15234 - Print','admin'),
  ('0011559569','2017-06-23','16:24:41',1,'W15233 - W15234 - Print','admin'),
  ('0011559695','2017-06-23','16:24:41',1,'W15233 - W15234 - Print','admin'),
  ('0011559878','2017-06-23','16:24:41',1,'W15233 - W15234 - Print','admin'),
  ('0011366322','2017-06-23','13:24:41',0,'W15234 - W15235 - Print','admin'),
  ('0011559695','2017-06-23','16:25:33',0,'W15234 - W15236 - Print','admin'),
  ('0011559569','2017-06-23','13:25:33',0,'W15234 - W15236 - Print','admin'),
  ('0011559878','2017-06-23','16:25:33',0,'W15234 - W15236 - Print','admin'),
  ('0011626187','2017-06-23','14:01:56',0,'W15227 - W15237 - Print','admin'),
  ('0011651486','2017-06-23','14:01:54',0,'W15227 - W15237 - Print','admin'),
  ('0011739633','2017-06-23','13:56:38',0,'W15226 - W15238 - Print','admin'),
  ('0011771269','2017-06-23','13:55:31',0,'W15225 - W15238 - Print','admin'),
  ('0011787421','2017-06-23','13:52:56',0,'W15224 - W15239 - Print','admin'),
  ('0011899088','2017-06-23','13:52:17',0,'W15223 - W15240 - Print','admin'),
  ('0011366322','2017-06-24','11:21:38',1,'W15235 - W15241 - Print','admin'),
  ('0011559569','2017-06-24','12:11:04',1,'W15236 - W15241 - Print','admin'),
  ('0011559695','2017-06-24','12:11:04',1,'W15236 - W15241 - Print','admin');
COMMIT;

#
# Data for the `tbreader` table  (LIMIT 0,500)
#

INSERT INTO `tbreader` (`id`, `rfid`) VALUES
  (1,'0011366322'),
  (2,'0011559695'),
  (3,'0011559569'),
  (4,'0011559878'),
  (5,'0011626187'),
  (6,'0011651486'),
  (7,'0011739633'),
  (8,'0011771269'),
  (9,'0011787421'),
  (10,'0011899088');
COMMIT;

#
# Data for the `tbrfid` table  (LIMIT 0,500)
#

INSERT INTO `tbrfid` (`nobukti`, `rfid`) VALUES
  ('W15202','0011899088'),
  ('W15202','0011787421'),
  ('W15202','0011771269'),
  ('W15203','0011366322'),
  ('W15203','0011559569'),
  ('W15203','0011559695'),
  ('W15203','0011559878'),
  ('W15203','0011626187'),
  ('W15203','0011651486'),
  ('W15204','0011899088'),
  ('W15204','0011787421'),
  ('W15204','0011771269'),
  ('W15204','0011739633'),
  ('W15204','0011651486'),
  ('W15204','0011626187'),
  ('W15204','0011559878'),
  ('W15204','0011559695'),
  ('W15204','0011559569'),
  ('W15204','0011366322'),
  ('W15205','0011145481'),
  ('W15205','0011145385'),
  ('W15205','0011145364'),
  ('W15205','0011145357'),
  ('W15205','0011145281'),
  ('W15205','0011145060'),
  ('W15205','0011144962'),
  ('W15205','0011144960'),
  ('W15206','0011366274'),
  ('W15206','0011146024'),
  ('W15206','0011146199'),
  ('W15206','0011221671'),
  ('W15206','0011221725'),
  ('W15206','0011221726'),
  ('W15206','0011366260'),
  ('W15206','0011145504'),
  ('W15207','0011144960'),
  ('W15207','0011144962'),
  ('W15207','0011145060'),
  ('W15208','0011144960'),
  ('W15208','0011144962'),
  ('W15208','0011145060'),
  ('W15209','0011144960'),
  ('W15209','0011144962'),
  ('W15209','0011145060'),
  ('W15210','0011144960'),
  ('W15210','0011144962'),
  ('W15210','0011145060'),
  ('W15211','0011145281'),
  ('W15211','0011145357'),
  ('W15211','0011145364'),
  ('W15211','0011145385'),
  ('W15212','0011145481'),
  ('W15213','0011144960'),
  ('W15213','0011144962'),
  ('W15213','0011145481'),
  ('W15214','0011144960'),
  ('W15214','0011144962'),
  ('W15214','0011145060'),
  ('W15214','0011145281'),
  ('W15214','0011145357'),
  ('W15214','0011145364'),
  ('W15214','0011145385'),
  ('W15214','0011145481'),
  ('W15215','0011144960'),
  ('W15215','0011144962'),
  ('W15215','0011145481'),
  ('W15215','0011145385'),
  ('W15215','0011145364'),
  ('W15215','0011145357'),
  ('W15215','0011145281'),
  ('W15215','0011145060'),
  ('W15215','0011144962'),
  ('W15215','0011144960'),
  ('W15216','0011899088'),
  ('W15216','0011787421'),
  ('W15217','0011899088'),
  ('W15217','0011787421'),
  ('W15217','0011771269'),
  ('W15218','0011146024'),
  ('W15218','0011146199'),
  ('W15219','0011221671'),
  ('W15219','0011221725'),
  ('W15220','0011145504'),
  ('W15220','0011366260'),
  ('W15220','0011221726'),
  ('W15221','0011145481'),
  ('W15221','0011145385'),
  ('W15222','0011145364'),
  ('W15222','0011145357'),
  ('W15223','0011899088'),
  ('W15224','0011787421'),
  ('W15225','0011771269'),
  ('W15226','0011739633'),
  ('W15227','0011651486'),
  ('W15227','0011626187'),
  ('W15227','0011559878'),
  ('W15228','0011559695'),
  ('W15229','0011366322'),
  ('W15229','0011559569'),
  ('W15230','0011366322'),
  ('W15230','0011559569'),
  ('W15230','0011559695'),
  ('W15230','0011559878'),
  ('W15231','0011366322'),
  ('W15231','0011559569'),
  ('W15231','0011559695'),
  ('W15231','0011559878'),
  ('W15232','0011366322'),
  ('W15232','0011559569'),
  ('W15232','0011559695'),
  ('W15232','0011559878'),
  ('W15233','0011366322'),
  ('W15233','0011559569'),
  ('W15233','0011559695'),
  ('W15233','0011559878'),
  ('W15234','0011366322'),
  ('W15234','0011559569'),
  ('W15234','0011559695'),
  ('W15234','0011559878'),
  ('W15235','0011366322'),
  ('W15236','0011559695'),
  ('W15236','0011559569'),
  ('W15236','0011559878'),
  ('W15237','0011626187'),
  ('W15237','0011651486'),
  ('W15238','0011739633'),
  ('W15238','0011771269'),
  ('W15239','0011787421'),
  ('W15240','0011899088'),
  ('W15241','0011366322'),
  ('W15241','0011559569'),
  ('W15241','0011559695');
COMMIT;

#
# Data for the `tbrfiddeposit` table  (LIMIT 0,500)
#

INSERT INTO `tbrfiddeposit` (`nodeposit`, `rfid`, `hargarfid`) VALUES
  ('D17033','0011899088',15000),
  ('D17033','0011787421',15000),
  ('D17033','0011771269',15000),
  ('D17034','0011899088',15000),
  ('D17035','0011144960',15000),
  ('D17035','0011144962',15000),
  ('D17035','0011145060',15000),
  ('D17035','0011145357',15000),
  ('D17035','0011145364',15000),
  ('D17035','0011145385',15000),
  ('D17035','0011145481',15000),
  ('D17035','0011145281',15000);
COMMIT;

#
# Data for the `tbsuplier` table  (LIMIT 0,500)
#

INSERT INTO `tbsuplier` (`kdsuplier`, `nmsuplier`, `alamat`, `telp`, `tgl_gabung`, `nama_rek`, `no_rek`, `bank`) VALUES
  ('1','CHIP','Jl. Adinegoro no.11 Padang','','2017-06-17','','','');
COMMIT;



/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;