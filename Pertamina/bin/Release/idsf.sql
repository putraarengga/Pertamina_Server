-- MySqlBackup.NET 2.0.9.2
-- Dump Time: 2017-01-20 06:45:28
-- --------------------------------------
-- Server version 5.1.73 Source distribution

-- 
-- Create schema idsf
-- 

CREATE DATABASE IF NOT EXISTS `idsf` /*!40100 DEFAULT CHARACTER SET utf8 */;
Use `idsf`;



/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;


-- 
-- Definition of cpengguna
-- 

DROP TABLE IF EXISTS `cpengguna`;
CREATE TABLE IF NOT EXISTS `cpengguna` (
  `ida` int(2) DEFAULT NULL,
  `idb` int(2) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- 
-- Dumping data for table cpengguna
-- 

/*!40000 ALTER TABLE `cpengguna` DISABLE KEYS */;
INSERT INTO `cpengguna`(`ida`,`idb`) VALUES
(1,2),
(1,1),
(2,1),
(2,2);
/*!40000 ALTER TABLE `cpengguna` ENABLE KEYS */;

-- 
-- Definition of cuser
-- 

DROP TABLE IF EXISTS `cuser`;
CREATE TABLE IF NOT EXISTS `cuser` (
  `id` int(2) DEFAULT NULL,
  `nama` varchar(30) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- 
-- Dumping data for table cuser
-- 

/*!40000 ALTER TABLE `cuser` DISABLE KEYS */;
INSERT INTO `cuser`(`id`,`nama`) VALUES
(1,'andre'),
(2,'putra');
/*!40000 ALTER TABLE `cuser` ENABLE KEYS */;

-- 
-- Definition of tdatatujuan
-- 

DROP TABLE IF EXISTS `tdatatujuan`;
CREATE TABLE IF NOT EXISTS `tdatatujuan` (
  `IDTujuan` int(7) NOT NULL,
  `namaTujuan` varchar(30) NOT NULL,
  `alamatTujuan` varchar(50) NOT NULL,
  PRIMARY KEY (`IDTujuan`),
  UNIQUE KEY `IDTujuan` (`IDTujuan`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

-- 
-- Dumping data for table tdatatujuan
-- 

/*!40000 ALTER TABLE `tdatatujuan` DISABLE KEYS */;
INSERT INTO `tdatatujuan`(`IDTujuan`,`namaTujuan`,`alamatTujuan`) VALUES
(764089,'PLTD Waena','alan Kambolker Perumnas II, Yabansai, Heram, Kota '),
(764222,'PLTD Yarmokh','Kota Jayapura, Yarmokh ,Papua Barat'),
(761131,'PLTD Sentani','Sentani Kota, Sentani, Jayapura, Papua'),
(768984,'PLTD ARSO','');
/*!40000 ALTER TABLE `tdatatujuan` ENABLE KEYS */;

-- 
-- Definition of tdatauser
-- 

DROP TABLE IF EXISTS `tdatauser`;
CREATE TABLE IF NOT EXISTS `tdatauser` (
  `IDUser` int(20) NOT NULL AUTO_INCREMENT,
  `NamaUser` varchar(50) DEFAULT NULL,
  `Password` varchar(30) DEFAULT NULL,
  `NamaLengkap` varchar(30) NOT NULL,
  `NoKTP` varchar(16) DEFAULT NULL,
  `IDJenisUser` int(10) DEFAULT NULL,
  PRIMARY KEY (`IDUser`)
) ENGINE=MyISAM AUTO_INCREMENT=11 DEFAULT CHARSET=latin1;

-- 
-- Dumping data for table tdatauser
-- 

/*!40000 ALTER TABLE `tdatauser` DISABLE KEYS */;
INSERT INTO `tdatauser`(`IDUser`,`NamaUser`,`Password`,`NamaLengkap`,`NoKTP`,`IDJenisUser`) VALUES
(1,'admin','admin','adminstrator','0123456789012345',1),
(8,'Bachtiar','Bachtiar','Bachtiar','1234567890123',2),
(3,'alif','alif','muhammad alif','0123456789801234',5),
(9,'putra','putra','Bayu Arengga Putra','3506251009930003',10),
(0,'null','null','-','null',0),
(10,'bayu','bayu','Bayu Setiawan','12376518741',10);
/*!40000 ALTER TABLE `tdatauser` ENABLE KEYS */;

-- 
-- Definition of tdatauserclient
-- 

DROP TABLE IF EXISTS `tdatauserclient`;
CREATE TABLE IF NOT EXISTS `tdatauserclient` (
  `IDUser` int(20) NOT NULL AUTO_INCREMENT,
  `NamaUser` varchar(50) DEFAULT NULL,
  `Password` varchar(30) DEFAULT NULL,
  `NamaLengkap` varchar(30) NOT NULL,
  `NoKTP` varchar(16) DEFAULT NULL,
  `IDJenisUser` int(10) DEFAULT NULL,
  PRIMARY KEY (`IDUser`)
) ENGINE=MyISAM AUTO_INCREMENT=7 DEFAULT CHARSET=latin1;

-- 
-- Dumping data for table tdatauserclient
-- 

/*!40000 ALTER TABLE `tdatauserclient` DISABLE KEYS */;
INSERT INTO `tdatauserclient`(`IDUser`,`NamaUser`,`Password`,`NamaLengkap`,`NoKTP`,`IDJenisUser`) VALUES
(1,'admin','admin','adminstrator','0123456789012345',1),
(2,'putra','putra','bayu arengga putra','5432109876543210',3),
(3,'alif','alif','muhammad alif','0123456789801234',2),
(5,'andre','andre','andrie yuwono','1111111111111',5),
(6,'nando','nando','nando suprapto','60251009934',1);
/*!40000 ALTER TABLE `tdatauserclient` ENABLE KEYS */;

-- 
-- Definition of tdistribusi
-- 

DROP TABLE IF EXISTS `tdistribusi`;
CREATE TABLE IF NOT EXISTS `tdistribusi` (
  `IDDistribusi` int(25) NOT NULL AUTO_INCREMENT,
  `IDKendaraan` varchar(25) DEFAULT NULL,
  `IDUser` int(10) DEFAULT NULL,
  `IDTujuan` int(7) NOT NULL,
  `NoDO` varchar(20) DEFAULT NULL,
  `wktMuat` time DEFAULT NULL,
  `wktSampai` time DEFAULT NULL,
  `tglMuat` date DEFAULT NULL,
  `tglSampai` date DEFAULT NULL,
  `dataBarcode` varchar(15) DEFAULT NULL,
  `Keterangan` varchar(50) DEFAULT NULL,
  `tempatLoading` varchar(25) NOT NULL,
  `IDUserClient` int(11) NOT NULL,
  PRIMARY KEY (`IDDistribusi`)
) ENGINE=MyISAM AUTO_INCREMENT=179 DEFAULT CHARSET=latin1;

-- 
-- Dumping data for table tdistribusi
-- 

/*!40000 ALTER TABLE `tdistribusi` DISABLE KEYS */;
INSERT INTO `tdistribusi`(`IDDistribusi`,`IDKendaraan`,`IDUser`,`IDTujuan`,`NoDO`,`wktMuat`,`wktSampai`,`tglMuat`,`tglSampai`,`dataBarcode`,`Keterangan`,`tempatLoading`,`IDUserClient`) VALUES
(116,'9',1,764089,'8005527501','10:07:00',NULL,'2016-12-26 00:00:00',NULL,'201612260116','LEVEL CHECKED','',0),
(173,'9',1,764089,'123456789','03:27:00','03:26:00','2016-12-28 00:00:00','2016-12-28 00:00:00','201612280173','ACCEPTED','PLTD Waena',9),
(174,'36',1,761131,'2314657161','03:35:00','03:30:00','2016-12-28 00:00:00','2016-12-28 00:00:00','201612280174','REJECTED','PLTD Waena',9),
(112,'11',1,764089,'123131','08:18:00','08:53:00','2016-12-26 00:00:00','2016-12-26 00:00:00','201612260000','ACCEPTED','PLTD Waena',1),
(175,'9',1,764089,'123456780','03:26:00',NULL,'2016-12-01 00:00:00',NULL,'201612010175','REGISTERED','',0),
(166,'21',1,764089,'8005690600','06:26:00','03:47:00','2016-12-27 00:00:00','2016-12-28 00:00:00','201612270166','ACCEPTED','PLTD Waena',9),
(167,'9',1,764222,'8005690700','06:29:00','11:18:00','2016-12-27 00:00:00','2016-12-27 00:00:00','201612270167','REJECTED','PLTD Waena',9),
(168,'9',1,761131,'8005690701','06:29:00','11:11:00','2016-12-27 00:00:00','2016-12-27 00:00:00','201612270168','REJECTED','PLTD Waena',9),
(171,'14',1,764089,'8005690704','06:37:00','11:21:00','2016-12-27 00:00:00','2016-12-27 00:00:00','201612270171','ACCEPTED','PLTD Waena',9),
(172,'13',1,761131,'8005690705','07:38:00','07:47:00','2016-12-27 00:00:00','2016-12-27 00:00:00','201612270172','REJECTED','PLTD Waena',9),
(176,'32',1,764222,'222333','06:32:00','06:40:00','2016-12-29 00:00:00','2016-12-29 00:00:00','201612290176','REJECTED','PLTD Waena',9),
(177,'33',1,764222,'656535353','07:47:00',NULL,'2017-01-02 00:00:00',NULL,'201701020177','REGISTERED','',0),
(178,'11',1,764089,'1212121','07:50:00',NULL,'2017-01-02 00:00:00',NULL,'201701020178','REGISTERED','',0);
/*!40000 ALTER TABLE `tdistribusi` ENABLE KEYS */;

-- 
-- Definition of tjenisuser
-- 

DROP TABLE IF EXISTS `tjenisuser`;
CREATE TABLE IF NOT EXISTS `tjenisuser` (
  `IDJenisUser` int(10) NOT NULL AUTO_INCREMENT,
  `JenisUser` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`IDJenisUser`)
) ENGINE=MyISAM AUTO_INCREMENT=12 DEFAULT CHARSET=latin1;

-- 
-- Dumping data for table tjenisuser
-- 

/*!40000 ALTER TABLE `tjenisuser` DISABLE KEYS */;
INSERT INTO `tjenisuser`(`IDJenisUser`,`JenisUser`) VALUES
(1,'admin'),
(2,'PERTAMINA'),
(3,'PLTD Arso'),
(5,'PLTD Sentani'),
(9,'PLTD YSK'),
(10,'PLTD Waena'),
(0,'Not Identified');
/*!40000 ALTER TABLE `tjenisuser` ENABLE KEYS */;

-- 
-- Definition of tkendaraan
-- 

DROP TABLE IF EXISTS `tkendaraan`;
CREATE TABLE IF NOT EXISTS `tkendaraan` (
  `IDKendaraan` int(20) NOT NULL AUTO_INCREMENT,
  `noPolKendaraan` varchar(10) DEFAULT NULL,
  `namaSopir` varchar(25) DEFAULT NULL,
  `namaKernet` varchar(25) DEFAULT NULL,
  `kapasitasTruk` int(10) DEFAULT NULL,
  `callCenter` varchar(15) NOT NULL,
  `namaPerusahaan` varchar(25) NOT NULL,
  PRIMARY KEY (`IDKendaraan`)
) ENGINE=MyISAM AUTO_INCREMENT=37 DEFAULT CHARSET=latin1;

-- 
-- Dumping data for table tkendaraan
-- 

/*!40000 ALTER TABLE `tkendaraan` DISABLE KEYS */;
INSERT INTO `tkendaraan`(`IDKendaraan`,`noPolKendaraan`,`namaSopir`,`namaKernet`,`kapasitasTruk`,`callCenter`,`namaPerusahaan`) VALUES
(9,'DS 9412 AL','JUMADE','FAHARUDDIN',5000,'0967 - 536419','PT.USAHA PANGKEP MANDIRI'),
(10,'DS 9364 A','UMAR','FIRMAN',5000,'0967 - 536419','PT.USAHA PANGKEP MANDIRI'),
(11,'DS 9973 AD','ARIFIN','CANDRA',5000,'0967 - 536419','PT.USAHA PANGKEP MANDIRI'),
(12,'DS 9730 AD','WAWAN','JEFRI',5000,'0967 - 536419','PT.USAHA PANGKEP MANDIRI'),
(13,'DS 9767 AD','PANDI','SYARIF',5000,'0967 - 536419','PT.USAHA PANGKEP MANDIRI'),
(14,'DS 9747 AC','ANSAR','ARIFUDIN',10000,'0967 - 536419','PT.USAHA PANGKEP MANDIRI'),
(15,'DS 9915 AF','SUDIRMAN','ANSAR',16000,'0967 - 536419','PT.USAHA PANGKEP MANDIRI'),
(16,'L 9744 UQ','INAL','ISMAIL',10000,'0967 - 536419','PT.USAHA PANGKEP MANDIRI'),
(17,'DS 9631 AF','-','-',5000,'0967 - 591833','PT. TRULI UTAMA INDAH'),
(18,'DS 9708 AF','-','-',10000,'0967 - 591833','PT. TRULI UTAMA INDAH'),
(19,'DS 9758 AF','-','-',10000,'0967 - 591833','PT. TRULI UTAMA INDAH'),
(20,'DS 9695 JL','ISMAIL','RIFALDY LATUPU',10000,'0967 - 534418','PT. WIRA SEMBADA PERKASA'),
(21,'DS 9724 JL','SUMARYANTO','ROMI NGADI',10000,'0967 - 534418','PT. WIRA SEMBADA PERKASA'),
(22,'DS9913 AE','LEONARD WATTIMENA','PREDRIK SOHILAIT',5000,'0967 - 934419','PT. ANDARIA JAYA UTAMA'),
(23,'DS 9914 AE','ABDUL MAJID','ASRIADI ALI',5000,'0967 - 934419','PT. ANDARIA JAYA UTAMA'),
(24,'DS 9686 AF','SUMARNO','SUKAMIN',10000,'0967 - 934419','PT. ANDARIA JAYA UTAMA'),
(25,'DS 9610 AF','RANCE RUNTUWENE','AMIR SYARIFUDIN',5000,'0967 - 934419','PT. ANDARIA JAYA UTAMA'),
(26,'DS 9719 AE','MAHMUD','IKSAN MAKMUR',5000,'0967 - 224401','PT. NAGOYA SEJATI PAPUA B'),
(27,'DS 9726 AC','PETRUS MARSIAT','ISMAIL SAMAD',10000,'0967 - 224401','PT. NAGOYA SEJATI PAPUA B'),
(28,'DS 9629 AA','AHMAD HAER DANI','DAUD PASANG',10000,'0967 - 224401','PT. NAGOYA SEJATI PAPUA B'),
(29,'DS 9804 AD','ISWANDI','INDRA RUDIANTO',5000,'0967 - 124345','PT. NURSADY SEJATI'),
(30,'DS 9894 AE','YACOBUS SAMBA','WENDI PRATAMA WARDANA',5000,'0967 - 124345','PT. NURSADY SEJATI'),
(31,'DS 9693 AC','REMON SINAI','-',10000,'0967 - 124345','PT. NURSADY SEJATI'),
(32,'B 9343 UFA','-','-',10000,'0967 - 124345','PT. NURSADY SEJATI'),
(33,'DS 9965 AF','DODI WAISAPI','-',10000,'0967 - 120264','PT. SINDITA SALSABILA'),
(34,'DS 9745 AF','JASMINTO','-',10000,'0967 - 120264','PT. SINDITA SALSABILA'),
(35,'DS 9703 AF','EFRAYIN ROSYANTO','-',10000,'0967 - 120264','PT. SINDITA SALSABILA'),
(36,'DS 9923 AF','JAMALUDIN','-',16000,'0967 - 120264','PT. SINDITA SALSABILA');
/*!40000 ALTER TABLE `tkendaraan` ENABLE KEYS */;


/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;


-- Dump completed on 2017-01-20 06:45:30
-- Total time: 0:0:0:1:711 (d:h:m:s:ms)
