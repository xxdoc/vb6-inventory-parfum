/*
Navicat MySQL Data Transfer

Source Server         : MyLocalhost
Source Server Version : 50611
Source Host           : 127.0.0.1:3306
Source Database       : gencil_parfum

Target Server Type    : MYSQL
Target Server Version : 50611
File Encoding         : 65001

Date: 2015-06-15 06:52:41
*/

SET FOREIGN_KEY_CHECKS=0;

-- ----------------------------
-- Table structure for botol
-- ----------------------------
DROP TABLE IF EXISTS `botol`;
CREATE TABLE `botol` (
  `botol_id` int(11) NOT NULL AUTO_INCREMENT,
  `botol_tipe` varchar(255) DEFAULT NULL,
  `botol_ukuran` double DEFAULT NULL,
  `botol_stok` int(11) DEFAULT NULL,
  PRIMARY KEY (`botol_id`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records of botol
-- ----------------------------
INSERT INTO `botol` VALUES ('1', 'Cantik', '100', '200');
INSERT INTO `botol` VALUES ('3', 'Cantik', '150', '50');

-- ----------------------------
-- Table structure for kategori
-- ----------------------------
DROP TABLE IF EXISTS `kategori`;
CREATE TABLE `kategori` (
  `kategori_id` int(11) NOT NULL AUTO_INCREMENT,
  `kategori_nama` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`kategori_id`)
) ENGINE=InnoDB AUTO_INCREMENT=17 DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records of kategori
-- ----------------------------
INSERT INTO `kategori` VALUES ('1', 'Men');
INSERT INTO `kategori` VALUES ('3', 'Fruit');
INSERT INTO `kategori` VALUES ('6', 'kategori 2');
INSERT INTO `kategori` VALUES ('7', 'kategori 3');
INSERT INTO `kategori` VALUES ('8', 'kategori 4');
INSERT INTO `kategori` VALUES ('9', 'kategori 5');
INSERT INTO `kategori` VALUES ('10', 'kategori 6');
INSERT INTO `kategori` VALUES ('11', 'kategori 7');
INSERT INTO `kategori` VALUES ('12', 'kategori 8');
INSERT INTO `kategori` VALUES ('13', 'kategori 9');
INSERT INTO `kategori` VALUES ('14', 'kategori 10');
INSERT INTO `kategori` VALUES ('15', 'test');

-- ----------------------------
-- Table structure for kecelakaan
-- ----------------------------
DROP TABLE IF EXISTS `kecelakaan`;
CREATE TABLE `kecelakaan` (
  `kecelakaan_id` int(11) NOT NULL AUTO_INCREMENT,
  `kecelakaan_tanggal` datetime DEFAULT NULL,
  `kecelakaan_parfum_id` int(11) DEFAULT NULL,
  `kecelakaan_botol_id` int(11) DEFAULT NULL,
  `kecelakaan_jumlah` double DEFAULT NULL,
  `kecelakaan_keterangan` varchar(255) DEFAULT NULL,
  `kecelakaan_user_id` int(11) DEFAULT NULL,
  PRIMARY KEY (`kecelakaan_id`),
  KEY `fk_kecelakaan_parfum` (`kecelakaan_parfum_id`),
  KEY `fk_kecelakaan_botol` (`kecelakaan_botol_id`),
  KEY `fk_kecelakaan_user` (`kecelakaan_user_id`),
  CONSTRAINT `fk_kecelakaan_botol` FOREIGN KEY (`kecelakaan_botol_id`) REFERENCES `botol` (`botol_id`) ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `fk_kecelakaan_parfum` FOREIGN KEY (`kecelakaan_parfum_id`) REFERENCES `parfum` (`parfum_id`) ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `fk_kecelakaan_user` FOREIGN KEY (`kecelakaan_user_id`) REFERENCES `user` (`user_id`) ON DELETE SET NULL ON UPDATE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=11 DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records of kecelakaan
-- ----------------------------
INSERT INTO `kecelakaan` VALUES ('2', '2015-02-18 00:56:42', '4', null, '33', 'hilang dicuri orang', '1');
INSERT INTO `kecelakaan` VALUES ('3', '2015-02-18 00:57:21', null, '3', '2', 'pecah bro', '1');
INSERT INTO `kecelakaan` VALUES ('4', '2015-03-29 00:30:53', '2', null, '2', 'menguap (aja), disuruh serius', '1');
INSERT INTO `kecelakaan` VALUES ('7', '2015-06-14 23:35:39', '6', null, '150', 'harusnya dikurang 150', '1');
INSERT INTO `kecelakaan` VALUES ('8', '2015-06-14 23:38:01', '6', null, '50', 'stok jadinya 800', '1');
INSERT INTO `kecelakaan` VALUES ('9', '2015-06-14 23:39:06', '6', null, '100', 'jadinya harus 700', '1');
INSERT INTO `kecelakaan` VALUES ('10', '2015-06-14 23:46:20', null, '1', '40', 'harusnya jadi 260', '1');

-- ----------------------------
-- Table structure for output
-- ----------------------------
DROP TABLE IF EXISTS `output`;
CREATE TABLE `output` (
  `output_id` int(11) NOT NULL AUTO_INCREMENT,
  `output_tanggal` datetime DEFAULT NULL,
  `output_keterangan` text,
  `output_user_id` int(11) DEFAULT NULL,
  PRIMARY KEY (`output_id`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records of output
-- ----------------------------
INSERT INTO `output` VALUES ('1', '2015-03-28 00:00:00', 'test', '1');
INSERT INTO `output` VALUES ('2', '2015-03-28 00:00:00', 'dibeli oleh Anni, minta dianter secepatnya. catat apa ajalah pokoknya', '1');
INSERT INTO `output` VALUES ('3', '2015-03-28 00:00:00', 'test', '1');
INSERT INTO `output` VALUES ('4', '2015-03-28 00:00:00', 'asdasdasd sd', '1');
INSERT INTO `output` VALUES ('5', '2015-03-28 00:00:00', 'test keluar pertama', '1');
INSERT INTO `output` VALUES ('6', '2015-03-28 00:00:00', 'test pengurangan stok', '1');

-- ----------------------------
-- Table structure for output_detail
-- ----------------------------
DROP TABLE IF EXISTS `output_detail`;
CREATE TABLE `output_detail` (
  `odetail_id` int(11) NOT NULL AUTO_INCREMENT,
  `odetail_output_id` int(11) DEFAULT NULL,
  `odetail_parfum_id` int(11) DEFAULT NULL,
  `odetail_botol_id` int(11) DEFAULT NULL,
  `odetail_jml` int(11) DEFAULT NULL,
  `odetail_keterangan` text,
  PRIMARY KEY (`odetail_id`),
  KEY `fk_output_detail` (`odetail_output_id`),
  KEY `fk_output_parfum` (`odetail_parfum_id`),
  KEY `fk_output_botol` (`odetail_botol_id`),
  CONSTRAINT `fk_output_botol` FOREIGN KEY (`odetail_botol_id`) REFERENCES `botol` (`botol_id`) ON UPDATE CASCADE,
  CONSTRAINT `fk_output_detail` FOREIGN KEY (`odetail_output_id`) REFERENCES `output` (`output_id`) ON UPDATE CASCADE,
  CONSTRAINT `fk_output_parfum` FOREIGN KEY (`odetail_parfum_id`) REFERENCES `parfum` (`parfum_id`) ON UPDATE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=10 DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records of output_detail
-- ----------------------------
INSERT INTO `output_detail` VALUES ('1', '1', null, '1', '3', 'asd');
INSERT INTO `output_detail` VALUES ('2', '1', '2', null, '3', 'test parfum');
INSERT INTO `output_detail` VALUES ('3', '2', null, '1', '1', '-');
INSERT INTO `output_detail` VALUES ('4', '2', '4', null, '14', '-');
INSERT INTO `output_detail` VALUES ('5', '3', null, '3', '1', '');
INSERT INTO `output_detail` VALUES ('6', '3', '2', null, '12', '');
INSERT INTO `output_detail` VALUES ('7', '4', '5', null, '12', 'asda');
INSERT INTO `output_detail` VALUES ('8', '5', '6', null, '300', '300 untuk test parfum');
INSERT INTO `output_detail` VALUES ('9', '6', null, '1', '60', 'harusnya sisa 200');

-- ----------------------------
-- Table structure for parfum
-- ----------------------------
DROP TABLE IF EXISTS `parfum`;
CREATE TABLE `parfum` (
  `parfum_id` int(11) NOT NULL AUTO_INCREMENT,
  `parfum_nama` varchar(255) DEFAULT NULL,
  `parfum_tanggal` datetime DEFAULT NULL,
  `parfum_remarks` varchar(255) DEFAULT NULL,
  `parfum_status` enum('0','1') DEFAULT NULL,
  `parfum_stok` double DEFAULT NULL,
  PRIMARY KEY (`parfum_id`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records of parfum
-- ----------------------------
INSERT INTO `parfum` VALUES ('1', 'nama', '2015-02-07 12:37:17', 'keterangan', '1', '1');
INSERT INTO `parfum` VALUES ('2', 'Parfum 2', '2015-02-09 00:00:51', 'test', '0', '0');
INSERT INTO `parfum` VALUES ('3', 'parfum 3', '2015-02-09 00:06:11', 'test cek sound cek', '1', '10');
INSERT INTO `parfum` VALUES ('4', 'Parfum terbaru', '2015-02-09 23:47:13', 'Akuu ediiitttt', '1', '125');
INSERT INTO `parfum` VALUES ('5', 'AKU EDIT', '2015-02-10 00:56:04', 'ini apa ya', '1', '23');
INSERT INTO `parfum` VALUES ('6', 'TEST PARFUM', '2015-06-14 23:07:29', 'testing parfum', '1', '700');

-- ----------------------------
-- Table structure for parfum_kategori
-- ----------------------------
DROP TABLE IF EXISTS `parfum_kategori`;
CREATE TABLE `parfum_kategori` (
  `pk_id` int(11) NOT NULL AUTO_INCREMENT,
  `pk_parfum_id` int(11) DEFAULT NULL,
  `pk_kategori_id` int(11) DEFAULT NULL,
  PRIMARY KEY (`pk_id`),
  UNIQUE KEY `uq_pk` (`pk_parfum_id`,`pk_kategori_id`),
  KEY `fk_pk_kategori` (`pk_kategori_id`),
  KEY `fk_pk_parfum` (`pk_parfum_id`),
  CONSTRAINT `fk_pk_kategori` FOREIGN KEY (`pk_kategori_id`) REFERENCES `kategori` (`kategori_id`) ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `fk_pk_parfum` FOREIGN KEY (`pk_parfum_id`) REFERENCES `parfum` (`parfum_id`) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=35 DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records of parfum_kategori
-- ----------------------------
INSERT INTO `parfum_kategori` VALUES ('28', '1', '3');
INSERT INTO `parfum_kategori` VALUES ('30', '1', '12');
INSERT INTO `parfum_kategori` VALUES ('24', '2', '1');
INSERT INTO `parfum_kategori` VALUES ('23', '2', '3');
INSERT INTO `parfum_kategori` VALUES ('33', '6', '3');
INSERT INTO `parfum_kategori` VALUES ('34', '6', '13');

-- ----------------------------
-- Table structure for sirkulasi
-- ----------------------------
DROP TABLE IF EXISTS `sirkulasi`;
CREATE TABLE `sirkulasi` (
  `sirkulasi_id` int(11) NOT NULL AUTO_INCREMENT,
  `sirkulasi_tanggal_pesan` datetime DEFAULT NULL,
  `sirkulasi_tanggal_terima` datetime DEFAULT NULL,
  `sirkulasi_user_id` int(11) DEFAULT NULL,
  `sirkulasi_status` enum('0','1','99') DEFAULT '0' COMMENT '0 = pesan, 1 = diterima, 99 = batal',
  PRIMARY KEY (`sirkulasi_id`),
  KEY `sirkulasi_user` (`sirkulasi_user_id`),
  CONSTRAINT `sirkulasi_user` FOREIGN KEY (`sirkulasi_user_id`) REFERENCES `user` (`user_id`) ON UPDATE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records of sirkulasi
-- ----------------------------
INSERT INTO `sirkulasi` VALUES ('1', '2015-04-01 15:23:50', '2015-04-08 15:24:16', '1', '1');
INSERT INTO `sirkulasi` VALUES ('2', '2015-04-06 15:31:40', '2015-04-08 15:31:44', '1', '0');

-- ----------------------------
-- Table structure for sirkulasi_detail
-- ----------------------------
DROP TABLE IF EXISTS `sirkulasi_detail`;
CREATE TABLE `sirkulasi_detail` (
  `detail_id` int(11) NOT NULL AUTO_INCREMENT,
  `detail_sirkulasi_id` int(11) DEFAULT NULL,
  `detail_parfum_id` int(11) DEFAULT NULL,
  `detail_botol_id` int(11) DEFAULT NULL,
  `detail_jml_pesan` double DEFAULT NULL,
  `detail_jml_terima_kotor` double DEFAULT NULL,
  `detail_jml_terima_bersih` double DEFAULT NULL,
  `detail_keterangan` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`detail_id`),
  KEY `sirkulasi_detail` (`detail_sirkulasi_id`),
  KEY `sirkulasi_parfum` (`detail_parfum_id`),
  KEY `sirkulasi_botol` (`detail_botol_id`),
  CONSTRAINT `sirkulasi_botol` FOREIGN KEY (`detail_botol_id`) REFERENCES `botol` (`botol_id`) ON UPDATE CASCADE,
  CONSTRAINT `sirkulasi_detail` FOREIGN KEY (`detail_sirkulasi_id`) REFERENCES `sirkulasi` (`sirkulasi_id`) ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `sirkulasi_parfum` FOREIGN KEY (`detail_parfum_id`) REFERENCES `parfum` (`parfum_id`) ON UPDATE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records of sirkulasi_detail
-- ----------------------------
INSERT INTO `sirkulasi_detail` VALUES ('1', '1', '4', null, '10', '10', '8', 'bagus');
INSERT INTO `sirkulasi_detail` VALUES ('2', '1', null, '3', '5', '4', '4', '1 rusak');
INSERT INTO `sirkulasi_detail` VALUES ('3', '1', '2', null, '12', '12', '11', 'sip');
INSERT INTO `sirkulasi_detail` VALUES ('4', '2', '4', null, '23', '13', '12', 'ok');

-- ----------------------------
-- Table structure for supplier
-- ----------------------------
DROP TABLE IF EXISTS `supplier`;
CREATE TABLE `supplier` (
  `supplier_id` int(11) NOT NULL AUTO_INCREMENT,
  `nama_supplier` varchar(30) DEFAULT NULL,
  `cp_supplier` varchar(13) DEFAULT NULL,
  `alamat` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`supplier_id`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records of supplier
-- ----------------------------
INSERT INTO `supplier` VALUES ('2', 'supp', 'asd', 'qaswdasd');
INSERT INTO `supplier` VALUES ('4', 'supp as', 'asdas', 'asdasd');

-- ----------------------------
-- Table structure for user
-- ----------------------------
DROP TABLE IF EXISTS `user`;
CREATE TABLE `user` (
  `user_id` int(11) NOT NULL AUTO_INCREMENT,
  `user_nama` varchar(255) DEFAULT NULL,
  `user_password` varchar(255) DEFAULT NULL,
  `user_level` int(11) DEFAULT '0',
  PRIMARY KEY (`user_id`),
  UNIQUE KEY `uq_username` (`user_nama`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records of user
-- ----------------------------
INSERT INTO `user` VALUES ('1', 'admin', 'admin', '1');
INSERT INTO `user` VALUES ('2', 'user', 'user', '0');

-- ----------------------------
-- View structure for vw_inventory_output
-- ----------------------------
DROP VIEW IF EXISTS `vw_inventory_output`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER  VIEW `vw_inventory_output` AS SELECT
	od.odetail_parfum_id as parfum_id, od.odetail_botol_id as botol_id, SUM(od.odetail_jml) as jml_keluar
FROM
output_detail od
JOIN output o ON od.odetail_output_id = o.output_id
GROUP BY od.odetail_parfum_id, od.odetail_botol_id ;

-- ----------------------------
-- View structure for vw_inventory_sirkulasi
-- ----------------------------
DROP VIEW IF EXISTS `vw_inventory_sirkulasi`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost`  VIEW `vw_inventory_sirkulasi` AS SELECT
	sd.detail_parfum_id as parfum_id,
	sd.detail_botol_id as botol_id,
	SUM(detail_jml_terima_bersih) as jml_bersih,
	SUM(detail_jml_terima_kotor) as jml_kotor
FROM 
sirkulasi_detail sd
JOIN sirkulasi s on sd.detail_sirkulasi_id = s.sirkulasi_id
WHERE s.sirkulasi_status = '1'
GROUP BY sd.detail_parfum_id, sd.detail_botol_id ;

-- ----------------------------
-- View structure for vw_kecelakaan
-- ----------------------------
DROP VIEW IF EXISTS `vw_kecelakaan`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost`  VIEW `vw_kecelakaan` AS SELECT 
	kecelakaan_id, kecelakaan_tanggal,
	botol_id, botol_tipe, botol_ukuran, 
	parfum_id, parfum_nama, kecelakaan_jumlah, kecelakaan_keterangan,
	user_id, user_nama
FROM 
	kecelakaan k
LEFT JOIN parfum p on k.kecelakaan_parfum_id = p.parfum_id
LEFT JOIN botol b on k.kecelakaan_botol_id = b.botol_id
LEFT JOIN user u on k.kecelakaan_user_id = u.user_id ;

-- ----------------------------
-- View structure for vw_summary_inventory
-- ----------------------------
DROP VIEW IF EXISTS `vw_summary_inventory`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER  VIEW `vw_summary_inventory` AS  ;
