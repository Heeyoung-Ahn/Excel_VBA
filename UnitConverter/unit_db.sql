-- --------------------------------------------------------
-- Host:                         172.17.110.91
-- Server version:               10.4.7-MariaDB - mariadb.org binary distribution
-- Server OS:                    Win64
-- HeidiSQL Version:             11.0.0.5919
-- --------------------------------------------------------

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET NAMES utf8 */;
/*!50503 SET NAMES utf8mb4 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;

-- Dumping structure for table overseas.unit_conversion
DROP TABLE IF EXISTS `unit_conversion`;
CREATE TABLE IF NOT EXISTS `unit_conversion` (
  `unit_id1` int(3) unsigned NOT NULL,
  `unit_id2` int(3) unsigned NOT NULL,
  `value` decimal(20,10) NOT NULL DEFAULT 0.0000000000,
  `cmt` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`unit_id1`,`unit_id2`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Dumping data for table overseas.unit_conversion: ~14 rows (approximately)
/*!40000 ALTER TABLE `unit_conversion` DISABLE KEYS */;
INSERT INTO `unit_conversion` (`unit_id1`, `unit_id2`, `value`, `cmt`) VALUES
	(1, 2, 39.3700790000, NULL),
	(1, 3, 3.2808400000, NULL),
	(1, 4, 0.0006213710, NULL),
	(1, 5, 1.0936130000, NULL),
	(6, 7, 0.3025000000, NULL),
	(6, 8, 0.0100000000, NULL),
	(6, 9, 0.0001000000, NULL),
	(6, 10, 10.7639100000, NULL),
	(6, 11, 1.1959900000, NULL),
	(6, 12, 0.0002471050, NULL),
	(13, 14, 15432.3584000000, NULL),
	(13, 15, 35.2739620000, NULL),
	(13, 16, 2.2046230000, NULL),
	(13, 17, 1.6666670000, NULL);
/*!40000 ALTER TABLE `unit_conversion` ENABLE KEYS */;

-- Dumping structure for table overseas.unit_type
DROP TABLE IF EXISTS `unit_type`;
CREATE TABLE IF NOT EXISTS `unit_type` (
  `unit_id` int(3) unsigned NOT NULL AUTO_INCREMENT,
  `unit_gb` varchar(20) NOT NULL,
  `unit_gb_ko` varchar(20) NOT NULL,
  `unit_standard` tinyint(1) NOT NULL COMMENT '각 단위 카테고리별 기본 단위의 id',
  `unit` varchar(20) NOT NULL,
  `sort_order` smallint(4) unsigned NOT NULL,
  `suspended` tinyint(1) unsigned NOT NULL DEFAULT 0 COMMENT '1: suspended',
  PRIMARY KEY (`unit_id`)
) ENGINE=InnoDB AUTO_INCREMENT=22 DEFAULT CHARSET=utf8 COMMENT='단위의 종류';

-- Dumping data for table overseas.unit_type: ~17 rows (approximately)
/*!40000 ALTER TABLE `unit_type` DISABLE KEYS */;
INSERT INTO `unit_type` (`unit_id`, `unit_gb`, `unit_gb_ko`, `unit_standard`, `unit`, `sort_order`, `suspended`) VALUES
	(1, 'length', '길이', 1, '미터(m)', 10, 0),
	(2, 'length', '길이', 1, '인치(in)', 20, 0),
	(3, 'length', '길이', 1, '피트(ft)', 30, 0),
	(4, 'length', '길이', 1, '마일(mile)', 40, 0),
	(5, 'length', '길이', 1, '야드(yd)', 50, 0),
	(6, 'area', '넓이', 6, '제곱미터(m2)', 60, 0),
	(7, 'area', '넓이', 6, '평', 70, 0),
	(8, 'area', '넓이', 6, '아르(a)', 80, 0),
	(9, 'area', '넓이', 6, '헥타르(ha)', 90, 0),
	(10, 'area', '넓이', 6, '제곱피트(ft2)', 100, 0),
	(11, 'area', '넓이', 6, '제곱야드(yd2)', 110, 0),
	(12, 'area', '넓이', 6, '에이커(ac)', 120, 0),
	(13, 'weight', '무게', 13, '킬로그램(kg)', 130, 0),
	(14, 'weight', '무게', 13, '그레인(gr)', 140, 0),
	(15, 'weight', '무게', 13, '온스(oz)', 150, 0),
	(16, 'weight', '무게', 13, '파운드(lb)', 160, 0),
	(17, 'weight', '무게', 13, '근', 170, 0);
/*!40000 ALTER TABLE `unit_type` ENABLE KEYS */;

/*!40101 SET SQL_MODE=IFNULL(@OLD_SQL_MODE, '') */;
/*!40014 SET FOREIGN_KEY_CHECKS=IF(@OLD_FOREIGN_KEY_CHECKS IS NULL, 1, @OLD_FOREIGN_KEY_CHECKS) */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
