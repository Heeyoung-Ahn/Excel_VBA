-- --------------------------------------------------------
-- Host:                         127.0.0.1
-- Server version:               10.4.7-MariaDB - mariadb.org binary distribution
-- Server OS:                    Win64
-- HeidiSQL Version:             10.3.0.5771
-- --------------------------------------------------------

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET NAMES utf8 */;
/*!50503 SET NAMES utf8mb4 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;


-- Dumping database structure for overseas
CREATE DATABASE IF NOT EXISTS `overseas` /*!40100 DEFAULT CHARACTER SET utf8 */;
USE `overseas`;

-- Dumping structure for table overseas.currency_cal
CREATE TABLE IF NOT EXISTS `currency_cal` (
  `currency_id` smallint(3) unsigned NOT NULL,
  `currency_un` varchar(3) NOT NULL,
  `refer_dt` date NOT NULL,
  `fx_rate_krw` double unsigned NOT NULL,
  `fx_rate_usd` double unsigned NOT NULL,
  `user_id` tinyint(3) unsigned NOT NULL,
  PRIMARY KEY (`currency_id`),
  KEY `user_id` (`user_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COMMENT='화폐 목록';

-- Dumping data for table overseas.currency_cal: ~0 rows (approximately)
/*!40000 ALTER TABLE `currency_cal` DISABLE KEYS */;
/*!40000 ALTER TABLE `currency_cal` ENABLE KEYS */;

/*!40101 SET SQL_MODE=IFNULL(@OLD_SQL_MODE, '') */;
/*!40014 SET FOREIGN_KEY_CHECKS=IF(@OLD_FOREIGN_KEY_CHECKS IS NULL, 1, @OLD_FOREIGN_KEY_CHECKS) */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
