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


-- Dumping database structure for common
CREATE DATABASE IF NOT EXISTS `common` /*!40100 DEFAULT CHARACTER SET utf8 */;
USE `common`;

-- Dumping structure for table common.users
CREATE TABLE IF NOT EXISTS `users` (
  `user_id` smallint(3) unsigned NOT NULL AUTO_INCREMENT COMMENT '사용자 id',
  `user_nm` varchar(20) NOT NULL COMMENT '엑셀의 사용자 이름으로 사용',
  `user_gb` varchar(5) NOT NULL DEFAULT 'WP' COMMENT 'SA, AM(리포트), MG(실무관리), WP(실무)',
  `user_pw` varchar(128) NOT NULL DEFAULT '1' COMMENT '비밀번호',
  `pw_initialize` tinyint(1) unsigned NOT NULL DEFAULT 1 COMMENT '1: 최초접속(비밀번호 초기화)',
  `user_ip` varchar(20) DEFAULT NULL,
  `programv` varchar(20) NOT NULL DEFAULT 'programv' COMMENT '프로그램버전',
  `argIP` varchar(20) NOT NULL DEFAULT 'DBIP' COMMENT '작업용DB IP',
  `argDB` varchar(30) NOT NULL DEFAULT 'common' COMMENT '작업용DB 스키마',
  `argUN` varchar(30) NOT NULL DEFAULT 'common' COMMENT '작업용DB UN',
  `argPW` varchar(30) NOT NULL DEFAULT 'common pw' COMMENT '작업용DB PW',
  `suspended` tinyint(1) unsigned NOT NULL DEFAULT 0 COMMENT '1: suspended',
  `time_stamp` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp(),
  PRIMARY KEY (`user_id`),
  KEY `user_gb` (`user_gb`),
  KEY `user_nm` (`user_nm`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COMMENT='프로그램 사용자';

-- Dumping data for table common.users: ~0 rows (approximately)
/*!40000 ALTER TABLE `users` DISABLE KEYS */;
/*!40000 ALTER TABLE `users` ENABLE KEYS */;

/*!40101 SET SQL_MODE=IFNULL(@OLD_SQL_MODE, '') */;
/*!40014 SET FOREIGN_KEY_CHECKS=IF(@OLD_FOREIGN_KEY_CHECKS IS NULL, 1, @OLD_FOREIGN_KEY_CHECKS) */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
