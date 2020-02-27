CREATE DATABASE IF NOT EXISTS `common`;
USE `common`;

CREATE TABLE IF NOT EXISTS `logs` (
  `log_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `procedure_nm` varchar(200) NOT NULL COMMENT 'Function or Procedure name',
  `table_nm` varchar(50) NOT NULL,
  `form_nm` varchar(30) DEFAULT NULL,
  `job_nm` varchar(20) DEFAULT NULL,
  `error_cd` tinyint(1) unsigned NOT NULL COMMENT '1: 오류, 0: 오류 아님',
  `affectedCount` mediumint(10) unsigned NOT NULL DEFAULT 0 COMMENT '영향받은 레코드 수',
  `sql_script` varchar(700) NOT NULL,
  `user_id` smallint(3) unsigned NOT NULL DEFAULT 1,
  `time_stamp` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp(),
  PRIMARY KEY (`log_id`),
  KEY `personincharge` (`user_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COMMENT='프로그램 로그';

CREATE TABLE IF NOT EXISTS `users` (
  `user_id` smallint(3) unsigned NOT NULL AUTO_INCREMENT COMMENT '사용자 id',
  `user_nm` varchar(20) NOT NULL COMMENT '엑셀의 사용자 이름으로 사용',
  `user_gb` varchar(5) NOT NULL DEFAULT 'WP' COMMENT 'SA, AM(리포트), MG(실무관리), WP(실무)',
  `user_pw` varchar(128) DEFAULT NULL COMMENT '비밀번호',
  `pw_initialize` tinyint(1) unsigned NOT NULL DEFAULT 1 COMMENT '1: 최초접속(비밀번호 초기화)',
  `user_ip` varchar(20) DEFAULT NULL,
  `user_dept` varchar(20) NOT NULL,
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