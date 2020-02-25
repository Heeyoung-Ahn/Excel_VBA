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

-- Dumping structure for table overseas.currencies
CREATE TABLE IF NOT EXISTS `currencies` (
  `currency_id` smallint(3) unsigned NOT NULL AUTO_INCREMENT,
  `currency_un` varchar(3) NOT NULL,
  `currency_nm` varchar(30) NOT NULL,
  `currency_cmt` varchar(100) DEFAULT NULL,
  `sort_order` smallint(5) unsigned NOT NULL DEFAULT 10000,
  `suspended` tinyint(1) unsigned NOT NULL DEFAULT 0 COMMENT '1: suspended',
  PRIMARY KEY (`currency_id`)
) ENGINE=InnoDB AUTO_INCREMENT=162 DEFAULT CHARSET=utf8 COMMENT='화폐 목록';

-- Dumping data for table overseas.currencies: ~161 rows (approximately)
/*!40000 ALTER TABLE `currencies` DISABLE KEYS */;
INSERT INTO `currencies` (`currency_id`, `currency_un`, `currency_nm`, `currency_cmt`, `sort_order`, `suspended`) VALUES
	(1, 'AED', 'Emirati Dirham', '', 10, 0),
	(2, 'AFN', 'Afghan Afghani', '', 20, 0),
	(3, 'ALL', 'Albanian Lek', '', 30, 0),
	(4, 'AMD', 'Armenian Dram', '', 40, 0),
	(5, 'ANG', 'Dutch Guilder', '', 50, 0),
	(6, 'AOA', 'Angolan Kwanza', '', 60, 0),
	(7, 'ARS', 'Argentine Peso', '', 70, 0),
	(8, 'ATS', 'SCHILLING', '', 80, 0),
	(9, 'AUD', 'Australian Dollar', '', 90, 0),
	(10, 'AWG', 'Aruban or Dutch Guilder', '', 100, 0),
	(11, 'AZN', 'Azerbaijan Manat', '', 110, 0),
	(12, 'BBD', 'Barbadian or Bajan Dollar', '', 120, 0),
	(13, 'BDT', 'Bangladeshi Taka', '', 130, 0),
	(14, 'BEF', 'FRANC', '', 140, 0),
	(15, 'BGN', 'Bulgarian Lev', '', 150, 0),
	(16, 'BHD', 'Bahraini Dinar', '', 160, 0),
	(17, 'BIF', 'Burundian Franc', '', 170, 0),
	(18, 'BMD', 'Bermudian Dollar', '', 180, 0),
	(19, 'BND', 'Bruneian Dollar', '', 190, 0),
	(20, 'BOB', 'Bolivian Bolíviano', '', 200, 0),
	(21, 'BRL', 'Brazilian Real', '', 210, 0),
	(22, 'BSD', 'Bahamian Dollar', '', 220, 0),
	(23, 'BTN', 'Bhutanese Ngultrum', '', 230, 0),
	(24, 'BWP', 'Botswana Pula', '', 240, 0),
	(25, 'BYN', 'Belarusian Ruble', '', 250, 0),
	(26, 'BZD', 'Belizean Dollar', '', 260, 0),
	(27, 'CAD', 'Canadian Dollar', '', 270, 0),
	(28, 'CHF', 'Swiss Franc', '', 280, 0),
	(29, 'CLP', 'Chilean Peso', '', 290, 0),
	(30, 'CNY', 'Chinese Yuan Renminbi', '', 300, 0),
	(31, 'COP', 'Colombian Peso', '', 310, 0),
	(32, 'CRC', 'Costa Rican Colon', '', 320, 0),
	(33, 'CUP', 'Cuban Peso', '', 330, 0),
	(34, 'CVE', 'Cape Verdean Escudo', '', 340, 0),
	(35, 'CZK', 'Czech Koruna', '', 350, 0),
	(36, 'DEM', 'MARK', '', 360, 0),
	(37, 'DJF', 'Djiboutian Franc', '', 370, 0),
	(38, 'DKK', 'Danish Krone', '', 380, 0),
	(39, 'DOP', 'Dominican Peso', '', 390, 0),
	(40, 'DZD', 'Algerian Dinar', '', 400, 0),
	(41, 'ECS', 'SUCRE', '', 410, 0),
	(42, 'EGP', 'Egyptian Pound', '', 420, 0),
	(43, 'ESP', 'PESETA', '', 430, 0),
	(44, 'ETB', 'Ethiopian Birr', '', 440, 0),
	(45, 'EUR', 'Euro', '', 450, 0),
	(46, 'FIM', 'MARKKA', '', 460, 0),
	(47, 'FJD', 'Fijian Dollar', '', 470, 0),
	(48, 'FKP', 'Falkland Island Pound', '', 480, 0),
	(49, 'FRF', 'FRANC', '', 490, 0),
	(50, 'GAF', 'CFA Franc BEAC', '', 500, 0),
	(51, 'GBP', 'British Pound', '', 510, 0),
	(52, 'GEL', 'Georgian Lari', '', 520, 0),
	(53, 'GHS', 'Ghanaian Cedi', '', 530, 0),
	(54, 'GIP', 'Gibraltar Pound', '', 540, 0),
	(55, 'GMD', 'Gambian Dalasi', '', 550, 0),
	(56, 'GNF', 'Guinean Franc', '', 560, 0),
	(57, 'GTQ', 'Guatemalan Quetzal', '', 570, 0),
	(58, 'GYD', 'Guyanese Dollar', '', 580, 0),
	(59, 'HKD', 'Hong Kong Dollar', '', 590, 0),
	(60, 'HNL', 'Honduran Lempira', '', 600, 0),
	(61, 'HRK', 'Croatian Kuna', '', 610, 0),
	(62, 'HTG', 'Haitian Gourde', '', 620, 0),
	(63, 'HUF', 'Hungarian Forint', '', 630, 0),
	(64, 'IDR', 'Indonesian Rupiah', '', 640, 0),
	(65, 'ILS', 'Israeli Shekel', '', 650, 0),
	(66, 'INR', 'Indian Rupee', '', 660, 0),
	(67, 'IQD', 'Iraqi Dinar', '', 670, 0),
	(68, 'IRR', 'Iranian Rial', '', 680, 0),
	(69, 'ISK', 'Icelandic Krona', '', 690, 0),
	(70, 'ITL', 'LIRA', '', 700, 0),
	(71, 'JMD', 'Jamaican Dollar', '', 710, 0),
	(72, 'JOD', 'Jordanian Dinar', '', 720, 0),
	(73, 'JPY', 'Japanese Yen', '', 730, 0),
	(74, 'KES', 'Kenyan Shilling', '', 740, 0),
	(75, 'KGS', 'Kyrgyzstani Som', '', 750, 0),
	(76, 'KHR', 'Cambodian Riel', '', 760, 0),
	(77, 'KMF', 'Comorian Franc', '', 770, 0),
	(78, 'KPW', 'North Korean Won', '', 780, 0),
	(79, 'KRW', 'Korean Won', '', 790, 0),
	(80, 'KWD', 'Kuwaiti Dinar', '', 800, 0),
	(81, 'KYD', 'Caymanian Dollar', '', 810, 0),
	(82, 'KZT', 'Kazakhstani Tenge', '', 820, 0),
	(83, 'LAK', 'Lao Kip', '', 830, 0),
	(84, 'LBP', 'Lebanese Pound', '', 840, 0),
	(85, 'LKR', 'Sri Lankan Rupee', '', 850, 0),
	(86, 'LRD', 'Liberian Dollar', '', 860, 0),
	(87, 'LSL', 'Basotho Loti', '', 870, 0),
	(88, 'LYD', 'Libyan Dinar', '', 880, 0),
	(89, 'MAD', 'Moroccan Dirham', '', 890, 0),
	(90, 'MDL', 'Moldovan Leu', '', 900, 0),
	(91, 'MGA', 'Malagasy Ariary', '', 910, 0),
	(92, 'MMK', 'Burmese Kyat', '', 920, 0),
	(93, 'MNT', 'Mongolian Tughrik', '', 930, 0),
	(94, 'MOP', 'Macau Pataca', '', 940, 0),
	(95, 'MRO', 'Ouguiyas', '', 950, 0),
	(96, 'MUR', 'Mauritian Rupee', '', 960, 0),
	(97, 'MVR', 'Maldivian Rufiyaa', '', 970, 0),
	(98, 'MWK', 'Malawian Kwacha', '', 980, 0),
	(99, 'MXN', 'Mexican Peso', '', 990, 0),
	(100, 'MYR', 'Malaysian Ringgit', '', 1000, 0),
	(101, 'MZN', 'Mozambican Metical', '', 1010, 0),
	(102, 'NAD', 'Namibian Dollar', '', 1020, 0),
	(103, 'NGN', 'Nigerian Naira', '', 1030, 0),
	(104, 'NIO', 'Nicaraguan Cordoba', '', 1040, 0),
	(105, 'NLG', 'GUILDER', '', 1050, 0),
	(106, 'NOK', 'Norwegian Krone', '', 1060, 0),
	(107, 'NPR', 'Nepalese Rupee', '', 1070, 0),
	(108, 'NZD', 'New Zealand Dollar', '', 1080, 0),
	(109, 'OMR', 'Omani Rial', '', 1090, 0),
	(110, 'PAB', 'Panamanian Balboa', '', 1100, 0),
	(111, 'PEN', 'Peruvian Sol', '', 1110, 0),
	(112, 'PGK', 'Papua New Guinean Kina', '', 1120, 0),
	(113, 'PHP', 'Philippine Peso', '', 1130, 0),
	(114, 'PKR', 'Pakistani Rupee', '', 1140, 0),
	(115, 'PLN', 'Polish Zloty', '', 1150, 0),
	(116, 'PYG', 'Paraguayan Guarani', '', 1160, 0),
	(117, 'QAR', 'Qatari Riyal', '', 1170, 0),
	(118, 'RON', 'Romanian Leu', '', 1180, 0),
	(119, 'RSD', 'Serbian Dinar', '', 1190, 0),
	(120, 'RUB', 'Russian Ruble', '', 1200, 0),
	(121, 'RWF', 'Rwandan Franc', '', 1210, 0),
	(122, 'SAR', 'Saudi Arabian Riyal', '', 1220, 0),
	(123, 'SBD', 'Solomon Islander Dollar', '', 1230, 0),
	(124, 'SCR', 'Seychellois Rupee', '', 1240, 0),
	(125, 'SDG', 'Sudanese Pound', '', 1250, 0),
	(126, 'SDR', 'SPECIAL DRAWING RIGHTS(USD)', '', 1260, 0),
	(127, 'SEK', 'Swedish Krona', '', 1270, 0),
	(128, 'SGD', 'Singapore Dollar', '', 1280, 0),
	(129, 'SHP', 'Saint Helenian Pound', '', 1290, 0),
	(130, 'SKK', 'Koruny', '', 1300, 0),
	(131, 'SLL', 'Sierra Leonean Leone', '', 1310, 0),
	(132, 'SOS', 'Somali Shilling', '', 1320, 0),
	(133, 'SRD', 'Surinamese Dollar', '', 1330, 0),
	(134, 'STD', 'Dobras', '', 1340, 0),
	(135, 'SVC', 'Salvadoran Colon', '', 1350, 0),
	(136, 'SYP', 'Syrian Pound', '', 1360, 0),
	(137, 'SZL', 'Swazi Lilangeni', '', 1370, 0),
	(138, 'THB', 'Thai Baht', '', 1380, 0),
	(139, 'TMT', 'Turkmenistani Manat', '', 1390, 0),
	(140, 'TND', 'Tunisian Dinar', '', 1400, 0),
	(141, 'TOP', 'Tongan Paanga', '', 1410, 0),
	(142, 'TRY', 'Turkish Lira', '', 1420, 0),
	(143, 'TTD', 'Trinidadian Dollar', '', 1430, 0),
	(144, 'TWD', 'Taiwan New Dollar', '', 1440, 0),
	(145, 'TZS', 'Tanzanian Shilling', '', 1450, 0),
	(146, 'UAH', 'Ukrainian Hryvnia', '', 1460, 0),
	(147, 'UGX', 'Ugandan Shilling', '', 1470, 0),
	(148, 'USD', 'US Dollar', '', 1480, 0),
	(149, 'UYU', 'Uruguayan Peso', '', 1490, 0),
	(150, 'UZS', 'Uzbekistani Som', '', 1500, 0),
	(151, 'VES', 'Venezuelan Bolívar', '', 1510, 0),
	(152, 'VND', 'Vietnamese Dong', '', 1520, 0),
	(153, 'VUV', 'Ni-Vanuatu Vatu', '', 1530, 0),
	(154, 'WST', 'Samoan Tala', '', 1540, 0),
	(155, 'XAF', 'Central African CFA Franc BEAC', '', 1550, 0),
	(156, 'XCD', 'East Caribbean Dollar', '', 1560, 0),
	(157, 'XOF', 'CFA Franc', '', 1570, 0),
	(158, 'XPF', 'CFP Franc', '', 1580, 0),
	(159, 'YER', 'Yemeni Rial', '', 1590, 0),
	(160, 'ZAR', 'South African Rand', '', 1600, 0),
	(161, 'ZMW', 'Zambian Kwacha', '', 1610, 0);
/*!40000 ALTER TABLE `currencies` ENABLE KEYS */;

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
