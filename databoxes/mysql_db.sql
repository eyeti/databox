-- MySQL Administrator dump 1.4
--
-- ------------------------------------------------------
-- Server version	5.1.39-community


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;


--
-- Create schema mysql_db
--

CREATE DATABASE IF NOT EXISTS mysql_db;
USE mysql_db;

--
-- Definition of table `tblchangelog`
--

DROP TABLE IF EXISTS `tblchangelog`;
CREATE TABLE `tblchangelog` (
  `major` int(11) NOT NULL,
  `minor` int(11) NOT NULL,
  `build` int(11) NOT NULL,
  `about` varchar(50) NOT NULL,
  `dependency` varchar(255) NOT NULL,
  `changelog` text NOT NULL,
  PRIMARY KEY (`major`,`minor`,`build`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

--
-- Dumping data for table `tblchangelog`
--

/*!40000 ALTER TABLE `tblchangelog` DISABLE KEYS */;
INSERT INTO `tblchangelog` (`major`,`minor`,`build`,`about`,`dependency`,`changelog`) VALUES 
 (0,1,0,'MySql_DB first version. DB I/O with logging','MySql.Data v6.3.6','Initial release');
/*!40000 ALTER TABLE `tblchangelog` ENABLE KEYS */;




/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
