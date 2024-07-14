/*
SQLyog Ultimate v12.09 (32 bit)
MySQL - 5.7.25 : Database - modeloagua
*********************************************************************
*/

/*!40101 SET NAMES utf8 */;

/*!40101 SET SQL_MODE=''*/;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;
CREATE DATABASE /*!32312 IF NOT EXISTS*/`modeloagua` /*!40100 DEFAULT CHARACTER SET latin1 */;

USE `modeloagua`;

/*Table structure for table `alicuota` */

DROP TABLE IF EXISTS `alicuota`;

CREATE TABLE `alicuota` (
  `iva_cf` decimal(5,2) NOT NULL DEFAULT '0.00',
  `iva` decimal(5,2) NOT NULL DEFAULT '0.00',
  `rni` decimal(6,3) NOT NULL DEFAULT '0.000',
  `fecha` date NOT NULL,
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(50) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`fecha`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `anotador` */

DROP TABLE IF EXISTS `anotador`;

CREATE TABLE `anotador` (
  `anotador_id` int(4) NOT NULL AUTO_INCREMENT,
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `anotacion` text NOT NULL,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`anotador_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `categoria_socio` */

DROP TABLE IF EXISTS `categoria_socio`;

CREATE TABLE `categoria_socio` (
  `categoriasocio_id` smallint(2) NOT NULL DEFAULT '0',
  `nombre` varchar(100) NOT NULL DEFAULT '',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`categoriasocio_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `cliente` */

DROP TABLE IF EXISTS `cliente`;

CREATE TABLE `cliente` (
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `fecha_alta` date NOT NULL,
  `fecha_baja` date DEFAULT NULL,
  `apellido` varchar(50) NOT NULL DEFAULT '',
  `nombre` varchar(50) NOT NULL DEFAULT '',
  `numero_socio` varchar(10) DEFAULT NULL,
  `inmueble_calle` varchar(50) NOT NULL DEFAULT '',
  `inmueble_puerta` varchar(10) NOT NULL DEFAULT '',
  `inmueble_piso` varchar(10) NOT NULL DEFAULT '',
  `inmueble_dpto` varchar(10) NOT NULL DEFAULT '',
  `inmueble_localidad` varchar(50) NOT NULL DEFAULT '',
  `inmueble_provincia` varchar(50) NOT NULL DEFAULT '',
  `inmueble_codpostal` smallint(2) NOT NULL DEFAULT '0',
  `fiscal_calle` varchar(50) NOT NULL DEFAULT '',
  `fiscal_puerta` varchar(10) NOT NULL DEFAULT '',
  `fiscal_piso` varchar(10) NOT NULL DEFAULT '',
  `fiscal_dpto` varchar(10) NOT NULL DEFAULT '',
  `fiscal_localidad` varchar(50) NOT NULL DEFAULT '',
  `fiscal_provincia` varchar(50) NOT NULL DEFAULT '',
  `fiscal_codpostal` smallint(2) NOT NULL DEFAULT '0',
  `cuit` varchar(11) NOT NULL DEFAULT '',
  `situacion_iva` smallint(2) NOT NULL DEFAULT '0',
  `nombre_categoria` varchar(50) NOT NULL DEFAULT '',
  `categoria` smallint(2) NOT NULL DEFAULT '0',
  `servicio` smallint(2) NOT NULL DEFAULT '0',
  `cobro` smallint(2) NOT NULL DEFAULT '0',
  `zona` smallint(2) NOT NULL DEFAULT '0',
  `ruta` smallint(2) NOT NULL DEFAULT '0',
  `orden` smallint(2) NOT NULL DEFAULT '0',
  `estado_id` smallint(2) NOT NULL DEFAULT '0',
  `fecha_nacimiento` date DEFAULT NULL,
  `categoriasocio_id` smallint(2) NOT NULL DEFAULT '0',
  `destino_id` smallint(2) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` datetime DEFAULT NULL,
  `updated` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`cliente_id`,`fecha_alta`),
  KEY `auto_id` (`auto_id`),
  KEY `apellido` (`apellido`,`nombre`),
  KEY `cliente_id` (`cliente_id`),
  KEY `fecha_alta` (`fecha_alta`),
  KEY `fecha_baja` (`fecha_baja`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `cliente_notificacion` */

DROP TABLE IF EXISTS `cliente_notificacion`;

CREATE TABLE `cliente_notificacion` (
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `ultima_notificacion_15` date DEFAULT NULL,
  `ultima_notificacion_48` date DEFAULT NULL,
  `ultima_notificacion_corte` date DEFAULT NULL,
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` datetime NOT NULL,
  `updated` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `uuid` varchar(32) NOT NULL DEFAULT '',
  PRIMARY KEY (`cliente_id`),
  KEY `auto_id` (`auto_id`),
  KEY `uuid` (`uuid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `cliente_volumen` */

DROP TABLE IF EXISTS `cliente_volumen`;

CREATE TABLE `cliente_volumen` (
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `periodo_id` smallint(2) NOT NULL DEFAULT '0',
  `medidor_id_actual` varchar(20) NOT NULL DEFAULT '',
  `estado_actual` int(4) NOT NULL DEFAULT '0',
  `medidor_id_anterior` varchar(20) NOT NULL DEFAULT '',
  `estado_anterior` int(4) NOT NULL DEFAULT '0',
  `consumido` int(4) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`cliente_id`,`periodo_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `clientedato` */

DROP TABLE IF EXISTS `clientedato`;

CREATE TABLE `clientedato` (
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `documento` decimal(10,0) NOT NULL DEFAULT '0',
  `email` varchar(150) NOT NULL DEFAULT '',
  `fijo` varchar(100) NOT NULL DEFAULT '',
  `celular` varchar(100) NOT NULL DEFAULT '',
  `uid` varchar(50) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`cliente_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `compafip` */

DROP TABLE IF EXISTS `compafip`;

CREATE TABLE `compafip` (
  `comprobante_id` smallint(2) NOT NULL DEFAULT '0',
  `nombre` varchar(150) NOT NULL DEFAULT '',
  `label` varchar(150) NOT NULL DEFAULT '',
  `uid` varchar(50) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`comprobante_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `cuota` */

DROP TABLE IF EXISTS `cuota`;

CREATE TABLE `cuota` (
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `plan_id` smallint(2) NOT NULL DEFAULT '0',
  `cuota_id` smallint(2) NOT NULL DEFAULT '0',
  `fecha_vencimiento` date NOT NULL,
  `fecha_pago` date DEFAULT NULL,
  `importe` decimal(16,2) NOT NULL DEFAULT '0.00',
  `cancelada` tinyint(1) NOT NULL DEFAULT '0',
  `plan_id_cancela` smallint(2) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`cliente_id`,`plan_id`,`cuota_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `desconexion` */

DROP TABLE IF EXISTS `desconexion`;

CREATE TABLE `desconexion` (
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `fecha_desconexion` date NOT NULL,
  `fecha_reconexion` date DEFAULT NULL,
  `motivo` varchar(100) NOT NULL DEFAULT '',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(50) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`cliente_id`,`fecha_desconexion`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `destino_servicio` */

DROP TABLE IF EXISTS `destino_servicio`;

CREATE TABLE `destino_servicio` (
  `destino_id` smallint(2) NOT NULL DEFAULT '0',
  `nombre` varchar(100) NOT NULL DEFAULT '',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`destino_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `detalle` */

DROP TABLE IF EXISTS `detalle`;

CREATE TABLE `detalle` (
  `prefijo_id` smallint(2) NOT NULL DEFAULT '0',
  `factura_id` int(4) NOT NULL DEFAULT '0',
  `rubro_id` smallint(2) NOT NULL DEFAULT '0',
  `concepto` varchar(100) NOT NULL DEFAULT '',
  `cantidad` decimal(16,2) NOT NULL DEFAULT '0.00',
  `precio_unitario` decimal(16,2) NOT NULL DEFAULT '0.00',
  `iva` tinyint(1) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(50) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`prefijo_id`,`factura_id`,`rubro_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `deuda` */

DROP TABLE IF EXISTS `deuda`;

CREATE TABLE `deuda` (
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `plan_id` smallint(2) NOT NULL DEFAULT '0',
  `deuda` decimal(16,2) NOT NULL DEFAULT '0.00',
  `cuotas` smallint(2) NOT NULL DEFAULT '0',
  `cuotas_pagadas` smallint(2) NOT NULL DEFAULT '0',
  `tasa` decimal(6,4) NOT NULL DEFAULT '0.0000',
  `pagado` tinyint(1) NOT NULL DEFAULT '0',
  `periodo` smallint(2) NOT NULL DEFAULT '0',
  `cancelada` tinyint(1) NOT NULL DEFAULT '0',
  `plan_id_cancela` smallint(2) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`cliente_id`,`plan_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `estado` */

DROP TABLE IF EXISTS `estado`;

CREATE TABLE `estado` (
  `estado_id` smallint(2) NOT NULL DEFAULT '0',
  `nombre` varchar(100) NOT NULL DEFAULT '',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`estado_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `factura` */

DROP TABLE IF EXISTS `factura`;

CREATE TABLE `factura` (
  `prefijo_id` smallint(2) NOT NULL DEFAULT '0',
  `factura_id` int(4) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `periodo_id` smallint(2) NOT NULL DEFAULT '0',
  `situacion_iva` smallint(2) NOT NULL DEFAULT '0',
  `tasa` decimal(6,4) NOT NULL DEFAULT '0.0000',
  `descuento` decimal(16,2) NOT NULL DEFAULT '0.00',
  `pagada` tinyint(1) NOT NULL DEFAULT '0',
  `fecha_pago` date DEFAULT NULL,
  `tipo_id` smallint(2) NOT NULL DEFAULT '0',
  `anulada` tinyint(1) NOT NULL DEFAULT '0',
  `total` decimal(16,2) NOT NULL DEFAULT '0.00',
  `interes` decimal(16,2) NOT NULL DEFAULT '0.00',
  `letras` varchar(200) NOT NULL DEFAULT '',
  `prefijo_id_interes` smallint(2) NOT NULL DEFAULT '0',
  `factura_id_interes` int(4) NOT NULL DEFAULT '0',
  `iva_cf` decimal(16,2) NOT NULL DEFAULT '0.00',
  `iva_ri` decimal(16,2) NOT NULL DEFAULT '0.00',
  `iva_rn` decimal(16,2) NOT NULL DEFAULT '0.00',
  `periodo_id_fin` smallint(2) NOT NULL DEFAULT '0',
  `cancelada` tinyint(1) NOT NULL DEFAULT '0',
  `plan_id_cancela` smallint(2) DEFAULT NULL,
  `pf_codigo` varchar(50) NOT NULL DEFAULT '',
  `pf_barras` varchar(50) NOT NULL DEFAULT '',
  `cajamovimiento_id` int(4) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(50) NOT NULL DEFAULT '',
  `created` datetime DEFAULT NULL,
  `updated` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`prefijo_id`,`factura_id`),
  KEY `auto_id` (`auto_id`),
  KEY `cliente_id` (`cliente_id`,`periodo_id`),
  KEY `anulada` (`anulada`),
  KEY `fecha_pago` (`fecha_pago`),
  KEY `periodo_id` (`periodo_id`),
  KEY `cajamovimiento_id` (`cajamovimiento_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `fedetalle` */

DROP TABLE IF EXISTS `fedetalle`;

CREATE TABLE `fedetalle` (
  `tipo_id` smallint(2) NOT NULL DEFAULT '0',
  `prefijo` smallint(2) NOT NULL DEFAULT '0',
  `numero` int(4) NOT NULL DEFAULT '0',
  `item` smallint(2) NOT NULL DEFAULT '0',
  `rubro_id` smallint(2) NOT NULL DEFAULT '0',
  `cantidad` decimal(10,2) NOT NULL DEFAULT '0.00',
  `unitario_sin_iva` decimal(16,2) NOT NULL DEFAULT '0.00',
  `unitario_con_iva` decimal(16,2) NOT NULL DEFAULT '0.00',
  `iva105` tinyint(1) NOT NULL DEFAULT '0',
  `exento` tinyint(1) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `concepto` varchar(200) NOT NULL DEFAULT '',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`tipo_id`,`prefijo`,`numero`,`item`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `fefactura` */

DROP TABLE IF EXISTS `fefactura`;

CREATE TABLE `fefactura` (
  `tipo_id` smallint(2) NOT NULL DEFAULT '0',
  `prefijo` smallint(2) NOT NULL DEFAULT '0',
  `numero` int(4) NOT NULL DEFAULT '0',
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `importe` decimal(16,2) NOT NULL DEFAULT '0.00',
  `neto27` decimal(16,2) NOT NULL DEFAULT '0.00',
  `neto` decimal(16,2) NOT NULL DEFAULT '0.00',
  `iva27` decimal(16,2) NOT NULL DEFAULT '0.00',
  `iva` decimal(16,2) NOT NULL DEFAULT '0.00',
  `exento` decimal(16,2) NOT NULL DEFAULT '0.00',
  `recibo` tinyint(1) NOT NULL DEFAULT '0',
  `anulada` tinyint(1) NOT NULL DEFAULT '0',
  `tipo_compro` varchar(1) NOT NULL DEFAULT '',
  `letras` varchar(200) NOT NULL DEFAULT '',
  `observaciones` varchar(300) NOT NULL DEFAULT '',
  `cae` varchar(50) NOT NULL DEFAULT '',
  `cae_vencimiento` varchar(20) NOT NULL DEFAULT '',
  `cae_barras` varchar(50) NOT NULL DEFAULT '',
  `prefijo_id` smallint(2) NOT NULL DEFAULT '0',
  `factura_id` int(4) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(50) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`tipo_id`,`prefijo`,`numero`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `imputado` */

DROP TABLE IF EXISTS `imputado`;

CREATE TABLE `imputado` (
  `serie_id` smallint(2) NOT NULL DEFAULT '0',
  `numero_id` int(4) NOT NULL DEFAULT '0',
  `tipo_id` smallint(2) NOT NULL DEFAULT '0',
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `comp_serie_id` smallint(2) NOT NULL DEFAULT '0',
  `comp_numero_id` int(4) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`serie_id`,`numero_id`,`tipo_id`,`cliente_id`,`comp_serie_id`,`comp_numero_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `lectura` */

DROP TABLE IF EXISTS `lectura`;

CREATE TABLE `lectura` (
  `medidor_id` varchar(20) NOT NULL DEFAULT '',
  `periodo_id` smallint(2) NOT NULL DEFAULT '0',
  `fecha_lectura` date NOT NULL,
  `estado` int(4) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(50) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`medidor_id`,`periodo_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `libro_socio` */

DROP TABLE IF EXISTS `libro_socio`;

CREATE TABLE `libro_socio` (
  `numero_socio` smallint(2) NOT NULL DEFAULT '0',
  `anho` smallint(2) NOT NULL DEFAULT '0',
  `nombre_apellido` varchar(150) NOT NULL DEFAULT '',
  `domicilio` varchar(150) NOT NULL DEFAULT '',
  `documento` decimal(10,0) NOT NULL DEFAULT '0',
  `estado` varchar(50) NOT NULL DEFAULT '',
  `edad` smallint(2) NOT NULL DEFAULT '0',
  `categoria` varchar(50) NOT NULL DEFAULT '',
  `ingreso` date NOT NULL,
  `enero` date DEFAULT NULL,
  `febrero` date DEFAULT NULL,
  `marzo` date DEFAULT NULL,
  `abril` date DEFAULT NULL,
  `mayo` date DEFAULT NULL,
  `junio` date DEFAULT NULL,
  `julio` date DEFAULT NULL,
  `agosto` date DEFAULT NULL,
  `setiembre` date DEFAULT NULL,
  `octubre` date DEFAULT NULL,
  `noviembre` date DEFAULT NULL,
  `diciembre` date DEFAULT NULL,
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`numero_socio`,`anho`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `listado` */

DROP TABLE IF EXISTS `listado`;

CREATE TABLE `listado` (
  `c1` varchar(150) NOT NULL DEFAULT '',
  `c2` varchar(150) NOT NULL DEFAULT '',
  `c3` varchar(150) NOT NULL DEFAULT '',
  `c4` varchar(150) NOT NULL DEFAULT '',
  `c5` varchar(150) NOT NULL DEFAULT '',
  `c6` varchar(150) NOT NULL DEFAULT '',
  `n1` decimal(16,2) NOT NULL DEFAULT '0.00',
  `n2` decimal(16,2) NOT NULL DEFAULT '0.00',
  `n3` decimal(16,2) NOT NULL DEFAULT '0.00',
  `n4` decimal(16,2) NOT NULL DEFAULT '0.00',
  `n5` decimal(16,2) NOT NULL DEFAULT '0.00',
  `n6` decimal(16,2) NOT NULL DEFAULT '0.00',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `medicion` */

DROP TABLE IF EXISTS `medicion`;

CREATE TABLE `medicion` (
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `periodo_id` smallint(2) NOT NULL DEFAULT '0',
  `medidor_id` varchar(20) NOT NULL DEFAULT '',
  `fecha_lectura` date DEFAULT NULL,
  `estado` int(4) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(50) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`cliente_id`,`periodo_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `medidor` */

DROP TABLE IF EXISTS `medidor`;

CREATE TABLE `medidor` (
  `medidor_id` varchar(20) NOT NULL DEFAULT '',
  `fecha_alta` datetime NOT NULL,
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `fecha_colocacion` date DEFAULT NULL,
  `fecha_retiro` date DEFAULT NULL,
  `motivo_retiro` smallint(2) NOT NULL DEFAULT '0',
  `estado_inicio` int(4) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`medidor_id`,`fecha_alta`,`cliente_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `ncredito` */

DROP TABLE IF EXISTS `ncredito`;

CREATE TABLE `ncredito` (
  `serie_id` smallint(2) NOT NULL DEFAULT '0',
  `numero` int(4) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `situacion_iva` smallint(2) NOT NULL DEFAULT '0',
  `anulado` tinyint(1) NOT NULL DEFAULT '0',
  `total` decimal(16,2) NOT NULL DEFAULT '0.00',
  `prefijo_id` smallint(2) NOT NULL DEFAULT '0',
  `factura_id` int(4) NOT NULL DEFAULT '0',
  `iva_cf` decimal(16,2) NOT NULL DEFAULT '0.00',
  `iva_ri` decimal(16,2) NOT NULL DEFAULT '0.00',
  `iva_rn` decimal(16,2) NOT NULL DEFAULT '0.00',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`serie_id`,`numero`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `notificacion` */

DROP TABLE IF EXISTS `notificacion`;

CREATE TABLE `notificacion` (
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `tiponotificacion_id` smallint(2) NOT NULL DEFAULT '0',
  `vencimiento` date DEFAULT NULL,
  `notificacion_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` datetime DEFAULT NULL,
  `updated` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `uuid` varchar(32) NOT NULL DEFAULT '',
  PRIMARY KEY (`cliente_id`,`fecha`),
  KEY `notificacion_id` (`notificacion_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `novedad` */

DROP TABLE IF EXISTS `novedad`;

CREATE TABLE `novedad` (
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `periodo_id` smallint(2) NOT NULL DEFAULT '0',
  `rubro_id` smallint(2) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `porcentaje` decimal(6,2) NOT NULL DEFAULT '0.00',
  `cantidad` decimal(10,2) NOT NULL DEFAULT '0.00',
  `importe` decimal(16,2) NOT NULL DEFAULT '0.00',
  `veces` smallint(2) NOT NULL DEFAULT '0',
  `veces_cobradas` smallint(2) NOT NULL DEFAULT '0',
  `indefinida` tinyint(1) NOT NULL DEFAULT '0',
  `periodo_id_suspension` smallint(2) DEFAULT NULL,
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`cliente_id`,`periodo_id`,`rubro_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `operador` */

DROP TABLE IF EXISTS `operador`;

CREATE TABLE `operador` (
  `operador_id` smallint(2) NOT NULL DEFAULT '0',
  `razon_social` varchar(150) NOT NULL DEFAULT '',
  `calle` varchar(50) NOT NULL DEFAULT '',
  `puerta` varchar(20) NOT NULL DEFAULT '',
  `piso` varchar(10) NOT NULL DEFAULT '',
  `dpto` varchar(10) NOT NULL DEFAULT '',
  `codigo_postal` smallint(2) NOT NULL DEFAULT '0',
  `localidad` varchar(50) NOT NULL DEFAULT '',
  `provincia` varchar(50) NOT NULL DEFAULT '',
  `telefono` varchar(50) NOT NULL DEFAULT '',
  `cuit` varchar(11) NOT NULL DEFAULT '',
  `ingresos_brutos` varchar(20) NOT NULL DEFAULT '',
  `situacion_iva` smallint(2) NOT NULL DEFAULT '0',
  `numero_epas` varchar(10) NOT NULL DEFAULT '',
  `fecha_inicio` date NOT NULL,
  `servicio` smallint(2) NOT NULL DEFAULT '0',
  `prefijo_id` smallint(2) NOT NULL DEFAULT '0',
  `factura_id` int(4) NOT NULL DEFAULT '0',
  `periodo_factura` smallint(2) NOT NULL DEFAULT '0',
  `resolucion` varchar(20) NOT NULL DEFAULT '',
  `personeria` varchar(20) NOT NULL DEFAULT '',
  `recibo_serie` smallint(2) NOT NULL DEFAULT '0',
  `recibo` int(4) NOT NULL DEFAULT '0',
  `ncredito_serie` smallint(2) NOT NULL DEFAULT '0',
  `ncredito` int(4) NOT NULL DEFAULT '0',
  `cai` varchar(25) DEFAULT NULL,
  `cai_vencimiento` date DEFAULT NULL,
  `preimpreso` tinyint(1) NOT NULL DEFAULT '0',
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`operador_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `parametro` */

DROP TABLE IF EXISTS `parametro`;

CREATE TABLE `parametro` (
  `parametro_id` int(4) NOT NULL AUTO_INCREMENT,
  `fe_produccion` tinyint(1) NOT NULL DEFAULT '0',
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`parametro_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `periodo` */

DROP TABLE IF EXISTS `periodo`;

CREATE TABLE `periodo` (
  `periodo_id` smallint(2) NOT NULL DEFAULT '0',
  `descripcion` varchar(30) NOT NULL DEFAULT '',
  `fecha_inicio` date NOT NULL,
  `fecha_fin` date NOT NULL,
  `fecha_primero` date NOT NULL,
  `fecha_segundo` date NOT NULL,
  `tasa` decimal(6,4) NOT NULL DEFAULT '0.0000',
  `leyenda` varchar(300) NOT NULL DEFAULT '',
  `liquidado` decimal(16,2) NOT NULL DEFAULT '0.00',
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`periodo_id`),
  KEY `fecha_fin` (`fecha_fin`),
  KEY `fecha_inicio` (`fecha_inicio`),
  KEY `fecha_primero` (`fecha_primero`),
  KEY `fecha_segundo` (`fecha_segundo`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `pffile` */

DROP TABLE IF EXISTS `pffile`;

CREATE TABLE `pffile` (
  `file_name` varchar(50) NOT NULL DEFAULT '',
  `path` varchar(150) NOT NULL DEFAULT '',
  `import` tinyint(1) NOT NULL DEFAULT '0',
  `fecha_import` date DEFAULT NULL,
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`file_name`),
  KEY `auto_id` (`auto_id`),
  KEY `file_name` (`file_name`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `pfrecord` */

DROP TABLE IF EXISTS `pfrecord`;

CREATE TABLE `pfrecord` (
  `file_name` varchar(50) NOT NULL DEFAULT '',
  `line` varchar(150) NOT NULL DEFAULT '',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`auto_id`),
  KEY `file_name` (`file_name`,`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `proveedor` */

DROP TABLE IF EXISTS `proveedor`;

CREATE TABLE `proveedor` (
  `proveedor_id` int(4) NOT NULL DEFAULT '0',
  `razon_social` varchar(100) NOT NULL DEFAULT '',
  `nombre_fantasia` varchar(100) NOT NULL DEFAULT '',
  `cuit` varchar(20) NOT NULL DEFAULT '',
  `domicilio` varchar(50) NOT NULL DEFAULT '',
  `telefono` varchar(20) NOT NULL DEFAULT '',
  `fax` varchar(20) NOT NULL DEFAULT '',
  `email` varchar(50) NOT NULL DEFAULT '',
  `posicion_iva` smallint(2) NOT NULL DEFAULT '0',
  `celular` varchar(25) NOT NULL DEFAULT '',
  `ingresos_brutos` varchar(20) NOT NULL DEFAULT '',
  `contacto` varchar(250) NOT NULL DEFAULT '',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`proveedor_id`),
  KEY `cuit` (`cuit`),
  KEY `razon_social` (`razon_social`),
  KEY `posicion_iva` (`posicion_iva`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `proveedor_movimiento` */

DROP TABLE IF EXISTS `proveedor_movimiento`;

CREATE TABLE `proveedor_movimiento` (
  `proveedor_id` int(4) NOT NULL DEFAULT '0',
  `comprobante_id` smallint(2) NOT NULL DEFAULT '0',
  `prefijo` smallint(2) NOT NULL DEFAULT '0',
  `nro_comprobante` int(4) NOT NULL DEFAULT '0',
  `empresa_id` smallint(2) NOT NULL DEFAULT '0',
  `fecha_comprobante` date NOT NULL,
  `fecha_vencimiento` date DEFAULT NULL,
  `total` decimal(16,2) NOT NULL DEFAULT '0.00',
  `total_cancelado` decimal(16,2) NOT NULL DEFAULT '0.00',
  `neto` decimal(16,2) NOT NULL DEFAULT '0.00',
  `importe_iva1` decimal(16,2) NOT NULL DEFAULT '0.00',
  `importe_iva2` decimal(16,2) NOT NULL DEFAULT '0.00',
  `importe_iva3` decimal(16,2) NOT NULL DEFAULT '0.00',
  `percepcion_iva` decimal(16,2) NOT NULL DEFAULT '0.00',
  `percepcion_iibb` decimal(16,2) NOT NULL DEFAULT '0.00',
  `gastos_no_gravados` decimal(16,2) NOT NULL DEFAULT '0.00',
  `ajustes` decimal(16,2) NOT NULL DEFAULT '0.00',
  `monotributo` tinyint(1) NOT NULL DEFAULT '0',
  `cuenta_movimiento_id` bigint(20) NOT NULL DEFAULT '0',
  `observaciones` text NOT NULL,
  `proveedor_movimiento_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`proveedor_id`,`comprobante_id`,`prefijo`,`nro_comprobante`),
  KEY `proveedor_id` (`proveedor_id`,`comprobante_id`,`prefijo`,`nro_comprobante`),
  KEY `fecha_vencimiento` (`fecha_vencimiento`),
  KEY `proveedor_id_2` (`proveedor_id`,`comprobante_id`,`fecha_comprobante`,`prefijo`,`nro_comprobante`,`total`),
  KEY `comprobante_id` (`comprobante_id`),
  KEY `cuenta_movimiento_id` (`cuenta_movimiento_id`),
  KEY `proveedor_movimiento_id` (`proveedor_movimiento_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `proveedor_pago` */

DROP TABLE IF EXISTS `proveedor_pago`;

CREATE TABLE `proveedor_pago` (
  `proveedor_movimiento_id_deuda` int(4) NOT NULL DEFAULT '0',
  `proveedor_movimiento_id_aplicado` int(4) NOT NULL DEFAULT '0',
  `importe_aplicado` decimal(16,2) NOT NULL DEFAULT '0.00',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`proveedor_movimiento_id_deuda`,`proveedor_movimiento_id_aplicado`),
  KEY `proveedor_movimiento_id_deuda` (`proveedor_movimiento_id_deuda`),
  KEY `proveedor_movimiento_id_aplicado` (`proveedor_movimiento_id_aplicado`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `proveedor_saldo` */

DROP TABLE IF EXISTS `proveedor_saldo`;

CREATE TABLE `proveedor_saldo` (
  `proveedor_id` int(4) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `saldo` decimal(16,2) NOT NULL,
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`proveedor_id`,`fecha`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `rango` */

DROP TABLE IF EXISTS `rango`;

CREATE TABLE `rango` (
  `categoria` smallint(2) NOT NULL DEFAULT '0',
  `rango_id` smallint(2) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `limite_inferior` decimal(10,2) NOT NULL DEFAULT '0.00',
  `limite_superior` decimal(10,2) NOT NULL DEFAULT '0.00',
  `tarifa` decimal(10,2) NOT NULL DEFAULT '0.00',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`categoria`,`rango_id`,`fecha`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `recibo` */

DROP TABLE IF EXISTS `recibo`;

CREATE TABLE `recibo` (
  `serie_id` smallint(2) NOT NULL DEFAULT '0',
  `numero` int(4) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `situacion_iva` smallint(2) NOT NULL DEFAULT '0',
  `anulado` tinyint(1) NOT NULL DEFAULT '0',
  `total` decimal(16,2) NOT NULL DEFAULT '0.00',
  `imputado` tinyint(1) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`serie_id`,`numero`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `registrocae` */

DROP TABLE IF EXISTS `registrocae`;

CREATE TABLE `registrocae` (
  `tipo_id` smallint(2) NOT NULL DEFAULT '0',
  `prefijo` smallint(2) NOT NULL DEFAULT '0',
  `numero` int(4) NOT NULL DEFAULT '0',
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `total` decimal(16,2) NOT NULL DEFAULT '0.00',
  `exento` decimal(16,2) NOT NULL DEFAULT '0.00',
  `neto27` decimal(16,2) NOT NULL DEFAULT '0.00',
  `neto` decimal(16,2) NOT NULL DEFAULT '0.00',
  `neto105` decimal(16,2) NOT NULL DEFAULT '0.00',
  `iva27` decimal(16,2) NOT NULL DEFAULT '0.00',
  `iva` decimal(16,2) NOT NULL DEFAULT '0.00',
  `iva105` decimal(16,2) NOT NULL DEFAULT '0.00',
  `cae` varchar(30) NOT NULL DEFAULT '',
  `fecha` varchar(20) NOT NULL DEFAULT '',
  `cae_vencimiento` varchar(20) NOT NULL DEFAULT '',
  `barras` varchar(50) NOT NULL DEFAULT '',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`tipo_id`,`prefijo`,`numero`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `rubro` */

DROP TABLE IF EXISTS `rubro`;

CREATE TABLE `rubro` (
  `rubro_id` smallint(2) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `rango_id` smallint(2) NOT NULL DEFAULT '0',
  `concepto` varchar(100) NOT NULL DEFAULT '',
  `precio_unitario` decimal(16,2) NOT NULL DEFAULT '0.00',
  `iva` tinyint(1) NOT NULL DEFAULT '0',
  `comun` tinyint(1) NOT NULL DEFAULT '0',
  `comun_socio` tinyint(1) NOT NULL DEFAULT '0',
  `cobro` smallint(2) NOT NULL DEFAULT '0',
  `desconectado` tinyint(1) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`rubro_id`,`fecha`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `rubrovigente` */

DROP TABLE IF EXISTS `rubrovigente`;

CREATE TABLE `rubrovigente` (
  `rubro_id` smallint(2) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `rango_id` smallint(2) NOT NULL DEFAULT '0',
  `concepto` varchar(100) NOT NULL DEFAULT '',
  `precio_unitario` decimal(16,2) NOT NULL DEFAULT '0.00',
  `iva` tinyint(1) NOT NULL DEFAULT '0',
  `comun` tinyint(1) NOT NULL DEFAULT '0',
  `comun_socio` tinyint(1) NOT NULL DEFAULT '0',
  `cobro` smallint(2) NOT NULL DEFAULT '0',
  `desconectado` tinyint(1) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`rubro_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `suspension` */

DROP TABLE IF EXISTS `suspension`;

CREATE TABLE `suspension` (
  `tipo` varchar(5) NOT NULL DEFAULT '',
  `numero` int(4) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `periodo_id` smallint(2) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`tipo`,`numero`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `suspfactura` */

DROP TABLE IF EXISTS `suspfactura`;

CREATE TABLE `suspfactura` (
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `periodo_id_inicio` smallint(2) NOT NULL DEFAULT '0',
  `periodo_id_fin` smallint(2) DEFAULT NULL,
  `motivo` varchar(100) NOT NULL DEFAULT '',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`cliente_id`,`periodo_id_inicio`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `tipo_notificacion` */

DROP TABLE IF EXISTS `tipo_notificacion`;

CREATE TABLE `tipo_notificacion` (
  `tiponotificacion_id` smallint(2) NOT NULL AUTO_INCREMENT,
  `nombre` varchar(50) NOT NULL DEFAULT '',
  `valor_socio` decimal(10,2) NOT NULL DEFAULT '0.00',
  `valor_no_socio` decimal(10,2) NOT NULL DEFAULT '0.00',
  `created` datetime DEFAULT NULL,
  `updated` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `uuid` varchar(32) NOT NULL DEFAULT '',
  PRIMARY KEY (`tiponotificacion_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `tipo_transaccion` */

DROP TABLE IF EXISTS `tipo_transaccion`;

CREATE TABLE `tipo_transaccion` (
  `tipo_transaccion_id` smallint(2) NOT NULL DEFAULT '0',
  `descripcion` varchar(100) NOT NULL DEFAULT '',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`tipo_transaccion_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `tipocomprobante` */

DROP TABLE IF EXISTS `tipocomprobante`;

CREATE TABLE `tipocomprobante` (
  `tipo_id` smallint(2) NOT NULL AUTO_INCREMENT,
  `descripcion` varchar(200) NOT NULL DEFAULT '',
  `modulo` smallint(2) NOT NULL DEFAULT '0',
  `aplica_pendiente` tinyint(1) NOT NULL DEFAULT '0',
  `cuenta_corriente` tinyint(1) NOT NULL DEFAULT '0',
  `debita` tinyint(1) NOT NULL DEFAULT '0',
  `iva` tinyint(1) NOT NULL DEFAULT '0',
  `aplicable` tinyint(1) NOT NULL DEFAULT '0',
  `libro_iva` tinyint(1) NOT NULL DEFAULT '0',
  `tipo_comprobante` varchar(1) NOT NULL DEFAULT '',
  `recibo` tinyint(1) NOT NULL DEFAULT '0',
  `contado` tinyint(1) NOT NULL DEFAULT '0',
  `punto_venta` smallint(2) NOT NULL DEFAULT '0',
  `comprobante_id` smallint(2) NOT NULL DEFAULT '0',
  `factura_electronica` tinyint(1) NOT NULL DEFAULT '0',
  `uid` varchar(20) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`tipo_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `tipopago` */

DROP TABLE IF EXISTS `tipopago`;

CREATE TABLE `tipopago` (
  `tipo_id` smallint(2) NOT NULL DEFAULT '0',
  `nombre` varchar(20) NOT NULL DEFAULT '',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `uid` varchar(50) NOT NULL DEFAULT '',
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`tipo_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `valor` */

DROP TABLE IF EXISTS `valor`;

CREATE TABLE `valor` (
  `valor_id` smallint(2) NOT NULL DEFAULT '0',
  `concepto` varchar(50) NOT NULL DEFAULT '',
  `cuenta` int(4) NOT NULL DEFAULT '0',
  `numerable` tinyint(1) NOT NULL DEFAULT '0',
  `fecha_emision` tinyint(1) NOT NULL DEFAULT '0',
  `fecha_vencimiento` tinyint(1) NOT NULL DEFAULT '0',
  `titular` tinyint(1) NOT NULL DEFAULT '0',
  `banco` tinyint(1) NOT NULL DEFAULT '0',
  `cheque_tercero` tinyint(1) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`valor_id`),
  KEY `auto_id` (`auto_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `valor_movimiento` */

DROP TABLE IF EXISTS `valor_movimiento`;

CREATE TABLE `valor_movimiento` (
  `valor_id` smallint(2) NOT NULL DEFAULT '0',
  `empresa_id` smallint(2) NOT NULL DEFAULT '1',
  `cliente_id` int(4) NOT NULL DEFAULT '0',
  `proveedor_id` int(4) NOT NULL DEFAULT '0',
  `fecha` date NOT NULL,
  `comprobante_id` smallint(2) NOT NULL DEFAULT '0',
  `numero` int(4) NOT NULL DEFAULT '0',
  `importe` decimal(16,2) NOT NULL DEFAULT '0.00',
  `cuenta` int(4) NOT NULL DEFAULT '0',
  `cliente_movimiento_id` int(4) NOT NULL DEFAULT '0',
  `proveedor_movimiento_id` int(4) NOT NULL DEFAULT '0',
  `caja_movimiento_id` int(4) NOT NULL DEFAULT '0',
  `titular` varchar(50) NOT NULL DEFAULT '',
  `banco` varchar(50) NOT NULL DEFAULT '',
  `fecha_emision` date DEFAULT NULL,
  `fecha_vencimiento` date DEFAULT NULL,
  `cuenta_movimiento_id` bigint(20) NOT NULL DEFAULT '0',
  `auto_id` int(4) NOT NULL AUTO_INCREMENT,
  `created` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`auto_id`),
  KEY `valor_id` (`valor_id`),
  KEY `cliente_id` (`cliente_id`),
  KEY `auto_id` (`auto_id`),
  KEY `cliente_movimiento_id` (`cliente_movimiento_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

/*Table structure for table `vw_clicor` */

DROP TABLE IF EXISTS `vw_clicor`;

/*!50001 DROP VIEW IF EXISTS `vw_clicor` */;
/*!50001 DROP TABLE IF EXISTS `vw_clicor` */;

/*!50001 CREATE TABLE  `vw_clicor`(
 `cliente_id` int(4) ,
 `apellido` varchar(50) ,
 `nombre` varchar(50) ,
 `inmueble_calle` varchar(50) ,
 `inmueble_puerta` varchar(10) ,
 `inmueble_piso` varchar(10) ,
 `inmueble_dpto` varchar(10) ,
 `cobro` smallint(2) 
)*/;

/*Table structure for table `vw_clihab` */

DROP TABLE IF EXISTS `vw_clihab`;

/*!50001 DROP VIEW IF EXISTS `vw_clihab` */;
/*!50001 DROP TABLE IF EXISTS `vw_clihab` */;

/*!50001 CREATE TABLE  `vw_clihab`(
 `cliente_id` int(4) ,
 `medidor_id` varchar(20) ,
 `fecha_retiro` date 
)*/;

/*Table structure for table `vw_climed` */

DROP TABLE IF EXISTS `vw_climed`;

/*!50001 DROP VIEW IF EXISTS `vw_climed` */;
/*!50001 DROP TABLE IF EXISTS `vw_climed` */;

/*!50001 CREATE TABLE  `vw_climed`(
 `cliente_id` int(4) ,
 `fecha` datetime ,
 `medidor_id` varchar(20) 
)*/;

/*Table structure for table `vw_clires` */

DROP TABLE IF EXISTS `vw_clires`;

/*!50001 DROP VIEW IF EXISTS `vw_clires` */;
/*!50001 DROP TABLE IF EXISTS `vw_clires` */;

/*!50001 CREATE TABLE  `vw_clires`(
 `cliente_id` int(4) ,
 `apellido` varchar(50) ,
 `nombre` varchar(50) ,
 `inmueble_calle` varchar(50) ,
 `inmueble_puerta` varchar(10) ,
 `inmueble_piso` varchar(10) ,
 `inmueble_dpto` varchar(10) ,
 `cobro` smallint(2) 
)*/;

/*Table structure for table `vw_corcue` */

DROP TABLE IF EXISTS `vw_corcue`;

/*!50001 DROP VIEW IF EXISTS `vw_corcue` */;
/*!50001 DROP TABLE IF EXISTS `vw_corcue` */;

/*!50001 CREATE TABLE  `vw_corcue`(
 `cliente_id` int(4) ,
 `cantidad` bigint(21) 
)*/;

/*Table structure for table `vw_corsel` */

DROP TABLE IF EXISTS `vw_corsel`;

/*!50001 DROP VIEW IF EXISTS `vw_corsel` */;
/*!50001 DROP TABLE IF EXISTS `vw_corsel` */;

/*!50001 CREATE TABLE  `vw_corsel`(
 `prefijo_id` smallint(2) ,
 `factura_id` int(4) ,
 `cliente_id` int(4) ,
 `fecha_primero` date ,
 `numero_id` int(4) 
)*/;

/*Table structure for table `vw_factpendientes` */

DROP TABLE IF EXISTS `vw_factpendientes`;

/*!50001 DROP VIEW IF EXISTS `vw_factpendientes` */;
/*!50001 DROP TABLE IF EXISTS `vw_factpendientes` */;

/*!50001 CREATE TABLE  `vw_factpendientes`(
 `cliente_id` int(4) ,
 `fecha` date ,
 `periodo_id` smallint(2) ,
 `prefijo_id` smallint(2) ,
 `factura_id` int(4) ,
 `total` decimal(16,2) ,
 `total_final` decimal(39,2) 
)*/;

/*Table structure for table `vw_liqperiodo` */

DROP TABLE IF EXISTS `vw_liqperiodo`;

/*!50001 DROP VIEW IF EXISTS `vw_liqperiodo` */;
/*!50001 DROP TABLE IF EXISTS `vw_liqperiodo` */;

/*!50001 CREATE TABLE  `vw_liqperiodo`(
 `periodo_id` smallint(2) ,
 `liquidado` decimal(38,2) 
)*/;

/*Table structure for table `vw_medlist` */

DROP TABLE IF EXISTS `vw_medlist`;

/*!50001 DROP VIEW IF EXISTS `vw_medlist` */;
/*!50001 DROP TABLE IF EXISTS `vw_medlist` */;

/*!50001 CREATE TABLE  `vw_medlist`(
 `cliente_id` int(4) ,
 `medidor_id` varchar(20) ,
 `fecha_retiro` date ,
 `motivo_retiro` smallint(2) 
)*/;

/*Table structure for table `vw_medret` */

DROP TABLE IF EXISTS `vw_medret`;

/*!50001 DROP VIEW IF EXISTS `vw_medret` */;
/*!50001 DROP TABLE IF EXISTS `vw_medret` */;

/*!50001 CREATE TABLE  `vw_medret`(
 `medidor_id` varchar(20) ,
 `fecha` date ,
 `motivo_retiro` smallint(2) 
)*/;

/*Table structure for table `vw_movcli` */

DROP TABLE IF EXISTS `vw_movcli`;

/*!50001 DROP VIEW IF EXISTS `vw_movcli` */;
/*!50001 DROP TABLE IF EXISTS `vw_movcli` */;

/*!50001 CREATE TABLE  `vw_movcli`(
 `cliente_id` int(4) ,
 `nombre` varchar(102) ,
 `descripcion` varchar(30) ,
 `fecha` date ,
 `numero` varchar(18) ,
 `total` decimal(16,2) ,
 `fecha_pago` date ,
 `fecha_primero` date 
)*/;

/*Table structure for table `vw_plancuota` */

DROP TABLE IF EXISTS `vw_plancuota`;

/*!50001 DROP VIEW IF EXISTS `vw_plancuota` */;
/*!50001 DROP TABLE IF EXISTS `vw_plancuota` */;

/*!50001 CREATE TABLE  `vw_plancuota`(
 `cliente_id` int(4) ,
 `apellido` varchar(50) ,
 `nombre` varchar(50) ,
 `plan_id_cancela` smallint(2) ,
 `tipo_id` int(1) ,
 `tipo` varchar(5) ,
 `prefijo` smallint(2) ,
 `numero` smallint(2) ,
 `vencimiento` date ,
 `total` decimal(16,2) 
)*/;

/*Table structure for table `vw_plandetalle` */

DROP TABLE IF EXISTS `vw_plandetalle`;

/*!50001 DROP VIEW IF EXISTS `vw_plandetalle` */;
/*!50001 DROP TABLE IF EXISTS `vw_plandetalle` */;

/*!50001 CREATE TABLE  `vw_plandetalle`(
 `cliente_id` int(11) ,
 `apellido` varchar(50) ,
 `nombre` varchar(50) ,
 `plan_id_cancela` smallint(6) ,
 `tipo_id` int(11) ,
 `tipo` varchar(11) ,
 `prefijo` smallint(6) ,
 `numero` int(11) ,
 `vencimiento` date ,
 `total` decimal(16,2) 
)*/;

/*Table structure for table `vw_planfactura` */

DROP TABLE IF EXISTS `vw_planfactura`;

/*!50001 DROP VIEW IF EXISTS `vw_planfactura` */;
/*!50001 DROP TABLE IF EXISTS `vw_planfactura` */;

/*!50001 CREATE TABLE  `vw_planfactura`(
 `cliente_id` int(4) ,
 `apellido` varchar(50) ,
 `nombre` varchar(50) ,
 `plan_id_cancela` smallint(2) ,
 `tipo_id` int(1) ,
 `tipo` varchar(11) ,
 `prefijo` smallint(2) ,
 `numero` int(4) ,
 `vencimiento` date ,
 `total` decimal(16,2) 
)*/;

/*Table structure for table `vw_rescue` */

DROP TABLE IF EXISTS `vw_rescue`;

/*!50001 DROP VIEW IF EXISTS `vw_rescue` */;
/*!50001 DROP TABLE IF EXISTS `vw_rescue` */;

/*!50001 CREATE TABLE  `vw_rescue`(
 `cliente_id` int(4) ,
 `cantidad` bigint(21) 
)*/;

/*Table structure for table `vw_resfil` */

DROP TABLE IF EXISTS `vw_resfil`;

/*!50001 DROP VIEW IF EXISTS `vw_resfil` */;
/*!50001 DROP TABLE IF EXISTS `vw_resfil` */;

/*!50001 CREATE TABLE  `vw_resfil`(
 `cliente_id` int(11) 
)*/;

/*Table structure for table `vw_resmas` */

DROP TABLE IF EXISTS `vw_resmas`;

/*!50001 DROP VIEW IF EXISTS `vw_resmas` */;
/*!50001 DROP TABLE IF EXISTS `vw_resmas` */;

/*!50001 CREATE TABLE  `vw_resmas`(
 `prefijo_id` smallint(2) ,
 `factura_id` int(4) ,
 `cliente_id` int(4) ,
 `fecha_primero` date ,
 `numero_id` int(4) 
)*/;

/*Table structure for table `vw_resuna` */

DROP TABLE IF EXISTS `vw_resuna`;

/*!50001 DROP VIEW IF EXISTS `vw_resuna` */;
/*!50001 DROP TABLE IF EXISTS `vw_resuna` */;

/*!50001 CREATE TABLE  `vw_resuna`(
 `prefijo_id` smallint(2) ,
 `factura_id` int(4) ,
 `cliente_id` int(4) ,
 `fecha_primero` date ,
 `numero_id` int(4) 
)*/;

/*Table structure for table `vw_totpen` */

DROP TABLE IF EXISTS `vw_totpen`;

/*!50001 DROP VIEW IF EXISTS `vw_totpen` */;
/*!50001 DROP TABLE IF EXISTS `vw_totpen` */;

/*!50001 CREATE TABLE  `vw_totpen`(
 `cliente_id` int(4) ,
 `nombre` varchar(102) ,
 `descripcion` varchar(30) ,
 `fecha` date ,
 `numero` varchar(18) ,
 `total` decimal(16,2) 
)*/;

/*View structure for view vw_clicor */

/*!50001 DROP TABLE IF EXISTS `vw_clicor` */;
/*!50001 DROP VIEW IF EXISTS `vw_clicor` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_clicor` AS select `cliente`.`cliente_id` AS `cliente_id`,`cliente`.`apellido` AS `apellido`,`cliente`.`nombre` AS `nombre`,`cliente`.`inmueble_calle` AS `inmueble_calle`,`cliente`.`inmueble_puerta` AS `inmueble_puerta`,`cliente`.`inmueble_piso` AS `inmueble_piso`,`cliente`.`inmueble_dpto` AS `inmueble_dpto`,`cliente`.`cobro` AS `cobro` from (`cliente` join `vw_corcue` on((`cliente`.`cliente_id` = `vw_corcue`.`cliente_id`))) where ((`vw_corcue`.`cantidad` > 1) and isnull(`cliente`.`fecha_baja`)) group by `cliente`.`cliente_id`,`cliente`.`apellido`,`cliente`.`nombre`,`cliente`.`inmueble_calle`,`cliente`.`inmueble_puerta`,`cliente`.`inmueble_piso`,`cliente`.`inmueble_dpto`,`cliente`.`cobro` order by `cliente`.`apellido`,`cliente`.`nombre` */;

/*View structure for view vw_clihab */

/*!50001 DROP TABLE IF EXISTS `vw_clihab` */;
/*!50001 DROP VIEW IF EXISTS `vw_clihab` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_clihab` AS (select distinct `c`.`cliente_id` AS `cliente_id`,`m`.`medidor_id` AS `medidor_id`,`m`.`fecha_retiro` AS `fecha_retiro` from (`cliente` `c` join `medidor` `m` on((`c`.`cliente_id` = `m`.`cliente_id`))) where (isnull(`m`.`fecha_retiro`) and isnull(`c`.`fecha_baja`))) */;

/*View structure for view vw_climed */

/*!50001 DROP TABLE IF EXISTS `vw_climed` */;
/*!50001 DROP VIEW IF EXISTS `vw_climed` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_climed` AS (select `c`.`cliente_id` AS `cliente_id`,max(`m`.`fecha_alta`) AS `fecha`,`m`.`medidor_id` AS `medidor_id` from (`cliente` `c` join `medidor` `m` on((`c`.`cliente_id` = `m`.`cliente_id`))) where (isnull(`c`.`fecha_baja`) and isnull(`m`.`fecha_retiro`)) group by `c`.`cliente_id`,`m`.`medidor_id`,`c`.`zona`,`c`.`ruta`,`c`.`orden` order by `c`.`zona`,`c`.`ruta`,`c`.`orden`) */;

/*View structure for view vw_clires */

/*!50001 DROP TABLE IF EXISTS `vw_clires` */;
/*!50001 DROP VIEW IF EXISTS `vw_clires` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_clires` AS select `cliente`.`cliente_id` AS `cliente_id`,`cliente`.`apellido` AS `apellido`,`cliente`.`nombre` AS `nombre`,`cliente`.`inmueble_calle` AS `inmueble_calle`,`cliente`.`inmueble_puerta` AS `inmueble_puerta`,`cliente`.`inmueble_piso` AS `inmueble_piso`,`cliente`.`inmueble_dpto` AS `inmueble_dpto`,`cliente`.`cobro` AS `cobro` from (`cliente` join `vw_resfil` on((`cliente`.`cliente_id` = `vw_resfil`.`cliente_id`))) where isnull(`cliente`.`fecha_baja`) group by `cliente`.`cliente_id`,`cliente`.`apellido`,`cliente`.`nombre`,`cliente`.`inmueble_calle`,`cliente`.`inmueble_puerta`,`cliente`.`inmueble_piso`,`cliente`.`inmueble_dpto`,`cliente`.`cobro` order by `cliente`.`apellido`,`cliente`.`nombre` */;

/*View structure for view vw_corcue */

/*!50001 DROP TABLE IF EXISTS `vw_corcue` */;
/*!50001 DROP VIEW IF EXISTS `vw_corcue` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_corcue` AS (select `vw_corsel`.`cliente_id` AS `cliente_id`,count(0) AS `cantidad` from `vw_corsel` group by `vw_corsel`.`cliente_id`) */;

/*View structure for view vw_corsel */

/*!50001 DROP TABLE IF EXISTS `vw_corsel` */;
/*!50001 DROP VIEW IF EXISTS `vw_corsel` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_corsel` AS select `factura`.`prefijo_id` AS `prefijo_id`,`factura`.`factura_id` AS `factura_id`,`factura`.`cliente_id` AS `cliente_id`,`periodo`.`fecha_primero` AS `fecha_primero`,`imputado`.`numero_id` AS `numero_id` from ((`periodo` join `factura` on((`periodo`.`periodo_id` = `factura`.`periodo_id`))) left join `imputado` on(((`imputado`.`comp_serie_id` = `factura`.`prefijo_id`) and (`imputado`.`comp_numero_id` = `factura`.`factura_id`) and isnull(`imputado`.`numero_id`)))) where ((`factura`.`pagada` = 0) and (`factura`.`anulada` = 0) and (`factura`.`cancelada` = 0) and (`periodo`.`fecha_primero` < (curdate() + interval -(15) day))) order by `factura`.`cliente_id` */;

/*View structure for view vw_factpendientes */

/*!50001 DROP TABLE IF EXISTS `vw_factpendientes` */;
/*!50001 DROP VIEW IF EXISTS `vw_factpendientes` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_factpendientes` AS (select `f`.`cliente_id` AS `cliente_id`,`f`.`fecha` AS `fecha`,`f`.`periodo_id` AS `periodo_id`,`f`.`prefijo_id` AS `prefijo_id`,`f`.`factura_id` AS `factura_id`,`f`.`total` AS `total`,(`f`.`total` - if(isnull(sum(`n`.`total`)),0,sum(`n`.`total`))) AS `total_final` from (`factura` `f` left join `ncredito` `n` on(((`f`.`prefijo_id` = `n`.`prefijo_id`) and (`f`.`factura_id` = `n`.`factura_id`) and (`n`.`anulado` = 0)))) where ((`f`.`pagada` = 0) and (`f`.`cancelada` = 0) and (`f`.`anulada` = 0)) group by `f`.`prefijo_id`,`f`.`factura_id`) */;

/*View structure for view vw_liqperiodo */

/*!50001 DROP TABLE IF EXISTS `vw_liqperiodo` */;
/*!50001 DROP VIEW IF EXISTS `vw_liqperiodo` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_liqperiodo` AS (select `factura`.`periodo_id` AS `periodo_id`,sum(`factura`.`total`) AS `liquidado` from `factura` where ((`factura`.`anulada` = 0) and (`factura`.`cancelada` = 0)) group by `factura`.`periodo_id`) */;

/*View structure for view vw_medlist */

/*!50001 DROP TABLE IF EXISTS `vw_medlist` */;
/*!50001 DROP VIEW IF EXISTS `vw_medlist` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_medlist` AS (select distinct `c`.`cliente_id` AS `cliente_id`,`m`.`medidor_id` AS `medidor_id`,`m`.`fecha` AS `fecha_retiro`,`m`.`motivo_retiro` AS `motivo_retiro` from (`vw_medret` `m` left join `vw_clihab` `c` on((`c`.`medidor_id` = `m`.`medidor_id`))) where isnull(`c`.`cliente_id`)) */;

/*View structure for view vw_medret */

/*!50001 DROP TABLE IF EXISTS `vw_medret` */;
/*!50001 DROP VIEW IF EXISTS `vw_medret` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_medret` AS (select `m`.`medidor_id` AS `medidor_id`,max(`m`.`fecha_retiro`) AS `fecha`,`m`.`motivo_retiro` AS `motivo_retiro` from `medidor` `m` where (`m`.`motivo_retiro` > 0) group by `m`.`medidor_id`,`m`.`motivo_retiro`) */;

/*View structure for view vw_movcli */

/*!50001 DROP TABLE IF EXISTS `vw_movcli` */;
/*!50001 DROP VIEW IF EXISTS `vw_movcli` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_movcli` AS (select `c`.`cliente_id` AS `cliente_id`,concat(`c`.`apellido`,', ',`c`.`nombre`) AS `nombre`,`p`.`descripcion` AS `descripcion`,`f`.`fecha` AS `fecha`,concat(`f`.`prefijo_id`,'-',`f`.`factura_id`) AS `numero`,`f`.`total` AS `total`,`f`.`fecha_pago` AS `fecha_pago`,`p`.`fecha_primero` AS `fecha_primero` from ((`periodo` `p` join `factura` `f` on((`p`.`periodo_id` = `f`.`periodo_id`))) join `cliente` `c` on((`c`.`cliente_id` = `f`.`cliente_id`))) where ((`f`.`anulada` = 0) and isnull(`c`.`fecha_baja`)) order by `c`.`apellido`,`c`.`nombre`,`f`.`fecha`) */;

/*View structure for view vw_plancuota */

/*!50001 DROP TABLE IF EXISTS `vw_plancuota` */;
/*!50001 DROP VIEW IF EXISTS `vw_plancuota` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_plancuota` AS (select `t`.`cliente_id` AS `cliente_id`,`c`.`apellido` AS `apellido`,`c`.`nombre` AS `nombre`,`t`.`plan_id_cancela` AS `plan_id_cancela`,2 AS `tipo_id`,'Cuota' AS `tipo`,`t`.`plan_id` AS `prefijo`,`t`.`cuota_id` AS `numero`,`t`.`fecha_vencimiento` AS `vencimiento`,`t`.`importe` AS `total` from (`cuota` `t` join `cliente` `c` on(((`c`.`cliente_id` = `t`.`cliente_id`) and isnull(`c`.`fecha_baja`)))) where (`t`.`cancelada` = 1)) */;

/*View structure for view vw_plandetalle */

/*!50001 DROP TABLE IF EXISTS `vw_plandetalle` */;
/*!50001 DROP VIEW IF EXISTS `vw_plandetalle` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_plandetalle` AS select `vw_planfactura`.`cliente_id` AS `cliente_id`,`vw_planfactura`.`apellido` AS `apellido`,`vw_planfactura`.`nombre` AS `nombre`,`vw_planfactura`.`plan_id_cancela` AS `plan_id_cancela`,`vw_planfactura`.`tipo_id` AS `tipo_id`,`vw_planfactura`.`tipo` AS `tipo`,`vw_planfactura`.`prefijo` AS `prefijo`,`vw_planfactura`.`numero` AS `numero`,`vw_planfactura`.`vencimiento` AS `vencimiento`,`vw_planfactura`.`total` AS `total` from `vw_planfactura` union all select `vw_plancuota`.`cliente_id` AS `cliente_id`,`vw_plancuota`.`apellido` AS `apellido`,`vw_plancuota`.`nombre` AS `nombre`,`vw_plancuota`.`plan_id_cancela` AS `plan_id_cancela`,`vw_plancuota`.`tipo_id` AS `tipo_id`,`vw_plancuota`.`tipo` AS `tipo`,`vw_plancuota`.`prefijo` AS `prefijo`,`vw_plancuota`.`numero` AS `numero`,`vw_plancuota`.`vencimiento` AS `vencimiento`,`vw_plancuota`.`total` AS `total` from `vw_plancuota` */;

/*View structure for view vw_planfactura */

/*!50001 DROP TABLE IF EXISTS `vw_planfactura` */;
/*!50001 DROP VIEW IF EXISTS `vw_planfactura` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_planfactura` AS (select `f`.`cliente_id` AS `cliente_id`,`c`.`apellido` AS `apellido`,`c`.`nombre` AS `nombre`,`f`.`plan_id_cancela` AS `plan_id_cancela`,1 AS `tipo_id`,'Liquidación' AS `tipo`,`f`.`prefijo_id` AS `prefijo`,`f`.`factura_id` AS `numero`,`p`.`fecha_primero` AS `vencimiento`,`f`.`total` AS `total` from ((`factura` `f` join `cliente` `c` on(((`c`.`cliente_id` = `f`.`cliente_id`) and isnull(`c`.`fecha_baja`)))) join `periodo` `p` on((`p`.`periodo_id` = `f`.`periodo_id`))) where (`f`.`cancelada` = 1)) */;

/*View structure for view vw_rescue */

/*!50001 DROP TABLE IF EXISTS `vw_rescue` */;
/*!50001 DROP VIEW IF EXISTS `vw_rescue` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_rescue` AS (select `vw_resmas`.`cliente_id` AS `cliente_id`,count(0) AS `cantidad` from `vw_resmas` group by `vw_resmas`.`cliente_id`) */;

/*View structure for view vw_resfil */

/*!50001 DROP TABLE IF EXISTS `vw_resfil` */;
/*!50001 DROP VIEW IF EXISTS `vw_resfil` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_resfil` AS select `vw_rescue`.`cliente_id` AS `cliente_id` from `vw_rescue` where (`vw_rescue`.`cantidad` > 1) union select `vw_resuna`.`cliente_id` AS `cliente_id` from `vw_resuna` */;

/*View structure for view vw_resmas */

/*!50001 DROP TABLE IF EXISTS `vw_resmas` */;
/*!50001 DROP VIEW IF EXISTS `vw_resmas` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_resmas` AS select `factura`.`prefijo_id` AS `prefijo_id`,`factura`.`factura_id` AS `factura_id`,`factura`.`cliente_id` AS `cliente_id`,`periodo`.`fecha_primero` AS `fecha_primero`,`imputado`.`numero_id` AS `numero_id` from ((`periodo` join `factura` on((`periodo`.`periodo_id` = `factura`.`periodo_id`))) left join `imputado` on(((`imputado`.`comp_numero_id` = `factura`.`factura_id`) and (`imputado`.`comp_serie_id` = `factura`.`prefijo_id`) and isnull(`imputado`.`numero_id`)))) where ((`factura`.`pagada` = 0) and (`factura`.`anulada` = 0) and (`factura`.`cancelada` = 0) and (`periodo`.`fecha_primero` < curdate())) */;

/*View structure for view vw_resuna */

/*!50001 DROP TABLE IF EXISTS `vw_resuna` */;
/*!50001 DROP VIEW IF EXISTS `vw_resuna` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_resuna` AS select `factura`.`prefijo_id` AS `prefijo_id`,`factura`.`factura_id` AS `factura_id`,`factura`.`cliente_id` AS `cliente_id`,`periodo`.`fecha_primero` AS `fecha_primero`,`imputado`.`numero_id` AS `numero_id` from ((`periodo` join `factura` on((`periodo`.`periodo_id` = `factura`.`periodo_id`))) left join `imputado` on(((`imputado`.`comp_numero_id` = `factura`.`factura_id`) and (`imputado`.`comp_serie_id` = `factura`.`prefijo_id`) and isnull(`imputado`.`numero_id`)))) where ((`factura`.`pagada` = 0) and (`factura`.`anulada` = 0) and (`factura`.`cancelada` = 0) and (`periodo`.`fecha_primero` < (curdate() + interval -(61) day))) */;

/*View structure for view vw_totpen */

/*!50001 DROP TABLE IF EXISTS `vw_totpen` */;
/*!50001 DROP VIEW IF EXISTS `vw_totpen` */;

/*!50001 CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` SQL SECURITY DEFINER VIEW `vw_totpen` AS (select `c`.`cliente_id` AS `cliente_id`,concat(`c`.`apellido`,', ',`c`.`nombre`) AS `nombre`,`p`.`descripcion` AS `descripcion`,`f`.`fecha` AS `fecha`,concat(`f`.`prefijo_id`,'-',`f`.`factura_id`) AS `numero`,`f`.`total` AS `total` from ((`periodo` `p` join `factura` `f` on((`p`.`periodo_id` = `f`.`periodo_id`))) join `cliente` `c` on((`c`.`cliente_id` = `f`.`cliente_id`))) where ((`f`.`pagada` = 0) and (`f`.`anulada` = 0) and (`f`.`cancelada` = 0) and isnull(`c`.`fecha_baja`)) order by `c`.`apellido`,`c`.`nombre`,`f`.`fecha`) */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;
