-- public.si_actsubact definition

-- Drop table

-- DROP TABLE public.si_actsubact;

CREATE TABLE public.si_actsubact (
	id_act int2 NULL,
	id_subact int2 NULL,
	descrip bpchar(120) NULL,
	altoriesgo int2 NULL,
	idcnbv bpchar(7) NULL,
	status bpchar(1) DEFAULT 'A'::bpchar NULL,
	factivacion date NULL,
	fcancelacion date NULL,
	usuario bpchar(8) NULL
);
INSERT INTO public.si_actsubact (id_act,id_subact,descrip,altoriesgo,idcnbv,status,factivacion,fcancelacion,usuario) VALUES
	 (4,2,'Armada                                                                                                                  ',0,'9400003','A','2010-04-21',NULL,'92862721'),
	 (4,4,'Marina/Militar                                                                                                          ',0,'9400003','A','2010-04-21',NULL,'92862721'),
	 (4,5,'Servicios de Seguridad Privada                                                                                          ',0,'8429038','A','2010-04-21',NULL,'92862721'),
	 (5,0,'Empleado                                                                                                                ',0,'0000000','A','2010-04-21',NULL,'92862721'),
	 (5,1,'Abarrotes                                                                                                               ',0,'6131023','A','2010-04-21',NULL,'92862721'),
	 (5,2,'Agricultura                                                                                                             ',0,'0100008','A','2010-04-21',NULL,'92862721'),
	 (5,3,'Comercio                                                                                                                ',0,'6999900','A','2010-04-21',NULL,'92862721'),
	 (5,4,'Comunicaciones                                                                                                          ',0,'7600001','A','2010-04-21',NULL,'92862721'),
	 (5,12,'Gasolineras                                                                                                             ',0,'3111010','A','2010-04-21',NULL,'92862721'),
	 (5,13,'Industria                                                                                                               ',0,'3999903','A','2010-04-21',NULL,'92862721');
INSERT INTO public.si_actsubact (id_act,id_subact,descrip,altoriesgo,idcnbv,status,factivacion,fcancelacion,usuario) VALUES
	 (5,16,'Pesca                                                                                                                   ',0,'0400002','A','2010-04-21',NULL,'92862721'),
	 (5,17,'Servicios Administrativos, Financieros y Contables                                                                      ',0,'8413015','A','2010-04-21',NULL,'92862721'),
	 (5,19,'Servicios de Salud y Hospitalarios                                                                                      ',0,'9211012','A','2010-04-21',NULL,'92862721'),
	 (5,21,'Tiendas de Autoservicio                                                                                                 ',0,'6411011','A','2010-04-21',NULL,'92862721'),
	 (5,22,'Transportes                                                                                                             ',0,'7100001','A','2010-04-21',NULL,'92862721'),
	 (5,23,'Otros Servicios                                                                                                         ',0,'8429012','A','2010-04-21',NULL,'92862721'),
	 (6,99,'Desempleado                                                                                                             ',0,'9900905','A','2010-04-21',NULL,'92862721'),
	 (7,99,'Ama de Casa                                                                                                             ',0,'9900908','A','2010-04-21',NULL,'92862721'),
	 (8,99,'Estudiante                                                                                                              ',0,'9900904','A','2010-04-21',NULL,'92862721'),
	 (9,99,'Pensionado o Jubilado                                                                                                   ',0,'9900906','A','2010-04-21',NULL,'92862721');
INSERT INTO public.si_actsubact (id_act,id_subact,descrip,altoriesgo,idcnbv,status,factivacion,fcancelacion,usuario) VALUES
	 (10,99,'Migrante                                                                                                                ',0,'8429012','A','2010-04-21',NULL,'92862721'),
	 (1,8,'Otros Servicios                                                                                                         ',0,'8429012','A','2010-04-21',NULL,'92862721'),
	 (3,52,'Centros Nocturnos                                                                                                       ',1,'8831019','A','2010-04-21',NULL,'92862721'),
	 (3,53,'Venta de Autos Usados                                                                                                   ',0,'6812011','A','2010-04-21',NULL,'92862721'),
	 (4,6,'Notario Publico                                                                                                         ',0,'8411019','A','2010-04-21',NULL,'92862721'),
	 (3,17,'Escuela Privada                                                                                                         ',0,'9119018','A','2010-04-21',NULL,'92862721'),
	 (3,21,'Gasolineras                                                                                                             ',1,'3111010','A','2010-04-21',NULL,'92862721'),
	 (3,22,'Hotel                                                                                                                   ',0,'8611015','A','2010-04-21',NULL,'92862721'),
	 (3,23,'Importadora y/o Exportadora                                                                                             ',1,'7513014','A','2010-04-21',NULL,'92862721'),
	 (3,24,'Imprenta                                                                                                                ',0,'2921056','A','2010-04-21',NULL,'92862721');
INSERT INTO public.si_actsubact (id_act,id_subact,descrip,altoriesgo,idcnbv,status,factivacion,fcancelacion,usuario) VALUES
	 (3,33,'Refaccionaria                                                                                                           ',0,'6700000','A','2010-04-21',NULL,'92862721'),
	 (3,34,'Restaurante                                                                                                             ',0,'8711021','A','2010-04-21',NULL,'92862721'),
	 (3,36,'Servicios Administrativos, Financieros y Contables                                                                      ',0,'8413015','A','2010-04-21',NULL,'92862721'),
	 (3,38,'Servicios de Salud y Hospitalarios                                                                                      ',0,'9211012','A','2010-04-21',NULL,'92862721'),
	 (3,43,'Tiendas de Autoservicio                                                                                                 ',0,'6411011','A','2010-04-21',NULL,'92862721'),
	 (3,45,'Transportes                                                                                                             ',0,'7100001','A','2010-04-21',NULL,'92862721'),
	 (3,46,'Venta de Armas                                                                                                          ',1,'6991013','A','2010-04-21',NULL,'92862721'),
	 (3,47,'Venta de Autos Nuevos                                                                                                   ',0,'6811013','A','2010-04-21',NULL,'92862721'),
	 (3,48,'Venta de Inmuebles                                                                                                      ',0,'6900006','A','2010-04-21',NULL,'92862721'),
	 (3,49,'Venta de Joyas y Obras de Arte                                                                                          ',1,'6225024','A','2010-04-21',NULL,'92862721');
INSERT INTO public.si_actsubact (id_act,id_subact,descrip,altoriesgo,idcnbv,status,factivacion,fcancelacion,usuario) VALUES
	 (3,50,'Veterinaria                                                                                                             ',0,'9200007','A','2010-04-21',NULL,'92862721'),
	 (3,51,'Otros Servicios                                                                                                         ',0,'8429012','A','2010-04-21',NULL,'92862721'),
	 (4,1,'Abogado Litigante                                                                                                       ',0,'8412017','A','2010-04-21',NULL,'92862721'),
	 (1,0,'Profesionista Independiente                                                                                             ',0,'0000000','A','2010-04-21',NULL,'92862721'),
	 (2,1,'Mecanico                                                                                                                ',0,'8913015','A','2010-04-21',NULL,'92862721'),
	 (3,4,'Cajas de Ahorro, Piramides y Empe?o                                                                                     ',1,'9503005','A','2010-04-21',NULL,'92862721'),
	 (3,5,'Carniceria, Cremeria y Fruteria                                                                                         ',0,'6131023','A','2010-04-21',NULL,'92862721'),
	 (3,6,'Carpinteria                                                                                                             ',0,'4200002','A','2010-04-21',NULL,'92862721'),
	 (3,10,'Cerveceria                                                                                                              ',0,'8722010','A','2010-04-21',NULL,'92862721'),
	 (3,11,'Cocina Economica o Fonda                                                                                                ',0,'8700008','A','2010-04-21',NULL,'92862721');
INSERT INTO public.si_actsubact (id_act,id_subact,descrip,altoriesgo,idcnbv,status,factivacion,fcancelacion,usuario) VALUES
	 (3,13,'Construccion                                                                                                            ',0,'3800001','A','2010-04-21',NULL,'92862721'),
	 (3,14,'Consultorio Medico                                                                                                      ',0,'9200007','A','2010-04-21',NULL,'92862721'),
	 (3,15,'Dulceria                                                                                                                ',0,'6132013','A','2010-04-21',NULL,'92862721'),
	 (1,2,'Agricultor                                                                                                              ',0,'0100008','A','2010-04-21',NULL,'92862721'),
	 (1,3,'Carpintero                                                                                                              ',0,'4200002','A','2010-04-21',NULL,'92862721'),
	 (1,4,'Cirquero                                                                                                                ',0,'8800006','A','2010-04-21',NULL,'92862721'),
	 (1,5,'Comerciante Ambulante                                                                                                   ',0,'6999900','A','2010-04-21',NULL,'92862721'),
	 (1,6,'Electricista                                                                                                            ',0,'5011903','A','2010-04-21',NULL,'92862721'),
	 (1,7,'Ganadero                                                                                                                ',0,'0200006','A','2010-04-21',NULL,'92862721'),
	 (2,0,'Oficio                                                                                                                  ',0,'0000000','A','2010-04-21',NULL,'92862721');
INSERT INTO public.si_actsubact (id_act,id_subact,descrip,altoriesgo,idcnbv,status,factivacion,fcancelacion,usuario) VALUES
	 (2,2,'Mesero                                                                                                                  ',0,'8711021','A','2010-04-21',NULL,'92862721'),
	 (2,3,'Obrero                                                                                                                  ',0,'4200002','A','2010-04-21',NULL,'92862721'),
	 (2,4,'Pescador                                                                                                                ',0,'8819015','A','2010-04-21',NULL,'92862721'),
	 (2,5,'Pintor                                                                                                                  ',0,'4200002','A','2010-04-21',NULL,'92862721'),
	 (2,6,'Plomero                                                                                                                 ',0,'4200002','A','2010-04-21',NULL,'92862721'),
	 (2,7,'Taxista                                                                                                                 ',0,'7113012','A','2010-04-21',NULL,'92862721'),
	 (2,8,'Otros Oficios                                                                                                           ',0,'4200002','A','2010-04-21',NULL,'92862721'),
	 (3,0,'Negocio Propio                                                                                                          ',0,'0000000','A','2010-04-21',NULL,'92862721'),
	 (3,1,'Abarrotes                                                                                                               ',0,'6131023','A','2010-04-21',NULL,'92862721'),
	 (3,2,'Agencia Aduanal                                                                                                         ',1,'7513014','A','2010-04-21',NULL,'92862721');
INSERT INTO public.si_actsubact (id_act,id_subact,descrip,altoriesgo,idcnbv,status,factivacion,fcancelacion,usuario) VALUES
	 (3,3,'Agencia de Viaje                                                                                                        ',1,'7512016','A','2010-04-21',NULL,'92862721'),
	 (3,7,'Casa de Cambio                                                                                                          ',1,'8219025','A','2010-04-21',NULL,'92862721'),
	 (3,8,'Casinos                                                                                                                 ',1,'8839013','A','2010-04-21',NULL,'92862721'),
	 (3,9,'Bares y Cantinas                                                                                                        ',1,'8721012','A','2010-04-21',NULL,'92862721'),
	 (3,12,'Comercio                                                                                                                ',0,'6999900','A','2010-04-21',NULL,'92862721'),
	 (3,16,'Educacion/Cultura                                                                                                       ',0,'9115016','A','2010-04-21',NULL,'92862721'),
	 (3,18,'Estacionamientos Publicos                                                                                               ',1,'8916027','A','2010-04-21',NULL,'92862721'),
	 (3,19,'Estetica                                                                                                                ',0,'8932015','A','2010-04-21',NULL,'92862721'),
	 (3,20,'Ferreteria                                                                                                              ',0,'6600002','A','2010-04-21',NULL,'92862721'),
	 (3,25,'Libreria                                                                                                                ',0,'6227012','A','2010-04-21',NULL,'92862721');
INSERT INTO public.si_actsubact (id_act,id_subact,descrip,altoriesgo,idcnbv,status,factivacion,fcancelacion,usuario) VALUES
	 (3,26,'Licoreria                                                                                                               ',0,'6136023','A','2010-04-21',NULL,'92862721'),
	 (3,27,'Mantenimiento y Reparacion                                                                                              ',0,'8427016','A','2010-04-21',NULL,'92862721'),
	 (3,28,'Neveria                                                                                                                 ',0,'8714017','A','2010-04-21',NULL,'92862721'),
	 (3,29,'Organizaciones de Caridad, Sindicales, Politicas y Religiosas                                                           ',1,'9300005','A','2010-04-21',NULL,'92862721'),
	 (3,30,'Panaderia                                                                                                               ',0,'2071017','A','2010-04-21',NULL,'92862721'),
	 (3,31,'Papeleria                                                                                                               ',0,'6233019','A','2010-04-21',NULL,'92862721'),
	 (3,35,'Sastreria                                                                                                               ',0,'2412039','A','2010-04-21',NULL,'92862721'),
	 (3,37,'Servicios de Hoteleria y Recreativos                                                                                    ',0,'8611015','A','2010-04-21',NULL,'92862721'),
	 (3,39,'Servicios de Tecnologia e Informatica                                                                                   ',0,'8421018','A','2010-04-21',NULL,'92862721'),
	 (3,40,'Tabaqueria                                                                                                              ',0,'6100002','A','2010-04-21',NULL,'92862721');
INSERT INTO public.si_actsubact (id_act,id_subact,descrip,altoriesgo,idcnbv,status,factivacion,fcancelacion,usuario) VALUES
	 (3,41,'Taller Electrico                                                                                                        ',0,'4222014','A','2010-04-21',NULL,'92862721'),
	 (3,42,'Taller Mecanico                                                                                                         ',0,'3699024','A','2010-04-21',NULL,'92862721'),
	 (3,44,'Tortilleria                                                                                                             ',0,'2093011','A','2010-04-21',NULL,'92862721'),
	 (4,3,'Servicios de Seguridad Publica (Federal o Estatal)                                                                      ',0,'9400003','A','2010-04-21',NULL,'92862721'),
	 (4,7,'Servicios de Seguridad Publica (Municipal)                                                                              ',0,NULL,'A',NULL,NULL,NULL),
	 (5,5,'Construccion                                                                                                            ',0,'4111027','A','2010-04-21',NULL,'92862721'),
	 (5,6,'Consultorio Medico                                                                                                      ',0,'9200007','A','2010-04-21',NULL,'92862721'),
	 (5,7,'Educacion/Cultura                                                                                                       ',0,'9115016','A','2010-04-21',NULL,'92862721'),
	 (5,8,'Empleado Publico (Gobierno Federal, Estatal, Municipal)                                                                 ',0,'9411998','A','2010-04-21',NULL,'92862721'),
	 (5,9,'Estetica                                                                                                                ',0,'8932015','A','2010-04-21',NULL,'92862721');
INSERT INTO public.si_actsubact (id_act,id_subact,descrip,altoriesgo,idcnbv,status,factivacion,fcancelacion,usuario) VALUES
	 (5,10,'Ferreteria                                                                                                              ',0,'6600002','A','2010-04-21',NULL,'92862721'),
	 (5,11,'Ganaderia                                                                                                               ',0,'0200006','A','2010-04-21',NULL,'92862721'),
	 (5,14,'Mantenimiento y Reparacion                                                                                              ',0,'8427016','A','2010-04-21',NULL,'92862721'),
	 (5,15,'Mineria                                                                                                                 ',0,'1100007','A','2010-04-21',NULL,'92862721'),
	 (5,18,'Servicios de Hoteleria y Recreativos                                                                                    ',0,'8611015','A','2010-04-21',NULL,'92862721'),
	 (5,20,'Servicios de Tecnologia e Informatica                                                                                   ',0,'8421018','A','2010-04-21',NULL,'92862721'),
	 (1,1,'Albanil                                                                                                                 ',0,'4200002','A','2010-04-21',NULL,'92862721'),
	 (4,0,'Abogado o Policia Judicial o Ministerial/ Seguridad                                                                     ',0,'0000000','A','2010-04-21',NULL,'92862721');


-- DROP FUNCTION public.buscar_actsubact(varchar);

CREATE OR REPLACE FUNCTION public.buscar_actsubact(p_descrip character varying)
 RETURNS TABLE(id_act integer, descrip character varying, id_subact integer)
 LANGUAGE plpgsql
AS $function$
BEGIN
  IF p_descrip IS NULL OR p_descrip = '' THEN
    RETURN QUERY
    SELECT sa.id_act, sa.descrip, sa.id_subact
    FROM si_actsubact sa;
  ELSE
    RETURN QUERY
    SELECT sa.id_act, sa.descrip, sa.id_subact
    FROM si_actsubact sa
    WHERE sa.descrip LIKE '%' || p_descrip || '%';
  END IF;
END;
$function$
;
