-- create database ImportExcel;
-- use ImportExcel;

create table SOGIP_Usuario(
	id int not null identity,
	cedula varchar(45) not null unique,
	contrasena varchar(45) not null,
	fecha_expiracion datetime not null,
	constraint pkSOGIP_Usuario primary key(id)
);

select * from SOGIP_Usuario;