create table goods
(
g_code varchar(20) not  null,
g_part varchar(10) not null,
g_name varchar(20) not null,
g_maker varchar(20) not null,
g_origin_price int not null,
g_sellprice int not null,
g_updateday varchar(30) not null,
g_image varchar(20) null,
g_content text not null
)


create table customer
(
c_code varchar(30) not null,
c_name varchar(10) not null,
c_address varchar(50) not null,
c_tel varchar(20) not null
)

create table imsi_buy
(
imsi_memid varchar(20) not null,
imsi_goodscode varchar(20) not null,
imsi_ea smallint
)

create table buy
(
b_code varchar(30) not null,
c_code varchar(30) not null,
b_date varchar(20) not null,
b_summary text null,
b_totalprice int,
b_bank varchar(10) not null
)