
CREATE TABLE address_family (
  add_fam_id integer primary key,
  family text
);
insert into address_family values(1, 'IPv4');
insert into address_family values(2, 'IPv6');


CREATE TABLE vrf (
  vrf_id integer primary key,
  name text,
  rd text,
  description text,
  app_id integer
);

CREATE TABLE route_map (
  map_id integer primary key,
  vrf_id integer,
  name text,
  type text,
  map_conf blob
);

CREATE TABLE static_route (
  route_id integer primary key,
  next_hop_ip integer,
  next_hop_int integer,
  vrf_id integer,
  subnet_id integer,
  ad_distance integer,
  name text,
  add_fam_id integer,
  app_id integer
);

CREATE TABLE ip_address (
  ip_id integer primary key,
  add_fam_id integer,
  address text,
  subnet_id integer,
  address_type text,
  int_id integer
);

CREATE TABLE log_file (
  file_id integer primary key,
  app_id integer,
  filename text,
  file_text blob,
  importance integer
);

CREATE TABLE site (
  site_id integer primary key,
  name text
);

CREATE TABLE role (
  role_id integer primary key,
  role_name text,
  app_id integer
);

CREATE TABLE vlan (
  vlan_id integer primary key,
  vlan_no integer,
  name text,
  app_id integer,
  exist integer
);

CREATE TABLE int_vlan (
  int_id integer ,
  vlan_id integer,
  vlan_mode text,
  int_vlan_id integer primary key
);

CREATE TABLE import_rt (
  vrf_id integer,
  add_fam_id integer,
  rt_import text
);

CREATE TABLE export_rt (
  vrf_id integer,
  add_fam_id integer,
  rt_export text
);

CREATE TABLE appliance (
  app_id integer primary key,
  role_id integer,
  site_id integer,
  hostname text,
  stack integer
);

CREATE TABLE interface (
  int_id integer primary key,
  type text,
  number text,
  description text,
  mode text,
  member_of integer,
  app_id integer,
  vrf_id integer,
  tunnel_id integer,
  status text
);

CREATE TABLE tunnel_int (
  tunnel_id integer primary key,
  source_int integer,
  source_ip integer,
  dest_ip integer
);

CREATE TABLE subnet (
  subnet_id integer primary key,
  network_id text
);

CREATE TABLE pending_interface (
  int_id integer primary key,
  type text,
  number text,
  description text,
  mode text,
  member_of integer,
  app_id integer,
  vrf_name text,
  tunnel_id integer,
  status text,
  pend_reason text
);

CREATE TABLE pending_tunnel_int (
  tunnel_id integer primary key,
  source_int text,
  source_ip text,
  dest_ip text
);

CREATE TABLE pend_ip_address (
  ip_id integer primary key,
  add_fam_id integer,
  address text,
  subnet_id integer,
  address_type text,
  int_id integer
);

CREATE TABLE pend_int_vlan (
  int_id integer ,
  vlan_id integer,
  vlan_mode text,
  int_vlan_id integer primary key
);

CREATE TABLE pend_static_route (
  route_id integer primary key,
  next_hop_ip text,
  next_hop_int_type text,
  next_hop_int_number text,
  vrf_name text,
  subnet_id integer,
  ad_distance integer,
  name text,
  add_fam_id integer,
  app_id integer
);