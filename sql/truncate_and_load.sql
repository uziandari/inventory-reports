truncate inventory.bucket;
truncate inventory.big_commerce_daily;
truncate inventory.ca_inventory;
truncate inventory.relists;
truncate inventory.main_inventory;
truncate inventory.ns_inventory;
truncate inventory.receipt_date;
truncate inventory.removes_daily;
truncate inventory.wholesale_pending;


LOAD DATA LOW_PRIORITY LOCAL INFILE 'C:\\Users\\uzian\\Desktop\\database_files\\web\\import_files\\bucket.csv' REPLACE INTO TABLE `inventory`.`bucket` CHARACTER SET latin1 FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"' ESCAPED BY '"' LINES TERMINATED BY '\r\n' IGNORE 1 LINES (`sku`);

LOAD DATA LOW_PRIORITY LOCAL INFILE 'C:\\Users\\uzian\\Desktop\\database_files\\web\\import_files\\ca_import.csv' REPLACE INTO TABLE `inventory`.`ca_inventory` CHARACTER SET latin1 FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"' ESCAPED BY '"' LINES TERMINATED BY '\r\n' IGNORE 1 LINES (`sku`, `description`, `total`, `available`, `pending_checkout`, `pending_payment`, `pending_shipment`, `flag`, `blocked`, `parent_sku`, `label`, `img_url`);

LOAD DATA LOW_PRIORITY LOCAL INFILE 'C:\\Users\\uzian\\Desktop\\database_files\\web\\import_files\\relist.csv' REPLACE INTO TABLE `inventory`.`relists` CHARACTER SET latin1 FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"' ESCAPED BY '"' LINES TERMINATED BY '\r\n' IGNORE 1 LINES (`sku`);

LOAD DATA LOW_PRIORITY LOCAL INFILE 'C:\\Users\\uzian\\Desktop\\database_files\\web\\import_files\\ns_import.csv' REPLACE INTO TABLE `inventory`.`ns_inventory` CHARACTER SET latin1 FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"' ESCAPED BY '"' LINES TERMINATED BY '\r\n' IGNORE 1 LINES (`sku`, `description`, `location`, `backstock`, `upc`, `stock`, `committed`, `head_cover`, `inline`, `purchase_price`);

LOAD DATA LOW_PRIORITY LOCAL INFILE 'C:\\Users\\uzian\\Desktop\\database_files\\web\\import_files\\receipt_date_import.csv' REPLACE INTO TABLE `inventory`.`receipt_date` CHARACTER SET latin1 FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"' ESCAPED BY '"' LINES TERMINATED BY '\r\n' IGNORE 1 LINES (`sku`, `receipt`);

LOAD DATA LOW_PRIORITY LOCAL INFILE 'C:\\Users\\uzian\\Desktop\\database_files\\web\\import_files\\prevdelists.csv' REPLACE INTO TABLE `inventory`.`removes_daily` CHARACTER SET latin1 FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"' ESCAPED BY '"' LINES TERMINATED BY '\r\n' IGNORE 1 LINES (`sku`);

LOAD DATA LOW_PRIORITY LOCAL INFILE 'C:\\Users\\uzian\\Desktop\\database_files\\web\\import_files\\wholesale_committed.csv' REPLACE INTO TABLE `inventory`.`wholesale_pending` CHARACTER SET latin1 FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"' ESCAPED BY '"' LINES TERMINATED BY '\r\n' IGNORE 1 LINES (`sku`);

LOAD DATA LOW_PRIORITY LOCAL INFILE 'C:\\Users\\uzian\\Desktop\\database_files\\web\\import_files\\daily_bc.csv' REPLACE INTO TABLE `inventory`.`big_commerce_daily` CHARACTER SET latin1 FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"' ESCAPED BY '"' LINES TERMINATED BY '\r\n' IGNORE 1 LINES (`sku`);