LOAD DATA LOW_PRIORITY LOCAL INFILE 'C:\\Users\\uzian\\Desktop\\adjustments\\local_db_daily_adjustments.csv' REPLACE INTO TABLE `invrec`.`quantity_adjustments` CHARACTER SET latin1 FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '"' ESCAPED BY '"' LINES TERMINATED BY '\r\n' IGNORE 1 LINES (`SKU`, `Description`, `NS_Adj_Req`);

SELECT * FROM invrec.quantity_adjustments WHERE DATE(Date) = CURDATE();