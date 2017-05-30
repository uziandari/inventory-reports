SET SQL_SAFE_UPDATES=0;
UPDATE inventory.receipt_to_stock
LEFT JOIN ns_inventory
ON receipt_to_stock.sku = ns_inventory.sku
INNER JOIN invrec.location_table
ON ns_inventory.location = invrec.location_table.location

SET receipt_to_stock.bin_size = 
  CASE WHEN (receipt_to_stock.sku LIKE "01%" OR receipt_to_stock.sku LIKE "15%" OR receipt_to_stock.sku LIKE "19%") AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received > 10 THEN "BALL"
        WHEN (receipt_to_stock.sku LIKE "01%" OR receipt_to_stock.sku LIKE "15%" OR receipt_to_stock.sku LIKE "19%") AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received <= 10 THEN "LQBALL"
        WHEN (receipt_to_stock.sku LIKE "02%" OR receipt_to_stock.sku LIKE "03%" OR receipt_to_stock.sku LIKE "04%" OR receipt_to_stock.sku LIKE "07%" OR receipt_to_stock.sku LIKE "08%" OR (receipt_to_stock.sku LIKE "09%" OR receipt_to_stock.sku LIKE "10%" AND (receipt_to_stock.description LIKE "%combo irons%" OR receipt_to_stock.description LIKE "%complete set%"))) AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received > 5 THEN "BULK"
        WHEN (receipt_to_stock.sku LIKE "02%" OR receipt_to_stock.sku LIKE "03%" OR receipt_to_stock.sku LIKE "04%" OR receipt_to_stock.sku LIKE "07%" OR receipt_to_stock.sku LIKE "08%" OR (receipt_to_stock.sku LIKE "09%" OR receipt_to_stock.sku LIKE "10%" AND (receipt_to_stock.description LIKE "%combo irons%" OR receipt_to_stock.description LIKE "%complete set%"))) AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received <= 5 THEN "BULKLQ"
        WHEN receipt_to_stock.sku LIKE "06%" AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received <= 14 THEN "XS"
        WHEN receipt_to_stock.sku LIKE "06%" AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received <= 40 THEN "SM"
        WHEN receipt_to_stock.sku LIKE "06%" AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received <= 90 THEN "MD"
        WHEN receipt_to_stock.sku LIKE "06%" AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received > 90 THEN "LG"
        WHEN (receipt_to_stock.sku LIKE "09%" OR receipt_to_stock.sku LIKE "10%" OR receipt_to_stock.sku LIKE "11%" OR receipt_to_stock.sku LIKE "12%" OR receipt_to_stock.sku LIKE "14%" OR receipt_to_stock.sku LIKE "21%") AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") OR (receipt_to_stock.sku LIKE "19%" AND (receipt_to_stock.description LIKE "%retriever%")) AND qty_received > 4 THEN "CL"
        WHEN (receipt_to_stock.sku LIKE "09%" OR receipt_to_stock.sku LIKE "10%" OR receipt_to_stock.sku LIKE "11%" OR receipt_to_stock.sku LIKE "12%" OR receipt_to_stock.sku LIKE "14%" OR receipt_to_stock.sku LIKE "21%") AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") OR (receipt_to_stock.sku LIKE "19%" AND (receipt_to_stock.description LIKE "%retriever%")) AND qty_received <= 4 THEN "CLLQ"
        WHEN receipt_to_stock.sku LIKE "16%" AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received <= 6 THEN "SUNLQ"
        WHEN receipt_to_stock.sku LIKE "16%" AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received > 6 THEN "SUN"
        #WHEN receipt_to_stock.sku LIKE "17%" AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received <= 40 THEN "SM"
        WHEN receipt_to_stock.sku LIKE "18%" AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received <= 9 THEN "SHOELQ"
        WHEN receipt_to_stock.sku LIKE "18%" AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received > 9 THEN "SHOE"
        WHEN (receipt_to_stock.sku LIKE "20%" OR receipt_to_stock.sku LIKE "32%" OR receipt_to_stock.sku LIKE "33%" OR receipt_to_stock.sku LIKE "35%") AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received <= 4 THEN "XS"
        WHEN (receipt_to_stock.sku LIKE "20%" OR receipt_to_stock.sku LIKE "32%" OR receipt_to_stock.sku LIKE "33%" OR receipt_to_stock.sku LIKE "35%") AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received <= 8 THEN "SM"
        WHEN (receipt_to_stock.sku LIKE "20%" OR receipt_to_stock.sku LIKE "32%" OR receipt_to_stock.sku LIKE "33%" OR receipt_to_stock.sku LIKE "35%") AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received <= 16 THEN "MD"
        WHEN (receipt_to_stock.sku LIKE "20%" OR receipt_to_stock.sku LIKE "32%" OR receipt_to_stock.sku LIKE "33%" OR receipt_to_stock.sku LIKE "35%") AND (old_bin IN ("", "NA") OR old_bin LIKE "%DROPSHIP") AND qty_received > 16 THEN "LG"
        
        #With Previous Quantity
        WHEN (receipt_to_stock.sku LIKE "01%" OR receipt_to_stock.sku LIKE "15%" OR receipt_to_stock.sku LIKE "19%") AND old_bin NOT IN ("", "NA", "DROPSHIP") AND qty_received + ns_inventory.stock > 10 AND invrec.location_table.bin_size <> "BALL" THEN "BALL"
        WHEN (receipt_to_stock.sku LIKE "02%" OR receipt_to_stock.sku LIKE "03%" OR receipt_to_stock.sku LIKE "04%" OR receipt_to_stock.sku LIKE "07%" OR receipt_to_stock.sku LIKE "08%" OR (receipt_to_stock.sku LIKE "09%" or receipt_to_stock.sku LIKE "10%" AND (receipt_to_stock.description LIKE "%combo irons%" OR receipt_to_stock.description LIKE "%complete set%"))) AND old_bin NOT IN ("", "NA", "DROPSHIP") AND qty_received + ns_inventory.stock > 5 AND invrec.location_table.bin_size <> "BULK" THEN "BULK"
        WHEN receipt_to_stock.sku LIKE "06%" AND old_bin NOT IN ("", "NA", "DROPSHIP") AND qty_received + ns_inventory.stock <= 40 AND invrec.location_table.bin_size <> "SM" THEN  "SM"
        WHEN receipt_to_stock.sku LIKE "06%" AND old_bin NOT IN ("", "NA", "DROPSHIP") AND qty_received + ns_inventory.stock <= 90 AND invrec.location_table.bin_size <> "MD" THEN "MD"
        WHEN receipt_to_stock.sku LIKE "06%" AND old_bin NOT IN ("", "NA", "DROPSHIP") AND qty_received + ns_inventory.stock > 90 AND invrec.location_table.bin_size <> "LG" THEN "LG"
        WHEN (receipt_to_stock.sku LIKE "09%" OR receipt_to_stock.sku LIKE "10%" OR receipt_to_stock.sku LIKE "11%" OR receipt_to_stock.sku LIKE "12%" OR receipt_to_stock.sku LIKE "14%" OR receipt_to_stock.sku LIKE "21%") AND old_bin NOT IN ("", "NA", "DROPSHIP") OR (receipt_to_stock.sku LIKE "19%" AND (receipt_to_stock.description LIKE "%retriever%")) AND qty_received + ns_inventory.stock > 4 AND invrec.location_table.bin_size <> "CL" THEN "CL"
        WHEN receipt_to_stock.sku LIKE "16%" AND old_bin NOT IN ("", "NA", "DROPSHIP") AND qty_received + ns_inventory.stock > 6 AND invrec.location_table.bin_size <> "SUN" THEN "SUN"
        WHEN receipt_to_stock.sku LIKE "18%" AND old_bin NOT IN ("", "NA", "DROPSHIP") AND qty_received + ns_inventory.stock > 9 AND invrec.location_table.bin_size <> "SHOE" THEN "SHOE"
        WHEN (receipt_to_stock.sku LIKE "20%" OR receipt_to_stock.sku LIKE "32%" OR receipt_to_stock.sku LIKE "33%" OR receipt_to_stock.sku LIKE "35%") AND old_bin NOT IN ("", "NA", "DROPSHIP") AND qty_received + ns_inventory.stock <= 8 AND invrec.location_table.bin_size <> "SM" THEN "SM"
        WHEN (receipt_to_stock.sku LIKE "20%" OR receipt_to_stock.sku LIKE "32%" OR receipt_to_stock.sku LIKE "33%" OR receipt_to_stock.sku LIKE "35%") AND old_bin NOT IN ("", "NA", "DROPSHIP") AND qty_received + ns_inventory.stock <= 16 AND invrec.location_table.bin_size <> "MD" THEN "MD"
        WHEN (receipt_to_stock.sku LIKE "20%" OR receipt_to_stock.sku LIKE "32%" OR receipt_to_stock.sku LIKE "33%" OR receipt_to_stock.sku LIKE "35%") AND old_bin NOT IN ("", "NA", "DROPSHIP") AND qty_received + ns_inventory.stock > 16 AND invrec.location_table.bin_size <> "LG" THEN "LG"

        END
;
/*
CALL looplocations();

DELIMITER $$

DROP PROCEDURE IF EXISTS loopLocations $$
CREATE PROCEDURE loopLocations() 

BEGIN
    DECLARE counter int(3) DEFAULT 0;

    WHILE (counter < 50) DO
        UPDATE receipt_to_stock
		SET receipt_to_stock.new_bin = (SELECT location FROM fl WHERE fl.bin_size = receipt_to_stock.bin_size AND fl.occupied = 0 ORDER BY RAND() LIMIT 1);

		UPDATE fl
		LEFT JOIN receipt_to_stock
		ON fl.location = receipt_to_stock.new_bin
		SET fl.occupied = 1
		WHERE receipt_to_stock.new_bin IS NOT NULL;

		UPDATE receipt_to_stock
		SET new_bin = NULL
		WHERE new_bin IN (SELECT new_bin FROM (SELECT new_bin FROM receipt_to_stock) as b GROUP BY new_bin HAVING COUNT(*) > 1);

		UPDATE fl
		LEFT JOIN receipt_to_stock
		ON fl.location = receipt_to_stock.new_bin
		SET fl.occupied = 0
		WHERE receipt_to_stock.new_bin IS NULL;


        set counter := counter +1;
    END WHILE;


END $$
*/


