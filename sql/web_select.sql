TRUNCATE inventory.main_inventory;

#DATE STRING FOR FLAGS
SET @today = CONCAT(MONTH(CURDATE()), "/", DAY(CURDATE()));


INSERT INTO inventory.main_inventory (`sku`,
`description`,
`total`,
`available`,
`pending_checkout`,
`pending_payment`,
`pending_shipment`,
`flag`,
`parent_sku`,
`label`,
`ns_description`,
`location`,
`backstock`,
`upc`,
`stock`,
`committed`,
`head_cover`,
`inline`,
`purchase_price`,
`receipt`,
`is_blocked`,
`img_url`)

SELECT ci.sku, ci.description, ci.total, ci.available, ci.pending_checkout, ci.pending_payment, ci.pending_shipment, ci.flag, ci.parent_sku,
		ci.label, ni.description, ni.location, ni.backstock, ni.upc, ni.stock, ni.`committed`, ni.head_cover, ni.inline, ni.purchase_price, rd.receipt, ci.blocked, ci.img_url
FROM inventory.ca_inventory ci
INNER JOIN inventory.ns_inventory ni
ON ci.sku = ni.sku
LEFT JOIN inventory.receipt_date rd
ON ci.sku = rd.sku
;

SET SQL_SAFE_UPDATES=0;

UPDATE inventory.main_inventory mi
SET mi.location = "NA"
WHERE mi.location = ""
;

UPDATE inventory.main_inventory mi
SET mi.inline = "No"
WHERE mi.inline = ""
;

/*
#DELETE inventory.removes and bucket
DELETE mi
FROM inventory.main_inventory mi
JOIN inventory.removes r ON mi.sku = r.sku WHERE r.sku IS NOT NULL;

DELETE mi
FROM inventory.main_inventory mi
JOIN bucket b ON mi.sku = b.sku WHERE b.sku IS NOT NULL;

*/

UPDATE inventory.main_inventory
JOIN inventory.bucket
ON inventory.main_inventory.sku = inventory.bucket.sku
SET inventory.main_inventory.bucket = 1
WHERE inventory.bucket.sku IS NOT NULL;

#Finds BC SKUs for john daily site
UPDATE inventory.main_inventory
JOIN inventory.big_commerce_daily ON inventory.main_inventory.sku = inventory.big_commerce_daily.sku
SET inventory.main_inventory.jdaily_quantity = 1
WHERE inventory.big_commerce_daily.sku IS NOT NULL;

UPDATE inventory.main_inventory
JOIN inventory.removes ON inventory.main_inventory.sku = inventory.removes.sku
SET inventory.main_inventory.remove = 1
WHERE inventory.removes.sku IS NOT NULL;

/* Removed 1/19/17
#UPDATE EBAY
UPDATE inventory.main_inventory
SET inventory.main_inventory.ebay = 1
WHERE flag LIKE "%ebay%"
;
*/

#UPDATE RELIST FLAGS
UPDATE inventory.main_inventory
INNER JOIN inventory.relists
ON inventory.main_inventory.sku = inventory.relists.sku
SET inventory.main_inventory.relist = 1
WHERE inventory.relists.sku IS NOT NULL
;

UPDATE inventory.main_inventory
SET inventory.main_inventory.relist = 1
WHERE inventory.main_inventory.flag LIKE CONCAT("final recount ", @today, "%")
;

#UPDATE inventory.main_inventory/daily inventory.removes
UPDATE inventory.main_inventory
INNER JOIN inventory.removes_daily
ON inventory.main_inventory.sku = inventory.removes_daily.sku
SET inventory.main_inventory.remove = 1
WHERE inventory.removes_daily.sku IS NOT NULL
;


#UPDATE LESS THAN NINE
UPDATE inventory.main_inventory
SET inventory.main_inventory.less_nine = 1
WHERE (available <= 9 and available > 0)
AND (flag NOT LIKE "%final%" AND flag NOT LIKE "%Inline%")
AND (location <> "NA" AND location <> "" AND location NOT LIKE "%Dropship%")
AND bucket = 0
/* Pull out gift cards --12/30 */
AND sku NOT LIKE "9%"
/* End office SKUs */
;

#UPDATE DELIST --updated 1/19 to reflect new changes to ebay
UPDATE inventory.main_inventory
SET inventory.main_inventory.delist = 1
WHERE available = 0
/* Pull out gift cards --12/30 */
AND sku NOT LIKE "9%"
/* End office SKUs */
AND flag NOT LIKE "%absolute final%"
AND location NOT IN  ("NA", "NA_", "", "Dropship")
AND description NOT LIKE ""
AND inline NOT LIKE "Yes%"
/* added 12/29 --attempting to only add used clubs to delist @ 0, not < 3 */
AND sku NOT LIKE "6%"
/* end 12/29 used club add */
AND stock = 0 # changed 1/19, new ebay policy
AND relist = 0
AND remove = 0
AND flag NOT LIKE "%recount%"


/*added 3/28 */
OR ( sku NOT LIKE "9%"
/* End office SKUs */
AND stock = committed
AND location NOT IN  ("NA", "NA_", "")
AND location NOT LIKE ("%DROPSHIP")
AND flag NOT LIKE "%absolute final%"
AND relist = 0
AND remove = 0
AND flag NOT LIKE "%recount%"
)
/* 3/28 */

OR (inline LIKE "Yes%"
/* Pull out gift cards --12/30 */
AND sku NOT LIKE "9%"
/* End office SKUs */
AND available = 0
AND stock > 0
AND stock = committed
#AND flag NOT LIKE "%ebay%"
AND flag NOT LIKE "%absolute final%"
AND relist = 0
AND remove = 0
AND flag NOT LIKE "%recount%"
)

#Experimental
OR (stock > 0
/* Pull out gift cards --12/30 */
AND sku NOT LIKE "9%"
/* End office SKUs */
AND stock = `committed`
AND flag NOT LIKE "%absolute final%"
AND location NOT LIKE "DROPSHIP"
AND relist = 0
AND remove = 0
AND flag NOT LIKE "%recount%")
;
# END EXPERIMENTAL

#UPDATE ALERTS
UPDATE inventory.main_inventory
SET inventory.main_inventory.alerts = 1

WHERE (available = 0
/* Pull out gift cards --12/30 */
AND sku NOT LIKE "9%"
/* End office SKUs */
AND flag NOT LIKE "%absolute final%"
AND flag NOT LIKE "%ebay%"
AND flag NOT LIKE "%hold%"
AND location NOT IN  ("NA", "NA_", "", "Dropship")
AND description NOT LIKE ""
AND flag NOT LIKE "%recount%"
#AND ((stock - `committed`) > (`committed` * 2) OR stock <= 0)
AND less_nine = 0
AND delist = 0
AND relist = 0
AND remove = 0
AND bucket = 0)
;
/* Removed for ebay change 1/20
OR (available = 0
# Pull out gift cards --12/30
AND sku NOT LIKE "9%"
# End office SKUs
AND inline LIKE "yes%"
AND location NOT IN  ("NA", "NA_", "", "Dropship")
AND description NOT LIKE ""
AND flag NOT LIKE "%absolute final%"
AND stock <= 0
AND less_nine = 0
AND delist = 0
AND relist = 0
AND remove = 0
AND bucket = 0
AND flag NOT LIKE "%recount%")
;
*/

#PULL PENDING FROM ALERTS
UPDATE inventory.main_inventory
SET inventory.main_inventory.alerts = 0
WHERE inventory.main_inventory.alerts = 1
AND stock > 0
AND (inventory.main_inventory.pending_checkout + inventory.main_inventory.pending_payment = stock);

#PULL COMMITTED FROM ALERTS WITH SAFETY STOCK
UPDATE inventory.main_inventory
SET inventory.main_inventory.alerts = 0
WHERE inventory.main_inventory.alerts = 1
AND stock > 0
AND (GREATEST(pending_shipment, committed) > 0)
AND (GREATEST(pending_shipment, committed) + 2 >= stock);

#RELIST
UPDATE inventory.main_inventory
SET inventory.main_inventory.relist_pushed = 1, inventory.main_inventory.relist = 0
WHERE (inventory.main_inventory.relist = 1
	AND (total <> 0
    OR `committed` <> 0
    OR bucket = 1))
;

#UPDATE WHERE DUPLICATE LOCATIONS
UPDATE inventory.main_inventory
INNER JOIN (SELECT location FROM inventory.main_inventory
WHERE stock > 0
AND location NOT LIKE "C%"
AND location NOT LIKE "Z CAB%"
AND location NOT LIKE "L%"
AND location NOT LIKE "M%"
AND location NOT LIKE "N%"
AND location NOT LIKE "O%"
AND location NOT LIKE "J%"
AND location NOT LIKE "K%"
AND location NOT LIKE "I24%"
AND location NOT LIKE "I26%"
AND location NOT LIKE "%dropship%"
AND location NOT LIKE "GNC%"
GROUP BY location HAVING COUNT(location) > 1) AS secondary_inventory
ON inventory.main_inventory.location = secondary_inventory.location
SET inventory.main_inventory.dupe_loc = 1
;

#UPDATE WHERE DUPLICATE UPCs
UPDATE inventory.main_inventory
INNER JOIN (SELECT upc FROM inventory.main_inventory
/* WHERE stock > 0 */
WHERE upc NOT LIKE ""
/* switch to digit matching
AND upc NOT LIKE "000000000001"
AND upc NOT LIKE "000000000010"
AND location NOT LIKE "%Dropship%" */
AND upc NOT LIKE "0000000000%%"
GROUP BY upc HAVING COUNT(upc) > 1) AS secondary_inventory
ON inventory.main_inventory.upc = secondary_inventory.upc
SET inventory.main_inventory.dupe_upc = 1;

UPDATE inventory.main_inventory mi
SET mi.location = UPPER( `location` )
;

UPDATE inventory.main_inventory mi
SET mi.location = "NA"
WHERE mi.location LIKE "";

#Find Committed Stock to Wholesale
UPDATE inventory.main_inventory
INNER JOIN inventory.wholesale_pending
ON inventory.main_inventory.sku = inventory.wholesale_pending.sku
SET inventory.main_inventory.wholesale_committed = 1
WHERE inventory.main_inventory.available = 0
AND inventory.main_inventory.location NOT IN  ("NA", "NA_", "DROPSHIP")
AND inventory.main_inventory.`committed` > 0
;

SET SQL_SAFE_UPDATES=1;

#Find Free Locations
TRUNCATE inventory.free_locations;
INSERT INTO inventory.free_locations (location, bin_size)
SELECT invrec.location_table.location, invrec.location_table.bin_size
FROM invrec.location_table
WHERE invrec.location_table.location NOT IN
    (SELECT ns_inventory.location
     FROM inventory.ns_inventory)
AND invrec.location_table.location NOT IN
	(SELECT inventory.location_removes.location
	FROM inventory.location_removes)
;
