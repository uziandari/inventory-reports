SELECT dci.sku AS 'kit SKU', dci.description, di.sku AS 'base SKU', di.upc, di.location
FROM web_inventory.ca_inventory dci
JOIN web_inventory.ns_inventory di
ON dci.sku LIKE CONCAT(di.SKU, "%")
WHERE dci.sku LIKE "%OPEN"
OR dci.sku LIKE "%LOGO"
OR dci.sku LIKE "%KIT"
;