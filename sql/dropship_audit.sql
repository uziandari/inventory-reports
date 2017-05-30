SELECT * FROM main_inventory
LEFT JOIN dropship
ON main_inventory.sku = dropship.sku
WHERE main_inventory.sku IN ();