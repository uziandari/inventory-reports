SELECT `main_inventory`.`sku`, `main_inventory`.`ns_description`, `main_inventory`.`stock`, `main_inventory`.`committed`, " _
  & "`main_inventory`.`upc`, `main_inventory`.`location`, `main_inventory`.`backstock` FROM `main_inventory` " _
  & "JOIN `invrec`.`location_table` ON `main_inventory`.`location` = `location_table`.`location` "
WHERE (stock < 11
AND `location_table`.bin_size IN ("LG", "BULK"))
OR (stock < 5
AND `location_table`.bin_size IN ("MD", "CL"))
OR (stock = 1
AND `location_table`.bin_size IN ("SM", "SHOE"))
ORDER BY `main_inventory`.location
;
