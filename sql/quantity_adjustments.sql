TRUNCATE invrec.daily_adjustment;

INSERT INTO invrec.daily_adjustment (SKU, QTY)
SELECT SKU, NS_Adj_Req
FROM invrec.quantity_adjustments qa
WHERE DATE(Date) = CURDATE()  
;

UPDATE invrec.daily_adjustment
SET invrec.daily_adjustment.`Inventory Adjustment #` = 2522 /*Add new soft goods code*/
WHERE invrec.daily_adjustment.SKU LIKE "01%"
OR invrec.daily_adjustment.SKU LIKE "02%"
OR invrec.daily_adjustment.SKU LIKE "03%"
OR invrec.daily_adjustment.SKU LIKE "04%"
OR invrec.daily_adjustment.SKU LIKE "05%"
OR invrec.daily_adjustment.SKU LIKE "06%"
OR invrec.daily_adjustment.SKU LIKE "16%"
OR invrec.daily_adjustment.SKU LIKE "17%"
OR invrec.daily_adjustment.SKU LIKE "18%"
OR invrec.daily_adjustment.SKU LIKE "19%"
OR invrec.daily_adjustment.SKU LIKE "20%"
OR invrec.daily_adjustment.SKU LIKE "30%"
OR invrec.daily_adjustment.SKU LIKE "31%"
OR invrec.daily_adjustment.SKU LIKE "32%"
OR invrec.daily_adjustment.SKU LIKE "33%"
OR invrec.daily_adjustment.SKU LIKE "34%"
OR invrec.daily_adjustment.SKU LIKE "35%"
OR invrec.daily_adjustment.SKU LIKE "36%"
;

UPDATE invrec.daily_adjustment
SET invrec.daily_adjustment.`Inventory Adjustment #` = 2523 /*add new hard goods code */
WHERE invrec.daily_adjustment.SKU LIKE "07%"
OR invrec.daily_adjustment.SKU LIKE "08%"
OR invrec.daily_adjustment.SKU LIKE "09%"
OR invrec.daily_adjustment.SKU LIKE "10%"
OR invrec.daily_adjustment.SKU LIKE "11%"
OR invrec.daily_adjustment.SKU LIKE "12%"
OR invrec.daily_adjustment.SKU LIKE "13%"
OR invrec.daily_adjustment.SKU LIKE "14%"
OR invrec.daily_adjustment.SKU LIKE "15%"
OR invrec.daily_adjustment.SKU LIKE "16%"
OR invrec.daily_adjustment.SKU LIKE "21%"
OR invrec.daily_adjustment.SKU LIKE "6%"
;