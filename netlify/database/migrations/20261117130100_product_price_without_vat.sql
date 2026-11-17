ALTER TABLE products ADD COLUMN IF NOT EXISTS sale_price_without_vat NUMERIC(16,4) NOT NULL DEFAULT 0;
UPDATE products SET sale_price_without_vat = CASE WHEN vat > 0 THEN ROUND((sale_price / (1 + vat / 100.0))::numeric, 4) ELSE sale_price END WHERE sale_price_without_vat = 0 AND sale_price <> 0;
