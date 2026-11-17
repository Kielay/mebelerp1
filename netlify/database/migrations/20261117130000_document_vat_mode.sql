ALTER TABLE documents ADD COLUMN IF NOT EXISTS vat_mode TEXT NOT NULL DEFAULT 'inclusive' CHECK (vat_mode IN ('inclusive','exclusive','none'));
