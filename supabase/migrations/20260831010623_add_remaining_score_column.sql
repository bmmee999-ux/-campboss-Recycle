/*
# Add remaining_score column to accumulated_scores and backup tables

1. Modified Tables
- `accumulated_scores`: add `remaining_score` (numeric, default 0) — คะแนนคงเหลือ
- `accumulated_scores_backup`: add `remaining_score` (numeric, default 0) — คะแนนคงเหลือ

2. Security
- No RLS policy changes needed; existing policies already cover all columns.

3. Notes
- The new column stores the remaining score after exchanges (คะแนนที่เหลือจากการหัก).
- Default is 0 so existing rows are not affected.
*/

DO $$ BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_name = 'accumulated_scores' AND column_name = 'remaining_score'
  ) THEN
    ALTER TABLE accumulated_scores ADD COLUMN remaining_score numeric NOT NULL DEFAULT 0;
  END IF;
END $$;

DO $$ BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_name = 'accumulated_scores_backup' AND column_name = 'remaining_score'
  ) THEN
    ALTER TABLE accumulated_scores_backup ADD COLUMN remaining_score numeric NOT NULL DEFAULT 0;
  END IF;
END $$;
