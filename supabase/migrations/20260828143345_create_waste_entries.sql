/*
# Create waste_entries table (single-tenant, no Supabase auth)

1. New Tables
- `waste_entries`
  - `id` (uuid, primary key, auto-generated)
  - `date` (text, not null) — the entry date as YYYY-MM-DD
  - `code` (text, nullable) — employee/participant code
  - `name` (text, not null) — participant full name
  - `unit_org` (text, nullable) — department/organization
  - `room` (text, not null) — room label or "ห้องรวม" for totals
  - `waste_type` (text, not null) — waste type key
  - `quantity` (text, nullable) — quantity as string
  - `qty_unit` (text, nullable) — unit of quantity
  - `points` (text, nullable) — calculated points as string
  - `reward` (text, nullable) — reward text
  - `signature` (text, nullable) — signature
  - `created_at` (timestamptz, default now()) — record creation time

2. Security
- Enable RLS on `waste_entries`.
- Allow anon + authenticated full CRUD because the app uses a simple
  client-side login (not Supabase auth) and the data is intentionally
  shared among all program operators.

3. Notes
- The app inserts rows via the anon-key Supabase client, so anon policies
  are required for the app to function.
- Realtime is enabled via supabase channel subscription in the frontend.
*/

CREATE TABLE IF NOT EXISTS waste_entries (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  date text NOT NULL,
  code text,
  name text NOT NULL,
  unit_org text,
  room text NOT NULL,
  waste_type text NOT NULL,
  quantity text,
  qty_unit text,
  points text,
  reward text,
  signature text,
  created_at timestamptz DEFAULT now()
);

ALTER TABLE waste_entries ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anon_select_waste_entries" ON waste_entries;
CREATE POLICY "anon_select_waste_entries" ON waste_entries FOR SELECT
  TO anon, authenticated USING (true);

DROP POLICY IF EXISTS "anon_insert_waste_entries" ON waste_entries;
CREATE POLICY "anon_insert_waste_entries" ON waste_entries FOR INSERT
  TO anon, authenticated WITH CHECK (true);

DROP POLICY IF EXISTS "anon_update_waste_entries" ON waste_entries;
CREATE POLICY "anon_update_waste_entries" ON waste_entries FOR UPDATE
  TO anon, authenticated USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "anon_delete_waste_entries" ON waste_entries;
CREATE POLICY "anon_delete_waste_entries" ON waste_entries FOR DELETE
  TO anon, authenticated USING (true);
