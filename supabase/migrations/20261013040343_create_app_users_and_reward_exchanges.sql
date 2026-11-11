/*
# Create app_users and reward_exchanges tables

1. New Tables
- `app_users` — stores system users (admins and regular users)
- `reward_exchanges` — log of reward redemptions

2. Security
- RLS enabled on both, anon+authenticated CRUD (app uses own login, not Supabase Auth)
*/

CREATE TABLE IF NOT EXISTS app_users (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  username text UNIQUE NOT NULL,
  password_hash text NOT NULL,
  full_name text NOT NULL,
  role text NOT NULL DEFAULT 'user',
  department text,
  is_active boolean NOT NULL DEFAULT true,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

ALTER TABLE app_users ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anon_select_app_users" ON app_users;
CREATE POLICY "anon_select_app_users" ON app_users FOR SELECT
  TO anon, authenticated USING (true);

DROP POLICY IF EXISTS "anon_insert_app_users" ON app_users;
CREATE POLICY "anon_insert_app_users" ON app_users FOR INSERT
  TO anon, authenticated WITH CHECK (true);

DROP POLICY IF EXISTS "anon_update_app_users" ON app_users;
CREATE POLICY "anon_update_app_users" ON app_users FOR UPDATE
  TO anon, authenticated USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "anon_delete_app_users" ON app_users;
CREATE POLICY "anon_delete_app_users" ON app_users FOR DELETE
  TO anon, authenticated USING (true);

CREATE TABLE IF NOT EXISTS reward_exchanges (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  employee_code text,
  employee_name text NOT NULL,
  reward_item text NOT NULL,
  quantity numeric NOT NULL DEFAULT 0,
  points_used numeric NOT NULL DEFAULT 0,
  exchanged_by text,
  exchanged_at timestamptz DEFAULT now()
);

ALTER TABLE reward_exchanges ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anon_select_reward_exchanges" ON reward_exchanges;
CREATE POLICY "anon_select_reward_exchanges" ON reward_exchanges FOR SELECT
  TO anon, authenticated USING (true);

DROP POLICY IF EXISTS "anon_insert_reward_exchanges" ON reward_exchanges;
CREATE POLICY "anon_insert_reward_exchanges" ON reward_exchanges FOR INSERT
  TO anon, authenticated WITH CHECK (true);

DROP POLICY IF EXISTS "anon_update_reward_exchanges" ON reward_exchanges;
CREATE POLICY "anon_update_reward_exchanges" ON reward_exchanges FOR UPDATE
  TO anon, authenticated USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "anon_delete_reward_exchanges" ON reward_exchanges;
CREATE POLICY "anon_delete_reward_exchanges" ON reward_exchanges FOR DELETE
  TO anon, authenticated USING (true);

INSERT INTO app_users (username, password_hash, full_name, role, department)
SELECT 'admin', encode(convert_to(reverse('admin123'), 'UTF8'), 'base64'), 'Admin', 'admin', 'System'
WHERE NOT EXISTS (SELECT 1 FROM app_users WHERE username = 'admin');