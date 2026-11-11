/*
# Create accumulated_scores and accumulated_scores_backup tables

1. New Tables
- `accumulated_scores`
  - `id` (uuid, primary key, auto-generated)
  - `employee_code` (text, not null) — รหัสพนักงาน
  - `employee_name` (text, not null) — ชื่อ-นามสกุล
  - `department` (text) — หน่วยงาน
  - `total_score` (numeric, default 0) — คะแนนสะสมรวม
  - `import_date` (date, not null) — วันที่นำเข้าข้อมูล
  - `file_name` (text) — ชื่อไฟล์ Excel ที่นำเข้า
  - `updated_at` (timestamptz, default now()) — เวลาอัปเดตล่าสุด
  - Unique constraint on employee_code เพื่อป้องกันข้อมูลซ้ำ

- `accumulated_scores_backup`
  - โครงสร้างเหมือน accumulated_scores ทุกประการ
  - เพิ่ม `backup_id` (uuid, primary key) และ `backup_created_at` (timestamptz)
  - ใช้สำหรับเก็บข้อมูลก่อนการ import ใหม่ (ระบบ Backup อัตโนมัติ)

2. Indexes
- Index บน employee_code สำหรับค้นหาเร็ว
- Index บน employee_name สำหรับค้นหาแบบ ILIKE
- Index บน department สำหรับค้นหาแบบ ILIKE

3. Security
- Enable RLS บนทั้งสองตาราง
- Allow anon + authenticated full CRUD (single-tenant, no Supabase auth)
- ข้อมูลนี้แยกออกจาก waste_entries อย่างเด็ดขาด ไม่มี FK เชื่อมโยง

4. Notes
- ตารางนี้ใช้สำหรับเก็บคะแนนสะสมจาก Excel เพื่อค้นหาเท่านั้น
- ไม่มีผลกระทบใดๆ ต่อข้อมูล waste_entries (Daily Records)
*/

CREATE TABLE IF NOT EXISTS accumulated_scores (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  employee_code text NOT NULL,
  employee_name text NOT NULL,
  department text,
  total_score numeric NOT NULL DEFAULT 0,
  import_date date NOT NULL,
  file_name text,
  updated_at timestamptz DEFAULT now()
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_accumulated_scores_employee_code
  ON accumulated_scores (employee_code);

CREATE INDEX IF NOT EXISTS idx_accumulated_scores_employee_name
  ON accumulated_scores (employee_name);

CREATE INDEX IF NOT EXISTS idx_accumulated_scores_department
  ON accumulated_scores (department);

ALTER TABLE accumulated_scores ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anon_select_accumulated_scores" ON accumulated_scores;
CREATE POLICY "anon_select_accumulated_scores" ON accumulated_scores FOR SELECT
  TO anon, authenticated USING (true);

DROP POLICY IF EXISTS "anon_insert_accumulated_scores" ON accumulated_scores;
CREATE POLICY "anon_insert_accumulated_scores" ON accumulated_scores FOR INSERT
  TO anon, authenticated WITH CHECK (true);

DROP POLICY IF EXISTS "anon_update_accumulated_scores" ON accumulated_scores;
CREATE POLICY "anon_update_accumulated_scores" ON accumulated_scores FOR UPDATE
  TO anon, authenticated USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "anon_delete_accumulated_scores" ON accumulated_scores;
CREATE POLICY "anon_delete_accumulated_scores" ON accumulated_scores FOR DELETE
  TO anon, authenticated USING (true);

-- Backup table (same structure + backup metadata)
CREATE TABLE IF NOT EXISTS accumulated_scores_backup (
  backup_id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  backup_created_at timestamptz DEFAULT now(),
  id uuid,
  employee_code text NOT NULL,
  employee_name text NOT NULL,
  department text,
  total_score numeric NOT NULL DEFAULT 0,
  import_date date NOT NULL,
  file_name text,
  updated_at timestamptz DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_accumulated_scores_backup_code
  ON accumulated_scores_backup (employee_code);

CREATE INDEX IF NOT EXISTS idx_accumulated_scores_backup_created
  ON accumulated_scores_backup (backup_created_at);

ALTER TABLE accumulated_scores_backup ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anon_select_accumulated_scores_backup" ON accumulated_scores_backup;
CREATE POLICY "anon_select_accumulated_scores_backup" ON accumulated_scores_backup FOR SELECT
  TO anon, authenticated USING (true);

DROP POLICY IF EXISTS "anon_insert_accumulated_scores_backup" ON accumulated_scores_backup;
CREATE POLICY "anon_insert_accumulated_scores_backup" ON accumulated_scores_backup FOR INSERT
  TO anon, authenticated WITH CHECK (true);

DROP POLICY IF EXISTS "anon_delete_accumulated_scores_backup" ON accumulated_scores_backup;
CREATE POLICY "anon_delete_accumulated_scores_backup" ON accumulated_scores_backup FOR DELETE
  TO anon, authenticated USING (true);
