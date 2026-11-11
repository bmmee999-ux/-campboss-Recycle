import * as XLSX from "xlsx";
import { supabase } from "@/integrations/supabase/client";
import { REWARDS, LEGACY_REWARD_NAMES } from "@/components/RewardExchange";

const PAGE_SIZE = 1000;
const BACKUP_KEY = "waste-rewards-local-backups";
const MAX_BACKUPS = 20;

async function fetchAll<T>(table: string, orderCol: string): Promise<T[]> {
  const all: T[] = [];
  let from = 0;
  let more = true;
  while (more) {
    const { data, error } = await supabase
      .from(table)
      .select("*")
      .order(orderCol, { ascending: true })
      .range(from, from + PAGE_SIZE - 1);
    if (error) throw error;
    const rows = (data ?? []) as T[];
    all.push(...rows);
    more = rows.length === PAGE_SIZE;
    from += PAGE_SIZE;
  }
  return all;
}

async function deleteAll(table: string): Promise<void> {
  let more = true;
  while (more) {
    const { data, error } = await supabase
      .from(table)
      .delete()
      .neq("id", "00000000-0000-0000-0000-000000000000")
      .select("id")
      .limit(PAGE_SIZE);
    if (error) throw error;
    more = (data ?? []).length === PAGE_SIZE;
  }
}

async function insertBatched<T>(table: string, rows: T[], batchSize = 500): Promise<void> {
  for (let i = 0; i < rows.length; i += batchSize) {
    const batch = rows.slice(i, i + batchSize);
    const { error } = await supabase.from(table).insert(batch);
    if (error) throw error;
  }
}

type ScoreRecord = {
  employee_code: string;
  employee_name: string;
  department: string | null;
  total_score: number;
  remaining_score: number;
};

type ExchangeRecord = {
  reward_item: string;
  quantity: number;
  points_used: number;
};

const round2 = (n: number) => Math.round(n * 100) / 100;

const fmtInt = (n: number) => Math.round(n).toLocaleString("en-US");

// Same layout as Easy_To_Read_Waste_Summary.xlsx: merged cells, headings and live formulas.
// All numbers are calculated from the database at the moment of backup.
function buildSummarySheet(scores: ScoreRecord[], exchanges: ExchangeRecord[]): XLSX.WorkSheet {
  const totalPoints = scores.reduce((sum, r) => sum + Number(r.total_score), 0);
  const remainingPoints = scores.reduce((sum, r) => sum + Number(r.remaining_score), 0);
  const redeemedPoints = totalPoints - remainingPoints;
  const departmentCount = new Set(scores.map((r) => (r.department ?? "").trim()).filter(Boolean)).size;

  // Aggregate redemptions per reward (old names like "ไข่" count as "ไข่ไก่")
  const byReward = new Map<string, { qty: number; points: number }>();
  for (const ex of exchanges) {
    const name = LEGACY_REWARD_NAMES[ex.reward_item] ?? ex.reward_item;
    const agg = byReward.get(name) ?? { qty: 0, points: 0 };
    agg.qty += Number(ex.quantity);
    agg.points += Number(ex.points_used);
    byReward.set(name, agg);
  }

  // Rows: every reward in the attachment's order (0 when never redeemed), then any unknown extras
  type Line = { name: string; qty: number; cost: number; unit: string; showSub: boolean };
  const lines: Line[] = REWARDS.map((rw, i) => ({
    name: rw.label.th,
    qty: byReward.get(rw.label.th)?.qty ?? 0,
    cost: rw.cost,
    unit: rw.unit.th,
    showSub: i < 3, // the attachment fills "จำนวน/หน่วยย่อย" only for the first three items
  }));
  const known = new Set(REWARDS.map((r) => r.label.th));
  for (const [name, agg] of byReward) {
    if (known.has(name)) continue;
    lines.push({ name, qty: agg.qty, cost: agg.qty ? round2(agg.points / agg.qty) : 0, unit: "", showSub: false });
  }

  const HEAD = 10; // header row (1-based), data starts at HEAD + 1
  const first = HEAD + 1;
  const last = HEAD + lines.length;
  const totalRow = last + 1;

  const aoa: (string | number | null)[][] = [];
  aoa[0] = ["สรุปภาพรวมโครงการขยะสะสมคะแนน แลกของรางวัล"];
  aoa[2] = ["ภาพรวมสถิติโครงการ"];
  aoa[3] = ["พนักงานที่เข้าร่วม", null, "หน่วยงานที่เข้าร่วม", null, "คะแนนที่เหลือยังไม่แลก"];
  aoa[4] = [
    `${fmtInt(scores.length)} คน`,
    null,
    `${fmtInt(departmentCount)} หน่วยงาน`,
    null,
    `${fmtInt(remainingPoints)} คะแนน`,
  ];
  aoa[6] = [
    `   📊 คะแนนสะสมทั้งหมด: ${fmtInt(totalPoints)} คะแนน     |     🎁 แลกของรางวัลไปแล้ว: ${fmtInt(redeemedPoints)} คะแนน      |     ⭐ คงเหลือ: ${fmtInt(remainingPoints)} คะแนน`,
  ];
  aoa[8] = ["รายละเอียดการแลกของรางวัลแยกตามประเภท"];
  aoa[HEAD - 1] = ["ลำดับ", "รายการของรางวัล", "จำนวน/หน่วยย่อย", "จำนวนที่แลก", "เกณฑ์คะแนน", "รวมคะแนนที่แลก", "หน่วย"];
  lines.forEach((l, i) => {
    aoa[HEAD + i] = [i + 1, l.name, l.showSub ? l.qty : null, l.qty, l.cost, null, l.unit];
  });
  aoa[totalRow - 1] = ["รวมคะแนนที่แลกของรางวัล ทั้งหมด", null, null, null, null, null, "คะแนน"];
  for (let r = 0; r < aoa.length; r++) if (!aoa[r]) aoa[r] = [];

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // live formulas, like the attachment
  lines.forEach((_, i) => {
    const row = first + i;
    ws[`F${row}`] = { t: "n", f: `D${row}*E${row}`, v: lines[i].qty * lines[i].cost, z: "#,##0" };
    ws[`D${row}`] && (ws[`D${row}`].z = "#,##0");
    ws[`E${row}`] && (ws[`E${row}`].z = "#,##0");
    ws[`C${row}`] && (ws[`C${row}`].z = "#,##0");
  });
  ws[`F${totalRow}`] = {
    t: "n",
    f: `SUM(F${first}:F${last})`,
    v: lines.reduce((s, l) => s + l.qty * l.cost, 0),
    z: "#,##0",
  };

  const m = (s: string) => XLSX.utils.decode_range(s);
  ws["!merges"] = [
    m("A1:G1"),
    m("A4:B4"), m("C4:D4"), m("E4:F4"),
    m("A5:B5"), m("C5:D5"), m("E5:F5"),
    m("A7:G7"),
    m(`A${totalRow}:E${totalRow}`),
  ];
  ws["!cols"] = [{ wch: 10 }, { wch: 34 }, { wch: 18 }, { wch: 16 }, { wch: 14 }, { wch: 18 }, { wch: 12 }];
  return ws;
}

// ===== Department reports (same layout as the signed PDF reports) =====

const compact = (v: string | null | undefined) => (v ?? "").replace(/\s+/g, "").toLowerCase();

// Table 2: the 20 "หน่วยงาน" accounts. They are stored in accumulated_scores as rows whose
// employee_name is the unit name. `match` is compared with whitespace removed + lowercase.
const UNIT_ACCOUNTS: { display: string; match: string }[] = [
  { display: "Welfare Affair (Canteen )", match: "welfareaffair(canteen)" },
  { display: "Camp Affair", match: "campaffair" },
  { display: "IT Center", match: "itcenter" },
  { display: "Laboratory", match: "laboratory" },
  { display: "Warehouse&Fuel (น้ำมัน)", match: "warehouse&fuel(น้ำมัน)" },
  { display: "บริษัทสิงห์แก้ว", match: "บริษัทสิงห์แก้ว" },
  { display: "Environment", match: "environment" },
  { display: "Welfare Affair (convenient)", match: "welfareaffair(convenient)" },
  { display: "Process Material", match: "processmaterial" },
  { display: "Site Agent 2", match: "siteagent2" },
  { display: "General Affair", match: "generalaffair" },
  { display: "First Aid", match: "firstaid" },
  { display: "Asset Control", match: "assetcontrol" },
  { display: "Account & Finance", match: "account&finance" },
  { display: "หน่วยงาน QC", match: "หน่วยงานqc" },
  { display: "บริษัท Whoessoe", match: "บริษัทwhoessoe" },
  { display: "หน่วยงาน E&M", match: "หน่วยงานe&m" },
  { display: "หน่วยงาน Package 1", match: "หน่วยงานpackage1" },
  { display: "Andritz", match: "andritz" },
  { display: "Property Control", match: "หน่วยงาน" },
];

// Table 1: staff groups, in the order of the report. Any other department is a contractor.
const STAFF_GROUPS = [
  "Engineering", "PQM", "Administration", "Process Material", "E&M",
  "Site Agent 1", "Site Agent 2", "Site Agent 3",
];
const CONTRACTOR_GROUP = "ผู้รับเหมา";
const DEPT_ALIASES: Record<string, string> = { pm: "PQM" }; // "PM" is part of PQM

function staffGroupOf(dept: string | null | undefined): string {
  const key = compact(dept);
  if (!key) return "";
  const alias = DEPT_ALIASES[key];
  if (alias) return alias;
  const hit = STAFF_GROUPS.find((g) => compact(g) === key);
  return hit ?? CONTRACTOR_GROUP;
}

const thaiDate = (d: Date) =>
  `${String(d.getDate()).padStart(2, "0")}.${String(d.getMonth() + 1).padStart(2, "0")}.${d.getFullYear()}`;

const NUM2 = "#,##0.00";

type Acc = { count: number; total: number; remaining: number };

function splitScores(scores: ScoreRecord[]) {
  const unitMap = new Map<string, Acc>();
  const staffMap = new Map<string, Acc>();
  for (const r of scores) {
    const total = Number(r.total_score) || 0;
    const remaining = Number(r.remaining_score) || 0;
    const unit = UNIT_ACCOUNTS.find((u) => u.match === compact(r.employee_name));
    const group = unit ? unit.display : staffGroupOf(r.department) || "ไม่ระบุหน่วยงาน";
    const map = unit ? unitMap : staffMap;
    const a = map.get(group) ?? { count: 0, total: 0, remaining: 0 };
    a.count += 1;
    a.total += total;
    a.remaining += remaining;
    map.set(group, a);
  }
  return { unitMap, staffMap };
}

const STAFF_SHEET = "จำนวนพนักงาน+คะแนน";
const UNIT_SHEET = "คะแนนหน่วยงาน";

function buildStaffSheet(staffMap: Map<string, Acc>, date: string): { ws: XLSX.WorkSheet; totalRow: number } {
  const order = [...STAFF_GROUPS, CONTRACTOR_GROUP];
  if (staffMap.has("ไม่ระบุหน่วยงาน")) order.push("ไม่ระบุหน่วยงาน");
  const aoa: (string | number | null)[][] = [
    [`รายงานจำนวนพนักงานที่เข้าร่วมโครงการขยะสะสมคะแนน และ คะแนนที่สะสมที่ยังไม่ได้แลก Update ${date}`],
    ["ตารางคะแนน พนักงานที่เข้าร่วมกิจกรรม โครงการขยะสะสมคะแนน แลกรางวัล"],
    ["ลำดับ", "หน่วยงาน", "จำนวนพนักงาน", "คะแนนที่สะสม", "คะแนนที่แลกไปแล้ว", "คะแนนที่ยังไม่ได้แลก"],
  ];
  const first = aoa.length + 1;
  order.forEach((g, i) => {
    const a = staffMap.get(g) ?? { count: 0, total: 0, remaining: 0 };
    aoa.push([i + 1, g, a.count, round2(a.total), null, round2(a.remaining)]);
  });
  const last = aoa.length;
  const totalRow = last + 1;
  aoa.push(["รวมจำนวนพนักงาน และ คะแนน", null, null, null, null, null]);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  order.forEach((g, i) => {
    const row = first + i;
    const a = staffMap.get(g) ?? { count: 0, total: 0, remaining: 0 };
    ws[`E${row}`] = { t: "n", f: `D${row}-F${row}`, v: round2(a.total - a.remaining), z: NUM2 };
    ws[`C${row}`].z = "#,##0";
    ws[`D${row}`].z = NUM2;
    ws[`F${row}`].z = NUM2;
  });
  const sum = (col: string, f: (a: Acc) => number, z: string) => {
    const v = round2(order.reduce((t, g) => t + f(staffMap.get(g) ?? { count: 0, total: 0, remaining: 0 }), 0));
    ws[`${col}${totalRow}`] = { t: "n", f: `SUM(${col}${first}:${col}${last})`, v, z };
  };
  sum("C", (a) => a.count, "#,##0");
  sum("D", (a) => a.total, NUM2);
  sum("E", (a) => a.total - a.remaining, NUM2);
  sum("F", (a) => a.remaining, NUM2);
  ws["!merges"] = [XLSX.utils.decode_range("A1:F1"), XLSX.utils.decode_range("A2:F2"), XLSX.utils.decode_range(`A${totalRow}:B${totalRow}`)];
  ws["!cols"] = [{ wch: 8 }, { wch: 28 }, { wch: 16 }, { wch: 18 }, { wch: 20 }, { wch: 22 }];
  return { ws, totalRow };
}

function buildUnitSheet(unitMap: Map<string, Acc>, staffTotalRow: number, staffTotals: Acc, date: string): XLSX.WorkSheet {
  const aoa: (string | number | null)[][] = [
    [`รายงานจำนวนพนักงานที่เข้าร่วมโครงการขยะสะสมคะแนน และ คะแนนที่สะสมที่ยังไม่ได้แลก  Update ${date}`],
    ["รายงานหน่วยงานที่เข้าร่วมกิจกรรม"],
    ["ลำดับ", "หน่วยงาน", "คะแนนที่สะสม", "คะแนนที่แลกไปแล้ว", "คะแนนที่ยังไม่ได้แลก"],
  ];
  const first = aoa.length + 1;
  UNIT_ACCOUNTS.forEach((u, i) => {
    const a = unitMap.get(u.display) ?? { count: 0, total: 0, remaining: 0 };
    aoa.push([i + 1, u.display, round2(a.total), null, round2(a.remaining)]);
  });
  const last = aoa.length;
  const unitTotalRow = last + 1;
  const grandRow = last + 2;
  aoa.push(["รวมคะแนนหน่วยงาน", null, null, null, null]);
  aoa.push(["รวมคะแนนทั้งหมด", null, null, null, null]);
  const ws = XLSX.utils.aoa_to_sheet(aoa);

  const unitAcc = UNIT_ACCOUNTS.map((u) => unitMap.get(u.display) ?? { count: 0, total: 0, remaining: 0 });
  unitAcc.forEach((a, i) => {
    const row = first + i;
    ws[`D${row}`] = { t: "n", f: `C${row}-E${row}`, v: round2(a.total - a.remaining), z: NUM2 };
    ws[`C${row}`].z = NUM2;
    ws[`E${row}`].z = NUM2;
  });
  const uTotal = unitAcc.reduce((t, a) => t + a.total, 0);
  const uRem = unitAcc.reduce((t, a) => t + a.remaining, 0);
  const put = (cell: string, f: string, v: number) => (ws[cell] = { t: "n", f, v: round2(v), z: NUM2 });
  put(`C${unitTotalRow}`, `SUM(C${first}:C${last})`, uTotal);
  put(`D${unitTotalRow}`, `SUM(D${first}:D${last})`, uTotal - uRem);
  put(`E${unitTotalRow}`, `SUM(E${first}:E${last})`, uRem);
  // grand total = unit accounts + staff groups (formulas point at the staff sheet)
  const ref = (col: string) => `'${STAFF_SHEET}'!${col}${staffTotalRow}`;
  put(`C${grandRow}`, `C${unitTotalRow}+${ref("D")}`, uTotal + staffTotals.total);
  put(`D${grandRow}`, `D${unitTotalRow}+${ref("E")}`, uTotal - uRem + (staffTotals.total - staffTotals.remaining));
  put(`E${grandRow}`, `E${unitTotalRow}+${ref("F")}`, uRem + staffTotals.remaining);

  ws["!merges"] = [
    XLSX.utils.decode_range("A1:E1"),
    XLSX.utils.decode_range("A2:E2"),
    XLSX.utils.decode_range(`A${unitTotalRow}:B${unitTotalRow}`),
    XLSX.utils.decode_range(`A${grandRow}:B${grandRow}`),
  ];
  ws["!cols"] = [{ wch: 8 }, { wch: 32 }, { wch: 18 }, { wch: 20 }, { wch: 22 }];
  return ws;
}

export async function exportFullBackup(): Promise<void> {
  const [entries, scores, scoreBackups, exchanges] = await Promise.all([
    fetchAll("waste_entries", "created_at"),
    fetchAll<ScoreRecord>("accumulated_scores", "employee_code"),
    fetchAll("accumulated_scores_backup", "backup_created_at"),
    fetchAll<ExchangeRecord>("reward_exchanges", "exchanged_at"),
  ]);

  const today = new Date().toISOString().slice(0, 10);
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, buildSummarySheet(scores, exchanges), "สรุปภาพรวมโครงการขยะ");

  const { unitMap, staffMap } = splitScores(scores);
  const reportDate = thaiDate(new Date());
  const { ws: staffWs, totalRow: staffTotalRow } = buildStaffSheet(staffMap, reportDate);
  const staffTotals = [...staffMap.values()].reduce(
    (t, a) => ({ count: t.count + a.count, total: t.total + a.total, remaining: t.remaining + a.remaining }),
    { count: 0, total: 0, remaining: 0 },
  );
  XLSX.utils.book_append_sheet(wb, staffWs, STAFF_SHEET);
  XLSX.utils.book_append_sheet(wb, buildUnitSheet(unitMap, staffTotalRow, staffTotals, reportDate), UNIT_SHEET);

  const employees = scores.map((s) => {
    const r = s as Record<string, unknown>;
    return {
      Code: r.employee_code ?? "",
      Name: r.employee_name ?? "",
      Department: r.department ?? "",
      TotalPoints: r.total_score ?? 0,
      RemainingPoints: r.remaining_score ?? 0,
      ImportDate: r.import_date ?? "",
    };
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(employees), "Employees");

  const transactions = entries.map((e) => {
    const r = e as Record<string, unknown>;
    return {
      Date: r.date ?? "",
      Code: r.code ?? "",
      Name: r.name ?? "",
      Department: r.unit_org ?? "",
      Room: r.room ?? "",
      WasteType: r.waste_type ?? "",
      Quantity: r.quantity ?? "",
      Unit: r.qty_unit ?? "",
      Points: r.points ?? "",
      Reward: r.reward ?? "",
      Signature: r.signature ?? "",
      CreatedAt: r.created_at ?? "",
    };
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(transactions), "DailyTransactions");

  const rewardHistory = scoreBackups.map((b) => {
    const r = b as Record<string, unknown>;
    return {
      BackupId: r.backup_id ?? "",
      BackupDate: r.backup_created_at ?? "",
      Code: r.employee_code ?? "",
      Name: r.employee_name ?? "",
      Department: r.department ?? "",
      TotalScore: r.total_score ?? 0,
      RemainingScore: r.remaining_score ?? 0,
      ImportDate: r.import_date ?? "",
      FileName: r.file_name ?? "",
    };
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rewardHistory), "RewardHistory");

  const settings = [{
    ExportDate: today,
    ExportTime: new Date().toISOString(),
    TotalEmployees: scores.length,
    TotalTransactions: entries.length,
    TotalBackups: scoreBackups.length,
    AppVersion: "1.0",
  }];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(settings), "Settings");

  XLSX.writeFile(wb, `Waste_Backup_${today}.xlsx`);
}

type RestoreData = {
  entries: Record<string, unknown>[];
  scores: Record<string, unknown>[];
  scoreBackups: Record<string, unknown>[];
};

export async function importBackup(file: File): Promise<void> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });

  const readSheet = (name: string): Record<string, unknown>[] => {
    const ws = wb.Sheets[name];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: "" });
  };

  const empRows = readSheet("Employees");
  const txRows = readSheet("DailyTransactions");
  const rewardRows = readSheet("RewardHistory");

  const scores = empRows.map((r) => ({
    employee_code: String(r.Code ?? r.code ?? r.employee_code ?? ""),
    employee_name: String(r.Name ?? r.name ?? r.employee_name ?? ""),
    department: String(r.Department ?? r.department ?? "") || null,
    total_score: Number(r.TotalPoints ?? r.total_score ?? r.total_score ?? 0) || 0,
    remaining_score: Number(r.RemainingPoints ?? r.remaining_score ?? r.remaining_score ?? 0) || 0,
    import_date: String(r.ImportDate ?? r.import_date ?? new Date().toISOString().slice(0, 10)),
    file_name: null,
  }));

  const entries = txRows.map((r) => ({
    date: String(r.Date ?? r.date ?? ""),
    code: String(r.Code ?? r.code ?? "") || null,
    name: String(r.Name ?? r.name ?? ""),
    unit_org: String(r.Department ?? r.unit_org ?? "") || null,
    room: String(r.Room ?? r.room ?? ""),
    waste_type: String(r.WasteType ?? r.waste_type ?? ""),
    quantity: String(r.Quantity ?? r.quantity ?? "0"),
    qty_unit: String(r.Unit ?? r.qty_unit ?? "") || null,
    points: String(r.Points ?? r.points ?? "0"),
    reward: String(r.Reward ?? r.reward ?? "") || null,
    signature: String(r.Signature ?? r.signature ?? "") || null,
  }));

  const scoreBackups = rewardRows.map((r) => ({
    employee_code: String(r.Code ?? r.employee_code ?? ""),
    employee_name: String(r.Name ?? r.employee_name ?? ""),
    department: String(r.Department ?? r.department ?? "") || null,
    total_score: Number(r.TotalScore ?? r.total_score ?? 0) || 0,
    remaining_score: Number(r.RemainingScore ?? r.remaining_score ?? 0) || 0,
    import_date: String(r.ImportDate ?? r.import_date ?? new Date().toISOString().slice(0, 10)),
    file_name: String(r.FileName ?? r.file_name ?? "") || null,
  }));

  await deleteAll("waste_entries");
  await deleteAll("accumulated_scores");
  await deleteAll("accumulated_scores_backup");

  if (entries.length > 0) await insertBatched("waste_entries", entries);
  if (scores.length > 0) await insertBatched("accumulated_scores", scores);
  if (scoreBackups.length > 0) await insertBatched("accumulated_scores_backup", scoreBackups);
}

type LocalBackup = {
  timestamp: string;
  entries: number;
  scores: number;
  data: RestoreData;
};

export function saveLocalBackup(entries: Record<string, unknown>[], scores: Record<string, unknown>[]): void {
  try {
    const backups: LocalBackup[] = JSON.parse(localStorage.getItem(BACKUP_KEY) || "[]");
    backups.unshift({
      timestamp: new Date().toISOString(),
      entries: entries.length,
      scores: scores.length,
      data: { entries, scores, scoreBackups: [] },
    });
    if (backups.length > MAX_BACKUPS) backups.length = MAX_BACKUPS;
    localStorage.setItem(BACKUP_KEY, JSON.stringify(backups));
  } catch {
    // localStorage may be full with large data — skip silently
  }
}

export function getLocalBackups(): LocalBackup[] {
  try {
    return JSON.parse(localStorage.getItem(BACKUP_KEY) || "[]");
  } catch {
    return [];
  }
}

export async function restoreLocalBackup(index: number): Promise<void> {
  const backups = getLocalBackups();
  const backup = backups[index];
  if (!backup) throw new Error("Backup not found");

  await deleteAll("waste_entries");
  await deleteAll("accumulated_scores");

  if (backup.data.entries.length > 0) await insertBatched("waste_entries", backup.data.entries);
  if (backup.data.scores.length > 0) await insertBatched("accumulated_scores", backup.data.scores);
}
