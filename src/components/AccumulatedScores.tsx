import { Fragment, useEffect, useMemo, useRef, useState } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card } from "@/components/ui/card";
import {
  Upload,
  Download,
  Search,
  FileSpreadsheet,
  AlertTriangle,
  CheckCircle2,
  Database,
  History,
  X,
  Loader2,
  Trash2,
  Gift,
} from "lucide-react";
import { toast } from "sonner";
import * as XLSX from "xlsx";
import { supabase } from "@/integrations/supabase/client";
import { RewardExchangePanel } from "@/components/RewardExchange";
import type { AppUser } from "@/lib/auth";

export type Lang = "th" | "lo" | "en";
type L = { th: string; lo: string; en: string };
const tr = (l: L, lang: Lang) => l[lang];

const LOC: Record<string, L> = {
  exchangeBtn: { th: "แลก", lo: "ແລກ", en: "Redeem" },
  closeBtn: { th: "ปิด", lo: "ປິດ", en: "Close" },
  importTitle: { th: "นำเข้าคะแนนสะสมจาก Excel", lo: "ນຳເຂົ້າຄະແນນສະສົມຈາກ Excel", en: "Import Accumulated Scores from Excel" },
  importSubtitle: { th: "นำเข้าข้อมูลคะแนนสะสมจากไฟล์ Excel เพื่อใช้ค้นหา", lo: "ນຳເຂົ້າຂໍ້ມູນຄະແນນສະສົມຈາກ Excel ເພື່ອຄົ້ນຫາ", en: "Import accumulated score data from Excel for searching" },
  warningMsg: { th: "การนำเข้าคะแนนสะสมจะไม่กระทบข้อมูลการบันทึกรายวันในระบบ", lo: "ການນຳເຂົ້າຄະແນນສະສົມຈະບໍ່ກະທົນຂໍ້ມູນບັນທຶກປະຈຳວັນໃນລະບົບ", en: "Importing accumulated scores will not affect daily records in the system" },
  uploadBtn: { th: "อัปโหลด Excel (.xlsx)", lo: "ອັບໂຫລດ Excel (.xlsx)", en: "Upload Excel (.xlsx)" },
  templateBtn: { th: "ดาวน์โหลด Template", lo: "ດາວໂຫລດ Template", en: "Download Template" },
  exportBtn: { th: "Export คะแนนสะสมปัจจุบัน", lo: "ສົ່ງອອກຄະແນນສະສົມປັດຈຸບັນ", en: "Export Current Scores" },
  preview: { th: "พรีวิวข้อมูล", lo: "ພຣີວິວຂໍ້ມູນ", en: "Data Preview" },
  confirmImport: { th: "ยืนยันการนำเข้า", lo: "ຢືນຢັນການນຳເຂົ້າ", en: "Confirm Import" },
  cancel: { th: "ยกเลิก", lo: "ຍົກເລີກ", en: "Cancel" },
  importing: { th: "กำลังนำเข้า...", lo: "ກຳລັງນຳເຂົ້າ...", en: "Importing..." },
  backupDone: { th: "สำรองข้อมูลเก่าเรียบร้อย", lo: "ສຳຮອງຂໍ້ມູນເກົ່າຮຽບຮ້ອຍ", en: "Backup completed" },
  importSuccess: { th: "นำเข้าข้อมูลสำเร็จ", lo: "ນຳເຂົ້າຂໍ້ມູນສຳເລັດ", en: "Import successful" },
  noFile: { th: "ยังไม่ได้เลือกไฟล์", lo: "ຍັງບໍ່ໄດ້ເລືອກໄຟລ໌", en: "No file selected" },
  duplicates: { th: "พบข้อมูลซ้ำ", lo: "ພົບຂໍ້ມູນຊ້ຳ", en: "Duplicates found" },
  validRows: { th: "ข้อมูลถูกต้อง", lo: "ຂໍ້ມູນຖືກຕ້ອງ", en: "Valid rows" },
  invalidRows: { th: "ข้อมูลไม่ถูกต้อง", lo: "ຂໍ້ມູນບໍ່ຖືກຕ້ອງ", en: "Invalid rows" },
  totalRows: { th: "ทั้งหมด", lo: "ລວມທັງໝົດ", en: "Total" },
  rows: { th: "รายการ", lo: "ລາຍການ", en: "rows" },
  empCode: { th: "รหัสพนักงาน", lo: "ລະຫັດພະນັກງານ", en: "Employee Code" },
  empName: { th: "ชื่อ-นามสกุล", lo: "ຊື່-ນາມສະກຸນ", en: "Full Name" },
  department: { th: "หน่วยงาน", lo: "ໜ່ວຍງານ", en: "Department" },
  totalScore: { th: "คะแนนสะสม", lo: "ຄະແນນສະສົມ", en: "Total Score" },
  importDate: { th: "วันที่นำเข้า", lo: "ວັນທີນຳເຂົ້າ", en: "Import Date" },
  searchTitle: { th: "ค้นหาคะแนนสะสม", lo: "ຄົ້ນຫາຄະແນນສະສົມ", en: "Search Accumulated Scores" },
  searchPlaceholder: { th: "ค้นหาด้วย รหัสพนักงาน / ชื่อ / หน่วยงาน", lo: "ຄົ້ນຫາດ້ວຍ ລະຫັດ / ຊື່ / ໜ່ວຍງານ", en: "Search by Code / Name / Department" },
  noResults: { th: "ไม่พบข้อมูล", lo: "ບໍ່ພົບຂໍ້ມູນ", en: "No results found" },
  noData: { th: "ยังไม่มีข้อมูลคะแนนสะสม กรุณานำเข้าจาก Excel ก่อน", lo: "ຍັງບໍ່ມີຂໍ້ມູນຄະແນນສະສົມ ກະລຸນານຳເຂົ້າຈາກ Excel ກ່ອນ", en: "No accumulated score data yet. Please import from Excel first." },
  resultsCount: { th: "พบ {n} รายการ", lo: "ພົບ {n} ລາຍການ", en: "{n} results found" },
  clearSearch: { th: "ล้าง", lo: "ລ້າງ", en: "Clear" },
  backupHistory: { th: "ประวัติการสำรองข้อมูล", lo: "ປະຫວັດການສຳຮອງຂໍ້ມູນ", en: "Backup History" },
  noBackups: { th: "ยังไม่มีประวัติการสำรองข้อมูล", lo: "ຍັງບໍ່ມີປະຫວັດການສຳຮອງ", en: "No backup history yet" },
  currentRecords: { th: "คะแนนสะสมในระบบ", lo: "ຄະແນນສະສົມໃນລະບົບ", en: "Accumulated Scores in System" },
  fileCol: { th: "ไฟล์ต้นทาง", lo: "ໄຟລ໌ຕົ້ນທາງ", en: "Source File" },
  updated: { th: "อัปเดตล่าสุด", lo: "ອັບເດດຫຼ້າສຸດ", en: "Last Updated" },
  errorReadFile: { th: "ไม่สามารถอ่านไฟล์ Excel ได้", lo: "ບໍ່ສາມາດອ່ານໄຟລ໌ Excel ໄດ້", en: "Cannot read Excel file" },
  errorNoData: { th: "ไฟล์ไม่มีข้อมูล หรือคอลัมน์ไม่ถูกต้อง", lo: "ໄຟລ໌ບໍ່ມີຂໍ້ມູນ ຫຼື ຖັນບໍ່ຖືກຕ້ອງ", en: "File has no data or incorrect columns" },
  errorMissingName: { th: "ชื่อ-นามสกุล ไม่สามารถเว้นว่างได้", lo: "ຊື່-ນາມສະກຸນບໍ່ສາມາດຫວ່າງໄດ້", en: "Name cannot be empty" },
  errorMissingCode: { th: "รหัสพนักงาน ไม่สามารถเว้นว่างได้", lo: "ລະຫັດພະນັກງານບໍ່ສາມາດຫວ່າງໄດ້", en: "Employee code cannot be empty" },
  errorInvalidScore: { th: "คะแนนต้องเป็นตัวเลข", lo: "ຄະແນນຕ້ອງເປັນຕົວເລກ", en: "Score must be a number" },
  largeFile: { th: "กำลังประมวลผลไฟล์ขนาดใหญ่...", lo: "ກຳລັງປະມວນຜົນໄຟລ໌ຂະໜາດໃຫຍ່...", en: "Processing large file..." },
  recordsCount: { th: "ระเบียนในระบบ", lo: "ບັນທຶກໃນລະບົບ", en: "records in system" },
  loading: { th: "กำลังโหลด...", lo: "ກຳລັງໂຫລດ...", en: "Loading..." },
  status: { th: "สถานะ", lo: "ສະຖານະ", en: "Status" },
  remainingScore: { th: "คะแนนคงเหลือ", lo: "ຄະແນນຄົງເຫຼືອ", en: "Remaining Score" },
  clearDataBtn: { th: "ล้างข้อมูลทั้งหมด", lo: "ລ້າງຂໍ້ມູນທັງໝົດ", en: "Clear All Data" },
  clearDataConfirm: { th: "คุณแน่ใจหรือไม่ว่าต้องการล้างข้อมูลคะแนนสะสมทั้งหมด? ข้อมูลเดิมจะถูกสำรองไว้ก่อน", lo: "ທ່ານແນ່ໃ�ຫຼືບໍ່ວ່າຕ້ອງການລ້າງຂໍ້ມູນຄະແນນສະສົມທັງໝົດ? ຂໍ້ມູນເກົ່າຈະຖືກສຳຮອງໄວ້ກ່ອນ", en: "Are you sure you want to clear all accumulated score data? Old data will be backed up first" },
  clearSuccess: { th: "ล้างข้อมูลสำเร็จ", lo: "ລ້າງຂໍ້ມູນສຳເລັດ", en: "Data cleared successfully" },
  clearing: { th: "กำลังล้างข้อมูล...", lo: "ກຳລັງລ້າງຂໍ້ມູນ...", en: "Clearing data..." },
};

const lt = (k: keyof typeof LOC, lang: Lang) => tr(LOC[k], lang);

const PAGE_SIZE = 1000;

async function fetchAll<T>(
  table: string,
  orderCol: string,
  ascending = true,
): Promise<T[]> {
  const all: T[] = [];
  let from = 0;
  let more = true;
  while (more) {
    const { data, error } = await supabase
      .from(table)
      .select("*")
      .order(orderCol, { ascending })
      .range(from, from + PAGE_SIZE - 1);
    if (error) throw error;
    const rows = (data ?? []) as T[];
    all.push(...rows);
    more = rows.length === PAGE_SIZE;
    from += PAGE_SIZE;
  }
  return all;
}

async function insertBatched<T>(
  table: string,
  rows: T[],
  batchSize = 500,
): Promise<void> {
  for (let i = 0; i < rows.length; i += batchSize) {
    const batch = rows.slice(i, i + batchSize);
    const { error } = await supabase.from(table).insert(batch);
    if (error) throw error;
  }
}

type ScoreRow = {
  id: string;
  employee_code: string;
  employee_name: string;
  department: string | null;
  total_score: number;
  remaining_score: number;
  import_date: string;
  file_name: string | null;
  updated_at: string;
};

type ParsedRow = {
  employee_code: string;
  employee_name: string;
  department: string;
  total_score: number;
  remaining_score: number;
  valid: boolean;
  error?: string;
  duplicate?: boolean;
};

type BackupRow = {
  backup_id: string;
  backup_created_at: string;
  employee_code: string;
  employee_name: string;
  department: string | null;
  total_score: number;
  remaining_score: number;
  file_name: string | null;
};

// ===== Import Scores Component =====
export function ImportScores({ lang, onImported }: { lang: Lang; onImported?: () => void }) {
  const [parsed, setParsed] = useState<ParsedRow[]>([]);
  const [fileName, setFileName] = useState("");
  const [importing, setImporting] = useState(false);
  const [parsing, setParsing] = useState(false);
  const [existingCount, setExistingCount] = useState(0);
  const [backups, setBackups] = useState<BackupRow[]>([]);
  const [clearing, setClearing] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    (async () => {
      const { count } = await supabase
        .from("accumulated_scores")
        .select("*", { count: "exact", head: true });
      setExistingCount(count ?? 0);

      const { data: bk } = await supabase
        .from("accumulated_scores_backup")
        .select("*")
        .order("backup_created_at", { ascending: false })
        .limit(20);
      setBackups((bk as BackupRow[]) ?? []);
    })();
  }, []);

  const handleFile = async (file: File) => {
    setParsing(true);
    setParsed([]);
    setFileName(file.name);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      const seenCodes = new Set<string>();
      const allRows: ParsedRow[] = [];

      for (const sheetName of wb.SheetNames) {
        const ws = wb.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: "" });
        if (json.length === 0) continue;

        // Find header row: first row (within first 5) containing "ชื่อ" or "name"
        let headerIdx = -1;
        for (let i = 0; i < Math.min(5, json.length); i++) {
          if (json[i].some((c) => String(c).includes("ชื่อ") || String(c).toLowerCase().includes("name"))) {
            headerIdx = i;
            break;
          }
        }
        if (headerIdx === -1) continue;

        const headers = json[headerIdx].map((h) => String(h).trim());

        const findColIdx = (candidates: string[]): number => {
          for (const cand of candidates) {
            if (!cand) continue;
            const cl = cand.toLowerCase();
            const idx = headers.findIndex((h) => {
              if (!h) return false;
              const hl = h.toLowerCase();
              return hl === cl || hl.includes(cl) || cl.includes(hl);
            });
            if (idx !== -1) return idx;
          }
          return -1;
        };

        const codeIdx = findColIdx(["รหัสพนักงาน", "รหัส", "employee_code", "code", "Employee Code", "ID"]);
        const nameIdx = findColIdx(["ชื่อ-นามสกุล", "ชื่อ", "employee_name", "name", "Full Name", "Name"]);
        const deptIdx = findColIdx(["หน่วยงาน", "department", "org", "Department", "Organization"]);
        const scoreIdx = findColIdx(["คะแนนสะสมที่ได้ทั้งหมด", "คะแนนสะสม", "total_score", "score", "points", "Total Score", "Accumulated"]);
        const remainingIdx = findColIdx(["คะแนนที่เหลือจากการหัก", "คะแนนที่เหลือ", "คะแนนคงเหลือ", "remaining_score", "remaining", "Remaining"]);

        if (nameIdx === -1) continue;

        let autoCodeCounter = 0;

        for (let i = headerIdx + 1; i < json.length; i++) {
          const row = json[i];
          const name = String(row[nameIdx] ?? "").trim();
          if (!name) continue;

          let code = codeIdx >= 0 ? String(row[codeIdx] ?? "").trim() : "";
          if (!code) {
            autoCodeCounter++;
            code = `${sheetName}-${autoCodeCounter}`;
          }

          const dept = deptIdx >= 0 ? String(row[deptIdx] ?? "").trim() : "";
          const scoreRaw = scoreIdx >= 0 ? String(row[scoreIdx] ?? "0").trim() : "0";
          const score = parseFloat(scoreRaw) || 0;
          const remainingRaw = remainingIdx >= 0 ? String(row[remainingIdx] ?? "0").trim() : "0";
          const remaining = parseFloat(remainingRaw) || 0;

          let valid = true;
          let error: string | undefined;
          let duplicate = false;

          if (seenCodes.has(code)) {
            duplicate = true;
          } else {
            seenCodes.add(code);
          }

          allRows.push({ employee_code: code, employee_name: name, department: dept, total_score: score, remaining_score: remaining, valid, error, duplicate });
        }
      }

      if (allRows.length === 0) {
        toast.error(lt("errorNoData", lang));
        setParsing(false);
        return;
      }

      setParsed(allRows);
      const dupCount = allRows.filter((r) => r.duplicate).length;
      if (dupCount > 0) {
        toast.warning(`${lt("duplicates", lang)}: ${dupCount} ${lt("rows", lang)}`);
      }
      toast.success(`${lt("validRows", lang)}: ${allRows.filter((r) => r.valid && !r.duplicate).length} / ${allRows.length} ${lt("rows", lang)}`);
    } catch {
      toast.error(lt("errorReadFile", lang));
    }
    setParsing(false);
  };

  const doImport = async () => {
    const validRows = parsed.filter((r) => r.valid && !r.duplicate);
    if (validRows.length === 0) {
      toast.error(lt("noFile", lang));
      return;
    }

    setImporting(true);
    const today = new Date().toISOString().slice(0, 10);

    // Step 1: Backup existing data
    try {
      const existing = await fetchAll<ScoreRow>("accumulated_scores", "employee_code");
      if (existing.length > 0) {
        const backupPayload = existing.map((r) => ({
          id: r.id,
          employee_code: r.employee_code,
          employee_name: r.employee_name,
          department: r.department,
          total_score: r.total_score,
          remaining_score: r.remaining_score,
          import_date: r.import_date,
          file_name: r.file_name,
          updated_at: r.updated_at,
        }));
        await insertBatched("accumulated_scores_backup", backupPayload);
        toast.success(lt("backupDone", lang));
      }
    } catch (e) {
      toast.error("Backup failed: " + (e as Error).message);
      setImporting(false);
      return;
    }

    // Step 2: Delete old data
    const { error: delErr } = await supabase.from("accumulated_scores").delete().neq("id", "00000000-0000-0000-0000-000000000000");
    if (delErr) {
      toast.error("Clear old data failed: " + delErr.message);
      setImporting(false);
      return;
    }

    // Step 3: Insert new data in batches
    const insertPayload = validRows.map((r) => ({
      employee_code: r.employee_code,
      employee_name: r.employee_name,
      department: r.department || null,
      total_score: r.total_score,
      remaining_score: r.remaining_score,
      import_date: today,
      file_name: fileName,
    }));
    try {
      await insertBatched("accumulated_scores", insertPayload);
    } catch (e) {
      toast.error("Import failed: " + (e as Error).message);
      setImporting(false);
      return;
    }

    setImporting(false);
    setParsed([]);
    setFileName("");
    setExistingCount(validRows.length);
    toast.success(`${lt("importSuccess", lang)}: ${validRows.length} ${lt("rows", lang)}`);

    // Refresh backups
    const { data: bk } = await supabase
      .from("accumulated_scores_backup")
      .select("*")
      .order("backup_created_at", { ascending: false })
      .limit(20);
    setBackups((bk as BackupRow[]) ?? []);

    // Notify parent to reload employee master
    onImported?.();
  };

  const downloadTemplate = () => {
    const data = [
      { "รหัสพนักงาน": "EMP001", "ชื่อ-นามสกุล": "สมชาย ใจดี", "หน่วยงาน": "ฝ่ายการเงิน", "คะแนนสะสม": 150 },
      { "รหัสพนักงาน": "EMP002", "ชื่อ-นามสกุล": "สมหญิง รักดี", "หน่วยงาน": "ฝ่ายบุคคล", "คะแนนสะสม": 85 },
      { "รหัสพนักงาน": "EMP003", "ชื่อ-นามสกุล": "วิชัย ตั้งใจ", "หน่วยงาน": "ฝ่ายผลิต", "คะแนนสะสม": 230 },
    ];
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{ wch: 15 }, { wch: 25 }, { wch: 20 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, "template-accumulated-scores.xlsx");
    toast.success("Template downloaded");
  };

  const exportCurrent = async () => {
    let data: ScoreRow[];
    try {
      data = await fetchAll<ScoreRow>("accumulated_scores", "employee_code");
    } catch (e) {
      toast.error("Export failed: " + (e as Error).message);
      return;
    }
    if (data.length === 0) {
      toast.error(lt("noData", lang));
      return;
    }
    const rows = data.map((r) => ({
      "รหัสพนักงาน": r.employee_code,
      "ชื่อ-นามสกุล": r.employee_name,
      "หน่วยงาน": r.department ?? "",
      "คะแนนสะสม": r.total_score,
      "คะแนนคงเหลือ": r.remaining_score,
      "วันที่นำเข้า": r.import_date,
      "ไฟล์ต้นทาง": r.file_name ?? "",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Accumulated Scores");
    XLSX.writeFile(wb, `accumulated-scores-export-${new Date().toISOString().slice(0, 10)}.xlsx`);
    toast.success("Export completed");
  };

  const validCount = parsed.filter((r) => r.valid && !r.duplicate).length;
  const invalidCount = parsed.filter((r) => !r.valid).length;
  const dupCount = parsed.filter((r) => r.duplicate).length;

  const clearAllData = async () => {
    if (!window.confirm(lt("clearDataConfirm", lang))) return;
    setClearing(true);
    try {
      const existing = await fetchAll<ScoreRow>("accumulated_scores", "employee_code");
      if (existing.length > 0) {
        const backupPayload = existing.map((r) => ({
          id: r.id,
          employee_code: r.employee_code,
          employee_name: r.employee_name,
          department: r.department,
          total_score: r.total_score,
          remaining_score: r.remaining_score,
          import_date: r.import_date,
          file_name: r.file_name,
          updated_at: r.updated_at,
        }));
        await insertBatched("accumulated_scores_backup", backupPayload);
      }
      const { error: delErr } = await supabase.from("accumulated_scores").delete().neq("id", "00000000-0000-0000-0000-000000000000");
      if (delErr) {
        toast.error("Clear failed: " + delErr.message);
        setClearing(false);
        return;
      }
      setExistingCount(0);
      setParsed([]);
      setFileName("");
      toast.success(lt("clearSuccess", lang));
      onImported?.();
      const { data: bk } = await supabase
        .from("accumulated_scores_backup")
        .select("*")
        .order("backup_created_at", { ascending: false })
        .limit(20);
      setBackups((bk as BackupRow[]) ?? []);
    } catch {
      toast.error(lt("errorReadFile", lang));
    }
    setClearing(false);
  };

  return (
    <div className="w-full min-w-0 space-y-4 overflow-x-hidden sm:space-y-6">
      {/* Warning banner */}
      <div className="flex items-start gap-3 rounded-lg border border-amber-300 bg-amber-50 p-4">
        <AlertTriangle className="mt-0.5 h-5 w-5 flex-shrink-0 text-amber-600" />
        <p className="text-sm text-amber-800">{lt("warningMsg", lang)}</p>
      </div>

      {/* Upload card */}
      <Card className="p-3 sm:p-4 lg:p-6">
        <div className="mb-3 flex items-center gap-2 sm:mb-4">
          <FileSpreadsheet className="h-5 w-5 text-primary" />
          <h2 className="text-base font-semibold sm:text-lg">{lt("importTitle", lang)}</h2>
        </div>
        <p className="mb-4 text-sm text-muted-foreground">{lt("importSubtitle", lang)}</p>

        <div className="flex flex-wrap gap-2">
          <input
            ref={fileRef}
            type="file"
            accept=".xlsx,.xls"
            className="hidden"
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) handleFile(f);
              e.target.value = "";
            }}
          />
          <Button onClick={() => fileRef.current?.click()} disabled={parsing || importing}>
            {parsing ? (
              <>
                <Loader2 className="mr-1 h-4 w-4 animate-spin" /> {lt("largeFile", lang)}
              </>
            ) : (
              <>
                <Upload className="mr-1 h-4 w-4" /> {lt("uploadBtn", lang)}
              </>
            )}
          </Button>
          <Button variant="outline" onClick={downloadTemplate}>
            <Download className="mr-1 h-4 w-4" /> {lt("templateBtn", lang)}
          </Button>
          <Button variant="outline" onClick={exportCurrent}>
            <Download className="mr-1 h-4 w-4" /> {lt("exportBtn", lang)}
          </Button>
          {existingCount > 0 && (
            <Button variant="destructive" onClick={clearAllData} disabled={clearing || importing || parsing}>
              {clearing ? (
                <>
                  <Loader2 className="mr-1 h-4 w-4 animate-spin" /> {lt("clearing", lang)}
                </>
              ) : (
                <>
                  <Trash2 className="mr-1 h-4 w-4" /> {lt("clearDataBtn", lang)}
                </>
              )}
            </Button>
          )}
        </div>

        {existingCount > 0 && (
          <div className="mt-3 flex items-center gap-2 text-sm text-muted-foreground">
            <Database className="h-4 w-4" />
            {existingCount} {lt("recordsCount", lang)}
          </div>
        )}

        {fileName && (
          <p className="mt-2 text-xs text-muted-foreground">
            {lt("fileCol", lang)}: {fileName}
          </p>
        )}
      </Card>

      {/* Preview */}
      {parsed.length > 0 && (
        <Card className="p-3 sm:p-4 lg:p-6">
          <div className="mb-4 flex flex-wrap items-center justify-between gap-2">
            <h2 className="text-lg font-semibold">{lt("preview", lang)}</h2>
            <div className="flex flex-wrap gap-2 text-xs">
              <span className="inline-flex items-center gap-1 rounded-full bg-green-100 px-2 py-1 text-green-700">
                <CheckCircle2 className="h-3 w-3" /> {lt("validRows", lang)}: {validCount}
              </span>
              {invalidCount > 0 && (
                <span className="inline-flex items-center gap-1 rounded-full bg-red-100 px-2 py-1 text-red-700">
                  <X className="h-3 w-3" /> {lt("invalidRows", lang)}: {invalidCount}
                </span>
              )}
              {dupCount > 0 && (
                <span className="inline-flex items-center gap-1 rounded-full bg-amber-100 px-2 py-1 text-amber-700">
                  <AlertTriangle className="h-3 w-3" /> {lt("duplicates", lang)}: {dupCount}
                </span>
              )}
              <span className="text-muted-foreground">
                {lt("totalRows", lang)}: {parsed.length}
              </span>
            </div>
          </div>

          <div className="max-h-96 overflow-y-auto overflow-x-hidden rounded-lg border">
            <table className="w-full border-collapse text-sm">
              <thead className="sticky top-0 bg-muted">
                <tr>
                  <th className="border px-2 py-2 text-left">{lt("empCode", lang)}</th>
                  <th className="border px-2 py-2 text-left">{lt("empName", lang)}</th>
                  <th className="border px-2 py-2 text-left">{lt("department", lang)}</th>
                  <th className="border px-2 py-2 text-right">{lt("totalScore", lang)}</th>
                  <th className="border px-2 py-2 text-right">{lt("remainingScore", lang)}</th>
                  <th className="border px-2 py-2 text-center">{lt("status", lang)}</th>
                </tr>
              </thead>
              <tbody>
                {parsed.slice(0, 200).map((r, i) => (
                  <tr
                    key={i}
                    className={`border-b ${
                      !r.valid
                        ? "bg-red-50"
                        : r.duplicate
                          ? "bg-amber-50"
                          : "hover:bg-muted/30"
                    }`}
                  >
                    <td className="border px-2 py-1.5">{r.employee_code}</td>
                    <td className="border px-2 py-1.5">{r.employee_name}</td>
                    <td className="border px-2 py-1.5 text-muted-foreground">{r.department || "-"}</td>
                    <td className="border px-2 py-1.5 text-right font-semibold">{r.total_score}</td>
                    <td className="border px-2 py-1.5 text-right font-semibold text-muted-foreground">{r.remaining_score}</td>
                    <td className="border px-2 py-1.5 text-center text-xs">
                      {!r.valid ? (
                        <span className="text-red-600">{r.error}</span>
                      ) : r.duplicate ? (
                        <span className="text-amber-600">{lt("duplicates", lang)}</span>
                      ) : (
                        <CheckCircle2 className="mx-auto h-4 w-4 text-green-600" />
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {parsed.length > 200 && (
            <p className="mt-2 text-xs text-muted-foreground">
              {lt("preview", lang)}: 200 / {parsed.length} {lt("rows", lang)}
            </p>
          )}

          <div className="mt-4 flex gap-2">
            <Button onClick={doImport} disabled={importing || validCount === 0}>
              {importing ? (
                <>
                  <Loader2 className="mr-1 h-4 w-4 animate-spin" /> {lt("importing", lang)}
                </>
              ) : (
                <>
                  <CheckCircle2 className="mr-1 h-4 w-4" /> {lt("confirmImport", lang)} ({validCount})
                </>
              )}
            </Button>
            <Button variant="outline" onClick={() => { setParsed([]); setFileName(""); }}>
              <X className="mr-1 h-4 w-4" /> {lt("cancel", lang)}
            </Button>
          </div>
        </Card>
      )}

      {/* Backup History */}
      {backups.length > 0 && (
        <Card className="p-3 sm:p-4 lg:p-6">
          <div className="mb-3 flex items-center gap-2 sm:mb-4">
            <History className="h-5 w-5 text-primary" />
            <h2 className="text-base font-semibold sm:text-lg">{lt("backupHistory", lang)}</h2>
          </div>
          <div className="max-h-48 overflow-y-auto overflow-x-hidden rounded-lg border">
            <table className="w-full border-collapse text-sm">
              <thead className="sticky top-0 bg-muted">
                <tr>
                  <th className="border px-2 py-2 text-left">{lt("empCode", lang)}</th>
                  <th className="border px-2 py-2 text-left">{lt("empName", lang)}</th>
                  <th className="border px-2 py-2 text-right">{lt("totalScore", lang)}</th>
                  <th className="border px-2 py-2 text-right">{lt("remainingScore", lang)}</th>
                  <th className="border px-2 py-2 text-left">{lt("fileCol", lang)}</th>
                  <th className="border px-2 py-2 text-left">{lt("updated", lang)}</th>
                </tr>
              </thead>
              <tbody>
                {backups.map((b) => (
                  <tr key={b.backup_id} className="border-b hover:bg-muted/30">
                    <td className="border px-2 py-1.5">{b.employee_code}</td>
                    <td className="border px-2 py-1.5">{b.employee_name}</td>
                    <td className="border px-2 py-1.5 text-right">{b.total_score}</td>
                    <td className="border px-2 py-1.5 text-right">{b.remaining_score}</td>
                    <td className="border px-2 py-1.5 text-xs text-muted-foreground">{b.file_name ?? "-"}</td>
                    <td className="border px-2 py-1.5 text-xs text-muted-foreground">
                      {new Date(b.backup_created_at).toLocaleString("th-TH")}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}
    </div>
  );
}

// ===== Search Scores Component =====
export function SearchScores({ lang, currentUser }: { lang: Lang; currentUser: AppUser }) {
  const [scores, setScores] = useState<ScoreRow[]>([]);
  const [loaded, setLoaded] = useState(false);
  const [searchText, setSearchText] = useState("");
  const [debounced, setDebounced] = useState("");
  const [exchangeId, setExchangeId] = useState<string | null>(null);

  useEffect(() => {
    const t = setTimeout(() => setDebounced(searchText), 300);
    return () => clearTimeout(t);
  }, [searchText]);

  const loadScores = async () => {
    try {
      const data = await fetchAll<ScoreRow>("accumulated_scores", "total_score", false);
      setScores(data);
    } catch (e) {
      toast.error("Load failed: " + (e as Error).message);
    }
    setLoaded(true);
  };

  useEffect(() => {
    loadScores();
  }, []);

  const filtered = useMemo(() => {
    if (!debounced.trim()) return scores;
    const kw = debounced.toLowerCase();
    return scores.filter(
      (s) =>
        s.employee_code.toLowerCase().includes(kw) ||
        s.employee_name.toLowerCase().includes(kw) ||
        (s.department ?? "").toLowerCase().includes(kw),
    );
  }, [scores, debounced]);

  function Highlight({ text, keyword }: { text: string; keyword: string }) {
    if (!keyword.trim()) return <>{text}</>;
    const kw = keyword.toLowerCase();
    const parts: React.ReactNode[] = [];
    let remaining = text;
    let i = 0;
    while (remaining.length > 0) {
      const idx = remaining.toLowerCase().indexOf(kw);
      if (idx === -1) {
        parts.push(<span key={i}>{remaining}</span>);
        break;
      }
      if (idx > 0) parts.push(<span key={i}>{remaining.slice(0, idx)}</span>);
      parts.push(
        <mark key={`h${i}`} className="rounded bg-primary/20 px-0.5 font-semibold text-primary">
          {remaining.slice(idx, idx + kw.length)}
        </mark>,
      );
      remaining = remaining.slice(idx + kw.length);
      i++;
    }
    return <>{parts}</>;
  }

  return (
    <div className="w-full min-w-0 space-y-4 overflow-x-hidden sm:space-y-6">
      <Card className="p-3 sm:p-4 lg:p-6">
        <div className="mb-3 flex items-center gap-2 sm:mb-4">
          <Search className="h-5 w-5 text-primary" />
          <h2 className="text-base font-semibold sm:text-lg">{lt("searchTitle", lang)}</h2>
        </div>

        <div className="mb-4 flex flex-wrap items-center gap-2">
          <div className="relative min-w-0 flex-1">
            <Search className="absolute left-2.5 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
            <Input
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") setDebounced(searchText); }}
              placeholder={lt("searchPlaceholder", lang)}
              className="pl-9 pr-8"
            />
            {searchText && (
              <button
                onClick={() => setSearchText("")}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground"
              >
                <X className="h-4 w-4" />
              </button>
            )}
          </div>
          <Button size="sm" onClick={() => setDebounced(searchText)}>
            <Search className="h-4 w-4" />
          </Button>
          {debounced.trim() && (
            <Button variant="ghost" size="sm" onClick={() => { setSearchText(""); setDebounced(""); }}>
              {lt("clearSearch", lang)}
            </Button>
          )}
        </div>

        {debounced.trim() && (
          <p className="mb-2 text-xs text-muted-foreground">
            {lt("resultsCount", lang).replace("{n}", String(filtered.length))}
          </p>
        )}

        {!loaded ? (
          <div className="flex items-center justify-center py-8 text-muted-foreground">
            <Loader2 className="mr-2 h-5 w-5 animate-spin" /> {lt("loading", lang)}
          </div>
        ) : scores.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-12 text-center">
            <Database className="mb-3 h-10 w-10 text-muted-foreground/50" />
            <p className="text-sm text-muted-foreground">{lt("noData", lang)}</p>
          </div>
        ) : filtered.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-12 text-center">
            <Search className="mb-3 h-10 w-10 text-muted-foreground/50" />
            <p className="text-sm text-muted-foreground">{lt("noResults", lang)}</p>
          </div>
        ) : (
          <div className="search-results-scroll overflow-x-auto rounded-lg border">
            <table className="w-full min-w-[480px] border-collapse text-sm">
              <thead className="bg-muted">
                <tr>
                  <th className="border px-2 py-2 text-left">{lt("empCode", lang)}</th>
                  <th className="border px-2 py-2 text-left">{lt("empName", lang)}</th>
                  <th className="border px-2 py-2 text-left">{lt("department", lang)}</th>
                  <th className="border px-2 py-2 text-right">{lt("totalScore", lang)}</th>
                  <th className="border px-2 py-2 text-right">{lt("remainingScore", lang)}</th>
                  <th className="border px-2 py-2 text-left">{lt("importDate", lang)}</th>
                  <th className="border px-2 py-2 text-center">{lt("exchangeBtn", lang)}</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((s) => (
                  <Fragment key={s.id}>
                  <tr className="border-b hover:bg-muted/30">
                    <td className="border px-2 py-2">
                      <Highlight text={s.employee_code} keyword={debounced} />
                    </td>
                    <td className="border px-2 py-2 font-medium">
                      <Highlight text={s.employee_name} keyword={debounced} />
                    </td>
                    <td className="border px-2 py-2 text-muted-foreground">
                      <Highlight text={s.department ?? "-"} keyword={debounced} />
                    </td>
                    <td className="border px-2 py-2 text-right font-bold text-primary">
                      {s.total_score}
                    </td>
                    <td className="border px-2 py-2 text-right font-semibold text-muted-foreground">
                      {s.remaining_score}
                    </td>
                    <td className="border px-2 py-2 text-xs text-muted-foreground">
                      {s.import_date}
                    </td>
                    <td className="border px-2 py-2 text-center">
                      <Button
                        size="sm"
                        variant={exchangeId === s.id ? "default" : "outline"}
                        onClick={() => setExchangeId(exchangeId === s.id ? null : s.id)}
                        className="h-7"
                      >
                        <Gift className="mr-1 h-3.5 w-3.5" />
                        {exchangeId === s.id ? lt("closeBtn", lang) : lt("exchangeBtn", lang)}
                      </Button>
                    </td>
                  </tr>
                  {exchangeId === s.id && (
                    <tr>
                      <td colSpan={7} className="border bg-accent/5 p-3">
                        <RewardExchangePanel
                          lang={lang}
                          currentUser={currentUser}
                          personName={s.employee_name}
                          personCode={s.employee_code}
                          personId={s.id}
                          onExchanged={loadScores}
                        />
                      </td>
                    </tr>
                  )}
                  </Fragment>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {scores.length > 0 && (
          <p className="mt-3 text-xs text-muted-foreground">
            {lt("currentRecords", lang)}: {scores.length} {lt("recordsCount", lang)}
          </p>
        )}
      </Card>
    </div>
  );
}
