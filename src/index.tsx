import { Fragment, useEffect, useMemo, useRef, useState } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card } from "@/components/ui/card";
import { Trash2, Plus, Download, FileText, Gift, Cloud, Pencil, Check, X, Languages, Search, Recycle, Upload, Database, ChevronDown, HardDriveDownload, ArchiveRestore, History, UserSearch, Shield, Loader2 } from "lucide-react";
import { toast } from "sonner";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { supabase } from "@/integrations/supabase/client";
import { ImportScores, SearchScores } from "@/components/AccumulatedScores";
import type { Lang } from "@/components/AccumulatedScores";
import { exportFullBackup, importBackup, saveLocalBackup, getLocalBackups, restoreLocalBackup } from "@/lib/backup";
import { loginUser, getStoredSession, logoutUser, type AppUser } from "@/lib/auth";
import { AdminUsers } from "@/components/AdminUsers";
import { RewardExchange, RewardExchangePanel, fetchAllExchanges, type ExchangeRow } from "@/components/RewardExchange";

// ===== i18n =====
// Lang is imported from AccumulatedScores

type L = { th: string; lo: string; en: string };
const tr = (l: L, lang: Lang) => l[lang];

const UI: Record<string, L> = {
  appTitle: { th: "โครงการขยะแลกของรางวัล", lo: "ໂຄງການຂີ້ເຫຍື້ອແລກຂອງລາງວັນ", en: "Waste-for-Rewards Program" },
  appSubtitle: { th: "บันทึก คำนวณคะแนนอัตโนมัติ และสรุปตามห้อง", lo: "ບັນທຶກ, ຄຳນວນຄະແນນອັດຕະໂນມັດ ແລະ ສະຫຼຸບຕາມຫ້ອງ", en: "Record, auto-calculate points, summarize by room" },
  syncing: { th: "กำลังซิงค์...", lo: "ກຳລັງຊິງຄ໌...", en: "Syncing..." },
  synced: { th: "ซิงค์กับ Cloud", lo: "ຊິງຄ໌ກັບ Cloud", en: "Cloud synced" },
  loading: { th: "กำลังโหลด...", lo: "ກຳລັງໂຫລດ...", en: "Loading..." },
  excel: { th: "Excel", lo: "Excel", en: "Excel" },
  pdf: { th: "PDF", lo: "PDF", en: "PDF" },
  dailyEntry: { th: "บันทึกข้อมูลรายวัน", lo: "ບັນທຶກຂໍ້ມູນປະຈຳວັນ", en: "Daily Entry" },
  date: { th: "วันที่", lo: "ວັນທີ", en: "Date" },
  code: { th: "รหัส", lo: "ລະຫັດ", en: "Code" },
  name: { th: "ชื่อ-นามสกุล", lo: "ຊື່-ນາມສະກຸນ", en: "Full Name" },
  nameRequired: { th: "ชื่อ-นามสกุล *", lo: "ຊື່-ນາມສະກຸນ *", en: "Full Name *" },
  org: { th: "หน่วยงาน", lo: "ໜ່ວຍງານ", en: "Department" },
  reward: { th: "รางวัล", lo: "ລາງວັນ", en: "Reward" },
  signature: { th: "ลายเซ็น", lo: "ລາຍເຊັນ", en: "Signature" },
  matrixHint: { th: "กรอกจำนวนตามชนิดขยะและห้อง (คะแนนคำนวณอัตโนมัติ)", lo: "ປ້ອນຈຳນວນຕາມປະເພດແລະຫ້ອງ (ຄະແນນຄຳນວນອັດຕະໂນມັດ)", en: "Enter amounts by type & room (points auto-calculated)" },
  wasteType: { th: "ชนิดขยะ", lo: "ປະເພດຂີ້ເຫຍື້ອ", en: "Waste Type" },
  criteria: { th: "เกณฑ์", lo: "ເກນ", en: "Criteria" },
  totalRoom: { th: "ห้องรวม", lo: "ຫ້ອງລວມ", en: "Total" },
  points: { th: "คะแนน", lo: "ຄະແນນ", en: "Points" },
  pointPerOne: { th: "= 1 คะแนน", lo: "= 1 ຄະແນນ", en: "= 1 point" },
  perRoomQty: { th: "รวมต่อห้อง (จำนวน)", lo: "ລວມຕໍ່ຫ້ອງ (ຈຳນວນ)", en: "Per-room total (qty)" },
  perRoomPts: { th: "คะแนนต่อห้อง", lo: "ຄະແນນຕໍ່ຫ້ອງ", en: "Per-room points" },
  grandTotalPts: { th: "รวมทั้งหมด", lo: "ລວມທັງໝົດ", en: "Grand total" },
  pointsWord: { th: "คะแนน", lo: "ຄະແນນ", en: "points" },
  saveEntry: { th: "บันทึกรายการ", lo: "ບັນທຶກລາຍການ", en: "Save Entry" },
  clearTable: { th: "ล้างตาราง", lo: "ລ້າງຕາຕະລາງ", en: "Clear Table" },
  clearForm: { th: "ล้างฟอร์มทั้งหมด", lo: "ລ້າງຟອມທັງໝົດ", en: "Clear All Fields" },
  clearAll: { th: "ล้างข้อมูล", lo: "ລ້າງຂໍ້ມູນ", en: "Clear Data" },
  totalPointsAll: { th: "คะแนนรวมทั้งหมด", lo: "ຄະແນນລວມທັງໝົດ", en: "Total Points" },
  perPersonSummary: { th: "สรุปคะแนนรายบุคคล", lo: "ສະຫຼຸບຄະແນນລາຍບຸກຄົນ", en: "Per-Person Summary" },
  accumulated: { th: "คะแนนสะสม", lo: "ຄະແນນສະສົມ", en: "Accumulated" },
  status: { th: "สถานะ", lo: "ສະຖານະ", en: "Status" },
  canRedeem: { th: "แลกของรางวัลได้", lo: "ແລກຂອງລາງວັນໄດ້", en: "Reward eligible" },
  needMore: { th: "อีก", lo: "ຍັງອີກ", en: "Need" },
  more: { th: "คะแนน", lo: "ຄະແນນ", en: "more pts" },
  dailySummary: { th: "สรุปขยะประจำวัน", lo: "ສະຫຼຸບຂີ້ເຫຍື້ອປະຈຳວັນ", en: "Daily Waste Summary" },
  saveNote: { th: "(รวมอยู่ในปุ่ม Save ด้านบน)", lo: "(ລວມຢູ່ໃນປຸ່ມ Save ດ້ານເທິງ)", en: "(included in Save above)" },
  noData: { th: "ยังไม่มีข้อมูล", lo: "ຍັງບໍ່ມີຂໍ້ມູນ", en: "No data yet" },
  noDataDay: { th: "ยังไม่มีข้อมูลในวันที่เลือก", lo: "ຍັງບໍ່ມີຂໍ້ມູນໃນວັນທີເລືອກ", en: "No data for selected date" },
  rank: { th: "อันดับ", lo: "ອັນດັບ", en: "Rank" },
  qty: { th: "จำนวนรวม", lo: "ຈຳນວນລວມ", en: "Total Qty" },
  unit: { th: "หน่วย", lo: "ໜ່ວຍ", en: "Unit" },
  allEntries: { th: "รายการบันทึกทั้งหมด", lo: "ລາຍການບັນທຶກທັງໝົດ", en: "All Entries" },
  redeemNow: { th: "แลกได้ตอนนี้", lo: "ແລກໄດ້ດຽວນີ້", en: "Available now" },
  searchPlaceholder: { th: "ค้นหาด้วย รหัสพนักงาน / ชื่อ / หน่วยงาน", lo: "ຄົ້ນຫາດ້ວຍ ລະຫັດພະນັກງານ / ຊື່ / ຫນ່ວຍງານ", en: "Search by ID / Name / Department" },
  searchResults: { th: "พบ {n} รายการ", lo: "ພົບ {n} ລາຍການ", en: "{n} results found" },
  noResults: { th: "ไม่พบข้อมูลผู้เข้าร่วม", lo: "ບໍ່ພົບຂໍ້ມູນຜູ້ເຂົ້າຮ່ວມ", en: "No participants found" },
  clearSearch: { th: "ล้าง", lo: "ລ້າງ", en: "Clear" },
};

const t = (k: keyof typeof UI, lang: Lang) => tr(UI[k], lang);

// ===== Config =====
type WasteType = {
  key: string;
  label: L;
  perPoint: number;
  unit: L;
  criteriaLabel?: L;
};

const WASTE_TYPES: WasteType[] = [
  { key: "paper", label: { th: "กระดาษ (เจ้ย)", lo: "ເຈ້ຍ", en: "Paper" }, perPoint: 5, unit: { th: "กก.", lo: "ກກ.", en: "kg" } },
  { key: "paper_box", label: { th: "กล่องกระดาษ", lo: "ກ່ອງເຈ້ຍ", en: "Cardboard box" }, perPoint: 10, unit: { th: "กก.", lo: "ກກ.", en: "kg" } },
  { key: "plastic", label: { th: "พลาสติก, ถุงพลาสติก", lo: "ພລາສຕິກ, ຖົງພລາສຕິກ", en: "Plastic / bags" }, perPoint: 5, unit: { th: "กก.", lo: "ກກ.", en: "kg" } },
  { key: "water_small", label: { th: "ขวดพลาสติกเล็ก 250 ม.ล", lo: "ຂວດພລາສຕິກນ້ອຍ 250 ມລ", en: "Small plastic bottle 250 ml" }, perPoint: 30, unit: { th: "ขวด", lo: "ຂວດ", en: "bottle" } },
  { key: "water_mid", label: { th: "ขวดพลาสติกกลาง 550 ม.ล", lo: "ຂວດພລາສຕິກກາງ 550 ມລ", en: "Medium plastic bottle 550 ml" }, perPoint: 20, unit: { th: "ขวด", lo: "ຂວດ", en: "bottle" } },
  { key: "water_big", label: { th: "ขวดพลาสติกใหญ่ 1,500 ม.ล", lo: "ຂວດພລາສຕິກໃຫຍ່ 1,500 ມລ", en: "Large plastic bottle 1500 ml" }, perPoint: 12, unit: { th: "ขวด", lo: "ຂວດ", en: "bottle" } },
  { key: "can_drink", label: { th: "กระป๋องเครื่องดื่มทุกชนิด", lo: "ກະປ໋ອງເຄື່ອງດື່ມທຸກຊະນິດ", en: "Beverage cans (all)" }, perPoint: 1, unit: { th: "กก.", lo: "ກກ.", en: "kg" } },
  { key: "can_spray", label: { th: "กระป๋องสเปรย์ฉีดยุง", lo: "ກະປ໋ອງສະເປຣ", en: "Spray can" }, perPoint: 5, unit: { th: "ขวด", lo: "ຂວດ", en: "can" } },
  { key: "m150", label: { th: "M150", lo: "M150", en: "M150" }, perPoint: 20, unit: { th: "ขวด", lo: "ຂວດ", en: "bottle" } },
  { key: "beer_small", label: { th: "ขวดเบียร์+Somersby (330 มล.)", lo: "ຂວດເບຍ+Somersby (330 ມລ)", en: "Beer/Somersby 330 ml" }, perPoint: 12, unit: { th: "ขวด", lo: "ຂວດ", en: "bottle" } },
  { key: "beer_big", label: { th: "ขวดเบียร์ (640 มล.)", lo: "ຂວດເບຍ (640 ມລ)", en: "Beer 640 ml" }, perPoint: 6, unit: { th: "ขวด", lo: "ຂວດ", en: "bottle" } },
  { key: "glass_other", label: { th: "ขวดแก้วทุกชนิด (700 มล.)", lo: "ຂວດແກ້ວທຸກຊະນິດ (700 ມລ)", en: "Glass bottle (700 ml)" }, perPoint: 20, unit: { th: "ขวด", lo: "ຂວດ", en: "bottle" } },
  { key: "kitchenware", label: { th: "เครื่องครัวไม่ใช้แล้วต้องการทิ้ง", lo: "ເຄື່ອງຄົວທີ່ບໍ່ໃຊ້ແລ້ວ", en: "Unused kitchenware" }, perPoint: 20, unit: { th: "ชิ้น", lo: "ຊິ້ນ", en: "pc" } },
  { key: "appliance_small", label: { th: "เครื่องใช้ไฟฟ้าจำนวน", lo: "ເຄື່ອງໃຊ້ໄຟຟ້າ", en: "Small appliance" }, perPoint: 10, unit: { th: "ชิ้น", lo: "ຊິ້ນ", en: "pc" } },
  { key: "appliance_big", label: { th: "เครื่องซักผ้าเสีย, TV, หม้อ", lo: "ເຄື່ອງຊັກຜ້າເສຍ, TV, ໝໍ້", en: "Washer/TV/pot (broken)" }, perPoint: 1, unit: { th: "ชิ้น", lo: "ຊິ້ນ", en: "pc" } },
  { key: "ewaste", label: { th: "ขยะอิเล็กทรอนิกส์, ถ่านไฟฉาย", lo: "ຂີ້ເຫຍື້ອອີເລັກໂທຣນິກ, ຖ່ານໄຟສາຍ", en: "E-waste / batteries" }, perPoint: 10, unit: { th: "ชิ้น", lo: "ຊິ້ນ", en: "pc" } },
  { key: "clothes", label: { th: "เสื้อผ้า", lo: "ເສື້ອຜ້າ", en: "Clothes" }, perPoint: 10, unit: { th: "ชิ้น", lo: "ຊິ້ນ", en: "pc" } },
  {
    key: "oil",
    label: { th: "น้ำมันหมู น้ำมันพืช", lo: "ນ້ຳມັນໝູ ນ້ຳມັນພືດ", en: "Lard / vegetable oil" },
    perPoint: 1 / 3,
    unit: { th: "ขวด 1 ลิตร", lo: "ຂວດ 1 ລິດ", en: "bottle (1 L)" },
    criteriaLabel: {
      th: "1 ขวด 1 ลิตร = 3 คะแนน",
      lo: "1 ຂວດ 1 ລິດ = 3 ຄະແນນ",
      en: "1 bottle (1 L) = 3 points",
    },
  },
];

const ROOMS_L: L[] = [
  { th: "ห้อง 1", lo: "ຫ້ອງ 1", en: "Room 1" },
  { th: "ห้อง 2", lo: "ຫ້ອງ 2", en: "Room 2" },
  { th: "ห้อง 3", lo: "ຫ້ອງ 3", en: "Room 3" },
  { th: "ห้อง 4", lo: "ຫ້ອງ 4", en: "Room 4" },
  { th: "ห้อง 5", lo: "ຫ້ອງ 5", en: "Room 5" },
];
const ROOMS = ["ห้อง 1", "ห้อง 2", "ห้อง 3", "ห้อง 4", "ห้อง 5"] as const;
const TOTAL_ROOM = "ห้องรวม";

const WASTE_CATEGORIES: { title: L; icon: string; start: number; end: number }[] = [
  { title: { th: "หมวดที่ 1 : ขยะรีไซเคิลทั่วไป", lo: "ໝວດທີ 1 : ຂີ້ເຫຍື້ອຣີໄຊເຄີນທົ່ວໄປ", en: "Category 1: General Recyclables" }, icon: "📦", start: 0, end: 5 },
  { title: { th: "หมวดที่ 2 : ภาชนะและเครื่องดื่ม", lo: "ໝວດທີ 2 : ພາຊະນະແລະເຄື່ອງດື່ມ", en: "Category 2: Containers & Beverages" }, icon: "🥫", start: 6, end: 11 },
  { title: { th: "หมวดที่ 3 : ไฟฟ้าและอื่นๆ", lo: "ໝວດທີ 3 : ໄຟຟ້າແລະອື່ນໆ", en: "Category 3: Electrical & Others" }, icon: "🔌", start: 12, end: 17 },
];

const categoryOfIndex = (idx: number) => WASTE_CATEGORIES.find((c) => idx >= c.start && idx <= c.end);

const FAVORITE_KEYS = ["water_small", "water_mid", "water_big", "beer_small", "beer_big", "glass_other"];

const MOBILE_CATS: { title: L; icon: string; keys: string[] }[] = [
  { title: { th: "⭐ ใช้งานบ่อย", lo: "⭐ ໃຊ້ເລ້ານິຍົມ", en: "⭐ Frequently Used" }, icon: "⭐", keys: FAVORITE_KEYS },
  ...WASTE_CATEGORIES.map((c) => ({
    title: c.title,
    icon: c.icon,
    keys: WASTE_TYPES.slice(c.start, c.end + 1).map((w) => w.key),
  })),
];

const REWARD_THRESHOLD = 10;

type Row = {
  id: string;
  date: string;
  code: string;
  name: string;
  unit_org: string;
  room: string;
  wasteType: string;
  quantity: string;
  qtyUnit: string;
  points: string;
  reward: string;
  signature: string;
};

const calcPoints = (wasteKey: string, qty: number) => {
  const w = WASTE_TYPES.find((x) => x.key === wasteKey);
  if (!w || !w.perPoint) return 0;
  return Math.round((qty / w.perPoint) * 100) / 100;
};

const fmtPts = (n: number) =>
  Number.isInteger(n) ? String(n) : n.toFixed(2).replace(/\.?0+$/, "");

const REWARDS: { label: L; unit: L; cost: number }[] = [
  { label: { th: "ไข่", lo: "ໄຂ່", en: "Egg" }, unit: { th: "ฟอง", lo: "ໜ່ວຍ", en: "pc" }, cost: 1 },
  { label: { th: "มาม่า", lo: "ໝີ່ກຶ່ງສຳເລັດຮູບ", en: "Instant noodle" }, unit: { th: "ซอง", lo: "ຊອງ", en: "pack" }, cost: 2 },
  { label: { th: "ปลากระป๋อง", lo: "ປາກະປ໋ອງ", en: "Canned fish" }, unit: { th: "กระป๋อง", lo: "ກະປ໋ອງ", en: "can" }, cost: 6 },
];

type Matrix = Record<string, Record<string, string>>;

type FormState = {
  date: string;
  code: string;
  name: string;
  unit_org: string;
  reward: string;
  signature: string;
  matrix: Matrix;
};

const emptyMatrix = (): Matrix =>
  Object.fromEntries(
    WASTE_TYPES.map((w) => [w.key, Object.fromEntries(ROOMS.map((r) => [r, ""]))]),
  );

const newForm = (): FormState => ({
  date: new Date().toISOString().slice(0, 10),
  code: "",
  name: "",
  unit_org: "",
  reward: "",
  signature: "",
  matrix: emptyMatrix(),
});

type DbRow = {
  id: string;
  date: string;
  code: string | null;
  name: string;
  unit_org: string | null;
  room: string;
  waste_type: string;
  quantity: string | null;
  qty_unit: string | null;
  points: string | null;
  reward: string | null;
  signature: string | null;
};

const fromDb = (r: DbRow): Row => ({
  id: r.id,
  date: r.date,
  code: r.code ?? "",
  name: r.name,
  unit_org: r.unit_org ?? "",
  room: r.room,
  wasteType: r.waste_type,
  quantity: r.quantity ?? "0",
  qtyUnit: r.qty_unit ?? "",
  points: r.points ?? "0",
  reward: r.reward ?? "",
  signature: r.signature ?? "",
});

const criteriaText = (w: WasteType, lang: Lang) => {
  if (w.criteriaLabel) return tr(w.criteriaLabel, lang);
  return `${w.perPoint} ${tr(w.unit, lang)} ${t("pointPerOne", lang)}`;
};

function LoginScreen({ onLogin }: { onLogin: (user: AppUser) => void }) {
  const [u, setU] = useState("");
  const [p, setP] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setBusy(true);
    setErr("");
    const { user, error } = await loginUser(u.trim(), p);
    setBusy(false);
    if (error || !user) {
      setErr(error || "เข้าสู่ระบบไม่สำเร็จ");
    } else {
      onLogin(user);
    }
  };
  return (
    <div className="flex min-h-screen items-center justify-center bg-gradient-to-br from-primary/10 via-background to-accent/30 p-4">
      <Card className="w-full max-w-sm p-6 shadow-lg">
        <div className="mb-6 flex items-center gap-3">
          <img
            src="/S__73916420.jpg"
            alt="โลโก้โครงการขยะแลกของรางวัล"
            className="h-11 w-11 rounded-xl object-cover shadow-md"
          />
          <div>
            <h1 className="text-base font-semibold">เข้าสู่ระบบ</h1>
            <p className="text-xs text-muted-foreground">โครงการขยะแลกของรางวัล</p>
          </div>
        </div>
        <form onSubmit={submit} className="space-y-3">
          <div>
            <label className="mb-1 block text-sm font-medium">Username</label>
            <Input value={u} onChange={(e) => setU(e.target.value)} autoFocus disabled={busy} />
          </div>
          <div>
            <label className="mb-1 block text-sm font-medium">Password</label>
            <Input type="password" value={p} onChange={(e) => setP(e.target.value)} disabled={busy} />
          </div>
          {err && <p className="text-sm text-destructive">{err}</p>}
          <Button type="submit" className="w-full" disabled={busy}>
            {busy ? <Loader2 className="mr-1 h-4 w-4 animate-spin" /> : null}
            เข้าสู่ระบบ
          </Button>
        </form>
      </Card>
    </div>
  );
}

export default function WasteRewardsApp() {
  const [user, setUser] = useState<AppUser | null>(null);
  const [checked, setChecked] = useState(false);
  useEffect(() => {
    setUser(getStoredSession());
    setChecked(true);
  }, []);
  if (!checked) return null;
  if (!user) return <LoginScreen onLogin={(u) => setUser(u)} />;
  return <IndexApp currentUser={user} onLogout={() => { logoutUser(); setUser(null); }} />;
}

function IndexApp({ currentUser, onLogout }: { currentUser: AppUser; onLogout: () => void }) {
  const [lang, setLang] = useState<Lang>(() => {
    if (typeof window === "undefined") return "th";
    return (localStorage.getItem("waste-rewards-lang") as Lang) || "th";
  });
  useEffect(() => {
    if (typeof window !== "undefined") localStorage.setItem("waste-rewards-lang", lang);
  }, [lang]);

  const [rows, setRows] = useState<Row[]>([]);
  const [loaded, setLoaded] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const [form, setForm] = useState<FormState>(newForm());
  const [activeTab, setActiveTab] = useState<"daily" | "import" | "search" | "admin">("daily");
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [expandedPerson, setExpandedPerson] = useState<string | null>(null);
  const [personScores, setPersonScores] = useState<Map<string, { total: number; remaining: number }>>(new Map());
  const [openCats, setOpenCats] = useState<Set<number>>(new Set([1]));

  const [searchText, setSearchText] = useState("");
  const [debouncedSearch, setDebouncedSearch] = useState("");
  useEffect(() => {
    const timer = setTimeout(() => setDebouncedSearch(searchText), 300);
    return () => clearTimeout(timer);
  }, [searchText]);

  const [personSearchText, setPersonSearchText] = useState("");
  const [debouncedPersonSearch, setDebouncedPersonSearch] = useState("");
  useEffect(() => {
    const timer = setTimeout(() => setDebouncedPersonSearch(personSearchText), 300);
    return () => clearTimeout(timer);
  }, [personSearchText]);

  useEffect(() => {
    let mounted = true;
    (async () => {
      const { data, error } = await supabase
        .from("waste_entries")
        .select("*")
        .order("created_at", { ascending: true });
      if (!mounted) return;
      if (error) toast.error("Load failed: " + error.message);
      else setRows((data as DbRow[]).map(fromDb));
      setLoaded(true);
    })();

    loadEmpMaster();

    const channel = supabase
      .channel("waste_entries_changes")
      .on(
        "postgres_changes",
        { event: "INSERT", schema: "public", table: "waste_entries" },
        (payload) => {
          const r = fromDb(payload.new as DbRow);
          setRows((rs) => (rs.some((x) => x.id === r.id) ? rs : [...rs, r]));
        },
      )
      .on(
        "postgres_changes",
        { event: "DELETE", schema: "public", table: "waste_entries" },
        (payload) => {
          const id = (payload.old as { id?: string }).id;
          if (id) setRows((rs) => rs.filter((x) => x.id !== id));
        },
      )
      .subscribe();

    return () => {
      mounted = false;
      supabase.removeChannel(channel);
    };
  }, []);

  const wasteTotals = useMemo(() => {
    const m = new Map<string, { sum: number; points: number }>();
    WASTE_TYPES.forEach((w) => {
      const sum = ROOMS.reduce((s, r) => s + (parseFloat(form.matrix[w.key]?.[r]) || 0), 0);
      m.set(w.key, { sum, points: calcPoints(w.key, sum) });
    });
    return m;
  }, [form.matrix]);

  const roomTotals = useMemo(() => {
    const m = new Map<string, { qty: number; points: number }>();
    ROOMS.forEach((r) => {
      let qty = 0;
      let points = 0;
      WASTE_TYPES.forEach((w) => {
        const q = parseFloat(form.matrix[w.key]?.[r]) || 0;
        qty += q;
        points += calcPoints(w.key, q);
      });
      m.set(r, { qty, points });
    });
    return m;
  }, [form.matrix]);

  const grandQty = useMemo(
    () => ROOMS.reduce((s, r) => s + (roomTotals.get(r)?.qty || 0), 0),
    [roomTotals],
  );
  const grandPoints = useMemo(
    () => WASTE_TYPES.reduce((s, w) => s + (wasteTotals.get(w.key)?.points || 0), 0),
    [wasteTotals],
  );

  const toggleCat = (idx: number) =>
    setOpenCats((prev) => {
      const next = new Set(prev);
      if (next.has(idx)) next.delete(idx);
      else next.add(idx);
      return next;
    });

  const catTotals = useMemo(
    () =>
      MOBILE_CATS.map((cat) => {
        let qty = 0;
        let points = 0;
        cat.keys.forEach((key) => {
          const tt = wasteTotals.get(key);
          if (tt) {
            qty += tt.sum;
            points += tt.points;
          }
        });
        return { qty, points };
      }),
    [wasteTotals],
  );

  const savingRef = useRef(false);
  const addRow = async () => {
    if (savingRef.current || syncing) return;
    if (!form.name.trim()) {
      toast.error(t("nameRequired", lang));
      return;
    }

    const merged: { waste: string; qty: number; points: number }[] = [];
    WASTE_TYPES.forEach((w) => {
      const tt = wasteTotals.get(w.key);
      if (tt && tt.sum > 0) merged.push({ waste: w.key, qty: tt.sum, points: tt.points });
    });
    if (merged.length === 0) {
      toast.error("กรุณากรอกจำนวนอย่างน้อย 1 ช่อง");
      return;
    }

    const payload = merged.map((c) => ({
      date: form.date,
      code: form.code,
      name: form.name,
      unit_org: form.unit_org,
      room: TOTAL_ROOM,
      waste_type: c.waste,
      quantity: String(c.qty),
      qty_unit: tr(WASTE_TYPES.find((w) => w.key === c.waste)!.unit, lang),
      points: fmtPts(c.points),
      reward: "",
      signature: "",
    }));

    savingRef.current = true;
    setSyncing(true);
    const { data, error } = await supabase.from("waste_entries").insert(payload).select();
    setSyncing(false);
    savingRef.current = false;
    if (error) {
      toast.error("Save failed: " + error.message);
      return;
    }
    if (data) {
      const inserted = (data as DbRow[]).map(fromDb);
      setRows((rs) => {
        const ids = new Set(rs.map((r) => r.id));
        return [...rs, ...inserted.filter((r) => !ids.has(r.id))];
      });
    }

    saveLocalBackup(rows as unknown as Record<string, unknown>[], []);
    toast.success(`${t("saveEntry", lang)} (+${fmtPts(grandPoints)} ${t("pointsWord", lang)})`);

    // Auto-update accumulated_scores
    (async () => {
      const empName = form.name.trim();
      const empCode = form.code.trim();
      const unit = (form.unit_org ?? "").trim();
      const { data: matches, error: findErr } = empCode
        ? await supabase.from("accumulated_scores").select("*").eq("employee_code", empCode)
        : await supabase.from("accumulated_scores").select("*").ilike("employee_name", empName);
      if (findErr) {
        toast.error("Score sync failed: " + findErr.message);
        return;
      }

      const candidates = empCode
        ? (matches ?? [])
        : (matches ?? []).filter((m) => (m.department ?? "").trim() === unit);
      const existing = candidates.length === 1 ? candidates[0] : null;

      const { error: saveErr } = existing
        ? await supabase
            .from("accumulated_scores")
            .update({
              total_score: (existing.total_score || 0) + grandPoints,
              remaining_score: (existing.remaining_score || 0) + grandPoints,
              updated_at: new Date().toISOString(),
            })
            .eq("id", existing.id)
        : await supabase.from("accumulated_scores").insert({
            employee_code: empCode || `AUTO-${Date.now()}`,
            employee_name: empName,
            department: form.unit_org || null,
            total_score: grandPoints,
            remaining_score: grandPoints,
            import_date: form.date,
            file_name: "auto-from-daily-entry",
          });
      if (saveErr) toast.error("Score sync failed: " + saveErr.message);
      else await Promise.all([loadPersonScores(), loadEmpMaster()]);
    })();

    const totalForPerson =
      rows
        .filter((x) => x.name.trim() === form.name.trim())
        .reduce((s, x) => s + (parseFloat(x.points) || 0), 0) + grandPoints;
    if (totalForPerson >= REWARD_THRESHOLD) {
      toast.success(`${form.name} = ${fmtPts(totalForPerson)} ${t("pointsWord", lang)} — ${t("canRedeem", lang)}`, {
        duration: 6000,
      });
    }

    setForm({ ...newForm(), date: form.date });
  };

  const removeRow = async (id: string) => {
    if (!confirm("Delete this entry?")) return;
    const prev = rows;
    setRows((rs) => rs.filter((r) => r.id !== id));
    const { error } = await supabase.from("waste_entries").delete().eq("id", id);
    if (error) {
      toast.error("Delete failed: " + error.message);
      setRows(prev);
    }
  };

  const [backupBusy, setBackupBusy] = useState(false);
  const [showBackups, setShowBackups] = useState(false);
  const [localBackups, setLocalBackups] = useState<{ timestamp: string; entries: number; scores: number }[]>([]);
  const fileRestoreRef = useRef<HTMLInputElement>(null);

  // Employee Master from accumulated_scores (populated via Import Excel)
  type EmpMaster = { employee_code: string; employee_name: string; department: string | null };
  const [empMaster, setEmpMaster] = useState<EmpMaster[]>([]);
  const [empLoading, setEmpLoading] = useState(false);
  const [codeSuggest, setCodeSuggest] = useState<EmpMaster[]>([]);
  const [nameSuggest, setNameSuggest] = useState<EmpMaster[]>([]);
  const [showCodeList, setShowCodeList] = useState(false);
  const [showNameList, setShowNameList] = useState(false);

  const loadEmpMaster = async () => {
    setEmpLoading(true);
    try {
      let all: EmpMaster[] = [];
      let from = 0;
      const PAGE = 1000;
      let more = true;
      while (more) {
        const { data, error } = await supabase
          .from("accumulated_scores")
          .select("employee_code, employee_name, department")
          .order("employee_code", { ascending: true })
          .range(from, from + PAGE - 1);
        if (error) throw error;
        const rows = (data ?? []) as EmpMaster[];
        all = all.concat(rows);
        more = rows.length === PAGE;
        from += PAGE;
      }
      setEmpMaster(all);
    } catch {
      // silently fail — autocomplete just won't work
    }
    setEmpLoading(false);
  };

  const [editingId, setEditingId] = useState<string | null>(null);
  const [editDraft, setEditDraft] = useState<Row | null>(null);

  const startEdit = (r: Row) => {
    setEditingId(r.id);
    setEditDraft({ ...r });
  };
  const cancelEdit = () => {
    setEditingId(null);
    setEditDraft(null);
  };
  const saveEdit = async () => {
    if (!editDraft) return;
    const qty = parseFloat(editDraft.quantity) || 0;
    const points = calcPoints(editDraft.wasteType, qty);
    const updated: Row = { ...editDraft, points: fmtPts(points) };
    const prev = rows;
    setRows((rs) => rs.map((r) => (r.id === updated.id ? updated : r)));
    setEditingId(null);
    setEditDraft(null);
    const { error } = await supabase
      .from("waste_entries")
      .update({
        date: updated.date,
        code: updated.code,
        name: updated.name,
        unit_org: updated.unit_org,
        room: updated.room,
        waste_type: updated.wasteType,
        quantity: updated.quantity,
        qty_unit: updated.qtyUnit,
        points: updated.points,
        reward: "",
        signature: "",
      })
      .eq("id", updated.id);
    if (error) {
      toast.error("Update failed: " + error.message);
      setRows(prev);
    } else {
      toast.success("Updated");
    }
  };

  const clearAll = async () => {
    if (!confirm("Clear ALL cloud data?")) return;
    const ids = rows.map((r) => r.id);
    if (ids.length === 0) return;
    const { error } = await supabase.from("waste_entries").delete().in("id", ids);
    if (error) toast.error("Clear failed: " + error.message);
    else setRows([]);
  };

  const [selectedDate, setSelectedDate] = useState(() => new Date().toISOString().slice(0, 10));

  const totalPoints = useMemo(
    () => rows.filter((r) => r.date === selectedDate).reduce((s, r) => s + (parseFloat(r.points) || 0), 0),
    [rows, selectedDate],
  );

  const [summaryDate, setSummaryDate] = useState(() => new Date().toISOString().slice(0, 10));

  const dateDefaulted = useRef(false);
  useEffect(() => {
    if (!loaded || dateDefaulted.current) return;
    dateDefaulted.current = true;
    const today = new Date().toISOString().slice(0, 10);
    if (rows.some((r) => r.date === today)) return;
    const latest = rows.reduce((max, r) => (r.date > max ? r.date : max), "");
    if (latest) {
      setSelectedDate(latest);
      setSummaryDate(latest);
    }
  }, [loaded, rows]);

  const dailySummary = useMemo(() => {
    const dayRows = rows.filter((r) => r.date === summaryDate);
    const byType = new Map<string, { qty: number; points: number }>();
    WASTE_TYPES.forEach((w) => byType.set(w.key, { qty: 0, points: 0 }));
    dayRows.forEach((r) => {
      const tt = byType.get(r.wasteType);
      if (!tt) return;
      tt.qty += parseFloat(r.quantity) || 0;
      tt.points += parseFloat(r.points) || 0;
    });
    const totalQty = dayRows.reduce((s, r) => s + (parseFloat(r.quantity) || 0), 0);
    const totalPts = dayRows.reduce((s, r) => s + (parseFloat(r.points) || 0), 0);
    return { byType, totalQty, totalPts, count: dayRows.length };
  }, [rows, summaryDate]);

  const byPerson = useMemo(() => {
    const m = new Map<
      string,
      {
        name: string;
        code: string;
        unit_org: string;
        points: number;
        byWaste: Record<string, number>;
      }
    >();
    rows.filter((r) => r.date === selectedDate).forEach((r) => {
      const key = `${r.name || "-"}||${r.unit_org || "-"}`;
      const prev = m.get(key) ?? {
        name: r.name || "-",
        code: r.code,
        unit_org: r.unit_org,
        points: 0,
        byWaste: {},
      };
      prev.points += parseFloat(r.points) || 0;
      prev.byWaste[r.wasteType] =
        (prev.byWaste[r.wasteType] || 0) + (parseFloat(r.quantity) || 0);
      if (!prev.code && r.code) prev.code = r.code;
      if (!prev.unit_org && r.unit_org) prev.unit_org = r.unit_org;
      m.set(key, prev);
    });
    return Array.from(m.values()).sort((a, b) => b.points - a.points);
  }, [rows, selectedDate]);

  const filteredByPerson = useMemo(() => {
    if (!debouncedPersonSearch.trim()) return byPerson;
    const kw = debouncedPersonSearch.toLowerCase();
    return byPerson.filter(
      (p) =>
        p.name.toLowerCase().includes(kw) ||
        (p.code || "").toLowerCase().includes(kw) ||
        (p.unit_org || "").toLowerCase().includes(kw),
    );
  }, [byPerson, debouncedPersonSearch]);

  // Load ALL scores independent of render state (avoids stale `byPerson` closure).
  // Keyed by employee code AND by name+department so duplicate names don't collide.
  const personKey = (p: { name: string; unit_org: string }) => `${p.name}||${p.unit_org || "-"}`;
  const nd = (name: string, dept: string | null | undefined) =>
    `nd:${(name || "").trim().toLowerCase()}|${(dept || "").trim().toLowerCase()}`;
  const getPersonScore = (p: { name: string; code: string; unit_org: string }) =>
    (p.code && personScores.get(`c:${p.code.trim()}`)) || personScores.get(nd(p.name, p.unit_org));

  const loadPersonScores = async () => {
    const m = new Map<string, { total: number; remaining: number }>();
    const PAGE = 1000;
    let from = 0;
    let more = true;
    while (more) {
      const { data, error } = await supabase
        .from("accumulated_scores")
        .select("employee_code, employee_name, department, total_score, remaining_score")
        .order("employee_code", { ascending: true })
        .range(from, from + PAGE - 1);
      if (error) {
        toast.error("Load scores failed: " + error.message);
        return;
      }
      const batch = (data ?? []) as {
        employee_code: string;
        employee_name: string;
        department: string | null;
        total_score: number;
        remaining_score: number;
      }[];
      for (const r of batch) {
        const v = { total: r.total_score, remaining: r.remaining_score };
        if (r.employee_code) m.set(`c:${r.employee_code.trim()}`, v);
        m.set(nd(r.employee_name, r.department), v);
      }
      more = batch.length === PAGE;
      from += PAGE;
    }
    setPersonScores(m);
  };

  // Load on mount only; explicit reloads happen after save / exchange.
  useEffect(() => {
    loadPersonScores();
  }, []);

  // Realtime: any change to accumulated_scores (daily save, exchange, import, other devices)
  // refreshes scores + autocomplete. Debounced because Excel imports fire thousands of events.
  useEffect(() => {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const refresh = () => {
      clearTimeout(timer);
      timer = setTimeout(() => {
        loadPersonScores();
        loadEmpMaster();
      }, 600);
    };
    const channel = supabase
      .channel("accumulated_scores_changes")
      .on("postgres_changes", { event: "*", schema: "public", table: "accumulated_scores" }, refresh)
      .subscribe();
    return () => {
      clearTimeout(timer);
      supabase.removeChannel(channel);
    };
  }, []);

  const wasteLabel = (k: string) =>
    tr(WASTE_TYPES.find((w) => w.key === k)?.label ?? { th: k, lo: k, en: k }, lang);

  const filteredRows = useMemo(() => {
    if (!debouncedSearch.trim()) return rows;
    const kw = debouncedSearch.toLowerCase();
    return rows.filter(
      (r) =>
        r.code.toLowerCase().includes(kw) ||
        r.name.toLowerCase().includes(kw) ||
        r.unit_org.toLowerCase().includes(kw),
    );
  }, [rows, debouncedSearch]);

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

  const exportExcel = async () => {
    const data = rows.map((r) => ({
      [t("date", lang)]: r.date,
      [t("code", lang)]: r.code,
      [t("name", lang)]: r.name,
      [t("org", lang)]: r.unit_org,
      Room: r.room,
      [t("wasteType", lang)]: wasteLabel(r.wasteType),
      [t("qty", lang)]: r.quantity,
      [t("unit", lang)]: r.qtyUnit,
      [t("points", lang)]: r.points,
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), "Entries");

    const summary = [
      ...byPerson.map((p, i) => {
        const sc = getPersonScore(p);
        return {
          "#": i + 1,
          [t("code", lang)]: p.code,
          [t("name", lang)]: p.name,
          [t("org", lang)]: p.unit_org,
          [t("totalPointsAll", lang)]: fmtPts(p.points),
          "คะแนนสะสมรวม": sc ? sc.total : 0,
          "คะแนนคงเหลือ": sc ? sc.remaining : 0,
          [t("status", lang)]: p.points >= REWARD_THRESHOLD ? t("canRedeem", lang) : "-",
        };
      }),
      {
        "#": "",
        [t("code", lang)]: "",
        [t("name", lang)]: t("grandTotalPts", lang),
        [t("org", lang)]: "",
        [t("totalPointsAll", lang)]: fmtPts(totalPoints),
        "คะแนนสะสมรวม": "",
        "คะแนนคงเหลือ": "",
        [t("status", lang)]: "",
      },
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "PerPerson");

    const detail = byPerson.map((p) => {
      const row: Record<string, string | number> = {
        [t("name", lang)]: p.name,
        [t("code", lang)]: p.code,
        [t("org", lang)]: p.unit_org,
      };
      WASTE_TYPES.forEach((w) => {
        row[`${tr(w.label, lang)} (${tr(w.unit, lang)})`] = p.byWaste[w.key] || 0;
      });
      row[t("totalPointsAll", lang)] = fmtPts(p.points);
      return row;
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detail), "Detail");

    {
      const daily: Record<string, string | number>[] = WASTE_TYPES.map((w, i) => {
        const tt = dailySummary.byType.get(w.key);
        return {
          [t("rank", lang)]: i + 1,
          [t("wasteType", lang)]: tr(w.label, lang),
          [t("criteria", lang)]: criteriaText(w, lang),
          [t("qty", lang)]: tt?.qty ?? 0,
          [t("unit", lang)]: tr(w.unit, lang),
          [t("points", lang)]: tt?.points ?? 0,
        };
      });
      daily.push({
        [t("rank", lang)]: "",
        [t("wasteType", lang)]: t("grandTotalPts", lang),
        [t("criteria", lang)]: "",
        [t("qty", lang)]: dailySummary.totalQty,
        [t("unit", lang)]: "-",
        [t("points", lang)]: dailySummary.totalPts,
      });
      const dailyWs = XLSX.utils.json_to_sheet(daily);
      dailyWs["!cols"] = [
        { wch: 6 },
        { wch: 28 },
        { wch: 32 },
        { wch: 12 },
        { wch: 14 },
        { wch: 10 },
      ];
      XLSX.utils.book_append_sheet(wb, dailyWs, `Daily ${summaryDate}`);
    }

    let allExchanges: ExchangeRow[] = [];
    try {
      allExchanges = await fetchAllExchanges();
    } catch {
      // exchanges table may be empty
    }
    if (allExchanges.length > 0) {
      const exData = allExchanges.map((ex) => ({
        "รหัสพนักงาน": ex.employee_code ?? "",
        "ชื่อ-นามสกุล": ex.employee_name,
        "ของรางวัล": ex.reward_item,
        "จำนวน": ex.quantity,
        "คะแนนที่ใช้": ex.points_used,
        "ดำเนินการโดย": ex.exchanged_by ?? "",
        "วันที่แลก": new Date(ex.exchanged_at).toLocaleString("th-TH"),
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(exData), "RewardExchanges");
    }

    XLSX.writeFile(wb, `waste-rewards-${new Date().toISOString().slice(0, 10)}.xlsx`);
    toast.success("Excel exported");
  };

  const exportPdf = () => {
    const doc = new jsPDF({ orientation: "landscape" });
    doc.text("Waste Recycling Rewards Report", 14, 14);
    autoTable(doc, {
      startY: 20,
      head: [["Date", "Code", "Name", "Org", "Room", "Type", "Qty", "Unit", "Pts"]],
      body: rows.map((r) => [
        r.date,
        r.code,
        r.name,
        r.unit_org,
        r.room,
        tr(WASTE_TYPES.find((w) => w.key === r.wasteType)?.label ?? { th: r.wasteType, lo: r.wasteType, en: r.wasteType }, "en"),
        r.quantity,
        r.qtyUnit,
        r.points,
      ]),
      styles: { fontSize: 8 },
    });
    const finalY =
      (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 8;
    doc.text("Per-Person Summary", 14, finalY);
    autoTable(doc, {
      startY: finalY + 4,
      head: [["#", "Code", "Name", "Org", "Total Points", "Status"]],
      body: [
        ...byPerson.map((p, i) => [
          String(i + 1),
          p.code,
          p.name,
          p.unit_org,
          fmtPts(p.points),
          p.points >= REWARD_THRESHOLD ? "Reward eligible" : "-",
        ]),
        ["", "", "TOTAL", "", fmtPts(totalPoints), ""],
      ],
      styles: { fontSize: 9 },
    });

    if (dailySummary.count > 0) {
      const finalY2 =
        (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 8;
      doc.text(`Daily Waste Summary - ${summaryDate}`, 14, finalY2);
      autoTable(doc, {
        startY: finalY2 + 4,
        head: [["#", "Type", "Criteria", "Total Qty", "Unit", "Points"]],
        body: [
          ...WASTE_TYPES.map((w, i) => {
            const tt = dailySummary.byType.get(w.key)!;
            return [
              String(i + 1),
              tr(w.label, "en"),
              w.criteriaLabel ? tr(w.criteriaLabel, "en") : `${w.perPoint} ${tr(w.unit, "en")} = 1`,
              fmtPts(tt.qty),
              tr(w.unit, "en"),
              fmtPts(tt.points),
            ];
          }),
          ["", "TOTAL", "", fmtPts(dailySummary.totalQty), "-", fmtPts(dailySummary.totalPts)],
        ],
        styles: { fontSize: 8 },
      });
    }

    doc.save(`waste-rewards-${new Date().toISOString().slice(0, 10)}.pdf`);
    toast.success("PDF exported");
  };

  const setCell = (wasteKey: string, room: string, value: string) =>
    setForm((f) => ({
      ...f,
      matrix: {
        ...f.matrix,
        [wasteKey]: { ...f.matrix[wasteKey], [room]: value },
      },
    }));

  return (
    <div className="min-h-screen w-full overflow-x-hidden bg-background">
      <header className="sticky top-0 z-10 w-full overflow-x-hidden border-b bg-card/95 shadow-sm backdrop-blur">
        <div className="mx-auto flex w-full max-w-7xl min-w-0 flex-wrap items-center justify-between gap-2 px-2 py-2.5 sm:px-6 sm:py-4">
          <div className="flex min-w-0 max-w-full items-center gap-2 sm:gap-3">
            <img
              src="/S__73916420.jpg"
              alt="โลโก้โครงการขยะแลกของรางวัล"
              className="h-8 w-8 shrink-0 rounded-lg object-cover shadow-sm sm:h-10 sm:w-10"
            />
            <div className="min-w-0">
              <h1 className="break-words text-sm font-semibold leading-tight sm:text-base lg:text-lg">
                {t("appTitle", lang)}
              </h1>
              <p className="truncate text-[11px] text-muted-foreground sm:text-xs">{t("appSubtitle", lang)}</p>
            </div>
          </div>
          <div className="flex w-full min-w-0 flex-wrap items-center gap-1.5 sm:w-auto sm:gap-2">
            <div className="flex shrink-0 items-center rounded-md border bg-background p-0.5">
              <Languages className="ml-1.5 mr-1 h-3.5 w-3.5 text-muted-foreground" />
              {(["th", "lo", "en"] as Lang[]).map((l) => (
                <button
                  key={l}
                  onClick={() => setLang(l)}
                  className={`rounded px-2 py-1 text-xs font-medium transition-colors ${
                    lang === l
                      ? "bg-primary text-primary-foreground"
                      : "text-muted-foreground hover:bg-muted"
                  }`}
                  aria-pressed={lang === l}
                >
                  {l === "th" ? "ไทย" : l === "lo" ? "ລາວ" : "EN"}
                </button>
              ))}
            </div>
            <span className="hidden items-center gap-1 rounded-full bg-primary/10 px-2 py-1 text-xs text-primary sm:inline-flex">
              <Cloud className="h-3.5 w-3.5" />
              {syncing ? t("syncing", lang) : loaded ? t("synced", lang) : t("loading", lang)}
            </span>
            <Button variant="outline" size="sm" onClick={exportExcel}>
              <Download className="mr-1 h-4 w-4" /> {t("excel", lang)}
            </Button>
            <Button variant="outline" size="sm" onClick={exportPdf}>
              <FileText className="mr-1 h-4 w-4" /> {t("pdf", lang)}
            </Button>
            <Button
              variant="outline"
              size="sm"
              disabled={backupBusy}
              onClick={async () => {
                setBackupBusy(true);
                try {
                  await exportFullBackup();
                  toast.success(lang === "th" ? "ดาวน์โหลดฐานข้อมูลสำเร็จ" : lang === "lo" ? "ດາວໂຫລດຖານຂໍ້ມູນສຳເລັດ" : "Database downloaded");
                } catch (e) {
                  toast.error((e as Error).message);
                }
                setBackupBusy(false);
              }}
            >
              <HardDriveDownload className="mr-1 h-4 w-4" /> {lang === "th" ? "สำรอง" : lang === "lo" ? "ສຳຮອງ" : "Backup"}
            </Button>
            <Button
              variant="outline"
              size="sm"
              disabled={backupBusy}
              onClick={() => fileRestoreRef.current?.click()}
            >
              <ArchiveRestore className="mr-1 h-4 w-4" /> {lang === "th" ? "กู้คืน" : lang === "lo" ? "ກູ້ຄືນ" : "Restore"}
            </Button>
            <Button
              variant="outline"
              size="sm"
              onClick={() => {
                setLocalBackups(getLocalBackups());
                setShowBackups(true);
              }}
            >
              <History className="mr-1 h-4 w-4" /> {lang === "th" ? "เวอร์ชัน" : lang === "lo" ? "ເວີຊັນ" : "Versions"}
            </Button>
            <input
              ref={fileRestoreRef}
              type="file"
              accept=".xlsx"
              className="hidden"
              onChange={async (e) => {
                const file = e.target.files?.[0];
                if (!file) return;
                if (!confirm(lang === "th" ? "การกู้คืนจะล้างข้อมูลปัจจุบันทั้งหมดและแทนที่ด้วยข้อมูลจากไฟล์ ต้องการดำเนินการต่อ?" : lang === "lo" ? "ການກູ້ຄືນຈະລ້າງຂໍ້ມູນປັດຈຸບັນທັງໝົດແລ້ວແທນທີ່ດ້ວຍຂໍ້ມູນຈາກໄຟລ໌ ຕ້ອງການດຳເນີນການຕໍ່?" : "Restore will clear all current data and replace with data from file. Continue?")) return;
                setBackupBusy(true);
                try {
                  await importBackup(file);
                  toast.success(lang === "th" ? "กู้คืนฐานข้อมูลสำเร็จ" : lang === "lo" ? "ກູ້ຄືນຖານຂໍ້ມູນສຳເລັດ" : "Database restored");
                  window.location.reload();
                } catch (e) {
                  toast.error((e as Error).message);
                  setBackupBusy(false);
                }
                e.target.value = "";
              }}
            />
            <span className="hidden items-center gap-1.5 text-xs text-muted-foreground sm:inline-flex">
              {currentUser.role === "admin" ? <Shield className="h-3.5 w-3.5 text-primary" /> : null}
              {currentUser.full_name}
            </span>
            <Button variant="outline" size="sm" onClick={onLogout}>
              ออกจากระบบ
            </Button>
          </div>
        </div>
      </header>

      {/* Desktop Tab Navigation */}
      <nav className="sticky top-[73px] z-10 hidden border-b bg-card/95 backdrop-blur lg:block">
        <div className="mx-auto flex max-w-7xl gap-1 px-4 sm:px-6">
          {([
            { key: "daily" as const, label: lang === "th" ? "บันทึกรายวัน" : lang === "lo" ? "ບັນທຶກປະຈຳວັນ" : "Daily Records", icon: Recycle },
            { key: "import" as const, label: lang === "th" ? "นำเข้าคะแนนสะสม" : lang === "lo" ? "ນຳເຂົ້າຄະແນນສະສົມ" : "Import Scores", icon: Upload },
            { key: "search" as const, label: lang === "th" ? "ค้นหาคะแนนสะสม" : lang === "lo" ? "ຄົ້ນຫາຄະແນນສະສົມ" : "Search Scores", icon: Database },
            ...(currentUser.role === "admin" ? [{ key: "admin" as const, label: lang === "th" ? "จัดการผู้ใช้งาน" : lang === "lo" ? "ຈັດການຜູ້ໃຊ້" : "Manage Users", icon: Shield }] : []),
          ] as { key: "daily" | "import" | "search" | "admin"; label: string; icon: typeof Recycle }[]).map((tab) => {
            const Icon = tab.icon;
            return (
              <button
                key={tab.key}
                onClick={() => setActiveTab(tab.key)}
                className={`flex items-center gap-1.5 border-b-2 px-3 py-3 text-sm font-medium transition-colors ${
                  activeTab === tab.key
                    ? "border-primary text-primary"
                    : "border-transparent text-muted-foreground hover:text-foreground"
                }`}
              >
                <Icon className="h-4 w-4" />
                {tab.label}
              </button>
            );
          })}
        </div>
      </nav>

      {/* Mobile Accordion Navigation */}
      <nav className="w-full overflow-x-hidden border-b bg-card/95 backdrop-blur lg:hidden">
        {([
          { key: "daily" as const, label: lang === "th" ? "บันทึกรายวัน" : lang === "lo" ? "ບັນທຶກປະຈຳວັນ" : "Daily Records", icon: Recycle },
          { key: "import" as const, label: lang === "th" ? "นำเข้าคะแนนสะสม" : lang === "lo" ? "ນຳເຂົ້າຄະແນນສະສົມ" : "Import Scores", icon: Upload },
          { key: "search" as const, label: lang === "th" ? "ค้นหาคะแนนสะสม" : lang === "lo" ? "ຄົ້ນຫາຄະແນນສະສົມ" : "Search Scores", icon: Database },
          ...(currentUser.role === "admin" ? [{ key: "admin" as const, label: lang === "th" ? "จัดการผู้ใช้งาน" : lang === "lo" ? "ຈັດການຜູ້ໃຊ້" : "Manage Users", icon: Shield }] : []),
        ] as { key: "daily" | "import" | "search" | "admin"; label: string; icon: typeof Recycle }[]).map((tab) => {
          const Icon = tab.icon;
          const isActive = activeTab === tab.key;
          if (!mobileNavOpen && !isActive) return null;
          return (
            <button
              key={tab.key}
              onClick={() => {
                if (isActive) setMobileNavOpen((open) => !open);
                else {
                  setActiveTab(tab.key);
                  setMobileNavOpen(false);
                }
              }}
              className={`flex w-full items-center justify-between border-b px-3 py-2.5 text-sm font-medium transition-colors last:border-b-0 sm:px-4 sm:py-3 ${
                isActive
                  ? "bg-primary/5 text-primary"
                  : "text-muted-foreground hover:text-foreground"
              }`}
            >
              <span className="flex min-w-0 items-center gap-2 text-left">
                <Icon className="h-4 w-4 shrink-0" />
                <span className="break-words">{tab.label}</span>
              </span>
              <ChevronDown
                className={`h-4 w-4 shrink-0 transition-transform duration-200 ${
                  isActive && mobileNavOpen ? "rotate-180" : ""
                }`}
              />
            </button>
          );
        })}
      </nav>

      {showBackups && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4" onClick={() => setShowBackups(false)}>
          <div className="max-h-[80vh] w-full max-w-md overflow-y-auto rounded-lg bg-card p-4 shadow-xl" onClick={(e) => e.stopPropagation()}>
            <h3 className="mb-3 text-base font-semibold">{lang === "th" ? "เวอร์ชันสำรองล่าสุด" : lang === "lo" ? "ເວີຊັນສຳຮອງຫຼ້າສຸດ" : "Recent Backup Versions"}</h3>
            {localBackups.length === 0 ? (
              <p className="text-sm text-muted-foreground">{lang === "th" ? "ยังไม่มีเวอร์ชันสำรอง" : lang === "lo" ? "ຍັງບໍ່ມີເວີຊັນສຳຮອງ" : "No backup versions yet"}</p>
            ) : (
              <div className="space-y-2">
                {localBackups.map((b, i) => (
                  <div key={i} className="flex items-center justify-between rounded-md border p-2.5">
                    <div className="min-w-0">
                      <p className="text-sm font-medium">{new Date(b.timestamp).toLocaleString()}</p>
                      <p className="text-xs text-muted-foreground">{b.entries} {lang === "th" ? "รายการ" : lang === "lo" ? "ລາຍການ" : "entries"}</p>
                    </div>
                    <Button
                      variant="outline"
                      size="sm"
                      disabled={backupBusy}
                      onClick={async () => {
                        if (!confirm(lang === "th" ? "กู้คืนเวอร์ชันนี้?" : lang === "lo" ? "ກູ້ຄືນເວີຊັນນີ້?" : "Restore this version?")) return;
                        setBackupBusy(true);
                        try {
                          await restoreLocalBackup(i);
                          toast.success(lang === "th" ? "กู้คืนสำเร็จ" : lang === "lo" ? "ກູ້ຄືນສຳເລັດ" : "Restored");
                          window.location.reload();
                        } catch (e) {
                          toast.error((e as Error).message);
                          setBackupBusy(false);
                        }
                      }}
                    >
                      <ArchiveRestore className="mr-1 h-3.5 w-3.5" /> {lang === "th" ? "กู้คืน" : lang === "lo" ? "ກູ້ຄືນ" : "Restore"}
                    </Button>
                  </div>
                ))}
              </div>
            )}
            <Button variant="outline" size="sm" className="mt-3 w-full" onClick={() => setShowBackups(false)}>
              {lang === "th" ? "ปิด" : lang === "lo" ? "ປິດ" : "Close"}
            </Button>
          </div>
        </div>
      )}

      <main className="mx-auto w-full max-w-7xl min-w-0 space-y-4 overflow-x-hidden px-2 py-4 sm:space-y-6 sm:px-6 sm:py-6">
        {activeTab === "import" && <ImportScores lang={lang} onImported={loadEmpMaster} />}
        {activeTab === "search" && <SearchScores lang={lang} currentUser={currentUser} />}
        {activeTab === "admin" && currentUser.role === "admin" && <AdminUsers currentUserUsername={currentUser.username} />}
        {activeTab === "daily" && (
        <>
        <Card className="p-3 sm:p-4 lg:p-6">
          <h2 className="mb-3 text-base font-semibold sm:mb-4 sm:text-lg">{t("dailyEntry", lang)}</h2>
          <div className="grid min-w-0 grid-cols-1 gap-2 min-[360px]:grid-cols-2 sm:grid-cols-3 sm:gap-3 lg:grid-cols-4">
            <Field label={t("date", lang)}>
              <Input
                type="date"
                value={form.date}
                onChange={(e) => setForm({ ...form, date: e.target.value })}
              />
            </Field>
            <Field label={t("code", lang)}>
              <div className="relative">
                <Input
                  value={form.code}
                  onChange={(e) => {
                    const val = e.target.value;
                    setForm({ ...form, code: val });
                    if (val.trim().length >= 1) {
                      const q = val.toLowerCase().trim();
                      const matches = empMaster.filter((m) => m.employee_code.toLowerCase().includes(q)).slice(0, 10);
                      setCodeSuggest(matches);
                      setShowCodeList(matches.length > 0);
                    } else {
                      setShowCodeList(false);
                    }
                  }}
                  onFocus={() => { if (form.code.trim().length >= 1) setShowCodeList(codeSuggest.length > 0); }}
                  onBlur={() => setTimeout(() => setShowCodeList(false), 200)}
                  placeholder={empLoading ? (lang === "th" ? "กำลังโหลด..." : lang === "lo" ? "ກຳລັງໂຫລດ..." : "Loading...") : ""}
                />
                {showCodeList && (
                  <div className="absolute z-20 mt-1 w-full rounded-md border bg-popover shadow-md">
                    {codeSuggest.map((m, i) => (
                      <button
                        key={i}
                        type="button"
                        className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm hover:bg-accent"
                        onMouseDown={(e) => {
                          e.preventDefault();
                          setForm({ ...form, code: m.employee_code, name: m.employee_name, unit_org: m.department ?? "" });
                          setShowCodeList(false);
                          setShowNameList(false);
                        }}
                      >
                        <span className="font-mono text-xs font-semibold text-primary">{m.employee_code}</span>
                        <span className="min-w-0 flex-1 truncate">{m.employee_name}</span>
                        {m.department && <span className="text-xs text-muted-foreground">{m.department}</span>}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            </Field>
            <Field label={t("nameRequired", lang)}>
              <div className="relative">
                <Input
                  value={form.name}
                  onChange={(e) => {
                    const val = e.target.value;
                    setForm({ ...form, name: val });
                    if (val.trim().length >= 1) {
                      const q = val.toLowerCase().trim();
                      const matches = empMaster.filter((m) => m.employee_name.toLowerCase().includes(q));
                      setNameSuggest(matches);
                      setShowNameList(matches.length > 0);
                    } else {
                      setShowNameList(false);
                    }
                  }}
                  onFocus={() => { if (form.name.trim().length >= 1) setShowNameList(nameSuggest.length > 0); }}
                  onBlur={() => setTimeout(() => setShowNameList(false), 200)}
                  placeholder={empMaster.length > 0 ? (lang === "th" ? "พิมพ์ชื่อเพื่อค้นหา..." : lang === "lo" ? "ພິມຊື່ເພື່ອຄົ້ນຫາ..." : "Type name to search...") : ""}
                />
                {showNameList && (
                  <div className="absolute z-20 mt-1 max-h-64 w-full overflow-y-auto rounded-md border bg-popover shadow-md">
                    {nameSuggest.map((m, i) => (
                      <button
                        key={`${m.employee_code}-${i}`}
                        type="button"
                        className="flex w-full flex-col gap-0.5 px-3 py-2 text-left text-sm hover:bg-accent sm:flex-row sm:items-center sm:gap-2"
                        onMouseDown={(e) => {
                          e.preventDefault();
                          setForm({ ...form, code: m.employee_code, name: m.employee_name, unit_org: m.department ?? "" });
                          setShowNameList(false);
                          setShowCodeList(false);
                        }}
                      >
                        <span className="font-medium">{m.employee_name}</span>
                        <span className="flex min-w-0 items-center gap-2 text-xs">
                          <span className="font-mono font-semibold text-primary">{m.employee_code}</span>
                          {m.department && <span className="truncate text-muted-foreground">{m.department}</span>}
                        </span>
                      </button>
                    ))}
                  </div>
                )}
              </div>
            </Field>
            <Field label={t("org", lang)}>
              <Input
                value={form.unit_org}
                onChange={(e) => setForm({ ...form, unit_org: e.target.value })}
              />
            </Field>
            {empMaster.length > 0 && (
              <div className="flex items-center gap-1.5 text-xs text-muted-foreground">
                <UserSearch className="h-3.5 w-3.5" />
                {empMaster.length} {lang === "th" ? "พนักงานพร้อมค้นหา" : lang === "lo" ? "ພະນັກງານພ້ອມຄົ້ນຫາ" : "employees ready"}
              </div>
            )}
          </div>

          <div className="mt-6">
            <h3 className="mb-2 break-words text-sm font-semibold">{t("matrixHint", lang)}</h3>
            <div className="hidden overflow-x-auto rounded-lg border md:block">
              <table className="w-full border-collapse text-sm">
                <thead className="bg-muted">
                  <tr>
                    <th className="border px-2 py-2 text-left font-semibold">{t("wasteType", lang)}</th>
                    <th className="border px-2 py-2 text-center font-medium">{t("criteria", lang)}</th>
                    {ROOMS_L.map((r, i) => (
                      <th key={ROOMS[i]} className="border px-2 py-2 text-center font-medium">
                        {tr(r, lang)}
                      </th>
                    ))}
                    <th className="border bg-primary/10 px-2 py-2 text-center font-semibold text-primary">
                      {t("totalRoom", lang)}
                    </th>
                    <th className="border bg-primary/10 px-2 py-2 text-center font-semibold text-primary">
                      {t("points", lang)}
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {WASTE_TYPES.map((w, i) => {
                    const tt = wasteTotals.get(w.key)!;
                    const cat = categoryOfIndex(i);
                    const showHeader = cat && i === cat.start;
                    return (
                      <Fragment key={w.key}>
                        {showHeader && cat && (
                          <tr className="bg-accent/40">
                            <td className="border px-2 py-2 font-bold" colSpan={ROOMS.length + 4}>
                              <span className="text-sm">{cat.icon} {tr(cat.title, lang)}</span>
                              <span className="ml-2 text-xs text-muted-foreground">
                                ({cat.start + 1}-{cat.end + 1})
                              </span>
                            </td>
                          </tr>
                        )}
                        <tr className="hover:bg-muted/30">
                          <td className="border px-2 py-1.5 font-medium">
                            <span className="text-xs text-muted-foreground">{i + 1}.</span>{" "}
                            {tr(w.label, lang)}
                          </td>
                          <td className="border px-2 py-1.5 text-center text-xs text-muted-foreground">
                            {criteriaText(w, lang)}
                          </td>
                          {ROOMS.map((r) => (
                            <td key={r} className="border px-1 py-1">
                              <Input
                                type="number"
                                inputMode="numeric"
                                min="0"
                                value={form.matrix[w.key][r]}
                                onChange={(e) => setCell(w.key, r, e.target.value)}
                                className="h-9 text-center"
                                placeholder="0"
                              />
                            </td>
                          ))}
                          <td className="border bg-primary/5 px-2 py-1.5 text-center font-semibold">
                            {tt.sum}
                          </td>
                          <td className="border bg-primary/5 px-2 py-1.5 text-center font-bold text-primary">
                            {fmtPts(tt.points)}
                          </td>
                        </tr>
                      </Fragment>
                    );
                  })}
                  <tr className="bg-muted font-semibold">
                    <td className="border px-2 py-2" colSpan={2}>
                      {t("perRoomQty", lang)}
                    </td>
                    {ROOMS.map((r) => (
                      <td key={r} className="border px-2 py-2 text-center">
                        {roomTotals.get(r)?.qty || 0}
                      </td>
                    ))}
                    <td className="border bg-primary/10 px-2 py-2 text-center">{grandQty}</td>
                    <td className="border bg-primary/10 px-2 py-2 text-center text-primary">
                      {fmtPts(grandPoints)}
                    </td>
                  </tr>
                  <tr className="bg-muted font-semibold">
                    <td className="border px-2 py-2" colSpan={2}>
                      {t("perRoomPts", lang)}
                    </td>
                    {ROOMS.map((r) => (
                      <td key={r} className="border px-2 py-2 text-center text-primary">
                        {fmtPts(roomTotals.get(r)?.points || 0)}
                      </td>
                    ))}
                    <td className="border bg-primary/10 px-2 py-2 text-center" colSpan={2}>
                      {t("grandTotalPts", lang)} {fmtPts(grandPoints)} {t("pointsWord", lang)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            {/* Mobile accordion input */}
            <div className="mt-3 space-y-2.5 sm:mt-4 sm:space-y-3 md:hidden">
              {MOBILE_CATS.map((cat, ci) => {
                const isOpen = ci === 0 || openCats.has(ci);
                const catWastes = cat.keys.map((k) => WASTE_TYPES.find((w) => w.key === k)!);
                return (
                  <div key={cat.icon + ci} className="overflow-hidden rounded-lg border border-border bg-card">
                    <button
                      onClick={() => ci !== 0 && toggleCat(ci)}
                      className="flex w-full items-center justify-between bg-accent/40 px-2.5 py-2 sm:px-3 sm:py-2.5"
                    >
                      <span className="text-xs font-bold sm:text-sm">
                        {cat.icon} {tr(cat.title, lang)}
                      </span>
                      <span className="flex items-center gap-2">
                        <span className="text-xs font-semibold text-primary">
                          {fmtPts(catTotals[ci].points)} {t("pointsWord", lang)}
                        </span>
                        {ci !== 0 && (
                          <ChevronDown
                            className={`h-4 w-4 text-muted-foreground transition-transform ${isOpen ? "rotate-180" : ""}`}
                          />
                        )}
                      </span>
                    </button>
                    {isOpen && (
                      <div className="divide-y divide-border">
                        {catWastes.map((w) => {
                          const idx = WASTE_TYPES.indexOf(w);
                          const tt = wasteTotals.get(w.key)!;
                          return (
                            <div key={w.key} className="p-2.5 sm:p-3">
                              <div className="mb-1.5 flex items-start justify-between gap-2 sm:mb-2">
                                <div className="flex-1">
                                  <p className="text-sm font-medium">
                                    <span className="text-xs text-muted-foreground">{idx + 1}.</span>{" "}
                                    {tr(w.label, lang)}
                                  </p>
                                  <p className="text-xs text-muted-foreground">{criteriaText(w, lang)}</p>
                                </div>
                                <div className="text-right">
                                  <p className="text-xs text-muted-foreground">{t("points", lang)}</p>
                                  <p className="font-bold text-primary">{fmtPts(tt.points)}</p>
                                </div>
                              </div>
                              <div className="grid min-w-0 grid-cols-2 gap-1 min-[360px]:grid-cols-3 sm:grid-cols-5 sm:gap-1.5">
                                {ROOMS.map((r, ri) => (
                                  <div key={r}>
                                    <label className="mb-0.5 block text-center text-[10px] text-muted-foreground">
                                      {tr(ROOMS_L[ri], lang)}
                                    </label>
                                    <Input
                                      type="number"
                                      inputMode="numeric"
                                      min="0"
                                      value={form.matrix[w.key][r]}
                                      onChange={(e) => setCell(w.key, r, e.target.value)}
                                      className="h-8 px-0.5 text-center text-xs sm:h-9 sm:px-1 sm:text-sm"
                                      placeholder="0"
                                    />
                                  </div>
                                ))}
                              </div>
                              <div className="mt-1.5 flex min-w-0 items-center justify-between gap-2 text-xs">
                                <span className="text-muted-foreground">{t("totalRoom", lang)}:</span>
                                <span className="font-semibold">{tt.sum} {tr(w.unit, lang)}</span>
                              </div>
                            </div>
                          );
                        })}
                        <div className="flex items-center justify-between bg-primary/5 px-3 py-2 text-sm">
                          <span className="font-semibold">{t("grandTotalPts", lang)}</span>
                          <span className="font-bold text-primary">
                            {fmtPts(catTotals[ci].points)} {t("pointsWord", lang)}
                          </span>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            {/* Mobile total summary */}
            <div className="mt-3 space-y-2.5 sm:mt-4 sm:space-y-3 md:hidden">
              <div className="rounded-lg border border-primary bg-primary p-3 text-primary-foreground shadow-md sm:p-4">
                <p className="text-xs opacity-80">{t("totalPointsAll", lang)}</p>
                <p className="mt-1 text-2xl font-bold sm:text-3xl">{fmtPts(grandPoints)}</p>
                <p className="text-xs opacity-80">{t("pointsWord", lang)}</p>
              </div>
            </div>
          </div>

          <div className="mt-3 flex min-w-0 flex-wrap gap-1.5 sm:mt-4 sm:gap-2">
            <Button onClick={addRow} disabled={syncing}>
              <Plus className="mr-1 h-4 w-4" /> {syncing ? t("syncing", lang) : t("saveEntry", lang)}
            </Button>

            <Button variant="outline" onClick={() => setForm({ ...form, matrix: emptyMatrix() })}>
              {t("clearTable", lang)}
            </Button>
            <Button variant="outline" onClick={() => setForm(newForm())}>
              {t("clearForm", lang)}
            </Button>
          </div>
        </Card>

        <div className="mb-2 flex items-center justify-end gap-2 text-sm">
          <span className="text-muted-foreground">{t("date", lang)}</span>
          <Input
            type="date"
            value={selectedDate}
            onChange={(e) => setSelectedDate(e.target.value)}
            className="h-9 w-40"
          />
        </div>
        <div className="grid gap-2 sm:gap-3 md:grid-cols-4">
          <Card className="border-primary bg-primary p-3 text-primary-foreground shadow-md sm:p-4">
            <p className="text-xs opacity-80">{t("totalPointsAll", lang)}</p>
            <p className="mt-1 text-2xl font-bold sm:text-3xl">{fmtPts(totalPoints)}</p>
            <p className="text-xs opacity-80">{t("pointsWord", lang)}</p>
          </Card>
          {REWARDS.map((rw) => {
            const exchangeable = Math.floor(totalPoints / rw.cost);
            return (
              <Card
                key={rw.label.en}
                className={`p-3 sm:p-4 ${exchangeable > 0 ? "border-primary bg-primary/5" : ""}`}
              >
                <div className="flex items-center gap-2 text-sm font-semibold">
                  <Gift className="h-4 w-4 text-primary" /> {tr(rw.label, lang)}
                </div>
                <p className="mt-1 text-xs text-muted-foreground">
                  {rw.cost} {t("pointsWord", lang)} / 1 {tr(rw.unit, lang)}
                </p>
                <p className="mt-2 text-xl font-bold sm:text-2xl">
                  {exchangeable}{" "}
                  <span className="text-sm font-normal text-muted-foreground">{tr(rw.unit, lang)}</span>
                </p>
                <p className="text-xs text-muted-foreground">{t("redeemNow", lang)}</p>
              </Card>
            );
          })}
        </div>

        {(byPerson.length > 0 || debouncedPersonSearch.trim()) && (
          <Card className="p-3 sm:p-4 lg:p-6">
            <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-base font-semibold sm:text-lg">{t("perPersonSummary", lang)}</h2>
              <div className="flex items-center gap-2">
                <div className="relative">
                  <Search className="absolute left-2 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                  <Input
                    type="text"
                    placeholder={t("searchPlaceholder", lang)}
                    value={personSearchText}
                    onChange={(e) => setPersonSearchText(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter") setDebouncedPersonSearch(personSearchText);
                    }}
                    className="h-9 w-48 pl-8 sm:w-64"
                  />
                </div>
                {debouncedPersonSearch.trim() && (
                  <Button variant="ghost" size="sm" onClick={() => { setPersonSearchText(""); setDebouncedPersonSearch(""); }}>
                    {t("clearSearch", lang)}
                  </Button>
                )}
              </div>
            </div>
            {debouncedPersonSearch.trim() && (
              <p className="mb-2 text-xs text-muted-foreground">
                {t("searchResults", lang).replace("{n}", String(filteredByPerson.length))}
              </p>
            )}
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead className="bg-muted">
                  <tr>
                    <th className="border px-2 py-2 text-left">{t("name", lang)}</th>
                    <th className="border px-2 py-2 text-left">{t("org", lang)}</th>
                    <th className="border px-2 py-2 text-right">{t("accumulated", lang)}</th>
                    <th className="border px-2 py-2 text-right">คะแนนสะสมรวม</th>
                    <th className="border px-2 py-2 text-right">คะแนนคงเหลือ</th>
                    <th className="border px-2 py-2 text-center">{t("status", lang)}</th>
                    <th className="border px-2 py-2 text-center">แลกของ</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredByPerson.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="py-6 text-center text-muted-foreground">
                        {t("noResults", lang)}
                      </td>
                    </tr>
                  ) : (
                    filteredByPerson.map((p, i) => {
                      const sc = getPersonScore(p);
                      const remaining = sc ? sc.remaining : 0;
                      const isExpanded = expandedPerson === personKey(p);
                      return (
                        <Fragment key={i}>
                          <tr className={`border-b hover:bg-muted/30 ${isExpanded ? "bg-primary/5" : ""}`}>
                            <td className="border px-2 py-2">
                              <Highlight text={p.name} keyword={debouncedPersonSearch} />
                            </td>
                            <td className="border px-2 py-2 text-muted-foreground">
                              <Highlight text={p.unit_org || "-"} keyword={debouncedPersonSearch} />
                            </td>
                            <td className="border px-2 py-2 text-right font-semibold">{fmtPts(p.points)}</td>
                            <td className="border px-2 py-2 text-right font-semibold text-muted-foreground">
                              {sc ? sc.total : "-"}
                            </td>
                            <td className="border px-2 py-2 text-right font-bold text-primary">
                              {sc ? remaining : "-"}
                            </td>
                            <td className="border px-2 py-2 text-center">
                              {p.points >= REWARD_THRESHOLD ? (
                                <span className="inline-flex items-center gap-1 rounded-full bg-primary/10 px-2 py-0.5 text-xs font-medium text-primary">
                                  <Gift className="h-3 w-3" /> {t("canRedeem", lang)}
                                </span>
                              ) : (
                                <span className="text-xs text-muted-foreground">
                                  {t("needMore", lang)} {fmtPts(REWARD_THRESHOLD - p.points)} {t("more", lang)}
                                </span>
                              )}
                            </td>
                            <td className="border px-2 py-2 text-center">
                              <Button
                                size="sm"
                                variant={isExpanded ? "default" : "outline"}
                                onClick={() => setExpandedPerson(isExpanded ? null : personKey(p))}
                                className="h-7"
                              >
                                <Gift className="mr-1 h-3.5 w-3.5" />
                                {isExpanded ? "ปิด" : "แลก"}
                              </Button>
                            </td>
                          </tr>
                          {isExpanded && (
                            <tr>
                              <td colSpan={7} className="border bg-accent/5 p-3">
                                <RewardExchangePanel
                                  lang={lang}
                                  currentUser={currentUser}
                                  personName={p.name}
                                  personCode={p.code}
                                  onExchanged={() => { loadPersonScores(); loadEmpMaster(); }}
                                />
                              </td>
                            </tr>
                          )}
                        </Fragment>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </Card>
        )}

        <Card className="p-3 sm:p-4 lg:p-6">
          <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
            <h2 className="text-base font-semibold sm:text-lg">{t("dailySummary", lang)}</h2>
            <div className="flex flex-wrap items-center gap-2">
              <label className="text-xs text-muted-foreground">{t("date", lang)}:</label>
              <Input
                type="date"
                value={summaryDate}
                onChange={(e) => setSummaryDate(e.target.value)}
                className="h-9 w-auto"
              />
              <span className="text-xs text-muted-foreground">{t("saveNote", lang)}</span>
            </div>
          </div>
          {dailySummary.count === 0 ? (
            <p className="py-6 text-center text-sm text-muted-foreground">{t("noDataDay", lang)}</p>
          ) : (
            <>
            <div className="hidden overflow-x-auto rounded-lg border md:block">
              <table className="w-full border-collapse text-sm">
                <thead className="bg-muted">
                  <tr>
                    <th className="border px-2 py-2 text-center w-12">{t("rank", lang)}</th>
                    <th className="border px-2 py-2 text-left">{t("wasteType", lang)}</th>
                    <th className="border px-2 py-2 text-center">{t("criteria", lang)}</th>
                    <th className="border px-2 py-2 text-right">{t("qty", lang)}</th>
                    <th className="border px-2 py-2 text-center">{t("unit", lang)}</th>
                    <th className="border px-2 py-2 text-right">{t("points", lang)}</th>
                  </tr>
                </thead>
                <tbody>
                  {WASTE_TYPES.map((w, i) => {
                    const tt = dailySummary.byType.get(w.key)!;
                    const hasData = tt.qty > 0;
                    const cat = categoryOfIndex(i);
                    const showHeader = cat && i === cat.start;
                    return (
                      <Fragment key={w.key}>
                        {showHeader && cat && (
                          <tr className="bg-accent/40">
                            <td className="border px-2 py-2 font-bold" colSpan={6}>
                              <span className="text-sm">{cat.icon} {tr(cat.title, lang)}</span>
                              <span className="ml-2 text-xs text-muted-foreground">
                                ({cat.start + 1}-{cat.end + 1})
                              </span>
                            </td>
                          </tr>
                        )}
                        <tr className={hasData ? "" : "text-muted-foreground"}>
                          <td className="border px-2 py-1.5 text-center font-semibold">{i + 1}</td>
                          <td className="border px-2 py-1.5">{tr(w.label, lang)}</td>
                          <td className="border px-2 py-1.5 text-center text-xs">
                            {criteriaText(w, lang)}
                          </td>
                          <td className="border px-2 py-1.5 text-right font-semibold">
                            {fmtPts(tt.qty)}
                          </td>
                          <td className="border px-2 py-1.5 text-center">{tr(w.unit, lang)}</td>
                          <td className="border px-2 py-1.5 text-right font-bold text-primary">
                            {fmtPts(tt.points)}
                          </td>
                        </tr>
                      </Fragment>
                    );
                  })}
                  <tr className="bg-primary/10 font-bold">
                    <td className="border px-2 py-2 text-center" colSpan={3}>
                      {t("grandTotalPts", lang)}
                    </td>
                    <td className="border px-2 py-2 text-right">{fmtPts(dailySummary.totalQty)}</td>
                    <td className="border px-2 py-2 text-center">-</td>
                    <td className="border px-2 py-2 text-right text-primary">
                      {fmtPts(dailySummary.totalPts)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            {/* Mobile card-based daily summary */}
            <div className="space-y-4 md:hidden">
              {WASTE_CATEGORIES.map((cat) => (
                <div key={cat.icon} className="rounded-lg border border-border bg-card">
                  <div className="rounded-t-lg bg-accent/40 px-3 py-2">
                    <p className="text-sm font-bold">
                      {cat.icon} {tr(cat.title, lang)}
                    </p>
                    <p className="text-xs text-muted-foreground">
                      {cat.start + 1}-{cat.end + 1}
                    </p>
                  </div>
                  <div className="divide-y divide-border">
                    {WASTE_TYPES.slice(cat.start, cat.end + 1).map((w) => {
                      const idx = WASTE_TYPES.indexOf(w);
                      const tt = dailySummary.byType.get(w.key)!;
                      const hasData = tt.qty > 0;
                      return (
                        <div key={w.key} className={`p-3 ${!hasData ? "opacity-50" : ""}`}>
                          <div className="flex items-start justify-between gap-2">
                            <div className="flex-1">
                              <p className="text-sm font-medium">
                                <span className="text-xs text-muted-foreground">{idx + 1}.</span>{" "}
                                {tr(w.label, lang)}
                              </p>
                              <p className="text-xs text-muted-foreground">{criteriaText(w, lang)}</p>
                            </div>
                            <div className="text-right">
                              <p className="text-xs text-muted-foreground">{t("qty", lang)}</p>
                              <p className="text-sm font-semibold">
                                {fmtPts(tt.qty)} {tr(w.unit, lang)}
                              </p>
                            </div>
                            <div className="text-right">
                              <p className="text-xs text-muted-foreground">{t("points", lang)}</p>
                              <p className="text-sm font-bold text-primary">{fmtPts(tt.points)}</p>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ))}
              <div className="rounded-lg border border-primary/30 bg-primary/10 p-3">
                <div className="flex items-center justify-between">
                  <p className="text-sm font-bold">{t("grandTotalPts", lang)}</p>
                  <div className="flex items-center gap-3">
                    <span className="text-sm font-semibold">
                      {fmtPts(dailySummary.totalQty)} {t("qty", lang)}
                    </span>
                    <span className="text-lg font-bold text-primary">
                      {fmtPts(dailySummary.totalPts)} {t("pointsWord", lang)}
                    </span>
                  </div>
                </div>
              </div>
            </div>
            </>
          )}
        </Card>

        <Card className="p-3 sm:p-4 lg:p-6">
          <div className="mb-3 flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
            <h2 className="text-base font-semibold sm:text-lg">
              {t("allEntries", lang)} ({rows.length})
            </h2>
            <div className="flex flex-col gap-2 sm:items-end">
              <div className="flex items-center gap-2">
                <div className="relative">
                  <Search className="absolute left-2.5 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
                  <Input
                    value={searchText}
                    onChange={(e) => setSearchText(e.target.value)}
                    onKeyDown={(e) => { if (e.key === "Enter") setDebouncedSearch(searchText); }}
                    placeholder={t("searchPlaceholder", lang)}
                    className="h-9 w-56 pl-9 pr-8 text-sm"
                  />
                  {searchText && (
                    <button
                      onClick={() => setSearchText("")}
                      className="absolute right-2 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground"
                      aria-label={t("clearSearch", lang)}
                    >
                      <X className="h-4 w-4" />
                    </button>
                  )}
                </div>
                <Button
                  size="sm"
                  onClick={() => setDebouncedSearch(searchText)}
                  aria-label="Search"
                >
                  <Search className="h-4 w-4" />
                </Button>
                {rows.length > 0 && (
                  <Button variant="outline" size="sm" onClick={clearAll}>
                    {t("clearAll", lang)}
                  </Button>
                )}
              </div>
              {debouncedSearch.trim() && (
                <span className="text-xs text-muted-foreground">
                  {t("searchResults", lang).replace("{n}", String(filteredRows.length))}
                </span>
              )}
            </div>
          </div>
          {rows.length === 0 ? (
            <p className="py-8 text-center text-sm text-muted-foreground">{t("noData", lang)}</p>
          ) : filteredRows.length === 0 ? (
            <p className="py-8 text-center text-sm text-muted-foreground">{t("noResults", lang)}</p>
          ) : (
            <div className="overflow-x-auto rounded-lg border">
              <table className="w-full border-collapse text-xs sm:text-sm">
                <thead className="bg-muted">
                  <tr>
                    {[
                      t("date", lang),
                      t("code", lang),
                      t("name", lang),
                      t("org", lang),
                      "Room",
                      t("wasteType", lang),
                      t("qty", lang),
                      t("unit", lang),
                      t("points", lang),
                      "",
                    ].map((h, i) => (
                      <th key={i} className="border px-2 py-2 text-left font-semibold">
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredRows.map((r) => {
                    const isEditing = editingId === r.id && editDraft;
                    if (isEditing && editDraft) {
                      const liveQty = parseFloat(editDraft.quantity) || 0;
                      const livePts = fmtPts(calcPoints(editDraft.wasteType, liveQty));
                      return (
                        <tr key={r.id} className="border-b bg-accent/30">
                          <td className="border px-1 py-1">
                            <Input
                              type="date"
                              value={editDraft.date}
                              onChange={(e) => setEditDraft({ ...editDraft, date: e.target.value })}
                              className="h-8"
                            />
                          </td>
                          <td className="border px-1 py-1">
                            <Input
                              value={editDraft.code}
                              onChange={(e) => setEditDraft({ ...editDraft, code: e.target.value })}
                              className="h-8"
                            />
                          </td>
                          <td className="border px-1 py-1">
                            <Input
                              value={editDraft.name}
                              onChange={(e) => setEditDraft({ ...editDraft, name: e.target.value })}
                              className="h-8"
                            />
                          </td>
                          <td className="border px-1 py-1">
                            <Input
                              value={editDraft.unit_org}
                              onChange={(e) => setEditDraft({ ...editDraft, unit_org: e.target.value })}
                              className="h-8"
                            />
                          </td>
                          <td className="border px-1 py-1">
                            <Input
                              value={editDraft.room}
                              onChange={(e) => setEditDraft({ ...editDraft, room: e.target.value })}
                              className="h-8"
                            />
                          </td>
                          <td className="border px-1 py-1">
                            <select
                              value={editDraft.wasteType}
                              onChange={(e) => {
                                const wt = e.target.value;
                                const w = WASTE_TYPES.find((x) => x.key === wt);
                                const unit = w ? tr(w.unit, lang) : editDraft.qtyUnit;
                                setEditDraft({ ...editDraft, wasteType: wt, qtyUnit: unit });
                              }}
                              className="h-8 w-full rounded border bg-background px-1 text-xs"
                            >
                              {WASTE_TYPES.map((w) => (
                                <option key={w.key} value={w.key}>
                                  {tr(w.label, lang)}
                                </option>
                              ))}
                            </select>
                          </td>
                          <td className="border px-1 py-1">
                            <Input
                              type="number"
                              value={editDraft.quantity}
                              onChange={(e) => setEditDraft({ ...editDraft, quantity: e.target.value })}
                              className="h-8 text-right"
                            />
                          </td>
                          <td className="border px-1 py-1">
                            <Input
                              value={editDraft.qtyUnit}
                              onChange={(e) => setEditDraft({ ...editDraft, qtyUnit: e.target.value })}
                              className="h-8"
                            />
                          </td>
                          <td className="border px-2 py-1.5 text-right font-semibold text-primary">
                            {livePts}
                          </td>
                          <td className="border px-1 py-1.5">
                            <div className="flex gap-1">
                              <button
                                onClick={saveEdit}
                                className="flex h-7 w-7 items-center justify-center rounded bg-primary text-primary-foreground hover:bg-primary/90"
                                aria-label="save"
                                title="save"
                              >
                                <Check className="h-3.5 w-3.5" />
                              </button>
                              <button
                                onClick={cancelEdit}
                                className="flex h-7 w-7 items-center justify-center rounded border text-muted-foreground hover:bg-muted"
                                aria-label="cancel"
                                title="cancel"
                              >
                                <X className="h-3.5 w-3.5" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    }
                    return (
                      <tr
                        key={r.id}
                        className={`border-b hover:bg-muted/40 ${
                          r.room === TOTAL_ROOM ? "bg-primary/5 font-medium" : ""
                        }`}
                      >
                        <td className="border px-2 py-1.5">{r.date}</td>
                        <td className="border px-2 py-1.5"><Highlight text={r.code} keyword={debouncedSearch} /></td>
                        <td className="border px-2 py-1.5"><Highlight text={r.name} keyword={debouncedSearch} /></td>
                        <td className="border px-2 py-1.5"><Highlight text={r.unit_org} keyword={debouncedSearch} /></td>
                        <td className="border px-2 py-1.5">{r.room}</td>
                        <td className="border px-2 py-1.5">{wasteLabel(r.wasteType)}</td>
                        <td className="border px-2 py-1.5 text-right">{r.quantity}</td>
                        <td className="border px-2 py-1.5">{r.qtyUnit}</td>
                        <td className="border px-2 py-1.5 text-right font-semibold">{r.points}</td>
                        <td className="border px-1 py-1.5">
                          <div className="flex gap-1">
                            <button
                              onClick={() => startEdit(r)}
                              className="flex h-7 w-7 items-center justify-center rounded text-muted-foreground hover:bg-primary hover:text-primary-foreground"
                              aria-label="edit"
                              title="edit"
                            >
                              <Pencil className="h-3.5 w-3.5" />
                            </button>
                            <button
                              onClick={() => removeRow(r.id)}
                              className="flex h-7 w-7 items-center justify-center rounded text-muted-foreground hover:bg-destructive hover:text-destructive-foreground"
                              aria-label="delete"
                              title="delete"
                            >
                              <Trash2 className="h-3.5 w-3.5" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </Card>

        </>
        )}
      </main>
    </div>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="min-w-0 space-y-1">
      <label className="text-xs font-medium text-muted-foreground">{label}</label>
      {children}
    </div>
  );
}
