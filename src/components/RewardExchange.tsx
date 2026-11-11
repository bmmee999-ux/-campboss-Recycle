import { useEffect, useState } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Gift, Loader2, History, Egg, Soup, Fish, ChevronDown, Bed, BedDouble, Droplet, CreditCard, Package } from "lucide-react";
import { toast } from "sonner";
import { supabase } from "@/integrations/supabase/client";
import type { Lang } from "@/components/AccumulatedScores";
import type { AppUser } from "@/lib/auth";

type L = { th: string; lo: string; en: string };
const tr = (l: L, lang: Lang) => l[lang];

// Order, names, costs and units follow Easy_To_Read_Waste_Summary.xlsx (backup Summary sheet)
export const REWARDS: { key: string; label: L; unit: L; cost: number; icon: typeof Egg }[] = [
  { key: "egg", label: { th: "ไข่ไก่", lo: "ໄຂ່ໄກ່", en: "Egg" }, unit: { th: "ฟอง", lo: "ໜ່ວຍ", en: "pc" }, cost: 1, icon: Egg },
  { key: "canned_fish", label: { th: "ปลากระป๋อง", lo: "ປາກະປ໋ອງ", en: "Canned fish" }, unit: { th: "กระป๋อง", lo: "ກະປ໋ອງ", en: "can" }, cost: 6, icon: Fish },
  { key: "noodle", label: { th: "มาม่า", lo: "ໝີ່ກຶ່ງສຳເລັດຮູບ", en: "Instant noodle" }, unit: { th: "ห่อ", lo: "ຊອງ", en: "pack" }, cost: 2, icon: Soup },
  { key: "blanket", label: { th: "ผ้าห่ม TOTO", lo: "ຜ້າຫົ່ມ TOTO", en: "TOTO blanket" }, unit: { th: "ผืน", lo: "ຜືນ", en: "pc" }, cost: 100, icon: Bed },
  { key: "pillow", label: { th: "หมอนใยสังเคราะห์", lo: "ໝອນໃຍສັງເຄາະ", en: "Fiber pillow" }, unit: { th: "ใบ", lo: "ໃບ", en: "pc" }, cost: 80, icon: Package },
  { key: "pillowcase", label: { th: "ปลอกหมอน", lo: "ປອກໝອນ", en: "Pillowcase" }, unit: { th: "ผืน", lo: "ຜືນ", en: "pc" }, cost: 50, icon: Package },
  { key: "dipper", label: { th: "ขันน้ำ", lo: "ຂັນນ້ຳ", en: "Water dipper" }, unit: { th: "อัน", lo: "ອັນ", en: "pc" }, cost: 5, icon: Droplet },
  { key: "mattress", label: { th: "ที่นอน 3 พับ", lo: "ບ່ອນນອນ 3 ພັບ", en: "3-fold mattress" }, unit: { th: "ผืน", lo: "ຜືນ", en: "pc" }, cost: 110, icon: BedDouble },
  { key: "canteen_card", label: { th: "บัตรแคนทีน 500,000 กีบ", lo: "ບັດແຄນທີນ 500,000 ກີບ", en: "Canteen card 500,000 kip" }, unit: { th: "บัตร", lo: "ບັດ", en: "card" }, cost: 130, icon: CreditCard },
];

// Names used by older exchange records, mapped to the current reward names
export const LEGACY_REWARD_NAMES: Record<string, string> = { "ไข่": "ไข่ไก่" };

export type ScoreRow = {
  id: string;
  employee_code: string;
  employee_name: string;
  department: string | null;
  total_score: number;
  remaining_score: number;
};

export type ExchangeRow = {
  id: string;
  employee_code: string | null;
  employee_name: string;
  reward_item: string;
  quantity: number;
  points_used: number;
  exchanged_by: string | null;
  exchanged_at: string;
};

export async function fetchScoreForPerson(name: string, code: string, id?: string): Promise<ScoreRow | null> {
  if (id) {
    const { data, error } = await supabase.from("accumulated_scores").select("*").eq("id", id).maybeSingle();
    return error ? null : (data as ScoreRow | null);
  }
  let query = supabase.from("accumulated_scores").select("*").ilike("employee_name", name.trim());
  if (code.trim()) query = query.eq("employee_code", code.trim());
  const { data, error } = await query.maybeSingle();
  if (error) return null;
  return data as ScoreRow | null;
}

export async function fetchExchangesForPerson(name: string): Promise<ExchangeRow[]> {
  const { data } = await supabase
    .from("reward_exchanges")
    .select("*")
    .ilike("employee_name", name.trim())
    .order("exchanged_at", { ascending: false })
    .limit(10);
  return (data as ExchangeRow[]) ?? [];
}

export async function fetchAllExchanges(): Promise<ExchangeRow[]> {
  const all: ExchangeRow[] = [];
  let from = 0;
  const PAGE = 1000;
  let more = true;
  while (more) {
    const { data } = await supabase
      .from("reward_exchanges")
      .select("*")
      .order("exchanged_at", { ascending: false })
      .range(from, from + PAGE - 1);
    const rows = (data ?? []) as ExchangeRow[];
    all.push(...rows);
    more = rows.length === PAGE;
    from += PAGE;
  }
  return all;
}

export function RewardExchangePanel({
  lang,
  currentUser,
  personName,
  personCode,
  personId,
  onExchanged,
}: {
  lang: Lang;
  currentUser: AppUser;
  personName: string;
  personCode: string;
  personId?: string; // exact accumulated_scores row (avoids duplicate name/code ambiguity)
  onExchanged?: () => void;
}) {
  const [scoreRow, setScoreRow] = useState<ScoreRow | null>(null);
  const [loading, setLoading] = useState(false);
  const [quantities, setQuantities] = useState<Record<string, string>>({});
  const [exchanging, setExchanging] = useState(false);
  const [exchanges, setExchanges] = useState<ExchangeRow[]>([]);

  const reload = async () => {
    setLoading(true);
    const [score, exch] = await Promise.all([
      fetchScoreForPerson(personName, personCode, personId),
      fetchExchangesForPerson(personName),
    ]);
    setScoreRow(score);
    setExchanges(exch);
    setLoading(false);
  };

  useEffect(() => {
    reload();
  }, [personName, personCode, personId]);

  const handleExchange = async (rewardKey: string) => {
    const reward = REWARDS.find((r) => r.key === rewardKey);
    if (!reward) return;
    const qtyStr = quantities[rewardKey] || "";
    const qty = parseInt(qtyStr, 10);
    if (!qty || qty <= 0) {
      toast.error("กรุณากรอกจำนวนที่ต้องการแลก");
      return;
    }
    if (!scoreRow) {
      toast.error("ไม่พบข้อมูลคะแนนสะสมของพนักงานนี้");
      return;
    }

    const pointsNeeded = qty * reward.cost;
    if (scoreRow.remaining_score < pointsNeeded) {
      toast.error(`คะแนนไม่พอ! ต้องการ ${pointsNeeded} คะแนน แต่เหลือเพียง ${scoreRow.remaining_score} คะแนน`);
      return;
    }

    setExchanging(true);
    const newRemaining = scoreRow.remaining_score - pointsNeeded;

    const { error: updateErr } = await supabase
      .from("accumulated_scores")
      .update({ remaining_score: newRemaining, updated_at: new Date().toISOString() })
      .eq("id", scoreRow.id);

    if (updateErr) {
      toast.error("อัปเดตคะแนนไม่สำเร็จ: " + updateErr.message);
      setExchanging(false);
      return;
    }

    const { error: logErr } = await supabase.from("reward_exchanges").insert({
      employee_code: personCode || scoreRow.employee_code,
      employee_name: personName || scoreRow.employee_name,
      reward_item: tr(reward.label, "th"),
      quantity: qty,
      points_used: pointsNeeded,
      exchanged_by: currentUser.username,
    });

    if (logErr) {
      toast.error("บันทึกประวัติการแลกไม่สำเร็จ: " + logErr.message);
      setExchanging(false);
      return;
    }

    setScoreRow({ ...scoreRow, remaining_score: newRemaining });
    setQuantities({ ...quantities, [rewardKey]: "" });
    toast.success(`แลก ${tr(reward.label, lang)} ${qty} ${tr(reward.unit, lang)} สำเร็จ (ใช้ ${pointsNeeded} คะแนน)`);
    reload();
    onExchanged?.();
    setExchanging(false);
  };

  return (
    <div className="rounded-lg border bg-accent/10 p-4">
      <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
        <div className="flex items-center gap-2">
          <Gift className="h-4 w-4 text-primary" />
          <span className="text-sm font-semibold">แลกของรางวัล — {personName}</span>
        </div>
        <div className="text-right">
          {loading ? (
            <Loader2 className="h-5 w-5 animate-spin text-muted-foreground" />
          ) : (
            <div className="flex items-center gap-3">
              <div>
                <p className="text-[10px] text-muted-foreground">คะแนนสะสม</p>
                <p className="text-sm font-semibold">{scoreRow?.total_score ?? 0}</p>
              </div>
              <div>
                <p className="text-[10px] text-muted-foreground">คะแนนคงเหลือ</p>
                <p className="text-lg font-bold text-primary">{scoreRow?.remaining_score ?? 0}</p>
              </div>
              {scoreRow && (
                <div>
                  <p className="text-[10px] text-muted-foreground">ใช้ไปแล้ว</p>
                  <p className="text-sm font-semibold text-muted-foreground">{scoreRow.total_score - scoreRow.remaining_score}</p>
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
        {REWARDS.map((rw) => {
          const Icon = rw.icon;
          const qty = parseInt(quantities[rw.key] || "0", 10) || 0;
          const canAfford = scoreRow ? scoreRow.remaining_score >= rw.cost * (qty || 1) : false;
          return (
            <div
              key={rw.key}
              className={`rounded-lg border p-3 transition-all ${qty > 0 && canAfford ? "border-primary bg-primary/5" : "border-border"}`}
            >
              <div className="mb-1.5 flex items-center gap-2">
                <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-primary/10">
                  <Icon className="h-4 w-4 text-primary" />
                </div>
                <div>
                  <p className="text-xs font-semibold">{tr(rw.label, lang)}</p>
                  <p className="text-[10px] text-muted-foreground">{rw.cost} คะแนน / {tr(rw.unit, lang)}</p>
                </div>
              </div>
              <div className="flex items-center gap-1.5">
                <Input
                  type="number"
                  min="0"
                  value={quantities[rw.key] || ""}
                  onChange={(e) => setQuantities({ ...quantities, [rw.key]: e.target.value })}
                  placeholder="0"
                  className="h-8 w-16 text-center text-sm"
                />
                <span className="text-[10px] text-muted-foreground">{tr(rw.unit, lang)}</span>
                <Button
                  size="sm"
                  disabled={exchanging || !qty || !canAfford}
                  onClick={() => handleExchange(rw.key)}
                  className="ml-auto h-8"
                >
                  <Gift className="mr-1 h-3.5 w-3.5" /> แลก
                </Button>
              </div>
            </div>
          );
        })}
      </div>

      {exchanges.length > 0 && (
        <div className="mt-3">
          <div className="mb-1.5 flex items-center gap-1.5">
            <History className="h-3.5 w-3.5 text-muted-foreground" />
            <span className="text-xs font-semibold text-muted-foreground">ประวัติการแลกล่าสุด</span>
          </div>
          <div className="overflow-x-auto rounded border">
            <table className="w-full border-collapse text-xs">
              <thead className="bg-muted">
                <tr>
                  <th className="border px-1.5 py-1.5 text-left">วันที่</th>
                  <th className="border px-1.5 py-1.5 text-left">ของรางวัล</th>
                  <th className="border px-1.5 py-1.5 text-right">จำนวน</th>
                  <th className="border px-1.5 py-1.5 text-right">ใช้คะแนน</th>
                  <th className="border px-1.5 py-1.5 text-left">โดย</th>
                </tr>
              </thead>
              <tbody>
                {exchanges.map((ex) => (
                  <tr key={ex.id} className="border-b">
                    <td className="border px-1.5 py-1 text-muted-foreground">
                      {new Date(ex.exchanged_at).toLocaleString("th-TH")}
                    </td>
                    <td className="border px-1.5 py-1 font-medium">{ex.reward_item}</td>
                    <td className="border px-1.5 py-1 text-right">{ex.quantity}</td>
                    <td className="border px-1.5 py-1 text-right font-semibold text-primary">{ex.points_used}</td>
                    <td className="border px-1.5 py-1 text-muted-foreground">{ex.exchanged_by ?? "-"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

// Standalone card version for backward compatibility (used at bottom of daily tab)
export function RewardExchange({
  lang,
  currentUser,
  selectedName,
  selectedCode,
}: {
  lang: Lang;
  currentUser: AppUser;
  selectedName: string;
  selectedCode: string;
}) {
  const [expanded, setExpanded] = useState(false);

  if (!selectedName.trim()) return null;

  return (
    <div className="rounded-lg border bg-card">
      <button
        onClick={() => setExpanded(!expanded)}
        className="flex w-full items-center justify-between p-4"
      >
        <div className="flex items-center gap-2">
          <Gift className="h-5 w-5 text-primary" />
          <h2 className="text-base font-semibold sm:text-lg">แลกของรางวัล — {selectedName}</h2>
        </div>
        <ChevronDown className={`h-5 w-5 text-muted-foreground transition-transform ${expanded ? "rotate-180" : ""}`} />
      </button>
      {expanded && (
        <div className="px-4 pb-4">
          <RewardExchangePanel
            lang={lang}
            currentUser={currentUser}
            personName={selectedName}
            personCode={selectedCode}
          />
        </div>
      )}
    </div>
  );
}
