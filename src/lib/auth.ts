import { supabase } from "@/integrations/supabase/client";

export type AppUser = {
  id: string;
  username: string;
  full_name: string;
  role: "admin" | "user";
  department: string | null;
};

const SESSION_KEY = "waste-rewards-session";

function hashPassword(pw: string): string {
  return btoa(pw.split("").reverse().join(""));
}

function verifyPassword(pw: string, hash: string): boolean {
  return hashPassword(pw) === hash;
}

export async function loginUser(
  username: string,
  password: string,
): Promise<{ user: AppUser; error?: string }> {
  const { data, error } = await supabase
    .from("app_users")
    .select("*")
    .eq("username", username)
    .eq("is_active", true)
    .maybeSingle();

  if (error) return { user: null as unknown as AppUser, error: error.message };
  if (!data) return { user: null as unknown as AppUser, error: "ไม่พบผู้ใช้งาน หรือบัญชีถูกปิดใช้งาน" };
  if (!verifyPassword(password, data.password_hash)) {
    return { user: null as unknown as AppUser, error: "รหัสผ่านไม่ถูกต้อง" };
  }

  const user: AppUser = {
    id: data.id,
    username: data.username,
    full_name: data.full_name,
    role: data.role,
    department: data.department,
  };

  localStorage.setItem(SESSION_KEY, JSON.stringify(user));
  return { user };
}

export function getStoredSession(): AppUser | null {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    if (!raw) return null;
    return JSON.parse(raw) as AppUser;
  } catch {
    return null;
  }
}

export function logoutUser() {
  localStorage.removeItem(SESSION_KEY);
}

export async function fetchAllUsers(): Promise<
  (AppUser & { is_active: boolean; created_at: string })[]
> {
  const { data, error } = await supabase
    .from("app_users")
    .select("*")
    .order("created_at", { ascending: true });

  if (error) throw error;
  return (data ?? []).map((r: Record<string, unknown>) => ({
    id: r.id as string,
    username: r.username as string,
    full_name: r.full_name as string,
    role: r.role as "admin" | "user",
    department: r.department as string | null,
    is_active: r.is_active as boolean,
    created_at: r.created_at as string,
  }));
}

export async function createUser(input: {
  username: string;
  password: string;
  full_name: string;
  role: "admin" | "user";
  department?: string;
}): Promise<{ error?: string }> {
  const { error } = await supabase.from("app_users").insert({
    username: input.username,
    password_hash: hashPassword(input.password),
    full_name: input.full_name,
    role: input.role,
    department: input.department || null,
  });
  if (error) return { error: error.message };
  return {};
}

export async function updateUser(
  id: string,
  input: {
    full_name?: string;
    role?: "admin" | "user";
    department?: string;
    is_active?: boolean;
    password?: string;
  },
): Promise<{ error?: string }> {
  const update: Record<string, unknown> = { updated_at: new Date().toISOString() };
  if (input.full_name !== undefined) update.full_name = input.full_name;
  if (input.role !== undefined) update.role = input.role;
  if (input.department !== undefined) update.department = input.department;
  if (input.is_active !== undefined) update.is_active = input.is_active;
  if (input.password) update.password_hash = hashPassword(input.password);

  const { error } = await supabase.from("app_users").update(update).eq("id", id);
  if (error) return { error: error.message };
  return {};
}

export async function deleteUser(id: string): Promise<{ error?: string }> {
  const { error } = await supabase.from("app_users").delete().eq("id", id);
  if (error) return { error: error.message };
  return {};
}
