import { useEffect, useState } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card } from "@/components/ui/card";
import { Plus, Trash2, Pencil, Check, X, Users, Shield, UserCog, Loader2 } from "lucide-react";
import { toast } from "sonner";
import {
  fetchAllUsers,
  createUser,
  updateUser,
  deleteUser,
  type AppUser,
} from "@/lib/auth";

type UserRow = AppUser & { is_active: boolean; created_at: string };

export function AdminUsers({ currentUserUsername }: { currentUserUsername: string }) {
  const [users, setUsers] = useState<UserRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [showAdd, setShowAdd] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editDraft, setEditDraft] = useState<UserRow | null>(null);

  const [newUser, setNewUser] = useState({
    username: "",
    password: "",
    full_name: "",
    role: "user" as "admin" | "user",
    department: "",
  });

  const loadUsers = async () => {
    setLoading(true);
    try {
      const data = await fetchAllUsers();
      setUsers(data);
    } catch (e) {
      toast.error("โหลดข้อมูลผู้ใช้งานไม่สำเร็จ: " + (e as Error).message);
    }
    setLoading(false);
  };

  useEffect(() => {
    loadUsers();
  }, []);

  const handleAdd = async () => {
    if (!newUser.username.trim() || !newUser.password.trim() || !newUser.full_name.trim()) {
      toast.error("กรุณากรอกข้อมูลให้ครบ");
      return;
    }
    const { error } = await createUser({
      username: newUser.username.trim(),
      password: newUser.password,
      full_name: newUser.full_name.trim(),
      role: newUser.role,
      department: newUser.department.trim() || undefined,
    });
    if (error) {
      toast.error("เพิ่มผู้ใช้งานไม่สำเร็จ: " + error);
      return;
    }
    toast.success(`เพิ่มผู้ใช้งาน "${newUser.username}" สำเร็จ`);
    setNewUser({ username: "", password: "", full_name: "", role: "user", department: "" });
    setShowAdd(false);
    loadUsers();
  };

  const handleSaveEdit = async () => {
    if (!editDraft) return;
    const { error } = await updateUser(editDraft.id, {
      full_name: editDraft.full_name,
      role: editDraft.role,
      department: editDraft.department ?? "",
      is_active: editDraft.is_active,
    });
    if (error) {
      toast.error("แก้ไขไม่สำเร็จ: " + error);
      return;
    }
    toast.success("แก้ไขผู้ใช้งานสำเร็จ");
    setEditingId(null);
    setEditDraft(null);
    loadUsers();
  };

  const handleDelete = async (id: string, username: string) => {
    if (username === currentUserUsername) {
      toast.error("ไม่สามารถลบบัญชีตัวเองได้");
      return;
    }
    if (!confirm(`ต้องการลบผู้ใช้งาน "${username}" ใช่หรือไม่?`)) return;
    const { error } = await deleteUser(id);
    if (error) {
      toast.error("ลบไม่สำเร็จ: " + error);
      return;
    }
    toast.success("ลบผู้ใช้งานสำเร็จ");
    loadUsers();
  };

  const handleResetPassword = async (id: string, username: string) => {
    const newPw = prompt(`ตั้งรหัสผ่านใหม่สำหรับ "${username}":`);
    if (!newPw || !newPw.trim()) return;
    const { error } = await updateUser(id, { password: newPw });
    if (error) {
      toast.error("เปลี่ยนรหัสผ่านไม่สำเร็จ: " + error);
      return;
    }
    toast.success("เปลี่ยนรหัสผ่านสำเร็จ");
  };

  return (
    <div className="w-full min-w-0 space-y-4 overflow-x-hidden sm:space-y-6">
      <Card className="p-3 sm:p-4 lg:p-6">
        <div className="mb-4 flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <Users className="h-5 w-5 text-primary" />
            <h2 className="text-base font-semibold sm:text-lg">จัดการผู้ใช้งาน</h2>
          </div>
          <Button size="sm" onClick={() => setShowAdd(!showAdd)}>
            <Plus className="mr-1 h-4 w-4" /> เพิ่มผู้ใช้งาน
          </Button>
        </div>

        {showAdd && (
          <div className="mb-4 rounded-lg border bg-accent/20 p-4">
            <h3 className="mb-3 text-sm font-semibold">เพิ่มผู้ใช้งานใหม่</h3>
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
              <div>
                <label className="mb-1 block text-xs font-medium text-muted-foreground">Username *</label>
                <Input value={newUser.username} onChange={(e) => setNewUser({ ...newUser, username: e.target.value })} />
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-muted-foreground">Password *</label>
                <Input type="password" value={newUser.password} onChange={(e) => setNewUser({ ...newUser, password: e.target.value })} />
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-muted-foreground">ชื่อ-นามสกุล *</label>
                <Input value={newUser.full_name} onChange={(e) => setNewUser({ ...newUser, full_name: e.target.value })} />
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-muted-foreground">บทบาท</label>
                <select
                  value={newUser.role}
                  onChange={(e) => setNewUser({ ...newUser, role: e.target.value as "admin" | "user" })}
                  className="h-10 w-full rounded-md border bg-background px-3 text-sm"
                >
                  <option value="user">User (ผู้ใช้ทั่วไป)</option>
                  <option value="admin">Admin (ผู้ดูแลระบบ)</option>
                </select>
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-muted-foreground">หน่วยงาน</label>
                <Input value={newUser.department} onChange={(e) => setNewUser({ ...newUser, department: e.target.value })} />
              </div>
            </div>
            <div className="mt-3 flex gap-2">
              <Button size="sm" onClick={handleAdd}>
                <Check className="mr-1 h-4 w-4" /> บันทึก
              </Button>
              <Button size="sm" variant="outline" onClick={() => setShowAdd(false)}>
                <X className="mr-1 h-4 w-4" /> ยกเลิก
              </Button>
            </div>
          </div>
        )}

        {loading ? (
          <div className="flex items-center justify-center py-8 text-muted-foreground">
            <Loader2 className="mr-2 h-5 w-5 animate-spin" /> กำลังโหลด...
          </div>
        ) : users.length === 0 ? (
          <p className="py-8 text-center text-sm text-muted-foreground">ยังไม่มีผู้ใช้งาน</p>
        ) : (
          <div className="overflow-x-auto rounded-lg border">
            <table className="w-full border-collapse text-sm">
              <thead className="bg-muted">
                <tr>
                  <th className="border px-2 py-2 text-left">Username</th>
                  <th className="border px-2 py-2 text-left">ชื่อ-นามสกุล</th>
                  <th className="border px-2 py-2 text-left">บทบาท</th>
                  <th className="border px-2 py-2 text-left">หน่วยงาน</th>
                  <th className="border px-2 py-2 text-center">สถานะ</th>
                  <th className="border px-2 py-2 text-center">จัดการ</th>
                </tr>
              </thead>
              <tbody>
                {users.map((u) => {
                  const isEditing = editingId === u.id && editDraft;
                  if (isEditing && editDraft) {
                    return (
                      <tr key={u.id} className="border-b bg-accent/30">
                        <td className="border px-2 py-1.5 font-mono">{u.username}</td>
                        <td className="border px-1 py-1">
                          <Input value={editDraft.full_name} onChange={(e) => setEditDraft({ ...editDraft, full_name: e.target.value })} className="h-8" />
                        </td>
                        <td className="border px-1 py-1">
                          <select
                            value={editDraft.role}
                            onChange={(e) => setEditDraft({ ...editDraft, role: e.target.value as "admin" | "user" })}
                            className="h-8 w-full rounded border bg-background px-1 text-xs"
                          >
                            <option value="user">User</option>
                            <option value="admin">Admin</option>
                          </select>
                        </td>
                        <td className="border px-1 py-1">
                          <Input value={editDraft.department ?? ""} onChange={(e) => setEditDraft({ ...editDraft, department: e.target.value })} className="h-8" />
                        </td>
                        <td className="border px-1 py-1 text-center">
                          <button
                            onClick={() => setEditDraft({ ...editDraft, is_active: !editDraft.is_active })}
                            className={`rounded px-2 py-1 text-xs font-medium ${editDraft.is_active ? "bg-green-100 text-green-700" : "bg-red-100 text-red-700"}`}
                          >
                            {editDraft.is_active ? "ใช้งาน" : "ปิดใช้งาน"}
                          </button>
                        </td>
                        <td className="border px-1 py-1.5">
                          <div className="flex justify-center gap-1">
                            <button onClick={handleSaveEdit} className="flex h-7 w-7 items-center justify-center rounded bg-primary text-primary-foreground hover:bg-primary/90" title="save">
                              <Check className="h-3.5 w-3.5" />
                            </button>
                            <button onClick={() => { setEditingId(null); setEditDraft(null); }} className="flex h-7 w-7 items-center justify-center rounded border text-muted-foreground hover:bg-muted" title="cancel">
                              <X className="h-3.5 w-3.5" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  }
                  return (
                    <tr key={u.id} className="border-b hover:bg-muted/30">
                      <td className="border px-2 py-1.5 font-mono font-medium">{u.username}</td>
                      <td className="border px-2 py-1.5">{u.full_name}</td>
                      <td className="border px-2 py-1.5">
                        {u.role === "admin" ? (
                          <span className="inline-flex items-center gap-1 rounded-full bg-primary/10 px-2 py-0.5 text-xs font-medium text-primary">
                            <Shield className="h-3 w-3" /> Admin
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 rounded-full bg-muted px-2 py-0.5 text-xs font-medium text-muted-foreground">
                            <UserCog className="h-3 w-3" /> User
                          </span>
                        )}
                      </td>
                      <td className="border px-2 py-1.5 text-muted-foreground">{u.department || "-"}</td>
                      <td className="border px-2 py-1.5 text-center">
                        <span className={`rounded-full px-2 py-0.5 text-xs font-medium ${u.is_active ? "bg-green-100 text-green-700" : "bg-red-100 text-red-700"}`}>
                          {u.is_active ? "ใช้งาน" : "ปิดใช้งาน"}
                        </span>
                      </td>
                      <td className="border px-1 py-1.5">
                        <div className="flex justify-center gap-1">
                          <button
                            onClick={() => { setEditingId(u.id); setEditDraft({ ...u }); }}
                            className="flex h-7 w-7 items-center justify-center rounded text-muted-foreground hover:bg-primary hover:text-primary-foreground"
                            title="edit"
                          >
                            <Pencil className="h-3.5 w-3.5" />
                          </button>
                          <button
                            onClick={() => handleResetPassword(u.id, u.username)}
                            className="flex h-7 w-7 items-center justify-center rounded text-muted-foreground hover:bg-amber-500 hover:text-white"
                            title="reset password"
                          >
                            <span className="text-xs font-bold">PW</span>
                          </button>
                          <button
                            onClick={() => handleDelete(u.id, u.username)}
                            className="flex h-7 w-7 items-center justify-center rounded text-muted-foreground hover:bg-destructive hover:text-destructive-foreground"
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

        <div className="mt-4 rounded-lg border border-amber-300 bg-amber-50 p-4">
          <div className="flex items-start gap-3">
            <Shield className="mt-0.5 h-5 w-5 flex-shrink-0 text-amber-600" />
            <div className="text-sm text-amber-800">
              <p className="font-semibold">หมายเหตุ</p>
              <p className="mt-1">Admin เข้าถึงทุกหน้าได้ รวมถึงหน้าจัดการผู้ใช้งาน | User เข้าได้เฉพาะหน้าบันทึกรายวันและแลกของรางวัล</p>
              <p className="mt-1">Admin เริ่มต้น: username <code className="font-mono font-bold">admin</code> password <code className="font-mono font-bold">admin123</code> (กรุณาเปลี่ยนรหัสผ่านหลังเข้าสู่ระบบ)</p>
            </div>
          </div>
        </div>
      </Card>
    </div>
  );
}
