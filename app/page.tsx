"use client";

import {
  useEffect,
  useMemo,
  useState,
} from "react";
import * as XLSX from "xlsx";
import { supabase } from "@/lib/supabase";

type Employee = {
  id?: string
  employee_key: string | null;

  employee_code: string;
  vendor: string;
  branch: string;
  title: string;
  first_name: string;
  last_name: string;
  gender: string;
  date_of_birth: string | null;
  id_card: string;
  employment_date: string | null;
  effective_date: string | null;
  plan: number | null;
  insurance_type: string;
  department: string;
  bank_account: string;
  bank_name: string;
  phone: string;
  remark: string;
  resignation_date: string | null;
  status: string;
  insurance_card_no: string;
life_plan: string;
};
type Claim = {
  id?: string;

  policy_name: string;
  reference_no: string;
  claim_status: string;
  relationship: string;
  employee_name: string;
  id_card: string;
  plan: number | null;
  claim_type: string;
  hospital_date: string | null;
  receipt_amount: number;
  insurance_paid: number;
  hospital_name: string;
  payment_type: string;
  claim_no: string;
  affiliation: string;
  claim_key?: string;
};

// =========================
// กลุ่มรายชื่อแจ้งออกที่ต้องดูเพิ่ม
// =========================
type OutGroupKey =
  | "resigned"
  | "dateDiff"
  | "noDate"
  | "codeNameDiff"
  | "nameOnly"
  | "notFound"
  | "dupInFile";

type OutProblem = {
  row: any;
  group: OutGroupKey;
  reason: string;
};

const OUT_GROUPS: Record<
  OutGroupKey,
  {
    label: string;
    hint: string;
    card: string;
    text: string;
    check: boolean;
  }
> = {
  resigned: {
    label: "แจ้งออกไปแล้ว",
    hint: "ในระบบลาออกแล้ว ไม่ต้องทำอะไร",
    card: "border-slate-200 bg-slate-50",
    text: "text-slate-600",
    check: false,
  },
  dateDiff: {
    label: "วันที่ลาออกไม่ตรง",
    hint: "ลาออกในระบบแล้ว แต่วันที่ต่างจากไฟล์",
    card: "border-amber-300 bg-amber-50",
    text: "text-amber-600",
    check: true,
  },
  noDate: {
    label: "ไม่มีวันที่ลาออก",
    hint: "ไฟล์ไม่ระบุวันที่ ระบบจึงไม่บันทึกให้",
    card: "border-orange-300 bg-orange-50",
    text: "text-orange-600",
    check: true,
  },
  codeNameDiff: {
    label: "รหัสตรงแต่ชื่อไม่ตรง",
    hint: "รหัสอาจถูกใช้กับคนอื่น",
    card: "border-rose-300 bg-rose-50",
    text: "text-rose-600",
    check: true,
  },
  nameOnly: {
    label: "ตรงแค่ชื่อ",
    hint: "ไม่พบรหัส แต่ชื่อตรงกับ 1 คนในระบบ",
    card: "border-yellow-300 bg-yellow-50",
    text: "text-yellow-700",
    check: true,
  },
  notFound: {
    label: "ไม่พบในระบบ",
    hint: "ไม่เจอทั้งรหัสและชื่อ",
    card: "border-purple-300 bg-purple-50",
    text: "text-purple-600",
    check: true,
  },
  dupInFile: {
    label: "รหัสซ้ำในไฟล์",
    hint: "รหัสเดียวกันมีหลายแถว บันทึกแค่แถวแรก",
    card: "border-sky-300 bg-sky-50",
    text: "text-sky-600",
    check: true,
  },
};

// =========================
// แปลงวันที่จาก Excel
// =========================
function formatExcelDate(value: any): string | null {
  if (!value) return null;

    if (value instanceof Date) {
  const year = value.getUTCFullYear();
  const month = String(value.getUTCMonth() + 1).padStart(2, "0");
  const day = String(value.getUTCDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value);

    if (!date) return null;

    const year = date.y;
    const month = String(date.m).padStart(2, "0");
    const day = String(date.d).padStart(2, "0");

    return `${year}-${month}-${day}`;
  }

  const text = String(value).trim();

  if (!text) return null;


    // รูปแบบ วัน/เดือน/ปี เช่น 01/01/2026 หรือ 1-2-2569
  const dmy = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);

  if (dmy) {
    const day = Number(dmy[1]);
    const month = Number(dmy[2]);
    let year = Number(dmy[3]);

    // ปี พ.ศ. แปลงเป็น ค.ศ.
    if (year > 2400) year -= 543;

    if (month < 1 || month > 12 || day < 1 || day > 31) return null;

    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  const parsed = new Date(text);

  if (!isNaN(parsed.getTime())) {
    const year = parsed.getFullYear();
    const month = String(parsed.getMonth() + 1).padStart(2, "0");
    const day = String(parsed.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
  }

  return null;
}
// =========================
// คำนวณวันที่มีผลประกัน
// วันที่เริ่มงาน + 119 วัน
// =========================
function calculateEffectiveDate(
  employmentDate: string | null
): string | null {
  if (!employmentDate) return null;

  const date = new Date(`${employmentDate}T00:00:00Z`);

  if (isNaN(date.getTime())) {
    return null;
  }

  date.setUTCDate(date.getUTCDate() + 119);

  const year = date.getUTCFullYear();
  const month = String(
    date.getUTCMonth() + 1
  ).padStart(2, "0");
  const day = String(
    date.getUTCDate()
  ).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

// =========================
// ประเภทปีประกัน
// ดูจากวันที่มีผลประกัน
// =========================
function getInsuranceType(effectiveDate: string | null): string {
  if (!effectiveDate) return "";

  const [year, month, day] = effectiveDate.split("-").map(Number);

  if (!year || !month || !day) return "";

  return year < 2026 || (year === 2026 && month === 1 && day === 1)
    ? "เต็มปี"
    : "ไม่เต็มปี";
}
// =========================
// สร้างชื่อสำหรับ Match
// =========================
const normalizeName = (value: any) => {
  return String(value ?? "")
    .replace(/\u00A0/g, " ")

    // คำนำหน้าชื่อ
    .replace(/นาย/g, "")
    .replace(/นางสาว/g, "")
    .replace(/นาง/g, "")

    .replace(/น\.\s*ส\./g, "")
    .replace(/น\s*ส\./g, "")
    .replace(/น\.ส/g, "")

    // อังกฤษ
    .replace(/\bmr\.?/gi, "")
    .replace(/\bmrs\.?/gi, "")
    .replace(/\bmiss\.?/gi, "")
    .replace(/\bms\.?/gi, "")

    // ลบอักขระแปลก
    .replace(/[.,\-()/\\]/g, "")

    // ลบช่องว่างทั้งหมด
    .replace(/\s+/g, "")

    .trim()
    .toLowerCase();
};
// =========================
// Normalize ข้อมูลสำหรับ Match
// =========================
const normalizeText = (value: any) => {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
};

// =========================
// สร้าง Employee Key
// Vendor + Employee Code + Effective Date
// หรือ Vendor + ID Card + Effective Date
// หรือ Vendor + ชื่อ + นามสกุล + Effective Date
// =========================
const createEmployeeKey = (
  vendor: string,
  employeeCode: string,
  firstName: string,
  lastName: string,
  idCard: string
) => {
  const cleanVendor = normalizeText(vendor);
  const cleanCode = normalizeText(employeeCode);
  const cleanFirstName = normalizeText(firstName);
  const cleanLastName = normalizeText(lastName);
  const cleanIdCard = normalizeText(idCard);

  if (cleanCode) {
    return `${cleanVendor}|code|${cleanCode}`;
  }

  if (cleanIdCard) {
    return `${cleanVendor}|id|${cleanIdCard}`;
  }

  return `${cleanVendor}|name|${cleanFirstName}|${cleanLastName}`;
};

// =========================
// แปลงชื่อชีตที่เป็นวันที่ เช่น "10-8-69" เป็นตัวเลขเวลา
// ใช้เลือกชีตล่าสุด (ถ้าไม่ใช่รูปแบบวันที่ จะได้ -1)
// =========================
const sheetNameToTime = (name: string): number => {
  const m = name.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);

  if (!m) return -1;

  const day = Number(m[1]);
  const month = Number(m[2]);
  let year = Number(m[3]);

  if (year < 100) year += 2500; // 69 -> 2569
  if (year > 2400) year -= 543; // พ.ศ. -> ค.ศ.

  return Date.UTC(year, month - 1, day);
};

// =========================
// ค่าเบี้ยประกันตามแผน
// =========================
const planPremium: Record<number, number> = {
  1: 3417,
  2: 3900,
  3: 4584,
};


// =========================
// คำนวณจำนวนวันที่มีประกัน
// =========================
const calculateDays = (
  effectiveDate: string | null,
  resignationDate: string | null
) => {
  if (!effectiveDate) return 0;

  const start = new Date(effectiveDate);

  const end = resignationDate
    ? new Date(resignationDate)
    : new Date();

  if (isNaN(start.getTime()) || isNaN(end.getTime())) {
    return 0;
  }

  const diffTime = end.getTime() - start.getTime();

  const diffDays = Math.floor(
    diffTime / (1000 * 60 * 60 * 24)
  );

  return Math.max(diffDays, 0);
};


// =========================
// ชื่อพนักงานเต็ม
// =========================
const getEmployeeFullName = (employee: Employee) => {
  return `${employee.title}${employee.first_name} ${employee.last_name}`
    .replace(/\s+/g, " ")
    .trim();
};
// =========================
// ฟังก์ชันจับคู่พนักงานกลาง
// ใช้ร่วมกันใน แจ้งเข้า / แจ้งออก / ประกันส่งกลับ
// =========================
const parseAmount = (v: any) => {
  if (typeof v === "number") return v;
  const n = Number(String(v ?? "").replace(/[,\s฿]/g, ""));
  return isNaN(n) ? 0 : n;
};
const normId = (value: unknown) =>
  String(value ?? "").replace(/\D/g, "");

const normCode = (value: unknown) =>
  String(value ?? "")
    .trim()
    .replace(/\s+/g, "")
    .toUpperCase();

const nameKey = (first: unknown, last: unknown) =>
  normalizeName(first) + "|" + normalizeName(last);

type MatchResult = {
  employee: Employee | null;
  by: "id_card" | "code" | "name" | "none";
  warning?: string;
};

// สร้างตารางค้นหา 3 แบบจากข้อมูลใน DB
const buildLookup = (dbEmployees: Employee[]) => {
  const byId = new Map<string, Employee>();
  const byCode = new Map<string, Employee>();
  const byName = new Map<string, Employee[]>();

  for (const e of dbEmployees) {
    const id = normId(e.id_card);
    if (id) byId.set(id, e);

    const code = normCode(e.employee_code);
    if (code) byCode.set(code, e);

    const nk = nameKey(e.first_name, e.last_name);
    byName.set(nk, [...(byName.get(nk) ?? []), e]);
  }

  return { byId, byCode, byName };
};

// หาพนักงานจาก 1 แถวในไฟล์
const findEmployee = (
  row: {
    id_card?: string;
    employee_code?: string;
    first_name?: string;
    last_name?: string;
  },
  lookup: ReturnType<typeof buildLookup>
): MatchResult => {
  const id = normId(row.id_card);
  const code = normCode(row.employee_code);
  const nk = nameKey(row.first_name, row.last_name);

  // 1) เลขบัตรประชาชน (แม่นที่สุด)
  if (id && lookup.byId.has(id)) {
    const hit = lookup.byId.get(id)!;

    const warning =
      nameKey(hit.first_name, hit.last_name) !== nk
        ? "เลขบัตรตรงแต่ชื่อไม่ตรง"
        : code && normCode(hit.employee_code) !== code
        ? `รหัสไม่ตรง (DB: ${hit.employee_code})`
        : undefined;

    return { employee: hit, by: "id_card", warning };
  }

  // 2) รหัสพนักงาน (เฉพาะแถวที่ไม่มีเลขบัตร เช่น ไฟล์แจ้งออก)
  //    ต้องชื่อตรงด้วย เพราะรหัสอาจถูกนำไปใช้กับคนใหม่
  if (!id && code && lookup.byCode.has(code)) {
    const hit = lookup.byCode.get(code)!;

    if (nameKey(hit.first_name, hit.last_name) === nk) {
      return { employee: hit, by: "code" };
    }

    return {
      employee: null,
      by: "none",
      warning: `รหัสตรงแต่ชื่อไม่ตรง (DB: ${hit.first_name} ${hit.last_name})`,
    };
  }

  // 3) ชื่ออย่างเดียว: ใช้ได้เมื่อเจอคนเดียวเท่านั้น และต้องตรวจเสมอ
  const nameHits = lookup.byName.get(nk) ?? [];

  if (nameHits.length === 1) {
    return {
      employee: nameHits[0],
      by: "name",
      warning: "จับคู่ด้วยชื่อเท่านั้น ควรตรวจสอบ",
    };
  }

  return { employee: null, by: "none" };
};

// =========================
// สรุป Claims ต่อพนักงาน
// ลองเลขบัตรก่อน ถ้าไม่เจอใช้ชื่อ (เฉพาะกรณีเจอคนเดียว)
// =========================
type ClaimSummary = { opd: number; ipd: number; count: number };

const getClaimSummary = (claims: Claim[], employees: Employee[]) => {
  const byId = new Map<string, Employee>();
  const byName = new Map<string, Employee[]>();

  for (const e of employees) {
    const id = normId(e.id_card);
    if (id) byId.set(id, e);

    const nk = normalizeName(`${e.first_name}${e.last_name}`);
    if (nk) byName.set(nk, [...(byName.get(nk) ?? []), e]);
  }

  const summary = new Map<string, ClaimSummary>();
  const unmatched: Claim[] = [];
  let employeeClaims = 0;
  let dependentClaims = 0;

  for (const claim of claims) {
    if (String(claim.relationship ?? "").trim().toUpperCase() !== "E") {
      dependentClaims++;
      continue;
    }
    employeeClaims++;

    let emp = byId.get(normId(claim.id_card));
    if (!emp) {
      const hits = byName.get(normalizeName(claim.employee_name)) ?? [];
      if (hits.length === 1) emp = hits[0];
    }
    if (!emp?.id) {
      unmatched.push(claim);
      continue;
    }

    const cur = summary.get(emp.id) ?? { opd: 0, ipd: 0, count: 0 };
    const type = String(claim.claim_type ?? "").toUpperCase();
    const paid = Number(claim.insurance_paid ?? 0);

    if (type.includes("OPD") || type.includes("ผู้ป่วยนอก")) cur.opd += paid;
    if (type.includes("IPD") || type.includes("ผู้ป่วยใน")) cur.ipd += paid;
    cur.count += 1;

    summary.set(emp.id, cur);
  }

  return { summary, unmatched, employeeClaims, dependentClaims };
};

// =========================
// โหลดพนักงานทั้งหมดจาก Supabase
// =========================
const loadAllEmployees = async () => {
  const allEmployees: Employee[] = [];

  const pageSize = 1000;

  let from = 0;
  let to = pageSize - 1;

  while (true) {
    const {
      data,
      error,
    } = await supabase
      .from("employees")
      .select("*")
      .order("id", {
        ascending: true,
      })
      .range(from, to);

    if (error) {
      throw error;
    }

    if (!data || data.length === 0) {
      break;
    }

    allEmployees.push(...data);

    if (data.length < pageSize) {
      break;
    }

    from += pageSize;
    to += pageSize;
  }

  return allEmployees;
};
// =========================
// โหลด Claims ทั้งหมดจาก Supabase
// =========================
const loadAllClaims = async () => {
  const allClaims: Claim[] = [];

  const pageSize = 1000;

  let from = 0;
  let to = pageSize - 1;

  while (true) {
    const {
      data,
      error,
    } = await supabase
      .from("claims")
      .select("*")
      .order("id", {
        ascending: true,
      })
      .range(from, to);

    if (error) {
      throw error;
    }

    if (!data || data.length === 0) {
      break;
    }

    allClaims.push(...data);

    if (data.length < pageSize) {
      break;
    }

    from += pageSize;
    to += pageSize;
  }

  return allClaims;
};

export default function Home() {
  const [fileName, setFileName] = useState("");
  const [rowCount, setRowCount] = useState<number | null>(null);
  const [error, setError] = useState("");
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [showInNames, setShowInNames] = useState(false);
const [showOutNames, setShowOutNames] = useState(false);
  const [claims, setClaims] = useState<Claim[]>([]);
  // รายชื่อพนักงานแจ้งเข้า แยกเป็น 2 กลุ่ม
const [inNameTab, setInNameTab] = useState<"new" | "existing" | "changed">("new");
const [inNewEmployees, setInNewEmployees] = useState<Employee[]>([]);

const [inExistingEmployees, setInExistingEmployees] = useState<Employee[]>([]);
const [inChanged, setInChanged] = useState<{ emp: Employee; db: Employee }[]>([]);
  const [employeeSearchInput, setEmployeeSearchInput] = useState("");
  const [employeeSearchTerm, setEmployeeSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState("ทั้งหมด");
  // =========================
// สรุป Claims แบบรวดเร็ว
// =========================
const filteredEmployees = useMemo(() => {
  const search = employeeSearchTerm.trim().toLowerCase();

  return employees.filter((employee) => {
    const employeeCode = String(employee.employee_code ?? "")
      .trim()
      .toLowerCase();

    const fullName = `${employee.title ?? ""} ${employee.first_name ?? ""} ${employee.last_name ?? ""}`
      .trim()
      .toLowerCase();

    const matchesSearch =
      !search ||
      employeeCode.includes(search) ||
      fullName.includes(search);

    const matchesStatus =
      statusFilter === "ทั้งหมด" ||
      employee.status === statusFilter;

    return matchesSearch && matchesStatus;
  });
}, [employees, employeeSearchTerm, statusFilter]);
const claimResult = useMemo(
  () => getClaimSummary(claims, employees),
  [claims, employees]
);

const claimSummaryMap = claimResult.summary;

const matchStats = {
  total: claims.length,
  employeeClaims: claimResult.employeeClaims,
  dependentClaims: claimResult.dependentClaims,
  matched: claimResult.employeeClaims - claimResult.unmatched.length,
  unmatched: claimResult.unmatched.length,
};



  const [isLoading, setIsLoading] = useState(false);
  const [isImporting, setIsImporting] = useState(false);

  const [claimFileName, setClaimFileName] = useState("");
  const [claimRowCount, setClaimRowCount] = useState<number | null>(null);
  const [claimPreview, setClaimPreview] = useState<Claim[]>([]);
  const [isClaimLoading, setIsClaimLoading] = useState(false);
  const [isClaimImporting, setIsClaimImporting] = useState(false);

  const [successMessage, setSuccessMessage] = useState("");
  // =========================
// แจ้งเข้า
// =========================
const [inFileName, setInFileName] = useState("");
const [inRowCount, setInRowCount] = useState<number | null>(null);
const [inPreview, setInPreview] = useState<Employee[]>([]);
const [inAllEmployees, setInAllEmployees] = useState<Employee[]>([]);
const [isInLoading, setIsInLoading] = useState(false);
const [isInImporting, setIsInImporting] = useState(false);
const [inNewCount, setInNewCount] = useState<number | null>(null);
const [inDuplicateCount, setInDuplicateCount] = useState<number | null>(null);
const [inSuccessMessage, setInSuccessMessage] = useState("");

// แจ้งออก
const [outFileName, setOutFileName] = useState("");
const [outRowCount, setOutRowCount] = useState<number | null>(null);
const [outPreview, setOutPreview] = useState<any[]>([]);
const [outProblems, setOutProblems] = useState<OutProblem[]>([]);
const [outGroupTab, setOutGroupTab] = useState<OutGroupKey | "allCheck">("allCheck");
const [isOutLoading, setIsOutLoading] = useState(false);
const [isOutImporting, setIsOutImporting] = useState(false);
const [showOutList, setShowOutList] = useState(false);
const [outSuccessMessage, setOutSuccessMessage] = useState("");

// ประกันส่งกลับ
const [insuranceFileName, setInsuranceFileName] = useState("");
const [insuranceRowCount, setInsuranceRowCount] = useState<number | null>(null);
const [isInsuranceLoading, setIsInsuranceLoading] = useState(false);
const [insuranceMatched, setInsuranceMatched] = useState<Employee[]>([]);
const [insuranceUnmatched, setInsuranceUnmatched] = useState<Employee[]>([]);
const [isInsuranceImporting, setIsInsuranceImporting] =
  useState(false);

      // =========================
  // โหลดข้อมูลพนักงานจาก Supabase
  // =========================
useEffect(() => {
  const loadData = async () => {
    setIsLoading(true);
    setError("");

    // ==========================================
    // โหลด Employees
    // ==========================================
    try {

  const employeeData =
    await loadAllEmployees();

  setEmployees(employeeData);

  setRowCount(employeeData.length);

} catch (employeeError) {

  console.error(
    "Supabase Employee Error:",
    employeeError
  );

  setError(
    `ไม่สามารถโหลดข้อมูลพนักงานได้: ${
      employeeError instanceof Error
        ? employeeError.message
        : "เกิดข้อผิดพลาด"
    }`
  );

}
// ==========================================
// โหลด Claims จาก Supabase
// ==========================================
try {

  const claimData =
    await loadAllClaims();

  console.log(
    "โหลด Claims สำเร็จ:",
    claimData.length
  );

  setClaims(claimData);

} catch (claimError) {

  console.error(
    "Supabase Claim Error:",
    claimError
  );

  setClaims([]);

}
    setIsLoading(false);
  };

  loadData();
}, []);
  // =========================
  // อ่านไฟล์ Excel
  // =========================
 const handleFileUpload = (
  event: React.ChangeEvent<HTMLInputElement>
) => {
  const file = event.target.files?.[0];

  console.log("📂 FILE ที่เลือก:", file);

  if (!file) {
    console.log("❌ ไม่มีไฟล์ถูกเลือก");
    return;
  }

  setFileName(file.name);

  setFileName(file.name);
  setError("");
  setSuccessMessage("");
  setRowCount(null);
  setEmployees([]);
  setIsLoading(true);

  const reader = new FileReader();

  reader.onload = async (e) => {
    try {
      const data = e.target?.result;

      if (!data) {
        throw new Error("ไม่พบข้อมูลไฟล์");
      }

            const workbook = XLSX.read(data, {
        type: "array",
        cellDates: true,
        sheets: "ทำ",
      });

      const worksheet = workbook.Sheets["ทำ"];
      if (!worksheet) {
        throw new Error(
        `ไม่พบชีต "ทำ" ในไฟล์ (ชีตที่มี: ${workbook.SheetNames.join(", ")})`
      );
}

      const rows = XLSX.utils.sheet_to_json<any[]>(worksheet, {
        header: 1,
        defval: "",
      });

     
      if (rows.length < 3) {
        throw new Error(
          "ไฟล์ Excel ไม่มีข้อมูลพนักงาน หรือรูปแบบไฟล์ไม่ถูกต้อง"
        );
      }

      // ---------------------------------
      // ไฟล์ตั้งต้น
      //
      // แถวที่ 1 = Header ภาษาไทย
      // แถวที่ 2 = Header ภาษาอังกฤษ
      // แถวที่ 3 เป็นต้นไป = ข้อมูลพนักงาน
      // ---------------------------------
      const dataRows = rows.slice(4);
      console.log(
  "แถวตัวอย่างไฟล์ตั้งต้น:",
  JSON.stringify(
    dataRows.slice(0, 3).map((r) =>
      r.slice(0, 12).map((v: any, i: number) => [i, String(v)])
    )
  )
);


const mappedEmployees: Employee[] = dataRows
  .filter((row) => {
    const employeeCode = String(row[2] ?? "").trim();
    const firstName = String(row[4] ?? "").trim();
    const lastName = String(row[5] ?? "").trim();

    return (
      employeeCode !== "" &&
      (firstName !== "" || lastName !== "")
    );
  })
  .map((row) => {
    const vendor = String(row[1] ?? "").trim();

    const employeeCode = String(row[2] ?? "")
      .trim()
      .replace(/\s+/g, "")
      .toUpperCase();

    const title = String(row[3] ?? "").trim();
    const firstName = String(row[4] ?? "").trim();
    const lastName = String(row[5] ?? "").trim();

    const idCard = String(row[8] ?? "")
      .trim()
      .replace(/\D/g, "");

    const employmentDate = formatExcelDate(row[9]);

    // วันที่มีผลประกันจริงจะมาจากไฟล์ประกันส่งกลับ
    const effectiveDate: string | null = null;

    return {
      employee_key: createEmployeeKey(
        vendor,
        employeeCode,
        firstName,
        lastName,
        idCard
      ),

      employee_code: employeeCode,
      vendor,
      branch: String(row[1] ?? "").trim(),
      title,
      first_name: firstName,
      last_name: lastName,
      gender: String(row[6] ?? "").trim(),

      date_of_birth: formatExcelDate(row[7]),
      id_card: idCard,

      employment_date: employmentDate,
      effective_date: effectiveDate,

      plan:
        row[11] !== "" &&
        row[11] !== null &&
        row[11] !== undefined
          ? Number(row[11])
          : null,

      insurance_type: getInsuranceType(effectiveDate),

      department: String(row[12] ?? "").trim(),
      bank_account: String(row[13] ?? "").trim(),
      bank_name: String(row[14] ?? "").trim(),
      phone: String(row[15] ?? "").trim(),

      remark: String(row[17] ?? "").trim(),

      resignation_date: null,
      status: "มีผลประกัน",

      insurance_card_no: "",
      life_plan: "",
    };
  });

// =========================
// รวมข้อมูลพนักงานที่มีหลายแถว
// เช่น แถวทำงาน + แถวลาออก
// ให้เหลือพนักงาน 1 คน
// =========================
const employeeMap = new Map<string, Employee>();
const skippedNoId: Employee[] = [];

for (const employee of mappedEmployees) {
  const idCard = String(employee.id_card ?? "")
    .trim()
    .replace(/\D/g, "");

  // ถ้าไม่มีเลขบัตรประชาชน ให้ข้าม
    if (!idCard) {
    skippedNoId.push(employee);
    continue;
  }

  const existing = employeeMap.get(idCard);

  // คนนี้ยังไม่เคยเจอ
  if (!existing) {
    employeeMap.set(idCard, employee);
    continue;
  }

  // ถ้าเจอข้อมูลซ้ำ ให้เติมข้อมูลที่ขาดจากแถวใหม่
if (existing) {
  employeeMap.set(idCard, {
    ...existing,

    employment_date:
      existing.employment_date ||
      employee.employment_date,

    effective_date:
      existing.effective_date ||
      employee.effective_date,

    insurance_type:
      existing.insurance_type ||
      employee.insurance_type,

    plan:
      existing.plan ?? employee.plan,

    // ถ้ามีวันที่ลาออก ให้ถือว่าลาออก
    resignation_date:
      employee.resignation_date ||
      existing.resignation_date,

    status:
      employee.resignation_date ||
      existing.resignation_date
        ? "ลาออก"
        : existing.status,
  });
}
}

if (skippedNoId.length > 0) {
  console.warn("ข้ามพนักงานที่ไม่มีเลขบัตร:", skippedNoId);
  setError(
    `ข้ามพนักงานที่ไม่มีเลขบัตร ${skippedNoId.length} คน: ` +
      skippedNoId
        .slice(0, 10)
        .map((e) => `${e.employee_code} ${e.first_name}`)
        .join(", ")
  );
}

const uniqueEmployees =
  Array.from(employeeMap.values());
  console.log(
  "🔎 ตรวจ effective_date / insurance_type:",
  uniqueEmployees.slice(0, 20).map((employee) => ({
    name: `${employee.first_name} ${employee.last_name}`,
    employment_date: employee.employment_date,
    effective_date: employee.effective_date,
    insurance_type: employee.insurance_type,
    status: employee.status,
  }))
);

const resignedCount = uniqueEmployees.filter(
  (employee) => employee.status === "ลาออก"
).length;

const workingCount = uniqueEmployees.filter(
  (employee) => employee.status === "มีผลประกัน"
).length;


console.log(
  "🔴 จำนวนคนที่เป็นลาออก:",
  uniqueEmployees.filter(
    (employee) => employee.status === "ลาออก"
  ).length
);

console.log(
  "🟢 จำนวนคนที่เป็นมีผลประกัน:",
  uniqueEmployees.filter(
    (employee) => employee.status === "มีผลประกัน"
  ).length
);

console.log(
  "🔴 ตัวอย่างข้อมูลคนที่ถูกมองว่าลาออก:",
  uniqueEmployees
    .filter(
      (employee) => employee.status === "ลาออก"
    )
    .slice(0, 20)
);

console.log(
  "จำนวนแถวในไฟล์:",
  mappedEmployees.length
);

console.log(
  "จำนวนพนักงานหลังรวม:",
  uniqueEmployees.length
);
setEmployees(uniqueEmployees);
setRowCount(uniqueEmployees.length);


    } catch (err) {

      console.error(err);

      setError(
        err instanceof Error
          ? err.message
          : "ไม่สามารถอ่านไฟล์แจ้งเข้าได้"
      );

    } finally {

      setIsLoading(false);

    }
  };

  reader.onerror = () => {

    setError(
      "เกิดข้อผิดพลาดในการอ่านไฟล์แจ้งเข้า"
    );

    setIsLoading(false);

  };

  reader.readAsArrayBuffer(file);
};
const handleInFileUpload = (
  event: React.ChangeEvent<HTMLInputElement>
) => {
  const file = event.target.files?.[0];

  if (!file) return;

  console.log("📄 ไฟล์แจ้งเข้า:", file.name);

  setInFileName(file.name);
  setInChanged([]);
  setInRowCount(null);
  setInPreview([]);
  setInNewCount(null);
  setInDuplicateCount(null);
  setIsInLoading(true);
  setError("");
  setInSuccessMessage("");
  setSuccessMessage("");

  const reader = new FileReader();

  reader.onload = async (e) => {
    try {
      const data = e.target?.result;

      if (!data) {
        throw new Error("ไม่พบข้อมูลไฟล์แจ้งเข้า");
      }

      const workbook = XLSX.read(data, {
        type: "array",
        cellDates: true,
      });

            const firstSheetName =
        workbook.SheetNames.find(
          (_name, index) =>
            (workbook.Workbook?.Sheets?.[index]?.Hidden ?? 0) === 0
        ) ?? workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      const rows = XLSX.utils.sheet_to_json<any[]>(worksheet, {
        header: 1,
        defval: "",
      });

      console.log("========== ไฟล์แจ้งเข้า ==========");
      console.log("ข้อมูลทั้งหมด:", rows);

 const dataRows = rows
  .slice(9)
  .filter((row) => {
    const employeeCode = String(row[2] ?? "").trim();
    const firstName = String(row[4] ?? "").trim();
    const lastName = String(row[5] ?? "").trim();

    return (
      employeeCode !== "" &&
      (firstName !== "" || lastName !== "")
    );
  });

      console.log(
        "จำนวนข้อมูลแจ้งเข้า:",
        dataRows.length
      );

      const mappedEmployees: Employee[] =
        dataRows.map((row) => {
          const vendor = String(
            row[1] ?? ""
          ).trim();

          const employeeCode = String(
            row[2] ?? ""
          )
            .trim()
            .replace(/\s+/g, "")
            .toUpperCase();

          const firstName = String(
            row[4] ?? ""
          ).trim();

          const lastName = String(
            row[5] ?? ""
          ).trim();

          const idCard = String(
            row[8] ?? ""
          )
            .trim()
            .replace(/\D/g, "");

          const employmentDate =
  formatExcelDate(row[9]);

const effectiveDate: string | null = null;

          return {
            employee_key: createEmployeeKey(
            vendor,
            employeeCode,
            firstName,
            lastName,
            idCard
          ),

            employee_code: employeeCode,

            vendor,

            branch: vendor,

            title: String(
              row[3] ?? ""
            ).trim(),

            first_name: firstName,

            last_name: lastName,

            gender: String(
              row[6] ?? ""
            ).trim(),

            date_of_birth:
              formatExcelDate(row[7]),

            id_card: idCard,

            employment_date:
              employmentDate,

            effective_date:
              effectiveDate,

            plan:
              row[11] !== "" &&
              row[11] !== null &&
              row[11] !== undefined
                ? Number(row[11])
                : null,

            insurance_type:
            getInsuranceType(effectiveDate),

            department: String(
              row[12] ?? ""
            ).trim(),

            bank_account: String(
              row[13] ?? ""
            ).trim(),

            bank_name: String(
              row[14] ?? ""
            ).trim(),

            phone: String(
              row[15] ?? ""
            ).trim(),

            remark: String(
              row[17] ?? ""
            ).trim(),

            resignation_date: null,

            status: "มีผลประกัน",

            insurance_card_no: "",

            life_plan: "",
          };
        });

      console.log(
        "ข้อมูลแจ้งเข้าที่ Mapping แล้ว:",
        mappedEmployees
      );

       setInAllEmployees(mappedEmployees);

// โหลดข้อมูลสดจากฐานข้อมูล ไม่พึ่ง state employees
const dbEmployees = await loadAllEmployees();

const existingIdCards = new Set(
  dbEmployees
    .map((employee) =>
      String(employee.id_card ?? "").trim().replace(/\D/g, "")
    )
    .filter((idCard) => idCard !== "")
);

const dbCodes = new Set(
  dbEmployees.map((employee) =>
    String(employee.employee_code ?? "").trim().toUpperCase()
  )
);

// Log ตรวจสอบ (ใช้ JSON.stringify เพื่อให้ Safari แสดงข้อความจริง)
console.log(
  "ผลตรวจ:",
  JSON.stringify(
    {
      sheet: firstSheetName,
      fileName: file.name,
      dbRows: dbEmployees.length,
      fileRows: mappedEmployees.length,
      matchByCode: mappedEmployees.filter((e) =>
        dbCodes.has(e.employee_code)
      ).length,
      matchByIdCard: mappedEmployees.filter((e) =>
        existingIdCards.has(e.id_card)
      ).length,
      emptyIdCard: mappedEmployees.filter((e) => !e.id_card).length,
      first3File: mappedEmployees
        .slice(0, 3)
        .map((e) => [e.employee_code, e.first_name, e.id_card.length]),
      first3Db: dbEmployees
        .slice(0, 3)
        .map((e) => [e.employee_code, e.first_name]),
    },
    null, 
    2
  )
);
console.log(
  "รหัสที่ไม่ตรงกับ DB:",
  JSON.stringify(
    mappedEmployees
      .filter((e) => !dbCodes.has(e.employee_code))
      .map((e) => [e.employee_code, e.first_name, e.last_name])
  )
);

const lookup = buildLookup(dbEmployees);

const newEmployees: Employee[] = [];
const duplicateEmployees: Employee[] = [];
const needReview: { emp: Employee; reason: string }[] = [];
const changedList: { emp: Employee; db: Employee }[] = [];
const seenIds = new Set<string>();

for (const emp of mappedEmployees) {
  const id = normId(emp.id_card);
  const nk = nameKey(emp.first_name, emp.last_name);
  const hit = id ? lookup.byId.get(id) : undefined;

  if (hit) {
        const nameMatches = nameKey(hit.first_name, hit.last_name) === nk;
    const codeChanged =
      normCode(hit.employee_code) !== normCode(emp.employee_code);

    if (nameMatches && codeChanged) {
      // เลขบัตรและชื่อตรง แต่รหัสพนักงานเปลี่ยน
      changedList.push({ emp, db: hit });
    } else {
      duplicateEmployees.push(emp);
      if (!nameMatches) {
        needReview.push({ emp, reason: "เลขบัตรตรงแต่ชื่อไม่ตรง" });
      }
    }
  } else if ((lookup.byName.get(nk) ?? []).length > 0) {
    needReview.push({
      emp,
      reason: id ? "ชื่อตรงแต่เลขบัตรไม่ตรง" : "ชื่อตรงแต่ไม่มีเลขบัตร",
    });
  } else if (!id) {
    needReview.push({ emp, reason: "ไม่มีเลขบัตร" });
  } else if (seenIds.has(id)) {
    needReview.push({ emp, reason: "เลขบัตรซ้ำกันเองในไฟล์" });
  } else {
    newEmployees.push(emp);
    seenIds.add(id);
  }
}

console.log(
  "สรุป:",
  JSON.stringify({
    new: newEmployees.length,
    duplicate: duplicateEmployees.length,
    needReview: needReview.length,
    changed: changedList.length,
  })
);
console.log(
  "ต้องตรวจ:",
  JSON.stringify(
    needReview.map((r) => [r.emp.employee_code, r.emp.first_name, r.reason])
  )
);

setInNewEmployees(newEmployees);
setInChanged(changedList);

console.log(
  "รหัสเปลี่ยน:",
  JSON.stringify(
    changedList.map((c) => [
      c.db.employee_code,
      c.emp.employee_code,
      c.emp.first_name,
      c.db.status,
    ])
  )
);

setInExistingEmployees(duplicateEmployees);

setInPreview(newEmployees);
setInRowCount(mappedEmployees.length);
setInNewCount(newEmployees.length);
setInDuplicateCount(duplicateEmployees.length);

    } catch (err) {
      console.error(
        "In File Error:",
        err
      );

      setError(
        err instanceof Error
          ? err.message
          : "ไม่สามารถอ่านไฟล์แจ้งเข้าได้"
      );
    } finally {
      setIsInLoading(false);
    }
  };

  reader.onerror = () => {
    setError(
      "ไม่สามารถอ่านไฟล์แจ้งเข้าได้"
    );

    setIsInLoading(false);
  };

  reader.readAsArrayBuffer(file);
};
// บันทึก "แจ้งเข้า"
// เพิ่มเฉพาะพนักงานใหม่
// =========================
const handleImportInToSupabase = async () => {
  console.log("🔥🔥🔥 HANDLE IMPORT IN ถูกเรียกแล้ว");

  console.log(
    "inPreview:",
    inPreview.length
  );

  console.log("🔥 กดปุ่มบันทึกแจ้งเข้าแล้ว");

  console.log(
    "จำนวน inPreview:",
    inPreview.length
  );

  console.log(
  "📌 เริ่มกระบวนการบันทึกแจ้งเข้า"
);

  setIsInImporting(true);
  setError("");
  setSuccessMessage("");
  setInSuccessMessage("");

  try {

// ==========================================
// เตรียมข้อมูลพนักงานใหม่สำหรับบันทึก
// ==========================================
const newEmployees = inPreview;

// ==========================================
// เติมวันที่เริ่มงานจากไฟล์แจ้งเข้า
// ให้กับพนักงานที่มีอยู่แล้วใน Master
// ==========================================
const currentEmployees = await loadAllEmployees();

const lookup = buildLookup(currentEmployees);

let updatedEmploymentCount = 0;

for (const inEmployee of inAllEmployees) {
  const code = normCode(inEmployee.employee_code);

const match = findEmployee(inEmployee, lookup);
const masterEmployee =
  match.by === "id_card" || match.by === "code" ? match.employee : null;
  if (!masterEmployee) continue;

  // เติมเฉพาะกรณี Master ไม่มีวันที่เริ่มงาน
  if (
    !masterEmployee.employment_date &&
    inEmployee.employment_date
  ) {
    const { error } = await supabase
      .from("employees")
      .update({
        employment_date: inEmployee.employment_date,
        updated_at: new Date().toISOString(),
      })
      .eq("id", masterEmployee.id);

    if (error) {
      throw new Error(
        `อัปเดตวันที่เริ่มงาน ${code} ไม่สำเร็จ: ${error.message}`
      );
    }

    updatedEmploymentCount++;
  }
}

console.log(
  "📅 เติมวันที่เริ่มงานจากแจ้งเข้า:",
  updatedEmploymentCount,
  "คน"
);


// ==========================================
// อัปเดตคนที่รหัสเปลี่ยน / กลับเข้าทำงาน
// ==========================================
let updatedChangedCount = 0;

for (const item of inChanged) {
  // ใช้ข้อมูลสดที่เพิ่งโหลดจากฐานข้อมูล
  const fresh = currentEmployees.find((e) => e.id === item.db.id);

  if (!fresh) continue;

  const oldCode = normCode(fresh.employee_code);
  const newCode = normCode(item.emp.employee_code);

  // รหัสในฐานข้อมูลเปลี่ยนไปแล้วหรือไม่ใช่รหัสที่เห็นตอนอัปโหลด ให้ข้าม
  if (oldCode !== normCode(item.db.employee_code)) continue;
  if (!newCode || newCode === oldCode) continue;

  // รหัสใหม่ถูกใช้โดยคนอื่นอยู่แล้ว ให้ข้ามเพื่อกันข้อมูลชนกัน
  const used = lookup.byCode.get(newCode);
  if (used && used.id !== fresh.id) {
    console.warn("ข้ามรหัสเปลี่ยน เพราะรหัสใหม่มีคนใช้แล้ว:", newCode);
    continue;
  }

  const wasResigned = fresh.status === "ลาออก";

  const oldNote = `[รหัสเดิม ${oldCode}${
    fresh.resignation_date ? `, ลาออกเดิม ${fresh.resignation_date}` : ""
  }]`;

  const { error } = await supabase
    .from("employees")
    .update({
      employee_code: newCode,
      employee_key: createEmployeeKey(
        fresh.vendor,
        newCode,
        fresh.first_name,
        fresh.last_name,
        fresh.id_card
      ),
      ...(wasResigned
        ? { status: "มีผลประกัน", resignation_date: null }
        : {}),
      remark: `${fresh.remark ?? ""} ${oldNote}`.trim(),
      updated_at: new Date().toISOString(),
    })
    .eq("id", fresh.id);

  if (error) {
    throw new Error(
      `อัปเดตรหัสใหม่ ${oldCode} → ${newCode} ไม่สำเร็จ: ${error.message}`
    );
  }

  updatedChangedCount++;
}

if (updatedChangedCount > 0) {
  const refreshed = await loadAllEmployees();
  setEmployees(refreshed);
  setRowCount(refreshed.length);
  setInChanged([]);
}

console.log("🔁 อัปเดตรหัสเปลี่ยน/กลับเข้าทำงาน:", updatedChangedCount, "คน");

console.log(
  "📥 จำนวนพนักงานใหม่ที่จะบันทึก:",
  newEmployees.length
);
 // ==========================================
// DEBUG
// ==========================================
console.log(
  "========== DEBUG แจ้งเข้า =========="
);

console.log(
  "จำนวนพนักงานใหม่ที่จะบันทึก:",
  newEmployees.length
);

console.log(
  "ตัวอย่างพนักงานใหม่:",
  newEmployees.slice(0, 3)
);

console.log(
  "===================================="
);
// ==========================================
// ไม่มีพนักงานใหม่
// ==========================================
if (newEmployees.length === 0) {

  console.log(
    "ℹ️ ไม่มีพนักงานใหม่ ไม่ต้องเพิ่มข้อมูล"
  );

  // ให้ปุ่มแสดงสถานะกำลังบันทึกสั้น ๆ
  await new Promise((resolve) =>
    setTimeout(resolve, 700)
  );

      setInSuccessMessage(
    `บันทึกเข้าสู่ฐานข้อมูลแล้ว ไม่มีพนักงานใหม่` +
      (updatedChangedCount > 0
        ? ` | อัปเดตรหัสใหม่/กลับเข้าทำงาน ${updatedChangedCount} คน`
        : ` ข้อมูล ${(inRowCount ?? 0).toLocaleString()} รายการมีอยู่ในระบบแล้ว`)
  );

  return;
}
    // ==========================================
// Insert เฉพาะคนใหม่
// ==========================================
console.log(
  "🚀 กำลังจะ Insert เข้า Supabase:",
  newEmployees.length,
  "รายการ"
);

const chunkSize = 500;

let totalImported = 0;

for (
  let i = 0;
  i < newEmployees.length;
  i += chunkSize
) {

  const chunk =
    newEmployees.slice(
      i,
      i + chunkSize
    );

  console.log(
    "กำลัง Insert:",
    chunk.length,
    "รายการ"
  );

  const {
    error,
  } = await supabase
    .from("employees")
    .insert(chunk);

  if (error) {

    console.error(
      "❌ Supabase แจ้งเข้า Error:",
      error
    );

    throw new Error(
      `บันทึกแจ้งเข้าไม่สำเร็จ: ${error.message}`
    );
  }

  totalImported +=
    chunk.length;
}
    // ==========================================
    // โหลดข้อมูลใหม่
    // ==========================================
    console.log(
      "4️⃣ Insert สำเร็จ กำลังโหลดข้อมูลใหม่..."
    );

    const latestEmployees =
      await loadAllEmployees();

    setEmployees(
      latestEmployees
    );

    setRowCount(
      latestEmployees.length
    );

    // ==========================================
    // แจ้งผล
    // ==========================================
    
        setInSuccessMessage(
  `บันทึกเข้าสู่ฐานข้อมูลแล้ว เพิ่มพนักงานใหม่ ${totalImported.toLocaleString()} คน` +
    (updatedChangedCount > 0
      ? ` | อัปเดตรหัสใหม่/กลับเข้าทำงาน ${updatedChangedCount} คน`
      : "")
);

    setInPreview([]);
    setInNewCount(null);
    setInDuplicateCount(null);

    console.log(
      "✅ แจ้งเข้าสำเร็จ:",
      totalImported
    );

  } catch (err) {

    console.error(
      "💥 Import แจ้งเข้า Error:",
      err
    );

    setError(
      err instanceof Error
        ? err.message
        : "ไม่สามารถบันทึกข้อมูลแจ้งเข้าได้"
    );

  } finally {

    setIsInImporting(false);

    console.log(
      "🏁 จบการทำงานแจ้งเข้า"
    );
  }
};
// =========================
// อ่านไฟล์ Excel "แจ้งออก"
// =========================
const handleOutFileUpload = (
  event: React.ChangeEvent<HTMLInputElement>
) => {

  const file =
    event.target.files?.[0];

  if (!file) return;

  setOutFileName(
    file.name
  );

  setOutRowCount(null);
  setOutPreview([]);
  setShowOutList(false);
  setOutSuccessMessage("");
  setOutProblems([]);

  setError("");
  setSuccessMessage("");

  setIsOutLoading(true);

  const reader =
    new FileReader();

    reader.onload = async (e) => {

    try {

      const data =
        e.target?.result;

      if (!data) {
        throw new Error(
          "ไม่พบข้อมูลไฟล์แจ้งออก"
        );
      }

      const workbook =
        XLSX.read(data, {
          type: "array",
          cellDates: true,
        });

              // เลือกชีตที่ชื่อเป็นวันที่ล่าสุด
      // ถ้าไม่มีชีตไหนชื่อเป็นวันที่ ใช้ชีตแรก
      const outSheetName =
        [...workbook.SheetNames].sort(
          (a, b) => sheetNameToTime(b) - sheetNameToTime(a)
        )[0] ?? workbook.SheetNames[0];

      console.log(
        "ชีตทั้งหมดในไฟล์แจ้งออก:",
        JSON.stringify(
          workbook.SheetNames.map((name, index) => [
            name,
            workbook.Workbook?.Sheets?.[index]?.Hidden ?? "ไม่มีข้อมูล",
          ])
        )
      );

      console.log("ชีตที่อ่าน (แจ้งออก):", JSON.stringify(outSheetName));

      const worksheet = workbook.Sheets[outSheetName];

      const rows =
        XLSX.utils.sheet_to_json<any[]>(
          worksheet,
          {
            header: 1,
            defval: "",
          }
        );

      console.log(
        "========== แจ้งออก =========="
      );

      console.log(
        "ข้อมูลทั้งหมด:",
        rows
      );

      if (rows.length < 3) {

        throw new Error(
          "ไฟล์แจ้งออกไม่มีข้อมูล หรือรูปแบบไฟล์ไม่ถูกต้อง"
        );

      }

      // ==========================================
      // Header 2 แถว
      // ข้อมูลเริ่มแถวที่ 3
      // ==========================================
      const dataRows =
        rows.slice(2);

      const mappedRows =
  dataRows
    .filter((row) => {
      const employeeCode = String(row[1] ?? "")
        .trim()
        .replace(/\s+/g, "")
        .toUpperCase();

      const firstName = String(row[3] ?? "").trim();
      const lastName = String(row[4] ?? "").trim();

      // ตัดแถวหัวตาราง / ตัวอย่างข้อมูลออก
      if (
        employeeCode === "EMP.CODE" ||
        employeeCode === "รหัสพนักงาน"
      ) {
        return false;
      }

      return (
        employeeCode !== "" &&
        (firstName !== "" || lastName !== "")
      );
    })
          .map((row) => {

            return {

              employee_code:
                String(row[1] ?? "")
                  .trim()
                  .replace(/\s+/g, "")
                  .toUpperCase(),

              title:
                String(row[2] ?? "")
                  .trim(),

              first_name:
                String(row[3] ?? "")
                  .trim(),

              last_name:
                String(row[4] ?? "")
                  .trim(),

              resignation_date:
                formatExcelDate(
                  row[5]
                ),

              plan:
                row[6] !== ""
                  ? Number(row[6])
                  : null,

              remark:
                String(row[7] ?? "")
                  .trim(),

            };

          });

      console.log(
        "ข้อมูลแจ้งออกหลัง Mapping:",
        mappedRows
      );


      // ตรวจเทียบกับฐานข้อมูล (ดูอย่างเดียว ไม่เขียนข้อมูล)
      const dbForOut = await loadAllEmployees();
      const outLookup = buildLookup(dbForOut);

      const outCheck = mappedRows.map((row) => {
        const m = findEmployee(row, outLookup);
        return { row, by: m.by, warning: m.warning, db: m.employee };
      });

      console.log(
        "แจ้งออก ตรวจเทียบ DB:",
        JSON.stringify({
          total: outCheck.length,
          byCode: outCheck.filter((c) => c.by === "code").length,
          byName: outCheck.filter((c) => c.by === "name").length,
          notFound: outCheck.filter((c) => c.by === "none").length,
          alreadyResigned: outCheck.filter(
            (c) => c.by === "code" && c.db?.status === "ลาออก"
          ).length,
        })
      );

      console.log(
        "แจ้งออก ที่ต้องตรวจ:",
        JSON.stringify(
          outCheck
            .filter((c) => c.by !== "code")
            .map((c) => [
              c.row.employee_code,
              c.row.first_name,
              c.row.last_name,
              c.by,
              c.warning ?? "",
            ])
        )
      );

            // ==========================================
      // แยกกลุ่ม: พร้อมแจ้งออก / กลุ่มที่ต้องตรวจ
      // ==========================================
      const okRows: any[] = [];
      const problemRows: OutProblem[] = [];
      const seenCodes = new Set<string>();

      for (const c of outCheck) {
        const code = normCode(c.row.employee_code);

        // รหัสซ้ำกันเองในไฟล์ (เก็บแถวแรก แถวถัดไปให้ตรวจ)
        if (seenCodes.has(code)) {
          problemRows.push({
            row: c.row,
            group: "dupInFile",
            reason: "รหัสนี้ซ้ำกันเองในไฟล์",
          });
          continue;
        }

        seenCodes.add(code);

        if (c.by === "code") {
          // รหัสตรง + ชื่อตรง
          if (c.db?.status === "ลาออก") {
            const fileDate = c.row.resignation_date;
            const dbDate = c.db.resignation_date;

            if (fileDate && dbDate && fileDate !== dbDate) {
              problemRows.push({
                row: c.row,
                group: "dateDiff",
                reason: `วันที่ในไฟล์ ${fileDate} ไม่ตรงกับในระบบ ${dbDate}`,
              });
            } else {
              problemRows.push({
                row: c.row,
                group: "resigned",
                reason: `ลาออกในระบบแล้ว (${dbDate ?? "ไม่มีวันที่"})`,
              });
            }
          } else if (!c.row.resignation_date) {
            problemRows.push({
              row: c.row,
              group: "noDate",
              reason: "ไฟล์ไม่มีวันที่ลาออก",
            });
          } else {
            okRows.push(c.row);
          }
        } else if (c.by === "name") {
          problemRows.push({
            row: c.row,
            group: "nameOnly",
            reason: "ไม่พบรหัส แต่ชื่อตรงกับ 1 คนในระบบ",
          });
        } else if (c.warning?.startsWith("รหัสตรงแต่ชื่อไม่ตรง")) {
          problemRows.push({
            row: c.row,
            group: "codeNameDiff",
            reason: c.warning,
          });
        } else {
          problemRows.push({
            row: c.row,
            group: "notFound",
            reason: c.warning ?? "ไม่พบในฐานข้อมูล",
          });
        }
      }

      console.log(
        "แจ้งออก แยกกลุ่ม:",
        JSON.stringify({
          ok: okRows.length,
          problems: problemRows.length,
        })
      );

      console.log(
        "แจ้งออก รายกลุ่ม:",
        JSON.stringify(
          problemRows.reduce<Record<string, number>>((acc, p) => {
            acc[p.group] = (acc[p.group] ?? 0) + 1;
            return acc;
          }, {})
        )
      );

      setOutPreview(okRows);             // เฉพาะคนที่จะบันทึกจริง
      setOutProblems(problemRows);       // ทุกกลุ่มที่ต้องดู/ตรวจ
      setOutRowCount(mappedRows.length); // จำนวนทั้งหมดในไฟล์

    } catch (err) {

      console.error(err);

      setError(
        err instanceof Error
          ? err.message
          : "ไม่สามารถอ่านไฟล์แจ้งออกได้"
      );

    } finally {

      setIsOutLoading(false);

    }

  };

  reader.onerror = () => {

    setError(
      "เกิดข้อผิดพลาดในการอ่านไฟล์แจ้งออก"
    );

    setIsOutLoading(false);

  };

  reader.readAsArrayBuffer(file);
};
// =========================
// บันทึก "แจ้งออก"
// เปลี่ยนสถานะ ไม่ลบข้อมูล
// =========================
const handleImportOutToSupabase =
  async () => {

    if (outPreview.length === 0) {

      setError(
        "ยังไม่มีข้อมูลแจ้งออกสำหรับดำเนินการ"
      );

      return;
    }

    setIsOutImporting(true);

    setError("");
    setSuccessMessage("");
    setOutSuccessMessage("");

    try {

      // ==========================================
      // โหลดข้อมูลล่าสุดจาก Supabase
      // ==========================================
      const currentEmployees =
        await loadAllEmployees();

            const lookup = buildLookup(currentEmployees);

      let updatedCount = 0;

      let notFoundCount = 0;

      const notFoundEmployees:
        string[] = [];

      // ==========================================
      // วนข้อมูลแจ้งออก
      // ==========================================
      for (
        const row of outPreview
      ) {

        const employeeCode = normCode(row.employee_code);

            // จับคู่ด้วย findEmployee เอาเฉพาะรหัสตรง + ชื่อตรง
      const match = findEmployee(row, lookup);
      const employee = match.by === "code" ? match.employee : null;

        // ========================================
        // ถ้าไม่พบรหัส
        // ========================================
        if (!employee) {

          notFoundCount++;

        notFoundEmployees.push(
          `${employeeCode} | ${row.title ?? ""}${row.first_name ?? ""} ${row.last_name ?? ""}` +
            (match.warning ? ` | ${match.warning}` : "")
        );

          continue;
        }

        // ========================================
        // UPDATE คนเดิม
        // ไม่ DELETE
        // ========================================
        const {
          error,
        } = await supabase
          .from("employees")
                    .update({

            status:
              "ลาออก",

            resignation_date:
              row.resignation_date,

            remark: row.remark
              ? [employee.remark, row.remark].filter(Boolean).join(" | ")
              : employee.remark,

            updated_at:
              new Date().toISOString(),

          })
          .eq(
            "id",
            employee.id
          );

        if (error) {

          throw new Error(
            `อัปเดตพนักงาน ${employeeCode} ไม่สำเร็จ: ${error.message}`
          );

        }

        updatedCount++;

      }

      // ==========================================
      // โหลดข้อมูลใหม่
      // ==========================================
      const latestEmployees =
        await loadAllEmployees();

      setEmployees(
        latestEmployees
      );

      setRowCount(
        latestEmployees.length
      );

      // ==========================================
      // สรุปผล
      // ==========================================
      let message =
        `แจ้งออกสำเร็จ ${updatedCount.toLocaleString()} รายการ`;

      if (
        notFoundCount > 0
      ) {

        message +=
          ` | ไม่พบรหัสพนักงาน ${notFoundCount.toLocaleString()} รายการ`;

      }

        setOutSuccessMessage(message);

      // ==========================================
      // Debug คนที่หาไม่เจอ
      // ==========================================
      if (
        notFoundEmployees.length > 0
      ) {

        console.warn(
          "ไม่พบพนักงานจากไฟล์แจ้งออก:",
          notFoundEmployees
        );

      }

      setOutPreview([]);

    } catch (err) {

      console.error(
        "Import แจ้งออก Error:",
        err
      );

      setError(
        err instanceof Error
          ? err.message
          : "ไม่สามารถดำเนินการแจ้งออกได้"
      );

    } finally {

      setIsOutImporting(false);

    }
  };
  // =========================
// อ่านไฟล์ Excel รายงานเคลม
// =========================
const handleClaimFileUpload = (
  event: React.ChangeEvent<HTMLInputElement>
) => {
  const file = event.target.files?.[0];

  if (!file) return;

  setClaimFileName(file.name);
  setClaimRowCount(null);
  setClaimPreview([]);
  setError("");
  setSuccessMessage("");
  setIsClaimLoading(true);

  const reader = new FileReader();

  reader.onload = async (e) => {
    try {
      const data = e.target?.result;

      if (!data) {
        throw new Error("ไม่พบข้อมูลไฟล์รายงานเคลม");
      }

      const workbook = XLSX.read(data, {
        type: "array",
        cellDates: true,
      });

            // เลือกชีตเคลมอัตโนมัติ จากชื่อชีต + หัวตารางในชีต
      const normSheet = (s: string) =>
        String(s ?? "").toLowerCase().replace(/\s+/g, " ").trim();

      const CLAIM_NAME_WORDS = ["claim", "เคลม", "health"];
      const CLAIM_HEADER_WORDS = [
        "ความสัมพันธ์",
        "เลขที่รับเรื่อง",
        "เลขที่อ้างอิง",
        "ประเภทการเคลม",
        "สถานพยาบาล",
        "บริษัทฯจ่าย",
        "บริษัทจ่าย",
        "เลขบัตร",
        "ใบเสร็จ",
        "ชื่อ-สกุล",
      ];

      const visibleSheets = workbook.SheetNames.filter(
        (_name, index) =>
          (workbook.Workbook?.Sheets?.[index]?.Hidden ?? 0) === 0
      );

      const scored = visibleSheets.map((name) => {
        const sheetRows = XLSX.utils.sheet_to_json<any[]>(
          workbook.Sheets[name],
          { header: 1, defval: "" }
        );

        // รวมข้อความ 20 แถวแรกไว้หาคำหัวตาราง
        const headText = normSheet(
          sheetRows
            .slice(0, 20)
            .map((r) => r.join(" "))
            .join(" ")
        );

        const headerHits = CLAIM_HEADER_WORDS.filter((w) =>
          headText.includes(normSheet(w))
        ).length;

        const nameHits = CLAIM_NAME_WORDS.filter((w) =>
          normSheet(name).includes(w)
        ).length;

        return {
          name,
          headerHits,
          rowCount: sheetRows.length,
          score: headerHits * 10 + nameHits * 3,
        };
      });

      scored.sort(
        (a, b) => b.score - a.score || b.rowCount - a.rowCount
      );

      console.log("ชีตเคลม ให้คะแนน:", JSON.stringify(scored));

      const best = scored[0];

      // มีชีตเดียว ใช้เลย ไม่ต้องตรวจหัวตาราง
      // มีหลายชีต ต้องมีหัวตารางเคลมอย่างน้อย 1 คำ
      if (!best || (visibleSheets.length > 1 && best.headerHits === 0)) {
        throw new Error(
          `ไม่พบชีตที่เป็นข้อมูลเคลมในไฟล์ (ชีตที่มี: ${workbook.SheetNames.join(", ")})`
        );
      }

      const claimSheetName = best.name;

      console.log("ชีตที่เลือก (เคลม):", JSON.stringify(claimSheetName));

      const worksheet = workbook.Sheets[claimSheetName];

      const rows = XLSX.utils.sheet_to_json<any[]>(worksheet, {
        header: 1,
        defval: "",
      });

      console.log("ข้อมูล Claims จาก Excel:", rows);

      if (rows.length < 2) {
        throw new Error(
          "ไฟล์รายงานเคลมไม่มีข้อมูล หรือรูปแบบไฟล์ไม่ถูกต้อง"
        );
      }
      console.log("========== CLAIM ROW DEBUG ==========");

console.log(
  rows[1].map((value: any, index: number) => ({
    column: index,
    value: value,
  }))
);

console.log("====================================");

// ==================================================
// Mapping รายงานเคลม
//
// 0  = Policy Name
// 1  = เลขที่อ้างอิง
// 2  = สถานะ
// 3  = ความสัมพันธ์
// 4  = ชื่อ-สกุล
// 5  = ช่องข้อมูลอื่น
// 6  = ช่องข้อมูลอื่น
// 7  = ช่องข้อมูลอื่น
// 8  = ช่องข้อมูลอื่น
// 9  = ประเภทการเคลม
// 10 = วันเข้า รพ.
// 11 = ใบเสร็จ
// 12 = บริษัทฯจ่าย
// 13 = สถานพยาบาล
// 14 = ประเภทการจ่าย
// 15 = เลขที่รับเรื่อง
// 16 = สังกัด
// 17 = เลขบัตรประชาชน
//
// ==================================================

      const dataRows = rows.slice(1);

      const mappedClaims: Claim[] = dataRows
  .filter((row) => {
    const employeeName = String(row[4] ?? "").trim();

    return (
      employeeName !== "" &&
      employeeName !== "ชื่อ-สกุล" &&
      employeeName !== "ชื่อผู้เอาประกันภัย"
    );
  })
  .map((row) => ({
    policy_name: String(row[0] ?? "").trim(),

    reference_no: String(row[1] ?? "").trim(),

    claim_status: String(row[2] ?? "").trim(),

    relationship: String(row[3] ?? "").trim(),

    employee_name: String(row[4] ?? "").trim(),

    // เลขบัตรประชาชน
    // Excel คอลัมน์ R = index 17
    id_card: String(row[17] ?? "")
      .trim()
      .replace(/\D/g, ""),

    // ตอนนี้ยังไม่ใช้แผนจากไฟล์เคลม
    plan: null,

    // ประเภทการเคลม
    claim_type: String(row[9] ?? "").trim(),

    // วันเข้าโรงพยาบาล
    hospital_date: formatExcelDate(row[10]),

        // จำนวนเงินตามใบเสร็จ
    receipt_amount: parseAmount(row[11]),

    // จำนวนเงินที่บริษัทประกันจ่าย
    insurance_paid: parseAmount(row[12]),

    // สถานพยาบาล
    hospital_name: String(row[13] ?? "").trim(),

    // ประเภทการจ่าย
    payment_type: String(row[14] ?? "").trim(),

    // เลขที่รับเรื่อง
    claim_no: String(row[15] ?? "").trim(),

    // สังกัด
    affiliation: String(row[16] ?? "").trim(),
  }));

      console.log("Claims หลัง Mapping:", mappedClaims);
      console.log(
  "========== CLAIM ID CARD DEBUG =========="
);

console.log(
  mappedClaims.slice(0, 20).map((claim) => ({
    ชื่อ: claim.employee_name,
    ความสัมพันธ์: claim.relationship,
    เลขบัตรประชาชน: claim.id_card,
    จำนวนหลัก: claim.id_card.length,
  }))
);

console.log(
  "========================================="
);

            const keyCount = new Map<string, number>();
      const withKeys = mappedClaims.map((c) => {
        const base = [
          c.reference_no,
          c.claim_no,
          c.id_card,
          c.claim_type,
          c.hospital_date ?? "",
        ].join("|");
        const n = (keyCount.get(base) ?? 0) + 1;
        keyCount.set(base, n);
        return { ...c, claim_key: `${base}|${n}` };
      });

      setClaimPreview(withKeys);
      setClaimRowCount(withKeys.length);

    } catch (err) {
      console.error(err);

      setError(
        err instanceof Error
          ? err.message
          : "ไม่สามารถอ่านไฟล์รายงานเคลมได้"
      );
    } finally {
      setIsClaimLoading(false);
    }
  };
  

  reader.onerror = () => {
    setError("เกิดข้อผิดพลาดในการอ่านไฟล์รายงานเคลม");
    setIsClaimLoading(false);
  };

  reader.readAsArrayBuffer(file);
};


// =========================
// บันทึก Claims ลง Supabase
// =========================
const handleImportClaimsToSupabase = async () => {
  if (claimPreview.length === 0) {
    setError("ยังไม่มีข้อมูลรายงานเคลมสำหรับนำเข้า");
    return;
  }

  setIsClaimImporting(true);
  setError("");
  setSuccessMessage("");

  try {
    const chunkSize = 500;

    let totalImported = 0;

    // ==========================================
    // บันทึก Claims ทีละ 500 รายการ
    // ==========================================
    for (
      let i = 0;
      i < claimPreview.length;
      i += chunkSize
    ) {
      const chunk = claimPreview.slice(
        i,
        i + chunkSize
      );

      console.log(
        `กำลังบันทึก Claim ${
          i + 1
        } - ${
          i + chunk.length
        } / ${
          claimPreview.length
        }`
      );

            const { error } = await supabase
        .from("claims")
        .upsert(chunk, { onConflict: "claim_key" });

      if (error) {
        console.error(
          "Supabase Claims Error:",
          error
        );

        throw new Error(
          `บันทึกข้อมูลเคลมไม่สำเร็จ: ${error.message}`
        );
      }

      totalImported += chunk.length;
    }

    // ==========================================
    // โหลด Claims ใหม่ทั้งหมดจาก Supabase
    // ==========================================
    console.log(
      "กำลังโหลด Claims ใหม่..."
    );

    const latestClaims =
      await loadAllClaims();

    console.log(
      "Claims ตัวอย่าง 10 รายการ:",
      latestClaims.slice(0, 10)
    );
    console.log(
      "โหลด Claims ใหม่สำเร็จ:",
      latestClaims.length
    );

        setClaims(latestClaims);

    setSuccessMessage(
      `นำเข้ารายงานเคลมสำเร็จ ${totalImported.toLocaleString()} รายการ`
    );

    setClaimPreview([]);

    setClaimRowCount(null);

  } catch (err) {
    console.error(
  "Import Claims Error:",
  err
);

    setError(
      err instanceof Error
        ? err.message
        : "ไม่สามารถบันทึกรายงานเคลมได้"
    );

  } finally {

    setIsClaimImporting(false);

  }
};
const handleInsuranceFileUpload = (
  event: React.ChangeEvent<HTMLInputElement>
) => {
  const file = event.target.files?.[0];

  if (!file) return;

  console.log("📄 ไฟล์ประกันส่งกลับ:", file.name);

  setInsuranceFileName(file.name);
  setInsuranceRowCount(null);
  setIsInsuranceLoading(true);
  setError("");
  setSuccessMessage("");

  const reader = new FileReader();

  reader.onload = async (e) => {
    try {
      const data = e.target?.result;

      if (!data) {
        throw new Error("ไม่พบข้อมูลไฟล์ประกันส่งกลับ");
      }

      const workbook = XLSX.read(data, {
        type: "array",
        cellDates: true,
      });

            const firstSheetName =
        workbook.SheetNames.find(
          (_name, index) =>
            (workbook.Workbook?.Sheets?.[index]?.Hidden ?? 0) === 0
        ) ?? workbook.SheetNames[0];

      console.log("ชีตที่อ่าน (ประกันส่งกลับ):", JSON.stringify(firstSheetName));
      const worksheet = workbook.Sheets[firstSheetName];

      const rows = XLSX.utils.sheet_to_json<any[]>(worksheet, {
        header: 1,
        defval: "",
      });

      console.log("========== ประกันส่งกลับ ==========");
      console.log("ข้อมูลทั้งหมด:", rows);

// ==========================================
// แปลงข้อมูลไฟล์ประกันส่งกลับ
// ==========================================

// หาแถวหัวตาราง (คอลัมน์ C = "รหัสพนักงาน") แล้วเริ่มอ่านถัดจากนั้น
const insHeaderIndex = rows.findIndex((row) =>
  String(row[2] ?? "").includes("รหัสพนักงาน")
);

if (insHeaderIndex === -1) {
  throw new Error("ไม่พบแถวหัวตาราง 'รหัสพนักงาน' ในไฟล์ประกันส่งกลับ");
}

const dataRows = rows
  .slice(insHeaderIndex + 1)
  .filter((row) => {
    const code = String(row[2] ?? "").trim();
    const first = String(row[4] ?? "").trim();
    const last = String(row[5] ?? "").trim();
    return code !== "" && (first !== "" || last !== "");
  });

console.log(
  "จำนวนข้อมูลประกันส่งกลับ:",
  dataRows.length
);


console.log(
  "ตัวอย่างวันที่ (ประกันส่งกลับ):",
  JSON.stringify({
    textDates: dataRows.filter(
      (r) => typeof r[20] === "string" && String(r[20]).trim() !== ""
    ).length,
    sample: dataRows
      .slice(0, 5)
      .map((r) => [
        String(r[20]),
        typeof r[20],
        formatExcelDate(r[20]),
        formatExcelDate(r[21]),
      ]),
  })
);

const mappedEmployees: Employee[] =
  dataRows.map((row) => {

    const employeeCode =
      String(row[2] ?? "")
        .trim()
        .replace(/\s+/g, "")
        .toUpperCase();

    return {
      employee_key: "",
      
      employee_code:
        employeeCode,

      vendor:
        String(row[1] ?? "").trim(),

      branch:
        String(row[12] ?? "").trim(),

      title:
        String(row[3] ?? "").trim(),

      first_name:
        String(row[4] ?? "").trim(),

      last_name:
        String(row[5] ?? "").trim(),

      gender:
        String(row[8] ?? "").trim(),

      date_of_birth:
        formatExcelDate(row[9]),

      id_card:
  String(row[11] ?? "")
    .trim()
    .replace(/\D/g, ""),

      employment_date:
        null,

      effective_date:
        formatExcelDate(row[20]),

      plan:
        row[23] !== "" &&
        row[23] !== null &&
        row[23] !== undefined
          ? Number(row[23])
          : null,

      insurance_type:
        "",

      department:
        String(row[15] ?? "").trim(),

      bank_account:
        String(row[19] ?? "").trim(),

      bank_name:
        String(row[18] ?? "").trim(),

      phone:
        String(row[25] ?? "").trim(),

      remark:
        "",

      resignation_date:
        formatExcelDate(row[21]),

      status:
        String(row[6] ?? "").trim(),

      insurance_card_no:
        String(row[27] ?? "").trim(),

      life_plan:
        String(row[22] ?? "").trim(),
    };
  });

console.log(
  "ข้อมูลประกันส่งกลับที่ Mapping แล้ว:",
  mappedEmployees
);

setInsuranceRowCount(
  mappedEmployees.length
);

// ==========================================
// จับคู่ด้วยเลขบัตรประชาชนเป็นหลัก
// ==========================================
const dbForInsurance = await loadAllEmployees();
const matchLookup = buildLookup(dbForInsurance);

const matched: Employee[] = [];
const unmatched: Employee[] = [];
const needCheck: { row: Employee; db: Employee }[] = [];

mappedEmployees.forEach((insuranceEmployee) => {
  const id = normId(insuranceEmployee.id_card);
  const masterEmployee = id ? matchLookup.byId.get(id) : undefined;

  // ไม่เจอเลขบัตรในฐานข้อมูล
  if (!masterEmployee) {
    unmatched.push(insuranceEmployee);
    return;
  }

  // เลขบัตรตรง แต่รหัสพนักงานต่างกัน ไม่บันทึกอัตโนมัติ
  if (
    normCode(masterEmployee.employee_code) !==
    normCode(insuranceEmployee.employee_code)
  ) {
    needCheck.push({ row: insuranceEmployee, db: masterEmployee });
    return;
  }

  matched.push({
    ...masterEmployee,

    // วันที่มีผลประกันจริงจากไฟล์ประกันส่งกลับ
    effective_date: insuranceEmployee.effective_date,

    // ประเภทปีประกัน ดูจากวันที่มีผลจริง
    insurance_type: getInsuranceType(insuranceEmployee.effective_date),

    plan: insuranceEmployee.plan,

    insurance_card_no: insuranceEmployee.insurance_card_no,

    life_plan: insuranceEmployee.life_plan,
  });
});

console.log(
  "========== ผลการจับคู่ =========="
);

console.log(
  "จับคู่ได้:",
  matched.length
);

console.log(
  "จับคู่ไม่ได้:",
  unmatched.length
);

console.log(
  "รายการที่จับคู่ไม่ได้:",
  unmatched
);


console.log(
  "ต้องตรวจ (เลขบัตรตรงแต่รหัสต่าง):",
  needCheck.length,
  JSON.stringify(
    needCheck
      .slice(0, 10)
      .map((c) => [
        c.row.employee_code,
        c.row.first_name,
        c.db.employee_code,
        c.db.status,
      ])
  )
);

const idCounts = new Map<string, number>();
mappedEmployees.forEach((e) => {
  const id = normId(e.id_card);
  if (id) idCounts.set(id, (idCounts.get(id) ?? 0) + 1);
});
console.log(
  "เลขบัตรซ้ำกันในไฟล์ประกัน:",
  Array.from(idCounts.values()).filter((n) => n > 1).length
);


const dupIds = Array.from(idCounts.entries())
  .filter(([, n]) => n > 1)
  .map(([id]) => id);

console.log(
  "แถวที่เลขบัตรซ้ำ:",
  JSON.stringify(
    mappedEmployees
      .filter((e) => dupIds.includes(normId(e.id_card)))
      .map((e) => [
        e.employee_code,
        e.first_name,
        e.last_name,
        e.effective_date,
        e.plan,
        e.insurance_card_no,
      ])
  )
);

// ตรวจค่าที่จะบันทึก (เฉพาะกลุ่มพร้อมบันทึก)
const countBy = (values: (string | number | null)[]) => {
  const result: Record<string, number> = {};
  values.forEach((v) => {
    const key = v === null || v === "" ? "(ว่าง)" : String(v);
    result[key] = (result[key] ?? 0) + 1;
  });
  return result;
};

const effCount = countBy(matched.map((m) => m.effective_date));

console.log(
  "วันที่มีผล (ที่จะบันทึก):",
  JSON.stringify({
    distinct: Object.keys(effCount).length,
    empty: effCount["(ว่าง)"] ?? 0,
    top: Object.entries(effCount)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 12),
  })
);

console.log(
  "แผน (ที่จะบันทึก):",
  JSON.stringify(countBy(matched.map((m) => m.plan)))
);

console.log(
  "ไม่มีหมายเลขบัตรประกัน:",
  matched.filter((m) => !m.insurance_card_no).length
);

// ตรวจเทียบแบบละเอียด (ดูอย่างเดียว ไม่เขียนข้อมูล)
const insLookup = buildLookup(dbForInsurance);

const sameNameAs = (a: Employee, b: Employee) =>
  nameKey(a.first_name, a.last_name) === nameKey(b.first_name, b.last_name);

const cats = {
  codeAndNameOk: [] as Employee[],
  codeButNameDiffers: [] as [Employee, Employee][],
  idOnly: [] as [Employee, Employee][],
  nameOnly: [] as [Employee, Employee][],
  notFound: [] as Employee[],
};

for (const row of mappedEmployees) {
  const byCode = insLookup.byCode.get(normCode(row.employee_code));
  const byId = row.id_card
    ? insLookup.byId.get(normId(row.id_card))
    : undefined;

  if (byCode && sameNameAs(byCode, row)) {
    cats.codeAndNameOk.push(row);
  } else if (byCode) {
    cats.codeButNameDiffers.push([row, byCode]);
  } else if (byId) {
    cats.idOnly.push([row, byId]);
  } else {
    const hits =
      insLookup.byName.get(nameKey(row.first_name, row.last_name)) ?? [];
    if (hits.length === 1) {
      cats.nameOnly.push([row, hits[0]]);
    } else {
      cats.notFound.push(row);
    }
  }
}

console.log(
  "ประกันส่งกลับ ตรวจเทียบ:",
  JSON.stringify({
    total: mappedEmployees.length,
    codeAndNameOk: cats.codeAndNameOk.length,
    codeButNameDiffers: cats.codeButNameDiffers.length,
    idOnly: cats.idOnly.length,
    nameOnly: cats.nameOnly.length,
    notFound: cats.notFound.length,
  })
);

console.log(
  "ตัวอย่าง รหัสตรงแต่ชื่อต่าง:",
  JSON.stringify(
    cats.codeButNameDiffers
      .slice(0, 10)
      .map(([r, d]) => [r.employee_code, r.first_name, d.first_name])
  )
);

console.log(
  "ตัวอย่าง รหัสไม่เจอแต่เลขบัตรเจอ:",
  JSON.stringify(
    cats.idOnly
      .slice(0, 10)
      .map(([r, d]) => [r.employee_code, r.first_name, d.employee_code, d.status])
  )
);

console.log(
  "ตัวอย่าง ไม่เจอเลย:",
  JSON.stringify(
    cats.notFound
      .slice(0, 10)
      .map((r) => [r.employee_code, r.first_name, r.last_name, r.status])
  )
);


console.log(
  "รหัสตรงแต่ชื่อต่าง เทียบเลขบัตร (4 หลักท้าย ไฟล์/DB):",
  JSON.stringify(
    cats.codeButNameDiffers.map(([r, d]) => [
      r.employee_code,
      `${r.first_name} ${r.last_name}`,
      `${d.first_name} ${d.last_name}`,
      normId(r.id_card).slice(-4),
      normId(d.id_card).slice(-4),
      insLookup.byId.has(normId(r.id_card)) ? "เลขบัตรไฟล์มีใน DB" : "ไม่มีใน DB",
    ])
  )
);

const statusCount: Record<string, number> = {};
for (const row of mappedEmployees) {
  const s = String(row.status ?? "").trim() || "(ว่าง)";
  statusCount[s] = (statusCount[s] ?? 0) + 1;
}
console.log("ประกันส่งกลับ นับตามสถานะ (คอลัมน์ G):", JSON.stringify(statusCount));

const junkRows = cats.notFound.filter((r) => !r.first_name && !r.last_name);
const realNotFound = cats.notFound.filter((r) => r.first_name || r.last_name);

console.log(
  "ไม่เจอเลย แยกประเภท:",
  JSON.stringify({
    junkRows: junkRows.length,
    realPeople: realNotFound.length,
    realWithIdCard: realNotFound.filter((r) => r.id_card).length,
    realNoIdCard: realNotFound.filter((r) => !r.id_card).length,
  })
);

console.log(
  "ตัวอย่าง รหัสตรงแต่ชื่อต่าง พร้อมสถานะ:",
  JSON.stringify(
    cats.codeButNameDiffers
      .slice(0, 15)
      .map(([r, d]) => [
        r.employee_code,
        r.first_name,
        r.last_name,
        r.status,
        d.first_name,
        d.last_name,
      ])
  )
);

setInsuranceMatched(matched);
setInsuranceUnmatched(unmatched);

    } catch (err) {
      console.error(
        "Insurance File Error:",
        err
      );

      setError(
        err instanceof Error
          ? err.message
          : "ไม่สามารถอ่านไฟล์ประกันส่งกลับได้"
      );
    } finally {
      setIsInsuranceLoading(false);
    }
  };

  reader.onerror = () => {
    setError("ไม่สามารถอ่านไฟล์ประกันส่งกลับได้");
    setIsInsuranceLoading(false);
  };

    reader.readAsArrayBuffer(file);
};

// ==========================================
// บันทึกข้อมูลประกันส่งกลับลง Supabase
// ==========================================

const handleImportInsuranceToSupabase = async () => {
  if (insuranceMatched.length === 0) {
    setError("ไม่มีข้อมูลประกันส่งกลับที่จับคู่ได้");
    return;
  }

  setIsInsuranceImporting(true);
  setError("");
  setSuccessMessage("");

  try {
    const chunkSize = 100;
    let successCount = 0;

    console.log(
      "🚀 เริ่มบันทึกข้อมูลประกันส่งกลับ:",
      insuranceMatched.length,
      "รายการ"
    );

    for (let i = 0; i < insuranceMatched.length; i += chunkSize) {
      const chunk = insuranceMatched.slice(i, i + chunkSize);

      const uniqueChunk = Array.from(
        new Map(
          chunk
            .filter((employee) => employee.id)
            .map((employee) => [employee.id, employee])
        ).values()
      );

      console.log(
        `💾 กำลังบันทึก ${Math.min(
          i + chunkSize,
          insuranceMatched.length
        )} / ${insuranceMatched.length}`
      );

      await Promise.all(
        uniqueChunk.map(async (employee) => {
          const { error } = await supabase
            .from("employees")
            .update({
          effective_date: employee.effective_date,
          insurance_type: employee.insurance_type,
          plan: employee.plan,
          insurance_card_no: employee.insurance_card_no,
          life_plan: employee.life_plan,
          updated_at: new Date().toISOString(),
        })
            .eq("id", employee.id);

          if (error) {
            throw error;
          }
        })
      );

      successCount += uniqueChunk.length;
    }

    console.log(
      "✅ บันทึกข้อมูลประกันส่งกลับสำเร็จ:",
      successCount,
      "รายการ"
    );

    const latestEmployees = await loadAllEmployees();
    setEmployees(latestEmployees);

    setSuccessMessage(
      `บันทึกข้อมูลประกันส่งกลับสำเร็จ ${successCount.toLocaleString()} รายการ`
    );
  } catch (err: any) {
    console.error("❌ Insurance Import Error:", err);
    console.error("❌ Error JSON:", JSON.stringify(err, null, 2));
    console.error("❌ Error message:", err?.message);
    console.error("❌ Error details:", err?.details);
    console.error("❌ Error hint:", err?.hint);
    console.error("❌ Error code:", err?.code);

    setError(
      err?.message ||
        err?.details ||
        err?.hint ||
        "ไม่สามารถบันทึกข้อมูลประกันส่งกลับได้"
    );
  } finally {
    setIsInsuranceImporting(false);
  }
};

const handleImportToSupabase = async () => {
  if (employees.length === 0) {
    setError("ยังไม่มีข้อมูลสำหรับนำเข้า");
    return;
  }

  setIsImporting(true);
  setError("");
  setSuccessMessage("");

  try {

    // ==========================================
    // 1. สร้าง employee_key ใหม่ให้ทุกคน
    // ==========================================
    const cleanedEmployees = employees
  .map((employee) => {
    const vendor = String(employee.vendor ?? "").trim();

    const employeeCode = String(
      employee.employee_code ?? ""
    )
      .trim()
      .replace(/\s+/g, "")
      .toUpperCase();

    const firstName = String(
      employee.first_name ?? ""
    ).trim();

    const lastName = String(
      employee.last_name ?? ""
    ).trim();

    const idCard = String(
      employee.id_card ?? ""
    ).trim();

    return {
      ...employee,

      employee_code: employeeCode,

  employee_key: createEmployeeKey(
  vendor,
  employeeCode,
  firstName,
  lastName,
  idCard
),
    };
  })
  .filter(
    (employee) =>
      employee.employee_key !== ""
  );


    // ==========================================
    // 2. ตรวจสอบ duplicate employee_key
    // ==========================================
    const employeeMap =
      new Map<string, Employee>();

    const duplicateKeys: string[] = [];

    for (const employee of cleanedEmployees) {

      const key =
        employee.employee_key;

      if (employeeMap.has(key)) {
        duplicateKeys.push(key);
      }

      employeeMap.set(
        key,
        employee
      );
    }


    // ==========================================
    // 3. ถ้ามีชื่อซ้ำในบริษัทเดียวกัน
    // ==========================================
    if (duplicateKeys.length > 0) {

      const uniqueDuplicateKeys =
        [...new Set(duplicateKeys)];

    // ==========================================
    // 3. ถ้ามีชื่อซ้ำในบริษัทเดียวกัน
    // ==========================================
const duplicateEmployees = uniqueDuplicateKeys.map((key) => {
  const duplicateRows = cleanedEmployees.filter(
    (employee) => employee.employee_key === key
  );

  return duplicateRows
    .map(
      (employee) =>
        `${employee.vendor} | ${employee.employee_code} | ${employee.title}${employee.first_name} ${employee.last_name} | บัตรประชาชน: ${employee.id_card}`
    )
    .join(" / ");
});

throw new Error(
  `พบข้อมูลซ้ำ ${uniqueDuplicateKeys.length} รายการ:\n\n${duplicateEmployees.join(
    "\n"
  )}`
);

    }


    // ==========================================
    // 4. ข้อมูลที่จะบันทึก
    // ==========================================
    const uniqueEmployees =
      Array.from(
        employeeMap.values()
      );


    console.log(
      "================================="
    );

    console.log(
      "จำนวนจาก Excel:",
      employees.length
    );

    console.log(
      "จำนวนหลัง Clean:",
      cleanedEmployees.length
    );

    console.log(
      "จำนวนที่จะบันทึก:",
      uniqueEmployees.length
    );

    console.log(
      "================================="
    );

    // ==========================================
// 5. บันทึกข้อมูลเป็นชุดละ 500
// ==========================================

// โหลดข้อมูลที่มีอยู่ใน Database ครั้งเดียว
const currentEmployees =
  await loadAllEmployees();

console.log(
  "📦 ข้อมูลเดิมใน Database:",
  currentEmployees.length,
  "รายการ"
);

// สร้าง Map จาก employee_key
const currentEmployeeMap =
  new Map<string, Employee>();

currentEmployees.forEach((employee) => {
  if (employee.employee_key) {
    currentEmployeeMap.set(
      employee.employee_key,
      employee
    );
  }
});

// เตรียมข้อมูลสำหรับ Upsert
const rowsToUpsert =
  uniqueEmployees.map((employee) => {

    const existingEmployee =
      employee.employee_key
        ? currentEmployeeMap.get(
            employee.employee_key
          )
        : undefined;

    return {
      // ถ้ามีข้อมูลเดิม → ใช้ id เดิม
      ...(existingEmployee?.id
        ? {
            id: existingEmployee.id,
          }
        : {}),

      employee_key:
        employee.employee_key,

      employee_code:
        employee.employee_code,

      vendor:
        employee.vendor,

      branch:
        employee.branch,

      title:
        employee.title,

      first_name:
        employee.first_name,

      last_name:
        employee.last_name,

      gender:
        employee.gender,

      date_of_birth:
        employee.date_of_birth,

      id_card:
        employee.id_card,

      // Master เป็นเจ้าของวันที่เริ่มงาน
      employment_date:
        employee.employment_date,

      plan:
        existingEmployee?.plan ?? employee.plan,

      department:
        employee.department,

      bank_account:
        employee.bank_account,

      bank_name:
        employee.bank_name,

      phone:
        employee.phone,

      remark:
        employee.remark,

            resignation_date:
        existingEmployee?.status === "ลาออก"
          ? existingEmployee.resignation_date
          : employee.resignation_date,

      status:
        existingEmployee?.status === "ลาออก"
          ? "ลาออก"
          : employee.status,

      // ==========================================
      // ข้อมูลจากไฟล์ประกันส่งกลับ
      // ถ้ามีอยู่แล้ว ให้เก็บค่าเดิมไว้
      // ==========================================

      effective_date:
        existingEmployee?.effective_date ??
        employee.effective_date,

      insurance_type:
        existingEmployee?.insurance_type ??
        employee.insurance_type,

      insurance_card_no:
        existingEmployee?.insurance_card_no ??
        employee.insurance_card_no,

      life_plan:
        existingEmployee?.life_plan ??
        employee.life_plan,

      updated_at:
        new Date().toISOString(),
    };
  });

console.log(
  "📦 จำนวนข้อมูลที่จะ Upsert:",
  rowsToUpsert.length
);

const chunkSize = 500;

let totalImported = 0;

// คนเดิม (มี id) → upsert / คนใหม่ (ไม่มี id) → insert แยกกัน
const rowsToUpdate = rowsToUpsert.filter((r) => "id" in r);
const rowsToInsert = rowsToUpsert.filter((r) => !("id" in r));

const saveRows = async (
  rows: typeof rowsToUpsert,
  mode: "upsert" | "insert"
) => {
  for (let i = 0; i < rows.length; i += chunkSize) {
    const chunk = rows.slice(i, i + chunkSize);

    const { error } =
      mode === "upsert"
        ? await supabase.from("employees").upsert(chunk)
        : await supabase.from("employees").insert(chunk);

    if (error) {
      console.error("❌ Supabase Error:", error);
      throw new Error(`บันทึกข้อมูลพนักงานไม่สำเร็จ: ${error.message}`);
    }

    totalImported += chunk.length;
  }
};

await saveRows(rowsToUpdate, "upsert");
await saveRows(rowsToInsert, "insert");

console.log(
  "✅ บันทึกข้อมูลสำเร็จ:",
  totalImported,
  "รายการ"
);
    
    // ==========================================
    // 6. โหลดข้อมูลใหม่จาก Database
    // ==========================================
try {

  const latestEmployees =
    await loadAllEmployees();

  setEmployees(latestEmployees);

  setRowCount(latestEmployees.length);

} catch (reloadError) {

  console.error(
    "Reload Error:",
    reloadError
  );

}

    // ==========================================
    // 7. สำเร็จ
    // ==========================================
    setSuccessMessage(
      `นำเข้าข้อมูลสำเร็จ ${totalImported.toLocaleString()} รายการ`
    );


  } catch (err) {

    console.error(
      "Import Employee Error:",
      err
    );

    setError(
      err instanceof Error
        ? err.message
        : "ไม่สามารถบันทึกข้อมูลลง Supabase ได้"
    );

  } finally {

    setIsImporting(false);

  }
};

  const outGroupKeys = Object.keys(OUT_GROUPS) as OutGroupKey[];

  const outGroupCounts = outGroupKeys.reduce((acc, key) => {
    acc[key] = outProblems.filter((p) => p.group === key).length;
    return acc;
  }, {} as Record<OutGroupKey, number>);

  const outCheckTotal = outGroupKeys
    .filter((key) => OUT_GROUPS[key].check)
    .reduce((sum, key) => sum + outGroupCounts[key], 0);

  const outModalList = outProblems.filter((p) =>
    outGroupTab === "allCheck"
      ? OUT_GROUPS[p.group].check
      : p.group === outGroupTab
  );

  const inNameList: Employee[] =
  inNameTab === "new"
    ? inNewEmployees
    : inNameTab === "existing"
    ? inExistingEmployees
    : inChanged.map((c) => ({
        ...c.emp,
        // ใช้ช่องบริษัทแสดงรหัสเดิมและสถานะเดิม (เฉพาะแสดงผล)
        vendor: `รหัสเดิม ${c.db.employee_code} (${c.db.status})`,
      }));

  return (
    <main className="min-h-screen bg-slate-100 p-8">
      <div className="mx-auto max-w-7xl">

        {/* Header */}
        <div className="mb-8">
          <h1 className="text-3xl font-bold text-slate-900">
            Group Insurance Management
          </h1>

          <p className="mt-2 text-slate-500">
            ระบบจัดการข้อมูลประกันกลุ่ม
          </p>
        </div>

        {/* =========================
            Import Initial Data
        ========================= */}
        <div className="mb-8 rounded-2xl bg-white p-6 shadow-sm">

          <h2 className="text-xl font-bold text-slate-900">
            นำเข้าข้อมูลตั้งต้น
          </h2>

          <p className="mt-2 text-sm text-slate-500">
            อัปโหลดไฟล์ Excel ข้อมูลรายชื่อผู้เอาประกัน
          </p>

          <div className="mt-5">
            <label
              htmlFor="excel-upload"
              className="inline-flex cursor-pointer items-center rounded-xl bg-blue-600 px-5 py-3 font-medium text-white transition hover:bg-blue-700"
            >
              📂 เลือกไฟล์ Excel
            </label>

            <input
  id="excel-upload"
  type="file"
  accept=".xlsx,.xls"
  onChange={(e) => {
    console.log("🟢 INPUT FILE ทำงานแล้ว");
    handleFileUpload(e);
  }}
  className="hidden"
/>
          </div>

          {/* File Status */}
          {fileName && (
            <div className="mt-5 rounded-xl bg-slate-50 p-4">

              <p className="text-sm text-slate-500">
                ไฟล์ที่เลือก
              </p>

              <p className="mt-1 font-semibold text-slate-900">
                {fileName}
              </p>

              {/* Loading */}
              {isLoading && (
                <div className="mt-3 flex items-center gap-3 text-blue-600">

                  <span className="h-5 w-5 animate-spin rounded-full border-2 border-blue-600 border-t-transparent"></span>

                  <span className="font-medium">
                    กำลังอ่านไฟล์ กรุณารอสักครู่...
                  </span>

                </div>
              )}
              {/* Save Button */}
{!isLoading && fileName && (
  <button
    type="button"
    onClick={handleImportToSupabase}
    disabled={false}
    className="mt-4 inline-flex items-center rounded-xl bg-green-600 px-5 py-3 font-medium text-white transition hover:bg-green-700 disabled:cursor-not-allowed disabled:opacity-70"
  >
    {isImporting ? (
      <>
        <span className="mr-2 h-5 w-5 animate-spin rounded-full border-2 border-white border-t-transparent" />
        กำลังบันทึกข้อมูล...
      </>
    ) : (
      <>
        💾 บันทึกข้อมูลเข้าสู่ฐานข้อมูล
      </>
    )}
  </button>
)}



              {/* Error */}
              {error && (
                <p className="mt-2 text-red-500">
                  {error}
                </p>
              )}
              {/* Success Message */}
{successMessage && (
  <p className="mt-2 font-medium text-green-600">
    ✓ {successMessage}
  </p>
)}

            </div>
          )}

        </div>

{/* =========================
    แจ้งเข้า
========================= */}
<div className="mb-8 rounded-2xl bg-white p-6 shadow-sm">

  <h2 className="text-xl font-bold text-slate-900">
    แจ้งเข้า
  </h2>

  <p className="mt-2 text-sm text-slate-500">
    อัปโหลดรายชื่อพนักงานใหม่ ระบบจะเพิ่มเฉพาะพนักงานที่ยังไม่มีในฐานข้อมูล
  </p>

  <div className="mt-5">

   <label
  htmlFor="in-upload"
  className="inline-flex cursor-pointer items-center rounded-xl bg-blue-600 px-5 py-3 font-medium text-white transition hover:bg-blue-700"
>
  📂 เลือกไฟล์แจ้งเข้า
</label>

    <input
      id="in-upload"
      type="file"
      accept=".xlsx,.xls"
      onChange={handleInFileUpload}
      className="hidden"
    />

  </div>

  {inFileName && (
    <div className="mt-5 rounded-xl bg-slate-50 p-4">

      <p className="text-sm text-slate-500">
        ไฟล์ที่เลือก
      </p>

      <p className="mt-1 font-semibold text-slate-900">
        {inFileName}
      </p>

      {isInLoading && (
        <div className="mt-3 flex items-center gap-3 text-blue-600">

          <span className="h-5 w-5 animate-spin rounded-full border-2 border-blue-600 border-t-transparent" />

          <span className="font-medium">
            กำลังอ่านไฟล์แจ้งเข้า...
          </span>

        </div>
      )}

      {!isInLoading &&
        inRowCount !== null && (
          <p className="mt-3 font-medium text-green-600">
            ✓ อ่านข้อมูลสำเร็จ{" "}
            {inRowCount.toLocaleString()} รายการ
          </p>
        )}
{!isInLoading && inRowCount !== null && (
  <div className="mt-4 grid grid-cols-1 gap-3 sm:grid-cols-3">

    {/* พนักงานใหม่ */}
    <div className="rounded-xl border border-[#7ED957]/20 bg-[#7ED957]/5 p-4">
      
      <div className="text-xs text-gray-400">
        พนักงานใหม่
      </div>

      <div className="mt-1 text-2xl font-black text-[#7ED957]">
        {inNewCount?.toLocaleString() ?? 0} คน
      </div>

      {inNewEmployees.length > 0 && (
        <button
          type="button"
          onClick={() => {
            setInNameTab("new");
            setShowInNames(true);
          }}
          className="mt-3 text-sm font-medium text-[#5fbd42] underline hover:text-[#4da832]"
        >
          👀 ดูรายชื่อ
        </button>
      )}

    </div>

        {/* รหัสเปลี่ยน / กลับเข้าทำงาน */}
    <div className="rounded-xl border border-amber-300 bg-amber-50 p-4">

      <div className="text-xs text-amber-700">
        รหัสเปลี่ยน / กลับเข้าทำงาน
      </div>

      <div className="mt-1 text-2xl font-black text-amber-600">
        {inChanged.length.toLocaleString()} คน
      </div>

      {inChanged.length > 0 && (
        <button
          type="button"
          onClick={() => {
            setInNameTab("changed");
            setShowInNames(true);
          }}
          className="mt-3 text-sm font-medium text-amber-700 underline hover:text-amber-800"
        >
          👀 ดูรายชื่อ
        </button>
      )}

    </div> 

    {/* มีอยู่แล้วในข้อมูลตั้งต้น */}
    <div className="rounded-xl border border-white/10 bg-white/5 p-4">
      
      <div className="text-xs text-gray-400">
        มีอยู่แล้วในข้อมูลตั้งต้น
      </div>

      <div className="mt-1 text-2xl font-black text-white">
        {inDuplicateCount?.toLocaleString() ?? 0} คน
      </div>

      {inExistingEmployees.length > 0 && (
        <button
          type="button"
          onClick={() => {
            setInNameTab("existing");
            setShowInNames(true);
          }}
          className="mt-3 text-sm font-medium text-white underline hover:text-slate-200"
        >
          👀 ดูรายชื่อ
        </button>
      )}

    </div>

  </div>
)}
{/* =========================
    Modal ดูรายชื่อแจ้งเข้า
========================= */}
{showInNames && (
  <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">

    <div className="w-full max-w-3xl overflow-hidden rounded-2xl bg-white shadow-2xl">

      {/* หัวหน้าต่าง */}
      <div className="flex items-center justify-between border-b border-slate-200 px-6 py-4">

        <div>
          <h3 className="text-lg font-bold text-slate-900">
            รายชื่อพนักงานแจ้งเข้า
          </h3>

          <p className="mt-1 text-sm text-slate-500">
                        {inNameTab === "new"
              ? `พนักงานใหม่ ${inNewEmployees.length.toLocaleString()} คน`
              : inNameTab === "changed"
              ? `รหัสเปลี่ยน / กลับเข้าทำงาน ${inChanged.length.toLocaleString()} คน`
              : `มีอยู่แล้วในข้อมูลตั้งต้น ${inExistingEmployees.length.toLocaleString()} คน`}
          </p>
        </div>

        <button
          type="button"
          onClick={() => setShowInNames(false)}
          className="rounded-lg px-3 py-2 text-xl text-slate-400 hover:bg-slate-100 hover:text-slate-700"
        >
          ✕
        </button>

      </div>


      {/* เลือกกลุ่มรายชื่อ */}
      <div className="flex gap-2 border-b border-slate-200 px-6 py-4">

        <button
          type="button"
          onClick={() => setInNameTab("new")}
          className={`rounded-lg px-4 py-2 text-sm font-medium ${
            inNameTab === "new"
              ? "bg-[#7ED957] text-white"
              : "bg-slate-100 text-slate-600 hover:bg-slate-200"
          }`}
        >
          พนักงานใหม่ {inNewEmployees.length} คน
        </button>

        <button
          type="button"
          onClick={() => setInNameTab("existing")}
          className={`rounded-lg px-4 py-2 text-sm font-medium ${
            inNameTab === "existing"
              ? "bg-slate-600 text-white"
              : "bg-slate-100 text-slate-600 hover:bg-slate-200"
          }`}
        >
          มีอยู่แล้ว {inExistingEmployees.length} คน
        </button>

      
        <button
          type="button"
          onClick={() => setInNameTab("changed")}
          className={`rounded-lg px-4 py-2 text-sm font-medium ${
            inNameTab === "changed"
              ? "bg-amber-500 text-white"
              : "bg-slate-100 text-slate-600 hover:bg-slate-200"
          }`}
        >
          รหัสเปลี่ยน {inChanged.length} คน
        </button>

      </div>


            {/* รายชื่อ */}
      <div className="max-h-[60vh] overflow-y-auto px-6 py-3">

        {inNameList.map((employee, index) => (

          <div
            key={`${employee.employee_code}-${index}`}
            className="flex items-center justify-between border-b border-slate-100 py-3"
          >

            <div>
              <div className="font-medium text-slate-800">
                {employee.title}
                {employee.first_name} {employee.last_name}
              </div>

              <div className="mt-1 text-xs text-slate-400">
                รหัสพนักงาน: {employee.employee_code}
                {employee.vendor
                  ? ` • ${employee.vendor}`
                  : ""}
              </div>
            </div>

            <div className="text-sm text-slate-400">
              {index + 1}
            </div>

          </div>

        ))}

        {inNameList.length === 0 && (
          <div className="py-10 text-center text-sm text-slate-400">
            ไม่พบรายชื่อ
          </div>
        )}

      </div>


      {/* ปุ่มปิด */}
      <div className="flex justify-end border-t border-slate-200 px-6 py-4">

        <button
          type="button"
          onClick={() => setShowInNames(false)}
          className="rounded-xl bg-slate-100 px-5 py-2.5 text-sm font-medium text-slate-700 hover:bg-slate-200"
        >
          ปิด
        </button>

      </div>

    </div>

  </div>
)}
{inSuccessMessage && (
  <p className="mt-4 font-medium text-green-600">
    ✓ {inSuccessMessage}
  </p>
)}

{!isInLoading &&
  inRowCount !== null && (
    <button
      type="button"
      onClick={() => {
        console.log("🟢 CLICK แจ้งเข้า");
        handleImportInToSupabase();
      }}
      disabled={isInImporting}
      className="mt-4 inline-flex items-center rounded-xl bg-green-600 px-5 py-3 font-medium text-white transition hover:bg-green-700 disabled:cursor-not-allowed disabled:opacity-70"
    >
      {isInImporting ? (
        <>
          <span className="mr-2 h-5 w-5 animate-spin rounded-full border-2 border-white border-t-transparent" />
          กำลังบันทึกข้อมูล...
        </>
      ) : (
        <>
          💾 บันทึกข้อมูลเข้าสู่ฐานข้อมูล
        </>
      )}
    </button>
  )}

    </div>
  )}

</div>


{/* =========================
    แจ้งออก
========================= */}
<div className="mb-8 rounded-2xl bg-white p-6 shadow-sm">

  <h2 className="text-xl font-bold text-slate-900">
    แจ้งออก
  </h2>

  <p className="mt-2 text-sm text-slate-500">
    อัปโหลดรายชื่อพนักงานลาออก ระบบจะเปลี่ยนสถานะเป็นลาออก โดยไม่ลบข้อมูลออกจากฐานข้อมูล
  </p>

  <div className="mt-5">

    <label
      htmlFor="out-upload"
      className="inline-flex cursor-pointer items-center rounded-xl bg-red-500 px-5 py-3 font-medium text-white transition hover:bg-red-600"
    >
      📂 เลือกไฟล์แจ้งออก
    </label>

    <input
      id="out-upload"
      type="file"
      accept=".xlsx,.xls"
      onChange={handleOutFileUpload}
      className="hidden"
    />

  </div>

  {outFileName && (

    <div className="mt-5 rounded-xl bg-slate-50 p-4">

      <p className="text-sm text-slate-500">
        ไฟล์ที่เลือก
      </p>

      <p className="mt-1 font-semibold text-slate-900">
        {outFileName}
      </p>

      {isOutLoading && (
        <div className="mt-3 flex items-center gap-3 text-red-500">

          <span className="h-5 w-5 animate-spin rounded-full border-2 border-red-500 border-t-transparent" />

          <span className="font-medium">
            กำลังอ่านไฟล์แจ้งออก...
          </span>

        </div>
      )}

      {!isOutLoading && outRowCount !== null && (
        <>
          <div className="mt-4 grid grid-cols-1 gap-3 sm:grid-cols-3">

            {/* ทั้งหมดในไฟล์ */}
            <div className="rounded-xl border border-slate-200 bg-white p-4">
              <div className="text-xs text-slate-500">
                ทั้งหมดในไฟล์
              </div>

              <div className="mt-1 text-2xl font-black text-slate-800">
                {outRowCount.toLocaleString()} คน
              </div>
            </div>

            {/* พร้อมแจ้งออก */}
            <div className="rounded-xl border border-red-200 bg-red-50 p-4">
              <div className="text-xs text-red-600">
                พร้อมแจ้งออก
              </div>

              <div className="mt-1 text-2xl font-black text-red-500">
                {outPreview.length.toLocaleString()} คน
              </div>

              <div className="mt-1 text-xs text-slate-500">
                รหัส+ชื่อตรง ยังไม่ลาออก มีวันที่ครบ
              </div>

              {outPreview.length > 0 && (
                <button
                  type="button"
                  onClick={() => setShowOutList(!showOutList)}
                  className="mt-3 text-sm font-medium text-red-600 underline hover:text-red-700"
                >
                  {showOutList ? "ซ่อนรายชื่อ" : "👀 ดูรายชื่อ"}
                </button>
              )}
            </div>

            {/* การ์ดแต่ละกลุ่ม (แสดงเฉพาะกลุ่มที่มีคน) */}
            {outGroupKeys.map((key) => {
              const count = outGroupCounts[key];

              if (count === 0) return null;

              const g = OUT_GROUPS[key];

              return (
                <div
                  key={key}
                  className={`rounded-xl border p-4 ${g.card}`}
                >
                  <div className={`text-xs ${g.text}`}>
                    {g.check ? "⚠️ " : ""}
                    {g.label}
                  </div>

                  <div className={`mt-1 text-2xl font-black ${g.text}`}>
                    {count.toLocaleString()} คน
                  </div>

                  <div className="mt-1 text-xs text-slate-500">
                    {g.hint}
                  </div>

                  <button
                    type="button"
                    onClick={() => {
                      setOutGroupTab(key);
                      setShowOutNames(true);
                    }}
                    className={`mt-3 text-sm font-medium underline ${g.text}`}
                  >
                    👀 ดูรายชื่อ
                  </button>
                </div>
              );
            })}

          </div>

          {/* แถบเตือนรวม */}
          {outCheckTotal > 0 && (
            <div className="mt-4 flex items-center justify-between gap-4 rounded-xl border border-amber-300 bg-amber-50 px-4 py-3">
              <p className="text-sm font-medium text-amber-700">
                ⚠️ มี {outCheckTotal.toLocaleString()} คนที่ต้องตรวจสอบ
                (ยังไม่ถูกบันทึก ให้ไปเช็คระบบหลังบ้าน)
              </p>

              <button
                type="button"
                onClick={() => {
                  setOutGroupTab("allCheck");
                  setShowOutNames(true);
                }}
                className="shrink-0 rounded-lg border border-amber-300 bg-white px-4 py-2 text-sm font-medium text-amber-700 hover:bg-amber-100"
              >
                ดูทั้งหมดที่ต้องตรวจ
              </button>
            </div>
          )}

          {/* รายชื่อพร้อมแจ้งออก */}
          {showOutList && outPreview.length > 0 && (
            <div className="mt-4 max-h-80 overflow-y-auto rounded-xl border border-slate-200 bg-white">
              <div className="divide-y divide-slate-100">
                {outPreview.map((employee, index) => (
                  <div
                    key={`${employee.employee_code}-${index}`}
                    className="flex items-center justify-between gap-4 px-4 py-3"
                  >
                    <div>
                      <p className="font-medium text-slate-900">
                        {employee.title}
                        {employee.first_name} {employee.last_name}
                      </p>

                      <p className="mt-1 text-xs text-slate-500">
                        รหัสพนักงาน: {employee.employee_code}
                        {employee.resignation_date
                          ? ` • วันที่ออก ${employee.resignation_date}`
                          : ""}
                      </p>
                    </div>

                    <span className="shrink-0 rounded-full bg-red-50 px-3 py-1 text-xs font-semibold text-red-500">
                      ลาออก
                    </span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Modal รายชื่อกลุ่มที่ต้องตรวจ */}
          {showOutNames && (
            <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
              <div className="w-full max-w-3xl overflow-hidden rounded-2xl bg-white shadow-2xl">

                <div className="flex items-center justify-between border-b border-slate-200 px-6 py-4">
                  <div>
                    <h3 className="text-lg font-bold text-slate-900">
                      รายชื่อแจ้งออกที่ต้องตรวจสอบ
                    </h3>

                    <p className="mt-1 text-sm text-slate-500">
                      {outModalList.length.toLocaleString()} คน
                      (ยังไม่ถูกบันทึก ให้ตรวจในระบบหลังบ้าน)
                    </p>
                  </div>

                  <button
                    type="button"
                    onClick={() => setShowOutNames(false)}
                    className="rounded-lg px-3 py-2 text-xl text-slate-400 hover:bg-slate-100 hover:text-slate-700"
                  >
                    ✕
                  </button>
                </div>

                {/* แท็บเลือกกลุ่ม */}
                <div className="flex flex-wrap gap-2 border-b border-slate-200 px-6 py-4">
                  <button
                    type="button"
                    onClick={() => setOutGroupTab("allCheck")}
                    className={`rounded-lg px-4 py-2 text-sm font-medium ${
                      outGroupTab === "allCheck"
                        ? "bg-amber-500 text-white"
                        : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                    }`}
                  >
                    ต้องตรวจทั้งหมด {outCheckTotal}
                  </button>

                  {outGroupKeys.map((key) => {
                    if (outGroupCounts[key] === 0) return null;

                    return (
                      <button
                        key={key}
                        type="button"
                        onClick={() => setOutGroupTab(key)}
                        className={`rounded-lg px-4 py-2 text-sm font-medium ${
                          outGroupTab === key
                            ? "bg-slate-700 text-white"
                            : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                        }`}
                      >
                        {OUT_GROUPS[key].label} {outGroupCounts[key]}
                      </button>
                    );
                  })}
                </div>

                <div className="max-h-[60vh] overflow-y-auto px-6 py-3">
                  {outModalList.map((p, index) => (
                    <div
                      key={`${p.row.employee_code}-${index}`}
                      className="flex items-center justify-between gap-4 border-b border-slate-100 py-3"
                    >
                      <div>
                        <div className="font-medium text-slate-800">
                          {p.row.title}
                          {p.row.first_name} {p.row.last_name}
                        </div>

                        <div className="mt-1 text-xs text-slate-400">
                          รหัสพนักงาน: {p.row.employee_code}
                          {p.row.resignation_date
                            ? ` • วันที่ออก ${p.row.resignation_date}`
                            : ""}
                        </div>

                        <div
                          className={`mt-1 text-xs font-medium ${OUT_GROUPS[p.group].text}`}
                        >
                          {OUT_GROUPS[p.group].label}: {p.reason}
                        </div>
                      </div>

                      <div className="text-sm text-slate-400">
                        {index + 1}
                      </div>
                    </div>
                  ))}

                  {outModalList.length === 0 && (
                    <div className="py-10 text-center text-sm text-slate-400">
                      ไม่พบรายชื่อ
                    </div>
                  )}
                </div>

                <div className="flex justify-end border-t border-slate-200 px-6 py-4">
                  <button
                    type="button"
                    onClick={() => setShowOutNames(false)}
                    className="rounded-xl bg-slate-100 px-5 py-2.5 text-sm font-medium text-slate-700 hover:bg-slate-200"
                  >
                    ปิด
                  </button>
                </div>

              </div>
            </div>
          )}

          {outSuccessMessage && (
            <p className="mt-4 font-medium text-green-600">
              ✓ {outSuccessMessage}
            </p>
          )}

          <button
            type="button"
            onClick={handleImportOutToSupabase}
            disabled={isOutImporting || outPreview.length === 0}
            className="mt-4 inline-flex items-center rounded-xl bg-red-500 px-5 py-3 font-medium text-white transition hover:bg-red-600 disabled:cursor-not-allowed disabled:opacity-50"
          >
            {isOutImporting ? (
              <>
                <span className="mr-2 h-5 w-5 animate-spin rounded-full border-2 border-white border-t-transparent" />
                กำลังแจ้งออก...
              </>
            ) : (
              <>
                💾 บันทึกแจ้งออก ({outPreview.length.toLocaleString()} คน)
              </>
            )}
          </button>
        </>
      )}

    </div>

  )}

</div>

{/* =========================
    ประกันส่งกลับ
========================= */}

<div className="mb-8 rounded-2xl bg-white p-6 shadow-sm">

  <h2 className="text-xl font-bold text-slate-900">
    ประกันส่งกลับ
  </h2>

  <p className="mt-2 text-sm text-slate-500">
    อัปโหลดไฟล์ข้อมูลที่ได้รับกลับจากบริษัทประกัน
  </p>

  <div className="mt-5">

    <label
      htmlFor="insurance-upload"
      className="inline-flex cursor-pointer items-center rounded-xl bg-orange-500 px-5 py-3 font-medium text-white transition hover:bg-orange-600"
    >
      📂 เลือกไฟล์ประกันส่งกลับ
    </label>

    <input
      id="insurance-upload"
      type="file"
      accept=".xlsx,.xls"
      onChange={handleInsuranceFileUpload}
      className="hidden"
    />

  </div>

  {insuranceFileName && (

    <div className="mt-5 rounded-xl bg-slate-50 p-4">

      <p className="text-sm text-slate-500">
        ไฟล์ที่เลือก
      </p>

      <p className="mt-1 font-semibold text-slate-900">
        {insuranceFileName}
      </p>

      {isInsuranceLoading && (

        <div className="mt-3 flex items-center gap-3 text-orange-500">

          <span className="h-5 w-5 animate-spin rounded-full border-2 border-orange-500 border-t-transparent" />

          <span className="font-medium">
            กำลังอ่านไฟล์ประกันส่งกลับ...
          </span>

        </div>

      )}

      {!isInsuranceLoading &&
        insuranceRowCount !== null && (

          <p className="mt-3 font-medium text-green-600">
            ✓ อ่านข้อมูลสำเร็จ{" "}
            {insuranceRowCount.toLocaleString()} รายการ
          </p>

        )}
        {/* ปุ่มบันทึกข้อมูลประกันส่งกลับ */}
{insuranceMatched.length > 0 && (
  <button
    type="button"
    onClick={handleImportInsuranceToSupabase}
    disabled={isInsuranceImporting}
    className="
      mt-4 w-full
      rounded-xl
      bg-[#7ED957]
      px-5 py-3
      text-sm font-black
      text-black
      transition
      hover:bg-[#8FEA68]
      disabled:cursor-not-allowed
      disabled:opacity-50
    "
  >
    {isInsuranceImporting
      ? "กำลังบันทึกข้อมูล..."
      : `💾 บันทึกข้อมูลประกันส่งกลับ (${insuranceMatched.length} รายการ)`}
  </button>
)}

    </div>

  )}

</div> 
{/* =========================
    Import Claims
========================= */}
<div className="mb-8 rounded-2xl bg-white p-6 shadow-sm">

  <h2 className="text-xl font-bold text-slate-900">
    นำเข้ารายงานการเคลม
  </h2>

  <p className="mt-2 text-sm text-slate-500">
    อัปโหลดไฟล์ Excel รายงานการเคลมจากบริษัทประกัน
  </p>

  <div className="mt-5">

    <label
      htmlFor="claim-upload"
      className="inline-flex cursor-pointer items-center rounded-xl bg-purple-600 px-5 py-3 font-medium text-white transition hover:bg-purple-700"
    >
      📄 เลือกไฟล์รายงานเคลม
    </label>

    <input
      id="claim-upload"
      type="file"
      accept=".xlsx,.xls"
      onChange={handleClaimFileUpload}
      className="hidden"
    />

  </div>


  {/* =========================
      Claim File Status
  ========================= */}

  {claimFileName && (

    <div className="mt-5 rounded-xl bg-slate-50 p-4">

      <p className="text-sm text-slate-500">
        ไฟล์ที่เลือก
      </p>

      <p className="mt-1 font-semibold text-slate-900">
        {claimFileName}
      </p>


      {/* Loading */}

      {isClaimLoading && (

        <div className="mt-3 flex items-center gap-3 text-purple-600">

          <span className="h-5 w-5 animate-spin rounded-full border-2 border-purple-600 border-t-transparent"></span>

          <span className="font-medium">
            กำลังอ่านไฟล์รายงานเคลม...
          </span>

        </div>

      )}


      {/* Success */}

      {!isClaimLoading &&
        claimRowCount !== null && (

        <p className="mt-3 font-medium text-green-600">

          ✓ อ่านข้อมูลรายงานเคลมสำเร็จ{" "}

          {claimRowCount.toLocaleString()} รายการ

        </p>

      )}


      {/* Import Button */}

      {!isClaimLoading &&
        claimPreview.length > 0 && (

        <button
          type="button"
          onClick={handleImportClaimsToSupabase}
          disabled={isClaimImporting}
          className="mt-4 inline-flex items-center rounded-xl bg-purple-600 px-5 py-3 font-medium text-white transition hover:bg-purple-700 disabled:cursor-not-allowed disabled:opacity-50"
        >

          {isClaimImporting ? (

            <>
              <span className="mr-2 h-5 w-5 animate-spin rounded-full border-2 border-white border-t-transparent"></span>

              กำลังบันทึกรายงานเคลม...

            </>

          ) : (

            <>
              💾 บันทึกรายงานเคลมเข้าระบบ
            </>

          )}

        </button>

      )}

    </div>

  )}

</div>


        {/* =========================
            Dashboard
        ========================= */}
                <h2 className="mb-3 text-lg font-bold text-slate-900">
          สรุปพนักงาน
        </h2>

        <div className="grid gap-5 md:grid-cols-3">

          <div className="rounded-2xl bg-white p-6 shadow-sm">
            <p className="text-sm text-slate-500">
              พนักงานทั้งหมด
            </p>

            <h2 className="mt-2 text-3xl font-bold text-slate-900">
              {employees.length}
            </h2>
          </div>

          <div className="rounded-2xl bg-white p-6 shadow-sm">
            <p className="text-sm text-slate-500">
              มีผลประกัน
            </p>

            <h2 className="mt-2 text-3xl font-bold text-green-600">
              {employees.filter(
                (employee) => employee.status !== "ลาออก"
              ).length}
            </h2>
          </div>

            <div className="rounded-2xl bg-white p-6 shadow-sm">
            <p className="text-sm text-slate-500">
              ลาออก
            </p>

            <h2 className="mt-2 text-3xl font-bold text-red-500">
              {employees.filter(
                (employee) => employee.status === "ลาออก"
              ).length}
            </h2>
          </div>

                  </div>

        <h2 className="mb-3 mt-8 text-lg font-bold text-slate-900">
          สรุปรายงานเคลม
        </h2>

        <div className="grid gap-5 md:grid-cols-2 lg:grid-cols-4">

          <div className="rounded-2xl bg-white p-6 shadow-sm">
  <p className="text-sm text-slate-500">
    Claims ทั้งหมด
  </p>

  <h2 className="mt-2 text-3xl font-bold text-purple-600">
    {matchStats.total.toLocaleString()}
  </h2>
</div>

<div className="rounded-2xl bg-white p-6 shadow-sm">
  <p className="text-sm text-slate-500">
    Claims พนักงาน (E)
  </p>

  <h2 className="mt-2 text-3xl font-bold text-slate-900">
    {matchStats.employeeClaims.toLocaleString()}
  </h2>
</div>

<div className="rounded-2xl bg-white p-6 shadow-sm">
  <p className="text-sm text-slate-500">
    Match ได้
  </p>

  <h2 className="mt-2 text-3xl font-bold text-green-600">
    {matchStats.matched.toLocaleString()}
  </h2>
</div>

<div className="rounded-2xl bg-white p-6 shadow-sm">
  <p className="text-sm text-slate-500">
    Match ไม่ได้
  </p>

  <h2 className="mt-2 text-3xl font-bold text-red-600">
    {matchStats.unmatched.toLocaleString()}
  </h2>
</div>

        </div>

{/* =========================
    Employee Table
========================= */}
<div className="mt-8 overflow-hidden rounded-2xl bg-white shadow-sm">

  <div className="border-b p-6">

    <h2 className="text-xl font-bold text-slate-900">
      รายชื่อผู้เอาประกัน
    </h2>

<div className="mt-4 flex flex-col gap-3 sm:flex-row">
  <input
    type="text"
    value={employeeSearchInput}
    onChange={(e) =>
      setEmployeeSearchInput(e.target.value)
    }
    onKeyDown={(e) => {
      if (e.key === "Enter") {
        setEmployeeSearchTerm(employeeSearchInput);
      }
    }}
    placeholder="ค้นหารหัสพนักงาน หรือ ชื่อ-นามสกุล..."
    className="w-full rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-blue-400 focus:ring-2 focus:ring-blue-100 sm:max-w-md"
  />

  <button
    type="button"
    onClick={() => {
      setEmployeeSearchTerm(
        employeeSearchInput
      );
    }}
    className="inline-flex items-center justify-center rounded-xl bg-blue-600 px-5 py-3 text-sm font-semibold text-white transition hover:bg-blue-700"
  >
    🔍 ค้นหาพนักงาน
  </button>
    <select
    value={statusFilter}
    onChange={(e) => setStatusFilter(e.target.value)}
    className="rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none transition focus:border-blue-400 focus:ring-2 focus:ring-blue-100"
  >
    <option value="ทั้งหมด">ทั้งหมด</option>
    <option value="มีผลประกัน">ทำงาน</option>
    <option value="ลาออก">ลาออก</option>
  </select>

  {employeeSearchTerm && (
    <button
      type="button"
      onClick={() => {
        setEmployeeSearchInput("");
        setEmployeeSearchTerm("");
      }}
      className="inline-flex items-center justify-center rounded-xl border border-slate-200 bg-white px-5 py-3 text-sm font-semibold text-slate-600 transition hover:bg-slate-50"
    >
      ล้างการค้นหา
    </button>
  )}
</div>
    {employees.length > 0 && (
  <p className="mt-1 text-sm text-slate-500">
    {employeeSearchTerm
      ? `พบ ${filteredEmployees.length.toLocaleString()} รายการ จากทั้งหมด ${employees.length.toLocaleString()} รายการ`
      : `แสดงทั้งหมด ${employees.length.toLocaleString()} รายการ`}
    {" • "}
    มีข้อมูลเคลม{" "}
    {claims.length.toLocaleString()} รายการ
  </p>
)}

  </div>

  <div className="overflow-x-auto">

    <table className="w-full text-left">

    <thead className="bg-slate-50 text-sm text-slate-600">
  <tr className="border-b border-slate-200">

    <th className="w-[120px] px-4 py-3 text-center font-semibold whitespace-nowrap">
      Vendor
    </th>

    <th className="w-[70px] px-4 py-3 text-center font-semibold whitespace-nowrap">
      ลำดับ
    </th>

    <th className="w-[180px] px-4 py-3 text-left font-semibold whitespace-nowrap">
      ชื่อ
    </th>

    <th className="w-[70px] px-4 py-3 text-center font-semibold whitespace-nowrap">
      แผน
    </th>

    <th className="w-[90px] px-4 py-3 text-center font-semibold">
      ประเภทปี
    </th>

    <th className="w-[110px] px-4 py-3 text-center font-semibold">
      มีผลประกัน
    </th>

    <th className="w-[100px] px-4 py-3 text-center font-semibold">
      ลาออก
    </th>

    <th className="w-[110px] px-4 py-3 text-center font-semibold">
      สถานะ
    </th>

    <th className="w-[120px] px-4 py-3 text-center font-semibold">
      <span className="block">อยู่มาแล้ว</span>
      <span className="block">กี่วัน</span>
    </th>

    <th className="w-[110px] px-4 py-3 text-center font-semibold">
      <span className="block">ค่าเบี้ย</span>
      <span className="block">ต่อคน</span>
    </th>

    <th className="w-[110px] px-4 py-3 text-center font-semibold">
      <span className="block">ค่าเบี้ย</span>
      <span className="block">ต่อวัน</span>
    </th>

    <th className="w-[80px] px-4 py-3 text-center font-semibold whitespace-nowrap">
      OPD
    </th>

    <th className="w-[80px] px-4 py-3 text-center font-semibold whitespace-nowrap">
      IPD
    </th>

    <th className="w-[130px] px-4 py-3 text-center font-semibold">
      <span className="block">หมายเลข</span>
      <span className="block">บัตร</span>
    </th>

    <th className="w-[70px] px-4 py-3 text-center font-semibold whitespace-nowrap">
      ครั้ง
    </th>

  </tr>
</thead> 

      <tbody>

        {employees.length === 0 ? (

          <tr className="border-t">

            <td
              colSpan={14}
              className="px-6 py-10 text-center text-slate-400"
            >
              {isLoading
                ? "กำลังเตรียมข้อมูล..."
                : "ยังไม่มีข้อมูล กรุณานำเข้าข้อมูลพนักงาน"}
            </td>

          </tr>

        ) : (

          filteredEmployees.map((employee, index) => {

            // =========================
            // จำนวนวันที่มีประกัน
            // =========================
            const days = calculateDays(
              employee.effective_date,
              employee.resignation_date
            );

            // =========================
            // ค่าเบี้ยตามแผน
            // =========================
            const premium =
              employee.plan !== null
                ? planPremium[employee.plan] ?? 0
                : 0;

            // =========================
            // ค่าเบี้ยต่อวัน
            // =========================
            const dailyPremium =
              premium > 0
                ? premium / 365
                : 0;
            // =========================
            // ข้อมูลการเคลม
            // =========================
const claimSummary =
  claimSummaryMap.get(employee.id ?? "") ?? {
    opd: 0,
    ipd: 0,
    count: 0,
  };
            return (

              <tr
                key={employee.id ?? `${employee.employee_key}-${index}`}
                className="border-t text-sm hover:bg-slate-50"
              >

                {/* บริษัท */}
                <td className="px-6 py-4">
                  {employee.vendor || "-"}
                </td>

                {/* ลำดับ */}
                <td className="px-6 py-4">
                  {index + 1}
                </td>

                {/* ชื่อ */}
                <td className="px-6 py-4">
                  {employee.title}{" "}
                  {employee.first_name}{" "}
                  {employee.last_name}
                </td>

                {/* แผน */}
                <td className="px-6 py-4">
                  {employee.plan ?? "-"}
                </td>

                <td>
                  {employee.insurance_type || "-"}
                </td>

                {/* วันที่มีผล */}
                <td className="px-6 py-4">
                  {employee.effective_date || "-"}
                </td>

                {/* วันที่ลาออก */}
                <td className="px-6 py-4">
                  {employee.resignation_date || "-"}
                </td>

                {/* สถานะ */}
                <td className="px-6 py-4">

                  {employee.status === "ลาออก" ? (
                    <span className="font-medium text-red-500">
                      ลาออก
                    </span>
                  ) : (
                    <span className="font-medium text-green-600">
                      ทำงาน
                    </span>
                  )}

                </td>

                {/* จำนวนวัน */}
                <td className="px-6 py-4">
                  {days.toLocaleString()} วัน
                </td>

                {/* ค่าเบี้ยต่อคน */}
                <td className="px-6 py-4">
                  {premium > 0
                    ? `${premium.toLocaleString()} บาท`
                    : "-"}
                </td>

                {/* ค่าเบี้ยต่อวัน */}
                <td className="px-6 py-4">
                  {dailyPremium > 0
                    ? `${dailyPremium.toFixed(2)} บาท`
                    : "-"}
                </td>

                {/* OPD */}
                <td className="px-6 py-4">
                  {claimSummary.opd.toLocaleString()} บาท
                </td>

                {/* IPD */}
                <td className="px-6 py-4">
                  {claimSummary.ipd.toLocaleString()} บาท
                </td>

                {/* หมายเลขบัตร */}
                <td className="px-6 py-4">
                  {employee.insurance_card_no || "-"}
                </td>

                {/* จำนวนครั้ง */}
                <td className="px-6 py-4">
                  {claimSummary.count}
                </td>
              </tr>

            );

          })

        )}

      </tbody>
    </table>
  </div>
</div>
</div>
    </main>
  );
}