"use client";

import {
  useEffect,
  useMemo,
  useState,
} from "react";
import * as XLSX from "xlsx";
import { supabase } from "@/lib/supabase";

type Employee = {
  id?: string;
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
};

// =========================
// แปลงวันที่จาก Excel
// =========================
function formatExcelDate(value: any): string | null {
  if (!value) return null;

  if (value instanceof Date) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");

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
  idCard: string,
  effectiveDate: string | null
) => {
  const cleanVendor = normalizeText(vendor);
  const cleanCode = normalizeText(employeeCode);
  const cleanFirstName = normalizeText(firstName);
  const cleanLastName = normalizeText(lastName);
  const cleanIdCard = normalizeText(idCard);
  const cleanEffectiveDate = normalizeText(effectiveDate);

  // ถ้ามีรหัสพนักงาน + วันที่มีผล
  if (cleanCode) {
    return `${cleanVendor}|code|${cleanCode}|effective|${cleanEffectiveDate}`;
  }

  // ถ้าไม่มีรหัสพนักงาน แต่มีบัตรประชาชน
  if (cleanIdCard) {
    return `${cleanVendor}|id|${cleanIdCard}|effective|${cleanEffectiveDate}`;
  }

  // ถ้าไม่มีทั้งรหัสและบัตรประชาชน
  return `${cleanVendor}|name|${cleanFirstName}|${cleanLastName}|effective|${cleanEffectiveDate}`;
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
// สร้าง Map สรุป Claims
// =========================
const getClaimSummaryMap = (
  claims: Claim[]
) => {

  const claimMap = new Map<
    string,
    {
      opd: number;
      ipd: number;
      count: number;
    }
  >();

  claims.forEach((claim) => {
    if (
  String(claim.relationship ?? "")
    .trim()
    .toUpperCase() !== "E"
) {
  return;
}

const idCard = String(
  claim.id_card ?? ""
)
  .trim()
  .replace(/\D/g, "");

if (!idCard) return;

const current =
  claimMap.get(idCard) ?? {
        opd: 0,
        ipd: 0,
        count: 0,
      };

    const claimType = String(
      claim.claim_type ?? ""
    )
      .trim()
      .toUpperCase();

    const paid = Number(
      claim.insurance_paid ?? 0
    );

    // OPD
    if (
      claimType.includes("OPD") ||
      claimType.includes("ผู้ป่วยนอก")
    ) {
      current.opd += paid;
    }

    // IPD
    if (
      claimType.includes("IPD") ||
      claimType.includes("ผู้ป่วยใน")
    ) {
      current.ipd += paid;
    }

    // จำนวนรายการเคลม
    current.count += 1;

    claimMap.set(
  idCard,
  current
);

  });

  console.log(
    "Claim Summary Map:",
    Array.from(claimMap.entries()).slice(0, 20)
  );

  return claimMap;
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
  const [claims, setClaims] = useState<Claim[]>([]);
  const [employeeSearchInput, setEmployeeSearchInput] = useState("");
const [employeeSearchTerm, setEmployeeSearchTerm] = useState("");
  // =========================
// สรุป Claims แบบรวดเร็ว
// =========================
const filteredEmployees = useMemo(() => {
  const search = employeeSearchTerm
    .trim()
    .toLowerCase();

  if (!search) {
    return employees;
  }

  return employees.filter((employee) => {
    const employeeCode = String(
      employee.employee_code ?? ""
    )
      .trim()
      .toLowerCase();

    const fullName = `${employee.title ?? ""} ${employee.first_name ?? ""} ${employee.last_name ?? ""}`
      .trim()
      .toLowerCase();

    return (
      employeeCode.includes(search) ||
      fullName.includes(search)
    );
  });
}, [employees, employeeSearchTerm]);
const claimSummaryMap = useMemo(() => {

  return getClaimSummaryMap(
    claims
  );

}, [claims]);
const matchStats = useMemo(() => {
  if (claims.length === 0) {
    return {
      total: 0,
      matched: 0,
      unmatched: 0,
      employeeClaims: 0,
      dependentClaims: 0,
    };
  }

  const employeeIdCardSet = new Set(
    employees
      .map((employee) =>
        String(employee.id_card ?? "")
          .trim()
          .replace(/\D/g, "")
      )
      .filter(Boolean)
  );

  let matched = 0;
  let unmatched = 0;
  let employeeClaims = 0;
  let dependentClaims = 0;

  claims.forEach((claim) => {
    const relationship = String(
      claim.relationship ?? ""
    )
      .trim()
      .toUpperCase();

    const claimIdCard = String(
      claim.id_card ?? ""
    )
      .trim()
      .replace(/\D/g, "");

    // พนักงาน
    if (relationship === "E") {
      employeeClaims++;

      if (
        claimIdCard &&
        employeeIdCardSet.has(claimIdCard)
      ) {
        matched++;
      } else {
        unmatched++;
      }
    } else {
      // คู่สมรส / บุตร / บุคคลอื่น
      dependentClaims++;
    }
  });

  return {
    total: claims.length,
    matched,
    unmatched,
    employeeClaims,
    dependentClaims,
  };
}, [claims, employees]);
const matchedEmployeeCount = useMemo(() => {

  return employees.filter((employee) => {

    const employeeIdCard = String(
      employee.id_card ?? ""
    )
      .trim()
      .replace(/\D/g, "");

    if (!employeeIdCard) {
      return false;
    }

    return claimSummaryMap.has(employeeIdCard);

  }).length;

}, [employees, claimSummaryMap]);
useEffect(() => {
  if (employees.length === 0 || claims.length === 0) return;

const employeeIdCardSet = new Set(
  employees
    .map((employee) =>
      String(employee.id_card ?? "")
        .trim()
        .replace(/\D/g, "")
    )
    .filter(Boolean)
);

const unmatchedClaims = claims.filter((claim) => {

  const claimIdCard = String(
    claim.id_card ?? ""
  )
    .trim()
    .replace(/\D/g, "");

  return (
    claimIdCard === "" ||
    !employeeIdCardSet.has(claimIdCard)
  );
});

  console.log("========== ตรวจสอบ Claims ==========");
  console.log("Claims ทั้งหมด:", claims.length);
  console.log("Match ไม่ได้:", unmatchedClaims.length);

console.log(
  "Claims ที่ Match ไม่ได้ทั้งหมด:",
  unmatchedClaims.map((claim, index) => ({
    ลำดับ: index + 1,
    ชื่อในไฟล์เคลม: claim.employee_name,
    บัตรประชาชนในไฟล์เคลม: claim.id_card,
    ความสัมพันธ์: claim.relationship,
    ประเภทเคลม: claim.claim_type,
    สังกัด: claim.affiliation,
    เลขที่เคลม: claim.claim_no,
  }))
);

  console.log("====================================");
}, [employees, claims]);

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
const [isInLoading, setIsInLoading] = useState(false);
const [isInImporting, setIsInImporting] = useState(false);
const [inNewCount, setInNewCount] = useState<number | null>(null);
const [inDuplicateCount, setInDuplicateCount] = useState<number | null>(null);
const [inSuccessMessage, setInSuccessMessage] = useState("");

// แจ้งออก
const [outFileName, setOutFileName] = useState("");
const [outRowCount, setOutRowCount] = useState<number | null>(null);
const [outPreview, setOutPreview] = useState<any[]>([]);
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

    if (!file) return;

    setFileName(file.name);
    setError("");
    setSuccessMessage("");
    setRowCount(null);
    setEmployees([]);
    setIsLoading(true);

    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;

        if (!data) {
          throw new Error("ไม่พบข้อมูลไฟล์");
        }

        const workbook = XLSX.read(data, {
          type: "array",
          cellDates: true,
        });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // อ่าน Excel เป็น Array
        // เพราะไฟล์จริงของเรามี Header 2 แถว
        const rows = XLSX.utils.sheet_to_json<any[]>(worksheet, {
          header: 1,
          defval: "",
        });

        console.log("ข้อมูลทั้งหมดจาก Excel:", rows);

        // ---------------------------------
        // ตรวจสอบว่ามีข้อมูลหรือไม่
        // ---------------------------------
        if (rows.length < 5) {
          throw new Error(
            "ไฟล์ Excel ไม่มีข้อมูลพนักงาน หรือรูปแบบไฟล์ไม่ถูกต้อง"
          );
        }

        // ---------------------------------
        // ไฟล์ตั้งต้นมี 4 แถวก่อนข้อมูลจริง
        //
        // แถวที่ 1 = Header ภาษาไทย
        // แถวที่ 2 = Header ภาษาอังกฤษ
        // แถวที่ 3 = ชื่อบริษัท
        // แถวที่ 4 = Company
        // แถวที่ 5 เป็นต้นไป = ข้อมูลพนักงานจริง
        // ---------------------------------
        const dataRows = rows.slice(4);

        // ---------------------------------
        // แปลงข้อมูล Excel → Employee
        // ---------------------------------
const mappedEmployees: Employee[] = dataRows
  .filter((row) => {
    const firstName = String(row[5] ?? "").trim();
    const lastName = String(row[6] ?? "").trim();

    return (
      firstName !== "" &&
      lastName !== ""
    );
  })
  .map((row) => {

    const vendor = String(row[1] ?? "").trim();

    const employeeCode = String(row[3] ?? "")
      .trim()
      .replace(/\s+/g, "")
      .toUpperCase();

    const firstName = String(row[5] ?? "").trim();

    const lastName = String(row[6] ?? "").trim();

    const idCard = String(row[9] ?? "").trim();

    return {
      employee_key: createEmployeeKey(
  vendor,
  employeeCode,
  firstName,
  lastName,
  idCard,
  formatExcelDate(row[11])
),

      employee_code: employeeCode,

      vendor,

      branch: String(row[2] ?? "").trim(),

      title: String(row[4] ?? "").trim(),

      first_name: firstName,

      last_name: lastName,

      gender: String(row[7] ?? "").trim(),

      date_of_birth: formatExcelDate(row[8]),

      id_card: idCard,

      employment_date: formatExcelDate(row[10]),

      effective_date: formatExcelDate(row[11]),

      plan:
        row[12] !== ""
          ? Number(row[12])
          : null,

      insurance_type: "เต็มปี",

      department: String(row[13] ?? "").trim(),

      bank_account: String(row[14] ?? "").trim(),

      bank_name: String(row[15] ?? "").trim(),

      phone: String(row[16] ?? "").trim(),

      remark: String(row[17] ?? "").trim(),

      resignation_date:
        formatExcelDate(row[19]),

      status:
        String(row[18] ?? "").trim() !== ""
          ? "ลาออก"
          : "มีผลประกัน",
          insurance_card_no: "",

life_plan: "",
    };
  });
        console.log("ข้อมูลหลัง Mapping:", mappedEmployees);

        setEmployees(mappedEmployees);
        setRowCount(mappedEmployees.length);

      } catch (err) {
        console.error(err);

        setError(
          err instanceof Error
            ? err.message
            : "ไม่สามารถอ่านไฟล์ Excel ได้"
        );
      } finally {
        setIsLoading(false);
      }
    };

    reader.onerror = () => {
      setError("เกิดข้อผิดพลาดในการอ่านไฟล์");
      setIsLoading(false);
    };

    reader.readAsArrayBuffer(file);
  };
  // =========================
// อ่านไฟล์ Excel "แจ้งเข้า"
// =========================
const handleInFileUpload = (
  event: React.ChangeEvent<HTMLInputElement>
) => {
  console.log("🟦🟦🟦 HANDLE FILE UPLOAD ถูกเรียก");

  const file = event.target.files?.[0];

  console.log("📁 ไฟล์ที่เลือก:", file?.name);

  if (!file) return;

  setInFileName(file.name);
  setInRowCount(null);
  setInPreview([]);
  setInNewCount(null);
  setInDuplicateCount(null);
  setError("");
  setSuccessMessage("");
  setIsInLoading(true);
  setInSuccessMessage("");
  

  const reader = new FileReader();

  reader.onload = (e) => {
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
        workbook.SheetNames[0];

      const worksheet =
        workbook.Sheets[firstSheetName];

      const rows =
        XLSX.utils.sheet_to_json<any[]>(
          worksheet,
          {
            header: 1,
            defval: "",
          }
        );

      console.log(
        "========== แจ้งเข้า =========="
      );

      console.log(
        "ข้อมูลทั้งหมด:",
        rows
      );

      if (rows.length < 3) {
        throw new Error(
          "ไฟล์แจ้งเข้าไม่มีข้อมูล หรือรูปแบบไฟล์ไม่ถูกต้อง"
        );
      }

      // ==========================================
      // ไฟล์แจ้งเข้ามี Header 2 แถว
      //
      // แถวที่ 1 = ภาษาไทย
      // แถวที่ 2 = ภาษาอังกฤษ
      // แถวที่ 3 เป็นต้นไป = ข้อมูล
      // ==========================================
      const dataRows =
        rows.slice(2);

      // ==========================================
      // Mapping Excel → Employee
      // ==========================================
      const mappedEmployees: Employee[] =
        dataRows
          .filter((row) => {

            const employeeCode =
              String(row[2] ?? "")
                .trim();

            const firstName =
              String(row[4] ?? "")
                .trim();

            const lastName =
              String(row[5] ?? "")
                .trim();

            return (
              employeeCode !== "" ||
              (
                firstName !== "" &&
                lastName !== ""
              )
            );
          })
          .map((row) => {

            const vendor =
              String(row[1] ?? "")
                .trim();

            const employeeCode =
              String(row[2] ?? "")
                .trim()
                .replace(/\s+/g, "")
                .toUpperCase();

            const firstName =
              String(row[4] ?? "")
                .trim();

            const lastName =
              String(row[5] ?? "")
                .trim();

            const idCard =
              String(row[8] ?? "")
                .trim()
                .replace(/\D/g, "");

            const effectiveDate =
              formatExcelDate(row[10]);

            return {

              employee_key:
                createEmployeeKey(
                  vendor,
                  employeeCode,
                  firstName,
                  lastName,
                  idCard,
                  effectiveDate
                ),

              employee_code:
                employeeCode,

              vendor:
                vendor,

              branch:
                vendor,

              title:
                String(row[3] ?? "")
                  .trim(),

              first_name:
                firstName,

              last_name:
                lastName,

              gender:
                String(row[6] ?? "")
                  .trim(),

              date_of_birth:
                formatExcelDate(row[7]),

              id_card:
                idCard,

              employment_date:
                formatExcelDate(row[9]),

              effective_date:
                effectiveDate,

              plan:
                row[11] !== ""
                  ? Number(row[11])
                  : null,

              insurance_type:
                "เต็มปี",

              department:
                String(row[12] ?? "")
                  .trim(),

              bank_account:
                String(row[13] ?? "")
                  .trim(),

              bank_name:
                String(row[14] ?? "")
                  .trim(),

              phone:
                String(row[15] ?? "")
                  .trim(),

              remark:
                String(row[16] ?? "")
                  .trim(),

              resignation_date:
                null,

              status:
                "มีผลประกัน",
              
                insurance_card_no:
                "",

              life_plan:
                "",
            };
          });

      console.log(
        "ข้อมูลแจ้งเข้าหลัง Mapping:",
        mappedEmployees
      );

      // ========================================
// ตรวจว่าพนักงานในไฟล์แจ้งเข้าเป็นคนใหม่หรือไม่
// เทียบกับข้อมูลตั้งต้นที่มีอยู่ใน employees
// ========================================

const existingIdCards = new Set(
  employees
    .map((employee) =>
      String(employee.id_card ?? "")
        .trim()
        .replace(/\D/g, "")
    )
    .filter(Boolean)
);

const newEmployees = mappedEmployees.filter((employee) => {
  const idCard = String(employee.id_card ?? "")
    .trim()
    .replace(/\D/g, "");

  // ไม่มีเลขบัตรประชาชน ไม่ถือว่าเป็นพนักงานใหม่
  if (!idCard) {
    return false;
  }

  // ถ้ามีเลขบัตรนี้อยู่ในข้อมูลตั้งต้นแล้ว = คนเดิม
  return !existingIdCards.has(idCard);
});

const duplicateCount = mappedEmployees.length - newEmployees.length;

console.log("========== ตรวจสอบแจ้งเข้า ==========");
console.log("ข้อมูลในไฟล์แจ้งเข้าทั้งหมด:", mappedEmployees.length);
console.log("พนักงานใหม่:", newEmployees.length);
console.log("มีอยู่แล้วในระบบ:", duplicateCount);
console.log("ข้อมูลพนักงานใหม่:", newEmployees);

// เก็บเฉพาะ "คนใหม่" ไว้สำหรับการบันทึก
setInPreview(newEmployees);

setInRowCount(mappedEmployees.length);
setInNewCount(newEmployees.length);
setInDuplicateCount(duplicateCount);

    } catch (err) {

      console.error(err);

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
      "เกิดข้อผิดพลาดในการอ่านไฟล์แจ้งเข้า"
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
    `บันทึกเข้าสู่ฐานข้อมูลแล้ว ไม่มีพนักงานใหม่ ข้อมูล ${inPreview.length.toLocaleString()} รายการมีอยู่ในระบบแล้ว`
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
  `บันทึกเข้าสู่ฐานข้อมูลแล้ว เพิ่มพนักงานใหม่ ${totalImported.toLocaleString()} คน`
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

  setError("");
  setSuccessMessage("");

  setIsOutLoading(true);

  const reader =
    new FileReader();

  reader.onload = (e) => {

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

      const worksheet =
        workbook.Sheets[
          workbook.SheetNames[0]
        ];

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

      setOutPreview(
        mappedRows
      );

      setOutRowCount(
        mappedRows.length
      );

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

    try {

      // ==========================================
      // โหลดข้อมูลล่าสุดจาก Supabase
      // ==========================================
      const currentEmployees =
        await loadAllEmployees();

      // ==========================================
      // สร้าง Map จาก employee_code
      // ==========================================
      const employeeMap =
        new Map<string, Employee>();

      currentEmployees.forEach(
        (employee) => {

          const code =
            String(
              employee.employee_code ?? ""
            )
              .trim()
              .replace(/\s+/g, "")
              .toUpperCase();

          if (code) {

            employeeMap.set(
              code,
              employee
            );

          }

        }
      );

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

        const employeeCode =
          String(
            row.employee_code ?? ""
          )
            .trim()
            .replace(/\s+/g, "")
            .toUpperCase();

        // หา employee
        const employee =
          employeeMap.get(
            employeeCode
          );

        // ========================================
        // ถ้าไม่พบรหัส
        // ========================================
        if (!employee) {

          notFoundCount++;

          notFoundEmployees.push(
            `${employeeCode} | ${row.title ?? ""}${row.first_name ?? ""} ${row.last_name ?? ""}`
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

            remark:
              row.remark ||
              employee.remark,

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

      setSuccessMessage(
        message
      );

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
      setOutRowCount(null);

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

  reader.onload = (e) => {
    try {
      const data = e.target?.result;

      if (!data) {
        throw new Error("ไม่พบข้อมูลไฟล์รายงานเคลม");
      }

      const workbook = XLSX.read(data, {
        type: "array",
        cellDates: true,
      });

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

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
    hospital_date: String(row[10] ?? "").trim(),

    // จำนวนเงินตามใบเสร็จ
    receipt_amount:
      row[11] !== "" &&
      row[11] !== null &&
      row[11] !== undefined
        ? Number(row[11])
        : 0,

    // จำนวนเงินที่บริษัทประกันจ่าย
    insurance_paid:
      row[12] !== "" &&
      row[12] !== null &&
      row[12] !== undefined
        ? Number(row[12])
        : 0,

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

      setClaimPreview(mappedClaims);
      setClaimRowCount(mappedClaims.length);

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
        .insert(chunk);

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
console.log("========== DEBUG MATCH ==========");

console.log(
  "Employee ตัวอย่าง:",
  employees.slice(0, 10).map((employee) => ({
    ชื่อ: getEmployeeFullName(employee),
    id_card: employee.id_card,
    normalized_id_card: String(employee.id_card ?? "")
      .trim()
      .replace(/\D/g, ""),
  }))
);

console.log(
  "Claim ตัวอย่าง:",
  latestClaims.slice(0, 10).map((claim) => ({
    ชื่อ: claim.employee_name,
    id_card: claim.id_card,
    normalized_id_card: String(claim.id_card ?? "")
      .trim()
      .replace(/\D/g, ""),
  }))
);

console.log("================================");
    
    setClaims(latestClaims);
    // ==========================================
// ตรวจสอบชื่อที่ Match ไม่ได้
// ==========================================

const employeeIdCardSet = new Set(
  employees
    .map((employee) =>
      String(employee.id_card ?? "")
        .trim()
        .replace(/\D/g, "")
    )
    .filter(Boolean)
);

const unmatchedClaims = latestClaims.filter((claim) => {
  // เอาเฉพาะพนักงาน
  if (
    String(claim.relationship ?? "")
      .trim()
      .toUpperCase() !== "E"
  ) {
    return false;
  }

  const claimIdCard = String(
    claim.id_card ?? ""
  )
    .trim()
    .replace(/\D/g, "");

  return (
    claimIdCard === "" ||
    !employeeIdCardSet.has(claimIdCard)
  );
});
// ==========================================
// หา Employee ที่ชื่อใกล้เคียงกับ Claim
// ==========================================

const employeeNameList = employees.map(
  (employee) => ({
    original: getEmployeeFullName(employee),
    normalized: normalizeName(
      getEmployeeFullName(employee)
    ),
    vendor: employee.vendor,
  })
);

const unmatchedWithSuggestions = unmatchedClaims
  .map((claim) => {

    const claimNormalized = normalizeName(
      claim.employee_name
    );

    // หาชื่อที่มีส่วนของชื่อเหมือนกัน
    const suggestions = employeeNameList
      .filter((employee) => {

        if (!claimNormalized) return false;

        return (
          employee.normalized.includes(
            claimNormalized.slice(0, 5)
          ) ||
          claimNormalized.includes(
            employee.normalized.slice(0, 5)
          )
        );
      })
      .slice(0, 5);

    return {
      claimName: claim.employee_name,
      normalized: claimNormalized,
      relationship: claim.relationship,
      claimType: claim.claim_type,
      suggestions,
    };
  });

console.log(
  "===================================="
);

console.log(
  "ชื่อที่ Match ไม่ได้ พร้อมชื่อพนักงานที่ใกล้เคียง:"
);

console.table(
  unmatchedWithSuggestions
);

console.log(
  "===================================="
);

console.log(
  "===================================="
);

console.log(
  "จำนวน Claims ทั้งหมด:",
  latestClaims.length
);

console.log(
  "จำนวน Claims ที่ Match ไม่ได้:",
  unmatchedClaims.length
);

console.log(
  "รายชื่อ Claims ที่ Match ไม่ได้ทั้งหมด:",
  unmatchedClaims.map(
    (claim) => ({
      original: claim.employee_name,
      id_card: claim.id_card,
      normalized: normalizeName(
        claim.employee_name
      ),
      relationship: claim.relationship,
      claimType: claim.claim_type,
      affiliation: claim.affiliation,
    })
  )
);

console.log(
  "===================================="
);
console.log(
  "ตัวอย่าง Claims:",
  latestClaims.slice(0, 20).map((claim) => ({
    ชื่อ: claim.employee_name,
    บัตรประชาชน: claim.id_card,
    ความสัมพันธ์: claim.relationship,
    แผน: claim.plan,
    ประเภทเคลม: claim.claim_type,
    จ่าย: claim.insurance_paid,
  }))
);
console.log(
  "ตัวอย่างชื่อจาก Employees:",
  employees.slice(0, 20).map((employee) => ({
    original: getEmployeeFullName(employee),
    normalized: normalizeName(
      getEmployeeFullName(employee)
    ),
  }))
);

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

  reader.onload = (e) => {
    try {
      const data = e.target?.result;

      if (!data) {
        throw new Error("ไม่พบข้อมูลไฟล์ประกันส่งกลับ");
      }

      const workbook = XLSX.read(data, {
        type: "array",
        cellDates: true,
      });

      const firstSheetName = workbook.SheetNames[0];
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

const dataRows = rows
  .slice(1)
  .filter((row) =>
    row.some(
      (cell: any) =>
        String(cell ?? "").trim() !== ""
    )
  );

console.log(
  "จำนวนข้อมูลประกันส่งกลับ:",
  dataRows.length
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
// จับคู่รหัสพนักงานกับ Master Database
// ==========================================

const employeeMap = new Map<string, Employee>();

employees.forEach((employee) => {
  const code = String(
    employee.employee_code ?? ""
  )
    .trim()
    .replace(/\s+/g, "")
    .toUpperCase();

  if (code !== "") {
    employeeMap.set(code, employee);
  }
});

const matched: Employee[] = [];
const unmatched: Employee[] = [];

mappedEmployees.forEach((insuranceEmployee) => {

  const code = String(
    insuranceEmployee.employee_code ?? ""
  )
    .trim()
    .replace(/\s+/g, "")
    .toUpperCase();

  const masterEmployee =
    employeeMap.get(code);

  if (masterEmployee) {
  matched.push({
    ...masterEmployee,

    plan:
      insuranceEmployee.plan,

    status:
      insuranceEmployee.status,

    resignation_date:
      insuranceEmployee.resignation_date,

    insurance_card_no:
      insuranceEmployee.insurance_card_no,

    life_plan:
      insuranceEmployee.life_plan,
  });
} else {
  unmatched.push(insuranceEmployee);
}
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
    const chunkSize = 500;
    let successCount = 0;

    console.log(
      "🚀 เริ่มบันทึกข้อมูลประกันส่งกลับ:",
      insuranceMatched.length,
      "รายการ"
    );

    for (let i = 0; i < insuranceMatched.length; i += chunkSize) {
      const chunk = insuranceMatched.slice(i, i + chunkSize);

      console.log(
        `💾 กำลังบันทึกชุด ${Math.floor(i / chunkSize) + 1} / ${Math.ceil(
          insuranceMatched.length / chunkSize
        )} (${chunk.length} รายการ)`
      );

      const rows = chunk
        .filter((employee) => employee.id)
        .map((employee) => ({
          id: employee.id,
          employee_key: employee.employee_key,
          employee_code: employee.employee_code,
          vendor: employee.vendor,
          branch: employee.branch,
          title: employee.title,
          first_name: employee.first_name,
          last_name: employee.last_name,
          gender: employee.gender,
          date_of_birth: employee.date_of_birth,
          id_card: employee.id_card,
          employment_date: employee.employment_date,
          effective_date: employee.effective_date,
          plan: employee.plan,
          insurance_type: employee.insurance_type,
          department: employee.department,
          bank_account: employee.bank_account,
          bank_name: employee.bank_name,
          phone: employee.phone,
          remark: employee.remark,
          resignation_date: employee.resignation_date,
          status: employee.status,
          insurance_card_no: employee.insurance_card_no,
          life_plan: employee.life_plan,
          updated_at: new Date().toISOString(),
        }));

      if (rows.length === 0) {
        continue;
      }

      const { error } = await supabase
        .from("employees")
        .upsert(rows, {
          onConflict: "id",
        });

      if (error) {
        throw error;
      }

      successCount += rows.length;
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
  } catch (err) {
    console.error("❌ Insurance Import Error:", err);

    setError(
      err instanceof Error
        ? err.message
        : "ไม่สามารถบันทึกข้อมูลประกันส่งกลับได้"
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
  idCard,
  employee.effective_date
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
    // 5. แบ่งข้อมูลเป็นชุดละ 500
    // ==========================================
    const chunkSize = 500;

    let totalImported = 0;


    for (
      let i = 0;
      i < uniqueEmployees.length;
      i += chunkSize
    ) {

      const chunk =
        uniqueEmployees.slice(
          i,
          i + chunkSize
        );


      console.log(
        `กำลังบันทึก ${
          i + 1
        } - ${
          i + chunk.length
        } / ${
          uniqueEmployees.length
        }`
      );


      // ========================================
      // ใช้ employee_key เป็นตัวหลัก
      // ========================================
      const {
        error
      } = await supabase
        .from("employees")
        .upsert(
          chunk,
          {
            onConflict:
              "employee_key",
          }
        );


      if (error) {

        console.error(
          "Supabase Error:",
          error
        );

        throw new Error(
          `บันทึกข้อมูลชุดที่ ${
            Math.floor(
              i / chunkSize
            ) + 1
          } ไม่สำเร็จ: ${
            error.message
          }`
        );
      }


      totalImported +=
        chunk.length;
    }

    console.log(
    "✅ INSERT สำเร็จ:",
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
              onChange={handleFileUpload}
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

{/* Import Button */}
{!isLoading && employees.length > 0 && (
  <button
    type="button"
    onClick={handleImportToSupabase}
    disabled={isImporting}
    className="mt-4 inline-flex items-center rounded-xl bg-green-600 px-5 py-3 font-medium text-white transition hover:bg-green-700 disabled:cursor-not-allowed disabled:opacity-50"
  >
    {isImporting ? (
      <>
        <span className="mr-2 h-5 w-5 animate-spin rounded-full border-2 border-white border-t-transparent"></span>
        กำลังบันทึกข้อมูล...
      </>
    ) : (
      <>
        💾 บันทึกข้อมูลเข้าระบบ
      </>
    )}
  </button>
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
  <div className="mt-4 grid grid-cols-1 gap-3 sm:grid-cols-2">

    <div className="rounded-xl border border-[#7ED957]/20 bg-[#7ED957]/5 p-4">
      <div className="text-xs text-gray-400">
        พนักงานใหม่
      </div>

      <div className="mt-1 text-2xl font-black text-[#7ED957]">
        {inNewCount?.toLocaleString() ?? 0} คน
      </div>
    </div>

    <div className="rounded-xl border border-white/10 bg-white/5 p-4">
      <div className="text-xs text-gray-400">
        มีอยู่แล้วในข้อมูลตั้งต้น
      </div>

      <div className="mt-1 text-2xl font-black text-white">
        {inDuplicateCount?.toLocaleString() ?? 0} คน
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

      {!isOutLoading &&
  outRowCount !== null && (
    <>
      <div className="mt-3 flex items-center justify-between">
        <p className="font-medium text-red-500">
          ⚠️ แจ้งออก {outRowCount.toLocaleString()} คน
        </p>

        {outPreview.length > 0 && (
          <button
            type="button"
            onClick={() =>
              setShowOutList(!showOutList)
            }
            className="rounded-lg border border-red-200 bg-white px-4 py-2 text-sm font-medium text-red-500 transition hover:bg-red-50"
          >
            {showOutList
              ? "ซ่อนรายชื่อ"
              : "ดูรายชื่อพนักงานออก"}
          </button>
        )}
      </div>

      {showOutList &&
        outPreview.length > 0 && (
          <div className="mt-4 max-h-80 overflow-y-auto rounded-xl border border-slate-200 bg-white">
            <div className="divide-y divide-slate-100">
              {outPreview.map(
                (employee, index) => (
                  <div
                    key={`${employee.employee_code}-${index}`}
                    className="flex items-center justify-between gap-4 px-4 py-3"
                  >
                    <div>
                      <p className="font-medium text-slate-900">
                        {employee.title}{" "}
                        {employee.first_name}{" "}
                        {employee.last_name}
                      </p>

                      <p className="mt-1 text-xs text-slate-500">
                        รหัสพนักงาน:{" "}
                        {employee.employee_code}
                      </p>
                    </div>

                    <span className="shrink-0 rounded-full bg-red-50 px-3 py-1 text-xs font-semibold text-red-500">
                      ลาออก
                    </span>
                  </div>
                )
              )}
            </div>
          </div>
        )}

      {outSuccessMessage && (
        <p className="mt-4 font-medium text-green-600">
          ✓ {outSuccessMessage}
        </p>
      )}

      {outRowCount !== null && (
        <button
          type="button"
          onClick={handleImportOutToSupabase}
          disabled={isOutImporting}
          className="mt-4 inline-flex items-center rounded-xl bg-red-500 px-5 py-3 font-medium text-white transition hover:bg-red-600 disabled:cursor-not-allowed disabled:opacity-50"
        >
          {isOutImporting ? (
            <>
              <span className="mr-2 h-5 w-5 animate-spin rounded-full border-2 border-white border-t-transparent" />
              กำลังแจ้งออก...
            </>
          ) : (
            <>
              💾 บันทึกแจ้งออก
            </>
          )}
        </button>
      )}
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
        <div className="grid gap-5 md:grid-cols-2 lg:grid-cols-5">

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

      <thead className="bg-slate-50 text-sm text-slate-500">
        <tr>

          <th className="px-6 py-4">
            Vendor
          </th>

          <th className="px-6 py-4">
            ลำดับ
          </th>

          <th className="px-6 py-4">
            ชื่อ
          </th>

          <th className="px-6 py-4">
            แผน
          </th>

          <th className="px-6 py-4">
            มีผลประกัน
          </th>

          <th className="px-6 py-4">
            ลาออก
          </th>

          <th className="px-6 py-4">
            สถานะ
          </th>

          <th className="px-6 py-4">
            อยู่มาแล้วกี่วัน
          </th>

          <th className="px-6 py-4">
            ค่าเบี้ยต่อคน
          </th>

          <th className="px-6 py-4">
            ค่าเบี้ยต่อวัน
          </th>

          <th className="px-6 py-4">
            OPD
          </th>

          <th className="px-6 py-4">
            IPD
          </th>

          <th className="px-6 py-4">
            หมายเลขบัตร
          </th>

          <th className="px-6 py-4">
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
const employeeIdCard = String(
  employee.id_card ?? ""
)
  .trim()
  .replace(/\D/g, "");

const claimSummary =
  claimSummaryMap.get(employeeIdCard) ?? {
    opd: 0,
    ipd: 0,
    count: 0,
  };

          if (
            claimSummary.count > 0
          ) {
           console.log(
  "MATCH เจอ:",
  employeeIdCard,
  claimSummary
);
          }
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