"use client";

import { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import { supabase } from "@/lib/supabase";

type Employee = {
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
};
type Claim = {
  policy_name: string;
  reference_no: string;
  claim_status: string;
  relationship: string;
  employee_name: string;
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

export default function Home() {
  const [fileName, setFileName] = useState("");
  const [rowCount, setRowCount] = useState<number | null>(null);
  const [error, setError] = useState("");
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [claims, setClaims] = useState<Claim[]>([]);

  const [isLoading, setIsLoading] = useState(false);
  const [isImporting, setIsImporting] = useState(false);

  const [claimFileName, setClaimFileName] = useState("");
  const [claimRowCount, setClaimRowCount] = useState<number | null>(null);
  const [claimPreview, setClaimPreview] = useState<Claim[]>([]);
  const [isClaimLoading, setIsClaimLoading] = useState(false);
  const [isClaimImporting, setIsClaimImporting] = useState(false);

  const [successMessage, setSuccessMessage] = useState("");
    // =========================
  // สรุปข้อมูลการเคลมของพนักงาน
  // =========================
  const getClaimSummary = (employee: Employee) => {
    const employeeFullName = getEmployeeFullName(employee);

    const employeeClaims = claims.filter(
      (claim) =>
        claim.employee_name
          .trim()
          .replace(/\s+/g, " ")
          .toLowerCase() ===
        employeeFullName
          .trim()
          .replace(/\s+/g, " ")
          .toLowerCase()
    );

    const opdClaims = employeeClaims.filter(
      (claim) =>
        claim.claim_type.trim().toUpperCase() === "OPD"
    );

    const ipdClaims = employeeClaims.filter(
      (claim) =>
        claim.claim_type.trim().toUpperCase() === "IPD"
    );

    const opd = opdClaims.reduce(
      (sum, claim) =>
        sum + Number(claim.insurance_paid || 0),
      0
    );

    const ipd = ipdClaims.reduce(
      (sum, claim) =>
        sum + Number(claim.insurance_paid || 0),
      0
    );

    return {
      opd,
      ipd,
      count: employeeClaims.length,
    };
  };
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
    const {
      data: employeeData,
      error: employeeError,
    } = await supabase
      .from("employees")
      .select("*")
      .order("id", { ascending: true });

    if (employeeError) {
      console.error(
        "Supabase Employee Error:",
        employeeError
      );

      setError(
        `ไม่สามารถโหลดข้อมูลพนักงานได้: ${employeeError.message}`
      );
    } else {
      setEmployees(employeeData ?? []);
      setRowCount(employeeData?.length ?? 0);
    }

    // ==========================================
    // โหลด Claims แยกออกจาก Employees
    // ==========================================
    const {
      data: claimData,
      error: claimError,
    } = await supabase
      .from("claims")
      .select("*");

    if (claimError) {
      console.error(
        "Supabase Claim Error:",
        claimError
      );

      // ไม่ต้อง throw
      // เพราะไม่ควรทำให้ข้อมูลพนักงานพังตาม
      setClaims([]);

    } else {
      setClaims(claimData ?? []);
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
        const employeeCode = String(row[3] ?? "").trim();

        // ต้องมีรหัสพนักงานจริง
        return (
          employeeCode !== "" &&
          employeeCode.toLowerCase() !== "employee id" &&
          employeeCode.toLowerCase() !== "employee code" &&
          employeeCode.toLowerCase() !== "no."
        );
      })
          .map((row) => ({
            employee_code: String(row[3] ?? "").trim(),

            vendor: String(row[1] ?? "").trim(),

            branch: String(row[2] ?? "").trim(),

            title: String(row[4] ?? "").trim(),

            first_name: String(row[5] ?? "").trim(),

            last_name: String(row[6] ?? "").trim(),

            gender: String(row[7] ?? "").trim(),

            date_of_birth: formatExcelDate(row[8]),

            id_card: String(row[9] ?? "").trim(),

            employment_date: formatExcelDate(row[10]),

            effective_date: formatExcelDate(row[11]),

            plan: row[12] !== "" ? Number(row[12]) : null,

            // ค่าเริ่มต้น
            insurance_type: "เต็มปี",

            department: String(row[13] ?? "").trim(),

            bank_account: String(row[14] ?? "").trim(),

            bank_name: String(row[15] ?? "").trim(),

            phone: String(row[16] ?? "").trim(),

            remark: String(row[17] ?? "").trim(),

            resignation_date: formatExcelDate(row[19]),

          status:
          String(row[18] ?? "").trim() !== ""
            ? "ลาออก"
            : "มีผลประกัน",
          }));

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
  // บันทึกข้อมูลลง Supabase
  // =========================
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

      // ==================================================
      // รายงานเคลม:
      //
      // 0 Policy Name
      // 1 เลขที่อ้างอิง
      // 2 สถานะ
      // 3 ความสัมพันธ์
      // 4 ชื่อ-สกุล
      // 5 แผน
      // 6 ประเภท
      // 7 วันเข้า รพ.
      // 8 ใบเสร็จ
      // 9 บริษัทฯจ่าย
      // 10 สถานพยาบาล
      // 11 ประเภทการจ่าย
      // 12 เลขที่รับเรื่อง
      // 13 สังกัด
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

          plan:
            row[5] !== "" &&
            row[5] !== null &&
            row[5] !== undefined
              ? Number(row[5])
              : null,

          claim_type: String(row[6] ?? "").trim(),

          hospital_date: formatExcelDate(row[7]),

          receipt_amount:
            row[8] !== "" &&
            row[8] !== null &&
            row[8] !== undefined
              ? Number(row[8])
              : 0,

          insurance_paid:
            row[9] !== "" &&
            row[9] !== null &&
            row[9] !== undefined
              ? Number(row[9])
              : 0,

          hospital_name: String(row[10] ?? "").trim(),

          payment_type: String(row[11] ?? "").trim(),

          claim_no: String(row[12] ?? "").trim(),

          affiliation: String(row[13] ?? "").trim(),
        }));

      console.log("Claims หลัง Mapping:", mappedClaims);

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
    const { data, error } = await supabase
      .from("claims")
      .insert(claimPreview)
      .select();

    if (error) {
      console.error("Supabase Claims Error:", error);
      throw new Error(
        `บันทึกข้อมูลเคลมไม่สำเร็จ: ${error.message}`
      );
    }

    setClaims((prev) => [
      ...prev,
      ...(data ?? []),
    ]);

    setSuccessMessage(
      `นำเข้ารายงานเคลมสำเร็จ ${claimPreview.length.toLocaleString()} รายการ`
    );

    setClaimPreview([]);

  } catch (err) {
    console.error("Import Claims Error:", err);

    setError(
      err instanceof Error
        ? err.message
        : "ไม่สามารถบันทึกรายงานเคลมได้"
    );
  } finally {
    setIsClaimImporting(false);
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
    // 1. ทำความสะอาดข้อมูล
    // ==========================================
    const cleanedEmployees = employees
      .map((employee) => ({
        ...employee,

        employee_code: String(employee.employee_code ?? "")
          .trim()
          .replace(/\s+/g, "")
          .toUpperCase(),
      }))
      .filter(
        (employee) =>
          employee.employee_code !== ""
      );

    // ==========================================
    // 2. ตรวจสอบ duplicate employee_code
    // ==========================================
    const employeeMap = new Map<
      string,
      Employee
    >();

    const duplicateCodes: string[] = [];

    for (const employee of cleanedEmployees) {
      const code = employee.employee_code;

      if (employeeMap.has(code)) {
        duplicateCodes.push(code);
      }

      employeeMap.set(code, employee);
    }

    const uniqueEmployees =
      Array.from(employeeMap.values());

    // ==========================================
    // 3. แสดงข้อมูลใน Console
    // ==========================================
    console.log(
      "================================="
    );

    console.log(
      "จำนวนจาก Excel:",
      employees.length
    );

    console.log(
      "หลัง Clean:",
      cleanedEmployees.length
    );

    console.log(
      "หลังตัด Duplicate:",
      uniqueEmployees.length
    );

    console.log(
      "Duplicate Codes:",
      duplicateCodes
    );

    console.log(
      "จำนวน Duplicate:",
      duplicateCodes.length
    );

    console.log(
      "================================="
    );

    // ==========================================
    // 4. ถ้ามี Duplicate ให้หยุดก่อน
    // ==========================================
    if (duplicateCodes.length > 0) {

      const uniqueDuplicateCodes =
        [...new Set(duplicateCodes)];

      console.error(
        "พบ employee_code ซ้ำ:",
        uniqueDuplicateCodes
      );

      throw new Error(
        `พบรหัสพนักงานซ้ำในไฟล์ ${uniqueDuplicateCodes.length} รหัส เช่น ${uniqueDuplicateCodes.slice(0, 10).join(", ")}`
      );
    }

    // ==========================================
    // 5. แบ่งข้อมูลเป็นชุดละ 500 รายการ
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
        `กำลังบันทึก ${i + 1} - ${
          i + chunk.length
        } / ${
          uniqueEmployees.length
        }`
      );

      const { error } =
        await supabase
          .from("employees")
          .upsert(chunk, {
            onConflict: "employee_code",
          });

      if (error) {
        console.error(
          "Supabase Error:",
          error
        );

        throw new Error(
          `บันทึกข้อมูลชุดที่ ${
            Math.floor(i / chunkSize) + 1
          } ไม่สำเร็จ: ${error.message}`
        );
      }

      totalImported += chunk.length;
    }

    // ==========================================
    // 6. สำเร็จ
    // ==========================================
    setSuccessMessage(
      `นำเข้าข้อมูลสำเร็จ ${totalImported.toLocaleString()} รายการ`
    );

    console.log(
      "นำเข้าสำเร็จ:",
      totalImported
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

              {/* Success */}
              {!isLoading && rowCount !== null && (
                <p className="mt-2 font-medium text-green-600">
                  ✓ อ่านข้อมูลสำเร็จ{" "}
                  {rowCount.toLocaleString()} รายการ
                </p>
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
            Dashboard
        ========================= */}
        <div className="grid gap-5 md:grid-cols-2 lg:grid-cols-4">

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
              รายการเคลมทั้งหมด
            </p>

            <h2 className="mt-2 text-3xl font-bold text-purple-600">
              {claims.length.toLocaleString()}
            </h2>
          </div>

        </div>

        {/* =========================
            Menu
        ========================= */}
        <div className="mt-8 grid gap-5 md:grid-cols-2 lg:grid-cols-4">

          <button className="rounded-2xl bg-white p-6 text-left shadow-sm transition hover:shadow-md">
            <div className="text-3xl">🟢</div>

            <h2 className="mt-4 text-xl font-bold text-slate-900">
              แจ้งเข้า
            </h2>

            <p className="mt-2 text-sm text-slate-500">
              อัปโหลดรายชื่อพนักงานเข้าใหม่
            </p>
          </button>

          <button className="rounded-2xl bg-white p-6 text-left shadow-sm transition hover:shadow-md">
            <div className="text-3xl">🔴</div>

            <h2 className="mt-4 text-xl font-bold text-slate-900">
              แจ้งออก
            </h2>

            <p className="mt-2 text-sm text-slate-500">
              อัปโหลดรายชื่อพนักงานลาออก
            </p>
          </button>

          <button className="rounded-2xl bg-white p-6 text-left shadow-sm transition hover:shadow-md">
            <div className="text-3xl">🏢</div>

            <h2 className="mt-4 text-xl font-bold text-slate-900">
              ไฟล์จากประกัน
            </h2>

            <p className="mt-2 text-sm text-slate-500">
              อัปโหลดและตรวจสอบข้อมูลจากบริษัทประกัน
            </p>
          </button>

          <div className="rounded-2xl bg-white p-6 text-left shadow-sm transition hover:shadow-md">

  <div className="text-3xl">
    🏥
  </div>

  <h2 className="mt-4 text-xl font-bold text-slate-900">
    รายงานการเคลม
  </h2>

  <p className="mt-2 text-sm text-slate-500">
    อัปโหลดรายงานการเคลมจากบริษัทประกัน
  </p>

  <label
    htmlFor="claim-excel-upload"
    className="mt-5 inline-flex cursor-pointer items-center rounded-xl bg-purple-600 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-purple-700"
  >
    📂 เลือกไฟล์รายงานเคลม
  </label>

  <input
    id="claim-excel-upload"
    type="file"
    accept=".xlsx,.xls"
    onChange={handleClaimFileUpload}
    className="hidden"
  />

  {claimFileName && (
    <div className="mt-4 rounded-xl bg-slate-50 p-4">

      <p className="text-xs text-slate-500">
        ไฟล์ที่เลือก
      </p>

      <p className="mt-1 text-sm font-semibold text-slate-900">
        {claimFileName}
      </p>

      {isClaimLoading && (
        <div className="mt-3 flex items-center gap-3 text-purple-600">

          <span className="h-5 w-5 animate-spin rounded-full border-2 border-purple-600 border-t-transparent"></span>

          <span className="text-sm font-medium">
            กำลังอ่านรายงานเคลม...
          </span>

        </div>
      )}

      {!isClaimLoading && claimRowCount !== null && (
        <p className="mt-2 text-sm font-medium text-green-600">
          ✓ อ่านข้อมูลสำเร็จ{" "}
          {claimRowCount.toLocaleString()} รายการ
        </p>
      )}

      {!isClaimLoading && claimPreview.length > 0 && (
        <button
          type="button"
          onClick={handleImportClaimsToSupabase}
          disabled={isClaimImporting}
          className="mt-4 inline-flex items-center rounded-xl bg-green-600 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-green-700 disabled:cursor-not-allowed disabled:opacity-50"
        >

          {isClaimImporting ? (
            <>
              <span className="mr-2 h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent"></span>
              กำลังบันทึก...
            </>
          ) : (
            <>
              💾 บันทึกเข้าระบบ
            </>
          )}

        </button>
      )}

    </div>
  )}

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

    {employees.length > 0 && (
      <p className="mt-1 text-sm text-slate-500">
        แสดงทั้งหมด{" "}
        {employees.length.toLocaleString()} รายการ
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
            ครั้ง
          </th>

        </tr>
      </thead>

      <tbody>

        {employees.length === 0 ? (

          <tr className="border-t">

            <td
              colSpan={13}
              className="px-6 py-10 text-center text-slate-400"
            >
              {isLoading
                ? "กำลังเตรียมข้อมูล..."
                : "ยังไม่มีข้อมูล กรุณานำเข้าข้อมูลพนักงาน"}
            </td>

          </tr>

        ) : (

          employees.map((employee, index) => {

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
              getClaimSummary(employee);

            return (

              <tr
                key={employee.employee_code || index}
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