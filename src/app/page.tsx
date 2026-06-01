"use client";

import React, { useState, useCallback, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import {
  FileSpreadsheet,
  FileArchive,
  ArrowRight,
  CheckCircle2,
  AlertCircle,
  Download,
  RefreshCw,
  FileText,
  Search,
  Info,
  ShieldCheck,
  FolderOpen,
  Sparkles,
  HelpCircle,
  Database,
  Trash2,
  AlertTriangle,
  Flame,
  Binary
} from "lucide-react";

interface ExcelRow {
  [key: string]: any;
}

interface MatchResult {
  originalName: string;
  originalPath: string;
  nicKey: string;
  empNo: string | null;
  name: string | null;
  phones: string[];
  status: "matched" | "unmatched_zip" | "duplicate_nic" | "duplicate_emp";
  matchType?: "exact" | "substring" | "fuzzy" | "none";
  matchScore?: number;
  details?: string;
}

interface UnmatchedExcelRow {
  nic: string;
  empNo: string;
  rowNumber: number;
  originalRow: any;
}

export default function Home() {
  // Application State
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [zipFile, setZipFile] = useState<File | null>(null);
  const [excelData, setExcelData] = useState<ExcelRow[]>([]);
  const [sheets, setSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>("");
  const [headers, setHeaders] = useState<string[]>([]);
  const [nicColumn, setNicColumn] = useState<string>("");
  const [empColumn, setEmpColumn] = useState<string>("");
  const [nameColumn, setNameColumn] = useState<string>("");
  const [phoneColumn, setPhoneColumn] = useState<string>("");
  const [phoneColumn2, setPhoneColumn2] = useState<string>("");
  
  // ZIP State
  const [zipInstance, setZipInstance] = useState<JSZip | null>(null);
  const [zipFileList, setZipFileList] = useState<{ name: string; path: string; fileObj: JSZip.JSZipObject }[]>([]);
  
  // Processing States
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressText, setProgressText] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const [filterType, setFilterType] = useState<"all" | "nic" | "emp" | "name">("all");
  const [statusFilter, setStatusFilter] = useState<"all" | "matched" | "unmatched_zip" | "unmatched_excel">("all");
  
  // Renaming Settings
  const [includeUnmatched, setIncludeUnmatched] = useState(true);
  const [filenameSuffix, setFilenameSuffix] = useState("");
  const [outputZipName, setOutputZipName] = useState("renamed_pdfs.zip");
  
  // Drop zone drag states
  const [isDragExcel, setIsDragExcel] = useState(false);
  const [isDragZip, setIsDragZip] = useState(false);

  // Tab State for bilingual help
  const [lang, setLang] = useState<"si" | "en">("en");

  // WhatsApp sent counts tracking
  const [whatsappSentCounts, setWhatsappSentCounts] = useState<Record<string, number>>({});

  // Load WhatsApp sent counts from localStorage on mount
  useEffect(() => {
    if (typeof window !== "undefined") {
      const saved = localStorage.getItem("whatsapp_sent_counts");
      if (saved) {
        try {
          setWhatsappSentCounts(JSON.parse(saved));
        } catch (e) {
          console.error("Failed to parse WhatsApp sent counts", e);
        }
      }
    }
  }, []);

  // Update WhatsApp sent count helper
  const updateSentCount = useCallback((path: string) => {
    setWhatsappSentCounts((prev) => {
      const updated = { ...prev, [path]: (prev[path] || 0) + 1 };
      localStorage.setItem("whatsapp_sent_counts", JSON.stringify(updated));
      return updated;
    });
  }, []);

  // Reset WhatsApp sent counts helper
  const handleResetSentCounts = useCallback(() => {
    if (
      confirm(
        lang === "si"
          ? "සියලුම WhatsApp යැවීම් වාර ගණන් බිංදු කිරීමට ඔබට විශ්වාසද?"
          : "Are you sure you want to reset all WhatsApp sent counts?"
      )
    ) {
      setWhatsappSentCounts({});
      localStorage.removeItem("whatsapp_sent_counts");
    }
  }, [lang]);

  // Pass Month (Auto detected based on date but fully editable)
  const [passMonth, setPassMonth] = useState<string>(() => {
    const today = new Date();
    const day = today.getDate();
    let monthIndex = today.getMonth();
    // If today is the 20th or later, target the next month's pass!
    if (day >= 20) {
      monthIndex = (monthIndex + 1) % 12;
    }
    const months = [
      "January", "February", "March", "April", "May", "June", 
      "July", "August", "September", "October", "November", "December"
    ];
    return months[monthIndex];
  });

  // Helper to normalize strings for matching (removes spaces, symbols, lowercase)
  const normalizeNIC = (val: any): string => {
    if (val === undefined || val === null) return "";
    return String(val)
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]/g, ""); // keep only alphanumeric characters
  };

  // Helper to format phone number for WhatsApp API
  const formatPhoneForWhatsApp = (phone: any): string => {
    if (phone === undefined || phone === null) return "";
    let clean = String(phone).replace(/[^0-9]/g, ""); // Keep digits only
    
    // Sri Lankan number format (local 07xxxxxxxx -> international 947xxxxxxxx)
    if (clean.startsWith("0") && clean.length === 10) {
      clean = "94" + clean.substring(1);
    } else if (clean.length === 9 && clean.startsWith("7")) {
      clean = "94" + clean;
    }
    return clean;
  };

  // Helper to extract multiple phone numbers from fields, splitting by common separators
  const extractPhoneNumbers = (val1: any, val2: any): string[] => {
    const rawList: string[] = [];
    
    const addNumbers = (val: any) => {
      if (val === undefined || val === null) return;
      const str = String(val).trim();
      if (!str) return;
      
      // Split by common separators: /, comma, semicolon, |, "and", "or", backslash
      const parts = str.split(/[\/,;|\\]|\band\b|\bor\b/i);
      parts.forEach(p => {
        const cleaned = p.trim();
        if (cleaned && !rawList.includes(cleaned)) {
          rawList.push(cleaned);
        }
      });
    };
    
    addNumbers(val1);
    addNumbers(val2);
    
    return rawList;
  };

  // Helper to validate and identify phone number type for WhatsApp
  const validatePhoneForWhatsApp = (phone: any): { isValid: boolean; type: "mobile" | "landline" | "invalid"; message: string } => {
    if (phone === undefined || phone === null || String(phone).trim() === "") {
      return { isValid: false, type: "invalid", message: "අංකයක් නොමැත / No number" };
    }
    
    let clean = String(phone).replace(/[^0-9]/g, ""); // Keep digits only
    
    if (clean.length === 0) {
      return { isValid: false, type: "invalid", message: "වැරදි අංකයක් / Invalid number" };
    }

    // Sri Lankan mobile: 07xxxxxxx or 947xxxxxxx
    const isSLMobile = (clean.startsWith("07") && clean.length === 10) || 
                       (clean.startsWith("947") && clean.length === 11) ||
                       (clean.length === 9 && clean.startsWith("7"));

    // Sri Lankan landline (starts with 0, not 07, length 10 or starts with 94, not 947, length 11)
    const isSLHome = (clean.startsWith("0") && !clean.startsWith("07") && clean.length === 10) ||
                     (clean.startsWith("94") && !clean.startsWith("947") && clean.length === 11);

    if (isSLMobile) {
      return { isValid: true, type: "mobile", message: "WhatsApp සක්‍රිය විය හැක (Mobile)" };
    }
    
    if (isSLHome) {
      return { isValid: false, type: "landline", message: "ලෑන්ඩ්ලයින් අංකයකි (Landline) - WhatsApp නැත" };
    }

    // Generic international format (8 to 15 digits)
    if (clean.length >= 8 && clean.length <= 15) {
      return { isValid: true, type: "mobile", message: "විදේශීය/වෙනත් අංකයකි (International)" };
    }

    return { isValid: false, type: "invalid", message: "අසම්පූර්ණ අංකයකි / Format error" };
  };

  // Open direct chat to verify if number exists on WhatsApp
  const handleVerifyWhatsApp = (phoneNumber: string) => {
    if (!phoneNumber) {
      alert("දුරකථන අංකයක් හමු නොවුණි. / Phone number not found.");
      return;
    }
    const formattedPhone = formatPhoneForWhatsApp(phoneNumber);
    if (!formattedPhone) {
      alert("දුරකථන අංකය වැරදි ආකෘතියක පවතී. / Invalid phone number format.");
      return;
    }
    
    const verifyUrl = `https://web.whatsapp.com/send?phone=${formattedPhone}`;
    window.open(verifyUrl, "whatsapp_verify");
  };

  // Helper to calculate Levenshtein distance for fuzzy matching
  const getLevenshteinDistance = (a: string, b: string): number => {
    const matrix: number[][] = [];
    for (let i = 0; i <= b.length; i++) {
      matrix[i] = [i];
    }
    for (let j = 0; j <= a.length; j++) {
      matrix[0][j] = j;
    }
    for (let i = 1; i <= b.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        if (b.charAt(i - 1) === a.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1, // substitution
            Math.min(
              matrix[i][j - 1] + 1, // insertion
              matrix[i - 1][j] + 1  // deletion
            )
          );
        }
      }
    }
    return matrix[b.length][a.length];
  };

  // Extract NIC key from a filename (removes path, extension, and standardizes)
  const getNicFromFilename = (filename: string): string => {
    // Get file name without folders
    const baseName = filename.split("/").pop() || filename;
    // Remove extension .pdf
    const nameWithoutExt = baseName.replace(/\.[^/.]+$/, "");
    return normalizeNIC(nameWithoutExt);
  };

  // Reset Excel State
  const handleResetExcel = () => {
    setExcelFile(null);
    setExcelData([]);
    setSheets([]);
    setSelectedSheet("");
    setHeaders([]);
    setNicColumn("");
    setEmpColumn("");
    setNameColumn("");
    setPhoneColumn("");
    setPhoneColumn2("");
  };

  // Reset ZIP State
  const handleResetZip = () => {
    setZipFile(null);
    setZipInstance(null);
    setZipFileList([]);
  };

  // Reset All
  const handleResetAll = () => {
    handleResetExcel();
    handleResetZip();
    setProgress(0);
    setProgressText("");
    setWhatsappSentCounts({});
    localStorage.removeItem("whatsapp_sent_counts");
  };

  // Parse Excel File
  const parseExcelFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        setSheets(workbook.SheetNames);
        
        if (workbook.SheetNames.length > 0) {
          const defaultSheet = workbook.SheetNames[0];
          setSelectedSheet(defaultSheet);
          processSheet(workbook.Sheets[defaultSheet]);
        }
        setExcelFile(file);
      } catch (error) {
        alert("Excel ගොනුව කියවීමේදී දෝෂයක් ඇතිවිය. / Error reading Excel file: " + (error as Error).message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Process a selected Excel sheet
  const processSheet = (sheet: XLSX.WorkSheet) => {
    // Convert to JSON with header rows
    const rows = XLSX.utils.sheet_to_json<any>(sheet, { defval: "" });
    if (rows.length === 0) {
      alert("තෝරාගත් Sheet එකෙහි දත්ත කිසිවක් හමුනොවිය. / No data found in the selected sheet.");
      return;
    }

    setExcelData(rows);
    
    // Extract headers from the keys of the first row
    const extractedHeaders = Object.keys(rows[0]);
    setHeaders(extractedHeaders);

    // Auto-detect columns (NIC, Employee No, Name, Phone 1, Phone 2)
    let autoNic = "";
    let autoEmp = "";
    let autoName = "";
    let autoPhone = "";

    const nicKeywords = ["nic", "identity", "id", "හඳුනුම්පත්", "ජා.හැ.", "card"];
    const empKeywords = ["emp", "employee", "member", "නොම්මර", "අංකය", "id", "code", "no"];
    const nameKeywords = ["name", "nama", "නම", "සේවක නම", "employee name", "emp name", "full name"];
    const phoneKeywords = ["phone", "mobile", "contact", "tel", "telephone", "දුරකථන", "whatsapp", "no", "number", "contact no"];

    // First try exact matches, then keyword contains
    for (const header of extractedHeaders) {
      const lowerHeader = header.toLowerCase();
      
      // Look for NIC
      if (!autoNic) {
        if (lowerHeader === "nic" || lowerHeader === "nic no" || lowerHeader === "nic number") {
          autoNic = header;
        }
      }
      
      // Look for Employee Number
      if (!autoEmp) {
        if (
          lowerHeader === "emp no" || 
          lowerHeader === "emp number" || 
          lowerHeader === "employee number" || 
          lowerHeader === "employee no" ||
          lowerHeader === "employee id" ||
          lowerHeader === "emp_id" ||
          lowerHeader === "emp_no"
        ) {
          autoEmp = header;
        }
      }

      // Look for Name
      if (!autoName) {
        if (
          lowerHeader === "name" ||
          lowerHeader === "employee name" ||
          lowerHeader === "emp name" ||
          lowerHeader === "nama"
        ) {
          autoName = header;
        }
      }

      // Look for Phone
      if (!autoPhone) {
        if (
          lowerHeader === "phone" ||
          lowerHeader === "phone no" ||
          lowerHeader === "phone number" ||
          lowerHeader === "mobile" ||
          lowerHeader === "mobile no" ||
          lowerHeader === "mobile number" ||
          lowerHeader === "whatsapp" ||
          lowerHeader === "contact" ||
          lowerHeader === "contact no"
        ) {
          autoPhone = header;
        }
      }
    }

    // Fuzzy matching if not found
    if (!autoNic) {
      autoNic = extractedHeaders.find(h => 
        nicKeywords.some(keyword => h.toLowerCase().includes(keyword))
      ) || "";
    }
    
    if (!autoEmp) {
      autoEmp = extractedHeaders.find(h => 
        empKeywords.some(keyword => h.toLowerCase().includes(keyword)) && h !== autoNic
      ) || "";
    }

    if (!autoName) {
      autoName = extractedHeaders.find(h => 
        nameKeywords.some(keyword => h.toLowerCase().includes(keyword)) && h !== autoNic && h !== autoEmp
      ) || "";
    }

    if (!autoPhone) {
      autoPhone = extractedHeaders.find(h => 
        phoneKeywords.some(keyword => h.toLowerCase().includes(keyword)) && h !== autoNic && h !== autoEmp && h !== autoName
      ) || "";
    }

    // Look for Phone 2
    let autoPhone2 = "";
    for (const header of extractedHeaders) {
      if (header === autoPhone) continue;
      const lowerHeader = header.toLowerCase();
      if (
        lowerHeader === "phone 2" ||
        lowerHeader === "phone2" ||
        lowerHeader === "mobile 2" ||
        lowerHeader === "mobile2" ||
        lowerHeader === "alt phone" ||
        lowerHeader === "alternative phone" ||
        lowerHeader === "home" ||
        lowerHeader === "home phone" ||
        lowerHeader === "landline" ||
        lowerHeader === "telephone"
      ) {
        autoPhone2 = header;
        break;
      }
    }

    if (!autoPhone2) {
      autoPhone2 = extractedHeaders.find(h => 
        phoneKeywords.some(keyword => h.toLowerCase().includes(keyword)) && 
        h !== autoNic && h !== autoEmp && h !== autoName && h !== autoPhone
      ) || "";
    }

    // Fallbacks
    setNicColumn(autoNic || extractedHeaders[0] || "");
    setEmpColumn(autoEmp || (extractedHeaders[1] !== autoNic ? extractedHeaders[1] : extractedHeaders[0]) || "");
    setNameColumn(autoName || extractedHeaders.find(h => h !== autoNic && h !== autoEmp) || "");
    setPhoneColumn(autoPhone || extractedHeaders.find(h => h !== autoNic && h !== autoEmp && h !== autoName) || "");
    setPhoneColumn2(autoPhone2 || "");
  };

  // Handle Sheet Change
  const handleSheetChange = (sheetName: string) => {
    if (!excelFile) return;
    setSelectedSheet(sheetName);
    
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[sheetName];
      processSheet(sheet);
    };
    reader.readAsArrayBuffer(excelFile);
  };

  // Parse ZIP File
  const parseZipFile = async (file: File) => {
    try {
      setZipFile(file);
      const zip = new JSZip();
      const loadedZip = await zip.loadAsync(file);
      setZipInstance(loadedZip);

      const files: { name: string; path: string; fileObj: JSZip.JSZipObject }[] = [];
      loadedZip.forEach((relativepath, fileObj) => {
        // Only include PDF files and ignore directory markers
        if (!fileObj.dir && relativepath.toLowerCase().endsWith(".pdf")) {
          // Get filename
          const filename = relativepath.split("/").pop() || relativepath;
          files.push({
            name: filename,
            path: relativepath,
            fileObj: fileObj
          });
        }
      });

      setZipFileList(files);
    } catch (error) {
      alert("ZIP ගොනුව කියවීමේදී දෝෂයක් ඇතිවිය. / Error reading ZIP file: " + (error as Error).message);
      setZipFile(null);
    }
  };

  // Drop Event Handlers for Excel
  const handleDragExcel = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragExcel(e.type === "dragover" || e.type === "dragenter");
  };

  const handleDropExcel = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragExcel(false);
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      const file = files[0];
      if (file.name.endsWith(".xlsx") || file.name.endsWith(".xls") || file.name.endsWith(".csv")) {
        parseExcelFile(file);
      } else {
        alert("කරුණාකර වලංගු Excel ගොනුවක් ලබාදෙන්න (.xlsx, .xls). / Please upload a valid Excel file.");
      }
    }
  };

  const handleFileChangeExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      parseExcelFile(files[0]);
    }
  };

  // Drop Event Handlers for ZIP
  const handleDragZip = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragZip(e.type === "dragover" || e.type === "dragenter");
  };

  const handleDropZip = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragZip(false);
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      const file = files[0];
      if (file.name.endsWith(".zip")) {
        parseZipFile(file);
      } else {
        alert("කරුණාකර වලංගු ZIP ගොනුවක් ලබාදෙන්න (.zip). / Please upload a valid ZIP file.");
      }
    }
  };

  const handleFileChangeZip = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      parseZipFile(files[0]);
    }
  };

  // Computations for Mappings & Matches
  const matchingData = useMemo(() => {
    if (!nicColumn || !empColumn) return { matchedList: [], unmatchedExcelList: [], stats: { matched: 0, unmatchedZip: 0, unmatchedExcel: 0, totalZip: 0 } };

    // 1. Extract and normalize Excel data into an array of entries
    const excelEntries: { 
      normalizedNic: string; 
      empNo: string; 
      name: string; 
      phones: string[];
      row: ExcelRow; 
      rowIdx: number; 
      originalNic: string; 
      matched: boolean 
    }[] = [];

    excelData.forEach((row, index) => {
      const rawNic = row[nicColumn];
      const rawEmp = row[empColumn];
      const rawName = nameColumn ? row[nameColumn] : "";
      const rawPhone = phoneColumn ? row[phoneColumn] : "";
      const rawPhone2 = phoneColumn2 ? row[phoneColumn2] : "";
      
      if (rawNic !== undefined && rawNic !== null && String(rawNic).trim() !== "") {
        const nicKey = normalizeNIC(rawNic);
        const empNo = String(rawEmp).trim();
        const name = String(rawName).trim();
        const phones = extractPhoneNumbers(rawPhone, rawPhone2);
        
        excelEntries.push({
          normalizedNic: nicKey,
          empNo: empNo,
          name: name,
          phones: phones,
          row: row,
          rowIdx: index + 2, // 1-indexed + header row
          originalNic: String(rawNic).trim(),
          matched: false
        });
      }
    });

    // 2. Map ZIP Files using multi-tier smart match (Exact, Sub-string/Digits, Fuzzy)
    const matchedList: MatchResult[] = [];
    const matchedExcelNicsSeen = new Set<string>();
    const empNosUsed = new Set<string>();

    zipFileList.forEach((zipFileItem) => {
      const nicKey = getNicFromFilename(zipFileItem.name);
      
      if (!nicKey) {
        matchedList.push({
          originalName: zipFileItem.name,
          originalPath: zipFileItem.path,
          nicKey: "",
          empNo: null,
          name: null,
          phones: [],
          status: "unmatched_zip",
          matchType: "none",
          details: "ගොනු නාමයෙන් NIC එකක් හඳුනාගත නොහැක / Cannot extract NIC from filename"
        });
        return;
      }

      // Try to find the best match for this ZIP file
      let bestMatch: typeof excelEntries[0] | null = null;
      let matchType: "exact" | "substring" | "fuzzy" | "none" = "none";
      let matchScore = 0;

      // Stage 1: Exact Match
      const exactMatch = excelEntries.find((e) => e.normalizedNic === nicKey);
      if (exactMatch) {
        bestMatch = exactMatch;
        matchType = "exact";
      } else {
        // Stage 2: Sub-string & Digit-only Match
        const substringMatches = excelEntries.filter((e) => {
          const eNic = e.normalizedNic;
          const zNic = nicKey;

          // Standard substring contains
          if (eNic.length >= 5 && zNic.length >= 5) {
            if (zNic.includes(eNic) || eNic.includes(zNic)) {
              return true;
            }
          }

          // Digit-only comparison (handles missing 'V' or extra characters)
          const eDigits = eNic.replace(/[^0-9]/g, "");
          const zDigits = zNic.replace(/[^0-9]/g, "");
          if (eDigits.length >= 7 && zDigits.length >= 7) {
            if (eDigits === zDigits || zDigits.includes(eDigits) || eDigits.includes(zDigits)) {
              return true;
            }
          }

          return false;
        });

        if (substringMatches.length > 0) {
          // Pick the one with closest length
          substringMatches.sort((a, b) => Math.abs(a.normalizedNic.length - nicKey.length) - Math.abs(b.normalizedNic.length - nicKey.length));
          bestMatch = substringMatches[0];
          matchType = "substring";
        } else {
          // Stage 3: Fuzzy Match (Levenshtein Distance <= 3)
          let minDistance = 999;
          let bestFuzzy: typeof excelEntries[0] | null = null;

          excelEntries.forEach((e) => {
            if (e.normalizedNic.length >= 5 && nicKey.length >= 5) {
              const distance = getLevenshteinDistance(nicKey, e.normalizedNic);
              if (distance < minDistance) {
                minDistance = distance;
                bestFuzzy = e;
              }
            }
          });

          if (bestFuzzy && minDistance <= 3) {
            bestMatch = bestFuzzy;
            matchType = "fuzzy";
            matchScore = minDistance;
          }
        }
      }

      // Check duplicates & finalize status
      if (bestMatch) {
        const targetEmpNo = bestMatch.empNo;
        const targetNic = bestMatch.normalizedNic;
        const targetName = bestMatch.name;
        bestMatch.matched = true; // Mark as matched

        if (matchedExcelNicsSeen.has(targetNic)) {
          matchedList.push({
            originalName: zipFileItem.name,
            originalPath: zipFileItem.path,
            nicKey: nicKey,
            empNo: targetEmpNo,
            name: targetName,
            phones: bestMatch.phones,
            status: "duplicate_nic",
            matchType: matchType,
            matchScore: matchScore,
            details: `Excel පේළිය (${bestMatch.originalNic}) දැනටමත් වෙනත් ගොනුවකට ගළපා ඇත / Excel row already matched`
          });
        } else if (empNosUsed.has(targetEmpNo)) {
          matchedExcelNicsSeen.add(targetNic);
          matchedList.push({
            originalName: zipFileItem.name,
            originalPath: zipFileItem.path,
            nicKey: nicKey,
            empNo: targetEmpNo,
            name: targetName,
            phones: bestMatch.phones,
            status: "duplicate_emp",
            matchType: matchType,
            matchScore: matchScore,
            details: `Employee Number (${targetEmpNo}) දැනටමත් වෙනත් NIC එකකට භාවිතා කර ඇත / Employee No already used`
          });
        } else {
          matchedExcelNicsSeen.add(targetNic);
          empNosUsed.add(targetEmpNo);
          matchedList.push({
            originalName: zipFileItem.name,
            originalPath: zipFileItem.path,
            nicKey: nicKey,
            empNo: targetEmpNo,
            name: targetName,
            phones: bestMatch.phones,
            status: "matched",
            matchType: matchType,
            matchScore: matchScore,
            details: matchType === "exact"
              ? "100% ගැලපේ (Exact Match)"
              : matchType === "substring"
                ? `පූර්ණ නොවන ගැලපීම (Partial Match: ${bestMatch.originalNic})`
                : `සැක සහිත ගැලපීම (Fuzzy Match: ${bestMatch.originalNic}, වෙනස: ${matchScore})`
          });
        }
      } else {
        matchedList.push({
          originalName: zipFileItem.name,
          originalPath: zipFileItem.path,
          nicKey: nicKey,
          empNo: null,
          name: null,
          phones: [],
          status: "unmatched_zip",
          matchType: "none",
          details: "Excel පත්‍රයේ ගැලපෙන අගයක් හමුනොවිය / No match found in Excel"
        });
      }
    });

    // 3. Find unmatched Excel Rows
    const unmatchedExcelList: UnmatchedExcelRow[] = [];
    excelEntries.forEach((val) => {
      if (!val.matched) {
        unmatchedExcelList.push({
          nic: val.normalizedNic,
          empNo: val.empNo,
          rowNumber: val.rowIdx,
          originalRow: val.row
        });
      }
    });

    // Stats
    const stats = {
      matched: matchedList.filter((m) => m.status === "matched").length,
      unmatchedZip: matchedList.filter((m) => m.status !== "matched").length,
      unmatchedExcel: unmatchedExcelList.length,
      totalZip: zipFileList.length
    };

    return { matchedList, unmatchedExcelList, stats };
  }, [excelData, zipFileList, nicColumn, empColumn, nameColumn, phoneColumn, phoneColumn2]);

  // Filtered lists for rendering preview
  const filteredList = useMemo(() => {
    const { matchedList } = matchingData;
    let items = matchedList;

    if (statusFilter === "matched") {
      items = matchedList.filter((m) => m.status === "matched");
    } else if (statusFilter === "unmatched_zip") {
      items = matchedList.filter((m) => m.status !== "matched");
    }

    if (searchQuery.trim() !== "") {
      const q = searchQuery.toLowerCase();
      const normQ = normalizeNIC(q);
      items = items.filter((m) => {
        const matchesNic = m.nicKey.toLowerCase().includes(q) || (normQ && m.nicKey.toLowerCase().includes(normQ));
        const matchesEmp = m.empNo ? m.empNo.toLowerCase().includes(q) : false;
        const matchesName = m.name ? m.name.toLowerCase().includes(q) : false;
        const matchesOrigName = m.originalName.toLowerCase().includes(q);

        if (filterType === "nic") return matchesNic;
        if (filterType === "emp") return matchesEmp;
        if (filterType === "name") return matchesName;
        
        // "all"
        return matchesNic || matchesEmp || matchesName || matchesOrigName;
      });
    }

    return items;
  }, [matchingData, statusFilter, searchQuery, filterType]);

  // Send renamed PDF file to WhatsApp
  const handleSendWhatsApp = async (item: MatchResult, phoneNumber: string) => {
    if (!phoneNumber) {
      alert("දුරකථන අංකයක් හමු නොවුණි. / Phone number not found.");
      return;
    }

    const formattedPhone = formatPhoneForWhatsApp(phoneNumber);
    if (!formattedPhone) {
      alert("දුරකථන අංකය වැරදි ආකෘතියක පවතී. / Invalid phone number format.");
      return;
    }

    const fileName = item.empNo 
      ? `${item.empNo} (${item.nicKey.toUpperCase()})${filenameSuffix}.pdf` 
      : `${item.nicKey.toUpperCase()}${filenameSuffix}.pdf`;
    
    // Find JSZipObject
    const origZipFile = zipFileList.find((f) => f.path === item.originalPath);
    if (!origZipFile) {
      alert("අදාළ PDF ගොනුව ZIP එකෙහි හමු නොවුණි. / File not found in ZIP archive.");
      return;
    }

    setIsProcessing(true);
    setProgress(30);
    setProgressText(`ගොනුව සකසමින්: ${fileName}... / Extracting file: ${fileName}...`);

    try {
      // Extract binary data of PDF
      const content = await origZipFile.fileObj.async("uint8array");
      const blob = new Blob([content as any], { type: "application/pdf" });
      
      setProgress(70);
      setProgressText("WhatsApp සම්බන්ධ කරමින්... / Directing to WhatsApp...");

      // Download file locally
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      // Update download and sent status count
      updateSentCount(item.originalPath);
      
      // Open WhatsApp Desktop App directly using the custom whatsapp:// URI scheme!
      const whatsappText = lang === "si"
        ? `ආයුබෝවන් ${item.name || ""},\n\nමෙන්න ඔබගේ ${passMonth} මාසික බලපත්‍රය (monthly pass).\n\nනම (Name): ${item.name || ""}\nසේවක අංකය (Employee ID): ${item.empNo || ""}`
        : `Hello ${item.name || ""},\n\nHere is your ${passMonth} monthly pass.\n\nName: ${item.name || ""}\nEmployee ID: ${item.empNo || ""}`;
      const whatsappUrl = `whatsapp://send?phone=${formattedPhone}&text=${encodeURIComponent(whatsappText)}`;
      window.location.href = whatsappUrl;
      
      setProgress(100);
      setTimeout(() => {
        setIsProcessing(false);
        setProgress(0);
        setProgressText("");
      }, 1000);
      
    } catch (err) {
      alert("WhatsApp හරහා යැවීමේදී දෝෂයක් ඇති විය / Error sharing via WhatsApp: " + (err as Error).message);
      setIsProcessing(false);
      setProgress(0);
      setProgressText("");
    }
  };

  // Generate Renamed ZIP File
  const handleRenameAndDownload = async () => {
    if (!zipInstance || (matchingData.stats.matched === 0 && !includeUnmatched)) {
      alert("බාගත කිරීමට ගොනු නොමැත. / No files to download.");
      return;
    }

    setIsProcessing(true);
    setProgress(5);
    setProgressText("නව ZIP ගොනුව සකසමින් පවතී... / Initializing new ZIP archive...");

    try {
      const newZip = new JSZip();
      const { matchedList } = matchingData;
      const totalFiles = matchedList.length;

      let processedCount = 0;

      for (let i = 0; i < totalFiles; i++) {
        const item = matchedList[i];
        
        // Find JSZipObject
        const origZipFile = zipFileList.find((f) => f.path === item.originalPath);

        if (!origZipFile) continue;

        const isMatched = item.status === "matched";

        if (isMatched && item.empNo) {
          const newName = `${item.empNo} (${item.nicKey.toUpperCase()})${filenameSuffix}.pdf`;
          setProgressText(`පරිවර්තනය කරමින්: ${item.originalName} -> ${newName}`);
          
          // Get PDF Blob or Uint8Array
          const content = await origZipFile.fileObj.async("uint8array");
          
          // Add to new zip
          newZip.file(newName, content);
          processedCount++;
        } else if (includeUnmatched) {
          setProgressText(`එලෙසම ඇතුලත් කරමින්: ${item.originalName}`);
          const content = await origZipFile.fileObj.async("uint8array");
          newZip.file(item.originalName, content);
          processedCount++;
        }

        // Update progress dynamically
        const progressVal = Math.floor((i / totalFiles) * 70) + 10;
        setProgress(progressVal);
      }

      setProgress(85);
      setProgressText("ZIP ගොනුව සම්පීඩනය කරමින්... (මේ සඳහා සුළු වේලාවක් ගතවිය හැක) / Compressing ZIP... (This may take a moment)");

      // Generate the ZIP blob
      const contentBlob = await newZip.generateAsync({ type: "blob" }, (metadata) => {
        const zipProgress = Math.floor(metadata.percent * 0.15) + 85;
        setProgress(zipProgress);
      });

      setProgress(100);
      setProgressText("පරිවර්තනය සාර්ථකයි! බාගත කිරීම ආරම්භ වේ... / Rename successful! Download starting...");

      // Download trigger
      const link = document.createElement("a");
      link.href = URL.createObjectURL(contentBlob);
      link.download = outputZipName.endsWith(".zip") ? outputZipName : `${outputZipName}.zip`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      setTimeout(() => {
        setIsProcessing(false);
        setProgress(0);
        setProgressText("");
      }, 3000);

    } catch (err) {
      alert("ZIP ගොනුව සැකසීමේදී දෝෂයක් සිදු විය / Error during ZIP processing: " + (err as Error).message);
      setIsProcessing(false);
      setProgress(0);
      setProgressText("");
    }
  };

  // Generate Sample Demo Files so the user can easily test the app!
  const generateDemoFiles = async () => {
    setIsProcessing(true);
    setProgress(10);
    setProgressText("ආදර්ශ දත්ත සකසමින්... / Creating demo Excel data...");

    try {
      // 1. Create Demo Excel Data
      const demoData = [
        { "Employee Name": "Chathura Prasad", "NIC Number": "951234567V", "Employee Number": "EMP-2026-001", "Phone Number": "0771234567", "Alt Phone": "0788888888", "Department": "IT" },
        { "Employee Name": "Nipuni Silva", "NIC Number": "978564321V", "Employee Number": "EMP-2026-002", "Phone Number": "0785643210", "Alt Phone": "0112345678", "Department": "HR" },
        { "Employee Name": "Kasun Perera", "NIC Number": "199212345678", "Employee Number": "EMP-2026-003", "Phone Number": "0712345678", "Alt Phone": "0723333333", "Department": "Finance" },
        { "Employee Name": "Ruwan Fernando", "NIC Number": "908765432V", "Employee Number": "EMP-2026-004", "Phone Number": "0768765432", "Alt Phone": "", "Department": "Operations" },
        { "Employee Name": "Sanduni Jayasinghe", "NIC Number": "981230987V", "Employee Number": "EMP-2026-005", "Phone Number": "0771230987", "Alt Phone": "0794444444", "Department": "Marketing" }
      ];

      const worksheet = XLSX.utils.json_to_sheet(demoData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Employees");
      
      const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      const excelBlob = new Blob([excelBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      
      setProgress(40);
      setProgressText("ආදර්ශ PDF සහිත ZIP ගොනුවක් සකසමින්... / Creating demo PDFs in ZIP...");

      // 2. Create Demo ZIP file containing mock PDFs with NIC filenames
      const zip = new JSZip();
      
      // We will create 5 small text/pdf-like empty mock documents
      const pdfMockContent = "%PDF-1.4 Mock PDF Content representing salary slip or certificate for testing purposes.";
      
      zip.file("951234567V.pdf", pdfMockContent);
      zip.file("978564321v.pdf", pdfMockContent); // case variation test
      zip.file("199212345678.pdf", pdfMockContent); // 12-digit test
      zip.file("908765432V.pdf", pdfMockContent);
      zip.file("999999999V.pdf", pdfMockContent); // unmatched in Excel test
      
      const zipBlob = await zip.generateAsync({ type: "blob" });

      setProgress(80);
      setProgressText("ගොනු බාගත කරමින්... / Downloading mock files...");

      // Download Excel
      const excelLink = document.createElement("a");
      excelLink.href = URL.createObjectURL(excelBlob);
      excelLink.download = "sample_employee_list.xlsx";
      document.body.appendChild(excelLink);
      excelLink.click();
      document.body.removeChild(excelLink);

      // Download Zip
      const zipLink = document.createElement("a");
      zipLink.href = URL.createObjectURL(zipBlob);
      zipLink.download = "sample_nic_pdfs.zip";
      document.body.appendChild(zipLink);
      zipLink.click();
      document.body.removeChild(zipLink);

      setProgress(100);
      setProgressText("සාර්ථකයි! ගොනු 2ක් බාගත විය. / Success! 2 mock files downloaded. Drag and drop them to test!");

      setTimeout(() => {
        setIsProcessing(false);
        setProgress(0);
        setProgressText("");
      }, 4000);

    } catch (err) {
      alert("Error generating demo data: " + (err as Error).message);
      setIsProcessing(false);
    }
  };

  return (
    <div className="flex-1 min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans select-none overflow-x-hidden selection:bg-teal-500 selection:text-slate-900">
      
      {/* Background Ambient Glows */}
      <div className="absolute top-0 left-1/4 w-96 h-96 bg-teal-500/10 rounded-full blur-[120px] pointer-events-none" />
      <div className="absolute top-1/3 right-1/4 w-96 h-96 bg-emerald-500/10 rounded-full blur-[120px] pointer-events-none" />
      <div className="absolute bottom-10 left-1/3 w-96 h-96 bg-cyan-500/10 rounded-full blur-[120px] pointer-events-none" />

      {/* Premium Header */}
      <header className="border-b border-slate-800/80 bg-slate-900/60 backdrop-blur-md sticky top-0 z-50 transition-all duration-300">
        <div className="max-w-[95%] 2xl:max-w-[1650px] mx-auto px-4 sm:px-6 lg:px-8 h-20 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="h-11 w-11 rounded-xl bg-gradient-to-tr from-teal-500 to-emerald-400 p-0.5 shadow-lg shadow-teal-500/20 flex items-center justify-center">
              <div className="h-full w-full rounded-[10px] bg-slate-900 flex items-center justify-center">
                <Binary className="h-5 w-5 text-teal-400 animate-pulse" />
              </div>
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xl font-bold tracking-tight bg-gradient-to-r from-white via-slate-100 to-teal-400 bg-clip-text text-transparent">
                  ZIP Rename Nexus
                </span>
                <span className="text-[10px] px-2 py-0.5 rounded-full bg-teal-500/10 border border-teal-500/20 text-teal-400 font-medium">
                  v1.2
                </span>
              </div>
              <p className="text-xs text-slate-400 font-light mt-0.5">
                NIC to Employee Code PDF Smart Relabeler
              </p>
            </div>
          </div>

          <div className="flex items-center gap-4">
            {/* Language Switcher */}
            <div className="bg-slate-900 border border-slate-800 p-1 rounded-lg flex gap-1">
              <button 
                onClick={() => setLang("si")}
                className={`text-xs px-3 py-1.5 rounded-md font-medium transition-all duration-200 ${
                  lang === "si" 
                    ? "bg-teal-500 text-slate-950 shadow-md shadow-teal-500/20 font-bold" 
                    : "text-slate-400 hover:text-slate-200"
                }`}
              >
                සිංහල
              </button>
              <button 
                onClick={() => setLang("en")}
                className={`text-xs px-3 py-1.5 rounded-md font-medium transition-all duration-200 ${
                  lang === "en" 
                    ? "bg-teal-500 text-slate-950 shadow-md shadow-teal-500/20 font-bold" 
                    : "text-slate-400 hover:text-slate-200"
                }`}
              >
                English
              </button>
            </div>

            {/* Test Demo Button */}
            <button
              onClick={generateDemoFiles}
              disabled={isProcessing}
              className="hidden sm:flex items-center gap-2 text-xs font-semibold px-4 py-2 rounded-lg bg-slate-800 hover:bg-slate-700/80 border border-slate-700/80 transition-all duration-200 disabled:opacity-50 text-teal-400 hover:text-teal-300 active:scale-95"
            >
              <Sparkles className="h-3.5 w-3.5" />
              {lang === "si" ? "ආදර්ශ ගොනු ලබාගන්න" : "Get Sample Files"}
            </button>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 max-w-[95%] 2xl:max-w-[1650px] w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 flex flex-col gap-8 relative z-10">
        
        {/* Bilingual Summary Explanation banner */}
        <div className="bg-gradient-to-r from-teal-950/40 to-emerald-950/20 border border-teal-500/20 rounded-2xl p-5 shadow-xl relative overflow-hidden backdrop-blur-sm">
          <div className="absolute right-0 top-0 translate-x-12 -translate-y-12 w-48 h-48 bg-teal-500/5 rounded-full blur-3xl pointer-events-none" />
          <div className="flex gap-4 items-start">
            <div className="p-3 bg-teal-500/10 border border-teal-500/20 rounded-xl text-teal-400 shadow-inner mt-1">
              <Info className="h-5 w-5" />
            </div>
            <div className="flex-1">
              {lang === "si" ? (
                <>
                  <h2 className="text-base font-semibold text-teal-300">භාවිතා කරන ආකාරය (How to Use)</h2>
                  <p className="text-sm text-slate-300 mt-1 leading-relaxed">
                    1. සේවක අංකය (Employee No) සහ ජාතික හැඳුනුම්පත් අංකය (NIC) සහිත <strong>Excel ගොනුව</strong> පළමුව උඩුගත කරන්න.<br />
                    2. PDF ගොනු ඇතුලත් <strong>ZIP ගොනුව</strong> උඩුගත කරන්න (එහි PDF වල නම ලෙස NIC අංකය තිබිය යුතුය).<br />
                    3. මෙම පද්ධතිය මඟින් ස්වයංක්‍රීයව NIC අංක ගළපා PDF ගොනු වල නම සේවක අංකයට වෙනස් කර නව ZIP ගොනුවක් සකසා දෙනු ඇත.<br />
                    4. 💡 <strong>WhatsApp සත්‍යාපනය (Verification)</strong>: ඕනෑම දුරකථන අංකයක් ඉදිරියෙන් ඇති 🔍 බොත්තම ක්ලික් කිරීමෙන් එම අංකයට WhatsApp ගිණුමක් සක්‍රීයව පවතීද යන්න බ්‍රවුසරය හරහාම නොමිලේ පරීක්ෂා කර බැලිය හැක.
                  </p>
                </>
              ) : (
                <>
                  <h2 className="text-base font-semibold text-teal-300">How It Works</h2>
                  <p className="text-sm text-slate-300 mt-1 leading-relaxed">
                    1. Upload your <strong>Excel sheet</strong> containing both Employee Numbers and NIC Numbers.<br />
                    2. Upload your <strong>ZIP file</strong> containing PDFs currently named with NIC numbers (e.g. <code className="bg-slate-900 px-1 py-0.5 rounded text-teal-400 font-mono text-xs">951234567V.pdf</code>).<br />
                    3. The app will automatically cross-match NICs, map them to their corresponding Employee Codes, rename the PDFs, and generate a new ZIP to download!<br />
                    4. 💡 <strong>WhatsApp Verification</strong>: Click the 🔍 button next to any number to instantly check if the number is registered on WhatsApp for free without expensive APIs.
                  </p>
                </>
              )}
              
              <div className="mt-3 flex sm:hidden">
                <button
                  onClick={generateDemoFiles}
                  className="flex items-center gap-1.5 text-xs font-semibold text-teal-400 hover:text-teal-300"
                >
                  <Sparkles className="h-3 w-3" />
                  {lang === "si" ? "ආදර්ශ ගොනු ලබාගෙන පරීක්ෂා කරන්න" : "Download sample test files"}
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* File Dropping Section */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          
          {/* File Input 1: Excel Sheet */}
          <div 
            className={`relative rounded-2xl border-2 transition-all duration-300 overflow-hidden bg-slate-900/40 backdrop-blur-sm shadow-lg ${
              isDragExcel 
                ? "border-teal-400 bg-teal-500/5 shadow-teal-500/5" 
                : excelFile 
                  ? "border-emerald-500/40 hover:border-emerald-500/60" 
                  : "border-slate-800 hover:border-slate-700/80"
            }`}
            onDragOver={handleDragExcel}
            onDragLeave={handleDragExcel}
            onDrop={handleDropExcel}
          >
            {/* Top design line */}
            <div className={`h-1.5 w-full ${excelFile ? "bg-emerald-500" : "bg-teal-500"}`} />

            <div className="p-6 sm:p-8 flex flex-col items-center text-center">
              <div className={`h-14 w-14 rounded-2xl flex items-center justify-center mb-4 transition-all duration-300 ${
                excelFile 
                  ? "bg-emerald-500/10 border border-emerald-500/30 text-emerald-400" 
                  : "bg-teal-500/10 border border-teal-500/30 text-teal-400"
              }`}>
                <FileSpreadsheet className="h-7 w-7" />
              </div>

              {!excelFile ? (
                <>
                  <h3 className="text-base font-bold text-slate-100">
                    {lang === "si" ? "Excel ගොනුව මෙතැනට Drag & Drop කරන්න" : "Drag & Drop Excel File Here"}
                  </h3>
                  <p className="text-xs text-slate-400 max-w-sm mt-1.5">
                    {lang === "si" ? "හෝ පරිගණකයෙන් තෝරන්න (.xlsx, .xls, .csv)" : "or browse your local files (.xlsx, .xls, .csv)"}
                  </p>
                  
                  <label className="mt-5 px-5 py-2.5 bg-slate-800 hover:bg-slate-700 border border-slate-700 hover:border-slate-600 rounded-xl text-xs font-semibold text-slate-200 transition-all duration-200 cursor-pointer shadow-sm active:scale-95">
                    {lang === "si" ? "Excel ගොනුව තෝරන්න" : "Browse Excel File"}
                    <input 
                      type="file" 
                      accept=".xlsx,.xls,.csv" 
                      className="hidden" 
                      onChange={handleFileChangeExcel}
                    />
                  </label>
                </>
              ) : (
                <div className="w-full">
                  <div className="flex items-center justify-between gap-3 bg-slate-950/60 p-4 rounded-xl border border-slate-800/80 mb-5">
                    <div className="flex items-center gap-3 overflow-hidden">
                      <div className="bg-emerald-500/20 text-emerald-400 p-2 rounded-lg shrink-0">
                        <CheckCircle2 className="h-4 w-4" />
                      </div>
                      <div className="text-left overflow-hidden">
                        <p className="text-xs font-bold text-slate-200 truncate">{excelFile.name}</p>
                        <p className="text-[10px] text-slate-500 font-mono mt-0.5">{(excelFile.size / 1024).toFixed(1)} KB</p>
                      </div>
                    </div>
                    <button
                      onClick={handleResetExcel}
                      className="p-1.5 hover:bg-rose-500/10 text-slate-500 hover:text-rose-400 rounded-lg transition-all duration-200"
                      title="Remove file"
                    >
                      <Trash2 className="h-4 w-4" />
                    </button>
                  </div>

                  {/* Excel Details & Config */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-left">
                    
                    {/* Sheet Selection */}
                    {sheets.length > 1 && (
                      <div className="sm:col-span-2">
                        <label className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">
                          {lang === "si" ? "පත්‍රිකාව තෝරන්න (Select Sheet)" : "Select Sheet"}
                        </label>
                        <select
                          value={selectedSheet}
                          onChange={(e) => handleSheetChange(e.target.value)}
                          className="mt-1 w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs font-medium text-slate-200 outline-none focus:border-teal-500 transition-all duration-200"
                        >
                          {sheets.map((s) => (
                            <option key={s} value={s}>{s}</option>
                          ))}
                        </select>
                      </div>
                    )}

                    {/* NIC Column Mapped */}
                    <div>
                      <label className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">
                        {lang === "si" ? "NIC තීරුව (NIC Column)" : "NIC Column"}
                      </label>
                      <select
                        value={nicColumn}
                        onChange={(e) => setNicColumn(e.target.value)}
                        className="mt-1 w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs font-medium text-slate-200 outline-none focus:border-teal-500 transition-all duration-200"
                      >
                        <option value="" disabled>-- Select Column --</option>
                        {headers.map((h) => (
                          <option key={h} value={h}>{h}</option>
                        ))}
                      </select>
                      <p className="text-[10px] text-slate-500 mt-1">
                        {lang === "si" ? "PDF නාමය සමඟ සසඳන අගය" : "Value matching PDF filename"}
                      </p>
                    </div>

                    {/* Employee Code Column Mapped */}
                    <div>
                      <label className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">
                        {lang === "si" ? "සේවක අංක තීරුව (Employee No)" : "Employee No Column"}
                      </label>
                      <select
                        value={empColumn}
                        onChange={(e) => setEmpColumn(e.target.value)}
                        className="mt-1 w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs font-medium text-slate-200 outline-none focus:border-teal-500 transition-all duration-200"
                      >
                        <option value="" disabled>-- Select Column --</option>
                        {headers.map((h) => (
                          <option key={h} value={h}>{h}</option>
                        ))}
                      </select>
                      <p className="text-[10px] text-slate-500 mt-1">
                        {lang === "si" ? "නව PDF නාමය ලෙස යොදන අගය" : "Target filename value"}
                      </p>
                    </div>

                    {/* Name Column Mapped */}
                    <div>
                      <label className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">
                        {lang === "si" ? "නම තීරුව (Name Column)" : "Name Column"}
                      </label>
                      <select
                        value={nameColumn}
                        onChange={(e) => setNameColumn(e.target.value)}
                        className="mt-1 w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs font-medium text-slate-200 outline-none focus:border-teal-500 transition-all duration-200"
                      >
                        <option value="" disabled>-- Select Column --</option>
                        {headers.map((h) => (
                          <option key={h} value={h}>{h}</option>
                        ))}
                      </select>
                      <p className="text-[10px] text-slate-500 mt-1">
                        {lang === "si" ? "සේවකයාගේ නම සඳහන් තීරුව" : "Employee name column"}
                      </p>
                    </div>

                    {/* Phone Column 1 Mapped */}
                    <div>
                      <label className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">
                        {lang === "si" ? "ප්‍රධාන දුරකථන තීරුව (Primary Phone)" : "Primary Phone Column"}
                      </label>
                      <select
                        value={phoneColumn}
                        onChange={(e) => setPhoneColumn(e.target.value)}
                        className="mt-1 w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs font-medium text-slate-200 outline-none focus:border-teal-500 transition-all duration-200"
                      >
                        <option value="" disabled>-- Select Column --</option>
                        <option value="">-- None / නැත --</option>
                        {headers.map((h) => (
                          <option key={h} value={h}>{h}</option>
                        ))}
                      </select>
                      <p className="text-[10px] text-slate-500 mt-1">
                        {lang === "si" ? "ප්‍රධාන WhatsApp පණිවිඩය යවන අංකය" : "Primary WhatsApp recipient phone column"}
                      </p>
                    </div>

                    {/* Phone Column 2 Mapped */}
                    <div>
                      <label className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">
                        {lang === "si" ? "අතිරේක දුරකථන තීරුව (Secondary Phone)" : "Secondary Phone Column"}
                      </label>
                      <select
                        value={phoneColumn2}
                        onChange={(e) => setPhoneColumn2(e.target.value)}
                        className="mt-1 w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs font-medium text-slate-200 outline-none focus:border-teal-500 transition-all duration-200"
                      >
                        <option value="" disabled>-- Select Column --</option>
                        <option value="">-- None / නැත --</option>
                        {headers.map((h) => (
                          <option key={h} value={h}>{h}</option>
                        ))}
                      </select>
                      <p className="text-[10px] text-slate-500 mt-1">
                        {lang === "si" ? "අතිරේක WhatsApp පණිවිඩය යවන අංකය" : "Secondary WhatsApp recipient phone column"}
                      </p>
                    </div>
                  </div>

                  <div className="mt-4 p-3 bg-slate-950/40 rounded-xl border border-slate-850 text-left">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                      {lang === "si" ? "දත්ත පෙරදසුන (Excel Rows)" : "Data Preview (First 2 rows)"}
                    </p>
                    <div className="overflow-x-auto mt-1.5">
                      <table className="w-full text-left text-xs border-collapse">
                        <thead>
                          <tr className="border-b border-slate-850 text-slate-400 font-semibold">
                            <th className="pb-1 pr-3">Row</th>
                            <th className="pb-1 pr-3 truncate max-w-[100px]">{nicColumn || "NIC"}</th>
                            <th className="pb-1 pr-3 truncate max-w-[100px]">{empColumn || "Employee No"}</th>
                            <th className="pb-1 pr-3 truncate max-w-[100px]">{nameColumn || "Name"}</th>
                            <th className="pb-1 pr-3 truncate max-w-[100px]">{phoneColumn || "Phone 1"}</th>
                            <th className="pb-1 truncate max-w-[100px]">{phoneColumn2 || "Phone 2"}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {excelData.slice(0, 2).map((row, idx) => (
                            <tr key={idx} className="text-slate-300 font-mono">
                              <td className="py-1 pr-3">#{idx + 2}</td>
                              <td className="py-1 pr-3 truncate max-w-[100px]">{String(row[nicColumn] || "")}</td>
                              <td className="py-1 pr-3 truncate max-w-[100px]">{String(row[empColumn] || "")}</td>
                              <td className="py-1 pr-3 truncate max-w-[100px]">{String(row[nameColumn] || "")}</td>
                              <td className="py-1 pr-3 truncate max-w-[100px]">{String(row[phoneColumn] || "")}</td>
                              <td className="py-1 truncate max-w-[100px]">{String(row[phoneColumn2] || "")}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* File Input 2: ZIP Archive */}
          <div 
            className={`relative rounded-2xl border-2 transition-all duration-300 overflow-hidden bg-slate-900/40 backdrop-blur-sm shadow-lg ${
              isDragZip 
                ? "border-teal-400 bg-teal-500/5 shadow-teal-500/5" 
                : zipFile 
                  ? "border-emerald-500/40 hover:border-emerald-500/60" 
                  : "border-slate-800 hover:border-slate-700/80"
            }`}
            onDragOver={handleDragZip}
            onDragLeave={handleDragZip}
            onDrop={handleDropZip}
          >
            {/* Top design line */}
            <div className={`h-1.5 w-full ${zipFile ? "bg-emerald-500" : "bg-teal-500"}`} />

            <div className="p-6 sm:p-8 flex flex-col items-center text-center">
              <div className={`h-14 w-14 rounded-2xl flex items-center justify-center mb-4 transition-all duration-300 ${
                zipFile 
                  ? "bg-emerald-500/10 border border-emerald-500/30 text-emerald-400" 
                  : "bg-teal-500/10 border border-teal-500/30 text-teal-400"
              }`}>
                <FileArchive className="h-7 w-7" />
              </div>

              {!zipFile ? (
                <>
                  <h3 className="text-base font-bold text-slate-100">
                    {lang === "si" ? "PDF සහිත ZIP ගොනුව මෙතැනට Drag & Drop කරන්න" : "Drag & Drop ZIP File Here"}
                  </h3>
                  <p className="text-xs text-slate-400 max-w-sm mt-1.5">
                    {lang === "si" ? "හෝ පරිගණකයෙන් තෝරන්න (.zip)" : "or browse your local files (.zip)"}
                  </p>
                  
                  <label className="mt-5 px-5 py-2.5 bg-slate-800 hover:bg-slate-700 border border-slate-700 hover:border-slate-600 rounded-xl text-xs font-semibold text-slate-200 transition-all duration-200 cursor-pointer shadow-sm active:scale-95">
                    {lang === "si" ? "ZIP ගොනුව තෝරන්න" : "Browse ZIP File"}
                    <input 
                      type="file" 
                      accept=".zip" 
                      className="hidden" 
                      onChange={handleFileChangeZip}
                    />
                  </label>
                </>
              ) : (
                <div className="w-full h-full flex flex-col justify-between">
                  <div className="flex items-center justify-between gap-3 bg-slate-950/60 p-4 rounded-xl border border-slate-800/80 mb-5">
                    <div className="flex items-center gap-3 overflow-hidden">
                      <div className="bg-emerald-500/20 text-emerald-400 p-2 rounded-lg shrink-0">
                        <CheckCircle2 className="h-4 w-4" />
                      </div>
                      <div className="text-left overflow-hidden">
                        <p className="text-xs font-bold text-slate-200 truncate">{zipFile.name}</p>
                        <p className="text-[10px] text-slate-500 font-mono mt-0.5">{(zipFile.size / (1024 * 1024)).toFixed(2)} MB</p>
                      </div>
                    </div>
                    <button
                      onClick={handleResetZip}
                      className="p-1.5 hover:bg-rose-500/10 text-slate-500 hover:text-rose-400 rounded-lg transition-all duration-200"
                      title="Remove file"
                    >
                      <Trash2 className="h-4 w-4" />
                    </button>
                  </div>

                  <div className="bg-slate-950/40 p-4 rounded-xl border border-slate-850 text-left mb-4">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                      {lang === "si" ? "ZIP ගොනුවේ අන්තර්ගතය" : "ZIP Archive Details"}
                    </p>
                    <div className="mt-2 grid grid-cols-2 gap-4">
                      <div>
                        <span className="text-xs text-slate-400">
                          {lang === "si" ? "මුළු PDF ගොනු ගණන" : "Total PDF files"}
                        </span>
                        <p className="text-xl font-bold font-mono text-teal-400 mt-0.5">
                          {zipFileList.length}
                        </p>
                      </div>
                      <div>
                        <span className="text-xs text-slate-400">
                          {lang === "si" ? "ගොනු ආකෘතිය" : "Detected Format"}
                        </span>
                        <p className="text-xs font-bold text-slate-200 mt-1 flex items-center gap-1.5">
                          <span className="h-2 w-2 rounded-full bg-emerald-400" />
                          Adobe PDF (.pdf)
                        </p>
                      </div>
                    </div>
                  </div>

                  <div className="p-3 bg-slate-950/30 rounded-xl border border-slate-850/50 text-left">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                      {lang === "si" ? "හඳුනාගත් ZIP ගොනු නාමයන් (පළමු 2)" : "Extracted ZIP Names (First 2)"}
                    </p>
                    <ul className="mt-1.5 space-y-1">
                      {zipFileList.slice(0, 2).map((item, idx) => (
                        <li key={idx} className="text-xs text-slate-300 font-mono flex items-center gap-1.5 truncate">
                          <FileText className="h-3.5 w-3.5 text-slate-500 shrink-0" />
                          <span className="truncate">{item.name}</span>
                        </li>
                      ))}
                    </ul>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Global Progress Overlay if processing */}
        {isProcessing && (
          <div className="bg-slate-900 border border-teal-500/30 rounded-2xl p-6 shadow-xl shadow-teal-950/20 relative overflow-hidden transition-all duration-300">
            <div className="absolute top-0 left-0 h-1 bg-teal-500/20 w-full" />
            <div className="flex flex-col gap-4">
              <div className="flex items-center justify-between gap-4">
                <div className="flex items-center gap-2.5">
                  <RefreshCw className="h-4.5 w-4.5 text-teal-400 animate-spin" />
                  <span className="text-sm font-semibold text-slate-200">
                    {lang === "si" ? "ක්‍රියාවලිය සිදුවෙමින් පවතී..." : "Processing File Operations..."}
                  </span>
                </div>
                <span className="text-sm font-bold text-teal-400 font-mono">{progress}%</span>
              </div>
              
              {/* Progress Bar Container */}
              <div className="w-full bg-slate-950 h-2.5 rounded-full overflow-hidden border border-slate-800">
                <div 
                  className="bg-gradient-to-r from-teal-500 to-emerald-400 h-full rounded-full transition-all duration-300 shadow-md"
                  style={{ width: `${progress}%` }}
                />
              </div>
              
              <p className="text-xs text-slate-400 italic truncate font-mono">
                {progressText}
              </p>
            </div>
          </div>
        )}

        {/* Matching Analysis and Actions Panel */}
        {excelFile && zipFile && (
          <div className="flex flex-col gap-6">
            
            {/* Quick Statistics Grid */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              
              {/* Stat 1: Matched */}
              <div className="bg-slate-900/60 border border-slate-800/80 p-4 rounded-xl text-left backdrop-blur-sm">
                <span className="text-[10px] uppercase font-bold text-slate-400 tracking-wider block">
                  {lang === "si" ? "පූර්ණ ලෙස ගැලපෙන" : "Fully Matched"}
                </span>
                <div className="flex items-baseline gap-2 mt-1">
                  <span className="text-2xl font-bold font-mono text-emerald-400">
                    {matchingData.stats.matched}
                  </span>
                  <span className="text-xs text-slate-500">
                    / {zipFileList.length}
                  </span>
                </div>
                <p className="text-[10px] text-slate-400 mt-1 truncate">
                  {lang === "si" ? "සේවක අංකයට වෙනස් වේ" : "Will be renamed successfully"}
                </p>
              </div>

              {/* Stat 2: Unmatched inside ZIP */}
              <div className="bg-slate-900/60 border border-slate-800/80 p-4 rounded-xl text-left backdrop-blur-sm">
                <span className="text-[10px] uppercase font-bold text-slate-400 tracking-wider block">
                  {lang === "si" ? "Excel හි නොමැති PDF" : "ZIP PDFs without Excel row"}
                </span>
                <div className="flex items-baseline gap-2 mt-1">
                  <span className="text-2xl font-bold font-mono text-amber-400">
                    {matchingData.matchedList.filter(m => m.status === "unmatched_zip").length}
                  </span>
                  <span className="text-xs text-slate-500">
                    / {zipFileList.length}
                  </span>
                </div>
                <p className="text-[10px] text-slate-400 mt-1 truncate">
                  {lang === "si" ? "NIC එක Excel ලැයිස්තුවේ නැත" : "NIC not present in sheet"}
                </p>
              </div>

              {/* Stat 3: Unmatched Excel rows */}
              <div className="bg-slate-900/60 border border-slate-800/80 p-4 rounded-xl text-left backdrop-blur-sm">
                <span className="text-[10px] uppercase font-bold text-slate-400 tracking-wider block">
                  {lang === "si" ? "ගොනු නොමැති සේවකයින්" : "Excel Employees without PDF"}
                </span>
                <div className="flex items-baseline gap-2 mt-1">
                  <span className="text-2xl font-bold font-mono text-cyan-400">
                    {matchingData.stats.unmatchedExcel}
                  </span>
                  <span className="text-xs text-slate-500">
                    / {excelData.length}
                  </span>
                </div>
                <p className="text-[10px] text-slate-400 mt-1 truncate">
                  {lang === "si" ? "ගැලපෙන PDF එකක් ZIP එකේ නැත" : "No matching PDF in archive"}
                </p>
              </div>

              {/* Stat 4: Duplicates & Warnings */}
              <div className="bg-slate-900/60 border border-slate-800/80 p-4 rounded-xl text-left backdrop-blur-sm">
                <span className="text-[10px] uppercase font-bold text-slate-400 tracking-wider block">
                  {lang === "si" ? "අනතුරු ඇඟවීම්" : "Duplicate Warnings"}
                </span>
                <div className="flex items-baseline gap-2 mt-1">
                  <span className={`text-2xl font-bold font-mono ${
                    matchingData.matchedList.filter(m => m.status === "duplicate_nic" || m.status === "duplicate_emp").length > 0 
                      ? "text-rose-400" 
                      : "text-slate-400"
                  }`}>
                    {matchingData.matchedList.filter(m => m.status === "duplicate_nic" || m.status === "duplicate_emp").length}
                  </span>
                  <span className="text-xs text-slate-500">
                    {lang === "si" ? "දෝෂ" : "conflicts"}
                  </span>
                </div>
                <p className="text-[10px] text-slate-400 mt-1 truncate">
                  {lang === "si" ? "සැක සහිත හෝ අනුපිටපත් ගොනු" : "Duplicate NIC/Employee ID checks"}
                </p>
              </div>

            </div>

            {/* Options and Action Button Card */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-xl relative overflow-hidden">
              <div className="absolute right-0 bottom-0 translate-x-10 translate-y-10 w-32 h-32 bg-emerald-500/5 rounded-full blur-2xl pointer-events-none" />
              
              <h3 className="text-base font-bold text-slate-100 flex items-center gap-2">
                <Database className="h-5 w-5 text-teal-400" />
                {lang === "si" ? "පරිවර්තන සැකසුම් සහ බාගත කිරීම්" : "Relabeling Settings & Output Archive"}
              </h3>
              
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-6 mt-5 items-end">
                
                {/* Setting 1: Custom Suffix */}
                <div>
                  <label className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">
                    {lang === "si" ? "ගොනු නාමයේ අගට එකතු කිරීම්" : "Filename Suffix (Optional)"}
                  </label>
                  <input
                    type="text"
                    value={filenameSuffix}
                    onChange={(e) => setFilenameSuffix(e.target.value)}
                    placeholder="e.g. _slip, _2026"
                    className="mt-1.5 w-full bg-slate-950 border border-slate-800 hover:border-slate-700/80 rounded-xl px-3.5 py-2.5 text-xs text-slate-200 font-mono outline-none focus:border-teal-500 transition-all duration-200"
                  />
                  <span className="text-[9px] text-slate-500 mt-1 block">
                    {lang === "si" ? "උදා: EMP-001 (951234567V)_slip.pdf" : "Result: EMP-001 (951234567V)_slip.pdf"}
                  </span>
                </div>

                {/* Setting 2: Pass Month (Auto Detected / Editable) */}
                <div>
                  <div className="flex items-center gap-1.5 justify-between">
                    <label className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">
                      {lang === "si" ? "බලපත්‍ර මාසය" : "Pass Month (Auto)"}
                    </label>
                    <span className="text-[9px] text-teal-400 font-bold bg-teal-500/10 px-1 py-0.2 rounded animate-pulse shrink-0">
                      Auto
                    </span>
                  </div>
                  <select
                    value={passMonth}
                    onChange={(e) => setPassMonth(e.target.value)}
                    className="mt-1.5 w-full bg-slate-950 border border-slate-800 hover:border-slate-700/80 rounded-xl px-3.5 py-2.5 text-xs text-slate-200 outline-none focus:border-teal-500 transition-all duration-200 cursor-pointer"
                  >
                    {[
                      "January", "February", "March", "April", "May", "June", 
                      "July", "August", "September", "October", "November", "December"
                    ].map((month) => (
                      <option key={month} value={month}>{month}</option>
                    ))}
                  </select>
                </div>

                {/* Setting 3: Unmatched Action toggle */}
                <div className="flex flex-col gap-2">
                  <span className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">
                    {lang === "si" ? "නොගැලපෙන ZIP PDF ගොනු" : "ZIP PDFs without match"}
                  </span>
                  
                  <label className="flex items-center gap-2.5 bg-slate-950 border border-slate-800 hover:border-slate-700/80 rounded-xl px-3.5 py-2.5 text-xs text-slate-300 font-medium cursor-pointer transition-all duration-200">
                    <input
                      type="checkbox"
                      checked={includeUnmatched}
                      onChange={(e) => setIncludeUnmatched(e.target.checked)}
                      className="rounded bg-slate-900 border-slate-800 text-teal-500 focus:ring-teal-500 h-4 w-4 shrink-0"
                    />
                    <span>
                      {lang === "si" ? "පෙර නමින්ම ZIP එකට එක්කරන්න" : "Keep in ZIP with old name"}
                    </span>
                  </label>
                </div>

                {/* Setting 4: Output ZIP File name */}
                <div>
                  <label className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">
                    {lang === "si" ? "ප්‍රතිදාන ZIP නම" : "Output ZIP Filename"}
                  </label>
                  <input
                    type="text"
                    value={outputZipName}
                    onChange={(e) => setOutputZipName(e.target.value)}
                    placeholder="renamed_pdfs.zip"
                    className="mt-1.5 w-full bg-slate-950 border border-slate-800 hover:border-slate-700/80 rounded-xl px-3.5 py-2.5 text-xs text-slate-200 font-mono outline-none focus:border-teal-500 transition-all duration-200"
                  />
                </div>

              </div>

              {/* Main Action Button */}
              <div className="mt-6 pt-6 border-t border-slate-800 flex flex-col sm:flex-row items-center justify-between gap-4">
                
                <div className="text-left">
                  <span className="text-xs text-slate-400 flex items-center gap-1.5">
                    <ShieldCheck className="h-4 w-4 text-emerald-400 shrink-0" />
                    {lang === "si" 
                      ? "සැකසීම සම්පූර්ණයෙන්ම බ්‍රවුසරය තුල සිදුවේ. කිසිදු ගොනුවක් අපගේ සර්වර් වෙත යොමු නොවේ." 
                      : "Processed 100% locally. Your private files never leave your computer."
                    }
                  </span>
                </div>

                <button
                  onClick={handleRenameAndDownload}
                  disabled={isProcessing || zipFileList.length === 0}
                  className="w-full sm:w-auto px-8 py-3.5 rounded-xl bg-gradient-to-r from-teal-500 to-emerald-400 hover:from-teal-400 hover:to-emerald-300 text-slate-950 font-bold text-sm tracking-wide shadow-lg shadow-teal-500/10 hover:shadow-teal-400/20 flex items-center justify-center gap-2.5 transition-all duration-300 disabled:opacity-40 disabled:cursor-not-allowed hover:-translate-y-0.5 active:translate-y-0 active:scale-98"
                >
                  <Download className="h-4.5 w-4.5" />
                  {lang === "si" 
                    ? `පරිවර්තනය කර බාගන්න (ගොනු ${includeUnmatched ? zipFileList.length : matchingData.stats.matched} ක්)` 
                    : `Rename & Download ZIP (${includeUnmatched ? zipFileList.length : matchingData.stats.matched} files)`
                  }
                </button>
              </div>
            </div>

            {/* Renaming Preview & List Panel */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-xl">
              
              {/* Header inside Panel */}
              <div className="px-6 py-5 border-b border-slate-800 bg-slate-900/80 flex flex-col md:flex-row md:items-center justify-between gap-4">
                
                <div>
                  <h3 className="text-sm font-bold text-slate-100 uppercase tracking-wider">
                    {lang === "si" ? "පරිවර්තන පෙරදසුන (Rename Preview Table)" : "Relabeling Mapping Results"}
                  </h3>
                  <p className="text-xs text-slate-400 mt-1">
                    {lang === "si" ? "ZIP ගොනු සැබෑ ලෙස වෙනස් වන ආකාරය" : "Compare original filenames with targets"}
                  </p>
                </div>

                {/* Filter and search controls */}
                <div className="flex flex-wrap items-center gap-3">
                  
                  {/* Search Field Filter Selector */}
                  <select
                    value={filterType}
                    onChange={(e) => setFilterType(e.target.value as any)}
                    className="bg-slate-950 border border-slate-800 hover:border-slate-700/80 rounded-xl px-3 py-2.5 text-xs text-slate-300 font-medium outline-none focus:border-teal-500 transition-all duration-200 cursor-pointer"
                  >
                    <option value="all">{lang === "si" ? "සියල්ල (All)" : "All Fields"}</option>
                    <option value="nic">{lang === "si" ? "NIC අංකයෙන් පමණි" : "NIC Only"}</option>
                    <option value="emp">{lang === "si" ? "සේවක අංකයෙන් (EMP)" : "Employee ID Only"}</option>
                    <option value="name">{lang === "si" ? "නමින් පමණි" : "Name Only"}</option>
                  </select>

                  {/* Reset Sent Counts Button */}
                  {Object.keys(whatsappSentCounts).length > 0 && (
                    <button
                      onClick={handleResetSentCounts}
                      className="text-xs px-3.5 py-2 rounded-xl border border-rose-500/30 bg-rose-500/10 text-rose-400 hover:bg-rose-500 hover:text-slate-950 transition-all duration-200 font-bold active:scale-95 cursor-pointer flex items-center gap-1.5"
                      title={lang === "si" ? "සියලුම WhatsApp යැවීම් බිංදු කරන්න" : "Reset all WhatsApp sent counts"}
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                      <span>
                        {lang === "si" ? "ගණන් බිංදු කරන්න" : "Reset Counts"}
                      </span>
                    </button>
                  )}

                  {/* Search Bar */}
                  <div className="relative max-w-xs w-full">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-slate-500" />
                    <input
                      type="text"
                      placeholder={
                        filterType === "nic" 
                          ? (lang === "si" ? "NIC අංකය සොයන්න..." : "Search by NIC...")
                          : filterType === "emp"
                            ? (lang === "si" ? "සේවක අංකය සොයන්න..." : "Search by Employee ID...")
                            : filterType === "name"
                              ? (lang === "si" ? "නම සොයන්න..." : "Search by Name...")
                              : (lang === "si" ? "සොයන්න (NIC/EMP/නම)..." : "Search files...")
                      }
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="bg-slate-950 border border-slate-800 hover:border-slate-700/80 rounded-xl pl-9 pr-4 py-2 text-xs text-slate-200 outline-none focus:border-teal-500 transition-all duration-200"
                    />
                  </div>

                  {/* Filter selector */}
                  <div className="flex bg-slate-950 p-1 rounded-xl border border-slate-800">
                    <button
                      onClick={() => setStatusFilter("all")}
                      className={`text-xs px-3 py-1.5 rounded-lg transition-all duration-150 font-medium ${
                        statusFilter === "all" ? "bg-slate-800 text-teal-400" : "text-slate-400 hover:text-slate-300"
                      }`}
                    >
                      {lang === "si" ? "සියල්ල" : "All"} ({matchingData.matchedList.length})
                    </button>
                    <button
                      onClick={() => setStatusFilter("matched")}
                      className={`text-xs px-3 py-1.5 rounded-lg transition-all duration-150 font-medium ${
                        statusFilter === "matched" ? "bg-slate-800 text-emerald-400" : "text-slate-400 hover:text-slate-300"
                      }`}
                    >
                      {lang === "si" ? "ගැලපෙන" : "Matched"} ({matchingData.stats.matched})
                    </button>
                    <button
                      onClick={() => setStatusFilter("unmatched_zip")}
                      className={`text-xs px-3 py-1.5 rounded-lg transition-all duration-150 font-medium ${
                        statusFilter === "unmatched_zip" ? "bg-slate-800 text-amber-400" : "text-slate-400 hover:text-slate-300"
                      }`}
                    >
                      {lang === "si" ? "නොගැලපෙන" : "Unmatched"} ({matchingData.stats.unmatchedZip})
                    </button>
                  </div>

                </div>

              </div>

              {/* Table */}
              <div className="overflow-auto max-h-[950px] border border-slate-800/80 rounded-xl relative scrollbar-thin scrollbar-thumb-slate-850 scrollbar-track-transparent">
                {filteredList.length === 0 ? (
                  <div className="p-8 text-center flex flex-col items-center">
                    <AlertCircle className="h-8 w-8 text-slate-600 mb-2" />
                    <span className="text-sm text-slate-400 font-medium">
                      {lang === "si" ? "කිසිදු ප්‍රතිඵලයක් සොයාගත නොහැක" : "No results match the current filters"}
                    </span>
                  </div>
                ) : (
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="sticky top-0 bg-slate-900 z-20 shadow-[0_1px_0_0_rgba(255,255,255,0.05)]">
                      <tr className="text-slate-400 font-bold border-b border-slate-800 bg-slate-900">
                        <th className="py-4 px-6">{lang === "si" ? "මුල් PDF (ZIP)" : "Original PDF (ZIP)"}</th>
                        <th className="py-4 px-4">{lang === "si" ? "NIC" : "NIC"}</th>
                        <th className="py-4 px-4">{lang === "si" ? "නම" : "Name"}</th>
                        <th className="py-4 px-4">{lang === "si" ? "දුරකථන අංක" : "Phone Numbers"}</th>
                        <th className="py-4 px-4">{lang === "si" ? "නව ගොනු නම" : "Target Filename"}</th>
                        <th className="py-4 px-4 text-center">{lang === "si" ? "තත්ත්වය" : "Status"}</th>
                        <th className="py-4 px-4 text-center">{lang === "si" ? "යැවූ තත්ත්වය" : "Sent Status"}</th>
                        <th className="py-4 px-6 text-right">{lang === "si" ? "ක්‍රියාවන් (WhatsApp)" : "Actions (WhatsApp)"}</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-850/40">
                      {filteredList.map((item, idx) => (
                        <tr key={idx} className="hover:bg-slate-900/50 hover:backdrop-blur-sm border-b border-slate-850/40 last:border-0 transition-all duration-200">
                          
                          {/* Original PDF Name */}
                          <td className="py-4 px-6 font-mono text-slate-300 truncate max-w-[260px] hover:text-teal-400 transition-colors duration-150" title={item.originalName}>
                            {item.originalName}
                          </td>

                          {/* Extracted NIC */}
                          <td className="py-4 px-4 font-mono">
                            {item.nicKey ? (
                              <span className="text-slate-300 bg-slate-950 px-2.5 py-1 rounded-lg border border-slate-800/80 shadow-sm font-medium">
                                {item.nicKey}
                              </span>
                            ) : (
                              <span className="text-slate-500 italic font-light">None</span>
                            )}
                          </td>

                          {/* Name */}
                          <td className="py-4 px-4 text-slate-200 font-semibold truncate max-w-[220px]" title={item.name || ""}>
                            {item.name || <span className="text-slate-500 italic font-normal">-</span>}
                          </td>

                          {/* Phone */}
                          <td className="py-4 px-4 font-mono">
                            <div className="flex flex-col gap-1.5 min-w-[160px] max-w-[240px]">
                              {item.phones && item.phones.length > 0 ? (
                                item.phones.map((phoneNumber, index) => {
                                  const validation = validatePhoneForWhatsApp(phoneNumber);
                                  return (
                                    <div key={index} className="flex items-center justify-between gap-2 bg-slate-950/40 px-2 py-1.5 rounded border border-slate-800/80 transition-all duration-200">
                                      <div className="flex items-center gap-1.5 truncate">
                                        <span className={`text-[9px] font-bold px-1 py-0.5 rounded shrink-0 ${
                                          index === 0 ? "bg-teal-500/15 text-teal-400 border border-teal-500/10" : "bg-cyan-500/15 text-cyan-400 border border-cyan-500/10"
                                        }`}>
                                          P{index + 1}
                                        </span>
                                        <span className="text-xs text-slate-200 select-all truncate">{phoneNumber}</span>
                                      </div>
                                      
                                      {/* Blinking Status Indicator Dot */}
                                      <div className="flex items-center gap-1.5 shrink-0" title={validation.message}>
                                        <span className="relative flex h-2.5 w-2.5">
                                          {validation.type === "mobile" ? (
                                            <>
                                              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                                              <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.8)]"></span>
                                            </>
                                          ) : validation.type === "landline" ? (
                                            <>
                                              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-rose-400 opacity-75"></span>
                                              <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-rose-500 shadow-[0_0_8px_rgba(244,63,94,0.8)]"></span>
                                            </>
                                          ) : (
                                            <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-slate-600"></span>
                                          )}
                                        </span>
                                      </div>
                                    </div>
                                  );
                                })
                              ) : (
                                <span className="text-slate-500 italic text-xs">-</span>
                              )}
                            </div>
                          </td>

                          {/* Target Renamed Name */}
                          <td className="py-4 px-4 font-mono text-sm max-w-[280px] break-words">
                            {item.status === "matched" && item.empNo ? (
                              <span className="text-emerald-400 font-bold bg-emerald-500/5 border border-emerald-500/10 px-2 py-1.5 rounded-lg shadow-sm block text-center truncate" title={`${item.empNo} (${item.nicKey.toUpperCase()})${filenameSuffix}.pdf`}>
                                {item.empNo} ({item.nicKey.toUpperCase()}){filenameSuffix}.pdf
                              </span>
                            ) : includeUnmatched && item.status !== "matched" ? (
                              <span className="text-slate-400 italic bg-slate-950/40 px-2 py-1.5 rounded-lg border border-slate-900 block text-center truncate" title={item.originalName}>
                                {item.originalName}
                              </span>
                            ) : (
                              <span className="text-rose-400/60 line-through bg-rose-500/5 px-2 py-1.5 rounded-lg border border-rose-500/10 block text-center">
                                (Excluded)
                              </span>
                            )}
                          </td>

                          {/* Status Badge */}
                          <td className="py-4 px-4">
                            <div className="flex items-center justify-center">
                              {item.status === "matched" ? (
                                item.matchType === "exact" ? (
                                  <span 
                                    className="bg-emerald-500/10 text-emerald-400 px-2.5 py-1 rounded-full border border-emerald-500/20 font-semibold text-[10px] flex items-center gap-1.5 cursor-help"
                                    title={item.details}
                                  >
                                    <span className="h-1.5 w-1.5 rounded-full bg-emerald-400" />
                                    {lang === "si" ? "100% ගැලපේ" : "Exact Match"}
                                  </span>
                                ) : item.matchType === "substring" ? (
                                  <span 
                                    className="bg-teal-500/10 text-teal-400 px-2.5 py-1 rounded-full border border-teal-500/20 font-semibold text-[10px] flex items-center gap-1.5 cursor-help"
                                    title={item.details}
                                  >
                                    <span className="h-1.5 w-1.5 rounded-full bg-teal-400" />
                                    {lang === "si" ? "කොටසක් ගැලපේ" : "Partial Match"}
                                  </span>
                                ) : (
                                  <span 
                                    className="bg-cyan-500/10 text-cyan-400 px-2.5 py-1 rounded-full border border-cyan-500/20 font-semibold text-[10px] flex items-center gap-1.5 cursor-help"
                                    title={item.details}
                                  >
                                    <span className="h-1.5 w-1.5 rounded-full bg-cyan-400" />
                                    {lang === "si" ? `සැක සහිත (${item.matchScore})` : `Fuzzy Match (${item.matchScore})`}
                                  </span>
                                )
                              ) : item.status === "unmatched_zip" ? (
                                <span 
                                  className="bg-slate-500/10 text-slate-400 px-2.5 py-1 rounded-full border border-slate-500/20 font-semibold text-[10px] flex items-center gap-1.5 cursor-help"
                                  title={item.details}
                                >
                                  <span className="h-1.5 w-1.5 rounded-full bg-slate-500" />
                                  {lang === "si" ? "නොගැලපේ (මුල් නම)" : "No match (Kept)"}
                                </span>
                              ) : (
                                <span 
                                  className="bg-rose-500/10 text-rose-400 px-2.5 py-1 rounded-full border border-rose-500/20 font-semibold text-[10px] flex items-center gap-1.5 cursor-help"
                                  title={item.details}
                                >
                                  <span className="h-1.5 w-1.5 rounded-full bg-rose-400" />
                                  {lang === "si" ? "අනුපිටපතක්" : "Duplicate"}
                                </span>
                              )}
                            </div>
                          </td>

                          {/* Sent Status */}
                          <td className="py-4 px-4 text-center">
                            <div className="flex items-center justify-center">
                              {whatsappSentCounts[item.originalPath] ? (
                                <div className="flex items-center gap-1.5 bg-emerald-500/10 text-emerald-400 px-2.5 py-1 rounded-full border border-emerald-500/20 font-semibold text-[10px] shadow-[0_0_8px_rgba(16,185,129,0.1)] animate-pulse">
                                  <CheckCircle2 className="h-3.5 w-3.5 text-emerald-400 shrink-0" />
                                  <span>
                                    {lang === "si" 
                                      ? `යවා ඇත (${whatsappSentCounts[item.originalPath]})` 
                                      : `Sent (${whatsappSentCounts[item.originalPath]})`
                                    }
                                  </span>
                                </div>
                              ) : (
                                <span className="text-slate-500 text-xs italic font-light">
                                  {lang === "si" ? "යවා නැත" : "Not Sent"}
                                </span>
                              )}
                            </div>
                          </td>

                          {/* WhatsApp Action Buttons */}
                          <td className="py-4 px-6 text-right">
                            <div className="flex flex-col gap-1.5 items-end justify-center">
                              {item.status === "matched" && item.phones && item.phones.length > 0 ? (
                                item.phones.map((phoneNumber, index) => {
                                  const validation = validatePhoneForWhatsApp(phoneNumber);
                                  return (
                                    <div key={index} className="flex items-center gap-1.5 bg-slate-950/30 p-1.5 rounded-xl border border-slate-850 hover:border-slate-800 transition-all duration-200">
                                      {/* Phone Label Badge */}
                                      <span className={`text-[9px] font-bold px-1.5 py-0.5 rounded shrink-0 ${
                                        index === 0 ? "bg-teal-500/10 text-teal-400" : "bg-cyan-500/10 text-cyan-400"
                                      }`}>
                                        #{index + 1}
                                      </span>
                                      
                                      {/* Number (hidden on tiny screens, shown on desktop) */}
                                      <span className="text-[10px] text-slate-400 font-mono hidden sm:inline truncate max-w-[85px]">
                                        {phoneNumber}
                                      </span>

                                      {/* Separator line */}
                                      <span className="h-3 w-px bg-slate-800 hidden sm:inline" />

                                      {/* Direct Verification Check Button */}
                                      <button
                                        onClick={() => handleVerifyWhatsApp(phoneNumber)}
                                        className="p-1 rounded-lg bg-slate-900/60 hover:bg-teal-500/10 hover:text-teal-400 text-slate-400 border border-slate-800/80 transition-all duration-200 active:scale-95 flex items-center justify-center cursor-pointer shrink-0"
                                        title={lang === "si" ? `WhatsApp අංකයදැයි පරීක්ෂා කරන්න (#${index + 1})` : `Verify WhatsApp Active status (#${index + 1})`}
                                      >
                                        <Search className="h-3 w-3" />
                                      </button>

                                      {/* Direct Send PDF Button */}
                                      <button
                                        onClick={() => handleSendWhatsApp(item, phoneNumber)}
                                        disabled={validation.type === "landline"}
                                        className={`inline-flex items-center justify-center p-1 rounded-lg border transition-all duration-300 active:scale-95 cursor-pointer shrink-0 ${
                                          validation.type === "landline"
                                            ? "bg-slate-900/40 border-slate-850 text-slate-600 cursor-not-allowed opacity-40"
                                            : "bg-emerald-500/10 hover:bg-emerald-500 text-emerald-400 hover:text-slate-950 border-emerald-500/20 hover:border-emerald-400 hover:shadow-emerald-500/20"
                                        }`}
                                        title={lang === "si" ? `WhatsApp හරහා PDF යවන්න (#${index + 1})` : `Send PDF via WhatsApp (#${index + 1})`}
                                      >
                                        <svg className="h-3.5 w-3.5 fill-current" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                                          <path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L0 24l6.335-1.662c1.746.953 3.71 1.458 5.704 1.459h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z" />
                                        </svg>
                                      </button>
                                    </div>
                                  );
                                })
                              ) : (
                                <span className="text-[10px] text-slate-500 italic font-light">
                                  {lang === "si" ? "අංකයන් නොමැත" : "No Phone Numbers"}
                                </span>
                              )}
                            </div>
                          </td>

                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>

              {/* Extra Unmatched Excel rows section if any */}
              {matchingData.unmatchedExcelList.length > 0 && (
                <div className="bg-slate-950/60 p-6 border-t border-slate-800">
                  <div className="flex items-center gap-2 text-cyan-400 mb-3">
                    <AlertTriangle className="h-4 w-4" />
                    <span className="text-xs font-bold uppercase tracking-wider">
                      {lang === "si" 
                        ? `ZIP එක තුල PDF ගොනු නොමැති සේවකයින් (${matchingData.unmatchedExcelList.length})` 
                        : `Employees listed in Excel but missing in ZIP (${matchingData.unmatchedExcelList.length})`
                      }
                    </span>
                  </div>
                  <div className="max-h-36 overflow-y-auto border border-slate-850 rounded-xl divide-y divide-slate-850">
                    {matchingData.unmatchedExcelList.map((rowItem, idx) => (
                      <div key={idx} className="p-2.5 flex items-center justify-between text-[11px] font-mono hover:bg-slate-900/40">
                        <span className="text-slate-400">
                          {lang === "si" ? "පේළිය" : "Row"} #{rowItem.rowNumber}
                        </span>
                        <div className="flex items-center gap-4">
                          <span className="text-slate-300">
                            NIC: <strong className="text-slate-100">{rowItem.nic}</strong>
                          </span>
                          <span className="text-slate-500">|</span>
                          <span className="text-cyan-400 font-bold">
                            EMP Code: {rowItem.empNo}
                          </span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

            </div>

          </div>
        )}

      </main>

      {/* Premium Footer */}
      <footer className="mt-auto border-t border-slate-800/80 bg-slate-950 py-8 relative z-10">
        <div className="max-w-[95%] 2xl:max-w-[1650px] mx-auto px-4 sm:px-6 lg:px-8 flex flex-col md:flex-row items-center justify-between gap-4 text-center md:text-left">
          <div>
            <p className="text-xs text-slate-400 font-light">
              &copy; 2026 ZIP Rename Nexus. Built for lightning-fast locally processed PDF relabeling.
            </p>
            <p className="text-[10px] text-slate-600 mt-1">
              Developed using Next.js 16, SheetJS, JSZip & Tailwind CSS v4. No files are uploaded to any server.
            </p>
          </div>
          <div className="flex items-center gap-6">
            <div className="flex items-center gap-1.5 text-xs text-slate-400">
              <span className="h-2 w-2 rounded-full bg-emerald-400 inline-block animate-ping" />
              <span>{lang === "si" ? "දේශීයව ක්‍රියාත්මකයි" : "Running locally"}</span>
            </div>
            <a 
              href="https://github.com" 
              target="_blank" 
              rel="noopener noreferrer" 
              className="text-xs text-slate-500 hover:text-slate-300 transition-colors"
            >
              Documentation
            </a>
          </div>
        </div>
      </footer>

    </div>
  );
}
