
import React, { useState, useEffect, useRef } from 'react';
import { Html5Qrcode } from 'html5-qrcode';
import * as XLSX from 'xlsx';
import QRCode from 'qrcode';
import { jsPDF } from 'jspdf';

// --- TYPE DEFINITIONS ---
interface StudentData {
  id: string; // Full text/name acts as ID
  name: string;
  school: string;
}

interface ScanRecord {
  studentId: string;
  timestamp: number;
  dateStr: string; // YYYY-MM-DD
}

// --- CONSTANTS ---
const STORAGE_KEY_DATA = 'qr_app_data_v6'; 
const STORAGE_KEY_HISTORY = 'qr_app_history_v6';

// List of Schools
const FIXED_SCHOOLS = [
    "2ο ΓΥΜΝΆΣΙΟ ΖΑΚΎΝΘΟΥ",
    "3ο ΓΥΜΝΆΣΙΟ ΖΑΚΎΝΘΟΥ",
    "ΓΥΜΝΆΣΙΟ ΛΤ ΒΑΝΆΤΟ",
    "ΓΥΜΝΆΣΙΟ ΛΤ ΒΟΛΙΜΏΝ",
    "ΓΥΜΝΆΣΙΟ ΛΙΘΑΚΙΆΣ",
    "ΓΥΜΝΆΣΙΟ ΚΑΤΑΣΤΑΡΊΟΥ",
    "ΓΥΜΝΆΣΙΟ ΜΑΧΑΙΡΆΔΟΥ",
    "ΓΥΜΝΆΣΙΟ ΜΟΥΣΙΚΌ",
    "1Ο ΓΕΛ ΖΑΚΎΝΘΟΥ",
    "2Ο ΓΕΛ ΖΑΚΎΝΘΟΥ",
    "ΓΕΛ ΚΑΤΑΣΤΑΡΊΟΥ",
    "ΕΠΑΛ ΖΑΚΎΝΘΟΥ"
];

// --- MATCHING LOGIC ---

// Helper: Normalize string for comparison (remove accents, uppercase, remove spaces/symbols)
const cleanStr = (str: string) => {
    return str.normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "") // remove accents
        .toUpperCase()
        .replace(/[^A-ZΑ-Ω0-9]/g, ""); // keep only letters and numbers
};

// Keywords mapping for robust detection
const SCHOOL_KEYWORDS: { [key: string]: string } = {
    "2ΟΓΥΜΝ": "2ο ΓΥΜΝΆΣΙΟ ΖΑΚΎΝΘΟΥ",
    "3ΟΓΥΜΝ": "3ο ΓΥΜΝΆΣΙΟ ΖΑΚΎΝΘΟΥ",
    "ΒΑΝΑΤ": "ΓΥΜΝΆΣΙΟ ΛΤ ΒΑΝΆΤΟ",
    "ΒΟΛΙΜ": "ΓΥΜΝΆΣΙΟ ΛΤ ΒΟΛΙΜΏΝ",
    "ΛΙΘΑΚ": "ΓΥΜΝΆΣΙΟ ΛΙΘΑΚΙΆΣ",
    "ΓΥΜΝΑΣΙΟΚΑΤΑΣΤ": "ΓΥΜΝΆΣΙΟ ΚΑΤΑΣΤΑΡΊΟΥ", // Distinguish from GEL
    "ΜΑΧΑΙΡ": "ΓΥΜΝΆΣΙΟ ΜΑΧΑΙΡΆΔΟΥ",
    "ΜΟΥΣΙΚ": "ΓΥΜΝΆΣΙΟ ΜΟΥΣΙΚΌ",
    "1ΟΓΕΛ": "1Ο ΓΕΛ ΖΑΚΎΝΘΟΥ",
    "2ΟΓΕΛ": "2Ο ΓΕΛ ΖΑΚΎΝΘΟΥ",
    "ΓΕΛΚΑΤΑΣΤ": "ΓΕΛ ΚΑΤΑΣΤΑΡΊΟΥ",
    "ΕΠΑΛ": "ΕΠΑΛ ΖΑΚΎΝΘΟΥ"
};

const getTodayStr = () => new Date().toISOString().split('T')[0];

// Helper to format date for display (DD/MM/YYYY)
const formatDateDisplay = (dateStr: string) => {
    if (!dateStr) return "-";
    const parts = dateStr.split('-');
    if (parts.length !== 3) return dateStr;
    return `${parts[2]}/${parts[1]}/${parts[0]}`;
};

const App: React.FC = () => {
    // --- STATE ---
    const [activeTab, setActiveTab] = useState<'scan' | 'manage' | 'history'>('scan'); 
    const [isAdmin, setIsAdmin] = useState(false); // Admin Authentication State
    
    // Login Modal State
    const [showLoginModal, setShowLoginModal] = useState(false);
    const [loginPin, setLoginPin] = useState('');

    const [students, setStudents] = useState<StudentData[]>([]);
    const [scanHistory, setScanHistory] = useState<ScanRecord[]>([]);

    // Scan Logic
    const [selectedSchoolForScan, setSelectedSchoolForScan] = useState<string | null>(null);
    const [isScanning, setIsScanning] = useState(false);
    const [scanFeedback, setScanFeedback] = useState<{msg: string, type: 'success'|'error'|'warning'|'info'}>({msg: '', type: 'info'});
    
    // History View State
    const [selectedHistoryDate, setSelectedHistoryDate] = useState<string | null>(null);

    // PDF Logic
    const [isPdfLoading, setIsPdfLoading] = useState(false);

    // --- REFS ---
    const scannerRef = useRef<Html5Qrcode | null>(null);
    const lastScannedRef = useRef<string | null>(null);
    const studentsRef = useRef(students);
    const scanHistoryRef = useRef(scanHistory);
    const selectedSchoolRef = useRef(selectedSchoolForScan);

    useEffect(() => { studentsRef.current = students; }, [students]);
    useEffect(() => { scanHistoryRef.current = scanHistory; }, [scanHistory]);
    useEffect(() => { selectedSchoolRef.current = selectedSchoolForScan; }, [selectedSchoolForScan]);

    // --- INIT / PERSISTENCE ---
    useEffect(() => {
        const savedData = localStorage.getItem(STORAGE_KEY_DATA);
        const savedHistory = localStorage.getItem(STORAGE_KEY_HISTORY);

        if (savedData) {
            try {
                const parsed = JSON.parse(savedData);
                setStudents(parsed.students || []);
            } catch (e) { console.error("Error loading data", e); }
        }

        if (savedHistory) {
            try {
                setScanHistory(JSON.parse(savedHistory));
            } catch (e) { console.error("Error loading history", e); }
        }
    }, []);

    useEffect(() => {
        localStorage.setItem(STORAGE_KEY_HISTORY, JSON.stringify(scanHistory));
    }, [scanHistory]);

    // --- ADMIN LOGIC ---
    const toggleAdmin = () => {
        if (isAdmin) {
            setIsAdmin(false);
            setActiveTab('scan');
        } else {
            setShowLoginModal(true);
            setLoginPin('');
        }
    };

    const handleLoginSubmit = (e: React.FormEvent) => {
        e.preventDefault();
        if (loginPin === '1234') {
            setIsAdmin(true);
            setShowLoginModal(false);
            setActiveTab('history'); // Auto switch to history so user sees the change
            setLoginPin('');
        } else {
            alert("Λάθος κωδικός");
        }
    };

    // --- EXPORT LOGIC (Excel) ---
    const handleExportHistory = () => {
        if (scanHistory.length === 0) {
            alert("Δεν υπάρχει ιστορικό για εξαγωγή.");
            return;
        }

        try {
            const wb = XLSX.utils.book_new();

            // Sheet 1: Summary
            const grouped = getHistoryGroupedByDateRaw(); // Helper defined below
            const summaryData = grouped.map(([date, count]) => ({
                'Ημερομηνία': formatDateDisplay(date),
                'Σύνολο Σαρώσεων': count
            }));
            const wsSummary = XLSX.utils.json_to_sheet(summaryData);
            XLSX.utils.book_append_sheet(wb, wsSummary, "Σύνολα Ημερών");

            // Sheet 2: Detailed Logs
            const detailedData = scanHistory.map(h => {
                const s = students.find(st => st.id === h.studentId);
                return {
                    'Ημερομηνία': formatDateDisplay(h.dateStr),
                    'Ώρα': new Date(h.timestamp).toLocaleTimeString('el-GR'),
                    'Ονοματεπώνυμο': s ? s.name : h.studentId,
                    'Σχολείο': s ? s.school : 'Άγνωστο',
                    'ID': h.studentId
                };
            });
            const wsDetails = XLSX.utils.json_to_sheet(detailedData);
            XLSX.utils.book_append_sheet(wb, wsDetails, "Αναλυτικό Ιστορικό");

            // Download
            XLSX.writeFile(wb, `History_Export_${getTodayStr()}.xlsx`);
        } catch (error) {
            console.error("Export error", error);
            alert("Σφάλμα κατά την εξαγωγή.");
        }
    };

    // --- MANAGE TAB: FILE UPLOAD (ROBUST) ---
    const identifySchool = (text: string): string | null => {
        const cleaned = cleanStr(text);
        if (cleaned.length < 3) return null; // Too short

        // 1. Check Keywords
        for (const [keyword, schoolName] of Object.entries(SCHOOL_KEYWORDS)) {
            if (cleaned.includes(keyword)) {
                return schoolName;
            }
        }

        // 2. Check Fixed List (Fuzzy)
        for (const school of FIXED_SCHOOLS) {
            if (cleaned.includes(cleanStr(school)) || cleanStr(school).includes(cleaned)) {
                return school;
            }
        }

        return null;
    };

    const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
        if (!event.target.files || event.target.files.length === 0) return;
        const file = event.target.files[0];
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target?.result as ArrayBuffer);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const json: (string | number)[][] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

                const loadedStudents: StudentData[] = [];
                const debugUnmatched: string[] = []; // Store samples of unmatched for feedback

                json.forEach((row) => {
                    // Convert row to string array and filter empty cells
                    const cells = (row as (string | number)[]).map(c => String(c || '').trim()).filter(c => c !== '');
                    if (cells.length < 2) return; // Need at least Name and School (or just Name if logic was different, but we need School)
                    
                    // Strategy: Scan ALL cells in the row to find a School Name
                    let foundSchool: string | null = null;
                    let nameParts: string[] = [];

                    // First pass: Look for school
                    let schoolIndex = -1;
                    for (let i = 0; i < cells.length; i++) {
                        const identified = identifySchool(cells[i]);
                        if (identified) {
                            foundSchool = identified;
                            schoolIndex = i;
                            break; // Found the school column
                        }
                    }

                    if (foundSchool) {
                        // Everything else is the name
                        nameParts = cells.filter((_, idx) => idx !== schoolIndex);
                    } else {
                        // Fallback: Assume last column is school (even if not matched yet) to report error
                        debugUnmatched.push(cells[cells.length-1] || "Άγνωστο");
                        return;
                    }

                    const fullName = nameParts.join(' ');
                    loadedStudents.push({ id: fullName, name: fullName, school: foundSchool });
                });

                if (loadedStudents.length === 0) {
                    let errorMsg = "Δεν βρέθηκαν έγκυρες εγγραφές!\n\n";
                    if (debugUnmatched.length > 0) {
                        errorMsg += `Το σύστημα διάβασε πιθανά σχολεία όπως: "${debugUnmatched.slice(0, 3).join('", "')}" αλλά δεν ταίριαζαν με τη λίστα.\n\nΠαρακαλώ ελέγξτε ότι τα ονόματα σχολείων στο Excel περιέχουν λέξεις κλειδιά όπως: '2ο Γυμνάσιο', 'Βανάτο', 'ΕΠΑΛ' κλπ.`;
                    }
                    alert(errorMsg);
                    return;
                }

                // MERGE LOGIC
                setStudents(prev => {
                    const existingMap = new Map(prev.map(s => [`${s.id}_${s.school}`, s]));
                    loadedStudents.forEach(s => {
                        existingMap.set(`${s.id}_${s.school}`, s);
                    });
                    const merged = Array.from(existingMap.values());
                    localStorage.setItem(STORAGE_KEY_DATA, JSON.stringify({ students: merged }));
                    return merged;
                });

                // Find unique schools in this upload for the success message
                const schoolsInUpload = Array.from(new Set(loadedStudents.map(s => s.school)));
                
                alert(`Επιτυχία!\n\nΦορτώθηκαν ${loadedStudents.length} μαθητές.\nΑναγνωρίστηκαν τα σχολεία:\n- ${schoolsInUpload.join('\n- ')}`);
                event.target.value = '';
                
            } catch (err) {
                console.error(err);
                alert('Σφάλμα κατά την ανάγνωση του αρχείου.');
            }
        };
        reader.readAsArrayBuffer(file);
    };

    const clearAllData = () => {
        if (confirm("ΠΡΟΣΟΧΗ: Θα διαγραφούν ΟΛΑ τα δεδομένα μαθητών και το ιστορικό σαρώσεων. Συνέχεια;")) {
            localStorage.removeItem(STORAGE_KEY_DATA);
            localStorage.removeItem(STORAGE_KEY_HISTORY);
            setStudents([]);
            setScanHistory([]);
            setScanFeedback({msg: '', type: 'info'});
        }
    };

    const clearSchoolData = (schoolName: string) => {
        if (confirm(`Διαγραφή όλων των μαθητών για: ${schoolName};`)) {
            const newStudents = students.filter(s => s.school !== schoolName);
            setStudents(newStudents);
            localStorage.setItem(STORAGE_KEY_DATA, JSON.stringify({ students: newStudents }));
        }
    };

    // --- MANAGE TAB: PDF GENERATION ---
    const generatePDF = async (schoolFilter?: string) => {
        const dataToPrint = schoolFilter 
            ? students.filter(s => s.school === schoolFilter)
            : students;

        if (dataToPrint.length === 0) {
            alert("Δεν υπάρχουν δεδομένα για εκτύπωση.");
            return;
        }
        
        setIsPdfLoading(true);

        try {
            const doc = new jsPDF();

            // 1. Load Greek-supporting font (Roboto)
            // We use a full version of Roboto from a reliable CDN that includes Greek glyphs
            const fontUrl = 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/fonts/Roboto/Roboto-Regular.ttf';
            
            const response = await fetch(fontUrl);
            if (!response.ok) throw new Error("Failed to load font");
            const blob = await response.blob();
            
            const base64Font = await new Promise<string>((resolve) => {
                const reader = new FileReader();
                reader.onloadend = () => {
                    const res = reader.result as string;
                    resolve(res.split(',')[1]);
                }
                reader.readAsDataURL(blob);
            });

            doc.addFileToVFS("Roboto-Regular.ttf", base64Font);
            doc.addFont("Roboto-Regular.ttf", "Roboto", "normal");
            doc.setFont("Roboto");

            // Layout Settings
            const pageWidth = doc.internal.pageSize.getWidth();
            const pageHeight = doc.internal.pageSize.getHeight();
            const cols = 3;
            const rows = 5;
            const cellWidth = pageWidth / cols;
            const cellHeight = pageHeight / rows;
            const qrSize = 35;

            let currentStudentIndex = 0;

            const getQR = async (text: string) => {
                return await QRCode.toDataURL(text, { errorCorrectionLevel: 'M', margin: 1 });
            };

            while (currentStudentIndex < dataToPrint.length) {
                if (currentStudentIndex > 0) doc.addPage();

                for (let r = 0; r < rows; r++) {
                    for (let c = 0; c < cols; c++) {
                        if (currentStudentIndex >= dataToPrint.length) break;

                        const s = dataToPrint[currentStudentIndex];
                        const x = c * cellWidth;
                        const y = r * cellHeight;

                        // Draw Border for cutting
                        doc.setDrawColor(220);
                        doc.rect(x, y, cellWidth, cellHeight);

                        // QR Code
                        const qrDataUrl = await getQR(s.id);
                        // Center QR
                        doc.addImage(qrDataUrl, 'PNG', x + (cellWidth - qrSize) / 2, y + 10, qrSize, qrSize);

                        // Text
                        doc.setFontSize(10);
                        doc.setTextColor(0);
                        
                        // Handle long names
                        const splitName = doc.splitTextToSize(s.name, cellWidth - 10);
                        // Calculate vertical position to center text block
                        doc.text(splitName, x + cellWidth / 2, y + qrSize + 20, { align: 'center' });
                        
                        doc.setFontSize(8);
                        doc.setTextColor(100);
                        const splitSchool = doc.splitTextToSize(s.school, cellWidth - 10);
                        // Position below name
                        const nameHeight = splitName.length * 4; // approx 4 units per line
                        doc.text(splitSchool, x + cellWidth / 2, y + qrSize + 22 + nameHeight, { align: 'center' });

                        currentStudentIndex++;
                    }
                }
            }

            const fileName = schoolFilter ? `QR_${schoolFilter}.pdf` : 'QR_ALL_SCHOOLS.pdf';
            doc.save(fileName);
        } catch (error) {
            console.error("PDF Error:", error);
            alert("Σφάλμα κατά τη δημιουργία PDF: " + (error as any).message + "\n\nΒεβαιωθείτε ότι έχετε σύνδεση στο Internet για τη φόρτωση της Ελληνικής γραμματοσειράς.");
        } finally {
            setIsPdfLoading(false);
        }
    };

    // --- SCANNER LOGIC ---
    
    const onScanSuccess = (decodedText: string, decodedResult: any) => {
        if (lastScannedRef.current === decodedText) return;

        // 1. Check Weekend (Requirement)
        const now = new Date();
        const dayOfWeek = now.getDay(); // 0 = Sunday, 6 = Saturday
        if (dayOfWeek === 0 || dayOfWeek === 6) {
             // Block scan
             if (lastScannedRef.current !== 'blocked_weekend') {
                playErrorSound();
                setScanFeedback({ msg: 'Η σάρωση δεν επιτρέπεται Σαββατοκύριακα!', type: 'error' });
                lastScannedRef.current = 'blocked_weekend';
                setTimeout(() => { lastScannedRef.current = null; }, 3000);
             }
             return;
        }

        lastScannedRef.current = decodedText;
        setTimeout(() => { lastScannedRef.current = null; }, 3000);

        // Match ID
        const student = studentsRef.current.find(s => s.id === decodedText);
        
        if (!student) {
            playErrorSound();
            setScanFeedback({ msg: `Άγνωστος Κωδικός: ${decodedText}`, type: 'error' });
            return;
        }

        // Match School
        if (selectedSchoolRef.current && student.school !== selectedSchoolRef.current) {
             playErrorSound();
             setScanFeedback({ 
                 msg: `Λάθος Σχολείο! Ανήκει στο: ${student.school}`, 
                 type: 'warning' 
             });
             return;
        }

        // Duplicate Check
        const today = getTodayStr();
        const alreadyScanned = scanHistoryRef.current.find(
            h => h.studentId === decodedText && h.dateStr === today
        );

        if (alreadyScanned) {
            playErrorSound();
            setScanFeedback({ msg: `Ήδη σαρρώθηκε σήμερα: ${student.name}`, type: 'warning' });
            return;
        }

        playSuccessSound();
        const newRecord: ScanRecord = {
            studentId: student.id,
            timestamp: Date.now(),
            dateStr: today
        };

        setScanHistory(prev => [newRecord, ...prev]);
        setScanFeedback({ msg: `Επιτυχία: ${student.name}`, type: 'success' });
    };

    const playSuccessSound = () => {
        const ctx = new (window.AudioContext || (window as any).webkitAudioContext)();
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.type = 'sine';
        osc.frequency.setValueAtTime(500, ctx.currentTime);
        osc.frequency.exponentialRampToValueAtTime(1000, ctx.currentTime + 0.1);
        gain.gain.setValueAtTime(0.1, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.1);
        osc.start();
        osc.stop(ctx.currentTime + 0.2);
    };

    const playErrorSound = () => {
        const ctx = new (window.AudioContext || (window as any).webkitAudioContext)();
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.type = 'sawtooth';
        osc.frequency.setValueAtTime(150, ctx.currentTime);
        osc.frequency.linearRampToValueAtTime(100, ctx.currentTime + 0.3);
        gain.gain.setValueAtTime(0.1, ctx.currentTime);
        gain.gain.linearRampToValueAtTime(0.01, ctx.currentTime + 0.3);
        osc.start();
        osc.stop(ctx.currentTime + 0.3);
    };

    // --- SCANNER UI HANDLERS ---
    const startScanner = (school: string) => {
        setSelectedSchoolForScan(school);
        setIsScanning(true);
        setScanFeedback({ msg: 'Έτοιμο για σάρωση...', type: 'info' });
    };

    const stopScanner = () => {
        if (scannerRef.current) {
            scannerRef.current.stop().then(() => {
                scannerRef.current?.clear();
                scannerRef.current = null;
            }).catch(err => console.error("Failed to stop scanner", err));
        }
        setIsScanning(false);
        setSelectedSchoolForScan(null);
        lastScannedRef.current = null;
    };

    useEffect(() => {
        let html5QrcodeScanner: Html5Qrcode | null = null;

        if (isScanning) {
            const timer = setTimeout(() => {
                html5QrcodeScanner = new Html5Qrcode("qr-reader");
                scannerRef.current = html5QrcodeScanner;

                const config = { fps: 10, qrbox: { width: 250, height: 250 } };
                
                html5QrcodeScanner.start(
                    { facingMode: "environment" },
                    config,
                    onScanSuccess,
                    (errorMessage) => {}
                ).catch(err => {
                    console.error("Error starting scanner", err);
                    setScanFeedback({ msg: 'Δεν βρέθηκε κάμερα', type: 'error' });
                });
            }, 100);

            return () => {
                clearTimeout(timer);
                if (html5QrcodeScanner && html5QrcodeScanner.isScanning) {
                    html5QrcodeScanner.stop().then(() => html5QrcodeScanner?.clear());
                }
            };
        }
    }, [isScanning]);

    // --- HELPERS ---
    const hasDataForSchool = (schoolName: string) => {
        return students.some(s => s.school === schoolName);
    };

    const getScansForSchoolToday = (schoolName: string) => {
        const today = getTodayStr();
        const schoolStudentIds = new Set(students.filter(s => s.school === schoolName).map(s => s.id));
        const todayScans = scanHistory.filter(h => 
            h.dateStr === today && schoolStudentIds.has(h.studentId)
        );
        return todayScans.length;
    };

    // --- HISTORY TAB LOGIC ---
    const getHistoryGroupedByDateRaw = () => {
        const grouped: {[key: string]: number} = {};
        scanHistory.forEach(h => {
            grouped[h.dateStr] = (grouped[h.dateStr] || 0) + 1;
        });
        // Sort descending (newest first)
        return Object.entries(grouped).sort((a, b) => b[0].localeCompare(a[0]));
    };

    const getHistoryGroupedByDate = () => getHistoryGroupedByDateRaw();

    return (
        <div className="min-h-screen flex flex-col font-sans bg-gray-50 relative">
            {/* HEADER */}
            <header className="bg-blue-700 text-white p-4 shadow-md sticky top-0 z-40 flex justify-between items-center">
                <div className="flex-1">
                    <h1 className="text-xl font-bold">QR Scanner Σχολείων</h1>
                </div>
                <button 
                    onClick={toggleAdmin}
                    className={`p-2 rounded-full transition-colors ${isAdmin ? 'bg-red-500 hover:bg-red-600' : 'bg-blue-600 hover:bg-blue-500'}`}
                    title={isAdmin ? "Έξοδος Διαχειριστή" : "Είσοδος Διαχειριστή"}
                >
                    {isAdmin ? (
                        <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17 16l4-4m0 0l-4-4m4 4H7m6 4v1a3 3 0 01-3 3H6a3 3 0 01-3-3V7a3 3 0 013-3h4a3 3 0 013 3v1" /></svg>
                    ) : (
                        <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" /></svg>
                    )}
                </button>
            </header>

            {/* TABS */}
            <div className="flex border-b bg-white shadow-sm sticky top-16 z-30 text-sm sm:text-base">
                <button 
                    className={`flex-1 py-3 sm:py-4 text-center font-semibold ${activeTab === 'scan' ? 'border-b-4 border-blue-600 text-blue-700 bg-blue-50' : 'text-gray-500'}`}
                    onClick={() => setActiveTab('scan')}
                >
                    Σάρωση
                </button>
                {isAdmin && (
                    <button 
                        className={`flex-1 py-3 sm:py-4 text-center font-semibold ${activeTab === 'history' ? 'border-b-4 border-blue-600 text-blue-700 bg-blue-50' : 'text-gray-500'}`}
                        onClick={() => setActiveTab('history')}
                    >
                        Ιστορικό
                    </button>
                )}
                {isAdmin && (
                    <button 
                        className={`flex-1 py-3 sm:py-4 text-center font-semibold ${activeTab === 'manage' ? 'border-b-4 border-blue-600 text-blue-700 bg-blue-50' : 'text-gray-500'}`}
                        onClick={() => setActiveTab('manage')}
                    >
                        Διαχείριση
                    </button>
                )}
            </div>

            {/* CONTENT */}
            <main className="flex-1 p-4 max-w-4xl mx-auto w-full">
                
                {/* --- TAB: SCAN --- */}
                {activeTab === 'scan' && (
                    <div>
                        {!isScanning ? (
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                {FIXED_SCHOOLS.map((school) => {
                                    const hasData = hasDataForSchool(school);
                                    const countToday = getScansForSchoolToday(school);
                                    
                                    return (
                                        <button
                                            key={school}
                                            onClick={() => hasData && startScanner(school)}
                                            disabled={!hasData}
                                            className={`p-6 rounded-xl shadow-sm border-2 text-left transition-all duration-200 relative overflow-hidden group
                                                ${hasData 
                                                    ? 'bg-white border-blue-100 hover:border-blue-400 hover:shadow-md cursor-pointer' 
                                                    : 'bg-gray-100 border-transparent opacity-60 cursor-not-allowed'
                                                }
                                            `}
                                        >
                                            <div className="relative z-10">
                                                <h3 className={`font-bold text-lg mb-1 ${hasData ? 'text-gray-800' : 'text-gray-500'}`}>
                                                    {school}
                                                </h3>
                                                <div className="flex items-center justify-between">
                                                    <span className={`text-sm px-2 py-0.5 rounded-full ${hasData ? 'bg-green-100 text-green-800' : 'bg-gray-200 text-gray-500'}`}>
                                                        {hasData ? 'Ενεργό' : 'Χωρίς Δεδομένα'}
                                                    </span>
                                                    {hasData && (
                                                        <span className="text-sm font-medium text-blue-600">
                                                            Σημερινές: {countToday}
                                                        </span>
                                                    )}
                                                </div>
                                            </div>
                                            {hasData && (
                                                <div className="absolute bottom-0 right-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
                                                    <svg className="w-12 h-12 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v1m6 11h2m-6 0h-2v4m0-11v3m0 0h.01M12 12h4.01M16 20h4M4 12h4m12 0h.01M5 8h2a1 1 0 001-1V5a1 1 0 00-1-1H5a1 1 0 00-1 1v2a1 1 0 001 1zm12 0h2a1 1 0 001-1V5a1 1 0 00-1-1h-2a1 1 0 00-1 1v2a1 1 0 001 1zM5 20h2a1 1 0 001-1v-2a1 1 0 00-1-1H5a1 1 0 00-1 1v2a1 1 0 001 1z" /></svg>
                                                </div>
                                            )}
                                        </button>
                                    );
                                })}
                            </div>
                        ) : (
                            <div id="qr-reader-container">
                                <div className="w-full max-w-lg relative flex flex-col items-center">
                                    <div className="w-full flex justify-between items-center mb-4 text-white px-2">
                                        <div>
                                            <h2 className="font-bold text-lg">{selectedSchoolForScan}</h2>
                                            <p className="text-xs opacity-80">Σάρωση μαθητών...</p>
                                        </div>
                                        <button 
                                            onClick={stopScanner}
                                            className="bg-white/20 hover:bg-white/30 text-white px-4 py-2 rounded-full text-sm backdrop-blur-sm"
                                        >
                                            Κλείσιμο
                                        </button>
                                    </div>

                                    <div id="qr-reader"></div>
                                    <div className="scanner-overlay"><div></div></div>

                                    <div className={`mt-6 w-full p-4 rounded-lg text-center font-medium text-lg shadow-lg transition-all transform duration-300 ${
                                        scanFeedback.type === 'success' ? 'bg-green-500 text-white scale-105' :
                                        scanFeedback.type === 'error' ? 'bg-red-500 text-white shake' :
                                        scanFeedback.type === 'warning' ? 'bg-orange-500 text-white' :
                                        'bg-gray-800 text-gray-300'
                                    }`}>
                                        {scanFeedback.msg}
                                    </div>
                                </div>
                            </div>
                        )}
                    </div>
                )}

                {/* --- TAB: HISTORY --- */}
                {activeTab === 'history' && isAdmin && (
                    <div className="space-y-6">
                        {!selectedHistoryDate ? (
                            // LIST OF DATES
                            <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
                                <div className="p-4 bg-gray-50 border-b border-gray-200 flex justify-between items-center">
                                    <h2 className="text-lg font-bold text-gray-800">Ιστορικό ανά Ημέρα</h2>
                                    <button
                                        onClick={handleExportHistory}
                                        className="flex items-center bg-green-600 text-white px-3 py-1.5 rounded hover:bg-green-700 text-sm font-medium shadow-sm"
                                    >
                                        <svg className="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
                                        Εξαγωγή Excel (Drive)
                                    </button>
                                </div>
                                <div className="divide-y divide-gray-100">
                                    {getHistoryGroupedByDate().length === 0 ? (
                                        <div className="p-8 text-center text-gray-500 italic">Δεν υπάρχουν καταγεγραμμένες σαρώσεις.</div>
                                    ) : (
                                        getHistoryGroupedByDate().map(([dateStr, count]) => (
                                            <button 
                                                key={dateStr}
                                                onClick={() => setSelectedHistoryDate(dateStr)}
                                                className="w-full flex items-center justify-between p-4 hover:bg-blue-50 transition-colors text-left"
                                            >
                                                <div className="flex items-center">
                                                    <div className="bg-blue-100 text-blue-700 p-2 rounded-lg mr-3">
                                                        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" /></svg>
                                                    </div>
                                                    <span className="font-medium text-gray-800 text-lg">{formatDateDisplay(dateStr)}</span>
                                                </div>
                                                <div className="flex items-center text-gray-500">
                                                    <span className="mr-2 font-bold text-blue-600">{count}</span>
                                                    <span className="text-xs uppercase tracking-wide">Σαρωσεις</span>
                                                    <svg className="w-4 h-4 ml-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" /></svg>
                                                </div>
                                            </button>
                                        ))
                                    )}
                                </div>
                            </div>
                        ) : (
                            // DETAILS FOR SELECTED DATE
                            <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden flex flex-col h-[calc(100vh-150px)]">
                                <div className="p-4 bg-gray-50 border-b border-gray-200 flex items-center justify-between shrink-0">
                                    <button 
                                        onClick={() => setSelectedHistoryDate(null)}
                                        className="flex items-center text-blue-600 hover:text-blue-800 font-medium text-sm"
                                    >
                                        <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" /></svg>
                                        Πίσω
                                    </button>
                                    <h2 className="text-lg font-bold text-gray-800">{formatDateDisplay(selectedHistoryDate)}</h2>
                                    <div className="w-10"></div> {/* Spacer for centering */}
                                </div>
                                <div className="overflow-y-auto p-0 flex-1">
                                    <table className="w-full text-sm text-left">
                                        <thead className="text-xs text-gray-700 uppercase bg-gray-50 sticky top-0">
                                            <tr>
                                                <th className="px-4 py-3">Ωρα</th>
                                                <th className="px-4 py-3">Ονοματεπωνυμο</th>
                                                <th className="px-4 py-3">Σχολειο</th>
                                            </tr>
                                        </thead>
                                        <tbody className="divide-y divide-gray-100">
                                            {scanHistory
                                                .filter(h => h.dateStr === selectedHistoryDate)
                                                .sort((a, b) => b.timestamp - a.timestamp)
                                                .map((record, idx) => {
                                                    // Find student details to show school
                                                    const studentDetails = students.find(s => s.id === record.studentId);
                                                    const timeStr = new Date(record.timestamp).toLocaleTimeString('el-GR', {hour: '2-digit', minute: '2-digit'});
                                                    
                                                    return (
                                                        <tr key={idx} className="hover:bg-gray-50">
                                                            <td className="px-4 py-3 font-mono text-gray-500">{timeStr}</td>
                                                            <td className="px-4 py-3 font-medium text-gray-900">{record.studentId}</td>
                                                            <td className="px-4 py-3 text-gray-500 text-xs">
                                                                {studentDetails ? studentDetails.school : "Άγνωστο"}
                                                            </td>
                                                        </tr>
                                                    );
                                                })}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        )}
                    </div>
                )}

                {/* --- TAB: MANAGE --- */}
                {activeTab === 'manage' && isAdmin && (
                    <div className="space-y-8">
                        
                        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
                            <h2 className="text-lg font-bold mb-4 flex items-center text-gray-800">
                                <span className="bg-blue-100 text-blue-600 p-2 rounded-lg mr-3">1</span>
                                Φόρτωση Δεδομένων (Excel)
                            </h2>
                            <p className="text-sm text-gray-600 mb-4">
                                Ανεβάστε αρχείο Excel. Το σύστημα θα αναγνωρίσει αυτόματα το σχολείο (π.χ. "2ο Γυμνάσιο", "Βανάτο") ανεξάρτητα από τη στήλη.
                            </p>
                            
                            <label className="block w-full cursor-pointer">
                                <input 
                                    type="file" 
                                    accept=".xlsx, .xls" 
                                    onChange={handleFileUpload} 
                                    className="block w-full text-sm text-gray-500
                                    file:mr-4 file:py-2.5 file:px-4
                                    file:rounded-full file:border-0
                                    file:text-sm file:font-semibold
                                    file:bg-blue-50 file:text-blue-700
                                    hover:file:bg-blue-100"
                                />
                            </label>
                        </div>

                        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
                            <h2 className="text-lg font-bold mb-4 flex items-center text-gray-800">
                                <span className="bg-indigo-100 text-indigo-600 p-2 rounded-lg mr-3">2</span>
                                Διαχείριση & Εκτύπωση ανά Σχολείο
                            </h2>

                            {students.length === 0 ? (
                                <p className="text-gray-400 text-center py-4 italic">Δεν υπάρχουν δεδομένα.</p>
                            ) : (
                                <div className="space-y-4">
                                    {FIXED_SCHOOLS.map(school => {
                                        const count = students.filter(s => s.school === school).length;
                                        if (count === 0) return null;

                                        return (
                                            <div key={school} className="flex flex-col sm:flex-row sm:items-center justify-between bg-gray-50 p-4 rounded-lg border border-gray-100 hover:border-indigo-200 transition-colors">
                                                <div className="mb-3 sm:mb-0">
                                                    <h3 className="font-bold text-gray-800">{school}</h3>
                                                    <span className="text-xs text-gray-500">{count} Μαθητές</span>
                                                </div>
                                                <div className="flex gap-2">
                                                    <button
                                                        onClick={() => generatePDF(school)}
                                                        disabled={isPdfLoading}
                                                        className={`flex items-center px-3 py-2 border rounded text-sm font-medium transition-colors ${isPdfLoading ? 'bg-gray-100 text-gray-400 border-gray-200' : 'bg-white border-purple-200 text-purple-700 hover:bg-purple-50'}`}
                                                    >
                                                        {isPdfLoading ? (
                                                             <span className="animate-pulse">Φόρτωση...</span>
                                                        ) : (
                                                            <>
                                                                <svg className="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" /></svg>
                                                                PDF
                                                            </>
                                                        )}
                                                    </button>
                                                    <button
                                                        onClick={() => clearSchoolData(school)}
                                                        className="flex items-center px-3 py-2 bg-white border border-red-200 text-red-600 rounded hover:bg-red-50 text-sm font-medium"
                                                        title="Διαγραφή δεδομένων αυτού του σχολείου"
                                                    >
                                                        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" /></svg>
                                                    </button>
                                                </div>
                                            </div>
                                        );
                                    })}
                                </div>
                            )}
                        </div>

                        <div className="flex justify-between items-center pt-4 border-t">
                             <div className="text-sm text-gray-500">
                                 Σύνολο Μαθητών: <b>{students.length}</b> | Σαρώσεις: <b>{scanHistory.length}</b>
                             </div>
                             <button 
                                onClick={clearAllData}
                                className="text-red-500 text-sm hover:text-red-700 hover:underline"
                            >
                                Διαγραφή ΟΛΩΝ
                            </button>
                        </div>

                    </div>
                )}
            </main>

            {/* --- LOGIN MODAL --- */}
            {showLoginModal && (
                <div className="fixed inset-0 bg-black/50 z-[60] flex items-center justify-center p-4 backdrop-blur-sm">
                    <div className="bg-white rounded-xl p-6 w-full max-w-sm shadow-2xl animate-[fadeIn_0.2s_ease-out]">
                        <div className="text-center mb-6">
                            <div className="bg-blue-100 w-16 h-16 rounded-full flex items-center justify-center mx-auto mb-4">
                                <svg className="w-8 h-8 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" /></svg>
                            </div>
                            <h3 className="text-xl font-bold text-gray-800">Είσοδος Διαχειριστή</h3>
                            <p className="text-sm text-gray-500 mt-1">Πληκτρολογήστε το PIN για πρόσβαση.</p>
                        </div>
                        <form onSubmit={handleLoginSubmit}>
                            <input
                                type="password"
                                value={loginPin}
                                onChange={e => setLoginPin(e.target.value)}
                                className="w-full border-2 border-gray-200 focus:border-blue-500 focus:ring-blue-500 p-3 rounded-xl mb-6 text-2xl text-center tracking-widest outline-none transition-all"
                                placeholder="PIN"
                                autoFocus
                                pattern="[0-9]*" 
                                inputMode="numeric"
                            />
                            <div className="flex gap-3">
                                <button
                                    type="button"
                                    onClick={() => { setShowLoginModal(false); setLoginPin(''); }}
                                    className="flex-1 py-3 text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-xl font-medium transition-colors"
                                >
                                    Ακύρωση
                                </button>
                                <button
                                    type="submit"
                                    className="flex-1 py-3 text-white bg-blue-600 hover:bg-blue-700 rounded-xl font-bold shadow-md shadow-blue-200 transition-colors"
                                >
                                    Είσοδος
                                </button>
                            </div>
                        </form>
                    </div>
                </div>
            )}
        </div>
    );
};

export default App;
