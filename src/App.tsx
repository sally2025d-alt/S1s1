import React, { useState, useEffect } from "react";
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { useFinanceStore } from "./hooks/useFinanceStore";
import { FinanceTab } from "./components/FinanceTab";
import { PaymentTab } from "./components/PaymentTab";
import { ReportsTab } from "./components/ReportsTab";
import { ArchiveTab } from "./components/ArchiveTab";
import { Department, Person, BASE_AMOUNT } from "./types";
import { motion, AnimatePresence } from "motion/react";
import { Briefcase, CreditCard, BarChart2, Settings, CheckCircle2, X, FileSpreadsheet, RotateCcw, Info, Upload, Archive } from "lucide-react";
import { Input } from "./components/ui/Input";
import { Button } from "./components/ui/Button";
import * as XLSX from "xlsx";

type Tab = "finance" | "payment" | "reports" | "archive";

export default function App() {
  const { state, updateCovenant, addExpense, confirmPayment, cancelPayment, resetDay, importData, deleteArchive, clearArchives, stats } = useFinanceStore();
  const [activeTab, setActiveTab] = useState<Tab>("finance");
  
  // Modal State
  const [selectedPerson, setSelectedPerson] = useState<{ deptId: number; person: Person } | null>(null);
  const [bonus, setBonus] = useState<number>(0);
  const [bonusInput, setBonusInput] = useState("");
  
  // UI State
  const [showSuccess, setShowSuccess] = useState(false);
  const [showSettingsModal, setShowSettingsModal] = useState(false);
  const [showImportModal, setShowImportModal] = useState(false);
  const [showCovenantModal, setShowCovenantModal] = useState(false);
  const [showResetModal, setShowResetModal] = useState(false);
  const [newCovenantInput, setNewCovenantInput] = useState("");
  const [importFile, setImportFile] = useState<File | null>(null);
  const [toast, setToast] = useState<{ message: string; visible: boolean }>({ message: "", visible: false });

  const showToast = (message: string) => {
    setToast({ message, visible: true });
    setTimeout(() => setToast({ message: "", visible: false }), 3000);
  };

  const handleSelectPerson = (deptId: number, personId: string) => {
    const dept = state.departments.find(d => d.id === deptId);
    const person = dept?.persons.find(p => p.id === personId);
    
    if (!person || !dept) return;
    if (person.received) {
      showToast("⚠️ استلم مسبقاً");
      return;
    }

    const personTotal = person.baseAmount;
    if (personTotal > stats.remaining) {
      showToast("❌ العهد غير كافٍ!");
      return;
    }

    setSelectedPerson({ deptId, person });
    setBonus(0);
    setBonusInput("");
  };

  const handleAddBonus = () => {
    const amount = parseInt(bonusInput);
    if (!isNaN(amount) && amount > 0) {
      const base = selectedPerson?.person.baseAmount || BASE_AMOUNT;
      if ((base + bonus + amount) > stats.remaining) {
        showToast("❌ الزيادة تتجاوز العهد!");
        return;
      }
      setBonus(prev => prev + amount);
      setBonusInput("");
      showToast(`✓ +${amount.toLocaleString()} ر.ي`);
    }
  };

  const handleConfirmPayment = () => {
    if (!selectedPerson) return;
    
    const base = selectedPerson.person.baseAmount;
    const total = base + bonus;
    
    if (total > stats.remaining) {
      showToast("❌ العهد غير كافٍ!");
      setSelectedPerson(null);
      return;
    }

    confirmPayment(selectedPerson.deptId, selectedPerson.person.id, bonus);
    setSelectedPerson(null);
    setShowSuccess(true);
    
    if (navigator.vibrate) navigator.vibrate([100, 50, 100]);
    
    setTimeout(() => {
      setShowSuccess(false);
    }, 2000);
  };

  const processImportFile = () => {
    if (!importFile) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const ab = evt.target?.result;
        const wb = XLSX.read(ab, { type: "array" });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        
        // Filter out empty rows
        const rows = data.filter((row: any) => row && row.length > 0 && row.some((cell: any) => cell !== undefined && cell !== null && cell !== ''));
        
        if (rows && rows.length > 1) {
          const rowsCount = rows.length - 1; // Subtract header
          importData(rows);
          showToast(`✓ تم استيراد ${rowsCount} سجل بنجاح`);
          setShowImportModal(false);
          setImportFile(null);
        } else {
          showToast("❌ الملف فارغ أو لا يحتوي على بيانات صحيحة");
        }
      } catch (error) {
        console.error("Error importing file:", error);
        showToast("❌ حدث خطأ أثناء قراءة الملف");
      }
    };
    reader.readAsArrayBuffer(importFile);
  };

  const exportToExcel = async () => {
    const receivedData: any[] = [];
    const pendingData: any[] = [];
    const expensesData: any[] = [];
    let totalPaid = 0;

    state.departments.forEach(d => {
      d.persons.forEach(p => {
        const amount = p.totalAmount;
        
        const rowData = {
          dept: d.name,
          rank: p.rank,
          name: p.name,
          militaryNumber: p.militaryNumber,
          baseAmount: p.baseAmount,
          bonus: p.bonus,
          total: amount,
          date: p.date || "-",
          time: p.time || "-"
        };

        if (p.received) {
          totalPaid += amount;
          receivedData.push(rowData);
        } else {
          pendingData.push(rowData);
        }
      });
    });

    state.expenses.forEach(e => {
      totalPaid += e.amount;
      expensesData.push({
        recipient: e.recipient,
        purpose: e.purpose,
        amount: e.amount,
        date: e.date,
        time: e.time
      });
    });

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'صراف اللواء';
    workbook.created = new Date();

    const createStyledSheet = (name: string, data: any[], type: 'summary' | 'persons' | 'expenses' = 'persons') => {
      const ws = workbook.addWorksheet(name, { views: [{ rightToLeft: true }] });
      
      if (type === 'summary') {
        ws.columns = [
          { header: 'البيان', key: 'label', width: 30 },
          { header: 'القيمة', key: 'value', width: 20 }
        ];
      } else if (type === 'expenses') {
        ws.columns = [
          { header: 'المستلم', key: 'recipient', width: 30 },
          { header: 'الغرض', key: 'purpose', width: 40 },
          { header: 'المبلغ', key: 'amount', width: 20 },
          { header: 'التاريخ', key: 'date', width: 15 },
          { header: 'الوقت', key: 'time', width: 15 },
        ];
      } else {
        ws.columns = [
          { header: 'القسم', key: 'dept', width: 20 },
          { header: 'الرتبة', key: 'rank', width: 15 },
          { header: 'الاسم', key: 'name', width: 30 },
          { header: 'الرقم العسكري', key: 'militaryNumber', width: 20 },
          { header: 'المبلغ الأساسي', key: 'baseAmount', width: 18 },
          { header: 'الزيادة', key: 'bonus', width: 15 },
          { header: 'الإجمالي', key: 'total', width: 18 },
          { header: 'التاريخ', key: 'date', width: 15 },
          { header: 'الوقت', key: 'time', width: 15 },
        ];
      }

      // Style Header
      const headerRow = ws.getRow(1);
      headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Arial' };
      headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }; // Slate 800
      headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
      headerRow.height = 30;

      // Add Data
      data.forEach((item, index) => {
        const row = ws.addRow(item);
        row.alignment = { vertical: 'middle', horizontal: 'center' };
        row.font = { size: 11, name: 'Arial', bold: type === 'summary' };
        row.height = 25;
        
        // Alternating row colors
        if (index % 2 === 0) {
          row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } }; // Slate 50
        } else {
          row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } }; // White
        }

        // Highlight specific summary rows
        if (type === 'summary' && item.label === 'المتبقي من العهد') {
          row.font = { size: 12, name: 'Arial', bold: true, color: { argb: 'FF16A34A' } }; // Green
        }
        if (type === 'summary' && item.label === 'إجمالي المصروف') {
          row.font = { size: 12, name: 'Arial', bold: true, color: { argb: 'FFDC2626' } }; // Red
        }
        
        // Borders for all cells in row
        row.eachCell((cell) => {
          cell.border = {
            top: {style:'thin', color: {argb:'FFE2E8F0'}},
            left: {style:'thin', color: {argb:'FFE2E8F0'}},
            bottom: {style:'thin', color: {argb:'FFE2E8F0'}},
            right: {style:'thin', color: {argb:'FFE2E8F0'}}
          };
        });
      });

      // Header borders
      headerRow.eachCell((cell) => {
        cell.border = {
          top: {style:'thin', color: {argb:'FF94A3B8'}},
          left: {style:'thin', color: {argb:'FF94A3B8'}},
          bottom: {style:'thin', color: {argb:'FF94A3B8'}},
          right: {style:'thin', color: {argb:'FF94A3B8'}}
        };
      });
    };

    // 1. المستلمين (Received)
    createStyledSheet("المستلمين", receivedData, 'persons');

    // 2. غير المستلمين (Pending)
    createStyledSheet("غير المستلمين", pendingData, 'persons');
    
    // 3. أوامر الصرف (Expenses)
    createStyledSheet("أوامر الصرف", expensesData, 'expenses');

    // 4. الملخص المالي (Summary)
    const summaryData = [
      { label: "إجمالي العهد", value: state.covenant },
      { label: "إجمالي المصروف", value: totalPaid },
      { label: "المتبقي من العهد", value: state.covenant - totalPaid },
      { label: "إجمالي الأفراد", value: stats.totalPersons },
      { label: "عدد المستلمين", value: stats.receivedPersons },
      { label: "عدد المتبقين", value: stats.pendingPersons },
      { label: "عدد أوامر الصرف", value: state.expenses.length },
    ];
    createStyledSheet("الملخص المالي", summaryData, 'summary');

    const date = new Date().toLocaleDateString("ar-SA").replace(/\//g, "-");
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `تقرير_صرف_اللواء_${date}.xlsx`);
    showToast("📊 تم تصدير التقرير باحترافية");
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans selection:bg-blue-200">
      {/* Header */}
      <header className="bg-gradient-to-br from-slate-900 to-slate-800 text-white p-4 pb-6 shadow-lg sticky top-0 z-10">
        <div className="flex justify-between items-center mb-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-white/10 rounded-xl flex items-center justify-center text-xl backdrop-blur-sm border border-white/10">
              💰
            </div>
            <h1 className="font-black text-xl tracking-tight">صراف اللواء</h1>
          </div>
          <button 
            onClick={() => setShowSettingsModal(true)}
            className="w-10 h-10 rounded-xl bg-white/10 flex items-center justify-center hover:bg-white/20 transition-colors border border-white/10 active:scale-95"
          >
            <Settings className="w-5 h-5" />
          </button>
        </div>

        {/* Quick Stats Scroll */}
        <div className="flex gap-3 overflow-x-auto no-scrollbar pb-1">
          <div className="bg-white/10 border border-white/10 backdrop-blur-md rounded-2xl p-3 min-w-[140px]">
            <div className="text-xs font-bold opacity-80 mb-1">💼 العهد</div>
            <div className="text-lg font-black">{state.covenant.toLocaleString()} ر.ي</div>
          </div>
          <div className="bg-rose-500/20 border border-rose-500/30 backdrop-blur-md rounded-2xl p-3 min-w-[140px]">
            <div className="text-xs font-bold opacity-80 mb-1 text-rose-100">💸 المصروف</div>
            <div className="text-lg font-black text-rose-50">{stats.totalSpent.toLocaleString()} ر.ي</div>
          </div>
          <div className="bg-emerald-500/20 border border-emerald-500/30 backdrop-blur-md rounded-2xl p-3 min-w-[140px]">
            <div className="text-xs font-bold opacity-80 mb-1 text-emerald-100">✅ المتبقي</div>
            <div className="text-lg font-black text-emerald-50">{stats.remaining.toLocaleString()} ر.ي</div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="p-4">
        <AnimatePresence mode="wait">
          {activeTab === "finance" && (
            <motion.div key="finance" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }}>
              <FinanceTab covenant={state.covenant} stats={stats} onUpdateCovenant={updateCovenant} onAddExpense={addExpense} departments={state.departments} expenses={state.expenses} />
            </motion.div>
          )}
          {activeTab === "payment" && (
            <motion.div key="payment" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }}>
              <PaymentTab departments={state.departments} onSelectPerson={handleSelectPerson} stats={stats} onImportData={importData} onCancelPayment={cancelPayment} />
            </motion.div>
          )}
          {activeTab === "reports" && (
            <motion.div key="reports" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }}>
              <ReportsTab covenant={state.covenant} stats={stats} departments={state.departments} expenses={state.expenses} />
            </motion.div>
          )}
          {activeTab === "archive" && (
            <motion.div key="archive" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }}>
              <ArchiveTab archives={state.archives || []} onDeleteArchive={deleteArchive} onClearArchives={clearArchives} />
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Bottom Navigation */}
      <nav className="fixed bottom-0 left-0 right-0 bg-white border-t border-slate-200 px-4 py-3 pb-safe flex justify-between items-center z-40 shadow-[0_-10px_40px_rgba(0,0,0,0.05)]">
        <NavItem 
          icon={<Briefcase className="w-6 h-6" />} 
          label="المالية" 
          active={activeTab === "finance"} 
          onClick={() => setActiveTab("finance")} 
        />
        <NavItem 
          icon={<CreditCard className="w-6 h-6" />} 
          label="الصرف" 
          active={activeTab === "payment"} 
          onClick={() => setActiveTab("payment")} 
        />
        <NavItem 
          icon={<BarChart2 className="w-6 h-6" />} 
          label="التقارير" 
          active={activeTab === "reports"} 
          onClick={() => setActiveTab("reports")} 
        />
        <NavItem 
          icon={<Archive className="w-6 h-6" />} 
          label="الأرشيف" 
          active={activeTab === "archive"} 
          onClick={() => setActiveTab("archive")} 
        />
      </nav>

      {/* Payment Modal */}
      <AnimatePresence>
        {selectedPerson && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.15 }}
            className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-end sm:items-center justify-center p-4 pb-safe"
            onClick={() => setSelectedPerson(null)}
          >
            <motion.div
              initial={{ y: "100%" }}
              animate={{ y: 0 }}
              exit={{ y: "100%" }}
              transition={{ type: "spring", damping: 25, stiffness: 300 }}
              className="bg-white rounded-3xl w-full max-w-md overflow-hidden shadow-2xl"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="p-4 border-b border-slate-100 flex justify-between items-center bg-slate-50">
                <div>
                  <h2 className="text-lg font-black text-slate-900">{selectedPerson.person.name}</h2>
                  <p className="text-xs font-bold text-slate-500">{selectedPerson.person.militaryNumber} • {state.departments.find(d => d.id === selectedPerson.deptId)?.name}</p>
                </div>
                <button
                  onClick={() => setSelectedPerson(null)}
                  className="w-8 h-8 bg-slate-200 rounded-full flex items-center justify-center text-slate-600 hover:bg-slate-300 transition-colors"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>

              <div className="p-4 space-y-4">
                <div className="flex justify-between items-center bg-emerald-50 p-3 rounded-xl border border-emerald-100">
                  <span className="text-sm font-bold text-emerald-700">المبلغ الأساسي</span>
                  <span className="text-xl font-black text-emerald-700">{selectedPerson.person.baseAmount.toLocaleString()} ر.ي</span>
                </div>

                <div className="flex gap-2">
                  <Input
                    type="number"
                    placeholder="مبلغ الزيادة (اختياري)"
                    value={bonusInput}
                    onChange={(e) => setBonusInput(e.target.value)}
                    className="bg-slate-50 border-slate-200 text-center font-bold h-12 flex-1"
                  />
                  <Button onClick={handleAddBonus} className="bg-slate-800 hover:bg-slate-900 text-white font-bold px-4 h-12">
                    إضافة
                  </Button>
                </div>

                {bonus > 0 && (
                  <div className="flex justify-between items-center bg-blue-50 p-3 rounded-xl border border-blue-100">
                    <span className="text-sm font-bold text-blue-700">الزيادة المضافة</span>
                    <span className="text-lg font-black text-blue-700">+{bonus.toLocaleString()} ر.ي</span>
                  </div>
                )}

                <Button onClick={handleConfirmPayment} className="w-full h-14 rounded-2xl bg-emerald-500 hover:bg-emerald-600 text-white text-lg font-black shadow-lg shadow-emerald-500/20 mt-2">
                  <CheckCircle2 className="w-6 h-6 ml-2" />
                  تأكيد صرف {(selectedPerson.person.baseAmount + bonus).toLocaleString()} ر.ي
                </Button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Success Overlay */}
      <AnimatePresence>
        {showSuccess && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-emerald-500 z-[60] flex flex-col items-center justify-center text-white"
          >
            <motion.div
              initial={{ scale: 0.5, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              transition={{ type: "spring", bounce: 0.5 }}
              className="w-32 h-32 bg-white/20 rounded-full flex items-center justify-center mb-6"
            >
              <CheckCircle2 className="w-16 h-16" />
            </motion.div>
            <h2 className="text-3xl font-black mb-2">تم الصرف بنجاح!</h2>
            <div className="text-5xl font-black tracking-tight mb-2">
              {((selectedPerson?.person.baseAmount || BASE_AMOUNT) + bonus).toLocaleString()}
            </div>
            <div className="text-xl font-bold opacity-90">ريال يمني</div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Settings Modal */}
      <AnimatePresence>
        {showSettingsModal && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4"
            onClick={() => setShowSettingsModal(false)}
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="bg-white rounded-3xl w-full max-w-sm overflow-hidden shadow-2xl"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="p-6 text-center border-b border-slate-100 relative">
                <button 
                  onClick={() => setShowSettingsModal(false)}
                  className="absolute left-4 top-4 w-8 h-8 bg-slate-100 rounded-full flex items-center justify-center text-slate-500 hover:bg-slate-200"
                >
                  <X className="w-4 h-4" />
                </button>
                <div className="w-16 h-16 bg-slate-100 rounded-2xl mx-auto mb-4 flex items-center justify-center text-slate-600">
                  <Settings className="w-8 h-8" />
                </div>
                <h2 className="text-xl font-black text-slate-900">الإعدادات</h2>
              </div>

              <div className="p-6 space-y-4">
                {/* Actions */}
                <div className="space-y-2">
                  <Button variant="outline" className="w-full justify-start h-12 font-bold" onClick={() => { exportToExcel(); setShowSettingsModal(false); }}>
                    <FileSpreadsheet className="w-5 h-5 ml-2 text-emerald-600" />
                    تصدير التقرير الشامل
                  </Button>
                  <Button variant="outline" className="w-full justify-start h-12 font-bold" onClick={() => {
                    setNewCovenantInput(state.covenant.toString());
                    setShowSettingsModal(false);
                    setShowCovenantModal(true);
                  }}>
                    <Briefcase className="w-5 h-5 ml-2 text-purple-600" />
                    تعديل إجمالي العهد
                  </Button>
                  <Button variant="outline" className="w-full justify-start h-12 font-bold" onClick={() => {
                    setShowSettingsModal(false);
                    setShowImportModal(true);
                  }}>
                    <Upload className="w-5 h-5 ml-2 text-blue-600" />
                    استيراد الأفراد من Excel
                  </Button>
                  <Button variant="outline" className="w-full justify-start h-12 font-bold text-rose-600 hover:text-rose-700 hover:bg-rose-50" onClick={() => {
                    setShowSettingsModal(false);
                    setShowResetModal(true);
                  }}>
                    <RotateCcw className="w-5 h-5 ml-2" />
                    بدء يوم جديد (إعادة تعيين)
                  </Button>
                </div>

                {/* About App */}
                <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 mt-6 text-center">
                  <h3 className="font-black text-slate-800 mb-2 flex items-center justify-center gap-1">
                    <Info className="w-4 h-4 text-blue-600" />
                    حول التطبيق
                  </h3>
                  <p className="text-xs font-semibold text-slate-500 mb-4 leading-relaxed">
                    نظام مالي متكامل لإدارة العهد ومتابعة عمليات الصرف للأفراد والأقسام بدقة وسهولة.
                  </p>
                  <div className="inline-block bg-blue-50 text-blue-700 px-3 py-1.5 rounded-lg text-xs font-black border border-blue-100">
                    اعداد وتطوير درويش رياض
                  </div>
                </div>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Import Modal */}
      <AnimatePresence>
        {showImportModal && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] flex items-center justify-center p-4"
            onClick={() => {
              setShowImportModal(false);
              setImportFile(null);
            }}
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="bg-white rounded-3xl w-full max-w-sm overflow-hidden shadow-2xl"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="p-6 text-center border-b border-slate-100 relative">
                <button 
                  onClick={() => {
                    setShowImportModal(false);
                    setImportFile(null);
                  }}
                  className="absolute left-4 top-4 w-8 h-8 bg-slate-100 rounded-full flex items-center justify-center text-slate-500 hover:bg-slate-200"
                >
                  <X className="w-4 h-4" />
                </button>
                <div className="w-16 h-16 bg-blue-50 rounded-2xl mx-auto mb-4 flex items-center justify-center text-blue-600">
                  <Upload className="w-8 h-8" />
                </div>
                <h2 className="text-xl font-black text-slate-900">استيراد الأفراد</h2>
                <p className="text-sm text-slate-500 mt-1">اختر ملف Excel (.xlsx, .xls)</p>
              </div>

              <div className="p-6 space-y-4">
                {!importFile ? (
                  <div className="relative border-2 border-dashed border-slate-200 rounded-2xl p-8 text-center hover:border-blue-400 hover:bg-blue-50 transition-colors cursor-pointer">
                    <input
                      type="file"
                      accept=".xlsx, .xls"
                      onChange={(e) => {
                        const file = e.target.files?.[0];
                        if (file) setImportFile(file);
                      }}
                      className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                    />
                    <FileSpreadsheet className="w-10 h-10 text-slate-400 mx-auto mb-3" />
                    <p className="text-sm font-bold text-slate-600">اضغط لاختيار ملف</p>
                  </div>
                ) : (
                  <div className="bg-blue-50 border border-blue-100 rounded-2xl p-4 flex items-center gap-3">
                    <div className="w-10 h-10 bg-blue-100 rounded-xl flex items-center justify-center text-blue-600 shrink-0">
                      <FileSpreadsheet className="w-5 h-5" />
                    </div>
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-bold text-slate-800 truncate">{importFile.name}</p>
                      <p className="text-xs text-slate-500">{(importFile.size / 1024).toFixed(1)} KB</p>
                    </div>
                    <button 
                      onClick={() => setImportFile(null)}
                      className="w-8 h-8 bg-white rounded-full flex items-center justify-center text-rose-500 shadow-sm shrink-0"
                    >
                      <X className="w-4 h-4" />
                    </button>
                  </div>
                )}

                <div className="pt-2">
                  <Button 
                    onClick={processImportFile} 
                    disabled={!importFile}
                    className="w-full h-14 rounded-2xl bg-blue-600 hover:bg-blue-700 text-white text-base shadow-lg shadow-blue-500/20 disabled:opacity-50 disabled:shadow-none"
                  >
                    بدء الاستيراد
                  </Button>
                </div>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Covenant Modal */}
      <AnimatePresence>
        {showCovenantModal && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[70] flex items-center justify-center p-4"
            onClick={() => setShowCovenantModal(false)}
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="bg-white rounded-3xl w-full max-w-sm overflow-hidden shadow-2xl p-6 text-center"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="w-16 h-16 bg-purple-100 rounded-full mx-auto mb-4 flex items-center justify-center text-purple-600">
                <Briefcase className="w-8 h-8" />
              </div>
              <h3 className="text-xl font-black text-slate-900 mb-2">تعديل العهد</h3>
              <p className="text-slate-500 mb-6 font-semibold">أدخل مبلغ العهد الجديد بالريال اليمني</p>
              
              <Input
                type="number"
                value={newCovenantInput}
                onChange={(e) => setNewCovenantInput(e.target.value)}
                className="mb-6 text-center font-bold text-lg h-14 bg-slate-50 border-slate-200"
                autoFocus
              />

              <div className="flex gap-3">
                <Button
                  onClick={() => setShowCovenantModal(false)}
                  className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold h-12 rounded-xl"
                >
                  إلغاء
                </Button>
                <Button
                  onClick={() => {
                    const amount = parseInt(newCovenantInput);
                    if (!isNaN(amount) && amount >= 0) {
                      updateCovenant(amount);
                      showToast("✓ تم تحديث العهد بنجاح");
                      setShowCovenantModal(false);
                    } else {
                      showToast("❌ يرجى إدخال مبلغ صحيح");
                    }
                  }}
                  className="flex-1 bg-purple-600 hover:bg-purple-700 text-white font-bold h-12 rounded-xl shadow-lg shadow-purple-500/20"
                >
                  حفظ
                </Button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Reset Day Modal */}
      <AnimatePresence>
        {showResetModal && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[70] flex items-center justify-center p-4"
            onClick={() => setShowResetModal(false)}
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="bg-white rounded-3xl w-full max-w-sm overflow-hidden shadow-2xl p-6 text-center"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="w-16 h-16 bg-rose-100 rounded-full mx-auto mb-4 flex items-center justify-center text-rose-600">
                <RotateCcw className="w-8 h-8" />
              </div>
              <h3 className="text-xl font-black text-slate-900 mb-2">بدء يوم جديد</h3>
              <p className="text-slate-500 mb-6 font-semibold leading-relaxed">
                هل أنت متأكد من بدء يوم جديد؟<br/>
                <span className="text-rose-600 font-bold">سيتم تصفير جميع حالات الاستلام والمصروفات.</span>
              </p>
              
              <div className="flex gap-3">
                <Button
                  onClick={() => setShowResetModal(false)}
                  className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold h-12 rounded-xl"
                >
                  تراجع
                </Button>
                <Button
                  onClick={() => {
                    resetDay();
                    showToast("🌅 تم بدء يوم جديد بنجاح");
                    setShowResetModal(false);
                  }}
                  className="flex-1 bg-rose-600 hover:bg-rose-700 text-white font-bold h-12 rounded-xl shadow-lg shadow-rose-500/20"
                >
                  تأكيد
                </Button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Toast */}
      <AnimatePresence>
        {toast.visible && (
          <motion.div
            initial={{ opacity: 0, y: -50, x: "-50%" }}
            animate={{ opacity: 1, y: 0, x: "-50%" }}
            exit={{ opacity: 0, y: -20, x: "-50%" }}
            className="fixed top-6 left-1/2 z-[70] bg-slate-900 text-white px-6 py-3 rounded-full font-bold shadow-xl flex items-center gap-2"
          >
            {toast.message}
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

function NavItem({ icon, label, active, onClick }: { icon: React.ReactNode; label: string; active: boolean; onClick: () => void }) {
  return (
    <button
      onClick={onClick}
      className={`flex flex-col items-center gap-1 p-2 w-20 transition-colors ${
        active ? "text-blue-600" : "text-slate-400 hover:text-slate-600"
      }`}
    >
      <motion.div
        animate={{ y: active ? -4 : 0 }}
        className={`p-2 rounded-2xl ${active ? "bg-blue-50" : ""}`}
      >
        {icon}
      </motion.div>
      <span className={`text-[10px] font-bold ${active ? "opacity-100" : "opacity-0"}`}>
        {label}
      </span>
    </button>
  );
}
