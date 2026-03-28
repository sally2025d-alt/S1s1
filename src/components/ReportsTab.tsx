import { Card, CardContent, CardHeader, CardTitle } from "./ui/Card";
import { Department, Expense, BASE_AMOUNT } from "../types";
import { formatCurrency } from "../lib/utils";
import { FileSpreadsheet, Clock, BarChart3, Users, DollarSign, Building2, Download, Printer } from "lucide-react";
import { Button } from "./ui/Button";
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { useRef, useState } from "react";
import { useReactToPrint } from "react-to-print";
import { PrintReport } from "./PrintReport";
import html2pdf from "html2pdf.js";

interface ReportsTabProps {
  covenant: number;
  stats: {
    totalSpent: number;
    remaining: number;
    totalPersons: number;
    receivedPersons: number;
    pendingPersons: number;
  };
  departments: Department[];
  expenses: Expense[];
}

export function ReportsTab({ covenant, stats, departments, expenses }: ReportsTabProps) {
  const percentage = covenant > 0 ? Math.round((stats.totalSpent / covenant) * 100) : 0;
  const average = stats.receivedPersons > 0 ? Math.round(stats.totalSpent / stats.receivedPersons) : 0;
  const [isDownloading, setIsDownloading] = useState(false);

  let maxAmount = 0;
  let minAmount = Infinity;
  const transactions: any[] = [];

  departments.forEach(d => {
    d.persons.forEach(p => {
      if (p.received) {
        if (p.totalAmount > maxAmount) maxAmount = p.totalAmount;
        if (p.totalAmount < minAmount) minAmount = p.totalAmount;
        transactions.push({
          name: p.name,
          rank: p.rank,
          amount: p.totalAmount,
          date: p.date,
          time: p.time,
          type: 'person'
        });
      }
    });
  });

  expenses.forEach(e => {
    if (e.amount > maxAmount) maxAmount = e.amount;
    if (e.amount < minAmount) minAmount = e.amount;
    transactions.push({
      name: e.recipient,
      rank: 'أمر صرف',
      amount: e.amount,
      date: e.date,
      time: e.time,
      type: 'expense'
    });
  });

  if (minAmount === Infinity) minAmount = 0;

  transactions.sort((a, b) => {
    const dateA = new Date(a.date + " " + a.time).getTime();
    const dateB = new Date(b.date + " " + b.time).getTime();
    return dateB - dateA;
  });

  const recentTransactions = transactions.slice(0, 5);

  const printRef = useRef<HTMLDivElement>(null);
  
  const handlePrint = useReactToPrint({
    contentRef: printRef,
    documentTitle: `التقرير_المالي_العام_${new Date().toLocaleDateString("ar-SA").replace(/\//g, "-")}`,
  });

  const handleDownloadPdf = async () => {
    if (!printRef.current) return;
    setIsDownloading(true);
    try {
      const container = document.getElementById('pdf-container');
      if (container) {
        container.classList.remove('hidden');
        container.style.position = 'fixed';
        container.style.top = '0';
        container.style.left = '0';
        container.style.width = '210mm';
        container.style.zIndex = '-9999';
        container.style.backgroundColor = 'white';
      }

      const element = printRef.current;
      const date = new Date().toLocaleDateString("ar-SA").replace(/\//g, "-");
      
      const opt = {
        margin:       10,
        filename:     `التقرير_المالي_العام_${date}.pdf`,
        image:        { type: 'jpeg' as const, quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true, logging: false },
        jsPDF:        { unit: 'mm' as const, format: 'a4' as const, orientation: 'portrait' as const },
        pagebreak:    { mode: ['css', 'legacy'] }
      };

      await html2pdf().set(opt).from(element).save();

      if (container) {
        container.classList.add('hidden');
        container.style.position = '';
        container.style.top = '';
        container.style.left = '';
        container.style.width = '';
        container.style.zIndex = '';
        container.style.backgroundColor = '';
      }
    } catch (error) {
      console.error("Error generating PDF:", error);
    } finally {
      setIsDownloading(false);
    }
  };

  const exportDepartmentsToExcel = async () => {
    const data = departments.map(dept => {
      const total = dept.persons.length;
      const received = dept.persons.filter(p => p.received).length;
      const remaining = total - received;
      const percent = total > 0 ? Math.round((received / total) * 100) : 0;

      return {
        dept: dept.name,
        total: total,
        received: received,
        remaining: remaining,
        percent: `${percent}%`
      };
    });

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'صراف اللواء';
    workbook.created = new Date();

    const ws = workbook.addWorksheet("تقرير الأقسام", { views: [{ rightToLeft: true }] });
    
    ws.columns = [
      { header: 'القسم', key: 'dept', width: 25 },
      { header: 'إجمالي الأفراد', key: 'total', width: 20 },
      { header: 'تم الصرف لهم', key: 'received', width: 20 },
      { header: 'المتبقي', key: 'remaining', width: 20 },
      { header: 'نسبة الصرف', key: 'percent', width: 15 }
    ];

    // Style Header
    const headerRow = ws.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Arial' };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } }; // Indigo 600
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow.height = 30;

    // Add Data
    data.forEach((item, index) => {
      const row = ws.addRow(item);
      row.alignment = { vertical: 'middle', horizontal: 'center' };
      row.font = { size: 11, name: 'Arial', bold: true };
      row.height = 25;
      
      // Alternating row colors
      if (index % 2 === 0) {
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } }; // Slate 50
      } else {
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } }; // White
      }

      // Highlight specific columns
      row.getCell('received').font = { size: 11, name: 'Arial', bold: true, color: { argb: 'FF16A34A' } }; // Green
      row.getCell('remaining').font = { size: 11, name: 'Arial', bold: true, color: { argb: 'FFDC2626' } }; // Red
      
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
        top: {style:'thin', color: {argb:'FFC7D2FE'}},
        left: {style:'thin', color: {argb:'FFC7D2FE'}},
        bottom: {style:'thin', color: {argb:'FFC7D2FE'}},
        right: {style:'thin', color: {argb:'FFC7D2FE'}}
      };
    });

    const date = new Date().toLocaleDateString("ar-SA").replace(/\//g, "-");
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `تقرير_الأقسام_${date}.xlsx`);
  };

  return (
    <div className="space-y-4 pb-24">
      {/* 1. التقرير المالي العام */}
      <Card className="border-none shadow-sm">
        <CardHeader className="pb-2 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
          <CardTitle className="flex items-center gap-2 text-lg">
            <BarChart3 className="w-5 h-5 text-blue-600" />
            التقرير المالي العام
          </CardTitle>
          <div className="flex flex-wrap gap-2 w-full sm:w-auto">
            <Button 
              onClick={handleDownloadPdf} 
              disabled={isDownloading}
              className="bg-emerald-600 hover:bg-emerald-700 text-white gap-2 rounded-xl flex-1 sm:flex-none"
              size="sm"
            >
              <Download className="w-4 h-4" />
              {isDownloading ? "جاري التنزيل..." : "تنزيل PDF"}
            </Button>
            <Button 
              onClick={() => handlePrint()} 
              className="bg-slate-900 hover:bg-slate-800 text-white gap-2 rounded-xl flex-1 sm:flex-none"
              size="sm"
            >
              <Printer className="w-4 h-4" />
              طباعة
            </Button>
          </div>
        </CardHeader>
        <CardContent className="space-y-2">
          <ReportRow label="إجمالي العهد" value={`${covenant.toLocaleString()} ر.ي`} />
          <ReportRow label="إجمالي المصروفات" value={`${stats.totalSpent.toLocaleString()} ر.ي`} />
          <ReportRow label="المبلغ المتبقي" value={`${stats.remaining.toLocaleString()} ر.ي`} />
          <ReportRow
            label="نسبة الصرف"
            value={`${percentage}%`}
            isHighlight
          />
        </CardContent>
      </Card>

      {/* 2. تقرير حالة الصرف للأفراد */}
      <Card className="border-none shadow-sm">
        <CardHeader className="pb-2">
          <CardTitle className="flex items-center gap-2 text-lg">
            <Users className="w-5 h-5 text-purple-600" />
            تقرير حالة الصرف للأفراد
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-2">
          <ReportRow label="إجمالي عدد الأفراد" value={stats.totalPersons.toString()} />
          <ReportRow label="تم الصرف لهم" value={stats.receivedPersons.toString()} />
          <ReportRow label="عدد المتبقين" value={stats.pendingPersons.toString()} />
        </CardContent>
      </Card>

      {/* 3. تقرير تحليل المبالغ */}
      <Card className="border-none shadow-sm">
        <CardHeader className="pb-2">
          <CardTitle className="flex items-center gap-2 text-lg">
            <DollarSign className="w-5 h-5 text-emerald-600" />
            تقرير تحليل المبالغ
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-2">
          <ReportRow label="أعلى مبلغ مصروف" value={`${maxAmount.toLocaleString()} ر.ي`} />
          <ReportRow label="أقل مبلغ مصروف" value={`${minAmount.toLocaleString()} ر.ي`} />
          <ReportRow label="متوسط الصرف" value={`${average.toLocaleString()} ر.ي`} />
        </CardContent>
      </Card>

      {/* 4. تقرير الصرف حسب الأقسام */}
      <Card className="border-none shadow-sm">
        <CardHeader className="pb-2 flex flex-row items-center justify-between">
          <CardTitle className="flex items-center gap-2 text-lg">
            <Building2 className="w-5 h-5 text-indigo-600" />
            تقرير الصرف حسب الأقسام
          </CardTitle>
          <Button variant="outline" size="sm" onClick={exportDepartmentsToExcel} className="h-8 gap-1 rounded-lg text-xs font-bold">
            <Download className="w-3.5 h-3.5" />
            تصدير
          </Button>
        </CardHeader>
        <CardContent className="space-y-3">
          {departments.map(dept => {
            const total = dept.persons.length;
            const received = dept.persons.filter(p => p.received).length;
            const percent = total > 0 ? Math.round((received / total) * 100) : 0;
            
            return (
              <div key={dept.id} className="bg-slate-50 p-4 rounded-xl border border-slate-100">
                <div className="flex justify-between items-center mb-2">
                  <span className="font-bold text-slate-800">{dept.name}</span>
                  <span className="font-black text-indigo-600">{percent}%</span>
                </div>
                <div className="w-full bg-slate-200 rounded-full h-2 mb-3 overflow-hidden">
                  <div className="bg-indigo-600 h-2 rounded-full transition-all duration-500" style={{ width: `${percent}%` }}></div>
                </div>
                <div className="flex justify-between text-xs font-bold text-slate-500">
                  <span className="bg-white px-2 py-1 rounded-md border border-slate-100 shadow-sm">الإجمالي: {total}</span>
                  <span className="bg-emerald-50 text-emerald-700 px-2 py-1 rounded-md border border-emerald-100 shadow-sm">استلم: {received}</span>
                  <span className="bg-rose-50 text-rose-700 px-2 py-1 rounded-md border border-rose-100 shadow-sm">متبقي: {total - received}</span>
                </div>
              </div>
            );
          })}
        </CardContent>
      </Card>

      {/* 5. تقرير آخر العمليات */}
      <Card className="border-none shadow-sm">
        <CardHeader className="pb-2">
          <CardTitle className="flex items-center gap-2 text-lg">
            <Clock className="w-5 h-5 text-amber-600" />
            تقرير آخر العمليات
          </CardTitle>
        </CardHeader>
        <CardContent>
          {recentTransactions.length === 0 ? (
            <div className="text-center py-6 text-slate-400 font-semibold">
              لا توجد عمليات
            </div>
          ) : (
            <div className="space-y-3">
              {recentTransactions.map((t, i) => (
                <div key={i} className="flex items-center p-3 bg-slate-50 rounded-xl border border-slate-100">
                  <div className={`w-12 h-12 rounded-xl flex items-center justify-center ml-3 shrink-0 ${t.type === 'expense' ? 'bg-rose-100 text-rose-600' : 'bg-emerald-100 text-emerald-600'}`}>
                    <DollarSign className="w-6 h-6" />
                  </div>
                  <div className="flex-1">
                    <div className="font-bold text-slate-800 mb-0.5">
                      {t.rank} {t.name}
                    </div>
                    <div className="text-xs font-semibold text-slate-500">
                      {t.date} • {t.time}
                    </div>
                  </div>
                  <div className={`font-black text-lg ${t.type === 'expense' ? 'text-rose-600' : 'text-emerald-600'}`}>
                    {t.amount.toLocaleString()} ر.ي
                  </div>
                </div>
              ))}
            </div>
          )}
        </CardContent>
      </Card>

      {/* Hidden Print Component */}
      <div id="pdf-container" className="hidden">
        <PrintReport 
          ref={printRef} 
          covenant={covenant} 
          stats={stats} 
          departments={departments} 
          expenses={expenses}
        />
      </div>
    </div>
  );
}

function ReportRow({ label, value, isHighlight = false }: { label: string; value: string; isHighlight?: boolean }) {
  return (
    <div
      className={`flex justify-between items-center p-4 rounded-xl ${
        isHighlight ? "bg-slate-900 text-white" : "bg-slate-50 text-slate-800"
      }`}
    >
      <span className="font-bold text-sm">{label}</span>
      <span className={`font-black ${isHighlight ? "text-xl" : "text-lg"}`}>{value}</span>
    </div>
  );
}
