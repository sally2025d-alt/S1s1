import React, { useState } from 'react';
import { ArchiveRecord } from '../types';
import { motion, AnimatePresence } from 'motion/react';
import { Calendar, ChevronDown, ChevronUp, FileText, Users, DollarSign, Download, Trash2, Search, Building2, AlertTriangle } from 'lucide-react';
import { cn } from '../lib/utils';
import { Input } from './ui/Input';
import { Button } from './ui/Button';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

interface ArchiveTabProps {
  archives: ArchiveRecord[];
  onDeleteArchive: (id: string) => void;
  onClearArchives: () => void;
}

export function ArchiveTab({ archives, onDeleteArchive, onClearArchives }: ArchiveTabProps) {
  const [expandedId, setExpandedId] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [showClearConfirm, setShowClearConfirm] = useState(false);

  const filteredArchives = archives.filter(a => a.date.includes(searchQuery));

  const exportArchiveToExcel = async (archive: ArchiveRecord) => {
    const receivedData: any[] = [];
    const pendingData: any[] = [];
    const expensesData: any[] = [];
    let totalPaid = 0;

    archive.departments.forEach(d => {
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

    archive.expenses.forEach(e => {
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

      const headerRow = ws.getRow(1);
      headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Arial' };
      headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
      headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
      headerRow.height = 30;

      data.forEach((item, index) => {
        const row = ws.addRow(item);
        row.alignment = { vertical: 'middle', horizontal: 'center' };
        row.font = { size: 11, name: 'Arial', bold: type === 'summary' };
        row.height = 25;
        
        if (index % 2 === 0) {
          row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        } else {
          row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
        }

        if (type === 'summary' && item.label === 'المتبقي من العهد') {
          row.font = { size: 12, name: 'Arial', bold: true, color: { argb: 'FF16A34A' } };
        }
        if (type === 'summary' && item.label === 'إجمالي المصروف') {
          row.font = { size: 12, name: 'Arial', bold: true, color: { argb: 'FFDC2626' } };
        }
        
        row.eachCell((cell) => {
          cell.border = {
            top: {style:'thin', color: {argb:'FFE2E8F0'}},
            left: {style:'thin', color: {argb:'FFE2E8F0'}},
            bottom: {style:'thin', color: {argb:'FFE2E8F0'}},
            right: {style:'thin', color: {argb:'FFE2E8F0'}}
          };
        });
      });

      headerRow.eachCell((cell) => {
        cell.border = {
          top: {style:'thin', color: {argb:'FF94A3B8'}},
          left: {style:'thin', color: {argb:'FF94A3B8'}},
          bottom: {style:'thin', color: {argb:'FF94A3B8'}},
          right: {style:'thin', color: {argb:'FF94A3B8'}}
        };
      });
    };

    createStyledSheet("المستلمين", receivedData, 'persons');
    createStyledSheet("غير المستلمين", pendingData, 'persons');
    createStyledSheet("أوامر الصرف", expensesData, 'expenses');

    const summaryData = [
      { label: "إجمالي العهد", value: archive.covenant },
      { label: "إجمالي المصروف", value: totalPaid },
      { label: "المتبقي من العهد", value: archive.covenant - totalPaid },
      { label: "إجمالي الأفراد", value: archive.stats.totalPersons },
      { label: "عدد المستلمين", value: archive.stats.receivedPersons },
      { label: "عدد المتبقين", value: archive.stats.pendingPersons },
      { label: "عدد أوامر الصرف", value: archive.expenses.length },
    ];
    createStyledSheet("الملخص المالي", summaryData, 'summary');

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `أرشيف_صرف_اللواء_${archive.date.replace(/\//g, "-")}.xlsx`);
  };

  if (!archives || archives.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center py-20 text-slate-400">
        <div className="w-24 h-24 bg-slate-100 rounded-full flex items-center justify-center mb-4">
          <Calendar className="w-10 h-10 text-slate-300" />
        </div>
        <h2 className="text-xl font-bold text-slate-600 mb-2">لا يوجد أرشيف بعد</h2>
        <p className="text-sm text-center max-w-xs leading-relaxed">
          عند بدء يوم جديد، سيتم حفظ بيانات اليوم السابق هنا للرجوع إليها لاحقاً.
        </p>
      </div>
    );
  }

  return (
    <div className="space-y-4 pb-24">
      <div className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 mb-6 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
        <div>
          <h2 className="text-lg font-black text-slate-800 flex items-center gap-2">
            <Calendar className="w-5 h-5 text-blue-600" />
            السجل التاريخي
          </h2>
          <p className="text-sm text-slate-500 mt-1 font-semibold">
            يحتوي على {archives.length} سجلات سابقة
          </p>
        </div>
        <Button 
          variant="outline" 
          onClick={() => setShowClearConfirm(true)}
          className="text-rose-600 border-rose-200 hover:bg-rose-50 hover:border-rose-300 font-bold"
        >
          <Trash2 className="w-4 h-4 ml-2" />
          مسح الأرشيف
        </Button>
      </div>

      {/* Search */}
      <div className="relative">
        <Search className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400 w-5 h-5" />
        <Input
          placeholder="ابحث بالتاريخ (مثال: 2024/5/12)..."
          value={searchQuery}
          onChange={(e) => setSearchQuery(e.target.value)}
          className="pl-4 pr-12 h-14 bg-white border-slate-200 shadow-sm rounded-2xl text-base font-semibold"
        />
      </div>

      <div className="space-y-4 mt-4">
        {filteredArchives.length === 0 ? (
          <div className="text-center py-10 text-slate-400 font-bold">
            لا توجد سجلات مطابقة لبحثك
          </div>
        ) : (
          filteredArchives.map((archive) => (
            <div key={archive.id} className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-sm">
              <div 
                className="p-4 flex items-center justify-between cursor-pointer hover:bg-slate-50 transition-colors"
                onClick={() => setExpandedId(expandedId === archive.id ? null : archive.id)}
              >
                <div className="flex items-center gap-3">
                  <div className="w-12 h-12 bg-blue-50 rounded-xl flex items-center justify-center text-blue-600 shrink-0">
                    <FileText className="w-6 h-6" />
                  </div>
                  <div>
                    <h3 className="font-bold text-slate-900 text-lg">{archive.date}</h3>
                    <div className="text-sm text-slate-500 font-semibold">
                      المصروف: {archive.stats.totalSpent.toLocaleString()} ر.ي
                    </div>
                  </div>
                </div>
                <div className="w-8 h-8 rounded-full bg-slate-100 flex items-center justify-center text-slate-500 shrink-0">
                  {expandedId === archive.id ? <ChevronUp className="w-5 h-5" /> : <ChevronDown className="w-5 h-5" />}
                </div>
              </div>

              <AnimatePresence>
                {expandedId === archive.id && (
                  <motion.div
                    initial={{ height: 0, opacity: 0 }}
                    animate={{ height: 'auto', opacity: 1 }}
                    exit={{ height: 0, opacity: 0 }}
                    className="overflow-hidden border-t border-slate-100 bg-slate-50"
                  >
                    <div className="p-4 space-y-6">
                      {/* Actions */}
                      <div className="flex gap-2">
                        <Button 
                          onClick={() => exportArchiveToExcel(archive)}
                          className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold h-10 rounded-xl shadow-sm"
                        >
                          <Download className="w-4 h-4 ml-2" />
                          تصدير التقرير (Excel)
                        </Button>
                        <Button 
                          onClick={() => {
                            if (window.confirm('هل أنت متأكد من حذف هذا السجل؟ لا يمكن التراجع عن هذا الإجراء.')) {
                              onDeleteArchive(archive.id);
                            }
                          }}
                          variant="outline"
                          className="text-rose-600 border-rose-200 hover:bg-rose-50 h-10 rounded-xl"
                        >
                          <Trash2 className="w-4 h-4" />
                        </Button>
                      </div>

                      {/* Stats Grid */}
                      <div className="grid grid-cols-2 gap-3">
                        <div className="bg-white p-3 rounded-xl border border-slate-200">
                          <div className="text-xs text-slate-500 font-bold mb-1 flex items-center gap-1">
                            <DollarSign className="w-3 h-3" /> إجمالي العهد
                          </div>
                          <div className="font-black text-slate-800">{archive.covenant.toLocaleString()}</div>
                        </div>
                        <div className="bg-white p-3 rounded-xl border border-slate-200">
                          <div className="text-xs text-emerald-600 font-bold mb-1 flex items-center gap-1">
                            <DollarSign className="w-3 h-3" /> المتبقي
                          </div>
                          <div className="font-black text-emerald-700">{archive.stats.remaining.toLocaleString()}</div>
                        </div>
                        <div className="bg-white p-3 rounded-xl border border-slate-200">
                          <div className="text-xs text-blue-600 font-bold mb-1 flex items-center gap-1">
                            <Users className="w-3 h-3" /> المستلمين
                          </div>
                          <div className="font-black text-blue-700">{archive.stats.receivedPersons} / {archive.stats.totalPersons}</div>
                        </div>
                        <div className="bg-white p-3 rounded-xl border border-slate-200">
                          <div className="text-xs text-rose-600 font-bold mb-1 flex items-center gap-1">
                            <FileText className="w-3 h-3" /> أوامر الصرف
                          </div>
                          <div className="font-black text-rose-700">{archive.expenses.length}</div>
                        </div>
                      </div>

                      {/* Departments Breakdown */}
                      <div>
                        <h4 className="font-bold text-slate-800 mb-3 text-sm border-b border-slate-200 pb-2 flex items-center gap-2">
                          <Building2 className="w-4 h-4 text-slate-500" />
                          تفاصيل الأقسام
                        </h4>
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                          {archive.departments.map(dept => {
                            const received = dept.persons.filter(p => p.received).length;
                            const total = dept.persons.length;
                            const percent = total > 0 ? Math.round((received / total) * 100) : 0;
                            return (
                              <div key={dept.id} className="bg-white p-3 rounded-xl border border-slate-200 flex justify-between items-center">
                                <span className="font-bold text-sm text-slate-700">{dept.name}</span>
                                <div className="flex items-center gap-2">
                                  <span className="text-xs font-semibold text-slate-500">{received}/{total}</span>
                                  <span className={cn(
                                    "text-xs font-black px-2 py-0.5 rounded-md",
                                    percent === 100 ? "bg-emerald-100 text-emerald-700" : "bg-slate-100 text-slate-600"
                                  )}>{percent}%</span>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>

                      {/* Expenses List */}
                      {archive.expenses && archive.expenses.length > 0 && (
                        <div>
                          <h4 className="font-bold text-slate-800 mb-3 text-sm border-b border-slate-200 pb-2 flex items-center gap-2">
                            <FileText className="w-4 h-4 text-slate-500" />
                            أوامر الصرف
                          </h4>
                          <div className="space-y-2">
                            {archive.expenses.map(exp => (
                              <div key={exp.id} className="bg-white p-3 rounded-xl border border-slate-200 flex justify-between items-center">
                                <div>
                                  <div className="font-bold text-sm text-slate-800">{exp.recipient}</div>
                                  <div className="text-xs text-slate-500 font-semibold">{exp.purpose}</div>
                                </div>
                                <div className="font-black text-rose-600 text-sm">
                                  {exp.amount.toLocaleString()} ر.ي
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          ))
        )}
      </div>

      {/* Clear All Confirmation Modal */}
      <AnimatePresence>
        {showClearConfirm && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[70] flex items-center justify-center p-4"
            onClick={() => setShowClearConfirm(false)}
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="bg-white rounded-3xl w-full max-w-sm overflow-hidden shadow-2xl p-6 text-center"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="w-16 h-16 bg-rose-100 rounded-full mx-auto mb-4 flex items-center justify-center text-rose-600">
                <AlertTriangle className="w-8 h-8" />
              </div>
              <h3 className="text-xl font-black text-slate-900 mb-2">مسح الأرشيف بالكامل</h3>
              <p className="text-slate-500 mb-6 font-semibold leading-relaxed">
                هل أنت متأكد من مسح جميع السجلات التاريخية؟<br/>
                <span className="text-rose-600 font-bold">هذا الإجراء نهائي ولا يمكن التراجع عنه.</span>
              </p>
              
              <div className="flex gap-3">
                <Button
                  onClick={() => setShowClearConfirm(false)}
                  className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold h-12 rounded-xl"
                >
                  تراجع
                </Button>
                <Button
                  onClick={() => {
                    onClearArchives();
                    setShowClearConfirm(false);
                  }}
                  className="flex-1 bg-rose-600 hover:bg-rose-700 text-white font-bold h-12 rounded-xl shadow-lg shadow-rose-500/20"
                >
                  نعم، امسح الكل
                </Button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
