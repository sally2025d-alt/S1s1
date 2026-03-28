import React from 'react';
import { Department, Expense } from '../types';

interface PrintReportProps {
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

export const PrintReport = React.forwardRef<HTMLDivElement, PrintReportProps>(
  ({ covenant, stats, departments, expenses }, ref) => {
    const today = new Date().toLocaleDateString('ar-SA');

    return (
      <div ref={ref} className="p-8 print-safe-colors" dir="rtl" style={{ fontFamily: 'Arial, sans-serif', backgroundColor: '#ffffff', color: '#000000' }}>
        {/* Page 1: General Financial Report */}
        <div className="page-break-after-always html2pdf__page-break min-h-[297mm] flex flex-col">
          <div className="flex justify-between items-start mb-10 border-b-4 pb-6" style={{ borderColor: '#000000' }}>
            <div className="text-right">
              <h3 className="font-bold text-lg">الجمهورية اليمنية</h3>
              <h3 className="font-bold text-lg">وزارة الدفاع</h3>
              <h3 className="font-bold text-lg">الشؤون المالية</h3>
            </div>
            <div className="text-center">
              <h1 className="text-3xl font-black mb-2 underline decoration-double underline-offset-8">التقرير المالي العام</h1>
              <p className="text-lg font-bold mt-4" style={{ color: '#4b5563' }}>تاريخ التقرير: {today}</p>
            </div>
            <div className="text-left opacity-0">
              <h3 className="font-bold text-lg">الجمهورية اليمنية</h3>
              <h3 className="font-bold text-lg">وزارة الدفاع</h3>
              <h3 className="font-bold text-lg">الشؤون المالية</h3>
            </div>
          </div>

          <div className="grid grid-cols-2 gap-8 mb-12">
            <div className="border rounded-lg p-6" style={{ borderColor: '#d1d5db', backgroundColor: '#f9fafb' }}>
              <h2 className="text-2xl font-bold mb-6 border-b pb-2" style={{ color: '#1f2937', borderColor: '#d1d5db' }}>الملخص المالي</h2>
              <div className="space-y-4 text-lg">
                <div className="flex justify-between">
                  <span className="font-semibold">إجمالي العهد:</span>
                  <span className="font-bold">{covenant.toLocaleString()} ر.ي</span>
                </div>
                <div className="flex justify-between" style={{ color: '#dc2626' }}>
                  <span className="font-semibold">إجمالي المصروفات:</span>
                  <span className="font-bold">{stats.totalSpent.toLocaleString()} ر.ي</span>
                </div>
                <div className="flex justify-between" style={{ color: '#16a34a' }}>
                  <span className="font-semibold">المبلغ المتبقي:</span>
                  <span className="font-bold">{stats.remaining.toLocaleString()} ر.ي</span>
                </div>
                <div className="flex justify-between pt-4 border-t" style={{ borderColor: '#d1d5db' }}>
                  <span className="font-semibold">نسبة الصرف:</span>
                  <span className="font-bold">{covenant > 0 ? Math.round((stats.totalSpent / covenant) * 100) : 0}%</span>
                </div>
              </div>
            </div>

            <div className="border rounded-lg p-6" style={{ borderColor: '#d1d5db', backgroundColor: '#f9fafb' }}>
              <h2 className="text-2xl font-bold mb-6 border-b pb-2" style={{ color: '#1f2937', borderColor: '#d1d5db' }}>ملخص الأفراد</h2>
              <div className="space-y-4 text-lg">
                <div className="flex justify-between">
                  <span className="font-semibold">إجمالي الأفراد:</span>
                  <span className="font-bold">{stats.totalPersons}</span>
                </div>
                <div className="flex justify-between" style={{ color: '#16a34a' }}>
                  <span className="font-semibold">تم الصرف لهم:</span>
                  <span className="font-bold">{stats.receivedPersons}</span>
                </div>
                <div className="flex justify-between" style={{ color: '#dc2626' }}>
                  <span className="font-semibold">المتبقي:</span>
                  <span className="font-bold">{stats.pendingPersons}</span>
                </div>
                <div className="flex justify-between pt-4 border-t" style={{ borderColor: '#d1d5db' }}>
                  <span className="font-semibold">نسبة الإنجاز:</span>
                  <span className="font-bold">{stats.totalPersons > 0 ? Math.round((stats.receivedPersons / stats.totalPersons) * 100) : 0}%</span>
                </div>
              </div>
            </div>
          </div>

          <div className="border rounded-lg p-6" style={{ borderColor: '#d1d5db' }}>
            <h2 className="text-2xl font-bold mb-6 border-b pb-2" style={{ color: '#1f2937', borderColor: '#d1d5db' }}>ملخص الأقسام</h2>
            <table className="w-full text-right border-collapse">
              <thead>
                <tr style={{ backgroundColor: '#f3f4f6' }}>
                  <th className="border p-3 font-bold" style={{ borderColor: '#d1d5db' }}>القسم</th>
                  <th className="border p-3 font-bold" style={{ borderColor: '#d1d5db' }}>الإجمالي</th>
                  <th className="border p-3 font-bold" style={{ borderColor: '#d1d5db' }}>تم الصرف</th>
                  <th className="border p-3 font-bold" style={{ borderColor: '#d1d5db' }}>المتبقي</th>
                  <th className="border p-3 font-bold" style={{ borderColor: '#d1d5db' }}>النسبة</th>
                </tr>
              </thead>
              <tbody>
                {departments.map((dept) => {
                  const total = dept.persons.length;
                  const received = dept.persons.filter(p => p.received).length;
                  const remaining = total - received;
                  const percent = total > 0 ? Math.round((received / total) * 100) : 0;
                  
                  return (
                    <tr key={dept.id}>
                      <td className="border p-3 font-semibold" style={{ borderColor: '#d1d5db' }}>{dept.name}</td>
                      <td className="border p-3" style={{ borderColor: '#d1d5db' }}>{total}</td>
                      <td className="border p-3" style={{ color: '#16a34a', borderColor: '#d1d5db' }}>{received}</td>
                      <td className="border p-3" style={{ color: '#dc2626', borderColor: '#d1d5db' }}>{remaining}</td>
                      <td className="border p-3" dir="ltr" style={{ borderColor: '#d1d5db' }}>{percent}%</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div className="mt-auto pt-16 grid grid-cols-2 gap-8 text-center font-bold text-lg">
            <div>
              <p className="mb-12">المسؤول المالي</p>
              <p className="border-t pt-2 w-3/4 mx-auto" style={{ borderColor: '#9ca3af' }}>الاسم والتوقيع</p>
            </div>
            <div>
              <p className="mb-12">الاعتماد</p>
              <p className="border-t pt-2 w-3/4 mx-auto" style={{ borderColor: '#9ca3af' }}>الاسم والتوقيع</p>
            </div>
          </div>
        </div>

        {/* Pages 2+: Departments details */}
        {departments.map((dept, index) => {
          const receivedPersons = dept.persons.filter(p => p.received);
          
          if (receivedPersons.length === 0) return null;

          return (
            <div key={dept.id} className="page-break-before-always html2pdf__page-break min-h-[297mm] pt-8">
              <div className="flex justify-between items-start mb-8 border-b-2 pb-4" style={{ borderColor: '#000000' }}>
                <div className="text-right">
                  <h3 className="font-bold text-lg">الجمهورية اليمنية</h3>
                  <h3 className="font-bold text-lg">وزارة الدفاع</h3>
                  <h3 className="font-bold text-lg">الشؤون المالية</h3>
                </div>
                <div className="text-center">
                  <h2 className="text-2xl font-black mb-2 underline decoration-double underline-offset-8">كشف تسليم المستحقات المالية</h2>
                  <h3 className="text-xl font-bold mt-3">القسم: {dept.name}</h3>
                </div>
                <div className="text-left">
                  <h3 className="font-bold text-lg">التاريخ: {today}</h3>
                  <h3 className="font-bold text-lg mt-1">عدد الأفراد: {receivedPersons.length}</h3>
                </div>
              </div>

              <table className="w-full text-right border-collapse text-base" style={{ borderColor: '#000000' }}>
                <thead>
                  <tr style={{ backgroundColor: '#e5e7eb' }}>
                    <th className="border-2 p-3 font-bold w-12 text-center" style={{ borderColor: '#000000' }}>م</th>
                    <th className="border-2 p-3 font-bold w-32 text-center" style={{ borderColor: '#000000' }}>الرقم العسكري</th>
                    <th className="border-2 p-3 font-bold w-24 text-center" style={{ borderColor: '#000000' }}>الرتبة</th>
                    <th className="border-2 p-3 font-bold" style={{ borderColor: '#000000' }}>الاسم الرباعي</th>
                    <th className="border-2 p-3 font-bold w-40 text-center" style={{ borderColor: '#000000' }}>تاريخ ووقت الاستلام</th>
                    <th className="border-2 p-3 font-bold w-48 text-center" style={{ borderColor: '#000000' }}>التوقيع / البصمة</th>
                  </tr>
                </thead>
                <tbody>
                  {receivedPersons.map((person, idx) => (
                    <tr key={person.id} className="break-inside-avoid">
                      <td className="border-2 p-4 text-center font-bold" style={{ borderColor: '#000000' }}>{idx + 1}</td>
                      <td className="border-2 p-4 text-center font-semibold" style={{ borderColor: '#000000' }}>{person.militaryNumber}</td>
                      <td className="border-2 p-4 text-center font-semibold" style={{ borderColor: '#000000' }}>{person.rank}</td>
                      <td className="border-2 p-4 font-bold text-lg" style={{ borderColor: '#000000' }}>{person.name}</td>
                      <td className="border-2 p-4 text-center text-sm font-semibold" dir="ltr" style={{ borderColor: '#000000' }}>{person.date}<br/>{person.time}</td>
                      <td className="border-2 p-4" style={{ borderColor: '#000000' }}></td>
                    </tr>
                  ))}
                </tbody>
              </table>

              <div className="mt-16 grid grid-cols-3 gap-8 text-center font-bold text-lg break-inside-avoid">
                <div>
                  <p className="mb-12">المسؤول المالي</p>
                  <p className="border-t-2 pt-2 w-3/4 mx-auto" style={{ borderColor: '#000000' }}>الاسم والتوقيع</p>
                </div>
                <div>
                  <p className="mb-12">المراجعة والتدقيق</p>
                  <p className="border-t-2 pt-2 w-3/4 mx-auto" style={{ borderColor: '#000000' }}>الاسم والتوقيع</p>
                </div>
                <div>
                  <p className="mb-12">الاعتماد</p>
                  <p className="border-t-2 pt-2 w-3/4 mx-auto" style={{ borderColor: '#000000' }}>الاسم والتوقيع</p>
                </div>
              </div>
            </div>
          );
        })}

        {/* Expenses details */}
        {expenses && expenses.length > 0 && (
          <div className="page-break-before-always html2pdf__page-break min-h-[297mm] pt-8">
            <div className="flex justify-between items-start mb-8 border-b-2 pb-4" style={{ borderColor: '#000000' }}>
              <div className="text-right">
                <h3 className="font-bold text-lg">الجمهورية اليمنية</h3>
                <h3 className="font-bold text-lg">وزارة الدفاع</h3>
                <h3 className="font-bold text-lg">الشؤون المالية</h3>
              </div>
              <div className="text-center">
                <h2 className="text-2xl font-black mb-2 underline decoration-double underline-offset-8">كشف أوامر الصرف المستقلة</h2>
              </div>
              <div className="text-left">
                <h3 className="font-bold text-lg">التاريخ: {today}</h3>
                <h3 className="font-bold text-lg mt-1">العدد: {expenses.length}</h3>
              </div>
            </div>

            <table className="w-full text-right border-collapse text-base" style={{ borderColor: '#000000' }}>
              <thead>
                <tr style={{ backgroundColor: '#e5e7eb' }}>
                  <th className="border-2 p-3 font-bold w-12 text-center" style={{ borderColor: '#000000' }}>م</th>
                  <th className="border-2 p-3 font-bold w-48" style={{ borderColor: '#000000' }}>المستلم</th>
                  <th className="border-2 p-3 font-bold" style={{ borderColor: '#000000' }}>الغرض</th>
                  <th className="border-2 p-3 font-bold w-32 text-center" style={{ borderColor: '#000000' }}>المبلغ</th>
                  <th className="border-2 p-3 font-bold w-40 text-center" style={{ borderColor: '#000000' }}>التاريخ والوقت</th>
                  <th className="border-2 p-3 font-bold w-48 text-center" style={{ borderColor: '#000000' }}>التوقيع / البصمة</th>
                </tr>
              </thead>
              <tbody>
                {expenses.map((expense, idx) => (
                  <tr key={expense.id} className="break-inside-avoid">
                    <td className="border-2 p-4 text-center font-bold" style={{ borderColor: '#000000' }}>{idx + 1}</td>
                    <td className="border-2 p-4 font-bold text-lg" style={{ borderColor: '#000000' }}>{expense.recipient}</td>
                    <td className="border-2 p-4 font-semibold" style={{ borderColor: '#000000' }}>{expense.purpose}</td>
                    <td className="border-2 p-4 font-bold text-center" style={{ borderColor: '#000000' }}>{expense.amount.toLocaleString()}</td>
                    <td className="border-2 p-4 text-center text-sm font-semibold" dir="ltr" style={{ borderColor: '#000000' }}>{expense.date}<br/>{expense.time}</td>
                    <td className="border-2 p-4" style={{ borderColor: '#000000' }}></td>
                  </tr>
                ))}
              </tbody>
            </table>

            <div className="mt-16 grid grid-cols-3 gap-8 text-center font-bold text-lg break-inside-avoid">
              <div>
                <p className="mb-12">المسؤول المالي</p>
                <p className="border-t-2 pt-2 w-3/4 mx-auto" style={{ borderColor: '#000000' }}>الاسم والتوقيع</p>
              </div>
              <div>
                <p className="mb-12">المراجعة والتدقيق</p>
                <p className="border-t-2 pt-2 w-3/4 mx-auto" style={{ borderColor: '#000000' }}>الاسم والتوقيع</p>
              </div>
              <div>
                <p className="mb-12">الاعتماد</p>
                <p className="border-t-2 pt-2 w-3/4 mx-auto" style={{ borderColor: '#000000' }}>الاسم والتوقيع</p>
              </div>
            </div>
          </div>
        )}

        <style>
          {`
            /* Force all elements in the print container to have safe computed colors */
            .print-safe-colors {
              box-shadow: none !important;
              text-shadow: none !important;
            }

            @media print {
              @page {
                size: A4;
                margin: 15mm;
              }
              body {
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
              }
              .page-break-after-always {
                page-break-after: always;
              }
              .page-break-before-always {
                page-break-before: always;
              }
              .break-inside-avoid {
                page-break-inside: avoid;
              }
            }
          `}
        </style>
      </div>
    );
  }
);

PrintReport.displayName = 'PrintReport';
