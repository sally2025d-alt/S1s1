import React, { useState, useMemo, useRef } from "react";
import { Department, Person, BASE_AMOUNT } from "../types";
import { Input } from "./ui/Input";
import { Button } from "./ui/Button";
import { Search, X, Check, Upload, ChevronDown, ChevronUp } from "lucide-react";
import { cn, normalizeArabic } from "../lib/utils";
import { motion, AnimatePresence } from "motion/react";
import * as XLSX from "xlsx";

interface PaymentTabProps {
  departments: Department[];
  onSelectPerson: (deptId: number, personId: string) => void;
  stats: {
    totalPersons: number;
    pendingPersons: number;
    receivedPersons: number;
  };
  onImportData?: (data: any[]) => void;
  onCancelPayment?: (deptId: number, personId: string) => void;
}

type FilterType = "all" | "pending" | "received";

export function PaymentTab({ departments, onSelectPerson, stats, onImportData, onCancelPayment }: PaymentTabProps) {
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState<FilterType>("all");
  const [deptFilter, setDeptFilter] = useState<number | "all">("all");
  const [collapsedDepts, setCollapsedDepts] = useState<Record<number, boolean>>({});
  const [personToCancel, setPersonToCancel] = useState<{deptId: number, person: Person} | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const toggleDept = (deptId: number) => {
    setCollapsedDepts(prev => ({ ...prev, [deptId]: !prev[deptId] }));
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: "binary" });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
      if (onImportData) {
        onImportData(data);
      }
    };
    reader.readAsBinaryString(file);
    
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const filteredDepartments = useMemo(() => {
    const normalizedSearch = normalizeArabic(search);
    
    return departments
      .filter(dept => deptFilter === "all" || dept.id === deptFilter)
      .map(dept => {
        const persons = dept.persons.filter(p => {
          // Search filter
          const nameMatch = normalizeArabic(p.name).includes(normalizedSearch);
          const numberMatch = p.militaryNumber.includes(search);
          const rankMatch = normalizeArabic(p.rank).includes(normalizedSearch);
          const matchesSearch = nameMatch || numberMatch || rankMatch;

          // Status filter
          let matchesStatus = true;
          if (filter === "received") matchesStatus = p.received;
          if (filter === "pending") matchesStatus = !p.received;

          return matchesSearch && matchesStatus;
        });
        return { ...dept, persons };
      }).filter(dept => dept.persons.length > 0);
  }, [departments, search, filter, deptFilter]);

  return (
    <div className="space-y-4 pb-24">
      {/* Search Bar & Import */}
      <div className="flex gap-2">
        <div className="relative flex-1">
          <Search className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400 w-5 h-5" />
          <Input
            placeholder="ابحث بالاسم أو الرقم العسكري..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            className="pl-12 pr-12 h-14 bg-white border-slate-200 shadow-sm rounded-2xl text-base font-semibold"
          />
          {search && (
            <button
              onClick={() => setSearch("")}
              className="absolute left-4 top-1/2 -translate-y-1/2 w-8 h-8 flex items-center justify-center rounded-full bg-slate-100 text-slate-500 hover:bg-slate-200 transition-colors"
            >
              <X className="w-4 h-4" />
            </button>
          )}
        </div>
        <input 
          type="file" 
          accept=".xlsx, .xls" 
          className="hidden" 
          ref={fileInputRef} 
          onChange={handleFileUpload} 
        />
        <Button 
          onClick={() => fileInputRef.current?.click()}
          className="h-14 w-14 rounded-2xl bg-white border border-slate-200 text-slate-600 hover:bg-slate-50 shadow-sm shrink-0"
          title="استيراد من إكسل"
        >
          <Upload className="w-5 h-5" />
        </Button>
      </div>

      {/* Filters */}
      <div className="space-y-2">
        {/* Status Filters */}
        <div className="flex gap-2 overflow-x-auto no-scrollbar pb-1">
          <FilterChip
            active={filter === "all"}
            onClick={() => setFilter("all")}
            label="حالة الصرف: الكل"
            count={stats.totalPersons}
          />
          <FilterChip
            active={filter === "pending"}
            onClick={() => setFilter("pending")}
            label="باقي"
            count={stats.pendingPersons}
          />
          <FilterChip
            active={filter === "received"}
            onClick={() => setFilter("received")}
            label="تم"
            count={stats.receivedPersons}
          />
        </div>

        {/* Department Filters */}
        <div className="relative mb-2">
          <select
            value={deptFilter}
            onChange={(e) => setDeptFilter(e.target.value === "all" ? "all" : Number(e.target.value))}
            className="w-full bg-white border border-slate-200 text-slate-900 text-sm rounded-xl focus:ring-2 focus:ring-blue-500 block p-3 pr-4 pl-10 appearance-none font-bold outline-none shadow-sm"
          >
            <option value="all">كل الأقسام ({stats.totalPersons})</option>
            {departments.map(d => (
              <option key={d.id} value={d.id}>{d.name} ({d.persons.length})</option>
            ))}
          </select>
          <ChevronDown className="absolute left-3 top-3 w-5 h-5 text-slate-500 pointer-events-none" />
        </div>
      </div>

      {/* Departments List */}
      <div className="space-y-6">
        <AnimatePresence mode="popLayout">
          {filteredDepartments.length === 0 ? (
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="text-center py-16 text-slate-400"
            >
              <div className="text-6xl mb-4 opacity-30">😕</div>
              <div className="font-bold text-lg">لا توجد نتائج</div>
            </motion.div>
          ) : (
            filteredDepartments.map(dept => (
              <motion.div
                key={dept.id}
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0 }}
                transition={{ duration: 0.15 }}
                className="space-y-3"
              >
                <DepartmentHeader 
                  dept={dept} 
                  isCollapsed={collapsedDepts[dept.id] || false} 
                  onToggle={() => toggleDept(dept.id)} 
                />
                <AnimatePresence initial={false}>
                  {!collapsedDepts[dept.id] && (
                    <motion.div
                      initial={{ height: 0, opacity: 0 }}
                      animate={{ height: "auto", opacity: 1 }}
                      exit={{ height: 0, opacity: 0 }}
                      transition={{ duration: 0.2, ease: "easeInOut" }}
                      className="overflow-hidden"
                    >
                      <div className="flex flex-col gap-2 mt-3">
                        {dept.persons.map(person => (
                          <CompactPersonCard
                            key={person.id}
                            person={person}
                            onClick={() => onSelectPerson(dept.id, person.id)}
                            onCancel={() => setPersonToCancel({ deptId: dept.id, person })}
                          />
                        ))}
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </motion.div>
            ))
          )}
        </AnimatePresence>
      </div>

      {/* Cancel Confirmation Modal */}
      <AnimatePresence>
        {personToCancel && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[70] flex items-center justify-center p-4"
            onClick={() => setPersonToCancel(null)}
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="bg-white rounded-3xl w-full max-w-sm overflow-hidden shadow-2xl p-6 text-center"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="w-16 h-16 bg-rose-100 rounded-full mx-auto mb-4 flex items-center justify-center text-rose-600">
                <X className="w-8 h-8" />
              </div>
              <h3 className="text-xl font-black text-slate-900 mb-2">إلغاء الصرف</h3>
              <p className="text-slate-500 mb-6 font-semibold">
                هل أنت متأكد من إلغاء الصرف للفرد <span className="text-slate-800 font-bold">{personToCancel.person.name}</span>؟
              </p>
              <div className="flex gap-3">
                <Button
                  onClick={() => setPersonToCancel(null)}
                  className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold h-12 rounded-xl"
                >
                  تراجع
                </Button>
                <Button
                  onClick={() => {
                    if (onCancelPayment) {
                      onCancelPayment(personToCancel.deptId, personToCancel.person.id);
                    }
                    setPersonToCancel(null);
                  }}
                  className="flex-1 bg-rose-500 hover:bg-rose-600 text-white font-bold h-12 rounded-xl shadow-lg shadow-rose-500/20"
                >
                  نعم، إلغاء
                </Button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

function FilterChip({ active, onClick, label, count }: { active: boolean; onClick: () => void; label: string; count: number }) {
  return (
    <button
      onClick={onClick}
      className={cn(
        "flex items-center gap-2 px-5 py-2.5 rounded-full text-sm font-bold whitespace-nowrap transition-all",
        active
          ? "bg-blue-600 text-white shadow-md shadow-blue-500/20"
          : "bg-white text-slate-500 border border-slate-200 hover:bg-slate-50"
      )}
    >
      {label}
      <span
        className={cn(
          "px-2 py-0.5 rounded-lg text-xs",
          active ? "bg-white/20" : "bg-slate-100"
        )}
      >
        {count}
      </span>
    </button>
  );
}

function DepartmentHeader({ dept, isCollapsed, onToggle }: { dept: Department, isCollapsed: boolean, onToggle: () => void }) {
  const total = dept.persons.length;
  const received = dept.persons.filter(p => p.received).length;
  const percent = total > 0 ? Math.round((received / total) * 100) : 0;
  const circumference = 2 * Math.PI * 22;
  const offset = circumference - (percent / 100) * circumference;

  return (
    <div 
      onClick={onToggle}
      className="bg-gradient-to-br from-slate-900 to-slate-800 text-white p-3 rounded-2xl flex justify-between items-center shadow-md cursor-pointer select-none transition-transform active:scale-[0.98]"
    >
      <div className="flex items-center gap-3">
        <div className={cn("transition-transform duration-300", isCollapsed ? "rotate-180" : "")}>
          <ChevronUp className="w-5 h-5 text-slate-400" />
        </div>
        <div>
          <h3 className="font-bold text-base mb-0.5">{dept.name}</h3>
          <div className="text-xs font-semibold opacity-70">
            {received}/{total} مستلم • {percent}%
          </div>
        </div>
      </div>
      <div className="relative w-10 h-10 shrink-0">
        <svg className="w-full h-full -rotate-90" viewBox="0 0 50 50">
          <circle
            cx="25"
            cy="25"
            r="22"
            fill="none"
            stroke="rgba(255,255,255,0.1)"
            strokeWidth="4"
          />
          <circle
            cx="25"
            cy="25"
            r="22"
            fill="none"
            stroke="white"
            strokeWidth="4"
            strokeLinecap="round"
            strokeDasharray={circumference}
            strokeDashoffset={offset}
            className="transition-all duration-1000 ease-out"
          />
        </svg>
        <div className="absolute inset-0 flex items-center justify-center text-[9px] font-black">
          {percent}%
        </div>
      </div>
    </div>
  );
}

function CompactPersonCard({ person, onClick, onCancel }: { person: Person; onClick: () => void; onCancel?: () => void; key?: React.Key }) {
  return (
    <motion.div
      onClick={onClick}
      whileTap={{ scale: 0.98 }}
      className={cn(
        "bg-white rounded-xl p-3 cursor-pointer border flex items-center gap-3 transition-colors shadow-sm relative group",
        person.received
          ? "border-emerald-200 bg-emerald-50/30"
          : "border-slate-100 hover:border-blue-300"
      )}
    >
      {/* Avatar */}
      <div
        className={cn(
          "w-10 h-10 rounded-lg flex items-center justify-center font-black text-white shrink-0",
          person.received
            ? "bg-emerald-500"
            : "bg-gradient-to-br from-blue-500 to-blue-400"
        )}
      >
        {person.received ? <Check className="w-5 h-5 stroke-[3]" /> : person.rank[0]}
      </div>

      {/* Info */}
      <div className="flex-1 min-w-0">
        <div className="font-bold text-sm text-slate-800 truncate">{person.name}</div>
        <div className="text-[11px] font-semibold text-slate-400 truncate">{person.militaryNumber} • {person.rank}</div>
      </div>

      {/* Amount & Status */}
      <div className="text-left shrink-0 flex items-center gap-2">
        <div
          className={cn(
            "px-2 py-1 rounded-md text-xs font-black",
            person.received
              ? "bg-emerald-100 text-emerald-700"
              : "bg-slate-100 text-slate-600"
          )}
        >
          {person.totalAmount.toLocaleString()} ر.ي
        </div>
        
        {person.received && onCancel && (
          <button
            onClick={(e) => {
              e.stopPropagation();
              onCancel();
            }}
            className="w-8 h-8 rounded-full bg-rose-100 text-rose-600 flex items-center justify-center hover:bg-rose-200 transition-colors shadow-sm shrink-0"
            title="إلغاء الصرف"
          >
            <X className="w-4 h-4" />
          </button>
        )}
      </div>
    </motion.div>
  );
}
