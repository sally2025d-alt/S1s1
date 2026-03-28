import React, { useState } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "./ui/Card";
import { Input } from "./ui/Input";
import { Button } from "./ui/Button";
import { formatCurrency } from "../lib/utils";
import { Briefcase, TrendingDown, CheckCircle, Users, ArrowUpRight, ArrowDownRight, Activity, Clock, Receipt, Wallet, X, CheckCircle2 } from "lucide-react";
import { Doughnut } from "react-chartjs-2";
import { Chart as ChartJS, ArcElement, Tooltip, Legend } from "chart.js";
import { Department, Expense } from "../types";
import { motion, AnimatePresence } from "motion/react";

ChartJS.register(ArcElement, Tooltip, Legend);

interface FinanceTabProps {
  covenant: number;
  stats: {
    totalSpent: number;
    remaining: number;
    pendingPersons: number;
    totalPersons: number;
    receivedPersons: number;
  };
  onUpdateCovenant: (amount: number) => void;
  onAddExpense: (amount: number, recipient: string, purpose: string) => void;
  departments: Department[];
  expenses: Expense[];
}

export function FinanceTab({ covenant, stats, onUpdateCovenant, onAddExpense, departments, expenses }: FinanceTabProps) {
  const [amountInput, setAmountInput] = useState("");
  const [showExpenseModal, setShowExpenseModal] = useState(false);
  const [expenseRecipient, setExpenseRecipient] = useState("");
  const [expensePurpose, setExpensePurpose] = useState("");
  const [showSuccess, setShowSuccess] = useState(false);

  const handleDeposit = () => {
    const amount = parseInt(amountInput);
    if (!isNaN(amount) && amount > 0) {
      onUpdateCovenant(covenant + amount);
      setAmountInput("");
    }
  };

  const handleWithdrawClick = () => {
    const amount = parseInt(amountInput);
    if (isNaN(amount) || amount <= 0) {
      alert("الرجاء إدخال مبلغ صحيح للصرف");
      return;
    }
    if (amount > stats.remaining) {
      alert("المبلغ المطلوب صرفه أكبر من المتبقي من العهدة!");
      return;
    }
    setShowExpenseModal(true);
  };

  const handleConfirmExpense = () => {
    const amount = parseInt(amountInput);
    if (!expenseRecipient.trim() || !expensePurpose.trim()) {
      alert("الرجاء إدخال المستلم والغرض");
      return;
    }
    
    onAddExpense(amount, expenseRecipient, expensePurpose);
    setAmountInput("");
    setExpenseRecipient("");
    setExpensePurpose("");
    setShowExpenseModal(false);
    
    setShowSuccess(true);
    setTimeout(() => {
      setShowSuccess(false);
    }, 2000);
  };

  const chartData = {
    labels: ["المصروف", "المتبقي"],
    datasets: [
      {
        data: [stats.totalSpent, Math.max(0, stats.remaining)],
        backgroundColor: ["#10b981", "#e2e8f0"],
        borderWidth: 0,
      },
    ],
  };

  // Get recent transactions (combining persons and expenses)
  const recentPersonTxs = departments
    .flatMap(d => d.persons.filter(p => p.received).map(p => ({ 
      id: p.id,
      name: p.name, 
      deptName: d.name, 
      time: p.time, 
      date: p.date,
      totalAmount: p.totalAmount,
      type: 'person'
    })));
    
  const recentExpenseTxs = expenses.map(e => ({
    id: e.id,
    name: e.recipient,
    deptName: e.purpose,
    time: e.time,
    date: e.date,
    totalAmount: e.amount,
    type: 'expense'
  }));

  const recentTransactions = [...recentPersonTxs, ...recentExpenseTxs]
    .sort((a, b) => {
      // Simple sort by assuming later added are at the end, or we can just reverse the combined array if we assume chronological addition
      // Since we don't have full timestamps, we'll just reverse the concatenated array for simplicity, or sort by id if they contain timestamps
      return 0; // We will just reverse the whole array below
    })
    .reverse()
    .slice(0, 5);

  const spentPercentage = covenant > 0 ? ((stats.totalSpent / covenant) * 100).toFixed(1) : "0";
  const avgPayment = stats.receivedPersons > 0 ? (stats.totalSpent / stats.receivedPersons).toFixed(0) : "0";

  return (
    <div className="space-y-4 pb-24">
      {/* Main Covenant Card */}
      <Card className="border-none shadow-sm overflow-hidden">
        <div className="bg-gradient-to-br from-slate-900 to-slate-800 p-6 text-white relative">
          <div className="absolute top-0 right-0 p-32 bg-white/5 rounded-full blur-3xl -mr-16 -mt-16 pointer-events-none"></div>
          <div className="relative z-10">
            <div className="flex items-center justify-between mb-6">
              <div className="flex items-center gap-2">
                <div className="w-10 h-10 bg-white/10 rounded-xl flex items-center justify-center backdrop-blur-sm">
                  <Wallet className="w-5 h-5 text-blue-400" />
                </div>
                <div>
                  <h2 className="text-sm font-medium text-slate-300">إجمالي العهدة الحالية</h2>
                  <p className="text-xs text-slate-400">الرصيد المتاح للصرف</p>
                </div>
              </div>
              <div className="text-left">
                <div className="text-3xl font-black tracking-tight">{covenant.toLocaleString()}</div>
                <div className="text-sm text-blue-400 font-bold">ريال يمني</div>
              </div>
            </div>

            <div className="bg-white/10 backdrop-blur-md rounded-2xl p-4 border border-white/10">
              <div className="flex gap-2">
                <Input
                  type="number"
                  placeholder="أدخل المبلغ..."
                  value={amountInput}
                  onChange={(e) => setAmountInput(e.target.value)}
                  className="bg-white/5 border-white/10 text-white placeholder:text-slate-400 text-center font-bold h-12 flex-1"
                />
                <Button
                  onClick={handleDeposit}
                  className="bg-emerald-500 hover:bg-emerald-600 text-white h-12 px-4 flex items-center gap-1"
                >
                  <ArrowUpRight className="w-4 h-4" />
                  إيداع
                </Button>
                <Button
                  onClick={handleWithdrawClick}
                  className="bg-rose-500 hover:bg-rose-600 text-white h-12 px-4 flex items-center gap-1"
                >
                  <ArrowDownRight className="w-4 h-4" />
                  صرف
                </Button>
              </div>
            </div>
          </div>
        </div>
      </Card>

      {/* Financial Indicators */}
      <div className="grid grid-cols-2 gap-3">
        <StatBox
          icon={<TrendingDown className="w-5 h-5 text-emerald-600" />}
          value={(stats.totalSpent / 1000).toFixed(1) + "K"}
          label="إجمالي المصروف"
          bg="bg-emerald-100"
        />
        <StatBox
          icon={<CheckCircle className="w-5 h-5 text-amber-600" />}
          value={(stats.remaining / 1000).toFixed(1) + "K"}
          label="المتبقي من العهدة"
          bg="bg-amber-100"
        />
        <StatBox
          icon={<Activity className="w-5 h-5 text-blue-600" />}
          value={`${spentPercentage}%`}
          label="نسبة الصرف"
          bg="bg-blue-100"
        />
        <StatBox
          icon={<Receipt className="w-5 h-5 text-purple-600" />}
          value={Number(avgPayment).toLocaleString()}
          label="متوسط الصرف للفرد"
          bg="bg-purple-100"
        />
      </div>

      {/* Charts & Analysis */}
      <Card className="border-none shadow-sm">
        <CardHeader className="pb-2 border-b border-slate-50">
          <CardTitle className="text-base font-bold flex items-center gap-2">
            <Activity className="w-5 h-5 text-blue-500" />
            تحليل الميزانية
          </CardTitle>
        </CardHeader>
        <CardContent className="pt-6">
          <div className="h-[220px] relative flex justify-center">
            <Doughnut
              data={chartData}
              options={{
                responsive: true,
                maintainAspectRatio: false,
                cutout: "75%",
                plugins: {
                  legend: {
                    position: "bottom",
                    labels: {
                      font: { family: "Cairo", size: 13, weight: "bold" },
                      padding: 20,
                      usePointStyle: true,
                    },
                  },
                },
              }}
            />
            <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none pb-8">
              <span className="text-3xl font-black text-slate-800">{spentPercentage}%</span>
              <span className="text-xs font-bold text-slate-400">تم صرفه</span>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* Recent Transactions */}
      <Card className="border-none shadow-sm">
        <CardHeader className="pb-2 border-b border-slate-50">
          <CardTitle className="text-base font-bold flex items-center gap-2">
            <Clock className="w-5 h-5 text-indigo-500" />
            أحدث عمليات الصرف
          </CardTitle>
        </CardHeader>
        <CardContent className="pt-4 p-0">
          {recentTransactions.length > 0 ? (
            <div className="divide-y divide-slate-50">
              {recentTransactions.map((tx, idx) => (
                <div key={idx} className="p-4 flex items-center justify-between hover:bg-slate-50 transition-colors">
                  <div className="flex items-center gap-3">
                    <div className={`w-10 h-10 rounded-full flex items-center justify-center shrink-0 ${tx.type === 'expense' ? 'bg-rose-100 text-rose-600' : 'bg-emerald-100 text-emerald-600'}`}>
                      {tx.type === 'expense' ? <ArrowDownRight className="w-5 h-5" /> : <CheckCircle className="w-5 h-5" />}
                    </div>
                    <div>
                      <p className="text-sm font-bold text-slate-800">{tx.name}</p>
                      <p className="text-xs text-slate-500">{tx.deptName} • {tx.time}</p>
                    </div>
                  </div>
                  <div className="text-left">
                    <p className={`text-sm font-black ${tx.type === 'expense' ? 'text-rose-600' : 'text-emerald-600'}`}>-{tx.totalAmount.toLocaleString()}</p>
                    <p className="text-xs text-slate-400">ر.ي</p>
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <div className="p-8 text-center text-slate-500 text-sm">
              لا توجد عمليات صرف حديثة
            </div>
          )}
        </CardContent>
      </Card>

      {/* Expense Modal */}
      <AnimatePresence>
        {showExpenseModal && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4"
            onClick={() => setShowExpenseModal(false)}
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
                  onClick={() => setShowExpenseModal(false)}
                  className="absolute left-4 top-4 w-8 h-8 bg-slate-100 rounded-full flex items-center justify-center text-slate-500 hover:bg-slate-200"
                >
                  <X className="w-4 h-4" />
                </button>
                <div className="w-16 h-16 bg-rose-50 rounded-2xl mx-auto mb-4 flex items-center justify-center text-rose-600">
                  <ArrowDownRight className="w-8 h-8" />
                </div>
                <h2 className="text-xl font-black text-slate-900">توجيهات الصرف</h2>
                <p className="text-sm font-bold text-rose-600 mt-1">المبلغ: {parseInt(amountInput).toLocaleString()} ر.ي</p>
              </div>

              <div className="p-6 space-y-4">
                <div className="space-y-2">
                  <label className="text-sm font-bold text-slate-700">المستلم</label>
                  <Input
                    placeholder="اسم المستلم..."
                    value={expenseRecipient}
                    onChange={(e) => setExpenseRecipient(e.target.value)}
                    className="bg-slate-50 border-slate-200"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-sm font-bold text-slate-700">الغرض</label>
                  <Input
                    placeholder="غرض الصرف..."
                    value={expensePurpose}
                    onChange={(e) => setExpensePurpose(e.target.value)}
                    className="bg-slate-50 border-slate-200"
                  />
                </div>

                <div className="pt-2">
                  <Button 
                    onClick={handleConfirmExpense} 
                    className="w-full h-14 rounded-2xl bg-rose-500 hover:bg-rose-600 text-white text-base font-bold shadow-lg shadow-rose-500/20"
                  >
                    تأكيد الصرف
                  </Button>
                </div>
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
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

function StatBox({ icon, value, label, bg }: { icon: React.ReactNode; value: string; label: string; bg: string }) {
  return (
    <div className="bg-white rounded-2xl p-4 text-center border border-slate-100 shadow-sm">
      <div className={`w-10 h-10 mx-auto rounded-xl flex items-center justify-center mb-2 ${bg}`}>
        {icon}
      </div>
      <div className="text-lg font-black text-slate-800 mb-0.5">{value}</div>
      <div className="text-xs font-bold text-slate-500">{label}</div>
    </div>
  );
}
