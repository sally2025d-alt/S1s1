import { useState, useEffect, useCallback } from 'react';
import { Department, Person, SystemState, Expense, ArchiveRecord, BASE_AMOUNT } from '../types';

const STORAGE_KEY = 'militaryPaymentSystem_v2';

function createDefaultData(): SystemState {
  const depts = ['القسم الأول', 'القسم الثاني', 'القسم الثالث', 'القسم الرابع'];
  const ranks = ['جندي', 'عريف', 'رقيب', 'رقيب أول', 'ملازم', 'ملازم أول', 'نقيب'];
  
  const departments: Department[] = depts.map((name, idx) => ({
      id: idx + 1,
      name: name,
      persons: Array.from({length: 20}, (_, i) => ({
          id: `d${idx+1}-p${i+1}`,
          name: `فرد ${i + 1}`,
          rank: ranks[Math.floor(Math.random() * ranks.length)],
          militaryNumber: `${2024}${(idx+1).toString().padStart(2,'0')}${(i+1).toString().padStart(3,'0')}`,
          received: false,
          date: null,
          time: null,
          baseAmount: BASE_AMOUNT,
          bonus: 0,
          totalAmount: BASE_AMOUNT
      }))
  }));
  
  return {
    departments,
    expenses: [],
    covenant: 320000, // Default covenant: 20 persons * 4 depts * 4000
    lastUpdate: new Date().toISOString(),
    archives: []
  };
}

export function useFinanceStore() {
  const [state, setState] = useState<SystemState>(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const parsedState = JSON.parse(saved) as SystemState;
        
        // Migration: ensure all persons have baseAmount, bonus, and totalAmount
        const migratedDepartments = parsedState.departments.map(dept => ({
          ...dept,
          persons: dept.persons.map(p => ({
            ...p,
            baseAmount: p.baseAmount ?? BASE_AMOUNT,
            bonus: p.bonus ?? 0,
            totalAmount: p.totalAmount ?? (p.baseAmount ?? BASE_AMOUNT)
          }))
        }));

        return {
          ...parsedState,
          departments: migratedDepartments,
          expenses: parsedState.expenses || [],
          archives: parsedState.archives || []
        };
      } catch (e) {
        console.error("Failed to parse saved state", e);
      }
    }
    return createDefaultData();
  });

  // Save to local storage whenever state changes
  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }, [state]);

  const updateCovenant = useCallback((amount: number) => {
    setState(prev => ({ ...prev, covenant: amount, lastUpdate: new Date().toISOString() }));
  }, []);

  const addExpense = useCallback((amount: number, recipient: string, purpose: string) => {
    setState(prev => {
      const now = new Date();
      const newExpense: Expense = {
        id: `exp-${Date.now()}`,
        amount,
        recipient,
        purpose,
        date: now.toLocaleDateString('ar-SA'),
        time: now.toLocaleTimeString('ar-SA', { hour: '2-digit', minute: '2-digit' })
      };
      return {
        ...prev,
        expenses: [...prev.expenses, newExpense],
        lastUpdate: now.toISOString()
      };
    });
  }, []);

  const confirmPayment = useCallback((deptId: number, personId: string, bonus: number) => {
    setState(prev => {
      const now = new Date();
      const newDepts = prev.departments.map(dept => {
        if (dept.id !== deptId) return dept;
        return {
          ...dept,
          persons: dept.persons.map(p => {
            if (p.id !== personId) return p;
            const total = p.baseAmount + bonus;
            return {
              ...p,
              received: true,
              date: now.toLocaleDateString('ar-SA'),
              time: now.toLocaleTimeString('ar-SA', {hour: '2-digit', minute:'2-digit'}),
              bonus,
              totalAmount: total
            };
          })
        };
      });
      return { ...prev, departments: newDepts, lastUpdate: now.toISOString() };
    });
  }, []);

  const cancelPayment = useCallback((deptId: number, personId: string) => {
    setState(prev => {
      const now = new Date();
      const newDepts = prev.departments.map(dept => {
        if (dept.id !== deptId) return dept;
        return {
          ...dept,
          persons: dept.persons.map(p => {
            if (p.id !== personId) return p;
            return {
              ...p,
              received: false,
              date: null,
              time: null,
              bonus: 0,
              totalAmount: p.baseAmount
            };
          })
        };
      });
      return { ...prev, departments: newDepts, lastUpdate: now.toISOString() };
    });
  }, []);

  const deleteArchive = useCallback((id: string) => {
    setState(prev => ({
      ...prev,
      archives: (prev.archives || []).filter(a => a.id !== id),
      lastUpdate: new Date().toISOString()
    }));
  }, []);

  const clearArchives = useCallback(() => {
    setState(prev => ({
      ...prev,
      archives: [],
      lastUpdate: new Date().toISOString()
    }));
  }, []);

  const resetDay = useCallback(() => {
    setState(prev => {
      // Calculate stats for the archive
      const totalSpentPersons = prev.departments.reduce((sum, d) => 
        sum + d.persons.filter(p => p.received).reduce((s, p) => s + p.totalAmount, 0), 0);
      const totalSpentExpenses = prev.expenses.reduce((sum, exp) => sum + exp.amount, 0);
      const totalSpent = totalSpentPersons + totalSpentExpenses;
      const remaining = prev.covenant - totalSpent;
      const totalPersons = prev.departments.reduce((sum, d) => sum + d.persons.length, 0);
      const receivedPersons = prev.departments.reduce((sum, d) => sum + d.persons.filter(p => p.received).length, 0);
      const pendingPersons = totalPersons - receivedPersons;

      const archiveRecord: ArchiveRecord = {
        id: `arch-${Date.now()}`,
        date: new Date().toLocaleDateString('ar-SA'),
        covenant: prev.covenant,
        expenses: [...prev.expenses],
        departments: JSON.parse(JSON.stringify(prev.departments)), // Deep copy to preserve state
        stats: {
          totalSpent,
          remaining,
          totalPersons,
          receivedPersons,
          pendingPersons
        }
      };

      return {
        ...prev,
        covenant: remaining,
        departments: prev.departments.map(dept => ({
          ...dept,
          persons: dept.persons.map(p => ({
            ...p,
            received: false,
            date: null,
            time: null,
            bonus: 0,
            totalAmount: p.baseAmount
          }))
        })),
        expenses: [],
        archives: [archiveRecord, ...(prev.archives || [])],
        lastUpdate: new Date().toISOString()
      };
    });
  }, []);

  const importData = useCallback((importedData: any[]) => {
    setState(prev => {
      const newDepartments: Department[] = [];
      let deptIdCounter = 1;

      // Find the header row to determine column indices
      let headerRowIndex = 0;
      let codeIdx = 0;
      let nameIdx = 1;
      let deptIdx = 2;

      // Try to find the header row
      for (let i = 0; i < Math.min(5, importedData.length); i++) {
        const row = importedData[i];
        if (!row) continue;
        
        const codeMatch = row.findIndex((cell: any) => typeof cell === 'string' && cell.includes('الكود'));
        const nameMatch = row.findIndex((cell: any) => typeof cell === 'string' && cell.includes('الاسم'));
        const deptMatch = row.findIndex((cell: any) => typeof cell === 'string' && cell.includes('القسم'));

        if (codeMatch !== -1 || nameMatch !== -1 || deptMatch !== -1) {
          headerRowIndex = i;
          if (codeMatch !== -1) codeIdx = codeMatch;
          if (nameMatch !== -1) nameIdx = nameMatch;
          if (deptMatch !== -1) deptIdx = deptMatch;
          break;
        }
      }

      for (let i = headerRowIndex + 1; i < importedData.length; i++) {
        const row = importedData[i];
        if (!row) continue;

        const code = row[codeIdx];
        const name = row[nameIdx];
        const deptName = row[deptIdx];

        if (!code || !name || !deptName) continue;

        let dept = newDepartments.find(d => d.name === deptName.toString().trim());
        if (!dept) {
          dept = {
            id: deptIdCounter++,
            name: deptName.toString().trim(),
            persons: []
          };
          newDepartments.push(dept);
        }

        dept.persons.push({
          id: `imported-${code}`,
          name: name.toString().trim(),
          rank: 'جندي', // Default rank
          militaryNumber: code.toString().trim(),
          received: false,
          date: null,
          time: null,
          baseAmount: BASE_AMOUNT,
          bonus: 0,
          totalAmount: BASE_AMOUNT
        });
      }

      if (newDepartments.length === 0) return prev;

      return {
        ...prev,
        departments: newDepartments,
        lastUpdate: new Date().toISOString()
      };
    });
  }, []);

  // Derived state
  const totalSpentPersons = state.departments.reduce((sum, d) => 
    sum + d.persons.filter(p => p.received).reduce((s, p) => s + p.totalAmount, 0), 0);
  
  const totalSpentExpenses = state.expenses.reduce((sum, exp) => sum + exp.amount, 0);

  const totalSpent = totalSpentPersons + totalSpentExpenses;
  
  const remaining = state.covenant - totalSpent;
  
  const totalPersons = state.departments.reduce((sum, d) => sum + d.persons.length, 0);
  const receivedPersons = state.departments.reduce((sum, d) => sum + d.persons.filter(p => p.received).length, 0);
  const pendingPersons = totalPersons - receivedPersons;

  return {
    state,
    updateCovenant,
    addExpense,
    confirmPayment,
    cancelPayment,
    deleteArchive,
    clearArchives,
    resetDay,
    importData,
    stats: {
      totalSpent,
      remaining,
      totalPersons,
      receivedPersons,
      pendingPersons
    }
  };
}
