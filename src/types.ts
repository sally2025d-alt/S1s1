export interface Person {
  id: string;
  name: string;
  rank: string;
  militaryNumber: string;
  received: boolean;
  date: string | null;
  time: string | null;
  baseAmount: number;
  bonus: number;
  totalAmount: number;
}

export interface Department {
  id: number;
  name: string;
  persons: Person[];
}

export interface Expense {
  id: string;
  amount: number;
  recipient: string;
  purpose: string;
  date: string;
  time: string;
}

export interface ArchiveRecord {
  id: string;
  date: string;
  covenant: number;
  expenses: Expense[];
  departments: Department[];
  stats: {
    totalSpent: number;
    remaining: number;
    totalPersons: number;
    receivedPersons: number;
    pendingPersons: number;
  };
}

export interface SystemState {
  departments: Department[];
  expenses: Expense[];
  covenant: number;
  lastUpdate: string;
  archives: ArchiveRecord[];
}

export const BASE_AMOUNT = 4000;
