import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export function formatCurrency(amount: number) {
  return new Intl.NumberFormat("ar-SA", {
    style: "currency",
    currency: "YER",
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(amount).replace("ر.ي.‏", "ر.ي");
}

export function normalizeArabic(text: string) {
  return text
    .replace(/[\u064B-\u065F\u0670\u0640]/g, "")
    .replace(/أ|إ|آ/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/ؤ/g, "و")
    .replace(/ئ/g, "ي")
    .toLowerCase();
}
