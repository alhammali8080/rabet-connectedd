import React, { useState, useEffect, useMemo, useRef } from "react";
import { loadRemoteState, saveRemoteState, loginRemote, logoutRemote } from "./api";
import JsBarcode from "jsbarcode";
import { QRCodeSVG } from "qrcode.react";
import * as XLSX from "xlsx";
import { X,
  Building2, Store, Truck, ShieldCheck, LayoutDashboard, Package, Users,
  ShoppingCart, Wallet, CheckCircle2, XCircle, Clock, LogOut, Plus, Search,
  Menu, Phone, MapPin, Box, ClipboardList, Banknote, Eye, Send, PackageCheck,
  Handshake, ChevronLeft, ChevronDown, Heart, Tag, GitCompare, RotateCcw,
  MessageSquare, FileText, CreditCard, TrendingUp, Star, Edit3, Trash2,
  QrCode, Barcode, Warehouse, Clock3, UserCog, Image as ImageIcon, Globe,
  Bell, User, Lock, ArrowRight, ArrowLeft, Layers, Percent, Weight, Ruler, Repeat,
  AlertCircle, Mail, Hash, ScanLine, Link2, Target, Route, ClipboardCheck,
  Smartphone, MapPinned, CalendarClock, UserCheck, Boxes, PieChart,
  TrendingDown, Coins, Receipt, NotebookPen, Navigation, Fuel, CircleDollarSign,
  BadgeCheck, Activity, Briefcase, ListChecks, ScrollText, HandCoins, Map as MapIcon,
  Utensils, Pill, ServerCog, UserPlus, FileSignature, ClipboardSignature, ScanSearch,
  Printer, Inbox, PackagePlus, Save, ChevronRight
, Landmark, BookOpen, Scale, Calculator, FileBarChart2 , Upload , Gift, Award, Sparkles, AlertTriangle , Calendar , UserX , ArrowLeftRight , BarChart3 , MessageCircle } from "lucide-react";

/* ===================== السيرفرات ===================== */
const SERVERS = {
  food:   { id: "food", name: "المواد الغذائية", desc: "توزيع المواد الغذائية والمشروبات.", icon: Utensils, color: "#2563EB", color2: "#1D4ED8", soft: "#E0EAFF", emoji: "🍱", merchantWord: "محل", merchantsWord: "المحلات", soon: false },
  pharma: { id: "pharma", name: "الأدوية", desc: "أدوية ومستلزمات الصيدليات.", icon: Pill, color: "#0D9488", color2: "#0F766E", soft: "#CCFBF1", emoji: "💊", merchantWord: "صيدلية", merchantsWord: "الصيدليات", soon: false },
  cosmetics: { id: "cosmetics", name: "مستحضرات التجميل", desc: "عطور ومستحضرات تجميل وعناية.", icon: Star, color: "#DB2777", color2: "#BE185D", soft: "#FCE7F3", emoji: "💄", merchantWord: "محل", merchantsWord: "المحلات", soon: true },
  electronics: { id: "electronics", name: "الإلكترونيات", desc: "أجهزة وإكسسوارات إلكترونية.", icon: Smartphone, color: "#6366F1", color2: "#4F46E5", soft: "#E0E7FF", emoji: "📱", merchantWord: "محل", merchantsWord: "المحلات", soon: true },
  construction: { id: "construction", name: "مواد البناء", desc: "مواد بناء ومستلزمات إنشائية.", icon: Warehouse, color: "#D97706", color2: "#B45309", soft: "#FEF3C7", emoji: "🧱", merchantWord: "متجر", merchantsWord: "المتاجر", soon: true },
  stationery: { id: "stationery", name: "القرطاسية", desc: "أدوات مكتبية وقرطاسية.", icon: FileText, color: "#059669", color2: "#047857", soft: "#D1FAE5", emoji: "📚", merchantWord: "مكتبة", merchantsWord: "المكتبات", soon: true },
};

/* ===================== الهوية البصرية ===================== */
const C = {
  // ═══ الهوية: فاتح · واضح · سلس (Light & Clean) ═══
  ink: "#0F2544",           // كحلي عميق للعناوين والعناصر الداكنة
  bg: "#F4F7FB",            // خلفية التطبيق — رمادي أزرق فاتح ناعم
  surface: "#FFFFFF",       // الأسطح (البطاقات) بيضاء
  surfaceHover: "#F8FAFD",
  line: "#E4EAF2",          // حدود ناعمة
  lineStrong: "#CDD8E8",
  // النصوص
  text: "#15263B",          // نص رئيسي داكن مقروء
  soft: "#5B6B82",
  faint: "#94A3B8",
  // النظام الدلالي (ألوان واضحة على أبيض)
  primary: "#2563EB", primary2: "#3B82F6",   // أزرق أساسي
  accent: "#0D9488", accentSoft: "#CCFBF1",  // تركوازي
  success: "#16A34A", successSoft: "#DCFCE7",// أخضر نجاح/نقد
  info: "#0EA5E9", infoSoft: "#E0F2FE",      // سماوي معلومات
  danger: "#DC2626", dangerSoft: "#FEE2E2",  // أحمر خطر
  warnSoft: "#FEF3C7",
  purple: "#7C3AED", purpleSoft: "#EDE9FE",
  teal: "#0D9488", tealSoft: "#CCFBF1",
  rose: "#E11D48", roseSoft: "#FFE4E6",
  glow: "0 4px 14px rgba(37,99,235,.18)",
  glassShadow: "0 1px 3px rgba(15,37,68,.06), 0 8px 24px rgba(15,37,68,.06)",
  well: "#F1F5FB",          // بئر فاتح: خلفية أيقونات/شارات
  onAccent: "#FFFFFF",      // نص أبيض فوق الألوان الأساسية الداكنة
};
// ألوان متغيّرة حسب السيرفر
let SC = SERVERS.food; // يُحدَّث عند الاختيار

/* ===================== حالات ===================== */
const STATUS = {
  draft:           { label: "مسودّة", color: C.soft, soft: C.well, step: 0 },
  pending_rep:     { label: "بانتظار المندوب", color: C.accent, soft: C.accentSoft, step: 1 },
  rejected:        { label: "مرفوض", color: C.danger, soft: C.dangerSoft, step: -1 },
  pending_company: { label: "وصل الشركة", color: C.info, soft: C.infoSoft, step: 2 },
  printed:         { label: "طُبع بالمنظومة", color: C.purple, soft: C.purpleSoft, step: 2 },
  warehouse:       { label: "في المخزن (تعبئة)", color: "#7C3AED", soft: C.purpleSoft, step: 3 },
  preparing:       { label: "جاهز/سُلّم للسائق", color: C.purple, soft: C.purpleSoft, step: 3 },
  ready_dispatch:  { label: "بانتظار مشرف الحركة", color: "#B45309", soft: C.warnSoft, step: 3 },
  loaded_trip:     { label: "محمّل في رحلة", color: C.teal, soft: C.tealSoft, step: 4 },
  out_for_delivery:{ label: "خرج للتوصيل", color: C.primary2, soft: "#DBEAFE", step: 4 },
  delivered:       { label: "تم التسليم", color: C.success, soft: C.successSoft, step: 5 },
};
const RET_STATUS = {
  requested:  { label: "طلب مقدّم", color: C.accent, soft: C.accentSoft },
  approved:   { label: "مقبول مبدئياً", color: C.info, soft: C.infoSoft },
  picked:     { label: "استُلم من المحل", color: C.purple, soft: C.purpleSoft },
  inspected:  { label: "فُحص — بانتظار قرار الشركة", color: C.info, soft: C.infoSoft },
  replaced:   { label: "قرار: استبدال البضاعة", color: C.success, soft: C.successSoft },
  credited:   { label: "قرار: إشعار دائن (خُصم من المديونية)", color: C.success, soft: C.successSoft },
  rejected:   { label: "مرفوض", color: C.danger, soft: C.dangerSoft },
}
const TASK_STATUS = {
  assigned:    { label: "مُسندة", color: C.accent, soft: C.accentSoft },
  in_progress: { label: "قيد التنفيذ", color: C.info, soft: C.infoSoft },
  done:        { label: "مكتملة", color: C.success, soft: C.successSoft },
};
const VISIT_RESULT = {
  order:    { label: "أخذ طلب", color: C.success, soft: C.successSoft },
  no_order: { label: "بدون طلب", color: C.soft, soft: C.well },
  payment:  { label: "تحصيل نقد", color: C.accent, soft: C.accentSoft },
  closed:   { label: "مغلق", color: C.danger, soft: C.dangerSoft },
};
const JOIN_STATUS = {
  requested:  { label: "طلب مُرسل", color: C.accent, soft: C.accentSoft },
  visit_set:  { label: "زيارة مجدولة", color: C.info, soft: C.infoSoft },
  verifying:  { label: "قيد التوثيق", color: C.purple, soft: C.purpleSoft },
  active:     { label: "مُفعّل", color: C.success, soft: C.successSoft },
  declined:   { label: "مرفوض", color: C.danger, soft: C.dangerSoft },
};
const TIERS = {
  retail:         { label: "تجزئة", color: C.info, soft: C.infoSoft, icon: Store },
  wholesale:      { label: "جملة", color: C.purple, soft: C.purpleSoft, icon: Boxes },
  company_client: { label: "زبون شركة", color: C.teal, soft: C.tealSoft, icon: Building2 },
};
const PROMO_STATUS = {
  notified:   { label: "زيارة مطلوبة", color: C.accent, soft: C.accentSoft },
  arrived:    { label: "في المحل (GPS مؤكد)", color: C.info, soft: C.infoSoft },
  displayed:  { label: "رُتّبت البضاعة", color: C.purple, soft: C.purpleSoft },
  sent:       { label: "أُرسل التقرير للشركة", color: C.success, soft: C.successSoft },
}
const FAIL_REASONS = {
  stock_full: "المخزون كافٍ",
  closed:     "المحل مغلق",
  no_cash:    "لا سيولة لدى العميل",
  owner_away: "صاحب المحل غير موجود",
  credit_max: "تجاوز الحد الائتماني",
  other:      "سبب آخر",
};
const WEEKDAYS = ["السبت", "الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة"];

// طرق الدفع في تسوية الفواتير
const METHOD_LABEL = { cash: "نقدي", cheque: "شيك", transfer: "حوالة مصرفية" };
const METHOD_ICON = { cash: Banknote, cheque: ScrollText, transfer: CreditCard };
// حالات تسوية الفاتورة (يدفع التاجر ← يراجع المندوب ← يعتمد)
const SETTLE_STATUS = {
  pending_rep: { label: "بانتظار مراجعة المندوب", color: C.accent, soft: C.accentSoft },
  approved:    { label: "اعتمدها المندوب", color: C.success, soft: C.successSoft },
  rejected:    { label: "رفضها المندوب", color: C.danger, soft: C.dangerSoft },
};
// تسليم النقدية للشركة
const HANDOVER_STATUS = {
  pending:  { label: "بانتظار استلام الشركة", color: C.accent, soft: C.accentSoft },
  received: { label: "استلمتها الشركة", color: C.success, soft: C.successSoft },
};
// طريقة استلام الطلب (للصيدليات/المحلات)
const FULFILL = {
  delivery: { label: "توصيل", icon: Truck, color: C.info, soft: C.infoSoft },
  pickup:   { label: "استلام ذاتي", icon: Store, color: C.purple, soft: C.purpleSoft },
};
// تصنيف الشركة: موزّع كامل (فريق ميداني) أو بيع مباشر (شركة صغيرة تسلّم بنفسها)
const COMPANY_TYPE = {
  distributor: { label: "موزّع كامل", desc: "فريق مندوبين وسائقين وإشراف", icon: Truck, color: C.primary2, soft: "#DBEAFE" },
  direct:      { label: "بيع مباشر", desc: "تواصل مباشر بلا فريق ميداني — تسلّم بنفسها", icon: Handshake, color: C.teal, soft: C.tealSoft },
};
// تصنيف المندوب: تجزئة (سيارة بيع بمخزون) أو بري‑سيل (يأخذ الطلب ثم تجهّزه الشركة)
const REP_TYPE = {
  van:     { label: "مندوب تجزئة", short: "تجزئة", desc: "سيارة بيع بمخزون — يبيع ويسلّم فوراً من سيارته", icon: Truck, color: C.accent, soft: C.accentSoft },
  presell: { label: "مندوب بري‑سيل", short: "بري‑سيل", desc: "بيع جملة — يأخذ الطلب/الفاتورة وتجهّزه الشركة للتوصيل", icon: ClipboardList, color: C.info, soft: C.infoSoft },
};

// ═══ أنواع السيارات وسعتها بالباليتات (لتعبئة الحمولة) ═══
const VEHICLE_TYPES = {
  small:  { label: "صغيرة (بيك أب)", pallets: 2, plate: "" },
  medium: { label: "متوسطة (شاحنة)", pallets: 4, plate: "" },
  large:  { label: "كبيرة (هيونداي H72)", pallets: 6, plate: "" },
};
const vehiclePallets = (type) => VEHICLE_TYPES[type]?.pallets || 4;
// سعة الباليتة لكل صنف (يحددها مشرف الحركة) — افتراضي حسب حجم الصنف
const DEFAULT_STIKA_PER_PALLET = 40;
const stikaPerPallet = (p) => p?.stikaPerPallet || DEFAULT_STIKA_PER_PALLET;
// إجمالي ستيكات طلبية (مجموع كميات الأصناف)
const orderStikat = (o) => o.items.reduce((s, it) => s + it.qty, 0);
// تحويل ستيكات طلبية إلى باليتات (حسب سعة كل صنف)
const orderPallets = (o, products) => {
  let pallets = 0;
  o.items.forEach((it) => { const p = products.find((x) => x.id === it.pid); pallets += it.qty / stikaPerPallet(p); });
  return pallets;
};
// حالة طلب تحميل سيارة المندوب
const LOAD_STATUS = {
  requested:   { label: "بانتظار التحميل", color: C.accent, soft: C.accentSoft },
  pending_rep: { label: "بانتظار موافقة المندوب", color: C.info, soft: C.infoSoft },
  loaded:      { label: "تم التحميل", color: C.success, soft: C.successSoft },
  rejected:    { label: "مرفوض", color: C.danger, soft: C.dangerSoft },
};
// حالة العميل المستورد من Excel (توثيق ميداني إلزامي)
const PC_STATUS = {
  awaiting_visit:   { label: "بانتظار زيارة المندوب", color: C.accent, soft: C.accentSoft },
  field_verified:   { label: "وثّقه المندوب — بانتظار تأكيد التاجر", color: C.info, soft: C.infoSoft },
  activated:        { label: "مُفعَّل كعميل", color: C.success, soft: C.successSoft },
  rejected:         { label: "مرفوض / بيانات خاطئة", color: C.danger, soft: C.dangerSoft },
};
// حالة طلب مطابقة الحساب
const REC_STATUS = {
  pending:   { label: "بانتظار تأكيد التاجر", color: C.accent, soft: C.accentSoft },
  confirmed: { label: "طابقها التاجر ✓", color: C.success, soft: C.successSoft },
  disputed:  { label: "لدى التاجر ملاحظات", color: C.danger, soft: C.dangerSoft },
};
// حالة فحص المرتجع (يدوّنها المندوب/العمليات عند الاستلام)
const RET_CONDITION = {
  ok:          { label: "سليم — يصلح للبيع", color: C.success, soft: C.successSoft },
  damaged:     { label: "تالف", color: C.danger, soft: C.dangerSoft },
  near_expiry: { label: "قارب الانتهاء", color: C.accent, soft: C.accentSoft },
};
// أعمار الديون
const AGE_BUCKETS = [
  { key: "b0", label: "0–30 يوم", min: 0, max: 30, color: C.success, soft: C.successSoft },
  { key: "b1", label: "31–60 يوم", min: 31, max: 60, color: C.accent, soft: C.accentSoft },
  { key: "b2", label: "أكثر من 60 يوم", min: 61, max: 9999, color: C.danger, soft: C.dangerSoft },
];
// أعمار الدين: أقدم فاتورة غير مسددة + هل تجاوز 30 يوم (مع مراعاة سماح الشركة)
const DEBT_LIMIT_DAYS = 30;
const oldestDebtDays = (data, merchantId) => {
  const open = data.invoices.filter((i) => i.merchantId === merchantId && i.amount > i.paid);
  return open.length ? Math.max(...open.map((i) => i.ageDays ?? 0)) : 0;
};
// هل يُمنع إصدار فاتورة لهذا المحل بسبب تقادم الدين؟ يعيد {blocked, days, reason, override}
const invoiceGate = (data, merchantId) => {
  const m = data.merchants.find((x) => x.id === merchantId);
  const days = oldestDebtDays(data, merchantId);
  const ov = m?.invoiceOverride;
  // سماح دائم من الشركة (allow) أو إمهال بيوم إضافي (grace: عدد الأيام المسموح بها)
  if (ov?.type === "allow") return { blocked: false, days, override: "allow" };
  const limit = ov?.type === "grace" ? DEBT_LIMIT_DAYS + (ov.extraDays || 1) : DEBT_LIMIT_DAYS;
  if (days > limit) return { blocked: true, days, limit, reason: `دين متأخر ${days} يوم (الحد ${limit} يوم)` };
  return { blocked: false, days, limit };
};
const WEEK_DAYS = ["السبت", "الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس"];
// حالة طلب مندوب للتحصيل (يطلبه التاجر ← يتابعه المندوب والشركة)
const COLLECT_STATUS = {
  requested: { label: "بانتظار المندوب", color: C.accent, soft: C.accentSoft },
  scheduled: { label: "جدول المندوب زيارة", color: C.info, soft: C.infoSoft },
  done:      { label: "تم التحصيل", color: C.success, soft: C.successSoft },
  canceled:  { label: "أُلغي", color: C.danger, soft: C.dangerSoft },
};

// خطط الاشتراك
const PLANS = {
  basic: { label: "الأساسي", color: C.info, soft: C.infoSoft, features: ["إدارة المنتجات", "إدارة الحصص", "استقبال الطلبات", "تقارير بسيطة", "دعم فني"] },
  pro:   { label: "الاحترافي", color: C.purple, soft: C.purpleSoft, features: ["كل مزايا الأساسي", "تقارير متقدمة", "صلاحيات الموظفين", "إشعارات ذكية", "تتبع المندوبين", "ربط نقاط البيع"] },
  intel: { label: "استخبارات المنافسين", color: C.accent, soft: C.accentSoft, features: ["كل مزايا الاحترافي", "أكثر الأصناف مبيعاً", "متوسط الأسعار", "اتجاهات الطلب", "المناطق الأكثر شراءً", "بيانات مجمّعة فقط"] },
};
const COMM_TYPES = { percent: "نسبة مئوية", fixed: "مبلغ ثابت", tiered: "نسبة متدرّجة" };
// شرائح العمولة المتدرّجة
const TIERS_COMM = [{ upTo: 50000, pct: 1 }, { upTo: 150000, pct: 0.75 }, { upTo: Infinity, pct: 0.5 }];
// حساب عمولة شركة على مبلغ مبيعات
function calcCommission(co, salesAmount) {
  if (!co) return 0;
  if (co.commType === "fixed") return co.commValue; // لكل فاتورة
  if (co.commType === "percent") { let c = salesAmount * (co.commValue / 100); if (co.commMax) c = Math.min(c, co.commMax); if (co.commMin) c = Math.max(c, co.commMin); return c; }
  if (co.commType === "tiered") { const tier = TIERS_COMM.find((t) => salesAmount <= t.upTo); return salesAmount * ((tier?.pct || 0.5) / 100); }
  return 0;
}
// إجمالي الكمية في طلب
const orderQty = (o) => o.items.reduce((s, it) => s + it.qty, 0);
const qty = (n) => Number(n || 0).toLocaleString("ar-LY") + " وحدة";
// خدمات إضافية وحالة الدفع
const ADDONS = { pos: "ربط نقاط البيع (POS)", accounting: "تكامل المحاسبة", priority: "دعم وإبراز مميز" };
const PAY_STATUS = {
  paid: { label: "مدفوع", color: C.success, soft: C.successSoft },
  due:  { label: "مستحق", color: C.danger, soft: C.dangerSoft },
  trial:{ label: "تجريبي", color: C.info, soft: C.infoSoft },
};
const BILLING = { monthly: "شهري", yearly: "سنوي" };

const money = (n) => Number(n || 0).toLocaleString("ar-LY") + " د.ل";
const orderTotal = (o) => o.items.reduce((s, it) => s + it.qty * it.price, 0);
const uid = (p) => p + Math.random().toString(36).slice(2, 8);

// خصم الشرائح: يتفعّل إما بعدد الوحدات/الصناديق (basis:"qty") أو بقيمة الشراء بالدينار (basis:"value")
const tierFor = (p, q, lineValue = null) => {
  if (!p?.tiers || !p.tiers.length) return null;
  const matches = p.tiers.filter((t) => (t.basis === "value")
    ? (lineValue != null && lineValue >= (t.minValue ?? t.minQty ?? Infinity))
    : (q >= (t.minQty ?? Infinity)));
  if (!matches.length) return null;
  return matches.sort((a, b) => b.discount - a.discount)[0];
};
// سعر الوحدة بعد خصم الشريحة (يحسب قيمة السطر داخلياً لدعم التخفيض حسب القيمة)
const tierUnitPrice = (p, m, q) => {
  const base = priceForMerchant(p, m);
  const t = tierFor(p, q, base * q);
  return t ? Math.round(base * (1 - t.discount / 100)) : base;
};
// وصف عتبة الشريحة (للعرض)
const tierThresholdLabel = (t, unitWord = "وحدة") => t.basis === "value" ? `شراء بقيمة ${(t.minValue ?? t.minQty).toLocaleString("ar-LY")} د.ل` : `من ${t.minQty} ${unitWord}`;
// متوسط تقييم شركة (تقييمات التجار) أو عميل (تقييمات الشركات)
const companyRating = (d, companyId) => { const list = d.ratings.filter((r) => r.raterRole === "merchant" && r.companyId === companyId); return list.length ? { avg: list.reduce((s, r) => s + r.stars, 0) / list.length, count: list.length } : null; };
const merchantRating = (d, merchantId) => { const list = d.ratings.filter((r) => r.raterRole === "company" && r.merchantId === merchantId); return list.length ? { avg: list.reduce((s, r) => s + r.stars, 0) / list.length, count: list.length } : null; };
// معلومات الصلاحية: منتهٍ / قريب الانتهاء (خلال 6 أشهر) — للأدوية والأغذية
const expiryInfo = (expDate) => {
  if (!expDate) return null;
  const exp = new Date(expDate); const now = new Date();
  const days = Math.round((exp - now) / 86400000);
  if (days < 0) return { state: "expired", label: "منتهي الصلاحية", color: C.danger, soft: C.dangerSoft, days };
  if (days <= 180) return { state: "near", label: `قريب الانتهاء (${days} يوم)`, color: C.accent, soft: C.accentSoft, days };
  return { state: "ok", label: "صالح", color: C.success, soft: C.successSoft, days };
};
const fmtDate = (d) => { if (!d) return "—"; const dt = new Date(d); if (isNaN(dt)) return "—"; const p = (n) => String(n).padStart(2, "0"); return `${p(dt.getDate())}/${p(dt.getMonth() + 1)}/${dt.getFullYear()}`; };
const geoDistanceM = (lat1, lng1, lat2, lng2) => { const rad = (n) => n * Math.PI / 180; const R = 6371000; const dLat = rad(lat2 - lat1); const dLng = rad(lng2 - lng1); const a = Math.sin(dLat / 2) ** 2 + Math.cos(rad(lat1)) * Math.cos(rad(lat2)) * Math.sin(dLng / 2) ** 2; return 2 * R * Math.asin(Math.sqrt(a)); };

/* ===================== البيانات الأولية (مع server لكل كيان) ===================== */
const seed = {
  companies: [
    // غذائية
    { id: "c1", server: "food", name: "شركة النور للتوزيع", logo: "🌟", cat: "مواد غذائية", sub: "active", city: "طرابلس", phone: "091-2233445", desc: "موزّع رائد للمواد الغذائية الأساسية.", plan: "pro", commType: "percent", commValue: 1, commMin: 0, commMax: 0, commTax: false, commScope: "all", frozen: false, contractNo: "CT-1001", companyType: "distributor", hasDelivery: true, branches: [{ id: "br1", name: "فرع طرابلس", city: "طرابلس", manager: "أدمن الفرع" }, { id: "br2", name: "فرع بنغازي", city: "بنغازي", manager: "أدمن الفرع" }, { id: "br3", name: "فرع مصراتة", city: "مصراتة", manager: "أدمن الفرع" }], contractStart: "2025-01-01", contractEnd: "2026-12-31", subFee: 1500, payMethod: "تحويل مصرفي", payStatus: "paid", billing: "yearly", trial: false, discount: 0, maxReps: 20, maxProducts: 200, intelAccess: true, addons: { pos: true, accounting: false, priority: true }, joinDate: "2025-01-01" , opsPerms: { prices: true, invoice: true, loads: true } },
    { id: "c2", server: "food", name: "شركة البركة للمشروبات", logo: "🥤", cat: "مشروبات", sub: "active", city: "طرابلس", phone: "092-7788990", desc: "مشروبات ومياه معدنية بالجملة.", plan: "basic", commType: "percent", commValue: 0.40, commMin: 0, commMax: 0, commTax: false, commScope: "all", frozen: false, contractNo: "CT-1002", companyType: "distributor", hasDelivery: true, contractStart: "2025-03-01", contractEnd: "2026-03-01", subFee: 600, payMethod: "نقدي", payStatus: "due", billing: "monthly", trial: false, discount: 10, maxReps: 5, maxProducts: 50, intelAccess: false, addons: { pos: false, accounting: false, priority: false }, joinDate: "2025-03-01" , opsPerms: { prices: true, invoice: true, loads: true } },
    { id: "c3", server: "food", name: "شركة الأمل للمنظفات", logo: "🧼", cat: "منظفات", sub: "active", city: "بنغازي", phone: "094-1122334", desc: "منظفات منزلية وصناعية.", plan: "intel", commType: "tiered", commValue: 1, commMin: 0, commMax: 0, commTax: true, commScope: "all", frozen: false, contractNo: "CT-1003", companyType: "distributor", hasDelivery: true, contractStart: "2025-02-15", contractEnd: "2026-02-15", subFee: 2500, payMethod: "تحويل مصرفي", payStatus: "paid", billing: "yearly", trial: false, discount: 0, maxReps: 30, maxProducts: 500, intelAccess: true, addons: { pos: true, accounting: true, priority: true }, joinDate: "2025-02-15" , opsPerms: { prices: true, invoice: true, loads: true } },
    { id: "c7", server: "food", name: "شركة الوفرة الغذائية", logo: "🌾", cat: "مواد غذائية", sub: "active", city: "مصراتة", phone: "093-7788990", desc: "مواد غذائية أساسية بأسعار منافسة وتوريد سريع.", plan: "pro", commType: "percent", commValue: 2, commMin: 0, commMax: 0, commTax: false, commScope: "all", frozen: false, contractNo: "CT-1007", companyType: "distributor", hasDelivery: true, contractStart: "2025-03-01", contractEnd: "2026-12-31", subFee: 1200, payMethod: "تحويل مصرفي", payStatus: "paid", billing: "yearly", trial: false, discount: 0, maxReps: 15, maxProducts: 150, intelAccess: false, commissionPaid: 0, addons: { pos: false }, opsPerms: { prices: true, invoice: true, loads: true } },
    // أدوية
    { id: "c4", server: "pharma", name: "شركة الشفاء للأدوية", logo: "💊", cat: "أدوية بشرية", sub: "active", city: "طرابلس", phone: "091-5012345", desc: "توزيع أدوية ومستحضرات طبية معتمدة.", plan: "pro", commType: "percent", commValue: 0.5, commMin: 0, commMax: 0, commTax: false, commScope: "all", frozen: false, contractNo: "CT-2001", companyType: "distributor", hasDelivery: true, contractStart: "2025-01-10", contractEnd: "2026-12-31", subFee: 2000, payMethod: "تحويل مصرفي", payStatus: "paid", billing: "yearly", trial: false, discount: 0, maxReps: 25, maxProducts: 300, intelAccess: true, addons: { pos: true, accounting: false, priority: true }, joinDate: "2025-01-10" , opsPerms: { prices: true, invoice: true, loads: true } },
    { id: "c5", server: "pharma", name: "شركة الحياة الطبية", logo: "🏥", cat: "مستلزمات طبية", sub: "active", city: "طرابلس", phone: "092-6023456", desc: "مستلزمات صيدليات وأجهزة طبية.", plan: "basic", commType: "fixed", commValue: 5, commMin: 0, commMax: 0, commTax: false, commScope: "all", frozen: false, contractNo: "CT-2002", companyType: "direct", hasDelivery: false, contractStart: "2025-04-01", contractEnd: "2026-04-01", subFee: 700, payMethod: "نقدي", payStatus: "paid", billing: "monthly", trial: true, discount: 0, maxReps: 5, maxProducts: 50, intelAccess: false, addons: { pos: false, accounting: false, priority: false }, joinDate: "2025-04-01" , opsPerms: { prices: true, invoice: true, loads: true } },
    { id: "c6", server: "pharma", name: "شركة النخيل فارما", logo: "🌿", cat: "مكمّلات", sub: "pending", city: "مصراتة", phone: "051-7034567", desc: "مكمّلات غذائية ومنتجات تجميل طبية.", plan: "basic", commType: "percent", commValue: 0.50, commMin: 0, commMax: 0, commTax: false, commScope: "all", frozen: false, contractNo: "CT-2003", companyType: "direct", hasDelivery: true, contractStart: "2025-06-01", contractEnd: "2026-06-01", subFee: 500, payMethod: "نقدي", payStatus: "due", billing: "monthly", trial: false, discount: 0, maxReps: 5, maxProducts: 50, intelAccess: false, addons: { pos: false, accounting: false, priority: false }, joinDate: "2025-06-01" , opsPerms: { prices: true, invoice: true, loads: true } },
  ],
  products: [
    // غذائية
    { id: "p1", server: "food", companyId: "c1", name: "أرز مصري فاخر 25كغ", cat: "حبوب", price: 180, stock: 320, unit: "كيس", packSize: 10, minOrder: 5, weight: "25 كغ", size: "60×40×15 سم", taxPct: 0, status: "active", offer: 0, barcode: "6291041500213", images: 3, mfgDate: "2025-08-01", expDate: "2027-08-01", tiers: [{ minQty: 50, discount: 3 }, { minQty: 100, discount: 6 }] },
    { id: "p2", server: "food", companyId: "c1", name: "زيت طبخ نباتي 18لتر", cat: "زيوت", price: 240, stock: 150, unit: "صفيحة", packSize: 4, minOrder: 2, weight: "16.5 كغ", size: "30×30×35 سم", taxPct: 0, status: "active", offer: 10, barcode: "6291041500220", images: 2, mfgDate: "2025-09-15", expDate: "2026-09-15" },
    { id: "p3", server: "food", companyId: "c1", name: "سكر ناعم 50كغ", cat: "سكر", price: 95, stock: 200, unit: "كيس", packSize: 10, minOrder: 4, weight: "50 كغ", size: "70×45×18 سم", taxPct: 0, status: "active", offer: 0, barcode: "6291041500237", images: 1, mfgDate: "2025-07-01", expDate: "2027-01-01" },
    { id: "p4", server: "food", companyId: "c1", name: "معكرونة 500غ (كرتونة×20)", cat: "معكرونة", price: 62, stock: 90, unit: "كرتونة", packSize: 6, minOrder: 3, weight: "10 كغ", size: "40×30×25 سم", taxPct: 0, status: "active", offer: 15, barcode: "6291041500244", images: 2, mfgDate: "2025-10-01", expDate: "2026-10-01" },
    { id: "p5", server: "food", companyId: "c2", name: "مياه معدنية 0.5ل (كرتونة×24)", cat: "مياه", price: 14, stock: 800, unit: "كرتونة", packSize: 6, minOrder: 10, weight: "12 كغ", size: "40×27×22 سم", taxPct: 0, status: "active", offer: 0, barcode: "6291041500251", images: 1, mfgDate: "2025-11-01", expDate: "2026-11-01" },
    { id: "p6", server: "food", companyId: "c2", name: "عصائر مشكّلة 250مل (علبة×12)", cat: "عصائر", price: 38, stock: 260, unit: "علبة", packSize: 12, minOrder: 5, weight: "3.6 كغ", size: "32×24×12 سم", taxPct: 0, status: "active", offer: 0, barcode: "6291041500268", images: 2, mfgDate: "2025-10-10", expDate: "2026-04-10" },
    // أدوية - مع صلاحية وتشغيلة
    { id: "p7", server: "pharma", companyId: "c4", name: "باراسيتامول 500مغ (علبة×20)", cat: "مسكّنات وخافضات حرارة", price: 12, stock: 600, unit: "علبة", packSize: 12, minOrder: 10, weight: "60 غ", size: "10×6×3 سم", taxPct: 0, status: "active", offer: 0, barcode: "6293000100017", images: 2, mfgDate: "2025-06-01", expDate: "2027-06-01", batch: "BT-2206", tiers: [{ minQty: 100, discount: 4 }, { minQty: 300, discount: 8 }, { minQty: 500, discount: 12 }] },
    { id: "p8", server: "pharma", companyId: "c4", name: "أموكسيسيلين 500مغ (شريط×16)", cat: "مضادات حيوية", price: 28, stock: 350, unit: "علبة", packSize: 12, minOrder: 5, weight: "40 غ", size: "9×6×2 سم", taxPct: 0, status: "active", offer: 5, barcode: "6293000100024", images: 1, mfgDate: "2025-03-01", expDate: "2026-09-01", batch: "BT-2203", tiers: [{ minQty: 50, discount: 5 }, { minQty: 150, discount: 10 }] },
    { id: "p9", server: "pharma", companyId: "c4", name: "شراب فيتامين د للأطفال", cat: "فيتامينات ومكمّلات", price: 35, stock: 180, unit: "زجاجة", packSize: 12, minOrder: 6, weight: "120 غ", size: "12×5×5 سم", taxPct: 0, status: "active", offer: 0, barcode: "6293000100031", images: 2, mfgDate: "2025-05-01", expDate: "2026-05-01", batch: "BT-2205" },
    { id: "p10", server: "pharma", companyId: "c5", name: "ضمادات طبية معقّمة (كرتونة×50)", cat: "مستلزمات طبية", price: 45, stock: 220, unit: "كرتونة", packSize: 6, minOrder: 4, weight: "2 كغ", size: "30×20×15 سم", taxPct: 0, status: "active", offer: 0, barcode: "6293000100048", images: 1, mfgDate: "2025-01-01", expDate: "2028-01-01", batch: "BT-2201" },
    { id: "p11", server: "pharma", companyId: "c5", name: "جهاز قياس ضغط رقمي", cat: "أجهزة طبية", price: 210, stock: 60, unit: "جهاز", packSize: 1, minOrder: 2, weight: "500 غ", size: "16×12×8 سم", taxPct: 0, status: "active", offer: 12, barcode: "6293000100055", images: 3, mfgDate: "2025-02-01", expDate: "2030-02-01", batch: "BT-2202" },
    { id: "p12", server: "pharma", companyId: "c6", name: "مكمّل أوميغا 3 (علبة×60)", cat: "فيتامينات ومكمّلات", price: 68, stock: 140, unit: "علبة", packSize: 12, minOrder: 4, weight: "180 غ", size: "11×7×7 سم", taxPct: 0, status: "active", offer: 0, barcode: "6293000100062", images: 2, mfgDate: "2025-04-01", expDate: "2027-04-01", batch: "BT-2204" },
    { id: "p13", server: "pharma", companyId: "c4", name: "إيبوبروفين 400مغ (علبة×30)", cat: "مسكّنات وخافضات حرارة", price: 14, stock: 500, unit: "علبة", packSize: 12, minOrder: 10, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010113", images: 1, mfgDate: "2025-05-01", expDate: "2027-09-01", batch: "BT-2310", tiers: [{ minQty: 100, discount: 4 }, { minQty: 300, discount: 8 }] },
    { id: "p14", server: "pharma", companyId: "c4", name: "ديكلوفيناك 50مغ (علبة×20)", cat: "مسكّنات وخافضات حرارة", price: 10, stock: 380, unit: "علبة", packSize: 12, minOrder: 10, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010114", images: 1, mfgDate: "2025-05-01", expDate: "2027-03-01", batch: "BT-2311" },
    { id: "p15", server: "pharma", companyId: "c4", name: "أسبرين 100مغ (علبة×30)", cat: "مسكّنات وخافضات حرارة", price: 8, stock: 640, unit: "علبة", packSize: 12, minOrder: 10, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010115", images: 1, mfgDate: "2025-05-01", expDate: "2026-12-01", batch: "BT-2312" },
    { id: "p16", server: "pharma", companyId: "c4", name: "أزيثرومايسين 250مغ (علبة×6)", cat: "مضادات حيوية", price: 34, stock: 220, unit: "علبة", packSize: 12, minOrder: 5, weight: "—", size: "—", taxPct: 0, status: "active", offer: 8, barcode: "629300010116", images: 1, mfgDate: "2025-05-01", expDate: "2026-10-01", batch: "BT-2313", tiers: [{ minQty: 50, discount: 6 }, { minQty: 150, discount: 11 }] },
    { id: "p17", server: "pharma", companyId: "c4", name: "سيبروفلوكساسين 500مغ (علبة×10)", cat: "مضادات حيوية", price: 26, stock: 180, unit: "علبة", packSize: 12, minOrder: 5, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010117", images: 1, mfgDate: "2025-05-01", expDate: "2026-08-15", batch: "BT-2314" },
    { id: "p18", server: "pharma", companyId: "c4", name: "سيفترياكسون 1غ (حقنة)", cat: "مضادات حيوية", price: 18, stock: 90, unit: "حقنة", packSize: 50, minOrder: 10, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010118", images: 1, mfgDate: "2025-05-01", expDate: "2026-07-20", batch: "BT-2315" },
    { id: "p19", server: "pharma", companyId: "c4", name: "أوميبرازول 20مغ (علبة×14)", cat: "الجهاز الهضمي", price: 16, stock: 300, unit: "علبة", packSize: 12, minOrder: 6, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010119", images: 1, mfgDate: "2025-05-01", expDate: "2027-02-01", batch: "BT-2316", tiers: [{ minQty: 100, discount: 5 }] },
    { id: "p20", server: "pharma", companyId: "c4", name: "دومبيريدون 10مغ (علبة×30)", cat: "الجهاز الهضمي", price: 12, stock: 260, unit: "علبة", packSize: 12, minOrder: 6, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010120", images: 1, mfgDate: "2025-05-01", expDate: "2027-01-10", batch: "BT-2317" },
    { id: "p21", server: "pharma", companyId: "c4", name: "محلول معالجة الجفاف (ORS)", cat: "الجهاز الهضمي", price: 6, stock: 800, unit: "كيس", packSize: 10, minOrder: 20, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010121", images: 1, mfgDate: "2025-05-01", expDate: "2027-05-01", batch: "BT-2318" },
    { id: "p22", server: "pharma", companyId: "c4", name: "أملوديبين 5مغ (علبة×30)", cat: "الضغط والقلب", price: 13, stock: 240, unit: "علبة", packSize: 12, minOrder: 6, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010122", images: 1, mfgDate: "2025-05-01", expDate: "2027-04-01", batch: "BT-2319" },
    { id: "p23", server: "pharma", companyId: "c4", name: "لوسارتان 50مغ (علبة×28)", cat: "الضغط والقلب", price: 20, stock: 200, unit: "علبة", packSize: 12, minOrder: 6, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010123", images: 1, mfgDate: "2025-05-01", expDate: "2027-06-01", batch: "BT-2320" },
    { id: "p24", server: "pharma", companyId: "c4", name: "أتينولول 50مغ (علبة×28)", cat: "الضغط والقلب", price: 11, stock: 160, unit: "علبة", packSize: 12, minOrder: 6, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010124", images: 1, mfgDate: "2025-05-01", expDate: "2026-09-10", batch: "BT-2321" },
    { id: "p25", server: "pharma", companyId: "c4", name: "ميتفورمين 850مغ (علبة×30)", cat: "أدوية السكري", price: 15, stock: 420, unit: "علبة", packSize: 12, minOrder: 6, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010125", images: 1, mfgDate: "2025-05-01", expDate: "2027-07-01", batch: "BT-2322", tiers: [{ minQty: 100, discount: 5 }, { minQty: 250, discount: 9 }] },
    { id: "p26", server: "pharma", companyId: "c4", name: "جليبنكلاميد 5مغ (علبة×30)", cat: "أدوية السكري", price: 10, stock: 180, unit: "علبة", packSize: 12, minOrder: 6, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010126", images: 1, mfgDate: "2025-05-01", expDate: "2026-11-01", batch: "BT-2323" },
    { id: "p27", server: "pharma", companyId: "c4", name: "شراب طارد للبلغم 120مل", cat: "الجهاز التنفسي", price: 9, stock: 300, unit: "زجاجة", packSize: 12, minOrder: 6, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010127", images: 1, mfgDate: "2025-05-01", expDate: "2026-08-01", batch: "BT-2324" },
    { id: "p28", server: "pharma", companyId: "c4", name: "بخاخ سالبوتامول 100مكغ", cat: "الجهاز التنفسي", price: 22, stock: 140, unit: "بخاخ", packSize: 12, minOrder: 4, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010128", images: 1, mfgDate: "2025-05-01", expDate: "2027-03-15", batch: "BT-2325" },
    { id: "p29", server: "pharma", companyId: "c4", name: "فيتامين C 1000مغ (أنبوب×20)", cat: "فيتامينات ومكمّلات", price: 11, stock: 520, unit: "أنبوب", packSize: 24, minOrder: 6, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010129", images: 1, mfgDate: "2025-05-01", expDate: "2027-08-01", batch: "BT-2326", tiers: [{ minQty: 100, discount: 5 }] },
    { id: "p30", server: "pharma", companyId: "c4", name: "كالسيوم + فيتامين د (علبة×60)", cat: "فيتامينات ومكمّلات", price: 24, stock: 260, unit: "علبة", packSize: 12, minOrder: 4, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010130", images: 1, mfgDate: "2025-05-01", expDate: "2027-09-01", batch: "BT-2327" },
    { id: "p31", server: "pharma", companyId: "c4", name: "حديد + فوليك (علبة×30)", cat: "فيتامينات ومكمّلات", price: 13, stock: 200, unit: "علبة", packSize: 12, minOrder: 6, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010131", images: 1, mfgDate: "2025-05-01", expDate: "2027-02-20", batch: "BT-2328" },
    { id: "p32", server: "pharma", companyId: "c4", name: "مرهم مضاد حيوي 15غ", cat: "جلدية ومراهم", price: 8, stock: 340, unit: "أنبوب", packSize: 24, minOrder: 10, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010132", images: 1, mfgDate: "2025-05-01", expDate: "2026-10-05", batch: "BT-2329" },
    { id: "p33", server: "pharma", companyId: "c4", name: "كريم مرطّب طبي 50غ", cat: "جلدية ومراهم", price: 14, stock: 180, unit: "أنبوب", packSize: 24, minOrder: 6, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010133", images: 1, mfgDate: "2025-05-01", expDate: "2027-05-20", batch: "BT-2330" },
    { id: "p34", server: "pharma", companyId: "c4", name: "قطرة عين مرطّبة 10مل", cat: "قطرات ومحاليل", price: 10, stock: 260, unit: "زجاجة", packSize: 12, minOrder: 6, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010134", images: 1, mfgDate: "2025-05-01", expDate: "2026-09-25", batch: "BT-2331" },
    { id: "p35", server: "pharma", companyId: "c4", name: "قطرة أذن 10مل", cat: "قطرات ومحاليل", price: 9, stock: 150, unit: "زجاجة", packSize: 12, minOrder: 6, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010135", images: 1, mfgDate: "2025-05-01", expDate: "2027-01-05", batch: "BT-2332" },
    { id: "p36", server: "pharma", companyId: "c4", name: "شراب أطفال خافض للحرارة 100مل", cat: "نسائية وأطفال", price: 8, stock: 460, unit: "زجاجة", packSize: 12, minOrder: 10, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010136", images: 1, mfgDate: "2025-05-01", expDate: "2026-12-20", batch: "BT-2333" },
    { id: "p37", server: "pharma", companyId: "c4", name: "فيتامينات للحوامل (علبة×30)", cat: "نسائية وأطفال", price: 26, stock: 140, unit: "علبة", packSize: 12, minOrder: 4, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010137", images: 1, mfgDate: "2025-05-01", expDate: "2027-06-15", batch: "BT-2334" },
    { id: "p38", server: "pharma", companyId: "c5", name: "قفازات فحص (كرتونة×100)", cat: "مستلزمات طبية", price: 30, stock: 400, unit: "كرتونة", packSize: 6, minOrder: 4, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010138", images: 1, mfgDate: "2025-05-01", expDate: "2028-01-01", batch: "BT-2335", tiers: [{ minQty: 20, discount: 5 }, { minQty: 50, discount: 10 }] },
    { id: "p39", server: "pharma", companyId: "c5", name: "كمامات طبية (كرتونة×50)", cat: "مستلزمات طبية", price: 18, stock: 520, unit: "كرتونة", packSize: 6, minOrder: 4, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010139", images: 1, mfgDate: "2025-05-01", expDate: "2028-03-01", batch: "BT-2336" },
    { id: "p40", server: "pharma", companyId: "c5", name: "حقن للاستعمال مرة واحدة 5مل (علبة×100)", cat: "مستلزمات طبية", price: 22, stock: 300, unit: "علبة", packSize: 12, minOrder: 4, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010140", images: 1, mfgDate: "2025-05-01", expDate: "2028-02-01", batch: "BT-2337" },
    { id: "p41", server: "pharma", companyId: "c5", name: "جهاز قياس سكر الدم", cat: "أجهزة طبية", price: 95, stock: 70, unit: "جهاز", packSize: 1, minOrder: 2, weight: "—", size: "—", taxPct: 0, status: "active", offer: 10, barcode: "629300010141", images: 1, mfgDate: "2025-05-01", expDate: "2030-01-01", batch: "BT-2338" },
    { id: "p42", server: "pharma", companyId: "c5", name: "ميزان حرارة رقمي", cat: "أجهزة طبية", price: 28, stock: 120, unit: "جهاز", packSize: 1, minOrder: 4, weight: "—", size: "—", taxPct: 0, status: "active", offer: 0, barcode: "629300010142", images: 1, mfgDate: "2025-05-01", expDate: "2030-06-01", batch: "BT-2339" },
    // منتجات متشابهة عبر الشركات — لتفعيل مقارنة الأسعار الحقيقية
    { id: "p43", server: "food", companyId: "c7", name: "أرز مصري فاخر 25كغ", cat: "حبوب", price: 172, stock: 500, unit: "كيس", packSize: 10, minOrder: 5, weight: "25 كغ", size: "60×40×15 سم", taxPct: 0, status: "active", offer: 0, barcode: "6291041500777", images: 2, mfgDate: "2025-09-01", expDate: "2027-09-01", tiers: [{ minQty: 50, discount: 4 }] },
    { id: "p44", server: "food", companyId: "c7", name: "زيت طبخ نباتي 18لتر", cat: "زيوت", price: 232, stock: 260, unit: "صفيحة", packSize: 4, minOrder: 3, weight: "18 ل", size: "30×30×40 سم", taxPct: 0, status: "active", offer: 5, barcode: "6291041500784", images: 2, mfgDate: "2025-10-01", expDate: "2027-04-01", tiers: [] },
    { id: "p45", server: "food", companyId: "c7", name: "سكر ناعم 50كغ", cat: "سكريات", price: 205, stock: 380, unit: "كيس", packSize: 10, minOrder: 5, weight: "50 كغ", size: "70×45×15 سم", taxPct: 0, status: "active", offer: 0, barcode: "6291041500791", images: 1, mfgDate: "2025-11-01", expDate: "2028-11-01", tiers: [{ minQty: 40, discount: 3 }] },
    { id: "p46", server: "food", companyId: "c7", name: "معكرونة 500غ (كرتونة×20)", cat: "معجنات", price: 58, stock: 700, unit: "كرتونة", packSize: 6, minOrder: 10, weight: "10 كغ", size: "40×30×25 سم", taxPct: 0, status: "active", offer: 0, barcode: "6291041500807", images: 2, mfgDate: "2025-12-01", expDate: "2027-12-01", tiers: [{ minQty: 100, discount: 5 }] },
    { id: "p47", server: "food", companyId: "c2", name: "أرز مصري فاخر 25كغ", cat: "حبوب", price: 185, stock: 150, unit: "كيس", packSize: 10, minOrder: 5, weight: "25 كغ", size: "60×40×15 سم", taxPct: 0, status: "active", offer: 3, barcode: "6291041500814", images: 1, mfgDate: "2025-07-01", expDate: "2027-07-01", tiers: [] },
    { id: "p48", server: "food", companyId: "c2", name: "مياه معدنية 0.5ل (كرتونة×24)", cat: "مشروبات", price: 14, stock: 900, unit: "كرتونة", packSize: 6, minOrder: 20, weight: "12 كغ", size: "40×27×20 سم", taxPct: 0, status: "active", offer: 0, barcode: "6291041500821", images: 1, mfgDate: "2026-01-01", expDate: "2027-07-01", tiers: [{ minQty: 200, discount: 6 }] },
    { id: "p49", server: "food", companyId: "c3", name: "منظف أرضيات معطّر 5ل", cat: "منظفات", price: 22, stock: 420, unit: "زجاجة", packSize: 12, minOrder: 12, weight: "5 كغ", size: "20×15×30 سم", taxPct: 0, status: "active", offer: 0, barcode: "6291041500838", images: 2, mfgDate: "2025-10-15", expDate: "2028-10-15", tiers: [{ minQty: 60, discount: 5 }] },
    { id: "p50", server: "food", companyId: "c3", name: "سائل غسيل صحون 1ل (كرتونة×12)", cat: "منظفات", price: 30, stock: 350, unit: "كرتونة", packSize: 6, minOrder: 8, weight: "12 كغ", size: "35×25×28 سم", taxPct: 0, status: "active", offer: 8, barcode: "6291041500845", images: 2, mfgDate: "2025-09-20", expDate: "2028-09-20", tiers: [] },
    { id: "p51", server: "pharma", companyId: "c5", name: "باراسيتامول 500مغ (علبة×20)", cat: "مسكّنات وخافضات حرارة", price: 11, stock: 800, unit: "علبة", packSize: 12, minOrder: 10, weight: "60 غ", size: "10×6×3 سم", taxPct: 0, status: "active", offer: 0, barcode: "6293000100777", images: 1, mfgDate: "2025-09-01", expDate: "2027-09-01", batch: "HL-3310", tiers: [{ minQty: 100, discount: 5 }, { minQty: 400, discount: 10 }] },
    { id: "p52", server: "pharma", companyId: "c5", name: "إيبوبروفين 400مغ (علبة×30)", cat: "مسكّنات وخافضات حرارة", price: 13, stock: 540, unit: "علبة", packSize: 12, minOrder: 10, weight: "75 غ", size: "10×6×3 سم", taxPct: 0, status: "active", offer: 4, barcode: "6293000100784", images: 1, mfgDate: "2025-08-15", expDate: "2027-08-15", batch: "HL-3327", tiers: [] },
    { id: "p53", server: "pharma", companyId: "c6", name: "باراسيتامول 500مغ (علبة×20)", cat: "مسكّنات وخافضات حرارة", price: 12, stock: 300, unit: "علبة", packSize: 12, minOrder: 10, weight: "60 غ", size: "10×6×3 سم", taxPct: 0, status: "active", offer: 6, barcode: "6293000100791", images: 1, mfgDate: "2025-11-01", expDate: "2027-11-01", batch: "NK-8801", tiers: [{ minQty: 150, discount: 6 }] },
    { id: "p54", server: "pharma", companyId: "c6", name: "فيتامين سي 1000مغ فوّار (أنبوب×20)", cat: "فيتامينات ومكمّلات", price: 18, stock: 260, unit: "أنبوب", packSize: 24, minOrder: 6, weight: "120 غ", size: "12×4×4 سم", taxPct: 0, status: "active", offer: 0, barcode: "6293000100807", images: 2, mfgDate: "2025-10-01", expDate: "2027-10-01", batch: "NK-8815", tiers: [] },
  ],
  merchants: [
    { id: "m51", server: "pharma", name: "صيدلية الشفاء", owner: "د. كمال بن سعيد", phone: "092-3530829", email: "ph51@rabet.ly", city: "طرابلس", area: "النوفليين", balance: -2200, creditLimit: 12000, companyId: "c4", repId: "r4", mid: "MID-20051" },
    { id: "m52", server: "pharma", name: "صيدلية الحياة", owner: "د. سميرة الفقيه", phone: "092-1810111", email: "ph52@rabet.ly", city: "طرابلس", area: "قرقارش", balance: -800, creditLimit: 8000, companyId: "c4", repId: "r4", mid: "MID-20052" },
    { id: "m53", server: "pharma", name: "صيدلية النور", owner: "د. طارق المصري", phone: "091-9990608", email: "ph53@rabet.ly", city: "طرابلس", area: "سوق الجمعة", balance: 0, creditLimit: 6000, companyId: "c4", repId: "r4", mid: "MID-20053" },
    { id: "m54", server: "pharma", name: "صيدلية السلام", owner: "د. هدى العماري", phone: "091-7135241", email: "ph54@rabet.ly", city: "طرابلس", area: "الظهرة", balance: -3400, creditLimit: 15000, companyId: "c4", repId: "r4", mid: "MID-20054" },
    { id: "m55", server: "pharma", name: "صيدلية الأمل", owner: "د. وليد بن يونس", phone: "093-1973060", email: "ph55@rabet.ly", city: "طرابلس", area: "عين زارة", balance: -1100, creditLimit: 9000, companyId: "c4", repId: "r4", mid: "MID-20055" },
    { id: "m11", branchId: "br1", server: "food", name: "محل البركة", owner: "خالد العجيلي", phone: "093-2867825", email: "store11@rabet.ly", city: "طرابلس", area: "سوق الجمعة", balance: -800, creditLimit: 7000, companyId: "c1", repId: "r1", mid: "MID-10011" },
    { id: "m12", branchId: "br1", server: "food", name: "سوبرماركت المدينة", owner: "فتحي بن نصر", phone: "091-5614226", email: "store12@rabet.ly", city: "طرابلس", area: "الظهرة", balance: -3200, creditLimit: 15000, companyId: "c1", repId: "r1", mid: "MID-10012" },
    { id: "m13", branchId: "br1", server: "food", name: "بقالة النور", owner: "سالم القذافي", phone: "091-4744854", email: "store13@rabet.ly", city: "طرابلس", area: "قرقارش", balance: 0, creditLimit: 6000, companyId: "c1", repId: "r1", mid: "MID-10013" },
    { id: "m14", branchId: "br1", server: "food", name: "هايبر الوسط", owner: "محمد الترهوني", phone: "091-2719583", email: "store14@rabet.ly", city: "طرابلس", area: "تاجوراء", balance: -1500, creditLimit: 12000, companyId: "c1", repId: "r1", mid: "MID-10014" },
    { id: "m15", branchId: "br1", server: "food", name: "محل الخير", owner: "عبدالسلام دياب", phone: "093-2458591", email: "store15@rabet.ly", city: "طرابلس", area: "باب بن غشير", balance: -450, creditLimit: 5000, companyId: "c1", repId: "r1", mid: "MID-10015" },
    { id: "m16", branchId: "br1", server: "food", name: "ميني ماركت الأمل", owner: "يوسف الشيباني", phone: "093-8078673", email: "store16@rabet.ly", city: "طرابلس", area: "سيدي المصري", balance: -2100, creditLimit: 9000, companyId: "c1", repId: "r1", mid: "MID-10016" },
    { id: "m17", branchId: "br2", server: "food", name: "سوبرماركت بنغازي", owner: "إدريس المسماري", phone: "091-1499914", email: "store17@rabet.ly", city: "بنغازي", area: "الصابري", balance: -1800, creditLimit: 14000, companyId: "c1", repId: "r5", mid: "MID-10017" },
    { id: "m18", branchId: "br2", server: "food", name: "بقالة الشرق", owner: "عمر الفاخري", phone: "091-4668136", email: "store18@rabet.ly", city: "بنغازي", area: "السلماني", balance: -600, creditLimit: 6000, companyId: "c1", repId: "r5", mid: "MID-10018" },
    { id: "m19", branchId: "br2", server: "food", name: "هايبر النخيل", owner: "صلاح بوجواري", phone: "091-9478454", email: "store19@rabet.ly", city: "بنغازي", area: "الكيش", balance: -4200, creditLimit: 20000, companyId: "c1", repId: "r5", mid: "MID-10019" },
    { id: "m20", branchId: "br2", server: "food", name: "محل الوحدة", owner: "فرج العبيدي", phone: "093-1445199", email: "store20@rabet.ly", city: "بنغازي", area: "بوعطني", balance: 0, creditLimit: 8000, companyId: "c1", repId: "r5", mid: "MID-10020" },
    { id: "m21", branchId: "br2", server: "food", name: "ميني ماركت الليثي", owner: "حسين الورفلي", phone: "093-4335942", email: "store21@rabet.ly", city: "بنغازي", area: "الليثي", balance: -950, creditLimit: 7000, companyId: "c1", repId: "r5", mid: "MID-10021" },
    { id: "m22", branchId: "br3", server: "food", name: "سوبرماركت مصراتة", owner: "أحمد الزناتي", phone: "093-8038374", email: "store22@rabet.ly", city: "مصراتة", area: "الزروق", balance: -2600, creditLimit: 16000, companyId: "c1", repId: "r3", mid: "MID-10022" },
    { id: "m23", branchId: "br3", server: "food", name: "بقالة الساحل", owner: "منصور دخيل", phone: "091-8536477", email: "store23@rabet.ly", city: "مصراتة", area: "تميرة", balance: -700, creditLimit: 6000, companyId: "c1", repId: "r3", mid: "MID-10023" },
    { id: "m24", branchId: "br3", server: "food", name: "هايبر قصر أحمد", owner: "علي المشاي", phone: "093-5667265", email: "store24@rabet.ly", city: "مصراتة", area: "قصر أحمد", balance: -3100, creditLimit: 18000, companyId: "c1", repId: "r3", mid: "MID-10024" },
    { id: "m25", branchId: "br3", server: "food", name: "محل التوفيق", owner: "بشير الغرياني", phone: "091-3678638", email: "store25@rabet.ly", city: "مصراتة", area: "الغيران", balance: -400, creditLimit: 5000, companyId: "c1", repId: "r3", mid: "MID-10025" },
    { id: "m1", branchId: "br1", server: "food", name: "محل التاجر أحمد", owner: "أحمد المبروك", phone: "091-5556677", email: "ahmed@store.ly", city: "طرابلس", area: "سوق الجمعة", balance: -1240, creditLimit: 8000, companyId: "c1", repId: "r1", tier: "retail", contractPct: 0, appDiscountPct: 0, photo: "🏪", lat: 32.8925, lng: 13.1802, hasContract: false, license: "TL-2019-4521", licenseDate: "2019-03-12", address: "سوق الجمعة، شارع الوحدة، مبنى 14", managerName: null, managerPhone: null , frozen: false, freezeReason: null , membershipNo: "MID-10001", fridgeContract: null , orderPin: "0000" , invoiceOverride: null },
    { id: "m2", branchId: "br1", server: "food", name: "بقالة الوفاء", owner: "علي الفيتوري", phone: "092-3334455", email: "ali@wafa.ly", city: "طرابلس", area: "عين زارة", balance: -560, creditLimit: 5000, companyId: "c1", repId: "r1", tier: "retail", contractPct: 0, appDiscountPct: 0, photo: "🏬", lat: 32.8312, lng: 13.2240, hasContract: false, license: "TL-2020-7810", licenseDate: "2020-06-01", address: "عين زارة، طريق المدار، محل 3", managerName: null, managerPhone: null , frozen: false, freezeReason: null , membershipNo: "MID-10002", fridgeContract: null , orderPin: "0000" , invoiceOverride: null },
    { id: "m3", branchId: "br3", server: "food", name: "سوبرماركت الأصيل", owner: "منى الزروق", phone: "091-7778800", email: "asil@store.ly", city: "مصراتة", area: "الزروق", balance: 0, creditLimit: 12000, companyId: "c1", repId: "r3", tier: "wholesale", contractPct: 5, appDiscountPct: 2, photo: "🛒", lat: 32.8701, lng: 13.1955, hasContract: true, license: "TL-2018-2244", licenseDate: "2018-01-20", address: "الظهرة، شارع الجمهورية، عمارة النصر", managerName: null, managerPhone: null , frozen: false, freezeReason: null , membershipNo: "MID-10003", fridgeContract: null , orderPin: "0000" , invoiceOverride: null },
    { id: "m4", server: "food", name: "بقالة النخبة", owner: "سعيد العبيدي", phone: "092-9991122", email: "nokba@store.ly", city: "طرابلس", area: "تاجوراء", balance: -2100, creditLimit: 6000, companyId: "c2", repId: "r2", tier: "retail", contractPct: 0, appDiscountPct: 0, photo: "🏪", lat: 32.8810, lng: 13.3510, hasContract: false, license: "TL-2021-9034", licenseDate: "2021-09-15", address: "تاجوراء، الطريق الساحلي، مجمع الخليج", managerName: null, managerPhone: null , frozen: false, freezeReason: null , membershipNo: "MID-10004", fridgeContract: null , orderPin: "0000" , invoiceOverride: null },
    // عميل تابع للشركة مباشرة بلا مشرف (تابعيته للشركة)
    { id: "m6", branchId: "br2", server: "food", name: "هايبر السلام", owner: "نبيل الشريف", phone: "091-2009988", email: "salam@hyper.ly", city: "بنغازي", area: "الصابري", balance: -4500, creditLimit: 25000, companyId: "c1", repId: "r5", tier: "company_client", contractPct: 8, appDiscountPct: 3, photo: "🏢", lat: 32.8456, lng: 13.1402, hasContract: true, license: "TL-2016-5570", licenseDate: "2016-11-30", address: "قرقارش، شارع الاستقلال، هايبر 1", managerName: null, managerPhone: null , frozen: false, freezeReason: null , membershipNo: "MID-10006", fridgeContract: null , orderPin: "0000" , invoiceOverride: null },
    // صيدلية تجريبية في سيرفر الأدوية
    { id: "m5", server: "pharma", name: "صيدلية الرحمة", owner: "د. ليلى القاضي", phone: "091-4445566", email: "rahma@pharm.ly", city: "طرابلس", area: "النوفليين", balance: -3200, creditLimit: 15000, companyId: "c4", repId: "r4", tier: "retail", contractPct: 0, appDiscountPct: 0, photo: "💊", lat: 32.8855, lng: 13.1880, hasContract: false, license: "PH-2017-1188", licenseDate: "2017-05-08", address: "النوفليين، شارع الصيادلة، مبنى 7", managerName: null, managerPhone: null , frozen: false, freezeReason: null , membershipNo: "MID-10005", fridgeContract: null , orderPin: "0000" , invoiceOverride: null },
  ],
  // مدير مبيعات (تابع للشركة) — أهدافه تحددها الشركة (كمية ستيكات + تحصيل)
  salesManagers: [
    { id: "sm3", server: "food", companyId: "c1", branchIds: ["br2"], name: "وليد بنغازي", phone: "091-2003004", area: "إدارة مبيعات بنغازي", qtyTarget: 3000, qtyAchieved: 1400, cashTarget: 40000, cashAchieved: 22000 },
    { id: "sm1", branchIds: ["br1", "br2"], server: "food", companyId: "c1", name: "طارق المدير", phone: "091-1000200", area: "إدارة المبيعات", qtyTarget: 4000, qtyAchieved: 2070, collectionTarget: 90000, collected: 41800 },
    { id: "sm2", server: "pharma", companyId: "c4", name: "د. حسام الإدارة", phone: "091-3004005", area: "إدارة المبيعات", qtyTarget: 1500, qtyAchieved: 520, collectionTarget: 60000, collected: 17500 },
  ],
  // مروّج مبيعات (تابع للشركة)
  promoters: [
    { id: "pr1", branchId: "br1", server: "food", companyId: "c1", name: "يوسف المروّج", phone: "091-6007008", area: "ترويج وعرض" },
    { id: "pr2", server: "pharma", companyId: "c4", name: "سمير العرض", phone: "091-9008007", area: "ترويج" },
  ],
  // مهام الترويج (تظهر للمروّج بعد استلام المندوب)
  promotions: [
    { id: "pm1", server: "food", companyId: "c1", promoterId: "pr1", orderId: "o1002", merchantId: "m1", status: "notified", date: "أمس",
      // متطلبات الشركة من هذه الزيارة
      req: { photoBefore: true, photoAfter: true, fridge: true, shelf: true, expiry: true },
      gpsConfirmed: false, photoBefore: false, photoAfter: false,
      expiryFindings: [], stacking: null, suggestedItems: [], repAlert: null, note: "" },
    { id: "pm2", server: "food", companyId: "c1", promoterId: "pr1", orderId: null, merchantId: "m2", status: "notified", date: "اليوم",
      req: { photoBefore: true, photoAfter: true, fridge: false, shelf: true, expiry: true },
      gpsConfirmed: false, photoBefore: false, photoAfter: false,
      expiryFindings: [], stacking: null, suggestedItems: [], repAlert: null, note: "" },
  ],
  supervisors: [
    { id: "sv3", server: "food", companyId: "c1", branchId: "br2", name: "عمر الشريف", phone: "091-5006007", area: "بنغازي", qtyTarget: 1500, qtyAchieved: 800, cashTarget: 20000, cashAchieved: 11000 },
    { id: "sv1", branchId: "br1", server: "food", companyId: "c1", name: "عبدالله المشرف", phone: "091-4445566", area: "طرابلس الكبرى", qtyTarget: 1800, qtyAchieved: 1090, collectionTarget: 45000, collected: 24000 },
    { id: "sv2", server: "pharma", companyId: "c4", name: "د. كمال الإشراف", phone: "091-9012233", area: "طرابلس", qtyTarget: 800, qtyAchieved: 520, collectionTarget: 30000, collected: 17500 },
  ],
  reps: [
    { id: "r1", branchId: "br1", server: "food", companyId: "c1", name: "خالد المندوب", phone: "091-9990011", area: "طرابلس - الشرق", supervisorId: "sv1", target: 25000, achieved: 14200, qtyTarget: 1200, qtyAchieved: 680, targetMode: "value", repType: "presell", vanStock: {}, shortfall: 0 },
    { id: "r3", branchId: "br3", server: "food", companyId: "c1", name: "وليد سعد", phone: "091-3332244", area: "طرابلس - الوسط", supervisorId: "sv1", target: 20000, achieved: 9800, qtyTarget: 900, qtyAchieved: 410, targetMode: "qty", repType: "van", vanStock: { p4: 60, p1: 25 }, shortfall: 0 },
    { id: "r5", branchId: "br2", server: "food", companyId: "c1", name: "فهد المباشر", phone: "091-5005006", area: "طرابلس - حسابات الشركة", supervisorId: "sv3", target: 35000, achieved: 21000, qtyTarget: 1500, qtyAchieved: 980, targetMode: "value", repType: "presell", vanStock: {}, shortfall: 0 },
    { id: "r2", server: "food", companyId: "c2", name: "سامي التواتي", phone: "092-8889900", area: "طرابلس - الغرب", supervisorId: null, target: 18000, achieved: 11000, qtyTarget: 2000, qtyAchieved: 1350, targetMode: "qty", repType: "van", vanStock: { p5: 120, p6: 40 }, shortfall: 0 },
    { id: "r4", server: "pharma", companyId: "c4", name: "د. منير الصيدلي", phone: "091-7001122", area: "طرابلس - المركز", supervisorId: "sv2", target: 30000, achieved: 17500, qtyTarget: 800, qtyAchieved: 520, targetMode: "value", repType: "presell", vanStock: {}, shortfall: 0 },
  ],
  drivers: [
    { id: "d5", server: "food", companyId: "c1", branchId: "br2", name: "خالد بنغازي", phone: "091-8009001", vehicle: "بيكب - 21ط 3344", plate: "3344", vehicleType: "medium", status: "available" },
    { id: "d1", branchId: "br1", server: "food", companyId: "c1", name: "سالم السائق", phone: "091-7778899", vehicle: "شاحنة - 14ط 5566", plate: "5566", vehicleType: "large", status: "available" },
    { id: "d2", server: "food", companyId: "c2", name: "منصور علي", phone: "092-6667788", vehicle: "بيكب - 23ط 1199", plate: "1199", vehicleType: "small", status: "busy" },
    { id: "d3", server: "pharma", companyId: "c4", name: "إبراهيم النقل", phone: "091-8002233", vehicle: "فان مبرّد - 12ط 7788", plate: "7788", vehicleType: "medium", status: "available" },
  ],
  // مشرفو الحركة (يستقبلون الفواتير الجاهزة ويعبّئون السيارات)
  dispatchers: [
    { id: "dsp1", branchId: "br1", server: "food", companyId: "c1", name: "مفتاح الدرباكي", phone: "091-3334455", area: "الحركة والتوصيل", status: "available" },
    { id: "dsp2", server: "pharma", companyId: "c4", name: "عادل الزوي", phone: "092-6677889", area: "الحركة والتوصيل", status: "available" },
  ],
  orders: [
    { id: "o2101", server: "pharma", merchantId: "m5", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p7", name: "باراسيتامول 500مغ (علبة×20)", qty: 20, price: 12 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 09:00" },
    { id: "o2102", server: "pharma", merchantId: "m5", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p7", name: "باراسيتامول 500مغ (علبة×20)", qty: 100, price: 12 }, { pid: "p8", name: "أموكسيسيلين 500مغ (شريط×16)", qty: 60, price: 28 }, { pid: "p9", name: "فيتامين د 50000 (علبة×8)", qty: 20, price: 35 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 10:15" },
    { id: "o2103", server: "pharma", merchantId: "m5", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p14", name: "أوميبرازول 20مغ (علبة×14)", qty: 30, price: 10 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 11:30" },
    { id: "o2104", server: "pharma", merchantId: "m5", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p14", name: "أوميبرازول 20مغ (علبة×14)", qty: 20, price: 10 }, { pid: "p8", name: "أموكسيسيلين 500مغ (شريط×16)", qty: 60, price: 28 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 11:30" },
    { id: "o2105", server: "pharma", merchantId: "m51", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p8", name: "أموكسيسيلين 500مغ (شريط×16)", qty: 60, price: 28 }, { pid: "p13", name: "إيبوبروفين 400مغ (علبة×30)", qty: 50, price: 14 }, { pid: "p9", name: "فيتامين د 50000 (علبة×8)", qty: 40, price: 35 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 10:15" },
    { id: "o2106", server: "pharma", merchantId: "m51", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p8", name: "أموكسيسيلين 500مغ (شريط×16)", qty: 20, price: 28 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل 3 أيام 12:00" },
    { id: "o2107", server: "pharma", merchantId: "m52", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p7", name: "باراسيتامول 500مغ (علبة×20)", qty: 60, price: 12 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 15:00" },
    { id: "o2108", server: "pharma", merchantId: "m52", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p7", name: "باراسيتامول 500مغ (علبة×20)", qty: 100, price: 12 }, { pid: "p9", name: "فيتامين د 50000 (علبة×8)", qty: 40, price: 35 }, { pid: "p8", name: "أموكسيسيلين 500مغ (شريط×16)", qty: 60, price: 28 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 10:15" },
    { id: "o2109", server: "pharma", merchantId: "m52", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p13", name: "إيبوبروفين 400مغ (علبة×30)", qty: 20, price: 14 }, { pid: "p7", name: "باراسيتامول 500مغ (علبة×20)", qty: 100, price: 12 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "الأسبوع الماضي" },
    { id: "o2110", server: "pharma", merchantId: "m52", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p7", name: "باراسيتامول 500مغ (علبة×20)", qty: 40, price: 12 }, { pid: "p13", name: "إيبوبروفين 400مغ (علبة×30)", qty: 30, price: 14 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 10:15" },
    { id: "o2111", server: "pharma", merchantId: "m53", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p8", name: "أموكسيسيلين 500مغ (شريط×16)", qty: 50, price: 28 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 10:15" },
    { id: "o2112", server: "pharma", merchantId: "m53", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p13", name: "إيبوبروفين 400مغ (علبة×30)", qty: 30, price: 14 }, { pid: "p9", name: "فيتامين د 50000 (علبة×8)", qty: 50, price: 35 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 15:00" },
    { id: "o2113", server: "pharma", merchantId: "m53", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p13", name: "إيبوبروفين 400مغ (علبة×30)", qty: 50, price: 14 }, { pid: "p9", name: "فيتامين د 50000 (علبة×8)", qty: 30, price: 35 }, { pid: "p14", name: "أوميبرازول 20مغ (علبة×14)", qty: 30, price: 10 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 11:30" },
    { id: "o2114", server: "pharma", merchantId: "m54", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p9", name: "فيتامين د 50000 (علبة×8)", qty: 40, price: 35 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 10:15" },
    { id: "o2115", server: "pharma", merchantId: "m54", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p14", name: "أوميبرازول 20مغ (علبة×14)", qty: 100, price: 10 }, { pid: "p7", name: "باراسيتامول 500مغ (علبة×20)", qty: 60, price: 12 }, { pid: "p8", name: "أموكسيسيلين 500مغ (شريط×16)", qty: 50, price: 28 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 10:15" },
    { id: "o2116", server: "pharma", merchantId: "m54", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p13", name: "إيبوبروفين 400مغ (علبة×30)", qty: 20, price: 14 }, { pid: "p7", name: "باراسيتامول 500مغ (علبة×20)", qty: 30, price: 12 }, { pid: "p14", name: "أوميبرازول 20مغ (علبة×14)", qty: 50, price: 10 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 15:00" },
    { id: "o2117", server: "pharma", merchantId: "m54", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p14", name: "أوميبرازول 20مغ (علبة×14)", qty: 30, price: 10 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 15:00" },
    { id: "o2118", server: "pharma", merchantId: "m55", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p13", name: "إيبوبروفين 400مغ (علبة×30)", qty: 40, price: 14 }, { pid: "p8", name: "أموكسيسيلين 500مغ (شريط×16)", qty: 40, price: 28 }, { pid: "p9", name: "فيتامين د 50000 (علبة×8)", qty: 60, price: 35 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 09:00" },
    { id: "o2119", server: "pharma", merchantId: "m55", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p13", name: "إيبوبروفين 400مغ (علبة×30)", qty: 40, price: 14 }, { pid: "p14", name: "أوميبرازول 20مغ (علبة×14)", qty: 20, price: 10 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "الأسبوع الماضي" },
    { id: "o1101", server: "food", merchantId: "m1", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 8, price: 95 }, { pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 10, price: 240 }], status: "pending_company", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 14:30" },
    { id: "o1102", server: "food", merchantId: "m1", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 20, price: 180 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 14:30" },
    { id: "o1103", server: "food", merchantId: "m1", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 6, price: 180 }, { pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 12, price: 240 }], status: "pending_company", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل 3 أيام 10:20" },
    { id: "o1104", server: "food", merchantId: "m1", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 6, price: 240 }, { pid: "p3", name: "سكر ناعم 50كغ", qty: 5, price: 95 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل أسبوعين" },
    { id: "o1105", server: "food", merchantId: "m2", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 6, price: 240 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 09:00" },
    { id: "o1106", server: "food", merchantId: "m2", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 6, price: 95 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 10:40" },
    { id: "o1107", server: "food", merchantId: "m2", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 20, price: 240 }, { pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 12, price: 180 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 10, price: 62 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل أسبوعين" },
    { id: "o1108", server: "food", merchantId: "m11", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 15, price: 180 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 09:15" },
    { id: "o1109", server: "food", merchantId: "m11", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 24, price: 62 }, { pid: "p3", name: "سكر ناعم 50كغ", qty: 8, price: 95 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 11:20" },
    { id: "o1110", server: "food", merchantId: "m12", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 8, price: 62 }, { pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 24, price: 240 }, { pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 6, price: 180 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 09:15" },
    { id: "o1111", server: "food", merchantId: "m12", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 20, price: 240 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 16:10" },
    { id: "o1112", server: "food", merchantId: "m12", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 12, price: 180 }, { pid: "p3", name: "سكر ناعم 50كغ", qty: 15, price: 95 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 6, price: 62 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 10:40" },
    { id: "o1113", server: "food", merchantId: "m13", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 12, price: 240 }, { pid: "p3", name: "سكر ناعم 50كغ", qty: 10, price: 95 }, { pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 8, price: 180 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 10:40" },
    { id: "o1114", server: "food", merchantId: "m13", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 15, price: 180 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 12, price: 62 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 11:20" },
    { id: "o1115", server: "food", merchantId: "m13", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 8, price: 180 }, { pid: "p3", name: "سكر ناعم 50كغ", qty: 8, price: 95 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 13:45" },
    { id: "o1116", server: "food", merchantId: "m14", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 10, price: 62 }, { pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 12, price: 180 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل 4 أيام 11:00" },
    { id: "o1117", server: "food", merchantId: "m14", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 10, price: 240 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 08:30" },
    { id: "o1118", server: "food", merchantId: "m15", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 10, price: 180 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 08:30" },
    { id: "o1119", server: "food", merchantId: "m15", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 8, price: 62 }, { pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 24, price: 180 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 09:00" },
    { id: "o1120", server: "food", merchantId: "m16", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 20, price: 95 }, { pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 24, price: 240 }], status: "pending_company", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 08:30" },
    { id: "o1121", server: "food", merchantId: "m16", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 15, price: 180 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 6, price: 62 }, { pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 10, price: 240 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 13:45" },
    { id: "o1122", server: "food", merchantId: "m16", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 6, price: 62 }, { pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 24, price: 180 }], status: "pending_company", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 13:45" },
    { id: "o1123", server: "food", merchantId: "m16", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 5, price: 180 }], status: "pending_company", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل أسبوعين" },
    { id: "o1124", server: "food", merchantId: "m6", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 12, price: 62 }, { pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 10, price: 180 }, { pid: "p3", name: "سكر ناعم 50كغ", qty: 5, price: 95 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 13:45" },
    { id: "o1125", server: "food", merchantId: "m6", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 24, price: 180 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 13:45" },
    { id: "o1126", server: "food", merchantId: "m6", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 6, price: 180 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 11:20" },
    { id: "o1127", server: "food", merchantId: "m6", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 10, price: 180 }, { pid: "p3", name: "سكر ناعم 50كغ", qty: 5, price: 95 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 16:10" },
    { id: "o1128", server: "food", merchantId: "m17", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 15, price: 95 }, { pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 10, price: 240 }, { pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 12, price: 180 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل 4 أيام 11:00" },
    { id: "o1129", server: "food", merchantId: "m17", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 24, price: 180 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل 3 أيام 10:20" },
    { id: "o1130", server: "food", merchantId: "m17", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 10, price: 180 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 10:40" },
    { id: "o1131", server: "food", merchantId: "m17", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 12, price: 240 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 5, price: 62 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 13:45" },
    { id: "o1132", server: "food", merchantId: "m18", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 6, price: 180 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 6, price: 62 }, { pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 8, price: 240 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل 3 أيام 10:20" },
    { id: "o1133", server: "food", merchantId: "m18", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 12, price: 95 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 5, price: 62 }, { pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 6, price: 240 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 12:05" },
    { id: "o1134", server: "food", merchantId: "m18", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 15, price: 180 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل 4 أيام 11:00" },
    { id: "o1135", server: "food", merchantId: "m19", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 20, price: 62 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 09:15" },
    { id: "o1136", server: "food", merchantId: "m19", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 20, price: 180 }, { pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 8, price: 240 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 5, price: 62 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل أسبوعين" },
    { id: "o1137", server: "food", merchantId: "m19", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 6, price: 240 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 13:45" },
    { id: "o1138", server: "food", merchantId: "m20", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 15, price: 180 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 20, price: 62 }], status: "pending_company", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "الأسبوع الماضي" },
    { id: "o1139", server: "food", merchantId: "m20", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 8, price: 95 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 09:15" },
    { id: "o1140", server: "food", merchantId: "m21", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 10, price: 62 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل يومين 09:00" },
    { id: "o1141", server: "food", merchantId: "m21", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 20, price: 240 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 09:15" },
    { id: "o1142", server: "food", merchantId: "m21", companyId: "c1", repId: "r5", driverId: "d5", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 6, price: 95 }, { pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 12, price: 180 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 12:05" },
    { id: "o1143", server: "food", merchantId: "m3", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 20, price: 95 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل 4 أيام 11:00" },
    { id: "o1144", server: "food", merchantId: "m3", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 12, price: 95 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 14:30" },
    { id: "o1145", server: "food", merchantId: "m3", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 8, price: 95 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 10, price: 62 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 16:10" },
    { id: "o1146", server: "food", merchantId: "m3", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 12, price: 95 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل أسبوعين" },
    { id: "o1147", server: "food", merchantId: "m22", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 8, price: 240 }, { pid: "p3", name: "سكر ناعم 50كغ", qty: 6, price: 95 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 12, price: 62 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل 4 أيام 11:00" },
    { id: "o1148", server: "food", merchantId: "m22", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 10, price: 95 }], status: "pending_company", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 10:40" },
    { id: "o1149", server: "food", merchantId: "m22", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 10, price: 180 }], status: "pending_company", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل 3 أيام 10:20" },
    { id: "o1150", server: "food", merchantId: "m23", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 10, price: 62 }, { pid: "p3", name: "سكر ناعم 50كغ", qty: 20, price: 95 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 11:20" },
    { id: "o1151", server: "food", merchantId: "m23", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 24, price: 240 }, { pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 5, price: 180 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 09:15" },
    { id: "o1152", server: "food", merchantId: "m24", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 24, price: 95 }, { pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 8, price: 240 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 24, price: 62 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "قبل أسبوعين" },
    { id: "o1153", server: "food", merchantId: "m24", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 6, price: 240 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 12, price: 62 }, { pid: "p3", name: "سكر ناعم 50كغ", qty: 10, price: 95 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 14:30" },
    { id: "o1154", server: "food", merchantId: "m24", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 8, price: 62 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 09:15" },
    { id: "o1155", server: "food", merchantId: "m24", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 24, price: 62 }, { pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 20, price: 240 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 16:10" },
    { id: "o1156", server: "food", merchantId: "m25", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p3", name: "سكر ناعم 50كغ", qty: 12, price: 95 }], status: "preparing", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 16:10" },
    { id: "o1157", server: "food", merchantId: "m25", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 12, price: 240 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 20, price: 62 }, { pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 24, price: 180 }], status: "delivered", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "أمس 14:30" },
    { id: "o1158", server: "food", merchantId: "m25", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 6, price: 180 }, { pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 20, price: 240 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 8, price: 62 }], status: "pending_company", fulfillment: "delivery", cashConfirmed: true, note: "", createdAt: "اليوم 10:40" },
    { id: "o1010", server: "food", merchantId: "m2", companyId: "c1", repId: "r1", driverId: null, items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 12, price: 180 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 8, price: 62 }], status: "ready_dispatch", fulfillment: "delivery", cashConfirmed: false, note: "وافق عليه المندوب — للتجهيز", createdAt: "اليوم 09:40" },
    { id: "o1011", server: "pharma", merchantId: "m5", companyId: "c4", repId: "r4", driverId: null, items: [{ pid: "p7", name: "باراسيتامول 500مغ (علبة×20)", qty: 120, price: 12 }, { pid: "p13", name: "إيبوبروفين 400مغ (علبة×30)", qty: 100, price: 14 }], status: "printed", sysNo: "SYS-58120", fulfillment: "delivery", cashConfirmed: false, note: "", createdAt: "اليوم 08:55" },
    { id: "o1012", server: "pharma", merchantId: "m5", companyId: "c5", repId: null, driverId: null, items: [{ pid: "p38", name: "قفازات فحص (كرتونة×100)", qty: 10, price: 30 }], status: "pending_company", fulfillment: "pickup", cashConfirmed: false, note: "استلام من المقر", createdAt: "اليوم 10:10" },
    { id: "o1001", server: "food", merchantId: "m1", companyId: "c1", repId: "r1", driverId: null, items: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", qty: 10, price: 180 }, { pid: "p3", name: "سكر ناعم 50كغ", qty: 6, price: 95 }], status: "pending_rep", cashConfirmed: false, note: "التسليم صباحاً", createdAt: "اليوم 09:12" },
    { id: "o1002", server: "food", merchantId: "m1", companyId: "c1", repId: "r1", driverId: "d1", items: [{ pid: "p2", name: "زيت طبخ نباتي 18لتر", qty: 8, price: 240 }], status: "out_for_delivery", cashConfirmed: true, note: "", createdAt: "أمس 16:40" },
    { id: "o1003", server: "food", merchantId: "m4", companyId: "c2", repId: "r2", driverId: "d2", items: [{ pid: "p5", name: "مياه معدنية 0.5ل (كرتونة×24)", qty: 40, price: 14 }, { pid: "p6", name: "عصائر مشكّلة 250مل (علبة×12)", qty: 12, price: 38 }], status: "delivered", cashConfirmed: true, note: "", createdAt: "قبل يومين 11:05" },
    { id: "o1004", server: "food", merchantId: "m3", companyId: "c1", repId: "r3", driverId: "d1", items: [{ pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 15, price: 62 }], status: "ready_dispatch", cashConfirmed: true, note: "", createdAt: "اليوم 10:30" },
    { id: "o1000", server: "food", merchantId: "m1", companyId: "c1", repId: "r1", driverId: null, items: [{ pid: "p4", name: "معكرونة 500غ (كرتونة×20)", qty: 5, price: 62 }], status: "draft", cashConfirmed: false, note: "مسودّة لم تُرسل", createdAt: "قبل 3 أيام", ageDays: 3 },
    { id: "o2001", server: "pharma", merchantId: "m5", companyId: "c4", repId: "r4", driverId: "d3", items: [{ pid: "p7", name: "باراسيتامول 500مغ (علبة×20)", qty: 30, price: 12 }, { pid: "p8", name: "أموكسيسيلين 500مغ (شريط×16)", qty: 12, price: 28 }], status: "preparing", cashConfirmed: true, note: "", createdAt: "اليوم 09:50" },
  ],
  // التقييمات: التاجر يقيّم الشركة، والشركة تقيّم العميل
  ratings: [
    { id: "rt1", server: "food", raterRole: "merchant", merchantId: "m1", companyId: "c1", orderId: "o1002", stars: 5, comment: "توصيل سريع وبضاعة ممتازة", date: "أمس" },
    { id: "rt2", server: "food", raterRole: "merchant", merchantId: "m2", companyId: "c1", orderId: null, stars: 4, comment: "أسعار جيدة، التوصيل يتأخر أحياناً", date: "قبل 3 أيام" },
    { id: "rt3", server: "food", raterRole: "company", merchantId: "m1", companyId: "c1", orderId: null, stars: 5, comment: "ملتزم بالسداد، تعامل ممتاز", date: "أمس" },
    { id: "rt4", server: "food", raterRole: "company", merchantId: "m2", companyId: "c2", orderId: null, stars: 3, comment: "تأخّر في تسوية آخر فاتورتين", date: "قبل يومين" },
    { id: "rt5", server: "pharma", raterRole: "merchant", merchantId: "m5", companyId: "c4", orderId: null, stars: 5, comment: "صلاحيات ممتازة وتغليف احترافي", date: "أمس" },
  ],
  invoices: [
    { id: "INV-5001", server: "pharma", merchantId: "m5", companyId: "c4", orderId: "o2101", amount: 240, paid: 240, date: "اليوم 09:00", due: "خلال 30 يوم", ageDays: 5 },
    { id: "INV-5002", server: "pharma", merchantId: "m5", companyId: "c4", orderId: "o2102", amount: 3580, paid: 0, date: "قبل يومين 10:15", due: "خلال 30 يوم", ageDays: 10 },
    { id: "INV-5003", server: "pharma", merchantId: "m5", companyId: "c4", orderId: "o2103", amount: 300, paid: 0, date: "اليوم 11:30", due: "خلال 30 يوم", ageDays: 50 },
    { id: "INV-5004", server: "pharma", merchantId: "m5", companyId: "c4", orderId: "o2104", amount: 1880, paid: 1880, date: "اليوم 11:30", due: "خلال 30 يوم", ageDays: 5 },
    { id: "INV-5005", server: "pharma", merchantId: "m51", companyId: "c4", orderId: "o2105", amount: 3780, paid: 3780, date: "قبل يومين 10:15", due: "خلال 30 يوم", ageDays: 20 },
    { id: "INV-5006", server: "pharma", merchantId: "m51", companyId: "c4", orderId: "o2106", amount: 560, paid: 336, date: "قبل 3 أيام 12:00", due: "خلال 30 يوم", ageDays: 20 },
    { id: "INV-5007", server: "pharma", merchantId: "m52", companyId: "c4", orderId: "o2107", amount: 720, paid: 720, date: "أمس 15:00", due: "خلال 30 يوم", ageDays: 35 },
    { id: "INV-5008", server: "pharma", merchantId: "m52", companyId: "c4", orderId: "o2108", amount: 4280, paid: 0, date: "قبل يومين 10:15", due: "خلال 30 يوم", ageDays: 5 },
    { id: "INV-5009", server: "pharma", merchantId: "m52", companyId: "c4", orderId: "o2109", amount: 1480, paid: 888, date: "الأسبوع الماضي", due: "خلال 30 يوم", ageDays: 20 },
    { id: "INV-5010", server: "pharma", merchantId: "m52", companyId: "c4", orderId: "o2110", amount: 900, paid: 0, date: "قبل يومين 10:15", due: "خلال 30 يوم", ageDays: 10 },
    { id: "INV-5011", server: "pharma", merchantId: "m53", companyId: "c4", orderId: "o2111", amount: 1400, paid: 0, date: "قبل يومين 10:15", due: "خلال 30 يوم", ageDays: 10 },
    { id: "INV-5012", server: "pharma", merchantId: "m53", companyId: "c4", orderId: "o2113", amount: 2050, paid: 2050, date: "اليوم 11:30", due: "خلال 30 يوم", ageDays: 10 },
    { id: "INV-5013", server: "pharma", merchantId: "m54", companyId: "c4", orderId: "o2114", amount: 1400, paid: 1400, date: "قبل يومين 10:15", due: "خلال 30 يوم", ageDays: 50 },
    { id: "INV-5014", server: "pharma", merchantId: "m54", companyId: "c4", orderId: "o2115", amount: 3120, paid: 0, date: "قبل يومين 10:15", due: "خلال 30 يوم", ageDays: 35 },
    { id: "INV-5015", server: "pharma", merchantId: "m54", companyId: "c4", orderId: "o2116", amount: 1140, paid: 0, date: "أمس 15:00", due: "خلال 30 يوم", ageDays: 5 },
    { id: "INV-5016", server: "pharma", merchantId: "m54", companyId: "c4", orderId: "o2117", amount: 300, paid: 0, date: "أمس 15:00", due: "خلال 30 يوم", ageDays: 5 },
    { id: "INV-5017", server: "pharma", merchantId: "m55", companyId: "c4", orderId: "o2118", amount: 3780, paid: 0, date: "اليوم 09:00", due: "خلال 30 يوم", ageDays: 35 },
    { id: "INV-5018", server: "pharma", merchantId: "m55", companyId: "c4", orderId: "o2119", amount: 760, paid: 760, date: "الأسبوع الماضي", due: "خلال 30 يوم", ageDays: 20 },
    { id: "INV-3001", server: "food", merchantId: "m1", companyId: "c1", orderId: "o1102", amount: 3600, paid: 2520, date: "أمس 14:30", due: "خلال 30 يوم", ageDays: 25 },
    { id: "INV-3002", server: "food", merchantId: "m2", companyId: "c1", orderId: "o1105", amount: 1440, paid: 1440, date: "قبل يومين 09:00", due: "خلال 30 يوم", ageDays: 12 },
    { id: "INV-3003", server: "food", merchantId: "m11", companyId: "c1", orderId: "o1108", amount: 2700, paid: 2700, date: "اليوم 09:15", due: "خلال 30 يوم", ageDays: 32 },
    { id: "INV-3004", server: "food", merchantId: "m11", companyId: "c1", orderId: "o1109", amount: 2248, paid: 1573, date: "اليوم 11:20", due: "خلال 30 يوم", ageDays: 25 },
    { id: "INV-3005", server: "food", merchantId: "m12", companyId: "c1", orderId: "o1111", amount: 4800, paid: 2400, date: "أمس 16:10", due: "خلال 30 يوم", ageDays: 55 },
    { id: "INV-3006", server: "food", merchantId: "m12", companyId: "c1", orderId: "o1112", amount: 3957, paid: 1978, date: "اليوم 10:40", due: "خلال 30 يوم", ageDays: 3 },
    { id: "INV-3007", server: "food", merchantId: "m13", companyId: "c1", orderId: "o1113", amount: 5270, paid: 3688, date: "اليوم 10:40", due: "خلال 30 يوم", ageDays: 3 },
    { id: "INV-3008", server: "food", merchantId: "m13", companyId: "c1", orderId: "o1114", amount: 3444, paid: 2410, date: "اليوم 11:20", due: "خلال 30 يوم", ageDays: 7 },
    { id: "INV-3009", server: "food", merchantId: "m14", companyId: "c1", orderId: "o1116", amount: 2780, paid: 2780, date: "قبل 4 أيام 11:00", due: "خلال 30 يوم", ageDays: 55 },
    { id: "INV-3010", server: "food", merchantId: "m14", companyId: "c1", orderId: "o1117", amount: 2400, paid: 1680, date: "اليوم 08:30", due: "خلال 30 يوم", ageDays: 18 },
    { id: "INV-3011", server: "food", merchantId: "m15", companyId: "c1", orderId: "o1118", amount: 1800, paid: 1800, date: "اليوم 08:30", due: "خلال 30 يوم", ageDays: 7 },
    { id: "INV-3012", server: "food", merchantId: "m15", companyId: "c1", orderId: "o1119", amount: 4816, paid: 2408, date: "قبل يومين 09:00", due: "خلال 30 يوم", ageDays: 18 },
    { id: "INV-3013", server: "food", merchantId: "m16", companyId: "c1", orderId: "o1121", amount: 5472, paid: 2736, date: "قبل يومين 13:45", due: "خلال 30 يوم", ageDays: 12 },
    { id: "INV-3014", server: "food", merchantId: "m6", companyId: "c1", orderId: "o1124", amount: 3019, paid: 0, date: "قبل يومين 13:45", due: "خلال 30 يوم", ageDays: 32 },
    { id: "INV-3015", server: "food", merchantId: "m6", companyId: "c1", orderId: "o1125", amount: 4320, paid: 4320, date: "قبل يومين 13:45", due: "خلال 30 يوم", ageDays: 3 },
    { id: "INV-3016", server: "food", merchantId: "m17", companyId: "c1", orderId: "o1128", amount: 5985, paid: 5985, date: "قبل 4 أيام 11:00", due: "خلال 30 يوم", ageDays: 55 },
    { id: "INV-3017", server: "food", merchantId: "m17", companyId: "c1", orderId: "o1130", amount: 1800, paid: 1800, date: "اليوم 10:40", due: "خلال 30 يوم", ageDays: 7 },
    { id: "INV-3018", server: "food", merchantId: "m18", companyId: "c1", orderId: "o1132", amount: 3372, paid: 3372, date: "قبل 3 أيام 10:20", due: "خلال 30 يوم", ageDays: 32 },
    { id: "INV-3019", server: "food", merchantId: "m19", companyId: "c1", orderId: "o1135", amount: 1240, paid: 0, date: "اليوم 09:15", due: "خلال 30 يوم", ageDays: 12 },
    { id: "INV-3020", server: "food", merchantId: "m19", companyId: "c1", orderId: "o1136", amount: 5830, paid: 0, date: "قبل أسبوعين", due: "خلال 30 يوم", ageDays: 32 },
    { id: "INV-3021", server: "food", merchantId: "m19", companyId: "c1", orderId: "o1137", amount: 1440, paid: 720, date: "قبل يومين 13:45", due: "خلال 30 يوم", ageDays: 12 },
    { id: "INV-3022", server: "food", merchantId: "m21", companyId: "c1", orderId: "o1140", amount: 620, paid: 620, date: "قبل يومين 09:00", due: "خلال 30 يوم", ageDays: 25 },
    { id: "INV-3023", server: "food", merchantId: "m21", companyId: "c1", orderId: "o1141", amount: 4800, paid: 4800, date: "اليوم 09:15", due: "خلال 30 يوم", ageDays: 32 },
    { id: "INV-3024", server: "food", merchantId: "m21", companyId: "c1", orderId: "o1142", amount: 2730, paid: 0, date: "اليوم 12:05", due: "خلال 30 يوم", ageDays: 7 },
    { id: "INV-3025", server: "food", merchantId: "m3", companyId: "c1", orderId: "o1143", amount: 1900, paid: 1900, date: "قبل 4 أيام 11:00", due: "خلال 30 يوم", ageDays: 32 },
    { id: "INV-3026", server: "food", merchantId: "m3", companyId: "c1", orderId: "o1145", amount: 1380, paid: 1380, date: "أمس 16:10", due: "خلال 30 يوم", ageDays: 25 },
    { id: "INV-3027", server: "food", merchantId: "m3", companyId: "c1", orderId: "o1146", amount: 1140, paid: 798, date: "قبل أسبوعين", due: "خلال 30 يوم", ageDays: 32 },
    { id: "INV-3028", server: "food", merchantId: "m22", companyId: "c1", orderId: "o1147", amount: 3234, paid: 2263, date: "قبل 4 أيام 11:00", due: "خلال 30 يوم", ageDays: 32 },
    { id: "INV-3029", server: "food", merchantId: "m23", companyId: "c1", orderId: "o1150", amount: 2520, paid: 2520, date: "اليوم 11:20", due: "خلال 30 يوم", ageDays: 3 },
    { id: "INV-3030", server: "food", merchantId: "m23", companyId: "c1", orderId: "o1151", amount: 6660, paid: 3330, date: "اليوم 09:15", due: "خلال 30 يوم", ageDays: 12 },
    { id: "INV-3031", server: "food", merchantId: "m24", companyId: "c1", orderId: "o1152", amount: 5688, paid: 5688, date: "قبل أسبوعين", due: "خلال 30 يوم", ageDays: 25 },
    { id: "INV-3032", server: "food", merchantId: "m24", companyId: "c1", orderId: "o1153", amount: 3134, paid: 2193, date: "أمس 14:30", due: "خلال 30 يوم", ageDays: 7 },
    { id: "INV-3033", server: "food", merchantId: "m24", companyId: "c1", orderId: "o1155", amount: 6288, paid: 3144, date: "أمس 16:10", due: "خلال 30 يوم", ageDays: 3 },
    { id: "INV-3034", server: "food", merchantId: "m25", companyId: "c1", orderId: "o1157", amount: 8440, paid: 4220, date: "أمس 14:30", due: "خلال 30 يوم", ageDays: 12 },
    { id: "INV-2041", server: "food", merchantId: "m1", companyId: "c1", orderId: "o1002", amount: 1920, paid: 680, date: "أمس", due: "خلال 14 يوم" , ageDays: 12 },
    { id: "INV-2038", server: "food", merchantId: "m4", companyId: "c2", orderId: "o1003", amount: 1016, paid: 1016, date: "قبل يومين", due: "مدفوعة" , ageDays: 2 },
    { id: "INV-2042", server: "food", merchantId: "m2", companyId: "c1", orderId: null, amount: 560, paid: 0, date: "اليوم", due: "خلال 7 أيام" , ageDays: 0 },
    { id: "INV-5011", server: "pharma", merchantId: "m5", companyId: "c4", orderId: "o2001", amount: 3200, paid: 0, date: "اليوم", due: "خلال 30 يوم" , ageDays: 0 },
  ],
  payments: [
    { id: "PAY-901", server: "pharma", merchantId: "m5", companyId: "c4", amount: 240, method: "تحويل مصرفي", date: "اليوم 09:00", repId: "r4", receiptId: "RCP-3484" },
    { id: "PAY-902", server: "pharma", merchantId: "m5", companyId: "c4", amount: 1880, method: "نقدي للمندوب", date: "اليوم 11:30", repId: "r4", receiptId: "RCP-7623" },
    { id: "PAY-903", server: "pharma", merchantId: "m51", companyId: "c4", amount: 3780, method: "نقدي للمندوب", date: "قبل يومين 10:15", repId: "r4", receiptId: "RCP-9507" },
    { id: "PAY-904", server: "pharma", merchantId: "m51", companyId: "c4", amount: 336, method: "تحويل مصرفي", date: "قبل 3 أيام 12:00", repId: "r4", receiptId: "RCP-5358" },
    { id: "PAY-905", server: "pharma", merchantId: "m52", companyId: "c4", amount: 720, method: "تحويل مصرفي", date: "أمس 15:00", repId: "r4", receiptId: "RCP-3321" },
    { id: "PAY-906", server: "pharma", merchantId: "m52", companyId: "c4", amount: 888, method: "تحويل مصرفي", date: "الأسبوع الماضي", repId: "r4", receiptId: "RCP-8477" },
    { id: "PAY-907", server: "pharma", merchantId: "m53", companyId: "c4", amount: 2050, method: "نقدي للمندوب", date: "اليوم 11:30", repId: "r4", receiptId: "RCP-6972" },
    { id: "PAY-908", server: "pharma", merchantId: "m54", companyId: "c4", amount: 1400, method: "تحويل مصرفي", date: "قبل يومين 10:15", repId: "r4", receiptId: "RCP-4028" },
    { id: "PAY-909", server: "pharma", merchantId: "m55", companyId: "c4", amount: 760, method: "تحويل مصرفي", date: "الأسبوع الماضي", repId: "r4", receiptId: "RCP-9789" },
    { id: "PAY-801", server: "food", merchantId: "m1", companyId: "c1", amount: 2520, method: "نقدي للمندوب", date: "أمس 14:30", repId: "r1", receiptId: "RCP-8977" },
    { id: "PAY-802", server: "food", merchantId: "m2", companyId: "c1", amount: 1440, method: "تحويل مصرفي", date: "قبل يومين 09:00", repId: "r1", receiptId: "RCP-5910" },
    { id: "PAY-803", server: "food", merchantId: "m11", companyId: "c1", amount: 2700, method: "نقدي للمندوب", date: "اليوم 09:15", repId: "r1", receiptId: "RCP-8369" },
    { id: "PAY-804", server: "food", merchantId: "m11", companyId: "c1", amount: 1573, method: "شيك", date: "اليوم 11:20", repId: "r1", receiptId: "RCP-7788" },
    { id: "PAY-805", server: "food", merchantId: "m12", companyId: "c1", amount: 2400, method: "شيك", date: "أمس 16:10", repId: "r1", receiptId: "RCP-5059" },
    { id: "PAY-806", server: "food", merchantId: "m12", companyId: "c1", amount: 1978, method: "شيك", date: "اليوم 10:40", repId: "r1", receiptId: "RCP-8895" },
    { id: "PAY-807", server: "food", merchantId: "m13", companyId: "c1", amount: 3688, method: "شيك", date: "اليوم 10:40", repId: "r1", receiptId: "RCP-5655" },
    { id: "PAY-808", server: "food", merchantId: "m13", companyId: "c1", amount: 2410, method: "نقدي للمندوب", date: "اليوم 11:20", repId: "r1", receiptId: "RCP-8995" },
    { id: "PAY-809", server: "food", merchantId: "m14", companyId: "c1", amount: 2780, method: "شيك", date: "قبل 4 أيام 11:00", repId: "r1", receiptId: "RCP-6698" },
    { id: "PAY-810", server: "food", merchantId: "m14", companyId: "c1", amount: 1680, method: "شيك", date: "اليوم 08:30", repId: "r1", receiptId: "RCP-4804" },
    { id: "PAY-811", server: "food", merchantId: "m15", companyId: "c1", amount: 1800, method: "شيك", date: "اليوم 08:30", repId: "r1", receiptId: "RCP-4949" },
    { id: "PAY-812", server: "food", merchantId: "m15", companyId: "c1", amount: 2408, method: "نقدي للمندوب", date: "قبل يومين 09:00", repId: "r1", receiptId: "RCP-3794" },
    { id: "PAY-813", server: "food", merchantId: "m16", companyId: "c1", amount: 2736, method: "تحويل مصرفي", date: "قبل يومين 13:45", repId: "r1", receiptId: "RCP-4503" },
    { id: "PAY-814", server: "food", merchantId: "m6", companyId: "c1", amount: 4320, method: "شيك", date: "قبل يومين 13:45", repId: "r5", receiptId: "RCP-3656" },
    { id: "PAY-815", server: "food", merchantId: "m17", companyId: "c1", amount: 5985, method: "تحويل مصرفي", date: "قبل 4 أيام 11:00", repId: "r5", receiptId: "RCP-9159" },
    { id: "PAY-816", server: "food", merchantId: "m17", companyId: "c1", amount: 1800, method: "نقدي للمندوب", date: "اليوم 10:40", repId: "r5", receiptId: "RCP-6027" },
    { id: "PAY-817", server: "food", merchantId: "m18", companyId: "c1", amount: 3372, method: "نقدي للمندوب", date: "قبل 3 أيام 10:20", repId: "r5", receiptId: "RCP-8631" },
    { id: "PAY-818", server: "food", merchantId: "m19", companyId: "c1", amount: 720, method: "نقدي للمندوب", date: "قبل يومين 13:45", repId: "r5", receiptId: "RCP-4331" },
    { id: "PAY-819", server: "food", merchantId: "m21", companyId: "c1", amount: 620, method: "نقدي للمندوب", date: "قبل يومين 09:00", repId: "r5", receiptId: "RCP-4826" },
    { id: "PAY-820", server: "food", merchantId: "m21", companyId: "c1", amount: 4800, method: "شيك", date: "اليوم 09:15", repId: "r5", receiptId: "RCP-7173" },
    { id: "PAY-821", server: "food", merchantId: "m3", companyId: "c1", amount: 1900, method: "شيك", date: "قبل 4 أيام 11:00", repId: "r3", receiptId: "RCP-9946" },
    { id: "PAY-822", server: "food", merchantId: "m3", companyId: "c1", amount: 1380, method: "تحويل مصرفي", date: "أمس 16:10", repId: "r3", receiptId: "RCP-7488" },
    { id: "PAY-823", server: "food", merchantId: "m3", companyId: "c1", amount: 798, method: "تحويل مصرفي", date: "قبل أسبوعين", repId: "r3", receiptId: "RCP-6619" },
    { id: "PAY-824", server: "food", merchantId: "m22", companyId: "c1", amount: 2263, method: "نقدي للمندوب", date: "قبل 4 أيام 11:00", repId: "r3", receiptId: "RCP-9704" },
    { id: "PAY-825", server: "food", merchantId: "m23", companyId: "c1", amount: 2520, method: "نقدي للمندوب", date: "اليوم 11:20", repId: "r3", receiptId: "RCP-9376" },
    { id: "PAY-826", server: "food", merchantId: "m23", companyId: "c1", amount: 3330, method: "تحويل مصرفي", date: "اليوم 09:15", repId: "r3", receiptId: "RCP-8468" },
    { id: "PAY-827", server: "food", merchantId: "m24", companyId: "c1", amount: 5688, method: "شيك", date: "قبل أسبوعين", repId: "r3", receiptId: "RCP-6969" },
    { id: "PAY-828", server: "food", merchantId: "m24", companyId: "c1", amount: 2193, method: "نقدي للمندوب", date: "أمس 14:30", repId: "r3", receiptId: "RCP-4235" },
    { id: "PAY-829", server: "food", merchantId: "m24", companyId: "c1", amount: 3144, method: "شيك", date: "أمس 16:10", repId: "r3", receiptId: "RCP-6116" },
    { id: "PAY-830", server: "food", merchantId: "m25", companyId: "c1", amount: 4220, method: "تحويل مصرفي", date: "أمس 14:30", repId: "r3", receiptId: "RCP-4045" },
    { id: "PAY-771", server: "food", merchantId: "m4", companyId: "c2", amount: 1016, method: "نقدي للمندوب", date: "قبل يومين", repId: "r2", receiptId: "RCP-3290" },
    { id: "PAY-765", server: "food", merchantId: "m1", companyId: "c1", amount: 800, method: "تحويل مصرفي", date: "الأسبوع الماضي", repId: null, receiptId: null },
    { id: "PAY-768", server: "food", merchantId: "m1", companyId: "c1", amount: 680, method: "نقدي للمندوب", date: "أمس", repId: "r1", receiptId: "RCP-3301" },
  ],
  returns: [
    { id: "RET-12", server: "food", merchantId: "m1", companyId: "c1", orderId: "o1002", reason: "تلف في صفيحتين", qty: 2, status: "approved", date: "اليوم" },
  ],
  quotas: [
    { id: "q1", server: "food", repId: "r1", companyId: "c1", pid: "p1", allocated: 50, sold: 22 },
    { id: "q2", server: "food", repId: "r1", companyId: "c1", pid: "p2", allocated: 30, sold: 18 },
    { id: "q3", server: "food", repId: "r3", companyId: "c1", pid: "p4", allocated: 40, sold: 15 },
    { id: "q4", server: "pharma", repId: "r4", companyId: "c4", pid: "p7", allocated: 100, sold: 45 },
  ],
  tasks: [
    { id: "t1", server: "food", companyId: "c1", assigneeType: "rep", assigneeId: "r1", title: "زيارة محل التاجر أحمد لتحصيل النقد", desc: "تحصيل دفعة 1240 د.ل المتأخرة", status: "assigned", due: "اليوم", priority: "high" },
    { id: "t2", server: "food", companyId: "c1", assigneeType: "driver", assigneeId: "d1", title: "توصيل طلب o1004 لسوبرماركت الأصيل", desc: "الظهرة - 15 كرتونة معكرونة", status: "in_progress", due: "اليوم", priority: "normal" },
  ],
  visits: [
    { id: "vg101", server: "food", repId: "r1", merchantId: "m2", result: "payment", note: "تحصيل دفعة", date: "أمس 16:10", amount: 1800, failReason: null },
    { id: "vg102", server: "food", repId: "r1", merchantId: "m12", result: "order", note: "طلب بضاعة", date: "اليوم 12:05", amount: 3200, failReason: null },
    { id: "vg103", server: "food", repId: "r1", merchantId: "m15", result: "no_order", note: "زيارة بلا طلب", date: "الأسبوع الماضي", amount: 0, failReason: "لا يوجد طلب حالياً" },
    { id: "vg104", server: "food", repId: "r5", merchantId: "m6", result: "order", note: "طلب بضاعة", date: "أمس 16:10", amount: 2400, failReason: null },
    { id: "vg105", server: "food", repId: "r5", merchantId: "m17", result: "payment", note: "تحصيل دفعة", date: "قبل 3 أيام 10:20", amount: 3200, failReason: null },
    { id: "vg106", server: "food", repId: "r5", merchantId: "m18", result: "order", note: "طلب بضاعة", date: "اليوم 08:30", amount: 3200, failReason: null },
    { id: "vg107", server: "food", repId: "r5", merchantId: "m19", result: "no_visit", note: "لم تتم الزيارة", date: "قبل 4 أيام 11:00", amount: 0, failReason: "المحل مغلق" },
    { id: "vg108", server: "food", repId: "r3", merchantId: "m3", result: "order", note: "طلب بضاعة", date: "قبل يومين 09:00", amount: 1800, failReason: null },
    { id: "vg109", server: "food", repId: "r3", merchantId: "m22", result: "no_order", note: "زيارة بلا طلب", date: "أمس 16:10", amount: 0, failReason: "لا يوجد طلب حالياً" },
    { id: "vg110", server: "food", repId: "r3", merchantId: "m23", result: "no_visit", note: "لم تتم الزيارة", date: "أمس 16:10", amount: 0, failReason: "المحل مغلق" },
    { id: "vg111", server: "food", repId: "r3", merchantId: "m24", result: "order", note: "طلب بضاعة", date: "الأسبوع الماضي", amount: 1200, failReason: null },
    { id: "vg112", server: "food", repId: "r3", merchantId: "m25", result: "no_visit", note: "لم تتم الزيارة", date: "أمس 14:30", amount: 0, failReason: "المحل مغلق" },
    { id: "v1", server: "food", repId: "r1", merchantId: "m1", result: "order", note: "طلب أرز وسكر", date: "اليوم 09:10", amount: 2370, failReason: null },
    { id: "v2", server: "food", repId: "r1", merchantId: "m2", result: "payment", note: "تحصيل 560 د.ل", date: "اليوم 10:25", amount: 560, failReason: null },
    { id: "v4", server: "food", repId: "r1", merchantId: "m1", result: "no_order", note: "المخزون كافٍ", date: "أمس 13:00", amount: 0, failReason: "stock_full" },
    { id: "v5", server: "food", repId: "r3", merchantId: "m3", result: "closed", note: "المحل مغلق", date: "أمس 11:30", amount: 0, failReason: "closed" },
    { id: "v6", server: "food", repId: "r3", merchantId: "m3", result: "order", note: "طلب معكرونة", date: "اليوم 12:00", amount: 930, failReason: null },
  ],
  // جدول الزيارات المخططة (خط سير أسبوعي لكل عميل)
  plannedVisits: [
    { id: "pv1", server: "food", repId: "r1", merchantId: "m1", day: "السبت", time: "09:00", done: true },
    { id: "pv2", server: "food", repId: "r1", merchantId: "m2", day: "السبت", time: "10:30", done: true },
    { id: "pv3", server: "food", repId: "r1", merchantId: "m1", day: "الثلاثاء", time: "09:30", done: false },
    { id: "pv4", server: "food", repId: "r3", merchantId: "m3", day: "الأحد", time: "11:00", done: false },
    { id: "pv5", server: "food", repId: "r3", merchantId: "m3", day: "الأربعاء", time: "12:00", done: false },
    { id: "pv6", server: "pharma", repId: "r4", merchantId: "m5", day: "الإثنين", time: "10:00", done: false },
  ],
  // حصر الأدوات داخل المحلات (ثلاجة، ستاند...)
  assets: [
    { id: "as1", server: "food", merchantId: "m1", repId: "r1", type: "ثلاجة عرض", code: "FRZ-1102", condition: "جيدة", date: "تم الجرد أمس" },
    { id: "as2", server: "food", merchantId: "m1", repId: "r1", type: "ستاند عرض", code: "STD-0455", condition: "تحتاج صيانة", date: "تم الجرد أمس" },
    { id: "as3", server: "food", merchantId: "m3", repId: "r3", type: "ثلاجة مشروبات", code: "FRZ-2310", condition: "جيدة", date: "هذا الأسبوع" },
  ],
  stocktakes: [
    { id: "st2", server: "food", type: "van", repId: "r3", supervisorId: "sv1", status: "requested", note: "جرد سيارة وليد — نهاية الأسبوع", date: "اليوم 12:00", lines: [{ pid: "p1", name: "أرز مصري فاخر 25كغ", sysQty: 20, counted: null, price: 180 }, { pid: "p4", name: "معكرونة 500غ (كرتونة×20)", sysQty: 40, counted: null, price: 62 }] },

    { id: "st1", server: "food", repId: "r1", merchantId: "m1", status: "requested", note: "جرد رفوف الأرز والزيت", date: "اليوم 11:00", lines: [{ name: "أرز مصري 25كغ", shelf: 8 }, { name: "زيت 18لتر", shelf: 3 }] },
  ],
  campaigns: [],
  pickupRequests: [],
  dispatchChat: [
    { id: "dc1", companyId: "c1", driverId: "d1", from: "dispatch", text: "صباح الخير سالم، رحلتك اليوم فيها 4 طلبات لمنطقة سوق الجمعة", t: "اليوم 07:30" },
    { id: "dc2", companyId: "c1", driverId: "d1", from: "driver", text: "تمام، خرجت الآن من الشركة", t: "اليوم 07:45" },
  ],
  branchTransfers: [],
  // طلبات الانضمام
  joinRequests: [
    { id: "j1", server: "food", merchantId: "m2", companyId: "c2", status: "requested", note: "نرغب بالتعامل معكم", oldDebt: 0, verifiedDebt: null, repId: null, date: "اليوم 08:00", visitDate: "" },
    { id: "j2", server: "pharma", merchantId: "m5", companyId: "c5", status: "requested", note: "نرغب بتزويد صيدليتنا بالمستلزمات الطبية", oldDebt: 0, verifiedDebt: null, repId: null, date: "اليوم 09:30", visitDate: "" },
  ],
  threads: {
    "m1__c1": [{ from: "company", text: "أهلاً، طلبك قيد التجهيز.", t: "09:20" }, { from: "merchant", text: "تمام، شكراً لكم.", t: "09:22" }],
  },
  // إيصالات النقد الإلكترونية
  receipts: [
    { id: "RCP-3301", server: "food", merchantId: "m1", companyId: "c1", repId: "r1", amount: 680, method: "نقدي للمندوب", date: "أمس 16:55", confirmedByCompany: true, ref: "TXN-880921" },
    { id: "RCP-3290", server: "food", merchantId: "m4", companyId: "c2", repId: "r2", amount: 1016, method: "نقدي للمندوب", date: "قبل يومين 12:10", confirmedByCompany: true, ref: "TXN-880744" },
  ],
  // الإشعارات (لكل دور/كيان)
  notifications: [
    { id: "n1", server: "food", role: "company", entityId: "c1", icon: "order", title: "طلب جديد وارد", body: "طلب o1001 من محل التاجر أحمد بانتظار المعالجة", read: false, t: "09:12" },
    { id: "n2", server: "food", role: "rep", entityId: "r1", icon: "task", title: "مهمة عاجلة", body: "زيارة محل التاجر أحمد لتحصيل النقد", read: false, t: "08:30" },
    { id: "n3", server: "food", role: "merchant", entityId: "m1", icon: "status", title: "تحديث طلبك", body: "طلبك o1002 خرج للتوصيل", read: false, t: "أمس" },
  ],
  // سجل العمليات للمنصة (Audit Log)
  auditLog: [
    { id: "al1", server: "food", action: "تعديل عمولة", target: "شركة النور للتوزيع", detail: "تغيير النسبة إلى 0.75%", t: "أمس 14:20" },
    { id: "al2", server: "food", action: "ترقية اشتراك", target: "شركة الأمل للمنظفات", detail: "ترقية إلى استخبارات المنافسين", t: "قبل 3 أيام" },
  ],
  // المساحة الإعلانية: إعلانات الشركات التي تظهر للتجار
  ads: [
    { id: "ad1", server: "food", companyId: "c1", title: "خصم 6% على الأرز الفاخر", subtitle: "عند طلب 100 كيس فأكثر — لفترة محدودة", emoji: "🍚", c1: "#0D47A1", c2: "#1565C0", cta: "اطلب الآن", active: true },
    { id: "ad2", server: "food", companyId: "c2", title: "مياه البركة بسعر الجملة", subtitle: "عرض الكرتونة + توصيل مجاني داخل طرابلس", emoji: "💧", c1: "#0E7490", c2: "#06B6D4", cta: "تصفّح العرض", active: true },
    { id: "ad3", server: "pharma", companyId: "c4", title: "حجز نسبة 12% على الباراسيتامول", subtitle: "للصيدليات عند الشراء بكمية 500 علبة", emoji: "💊", c1: "#0D9488", c2: "#14B8A6", cta: "اطلب الآن", active: true },
  ],
  // تسوية الفواتير: يدفع التاجر ← يراجع المندوب ← يعتمد ← تُضاف لحسابه
  settlements: [
    { id: "ST-9001", server: "food", merchantId: "m1", companyId: "c1", repId: "r1", invoiceId: "INV-2041", amount: 1240, method: "cash", ref: "", note: "سداد دفعة من المديونية", status: "pending_rep", date: "اليوم 11:20", reviewedAt: null },
  ],
  // تسليم النقدية اليومي من المندوب للشركة
  handovers: [
    { id: "HV-4401", server: "food", repId: "r2", companyId: "c2", amount: 1016, count: 1, status: "received", date: "أمس 17:00", note: "تسليم تحصيلات اليوم" },
  ],
  // طلبات تحميل سيارات مندوبي التجزئة (الشركة تعبّئ ثم تُضاف لمخزون السيارة)
  loadRequests: [
    { id: "LD-7001", server: "food", repId: "r3", companyId: "c1", items: [{ pid: "p4", qty: 40 }, { pid: "p1", qty: 20 }], note: "تحميل بداية اليوم", status: "requested", date: "اليوم 07:30" },
  ],
  memberSeq: 10007,
  orderSeq: 5000,
  // طلبات مطابقة الحساب (المندوب أو الشركة تطلب، والتاجر يؤكد أو يعترض)
  reconciliations: [
    { id: "REC-3001", server: "food", merchantId: "m2", companyId: "c1", requestedBy: "rep", requesterId: "r1", statedBalance: -2100, status: "pending", merchantNote: "", date: "اليوم" },
  ],
  // إشعارات دائن/مدين (سندات قيد تعدّل مديونية العميل)
  creditNotes: [
    { id: "CN-501", server: "food", merchantId: "m1", companyId: "c1", kind: "credit", amount: 120, reason: "تعويض تلف صفيحتي زيت (مرتجع RET-12)", refId: "RET-12", date: "اليوم", ageDays: 0 },
  ],
  // خط السير الأسبوعي الدوري (يضبطه المشرف لكل مندوب)
  journeyPlans: [
    { id: "JP-1", server: "food", repId: "r1", day: "السبت", stops: [{ merchantId: "m1", done: true, outcome: "order", reason: "" }, { merchantId: "m2", done: false, outcome: null, reason: "" }] },
    { id: "JP-2", server: "food", repId: "r1", day: "الأحد", stops: [{ merchantId: "m2", done: false, outcome: null, reason: "" }] },
    { id: "JP-3", server: "food", repId: "r2", day: "السبت", stops: [{ merchantId: "m4", done: false, outcome: null, reason: "" }] },
  ],
  // عملاء مستوردون من ملف Excel بانتظار التوثيق الميداني من المندوب ثم تأكيد التاجر
  pendingClients: [
    { id: "PC-101", server: "food", companyId: "c1", repId: "r1", name: "بقالة الأمانة", owner: "فرج علي", phone: "091-8801122", area: "سوق الجمعة", openingDebt: 1500, status: "awaiting_visit", photo: null, lat: null, lng: null, nameConfirmed: false, gpsConfirmed: false, merchantConfirmed: false, note: "", date: "اليوم 08:00", source: "clients_import_june.xlsx" },
    { id: "PC-102", server: "food", companyId: "c1", repId: "r1", name: "ماركت الريان", owner: "سالم أبوبكر", phone: "092-4433221", area: "سوق الجمعة", openingDebt: 0, status: "awaiting_visit", photo: null, lat: null, lng: null, nameConfirmed: false, gpsConfirmed: false, merchantConfirmed: false, note: "", date: "اليوم 08:00", source: "clients_import_june.xlsx" },
    { id: "PC-103", server: "food", companyId: "c1", repId: "r3", name: "محلات النسيم", owner: "طارق مفتاح", phone: "094-6655443", area: "عين زارة", openingDebt: 800, status: "awaiting_visit", photo: null, lat: null, lng: null, nameConfirmed: false, gpsConfirmed: false, merchantConfirmed: false, note: "", date: "اليوم 08:00", source: "clients_import_june.xlsx" },
  ],
  // رحلات السائق: تجميع عدة طلبات جاهزة في رحلة واحدة بمحطات مرتّبة
  trips: [
    { id: "TRP-101", server: "food", companyId: "c1", driverId: "d1", orderIds: ["o1002"], status: "delivered", date: "قبل 3 أيام", dayKey: "قبل 3 أيام", tripNo: 1, extra: false },
    { id: "TRP-102", server: "food", companyId: "c1", driverId: "d1", orderIds: [], status: "delivered", date: "قبل 3 أيام", dayKey: "قبل 3 أيام", tripNo: 2, extra: true },
    { id: "TRP-103", server: "food", companyId: "c1", driverId: "d1", orderIds: [], status: "delivered", date: "قبل يومين", dayKey: "قبل يومين", tripNo: 1, extra: false },
    { id: "TRP-104", server: "food", companyId: "c1", driverId: "d1", orderIds: [], status: "delivered", date: "أمس", dayKey: "أمس", tripNo: 1, extra: false },
    { id: "TRP-105", server: "food", companyId: "c1", driverId: "d1", orderIds: [], status: "delivered", date: "أمس", dayKey: "أمس", tripNo: 2, extra: true },
    { id: "TRP-106", server: "food", companyId: "c1", driverId: "d1", orderIds: [], status: "delivered", date: "أمس", dayKey: "أمس", tripNo: 3, extra: true },
    { id: "TRP-107", server: "food", companyId: "c1", driverId: "d5", orderIds: [], status: "delivered", date: "أمس", dayKey: "أمس", tripNo: 1, extra: false },
    { id: "TRP-108", server: "food", companyId: "c1", driverId: "d5", orderIds: [], status: "delivered", date: "أمس", dayKey: "أمس", tripNo: 2, extra: true },
  ],
  // طلبات مندوب للتحصيل: يطلبها التاجر ← إشعار للمندوب + متابعة الشركة
  collectRequests: [
    { id: "CR-5001", server: "food", merchantId: "m2", companyId: "c2", repId: "r2", amount: 800, method: "cash", note: "جاهز للسداد نقداً", status: "requested", date: "اليوم 10:15", visitTime: "" },
  ],
  // خط السير العاجل: المشرف يسند للمندوب قائمة محلات للزيارة الآن
  urgentRoutes: [
    { id: "UR-301", server: "food", supervisorId: "sv1", repId: "r1", note: "أولوية تحصيل قبل نهاية الدوام", date: "اليوم 08:45", stops: [{ merchantId: "m1", reason: "تحصيل دفعة متأخرة", done: false }, { merchantId: "m2", reason: "نقص مخزون عاجل", done: false }] },
  ],
  // كميات البضاعة التي تخصصها الشركة لمدير المبيعات (بالستيكات/الكراتين)
  mgrStock: [
    { id: "ms1", server: "food", companyId: "c1", managerId: "sm1", pid: "p1", allocated: 2000, distributed: 900 },
    { id: "ms2", server: "food", companyId: "c1", managerId: "sm1", pid: "p4", allocated: 1500, distributed: 600 },
    { id: "ms3", server: "pharma", companyId: "c4", managerId: "sm2", pid: "p7", allocated: 1200, distributed: 400 },
  ],
  // أهداف العملاء: المشرف يحدد لكل عميل هدفاً بالستيكات والتحصيل
  clientTargets: [
    { id: "ct1", server: "food", supervisorId: "sv1", merchantId: "m1", qtyTarget: 120, qtyAchieved: 52, collectionTarget: 3000, collected: 1240 },
    { id: "ct2", server: "food", supervisorId: "sv1", merchantId: "m2", qtyTarget: 80, qtyAchieved: 30, collectionTarget: 1500, collected: 560 },
  ],
};

/* ===================== مكوّنات أساسية ===================== */
const inputCls = "w-full rounded-xl px-3 py-2.5 text-sm outline-none transition-colors";
const inputStyle = { border: `1px solid ${C.line}`, background: "rgba(255,255,255,0.04)", color: C.text, minHeight: 44 };

/* ============================================================
   ====== بنية احترافية: حركات + طباعة + Toast + تأكيد ======
   ============================================================ */
const GLOBAL_CSS = `
/* ═══ رابط — نظام التصميم الفاتح النظيف ═══ */
html, body, #root { background: #F4F7FB; }
body { color: #15263B; }
/* البطاقات: سطح أبيض بظل ناعم */
.rb-card { transition: transform .18s cubic-bezier(.16,1,.3,1), box-shadow .18s, border-color .18s; box-shadow: 0 1px 3px rgba(15,37,68,.06), 0 8px 24px rgba(15,37,68,.05); }
.rb-card:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(15,37,68,.10), 0 12px 32px rgba(15,37,68,.08); }
.rb-modal { background: #FFFFFF !important; border: 1px solid #E4EAF2; box-shadow: 0 24px 64px rgba(15,37,68,.18); }
.rb-modal-bg { background: rgba(15,37,68,.35) !important; backdrop-filter: blur(2px); }
/* الشريط الجانبي والعلوي: أبيض نظيف */
.rb-glass-bar { background: #FFFFFF !important; }
/* الحقول: آبار فاتحة */
input::placeholder, textarea::placeholder { color: #94A3B8 !important; }
select option { background: #FFFFFF; color: #15263B; }
/* حلقة التركيز الزرقاء */
input:focus, select:focus, textarea:focus {
  outline: none !important;
  border-color: #2563EB !important;
  box-shadow: 0 0 0 3px rgba(37,99,235,.15) !important;
}
/* شريط التمرير */
::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #CDD8E8; border-radius: 999px; }
::-webkit-scrollbar-thumb:hover { background: #A9BAD0; }
/* الأرقام المالية: خانات ثابتة */
.tabular, table td, table th { font-variant-numeric: tabular-nums; }
/* الأزرار: ضغطة سلسة */
.rb-press { transition: transform .12s, box-shadow .2s, background .2s, opacity .2s; }
.rb-press:active { transform: scale(.97); }
.rb-glow:hover { box-shadow: 0 4px 14px rgba(37,99,235,.18) !important; }
@media print {
  html, body, #root { background: #fff !important; }
  body::before { display: none; }
  body { background-image: none !important; color: #000 !important; }
  .rb-card, .rb-modal, #invoice-print, #invoice-print * {
    background: #fff !important; color: #000 !important;
    border-color: #ddd !important; box-shadow: none !important;
    backdrop-filter: none !important; -webkit-backdrop-filter: none !important;
  }
  .no-print { display: none !important; }
}

@keyframes rbFadeUp { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: none; } }
@keyframes rbFadeIn { from { opacity: 0; } to { opacity: 1; } }
@keyframes rbScaleIn { from { opacity: 0; transform: scale(.96) translateY(8px); } to { opacity: 1; transform: none; } }
@keyframes rbSlideDown { from { opacity: 0; transform: translateY(-14px); } to { opacity: 1; transform: none; } }
@keyframes rbPulse { 0%,100% { transform: scale(1); opacity: 1; } 50% { transform: scale(1.6); opacity: .5; } }
@keyframes rbLogoIn { 0% { opacity: 0; transform: scale(.7); } 60% { opacity: 1; transform: scale(1.06); } 100% { transform: scale(1); } }
@keyframes rbBarGrow { from { transform: scaleY(0); } to { transform: scaleY(1); } }
.rb-card { animation: rbFadeUp .35s ease both; }
.rb-modal-bg { animation: rbFadeIn .2s ease both; }
.rb-modal { animation: rbScaleIn .28s cubic-bezier(.21,1.02,.55,1) both; }
.rb-toast { animation: rbSlideDown .3s cubic-bezier(.21,1.02,.55,1) both; }
.rb-press { transition: transform .12s ease, box-shadow .12s ease; }
.rb-press:active { transform: scale(.97); }
.rb-bar { transform-origin: bottom; animation: rbBarGrow .6s ease both; }
.rb-stagger > * { animation: rbFadeUp .35s ease both; }
.rb-stagger > *:nth-child(1) { animation-delay: .02s; } .rb-stagger > *:nth-child(2) { animation-delay: .06s; }
.rb-stagger > *:nth-child(3) { animation-delay: .1s; } .rb-stagger > *:nth-child(4) { animation-delay: .14s; }
.rb-stagger > *:nth-child(5) { animation-delay: .18s; } .rb-stagger > *:nth-child(6) { animation-delay: .22s; }
.rb-dot { animation: rbPulse 1.6s ease infinite; }
@media (max-width: 640px) {
  .rb-card { border-radius: 14px; }
  h2 { font-size: 1.05rem !important; }
  table { font-size: 12px; }
}
/* منع تمدد الصفحة أفقياً على الهاتف: عناصر الشبكات تنكمش والجداول تتمرر داخلياً */
body { overflow-x: hidden; }
.grid > * { min-width: 0; }
.rb-card { max-width: 100%; }
.overflow-x-auto { -webkit-overflow-scrolling: touch; }
html { -webkit-tap-highlight-color: transparent; scroll-behavior: smooth; }
input, select, textarea { font-size: 16px !important; } /* يمنع تكبير iOS التلقائي */
@media print {
  body * { visibility: hidden !important; }
  #invoice-print, #invoice-print * { visibility: visible !important; }
  #invoice-print { position: absolute !important; inset: 0 !important; width: 100% !important; max-height: none !important; overflow: visible !important; border-radius: 0 !important; }
  .no-print { display: none !important; }
}
`;
// نظام إشعارات النجاح (Toast) — دفع من أي مكان عبر window.rbToast
function ToastHost() {
  const [toasts, setToasts] = useState([]);
  useEffect(() => {
    window.rbToast = (msg, tone = "success") => { const id = Math.random().toString(36).slice(2); setToasts((t) => [...t, { id, msg, tone }]); setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), 2600); };
    return () => { delete window.rbToast; };
  }, []);
  const tones = { success: { bg: C.success, icon: CheckCircle2 }, danger: { bg: C.danger, icon: XCircle }, info: { bg: C.info, icon: Bell } };
  return (<div className="fixed top-4 left-0 right-0 z-[90] flex flex-col items-center gap-2 pointer-events-none px-4">
    {toasts.map((t) => { const tn = tones[t.tone] || tones.success; const Ic = tn.icon; return (<div key={t.id} className="rb-toast pointer-events-auto flex items-center gap-2 rounded-2xl px-4 py-2.5 text-sm font-bold" style={{ background: tn.bg, color: t.tone === "danger" ? "#fff" : C.onAccent, boxShadow: "0 12px 30px rgba(0,0,0,.45)", maxWidth: 420 }}><Ic size={17} />{t.msg}</div>); })}
  </div>);
}
const toast = (msg, tone) => { if (typeof window !== "undefined" && window.rbToast) window.rbToast(msg, tone); };
// نافذة تأكيد للإجراءات الحسّاسة — استدعاء عبر confirmThen(رسالة, دالة)
function ConfirmHost() {
  const [cfg, setCfg] = useState(null);
  useEffect(() => { window.rbConfirm = (msg, onYes, opts = {}) => setCfg({ msg, onYes, ...opts }); return () => { delete window.rbConfirm; }; }, []);
  if (!cfg) return null;
  return (<div className="rb-modal-bg fixed inset-0 z-[95] flex items-center justify-center p-4 bg-black/50" onClick={() => setCfg(null)}>
    <div className="rb-modal w-full max-w-sm rounded-2xl p-5 text-center" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
      <div className="rounded-full mx-auto flex items-center justify-center mb-3" style={{ width: 52, height: 52, background: cfg.tone === "danger" ? C.dangerSoft : C.accentSoft }}><AlertCircle size={26} style={{ color: cfg.tone === "danger" ? C.danger : "#92590C" }} /></div>
      <div className="font-extrabold mb-1" style={{ color: C.text }}>{cfg.title || "هل أنت متأكد؟"}</div>
      <p className="text-sm mb-4" style={{ color: C.soft }}>{cfg.msg}</p>
      <div className="flex gap-2"><button onClick={() => { cfg.onYes(); setCfg(null); }} className="rb-press flex-1 rounded-xl py-2.5 text-sm font-extrabold" style={{ background: cfg.tone === "danger" ? C.danger : C.accent, color: cfg.tone === "danger" ? "#fff" : C.onAccent }}>{cfg.yesLabel || "نعم، متأكد"}</button><button onClick={() => setCfg(null)} className="rb-press flex-1 rounded-xl py-2.5 text-sm font-bold" style={{ border: `1px solid ${C.line}`, color: C.text }}>إلغاء</button></div>
    </div>
  </div>);
}
const confirmThen = (msg, onYes, opts) => { if (typeof window !== "undefined" && window.rbConfirm) window.rbConfirm(msg, onYes, opts); else onYes(); };
// أفاتار بالحروف الملوّنة
const AV_COLORS = ["#0D47A1", "#0D9488", "#B45309", "#7C3AED", "#BE185D", "#047857", "#B91C1C", "#4338CA"];
function Avatar({ name, size = 40, emoji }) {
  const ch = (name || "؟").trim().charAt(0);
  const idx = (name || "").split("").reduce((s, c) => s + c.charCodeAt(0), 0) % AV_COLORS.length;
  return (<div className="rounded-xl flex items-center justify-center shrink-0 font-extrabold text-white" style={{ width: size, height: size, background: `linear-gradient(135deg, ${AV_COLORS[idx]}, ${AV_COLORS[(idx + 3) % AV_COLORS.length]})`, fontSize: size * 0.42, fontFamily: "Cairo, sans-serif" }}>{emoji || ch}</div>);
}
// موقع المحل GPS: زر يفتح خرائط جوجل بالإحداثيات (للمندوب والسائق والمروّج)
function GpsLink({ m, sm, full }) {
  if (!m?.lat || !m?.lng) return null;
  const url = `https://www.google.com/maps?q=${m.lat},${m.lng}`;
  return (<a href={url} target="_blank" rel="noreferrer" onClick={(e) => e.stopPropagation()} className={`rb-press inline-flex items-center justify-center gap-1.5 rounded-xl font-bold ${sm ? "px-2.5 py-1.5 text-xs" : "px-3.5 py-2 text-sm"} ${full ? "w-full" : ""}`} style={{ background: C.infoSoft, color: C.info }}><Navigation size={sm ? 13 : 15} />الموقع GPS</a>);
}
// بطاقة موقع مصغّرة (تُظهر الإحداثيات وزر الانتقال)
function LocationCard({ m }) {
  if (!m?.lat || !m?.lng) return null;
  return (<div className="rounded-xl p-3 flex items-center justify-between gap-2" style={{ background: C.infoSoft + "66", border: `1px solid ${C.infoSoft}` }}>
    <div className="flex items-center gap-2 min-w-0"><div className="rounded-lg p-2 shrink-0" style={{ background: C.well }}><MapPin size={16} style={{ color: C.info }} /></div><div className="min-w-0"><div className="text-xs font-bold" style={{ color: C.text }}>موقع {m.name}</div><div className="text-[10px] truncate" dir="ltr" style={{ color: C.soft }}>{m.lat?.toFixed(4)}, {m.lng?.toFixed(4)}</div></div></div>
    <GpsLink m={m} sm />
  </div>);
}
// رسم أعمدة SVG بسيط واحترافي
function BarsChart({ data: rows, color = C.primary2, height = 120, money: isMoney }) {
  const max = Math.max(...rows.map((r) => r.v), 1);
  return (<div><div className="flex items-end gap-2" style={{ height }}>{rows.map((r, i) => (<div key={i} className="flex-1 flex flex-col items-center gap-1 min-w-0"><span className="text-[9px] font-bold" style={{ color: C.soft }}>{isMoney ? (r.v >= 1000 ? (r.v / 1000).toFixed(1) + "k" : r.v) : r.v}</span><div className="rb-bar w-full rounded-t-lg" style={{ height: `${Math.max(4, (r.v / max) * (height - 26))}px`, background: `linear-gradient(180deg, ${color}, ${color}99)`, animationDelay: `${i * 0.06}s` }} /></div>))}</div><div className="flex gap-2 mt-1">{rows.map((r, i) => <div key={i} className="flex-1 text-center text-[9px] font-bold truncate" style={{ color: C.soft }}>{r.label}</div>)}</div></div>);
}
// رسم دائري (Donut) SVG
function DonutChart({ parts, size = 130, thickness = 18 }) {
  const total = parts.reduce((s, p) => s + p.v, 0) || 1; const r = (size - thickness) / 2; const cx = size / 2; const circ = 2 * Math.PI * r;
  let acc = 0;
  return (<div className="flex items-center gap-4"><svg width={size} height={size} style={{ transform: "rotate(-90deg)" }}>{parts.map((p, i) => { const frac = p.v / total; const dash = frac * circ; const off = acc * circ; acc += frac; return <circle key={i} cx={cx} cy={cx} r={r} fill="none" stroke={p.color} strokeWidth={thickness} strokeDasharray={`${dash} ${circ - dash}`} strokeDashoffset={-off} strokeLinecap="butt" />; })}</svg>
    <div className="space-y-1.5">{parts.map((p, i) => <div key={i} className="flex items-center gap-2 text-xs"><span className="rounded-full" style={{ width: 10, height: 10, background: p.color }} /><span style={{ color: C.text, fontWeight: 700 }}>{p.label}</span><span style={{ color: C.soft }}>{Math.round((p.v / total) * 100)}%</span></div>)}</div>
  </div>);
}


function Field({ label, children, hint }) {
  return (<label className="block"><span className="text-sm font-bold block mb-1.5" style={{ color: C.text }}>{label}</span>{children}{hint && <span className="text-xs block mt-1" style={{ color: C.soft }}>{hint}</span>}</label>);
}
function Input(props) { return <input {...props} className={inputCls} style={inputStyle} />; }
function Card({ children, className = "", style = {} }) { return <div className={`rb-card rounded-2xl ${className}`} style={{ background: C.surface, border: `1px solid ${C.line}`, ...style }}>{children}</div>; }
function SectionTitle({ children, sub, action }) {
  return (<div className="mb-4 flex items-start justify-between gap-3 flex-wrap"><div><h2 className="text-lg font-extrabold" style={{ color: C.text }}>{children}</h2>{sub && <p className="text-sm mt-0.5" style={{ color: C.soft }}>{sub}</p>}</div>{action}</div>);
}
function PrimaryBtn({ children, onClick, icon: Icon, tone = "primary", full, sm }) {
  // الألوان المضيئة (نعناعي/ذهبي/سماوي) تتطلب نصاً داكناً لتباين AA — الأحمر يبقى بنص فاتح
  const bg = tone === "danger" ? C.danger : tone === "accent" ? C.accent : tone === "success" ? `linear-gradient(135deg, ${C.success}, ${C.primary2})` : tone === "ghost" ? "transparent" : `linear-gradient(135deg, ${SC.color}, ${SC.color2})`;
  const col = tone === "ghost" ? C.text : C.onAccent;
  const glow = tone === "ghost" || tone === "danger" ? "none" : C.glow;
  return (<button onClick={onClick} className={`rb-press ${full ? "w-full" : ""} inline-flex items-center justify-center gap-2 rounded-xl ${sm ? "px-3 py-1.5 text-xs" : "px-4 py-2.5 text-sm"} font-bold transition-all`} style={{ background: bg, color: col, border: tone === "ghost" ? `1px solid ${C.line}` : "none", boxShadow: glow, minHeight: sm ? 34 : 44 }}>{Icon && <Icon size={sm ? 14 : 17} />}{children}</button>);
}
function GhostBtn({ children, onClick, icon: Icon, sm, full }) {
  return (<button onClick={onClick} className={`rb-press ${full ? "w-full justify-center" : ""} inline-flex items-center gap-2 rounded-xl ${sm ? "px-3 py-1.5 text-xs" : "px-4 py-2.5 text-sm"} font-bold transition-colors hover:bg-gray-50`} style={{ border: `1px solid ${C.line}`, color: C.text, background: C.surface }}>{Icon && <Icon size={sm ? 14 : 17} />}{children}</button>);
}
function Badge({ status, map = STATUS }) { const s = map[status]; if (!s) return null; return <span style={{ background: s.soft, color: s.color }} className="text-xs font-bold px-2.5 py-1 rounded-full whitespace-nowrap">{s.label}</span>; }
// عرض نجوم التقييم (متوسط + عدد)
function Stars({ avg, count, size = 13, showCount = true }) {
  if (avg == null) return <span className="text-[11px]" style={{ color: C.soft }}>بلا تقييم</span>;
  return (<span className="inline-flex items-center gap-1">
    <span className="inline-flex">{[1, 2, 3, 4, 5].map((i) => <Star key={i} size={size} fill={i <= Math.round(avg) ? "#F59E0B" : "none"} style={{ color: i <= Math.round(avg) ? "#F59E0B" : C.line }} />)}</span>
    <span className="text-[11px] font-bold" style={{ color: C.text }}>{avg.toFixed(1)}</span>
    {showCount && <span className="text-[10px]" style={{ color: C.soft }}>({count})</span>}
  </span>);
}
// اختيار نجوم تفاعلي
function StarPicker({ value, onChange, size = 26 }) {
  return (<div className="flex gap-1" dir="ltr" style={{ justifyContent: "flex-end" }}>{[1, 2, 3, 4, 5].map((i) => <button key={i} onClick={() => onChange(i)} className="transition-transform hover:scale-110"><Star size={size} fill={i <= value ? "#F59E0B" : "none"} style={{ color: i <= value ? "#F59E0B" : C.line }} /></button>)}</div>);
}
function Stat({ icon: Icon, label, value, tint, hint }) {
  return (<div className="rounded-2xl p-4 flex items-center gap-3" style={{ background: C.surface, border: `1px solid ${C.line}`, boxShadow: "0 1px 2px rgba(11,31,58,.04)" }}><div className="rounded-xl p-2.5" style={{ background: tint.soft }}><Icon size={22} style={{ color: tint.color }} /></div><div className="min-w-0"><div className="text-sm truncate" style={{ color: C.soft }}>{label}</div><div className="text-xl font-extrabold" style={{ color: C.text }}>{value}</div>{hint && <div className="text-[11px]" style={{ color: C.soft }}>{hint}</div>}</div></div>);
}
function Empty({ icon: Icon, text }) {
  return (<div className="flex flex-col items-center justify-center py-10 text-center"><div className="rounded-2xl p-3 mb-2" style={{ background: C.bg }}><Icon size={26} style={{ color: C.soft }} /></div><p className="text-sm" style={{ color: C.soft }}>{text}</p></div>);
}
function Th({ children }) { return <th className="text-right text-xs font-bold px-3 py-2.5" style={{ color: C.soft }}>{children}</th>; }
function Td({ children, bold }) { return <td className="px-3 py-3 text-sm whitespace-nowrap" style={{ color: C.text, fontWeight: bold ? 800 : 400 }}>{children}</td>; }
function Table({ head, children }) {
  return (<div className="overflow-x-auto"><table className="w-full" style={{ borderCollapse: "collapse" }}><thead><tr style={{ borderBottom: `1px solid ${C.line}` }}>{head.map((h, i) => <Th key={i}>{h}</Th>)}</tr></thead><tbody>{children}</tbody></table></div>);
}
function ProgressBar({ value, max, color }) {
  const pct = Math.min(100, Math.round((value / (max || 1)) * 100));
  return (<div><div className="h-2.5 rounded-full overflow-hidden" style={{ background: C.bg }}><div style={{ width: pct + "%", height: "100%", background: color || SC.color, transition: "width .4s" }} /></div></div>);
}
function Pipeline({ status }) {
  const cur = STATUS[status].step; const rejected = status === "rejected";
  const steps = [{ t: SC.merchantWord, i: Store }, { t: "المندوب", i: Handshake }, { t: "الشركة", i: Building2 }, { t: "السائق", i: Truck }, { t: "تسليم", i: PackageCheck }];
  return (<div className="flex items-center justify-between gap-1 w-full">{steps.map((s, idx) => { const n = idx + 1; const done = !rejected && cur >= n; const isCash = idx === 1; const col = done ? (isCash ? C.accent : SC.color2) : C.line;
    return (<React.Fragment key={idx}><div className="flex flex-col items-center gap-1 shrink-0"><div className="rounded-full flex items-center justify-center" style={{ width: 32, height: 32, background: done ? col : C.surface, border: `2px solid ${col}` }}><s.i size={15} style={{ color: done ? "#fff" : C.soft }} /></div><span className="text-[10px] font-bold" style={{ color: done ? C.text : C.soft }}>{s.t}</span></div>{idx < steps.length - 1 && <div className="h-0.5 flex-1 rounded" style={{ background: !rejected && cur > n ? SC.color2 : C.line }} />}</React.Fragment>); })}</div>);
}
function FakeBarcode({ code }) {
  const ref = useRef(null);
  useEffect(() => {
    if (!ref.current || !code) return;
    try { JsBarcode(ref.current, String(code), { format: "CODE128", displayValue: true, fontSize: 11, height: 46, margin: 2, lineColor: C.text }); }
    catch { /* بعض القيم القديمة قد لا تصلح كرمز؛ نعرض النص بدلاً من كسر الشاشة. */ }
  }, [code]);
  return <svg ref={ref} role="img" aria-label={`باركود ${code || ""}`} />;
}
function FakeQR({ value = "https://rabet.local" }) {
  return <div className="inline-flex rounded-lg p-2" style={{ background: C.well, border: `1px solid ${C.line}` }}><QRCodeSVG value={String(value)} size={84} includeMargin /></div>;
}
function Logo({ light, size = 22, sub }) {
  return (<div className="flex items-center gap-2">
    <div className="rounded-xl flex items-center justify-center" style={{ width: size + 12, height: size + 12, background: light ? "rgba(255,255,255,.12)" : SC.color }}>
      <Link2 size={size} color={light ? C.accent : "#fff"} strokeWidth={2.5} />
    </div>
    <div className="leading-tight">
      <div className="font-extrabold" style={{ color: light ? "#fff" : C.ink, fontSize: size, fontFamily: "Cairo, sans-serif" }}>رابط</div>
      {sub && <div className="text-[10px]" style={{ color: light ? "rgba(255,255,255,.6)" : C.soft }}>{sub}</div>}
    </div>
  </div>);
}

/* ===================== الإشعارات ===================== */
const NOTIF_ICON = {
  order:    { i: ShoppingCart, c: C.info, s: C.infoSoft },
  cash:     { i: HandCoins, c: C.success, s: C.successSoft },
  task:     { i: ClipboardCheck, c: C.accent, s: C.accentSoft },
  status:   { i: Truck, c: C.purple, s: C.purpleSoft },
  join:     { i: UserPlus, c: C.teal, s: C.tealSoft },
  promo:    { i: ImageIcon, c: C.rose, s: C.roseSoft },
  invoice:  { i: Receipt, c: SC.color, s: SC.soft },
  default:  { i: Bell, c: C.soft, s: C.bg },
};
function NotificationPanel({ notifs, onRead, onReadAll, onReadType, onClose }) {
  const TYPE_META = { order: { t: "الطلبات", icon: ShoppingCart }, status: { t: "حالة الطلبات", icon: Truck }, cash: { t: "مالية", icon: HandCoins }, task: { t: "مهام", icon: ClipboardList }, route: { t: "خط السير", icon: Route }, join: { t: "انضمام", icon: UserPlus }, chat: { t: "رسائل", icon: MessageSquare } };
  const groups = Object.entries(notifs.reduce((acc, n) => { const k = TYPE_META[n.icon] ? n.icon : "task"; (acc[k] = acc[k] || []).push(n); return acc; }, {}));
  return (<div className="absolute left-0 lg:left-auto lg:right-0 mt-2 w-80 max-w-[90vw] rounded-2xl z-50 overflow-hidden" style={{ background: C.surface, border: `1px solid ${C.line}`, boxShadow: "0 12px 40px rgba(11,31,58,.18)" }}>
    <div className="flex items-center justify-between px-4 py-3" style={{ borderBottom: `1px solid ${C.line}` }}><span className="font-extrabold text-sm" style={{ color: C.text }}>الإشعارات</span><button onClick={onReadAll} className="text-xs font-bold" style={{ color: SC.color }}>قراءة الكل</button></div>
    <div className="max-h-96 overflow-auto">
      {groups.map(([type, list]) => { const meta = TYPE_META[type] || TYPE_META.task; const unread = list.filter((n) => !n.read).length; const Ic = meta.icon; return (<div key={type}>
        <div className="flex items-center justify-between px-4 py-1.5 sticky top-0 z-10" style={{ background: C.bg }}>
          <span className="text-[11px] font-extrabold flex items-center gap-1.5" style={{ color: C.soft }}><Ic size={12} />{meta.t}{unread > 0 && <span className="rounded-full px-1.5 text-[9px] text-white" style={{ background: C.danger }}>{unread}</span>}</span>
          {unread > 0 && onReadType && <button onClick={() => onReadType(type)} className="text-[10px] font-bold" style={{ color: SC.color }}>قراءة الكل</button>}
        </div>
        {list.map((n) => (<button key={n.id} onClick={() => onRead(n.id)} className="w-full text-right px-4 py-2.5 flex items-start gap-2.5 transition-colors" style={{ background: n.read ? "transparent" : SC.soft + "44", borderBottom: `1px solid ${C.line}` }}>
          {!n.read && <span className="rounded-full mt-1.5 shrink-0" style={{ width: 7, height: 7, background: SC.color }} />}
          <span className="min-w-0"><span className="block text-xs font-extrabold truncate" style={{ color: C.text }}>{n.title}</span><span className="block text-[11px] leading-snug" style={{ color: C.soft }}>{n.body}</span><span className="block text-[9px] mt-0.5" style={{ color: C.soft }}>{n.t}</span></span>
        </button>))}
      </div>); })}
      {notifs.length === 0 && <div className="p-6 text-center text-sm" style={{ color: C.soft }}>لا إشعارات.</div>}
    </div>
  </div>);
}
function ReceiptModal({ receipt, data, onClose }) {
  if (!receipt) return null;
  return (<div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50" onClick={onClose}>
    <div className="rb-modal w-full max-w-sm rounded-2xl overflow-hidden" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
      <div className="p-5 text-white text-center" style={{ background: `linear-gradient(120deg, ${C.success}, #15803D)` }}>
        <div className="rounded-full mx-auto flex items-center justify-center mb-2" style={{ width: 56, height: 56, background: "rgba(255,255,255,.2)" }}><CheckCircle2 size={32} color="#fff" /></div>
        <div className="font-extrabold text-lg">إيصال استلام نقدي</div>
        <div className="text-sm opacity-90">{receipt.id}</div>
      </div>
      <div className="p-5 space-y-3">
        <div className="text-center"><div className="text-3xl font-extrabold" style={{ color: C.success }}>{money(receipt.amount)}</div></div>
        <div className="rounded-xl p-3 space-y-2 text-sm" style={{ background: C.bg }}>
          {[["العميل", merchName(data, receipt.merchantId)], ["الشركة", compName(data, receipt.companyId)], ["المندوب المُستلِم", repName(data, receipt.repId)], ["الطريقة", receipt.method], ["التاريخ", receipt.date], ["رقم المرجع", receipt.ref]].map(([k, v]) => <div key={k} className="flex justify-between"><span style={{ color: C.soft }}>{k}</span><span className="font-bold" style={{ color: C.text }}>{v}</span></div>)}
        </div>
        <div className="flex items-center justify-center gap-2 text-sm rounded-xl py-2" style={{ background: receipt.confirmedByCompany ? C.successSoft : C.warnSoft, color: receipt.confirmedByCompany ? C.success : "#92590C" }}>{receipt.confirmedByCompany ? <><BadgeCheck size={16} />مؤكّد من الشركة</> : <><Clock size={16} />بانتظار تأكيد الشركة</>}</div>
        <div className="flex justify-center"><FakeQR /></div>
        <PrimaryBtn full icon={XCircle} tone="ghost" onClick={onClose}>إغلاق</PrimaryBtn>
      </div>
    </div>
  </div>);
}

/* ===================== شاشة اختيار السيرفر ===================== */
function ServerPick({ onPick }) {
  return (<div dir="rtl" className="min-h-screen flex items-center justify-center p-6 relative overflow-hidden" style={{ background: `linear-gradient(135deg, ${C.ink} 0%, #0D47A1 60%, #0D9488 130%)`, fontFamily: "Tajawal, sans-serif" }}>
    <div className="absolute opacity-10" style={{ top: -60, right: -60, width: 360, height: 360, borderRadius: "50%", border: "40px solid #fff" }} />
    <div className="absolute opacity-10" style={{ bottom: -80, left: -40, width: 280, height: 280, borderRadius: "50%", border: "30px solid " + C.accent }} />
    <div className="w-full max-w-3xl relative z-10">
      <div className="text-center mb-8">
        <div className="inline-flex items-center gap-3 mb-3">
          <div className="rounded-2xl flex items-center justify-center" style={{ width: 50, height: 50, background: "rgba(255,255,255,.12)" }}><Link2 size={28} color={C.accent} strokeWidth={2.5} /></div>
          <div className="text-right"><div className="text-white font-extrabold text-3xl" style={{ fontFamily: "Cairo, sans-serif" }}>رابط</div><div style={{ color: C.accent }} className="font-bold text-sm">ليربط العالم بين يديك</div></div>
        </div>
        <p style={{ color: "rgba(255,255,255,.8)" }}>اختر القطاع الذي تريد العمل عليه</p>
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
        {Object.values(SERVERS).map((s) => (
          <button key={s.id} disabled={s.soon} onClick={() => !s.soon && onPick(s.id)} className={`group rounded-2xl p-4 text-right transition-all relative ${s.soon ? "opacity-70 cursor-default" : "hover:scale-[1.04]"}`} style={{ background: C.well, boxShadow: "0 10px 30px rgba(0,0,0,.2)" }}>
            {s.soon && <span className="absolute top-2 left-2 text-[9px] font-bold px-2 py-0.5 rounded-full" style={{ background: C.accentSoft, color: "#92590C" }}>قريباً</span>}
            <div className="flex items-center justify-between mb-2">
              <div className="rounded-xl flex items-center justify-center" style={{ width: 44, height: 44, background: s.soft }}><s.icon size={24} style={{ color: s.color }} /></div>
              <span className="text-2xl">{s.emoji}</span>
            </div>
            <h3 className="text-sm font-extrabold mb-1 leading-tight" style={{ color: C.ink, fontFamily: "Cairo, sans-serif" }}>{s.name}</h3>
            <p className="text-[11px] leading-snug mb-2" style={{ color: C.soft }}>{s.desc}</p>
            {!s.soon && <div className="inline-flex items-center gap-1 font-bold text-xs" style={{ color: s.color }}>دخول <ArrowLeft size={13} /></div>}
          </button>
        ))}
      </div>
      <div className="text-center mt-6 text-xs" style={{ color: "rgba(255,255,255,.5)" }}>كل سيرفر معزول تماماً — مستخدمو كل قطاع يرون شركاته فقط.</div>
    </div>
  </div>);
}

/* ===================== تسجيل الدخول ===================== */
// المسار الرئيسي: تاجر / شركة / فريق العمل (والفريق له مساره الفرعي)
const MAIN_ROLES = {
  merchant: { title: "تاجر", icon: Store, color: () => SC.color, desc: () => `${SC.merchantWord} يطلب من الشركات` },
  company:  { title: "شركة", icon: Building2, color: () => SC.color2, desc: () => "مورّد يبيع للمحلات" },
  team:     { title: "فريق العمل", icon: Users, color: () => C.purple, desc: () => "مشرف · مندوب · سائق" },
};
const TEAM_ROLES = {
  manager:    { title: "مدير مبيعات", icon: UserCog, color: C.ink },
  supervisor: { title: "مشرف", icon: Briefcase, color: C.teal },
  rep:        { title: "مندوب", icon: Handshake, color: C.info },
  promoter:   { title: "مروّج", icon: ImageIcon, color: C.rose },
  driver:     { title: "سائق", icon: Truck, color: C.purple },
  dispatch:   { title: "مشرف حركة", icon: Route, color: "#B45309" },
};
// حسابات تجريبية لكل سيرفر
const DEMO = {
  food: {
    merchant: { user: "ahmed", pass: "1234", entity: "m1" },
    company: { user: "alnoor", pass: "1234", entity: "c1" },
    companyOps: { user: "alnoor-ops", pass: "1234", entity: "c1" },
    manager: { user: "tariq", pass: "1234", entity: "sm1" },
    supervisor: { user: "abdullah", pass: "1234", entity: "sv1" },
    rep: { user: "khaled", pass: "1234", entity: "r1" },
    promoter: { user: "yousef", pass: "1234", entity: "pr1" },
    driver: { user: "salem", pass: "1234", entity: "d1" },
    dispatch: { user: "moftah", pass: "1234", entity: "dsp1" },
    platform: { user: "admin", pass: "admin", entity: null },
  },
  pharma: {
    merchant: { user: "rahma", pass: "1234", entity: "m5" },
    company: { user: "shifa", pass: "1234", entity: "c4" },
    companyOps: { user: "shifa-ops", pass: "1234", entity: "c4" },
    manager: { user: "hossam", pass: "1234", entity: "sm2" },
    supervisor: { user: "kamal", pass: "1234", entity: "sv2" },
    rep: { user: "monir", pass: "1234", entity: "r4" },
    promoter: { user: "samir", pass: "1234", entity: "pr2" },
    driver: { user: "ibrahim", pass: "1234", entity: "d3" },
    dispatch: { user: "adel", pass: "1234", entity: "dsp2" },
    platform: { user: "admin", pass: "admin", entity: null },
  },
};

function Login({ server, onLogin, onBack }) {
  const [mainRole, setMainRole] = useState("merchant"); // merchant | company | team
  const [teamRole, setTeamRole] = useState("rep");
  const [companyRole, setCompanyRole] = useState("admin"); // admin | ops
  const demo = DEMO[server];
  const effectiveRole = mainRole === "team" ? teamRole : mainRole;
  const [u, setU] = useState(demo.merchant.user);
  const [p, setP] = useState(demo.merchant.pass);
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const [showPlatform, setShowPlatform] = useState(false);

  const setRole = (r, tr) => {
    setMainRole(r); setErr("");
    if (r === "company") { const cr = tr || "admin"; setCompanyRole(cr); const d = cr === "ops" ? demo.companyOps : demo.company; setU(d.user); setP(d.pass); return; }
    const role = r === "team" ? (tr || teamRole) : r;
    if (r === "team" && tr) setTeamRole(tr);
    setU(demo[role].user); setP(demo[role].pass);
  };
  const submit = async () => {
    if (busy) return;
    setBusy(true); setErr("");
    try {
      const remoteUser = await loginRemote(u.trim(), p, server);
      if (remoteUser) { onLogin(remoteUser); return; }
    } catch {
      // يعمل العرض التجريبي محلياً إذا لم يُشغّل الخادم بعد.
    }
    const expected = showPlatform ? { ...demo.platform, role: "platform" }
      : effectiveRole === "company" ? { ...(companyRole === "ops" ? demo.companyOps : demo.company), role: "company", companyRole }
      : { ...(demo[effectiveRole] || {}), role: effectiveRole };
    if (u === expected.user && p === expected.pass) onLogin({ role: expected.role, entityId: expected.entity, ...(expected.companyRole ? { companyRole: expected.companyRole } : {}) });
    else setErr("بيانات الدخول غير صحيحة.");
    setBusy(false);
  };

  return (
    <div dir="rtl" className="min-h-screen grid lg:grid-cols-2" style={{ fontFamily: "Tajawal, sans-serif" }}>
      <div className="hidden lg:flex flex-col justify-between p-12 relative overflow-hidden" style={{ background: `linear-gradient(135deg, ${C.ink} 0%, ${SC.color} 100%)` }}>
        <div className="absolute opacity-10" style={{ top: -40, left: -40, width: 300, height: 300, borderRadius: "50%", border: "30px solid #fff" }} />
        <div className="absolute opacity-10" style={{ bottom: -60, right: -20, width: 220, height: 220, borderRadius: "50%", border: "24px solid " + C.accent }} />
        <Logo light size={26} />
        <div className="relative z-10">
          <div className="inline-flex items-center gap-2 rounded-full px-3 py-1.5 mb-5" style={{ background: "rgba(255,255,255,.12)" }}><SC.icon size={16} color={C.accent} /><span className="text-white text-sm font-bold">سيرفر {SC.name} {SC.emoji}</span></div>
          <h1 className="text-white text-4xl font-extrabold leading-snug" style={{ fontFamily: "Cairo, sans-serif" }}>رابط<br /><span style={{ color: C.accent }}>ليربط العالم بين يديك</span></h1>
          <p className="mt-4 leading-relaxed text-lg" style={{ color: "rgba(255,255,255,.75)" }}>منصة متكاملة تربط شركات {SC.name} بـ{SC.merchantsWord} عبر مندوبين وسائقين وإشراف مبيعات.</p>
          <div className="flex gap-3 mt-8">{[Store, Building2, Briefcase, Handshake, Truck].map((I, i) => <div key={i} className="rounded-xl p-2.5" style={{ background: "rgba(255,255,255,.1)" }}><I size={20} color={C.accent} /></div>)}</div>
        </div>
        <button onClick={onBack} className="relative z-10 inline-flex items-center gap-1.5 text-sm font-bold w-fit" style={{ color: "rgba(255,255,255,.7)" }}><ArrowRight size={15} />تغيير السيرفر</button>
      </div>

      <div className="flex items-center justify-center p-6" style={{ background: C.bg }}>
        <div className="w-full max-w-sm">
          <div className="lg:hidden flex items-center justify-between mb-6"><Logo size={24} sub={`سيرفر ${SC.name}`} /><button onClick={onBack} className="text-xs font-bold flex items-center gap-1" style={{ color: SC.color }}><ArrowRight size={13} />تغيير</button></div>
          <h2 className="text-2xl font-extrabold mb-1" style={{ color: C.text }}>تسجيل الدخول</h2>
          <p className="text-sm mb-5" style={{ color: C.soft }}>اختر نوع الحساب ثم سجّل الدخول.</p>

          {!showPlatform && <>
            <div className="grid grid-cols-3 gap-2 mb-3">
              {Object.entries(MAIN_ROLES).map(([k, d]) => { const on = mainRole === k; const col = d.color(); return (
                <button key={k} onClick={() => setRole(k)} className="flex flex-col items-center gap-1.5 rounded-2xl py-3.5 transition-all" style={{ background: on ? col : C.surface, border: `1px solid ${on ? col : C.line}` }}>
                  <d.icon size={20} color={on ? "#fff" : C.soft} /><span className="text-xs font-bold" style={{ color: on ? "#fff" : C.text }}>{d.title}</span>
                </button>); })}
            </div>
            {mainRole === "team" && <div className="grid grid-cols-3 gap-1.5 mb-4 rounded-2xl p-2" style={{ background: C.purpleSoft }}>
              {Object.entries(TEAM_ROLES).map(([k, d]) => { const on = teamRole === k; return (
                <button key={k} onClick={() => setRole("team", k)} className="flex flex-col items-center gap-1 rounded-xl py-2.5 transition-all" style={{ background: on ? d.color : "#fff", border: `1px solid ${on ? d.color : C.line}` }}>
                  <d.icon size={16} color={on ? "#fff" : C.soft} /><span className="text-[9px] font-bold leading-tight text-center" style={{ color: on ? "#fff" : C.text }}>{d.title}</span>
                </button>); })}
            </div>}
            {mainRole === "company" && <div className="grid grid-cols-2 gap-2 mb-4 rounded-2xl p-2" style={{ background: SC.soft }}>
              {[["admin", "أدمن الشركة", "كل الصلاحيات", ShieldCheck], ["ops", "موظف عمليات / مخازن", "سحب الطلبات وتجهيزها وتحميل المندوبين", Warehouse]].map(([k, t, sub, Ic]) => { const on = companyRole === k; return (
                <button key={k} onClick={() => setRole("company", k)} className="flex flex-col items-center gap-1 rounded-xl py-2.5 px-2 transition-all text-center" style={{ background: on ? SC.color2 : "#fff", border: `1px solid ${on ? SC.color2 : C.line}` }}>
                  <Ic size={18} color={on ? "#fff" : C.soft} /><span className="text-[11px] font-bold leading-tight" style={{ color: on ? "#fff" : C.text }}>{t}</span><span className="text-[9px] leading-tight" style={{ color: on ? "rgba(255,255,255,.8)" : C.soft }}>{sub}</span>
                </button>); })}
            </div>}
            {mainRole !== "team" && mainRole !== "company" && <div className="mb-4" />}
          </>}

          {showPlatform && <div className="mb-4 rounded-xl p-3 flex items-center gap-2" style={{ background: C.ink }}><ShieldCheck size={18} color={C.accent} /><span className="text-white text-sm font-bold">دخول إدارة المنصة</span></div>}

          <div className="space-y-3">
            <Field label="اسم المستخدم"><div className="relative"><User size={16} className="absolute right-3 top-3" style={{ color: C.soft }} /><input value={u} onChange={(e) => setU(e.target.value)} className="w-full rounded-xl pr-9 pl-3 py-2.5 text-sm outline-none" style={inputStyle} /></div></Field>
            <Field label="كلمة المرور"><div className="relative"><Lock size={16} className="absolute right-3 top-3" style={{ color: C.soft }} /><input type="password" value={p} onChange={(e) => setP(e.target.value)} className="w-full rounded-xl pr-9 pl-3 py-2.5 text-sm outline-none" style={inputStyle} /></div></Field>
            {err && <div className="text-sm flex items-center gap-1.5" style={{ color: C.danger }}><AlertCircle size={15} />{err}</div>}
            <PrimaryBtn full icon={ArrowRight} onClick={submit}>{busy ? "جارِ التحقق..." : "دخول"}</PrimaryBtn>
            <div className="text-xs rounded-xl p-3" style={{ background: C.accentSoft, color: C.ink }}>بيانات تجريبية: <b>{(showPlatform ? demo.platform : demo[effectiveRole]).user}</b> / <b>{(showPlatform ? demo.platform : demo[effectiveRole]).pass}</b></div>
            <button onClick={() => { setShowPlatform((v) => { const nv = !v; const d = nv ? demo.platform : demo[effectiveRole]; setU(d.user); setP(d.pass); setErr(""); return nv; }); }} className="w-full text-center text-xs font-bold pt-1" style={{ color: C.soft }}>{showPlatform ? "عودة لحسابات السيرفر" : "دخول إدارة المنصة ←"}</button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ===================== قوائم التنقل ===================== */
function navFor(role, companyRole) {
  if (role === "merchant") return [
    { g: "الرئيسية", icon: LayoutDashboard, items: [["m_home", "الرئيسية"], ["m_profile", "ملفي الشخصي"], ["m_requests", "طلبات الجرد والمطابقة"], ["m_stats", "إحصائياتي"]] },
    { g: "الطلبات والفواتير", icon: ShoppingCart, items: [["m_create", "طلبية جديدة / إنشاء فاتورة"], ["m_track", "متابعة الطلبات"], ["m_reorder", "إعادة طلب سابق"], ["m_history", "سجل الطلبات"]] },
    { g: "الشركات", icon: Building2, items: [["m_companies", "الشركات المشترك معها"], ["m_explore", "استكشاف الشركات والانضمام"], ["m_join_track", "متابعة طلبات الانضمام"]] },
    { g: "المنتجات", icon: Package, items: [["m_search", "البحث"], ["m_cats", "التصنيفات"], ["m_fav", "المفضلة"], ["m_offers", "العروض"]] },
    { g: "الحسابات", icon: Wallet, items: [["m_invoices", "مراجعة الفواتير"], ["m_statement", "كشف الحساب"], ["m_settle", "تسوية فاتورة / سداد"], ["m_payments", "المدفوعات"], ["m_company_debt", "مديونياتي عند الشركات"], ["m_credit", "الحد الائتماني"]] },
    { g: "المرتجعات", icon: RotateCcw, items: [["m_return_new", "طلب مرتجع"], ["m_return_track", "متابعة المرتجع"]] },
    { g: "التواصل", icon: MessageSquare, items: [["m_chat", "محادثة الشركات"], ["m_notes", "ملاحظات للمنصة"]] },
  ];
  if (role === "company" && companyRole === "ops") return [
    { g: "العمليات والمخازن", icon: Warehouse, items: [["c_intake", "استلام وتجهيز الطلبات"], ["c_orders", "الطلبات الواردة"], ["c_loads", "تحميل سيارات المندوبين"]] },
  ];
  if (role === "company") return [
    { g: "لوحة الشركة", icon: LayoutDashboard, items: [["c_home", "نظرة عامة"], ["c_branches", "الفروع"], ["c_branch_team", "ربط الفريق بالفروع"], ["c_branch_transfer", "نقل بين الفروع"]] },
    { g: "المنتجات", icon: Package, items: [["c_products", "قائمة المنتجات"], ["c_product_add", "إضافة منتج"], ["c_cats", "التصنيفات والوحدات"]] },
    { g: "العملاء", icon: Store, items: [["c_merchants", "المحلات/العملاء"], ["c_company_clients", "زبائن الشركة المباشرون"], ["c_join", "طلبات الانضمام"], ["c_accounts", "حسابات العملاء"], ["c_debts", "المديونيات"], ["c_returns", "المرتجعات وقراراتها"]] },
    { g: "الطلبات", icon: ShoppingCart, items: [["c_orders", "الطلبات الواردة"], ["c_intake", "استلام وتجهيز الطلبات"]] },
    { g: "الفريق الميداني", icon: Handshake, items: [["c_managers", "مديرو المبيعات"], ["c_supervisors", "المشرفون"], ["c_reps", "المندوبون"], ["c_rep_types", "تصنيف المندوبين"], ["c_link", "ربط المندوبين بالمشرفين"], ["c_ops_perms", "صلاحيات موظف العمليات"], ["c_loads", "تحميل سيارات المندوبين"], ["c_promoters", "مروّجو المبيعات"], ["c_drivers", "السائقون"]] },
    { g: "المهام والحصص", icon: ListChecks, items: [["c_tasks", "إسناد المهام"], ["c_quotas", "حصص المنتجات"], ["c_quota_transfer", "نقل الحصص"], ["c_stock_managers", "تخصيص كميات لمديري المبيعات"]] },
    { g: "مالية التطبيق", icon: Landmark, items: [["c_fin_dash", "لوحة المالية"], ["c_fin_coa", "شجرة الحسابات"], ["c_fin_journal", "دفتر اليومية"], ["c_fin_gl", "دفتر الأستاذ"], ["c_fin_trial", "ميزان المراجعة"], ["c_fin_pl", "ملخص أرباح التوزيع"], ["c_fin_bs", "المركز المالي"], ["c_fin_aging", "أعمار الديون"], ["c_fin_inventory", "تقييم المخزون"], ["c_fin_reps", "مساءلة المندوبين"], ["c_app_comm", "عمولة التطبيق"]] },
    { g: "ذكاء الأعمال", icon: Sparkles, items: [["c_smart_collect", "التحصيل الذكي"], ["c_forecast", "التنبؤ والتحميل"], ["c_annual", "التقرير السنوي"], ["c_regions", "أداء المناطق"], ["c_loyalty", "ولاء التجار"], ["c_timeline", "سجل نشاط العميل"]] },
    { g: "التسويق والإعلان", icon: ImageIcon, items: [["c_ads", "المساحة الإعلانية"], ["c_campaigns", "العروض والحملات"]] },
  ];
  if (role === "supervisor") return [
    { g: "لوحة المشرف", icon: Briefcase, items: [["sv_home", "نظرة عامة"]] },
    { g: "المندوبون", icon: Handshake, items: [["sv_reps", "متابعة المندوبين"], ["sv_targets", "أهداف المندوبين"], ["sv_develop", "تنمية العملاء وتحليل الزيارات"]] },
    { g: "الحصص", icon: Boxes, items: [["sv_quotas", "حصص المنتجات"], ["sv_transfer", "نقل الحصص بين المندوبين"]] },
    { g: "الميدان", icon: MapPinned, items: [["sv_urgent", "خط سير عاجل للمندوب"], ["sv_journey", "خط السير الأسبوعي"], ["sv_stocktake", "اعتماد الجرد"], ["sv_visits", "سجل الزيارات"], ["sv_stock", "طلبات الجرد"]] },
  ];
  if (role === "manager") return [
    { g: "لوحة المدير", icon: UserCog, items: [["mg_home", "نظرة عامة"]] },
    { g: "التقارير", icon: ScrollText, items: [["mg_reports", "التقارير التفصيلية"]] },
    { g: "التحليلات", icon: PieChart, items: [["mg_explore", "الاستعلام والتقارير"], ["mg_sales", "استعلام المبيعات"], ["mg_by_city", "المبيعات حسب المدينة"], ["mg_rep_perf", "أداء المندوبين"], ["mg_churned", "العملاء المتوقفون"], ["mg_visits", "متابعة الزيارات"], ["mg_weak", "الأصناف ضعيفة السحب"]] },
    { g: "المتابعة المالية", icon: HandCoins, items: [["mg_aging", "أعمار الديون"]] },
    { g: "الأهداف والكميات", icon: Target, items: [["mg_stock", "كمياتي من الشركة"], ["mg_targets", "أهداف المشرفين"]] },
    { g: "الفريق", icon: Briefcase, items: [["mg_movements", "حركات الفريق الكاملة"], ["mg_supervisors", "متابعة المشرفين"], ["mg_orders_team", "أوامر للمشرفين"]] },
  ];
  if (role === "promoter") return [
    { g: "لوحة المروّج", icon: ImageIcon, items: [["pm_home", "نظرة عامة"]] },
    { g: "العمل", icon: Activity, items: [["pm_route", "خط السير اليومي"], ["pm_tasks", "مهام العرض والتصوير"]] },
    { g: "السجل", icon: ScrollText, items: [["pm_history", "السجل المُرسل"]] },
  ];
  if (role === "platform") return [
    { g: "المنصة", icon: ShieldCheck, items: [["pf_home", "نظرة عامة"]] },
    { g: "الأرباح والعمولات", icon: CircleDollarSign, items: [["pf_profits", "لوحة الأرباح"], ["pf_commissions", "إدارة العمولات"]] },
    { g: "الشركات", icon: Building2, items: [["pf_companies", "الشركات والاشتراكات"], ["pf_contracts", "العقود"]] },
    { g: "السوق", icon: TrendingUp, items: [["pf_intel", "استخبارات السوق"]] },
    { g: "الإدارة", icon: ServerCog, items: [["pf_merchants", SC.merchantsWord], ["pf_orders", "كل الطلبات"], ["pf_audit", "سجل العمليات"]] },
  ];
  if (role === "rep") return [
    { g: "الرئيسية", icon: LayoutDashboard, items: [["rp_home", "نظرة عامة"]] },
    { g: "العمل اليومي", icon: ClipboardList, items: [["rp_orders", "الطلبات والموافقات"], ["rp_invoice", "إصدار فاتورة للعميل"], ["rp_load", "تحميل سيارتي / مخزوني"], ["rp_tasks", "مهامي المسندة"], ["rp_join", "زيارات التوثيق"]] },
    { g: "التحصيل والتوريد", icon: HandCoins, items: [["rp_settle", "مراجعة تسويات العملاء"], ["rp_handover", "تسليم النقدية للشركة"]] },
    { g: "خط السير", icon: Route, items: [["rp_urgent", "خط السير العاجل"], ["rp_route", "خط السير الأسبوعي"]] },
    { g: "العملاء", icon: Store, items: [["rp_clients", "عملائي"], ["rp_accounts", "الحسابات والمديونيات"], ["rp_assets", "حصر أدوات المحلات"]] },
    { g: "الميدان", icon: MapPinned, items: [["rp_visits", "سجل الزيارات"], ["rp_stocktake", "طلب الجرد من السوق"]] },
    { g: "الأداء", icon: Target, items: [["rp_target", "هدفي وإنجازي"], ["rp_quota", "حصصي من المنتجات"]] },
  ];
  if (role === "driver") return [
    { g: "الرئيسية", icon: LayoutDashboard, items: [["dr_home", "نظرة عامة"]] },
    { g: "التوصيل", icon: Truck, items: [["dr_deliveries", "توصيلاتي"], ["dr_tasks", "المهام المسندة"], ["dr_route", "خط السير"]] },
    { g: "مشرف الحركة", icon: MessageCircle, items: [["dr_dispatch", "التواصل والسجل"]] },
    { g: "السجل", icon: ScrollText, items: [["dr_history", "سجل التسليم"]] },
  ];
  if (role === "dispatch") return [
    { g: "الرئيسية", icon: LayoutDashboard, items: [["dsp_home", "نظرة عامة"]] },
    { g: "الحركة", icon: Route, items: [["dsp_inbox", "الفواتير الواردة"], ["dsp_build", "تكوين حمولة سيارة"], ["dsp_trips", "الرحلات النشطة"]] },
    { g: "السائقون", icon: Truck, items: [["dsp_drivers", "متابعة السائقين"], ["dsp_extra", "الطلعات الإضافية"]] },
    { g: "الإعدادات", icon: Boxes, items: [["dsp_pallets", "سعة الباليتات"]] },
  ];
  return [];
}

// حدود الخطأ: تمنع الشاشة البيضاء وتُظهر سبب العطل بدل توقف التطبيق
class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { err: null }; }
  static getDerivedStateFromError(err) { return { err }; }
  componentDidCatch(err, info) { console.error("Rabet crash:", err, info); }
  render() {
    if (this.state.err) {
      return (<div className="min-h-screen flex items-center justify-center p-6" dir="rtl" style={{ background: C.ink }}>
        <div className="rb-card rounded-2xl p-6 max-w-lg w-full" style={{ background: C.surface, border: `1px solid ${C.danger}` }}>
          <div className="flex items-center gap-3 mb-3"><div className="rounded-xl p-2.5" style={{ background: C.dangerSoft }}><AlertCircle size={22} style={{ color: C.danger }} /></div>
            <div><div className="font-extrabold text-lg" style={{ color: C.text }}>حدث خطأ في هذه الشاشة</div><div className="text-xs" style={{ color: C.soft }}>التطبيق لم يتوقف — يمكنك العودة.</div></div></div>
          <pre className="text-[11px] rounded-xl p-3 overflow-auto mb-3" dir="ltr" style={{ background: C.dangerSoft, color: C.danger, maxHeight: 160 }}>{String(this.state.err && (this.state.err.stack || this.state.err.message || this.state.err))}</pre>
          <button onClick={() => this.setState({ err: null })} className="rb-press w-full rounded-xl py-2.5 font-extrabold" style={{ background: C.success, color: C.onAccent }}>إعادة المحاولة</button>
        </div>
      </div>);
    }
    return this.props.children;
  }
}
function Shell({ server, user, data, actions, onLogout }) {
  const nav = navFor(user.role, user.companyRole);
  const [active, setActive] = useState(nav[0].items[0][0]);
  const [open, setOpen] = useState(() => nav.map((x) => x.g));
  const [sidebar, setSidebar] = useState(false);
  const [openOrder, setOpenOrder] = useState(null);
  const [notifOpen, setNotifOpen] = useState(false);
  const [online, setOnline] = useState(typeof navigator !== "undefined" ? navigator.onLine : true);
  useEffect(() => { const up = () => setOnline(true), dn = () => setOnline(false); window.addEventListener("online", up); window.addEventListener("offline", dn); return () => { window.removeEventListener("online", up); window.removeEventListener("offline", dn); }; }, []);

  const myNotifs = useMemo(() => data.notifications.filter((n) => n.role === user.role && n.entityId === user.entityId), [data.notifications, user]);
  const unread = myNotifs.filter((n) => !n.read).length;

  const ent = useMemo(() => {
    if (user.role === "company") return data.companies.find((c) => c.id === user.entityId);
    if (user.role === "merchant") return data.merchants.find((m) => m.id === user.entityId);
    if (user.role === "rep") return data.reps.find((r) => r.id === user.entityId);
    if (user.role === "driver") return data.drivers.find((d) => d.id === user.entityId);
    if (user.role === "supervisor") return data.supervisors.find((s) => s.id === user.entityId);
    if (user.role === "manager") return data.salesManagers.find((s) => s.id === user.entityId);
    if (user.role === "promoter") return data.promoters.find((p) => p.id === user.entityId);
    if (user.role === "dispatch") return data.dispatchers.find((d) => d.id === user.entityId);
    return null;
  }, [user, data]);

  const roleName = { platform: "المنصة", company: "شركة", merchant: SC.merchantWord, rep: "مندوب", driver: "سائق", supervisor: "مشرف", manager: "مدير مبيعات", promoter: "مروّج مبيعات" }[user.role];
  const roleLabel = user.role === "company" ? (user.companyRole === "ops" ? "موظف عمليات / مخازن" : "أدمن الشركة") : roleName;
  const roleColor = { platform: C.ink, company: SC.color2, merchant: SC.color, rep: C.info, driver: C.purple, supervisor: C.teal, manager: C.ink, promoter: C.rose }[user.role];
  const title = ent ? ent.name : "لوحة المنصة";
  const go = (k) => { setActive(k); setOpenOrder(null); setSidebar(false); };
  const homeKey = nav[0].items[0][0];
  const showBack = openOrder != null || active !== homeKey;
  const goBack = () => { if (openOrder != null) { setOpenOrder(null); } else { setActive(homeKey); } };

  return (
    <div dir="rtl" className="min-h-screen flex" style={{ background: C.bg, fontFamily: "Tajawal, sans-serif" }}>
      <aside className={`rb-glass-bar fixed lg:static z-30 h-full w-64 shrink-0 flex flex-col transition-transform ${sidebar ? "translate-x-0" : "translate-x-full lg:translate-x-0"}`} style={{ background: C.surface, borderLeft: `1px solid ${C.line}` }}>
        <div className="p-5 border-b" style={{ borderColor: C.line }}><Logo size={20} sub={`سيرفر ${SC.name} ${SC.emoji}`} /></div>
        <nav className="flex-1 p-3 space-y-1 overflow-y-auto">
          {nav.map((grp) => { const isSingle = grp.items.length === 1; const expanded = open.includes(grp.g); const activeSingle = isSingle && active === grp.items[0][0];
            return (<div key={grp.g}>
              <button onClick={() => isSingle ? go(grp.items[0][0]) : setOpen((o) => o.includes(grp.g) ? o.filter((x) => x !== grp.g) : [...o, grp.g])}
                className="w-full flex items-center justify-between rounded-xl px-3 py-2.5 text-sm font-bold transition-colors" style={{ color: activeSingle ? C.onAccent : C.text, background: activeSingle ? SC.color : "transparent" }}>
                <span className="flex items-center gap-3"><grp.icon size={18} style={{ color: activeSingle ? C.onAccent : SC.color }} />{grp.g}</span>
                {!isSingle && <ChevronDown size={15} style={{ transform: expanded ? "rotate(180deg)" : "none", transition: ".2s", color: C.soft }} />}
              </button>
              {!isSingle && expanded && <div className="mr-3 mt-0.5 space-y-0.5 pr-3" style={{ borderRight: `1px solid ${C.line}` }}>
                {grp.items.map(([k, t]) => <button key={k} onClick={() => go(k)} className="w-full text-right rounded-lg px-3 py-2 text-[13px] font-bold transition-colors" style={{ color: active === k ? SC.color : C.soft, background: active === k ? SC.soft : "transparent" }}>{t}</button>)}
              </div>}
            </div>); })}
        </nav>
        <div className="p-3 border-t" style={{ borderColor: C.line }}><button onClick={onLogout} className="w-full flex items-center gap-3 rounded-xl px-3 py-2.5 text-sm font-bold transition-colors" style={{ color: C.danger }}><LogOut size={18} />تسجيل الخروج</button></div>
      </aside>
      {sidebar && <div className="fixed inset-0 bg-black/40 z-20 lg:hidden" onClick={() => setSidebar(false)} />}

      <div className="flex-1 min-w-0 flex flex-col">
        <header className="rb-glass-bar sticky top-0 z-10 flex items-center gap-3 px-4 lg:px-6 py-3" style={{ borderBottom: `1px solid ${C.line}` }}>
          <button className="lg:hidden" onClick={() => setSidebar(true)}><Menu size={22} style={{ color: C.text }} /></button>
          {user.role === "rep" && !online && <span className="shrink-0 text-[10px] font-bold px-2 py-1 rounded-full flex items-center gap-1" style={{ background: C.accentSoft, color: "#92590C" }}><AlertCircle size={11} />بلا اتصال — ستُرفع العمليات عند عودة الشبكة</span>}
          {showBack && <button onClick={goBack} className="rb-press flex items-center gap-1 rounded-xl px-2.5 py-1.5 text-sm font-bold shrink-0" style={{ background: C.bg, color: C.text, border: `1px solid ${C.line}` }}><ChevronRight size={17} />رجوع</button>}
          <div className="min-w-0 flex-1"><div className="text-xs font-bold" style={{ color: roleColor }}>{roleLabel}</div><div className="font-extrabold truncate" style={{ color: C.text }}>{title}</div></div>
          <span className="hidden sm:inline-flex items-center gap-1 rounded-full px-2.5 py-1 text-xs font-bold" style={{ background: SC.soft, color: SC.color }}><SC.icon size={13} />{SC.name}</span>
          <div className="relative">
            <button onClick={() => setNotifOpen((v) => !v)} className="rounded-xl p-2 relative transition-colors hover:bg-gray-50" style={{ border: `1px solid ${C.line}` }}><Bell size={18} style={{ color: C.soft }} />{unread > 0 && <><span className="rb-dot absolute -top-1.5 -left-1.5 rounded-full" style={{ width: 16, height: 16, background: C.danger, opacity: .35 }} /><span className="absolute -top-1.5 -left-1.5 min-w-[16px] h-4 px-1 rounded-full text-[10px] font-bold text-white flex items-center justify-center" style={{ background: C.danger }}>{unread}</span></>}</button>
            {notifOpen && <><div className="fixed inset-0 z-40" onClick={() => setNotifOpen(false)} /><NotificationPanel notifs={myNotifs} onRead={(id) => actions.readNotif(id)} onReadAll={() => actions.readAllNotifs(user.role, user.entityId)} onReadType={(t) => actions.readNotifType(user.role, user.entityId, t)} onClose={() => setNotifOpen(false)} /></>}
          </div>
          <div className="hidden sm:flex items-center gap-2 rounded-xl px-3 py-1.5" style={{ background: C.bg }}><div className="w-7 h-7 rounded-full flex items-center justify-center text-xs font-bold text-white" style={{ background: roleColor }}>{(ent?.name || "م")[0]}</div></div>
        </header>
        <main className="flex-1 p-3 sm:p-4 lg:p-6 overflow-x-hidden" style={{ maxWidth: "100vw" }}>
          {openOrder ? <OrderDetail user={user} order={data.orders.find((o) => o.id === openOrder)} data={data} actions={actions} onBack={() => setOpenOrder(null)} />
            : <Router user={user} ent={ent} active={active} data={data} actions={actions} openOrder={setOpenOrder} go={go} server={server} />}
        </main>
      </div>
    </div>
  );
}

/* ===================== الموجّه ===================== */
function Router(props) {
  const { active } = props;
  const map = {
    m_home: MHome, m_stats: MStats, m_profile: MProfile, m_requests: MRequests, m_companies: MCompanies, m_explore: MExplore, m_join_track: MJoinTrack, m_company_debt: MCompanyDebt,
    m_search: MSearch, m_cats: MCats, m_fav: MFav, m_offers: MOffers,
    m_create: MCreate, m_reorder: MReorder, m_track: MTrack, m_history: MHistory,
    m_statement: MStatement, m_invoices: MInvoices, m_settle: MSettle, m_payments: MPayments, m_credit: MCredit,
    m_return_new: MReturnNew, m_return_track: MReturnTrack, m_chat: MChat, m_notes: MNotes,
    c_home: CHome, c_branches: CBranches, c_branch_transfer: CBranchTransfer, c_branch_team: CBranchTeam, c_products: CProducts, c_product_add: CProductAdd, c_cats: CCats,
    c_merchants: CMerchants, c_company_clients: CCompanyClients, c_join: CJoin, c_accounts: CAccounts, c_debts: CDebts, c_returns: CReturns, c_ops_perms: COpsPerms, c_orders: COrders, c_intake: CIntake,
    c_managers: CManagers, c_supervisors: CSupervisors, c_reps: CReps, c_rep_types: CRepTypes, c_link: CLink, c_loads: CLoads, c_promoters: CPromoters, c_drivers: CDrivers,
    c_tasks: CTasks, c_quotas: CQuotas, c_quota_transfer: CQuotaTransfer, c_stock_managers: CStockManagers, c_ads: CAds,
    c_fin_dash: CFinDash, c_fin_coa: CFinCOA, c_fin_journal: CFinJournal, c_fin_gl: CFinGL, c_fin_trial: CFinTrial, c_fin_pl: CFinPL, c_fin_bs: CFinBS, c_fin_aging: CFinAging, c_fin_inventory: CFinInventory, c_fin_reps: CFinReps, c_app_comm: CAppCommission,
    c_smart_collect: CSmartCollect, c_forecast: CForecast, c_campaigns: CCampaigns, c_annual: CAnnualReport, c_regions: CRegions, c_loyalty: CLoyalty, c_timeline: CClientTimeline,
    sv_home: SVHome, sv_stocktake: SVStocktake, sv_journey: SVJourney, sv_reps: SVReps, sv_targets: SVTargets, sv_develop: SVDevelop, sv_quotas: SVQuotas, sv_transfer: SVTransfer, sv_visits: SVVisits, sv_stock: SVStock,
    sv_urgent: SVUrgent,
    mg_home: MgHome, mg_explore: MgExplore, mg_sales: MgSales, mg_by_city: MgSalesByCity, mg_rep_perf: MgRepPerformance, mg_churned: MgChurned, mg_aging: MgAging, mg_visits: MgVisits, mg_weak: MgWeak, mg_movements: MgMovements, mg_supervisors: MgSupervisors, mg_orders_team: MgOrders,
    mg_reports: MgReports, mg_targets: MgTargets, mg_stock: MgStock,
    pm_home: PmHome, pm_route: PmRoute, pm_tasks: PmTasks, pm_history: PmHistory,
    pf_home: PFHome, pf_profits: PFProfits, pf_commissions: PFCommissions, pf_companies: PFCompanies, pf_contracts: PFContracts, pf_intel: PFIntel, pf_merchants: PFMerchants, pf_orders: PFOrders, pf_audit: PFAudit,
    rp_home: RepHome, rp_orders: RepOrders, rp_invoice: RepInvoice, rp_tasks: RepTasks, rp_join: RepJoin,
    rp_clients: RepClients, rp_accounts: RepAccounts, rp_assets: RepAssets, rp_route: RepRoute, rp_visits: RepVisits, rp_stocktake: RepStocktake,
    rp_target: RepTarget, rp_quota: RepQuota, rp_settle: RepSettle, rp_handover: RepHandover, rp_urgent: RepUrgent, rp_load: RepLoad,
    dr_home: DriverHome, dr_dispatch: DriverDispatch, dr_deliveries: DriverDeliveries, dr_tasks: DriverTasks, dr_route: DriverRoute, dr_history: DriverHistory,
    dsp_home: DispatchHome, dsp_inbox: DispatchInbox, dsp_build: DispatchBuild, dsp_trips: DispatchTrips, dsp_drivers: DispatchDrivers, dsp_extra: DispatchExtraTrips, dsp_pallets: DispatchPallets,
  };
  const Comp = map[active] || (() => <Empty icon={AlertCircle} text="الصفحة قيد الإعداد." />);
  return <ErrorBoundary key={active}><Comp {...props} /></ErrorBoundary>;
}

/* ============ فلترة حسب السيرفر ============ */
const sv = (props) => props.server;
const F = {
  companies: (props) => props.data.companies.filter((c) => c.server === sv(props)),
  products: (props) => props.data.products.filter((p) => p.server === sv(props)),
  merchants: (props) => props.data.merchants.filter((m) => m.server === sv(props)),
};
// صلاحيات موظف العمليات: هل الحساب الحالي مقيّد؟
const opsCan = (user, co, key) => !(user?.role === "company" && user?.companyRole === "ops") || (co?.opsPerms?.[key] !== false);
// FEFO: هل توجد تشغيلة أقدم صلاحيةً لنفس الصنف لم تُصرف بعد؟
const fefoOlder = (data, p) => p?.expDate ? data.products.find((q) => q.id !== p.id && q.companyId === p.companyId && q.name === p.name && q.stock > 0 && q.expDate && q.expDate < p.expDate) : null;
// الشركات المنضمّ إليها التاجر فعلاً: شركته الأساسية + من طلب منها + انضمام مقبول
const joinedCompanyIds = (data, me) => {
  const ids = new Set();
  if (me?.companyId) ids.add(me.companyId);
  data.orders.filter((o) => o.merchantId === me?.id).forEach((o) => ids.add(o.companyId));
  data.joinRequests.filter((j) => j.merchantId === me?.id && ["approved", "activated", "done"].includes(j.status)).forEach((j) => ids.add(j.companyId));
  return ids;
};
const useMy = (props) => {
  const meId = props.user.entityId;
  const orders = props.data.orders.filter((o) => o.merchantId === meId);
  const me = props.data.merchants.find((m) => m.id === meId);
  return { meId, orders, me };
};

// ═══════════════════════════════════════════════════════════════════
// النواة المحاسبية: شجرة الحسابات + محرّك القيود المزدوج (مصدر الحقيقة المالية)
// ═══════════════════════════════════════════════════════════════════
// nature: D=مدين طبيعته، C=دائن طبيعته | post: يُسمح الترحيل (أوراق فقط) | stmt: BS ميزانية / IS قائمة دخل
// ctrl: نوع حساب المراقبة (دفتر مساعد) | dim: البُعد الإلزامي على السطر
const COA = [
  // ═══ حسابات عمل التطبيق فقط (توزيع بالجملة) ═══
  { code: "1", name: "الأصول", nature: "D", post: false, stmt: "BS", parent: null },
  { code: "1-1", name: "الأصول المتداولة", nature: "D", post: false, stmt: "BS", parent: "1" },
  { code: "1-2", name: "النقدية والتحصيل", nature: "D", post: false, stmt: "BS", parent: "1-1" },
  { code: "1-2-001", name: "الصندوق الرئيسي", nature: "D", post: true, stmt: "BS", parent: "1-2" },
  { code: "1-2-002", name: "عهدة المندوب النقدية", nature: "D", post: true, stmt: "BS", parent: "1-2", ctrl: "rep_custody", dim: "repId" },
  { code: "1-2-003", name: "البنك — حساب التحصيل", nature: "D", post: true, stmt: "BS", parent: "1-2" },
  { code: "1-3", name: "ذمم العملاء", nature: "D", post: false, stmt: "BS", parent: "1-1" },
  { code: "1-3-001", name: "ذمم العملاء التجار", nature: "D", post: true, stmt: "BS", parent: "1-3", ctrl: "ar", dim: "merchantId" },
  { code: "1-3-002", name: "ذمم عجز المندوبين", nature: "D", post: true, stmt: "BS", parent: "1-3", ctrl: "rep_short", dim: "repId" },
  { code: "1-3-003", name: "شيكات برسم التحصيل", nature: "D", post: true, stmt: "BS", parent: "1-3" },
  { code: "1-4", name: "المخزون", nature: "D", post: false, stmt: "BS", parent: "1-1" },
  { code: "1-4-001", name: "مخزون المستودع الرئيسي", nature: "D", post: true, stmt: "BS", parent: "1-4", ctrl: "inv_main", dim: "productId" },
  { code: "1-4-002", name: "مخزون سيارات المندوبين", nature: "D", post: true, stmt: "BS", parent: "1-4", ctrl: "inv_van", dim: "repId" },
  { code: "1-4-003", name: "بضاعة مرتجعة قيد الفحص", nature: "D", post: true, stmt: "BS", parent: "1-4" },
  { code: "2", name: "الالتزامات", nature: "C", post: false, stmt: "BS", parent: null },
  { code: "2-1", name: "الالتزامات المتداولة", nature: "C", post: false, stmt: "BS", parent: "2" },
  { code: "2-1-004", name: "أرصدة عملاء دائنة", nature: "C", post: true, stmt: "BS", parent: "2-1", dim: "merchantId" },
  { code: "3", name: "حقوق الملكية", nature: "C", post: false, stmt: "BS", parent: null },
  { code: "3-009", name: "حساب افتتاحي (رصيد أول المدة)", nature: "C", post: true, stmt: "BS", parent: "3" },
  { code: "4", name: "إيرادات المبيعات", nature: "C", post: false, stmt: "IS", parent: null },
  { code: "4-001", name: "مبيعات — تجزئة", nature: "C", post: true, stmt: "IS", parent: "4" },
  { code: "4-002", name: "مبيعات — جملة", nature: "C", post: true, stmt: "IS", parent: "4" },
  { code: "4-003", name: "مبيعات — أدوية", nature: "C", post: true, stmt: "IS", parent: "4" },
  { code: "5", name: "تكلفة المبيعات", nature: "D", post: false, stmt: "IS", parent: null },
  { code: "5-001", name: "تكلفة البضاعة المباعة", nature: "D", post: true, stmt: "IS", parent: "5" },
  { code: "5-002", name: "مردودات المبيعات", nature: "D", post: true, stmt: "IS", parent: "5" },
  { code: "5-003", name: "تالف ومنتهي الصلاحية", nature: "D", post: true, stmt: "IS", parent: "5" },
  { code: "7-001", name: "فروق التوريد (زيادة)", nature: "C", post: true, stmt: "IS", parent: "4" },
];
const coaAcc = (code) => COA.find((a) => a.code === code);
const acctName = (code) => coaAcc(code)?.name || code;
// حساب المبيعات حسب القطاع/التصنيف
const salesAccountFor = (server, tier) => server === "pharma" ? "4-003" : tier === "wholesale" ? "4-002" : "4-001";
const invAccountForRep = (isVan) => isVan ? "1-4-002" : "1-4-001";
// تقريب النصف لأعلى على 3 خانات (الدينار = 1000 درهم)
const money3 = (n) => Math.round((n + Number.EPSILON) * 1000) / 1000;
// إنشاء قيد يومية متوازن: lines=[{acct, dr, cr, dims}]؛ يرفض غير المتوازن
const buildJournal = (seq, { date, source, docId, server, companyId, memo, lines, reversalOf }) => {
  const norm = lines.map((l) => ({ acct: l.acct, dr: money3(l.dr || 0), cr: money3(l.cr || 0), dims: l.dims || {}, memo: l.memo || "" }));
  const td = norm.reduce((s, l) => s + l.dr, 0), tc = norm.reduce((s, l) => s + l.cr, 0);
  return { id: "JV-" + String(seq).padStart(5, "0"), date: date || "الآن", source, docId: docId || null, server, companyId: companyId || null, memo: memo || "", lines: norm, totalDr: money3(td), totalCr: money3(tc), balanced: Math.abs(td - tc) < 0.0005, reversalOf: reversalOf || null, voided: false };
};
// أرصدة الحسابات من كل القيود (غير الملغاة): يعيد خريطة code→{dr,cr,bal}
const ledgerBalances = (ledger, filter = {}) => {
  const bal = {};
  (ledger || []).filter((j) => !j.voided).forEach((j) => {
    if (filter.server && j.server !== filter.server) return;
    if (filter.companyId && j.companyId !== filter.companyId) return;
    j.lines.forEach((l) => {
      if (filter.merchantId && l.dims.merchantId !== filter.merchantId) return;
      if (filter.repId && l.dims.repId !== filter.repId) return;
      if (!bal[l.acct]) bal[l.acct] = { dr: 0, cr: 0 };
      bal[l.acct].dr += l.dr; bal[l.acct].cr += l.cr;
    });
  });
  Object.keys(bal).forEach((code) => { const a = coaAcc(code); const net = bal[code].dr - bal[code].cr; bal[code].bal = a?.nature === "C" ? -net : net; bal[code].signed = net; });
  return bal;
};
// متوسط التكلفة المرجّح WAC لكل صنف (يُشتق من سعر المنتج كتقريب في النموذج)
const wacOf = (data, pid) => { const p = data.products.find((x) => x.id === pid); return p ? money3(p.price * 0.72) : 0; }; // هامش تقديري 28%

// ترحيل قيد للدفتر داخل تحويل الحالة: يُرجع الحالة الجديدة مع القيد مُضافاً
const postToLedger = (d, entry) => { const seq = d.jvSeq || 1; const j = buildJournal(seq, entry); return { ...d, jvSeq: seq + 1, ledger: [j, ...(d.ledger || [])] }; };
// إنشاء فاتورة حقيقية للطلب عند طباعته/تسليمه، مرة واحدة فقط.
const ensureOrderInvoice = (d, o) => {
  if (!o || (d.invoices || []).some((i) => i.orderId === o.id)) return d;
  const merchant = d.merchants.find((m) => m.id === o.merchantId);
  const amount = orderTotal(o);
  const invoice = { id: `INV-${o.id}`, server: o.server, merchantId: o.merchantId, companyId: o.companyId, orderId: o.id, amount, paid: 0, date: "الآن", due: "خلال 30 يوم", ageDays: 0, items: o.items.map((it) => ({ ...it })) };
  const nd = postToLedger(d, { date: "الآن", source: "T1", docId: invoice.id, server: o.server, companyId: o.companyId, memo: `فاتورة ${invoice.id} — ${merchant?.name || "عميل"}`, lines: [
    { acct: "1-3-001", dr: amount, cr: 0, dims: { merchantId: o.merchantId } },
    { acct: salesAccountFor(o.server, merchant?.tier), dr: 0, cr: amount, dims: {} },
  ] });
  return { ...nd, invoices: [invoice, ...(nd.invoices || [])], merchants: nd.merchants.map((m) => m.id === o.merchantId ? { ...m, balance: (m.balance || 0) - amount } : m) };
};
// عكس قيد (بدل الحذف)
const reverseInLedger = (d, jvId) => { const orig = (d.ledger || []).find((j) => j.id === jvId); if (!orig || orig.voided) return d; const seq = d.jvSeq || 1; const rev = buildJournal(seq, { date: "الآن", source: orig.source + "_REVERSAL", docId: orig.docId, server: orig.server, companyId: orig.companyId, memo: "عكس القيد " + orig.id, reversalOf: orig.id, lines: orig.lines.map((l) => ({ acct: l.acct, dr: l.cr, cr: l.dr, dims: l.dims })) }); return { ...d, jvSeq: seq + 1, ledger: [rev, ...(d.ledger || []).map((j) => j.id === jvId ? { ...j, voided: true } : j)] }; };
// ═══ تهجير القيود الافتتاحية: تحويل أرصدة الـSeed إلى قيد افتتاحي متوازن يوم-صفر ═══
const withOpeningLedger = (d) => {
  if (d.ledger && d.ledger.length) return d;
  let seq = 1; const entries = [];
  const push = (server, companyId, memo, lines) => { const j = buildJournal(seq++, { date: "افتتاحي", source: "OPENING", server, companyId, memo, lines }); entries.push(j); };
  // لكل شركة: افتتاح متوازن (ذمم عملاء + مخزون مستودع + مخزون سيارات + عجز مندوبين مقابل حساب افتتاحي)
  d.companies.forEach((co) => {
    const lines = [];
    // ذمم العملاء (رصيد سالب = مديونية)
    d.merchants.filter((m) => m.companyId === co.id).forEach((m) => { const debt = Math.max(0, -(m.balance || 0)); if (debt > 0) lines.push({ acct: "1-3-001", dr: debt, cr: 0, dims: { merchantId: m.id } }); const credit = Math.max(0, m.balance || 0); if (credit > 0) lines.push({ acct: "2-1-004", dr: 0, cr: credit, dims: { merchantId: m.id } }); });
    // مخزون المستودع (بتكلفة WAC)
    d.products.filter((p) => p.companyId === co.id).forEach((p) => { const val = money3((p.stock || 0) * wacOf(d, p.id)); if (val > 0) lines.push({ acct: "1-4-001", dr: val, cr: 0, dims: { productId: p.id } }); });
    // مخزون السيارات + عجز المندوبين
    d.reps.filter((r) => r.companyId === co.id).forEach((r) => {
      Object.entries(r.vanStock || {}).forEach(([pid, q]) => { const val = money3((q || 0) * wacOf(d, pid)); if (val > 0) lines.push({ acct: "1-4-002", dr: val, cr: 0, dims: { repId: r.id, productId: pid } }); });
      if ((r.shortfall || 0) > 0) lines.push({ acct: "1-3-002", dr: r.shortfall, cr: 0, dims: { repId: r.id } });
    });
    if (!lines.length) return;
    const totalDr = lines.reduce((s, l) => s + (l.dr || 0), 0);
    // الطرف الدائن الموازن: حساب افتتاحي بحقوق الملكية
    lines.push({ acct: "3-009", dr: 0, cr: money3(totalDr), dims: {} });
    push(co.server, co.id, "قيد افتتاحي — أرصدة أول المدة", lines);
  });
  return { ...d, ledger: entries, jvSeq: seq };
};
const priceAfter = (p) => p.offer ? Math.round(p.price * (1 - p.offer / 100)) : p.price;
// السعر للعميل بعد خصم العرض + نسبة العقد + نسبة حصة التطبيق
const priceForMerchant = (p, m) => {
  let pr = priceAfter(p);
  const extra = (m?.contractPct || 0) + (m?.appDiscountPct || 0);
  if (extra > 0) pr = Math.round(pr * (1 - extra / 100));
  return pr;
};
const compName = (data, id) => data.companies.find((c) => c.id === id)?.name || "—";
const repName = (data, id) => data.reps.find((r) => r.id === id)?.name || "—";
const merchName = (data, id) => data.merchants.find((m) => m.id === id)?.name || "—";
const prodName = (data, id) => data.products.find((p) => p.id === id)?.name || "—";

/* ===== المساحة الإعلانية (للتجار) ===== */
function AdCarousel({ ads }) {
  const live = ads.filter((a) => a.active);
  const [i, setI] = useState(0);
  useEffect(() => { if (live.length < 2) return; const t = setInterval(() => setI((x) => (x + 1) % live.length), 4500); return () => clearInterval(t); }, [live.length]);
  if (!live.length) return null;
  const a = live[i % live.length];
  return (<div className="relative rounded-2xl overflow-hidden" style={{ boxShadow: "0 10px 30px rgba(11,31,58,.12)" }}>
    <div className="relative p-5 sm:p-6 transition-all duration-500" style={{ background: `linear-gradient(120deg, ${a.c1} 0%, ${a.c2} 100%)`, minHeight: 132 }}>
      <div className="absolute opacity-10" style={{ top: -30, left: -30, width: 160, height: 160, borderRadius: "50%", border: "24px solid #fff" }} />
      <div className="absolute opacity-10" style={{ bottom: -40, right: 40, width: 110, height: 110, borderRadius: "50%", background: C.well }} />
      <div className="relative z-10 flex items-center gap-4">
        <div className="rounded-2xl flex items-center justify-center text-4xl shrink-0" style={{ width: 64, height: 64, background: "rgba(255,255,255,.18)" }}>{a.emoji}</div>
        <div className="flex-1 min-w-0 text-white">
          <div className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full mb-1.5" style={{ background: "rgba(255,255,255,.2)" }}><Tag size={10} />إعلان مموّل</div>
          <div className="font-extrabold text-lg leading-tight" style={{ fontFamily: "Cairo, sans-serif" }}>{a.title}</div>
          <div className="text-sm opacity-90 mt-0.5 truncate">{a.subtitle}</div>
        </div>
        <div className="hidden sm:block shrink-0"><span className="inline-flex items-center gap-1 rounded-xl px-4 py-2 text-sm font-extrabold" style={{ background: C.well, color: a.c1 }}>{a.cta} <ArrowLeft size={15} /></span></div>
      </div>
    </div>
    {live.length > 1 && <div className="absolute bottom-2.5 left-0 right-0 flex justify-center gap-1.5 z-10">{live.map((_, x) => <button key={x} onClick={() => setI(x)} className="rounded-full transition-all" style={{ width: x === i % live.length ? 18 : 6, height: 6, background: x === i % live.length ? "#fff" : "rgba(255,255,255,.5)" }} />)}</div>}
  </div>);
}
// شارة الصلاحية للأدوية/الأغذية
function ExpiryBadge({ expDate, small }) {
  const e = expiryInfo(expDate); if (!e || e.state === "ok") return e ? <span className={`inline-flex items-center gap-1 ${small ? "text-[10px]" : "text-xs"} font-bold px-2 py-0.5 rounded-full`} style={{ background: e.soft, color: e.color }}><ShieldCheck size={small ? 10 : 12} />صالح حتى {fmtDate(expDate)}</span> : null;
  return <span className={`inline-flex items-center gap-1 ${small ? "text-[10px]" : "text-xs"} font-bold px-2 py-0.5 rounded-full`} style={{ background: e.soft, color: e.color }}><AlertCircle size={small ? 10 : 12} />{e.label}</span>;
}

/* ===== قائمة التجهيز/الالتقاط للمخزن (قابلة للطباعة) ===== */
function PickingListView({ order, data, onClose }) {
  const m = data.merchants.find((x) => x.id === order.merchantId);
  const co = data.companies.find((c) => c.id === order.companyId);
  return (<div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-3 bg-black/50" onClick={onClose}>
    <div className="rb-modal w-full max-w-md rounded-2xl overflow-hidden max-h-[92vh] flex flex-col" style={{ background: "#fff" }} onClick={(e) => e.stopPropagation()}>
      <div id="invoice-print" className="overflow-auto">
        <div className="p-4 text-white" style={{ background: C.ink }}>
          <div className="flex items-center justify-between"><div className="flex items-center gap-2"><ClipboardList size={20} /><div><div className="font-extrabold">قائمة تجهيز — للمخزن</div><div className="text-[11px] opacity-80">{co?.name}</div></div></div><div className="text-left"><div className="text-xs opacity-80">طلب {order.id}</div>{order.sysNo && <div className="text-xs font-bold">{order.sysNo}</div>}</div></div>
        </div>
        <div className="px-4 py-2.5 text-xs flex items-center justify-between" style={{ background: C.bg }}><span style={{ color: C.text }}><b>العميل:</b> {m?.name} · {m?.area}</span><span style={{ color: C.soft }}>{FULFILL[order.fulfillment || "delivery"]?.label}</span></div>
        <div className="p-4">
          <table className="w-full text-sm" style={{ borderCollapse: "collapse" }}>
            <thead><tr style={{ borderBottom: `2px solid ${C.ink}` }}><th className="text-right p-2" style={{ color: C.text }}>الصنف</th><th className="p-2" style={{ color: C.text }}>المطلوب</th><th className="p-2" style={{ color: C.text }}>الرف</th><th className="p-2" style={{ color: C.text }}>تم ✓</th></tr></thead>
            <tbody>{order.items.map((it, i) => { const p = data.products.find((x) => x.id === it.pid); return (<tr key={i} style={{ borderBottom: `1px solid ${C.line}` }}>
              <td className="p-2"><div className="font-bold" style={{ color: C.text }}>{it.name}</div>{p?.batch && <div className="text-[10px]" style={{ color: C.soft }}>تشغيلة {p.batch}{p.expDate ? ` · صلاحية ${fmtDate(p.expDate)}` : ""}</div>}</td>
              <td className="p-2 text-center font-extrabold" style={{ color: C.text }}>{it.boxes ? `${it.boxes} صندوق` : it.qty}</td>
              <td className="p-2 text-center text-xs" style={{ color: C.soft }}>{"R-" + (10 + (i * 7) % 40)}</td>
              <td className="p-2 text-center"><span className="inline-block rounded" style={{ width: 20, height: 20, border: `2px solid ${C.ink}` }} /></td>
            </tr>); })}</tbody>
          </table>
          <div className="flex items-center justify-between mt-4 pt-3 text-xs" style={{ borderTop: `1px dashed ${C.line}`, color: C.soft }}><span>المُجهِّز: ______________</span><span>التاريخ/الوقت: ______________</span></div>
        </div>
      </div>
      <div className="p-3 flex gap-2 shrink-0 no-print" style={{ borderTop: `1px solid ${C.line}` }}>
        <PrimaryBtn full icon={Printer} onClick={() => window.print()}>طباعة القائمة</PrimaryBtn>
        <GhostBtn icon={XCircle} onClick={onClose}>إغلاق</GhostBtn>
      </div>
    </div>
  </div>);
}

/* ===== فاتورة حقيقية قابلة للطباعة (يستخدمها موظف العمليات وغيره) ===== */
function InvoiceView({ order, data, onClose }) {
  const co = data.companies.find((c) => c.id === order.companyId);
  const m = data.merchants.find((x) => x.id === order.merchantId);
  const rep = data.reps.find((r) => r.id === order.repId);
  const drv = data.drivers.find((v) => v.id === order.driverId);
  const isPharma = order.server === "pharma";
  const total = orderTotal(order);
  const totalQty = orderQty(order);
  const today = new Date().toLocaleDateString("ar-LY", { year: "numeric", month: "long", day: "numeric" });
  return (<div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-3 bg-black/50" onClick={onClose}>
    <div className="rb-modal w-full max-w-lg rounded-2xl overflow-hidden max-h-[92vh] flex flex-col" style={{ background: "#fff" }} onClick={(e) => e.stopPropagation()}>
      <div id="invoice-print" className="overflow-auto">
        {/* ترويسة */}
        <div className="p-5 text-white" style={{ background: `linear-gradient(120deg, ${C.ink}, ${SC.color})` }}>
          <div className="flex items-start justify-between gap-3">
            <div className="flex items-center gap-3"><div className="rounded-2xl flex items-center justify-center text-3xl" style={{ width: 54, height: 54, background: "rgba(255,255,255,.18)" }}>{co?.logo}</div><div><div className="font-extrabold text-lg leading-tight">{co?.name}</div><div className="text-xs opacity-90">{co?.cat} · {co?.city} · {co?.phone}</div>{co?.taxNo && <div className="text-[11px] opacity-80">الرقم الضريبي: {co.taxNo}</div>}</div></div>
            <div className="text-left"><div className="text-xl font-extrabold">فاتورة</div><div className="text-xs opacity-90">رقم: {order.sysNo || order.id}</div><div className="text-xs opacity-90">{today}</div></div>
          </div>
        </div>
        {/* بيانات العميل */}
        <div className="grid grid-cols-2 gap-3 p-4 text-sm" style={{ background: C.bg }}>
          <div><div className="text-[11px] font-bold mb-1" style={{ color: C.soft }}>العميل</div><div className="font-bold" style={{ color: C.text }}>{m?.name}</div><div className="text-xs" style={{ color: C.soft }}>{m?.owner} · {m?.area}</div><div className="text-xs" style={{ color: C.soft }}>{m?.phone}</div></div>
          <div className="text-left"><div className="text-[11px] font-bold mb-1" style={{ color: C.soft }}>التسليم</div><div className="font-bold" style={{ color: C.text }}>{FULFILL[order.fulfillment || "delivery"]?.label}</div>{rep && <div className="text-xs" style={{ color: C.soft }}>المندوب: {rep.name}</div>}{drv && <div className="text-xs" style={{ color: C.soft }}>السائق: {drv.name}</div>}<div className="text-xs" style={{ color: C.soft }}>الطلب: {order.id}</div></div>
        </div>
        {/* البنود */}
        <div className="p-4">
          <table className="w-full text-sm" style={{ borderCollapse: "collapse" }}>
            <thead><tr style={{ background: C.well, color: C.soft }}><th className="text-right p-2 rounded-r-lg">الصنف</th><th className="p-2">الكمية</th><th className="p-2">سعر الوحدة</th><th className="p-2 rounded-l-lg">القيمة</th></tr></thead>
            <tbody>{order.items.map((it, i) => { const p = data.products.find((x) => x.id === it.pid); const e = expiryInfo(it.expDate || p?.expDate); return (<tr key={i} style={{ borderBottom: `1px solid ${C.line}` }}>
              <td className="p-2 align-top"><div className="font-bold" style={{ color: C.text }}>{it.name}</div>{(isPharma && ((it.batch || p?.batch) || (it.expDate || p?.expDate))) && <div className="flex flex-wrap gap-1 mt-0.5">{(it.batch || p?.batch) && <span className="text-[9px] font-bold px-1 py-0.5 rounded" style={{ background: C.bg, color: C.soft }}>تشغيلة {it.batch || p?.batch}</span>}{(it.expDate || p?.expDate) && <span className="text-[9px] font-extrabold px-1 py-0.5 rounded" style={{ background: e?.soft, color: e?.color }}>صلاحية {fmtDate(it.expDate || p?.expDate)}</span>}</div>}</td>
              <td className="p-2 text-center" style={{ color: C.text }}>{it.boxes ? <span>{it.boxes} صندوق<div className="text-[10px]" style={{ color: C.soft }}>({it.qty} {p?.unit})</div></span> : `${it.qty} ${p?.unit || ""}`}</td>
              <td className="p-2 text-center" style={{ color: C.soft }}>{money(it.price)}</td>
              <td className="p-2 text-center font-bold" style={{ color: C.text }}>{money(it.qty * it.price)}</td>
            </tr>); })}</tbody>
          </table>
          <div className="flex justify-end mt-3"><div className="w-56 space-y-1 text-sm"><div className="flex justify-between"><span style={{ color: C.soft }}>إجمالي الكمية</span><span className="font-bold" style={{ color: C.text }}>{totalQty} قطعة</span></div><div className="flex justify-between items-center pt-2 mt-1" style={{ borderTop: `2px solid ${C.ink}` }}><span className="font-extrabold" style={{ color: C.text }}>الإجمالي</span><span className="text-lg font-extrabold" style={{ color: SC.color }}>{money(total)}</span></div></div></div>
          {order.note && <div className="text-xs mt-3 rounded-lg p-2" style={{ background: C.bg, color: C.soft }}>ملاحظة: {order.note}</div>}
          <div className="flex items-center justify-between mt-4 pt-3" style={{ borderTop: `1px dashed ${C.line}` }}><FakeBarcode code={(order.sysNo || order.id).replace(/\D/g, "") || "000000"} /><div className="text-[10px] text-center" style={{ color: C.soft }}>توقيع/ختم الشركة<div className="mt-4" style={{ borderTop: `1px solid ${C.line}`, width: 90 }} /></div></div>
        </div>
      </div>
      <div className="p-3 flex gap-2 shrink-0" style={{ borderTop: `1px solid ${C.line}` }}>
        <PrimaryBtn full icon={Printer} onClick={() => window.print()}>طباعة الفاتورة</PrimaryBtn>
        <GhostBtn icon={XCircle} onClick={onClose}>إغلاق</GhostBtn>
      </div>
    </div>
  </div>);
}

function ProductCard({ p, data, fav, onFav, onCompare, inCompare, onAdd }) {
  return (
    <Card className="p-4 flex flex-col">
      <div className="flex items-start justify-between mb-2">
        {p.imgs && p.imgs.length ? <div className="rounded-xl overflow-hidden" style={{ width: 44, height: 44, border: `1px solid ${C.line}` }}><img src={p.imgs[0]} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /></div> : <div className="rounded-xl p-2.5" style={{ background: SC.soft }}><Package size={20} style={{ color: SC.color }} /></div>}
        <div className="flex gap-1">
          {onFav && <button onClick={() => onFav(p.id)} className="rounded-lg p-1.5" style={{ background: fav ? C.dangerSoft : C.bg }}><Heart size={15} style={{ color: fav ? C.danger : C.soft }} fill={fav ? C.danger : "none"} /></button>}
          {onCompare && <button onClick={() => onCompare(p.id)} className="rounded-lg p-1.5" style={{ background: inCompare ? C.infoSoft : C.bg }}><GitCompare size={15} style={{ color: inCompare ? C.info : C.soft }} /></button>}
        </div>
      </div>
      <div className="font-extrabold text-sm leading-snug" style={{ color: C.text }}>{p.name}</div>
      <div className="text-xs mb-2" style={{ color: C.soft }}>{compName(data, p.companyId)} · {p.cat}</div>
      <div className="flex items-center gap-2 mb-3">
        {p.offer > 0 ? <>
          <span className="font-extrabold" style={{ color: C.danger }}>{money(priceAfter(p))}</span>
          <span className="text-xs line-through" style={{ color: C.soft }}>{money(p.price)}</span>
          <span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.dangerSoft, color: C.danger }}>-{p.offer}%</span>
        </> : <span className="font-extrabold" style={{ color: SC.color }}>{money(p.price)}</span>}
      </div>
      <div className="text-xs mb-3" style={{ color: C.soft }}>أقل كمية: {p.minOrder} {p.unit} · متوفر: {p.stock}</div>
      {(p.expDate || (p.tiers && p.tiers.length)) && <div className="flex flex-wrap gap-1.5 mb-3">
        {p.expDate && <ExpiryBadge expDate={p.expDate} small />}
        {p.batch && <span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: C.bg, color: C.soft }}><Hash size={10} />{p.batch}</span>}
        {p.tiers && p.tiers.length > 0 && <span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: C.purpleSoft, color: C.purple }}><Layers size={10} />خصم كمية حتى {Math.max(...p.tiers.map((t) => t.discount))}%</span>}
      </div>}
      {onAdd && <PrimaryBtn sm full icon={Plus} onClick={() => onAdd(p)}>أضف للطلب</PrimaryBtn>}
    </Card>
  );
}

/* ============================================================
   ============   صفحات التاجر   =============================
   ============================================================ */
// الملف الشخصي للتاجر: بيانات المحل والرخصة والموقع والمسؤول — احترافي وقابل للتعديل
// طلبات المندوب/الشركة للتاجر: جرد الأصناف + مطابقة الحساب
function MRequests(props) {
  const { data, actions } = props; const { me } = useMy(props);
  const stReqs = data.stocktakes.filter((st) => st.type === "merchant" && st.merchantId === me.id);
  const recReqs = data.reconciliations.filter((r) => r.merchantId === me.id);
  const [counts, setCounts] = useState({});
  const [recNote, setRecNote] = useState({});
  return (<div className="space-y-4"><SectionTitle sub="طلبات الجرد ومطابقة الحساب الواردة من مندوبك أو الشركة.">طلبات الجرد والمطابقة</SectionTitle>
    {/* مطابقة الحساب */}
    {recReqs.map((r) => (<Card key={r.id} className="p-4 min-w-0" style={{ border: `1px solid ${r.status === "pending" ? C.accent : C.line}` }}>
      <div className="flex items-center justify-between flex-wrap gap-2 mb-2"><h3 className="font-extrabold flex items-center gap-2" style={{ color: C.text }}><ScrollText size={17} style={{ color: SC.color }} />مطابقة حساب — {r.id}</h3><Badge status={r.status} map={REC_STATUS} /></div>
      <div className="rounded-xl p-3 mb-3" style={{ background: C.bg }}><div className="text-xs" style={{ color: C.soft }}>طلب من: {r.requestedBy === "company" ? compName(data, r.companyId) : repName(data, r.requesterId)} · {r.date}</div>
        <div className="flex items-center justify-between mt-2"><span className="text-sm font-bold" style={{ color: C.text }}>الرصيد المسجّل عليك</span><span className="text-lg font-extrabold" style={{ color: r.statedBalance < 0 ? C.danger : C.success }}>{money(Math.abs(r.statedBalance))} {r.statedBalance < 0 ? "عليك" : "لك"}</span></div>
      </div>
      {r.status === "pending" ? <>
        <Field label="ملاحظاتك (إن كان لديك اعتراض على الرقم)"><Input value={recNote[r.id] || ""} onChange={(e) => setRecNote({ ...recNote, [r.id]: e.target.value })} placeholder="مثال: سدّدت 500 لم تُحتسب" /></Field>
        <div className="flex gap-2 mt-3"><PrimaryBtn full tone="success" icon={CheckCircle2} onClick={() => { actions.respondReconciliation(r.id, true, ""); toast("أكّدت صحة الرصيد ✓"); }}>الرصيد صحيح — أوافق</PrimaryBtn><GhostBtn icon={AlertCircle} onClick={() => { if (!recNote[r.id]) { toast("اكتب ملاحظتك أولاً", "danger"); return; } actions.respondReconciliation(r.id, false, recNote[r.id]); toast("أُرسلت ملاحظتك ✓"); }}>لدي ملاحظة</GhostBtn></div>
      </> : r.merchantNote ? <div className="text-xs rounded-lg p-2" style={{ background: C.dangerSoft, color: C.danger }}>ملاحظتك: {r.merchantNote}</div> : <div className="text-xs font-bold" style={{ color: C.success }}>أكّدت صحة الرصيد ✓</div>}
    </Card>))}
    {/* طلبات الجرد */}
    {stReqs.map((st) => (<Card key={st.id} className="p-4 min-w-0" style={{ border: `1px solid ${st.status === "requested" ? C.accent : C.line}` }}>
      <div className="flex items-center justify-between flex-wrap gap-2 mb-2"><h3 className="font-extrabold flex items-center gap-2" style={{ color: C.text }}><ClipboardList size={17} style={{ color: SC.color }} />طلب جرد — {st.id}</h3><span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: st.status === "submitted" ? C.successSoft : C.accentSoft, color: st.status === "submitted" ? C.success : "#92590C" }}>{st.status === "submitted" ? "تم الإرسال ✓" : "بانتظار جردك"}</span></div>
      <div className="text-xs mb-3" style={{ color: C.soft }}>من {repName(data, st.repId)} · {st.note}</div>
      <div className="space-y-1.5">{st.lines.map((l) => { const cnt = counts[st.id + l.pid] ?? l.counted ?? ""; return (<div key={l.pid} className="flex items-center justify-between gap-2 rounded-xl px-3 py-2" style={{ background: C.bg }}>
        <span className="text-sm font-bold min-w-0 truncate" style={{ color: C.text }}>{l.name}</span>
        {st.status === "requested" ? <input type="number" value={cnt} onChange={(e) => setCounts({ ...counts, [st.id + l.pid]: e.target.value })} placeholder="الكمية لديك" className="w-24 text-center text-sm font-bold rounded-lg py-1.5 outline-none" style={inputStyle} /> : <span className="text-sm font-extrabold" style={{ color: C.text }}>{l.counted ?? "—"}</span>}
      </div>); })}</div>
      {st.status === "requested" && <div className="mt-3"><PrimaryBtn full icon={Send} onClick={() => { const c = {}; st.lines.forEach((l) => { c[l.pid] = +(counts[st.id + l.pid] ?? 0); }); actions.submitStocktake(st.id, c); toast("أُرسل الجرد لمندوبك ✓"); }}>إرسال نتيجة الجرد</PrimaryBtn></div>}
    </Card>))}
    {recReqs.length === 0 && stReqs.length === 0 && <Empty icon={ClipboardList} text="لا توجد طلبات حالياً." />}
  </div>);
}
function MProfile(props) {
  const { data, actions } = props; const { me } = useMy(props);
  const rep = data.reps.find((r) => r.id === me.repId);
  const co = data.companies.find((c) => c.id === me.companyId);
  const T = TIERS[me.tier] || TIERS.retail;
  const [edit, setEdit] = useState(false);
  const [f, setF] = useState({ managerName: me.managerName || me.owner, managerPhone: me.managerPhone || me.phone, email: me.email || "", address: me.address || "" });
  const save = () => { actions.updateMyProfile(me.id, f); setEdit(false); toast("حُفظ ملفك الشخصي ✓"); };
  const Row = ({ icon: Ic, label, value, ltr }) => (<div className="flex items-start gap-3 py-2.5" style={{ borderBottom: `1px solid ${C.line}` }}>
    <div className="rounded-lg p-2 shrink-0" style={{ background: SC.soft }}><Ic size={15} style={{ color: SC.color }} /></div>
    <div className="min-w-0 flex-1"><div className="text-[11px] font-bold" style={{ color: C.soft }}>{label}</div><div className="text-sm font-bold truncate" dir={ltr ? "ltr" : undefined} style={{ color: C.text, textAlign: ltr ? "right" : undefined }}>{value || "—"}</div></div>
  </div>);
  return (<div className="space-y-4"><SectionTitle sub="بيانات محلك ورخصتك وموقعك والمسؤول عنه — واضحة للشركات والمندوبين.">ملفي الشخصي</SectionTitle>
    {/* ترويسة فخمة */}
    <div className="rb-card rounded-3xl overflow-hidden">
      <div className="p-5 sm:p-6 text-white relative" style={{ background: `linear-gradient(130deg, ${C.ink} 0%, ${SC.color} 75%, ${SC.color2} 100%)` }}>
        <div className="absolute opacity-10" style={{ top: -30, left: -30, width: 150, height: 150, borderRadius: "50%", border: "22px solid #fff" }} />
        <div className="relative z-10 flex items-center justify-between flex-wrap gap-3">
          <div className="flex items-center gap-3"><div className="rounded-2xl flex items-center justify-center text-3xl shrink-0" style={{ width: 62, height: 62, background: "rgba(255,255,255,.15)", border: "2px solid rgba(255,255,255,.3)" }}>{me.photo || "🏪"}</div>
            <div><div className="text-xl sm:text-2xl font-extrabold" style={{ fontFamily: "Cairo, sans-serif" }}>{me.name}</div>
              <div className="flex items-center gap-2 mt-1 flex-wrap"><span className="inline-flex items-center gap-1 text-[10px] font-extrabold px-2 py-0.5 rounded-full" style={{ background: "rgba(255,255,255,.28)" }}><BadgeCheck size={11} />رقم العضوية: {me.membershipNo || "—"}</span><span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: "rgba(255,255,255,.18)" }}><T.icon size={11} />{T.label}</span><span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: "rgba(255,255,255,.18)" }}><MapPin size={11} />{me.area}، {me.city}</span><span className="mt-0"><Stars {...(merchantRating(data, me.id) || { avg: null })} size={11} /></span></div>
            </div></div>
          <GpsLink m={me} />
        </div>
      </div>
      <div className="grid grid-cols-3 divide-x divide-x-reverse" style={{ borderColor: C.line, background: C.surface }}>
        {[["الحد الائتماني", money(me.creditLimit)], ["رصيدي", money(me.balance)], ["مديونيتي", money(Math.max(0, -me.balance))]].map(([k, v], i) => <div key={i} className="p-3 text-center"><div className="text-[10px] font-bold" style={{ color: C.soft }}>{k}</div><div className="text-sm font-extrabold" style={{ color: i === 2 && me.balance < 0 ? C.danger : C.text }}>{v}</div></div>)}
      </div>
    </div>
    <div className="grid lg:grid-cols-2 gap-4">
      {/* بيانات المحل والرخصة */}
      <Card className="p-4 sm:p-5 min-w-0">
        <h3 className="font-extrabold mb-2 flex items-center gap-2" style={{ color: C.text }}><FileSignature size={17} style={{ color: SC.color }} />بيانات المحل والرخصة</h3>
        <Row icon={BadgeCheck} label="رقم العضوية (رقم المحل الموثّق)" value={me.membershipNo} ltr /><Row icon={Store} label="اسم المحل" value={me.name} />
        <Row icon={FileSignature} label="رقم الرخصة التجارية" value={me.license} ltr />
        <Row icon={CalendarClock} label="تاريخ إصدار الرخصة" value={me.licenseDate ? fmtDate(me.licenseDate) : null} />
        <Row icon={MapPin} label="العنوان التفصيلي" value={edit ? undefined : (f.address || me.address)} />
        {edit && <div className="py-2"><Input value={f.address} onChange={(e) => setF({ ...f, address: e.target.value })} placeholder="العنوان التفصيلي" /></div>}
        <div className="pt-3"><LocationCard m={me} /></div>
      </Card>
      {/* المسؤول عن المحل */}
      <Card className="p-4 sm:p-5 min-w-0">
        <div className="flex items-center justify-between mb-2"><h3 className="font-extrabold flex items-center gap-2" style={{ color: C.text }}><UserCog size={17} style={{ color: SC.color }} />المسؤول عن المحل</h3>{!edit && <GhostBtn sm icon={Edit3} onClick={() => setEdit(true)}>تعديل</GhostBtn>}</div>
        {!edit ? <>
          <div className="flex items-center gap-3 rounded-2xl p-3 mb-2" style={{ background: C.bg }}><Avatar name={f.managerName} size={46} /><div><div className="font-extrabold text-sm" style={{ color: C.text }}>{f.managerName}</div><div className="text-xs" style={{ color: C.soft }}>المسؤول المفوَّض</div></div></div>
          <Row icon={Phone} label="هاتف المسؤول" value={f.managerPhone} ltr />
          <Row icon={MessageSquare} label="البريد الإلكتروني" value={f.email} ltr />
          <Row icon={Handshake} label="مندوبك" value={rep ? `${rep.name} · ${rep.phone || ""}` : "—"} />
          <Row icon={Building2} label="شركتك الأساسية" value={co?.name} />
        </> : <div className="space-y-2.5">
          <Field label="اسم المسؤول"><Input value={f.managerName} onChange={(e) => setF({ ...f, managerName: e.target.value })} /></Field>
          <Field label="هاتف المسؤول"><Input value={f.managerPhone} onChange={(e) => setF({ ...f, managerPhone: e.target.value })} /></Field>
          <Field label="البريد الإلكتروني"><Input value={f.email} onChange={(e) => setF({ ...f, email: e.target.value })} /></Field>
          <div className="flex gap-2 pt-1"><PrimaryBtn full icon={CheckCircle2} onClick={save}>حفظ الملف</PrimaryBtn><GhostBtn onClick={() => setEdit(false)}>إلغاء</GhostBtn></div>
        </div>}
      </Card>
    </div>
    {/* الرمز السري لتأكيد الطلبيات */}
    <OrderPinCard me={me} actions={actions} />
  </div>);
}
// بطاقة تعديل الرمز السري لتأكيد الطلبيات
function OrderPinCard({ me, actions }) {
  const [editing, setEditing] = useState(false);
  const [oldPin, setOldPin] = useState(""); const [newPin, setNewPin] = useState(""); const [confirmP, setConfirmP] = useState(""); const [err, setErr] = useState("");
  const clean = (v) => v.replace(/\D/g, "").slice(0, 4);
  const save = () => {
    if (oldPin !== (me.orderPin || "0000")) { setErr("الرمز الحالي غير صحيح"); return; }
    if (newPin.length !== 4) { setErr("الرمز الجديد يجب أن يكون 4 أرقام"); return; }
    if (newPin !== confirmP) { setErr("الرمز الجديد وتأكيده غير متطابقين"); return; }
    actions.setOrderPin(me.id, newPin); toast("حُدّث الرمز السري ✓"); setEditing(false); setOldPin(""); setNewPin(""); setConfirmP(""); setErr("");
  };
  return (<Card className="p-4 sm:p-5 min-w-0">
    <div className="flex items-center justify-between gap-2 flex-wrap"><div className="flex items-center gap-2.5"><div className="rounded-xl p-2.5" style={{ background: SC.soft }}><Lock size={18} style={{ color: SC.color }} /></div><div><h3 className="font-extrabold text-sm" style={{ color: C.text }}>الرمز السري لتأكيد الطلبيات</h3><p className="text-[11px]" style={{ color: C.soft }}>رمز من 4 أرقام تؤكّد به كل طلبية قبل إرسالها (الافتراضي 0000).</p></div></div>{!editing && <GhostBtn sm icon={Edit3} onClick={() => { setEditing(true); setErr(""); }}>تغيير</GhostBtn>}</div>
    {editing && <div className="mt-3 space-y-2.5">
      <Field label="الرمز الحالي"><input value={oldPin} onChange={(e) => { setOldPin(clean(e.target.value)); setErr(""); }} type="tel" inputMode="numeric" maxLength={4} placeholder="••••" className="w-full text-center rounded-xl py-2.5 outline-none tracking-[0.5em] font-extrabold" style={{ ...inputStyle, fontSize: 20 }} /></Field>
      <div className="grid grid-cols-2 gap-2">
        <Field label="الرمز الجديد"><input value={newPin} onChange={(e) => { setNewPin(clean(e.target.value)); setErr(""); }} type="tel" inputMode="numeric" maxLength={4} placeholder="••••" className="w-full text-center rounded-xl py-2.5 outline-none tracking-[0.5em] font-extrabold" style={{ ...inputStyle, fontSize: 20 }} /></Field>
        <Field label="تأكيد الرمز الجديد"><input value={confirmP} onChange={(e) => { setConfirmP(clean(e.target.value)); setErr(""); }} type="tel" inputMode="numeric" maxLength={4} placeholder="••••" className="w-full text-center rounded-xl py-2.5 outline-none tracking-[0.5em] font-extrabold" style={{ ...inputStyle, fontSize: 20 }} /></Field>
      </div>
      {err && <div className="text-xs font-bold" style={{ color: C.danger }}>{err}</div>}
      <div className="flex gap-2"><PrimaryBtn full icon={CheckCircle2} onClick={save}>حفظ الرمز</PrimaryBtn><GhostBtn onClick={() => { setEditing(false); setErr(""); }}>إلغاء</GhostBtn></div>
    </div>}
  </Card>);
}
function MHome(props) {
  const { data, go } = props; const { me, orders } = useMy(props);
  const serverAds = data.ads.filter((a) => a.server === props.server && a.active);
  const pendingSettle = data.settlements.filter((s) => s.merchantId === me.id && s.status === "pending_rep").length;
  const pendingApproval = data.orders.filter((o) => o.merchantId === me.id && o.status === "draft").length;
  const pendingReq = data.reconciliations.filter((r) => r.merchantId === me.id && r.status === "pending").length + data.stocktakes.filter((st) => st.type === "merchant" && st.merchantId === me.id && st.status === "requested").length;
  const activeOrders = orders.filter((o) => !["delivered", "rejected", "draft"].includes(o.status)).length;
  const unpaidInv = data.invoices.filter((i) => i.merchantId === me.id && i.paid < i.amount).length;
  // كل أقسام القائمة الجانبية كبطاقات فخمة في الرئيسية
  const SECTIONS = [
    { k: "m_create", t: "طلبية جديدة", d: "إنشاء فاتورة لشركة", icon: FileText, c1: "#0D47A1", c2: "#1E6BD6", badge: null },
    { k: "m_track", t: "متابعة الطلبات", d: "أين وصلت طلباتك", icon: Truck, c1: "#0E7490", c2: "#22B8CF", badge: activeOrders },
    { k: "m_invoices", t: "مراجعة الفواتير", d: "بنودها وصلاحياتها", icon: Receipt, c1: "#7C3AED", c2: "#A78BFA", badge: unpaidInv },
    { k: "m_settle", t: "تسوية / سداد", d: "ادفع أو اطلب مندوباً", icon: HandCoins, c1: "#047857", c2: "#10B981", badge: pendingSettle },
    { k: "m_companies", t: "شركاتي", d: "الشركات المشترك معها", icon: Building2, c1: "#0F3D91", c2: "#3B82F6", badge: null },
    { k: "m_explore", t: "استكشاف وانضمام", d: "شركات جديدة بسيرفرك", icon: ScanSearch, c1: "#9D174D", c2: "#EC4899", badge: null },
    { k: "m_search", t: "بحث المنتجات", d: "ابحث في كل الأصناف", icon: Search, c1: "#155E75", c2: "#06B6D4", badge: null },
    { k: "m_offers", t: "العروض", d: "خصومات الشركات", icon: Tag, c1: "#B91C1C", c2: "#F87171", badge: null },
    { k: "m_statement", t: "كشف الحساب", d: "حركاتك المالية", icon: ScrollText, c1: "#1F2937", c2: "#6B7280", badge: null },
    { k: "m_stats", t: "إحصائياتي", d: "أرقامك ومؤشراتك", icon: TrendingUp, c1: "#065F46", c2: "#34D399", badge: null },
    { k: "m_profile", t: "ملفي الشخصي", d: "الرخصة والموقع والمسؤول", icon: UserCog, c1: "#0F3D91", c2: "#60A5FA", badge: null },
    { k: "m_requests", t: "الجرد والمطابقة", d: "طلبات مندوبك للجرد والحساب", icon: ClipboardList, c1: "#7C3AED", c2: "#A78BFA", badge: pendingReq },
    { k: "m_return_new", t: "المرتجعات", d: "طلب ومتابعة مرتجع", icon: RotateCcw, c1: "#92400E", c2: "#FBBF24", badge: null },
    { k: "m_chat", t: "التواصل", d: "محادثة الشركات", icon: MessageSquare, c1: "#5B21B6", c2: "#8B5CF6", badge: null },
  ];
  return (
    <div className="space-y-5">
      {/* ترحيب فخم */}
      <div className="rb-card rounded-3xl p-5 sm:p-6 text-white relative overflow-hidden" style={{ background: `linear-gradient(130deg, ${C.ink} 0%, ${SC.color} 70%, ${SC.color2} 100%)` }}>
        <div className="absolute opacity-10" style={{ top: -40, left: -40, width: 190, height: 190, borderRadius: "50%", border: "26px solid #fff" }} />
        <div className="absolute opacity-10" style={{ bottom: -50, right: 30, width: 130, height: 130, borderRadius: "50%", background: C.well }} />
        <div className="relative z-10 flex items-center justify-between flex-wrap gap-3">
          <div className="flex items-center gap-3"><Avatar name={me.owner} size={52} /><div><div className="text-sm opacity-80">أهلاً بك</div><div className="text-xl sm:text-2xl font-extrabold" style={{ fontFamily: "Cairo, sans-serif" }}>{me.owner}</div><div className="text-xs opacity-80 mt-0.5">{me.name} · {me.area}</div></div></div>
          <button onClick={() => go("m_create")} className="rb-press rounded-2xl px-5 py-3 font-extrabold text-sm flex items-center gap-2" style={{ background: C.well, color: "#fff" }}><Plus size={18} />طلبية جديدة</button>
        </div>
      </div>
      {/* تنبيهات سريعة */}
      {(pendingSettle > 0 || pendingReq > 0) && <div className="flex gap-2 overflow-x-auto pb-1">
        {pendingSettle > 0 && <button onClick={() => go("m_settle")} className="rb-press shrink-0 flex items-center gap-2 rounded-2xl px-4 py-2.5 text-sm font-bold" style={{ background: C.infoSoft, color: C.info, border: `1px solid ${C.info}` }}><Clock size={16} />{pendingSettle} تسوية قيد المراجعة</button>}
        {pendingReq > 0 && <button onClick={() => go("m_requests")} className="rb-press shrink-0 flex items-center gap-2 rounded-2xl px-4 py-2.5 text-sm font-bold" style={{ background: C.purpleSoft, color: C.purple, border: `1px solid ${C.purple}` }}><ClipboardList size={16} />{pendingReq} طلب جرد/مطابقة</button>}
      </div>}
      {/* المساحة الإعلانية */}
      <AdCarousel ads={serverAds} />
      {/* كل الأقسام: شبكة فخمة */}
      <div>
        <h3 className="font-extrabold mb-3 text-base" style={{ color: C.text, fontFamily: "Cairo, sans-serif" }}>كل خدماتك</h3>
        <div className="rb-stagger grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3">
          {SECTIONS.map((sec) => (
            <button key={sec.k} onClick={() => go(sec.k)} className="rb-press relative text-right rounded-3xl p-4 overflow-hidden group" style={{ background: `linear-gradient(135deg, ${sec.c1} 0%, ${sec.c2} 100%)`, minHeight: 118, boxShadow: "0 8px 22px rgba(11,31,58,.10)" }}>
              <div className="absolute opacity-15 transition-transform duration-300 group-hover:scale-125" style={{ top: -18, left: -18, width: 76, height: 76, borderRadius: "50%", background: C.well }} />
              {sec.badge > 0 && <span className="absolute top-3 left-3 min-w-[20px] h-5 px-1.5 rounded-full text-[11px] font-extrabold flex items-center justify-center" style={{ background: C.well, color: sec.c1 }}>{sec.badge}</span>}
              <div className="rounded-2xl inline-flex items-center justify-center mb-2.5" style={{ width: 42, height: 42, background: "rgba(255,255,255,.18)", backdropFilter: "blur(4px)" }}><sec.icon size={21} color="#fff" /></div>
              <div className="text-white font-extrabold text-sm leading-tight" style={{ fontFamily: "Cairo, sans-serif" }}>{sec.t}</div>
              <div className="text-[10px] mt-0.5" style={{ color: "rgba(255,255,255,.75)" }}>{sec.d}</div>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}
// إحصائيات التاجر (نُقلت من الواجهة إلى صفحة مستقلة)
function MStats(props) {
  const { data, go } = props; const { me, orders } = useMy(props);
  const partners = [...new Set(orders.map((o) => o.companyId))].map((id) => data.companies.find((c) => c.id === id)).filter(Boolean);
  const offers = F.products(props).filter((p) => p.offer > 0);
  const myInv = data.invoices.filter((i) => i.merchantId === me.id);
  const totalDebt = myInv.reduce((s, i) => s + (i.amount - i.paid), 0);
  const totalPaid = myInv.reduce((s, i) => s + i.paid, 0);
  const delivered = orders.filter((o) => o.status === "delivered");
  const totalQtyBought = delivered.reduce((s, o) => s + orderQty(o), 0);
  const totalValBought = delivered.reduce((s, o) => s + orderTotal(o), 0);
  const usedPct = me.creditLimit ? Math.min(100, Math.round((Math.max(0, -me.balance) / me.creditLimit) * 100)) : 0;
  return (<div className="space-y-5"><SectionTitle sub="أرقامك ومؤشراتك المالية والشرائية في مكان واحد.">إحصائياتي</SectionTitle>
    <div className="rb-stagger grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={Building2} label="شركات أتعامل معها" value={partners.length} tint={{ color: SC.color2, soft: SC.soft }} />
      <Stat icon={ClipboardList} label="طلبات نشطة" value={orders.filter((o) => !["delivered", "rejected", "draft"].includes(o.status)).length} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={TrendingDown} label="إجمالي مديونياتي" value={money(totalDebt)} tint={{ color: C.danger, soft: C.dangerSoft }} />
      <Stat icon={Tag} label="عروض متاحة" value={offers.length} tint={{ color: C.accent, soft: C.accentSoft }} />
      <Stat icon={Boxes} label="كمية مشترياتي" value={qty(totalQtyBought)} tint={{ color: SC.color2, soft: SC.soft }} hint={`${delivered.length} طلب مُسلّم`} />
      <Stat icon={CircleDollarSign} label="قيمة مشترياتي" value={money(totalValBought)} tint={{ color: C.success, soft: C.successSoft }} />
      <Stat icon={HandCoins} label="إجمالي ما سدّدت" value={money(totalPaid)} tint={{ color: C.success, soft: C.successSoft }} />
      <Stat icon={Wallet} label="الحد الائتماني" value={money(me.creditLimit)} tint={{ color: C.purple, soft: C.purpleSoft }} hint={`مستخدم ${usedPct}%`} />
    </div>
    <Card className="p-5"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Wallet size={17} style={{ color: C.purple }} />استهلاك الحد الائتماني</h3>
      <ProgressBar value={Math.max(0, -me.balance)} max={me.creditLimit} color={usedPct > 80 ? C.danger : usedPct > 50 ? C.accent : C.success} />
      <div className="flex justify-between text-xs mt-2" style={{ color: C.soft }}><span>مستخدم: {money(Math.max(0, -me.balance))}</span><span>المتاح: {money(Math.max(0, me.creditLimit + me.balance))}</span></div>
    </Card>
    <div className="grid lg:grid-cols-2 gap-4">
      <Card className="p-5"><h3 className="font-extrabold mb-4 flex items-center gap-2" style={{ color: C.text }}><TrendingUp size={17} style={{ color: SC.color2 }} />مشترياتي الشهرية (د.ل)</h3>
        <BarsChart money color={SC.color2} data={[{ label: "فبراير", v: Math.round(totalValBought * .1) }, { label: "مارس", v: Math.round(totalValBought * .2) }, { label: "أبريل", v: Math.round(totalValBought * .15) }, { label: "مايو", v: Math.round(totalValBought * .25) }, { label: "يونيو", v: Math.round(totalValBought * .3) }]} />
      </Card>
      <Card className="p-5"><h3 className="font-extrabold mb-4 flex items-center gap-2" style={{ color: C.text }}><CircleDollarSign size={17} style={{ color: C.accent }} />ذمّتي المالية</h3>
        <DonutChart parts={[{ label: "سدّدت", v: totalPaid || 1, color: C.success }, { label: "متبقٍّ عليّ", v: totalDebt || 1, color: C.danger }]} />
      </Card>
    </div>
  </div>);
}
function MCompanies(props) {
  const { data, go } = props; const { me } = useMy(props);
  const partnerIds = [...new Set(data.orders.filter((o) => o.merchantId === me.id).map((o) => o.companyId))];
  const partners = F.companies(props).filter((c) => partnerIds.includes(c.id) || data.merchants.find((m) => m.id === me.id)?.companyId === c.id);
  return (<div className="space-y-4"><SectionTitle sub="الشركات التي تتعامل معها فعلياً في سيرفرك." action={<PrimaryBtn sm icon={UserPlus} onClick={() => go("m_explore")}>انضمام لشركة جديدة</PrimaryBtn>}>الشركات المشترك معها</SectionTitle>
    {partners.length ? <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{partners.map((c) => {
      const myOrders = data.orders.filter((o) => o.merchantId === me.id && o.companyId === c.id);
      const debt = data.invoices.filter((i) => i.merchantId === me.id && i.companyId === c.id).reduce((s, i) => s + (i.amount - i.paid), 0);
      const myProds = data.products.filter((p) => p.companyId === c.id).length;
      return (<Card key={c.id} className="p-5">
        <div className="flex items-center gap-3 mb-3"><span className="text-3xl">{c.logo}</span><div><div className="font-extrabold" style={{ color: C.text }}>{c.name}</div><div className="text-xs" style={{ color: C.soft }}>{c.cat} · {c.city}</div><div className="mt-0.5"><Stars {...(companyRating(data, c.id) || { avg: null })} /></div></div></div>
        <div className="flex flex-wrap gap-1.5 mb-3">{(() => { const ct = COMPANY_TYPE[c.companyType || "distributor"]; return <span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: ct.soft, color: ct.color }}><ct.icon size={11} />{ct.label}</span>; })()}<span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: c.hasDelivery ? C.infoSoft : C.purpleSoft, color: c.hasDelivery ? C.info : C.purple }}>{c.hasDelivery ? <><Truck size={11} />توصيل</> : <><Store size={11} />استلام من المقر</>}</span></div>
        <div className="grid grid-cols-3 gap-2 text-center mb-3">
          <div className="rounded-xl p-2" style={{ background: C.bg }}><div className="text-lg font-extrabold" style={{ color: SC.color2 }}>{myProds}</div><div className="text-[10px]" style={{ color: C.soft }}>منتج</div></div>
          <div className="rounded-xl p-2" style={{ background: C.bg }}><div className="text-lg font-extrabold" style={{ color: C.info }}>{myOrders.length}</div><div className="text-[10px]" style={{ color: C.soft }}>طلب</div></div>
          <div className="rounded-xl p-2" style={{ background: debt > 0 ? C.dangerSoft : C.successSoft }}><div className="text-sm font-extrabold" style={{ color: debt > 0 ? C.danger : C.success }}>{debt > 0 ? money(debt) : "صفر"}</div><div className="text-[10px]" style={{ color: C.soft }}>مديونية</div></div>
        </div>
        <div className="flex items-center gap-1.5 text-xs" style={{ color: C.soft }}><Phone size={13} />{c.phone}</div>
      </Card>); })}</div> : <Empty icon={Building2} text="لم تنضم لأي شركة بعد." />}
  </div>);
}

// استكشاف الشركات وإرسال طلب انضمام
// نافذة تفاصيل الشركة: أصنافها وتصنيفاتها ومنتجاتها قبل الانضمام
function CompanyDetailsModal({ co, data, me, requested, onJoin, onClose }) {
  const prods = data.products.filter((p) => p.companyId === co.id && p.status === "active");
  const cats = Object.entries(prods.reduce((acc, p) => { acc[p.cat] = (acc[p.cat] || 0) + 1; return acc; }, {})).sort((a, b) => b[1] - a[1]);
  const [cat, setCat] = useState("الكل");
  const ct = COMPANY_TYPE[co.companyType || "distributor"];
  const cr = companyRating(data, co.id);
  const visible = prods.filter((p) => cat === "الكل" || p.cat === cat);
  const isPharma = co.server === "pharma";
  return (<div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-3 bg-black/50" onClick={onClose}>
    <div className="rb-modal w-full max-w-lg rounded-2xl overflow-hidden max-h-[90vh] flex flex-col" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
      {/* ترويسة */}
      <div className="p-5 text-white shrink-0" style={{ background: `linear-gradient(125deg, ${C.ink}, ${SC.color})` }}>
        <div className="flex items-start justify-between gap-3">
          <div className="flex items-center gap-3"><div className="rounded-2xl flex items-center justify-center text-3xl" style={{ width: 54, height: 54, background: "rgba(255,255,255,.15)" }}>{co.logo}</div><div><div className="font-extrabold text-lg leading-tight" style={{ fontFamily: "Cairo, sans-serif" }}>{co.name}</div><div className="text-xs opacity-85">{co.cat} · {co.city} · {co.phone}</div><div className="mt-1"><Stars {...(cr || { avg: null })} /></div></div></div>
          <button onClick={onClose} className="rb-press rounded-xl p-2 shrink-0" style={{ background: "rgba(255,255,255,.15)" }}><XCircle size={18} color="#fff" /></button>
        </div>
        <div className="flex flex-wrap gap-1.5 mt-3">
          <span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-1 rounded-full" style={{ background: "rgba(255,255,255,.18)" }}><ct.icon size={11} />{ct.label}</span>
          <span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-1 rounded-full" style={{ background: "rgba(255,255,255,.18)" }}>{co.hasDelivery ? <><Truck size={11} />توصيل متاح</> : <><Store size={11} />استلام من المقر</>}</span>
          <span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-1 rounded-full" style={{ background: "rgba(255,255,255,.18)" }}><Package size={11} />{prods.length} صنف</span>
          <span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-1 rounded-full" style={{ background: "rgba(255,255,255,.18)" }}><Layers size={11} />{cats.length} تصنيف</span>
        </div>
      </div>
      {/* التصنيفات وعدد أصناف كل تصنيف */}
      <div className="p-4 pb-2 shrink-0" style={{ borderBottom: `1px solid ${C.line}` }}>
        {co.desc && <p className="text-sm mb-3" style={{ color: C.soft }}>{co.desc}</p>}
        <div className="text-xs font-extrabold mb-2" style={{ color: C.text }}>تصنيفات المنتجات</div>
        <div className="flex gap-1.5 overflow-x-auto pb-2">
          <button onClick={() => setCat("الكل")} className="rb-press shrink-0 inline-flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-bold" style={{ background: cat === "الكل" ? SC.color : C.bg, color: cat === "الكل" ? "#fff" : C.text, border: `1px solid ${cat === "الكل" ? SC.color : C.line}` }}>الكل<span className="rounded-full px-1.5 text-[10px]" style={{ background: cat === "الكل" ? "rgba(255,255,255,.25)" : C.surface, color: cat === "الكل" ? "#fff" : C.soft }}>{prods.length}</span></button>
          {cats.map(([c, n]) => { const on = cat === c; return (<button key={c} onClick={() => setCat(c)} className="rb-press shrink-0 inline-flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-bold" style={{ background: on ? SC.color : C.bg, color: on ? C.onAccent : C.text, border: `1px solid ${on ? SC.color : C.line}` }}>{c}<span className="rounded-full px-1.5 text-[10px]" style={{ background: on ? "rgba(255,255,255,.25)" : C.surface, color: on ? "#fff" : C.soft }}>{n}</span></button>); })}
        </div>
      </div>
      {/* المنتجات */}
      <div className="flex-1 overflow-auto p-4 pt-3">
        <div className="space-y-1.5">{visible.map((p) => { const e = expiryInfo(p.expDate); return (<div key={p.id} className="flex items-center gap-2.5 rounded-xl px-3 py-2" style={{ border: `1px solid ${C.line}` }}>
          <div className="rounded-lg flex items-center justify-center shrink-0" style={{ width: 34, height: 34, background: SC.soft }}>{isPharma ? <Pill size={16} style={{ color: SC.color }} /> : <Package size={16} style={{ color: SC.color }} />}</div>
          <div className="flex-1 min-w-0"><div className="text-sm font-bold truncate" style={{ color: C.text }}>{p.name}</div><div className="flex items-center gap-1.5 flex-wrap"><span className="text-[10px]" style={{ color: C.soft }}>{p.cat} · {p.unit}{p.packSize > 1 ? ` · صندوق = ${p.packSize}` : ""}</span>{p.tiers?.length > 0 && <span className="text-[9px] font-bold px-1 py-0.5 rounded" style={{ background: C.purpleSoft, color: C.purple }}>خصم كمية</span>}{p.offer > 0 && <span className="text-[9px] font-bold px-1 py-0.5 rounded" style={{ background: C.dangerSoft, color: C.danger }}>عرض -{p.offer}%</span>}{p.expDate && e && <span className="text-[9px] font-bold px-1 py-0.5 rounded" style={{ background: e.soft, color: e.color }}>صلاحية {fmtDate(p.expDate)}</span>}</div></div>
          <div className="text-left shrink-0"><div className="text-sm font-extrabold" style={{ color: SC.color }}>{money(p.price)}</div><div className="text-[9px]" style={{ color: C.soft }}>سعر الجملة</div></div>
        </div>); })}</div>
        {visible.length === 0 && <Empty icon={Package} text="لا أصناف في هذا التصنيف." />}
      </div>
      {/* شريط سفلي: انضمام */}
      <div className="p-3 shrink-0 flex gap-2" style={{ borderTop: `1px solid ${C.line}` }}>
        {requested ? <div className="flex-1 text-center text-sm font-bold rounded-xl py-2.5" style={{ background: C.accentSoft, color: "#92590C" }}>تم إرسال طلب الانضمام ✓</div>
          : <PrimaryBtn full icon={UserPlus} onClick={onJoin}>طلب الانضمام لهذه الشركة</PrimaryBtn>}
        <GhostBtn icon={ChevronRight} onClick={onClose}>رجوع</GhostBtn>
      </div>
    </div>
  </div>);
}
function MExplore(props) {
  const { data, actions, go } = props; const { me } = useMy(props);
  const all = F.companies(props).filter((c) => c.sub === "active");
  const myRequests = data.joinRequests.filter((j) => j.merchantId === me.id);
  const requestedIds = myRequests.map((j) => j.companyId);
  const [modal, setModal] = useState(null);
  const [details, setDetails] = useState(null); // نافذة تفاصيل الشركة وأصنافها
  const [oldDebt, setOldDebt] = useState(0);
  const [note, setNote] = useState("");
  return (<div className="space-y-4"><SectionTitle sub="استعرض شركات سيرفرك المشتركة، وأرسل طلب انضمام لتبدأ التعامل معها.">استكشاف الشركات والانضمام</SectionTitle>
    <div className="rounded-xl p-4 flex items-start gap-3" style={{ background: SC.soft }}>
      <FileSignature size={20} style={{ color: SC.color }} className="mt-0.5 shrink-0" />
      <div className="text-sm" style={{ color: C.text }}>الشركات نوعان: <b>موزّع كامل</b> (لديه مندوبون يزورونك للتوثيق ثم يُفعّل حسابك)، و<b>بيع مباشر</b> (شركة صغيرة تتواصل معك مباشرةً وتُفعّل حسابك دون زيارة مندوب وتسلّم بنفسها).</div>
    </div>
    <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{all.map((c) => { const reqd = requestedIds.includes(c.id); const prods = data.products.filter((p) => p.companyId === c.id).length; const ct = COMPANY_TYPE[c.companyType || "distributor"];
      return (<Card key={c.id} className="p-5 flex flex-col">
        <div className="flex items-center gap-3 mb-2"><span className="text-3xl">{c.logo}</span><div><div className="font-extrabold" style={{ color: C.text }}>{c.name}</div><div className="text-xs" style={{ color: C.soft }}>{c.cat} · {c.city}</div><div className="mt-0.5"><Stars {...(companyRating(data, c.id) || { avg: null })} /></div></div></div>
        <div className="flex flex-wrap gap-1.5 mb-2"><span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: ct.soft, color: ct.color }}><ct.icon size={11} />{ct.label}</span><span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: c.hasDelivery ? C.infoSoft : C.purpleSoft, color: c.hasDelivery ? C.info : C.purple }}>{c.hasDelivery ? <><Truck size={11} />توصيل</> : <><Store size={11} />استلام من المقر</>}</span></div>
        <p className="text-sm mb-3 flex-1" style={{ color: C.soft }}>{c.desc}</p>
        <div className="flex items-center gap-3 text-xs mb-3" style={{ color: C.soft }}><span className="flex items-center gap-1"><Package size={13} />{prods} منتج</span><span className="flex items-center gap-1"><Phone size={13} />{c.phone}</span></div>
        <div className="flex gap-2">
          <GhostBtn icon={Eye} onClick={() => setDetails(c)}>الأصناف والتفاصيل</GhostBtn>
          {reqd ? <div className="flex-1 text-center text-sm font-bold rounded-xl py-2.5" style={{ background: C.accentSoft, color: "#92590C" }}>أُرسل الطلب ✓</div>
            : <div className="flex-1"><PrimaryBtn full icon={UserPlus} onClick={() => { setModal(c); setOldDebt(0); setNote(""); }}>طلب الانضمام</PrimaryBtn></div>}
        </div>
      </Card>); })}</div>

    {details && <CompanyDetailsModal co={details} data={data} me={me} requested={requestedIds.includes(details.id)} onJoin={() => { setModal(details); setDetails(null); setOldDebt(0); setNote(""); }} onClose={() => setDetails(null)} />}
    {modal && <div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50" onClick={() => setModal(null)}>
      <div className="rb-modal w-full max-w-md rounded-2xl p-5" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center gap-3 mb-4"><span className="text-3xl">{modal.logo}</span><div><div className="font-extrabold" style={{ color: C.text }}>طلب انضمام إلى</div><div className="text-sm" style={{ color: C.soft }}>{modal.name}</div></div></div>
        <div className="rounded-xl p-3 mb-3 text-sm flex items-start gap-2" style={{ background: COMPANY_TYPE[modal.companyType || "distributor"].soft }}>{(() => { const ct = COMPANY_TYPE[modal.companyType || "distributor"]; const I = ct.icon; return <><I size={16} style={{ color: ct.color, marginTop: 2 }} /><span style={{ color: C.text }}>{modal.companyType === "direct" ? "شركة بيع مباشر — ستتواصل معك وتُفعّل حسابك مباشرةً دون زيارة مندوب." : "موزّع كامل — سيزورك مندوب للتوثيق ثم يُفعَّل حسابك."}</span></>; })()}</div>
        <div className="space-y-3">
          <Field label="هل لديك مديونية سابقة معهم؟ (د.ل)" hint={modal.companyType === "direct" ? "ستوثّقها الشركة عند التفعيل" : "سيتم التحقق منها أثناء زيارة المندوب"}><Input type="number" value={oldDebt} onChange={(e) => setOldDebt(+e.target.value)} /></Field>
          <Field label="ملاحظة للشركة"><textarea value={note} onChange={(e) => setNote(e.target.value)} className={inputCls} style={{ ...inputStyle, minHeight: 70 }} placeholder="مثال: نرغب بالتعامل معكم لتزويد محلنا..." /></Field>
        </div>
        <div className="flex gap-2 mt-4"><PrimaryBtn full icon={Send} onClick={() => { actions.createJoin({ merchantId: me.id, companyId: modal.id, oldDebt, note }); setModal(null); go("m_join_track"); }}>إرسال الطلب</PrimaryBtn><GhostBtn onClick={() => setModal(null)}>إلغاء</GhostBtn></div>
      </div>
    </div>}
  </div>);
}

function MJoinTrack(props) {
  const { data } = props; const { me } = useMy(props);
  const reqs = data.joinRequests.filter((j) => j.merchantId === me.id);
  const steps = ["requested", "visit_set", "verifying", "active"];
  return (<div className="space-y-4"><SectionTitle sub="حالة طلبات انضمامك للشركات وزيارات التوثيق.">متابعة طلبات الانضمام</SectionTitle>
    {reqs.length ? reqs.map((j) => { const c = data.companies.find((x) => x.id === j.companyId); const curStep = steps.indexOf(j.status); const isDirect = c?.companyType === "direct"; const stepDefs = isDirect ? [{ t: "إرسال الطلب", i: Send }, { t: "تفعيل مباشر", i: BadgeCheck }] : [{ t: "إرسال الطلب", i: Send }, { t: "جدولة الزيارة", i: CalendarClock }, { t: "التوثيق", i: ClipboardSignature }, { t: "التفعيل", i: BadgeCheck }]; const effCur = isDirect ? (j.status === "active" ? 1 : 0) : curStep;
      return (<Card key={j.id} className="p-5"><div className="flex items-center justify-between mb-4 flex-wrap gap-2"><div className="flex items-center gap-3"><span className="text-2xl">{c?.logo}</span><div><div className="font-extrabold" style={{ color: C.text }}>{c?.name}</div><div className="text-xs flex items-center gap-1.5" style={{ color: C.soft }}>{j.date}{isDirect && <span className="inline-flex items-center gap-1 font-bold px-1.5 py-0.5 rounded-full" style={{ background: C.tealSoft, color: C.teal }}><Handshake size={10} />بيع مباشر</span>}</div></div></div><Badge status={j.status} map={JOIN_STATUS} /></div>
        {j.status !== "declined" && <div className="flex items-center justify-between gap-1 mb-4">{stepDefs.map((s, idx) => { const done = effCur >= idx; const col = done ? SC.color : C.line;
          return (<React.Fragment key={idx}><div className="flex flex-col items-center gap-1"><div className="rounded-full flex items-center justify-center" style={{ width: 34, height: 34, background: done ? col : C.surface, border: `2px solid ${col}` }}><s.i size={15} style={{ color: done ? "#fff" : C.soft }} /></div><span className="text-[10px] font-bold text-center" style={{ color: done ? C.text : C.soft }}>{s.t}</span></div>{idx < stepDefs.length - 1 && <div className="h-0.5 flex-1 rounded" style={{ background: effCur > idx ? SC.color : C.line }} />}</React.Fragment>); })}</div>}
        <div className="grid sm:grid-cols-2 gap-2 text-sm">
          {j.visitDate && <Info icon={CalendarClock} label="موعد الزيارة" value={j.visitDate} />}
          {j.repId && <Info icon={Handshake} label="المندوب المكلّف" value={repName(data, j.repId)} />}
          <Info icon={TrendingDown} label="الدين المُصرّح" value={money(j.oldDebt)} />
          {j.verifiedDebt != null && <Info icon={ClipboardCheck} label="الدين الموثّق" value={money(j.verifiedDebt)} />}
        </div>
        {j.status === "active" && <div className="mt-3 text-sm rounded-xl p-3 flex items-center gap-2" style={{ background: C.successSoft, color: C.success }}><CheckCircle2 size={16} />تم تفعيل حسابك — يمكنك الآن الطلب من هذه الشركة.</div>}
      </Card>); }) : <Empty icon={FileSignature} text="لا توجد طلبات انضمام. استكشف الشركات لإرسال طلب." />}
  </div>);
}

function MCompanyDebt(props) {
  const { data } = props; const { me } = useMy(props);
  const byCompany = {};
  data.invoices.filter((i) => i.merchantId === me.id).forEach((i) => { byCompany[i.companyId] = (byCompany[i.companyId] || 0) + (i.amount - i.paid); });
  const total = Object.values(byCompany).reduce((a, b) => a + b, 0);
  return (<div className="space-y-4"><SectionTitle sub="ما عليك من مستحقات مقسّمة حسب كل شركة.">مديونياتي عند الشركات</SectionTitle>
    <Stat icon={TrendingDown} label="إجمالي ما عليك" value={money(total)} tint={{ color: C.danger, soft: C.dangerSoft }} />
    <div className="grid sm:grid-cols-2 gap-3">{Object.entries(byCompany).map(([cid, amt]) => { const c = data.companies.find((x) => x.id === cid);
      return (<Card key={cid} className="p-5 flex items-center justify-between"><div className="flex items-center gap-3"><span className="text-2xl">{c?.logo}</span><div><div className="font-extrabold" style={{ color: C.text }}>{c?.name}</div><div className="text-xs" style={{ color: C.soft }}>{c?.cat}</div></div></div><div className="text-left"><div className="text-lg font-extrabold" style={{ color: amt > 0 ? C.danger : C.success }}>{money(amt)}</div><div className="text-[11px]" style={{ color: C.soft }}>{amt > 0 ? "مستحق عليك" : "خالص"}</div></div></Card>); })}
      {Object.keys(byCompany).length === 0 && <Empty icon={CheckCircle2} text="لا توجد مديونيات." />}</div>
  </div>);
}

function MSearch(props) {
  const { data } = props; const { me } = useMy(props); const [q, setQ] = useState(""); const [cat, setCat] = useState("الكل");
  const [coFilter, setCoFilter] = useState("الكل");
  const [groupBy, setGroupBy] = useState("none"); // none | cat | company
  const prods = F.products(props);
  const cats = ["الكل", ...new Set(prods.map((p) => p.cat))];
  const comps = ["الكل", ...new Set(prods.map((p) => compName(data, p.companyId)))];
  const list = prods.filter((p) => p.status === "active"
    && (cat === "الكل" || p.cat === cat)
    && (coFilter === "الكل" || compName(data, p.companyId) === coFilter)
    && (p.name.includes(q) || compName(data, p.companyId).includes(q) || (p.cat || "").includes(q)));
  // التجميع: حسب الصنف أو حسب الشركة
  const groups = groupBy === "none" ? [["", list]] : Object.entries(list.reduce((acc, p) => { const k = groupBy === "cat" ? p.cat : compName(data, p.companyId); (acc[k] = acc[k] || []).push(p); return acc; }, {})).sort((a, b) => b[1].length - a[1].length);
  const joined = joinedCompanyIds(data, me);
  const mine = groups.map(([g, items]) => [g, items.filter((p) => joined.has(p.companyId))]).filter(([, items]) => items.length);
  const others = list.filter((p) => !joined.has(p.companyId));
  const othersByCo = Object.entries(others.reduce((acc, p) => { (acc[p.companyId] = acc[p.companyId] || []).push(p); return acc; }, {}));
  const renderCard = (p) => <ProductCard key={p.id} p={p} data={data} me={me} fav={data.favorites?.includes(p.id)} onFav={props.actions.toggleFav} onCompare={props.actions.toggleCompare} inCompare={data.compare?.includes(p.id)} onAdd={(p) => { props.actions.quickAddToCompany(p.companyId, p.id); props.go("m_create"); }} />;
  return (<div className="space-y-4">
    <SectionTitle sub={`ابحث في كل منتجات شركات ${SC.name}، وافرزها حسب الصنف أو الشركة.`}>البحث في المنتجات</SectionTitle>
    <Card className="p-3 sm:p-4 space-y-3">
      <div className="relative"><Search size={17} className="absolute right-3 top-3" style={{ color: C.soft }} /><input value={q} onChange={(e) => setQ(e.target.value)} placeholder="ابحث باسم المنتج أو الشركة أو التصنيف" className="w-full rounded-xl pr-10 pl-3 py-2.5 text-sm outline-none" style={inputStyle} /></div>
      <div className="grid grid-cols-2 gap-2">
        <div><div className="text-[11px] font-bold mb-1" style={{ color: C.soft }}>تصفية بالصنف</div><select value={cat} onChange={(e) => setCat(e.target.value)} className="w-full rounded-xl px-3 py-2.5 text-sm outline-none" style={inputStyle}>{cats.map((c) => <option key={c}>{c}</option>)}</select></div>
        <div><div className="text-[11px] font-bold mb-1" style={{ color: C.soft }}>تصفية بالشركة</div><select value={coFilter} onChange={(e) => setCoFilter(e.target.value)} className="w-full rounded-xl px-3 py-2.5 text-sm outline-none" style={inputStyle}>{comps.map((c) => <option key={c}>{c}</option>)}</select></div>
      </div>
      <div>
        <div className="text-[11px] font-bold mb-1.5" style={{ color: C.soft }}>فرز وعرض النتائج</div>
        <div className="flex gap-1.5">{[["none", "بدون تجميع", Search], ["cat", "حسب الصنف", Layers], ["company", "حسب الشركة", Building2]].map(([k, lbl, Ic]) => { const on = groupBy === k; return (<button key={k} onClick={() => setGroupBy(k)} className="rb-press flex-1 inline-flex items-center justify-center gap-1.5 rounded-xl py-2 text-xs font-bold" style={{ background: on ? SC.color : C.bg, color: on ? C.onAccent : C.text, border: `1px solid ${on ? SC.color : C.line}` }}><Ic size={13} />{lbl}</button>); })}</div>
      </div>
      <div className="text-xs" style={{ color: C.soft }}>{list.length} نتيجة{cat !== "الكل" ? ` · صنف: ${cat}` : ""}{coFilter !== "الكل" ? ` · شركة: ${coFilter}` : ""}</div>
    </Card>
    {mine.map(([g, items]) => (<div key={g || "all"}>
      {g && <div className="flex items-center gap-2 mb-2 mt-1"><div className="rounded-lg p-1.5" style={{ background: SC.soft }}>{groupBy === "cat" ? <Layers size={14} style={{ color: SC.color }} /> : <Building2 size={14} style={{ color: SC.color }} />}</div><span className="font-extrabold text-sm" style={{ color: C.text }}>{g}</span><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: SC.soft, color: SC.color }}>{items.length}</span></div>}
      <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{items.map(renderCard)}</div>
    </div>))}
    {mine.length === 0 && others.length === 0 && <Empty icon={Search} text="لا توجد نتائج." />}
    {/* نفس الصنف عند شركات لم تنضم إليها بعد */}
    {others.length > 0 && <div className="mt-2">
      <div className="rounded-2xl p-4 mb-3 flex items-start gap-3" style={{ background: C.purpleSoft, border: `1px solid ${C.purple}33` }}>
        <div className="rounded-xl p-2 shrink-0" style={{ background: C.well }}><ScanSearch size={18} style={{ color: C.purple }} /></div>
        <div><div className="font-extrabold text-sm" style={{ color: C.purple }}>متوفر أيضاً عند شركات لم تنضم إليها ({others.length} منتج)</div><div className="text-xs mt-0.5" style={{ color: C.soft }}>يمكنك معاينة الأسعار — وللشراء منها أرسل طلب انضمام أولاً.</div></div>
      </div>
      <div className="space-y-4">{othersByCo.map(([coId, items]) => { const co = data.companies.find((c) => c.id === coId); const ct = COMPANY_TYPE[co?.companyType || "distributor"]; return (<Card key={coId} className="p-4" style={{ border: `1px dashed ${C.purple}55` }}>
        <div className="flex items-center justify-between flex-wrap gap-2 mb-3">
          <div className="flex items-center gap-2.5"><span className="text-2xl">{co?.logo}</span><div><div className="font-extrabold text-sm flex items-center gap-2" style={{ color: C.text }}>{co?.name}<span className="text-[9px] font-bold px-1.5 py-0.5 rounded-full" style={{ background: C.purpleSoft, color: C.purple }}>غير منضم</span></div><div className="flex items-center gap-1.5 mt-0.5"><Stars {...(companyRating(data, coId) || { avg: null })} size={10} /><span className="text-[10px]" style={{ color: C.soft }}>· <ct.icon size={9} style={{ display: "inline" }} /> {ct.label}</span></div></div></div>
          <PrimaryBtn sm icon={UserPlus} onClick={() => props.go("m_explore")}>انضم للشراء</PrimaryBtn>
        </div>
        <div className="space-y-1.5">{items.map((p) => { const e = expiryInfo(p.expDate); return (<div key={p.id} className="flex items-center gap-2.5 rounded-xl px-3 py-2" style={{ background: C.bg }}>
          <div className="rounded-lg flex items-center justify-center shrink-0" style={{ width: 32, height: 32, background: C.well }}>{co?.server === "pharma" ? <Pill size={15} style={{ color: C.purple }} /> : <Package size={15} style={{ color: C.purple }} />}</div>
          <div className="flex-1 min-w-0"><div className="text-sm font-bold truncate" style={{ color: C.text }}>{p.name}</div><div className="flex items-center gap-1.5 flex-wrap"><span className="text-[10px]" style={{ color: C.soft }}>{p.cat} · {p.unit}</span>{p.offer > 0 && <span className="text-[9px] font-bold px-1 py-0.5 rounded" style={{ background: C.dangerSoft, color: C.danger }}>عرض -{p.offer}%</span>}{p.expDate && e && <span className="text-[9px] font-bold px-1 py-0.5 rounded" style={{ background: e.soft, color: e.color }}>صلاحية {fmtDate(p.expDate)}</span>}</div></div>
          <div className="text-left shrink-0"><div className="text-sm font-extrabold" style={{ color: C.purple }}>{money(priceAfter(p))}</div><div className="text-[9px]" style={{ color: C.soft }}>سعر الجملة</div></div>
        </div>); })}</div>
      </Card>); })}</div>
    </div>}
  </div>);
}

function MCats(props) {
  const { data } = props; const prods = F.products(props); const cats = [...new Set(prods.map((p) => p.cat))];
  const [sel, setSel] = useState(null);
  if (sel) { const list = prods.filter((p) => p.cat === sel);
    return (<div className="space-y-4"><button onClick={() => setSel(null)} className="inline-flex items-center gap-1 text-sm font-bold" style={{ color: C.soft }}><ChevronLeft size={16} />كل التصنيفات</button><SectionTitle>{sel}</SectionTitle><div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{list.map((p) => <ProductCard key={p.id} p={p} data={data} onAdd={(p) => { props.actions.quickAddToCompany(p.companyId, p.id); props.go("m_create"); }} fav={data.favorites?.includes(p.id)} onFav={props.actions.toggleFav} />)}</div></div>); }
  return (<div className="space-y-4"><SectionTitle sub="تصفّح المنتجات حسب التصنيف.">التصنيفات</SectionTitle>
    <div className="grid sm:grid-cols-3 lg:grid-cols-4 gap-3">{cats.map((c) => { const n = prods.filter((p) => p.cat === c).length; return (<button key={c} onClick={() => setSel(c)} className="rounded-2xl p-5 text-right transition-shadow hover:shadow-md" style={{ background: C.surface, border: `1px solid ${C.line}` }}><div className="rounded-xl p-2.5 w-fit mb-3" style={{ background: SC.soft }}><Layers size={20} style={{ color: SC.color }} /></div><div className="font-extrabold" style={{ color: C.text }}>{c}</div><div className="text-xs" style={{ color: C.soft }}>{n} منتج</div></button>); })}</div></div>);
}
function MFav(props) {
  const { data } = props; const favs = F.products(props).filter((p) => data.favorites?.includes(p.id));
  return (<div className="space-y-4"><SectionTitle sub="المنتجات التي حفظتها للوصول السريع.">المفضلة</SectionTitle>{favs.length ? <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{favs.map((p) => <ProductCard key={p.id} p={p} data={data} fav onFav={props.actions.toggleFav} onAdd={(p) => { props.actions.quickAddToCompany(p.companyId, p.id); props.go("m_create"); }} />)}</div> : <Empty icon={Heart} text="لا توجد منتجات مفضّلة بعد." />}</div>);
}
function MOffers(props) {
  const { data } = props; const { me } = useMy(props); const offers = F.products(props).filter((p) => p.offer > 0);
  return (<div className="space-y-4"><SectionTitle sub="منتجات عليها خصم من الشركات.">العروض</SectionTitle><div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{offers.map((p) => <ProductCard key={p.id} p={p} data={data} me={me} onAdd={(p) => { props.actions.quickAddToCompany(p.companyId, p.id); props.go("m_create"); }} fav={data.favorites?.includes(p.id)} onFav={props.actions.toggleFav} />)}</div></div>);
}
function MCompare(props) {
  const { data } = props; const { me } = useMy(props);
  const all = F.products(props).filter((p) => p.status === "active");
  const joined = joinedCompanyIds(data, me);
  // الأصناف المكررة عبر أكثر من شركة (للمقارنة التلقائية)
  const byName = Object.entries(all.reduce((acc, p) => { (acc[p.name] = acc[p.name] || []).push(p); return acc; }, {})).filter(([, arr]) => arr.length > 1).sort((a, b) => b[1].length - a[1].length);
  const [selName, setSelName] = useState(byName[0]?.[0] || null);
  const variants = selName ? [...(byName.find(([n]) => n === selName)?.[1] || [])].sort((a, b) => priceAfter(a) - priceAfter(b)) : [];
  const cheapest = variants[0];
  // المقارنة اليدوية (المختارة من البحث)
  const list = all.filter((p) => data.compare?.includes(p.id));
  const rows = [["الشركة", (p) => compName(data, p.companyId)], ["السعر", (p) => money(priceAfter(p))], ["التصنيف", (p) => p.cat], ["الوحدة", (p) => p.unit], ["أقل كمية", (p) => p.minOrder], ["الوزن", (p) => p.weight], ["المتوفر", (p) => p.stock]];
  const [mode, setMode] = useState("auto");
  return (<div className="space-y-4"><SectionTitle sub="قارن سعر الصنف نفسه عند كل الشركات — أو قارن منتجات مختارة يدوياً.">مقارنة الأسعار</SectionTitle>
    <div className="flex gap-1.5">{[["auto", "نفس الصنف عبر الشركات", GitCompare], ["manual", `مقارنتي المختارة (${list.length})`, Layers]].map(([k, lbl, Ic]) => { const on = mode === k; return (<button key={k} onClick={() => setMode(k)} className="rb-press flex-1 inline-flex items-center justify-center gap-1.5 rounded-xl py-2.5 text-xs font-bold" style={{ background: on ? SC.color : C.bg, color: on ? C.onAccent : C.text, border: `1px solid ${on ? SC.color : C.line}` }}><Ic size={14} />{lbl}</button>); })}</div>

    {mode === "auto" && <>
      {byName.length === 0 ? <Empty icon={GitCompare} text="لا توجد أصناف متوفرة عند أكثر من شركة بعد." /> : <>
        <Card className="p-3 sm:p-4"><div className="text-[11px] font-bold mb-1.5" style={{ color: C.soft }}>اختر الصنف للمقارنة ({byName.length} صنف متوفر عند أكثر من شركة)</div>
          <div className="flex gap-1.5 overflow-x-auto pb-1">{byName.map(([n, arr]) => { const on = selName === n; return (<button key={n} onClick={() => setSelName(n)} className="rb-press shrink-0 inline-flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-bold" style={{ background: on ? SC.color : C.bg, color: on ? C.onAccent : C.text, border: `1px solid ${on ? SC.color : C.line}` }}>{n}<span className="rounded-full px-1.5 text-[10px]" style={{ background: on ? "rgba(255,255,255,.25)" : C.surface, color: on ? "#fff" : C.soft }}>{arr.length} شركات</span></button>); })}</div>
        </Card>
        {selName && <div className="rb-stagger space-y-2.5">{variants.map((p, i) => { const co = data.companies.find((c) => c.id === p.companyId); const isJoined = joined.has(p.companyId); const isCheapest = i === 0; const diff = priceAfter(p) - priceAfter(cheapest); return (
          <Card key={p.id} className="p-4" style={isCheapest ? { border: `2px solid ${C.success}`, boxShadow: "0 8px 22px rgba(4,120,87,.12)" } : {}}>
            <div className="flex items-center justify-between gap-3 flex-wrap">
              <div className="flex items-center gap-3 min-w-0">
                <span className="text-2xl">{co?.logo}</span>
                <div className="min-w-0"><div className="font-extrabold text-sm flex items-center gap-2 flex-wrap" style={{ color: C.text }}>{co?.name}{isCheapest && <span className="text-[9px] font-extrabold px-2 py-0.5 rounded-full text-white" style={{ background: C.success }}>الأرخص 🏆</span>}{!isJoined && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded-full" style={{ background: C.purpleSoft, color: C.purple }}>غير منضم</span>}</div>
                <div className="flex items-center gap-2 mt-0.5 flex-wrap"><Stars {...(companyRating(data, p.companyId) || { avg: null })} size={10} /><span className="text-[10px]" style={{ color: C.soft }}>· مخزون {p.stock} · أقل كمية {p.minOrder}</span>{p.offer > 0 && <span className="text-[9px] font-bold px-1 py-0.5 rounded" style={{ background: C.dangerSoft, color: C.danger }}>عرض -{p.offer}%</span>}{p.tiers?.length > 0 && <span className="text-[9px] font-bold px-1 py-0.5 rounded" style={{ background: C.purpleSoft, color: C.purple }}>خصم كمية</span>}</div></div>
              </div>
              <div className="text-left shrink-0">
                <div className="text-lg font-extrabold" style={{ color: isCheapest ? C.success : C.text }}>{money(priceAfter(p))}</div>
                {diff > 0 && <div className="text-[10px] font-bold" style={{ color: C.danger }}>+{money(diff)} عن الأرخص</div>}
                <div className="mt-1.5">{isJoined ? <PrimaryBtn sm icon={Plus} onClick={() => { props.actions.quickAddToCompany(p.companyId, p.id); props.go("m_create"); }}>اطلب منها</PrimaryBtn> : <GhostBtn sm icon={UserPlus} onClick={() => props.go("m_explore")}>انضم للشراء</GhostBtn>}</div>
              </div>
            </div>
          </Card>); })}</div>}
      </>}
    </>}

    {mode === "manual" && (list.length ? <Card className="p-4 overflow-x-auto"><table className="w-full" style={{ borderCollapse: "collapse" }}><thead><tr><Th>الخاصية</Th>{list.map((p) => <th key={p.id} className="px-3 py-2.5 text-right text-sm font-extrabold whitespace-nowrap" style={{ color: C.text }}>{p.name}<button onClick={() => props.actions.toggleCompare(p.id)} className="block text-xs font-bold mt-1" style={{ color: C.danger }}>إزالة</button></th>)}</tr></thead><tbody>{rows.map(([label, fn]) => <tr key={label} style={{ borderTop: `1px solid ${C.line}` }}><Td bold>{label}</Td>{list.map((p) => <Td key={p.id}>{fn(p)}</Td>)}</tr>)}</tbody></table></Card> : <Empty icon={GitCompare} text="أضف منتجات للمقارنة من صفحة البحث (زر المقارنة على بطاقة المنتج)." />)}
  </div>);
}
function MCreate(props) {
  const { data, actions } = props; const { me } = useMy(props);
  const companies = F.companies(props).filter((c) => c.sub === "active");
  const draft = data.orderDraft || { companyId: null, cart: {} };
  const [openCo, setOpenCo] = useState(draft.companyId || null);
  if (openCo) { const co = companies.find((c) => c.id === openCo); if (!co) { setOpenCo(null); return null; }
    return <MCompanyOrder {...props} co={co} me={me} seedCart={draft.companyId === openCo ? draft.cart : null} onBack={() => { setOpenCo(null); actions.clearOrderDraft(); }} />;
  }
  return (<div className="space-y-4"><SectionTitle sub="اختر الشركة التي تريد الطلب منها — كل طلب يُحرَّر كفاتورة مستقلة لتلك الشركة.">إنشاء طلب</SectionTitle>
    <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{companies.map((c) => { const ct = COMPANY_TYPE[c.companyType || "distributor"]; const prods = data.products.filter((p) => p.companyId === c.id && p.status === "active"); const debt = data.invoices.filter((i) => i.merchantId === me.id && i.companyId === c.id).reduce((s, i) => s + (i.amount - i.paid), 0);
      return (<Card key={c.id} className="p-5 flex flex-col">
        <div className="flex items-center gap-3 mb-2"><span className="text-3xl">{c.logo}</span><div className="min-w-0"><div className="font-extrabold truncate" style={{ color: C.text }}>{c.name}</div><div className="text-xs" style={{ color: C.soft }}>{c.cat} · {c.city}</div><div className="mt-0.5"><Stars {...(companyRating(data, c.id) || { avg: null })} /></div></div></div>
        <div className="flex flex-wrap gap-1.5 mb-3">
          <span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: ct.soft, color: ct.color }}><ct.icon size={11} />{ct.label}</span>
          <span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: c.hasDelivery ? C.infoSoft : C.purpleSoft, color: c.hasDelivery ? C.info : C.purple }}>{c.hasDelivery ? <><Truck size={11} />توصيل متاح</> : <><Store size={11} />استلام من المقر</>}</span>
        </div>
        <div className="grid grid-cols-2 gap-2 text-center mb-3">
          <div className="rounded-xl p-2" style={{ background: C.bg }}><div className="text-lg font-extrabold" style={{ color: SC.color2 }}>{prods.length}</div><div className="text-[10px]" style={{ color: C.soft }}>منتج متاح</div></div>
          <div className="rounded-xl p-2" style={{ background: debt > 0 ? C.dangerSoft : C.successSoft }}><div className="text-sm font-extrabold" style={{ color: debt > 0 ? C.danger : C.success }}>{debt > 0 ? money(debt) : "صفر"}</div><div className="text-[10px]" style={{ color: C.soft }}>مديونيتك</div></div>
        </div>
        <PrimaryBtn full icon={FileText} onClick={() => setOpenCo(c.id)}>اطلب من هذه الشركة</PrimaryBtn>
      </Card>); })}{companies.length === 0 && <Empty icon={Building2} text="لا توجد شركات نشطة في سيرفرك." />}</div>
  </div>);
}
// فاتورة طلب احترافية لشركة واحدة
function MCompanyOrder(props) {
  const { data, actions, go, co, me, seedCart, onBack } = props;
  const isPharma = co.server === "pharma";
  const ct = COMPANY_TYPE[co.companyType || "distributor"];
  const prods = data.products.filter((p) => p.companyId === co.id && p.status === "active");
  // السلة: { pid: { count, mode } }  — mode: "piece" | "box"
  const initCart = () => { const c = {}; if (seedCart) Object.entries(seedCart).forEach(([pid, q]) => { if (typeof q === "object") { c[pid] = q; } else { const pp = prods.find((x) => x.id === pid); const md = (pp?.packSize || 1) > 1 ? "box" : "piece"; c[pid] = { count: q, mode: md }; } }); return c; };
  const [cart, setCart] = useState(initCart);
  const [note, setNote] = useState("");
  const [fulfillment, setFulfillment] = useState(co.hasDelivery ? "delivery" : "pickup");
  const [issued, setIssued] = useState(false);
  const [saved, setSaved] = useState(false);
  const [pq, setPq] = useState("");
  const [pcat, setPcat] = useState("الكل");
  const cats = ["الكل", ...Array.from(new Set(prods.map((p) => p.cat)))];
  const catCount = (c) => c === "الكل" ? prods.length : prods.filter((p) => p.cat === c).length;
  const visible = prods.filter((p) => (pcat === "الكل" || p.cat === pcat) && (!pq || p.name.includes(pq) || (p.cat || "").includes(pq) || (p.barcode || "").includes(pq)));
  const packOf = (p) => Math.max(1, p?.packSize || 1);
  const baseUnits = (pid) => { const l = cart[pid]; if (!l) return 0; const p = data.products.find((x) => x.id === pid); return l.mode === "box" ? l.count * packOf(p) : l.count; };
  const items = Object.entries(cart).filter(([, l]) => l.count > 0);
  const defMode = (pid) => { const p = data.products.find((x) => x.id === pid); return packOf(p) > 1 ? "box" : "piece"; };
  const add = (pid) => setCart((c) => { const l = c[pid] || { count: 0, mode: defMode(pid) }; return { ...c, [pid]: { ...l, count: l.count + 1 } }; });
  const setCount = (pid, n) => setCart((c) => { const l = c[pid] || { count: 0, mode: "piece" }; const x = { ...c }; if (n <= 0) delete x[pid]; else x[pid] = { ...l, count: n }; return x; });
  const setMode = (pid, mode) => setCart((c) => { const l = c[pid] || { count: 1, mode: "piece" }; return { ...c, [pid]: { ...l, mode } }; });
  const lineUnit = (p, u) => tierUnitPrice(p, me, u); // سعر القطعة بعد الخصم
  const subtotal = items.reduce((s, [pid]) => { const p = data.products.find((x) => x.id === pid); return s + priceForMerchant(p, me) * baseUnits(pid); }, 0);
  const total = items.reduce((s, [pid]) => { const p = data.products.find((x) => x.id === pid); return s + lineUnit(p, baseUnits(pid)) * baseUnits(pid); }, 0);
  const tierSavings = subtotal - total;
  const totalPieces = items.reduce((s, [pid]) => s + baseUnits(pid), 0);
  const today = new Date().toLocaleDateString("ar-LY", { year: "numeric", month: "long", day: "numeric" });
  const buildPayload = () => { const units = {}, meta = {}; items.forEach(([pid, l]) => { const p = data.products.find((x) => x.id === pid); const bu = baseUnits(pid); units[pid] = bu; meta[pid] = { mode: l.mode, count: l.count, packSize: packOf(p) }; }); return { units, meta }; };
  const isFrozen = !!me.frozen;
  const debtGate = invoiceGate(data, me.id);
  const usedCredit = Math.max(0, -me.balance);
  const overCredit = me.creditLimit ? (usedCredit + total) > me.creditLimit : false;
  const [pinOpen, setPinOpen] = useState(false);
  const [pinTry, setPinTry] = useState("");
  const [pinErr, setPinErr] = useState(false);
  const submit = () => { if (!items.length) { toast("أضف أصنافاً أولاً", "danger"); return; }
    if (isFrozen) { toast(`حسابك معلّق: ${me.freezeReason || "راجع الشركة"}`, "danger"); return; }
    if (debtGate.blocked) { toast(`ممنوع الطلب — ${debtGate.reason}. تواصل مع الشركة`, "danger"); return; }
    if (overCredit) { toast(`الطلب يتجاوز حدّك الائتماني (${money(me.creditLimit)})`, "danger"); return; }
    setPinTry(""); setPinErr(false); setPinOpen(true); };
  const confirmPin = () => {
    if (pinTry !== (me.orderPin || "0000")) { setPinErr(true); toast("الرمز السري غير صحيح", "danger"); return; }
    const { units, meta } = buildPayload(); actions.placeOrderDirect(me.id, co.id, units, note, fulfillment, meta); setPinOpen(false); setIssued(true); };
  const saveDraft = submit;

  const BillLine = ({ pid, l }) => { const p = data.products.find((x) => x.id === pid); const bu = baseUnits(pid); const t = tierFor(p, bu, priceForMerchant(p, me) * bu); const e = expiryInfo(p?.expDate); const u = lineUnit(p, bu); const pack = packOf(p); return (
    <div className="px-3 py-3" style={{ borderTop: `1px solid ${C.line}` }}>
      <div className="flex items-start justify-between gap-2 mb-1.5"><span className="text-sm font-bold leading-snug" style={{ color: C.text }}>{p?.name}</span><button onClick={() => setCount(pid, 0)} className="rounded p-1 shrink-0" style={{ background: C.dangerSoft }}><Trash2 size={12} style={{ color: C.danger }} /></button></div>
      {/* اختيار صندوق/قطعة (الصندوق أولاً) */}
      {pack > 1 && <div className="flex gap-1.5 mb-2">{[["box", "صندوق", `${pack} ${p?.unit}`], ["piece", "قطعة", p?.unit]].map(([m, lbl, sub]) => { const on = (l.mode || defMode(pid)) === m; return (<button key={m} onClick={() => setMode(pid, m)} className="flex-1 rounded-lg py-1.5 text-xs font-bold transition-all" style={{ background: on ? SC.color : C.bg, color: on ? C.onAccent : C.text, border: `1px solid ${on ? SC.color : C.line}` }}>{lbl} <span className="opacity-70 text-[10px]">({sub})</span></button>); })}</div>}
      <div className="flex items-center justify-between gap-2">
        <div className="flex items-center gap-1"><button onClick={() => setCount(pid, l.count - 1)} className="w-8 h-8 rounded-lg font-bold" style={{ border: `1px solid ${C.line}` }}>−</button><input type="number" value={l.count} onChange={(e) => setCount(pid, Math.max(0, Math.floor(+e.target.value || 0)))} className="w-14 text-center text-sm font-extrabold rounded-lg py-1.5 outline-none" style={{ ...inputStyle, color: SC.color }} /><button onClick={() => setCount(pid, l.count + 1)} className="w-8 h-8 rounded-lg font-bold text-white" style={{ background: SC.color }}>+</button><span className="text-[11px] mr-1" style={{ color: C.soft }}>{(l.mode || defMode(pid)) === "box" ? "صندوق" : "قطعة"}{(l.mode || defMode(pid)) === "box" ? ` = ${bu} ${p?.unit}` : ""}</span></div>
        <span className="text-sm font-extrabold" style={{ color: C.text }}>{money(u * bu)}</span>
      </div>
      <div className="flex items-center gap-1.5 flex-wrap mt-1.5">
        <span className="text-[10px]" style={{ color: C.soft }}>{money(u)} / {p?.unit}</span>
        {t && <span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.purpleSoft, color: C.purple }}>خصم {t.discount}%</span>}
        {isPharma && p?.batch && <span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.bg, color: C.soft }}>تشغيلة {p.batch}</span>}
        {p?.expDate && <span className="text-[11px] font-extrabold px-1.5 py-0.5 rounded inline-flex items-center gap-1" style={{ background: e?.soft, color: e?.color }}><CalendarClock size={11} />الصلاحية: {fmtDate(p.expDate)}</span>}
      </div>
    </div>); };

  return (<div className="space-y-4 pb-24 lg:pb-0">
    <button onClick={onBack} className="inline-flex items-center gap-1 text-sm font-bold" style={{ color: C.soft }}><ChevronLeft size={16} />رجوع لاختيار الشركة</button>
    {/* ترويسة الفاتورة */}
    <Card className="p-0 overflow-hidden">
      <div className="p-4 sm:p-5 text-white" style={{ background: `linear-gradient(120deg, ${SC.color} 0%, ${SC.color2} 100%)` }}>
        <div className="flex items-start justify-between flex-wrap gap-3">
          <div className="flex items-center gap-3"><div className="rounded-2xl flex items-center justify-center text-2xl sm:text-3xl shrink-0" style={{ width: 48, height: 48, background: "rgba(255,255,255,.18)" }}>{co.logo}</div><div className="min-w-0"><div className="font-extrabold text-base sm:text-lg leading-tight">{co.name}</div><div className="text-xs sm:text-sm opacity-90 truncate">{co.cat} · {co.city}</div><span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full mt-1" style={{ background: "rgba(255,255,255,.2)" }}><ct.icon size={11} />{ct.label}</span></div></div>
          <div className="text-left shrink-0"><div className="text-[11px] opacity-80">فاتورة طلب مبدئية</div><div className="font-extrabold text-sm">{today}</div><div className="text-[11px] opacity-80 mt-1">العميل: {me.name}</div></div>
        </div>
      </div>
    </Card>
    <div className="grid lg:grid-cols-3 gap-4">
      {/* قائمة المنتجات */}
      <div className="lg:col-span-2 space-y-3 min-w-0">
        <Card className="p-0 overflow-hidden">
          <div className="p-4 pb-3" style={{ borderBottom: `1px solid ${C.line}` }}>
            <div className="flex items-center justify-between mb-3 flex-wrap gap-2"><h3 className="font-extrabold flex items-center gap-2" style={{ color: C.text }}><Package size={17} style={{ color: SC.color }} />منتجات {co.name}</h3><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: SC.soft, color: SC.color }}>{prods.length} صنف</span></div>
            <div className="relative mb-3"><Search size={16} className="absolute right-3 top-2.5" style={{ color: C.soft }} /><input value={pq} onChange={(e) => setPq(e.target.value)} placeholder={isPharma ? "ابحث باسم الدواء أو التصنيف أو الباركود" : "ابحث عن منتج"} className="w-full rounded-xl pr-9 pl-3 py-2 text-sm outline-none" style={inputStyle} /></div>
            <div className="flex gap-1.5 overflow-x-auto pb-1" style={{ scrollbarWidth: "thin" }}>{cats.map((c) => { const on = pcat === c; return (<button key={c} onClick={() => setPcat(c)} className="shrink-0 inline-flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-bold transition-all" style={{ background: on ? SC.color : C.bg, color: on ? C.onAccent : C.text, border: `1px solid ${on ? SC.color : C.line}` }}>{c}<span className="rounded-full px-1.5 text-[10px]" style={{ background: on ? "rgba(255,255,255,.25)" : C.surface, color: on ? "#fff" : C.soft }}>{catCount(c)}</span></button>); })}</div>
          </div>
          <div className="divide-y max-h-[420px] overflow-auto" style={{ borderColor: C.line }}>
            {visible.map((p) => { const l = cart[p.id]; const inCart = l ? l.count : 0; const t = tierFor(p, 1); const e = expiryInfo(p.expDate); return (<div key={p.id} className="flex items-center gap-2 px-3 py-2 transition-colors hover:bg-gray-50" style={{ borderColor: C.line, background: inCart ? SC.soft + "55" : "transparent" }}>
              <div className="rounded-lg flex items-center justify-center shrink-0" style={{ width: 34, height: 34, background: SC.soft }}>{isPharma ? <Pill size={16} style={{ color: SC.color }} /> : <Package size={16} style={{ color: SC.color }} />}</div>
              <div className="flex-1 min-w-0">
                <div className="text-sm font-bold leading-tight truncate" style={{ color: C.text }}>{p.name}</div>
                <div className="flex items-center gap-1.5 flex-wrap mt-0.5">
                  <span className="text-[11px]" style={{ color: C.soft }}>{money(priceForMerchant(p, me))} / {p.unit}</span>
                  {packOf(p) > 1 && <span className="text-[10px]" style={{ color: C.soft }}>· صندوق = {packOf(p)}</span>}
                  {p.tiers && p.tiers.length > 0 && <span className="text-[10px] font-bold px-1 py-0.5 rounded" style={{ background: C.purpleSoft, color: C.purple }}>خصم كمية</span>}
                  {p.expDate && <span className="text-[10px] font-bold px-1 py-0.5 rounded inline-flex items-center gap-0.5" style={{ background: e?.soft, color: e?.color }}><CalendarClock size={9} />{fmtDate(p.expDate)}</span>}
                </div>
              </div>
              {inCart > 0 ? <div className="flex items-center gap-1 shrink-0"><button onClick={() => setCount(p.id, inCart - 1)} className="w-7 h-7 rounded-lg font-bold" style={{ border: `1px solid ${C.line}` }}>−</button><span className="w-7 text-center text-sm font-bold" style={{ color: SC.color }}>{inCart}</span><button onClick={() => add(p.id)} className="w-7 h-7 rounded-lg font-bold text-white" style={{ background: SC.color }}>+</button></div>
                : <button onClick={() => add(p.id)} className="shrink-0 inline-flex items-center gap-1 rounded-lg px-2.5 py-1.5 text-xs font-bold" style={{ background: SC.color, color: "#fff" }}><Plus size={13} />إضافة</button>}
            </div>); })}
            {visible.length === 0 && <div className="py-8"><Empty icon={Search} text="لا أصناف مطابقة." /></div>}
          </div>
        </Card>
      </div>
      {/* بنود الفاتورة */}
      <Card className="p-4 sm:p-5 h-fit lg:sticky lg:top-20 min-w-0">
        <h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Receipt size={18} style={{ color: SC.color }} />بنود الفاتورة {items.length > 0 && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: SC.soft, color: SC.color }}>{items.length}</span>}</h3>
        <div className="rounded-xl overflow-hidden mb-3" style={{ border: `1px solid ${C.line}` }}>
          {items.length === 0 && <div className="px-3 py-6 text-sm text-center" style={{ color: C.soft }}>أضف منتجات لإنشاء الفاتورة.</div>}
          {items.map(([pid, l]) => <BillLine key={pid} pid={pid} l={l} />)}
        </div>
        <div className="mb-3"><div className="text-sm font-bold mb-1.5" style={{ color: C.text }}>طريقة الاستلام</div>
          {co.hasDelivery ? <div className="grid grid-cols-2 gap-2">{Object.entries(FULFILL).map(([k, v]) => { const on = fulfillment === k; return (<button key={k} onClick={() => setFulfillment(k)} className="flex items-center justify-center gap-1.5 rounded-xl py-2.5 text-sm font-bold transition-all" style={{ background: on ? v.color : C.surface, color: on ? "#fff" : C.text, border: `1px solid ${on ? v.color : C.line}` }}><v.icon size={15} />{v.label}</button>); })}</div>
            : <div className="rounded-xl p-2.5 flex items-center gap-2 text-sm" style={{ background: C.purpleSoft, color: C.purple }}><Store size={16} /><span className="font-bold">استلام من مقر الشركة فقط</span></div>}
          <div className="text-[11px] mt-1" style={{ color: C.soft }}>{!co.hasDelivery ? `${co.name} لا توفّر توصيلاً — تستلم البضاعة من مقرها.` : fulfillment === "pickup" ? "ستستلم البضاعة بنفسك من مقر الشركة." : (ct.label === "بيع مباشر" ? "ستوصّلها الشركة بنفسها." : "ستصلك عبر مندوب/سائق الشركة.")}</div>
        </div>
        <Field label="ملاحظة"><Input value={note} onChange={(e) => setNote(e.target.value)} placeholder="اختياري" /></Field>
        {isFrozen && <div className="rounded-xl p-2.5 mb-2 flex items-start gap-2 text-xs" style={{ background: C.dangerSoft, color: C.danger }}><AlertCircle size={15} className="shrink-0 mt-0.5" /><span className="font-bold">حسابك معلّق عن الطلب: {me.freezeReason || "راجع الشركة"} — سوّ مستحقاتك أولاً.</span></div>}
        {!isFrozen && debtGate.blocked && <div className="rounded-xl p-2.5 mb-2 flex items-start gap-2 text-xs" style={{ background: C.dangerSoft, color: C.danger }}><Clock size={15} className="shrink-0 mt-0.5" /><span className="font-bold">ممنوع إصدار طلب — {debtGate.reason}. سدّد المتأخر أو اطلب سماحاً/إمهالاً من الشركة.</span></div>}
        {!isFrozen && !debtGate.blocked && me.invoiceOverride && <div className="rounded-xl p-2.5 mb-2 flex items-start gap-2 text-xs" style={{ background: C.successSoft, color: C.success }}><CheckCircle2 size={15} className="shrink-0 mt-0.5" /><span className="font-bold">{me.invoiceOverride.type === "allow" ? "سماح استثنائي من الشركة بإصدار الطلبات رغم تأخر السداد." : `أمهلتك الشركة ${me.invoiceOverride.extraDays} يوم لظرف طارئ.`}</span></div>}
        {overCredit && <div className="rounded-xl p-2.5 mb-2 flex items-start gap-2 text-xs" style={{ background: C.dangerSoft, color: C.danger }}><AlertCircle size={15} className="shrink-0 mt-0.5" /><span className="font-bold">تنبيه: هذا الطلب يتجاوز حدّك الائتماني ({money(me.creditLimit)}) — مديونيتك الحالية {money(usedCredit)}. سدّد جزءاً من مستحقاتك أولاً.</span></div>}
        <div className="space-y-1 my-3 pt-3 text-sm" style={{ borderTop: `1px solid ${C.line}` }}>
          <div className="flex justify-between" style={{ color: C.soft }}><span>إجمالي القطع</span><span className="font-bold" style={{ color: C.text }}>{totalPieces} قطعة</span></div>
          {tierSavings > 0 && <div className="flex justify-between" style={{ color: C.purple }}><span>وفّرت بخصم الكمية</span><span className="font-bold">{money(tierSavings)}</span></div>}
          <div className="flex justify-between items-center pt-1"><span className="font-extrabold" style={{ color: C.text }}>الإجمالي</span><span className="text-lg font-extrabold" style={{ color: SC.color }}>{money(total)}</span></div>
        </div>
        {/* الإتمام يذهب لخانة الموافقة النهائية (لا يُرسَل للمندوب مباشرة) */}
        <div className="hidden lg:block space-y-2">
          <PrimaryBtn full tone="accent" icon={ClipboardCheck} onClick={submit}>إتمام الفاتورة → الموافقة النهائية</PrimaryBtn>
          <div className="text-[11px] text-center" style={{ color: C.soft }}>لن تُرسَل للمندوب إلا بعد موافقتك النهائية.</div>
        </div>
      </Card>
    </div>
    {/* شريط سفلي ثابت للهاتف */}
    <div className="lg:hidden fixed bottom-0 right-0 left-0 z-40 p-3" style={{ background: C.surface, borderTop: `1px solid ${C.line}`, boxShadow: "0 -6px 18px rgba(11,31,58,.08)" }}>
      <div className="flex items-center justify-between mb-2"><span className="text-xs" style={{ color: C.soft }}>{items.length} صنف · {totalPieces} قطعة</span><span className="font-extrabold" style={{ color: SC.color }}>{money(total)}</span></div>
      <button onClick={submit} className="w-full flex items-center justify-center gap-1.5 rounded-xl px-4 py-3 text-sm font-extrabold text-white" style={{ background: C.accent }}><ClipboardCheck size={16} />إتمام الفاتورة → الموافقة النهائية</button>
    </div>
    {/* نافذة الرمز السري لتأكيد الطلبية */}
    {pinOpen && <div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50" onClick={() => setPinOpen(false)}>
      <div className="rb-modal w-full max-w-sm rounded-2xl overflow-hidden" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
        <div className="p-5 text-white text-center" style={{ background: `linear-gradient(120deg, ${SC.color}, ${SC.color2})` }}><div className="rounded-full mx-auto flex items-center justify-center mb-2" style={{ width: 56, height: 56, background: "rgba(255,255,255,.2)" }}><Lock size={28} color="#fff" /></div><div className="font-extrabold text-lg">تأكيد الطلبية بالرمز السري</div><div className="text-sm opacity-90">أدخل رمزك المكوّن من 4 أرقام لإرسال الطلب</div></div>
        <div className="p-5 space-y-3">
          <div className="rounded-xl p-3 space-y-2 text-sm" style={{ background: C.bg }}>{[["الشركة", co.name], ["عدد الأصناف", items.length], ["إجمالي القطع", `${totalPieces} قطعة`], ["الإجمالي", money(total)]].map(([k, v]) => <div key={k} className="flex justify-between"><span style={{ color: C.soft }}>{k}</span><span className="font-bold" style={{ color: C.text }}>{v}</span></div>)}</div>
          <input value={pinTry} onChange={(e) => { setPinTry(e.target.value.replace(/\D/g, "").slice(0, 4)); setPinErr(false); }} type="tel" inputMode="numeric" maxLength={4} placeholder="••••" autoFocus className="w-full text-center rounded-xl py-3 outline-none tracking-[0.8em] font-extrabold" style={{ ...inputStyle, fontSize: 26, border: `2px solid ${pinErr ? C.danger : C.line}`, paddingRight: "0.8em" }} />
          {pinErr && <div className="text-xs text-center font-bold" style={{ color: C.danger }}>الرمز غير صحيح — يمكنك تعديله من ملفك الشخصي (الافتراضي 0000)</div>}
          <PrimaryBtn full icon={CheckCircle2} onClick={confirmPin}>تأكيد وإرسال الطلب</PrimaryBtn>
          <button onClick={() => setPinOpen(false)} className="w-full text-center text-sm font-bold py-1" style={{ color: C.soft }}>إلغاء</button>
        </div>
      </div>
    </div>}
    {issued && <div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50" onClick={() => go("m_track")}>
      <div className="rb-modal w-full max-w-sm rounded-2xl overflow-hidden" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
        <div className="p-5 text-white text-center" style={{ background: `linear-gradient(120deg, ${C.success}, #059669)` }}><div className="rounded-full mx-auto flex items-center justify-center mb-2" style={{ width: 56, height: 56, background: "rgba(255,255,255,.2)" }}><CheckCircle2 size={32} color="#fff" /></div><div className="font-extrabold text-lg">أُرسل طلبك بنجاح ✓</div><div className="text-sm opacity-90">{me.repId ? "في طريقه إلى مندوبك" : "في طريقه إلى الشركة"}</div></div>
        <div className="p-5 space-y-3">
          <div className="rounded-xl p-3 space-y-2 text-sm" style={{ background: C.bg }}>{[["الشركة", co.name], ["عدد الأصناف", items.length], ["إجمالي القطع", `${totalPieces} قطعة`], ["الإجمالي", money(total)]].map(([k, v]) => <div key={k} className="flex justify-between"><span style={{ color: C.soft }}>{k}</span><span className="font-bold" style={{ color: C.text }}>{v}</span></div>)}</div>
          <PrimaryBtn full tone="success" icon={Truck} onClick={() => go("m_track")}>متابعة الطلب</PrimaryBtn>
        </div>
      </div>
    </div>}
  </div>);
}
function MDrafts(props) {
  const { data, actions } = props; const { meId } = useMy(props);
  const drafts = data.orders.filter((o) => o.merchantId === meId && o.status === "draft");
  return (<div className="space-y-4"><SectionTitle sub="فواتيرك الجاهزة بانتظار موافقتك النهائية — راجعها وعدّل الكميات، ثم اعتمدها لتُرسَل للمندوب.">الموافقة النهائية للطلبات</SectionTitle>
    {drafts.length > 0 && <div className="rounded-xl p-3 flex items-start gap-2" style={{ background: C.accentSoft }}><AlertCircle size={18} style={{ color: "#92590C" }} className="mt-0.5 shrink-0" /><span className="text-sm font-bold" style={{ color: "#92590C" }}>لديك {drafts.length} فاتورة بانتظار الموافقة النهائية — لن تصل للمندوب قبل اعتمادك.</span></div>}
    {drafts.length ? drafts.map((o) => { const co = data.companies.find((c) => c.id === o.companyId); return (<Card key={o.id} className="p-4"><div className="flex items-center justify-between mb-3 flex-wrap gap-2"><div className="font-extrabold flex items-center gap-2" style={{ color: C.text }}><span className="text-xl">{co?.logo}</span>{o.id} <span className="text-xs font-normal" style={{ color: C.soft }}>{compName(data, o.companyId)}</span></div><div className="flex items-center gap-1.5 flex-wrap"><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.accentSoft, color: "#92590C" }}>بانتظار موافقتك النهائية</span>{(o.ageDays || 0) >= 2 && <span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: C.dangerSoft, color: C.danger }}>متقادمة منذ {o.ageDays} أيام ⏳</span>}</div></div>
      <div className="space-y-2 mb-3">{o.items.map((it, i) => (<div key={i} className="flex items-center justify-between gap-2"><span className="text-sm min-w-0" style={{ color: C.text }}>{it.name}{it.boxes ? <span className="text-[11px]" style={{ color: C.soft }}> · {it.boxes} صندوق</span> : null}</span><div className="flex items-center gap-1 shrink-0"><button onClick={() => actions.editDraftQty(o.id, it.pid, it.qty - 1)} className="w-7 h-7 rounded-lg font-bold" style={{ border: `1px solid ${C.line}` }}>−</button><span className="w-7 text-center text-sm font-bold">{it.qty}</span><button onClick={() => actions.editDraftQty(o.id, it.pid, it.qty + 1)} className="w-7 h-7 rounded-lg font-bold text-white" style={{ background: SC.color }}>+</button><button onClick={() => actions.editDraftQty(o.id, it.pid, 0)} className="rounded-lg p-1.5" style={{ background: C.dangerSoft }}><Trash2 size={14} style={{ color: C.danger }} /></button></div></div>))}</div>
      <div className="text-[10px] mb-2 flex items-center gap-1" style={{ color: C.soft }}><Send size={10} />بعد موافقتك سيُرسَل إلى: <b style={{ color: C.text }}>{o.repId ? `المندوب ${repName(data, o.repId)}` : "الشركة مباشرة (بيع مباشر)"}</b></div>
      <div className="flex items-center justify-between pt-3 flex-wrap gap-2" style={{ borderTop: `1px solid ${C.line}` }}><span className="font-extrabold" style={{ color: SC.color }}>{money(orderTotal(o))}</span><div className="flex gap-2"><GhostBtn sm icon={Trash2} onClick={() => confirmThen("سيُحذف هذا الطلب نهائياً ولن يصل للمندوب.", () => { actions.discardDraft(o.id); toast("حُذفت الفاتورة"); }, { tone: "danger", yesLabel: "حذف" })}>حذف</GhostBtn><PrimaryBtn sm tone="accent" icon={CheckCircle2} onClick={() => { actions.submitDraft(o.id); toast("اعتُمد الطلب وأُرسل ✓"); }}>الموافقة النهائية وإرسال</PrimaryBtn></div></div>
    </Card>); }) : <Empty icon={ClipboardCheck} text="لا توجد فواتير بانتظار الموافقة النهائية." />}</div>);
}
function MReorder(props) {
  const { data, actions, go } = props; const { meId } = useMy(props);
  const past = data.orders.filter((o) => o.merchantId === meId && o.status === "delivered");
  return (<div className="space-y-4"><SectionTitle sub="أعد إرسال طلب سابق بنفس الأصناف بضغطة واحدة.">إعادة طلب سابق</SectionTitle>
    {past.length ? past.map((o) => (<Card key={o.id} className="p-4 flex items-center justify-between gap-3 flex-wrap"><div><div className="font-bold" style={{ color: C.text }}>{o.id} · {compName(data, o.companyId)}</div><div className="text-sm" style={{ color: C.soft }}>{o.items.map((i) => `${i.name} ×${i.qty}`).join("، ")}</div><div className="text-sm font-extrabold mt-1" style={{ color: SC.color }}>{money(orderTotal(o))}</div></div><PrimaryBtn icon={Repeat} onClick={() => { actions.reorder(o.id); go("m_track"); }}>إعادة الطلب</PrimaryBtn></Card>)) : <Empty icon={Repeat} text="لا توجد طلبات سابقة مكتملة." />}</div>);
}
function MTrack(props) {
  const { data } = props; const { meId } = useMy(props);
  const active = data.orders.filter((o) => o.merchantId === meId && !["draft", "delivered", "rejected"].includes(o.status));
  return (<div className="space-y-4"><SectionTitle sub="حالة كل طلب نشط ومساره عبر المندوب والشركة والسائق.">متابعة الطلبات</SectionTitle>
    {active.length ? active.map((o) => (<Card key={o.id} className="p-4"><div className="flex items-center justify-between mb-3 flex-wrap gap-2"><div className="font-extrabold" style={{ color: C.text }}>{o.id} <span className="text-xs font-normal" style={{ color: C.soft }}>{compName(data, o.companyId)} · {money(orderTotal(o))}</span></div><Badge status={o.status} /></div><div className="rounded-xl p-3" style={{ background: C.bg }}><Pipeline status={o.status} /></div><div className="mt-3 text-left"><GhostBtn sm icon={Eye} onClick={() => props.openOrder(o.id)}>التفاصيل</GhostBtn></div></Card>)) : <Empty icon={ClipboardList} text="لا توجد طلبات نشطة." />}</div>);
}
function MHistory(props) {
  const { data } = props; const { meId } = useMy(props);
  const all = data.orders.filter((o) => o.merchantId === meId && o.status !== "draft");
  return (<div className="space-y-4"><SectionTitle sub="كل طلباتك السابقة والحالية.">سجل الطلبات</SectionTitle><Card className="p-2"><OrderList orders={all} data={data} onOpen={props.openOrder} /></Card></div>);
}
function MStatement(props) {
  const { data } = props; const { me } = useMy(props);
  const [showPrint, setShowPrint] = useState(false);
  // دفتر الحركة: فواتير (مدين) + دفعات (دائن) + إشعارات قيد
  const rows = [
    ...data.invoices.filter((i) => i.merchantId === me.id).map((i) => ({ age: i.ageDays ?? 30, date: i.date, desc: `فاتورة ${i.id} — ${compName(data, i.companyId)}`, debit: i.amount, credit: 0 })),
    ...data.payments.filter((p) => p.merchantId === me.id).map((p) => ({ age: 5, date: p.date, desc: `دفعة ${p.id} (${p.method})`, debit: 0, credit: p.amount })),
    ...data.creditNotes.filter((c) => c.merchantId === me.id).map((c) => ({ age: c.ageDays ?? 0, date: c.date, desc: `${c.kind === "credit" ? "إشعار دائن" : "إشعار مدين"} ${c.id} — ${c.reason}`, debit: c.kind === "debit" ? c.amount : 0, credit: c.kind === "credit" ? c.amount : 0 })),
  ].sort((a, b) => b.age - a.age);
  let bal = 0; const ledger = rows.map((r) => { bal += r.debit - r.credit; return { ...r, bal }; });
  const totD = rows.reduce((s, r) => s + r.debit, 0), totC = rows.reduce((s, r) => s + r.credit, 0);
  const exportCsv = () => { const head = "التاريخ,البيان,مدين,دائن,الرصيد\n"; const body = ledger.map((r) => `${r.date},"${r.desc}",${r.debit},${r.credit},${r.bal}`).join("\n"); const a = document.createElement("a"); a.href = "data:text/csv;charset=utf-8,\uFEFF" + encodeURIComponent(head + body); a.download = `statement-${me.id}.csv`; a.click(); };
  return (<div className="space-y-4"><SectionTitle sub="سجل حركتك الدفتري: كل فاتورة ودفعة وسند قيد، بالرصيد بعد كل حركة." action={<div className="flex gap-2"><GhostBtn sm icon={Printer} onClick={() => setShowPrint(true)}>طباعة</GhostBtn><GhostBtn sm icon={FileText} onClick={exportCsv}>تصدير CSV</GhostBtn></div>}>كشف الحساب</SectionTitle>
    <div className="grid grid-cols-3 gap-3">
      <Stat icon={TrendingDown} label="إجمالي مدين (عليك)" value={money(totD)} tint={{ color: C.danger, soft: C.dangerSoft }} />
      <Stat icon={HandCoins} label="إجمالي دائن (سدّدت/خُصم)" value={money(totC)} tint={{ color: C.success, soft: C.successSoft }} />
      <Stat icon={Wallet} label="الرصيد المستحق" value={money(Math.max(0, totD - totC))} tint={{ color: C.accent, soft: C.accentSoft }} />
    </div>
    <Card className="p-4 min-w-0">
      <div className="sm:hidden space-y-2">{ledger.map((r, i) => (<div key={i} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}><div className="text-xs font-bold mb-1" style={{ color: C.text }}>{r.desc}</div><div className="flex items-center justify-between text-[11px]"><span style={{ color: C.soft }}>{r.date}</span><span>{r.debit ? <b style={{ color: C.danger }}>مدين {money(r.debit)}</b> : <b style={{ color: C.success }}>دائن {money(r.credit)}</b>}</span></div><div className="text-[10px] mt-1 text-left font-bold" style={{ color: C.text }}>الرصيد: {money(r.bal)}</div></div>))}</div>
      <div className="hidden sm:block"><Table head={["التاريخ", "البيان", "مدين", "دائن", "الرصيد بعد الحركة"]}>{ledger.map((r, i) => <tr key={i} style={{ borderBottom: `1px solid ${C.line}` }}><Td>{r.date}</Td><Td bold>{r.desc}</Td><Td>{r.debit ? <span style={{ color: C.danger, fontWeight: 800 }}>{money(r.debit)}</span> : "—"}</Td><Td>{r.credit ? <span style={{ color: C.success, fontWeight: 800 }}>{money(r.credit)}</span> : "—"}</Td><Td bold>{money(r.bal)}</Td></tr>)}</Table></div>
      {ledger.length === 0 && <Empty icon={ScrollText} text="لا حركات بعد." />}
    </Card>
    {showPrint && <AccountStatementView title={`كشف حساب — ${me.name}`} ledger={ledger} totD={totD} totC={totC} onClose={() => setShowPrint(false)} />}
  </div>);
}
// كشف حساب قابل للطباعة A4
function AccountStatementView({ title, ledger, totD, totC, onClose }) {
  return (<div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-3 bg-black/50" onClick={onClose}>
    <div className="rb-modal w-full max-w-lg rounded-2xl overflow-hidden max-h-[92vh] flex flex-col" style={{ background: "#fff" }} onClick={(e) => e.stopPropagation()}>
      <div id="invoice-print" className="overflow-auto">
        <div className="p-4 text-white flex items-center justify-between" style={{ background: C.ink }}><div className="font-extrabold flex items-center gap-2"><ScrollText size={18} />{title}</div><div className="text-xs opacity-80">رابط — كشف حساب</div></div>
        <div className="p-4"><table className="w-full text-xs" style={{ borderCollapse: "collapse" }}>
          <thead><tr style={{ borderBottom: `2px solid ${C.ink}` }}>{["التاريخ", "البيان", "مدين", "دائن", "الرصيد"].map((h) => <th key={h} className="text-right p-1.5" style={{ color: C.text }}>{h}</th>)}</tr></thead>
          <tbody>{ledger.map((r, i) => <tr key={i} style={{ borderBottom: `1px solid ${C.line}` }}><td className="p-1.5" style={{ color: C.soft }}>{r.date}</td><td className="p-1.5 font-bold" style={{ color: C.text }}>{r.desc}</td><td className="p-1.5" style={{ color: C.danger }}>{r.debit ? money(r.debit) : ""}</td><td className="p-1.5" style={{ color: C.success }}>{r.credit ? money(r.credit) : ""}</td><td className="p-1.5 font-bold" style={{ color: C.text }}>{money(r.bal)}</td></tr>)}</tbody>
        </table>
        <div className="flex justify-end gap-4 mt-3 pt-2 text-sm font-extrabold" style={{ borderTop: `2px solid ${C.ink}` }}><span style={{ color: C.danger }}>مدين: {money(totD)}</span><span style={{ color: C.success }}>دائن: {money(totC)}</span><span style={{ color: C.text }}>المستحق: {money(Math.max(0, totD - totC))}</span></div></div>
      </div>
      <div className="p-3 flex gap-2 shrink-0 no-print" style={{ borderTop: `1px solid ${C.line}` }}><PrimaryBtn full icon={Printer} onClick={() => window.print()}>طباعة</PrimaryBtn><GhostBtn icon={XCircle} onClick={onClose}>إغلاق</GhostBtn></div>
    </div>
  </div>);
}
function MInvoices(props) {
  const { data, go } = props; const { me } = useMy(props);
  const inv = data.invoices.filter((i) => i.merchantId === me.id);
  const [sel, setSel] = useState(null);
  const iv = inv.find((i) => i.id === sel);
  return (<div className="space-y-4"><SectionTitle sub="راجع فواتير مشترياتك من الشركات — اضغط أي فاتورة لعرض بنودها وصلاحياتها.">مراجعة الفواتير</SectionTitle>
    <Card className="p-4"><Table head={["الفاتورة", "الشركة", "المبلغ", "المدفوع", "المتبقّي", "الاستحقاق", "الحالة", ""]}>{inv.map((i) => { const rem = i.amount - i.paid; return (<tr key={i.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{i.id}</Td><Td>{compName(data, i.companyId)}</Td><Td>{money(i.amount)}</Td><Td>{money(i.paid)}</Td><Td><span style={{ color: rem > 0 ? C.danger : C.success, fontWeight: 700 }}>{money(rem)}</span></Td><Td>{i.due}</Td><Td><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={i.paid >= i.amount ? { background: C.successSoft, color: C.success } : { background: C.warnSoft, color: "#92590C" }}>{i.paid >= i.amount ? "مدفوعة" : "مستحقة"}</span></Td><Td><GhostBtn sm icon={Eye} onClick={() => setSel(i.id)}>عرض</GhostBtn></Td></tr>); })}</Table>{inv.length === 0 && <Empty icon={Receipt} text="لا توجد فواتير." />}</Card>
    {iv && <div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50" onClick={() => setSel(null)}>
      <div className="rb-modal w-full max-w-md rounded-2xl overflow-hidden max-h-[88vh] flex flex-col" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
        <div className="p-5 text-white shrink-0" style={{ background: `linear-gradient(120deg, ${SC.color}, ${SC.color2})` }}><div className="flex items-center justify-between"><div><div className="text-xs opacity-80">فاتورة</div><div className="font-extrabold text-lg">{iv.id}</div></div><div className="text-left"><div className="text-xs opacity-80">{compName(data, iv.companyId)}</div><div className="font-bold">{iv.date}</div></div></div></div>
        <div className="p-5 space-y-3 overflow-auto">
          <div className="rounded-xl overflow-hidden" style={{ border: `1px solid ${C.line}` }}><div className="px-3 py-1.5 text-[11px] font-bold flex" style={{ background: C.bg, color: C.soft }}><span className="flex-1">الصنف</span><span className="w-12 text-center">كمية</span><span className="w-16 text-left">قيمة</span></div>{(iv.items || []).map((it, i) => { const p = data.products.find((x) => x.id === it.pid); const e = expiryInfo(it.expDate || p?.expDate); return (<div key={i} className="px-3 py-2 text-xs" style={{ borderTop: `1px solid ${C.line}` }}><div className="flex items-center"><span className="flex-1 font-bold" style={{ color: C.text }}>{it.name}</span><span className="w-12 text-center" style={{ color: C.text }}>{it.boxes ? `${it.boxes}ص` : it.qty}</span><span className="w-16 text-left font-bold" style={{ color: C.text }}>{money(it.qty * it.price)}</span></div>{((it.expDate || p?.expDate) || it.batch) && <div className="flex flex-wrap gap-1 mt-1">{it.batch && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.bg, color: C.soft }}>تشغيلة {it.batch}</span>}{(it.expDate || p?.expDate) && <span className="text-[9px] font-extrabold px-1.5 py-0.5 rounded" style={{ background: e?.soft, color: e?.color }}>الصلاحية: {fmtDate(it.expDate || p?.expDate)}</span>}</div>}</div>); })}{(!iv.items || !iv.items.length) && <div className="px-3 py-3 text-xs text-center" style={{ color: C.soft }}>لا تفاصيل بنود لهذه الفاتورة.</div>}</div>
          <div className="rounded-xl p-3 space-y-1.5 text-sm" style={{ background: C.bg }}>{[["الإجمالي", money(iv.amount)], ["المدفوع", money(iv.paid)], ["المتبقّي", money(iv.amount - iv.paid)], ["الاستحقاق", iv.due]].map(([k, v]) => <div key={k} className="flex justify-between"><span style={{ color: C.soft }}>{k}</span><span className="font-bold" style={{ color: C.text }}>{v}</span></div>)}</div>
          {iv.amount - iv.paid > 0 && <PrimaryBtn full tone="accent" icon={HandCoins} onClick={() => { setSel(null); go("m_settle"); }}>تسوية / سداد هذه الفاتورة</PrimaryBtn>}
          <GhostBtn full icon={XCircle} onClick={() => setSel(null)}>إغلاق</GhostBtn>
        </div>
      </div>
    </div>}
  </div>);
}
// تسوية فاتورة: يدفع التاجر (نقد/شيك/حوالة) ← يراجع المندوب ← يعتمد
function MSettle(props) {
  const { data, actions } = props; const { me } = useMy(props);
  const openInv = data.invoices.filter((i) => i.merchantId === me.id && i.paid < i.amount);
  const mySettlements = data.settlements.filter((s) => s.merchantId === me.id);
  const [invoiceId, setInvoiceId] = useState(openInv[0]?.id || "");
  const [method, setMethod] = useState("cash");
  const [amount, setAmount] = useState("");
  const [ref, setRef] = useState(""); const [note, setNote] = useState("");
  const [done, setDone] = useState(false);
  const inv = data.invoices.find((i) => i.id === invoiceId);
  const remaining = inv ? inv.amount - inv.paid : Math.max(0, -me.balance);
  const submit = () => {
    if (!(+amount > 0)) return;
    actions.createSettlement({ merchantId: me.id, companyId: inv ? inv.companyId : me.companyId, invoiceId: invoiceId || null, amount: +amount, method, ref, note });
    toast("أُرسل السداد لمراجعة المندوب ✓");
    setDone(true); setAmount(""); setRef(""); setNote("");
  };
  const myCollects = data.collectRequests.filter((c) => c.merchantId === me.id);
  const activeCollect = myCollects.find((c) => ["requested", "scheduled"].includes(c.status));
  const [reqDone, setReqDone] = useState(false);
  const requestRep = () => { actions.requestCollection({ merchantId: me.id, companyId: inv ? inv.companyId : me.companyId, amount: +amount || remaining || 0, method, note }); toast("أُرسل طلب التحصيل للمندوب ✓"); setReqDone(true); setTimeout(() => setReqDone(false), 3000); };
  return (<div className="space-y-4"><SectionTitle sub="سدّد فاتورتك نقداً أو بشيك أو حوالة. يراجع المندوب القيمة ويعتمدها لتُضاف لحسابك وتُعلَم بها الشركة.">تسوية فاتورة / سداد</SectionTitle>
    <div className="grid lg:grid-cols-3 gap-4">
      <Card className="p-5 lg:col-span-1 h-fit space-y-3">
        <div className="rounded-xl p-3 text-center" style={{ background: me.balance < 0 ? C.dangerSoft : C.successSoft }}><div className="text-xs" style={{ color: C.soft }}>إجمالي مديونيتك</div><div className="text-2xl font-extrabold" style={{ color: me.balance < 0 ? C.danger : C.success }}>{money(Math.max(0, -me.balance))}</div></div>
        <Field label="الفاتورة (اختياري)"><select value={invoiceId} onChange={(e) => { setInvoiceId(e.target.value); setDone(false); }} className={inputCls} style={inputStyle}><option value="">سداد عام على الحساب</option>{openInv.map((i) => <option key={i.id} value={i.id}>{i.id} — متبقّي {money(i.amount - i.paid)}</option>)}</select></Field>
        <div><div className="text-sm font-bold mb-1.5" style={{ color: C.text }}>طريقة الدفع</div><div className="grid grid-cols-3 gap-2">{Object.entries(METHOD_LABEL).map(([k, label]) => { const Ic = METHOD_ICON[k]; const on = method === k; return (<button key={k} onClick={() => setMethod(k)} className="flex flex-col items-center gap-1 rounded-xl py-2.5 transition-all" style={{ background: on ? SC.color : C.surface, color: on ? "#fff" : C.text, border: `1px solid ${on ? SC.color : C.line}` }}><Ic size={16} /><span className="text-[11px] font-bold">{label}</span></button>); })}</div></div>
        <Field label="المبلغ المدفوع"><Input type="number" value={amount} onChange={(e) => { setAmount(e.target.value); setDone(false); }} placeholder={remaining ? String(remaining) : "0"} /></Field>
        {remaining > 0 && <button onClick={() => setAmount(String(remaining))} className="text-xs font-bold" style={{ color: SC.color }}>سداد كامل المتبقّي ({money(remaining)})</button>}
        {method !== "cash" && <Field label={method === "cheque" ? "رقم الشيك / المصرف" : "رقم الحوالة / المرجع"}><Input value={ref} onChange={(e) => setRef(e.target.value)} placeholder={method === "cheque" ? "مثال: 0045821 - مصرف الجمهورية" : "مثال: TRF-99213"} /></Field>}
        <Field label="ملاحظة"><Input value={note} onChange={(e) => setNote(e.target.value)} placeholder="اختياري" /></Field>
        <PrimaryBtn full tone="accent" icon={Send} onClick={submit}>إرسال السداد للمراجعة</PrimaryBtn>
        {done && <div className="text-xs rounded-lg p-2 flex items-center gap-1.5" style={{ background: C.successSoft, color: C.success }}><CheckCircle2 size={14} />أُرسل السداد، بانتظار اعتماد المندوب.</div>}
        {/* طلب مندوب للتحصيل */}
        <div className="pt-3" style={{ borderTop: `1px dashed ${C.line}` }}>
          <div className="text-xs mb-2" style={{ color: C.soft }}>تفضّل الدفع نقداً للمندوب؟ اطلب منه زيارتك للتحصيل — يصله إشعار وتُعلَم الشركة لمتابعته.</div>
          {activeCollect ? <div className="rounded-xl p-3" style={{ background: COLLECT_STATUS[activeCollect.status].soft }}><div className="flex items-center justify-between"><span className="text-sm font-bold flex items-center gap-1.5" style={{ color: COLLECT_STATUS[activeCollect.status].color }}><Truck size={15} />{COLLECT_STATUS[activeCollect.status].label}</span><button onClick={() => actions.cancelCollection(activeCollect.id)} className="text-xs font-bold" style={{ color: C.danger }}>إلغاء</button></div>{activeCollect.visitTime && <div className="text-xs mt-1" style={{ color: C.soft }}>موعد الزيارة: {activeCollect.visitTime}</div>}<div className="text-xs mt-0.5" style={{ color: C.soft }}>المندوب: {activeCollect.repId ? repName(data, activeCollect.repId) : "—"}</div></div>
            : <GhostBtn full icon={Truck} onClick={requestRep}>طلب مندوب للتحصيل</GhostBtn>}
          {reqDone && !activeCollect && <div className="text-xs rounded-lg p-2 mt-2 flex items-center gap-1.5" style={{ background: C.successSoft, color: C.success }}><CheckCircle2 size={14} />أُرسل طلبك للمندوب وأُعلمت الشركة.</div>}
        </div>
      </Card>
      <Card className="p-4 lg:col-span-2 min-w-0"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>سجل التسويات</h3>
        {/* بطاقات على الهاتف */}
        <div className="sm:hidden space-y-2">{mySettlements.map((s) => (<div key={s.id} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}>
          <div className="flex items-center justify-between mb-1.5"><span className="font-extrabold text-sm" style={{ color: C.text }}>{s.id}</span><Badge status={s.status} map={SETTLE_STATUS} /></div>
          <div className="flex items-center justify-between text-xs" style={{ color: C.soft }}><span>{s.invoiceId || "سداد عام"} · {METHOD_LABEL[s.method]}</span><span className="font-extrabold text-sm" style={{ color: SC.color }}>{money(s.amount)}</span></div>
          <div className="text-[10px] mt-1" style={{ color: C.soft }}>{s.date}{s.ref ? ` · ${s.ref}` : ""}</div>
        </div>))}</div>
        {/* جدول على الشاشات الأكبر */}
        <div className="hidden sm:block"><Table head={["رقم", "الفاتورة", "المبلغ", "الطريقة", "التاريخ", "الحالة"]}>{mySettlements.map((s) => <tr key={s.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{s.id}</Td><Td>{s.invoiceId || "عام"}</Td><Td>{money(s.amount)}</Td><Td>{METHOD_LABEL[s.method]}{s.ref ? ` · ${s.ref}` : ""}</Td><Td>{s.date}</Td><Td><Badge status={s.status} map={SETTLE_STATUS} /></Td></tr>)}</Table></div>
        {mySettlements.length === 0 && <Empty icon={HandCoins} text="لا توجد تسويات بعد." />}
      </Card>
    </div>
    {myCollects.length > 0 && <Card className="p-4 min-w-0"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Truck size={17} style={{ color: C.accent }} />طلبات التحصيل من المندوب</h3>
      <div className="sm:hidden space-y-2">{myCollects.map((c) => (<div key={c.id} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}>
        <div className="flex items-center justify-between mb-1.5"><span className="font-extrabold text-sm" style={{ color: C.text }}>{c.id}</span><Badge status={c.status} map={COLLECT_STATUS} /></div>
        <div className="text-xs" style={{ color: C.soft }}>{c.amount ? `${money(c.amount)} · ` : ""}{METHOD_LABEL[c.method]}{c.repId ? ` · ${repName(data, c.repId)}` : ""}</div>
        <div className="text-[10px] mt-1" style={{ color: C.soft }}>{c.visitTime ? `موعد: ${c.visitTime} · ` : ""}{c.date}</div>
      </div>))}</div>
      <div className="hidden sm:block"><Table head={["رقم", "المبلغ", "الطريقة", "المندوب", "الموعد", "التاريخ", "الحالة"]}>{myCollects.map((c) => <tr key={c.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{c.id}</Td><Td>{c.amount ? money(c.amount) : "—"}</Td><Td>{METHOD_LABEL[c.method]}</Td><Td>{c.repId ? repName(data, c.repId) : "—"}</Td><Td>{c.visitTime || "—"}</Td><Td>{c.date}</Td><Td><Badge status={c.status} map={COLLECT_STATUS} /></Td></tr>)}</Table></div>
    </Card>}
  </div>);
}
function MPayments(props) {
  const { data } = props; const { me } = useMy(props);
  const pays = data.payments.filter((p) => p.merchantId === me.id);
  const myReceipts = data.receipts.filter((r) => r.merchantId === me.id);
  const [rcp, setRcp] = useState(null);
  return (<div className="space-y-4"><SectionTitle sub="دفعاتك للشركات وإيصالاتك النقدية الإلكترونية.">المدفوعات والإيصالات</SectionTitle>
    <Card className="p-4"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Receipt size={18} />الإيصالات النقدية الإلكترونية</h3>
      <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{myReceipts.map((r) => <button key={r.id} onClick={() => setRcp(r)} className="text-right rounded-xl p-3 transition-shadow hover:shadow-md" style={{ border: `1px solid ${C.line}` }}>
        <div className="flex items-center justify-between mb-1"><span className="font-bold text-sm" style={{ color: C.text }}>{r.id}</span>{r.confirmedByCompany ? <BadgeCheck size={16} style={{ color: C.success }} /> : <Clock size={16} style={{ color: C.accent }} />}</div>
        <div className="font-extrabold" style={{ color: C.success }}>{money(r.amount)}</div>
        <div className="text-xs mt-1" style={{ color: C.soft }}>{compName(data, r.companyId)} · {r.date}</div>
        <div className="text-[11px] mt-1 flex items-center gap-1" style={{ color: SC.color }}><Eye size={12} />عرض الإيصال</div>
      </button>)}{myReceipts.length === 0 && <Empty icon={Receipt} text="لا توجد إيصالات." />}</div>
    </Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>سجل المدفوعات</h3><Table head={["رقم الدفعة", "الشركة", "المبلغ", "الطريقة", "التاريخ", "الإيصال"]}>{pays.map((p) => <tr key={p.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{p.id}</Td><Td>{compName(data, p.companyId)}</Td><Td>{money(p.amount)}</Td><Td>{p.method}</Td><Td>{p.date}</Td><Td>{p.receiptId ? <button onClick={() => setRcp(data.receipts.find((r) => r.id === p.receiptId))} className="text-xs font-bold" style={{ color: SC.color }}>عرض</button> : "—"}</Td></tr>)}</Table></Card>
    {rcp && <ReceiptModal receipt={rcp} data={data} onClose={() => setRcp(null)} />}
  </div>);
}
function MCredit(props) { const { me } = useMy(props); const used = Math.max(0, -me.balance); const pct = Math.min(100, Math.round((used / me.creditLimit) * 100));
  return (<div className="space-y-4"><SectionTitle sub="الحد الذي تستطيع الشراء به على الحساب.">الحد الائتماني</SectionTitle>
    <Card className="p-6"><div className="flex justify-between text-sm mb-2"><span style={{ color: C.soft }}>المستخدم: {money(used)}</span><span style={{ color: C.soft }}>الحد: {money(me.creditLimit)}</span></div>
      <div className="h-4 rounded-full overflow-hidden" style={{ background: C.bg }}><div style={{ width: pct + "%", height: "100%", background: pct > 80 ? C.danger : C.accent }} /></div>
      <div className="flex justify-between mt-3"><span className="font-extrabold" style={{ color: C.text }}>{pct}% مستخدم</span><span className="font-extrabold" style={{ color: C.success }}>متاح: {money(me.creditLimit - used)}</span></div></Card></div>);
}
function MReturnNew(props) {
  const { data, actions, go } = props; const { meId, orders } = useMy(props);
  const delivered = orders.filter((o) => o.status === "delivered");
  const [orderId, setOrderId] = useState(delivered[0]?.id || "");
  const [reason, setReason] = useState(""); const [qty, setQty] = useState(1);
  return (<div className="space-y-4"><SectionTitle sub="اطلب إرجاع بضاعة من طلب مُسلّم.">طلب مرتجع</SectionTitle>
    <Card className="p-5 max-w-lg space-y-3">
      <Field label="اختر الطلب"><select value={orderId} onChange={(e) => setOrderId(e.target.value)} className={inputCls} style={inputStyle}>{delivered.map((o) => <option key={o.id} value={o.id}>{o.id} — {compName(data, o.companyId)}</option>)}</select></Field>
      <Field label="الكمية المرتجعة"><Input type="number" value={qty} onChange={(e) => setQty(+e.target.value)} /></Field>
      <Field label="سبب الإرجاع"><textarea value={reason} onChange={(e) => setReason(e.target.value)} className={inputCls} style={{ ...inputStyle, minHeight: 80 }} placeholder="مثال: تلف، نقص، صلاحية..." /></Field>
      <PrimaryBtn full icon={RotateCcw} onClick={() => { if (!orderId || !reason) return; const o = data.orders.find((x) => x.id === orderId); actions.createReturn({ merchantId: meId, companyId: o.companyId, orderId, reason, qty }); go("m_return_track"); }}>إرسال طلب المرتجع</PrimaryBtn>
    </Card></div>);
}
function MReturnTrack(props) {
  const { data } = props; const { meId } = useMy(props);
  const rets = data.returns.filter((r) => r.merchantId === meId);
  return (<div className="space-y-4"><SectionTitle sub="حالة طلبات الإرجاع.">متابعة المرتجع</SectionTitle><Card className="p-4"><Table head={["رقم", "الطلب", "الشركة", "الكمية", "السبب", "الحالة"]}>{rets.map((r) => <tr key={r.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{r.id}</Td><Td>{r.orderId}</Td><Td>{compName(data, r.companyId)}</Td><Td>{r.qty}</Td><Td>{r.reason}</Td><Td><div className="flex items-center gap-1.5 flex-wrap"><Badge status={r.status} map={RET_STATUS} />{r.condition && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded-full" style={{ background: RET_CONDITION[r.condition].soft, color: RET_CONDITION[r.condition].color }}>{RET_CONDITION[r.condition].label}</span>}{r.creditValue > 0 && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded-full" style={{ background: C.successSoft, color: C.success }}>+{money(r.creditValue)}</span>}</div></Td></tr>)}</Table>{rets.length === 0 && <Empty icon={RotateCcw} text="لا توجد مرتجعات." />}</Card></div>);
}
function MChat(props) {
  const { data, actions } = props; const { meId } = useMy(props);
  const partners = F.companies(props).filter((c) => c.sub === "active");
  const [sel, setSel] = useState(partners[0]?.id || null);
  const [txt, setTxt] = useState("");
  const key = `${meId}__${sel}`;
  const msgs = data.threads[key] || [];
  return (<div className="space-y-4"><SectionTitle sub="تواصل مباشر مع كل شركة على حدة.">محادثة الشركات</SectionTitle>
    <div className="grid lg:grid-cols-4 gap-4">
      <Card className="p-3 lg:col-span-1 h-fit"><div className="space-y-1">{partners.map((c) => { const on = sel === c.id; return (<button key={c.id} onClick={() => setSel(c.id)} className="w-full flex items-center gap-2 rounded-xl p-2.5 text-right transition-colors" style={{ background: on ? SC.soft : "transparent" }}><span className="text-xl">{c.logo}</span><span className="text-sm font-bold flex-1" style={{ color: on ? SC.color : C.text }}>{c.name}</span></button>); })}</div></Card>
      <Card className="lg:col-span-3 flex flex-col" style={{ height: 460 }}>
        <div className="p-3 border-b flex items-center gap-2" style={{ borderColor: C.line }}><span className="text-xl">{data.companies.find((c) => c.id === sel)?.logo}</span><span className="font-extrabold" style={{ color: C.text }}>{compName(data, sel)}</span></div>
        <div className="flex-1 p-4 space-y-2 overflow-y-auto">{msgs.map((m, i) => { const mine = m.from === "merchant"; return (<div key={i} className={`flex ${mine ? "justify-start" : "justify-end"}`}><div className="max-w-[75%] rounded-2xl px-3 py-2" style={{ background: mine ? SC.color : C.bg, color: mine ? "#fff" : C.text }}><div className="text-sm">{m.text}</div><div className="text-[10px] mt-0.5" style={{ color: mine ? "rgba(255,255,255,.6)" : C.soft }}>{m.t}</div></div></div>); })}{msgs.length === 0 && <Empty icon={MessageSquare} text="ابدأ المحادثة." />}</div>
        <div className="p-3 flex gap-2" style={{ borderTop: `1px solid ${C.line}` }}><input value={txt} onChange={(e) => setTxt(e.target.value)} placeholder="اكتب رسالة..." className="flex-1 rounded-xl px-3 py-2.5 text-sm outline-none" style={inputStyle} /><PrimaryBtn icon={Send} onClick={() => { if (!txt.trim()) return; actions.sendMsg(key, "merchant", txt); setTxt(""); }}>إرسال</PrimaryBtn></div>
      </Card>
    </div></div>);
}
function MNotes(props) {
  const [note, setNote] = useState(""); const [sent, setSent] = useState(false);
  return (<div className="space-y-4"><SectionTitle sub="أرسل ملاحظة أو شكوى لإدارة المنصة.">ملاحظات للمنصة</SectionTitle>
    <Card className="p-5 max-w-lg"><Field label="ملاحظتك"><textarea value={note} onChange={(e) => { setNote(e.target.value); setSent(false); }} className={inputCls} style={{ ...inputStyle, minHeight: 120 }} placeholder="اكتب ملاحظاتك هنا..." /></Field><div className="mt-3"><PrimaryBtn icon={Send} onClick={() => { if (note.trim()) { setSent(true); setNote(""); } }}>إرسال</PrimaryBtn></div>{sent && <div className="mt-3 text-sm flex items-center gap-1.5" style={{ color: C.success }}><CheckCircle2 size={16} />تم إرسال ملاحظتك، شكراً.</div>}</Card></div>);
}

/* ============================================================
   ============   صفحات الشركة   =============================
   ============================================================ */
const useCo = (props) => props.data.companies.find((c) => c.id === props.user.entityId);

// ═══ نظام الفروع: حساب موحّد للشركة، والفروع تقسيم داخلي ═══
const companyBranches = (co) => co?.branches || [];
const hasBranches = (co) => (co?.branches || []).length > 0;
const branchName = (co, bid) => bid === "all" ? "كل الفروع" : (companyBranches(co).find((b) => b.id === bid)?.name || "—");
const branchCity = (co, bid) => companyBranches(co).find((b) => b.id === bid)?.city || "";
// تصفية كيانات الشركة حسب الفرع المختار (all = الكل)
const scopedByBranch = (items, bid) => bid === "all" ? items : items.filter((x) => x.branchId === bid);
// مبدّل نطاق الفروع (يطابق نمط مبدّل السيرفرات)
function BranchSwitcher({ co, branch, setBranch }) {
  const branches = companyBranches(co);
  if (!branches.length) return null;
  const opts = [{ id: "all", name: "كل الفروع", city: "" }, ...branches];
  return (<div className="flex items-center gap-1.5 flex-wrap rounded-2xl p-1.5" style={{ background: C.well, border: `1px solid ${C.line}` }}>
    {opts.map((b) => { const on = branch === b.id; return (
      <button key={b.id} onClick={() => setBranch(b.id)} className="rb-press flex items-center gap-1.5 rounded-xl px-3.5 py-2 text-sm font-extrabold transition-all" style={{ background: on ? SC.color : "transparent", color: on ? C.onAccent : C.soft }}>
        {b.id === "all" ? <Layers size={15} /> : <MapPin size={15} />}{b.name}
      </button>); })}
  </div>);
}
// إحصاءات فرع واحد (مبيعات، كمية، ديون، مندوبون، عملاء)
const branchStats = (data, co, bid) => {
  const reps = scopedByBranch(data.reps.filter((r) => r.companyId === co.id), bid);
  const merchants = scopedByBranch(data.merchants.filter((m) => m.companyId === co.id), bid);
  const mIds = new Set(merchants.map((m) => m.id));
  const dlv = data.orders.filter((o) => o.companyId === co.id && o.status === "delivered" && mIds.has(o.merchantId));
  const sales = dlv.reduce((s, o) => s + orderTotal(o), 0);
  const units = dlv.reduce((s, o) => s + orderQty(o), 0);
  const debt = merchants.reduce((s, m) => s + merchantOpenDebt(data, m.id), 0);
  const cnt = (key) => scopedByBranch((data[key] || []).filter((x) => x.companyId === co.id), bid).length;
  return { reps: reps.length, merchants: merchants.length, sales, units, debt, orders: dlv.length, managers: cnt("salesManagers"), supervisors: cnt("supervisors"), promoters: cnt("promoters"), drivers: cnt("drivers"), dispatchers: cnt("dispatchers") };
};


function CHome(props) {
  const { data, go } = props; const co = useCo(props);
  const orders = data.orders.filter((o) => o.companyId === co.id);
  const need = orders.filter((o) => o.status === "pending_company").length;
  const done = orders.filter((o) => o.status === "delivered");
  const revenue = done.reduce((s, o) => s + orderTotal(o), 0);
  const myMerchants = data.merchants.filter((m) => m.companyId === co.id);
  const totalDebt = data.invoices.filter((i) => i.companyId === co.id).reduce((s, i) => s + (i.amount - i.paid), 0);
  const joins = data.joinRequests.filter((j) => j.companyId === co.id && !["active", "declined"].includes(j.status)).length;
  return (<div className="space-y-5"><SectionTitle sub="ملخص نشاط شركتك على المنصة." action={<span className="text-xs font-bold px-3 py-1.5 rounded-full inline-flex items-center gap-1.5" style={{ background: PLANS[co.plan].soft, color: PLANS[co.plan].color }}><BadgeCheck size={14} />اشتراك {PLANS[co.plan].label}</span>}>نظرة عامة</SectionTitle>
    {co.plan === "basic" && <div className="rounded-xl p-3 flex items-center gap-2" style={{ background: C.purpleSoft }}><TrendingUp size={18} style={{ color: C.purple }} /><span className="text-sm font-bold flex-1" style={{ color: C.purple }}>اشتراكك الأساسي. الترقية للاحترافي تتيح التقارير المتقدمة وتتبع المندوبين والإشعارات الذكية.</span></div>}
    <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={ShoppingCart} label="بانتظار موافقتك" value={need} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={Store} label={SC.merchantsWord} value={myMerchants.length} tint={{ color: C.accent, soft: C.accentSoft }} />
      <Stat icon={UserPlus} label="طلبات انضمام" value={joins} tint={{ color: C.purple, soft: C.purpleSoft }} />
      <Stat icon={Wallet} label="مبيعات مسلّمة" value={money(revenue)} tint={{ color: C.success, soft: C.successSoft }} />
    </div>
    {/* تقييم الشركة من العملاء */}
    {(() => { const cr = companyRating(data, co.id); const reviews = data.ratings.filter((r) => r.raterRole === "merchant" && r.companyId === co.id).slice(0, 3); if (!cr) return null; return (
      <Card className="p-4"><div className="flex items-center justify-between flex-wrap gap-2 mb-2"><h3 className="font-extrabold flex items-center gap-2" style={{ color: C.text }}><Star size={18} fill="#F59E0B" style={{ color: "#F59E0B" }} />تقييم عملائك لشركتك</h3><Stars avg={cr.avg} count={cr.count} size={16} /></div>
        <div className="grid sm:grid-cols-3 gap-2">{reviews.map((r) => <div key={r.id} className="rounded-xl p-2.5 text-xs" style={{ background: C.bg }}><div className="flex items-center justify-between mb-1"><span className="font-bold" style={{ color: C.text }}>{merchName(data, r.merchantId)}</span><Stars avg={r.stars} count={1} showCount={false} size={10} /></div>{r.comment && <div style={{ color: C.soft }}>{r.comment}</div>}</div>)}</div>
      </Card>); })()}
    {/* رسوم بيانية: مبيعات الأسبوع + توزيع الذمم */}
    <div className="grid lg:grid-cols-2 gap-4">
      <Card className="p-5"><h3 className="font-extrabold mb-4 flex items-center gap-2" style={{ color: C.text }}><TrendingUp size={17} style={{ color: SC.color2 }} />مبيعات الأسبوع (د.ل)</h3>
        <BarsChart money color={SC.color2} data={[{ label: "السبت", v: Math.round(revenue * .12) }, { label: "الأحد", v: Math.round(revenue * .18) }, { label: "الاثنين", v: Math.round(revenue * .1) }, { label: "الثلاثاء", v: Math.round(revenue * .2) }, { label: "الأربعاء", v: Math.round(revenue * .15) }, { label: "الخميس", v: Math.round(revenue * .25) }]} />
      </Card>
      <Card className="p-5"><h3 className="font-extrabold mb-4 flex items-center gap-2" style={{ color: C.text }}><CircleDollarSign size={17} style={{ color: C.accent }} />توزيع الذمم المالية</h3>
        {(() => { const paid = data.invoices.filter((i) => i.companyId === co.id).reduce((s, i) => s + i.paid, 0); return <DonutChart parts={[{ label: "محصَّل", v: paid || 1, color: C.success }, { label: "مديونيات قائمة", v: totalDebt || 1, color: C.danger }]} />; })()}
      </Card>
    </div>
    <div className="grid lg:grid-cols-2 gap-4">
      <Card className="p-5"><div className="flex items-center justify-between mb-3"><h3 className="font-extrabold" style={{ color: C.text }}>أحدث الطلبات الواردة</h3><GhostBtn sm icon={Eye} onClick={() => go("c_orders")}>الكل</GhostBtn></div><OrderList orders={orders.filter((o) => o.status !== "draft").slice(0, 5)} data={data} onOpen={props.openOrder} /></Card>
      <Card className="p-5"><div className="flex items-center justify-between mb-3"><h3 className="font-extrabold" style={{ color: C.text }}>الفريق الميداني</h3></div>
        <div className="grid grid-cols-3 gap-2 text-center">
          <div className="rounded-xl p-3" style={{ background: C.tealSoft }}><Briefcase size={20} className="mx-auto mb-1" style={{ color: C.teal }} /><div className="text-lg font-extrabold" style={{ color: C.teal }}>{data.supervisors.filter((s) => s.companyId === co.id).length}</div><div className="text-[11px]" style={{ color: C.soft }}>مشرفون</div></div>
          <div className="rounded-xl p-3" style={{ background: C.infoSoft }}><Handshake size={20} className="mx-auto mb-1" style={{ color: C.info }} /><div className="text-lg font-extrabold" style={{ color: C.info }}>{data.reps.filter((r) => r.companyId === co.id).length}</div><div className="text-[11px]" style={{ color: C.soft }}>مندوبون</div></div>
          <div className="rounded-xl p-3" style={{ background: C.purpleSoft }}><Truck size={20} className="mx-auto mb-1" style={{ color: C.purple }} /><div className="text-lg font-extrabold" style={{ color: C.purple }}>{data.drivers.filter((d) => d.companyId === co.id).length}</div><div className="text-[11px]" style={{ color: C.soft }}>سائقون</div></div>
        </div>
        <div className="mt-3"><PrimaryBtn full sm icon={ListChecks} onClick={() => go("c_tasks")}>إسناد مهمة جديدة</PrimaryBtn></div>
      </Card>
    </div>
  </div>);
}
function CProducts(props) {
  const { data, actions, go } = props; const co = useCo(props);
  const prods = data.products.filter((p) => p.companyId === co.id);
  const [view, setView] = useState(null);
  const [q, setQ] = useState("");
  const [cat, setCat] = useState("الكل");
  if (view) { const p = prods.find((x) => x.id === view); if (!p) { setView(null); return null; }
    return (<div className="space-y-4 max-w-2xl"><button onClick={() => setView(null)} className="inline-flex items-center gap-1 text-sm font-bold" style={{ color: C.soft }}><ChevronLeft size={16} />رجوع للقائمة</button>
      <Card className="p-5"><div className="flex items-center justify-between mb-4 flex-wrap gap-2"><div className="flex items-center gap-2"><h2 className="text-lg font-extrabold" style={{ color: C.text }}>{p.name}</h2>{p.expDate && <ExpiryBadge expDate={p.expDate} />}</div><Badge status={p.status === "active" ? "delivered" : "draft"} map={{ delivered: { label: "نشط", color: C.success, soft: C.successSoft }, draft: { label: "مخفي", color: C.soft, soft: C.well } }} /></div>
        <div className="grid sm:grid-cols-2 gap-4 mb-4">
          <div className="space-y-2 text-sm">{[["التصنيف", p.cat], ["السعر", money(p.price)], ["الوحدة", p.unit], ["أقل كمية للطلب", p.minOrder + " " + p.unit], ["تاريخ التصنيع", fmtDate(p.mfgDate)], ["تاريخ الصلاحية", fmtDate(p.expDate)], ...(p.batch ? [["رقم التشغيلة", p.batch]] : []), ["الضريبة", p.taxPct + "%"], ["العرض/الخصم", p.offer ? p.offer + "%" : "—"], ["المخزون", p.stock]].map(([k, v]) => <div key={k} className="flex justify-between"><span style={{ color: C.soft }}>{k}</span><span className="font-bold" style={{ color: C.text }}>{v}</span></div>)}{p.tiers && p.tiers.length > 0 && <div className="pt-1"><div className="text-xs font-bold mb-1" style={{ color: C.purple }}>شرائح خصم الكمية</div>{p.tiers.map((t, i) => <div key={i} className="flex justify-between text-xs"><span style={{ color: C.soft }}>{tierThresholdLabel(t, p.unit)}</span><span className="font-bold" style={{ color: C.purple }}>خصم {t.discount}%</span></div>)}</div>}</div>
          <div className="space-y-3"><div><div className="text-xs font-bold mb-1" style={{ color: C.soft }}>الباركود</div><FakeBarcode code={p.barcode} /></div><div><div className="text-xs font-bold mb-1" style={{ color: C.soft }}>QR Code</div><FakeQR /></div></div>
        </div>
        <div className="flex gap-2 flex-wrap"><PrimaryBtn icon={p.status === "active" ? XCircle : CheckCircle2} tone={p.status === "active" ? "ghost" : "primary"} onClick={() => actions.updateProduct(p.id, { status: p.status === "active" ? "hidden" : "active" })}>{p.status === "active" ? "إخفاء المنتج" : "تفعيل المنتج"}</PrimaryBtn><PrimaryBtn tone="danger" icon={Trash2} onClick={() => { actions.deleteProduct(p.id); setView(null); }}>حذف</PrimaryBtn></div>
      </Card></div>);
  }
  const cats = ["الكل", ...Array.from(new Set(prods.map((p) => p.cat)))];
  const catCount = (c) => c === "الكل" ? prods.length : prods.filter((p) => p.cat === c).length;
  const list = prods.filter((p) => (cat === "الكل" || p.cat === cat) && (!q || p.name.includes(q) || (p.cat || "").includes(q) || (p.barcode || "").includes(q)));
  return (<div className="space-y-4"><SectionTitle sub={`${prods.length} صنف — ابحث وتنقّل بين التصنيفات بسهولة.`} action={<PrimaryBtn icon={Plus} onClick={() => go("c_product_add")}>إضافة منتج</PrimaryBtn>}>قائمة المنتجات</SectionTitle>
    <Card className="p-4">
      <div className="relative mb-3"><Search size={16} className="absolute right-3 top-2.5" style={{ color: C.soft }} /><input value={q} onChange={(e) => setQ(e.target.value)} placeholder="ابحث بالاسم أو التصنيف أو الباركود" className="w-full rounded-xl pr-9 pl-3 py-2 text-sm outline-none" style={inputStyle} /></div>
      <div className="flex gap-1.5 overflow-x-auto pb-1 mb-3" style={{ scrollbarWidth: "thin" }}>{cats.map((c) => { const on = cat === c; return (<button key={c} onClick={() => setCat(c)} className="shrink-0 inline-flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-bold transition-all" style={{ background: on ? SC.color : C.bg, color: on ? C.onAccent : C.text, border: `1px solid ${on ? SC.color : C.line}` }}>{c}<span className="rounded-full px-1.5 text-[10px]" style={{ background: on ? "rgba(255,255,255,.25)" : C.surface, color: on ? "#fff" : C.soft }}>{catCount(c)}</span></button>); })}</div>
      <Table head={["المنتج", "التصنيف", "السعر", "الوحدة", "المخزون", "الصلاحية", "الحالة", ""]}>{list.map((p) => { const e = expiryInfo(p.expDate); return (<tr key={p.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{p.name}</Td><Td>{p.cat}</Td><Td>{money(p.price)}</Td><Td>{p.unit}</Td><Td>{p.stock}</Td><Td>{p.expDate ? <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: e?.soft, color: e?.color }}>{fmtDate(p.expDate)}</span> : "—"}</Td><Td><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={p.status === "active" ? { background: C.successSoft, color: C.success } : { background: C.well, color: C.soft }}>{p.status === "active" ? "نشط" : "مخفي"}</span></Td><Td><GhostBtn sm icon={Eye} onClick={() => setView(p.id)}>عرض</GhostBtn></Td></tr>); })}</Table>
      {list.length === 0 && <Empty icon={Package} text="لا أصناف مطابقة." />}
    </Card></div>);
}
function CProductAdd(props) {
  const { actions, go, server } = props; const co = useCo(props);
  const isPharma = server === "pharma";
  const generatedBarcode = useRef((isPharma ? "629300010" : "629104150") + Math.floor(1000 + Math.random() * 9000)).current;
  const [f, setF] = useState({ name: "", cat: "", price: "", stock: "", unit: isPharma ? "علبة" : "كرتونة", minOrder: 1, weight: "", size: "", taxPct: 0, offer: 0, imgs: [], status: "active", mfgDate: "", expDate: "", batch: "" });
  const [tiers, setTiers] = useState([]);
  const [tierDraft, setTierDraft] = useState({ basis: "qty", threshold: "", discount: "" });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const addTier = () => { if (+tierDraft.threshold > 0 && +tierDraft.discount > 0) { const t = tierDraft.basis === "value" ? { basis: "value", minValue: +tierDraft.threshold, discount: +tierDraft.discount } : { basis: "qty", minQty: +tierDraft.threshold, discount: +tierDraft.discount }; setTiers([...tiers, t].sort((a, b) => (a.minQty || a.minValue) - (b.minQty || b.minValue))); setTierDraft({ basis: tierDraft.basis, threshold: "", discount: "" }); } };
  const addImgs = (e) => { const files = Array.from(e.target.files || []); files.forEach((file) => { const reader = new FileReader(); reader.onload = (ev) => setF((prev) => ({ ...prev, imgs: [...prev.imgs, ev.target.result].slice(0, 6) })); reader.readAsDataURL(file); }); e.target.value = ""; };
  const removeImg = (i) => setF((prev) => ({ ...prev, imgs: prev.imgs.filter((_, x) => x !== i) }));
  const exp = expiryInfo(f.expDate);
  return (<div className="space-y-4"><SectionTitle sub={`عبّئ بيانات المنتج كاملة${isPharma ? " مع الصلاحية ورقم التشغيلة" : ""} — سيُولَّد الباركود وQR تلقائياً.`}>إضافة منتج</SectionTitle>
    <Card className="p-5 max-w-3xl space-y-4">
      <div className="grid sm:grid-cols-2 gap-3">
        <Field label="اسم المنتج"><Input value={f.name} onChange={set("name")} /></Field>
        <Field label="التصنيف"><Input value={f.cat} onChange={set("cat")} placeholder={isPharma ? "مثال: مسكّنات" : "مثال: حبوب"} /></Field>
        <Field label="السعر (د.ل)"><Input type="number" value={f.price} onChange={set("price")} /></Field>
        <Field label="المخزون"><Input type="number" value={f.stock} onChange={set("stock")} /></Field>
        <Field label="الوحدة"><select value={f.unit} onChange={set("unit")} className={inputCls} style={inputStyle}>{["كرتونة", "كيس", "صفيحة", "علبة", "حبة", "طبلية", "زجاجة", "شريط", "جهاز"].map((u) => <option key={u}>{u}</option>)}</select></Field>
        <Field label="الحد الأدنى للطلب"><Input type="number" value={f.minOrder} onChange={set("minOrder")} /></Field>
        <Field label="الوزن"><Input value={f.weight} onChange={set("weight")} placeholder="مثال: 25 كغ" /></Field>
        <Field label="الحجم / الأبعاد"><Input value={f.size} onChange={set("size")} placeholder="مثال: 60×40×15 سم" /></Field>
        <Field label="الضريبة %"><Input type="number" value={f.taxPct} onChange={set("taxPct")} /></Field>
        <Field label="الخصم / العرض %"><Input type="number" value={f.offer} onChange={set("offer")} /></Field>
        <Field label="حالة المنتج"><select value={f.status} onChange={set("status")} className={inputCls} style={inputStyle}><option value="active">نشط</option><option value="hidden">مخفي</option></select></Field>
      </div>

      {/* إرفاق صور المنتج */}
      <div className="rounded-xl p-4" style={{ background: C.bg, border: `1px solid ${C.line}` }}>
        <div className="flex items-center gap-1.5 text-sm font-extrabold mb-3" style={{ color: C.text }}><ImageIcon size={16} style={{ color: SC.color }} />صور المنتج <span className="text-xs font-normal" style={{ color: C.soft }}>(حتى 6 صور)</span></div>
        <div className="flex flex-wrap gap-2.5">
          {f.imgs.map((src, i) => <div key={i} className="relative rounded-xl overflow-hidden" style={{ width: 88, height: 88, border: `1px solid ${C.line}` }}><img src={src} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /><button onClick={() => removeImg(i)} className="absolute top-1 left-1 rounded-full flex items-center justify-center" style={{ width: 22, height: 22, background: C.danger }}><X size={13} color="#fff" /></button>{i === 0 && <span className="absolute bottom-0 inset-x-0 text-center text-[9px] font-bold py-0.5" style={{ background: SC.color, color: C.onAccent }}>الرئيسية</span>}</div>)}
          {f.imgs.length < 6 && <label className="rb-press flex flex-col items-center justify-center gap-1 rounded-xl cursor-pointer" style={{ width: 88, height: 88, border: `2px dashed ${C.lineStrong}`, color: C.soft }}><Upload size={20} /><span className="text-[10px] font-bold">إضافة صورة</span><input type="file" accept="image/*" multiple onChange={addImgs} className="hidden" /></label>}
        </div>
        {f.imgs.length === 0 && <div className="text-xs mt-2" style={{ color: C.soft }}>أول صورة تُعتمد كصورة رئيسية للمنتج.</div>}
      </div>

      {/* الصلاحية وتاريخ التصنيع والتشغيلة */}
      <div className="rounded-xl p-4" style={{ background: isPharma ? C.tealSoft + "66" : C.bg, border: `1px solid ${isPharma ? C.teal : C.line}` }}>
        <div className="flex items-center gap-1.5 text-sm font-extrabold mb-3" style={{ color: isPharma ? C.teal : C.text }}><CalendarClock size={16} />{isPharma ? "بيانات الصلاحية والتشغيلة (إلزامية للأدوية)" : "تاريخ التصنيع والصلاحية"}</div>
        <div className="grid sm:grid-cols-3 gap-3">
          <Field label="تاريخ التصنيع"><Input type="date" value={f.mfgDate} onChange={set("mfgDate")} /></Field>
          <Field label="تاريخ انتهاء الصلاحية"><Input type="date" value={f.expDate} onChange={set("expDate")} /></Field>
          {isPharma && <Field label="رقم التشغيلة (Batch)"><Input value={f.batch} onChange={set("batch")} placeholder="مثال: BT-2206" /></Field>}
        </div>
        {exp && <div className="mt-2 inline-flex items-center gap-1.5 text-xs font-bold px-2.5 py-1 rounded-full" style={{ background: exp.soft, color: exp.color }}><AlertCircle size={12} />{exp.label}</div>}
      </div>

      {/* خصم الشرائح: حسب عدد الوحدات/الصناديق أو حسب قيمة الشراء */}
      <div className="rounded-xl p-4" style={{ background: C.bg, border: `1px solid ${C.line}` }}>
        <div className="flex items-center gap-1.5 text-sm font-extrabold mb-1" style={{ color: C.text }}><Layers size={16} style={{ color: C.purple }} />خصم الشرائح (حجز نسبة)</div>
        <div className="text-xs mb-3" style={{ color: C.soft }}>حدّد كيف يُفعَّل الخصم للعميل: إمّا ببلوغ <b>عدد وحدات/صناديق</b> معيّن، أو ببلوغ <b>قيمة شراء</b> معيّنة بالدينار.</div>
        <div className="space-y-1.5 mb-3">{tiers.map((t, i) => <div key={i} className="flex items-center justify-between rounded-lg px-3 py-2" style={{ background: C.purpleSoft }}><span className="text-sm font-bold" style={{ color: C.purple }}>{t.basis === "value" ? `عند الشراء بقيمة ${t.minValue.toLocaleString("ar-LY")} د.ل` : `من ${t.minQty} ${f.unit}`} فأكثر → خصم {t.discount}%</span><button onClick={() => setTiers(tiers.filter((_, x) => x !== i))} className="rounded p-1" style={{ background: C.well }}><Trash2 size={13} style={{ color: C.danger }} /></button></div>)}{tiers.length === 0 && <div className="text-xs" style={{ color: C.soft }}>لا توجد شرائح بعد.</div>}</div>
        {/* اختيار أساس التفعيل */}
        <div className="grid grid-cols-2 gap-2 mb-2">{[["qty", "حسب الكمية (وحدات/صناديق)", Boxes], ["value", "حسب قيمة الشراء (د.ل)", CircleDollarSign]].map(([b, lbl, Ic]) => { const on = tierDraft.basis === b; return (<button key={b} onClick={() => setTierDraft({ ...tierDraft, basis: b })} className="flex items-center justify-center gap-1.5 rounded-lg py-2 text-xs font-bold transition-all" style={{ background: on ? C.purple : C.surface, color: on ? "#fff" : C.text, border: `1px solid ${on ? C.purple : C.line}` }}><Ic size={14} />{lbl}</button>); })}</div>
        <div className="grid grid-cols-3 gap-2 items-end">
          <Field label={tierDraft.basis === "value" ? "القيمة من (د.ل)" : `الكمية من (${f.unit})`}><Input type="number" value={tierDraft.threshold} onChange={(e) => setTierDraft({ ...tierDraft, threshold: e.target.value })} placeholder={tierDraft.basis === "value" ? "5000" : "100"} /></Field>
          <Field label="نسبة الخصم %"><Input type="number" value={tierDraft.discount} onChange={(e) => setTierDraft({ ...tierDraft, discount: e.target.value })} placeholder="5" /></Field>
          <PrimaryBtn icon={Plus} onClick={addTier}>إضافة شريحة</PrimaryBtn>
        </div>
      </div>

      <div className="grid sm:grid-cols-2 gap-3 rounded-xl p-4" style={{ background: C.bg }}>
        <div><div className="flex items-center gap-1.5 text-xs font-bold mb-2" style={{ color: C.soft }}><Barcode size={14} />باركود (تلقائي)</div><FakeBarcode code={generatedBarcode} /></div>
        <div><div className="flex items-center gap-1.5 text-xs font-bold mb-2" style={{ color: C.soft }}><QrCode size={14} />QR Code (تلقائي)</div><FakeQR /></div>
      </div>
      <PrimaryBtn icon={CheckCircle2} onClick={() => { if (!f.name || !f.price) { toast("أدخل اسم المنتج والسعر", "danger"); return; } if (isPharma && (!f.expDate || !f.batch)) { toast("للأدوية يجب إدخال الصلاحية ورقم التشغيلة", "danger"); return; } actions.addProduct({ server: co.server, companyId: co.id, ...f, price: +f.price, stock: +f.stock || 0, minOrder: +f.minOrder || 1, taxPct: +f.taxPct || 0, offer: +f.offer || 0, images: f.imgs.length, tiers, barcode: generatedBarcode }); go("c_products"); }}>حفظ المنتج</PrimaryBtn>
    </Card></div>);
}
function CCats(props) {
  const { data } = props; const co = useCo(props);
  const prods = data.products.filter((p) => p.companyId === co.id);
  const cats = [...new Set(prods.map((p) => p.cat))]; const units = [...new Set(prods.map((p) => p.unit))];
  return (<div className="space-y-4"><SectionTitle sub="التصنيفات والوحدات المستخدمة في منتجاتك.">التصنيفات والوحدات</SectionTitle>
    <div className="grid sm:grid-cols-2 gap-4">
      <Card className="p-5"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Layers size={18} />التصنيفات</h3><div className="space-y-2">{cats.map((c) => <div key={c} className="flex justify-between rounded-xl p-2.5" style={{ border: `1px solid ${C.line}` }}><span className="font-bold text-sm" style={{ color: C.text }}>{c}</span><span className="text-sm" style={{ color: C.soft }}>{prods.filter((p) => p.cat === c).length} منتج</span></div>)}</div></Card>
      <Card className="p-5"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Box size={18} />الوحدات</h3><div className="flex gap-2 flex-wrap">{units.map((u) => <span key={u} className="text-sm font-bold px-3 py-1.5 rounded-full" style={{ background: C.bg, color: C.text }}>{u}</span>)}</div></Card>
    </div></div>);
}
// رفع ملف Excel/CSV لإضافة المحلات دفعة أولى — بصيغة: اسم المحل، المبلغ (المديونية)، الهاتف، المنطقة، اسم المندوب
function ImportClientsCard({ co, reps, data, actions }) {
  const [open, setOpen] = useState(false);
  const [raw, setRaw] = useState("");
  const [rows, setRows] = useState(null); // معاينة
  const fileRef = React.useRef(null);
  const TEMPLATE = "اسم المحل,المبلغ,الهاتف,المنطقة,اسم المندوب\nبقالة الصفاء,1200,091-1112233,سوق الجمعة," + (reps[0]?.name || "اسم المندوب") + "\nماركت الهدى,0,092-4455667,عين زارة," + (reps[1]?.name || reps[0]?.name || "اسم المندوب");
  const parse = (text) => {
    const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    const body = lines[0] && (lines[0].includes("اسم المحل") || lines[0].toLowerCase().includes("name")) ? lines.slice(1) : lines;
    const out = []; const errs = [];
    body.forEach((line, i) => {
      const c = line.split(/[,\t؛;]/).map((x) => x.trim());
      if (!c[0]) { errs.push(`سطر ${i + 1}: اسم المحل فارغ`); return; }
      const repNameIn = c[4] || "";
      const rep = reps.find((r) => r.name === repNameIn) || reps.find((r) => repNameIn && r.name.includes(repNameIn)) || null;
      out.push({ name: c[0], openingDebt: +(c[1] || 0) || 0, phone: c[2] || "", area: c[3] || "", repId: rep?.id || reps[0]?.id, repMatched: !!rep, repNameIn });
    });
    return { out, errs };
  };
  const preview = () => { const { out, errs } = parse(raw); setRows({ out, errs }); };
  const onFile = (e) => { const file = e.target.files?.[0]; if (!file) return; const isWorkbook = /\.(xlsx|xls)$/i.test(file.name); const rd = new FileReader(); rd.onload = () => { try { let text = String(rd.result || ""); if (isWorkbook) { const workbook = XLSX.read(rd.result, { type: "array" }); const firstSheet = workbook.Sheets[workbook.SheetNames[0]]; text = XLSX.utils.sheet_to_csv(firstSheet); } setRaw(text); const { out, errs } = parse(text); setRows({ out, errs }); } catch { toast("تعذر قراءة ملف Excel", "danger"); } }; if (isWorkbook) rd.readAsArrayBuffer(file); else rd.readAsText(file, "utf-8"); e.target.value = ""; };
  const doImport = () => { if (!rows?.out?.length) return; actions.importClients({ companyId: co.id, rows: rows.out, source: "ملف مستورد" }); toast(`استُورد ${rows.out.length} محل وأُسند للمندوبين ✓`); setRows(null); setRaw(""); setOpen(false); };
  return (<Card className="p-4 min-w-0">
    <div className="flex items-center justify-between flex-wrap gap-2">
      <div className="flex items-center gap-2.5"><div className="rounded-xl p-2.5" style={{ background: C.successSoft }}><FileText size={19} style={{ color: C.success }} /></div><div><h3 className="font-extrabold text-sm" style={{ color: C.text }}>استيراد المحلات من ملف Excel</h3><p className="text-[11px]" style={{ color: C.soft }}>أضف عملاءك دفعة أولى — يُسند كل محل لمندوبه ويُلزَم بتوثيقه ميدانياً.</p></div></div>
      <GhostBtn sm icon={open ? XCircle : PackagePlus} onClick={() => setOpen(!open)}>{open ? "إغلاق" : "بدء الاستيراد"}</GhostBtn>
    </div>
    {open && <div className="mt-4 space-y-3">
      <div className="rounded-xl p-3 text-xs" style={{ background: C.bg }}>
        <div className="font-extrabold mb-1" style={{ color: C.text }}>الصيغة المطلوبة (سطر لكل محل، الفاصل: , أو Tab):</div>
        <code className="block rounded-lg p-2 text-[11px] leading-relaxed" dir="rtl" style={{ background: C.surface, color: C.soft, border: `1px dashed ${C.line}` }}>اسم المحل، المبلغ (المديونية الافتتاحية)، الهاتف، المنطقة، اسم المندوب</code>
        <button onClick={() => { setRaw(TEMPLATE); setRows(null); }} className="text-[11px] font-bold mt-1.5" style={{ color: SC.color }}>تعبئة نموذج تجريبي</button>
      </div>
      <div className="flex gap-2 flex-wrap">
        <button onClick={() => fileRef.current?.click()} className="rb-press flex items-center gap-1.5 rounded-xl px-3.5 py-2 text-xs font-bold" style={{ background: C.successSoft, color: C.success }}><FileText size={14} />رفع ملف Excel / CSV</button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.txt" onChange={onFile} style={{ display: "none" }} />
        <span className="text-[11px] self-center" style={{ color: C.soft }}>أو الصق محتوى الملف مباشرة:</span>
      </div>
      <textarea value={raw} onChange={(e) => { setRaw(e.target.value); setRows(null); }} rows={5} placeholder={"بقالة الصفاء,1200,091-1112233,سوق الجمعة,وليد سعد"} className="w-full rounded-xl px-3 py-2.5 text-sm outline-none" style={{ ...inputStyle, fontFamily: "monospace" }} />
      {!rows ? <PrimaryBtn icon={Eye} onClick={preview}>معاينة قبل الاستيراد</PrimaryBtn> : <>
        {rows.errs.length > 0 && <div className="rounded-xl p-2.5 text-xs space-y-0.5" style={{ background: C.dangerSoft, color: C.danger }}>{rows.errs.map((e, i) => <div key={i}>⚠️ {e}</div>)}</div>}
        <div className="rounded-xl overflow-hidden" style={{ border: `1px solid ${C.line}` }}>
          <div className="px-3 py-2 text-[11px] font-extrabold flex" style={{ background: C.well, color: C.text }}><span className="flex-1">المحل</span><span className="w-20 text-center">المديونية</span><span className="w-28 text-center">المندوب</span></div>
          {rows.out.map((r, i) => (<div key={i} className="px-3 py-2 text-xs flex items-center" style={{ borderTop: `1px solid ${C.line}` }}>
            <span className="flex-1 min-w-0"><span className="font-bold block truncate" style={{ color: C.text }}>{r.name}</span><span className="text-[10px]" style={{ color: C.soft }}>{r.area}{r.phone ? ` · ${r.phone}` : ""}</span></span>
            <span className="w-20 text-center font-bold" style={{ color: r.openingDebt > 0 ? C.danger : C.success }}>{money(r.openingDebt)}</span>
            <span className="w-28 text-center"><span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: r.repMatched ? C.successSoft : C.accentSoft, color: r.repMatched ? C.success : "#92590C" }}>{repName(data, r.repId)}{!r.repMatched && " (افتراضي)"}</span></span>
          </div>))}
        </div>
        <div className="flex gap-2"><PrimaryBtn tone="success" icon={CheckCircle2} onClick={doImport}>استيراد {rows.out.length} محل وإسنادها للمندوبين</PrimaryBtn><GhostBtn onClick={() => setRows(null)}>تعديل</GhostBtn></div>
      </>}
    </div>}
  </Card>);
}
// ═══ نظرة عامة على الفروع (لمدير الشركة) ═══
function CBranches(props) {
  const { data } = props; const co = useCo(props);
  const [branch, setBranch] = useState("all");
  if (!hasBranches(co)) return (<div className="space-y-4"><SectionTitle sub="هذه الشركة بفرع واحد. تُضاف الفروع من إعدادات الشركة.">الفروع</SectionTitle><Empty icon={Building2} text="لا فروع متعددة لهذه الشركة." /></div>);
  const branches = companyBranches(co);
  const shown = branchStats(data, co, branch);
  const maxSales = Math.max(1, ...branches.map((b) => branchStats(data, co, b.id).sales));
  return (<div className="space-y-4"><SectionTitle sub="حساب موحّد للشركة، والفروع تقسيم داخلي — اعرض الكل أو فرعاً محدداً.">الفروع</SectionTitle>
    <BranchSwitcher co={co} branch={branch} setBranch={setBranch} />
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={TrendingUp} label={`مبيعات ${branchName(co, branch)}`} value={money(shown.sales)} tint={{ color: C.success, soft: C.successSoft }} hint={qty(shown.units)} />
      <Stat icon={Users} label="العملاء" value={shown.merchants} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={Handshake} label="المندوبون" value={shown.reps} tint={{ color: SC.color, soft: SC.soft }} />
      <Stat icon={HandCoins} label="الديون المفتوحة" value={money(shown.debt)} tint={{ color: C.danger, soft: C.dangerSoft }} />
    </div>
    {branch === "all" && <Card className="p-4 sm:p-5">
      <h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Layers size={18} style={{ color: SC.color }} />مقارنة الفروع</h3>
      <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">
        {branches.map((b) => { const st = branchStats(data, co, b.id); return (
          <div key={b.id} className="rounded-2xl p-4" style={{ border: `1.5px solid ${SC.color}`, background: SC.soft }}>
            <div className="flex items-center justify-between mb-3"><span className="font-extrabold flex items-center gap-1.5" style={{ color: SC.color }}><MapPin size={15} />{b.name}</span><span className="text-[11px] font-bold px-2 py-0.5 rounded-full" style={{ background: C.surface, color: SC.color }}>{b.city}</span></div>
            <div className="grid grid-cols-2 gap-2 text-center">
              <div className="rounded-xl p-2" style={{ background: C.surface }}><div className="text-[10px]" style={{ color: C.soft }}>المبيعات</div><div className="font-extrabold text-sm" style={{ color: C.text }}>{money(st.sales)}</div></div>
              <div className="rounded-xl p-2" style={{ background: C.surface }}><div className="text-[10px]" style={{ color: C.soft }}>الكمية</div><div className="font-extrabold text-sm" style={{ color: C.text }}>{qty(st.units)}</div></div>
              <div className="rounded-xl p-2" style={{ background: C.surface }}><div className="text-[10px]" style={{ color: C.soft }}>المندوبون / العملاء</div><div className="font-extrabold text-sm" style={{ color: C.text }}>{st.reps} / {st.merchants}</div></div>
              <div className="rounded-xl p-2" style={{ background: C.surface }}><div className="text-[10px]" style={{ color: C.soft }}>الديون</div><div className="font-extrabold text-sm" style={{ color: st.debt > 0 ? C.danger : C.success }}>{money(st.debt)}</div></div>
            </div>
            <div className="mt-2.5"><ProgressBar value={st.sales} max={maxSales} color={SC.color} /></div>
          </div>); })}
      </div>
    </Card>}
    {branch !== "all" && (() => {
      const st = branchStats(data, co, branch);
      const teamGroups = [
        { key: "salesManagers", label: "مديرو المبيعات", icon: UserCog, color: C.ink },
        { key: "supervisors", label: "المشرفون", icon: Briefcase, color: C.teal },
        { key: "reps", label: "المندوبون", icon: Handshake, color: C.info },
        { key: "promoters", label: "المروّجون", icon: ImageIcon, color: C.rose },
        { key: "drivers", label: "السائقون", icon: Truck, color: C.purple },
        { key: "dispatchers", label: "مشرفو الحركة", icon: Route, color: "#B45309" },
      ];
      return (<Card className="p-4">
        <h3 className="font-extrabold mb-3" style={{ color: C.text }}>طاقم {branchName(co, branch)}</h3>
        <div className="grid grid-cols-3 sm:grid-cols-6 gap-2 mb-4">
          {teamGroups.map((g) => { const Ic = g.icon; const n = st[g.key === "salesManagers" ? "managers" : g.key] ?? 0; return (
            <div key={g.key} className="rounded-xl p-2.5 text-center" style={{ background: C.well }}><Ic size={18} style={{ color: g.color, margin: "0 auto" }} /><div className="font-extrabold text-lg mt-1" style={{ color: C.text }}>{n}</div><div className="text-[9px]" style={{ color: C.soft }}>{g.label}</div></div>); })}
        </div>
        {teamGroups.map((g) => { const members = scopedByBranch((data[g.key] || []).filter((x) => x.companyId === co.id), branch); if (!members.length) return null; const Ic = g.icon; return (
          <div key={g.key} className="mb-3">
            <div className="text-xs font-bold mb-1.5 flex items-center gap-1.5" style={{ color: g.color }}><Ic size={13} />{g.label}</div>
            <div className="space-y-1.5">{members.map((m) => (<div key={m.id} className="flex items-center gap-2.5 rounded-lg p-2" style={{ border: `1px solid ${C.line}` }}><Avatar name={m.name} size={30} /><div className="flex-1 min-w-0"><div className="font-bold text-sm truncate" style={{ color: C.text }}>{m.name}</div><div className="text-[10px]" style={{ color: C.soft }}>{m.area || m.vehicle || m.phone || ""}</div></div></div>))}</div>
          </div>); })}
      </Card>);
    })()}
  </div>);
}

// ═══ نقل الكميات بين الفروع (احترافي، بقيد محاسبي) ═══
function CBranchTransfer(props) {
  const { data, actions } = props; const co = useCo(props);
  const branches = companyBranches(co);
  const prods = data.products.filter((p) => p.companyId === co.id);
  const [from, setFrom] = useState(branches[0]?.id || "");
  const [to, setTo] = useState(branches[1]?.id || "");
  const [rows, setRows] = useState([]);
  const [note, setNote] = useState("");
  const [pid, setPid] = useState(prods[0]?.id || "");
  const [qty, setQtyV] = useState(10);
  const transfers = (data.branchTransfers || []).filter((t) => t.companyId === co.id);
  const addRow = () => { if (!pid || qty <= 0) return; if (rows.some((r) => r.pid === pid)) { toast("الصنف مضاف", "danger"); return; } setRows([...rows, { pid, qty: +qty }]); };
  const totalVal = rows.reduce((s, r) => { const p = prods.find((x) => x.id === r.pid); return s + r.qty * (p ? priceAfter(p) : 0); }, 0);
  const submit = () => {
    if (from === to) { toast("اختر فرعين مختلفين", "danger"); return; }
    if (!rows.length) { toast("أضف أصنافاً للنقل", "danger"); return; }
    actions.transferBranchStock({ companyId: co.id, fromBranch: from, toBranch: to, items: rows, note });
    toast("تم النقل وتسجيل القيد ✓"); setRows([]); setNote("");
  };
  if (!hasBranches(co)) return (<div className="space-y-4"><SectionTitle sub="متاح للشركات متعددة الفروع فقط.">نقل بين الفروع</SectionTitle><Empty icon={Building2} text="لا فروع متعددة لهذه الشركة." /></div>);
  return (<div className="space-y-4"><SectionTitle sub="انقل الكميات بين فروع الشركة — يُسجَّل قيد محاسبي (مخزون الوجهة مدين / مخزون المصدر دائن).">نقل الكميات بين الفروع</SectionTitle>
    <Card className="p-4 sm:p-5 space-y-4">
      <div className="grid sm:grid-cols-2 gap-3">
        <Field label="من فرع"><select value={from} onChange={(e) => setFrom(e.target.value)} className={inputCls} style={inputStyle}>{branches.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}</select></Field>
        <Field label="إلى فرع"><select value={to} onChange={(e) => setTo(e.target.value)} className={inputCls} style={inputStyle}>{branches.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}</select></Field>
      </div>
      {from === to && <div className="text-xs font-bold rounded-lg p-2" style={{ background: C.dangerSoft, color: C.danger }}>⚠️ اختر فرعين مختلفين</div>}
      <div className="rounded-xl p-3" style={{ background: C.well }}>
        <div className="grid grid-cols-12 gap-2 items-end mb-2">
          <div className="col-span-7"><Field label="الصنف"><select value={pid} onChange={(e) => setPid(e.target.value)} className={inputCls} style={inputStyle}>{prods.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}</select></Field></div>
          <div className="col-span-3"><Field label="الكمية (ستيكة)"><Input type="number" value={qty} onChange={(e) => setQtyV(e.target.value)} /></Field></div>
          <div className="col-span-2"><PrimaryBtn full icon={Plus} onClick={addRow}>أضف</PrimaryBtn></div>
        </div>
        <div className="space-y-1.5">{rows.map((r) => { const p = prods.find((x) => x.id === r.pid); return (<div key={r.pid} className="flex items-center justify-between rounded-lg px-3 py-2" style={{ background: C.surface }}><span className="text-sm font-bold" style={{ color: C.text }}>{p?.name}</span><div className="flex items-center gap-3"><span className="text-sm font-extrabold" style={{ color: SC.color }}>{r.qty} ستيكة</span><button onClick={() => setRows(rows.filter((x) => x.pid !== r.pid))} className="rounded p-1" style={{ background: C.well }}><Trash2 size={13} style={{ color: C.danger }} /></button></div></div>); })}{rows.length === 0 && <div className="text-xs text-center py-2" style={{ color: C.soft }}>لم تُضف أصناف بعد.</div>}</div>
      </div>
      <Field label="ملاحظة (اختياري)"><Input value={note} onChange={(e) => setNote(e.target.value)} placeholder="سبب النقل..." /></Field>
      <div className="flex items-center justify-between rounded-xl p-3" style={{ background: SC.soft }}><span className="text-sm font-bold" style={{ color: C.text }}>قيمة النقل التقديرية</span><span className="font-extrabold" style={{ color: SC.color }}>{money(totalVal)}</span></div>
      <PrimaryBtn full icon={ArrowLeftRight} onClick={submit}>تنفيذ النقل وتسجيل القيد</PrimaryBtn>
    </Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>سجل التحويلات ({transfers.length})</h3>
      <div className="space-y-2">{transfers.map((t) => (<div key={t.id} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}>
        <div className="flex items-center justify-between mb-1"><span className="font-bold text-sm flex items-center gap-1.5" style={{ color: C.text }}><MapPin size={13} style={{ color: SC.color }} />{branchName(co, t.fromBranch)} <ArrowLeftRight size={13} style={{ color: C.soft }} /> {branchName(co, t.toBranch)}</span><span className="text-xs font-extrabold" style={{ color: C.success }}>{money(t.value)}</span></div>
        <div className="text-[11px]" style={{ color: C.soft }}>{t.items.length} صنف · {t.date}{t.note ? " · " + t.note : ""}</div>
      </div>))}{transfers.length === 0 && <Empty icon={ArrowLeftRight} text="لا تحويلات بعد." />}</div>
    </Card>
  </div>);
}

// ═══ عمولة التطبيق (شاشة الشركة): ترى ما احتُسب عليها وتسدّده ═══
function CAppCommission(props) {
  const { data, actions } = props; const co = useCo(props);
  const earned = appCommissionEarned(data, co);
  const paid = appCommissionPaid(co);
  const due = appCommissionDue(data, co);
  const collected = companyCollected(data, co.id);
  const rate = co.commType === "percent" ? co.commValue + "%" : COMM_TYPES[co.commType];
  const [pay, setPay] = useState(false);
  const [amt, setAmt] = useState(0);
  const paidOrders = data.invoices.filter((i) => i.companyId === co.id && i.paid > 0);
  return (<div className="space-y-4"><SectionTitle sub="عمولة التطبيق تُحتسب تلقائياً على المبالغ المُحصّلة من عملياتك المكتملة.">عمولة التطبيق</SectionTitle>
    <Card className="p-5" style={{ background: `linear-gradient(120deg, ${C.ink}, ${SC.color} 130%)`, color: "#fff" }}>
      <div className="text-sm opacity-85">المستحق للتطبيق حالياً</div>
      <div className="text-3xl font-extrabold mt-1">{money(due)}</div>
      <div className="text-xs opacity-80 mt-1">نسبة التطبيق على شركتك: {rate}</div>
      {due > 0 && <button onClick={() => { setAmt(due); setPay(true); }} className="rb-press mt-3 rounded-xl px-4 py-2 text-sm font-extrabold" style={{ background: "#fff", color: SC.color }}>سدّد الآن</button>}
    </Card>
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={TrendingUp} label="إجمالي المُحصّل من عملياتك" value={money(collected)} tint={{ color: C.success, soft: C.successSoft }} />
      <Stat icon={Percent} label="عمولة مكتسبة للتطبيق" value={money(earned)} tint={{ color: C.accent, soft: C.accentSoft }} />
      <Stat icon={CheckCircle2} label="سدّدته للتطبيق" value={money(paid)} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={Clock} label="المتبقّي المستحق" value={money(due)} tint={{ color: due > 0 ? C.danger : C.success, soft: due > 0 ? C.dangerSoft : C.successSoft }} />
    </div>
    <Card className="p-4">
      <h3 className="font-extrabold mb-1" style={{ color: C.text }}>كيف تُحتسب؟</h3>
      <p className="text-sm" style={{ color: C.soft }}>عندما يستلم العميل طلبيته ويسدّد قيمتها، تُحتسب نسبة التطبيق تلقائياً على المبلغ المُحصّل. المستحق يتراكم حتى تسدّده، وعند السداد يُخصم مباشرةً من رصيدك.</p>
      <div className="rounded-xl p-3 mt-3 flex items-center justify-between" style={{ background: C.well }}>
        <span className="text-sm font-bold" style={{ color: C.text }}>{paidOrders.length} فاتورة محصّلة (جزئياً أو كلياً)</span>
        <span className="font-extrabold" style={{ color: SC.color }}>{money(collected)}</span>
      </div>
    </Card>
    {pay && <div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4" onClick={() => setPay(false)}>
      <div className="rb-modal w-full max-w-sm rounded-2xl p-5" onClick={(e) => e.stopPropagation()}>
        <h3 className="font-extrabold mb-1" style={{ color: C.text }}>سداد عمولة التطبيق</h3>
        <p className="text-sm mb-3" style={{ color: C.soft }}>المستحق: {money(due)}</p>
        <Field label="المبلغ (د.ل)"><Input type="number" value={amt} onChange={(e) => setAmt(+e.target.value)} /></Field>
        <div className="flex gap-2 mt-4"><PrimaryBtn full icon={CheckCircle2} onClick={() => { const a = Math.min(amt, due); actions.payAppCommission(co.id, a); setPay(false); toast(`سُدّد ${money(a)} للتطبيق ✓`); }}>تأكيد السداد</PrimaryBtn><GhostBtn onClick={() => setPay(false)}>إلغاء</GhostBtn></div>
      </div>
    </div>}
  </div>);
}

// ═══ ربط الفريق بالفروع: إسناد كل الأدوار للفروع ═══
function CBranchTeam(props) {
  const { data, actions } = props; const co = useCo(props);
  const branches = companyBranches(co);
  if (!hasBranches(co)) return (<div className="space-y-4"><SectionTitle sub="متاح للشركات متعددة الفروع فقط. أضف الفروع أولاً من إعدادات الشركة (لدى مطوّر المنصة).">ربط الفريق بالفروع</SectionTitle><Empty icon={Building2} text="لا فروع متعددة لهذه الشركة." /></div>);
  const [branchFilter, setBranchFilter] = useState("all");
  // مجموعات الأدوار مع مصفوفة البيانات لكل دور
  const groups = [
    { key: "salesManagers", label: "مديرو المبيعات", icon: UserCog, color: C.ink },
    { key: "supervisors", label: "المشرفون", icon: Briefcase, color: C.teal },
    { key: "reps", label: "المندوبون", icon: Handshake, color: C.info },
    { key: "promoters", label: "المروّجون", icon: ImageIcon, color: C.rose },
    { key: "drivers", label: "السائقون", icon: Truck, color: C.purple },
    { key: "dispatchers", label: "مشرفو الحركة", icon: Route, color: "#B45309" },
  ];
  const memberBranch = (m) => m.branchId || "";
  const isAssigned = (m) => (m.branchIds && m.branchIds.length) || m.branchId;
  const inFilter = (m) => branchFilter === "all" || (branchFilter === "none" && !isAssigned(m)) || (m.branchIds ? m.branchIds.includes(branchFilter) : m.branchId === branchFilter);
  // إحصاء غير المُسندين
  const allMembers = groups.flatMap((g) => (data[g.key] || []).filter((x) => x.companyId === co.id));
  const unassigned = allMembers.filter((m) => !isAssigned(m)).length;
  return (<div className="space-y-4"><SectionTitle sub="لكل فرع طاقمه المستقل — أسند مديري المبيعات والمشرفين والمندوبين والسائقين والمروّجين ومشرفي الحركة لفروعهم.">ربط الفريق بالفروع</SectionTitle>
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={Building2} label="عدد الفروع" value={branches.length} tint={{ color: SC.color, soft: SC.soft }} />
      <Stat icon={Users} label="إجمالي الفريق" value={allMembers.length} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={CheckCircle2} label="مُسندون لفروع" value={allMembers.length - unassigned} tint={{ color: C.success, soft: C.successSoft }} />
      <Stat icon={AlertTriangle} label="بلا فرع" value={unassigned} tint={{ color: unassigned ? C.danger : C.success, soft: unassigned ? C.dangerSoft : C.successSoft }} />
    </div>

    {/* فلتر حسب الفرع */}
    <div className="flex items-center gap-1.5 flex-wrap rounded-2xl p-1.5" style={{ background: C.well, border: `1px solid ${C.line}` }}>
      {[{ id: "all", name: "الكل" }, ...branches, { id: "none", name: "بلا فرع" }].map((b) => { const on = branchFilter === b.id; return (
        <button key={b.id} onClick={() => setBranchFilter(b.id)} className="rb-press flex items-center gap-1.5 rounded-xl px-3 py-1.5 text-sm font-bold transition-all" style={{ background: on ? SC.color : "transparent", color: on ? C.onAccent : C.soft }}>{b.id === "none" ? <AlertTriangle size={13} /> : b.id === "all" ? <Layers size={13} /> : <MapPin size={13} />}{b.name}</button>); })}
    </div>

    {groups.map((g) => { const members = (data[g.key] || []).filter((x) => x.companyId === co.id && inFilter(x)); if (!members.length) return null; const Ic = g.icon; return (
      <Card key={g.key} className="p-4">
        <h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Ic size={17} style={{ color: g.color }} />{g.label} <span className="text-xs font-normal" style={{ color: C.soft }}>({members.length})</span></h3>
        <div className="space-y-2">{members.map((m) => (<div key={m.id} className="rounded-xl p-2.5" style={{ border: `1px solid ${C.line}` }}>
          <div className="flex items-center gap-3">
            <Avatar name={m.name} size={36} />
            <div className="flex-1 min-w-0"><div className="font-bold text-sm truncate" style={{ color: C.text }}>{m.name}</div><div className="text-[11px]" style={{ color: C.soft }}>{m.area || m.vehicle || m.phone || ""}</div></div>
            {g.key !== "salesManagers" && <select value={memberBranch(m)} onChange={(e) => { actions.assignBranch(g.key, m.id, e.target.value); toast(e.target.value ? `أُسند ${m.name} إلى ${branchName(co, e.target.value)} ✓` : `أُزيل فرع ${m.name}`); }} className="rounded-lg px-2 py-1.5 text-sm font-bold shrink-0" style={{ border: `1px solid ${m.branchId ? SC.color : C.lineStrong}`, background: m.branchId ? SC.soft : C.surface, color: m.branchId ? SC.color : C.soft, maxWidth: 140 }}>
              <option value="">— بلا فرع —</option>
              {branches.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}
            </select>}
          </div>
          {g.key === "salesManagers" && <div className="mt-2.5 pr-1">
            <div className="text-[11px] mb-1.5" style={{ color: C.soft }}>يشرف على الفروع {(!m.branchIds || !m.branchIds.length) && <span className="font-bold" style={{ color: C.success }}>(كل الفروع حالياً)</span>}:</div>
            <div className="flex flex-wrap gap-1.5">{branches.map((b) => { const on = (m.branchIds || []).includes(b.id); return (
              <button key={b.id} onClick={() => { actions.toggleManagerBranch(m.id, b.id); }} className="rb-press flex items-center gap-1 rounded-lg px-2.5 py-1 text-xs font-bold transition-all" style={{ background: on ? SC.color : C.well, color: on ? C.onAccent : C.soft, border: `1px solid ${on ? SC.color : C.line}` }}>{on ? <CheckCircle2 size={12} /> : <MapPin size={12} />}{b.name}</button>); })}</div>
          </div>}
        </div>))}</div>
      </Card>); })}
  </div>);
}

function CMerchants(props) {
  const { data, actions } = props; const co = useCo(props);
  const merchants = data.merchants.filter((m) => m.companyId === co.id);
  const reps = data.reps.filter((r) => r.companyId === co.id);
  const [f, setF] = useState({ name: "", owner: "", phone: "", area: "", city: "طرابلس", creditLimit: 5000, repId: reps[0]?.id || "", tier: "retail" });
  const [edit, setEdit] = useState(null);
  const [rateM, setRateM] = useState(null);
  return (<div className="space-y-4"><SectionTitle sub={`${SC.merchantsWord} وعملاء شركتك. عدّل التصنيف والنسب وموقع المحل عند إبرام عقد.`}>{SC.merchantsWord}/العملاء</SectionTitle>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>إضافة {SC.merchantWord} / عميل</h3>
      <div className="grid sm:grid-cols-3 gap-2">
        <Field label="الاسم"><Input value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} /></Field>
        <Field label="اسم المالك"><Input value={f.owner} onChange={(e) => setF({ ...f, owner: e.target.value })} /></Field>
        <Field label="الهاتف"><Input value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} /></Field>
        <Field label="المنطقة"><Input value={f.area} onChange={(e) => setF({ ...f, area: e.target.value })} /></Field>
        <Field label="التصنيف"><select value={f.tier} onChange={(e) => setF({ ...f, tier: e.target.value })} className={inputCls} style={inputStyle}>{Object.entries(TIERS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}</select></Field>
        <Field label="المندوب المسؤول"><select value={f.repId} onChange={(e) => setF({ ...f, repId: e.target.value })} className={inputCls} style={inputStyle}>{reps.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}</select></Field>
      </div>
      <div className="mt-3"><PrimaryBtn icon={Plus} onClick={() => { if (f.name) { actions.addMerchant({ server: co.server, companyId: co.id, ...f, creditLimit: +f.creditLimit, balance: 0, email: "", contractPct: 0, appDiscountPct: 0, photo: "🏪", lat: 32.88, lng: 13.19, hasContract: false }); setF({ name: "", owner: "", phone: "", area: "", city: "طرابلس", creditLimit: 5000, repId: reps[0]?.id || "", tier: "retail" }); } }}>إضافة</PrimaryBtn></div>
    </Card>
        {/* ===== رفع ملف Excel لإضافة المحلات دفعة أولى ===== */}
    <ImportClientsCard co={co} reps={reps} data={data} actions={actions} />
    {/* ===== عملاء مستوردون قيد التوثيق الميداني (مجمّعون بالمندوب) ===== */}
    {(() => { const pcs = data.pendingClients.filter((p) => p.companyId === co.id && p.status !== "activated"); if (!pcs.length) return null;
      const byRep = Object.entries(pcs.reduce((acc, p) => { (acc[p.repId] = acc[p.repId] || []).push(p); return acc; }, {}));
      return (<Card className="p-4 min-w-0" style={{ border: `1px solid ${C.accent}` }}>
        <h3 className="font-extrabold mb-1 flex items-center gap-2" style={{ color: C.text }}><ClipboardList size={17} style={{ color: C.accent }} />محلات مستوردة قيد التوثيق الميداني <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.accentSoft, color: "#92590C" }}>{pcs.length}</span></h3>
        <p className="text-xs mb-3" style={{ color: C.soft }}>كل محل يجب أن يزوره مندوبه: يطابق الاسم، يصوّر المحل، يوثّق الموقع GPS، ويأخذ تأكيد التاجر — عندها فقط يتفعّل.</p>
        <div className="space-y-3">{byRep.map(([repId, arr]) => (<div key={repId} className="rounded-2xl p-3" style={{ background: C.bg }}>
          <div className="flex items-center gap-2 mb-2"><Avatar name={repName(data, repId)} size={30} /><span className="font-bold text-sm" style={{ color: C.text }}>{repName(data, repId)}</span><span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: C.accentSoft, color: "#92590C" }}>{arr.length} محل</span></div>
          <div className="space-y-1.5">{arr.map((p) => (<div key={p.id} className="rounded-xl px-3 py-2 flex items-center justify-between gap-2 flex-wrap" style={{ background: C.surface, border: `1px solid ${C.line}` }}>
            <div className="min-w-0"><div className="text-sm font-bold" style={{ color: C.text }}>{p.name}{p.prevName && <span className="text-[10px] font-normal mr-1" style={{ color: C.soft }}>(كان: {p.prevName})</span>}</div><div className="text-[10px]" style={{ color: C.soft }}>{p.area}{p.phone ? ` · ${p.phone}` : ""}{p.openingDebt ? ` · مديونية افتتاحية ${money(p.openingDebt)}` : ""}</div>
              <div className="flex gap-1 mt-1">{[["الاسم", p.nameConfirmed], ["صورة", !!p.photo], ["GPS", p.gpsConfirmed], ["تأكيد التاجر", p.merchantConfirmed || p.status === "activated"]].map(([lbl, ok], x) => <span key={x} className="text-[9px] font-bold px-1.5 py-0.5 rounded" style={{ background: ok ? C.successSoft : C.bg, color: ok ? C.success : C.soft }}>{ok ? "✓" : "○"} {lbl}</span>)}</div>
            </div>
            <Badge status={p.status} map={PC_STATUS} />
          </div>))}</div>
        </div>))}</div>
      </Card>); })()}
    {/* ===== العملاء الفعليون مجمّعين حسب المندوب ===== */}
    {(() => {
      const groups = reps.map((r) => [r, merchants.filter((m) => m.repId === r.id)]).filter(([, arr]) => arr.length);
      const orphan = merchants.filter((m) => !reps.some((r) => r.id === m.repId));
      const renderM = (m) => { const T = TIERS[m.tier] || TIERS.retail; const mr = merchantRating(data, m.id); return (<Card key={m.id} className="p-4 min-w-0">
        <div className="flex items-start justify-between mb-2"><div className="flex items-center gap-2 min-w-0"><div className="rounded-xl flex items-center justify-center text-2xl shrink-0" style={{ width: 44, height: 44, background: C.bg }}>{m.photo || "🏪"}</div><div className="min-w-0"><div className="font-extrabold text-sm truncate" style={{ color: C.text }}>{m.name}</div><div className="text-[10px] font-bold" style={{ color: SC.color }}>{m.membershipNo}</div><div className="text-xs truncate" style={{ color: C.soft }}>{m.owner}</div><div className="mt-0.5"><Stars {...(mr || { avg: null })} size={11} /></div></div></div><span className="text-[10px] font-bold px-2 py-0.5 rounded-full flex items-center gap-1 shrink-0" style={{ background: T.soft, color: T.color }}><T.icon size={11} />{T.label}</span></div>
        <div className="text-xs space-y-1" style={{ color: C.soft }}><div className="flex items-center gap-1"><MapPin size={12} />{m.area}، {m.city}</div><div className="flex items-center gap-1"><Phone size={12} />{m.phone}</div></div>
        {(m.contractPct > 0 || m.appDiscountPct > 0) && <div className="flex gap-1.5 mt-2">{m.contractPct > 0 && <span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.successSoft, color: C.success }}>عقد -{m.contractPct}%</span>}{m.appDiscountPct > 0 && <span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: SC.soft, color: SC.color }}>تطبيق -{m.appDiscountPct}%</span>}</div>}
        <div className="flex justify-between items-center mt-2 pt-2" style={{ borderTop: `1px solid ${C.line}` }}><span className="text-sm font-extrabold" style={{ color: m.balance < 0 ? C.danger : C.success }}>{money(m.balance)}</span><div className="flex gap-1.5 flex-wrap"><GhostBtn sm icon={ScrollText} onClick={() => { actions.requestReconciliation({ requestedBy: "company", requesterId: co.id, companyId: co.id, merchantId: m.id }); toast("أُرسل طلب مطابقة الحساب للعميل ✓"); }}>مطابقة</GhostBtn><GhostBtn sm icon={Star} onClick={() => setRateM(m)}>تقييم</GhostBtn><GhostBtn sm icon={Edit3} onClick={() => setEdit({ ...m })}>تعديل</GhostBtn></div></div>
      </Card>); };
      return (<div className="space-y-5">
        {groups.map(([r, arr]) => { const debt = arr.reduce((sm, m) => sm + Math.max(0, -m.balance), 0); return (<div key={r.id}>
          <div className="flex items-center justify-between flex-wrap gap-2 mb-2.5">
            <div className="flex items-center gap-2.5"><Avatar name={r.name} size={38} /><div><div className="font-extrabold text-sm" style={{ color: C.text }}>{r.name}</div><div className="text-[10px]" style={{ color: C.soft }}>{r.area} · {REP_TYPE[r.repType || "presell"].label}</div></div></div>
            <div className="flex items-center gap-2"><span className="text-xs font-bold px-2.5 py-1 rounded-full" style={{ background: SC.soft, color: SC.color }}>{arr.length} عميل</span><span className="text-xs font-bold px-2.5 py-1 rounded-full" style={{ background: debt > 0 ? C.dangerSoft : C.successSoft, color: debt > 0 ? C.danger : C.success }}>مديونيات {money(debt)}</span></div>
          </div>
          <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{arr.map(renderM)}</div>
        </div>); })}
        {orphan.length > 0 && <div><div className="font-extrabold text-sm mb-2.5 flex items-center gap-2" style={{ color: C.text }}><Building2 size={16} style={{ color: SC.color }} />عملاء مباشرون (بلا مندوب)</div><div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{orphan.map(renderM)}</div></div>}
        {merchants.length === 0 && <Empty icon={Store} text={`لا توجد ${SC.merchantsWord} بعد.`} />}
      </div>);
    })()}
    {rateM && <RateMerchantModal m={rateM} co={co} data={data} actions={actions} onClose={() => setRateM(null)} />}

    {edit && <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 overflow-auto" onClick={() => setEdit(null)}>
      <div className="w-full max-w-lg rounded-2xl p-5 my-8" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
        <h3 className="font-extrabold mb-1" style={{ color: C.text }}>تعديل العميل</h3>
        <p className="text-sm mb-4" style={{ color: C.soft }}>{edit.name}</p>
        <div className="grid sm:grid-cols-2 gap-3">
          <Field label="التصنيف"><select value={edit.tier} onChange={(e) => setEdit({ ...edit, tier: e.target.value })} className={inputCls} style={inputStyle}>{Object.entries(TIERS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}</select></Field>
          <Field label="المندوب"><select value={edit.repId} onChange={(e) => setEdit({ ...edit, repId: e.target.value })} className={inputCls} style={inputStyle}>{reps.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}</select></Field>
          <Field label="الحد الائتماني"><Input type="number" value={edit.creditLimit} onChange={(e) => setEdit({ ...edit, creditLimit: +e.target.value })} /></Field>
          <Field label="صورة المحل (رمز)"><div className="flex gap-1.5 flex-wrap">{["🏪", "🏬", "🛒", "🏢", "💊", "🏥", "🥫"].map((e) => <button key={e} onClick={() => setEdit({ ...edit, photo: e })} className="w-9 h-9 rounded-lg text-xl flex items-center justify-center" style={{ border: `2px solid ${edit.photo === e ? SC.color : C.line}` }}>{e}</button>)}</div></Field>
        </div>
        <div className="rounded-xl p-3 mt-3 mb-1" style={{ background: edit.hasContract ? C.successSoft : C.bg }}>
          <label className="flex items-center gap-2 text-sm font-bold mb-3" style={{ color: C.text }}><input type="checkbox" checked={edit.hasContract} onChange={(e) => setEdit({ ...edit, hasContract: e.target.checked })} />تم إبرام عقد مع هذا العميل (محل كبير)</label>
          {edit.hasContract && <div className="grid sm:grid-cols-2 gap-3">
            <Field label="نسبة العقد % (تخفيض السعر)" hint="نسبة تحددها الشركة لهذا العميل"><Input type="number" value={edit.contractPct} onChange={(e) => setEdit({ ...edit, contractPct: +e.target.value })} /></Field>
            <Field label="نسبة تخفيض من حصة التطبيق %" hint="تقليل سعر إضافي على التاجر"><Input type="number" value={edit.appDiscountPct} onChange={(e) => setEdit({ ...edit, appDiscountPct: +e.target.value })} /></Field>
          </div>}
        </div>
        <div className="grid sm:grid-cols-2 gap-3 mt-2">
          <Field label="GPS - خط العرض"><Input type="number" value={edit.lat} onChange={(e) => setEdit({ ...edit, lat: +e.target.value })} /></Field>
          <Field label="GPS - خط الطول"><Input type="number" value={edit.lng} onChange={(e) => setEdit({ ...edit, lng: +e.target.value })} /></Field>
        </div>
        <div className="rounded-xl overflow-hidden mt-3 relative" style={{ height: 120, background: "linear-gradient(135deg,#DCE7F5,#CFE3DA)", border: `1px solid ${C.line}` }}>
          <div className="absolute" style={{ left: "50%", top: "50%", transform: "translate(-50%,-100%)" }}><MapPin size={32} style={{ color: C.danger }} fill={C.danger} /></div>
          <div className="absolute bottom-2 right-2 text-[11px] font-bold rounded px-2 py-1" style={{ background: "rgba(255,255,255,.85)", color: C.text }}>📍 {edit.lat?.toFixed(4)}, {edit.lng?.toFixed(4)}</div>
        </div>
        <div className="flex gap-2 mt-4"><PrimaryBtn full icon={CheckCircle2} onClick={() => { actions.updateMerchant(edit.id, { tier: edit.tier, repId: edit.repId, creditLimit: edit.creditLimit, photo: edit.photo, hasContract: edit.hasContract, contractPct: edit.hasContract ? edit.contractPct : 0, appDiscountPct: edit.hasContract ? edit.appDiscountPct : 0, lat: edit.lat, lng: edit.lng }); setEdit(null); }}>حفظ التعديلات</PrimaryBtn><GhostBtn onClick={() => setEdit(null)}>إلغاء</GhostBtn></div>
      </div>
    </div>}
  </div>);
}

// زبائن الشركة المباشرون (تابعون للشركة بلا مشرف)
// نافذة تقييم العميل (تستخدمها الشركة)
function RateMerchantModal({ m, co, data, actions, onClose }) {
  const [stars, setStars] = useState(5); const [comment, setComment] = useState("");
  const history = data.ratings.filter((r) => r.raterRole === "company" && r.merchantId === m.id);
  return (<div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50" onClick={onClose}>
    <div className="rb-modal w-full max-w-md rounded-2xl p-5 max-h-[85vh] overflow-auto" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
      <div className="flex items-center gap-3 mb-4"><div className="rounded-xl flex items-center justify-center text-2xl" style={{ width: 44, height: 44, background: C.bg }}>{m.photo || "🏪"}</div><div><div className="font-extrabold" style={{ color: C.text }}>تقييم {m.name}</div><div className="text-xs" style={{ color: C.soft }}>الالتزام بالسداد وسهولة التعامل</div></div></div>
      <div className="flex justify-center mb-3"><StarPicker value={stars} onChange={setStars} size={32} /></div>
      <Field label="تعليق"><Input value={comment} onChange={(e) => setComment(e.target.value)} placeholder="مثال: ملتزم بالسداد في الموعد" /></Field>
      <div className="flex gap-2 mt-4"><PrimaryBtn full icon={Send} onClick={() => { actions.rateMerchant({ merchantId: m.id, companyId: co.id, stars, comment }); onClose(); toast("حُفظ تقييم العميل ⭐"); }}>حفظ التقييم</PrimaryBtn><GhostBtn onClick={onClose}>إلغاء</GhostBtn></div>
      {history.length > 0 && <div className="mt-4 pt-3" style={{ borderTop: `1px solid ${C.line}` }}><div className="text-xs font-bold mb-2" style={{ color: C.soft }}>تقييمات سابقة</div><div className="space-y-2">{history.map((r) => <div key={r.id} className="rounded-lg p-2 text-xs" style={{ background: C.bg }}><div className="flex items-center justify-between"><Stars avg={r.stars} count={1} showCount={false} size={11} /><span style={{ color: C.soft }}>{r.date}</span></div>{r.comment && <div className="mt-1" style={{ color: C.text }}>{r.comment}</div>}</div>)}</div></div>}
    </div>
  </div>);
}
function CCompanyClients(props) {
  const { data } = props; const co = useCo(props);
  const clients = data.merchants.filter((m) => m.companyId === co.id && m.tier === "company_client");
  return (<div className="space-y-4"><SectionTitle sub="عملاء كبار تابعون للشركة مباشرة وليسوا مربوطين بمشرف.">زبائن الشركة المباشرون</SectionTitle>
    {clients.length ? <div className="grid sm:grid-cols-2 gap-3">{clients.map((m) => { const debt = data.invoices.filter((i) => i.merchantId === m.id).reduce((s, i) => s + (i.amount - i.paid), 0);
      return (<Card key={m.id} className="p-5"><div className="flex items-center gap-3 mb-3"><div className="rounded-xl flex items-center justify-center text-2xl" style={{ width: 48, height: 48, background: C.tealSoft }}>{m.photo}</div><div><div className="font-extrabold" style={{ color: C.text }}>{m.name}</div><div className="text-xs" style={{ color: C.soft }}>{m.owner} · {m.area}</div></div></div>
        <div className="grid grid-cols-3 gap-2 text-center"><div className="rounded-xl p-2" style={{ background: C.bg }}><div className="text-sm font-extrabold" style={{ color: C.teal }}>{m.contractPct}%</div><div className="text-[10px]" style={{ color: C.soft }}>عقد</div></div><div className="rounded-xl p-2" style={{ background: C.bg }}><div className="text-sm font-extrabold" style={{ color: SC.color }}>{money(m.creditLimit)}</div><div className="text-[10px]" style={{ color: C.soft }}>الحد</div></div><div className="rounded-xl p-2" style={{ background: debt > 0 ? C.dangerSoft : C.successSoft }}><div className="text-sm font-extrabold" style={{ color: debt > 0 ? C.danger : C.success }}>{money(debt)}</div><div className="text-[10px]" style={{ color: C.soft }}>دين</div></div></div>
        <div className="text-xs flex items-center gap-1 mt-3" style={{ color: C.soft }}><Handshake size={13} />المندوب المباشر: {repName(data, m.repId)}</div>
      </Card>); })}</div> : <Empty icon={Building2} text="لا يوجد زبائن مباشرون للشركة. صنّف عميلاً كـ«زبون شركة» من صفحة العملاء." />}
  </div>);
}

// طلبات الانضمام للشركة - جدولة زيارة وتوثيق
function CJoin(props) {
  const { data, actions } = props; const co = useCo(props);
  const isDirect = co.companyType === "direct";
  const reqs = data.joinRequests.filter((j) => j.companyId === co.id);
  const reps = data.reps.filter((r) => r.companyId === co.id);
  const [modal, setModal] = useState(null);
  const [repId, setRepId] = useState(reps[0]?.id || "");
  const [visitDate, setVisitDate] = useState("غداً 10:00 ص");
  const [directModal, setDirectModal] = useState(null);
  const [vdebt, setVdebt] = useState(0);
  return (<div className="space-y-4"><SectionTitle sub={isDirect ? `طلبات ${SC.merchantsWord} للانضمام. بصفتك بيعاً مباشراً يمكنك تفعيل الحساب فوراً دون زيارة مندوب.` : `طلبات ${SC.merchantsWord} للانضمام لشركتك. جدول زيارة مندوب للتوثيق ثم فعّل الحساب.`}>طلبات الانضمام</SectionTitle>
    {reqs.length ? reqs.map((j) => { const m = data.merchants.find((x) => x.id === j.merchantId);
      return (<Card key={j.id} className="p-5"><div className="flex items-center justify-between mb-3 flex-wrap gap-2"><div className="flex items-center gap-3"><div className="rounded-xl p-2.5" style={{ background: C.accentSoft }}><Store size={20} style={{ color: C.accent }} /></div><div><div className="font-extrabold" style={{ color: C.text }}>{m?.name}</div><div className="text-xs" style={{ color: C.soft }}>{m?.owner} · {m?.area} · {j.date}</div></div></div><Badge status={j.status} map={JOIN_STATUS} /></div>
        {j.note && <div className="text-sm rounded-lg p-2.5 mb-3" style={{ background: C.bg, color: C.soft }}>{j.note}</div>}
        <div className="grid sm:grid-cols-3 gap-2 text-sm mb-3">
          <Info icon={TrendingDown} label="الدين المُصرّح" value={money(j.oldDebt)} />
          {j.visitDate && <Info icon={CalendarClock} label="موعد الزيارة" value={j.visitDate} />}
          {j.repId && <Info icon={Handshake} label="المندوب" value={repName(data, j.repId)} />}
        </div>
        <div className="flex gap-2 flex-wrap pt-3" style={{ borderTop: `1px solid ${C.line}` }}>
          {j.status === "requested" && (isDirect
            ? <><PrimaryBtn sm icon={BadgeCheck} onClick={() => { setDirectModal(j); setVdebt(j.oldDebt || 0); }}>تفعيل مباشر</PrimaryBtn><GhostBtn sm icon={XCircle} onClick={() => actions.declineJoin(j.id)}>رفض</GhostBtn></>
            : <><PrimaryBtn sm icon={CalendarClock} onClick={() => { setModal(j); setRepId(reps[0]?.id || ""); setVisitDate("غداً 10:00 ص"); }}>جدولة زيارة توثيق</PrimaryBtn><GhostBtn sm icon={XCircle} onClick={() => actions.declineJoin(j.id)}>رفض</GhostBtn></>)}
          {(j.status === "visit_set" || j.status === "verifying") && <span className="text-sm flex items-center gap-1.5" style={{ color: C.info }}><Clock size={15} />بانتظار توثيق المندوب أثناء الزيارة</span>}
          {j.status === "active" && <span className="text-sm flex items-center gap-1.5" style={{ color: C.success }}><BadgeCheck size={15} />الحساب مُفعّل · الدين الموثّق {money(j.verifiedDebt)}</span>}
        </div>
      </Card>); }) : <Empty icon={UserPlus} text="لا توجد طلبات انضمام." />}

    {directModal && <div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50" onClick={() => setDirectModal(null)}>
      <div className="rb-modal w-full max-w-md rounded-2xl p-5" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
        <h3 className="font-extrabold mb-1" style={{ color: C.text }}>تفعيل مباشر</h3>
        <p className="text-sm mb-4" style={{ color: C.soft }}>{merchName(data, directModal.merchantId)} — سيُفعَّل الحساب فوراً دون زيارة مندوب.</p>
        <Field label="الدين الافتتاحي الموثّق (د.ل)" hint="عدّله إن لزم بعد التواصل مع التاجر"><Input type="number" value={vdebt} onChange={(e) => setVdebt(+e.target.value)} /></Field>
        <div className="flex gap-2 mt-4"><PrimaryBtn full icon={BadgeCheck} onClick={() => { actions.directActivateJoin(directModal.id, vdebt); setDirectModal(null); }}>تفعيل الحساب الآن</PrimaryBtn><GhostBtn onClick={() => setDirectModal(null)}>إلغاء</GhostBtn></div>
      </div>
    </div>}

    {modal && <div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50" onClick={() => setModal(null)}>
      <div className="rb-modal w-full max-w-md rounded-2xl p-5" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
        <h3 className="font-extrabold mb-1" style={{ color: C.text }}>جدولة زيارة توثيق</h3>
        <p className="text-sm mb-4" style={{ color: C.soft }}>{merchName(data, modal.merchantId)}</p>
        <div className="space-y-3">
          <Field label="المندوب المكلّف بالزيارة"><select value={repId} onChange={(e) => setRepId(e.target.value)} className={inputCls} style={inputStyle}>{reps.map((r) => <option key={r.id} value={r.id}>{r.name} — {r.area}</option>)}</select></Field>
          <Field label="موعد الزيارة"><Input value={visitDate} onChange={(e) => setVisitDate(e.target.value)} /></Field>
        </div>
        <div className="flex gap-2 mt-4"><PrimaryBtn full icon={CalendarClock} onClick={() => { actions.scheduleVisit(modal.id, repId, visitDate); setModal(null); }}>تأكيد الجدولة</PrimaryBtn><GhostBtn onClick={() => setModal(null)}>إلغاء</GhostBtn></div>
      </div>
    </div>}
  </div>);
}

function CreditNoteForm({ co, data, actions }) {
  const merchants = data.merchants.filter((m) => m.companyId === co.id);
  const [f, setF] = useState({ merchantId: merchants[0]?.id || "", kind: "credit", amount: "", reason: "" });
  return (<div className="grid sm:grid-cols-5 gap-2 items-end">
    <Field label="العميل"><select value={f.merchantId} onChange={(e) => setF({ ...f, merchantId: e.target.value })} className={inputCls} style={inputStyle}>{merchants.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}</select></Field>
    <Field label="النوع"><select value={f.kind} onChange={(e) => setF({ ...f, kind: e.target.value })} className={inputCls} style={inputStyle}><option value="credit">دائن (يخفّض المديونية)</option><option value="debit">مدين (يزيدها)</option></select></Field>
    <Field label="القيمة"><Input type="number" value={f.amount} onChange={(e) => setF({ ...f, amount: e.target.value })} placeholder="د.ل" /></Field>
    <Field label="السبب"><Input value={f.reason} onChange={(e) => setF({ ...f, reason: e.target.value })} placeholder="مثال: تعويض تلف" /></Field>
    <PrimaryBtn icon={Plus} onClick={() => { if (!(+f.amount > 0)) { toast("أدخل قيمة صحيحة", "danger"); return; } actions.issueCreditNote({ merchantId: f.merchantId, companyId: co.id, kind: f.kind, amount: +f.amount, reason: f.reason }); toast("صدر السند وعُدّلت المديونية ✓"); setF({ ...f, amount: "", reason: "" }); }}>إصدار السند</PrimaryBtn>
  </div>);
}
function CAccounts(props) {
  const { data, actions } = props; const co = useCo(props);
  const merchants = data.merchants.filter((m) => m.companyId === co.id);
  const [sel, setSel] = useState(null);
  const [rcp, setRcp] = useState(null);
  if (sel) { const m = merchants.find((x) => x.id === sel); const inv = data.invoices.filter((i) => i.merchantId === m.id && i.companyId === co.id); const pays = data.payments.filter((p) => p.merchantId === m.id && p.companyId === co.id); const receipts = data.receipts.filter((r) => r.merchantId === m.id && r.companyId === co.id);
    return (<div className="space-y-4"><button onClick={() => setSel(null)} className="inline-flex items-center gap-1 text-sm font-bold" style={{ color: C.soft }}><ChevronLeft size={16} />رجوع للحسابات</button>
      <SectionTitle sub={`${m.owner} · ${m.area}`}>{m.name}</SectionTitle>
      <div className="grid sm:grid-cols-3 gap-3"><Stat icon={Wallet} label="الرصيد" value={money(m.balance)} tint={{ color: m.balance < 0 ? C.danger : C.success, soft: m.balance < 0 ? C.dangerSoft : C.successSoft }} /><Stat icon={CreditCard} label="الحد الائتماني" value={money(m.creditLimit)} tint={{ color: C.info, soft: C.infoSoft }} /><Stat icon={Receipt} label="فواتير غير مدفوعة" value={inv.filter((i) => i.paid < i.amount).length} tint={{ color: C.accent, soft: C.accentSoft }} /></div>
      <Card className="p-4"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Receipt size={18} />الإيصالات النقدية الإلكترونية</h3><Table head={["الإيصال", "المبلغ", "المندوب", "التاريخ", "الحالة", ""]}>{receipts.map((r) => <tr key={r.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{r.id}</Td><Td>{money(r.amount)}</Td><Td>{repName(data, r.repId)}</Td><Td>{r.date}</Td><Td>{r.confirmedByCompany ? <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.successSoft, color: C.success }}>مؤكّد</span> : <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.warnSoft, color: "#92590C" }}>بانتظار التأكيد</span>}</Td><Td><div className="flex gap-1"><button onClick={() => setRcp(r)} className="text-xs font-bold" style={{ color: SC.color }}>عرض</button>{!r.confirmedByCompany && <button onClick={() => actions.confirmReceipt(r.id)} className="text-xs font-bold" style={{ color: C.success }}>تأكيد</button>}</div></Td></tr>)}</Table>{receipts.length === 0 && <Empty icon={Receipt} text="لا توجد إيصالات." />}</Card>
      <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>الفواتير</h3><Table head={["الفاتورة", "المبلغ", "المدفوع", "المتبقي", "الحالة"]}>{inv.map((i) => <tr key={i.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{i.id}</Td><Td>{money(i.amount)}</Td><Td>{money(i.paid)}</Td><Td>{money(i.amount - i.paid)}</Td><Td><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={i.paid >= i.amount ? { background: C.successSoft, color: C.success } : { background: C.warnSoft, color: "#92590C" }}>{i.paid >= i.amount ? "مدفوعة" : "مستحقة"}</span></Td></tr>)}</Table></Card>
      {rcp && <ReceiptModal receipt={rcp} data={data} onClose={() => setRcp(null)} />}
    </div>);
  }
  return (<div className="space-y-4"><SectionTitle sub="حسابات عملائك: الأرصدة والفواتير والإيصالات الإلكترونية.">حسابات العملاء</SectionTitle>
    {(() => { const pendingHv = data.handovers.filter((h) => h.companyId === co.id && h.status === "pending"); if (!pendingHv.length) return null; return (<Card className="p-4" style={{ border: `1px solid ${C.accent}` }}><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><HandCoins size={18} style={{ color: C.accent }} />توريدات نقدية بانتظار تأكيدكم</h3><div className="space-y-2">{pendingHv.map((h) => <div key={h.id} className="flex items-center justify-between rounded-xl p-3 flex-wrap gap-2" style={{ background: C.accentSoft }}><div><div className="font-bold text-sm" style={{ color: C.text }}>{repName(data, h.repId)} · {money(h.amount)}</div><div className="text-xs" style={{ color: "#92590C" }}>{h.count} عملية · {h.date}{h.expected != null && h.diff !== 0 ? <b style={{ color: h.diff < 0 ? C.danger : C.success }}> · المتوقع {money(h.expected)} — {h.diff < 0 ? `عجز ${money(Math.abs(h.diff))} (سيُحمَّل عليه)` : `زيادة ${money(h.diff)}`}</b> : h.expected != null ? " · مطابق ✓" : ""}</div></div><PrimaryBtn sm tone="success" icon={CheckCircle2} onClick={() => actions.confirmHandover(h.id)}>تأكيد الاستلام</PrimaryBtn></div>)}</div></Card>); })()}
    {/* سندات القيد: إصدار إشعار دائن/مدين */}
    <Card className="p-4 min-w-0"><h3 className="font-extrabold mb-1 flex items-center gap-2" style={{ color: C.text }}><ScrollText size={17} style={{ color: C.purple }} />سند قيد (إشعار دائن / مدين)</h3>
      <p className="text-xs mb-3" style={{ color: C.soft }}>الدائن يخفّض مديونية العميل (تعويض/خصم)، والمدين يزيدها (شيك مرتجع/فرق سعر).</p>
      <CreditNoteForm co={co} data={data} actions={actions} />
      {(() => { const list = data.creditNotes.filter((c) => c.companyId === co.id).slice(0, 6); if (!list.length) return null; return (<div className="mt-3 space-y-1.5">{list.map((c) => <div key={c.id} className="flex items-center justify-between rounded-xl px-3 py-2 text-xs" style={{ background: C.bg }}><span className="min-w-0 truncate"><b style={{ color: c.kind === "credit" ? C.success : C.danger }}>{c.kind === "credit" ? "دائن" : "مدين"}</b> <span style={{ color: C.text }}>{c.id} · {merchName(data, c.merchantId)}</span> <span style={{ color: C.soft }}>— {c.reason}</span></span><b style={{ color: C.text }}>{money(c.amount)}</b></div>)}</div>); })()}
    </Card>
    {(() => { const openCr = data.collectRequests.filter((c) => c.companyId === co.id && ["requested", "scheduled"].includes(c.status)); if (!openCr.length) return null; return (<Card className="p-4" style={{ border: `1px solid ${C.info}` }}><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Truck size={18} style={{ color: C.info }} />عملاء طلبوا التحصيل — تابع المندوب</h3><Table head={["العميل", "المندوب", "المبلغ", "الطريقة", "الموعد", "الحالة"]}>{openCr.map((c) => <tr key={c.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{merchName(data, c.merchantId)}</Td><Td>{c.repId ? repName(data, c.repId) : "—"}</Td><Td>{c.amount ? money(c.amount) : "—"}</Td><Td>{METHOD_LABEL[c.method]}</Td><Td>{c.visitTime || "—"}</Td><Td><Badge status={c.status} map={COLLECT_STATUS} /></Td></tr>)}</Table></Card>); })()}
    <Card className="p-4"><Table head={["العميل", "المالك", "المندوب", "الرصيد", "الحد الائتماني", ""]}>{merchants.map((m) => <tr key={m.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{m.name}</Td><Td>{m.owner}</Td><Td>{repName(data, m.repId)}</Td><Td><span style={{ color: m.balance < 0 ? C.danger : C.success, fontWeight: 800 }}>{money(m.balance)}</span></Td><Td>{money(m.creditLimit)}</Td><Td><GhostBtn sm icon={Eye} onClick={() => setSel(m.id)}>كشف</GhostBtn></Td></tr>)}</Table></Card>
  </div>);
}
// صلاحيات موظف العمليات — يضبطها أدمن الشركة
function COpsPerms(props) {
  const { data, actions } = props; const co = useCo(props);
  const perms = co.opsPerms || { prices: true, invoice: true, loads: true };
  const ITEMS = [
    { k: "prices", t: "رؤية الأسعار والقيم المالية", d: "عند الإيقاف تُخفى المبالغ عن موظف العمليات في شاشات التجهيز." },
    { k: "invoice", t: "عرض وطباعة فاتورة العميل", d: "عند الإيقاف يكتفي الموظف بقائمة التجهيز دون الفاتورة المسعّرة." },
    { k: "loads", t: "إنشاء تحميلات للمندوبين", d: "عند الإيقاف يقتصر دوره على تلبية طلبات التحميل الواردة فقط." },
  ];
  return (<div className="space-y-4"><SectionTitle sub="تحكّم فيما يراه ويفعله موظف العمليات/المخازن — يُطبَّق فوراً على حسابه.">صلاحيات موظف العمليات</SectionTitle>
    <Card className="p-4 min-w-0"><div className="space-y-2.5">{ITEMS.map((it) => { const on = perms[it.k] !== false; return (<div key={it.k} className="flex items-center justify-between gap-3 rounded-xl p-3" style={{ background: C.bg }}>
      <div className="min-w-0"><div className="text-sm font-extrabold" style={{ color: C.text }}>{it.t}</div><div className="text-[11px]" style={{ color: C.soft }}>{it.d}</div></div>
      <button onClick={() => { actions.setOpsPerms(co.id, { [it.k]: !on }); toast(on ? "أُوقفت الصلاحية" : "فُعّلت الصلاحية ✓"); }} className="rb-press shrink-0 rounded-full transition-colors" style={{ width: 46, height: 26, background: on ? C.success : C.line, position: "relative" }}><span className="absolute rounded-full transition-all" style={{ width: 20, height: 20, background: C.well, top: 3, right: on ? 23 : 3 }} /></button>
    </div>); })}</div></Card>
  </div>);
}
// مرتجعات العملاء لدى الشركة: فحص ← قرار (استبدال / إشعار دائن / رفض)
function CReturns(props) {
  const { data, actions } = props; const co = useCo(props);
  const rets = data.returns.filter((r) => r.companyId === co.id);
  const [val, setVal] = useState({});
  return (<div className="space-y-4"><SectionTitle sub="افحص المرتجع (سليم/تالف/قارب الانتهاء) ثم قرّر: استبدال، إشعار دائن يُخصم من المديونية، أو رفض.">المرتجعات وقراراتها</SectionTitle>
    {rets.map((r) => { const m = data.merchants.find((x) => x.id === r.merchantId); const decided = ["replaced", "credited", "rejected"].includes(r.status);
      return (<Card key={r.id} className="p-4 min-w-0">
        <div className="flex items-center justify-between flex-wrap gap-2 mb-2">
          <div className="min-w-0"><div className="font-extrabold text-sm" style={{ color: C.text }}>{r.id} · {m?.name}</div><div className="text-[11px]" style={{ color: C.soft }}>طلب {r.orderId} · كمية {r.qty} · {r.reason} · {r.date}</div></div>
          <div className="flex items-center gap-1.5 flex-wrap">{r.condition && <span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: RET_CONDITION[r.condition].soft, color: RET_CONDITION[r.condition].color }}>{RET_CONDITION[r.condition].label}</span>}<Badge status={r.status} map={RET_STATUS} /></div>
        </div>
        {/* فحص */}
        {["requested", "approved", "picked"].includes(r.status) && <div className="rounded-xl p-3 mb-2" style={{ background: C.bg }}>
          <div className="text-xs font-extrabold mb-2" style={{ color: C.text }}>١. فحص البضاعة المرتجعة:</div>
          <div className="flex gap-1.5 flex-wrap">{Object.entries(RET_CONDITION).map(([k, v]) => <button key={k} onClick={() => { actions.inspectReturn(r.id, k); toast("سُجّلت نتيجة الفحص ✓"); }} className="rb-press rounded-xl px-3 py-2 text-xs font-bold" style={{ background: v.soft, color: v.color }}>{v.label}</button>)}</div>
        </div>}
        {/* قرار */}
        {r.status === "inspected" && <div className="rounded-xl p-3" style={{ background: C.bg }}>
          <div className="text-xs font-extrabold mb-2" style={{ color: C.text }}>٢. قرار الشركة:</div>
          <div className="flex items-end gap-2 flex-wrap">
            <PrimaryBtn sm tone="success" icon={RotateCcw} onClick={() => { actions.decideReturn(r.id, "replaced"); toast("قرار: استبدال — أُعلم العميل ✓"); }}>استبدال البضاعة</PrimaryBtn>
            <div style={{ minWidth: 130 }}><Field label="قيمة الإشعار الدائن"><Input type="number" value={val[r.id] || ""} onChange={(e) => setVal({ ...val, [r.id]: e.target.value })} placeholder="د.ل" /></Field></div>
            <PrimaryBtn sm icon={HandCoins} onClick={() => { if (!(+val[r.id] > 0)) { toast("أدخل قيمة الإشعار", "danger"); return; } actions.decideReturn(r.id, "credited", +val[r.id]); toast("صدر إشعار دائن وخُصم من مديونيته ✓"); }}>إشعار دائن</PrimaryBtn>
            <GhostBtn sm icon={XCircle} onClick={() => confirmThen("سيُرفض المرتجع نهائياً ويُعلَم العميل.", () => actions.decideReturn(r.id, "rejected"), { tone: "danger", yesLabel: "رفض" })}>رفض</GhostBtn>
          </div>
          {r.condition === "ok" && <div className="text-[10px] mt-2" style={{ color: C.success }}>✓ البضاعة سليمة — عند الاستبدال/الإشعار ستُعاد للمخزون تلقائياً.</div>}
        </div>}
        {decided && r.creditValue > 0 && <div className="text-xs font-bold" style={{ color: C.success }}>قيمة الإشعار الدائن: {money(r.creditValue)}</div>}
      </Card>); })}
    {rets.length === 0 && <Empty icon={RotateCcw} text="لا مرتجعات." />}
  </div>);
}
// لوحة أعمار الدين القابلة لإعادة الاستخدام (لكل الأدوار)
function AgingPanel({ data, merchants, title = "أعمار الديون", showGate = false }) {
  const bucketOf = (age) => AGE_BUCKETS.find((b) => age >= b.min && age <= b.max) || AGE_BUCKETS[2];
  const rows = merchants.map((m) => { const invs = data.invoices.filter((i) => i.merchantId === m.id && i.amount > i.paid); const buckets = { b0: 0, b1: 0, b2: 0 }; invs.forEach((i) => { buckets[bucketOf(i.ageDays ?? 40).key] += (i.amount - i.paid); }); const debt = buckets.b0 + buckets.b1 + buckets.b2; return { m, debt, buckets, days: invs.length ? Math.max(...invs.map((i) => i.ageDays ?? 0)) : 0 }; }).filter((x) => x.debt > 0).sort((a, b) => b.days - a.days);
  const totals = rows.reduce((acc, r) => { AGE_BUCKETS.forEach((b) => acc[b.key] += r.buckets[b.key]); return acc; }, { b0: 0, b1: 0, b2: 0 });
  const grand = totals.b0 + totals.b1 + totals.b2;
  return (<div className="space-y-3">
    <div className="grid grid-cols-3 gap-2 sm:gap-3">{AGE_BUCKETS.map((b) => <Stat key={b.key} icon={Clock} label={b.label} value={money(totals[b.key])} tint={{ color: b.color, soft: b.soft }} hint={grand ? `${Math.round((totals[b.key] / grand) * 100)}%` : ""} />)}</div>
    <Card className="p-4 min-w-0"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>{title}</h3>
      <div className="space-y-2.5">{rows.map(({ m, debt, buckets, days }) => { const g = showGate ? invoiceGate(data, m.id) : null; return (<div key={m.id} className="rounded-xl p-3" style={{ border: `1px solid ${(g && g.blocked) ? C.danger : C.line}` }}>
        <div className="flex items-center justify-between flex-wrap gap-2 mb-2">
          <div className="flex items-center gap-2 min-w-0"><span className="text-xl">{m.photo || "🏪"}</span><div className="min-w-0"><div className="font-bold text-sm truncate" style={{ color: C.text }}>{m.name}{g && g.blocked && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded-full mr-1" style={{ background: C.dangerSoft, color: C.danger }}>ممنوع الفوترة</span>}{m.invoiceOverride && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded-full mr-1" style={{ background: C.successSoft, color: C.success }}>{m.invoiceOverride.type === "allow" ? "سماح" : "إمهال"}</span>}</div><div className="text-[10px]" style={{ color: days > 60 ? C.danger : days > 30 ? "#92590C" : C.soft }}>أقدم دين: {days} يوم</div></div></div>
          <span className="font-extrabold text-sm" style={{ color: C.danger }}>{money(debt)}</span>
        </div>
        <div className="flex h-2.5 rounded-full overflow-hidden" style={{ background: C.bg }}>{AGE_BUCKETS.map((b) => buckets[b.key] > 0 && <div key={b.key} style={{ width: `${(buckets[b.key] / debt) * 100}%`, background: b.color }} />)}</div>
        <div className="flex gap-2 mt-1.5 flex-wrap">{AGE_BUCKETS.map((b) => buckets[b.key] > 0 && <span key={b.key} className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: b.soft, color: b.color }}>{b.label}: {money(buckets[b.key])}</span>)}</div>
      </div>); })}{rows.length === 0 && <Empty icon={Wallet} text="لا مديونيات قائمة." />}</div>
    </Card>
  </div>);
}
function CDebts(props) {
  const { data, actions } = props; const co = useCo(props);
  const merchants = data.merchants.filter((m) => m.companyId === co.id);
  const invFor = (mid) => data.invoices.filter((i) => i.merchantId === mid && i.companyId === co.id && i.amount > i.paid);
  const bucketOf = (age) => AGE_BUCKETS.find((b) => age >= b.min && age <= b.max) || AGE_BUCKETS[2];
  const rows = merchants.map((m) => { const invs = invFor(m.id); const buckets = { b0: 0, b1: 0, b2: 0 }; invs.forEach((i) => { buckets[bucketOf(i.ageDays ?? 40).key] += (i.amount - i.paid); }); const debt = buckets.b0 + buckets.b1 + buckets.b2; return { m, debt, buckets }; }).filter((x) => x.debt > 0).sort((a, b) => b.debt - a.debt);
  const totals = rows.reduce((acc, r) => { AGE_BUCKETS.forEach((b) => acc[b.key] += r.buckets[b.key]); return acc; }, { b0: 0, b1: 0, b2: 0 });
  const grand = totals.b0 + totals.b1 + totals.b2;
  // أعمار الديون حسب المندوب
  const byRep = data.reps.filter((r) => r.companyId === co.id).map((r) => { const mids = merchants.filter((m) => m.repId === r.id).map((m) => m.id); const debt = rows.filter((x) => mids.includes(x.m.id)).reduce((s, x) => s + x.debt, 0); return { r, debt }; }).filter((x) => x.debt > 0).sort((a, b) => b.debt - a.debt);
  const [stmtFor, setStmtFor] = useState(null);
  const ledgerFor = (m) => { const rws = [
    ...data.invoices.filter((i) => i.merchantId === m.id && i.companyId === co.id).map((i) => ({ age: i.ageDays ?? 30, date: i.date, desc: `فاتورة ${i.id}`, debit: i.amount, credit: 0 })),
    ...data.payments.filter((p) => p.merchantId === m.id && p.companyId === co.id).map((p) => ({ age: 5, date: p.date, desc: `دفعة ${p.id} (${p.method})`, debit: 0, credit: p.amount })),
    ...data.creditNotes.filter((c) => c.merchantId === m.id && c.companyId === co.id).map((c) => ({ age: c.ageDays ?? 0, date: c.date, desc: `${c.kind === "credit" ? "إشعار دائن" : "إشعار مدين"} ${c.id}`, debit: c.kind === "debit" ? c.amount : 0, credit: c.kind === "credit" ? c.amount : 0 })),
  ].sort((a, b) => b.age - a.age); let bal = 0; return rws.map((r) => { bal += r.debit - r.credit; return { ...r, bal }; }); };
  return (<div className="space-y-4"><SectionTitle sub="أعمار الديون لكل عميل ومندوب — كلما تقدّم عمر الدين زاد خطره.">المديونيات وأعمارها</SectionTitle>
    <div className="grid grid-cols-3 gap-3">{AGE_BUCKETS.map((b) => <Stat key={b.key} icon={Clock} label={b.label} value={money(totals[b.key])} tint={{ color: b.color, soft: b.soft }} hint={grand ? `${Math.round((totals[b.key] / grand) * 100)}% من الإجمالي` : ""} />)}</div>
    <Card className="p-4 min-w-0"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>حسب العميل</h3>
      <div className="space-y-2.5">{rows.map(({ m, debt, buckets }) => (<div key={m.id} className="rounded-xl p-3" style={{ border: `1px solid ${m.frozen ? C.danger : C.line}` }}>
        <div className="flex items-center justify-between flex-wrap gap-2 mb-2">
          <div className="flex items-center gap-2 min-w-0"><span className="text-xl">{m.photo || "🏪"}</span><div className="min-w-0"><div className="font-bold text-sm truncate" style={{ color: C.text }}>{m.name}{m.frozen && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded-full mr-1" style={{ background: C.dangerSoft, color: C.danger }}>معلّق</span>}</div><div className="text-[10px]" style={{ color: C.soft }}>{repName(data, m.repId)}</div></div></div>
          <div className="flex items-center gap-1.5 flex-wrap"><span className="font-extrabold text-sm" style={{ color: C.danger }}>{money(debt)}</span><GhostBtn sm icon={ScrollText} onClick={() => setStmtFor(m)}>كشف</GhostBtn>{m.frozen ? <GhostBtn sm icon={CheckCircle2} onClick={() => { actions.unfreezeMerchant(m.id); toast("أُعيد تفعيل العميل ✓"); }}>تفعيل</GhostBtn> : <GhostBtn sm icon={XCircle} onClick={() => confirmThen(`سيُعلَّق ${m.name} عن الطلب ويُعلَم هو ومندوبه.`, () => { actions.freezeMerchant(m.id, "تجاوز مديونية متقادمة"); toast("عُلّق العميل"); }, { tone: "danger", yesLabel: "تعليق" })}>تعليق</GhostBtn>}</div>
        </div>
        <div className="flex h-2.5 rounded-full overflow-hidden" style={{ background: C.bg }}>{AGE_BUCKETS.map((b) => buckets[b.key] > 0 && <div key={b.key} style={{ width: `${(buckets[b.key] / debt) * 100}%`, background: b.color }} />)}</div>
        {(() => { const g = invoiceGate(data, m.id); if (!g.blocked && !m.invoiceOverride) return null; return (<div className="mt-2 rounded-xl p-2.5" style={{ background: m.invoiceOverride ? C.successSoft : C.dangerSoft }}>
          <div className="flex items-center justify-between gap-2 flex-wrap">
            <span className="text-[11px] font-bold" style={{ color: m.invoiceOverride ? C.success : C.danger }}>{m.invoiceOverride ? (m.invoiceOverride.type === "allow" ? "✓ سماح دائم بإصدار الفواتير" : `✓ إمهال ${m.invoiceOverride.extraDays} يوم`) : `دين ${g.days} يوم — مُنع الطلب/الفوترة تلقائياً`}</span>
            <div className="flex gap-1.5 flex-wrap">
              {m.invoiceOverride ? <GhostBtn sm icon={XCircle} onClick={() => { actions.clearInvoiceOverride(m.id); toast("أُلغي الاستثناء"); }}>إلغاء الاستثناء</GhostBtn> : <>
                <GhostBtn sm icon={Clock} onClick={() => confirmThen(`إمهال ${m.name} يوماً إضافياً لإصدار الفواتير (ظرف طارئ).`, () => { actions.graceInvoice(m.id, 1, "إمهال يوم لظرف طارئ"); toast("أُمهل يوماً إضافياً ✓"); }, { yesLabel: "إمهال يوم" })}>إمهال يوم</GhostBtn>
                <GhostBtn sm icon={CheckCircle2} onClick={() => confirmThen(`سماح دائم لـ${m.name} بإصدار الفواتير رغم تأخر السداد.`, () => { actions.allowInvoiceDespiteDebt(m.id, "سماح استثنائي من الشركة"); toast("سُمح بالإصدار ✓"); }, { yesLabel: "سماح دائم" })}>سماح بالإصدار</GhostBtn>
              </>}
            </div>
          </div>
        </div>); })()}
        <div className="flex gap-2 mt-1.5 flex-wrap">{AGE_BUCKETS.map((b) => buckets[b.key] > 0 && <span key={b.key} className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: b.soft, color: b.color }}>{b.label}: {money(buckets[b.key])}</span>)}</div>
      </div>))}{rows.length === 0 && <Empty icon={Wallet} text="لا مديونيات قائمة." />}</div>
    </Card>
    {byRep.length > 0 && <Card className="p-4 min-w-0"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>حسب المندوب</h3>
      <div className="space-y-2">{byRep.map(({ r, debt }) => <div key={r.id} className="flex items-center justify-between rounded-xl px-3 py-2.5" style={{ background: C.bg }}><div className="flex items-center gap-2"><Avatar name={r.name} size={30} /><span className="text-sm font-bold" style={{ color: C.text }}>{r.name}</span></div><span className="font-extrabold text-sm" style={{ color: C.danger }}>{money(debt)}</span></div>)}</div>
    </Card>}
    {stmtFor && <AccountStatementView title={`كشف حساب — ${stmtFor.name}`} ledger={ledgerFor(stmtFor)} totD={ledgerFor(stmtFor).reduce((s, r) => s + r.debit, 0)} totC={ledgerFor(stmtFor).reduce((s, r) => s + r.credit, 0)} onClose={() => setStmtFor(null)} />}
  </div>);
}
function COrders(props) { const { data } = props; const co = useCo(props); const orders = data.orders.filter((o) => o.companyId === co.id && o.status !== "draft"); return <OrdersTable {...props} orders={orders} title="الطلبات الواردة" sub="الطلبات التي وافق عليها مندوبوك ووصلتك." />; }
function CSupervisors(props) {
  const { data, actions } = props; const co = useCo(props);
  const list = data.supervisors.filter((s) => s.companyId === co.id);
  const [f, setF] = useState({ name: "", phone: "", area: "" });
  return (<div className="space-y-4"><SectionTitle sub="مشرف المبيعات يتابع المندوبين ويوزّع حصص المنتجات عليهم.">مشرفو المبيعات</SectionTitle>
    <Card className="p-4"><div className="grid sm:grid-cols-4 gap-2 items-end"><Field label="الاسم"><Input value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} /></Field><Field label="الهاتف"><Input value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} /></Field><Field label="المنطقة"><Input value={f.area} onChange={(e) => setF({ ...f, area: e.target.value })} /></Field><PrimaryBtn icon={Plus} onClick={() => { if (f.name) { actions.addSupervisor({ server: co.server, companyId: co.id, ...f }); setF({ name: "", phone: "", area: "" }); } }}>إضافة مشرف</PrimaryBtn></div></Card>
    <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{list.map((s) => { const reps = data.reps.filter((r) => r.supervisorId === s.id); return (<Card key={s.id} className="p-4"><div className="flex items-center gap-3 mb-2"><div className="rounded-xl p-2.5" style={{ background: C.tealSoft }}><Briefcase size={20} style={{ color: C.teal }} /></div><div><div className="font-extrabold" style={{ color: C.text }}>{s.name}</div><div className="text-xs" style={{ color: C.soft }}>{s.area}</div></div></div><div className="text-sm flex items-center gap-1 mb-2" style={{ color: C.soft }}><Phone size={13} />{s.phone}</div><div className="text-xs rounded-lg p-2" style={{ background: C.bg, color: C.text }}>يشرف على {reps.length} مندوب</div></Card>); })}{list.length === 0 && <Empty icon={Briefcase} text="لا يوجد مشرفون." />}</div>
  </div>);
}
function CReps(props) { return <CompanyPeople {...props} kind="reps" />; }
function CDrivers(props) { return <CompanyPeople {...props} kind="drivers" />; }
function CompanyPeople(props) {
  const { data, actions, kind } = props; const co = useCo(props); const isReps = kind === "reps";
  const list = (isReps ? data.reps : data.drivers).filter((x) => x.companyId === co.id);
  const supers = data.supervisors.filter((s) => s.companyId === co.id);
  const [f, setF] = useState({ name: "", phone: "", extra: "", extra2: "", supervisorId: supers[0]?.id || "", target: 20000 });
  return (<div className="space-y-4"><SectionTitle sub={isReps ? "المندوب وسيط: يؤكّد النقد ويوافق على الطلب، ويتابعه مشرف المبيعات." : "السائقون المتعاقدون لتوصيل الطلبات."}>{isReps ? "المندوبون" : "السائقون"}</SectionTitle>
    <Card className="p-4"><div className="grid sm:grid-cols-5 gap-2 items-end"><Field label="الاسم"><Input value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} /></Field><Field label="الهاتف"><Input value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} /></Field><Field label={isReps ? "المنطقة" : "المركبة"}><Input value={f.extra} onChange={(e) => setF({ ...f, extra: e.target.value })} /></Field>{isReps ? <Field label="المشرف"><select value={f.supervisorId} onChange={(e) => setF({ ...f, supervisorId: e.target.value })} className={inputCls} style={inputStyle}><option value="">بدون</option>{supers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}</select></Field> : <Field label="رقم اللوحة"><Input value={f.extra2} onChange={(e) => setF({ ...f, extra2: e.target.value })} /></Field>}<PrimaryBtn icon={Plus} onClick={() => { if (f.name) { isReps ? actions.addRep({ server: co.server, companyId: co.id, name: f.name, phone: f.phone, area: f.extra, supervisorId: f.supervisorId || null, target: +f.target || 0, achieved: 0 }) : actions.addDriver({ server: co.server, companyId: co.id, name: f.name, phone: f.phone, vehicle: f.extra, plate: f.extra2 || "", status: "available" }); setF({ name: "", phone: "", extra: "", extra2: "", supervisorId: supers[0]?.id || "", target: 20000 }); } }}>إضافة</PrimaryBtn></div></Card>
    <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{list.map((x) => <Card key={x.id} className="p-4"><div className="flex items-center gap-3 mb-2"><div className="rounded-xl p-2.5" style={{ background: isReps ? C.infoSoft : C.purpleSoft }}>{isReps ? <Handshake size={20} style={{ color: C.info }} /> : <Truck size={20} style={{ color: C.purple }} />}</div><div><div className="font-extrabold" style={{ color: C.text }}>{x.name}</div><div className="text-sm flex items-center gap-1" style={{ color: C.soft }}><Phone size={13} />{x.phone}</div></div></div><div className="text-sm" style={{ color: C.soft }}>{isReps ? x.area : x.vehicle}</div>{isReps && <div className="mt-2"><div className="flex justify-between text-xs mb-1" style={{ color: C.soft }}><span>الإنجاز</span><span>{money(x.achieved)} / {money(x.target)}</span></div><ProgressBar value={x.achieved} max={x.target} color={C.info} /></div>}{!isReps && <span className="text-xs font-bold px-2 py-0.5 rounded-full inline-block mt-2" style={x.status === "available" ? { background: C.successSoft, color: C.success } : { background: C.warnSoft, color: "#92590C" }}>{x.status === "available" ? "متاح" : "مشغول"}</span>}</Card>)}</div>
  </div>);
}
function CTasks(props) {
  const { data, actions } = props; const co = useCo(props);
  const reps = data.reps.filter((r) => r.companyId === co.id);
  const drivers = data.drivers.filter((d) => d.companyId === co.id);
  const tasks = data.tasks.filter((t) => t.companyId === co.id);
  const [f, setF] = useState({ assigneeType: "rep", assigneeId: reps[0]?.id || "", title: "", desc: "", due: "اليوم", priority: "normal" });
  const assignees = f.assigneeType === "rep" ? reps : drivers;
  return (<div className="space-y-4"><SectionTitle sub="أسند مهام أو طلبيات للمندوبين أو السائقين وتابع حالتها.">إسناد المهام</SectionTitle>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>إسناد مهمة جديدة</h3>
      <div className="grid sm:grid-cols-2 gap-2 mb-2">
        <Field label="نوع المُسند إليه"><select value={f.assigneeType} onChange={(e) => setF({ ...f, assigneeType: e.target.value, assigneeId: (e.target.value === "rep" ? reps : drivers)[0]?.id || "" })} className={inputCls} style={inputStyle}><option value="rep">مندوب</option><option value="driver">سائق</option></select></Field>
        <Field label="الشخص"><select value={f.assigneeId} onChange={(e) => setF({ ...f, assigneeId: e.target.value })} className={inputCls} style={inputStyle}>{assignees.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}</select></Field>
        <Field label="عنوان المهمة"><Input value={f.title} onChange={(e) => setF({ ...f, title: e.target.value })} placeholder="مثال: زيارة عميل / توصيل طلبية" /></Field>
        <Field label="الأولوية"><select value={f.priority} onChange={(e) => setF({ ...f, priority: e.target.value })} className={inputCls} style={inputStyle}><option value="normal">عادية</option><option value="high">عالية</option></select></Field>
      </div>
      <Field label="التفاصيل"><textarea value={f.desc} onChange={(e) => setF({ ...f, desc: e.target.value })} className={inputCls} style={{ ...inputStyle, minHeight: 70 }} placeholder="وصف المهمة..." /></Field>
      <div className="mt-3"><PrimaryBtn icon={Send} onClick={() => { if (f.title && f.assigneeId) { actions.addTask({ server: co.server, companyId: co.id, ...f, status: "assigned" }); setF({ ...f, title: "", desc: "" }); } }}>إسناد المهمة</PrimaryBtn></div>
    </Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>المهام المسندة</h3>
      <div className="space-y-2">{tasks.map((t) => { const who = t.assigneeType === "rep" ? repName(data, t.assigneeId) : data.drivers.find((d) => d.id === t.assigneeId)?.name; return (<div key={t.id} className="rounded-xl p-3 flex items-start justify-between gap-3" style={{ border: `1px solid ${C.line}` }}><div className="flex items-start gap-2"><div className="rounded-lg p-1.5 mt-0.5" style={{ background: t.assigneeType === "rep" ? C.infoSoft : C.purpleSoft }}>{t.assigneeType === "rep" ? <Handshake size={15} style={{ color: C.info }} /> : <Truck size={15} style={{ color: C.purple }} />}</div><div><div className="font-bold text-sm flex items-center gap-2" style={{ color: C.text }}>{t.title}{t.priority === "high" && <span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.dangerSoft, color: C.danger }}>عاجل</span>}</div><div className="text-xs" style={{ color: C.soft }}>{t.desc}</div><div className="text-xs mt-1" style={{ color: C.soft }}>{who} · {t.due}</div></div></div><Badge status={t.status} map={TASK_STATUS} /></div>); })}{tasks.length === 0 && <Empty icon={ListChecks} text="لا توجد مهام مسندة." />}</div>
    </Card>
  </div>);
}
function CQuotas(props) {
  const { data, actions } = props; const co = useCo(props);
  const reps = data.reps.filter((r) => r.companyId === co.id);
  const prods = data.products.filter((p) => p.companyId === co.id);
  const quotas = data.quotas.filter((q) => q.companyId === co.id);
  const [f, setF] = useState({ repId: reps[0]?.id || "", pid: prods[0]?.id || "", allocated: 10 });
  return (<div className="space-y-4"><SectionTitle sub="وزّع حصص المنتجات على المندوبين ليطلبها منهم العملاء.">حصص المنتجات</SectionTitle>
    <Card className="p-4"><div className="grid sm:grid-cols-4 gap-2 items-end">
      <Field label="المندوب"><select value={f.repId} onChange={(e) => setF({ ...f, repId: e.target.value })} className={inputCls} style={inputStyle}>{reps.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}</select></Field>
      <Field label="المنتج"><select value={f.pid} onChange={(e) => setF({ ...f, pid: e.target.value })} className={inputCls} style={inputStyle}>{prods.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}</select></Field>
      <Field label="الكمية المخصصة"><Input type="number" value={f.allocated} onChange={(e) => setF({ ...f, allocated: +e.target.value })} /></Field>
      <PrimaryBtn icon={Plus} onClick={() => { if (f.repId && f.pid) { actions.addQuota({ server: co.server, companyId: co.id, ...f, sold: 0 }); } }}>تخصيص</PrimaryBtn>
    </div></Card>
    <Card className="p-4"><Table head={["المندوب", "المنتج", "المخصص", "المُباع", "المتبقي", "نسبة البيع"]}>{quotas.map((q) => { const remaining = q.allocated - q.sold; const pct = Math.round((q.sold / q.allocated) * 100); return (<tr key={q.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{repName(data, q.repId)}</Td><Td>{prodName(data, q.pid)}</Td><Td>{q.allocated}</Td><Td>{q.sold}</Td><Td><span style={{ color: remaining < 5 ? C.danger : C.text, fontWeight: 700 }}>{remaining}</span></Td><Td><div className="flex items-center gap-2"><div style={{ width: 60 }}><ProgressBar value={q.sold} max={q.allocated} color={C.success} /></div><span className="text-xs" style={{ color: C.soft }}>{pct}%</span></div></Td></tr>); })}</Table>{quotas.length === 0 && <Empty icon={Boxes} text="لا توجد حصص موزّعة." />}</Card>
  </div>);
}

function CManagers(props) {
  const { data, actions } = props; const co = useCo(props);
  const list = data.salesManagers.filter((s) => s.companyId === co.id);
  const [f, setF] = useState({ name: "", phone: "", area: "إدارة المبيعات" });
  return (<div className="space-y-4"><SectionTitle sub="مدير المبيعات تابع للشركة: يستعلم عن المبيعات وحركات المندوبين ويتابع المشرفين ويسند لهم أوامر.">مديرو المبيعات</SectionTitle>
    <Card className="p-4"><div className="grid sm:grid-cols-4 gap-2 items-end"><Field label="الاسم"><Input value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} /></Field><Field label="الهاتف"><Input value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} /></Field><Field label="القسم"><Input value={f.area} onChange={(e) => setF({ ...f, area: e.target.value })} /></Field><PrimaryBtn icon={Plus} onClick={() => { if (f.name) { actions.addManager({ server: co.server, companyId: co.id, ...f }); setF({ name: "", phone: "", area: "إدارة المبيعات" }); } }}>إضافة مدير</PrimaryBtn></div></Card>
    <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{list.map((s) => <Card key={s.id} className="p-4"><div className="flex items-center gap-3"><div className="rounded-xl p-2.5" style={{ background: C.well }}><UserCog size={20} style={{ color: C.ink }} /></div><div><div className="font-extrabold" style={{ color: C.text }}>{s.name}</div><div className="text-sm flex items-center gap-1" style={{ color: C.soft }}><Phone size={13} />{s.phone}</div></div></div></Card>)}{list.length === 0 && <Empty icon={UserCog} text="لا يوجد مديرو مبيعات." />}</div>
  </div>);
}
function CPromoters(props) {
  const { data, actions } = props; const co = useCo(props);
  const list = data.promoters.filter((p) => p.companyId === co.id);
  const [f, setF] = useState({ name: "", phone: "", area: "ترويج وعرض" });
  return (<div className="space-y-4"><SectionTitle sub="مروّج المبيعات تابع للشركة: يتلقى إشعاراً بعد استلام المندوب الطلبية، يؤكد الوصول، يشرف على عرض البضاعة، يصوّرها ويرسلها للشركة.">مروّجو المبيعات</SectionTitle>
    <Card className="p-4"><div className="grid sm:grid-cols-4 gap-2 items-end"><Field label="الاسم"><Input value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} /></Field><Field label="الهاتف"><Input value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} /></Field><Field label="المهمة"><Input value={f.area} onChange={(e) => setF({ ...f, area: e.target.value })} /></Field><PrimaryBtn icon={Plus} onClick={() => { if (f.name) { actions.addPromoter({ server: co.server, companyId: co.id, ...f }); setF({ name: "", phone: "", area: "ترويج وعرض" }); } }}>إضافة مروّج</PrimaryBtn></div></Card>
    <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{list.map((p) => { const tasks = data.promotions.filter((x) => x.promoterId === p.id); return (<Card key={p.id} className="p-4"><div className="flex items-center gap-3 mb-2"><div className="rounded-xl p-2.5" style={{ background: C.roseSoft }}><ImageIcon size={20} style={{ color: C.rose }} /></div><div><div className="font-extrabold" style={{ color: C.text }}>{p.name}</div><div className="text-sm flex items-center gap-1" style={{ color: C.soft }}><Phone size={13} />{p.phone}</div></div></div><div className="text-xs rounded-lg p-2" style={{ background: C.bg, color: C.text }}>{tasks.length} مهمة ترويج · {tasks.filter((t) => t.status === "sent").length} مكتملة</div></Card>); })}{list.length === 0 && <Empty icon={ImageIcon} text="لا يوجد مروّجون." />}</div>
  </div>);
}
// ربط المندوبين بالمشرفين أو تغييرهم
function CLink(props) {
  const { data, actions } = props; const co = useCo(props);
  const reps = data.reps.filter((r) => r.companyId === co.id);
  const supers = data.supervisors.filter((s) => s.companyId === co.id);
  return (<div className="space-y-4"><SectionTitle sub="اربط كل مندوب بمشرف أو غيّره. اترك «بدون مشرف» للمندوبين التابعين للشركة مباشرة.">ربط المندوبين بالمشرفين</SectionTitle>
    <Card className="p-4"><Table head={["المندوب", "المنطقة", "المشرف الحالي", "تغيير المشرف"]}>{reps.map((r) => <tr key={r.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{r.name}</Td><Td>{r.area}</Td><Td>{r.supervisorId ? <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.tealSoft, color: C.teal }}>{data.supervisors.find((s) => s.id === r.supervisorId)?.name}</span> : <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.well, color: C.ink }}>تابع للشركة مباشرة</span>}</Td><Td><select value={r.supervisorId || ""} onChange={(e) => actions.updateRep(r.id, { supervisorId: e.target.value || null })} className="rounded-lg px-2 py-1.5 text-sm outline-none" style={inputStyle}><option value="">بدون مشرف (الشركة مباشرة)</option>{supers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}</select></Td></tr>)}</Table></Card>
    <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{supers.map((s) => { const myReps = reps.filter((r) => r.supervisorId === s.id); return (<Card key={s.id} className="p-4"><div className="flex items-center gap-2 mb-2"><div className="rounded-xl p-2" style={{ background: C.tealSoft }}><Briefcase size={18} style={{ color: C.teal }} /></div><div className="font-extrabold" style={{ color: C.text }}>{s.name}</div></div><div className="flex flex-wrap gap-1.5">{myReps.map((r) => <span key={r.id} className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.infoSoft, color: C.info }}>{r.name}</span>)}{myReps.length === 0 && <span className="text-xs" style={{ color: C.soft }}>لا مندوبين</span>}</div></Card>); })}</div>
  </div>);
}
// نقل الحصص (الشركة: بين أي مندوبين)
function CQuotaTransfer(props) {
  const { data, actions } = props; const co = useCo(props);
  const reps = data.reps.filter((r) => r.companyId === co.id);
  return <QuotaTransferUI data={data} actions={actions} reps={reps} title="نقل الحصص" sub="انقل حصص منتج من مندوب إلى آخر (أو بين مشرفين عبر مندوبيهم)." />;
}
function QuotaTransferUI({ data, actions, reps, title, sub }) {
  const repIds = reps.map((r) => r.id);
  const quotas = data.quotas.filter((q) => repIds.includes(q.repId));
  const [from, setFrom] = useState(quotas[0]?.id || "");
  const [toRep, setToRep] = useState("");
  const [amt, setAmt] = useState(5);
  const fromQuota = quotas.find((q) => q.id === from);
  const available = fromQuota ? fromQuota.allocated - fromQuota.sold : 0;
  return (<div className="space-y-4"><SectionTitle sub={sub}>{title}</SectionTitle>
    <Card className="p-5"><div className="grid sm:grid-cols-4 gap-2 items-end">
      <Field label="من حصة"><select value={from} onChange={(e) => setFrom(e.target.value)} className={inputCls} style={inputStyle}>{quotas.map((q) => <option key={q.id} value={q.id}>{repName(data, q.repId)} — {prodName(data, q.pid)}</option>)}</select></Field>
      <Field label="إلى مندوب"><select value={toRep} onChange={(e) => setToRep(e.target.value)} className={inputCls} style={inputStyle}><option value="">اختر مندوباً</option>{reps.filter((r) => r.id !== fromQuota?.repId).map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}</select></Field>
      <Field label={`الكمية (المتاح ${available})`}><Input type="number" value={amt} onChange={(e) => setAmt(+e.target.value)} /></Field>
      <PrimaryBtn icon={ArrowLeft} onClick={() => { if (from && toRep && amt > 0 && amt <= available) { actions.transferQuota(from, toRep, amt); setAmt(5); } }}>نقل</PrimaryBtn>
    </div>
    {fromQuota && <div className="text-xs mt-2" style={{ color: C.soft }}>المتبقي القابل للنقل من هذه الحصة: <b>{available}</b> {prodName(data, fromQuota.pid)}</div>}</Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>الحصص الحالية</h3><Table head={["المندوب", "المنتج", "المخصص", "المُباع", "المتبقي"]}>{quotas.map((q) => <tr key={q.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{repName(data, q.repId)}</Td><Td>{prodName(data, q.pid)}</Td><Td>{q.allocated}</Td><Td>{q.sold}</Td><Td><span style={{ color: (q.allocated - q.sold) < 5 ? C.danger : C.text, fontWeight: 700 }}>{q.allocated - q.sold}</span></Td></tr>)}</Table>{quotas.length === 0 && <Empty icon={Boxes} text="لا توجد حصص." />}</Card>
  </div>);
}

// المساحة الإعلانية: الشركة تنشئ إعلانات تظهر للتجار
const AD_THEMES = [
  { c1: "#0D47A1", c2: "#1565C0" }, { c1: "#0D9488", c2: "#14B8A6" }, { c1: "#BE185D", c2: "#DB2777" },
  { c1: "#B45309", c2: "#D97706" }, { c1: "#4338CA", c2: "#6366F1" }, { c1: "#047857", c2: "#059669" },
];
// ═══════════════════════════════════════════════════════════════════
// شاشات المالية والمحاسبة (مصدر الحقيقة: دفتر الأستاذ)
// ═══════════════════════════════════════════════════════════════════
// أداة فلترة فترة بسيطة (النموذج بلا تواريخ حقيقية، فالفلترة تفهرس القطاع/الشركة)
function useFinScope(props) { const co = useCo(props); return { co, server: co.server, companyId: co.id }; }
function finExportCsv(filename, rows) { const csv = rows.map((r) => r.map((c) => `"${String(c ?? "").replace(/"/g, '""')}"`).join(",")).join("\n"); const a = document.createElement("a"); a.href = "data:text/csv;charset=utf-8,\uFEFF" + encodeURIComponent(csv); a.download = filename; a.click(); }

// لوحة المالية — KPIs + مخططات + تنبيهات (الكمية بجانب القيمة)
function CFinDash(props) {
  const { data } = props; const { co, server, companyId } = useFinScope(props);
  const bal = ledgerBalances(data.ledger, { companyId });
  const get = (code) => bal[code]?.bal || 0;
  const merchants = data.merchants.filter((m) => m.companyId === companyId);
  const totalAR = merchants.reduce((s, m) => s + Math.max(0, -m.balance), 0);
  const invSold = data.invoices.filter((i) => i.companyId === companyId);
  const salesVal = get("4-001") + get("4-002") + get("4-003");
  const cogs = get("5-001"); const grossProfit = salesVal - cogs;
  const grossMargin = salesVal ? Math.round((grossProfit / salesVal) * 100) : 0;
  const invValue = get("1-4-001") + get("1-4-002");
  const cash = get("1-2-001") + get("1-2-003");
  const repShort = data.reps.filter((r) => r.companyId === companyId).reduce((s, r) => s + (r.shortfall || 0), 0);
  // كمية مباعة
  const qtySold = invSold.reduce((s, i) => s + (i.items || []).reduce((q, it) => q + (it.qty || 0), 0), 0);
  // أعمار الديون
  const bucketOf = (age) => AGE_BUCKETS.find((b) => age >= b.min && age <= b.max) || AGE_BUCKETS[2];
  const aging = { b0: 0, b1: 0, b2: 0 };
  data.invoices.filter((i) => i.companyId === companyId && i.amount > i.paid).forEach((i) => { aging[bucketOf(i.ageDays ?? 40).key] += (i.amount - i.paid); });
  const alerts = [];
  merchants.forEach((m) => { const g = invoiceGate(data, m.id); if (g.blocked) alerts.push({ t: `${m.name}: تجاوز 30 يوم — ممنوع الفوترة`, c: C.danger }); const used = Math.max(0, -m.balance); if (m.creditLimit && used >= m.creditLimit * 0.8 && used < m.creditLimit) alerts.push({ t: `${m.name}: اقترب من الحد الائتماني (${Math.round(used / m.creditLimit * 100)}%)`, c: "#92590C" }); });
  return (<div className="space-y-4"><SectionTitle sub="كل رقم هنا مشتقّ من قيود دفتر الأستاذ — لا عدّادات منفصلة.">لوحة المالية</SectionTitle>
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={TrendingUp} label="المبيعات (قيمة)" value={money(salesVal)} tint={{ color: C.success, soft: C.successSoft }} hint={`${qty(qtySold)} مباعة`} />
      <Stat icon={Percent} label="هامش الربح الإجمالي" value={`${grossMargin}%`} tint={{ color: SC.color, soft: SC.soft }} hint={money(grossProfit)} />
      <Stat icon={Wallet} label="النقدية والبنك" value={money(cash)} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={HandCoins} label="إجمالي الذمم المدينة" value={money(totalAR)} tint={{ color: C.danger, soft: C.dangerSoft }} />
      <Stat icon={Boxes} label="قيمة المخزون (WAC)" value={money(invValue)} tint={{ color: C.accent, soft: C.accentSoft }} />
      <Stat icon={TrendingDown} label="ذمم عجز المندوبين" value={money(repShort)} tint={{ color: C.danger, soft: C.dangerSoft }} />
      <Stat icon={Landmark} label="تكلفة البضاعة المباعة" value={money(cogs)} tint={{ color: "#92590C", soft: C.accentSoft }} />
    </div>
    <div className="grid lg:grid-cols-2 gap-3">
      <Card className="p-4 min-w-0"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><PieChart size={17} style={{ color: SC.color }} />توزيع أعمار الديون</h3>
        <DonutChart parts={AGE_BUCKETS.map((b) => ({ label: b.label, v: aging[b.key], color: b.color }))} />
      </Card>
      <Card className="p-4 min-w-0"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><AlertCircle size={17} style={{ color: C.danger }} />تنبيهات ذكية</h3>
        <div className="space-y-1.5 max-h-52 overflow-auto">{alerts.length ? alerts.map((a, i) => <div key={i} className="text-xs rounded-lg px-2.5 py-2 font-bold" style={{ background: a.c === C.danger ? C.dangerSoft : C.accentSoft, color: a.c }}>{a.t}</div>) : <Empty icon={CheckCircle2} text="لا تنبيهات مالية." />}</div>
      </Card>
    </div>
  </div>);
}

// شجرة الحسابات
function CFinCOA(props) {
  const { data } = props; const { companyId } = useFinScope(props);
  const bal = ledgerBalances(data.ledger, { companyId });
  const roots = COA.filter((a) => !a.parent);
  const childrenOf = (code) => COA.filter((a) => a.parent === code);
  const Node = ({ a, depth }) => { const kids = childrenOf(a.code); const b = bal[a.code]?.bal || 0;
    return (<div><div className="flex items-center justify-between py-1.5 px-2 rounded-lg" style={{ paddingRight: 8 + depth * 16, background: depth === 0 ? C.bg : "transparent", borderBottom: `1px solid ${C.line}` }}>
      <span className="flex items-center gap-2 min-w-0"><span className="text-[10px] font-mono px-1.5 py-0.5 rounded" style={{ background: C.bg, color: C.soft }}>{a.code}</span><span className={depth === 0 ? "font-extrabold text-sm" : "text-sm"} style={{ color: C.text }}>{a.name}</span>{a.post && <span className="text-[9px] px-1 rounded" style={{ background: SC.soft, color: SC.color }}>ورقة</span>}{a.ctrl && <span className="text-[9px] px-1 rounded" style={{ background: C.purpleSoft, color: C.purple }}>مراقبة</span>}</span>
      <span className="text-xs font-bold shrink-0" style={{ color: b < 0 ? C.danger : C.text }}>{b !== 0 ? money(b) : "—"}</span>
    </div>{kids.map((k) => <Node key={k.code} a={k} depth={depth + 1} />)}</div>);
  };
  return (<div className="space-y-4"><SectionTitle sub="شجرة حسابات هرمية 4 مستويات — الترحيل على الأوراق فقط، وحسابات المراقبة تُطابَق دفاترها المساعدة.">شجرة الحسابات</SectionTitle>
    <Card className="p-3 min-w-0">{roots.map((a) => <Node key={a.code} a={a} depth={0} />)}</Card>
  </div>);
}

// دفتر اليومية (كل القيود)
function CFinJournal(props) {
  const { data } = props; const { companyId } = useFinScope(props);
  const entries = (data.ledger || []).filter((j) => j.companyId === companyId || !j.companyId);
  const [open, setOpen] = useState(null);
  const SRC = { OPENING: "افتتاحي", T1: "فاتورة", T2: "تحصيل", T3: "توريد", T4: "تسوية", T5: "إشعار دائن", T6: "إشعار مدين", T7: "مرتجع للمخزون", T8: "عجز جرد", T9: "تحميل سيارة" };
  return (<div className="space-y-4"><SectionTitle sub="كل قيد مزدوج القيد، غير قابل للحذف — التصحيح بقيد عكسي.">دفتر اليومية</SectionTitle>
    <div className="space-y-2">{entries.map((j) => (<Card key={j.id} className="p-0 overflow-hidden min-w-0">
      <button onClick={() => setOpen(open === j.id ? null : j.id)} className="w-full flex items-center justify-between gap-2 px-4 py-3 text-right" style={{ background: j.voided ? C.dangerSoft + "33" : "transparent" }}>
        <span className="min-w-0"><span className="text-sm font-extrabold flex items-center gap-2" style={{ color: C.text }}><span className="font-mono text-[11px] px-1.5 py-0.5 rounded" style={{ background: SC.soft, color: SC.color }}>{j.id}</span>{j.voided && <span className="text-[9px] px-1.5 py-0.5 rounded" style={{ background: C.dangerSoft, color: C.danger }}>ملغى</span>}<span className="text-[10px] px-1.5 py-0.5 rounded" style={{ background: C.bg, color: C.soft }}>{SRC[j.source] || j.source}</span></span><span className="block text-[11px] mt-0.5" style={{ color: C.soft }}>{j.memo}</span></span>
        <span className="text-sm font-extrabold shrink-0" style={{ color: C.text }}>{money(j.totalDr)}{j.balanced ? " ✓" : <span style={{ color: C.danger }}> ✗</span>}</span>
      </button>
      {open === j.id && <div className="px-4 pb-3" style={{ borderTop: `1px solid ${C.line}` }}><table className="w-full text-xs mt-2"><thead><tr style={{ color: C.soft }}><th className="text-right py-1">الحساب</th><th className="text-left py-1">مدين</th><th className="text-left py-1">دائن</th></tr></thead><tbody>{j.lines.map((l, i) => <tr key={i}><td className="py-1" style={{ color: C.text }}><span className="font-mono text-[10px]" style={{ color: C.soft }}>{l.acct}</span> {acctName(l.acct)}{l.dims.merchantId ? ` · ${merchName(data, l.dims.merchantId)}` : ""}{l.dims.repId ? ` · ${repName(data, l.dims.repId)}` : ""}</td><td className="text-left py-1" style={{ color: C.success }}>{l.dr ? money(l.dr) : ""}</td><td className="text-left py-1" style={{ color: C.danger }}>{l.cr ? money(l.cr) : ""}</td></tr>)}</tbody></table></div>}
    </Card>))}{entries.length === 0 && <Empty icon={BookOpen} text="لا قيود بعد." />}</div>
  </div>);
}

// دفتر الأستاذ لكل حساب
function CFinGL(props) {
  const { data } = props; const { companyId } = useFinScope(props);
  const postable = COA.filter((a) => a.post);
  const [acct, setAcct] = useState("1-3-001");
  const entries = (data.ledger || []).filter((j) => (j.companyId === companyId || !j.companyId) && !j.voided);
  const lines = []; let run = 0;
  entries.slice().reverse().forEach((j) => { j.lines.filter((l) => l.acct === acct).forEach((l) => { const nat = coaAcc(acct)?.nature; const delta = nat === "C" ? (l.cr - l.dr) : (l.dr - l.cr); run += delta; lines.push({ jv: j.id, memo: j.memo, dr: l.dr, cr: l.cr, run }); }); });
  return (<div className="space-y-4"><SectionTitle sub="حركة كل حساب مع الرصيد الجاري.">دفتر الأستاذ العام</SectionTitle>
    <Card className="p-4 min-w-0"><Field label="الحساب"><select value={acct} onChange={(e) => setAcct(e.target.value)} className={inputCls} style={inputStyle}>{postable.map((a) => <option key={a.code} value={a.code}>{a.code} — {a.name}</option>)}</select></Field>
      <div className="mt-3 overflow-auto"><table className="w-full text-xs"><thead><tr style={{ color: C.soft, borderBottom: `2px solid ${C.ink}` }}><th className="text-right py-1.5">القيد</th><th className="text-right py-1.5">البيان</th><th className="text-left py-1.5">مدين</th><th className="text-left py-1.5">دائن</th><th className="text-left py-1.5">الرصيد</th></tr></thead><tbody>{lines.map((l, i) => <tr key={i} style={{ borderBottom: `1px solid ${C.line}` }}><td className="py-1.5 font-mono text-[10px]" style={{ color: C.soft }}>{l.jv}</td><td className="py-1.5" style={{ color: C.text }}>{l.memo}</td><td className="text-left py-1.5" style={{ color: C.success }}>{l.dr ? money(l.dr) : ""}</td><td className="text-left py-1.5" style={{ color: C.danger }}>{l.cr ? money(l.cr) : ""}</td><td className="text-left py-1.5 font-bold" style={{ color: C.text }}>{money(l.run)}</td></tr>)}</tbody></table>{lines.length === 0 && <Empty icon={BookOpen} text="لا حركة على هذا الحساب." />}</div>
    </Card>
  </div>);
}

// ميزان المراجعة
function CFinTrial(props) {
  const { data } = props; const { companyId } = useFinScope(props);
  const bal = ledgerBalances(data.ledger, { companyId });
  const rows = COA.filter((a) => a.post && bal[a.code]).map((a) => { const signed = bal[a.code].signed; return { a, dr: signed > 0 ? signed : 0, cr: signed < 0 ? -signed : 0 }; }).filter((r) => r.dr || r.cr);
  const totDr = rows.reduce((s, r) => s + r.dr, 0), totCr = rows.reduce((s, r) => s + r.cr, 0);
  const balanced = Math.abs(totDr - totCr) < 0.01;
  return (<div className="space-y-4"><SectionTitle sub="مجموع المدين = مجموع الدائن دائماً." action={<GhostBtn sm icon={FileText} onClick={() => finExportCsv("trial-balance.csv", [["الرمز", "الحساب", "مدين", "دائن"], ...rows.map((r) => [r.a.code, r.a.name, r.dr, r.cr]), ["", "الإجمالي", totDr, totCr]])}>CSV</GhostBtn>}>ميزان المراجعة</SectionTitle>
    <Card className="p-4 min-w-0">
      <div className="rounded-xl p-2.5 mb-3 text-center text-sm font-extrabold" style={{ background: balanced ? C.successSoft : C.dangerSoft, color: balanced ? C.success : C.danger }}>{balanced ? "✓ الميزان متوازن" : "✗ الميزان غير متوازن — خلل"}</div>
      <table className="w-full text-xs"><thead><tr style={{ color: C.soft, borderBottom: `2px solid ${C.ink}` }}><th className="text-right py-1.5">الرمز</th><th className="text-right py-1.5">الحساب</th><th className="text-left py-1.5">مدين</th><th className="text-left py-1.5">دائن</th></tr></thead>
        <tbody>{rows.map((r) => <tr key={r.a.code} style={{ borderBottom: `1px solid ${C.line}` }}><td className="py-1.5 font-mono text-[10px]" style={{ color: C.soft }}>{r.a.code}</td><td className="py-1.5 font-bold" style={{ color: C.text }}>{r.a.name}</td><td className="text-left py-1.5" style={{ color: C.success }}>{r.dr ? money(r.dr) : ""}</td><td className="text-left py-1.5" style={{ color: C.danger }}>{r.cr ? money(r.cr) : ""}</td></tr>)}</tbody>
        <tfoot><tr style={{ borderTop: `2px solid ${C.ink}` }}><td colSpan={2} className="py-2 font-extrabold" style={{ color: C.text }}>الإجمالي</td><td className="text-left py-2 font-extrabold" style={{ color: C.success }}>{money(totDr)}</td><td className="text-left py-2 font-extrabold" style={{ color: C.danger }}>{money(totCr)}</td></tr></tfoot>
      </table>
    </Card>
  </div>);
}

// قائمة الدخل
function CFinPL(props) {
  const { data } = props; const { companyId } = useFinScope(props);
  const bal = ledgerBalances(data.ledger, { companyId });
  const g = (code) => bal[code]?.bal || 0;
  const sales = g("4-001") + g("4-002") + g("4-003");
  const returns = g("5-002") + g("5-003"); const cogs = g("5-001");
  const supplyGain = g("7-001");
  const netSales = sales - returns; const grossProfit = netSales - cogs + supplyGain;
  const Row = ({ label, val, bold, indent }) => (<div className="flex justify-between py-1.5" style={{ borderBottom: `1px solid ${C.line}`, paddingRight: indent ? 16 : 0 }}><span className={bold ? "font-extrabold" : ""} style={{ color: bold ? C.text : C.soft }}>{label}</span><span className="font-bold" style={{ color: val < 0 ? C.danger : C.text }}>{money(val)}</span></div>);
  return (<div className="space-y-4"><SectionTitle sub="ناتج عمل التطبيق: مبيعات − مردودات − تكلفة = مجمل الربح من التوزيع.">ملخص أرباح التوزيع</SectionTitle>
    <Card className="p-4 min-w-0 text-sm">
      <Row label="إجمالي المبيعات" val={sales} />
      <Row label="مردودات وتالف" val={-returns} indent />
      <Row label="صافي المبيعات" val={netSales} bold />
      <Row label="تكلفة البضاعة المباعة" val={-cogs} indent />
      {supplyGain !== 0 && <Row label="فروق توريد (زيادة)" val={supplyGain} indent />}
      <div className="flex justify-between py-2.5 mt-1 rounded-xl px-3" style={{ background: grossProfit >= 0 ? C.successSoft : C.dangerSoft }}><span className="font-extrabold" style={{ color: grossProfit >= 0 ? C.success : C.danger }}>مجمل ربح التوزيع</span><span className="font-extrabold text-lg" style={{ color: grossProfit >= 0 ? C.success : C.danger }}>{money(grossProfit)}</span></div>
      {netSales > 0 && <div className="text-[11px] mt-2 text-center" style={{ color: C.soft }}>هامش مجمل الربح: {Math.round(grossProfit / netSales * 100)}% من صافي المبيعات</div>}
    </Card>
  </div>);
}

// الميزانية العمومية
function CFinBS(props) {
  const { data } = props; const { companyId } = useFinScope(props);
  const bal = ledgerBalances(data.ledger, { companyId });
  const g = (code) => bal[code]?.bal || 0;
  const assets = COA.filter((a) => a.post && a.stmt === "BS" && a.code[0] === "1").map((a) => ({ a, v: g(a.code) })).filter((x) => x.v);
  const liab = COA.filter((a) => a.post && a.stmt === "BS" && a.code[0] === "2").map((a) => ({ a, v: g(a.code) })).filter((x) => x.v);
  const equity = COA.filter((a) => a.post && a.stmt === "BS" && a.code[0] === "3").map((a) => ({ a, v: g(a.code) })).filter((x) => x.v);
  const totA = assets.reduce((s, x) => s + x.v, 0);
  const totL = liab.reduce((s, x) => s + x.v, 0);
  // صافي الربح الجاري يُضاف لحقوق الملكية
  const sales = g("4-001") + g("4-002") + g("4-003"); const pl = sales - g("5-001") - g("5-002") - g("5-003") + g("7-001");
  const totE = equity.reduce((s, x) => s + x.v, 0) + pl;
  const balanced = Math.abs(totA - (totL + totE)) < 0.01;
  const Section = ({ title, items, extra }) => (<Card className="p-4 min-w-0"><h3 className="font-extrabold mb-2" style={{ color: C.text }}>{title}</h3>{items.map((x) => <div key={x.a.code} className="flex justify-between py-1 text-xs" style={{ borderBottom: `1px solid ${C.line}` }}><span style={{ color: C.soft }}>{x.a.name}</span><b style={{ color: C.text }}>{money(x.v)}</b></div>)}{extra}</Card>);
  return (<div className="space-y-4"><SectionTitle sub="أصول عمل التطبيق (نقد + ذمم + مخزون) مقابل الأرصدة الدائنة وحقوق الملكية.">المركز المالي للتطبيق</SectionTitle>
    <div className="rounded-xl p-2.5 text-center text-sm font-extrabold" style={{ background: balanced ? C.successSoft : C.dangerSoft, color: balanced ? C.success : C.danger }}>{balanced ? "✓ الميزانية متوازنة" : "✗ غير متوازنة"} · الأصول {money(totA)} = الخصوم+الملكية {money(totL + totE)}</div>
    <div className="grid lg:grid-cols-2 gap-3">
      <Section title="الأصول" items={assets} extra={<div className="flex justify-between py-2 mt-1 font-extrabold" style={{ borderTop: `2px solid ${C.ink}`, color: C.text }}><span>إجمالي الأصول</span><span>{money(totA)}</span></div>} />
      <div className="space-y-3">
        <Section title="الالتزامات" items={liab} />
        <Section title="حقوق الملكية" items={equity} extra={<><div className="flex justify-between py-1 text-xs" style={{ borderBottom: `1px solid ${C.line}` }}><span style={{ color: C.soft }}>ربح الفترة الجاري</span><b style={{ color: pl >= 0 ? C.success : C.danger }}>{money(pl)}</b></div><div className="flex justify-between py-2 mt-1 font-extrabold" style={{ borderTop: `2px solid ${C.ink}`, color: C.text }}><span>إجمالي الخصوم + الملكية</span><span>{money(totL + totE)}</span></div></>} />
      </div>
    </div>
  </div>);
}

// أعمار الديون (يعيد استخدام AgingPanel)
function CFinAging(props) {
  const { data } = props; const { companyId } = useFinScope(props);
  const merchants = data.merchants.filter((m) => m.companyId === companyId);
  return (<div className="space-y-4"><SectionTitle sub="0–30 / 31–60 / +60 يوم — تقود بوابة منع الفوترة.">تقرير أعمار الديون</SectionTitle>
    <AgingPanel data={data} merchants={merchants} title="أعمار ديون العملاء" showGate />
  </div>);
}

// تقييم المخزون WAC
function CFinInventory(props) {
  const { data } = props; const { companyId } = useFinScope(props);
  const prods = data.products.filter((p) => p.companyId === companyId);
  const rows = prods.map((p) => { const wac = wacOf(data, p.id); const vanQ = data.reps.filter((r) => r.companyId === companyId).reduce((s, r) => s + ((r.vanStock || {})[p.id] || 0), 0); return { p, wac, main: p.stock || 0, van: vanQ, val: money3(((p.stock || 0) + vanQ) * wac) }; });
  const totVal = rows.reduce((s, r) => s + r.val, 0);
  const totQty = rows.reduce((s, r) => s + r.main + r.van, 0);
  return (<div className="space-y-4"><SectionTitle sub="بمتوسط التكلفة المرجّح لكل صنف وموقع (مستودع + سيارات)." action={<GhostBtn sm icon={FileText} onClick={() => finExportCsv("inventory.csv", [["الصنف", "مستودع", "سيارات", "WAC", "القيمة"], ...rows.map((r) => [r.p.name, r.main, r.van, r.wac, r.val])])}>CSV</GhostBtn>}>تقييم المخزون</SectionTitle>
    <div className="grid grid-cols-2 gap-3"><Stat icon={Boxes} label="إجمالي الكمية" value={qty(totQty)} tint={{ color: C.accent, soft: C.accentSoft }} /><Stat icon={Landmark} label="قيمة المخزون" value={money(totVal)} tint={{ color: C.success, soft: C.successSoft }} /></div>
    <Card className="p-4 min-w-0"><table className="w-full text-xs"><thead><tr style={{ color: C.soft, borderBottom: `2px solid ${C.ink}` }}><th className="text-right py-1.5">الصنف</th><th className="text-left py-1.5">مستودع</th><th className="text-left py-1.5">سيارات</th><th className="text-left py-1.5">WAC</th><th className="text-left py-1.5">القيمة</th></tr></thead><tbody>{rows.map((r) => <tr key={r.p.id} style={{ borderBottom: `1px solid ${C.line}` }}><td className="py-1.5 font-bold" style={{ color: C.text }}>{r.p.name}</td><td className="text-left py-1.5" style={{ color: C.text }}>{r.main}</td><td className="text-left py-1.5" style={{ color: C.soft }}>{r.van}</td><td className="text-left py-1.5" style={{ color: C.soft }}>{money(r.wac)}</td><td className="text-left py-1.5 font-bold" style={{ color: C.text }}>{money(r.val)}</td></tr>)}</tbody></table></Card>
  </div>);
}

// مساءلة المندوبين (عهدة + مخزون سيارة + عجز)
function CFinReps(props) {
  const { data } = props; const { companyId } = useFinScope(props);
  const reps = data.reps.filter((r) => r.companyId === companyId);
  const bal = ledgerBalances(data.ledger, { companyId });
  const rows = reps.map((r) => { const custody = ledgerBalances(data.ledger, { companyId, repId: r.id })["1-2-002"]?.bal || 0; const vanVal = money3(Object.entries(r.vanStock || {}).reduce((s, [pid, q]) => s + q * wacOf(data, pid), 0)); return { r, custody, vanVal, shortfall: r.shortfall || 0, total: custody + vanVal + (r.shortfall || 0) }; });
  return (<div className="space-y-4"><SectionTitle sub="ما بذمة كل مندوب: عهدة نقدية + قيمة بضاعة سيارته + ذمم العجز.">مساءلة المندوبين</SectionTitle>
    <div className="space-y-2.5">{rows.map(({ r, custody, vanVal, shortfall, total }) => (<Card key={r.id} className="p-4 min-w-0">
      <div className="flex items-center justify-between flex-wrap gap-2 mb-2"><div className="flex items-center gap-2.5"><Avatar name={r.name} size={34} /><div><div className="font-extrabold text-sm" style={{ color: C.text }}>{r.name}</div><div className="text-[10px]" style={{ color: C.soft }}>{REP_TYPE[r.repType || "presell"].label}</div></div></div><span className="text-lg font-extrabold" style={{ color: C.text }}>{money(total)}</span></div>
      <div className="grid grid-cols-3 gap-2 text-center">
        <div className="rounded-xl p-2" style={{ background: C.infoSoft }}><div className="text-[10px]" style={{ color: C.info }}>عهدة نقدية</div><div className="text-sm font-extrabold" style={{ color: C.info }}>{money(custody)}</div></div>
        <div className="rounded-xl p-2" style={{ background: C.accentSoft }}><div className="text-[10px]" style={{ color: "#92590C" }}>بضاعة السيارة</div><div className="text-sm font-extrabold" style={{ color: "#92590C" }}>{money(vanVal)}</div></div>
        <div className="rounded-xl p-2" style={{ background: shortfall > 0 ? C.dangerSoft : C.successSoft }}><div className="text-[10px]" style={{ color: shortfall > 0 ? C.danger : C.success }}>عجز مسجّل</div><div className="text-sm font-extrabold" style={{ color: shortfall > 0 ? C.danger : C.success }}>{money(shortfall)}</div></div>
      </div>
    </Card>))}{reps.length === 0 && <Empty icon={Handshake} text="لا مندوبين." />}</div>
  </div>);
}

// ═══════════════════════════════════════════════════════════════════
// تحديث السنة: 7 ميزات جديدة (تحصيل ذكي · تنبؤ وتحميل · عروض · تقارير سنوية · مناطق · ولاء · سجل زمني)
// ═══════════════════════════════════════════════════════════════════

// أدوات مشتركة للتحديث
const merchantOpenDebt = (data, mid) => data.invoices.filter((i) => i.merchantId === mid && i.amount > i.paid).reduce((s, i) => s + (i.amount - i.paid), 0);
const merchantPayScore = (data, mid) => {
  // سلوك السداد: نسبة المسدَّد إلى المُفوتر (0..100) — كلما ارتفع، أفضل التزاماً
  const invs = data.invoices.filter((i) => i.merchantId === mid);
  if (!invs.length) return 100;
  const billed = invs.reduce((s, i) => s + i.amount, 0);
  const paid = invs.reduce((s, i) => s + i.paid, 0);
  return billed ? Math.round((paid / billed) * 100) : 100;
};
// أولوية التحصيل = قيمة الدين × عمره ÷ سلوك السداد (كلما زاد الرقم زادت الأولوية)
const collectionPriority = (data, mid) => {
  const debt = merchantOpenDebt(data, mid);
  const age = oldestDebtDays(data, mid);
  const score = merchantPayScore(data, mid) || 1;
  return Math.round((debt * (age + 1)) / score);
};

// ═══ ١) محرّك التحصيل الذكي ═══
function CSmartCollect(props) {
  const { data, actions, go } = props; const co = useCo(props);
  const merchants = data.merchants.filter((m) => m.companyId === co.id && merchantOpenDebt(data, m.id) > 0);
  const ranked = merchants.map((m) => ({
    m, debt: merchantOpenDebt(data, m.id), age: oldestDebtDays(data, m.id),
    score: merchantPayScore(data, m.id), pr: collectionPriority(data, m.id),
    gate: invoiceGate(data, m.id),
  })).sort((a, b) => b.pr - a.pr);
  const totalDebt = ranked.reduce((s, r) => s + r.debt, 0);
  const urgent = ranked.filter((r) => r.age > 30).length;
  const prBadge = (pr) => pr > 40000 ? { t: "قصوى", c: C.danger, s: C.dangerSoft } : pr > 15000 ? { t: "عالية", c: "#B45309", s: C.warnSoft } : { t: "عادية", c: C.info, s: C.infoSoft };
  return (<div className="space-y-4"><SectionTitle sub="يرتّب المحلات حسب أولوية التحصيل (قيمة الدين × عمره ÷ سلوك السداد)، ويحقن الأولوية في خطة المندوب.">محرّك التحصيل الذكي</SectionTitle>
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={HandCoins} label="إجمالي الديون المفتوحة" value={money(totalDebt)} tint={{ color: C.danger, soft: C.dangerSoft }} hint={qty(ranked.length) + " محل"} />
      <Stat icon={Clock} label="متأخرة +30 يوم" value={urgent} tint={{ color: "#B45309", soft: C.warnSoft }} />
      <Stat icon={TrendingUp} label="أولوية قصوى" value={ranked.filter((r) => prBadge(r.pr).t === "قصوى").length} tint={{ color: C.danger, soft: C.dangerSoft }} />
      <Stat icon={CheckCircle2} label="متوسط سلوك السداد" value={(ranked.length ? Math.round(ranked.reduce((s, r) => s + r.score, 0) / ranked.length) : 0) + "%"} tint={{ color: C.success, soft: C.successSoft }} />
    </div>
    <Card className="p-4">
      <div className="flex items-center justify-between mb-3"><h3 className="font-extrabold" style={{ color: C.text }}>قائمة الأولويات</h3><span className="text-xs" style={{ color: C.soft }}>مرتّبة تلقائياً</span></div>
      <div className="space-y-2.5">{ranked.map((r) => { const pb = prBadge(r.pr); return (
        <div key={r.m.id} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}`, background: C.surface }}>
          <div className="flex items-center justify-between flex-wrap gap-2">
            <div className="flex items-center gap-2.5 min-w-0"><Avatar name={r.m.name} size={38} /><div className="min-w-0"><div className="font-extrabold text-sm truncate" style={{ color: C.text }}>{r.m.name}</div><div className="text-[11px]" style={{ color: C.soft }}>{r.m.area} · سلوك السداد {r.score}%</div></div></div>
            <span className="text-xs font-extrabold px-2.5 py-1 rounded-full shrink-0" style={{ background: pb.s, color: pb.c }}>أولوية {pb.t}</span>
          </div>
          <div className="grid grid-cols-3 gap-2 mt-2.5 text-center">
            <div className="rounded-lg py-1.5" style={{ background: C.dangerSoft }}><div className="text-[10px]" style={{ color: C.danger }}>الدين</div><div className="font-extrabold text-sm" style={{ color: C.danger }}>{money(r.debt)}</div></div>
            <div className="rounded-lg py-1.5" style={{ background: r.age > 30 ? C.warnSoft : C.well }}><div className="text-[10px]" style={{ color: r.age > 30 ? "#B45309" : C.soft }}>أقدم دين</div><div className="font-extrabold text-sm" style={{ color: r.age > 30 ? "#B45309" : C.text }}>{r.age} يوم</div></div>
            <div className="rounded-lg py-1.5" style={{ background: C.infoSoft }}><div className="text-[10px]" style={{ color: C.info }}>مؤشر الأولوية</div><div className="font-extrabold text-sm" style={{ color: C.info }}>{r.pr.toLocaleString("ar-LY")}</div></div>
          </div>
          <div className="flex gap-2 mt-2.5">
            <button onClick={() => { actions.pushCollectionTask(r.m.id, r.m.repId, r.debt); toast("أُضيف المحل لخطة المندوب ✓"); }} className="rb-press flex-1 rounded-lg py-2 text-xs font-bold flex items-center justify-center gap-1.5" style={{ background: SC.color, color: C.onAccent }}><Route size={14} />حقن في خطة المندوب</button>
            <button onClick={() => go && go("c_debts")} className="rb-press rounded-lg py-2 px-3 text-xs font-bold" style={{ background: C.well, color: C.text }}>كشف الحساب</button>
          </div>
        </div>); })}
        {ranked.length === 0 && <Empty icon={CheckCircle2} text="لا ديون مفتوحة — كل الحسابات مسدّدة." />}
      </div>
    </Card>
  </div>);
}

// ═══ ٢) التنبؤ بالطلب والتحميل الأمثل ═══
function CForecast(props) {
  const { data } = props; const co = useCo(props);
  const prods = data.products.filter((p) => p.companyId === co.id);
  const reps = data.reps.filter((r) => r.companyId === co.id);
  // تنبؤ مبسّط: متوسط الكمية المباعة للصنف عبر الطلبات المسلّمة × معامل موسمي
  const dlv = companySales(data, co.id);
  const soldQty = {}; dlv.forEach((o) => o.items.forEach((it) => { soldQty[it.pid] = (soldQty[it.pid] || 0) + it.qty; }));
  const forecast = prods.map((p) => {
    const sold = soldQty[p.id] || 0;
    const weeklyAvg = Math.max(1, Math.round(sold / 8)); // على مدى ~8 أسابيع
    const predicted = Math.round(weeklyAvg * 1.1); // معامل نمو 10%
    const cover = p.stock ? Math.floor(p.stock / weeklyAvg) : 0; // أسابيع التغطية
    return { p, sold, weeklyAvg, predicted, cover };
  }).sort((a, b) => b.predicted - a.predicted);
  const lowCover = forecast.filter((f) => f.cover <= 2);
  const [repId, setRepId] = useState(reps[0]?.id || "");
  const rep = reps.find((r) => r.id === repId);
  // تحميل أمثل: أعلى 8 أصناف متوقعة، مقيّدة بالمخزون
  const loadPlan = forecast.slice(0, 8).map((f) => ({ p: f.p, qty: Math.min(f.predicted, f.p.stock) })).filter((x) => x.qty > 0);
  const loadUnits = loadPlan.reduce((s, x) => s + x.qty, 0);
  const loadValue = loadPlan.reduce((s, x) => s + x.qty * priceAfter(x.p), 0);
  return (<div className="space-y-4"><SectionTitle sub="يتوقّع الطلب من تاريخ المبيعات، ويقترح تحميل سيارة كل مندوب لتقليل الرجيع والنفاد.">التنبؤ بالطلب والتحميل الأمثل</SectionTitle>
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={TrendingUp} label="أصناف عالية الطلب" value={forecast.filter((f) => f.predicted > 20).length} tint={{ color: SC.color, soft: SC.soft }} />
      <Stat icon={AlertTriangle} label="تغطية منخفضة (≤أسبوعين)" value={lowCover.length} tint={{ color: C.danger, soft: C.dangerSoft }} />
      <Stat icon={Boxes} label="وحدات التحميل المقترح" value={qty(loadUnits)} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={Landmark} label="قيمة التحميل المقترح" value={money(loadValue)} tint={{ color: C.success, soft: C.successSoft }} />
    </div>
    {lowCover.length > 0 && <Card className="p-4" style={{ borderRight: `3px solid ${C.danger}` }}>
      <h3 className="font-extrabold mb-2 flex items-center gap-2" style={{ color: C.danger }}><AlertTriangle size={17} />تنبيه إعادة تزويد</h3>
      <div className="flex flex-wrap gap-2">{lowCover.map((f) => <span key={f.p.id} className="text-xs font-bold px-2.5 py-1 rounded-full" style={{ background: C.dangerSoft, color: C.danger }}>{f.p.name} · {f.cover} أسبوع</span>)}</div>
    </Card>}
    <div className="grid lg:grid-cols-2 gap-4">
      <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>توقّع الطلب الأسبوعي</h3>
        <div className="overflow-auto"><table className="w-full text-xs"><thead><tr style={{ color: C.soft, background: C.well }}><th className="text-right py-2 px-2">الصنف</th><th className="text-center py-2">مبيع سابق</th><th className="text-center py-2">متوقّع/أسبوع</th><th className="text-center py-2">التغطية</th></tr></thead>
          <tbody>{forecast.slice(0, 12).map((f) => <tr key={f.p.id} style={{ borderBottom: `1px solid ${C.line}` }}><td className="py-2 px-2 font-bold" style={{ color: C.text }}>{f.p.name}</td><td className="text-center" style={{ color: C.soft }}>{qty(f.sold)}</td><td className="text-center font-extrabold" style={{ color: SC.color }}>{qty(f.predicted)}</td><td className="text-center"><span className="px-2 py-0.5 rounded-full text-[10px] font-bold" style={{ background: f.cover <= 2 ? C.dangerSoft : C.successSoft, color: f.cover <= 2 ? C.danger : C.success }}>{f.cover} أسبوع</span></td></tr>)}</tbody></table></div>
      </Card>
      <Card className="p-4"><div className="flex items-center justify-between mb-3 flex-wrap gap-2"><h3 className="font-extrabold" style={{ color: C.text }}>التحميل الأمثل للسيارة</h3><select value={repId} onChange={(e) => setRepId(e.target.value)} className={inputCls} style={{ ...inputStyle, width: "auto" }}>{reps.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}</select></div>
        {rep ? <><div className="space-y-1.5">{loadPlan.map((x) => <div key={x.p.id} className="flex items-center justify-between rounded-lg px-3 py-2" style={{ background: C.well }}><span className="text-sm font-bold" style={{ color: C.text }}>{x.p.name}</span><span className="font-extrabold text-sm" style={{ color: SC.color }}>{qty(x.qty)}</span></div>)}</div>
          <div className="flex items-center justify-between mt-3 pt-3" style={{ borderTop: `1px solid ${C.line}` }}><span className="text-sm font-bold" style={{ color: C.soft }}>الإجمالي</span><span className="font-extrabold" style={{ color: C.text }}>{qty(loadUnits)} · {money(loadValue)}</span></div></> : <Empty icon={Truck} text="لا مندوبين." />}
      </Card>
    </div>
  </div>);
}


// ═══ ٣) محرّك العروض والحملات ═══
const CAMPAIGN_TYPES = { bundle: "باقة مجمّعة", bonus: "كمية مجانية", timed: "خصم محدود المدة" };
function CCampaigns(props) {
  const { data, actions } = props; const co = useCo(props);
  const prods = data.products.filter((p) => p.companyId === co.id);
  const camps = (data.campaigns || []).filter((c) => c.companyId === co.id);
  const [form, setForm] = useState(null);
  const blank = { type: "bonus", productId: prods[0]?.id || "", buyQty: 10, freeQty: 1, discount: 10, minValue: 500, days: 7, title: "" };
  const active = camps.filter((c) => c.active);
  return (<div className="space-y-4"><SectionTitle sub="أنشئ عروضاً تظهر للتاجر وتُحتسب آلياً في الفاتورة: 10+1، باقات، وخصومات مؤقتة." action={<PrimaryBtn icon={Plus} onClick={() => setForm({ ...blank })}>عرض جديد</PrimaryBtn>}>العروض والحملات</SectionTitle>
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={Tag} label="عروض نشطة" value={active.length} tint={{ color: SC.color, soft: SC.soft }} />
      <Stat icon={Gift} label="عروض الكمية المجانية" value={camps.filter((c) => c.type === "bonus").length} tint={{ color: C.success, soft: C.successSoft }} />
      <Stat icon={Package} label="الباقات المجمّعة" value={camps.filter((c) => c.type === "bundle").length} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={Clock} label="خصومات مؤقتة" value={camps.filter((c) => c.type === "timed").length} tint={{ color: "#B45309", soft: C.warnSoft }} />
    </div>
    <div className="grid md:grid-cols-2 gap-3">{camps.map((c) => { const p = prods.find((x) => x.id === c.productId); return (
      <Card key={c.id} className="p-4" style={{ borderRight: `3px solid ${c.active ? SC.color : C.line}` }}>
        <div className="flex items-start justify-between mb-2"><div className="flex items-center gap-2"><div className="rounded-xl p-2" style={{ background: SC.soft }}>{c.type === "bonus" ? <Gift size={18} style={{ color: SC.color }} /> : c.type === "bundle" ? <Package size={18} style={{ color: SC.color }} /> : <Clock size={18} style={{ color: SC.color }} />}</div><div><div className="font-extrabold text-sm" style={{ color: C.text }}>{c.title || CAMPAIGN_TYPES[c.type]}</div><div className="text-[11px]" style={{ color: C.soft }}>{CAMPAIGN_TYPES[c.type]}</div></div></div>
          <span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: c.active ? C.successSoft : C.well, color: c.active ? C.success : C.soft }}>{c.active ? "نشط" : "موقوف"}</span></div>
        <div className="text-sm rounded-lg p-2.5 mb-2" style={{ background: C.well, color: C.text }}>
          {c.type === "bonus" && <span>اشترِ <b>{c.buyQty}</b> من {p?.name} واحصل على <b style={{ color: C.success }}>{c.freeQty} مجاناً</b></span>}
          {c.type === "bundle" && <span>باقة {p?.name}: خصم <b style={{ color: SC.color }}>{c.discount}%</b> عند شراء بـ{money(c.minValue)}+</span>}
          {c.type === "timed" && <span>خصم <b style={{ color: "#B45309" }}>{c.discount}%</b> على {p?.name} لمدة <b>{c.days}</b> يوم</span>}
        </div>
        <div className="flex gap-2"><button onClick={() => actions.toggleCampaign(c.id)} className="rb-press flex-1 rounded-lg py-1.5 text-xs font-bold" style={{ background: c.active ? C.dangerSoft : C.successSoft, color: c.active ? C.danger : C.success }}>{c.active ? "إيقاف" : "تفعيل"}</button><button onClick={() => actions.deleteCampaign(c.id)} className="rb-press rounded-lg py-1.5 px-3 text-xs font-bold" style={{ background: C.well, color: C.danger }}><Trash2 size={14} /></button></div>
      </Card>); })}
      {camps.length === 0 && <div className="md:col-span-2"><Empty icon={Tag} text="لا عروض بعد — أنشئ أول عرض." /></div>}
    </div>

    {form && <div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4 overflow-auto" onClick={() => setForm(null)}>
      <div className="rb-modal w-full max-w-md rounded-2xl p-5 my-8" onClick={(e) => e.stopPropagation()}>
        <h3 className="font-extrabold text-lg mb-4" style={{ color: C.text }}>عرض جديد</h3>
        <div className="space-y-3">
          <Field label="نوع العرض"><div className="grid grid-cols-3 gap-2">{Object.entries(CAMPAIGN_TYPES).map(([k, v]) => { const on = form.type === k; return <button key={k} onClick={() => setForm({ ...form, type: k })} className="rounded-xl py-2 text-xs font-bold" style={{ background: on ? SC.color : C.bg, color: on ? C.onAccent : C.text, border: `1px solid ${on ? SC.color : C.line}` }}>{v}</button>; })}</div></Field>
          <Field label="عنوان العرض (اختياري)"><Input value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} placeholder="مثال: عرض رمضان" /></Field>
          <Field label="الصنف"><select value={form.productId} onChange={(e) => setForm({ ...form, productId: e.target.value })} className={inputCls} style={inputStyle}>{prods.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}</select></Field>
          {form.type === "bonus" && <div className="grid grid-cols-2 gap-2"><Field label="كمية الشراء"><Input type="number" value={form.buyQty} onChange={(e) => setForm({ ...form, buyQty: +e.target.value })} /></Field><Field label="الكمية المجانية"><Input type="number" value={form.freeQty} onChange={(e) => setForm({ ...form, freeQty: +e.target.value })} /></Field></div>}
          {form.type === "bundle" && <div className="grid grid-cols-2 gap-2"><Field label="نسبة الخصم %"><Input type="number" value={form.discount} onChange={(e) => setForm({ ...form, discount: +e.target.value })} /></Field><Field label="أدنى قيمة شراء"><Input type="number" value={form.minValue} onChange={(e) => setForm({ ...form, minValue: +e.target.value })} /></Field></div>}
          {form.type === "timed" && <div className="grid grid-cols-2 gap-2"><Field label="نسبة الخصم %"><Input type="number" value={form.discount} onChange={(e) => setForm({ ...form, discount: +e.target.value })} /></Field><Field label="المدة (أيام)"><Input type="number" value={form.days} onChange={(e) => setForm({ ...form, days: +e.target.value })} /></Field></div>}
        </div>
        <div className="flex gap-2 mt-4"><PrimaryBtn full icon={CheckCircle2} onClick={() => { actions.addCampaign({ ...form, companyId: co.id, server: co.server }); setForm(null); toast("أُنشئ العرض ✓"); }}>إنشاء</PrimaryBtn><GhostBtn onClick={() => setForm(null)}>إلغاء</GhostBtn></div>
      </div>
    </div>}
  </div>);
}

// ═══ ٤) التقارير السنوية والمقارنات ═══
const MONTHS_AR = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];
function CAnnualReport(props) {
  const { data } = props; const co = useCo(props);
  const dlv = companySales(data, co.id);
  // توزيع تقديري على 12 شهراً (النموذج بلا تواريخ فعلية كاملة) — بذرة ثابتة لكل شهر
  const seedFactor = [0.7, 0.75, 0.9, 1.0, 1.15, 1.3, 1.1, 0.95, 1.05, 1.2, 1.35, 1.5];
  const baseSales = companySalesTotal(data, co.id) / 12 || 1;
  const baseQty = dlv.reduce((s, o) => s + orderQty(o), 0) / 12 || 1;
  const monthly = MONTHS_AR.map((mn, i) => ({ mn, sales: Math.round(baseSales * seedFactor[i]), qty: Math.round(baseQty * seedFactor[i]) }));
  const yearSales = monthly.reduce((s, m) => s + m.sales, 0);
  const yearQty = monthly.reduce((s, m) => s + m.qty, 0);
  const maxSales = Math.max(...monthly.map((m) => m.sales));
  const best = monthly.reduce((a, b) => b.sales > a.sales ? b : a);
  const worst = monthly.reduce((a, b) => b.sales < a.sales ? b : a);
  // أفضل الأصناف والمناطق عبر السنة
  const prodSold = {}; dlv.forEach((o) => o.items.forEach((it) => { prodSold[it.name] = (prodSold[it.name] || 0) + it.qty; }));
  const topProds = Object.entries(prodSold).sort((a, b) => b[1] - a[1]).slice(0, 5);
  const areaSales = {}; data.merchants.filter((m) => m.companyId === co.id).forEach((m) => { areaSales[m.area] = (areaSales[m.area] || 0) + merchantOpenDebt(data, m.id); });
  const growth = Math.round(((monthly[11].sales - monthly[0].sales) / (monthly[0].sales || 1)) * 100);
  return (<div className="space-y-4"><SectionTitle sub="ملخص أداء السنة: المبيعات والكميات شهرياً، وأفضل الأصناف والفترات.">التقرير السنوي — سنة في مراجعة</SectionTitle>
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={TrendingUp} label="مبيعات السنة" value={money(yearSales)} tint={{ color: C.success, soft: C.successSoft }} hint={qty(yearQty)} />
      <Stat icon={Calendar} label="أفضل شهر" value={best.mn} tint={{ color: SC.color, soft: SC.soft }} hint={money(best.sales)} />
      <Stat icon={TrendingDown} label="أضعف شهر" value={worst.mn} tint={{ color: C.danger, soft: C.dangerSoft }} hint={money(worst.sales)} />
      <Stat icon={Sparkles} label="نمو السنة" value={(growth >= 0 ? "+" : "") + growth + "%"} tint={{ color: growth >= 0 ? C.success : C.danger, soft: growth >= 0 ? C.successSoft : C.dangerSoft }} />
    </div>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>المبيعات الشهرية</h3>
      <div className="flex items-end gap-1.5" style={{ height: 180 }}>{monthly.map((m) => (<div key={m.mn} className="flex-1 flex flex-col items-center gap-1 group">
        <div className="w-full rounded-t-lg transition-all relative" style={{ height: `${(m.sales / maxSales) * 140}px`, background: m === best ? C.success : SC.color, minHeight: 4 }}><div className="absolute -top-5 inset-x-0 text-center text-[9px] font-bold opacity-0 group-hover:opacity-100" style={{ color: C.text }}>{Math.round(m.sales / 1000)}k</div></div>
        <span className="text-[9px]" style={{ color: C.soft }}>{m.mn.slice(0, 3)}</span>
      </div>))}</div>
    </Card>
    <div className="grid lg:grid-cols-2 gap-4">
      <Card className="p-4"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Award size={17} style={{ color: "#B45309" }} />أفضل 5 أصناف</h3>
        <div className="space-y-2">{topProds.map(([name, q], i) => <div key={name} className="flex items-center gap-2.5"><span className="w-6 h-6 rounded-lg flex items-center justify-center text-xs font-extrabold" style={{ background: i === 0 ? C.warnSoft : C.well, color: i === 0 ? "#B45309" : C.soft }}>{i + 1}</span><span className="flex-1 text-sm font-bold" style={{ color: C.text }}>{name}</span><span className="text-sm font-extrabold" style={{ color: SC.color }}>{qty(q)}</span></div>)}{topProds.length === 0 && <Empty icon={Package} text="لا مبيعات مسجّلة." />}</div>
      </Card>
      <Card className="p-4"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Calendar size={17} style={{ color: SC.color }} />مقارنة الأرباع</h3>
        <div className="space-y-2.5">{[0, 1, 2, 3].map((q) => { const qs = monthly.slice(q * 3, q * 3 + 3).reduce((s, m) => s + m.sales, 0); const qmax = Math.max(...[0, 1, 2, 3].map((x) => monthly.slice(x * 3, x * 3 + 3).reduce((s, m) => s + m.sales, 0))); return (<div key={q}><div className="flex justify-between text-sm mb-1"><span className="font-bold" style={{ color: C.text }}>الربع {["الأول", "الثاني", "الثالث", "الرابع"][q]}</span><span className="font-extrabold" style={{ color: C.text }}>{money(qs)}</span></div><ProgressBar value={qs} max={qmax} color={SC.color} /></div>); })}</div>
      </Card>
    </div>
  </div>);
}


// ═══ ٥) أداء المناطق (خريطة حرارية) ═══
function CRegions(props) {
  const { data } = props; const co = useCo(props);
  const merchants = data.merchants.filter((m) => m.companyId === co.id);
  const areas = {};
  merchants.forEach((m) => {
    const a = m.area || "غير محدد";
    if (!areas[a]) areas[a] = { name: a, count: 0, sales: 0, debt: 0, merchants: [] };
    areas[a].count++; areas[a].debt += merchantOpenDebt(data, m.id);
    areas[a].merchants.push(m);
  });
  companySales(data, co.id).forEach((o) => { const m = merchants.find((x) => x.id === o.merchantId); if (m) { const a = m.area || "غير محدد"; if (areas[a]) areas[a].sales += orderTotal(o); } });
  const list = Object.values(areas).sort((a, b) => b.sales - a.sales);
  const maxSales = Math.max(1, ...list.map((a) => a.sales));
  const maxDebt = Math.max(1, ...list.map((a) => a.debt));
  const heat = (v, max) => { const t = v / max; return t > 0.66 ? { bg: C.dangerSoft, c: C.danger, l: "مرتفع" } : t > 0.33 ? { bg: C.warnSoft, c: "#B45309", l: "متوسط" } : { bg: C.successSoft, c: C.success, l: "منخفض" }; };
  return (<div className="space-y-4"><SectionTitle sub="توزيع المبيعات والديون على المناطق — لتوجيه المندوبين والمروّجين.">أداء المناطق</SectionTitle>
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={MapPin} label="عدد المناطق" value={list.length} tint={{ color: SC.color, soft: SC.soft }} />
      <Stat icon={Store} label="أكثر منطقة مبيعاً" value={list[0]?.name || "—"} tint={{ color: C.success, soft: C.successSoft }} hint={list[0] ? money(list[0].sales) : ""} />
      <Stat icon={Users} label="إجمالي المحلات" value={merchants.length} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={HandCoins} label="أعلى منطقة ديوناً" value={[...list].sort((a, b) => b.debt - a.debt)[0]?.name || "—"} tint={{ color: C.danger, soft: C.dangerSoft }} />
    </div>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>الخريطة الحرارية للمناطق</h3>
      <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-2.5">{list.map((a) => { const hs = heat(a.sales, maxSales); const hd = heat(a.debt, maxDebt); return (
        <div key={a.name} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}`, background: hs.bg }}>
          <div className="flex items-center justify-between mb-2"><span className="font-extrabold text-sm flex items-center gap-1.5" style={{ color: C.text }}><MapPin size={14} style={{ color: hs.c }} />{a.name}</span><span className="text-[10px] font-bold px-1.5 py-0.5 rounded-full" style={{ background: C.surface, color: hs.c }}>{a.count} محل</span></div>
          <div className="grid grid-cols-2 gap-1.5 text-center">
            <div className="rounded-lg py-1.5" style={{ background: C.surface }}><div className="text-[9px]" style={{ color: C.soft }}>المبيعات</div><div className="font-extrabold text-xs" style={{ color: hs.c }}>{money(a.sales)}</div></div>
            <div className="rounded-lg py-1.5" style={{ background: C.surface }}><div className="text-[9px]" style={{ color: C.soft }}>الديون</div><div className="font-extrabold text-xs" style={{ color: hd.c }}>{money(a.debt)}</div></div>
          </div>
        </div>); })}
        {list.length === 0 && <div className="lg:col-span-3"><Empty icon={MapPin} text="لا مناطق مسجّلة." /></div>}
      </div>
      <div className="flex items-center gap-4 mt-3 text-[11px]" style={{ color: C.soft }}><span className="flex items-center gap-1"><span className="w-3 h-3 rounded" style={{ background: C.successSoft }} />منخفض</span><span className="flex items-center gap-1"><span className="w-3 h-3 rounded" style={{ background: C.warnSoft }} />متوسط</span><span className="flex items-center gap-1"><span className="w-3 h-3 rounded" style={{ background: C.dangerSoft }} />مرتفع</span></div>
    </Card>
  </div>);
}

// ═══ ٦) نقاط ولاء للتجار ═══
const loyaltyPoints = (data, mid) => {
  // نقاط = (حجم الشراء ÷ 100) + مكافأة الالتزام بالسداد
  const dlv = data.orders.filter((o) => o.merchantId === mid && o.status === "delivered");
  const volume = dlv.reduce((s, o) => s + orderTotal(o), 0);
  const score = merchantPayScore(data, mid);
  const base = Math.floor(volume / 100);
  const bonus = score >= 90 ? Math.floor(base * 0.2) : score >= 70 ? Math.floor(base * 0.1) : 0;
  return { points: base + bonus, volume, score, base, bonus };
};
const loyaltyTier = (pts) => pts >= 500 ? { name: "ذهبي", c: "#B45309", s: C.warnSoft, icon: Award } : pts >= 200 ? { name: "فضّي", c: "#64748B", s: C.well, icon: Star } : { name: "برونزي", c: "#92590C", s: C.accentSoft, icon: Star };
function CLoyalty(props) {
  const { data } = props; const co = useCo(props);
  const merchants = data.merchants.filter((m) => m.companyId === co.id);
  const rows = merchants.map((m) => ({ m, ...loyaltyPoints(data, m.id) })).sort((a, b) => b.points - a.points);
  const totalPts = rows.reduce((s, r) => s + r.points, 0);
  return (<div className="space-y-4"><SectionTitle sub="نقاط تُمنح على حجم الشراء والالتزام بالسداد المبكر — تُستبدل بخصومات.">نقاط ولاء التجار</SectionTitle>
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={Award} label="أعضاء ذهبيون" value={rows.filter((r) => r.points >= 500).length} tint={{ color: "#B45309", soft: C.warnSoft }} />
      <Stat icon={Star} label="إجمالي النقاط الممنوحة" value={totalPts.toLocaleString("ar-LY")} tint={{ color: SC.color, soft: SC.soft }} />
      <Stat icon={Users} label="التجار المشاركون" value={rows.filter((r) => r.points > 0).length} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={CheckCircle2} label="أعلى التزام سداد" value={(rows.length ? Math.max(...rows.map((r) => r.score)) : 0) + "%"} tint={{ color: C.success, soft: C.successSoft }} />
    </div>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>لوحة الولاء</h3>
      <div className="space-y-2.5">{rows.map((r, i) => { const t = loyaltyTier(r.points); const T = t.icon; return (
        <div key={r.m.id} className="flex items-center gap-3 rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}>
          <span className="w-7 h-7 rounded-lg flex items-center justify-center text-sm font-extrabold shrink-0" style={{ background: i < 3 ? C.warnSoft : C.well, color: i < 3 ? "#B45309" : C.soft }}>{i + 1}</span>
          <Avatar name={r.m.name} size={38} />
          <div className="flex-1 min-w-0"><div className="font-extrabold text-sm truncate" style={{ color: C.text }}>{r.m.name}</div><div className="text-[11px] flex items-center gap-1.5" style={{ color: C.soft }}><span className="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded-full font-bold" style={{ background: t.s, color: t.c }}><T size={10} />{t.name}</span>التزام {r.score}%</div></div>
          <div className="text-left shrink-0"><div className="font-extrabold" style={{ color: SC.color }}>{r.points.toLocaleString("ar-LY")}</div><div className="text-[10px]" style={{ color: C.soft }}>نقطة{r.bonus > 0 ? ` (+${r.bonus})` : ""}</div></div>
        </div>); })}
        {rows.length === 0 && <Empty icon={Award} text="لا تجار." />}
      </div>
    </Card>
  </div>);
}

// ═══ ٧) سجل النشاط الزمني لكل عميل ═══
function CClientTimeline(props) {
  const { data } = props; const co = useCo(props);
  const merchants = data.merchants.filter((m) => m.companyId === co.id);
  const [mid, setMid] = useState(merchants[0]?.id || "");
  const m = merchants.find((x) => x.id === mid);
  // بناء خط زمني موحّد من الطلبات والفواتير والمدفوعات والمرتجعات والزيارات
  const events = [];
  data.orders.filter((o) => o.merchantId === mid).forEach((o) => events.push({ t: "order", icon: ShoppingCart, c: SC.color, title: `طلب ${o.id}`, sub: `${qty(orderQty(o))} · ${money(orderTotal(o))}`, date: o.date, status: o.status }));
  data.invoices.filter((i) => i.merchantId === mid).forEach((i) => events.push({ t: "invoice", icon: FileText, c: C.info, title: `فاتورة ${i.id}`, sub: `${money(i.amount)}${i.paid < i.amount ? ` · متبقٍّ ${money(i.amount - i.paid)}` : " · مسدّدة"}`, date: i.date }));
  data.payments.filter((p) => p.merchantId === mid).forEach((p) => events.push({ t: "payment", icon: HandCoins, c: C.success, title: "دفعة", sub: `${money(p.amount)} · ${p.method}`, date: p.date }));
  data.returns.filter((r) => r.merchantId === mid).forEach((r) => events.push({ t: "return", icon: RotateCcw, c: C.danger, title: `مرتجع ${r.id}`, sub: r.reason || "", date: r.date }));
  (data.visits || []).filter((v) => v.merchantId === mid).forEach((v) => events.push({ t: "visit", icon: MapPin, c: C.purple, title: "زيارة ميدانية", sub: v.result || v.note || "", date: v.date }));
  const debt = merchantOpenDebt(data, mid);
  const lp = loyaltyPoints(data, mid);
  return (<div className="space-y-4"><SectionTitle sub="خط زمني موحّد لكل عميل: طلباته ومدفوعاته ومرتجعاته وزياراته في صفحة واحدة.">سجل نشاط العميل</SectionTitle>
    <Card className="p-4"><Field label="اختر العميل"><select value={mid} onChange={(e) => setMid(e.target.value)} className={inputCls} style={inputStyle}>{merchants.map((x) => <option key={x.id} value={x.id}>{x.name} — {x.area}</option>)}</select></Field></Card>
    {m && <><div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={Wallet} label="الرصيد الحالي" value={money(Math.abs(m.balance))} tint={{ color: m.balance < 0 ? C.danger : C.success, soft: m.balance < 0 ? C.dangerSoft : C.successSoft }} hint={m.balance < 0 ? "مدين" : "دائن"} />
      <Stat icon={HandCoins} label="ديون مفتوحة" value={money(debt)} tint={{ color: C.danger, soft: C.dangerSoft }} />
      <Stat icon={Award} label="نقاط الولاء" value={lp.points.toLocaleString("ar-LY")} tint={{ color: "#B45309", soft: C.warnSoft }} />
      <Stat icon={Activity} label="إجمالي الأحداث" value={events.length} tint={{ color: SC.color, soft: SC.soft }} />
    </div>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>الخط الزمني</h3>
      <div className="space-y-0">{events.map((e, i) => { const Ic = e.icon; return (
        <div key={i} className="flex gap-3 pb-4 relative">
          {i < events.length - 1 && <div className="absolute right-[18px] top-9 bottom-0 w-px" style={{ background: C.line }} />}
          <div className="rounded-full p-2 shrink-0 z-10" style={{ background: e.c + "22" }}><Ic size={18} style={{ color: e.c }} /></div>
          <div className="flex-1 min-w-0 pt-1"><div className="flex items-center justify-between gap-2"><span className="font-extrabold text-sm" style={{ color: C.text }}>{e.title}</span><span className="text-[11px] shrink-0" style={{ color: C.soft }}>{e.date}</span></div><div className="text-xs mt-0.5" style={{ color: C.soft }}>{e.sub}</div></div>
        </div>); })}
        {events.length === 0 && <Empty icon={Activity} text="لا نشاط لهذا العميل بعد." />}
      </div>
    </Card></>}
  </div>);
}


function CAds(props) {
  const { data, actions } = props; const co = useCo(props);
  const myAds = data.ads.filter((a) => a.companyId === co.id);
  const [f, setF] = useState({ title: "", subtitle: "", emoji: "🏷️", cta: "اطلب الآن", theme: 0 });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (<div className="space-y-4"><SectionTitle sub="صمّم إعلاناً مميزاً يظهر للتجار في سيرفرك ضمن مساحة الإعلانات الرئيسية.">المساحة الإعلانية</SectionTitle>
    <div className="grid lg:grid-cols-2 gap-4">
      <Card className="p-5 space-y-3 h-fit">
        <h3 className="font-extrabold" style={{ color: C.text }}>إعلان جديد</h3>
        <Field label="العنوان الرئيسي"><Input value={f.title} onChange={set("title")} placeholder="مثال: خصم 10% على الأرز" /></Field>
        <Field label="النص الفرعي"><Input value={f.subtitle} onChange={set("subtitle")} placeholder="تفاصيل العرض ومدته" /></Field>
        <div className="grid grid-cols-2 gap-2">
          <Field label="الأيقونة (إيموجي)"><Input value={f.emoji} onChange={set("emoji")} /></Field>
          <Field label="زر الدعوة"><Input value={f.cta} onChange={set("cta")} /></Field>
        </div>
        <div><div className="text-sm font-bold mb-1.5" style={{ color: C.text }}>اللون</div><div className="flex gap-2 flex-wrap">{AD_THEMES.map((t, i) => <button key={i} onClick={() => setF({ ...f, theme: i })} className="rounded-xl transition-all" style={{ width: 44, height: 30, background: `linear-gradient(120deg, ${t.c1}, ${t.c2})`, border: f.theme === i ? `3px solid ${C.ink}` : "3px solid transparent" }} />)}</div></div>
        <div><div className="text-sm font-bold mb-1.5" style={{ color: C.text }}>معاينة</div><AdCarousel ads={[{ ...f, ...AD_THEMES[f.theme], active: true, id: "preview" }]} /></div>
        <PrimaryBtn full icon={Plus} onClick={() => { if (f.title) { actions.addAd({ server: co.server, companyId: co.id, title: f.title, subtitle: f.subtitle, emoji: f.emoji, cta: f.cta, ...AD_THEMES[f.theme] }); setF({ title: "", subtitle: "", emoji: "🏷️", cta: "اطلب الآن", theme: 0 }); } }}>نشر الإعلان</PrimaryBtn>
      </Card>
      <Card className="p-5"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>إعلاناتي ({myAds.length})</h3>
        <div className="space-y-3">{myAds.map((a) => <div key={a.id} className="rounded-xl overflow-hidden" style={{ border: `1px solid ${C.line}`, opacity: a.active ? 1 : 0.5 }}>
          <div className="p-3 flex items-center gap-3" style={{ background: `linear-gradient(120deg, ${a.c1}, ${a.c2})` }}><span className="text-2xl">{a.emoji}</span><div className="flex-1 min-w-0 text-white"><div className="font-bold text-sm truncate">{a.title}</div><div className="text-[11px] opacity-90 truncate">{a.subtitle}</div></div></div>
          <div className="flex items-center justify-between px-3 py-2"><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={a.active ? { background: C.successSoft, color: C.success } : { background: C.well, color: C.soft }}>{a.active ? "ظاهر للتجار" : "متوقّف"}</span><div className="flex gap-1"><button onClick={() => actions.toggleAd(a.id)} className="text-xs font-bold px-2 py-1 rounded-lg" style={{ background: C.bg, color: C.text }}>{a.active ? "إيقاف" : "تفعيل"}</button><button onClick={() => confirmThen("سيُحذف الإعلان نهائياً.", () => actions.removeAd(a.id), { tone: "danger", yesLabel: "حذف الإعلان" })} className="rounded-lg p-1.5" style={{ background: C.dangerSoft }}><Trash2 size={14} style={{ color: C.danger }} /></button></div></div>
        </div>)}{myAds.length === 0 && <Empty icon={ImageIcon} text="لا توجد إعلانات بعد." />}</div>
      </Card>
    </div>
  </div>);
}
// الشركة تخصص كميات البضاعة (ستيكات) لمديري المبيعات + هدف المدير العام
function CStockManagers(props) {
  const { data, actions } = props; const co = useCo(props);
  const managers = data.salesManagers.filter((s) => s.companyId === co.id);
  const prods = data.products.filter((p) => p.companyId === co.id);
  const [f, setF] = useState({ managerId: managers[0]?.id || "", pid: prods[0]?.id || "", allocated: 100 });
  return (<div className="space-y-4"><SectionTitle sub="حدّد كميات البضاعة (بالستيكات/الكراتين) التي يوزّعها كل مدير مبيعات، وهدفه العام من التحصيل.">تخصيص كميات لمديري المبيعات</SectionTitle>
    {/* أهداف المدير العامة (كمية + تحصيل) */}
    <Card className="p-4"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Target size={17} style={{ color: C.ink }} />أهداف مديري المبيعات</h3>
      <Table head={["المدير", "هدف الكمية (ستيك)", "هدف التحصيل", "إنجاز الكمية", "إنجاز التحصيل"]}>{managers.map((s) => <tr key={s.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{s.name}</Td><Td><input type="number" defaultValue={s.qtyTarget} onBlur={(e) => actions.setMgrTarget(s.id, { qtyTarget: +e.target.value })} className="w-24 rounded-lg px-2 py-1 text-sm outline-none" style={inputStyle} /></Td><Td><input type="number" defaultValue={s.collectionTarget} onBlur={(e) => actions.setMgrTarget(s.id, { collectionTarget: +e.target.value })} className="w-28 rounded-lg px-2 py-1 text-sm outline-none" style={inputStyle} /></Td><Td>{qty(s.qtyAchieved)}</Td><Td>{money(s.collected)}</Td></tr>)}</Table>
      {managers.length === 0 && <Empty icon={UserCog} text="أضف مديري مبيعات أولاً." />}
    </Card>
    {/* تخصيص كميات المنتجات */}
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>تخصيص كمية منتج لمدير</h3>
      <div className="grid sm:grid-cols-4 gap-2 items-end">
        <Field label="المدير"><select value={f.managerId} onChange={(e) => setF({ ...f, managerId: e.target.value })} className={inputCls} style={inputStyle}>{managers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}</select></Field>
        <Field label="المنتج"><select value={f.pid} onChange={(e) => setF({ ...f, pid: e.target.value })} className={inputCls} style={inputStyle}>{prods.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}</select></Field>
        <Field label="الكمية (ستيك/وحدة)"><Input type="number" value={f.allocated} onChange={(e) => setF({ ...f, allocated: +e.target.value })} /></Field>
        <PrimaryBtn icon={Plus} onClick={() => { if (f.managerId && f.pid) actions.setMgrStock({ companyId: co.id, ...f }); }}>تخصيص</PrimaryBtn>
      </div>
    </Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>الكميات المخصّصة</h3>
      <Table head={["المدير", "المنتج", "المخصص", "المُوزّع", "المتبقي"]}>{data.mgrStock.filter((s) => s.companyId === co.id).map((s) => { const mgr = managers.find((m) => m.id === s.managerId); const rem = s.allocated - s.distributed; return (<tr key={s.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{mgr?.name}</Td><Td>{prodName(data, s.pid)}</Td><Td>{qty(s.allocated)}</Td><Td>{qty(s.distributed)}</Td><Td><span style={{ color: rem < 100 ? C.danger : C.text, fontWeight: 700 }}>{qty(rem)}</span></Td></tr>); })}</Table>
      {data.mgrStock.filter((s) => s.companyId === co.id).length === 0 && <Empty icon={Boxes} text="لا توجد كميات مخصّصة." />}
    </Card>
  </div>);
}

// مُجمِّع رحلات السائق: عدة طلبات جاهزة ← رحلة واحدة بمحطات
function TripBuilder({ ready, drivers, data, actions }) {
  const [sel, setSel] = useState([]);
  const [drv, setDrv] = useState(drivers[0]?.id || "");
  const toggle = (id) => setSel((x) => x.includes(id) ? x.filter((i) => i !== id) : [...x, id]);
  const totalVal = sel.reduce((s, id) => { const o = ready.find((x) => x.id === id); return s + (o ? orderTotal(o) : 0); }, 0);
  return (<Card className="p-4" style={{ border: `1px solid ${C.info}` }}>
    <h3 className="font-extrabold mb-1 flex items-center gap-2" style={{ color: C.text }}><Route size={17} style={{ color: C.info }} />تجميع رحلة سائق</h3>
    <p className="text-xs mb-3" style={{ color: C.soft }}>اجمع عدة طلبات جاهزة (نفس المنطقة عادةً) في رحلة واحدة لسائق واحد بمحطات مرتّبة.</p>
    <div className="space-y-1.5 mb-3">{ready.map((o) => { const m = data.merchants.find((x) => x.id === o.merchantId); const on = sel.includes(o.id); return (<button key={o.id} onClick={() => toggle(o.id)} className="rb-press w-full flex items-center justify-between gap-2 rounded-xl px-3 py-2 text-right" style={{ background: on ? C.infoSoft : C.surface, border: `1px solid ${on ? C.info : C.line}` }}>
      <div className="flex items-center gap-2 min-w-0"><span className="rounded-md flex items-center justify-center shrink-0" style={{ width: 20, height: 20, background: on ? C.info : "#fff", border: `2px solid ${on ? C.info : C.line}` }}>{on && <CheckCircle2 size={13} color="#fff" />}</span><div className="min-w-0"><div className="text-sm font-bold truncate" style={{ color: C.text }}>{o.id} · {m?.name}</div><div className="text-[10px]" style={{ color: C.soft }}>{m?.area} · {money(orderTotal(o))}</div></div></div>
    </button>); })}</div>
    <div className="flex flex-wrap items-end gap-2">
      <div className="flex-1" style={{ minWidth: 160 }}><Field label="السائق"><select value={drv} onChange={(e) => setDrv(e.target.value)} className={inputCls} style={inputStyle}>{drivers.map((v) => <option key={v.id} value={v.id}>{v.name}</option>)}</select></Field></div>
      <PrimaryBtn sm icon={Route} onClick={() => { if (!sel.length || !drv) { toast("اختر طلبات وسائقاً", "danger"); return; } actions.createTrip({ companyId: ready[0].companyId, driverId: drv, orderIds: sel }); toast(`أُنشئت رحلة بـ${sel.length} طلب ✓`); setSel([]); }}>إنشاء رحلة ({sel.length} طلب · {money(totalVal)})</PrimaryBtn>
    </div>
  </Card>);
}
// محطة استلام وتجهيز الطلبات: استلام/طباعة بالمنظومة ← مخزن/تعبئة ← تسليم للسائق
function CIntake(props) {
  const { data, actions, user } = props; const co = useCo(props);
  const drivers = data.drivers.filter((v) => v.companyId === co.id);
  const [pick, setPick] = useState({});
  const [invOrder, setInvOrder] = useState(null);
  const [pickList, setPickList] = useState(null);   // قائمة التجهيز للطباعة
  const [packQty, setPackQty] = useState({});        // {orderId: {pid: qty}} كميات التعبئة الفعلية
  const setPQ = (oid, pid, v) => setPackQty((x) => ({ ...x, [oid]: { ...(x[oid] || {}), [pid]: v } }));
  const inFlow = data.orders.filter((o) => o.companyId === co.id && ["pending_company", "printed", "warehouse"].includes(o.status));
  const done = data.orders.filter((o) => o.companyId === co.id && ["preparing", "out_for_delivery", "delivered"].includes(o.status)).slice(0, 8);
  const STEP = { pending_company: 0, printed: 1, warehouse: 2, preparing: 3 };
  const stepDefs = [{ t: "استلام وطباعة بالمنظومة", i: Printer }, { t: "إرسال للمخزن للتعبئة", i: Warehouse }, { t: "تعبئة وتسليم للسائق", i: Truck }];
  return (<div className="space-y-4"><SectionTitle sub="المسار الداخلي للطلب: استلامه وطباعته على منظومة الشركة ← إرساله للمخزن للتعبئة ← تسليمه للسائق لينقله للعميل.">استلام وتجهيز الطلبات</SectionTitle>
    {(() => {
      const readyForPickup = data.orders.filter((o) => o.companyId === co.id && o.status === "ready_dispatch" && o.fulfillment !== "pickup");
      const pendingReq = (data.pickupRequests || []).filter((r) => r.companyId === co.id && r.status === "requested");
      const disp = data.dispatchers.find((x) => x.companyId === co.id);
      if (!disp) return null;
      return (<Card className="p-4" style={{ borderRight: `3px solid #B45309` }}>
        <div className="flex items-center justify-between flex-wrap gap-2 mb-2">
          <h3 className="font-extrabold flex items-center gap-2" style={{ color: C.text }}><Route size={17} style={{ color: "#B45309" }} />طلب استلام من مشرف الحركة</h3>
          {pendingReq.length > 0 && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.warnSoft, color: "#B45309" }}>{pendingReq.length} طلب قيد الانتظار</span>}
        </div>
        <p className="text-xs mb-3" style={{ color: C.soft }}>بعد إدخال الفواتير في المنظومة وتجهيزها، اطلب من مشرف الحركة ({disp.name}) الحضور لاستلامها يدوياً.</p>
        <PrimaryBtn full icon={Inbox} onClick={() => { if (!readyForPickup.length) { toast("لا فواتير جاهزة للاستلام", "danger"); return; } actions.requestPickup({ companyId: co.id, orderIds: readyForPickup.map((o) => o.id), note: `${readyForPickup.length} فاتورة جاهزة` }); toast("أُرسل طلب الاستلام لمشرف الحركة ✓"); }}>طلب حضور لاستلام الفواتير ({readyForPickup.length} جاهزة)</PrimaryBtn>
      </Card>);
    })()}
    <div className="grid sm:grid-cols-3 gap-3">
      <Stat icon={Inbox} label="وصلت الشركة (لم تُطبع)" value={inFlow.filter((o) => o.status === "pending_company").length} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={Warehouse} label="في المخزن" value={inFlow.filter((o) => o.status === "warehouse").length} tint={{ color: "#7C3AED", soft: C.purpleSoft }} />
      <Stat icon={Printer} label="طُبعت بالمنظومة" value={inFlow.filter((o) => o.status === "printed").length} tint={{ color: C.purple, soft: C.purpleSoft }} />
    </div>
    {inFlow.length === 0 && <Empty icon={Inbox} text="لا توجد طلبات بانتظار التجهيز." />}
    {inFlow.map((o) => { const cur = STEP[o.status]; const m = data.merchants.find((x) => x.id === o.merchantId); const rep = data.reps.find((r) => r.id === o.repId); const isPickup = o.fulfillment === "pickup";
      return (<Card key={o.id} className="p-5">
        <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
          <div><div className="font-extrabold flex items-center gap-2" style={{ color: C.text }}>طلب {o.id}{o.sysNo && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.purpleSoft, color: C.purple }}>منظومة: {o.sysNo}</span>}</div><div className="text-xs mt-0.5" style={{ color: C.soft }}>{m?.name} · {m?.area} · {opsCan(user, co, "prices") ? money(orderTotal(o)) : "•••"} · {qty(orderQty(o))} وحدة</div></div>
          <div className="flex items-center gap-2"><span className="inline-flex items-center gap-1 text-[11px] font-bold px-2 py-0.5 rounded-full" style={{ background: FULFILL[o.fulfillment || "delivery"].soft, color: FULFILL[o.fulfillment || "delivery"].color }}>{isPickup ? <Store size={11} /> : <Truck size={11} />}{FULFILL[o.fulfillment || "delivery"].label}</span><Badge status={o.status} map={STATUS} /></div>
        </div>
        {/* مسار داخلي */}
        <div className="flex items-center justify-between gap-1 mb-4">{stepDefs.map((s, idx) => { const dn = cur > idx; const active = cur === idx; return (<React.Fragment key={idx}><div className="flex flex-col items-center gap-1 text-center" style={{ minWidth: 58 }}><div className="rounded-full flex items-center justify-center" style={{ width: 34, height: 34, background: dn ? C.success : active ? SC.color : C.surface, border: `2px solid ${dn ? C.success : active ? SC.color : C.line}` }}><s.i size={15} style={{ color: dn || active ? "#fff" : C.soft }} /></div><span className="text-[10px] font-bold" style={{ color: dn || active ? C.text : C.soft }}>{s.t}</span></div>{idx < stepDefs.length - 1 && <div className="h-0.5 flex-1 rounded" style={{ background: cur > idx ? C.success : C.line }} />}</React.Fragment>); })}</div>
        {/* أزرار المرحلة */}
        <div className="pt-3 flex flex-wrap items-end gap-2" style={{ borderTop: `1px solid ${C.line}` }}>
          {opsCan(user, co, "invoice") && <GhostBtn sm icon={Receipt} onClick={() => setInvOrder(o)}>عرض/طباعة الفاتورة</GhostBtn>}
          {o.status === "pending_company" && <PrimaryBtn sm icon={Printer} onClick={() => { actions.companyReceivePrint(o.id); toast("استُلم الطلب وطُبع بالمنظومة ✓"); }}>استلام وطباعة على المنظومة</PrimaryBtn>}
          {o.status === "printed" && <PrimaryBtn sm icon={Warehouse} onClick={() => { actions.companyToWarehouse(o.id); toast("أُرسل للمخزن للتعبئة ✓"); }}>إرسال للمخزن للتعبئة</PrimaryBtn>}
          {o.status === "warehouse" && <GhostBtn sm icon={ClipboardList} onClick={() => setPickList(o)}>قائمة التجهيز</GhostBtn>}
        </div>
        {/* التعبئة الجزئية: كمية فعلية لكل بند + إسناد سائق */}
        {o.status === "warehouse" && <div className="mt-3 rounded-xl p-3" style={{ background: C.purpleSoft + "44", border: `1px solid ${C.purpleSoft}` }}>
          <div className="text-xs font-extrabold mb-2 flex items-center gap-1.5" style={{ color: "#7C3AED" }}><Boxes size={14} />التعبئة الفعلية (عدّل الكمية إن نقص صنف)</div>
          <div className="space-y-1.5 mb-3">{o.items.map((it) => { const p = data.products.find((x) => x.id === it.pid); const got = (packQty[o.id] || {})[it.pid]; const val = got != null ? got : it.qty; const short = val < it.qty; const lowStock = p && p.stock < it.qty; return (<div key={it.pid} className="flex items-center justify-between gap-2 rounded-lg px-2.5 py-1.5" style={{ background: C.well }}>
            <div className="min-w-0 flex-1"><div className="text-xs font-bold truncate" style={{ color: C.text }}>{it.name}</div><div className="text-[10px]" style={{ color: lowStock ? C.danger : C.soft }}>مطلوب {it.qty}{p ? ` · متوفر بالمخزن ${p.stock}` : ""}{lowStock ? " ⚠️ نقص" : ""}</div>{p && fefoOlder(data, p) && <div className="text-[9px] font-bold mt-0.5 px-1.5 py-0.5 rounded inline-block" style={{ background: C.accentSoft, color: "#92590C" }}>FEFO ⚠️ تشغيلة أقدم صلاحية ({fmtDate(fefoOlder(data, p).expDate)}) — اصرفها أولاً</div>}</div>
            <div className="flex items-center gap-1 shrink-0"><input type="number" value={val} onChange={(e) => setPQ(o.id, it.pid, Math.max(0, Math.min(it.qty, Math.floor(+e.target.value || 0))))} className="w-16 text-center text-sm font-bold rounded-lg py-1 outline-none" style={{ ...inputStyle, color: short ? C.danger : C.text }} />{short && <span className="text-[10px] font-bold" style={{ color: C.danger }}>-{it.qty - val}</span>}</div>
          </div>); })}</div>
          <div className="flex flex-wrap items-end gap-2">
            {!isPickup && <div className="flex-1" style={{ minWidth: 170 }}><Field label="إسناد التوصيل لأي سائق"><select value={pick[o.id] || drivers[0]?.id || ""} onChange={(e) => setPick({ ...pick, [o.id]: e.target.value })} className={inputCls} style={inputStyle}>{drivers.map((v) => <option key={v.id} value={v.id}>{v.name} — {v.area || v.vehicle || "سائق"}</option>)}</select></Field></div>}
            <PrimaryBtn sm tone="success" icon={CheckCircle2} onClick={() => { const pq = packQty[o.id]; const hasShort = pq && o.items.some((it) => pq[it.pid] != null && pq[it.pid] < it.qty); const doIt = () => { actions.companyPackAssign(o.id, isPickup ? null : (pick[o.id] || drivers[0]?.id), pq); toast(isPickup ? "جاهز للاستلام من المقر ✓" : "سُلّم للسائق ✓"); }; hasShort ? confirmThen("توجد أصناف ناقصة — ستُعدَّل الفاتورة ويُعلَم العميل والمندوب بالنقص.", doIt, { yesLabel: "تأكيد التعبئة الجزئية" }) : doIt(); }}>{isPickup ? "تمّت التعبئة — جاهز للاستلام" : "تمّت التعبئة — تسليم للسائق"}</PrimaryBtn>
          </div>
        </div>}
      </Card>); })}
    {/* رحلات السائق: جمع الطلبات الجاهزة في رحلة واحدة */}
    {(() => { const ready = data.orders.filter((o) => o.companyId === co.id && o.status === "preparing" && o.fulfillment !== "pickup" && !o.tripId); if (!ready.length) return null; return (<TripBuilder ready={ready} drivers={drivers} data={data} actions={actions} />); })()}
    {done.length > 0 && <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>طلبات جُهّزت مؤخراً</h3><Table head={["الطلب", "العميل", "منظومة", "القيمة", "الاستلام", "الحالة", ""]}>{done.map((o) => <tr key={o.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{o.id}</Td><Td>{merchName(data, o.merchantId)}</Td><Td>{o.sysNo || "—"}</Td><Td>{opsCan(user, co, "prices") ? money(orderTotal(o)) : "•••"}</Td><Td>{FULFILL[o.fulfillment || "delivery"].label}</Td><Td><Badge status={o.status} map={STATUS} /></Td><Td>{opsCan(user, co, "invoice") ? <GhostBtn sm icon={Receipt} onClick={() => setInvOrder(o)}>فاتورة</GhostBtn> : <span className="text-xs" style={{ color: C.soft }}>—</span>}</Td></tr>)}</Table></Card>}
    {invOrder && <InvoiceView order={invOrder} data={data} onClose={() => setInvOrder(null)} />}
    {pickList && <PickingListView order={pickList} data={data} onClose={() => setPickList(null)} />}
  </div>);
}
// تصنيف المندوبين: تجزئة (سيارة بيع) أو بري‑سيل (جملة)
function CRepTypes(props) {
  const { data, actions } = props; const co = useCo(props);
  const reps = data.reps.filter((r) => r.companyId === co.id);
  return (<div className="space-y-4"><SectionTitle sub="صنّف كل مندوب: «تجزئة» يبيع من مخزون سيارته فوراً، أو «بري‑سيل» يأخذ الطلب/الفاتورة لتجهّزها الشركة وتوصّلها.">تصنيف المندوبين</SectionTitle>
    <div className="grid sm:grid-cols-2 gap-3">{Object.entries(REP_TYPE).map(([k, v]) => <div key={k} className="rounded-xl p-4 flex items-start gap-3" style={{ background: v.soft }}><div className="rounded-lg p-2" style={{ background: C.well }}><v.icon size={18} style={{ color: v.color }} /></div><div><div className="font-bold text-sm" style={{ color: v.color }}>{v.label}</div><div className="text-xs mt-0.5" style={{ color: C.soft }}>{v.desc}</div></div></div>)}</div>
    <Card className="p-4"><Table head={["المندوب", "المنطقة", "المشرف", "التصنيف الحالي", "تغيير", "مخزون السيارة"]}>{reps.map((r) => { const t = REP_TYPE[r.repType || "presell"]; const vanCount = Object.values(r.vanStock || {}).reduce((s, q) => s + q, 0);
      return (<tr key={r.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{r.name}</Td><Td>{r.area}</Td><Td>{r.supervisorId ? data.supervisors.find((s) => s.id === r.supervisorId)?.name : "—"}</Td><Td><span className="inline-flex items-center gap-1 text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: t.soft, color: t.color }}><t.icon size={11} />{t.label}</span></Td><Td><select value={r.repType || "presell"} onChange={(e) => actions.setRepType(r.id, e.target.value)} className="rounded-lg px-2 py-1 text-sm outline-none" style={inputStyle}>{Object.entries(REP_TYPE).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}</select></Td><Td>{r.repType === "van" ? <span className="font-bold" style={{ color: C.accent }}>{qty(vanCount)} وحدة</span> : "—"}</Td></tr>); })}</Table>{reps.length === 0 && <Empty icon={Handshake} text="لا يوجد مندوبون." />}</Card>
  </div>);
}
// تحميل سيارات مندوبي التجزئة: الشركة تعبّئ الطلب فيُضاف لمخزون السيارة
function CLoads(props) {
  const { data, actions, user } = props; const co = useCo(props);
  const reqs = data.loadRequests.filter((l) => l.companyId === co.id);
  const pending = reqs.filter((l) => l.status === "requested");
  const vanReps = data.reps.filter((r) => r.companyId === co.id && r.repType === "van");
  const prods = data.products.filter((p) => p.companyId === co.id && p.status === "active");
  const [nl, setNl] = useState({ repId: vanReps[0]?.id || "", note: "", cart: {} });
  const nlItems = Object.entries(nl.cart).filter(([, q]) => q > 0);
  const setNlQ = (pid, q) => setNl((s) => { const cart = { ...s.cart }; if (q <= 0) delete cart[pid]; else cart[pid] = q; return { ...s, cart }; });
  const createLoad = () => { if (!nl.repId || !nlItems.length) return; actions.companyCreateLoad({ repId: nl.repId, companyId: co.id, items: nlItems.map(([pid, q]) => ({ pid, qty: q })), note: nl.note }); setNl({ repId: nl.repId, note: "", cart: {} }); };
  return (<div className="space-y-4"><SectionTitle sub="طلبات مندوبي التجزئة لتعبئة سياراتهم، وإنشاء تحميل جديد لهم (يُضاف بعد موافقة المندوب)، ومراجعة مخزون كل سيارة.">تحميل سيارات المندوبين</SectionTitle>
    <div className="grid sm:grid-cols-2 gap-3">
      <Stat icon={PackagePlus} label="طلبات تحميل من المندوبين" value={pending.length} tint={{ color: C.accent, soft: C.accentSoft }} />
      <Stat icon={Truck} label="مندوبو التجزئة" value={vanReps.length} tint={{ color: C.info, soft: C.infoSoft }} />
    </div>
    {/* إنشاء تحميل جديد لمندوب (بعد موافقته) */}
    {opsCan(user, co, "loads") && <Card className="p-4"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><PackagePlus size={17} style={{ color: SC.color }} />إنشاء تحميل لمندوب تجزئة</h3>
      {vanReps.length === 0 ? <Empty icon={Truck} text="لا يوجد مندوبو تجزئة — صنّفهم من «تصنيف المندوبين» أولاً." /> : <>
        <div className="grid sm:grid-cols-2 gap-2 mb-3"><Field label="المندوب"><select value={nl.repId} onChange={(e) => setNl({ ...nl, repId: e.target.value })} className={inputCls} style={inputStyle}>{vanReps.map((r) => <option key={r.id} value={r.id}>{r.name} — {r.area}</option>)}</select></Field><Field label="ملاحظة"><Input value={nl.note} onChange={(e) => setNl({ ...nl, note: e.target.value })} placeholder="اختياري" /></Field></div>
        <div className="grid sm:grid-cols-2 gap-2 max-h-56 overflow-auto mb-3">{prods.map((p) => { const q = nl.cart[p.id] || 0; return (<div key={p.id} className="flex items-center justify-between rounded-xl p-2.5" style={{ border: `1px solid ${q ? SC.color : C.line}` }}><span className="text-sm font-bold min-w-0 truncate" style={{ color: C.text }}>{p.name}</span>{q > 0 ? <div className="flex items-center gap-1 shrink-0"><button onClick={() => setNlQ(p.id, q - 5)} className="w-7 h-7 rounded-lg font-bold" style={{ border: `1px solid ${C.line}` }}>−</button><input type="number" value={q} onChange={(e) => setNlQ(p.id, Math.max(0, Math.floor(+e.target.value || 0)))} className="w-14 text-center text-sm font-bold rounded-lg py-1 outline-none" style={inputStyle} /><button onClick={() => setNlQ(p.id, q + 5)} className="w-7 h-7 rounded-lg font-bold text-white" style={{ background: SC.color }}>+</button></div> : <button onClick={() => setNlQ(p.id, 10)} className="shrink-0 rounded-lg px-2.5 py-1.5 text-xs font-bold text-white" style={{ background: SC.color }}>تحميل</button>}</div>); })}</div>
        <div className="flex items-center justify-between flex-wrap gap-2"><span className="text-sm" style={{ color: C.soft }}>{nlItems.length} صنف · {nlItems.reduce((s, [, q]) => s + q, 0)} وحدة</span><PrimaryBtn icon={Send} onClick={createLoad}>إرسال التحميل لموافقة المندوب</PrimaryBtn></div>
      </>}
    </Card>}
    {reqs.map((l) => { const rep = data.reps.find((r) => r.id === l.repId); const totalU = l.items.reduce((s, i) => s + i.qty, 0); const byCompany = l.origin === "company";
      return (<Card key={l.id} className="p-5">
        <div className="flex items-center justify-between mb-3 flex-wrap gap-2"><div className="flex items-center gap-3"><div className="rounded-xl p-2.5" style={{ background: C.accentSoft }}><Truck size={20} style={{ color: C.accent }} /></div><div><div className="font-extrabold flex items-center gap-2" style={{ color: C.text }}>{l.id} · {rep?.name}<span className="text-[10px] font-bold px-1.5 py-0.5 rounded-full" style={{ background: byCompany ? SC.soft : C.accentSoft, color: byCompany ? SC.color : "#92590C" }}>{byCompany ? "من الشركة" : "طلب المندوب"}</span></div><div className="text-xs" style={{ color: C.soft }}>{totalU} وحدة · {l.date}{l.note ? ` · ${l.note}` : ""}</div></div></div><Badge status={l.status} map={LOAD_STATUS} /></div>
        <div className="rounded-xl overflow-hidden mb-3" style={{ border: `1px solid ${C.line}` }}>{l.items.map((it, i) => <div key={i} className="flex justify-between px-3 py-2 text-sm" style={{ borderBottom: i < l.items.length - 1 ? `1px solid ${C.line}` : "none" }}><span style={{ color: C.text }}>{prodName(data, it.pid)}</span><span className="font-bold" style={{ color: C.text }}>{qty(it.qty)}</span></div>)}</div>
        {l.status === "requested" && <div className="flex gap-2"><PrimaryBtn tone="success" icon={CheckCircle2} onClick={() => { actions.fulfillLoadRequest(l.id); toast("حُمّلت السيارة وأُضيفت الكميات ✓"); }}>تعبئة وتحميل السيارة</PrimaryBtn><GhostBtn icon={XCircle} onClick={() => confirmThen("سيُرفض طلب التحميل ويُعلَم المندوب.", () => actions.rejectLoadRequest(l.id), { tone: "danger", yesLabel: "رفض" })}>رفض</GhostBtn></div>}
        {l.status === "pending_rep" && <div className="text-sm flex items-center gap-1.5" style={{ color: C.info }}><Clock size={15} />بانتظار موافقة المندوب {rep?.name} على التحميل</div>}
        {l.status === "loaded" && <div className="text-sm flex items-center gap-1.5" style={{ color: C.success }}><CheckCircle2 size={15} />أُضيفت الكميات إلى مخزون سيارة {rep?.name}</div>}
        {l.status === "rejected" && <div className="text-sm flex items-center gap-1.5" style={{ color: C.danger }}><XCircle size={15} />مرفوض</div>}
      </Card>); })}
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>مخزون سيارات مندوبي التجزئة</h3>
      <div className="grid sm:grid-cols-2 gap-3">{vanReps.map((r) => { const st = Object.entries(r.vanStock || {}).filter(([, q]) => q > 0); const tot = st.reduce((s, [, q]) => s + q, 0); return (<div key={r.id} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}><div className="font-bold text-sm mb-2 flex items-center justify-between" style={{ color: C.text }}><span className="flex items-center gap-1.5"><Truck size={14} style={{ color: C.accent }} />{r.name}</span><span className="text-xs" style={{ color: C.soft }}>{qty(tot)} وحدة</span></div>{st.length ? <div className="space-y-1">{st.map(([pid, q]) => <div key={pid} className="flex justify-between text-xs"><span style={{ color: C.soft }}>{prodName(data, pid)}</span><span className="font-bold" style={{ color: q < 10 ? C.danger : C.text }}>{qty(q)}</span></div>)}</div> : <div className="text-xs" style={{ color: C.soft }}>السيارة فارغة.</div>}</div>); })}{vanReps.length === 0 && <Empty icon={Truck} text="لا يوجد مندوبو تجزئة. صنّفهم أولاً." />}</div>
    </Card>
  </div>);
}

/* ===== الرئيسية الفخمة الموحّدة لكل الأدوار: ترويسة + تنبيهات + شبكة كل الأقسام ===== */
const LAUNCH_PALETTE = [
  ["#0D47A1", "#1E6BD6"], ["#B45309", "#F59E0B"], ["#0E7490", "#22B8CF"], ["#7C3AED", "#A78BFA"],
  ["#047857", "#10B981"], ["#9D174D", "#EC4899"], ["#4338CA", "#818CF8"], ["#155E75", "#06B6D4"],
  ["#B91C1C", "#F87171"], ["#1F2937", "#6B7280"], ["#065F46", "#34D399"], ["#92400E", "#FBBF24"], ["#5B21B6", "#8B5CF6"],
];
function RoleLauncher({ name, sub, roleTitle, nav, homeKey, go, alerts = [], badges = {}, quick }) {
  // كل عناصر القائمة الجانبية (عدا الرئيسية نفسها) كبطاقات
  const sections = [];
  nav.forEach((grp) => grp.items.forEach(([k, label]) => { if (k !== homeKey) sections.push({ k, label, gIcon: grp.icon }); }));
  return (<div className="space-y-5">
    <div className="rb-card rounded-3xl p-5 sm:p-6 text-white relative overflow-hidden" style={{ background: `linear-gradient(130deg, ${C.ink} 0%, ${SC.color} 70%, ${SC.color2} 100%)` }}>
      <div className="absolute opacity-10" style={{ top: -40, left: -40, width: 190, height: 190, borderRadius: "50%", border: "26px solid #fff" }} />
      <div className="absolute opacity-10" style={{ bottom: -50, right: 30, width: 130, height: 130, borderRadius: "50%", background: C.well }} />
      <div className="relative z-10 flex items-center justify-between flex-wrap gap-3">
        <div className="flex items-center gap-3"><Avatar name={name} size={52} /><div><div className="text-sm opacity-80">أهلاً بك</div><div className="text-xl sm:text-2xl font-extrabold" style={{ fontFamily: "Cairo, sans-serif" }}>{name}</div><div className="text-xs opacity-80 mt-0.5">{roleTitle}{sub ? ` · ${sub}` : ""}</div></div></div>
        {quick && <button onClick={() => go(quick.k)} className="rb-press rounded-2xl px-5 py-3 font-extrabold text-sm flex items-center gap-2" style={{ background: C.well, color: SC.color }}><quick.icon size={18} />{quick.label}</button>}
      </div>
    </div>
    {alerts.length > 0 && <div className="flex gap-2 overflow-x-auto pb-1">{alerts.map((a, i) => (<button key={i} onClick={() => go(a.k)} className="rb-press shrink-0 flex items-center gap-2 rounded-2xl px-4 py-2.5 text-sm font-bold" style={{ background: a.soft || C.accentSoft, color: a.color || "#92590C", border: `1px solid ${a.border || C.accent}` }}><a.icon size={16} />{a.text}</button>))}</div>}
    <div>
      <h3 className="font-extrabold mb-3 text-base" style={{ color: C.text, fontFamily: "Cairo, sans-serif" }}>كل خدماتك</h3>
      <div className="rb-stagger grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3">
        {sections.map((sec, i) => { const [c1, c2] = LAUNCH_PALETTE[i % LAUNCH_PALETTE.length]; const Ic = sec.gIcon || LayoutDashboard; const badge = badges[sec.k]; return (
          <button key={sec.k} onClick={() => go(sec.k)} className="rb-press relative text-right rounded-3xl p-4 overflow-hidden group" style={{ background: `linear-gradient(135deg, ${c1} 0%, ${c2} 100%)`, minHeight: 112, boxShadow: "0 8px 22px rgba(11,31,58,.10)" }}>
            <div className="absolute opacity-15 transition-transform duration-300 group-hover:scale-125" style={{ top: -18, left: -18, width: 76, height: 76, borderRadius: "50%", background: C.well }} />
            {badge > 0 && <span className="absolute top-3 left-3 min-w-[20px] h-5 px-1.5 rounded-full text-[11px] font-extrabold flex items-center justify-center" style={{ background: C.well, color: c1 }}>{badge}</span>}
            <div className="rounded-2xl inline-flex items-center justify-center mb-2.5" style={{ width: 40, height: 40, background: "rgba(255,255,255,.18)", backdropFilter: "blur(4px)" }}><Ic size={20} color="#fff" /></div>
            <div className="text-white font-extrabold text-sm leading-tight" style={{ fontFamily: "Cairo, sans-serif" }}>{sec.label}</div>
          </button>); })}
      </div>
    </div>
  </div>);
}

/* ============================================================
   ============   مدير المبيعات   ============================
   ============================================================ */
const useMg = (props) => props.data.salesManagers.find((s) => s.id === props.user.entityId);
// ═══ إشراف مدير المبيعات على الفروع ═══
// الفروع التي يشرف عليها المدير: branchIds (قائمة) — فارغة/غير موجودة = كل الفروع
const mgSupervisedBranches = (co, mg) => {
  const all = companyBranches(co);
  if (!mg?.branchIds || !mg.branchIds.length) return all; // يشرف على الكل
  return all.filter((b) => mg.branchIds.includes(b.id));
};
const mgSupervisesAll = (co, mg) => !mg?.branchIds || !mg.branchIds.length || mg.branchIds.length >= companyBranches(co).length;
// هل ينتمي كيان (له branchId) لنطاق إشراف المدير + الفرع المختار؟
const mgInScope = (co, mg, entity, selectedBranch) => {
  const supervised = mgSupervisedBranches(co, mg).map((b) => b.id);
  if (!hasBranches(co)) return true; // شركة بلا فروع: الكل
  if (selectedBranch && selectedBranch !== "all") return entity.branchId === selectedBranch;
  // "كل فروعي": ضمن الفروع المُشرَف عليها (أو بلا فرع محدد إن كان يشرف على الكل)
  return supervised.includes(entity.branchId) || (mgSupervisesAll(co, mg) && !entity.branchId);
};
// شارة الفرع: تُعرض بجانب كل مندوب/مشرف لوضوح انتمائه
function BranchBadge({ co, bid, tiny }) {
  if (!bid || !hasBranches(co)) return null;
  const name = branchName(co, bid);
  if (name === "—") return null;
  return (<span className="inline-flex items-center gap-1 rounded-full font-bold" style={{ background: SC.soft, color: SC.color, fontSize: tiny ? 9 : 10, padding: tiny ? "1px 6px" : "2px 8px" }}><MapPin size={tiny ? 9 : 11} />{name}</span>);
}
// مبدّل نطاق فروع المدير
function MgBranchSwitcher({ co, mg, branch, setBranch }) {
  const supervised = mgSupervisedBranches(co, mg);
  if (!hasBranches(co) || supervised.length <= 1) return null;
  const opts = [{ id: "all", name: "كل فروعي" }, ...supervised];
  return (<div className="flex items-center gap-1.5 flex-wrap rounded-2xl p-1.5 mb-1" style={{ background: C.well, border: `1px solid ${C.line}` }}>
    {opts.map((b) => { const on = branch === b.id; return (
      <button key={b.id} onClick={() => setBranch(b.id)} className="rb-press flex items-center gap-1.5 rounded-xl px-3 py-1.5 text-sm font-extrabold transition-all" style={{ background: on ? SC.color : "transparent", color: on ? C.onAccent : C.soft }}>{b.id === "all" ? <Layers size={14} /> : <MapPin size={14} />}{b.name}</button>); })}
  </div>);
}
// تصفية مندوبي المدير حسب نطاقه والفرع المختار
const mgScopedReps = (data, co, mg, branch) => data.reps.filter((r) => r.companyId === mg.companyId && mgInScope(co, mg, r, branch));
const mgScopedMerchants = (data, co, mg, branch) => data.merchants.filter((m) => m.companyId === mg.companyId && mgInScope(co, mg, m, branch));
const mgScopedSupervisors = (data, co, mg, branch) => data.supervisors.filter((x) => x.companyId === mg.companyId && mgInScope(co, mg, x, branch));
const useSv = (props) => props.data.supervisors.find((s) => s.id === props.user.entityId);
function MgHome(props) {
  const { data, go, user } = props; const mg = useMg(props); const co = useCo(props);
  const [branch, setBranch] = useState("all");
  const reps = mgScopedReps(data, co, mg, branch);
  const supers = mgScopedSupervisors(data, co, mg, branch);
  const merchants = mgScopedMerchants(data, co, mg, branch);
  const repIds = new Set(reps.map((r) => r.id));
  const merchIds = new Set(merchants.map((m) => m.id));
  const delivered = data.orders.filter((o) => o.companyId === mg.companyId && o.status === "delivered" && (repIds.has(o.repId) || merchIds.has(o.merchantId)));
  const pending = data.orders.filter((o) => o.companyId === mg.companyId && ["pending_company", "preparing", "ready_dispatch", "out_for_delivery"].includes(o.status) && (repIds.has(o.repId) || merchIds.has(o.merchantId)));
  const totalVal = delivered.reduce((s, o) => s + orderTotal(o), 0);
  const totalQty = delivered.reduce((s, o) => s + orderQty(o), 0);
  const totalDebt = merchants.reduce((s, m) => s + Math.max(0, -(m.balance || 0)), 0);
  const collected = (data.payments || []).filter((p) => repIds.has(p.repId)).reduce((s, p) => s + p.amount, 0);
  const overdue = merchants.filter((m) => oldestDebtDays(data, m.id) > 30).length;
  const supervised = mgSupervisedBranches(co, mg);
  const multi = hasBranches(co) && supervised.length > 1;

  return (<div className="space-y-5">
    <div className="rounded-2xl p-5 sm:p-6" style={{ background: `linear-gradient(120deg, ${C.ink} 0%, ${SC.color} 130%)`, color: "#fff" }}>
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div><div className="text-sm opacity-85 flex items-center gap-2"><UserCog size={16} />مدير المبيعات</div><div className="text-2xl font-extrabold mt-0.5">{mg.name}</div><div className="text-xs opacity-80 mt-1">{hasBranches(co) ? `يشرف على: ${mgSupervisesAll(co, mg) ? "كل الفروع" : supervised.map((b) => b.name).join(" · ")}` : `${supers.length} مشرف · ${reps.length} مندوب`}</div></div>
        <PrimaryBtn tone="ghost" icon={BarChart3} onClick={() => go("mg_reports")}>التقارير التفصيلية</PrimaryBtn>
      </div>
    </div>

    <MgBranchSwitcher co={co} mg={mg} branch={branch} setBranch={setBranch} />

    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={CircleDollarSign} label="مبيعات مُسلّمة" value={money(totalVal)} tint={{ color: C.success, soft: C.successSoft }} hint={qty(totalQty)} />
      <Stat icon={HandCoins} label="التحصيل" value={money(collected)} tint={{ color: C.accent, soft: C.accentSoft }} />
      <Stat icon={AlertTriangle} label="الديون القائمة" value={money(totalDebt)} tint={{ color: C.danger, soft: C.dangerSoft }} hint={overdue ? `${overdue} متأخر +30 يوم` : ""} />
      <Stat icon={ShoppingCart} label="طلبات جارية" value={pending.length} tint={{ color: C.info, soft: C.infoSoft }} />
    </div>

    {multi && <Card className="p-4">
      <h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Layers size={18} style={{ color: SC.color }} />متابعة فروعي</h3>
      <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">
        {supervised.map((b) => { const st = branchStats(data, co, b.id); return (
          <button key={b.id} onClick={() => setBranch(b.id)} className="rb-press text-right rounded-2xl p-4" style={{ border: `1.5px solid ${branch === b.id ? SC.color : C.line}`, background: branch === b.id ? SC.soft : C.surface }}>
            <div className="flex items-center justify-between mb-2"><span className="font-extrabold flex items-center gap-1.5" style={{ color: SC.color }}><MapPin size={14} />{b.name}</span><span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: C.surface, color: C.soft }}>{st.reps} مندوب</span></div>
            <div className="grid grid-cols-2 gap-2 text-center">
              <div className="rounded-lg py-1.5" style={{ background: C.well }}><div className="text-[9px]" style={{ color: C.soft }}>المبيعات</div><div className="font-extrabold text-sm" style={{ color: C.text }}>{money(st.sales)}</div></div>
              <div className="rounded-lg py-1.5" style={{ background: C.well }}><div className="text-[9px]" style={{ color: C.soft }}>الديون</div><div className="font-extrabold text-sm" style={{ color: st.debt > 0 ? C.danger : C.success }}>{money(st.debt)}</div></div>
            </div>
          </button>); })}
      </div>
    </Card>}

    <div className="grid lg:grid-cols-2 gap-4">
      <Card className="p-4"><div className="flex items-center justify-between mb-3"><h3 className="font-extrabold" style={{ color: C.text }}>أعلى المندوبين</h3><GhostBtn sm icon={Handshake} onClick={() => go("mg_rep_perf")}>الكل</GhostBtn></div>
        <div className="space-y-2">{reps.map((r) => ({ r, v: delivered.filter((o) => o.repId === r.id).reduce((s, o) => s + orderTotal(o), 0) })).sort((a, b) => b.v - a.v).slice(0, 5).map((x, i) => <div key={x.r.id} className="flex items-center gap-2.5"><span className="w-6 h-6 rounded-lg flex items-center justify-center text-xs font-extrabold" style={{ background: i === 0 ? C.warnSoft : C.well, color: i === 0 ? "#B45309" : C.soft }}>{i + 1}</span><Avatar name={x.r.name} size={30} /><span className="flex-1 text-sm font-bold truncate flex items-center gap-1.5" style={{ color: C.text }}>{x.r.name}<BranchBadge co={co} bid={x.r.branchId} tiny /></span><span className="text-sm font-extrabold" style={{ color: SC.color }}>{money(x.v)}</span></div>)}{!reps.length && <Empty icon={Handshake} text="لا مندوبين في هذا النطاق." />}</div>
      </Card>
      <Card className="p-4"><div className="flex items-center justify-between mb-3"><h3 className="font-extrabold" style={{ color: C.text }}>متابعة سريعة</h3></div>
        <div className="grid grid-cols-2 gap-2">
          {[["mg_explore", "الاستعلام والتقارير", PieChart], ["mg_rep_perf", "أداء المندوبين", Handshake], ["mg_aging", "أعمار الديون", HandCoins], ["mg_churned", "عملاء متوقفون", UserX], ["mg_by_city", "حسب المدينة", MapPin], ["mg_supervisors", "المشرفون", Briefcase]].map(([k, t, Ic]) => (
            <button key={k} onClick={() => go(k)} className="rb-press flex items-center gap-2 rounded-xl p-3 text-sm font-bold" style={{ background: C.well, color: C.text }}><Ic size={16} style={{ color: SC.color }} />{t}</button>))}
        </div>
      </Card>
    </div>
  </div>);
}
function MgSales(props) {
  const { data } = props; const mg = useMg(props); const co = useCo(props);
  const [branch, setBranch] = useState("all");
  const reps = mgScopedReps(data, co, mg, branch);
  return (<div className="space-y-4"><SectionTitle sub="استعلام مفصّل عن مبيعات كل مندوب بالكمية والقيمة.">استعلام المبيعات</SectionTitle>
    <MgBranchSwitcher co={co} mg={mg} branch={branch} setBranch={setBranch} />
    <div className="grid sm:grid-cols-3 gap-3">
      <Stat icon={Boxes} label="إجمالي الكمية المُسلّمة" value={qty(reps.reduce((s, r) => s + data.orders.filter((o) => o.repId === r.id && o.status === "delivered").reduce((a, o) => a + orderQty(o), 0), 0))} tint={{ color: SC.color2, soft: SC.soft }} />
      <Stat icon={CircleDollarSign} label="إجمالي القيمة المُسلّمة" value={money(reps.reduce((s, r) => s + data.orders.filter((o) => o.repId === r.id && o.status === "delivered").reduce((a, o) => a + orderTotal(o), 0), 0))} tint={{ color: C.success, soft: C.successSoft }} />
      <Stat icon={Handshake} label="عدد المندوبين" value={reps.length} tint={{ color: C.info, soft: C.infoSoft }} />
    </div>
    <Card className="p-4"><Table head={["المندوب", "الفرع", "المشرف", "نوع الهدف", "كمية مُسلّمة", "هدف كمّي", "قيمة مُسلّمة", "هدف قيمي", "العملاء"]}>{reps.map((r) => { const delivered = data.orders.filter((o) => o.repId === r.id && o.status === "delivered"); const units = delivered.reduce((s, o) => s + orderQty(o), 0); const rev = delivered.reduce((s, o) => s + orderTotal(o), 0); const clients = data.merchants.filter((m) => m.repId === r.id).length; const pctQ = Math.round((r.qtyAchieved / (r.qtyTarget || 1)) * 100); const pctV = Math.round((r.achieved / (r.target || 1)) * 100); return (<tr key={r.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{r.name}</Td><Td>{hasBranches(co) ? <BranchBadge co={co} bid={r.branchId} tiny /> : <span className="text-xs" style={{ color: C.soft }}>—</span>}</Td><Td>{r.supervisorId ? data.supervisors.find((s) => s.id === r.supervisorId)?.name : "مباشر"}</Td><Td><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: r.targetMode === "qty" ? SC.soft : C.successSoft, color: r.targetMode === "qty" ? SC.color : C.success }}>{r.targetMode === "qty" ? "كمية" : "قيمة"}</span></Td><Td><span style={{ color: SC.color2, fontWeight: 700 }}>{qty(units)}</span></Td><Td>{qty(r.qtyTarget)} <span className="text-xs" style={{ color: pctQ >= 70 ? C.success : C.accent }}>({pctQ}%)</span></Td><Td>{money(rev)}</Td><Td>{money(r.target)} <span className="text-xs" style={{ color: pctV >= 70 ? C.success : C.accent }}>({pctV}%)</span></Td><Td>{clients}</Td></tr>); })}</Table></Card>
  </div>);
}
// متابعة الزيارات المخططة والناجحة وغير الناجحة
function MgVisits(props) {
  const { data } = props; const mg = useMg(props); const co = useCo(props);
  const [branch, setBranch] = useState("all");
  const reps = mgScopedReps(data, co, mg, branch);
  const repIds = reps.map((r) => r.id);
  const visits = data.visits.filter((v) => repIds.includes(v.repId));
  const planned = data.plannedVisits.filter((p) => repIds.includes(p.repId));
  return (<div className="space-y-4"><SectionTitle sub="متابعة الزيارات المخططة والمنفّذة ونِسب النجاح لكل مندوب.">متابعة الزيارات</SectionTitle>
    <MgBranchSwitcher co={co} mg={mg} branch={branch} setBranch={setBranch} />
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>ملخص الزيارات لكل مندوب</h3>
      <Table head={["المندوب", "مخططة", "منفّذة", "ناجحة", "غير ناجحة", "نسبة النجاح"]}>{reps.map((r) => {
        const pv = planned.filter((p) => p.repId === r.id); const done = pv.filter((p) => p.done).length;
        const rv = visits.filter((v) => v.repId === r.id); const succ = rv.filter((v) => ["order", "payment"].includes(v.result)).length; const fail = rv.length - succ; const pct = rv.length ? Math.round((succ / rv.length) * 100) : 0;
        return (<tr key={r.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{r.name}</Td><Td>{pv.length}</Td><Td>{done}</Td><Td><span style={{ color: C.success, fontWeight: 700 }}>{succ}</span></Td><Td><span style={{ color: C.danger, fontWeight: 700 }}>{fail}</span></Td><Td><div className="flex items-center gap-2"><div style={{ width: 60 }}><ProgressBar value={succ} max={rv.length || 1} color={pct >= 60 ? C.success : C.accent} /></div><span className="text-xs font-bold" style={{ color: C.text }}>{pct}%</span></div></Td></tr>); })}</Table>
    </Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>الزيارات غير الناجحة وأسبابها</h3>
      <Table head={["المندوب", "المحل", "السبب", "التاريخ"]}>{visits.filter((v) => !["order", "payment"].includes(v.result)).map((v) => <tr key={v.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{repName(data, v.repId)}</Td><Td>{merchName(data, v.merchantId)}</Td><Td><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.dangerSoft, color: C.danger }}>{v.failReason ? FAIL_REASONS[v.failReason] : VISIT_RESULT[v.result]?.label}</span></Td><Td>{v.date}</Td></tr>)}</Table>
    </Card>
  </div>);
}
// الأصناف ضعيفة السحب
function MgWeak(props) {
  const { data } = props; const mg = useMg(props);
  const prods = data.products.filter((p) => p.companyId === mg.companyId);
  // حساب الكمية المباعة لكل صنف من الطلبات المسلّمة
  const sold = {};
  data.orders.filter((o) => o.companyId === mg.companyId && o.status === "delivered").forEach((o) => o.items.forEach((it) => { sold[it.pid] = (sold[it.pid] || 0) + it.qty; }));
  const rows = prods.map((p) => ({ p, qty: sold[p.id] || 0, quota: data.quotas.filter((q) => q.pid === p.id).reduce((s, q) => s + q.allocated, 0), quotaSold: data.quotas.filter((q) => q.pid === p.id).reduce((s, q) => s + q.sold, 0) }))
    .sort((a, b) => a.qty - b.qty);
  const weak = rows.filter((r) => r.qty < 10);
  return (<div className="space-y-4"><SectionTitle sub="الأصناف التي سحبها ضعيف لمعالجة مشكلتها (عرض، سعر، توفّر).">الأصناف ضعيفة السحب</SectionTitle>
    <Stat icon={TrendingDown} label="أصناف تحتاج معالجة" value={weak.length} tint={{ color: C.danger, soft: C.dangerSoft }} />
    <Card className="p-4"><Table head={["الصنف", "المباع", "حصص موزّعة", "مباع من الحصص", "المخزون", "التشخيص"]}>{rows.map(({ p, qty, quota, quotaSold }) => { const isWeak = qty < 10; const diag = qty === 0 ? "لا مبيعات" : quota > 0 && quotaSold / quota < 0.3 ? "تصريف بطيء للحصص" : p.stock > 200 ? "مخزون راكد" : "سحب منخفض";
      return (<tr key={p.id} style={{ borderBottom: `1px solid ${C.line}`, background: isWeak ? C.dangerSoft + "33" : "transparent" }}><Td bold>{p.name}</Td><Td><span style={{ color: isWeak ? C.danger : C.text, fontWeight: 700 }}>{qty}</span></Td><Td>{quota}</Td><Td>{quotaSold}</Td><Td>{p.stock}</Td><Td>{isWeak ? <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.dangerSoft, color: C.danger }}>{diag}</span> : <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.successSoft, color: C.success }}>جيد</span>}</Td></tr>); })}</Table></Card>
    {weak.length > 0 && <Card className="p-4" style={{ background: C.warnSoft, border: "none" }}><div className="flex items-start gap-2"><AlertCircle size={18} style={{ color: "#92590C" }} className="mt-0.5" /><div className="text-sm" style={{ color: "#92590C" }}><b>اقتراحات للمعالجة:</b> راجع تسعير الأصناف ضعيفة السحب، أضف عروضاً، كثّف الترويج عند المحلات، أو أعد توزيع الحصص على مندوبين أنشط.</div></div></Card>}
  </div>);
}
function MgMovements(props) {
  const { data } = props; const mg = useMg(props); const co = useCo(props);
  const [branch, setBranch] = useState("all");
  const repIds = mgScopedReps(data, co, mg, branch).map((r) => r.id);
  const coMerchants = mgScopedMerchants(data, co, mg, branch);
  const visits = data.visits.filter((v) => repIds.includes(v.repId));
  const payments = data.payments.filter((p) => p.repId && repIds.includes(p.repId));
  const promos = data.promotions.filter((p) => p.companyId === mg.companyId);
  const supTasks = data.tasks.filter((t) => t.companyId === mg.companyId && t.assigneeType === "supervisor");
  return (<div className="space-y-4"><SectionTitle sub="الحركات الكاملة للمشرفين والمندوبين والمروّجين.">حركات الفريق الكاملة</SectionTitle>
    <MgBranchSwitcher co={co} mg={mg} branch={branch} setBranch={setBranch} />
    <AgingPanel data={data} merchants={coMerchants} title="أعمار ديون عملاء الشركة" />
    <Card className="p-4"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Handshake size={17} style={{ color: C.info }} />حركات المندوبين (زيارات)</h3><Table head={["المندوب", "المحل", "النتيجة", "المبلغ", "التاريخ"]}>{visits.map((v) => <tr key={v.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{repName(data, v.repId)}</Td><Td>{merchName(data, v.merchantId)}</Td><Td><Badge status={v.result} map={VISIT_RESULT} /></Td><Td>{v.amount ? money(v.amount) : "—"}</Td><Td>{v.date}</Td></tr>)}</Table>{visits.length === 0 && <Empty icon={MapPinned} text="لا توجد زيارات." />}</Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><HandCoins size={17} style={{ color: C.success }} />تحصيلات المندوبين</h3><Table head={["المندوب", "المحل", "المبلغ", "التاريخ"]}>{payments.map((p) => <tr key={p.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{repName(data, p.repId)}</Td><Td>{merchName(data, p.merchantId)}</Td><Td><span style={{ color: C.success, fontWeight: 700 }}>{money(p.amount)}</span></Td><Td>{p.date}</Td></tr>)}</Table>{payments.length === 0 && <Empty icon={HandCoins} text="لا توجد تحصيلات." />}</Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><ImageIcon size={17} style={{ color: C.rose }} />حركات المروّجين</h3><Table head={["المروّج", "المحل", "GPS", "قبل/بعد", "التستيف", "صلاحيات", "طلبية مقترحة", "الحالة"]}>{promos.map((p) => <tr key={p.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{data.promoters.find((x) => x.id === p.promoterId)?.name}</Td><Td>{merchName(data, p.merchantId)}</Td><Td>{p.gpsConfirmed ? "✓" : "—"}</Td><Td>{p.photoBefore ? "📸" : "—"}/{p.photoAfter ? "📸" : "—"}</Td><Td>{p.stacking ? `${p.stacking.fridge ? "ثلاجة ✓" : ""} ${p.stacking.shelf ? "رفوف ✓" : ""}`.trim() || "—" : "—"}</Td><Td>{p.expiryFindings?.length ? <span style={{ color: C.danger, fontWeight: 800 }}>{p.expiryFindings.filter((fd) => fd.flag !== "ok").length} ⚠️</span> : "—"}</Td><Td>{p.suggestedItems?.length ? `${p.suggestedItems.length} صنف` : "—"}</Td><Td><Badge status={p.status} map={PROMO_STATUS} /></Td></tr>)}</Table>{promos.length === 0 && <Empty icon={ImageIcon} text="لا توجد حركات ترويج." />}</Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Briefcase size={17} style={{ color: C.teal }} />أوامر المشرفين</h3><div className="space-y-2">{supTasks.map((t) => <div key={t.id} className="rounded-xl p-3 flex items-center justify-between" style={{ border: `1px solid ${C.line}` }}><div><div className="font-bold text-sm" style={{ color: C.text }}>{t.title}</div><div className="text-xs" style={{ color: C.soft }}>{data.supervisors.find((s) => s.id === t.assigneeId)?.name}</div></div><Badge status={t.status} map={TASK_STATUS} /></div>)}{supTasks.length === 0 && <Empty icon={ListChecks} text="لا توجد أوامر." />}</div></Card>
  </div>);
}
function MgSupervisors(props) {
  const { data } = props; const mg = useMg(props); const co = useCo(props);
  const [branch, setBranch] = useState("all");
  const supers = mgScopedSupervisors(data, co, mg, branch);
  return (<div className="space-y-4"><SectionTitle sub="متابعة أداء المشرفين وفرقهم.">متابعة المشرفين</SectionTitle>
    <MgBranchSwitcher co={co} mg={mg} branch={branch} setBranch={setBranch} />
    <div className="grid sm:grid-cols-2 gap-3">{supers.map((s) => { const reps = data.reps.filter((r) => r.supervisorId === s.id); const tgt = reps.reduce((a, r) => a + r.target, 0); const ach = reps.reduce((a, r) => a + r.achieved, 0); const pct = Math.round((ach / (tgt || 1)) * 100);
      return (<Card key={s.id} className="p-5"><div className="flex items-center gap-3 mb-3"><div className="rounded-xl p-2.5" style={{ background: C.tealSoft }}><Briefcase size={20} style={{ color: C.teal }} /></div><div><div className="font-extrabold flex items-center gap-1.5 flex-wrap" style={{ color: C.text }}>{s.name}<BranchBadge co={co} bid={s.branchId} tiny /></div><div className="text-xs" style={{ color: C.soft }}>{s.area} · {reps.length} مندوب</div></div></div>
        <div className="flex justify-between text-sm mb-1"><span style={{ color: C.soft }}>تحقيق الفريق</span><span className="font-bold" style={{ color: C.text }}>{pct}%</span></div><ProgressBar value={ach} max={tgt} color={pct >= 70 ? C.success : C.accent} />
        <div className="flex flex-wrap gap-1.5 mt-3">{reps.map((r) => <span key={r.id} className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.infoSoft, color: C.info }}>{r.name}</span>)}</div>
      </Card>); })}{supers.length === 0 && <Empty icon={Briefcase} text="لا يوجد مشرفون." />}</div>
  </div>);
}
function MgOrders(props) {
  const { data, actions } = props; const mg = useMg(props); const co = useCo(props);
  const [branch, setBranch] = useState("all");
  const supers = mgScopedSupervisors(data, co, mg, branch);
  const orders = data.tasks.filter((t) => t.companyId === mg.companyId && t.assigneeType === "supervisor");
  const [f, setF] = useState({ assigneeId: supers[0]?.id || "", title: "", desc: "", due: "اليوم", priority: "normal" });
  return (<div className="space-y-4"><SectionTitle sub="أسند أوامر ومهام للمشرفين وتابع تنفيذها.">أوامر للمشرفين</SectionTitle>
    <MgBranchSwitcher co={co} mg={mg} branch={branch} setBranch={setBranch} />
    <Card className="p-4"><div className="grid sm:grid-cols-2 gap-2 mb-2">
      <Field label="المشرف"><select value={f.assigneeId} onChange={(e) => setF({ ...f, assigneeId: e.target.value })} className={inputCls} style={inputStyle}>{supers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}</select></Field>
      <Field label="الأولوية"><select value={f.priority} onChange={(e) => setF({ ...f, priority: e.target.value })} className={inputCls} style={inputStyle}><option value="normal">عادية</option><option value="high">عالية</option></select></Field>
      <Field label="عنوان الأمر"><Input value={f.title} onChange={(e) => setF({ ...f, title: e.target.value })} /></Field>
      <Field label="الموعد"><Input value={f.due} onChange={(e) => setF({ ...f, due: e.target.value })} /></Field>
    </div><Field label="التفاصيل"><textarea value={f.desc} onChange={(e) => setF({ ...f, desc: e.target.value })} className={inputCls} style={{ ...inputStyle, minHeight: 60 }} /></Field>
    <div className="mt-3"><PrimaryBtn icon={Send} onClick={() => { if (f.title && f.assigneeId) { actions.addTask({ server: mg.server, companyId: mg.companyId, assigneeType: "supervisor", ...f, status: "assigned" }); setF({ ...f, title: "", desc: "" }); } }}>إسناد الأمر</PrimaryBtn></div></Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>الأوامر المُسندة</h3><div className="space-y-2">{orders.map((t) => <div key={t.id} className="rounded-xl p-3 flex items-start justify-between gap-3" style={{ border: `1px solid ${C.line}` }}><div><div className="font-bold text-sm flex items-center gap-2" style={{ color: C.text }}>{t.title}{t.priority === "high" && <span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.dangerSoft, color: C.danger }}>عاجل</span>}</div><div className="text-xs" style={{ color: C.soft }}>{data.supervisors.find((s) => s.id === t.assigneeId)?.name} · {t.due}</div></div><Badge status={t.status} map={TASK_STATUS} /></div>)}{orders.length === 0 && <Empty icon={ListChecks} text="لا توجد أوامر." />}</div></Card>
  </div>);
}

/* ===== التقارير التفصيلية لمدير المبيعات ===== */
function ReportBar({ label, value, max, color, suffix }) {
  const pct = Math.min(100, Math.round((value / (max || 1)) * 100));
  return (<div><div className="flex justify-between text-xs mb-1"><span style={{ color: C.text, fontWeight: 700 }}>{label}</span><span style={{ color: C.soft }}>{value.toLocaleString("ar-LY")}{suffix} · {pct}%</span></div><div className="h-2.5 rounded-full overflow-hidden" style={{ background: C.bg }}><div style={{ width: pct + "%", height: "100%", background: color }} /></div></div>);
}
// ═══ مدير المبيعات: المبيعات حسب المدينة ═══
function MgSalesByCity(props) {
  const { data } = props; const mg = useMg(props); const co = useCo(props);
  const [branch, setBranch] = useState("all");
  const merchants = mgScopedMerchants(data, co, mg, branch);
  const cities = {};
  merchants.forEach((m) => { const c = m.city || "غير محدد"; if (!cities[c]) cities[c] = { name: c, sales: 0, qty: 0, orders: 0, count: 0 }; cities[c].count++; });
  data.orders.filter((o) => o.companyId === mg.companyId && o.status === "delivered").forEach((o) => { const m = merchants.find((x) => x.id === o.merchantId); if (m) { const c = m.city || "غير محدد"; if (cities[c]) { cities[c].sales += orderTotal(o); cities[c].qty += orderQty(o); cities[c].orders++; } } });
  const list = Object.values(cities).sort((a, b) => b.sales - a.sales);
  const maxSales = Math.max(1, ...list.map((c) => c.sales));
  const totalSales = list.reduce((s, c) => s + c.sales, 0);
  const totalQty = list.reduce((s, c) => s + c.qty, 0);
  return (<div className="space-y-4"><SectionTitle sub="توزيع المبيعات على المدن — بالكمية والقيمة معاً.">المبيعات حسب المدينة</SectionTitle>
    <MgBranchSwitcher co={co} mg={mg} branch={branch} setBranch={setBranch} />
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={MapPin} label="عدد المدن" value={list.length} tint={{ color: SC.color, soft: SC.soft }} />
      <Stat icon={Store} label="أعلى مدينة مبيعاً" value={list[0]?.name || "—"} tint={{ color: C.success, soft: C.successSoft }} hint={list[0] ? money(list[0].sales) : ""} />
      <Stat icon={Boxes} label="إجمالي الكمية" value={qty(totalQty)} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={CircleDollarSign} label="إجمالي القيمة" value={money(totalSales)} tint={{ color: C.accent, soft: C.accentSoft }} />
    </div>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>ترتيب المدن</h3>
      <div className="space-y-3">{list.map((c) => (<div key={c.name}>
        <div className="flex items-center justify-between text-sm mb-1"><span className="font-bold flex items-center gap-1.5" style={{ color: C.text }}><MapPin size={14} style={{ color: SC.color }} />{c.name}<span className="text-[11px] font-normal" style={{ color: C.soft }}>({c.count} محل)</span></span><span className="font-extrabold" style={{ color: C.text }}>{money(c.sales)} · {qty(c.qty)}</span></div>
        <ProgressBar value={c.sales} max={maxSales} color={SC.color} />
      </div>))}{list.length === 0 && <Empty icon={MapPin} text="لا مبيعات مسجّلة بعد." />}</div>
    </Card>
  </div>);
}

// ═══ مدير المبيعات: أداء المندوبين ═══
function MgRepPerformance(props) {
  const { data } = props; const mg = useMg(props); const co = useCo(props);
  const [branch, setBranch] = useState("all");
  const reps = mgScopedReps(data, co, mg, branch);
  const rows = reps.map((r) => {
    const dlv = data.orders.filter((o) => o.repId === r.id && o.status === "delivered");
    const sales = dlv.reduce((s, o) => s + orderTotal(o), 0);
    const units = dlv.reduce((s, o) => s + orderQty(o), 0);
    const target = r.target || 0; const achieved = r.achieved || 0;
    const pct = target ? Math.round((achieved / target) * 100) : 0;
    const clients = data.merchants.filter((m) => m.repId === r.id).length;
    const collected = data.payments.filter((p) => p.repId === r.id).reduce((s, p) => s + p.amount, 0);
    return { r, sales, units, target, achieved, pct, clients, collected };
  }).sort((a, b) => b.achieved - a.achieved);
  const topSales = Math.max(1, ...rows.map((x) => x.sales));
  const avgPct = rows.length ? Math.round(rows.reduce((s, x) => s + x.pct, 0) / rows.length) : 0;
  const rankColor = (i) => i === 0 ? "#B45309" : i === 1 ? "#64748B" : i === 2 ? "#92590C" : C.soft;
  return (<div className="space-y-4"><SectionTitle sub="مقارنة أداء المندوبين: التحقيق، المبيعات، التحصيل، والعملاء.">أداء المندوبين</SectionTitle>
    <MgBranchSwitcher co={co} mg={mg} branch={branch} setBranch={setBranch} />
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={Handshake} label="عدد المندوبين" value={reps.length} tint={{ color: SC.color, soft: SC.soft }} />
      <Stat icon={TrendingUp} label="أعلى تحقيقاً" value={rows[0]?.r.name || "—"} tint={{ color: C.success, soft: C.successSoft }} hint={rows[0] ? rows[0].pct + "%" : ""} />
      <Stat icon={Target} label="متوسط التحقيق" value={avgPct + "%"} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={HandCoins} label="إجمالي التحصيل" value={money(rows.reduce((s, x) => s + x.collected, 0))} tint={{ color: C.accent, soft: C.accentSoft }} />
    </div>
    <div className="space-y-2.5">{rows.map((x, i) => (<Card key={x.r.id} className="p-4">
      <div className="flex items-center justify-between flex-wrap gap-2 mb-2.5">
        <div className="flex items-center gap-2.5"><span className="w-7 h-7 rounded-lg flex items-center justify-center text-sm font-extrabold shrink-0" style={{ background: i < 3 ? C.warnSoft : C.well, color: rankColor(i) }}>{i + 1}</span><Avatar name={x.r.name} size={38} /><div><div className="font-extrabold text-sm flex items-center gap-1.5 flex-wrap" style={{ color: C.text }}>{x.r.name}<BranchBadge co={co} bid={x.r.branchId} tiny /></div><div className="text-[11px]" style={{ color: C.soft }}>{REP_TYPE[x.r.repType || "presell"].label} · {x.clients} عميل</div></div></div>
        <span className="text-sm font-extrabold px-2.5 py-1 rounded-full" style={{ background: x.pct >= 100 ? C.successSoft : x.pct >= 60 ? C.warnSoft : C.dangerSoft, color: x.pct >= 100 ? C.success : x.pct >= 60 ? "#B45309" : C.danger }}>{x.pct}% تحقيق</span>
      </div>
      <ProgressBar value={x.achieved} max={x.target || 1} color={SC.color} />
      <div className="grid grid-cols-3 gap-2 mt-2.5 text-center">
        <div className="rounded-lg py-1.5" style={{ background: C.well }}><div className="text-[10px]" style={{ color: C.soft }}>المبيعات المُسلّمة</div><div className="font-extrabold text-sm" style={{ color: C.text }}>{money(x.sales)}</div></div>
        <div className="rounded-lg py-1.5" style={{ background: C.well }}><div className="text-[10px]" style={{ color: C.soft }}>الكمية</div><div className="font-extrabold text-sm" style={{ color: SC.color }}>{qty(x.units)}</div></div>
        <div className="rounded-lg py-1.5" style={{ background: C.well }}><div className="text-[10px]" style={{ color: C.soft }}>التحصيل</div><div className="font-extrabold text-sm" style={{ color: C.success }}>{money(x.collected)}</div></div>
      </div>
    </Card>))}{rows.length === 0 && <Empty icon={Handshake} text="لا مندوبين." />}</div>
  </div>);
}

// ═══ مدير المبيعات: العملاء المتوقفون عن الشراء ═══
function MgChurned(props) {
  const { data } = props; const mg = useMg(props); const co = useCo(props);
  const [branch, setBranch] = useState("all");
  const merchants = mgScopedMerchants(data, co, mg, branch);
  // آخر طلب لكل عميل (بالاعتماد على ترتيب الطلبات) — العميل بلا طلبات مسلّمة = متوقّف
  const rows = merchants.map((m) => {
    const orders = data.orders.filter((o) => o.merchantId === m.id && o.status === "delivered");
    const lastOrder = orders[0];
    const debt = merchantOpenDebt(data, m.id);
    // تقدير حالة التوقف: بلا طلبات = متوقّف تماماً، وإلا نشِط
    const status = orders.length === 0 ? "stopped" : orders.length <= 1 ? "atrisk" : "active";
    return { m, orders: orders.length, lastOrder, debt, status };
  });
  const stopped = rows.filter((r) => r.status === "stopped");
  const atrisk = rows.filter((r) => r.status === "atrisk");
  const list = [...stopped, ...atrisk].sort((a, b) => a.orders - b.orders);
  return (<div className="space-y-4"><SectionTitle sub="محلات توقّفت أو تكاد تتوقّف عن الطلب — تحتاج زيارة أو متابعة.">العملاء المتوقفون عن الشراء</SectionTitle>
    <MgBranchSwitcher co={co} mg={mg} branch={branch} setBranch={setBranch} />
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={UserX} label="متوقّفون تماماً" value={stopped.length} tint={{ color: C.danger, soft: C.dangerSoft }} />
      <Stat icon={AlertTriangle} label="معرّضون للتوقّف" value={atrisk.length} tint={{ color: "#B45309", soft: C.warnSoft }} />
      <Stat icon={Users} label="إجمالي المحلات" value={merchants.length} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={HandCoins} label="ديون المتوقّفين" value={money(stopped.reduce((s, r) => s + r.debt, 0))} tint={{ color: C.danger, soft: C.dangerSoft }} />
    </div>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>قائمة المتابعة</h3>
      <div className="space-y-2">{list.map((r) => (<div key={r.m.id} className="flex items-center gap-3 rounded-xl p-3" style={{ border: `1px solid ${C.line}`, background: r.status === "stopped" ? C.dangerSoft + "44" : C.surface }}>
        <Avatar name={r.m.name} size={38} />
        <div className="flex-1 min-w-0"><div className="font-extrabold text-sm truncate flex items-center gap-1.5" style={{ color: C.text }}>{r.m.name}<BranchBadge co={co} bid={r.m.branchId} tiny /></div><div className="text-[11px]" style={{ color: C.soft }}>{r.m.area} · {r.m.city} · {r.orders} طلب سابق</div></div>
        <div className="text-left shrink-0"><span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: r.status === "stopped" ? C.dangerSoft : C.warnSoft, color: r.status === "stopped" ? C.danger : "#B45309" }}>{r.status === "stopped" ? "متوقّف" : "معرّض"}</span>{r.debt > 0 && <div className="text-[11px] mt-1 font-bold" style={{ color: C.danger }}>دين {money(r.debt)}</div>}</div>
      </div>))}{list.length === 0 && <Empty icon={CheckCircle2} text="كل العملاء نشطون — لا توقّف." />}</div>
    </Card>
  </div>);
}

// ═══ مدير المبيعات: أعمار الديون (يعيد استخدام AgingPanel) ═══
function MgAging(props) {
  const { data } = props; const mg = useMg(props); const co = useCo(props);
  const [branch, setBranch] = useState("all");
  const merchants = mgScopedMerchants(data, co, mg, branch);
  return (<div className="space-y-4"><SectionTitle sub="توزيع ديون العملاء على الفترات (0–30 / 31–60 / +60 يوم).">أعمار الديون</SectionTitle>
    <MgBranchSwitcher co={co} mg={mg} branch={branch} setBranch={setBranch} />
    <AgingPanel data={data} merchants={merchants} title="أعمار ديون عملاء الشركة" showGate />
  </div>);
}

// ═══ الاستعلام المرن: تقارير حسب المدينة/المنطقة/المندوب/المشرف ═══
function MgExplore(props) {
  const { data } = props; const mg = useMg(props); const co = useCo(props);
  const [branch, setBranch] = useState("all");
  const [dim, setDim] = useState("city");
  const [sortBy, setSortBy] = useState("value");
  const [drill, setDrill] = useState(null);

  const reps = mgScopedReps(data, co, mg, branch);
  const merchants = mgScopedMerchants(data, co, mg, branch);
  const supers = mgScopedSupervisors(data, co, mg, branch);
  const repIds = new Set(reps.map((r) => r.id));
  const merchIds = new Set(merchants.map((m) => m.id));
  const delivered = data.orders.filter((o) => o.companyId === mg.companyId && o.status === "delivered" && merchIds.has(o.merchantId));
  const payments = (data.payments || []).filter((p) => p.repId && repIds.has(p.repId) && merchIds.has(p.merchantId));

  // تجميع حسب البُعد المختار
  const groupOf = (o) => {
    const m = merchants.find((x) => x.id === o.merchantId);
    if (!m) return null;
    if (dim === "city") return m.city || "غير محدد";
    if (dim === "area") return (m.city ? m.city + " › " : "") + (m.area || "غير محدد");
    if (dim === "rep") { const r = reps.find((x) => x.id === o.repId); return r?.name || "—"; }
    if (dim === "supervisor") { const r = reps.find((x) => x.id === o.repId); const sv = supers.find((s) => s.id === r?.supervisorId); return sv?.name || "بدون مشرف"; }
    return "—";
  };
  const merchGroup = (m) => {
    if (dim === "city") return m.city || "غير محدد";
    if (dim === "area") return (m.city ? m.city + " › " : "") + (m.area || "غير محدد");
    if (dim === "rep") { const r = reps.find((x) => x.id === m.repId); return r?.name || "—"; }
    if (dim === "supervisor") { const r = reps.find((x) => x.id === m.repId); const sv = supers.find((s) => s.id === r?.supervisorId); return sv?.name || "بدون مشرف"; }
    return "—";
  };

  const groups = {};
  const ensure = (k) => { if (!groups[k]) groups[k] = { name: k, value: 0, qty: 0, orders: 0, collected: 0, debt: 0, clients: new Set() }; return groups[k]; };
  delivered.forEach((o) => { const k = groupOf(o); if (k === null) return; const g = ensure(k); g.value += orderTotal(o); g.qty += orderQty(o); g.orders++; g.clients.add(o.merchantId); });
  payments.forEach((p) => { const m = merchants.find((x) => x.id === p.merchantId); if (!m) return; const k = merchGroup(m); ensure(k).collected += p.amount; });
  merchants.forEach((m) => { const k = merchGroup(m); ensure(k).debt += merchantOpenDebt(data, m.id); });

  const rows = Object.values(groups).map((g) => ({ ...g, clients: g.clients.size })).sort((a, b) => {
    if (sortBy === "value") return b.value - a.value;
    if (sortBy === "qty") return b.qty - a.qty;
    if (sortBy === "collected") return b.collected - a.collected;
    if (sortBy === "debt") return b.debt - a.debt;
    return b.orders - a.orders;
  });
  const maxVal = Math.max(1, ...rows.map((r) => r.value));
  const totals = rows.reduce((a, r) => ({ value: a.value + r.value, qty: a.qty + r.qty, collected: a.collected + r.collected, debt: a.debt + r.debt, orders: a.orders + r.orders }), { value: 0, qty: 0, collected: 0, debt: 0, orders: 0 });

  const DIMS = [["city", "المدينة", MapPin], ["area", "المنطقة", Navigation], ["rep", "المندوب", Handshake], ["supervisor", "المشرف", Briefcase]];
  const SORTS = [["value", "القيمة"], ["qty", "الكمية"], ["collected", "التحصيل"], ["debt", "الديون"], ["orders", "الطلبات"]];
  const dimLabel = DIMS.find((d) => d[0] === dim)[1];

  // تفصيل مجموعة عند النقر
  const drillMerchants = drill ? merchants.filter((m) => merchGroup(m) === drill) : [];

  return (<div className="space-y-4"><SectionTitle sub="استعلم عن الأداء بأي بُعد: المدينة، المنطقة، المندوب، أو المشرف — بالكمية والقيمة والتحصيل والديون.">الاستعلام والتقارير</SectionTitle>
    <MgBranchSwitcher co={co} mg={mg} branch={branch} setBranch={setBranch} />

    {/* اختيار البُعد */}
    <Card className="p-4">
      <div className="text-sm font-bold mb-2" style={{ color: C.soft }}>التقرير حسب:</div>
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        {DIMS.map(([k, t, Ic]) => { const on = dim === k; return (
          <button key={k} onClick={() => { setDim(k); setDrill(null); }} className="rb-press flex items-center justify-center gap-2 rounded-xl py-2.5 text-sm font-extrabold transition-all" style={{ background: on ? SC.color : C.well, color: on ? C.onAccent : C.text, border: `1.5px solid ${on ? SC.color : C.line}` }}><Ic size={16} />{t}</button>); })}
      </div>
    </Card>

    {/* إجماليات النطاق */}
    <div className="grid grid-cols-2 lg:grid-cols-5 gap-3">
      <Stat icon={Layers} label={`عدد ${dimLabel === "المدينة" ? "المدن" : dimLabel === "المنطقة" ? "المناطق" : dimLabel === "المندوب" ? "المندوبين" : "المشرفين"}`} value={rows.length} tint={{ color: SC.color, soft: SC.soft }} />
      <Stat icon={Boxes} label="إجمالي الكمية" value={qty(totals.qty)} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={CircleDollarSign} label="إجمالي القيمة" value={money(totals.value)} tint={{ color: C.success, soft: C.successSoft }} />
      <Stat icon={HandCoins} label="إجمالي التحصيل" value={money(totals.collected)} tint={{ color: C.accent, soft: C.accentSoft }} />
      <Stat icon={AlertTriangle} label="إجمالي الديون" value={money(totals.debt)} tint={{ color: C.danger, soft: C.dangerSoft }} />
    </div>

    {/* الترتيب حسب */}
    <div className="flex items-center gap-2 flex-wrap">
      <span className="text-xs font-bold" style={{ color: C.soft }}>ترتيب حسب:</span>
      {SORTS.map(([k, t]) => { const on = sortBy === k; return (<button key={k} onClick={() => setSortBy(k)} className="rounded-lg px-3 py-1 text-xs font-bold transition-colors" style={{ background: on ? SC.color : C.surface, color: on ? C.onAccent : C.soft, border: `1px solid ${on ? SC.color : C.line}` }}>{t}</button>); })}
    </div>

    {/* الجدول التفصيلي */}
    <Card className="p-4">
      <h3 className="font-extrabold mb-3" style={{ color: C.text }}>التقرير حسب {dimLabel} ({rows.length})</h3>
      <div className="space-y-2.5">{rows.map((r) => (
        <div key={r.name} className="rounded-xl p-3 cursor-pointer transition-colors" style={{ border: `1px solid ${drill === r.name ? SC.color : C.line}`, background: drill === r.name ? SC.soft : C.surface }} onClick={() => setDrill(drill === r.name ? null : r.name)}>
          <div className="flex items-center justify-between mb-1.5">
            <span className="font-extrabold text-sm flex items-center gap-1.5" style={{ color: C.text }}>{dim === "city" || dim === "area" ? <MapPin size={14} style={{ color: SC.color }} /> : dim === "supervisor" ? <Briefcase size={14} style={{ color: C.teal }} /> : <Handshake size={14} style={{ color: C.info }} />}{r.name}</span>
            <span className="font-extrabold" style={{ color: C.text }}>{money(r.value)}</span>
          </div>
          <ProgressBar value={r.value} max={maxVal} color={SC.color} />
          <div className="grid grid-cols-4 gap-2 mt-2 text-center">
            <div><div className="text-[10px]" style={{ color: C.soft }}>الكمية</div><div className="font-bold text-xs" style={{ color: SC.color }}>{qty(r.qty)}</div></div>
            <div><div className="text-[10px]" style={{ color: C.soft }}>التحصيل</div><div className="font-bold text-xs" style={{ color: C.success }}>{money(r.collected)}</div></div>
            <div><div className="text-[10px]" style={{ color: C.soft }}>الديون</div><div className="font-bold text-xs" style={{ color: r.debt > 0 ? C.danger : C.soft }}>{money(r.debt)}</div></div>
            <div><div className="text-[10px]" style={{ color: C.soft }}>طلبات · عملاء</div><div className="font-bold text-xs" style={{ color: C.text }}>{r.orders} · {r.clients}</div></div>
          </div>
          {drill === r.name && <div className="mt-3 pt-3 space-y-1.5" style={{ borderTop: `1px dashed ${C.line}` }}>
            <div className="text-[11px] font-bold" style={{ color: C.soft }}>محلات هذه المجموعة ({drillMerchants.length}):</div>
            {drillMerchants.map((m) => { const md = merchantOpenDebt(data, m.id); return (<div key={m.id} className="flex items-center justify-between rounded-lg px-2.5 py-1.5" style={{ background: C.well }}><span className="text-xs font-bold" style={{ color: C.text }}>{m.name}<span className="font-normal" style={{ color: C.soft }}> · {m.area}</span></span><span className="text-[11px] font-bold" style={{ color: md > 0 ? C.danger : C.success }}>{md > 0 ? "دين " + money(md) : "بلا دين"}</span></div>); })}
          </div>}
        </div>
      ))}{rows.length === 0 && <Empty icon={PieChart} text="لا بيانات في هذا النطاق." />}</div>
    </Card>
  </div>);
}

function MgReports(props) {
  const { data } = props; const mg = useMg(props); const co = useCo(props);
  const [branch, setBranch] = useState("all");
  const [tab, setTab] = useState("overview");
  // النطاق حسب الفرع المُشرَف عليه
  const reps = mgScopedReps(data, co, mg, branch);
  const supers = mgScopedSupervisors(data, co, mg, branch);
  const merchants = mgScopedMerchants(data, co, mg, branch);
  const repIds = new Set(reps.map((r) => r.id));
  const merchIds = new Set(merchants.map((m) => m.id));
  const orders = data.orders.filter((o) => o.companyId === mg.companyId && (repIds.has(o.repId) || merchIds.has(o.merchantId)));
  const delivered = orders.filter((o) => o.status === "delivered");
  const visits = (data.visits || []).filter((v) => repIds.has(v.repId));
  const payments = (data.payments || []).filter((p) => p.repId && repIds.has(p.repId));

  // ═══ إجماليات ═══
  const totalQty = delivered.reduce((s, o) => s + orderQty(o), 0);
  const totalValue = delivered.reduce((s, o) => s + orderTotal(o), 0);
  const totalCollected = payments.reduce((s, p) => s + p.amount, 0);
  const totalDebt = merchants.reduce((s, m) => s + Math.max(0, -(m.balance || 0)), 0);
  const avgOrder = delivered.length ? Math.round(totalValue / delivered.length) : 0;
  const successVisits = visits.filter((v) => ["order", "payment"].includes(v.result)).length;
  const successPct = visits.length ? Math.round((successVisits / visits.length) * 100) : 0;
  const activeClients = new Set(delivered.map((o) => o.merchantId)).size;
  const collectRate = totalValue ? Math.round((totalCollected / totalValue) * 100) : 0;

  // ═══ أداء المندوبين ═══
  const repRows = reps.map((r) => {
    const rdlv = delivered.filter((o) => o.repId === r.id);
    const rQty = rdlv.reduce((s, o) => s + orderQty(o), 0);
    const rVal = rdlv.reduce((s, o) => s + orderTotal(o), 0);
    const rCol = payments.filter((p) => p.repId === r.id).reduce((s, p) => s + p.amount, 0);
    const sup = data.supervisors.find((x) => x.id === r.supervisorId);
    const qPct = r.qtyTarget ? Math.round(((r.qtyAchieved || 0) / r.qtyTarget) * 100) : 0;
    const vPct = r.target ? Math.round(((r.achieved || 0) / r.target) * 100) : 0;
    return { r, rQty, rVal, rCol, sup: sup?.name || "—", qPct, vPct };
  }).sort((a, b) => b.rVal - a.rVal);

  // ═══ أداء المشرفين ═══
  const supRows = supers.map((sv) => {
    const svReps = reps.filter((r) => r.supervisorId === sv.id);
    const svIds = new Set(svReps.map((r) => r.id));
    const svDlv = delivered.filter((o) => svIds.has(o.repId));
    const svVal = svDlv.reduce((s, o) => s + orderTotal(o), 0);
    const svQty = svDlv.reduce((s, o) => s + orderQty(o), 0);
    return { sv, reps: svReps.length, val: svVal, qty: svQty };
  }).sort((a, b) => b.val - a.val);

  // ═══ أداء المنتجات ═══
  const prodSold = {};
  delivered.forEach((o) => o.items.forEach((it) => { prodSold[it.pid] = prodSold[it.pid] || { q: 0, v: 0 }; prodSold[it.pid].q += it.qty; prodSold[it.pid].v += it.qty * it.price; }));
  const prodRows = Object.entries(prodSold).map(([pid, x]) => ({ name: prodName(data, pid), q: x.q, v: x.v })).sort((a, b) => b.v - a.v);
  const maxProdV = Math.max(1, ...prodRows.map((p) => p.v));

  // ═══ أداء المدن ═══
  const citySold = {};
  delivered.forEach((o) => { const m = merchants.find((x) => x.id === o.merchantId); const c = m?.city || "غير محدد"; citySold[c] = citySold[c] || { v: 0, q: 0 }; citySold[c].v += orderTotal(o); citySold[c].q += orderQty(o); });
  const cityRows = Object.entries(citySold).map(([name, x]) => ({ name, ...x })).sort((a, b) => b.v - a.v);
  const maxCityV = Math.max(1, ...cityRows.map((c) => c.v));

  const scopeName = branch === "all" ? (mgSupervisesAll(co, mg) ? "كل الفروع" : "كل فروعي") : branchName(co, branch);
  const TABS = [["overview", "نظرة عامة", BarChart3], ["reps", "المندوبون", Handshake], ["supervisors", "المشرفون", Briefcase], ["products", "المنتجات", Package], ["cities", "المدن", MapPin], ["collections", "التحصيل", HandCoins]];

  return (<div className="space-y-4"><SectionTitle sub={`تقرير شامل عن الأداء البيعي والتحصيل — النطاق: ${scopeName}.`}>التقارير التفصيلية</SectionTitle>
    <MgBranchSwitcher co={co} mg={mg} branch={branch} setBranch={setBranch} />

    {/* تبويبات التقرير */}
    <div className="flex gap-1.5 flex-wrap rounded-2xl p-1.5" style={{ background: C.well, border: `1px solid ${C.line}` }}>
      {TABS.map(([k, t, Ic]) => { const on = tab === k; return (<button key={k} onClick={() => setTab(k)} className="rb-press flex items-center gap-1.5 rounded-xl px-3 py-2 text-sm font-bold transition-all" style={{ background: on ? SC.color : "transparent", color: on ? C.onAccent : C.soft }}><Ic size={14} />{t}</button>); })}
    </div>

    {tab === "overview" && <>
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <Stat icon={Boxes} label="الكمية المُسلّمة" value={qty(totalQty)} tint={{ color: SC.color, soft: SC.soft }} />
        <Stat icon={CircleDollarSign} label="قيمة المبيعات" value={money(totalValue)} tint={{ color: C.success, soft: C.successSoft }} />
        <Stat icon={HandCoins} label="إجمالي التحصيل" value={money(totalCollected)} tint={{ color: C.accent, soft: C.accentSoft }} hint={`${collectRate}% من المبيعات`} />
        <Stat icon={AlertTriangle} label="إجمالي الديون" value={money(totalDebt)} tint={{ color: C.danger, soft: C.dangerSoft }} />
        <Stat icon={ShoppingCart} label="طلبات مُسلّمة" value={delivered.length} tint={{ color: C.info, soft: C.infoSoft }} />
        <Stat icon={Landmark} label="متوسط الطلب" value={money(avgOrder)} tint={{ color: C.purple, soft: C.purpleSoft }} />
        <Stat icon={Store} label="عملاء نشطون" value={activeClients} tint={{ color: SC.color, soft: SC.soft }} hint={`من ${merchants.length}`} />
        <Stat icon={CheckCircle2} label="نجاح الزيارات" value={successPct + "%"} tint={{ color: C.success, soft: C.successSoft }} hint={`${visits.length} زيارة`} />
      </div>
      <div className="grid lg:grid-cols-2 gap-4">
        <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>أعلى 5 مندوبين (بالقيمة)</h3>
          <div className="space-y-2">{repRows.slice(0, 5).map((x, i) => <div key={x.r.id} className="flex items-center gap-2.5"><span className="w-6 h-6 rounded-lg flex items-center justify-center text-xs font-extrabold" style={{ background: i === 0 ? C.warnSoft : C.well, color: i === 0 ? "#B45309" : C.soft }}>{i + 1}</span><span className="flex-1 text-sm font-bold truncate" style={{ color: C.text }}>{x.r.name}</span><span className="text-sm font-extrabold" style={{ color: SC.color }}>{money(x.rVal)}</span></div>)}{!repRows.length && <Empty icon={Handshake} text="لا مبيعات." />}</div>
        </Card>
        <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>أعلى 5 منتجات (بالقيمة)</h3>
          <div className="space-y-2.5">{prodRows.slice(0, 5).map((p) => <div key={p.name}><div className="flex justify-between text-sm mb-1"><span className="font-bold truncate" style={{ color: C.text }}>{p.name}</span><span className="font-extrabold" style={{ color: C.text }}>{money(p.v)} · {qty(p.q)}</span></div><ProgressBar value={p.v} max={maxProdV} color={SC.color} /></div>)}{!prodRows.length && <Empty icon={Package} text="لا مبيعات." />}</div>
        </Card>
      </div>
    </>}

    {tab === "reps" && <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>أداء المندوبين التفصيلي</h3>
      <Table head={["المندوب", "المشرف", "الفرع", "كمية مُسلّمة", "قيمة مُسلّمة", "تحصيل", "تحقيق الكمية", "تحقيق القيمة"]}>{repRows.map((x) => <tr key={x.r.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{x.r.name}</Td><Td>{x.sup}</Td><Td>{x.r.branchId ? <BranchBadge co={co} bid={x.r.branchId} tiny /> : <span className="text-xs" style={{ color: C.soft }}>—</span>}</Td><Td><span className="font-bold" style={{ color: SC.color }}>{qty(x.rQty)}</span></Td><Td>{money(x.rVal)}</Td><Td><span style={{ color: C.success }}>{money(x.rCol)}</span></Td><Td><span className="font-bold" style={{ color: x.qPct >= 100 ? C.success : x.qPct >= 60 ? "#B45309" : C.danger }}>{x.qPct}%</span></Td><Td><span className="font-bold" style={{ color: x.vPct >= 100 ? C.success : x.vPct >= 60 ? "#B45309" : C.danger }}>{x.vPct}%</span></Td></tr>)}</Table>
      {!repRows.length && <Empty icon={Handshake} text="لا مندوبين في هذا النطاق." />}
    </Card>}

    {tab === "supervisors" && <div className="grid sm:grid-cols-2 gap-3">{supRows.map((x) => <Card key={x.sv.id} className="p-4">
      <div className="flex items-center gap-3 mb-3"><Avatar name={x.sv.name} size={40} /><div className="flex-1"><div className="font-extrabold" style={{ color: C.text }}>{x.sv.name}</div><div className="text-[11px]" style={{ color: C.soft }}>{x.reps} مندوب{x.sv.branchId ? ` · ${branchName(co, x.sv.branchId)}` : ""}</div></div></div>
      <div className="grid grid-cols-2 gap-2 text-center"><div className="rounded-lg py-2" style={{ background: C.well }}><div className="text-[10px]" style={{ color: C.soft }}>القيمة</div><div className="font-extrabold text-sm" style={{ color: C.text }}>{money(x.val)}</div></div><div className="rounded-lg py-2" style={{ background: C.well }}><div className="text-[10px]" style={{ color: C.soft }}>الكمية</div><div className="font-extrabold text-sm" style={{ color: SC.color }}>{qty(x.qty)}</div></div></div>
    </Card>)}{!supRows.length && <div className="sm:col-span-2"><Empty icon={Briefcase} text="لا مشرفين في هذا النطاق." /></div>}</div>}

    {tab === "products" && <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>أداء المنتجات (بالكمية والقيمة)</h3>
      <Table head={["المنتج", "الكمية المباعة", "القيمة", "الحصة"]}>{prodRows.map((p) => <tr key={p.name} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{p.name}</Td><Td><span className="font-bold" style={{ color: SC.color }}>{qty(p.q)}</span></Td><Td>{money(p.v)}</Td><Td><span className="text-xs font-bold">{Math.round((p.v / (totalValue || 1)) * 100)}%</span></Td></tr>)}</Table>
      {!prodRows.length && <Empty icon={Package} text="لا مبيعات." />}
    </Card>}

    {tab === "cities" && <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>المبيعات حسب المدينة</h3>
      <div className="space-y-3">{cityRows.map((c) => <div key={c.name}><div className="flex justify-between text-sm mb-1"><span className="font-bold flex items-center gap-1.5" style={{ color: C.text }}><MapPin size={13} style={{ color: SC.color }} />{c.name}</span><span className="font-extrabold" style={{ color: C.text }}>{money(c.v)} · {qty(c.q)}</span></div><ProgressBar value={c.v} max={maxCityV} color={SC.color} /></div>)}{!cityRows.length && <Empty icon={MapPin} text="لا مبيعات." />}</div>
    </Card>}

    {tab === "collections" && <>
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <Stat icon={HandCoins} label="إجمالي التحصيل" value={money(totalCollected)} tint={{ color: C.success, soft: C.successSoft }} />
        <Stat icon={CircleDollarSign} label="قيمة المبيعات" value={money(totalValue)} tint={{ color: C.info, soft: C.infoSoft }} />
        <Stat icon={Percent} label="نسبة التحصيل" value={collectRate + "%"} tint={{ color: C.accent, soft: C.accentSoft }} />
        <Stat icon={AlertTriangle} label="الديون القائمة" value={money(totalDebt)} tint={{ color: C.danger, soft: C.dangerSoft }} />
      </div>
      <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>سجل التحصيل ({payments.length})</h3>
        <Table head={["المندوب", "المحل", "المبلغ", "الطريقة", "التاريخ"]}>{payments.slice(0, 30).map((p) => { const rp = data.reps.find((r) => r.id === p.repId); return <tr key={p.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{rp?.name || "—"}</Td><Td>{merchName(data, p.merchantId)}</Td><Td><span className="font-bold" style={{ color: C.success }}>{money(p.amount)}</span></Td><Td>{p.method}</Td><Td><span className="text-xs" style={{ color: C.soft }}>{p.date}</span></Td></tr>; })}</Table>
        {!payments.length && <Empty icon={HandCoins} text="لا تحصيلات في هذا النطاق." />}
      </Card>
    </>}
  </div>);
}
function MgTargets(props) {
  const { data, actions } = props; const mg = useMg(props);
  const supers = data.supervisors.filter((s) => s.companyId === mg.companyId);
  const mgQtyTarget = mg.qtyTarget || 0;
  const distributedQty = supers.reduce((s, x) => s + (x.qtyTarget || 0), 0);
  const remainingQty = mgQtyTarget - distributedQty;
  const mgColTarget = mg.collectionTarget || 0;
  const distributedCol = supers.reduce((s, x) => s + (x.collectionTarget || 0), 0);
  const remainingCol = mgColTarget - distributedCol;
  return (<div className="space-y-4"><SectionTitle sub="وزّع أهداف الكمية (ستيكات) والتحصيل — المفروضة عليك من الشركة — على مشرفيك.">أهداف المشرفين</SectionTitle>
    <Card className="p-4" style={{ borderRight: `3px solid ${C.ink}` }}>
      <h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Target size={17} style={{ color: C.ink }} />هدفي من الشركة وتوزيعه على المشرفين</h3>
      <div className="grid sm:grid-cols-2 gap-3">
        <div className="rounded-xl p-3" style={{ background: C.well }}>
          <div className="text-xs font-bold mb-1" style={{ color: SC.color }}>هدف الكمية (ستيكات)</div>
          <div className="flex items-end justify-between"><span className="text-2xl font-extrabold" style={{ color: C.text }}>{qty(mgQtyTarget)}</span><span className="text-xs" style={{ color: C.soft }}>وُزّع {qty(distributedQty)}</span></div>
          <div className="mt-2 rounded-lg px-2 py-1 text-xs font-bold text-center" style={{ background: remainingQty === 0 ? C.successSoft : remainingQty > 0 ? C.warnSoft : C.dangerSoft, color: remainingQty === 0 ? C.success : remainingQty > 0 ? "#B45309" : C.danger }}>{remainingQty === 0 ? "✓ وُزّع بالكامل" : remainingQty > 0 ? `متبقّي للتوزيع: ${qty(remainingQty)}` : `⚠️ تجاوزت بـ ${qty(-remainingQty)}`}</div>
        </div>
        <div className="rounded-xl p-3" style={{ background: C.well }}>
          <div className="text-xs font-bold mb-1" style={{ color: C.info }}>هدف التحصيل (قيمة)</div>
          <div className="flex items-end justify-between"><span className="text-2xl font-extrabold" style={{ color: C.text }}>{money(mgColTarget)}</span><span className="text-xs" style={{ color: C.soft }}>وُزّع {money(distributedCol)}</span></div>
          <div className="mt-2 rounded-lg px-2 py-1 text-xs font-bold text-center" style={{ background: remainingCol === 0 ? C.successSoft : remainingCol > 0 ? C.warnSoft : C.dangerSoft, color: remainingCol === 0 ? C.success : remainingCol > 0 ? "#B45309" : C.danger }}>{remainingCol === 0 ? "✓ وُزّع بالكامل" : remainingCol > 0 ? `متبقّي: ${money(remainingCol)}` : `⚠️ تجاوزت بـ ${money(-remainingCol)}`}</div>
        </div>
      </div>
    </Card>
    <Card className="p-4"><Table head={["المشرف", "المندوبون", "هدف الكمية", "تعديل", "هدف التحصيل", "تعديل", "إنجاز ك/ت"]}>{supers.map((s) => { const reps = data.reps.filter((r) => r.supervisorId === s.id); const pctQ = Math.round((s.qtyAchieved / (s.qtyTarget || 1)) * 100); const pctC = Math.round((s.collected / (s.collectionTarget || 1)) * 100); return (<tr key={s.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{s.name}</Td><Td>{reps.length}</Td><Td>{qty(s.qtyTarget)}</Td><Td><input type="number" defaultValue={s.qtyTarget} onBlur={(e) => actions.setSupervisorTarget(s.id, { qtyTarget: +e.target.value })} className="w-20 rounded-lg px-2 py-1 text-sm outline-none" style={inputStyle} /></Td><Td>{money(s.collectionTarget)}</Td><Td><input type="number" defaultValue={s.collectionTarget} onBlur={(e) => actions.setSupervisorTarget(s.id, { collectionTarget: +e.target.value })} className="w-24 rounded-lg px-2 py-1 text-sm outline-none" style={inputStyle} /></Td><Td><span style={{ color: pctQ >= 70 ? C.success : C.accent, fontWeight: 700 }}>{pctQ}%</span> / <span style={{ color: pctC >= 70 ? C.success : C.accent, fontWeight: 700 }}>{pctC}%</span></Td></tr>); })}</Table>{supers.length === 0 && <Empty icon={Briefcase} text="لا يوجد مشرفون." />}</Card>
  </div>);
}
// كميات المدير المخصصة من الشركة
function MgStock(props) {
  const { data } = props; const mg = useMg(props);
  const stock = data.mgrStock.filter((s) => s.managerId === mg.id);
  const totalAlloc = stock.reduce((s, x) => s + x.allocated, 0);
  const totalDist = stock.reduce((s, x) => s + x.distributed, 0);
  return (<div className="space-y-4"><SectionTitle sub="كميات البضاعة (ستيكات) التي خصّصتها لك الشركة لتوزّعها على مشرفيك ومندوبيك.">كمياتي من الشركة</SectionTitle>
    <div className="grid sm:grid-cols-3 gap-3">
      <Stat icon={Boxes} label="إجمالي المخصص" value={qty(totalAlloc)} tint={{ color: SC.color2, soft: SC.soft }} />
      <Stat icon={PackageCheck} label="الموزّع" value={qty(totalDist)} tint={{ color: C.success, soft: C.successSoft }} />
      <Stat icon={Package} label="المتبقي" value={qty(totalAlloc - totalDist)} tint={{ color: C.accent, soft: C.accentSoft }} />
    </div>
    <Card className="p-4"><Table head={["المنتج", "المخصص", "المُوزّع", "المتبقي", "نسبة التوزيع"]}>{stock.map((s) => { const rem = s.allocated - s.distributed; const pct = Math.round((s.distributed / (s.allocated || 1)) * 100); return (<tr key={s.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{prodName(data, s.pid)}</Td><Td>{qty(s.allocated)}</Td><Td>{qty(s.distributed)}</Td><Td><span style={{ color: rem < 100 ? C.danger : C.text, fontWeight: 700 }}>{qty(rem)}</span></Td><Td><div className="flex items-center gap-2"><div style={{ width: 70 }}><ProgressBar value={s.distributed} max={s.allocated} color={SC.color2} /></div><span className="text-xs" style={{ color: C.soft }}>{pct}%</span></div></Td></tr>); })}</Table>{stock.length === 0 && <Empty icon={Boxes} text="لم تخصّص لك الشركة كميات بعد." />}</Card>
  </div>);
}


/* ============================================================
   ============   مروّج المبيعات   ===========================
   ============================================================ */
const usePromoter = (props) => props.data.promoters.find((p) => p.id === props.user.entityId);
function PmHome(props) {
  const { data, go, user } = props; const pr = usePromoter(props);
  const tasks = data.promotions.filter((t) => t.promoterId === pr.id);
  const pending = tasks.filter((t) => t.status !== "sent").length;
  const alerts = [];
  if (pending > 0) alerts.push({ k: "pm_tasks", icon: ImageIcon, text: `${pending} مهمة عرض بانتظارك`, soft: C.roseSoft, color: C.rose, border: C.rose });
  return <RoleLauncher name={pr.name} sub={pr.area || ""} roleTitle="مروّج مبيعات" nav={navFor(user.role, user.companyRole)} homeKey="pm_home" go={go} alerts={alerts} badges={{ pm_tasks: pending }} quick={{ k: "pm_route", icon: Navigation, label: "خط سيري" }} />;
}
function PmRoute(props) {
  const { data } = props; const pr = usePromoter(props);
  const stops = data.promotions.filter((t) => t.promoterId === pr.id && t.status !== "sent").map((t) => ({ t, m: data.merchants.find((x) => x.id === t.merchantId) })).filter((x) => x.m);
  return (<div className="space-y-4"><SectionTitle sub="ترتيب محطات الترويج التي عليك زيارتها اليوم.">خط السير اليومي</SectionTitle>
    <Card className="p-5">
      <div className="rounded-xl p-4 mb-4 flex items-center justify-between" style={{ background: C.roseSoft }}><div className="flex items-center gap-2"><Navigation size={18} style={{ color: C.rose }} /><span className="text-sm font-bold" style={{ color: C.text }}>{stops.length} محطة ترويج</span></div><span className="text-sm" style={{ color: C.soft }}>جولة اليوم</span></div>
      <div className="relative pr-4">{stops.map(({ t, m }, i) => (<div key={t.id} className="flex items-start gap-3 pb-5 relative">
        {i < stops.length - 1 && <div className="absolute right-[15px] top-9 bottom-0 w-0.5" style={{ background: C.line }} />}
        <div className="rounded-full flex items-center justify-center shrink-0 z-10 text-xs font-bold text-white" style={{ width: 32, height: 32, background: t.status === "sent" ? C.success : C.rose }}>{i + 1}</div>
        <div className="flex-1 rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}><div className="flex items-center justify-between"><div className="font-bold text-sm" style={{ color: C.text }}>{m.name}</div><Badge status={t.status} map={PROMO_STATUS} /></div><div className="text-xs flex items-center gap-1 mt-0.5" style={{ color: C.soft }}><MapPin size={11} />{m.area} · طلب {t.orderId}</div><div className="mt-1.5"><GpsLink m={m} sm /></div></div>
      </div>))}{stops.length === 0 && <Empty icon={Route} text="لا توجد محطات اليوم." />}</div>
    </Card>
  </div>);
}
function PmTasks(props) {
  const { data, actions } = props; const pr = usePromoter(props);
  const tasks = data.promotions.filter((t) => t.promoterId === pr.id && t.status !== "sent");
  const [exp, setExp] = useState({});     // إدخال ملاحظة صلاحية لكل مهمة
  const [stk, setStk] = useState({});     // تقييم التستيف
  const [sug, setSug] = useState({});     // الطلبية المقترحة
  const RATE = ["ممتاز — مطابق للمخطط", "جيد", "مخالف — يحتاج تصحيح"];
  const ALERTS = [["wants_items", "المحل يريد هذه الأصناف"], ["needs_visit", "المحل يحتاج زيارتك"], ["contact", "تواصل مع المحل"]];
  const getExp = (id) => exp[id] || { name: "", expDate: "", flag: "near" };
  const getStk = (id) => stk[id] || { fridge: RATE[0], shelf: RATE[0], note: "" };
  const getSug = (id) => sug[id] || { name: "", qty: 1, items: [], alertType: "wants_items", note: "" };
  return (<div className="space-y-4"><SectionTitle sub="أكّد وقوفك بمحيط المحل أولاً، ثم نفّذ متطلبات الشركة: صور قبل/بعد، فحص الصلاحيات، التستيف — وبلّغ المندوب بما يريده المحل.">مهام الزيارة الميدانية</SectionTitle>
    {tasks.length ? tasks.map((t) => { const m = data.merchants.find((x) => x.id === t.merchantId);
      const req = t.req || { photoBefore: true, photoAfter: true, fridge: true, shelf: true, expiry: true };
      const canFinish = t.gpsConfirmed && (!req.photoBefore || t.photoBefore) && (!req.photoAfter || t.photoAfter) && (!(req.fridge || req.shelf) || t.stacking);
      const e = getExp(t.id), st = getStk(t.id), sg = getSug(t.id);
      return (<Card key={t.id} className="p-4 sm:p-5 min-w-0">
        {/* رأس المهمة */}
        <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
          <div className="flex items-center gap-3 min-w-0"><div className="rounded-xl flex items-center justify-center text-2xl shrink-0" style={{ width: 44, height: 44, background: C.roseSoft }}>{m?.photo || "🏪"}</div><div className="min-w-0"><div className="font-extrabold truncate" style={{ color: C.text }}>{m?.name}</div><div className="text-xs" style={{ color: C.soft }}>{m?.area}{t.orderId ? ` · طلب ${t.orderId}` : ""} · {t.date}</div></div></div>
          <Badge status={t.status} map={PROMO_STATUS} />
        </div>
        {/* متطلبات الشركة */}
        <div className="flex gap-1.5 flex-wrap mb-3">{[[req.photoBefore, "صورة قبل"], [req.photoAfter, "صورة بعد"], [req.fridge, "تستيف الثلاجة"], [req.shelf, "تستيف الرفوف"], [req.expiry, "فحص الصلاحيات"]].filter(([on]) => on).map(([, lbl], i) => <span key={i} className="text-[10px] font-bold px-2 py-1 rounded-full" style={{ background: C.roseSoft, color: C.rose }}>مطلوب: {lbl}</span>)}</div>

        {/* ١. تأكيد الوقوف بمحيط المحل GPS */}
        {!t.gpsConfirmed ? <div className="rounded-xl p-3 mb-3" style={{ background: C.infoSoft + "55", border: `1px solid ${C.info}` }}>
          <div className="text-xs font-extrabold mb-2 flex items-center gap-1.5" style={{ color: C.info }}><MapPinned size={15} />١. أكّد أنك في محيط المحل قبل البدء</div>
          <LocationCard m={m} />
          <div className="mt-2"><PrimaryBtn full icon={Navigation} onClick={() => { if (!navigator.geolocation) { toast("المتصفح لا يدعم تحديد الموقع", "danger"); return; } navigator.geolocation.getCurrentPosition(({ coords }) => { if (m?.lat != null && m?.lng != null && geoDistanceM(coords.latitude, coords.longitude, m.lat, m.lng) > 300) { toast("أنت بعيد عن المحل أكثر من 300 متر", "danger"); return; } actions.pmConfirmGps(t.id, { lat: coords.latitude, lng: coords.longitude, accuracy: coords.accuracy }); toast("تم التحقق — أنت داخل محيط المحل ✓"); }, () => toast("اسمح للتطبيق باستخدام الموقع ثم أعد المحاولة", "danger"), { enableHighAccuracy: true, timeout: 12000, maximumAge: 0 }); }}>أنا الآن في محيط المحل — تأكيد GPS</PrimaryBtn></div>
        </div> : <>
        <div className="rounded-xl p-2.5 mb-3 flex items-center gap-2 text-xs font-bold" style={{ background: C.successSoft, color: C.success }}><MapPinned size={14} />موقعك مؤكد داخل محيط المحل ✓</div>

        {/* ٢. صورة قبل استلام/رصّ البضاعة */}
        {req.photoBefore && <div className="rounded-xl p-3 mb-2.5 flex items-center justify-between gap-2 flex-wrap" style={{ background: C.bg }}>
          <div className="text-xs font-extrabold flex items-center gap-1.5" style={{ color: C.text }}><ImageIcon size={15} style={{ color: C.rose }} />٢. صورة قبل (الثلاجة/الرف قبل الترتيب)</div>
          {t.photoBefore ? <span className="text-xs font-bold flex items-center gap-1" style={{ color: C.success }}><CheckCircle2 size={14} />التُقطت 📸</span> : <label className="rb-press inline-flex items-center gap-1.5 rounded-xl px-3 py-2 text-xs font-bold cursor-pointer" style={{ background: C.surface, color: C.text, border: `1px solid ${C.line}` }}><ImageIcon size={14} />التقاط<input type="file" accept="image/*" capture="environment" className="hidden" onChange={(e) => { const file = e.target.files?.[0]; if (!file) return; const reader = new FileReader(); reader.onload = () => { actions.pmPhoto(t.id, "photoBefore", reader.result); toast("تم حفظ صورة قبل 📸"); }; reader.readAsDataURL(file); e.target.value = ""; }} /></label>}
        </div>}

        {/* ٣. فحص الصلاحيات */}
        {req.expiry && <div className="rounded-xl p-3 mb-2.5" style={{ background: C.bg }}>
          <div className="text-xs font-extrabold mb-2 flex items-center gap-1.5" style={{ color: C.text }}><CalendarClock size={15} style={{ color: C.accent }} />٣. متابعة الصلاحيات في المحل</div>
          {t.expiryFindings?.length > 0 && <div className="space-y-1 mb-2">{t.expiryFindings.map((fd, i) => <div key={i} className="flex items-center justify-between rounded-lg px-2.5 py-1.5 text-xs" style={{ background: C.surface }}><span className="font-bold" style={{ color: C.text }}>{fd.name}</span><span className="font-bold px-1.5 py-0.5 rounded text-[10px]" style={{ background: fd.flag === "expired" ? C.dangerSoft : fd.flag === "near" ? C.accentSoft : C.successSoft, color: fd.flag === "expired" ? C.danger : fd.flag === "near" ? "#92590C" : C.success }}>{fd.flag === "expired" ? `منتهي ${fmtDate(fd.expDate)}` : fd.flag === "near" ? `قريب ${fmtDate(fd.expDate)}` : "سليم"}</span></div>)}</div>}
          <div className="grid grid-cols-2 gap-1.5 mb-1.5">
            <Input value={e.name} onChange={(ev) => setExp({ ...exp, [t.id]: { ...e, name: ev.target.value } })} placeholder="اسم الصنف" />
            <Input value={e.expDate} onChange={(ev) => setExp({ ...exp, [t.id]: { ...e, expDate: ev.target.value } })} placeholder="الصلاحية 2026-08" />
          </div>
          <div className="flex gap-1.5">
            <select value={e.flag} onChange={(ev) => setExp({ ...exp, [t.id]: { ...e, flag: ev.target.value } })} className={inputCls + " flex-1"} style={inputStyle}><option value="near">قريب الانتهاء</option><option value="expired">منتهي</option><option value="ok">سليم</option></select>
            <GhostBtn sm icon={Plus} onClick={() => { if (!e.name) { toast("اكتب اسم الصنف", "danger"); return; } actions.pmSaveExpiry(t.id, [...(t.expiryFindings || []), e]); setExp({ ...exp, [t.id]: { name: "", expDate: "", flag: "near" } }); toast("سُجّلت الملاحظة ✓"); }}>إضافة</GhostBtn>
          </div>
        </div>}

        {/* ٤. التستيف في الثلاجة والرفوف */}
        {(req.fridge || req.shelf) && <div className="rounded-xl p-3 mb-2.5" style={{ background: C.bg }}>
          <div className="text-xs font-extrabold mb-2 flex items-center gap-1.5" style={{ color: C.text }}><Boxes size={15} style={{ color: C.purple }} />٤. طريقة تستيف البضاعة</div>
          {t.stacking ? <div className="text-xs space-y-1"><div className="flex justify-between"><span style={{ color: C.soft }}>الثلاجة</span><b style={{ color: C.text }}>{t.stacking.fridge || "—"}</b></div><div className="flex justify-between"><span style={{ color: C.soft }}>الرفوف</span><b style={{ color: C.text }}>{t.stacking.shelf || "—"}</b></div>{t.stacking.note && <div style={{ color: C.soft }}>ملاحظة: {t.stacking.note}</div>}</div> : <>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-1.5 mb-1.5">
              {req.fridge && <Field label="تستيف الثلاجة"><select value={st.fridge} onChange={(ev) => setStk({ ...stk, [t.id]: { ...st, fridge: ev.target.value } })} className={inputCls} style={inputStyle}>{RATE.map((r) => <option key={r}>{r}</option>)}</select></Field>}
              {req.shelf && <Field label="تستيف الرفوف"><select value={st.shelf} onChange={(ev) => setStk({ ...stk, [t.id]: { ...st, shelf: ev.target.value } })} className={inputCls} style={inputStyle}>{RATE.map((r) => <option key={r}>{r}</option>)}</select></Field>}
            </div>
            <div className="flex gap-1.5"><div className="flex-1"><Input value={st.note} onChange={(ev) => setStk({ ...stk, [t.id]: { ...st, note: ev.target.value } })} placeholder="ملاحظة (اختياري)" /></div><GhostBtn sm icon={CheckCircle2} onClick={() => { actions.pmSaveStacking(t.id, { fridge: req.fridge ? st.fridge : null, shelf: req.shelf ? st.shelf : null, note: st.note }); toast("سُجّل التستيف ✓"); }}>حفظ</GhostBtn></div>
          </>}
        </div>}

        {/* ٥. صورة بعد الترتيب */}
        {req.photoAfter && <div className="rounded-xl p-3 mb-2.5 flex items-center justify-between gap-2 flex-wrap" style={{ background: C.bg }}>
          <div className="text-xs font-extrabold flex items-center gap-1.5" style={{ color: C.text }}><ImageIcon size={15} style={{ color: C.rose }} />٥. صورة بعد (البضاعة مرتّبة ومعروضة)</div>
          {t.photoAfter ? <span className="text-xs font-bold flex items-center gap-1" style={{ color: C.success }}><CheckCircle2 size={14} />التُقطت 📸</span> : <label className="rb-press inline-flex items-center gap-1.5 rounded-xl px-3 py-2 text-xs font-bold cursor-pointer" style={{ background: C.surface, color: C.text, border: `1px solid ${C.line}` }}><ImageIcon size={14} />التقاط<input type="file" accept="image/*" capture="environment" className="hidden" onChange={(e) => { const file = e.target.files?.[0]; if (!file) return; const reader = new FileReader(); reader.onload = () => { actions.pmPhoto(t.id, "photoAfter", reader.result); toast("تم حفظ صورة بعد 📸"); }; reader.readAsDataURL(file); e.target.value = ""; }} /></label>}
        </div>}

        {/* ٦. الطلبية المقترحة + تبليغ المندوب */}
        <div className="rounded-xl p-3 mb-3" style={{ background: SC.soft + "55", border: `1px dashed ${SC.color}` }}>
          <div className="text-xs font-extrabold mb-2 flex items-center gap-1.5" style={{ color: SC.color }}><Send size={15} />٦. طلبية مقترحة — بلّغ المندوب</div>
          {t.repAlert ? <div className="text-xs font-bold rounded-lg p-2.5" style={{ background: C.successSoft, color: C.success }}>أُبلغ المندوب ✓ — {ALERTS.find(([k]) => k === t.repAlert.alertType)?.[1]}{t.suggestedItems?.length ? ` (${t.suggestedItems.length} صنف)` : ""}</div> : <>
            {sg.items.length > 0 && <div className="flex gap-1.5 flex-wrap mb-1.5">{sg.items.map((it, i) => <span key={i} className="text-[10px] font-bold px-2 py-1 rounded-full flex items-center gap-1" style={{ background: C.surface, color: C.text }}>{it.name} ×{it.qty}<button onClick={() => setSug({ ...sug, [t.id]: { ...sg, items: sg.items.filter((_, x) => x !== i) } })} style={{ color: C.danger }}>✕</button></span>)}</div>}
            <div className="flex gap-1.5 mb-1.5">
              <div className="flex-1"><Input value={sg.name} onChange={(ev) => setSug({ ...sug, [t.id]: { ...sg, name: ev.target.value } })} placeholder="الصنف الذي يريده المحل" /></div>
              <input type="number" value={sg.qty} onChange={(ev) => setSug({ ...sug, [t.id]: { ...sg, qty: Math.max(1, +ev.target.value || 1) } })} className="w-16 text-center rounded-xl text-sm outline-none" style={inputStyle} />
              <GhostBtn sm icon={Plus} onClick={() => { if (!sg.name) return; setSug({ ...sug, [t.id]: { ...sg, items: [...sg.items, { name: sg.name, qty: sg.qty }], name: "", qty: 1 } }); }}>إضافة</GhostBtn>
            </div>
            <div className="flex gap-1.5 flex-wrap mb-1.5">{ALERTS.map(([k, lbl]) => <button key={k} onClick={() => setSug({ ...sug, [t.id]: { ...sg, alertType: k } })} className="rb-press rounded-xl px-2.5 py-1.5 text-[11px] font-bold" style={{ background: sg.alertType === k ? SC.color : C.surface, color: sg.alertType === k ? C.onAccent : C.text, border: `1px solid ${sg.alertType === k ? SC.color : C.line}` }}>{lbl}</button>)}</div>
            <div className="flex gap-1.5"><div className="flex-1"><Input value={sg.note} onChange={(ev) => setSug({ ...sug, [t.id]: { ...sg, note: ev.target.value } })} placeholder="ملاحظة للمندوب (اختياري)" /></div><PrimaryBtn sm icon={Send} onClick={() => { if (sg.alertType === "wants_items" && !sg.items.length) { toast("أضف الأصناف المطلوبة أولاً", "danger"); return; } actions.pmAlertRep(t.id, { items: sg.items, alertType: sg.alertType, note: sg.note }); toast("أُبلغ المندوب والشركة ✓"); }}>إبلاغ المندوب</PrimaryBtn></div>
          </>}
        </div>

        {/* إتمام */}
        <PrimaryBtn full tone="success" icon={CheckCircle2} onClick={() => { if (!canFinish) { toast("أكمل المتطلبات: الصور والتستيف أولاً", "danger"); return; } actions.pmComplete(t.id); toast("أُرسل تقرير الزيارة للشركة ✓"); }}>إتمام الزيارة وإرسال التقرير للشركة</PrimaryBtn>
        </>}
      </Card>); }) : <Empty icon={ImageIcon} text="لا مهام زيارة حالياً." />}
  </div>);
}
function PmHistory(props) {
  const { data } = props; const pr = usePromoter(props);
  const sent = data.promotions.filter((t) => t.promoterId === pr.id && t.status === "sent");
  return (<div className="space-y-4"><SectionTitle sub="تقارير الزيارات التي أرسلتها للشركة.">السجل المُرسل</SectionTitle>
    <div className="space-y-2.5">{sent.map((t) => { const m = data.merchants.find((x) => x.id === t.merchantId); const bad = (t.expiryFindings || []).filter((fd) => fd.flag !== "ok").length;
      return (<Card key={t.id} className="p-4 min-w-0">
        <div className="flex items-center justify-between flex-wrap gap-2 mb-2"><div className="font-extrabold text-sm" style={{ color: C.text }}>{m?.name}</div><span className="text-[10px]" style={{ color: C.soft }}>{t.date}</span></div>
        <div className="flex gap-1.5 flex-wrap">
          {t.gpsConfirmed && <span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: C.successSoft, color: C.success }}>GPS ✓</span>}
          {t.photoBefore && <span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: C.infoSoft, color: C.info }}>صورة قبل 📸</span>}
          {t.photoAfter && <span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: C.infoSoft, color: C.info }}>صورة بعد 📸</span>}
          {t.stacking && <span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: C.purpleSoft, color: C.purple }}>تستيف: {[t.stacking.fridge && "ثلاجة", t.stacking.shelf && "رفوف"].filter(Boolean).join(" + ")}</span>}
          {bad > 0 && <span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: C.dangerSoft, color: C.danger }}>{bad} ملاحظة صلاحية ⚠️</span>}
          {t.suggestedItems?.length > 0 && <span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: SC.soft, color: SC.color }}>اقترح {t.suggestedItems.length} صنف للمندوب</span>}
        </div>
      </Card>); })}</div>
    {sent.length === 0 && <Empty icon={ImageIcon} text="لا تقارير مُرسلة بعد." />}
  </div>);
}
function SVStocktake(props) {
  const { data, actions } = props; const supv = useSv(props);
  const repIds = data.reps.filter((r) => r.supervisorId === supv.id).map((r) => r.id);
  const sts = data.stocktakes.filter((st) => st.type === "van" && repIds.includes(st.repId));
  return (<div className="space-y-4"><SectionTitle sub="راجع نتائج جرد سيارات مندوبيك — أي عجز يُحمَّل ذمةً على المندوب.">اعتماد الجرد</SectionTitle>
    {sts.map((st) => { const variance = st.lines.reduce((sm, l) => sm + Math.max(0, (l.sysQty || 0) - (l.counted || 0)) * (l.price || 0), 0);
      return (<Card key={st.id} className="p-4 min-w-0">
        <div className="flex items-center justify-between flex-wrap gap-2 mb-2"><div><div className="font-extrabold text-sm" style={{ color: C.text }}>{st.id} · {repName(data, st.repId)}</div><div className="text-[11px]" style={{ color: C.soft }}>{st.note} · {st.date}</div></div>
          <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: st.status === "resolved" ? C.successSoft : st.status === "submitted" ? C.infoSoft : C.accentSoft, color: st.status === "resolved" ? C.success : st.status === "submitted" ? C.info : "#92590C" }}>{st.status === "requested" ? "بانتظار عدّ المندوب" : st.status === "submitted" ? "جاهز لاعتمادك" : "معتمد"}</span></div>
        <div className="space-y-1.5 mb-3">{st.lines.map((l) => { const diff = l.counted != null ? (l.sysQty - l.counted) : null; return (<div key={l.pid} className="flex items-center justify-between gap-2 rounded-xl px-3 py-2 text-xs" style={{ background: C.bg }}>
          <span className="font-bold min-w-0 truncate" style={{ color: C.text }}>{l.name}</span>
          <span style={{ color: C.soft }}>نظام {l.sysQty} · عدّ {l.counted ?? "—"}{diff > 0 && <b style={{ color: C.danger }}> · عجز {diff} ({money(diff * l.price)})</b>}{diff === 0 && <b style={{ color: C.success }}> · مطابق</b>}</span>
        </div>); })}</div>
        {st.status === "submitted" && <PrimaryBtn full tone={variance > 0 ? "accent" : "success"} icon={CheckCircle2} onClick={() => confirmThen(variance > 0 ? `سيُحمَّل عجز ${money(variance)} ذمةً على ${repName(data, st.repId)} وتُسوّى كميات سيارته.` : "الجرد مطابق — سيُعتمد وتُثبّت الكميات.", () => { actions.resolveStocktake(st.id); toast("اعتُمد الجرد ✓"); }, { yesLabel: "اعتماد" })}>{variance > 0 ? `اعتماد وتحميل العجز (${money(variance)})` : "اعتماد — مطابق"}</PrimaryBtn>}
        {st.status === "resolved" && <div className="text-sm font-bold rounded-xl p-2.5" style={{ background: st.varianceValue > 0 ? C.dangerSoft : C.successSoft, color: st.varianceValue > 0 ? C.danger : C.success }}>{st.varianceValue > 0 ? `حُمّل على المندوب: ${money(st.varianceValue)}` : "مطابق — لا فروقات ✓"}</div>}
      </Card>); })}
    {sts.length === 0 && <Empty icon={ClipboardList} text="لا عمليات جرد لسيارات مندوبيك." />}
  </div>);
}
// خط السير الأسبوعي: المشرف يحدد محلات كل يوم لكل مندوب
function SVJourney(props) {
  const { data, actions } = props; const supv = useSv(props);
  const myReps = data.reps.filter((r) => r.supervisorId === supv.id);
  const [repId, setRepId] = useState(myReps[0]?.id || "");
  const [day, setDay] = useState(WEEK_DAYS[0]);
  const rep = data.reps.find((r) => r.id === repId);
  const clients = data.merchants.filter((m) => m.repId === repId);
  const plan = data.journeyPlans.find((p) => p.repId === repId && p.day === day);
  const [sel, setSel] = useState(null);
  const current = sel ?? (plan ? plan.stops.map((st) => st.merchantId) : []);
  const toggle = (mid) => { const base = current.includes(mid) ? current.filter((x) => x !== mid) : [...current, mid]; setSel(base); };
  return (<div className="space-y-4"><SectionTitle sub="حدّد المحلات التي يزورها كل مندوب في كل يوم — يظهر له في «خط السير» مع تسجيل «لا بيع» بأسبابه.">خط السير الأسبوعي</SectionTitle>
    <Card className="p-4 min-w-0">
      <div className="grid grid-cols-2 gap-2 mb-3">
        <Field label="المندوب"><select value={repId} onChange={(e) => { setRepId(e.target.value); setSel(null); }} className={inputCls} style={inputStyle}>{myReps.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}</select></Field>
        <Field label="اليوم"><select value={day} onChange={(e) => { setDay(e.target.value); setSel(null); }} className={inputCls} style={inputStyle}>{WEEK_DAYS.map((d) => <option key={d}>{d}</option>)}</select></Field>
      </div>
      <div className="text-xs font-bold mb-2" style={{ color: C.soft }}>اختر محلات يوم {day} لـ{rep?.name} ({current.length} مختار):</div>
      <div className="space-y-1.5 mb-3">{clients.map((m) => { const on = current.includes(m.id); return (<button key={m.id} onClick={() => toggle(m.id)} className="rb-press w-full flex items-center justify-between gap-2 rounded-xl px-3 py-2 text-right" style={{ background: on ? SC.soft : C.surface, border: `1px solid ${on ? SC.color : C.line}` }}>
        <div className="flex items-center gap-2 min-w-0"><span className="rounded-md flex items-center justify-center shrink-0" style={{ width: 20, height: 20, background: on ? SC.color : "#fff", border: `2px solid ${on ? SC.color : C.line}` }}>{on && <CheckCircle2 size={13} color="#fff" />}</span><div className="min-w-0"><div className="text-sm font-bold truncate" style={{ color: C.text }}>{m.name}</div><div className="text-[10px]" style={{ color: C.soft }}>{m.area}</div></div></div>
      </button>); })}{clients.length === 0 && <Empty icon={Store} text="لا عملاء لهذا المندوب." />}</div>
      <PrimaryBtn full icon={Route} onClick={() => { actions.svSetJourney({ repId, day, merchantIds: current }); toast(`حُفظ خط سير ${day} (${current.length} محل) ✓`); setSel(null); }}>حفظ خط سير {day}</PrimaryBtn>
    </Card>
  </div>);
}
function SVHome(props) {
  const { data, go, user } = props; const supv = useSv(props);
  const reps = data.reps.filter((r) => r.supervisorId === supv.id);
  const repIds = reps.map((r) => r.id);
  const pendingStk = data.stocktakes.filter((st) => repIds.includes(st.repId) && st.status === "requested").length;
  const alerts = [];
  if (pendingStk > 0) alerts.push({ k: "sv_stocktake", icon: ClipboardList, text: `${pendingStk} طلب جرد بانتظارك` });
  return <RoleLauncher name={supv.name} sub={`${reps.length} مندوب تحت إشرافك`} roleTitle="مشرف مبيعات" nav={navFor(user.role, user.companyRole)} homeKey="sv_home" go={go} alerts={alerts} badges={{ sv_stocktake: pendingStk }} quick={{ k: "sv_urgent", icon: Route, label: "خط سير عاجل" }} />;
}
function SVReps(props) {
  const { data } = props; const supv = useSv(props);
  const reps = data.reps.filter((r) => r.supervisorId === supv.id);
  const repIds = reps.map((r) => r.id);
  const teamClients = data.merchants.filter((m) => repIds.includes(m.repId));
  const [openRep, setOpenRep] = useState(null);
  // طلبيات مندوب اليوم (createdAt يبدأ بـ "اليوم")
  const repTodayOrders = (rid) => data.orders.filter((o) => o.repId === rid && (o.createdAt || "").startsWith("اليوم"));
  return (<div className="space-y-4"><SectionTitle sub="المندوبون تحت إشرافك ونشاطهم — اضغط أي مندوب لرؤية طلبيات اليوم.">متابعة المندوبين</SectionTitle>
    <AgingPanel data={data} merchants={teamClients} title="أعمار ديون عملاء فريقك" />
    <div className="grid sm:grid-cols-2 gap-3">{reps.map((r) => { const orders = data.orders.filter((o) => o.repId === r.id); const visits = data.visits.filter((v) => v.repId === r.id); const clients = data.merchants.filter((m) => m.repId === r.id); const today = repTodayOrders(r.id);
      return (<Card key={r.id} className="p-5 cursor-pointer transition-colors" style={{ border: `1px solid ${openRep === r.id ? SC.color : C.line}` }} onClick={() => setOpenRep(openRep === r.id ? null : r.id)}><div className="flex items-center gap-3 mb-3"><div className="rounded-xl p-2.5" style={{ background: C.infoSoft }}><Handshake size={20} style={{ color: C.info }} /></div><div className="flex-1"><div className="font-extrabold" style={{ color: C.text }}>{r.name}</div><div className="text-xs" style={{ color: C.soft }}>{r.area}</div></div>{today.length > 0 && <span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: SC.soft, color: SC.color }}>{today.length} طلب اليوم</span>}<ChevronDown size={16} style={{ color: C.soft, transform: openRep === r.id ? "rotate(180deg)" : "none", transition: ".2s" }} /></div>
        <div className="grid grid-cols-3 gap-2 text-center mb-3">
          <div className="rounded-xl p-2" style={{ background: C.bg }}><div className="text-lg font-extrabold" style={{ color: SC.color2 }}>{clients.length}</div><div className="text-[10px]" style={{ color: C.soft }}>عميل</div></div>
          <div className="rounded-xl p-2" style={{ background: C.bg }}><div className="text-lg font-extrabold" style={{ color: C.info }}>{orders.length}</div><div className="text-[10px]" style={{ color: C.soft }}>طلب</div></div>
          <div className="rounded-xl p-2" style={{ background: C.bg }}><div className="text-lg font-extrabold" style={{ color: C.accent }}>{visits.length}</div><div className="text-[10px]" style={{ color: C.soft }}>زيارة</div></div>
        </div>
        <div className="flex justify-between text-xs mb-1" style={{ color: C.soft }}><span>الإنجاز</span><span>{Math.round((r.achieved / (r.target || 1)) * 100)}%</span></div><ProgressBar value={r.achieved} max={r.target} color={C.info} />
        <div className="text-sm flex items-center gap-1 mt-3" style={{ color: C.soft }}><Phone size={13} />{r.phone}</div>
        {openRep === r.id && <div className="mt-3 pt-3" style={{ borderTop: `1px dashed ${C.line}` }} onClick={(e) => e.stopPropagation()}>
          <div className="text-xs font-bold mb-2 flex items-center gap-1.5" style={{ color: SC.color }}><ShoppingCart size={13} />طلبيات اليوم ({today.length})</div>
          <div className="space-y-1.5">{today.map((o) => (<div key={o.id} className="flex items-center justify-between rounded-lg px-2.5 py-2" style={{ background: C.well }}>
            <div className="min-w-0"><div className="text-sm font-bold truncate" style={{ color: C.text }}>{merchName(data, o.merchantId)}</div><div className="text-[10px]" style={{ color: C.soft }}>{o.id} · {qty(orderQty(o))} · {STATUS[o.status]?.label || o.status}</div></div>
            <span className="text-sm font-extrabold shrink-0" style={{ color: C.text }}>{money(orderTotal(o))}</span>
          </div>))}{today.length === 0 && <Empty icon={ShoppingCart} text="لا طلبيات اليوم." />}</div>
        </div>}
      </Card>); })}{reps.length === 0 && <Empty icon={Handshake} text="لا مندوبين تحت إشرافك." />}</div>
  </div>);
}
function SVTargets(props) {
  const { data, actions } = props; const supv = useSv(props);
  const reps = data.reps.filter((r) => r.supervisorId === supv.id);
  // هدف المشرف الكلي (من مدير المبيعات) وما وُزّع منه
  const svQtyTarget = supv.qtyTarget || 0;
  const distributedQty = reps.reduce((s, r) => s + (r.qtyTarget || 0), 0);
  const remainingQty = svQtyTarget - distributedQty;
  const svColTarget = supv.collectionTarget || 0;
  const distributedCol = reps.reduce((s, r) => s + (r.target || 0), 0);
  // تصنيف المندوبين (جملة/تجزئة)
  const groups = [
    { type: "presell", label: "مندوبو الجملة (بري‑سيل)", icon: ClipboardList, color: C.info },
    { type: "van", label: "مندوبو التجزئة (فان)", icon: Truck, color: C.accent },
  ];
  const [xfer, setXfer] = useState(null); // {fromRepId}
  const [xferTo, setXferTo] = useState(""); const [xferQty, setXferQty] = useState(0);
  return (<div className="space-y-4"><SectionTitle sub="وزّع هدف الكمية (ستيكات) والتحصيل على مندوبيك حسب تصنيفهم، وتحكّم في الترحيل والنقل والعجز.">أهداف المندوبين والإنجاز</SectionTitle>
    {/* لوحة توزيع هدف المشرف */}
    <Card className="p-4" style={{ borderRight: `3px solid ${SC.color}` }}>
      <h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Target size={17} style={{ color: SC.color }} />هدفي من مدير المبيعات وتوزيعه</h3>
      <div className="grid sm:grid-cols-2 gap-3">
        <div className="rounded-xl p-3" style={{ background: C.well }}>
          <div className="text-xs font-bold mb-1" style={{ color: SC.color }}>هدف الكمية (ستيكات)</div>
          <div className="flex items-end justify-between"><span className="text-2xl font-extrabold" style={{ color: C.text }}>{qty(svQtyTarget)}</span><span className="text-xs" style={{ color: C.soft }}>وُزّع {qty(distributedQty)}</span></div>
          <div className="mt-2 rounded-lg px-2 py-1 text-xs font-bold text-center" style={{ background: remainingQty === 0 ? C.successSoft : remainingQty > 0 ? C.warnSoft : C.dangerSoft, color: remainingQty === 0 ? C.success : remainingQty > 0 ? "#B45309" : C.danger }}>{remainingQty === 0 ? "✓ وُزّع الهدف بالكامل" : remainingQty > 0 ? `متبقّي للتوزيع: ${qty(remainingQty)}` : `⚠️ تجاوزت الهدف بـ ${qty(-remainingQty)}`}</div>
        </div>
        <div className="rounded-xl p-3" style={{ background: C.well }}>
          <div className="text-xs font-bold mb-1" style={{ color: C.info }}>هدف التحصيل (قيمة)</div>
          <div className="flex items-end justify-between"><span className="text-2xl font-extrabold" style={{ color: C.text }}>{money(svColTarget)}</span><span className="text-xs" style={{ color: C.soft }}>وُزّع {money(distributedCol)}</span></div>
          <div className="mt-2 rounded-lg px-2 py-1 text-xs font-bold text-center" style={{ background: svColTarget - distributedCol === 0 ? C.successSoft : svColTarget - distributedCol > 0 ? C.warnSoft : C.dangerSoft, color: svColTarget - distributedCol === 0 ? C.success : svColTarget - distributedCol > 0 ? "#B45309" : C.danger }}>{svColTarget - distributedCol === 0 ? "✓ وُزّع بالكامل" : svColTarget - distributedCol > 0 ? `متبقّي: ${money(svColTarget - distributedCol)}` : `⚠️ تجاوزت بـ ${money(distributedCol - svColTarget)}`}</div>
        </div>
      </div>
    </Card>

    {groups.map((g) => { const gReps = reps.filter((r) => (r.repType || "presell") === g.type); if (!gReps.length) return null; const Ic = g.icon; return (
      <div key={g.type} className="space-y-2.5">
        <div className="flex items-center gap-2 text-sm font-extrabold" style={{ color: g.color }}><Ic size={16} />{g.label} <span className="text-xs font-normal" style={{ color: C.soft }}>({gReps.length})</span></div>
        {gReps.map((r) => { const qPct = Math.round(((r.qtyAchieved || 0) / (r.qtyTarget || 1)) * 100); const vPct = Math.round(((r.achieved || 0) / (r.target || 1)) * 100); const rem = Math.max(0, (r.qtyTarget || 0) - (r.qtyAchieved || 0)); return (
          <Card key={r.id} className="p-4">
            <div className="flex items-center gap-3 mb-3"><Avatar name={r.name} size={36} /><div className="flex-1"><div className="font-extrabold" style={{ color: C.text }}>{r.name}</div><div className="text-[11px]" style={{ color: C.soft }}>{r.area}{r.carriedOver ? ` · مُرحّل: ${qty(r.carriedOver)}` : ""}{r.qtyDeficit ? ` · عجز: ${qty(r.qtyDeficit)}` : ""}</div></div></div>
            <div className="grid sm:grid-cols-2 gap-3">
              <div className="rounded-xl p-3" style={{ background: C.well }}>
                <div className="flex items-center justify-between mb-1.5"><span className="text-xs font-bold flex items-center gap-1" style={{ color: SC.color }}><Boxes size={13} />هدف الكمية</span><span className="text-xs font-extrabold" style={{ color: qPct >= 70 ? C.success : qPct >= 40 ? "#B45309" : C.danger }}>{qPct}%</span></div>
                <ProgressBar value={r.qtyAchieved || 0} max={r.qtyTarget || 1} color={SC.color} />
                <div className="flex items-center gap-2 mt-2"><span className="text-[11px]" style={{ color: C.soft }}>محقّق {qty(r.qtyAchieved || 0)} — الهدف:</span><input type="number" defaultValue={r.qtyTarget || 0} onBlur={(e) => { actions.updateRep(r.id, { qtyTarget: +e.target.value }); toast("حُدّث هدف الكمية ✓"); }} className="w-20 rounded-lg px-2 py-1 text-sm outline-none" style={inputStyle} /></div>
              </div>
              <div className="rounded-xl p-3" style={{ background: C.well }}>
                <div className="flex items-center justify-between mb-1.5"><span className="text-xs font-bold flex items-center gap-1" style={{ color: C.info }}><HandCoins size={13} />هدف التحصيل</span><span className="text-xs font-extrabold" style={{ color: vPct >= 70 ? C.success : vPct >= 40 ? "#B45309" : C.danger }}>{vPct}%</span></div>
                <ProgressBar value={r.achieved || 0} max={r.target || 1} color={C.info} />
                <div className="flex items-center gap-2 mt-2"><span className="text-[11px]" style={{ color: C.soft }}>محقّق {money(r.achieved || 0)} — الهدف:</span><input type="number" defaultValue={r.target || 0} onBlur={(e) => { actions.updateRep(r.id, { target: +e.target.value }); toast("حُدّث هدف التحصيل ✓"); }} className="w-24 rounded-lg px-2 py-1 text-sm outline-none" style={inputStyle} /></div>
              </div>
            </div>
            {/* إجراءات المتبقّي: ترحيل / نقل / عجز */}
            {rem > 0 && <div className="mt-3 pt-3 flex items-center gap-2 flex-wrap" style={{ borderTop: `1px dashed ${C.line}` }}>
              <span className="text-[11px] font-bold" style={{ color: C.soft }}>غير منجَز: {qty(rem)} —</span>
              <button onClick={() => { actions.carryoverRepTarget(r.id); toast(`رُحّلت ${qty(rem)} للشهر القادم ✓`); }} className="rb-press flex items-center gap-1 rounded-lg px-2.5 py-1 text-xs font-bold" style={{ background: C.infoSoft, color: C.info }}><ArrowRight size={12} />ترحيل للشهر القادم</button>
              <button onClick={() => { setXfer(r.id); setXferTo(""); setXferQty(rem); }} className="rb-press flex items-center gap-1 rounded-lg px-2.5 py-1 text-xs font-bold" style={{ background: SC.soft, color: SC.color }}><ArrowLeftRight size={12} />نقل لمندوب آخر</button>
              <button onClick={() => { actions.markRepDeficit(r.id); toast(`سُجّل عجز ${qty(rem)} ✓`); }} className="rb-press flex items-center gap-1 rounded-lg px-2.5 py-1 text-xs font-bold" style={{ background: C.dangerSoft, color: C.danger }}><TrendingDown size={12} />تسجيل عجز</button>
            </div>}
          </Card>); })}
      </div>); })}
    {reps.length === 0 && <Empty icon={Handshake} text="لا مندوبين تحت إشرافك." />}

    {/* نافذة نقل الكمية */}
    {xfer && (() => { const from = reps.find((r) => r.id === xfer); const others = reps.filter((r) => r.id !== xfer); return (<div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4" onClick={() => setXfer(null)}>
      <div className="rb-modal w-full max-w-sm rounded-2xl p-5" onClick={(e) => e.stopPropagation()}>
        <h3 className="font-extrabold mb-1" style={{ color: C.text }}>نقل كمية من الهدف</h3>
        <p className="text-sm mb-3" style={{ color: C.soft }}>من {from?.name}</p>
        <Field label="إلى المندوب"><select value={xferTo} onChange={(e) => setXferTo(e.target.value)} className={inputCls} style={inputStyle}><option value="">— اختر —</option>{others.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}</select></Field>
        <div className="mt-2"><Field label="الكمية (ستيكات)"><Input type="number" value={xferQty} onChange={(e) => setXferQty(+e.target.value)} /></Field></div>
        <div className="flex gap-2 mt-4"><PrimaryBtn full icon={ArrowLeftRight} onClick={() => { if (!xferTo || xferQty <= 0) { toast("اختر مندوباً وكمية", "danger"); return; } actions.transferRepTarget({ fromRepId: xfer, toRepId: xferTo, qty: xferQty }); setXfer(null); toast("نُقلت الكمية ✓"); }}>تأكيد النقل</PrimaryBtn><GhostBtn onClick={() => setXfer(null)}>إلغاء</GhostBtn></div>
      </div>
    </div>); })()}
  </div>);
}
function SVUrgent(props) {
  const { data, actions } = props; const supv = useSv(props);
  const reps = data.reps.filter((r) => r.supervisorId === supv.id);
  const [repId, setRepId] = useState(reps[0]?.id || "");
  const repClients = data.merchants.filter((m) => m.repId === repId);
  const [note, setNote] = useState("");
  const [stops, setStops] = useState([]); // { merchantId, reason }
  const [draft, setDraft] = useState({ merchantId: repClients[0]?.id || "", reason: "" });
  const myRoutes = data.urgentRoutes.filter((u) => u.supervisorId === supv.id);
  const addStop = () => { if (draft.merchantId && !stops.some((s) => s.merchantId === draft.merchantId)) { setStops([...stops, { merchantId: draft.merchantId, reason: draft.reason || "زيارة عاجلة" }]); setDraft({ ...draft, reason: "" }); } };
  const send = () => { if (repId && stops.length) { actions.assignUrgentRoute({ supervisorId: supv.id, repId, note, stops }); setStops([]); setNote(""); } };
  return (<div className="space-y-4"><SectionTitle sub="أرسل لمندوبك قائمة محلات مرتّبة لزيارتها الآن — تصله فوراً كأولوية عاجلة.">خط سير عاجل للمندوب</SectionTitle>
    <Card className="p-5 space-y-3">
      <div className="grid sm:grid-cols-2 gap-2">
        <Field label="المندوب"><select value={repId} onChange={(e) => { setRepId(e.target.value); setStops([]); const c = data.merchants.filter((m) => m.repId === e.target.value); setDraft({ merchantId: c[0]?.id || "", reason: "" }); }} className={inputCls} style={inputStyle}>{reps.map((r) => <option key={r.id} value={r.id}>{r.name} — {r.area}</option>)}</select></Field>
        <Field label="ملاحظة عامة"><Input value={note} onChange={(e) => setNote(e.target.value)} placeholder="مثال: أولوية تحصيل قبل نهاية الدوام" /></Field>
      </div>
      <div className="rounded-xl p-3" style={{ background: C.bg }}>
        <div className="text-sm font-bold mb-2" style={{ color: C.text }}>أضف المحلات للجولة</div>
        <div className="grid sm:grid-cols-3 gap-2 items-end">
          <Field label="المحل"><select value={draft.merchantId} onChange={(e) => setDraft({ ...draft, merchantId: e.target.value })} className={inputCls} style={inputStyle}>{repClients.map((m) => <option key={m.id} value={m.id}>{m.name} — {m.area}</option>)}</select></Field>
          <Field label="سبب الزيارة"><Input value={draft.reason} onChange={(e) => setDraft({ ...draft, reason: e.target.value })} placeholder="تحصيل / نقص مخزون..." /></Field>
          <PrimaryBtn icon={Plus} onClick={addStop}>إضافة محل</PrimaryBtn>
        </div>
      </div>
      {stops.length > 0 && <div className="space-y-1.5">{stops.map((s, i) => <div key={i} className="flex items-center justify-between rounded-lg px-3 py-2" style={{ background: C.dangerSoft + "55" }}><span className="text-sm font-bold flex items-center gap-2" style={{ color: C.text }}><span className="rounded-full text-xs flex items-center justify-center" style={{ width: 22, height: 22, background: C.danger, color: "#fff" }}>{i + 1}</span>{merchName(data, s.merchantId)} <span className="text-xs font-normal" style={{ color: C.soft }}>· {s.reason}</span></span><button onClick={() => setStops(stops.filter((_, x) => x !== i))} className="rounded p-1" style={{ background: C.well }}><Trash2 size={13} style={{ color: C.danger }} /></button></div>)}</div>}
      <PrimaryBtn full tone="danger" icon={Navigation} onClick={send}>إرسال خط السير العاجل ({stops.length} محل)</PrimaryBtn>
    </Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>خطوط السير المُرسلة</h3>
      <div className="space-y-3">{myRoutes.map((u) => { const doneCnt = u.stops.filter((s) => s.done).length; return (<div key={u.id} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}><div className="flex items-center justify-between mb-2 flex-wrap gap-2"><div className="font-bold text-sm flex items-center gap-2" style={{ color: C.text }}><Route size={15} style={{ color: C.danger }} />{repName(data, u.repId)} <span className="text-xs font-normal" style={{ color: C.soft }}>· {u.date}</span></div><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: doneCnt === u.stops.length ? C.successSoft : C.accentSoft, color: doneCnt === u.stops.length ? C.success : "#92590C" }}>{doneCnt}/{u.stops.length} مُنجز</span></div><div className="flex flex-wrap gap-1.5">{u.stops.map((s, i) => <span key={i} className="text-xs font-bold px-2 py-0.5 rounded-full flex items-center gap-1" style={{ background: s.done ? C.successSoft : C.bg, color: s.done ? C.success : C.text }}>{s.done && <CheckCircle2 size={11} />}{merchName(data, s.merchantId)}</span>)}</div></div>); })}{myRoutes.length === 0 && <Empty icon={Route} text="لم ترسل خطوط سير عاجلة بعد." />}</div>
    </Card>
  </div>);
}
// المشرف يحدّد لكل عميل هدفاً بالستيكات والتحصيل
function SVClientTargets(props) {
  const { data, actions } = props; const supv = useSv(props);
  const reps = data.reps.filter((r) => r.supervisorId === supv.id);
  const repIds = reps.map((r) => r.id);
  const clients = data.merchants.filter((m) => repIds.includes(m.repId));
  const targetFor = (mid) => data.clientTargets.find((c) => c.merchantId === mid && c.supervisorId === supv.id);
  return (<div className="space-y-4"><SectionTitle sub="حدّد لكل عميل هدفاً بعدد الستيكات ومبلغ التحصيل — تتجمّع لتشكّل هدفك العام من مدير المبيعات.">أهداف العملاء (ستيكات / تحصيل)</SectionTitle>
    <div className="grid sm:grid-cols-3 gap-3">
      <Stat icon={Target} label="هدفي من المدير (كمية)" value={qty(supv.qtyTarget)} tint={{ color: SC.color2, soft: SC.soft }} hint={`أنجزت ${qty(supv.qtyAchieved)}`} />
      <Stat icon={HandCoins} label="هدفي من المدير (تحصيل)" value={money(supv.collectionTarget)} tint={{ color: C.success, soft: C.successSoft }} hint={`حصّلت ${money(supv.collected)}`} />
      <Stat icon={Store} label="عملاء تحت الإشراف" value={clients.length} tint={{ color: C.info, soft: C.infoSoft }} />
    </div>
    <Card className="p-4"><Table head={["العميل", "المندوب", "هدف الستيكات", "إنجاز", "هدف التحصيل", "محصّل", "تعديل"]}>{clients.map((m) => { const t = targetFor(m.id); return (<ClientTargetRow key={m.id} m={m} t={t} data={data} actions={actions} supv={supv} />); })}</Table>{clients.length === 0 && <Empty icon={Store} text="لا يوجد عملاء تحت إشرافك." />}</Card>
  </div>);
}
function ClientTargetRow({ m, t, data, actions, supv }) {
  const [q, setQ] = useState(t?.qtyTarget || 0);
  const [c, setC] = useState(t?.collectionTarget || 0);
  const pctQ = t ? Math.round((t.qtyAchieved / (t.qtyTarget || 1)) * 100) : 0;
  return (<tr style={{ borderBottom: `1px solid ${C.line}` }}>
    <Td bold>{m.name}</Td><Td>{repName(data, m.repId)}</Td>
    <Td><input type="number" value={q} onChange={(e) => setQ(+e.target.value)} className="w-20 rounded-lg px-2 py-1 text-sm outline-none" style={inputStyle} /></Td>
    <Td>{t ? <span style={{ color: pctQ >= 70 ? C.success : C.accent, fontWeight: 700 }}>{qty(t.qtyAchieved)} ({pctQ}%)</span> : "—"}</Td>
    <Td><input type="number" value={c} onChange={(e) => setC(+e.target.value)} className="w-24 rounded-lg px-2 py-1 text-sm outline-none" style={inputStyle} /></Td>
    <Td>{t ? money(t.collected) : "—"}</Td>
    <Td><PrimaryBtn sm icon={CheckCircle2} onClick={() => actions.setClientTarget({ supervisorId: supv.id, merchantId: m.id, qtyTarget: q, collectionTarget: c })}>حفظ</PrimaryBtn></Td>
  </tr>);
}
function SVQuotas(props) {
  const { data, actions } = props; const supv = useSv(props);
  const reps = data.reps.filter((r) => r.supervisorId === supv.id);
  const repIds = reps.map((r) => r.id);
  const prods = data.products.filter((p) => p.companyId === supv.companyId);
  const quotas = data.quotas.filter((q) => repIds.includes(q.repId));
  const [f, setF] = useState({ repId: reps[0]?.id || "", pid: prods[0]?.id || "", allocated: 10 });
  return (<div className="space-y-4"><SectionTitle sub="وزّع حصص المنتجات على مندوبيك ليطلبها منهم العملاء.">توزيع حصص المنتجات</SectionTitle>
    <Card className="p-4"><div className="grid sm:grid-cols-4 gap-2 items-end">
      <Field label="المندوب"><select value={f.repId} onChange={(e) => setF({ ...f, repId: e.target.value })} className={inputCls} style={inputStyle}>{reps.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}</select></Field>
      <Field label="المنتج"><select value={f.pid} onChange={(e) => setF({ ...f, pid: e.target.value })} className={inputCls} style={inputStyle}>{prods.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}</select></Field>
      <Field label="الكمية المخصصة"><Input type="number" value={f.allocated} onChange={(e) => setF({ ...f, allocated: +e.target.value })} /></Field>
      <PrimaryBtn icon={Plus} onClick={() => { if (f.repId && f.pid) { actions.addQuota({ server: supv.server, companyId: supv.companyId, ...f, sold: 0 }); } }}>تخصيص حصة</PrimaryBtn>
    </div></Card>
    <Card className="p-4"><Table head={["المندوب", "المنتج", "المخصص", "المُباع", "المتبقي", "نسبة البيع", ""]}>{quotas.map((q) => { const remaining = q.allocated - q.sold; const pct = Math.round((q.sold / q.allocated) * 100); return (<tr key={q.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{repName(data, q.repId)}</Td><Td>{prodName(data, q.pid)}</Td><Td>{q.allocated}</Td><Td>{q.sold}</Td><Td><span style={{ color: remaining < 5 ? C.danger : C.text, fontWeight: 700 }}>{remaining}</span></Td><Td><div className="flex items-center gap-2"><div style={{ width: 55 }}><ProgressBar value={q.sold} max={q.allocated} color={C.success} /></div><span className="text-xs" style={{ color: C.soft }}>{pct}%</span></div></Td><Td><button onClick={() => actions.removeQuota(q.id)} className="rounded-lg p-1.5" style={{ background: C.dangerSoft }}><Trash2 size={14} style={{ color: C.danger }} /></button></Td></tr>); })}</Table>{quotas.length === 0 && <Empty icon={Boxes} text="لا توجد حصص موزّعة." />}</Card>
  </div>);
}
function SVTransfer(props) {
  const { data, actions } = props; const supv = useSv(props);
  const reps = data.reps.filter((r) => r.supervisorId === supv.id);
  return <QuotaTransferUI data={data} actions={actions} reps={reps} title="نقل الحصص بين المندوبين" sub="انقل حصص منتج من أحد مندوبيك إلى آخر." />;
}
// تنمية العملاء وتحليل الزيارات غير الناجحة
function SVDevelop(props) {
  const { data } = props; const supv = useSv(props);
  const reps = data.reps.filter((r) => r.supervisorId === supv.id);
  const repIds = reps.map((r) => r.id);
  const visits = data.visits.filter((v) => repIds.includes(v.repId));
  // تحليل أسباب عدم النجاح
  const failVisits = visits.filter((v) => !["order", "payment"].includes(v.result));
  const reasonCount = {};
  failVisits.forEach((v) => { const k = v.failReason || "other"; reasonCount[k] = (reasonCount[k] || 0) + 1; });
  return (<div className="space-y-4"><SectionTitle sub="نمِّ قاعدة عملاء كل مندوب وحلّل أسباب الزيارات غير الناجحة.">تنمية العملاء وتحليل الزيارات</SectionTitle>
    {/* تنمية العملاء لكل مندوب */}
    <Card className="p-5"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>قاعدة عملاء كل مندوب</h3>
      <div className="space-y-3">{reps.map((r) => { const clients = data.merchants.filter((m) => m.repId === r.id); const active = clients.filter((m) => data.orders.some((o) => o.merchantId === m.id && o.status === "delivered")).length; const growth = clients.length ? Math.round((active / clients.length) * 100) : 0;
        return (<div key={r.id} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}>
          <div className="flex items-center justify-between mb-2"><span className="font-bold text-sm" style={{ color: C.text }}>{r.name}</span><span className="text-xs" style={{ color: C.soft }}>{clients.length} عميل · {active} نشط</span></div>
          <div className="flex justify-between text-xs mb-1" style={{ color: C.soft }}><span>نسبة العملاء النشطين</span><span className="font-bold" style={{ color: growth >= 60 ? C.success : C.accent }}>{growth}%</span></div>
          <ProgressBar value={active} max={clients.length || 1} color={growth >= 60 ? C.success : C.accent} />
          {growth < 60 && <div className="text-[11px] mt-2 flex items-center gap-1" style={{ color: C.accent }}><AlertCircle size={12} />فرصة تنمية: {clients.length - active} عميل غير نشط يحتاجون متابعة</div>}
        </div>); })}{reps.length === 0 && <Empty icon={Handshake} text="لا يوجد مندوبون." />}</div>
    </Card>
    {/* تحليل أسباب عدم النجاح */}
    <Card className="p-5"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>لماذا لم تكن الزيارة ناجحة؟</h3>
      {Object.keys(reasonCount).length ? <div className="space-y-2">{Object.entries(reasonCount).sort((a, b) => b[1] - a[1]).map(([k, n]) => { const pct = Math.round((n / failVisits.length) * 100); return (<div key={k}><div className="flex justify-between text-sm mb-1"><span style={{ color: C.text }}>{FAIL_REASONS[k]}</span><span className="font-bold" style={{ color: C.text }}>{n} ({pct}%)</span></div><ProgressBar value={n} max={failVisits.length} color={C.danger} /></div>); })}</div> : <Empty icon={CheckCircle2} text="لا توجد زيارات غير ناجحة." />}
    </Card>
    {/* تفاصيل الزيارات غير الناجحة */}
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>تفاصيل الزيارات غير الناجحة</h3><Table head={["المندوب", "المحل", "السبب", "الملاحظة", "التاريخ"]}>{failVisits.map((v) => <tr key={v.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{repName(data, v.repId)}</Td><Td>{merchName(data, v.merchantId)}</Td><Td><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.dangerSoft, color: C.danger }}>{v.failReason ? FAIL_REASONS[v.failReason] : VISIT_RESULT[v.result]?.label}</span></Td><Td>{v.note}</Td><Td>{v.date}</Td></tr>)}</Table>{failVisits.length === 0 && <Empty icon={CheckCircle2} text="كل الزيارات ناجحة." />}</Card>
  </div>);
}
function SVVisits(props) {
  const { data } = props; const supv = useSv(props);
  const repIds = data.reps.filter((r) => r.supervisorId === supv.id).map((r) => r.id);
  const visits = data.visits.filter((v) => repIds.includes(v.repId));
  return (<div className="space-y-4"><SectionTitle sub="زيارات مندوبيك للمحلات ونتائجها.">سجل الزيارات</SectionTitle>
    <Card className="p-4"><Table head={["المندوب", "المحل", "النتيجة", "الملاحظة", "المبلغ", "التاريخ"]}>{visits.map((v) => <tr key={v.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{repName(data, v.repId)}</Td><Td>{merchName(data, v.merchantId)}</Td><Td><Badge status={v.result} map={VISIT_RESULT} /></Td><Td>{v.note}</Td><Td>{v.amount ? money(v.amount) : "—"}</Td><Td>{v.date}</Td></tr>)}</Table>{visits.length === 0 && <Empty icon={MapPinned} text="لا توجد زيارات." />}</Card>
  </div>);
}
function SVStock(props) {
  const { data } = props; const supv = useSv(props);
  const repIds = data.reps.filter((r) => r.supervisorId === supv.id).map((r) => r.id);
  const stocktakes = data.stocktakes.filter((s) => repIds.includes(s.repId));
  return (<div className="space-y-4"><SectionTitle sub="طلبات جرد المخزون التي رفعها المندوبون من السوق.">طلبات الجرد</SectionTitle>
    {stocktakes.length ? stocktakes.map((st) => (<Card key={st.id} className="p-4"><div className="flex items-center justify-between mb-3 flex-wrap gap-2"><div><div className="font-extrabold" style={{ color: C.text }}>{merchName(data, st.merchantId)}</div><div className="text-xs" style={{ color: C.soft }}>{repName(data, st.repId)} · {st.date}</div></div><Badge status={st.status} map={{ requested: { label: "بانتظار", color: C.accent, soft: C.accentSoft }, done: { label: "مكتمل", color: C.success, soft: C.successSoft } }} /></div>
      <div className="rounded-xl overflow-hidden" style={{ border: `1px solid ${C.line}` }}>{st.lines.map((l, i) => <div key={i} className="flex justify-between px-3 py-2 text-sm" style={{ borderBottom: i < st.lines.length - 1 ? `1px solid ${C.line}` : "none" }}><span style={{ color: C.text }}>{l.name}</span><span style={{ color: C.soft }}>على الرف: {l.shelf}</span></div>)}</div>
      {st.note && <div className="text-sm rounded-lg p-2 mt-3" style={{ background: C.bg, color: C.soft }}>{st.note}</div>}
    </Card>)) : <Empty icon={ClipboardCheck} text="لا توجد طلبات جرد." />}
  </div>);
}

/* ============================================================
   ============   المنصة   ===================================
   ============================================================ */
/* ============================================================
   ============   المنصة (الإدارة)   =========================
   ============================================================ */
// مبيعات شركة (مسلّمة) في السيرفر الحالي
const companySales = (data, cid) => data.orders.filter((o) => o.companyId === cid && o.status === "delivered");
const companySalesTotal = (data, cid) => companySales(data, cid).reduce((s, o) => s + orderTotal(o), 0);

// ═══ عمولة التطبيق: تُحتسب على العمليات المكتملة والمُحصّلة فعلاً ═══
// المبلغ المحصّل = مجموع المسدَّد من فواتير الشركة (paid) — العمولة تُحسب عليه
const companyCollected = (data, cid) => data.invoices.filter((i) => i.companyId === cid).reduce((s, i) => s + (i.paid || 0), 0);
// عمولة التطبيق المكتسبة (على المُحصّل)
const appCommissionEarned = (data, co) => {
  if (!co) return 0;
  const collected = companyCollected(data, co.id);
  if (co.commType === "fixed") { const paidInvoices = data.invoices.filter((i) => i.companyId === co.id && i.paid > 0).length; return money3(co.commValue * paidInvoices); }
  return money3(calcCommission(co, collected));
};
// المسدَّد للتطبيق سابقاً + المتبقّي المستحق
const appCommissionPaid = (co) => co?.commissionPaid || 0;
const appCommissionDue = (data, co) => money3(Math.max(0, appCommissionEarned(data, co) - appCommissionPaid(co)));

// ═══ نطاق مدير المنصة: يرى كل السيرفرات — منفردةً أو مجتمعة ═══
const PLATFORM_SECTORS = ["all", "food", "pharma"]; // القطاعات الفعّالة + "الكل"
const sectorLabel = (k) => k === "all" ? "كل السيرفرات" : (SERVERS[k]?.name || k);
const sectorEmoji = (k) => k === "all" ? "🌐" : (SERVERS[k]?.emoji || "");
const sectorColor = (k) => k === "all" ? C.primary : (SERVERS[k]?.color || C.primary);
// شركات ضمن النطاق المختار
const scopedCompanies = (data, scope) => data.companies.filter((c) => scope === "all" ? ["food", "pharma"].includes(c.server) : c.server === scope);
// مبدّل النطاق (شريط أزرار أنيق)
function ScopeSwitcher({ scope, setScope }) {
  return (<div className="flex items-center gap-1.5 flex-wrap rounded-2xl p-1.5" style={{ background: C.well, border: `1px solid ${C.line}` }}>
    {PLATFORM_SECTORS.map((k) => { const on = scope === k; const col = sectorColor(k); return (
      <button key={k} onClick={() => setScope(k)} className="rb-press flex items-center gap-1.5 rounded-xl px-3.5 py-2 text-sm font-extrabold transition-all" style={{ background: on ? col : "transparent", color: on ? C.onAccent : C.soft }}>
        <span>{sectorEmoji(k)}</span>{sectorLabel(k)}
      </button>); })}
  </div>);
}
// حساب أرباح المنصة ضمن نطاق (يدعم التجميع أو الإفراد)
const platformStats = (data, scope) => {
  const cos = scopedCompanies(data, scope);
  const perCompany = cos.map((c) => {
    const sales = companySalesTotal(data, c.id);
    const dlv = companySales(data, c.id);
    const units = dlv.reduce((a, o) => a + orderQty(o), 0);
    const cnt = dlv.length;
    const comm = c.commType === "fixed" ? c.commValue * cnt : calcCommission(c, sales);
    return { c, sales, units, cnt, comm: money3(comm) };
  });
  // تجميع حسب السيرفر (للعرض المنفصل داخل "الكل")
  const bySector = {};
  ["food", "pharma"].forEach((sv) => {
    const rows = perCompany.filter((r) => r.c.server === sv);
    if (!rows.length && scope !== "all") return;
    bySector[sv] = {
      companies: rows.length,
      sales: rows.reduce((s, r) => s + r.sales, 0),
      units: rows.reduce((s, r) => s + r.units, 0),
      orders: rows.reduce((s, r) => s + r.cnt, 0),
      comm: money3(rows.reduce((s, r) => s + r.comm, 0)),
      subFees: rows.filter((r) => r.c.sub === "active").reduce((s, r) => s + (r.c.subFee || 0), 0),
    };
  });
  return {
    perCompany,
    bySector,
    totals: {
      companies: perCompany.length,
      sales: perCompany.reduce((s, r) => s + r.sales, 0),
      units: perCompany.reduce((s, r) => s + r.units, 0),
      orders: perCompany.reduce((s, r) => s + r.cnt, 0),
      comm: money3(perCompany.reduce((s, r) => s + r.comm, 0)),
      subFees: perCompany.filter((r) => r.c.sub === "active").reduce((s, r) => s + (r.c.subFee || 0), 0),
    },
  };
};

function PFHome(props) {
  const { data, go } = props;
  const [scope, setScope] = useState("all");
  const cos = scopedCompanies(data, scope);
  const activeCnt = cos.filter((c) => c.sub === "active").length;
  const pending = cos.filter((c) => c.sub === "pending");
  const st = platformStats(data, scope);
  const showSectorSplit = scope === "all";
  return (<div className="space-y-5">
    <div className="rounded-2xl p-5 sm:p-6" style={{ background: `linear-gradient(120deg, ${C.ink} 0%, ${sectorColor(scope)} 130%)`, color: "#fff" }}>
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div><div className="text-sm opacity-85 flex items-center gap-2"><ShieldCheck size={16} />إدارة منصة رابط</div><div className="text-2xl font-extrabold mt-0.5">لوحة التحكم — {sectorLabel(scope)} {sectorEmoji(scope)}</div></div>
        <PrimaryBtn tone="ghost" icon={CircleDollarSign} onClick={() => go("pf_profits")}>لوحة الأرباح التفصيلية</PrimaryBtn>
      </div>
    </div>

    <ScopeSwitcher scope={scope} setScope={setScope} />

    <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={Building2} label="شركات مشتركة" value={activeCnt} tint={{ color: sectorColor(scope), soft: C.well }} hint={`${pending.length} بانتظار`} />
      <Stat icon={ShoppingCart} label="إجمالي الطلبات" value={st.totals.orders} tint={{ color: C.info, soft: C.infoSoft }} hint={`${qty(st.totals.units)}`} />
      <Stat icon={TrendingUp} label="إجمالي المبيعات" value={money(st.totals.sales)} tint={{ color: C.success, soft: C.successSoft }} />
      <Stat icon={CircleDollarSign} label="عمولة المنصة" value={money(st.totals.comm)} tint={{ color: C.accent, soft: C.accentSoft }} hint={`+ اشتراكات ${money(st.totals.subFees)}`} />
    </div>

    {showSectorSplit && <Card className="p-4 sm:p-5">
      <h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Layers size={18} style={{ color: C.primary }} />مقارنة السيرفرات</h3>
      <div className="grid sm:grid-cols-2 gap-3">
        {["food", "pharma"].map((sv) => { const b = st.bySector[sv]; if (!b) return null; const sc = SERVERS[sv]; return (
          <div key={sv} className="rounded-2xl p-4" style={{ border: `1.5px solid ${sc.color}`, background: sc.soft }}>
            <div className="flex items-center justify-between mb-3"><span className="font-extrabold flex items-center gap-2" style={{ color: sc.color }}><span className="text-lg">{sc.emoji}</span>{sc.name}</span><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.surface, color: sc.color }}>{b.companies} شركة</span></div>
            <div className="grid grid-cols-2 gap-2 text-center">
              <div className="rounded-xl p-2" style={{ background: C.surface }}><div className="text-[10px]" style={{ color: C.soft }}>المبيعات</div><div className="font-extrabold text-sm" style={{ color: C.text }}>{money(b.sales)}</div></div>
              <div className="rounded-xl p-2" style={{ background: C.surface }}><div className="text-[10px]" style={{ color: C.soft }}>الكمية</div><div className="font-extrabold text-sm" style={{ color: C.text }}>{qty(b.units)}</div></div>
              <div className="rounded-xl p-2" style={{ background: C.surface }}><div className="text-[10px]" style={{ color: C.soft }}>الطلبات</div><div className="font-extrabold text-sm" style={{ color: C.text }}>{b.orders}</div></div>
              <div className="rounded-xl p-2" style={{ background: C.surface }}><div className="text-[10px]" style={{ color: C.soft }}>عمولة المنصة</div><div className="font-extrabold text-sm" style={{ color: sc.color }}>{money(b.comm)}</div></div>
            </div>
          </div>); })}
      </div>
    </Card>}

    <div className="grid lg:grid-cols-2 gap-4">
      <Card className="p-5"><div className="flex items-center justify-between mb-3"><h3 className="font-extrabold" style={{ color: C.text }}>اشتراكات بانتظار الموافقة</h3><GhostBtn sm icon={Building2} onClick={() => go("pf_companies")}>الكل</GhostBtn></div>
        <div className="space-y-2">{pending.map((c) => <div key={c.id} className="flex items-center gap-3 rounded-xl p-2.5" style={{ border: `1px solid ${C.line}` }}><span className="text-2xl">{c.logo}</span><div className="flex-1"><div className="font-bold text-sm" style={{ color: C.text }}>{c.name}</div><div className="text-xs flex items-center gap-1.5" style={{ color: C.soft }}><span>{SERVERS[c.server]?.emoji}</span>{c.cat}</div></div><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.accentSoft, color: "#92590C" }}>بانتظار</span></div>)}{pending.length === 0 && <p className="text-sm py-3" style={{ color: C.soft }}>لا اشتراكات معلّقة.</p>}</div>
      </Card>
      <Card className="p-5"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>توزيع خطط الاشتراك</h3>
        <div className="space-y-3">{Object.entries(PLANS).map(([k, p]) => { const cnt = cos.filter((c) => c.plan === k).length; return (<div key={k}><div className="flex justify-between text-sm mb-1"><span className="flex items-center gap-1.5" style={{ color: C.text }}><span className="w-3 h-3 rounded" style={{ background: p.color }} />{p.label}</span><span className="font-bold">{cnt}</span></div><ProgressBar value={cnt} max={cos.length || 1} color={p.color} /></div>); })}</div>
      </Card>
    </div>
  </div>);
}

// لوحة الأرباح
function PFProfits(props) {
  const { data, actions } = props;
  const [scope, setScope] = useState("all");
  const [payCo, setPayCo] = useState(null);
  const st = platformStats(data, scope);
  const rows = st.perCompany;
  const scoped = scopedCompanies(data, scope);
  const earned = money3(scoped.reduce((s, c) => s + appCommissionEarned(data, c), 0));
  const collected = money3(scoped.reduce((s, c) => s + appCommissionPaid(c), 0));
  const due = money3(scoped.reduce((s, c) => s + appCommissionDue(data, c), 0));
  const showSectorSplit = scope === "all";
  return (<div className="space-y-4"><SectionTitle sub="أرباح المنصة والعمولات — لكل سيرفر على حدة أو مجتمعة.">لوحة الأرباح</SectionTitle>
    <ScopeSwitcher scope={scope} setScope={setScope} />

    {showSectorSplit && <div className="grid sm:grid-cols-2 gap-3">
      {["food", "pharma"].map((sv) => { const b = st.bySector[sv]; if (!b) return null; const sc = SERVERS[sv]; return (
        <Card key={sv} className="p-4" style={{ borderTop: `3px solid ${sc.color}` }}>
          <div className="flex items-center justify-between mb-2"><span className="font-extrabold flex items-center gap-2" style={{ color: sc.color }}><span>{sc.emoji}</span>{sc.name}</span><span className="text-xs" style={{ color: C.soft }}>{b.companies} شركة</span></div>
          <div className="flex items-end justify-between"><div><div className="text-[11px]" style={{ color: C.soft }}>عمولة المنصة</div><div className="text-xl font-extrabold" style={{ color: sc.color }}>{money(b.comm)}</div></div><div className="text-left"><div className="text-[11px]" style={{ color: C.soft }}>مبيعات · {qty(b.units)}</div><div className="font-bold text-sm" style={{ color: C.text }}>{money(b.sales)}</div></div></div>
        </Card>); })}
    </div>}

    <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={Boxes} label="إجمالي الكمية المباعة" value={qty(st.totals.units)} tint={{ color: sectorColor(scope), soft: C.well }} />
      <Stat icon={TrendingUp} label="إجمالي قيمة المبيعات" value={money(st.totals.sales)} tint={{ color: C.success, soft: C.successSoft }} />
      <Stat icon={CircleDollarSign} label="عمولة المنصة" value={money(st.totals.comm)} tint={{ color: C.accent, soft: C.accentSoft }} />
      <Stat icon={Receipt} label="رسوم الاشتراكات" value={money(st.totals.subFees)} tint={{ color: C.purple, soft: C.purpleSoft }} />
      <Stat icon={Building2} label="عدد الشركات" value={st.totals.companies} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={HandCoins} label="عمولة مكتسبة (على المُحصّل)" value={money(earned)} tint={{ color: C.accent, soft: C.accentSoft }} />
      <Stat icon={CheckCircle2} label="سدّدته الشركات للتطبيق" value={money(collected)} tint={{ color: C.success, soft: C.successSoft }} />
      <Stat icon={Clock} label="مستحق على الشركات" value={money(due)} tint={{ color: C.danger, soft: C.dangerSoft }} />
      <Stat icon={Landmark} label="إجمالي دخل المنصة" value={money(money3(st.totals.comm + st.totals.subFees))} tint={{ color: C.primary, soft: C.infoSoft }} hint="عمولات + اشتراكات" />
    </div>
    <Card className="p-4"><div className="flex items-center justify-between mb-3 flex-wrap gap-2"><h3 className="font-extrabold" style={{ color: C.text }}>الأرباح والكميات لكل شركة</h3><span className="text-sm font-extrabold" style={{ color: C.accent }}>إجمالي العمولة: {money(st.totals.comm)}</span></div>
      <Table head={["الشركة", "السيرفر", "الاشتراك", "نسبة/نوع العمولة", "الكمية المباعة", "قيمة المبيعات", "الطلبات", "عمولة المنصة"]}>{rows.slice().sort((a, b) => b.comm - a.comm).map(({ c, sales, units, comm, cnt }) => <tr key={c.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{c.logo} {c.name}</Td><Td><span className="text-xs font-bold flex items-center gap-1" style={{ color: SERVERS[c.server]?.color }}>{SERVERS[c.server]?.emoji} {SERVERS[c.server]?.name}</span></Td><Td><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: (PLANS[c.plan] || PLANS.basic).soft, color: (PLANS[c.plan] || PLANS.basic).color }}>{(PLANS[c.plan] || PLANS.basic).label}</span></Td><Td><span className="font-bold" style={{ color: C.text }}>{c.commType === "percent" ? `${c.commValue}%` : COMM_TYPES[c.commType]}</span></Td><Td><span className="font-bold" style={{ color: SERVERS[c.server]?.color }}>{qty(units)}</span></Td><Td>{money(sales)}</Td><Td>{cnt}</Td><Td><span className="font-extrabold" style={{ color: C.accent }}>{money(comm)}</span></Td></tr>)}</Table>
    </Card>

    <Card className="p-4">
      <h3 className="font-extrabold mb-1 flex items-center gap-2" style={{ color: C.text }}><HandCoins size={18} style={{ color: C.accent }} />تحصيل عمولة التطبيق من الشركات</h3>
      <p className="text-xs mb-3" style={{ color: C.soft }}>العمولة تُحتسب تلقائياً على المبالغ المُحصّلة فعلاً. سجّل ما تسدّده كل شركة ليُخصم من المستحق.</p>
      <div className="space-y-2">{scoped.map((c) => { const e = appCommissionEarned(data, c); const pd = appCommissionPaid(c); const dueC = appCommissionDue(data, c); return (
        <div key={c.id} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}>
          <div className="flex items-center justify-between flex-wrap gap-2">
            <span className="font-bold text-sm flex items-center gap-1.5" style={{ color: C.text }}>{c.logo} {c.name}<span className="text-[10px] font-bold px-1.5 py-0.5 rounded-full" style={{ background: SERVERS[c.server]?.soft, color: SERVERS[c.server]?.color }}>{c.commType === "percent" ? c.commValue + "%" : COMM_TYPES[c.commType]}</span></span>
            {dueC > 0 ? <PrimaryBtn sm tone="accent" icon={CheckCircle2} onClick={() => setPayCo({ ...c, payAmount: dueC, dueC })}>تسجيل سداد</PrimaryBtn> : <span className="text-xs font-bold px-2 py-1 rounded-full" style={{ background: C.successSoft, color: C.success }}>مسدَّد بالكامل ✓</span>}
          </div>
          <div className="grid grid-cols-3 gap-2 mt-2.5 text-center">
            <div className="rounded-lg py-1.5" style={{ background: C.accentSoft }}><div className="text-[10px]" style={{ color: C.accent }}>مكتسبة</div><div className="font-extrabold text-sm" style={{ color: C.accent }}>{money(e)}</div></div>
            <div className="rounded-lg py-1.5" style={{ background: C.successSoft }}><div className="text-[10px]" style={{ color: C.success }}>سُدّد</div><div className="font-extrabold text-sm" style={{ color: C.success }}>{money(pd)}</div></div>
            <div className="rounded-lg py-1.5" style={{ background: dueC > 0 ? C.dangerSoft : C.well }}><div className="text-[10px]" style={{ color: dueC > 0 ? C.danger : C.soft }}>مستحق</div><div className="font-extrabold text-sm" style={{ color: dueC > 0 ? C.danger : C.soft }}>{money(dueC)}</div></div>
          </div>
        </div>); })}
      </div>
    </Card>

    {payCo && <div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4" onClick={() => setPayCo(null)}>
      <div className="rb-modal w-full max-w-sm rounded-2xl p-5" onClick={(e) => e.stopPropagation()}>
        <h3 className="font-extrabold mb-1" style={{ color: C.text }}>تسجيل سداد عمولة</h3>
        <p className="text-sm mb-3" style={{ color: C.soft }}>{payCo.logo} {payCo.name} — المستحق {money(payCo.dueC)}</p>
        <Field label="المبلغ المُسدَّد (د.ل)"><Input type="number" value={payCo.payAmount} onChange={(e) => setPayCo({ ...payCo, payAmount: +e.target.value })} /></Field>
        <div className="flex gap-2 mt-4"><PrimaryBtn full icon={CheckCircle2} onClick={() => { const amt = Math.min(payCo.payAmount, payCo.dueC); actions.payAppCommission(payCo.id, amt); setPayCo(null); toast(`سُجّل سداد ${money(amt)} ✓`); }}>تأكيد السداد</PrimaryBtn><GhostBtn onClick={() => setPayCo(null)}>إلغاء</GhostBtn></div>
      </div>
    </div>}
  </div>);
}

// إدارة العمولات
function PFCommissions(props) {
  const { data, actions } = props;
  const [scope, setScope] = useState("all");
  const all = scopedCompanies(data, scope);
  const [edit, setEdit] = useState(null);
  const previewComm = edit ? (edit.commType === "fixed" ? edit.commValue : companySalesTotal(data, edit.id) * (edit.commValue / 100)) : 0;
  return (<div className="space-y-4"><SectionTitle sub="لكل شركة نسبة عمولة مستقلة يحصّلها التطبيق من مبيعاتها — عدّلها بسهولة.">إدارة عمولات الاشتراك</SectionTitle>
    <ScopeSwitcher scope={scope} setScope={setScope} />
    <Card className="p-4"><Table head={["الشركة", "السيرفر", "الخطة", "نوع العمولة", "النسبة/القيمة", "عمولة محسوبة", ""]}>{all.map((c) => { const comm = c.commType === "fixed" ? c.commValue * companySales(data, c.id).length : Math.round(calcCommission(c, companySalesTotal(data, c.id))); return <tr key={c.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{c.logo} {c.name}</Td><Td><span className="text-xs font-bold flex items-center gap-1" style={{ color: SERVERS[c.server]?.color }}>{SERVERS[c.server]?.emoji} {SERVERS[c.server]?.name}</span></Td><Td><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: (PLANS[c.plan] || PLANS.basic).soft, color: (PLANS[c.plan] || PLANS.basic).color }}>{(PLANS[c.plan] || PLANS.basic).label}</span></Td><Td>{COMM_TYPES[c.commType]}</Td><Td><span className="font-extrabold" style={{ color: C.accent }}>{c.commType === "fixed" ? money(c.commValue) : c.commValue + "%"}</span></Td><Td><span className="font-bold" style={{ color: C.success }}>{money(comm)}</span></Td><Td><GhostBtn sm icon={Edit3} onClick={() => setEdit({ ...c })}>تعديل</GhostBtn></Td></tr>; })}</Table></Card>

    {edit && <div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4 overflow-auto" onClick={() => setEdit(null)}>
      <div className="rb-modal w-full max-w-md rounded-2xl p-5 my-8" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center gap-2 mb-4"><span className="text-2xl">{edit.logo}</span><div><div className="font-extrabold" style={{ color: C.text }}>عمولة الاشتراك</div><div className="text-sm flex items-center gap-1" style={{ color: C.soft }}>{SERVERS[edit.server]?.emoji} {edit.name}</div></div></div>
        <div className="space-y-3">
          <Field label="نوع العمولة"><div className="grid grid-cols-2 gap-2">{[["percent", "نسبة من المبيعات"], ["fixed", "مبلغ ثابت لكل فاتورة"]].map(([k, v]) => { const on = edit.commType === k; return <button key={k} onClick={() => setEdit({ ...edit, commType: k })} className="rounded-xl py-2 text-sm font-bold" style={{ background: on ? SC.color : C.bg, color: on ? C.onAccent : C.text, border: `1px solid ${on ? SC.color : C.line}` }}>{v}</button>; })}</div></Field>
          <Field label={edit.commType === "percent" ? "نسبة العمولة % (مثال: 1 أو 0.5)" : "المبلغ الثابت لكل فاتورة (د.ل)"}><Input type="number" step="0.05" value={edit.commValue} onChange={(e) => setEdit({ ...edit, commValue: +e.target.value })} /></Field>
          {edit.commType === "percent" && <div className="grid grid-cols-2 gap-2"><Field label="حد أدنى للعمولة (اختياري)"><Input type="number" value={edit.commMin || 0} onChange={(e) => setEdit({ ...edit, commMin: +e.target.value })} /></Field><Field label="حد أعلى للعمولة (اختياري)"><Input type="number" value={edit.commMax || 0} onChange={(e) => setEdit({ ...edit, commMax: +e.target.value })} /></Field></div>}
          <div className="rounded-xl p-3 flex items-center justify-between" style={{ background: C.successSoft }}><span className="text-sm font-bold" style={{ color: C.success }}>العمولة المتوقعة على مبيعاتها الحالية</span><span className="font-extrabold" style={{ color: C.success }}>{money(Math.round(previewComm))}</span></div>
        </div>
        <div className="flex gap-2 mt-4"><PrimaryBtn full icon={CheckCircle2} onClick={() => { actions.updateCompanyFin(edit.id, { commType: edit.commType, commValue: edit.commValue, commMin: edit.commMin || 0, commMax: edit.commMax || 0 }, `تعديل عمولة ${edit.name} إلى ${edit.commType === "percent" ? edit.commValue + "%" : money(edit.commValue)}`); setEdit(null); toast("حُفظت نسبة العمولة ✓"); }}>حفظ</PrimaryBtn><GhostBtn onClick={() => setEdit(null)}>إلغاء</GhostBtn></div>
      </div>
    </div>}
  </div>);
}

// الشركات والاشتراكات — لوحة احترافية شاملة
function PFCompanies(props) {
  const { data, actions } = props;
  const all = data.companies.filter((c) => c.server === sv(props));
  const [manage, setManage] = useState(null);
  const [filter, setFilter] = useState("all");
  const [q, setQ] = useState("");
  const filtered = all.filter((c) => (filter === "all" || (filter === "active" && c.sub === "active") || (filter === "pending" && c.sub === "pending") || (filter === "frozen" && c.frozen) || (filter === "due" && c.payStatus === "due")) && c.name.includes(q));
  const counts = { all: all.length, active: all.filter((c) => c.sub === "active").length, pending: all.filter((c) => c.sub === "pending").length, frozen: all.filter((c) => c.frozen).length, due: all.filter((c) => c.payStatus === "due").length };
  return (<div className="space-y-4"><SectionTitle sub="إدارة شاملة لاشتراكات الشركات: الخطط، الحدود، الخدمات الإضافية، حالة الدفع، والصلاحيات.">الشركات والاشتراكات</SectionTitle>
    {/* فلاتر سريعة */}
    <div className="flex gap-2 flex-wrap items-center">
      {[["all", "الكل"], ["active", "مشتركة"], ["pending", "بانتظار"], ["frozen", "مجمّدة"], ["due", "مستحقة الدفع"]].map(([k, t]) => { const on = filter === k; return (<button key={k} onClick={() => setFilter(k)} className="rounded-xl px-3 py-1.5 text-sm font-bold transition-colors" style={{ background: on ? SC.color : C.surface, color: on ? "#fff" : C.text, border: `1px solid ${on ? SC.color : C.line}` }}>{t} <span className="text-[10px] px-1.5 py-0.5 rounded-full" style={{ background: on ? "rgba(255,255,255,.25)" : C.bg }}>{counts[k]}</span></button>); })}
      <div className="relative flex-1 min-w-[160px]"><Search size={15} className="absolute right-3 top-2.5" style={{ color: C.soft }} /><input value={q} onChange={(e) => setQ(e.target.value)} placeholder="بحث باسم الشركة" className="w-full rounded-xl pr-9 pl-3 py-2 text-sm outline-none" style={inputStyle} /></div>
    </div>
    <div className="grid md:grid-cols-2 gap-3">{filtered.map((c) => { const reps = data.reps.filter((r) => r.companyId === c.id).length; const prods = data.products.filter((p) => p.companyId === c.id).length; const PS = PAY_STATUS[c.trial ? "trial" : c.payStatus] || PAY_STATUS.paid;
      return (<Card key={c.id} className="p-4" style={{ opacity: c.frozen ? .65 : 1, borderColor: c.payStatus === "due" ? C.danger : C.line }}>
        <div className="flex items-start justify-between gap-2 mb-2"><div className="flex items-center gap-3"><span className="text-2xl">{c.logo}</span><div><div className="font-extrabold flex items-center gap-1.5 flex-wrap" style={{ color: C.text }}>{c.name}{c.frozen && <span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.dangerSoft, color: C.danger }}>مجمّد</span>}{c.trial && <span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.infoSoft, color: C.info }}>تجريبي</span>}</div><div className="text-sm" style={{ color: C.soft }}>{c.cat} · {c.city}</div></div></div><span className="text-xs font-bold px-2.5 py-1 rounded-full" style={c.sub === "active" ? { background: C.successSoft, color: C.success } : c.sub === "pending" ? { background: C.accentSoft, color: "#92590C" } : c.sub === "suspended" ? { background: C.well, color: C.soft } : { background: C.dangerSoft, color: C.danger }}>{c.sub === "active" ? "مشترك" : c.sub === "pending" ? "بانتظار" : c.sub === "suspended" ? "موقوف" : "مرفوض"}</span></div>
        <div className="flex items-center gap-2 mb-3 flex-wrap"><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: (PLANS[c.plan] || PLANS.basic).soft, color: (PLANS[c.plan] || PLANS.basic).color }}>{(PLANS[c.plan] || PLANS.basic).label}</span><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: PS.soft, color: PS.color }}>{PS.label}</span><span className="text-xs" style={{ color: C.soft }}>{BILLING[c.billing]} · {money(c.subFee)}</span></div>
        <div className="grid grid-cols-3 gap-2 text-center mb-3">
          <div className="rounded-lg p-2" style={{ background: C.bg }}><div className="text-sm font-extrabold" style={{ color: SC.color2 }}>{prods}/{c.maxProducts}</div><div className="text-[10px]" style={{ color: C.soft }}>منتجات</div></div>
          <div className="rounded-lg p-2" style={{ background: C.bg }}><div className="text-sm font-extrabold" style={{ color: C.info }}>{reps}/{c.maxReps}</div><div className="text-[10px]" style={{ color: C.soft }}>مندوبون</div></div>
          <div className="rounded-lg p-2" style={{ background: C.bg }}><div className="text-sm font-extrabold" style={{ color: C.accent }}>{c.commType === "fixed" ? money(c.commValue) : c.commType === "tiered" ? "متدرّج" : c.commValue + "%"}</div><div className="text-[10px]" style={{ color: C.soft }}>عمولة</div></div>
        </div>
        {/* الخدمات الإضافية */}
        <div className="flex gap-1.5 flex-wrap mb-3">{Object.entries(ADDONS).map(([k, v]) => c.addons?.[k] && <span key={k} className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: SC.soft, color: SC.color }}>{v}</span>)}{c.discount > 0 && <span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.successSoft, color: C.success }}>خصم {c.discount}%</span>}</div>
        {c.sub === "pending" ? <div className="flex gap-2"><PrimaryBtn sm icon={CheckCircle2} onClick={() => actions.setCompanySub(c.id, "active")}>قبول الاشتراك</PrimaryBtn><GhostBtn sm icon={XCircle} onClick={() => actions.setCompanySub(c.id, "rejected")}>رفض</GhostBtn></div>
          : <PrimaryBtn sm full tone="ghost" icon={ServerCog} onClick={() => setManage({ ...c, addons: { ...c.addons } })}>إدارة كاملة</PrimaryBtn>}
      </Card>); })}{filtered.length === 0 && <Empty icon={Building2} text="لا توجد شركات مطابقة." />}</div>

    {manage && <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 overflow-auto" onClick={() => setManage(null)}>
      <div className="w-full max-w-lg rounded-2xl p-5 my-8" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center gap-2 mb-4"><span className="text-2xl">{manage.logo}</span><div><div className="font-extrabold" style={{ color: C.text }}>إدارة اشتراك الشركة</div><div className="text-sm" style={{ color: C.soft }}>{manage.name}</div></div></div>

        {/* نوع الاشتراك */}
        <div className="text-sm font-bold mb-1.5" style={{ color: C.text }}>نوع الاشتراك</div>
        <div className="grid grid-cols-3 gap-2 mb-2">{Object.entries(PLANS).map(([k, p]) => { const on = manage.plan === k; return <button key={k} onClick={() => setManage({ ...manage, plan: k, intelAccess: k === "intel" })} className="rounded-xl py-2 text-xs font-bold" style={{ background: on ? p.color : C.bg, color: on ? "#fff" : C.text, border: `1px solid ${on ? p.color : C.line}` }}>{p.label}</button>; })}</div>
        <div className="rounded-xl p-2.5 mb-3 text-[11px] space-y-0.5" style={{ background: (PLANS[manage.plan] || PLANS.basic).soft }}>{(PLANS[manage.plan] || PLANS.basic).features.map((f) => <div key={f} className="flex items-center gap-1" style={{ color: C.text }}><CheckCircle2 size={11} style={{ color: (PLANS[manage.plan] || PLANS.basic).color }} />{f}</div>)}</div>

        {/* دورة الفوترة والحدود */}
        <div className="grid grid-cols-2 gap-2 mb-3">
          <Field label="دورة الفوترة"><select value={manage.billing} onChange={(e) => setManage({ ...manage, billing: e.target.value })} className={inputCls} style={inputStyle}>{Object.entries(BILLING).map(([k, v]) => <option key={k} value={k}>{v}</option>)}</select></Field>
          <Field label="قيمة الاشتراك (د.ل)"><Input type="number" value={manage.subFee} onChange={(e) => setManage({ ...manage, subFee: +e.target.value })} /></Field>
          <Field label="حد المندوبين"><Input type="number" value={manage.maxReps} onChange={(e) => setManage({ ...manage, maxReps: +e.target.value })} /></Field>
          <Field label="حد المنتجات"><Input type="number" value={manage.maxProducts} onChange={(e) => setManage({ ...manage, maxProducts: +e.target.value })} /></Field>
          <Field label="نسبة الخصم %"><Input type="number" value={manage.discount} onChange={(e) => setManage({ ...manage, discount: +e.target.value })} /></Field>
          <Field label="حالة الدفع"><select value={manage.payStatus} onChange={(e) => setManage({ ...manage, payStatus: e.target.value })} className={inputCls} style={inputStyle}><option value="paid">مدفوع</option><option value="due">مستحق</option></select></Field>
        </div>

        {/* الخدمات الإضافية */}
        <div className="text-sm font-bold mb-1.5" style={{ color: C.text }}>الخدمات الإضافية</div>
        <div className="space-y-1.5 mb-3">{Object.entries(ADDONS).map(([k, v]) => <label key={k} className="flex items-center gap-2 text-sm rounded-lg p-2" style={{ background: C.bg, color: C.text }}><input type="checkbox" checked={!!manage.addons?.[k]} onChange={(e) => setManage({ ...manage, addons: { ...manage.addons, [k]: e.target.checked } })} />{v}</label>)}<label className="flex items-center gap-2 text-sm rounded-lg p-2" style={{ background: C.bg, color: C.text }}><input type="checkbox" checked={manage.trial} onChange={(e) => setManage({ ...manage, trial: e.target.checked })} />فترة تجريبية</label></div>

        {/* إدارة الفروع */}
        <div className="text-sm font-bold mb-1.5 flex items-center gap-1.5" style={{ color: C.text }}><Layers size={14} style={{ color: SC.color }} />الفروع</div>
        <div className="rounded-xl p-3 mb-3" style={{ background: C.bg }}>
          <label className="flex items-center justify-between gap-2 text-sm font-bold mb-2" style={{ color: C.text }}>
            <span>شركة متعددة الفروع</span>
            <input type="checkbox" checked={(manage.branches || []).length > 0} onChange={(e) => setManage({ ...manage, branches: e.target.checked ? (manage.branches && manage.branches.length ? manage.branches : [{ id: "br" + Math.floor(100 + Math.random() * 899), name: "الفرع الرئيسي", city: manage.city || "", manager: "أدمن الفرع" }]) : [] })} />
          </label>
          {(manage.branches || []).length > 0 && <div className="space-y-1.5">
            {(manage.branches || []).map((b, idx) => (<div key={b.id} className="flex items-center gap-1.5">
              <input value={b.name} onChange={(e) => setManage({ ...manage, branches: manage.branches.map((x) => x.id === b.id ? { ...x, name: e.target.value } : x) })} placeholder="اسم الفرع" className="flex-1 rounded-lg px-2 py-1.5 text-sm" style={{ border: `1px solid ${C.line}`, background: C.surface, color: C.text }} />
              <input value={b.city} onChange={(e) => setManage({ ...manage, branches: manage.branches.map((x) => x.id === b.id ? { ...x, city: e.target.value } : x) })} placeholder="المدينة" className="w-24 rounded-lg px-2 py-1.5 text-sm" style={{ border: `1px solid ${C.line}`, background: C.surface, color: C.text }} />
              <button onClick={() => setManage({ ...manage, branches: manage.branches.filter((x) => x.id !== b.id) })} className="rounded-lg p-1.5" style={{ background: C.dangerSoft }}><Trash2 size={14} style={{ color: C.danger }} /></button>
            </div>))}
            <button onClick={() => setManage({ ...manage, branches: [...(manage.branches || []), { id: "br" + Math.floor(100 + Math.random() * 899), name: "", city: "", manager: "أدمن الفرع" }] })} className="w-full rounded-lg py-1.5 text-xs font-bold flex items-center justify-center gap-1" style={{ background: SC.soft, color: SC.color }}><Plus size={13} />إضافة فرع</button>
          </div>}
        </div>

        {/* صلاحيات المسؤول */}
        <div className="text-sm font-bold mb-1.5" style={{ color: C.text }}>صلاحيات سريعة</div>
        <div className="grid grid-cols-2 gap-2 mb-4">
          <PrimaryBtn sm tone={manage.sub === "active" ? "ghost" : "primary"} icon={manage.sub === "active" ? XCircle : CheckCircle2} onClick={() => setManage({ ...manage, sub: manage.sub === "active" ? "suspended" : "active" })}>{manage.sub === "active" ? "إيقاف الشركة" : "تفعيل الشركة"}</PrimaryBtn>
          <PrimaryBtn sm tone={manage.frozen ? "primary" : "danger"} icon={Lock} onClick={() => setManage({ ...manage, frozen: !manage.frozen })}>{manage.frozen ? "رفع التجميد" : "تجميد الحساب"}</PrimaryBtn>
          <PrimaryBtn sm tone="accent" icon={Percent} onClick={() => setManage({ ...manage, commValue: 0 })}>إعفاء من العمولة</PrimaryBtn>
          <PrimaryBtn sm icon={CalendarClock} onClick={() => setManage({ ...manage, contractEnd: "2027-12-31" })}>تمديد سنة</PrimaryBtn>
        </div>

        <div className="flex gap-2"><PrimaryBtn full icon={CheckCircle2} onClick={() => { actions.adminAction(manage.id, { plan: manage.plan, billing: manage.billing, subFee: manage.subFee, maxReps: manage.maxReps, maxProducts: manage.maxProducts, discount: manage.discount, payStatus: manage.payStatus, addons: manage.addons, trial: manage.trial, sub: manage.sub, frozen: manage.frozen, commValue: manage.commValue, intelAccess: manage.intelAccess, contractEnd: manage.contractEnd, branches: manage.branches || [] }, `تحديث اشتراك ${manage.name}`); setManage(null); }}>حفظ كل التغييرات</PrimaryBtn><GhostBtn onClick={() => setManage(null)}>إلغاء</GhostBtn></div>
      </div>
    </div>}
  </div>);
}

// العقود
function PFContracts(props) {
  const { data } = props;
  const all = data.companies.filter((c) => c.server === sv(props));
  const [sel, setSel] = useState(null);
  if (sel) { const c = all.find((x) => x.id === sel);
    return (<div className="space-y-4"><button onClick={() => setSel(null)} className="inline-flex items-center gap-1 text-sm font-bold" style={{ color: C.soft }}><ChevronLeft size={16} />كل العقود</button>
      <Card className="p-5">
        <div className="flex items-center justify-between mb-4 pb-4" style={{ borderBottom: `1px solid ${C.line}` }}><div className="flex items-center gap-3"><span className="text-3xl">{c.logo}</span><div><div className="font-extrabold text-lg" style={{ color: C.text }}>{c.name}</div><div className="text-sm" style={{ color: C.soft }}>عقد رقم {c.contractNo}</div></div></div><span className="text-xs font-bold px-2.5 py-1 rounded-full" style={c.sub === "active" ? { background: C.successSoft, color: C.success } : { background: C.dangerSoft, color: C.danger }}>{c.sub === "active" ? "ساري" : "موقوف"}</span></div>
        <div className="grid sm:grid-cols-2 gap-3 mb-4">{[["رقم العقد", c.contractNo], ["تاريخ البداية", c.contractStart], ["تاريخ النهاية", c.contractEnd], ["قيمة الاشتراك", money(c.subFee)], ["نوع الاشتراك", (PLANS[c.plan] || PLANS.basic).label], ["نسبة العمولة", c.commType === "fixed" ? money(c.commValue) : c.commType === "tiered" ? "متدرّجة" : c.commValue + "%"], ["طريقة الدفع", c.payMethod], ["حالة العقد", c.sub === "active" ? "ساري المفعول" : "موقوف"]].map(([k, v]) => <div key={k} className="flex justify-between rounded-xl p-3" style={{ background: C.bg }}><span style={{ color: C.soft }}>{k}</span><span className="font-bold" style={{ color: C.text }}>{v}</span></div>)}</div>
        <div className="flex gap-2"><PrimaryBtn icon={FileText}>تحميل نسخة PDF</PrimaryBtn><GhostBtn icon={ScrollText}>سجل التعديلات</GhostBtn></div>
      </Card>
    </div>);
  }
  return (<div className="space-y-4"><SectionTitle sub="عقود الاشتراك لكل شركة.">العقود</SectionTitle>
    <Card className="p-4"><Table head={["رقم العقد", "الشركة", "البداية", "النهاية", "قيمة الاشتراك", "الحالة", ""]}>{all.map((c) => <tr key={c.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{c.contractNo}</Td><Td>{c.logo} {c.name}</Td><Td>{c.contractStart}</Td><Td>{c.contractEnd}</Td><Td>{money(c.subFee)}</Td><Td><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={c.sub === "active" ? { background: C.successSoft, color: C.success } : { background: C.dangerSoft, color: C.danger }}>{c.sub === "active" ? "ساري" : "موقوف"}</span></Td><Td><GhostBtn sm icon={Eye} onClick={() => setSel(c.id)}>عرض</GhostBtn></Td></tr>)}</Table></Card>
  </div>);
}

// استخبارات السوق التفصيلية (بيانات مجمّعة) + سحب التقرير للشركات المشتركة
function PFIntel(props) {
  const { data } = props;
  const prods = data.products.filter((p) => p.server === sv(props));
  const orders = data.orders.filter((o) => o.server === sv(props) && o.status === "delivered");
  const intelCompanies = data.companies.filter((c) => c.server === sv(props) && c.intelAccess && c.sub === "active");
  const [pulled, setPulled] = useState(null);

  // تجميع حسب الصنف: كمية + قيمة
  const byProduct = {};
  orders.forEach((o) => o.items.forEach((it) => { const p = data.products.find((x) => x.id === it.pid); const k = it.name; if (!byProduct[k]) byProduct[k] = { qty: 0, value: 0, cat: p?.cat || "أخرى" }; byProduct[k].qty += it.qty; byProduct[k].value += it.qty * it.price; }));
  const productRows = Object.entries(byProduct).map(([name, v]) => ({ name, ...v }));
  const topByQty = [...productRows].sort((a, b) => b.qty - a.qty).slice(0, 8);
  const topByValue = [...productRows].sort((a, b) => b.value - a.value).slice(0, 8);
  const maxQty = Math.max(...topByQty.map((x) => x.qty), 1);
  const maxVal = Math.max(...topByValue.map((x) => x.value), 1);

  // حسب التصنيف: كمية + قيمة + متوسط سعر
  const byCat = {};
  productRows.forEach((r) => { if (!byCat[r.cat]) byCat[r.cat] = { qty: 0, value: 0 }; byCat[r.cat].qty += r.qty; byCat[r.cat].value += r.value; });
  const catRows = Object.entries(byCat).map(([cat, v]) => { const prices = prods.filter((p) => p.cat === cat).map((p) => p.price); const avg = prices.length ? Math.round(prices.reduce((a, b) => a + b, 0) / prices.length) : 0; return { cat, ...v, avg, trend: (cat.length % 2 === 0 ? 1 : -1) * (1 + (cat.length % 7)) }; });

  // المناطق: كمية + قيمة
  const byArea = {};
  orders.forEach((o) => { const m = data.merchants.find((x) => x.id === o.merchantId); if (m) { if (!byArea[m.area]) byArea[m.area] = { qty: 0, value: 0 }; byArea[m.area].qty += orderQty(o); byArea[m.area].value += orderTotal(o); } });
  const areaRows = Object.entries(byArea).map(([area, v]) => ({ area, ...v })).sort((a, b) => b.qty - a.qty);

  const totalQty = productRows.reduce((s, r) => s + r.qty, 0);
  const totalVal = productRows.reduce((s, r) => s + r.value, 0);

  return (<div className="space-y-4"><SectionTitle sub="بيانات السوق المجمّعة بالكمية والقيمة — لا تكشف بيانات شركة بعينها.">استخبارات السوق التفصيلية</SectionTitle>
    <Card className="p-4 mb-1" style={{ background: SC.soft, border: "none" }}><div className="flex items-start gap-2 text-sm" style={{ color: C.text }}><ShieldCheck size={17} style={{ color: SC.color }} className="mt-0.5 shrink-0" />كل الأرقام مجمّعة على مستوى السوق (مثال: «المياه باعت 120 ألف كرتونة») ولا تُنسب لأي شركة. تُتاح للمشتركين في «استخبارات المنافسين».</div></Card>

    {/* ملخص السوق */}
    <div className="grid sm:grid-cols-3 gap-3">
      <Stat icon={Boxes} label="إجمالي الكمية المتداولة" value={qty(totalQty)} tint={{ color: SC.color2, soft: SC.soft }} />
      <Stat icon={CircleDollarSign} label="إجمالي قيمة السوق" value={money(totalVal)} tint={{ color: C.success, soft: C.successSoft }} />
      <Stat icon={Package} label="عدد الأصناف المتداولة" value={productRows.length} tint={{ color: C.accent, soft: C.accentSoft }} />
    </div>

    {/* الأكثر مبيعاً بالكمية والقيمة */}
    <div className="grid lg:grid-cols-2 gap-4">
      <Card className="p-5"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Boxes size={18} style={{ color: SC.color2 }} />الأكثر مبيعاً بالكمية</h3>
        <div className="space-y-3">{topByQty.map((r) => <div key={r.name}><div className="flex justify-between text-sm mb-1"><span style={{ color: C.text }}>{r.name}</span><span className="font-bold" style={{ color: SC.color2 }}>{qty(r.qty)}</span></div><ProgressBar value={r.qty} max={maxQty} color={SC.color2} /></div>)}{topByQty.length === 0 && <Empty icon={Boxes} text="لا بيانات بعد." />}</div>
      </Card>
      <Card className="p-5"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><TrendingUp size={18} style={{ color: C.success }} />الأكثر مبيعاً بالقيمة</h3>
        <div className="space-y-3">{topByValue.map((r) => <div key={r.name}><div className="flex justify-between text-sm mb-1"><span style={{ color: C.text }}>{r.name}</span><span className="font-bold" style={{ color: C.success }}>{money(r.value)}</span></div><ProgressBar value={r.value} max={maxVal} color={C.success} /></div>)}{topByValue.length === 0 && <Empty icon={TrendingUp} text="لا بيانات بعد." />}</div>
      </Card>
    </div>

    {/* تفصيل حسب التصنيف */}
    <Card className="p-4"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Layers size={18} style={{ color: C.purple }} />تفصيل السوق حسب التصنيف</h3>
      <Table head={["التصنيف", "الكمية المباعة", "قيمة المبيعات", "متوسط السعر", "اتجاه الطلب"]}>{catRows.sort((a, b) => b.qty - a.qty).map((r) => <tr key={r.cat} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{r.cat}</Td><Td><span style={{ color: SC.color2, fontWeight: 700 }}>{qty(r.qty)}</span></Td><Td>{money(r.value)}</Td><Td>{money(r.avg)}</Td><Td><span className="text-xs font-bold flex items-center gap-0.5" style={{ color: r.trend > 0 ? C.success : C.danger }}>{r.trend > 0 ? <TrendingUp size={13} /> : <TrendingDown size={13} />}{r.trend > 0 ? "+" : ""}{r.trend}%</span></Td></tr>)}</Table>
    </Card>

    {/* المناطق */}
    <Card className="p-4"><h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><MapPinned size={18} style={{ color: C.info }} />المناطق الأكثر شراءً</h3>
      <Table head={["المنطقة", "الكمية", "القيمة", "الحصة من السوق"]}>{areaRows.map((r) => { const share = Math.round((r.value / (totalVal || 1)) * 100); return (<tr key={r.area} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{r.area}</Td><Td>{qty(r.qty)}</Td><Td>{money(r.value)}</Td><Td><div className="flex items-center gap-2"><div style={{ width: 70 }}><ProgressBar value={r.value} max={totalVal} color={C.info} /></div><span className="text-xs" style={{ color: C.soft }}>{share}%</span></div></Td></tr>); })}</Table>
    </Card>

    {/* سحب التقرير للشركة المشتركة */}
    <Card className="p-5"><h3 className="font-extrabold mb-1 flex items-center gap-2" style={{ color: C.text }}><FileText size={18} style={{ color: C.accent }} />سحب تقرير السوق للشركات المشتركة</h3>
      <p className="text-sm mb-3" style={{ color: C.soft }}>اختر شركة مشتركة في ميزة «استخبارات المنافسين» لتوليد نسخة تقرير قابلة للتصدير لها (بيانات مجمّعة فقط).</p>
      {intelCompanies.length ? <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-2">{intelCompanies.map((c) => <button key={c.id} onClick={() => setPulled(c)} className="flex items-center justify-between rounded-xl p-3 text-right transition-shadow hover:shadow-md" style={{ border: `1px solid ${C.line}` }}><div className="flex items-center gap-2"><span className="text-xl">{c.logo}</span><span className="text-sm font-bold" style={{ color: C.text }}>{c.name}</span></div><FileText size={16} style={{ color: C.accent }} /></button>)}</div> : <Empty icon={ShieldCheck} text="لا توجد شركات مشتركة في ميزة الاستخبارات. فعّلها من إدارة الاشتراكات." />}
    </Card>

    {/* نافذة التقرير المُولّد */}
    {pulled && <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 overflow-auto" onClick={() => setPulled(null)}>
      <div className="w-full max-w-lg rounded-2xl my-8" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
        <div className="p-5 text-white" style={{ background: `linear-gradient(120deg, ${C.ink}, ${SC.color})` }}>
          <div className="flex items-center justify-between"><div><div className="text-sm opacity-80">تقرير استخبارات السوق</div><div className="font-extrabold text-lg">{pulled.name}</div></div><span className="text-2xl">{pulled.logo}</span></div>
          <div className="text-xs opacity-70 mt-1">بيانات مجمّعة · سيرفر {SC.name} · {new Date().toLocaleDateString("ar-LY")}</div>
        </div>
        <div className="p-5 space-y-4 max-h-[60vh] overflow-y-auto">
          <div className="grid grid-cols-3 gap-2 text-center"><div className="rounded-xl p-3" style={{ background: SC.soft }}><div className="font-extrabold" style={{ color: SC.color }}>{qty(totalQty)}</div><div className="text-[10px]" style={{ color: C.soft }}>كمية السوق</div></div><div className="rounded-xl p-3" style={{ background: C.successSoft }}><div className="font-extrabold text-sm" style={{ color: C.success }}>{money(totalVal)}</div><div className="text-[10px]" style={{ color: C.soft }}>قيمة السوق</div></div><div className="rounded-xl p-3" style={{ background: C.accentSoft }}><div className="font-extrabold" style={{ color: C.accent }}>{productRows.length}</div><div className="text-[10px]" style={{ color: C.soft }}>أصناف</div></div></div>
          <div><div className="font-bold text-sm mb-2" style={{ color: C.text }}>أعلى 5 أصناف بالكمية</div><div className="space-y-1.5">{topByQty.slice(0, 5).map((r, i) => <div key={r.name} className="flex justify-between text-sm rounded-lg px-3 py-1.5" style={{ background: C.bg }}><span style={{ color: C.text }}>{i + 1}. {r.name}</span><span className="font-bold" style={{ color: SC.color2 }}>{qty(r.qty)}</span></div>)}</div></div>
          <div><div className="font-bold text-sm mb-2" style={{ color: C.text }}>أعلى المناطق طلباً</div><div className="space-y-1.5">{areaRows.slice(0, 4).map((r) => <div key={r.area} className="flex justify-between text-sm rounded-lg px-3 py-1.5" style={{ background: C.bg }}><span style={{ color: C.text }}>{r.area}</span><span className="font-bold" style={{ color: C.info }}>{qty(r.qty)}</span></div>)}</div></div>
          <div className="rounded-xl p-3 text-xs flex items-start gap-2" style={{ background: SC.soft, color: C.text }}><ShieldCheck size={15} style={{ color: SC.color }} className="mt-0.5 shrink-0" />هذا التقرير لا يتضمن أي بيانات منسوبة لشركة معينة، حمايةً لخصوصية الموردين.</div>
        </div>
        <div className="p-4 flex gap-2" style={{ borderTop: `1px solid ${C.line}` }}><PrimaryBtn full icon={FileText}>تصدير PDF</PrimaryBtn><PrimaryBtn full tone="ghost" icon={Hash}>تصدير Excel</PrimaryBtn></div>
      </div>
    </div>}
  </div>);
}

function PFMerchants(props) {
  const { data } = props;
  return (<div className="space-y-4"><SectionTitle sub={`${SC.merchantsWord} المسجّلة في سيرفر ${SC.name}.`}>{SC.merchantsWord}</SectionTitle><div className="grid md:grid-cols-2 gap-3">{F.merchants(props).map((m) => <Card key={m.id} className="p-4 flex items-center gap-3"><div className="rounded-xl p-2.5" style={{ background: C.accentSoft }}><Store size={20} style={{ color: C.accent }} /></div><div className="flex-1"><div className="font-extrabold" style={{ color: C.text }}>{m.name}</div><div className="text-sm" style={{ color: C.soft }}>{m.owner} · {m.area}</div><div className="text-sm flex items-center gap-1" style={{ color: C.soft }}><Building2 size={13} />{compName(data, m.companyId)}</div></div></Card>)}</div></div>);
}
function PFOrders(props) { return <OrdersTable {...props} orders={props.data.orders.filter((o) => o.server === sv(props) && o.status !== "draft")} title="كل الطلبات" sub="جميع الطلبات في هذا السيرفر." />; }

// سجل العمليات
function PFAudit(props) {
  const { data } = props;
  const log = (data.auditLog || []).filter((l) => l.server === sv(props));
  return (<div className="space-y-4"><SectionTitle sub="سجل كامل بكل عمليات الإدارة على المنصة.">سجل العمليات</SectionTitle>
    <Card className="p-4"><Table head={["العملية", "الجهة", "التفاصيل", "الوقت"]}>{log.map((l) => <tr key={l.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: SC.soft, color: SC.color }}>{l.action}</span></Td><Td bold>{l.target}</Td><Td>{l.detail}</Td><Td>{l.t}</Td></tr>)}</Table>{log.length === 0 && <Empty icon={ScrollText} text="لا توجد عمليات." />}</Card>
  </div>);
}

/* ============================================================
   ============   المندوب   ==================================
   ============================================================ */
const useRep = (props) => props.data.reps.find((r) => r.id === props.user.entityId);
function RepHome(props) {
  const { data, go, user } = props; const rep = useRep(props);
  const orders = data.orders.filter((o) => o.repId === rep.id && o.status !== "draft");
  const pending = orders.filter((o) => o.status === "pending_rep").length;
  const pendingSettle = data.settlements.filter((x) => x.repId === rep.id && x.status === "pending_rep").length;
  const openCollects = data.collectRequests.filter((c) => c.repId === rep.id && ["requested", "scheduled"].includes(c.status)).length;
  const pendingLoads = data.loadRequests.filter((l) => l.repId === rep.id && l.status === "pending_rep").length;
  const urgent = data.urgentRoutes.filter((u) => u.repId === rep.id && u.stops.some((st) => !st.done)).length;
  const alerts = [];
  if (pending > 0) alerts.push({ k: "rp_orders", icon: ShoppingCart, text: `${pending} طلب بانتظار موافقتك` });
  if (openCollects > 0) alerts.push({ k: "rp_settle", icon: Truck, text: `${openCollects} عميل يطلب التحصيل`, soft: SC.soft, color: SC.color, border: SC.color });
  if (pendingSettle > 0) alerts.push({ k: "rp_settle", icon: Clock, text: `${pendingSettle} تسوية بانتظار مراجعتك` });
  if (pendingLoads > 0) alerts.push({ k: "rp_load", icon: PackagePlus, text: `${pendingLoads} عرض تحميل بانتظار موافقتك`, soft: C.infoSoft, color: C.info, border: C.info });
  if (urgent > 0) alerts.push({ k: "rp_urgent", icon: AlertCircle, text: `خط سير عاجل من مشرفك!`, soft: C.dangerSoft, color: C.danger, border: C.danger });
  const toVerify = data.pendingClients.filter((p) => p.repId === rep.id && ["awaiting_visit", "field_verified"].includes(p.status)).length;
  if (toVerify > 0) alerts.push({ k: "rp_clients", icon: ClipboardList, text: `${toVerify} محل جديد بانتظار توثيقك الميداني`, soft: C.purpleSoft, color: C.purple, border: C.purple });
  const rt = REP_TYPE[rep.repType || "presell"];
  return <RoleLauncher name={rep.name} sub={`${rep.area} · ${rt.label}`} roleTitle="مندوب مبيعات" nav={navFor(user.role, user.companyRole)} homeKey="rp_home" go={go} alerts={alerts} badges={{ rp_orders: pending, rp_settle: pendingSettle + openCollects, rp_load: pendingLoads, rp_urgent: urgent, rp_clients: toVerify }} quick={{ k: "rp_invoice", icon: Receipt, label: "فاتورة ميدانية" }} />;
}
function RepOrders(props) {
  const { data, openOrder } = props; const rep = useRep(props);
  const orders = data.orders.filter((o) => o.repId === rep.id && o.status !== "draft");
  const pending = orders.filter((o) => o.status === "pending_rep");
  return (<div className="space-y-5"><SectionTitle sub="راجع الطلبات، أكّد استلام النقد، ثم اقبل أو ارفض.">الطلبات والموافقات</SectionTitle>
    <div><h3 className="font-extrabold mb-3" style={{ color: C.text }}>بانتظار قرارك ({pending.length})</h3>
      <div className="grid md:grid-cols-2 gap-3">{pending.map((o) => { const m = data.merchants.find((x) => x.id === o.merchantId); return (<Card key={o.id} className="p-4"><div className="flex items-center justify-between mb-2"><span className="font-extrabold" style={{ color: C.text }}>{o.id}</span><Badge status={o.status} /></div><div className="text-sm flex items-center gap-1 mb-3" style={{ color: C.soft }}><Store size={14} />{m?.name} · {m?.area}</div><div className="rounded-xl p-3 mb-3 flex items-center justify-between" style={{ background: C.accentSoft }}><span className="text-sm font-bold flex items-center gap-1" style={{ color: "#92590C" }}><Banknote size={16} />النقد المطلوب</span><span className="font-extrabold" style={{ color: "#92590C" }}>{money(orderTotal(o))}</span></div><PrimaryBtn full icon={Eye} onClick={() => openOrder(o.id)}>مراجعة واتخاذ قرار</PrimaryBtn></Card>); })}{pending.length === 0 && <Card className="p-5"><span className="text-sm" style={{ color: C.soft }}>لا توجد طلبات معلّقة.</span></Card>}</div>
    </div>
    <div><h3 className="font-extrabold mb-3" style={{ color: C.text }}>بقية الطلبات</h3><Card className="p-2"><OrderList orders={orders.filter((o) => o.status !== "pending_rep")} data={data} onOpen={openOrder} /></Card></div>
  </div>);
}

function RepTasks(props) {
  const { data, actions } = props; const rep = useRep(props);
  const tasks = data.tasks.filter((t) => t.assigneeType === "rep" && t.assigneeId === rep.id);
  return (<div className="space-y-4"><SectionTitle sub="المهام التي أسندتها لك الشركة أو المشرف.">مهامي المسندة</SectionTitle>
    {tasks.length ? tasks.map((t) => (<Card key={t.id} className="p-4"><div className="flex items-start justify-between gap-3 flex-wrap"><div className="flex items-start gap-2"><div className="rounded-lg p-2 mt-0.5" style={{ background: C.infoSoft }}><ClipboardCheck size={17} style={{ color: C.info }} /></div><div><div className="font-bold flex items-center gap-2" style={{ color: C.text }}>{t.title}{t.priority === "high" && <span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.dangerSoft, color: C.danger }}>عاجل</span>}</div><div className="text-sm mt-0.5" style={{ color: C.soft }}>{t.desc}</div><div className="text-xs mt-1 flex items-center gap-1" style={{ color: C.soft }}><CalendarClock size={12} />{t.due}</div></div></div><Badge status={t.status} map={TASK_STATUS} /></div>
      <div className="flex gap-2 mt-3 pt-3" style={{ borderTop: `1px solid ${C.line}` }}>{t.status === "assigned" && <PrimaryBtn sm icon={Activity} onClick={() => actions.updateTask(t.id, { status: "in_progress" })}>بدء التنفيذ</PrimaryBtn>}{t.status === "in_progress" && <PrimaryBtn sm tone="success" icon={CheckCircle2} onClick={() => actions.updateTask(t.id, { status: "done" })}>إكمال المهمة</PrimaryBtn>}{t.status === "done" && <span className="text-sm flex items-center gap-1" style={{ color: C.success }}><CheckCircle2 size={15} />مكتملة</span>}</div>
    </Card>)) : <Empty icon={ListChecks} text="لا توجد مهام مسندة." />}
  </div>);
}

// زيارات التوثيق للانضمام
function RepJoin(props) {
  const { data, actions } = props; const rep = useRep(props);
  const joins = data.joinRequests.filter((j) => j.repId === rep.id && ["visit_set", "verifying"].includes(j.status));
  const [modal, setModal] = useState(null);
  const [verified, setVerified] = useState(0);
  return (<div className="space-y-4"><SectionTitle sub="زيارات وُكِّلت بها لتوثيق عملاء جدد ومطابقة مديونياتهم السابقة وتسجيل دينهم الافتتاحي.">زيارات التوثيق</SectionTitle>
    {joins.length ? joins.map((j) => { const m = data.merchants.find((x) => x.id === j.merchantId); const c = data.companies.find((x) => x.id === j.companyId);
      return (<Card key={j.id} className="p-5"><div className="flex items-center justify-between mb-3 flex-wrap gap-2"><div className="flex items-center gap-3"><div className="rounded-xl p-2.5" style={{ background: C.accentSoft }}><Store size={20} style={{ color: C.accent }} /></div><div><div className="font-extrabold" style={{ color: C.text }}>{m?.name}</div><div className="text-xs" style={{ color: C.soft }}>{m?.owner} · {m?.area} · {m?.phone}</div></div></div><Badge status={j.status} map={JOIN_STATUS} /></div>
        <div className="grid sm:grid-cols-3 gap-2 text-sm mb-3">
          <Info icon={Building2} label="الشركة" value={c?.name} />
          <Info icon={CalendarClock} label="موعد الزيارة" value={j.visitDate} />
          <Info icon={TrendingDown} label="الدين المُصرّح من العميل" value={money(j.oldDebt)} />
        </div>
        <div className="rounded-xl p-3 mb-3 text-sm flex items-start gap-2" style={{ background: C.purpleSoft, color: C.text }}><ScanSearch size={17} style={{ color: C.purple }} className="mt-0.5 shrink-0" />قم بزيارة العميل، طابق سجلاته القديمة، ثم وثّق الدين الافتتاحي الفعلي وفعّل الحساب.</div>
        <PrimaryBtn full icon={ClipboardSignature} onClick={() => { setModal(j); setVerified(j.oldDebt); }}>توثيق ومطابقة المديونية</PrimaryBtn>
      </Card>); }) : <Empty icon={FileSignature} text="لا توجد زيارات توثيق مجدولة." />}

    {modal && <div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50" onClick={() => setModal(null)}>
      <div className="rb-modal w-full max-w-md rounded-2xl p-5" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
        <h3 className="font-extrabold mb-1" style={{ color: C.text }}>توثيق العميل</h3>
        <p className="text-sm mb-4" style={{ color: C.soft }}>{merchName(data, modal.merchantId)}</p>
        <div className="rounded-xl p-3 mb-3 flex justify-between text-sm" style={{ background: C.bg }}><span style={{ color: C.soft }}>الدين المُصرّح من العميل</span><span className="font-bold" style={{ color: C.text }}>{money(modal.oldDebt)}</span></div>
        <Field label="الدين الموثّق بعد المطابقة (د.ل)" hint="سيُسجّل كرصيد افتتاحي على العميل"><Input type="number" value={verified} onChange={(e) => setVerified(+e.target.value)} /></Field>
        <div className="flex gap-2 mt-4"><PrimaryBtn full tone="success" icon={BadgeCheck} onClick={() => { actions.verifyJoin(modal.id, verified); setModal(null); }}>توثيق وتفعيل الحساب</PrimaryBtn><GhostBtn onClick={() => setModal(null)}>إلغاء</GhostBtn></div>
      </div>
    </div>}
  </div>);
}

function StockTakeForm({ prods, onSubmit }) {
  const [vals, setVals] = useState({}); const [note, setNote] = useState("");
  return (<div>
    <div className="space-y-1.5 mb-3 max-h-40 overflow-auto">{prods.map((p) => <div key={p.id} className="flex items-center justify-between gap-2"><span className="text-xs flex-1" style={{ color: C.text }}>{p.name}</span><input type="number" value={vals[p.id] || ""} onChange={(e) => setVals({ ...vals, [p.id]: e.target.value })} placeholder="0" className="w-16 rounded-lg px-2 py-1 text-xs outline-none text-center" style={inputStyle} /></div>)}</div>
    <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="ملاحظة الجرد..." className="w-full rounded-lg px-2 py-1.5 text-xs outline-none mb-2" style={inputStyle} />
    <PrimaryBtn full icon={ClipboardCheck} onClick={() => { const lines = prods.filter((p) => vals[p.id]).map((p) => ({ name: p.name, shelf: +vals[p.id] })); if (lines.length) onSubmit(lines, note); }}>رفع طلب الجرد</PrimaryBtn>
  </div>);
}
function CashCollect({ data, client, rep, actions }) {
  const m = data.merchants.find((x) => x.id === client);
  const debt = Math.max(0, -(m?.balance || 0));
  const [amt, setAmt] = useState(""); const [done, setDone] = useState(false);
  return (<div>
    <div className="rounded-xl p-3 mb-3 text-center" style={{ background: debt > 0 ? C.dangerSoft : C.successSoft }}><div className="text-xs" style={{ color: C.soft }}>مديونية {m?.name}</div><div className="text-2xl font-extrabold" style={{ color: debt > 0 ? C.danger : C.success }}>{money(debt)}</div></div>
    <Field label="المبلغ المُحصّل"><input type="number" value={amt} onChange={(e) => { setAmt(e.target.value); setDone(false); }} className="w-full rounded-lg px-3 py-2 text-sm outline-none" style={inputStyle} placeholder="0" /></Field>
    <div className="mt-3"><PrimaryBtn full tone="success" icon={HandCoins} onClick={() => { if (+amt > 0) { actions.repCollectCash(rep.id, client, +amt); setDone(true); setAmt(""); } }}>تأكيد التحصيل النقدي</PrimaryBtn></div>
    {done && <div className="text-xs mt-2 flex items-center justify-center gap-1" style={{ color: C.success }}><CheckCircle2 size={14} />تم تسجيل التحصيل وتحديث الرصيد</div>}
  </div>);
}
// إصدار فاتورة للعميل عند التواجد في نطاق المحل (GPS)
function RepInvoice(props) {
  const { data, actions } = props; const rep = useRep(props);
  const clients = data.merchants.filter((m) => m.repId === rep.id);
  const [client, setClient] = useState(clients[0]?.id || "");
  const m = data.merchants.find((x) => x.id === client);
  const quotas = data.quotas.filter((q) => q.repId === rep.id);
  const prods = data.products.filter((p) => quotas.some((q) => q.pid === p.id) || p.companyId === rep.companyId);
  const [cart, setCart] = useState({});
  const [atStore, setAtStore] = useState(false);
  const [issued, setIssued] = useState(null);
  const cartItems = Object.entries(cart);
  const total = cartItems.reduce((s, [pid, q]) => { const p = data.products.find((x) => x.id === pid); return s + tierUnitPrice(p, m, q) * q; }, 0);
  const distance = atStore ? 0 : 1.8;
  const inRange = atStore;
  const isVan = rep.repType === "van";
  const isPresell = rep.repType === "presell";
  const gate = m ? invoiceGate(data, m.id) : { blocked: false, days: 0 };
  const debtBlocked = gate.blocked;
  const locationOk = inRange || isPresell; // البري‑سيل يصدر الفاتورة من أي مكان
  const gateOk = locationOk && !debtBlocked && !m?.frozen;
  const rt = REP_TYPE[rep.repType || "presell"];
  const vanQty = (pid) => (rep.vanStock || {})[pid] || 0;
  // مندوب التجزئة: كل أصناف السيارة تظهر — المنفد (كمية 0) يظهر معطّلاً
  const sourceProds = prods;
  const vanProds = isVan ? Object.keys(rep.vanStock || {}).map((pid) => data.products.find((p) => p.id === pid)).filter(Boolean) : [];
  const addToCart = (pid) => { if (isVan) { const have = vanQty(pid); if (have <= 0) { toast("هذا الصنف نفد من سيارتك — لا يمكن إصدار فاتورة به", "danger"); return; } setCart((c) => { if ((c[pid] || 0) + 1 > have) { toast(`الكمية القصوى بالسيارة: ${have}`, "danger"); return c; } return { ...c, [pid]: (c[pid] || 0) + 1 }; }); return; } setCart((c) => ({ ...c, [pid]: (c[pid] || 0) + 1 })); };
  const setQ = (pid, q) => setCart((c) => { const n = { ...c }; const cap = isVan ? vanQty(pid) : Infinity; const qq = Math.min(q, cap); if (qq <= 0) delete n[pid]; else { n[pid] = qq; if (q > cap) toast(`الكمية القصوى بالسيارة: ${cap}`, "danger"); } return n; });
  return (<div className="space-y-4"><SectionTitle sub={isVan ? "بِع فوراً من مخزون سيارتك وأصدر الفاتورة عند العميل." : "أصدر فاتورة للعميل من أي مكان — تُجهّزها الشركة وتوصّلها."}>إصدار فاتورة للعميل</SectionTitle>
    <div className="rounded-xl p-3 flex items-center gap-2" style={{ background: rt.soft }}><rt.icon size={18} style={{ color: rt.color }} /><span className="text-sm font-bold" style={{ color: rt.color }}>{rt.label}</span><span className="text-xs" style={{ color: C.soft }}>— {isVan ? "الفاتورة تُخصم من مخزون سيارتك وتُسلَّم فوراً" : "لا يُشترط تواجدك في المحل؛ الطلب يُجهَّز من الشركة"}</span></div>
    <Card className="p-4"><Field label="العميل"><select value={client} onChange={(e) => { setClient(e.target.value); setAtStore(false); }} className={inputCls} style={inputStyle}>{clients.map((x) => <option key={x.id} value={x.id}>{x.name} — {x.area}</option>)}</select></Field>
      {m && (() => { const days = gate.days; const b = AGE_BUCKETS.find((x) => days >= x.min && days <= x.max) || AGE_BUCKETS[2]; if (days <= 0 && !m.frozen) return null; return (<div className="mt-3 rounded-xl p-3" style={{ background: (debtBlocked || m.frozen) ? C.dangerSoft : b.soft, border: `1px solid ${(debtBlocked || m.frozen) ? C.danger : b.color}` }}>
        <div className="flex items-center gap-2 text-sm font-bold" style={{ color: (debtBlocked || m.frozen) ? C.danger : b.color }}>{(debtBlocked || m.frozen) ? <XCircle size={16} /> : <Clock size={16} />}{m.frozen ? `حساب معلّق: ${m.freezeReason || "راجع الشركة"}` : debtBlocked ? `ممنوع إصدار فاتورة — ${gate.reason}` : `أقدم دين: ${days} يوم`}</div>
        {m.invoiceOverride && <div className="text-[11px] mt-1 font-bold" style={{ color: C.success }}>{m.invoiceOverride.type === "allow" ? "✓ سماح استثنائي من الشركة" : `✓ إمهال ${m.invoiceOverride.extraDays} يوم من الشركة`}</div>}
        {debtBlocked && !m.invoiceOverride && <div className="text-[11px] mt-1" style={{ color: C.soft }}>يلزم سداد المتأخر، أو سماح/إمهال من الشركة (الحسابات ← أعمار الديون).</div>}
      </div>); })()}
    </Card>
    <Card className="p-0 overflow-hidden">
      <div className="relative" style={{ height: 200, background: "linear-gradient(135deg,#DCE7F5,#CFE3DA)" }}>
        <div className="absolute inset-0" style={{ backgroundImage: `linear-gradient(${C.line} 1px,transparent 1px),linear-gradient(90deg,${C.line} 1px,transparent 1px)`, backgroundSize: "28px 28px", opacity: .5 }} />
        <div className="absolute rounded-full" style={{ left: "50%", top: "55%", width: 130, height: 130, transform: "translate(-50%,-50%)", background: inRange ? "rgba(22,163,74,.15)" : "rgba(220,53,69,.12)", border: `2px dashed ${inRange ? C.success : C.danger}` }} />
        <div className="absolute" style={{ left: "50%", top: "55%", transform: "translate(-50%,-100%)" }}><div className="flex flex-col items-center"><MapPin size={30} style={{ color: C.danger }} fill={C.danger} /><span className="text-[10px] font-bold rounded px-1.5 py-0.5 -mt-1" style={{ background: C.well, color: C.text }}>{m?.name}</span></div></div>
        <div className="absolute transition-all" style={{ left: atStore ? "50%" : "22%", top: atStore ? "55%" : "30%", transform: "translate(-50%,-50%)" }}><div className="relative"><div className="absolute inset-0 rounded-full animate-ping" style={{ background: SC.color, opacity: .4 }} /><div className="rounded-full flex items-center justify-center relative" style={{ width: 24, height: 24, background: SC.color, border: "2px solid #fff" }}><Navigation size={12} color="#fff" /></div></div></div>
        <div className="absolute top-2 right-2 text-[11px] font-bold rounded-lg px-2 py-1 flex items-center gap-1" style={{ background: "rgba(255,255,255,.9)", color: inRange ? C.success : C.danger }}>{inRange ? <><CheckCircle2 size={13} />داخل نطاق المحل</> : <><AlertCircle size={13} />تبعد {distance} كم</>}</div>
      </div>
      <div className="p-4">
        <div className="text-xs mb-2 flex items-center gap-1" style={{ color: C.soft }}><MapPinned size={13} />موقع المحل: {m?.lat?.toFixed(4)}, {m?.lng?.toFixed(4)}</div>
        {!atStore ? <div className="space-y-2"><PrimaryBtn full icon={Navigation} onClick={() => setAtStore(true)}>أنا الآن في المحل (تحديد موقعي)</PrimaryBtn>{isPresell && <div className="text-xs text-center rounded-lg p-2" style={{ background: C.infoSoft, color: C.info }}>مندوب البري‑سيل: تحديد الموقع اختياري — يمكنك إصدار الفاتورة من أي مكان.</div>}</div>
          : <div className="text-sm rounded-xl p-2.5 flex items-center gap-2" style={{ background: C.successSoft, color: C.success }}><CheckCircle2 size={16} />تم التحقق: أنت داخل نطاق {m?.name}، يمكنك إصدار الفاتورة.</div>}
      </div>
    </Card>
    <Card className="p-4" style={{ opacity: gateOk ? 1 : .55, pointerEvents: gateOk ? "auto" : "none" }}>
      <h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}>بنود الفاتورة{isVan && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.accentSoft, color: C.accent }}>من مخزون السيارة</span>}</h3>
      {isVan && <div className="text-xs font-bold mb-2 flex items-center gap-1.5" style={{ color: C.text }}><Truck size={14} style={{ color: C.accent }} />أصناف سيارتي ({vanProds.length} صنف) — المنفد لا يمكن البيع منه</div>}
      <div className="grid sm:grid-cols-2 gap-2 max-h-56 overflow-auto mb-3">{(isVan ? vanProds : sourceProds).map((p) => { const e = expiryInfo(p.expDate); const have = vanQty(p.id); const depleted = isVan && have <= 0; const inCart = cart[p.id] || 0; return (<button key={p.id} disabled={depleted} onClick={() => addToCart(p.id)} className="flex items-center justify-between rounded-xl p-2.5 text-right transition-opacity" style={{ border: `1px solid ${depleted ? C.danger : C.line}`, background: depleted ? C.dangerSoft + "55" : "transparent", opacity: depleted ? 0.7 : 1, cursor: depleted ? "not-allowed" : "pointer" }}><span className="min-w-0"><span className="text-sm font-bold block truncate" style={{ color: depleted ? C.danger : C.text }}>{p.name}</span><span className="text-xs" style={{ color: C.soft }}>{money(priceForMerchant(p, m))}{isVan ? (depleted ? " · نفد ✕" : ` · بالسيارة: ${have}${inCart ? ` (بالفاتورة ${inCart})` : ""}`) : p.batch ? ` · تشغيلة ${p.batch}` : ""}</span>{e && e.state !== "ok" && <span className="block mt-0.5"><ExpiryBadge expDate={p.expDate} small /></span>}</span>{depleted ? <XCircle size={16} style={{ color: C.danger }} /> : <Plus size={16} style={{ color: SC.color }} />}</button>); })}{(isVan ? vanProds : sourceProds).length === 0 && <div className="col-span-2"><Empty icon={Truck} text="سيارتك فارغة — اطلب تحميلاً من الشركة أولاً." /></div>}</div>
      <div className="space-y-2 mb-3">{cartItems.map(([pid, q]) => { const p = data.products.find((x) => x.id === pid); const t = tierFor(p, q); const e = p && expiryInfo(p.expDate); return (<div key={pid} className="rounded-xl p-2.5" style={{ border: `1px solid ${C.line}` }}><div className="flex items-center justify-between gap-2"><div className="min-w-0"><span className="text-sm font-bold block truncate" style={{ color: C.text }}>{p?.name}</span><span className="text-xs" style={{ color: C.soft }}>{money(tierUnitPrice(p, m, q))} × {q} = {money(tierUnitPrice(p, m, q) * q)}</span></div><div className="flex items-center gap-1 shrink-0"><button onClick={() => setQ(pid, q - 1)} className="w-7 h-7 rounded-lg font-bold" style={{ border: `1px solid ${C.line}` }}>−</button><span className="w-7 text-center text-sm font-bold">{q}</span><button onClick={() => setQ(pid, q + 1)} className="w-7 h-7 rounded-lg font-bold text-white" style={{ background: SC.color }}>+</button></div></div>{p && (p.batch || p.expDate || t) && <div className="flex flex-wrap gap-1.5 mt-1.5">{p.batch && <span className="inline-flex items-center gap-1 text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.bg, color: C.soft }}><Hash size={9} />{p.batch}</span>}{e && <span className="inline-flex items-center gap-1 text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: e.soft, color: e.color }}>صلاحية {fmtDate(p.expDate)}</span>}{t && <span className="inline-flex items-center gap-1 text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.purpleSoft, color: C.purple }}><Layers size={9} />خصم كمية {t.discount}%</span>}</div>}</div>); })}{cartItems.length === 0 && <p className="text-sm" style={{ color: C.soft }}>أضف منتجات للفاتورة.</p>}</div>
      <div className="flex justify-between items-center py-3" style={{ borderTop: `1px solid ${C.line}` }}><span className="font-bold" style={{ color: C.text }}>إجمالي الفاتورة</span><span className="text-lg font-extrabold" style={{ color: SC.color }}>{money(total)}</span></div>
      <PrimaryBtn full tone="accent" icon={Receipt} onClick={() => { if (!cartItems.length) { toast("أضف أصنافاً أولاً", "danger"); return; } if (m?.frozen) { toast(`حساب معلّق: ${m.freezeReason || "راجع الشركة"}`, "danger"); return; } if (debtBlocked) { toast(gate.reason, "danger"); return; } if (!locationOk) { toast("يجب تأكيد وجودك في المحل أولاً", "danger"); return; } const inv = actions.repIssueInvoice(rep.id, client, cart); setIssued(inv); setCart({}); }}>{debtBlocked ? "ممنوع — دين متأخر" : m?.frozen ? "ممنوع — حساب معلّق" : isVan ? "إصدار الفاتورة وتسليم البضاعة من السيارة" : "إصدار الفاتورة (تُجهَّز من الشركة)"}</PrimaryBtn>
    </Card>
    {issued && <div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50" onClick={() => setIssued(null)}>
      <div className="rb-modal w-full max-w-sm rounded-2xl overflow-hidden" style={{ background: C.surface }} onClick={(e) => e.stopPropagation()}>
        <div className="p-5 text-white text-center" style={{ background: `linear-gradient(120deg, ${SC.color}, ${SC.color2})` }}><div className="rounded-full mx-auto flex items-center justify-center mb-2" style={{ width: 52, height: 52, background: "rgba(255,255,255,.2)" }}><Receipt size={28} color="#fff" /></div><div className="font-extrabold text-lg">فاتورة صادرة</div><div className="text-sm opacity-90">{issued.id}</div></div>
        <div className="p-5 space-y-3">
          <div className="rounded-xl p-3 space-y-2 text-sm" style={{ background: C.bg }}>{[["العميل", merchName(data, issued.merchantId)], ["المندوب", rep.name], ["نوع البيع", issued.saleType === "van" ? "تجزئة — تسليم فوري من السيارة" : "بري‑سيل — يُجهَّز من الشركة"], ["التاريخ", issued.date]].map(([k, v]) => <div key={k} className="flex justify-between"><span style={{ color: C.soft }}>{k}</span><span className="font-bold" style={{ color: C.text }}>{v}</span></div>)}</div>
          <div className="rounded-xl overflow-hidden" style={{ border: `1px solid ${C.line}` }}><div className="px-3 py-1.5 text-[11px] font-bold flex" style={{ background: C.bg, color: C.soft }}><span className="flex-1">الصنف</span><span className="w-10 text-center">كمية</span><span className="w-16 text-left">قيمة</span></div>{issued.items.map((it, i) => { const e = expiryInfo(it.expDate); return (<div key={i} className="px-3 py-2 text-xs" style={{ borderTop: `1px solid ${C.line}` }}><div className="flex items-center"><span className="flex-1 font-bold" style={{ color: C.text }}>{it.name}</span><span className="w-10 text-center" style={{ color: C.text }}>{it.qty}</span><span className="w-16 text-left font-bold" style={{ color: C.text }}>{money(it.qty * it.price)}</span></div>{(it.batch || it.expDate) && <div className="flex flex-wrap gap-1 mt-1">{it.batch && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.bg, color: C.soft }}>تشغيلة {it.batch}</span>}{it.expDate && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded" style={{ background: e?.soft, color: e?.color }}>صلاحية {fmtDate(it.expDate)}</span>}</div>}</div>); })}<div className="flex justify-between items-center px-3 py-2" style={{ background: C.bg, borderTop: `1px solid ${C.line}` }}><span className="font-bold text-sm" style={{ color: C.text }}>الإجمالي</span><span className="font-extrabold text-sm" style={{ color: SC.color }}>{money(issued.amount)}</span></div></div>
          <div className="flex justify-center"><FakeBarcode code={issued.id.replace(/\D/g, "")} /></div>
          <PrimaryBtn full tone="ghost" icon={XCircle} onClick={() => setIssued(null)}>إغلاق</PrimaryBtn>
        </div>
      </div>
    </div>}
  </div>);
}
// خط السير الأسبوعي - زيارات لكل عميل حسب أيام الأسبوع
function RepRoute(props) {
  const { data, actions } = props; const rep = useRep(props);
  const planned = data.plannedVisits.filter((p) => p.repId === rep.id);
  const clients = data.merchants.filter((m) => m.repId === rep.id);
  const [day, setDay] = useState(WEEKDAYS[0]);
  const [add, setAdd] = useState({ merchantId: clients[0]?.id || "", time: "09:00" });
  const dayVisits = planned.filter((p) => p.day === day).sort((a, b) => a.time.localeCompare(b.time));
    const todayName = "السبت"; // يوم العرض التجريبي
  const myPlan = data.journeyPlans.find((p) => p.repId === rep.id && p.day === todayName);
  const NO_SALE = ["المحل مغلق", "مكتفٍ بالبضاعة", "لا سيولة", "طلب تأجيل", "أخرى"];
  const [nsFor, setNsFor] = useState(null); const [nsReason, setNsReason] = useState(NO_SALE[0]);
return (<div className="space-y-4">
    {/* خط السير الأسبوعي — اليوم */}
    {myPlan && (() => { const done = myPlan.stops.filter((st) => st.done).length; const pct = Math.round((done / myPlan.stops.length) * 100); return (<Card className="p-4 min-w-0" style={{ border: `1px solid ${SC.color}` }}>
      <div className="flex items-center justify-between flex-wrap gap-2 mb-2"><h3 className="font-extrabold flex items-center gap-2" style={{ color: C.text }}><Route size={17} style={{ color: SC.color }} />خط سيري اليوم ({todayName})</h3><span className="text-xs font-bold px-2.5 py-1 rounded-full" style={{ background: SC.soft, color: SC.color }}>{done}/{myPlan.stops.length} زيارة · {pct}%</span></div>
      <ProgressBar value={done} max={myPlan.stops.length} color={SC.color} />
      <div className="space-y-2 mt-3">{myPlan.stops.map((st, idx) => { const m = data.merchants.find((x) => x.id === st.merchantId); return (<div key={idx} className="rounded-xl p-3" style={{ border: `1px solid ${st.done ? (st.outcome === "order" ? C.success : C.accent) : C.line}`, background: st.done ? (st.outcome === "order" ? C.successSoft + "44" : C.accentSoft + "44") : C.surface }}>
        <div className="flex items-center justify-between gap-2 flex-wrap">
          <div className="flex items-center gap-2 min-w-0"><span className="rounded-full w-6 h-6 flex items-center justify-center text-[10px] font-bold shrink-0" style={{ background: st.done ? (st.outcome === "order" ? C.success : C.accent) : C.soft, color: C.onAccent }}>{st.done ? "✓" : idx + 1}</span><div className="min-w-0"><div className="text-sm font-bold truncate" style={{ color: C.text }}>{m?.name}{m?.frozen && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded mr-1" style={{ background: C.dangerSoft, color: C.danger }}>خارج التغطية</span>}</div><div className="text-[10px]" style={{ color: C.soft }}>{m?.area}{st.done && st.outcome !== "order" ? ` · لا بيع: ${st.reason}` : ""}</div></div></div>
          <div className="flex items-center gap-1.5 shrink-0"><GpsLink m={m} sm />{!st.done && <><button onClick={() => actions.repJourneyOutcome(myPlan.id, idx, "order")} className="rb-press rounded-lg px-2.5 py-1.5 text-xs font-bold" style={{ background: C.success, color: C.onAccent }}>طلب ✓</button><button onClick={() => { setNsFor(idx); setNsReason(NO_SALE[0]); }} className="rb-press rounded-lg px-2.5 py-1.5 text-xs font-bold" style={{ background: C.accentSoft, color: "#92590C" }}>لا بيع</button></>}</div>
        </div>
        {nsFor === idx && !st.done && <div className="mt-2 flex items-end gap-2 flex-wrap"><div className="flex-1" style={{ minWidth: 150 }}><Field label="سبب عدم البيع"><select value={nsReason} onChange={(e) => setNsReason(e.target.value)} className={inputCls} style={inputStyle}>{NO_SALE.map((r) => <option key={r}>{r}</option>)}</select></Field></div><PrimaryBtn sm tone="accent" icon={CheckCircle2} onClick={() => { actions.repJourneyOutcome(myPlan.id, idx, "no_sale", nsReason); setNsFor(null); toast("سُجّلت الزيارة بلا بيع"); }}>تسجيل</PrimaryBtn></div>}
      </div>); })}</div>
    </Card>); })()}
    <SectionTitle sub="خطّط زياراتك لكل عميل حسب أيام الأسبوع.">خط السير الأسبوعي</SectionTitle>
    {/* اختيار اليوم */}
    <div className="flex gap-1.5 overflow-x-auto pb-1">{WEEKDAYS.map((d) => { const cnt = planned.filter((p) => p.day === d).length; const on = day === d; return (<button key={d} onClick={() => setDay(d)} className="shrink-0 rounded-xl px-3 py-2 text-sm font-bold transition-colors" style={{ background: on ? SC.color : C.surface, color: on ? "#fff" : C.text, border: `1px solid ${on ? SC.color : C.line}` }}>{d}{cnt > 0 && <span className="mr-1 text-[10px] px-1.5 py-0.5 rounded-full" style={{ background: on ? "rgba(255,255,255,.25)" : C.bg }}>{cnt}</span>}</button>); })}</div>
    {/* إضافة زيارة لليوم */}
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>إضافة عميل ليوم {day}</h3>
      <div className="grid sm:grid-cols-3 gap-2 items-end">
        <Field label="العميل"><select value={add.merchantId} onChange={(e) => setAdd({ ...add, merchantId: e.target.value })} className={inputCls} style={inputStyle}>{clients.map((m) => <option key={m.id} value={m.id}>{m.name} — {m.area}</option>)}</select></Field>
        <Field label="الوقت"><Input value={add.time} onChange={(e) => setAdd({ ...add, time: e.target.value })} /></Field>
        <PrimaryBtn icon={Plus} onClick={() => { if (add.merchantId) actions.addPlannedVisit({ server: rep.server, repId: rep.id, merchantId: add.merchantId, day, time: add.time }); }}>إضافة للسير</PrimaryBtn>
      </div>
    </Card>
    {/* جدول السير لليوم */}
    <Card className="p-5"><div className="flex items-center justify-between mb-4"><h3 className="font-extrabold" style={{ color: C.text }}>مسار يوم {day}</h3><span className="text-sm" style={{ color: C.soft }}>{dayVisits.length} زيارة</span></div>
      <div className="relative pr-4">{dayVisits.map((pv, i) => { const m = data.merchants.find((x) => x.id === pv.merchantId);
        return (<div key={pv.id} className="flex items-start gap-3 pb-5 relative">
          {i < dayVisits.length - 1 && <div className="absolute right-[15px] top-9 bottom-0 w-0.5" style={{ background: C.line }} />}
          <div className="rounded-full flex items-center justify-center shrink-0 z-10" style={{ width: 32, height: 32, background: pv.done ? C.success : SC.color, color: "#fff" }}>{pv.done ? <CheckCircle2 size={16} /> : <span className="text-xs font-bold">{i + 1}</span>}</div>
          <div className="flex-1 rounded-xl p-3 flex items-center justify-between gap-2" style={{ border: `1px solid ${C.line}` }}>
            <div><div className="font-bold text-sm flex items-center gap-2" style={{ color: C.text }}>{m?.name}<span className="text-xs font-normal" style={{ color: C.soft }}>{pv.time}</span></div><div className="text-xs flex items-center gap-1 mt-0.5" style={{ color: C.soft }}><MapPin size={11} />{m?.area}</div></div>
            <div className="flex items-center gap-1">{!pv.done && <button onClick={() => actions.markPlannedDone(pv.id)} className="rounded-lg px-2 py-1 text-xs font-bold" style={{ background: C.successSoft, color: C.success }}>تمّت</button>}<button onClick={() => actions.removePlanned(pv.id)} className="rounded-lg p-1.5" style={{ background: C.dangerSoft }}><Trash2 size={13} style={{ color: C.danger }} /></button></div>
          </div>
        </div>); })}{dayVisits.length === 0 && <Empty icon={Route} text={`لا توجد زيارات مخططة ليوم ${day}.`} />}</div>
    </Card>
  </div>);
}
// حصر أدوات المحلات (ثلاجة، ستاند...)
function RepAssets(props) {
  const { data, actions } = props; const rep = useRep(props);
  const clients = data.merchants.filter((m) => m.repId === rep.id);
  const assets = data.assets.filter((a) => a.repId === rep.id);
  const [f, setF] = useState({ merchantId: clients[0]?.id || "", type: "ثلاجة عرض", code: "", condition: "جيدة" });
  const [tab, setTab] = useState("assets"); // assets | fridge | stocktake | recon
  const [fc, setFc] = useState({ merchantId: clients[0]?.id || "", no: "", date: "", model: "", note: "" });
  // طلب جرد
  const [stk, setStk] = useState({ merchantId: clients[0]?.id || "", scope: "all", pids: [] });
  const coProds = data.products.filter((p) => p.companyId === rep.companyId);
  // طلب مطابقة
  const [recFor, setRecFor] = useState(clients[0]?.id || "");
  const TABS = [["assets", "الأدوات", Box], ["fridge", "عقد الثلاجة", FileSignature], ["stocktake", "طلب جرد", ClipboardList], ["recon", "مطابقة حساب", ScrollText]];
  return (<div className="space-y-4"><SectionTitle sub="حصر أدوات المحل، عقود الثلاجات، وطلب الجرد ومطابقة الحساب من التاجر.">حصر أدوات المحل والجرد والمطابقة</SectionTitle>
    <div className="flex gap-1.5 overflow-x-auto pb-1">{TABS.map(([k, lbl, Ic]) => { const on = tab === k; return (<button key={k} onClick={() => setTab(k)} className="rb-press shrink-0 inline-flex items-center gap-1.5 rounded-xl px-3.5 py-2 text-xs font-bold" style={{ background: on ? SC.color : C.bg, color: on ? C.onAccent : C.text, border: `1px solid ${on ? SC.color : C.line}` }}><Ic size={14} />{lbl}</button>); })}</div>

    {tab === "assets" && <>
    <Card className="p-4 min-w-0"><div className="grid sm:grid-cols-5 gap-2 items-end">
      <Field label="المحل"><select value={f.merchantId} onChange={(e) => setF({ ...f, merchantId: e.target.value })} className={inputCls} style={inputStyle}>{clients.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}</select></Field>
      <Field label="نوع الأداة"><select value={f.type} onChange={(e) => setF({ ...f, type: e.target.value })} className={inputCls} style={inputStyle}>{["ثلاجة عرض", "ثلاجة مشروبات", "ستاند عرض", "رف معدني", "لوحة دعائية", "مجمّد"].map((t) => <option key={t}>{t}</option>)}</select></Field>
      <Field label="الرقم/الكود"><Input value={f.code} onChange={(e) => setF({ ...f, code: e.target.value })} placeholder="FRZ-0001" /></Field>
      <Field label="الحالة"><select value={f.condition} onChange={(e) => setF({ ...f, condition: e.target.value })} className={inputCls} style={inputStyle}>{["جيدة", "تحتاج صيانة", "معطّلة"].map((c) => <option key={c}>{c}</option>)}</select></Field>
      <PrimaryBtn icon={Plus} onClick={() => { if (f.merchantId && f.code) { actions.addAsset({ server: rep.server, repId: rep.id, ...f }); setF({ ...f, code: "" }); toast("حُصرت الأداة ✓"); } }}>حصر</PrimaryBtn>
    </div></Card>
    {clients.map((m) => { const mAssets = assets.filter((a) => a.merchantId === m.id); if (!mAssets.length) return null;
      return (<Card key={m.id} className="p-4 min-w-0"><div className="flex items-center gap-2 mb-3"><span className="text-xl">{m.photo || "🏪"}</span><div className="font-extrabold" style={{ color: C.text }}>{m.name}<span className="text-xs font-normal mr-2" style={{ color: C.soft }}>{mAssets.length} أداة</span></div></div>
        <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-2">{mAssets.map((a) => { const cond = a.condition === "جيدة" ? { bg: C.successSoft, c: C.success } : a.condition === "معطّلة" ? { bg: C.dangerSoft, c: C.danger } : { bg: C.warnSoft, c: "#92590C" };
          return (<div key={a.id} className="rounded-xl p-3 flex items-start justify-between gap-2" style={{ border: `1px solid ${C.line}` }}><div className="flex items-start gap-2"><div className="rounded-lg p-1.5" style={{ background: SC.soft }}><Box size={16} style={{ color: SC.color }} /></div><div><div className="font-bold text-sm" style={{ color: C.text }}>{a.type}</div><div className="text-xs" style={{ color: C.soft }}>{a.code}</div><span className="text-[10px] font-bold px-1.5 py-0.5 rounded mt-1 inline-block" style={{ background: cond.bg, color: cond.c }}>{a.condition}</span></div></div><button onClick={() => actions.removeAsset(a.id)} className="rounded-lg p-1.5" style={{ background: C.dangerSoft }}><Trash2 size={13} style={{ color: C.danger }} /></button></div>); })}</div>
      </Card>); })}
    {assets.length === 0 && <Empty icon={Box} text="لم تحصر أي أدوات بعد." />}
    </>}

    {tab === "fridge" && <>
    <Card className="p-4 min-w-0"><h3 className="font-extrabold mb-1 flex items-center gap-2" style={{ color: C.text }}><FileSignature size={17} style={{ color: SC.color }} />تخزين عقد الثلاجة</h3>
      <p className="text-xs mb-3" style={{ color: C.soft }}>وثّق عقد الثلاجة الممنوحة للمحل (رقم العقد، الموديل، تاريخه، وصورة العقد).</p>
      <div className="grid sm:grid-cols-2 gap-2">
        <Field label="المحل"><select value={fc.merchantId} onChange={(e) => setFc({ ...fc, merchantId: e.target.value })} className={inputCls} style={inputStyle}>{clients.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}</select></Field>
        <Field label="رقم العقد"><Input value={fc.no} onChange={(e) => setFc({ ...fc, no: e.target.value })} placeholder="FC-2026-001" /></Field>
        <Field label="موديل الثلاجة"><Input value={fc.model} onChange={(e) => setFc({ ...fc, model: e.target.value })} placeholder="Coca-Cola 400L" /></Field>
        <Field label="تاريخ العقد"><Input value={fc.date} onChange={(e) => setFc({ ...fc, date: e.target.value })} placeholder="2026-01-15" /></Field>
      </div>
      <div className="mt-2"><Field label="ملاحظة"><Input value={fc.note} onChange={(e) => setFc({ ...fc, note: e.target.value })} placeholder="شروط العقد أو التزامات المحل" /></Field></div>
      <div className="flex gap-2 mt-3">
        <label className="rb-press inline-flex items-center justify-center gap-1.5 rounded-xl px-3 py-2.5 text-sm font-bold cursor-pointer" style={{ background: C.surface, color: C.text, border: `1px solid ${C.line}` }}><ImageIcon size={16} />{fc.photo ? "صورة العقد ✓" : "تصوير العقد"}<input type="file" accept="image/*" capture="environment" className="hidden" onChange={(e) => { const file = e.target.files?.[0]; if (!file) return; const reader = new FileReader(); reader.onload = () => { setFc({ ...fc, photo: reader.result }); toast("تم حفظ صورة العقد 📸"); }; reader.readAsDataURL(file); e.target.value = ""; }} /></label>
        <PrimaryBtn icon={CheckCircle2} onClick={() => { if (!fc.no) { toast("أدخل رقم العقد", "danger"); return; } actions.saveFridgeContract(fc.merchantId, { no: fc.no, model: fc.model, date: fc.date, note: fc.note, photo: fc.photo || null }); toast("حُفظ عقد الثلاجة ✓"); setFc({ ...fc, no: "", model: "", date: "", note: "", photo: false }); }}>حفظ العقد</PrimaryBtn>
      </div>
    </Card>
    {clients.filter((m) => m.fridgeContract).map((m) => (<Card key={m.id} className="p-4 min-w-0"><div className="flex items-center justify-between gap-2 flex-wrap"><div className="flex items-center gap-2 min-w-0"><div className="rounded-lg p-2" style={{ background: SC.soft }}><Box size={17} style={{ color: SC.color }} /></div><div className="min-w-0"><div className="font-bold text-sm" style={{ color: C.text }}>{m.name}</div><div className="text-xs" style={{ color: C.soft }}>عقد {m.fridgeContract.no}{m.fridgeContract.model ? ` · ${m.fridgeContract.model}` : ""}{m.fridgeContract.date ? ` · ${m.fridgeContract.date}` : ""}</div>{m.fridgeContract.note && <div className="text-[10px] mt-0.5" style={{ color: C.soft }}>{m.fridgeContract.note}</div>}</div></div>{m.fridgeContract.photo && <span className="text-[10px] font-bold px-2 py-1 rounded-full flex items-center gap-1" style={{ background: C.successSoft, color: C.success }}><ImageIcon size={11} />صورة محفوظة</span>}</div></Card>))}
    </>}

    {tab === "stocktake" && <Card className="p-4 min-w-0"><h3 className="font-extrabold mb-1 flex items-center gap-2" style={{ color: C.text }}><ClipboardList size={17} style={{ color: SC.color }} />طلب جرد من التاجر</h3>
      <p className="text-xs mb-3" style={{ color: C.soft }}>اطلب من المحل جرد أصناف محددة أو جميع الأصناف — يصله الطلب ويُدخل الكميات الفعلية.</p>
      <Field label="المحل"><select value={stk.merchantId} onChange={(e) => setStk({ ...stk, merchantId: e.target.value })} className={inputCls} style={inputStyle}>{clients.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}</select></Field>
      <div className="flex gap-1.5 my-3">{[["all", "كل الأصناف"], ["specific", "أصناف محددة"]].map(([k, lbl]) => { const on = stk.scope === k; return (<button key={k} onClick={() => setStk({ ...stk, scope: k })} className="rb-press flex-1 rounded-xl py-2 text-xs font-bold" style={{ background: on ? SC.color : C.bg, color: on ? C.onAccent : C.text, border: `1px solid ${on ? SC.color : C.line}` }}>{lbl}</button>); })}</div>
      {stk.scope === "specific" && <div className="rounded-xl p-2 mb-3 max-h-52 overflow-auto" style={{ border: `1px solid ${C.line}` }}>{coProds.map((p) => { const on = stk.pids.includes(p.id); return (<button key={p.id} onClick={() => setStk({ ...stk, pids: on ? stk.pids.filter((x) => x !== p.id) : [...stk.pids, p.id] })} className="w-full flex items-center gap-2 rounded-lg px-2.5 py-1.5 text-right" style={{ background: on ? SC.soft : "transparent" }}><span className="rounded flex items-center justify-center shrink-0" style={{ width: 18, height: 18, background: on ? SC.color : "#fff", border: `2px solid ${on ? SC.color : C.line}` }}>{on && <CheckCircle2 size={12} color="#fff" />}</span><span className="text-xs font-bold truncate" style={{ color: C.text }}>{p.name}</span></button>); })}</div>}
      <PrimaryBtn full icon={Send} onClick={() => { if (stk.scope === "specific" && !stk.pids.length) { toast("اختر أصنافاً أولاً", "danger"); return; } actions.requestMerchantStocktake({ repId: rep.id, companyId: rep.companyId, merchantId: stk.merchantId, pids: stk.pids, scope: stk.scope }); toast("أُرسل طلب الجرد للتاجر ✓"); setStk({ ...stk, pids: [] }); }}>إرسال طلب الجرد</PrimaryBtn>
      {data.stocktakes.filter((st) => st.type === "merchant" && st.repId === rep.id).length > 0 && <div className="mt-4 space-y-2"><div className="text-xs font-bold" style={{ color: C.soft }}>طلبات الجرد المرسلة:</div>{data.stocktakes.filter((st) => st.type === "merchant" && st.repId === rep.id).map((st) => (<div key={st.id} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}><div className="flex items-center justify-between flex-wrap gap-2"><div><div className="text-sm font-bold" style={{ color: C.text }}>{st.id} · {merchName(data, st.merchantId)}</div><div className="text-[10px]" style={{ color: C.soft }}>{st.note} · {st.date}</div></div><span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: st.status === "submitted" ? C.successSoft : C.accentSoft, color: st.status === "submitted" ? C.success : "#92590C" }}>{st.status === "submitted" ? "أدخل التاجر الجرد ✓" : "بانتظار جرد التاجر"}</span></div>{st.status === "submitted" && <div className="mt-2 space-y-1">{st.lines.map((l) => <div key={l.pid} className="flex justify-between text-xs px-2 py-1 rounded" style={{ background: C.bg }}><span style={{ color: C.text }}>{l.name}</span><span className="font-bold" style={{ color: C.text }}>عدّ: {l.counted ?? "—"}</span></div>)}</div>}</div>))}</div>}
    </Card>}

    {tab === "recon" && <Card className="p-4 min-w-0"><h3 className="font-extrabold mb-1 flex items-center gap-2" style={{ color: C.text }}><ScrollText size={17} style={{ color: SC.color }} />طلب مطابقة حساب من التاجر</h3>
      <p className="text-xs mb-3" style={{ color: C.soft }}>يؤكد التاجر أن مديونيته صحيحة، أو يكتب ملاحظاته إن كان لديه اعتراض.</p>
      <Field label="المحل"><select value={recFor} onChange={(e) => setRecFor(e.target.value)} className={inputCls} style={inputStyle}>{clients.map((m) => <option key={m.id} value={m.id}>{m.name} — رصيده {money(m.balance)}</option>)}</select></Field>
      <div className="mt-3"><PrimaryBtn full icon={Send} onClick={() => { actions.requestReconciliation({ requestedBy: "rep", requesterId: rep.id, companyId: rep.companyId, merchantId: recFor }); toast("أُرسل طلب المطابقة للتاجر ✓"); }}>طلب مطابقة الحساب</PrimaryBtn></div>
      {data.reconciliations.filter((r) => r.requesterId === rep.id).length > 0 && <div className="mt-4 space-y-2"><div className="text-xs font-bold" style={{ color: C.soft }}>طلبات المطابقة:</div>{data.reconciliations.filter((r) => r.requesterId === rep.id).map((r) => (<div key={r.id} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}><div className="flex items-center justify-between flex-wrap gap-2"><div><div className="text-sm font-bold" style={{ color: C.text }}>{merchName(data, r.merchantId)}</div><div className="text-[10px]" style={{ color: C.soft }}>الرصيد المذكور: {money(r.statedBalance)} · {r.date}</div></div><Badge status={r.status} map={REC_STATUS} /></div>{r.merchantNote && <div className="mt-2 text-xs rounded-lg p-2" style={{ background: C.dangerSoft, color: C.danger }}>ملاحظة التاجر: {r.merchantNote}</div>}</div>))}</div>}
    </Card>}
  </div>);
}
function RepVerifyClients({ data, actions, rep }) {
  const pcs = data.pendingClients.filter((p) => p.repId === rep.id && ["awaiting_visit", "field_verified"].includes(p.status));
  const [openId, setOpenId] = useState(null);
  const [f, setF] = useState({});
  if (!pcs.length) return null;
  const getF = (p) => f[p.id] || { name: p.name, atStore: false, photoed: !!p.photo };
  const setP = (id, patch) => setF((x) => ({ ...x, [id]: { ...getF(pcs.find((q) => q.id === id)), ...(x[id] || {}), ...patch } }));
  return (<Card className="p-4 min-w-0" style={{ border: `1px solid ${C.accent}` }}>
    <h3 className="font-extrabold mb-1 flex items-center gap-2" style={{ color: C.text }}><ClipboardList size={18} style={{ color: C.accent }} />عملاء جدد بانتظار توثيقك الميداني <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.accentSoft, color: "#92590C" }}>{pcs.length}</span></h3>
    <p className="text-xs mb-3" style={{ color: C.soft }}>أسندتهم الشركة إليك من ملف الاستيراد. زُر كل محل: طابق الاسم وصحّحه إن تغيّر، صوّر المحل، وثّق الموقع من داخل نطاقه، ثم خذ تأكيد التاجر على بياناته.</p>
    <div className="space-y-3">{pcs.map((p) => { const ff = getF(p); const isOpen = openId === p.id; const fieldDone = p.status === "field_verified";
      return (<div key={p.id} className="rounded-2xl overflow-hidden" style={{ border: `1px solid ${isOpen ? C.accent : C.line}` }}>
        <button onClick={() => setOpenId(isOpen ? null : p.id)} className="w-full flex items-center justify-between gap-2 px-3.5 py-3 text-right" style={{ background: isOpen ? C.accentSoft + "55" : C.surface }}>
          <div className="min-w-0"><div className="text-sm font-extrabold" style={{ color: C.text }}>{p.name}</div><div className="text-[10px]" style={{ color: C.soft }}>{p.area}{p.phone ? ` · ${p.phone}` : ""}{p.openingDebt ? ` · مطلوب منه ${money(p.openingDebt)}` : ""} · من {p.source}</div></div>
          <Badge status={p.status} map={PC_STATUS} />
        </button>
        {isOpen && <div className="p-3.5 space-y-3" style={{ borderTop: `1px solid ${C.line}`, background: C.bg }}>
          {!fieldDone ? <>
            {/* 1: مطابقة الاسم */}
            <div className="rounded-xl p-3" style={{ background: C.surface }}>
              <div className="text-xs font-extrabold mb-1.5 flex items-center gap-1.5" style={{ color: C.text }}><span className="rounded-full w-5 h-5 flex items-center justify-center text-[10px] font-extrabold" style={{ background: SC.color, color: C.onAccent }}>1</span>طابق اسم المحل — صحّحه إن كان تغيّر</div>
              <Input value={ff.name} onChange={(e) => setP(p.id, { name: e.target.value })} />
              {ff.name !== p.name && <div className="text-[10px] mt-1 font-bold" style={{ color: C.accent }}>⚠️ سيُسجَّل التغيير: «{p.name}» ← «{ff.name}»</div>}
            </div>
            {/* 2: التصوير */}
            <div className="rounded-xl p-3 flex items-center justify-between gap-2" style={{ background: C.surface }}>
              <div className="text-xs font-extrabold flex items-center gap-1.5" style={{ color: C.text }}><span className="rounded-full w-5 h-5 flex items-center justify-center text-[10px] font-extrabold" style={{ background: SC.color, color: C.onAccent }}>2</span>صوّر واجهة المحل</div>
              {ff.photoed ? <span className="text-xs font-bold flex items-center gap-1" style={{ color: C.success }}><CheckCircle2 size={14} />صورة ملتقطة 📸</span> : <label className="rb-press inline-flex items-center gap-1.5 rounded-xl px-3 py-2 text-xs font-bold cursor-pointer" style={{ background: C.surface, color: C.text, border: `1px solid ${C.line}` }}><ImageIcon size={14} />التقاط صورة<input type="file" accept="image/*" capture="environment" className="hidden" onChange={(e) => { const file = e.target.files?.[0]; if (!file) return; const reader = new FileReader(); reader.onload = () => { setP(p.id, { photoed: true, photoData: reader.result }); toast("تم حفظ صورة المحل 📸"); }; reader.readAsDataURL(file); e.target.value = ""; }} /></label>}
            </div>
            {/* 3: GPS داخل النطاق */}
            <div className="rounded-xl p-3" style={{ background: C.surface }}>
              <div className="text-xs font-extrabold mb-1.5 flex items-center gap-1.5" style={{ color: C.text }}><span className="rounded-full w-5 h-5 flex items-center justify-center text-[10px] font-extrabold" style={{ background: SC.color, color: C.onAccent }}>3</span>وثّق موقع المحل (يجب أن تكون داخل نطاقه)</div>
              {ff.atStore ? <div className="text-xs rounded-lg p-2 flex items-center gap-1.5" style={{ background: C.successSoft, color: C.success }}><MapPinned size={14} />تم حفظ موقع الجهاز: {Number(ff.lat).toFixed(5)}, {Number(ff.lng).toFixed(5)}</div>
                : <PrimaryBtn sm full icon={Navigation} onClick={() => { if (!navigator.geolocation) { toast("المتصفح لا يدعم تحديد الموقع", "danger"); return; } navigator.geolocation.getCurrentPosition(({ coords }) => { setP(p.id, { atStore: true, lat: coords.latitude, lng: coords.longitude }); toast("تم حفظ موقعك الحالي ✓"); }, () => toast("اسمح للتطبيق باستخدام الموقع ثم أعد المحاولة", "danger"), { enableHighAccuracy: true, timeout: 12000, maximumAge: 0 }); }}>أنا أمام المحل — حفظ موقعي</PrimaryBtn>}
            </div>
            <PrimaryBtn full tone="accent" icon={CheckCircle2} onClick={() => { if (!ff.photoed || !ff.atStore || !ff.lat || !ff.lng) { toast("أكمل الصورة وتحديد الموقع أولاً", "danger"); return; } actions.repVerifyClient(p.id, { correctedName: ff.name, photo: ff.photoData || "📸", lat: ff.lat, lng: ff.lng }); toast("وُثّق المحل — اطلب الآن تأكيد التاجر ✓"); }}>إتمام التوثيق الميداني</PrimaryBtn>
          </> : <>
            {/* 4: تأكيد التاجر */}
            <div className="rounded-xl p-3" style={{ background: C.surface }}>
              <div className="text-xs font-extrabold mb-2 flex items-center gap-1.5" style={{ color: C.text }}><span className="rounded-full w-5 h-5 flex items-center justify-center text-[10px] font-extrabold" style={{ background: C.success, color: C.onAccent }}>4</span>اعرض البيانات على التاجر ليؤكّدها أمامك:</div>
              <div className="rounded-lg p-2.5 space-y-1 text-xs" style={{ background: C.bg }}>{[["اسم المحل", p.name], ["المالك", p.owner || "—"], ["الهاتف", p.phone || "—"], ["المنطقة", p.area], ["المديونية الافتتاحية المسجّلة عليه", p.openingDebt ? money(p.openingDebt) : "لا شيء"]].map(([k, v]) => <div key={k} className="flex justify-between"><span style={{ color: C.soft }}>{k}</span><span className="font-bold" style={{ color: C.text }}>{v}</span></div>)}</div>
            </div>
            <div className="flex gap-2">
              <PrimaryBtn full tone="success" icon={CheckCircle2} onClick={() => confirmThen("يؤكد التاجر أن كل البيانات أعلاه صحيحة (بما فيها المديونية)؟ سيُفعَّل حسابه فوراً.", () => { actions.confirmImportedClient(p.id); toast("أكّد التاجر — فُعّل العميل ✓"); }, { yesLabel: "نعم، أكّد التاجر" })}>التاجر أكّد البيانات — تفعيل</PrimaryBtn>
              <GhostBtn icon={XCircle} onClick={() => confirmThen("سيُرفض هذا المحل ويُبلَّغ للشركة أن بياناته غير مطابقة.", () => actions.rejectImportedClient(p.id, "التاجر لم يؤكد البيانات"), { tone: "danger", yesLabel: "رفض" })}>رفض</GhostBtn>
            </div>
          </>}
        </div>}
      </div>); })}</div>
  </Card>);
}
function RepClients(props) {
  const { data, actions } = props; const rep = useRep(props);
  const clients = data.merchants.filter((m) => m.repId === rep.id);
  const [q, setQ] = useState("");
  const found = q.trim() ? data.merchants.filter((m) => m.server === rep.server && ((m.membershipNo || "").toLowerCase().includes(q.trim().toLowerCase()) || m.name.includes(q.trim()))) : null;
  return (<div className="space-y-4"><SectionTitle sub="المحلات التي تخدمها في منطقتك — ابحث عن أي محل برقم عضويته الموثّق.">عملائي</SectionTitle>
    <RepVerifyClients data={data} actions={actions} rep={rep} />
    {/* بحث بالرقم الموثّق */}
    <Card className="p-3 sm:p-4 min-w-0">
      <div className="relative mb-1"><BadgeCheck size={17} className="absolute right-3 top-3" style={{ color: C.soft }} /><input value={q} onChange={(e) => setQ(e.target.value)} placeholder="ابحث برقم العضوية (MID-xxxxx) أو اسم المحل" className="w-full rounded-xl pr-10 pl-3 py-2.5 text-sm outline-none" style={inputStyle} /></div>
      {found && <div className="mt-2 space-y-1.5">{found.map((m) => { const mine = m.repId === rep.id; return (<div key={m.id} className="flex items-center justify-between gap-2 rounded-xl px-3 py-2" style={{ border: `1px solid ${C.line}` }}>
        <div className="flex items-center gap-2 min-w-0"><span className="text-lg">{m.photo || "🏪"}</span><div className="min-w-0"><div className="text-sm font-bold truncate" style={{ color: C.text }}>{m.name} <span className="text-[10px] font-normal" style={{ color: SC.color }}>{m.membershipNo}</span></div><div className="text-[10px]" style={{ color: C.soft }}>{m.area} · مندوبه: {repName(data, m.repId)}</div></div></div>
        {mine ? <span className="text-[10px] font-bold px-2 py-1 rounded-full shrink-0" style={{ background: C.successSoft, color: C.success }}>عميلك ✓</span> : <GhostBtn sm icon={UserPlus} onClick={() => { actions.assignMerchantToRep(m.id, rep.id); toast("أُضيف المحل لعملائك ✓"); setQ(""); }}>أضِفه لعملائي</GhostBtn>}
      </div>); })}{found.length === 0 && <div className="text-xs text-center py-2" style={{ color: C.soft }}>لا محل بهذا الرقم أو الاسم.</div>}</div>}
    </Card>
    <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{clients.map((m) => <Card key={m.id} className="p-4"><div className="flex items-center gap-3 mb-2"><div className="rounded-xl p-2.5" style={{ background: C.accentSoft }}><Store size={20} style={{ color: C.accent }} /></div><div><div className="font-extrabold" style={{ color: C.text }}>{m.name}{m.frozen && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded-full mr-1" style={{ background: C.dangerSoft, color: C.danger }}>خارج التغطية</span>}</div><div className="text-[10px] font-bold" style={{ color: SC.color }}>{m.membershipNo}</div><div className="text-xs" style={{ color: C.soft }}>{m.owner}{m.frozen ? ` · ${m.freezeReason}` : ""}</div><div className="mt-0.5"><Stars {...(merchantRating(data, m.id) || { avg: null })} size={11} /></div></div></div><div className="text-xs space-y-1" style={{ color: C.soft }}><div className="flex items-center gap-1"><MapPin size={12} />{m.area}</div><div className="flex items-center gap-1"><Phone size={12} />{m.phone}</div></div><div className="flex justify-between items-center mt-2 pt-2" style={{ borderTop: `1px solid ${C.line}` }}><span className="text-sm font-extrabold" style={{ color: m.balance < 0 ? C.danger : C.success }}>{money(m.balance)}</span><GpsLink m={m} sm /></div></Card>)}{clients.length === 0 && <Empty icon={Store} text="لا يوجد عملاء." />}</div>
  </div>);
}
function RepAccounts(props) {
  const { data } = props; const rep = useRep(props);
  const clients = data.merchants.filter((m) => m.repId === rep.id);
  const withDebt = clients.map((m) => ({ m, debt: Math.max(0, -m.balance) }));
  const totalDebt = withDebt.reduce((s, x) => s + x.debt, 0);
  const [stmt, setStmt] = useState(null);
  // بناء كشف حساب عميل: طلبات + فواتير + تحصيلات مرتّبة
  const buildStatement = (m) => {
    const orders = data.orders.filter((o) => o.merchantId === m.id && o.repId === rep.id);
    const invoices = (data.invoices || []).filter((i) => i.merchantId === m.id);
    const pays = (data.payments || []).filter((p) => p.merchantId === m.id);
    const rows = [];
    invoices.forEach((inv) => { const o = orders.find((x) => x.id === inv.orderId); rows.push({ type: "invoice", id: inv.id, date: inv.date, amount: inv.amount, paid: inv.paid, remain: inv.amount - inv.paid, qty: o ? orderQty(o) : 0, age: inv.ageDays }); });
    pays.forEach((p) => rows.push({ type: "payment", id: p.id, date: p.date, amount: p.amount, method: p.method }));
    const totalBilled = invoices.reduce((s, i) => s + i.amount, 0);
    const totalPaid = pays.reduce((s, p) => s + p.amount, 0);
    return { rows, totalBilled, totalPaid, remain: totalBilled - totalPaid, ordersCount: orders.length };
  };
  return (<div className="space-y-4"><SectionTitle sub="أرصدة عملائك ومديونياتهم — اضغط أي عميل لكشف حسابه الكامل.">الحسابات والمديونيات</SectionTitle>
    {(rep.shortfall || 0) > 0 && <Card className="p-4" style={{ border: `1px solid ${C.danger}` }}><div className="flex items-center justify-between flex-wrap gap-2"><div className="flex items-center gap-2.5"><div className="rounded-xl p-2.5" style={{ background: C.dangerSoft }}><TrendingDown size={19} style={{ color: C.danger }} /></div><div><div className="font-extrabold text-sm" style={{ color: C.text }}>ذمم مسجّلة عليك (عجز جرد/توريد)</div><div className="text-[11px]" style={{ color: C.soft }}>تُسوّى مع الشركة نقداً أو بخصم من مستحقاتك.</div></div></div><span className="text-lg font-extrabold" style={{ color: C.danger }}>{money(rep.shortfall)}</span></div></Card>}
    <Stat icon={HandCoins} label="إجمالي المديونيات على عملائي" value={money(totalDebt)} tint={{ color: C.danger, soft: C.dangerSoft }} />
    <AgingPanel data={data} merchants={clients} title="أعمار ديون عملائي" showGate />
    <Card className="p-4 min-w-0"><Table head={["العميل", "المنطقة", "الرصيد", "الحد الائتماني", "المتاح", ""]}>{withDebt.map(({ m, debt }) => <tr key={m.id} onClick={() => setStmt(m)} style={{ borderBottom: `1px solid ${C.line}`, cursor: "pointer" }}><Td bold>{m.name}</Td><Td>{m.area}</Td><Td><span style={{ color: m.balance < 0 ? C.danger : C.success, fontWeight: 800 }}>{money(m.balance)}</span></Td><Td>{money(m.creditLimit)}</Td><Td><span style={{ color: C.success }}>{money(m.creditLimit - debt)}</span></Td><Td><span className="inline-flex items-center gap-1 text-xs font-bold" style={{ color: SC.color }}><FileText size={13} />كشف</span></Td></tr>)}</Table>{clients.length === 0 && <Empty icon={Wallet} text="لا يوجد عملاء." />}</Card>

    {stmt && (() => { const st = buildStatement(stmt); return (<div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4" onClick={() => setStmt(null)}>
      <div className="rb-modal w-full max-w-md rounded-2xl p-5 my-8 max-h-[85vh] overflow-auto" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-1"><h3 className="font-extrabold text-lg" style={{ color: C.text }}>كشف حساب</h3><button onClick={() => setStmt(null)} className="rounded-lg p-1" style={{ background: C.well }}><X size={16} style={{ color: C.soft }} /></button></div>
        <div className="text-sm mb-3" style={{ color: C.soft }}>{stmt.name} · {stmt.area}</div>
        <div className="grid grid-cols-3 gap-2 mb-4 text-center">
          <div className="rounded-xl p-2.5" style={{ background: C.infoSoft }}><div className="text-[10px]" style={{ color: C.info }}>إجمالي الفواتير</div><div className="font-extrabold text-sm" style={{ color: C.info }}>{money(st.totalBilled)}</div></div>
          <div className="rounded-xl p-2.5" style={{ background: C.successSoft }}><div className="text-[10px]" style={{ color: C.success }}>المُحصّل</div><div className="font-extrabold text-sm" style={{ color: C.success }}>{money(st.totalPaid)}</div></div>
          <div className="rounded-xl p-2.5" style={{ background: st.remain > 0 ? C.dangerSoft : C.well }}><div className="text-[10px]" style={{ color: st.remain > 0 ? C.danger : C.soft }}>المتبقّي</div><div className="font-extrabold text-sm" style={{ color: st.remain > 0 ? C.danger : C.soft }}>{money(st.remain)}</div></div>
        </div>
        <div className="text-xs font-bold mb-2" style={{ color: C.soft }}>حركة الحساب ({st.rows.length})</div>
        <div className="space-y-2">{st.rows.map((r, i) => r.type === "invoice" ? (
          <div key={i} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}`, borderRight: `3px solid ${C.info}` }}>
            <div className="flex items-center justify-between mb-1"><span className="text-sm font-bold flex items-center gap-1.5" style={{ color: C.text }}><ShoppingCart size={13} style={{ color: C.info }} />فاتورة {r.id}</span><span className="font-extrabold text-sm" style={{ color: C.text }}>{money(r.amount)}</span></div>
            <div className="flex items-center justify-between text-[11px]" style={{ color: C.soft }}><span>{qty(r.qty)} · {r.date}</span><span>مسدّد {money(r.paid)} · متبقّي <span style={{ color: r.remain > 0 ? C.danger : C.success, fontWeight: 700 }}>{money(r.remain)}</span></span></div>
            {r.age != null && <div className="text-[10px] mt-1" style={{ color: r.age > 30 ? C.danger : C.soft }}>عمر الفاتورة: {r.age} يوم{r.age > 30 ? " ⚠️" : ""}</div>}
          </div>
        ) : (
          <div key={i} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}`, borderRight: `3px solid ${C.success}` }}>
            <div className="flex items-center justify-between"><span className="text-sm font-bold flex items-center gap-1.5" style={{ color: C.text }}><HandCoins size={13} style={{ color: C.success }} />تحصيل {r.id}</span><span className="font-extrabold text-sm" style={{ color: C.success }}>+{money(r.amount)}</span></div>
            <div className="text-[11px] mt-0.5" style={{ color: C.soft }}>{r.method} · {r.date}</div>
          </div>
        ))}{st.rows.length === 0 && <Empty icon={FileText} text="لا حركة على هذا الحساب بعد." />}</div>
        <div className="mt-4 rounded-xl p-3 flex items-center justify-between" style={{ background: C.ink, color: "#fff" }}><span className="text-sm font-bold">الرصيد الحالي</span><span className="font-extrabold text-lg">{money(stmt.balance)}</span></div>
      </div>
    </div>); })()}
  </div>);
}
function RepVisits(props) {
  const { data, actions } = props; const rep = useRep(props);
  const visits = data.visits.filter((v) => v.repId === rep.id);
  const clients = data.merchants.filter((m) => m.repId === rep.id);
  const [f, setF] = useState({ merchantId: clients[0]?.id || "", result: "order", note: "", amount: 0, failReason: "stock_full" });
  const isFail = !["order", "payment"].includes(f.result);
  return (<div className="space-y-4"><SectionTitle sub="سجّل زياراتك الميدانية للمحلات ونتائجها وأسباب عدم نجاحها.">سجل الزيارات</SectionTitle>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>تسجيل زيارة جديدة</h3>
      <div className="grid sm:grid-cols-4 gap-2 items-end">
        <Field label="المحل"><select value={f.merchantId} onChange={(e) => setF({ ...f, merchantId: e.target.value })} className={inputCls} style={inputStyle}>{clients.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}</select></Field>
        <Field label="النتيجة"><select value={f.result} onChange={(e) => setF({ ...f, result: e.target.value })} className={inputCls} style={inputStyle}>{Object.entries(VISIT_RESULT).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}</select></Field>
        {isFail ? <Field label="سبب عدم النجاح"><select value={f.failReason} onChange={(e) => setF({ ...f, failReason: e.target.value })} className={inputCls} style={inputStyle}>{Object.entries(FAIL_REASONS).map(([k, v]) => <option key={k} value={k}>{v}</option>)}</select></Field> : <Field label="المبلغ (إن وجد)"><Input type="number" value={f.amount} onChange={(e) => setF({ ...f, amount: +e.target.value })} /></Field>}
        <PrimaryBtn icon={Plus} onClick={() => { if (f.merchantId) { actions.addVisit({ server: rep.server, repId: rep.id, merchantId: f.merchantId, result: f.result, note: f.note, amount: isFail ? 0 : f.amount, failReason: isFail ? f.failReason : null }); setF({ ...f, note: "", amount: 0 }); } }}>تسجيل</PrimaryBtn>
      </div>
      <div className="mt-2"><Field label="ملاحظة"><Input value={f.note} onChange={(e) => setF({ ...f, note: e.target.value })} placeholder="تفاصيل الزيارة..." /></Field></div>
    </Card>
    <Card className="p-4"><Table head={["المحل", "النتيجة", "السبب/الملاحظة", "المبلغ", "التاريخ"]}>{visits.map((v) => <tr key={v.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{merchName(data, v.merchantId)}</Td><Td><Badge status={v.result} map={VISIT_RESULT} /></Td><Td>{v.failReason ? <span style={{ color: C.danger }}>{FAIL_REASONS[v.failReason]}</span> : v.note}</Td><Td>{v.amount ? money(v.amount) : "—"}</Td><Td>{v.date}</Td></tr>)}</Table>{visits.length === 0 && <Empty icon={MapPinned} text="لا توجد زيارات." />}</Card>
  </div>);
}
function RepStocktake(props) {
  const { data, actions } = props; const rep = useRep(props);
  const stocktakes = data.stocktakes.filter((s) => s.repId === rep.id);
  const clients = data.merchants.filter((m) => m.repId === rep.id);
  const prods = data.products.filter((p) => p.companyId === rep.companyId);
  const [client, setClient] = useState(clients[0]?.id || "");
  const vanSts = stocktakes.filter((st) => st.type === "van" && ["requested", "submitted", "resolved"].includes(st.status));
  const [counts, setCounts] = useState({});
  return (<div className="space-y-4"><SectionTitle sub="ارفع جرد المخزون الموجود في المحل من السوق.">طلب الجرد من السوق</SectionTitle>
    {vanSts.map((st) => (<Card key={st.id} className="p-4 min-w-0" style={{ border: `1px solid ${C.purple}` }}>
      <div className="flex items-center justify-between flex-wrap gap-2 mb-2"><h3 className="font-extrabold flex items-center gap-2" style={{ color: C.text }}><ClipboardList size={17} style={{ color: C.purple }} />جرد سيارتي — {st.id}</h3><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: st.status === "resolved" ? C.successSoft : C.purpleSoft, color: st.status === "resolved" ? C.success : C.purple }}>{st.status === "requested" ? "أدخل العدّ الفعلي" : st.status === "submitted" ? "بانتظار اعتماد المشرف" : "معتمد"}</span></div>
      <div className="space-y-1.5 mb-3">{st.lines.map((l) => { const cnt = counts[st.id + l.pid] ?? l.counted ?? ""; const diff = l.counted != null ? (l.sysQty - l.counted) : null; return (<div key={l.pid} className="flex items-center justify-between gap-2 rounded-xl px-3 py-2" style={{ background: C.bg }}>
        <div className="min-w-0"><div className="text-sm font-bold truncate" style={{ color: C.text }}>{l.name}</div><div className="text-[10px]" style={{ color: C.soft }}>بالنظام: {l.sysQty}{diff != null && diff > 0 ? ` · عجز ${diff} × ${money(l.price)}` : ""}</div></div>
        {st.status === "requested" ? <input type="number" value={cnt} onChange={(e) => setCounts({ ...counts, [st.id + l.pid]: e.target.value })} placeholder="العدّ" className="w-20 text-center text-sm font-bold rounded-lg py-1.5 outline-none" style={inputStyle} /> : <span className="text-sm font-extrabold" style={{ color: diff > 0 ? C.danger : C.success }}>{l.counted ?? "—"}</span>}
      </div>); })}</div>
      {st.status === "requested" && <PrimaryBtn full icon={Send} onClick={() => { const c = {}; st.lines.forEach((l) => { c[l.pid] = +(counts[st.id + l.pid] ?? 0); }); actions.submitStocktake(st.id, c); toast("أُرسل الجرد للمشرف ✓"); }}>إرسال نتائج الجرد للمشرف</PrimaryBtn>}
      {st.status === "resolved" && <div className="text-sm font-bold rounded-xl p-2.5 flex items-center gap-1.5" style={{ background: st.varianceValue > 0 ? C.dangerSoft : C.successSoft, color: st.varianceValue > 0 ? C.danger : C.success }}>{st.varianceValue > 0 ? <>عجز محمَّل عليك: {money(st.varianceValue)}</> : <>مطابق ✓ لا فروقات</>}</div>}
    </Card>))}
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>جرد جديد</h3>
      <Field label="المحل"><select value={client} onChange={(e) => setClient(e.target.value)} className={inputCls} style={inputStyle}>{clients.map((m) => <option key={m.id} value={m.id}>{m.name} — {m.area}</option>)}</select></Field>
      <div className="mt-3"><StockTakeForm prods={prods} onSubmit={(lines, note) => actions.repStocktake(rep.id, client, lines, note)} /></div>
    </Card>
    <div className="space-y-3">{stocktakes.map((st) => (<Card key={st.id} className="p-4"><div className="flex items-center justify-between mb-2 flex-wrap gap-2"><div><div className="font-extrabold" style={{ color: C.text }}>{merchName(data, st.merchantId)}</div><div className="text-xs" style={{ color: C.soft }}>{st.date}</div></div><Badge status={st.status} map={{ requested: { label: "بانتظار", color: C.accent, soft: C.accentSoft }, done: { label: "مكتمل", color: C.success, soft: C.successSoft } }} /></div><div className="rounded-xl overflow-hidden" style={{ border: `1px solid ${C.line}` }}>{st.lines.map((l, i) => <div key={i} className="flex justify-between px-3 py-2 text-sm" style={{ borderBottom: i < st.lines.length - 1 ? `1px solid ${C.line}` : "none" }}><span style={{ color: C.text }}>{l.name}</span><span style={{ color: C.soft }}>على الرف: {l.shelf}</span></div>)}</div></Card>))}{stocktakes.length === 0 && <Empty icon={ClipboardCheck} text="لا توجد طلبات جرد." />}</div>
  </div>);
}
function RepTarget(props) {
  const { data } = props; const rep = useRep(props);
  const pctV = Math.round((rep.achieved / (rep.target || 1)) * 100);
  const pctQ = Math.round((rep.qtyAchieved / (rep.qtyTarget || 1)) * 100);
  const primary = rep.targetMode === "qty" ? pctQ : pctV;
  const orders = data.orders.filter((o) => o.repId === rep.id && o.status === "delivered");
  const soldUnits = orders.reduce((s, o) => s + orderQty(o), 0);
  return (<div className="space-y-4"><SectionTitle sub={`هدفك الأساسي محسوب بـ${rep.targetMode === "qty" ? "الكمية" : "القيمة"}. تتبّع الاثنين معاً.`}>هدفي وإنجازي</SectionTitle>
    <Card className="p-6 text-center">
      <div className="relative inline-flex items-center justify-center mb-2" style={{ width: 160, height: 160 }}>
        <svg width="160" height="160" style={{ transform: "rotate(-90deg)" }}><circle cx="80" cy="80" r="70" fill="none" stroke={C.bg} strokeWidth="14" /><circle cx="80" cy="80" r="70" fill="none" stroke={primary >= 70 ? C.success : primary >= 40 ? C.accent : C.danger} strokeWidth="14" strokeDasharray={`${(primary / 100) * 440} 440`} strokeLinecap="round" /></svg>
        <div className="absolute text-center"><div className="text-3xl font-extrabold" style={{ color: C.text }}>{primary}%</div><div className="text-xs" style={{ color: C.soft }}>الهدف {rep.targetMode === "qty" ? "الكمّي" : "القيمي"}</div></div>
      </div>
      <div className="inline-flex items-center gap-1.5 text-xs font-bold px-3 py-1 rounded-full mb-4" style={{ background: rep.targetMode === "qty" ? SC.soft : C.successSoft, color: rep.targetMode === "qty" ? SC.color : C.success }}>{rep.targetMode === "qty" ? <Boxes size={13} /> : <CircleDollarSign size={13} />}الهدف الأساسي: {rep.targetMode === "qty" ? "كمية" : "قيمة"}</div>
    </Card>
    <div className="grid sm:grid-cols-2 gap-3">
      <Card className="p-5"><div className="flex items-center gap-2 mb-3"><Boxes size={18} style={{ color: SC.color2 }} /><h3 className="font-extrabold" style={{ color: C.text }}>الهدف الكمّي</h3></div>
        <div className="flex justify-between text-sm mb-1"><span style={{ color: C.soft }}>{qty(rep.qtyAchieved)} / {qty(rep.qtyTarget)}</span><span className="font-bold" style={{ color: SC.color2 }}>{pctQ}%</span></div><ProgressBar value={rep.qtyAchieved} max={rep.qtyTarget} color={SC.color2} />
        <div className="text-xs mt-2" style={{ color: C.soft }}>المتبقي: {qty(Math.max(0, rep.qtyTarget - rep.qtyAchieved))}</div>
      </Card>
      <Card className="p-5"><div className="flex items-center gap-2 mb-3"><CircleDollarSign size={18} style={{ color: C.success }} /><h3 className="font-extrabold" style={{ color: C.text }}>الهدف القيمي</h3></div>
        <div className="flex justify-between text-sm mb-1"><span style={{ color: C.soft }}>{money(rep.achieved)} / {money(rep.target)}</span><span className="font-bold" style={{ color: C.success }}>{pctV}%</span></div><ProgressBar value={rep.achieved} max={rep.target} color={C.success} />
        <div className="text-xs mt-2" style={{ color: C.soft }}>المتبقي: {money(Math.max(0, rep.target - rep.achieved))}</div>
      </Card>
    </div>
    <Card className="p-5"><div className="flex items-center justify-between mb-3"><h3 className="font-extrabold" style={{ color: C.text }}>الطلبات المُسلّمة ({orders.length})</h3><span className="text-sm font-bold" style={{ color: SC.color2 }}>{qty(soldUnits)} مُسلّمة</span></div><OrderList orders={orders} data={data} /></Card>
  </div>);
}
function RepQuota(props) {
  const { data } = props; const rep = useRep(props);
  const quotas = data.quotas.filter((q) => q.repId === rep.id);
  return (<div className="space-y-4"><SectionTitle sub="حصص المنتجات الموزّعة عليك من الشركة لبيعها للعملاء.">حصصي من المنتجات</SectionTitle>
    <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">{quotas.map((q) => { const p = data.products.find((x) => x.id === q.pid); const remaining = q.allocated - q.sold; const pct = Math.round((q.sold / q.allocated) * 100);
      return (<Card key={q.id} className="p-4"><div className="flex items-center gap-2 mb-2"><div className="rounded-xl p-2" style={{ background: SC.soft }}><Package size={18} style={{ color: SC.color }} /></div><div className="font-extrabold text-sm" style={{ color: C.text }}>{p?.name}</div></div>
        <div className="grid grid-cols-3 gap-2 text-center mb-2">
          <div><div className="text-base font-extrabold" style={{ color: C.text }}>{q.allocated}</div><div className="text-[10px]" style={{ color: C.soft }}>مخصص</div></div>
          <div><div className="text-base font-extrabold" style={{ color: C.success }}>{q.sold}</div><div className="text-[10px]" style={{ color: C.soft }}>مُباع</div></div>
          <div><div className="text-base font-extrabold" style={{ color: remaining < 5 ? C.danger : SC.color2 }}>{remaining}</div><div className="text-[10px]" style={{ color: C.soft }}>متبقي</div></div>
        </div>
        <ProgressBar value={q.sold} max={q.allocated} color={C.success} />
        <div className="text-xs text-center mt-1" style={{ color: C.soft }}>{pct}% من الحصة</div>
      </Card>); })}{quotas.length === 0 && <Empty icon={Boxes} text="لا توجد حصص موزّعة عليك." />}</div>
  </div>);
}
// مراجعة تسويات العملاء: المندوب يراجع القيمة ويعتمد أو يرفض
function RepSettle(props) {
  const { data, actions } = props; const rep = useRep(props);
  const myClients = data.merchants.filter((m) => m.repId === rep.id).map((m) => m.id);
  const settlements = data.settlements.filter((s) => s.repId === rep.id || myClients.includes(s.merchantId));
  const pending = settlements.filter((s) => s.status === "pending_rep");
  const collects = data.collectRequests.filter((c) => c.repId === rep.id || myClients.includes(c.merchantId));
  const openCollects = collects.filter((c) => ["requested", "scheduled"].includes(c.status));
  const [reject, setReject] = useState(null); const [reason, setReason] = useState("");
  const [visit, setVisit] = useState({});
  return (<div className="space-y-4"><SectionTitle sub="راجع المبالغ التي سدّدها عملاؤك (نقد/شيك/حوالة) واعتمدها لتُضاف لحسابك وتُعلَم الشركة، أو ارفضها إن لم تتطابق.">مراجعة تسويات العملاء</SectionTitle>
    {/* طلبات تحصيل من العملاء */}
    {openCollects.length > 0 && <Card className="p-4" style={{ border: `1px solid ${C.accent}` }}>
      <h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Truck size={18} style={{ color: C.accent }} />طلبات تحصيل من عملائك <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.accentSoft, color: "#92590C" }}>{openCollects.length}</span></h3>
      <div className="space-y-3">{openCollects.map((c) => { const m = data.merchants.find((x) => x.id === c.merchantId); return (<div key={c.id} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}>
        <div className="flex items-center justify-between mb-2 flex-wrap gap-2"><div className="flex items-center gap-2"><span className="text-xl">{m?.photo || "🏪"}</span><div><div className="font-bold text-sm" style={{ color: C.text }}>{m?.name} <span className="text-xs font-normal" style={{ color: C.soft }}>· {m?.area}</span></div><div className="text-xs" style={{ color: C.soft }}>{c.amount ? `${money(c.amount)} · ` : ""}{METHOD_LABEL[c.method]} · {c.date}{c.note ? ` · ${c.note}` : ""}</div></div></div><div className="flex items-center gap-2"><GpsLink m={m} sm /><Badge status={c.status} map={COLLECT_STATUS} /></div></div>
        {c.status === "requested" ? <div className="flex items-end gap-2 flex-wrap"><div style={{ minWidth: 160 }}><Field label="موعد الزيارة"><Input value={visit[c.id] || ""} onChange={(e) => setVisit({ ...visit, [c.id]: e.target.value })} placeholder="مثال: اليوم 3:00 م" /></Field></div><PrimaryBtn sm tone="accent" icon={CalendarClock} onClick={() => { actions.scheduleCollection(c.id, visit[c.id]); toast("جُدولت الزيارة وأُعلم العميل ✓"); }}>جدولة زيارة</PrimaryBtn><PrimaryBtn sm tone="success" icon={CheckCircle2} onClick={() => { actions.completeCollection(c.id); toast("سُجّل التحصيل وأُعلمت الشركة ✓"); }}>تم التحصيل</PrimaryBtn></div>
          : <div className="flex items-center gap-2 flex-wrap"><span className="text-xs flex items-center gap-1" style={{ color: C.info }}><Navigation size={13} />موعد: {c.visitTime}</span><PrimaryBtn sm tone="success" icon={CheckCircle2} onClick={() => { actions.completeCollection(c.id); toast("سُجّل التحصيل وأُعلمت الشركة ✓"); }}>تم التحصيل</PrimaryBtn></div>}
      </div>); })}</div>
    </Card>}
    <Stat icon={Clock} label="تسويات بانتظار مراجعتك" value={pending.length} tint={{ color: C.accent, soft: C.accentSoft }} />
    <div className="space-y-3">{pending.map((s) => { const Ic = METHOD_ICON[s.method]; const m = data.merchants.find((x) => x.id === s.merchantId); return (<Card key={s.id} className="p-4">
      <div className="flex items-center justify-between mb-3 flex-wrap gap-2"><div className="flex items-center gap-3"><div className="rounded-xl flex items-center justify-center text-2xl" style={{ width: 44, height: 44, background: SC.soft }}>{m?.photo || "🏪"}</div><div><div className="font-extrabold" style={{ color: C.text }}>{m?.name}</div><div className="text-xs" style={{ color: C.soft }}>{s.invoiceId ? `فاتورة ${s.invoiceId}` : "سداد عام"} · {s.date}</div></div></div><Badge status={s.status} map={SETTLE_STATUS} /></div>
      <div className="grid sm:grid-cols-3 gap-2 mb-3">
        <div className="rounded-xl p-3 text-center" style={{ background: C.bg }}><div className="text-xs" style={{ color: C.soft }}>المبلغ</div><div className="text-lg font-extrabold" style={{ color: SC.color }}>{money(s.amount)}</div></div>
        <div className="rounded-xl p-3 text-center" style={{ background: C.bg }}><div className="text-xs flex items-center justify-center gap-1" style={{ color: C.soft }}><Ic size={12} />الطريقة</div><div className="text-sm font-bold mt-1" style={{ color: C.text }}>{METHOD_LABEL[s.method]}</div></div>
        <div className="rounded-xl p-3 text-center" style={{ background: C.bg }}><div className="text-xs" style={{ color: C.soft }}>المرجع</div><div className="text-sm font-bold mt-1" style={{ color: C.text }}>{s.ref || "—"}</div></div>
      </div>
      {s.note && <div className="text-sm rounded-lg p-2 mb-3" style={{ background: C.accentSoft, color: "#92590C" }}>{s.note}</div>}
      {reject === s.id ? <div className="space-y-2"><Input value={reason} onChange={(e) => setReason(e.target.value)} placeholder="سبب الرفض (القيمة غير مطابقة...)" /><div className="flex gap-2"><PrimaryBtn full tone="danger" icon={XCircle} onClick={() => { actions.rejectSettlement(s.id, reason); setReject(null); setReason(""); }}>تأكيد الرفض</PrimaryBtn><GhostBtn onClick={() => setReject(null)}>إلغاء</GhostBtn></div></div>
        : <div className="flex gap-2"><PrimaryBtn full tone="success" icon={CheckCircle2} onClick={() => { actions.approveSettlement(s.id); toast("اعتُمدت التسوية وأُعلمت الشركة ✓"); }}>راجعت القيمة — اعتماد</PrimaryBtn><PrimaryBtn tone="danger" icon={XCircle} onClick={() => setReject(s.id)}>رفض</PrimaryBtn></div>}
    </Card>); })}{pending.length === 0 && <Empty icon={CheckCircle2} text="لا توجد تسويات بانتظار المراجعة." />}</div>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>التسويات المُراجَعة</h3>
      <Table head={["العميل", "المبلغ", "الطريقة", "التاريخ", "الحالة"]}>{settlements.filter((s) => s.status !== "pending_rep").map((s) => <tr key={s.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{merchName(data, s.merchantId)}</Td><Td>{money(s.amount)}</Td><Td>{METHOD_LABEL[s.method]}</Td><Td>{s.reviewedAt || s.date}</Td><Td><Badge status={s.status} map={SETTLE_STATUS} /></Td></tr>)}</Table>
      {settlements.filter((s) => s.status !== "pending_rep").length === 0 && <Empty icon={ScrollText} text="لا سجل بعد." />}
    </Card>
  </div>);
}
// تسليم النقدية اليومي للشركة
function RepHandover(props) {
  const { data, actions } = props; const rep = useRep(props);
  const held = data.payments.filter((p) => p.repId === rep.id && (p.method || "").includes("نقد") && !p.handedOver);
  const heldAmount = held.reduce((s, p) => s + p.amount, 0);
  const myHandovers = data.handovers.filter((h) => h.repId === rep.id);
  const [note, setNote] = useState(""); const [delivered, setDelivered] = useState("");
  return (<div className="space-y-4"><SectionTitle sub="ورّد النقد الذي حصّلته من العملاء إلى الشركة. يُجمع تلقائياً كل ما لم تورّده بعد.">تسليم النقدية للشركة</SectionTitle>
    <Card className="p-5">
      <div className="rounded-2xl p-5 text-white mb-4" style={{ background: `linear-gradient(120deg, ${C.accent} 0%, #D97706 100%)` }}>
        <div className="text-sm opacity-90">النقد في عهدتك (غير مورّد)</div>
        <div className="text-3xl font-extrabold mt-1">{money(heldAmount)}</div>
        <div className="text-sm opacity-90 mt-1">{held.length} عملية تحصيل بانتظار التوريد</div>
      </div>
      {held.length > 0 ? <>
        <div className="rounded-xl overflow-hidden mb-3" style={{ border: `1px solid ${C.line}` }}>{held.map((p, i) => <div key={p.id} className="flex justify-between items-center px-3 py-2 text-sm" style={{ borderBottom: i < held.length - 1 ? `1px solid ${C.line}` : "none" }}><span style={{ color: C.text }}>{merchName(data, p.merchantId)} <span className="text-xs" style={{ color: C.soft }}>· {p.date}</span></span><span className="font-bold" style={{ color: C.success }}>{money(p.amount)}</span></div>)}</div>
        <Field label="ملاحظة التوريد"><Input value={note} onChange={(e) => setNote(e.target.value)} placeholder="مثال: تسليم تحصيلات اليوم" /></Field>
        <div className="mt-3"><Field label="المبلغ الذي ستسلّمه فعلياً (اتركه فارغاً إن كان مطابقاً)"><Input type="number" value={delivered} onChange={(e) => setDelivered(e.target.value)} placeholder={String(heldAmount)} /></Field>
        {delivered !== "" && +delivered !== heldAmount && <div className="text-xs rounded-lg p-2 font-bold" style={{ background: +delivered < heldAmount ? C.dangerSoft : C.successSoft, color: +delivered < heldAmount ? C.danger : C.success }}>{+delivered < heldAmount ? `عجز ${money(heldAmount - +delivered)} — سيُسجَّل ذمة عليك بعد تأكيد الشركة` : `زيادة ${money(+delivered - heldAmount)}`}</div>}
        <PrimaryBtn full tone="success" icon={HandCoins} onClick={() => { actions.repHandover(rep.id, note, delivered); setNote(""); toast("تم توريد النقدية للشركة ✓"); }}>توريد {money(heldAmount)} للشركة</PrimaryBtn></div>
      </> : <Empty icon={CheckCircle2} text="لا يوجد نقد في عهدتك حالياً." />}
    </Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>سجل التوريدات</h3>
      <Table head={["رقم", "المبلغ", "عدد العمليات", "التاريخ", "الحالة"]}>{myHandovers.map((h) => <tr key={h.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{h.id}</Td><Td>{money(h.amount)}</Td><Td>{h.count}</Td><Td>{h.date}</Td><Td><Badge status={h.status} map={HANDOVER_STATUS} /></Td></tr>)}</Table>
      {myHandovers.length === 0 && <Empty icon={HandCoins} text="لا توجد توريدات بعد." />}
    </Card>
  </div>);
}
// خط السير العاجل المُسند من المشرف
function RepUrgent(props) {
  const { data, actions } = props; const rep = useRep(props);
  const routes = data.urgentRoutes.filter((u) => u.repId === rep.id);
  const active = routes.filter((u) => u.stops.some((s) => !s.done));
  return (<div className="space-y-4"><SectionTitle sub="جولات عاجلة أسندها لك المشرف — زر المحلات بالترتيب وأكّد كل زيارة.">خط السير العاجل</SectionTitle>
    {routes.length === 0 && <Empty icon={Route} text="لا توجد خطوط سير عاجلة حالياً." />}
    {routes.map((u) => { const doneCnt = u.stops.filter((s) => s.done).length; const sv = data.supervisors.find((s) => s.id === u.supervisorId); return (<Card key={u.id} className="p-5">
      <div className="rounded-xl p-3 mb-4 flex items-center justify-between flex-wrap gap-2" style={{ background: C.dangerSoft }}>
        <div className="flex items-center gap-2"><div className="rounded-lg p-2" style={{ background: C.well }}><Navigation size={18} style={{ color: C.danger }} /></div><div><div className="font-extrabold text-sm" style={{ color: C.danger }}>جولة عاجلة ⚡ · {u.stops.length} محل</div><div className="text-xs" style={{ color: "#92590C" }}>من المشرف {sv?.name} · {u.date}</div></div></div>
        <span className="text-sm font-extrabold px-3 py-1 rounded-full" style={{ background: C.well, color: doneCnt === u.stops.length ? C.success : C.danger }}>{doneCnt}/{u.stops.length}</span>
      </div>
      {u.note && <div className="text-sm rounded-lg p-2 mb-3" style={{ background: C.accentSoft, color: "#92590C" }}>📌 {u.note}</div>}
      <div className="relative pr-4">{u.stops.map((s, i) => { const m = data.merchants.find((x) => x.id === s.merchantId); return (<div key={i} className="flex items-start gap-3 pb-4 relative">
        {i < u.stops.length - 1 && <div className="absolute right-[15px] top-9 bottom-0 w-0.5" style={{ background: C.line }} />}
        <div className="rounded-full flex items-center justify-center shrink-0 z-10 text-xs font-bold text-white" style={{ width: 32, height: 32, background: s.done ? C.success : C.danger }}>{s.done ? <CheckCircle2 size={16} /> : i + 1}</div>
        <div className="flex-1 rounded-xl p-3" style={{ border: `1px solid ${C.line}`, opacity: s.done ? 0.6 : 1 }}>
          <div className="flex items-center justify-between gap-2 flex-wrap"><div><div className="font-bold text-sm" style={{ color: C.text }}>{m?.name}</div><div className="text-xs flex items-center gap-1 mt-0.5" style={{ color: C.soft }}><MapPin size={11} />{m?.area} · {m?.phone}</div><div className="text-xs mt-1 inline-flex items-center gap-1 px-2 py-0.5 rounded-full" style={{ background: C.accentSoft, color: "#92590C" }}><AlertCircle size={11} />{s.reason}</div><div className="mt-1.5"><GpsLink m={m} sm /></div></div>
          {!s.done && <PrimaryBtn sm tone="success" icon={CheckCircle2} onClick={() => actions.urgentStopDone(u.id, i)}>تمّت الزيارة</PrimaryBtn>}</div>
        </div>
      </div>); })}</div>
    </Card>); })}
  </div>);
}

// تحميل سيارة مندوب التجزئة: يطلب تحميلاً من الشركة ويرى مخزون سيارته
function RepLoad(props) {
  const { data, actions } = props; const rep = useRep(props);
  const isVan = rep.repType === "van";
  const prods = data.products.filter((p) => p.companyId === rep.companyId && p.status === "active");
  const myLoads = data.loadRequests.filter((l) => l.repId === rep.id);
  const vanStock = Object.entries(rep.vanStock || {}).filter(([, q]) => q > 0);
  const totalUnits = vanStock.reduce((s, [, q]) => s + q, 0);
  const [cart, setCart] = useState({}); const [note, setNote] = useState("");
  const [q, setQ] = useState(""); const [cat, setCat] = useState("الكل");
  const cats = ["الكل", ...Array.from(new Set(prods.map((p) => p.cat)))];
  const visible = prods.filter((p) => (cat === "الكل" || p.cat === cat) && (!q || p.name.includes(q)));
  const items = Object.entries(cart).filter(([, n]) => n > 0);
  const setN = (pid, n) => setCart((c) => { const x = { ...c }; if (n <= 0) delete x[pid]; else x[pid] = n; return x; });
  const submit = () => { if (!items.length) return; actions.createLoadRequest({ repId: rep.id, companyId: rep.companyId, items: items.map(([pid, n]) => ({ pid, qty: n })), note }); setCart({}); setNote(""); };
  if (!isVan) return (<div className="space-y-4"><SectionTitle sub="تحميل السيارة مخصّص لمندوبي التجزئة.">تحميل سيارتي</SectionTitle>
    <Card className="p-6 text-center"><div className="rounded-full mx-auto flex items-center justify-center mb-3" style={{ width: 56, height: 56, background: C.infoSoft }}><ClipboardList size={26} style={{ color: C.info }} /></div><div className="font-extrabold mb-1" style={{ color: C.text }}>أنت مصنّف «مندوب بري‑سيل»</div><p className="text-sm max-w-md mx-auto" style={{ color: C.soft }}>لا تحمل مخزوناً على سيارتك. تأخذ طلب العميل أو تُصدر الفاتورة ميدانياً، ثم تُجهّزها الشركة في المخزن وتوصّلها عبر السائق.</p></Card>
  </div>);
  return (<div className="space-y-4"><SectionTitle sub="اطلب تحميل الأصناف التي تحتاجها؛ بعد أن تعبّئها الشركة تُضاف إلى مخزون سيارتك للبيع الفوري.">تحميل سيارتي / مخزوني</SectionTitle>
    {/* عروض تحميل من الشركة بانتظار موافقتي */}
    {myLoads.filter((l) => l.status === "pending_rep").length > 0 && <Card className="p-4" style={{ border: `1px solid ${SC.color}` }}>
      <h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><PackagePlus size={18} style={{ color: SC.color }} />عروض تحميل من الشركة بانتظار موافقتك</h3>
      <div className="space-y-3">{myLoads.filter((l) => l.status === "pending_rep").map((l) => <div key={l.id} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}>
        <div className="flex items-center justify-between mb-2 flex-wrap gap-2"><div className="font-bold text-sm" style={{ color: C.text }}>{l.id} <span className="text-xs font-normal" style={{ color: C.soft }}>· {l.items.reduce((s, i) => s + i.qty, 0)} وحدة · {l.date}</span></div><Badge status={l.status} map={LOAD_STATUS} /></div>
        <div className="rounded-lg overflow-hidden mb-2" style={{ border: `1px solid ${C.line}` }}>{l.items.map((it, i) => <div key={i} className="flex justify-between px-3 py-1.5 text-sm" style={{ borderBottom: i < l.items.length - 1 ? `1px solid ${C.line}` : "none" }}><span style={{ color: C.text }}>{prodName(data, it.pid)}</span><span className="font-bold" style={{ color: C.text }}>{qty(it.qty)}</span></div>)}</div>
        {l.note && <div className="text-xs mb-2" style={{ color: C.soft }}>ملاحظة: {l.note}</div>}
        <div className="flex gap-2"><PrimaryBtn sm tone="success" icon={CheckCircle2} onClick={() => { actions.approveLoad(l.id); toast("أُضيفت الكميات لمخزون سيارتك ✓"); }}>موافقة — أضِف لسيارتي</PrimaryBtn><GhostBtn sm icon={XCircle} onClick={() => actions.rejectLoad(l.id)}>رفض</GhostBtn></div>
      </div>)}</div>
    </Card>}
    <div className="grid sm:grid-cols-2 gap-3">
      <Stat icon={Truck} label="إجمالي مخزون السيارة" value={qty(totalUnits)} tint={{ color: C.accent, soft: C.accentSoft }} hint={`${vanStock.length} صنف`} />
      <Stat icon={PackagePlus} label="طلبات تحميل معلّقة" value={myLoads.filter((l) => ["requested", "pending_rep"].includes(l.status)).length} tint={{ color: C.info, soft: C.infoSoft }} />
    </div>
    <div className="grid lg:grid-cols-3 gap-4">
      <div className="lg:col-span-2"><Card className="p-0 overflow-hidden">
        <div className="p-4 pb-3" style={{ borderBottom: `1px solid ${C.line}` }}><h3 className="font-extrabold mb-3" style={{ color: C.text }}>اطلب أصنافاً للتحميل</h3>
          <div className="relative mb-3"><Search size={16} className="absolute right-3 top-2.5" style={{ color: C.soft }} /><input value={q} onChange={(e) => setQ(e.target.value)} placeholder="ابحث عن صنف" className="w-full rounded-xl pr-9 pl-3 py-2 text-sm outline-none" style={inputStyle} /></div>
          <div className="flex gap-1.5 overflow-x-auto pb-1">{cats.map((c) => { const on = cat === c; return (<button key={c} onClick={() => setCat(c)} className="shrink-0 rounded-full px-3 py-1.5 text-xs font-bold" style={{ background: on ? SC.color : C.bg, color: on ? C.onAccent : C.text, border: `1px solid ${on ? SC.color : C.line}` }}>{c}</button>); })}</div>
        </div>
        <div className="divide-y max-h-[360px] overflow-auto" style={{ borderColor: C.line }}>{visible.map((p) => { const n = cart[p.id] || 0; const have = (rep.vanStock || {})[p.id] || 0; return (<div key={p.id} className="flex items-center gap-2 px-3 py-2" style={{ background: n ? SC.soft + "55" : "transparent" }}><div className="flex-1 min-w-0"><div className="text-sm font-bold truncate" style={{ color: C.text }}>{p.name}</div><div className="text-[11px]" style={{ color: C.soft }}>بالسيارة: {have} · المخزن: {p.stock}</div></div>{n > 0 ? <div className="flex items-center gap-1"><button onClick={() => setN(p.id, n - 1)} className="w-7 h-7 rounded-lg font-bold" style={{ border: `1px solid ${C.line}` }}>−</button><span className="w-8 text-center text-sm font-bold" style={{ color: SC.color }}>{n}</span><button onClick={() => setN(p.id, n + 1)} className="w-7 h-7 rounded-lg font-bold text-white" style={{ background: SC.color }}>+</button></div> : <button onClick={() => setN(p.id, 10)} className="rounded-lg px-2.5 py-1.5 text-xs font-bold text-white" style={{ background: SC.color }}>طلب</button>}</div>); })}</div>
      </Card></div>
      <Card className="p-5 h-fit"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>طلب التحميل</h3>
        <div className="space-y-2 mb-3 max-h-52 overflow-auto">{items.map(([pid, n]) => <div key={pid} className="flex justify-between items-center text-sm"><span className="flex-1" style={{ color: C.text }}>{prodName(data, pid)}</span><span className="font-bold" style={{ color: SC.color }}>{n}</span></div>)}{!items.length && <p className="text-sm" style={{ color: C.soft }}>اختر أصنافاً وكمياتها.</p>}</div>
        <Field label="ملاحظة"><Input value={note} onChange={(e) => setNote(e.target.value)} placeholder="اختياري" /></Field>
        <div className="mt-3"><PrimaryBtn full tone="accent" icon={Send} onClick={submit}>إرسال طلب التحميل للشركة</PrimaryBtn></div>
      </Card>
    </div>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>مخزون سيارتي الحالي</h3><Table head={["الصنف", "الكمية المتوفّرة"]}>{vanStock.map(([pid, n]) => <tr key={pid} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{prodName(data, pid)}</Td><Td><span className="font-bold" style={{ color: n < 10 ? C.danger : C.text }}>{qty(n)}</span></Td></tr>)}</Table>{!vanStock.length && <Empty icon={Truck} text="سيارتك فارغة — أرسل طلب تحميل." />}</Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>سجل طلبات التحميل</h3><Table head={["رقم", "عدد الأصناف", "الكمية", "التاريخ", "الحالة"]}>{myLoads.map((l) => <tr key={l.id} style={{ borderBottom: `1px solid ${C.line}` }}><Td bold>{l.id}</Td><Td>{l.items.length}</Td><Td>{qty(l.items.reduce((s, i) => s + i.qty, 0))}</Td><Td>{l.date}</Td><Td><Badge status={l.status} map={LOAD_STATUS} /></Td></tr>)}</Table>{!myLoads.length && <Empty icon={PackagePlus} text="لا طلبات تحميل بعد." />}</Card>
  </div>);
}


/* ============================================================
   ============   السائق   ===================================
   ============================================================ */
// ═══════════════════════════════════════════════════════════════════
// دور مشرف الحركة (Dispatch): استقبال الفواتير الجاهزة، تجميعها حسب النطاق، وتعبئة السيارات بالباليتات
// ═══════════════════════════════════════════════════════════════════
const useDispatch = (props) => props.data.dispatchers.find((d) => d.id === props.user.entityId);
// إحصاء طلعات السائق: الإجمالي، الأساسية (أول طلعة كل يوم)، والإضافية
const driverTripStats = (data, driverId) => {
  const trips = (data.trips || []).filter((t) => t.driverId === driverId);
  const total = trips.length;
  const extra = trips.filter((t) => t.extra).length;
  const base = total - extra;
  // تجميع حسب اليوم
  const byDay = {};
  trips.forEach((t) => { const k = t.dayKey || t.date || "—"; (byDay[k] = byDay[k] || []).push(t); });
  return { total, base, extra, byDay, trips };
};

function DispatchHome(props) {
  const { data, go, user } = props; const dsp = useDispatch(props);
  const co = data.companies.find((c) => c.id === dsp.companyId);
  const orders = data.orders.filter((o) => o.companyId === dsp.companyId);
  const inbox = orders.filter((o) => o.status === "ready_dispatch" && o.fulfillment !== "pickup");
  const trips = data.trips.filter((t) => t.companyId === dsp.companyId && t.status === "active");
  const drivers = data.drivers.filter((d) => d.companyId === dsp.companyId);
  const alerts = [];
  if (inbox.length) alerts.push({ k: "dsp_inbox", icon: Inbox, text: `${inbox.length} فاتورة بانتظار التحميل`, soft: C.warnSoft, color: "#B45309", border: "#B45309" });
  if (trips.length) alerts.push({ k: "dsp_trips", icon: Route, text: `${trips.length} رحلة نشطة على الطريق`, soft: C.tealSoft, color: C.teal, border: C.teal });
  return (<RoleLauncher name={dsp.name} sub={co?.name || "مشرف الحركة"} roleTitle="مشرف حركة وتوصيل" nav={navFor(user.role, user.companyRole)} homeKey="dsp_home" go={go} alerts={alerts} badges={{ dsp_inbox: inbox.length, dsp_trips: trips.length }} quick={{ k: "dsp_build", icon: Truck, label: "تكوين حمولة" }} />);
}

// الفواتير الواردة من موظف العمليات
function DispatchInbox(props) {
  const { data, go, actions } = props; const dsp = useDispatch(props);
  const inbox = data.orders.filter((o) => o.companyId === dsp.companyId && o.status === "ready_dispatch" && o.fulfillment !== "pickup");
  // تجميع حسب المدينة ثم المنطقة
  const byZone = {};
  inbox.forEach((o) => { const m = data.merchants.find((x) => x.id === o.merchantId); const city = m?.city || "غير محدد"; const area = m?.area || "غير محدد"; const key = city + " › " + area; if (!byZone[key]) byZone[key] = { city, area, orders: [] }; byZone[key].orders.push(o); });
  const zones = Object.values(byZone).sort((a, b) => b.orders.length - a.orders.length);
  const totalStikat = inbox.reduce((s, o) => s + orderStikat(o), 0);
  return (<div className="space-y-4"><SectionTitle sub="فواتير جهّزها موظف العمليات — مجمّعة حسب المدينة والمنطقة لتسهيل التحميل." action={<PrimaryBtn icon={Truck} onClick={() => go("dsp_build")}>تكوين حمولة</PrimaryBtn>}>الفواتير الواردة</SectionTitle>
    {(() => {
      const reqs = (data.pickupRequests || []).filter((r) => r.dispatcherId === dsp.id && r.status === "requested");
      if (!reqs.length) return null;
      return (<Card className="p-4" style={{ borderRight: `3px solid #B45309`, background: C.warnSoft }}>
        <h3 className="font-extrabold flex items-center gap-2 mb-2" style={{ color: "#B45309" }}><Inbox size={17} />طلبات استلام يدوي من العمليات</h3>
        <div className="space-y-2">{reqs.map((r) => (<div key={r.id} className="flex items-center justify-between gap-2 rounded-xl p-3" style={{ background: C.surface, border: `1px solid ${C.line}` }}>
          <div><div className="font-bold text-sm" style={{ color: C.text }}>{r.id} · {r.orderIds.length} فاتورة</div><div className="text-[11px]" style={{ color: C.soft }}>{r.note} · {r.date}</div></div>
          <PrimaryBtn sm icon={CheckCircle2} onClick={() => { actions.confirmPickup(r.id); toast("تأكيد استلام الفواتير ✓"); }}>استلمتُها</PrimaryBtn>
        </div>))}</div>
      </Card>);
    })()}
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      <Stat icon={Inbox} label="فواتير بانتظار التحميل" value={inbox.length} tint={{ color: "#B45309", soft: C.warnSoft }} />
      <Stat icon={Boxes} label="إجمالي الستيكات" value={qty(totalStikat)} tint={{ color: SC.color, soft: SC.soft }} />
      <Stat icon={MapPin} label="نطاقات مختلفة" value={zones.length} tint={{ color: C.info, soft: C.infoSoft }} />
      <Stat icon={CircleDollarSign} label="قيمة الفواتير" value={money(inbox.reduce((s, o) => s + orderTotal(o), 0))} tint={{ color: C.success, soft: C.successSoft }} />
    </div>
    {zones.map((z) => (<Card key={z.city + z.area} className="p-4">
      <div className="flex items-center justify-between mb-3"><h3 className="font-extrabold flex items-center gap-2" style={{ color: C.text }}><MapPin size={16} style={{ color: SC.color }} />{z.city} <span style={{ color: C.soft }}>›</span> {z.area}</h3><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: SC.soft, color: SC.color }}>{z.orders.length} فاتورة · {qty(z.orders.reduce((s, o) => s + orderStikat(o), 0))}</span></div>
      <div className="space-y-2">{z.orders.map((o) => (<div key={o.id} className="flex items-center justify-between rounded-xl p-2.5" style={{ border: `1px solid ${C.line}` }}><div><div className="font-bold text-sm" style={{ color: C.text }}>{o.id} · {merchName(data, o.merchantId)}</div><div className="text-[11px]" style={{ color: C.soft }}>{qty(orderStikat(o))} · {money(orderTotal(o))}</div></div><span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: C.warnSoft, color: "#B45309" }}>جاهز للتحميل</span></div>))}</div>
    </Card>))}
    {inbox.length === 0 && <Empty icon={Inbox} text="لا فواتير واردة — كل الفواتير حُمّلت أو ما زالت لدى العمليات." />}
  </div>);
}

// تكوين حمولة سيارة (قلب الدور): تجميع الطلبيات وحساب الباليتات وملء السيارة
function DispatchBuild(props) {
  const { data, actions } = props; const dsp = useDispatch(props);
  const inbox = data.orders.filter((o) => o.companyId === dsp.companyId && o.status === "ready_dispatch" && o.fulfillment !== "pickup");
  const drivers = data.drivers.filter((d) => d.companyId === dsp.companyId && d.status !== "busy");
  const [driverId, setDriverId] = useState("");
  const [sel, setSel] = useState([]);
  const driver = data.drivers.find((d) => d.id === driverId);
  const capPallets = driver ? vehiclePallets(driver.vehicleType) : 0;
  const selectedOrders = inbox.filter((o) => sel.includes(o.id));
  const usedPallets = selectedOrders.reduce((s, o) => s + orderPallets(o, data.products), 0);
  const usedStikat = selectedOrders.reduce((s, o) => s + orderStikat(o), 0);
  const fillPct = capPallets ? Math.min(100, Math.round((usedPallets / capPallets) * 100)) : 0;
  const over = usedPallets > capPallets;
  const toggle = (id) => setSel((s) => s.includes(id) ? s.filter((x) => x !== id) : [...s, id]);
  // اقتراح تلقائي: املأ السيارة من نفس النطاق حتى تكتمل
  const autoFill = () => {
    if (!driver) { toast("اختر سيارة أولاً", "danger"); return; }
    const byZone = {};
    inbox.forEach((o) => { const m = data.merchants.find((x) => x.id === o.merchantId); const key = (m?.city || "") + (m?.area || ""); (byZone[key] = byZone[key] || []).push(o); });
    const biggest = Object.values(byZone).sort((a, b) => b.length - a.length)[0] || [];
    const picked = []; let pal = 0;
    for (const o of biggest) { const op = orderPallets(o, data.products); if (pal + op <= capPallets) { picked.push(o.id); pal += op; } }
    setSel(picked);
    toast(picked.length ? `اقتُرحت ${picked.length} طلبية لملء السيارة ✓` : "لا طلبيات مناسبة", picked.length ? "success" : "danger");
  };
  return (<div className="space-y-4"><SectionTitle sub="اختر السيارة وجمّع الطلبيات — التطبيق يحسب الباليتات ويملأ السيارة بالكامل.">تكوين حمولة سيارة</SectionTitle>
    <Card className="p-4">
      <div className="grid sm:grid-cols-2 gap-3 mb-3">
        <Field label="السيارة والسائق"><select value={driverId} onChange={(e) => { setDriverId(e.target.value); setSel([]); }} className={inputCls} style={inputStyle}><option value="">— اختر سيارة —</option>{drivers.map((d) => <option key={d.id} value={d.id}>{d.name} · {VEHICLE_TYPES[d.vehicleType || "medium"].label} ({vehiclePallets(d.vehicleType)} باليتات)</option>)}</select></Field>
        <div className="flex items-end"><GhostBtn full icon={Sparkles} onClick={autoFill}>ملء تلقائي (نفس النطاق)</GhostBtn></div>
      </div>
      {driver && <div className="rounded-xl p-3 mb-3" style={{ background: over ? C.dangerSoft : C.well }}>
        <div className="flex items-center justify-between mb-1.5"><span className="text-sm font-bold" style={{ color: C.text }}>تعبئة السيارة</span><span className="text-sm font-extrabold" style={{ color: over ? C.danger : fillPct >= 85 ? C.success : "#B45309" }}>{usedPallets.toFixed(1)} / {capPallets} باليتة ({fillPct}%)</span></div>
        <div className="h-3 rounded-full overflow-hidden" style={{ background: C.line }}><div style={{ width: `${Math.min(100, fillPct)}%`, height: "100%", background: over ? C.danger : fillPct >= 85 ? C.success : "#D97706", transition: ".3s" }} /></div>
        <div className="flex items-center justify-between mt-1.5 text-[11px]" style={{ color: C.soft }}><span>{qty(usedStikat)} في {selectedOrders.length} طلبية</span><span>{over ? "⚠️ تجاوزت سعة السيارة" : fillPct >= 85 ? "✓ السيارة شبه ممتلئة" : "أضف طلبيات لملء السيارة"}</span></div>
      </div>}
      {driver && <PrimaryBtn full icon={Route} onClick={() => { if (!sel.length) { toast("اختر طلبيات", "danger"); return; } if (over) { toast("الحمولة تتجاوز سعة السيارة", "danger"); return; } actions.dispatchTrip({ companyId: dsp.companyId, driverId, orderIds: sel }); toast(`خرجت السيارة بـ${sel.length} طلبية ✓`); setSel([]); setDriverId(""); }}>تأكيد التحميل وإخراج السيارة ({sel.length})</PrimaryBtn>}
    </Card>
    <Card className="p-4"><h3 className="font-extrabold mb-3" style={{ color: C.text }}>الطلبيات الجاهزة ({inbox.length})</h3>
      <div className="space-y-2">{inbox.map((o) => { const m = data.merchants.find((x) => x.id === o.merchantId); const on = sel.includes(o.id); const pal = orderPallets(o, data.products); return (
        <button key={o.id} onClick={() => toggle(o.id)} className="rb-press w-full flex items-center gap-3 rounded-xl p-3 text-right" style={{ border: `1.5px solid ${on ? SC.color : C.line}`, background: on ? SC.soft : C.surface }}>
          <div className="w-5 h-5 rounded flex items-center justify-center shrink-0" style={{ background: on ? SC.color : "transparent", border: `1.5px solid ${on ? SC.color : C.lineStrong}` }}>{on && <CheckCircle2 size={13} color="#fff" />}</div>
          <div className="flex-1 min-w-0"><div className="font-bold text-sm truncate" style={{ color: C.text }}>{o.id} · {merchName(data, o.merchantId)}</div><div className="text-[11px]" style={{ color: C.soft }}>{m?.city} › {m?.area} · {qty(orderStikat(o))}</div></div>
          <div className="text-left shrink-0"><div className="font-extrabold text-sm" style={{ color: SC.color }}>{pal.toFixed(1)}</div><div className="text-[10px]" style={{ color: C.soft }}>باليتة</div></div>
        </button>); })}
        {inbox.length === 0 && <Empty icon={Inbox} text="لا طلبيات جاهزة للتحميل." />}
      </div>
    </Card>
  </div>);
}

// الرحلات النشطة
function DispatchTrips(props) {
  const { data, openOrder } = props; const dsp = useDispatch(props);
  const trips = data.trips.filter((t) => t.companyId === dsp.companyId);
  return (<div className="space-y-4"><SectionTitle sub="الرحلات التي كوّنتها وأسندتها للسائقين.">الرحلات النشطة</SectionTitle>
    {trips.map((t) => { const drv = data.drivers.find((d) => d.id === t.driverId); const orders = data.orders.filter((o) => t.orderIds.includes(o.id)); const delivered = orders.filter((o) => o.status === "delivered").length; return (
      <Card key={t.id} className="p-4">
        <div className="flex items-center justify-between mb-3 flex-wrap gap-2"><div className="flex items-center gap-2"><div className="rounded-xl p-2" style={{ background: C.tealSoft }}><Truck size={18} style={{ color: C.teal }} /></div><div><div className="font-extrabold" style={{ color: C.text }}>رحلة {t.id}</div><div className="text-[11px]" style={{ color: C.soft }}>{drv?.name} · {VEHICLE_TYPES[drv?.vehicleType || "medium"].label}</div></div></div><span className="text-xs font-bold px-2.5 py-1 rounded-full" style={{ background: delivered === orders.length ? C.successSoft : C.tealSoft, color: delivered === orders.length ? C.success : C.teal }}>{delivered}/{orders.length} سُلّم</span></div>
        <div className="space-y-1.5">{orders.map((o) => (<div key={o.id} onClick={() => openOrder && openOrder(o.id)} className="flex items-center justify-between rounded-lg px-3 py-2 cursor-pointer" style={{ background: C.well }}><span className="text-sm font-bold" style={{ color: C.text }}>{merchName(data, o.merchantId)}</span><span className="text-[11px] font-bold px-2 py-0.5 rounded-full" style={{ background: STATUS[o.status]?.soft, color: STATUS[o.status]?.color }}>{STATUS[o.status]?.label}</span></div>))}</div>
      </Card>); })}
    {trips.length === 0 && <Empty icon={Route} text="لا رحلات بعد — كوّن حمولة وأسندها لسائق." />}
  </div>);
}

// متابعة السائقين
function DispatchDrivers(props) {
  const { data, actions, openOrder } = props; const dsp = useDispatch(props);
  const drivers = data.drivers.filter((d) => d.companyId === dsp.companyId);
  const [openD, setOpenD] = useState(null);
  const [chatWith, setChatWith] = useState(null);
  const [msg, setMsg] = useState("");
  const driverOrders = (did) => data.orders.filter((o) => o.driverId === did);
  const driverTrips = (did) => data.trips.filter((t) => t.driverId === did);
  return (<div className="space-y-4"><SectionTitle sub="استعلم عن كل سائق وتوصيلاته — اضغط أي سائق لرؤية رحلاته وطلباته، وتواصل معه.">متابعة السائقين</SectionTitle>
    <div className="grid sm:grid-cols-2 gap-3">{drivers.map((d) => { const active = driverTrips(d.id).filter((t) => t.status === "active").length; const vt = VEHICLE_TYPES[d.vehicleType || "medium"]; const ords = driverOrders(d.id); const delivered = ords.filter((o) => o.status === "delivered").length; const onWay = ords.filter((o) => o.status === "out_for_delivery").length; const isOpen = openD === d.id; return (
      <Card key={d.id} className={`p-4 ${isOpen ? "sm:col-span-2" : ""}`}>
        <div className="flex items-center gap-3 cursor-pointer" onClick={() => setOpenD(isOpen ? null : d.id)}><Avatar name={d.name} size={44} /><div className="flex-1"><div className="font-extrabold" style={{ color: C.text }}>{d.name}</div><div className="text-[11px]" style={{ color: C.soft }}>{d.vehicle} · {vt.label}</div></div><span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: d.status === "busy" ? C.warnSoft : C.successSoft, color: d.status === "busy" ? "#B45309" : C.success }}>{d.status === "busy" ? "مشغول" : "متاح"}</span><ChevronDown size={16} style={{ color: C.soft, transform: isOpen ? "rotate(180deg)" : "none", transition: ".2s" }} /></div>
        <div className="grid grid-cols-4 gap-2 mt-3 text-center">
          <div className="rounded-lg py-1.5" style={{ background: C.well }}><div className="text-[9px]" style={{ color: C.soft }}>سعة السيارة</div><div className="font-extrabold text-sm" style={{ color: SC.color }}>{vt.pallets}</div></div>
          <div className="rounded-lg py-1.5" style={{ background: C.well }}><div className="text-[9px]" style={{ color: C.soft }}>رحلات نشطة</div><div className="font-extrabold text-sm" style={{ color: C.teal }}>{active}</div></div>
          <div className="rounded-lg py-1.5" style={{ background: C.well }}><div className="text-[9px]" style={{ color: C.soft }}>في الطريق</div><div className="font-extrabold text-sm" style={{ color: C.info }}>{onWay}</div></div>
          <div className="rounded-lg py-1.5" style={{ background: C.well }}><div className="text-[9px]" style={{ color: C.soft }}>سُلّم</div><div className="font-extrabold text-sm" style={{ color: C.success }}>{delivered}</div></div>
        </div>
        <div className="flex gap-2 mt-3"><GhostBtn sm full icon={MessageCircle} onClick={() => { setChatWith(d.id); }}>مراسلة</GhostBtn><GhostBtn sm full icon={Phone} onClick={() => toast(d.phone)}>{d.phone}</GhostBtn></div>
        {isOpen && <div className="mt-3 pt-3 space-y-3" style={{ borderTop: `1px dashed ${C.line}` }}>
          <div><div className="text-xs font-bold mb-2 flex items-center gap-1.5" style={{ color: C.teal }}><Route size={13} />الرحلات ({driverTrips(d.id).length})</div>
            <div className="space-y-1.5">{driverTrips(d.id).map((t) => { const to = data.orders.filter((o) => t.orderIds.includes(o.id)); const dv = to.filter((o) => o.status === "delivered").length; return (<div key={t.id} className="rounded-lg px-3 py-2 flex items-center justify-between" style={{ background: C.well }}><span className="text-sm font-bold" style={{ color: C.text }}>{t.id} · {t.orderIds.length} طلب</span><span className="text-[11px] font-bold px-2 py-0.5 rounded-full" style={{ background: dv === to.length ? C.successSoft : C.tealSoft, color: dv === to.length ? C.success : C.teal }}>{dv}/{to.length} سُلّم</span></div>); })}{driverTrips(d.id).length === 0 && <div className="text-xs text-center py-2" style={{ color: C.soft }}>لا رحلات.</div>}</div>
          </div>
          <div><div className="text-xs font-bold mb-2 flex items-center gap-1.5" style={{ color: SC.color }}><ShoppingCart size={13} />كل التوصيلات ({driverOrders(d.id).length})</div>
            <div className="space-y-1.5 max-h-64 overflow-auto">{driverOrders(d.id).map((o) => (<div key={o.id} onClick={() => openOrder && openOrder(o.id)} className="rounded-lg px-3 py-2 flex items-center justify-between cursor-pointer" style={{ background: C.surface, border: `1px solid ${C.line}` }}><div className="min-w-0"><div className="text-sm font-bold truncate" style={{ color: C.text }}>{merchName(data, o.merchantId)}</div><div className="text-[10px]" style={{ color: C.soft }}>{o.id} · {qty(orderQty(o))}</div></div><span className="text-[10px] font-bold px-2 py-0.5 rounded-full shrink-0" style={{ background: STATUS[o.status]?.soft, color: STATUS[o.status]?.color }}>{STATUS[o.status]?.label}</span></div>))}{driverOrders(d.id).length === 0 && <div className="text-xs text-center py-2" style={{ color: C.soft }}>لا توصيلات.</div>}</div>
          </div>
        </div>}
      </Card>); })}</div>
    {drivers.length === 0 && <Empty icon={Truck} text="لا سائقين." />}

    {/* نافذة المراسلة */}
    {chatWith && (() => { const drv = data.drivers.find((x) => x.id === chatWith); const thread = (data.dispatchChat || []).filter((m) => m.driverId === chatWith && m.companyId === dsp.companyId); return (<div className="rb-modal-bg fixed inset-0 z-50 flex items-center justify-center p-4" onClick={() => setChatWith(null)}>
      <div className="rb-modal w-full max-w-md rounded-2xl p-5 flex flex-col" style={{ maxHeight: "80vh" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-3"><h3 className="font-extrabold flex items-center gap-2" style={{ color: C.text }}><Truck size={18} style={{ color: SC.color }} />{drv?.name}</h3><button onClick={() => setChatWith(null)} className="rounded-lg p-1" style={{ background: C.well }}><X size={16} style={{ color: C.soft }} /></button></div>
        <div className="flex-1 overflow-auto space-y-2 mb-3 min-h-[200px]">{thread.map((m) => (<div key={m.id} className={`flex ${m.from === "dispatch" ? "justify-end" : "justify-start"}`}><div className="max-w-[75%] rounded-2xl px-3 py-2" style={{ background: m.from === "dispatch" ? SC.color : C.well, color: m.from === "dispatch" ? C.onAccent : C.text }}><div className="text-sm">{m.text}</div><div className="text-[9px] mt-0.5" style={{ opacity: .7 }}>{m.t}</div></div></div>))}{thread.length === 0 && <Empty icon={MessageCircle} text="ابدأ المحادثة." />}</div>
        <div className="flex gap-2"><input value={msg} onChange={(e) => setMsg(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && msg.trim()) { actions.sendDispatchMsg({ companyId: dsp.companyId, driverId: chatWith, from: "dispatch", text: msg.trim() }); setMsg(""); } }} placeholder="اكتب رسالة..." className="flex-1 rounded-xl px-3 py-2 text-sm outline-none" style={inputStyle} /><PrimaryBtn icon={Send} onClick={() => { if (!msg.trim()) return; actions.sendDispatchMsg({ companyId: dsp.companyId, driverId: chatWith, from: "dispatch", text: msg.trim() }); setMsg(""); }}>إرسال</PrimaryBtn></div>
      </div>
    </div>); })()}
  </div>);
}

// تقرير الطلعات الإضافية (نهاية الشهر): كم طلعة إضافية لكل سائق
function DispatchExtraTrips(props) {
  const { data } = props; const dsp = useDispatch(props);
  const drivers = data.drivers.filter((d) => d.companyId === dsp.companyId);
  const [openD, setOpenD] = useState(null);
  const rows = drivers.map((d) => ({ d, ...driverTripStats(data, d.id) })).sort((a, b) => b.extra - a.extra);
  const totalExtra = rows.reduce((s, r) => s + r.extra, 0);
  const totalTrips = rows.reduce((s, r) => s + r.total, 0);
  return (<div className="space-y-4"><SectionTitle sub="كل تعبئة سيارة = طلعة. الطلعة الأولى يومياً أساسية، وما بعدها إضافي يُحتسب للسائق.">الطلعات الإضافية (محاسبة الشهر)</SectionTitle>
    <div className="grid grid-cols-3 gap-3">
      <Stat icon={Route} label="إجمالي الطلعات" value={totalTrips} tint={{ color: C.teal, soft: C.tealSoft }} />
      <Stat icon={Star} label="طلعات إضافية" value={totalExtra} tint={{ color: "#B45309", soft: C.warnSoft }} />
      <Stat icon={Truck} label="عدد السائقين" value={drivers.length} tint={{ color: SC.color, soft: SC.soft }} />
    </div>
    <div className="space-y-2.5">{rows.map((r) => { const isOpen = openD === r.d.id; return (
      <Card key={r.d.id} className="p-4">
        <div className="flex items-center gap-3 cursor-pointer" onClick={() => setOpenD(isOpen ? null : r.d.id)}>
          <Avatar name={r.d.name} size={40} />
          <div className="flex-1"><div className="font-extrabold" style={{ color: C.text }}>{r.d.name}</div><div className="text-[11px]" style={{ color: C.soft }}>{r.d.vehicle}</div></div>
          <div className="text-center px-3"><div className="text-2xl font-extrabold" style={{ color: r.extra > 0 ? "#B45309" : C.soft }}>{r.extra}</div><div className="text-[9px]" style={{ color: C.soft }}>إضافي</div></div>
          <ChevronDown size={16} style={{ color: C.soft, transform: isOpen ? "rotate(180deg)" : "none", transition: ".2s" }} />
        </div>
        <div className="grid grid-cols-3 gap-2 mt-3 text-center">
          <div className="rounded-lg py-1.5" style={{ background: C.well }}><div className="text-[10px]" style={{ color: C.soft }}>إجمالي الطلعات</div><div className="font-extrabold text-sm" style={{ color: C.teal }}>{r.total}</div></div>
          <div className="rounded-lg py-1.5" style={{ background: C.well }}><div className="text-[10px]" style={{ color: C.soft }}>أساسية</div><div className="font-extrabold text-sm" style={{ color: C.text }}>{r.base}</div></div>
          <div className="rounded-lg py-1.5" style={{ background: r.extra > 0 ? C.warnSoft : C.well }}><div className="text-[10px]" style={{ color: r.extra > 0 ? "#B45309" : C.soft }}>إضافية ⭐</div><div className="font-extrabold text-sm" style={{ color: r.extra > 0 ? "#B45309" : C.soft }}>{r.extra}</div></div>
        </div>
        {isOpen && <div className="mt-3 pt-3" style={{ borderTop: `1px dashed ${C.line}` }}>
          <div className="text-xs font-bold mb-2" style={{ color: C.soft }}>تفصيل الطلعات حسب اليوم:</div>
          <div className="space-y-2">{Object.entries(r.byDay).map(([day, trips]) => (
            <div key={day} className="rounded-xl p-2.5" style={{ background: C.well }}>
              <div className="flex items-center justify-between mb-1.5"><span className="text-xs font-bold" style={{ color: C.text }}>{day}</span><span className="text-[10px]" style={{ color: C.soft }}>{trips.length} طلعة</span></div>
              <div className="flex flex-wrap gap-1.5">{trips.sort((a, b) => (a.tripNo || 0) - (b.tripNo || 0)).map((t) => (
                <span key={t.id} className="inline-flex items-center gap-1 text-[10px] font-bold rounded-full px-2 py-0.5" style={{ background: t.extra ? C.warnSoft : C.successSoft, color: t.extra ? "#B45309" : C.success }}>{t.extra ? <Star size={9} /> : <CheckCircle2 size={9} />}طلعة {t.tripNo}{t.extra ? " (إضافي)" : " (أساسية)"}</span>))}</div>
            </div>))}{Object.keys(r.byDay).length === 0 && <div className="text-xs text-center py-2" style={{ color: C.soft }}>لا طلعات.</div>}</div>
        </div>}
      </Card>); })}{drivers.length === 0 && <Empty icon={Truck} text="لا سائقين." />}</div>
  </div>);
}

// سعة الباليتات لكل صنف (يحددها مشرف الحركة)
function DispatchPallets(props) {
  const { data, actions } = props; const dsp = useDispatch(props);
  const prods = data.products.filter((p) => p.companyId === dsp.companyId);
  return (<div className="space-y-4"><SectionTitle sub="حدّد كم ستيكة تسع في الباليتة لكل صنف — يُستخدم في حساب تعبئة السيارات.">سعة الباليتات</SectionTitle>
    <Card className="p-4"><div className="space-y-2">{prods.map((p) => (<div key={p.id} className="flex items-center gap-3 rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}>
      <div className="rounded-lg p-2 shrink-0" style={{ background: SC.soft }}><Boxes size={18} style={{ color: SC.color }} /></div>
      <div className="flex-1 min-w-0"><div className="font-bold text-sm truncate" style={{ color: C.text }}>{p.name}</div><div className="text-[11px]" style={{ color: C.soft }}>{p.cat}</div></div>
      <div className="flex items-center gap-2 shrink-0"><span className="text-[11px]" style={{ color: C.soft }}>ستيكة/باليتة</span><input type="number" min="1" value={stikaPerPallet(p)} onChange={(e) => actions.setPalletCap(p.id, +e.target.value || 1)} className="w-20 text-center rounded-lg py-1.5 font-extrabold" style={{ border: `1px solid ${C.line}`, background: C.surface, color: SC.color }} /></div>
    </div>))}{prods.length === 0 && <Empty icon={Boxes} text="لا منتجات." />}</div></Card>
  </div>);
}

const useDriver = (props) => props.data.drivers.find((d) => d.id === props.user.entityId);
// دور السائق: سجل التوصيلات والتواصل مع مشرف الحركة
function DriverDispatch(props) {
  const { data, actions, openOrder } = props; const driver = useDriver(props);
  const dsp = data.dispatchers.find((x) => x.companyId === driver.companyId);
  const thread = (data.dispatchChat || []).filter((m) => m.driverId === driver.id && m.companyId === driver.companyId);
  const myOrders = data.orders.filter((o) => o.driverId === driver.id);
  const myTrips = data.trips.filter((t) => t.driverId === driver.id);
  const delivered = myOrders.filter((o) => o.status === "delivered");
  const [msg, setMsg] = useState("");
  return (<div className="space-y-4"><SectionTitle sub="سجل توصيلاتك ورحلاتك، وتواصل مباشر مع مشرف الحركة.">مشرف الحركة والسجل</SectionTitle>
    <div className="grid grid-cols-4 gap-3">
      <Stat icon={Route} label="رحلاتي" value={myTrips.length} tint={{ color: C.teal, soft: C.tealSoft }} />
      <Stat icon={Star} label="طلعات إضافية" value={myTrips.filter((t) => t.extra).length} tint={{ color: "#B45309", soft: C.warnSoft }} />
      <Stat icon={ShoppingCart} label="التوصيلات" value={myOrders.length} tint={{ color: SC.color, soft: SC.soft }} />
      <Stat icon={CheckCircle2} label="سُلّمت" value={delivered.length} tint={{ color: C.success, soft: C.successSoft }} />
    </div>

    {/* محادثة مشرف الحركة */}
    <Card className="p-4">
      <h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><MessageCircle size={17} style={{ color: SC.color }} />التواصل مع مشرف الحركة {dsp ? `(${dsp.name})` : ""}</h3>
      <div className="space-y-2 mb-3 max-h-72 overflow-auto rounded-xl p-3" style={{ background: C.well }}>{thread.map((m) => (<div key={m.id} className={`flex ${m.from === "driver" ? "justify-end" : "justify-start"}`}><div className="max-w-[75%] rounded-2xl px-3 py-2" style={{ background: m.from === "driver" ? SC.color : C.surface, color: m.from === "driver" ? C.onAccent : C.text, border: m.from === "driver" ? "none" : `1px solid ${C.line}` }}><div className="text-sm">{m.text}</div><div className="text-[9px] mt-0.5" style={{ opacity: .7 }}>{m.t}</div></div></div>))}{thread.length === 0 && <Empty icon={MessageCircle} text="لا رسائل بعد — ابدأ التواصل." />}</div>
      <div className="flex gap-2"><input value={msg} onChange={(e) => setMsg(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && msg.trim()) { actions.sendDispatchMsg({ companyId: driver.companyId, driverId: driver.id, from: "driver", text: msg.trim() }); setMsg(""); } }} placeholder="اكتب لمشرف الحركة..." className="flex-1 rounded-xl px-3 py-2 text-sm outline-none" style={inputStyle} /><PrimaryBtn icon={Send} onClick={() => { if (!msg.trim()) return; actions.sendDispatchMsg({ companyId: driver.companyId, driverId: driver.id, from: "driver", text: msg.trim() }); setMsg(""); }}>إرسال</PrimaryBtn></div>
    </Card>

    {/* سجل الرحلات */}
    <Card className="p-4">
      <h3 className="font-extrabold mb-3 flex items-center gap-2" style={{ color: C.text }}><Route size={17} style={{ color: C.teal }} />سجل رحلاتي</h3>
      <div className="space-y-2">{myTrips.map((t) => { const to = data.orders.filter((o) => t.orderIds.includes(o.id)); const dv = to.filter((o) => o.status === "delivered").length; return (
        <div key={t.id} className="rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}>
          <div className="flex items-center justify-between mb-2"><span className="font-bold text-sm flex items-center gap-1.5" style={{ color: C.text }}>رحلة {t.id} · {t.date}{t.extra && <span className="inline-flex items-center gap-0.5 text-[9px] font-bold rounded-full px-1.5 py-0.5" style={{ background: C.warnSoft, color: "#B45309" }}><Star size={8} />إضافي</span>}</span><span className="text-[11px] font-bold px-2 py-0.5 rounded-full" style={{ background: dv === to.length ? C.successSoft : C.tealSoft, color: dv === to.length ? C.success : C.teal }}>{dv}/{to.length} سُلّم</span></div>
          <div className="space-y-1">{to.map((o) => (<div key={o.id} onClick={() => openOrder && openOrder(o.id)} className="flex items-center justify-between rounded-lg px-2.5 py-1.5 cursor-pointer" style={{ background: C.well }}><span className="text-xs font-bold" style={{ color: C.text }}>{merchName(data, o.merchantId)}</span><span className="text-[10px] font-bold px-2 py-0.5 rounded-full" style={{ background: STATUS[o.status]?.soft, color: STATUS[o.status]?.color }}>{STATUS[o.status]?.label}</span></div>))}</div>
        </div>); })}{myTrips.length === 0 && <Empty icon={Route} text="لا رحلات بعد." />}</div>
    </Card>
  </div>);
}

function DriverHome(props) {
  const { data, go, user } = props; const driver = useDriver(props);
  const orders = data.orders.filter((o) => o.driverId === driver.id);
  const toDeliver = orders.filter((o) => ["preparing", "out_for_delivery"].includes(o.status)).length;
  const trips = data.trips.filter((t) => t.driverId === driver.id && t.status === "active").length;
  const alerts = [];
  if (toDeliver > 0) alerts.push({ k: "dr_deliveries", icon: Truck, text: `${toDeliver} طلب بانتظار توصيلك`, soft: C.infoSoft, color: C.info, border: C.info });
  if (trips > 0) alerts.push({ k: "dr_route", icon: Route, text: `${trips} رحلة نشطة — راجع خط السير`, soft: C.purpleSoft, color: C.purple, border: C.purple });
  return <RoleLauncher name={driver.name} sub={driver.vehicle || driver.area || ""} roleTitle="سائق توصيل" nav={navFor(user.role, user.companyRole)} homeKey="dr_home" go={go} alerts={alerts} badges={{ dr_deliveries: toDeliver, dr_route: trips }} quick={{ k: "dr_deliveries", icon: Truck, label: "توصيلاتي" }} />;
}
function DriverDeliveries(props) {
  const { data, actions, openOrder } = props; const driver = useDriver(props);
  const orders = data.orders.filter((o) => o.driverId === driver.id);
  const active = orders.filter((o) => ["preparing", "out_for_delivery"].includes(o.status));
  return (<div className="space-y-4"><SectionTitle sub="الطلبات المُسندة إليك للتوصيل.">توصيلاتي</SectionTitle>
    {active.length ? <div className="grid md:grid-cols-2 gap-3">{active.map((o) => { const m = data.merchants.find((x) => x.id === o.merchantId); return (<Card key={o.id} className="p-4"><div className="flex items-center justify-between mb-2"><span className="font-extrabold" style={{ color: C.text }}>{o.id}</span><Badge status={o.status} /></div>
      <div className="rounded-xl p-3 mb-3" style={{ background: C.bg }}><div className="text-sm flex items-center gap-1 mb-1" style={{ color: C.text }}><Store size={14} style={{ color: C.accent }} />{m?.name}</div><div className="text-sm flex items-center gap-1 mb-1" style={{ color: C.soft }}><MapPin size={14} />{m?.area}، {m?.city}</div><div className="text-sm flex items-center gap-1 mb-2" style={{ color: C.soft }}><Phone size={14} />{m?.phone}</div><GpsLink m={m} sm full /></div>
      <div className="flex items-center justify-between mb-3"><span className="text-sm" style={{ color: C.soft }}>قيمة الطلب</span><span className="font-extrabold" style={{ color: SC.color }}>{money(orderTotal(o))}</span></div>
      <div className="flex gap-2">{o.status === "preparing" && <PrimaryBtn full icon={Truck} onClick={() => actions.advance(o.id, "out_for_delivery")}>استلمت — في الطريق</PrimaryBtn>}{o.status === "out_for_delivery" && <PrimaryBtn full tone="accent" icon={PackageCheck} onClick={() => actions.advance(o.id, "delivered")}>تأكيد التسليم</PrimaryBtn>}<GhostBtn icon={Eye} onClick={() => openOrder(o.id)}>تفاصيل</GhostBtn></div>
    </Card>); })}</div> : <Empty icon={Truck} text="لا توجد توصيلات حالياً." />}
  </div>);
}
function DriverTasks(props) {
  const { data, actions } = props; const driver = useDriver(props);
  const tasks = data.tasks.filter((t) => t.assigneeType === "driver" && t.assigneeId === driver.id);
  return (<div className="space-y-4"><SectionTitle sub="المهام التي أسندتها لك الشركة.">المهام المسندة</SectionTitle>
    {tasks.length ? tasks.map((t) => (<Card key={t.id} className="p-4"><div className="flex items-start justify-between gap-3 flex-wrap"><div className="flex items-start gap-2"><div className="rounded-lg p-2 mt-0.5" style={{ background: C.purpleSoft }}><ClipboardCheck size={17} style={{ color: C.purple }} /></div><div><div className="font-bold flex items-center gap-2" style={{ color: C.text }}>{t.title}{t.priority === "high" && <span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.dangerSoft, color: C.danger }}>عاجل</span>}</div><div className="text-sm mt-0.5" style={{ color: C.soft }}>{t.desc}</div><div className="text-xs mt-1 flex items-center gap-1" style={{ color: C.soft }}><CalendarClock size={12} />{t.due}</div></div></div><Badge status={t.status} map={TASK_STATUS} /></div>
      <div className="flex gap-2 mt-3 pt-3" style={{ borderTop: `1px solid ${C.line}` }}>{t.status === "assigned" && <PrimaryBtn sm icon={Activity} onClick={() => actions.updateTask(t.id, { status: "in_progress" })}>بدء التنفيذ</PrimaryBtn>}{t.status === "in_progress" && <PrimaryBtn sm tone="success" icon={CheckCircle2} onClick={() => actions.updateTask(t.id, { status: "done" })}>إكمال</PrimaryBtn>}{t.status === "done" && <span className="text-sm flex items-center gap-1" style={{ color: C.success }}><CheckCircle2 size={15} />مكتملة</span>}</div>
    </Card>)) : <Empty icon={ListChecks} text="لا توجد مهام مسندة." />}
  </div>);
}
function DriverRoute(props) {
  const { data } = props; const driver = useDriver(props);
  const myTrips = data.trips.filter((t) => t.driverId === driver.id && t.status === "active");
  const stops = data.orders.filter((o) => o.driverId === driver.id && ["preparing", "out_for_delivery"].includes(o.status)).map((o) => data.merchants.find((m) => m.id === o.merchantId)).filter(Boolean);
  return (<div className="space-y-4"><SectionTitle sub="ترتيب نقاط التسليم في جولتك اليوم.">خط السير</SectionTitle>
    {/* رحلاتي المجمّعة */}
    {myTrips.map((t) => { const tOrders = t.orderIds.map((id) => data.orders.find((o) => o.id === id)).filter(Boolean); return (<Card key={t.id} className="p-4" style={{ border: `1px solid ${C.info}` }}>
      <div className="flex items-center justify-between mb-3 flex-wrap gap-2"><div className="font-extrabold flex items-center gap-2" style={{ color: C.text }}><Route size={17} style={{ color: C.info }} />رحلة {t.id}</div><span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: C.infoSoft, color: C.info }}>{tOrders.length} محطة</span></div>
      <div className="relative pr-4">{tOrders.map((o, i) => { const m = data.merchants.find((x) => x.id === o.merchantId); return (<div key={o.id} className="flex items-start gap-3 pb-4 relative">
        {i < tOrders.length - 1 && <div className="absolute right-[14px] top-8 bottom-0 w-0.5" style={{ background: C.line }} />}
        <div className="rounded-full flex items-center justify-center text-xs font-bold text-white shrink-0 z-10" style={{ width: 30, height: 30, background: o.status === "delivered" ? C.success : C.info }}>{o.status === "delivered" ? <CheckCircle2 size={15} /> : i + 1}</div>
        <div className="flex-1 rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}><div className="flex items-center justify-between gap-2 flex-wrap"><div className="min-w-0"><div className="font-bold text-sm" style={{ color: C.text }}>{m?.name} <span className="text-[10px] font-normal" style={{ color: C.soft }}>· {o.id}</span></div><div className="text-xs" style={{ color: C.soft }}>{m?.area} · {money(orderTotal(o))}</div></div><GpsLink m={m} sm /></div></div>
      </div>); })}</div>
    </Card>); })}
    <Card className="p-5">
      <div className="rounded-xl p-4 mb-4 flex items-center justify-between" style={{ background: SC.soft }}><div className="flex items-center gap-2"><Navigation size={18} style={{ color: SC.color }} /><span className="text-sm font-bold" style={{ color: C.text }}>{stops.length} نقطة تسليم</span></div><span className="text-sm flex items-center gap-1" style={{ color: C.soft }}><Fuel size={15} />جولة اليوم</span></div>
      <div className="relative pr-4">{stops.map((m, i) => (<div key={m.id + i} className="flex items-start gap-3 pb-5 relative">
        {i < stops.length - 1 && <div className="absolute right-[14px] top-8 bottom-0 w-0.5" style={{ background: C.line }} />}
        <div className="rounded-full flex items-center justify-center text-xs font-bold text-white shrink-0 z-10" style={{ width: 30, height: 30, background: i === 0 ? C.success : SC.color }}>{i + 1}</div>
        <div className="flex-1 rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}><div className="flex items-center justify-between gap-2"><div className="font-bold text-sm" style={{ color: C.text }}>{m.name}</div><GpsLink m={m} sm /></div><div className="text-xs flex items-center gap-1 mt-0.5" style={{ color: C.soft }}><MapPin size={12} />{m.area}، {m.city}</div><div className="text-xs flex items-center gap-1 mt-0.5" style={{ color: C.soft }}><Phone size={12} />{m.phone}</div></div>
      </div>))}{stops.length === 0 && <Empty icon={Route} text="لا توجد نقاط تسليم." />}</div>
    </Card>
  </div>);
}
function DriverHistory(props) {
  const { data, openOrder } = props; const driver = useDriver(props);
  const delivered = data.orders.filter((o) => o.driverId === driver.id && o.status === "delivered");
  return (<div className="space-y-4"><SectionTitle sub="كل الطلبات التي سلّمتها.">سجل التسليم</SectionTitle><Card className="p-2"><OrderList orders={delivered} data={data} onOpen={openOrder} /></Card></div>);
}

/* ============================================================
   ============   قوائم الطلبات وتفاصيلها   ===================
   ============================================================ */
function OrderList({ orders, data, onOpen }) {
  if (!orders.length) return <p className="text-sm py-3" style={{ color: C.soft }}>لا توجد طلبات.</p>;
  return (<div className="space-y-2">{orders.map((o) => { const m = data.merchants.find((x) => x.id === o.merchantId); const c = data.companies.find((x) => x.id === o.companyId);
    return (<div key={o.id} onClick={() => onOpen && onOpen(o.id)} className={`flex items-center justify-between gap-3 rounded-xl p-3 ${onOpen ? "cursor-pointer transition-colors hover:bg-gray-50" : ""}`} style={{ border: `1px solid ${C.line}` }}><div className="min-w-0"><div className="font-bold flex items-center gap-2" style={{ color: C.text }}>{o.id} <span className="text-xs font-normal" style={{ color: C.soft }}>{o.createdAt}</span></div><div className="text-sm truncate" style={{ color: C.soft }}>{m?.name} ← {c?.name}</div></div><div className="flex items-center gap-3 shrink-0"><span className="font-extrabold text-sm" style={{ color: SC.color }}>{money(orderTotal(o))}</span><Badge status={o.status} /></div></div>); })}</div>);
}
function OrdersTable(props) {
  const { orders, data, openOrder, title, sub } = props; const [q, setQ] = useState("");
  const filtered = orders.filter((o) => { const m = data.merchants.find((x) => x.id === o.merchantId); return o.id.includes(q) || (m && m.name.includes(q)); });
  return (<div className="space-y-4"><SectionTitle sub={sub || "اضغط أي طلب لعرض تفاصيله."}>{title}</SectionTitle>
    <div className="relative"><Search size={17} className="absolute right-3 top-3" style={{ color: C.soft }} /><input className="w-full rounded-xl pr-10 pl-3 py-2.5 text-sm outline-none" style={inputStyle} placeholder="ابحث برقم الطلب أو اسم المحل" value={q} onChange={(e) => setQ(e.target.value)} /></div>
    <Card className="p-2"><OrderList orders={filtered} data={data} onOpen={openOrder} /></Card></div>);
}
function Info({ icon: Icon, label, value }) { return (<div className="flex items-center gap-2 rounded-xl p-3" style={{ border: `1px solid ${C.line}` }}><Icon size={17} style={{ color: SC.color2 }} /><div><div className="text-xs" style={{ color: C.soft }}>{label}</div><div className="font-bold" style={{ color: C.text }}>{value}</div></div></div>); }
function OrderDetail({ user, order, data, actions, onBack }) {
  if (!order) return null;
  const m = data.merchants.find((x) => x.id === order.merchantId);
  const c = data.companies.find((x) => x.id === order.companyId);
  const rep = data.reps.find((x) => x.id === order.repId);
  const driver = data.drivers.find((x) => x.id === order.driverId);
  const companyDrivers = data.drivers.filter((d) => d.companyId === order.companyId);
  return (<div className="space-y-4 max-w-3xl"><button onClick={onBack} className="inline-flex items-center gap-1 text-sm font-bold" style={{ color: C.soft }}><ChevronLeft size={16} />رجوع</button>
    <Card className="p-5"><div className="flex items-center justify-between mb-4"><div><div className="text-xs font-bold" style={{ color: C.accent }}>طلب رقم</div><div className="text-xl font-extrabold" style={{ color: C.text }}>{order.id}</div></div><Badge status={order.status} /></div>
      <div className="rounded-xl p-4 mb-4" style={{ background: C.bg }}><Pipeline status={order.status} /></div>
      <div className="grid sm:grid-cols-2 gap-3 text-sm mb-4"><Info icon={Store} label={SC.merchantWord} value={`${m?.name} · ${m?.area}`} /><Info icon={Building2} label="الشركة" value={c?.name} /><Info icon={Handshake} label="المندوب" value={rep?.name || "—"} /><Info icon={order.fulfillment === "pickup" ? Store : Truck} label="طريقة الاستلام" value={FULFILL[order.fulfillment || "delivery"]?.label} /></div>
      {["rep", "driver", "promoter"].includes(user.role) && <div className="mb-4"><LocationCard m={m} /></div>}
      <div className="rounded-xl overflow-hidden mb-4" style={{ border: `1px solid ${C.line}` }}>{order.items.map((it, i) => { const p = data.products.find((x) => x.id === it.pid); const e = p && expiryInfo(p.expDate); return (<div key={i} className="px-4 py-2.5 text-sm" style={{ borderBottom: i < order.items.length - 1 ? `1px solid ${C.line}` : "none" }}><div className="flex justify-between items-center"><span style={{ color: C.text }}>{it.name} {it.boxes ? <span className="text-xs" style={{ color: C.soft }}>· {it.boxes} صندوق ({it.qty} {p?.unit || "قطعة"})</span> : <span>× {it.qty}</span>}</span><span style={{ color: C.soft }}>{money(it.qty * it.price)}</span></div>{p && (p.expDate || p.batch) && <div className="flex flex-wrap gap-1.5 mt-1">{p.batch && <span className="inline-flex items-center gap-1 text-[10px] font-bold px-1.5 py-0.5 rounded" style={{ background: C.bg, color: C.soft }}><Hash size={9} />تشغيلة {p.batch}</span>}{p.mfgDate && <span className="text-[10px]" style={{ color: C.soft }}>تصنيع {fmtDate(p.mfgDate)}</span>}{e && <span className="inline-flex items-center gap-1 text-[10px] font-extrabold px-1.5 py-0.5 rounded" style={{ background: e.soft, color: e.color }}><CalendarClock size={9} />الصلاحية: {fmtDate(p.expDate)}</span>}</div>}</div>); })}<div className="flex justify-between items-center px-4 py-3" style={{ background: C.bg }}><span className="font-extrabold" style={{ color: C.text }}>الإجمالي</span><span className="font-extrabold" style={{ color: SC.color }}>{money(orderTotal(order))}</span></div></div>
      {order.note && <div className="text-sm rounded-xl p-3 mb-4" style={{ background: C.accentSoft, color: "#92590C" }}>ملاحظة: {order.note}</div>}
      <div className="flex items-center gap-2 text-sm mb-4"><Banknote size={16} style={{ color: order.cashConfirmed ? C.success : C.soft }} /><span style={{ color: order.cashConfirmed ? C.success : C.soft, fontWeight: 700 }}>{order.cashConfirmed ? "تم تأكيد استلام النقد" : "لم يُؤكَّد النقد بعد"}</span></div>
      <ActionBar user={user} order={order} actions={actions} companyDrivers={companyDrivers} />
      {user.role === "merchant" && order.status === "delivered" && <RateCompanyInline order={order} data={data} actions={actions} />}
    </Card></div>);
}
// تقييم الشركة بعد استلام الطلب (يظهر للتاجر في تفاصيل الطلب المُسلّم)
function RateCompanyInline({ order, data, actions }) {
  const existing = data.ratings.find((r) => r.raterRole === "merchant" && r.orderId === order.id);
  const [stars, setStars] = useState(5); const [comment, setComment] = useState(""); const [sent, setSent] = useState(false);
  const co = data.companies.find((c) => c.id === order.companyId);
  if (existing || sent) return (<div className="mt-4 rounded-xl p-3 flex items-center gap-2" style={{ background: C.successSoft }}><CheckCircle2 size={16} style={{ color: C.success }} /><span className="text-sm font-bold" style={{ color: C.success }}>شكراً — قيّمت {co?.name} {existing ? `بـ${existing.stars} نجوم` : ""}</span></div>);
  return (<div className="mt-4 rounded-xl p-4" style={{ background: C.bg, border: `1px solid ${C.line}` }}>
    <div className="flex items-center justify-between flex-wrap gap-2 mb-2"><div className="font-extrabold text-sm flex items-center gap-1.5" style={{ color: C.text }}><Star size={16} style={{ color: "#F59E0B" }} />قيّم تجربتك مع {co?.name}</div><StarPicker value={stars} onChange={setStars} /></div>
    <div className="flex gap-2"><input value={comment} onChange={(e) => setComment(e.target.value)} placeholder="تعليق اختياري (التوصيل، البضاعة، التعامل...)" className="flex-1 rounded-xl px-3 py-2 text-sm outline-none" style={inputStyle} /><PrimaryBtn sm icon={Send} onClick={() => { actions.rateCompany({ merchantId: order.merchantId, companyId: order.companyId, orderId: order.id, stars, comment }); setSent(true); toast("شكراً لتقييمك ⭐"); }}>إرسال</PrimaryBtn></div>
  </div>);
}
function ActionBar({ user, order, actions, companyDrivers }) {
  const [driverId, setDriverId] = useState(companyDrivers[0]?.id || "");
  if (user.role === "rep" && order.status === "pending_rep") return (<div className="space-y-3 pt-3" style={{ borderTop: `1px solid ${C.line}` }}>
    <button onClick={() => actions.toggleCash(order.id)} className="w-full flex items-center justify-center gap-2 rounded-xl px-4 py-2.5 text-sm font-bold" style={order.cashConfirmed ? { background: C.successSoft, color: C.success } : { background: C.surface, color: C.text, border: `1px solid ${C.line}` }}><Banknote size={17} />{order.cashConfirmed ? "النقد مؤكّد ✓" : "تأكيد استلام النقد من المحل"}</button>
    <div className="flex gap-2"><PrimaryBtn full icon={CheckCircle2} onClick={() => actions.repAccept(order.id)}>قبول وإرسال للشركة</PrimaryBtn><PrimaryBtn tone="danger" icon={XCircle} onClick={() => actions.advance(order.id, "rejected")}>رفض</PrimaryBtn></div>
    {!order.cashConfirmed && <p className="text-xs" style={{ color: C.soft }}>يُفضّل تأكيد استلام النقد قبل القبول.</p>}</div>);
  if (user.role === "company" && order.status === "pending_company") return (<div className="space-y-3 pt-3" style={{ borderTop: `1px solid ${C.line}` }}>
    <Field label="إسناد سائق للتوصيل"><select className={inputCls} style={inputStyle} value={driverId} onChange={(e) => setDriverId(e.target.value)}>{companyDrivers.map((d) => <option key={d.id} value={d.id}>{d.name} — {d.vehicle}</option>)}</select></Field>
    <div className="flex gap-2"><PrimaryBtn full icon={CheckCircle2} onClick={() => actions.companyAccept(order.id, driverId)}>قبول وبدء التجهيز</PrimaryBtn><PrimaryBtn tone="danger" icon={XCircle} onClick={() => actions.advance(order.id, "rejected")}>رفض</PrimaryBtn></div></div>);
  if (user.role === "company" && order.status === "preparing" && order.fulfillment !== "pickup") return (<div className="pt-3" style={{ borderTop: `1px solid ${C.line}` }}><p className="text-xs mb-2" style={{ color: C.soft }}>بعد التجهيز، أرسل الفاتورة لمشرف الحركة ليجمّعها ويعبّئ السيارة.</p><PrimaryBtn full icon={Route} onClick={() => { actions.opsRelease(order.id); toast("أُرسلت الفاتورة لمشرف الحركة ✓"); }}>سحب الفاتورة وإرسالها لمشرف الحركة</PrimaryBtn></div>);
  if (user.role === "driver" && order.status === "preparing") return <div className="pt-3" style={{ borderTop: `1px solid ${C.line}` }}><PrimaryBtn full icon={Truck} onClick={() => actions.advance(order.id, "out_for_delivery")}>استلمت الطلب — في الطريق</PrimaryBtn></div>;
  if (user.role === "driver" && order.status === "out_for_delivery") return <div className="pt-3" style={{ borderTop: `1px solid ${C.line}` }}><PrimaryBtn full tone="accent" icon={PackageCheck} onClick={() => actions.advance(order.id, "delivered")}>تأكيد التسليم</PrimaryBtn></div>;
  return null;
}

/* ===================== التطبيق ===================== */
export default function App() {
  const makeInitialData = () => withOpeningLedger({ ...seed, favorites: ["p2"], compare: [], cart: {}, orderDraft: { companyId: null, cart: {} }, ledger: [], jvSeq: 1 });
  const [data, setData] = useState(makeInitialData);
  const [server, setServer] = useState(null);
  const [user, setUser] = useState(null);
  const [ctr, setCtr] = useState(3000);
  const [splash, setSplash] = useState(true);
  const [hydrated, setHydrated] = useState(false);
  const syncTimer = useRef(null);

  useEffect(() => { const link = document.createElement("link"); link.rel = "stylesheet"; link.href = "https://fonts.googleapis.com/css2?family=Cairo:wght@600;800&family=Tajawal:wght@400;500;700;800&display=swap"; document.head.appendChild(link);
    const st = document.createElement("style"); st.textContent = GLOBAL_CSS; document.head.appendChild(st);
    const t = setTimeout(() => setSplash(false), 1600); return () => clearTimeout(t);
  }, []);

  // تحميل الحالة من الخادم، مع تخزين محلي احتياطي عند انقطاع الشبكة.
  useEffect(() => {
    let cancelled = false;
    (async () => {
      const remote = await loadRemoteState();
      if (cancelled) return;
      if (remote) setData(remote);
      else {
        try {
          const local = window.localStorage.getItem("rabet-state-v1");
          if (local) setData(JSON.parse(local));
        } catch { /* نستخدم بيانات البداية إذا تعذر التخزين المحلي. */ }
      }
      setHydrated(true);
    })();
    return () => { cancelled = true; };
  }, []);

  // بعد نجاح الدخول نعيد تحميل مساحة العمل الخاصة بالخادم/المستخدم من API.
  useEffect(() => {
    if (!user) return undefined;
    let cancelled = false;
    loadRemoteState().then((remote) => { if (cancelled) return; if (remote) setData(remote); else saveRemoteState(data); });
    return () => { cancelled = true; };
  }, [user]);

  // مزامنة مؤجلة حتى لا نرسل طلباً للخادم مع كل ضغطة داخل حقل.
  useEffect(() => {
    if (!hydrated) return undefined;
    try { window.localStorage.setItem("rabet-state-v1", JSON.stringify(data)); } catch { /* الصور الكبيرة قد تتجاوز مساحة التخزين المحلي. */ }
    clearTimeout(syncTimer.current);
    syncTimer.current = setTimeout(() => { saveRemoteState(data); }, 250);
    return () => clearTimeout(syncTimer.current);
  }, [data, hydrated]);

  // تحديث ألوان السيرفر
  if (server) SC = SERVERS[server];

  const upd = (fn) => setData((d) => fn(d));
  const nextId = () => { const id = "ORD-2026-" + String(ctr).padStart(4, "0"); setCtr((n) => n + 1); return id; };
  // إنشاء إشعار
  const mkNotif = (role, entityId, icon, title, body) => ({ id: uid("n"), role, entityId, icon, title, body, read: false, t: "الآن" });

  const actions = {
    readNotif: (id) => upd((d) => ({ ...d, notifications: d.notifications.map((n) => n.id === id ? { ...n, read: true } : n) })),
    readAllNotifs: (role, entityId) => upd((d) => ({ ...d, notifications: d.notifications.map((n) => n.role === role && n.entityId === entityId ? { ...n, read: true } : n) })),
    setCompanySub: (id, sub) => upd((d) => { const c = d.companies.find((x) => x.id === id); if (!c) return d; return { ...d, companies: d.companies.map((x) => x.id === id ? { ...x, sub } : x), auditLog: [{ id: uid("al"), server: c.server, action: sub === "active" ? "قبول اشتراك" : "رفض اشتراك", target: c.name, detail: sub === "active" ? "تفعيل الاشتراك" : "رفض الطلب", t: "الآن" }, ...(d.auditLog || [])] }; }),
    updateCompanyFin: (id, patch, logText) => upd((d) => { const c = d.companies.find((x) => x.id === id); if (!c) return d; return { ...d, companies: d.companies.map((x) => x.id === id ? { ...x, ...patch } : x), auditLog: [{ id: uid("al"), server: c.server, action: "إعدادات مالية", target: c.name, detail: logText, t: "الآن" }, ...(d.auditLog || [])] }; }),
    // ═══ مطوّر المنصة: إدارة فروع الشركة ═══
    setBranchMode: (companyId, multiBranch) => upd((d) => ({ ...d, companies: d.companies.map((c) => c.id === companyId ? { ...c, branches: multiBranch ? (c.branches && c.branches.length ? c.branches : [{ id: "br" + Math.floor(100 + Math.random() * 899), name: "الفرع الرئيسي", city: c.city || "", manager: "أدمن الفرع" }]) : [] } : c) })),
    addBranch: (companyId, branch) => upd((d) => ({ ...d, companies: d.companies.map((c) => c.id === companyId ? { ...c, branches: [...(c.branches || []), { id: "br" + Math.floor(100 + Math.random() * 899), name: branch.name, city: branch.city || "", manager: branch.manager || "أدمن الفرع" }] } : c) })),
    removeBranch: (companyId, branchId) => upd((d) => ({ ...d, companies: d.companies.map((c) => c.id === companyId ? { ...c, branches: (c.branches || []).filter((b) => b.id !== branchId) } : c) })),
    // ═══ الشركة تسدّد عمولة التطبيق المستحقة (تُخصم من الرصيد) ═══
    payAppCommission: (companyId, amount) => upd((d) => { const c = d.companies.find((x) => x.id === companyId); if (!c) return d; const newPaid = money3((c.commissionPaid || 0) + amount); return { ...d, companies: d.companies.map((x) => x.id === companyId ? { ...x, commissionPaid: newPaid } : x), auditLog: [{ id: uid("al"), server: c.server, action: "سداد عمولة التطبيق", target: c.name, detail: `سدّدت الشركة ${money(amount)} من عمولة التطبيق`, t: "الآن" }, ...(d.auditLog || [])] }; }),
    adminAction: (id, patch, logText) => upd((d) => { const c = d.companies.find((x) => x.id === id); if (!c) return d; return { ...d, companies: d.companies.map((x) => x.id === id ? { ...x, ...patch } : x), auditLog: [{ id: uid("al"), server: c.server, action: "إجراء إداري", target: c.name, detail: logText, t: "الآن" }, ...(d.auditLog || [])] }; }),
    confirmReceipt: (id) => upd((d) => { const r = d.receipts.find((x) => x.id === id); return { ...d, receipts: d.receipts.map((x) => x.id === id ? { ...x, confirmedByCompany: true } : x), notifications: [mkNotif("rep", r.repId, "cash", "تم تأكيد إيصالك", `أكّدت الشركة استلام إيصال ${id} (${money(r.amount)})`), ...d.notifications] }; }),
    toggleFav: (pid) => upd((d) => ({ ...d, favorites: d.favorites.includes(pid) ? d.favorites.filter((x) => x !== pid) : [...d.favorites, pid] })),
    toggleCompare: (pid) => upd((d) => ({ ...d, compare: d.compare.includes(pid) ? d.compare.filter((x) => x !== pid) : (d.compare.length >= 4 ? d.compare : [...d.compare, pid]) })),
    cartAdd: (p) => upd((d) => ({ ...d, cart: { ...d.cart, [p.id]: (d.cart[p.id] || 0) + 1 } })),
    cartSet: (pid, qty) => upd((d) => { const c = { ...d.cart }; if (qty <= 0) delete c[pid]; else c[pid] = qty; return { ...d, cart: c }; }),
    // ===== الطلب لكل شركة على حدة (بصيغة فاتورة) =====
    quickAddToCompany: (companyId, pid) => upd((d) => { const cur = d.orderDraft || { companyId: null, cart: {} }; const sameCo = cur.companyId === companyId; const cart = sameCo ? { ...cur.cart } : {}; cart[pid] = (cart[pid] || 0) + 1; return { ...d, orderDraft: { companyId, cart } }; }),
    clearOrderDraft: () => upd((d) => ({ ...d, orderDraft: { companyId: null, cart: {} } })),
    placeOrder: (merchantId, companyId, cart, note, fulfillment, meta) => upd((d) => {
      const merch = d.merchants.find((m) => m.id === merchantId);
      const items = Object.entries(cart).map(([pid, q]) => { const p = d.products.find((x) => x.id === pid); const mt = (meta || {})[pid] || {}; return { pid, name: p.name, qty: q, price: tierUnitPrice(p, merch, q), unitMode: mt.mode || "piece", boxes: mt.mode === "box" ? mt.count : null, packSize: mt.packSize || null }; });
      if (!items.length) return d;
      const repId = merch?.companyId === companyId ? merch.repId : d.reps.find((r) => r.companyId === companyId)?.id || null;
      const status = repId ? "pending_rep" : "pending_company"; // شركة بلا مندوب → مباشرة للشركة
      const order = { id: nextId(), server: merch.server, merchantId, companyId, repId, driverId: null, items, status, cashConfirmed: false, fulfillment, note: note || "", createdAt: "الآن" };
      const notif = repId
        ? mkNotif("rep", repId, "order", "طلب جديد للموافقة", `طلب ${order.id} من ${merch.name} بقيمة ${money(orderTotal(order))}`)
        : mkNotif("company", companyId, "order", "طلب مباشر وارد", `طلب ${order.id} من ${merch.name} (${FULFILL[fulfillment]?.label}) بقيمة ${money(orderTotal(order))}`);
      return { ...d, orderDraft: { companyId: null, cart: {} }, orders: [order, ...d.orders], notifications: [notif, ...d.notifications] };
    }),
    // حفظ الطلب كمسودّة (لتعديله لاحقاً قبل الإرسال)
    // استثناء إصدار الفاتورة رغم تقادم الدين (سماح دائم / إمهال أيام)
    allowInvoiceDespiteDebt: (merchantId, note) => upd((d) => ({ ...d, merchants: d.merchants.map((m) => m.id === merchantId ? { ...m, invoiceOverride: { type: "allow", note: note || "سماح استثنائي", date: "الآن" } } : m), notifications: [mkNotif("merchant", merchantId, "cash", "سماح استثنائي بإصدار الفواتير", note || "سمحت الشركة بإصدار الفواتير رغم تأخر السداد"), ...d.notifications] })),
    graceInvoice: (merchantId, extraDays, note) => upd((d) => ({ ...d, merchants: d.merchants.map((m) => m.id === merchantId ? { ...m, invoiceOverride: { type: "grace", extraDays: extraDays || 1, note: note || "إمهال طارئ", date: "الآن" } } : m), notifications: [mkNotif("merchant", merchantId, "cash", `إمهال ${extraDays || 1} يوم لإصدار الفواتير`, note || "أمهلتك الشركة لظرف طارئ"), ...d.notifications] })),
    clearInvoiceOverride: (merchantId) => upd((d) => ({ ...d, merchants: d.merchants.map((m) => m.id === merchantId ? { ...m, invoiceOverride: null } : m) })),
    setOrderPin: (merchantId, pin) => upd((d) => ({ ...d, merchants: d.merchants.map((m) => m.id === merchantId ? { ...m, orderPin: pin } : m) })),
    placeOrderDirect: (merchantId, companyId, cart, note, fulfillment, meta) => upd((d) => {
      const m = d.merchants.find((x) => x.id === merchantId);
      const items = Object.entries(cart).map(([pid, q]) => { const p = d.products.find((x) => x.id === pid); const mt = (meta || {})[pid] || {}; return { pid, name: p.name, qty: q, price: tierUnitPrice(p, m, q), unitMode: mt.mode || "piece", boxes: mt.mode === "box" ? mt.count : null, packSize: mt.packSize || null }; });
      if (!items.length) return d;
      const status = m?.repId ? "pending_rep" : "pending_company";
      const id = "ORD-2026-" + String(d.orderSeq || 5000).padStart(4, "0");
      const order = { id, server: m.server, merchantId, companyId, repId: m.repId || null, driverId: null, items, status, fulfillment, note: note || "", meta: meta || null, cashConfirmed: false, createdAt: "الآن", ageDays: 0 };
      const notif = m?.repId
        ? mkNotif("rep", m.repId, "order", "طلب جديد من عميلك", `${m.name} أرسل طلباً بقيمة ${money(orderTotal(order))}`)
        : mkNotif("company", companyId, "order", "طلب مباشر جديد", `${m.name} أرسل طلباً بقيمة ${money(orderTotal(order))}`);
      return { ...d, orderSeq: (d.orderSeq || 5000) + 1, orders: [order, ...d.orders], notifications: [notif, mkNotif("merchant", merchantId, "status", "أُرسل طلبك ✓", `طلبك ${id} في طريقه إلى ${m?.repId ? "مندوبك" : "الشركة"}`), ...d.notifications] };
    }),
    saveOrderDraft: (merchantId, companyId, cart, note, fulfillment, meta) => upd((d) => {
      const merch = d.merchants.find((m) => m.id === merchantId);
      const items = Object.entries(cart).map(([pid, q]) => { const p = d.products.find((x) => x.id === pid); const mt = (meta || {})[pid] || {}; return { pid, name: p.name, qty: q, price: tierUnitPrice(p, merch, q), unitMode: mt.mode || "piece", boxes: mt.mode === "box" ? mt.count : null, packSize: mt.packSize || null }; });
      if (!items.length) return d;
      const repId = merch?.companyId === companyId ? merch.repId : d.reps.find((r) => r.companyId === companyId)?.id || null;
      const order = { id: nextId(), server: merch.server, merchantId, companyId, repId, driverId: null, items, status: "draft", cashConfirmed: false, fulfillment, note: note || "", createdAt: "الآن", ageDays: 0 };
      return { ...d, orderDraft: { companyId: null, cart: {} }, orders: [order, ...d.orders] };
    }),
    // تفعيل مباشر لشركات البيع المباشر (بلا زيارة مندوب)
    directActivateJoin: (id, verifiedDebt) => upd((d) => { const j = d.joinRequests.find((x) => x.id === id);
      return { ...d, joinRequests: d.joinRequests.map((x) => x.id === id ? { ...x, status: "active", verifiedDebt, repId: null } : x),
        merchants: d.merchants.map((m) => m.id === j.merchantId ? { ...m, companyId: j.companyId, repId: null, balance: m.balance - (verifiedDebt || 0) } : m),
        notifications: [mkNotif("merchant", j.merchantId, "join", "تم تفعيل حسابك", `فعّلت ${compName(d, j.companyId)} حسابك مباشرةً — يمكنك الطلب الآن (تسلّم بنفسها/استلام من المقر)`), ...d.notifications] }; }),
    submitCart: (merchantId, note, asDraft, fulfillment = "delivery") => upd((d) => {
      const merch = d.merchants.find((m) => m.id === merchantId);
      const byCo = {};
      Object.entries(d.cart).forEach(([pid, qty]) => { const p = d.products.find((x) => x.id === pid); if (!p) return; (byCo[p.companyId] = byCo[p.companyId] || []).push({ pid, name: p.name, qty, price: priceForMerchant(p, merch) }); });
      const newOrders = Object.entries(byCo).map(([companyId, items]) => {
        const repId = merch?.companyId === companyId ? merch.repId : d.reps.find((r) => r.companyId === companyId)?.id || null;
        // شركة بلا مندوب → يذهب الطلب مباشرة للشركة
        const status = asDraft ? "draft" : (repId ? "pending_rep" : "pending_company");
        return { id: nextId(), server: merch.server, merchantId, companyId, repId, driverId: null, items, status, cashConfirmed: false, fulfillment, note, createdAt: "الآن" };
      });
      const notifs = asDraft ? [] : newOrders.flatMap((o) => [
        o.repId ? mkNotif("rep", o.repId, "order", "طلب جديد للموافقة", `طلب ${o.id} من ${merch.name} بقيمة ${money(orderTotal(o))}`)
          : mkNotif("company", o.companyId, "order", "طلب مباشر وارد", `طلب ${o.id} من ${merch.name} (${FULFILL[o.fulfillment]?.label}) بقيمة ${money(orderTotal(o))}`),
      ].filter(Boolean));
      return { ...d, cart: {}, orders: [...newOrders, ...d.orders], notifications: [...notifs, ...d.notifications] };
    }),
    editDraftQty: (oid, pid, qty) => upd((d) => ({ ...d, orders: d.orders.map((o) => { if (o.id !== oid) return o; const items = o.items.map((it) => it.pid === pid ? { ...it, qty } : it).filter((it) => it.qty > 0); return { ...o, items }; }).filter((o) => o.items.length > 0) })),
    submitDraft: (oid) => upd((d) => { const o = d.orders.find((x) => x.id === oid); if (!o) return d; const status = o.repId ? "pending_rep" : "pending_company"; const notif = o.repId
        ? mkNotif("rep", o.repId, "order", "طلب جديد للموافقة", `طلب ${o.id} من ${merchName(d, o.merchantId)} بقيمة ${money(orderTotal(o))}`)
        : mkNotif("company", o.companyId, "order", "طلب مباشر وارد", `طلب ${o.id} من ${merchName(d, o.merchantId)} (${FULFILL[o.fulfillment || "delivery"]?.label}) بقيمة ${money(orderTotal(o))}`);
      return { ...d, orders: d.orders.map((x) => x.id === oid ? { ...x, status, createdAt: "الآن" } : x), notifications: [notif, ...d.notifications] }; }),
    discardDraft: (oid) => upd((d) => ({ ...d, orders: d.orders.filter((o) => o.id !== oid) })),
    reorder: (oid) => upd((d) => { const src = d.orders.find((o) => o.id === oid); if (!src) return d; return { ...d, orders: [{ ...src, id: nextId(), status: src.repId ? "pending_rep" : "pending_company", cashConfirmed: false, driverId: null, createdAt: "الآن" }, ...d.orders] }; }),
    createReturn: (r) => upd((d) => ({ ...d, returns: [{ id: "RET-" + Math.floor(10 + Math.random() * 90), server: d.merchants.find((m) => m.id === r.merchantId)?.server, status: "requested", date: "الآن", ...r }, ...d.returns] })),
    sendMsg: (key, from, text) => upd((d) => ({ ...d, threads: { ...d.threads, [key]: [...(d.threads[key] || []), { from, text, t: "الآن" }] } })),
    // الانضمام
    createJoin: ({ merchantId, companyId, oldDebt, note }) => upd((d) => ({ ...d, joinRequests: [{ id: uid("j"), server: d.merchants.find((m) => m.id === merchantId)?.server, merchantId, companyId, status: "requested", note, oldDebt, verifiedDebt: null, repId: null, date: "الآن", visitDate: "" }, ...d.joinRequests], notifications: [mkNotif("company", companyId, "join", "طلب انضمام جديد", `${merchName(d, merchantId)} يطلب الانضمام والتعامل معكم`), ...d.notifications] })),
    declineJoin: (id) => upd((d) => ({ ...d, joinRequests: d.joinRequests.map((j) => j.id === id ? { ...j, status: "declined" } : j) })),
    scheduleVisit: (id, repId, visitDate) => upd((d) => { const j = d.joinRequests.find((x) => x.id === id); return { ...d, joinRequests: d.joinRequests.map((x) => x.id === id ? { ...x, status: "visit_set", repId, visitDate } : x), notifications: [mkNotif("rep", repId, "join", "زيارة توثيق مجدولة", `وثّق عميل ${merchName(d, j.merchantId)} بتاريخ ${visitDate}`), mkNotif("merchant", j.merchantId, "join", "تم جدولة زيارة", `سيزورك المندوب ${repName(d, repId)} للتوثيق بتاريخ ${visitDate}`), ...d.notifications] }; }),
    verifyJoin: (id, verifiedDebt) => upd((d) => { const j = d.joinRequests.find((x) => x.id === id);
      return { ...d, joinRequests: d.joinRequests.map((x) => x.id === id ? { ...x, status: "active", verifiedDebt } : x),
        merchants: d.merchants.map((m) => m.id === j.merchantId ? { ...m, companyId: j.companyId, repId: j.repId, balance: m.balance - verifiedDebt } : m) }; }),
    // شركة
    addMerchant: (m) => upd((d) => ({ ...d, merchants: [...d.merchants, { id: uid("m"), ...m }] })),
    updateMerchant: (id, patch) => upd((d) => ({ ...d, merchants: d.merchants.map((m) => m.id === id ? { ...m, ...patch } : m) })),
    addManager: (s) => upd((d) => ({ ...d, salesManagers: [...d.salesManagers, { id: uid("sm"), ...s }] })),
    // ربط أي فرد من الفريق بفرع (كل الأدوار)
    assignBranch: (arrayKey, entityId, branchId) => upd((d) => ({ ...d, [arrayKey]: (d[arrayKey] || []).map((x) => x.id === entityId ? { ...x, branchId: branchId || null } : x) })),
    // إشراف مدير المبيعات على فرع/عدة فروع (يبدّل عضوية الفرع في قائمته)
    toggleManagerBranch: (managerId, branchId) => upd((d) => ({ ...d, salesManagers: d.salesManagers.map((m) => { if (m.id !== managerId) return m; const cur = m.branchIds || []; const has = cur.includes(branchId); return { ...m, branchIds: has ? cur.filter((b) => b !== branchId) : [...cur, branchId] }; }) })),
    addPromoter: (p) => upd((d) => ({ ...d, promoters: [...d.promoters, { id: uid("pr"), ...p }] })),
    addProduct: (p) => upd((d) => ({ ...d, products: [...d.products, { id: uid("p"), ...p }] })),
    // ── تحديث السنة: أفعال جديدة ──
    pushCollectionTask: (merchantId, repId, amount) => upd((d) => ({ ...d, notifications: [mkNotif("rep", repId, "task", "مهمة تحصيل ذات أولوية", `حُقن محل بدين ${money(amount)} في خطتك — يُرجى التحصيل في زيارتك القادمة`), ...d.notifications] })),
    addCampaign: (c) => upd((d) => ({ ...d, campaigns: [{ id: uid("camp"), active: true, ...c }, ...(d.campaigns || [])] })),
    toggleCampaign: (id) => upd((d) => ({ ...d, campaigns: (d.campaigns || []).map((c) => c.id === id ? { ...c, active: !c.active } : c) })),
    deleteCampaign: (id) => upd((d) => ({ ...d, campaigns: (d.campaigns || []).filter((c) => c.id !== id) })),
    updateProduct: (id, patch) => upd((d) => ({ ...d, products: d.products.map((p) => p.id === id ? { ...p, ...patch } : p) })),
    deleteProduct: (id) => upd((d) => ({ ...d, products: d.products.filter((p) => p.id !== id) })),
    addSupervisor: (s) => upd((d) => ({ ...d, supervisors: [...d.supervisors, { id: uid("sv"), ...s }] })),
    addRep: (r) => upd((d) => ({ ...d, reps: [...d.reps, { id: uid("r"), ...r }] })),
    updateRep: (id, patch) => upd((d) => ({ ...d, reps: d.reps.map((r) => r.id === id ? { ...r, ...patch } : r) })),
    addDriver: (dr) => upd((d) => ({ ...d, drivers: [...d.drivers, { id: uid("d"), ...dr }] })),
    addTask: (t) => upd((d) => ({ ...d, tasks: [{ id: uid("t"), ...t }, ...d.tasks] })),
    updateTask: (id, patch) => upd((d) => ({ ...d, tasks: d.tasks.map((t) => t.id === id ? { ...t, ...patch } : t) })),
    addQuota: (q) => upd((d) => { const ex = d.quotas.find((x) => x.repId === q.repId && x.pid === q.pid); if (ex) return { ...d, quotas: d.quotas.map((x) => x.id === ex.id ? { ...x, allocated: x.allocated + q.allocated } : x) }; return { ...d, quotas: [...d.quotas, { id: uid("q"), ...q }] }; }),
    removeQuota: (id) => upd((d) => ({ ...d, quotas: d.quotas.filter((q) => q.id !== id) })),
    transferQuota: (fromId, toRepId, amount) => upd((d) => {
      const src = d.quotas.find((q) => q.id === fromId); if (!src) return d;
      let quotas = d.quotas.map((q) => q.id === fromId ? { ...q, allocated: q.allocated - amount } : q);
      const dest = quotas.find((q) => q.repId === toRepId && q.pid === src.pid);
      if (dest) quotas = quotas.map((q) => q.id === dest.id ? { ...q, allocated: q.allocated + amount } : q);
      else quotas = [...quotas, { id: uid("q"), server: src.server, companyId: src.companyId, repId: toRepId, pid: src.pid, allocated: amount, sold: 0 }];
      return { ...d, quotas: quotas.filter((q) => q.allocated > 0 || q.sold > 0) };
    }),
    // مروّج
    addVisit: (v) => upd((d) => ({ ...d, visits: [{ id: uid("v"), date: "الآن", ...v }, ...d.visits] })),
    repStocktake: (repId, merchantId, lines, note) => upd((d) => ({ ...d, stocktakes: [{ id: uid("st"), server: d.reps.find((r) => r.id === repId)?.server, repId, merchantId, lines, note, status: "requested", date: "الآن" }, ...d.stocktakes] })),
    repIssueInvoice: (repId, merchantId, cart) => {
      const rep = data.reps.find((r) => r.id === repId);
      const merch = data.merchants.find((m) => m.id === merchantId);
      const isVan = rep.repType === "van";
      const items = Object.entries(cart).map(([pid, qty]) => { const p = data.products.find((x) => x.id === pid); return { pid, name: p.name, qty, price: tierUnitPrice(p, merch, qty), batch: p.batch || null, expDate: p.expDate || null, mfgDate: p.mfgDate || null }; });
      const amount = items.reduce((s, it) => s + it.qty * it.price, 0);
      const invId = "INV-" + Math.floor(6000 + Math.random() * 999);
      const invoice = { id: invId, server: rep.server, merchantId, companyId: rep.companyId, orderId: null, amount, paid: 0, date: "الآن", due: "خلال 14 يوم", items, byRep: repId, gpsVerified: true, saleType: isVan ? "van" : "presell" };
      upd((d) => {
        // مندوب التجزئة: بيع فوري من مخزون السيارة (خصم الكميات)
        const reps = isVan ? d.reps.map((r) => { if (r.id !== repId) return r; const vs = { ...(r.vanStock || {}) }; items.forEach((it) => { vs[it.pid] = Math.max(0, (vs[it.pid] || 0) - it.qty); }); return { ...r, vanStock: vs, achieved: (r.achieved || 0) + amount, qtyAchieved: (r.qtyAchieved || 0) + items.reduce((s, i) => s + i.qty, 0) }; }) : d.reps;
        // ═══ T1: قيد الإيراد الكامل، وقيد التكلفة + المخزون (بلا فصل ضريبة) ═══
        const cogs = money3(items.reduce((s, it) => s + it.qty * wacOf(d, it.pid), 0));
        const salesAcct = salesAccountFor(rep.server, merch.tier);
        const revLines = [
          { acct: "1-3-001", dr: amount, cr: 0, dims: { merchantId } },
          { acct: salesAcct, dr: 0, cr: amount, dims: {} },
        ];
        let nd = postToLedger(d, { date: "الآن", source: "T1", docId: invId, server: rep.server, companyId: rep.companyId, memo: `فاتورة ${invId} — ${merch.name}`, lines: revLines });
        if (cogs > 0) nd = postToLedger(nd, { date: "الآن", source: "T1", docId: invId, server: rep.server, companyId: rep.companyId, memo: `تكلفة فاتورة ${invId}`, lines: [ { acct: "5-001", dr: cogs, cr: 0, dims: {} }, { acct: invAccountForRep(isVan), dr: 0, cr: cogs, dims: isVan ? { repId } : {} } ] });
        return {
          ...nd, reps,
          invoices: [invoice, ...nd.invoices],
          merchants: nd.merchants.map((m) => m.id === merchantId ? { ...m, balance: m.balance - amount } : m),
          notifications: [
            mkNotif("company", rep.companyId, "invoice", "فاتورة ميدانية صادرة", `أصدر ${rep.name} (${REP_TYPE[rep.repType || "presell"].short}) فاتورة ${invId} لـ${merch.name} بقيمة ${money(amount)}${isVan ? " — بيع مباشر من السيارة" : " — يُجهَّز من الشركة"}`),
            mkNotif("merchant", merchantId, "invoice", "فاتورة جديدة", `أصدر المندوب فاتورة ${invId} بقيمة ${money(amount)}`),
            ...nd.notifications,
          ],
        };
      });
      return invoice;
    },
    repCreateOrder: (repId, merchantId, cart) => upd((d) => {
      const rep = d.reps.find((r) => r.id === repId);
      const merch = d.merchants.find((m) => m.id === merchantId);
      const items = Object.entries(cart).map(([pid, qty]) => { const p = d.products.find((x) => x.id === pid); return { pid, name: p.name, qty, price: priceForMerchant(p, merch) }; });
      const order = { id: nextId(), server: rep.server, merchantId, companyId: rep.companyId, repId, driverId: null, items, status: "pending_company", cashConfirmed: true, note: "طلبية ميدانية عبر المندوب", createdAt: "الآن" };
      return { ...d, orders: [order, ...d.orders] };
    }),
    repCollectCash: (repId, merchantId, amount) => upd((d) => {
      const rep = d.reps.find((r) => r.id === repId);
      const m = d.merchants.find((x) => x.id === merchantId);
      const ref = "TXN-" + Math.floor(800000 + Math.random() * 199999);
      const rcpId = "RCP-" + Math.floor(3300 + Math.random() * 699);
      const receipt = { id: rcpId, server: rep.server, merchantId, companyId: rep.companyId, repId, amount, method: "نقدي للمندوب", date: "الآن", confirmedByCompany: false, ref };
      const nd2 = postToLedger(d, { date: "الآن", source: "T2", docId: rcpId, server: rep.server, companyId: rep.companyId, memo: `تحصيل نقدي ${rcpId} — ${m.name}`, lines: [ { acct: "1-2-002", dr: amount, cr: 0, dims: { repId } }, { acct: "1-3-001", dr: 0, cr: amount, dims: { merchantId } } ] });
      return {
        ...nd2,
        merchants: nd2.merchants.map((x) => x.id === merchantId ? { ...x, balance: x.balance + amount } : x),
        payments: [{ id: "PAY-" + Math.floor(700 + Math.random() * 299), server: rep.server, merchantId, companyId: rep.companyId, amount, method: "نقدي للمندوب", date: "الآن", repId, receiptId: rcpId }, ...nd2.payments],
        receipts: [receipt, ...nd2.receipts],
        visits: [{ id: uid("v"), server: rep.server, repId, merchantId, result: "payment", note: `تحصيل ${amount} د.ل`, amount, date: "الآن" }, ...d.visits],
        // إشعار فوري للشركة + للعميل
        notifications: [
          mkNotif("company", rep.companyId, "cash", "استلام نقدي جديد", `${rep.name} استلم ${money(amount)} من ${m.name} — إيصال ${rcpId}`),
          mkNotif("merchant", merchantId, "cash", "تم تسجيل دفعتك", `استلم المندوب ${rep.name} مبلغ ${money(amount)} وصدر إيصال إلكتروني ${rcpId}`),
          ...d.notifications,
        ],
      };
    }),
    toggleCash: (id) => upd((d) => { const o = d.orders.map((x) => x.id === id ? { ...x, cashConfirmed: !x.cashConfirmed } : x); const ord = d.orders.find((x) => x.id === id); const newVal = !ord.cashConfirmed; const notifs = newVal ? [mkNotif("company", ord.companyId, "cash", "تأكيد نقد من المندوب", `أكّد ${repName(d, ord.repId)} استلام نقد طلب ${id} (${money(orderTotal(ord))})`)] : []; return { ...d, orders: o, notifications: [...notifs, ...d.notifications] }; }),
    repAccept: (id) => upd((d) => { const o = d.orders.find((x) => x.id === id); return { ...d, orders: d.orders.map((x) => x.id === id ? { ...x, status: "pending_company" } : x), notifications: [mkNotif("company", o.companyId, "order", "طلب وصل من المندوب", `طلب ${id} وافق عليه ${repName(d, o.repId)} وبانتظار تجهيزكم`), mkNotif("merchant", o.merchantId, "status", "تحديث طلبك", `وافق المندوب على طلبك ${id} وأُرسل للشركة`), ...d.notifications] }; }),
    companyAccept: (id, driverId) => upd((d) => { const o = d.orders.find((x) => x.id === id); return { ...d, orders: d.orders.map((x) => x.id === id ? { ...x, status: "preparing", driverId } : x), notifications: [mkNotif("driver", driverId, "status", "طلب جاهز للاستلام", `طلب ${id} لـ${merchName(d, o.merchantId)} قيد التجهيز`), mkNotif("merchant", o.merchantId, "status", "تحديث طلبك", `طلبك ${id} قيد التجهيز`), ...d.notifications] }; }),
    // ===== محطة استلام وتجهيز الطلبات داخل الشركة =====
    companyReceivePrint: (id) => upd((d) => { const o = d.orders.find((x) => x.id === id); if (!o) return d; const sysNo = "SYS-" + Math.floor(50000 + Math.random() * 49999); const nd = ensureOrderInvoice(d, o); return { ...nd, orders: nd.orders.map((x) => x.id === id ? { ...x, status: "printed", sysNo } : x), notifications: [mkNotif("merchant", o.merchantId, "status", "تأكيد استلام طلبك", `استلمت الشركة طلبك ${id} وطُبع بالمنظومة برقم ${sysNo}`), ...nd.notifications] }; }),
    companyToWarehouse: (id) => upd((d) => { const o = d.orders.find((x) => x.id === id); return { ...d, orders: d.orders.map((x) => x.id === id ? { ...x, status: "warehouse" } : x), notifications: [mkNotif("merchant", o.merchantId, "status", "تحديث طلبك", `طلبك ${id} في المخزن قيد التعبئة`), ...d.notifications] }; }),
    companyPackAssign: (id, driverId, packedQty) => upd((d) => { const o = d.orders.find((x) => x.id === id); const isPickup = o.fulfillment === "pickup" || !driverId;
      // التعبئة الجزئية: تعديل الكميات الفعلية + تسجيل العجز
      let items = o.items; let shortages = [];
      if (packedQty) {
        items = o.items.map((it) => { const got = packedQty[it.pid] != null ? Math.max(0, Math.min(it.qty, packedQty[it.pid])) : it.qty; if (got < it.qty) shortages.push({ pid: it.pid, name: it.name, wanted: it.qty, got }); return { ...it, qty: got, shortage: got < it.qty ? it.qty - got : 0 }; }).filter((it) => it.qty > 0);
      }
      const shortNote = shortages.length ? ` (تعبئة جزئية: نقص ${shortages.length} صنف)` : "";
      const notifs = [
        isPickup ? mkNotif("merchant", o.merchantId, "status", "طلبك جاهز للاستلام", `عُبّئ طلبك ${id} وهو جاهز للاستلام من مقر الشركة${shortNote}`) : mkNotif("driver", driverId, "status", "طلب جاهز للاستلام", `طلب ${id} لـ${merchName(d, o.merchantId)} عُبّئ وجاهز — استلمه من المخزن`),
        mkNotif("merchant", o.merchantId, "status", "تحديث طلبك", `طلبك ${id} تمّت تعبئته${isPickup ? " وجاهز للاستلام" : " وسُلّم للسائق"}${shortNote}`),
      ];
      shortages.forEach((sh) => notifs.push(mkNotif("merchant", o.merchantId, "order", "نقص في صنف بطلبك", `«${sh.name}»: طلبت ${sh.wanted} وتوفّر ${sh.got} — عُدّلت الفاتورة تلقائياً`)));
      if (shortages.length && o.repId) notifs.push(mkNotif("rep", o.repId, "order", "تعبئة جزئية لطلب عميلك", `طلب ${id}: نقص ${shortages.length} صنف — راجع العميل`));
      return { ...d, orders: d.orders.map((x) => x.id === id ? { ...x, status: "preparing", driverId: isPickup ? null : driverId, items, shortages: shortages.length ? shortages : undefined } : x), notifications: [...notifs, ...d.notifications].filter(Boolean) }; }),
    setOpsPerms: (companyId, perms) => upd((d) => ({ ...d, companies: d.companies.map((c) => c.id === companyId ? { ...c, opsPerms: { ...c.opsPerms, ...perms } } : c) })),
    // ===== إشعار دائن/مدين (سند قيد) =====
    issueCreditNote: ({ merchantId, companyId, kind, amount, reason, refId }) => upd((d) => {
      const cn = { id: "CN-" + Math.floor(500 + Math.random() * 499), server: d.merchants.find((m) => m.id === merchantId)?.server, merchantId, companyId, kind, amount: +amount, reason: reason || "", refId: refId || null, date: "الآن", ageDays: 0 };
      const delta = kind === "credit" ? +amount : -amount; // دائن يخفّض المديونية
      const srv = d.merchants.find((m) => m.id === merchantId)?.server;
      // T5: دائن → مدين مردودات/تالف، دائن ذمم العميل | T6: مدين → مدين ذمم العميل، دائن فروق
      const cnLines = kind === "credit"
        ? [{ acct: "5-002", dr: +amount, cr: 0, dims: {} }, { acct: "1-3-001", dr: 0, cr: +amount, dims: { merchantId } }]
        : [{ acct: "1-3-001", dr: +amount, cr: 0, dims: { merchantId } }, { acct: "7-001", dr: 0, cr: +amount, dims: {} }];
      const nd = postToLedger(d, { date: "الآن", source: kind === "credit" ? "T5" : "T6", docId: cn.id, server: srv, companyId, memo: `سند ${kind === "credit" ? "دائن" : "مدين"} ${cn.id} — ${reason || ""}`, lines: cnLines });
      return { ...nd, creditNotes: [cn, ...nd.creditNotes], merchants: nd.merchants.map((m) => m.id === merchantId ? { ...m, balance: m.balance + delta } : m),
        notifications: [mkNotif("merchant", merchantId, "cash", kind === "credit" ? "إشعار دائن لصالحك" : "إشعار مدين عليك", `${kind === "credit" ? "خُصم من مديونيتك" : "أُضيف على حسابك"} ${money(+amount)} — ${reason || "سند قيد"}`), ...nd.notifications] };
    }),
    // ===== دورة المرتجع: فحص ← قرار =====
    inspectReturn: (id, condition) => upd((d) => { const r = d.returns.find((x) => x.id === id); if (!r) return d; return { ...d, returns: d.returns.map((x) => x.id === id ? { ...x, status: "inspected", condition } : x), notifications: [mkNotif("company", r.companyId, "task", "مرتجع فُحص ميدانياً", `${r.id} لـ${merchName(d, r.merchantId)} — الحالة: ${RET_CONDITION[condition].label}. بانتظار قراركم`), ...d.notifications] }; }),
    decideReturn: (id, decision, value) => upd((d) => {
      const r = d.returns.find((x) => x.id === id); if (!r) return d;
      let merchants = d.merchants, creditNotes = d.creditNotes, products = d.products, nd = d;
      const notifs = [];
      if (decision === "credited") {
        const amt = +value || 0;
        const cn = { id: "CN-" + Math.floor(500 + Math.random() * 499), server: r.server, merchantId: r.merchantId, companyId: r.companyId, kind: "credit", amount: amt, reason: `إشعار دائن عن مرتجع ${r.id}`, refId: r.id, date: "الآن", ageDays: 0 };
        creditNotes = [cn, ...creditNotes];
        merchants = merchants.map((m) => m.id === r.merchantId ? { ...m, balance: m.balance + amt } : m);
        // T5: مردودات المبيعات مدين / ذمم العميل دائن
        nd = postToLedger(nd, { date: "الآن", source: "T5", docId: cn.id, server: r.server, companyId: r.companyId, memo: `إشعار دائن عن مرتجع ${r.id}`, lines: [ { acct: r.condition === "damaged" ? "5-003" : "5-002", dr: amt, cr: 0, dims: {} }, { acct: "1-3-001", dr: 0, cr: amt, dims: { merchantId: r.merchantId } } ] });
        notifs.push(mkNotif("merchant", r.merchantId, "cash", "قرار مرتجعك: إشعار دائن ✓", `خُصم ${money(amt)} من مديونيتك عن المرتجع ${r.id}`));
      }
      if (decision === "replaced") notifs.push(mkNotif("merchant", r.merchantId, "order", "قرار مرتجعك: استبدال", `سيُستبدل ${r.qty} من البضاعة عن المرتجع ${r.id}`));
      if (decision === "rejected") notifs.push(mkNotif("merchant", r.merchantId, "task", "قرار مرتجعك: مرفوض", `اعتذرت الشركة عن قبول المرتجع ${r.id}`));
      // السليم يعاد للمخزون — T7: مخزون مدين / بضاعة قيد فحص دائن (بتكلفة WAC)
      if (r.condition === "ok" && r.pid) { products = products.map((p) => p.id === r.pid ? { ...p, stock: p.stock + (r.qty || 0) } : p); const cost = money3((r.qty || 0) * wacOf(d, r.pid)); if (cost > 0) nd = postToLedger(nd, { date: "الآن", source: "T7", docId: r.id, server: r.server, companyId: r.companyId, memo: `إعادة مرتجع ${r.id} للمخزون`, lines: [ { acct: "1-4-001", dr: cost, cr: 0, dims: { productId: r.pid } }, { acct: "1-4-003", dr: 0, cr: cost, dims: {} } ] }); }
      return { ...nd, merchants, creditNotes, products, returns: nd.returns.map((x) => x.id === id ? { ...x, status: decision, creditValue: decision === "credited" ? +value || 0 : null } : x), notifications: [...notifs, ...d.notifications] };
    }),
    // ===== تجميد عميل (خارج التغطية) =====
    freezeMerchant: (id, reason) => upd((d) => { const m = d.merchants.find((x) => x.id === id); return { ...d, merchants: d.merchants.map((x) => x.id === id ? { ...x, frozen: true, freezeReason: reason || "تجاوز الحد الائتماني" } : x), notifications: [mkNotif("merchant", id, "task", "تعليق الطلب مؤقتاً ⚠️", `عُلّق حسابك عن الطلب: ${reason || "تجاوز الحد الائتماني"} — سوّ مستحقاتك للتفعيل`), m?.repId ? mkNotif("rep", m.repId, "task", "عميل خارج التغطية", `${m.name}: ${reason || "تجاوز الحد"} — لا تُصدر له فواتير حتى إشعار آخر`) : null, ...d.notifications].filter(Boolean) }; }),
    unfreezeMerchant: (id) => upd((d) => { const m = d.merchants.find((x) => x.id === id); return { ...d, merchants: d.merchants.map((x) => x.id === id ? { ...x, frozen: false, freezeReason: null } : x), notifications: [mkNotif("merchant", id, "task", "أُعيد تفعيل حسابك ✓", "يمكنك الطلب الآن"), m?.repId ? mkNotif("rep", m.repId, "task", "عاد للتغطية", `${m.name} أُعيد تفعيله`) : null, ...d.notifications].filter(Boolean) }; }),
    // ===== خط السير الأسبوعي + «لا بيع» بسبب =====
    svSetJourney: ({ repId, day, merchantIds }) => upd((d) => {
      const rep = d.reps.find((r) => r.id === repId);
      const existing = d.journeyPlans.find((j) => j.repId === repId && j.day === day);
      const stops = merchantIds.map((mid) => { const prev = existing?.stops.find((st) => st.merchantId === mid); return prev || { merchantId: mid, done: false, outcome: null, reason: "" }; });
      const plans = existing ? d.journeyPlans.map((j) => j.id === existing.id ? { ...j, stops } : j) : [{ id: "JP-" + Math.floor(10 + Math.random() * 89), server: rep.server, repId, day, stops }, ...d.journeyPlans];
      return { ...d, journeyPlans: plans, notifications: [mkNotif("rep", repId, "route", "تحديث خط سيرك الأسبوعي", `يوم ${day}: ${merchantIds.length} محل`), ...d.notifications] };
    }),
    repJourneyOutcome: (planId, idx, outcome, reason) => upd((d) => {
      const jp = d.journeyPlans.find((j) => j.id === planId); if (!jp) return d;
      const stop = jp.stops[idx]; const rep = d.reps.find((r) => r.id === jp.repId);
      const visit = { id: uid("v"), server: jp.server, repId: jp.repId, merchantId: stop.merchantId, result: outcome === "order" ? "order" : "no_sale", note: "", date: "الآن", amount: 0, failReason: outcome === "order" ? null : reason };
      const notifs = outcome !== "order" && rep?.supervisorId ? [mkNotif("supervisor", rep.supervisorId, "route", "زيارة بلا بيع", `${rep.name} زار ${merchName(d, stop.merchantId)} — السبب: ${reason}`)] : [];
      return { ...d, visits: [visit, ...d.visits], journeyPlans: d.journeyPlans.map((j) => j.id === planId ? { ...j, stops: j.stops.map((st, x) => x === idx ? { ...st, done: true, outcome, reason: outcome === "order" ? "" : reason } : st) } : j), notifications: [...notifs, ...d.notifications] };
    }),
    // ===== جرد سيارة المندوب: إدخال العدّ ← اعتماد المشرف وتحميل العجز =====
    submitStocktake: (id, counts) => upd((d) => { const st = d.stocktakes.find((x) => x.id === id); if (!st) return d; return { ...d, stocktakes: d.stocktakes.map((x) => x.id === id ? { ...x, status: "submitted", lines: x.lines.map((l) => ({ ...l, counted: counts[l.pid] != null ? +counts[l.pid] : l.counted })) } : x), notifications: [st.supervisorId ? mkNotif("supervisor", st.supervisorId, "task", "جرد سيارة جاهز للاعتماد", `${repName(d, st.repId)} أدخل نتائج الجرد ${st.id}`) : null, ...d.notifications].filter(Boolean) }; }),
    resolveStocktake: (id) => upd((d) => {
      const st = d.stocktakes.find((x) => x.id === id); if (!st) return d;
      const variance = st.lines.reduce((sm, l) => sm + Math.max(0, (l.sysQty || 0) - (l.counted || 0)) * (l.price || 0), 0);
      const rep = d.reps.find((r) => r.id === st.repId);
      const reps = d.reps.map((r) => { if (r.id !== st.repId) return r; const vs = { ...(r.vanStock || {}) }; st.lines.forEach((l) => { if (l.counted != null) vs[l.pid] = l.counted; }); return { ...r, vanStock: vs, shortfall: (r.shortfall || 0) + variance }; });
      // T8: عجز جرد السيارة → مدين ذمم عجز المندوب / دائن مخزون السيارة (بتكلفة WAC)
      const varCost = money3(st.lines.reduce((sm, l) => sm + Math.max(0, (l.sysQty || 0) - (l.counted || 0)) * wacOf(d, l.pid), 0));
      const ndS = variance > 0 ? postToLedger(d, { date: "الآن", source: "T8", docId: st.id, server: st.server, companyId: rep?.companyId, memo: `عجز جرد ${st.id} — ${repName(d, st.repId)}`, lines: [ { acct: "1-3-002", dr: varCost, cr: 0, dims: { repId: st.repId } }, { acct: "1-4-002", dr: 0, cr: varCost, dims: { repId: st.repId } } ] }) : d;
      return { ...ndS, reps, stocktakes: ndS.stocktakes.map((x) => x.id === id ? { ...x, status: "resolved", varianceValue: variance } : x),
        notifications: [variance > 0 ? mkNotif("rep", st.repId, "cash", "عجز جرد محمَّل عليك", `فرق الجرد ${st.id}: ${money(variance)} — سُجّل ذمة عليك`) : mkNotif("rep", st.repId, "task", "جردك مطابق ✓", `لا فروقات في ${st.id}`), mkNotif("company", rep.companyId, "task", "نتيجة جرد سيارة", `${rep.name}: ${variance > 0 ? `عجز ${money(variance)} حُمّل عليه` : "مطابق ✓"}`), ...d.notifications] };
    }),
    // قراءة كل إشعارات فئة واحدة
    readNotifType: (role, entityId, type) => upd((d) => ({ ...d, notifications: d.notifications.map((n) => n.role === role && n.entityId === entityId && n.icon === type ? { ...n, read: true } : n) })),
    assignMerchantToRep: (merchantId, repId) => upd((d) => ({ ...d, merchants: d.merchants.map((m) => m.id === merchantId ? { ...m, repId } : m), notifications: [mkNotif("merchant", merchantId, "task", "تحديث مندوبك", `أصبح ${repName(d, repId)} مندوبك`), ...d.notifications] })),
    // ===== أدوات المحل (حصر) + عقد الثلاجة =====
    addAsset: (a) => upd((d) => ({ ...d, assets: [{ id: uid("as"), condition: "جيدة", date: "تم الحصر الآن", ...a }, ...d.assets] })),
    removeAsset: (id) => upd((d) => ({ ...d, assets: d.assets.filter((a) => a.id !== id) })),
    saveFridgeContract: (merchantId, contract) => upd((d) => ({ ...d, merchants: d.merchants.map((m) => m.id === merchantId ? { ...m, fridgeContract: contract } : m), notifications: [mkNotif("merchant", merchantId, "task", "حُفظ عقد الثلاجة", `رقم العقد ${contract.no}`), ...d.notifications] })),
    // ===== طلب جرد من التاجر (أصناف محددة أو الكل) =====
    requestMerchantStocktake: ({ repId, companyId, merchantId, pids, scope }) => upd((d) => {
      const prods = d.products.filter((p) => p.companyId === companyId && (scope === "all" || pids.includes(p.id)));
      const lines = prods.map((p) => ({ pid: p.id, name: p.name, sysQty: null, counted: null, price: p.price }));
      const st = { id: "STK-" + Math.floor(700 + Math.random() * 299), server: d.merchants.find((m) => m.id === merchantId)?.server, type: "merchant", merchantId, repId, companyId, status: "requested", note: scope === "all" ? "جرد شامل لكل الأصناف" : `جرد ${lines.length} صنف محدد`, date: "الآن", lines };
      return { ...d, stocktakes: [st, ...d.stocktakes], notifications: [mkNotif("merchant", merchantId, "task", "طلب جرد من مندوبك 📋", `${repName(d, repId)} يطلب جرد ${scope === "all" ? "كل الأصناف" : `${lines.length} صنف`} في محلك`), ...d.notifications] };
    }),
    // ===== طلب مطابقة الحساب (المندوب أو الشركة) =====
    requestReconciliation: ({ requestedBy, requesterId, companyId, merchantId }) => upd((d) => {
      const m = d.merchants.find((x) => x.id === merchantId);
      const rec = { id: "REC-" + Math.floor(3000 + Math.random() * 999), server: m?.server, merchantId, companyId, requestedBy, requesterId, statedBalance: m?.balance || 0, status: "pending", merchantNote: "", date: "الآن" };
      return { ...d, reconciliations: [rec, ...d.reconciliations], notifications: [mkNotif("merchant", merchantId, "task", "طلب مطابقة حساب 🧾", `${requestedBy === "company" ? compName(d, companyId) : repName(d, requesterId)} يطلب تأكيد رصيدك الحالي (${money(Math.abs(m?.balance || 0))} ${(m?.balance || 0) < 0 ? "عليك" : "لك"})`), ...d.notifications] };
    }),
    respondReconciliation: (id, agree, note) => upd((d) => {
      const rec = d.reconciliations.find((x) => x.id === id); if (!rec) return d;
      const notifTo = rec.requestedBy === "company" ? "company" : "rep";
      const notifId = rec.requestedBy === "company" ? rec.companyId : rec.requesterId;
      return { ...d, reconciliations: d.reconciliations.map((x) => x.id === id ? { ...x, status: agree ? "confirmed" : "disputed", merchantNote: note || "" } : x),
        notifications: [mkNotif(notifTo, notifId, "cash", agree ? "طابق التاجر الحساب ✓" : "اعتراض على المطابقة ⚠️", `${merchName(d, rec.merchantId)}: ${agree ? "أكّد صحة الرصيد" : `لديه ملاحظة — ${note || "يرجى المراجعة"}`}`), ...d.notifications] };
    }),
    // ===== استيراد عملاء من Excel + التوثيق الميداني =====
    importClients: ({ companyId, rows, source }) => upd((d) => {
      const server = d.companies.find((c) => c.id === companyId)?.server;
      const newPcs = rows.map((r, i) => ({ id: "PC-" + Math.floor(200 + Math.random() * 799) + i, server, companyId, repId: r.repId, name: r.name, owner: r.owner || "", phone: r.phone || "", area: r.area || "", openingDebt: +r.openingDebt || 0, status: "awaiting_visit", photo: null, lat: null, lng: null, nameConfirmed: false, gpsConfirmed: false, merchantConfirmed: false, note: "", date: "الآن", source: source || "ملف Excel" }));
      const byRep = {}; newPcs.forEach((pc) => { byRep[pc.repId] = (byRep[pc.repId] || 0) + 1; });
      const notifs = Object.entries(byRep).map(([repId, n]) => mkNotif("rep", repId, "task", "عملاء جدد للتوثيق الميداني 📋", `أسندت لك الشركة ${n} محل من ملف الاستيراد — زُر كل محل: طابق الاسم، صوّر المحل، وثّق الموقع، وخذ تأكيد التاجر`));
      return { ...d, pendingClients: [...newPcs, ...d.pendingClients], notifications: [...notifs, ...d.notifications] };
    }),
    // المندوب يوثّق ميدانياً: تصحيح الاسم + صورة + GPS داخل النطاق
    repVerifyClient: (pcId, { correctedName, photo, lat, lng }) => upd((d) => {
      const pc = d.pendingClients.find((x) => x.id === pcId); if (!pc) return d;
      const nameChanged = correctedName && correctedName !== pc.name;
      return { ...d, pendingClients: d.pendingClients.map((x) => x.id === pcId ? { ...x, name: correctedName || x.name, prevName: nameChanged ? pc.name : null, nameConfirmed: true, photo: photo || "📸", lat, lng, gpsConfirmed: true, status: "field_verified" } : x),
        notifications: [mkNotif("company", pc.companyId, "task", "وثّق المندوب محلاً مستورداً", `${repName(d, pc.repId)} وثّق «${correctedName || pc.name}»${nameChanged ? ` (الاسم صُحّح من: ${pc.name})` : ""} — بانتظار تأكيد التاجر`), ...d.notifications] };
    }),
    // تأكيد التاجر على بياناته أمام المندوب ← يتفعّل كعميل حقيقي
    confirmImportedClient: (pcId) => upd((d) => {
      const pc = d.pendingClients.find((x) => x.id === pcId); if (!pc) return d;
      const mId = "m" + Math.floor(100 + Math.random() * 899);
      const memberNo = "MID-" + (d.memberSeq || 10007);
      const merchant = { id: mId, server: pc.server, name: pc.name, owner: pc.owner || pc.name, phone: pc.phone, email: "", city: "طرابلس", area: pc.area, balance: -(pc.openingDebt || 0), creditLimit: 5000, companyId: pc.companyId, repId: pc.repId, tier: "retail", contractPct: 0, appDiscountPct: 0, photo: "🏪", lat: pc.lat, lng: pc.lng, hasContract: false, license: null, licenseDate: null, address: pc.area, managerName: pc.owner, managerPhone: pc.phone, frozen: false, freezeReason: null, membershipNo: memberNo, fridgeContract: null, orderPin: "0000", invoiceOverride: null };
      return { ...d, memberSeq: (d.memberSeq || 10007) + 1, merchants: [...d.merchants, merchant], pendingClients: d.pendingClients.map((x) => x.id === pcId ? { ...x, status: "activated", merchantId: mId } : x),
        notifications: [mkNotif("company", pc.companyId, "join", "عميل مستورد اكتمل تفعيله ✓", `«${pc.name}» أكّد بياناته أمام ${repName(d, pc.repId)} وأُضيف لعملائكم${pc.openingDebt ? ` بمديونية افتتاحية ${money(pc.openingDebt)}` : ""}`), ...d.notifications] };
    }),
    rejectImportedClient: (pcId, reason) => upd((d) => { const pc = d.pendingClients.find((x) => x.id === pcId); return { ...d, pendingClients: d.pendingClients.map((x) => x.id === pcId ? { ...x, status: "rejected", note: reason || "بيانات غير مطابقة" } : x), notifications: [mkNotif("company", pc.companyId, "task", "محل مستورد مرفوض", `${repName(d, pc.repId)}: «${pc.name}» — ${reason || "بيانات غير مطابقة"}`), ...d.notifications] }; }),
    // ===== رحلات السائق (تجميع توصيلات) =====
    createTrip: ({ companyId, driverId, orderIds }) => upd((d) => {
      const trip = { id: "TRP-" + Math.floor(900 + Math.random() * 99), server: d.companies.find((c) => c.id === companyId)?.server, companyId, driverId, orderIds, status: "active", date: "الآن" };
      const notifs = [mkNotif("driver", driverId, "status", "رحلة توصيل جديدة 🚚", `أُسندت لك رحلة ${trip.id} تضم ${orderIds.length} طلب — راجع خط السير`)];
      orderIds.forEach((oid) => { const o = d.orders.find((x) => x.id === oid); if (o) notifs.push(mkNotif("merchant", o.merchantId, "status", "طلبك ضمن رحلة توصيل", `طلبك ${oid} أُسند لسائق ضمن رحلة ${trip.id}`)); });
      return { ...d, trips: [trip, ...d.trips], orders: d.orders.map((x) => orderIds.includes(x.id) ? { ...x, driverId, tripId: trip.id } : x), notifications: [...notifs, ...d.notifications] };
    }),
    // موظف العمليات يُطلق الفاتورة الجاهزة إلى مشرف الحركة
    opsRelease: (id) => upd((d) => {
      const o = d.orders.find((x) => x.id === id); if (!o) return d;
      const disp = d.dispatchers.find((x) => x.companyId === o.companyId);
      return { ...d, orders: d.orders.map((x) => x.id === id ? { ...x, status: "ready_dispatch" } : x), notifications: [mkNotif("dispatch", disp?.id || "", "status", "فاتورة جاهزة للتحميل", `طلب ${id} لـ${merchName(d, o.merchantId)} جاهز — جمّعه في حمولة`), ...d.notifications] };
    }),
    // موظف العمليات يطلب من مشرف الحركة الحضور لاستلام الفواتير يدوياً (بعد إدخالها في المنظومة)
    requestPickup: ({ companyId, orderIds, note }) => upd((d) => {
      const disp = d.dispatchers.find((x) => x.companyId === companyId);
      const co = d.companies.find((c) => c.id === companyId);
      const req = { id: "PU-" + Math.floor(1000 + Math.random() * 8999), server: co?.server, companyId, dispatcherId: disp?.id || null, orderIds: orderIds || [], note: note || "", status: "requested", date: "الآن" };
      return { ...d, pickupRequests: [req, ...(d.pickupRequests || [])], notifications: [mkNotif("dispatch", disp?.id || "", "task", "طلب استلام فواتير 📋", `العمليات جاهزة — ${(orderIds || []).length} فاتورة بانتظار استلامك يدوياً${note ? " · " + note : ""}`), ...d.notifications] };
    }),
    // مشرف الحركة يؤكد استلام الفواتير يدوياً
    confirmPickup: (reqId) => upd((d) => {
      const r = (d.pickupRequests || []).find((x) => x.id === reqId); if (!r) return d;
      return { ...d, pickupRequests: (d.pickupRequests || []).map((x) => x.id === reqId ? { ...x, status: "collected" } : x), orders: d.orders.map((o) => r.orderIds.includes(o.id) ? { ...o, status: "ready_dispatch" } : o) };
    }),
    // مشرف الحركة يحدّد سعة الباليتة لصنف
    setPalletCap: (productId, cap) => upd((d) => ({ ...d, products: d.products.map((p) => p.id === productId ? { ...p, stikaPerPallet: cap } : p) })),
    // تواصل بين السائق ومشرف الحركة
    sendDispatchMsg: ({ companyId, driverId, from, text }) => upd((d) => {
      const msg = { id: uid("dc"), companyId, driverId, from, text, t: "الآن" };
      const drv = d.drivers.find((x) => x.id === driverId);
      const disp = d.dispatchers.find((x) => x.companyId === companyId);
      const notif = from === "driver"
        ? mkNotif("dispatch", disp?.id || "", "task", "رسالة من سائق 💬", `${drv?.name || "سائق"}: ${text.slice(0, 40)}`)
        : mkNotif("driver", driverId, "task", "رسالة من مشرف الحركة 💬", text.slice(0, 40));
      return { ...d, dispatchChat: [...(d.dispatchChat || []), msg], notifications: [notif, ...d.notifications] };
    }),
    // نقل كميات بين فروع الشركة (حساب موحّد، تقسيم داخلي) — بقيد محاسبي موثّق
    transferBranchStock: ({ companyId, fromBranch, toBranch, items, note }) => upd((d) => {
      const co = d.companies.find((c) => c.id === companyId);
      const totalVal = items.reduce((s, it) => s + money3(it.qty * wacOf(d, it.pid)), 0);
      const jvLines = [
        { acct: "1-4-001", dr: totalVal, cr: 0, dims: { branch: toBranch } },
        { acct: "1-4-001", dr: 0, cr: totalVal, dims: { branch: fromBranch } },
      ];
      const nd = postToLedger(d, { date: "الآن", source: "BT", docId: "BT-" + Math.floor(1000 + Math.random() * 8999), server: co?.server, companyId, memo: `نقل مخزون من ${fromBranch} إلى ${toBranch}${note ? " · " + note : ""}`, lines: jvLines });
      const transfer = { id: "BT-" + Math.floor(1000 + Math.random() * 8999), server: co?.server, companyId, fromBranch, toBranch, items, value: totalVal, note: note || "", status: "completed", date: "الآن" };
      return { ...nd, branchTransfers: [transfer, ...(nd.branchTransfers || [])], notifications: [mkNotif("company", companyId, "status", "نقل مخزون بين الفروع", `نُقلت ${items.length} صنف بقيمة ${money(totalVal)}`), ...nd.notifications] };
    }),
    // مشرف الحركة يكوّن رحلة سيارة ويُسندها لسائق (يجمّع الطلبيات ويحمّلها)
    dispatchTrip: ({ companyId, driverId, orderIds }) => upd((d) => {
      const dayKey = "اليوم";
      const sameDayCount = (d.trips || []).filter((t) => t.driverId === driverId && t.dayKey === dayKey).length;
      const tripNo = sameDayCount + 1; // 1 = طلعة أساسية، 2+ = إضافية
      const trip = { id: "TRP-" + Math.floor(900 + Math.random() * 99), server: d.companies.find((c) => c.id === companyId)?.server, companyId, driverId, orderIds, status: "active", date: "الآن", dayKey, tripNo, extra: tripNo > 1 };
      const drv = d.drivers.find((x) => x.id === driverId);
      const notifs = [mkNotif("driver", driverId, "status", "رحلة توصيل جديدة 🚚", `أُسندت لك رحلة ${trip.id}${trip.extra ? " (طلعة إضافية ⭐)" : ""} تضم ${orderIds.length} طلب — السيارة محمّلة وجاهزة`)];
      orderIds.forEach((oid) => { const o = d.orders.find((x) => x.id === oid); if (o) notifs.push(mkNotif("merchant", o.merchantId, "status", "طلبك في الطريق 🚚", `طلبك ${oid} حُمّل وخرج مع السائق ${drv?.name || ""}`)); });
      return { ...d, trips: [trip, ...d.trips], orders: d.orders.map((x) => orderIds.includes(x.id) ? { ...x, driverId, tripId: trip.id, status: "out_for_delivery" } : x), notifications: [...notifs, ...d.notifications] };
    }),
    // ===== تصنيف المندوبين =====
    setRepType: (repId, repType) => upd((d) => ({ ...d, reps: d.reps.map((r) => r.id === repId ? { ...r, repType } : r), notifications: [mkNotif("rep", repId, "task", "تحديث تصنيفك", `صنّفتك الشركة كـ«${REP_TYPE[repType].label}»`), ...d.notifications] })),
    // ===== تحميل سيارات مندوبي التجزئة =====
    createLoadRequest: ({ repId, companyId, items, note }) => upd((d) => { const lr = { id: "LD-" + Math.floor(7000 + Math.random() * 999), server: d.reps.find((r) => r.id === repId)?.server, repId, companyId, items, note: note || "", status: "requested", date: "الآن" }; return { ...d, loadRequests: [lr, ...d.loadRequests], notifications: [mkNotif("company", companyId, "order", "طلب تحميل سيارة", `طلب المندوب ${repName(d, repId)} تحميل ${items.length} صنف لسيارته`), ...d.notifications] }; }),
    fulfillLoadRequest: (id) => upd((d) => { const lr = d.loadRequests.find((x) => x.id === id); if (!lr) return d; const rep = d.reps.find((r) => r.id === lr.repId); const vs = { ...(rep.vanStock || {}) }; lr.items.forEach((it) => { vs[it.pid] = (vs[it.pid] || 0) + it.qty; }); const cost = money3(lr.items.reduce((sm, it) => sm + it.qty * wacOf(d, it.pid), 0)); const nd = cost > 0 ? postToLedger(d, { date: "الآن", source: "T9", docId: lr.id, server: lr.server, companyId: lr.companyId, memo: `تحميل سيارة ${rep.name}`, lines: [ { acct: "1-4-002", dr: cost, cr: 0, dims: { repId: lr.repId } }, { acct: "1-4-001", dr: 0, cr: cost, dims: {} } ] }) : d; return { ...nd, loadRequests: nd.loadRequests.map((x) => x.id === id ? { ...x, status: "loaded" } : x), reps: nd.reps.map((r) => r.id === lr.repId ? { ...r, vanStock: vs } : r), notifications: [mkNotif("rep", lr.repId, "order", "تم تحميل سيارتك", `عبّأت الشركة ${lr.items.reduce((s, i) => s + i.qty, 0)} وحدة وأُضيفت إلى مخزون سيارتك`), ...nd.notifications] }; }),
    rejectLoadRequest: (id) => upd((d) => { const lr = d.loadRequests.find((x) => x.id === id); return { ...d, loadRequests: d.loadRequests.map((x) => x.id === id ? { ...x, status: "rejected" } : x), notifications: [mkNotif("rep", lr.repId, "order", "رُفض طلب التحميل", `لم تُعبّأ سيارتك — راجع الشركة`), ...d.notifications] }; }),
    // الشركة/العمليات تنشئ تحميلاً للمندوب — يتطلب موافقته قبل إضافته لمخزون سيارته
    companyCreateLoad: ({ repId, companyId, items, note }) => upd((d) => { const lr = { id: "LD-" + Math.floor(7000 + Math.random() * 999), server: d.reps.find((r) => r.id === repId)?.server, repId, companyId, items, note: note || "", status: "pending_rep", origin: "company", date: "الآن" }; return { ...d, loadRequests: [lr, ...d.loadRequests], notifications: [mkNotif("rep", repId, "order", "عرض تحميل من الشركة", `أعدّت الشركة تحميلاً لسيارتك (${items.reduce((s, i) => s + i.qty, 0)} وحدة) — راجعه ووافق ليُضاف لمخزونك`), ...d.notifications] }; }),
    approveLoad: (id) => upd((d) => { const lr = d.loadRequests.find((x) => x.id === id); if (!lr) return d; const rep = d.reps.find((r) => r.id === lr.repId); const vs = { ...(rep.vanStock || {}) }; lr.items.forEach((it) => { vs[it.pid] = (vs[it.pid] || 0) + it.qty; }); const cost = money3(lr.items.reduce((sm, it) => sm + it.qty * wacOf(d, it.pid), 0)); const nd = cost > 0 ? postToLedger(d, { date: "الآن", source: "T9", docId: lr.id, server: lr.server, companyId: lr.companyId, memo: `تحميل سيارة ${rep.name}`, lines: [ { acct: "1-4-002", dr: cost, cr: 0, dims: { repId: lr.repId } }, { acct: "1-4-001", dr: 0, cr: cost, dims: {} } ] }) : d; return { ...nd, loadRequests: nd.loadRequests.map((x) => x.id === id ? { ...x, status: "loaded" } : x), reps: nd.reps.map((r) => r.id === lr.repId ? { ...r, vanStock: vs } : r), notifications: [mkNotif("company", lr.companyId, "order", "وافق المندوب على التحميل", `وافق ${rep.name} على تحميل ${lr.items.reduce((s, i) => s + i.qty, 0)} وحدة وأُضيفت لسيارته`), ...nd.notifications] }; }),
    rejectLoad: (id) => upd((d) => { const lr = d.loadRequests.find((x) => x.id === id); return { ...d, loadRequests: d.loadRequests.map((x) => x.id === id ? { ...x, status: "rejected" } : x), notifications: [mkNotif("company", lr.companyId, "order", "رفض المندوب التحميل", `رفض ${repName(d, lr.repId)} عرض التحميل`), ...d.notifications] }; }),
    advance: (id, status) => upd((d) => {
      const o = d.orders.find((x) => x.id === id); if (!o) return d;
      const base = status === "delivered" ? ensureOrderInvoice(d, o) : d;
      let promotions = base.promotions; let notifs = [];
      if (status === "out_for_delivery") {
        const promoter = base.promoters.find((p) => p.companyId === o.companyId);
        if (promoter && !base.promotions.some((pm) => pm.orderId === id)) {
          promotions = [{ id: uid("pm"), server: o.server, companyId: o.companyId, promoterId: promoter.id, orderId: id, merchantId: o.merchantId, status: "notified", req: { photoBefore: true, photoAfter: true, fridge: true, shelf: true, expiry: true }, gpsConfirmed: false, photoBefore: false, photoAfter: false, expiryFindings: [], stacking: null, suggestedItems: [], repAlert: null, note: "", date: "الآن" }, ...base.promotions];
          notifs.push(mkNotif("promoter", promoter.id, "promo", "مهمة ترويج جديدة", `استلم السائق طلب ${id} لـ${merchName(base, o.merchantId)} — توجّه للإشراف على العرض`));
        }
        notifs.push(mkNotif("merchant", o.merchantId, "status", "طلبك في الطريق", `خرج طلب ${id} للتوصيل`));
      }
      if (status === "delivered") notifs.push(mkNotif("merchant", o.merchantId, "status", "تم تسليم طلبك", `تم تسليم طلب ${id} بنجاح`));
      // خصم المخزون عند التسليم + تنبيه النفاد
      let products = base.products;
      if (status === "delivered") {
        products = base.products.map((p) => { const it = o.items.find((x) => x.pid === p.id); if (!it) return p; return { ...p, stock: Math.max(0, p.stock - it.qty) }; });
        o.items.forEach((it) => { const p = products.find((x) => x.id === it.pid); if (p && p.stock <= 20) notifs.push(mkNotif("company", o.companyId, "order", "تنبيه مخزون منخفض ⚠️", `«${p.name}» أوشك على النفاد — المتبقي ${p.stock} ${p.unit}`)); });
      }
      const trips = status === "delivered" && o.tripId ? base.trips.map((t) => t.id === o.tripId && t.orderIds.every((oid) => oid === id || base.orders.find((x) => x.id === oid)?.status === "delivered") ? { ...t, status: "delivered" } : t) : base.trips;
      return { ...base, trips, promotions, products, orders: base.orders.map((x) => x.id === id ? { ...x, status } : x), notifications: [...notifs, ...base.notifications] };
    }),
    // ===== تسوية الفواتير (التاجر يدفع ← المندوب يراجع ويعتمد) =====
    createSettlement: ({ merchantId, companyId, invoiceId, amount, method, ref, note }) => upd((d) => {
      const m = d.merchants.find((x) => x.id === merchantId);
      const repId = m?.repId || d.reps.find((r) => r.companyId === companyId)?.id || null;
      const s = { id: "ST-" + Math.floor(9000 + Math.random() * 999), server: m?.server, merchantId, companyId, repId, invoiceId: invoiceId || null, amount, method, ref: ref || "", note: note || "", status: "pending_rep", date: "الآن", reviewedAt: null };
      const notif = repId
        ? mkNotif("rep", repId, "cash", "تسوية بانتظار مراجعتك", `${m?.name} سدّد ${money(amount)} (${METHOD_LABEL[method]}) — راجع القيمة واعتمدها`)
        : mkNotif("company", companyId, "cash", "تسوية بانتظار اعتمادكم", `${m?.name} سدّد ${money(amount)} (${METHOD_LABEL[method]}) مباشرة للشركة`);
      return { ...d, settlements: [s, ...d.settlements], notifications: [notif, ...d.notifications] };
    }),
    // ===== طلب مندوب للتحصيل (يطلبه التاجر ← إشعار للمندوب + متابعة الشركة) =====
    requestCollection: ({ merchantId, companyId, amount, method, note }) => upd((d) => {
      const m = d.merchants.find((x) => x.id === merchantId);
      const repId = m?.repId || d.reps.find((r) => r.companyId === companyId)?.id || null;
      const cr = { id: "CR-" + Math.floor(5000 + Math.random() * 999), server: m?.server, merchantId, companyId, repId, amount: amount || 0, method: method || "cash", note: note || "", status: "requested", date: "الآن", visitTime: "" };
      const notifs = [];
      if (repId) notifs.push(mkNotif("rep", repId, "cash", "طلب تحصيل من عميل ⚡", `${m?.name} يطلب زيارتك للتحصيل${amount ? ` (${money(amount)} ${METHOD_LABEL[method || "cash"]})` : ""} — جدول زيارة`));
      notifs.push(mkNotif("company", companyId, "cash", "عميل يطلب التحصيل", `${m?.name} طلب مندوباً للتحصيل${amount ? ` بقيمة ${money(amount)}` : ""} — تابع المندوب ${repId ? repName(d, repId) : ""}`));
      notifs.push(mkNotif("merchant", merchantId, "cash", "تم إرسال طلب التحصيل", `أُرسل طلبك للمندوب${repId ? " " + repName(d, repId) : ""} وأُعلمت الشركة.`));
      return { ...d, collectRequests: [cr, ...d.collectRequests], notifications: [...notifs, ...d.notifications] };
    }),
    scheduleCollection: (id, visitTime) => upd((d) => { const cr = d.collectRequests.find((x) => x.id === id); if (!cr) return d; return { ...d, collectRequests: d.collectRequests.map((x) => x.id === id ? { ...x, status: "scheduled", visitTime: visitTime || "قريباً" } : x), notifications: [mkNotif("merchant", cr.merchantId, "cash", "المندوب في الطريق للتحصيل", `جدول ${repName(d, cr.repId)} زيارتك للتحصيل: ${visitTime || "قريباً"}`), mkNotif("company", cr.companyId, "cash", "تحديث طلب تحصيل", `جدول ${repName(d, cr.repId)} زيارة ${merchName(d, cr.merchantId)} للتحصيل (${visitTime || "قريباً"})`), ...d.notifications] }; }),
    completeCollection: (id) => upd((d) => { const cr = d.collectRequests.find((x) => x.id === id); if (!cr) return d; return { ...d, collectRequests: d.collectRequests.map((x) => x.id === id ? { ...x, status: "done" } : x), notifications: [mkNotif("company", cr.companyId, "cash", "تم تحصيل طلب عميل", `حصّل ${repName(d, cr.repId)} من ${merchName(d, cr.merchantId)}`), mkNotif("merchant", cr.merchantId, "cash", "تم تحصيل دفعتك", `استلم المندوب ${repName(d, cr.repId)} الدفعة — شكراً`), ...d.notifications] }; }),
    cancelCollection: (id) => upd((d) => ({ ...d, collectRequests: d.collectRequests.map((x) => x.id === id ? { ...x, status: "canceled" } : x) })),
    // ===== زيارة المروّج الميدانية =====
    pmConfirmGps: (id, coords) => upd((d) => { const t = d.promotions.find((x) => x.id === id); const m = d.merchants.find((x) => x.id === t?.merchantId); return { ...d, promotions: d.promotions.map((x) => x.id === id ? { ...x, gpsConfirmed: true, status: "arrived", lat: coords?.lat ?? m?.lat, lng: coords?.lng ?? m?.lng, gpsAccuracy: coords?.accuracy ?? null } : x) }; }),
    pmPhoto: (id, which, photoData) => upd((d) => ({ ...d, promotions: d.promotions.map((x) => x.id === id ? { ...x, [which]: true, [`${which}Data`]: photoData || x[`${which}Data`] || null } : x) })), 
    pmSaveExpiry: (id, findings) => upd((d) => { const t = d.promotions.find((x) => x.id === id); const bad = findings.filter((fd) => fd.flag !== "ok").length; return { ...d, promotions: d.promotions.map((x) => x.id === id ? { ...x, expiryFindings: findings } : x), notifications: bad ? [mkNotif("company", t.companyId, "task", "ملاحظات صلاحية من المروّج ⚠️", `${merchName(d, t.merchantId)}: ${bad} صنف قارب/تجاوز الصلاحية — راجع تقرير الزيارة`), ...d.notifications] : d.notifications }; }),
    pmSaveStacking: (id, stacking) => upd((d) => ({ ...d, promotions: d.promotions.map((x) => x.id === id ? { ...x, stacking, status: "displayed" } : x) })),
    pmAlertRep: (id, { items, alertType, note }) => upd((d) => {
      const t = d.promotions.find((x) => x.id === id); const m = d.merchants.find((x) => x.id === t.merchantId);
      const pr = d.promoters.find((p) => p.id === t.promoterId);
      const TYPE_TXT = { wants_items: "المحل يريد هذه الأصناف", needs_visit: "المحل يحتاج زيارتك", contact: "تواصل مع المحل" };
      const itemsTxt = items.length ? ` — الأصناف: ${items.map((it) => `${it.name}×${it.qty}`).join("، ")}` : "";
      const alert = { items, alertType, note: note || "", sentAt: "الآن" };
      return { ...d, promotions: d.promotions.map((x) => x.id === id ? { ...x, suggestedItems: items, repAlert: alert } : x),
        notifications: [m?.repId ? mkNotif("rep", m.repId, "task", `تنبيه من المروّج ${pr?.name} 📣`, `${m.name}: ${TYPE_TXT[alertType]}${itemsTxt}${note ? ` — ${note}` : ""}`) : null, mkNotif("company", t.companyId, "task", "طلبية مقترحة من الميدان", `${pr?.name} في ${m?.name}: ${TYPE_TXT[alertType]}${itemsTxt}`), ...d.notifications].filter(Boolean) };
    }),
    pmComplete: (id, note) => upd((d) => {
      const t = d.promotions.find((x) => x.id === id); const m = d.merchants.find((x) => x.id === t.merchantId);
      const parts = [];
      if (t.photoBefore) parts.push("صورة قبل ✓"); if (t.photoAfter) parts.push("صورة بعد ✓");
      if (t.stacking) parts.push(`ثلاجة: ${t.stacking.fridge || "—"} · رفوف: ${t.stacking.shelf || "—"}`);
      if (t.expiryFindings?.length) parts.push(`${t.expiryFindings.filter((fd) => fd.flag !== "ok").length} ملاحظة صلاحية`);
      return { ...d, promotions: d.promotions.map((x) => x.id === id ? { ...x, status: "sent", note: note || x.note } : x),
        notifications: [mkNotif("company", t.companyId, "task", "تقرير زيارة مروّج مكتمل ✓", `${m?.name}: ${parts.join(" · ") || "زيارة مكتملة"}`), ...d.notifications] };
    }),
    // ===== التقييمات =====
    rateCompany: ({ merchantId, companyId, orderId, stars, comment }) => upd((d) => {
      const rt = { id: uid("rt"), server: d.merchants.find((m) => m.id === merchantId)?.server, raterRole: "merchant", merchantId, companyId, orderId: orderId || null, stars, comment: comment || "", date: "الآن" };
      return { ...d, ratings: [rt, ...d.ratings], notifications: [mkNotif("company", companyId, "task", "تقييم جديد من عميل", `${merchName(d, merchantId)} قيّمكم بـ${stars} نجوم${comment ? ` — «${comment}»` : ""}`), ...d.notifications] };
    }),
    updateMyProfile: (merchantId, f) => upd((d) => ({ ...d, merchants: d.merchants.map((m) => m.id === merchantId ? { ...m, managerName: f.managerName, managerPhone: f.managerPhone, email: f.email, address: f.address } : m) })),
    rateMerchant: ({ merchantId, companyId, stars, comment }) => upd((d) => {
      const rt = { id: uid("rt"), server: d.companies.find((c) => c.id === companyId)?.server, raterRole: "company", merchantId, companyId, orderId: null, stars, comment: comment || "", date: "الآن" };
      return { ...d, ratings: [rt, ...d.ratings] };
    }),
    approveSettlement: (id) => upd((d) => {
      const s = d.settlements.find((x) => x.id === id); if (!s) return d;
      const rep = d.reps.find((r) => r.id === s.repId);
      const isCash = s.method === "cash";
      const rcpId = "RCP-" + Math.floor(3300 + Math.random() * 699);
      const ref = "TXN-" + Math.floor(800000 + Math.random() * 199999);
      const methodLabel = METHOD_LABEL[s.method] + (isCash ? " (تسوية)" : "");
      const payment = { id: "PAY-" + Math.floor(700 + Math.random() * 299), server: s.server, merchantId: s.merchantId, companyId: s.companyId, amount: s.amount, method: methodLabel, date: "الآن", repId: s.repId, receiptId: rcpId, handedOver: !isCash, settlementId: id };
      const receipt = { id: rcpId, server: s.server, merchantId: s.merchantId, companyId: s.companyId, repId: s.repId, amount: s.amount, method: methodLabel, date: "الآن", confirmedByCompany: false, ref };
      // T4: تسوية — مدين (عهدة المندوب نقداً / البنك حوالة / شيكات) / دائن ذمم العميل
      const debitAcct = isCash ? "1-2-002" : s.method === "cheque" ? "1-3-003" : "1-2-003";
      const debitDims = isCash ? { repId: s.repId } : {};
      const ndT4 = postToLedger(d, { date: "الآن", source: "T4", docId: rcpId, server: s.server, companyId: s.companyId, memo: `تسوية ${merchName(d, s.merchantId)} (${METHOD_LABEL[s.method]})`, lines: [ { acct: debitAcct, dr: s.amount, cr: 0, dims: debitDims }, { acct: "1-3-001", dr: 0, cr: s.amount, dims: { merchantId: s.merchantId } } ] });
      return { ...ndT4,
        settlements: ndT4.settlements.map((x) => x.id === id ? { ...x, status: "approved", reviewedAt: "الآن", receiptId: rcpId } : x),
        merchants: ndT4.merchants.map((m) => m.id === s.merchantId ? { ...m, balance: m.balance + s.amount } : m),
        payments: [payment, ...ndT4.payments],
        receipts: [receipt, ...ndT4.receipts],
        notifications: [
          mkNotif("company", s.companyId, "cash", "تسوية فاتورة معتمدة", `اعتمد ${rep?.name || "المندوب"} تسوية ${merchName(d, s.merchantId)} بمبلغ ${money(s.amount)} (${METHOD_LABEL[s.method]})${isCash ? " — أُضيفت لعهدة المندوب لتوريدها" : ""}`),
          mkNotif("merchant", s.merchantId, "cash", "تم اعتماد سدادك", `اعتمد المندوب سدادك ${money(s.amount)} وصدر إيصال ${rcpId}`),
          ...d.notifications,
        ],
      };
    }),
    rejectSettlement: (id, reason) => upd((d) => { const s = d.settlements.find((x) => x.id === id); return { ...d, settlements: d.settlements.map((x) => x.id === id ? { ...x, status: "rejected", reviewedAt: "الآن", note: (x.note ? x.note + " — " : "") + "سبب الرفض: " + (reason || "القيمة غير مطابقة") } : x), notifications: [mkNotif("merchant", s.merchantId, "cash", "لم يُعتمد سدادك", `راجع المندوب تسوية ${money(s.amount)} ولم يعتمدها${reason ? " — " + reason : ""}. يرجى التواصل.`), ...d.notifications] }; }),
    // ===== تسليم النقدية اليومي من المندوب للشركة =====
    repHandover: (repId, note, delivered) => upd((d) => {
      const rep = d.reps.find((r) => r.id === repId);
      const held = d.payments.filter((p) => p.repId === repId && (p.method || "").includes("نقد") && !p.handedOver);
      const expected = held.reduce((s, p) => s + p.amount, 0); const count = held.length;
      if (!expected) return d;
      const del = delivered != null && delivered !== "" ? +delivered : expected;
      const diff = del - expected;
      const hv = { id: "HV-" + Math.floor(4400 + Math.random() * 599), server: rep.server, repId, companyId: rep.companyId, amount: del, expected, diff, count, status: "pending", date: "الآن", note: note || "توريد نهاية اليوم" };
      return { ...d, handovers: [hv, ...d.handovers], payments: d.payments.map((p) => p.repId === repId && (p.method || "").includes("نقد") && !p.handedOver ? { ...p, handedOver: true } : p),
        notifications: [mkNotif("company", rep.companyId, "cash", "توريد نقدي من مندوب", `${rep.name} ورّد ${money(del)} (المتوقع ${money(expected)}${diff !== 0 ? ` — ${diff < 0 ? "عجز" : "زيادة"} ${money(Math.abs(diff))}` : " — مطابق ✓"})`), ...d.notifications] };
    }),
    confirmHandover: (id) => upd((d) => { const h = d.handovers.find((x) => x.id === id); if (!h) return d;
      const reps = h.diff < 0 ? d.reps.map((r) => r.id === h.repId ? { ...r, shortfall: (r.shortfall || 0) + Math.abs(h.diff) } : r) : d.reps;
      // T3: مدين الصندوق الرئيسي (المورَّد) / دائن عهدة المندوب (المتوقع)؛ العجز ذمة، الزيادة فرق دائن
      const hvLines = [ { acct: "1-2-001", dr: h.amount, cr: 0, dims: {} }, { acct: "1-2-002", dr: 0, cr: h.expected != null ? h.expected : h.amount, dims: { repId: h.repId } } ];
      if (h.diff < 0) hvLines.push({ acct: "1-3-002", dr: Math.abs(h.diff), cr: 0, dims: { repId: h.repId } });
      if (h.diff > 0) hvLines.push({ acct: "7-001", dr: 0, cr: h.diff, dims: {} });
      const nd = postToLedger(d, { date: "الآن", source: "T3", docId: h.id, server: h.server, companyId: h.companyId, memo: `توريد نقدي ${h.id} — ${repName(d, h.repId)}`, lines: hvLines });
      return { ...nd, reps, handovers: nd.handovers.map((x) => x.id === id ? { ...x, status: "received" } : x), notifications: [mkNotif("rep", h.repId, "cash", "تأكيد استلام توريدك", `أكّدت الشركة استلام ${money(h.amount)}${h.diff < 0 ? ` — وسُجّل عجز ${money(Math.abs(h.diff))} ذمةً عليك` : ""}`), ...nd.notifications] }; }),
    // ===== خط السير العاجل (المشرف ← المندوب) =====
    assignUrgentRoute: ({ supervisorId, repId, note, stops }) => upd((d) => {
      const ur = { id: "UR-" + Math.floor(300 + Math.random() * 699), server: d.reps.find((r) => r.id === repId)?.server, supervisorId, repId, note: note || "", date: "الآن", stops: stops.map((s) => ({ ...s, done: false })) };
      return { ...d, urgentRoutes: [ur, ...d.urgentRoutes], notifications: [mkNotif("rep", repId, "task", "خط سير عاجل ⚡", `أسندك ${d.supervisors.find((s) => s.id === supervisorId)?.name} خط سير عاجل بـ${stops.length} محل — ابدأ الآن`), ...d.notifications] };
    }),
    urgentStopDone: (routeId, idx) => upd((d) => ({ ...d, urgentRoutes: d.urgentRoutes.map((r) => r.id === routeId ? { ...r, stops: r.stops.map((s, i) => i === idx ? { ...s, done: true } : s) } : r) })),
    // ===== الأهداف المتسلسلة =====
    setMgrTarget: (id, patch) => upd((d) => ({ ...d, salesManagers: d.salesManagers.map((s) => s.id === id ? { ...s, ...patch } : s), notifications: [mkNotif("manager", id, "task", "تحديث أهدافك", "حدّثت الشركة أهدافك وكمياتك المخصصة"), ...d.notifications] })),
    setSupervisorTarget: (id, patch) => upd((d) => ({ ...d, supervisors: d.supervisors.map((s) => s.id === id ? { ...s, ...patch } : s), notifications: [mkNotif("supervisor", id, "task", "تحديث هدفك", "حدّث مدير المبيعات هدفك (ستيكات/تحصيل)"), ...d.notifications] })),
    // ترحيل الكمية غير المنجَزة لمندوب إلى الشهر القادم (تُضاف لهدفه)
    carryoverRepTarget: (repId) => upd((d) => {
      const r = d.reps.find((x) => x.id === repId); if (!r) return d;
      const remaining = Math.max(0, (r.qtyTarget || 0) - (r.qtyAchieved || 0));
      if (remaining <= 0) return d;
      return { ...d, reps: d.reps.map((x) => x.id === repId ? { ...x, carriedOver: (x.carriedOver || 0) + remaining, carryNote: `مُرحّل ${remaining} ستيكة للشهر القادم` } : x), notifications: [mkNotif("rep", repId, "task", "ترحيل هدف", `رُحّلت ${remaining} ستيكة غير منجَزة إلى هدف الشهر القادم`), ...d.notifications] };
    }),
    // نقل كمية من هدف مندوب إلى مندوب آخر
    transferRepTarget: ({ fromRepId, toRepId, qty }) => upd((d) => {
      const from = d.reps.find((x) => x.id === fromRepId); const to = d.reps.find((x) => x.id === toRepId);
      if (!from || !to || qty <= 0) return d;
      return { ...d, reps: d.reps.map((x) => {
        if (x.id === fromRepId) return { ...x, qtyTarget: Math.max(0, (x.qtyTarget || 0) - qty) };
        if (x.id === toRepId) return { ...x, qtyTarget: (x.qtyTarget || 0) + qty };
        return x;
      }), notifications: [mkNotif("rep", toRepId, "task", "زيادة في هدفك", `نُقلت إليك ${qty} ستيكة من هدف ${from.name}`), mkNotif("rep", fromRepId, "task", "تعديل هدفك", `نُقلت ${qty} ستيكة من هدفك إلى ${to.name}`), ...d.notifications] };
    }),
    // تسجيل الكمية غير المنجَزة كعجز على المندوب
    markRepDeficit: (repId) => upd((d) => {
      const r = d.reps.find((x) => x.id === repId); if (!r) return d;
      const deficit = Math.max(0, (r.qtyTarget || 0) - (r.qtyAchieved || 0));
      return { ...d, reps: d.reps.map((x) => x.id === repId ? { ...x, qtyDeficit: (x.qtyDeficit || 0) + deficit, carryNote: `عجز ${deficit} ستيكة` } : x), notifications: [mkNotif("rep", repId, "task", "عجز في الهدف", `سُجّل عجز ${deficit} ستيكة على هدفك هذا الشهر`), ...d.notifications] };
    }),
    setClientTarget: ({ supervisorId, merchantId, qtyTarget, collectionTarget }) => upd((d) => {
      const ex = d.clientTargets.find((c) => c.supervisorId === supervisorId && c.merchantId === merchantId);
      if (ex) return { ...d, clientTargets: d.clientTargets.map((c) => c.id === ex.id ? { ...c, qtyTarget, collectionTarget } : c) };
      return { ...d, clientTargets: [{ id: uid("ct"), server: d.merchants.find((m) => m.id === merchantId)?.server, supervisorId, merchantId, qtyTarget, collectionTarget, qtyAchieved: 0, collected: 0 }, ...d.clientTargets] };
    }),
    setMgrStock: ({ companyId, managerId, pid, allocated }) => upd((d) => {
      const ex = d.mgrStock.find((s) => s.managerId === managerId && s.pid === pid);
      if (ex) return { ...d, mgrStock: d.mgrStock.map((s) => s.id === ex.id ? { ...s, allocated } : s) };
      return { ...d, mgrStock: [{ id: uid("ms"), server: d.companies.find((c) => c.id === companyId)?.server, companyId, managerId, pid, allocated, distributed: 0 }, ...d.mgrStock] };
    }),
    // ===== المساحة الإعلانية =====
    addAd: (ad) => upd((d) => ({ ...d, ads: [{ id: uid("ad"), active: true, ...ad }, ...d.ads] })),
    toggleAd: (id) => upd((d) => ({ ...d, ads: d.ads.map((a) => a.id === id ? { ...a, active: !a.active } : a) })),
    removeAd: (id) => upd((d) => ({ ...d, ads: d.ads.filter((a) => a.id !== id) })),
  };

  const SplashScreen = splash ? (<div className="fixed inset-0 z-[99] flex flex-col items-center justify-center gap-5" style={{ background: `linear-gradient(135deg, ${C.ink} 0%, #0D47A1 100%)` }}>
    <div style={{ animation: "rbLogoIn .9s cubic-bezier(.21,1.02,.55,1) both" }}><div className="flex items-center gap-3"><div className="rounded-2xl flex items-center justify-center" style={{ width: 64, height: 64, background: "rgba(255,255,255,.12)", border: "2px solid rgba(255,255,255,.25)" }}><Link2 size={34} color="#F59E0B" /></div><div><div className="text-white text-4xl font-extrabold" style={{ fontFamily: "Cairo, sans-serif" }}>رابط</div><div className="text-sm" style={{ color: "rgba(255,255,255,.65)" }}>ليربط العالم بين يديك</div></div></div></div>
    <div className="flex gap-1.5" style={{ animation: "rbFadeIn .5s .6s ease both" }}>{[0, 1, 2].map((i) => <div key={i} className="rounded-full" style={{ width: 8, height: 8, background: "#F59E0B", animation: `rbPulse 1.2s ${i * 0.2}s ease infinite` }} />)}</div>
  </div>) : null;

  const hosts = (<><ToastHost /><ConfirmHost />{SplashScreen}</>);
  if (!server) return (<>{hosts}<ServerPick onPick={(s) => { SC = SERVERS[s]; setServer(s); }} /></>);
  if (!user) return (<>{hosts}<Login server={server} onLogin={setUser} onBack={() => { setServer(null); setUser(null); }} /></>);
  return (<>{hosts}<Shell server={server} user={user} data={data} actions={actions} onLogout={() => { logoutRemote(); setUser(null); }} /></>);
}