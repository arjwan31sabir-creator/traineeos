import { useState, useEffect, useRef } from "react";
import { supabase } from "./supabaseClient";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";

const CLAUDE_KEY  = import.meta.env.VITE_CLAUDE_API_KEY;
const WORK_LAT    = 23.5896598;
const WORK_LNG    = 58.4125277;
const WORK_RADIUS = 1000;
const MAX_SIGNIN  = "09:00";
const PENALTY_PCT = 8.33;
const MANAGER_NAMES = ["Arjwan","Sultan","Mohammed","Sana","Osama"];

// ── TRANSLATIONS ────────────────────────────────────
const T = {
  en: {
    appName:"Huawei TechTrack", poweredBy:"Powered by Arjwan Sabir",
    signIn:"Sign In →", signingIn:"Signing in…", createAccount:"Create Trainee Account",
    email:"Email", password:"Password", confirmPassword:"Confirm Password",
    welcome:"Welcome! 🎉", completeProfile:"Complete your profile once to get started.",
    fullName:"Full Name", civilId:"Civil ID", dob:"Date of Birth", gender:"Gender",
    nationality:"Nationality", university:"University", department:"Department",
    gpa:"GPA", ojtEndDate:"OJT End Date", mentor:"Assigned Mentor", phone:"Phone Number",
    personalInfo:"👤 Personal Info", academicInfo:"🎓 Academic & OJT Info", contact:"📞 Contact",
    completeProfileBtn:"Complete Profile & Start →",
    hi:"Hi", onTime:"✅ On Time", pastDeadline:"⚠️ Past 9 AM",
    greeting:"👋 Greeting", reminder:"⏰ Reminder", signOut:"Sign out",
    attend:"Attend", weekly:"Weekly", sick:"Sick", goals:"Goals",
    signAttendance:"✅ Sign Attendance", date:"Date", signinTime:"Sign-in Time (auto)",
    confirmAttendance:"I confirm attendance today",
    penaltyWarning:"⚠️ Penalty Warning: Not marked as attended — 8.33% deduction",
    timeLocked:"Time Locked", deadlineWas:"Deadline was 9:00 AM.",
    penaltyWillApply:"8.33% penalty will apply.",
    locationVerification:"📍 Location Verification",
    mustBeWithin:"Must be within", ofWorkplace:"of workplace.",
    checkingLocation:"📍 Checking…", locationVerified:"✅ Location Verified",
    verifyLocation:"📍 Verify My Location",
    submitAttendance:"✅ Submit Attendance", saving:"Saving…",
    signOutBtn:"🚪 Sign Out", recordTimeOut:"🚪 Record Time Out",
    signedOutAt:"Signed out at", timeRecorded:"Time out recorded",
    myStatus:"📋 My Status", payment:"Payment", laptop:"Laptop",
    paid:"✅ Paid", unpaid:"❌ Unpaid", received:"✅ Received",
    notYet:"❌ Not Yet", returned:"🔄 Returned",
    returnedLaptop:"🔄 I Returned My Laptop to HR",
    receivedLaptop:"✅ I Received My Laptop",
    ojtEndDateLabel:"🎓 OJT End Date", daysRemaining:"days remaining", ojtEnded:"OJT period ended",
    excuseRequired:"⚠️ Excuse Required", excuseType:"Excuse Type",
    selectReason:"Select reason…", traffic:"🚗 Road Traffic", limitReached:"(LIMIT REACHED)",
    medical:"🏥 Medical Emergency", family:"👨‍👩‍👧 Family Emergency", other:"📋 Other",
    description:"Description", proofPhoto:"Proof Photo", required:"(Required)", optional:"(Optional)",
    uploadProof:"Upload Proof", changePhoto:"Change Photo",
    currentWeek:"Current Week", sundayThursday:"Sunday — Thursday",
    submitted:"Submitted", pending:"Pending", reportSubmitted:"✅ Report Submitted",
    edit:"✏️ Edit", weekSummary:"Week Summary", talentNotes:"🌟 Talent Notes",
    weeklyKpi:"Weekly KPI Score", outOf:"out of 100", aiAnalysis:"🤖 AI Analysis",
    thisWeekTasks:"📅 This Week's Tasks",
    writeEachDay:"Write what you did each day. AI will analyze and score your report.",
    weeklyTasksLabel:"Weekly Tasks", proofPhotoOpt:"📸 Proof Photo (optional)",
    submitWeekly:"📅 Submit Weekly Report", aiAnalyzing:"🤖 AI Analyzing…",
    sickLeavePolicy:"ℹ️ Sick Leave Policy",
    first2days:"First 2 days (48 hours) — No penalty ✅",
    after48:"After 48 hours — 8.33% penalty per day ⚠️",
    mySickLeaves:"🏥 My Sick Leaves", submitSickLeave:"+ Submit",
    cancel:"✕ Cancel", startDate:"Start Date", endDate:"End Date",
    reason:"Reason", medicalProof:"📸 Medical Proof (Recommended)",
    submitSickBtn:"🏥 Submit Sick Leave", submitting:"Submitting…",
    noSickLeaves:"No sick leaves recorded.", total:"Total",
    freeDays:"Free Days", penaltyDays:"Penalty Days", nopenalty:"✅ No penalty",
    totalGoals:"Total", inProgress:"In Progress", completed:"Completed", overdue:"Overdue",
    addGoal:"+ Add Goal", newGoal:"🎯 New Goal", kraCategory:"KRA Category",
    goalTitle:"Goal Title", target:"Target", unit:"Unit", dueDate:"Due Date",
    setGoal:"🎯 Set Goal", noGoals:"No goals yet. Tap + Add Goal!",
    update:"Update", complete:"✅ Complete",
    allTrainees:"👥 All Trainees", searchPlaceholder:"🔍 Search by name, department, civil ID...",
    trainees:"Trainees", live:"Live", analytics:"Analytics", okr:"OKR", access:"Access",
    exportExcel:"📊 Export Excel", attendanceSheets:"📋 Attendance Sheets",
    back:"← Back", reactivate:"✅ Reactivate", transfer:"🔄 Transfer",
    inactive:"⏸ Inactive", drop:"🔴 Drop",
    timeline:"Timeline", reports:"Reports", penalties:"Penalties",
    paymentTab:"Payment", laptopTab:"Laptop",
    activityTimeline:"Activity Timeline", noActivity:"No activity yet.",
    editProfile:"Edit Profile", saveChanges:"Save Changes",
    joiningDate:"Joining Date", quittingDate:"Quitting Date", quittingReason:"Quitting Reason",
    paymentStatus:"💰 Payment Status", paymentNotes:"Payment Notes",
    savePayment:"💾 Save Payment Status",
    laptopStatus:"💻 Laptop Status", markReceived:"✅ Mark as Received",
    markReturned:"🔄 Mark as Returned to HR", reset:"❌ Reset to Not Received",
    laptopSerial:"Laptop Serial Number", dateReceived:"Date Received",
    dateReturned:"Date Returned to HR", saveLaptop:"💾 Save Laptop Status",
    goalsTab:"Goals", noGoalsSet:"No goals set yet.",
    liveFeed:"📡 Live Feed", waitingSignins:"Waiting for sign-ins…",
    signedIn:"Signed In", absent:"Absent", onTimeLabel:"On Time", late:"Late",
    kpiLeaderboard:"🏆 KPI Leaderboard", noData:"No data yet.",
    atRisk:"⚠️ At Risk", avgKpi:"Avg KPI",
    active:"Active", transferred:"Transferred", dropped:"Dropped",
    ocrTracking:"🎯 OKR Tracking", addOkr:"+ Add", addOkrTitle:"Add OKR",
    objective:"Objective", keyResult:"Key Result",
    onTrack:"On Track", behind:"Behind",
    accessLog:"🔐 Access Log", loggedInAs:"Logged in as:",
    refresh:"🔄 Refresh", totalActions:"Total Actions", logins:"Logins",
    profileViews:"Profile Views", profileEdits:"Profile Edits",
    statusChanges:"Status Changes", paymentUpdates:"Payment Updates",
    laptopUpdates:"Laptop Updates", exports:"Exports",
    accessByPerson:"👤 Access by Person", fullActivityLog:"📋 Full Activity Log",
    noAccessYet:"No named access yet.", selectNameToTrack:"Select your name on the greeting popup to start tracking.",
    noActivityYet:"No activity yet.", loginToTrack:"Log in with your name to start tracking all actions.",
    managementHub:"Management Hub", whoAccessing:"👤 Who is accessing today?",
    selectNameGreeting:"Select your name to receive a personalized greeting",
    enterDashboard:"Enter Dashboard 🚀", letsGo:"Let's Go! 🚀",
    todayMotivation:"🈲 Today's Motivation", haveNiceDay:"🌟 Have a nice day!",
    transferTrainee:"Transfer Trainee", currentDept:"Current dept:",
    currentMentor:"Current mentor:", newDepartment:"New Department",
    newMentor:"New Mentor (optional)", confirmTransfer:"🔄 Confirm",
    dropTrainee:"Drop Trainee", confirmDrop:"🔴 Confirm Drop",
    exportAttendance:"Export Attendance Sheets",
    selectMonthYear:"Select month and year to generate attendance sheets for all active trainees",
    month:"Month", year:"Year", export:"📋 Export",
    attendanceRemark:"Attendance/Remark", attendanceSummary:"Attendance Summary",
    info:"Info", days:"Days", present:"Present",
    sickLeave:"Sick Leave", weekend:"Weekend", holiday:"Holiday",
    visitSite:"Visit Site", totalDaysMonth:"Total days in month",
    traineeSignature:"Trainee Signature:", hrSignature:"HR Manager Signature:",
    approvedBy:"Approved By PM/Manager:",
    noTrainees:"No trainees match your search.",
    of:"of",
  },
  ar: {
    appName:"هواوي تيك تراك", poweredBy:"من تطوير أروان صابر",
    signIn:"تسجيل الدخول →", signingIn:"جاري تسجيل الدخول…", createAccount:"إنشاء حساب متدرب",
    email:"البريد الإلكتروني", password:"كلمة المرور", confirmPassword:"تأكيد كلمة المرور",
    welcome:"أهلاً وسهلاً! 🎉", completeProfile:"أكمل ملفك الشخصي مرة واحدة للبدء.",
    fullName:"الاسم الكامل", civilId:"رقم الهوية", dob:"تاريخ الميلاد", gender:"الجنس",
    nationality:"الجنسية", university:"الجامعة", department:"القسم",
    gpa:"المعدل التراكمي", ojtEndDate:"تاريخ انتهاء التدريب", mentor:"المشرف المعين", phone:"رقم الهاتف",
    personalInfo:"👤 المعلومات الشخصية", academicInfo:"🎓 المعلومات الأكاديمية والتدريبية", contact:"📞 التواصل",
    completeProfileBtn:"إتمام الملف الشخصي والبدء →",
    hi:"مرحباً", onTime:"✅ في الوقت المحدد", pastDeadline:"⚠️ بعد الساعة 9 صباحاً",
    greeting:"👋 التحية", reminder:"⏰ تذكير", signOut:"تسجيل الخروج",
    attend:"الحضور", weekly:"أسبوعي", sick:"إجازة مرضية", goals:"الأهداف",
    signAttendance:"✅ تسجيل الحضور", date:"التاريخ", signinTime:"وقت الدخول (تلقائي)",
    confirmAttendance:"أؤكد حضوري اليوم",
    penaltyWarning:"⚠️ تحذير: لم يتم تأكيد الحضور — خصم 8.33%",
    timeLocked:"الوقت محجوب", deadlineWas:"الموعد النهائي كان 9:00 صباحاً.",
    penaltyWillApply:"سيتم تطبيق خصم 8.33%.",
    locationVerification:"📍 التحقق من الموقع",
    mustBeWithin:"يجب أن تكون ضمن", ofWorkplace:"من مكان العمل.",
    checkingLocation:"📍 جاري التحقق…", locationVerified:"✅ تم التحقق من الموقع",
    verifyLocation:"📍 التحقق من موقعي",
    submitAttendance:"✅ تسجيل الحضور", saving:"جاري الحفظ…",
    signOutBtn:"🚪 تسجيل الخروج", recordTimeOut:"🚪 تسجيل وقت الخروج",
    signedOutAt:"تم تسجيل الخروج في", timeRecorded:"تم تسجيل وقت الخروج",
    myStatus:"📋 حالتي", payment:"الدفع", laptop:"الحاسوب",
    paid:"✅ مدفوع", unpaid:"❌ غير مدفوع", received:"✅ تم الاستلام",
    notYet:"❌ لم يتم بعد", returned:"🔄 تم الإرجاع",
    returnedLaptop:"🔄 أعدت الحاسوب إلى الموارد البشرية",
    receivedLaptop:"✅ استلمت حاسوبي",
    ojtEndDateLabel:"🎓 تاريخ انتهاء التدريب", daysRemaining:"أيام متبقية", ojtEnded:"انتهت فترة التدريب",
    excuseRequired:"⚠️ مطلوب عذر", excuseType:"نوع العذر",
    selectReason:"اختر السبب…", traffic:"🚗 ازدحام مروري", limitReached:"(تم الوصول للحد)",
    medical:"🏥 طوارئ طبية", family:"👨‍👩‍👧 طوارئ عائلية", other:"📋 أخرى",
    description:"الوصف", proofPhoto:"صورة إثبات", required:"(مطلوب)", optional:"(اختياري)",
    uploadProof:"رفع الإثبات", changePhoto:"تغيير الصورة",
    currentWeek:"الأسبوع الحالي", sundayThursday:"الأحد — الخميس",
    submitted:"تم التقديم", pending:"قيد الانتظار", reportSubmitted:"✅ تم تقديم التقرير",
    edit:"✏️ تعديل", weekSummary:"ملخص الأسبوع", talentNotes:"🌟 ملاحظات المواهب",
    weeklyKpi:"نقاط KPI الأسبوعية", outOf:"من 100", aiAnalysis:"🤖 تحليل الذكاء الاصطناعي",
    thisWeekTasks:"📅 مهام هذا الأسبوع",
    writeEachDay:"اكتب ما قمت به كل يوم. سيقوم الذكاء الاصطناعي بتحليل تقريرك وتقييمه.",
    weeklyTasksLabel:"المهام الأسبوعية", proofPhotoOpt:"📸 صورة إثبات (اختياري)",
    submitWeekly:"📅 تقديم التقرير الأسبوعي", aiAnalyzing:"🤖 جاري التحليل…",
    sickLeavePolicy:"ℹ️ سياسة الإجازة المرضية",
    first2days:"أول يومين (48 ساعة) — بدون خصم ✅",
    after48:"بعد 48 ساعة — خصم 8.33% يومياً ⚠️",
    mySickLeaves:"🏥 إجازاتي المرضية", submitSickLeave:"+ تقديم",
    cancel:"✕ إلغاء", startDate:"تاريخ البدء", endDate:"تاريخ الانتهاء",
    reason:"السبب", medicalProof:"📸 إثبات طبي (موصى به)",
    submitSickBtn:"🏥 تقديم إجازة مرضية", submitting:"جاري التقديم…",
    noSickLeaves:"لا توجد إجازات مرضية مسجلة.", total:"الإجمالي",
    freeDays:"أيام مجانية", penaltyDays:"أيام الخصم", nopenalty:"✅ بدون خصم",
    totalGoals:"الإجمالي", inProgress:"قيد التنفيذ", completed:"مكتمل", overdue:"متأخر",
    addGoal:"+ إضافة هدف", newGoal:"🎯 هدف جديد", kraCategory:"فئة KRA",
    goalTitle:"عنوان الهدف", target:"الهدف", unit:"الوحدة", dueDate:"تاريخ الاستحقاق",
    setGoal:"🎯 تحديد الهدف", noGoals:"لا توجد أهداف بعد. اضغط + إضافة هدف!",
    update:"تحديث", complete:"✅ اكتمل",
    allTrainees:"👥 جميع المتدربين", searchPlaceholder:"🔍 البحث بالاسم أو القسم أو الهوية...",
    trainees:"المتدربون", live:"مباشر", analytics:"التحليلات", okr:"OKR", access:"السجل",
    exportExcel:"📊 تصدير Excel", attendanceSheets:"📋 كشوف الحضور",
    back:"→ رجوع", reactivate:"✅ إعادة تفعيل", transfer:"🔄 نقل",
    inactive:"⏸ غير نشط", drop:"🔴 إنهاء",
    timeline:"السجل الزمني", reports:"التقارير", penalties:"الخصومات",
    paymentTab:"الدفع", laptopTab:"الحاسوب",
    activityTimeline:"السجل الزمني للنشاط", noActivity:"لا يوجد نشاط بعد.",
    editProfile:"تعديل الملف الشخصي", saveChanges:"حفظ التغييرات",
    joiningDate:"تاريخ الانضمام", quittingDate:"تاريخ المغادرة", quittingReason:"سبب المغادرة",
    paymentStatus:"💰 حالة الدفع", paymentNotes:"ملاحظات الدفع",
    savePayment:"💾 حفظ حالة الدفع",
    laptopStatus:"💻 حالة الحاسوب", markReceived:"✅ تعيين كمستلم",
    markReturned:"🔄 تعيين كمُعاد للموارد البشرية", reset:"❌ إعادة تعيين",
    laptopSerial:"الرقم التسلسلي للحاسوب", dateReceived:"تاريخ الاستلام",
    dateReturned:"تاريخ الإرجاع", saveLaptop:"💾 حفظ حالة الحاسوب",
    goalsTab:"الأهداف", noGoalsSet:"لم يتم تحديد أهداف بعد.",
    liveFeed:"📡 البث المباشر", waitingSignins:"في انتظار تسجيلات الحضور…",
    signedIn:"حاضر", absent:"غائب", onTimeLabel:"في الوقت", late:"متأخر",
    kpiLeaderboard:"🏆 لوحة ترتيب KPI", noData:"لا توجد بيانات بعد.",
    atRisk:"⚠️ في خطر", avgKpi:"متوسط KPI",
    active:"نشط", transferred:"منقول", dropped:"منتهي",
    ocrTracking:"🎯 تتبع OKR", addOkr:"+ إضافة", addOkrTitle:"إضافة OKR",
    objective:"الهدف", keyResult:"النتيجة الرئيسية",
    onTrack:"على المسار", behind:"متأخر",
    accessLog:"🔐 سجل الوصول", loggedInAs:"مسجل الدخول كـ:",
    refresh:"🔄 تحديث", totalActions:"إجمالي الإجراءات", logins:"تسجيلات الدخول",
    profileViews:"مشاهدات الملف", profileEdits:"تعديلات الملف",
    statusChanges:"تغييرات الحالة", paymentUpdates:"تحديثات الدفع",
    laptopUpdates:"تحديثات الحاسوب", exports:"التصديرات",
    accessByPerson:"👤 الوصول حسب الشخص", fullActivityLog:"📋 سجل النشاط الكامل",
    noAccessYet:"لا يوجد وصول مسمى بعد.", selectNameToTrack:"اختر اسمك في نافذة الترحيب لبدء التتبع.",
    noActivityYet:"لا يوجد نشاط بعد.", loginToTrack:"سجل الدخول باسمك لبدء تتبع جميع الإجراءات.",
    managementHub:"مركز الإدارة", whoAccessing:"👤 من يصل اليوم؟",
    selectNameGreeting:"اختر اسمك لتلقي تحية شخصية",
    enterDashboard:"الدخول للوحة التحكم 🚀", letsGo:"هيا بنا! 🚀",
    todayMotivation:"🈲 دافع اليوم", haveNiceDay:"🌟 أتمنى لك يوماً سعيداً!",
    transferTrainee:"نقل متدرب", currentDept:"القسم الحالي:",
    currentMentor:"المشرف الحالي:", newDepartment:"القسم الجديد",
    newMentor:"المشرف الجديد (اختياري)", confirmTransfer:"🔄 تأكيد النقل",
    dropTrainee:"إنهاء تدريب متدرب", confirmDrop:"🔴 تأكيد الإنهاء",
    exportAttendance:"تصدير كشوف الحضور",
    selectMonthYear:"اختر الشهر والسنة لإنشاء كشوف حضور لجميع المتدربين النشطين",
    month:"الشهر", year:"السنة", export:"📋 تصدير",
    attendanceRemark:"الحضور/الملاحظة", attendanceSummary:"ملخص الحضور",
    info:"البيان", days:"الأيام", present:"حاضر",
    sickLeave:"إجازة مرضية", weekend:"عطلة نهاية الأسبوع", holiday:"إجازة رسمية",
    visitSite:"زيارة ميدانية", totalDaysMonth:"إجمالي أيام الشهر",
    traineeSignature:"توقيع المتدرب:", hrSignature:"توقيع مدير الموارد البشرية:",
    approvedBy:"اعتماد المدير:",
    noTrainees:"لا يوجد متدربون يطابقون بحثك.",
    of:"من",
  }
};

const KRA_CATEGORIES = [
  { id:"attendance",  label:"Attendance & Punctuality",  labelAr:"الحضور والانضباط",   icon:"⏰", color:"#CF0A2C" },
  { id:"technical",   label:"Technical Skills",          labelAr:"المهارات التقنية",    icon:"🔧", color:"#FFA500" },
  { id:"reporting",   label:"Reporting & Documentation", labelAr:"التقارير والتوثيق",   icon:"📋", color:"#34d399" },
  { id:"teamwork",    label:"Teamwork & Communication",  labelAr:"العمل الجماعي",       icon:"🤝", color:"#4f8ef7" },
  { id:"learning",    label:"Learning & Development",    labelAr:"التعلم والتطوير",     icon:"📚", color:"#7c5cfc" },
  { id:"initiative",  label:"Initiative & Innovation",   labelAr:"المبادرة والابتكار",  icon:"💡", color:"#FFD700" },
];

const CHINESE_QUOTES = [
  { chinese:"千里之行，始于足下", english:"A journey of a thousand miles begins with a single step.", author:"老子 Lao Tzu" },
  { chinese:"学而不思则罔，思而不学则殆", english:"Learning without thought is labor lost; thought without learning is perilous.", author:"孔子 Confucius" },
  { chinese:"不积跬步，无以至千里", english:"Without accumulating small steps, one cannot walk a thousand miles.", author:"荀子 Xunzi" },
  { chinese:"志不强者智不达", english:"Those with weak ambition cannot achieve great wisdom.", author:"墨子 Mozi" },
  { chinese:"勤能补拙", english:"Diligence can make up for lack of talent.", author:"中国谚语 Chinese Proverb" },
  { chinese:"宝剑锋从磨砺出，梅花香自苦寒来", english:"A sharp sword comes from grinding; plum blossoms come from bitter cold.", author:"中国谚语 Chinese Proverb" },
  { chinese:"失败是成功之母", english:"Failure is the mother of success.", author:"中国谚语 Chinese Proverb" },
  { chinese:"活到老，学到老", english:"Live and learn until old age.", author:"中国谚语 Chinese Proverb" },
  { chinese:"知己知彼，百战不殆", english:"Know yourself and your enemy, and you will never be defeated.", author:"孙子 Sun Tzu" },
  { chinese:"人无远虑，必有近忧", english:"He who does not think ahead will find trouble at his doorstep.", author:"孔子 Confucius" },
];

const THEMES = [
  { name:"Huawei Red",    primary:"#CF0A2C", darkPrimary:"#A00820", dark:"#0D0D0D", surface:"#1E1E1E", surface2:"#2A2A2A", border:"#333333", text:"#F5F5F5", muted:"#888888", white:"#FFFFFF" },
  { name:"Ocean Blue",    primary:"#0066FF", darkPrimary:"#0044BB", dark:"#050E1F", surface:"#0A1628", surface2:"#0F2040", border:"#1A3A6A", text:"#F0F8FF", muted:"#7BA7D4", white:"#FFFFFF" },
  { name:"Royal Purple",  primary:"#7C3AED", darkPrimary:"#5B21B6", dark:"#0D0A1F", surface:"#150E2E", surface2:"#1E1440", border:"#2D1F60", text:"#F5F0FF", muted:"#A78BCA", white:"#FFFFFF" },
  { name:"Emerald Green", primary:"#059669", darkPrimary:"#047857", dark:"#030F0A", surface:"#071A10", surface2:"#0A2518", border:"#0F3D26", text:"#F0FFF8", muted:"#6BB891", white:"#FFFFFF" },
  { name:"Sunset Orange", primary:"#EA580C", darkPrimary:"#C2410C", dark:"#100704", surface:"#1C0D06", surface2:"#28140A", border:"#4A200E", text:"#FFF7F0", muted:"#C48A6A", white:"#FFFFFF" },
  { name:"Rose Pink",     primary:"#E11D78", darkPrimary:"#BE1065", dark:"#0F040A", surface:"#1C0712", surface2:"#28091A", border:"#4A1030", text:"#FFF0F7", muted:"#C47A9E", white:"#FFFFFF" },
  { name:"Sky Teal",      primary:"#0891B2", darkPrimary:"#0E7490", dark:"#030D10", surface:"#071820", surface2:"#0A2230", border:"#0F3A4A", text:"#F0FEFF", muted:"#60A8BB", white:"#FFFFFF" },
  { name:"Golden Sun",    primary:"#D97706", darkPrimary:"#B45309", dark:"#0F0A02", surface:"#1C1404", surface2:"#281E06", border:"#4A380A", text:"#FFFBF0", muted:"#BBA050", white:"#FFFFFF" },
];

const randomTheme = THEMES[Math.floor(Math.random()*THEMES.length)];
const HW = {
  red:randomTheme.primary, darkRed:randomTheme.darkPrimary,
  black:"#1A1A1A", dark:randomTheme.dark,
  surface:randomTheme.surface, surface2:randomTheme.surface2,
  border:randomTheme.border, text:randomTheme.text,
  muted:randomTheme.muted, white:randomTheme.white,
};

function speak(text){
  if(!("speechSynthesis" in window)) return;
  window.speechSynthesis.cancel();
  const msg=new SpeechSynthesisUtterance(text);
  msg.rate=0.9;msg.pitch=1;msg.volume=1;
  const setVoice=()=>{
    const voices=window.speechSynthesis.getVoices();
    const preferred=voices.find(v=>v.lang.startsWith("en")&&v.name.includes("Female"))||voices.find(v=>v.lang.startsWith("en"))||voices[0];
    if(preferred) msg.voice=preferred;
    setTimeout(()=>window.speechSynthesis.speak(msg),400);
  };
  if(window.speechSynthesis.getVoices().length>0) setVoice();
  else window.speechSynthesis.onvoiceschanged=setVoice;
}

async function writeAccessLog(userEmail,managerName,actionType,description,metadata={}){
  try{
    await supabase.from("access_logs").insert({
      manager_name:managerName||"Unknown",manager_email:userEmail||"—",
      action_type:actionType,description:description,
      metadata:{...metadata,day:new Date().toLocaleDateString("en-GB",{weekday:"long"}),date:new Date().toLocaleDateString("en-GB",{day:"numeric",month:"long",year:"numeric"}),time:new Date().toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit",second:"2-digit"})},
    });
  }catch(e){console.error("Access log error:",e);}
}

function getCurrentWeek(){
  const now=new Date();const day=now.getDay();
  const sunday=new Date(now);sunday.setDate(now.getDate()-day);
  const thursday=new Date(sunday);thursday.setDate(sunday.getDate()+4);
  const fmt=(d)=>d.toISOString().split("T")[0];
  return{week_start:fmt(sunday),week_end:fmt(thursday),label:`${sunday.toLocaleDateString("en-GB",{day:"numeric",month:"short"})} – ${thursday.toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"})}`};
}

function HuaweiLogo({size=32}){
  return(<svg width={size} height={size} viewBox="0 0 100 100">{[0,60,120,180,240,300].map(r=>(<ellipse key={r} cx="50" cy="22" rx="8" ry="20" fill={HW.red} transform={`rotate(${r} 50 50)`}/>))}</svg>);
}

function LangToggle({lang,setLang}){
  return(
    <button onClick={()=>setLang(lang==="en"?"ar":"en")}
      style={{background:HW.surface2,border:`1px solid ${HW.border}`,color:HW.text,
        borderRadius:20,padding:"6px 12px",cursor:"pointer",fontSize:13,fontWeight:700,
        display:"flex",alignItems:"center",gap:6}}>
      🌐 {lang==="en"?"العربية":"English"}
    </button>
  );
}

function getDistance(lat1,lng1,lat2,lng2){
  const R=6371000,dLat=(lat2-lat1)*Math.PI/180,dLng=(lng2-lng1)*Math.PI/180;
  const a=Math.sin(dLat/2)**2+Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLng/2)**2;
  return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}

async function analyzeReport(text){
  const res=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json","x-api-key":CLAUDE_KEY,"anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},body:JSON.stringify({model:"claude-haiku-4-5-20251001",max_tokens:1024,messages:[{role:"user",content:`Analyze this trainee weekly report and respond ONLY with JSON:\nReport: "${text}"\n{"kpi_score":<0-100>,"pie_chart":{"Tasks Completed":<pct>,"Planning":<pct>,"Challenges":<pct>,"Learning":<pct>},"talent_notes":"<2-3 sentences>","summary":"<one sentence>"}`}]})});
  const data=await res.json();
  return JSON.parse(data.content[0].text.replace(/```json|```/g,"").trim());
}

function PieChart({data}){
  const colors=[HW.red,"#FF6B6B","#FF9999","#FFB3B3"];
  const entries=Object.entries(data);let cum=0;
  const slices=entries.map(([label,pct],i)=>{
    const val=pct/100,s=cum*2*Math.PI;cum+=val;const e=cum*2*Math.PI;
    const x1=Math.cos(s-Math.PI/2),y1=Math.sin(s-Math.PI/2),x2=Math.cos(e-Math.PI/2),y2=Math.sin(e-Math.PI/2);
    return{label,pct,color:colors[i],d:`M 0 0 L ${x1} ${y1} A 1 1 0 ${val>.5?1:0} 1 ${x2} ${y2} Z`};
  });
  return(<div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:12}}><svg viewBox="-1.1 -1.1 2.2 2.2" width="140" height="140">{slices.map((s,i)=>(<path key={i} d={s.d} fill={s.color} stroke={HW.surface} strokeWidth="0.03"/>))}<circle cx="0" cy="0" r="0.55" fill={HW.surface}/></svg><div style={{display:"flex",flexWrap:"wrap",gap:8,justifyContent:"center"}}>{slices.map((s,i)=>(<div key={i} style={{display:"flex",alignItems:"center",gap:6,fontSize:12}}><div style={{width:8,height:8,borderRadius:"50%",background:s.color,flexShrink:0}}/><span style={{color:HW.muted}}>{s.label}</span><span style={{fontWeight:700,color:s.color}}>{s.pct}%</span></div>))}</div></div>);
}

function OKRBar({okr,onUpdate}){
  const pct=Math.min((okr.current/okr.target)*100,100).toFixed(0);
  const color=pct>=80?HW.red:pct>=50?"#FFA500":"#666";
  return(<div style={{background:HW.surface2,borderRadius:12,padding:14,marginBottom:12}}><div style={{marginBottom:6}}><div style={{fontSize:11,color:HW.muted,marginBottom:2}}>{okr.department}</div><div style={{fontWeight:600,fontSize:14,marginBottom:2}}>{okr.objective}</div><div style={{fontSize:12,color:HW.muted}}>{okr.key_result}</div></div><div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:6}}><span style={{color:HW.muted}}>Progress</span><span style={{fontWeight:700,color}}>{pct}% ({okr.current}/{okr.target} {okr.unit})</span></div><div style={{height:8,background:HW.border,borderRadius:10,overflow:"hidden",marginBottom:8}}><div style={{height:"100%",borderRadius:10,background:color,width:`${pct}%`,transition:"width .6s ease"}}/></div>{onUpdate&&(<div style={{display:"flex",gap:8,alignItems:"center"}}><input type="number" defaultValue={okr.current} style={{background:HW.surface,border:`1px solid ${HW.border}`,color:HW.text,borderRadius:6,padding:"6px 8px",flex:1,fontSize:14}} onBlur={e=>onUpdate(okr.id,parseFloat(e.target.value))}/><span style={{fontSize:12,color:HW.muted}}>Update</span></div>)}</div>);
}

function GoalCard({goal,onUpdate,onDelete,isTrainee,lang}){
  const t=T[lang];
  const kra=KRA_CATEGORIES.find(k=>k.id===goal.kra)||KRA_CATEGORIES[0];
  const kraLabel=lang==="ar"?kra.labelAr:kra.label;
  const pct=Math.min((goal.current_value/goal.target_value)*100,100).toFixed(0);
  const isOverdue=goal.due_date&&new Date(goal.due_date)<new Date()&&goal.status!=="completed";
  const status=isOverdue&&goal.status!=="completed"?"overdue":goal.status;
  const sc={not_started:{bg:"rgba(136,136,136,.15)",color:"#888"},in_progress:{bg:"rgba(79,142,247,.15)",color:"#4f8ef7"},completed:{bg:"rgba(52,211,153,.15)",color:"#34d399"},overdue:{bg:"rgba(248,113,113,.15)",color:"#f87171"}};
  return(<div style={{background:HW.surface2,borderRadius:14,padding:16,marginBottom:12,borderLeft:`4px solid ${kra.color}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}><div style={{flex:1,marginRight:8}}><div style={{display:"flex",alignItems:"center",gap:6,marginBottom:4}}><span style={{fontSize:16}}>{kra.icon}</span><span style={{fontSize:10,color:kra.color,fontWeight:700,textTransform:"uppercase"}}>{kraLabel}</span></div><div style={{fontWeight:700,fontSize:14,marginBottom:2}}>{goal.goal_title}</div>{goal.description&&<div style={{fontSize:12,color:HW.muted}}>{goal.description}</div>}</div><span style={{padding:"3px 8px",borderRadius:20,fontSize:10,fontWeight:700,whiteSpace:"nowrap",background:sc[status]?.bg,color:sc[status]?.color}}>{status==="not_started"?"⬜ "+t.totalGoals:status==="in_progress"?"🔵 "+t.inProgress:status==="completed"?"✅ "+t.completed:"🔴 "+t.overdue}</span></div><div style={{marginBottom:10}}><div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:4}}><span style={{color:HW.muted}}>Progress</span><span style={{fontWeight:700,color:kra.color}}>{goal.current_value}/{goal.target_value} {goal.unit} ({pct}%)</span></div><div style={{height:8,background:HW.border,borderRadius:10,overflow:"hidden"}}><div style={{height:"100%",borderRadius:10,background:kra.color,width:`${pct}%`,transition:"width .6s ease"}}/></div></div>{isTrainee&&goal.status!=="completed"&&(<div style={{display:"flex",gap:8,flexWrap:"wrap"}}><input type="number" placeholder="New value" id={`prog_${goal.id}`} style={{background:HW.surface,border:`1px solid ${HW.border}`,color:HW.text,borderRadius:6,padding:"8px 10px",flex:1,fontSize:14}}/><button style={{padding:"8px 14px",borderRadius:6,border:"none",background:kra.color,color:HW.white,fontWeight:700,fontSize:13,cursor:"pointer"}} onClick={()=>{const input=document.getElementById(`prog_${goal.id}`);if(input&&input.value)onUpdate(goal.id,parseFloat(input.value));}}>{t.update}</button>{parseFloat(pct)>=100&&(<button style={{padding:"8px 14px",borderRadius:6,border:"none",background:"rgba(52,211,153,.15)",color:"#34d399",fontWeight:700,fontSize:13,cursor:"pointer"}} onClick={()=>onUpdate(goal.id,goal.target_value,"completed")}>✅ {t.complete}</button>)}{onDelete&&(<button style={{padding:"8px 14px",borderRadius:6,border:"none",background:"rgba(248,113,113,.1)",color:"#f87171",fontWeight:700,fontSize:13,cursor:"pointer"}} onClick={()=>onDelete(goal.id)}>🗑</button>)}</div>)}</div>);
}

function ReminderPopup({onDismiss,lang}){
  const t=T[lang];
  return(<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.85)",display:"flex",alignItems:"flex-end",justifyContent:"center",zIndex:999,backdropFilter:"blur(4px)"}}><div style={{background:HW.surface,border:`2px solid ${HW.red}`,borderRadius:"24px 24px 0 0",padding:32,width:"100%",textAlign:"center",direction:lang==="ar"?"rtl":"ltr"}}><div style={{fontSize:48,marginBottom:12}}>⏰</div><h3 style={{color:HW.red,fontSize:20,marginBottom:8}}>{lang==="ar"?"تذكير بالتقرير الأسبوعي":"Weekly Report Reminder"}</h3><p style={{color:HW.muted,fontSize:14,lineHeight:1.6,marginBottom:20}}>{lang==="ar"?"لا تنس تقديم تقريرك الأسبوعي قبل الخميس!":"Don't forget to submit your weekly report before Thursday!"}</p><button onClick={onDismiss} style={{background:HW.red,color:HW.white,border:"none",borderRadius:12,padding:"14px 32px",fontWeight:700,fontSize:16,cursor:"pointer",width:"100%"}}>{lang==="ar"?"حسناً!":"Got it!"}</button></div></div>);
}

function TransferPopup({trainee,onConfirm,onCancel,lang}){
  const t=T[lang];
  const [newDept,setNewDept]=useState(trainee?.department||"");
  const [newMentor,setNewMentor]=useState(trainee?.assigned_mentor||"");
  const inp={background:HW.surface2,border:`1px solid ${HW.border}`,color:HW.text,borderRadius:10,padding:"12px 14px",width:"100%",fontFamily:"inherit",fontSize:16,boxSizing:"border-box"};
  const lbl={fontSize:11,fontWeight:700,color:HW.muted,textTransform:"uppercase",letterSpacing:".06em",display:"block",marginBottom:6};
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.88)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000,backdropFilter:"blur(6px)",padding:16}}>
      <div style={{background:HW.surface,border:`2px solid #FFA500`,borderRadius:24,padding:28,width:"100%",maxWidth:420,direction:lang==="ar"?"rtl":"ltr"}}>
        <div style={{textAlign:"center",marginBottom:20}}><div style={{fontSize:36,marginBottom:8}}>🔄</div><h3 style={{fontSize:20,fontWeight:800,color:"#FFA500",margin:"0 0 6px"}}>{t.transferTrainee}</h3><p style={{fontSize:13,color:HW.muted,margin:0}}><b style={{color:HW.text}}>{trainee?.full_name}</b><br/>{t.currentDept} <b style={{color:"#FFA500"}}>{trainee?.department||"—"}</b></p></div>
        <label style={lbl}>{t.newDepartment} *</label>
        <input style={{...inp,marginBottom:14,border:`1px solid #FFA50060`}} placeholder={lang==="ar"?"مثال: البنية التحتية لتقنية المعلومات":"e.g. IT Infrastructure"} value={newDept} onChange={e=>setNewDept(e.target.value)} autoFocus/>
        <label style={lbl}>{t.newMentor}</label>
        <input style={{...inp,marginBottom:20}} placeholder={lang==="ar"?"مثال: د. أحمد":"e.g. Dr. Ahmed"} value={newMentor} onChange={e=>setNewMentor(e.target.value)}/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <button style={{padding:14,borderRadius:12,border:`1px solid ${HW.border}`,background:HW.surface2,color:HW.muted,fontWeight:700,fontSize:14,cursor:"pointer"}} onClick={onCancel}>{t.cancel}</button>
          <button style={{padding:14,borderRadius:12,border:"none",background:"#FFA500",color:"#000",fontWeight:800,fontSize:14,cursor:"pointer",opacity:!newDept.trim()?0.4:1}} disabled={!newDept.trim()} onClick={()=>onConfirm(newDept.trim(),newMentor.trim())}>{t.confirmTransfer}</button>
        </div>
      </div>
    </div>
  );
}

function DropPopup({trainee,onConfirm,onCancel,lang}){
  const t=T[lang];
  const [reason,setReason]=useState("");
  const [quitDate,setQuitDate]=useState(new Date().toISOString().split("T")[0]);
  const inp={background:HW.surface2,border:`1px solid ${HW.border}`,color:HW.text,borderRadius:10,padding:"12px 14px",width:"100%",fontFamily:"inherit",fontSize:16,boxSizing:"border-box"};
  const lbl={fontSize:11,fontWeight:700,color:HW.muted,textTransform:"uppercase",letterSpacing:".06em",display:"block",marginBottom:6};
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.88)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000,backdropFilter:"blur(6px)",padding:16}}>
      <div style={{background:HW.surface,border:`2px solid ${HW.red}`,borderRadius:24,padding:28,width:"100%",maxWidth:420,direction:lang==="ar"?"rtl":"ltr"}}>
        <div style={{textAlign:"center",marginBottom:20}}><div style={{fontSize:36,marginBottom:8}}>🔴</div><h3 style={{fontSize:20,fontWeight:800,color:HW.red,margin:"0 0 6px"}}>{t.dropTrainee}</h3><p style={{fontSize:13,color:HW.muted,margin:0}}><b style={{color:HW.text}}>{trainee?.full_name}</b></p></div>
        <label style={lbl}>{t.quittingDate}</label>
        <input style={{...inp,marginBottom:14}} type="date" value={quitDate} onChange={e=>setQuitDate(e.target.value)}/>
        <label style={lbl}>{t.quittingReason} *</label>
        <textarea style={{...inp,height:90,resize:"vertical",marginBottom:20}} placeholder={lang==="ar"?"مثال: أسباب شخصية، وجد وظيفة، مشاكل صحية…":"e.g. Personal reasons, found a job, health issues…"} value={reason} onChange={e=>setReason(e.target.value)}/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <button style={{padding:14,borderRadius:12,border:`1px solid ${HW.border}`,background:HW.surface2,color:HW.muted,fontWeight:700,fontSize:14,cursor:"pointer"}} onClick={onCancel}>{t.cancel}</button>
          <button style={{padding:14,borderRadius:12,border:"none",background:HW.red,color:HW.white,fontWeight:800,fontSize:14,cursor:"pointer",opacity:!reason.trim()?0.4:1}} disabled={!reason.trim()} onClick={()=>onConfirm(quitDate,reason.trim())}>{t.confirmDrop}</button>
        </div>
      </div>
    </div>
  );
}

function ManagerGreetingPopup({onDismiss,onLogin,lang}){
  const t=T[lang];
  const [selectedName,setSelectedName]=useState("");
  const [greeted,setGreeted]=useState(false);
  const now=new Date();const hour=now.getHours();
  const timeGreeting=hour<12?"Good morning":hour<17?"Good afternoon":"Good evening";
  function doGreet(name){
    setSelectedName(name);setGreeted(true);
    speak(`${timeGreeting}, ${name}! Welcome to Huawei TechTrack, Management Hub. Your insights drive our success. Let's review team performance, track attendance data, and empower the next generation of talent today.`);
    if(onLogin) onLogin(name);
  }
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.92)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,backdropFilter:"blur(8px)",padding:16}}>
      <div style={{background:HW.surface,border:`2px solid ${HW.red}`,borderRadius:24,padding:32,width:"100%",maxWidth:440,textAlign:"center",maxHeight:"90vh",overflowY:"auto",direction:lang==="ar"?"rtl":"ltr"}}>
        <style>{`@keyframes float{0%,100%{transform:translateY(0)}50%{transform:translateY(-6px)}}`}</style>
        <div style={{animation:"float 3s ease-in-out infinite",marginBottom:12}}><HuaweiLogo size={56}/></div>
        <h2 style={{fontSize:22,fontWeight:800,color:HW.red,margin:"0 0 6px"}}>{t.managementHub}</h2>
        <p style={{fontSize:13,color:HW.muted,marginBottom:20}}>{t.appName}</p>
        {!greeted?(
          <>
            <div style={{height:1,background:HW.border,marginBottom:20}}/>
            <p style={{fontSize:15,color:HW.text,fontWeight:600,marginBottom:16}}>{t.whoAccessing}</p>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:20}}>
              {MANAGER_NAMES.map(name=>(
                <button key={name} style={{padding:"14px 10px",borderRadius:12,border:`2px solid ${HW.border}`,background:HW.surface2,color:HW.text,fontWeight:700,fontSize:15,cursor:"pointer"}}
                  onMouseOver={e=>{e.currentTarget.style.background=HW.red;e.currentTarget.style.color=HW.white;e.currentTarget.style.border=`2px solid ${HW.red}`;}}
                  onMouseOut={e=>{e.currentTarget.style.background=HW.surface2;e.currentTarget.style.color=HW.text;e.currentTarget.style.border=`2px solid ${HW.border}`;}}
                  onClick={()=>doGreet(name)}>{name}</button>
              ))}
            </div>
            <p style={{fontSize:12,color:HW.muted}}>{t.selectNameGreeting}</p>
          </>
        ):(
          <>
            <div style={{height:1,background:HW.border,marginBottom:20}}/>
            <div style={{fontSize:32,marginBottom:8}}>👋</div>
            <h3 style={{fontSize:22,fontWeight:800,color:HW.text,margin:"0 0 8px"}}>{timeGreeting}, {selectedName}!</h3>
            <div style={{background:HW.surface2,borderRadius:14,padding:16,marginBottom:20,border:`1px solid ${HW.border}`}}>
              <p style={{fontSize:13,color:HW.muted,lineHeight:1.7,margin:0}}>{lang==="ar"?`أهلاً بك في هواوي تيك تراك، مركز الإدارة.\nرؤيتك تقود نجاحنا.\nلنستعرض أداء الفريق، ونتابع بيانات الحضور،\nونمكّن الجيل القادم من المواهب اليوم.`:"Welcome to Huawei TechTrack, Management Hub.\nYour insights drive our success.\nLet's review team performance, track attendance data,\nand empower the next generation of talent today."}</p>
            </div>
            <div style={{fontSize:13,color:HW.muted,marginBottom:20}}>{now.toLocaleDateString(lang==="ar"?"ar-SA":"en-GB",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</div>
            <button onClick={onDismiss} style={{background:HW.red,color:HW.white,border:"none",borderRadius:12,padding:"16px 40px",fontWeight:800,fontSize:16,cursor:"pointer",width:"100%"}}>{t.enterDashboard}</button>
          </>
        )}
      </div>
    </div>
  );
}

function TraineeGreetingPopup({name,onDismiss,lang}){
  const t=T[lang];
  const now=new Date();const hour=now.getHours();
  const greeting=lang==="ar"?(hour<12?"صباح الخير":hour<17?"مساء الخير":"مساء النور"):(hour<12?"Good morning":hour<17?"Good afternoon":"Good evening");
  const greetingAr=hour<12?"صباح الخير":hour<17?"مساء الخير":"مساء النور";
  const quote=CHINESE_QUOTES[Math.floor(Math.random()*CHINESE_QUOTES.length)];
  useEffect(()=>{
    speak(`${hour<12?"Good morning":hour<17?"Good afternoon":"Good evening"}, ${name}! Welcome to Huawei TechTrack. Together, we step into the future of talent. Let's track your progress, celebrate your milestones, and build your success today. Have a productive session!`);
    return ()=>window.speechSynthesis?.cancel();
  },[]);
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.9)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,backdropFilter:"blur(8px)",padding:16}}>
      <div style={{background:HW.surface,border:`2px solid ${HW.red}`,borderRadius:24,padding:32,width:"100%",maxWidth:420,textAlign:"center",maxHeight:"90vh",overflowY:"auto",direction:lang==="ar"?"rtl":"ltr"}}>
        <style>{`@keyframes float{0%,100%{transform:translateY(0)}50%{transform:translateY(-6px)}}`}</style>
        <div style={{animation:"float 3s ease-in-out infinite",marginBottom:12}}><HuaweiLogo size={56}/></div>
        <div style={{fontSize:12,color:HW.muted,marginBottom:4,letterSpacing:".1em",textTransform:"uppercase"}}>{greetingAr}</div>
        <h2 style={{fontSize:28,fontWeight:800,color:HW.text,margin:"0 0 2px"}}>{greeting},</h2>
        <h2 style={{fontSize:32,fontWeight:800,color:HW.red,margin:"0 0 20px"}}>{name}! 👋</h2>
        <div style={{height:1,background:HW.border,marginBottom:20}}/>
        <div style={{background:HW.surface2,borderRadius:14,padding:16,marginBottom:14,border:`1px solid ${HW.border}`}}>
          <div style={{fontSize:10,color:HW.red,fontWeight:700,textTransform:"uppercase",letterSpacing:".1em",marginBottom:10}}>{t.todayMotivation}</div>
          <div style={{fontSize:20,fontWeight:800,color:HW.text,marginBottom:8,lineHeight:1.4,fontFamily:"serif",letterSpacing:3}}>{quote.chinese}</div>
          <div style={{fontSize:13,color:HW.muted,lineHeight:1.6,fontStyle:"italic",marginBottom:6}}>"{quote.english}"</div>
          <div style={{fontSize:11,color:HW.red,fontWeight:600}}>— {quote.author}</div>
        </div>
        <div style={{background:HW.surface2,borderRadius:12,padding:14,marginBottom:16,border:`1px solid ${HW.border}`}}>
          <p style={{fontSize:13,color:HW.muted,lineHeight:1.7,margin:0}}>{lang==="ar"?"أهلاً بك في هواوي تيك تراك. معاً نخطو نحو مستقبل المواهب. لنتابع تقدمك، ونحتفل بإنجازاتك، ونبني نجاحك اليوم. أتمنى لك جلسة منتجة!":"Welcome to Huawei TechTrack. Together, we step into the future of talent. Let's track your progress, celebrate your milestones, and build your success today. Have a productive session!"}</p>
        </div>
        <div style={{fontSize:13,color:HW.muted,marginBottom:16}}>{now.toLocaleDateString(lang==="ar"?"ar-SA":"en-GB",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</div>
        <div style={{fontSize:12,color:HW.red,fontWeight:700,marginBottom:16,background:`${HW.red}15`,borderRadius:8,padding:"6px 12px",display:"inline-block"}}>🎨 {randomTheme.name}</div>
        <button onClick={onDismiss} style={{background:HW.red,color:HW.white,border:"none",borderRadius:12,padding:"16px 40px",fontWeight:800,fontSize:16,cursor:"pointer",width:"100%"}}>{t.letsGo}</button>
      </div>
    </div>
  );
}

export default function App(){
  const [lang,setLang] = useState("en");
  const t = T[lang];
  const isAr = lang==="ar";
  const dir = isAr?"rtl":"ltr";

  const [view,setView]               = useState("login");
  const [traineeTab,setTraineeTab]   = useState("attendance");
  const [mgmtTab,setMgmtTab]         = useState("trainees");
  const [user,setUser]               = useState(null);
  const [traineeId,setTraineeId]     = useState(null);
  const [trainees,setTrainees]       = useState([]);
  const [selected,setSelected]       = useState(null);
  const [reports,setReports]         = useState([]);
  const [logs,setLogs]               = useState([]);
  const [penalties,setPenalties]     = useState([]);
  const [goals,setGoals]             = useState([]);
  const [selectedGoals,setSelectedGoals] = useState([]);
  const [okrs,setOkrs]               = useState([]);
  const [allReports,setAllReports]   = useState([]);
  const [allPenalties,setAllPenalties] = useState([]);
  const [liveSignins,setLiveSignins] = useState([]);
  const [sickLeaves,setSickLeaves]   = useState([]);
  const [selectedSickLeaves,setSelectedSickLeaves] = useState([]);
  const [allSickLeaves,setAllSickLeaves] = useState([]);
  const [myTraineeData,setMyTraineeData] = useState(null);
  const [accessLogs,setAccessLogs]   = useState([]);
  const [loading,setLoading]         = useState(false);
  const [aiLoading,setAiLoading]     = useState(false);
  const [msg,setMsg]                 = useState("");
  const [email,setEmail]             = useState("");
  const [password,setPassword]       = useState("");
  const [confirmPwd,setConfirmPwd]   = useState("");
  const [profileTab,setProfileTab]   = useState("timeline");
  const [aiResult,setAiResult]       = useState(null);
  const [weeklyPhotoFile,setWeeklyPhotoFile] = useState(null);
  const [weeklyPhotoPreview,setWeeklyPhotoPreview] = useState(null);
  const [sickProofFile,setSickProofFile] = useState(null);
  const [sickProofPreview,setSickProofPreview] = useState(null);
  const [geoStatus,setGeoStatus]     = useState("idle");
  const [geoMsg,setGeoMsg]           = useState("");
  const [locationOk,setLocationOk]   = useState(false);
  const [showReminder,setShowReminder] = useState(false);
  const [showGreeting,setShowGreeting] = useState(false);
  const [showMgrGreeting,setShowMgrGreeting] = useState(false);
  const [showTransferPopup,setShowTransferPopup] = useState(false);
  const [showDropPopup,setShowDropPopup] = useState(false);
  const [showAttendanceExport,setShowAttendanceExport] = useState(false);
  const [traineeName,setTraineeName] = useState("");
  const [currentManagerName,setCurrentManagerName] = useState("");
  const [trafficCount,setTrafficCount] = useState(0);
  const [isLate,setIsLate]           = useState(false);
  const [excuseType,setExcuseType]   = useState("");
  const [excuseText,setExcuseText]   = useState("");
  const [excusePhoto,setExcusePhoto] = useState(null);
  const [excusePreview,setExcusePreview] = useState(null);
  const [showAddSick,setShowAddSick] = useState(false);
  const [newSick,setNewSick]         = useState({start_date:new Date().toISOString().split("T")[0],end_date:new Date().toISOString().split("T")[0],reason:""});
  const [newOkr,setNewOkr]           = useState({department:"",objective:"",key_result:"",target:100,current:0,unit:"%",due_date:""});
  const [showAddOkr,setShowAddOkr]   = useState(false);
  const [showAddGoal,setShowAddGoal] = useState(false);
  const [newGoal,setNewGoal]         = useState({kra:"attendance",goal_title:"",description:"",target_value:100,current_value:0,unit:"%",start_date:new Date().toISOString().split("T")[0],due_date:"",status:"not_started"});
  const [goalFilter,setGoalFilter]   = useState("all");
  const [searchQuery,setSearchQuery] = useState("");
  const [currentTime,setCurrentTime] = useState(new Date());
  const [signedOut,setSignedOut]     = useState(false);
  const [signoutTime,setSignoutTime] = useState("");
  const [currentWeek]                = useState(getCurrentWeek());
  const [weeklyText,setWeeklyText]   = useState("");
  const [weeklySubmitted,setWeeklySubmitted] = useState(false);
  const [attendanceExportMonth,setAttendanceExportMonth] = useState(String(new Date().getMonth()+1).padStart(2,"0"));
  const [attendanceExportYear,setAttendanceExportYear]   = useState(String(new Date().getFullYear()));
  const [setupProfile,setSetupProfile] = useState({full_name:"",civil_id:"",phone_number:"",department:"",assigned_mentor:"",gpa:"",date_of_birth:"",gender:"",nationality:"Omani",university:"",ojt_end_date:""});

  const excuseRef=useRef(),weeklyPhotoRef=useRef(),sickProofRef=useRef();
  const [attendance,setAttendance]=useState({report_date:new Date().toISOString().split("T")[0],attended:false});

  const s={
    page:{minHeight:"100vh",background:HW.dark,color:HW.text,fontFamily:isAr?"'Segoe UI',Tahoma,sans-serif":"sans-serif",padding:"24px 32px 100px 32px",maxWidth:"100%",direction:dir},
    card:{background:HW.surface,border:`1px solid ${HW.border}`,borderRadius:16,padding:16,marginBottom:16},
    input:{background:HW.surface2,border:`1px solid ${HW.border}`,color:HW.text,borderRadius:10,padding:"12px 14px",width:"100%",fontFamily:"inherit",fontSize:16,boxSizing:"border-box",direction:dir},
    label:{fontSize:11,fontWeight:700,color:HW.muted,textTransform:"uppercase",letterSpacing:".06em",display:"block",marginBottom:6},
    btn:{padding:"12px 20px",borderRadius:10,border:"none",fontWeight:700,cursor:"pointer",fontSize:14},
  };

  const statusColors={active:{bg:`${HW.red}20`,color:HW.red},inactive:{bg:"rgba(136,136,136,.15)",color:HW.muted},transferred:{bg:"rgba(255,165,0,.15)",color:"#FFA500"},dropped:{bg:"rgba(136,136,136,.2)",color:"#666"}};
  const eventIcons={joined:"🟢",dropped:"🔴",transferred:"🔄",mentor_changed:"👤",dept_changed:"🏢",reactivated:"✅",note:"📝",inactive:"⏸"};
  function kpiColor(score){return score>=80?HW.red:score>=60?"#FFA500":"#666";}

  useEffect(()=>{
    const interval=setInterval(()=>{const now=new Date();setCurrentTime(now);const tt=`${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;setIsLate(tt>MAX_SIGNIN);},1000);
    return ()=>clearInterval(interval);
  },[]);

  useEffect(()=>{
    const check=()=>{const now=new Date();if(now.getHours()===16&&now.getMinutes()===30){setShowReminder(true);if(Notification.permission==="granted")new Notification("Huawei TechTrack",{body:"Reminder: Submit your weekly report!"});}};
    if(Notification.permission==="default")Notification.requestPermission();
    const interval=setInterval(check,60000);return ()=>clearInterval(interval);
  },[]);

  useEffect(()=>{
    supabase.auth.getSession().then(({data:{session}})=>{if(session)handleSession(session.user);});
    const{data:listener}=supabase.auth.onAuthStateChange((_e,session)=>{if(session)handleSession(session.user);else{setUser(null);setView("login");}});
    return ()=>listener.subscription.unsubscribe();
  },[]);

  useEffect(()=>{
    if(view!=="mgmt")return;
    fetchTodaySignins();
    const channel=supabase.channel("live-signins").on("postgres_changes",{event:"INSERT",schema:"public",table:"daily_reports"},async(payload)=>{
      const r=payload.new;const{data:tt}=await supabase.from("trainees").select("full_name,department").eq("id",r.trainee_id).single();
      setLiveSignins(prev=>[{id:r.id,full_name:tt?.full_name||"Unknown",department:tt?.department||"—",signin_time:r.signin_time,attended:r.attended,report_date:r.report_date,penalty_applied:r.penalty_applied,timestamp:new Date()},...prev].slice(0,50));
    }).subscribe();
    return ()=>supabase.removeChannel(channel);
  },[view]);

  async function fetchTodaySignins(){
    const today=new Date().toISOString().split("T")[0];
    const{data}=await supabase.from("daily_reports").select("*,trainees(full_name,department)").eq("report_date",today).order("created_at",{ascending:false});
    if(data)setLiveSignins(data.map(r=>({id:r.id,full_name:r.trainees?.full_name||"Unknown",department:r.trainees?.department||"—",signin_time:r.signin_time,attended:r.attended,report_date:r.report_date,penalty_applied:r.penalty_applied,timestamp:new Date(r.created_at)})));
  }

  async function handleSession(authUser){
    setUser(authUser);
    const{data}=await supabase.from("profiles").select("role,trainee_id,profile_completed").eq("id",authUser.id).single();
    if(data){
      if(data.role==="management"){
        setView("mgmt");fetchTrainees();fetchOkrs();fetchAllData();fetchAccessLogs();
        setTimeout(()=>setShowMgrGreeting(true),500);
      } else {
        if(!data.profile_completed)setView("setup");
        else{
          setView("trainee");
          if(data.trainee_id){
            setTraineeId(data.trainee_id);
            const{data:tt}=await supabase.from("trainees").select("*").eq("id",data.trainee_id).single();
            if(tt){setMyTraineeData(tt);setTraineeName(tt.full_name.split(" ")[0]);setTimeout(()=>setShowGreeting(true),500);}
            fetchGoals(data.trainee_id);fetchWeeklyReport(data.trainee_id);fetchSickLeaves(data.trainee_id);
            const today=new Date().toISOString().split("T")[0];
            const{data:todayReport}=await supabase.from("daily_reports").select("signout_time").eq("trainee_id",data.trainee_id).eq("report_date",today).single();
            if(todayReport?.signout_time){setSignedOut(true);setSignoutTime(todayReport.signout_time);}
            const start=new Date();start.setDate(1);
            const{data:tc}=await supabase.from("traffic_excuses").select("id").eq("trainee_id",data.trainee_id).gte("created_at",start.toISOString());
            if(tc)setTrafficCount(tc.length);
          }
        }
      }
    }
  }

  async function refreshMyTraineeData(){
    if(!traineeId)return;
    const{data}=await supabase.from("trainees").select("*").eq("id",traineeId).single();
    if(data)setMyTraineeData(data);
  }

  async function fetchWeeklyReport(tid){
    const week=getCurrentWeek();
    const{data}=await supabase.from("daily_reports").select("*").eq("trainee_id",tid).eq("week_start",week.week_start).single();
    if(data){setWeeklySubmitted(true);setWeeklyText(data.weekly_tasks||"");if(data.kpi_score)setAiResult({kpi_score:data.kpi_score,pie_chart:data.pie_chart_json?JSON.parse(data.pie_chart_json):null,talent_notes:data.talent_notes,summary:data.report_text});}
  }

  async function fetchSickLeaves(tid){const{data}=await supabase.from("sick_leaves").select("*").eq("trainee_id",tid).order("start_date",{ascending:false});if(data)setSickLeaves(data);}
  async function fetchSelectedSickLeaves(tid){const{data}=await supabase.from("sick_leaves").select("*").eq("trainee_id",tid).order("start_date",{ascending:false});if(data)setSelectedSickLeaves(data);}
  async function fetchAccessLogs(){const{data}=await supabase.from("access_logs").select("*").order("created_at",{ascending:false}).limit(300);if(data)setAccessLogs(data);}

  async function logAccess(actionType,description,metadata={}){
    const name=currentManagerName||"Unknown";
    await writeAccessLog(user?.email,name,actionType,description,metadata);
    setTimeout(fetchAccessLogs,500);
  }

  async function traineeLaptopAction(action){
    if(!traineeId)return;setLoading(true);
    const updates={};
    if(action==="received"){updates.laptop_received=true;updates.laptop_received_date=new Date().toISOString().split("T")[0];updates.laptop_returned=false;}
    if(action==="returned"){updates.laptop_returned=true;updates.laptop_returned_date=new Date().toISOString().split("T")[0];}
    await supabase.from("trainees").update(updates).eq("id",traineeId);
    await refreshMyTraineeData();setMsg("✅ "+t.saveLaptop);setLoading(false);
  }

  async function submitSickLeave(){
    if(!newSick.reason){setMsg(isAr?"الرجاء إدخال السبب":"Please enter a reason.");return;}
    setLoading(true);setMsg("");
    const tid=traineeId;const ts=Date.now();
    const start=new Date(newSick.start_date);const end=new Date(newSick.end_date);
    const totalDays=Math.ceil((end-start)/(1000*60*60*24))+1;
    const penaltyDays=Math.max(0,totalDays-2);const penaltyApplied=penaltyDays>0;
    let proofUrl=null;
    if(sickProofFile)proofUrl=await uploadFile("report-photos",`${tid}/sick_${ts}.jpg`,sickProofFile);
    const{error}=await supabase.from("sick_leaves").insert({trainee_id:tid,start_date:newSick.start_date,end_date:newSick.end_date,reason:newSick.reason,proof_url:proofUrl,total_days:totalDays,penalty_days:penaltyDays,penalty_applied:penaltyApplied});
    if(penaltyApplied){for(let i=2;i<totalDays;i++){const d=new Date(start);d.setDate(start.getDate()+i);await supabase.from("penalties").insert({trainee_id:tid,report_date:d.toISOString().split("T")[0],reason:`Sick leave exceeded 48 hours (day ${i+1})`,amount:PENALTY_PCT});}}
    if(error)setMsg("Error: "+error.message);
    else{setMsg(penaltyApplied?`✅ ${isAr?"تم التقديم":"Submitted"} — ⚠️ ${penaltyDays} ${isAr?"أيام خصم":"penalty day(s)"}`:`✅ ${isAr?"تم التقديم — بدون خصم":"Submitted — No penalty"}`);setShowAddSick(false);setNewSick({start_date:new Date().toISOString().split("T")[0],end_date:new Date().toISOString().split("T")[0],reason:""});setSickProofFile(null);setSickProofPreview(null);fetchSickLeaves(tid);}
    setLoading(false);
  }

  async function saveSetupProfile(){
    if(!setupProfile.full_name||!setupProfile.civil_id){setMsg(isAr?"يرجى ملء الاسم الكامل ورقم الهوية":"Please fill Full Name and Civil ID.");return;}
    setLoading(true);setMsg("");
    const{data:tData,error:tErr}=await supabase.from("trainees").upsert({full_name:setupProfile.full_name,civil_id:setupProfile.civil_id,phone_number:setupProfile.phone_number,department:setupProfile.department,assigned_mentor:setupProfile.assigned_mentor,gpa:setupProfile.gpa?parseFloat(setupProfile.gpa):null,date_of_birth:setupProfile.date_of_birth||null,gender:setupProfile.gender||null,nationality:setupProfile.nationality||null,university:setupProfile.university||null,ojt_end_date:setupProfile.ojt_end_date||null,joining_date:new Date().toISOString().split("T")[0]},{onConflict:"civil_id"}).select().single();
    if(tErr){setMsg("Error: "+tErr.message);setLoading(false);return;}
    await supabase.from("profiles").update({trainee_id:tData.id,profile_completed:true}).eq("id",user.id);
    setTraineeId(tData.id);setMyTraineeData(tData);setTraineeName(setupProfile.full_name.split(" ")[0]);
    fetchGoals(tData.id);setLoading(false);setView("trainee");setTimeout(()=>setShowGreeting(true),500);
  }

  async function login(){setLoading(true);setMsg("");const{error}=await supabase.auth.signInWithPassword({email,password});if(error)setMsg(error.message);setLoading(false);}

  async function signup(){
    if(!email||!password){setMsg(isAr?"يرجى إدخال البريد الإلكتروني وكلمة المرور":"Please enter email and password.");return;}
    if(password!==confirmPwd){setMsg(isAr?"كلمتا المرور غير متطابقتين":"Passwords do not match.");return;}
    if(password.length<6){setMsg(isAr?"كلمة المرور يجب أن تكون 6 أحرف على الأقل":"Password must be at least 6 characters.");return;}
    setLoading(true);setMsg("");
    const{data:authData,error}=await supabase.auth.signUp({email,password});
    if(error){setMsg(error.message);setLoading(false);return;}
    await supabase.from("profiles").insert({id:authData.user.id,email,role:"trainee",profile_completed:false});
    setMsg("✅ "+(isAr?"تم إنشاء الحساب! يمكنك الآن تسجيل الدخول.":"Account created! You can now log in."));
    setLoading(false);setTimeout(()=>{setView("login");setMsg("");setConfirmPwd("");},2000);
  }

  async function logout(){
    if(currentManagerName)await writeAccessLog(user?.email,currentManagerName,"logout",`${currentManagerName} logged out of Management Hub`);
    await supabase.auth.signOut();
    setView("login");setUser(null);setSelected(null);setMsg("");
    setAiResult(null);setWeeklyPhotoFile(null);setWeeklyPhotoPreview(null);
    setLocationOk(false);setGeoStatus("idle");setGeoMsg("");
    setTraineeId(null);setGoals([]);setLiveSignins([]);
    setShowGreeting(false);setShowMgrGreeting(false);setTraineeName("");
    setSignedOut(false);setSignoutTime("");setWeeklySubmitted(false);setWeeklyText("");
    setSickLeaves([]);setMyTraineeData(null);setSearchQuery("");
    setAccessLogs([]);setCurrentManagerName("");setShowTransferPopup(false);setShowDropPopup(false);
    window.speechSynthesis?.cancel();
  }

  function checkLocation(){
    setGeoStatus("checking");setGeoMsg("📍 "+(isAr?"جاري التحقق من موقعك…":"Checking your location…"));
    if(!navigator.geolocation){setGeoStatus("error");setGeoMsg("❌ "+(isAr?"المتصفح لا يدعم تحديد الموقع":"Browser does not support location."));return;}
    navigator.geolocation.getCurrentPosition(
      (pos)=>{const dist=getDistance(pos.coords.latitude,pos.coords.longitude,WORK_LAT,WORK_LNG);if(dist<=WORK_RADIUS){setLocationOk(true);setGeoStatus("ok");setGeoMsg(`✅ ${isAr?"تم التحقق":"Verified"} — ${Math.round(dist)}m`);}else{setLocationOk(false);setGeoStatus("error");setGeoMsg(`❌ ${Math.round(dist)}m — ${isAr?"يجب أن تكون ضمن":"Must be within"} ${WORK_RADIUS}m`);}},
      ()=>{setGeoStatus("error");setGeoMsg("❌ "+(isAr?"تم رفض الوصول للموقع":"Location denied."));},{enableHighAccuracy:true,timeout:10000}
    );
  }

  function handleExcusePhoto(e){const f=e.target.files[0];if(!f)return;setExcusePhoto(f);setExcusePreview(URL.createObjectURL(f));}
  function handleWeeklyPhoto(e){const f=e.target.files[0];if(!f)return;setWeeklyPhotoFile(f);setWeeklyPhotoPreview(URL.createObjectURL(f));}
  function handleSickProof(e){const f=e.target.files[0];if(!f)return;setSickProofFile(f);setSickProofPreview(URL.createObjectURL(f));}

  async function uploadFile(bucket,path,file){
    const{error}=await supabase.storage.from(bucket).upload(path,file);
    if(error)return null;
    return supabase.storage.from(bucket).getPublicUrl(path).data.publicUrl;
  }

  async function submitAttendance(){
    if(!locationOk){setMsg("📍 "+(isAr?"يرجى التحقق من موقعك أولاً":"Please verify your location first."));return;}
    const now=new Date();const timeStr=`${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;
    if(timeStr>MAX_SIGNIN&&!excuseText){setMsg("⚠️ "+(isAr?"بعد 9:00 صباحاً — يرجى تقديم عذر":"Past 9:00 AM — please provide an excuse."));return;}
    if(excuseType==="traffic"){if(trafficCount>=2){setMsg("❌ "+(isAr?"تم الوصول لحد عذر المرور":"Traffic excuse limit reached."));return;}if(!excusePhoto){setMsg("⚠️ "+(isAr?"عذر المرور يتطلب صورة إثبات":"Traffic excuse requires proof photo."));return;}}
    const penaltyApplied=!attendance.attended||(timeStr>MAX_SIGNIN&&!excuseText);
    setLoading(true);setMsg("");const tid=traineeId;const ts=Date.now();
    let excusePhotoUrl=null;
    if(excusePhoto)excusePhotoUrl=await uploadFile("report-photos",`${tid}/excuse_${ts}.jpg`,excusePhoto);
    if(excuseType==="traffic"&&excusePhotoUrl){await supabase.from("traffic_excuses").insert({trainee_id:tid,report_date:attendance.report_date,photo_url:excusePhotoUrl});setTrafficCount(c=>c+1);}
    if(penaltyApplied)await supabase.from("penalties").insert({trainee_id:tid,report_date:attendance.report_date,reason:!attendance.attended?"Absent":"Late sign-in after 9:00 AM",amount:PENALTY_PCT});
    await supabase.from("daily_reports").upsert({trainee_id:tid,report_date:attendance.report_date,attended:attendance.attended,signin_time:timeStr,excuse_type:excuseType||null,excuse_text:excuseText||null,excuse_photo_url:excusePhotoUrl,traffic_excuse:excuseType==="traffic",penalty_applied:penaltyApplied,penalty_amount:penaltyApplied?PENALTY_PCT:0},{onConflict:"trainee_id,report_date"});
    setMsg(penaltyApplied?`✅ ${isAr?"تم تسجيل الحضور":"Attendance recorded"} — ⚠️ ${isAr?"خصم":"Penalty"} ${PENALTY_PCT}%`:`✅ ${isAr?"تم تسجيل الحضور بنجاح!":"Attendance recorded successfully!"}`);
    setLoading(false);
  }

  async function submitSignOut(){
    if(!traineeId){setMsg(isAr?"يرجى تسجيل الحضور أولاً":"Please submit attendance first.");return;}
    setLoading(true);setMsg("");
    const now=new Date();const timeStr=`${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;
    const today=new Date().toISOString().split("T")[0];
    const{error}=await supabase.from("daily_reports").update({signout_time:timeStr}).eq("trainee_id",traineeId).eq("report_date",today);
    if(error)setMsg("Error: "+error.message);
    else{setSignedOut(true);setSignoutTime(timeStr);setMsg(`✅ ${t.signedOutAt} ${timeStr}`);}
    setLoading(false);
  }

  async function submitWeeklyReport(){
    if(!weeklyText){setMsg(isAr?"يرجى كتابة مهامك الأسبوعية":"Please write your weekly tasks.");return;}
    setLoading(true);setAiLoading(true);setMsg("");setAiResult(null);
    const tid=traineeId;if(!tid){setMsg(isAr?"يرجى تسجيل الحضور أولاً":"Please submit attendance first.");setLoading(false);setAiLoading(false);return;}
    const ts=Date.now();const week=getCurrentWeek();
    const photoUrl=weeklyPhotoFile?await uploadFile("report-photos",`${tid}/weekly_${ts}.jpg`,weeklyPhotoFile):null;
    setMsg("🤖 "+(isAr?"جاري تحليل التقرير…":"AI analyzing your report…"));
    let ai=null;try{ai=await analyzeReport(weeklyText);setAiResult(ai);}catch(e){console.error(e);}
    await supabase.from("daily_reports").upsert({trainee_id:tid,report_date:week.week_start,week_start:week.week_start,week_end:week.week_end,weekly_tasks:weeklyText,weekly_photo_url:photoUrl,report_text:ai?.summary||weeklyText.substring(0,100),kpi_score:ai?.kpi_score||null,pie_chart_json:ai?.pie_chart?JSON.stringify(ai.pie_chart):null,talent_notes:ai?.talent_notes||null},{onConflict:"trainee_id,report_date"});
    setWeeklySubmitted(true);setMsg("✅ "+(isAr?"تم تقديم التقرير الأسبوعي!":"Weekly report submitted!"));setLoading(false);setAiLoading(false);
  }

  async function fetchGoals(tid){const{data}=await supabase.from("goals").select("*").eq("trainee_id",tid).order("created_at",{ascending:false});if(data)setGoals(data);}
  async function addGoal(){
    if(!newGoal.goal_title){setMsg(isAr?"يرجى إدخال عنوان الهدف":"Please enter a goal title.");return;}
    setLoading(true);setMsg("");
    const{error}=await supabase.from("goals").insert({...newGoal,trainee_id:traineeId});
    if(error)setMsg(error.message);
    else{setMsg("✅ "+(isAr?"تم إضافة الهدف!":"Goal added!"));setShowAddGoal(false);setNewGoal({kra:"attendance",goal_title:"",description:"",target_value:100,current_value:0,unit:"%",start_date:new Date().toISOString().split("T")[0],due_date:"",status:"not_started"});fetchGoals(traineeId);}
    setLoading(false);
  }
  async function updateGoal(id,newValue,newStatus){
    const updates={current_value:newValue};
    if(newStatus)updates.status=newStatus;
    else if(newValue>=goals.find(g=>g.id===id)?.target_value)updates.status="completed";
    await supabase.from("goals").update(updates).eq("id",id);fetchGoals(traineeId);setMsg("✅ "+(isAr?"تم تحديث الهدف":"Goal updated!"));
  }
  async function deleteGoal(id){await supabase.from("goals").delete().eq("id",id);fetchGoals(traineeId);setMsg("✅ "+(isAr?"تم حذف الهدف":"Goal deleted."));}
  async function fetchSelectedGoals(tid){const{data}=await supabase.from("goals").select("*").eq("trainee_id",tid).order("created_at",{ascending:false});if(data)setSelectedGoals(data);}
  async function fetchTrainees(){const{data}=await supabase.from("trainees").select("*").order("full_name");if(data)setTrainees(data);}
  async function fetchOkrs(){const{data}=await supabase.from("okrs").select("*").order("department");if(data)setOkrs(data);}
  async function fetchAllData(){
    const{data:reps}=await supabase.from("daily_reports").select("*");
    const{data:pens}=await supabase.from("penalties").select("*");
    const{data:sick}=await supabase.from("sick_leaves").select("*");
    if(reps)setAllReports(reps);if(pens)setAllPenalties(pens);if(sick)setAllSickLeaves(sick);
  }
  async function fetchReports(tid){const{data}=await supabase.from("daily_reports").select("*").eq("trainee_id",tid).order("report_date",{ascending:false});if(data)setReports(data);}
  async function fetchLogs(tid){const{data}=await supabase.from("trainee_logs").select("*").eq("trainee_id",tid).order("created_at",{ascending:false});if(data)setLogs(data);}
  async function fetchPenalties(tid){const{data}=await supabase.from("penalties").select("*").eq("trainee_id",tid).order("created_at",{ascending:false});if(data)setPenalties(data);}

  async function openProfile(tt){
    setSelected({...tt});setProfileTab("timeline");setMsg("");
    await fetchReports(tt.id);await fetchLogs(tt.id);await fetchPenalties(tt.id);await fetchSelectedGoals(tt.id);await fetchSelectedSickLeaves(tt.id);
    await writeAccessLog(user?.email,currentManagerName,"profile_view",`Viewed profile of ${tt.full_name}`,{trainee:tt.full_name});
    setTimeout(fetchAccessLogs,500);
  }

  async function updatePaymentStatus(status,notes){
    await supabase.from("trainees").update({payment_status:status,payment_notes:notes||null}).eq("id",selected.id);
    setSelected({...selected,payment_status:status,payment_notes:notes});
    fetchTrainees();setMsg("✅ "+(isAr?"تم تحديث حالة الدفع":"Payment status updated!"));
    await writeAccessLog(user?.email,currentManagerName,"payment_update",`Updated payment for ${selected.full_name} to ${status}`,{trainee:selected.full_name,payment_status:status});
    setTimeout(fetchAccessLogs,500);
  }

  async function updateLaptopStatus(received,serial,date,returned,returnedDate){
    await supabase.from("trainees").update({laptop_received:received||false,laptop_serial:serial||null,laptop_received_date:date||null,laptop_returned:returned||false,laptop_returned_date:returnedDate||null}).eq("id",selected.id);
    setSelected({...selected,laptop_received:received,laptop_serial:serial,laptop_received_date:date,laptop_returned:returned,laptop_returned_date:returnedDate});
    fetchTrainees();setMsg("✅ "+(isAr?"تم تحديث حالة الحاسوب":"Laptop status updated!"));
    await writeAccessLog(user?.email,currentManagerName,"laptop_update",`Updated laptop for ${selected.full_name}`,{trainee:selected.full_name,serial,received,returned});
    setTimeout(fetchAccessLogs,500);
  }

  async function logEvent(tid,type,desc,old="",nw=""){await supabase.from("trainee_logs").insert({trainee_id:tid,event_type:type,description:desc,old_value:old,new_value:nw,logged_by:user?.email||"manager"});}

  async function saveProfile(){
    const changes=[];
    if(selected.department!==selected._original?.department)changes.push(logEvent(selected.id,"dept_changed","Department changed",selected._original?.department,selected.department));
    if(selected.assigned_mentor!==selected._original?.assigned_mentor)changes.push(logEvent(selected.id,"mentor_changed","Mentor changed",selected._original?.assigned_mentor,selected.assigned_mentor));
    const{error}=await supabase.from("trainees").update({department:selected.department,assigned_mentor:selected.assigned_mentor,gpa:selected.gpa,joining_date:selected.joining_date,university:selected.university||null,ojt_end_date:selected.ojt_end_date||null}).eq("id",selected.id);
    if(error){setMsg(error.message);return;}
    await Promise.all(changes);
    setMsg("✅ "+(isAr?"تم الحفظ":"Saved!"));fetchTrainees();fetchLogs(selected.id);
    await writeAccessLog(user?.email,currentManagerName,"profile_edit",`Edited profile of ${selected.full_name}`,{trainee:selected.full_name,department:selected.department});
    setTimeout(fetchAccessLogs,500);
  }

  async function handleTransferConfirm(newDept,newMentor){
    setShowTransferPopup(false);
    const oldDept=selected.department;const oldMentor=selected.assigned_mentor;
    const updateData={status:"transferred",department:newDept,assigned_mentor:newMentor||selected.assigned_mentor};
    const{error}=await supabase.from("trainees").update(updateData).eq("id",selected.id);
    if(error){setMsg(error.message);return;}
    await logEvent(selected.id,"transferred",`Transferred from ${oldDept} to ${newDept}`,oldDept,newDept);
    if(newMentor&&newMentor!==oldMentor)await logEvent(selected.id,"mentor_changed",`Mentor changed to ${newMentor}`,oldMentor,newMentor);
    setSelected({...selected,...updateData});fetchTrainees();fetchLogs(selected.id);
    setMsg(`✅ ${isAr?"تم النقل إلى":"Transferred to"} ${newDept}`);
    await writeAccessLog(user?.email,currentManagerName,"status_change",`Transferred ${selected.full_name} from ${oldDept} to ${newDept}`,{trainee:selected.full_name,old_department:oldDept,new_department:newDept,new_mentor:newMentor||oldMentor,old_status:selected.status,new_status:"transferred"});
    setTimeout(fetchAccessLogs,500);
  }

  async function handleDropConfirm(quitDate,reason){
    setShowDropPopup(false);
    const updateData={status:"dropped",quitting_date:quitDate,quitting_reason:reason};
    const{error}=await supabase.from("trainees").update(updateData).eq("id",selected.id);
    if(error){setMsg(error.message);return;}
    await logEvent(selected.id,"dropped",`Dropped — Reason: ${reason}`,selected.status,"dropped");
    setSelected({...selected,...updateData});fetchTrainees();fetchLogs(selected.id);
    setMsg(`✅ ${isAr?"تم إنهاء التدريب":"Trainee marked as dropped"}`);
    await writeAccessLog(user?.email,currentManagerName,"status_change",`Dropped ${selected.full_name} — ${reason}`,{trainee:selected.full_name,reason,quit_date:quitDate,old_status:selected.status,new_status:"dropped"});
    setTimeout(fetchAccessLogs,500);
  }

  async function changeStatus(newStatus){
    if(newStatus==="transferred"){setShowTransferPopup(true);return;}
    if(newStatus==="dropped"){setShowDropPopup(true);return;}
    const updateData={status:newStatus};
    if(newStatus==="active"){updateData.quitting_date=null;updateData.quitting_reason=null;}
    const{error}=await supabase.from("trainees").update(updateData).eq("id",selected.id);
    if(error){setMsg(error.message);return;}
    await logEvent(selected.id,newStatus==="active"?"reactivated":newStatus,`Status changed to ${newStatus}`,selected.status,newStatus);
    setSelected({...selected,status:newStatus,...updateData});fetchTrainees();fetchLogs(selected.id);
    setMsg(`✅ ${isAr?"تم تحديث الحالة إلى":"Status updated to"} ${newStatus}`);
    await writeAccessLog(user?.email,currentManagerName,"status_change",`Changed ${selected.full_name} status to ${newStatus}`,{trainee:selected.full_name,old_status:selected.status,new_status:newStatus});
    setTimeout(fetchAccessLogs,500);
  }

  async function updateOkr(id,newValue){await supabase.from("okrs").update({current:newValue}).eq("id",id);fetchOkrs();}
  async function addOkr(){
    if(!newOkr.department||!newOkr.objective||!newOkr.key_result){setMsg(isAr?"يرجى ملء جميع حقول OKR":"Please fill all OKR fields.");return;}
    await supabase.from("okrs").insert({...newOkr,created_by:user?.email});
    setShowAddOkr(false);setNewOkr({department:"",objective:"",key_result:"",target:100,current:0,unit:"%",due_date:""});fetchOkrs();setMsg("✅ OKR "+(isAr?"تمت الإضافة":"added!"));
    await writeAccessLog(user?.email,currentManagerName,"okr_add",`Added OKR: ${newOkr.objective}`,{department:newOkr.department});
    setTimeout(fetchAccessLogs,500);
  }

  async function exportAttendanceSheets(){
    setMsg(isAr?"📋 جاري إنشاء كشوف الحضور…":"📋 Generating attendance sheets…");
    setShowAttendanceExport(false);
    await writeAccessLog(user?.email,currentManagerName,"excel_export","Exported attendance sheets");
    const year=parseInt(attendanceExportYear);const month=parseInt(attendanceExportMonth);
    const monthName=new Date(year,month-1,1).toLocaleDateString("en-GB",{month:"long",year:"numeric"});
    const daysInMonth=new Date(year,month,0).getDate();const monthPad=String(month).padStart(2,"0");
    const{data:allT}=await supabase.from("trainees").select("*").in("status",["active","transferred"]).order("full_name");
    const{data:allR}=await supabase.from("daily_reports").select("*").gte("report_date",`${year}-${monthPad}-01`).lte("report_date",`${year}-${monthPad}-${String(daysInMonth).padStart(2,"0")}`);
    const XL=XLSX;const wb=XL.utils.book_new();
    for(const trainee of (allT||[])){
      const days=[];
      for(let d=1;d<=daysInMonth;d++){
        const dateStr=`${year}-${monthPad}-${String(d).padStart(2,"0")}`;
        const dateObj=new Date(dateStr+"T00:00:00");
        const dow=dateObj.getDay();const isWeekend=dow===5||dow===6;
        const report=(allR||[]).find(rep=>rep.trainee_id===trainee.id&&rep.report_date===dateStr&&!rep.week_start);
        let remark="";
        if(isWeekend)remark="WK";
        else if(report){if(report.attended){remark=report.signin_time&&report.signin_time>MAX_SIGNIN?"LATE":"P";}else if(report.excuse_type==="medical"||report.excuse_type==="sick")remark="SL";else remark="A";}
        days.push({date:dateStr,label:dateObj.toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}).replace(/ /g,"-"),isWeekend,report,remark});
      }
      const counts={P:0,LATE:0,A:0,SL:0,WK:0,H:0,V:0};
      days.forEach(d=>{if(counts[d.remark]!==undefined)counts[d.remark]++;});
      const half=Math.ceil(days.length/2);
      const left=days.slice(0,half);const right=days.slice(half);const maxR=Math.max(left.length,right.length);
      const aoa=[];
      aoa.push(["Project Name: Internship Program","","","","Huawei Tech Investment (Oman) LLC","","","","","Attendance Sheet for "+monthName,"",""]);
      aoa.push(["","","","","","","","","","","",""]);
      aoa.push([`Trainee Name: ${trainee.full_name}`,"","","",`HUAWEI ID: ${trainee.civil_id||"—"}`,"","","","",`Position/Designation: ${trainee.department||"—"}`,"",""]);
      aoa.push(["","","","","","","","","","","",""]);
      aoa.push(["Date","Time In","Time Out","Attendance/Remark","Date","Time In","Time Out","Attendance/Remark","Attendance Summary","Info","Days",""]);
      for(let i=0;i<maxR;i++){
        const ld=left[i];const rd=right[i];
        const row=["","","","","","","","","","","",""];
        if(ld){row[0]=ld.label;if(ld.isWeekend){row[1]="Weekend";row[2]="";row[3]="WK";}else{row[1]=ld.report?.signin_time||"";row[2]=ld.report?.signout_time||"";row[3]=ld.remark;}}
        if(rd){row[4]=rd.label;if(rd.isWeekend){row[5]="Weekend";row[6]="";row[7]="WK";}else{row[5]=rd.report?.signin_time||"";row[6]=rd.report?.signout_time||"";row[7]=rd.remark;}}
        if(i===0){row[8]="Present";row[9]="P";row[10]=counts.P;}
        if(i===1){row[8]="Absent";row[9]="A";row[10]=counts.A;}
        if(i===2){row[8]="Late";row[9]="LATE";row[10]=counts.LATE;}
        if(i===3){row[8]="Sick Leave";row[9]="SL";row[10]=counts.SL;}
        if(i===4){row[8]="Weekend";row[9]="WK";row[10]=counts.WK;}
        if(i===5){row[8]="Holiday";row[9]="H";row[10]=counts.H;}
        if(i===6){row[8]="Visit Site";row[9]="V";row[10]=counts.V;}
        if(i===7){row[8]="Total days in month";row[9]="";row[10]=daysInMonth;}
        aoa.push(row);
      }
      aoa.push(["","","","","","","","","","","",""]);
      aoa.push(["Before signing make sure the details are correct. If the student is found to be absent, strict action shall be taken against the student and supervisor both.","","","","","","","","","","",""]);
      aoa.push(["","","","","","","","","","","",""]);
      aoa.push(["Trainee Signature:","","","HR Manager Signature:","","","","Approved By PM/Manager:","","","",""]);
      aoa.push(["","","","","","","","","","","",""]);
      aoa.push(["","","","","","","","","","","",""]);
      const ws=XL.utils.aoa_to_sheet(aoa);
      ws["!cols"]=[{wch:16},{wch:10},{wch:10},{wch:14},{wch:16},{wch:10},{wch:10},{wch:14},{wch:20},{wch:8},{wch:8},{wch:8}];
      const merges=[{s:{r:0,c:0},e:{r:0,c:3}},{s:{r:0,c:4},e:{r:0,c:8}},{s:{r:0,c:9},e:{r:0,c:11}},{s:{r:2,c:0},e:{r:2,c:3}},{s:{r:2,c:4},e:{r:2,c:8}},{s:{r:2,c:9},e:{r:2,c:11}},{s:{r:aoa.length-5,c:0},e:{r:aoa.length-5,c:11}},{s:{r:aoa.length-3,c:0},e:{r:aoa.length-3,c:2}},{s:{r:aoa.length-3,c:3},e:{r:aoa.length-3,c:6}},{s:{r:aoa.length-3,c:7},e:{r:aoa.length-3,c:11}},{s:{r:aoa.length-2,c:0},e:{r:aoa.length-2,c:2}},{s:{r:aoa.length-2,c:3},e:{r:aoa.length-2,c:6}},{s:{r:aoa.length-2,c:7},e:{r:aoa.length-2,c:11}}];
      const startRow=5;
      for(let i=0;i<maxR;i++){const ld=left[i];const rd=right[i];if(ld?.isWeekend)merges.push({s:{r:startRow+i,c:1},e:{r:startRow+i,c:3}});if(rd?.isWeekend)merges.push({s:{r:startRow+i,c:5},e:{r:startRow+i,c:7}});}
      ws["!merges"]=merges;
      const sheetName=trainee.full_name.substring(0,28).replace(/[\\/?*[\]]/g,"");
      XL.utils.book_append_sheet(wb,ws,sheetName);
    }
    XL.writeFile(wb,`Attendance_${monthName.replace(" ","_")}.xlsx`);
    setMsg(`✅ ${isAr?"تم تصدير كشوف الحضور لـ":"Attendance sheets exported for"} ${(allT||[]).length} ${isAr?"متدرب":"trainees"}!`);
    setTimeout(fetchAccessLogs,500);
  }

  async function exportExcel(){
    setMsg(isAr?"📊 جاري تحضير التصدير…":"📊 Preparing export…");
    await writeAccessLog(user?.email,currentManagerName,"excel_export","Exported Excel report");
    const{data:allT}=await supabase.from("trainees").select("*").order("full_name");
    const{data:allR}=await supabase.from("daily_reports").select("*").order("report_date");
    const{data:allP}=await supabase.from("penalties").select("*").order("created_at");
    const{data:allG}=await supabase.from("goals").select("*").order("created_at");
    const{data:allSL}=await supabase.from("sick_leaves").select("*").order("start_date");
    const{data:allAL}=await supabase.from("access_logs").select("*").order("created_at",{ascending:false});
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet((allT||[]).map(t=>({
      "Full Name":t.full_name,"Civil ID":t.civil_id,"Phone":t.phone_number,
      "University":t.university||"—","Department":t.department,"Mentor":t.assigned_mentor,
      "GPA":t.gpa,"Gender":t.gender||"—","Nationality":t.nationality||"—",
      "Status":t.status,"Joining Date":t.joining_date||"—","OJT End Date":t.ojt_end_date||"—",
      "Quitting Date":t.quitting_date||"—","Quitting Reason":t.quitting_reason||"—",
      "Payment Status":t.payment_status||"unpaid","Payment Notes":t.payment_notes||"—",
      "Laptop Received":t.laptop_received?"Yes":"No","Laptop Serial":t.laptop_serial||"—",
      "Laptop Received Date":t.laptop_received_date||"—",
      "Laptop Returned":t.laptop_returned?"Yes":"No","Laptop Returned Date":t.laptop_returned_date||"—",
    }))),"Trainees");
    const transferred=(allT||[]).filter(t=>t.status==="transferred");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(transferred.length>0?transferred.map(t=>{
      const tLog=allAL?.find(l=>l.action_type==="status_change"&&l.metadata?.trainee===t.full_name&&l.metadata?.new_status==="transferred");
      return{"Full Name":t.full_name,"Civil ID":t.civil_id,"University":t.university||"—","Previous Department":tLog?.metadata?.old_department||"—","New Department":t.department||"—","New Mentor":t.assigned_mentor||"—","Transfer Date":tLog?new Date(tLog.created_at).toLocaleDateString("en-GB",{day:"numeric",month:"long",year:"numeric"}):"—","Transferred By":tLog?.manager_name||"—","GPA":t.gpa||"—","Phone":t.phone_number||"—"};
    }):[{"Info":"No transferred trainees yet"}]),"Transferred Trainees");
    const dropped=(allT||[]).filter(t=>t.status==="dropped");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(dropped.length>0?dropped.map(t=>({
      "Full Name":t.full_name,"Civil ID":t.civil_id,"University":t.university||"—","Department":t.department||"—","GPA":t.gpa||"—","Joining Date":t.joining_date||"—","Quitting Date":t.quitting_date||"—","Quitting Reason":t.quitting_reason||"—","OJT End Date":t.ojt_end_date||"—",
    })):[{"Info":"No dropped trainees yet"}]),"Dropped Trainees");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet((allR||[]).filter(r=>!r.week_start).map(r=>{
      const trainee=(allT||[]).find(tt=>tt.id===r.trainee_id);const dateObj=new Date(r.report_date+"T00:00:00");
      return{"Full Name":trainee?.full_name||"—","Department":trainee?.department||"—","Day":dateObj.toLocaleDateString("en-GB",{weekday:"long"}),"Date":dateObj.toLocaleDateString("en-GB",{day:"numeric",month:"long",year:"numeric"}),"Attended":r.attended?"Yes":"No","Time In":r.signin_time||"—","Time Out":r.signout_time||"—","Penalty":r.penalty_applied?`-${r.penalty_amount}%`:"None"};
    })),"Daily Attendance");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet((allR||[]).filter(r=>r.weekly_tasks).map(r=>{
      const trainee=(allT||[]).find(tt=>tt.id===r.trainee_id);
      return{"Full Name":trainee?.full_name||"—","Department":trainee?.department||"—","Week":r.week_start+" to "+r.week_end,"KPI Score":r.kpi_score||"—","Weekly Tasks":r.weekly_tasks||"—","AI Summary":r.report_text||"—"};
    })),"Weekly Reports");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet((allSL||[]).map(sl=>{const trainee=(allT||[]).find(tt=>tt.id===sl.trainee_id);return{"Full Name":trainee?.full_name||"—","Department":trainee?.department||"—","Start":sl.start_date,"End":sl.end_date,"Total Days":sl.total_days,"Reason":sl.reason||"—","Penalty Days":sl.penalty_days||0,"Penalty Applied":sl.penalty_applied?"Yes":"No"};})),"Sick Leaves");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet((allP||[]).map(p=>({"Full Name":(allT||[]).find(tt=>tt.id===p.trainee_id)?.full_name||"—","Date":p.report_date,"Reason":p.reason,"Deduction":`-${p.amount}%`}))),"Penalties");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet((allG||[]).map(g=>({"Full Name":(allT||[]).find(tt=>tt.id===g.trainee_id)?.full_name||"—","KRA":g.kra,"Goal":g.goal_title,"Target":g.target_value,"Current":g.current_value,"Unit":g.unit,"Status":g.status,"Progress":Math.min((g.current_value/g.target_value)*100,100).toFixed(0)+"%"}))),"Goals");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet((allT||[]).map(t=>({"Full Name":t.full_name,"Department":t.department,"Status":t.status,"Payment Status":t.payment_status||"unpaid","Payment Notes":t.payment_notes||"—","Laptop Received":t.laptop_received?"Yes":"No","Laptop Serial":t.laptop_serial||"—","Received Date":t.laptop_received_date||"—","Returned to HR":t.laptop_returned?"Yes":"No","Returned Date":t.laptop_returned_date||"—"}))),"Payment & Laptop");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet((allAL||[]).map(al=>({"Manager Name":al.manager_name||"—","Manager Email":al.manager_email||"—","Action":al.action_type?.replace(/_/g," ")||"—","Description":al.description||"—","Day":al.metadata?.day||"—","Date":al.metadata?.date||"—","Time":al.metadata?.time||"—"}))),"Access Log");
    XLSX.writeFile(wb,`HuaweiTechTrack_${new Date().toISOString().split("T")[0]}.xlsx`);
    setMsg("✅ "+(isAr?"تم التصدير بنجاح!":"Excel exported!"));
    setTimeout(fetchAccessLogs,500);
  }

  async function exportPDF(trainee){
    const doc=new jsPDF();const tt=trainee||selected;
    doc.setFillColor(207,10,44);doc.rect(0,0,210,40,"F");
    doc.setTextColor(255,255,255);doc.setFontSize(18);doc.setFont("helvetica","bold");doc.text("Huawei TechTrack — Performance Report",14,18);
    doc.setFontSize(10);doc.setFont("helvetica","normal");doc.text(`Generated: ${new Date().toLocaleDateString("en-GB")}`,14,30);
    doc.setTextColor(0,0,0);doc.setFontSize(13);doc.setFont("helvetica","bold");doc.text("Trainee Information",14,52);
    autoTable(doc,{startY:56,head:[["Field","Details"]],body:[["Full Name",tt.full_name||"—"],["Civil ID",tt.civil_id||"—"],["University",tt.university||"—"],["Department",tt.department||"—"],["Mentor",tt.assigned_mentor||"—"],["Status",tt.status||"—"],["Joining Date",tt.joining_date||"—"],["OJT End Date",tt.ojt_end_date||"—"],...(tt.quitting_date?[["Quitting Date",tt.quitting_date],["Quitting Reason",tt.quitting_reason||"—"]]:[]),["Payment",tt.payment_status||"unpaid"],["Laptop",tt.laptop_returned?"Returned to HR":tt.laptop_received?`Received — SN: ${tt.laptop_serial||"—"}`:"Not Received"]],headStyles:{fillColor:[207,10,44],textColor:[255,255,255]},alternateRowStyles:{fillColor:[245,245,245]}});
    const penY=doc.lastAutoTable.finalY+10;doc.setFontSize(13);doc.setFont("helvetica","bold");doc.text("Penalties",14,penY);
    const{data:pens}=await supabase.from("penalties").select("*").eq("trainee_id",tt.id);
    autoTable(doc,{startY:penY+4,head:[["Date","Reason","Deduction"]],body:pens?.length>0?pens.map(p=>[p.report_date,p.reason,`-${p.amount}%`]):[["—","No penalties","—"]],headStyles:{fillColor:[207,10,44],textColor:[255,255,255]},alternateRowStyles:{fillColor:[245,245,245]}});
    doc.setTextColor(150,150,150);doc.setFontSize(9);doc.setFont("helvetica","normal");doc.text("Huawei TechTrack • Powered by Arjwan Sabir • Confidential",14,285);
    doc.save(`${tt.full_name}_report.pdf`);
    await writeAccessLog(user?.email,currentManagerName,"pdf_export",`Exported PDF for ${tt.full_name}`,{trainee:tt.full_name});
    setTimeout(fetchAccessLogs,500);
  }

  function getAnalytics(){
    const active=trainees.filter(tt=>tt.status==="active");
    const totalReports=allReports.length;const attended=allReports.filter(r=>r.attended).length;
    const attendanceRate=totalReports>0?((attended/totalReports)*100).toFixed(1):0;
    const kpiReports=allReports.filter(r=>r.kpi_score);
    const avgKpi=kpiReports.length>0?(kpiReports.reduce((a,r)=>a+r.kpi_score,0)/kpiReports.length).toFixed(1):0;
    const traineeKpi=trainees.map(tt=>{
      const tr=allReports.filter(r=>r.trainee_id===tt.id&&r.kpi_score);
      const avg=tr.length>0?tr.reduce((a,r)=>a+r.kpi_score,0)/tr.length:0;
      const tp=allPenalties.filter(p=>p.trainee_id===tt.id).length;
      return{...tt,avgKpi:avg.toFixed(1),penalties:tp};
    }).sort((a,b)=>b.avgKpi-a.avgKpi);
    const atRisk=traineeKpi.filter(tt=>parseFloat(tt.avgKpi)<60||tt.penalties>2);
    return{active,attendanceRate,avgKpi,totalPenalties:allPenalties.length,traineeKpi,atRisk};
  }

  const filteredGoals=goalFilter==="all"?goals:goals.filter(g=>g.kra===goalFilter||g.status===goalFilter);
  const filteredTrainees=trainees.filter(tt=>{
    if(!searchQuery)return true;const q=searchQuery.toLowerCase();
    if(q==="active")return tt.status==="active";if(q==="paid")return tt.payment_status==="paid";
    if(q==="unpaid")return tt.payment_status!=="paid";if(q==="laptop_received")return tt.laptop_received&&!tt.laptop_returned;
    if(q==="laptop_returned")return tt.laptop_returned;if(q==="transferred")return tt.status==="transferred";
    if(q==="dropped")return tt.status==="dropped";
    return(tt.full_name?.toLowerCase().includes(q)||tt.department?.toLowerCase().includes(q)||tt.civil_id?.toLowerCase().includes(q)||tt.assigned_mentor?.toLowerCase().includes(q)||tt.laptop_serial?.toLowerCase().includes(q)||tt.university?.toLowerCase().includes(q));
  });

  const hh=String(currentTime.getHours()).padStart(2,"0");
  const mm=String(currentTime.getMinutes()).padStart(2,"0");
  const ss=String(currentTime.getSeconds()).padStart(2,"0");
  const clockColor=isLate?HW.red:"#34d399";

  const TraineeNav=()=>(<div style={{position:"fixed",bottom:0,left:0,right:0,background:HW.surface,borderTop:`1px solid ${HW.border}`,display:"flex",zIndex:100,paddingBottom:"env(safe-area-inset-bottom)"}}>{[{id:"attendance",icon:"✅",label:t.attend},{id:"weekly",icon:"📅",label:t.weekly},{id:"sick",icon:"🏥",label:t.sick},{id:"goals",icon:"🎯",label:t.goals}].map(tab=>(<button key={tab.id} onClick={()=>setTraineeTab(tab.id)} style={{flex:1,padding:"10px 4px 8px",border:"none",background:"none",cursor:"pointer",borderTop:traineeTab===tab.id?`3px solid ${HW.red}`:"3px solid transparent",display:"flex",flexDirection:"column",alignItems:"center",gap:3}}><span style={{fontSize:20}}>{tab.icon}</span><span style={{fontSize:9,fontWeight:700,color:traineeTab===tab.id?HW.red:HW.muted}}>{tab.label}</span></button>))}</div>);

  const MgmtNav=()=>(<div style={{position:"fixed",bottom:0,left:0,right:0,background:HW.surface,borderTop:`1px solid ${HW.border}`,display:"flex",zIndex:100,paddingBottom:"env(safe-area-inset-bottom)"}}>{["trainees","live","analytics","okr","access"].map(tab=>(<button key={tab} onClick={()=>{setMgmtTab(tab);setSelected(null);setMsg("");setSearchQuery("");}} style={{flex:1,padding:"10px 4px 8px",border:"none",background:"none",cursor:"pointer",borderTop:mgmtTab===tab?`3px solid ${HW.red}`:"3px solid transparent",display:"flex",flexDirection:"column",alignItems:"center",gap:3}}><span style={{fontSize:18}}>{tab==="trainees"?"👥":tab==="live"?"📡":tab==="analytics"?"📊":tab==="okr"?"🎯":"🔐"}</span><span style={{fontSize:9,fontWeight:700,color:mgmtTab===tab?HW.red:HW.muted}}>{tab==="trainees"?t.trainees:tab==="live"?t.live:tab==="analytics"?t.analytics:tab==="okr"?t.okr:t.access}</span></button>))}</div>);

  // ══ LOGIN ══
  if(view==="login")return(
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",padding:20,background:`linear-gradient(135deg,${HW.dark} 0%,${HW.surface} 50%,${HW.dark} 100%)`,fontFamily:isAr?"'Segoe UI',Tahoma,sans-serif":"sans-serif",color:HW.text,direction:dir}}>
      <div style={{width:"100%",maxWidth:420}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <HuaweiLogo size={64}/>
          <h1 style={{margin:"16px 0 4px",fontSize:28,color:HW.text}}>{t.appName}</h1>
          <p style={{color:HW.red,fontSize:12,fontWeight:700,letterSpacing:".1em",textTransform:"uppercase"}}>{t.poweredBy}</p>
          <div style={{fontSize:11,color:HW.muted,marginTop:6}}>🎨 {randomTheme.name}</div>
        </div>
        <div style={{display:"flex",justifyContent:"flex-end",marginBottom:16}}>
          <LangToggle lang={lang} setLang={setLang}/>
        </div>
        <label style={s.label}>{t.email}</label>
        <input style={{...s.input,marginBottom:16}} value={email} onChange={e=>setEmail(e.target.value)} placeholder="you@example.com" type="email" autoComplete="email"/>
        <label style={s.label}>{t.password}</label>
        <input style={{...s.input,marginBottom:24}} type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder={t.password} autoComplete="current-password" onKeyDown={e=>e.key==="Enter"&&login()}/>
        <button style={{...s.btn,background:HW.red,color:HW.white,width:"100%",padding:16,fontSize:17,opacity:loading?0.6:1}} onClick={login} disabled={loading}>{loading?t.signingIn:t.signIn}</button>
        <div style={{display:"flex",alignItems:"center",gap:10,margin:"20px 0"}}><div style={{flex:1,height:1,background:HW.border}}/><span style={{color:HW.muted,fontSize:12}}>OR</span><div style={{flex:1,height:1,background:HW.border}}/></div>
        <button style={{...s.btn,background:HW.surface2,color:HW.text,width:"100%",padding:16,border:`1px solid ${HW.border}`,fontSize:15}} onClick={()=>{setView("signup");setMsg("");}}>{t.createAccount}</button>
        {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,fontSize:14,marginTop:16,textAlign:"center"}}>{msg}</p>}
      </div>
    </div>
  );

  // ══ SIGNUP ══
  if(view==="signup")return(
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",padding:20,background:`linear-gradient(135deg,${HW.dark} 0%,${HW.surface} 50%,${HW.dark} 100%)`,fontFamily:isAr?"'Segoe UI',Tahoma,sans-serif":"sans-serif",color:HW.text,direction:dir}}>
      <div style={{width:"100%",maxWidth:420}}>
        <div style={{textAlign:"center",marginBottom:28}}><HuaweiLogo size={48}/><h2 style={{margin:"12px 0 4px",color:HW.text}}>{t.createAccount}</h2></div>
        <div style={{display:"flex",justifyContent:"flex-end",marginBottom:16}}><LangToggle lang={lang} setLang={setLang}/></div>
        <label style={s.label}>{t.email}</label>
        <input style={{...s.input,marginBottom:14}} value={email} onChange={e=>setEmail(e.target.value)} placeholder="yourname@example.com" type="email"/>
        <label style={s.label}>{t.password}</label>
        <input style={{...s.input,marginBottom:14}} type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder={isAr?"6 أحرف على الأقل":"Minimum 6 characters"}/>
        <label style={s.label}>{t.confirmPassword}</label>
        <input style={{...s.input,marginBottom:24}} type="password" value={confirmPwd} onChange={e=>setConfirmPwd(e.target.value)} placeholder={isAr?"أعد كتابة كلمة المرور":"Repeat your password"} onKeyDown={e=>e.key==="Enter"&&signup()}/>
        <button style={{...s.btn,background:HW.red,color:HW.white,width:"100%",padding:16,fontSize:16,opacity:loading?0.6:1}} onClick={signup} disabled={loading}>{loading?(isAr?"جاري الإنشاء…":"Creating account…"):`${t.createAccount} →`}</button>
        <button style={{...s.btn,background:"none",color:HW.muted,width:"100%",padding:14,marginTop:8,fontSize:15}} onClick={()=>{setView("login");setMsg("");}}>{isAr?"→ العودة لتسجيل الدخول":"← Back to Login"}</button>
        {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,fontSize:14,marginTop:12,textAlign:"center"}}>{msg}</p>}
      </div>
    </div>
  );

  // ══ SETUP ══
  if(view==="setup")return(
    <div style={{...s.page,paddingBottom:32}}>
      <div style={{display:"flex",justifyContent:"flex-end",marginBottom:12}}><LangToggle lang={lang} setLang={setLang}/></div>
      <div style={{textAlign:"center",marginBottom:24}}><HuaweiLogo size={52}/><h2 style={{margin:"12px 0 6px",fontSize:22,color:HW.text}}>{t.welcome}</h2><p style={{color:HW.muted,fontSize:13}}>{t.completeProfile}</p></div>
      <div style={s.card}>
        <h3 style={{marginBottom:16,color:HW.red,fontSize:16}}>{t.personalInfo}</h3>
        <label style={s.label}>{t.fullName} *</label><input style={{...s.input,marginBottom:12}} placeholder={isAr?"أحمد محمد الراشدي":"Ahmed Mohammed Al-Rashidi"} value={setupProfile.full_name} onChange={e=>setSetupProfile({...setupProfile,full_name:e.target.value})}/>
        <label style={s.label}>{t.civilId} *</label><input style={{...s.input,marginBottom:12}} placeholder="10234567" value={setupProfile.civil_id} onChange={e=>setSetupProfile({...setupProfile,civil_id:e.target.value})}/>
        <label style={s.label}>{t.dob}</label><input style={{...s.input,marginBottom:12}} type="date" value={setupProfile.date_of_birth} onChange={e=>setSetupProfile({...setupProfile,date_of_birth:e.target.value})}/>
        <label style={s.label}>{t.gender}</label>
        <select style={{...s.input,marginBottom:12}} value={setupProfile.gender} onChange={e=>setSetupProfile({...setupProfile,gender:e.target.value})}><option value="">{isAr?"اختر…":"Select…"}</option><option value="Male">{isAr?"ذكر":"Male"}</option><option value="Female">{isAr?"أنثى":"Female"}</option></select>
        <label style={s.label}>{t.nationality}</label><input style={s.input} placeholder={isAr?"مثال: عُماني":"e.g. Omani"} value={setupProfile.nationality} onChange={e=>setSetupProfile({...setupProfile,nationality:e.target.value})}/>
      </div>
      <div style={s.card}>
        <h3 style={{marginBottom:16,color:HW.red,fontSize:16}}>{t.academicInfo}</h3>
        <label style={s.label}>{t.university}</label><input style={{...s.input,marginBottom:12}} placeholder={isAr?"مثال: جامعة السلطان قابوس":"e.g. Sultan Qaboos University"} value={setupProfile.university} onChange={e=>setSetupProfile({...setupProfile,university:e.target.value})}/>
        <label style={s.label}>{t.department}</label><input style={{...s.input,marginBottom:12}} placeholder={isAr?"مثال: الهندسة":"e.g. Engineering"} value={setupProfile.department} onChange={e=>setSetupProfile({...setupProfile,department:e.target.value})}/>
        <label style={s.label}>{t.gpa}</label><input style={{...s.input,marginBottom:12}} type="number" placeholder="3.85" min="0" max="4" step="0.01" value={setupProfile.gpa} onChange={e=>setSetupProfile({...setupProfile,gpa:e.target.value})}/>
        <label style={s.label}>{t.ojtEndDate}</label><input style={{...s.input,marginBottom:12}} type="date" value={setupProfile.ojt_end_date} onChange={e=>setSetupProfile({...setupProfile,ojt_end_date:e.target.value})}/>
        <label style={s.label}>{t.mentor}</label><input style={s.input} placeholder={isAr?"مثال: د. فاطمة":"e.g. Dr. Fatima"} value={setupProfile.assigned_mentor} onChange={e=>setSetupProfile({...setupProfile,assigned_mentor:e.target.value})}/>
      </div>
      <div style={s.card}>
        <h3 style={{marginBottom:16,color:HW.red,fontSize:16}}>{t.contact}</h3>
        <label style={s.label}>{t.phone}</label><input style={s.input} placeholder="+968-9100-0001" value={setupProfile.phone_number} type="tel" onChange={e=>setSetupProfile({...setupProfile,phone_number:e.target.value})}/>
      </div>
      <button style={{...s.btn,background:HW.red,color:HW.white,width:"100%",padding:16,fontSize:16,opacity:loading?0.6:1}} onClick={saveSetupProfile} disabled={loading}>{loading?(isAr?"جاري الحفظ…":"Saving…"):t.completeProfileBtn}</button>
      {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,fontSize:14,marginTop:12,textAlign:"center"}}>{msg}</p>}
    </div>
  );

  // ══ TRAINEE ══
  if(view==="trainee")return(
    <div style={s.page}>
      {showReminder&&<ReminderPopup onDismiss={()=>setShowReminder(false)} lang={lang}/>}
      {showGreeting&&<TraineeGreetingPopup name={traineeName} onDismiss={()=>setShowGreeting(false)} lang={lang}/>}
      <TraineeNav/>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,paddingTop:4}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}><HuaweiLogo size={28}/><div><div style={{fontWeight:700,fontSize:15,color:HW.text}}>{traineeName?`${t.hi}, ${traineeName}!`:t.appName}</div><div style={{fontSize:10,color:HW.muted}}>{user?.email}</div></div></div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <LangToggle lang={lang} setLang={setLang}/>
          <div style={{textAlign:"right"}}><div style={{fontSize:18,fontWeight:800,color:clockColor,fontFamily:"monospace"}}>{hh}:{mm}:{ss}</div><div style={{fontSize:10,color:clockColor,fontWeight:700}}>{isLate?t.pastDeadline:t.onTime}</div></div>
        </div>
      </div>
      <div style={{display:"flex",gap:6,marginBottom:4}}>
        <button style={{...s.btn,background:HW.surface2,color:HW.muted,fontSize:11,padding:"6px 10px"}} onClick={()=>setShowGreeting(true)}>{t.greeting}</button>
        <button style={{...s.btn,background:HW.surface2,color:HW.muted,fontSize:11,padding:"6px 10px"}} onClick={()=>setShowReminder(true)}>{t.reminder}</button>
        <button style={{...s.btn,background:`${HW.red}20`,color:HW.red,fontSize:11,padding:"6px 10px",marginLeft:"auto"}} onClick={logout}>{t.signOut}</button>
      </div>
      <div style={{height:1,background:HW.border,margin:"12px 0 16px"}}/>

      {traineeTab==="attendance"&&(
        <div>
          <div style={{...s.card,textAlign:"center",border:`1px solid ${clockColor}40`,background:`linear-gradient(135deg,${HW.surface},${HW.red}08)`}}>
            <div style={{fontSize:48,fontWeight:800,color:clockColor,fontFamily:"monospace",letterSpacing:4,textShadow:`0 0 30px ${clockColor}40`}}>{hh}:{mm}:{ss}</div>
            <div style={{fontSize:12,color:HW.muted,marginTop:4}}>{currentTime.toLocaleDateString(isAr?"ar-SA":"en-GB",{weekday:"long",day:"numeric",month:"long"})}</div>
            <div style={{marginTop:10,display:"inline-block",padding:"5px 16px",borderRadius:20,fontWeight:700,fontSize:12,background:`${clockColor}15`,color:clockColor,border:`1px solid ${clockColor}40`}}>{isLate?(isAr?"⚠️ بعد الساعة 9 — مطلوب عذر":"⚠️ Past 9:00 AM — Excuse Required"):(isAr?"✅ في الوقت — جاهز للتسجيل":"✅ On Time — Ready to Sign In")}</div>
          </div>
          {isLate&&(<div style={{background:`${HW.red}15`,border:`1px solid ${HW.red}40`,borderRadius:12,padding:14,marginBottom:16,display:"flex",gap:10}}><span style={{fontSize:24}}>🔒</span><div><div style={{fontWeight:700,color:HW.red,fontSize:14}}>{t.timeLocked}</div><div style={{fontSize:12,color:HW.muted,marginTop:2}}>{t.deadlineWas} <b style={{color:HW.red}}>{t.penaltyWillApply}</b></div></div></div>)}
          <div style={s.card}>
            <h3 style={{marginBottom:14,color:HW.red,fontSize:16}}>{t.signAttendance}</h3>
            <label style={s.label}>{t.date}</label><input style={{...s.input,marginBottom:14}} type="date" value={attendance.report_date} onChange={e=>setAttendance({...attendance,report_date:e.target.value})}/>
            <label style={s.label}>{t.signinTime}</label>
            <div style={{background:HW.surface2,border:`1px solid ${HW.border}`,borderRadius:10,padding:"12px 14px",fontSize:20,color:clockColor,fontWeight:800,fontFamily:"monospace",marginBottom:14}}>{hh}:{mm}:{ss}</div>
            <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:14,background:HW.surface2,borderRadius:10,padding:14}}>
              <input type="checkbox" id="att" checked={attendance.attended} style={{width:22,height:22}} onChange={e=>setAttendance({...attendance,attended:e.target.checked})}/>
              <label htmlFor="att" style={{fontSize:15,fontWeight:600,cursor:"pointer",color:HW.text}}>{t.confirmAttendance}</label>
            </div>
            {!attendance.attended&&(<div style={{background:`${HW.red}15`,border:`1px solid ${HW.red}40`,borderRadius:10,padding:12,marginBottom:14,fontSize:13,color:HW.red}}>{t.penaltyWarning}</div>)}
            {isLate&&(<div style={{background:HW.surface2,borderRadius:12,padding:14,border:`1px solid ${HW.red}40`,marginBottom:14}}>
              <div style={{fontWeight:700,color:HW.red,marginBottom:12,fontSize:15}}>{t.excuseRequired}</div>
              <label style={s.label}>{t.excuseType}</label>
              <select style={{...s.input,marginBottom:12}} value={excuseType} onChange={e=>setExcuseType(e.target.value)}>
                <option value="">{t.selectReason}</option>
                <option value="traffic">{t.traffic} {trafficCount>=2?t.limitReached:""}</option>
                <option value="medical">{t.medical}</option>
                <option value="family">{t.family}</option>
                <option value="other">{t.other}</option>
              </select>
              <label style={s.label}>{t.description} *</label>
              <textarea style={{...s.input,height:90,resize:"vertical",marginBottom:12}} placeholder={isAr?"يرجى وصف سببك…":"Please describe your reason…"} value={excuseText} onChange={e=>setExcuseText(e.target.value)}/>
              <label style={s.label}>{t.proofPhoto} {excuseType==="traffic"?t.required:t.optional}</label>
              <div style={{display:"flex",gap:10,alignItems:"center"}}><button style={{...s.btn,background:HW.surface,color:HW.text,border:`1px dashed ${HW.border}`,flex:1}} onClick={()=>excuseRef.current.click()}>📷 {excusePhoto?t.changePhoto:t.uploadProof}</button><input ref={excuseRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleExcusePhoto}/>{excusePreview&&<img src={excusePreview} alt="proof" style={{width:60,height:60,objectFit:"cover",borderRadius:8}}/>}</div>
            </div>)}
            <div style={{background:HW.surface2,borderRadius:12,padding:14,marginBottom:14,border:locationOk?`1px solid ${HW.red}50`:`1px solid ${HW.border}`}}>
              <div style={{fontWeight:700,color:HW.red,marginBottom:10,fontSize:15}}>{t.locationVerification}</div>
              <p style={{fontSize:12,color:HW.muted,marginBottom:10}}>{t.mustBeWithin} <b style={{color:HW.text}}>{WORK_RADIUS}m</b> {t.ofWorkplace}</p>
              <button style={{...s.btn,background:locationOk?`${HW.red}20`:HW.red,color:locationOk?HW.red:HW.white,width:"100%",opacity:geoStatus==="checking"?0.6:1}} onClick={checkLocation} disabled={geoStatus==="checking"}>{geoStatus==="checking"?t.checkingLocation:locationOk?t.locationVerified:t.verifyLocation}</button>
              {geoMsg&&<p style={{fontSize:12,marginTop:8,color:geoStatus==="ok"?"#34d399":HW.red}}>{geoMsg}</p>}
            </div>
            <button style={{...s.btn,background:locationOk?HW.red:HW.surface2,color:locationOk?HW.white:HW.muted,width:"100%",padding:16,fontSize:16,opacity:loading?0.6:1,cursor:locationOk?"pointer":"not-allowed"}} onClick={submitAttendance} disabled={loading||!locationOk}>{loading?t.saving:t.submitAttendance}</button>
            <div style={{marginTop:12,background:HW.surface2,borderRadius:12,padding:14,border:`1px solid ${HW.border}`}}>
              <div style={{fontWeight:700,color:"#4f8ef7",marginBottom:10,fontSize:15}}>{t.signOutBtn}</div>
              {signedOut?(<div style={{display:"flex",alignItems:"center",gap:10}}><div style={{width:44,height:44,borderRadius:"50%",background:"rgba(79,142,247,.15)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22}}>✅</div><div><div style={{fontWeight:700,color:"#4f8ef7",fontSize:15}}>{t.signedOutAt} {signoutTime}</div><div style={{fontSize:12,color:HW.muted}}>{t.timeRecorded}</div></div></div>):(<button style={{...s.btn,background:"rgba(79,142,247,.15)",color:"#4f8ef7",width:"100%",padding:14,fontSize:15,border:"1px solid rgba(79,142,247,.3)",opacity:loading?0.6:1}} onClick={submitSignOut} disabled={loading}>{t.recordTimeOut}</button>)}
            </div>
            <div style={{marginTop:12,background:HW.surface2,borderRadius:12,padding:14,border:`1px solid ${HW.border}`}}>
              <div style={{fontWeight:700,color:HW.text,marginBottom:12,fontSize:15}}>{t.myStatus}</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
                <div style={{background:HW.surface,borderRadius:10,padding:12,textAlign:"center"}}><div style={{fontSize:20,marginBottom:4}}>💰</div><div style={{fontSize:11,color:HW.muted,marginBottom:4}}>{t.payment}</div><div style={{fontSize:13,fontWeight:800,color:myTraineeData?.payment_status==="paid"?"#34d399":HW.red}}>{myTraineeData?.payment_status==="paid"?t.paid:t.unpaid}</div>{myTraineeData?.payment_notes&&<div style={{fontSize:11,color:HW.muted,marginTop:4}}>{myTraineeData.payment_notes}</div>}</div>
                <div style={{background:HW.surface,borderRadius:10,padding:12,textAlign:"center"}}><div style={{fontSize:20,marginBottom:4}}>💻</div><div style={{fontSize:11,color:HW.muted,marginBottom:4}}>{t.laptop}</div><div style={{fontSize:13,fontWeight:800,color:myTraineeData?.laptop_returned?"#FFA500":myTraineeData?.laptop_received?"#34d399":HW.muted}}>{myTraineeData?.laptop_returned?t.returned:myTraineeData?.laptop_received?t.received:t.notYet}</div>{myTraineeData?.laptop_serial&&<div style={{fontSize:11,color:HW.muted,marginTop:4}}>SN: {myTraineeData.laptop_serial}</div>}</div>
              </div>
              {myTraineeData?.ojt_end_date&&(<div style={{background:HW.surface,borderRadius:10,padding:12,textAlign:"center",marginBottom:10}}><div style={{fontSize:11,color:HW.muted,marginBottom:4}}>{t.ojtEndDateLabel}</div><div style={{fontSize:13,fontWeight:800,color:HW.red}}>{new Date(myTraineeData.ojt_end_date).toLocaleDateString(isAr?"ar-SA":"en-GB",{day:"numeric",month:"long",year:"numeric"})}</div>{(()=>{const days=Math.ceil((new Date(myTraineeData.ojt_end_date)-new Date())/(1000*60*60*24));return days>0?<div style={{fontSize:11,color:HW.muted,marginTop:4}}>{days} {t.daysRemaining}</div>:<div style={{fontSize:11,color:HW.red,marginTop:4}}>{t.ojtEnded}</div>;})()}</div>)}
              {myTraineeData?.laptop_received&&!myTraineeData?.laptop_returned&&(<button style={{...s.btn,background:"rgba(255,165,0,.15)",color:"#FFA500",width:"100%",padding:12,fontSize:14,marginBottom:8,border:"1px solid rgba(255,165,0,.3)",opacity:loading?0.6:1}} onClick={()=>traineeLaptopAction("returned")} disabled={loading}>{t.returnedLaptop}</button>)}
              {!myTraineeData?.laptop_received&&(<button style={{...s.btn,background:"rgba(52,211,153,.15)",color:"#34d399",width:"100%",padding:12,fontSize:14,border:"1px solid rgba(52,211,153,.3)",opacity:loading?0.6:1}} onClick={()=>traineeLaptopAction("received")} disabled={loading}>{t.receivedLaptop}</button>)}
            </div>
            {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,fontSize:14,marginTop:12,textAlign:"center"}}>{msg}</p>}
          </div>
        </div>
      )}

      {traineeTab==="weekly"&&(
        <div>
          <div style={{...s.card,background:`linear-gradient(135deg,${HW.red}15,${HW.red}05)`,border:`1px solid ${HW.red}40`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><div style={{fontSize:11,color:HW.muted,fontWeight:700,textTransform:"uppercase",marginBottom:4}}>{t.currentWeek}</div><div style={{fontSize:18,fontWeight:800,color:HW.text}}>{currentWeek.label}</div><div style={{fontSize:12,color:HW.muted,marginTop:2}}>{t.sundayThursday}</div></div>
              <div style={{background:weeklySubmitted?"rgba(52,211,153,.15)":"rgba(255,165,0,.15)",border:weeklySubmitted?"1px solid rgba(52,211,153,.3)":"1px solid rgba(255,165,0,.3)",borderRadius:12,padding:"10px 14px",textAlign:"center"}}><div style={{fontSize:20}}>{weeklySubmitted?"✅":"📝"}</div><div style={{fontSize:11,fontWeight:700,marginTop:4,color:weeklySubmitted?"#34d399":"#FFA500"}}>{weeklySubmitted?t.submitted:t.pending}</div></div>
            </div>
          </div>
          {weeklySubmitted&&aiResult?(<div><div style={{...s.card,border:"1px solid rgba(52,211,153,.3)"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}><h3 style={{margin:0,color:"#34d399",fontSize:16}}>{t.reportSubmitted}</h3><button style={{...s.btn,background:`${HW.red}15`,color:HW.red,fontSize:12,padding:"6px 12px"}} onClick={()=>setWeeklySubmitted(false)}>{t.edit}</button></div><div style={{background:HW.surface2,borderRadius:10,padding:12,borderLeft:`3px solid ${HW.red}`,fontSize:13,lineHeight:1.7,color:HW.text}}>{weeklyText}</div></div><div style={{...s.card,border:`1px solid ${HW.red}40`}}><h3 style={{marginBottom:16,color:HW.red,fontSize:16}}>{t.aiAnalysis}</h3><div style={{background:HW.surface2,borderRadius:12,padding:16,textAlign:"center",marginBottom:16}}><div style={{fontSize:11,color:HW.muted,fontWeight:700,textTransform:"uppercase",marginBottom:8}}>{t.weeklyKpi}</div><div style={{fontSize:56,fontWeight:800,color:kpiColor(aiResult.kpi_score)}}>{aiResult.kpi_score}</div><div style={{fontSize:11,color:HW.muted}}>{t.outOf}</div></div>{aiResult.pie_chart&&<PieChart data={aiResult.pie_chart}/>}{aiResult.summary&&<div style={{background:HW.surface2,borderRadius:10,padding:12,marginTop:12,borderLeft:`3px solid ${HW.red}`}}><div style={{fontSize:10,color:HW.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>{t.weekSummary}</div><div style={{fontSize:13,color:HW.text}}>{aiResult.summary}</div></div>}{aiResult.talent_notes&&<div style={{background:HW.surface2,borderRadius:10,padding:12,marginTop:10,borderLeft:"3px solid #FFA500"}}><div style={{fontSize:10,color:"#FFA500",fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>{t.talentNotes}</div><div style={{fontSize:13,lineHeight:1.6,color:HW.text}}>{aiResult.talent_notes}</div></div>}</div></div>):(<div style={s.card}><h3 style={{marginBottom:8,color:HW.red,fontSize:16}}>{t.thisWeekTasks}</h3><p style={{color:HW.muted,fontSize:13,marginBottom:12,lineHeight:1.5}}>{t.writeEachDay}</p><div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:12}}>{[isAr?"أح":"Sun",isAr?"إث":"Mon",isAr?"ثل":"Tue",isAr?"أر":"Wed",isAr?"خم":"Thu"].map(day=>(<span key={day} style={{background:HW.surface2,borderRadius:6,padding:"4px 10px",fontSize:12,color:HW.muted,border:`1px solid ${HW.border}`}}>{day}</span>))}</div><label style={s.label}>{t.weeklyTasksLabel} *</label><textarea style={{...s.input,height:200,resize:"vertical",marginBottom:14,fontSize:14,lineHeight:1.6}} placeholder={isAr?"الأحد: ...\nالإثنين: ...\nالثلاثاء: ...\nالأربعاء: ...\nالخميس: ...":"Sunday: ...\nMonday: ...\nTuesday: ...\nWednesday: ...\nThursday: ..."} value={weeklyText} onChange={e=>setWeeklyText(e.target.value)}/><label style={s.label}>{t.proofPhotoOpt}</label><div style={{display:"flex",gap:10,alignItems:"center",marginBottom:16}}><button style={{...s.btn,background:HW.surface2,color:HW.text,border:`1px dashed ${HW.border}`,flex:1}} onClick={()=>weeklyPhotoRef.current.click()}>{weeklyPhotoFile?`📷 ${t.changePhoto}`:`📷 ${t.uploadProof}`}</button><input ref={weeklyPhotoRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleWeeklyPhoto}/>{weeklyPhotoPreview&&<img src={weeklyPhotoPreview} alt="proof" style={{width:70,height:70,objectFit:"cover",borderRadius:10,border:`2px solid ${HW.border}`}}/>}</div><button style={{...s.btn,background:HW.red,color:HW.white,width:"100%",padding:16,fontSize:16,opacity:(loading||aiLoading)?0.6:1}} onClick={submitWeeklyReport} disabled={loading||aiLoading}>{aiLoading?t.aiAnalyzing:loading?t.saving:t.submitWeekly}</button>{msg&&<p style={{color:msg.startsWith("✅")?"#34d399":msg.includes("🤖")?"#FFA500":HW.red,fontSize:14,marginTop:12,textAlign:"center"}}>{msg}</p>}</div>)}
        </div>
      )}

      {traineeTab==="sick"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:16}}>{[{label:t.total,value:sickLeaves.length,color:"#4f8ef7"},{label:t.freeDays,value:sickLeaves.reduce((a,sl)=>a+Math.min(sl.total_days,2),0),color:"#34d399"},{label:t.penaltyDays,value:sickLeaves.reduce((a,sl)=>a+(sl.penalty_days||0),0),color:HW.red}].map((stat,i)=>(<div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,borderRadius:12,padding:14,textAlign:"center",borderTop:`3px solid ${stat.color}`}}><div style={{fontSize:24,fontWeight:800,color:stat.color}}>{stat.value}</div><div style={{fontSize:11,color:HW.muted,marginTop:3}}>{stat.label}</div></div>))}</div>
          <div style={{...s.card,background:"rgba(79,142,247,.08)",border:"1px solid rgba(79,142,247,.3)"}}><div style={{fontSize:14,fontWeight:700,color:"#4f8ef7",marginBottom:6}}>{t.sickLeavePolicy}</div><div style={{fontSize:13,color:HW.muted,lineHeight:1.6}}>{t.first2days}<br/>{t.after48}</div></div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}><h3 style={{margin:0,fontSize:16,color:HW.text}}>{t.mySickLeaves} ({sickLeaves.length})</h3><button style={{...s.btn,background:HW.red,color:HW.white,padding:"10px 16px"}} onClick={()=>setShowAddSick(!showAddSick)}>{showAddSick?t.cancel:t.submitSickLeave}</button></div>
          {showAddSick&&(<div style={{...s.card,border:`1px solid ${HW.red}40`,marginBottom:12}}><h4 style={{marginBottom:14,color:HW.red,fontSize:15}}>🏥 {isAr?"تقديم إجازة مرضية":"Submit Sick Leave"}</h4><label style={s.label}>{t.startDate}</label><input style={{...s.input,marginBottom:12}} type="date" value={newSick.start_date} onChange={e=>setNewSick({...newSick,start_date:e.target.value})}/><label style={s.label}>{t.endDate}</label><input style={{...s.input,marginBottom:12}} type="date" value={newSick.end_date} onChange={e=>setNewSick({...newSick,end_date:e.target.value})}/>{newSick.start_date&&newSick.end_date&&(()=>{const days=Math.ceil((new Date(newSick.end_date)-new Date(newSick.start_date))/(1000*60*60*24))+1;const penDays=Math.max(0,days-2);return(<div style={{background:penDays>0?`${HW.red}15`:"rgba(52,211,153,.1)",border:penDays>0?`1px solid ${HW.red}40`:"1px solid rgba(52,211,153,.3)",borderRadius:10,padding:12,marginBottom:12}}><div style={{fontSize:13,fontWeight:700,color:penDays>0?HW.red:"#34d399"}}>{days} {isAr?"أيام":"day(s)"} — {penDays>0?`⚠️ ${penDays} ${isAr?"أيام خصم":"penalty day(s)"}`:`✅ ${isAr?"بدون خصم":"No penalty"}`}</div></div>);})()}<label style={s.label}>{t.reason} *</label><textarea style={{...s.input,height:80,resize:"vertical",marginBottom:12}} placeholder={isAr?"صف مرضك…":"Describe your illness…"} value={newSick.reason} onChange={e=>setNewSick({...newSick,reason:e.target.value})}/><label style={s.label}>{t.medicalProof}</label><div style={{display:"flex",gap:10,alignItems:"center",marginBottom:14}}><button style={{...s.btn,background:HW.surface2,color:HW.text,border:`1px dashed ${HW.border}`,flex:1}} onClick={()=>sickProofRef.current.click()}>{sickProofFile?`📷 ${t.changePhoto}`:`📷 ${t.uploadProof}`}</button><input ref={sickProofRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleSickProof}/>{sickProofPreview&&<img src={sickProofPreview} alt="proof" style={{width:70,height:70,objectFit:"cover",borderRadius:10,border:`2px solid ${HW.border}`}}/>}</div><button style={{...s.btn,background:HW.red,color:HW.white,width:"100%",padding:14,opacity:loading?0.6:1}} onClick={submitSickLeave} disabled={loading}>{loading?t.submitting:t.submitSickBtn}</button>{msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,fontSize:13,marginTop:8,textAlign:"center"}}>{msg}</p>}</div>)}
          {sickLeaves.length===0?(<div style={{...s.card,textAlign:"center",padding:32}}><div style={{fontSize:36,marginBottom:10}}>🏥</div><p style={{color:HW.muted}}>{t.noSickLeaves}</p></div>):(sickLeaves.map(sl=>(<div key={sl.id} style={{...s.card,borderLeft:`4px solid ${sl.penalty_applied?HW.red:"#34d399"}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}><div><div style={{fontWeight:700,fontSize:14,color:HW.text}}>📅 {sl.start_date} → {sl.end_date}</div><div style={{fontSize:12,color:HW.muted,marginTop:2}}>{sl.reason}</div></div><div style={{textAlign:"right"}}><div style={{fontSize:13,fontWeight:700,color:sl.penalty_applied?HW.red:"#34d399"}}>{sl.total_days} {isAr?"أيام":"day(s)"}</div><div style={{fontSize:11,color:HW.muted}}>{sl.penalty_applied?`⚠️ -${(sl.penalty_days*PENALTY_PCT).toFixed(2)}%`:t.nopenalty}</div></div></div>{sl.proof_url&&<div style={{fontSize:12,color:"#4f8ef7"}}>📎 {isAr?"إثبات طبي مرفق":"Medical proof attached"}</div>}</div>)))}
        </div>
      )}

      {traineeTab==="goals"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginBottom:16}}>{[{label:t.totalGoals,value:goals.length,color:HW.red},{label:t.inProgress,value:goals.filter(g=>g.status==="in_progress").length,color:"#4f8ef7"},{label:t.completed,value:goals.filter(g=>g.status==="completed").length,color:"#34d399"},{label:t.overdue,value:goals.filter(g=>g.due_date&&new Date(g.due_date)<new Date()&&g.status!=="completed").length,color:"#f87171"}].map((stat,i)=>(<div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,borderRadius:12,padding:14,textAlign:"center",borderTop:`3px solid ${stat.color}`}}><div style={{fontSize:26,fontWeight:800,color:stat.color}}>{stat.value}</div><div style={{fontSize:11,color:HW.muted,marginTop:3}}>{stat.label}</div></div>))}</div>
          <div style={{display:"flex",gap:8,overflowX:"auto",paddingBottom:8,marginBottom:12}}><button style={{...s.btn,background:goalFilter==="all"?HW.red:HW.surface2,color:goalFilter==="all"?HW.white:HW.muted,padding:"8px 14px",fontSize:12,whiteSpace:"nowrap"}} onClick={()=>setGoalFilter("all")}>{isAr?"الكل":"All"}</button>{KRA_CATEGORIES.map(kra=>(<button key={kra.id} style={{...s.btn,background:goalFilter===kra.id?kra.color:HW.surface2,color:goalFilter===kra.id?HW.white:HW.muted,padding:"8px 14px",fontSize:12,whiteSpace:"nowrap",border:goalFilter===kra.id?"none":`1px solid ${HW.border}`}} onClick={()=>setGoalFilter(goalFilter===kra.id?"all":kra.id)}>{kra.icon} {(isAr?kra.labelAr:kra.label).split(" ")[0]}</button>))}</div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}><div style={{fontSize:14,fontWeight:700,color:HW.text}}>{filteredGoals.length} {isAr?"هدف":"goal(s)"}</div><button style={{...s.btn,background:HW.red,color:HW.white,padding:"10px 16px"}} onClick={()=>setShowAddGoal(!showAddGoal)}>{showAddGoal?t.cancel:t.addGoal}</button></div>
          {showAddGoal&&(<div style={{...s.card,border:`1px solid ${HW.red}40`,marginBottom:12}}><h4 style={{marginBottom:14,color:HW.red,fontSize:15}}>{t.newGoal}</h4><label style={s.label}>{t.kraCategory}</label><select style={{...s.input,marginBottom:12}} value={newGoal.kra} onChange={e=>setNewGoal({...newGoal,kra:e.target.value})}>{KRA_CATEGORIES.map(k=><option key={k.id} value={k.id}>{k.icon} {isAr?k.labelAr:k.label}</option>)}</select><label style={s.label}>{t.goalTitle} *</label><input style={{...s.input,marginBottom:12}} placeholder={isAr?"مثال: تحقيق حضور 95%":"e.g. Achieve 95% attendance"} value={newGoal.goal_title} onChange={e=>setNewGoal({...newGoal,goal_title:e.target.value})}/><label style={s.label}>{isAr?"الوصف":"Description"}</label><textarea style={{...s.input,height:70,resize:"vertical",marginBottom:12}} placeholder={isAr?"كيف تخطط لتحقيق هذا؟":"How do you plan to achieve this?"} value={newGoal.description} onChange={e=>setNewGoal({...newGoal,description:e.target.value})}/><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}><div><label style={s.label}>{t.target}</label><input style={s.input} type="number" value={newGoal.target_value} onChange={e=>setNewGoal({...newGoal,target_value:parseFloat(e.target.value)})}/></div><div><label style={s.label}>{t.unit}</label><select style={s.input} value={newGoal.unit} onChange={e=>setNewGoal({...newGoal,unit:e.target.value})}><option value="%">%</option><option value="days">{isAr?"أيام":"Days"}</option><option value="sessions">{isAr?"جلسات":"Sessions"}</option><option value="tasks">{isAr?"مهام":"Tasks"}</option><option value="hours">{isAr?"ساعات":"Hours"}</option><option value="score">{isAr?"نقاط":"Score"}</option></select></div></div><label style={s.label}>{t.dueDate}</label><input style={{...s.input,marginBottom:14}} type="date" value={newGoal.due_date} onChange={e=>setNewGoal({...newGoal,due_date:e.target.value})}/><button style={{...s.btn,background:HW.red,color:HW.white,width:"100%",padding:14,opacity:loading?0.6:1}} onClick={addGoal} disabled={loading}>{loading?t.saving:t.setGoal}</button></div>)}
          {filteredGoals.length===0?(<div style={{...s.card,textAlign:"center",padding:32}}><div style={{fontSize:36,marginBottom:10}}>🎯</div><p style={{color:HW.muted}}>{t.noGoals}</p></div>):(filteredGoals.map(goal=><GoalCard key={goal.id} goal={goal} onUpdate={updateGoal} onDelete={deleteGoal} isTrainee={true} lang={lang}/>))}
          {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,fontSize:14,marginTop:8,textAlign:"center"}}>{msg}</p>}
        </div>
      )}
    </div>
  );

  // ══ MANAGEMENT ══
  if(view==="mgmt"){
    const analytics=getAnalytics();
    return(
      <div style={s.page}>
        {showMgrGreeting&&<ManagerGreetingPopup onDismiss={()=>setShowMgrGreeting(false)} lang={lang} onLogin={async(name)=>{setCurrentManagerName(name);await writeAccessLog(user?.email,name,"login",`${name} logged into Management Hub`);setTimeout(fetchAccessLogs,500);}}/>}
        {showTransferPopup&&selected&&<TransferPopup trainee={selected} onConfirm={handleTransferConfirm} onCancel={()=>setShowTransferPopup(false)} lang={lang}/>}
        {showDropPopup&&selected&&<DropPopup trainee={selected} onConfirm={handleDropConfirm} onCancel={()=>setShowDropPopup(false)} lang={lang}/>}
        {showAttendanceExport&&(
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.88)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000,backdropFilter:"blur(6px)",padding:16}}>
            <div style={{background:HW.surface,border:`2px solid ${HW.red}`,borderRadius:24,padding:28,width:"100%",maxWidth:380,textAlign:"center",direction:dir}}>
              <div style={{fontSize:36,marginBottom:12}}>📋</div>
              <h3 style={{fontSize:20,fontWeight:800,color:HW.red,margin:"0 0 6px"}}>{t.exportAttendance}</h3>
              <p style={{fontSize:13,color:HW.muted,marginBottom:20}}>{t.selectMonthYear}</p>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:20}}>
                <div>
                  <label style={s.label}>{t.month}</label>
                  <select style={{...s.input}} value={attendanceExportMonth} onChange={e=>setAttendanceExportMonth(e.target.value)}>
                    {["01","02","03","04","05","06","07","08","09","10","11","12"].map((m,i)=>(
                      <option key={m} value={m}>{new Date(2026,i,1).toLocaleDateString(isAr?"ar-SA":"en-GB",{month:"long"})}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label style={s.label}>{t.year}</label>
                  <select style={{...s.input}} value={attendanceExportYear} onChange={e=>setAttendanceExportYear(e.target.value)}>
                    {["2025","2026","2027"].map(y=><option key={y} value={y}>{y}</option>)}
                  </select>
                </div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                <button style={{padding:14,borderRadius:12,border:`1px solid ${HW.border}`,background:HW.surface2,color:HW.muted,fontWeight:700,fontSize:14,cursor:"pointer"}} onClick={()=>setShowAttendanceExport(false)}>{t.cancel}</button>
                <button style={{padding:14,borderRadius:12,border:"none",background:HW.red,color:HW.white,fontWeight:800,fontSize:14,cursor:"pointer"}} onClick={exportAttendanceSheets}>{t.export}</button>
              </div>
            </div>
          </div>
        )}
        <MgmtNav/>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,paddingTop:4}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}><HuaweiLogo size={28}/><div><div style={{fontWeight:700,fontSize:14,color:HW.text}}>{t.appName} {currentManagerName&&`— ${currentManagerName}`}</div><div style={{fontSize:10,color:HW.muted}}>{user?.email}</div></div></div>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <LangToggle lang={lang} setLang={setLang}/>
            <div style={{textAlign:"right"}}><div style={{fontSize:16,fontWeight:800,color:clockColor,fontFamily:"monospace"}}>{hh}:{mm}:{ss}</div><div style={{fontSize:10,color:clockColor,fontWeight:700}}>{isLate?t.pastDeadline:t.onTime}</div></div>
          </div>
        </div>
        <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap"}}>
          <button style={{...s.btn,background:`${HW.red}20`,color:HW.red,fontSize:12,padding:"8px 12px"}} onClick={exportExcel}>{t.exportExcel}</button>
          <button style={{...s.btn,background:`${HW.red}20`,color:HW.red,fontSize:12,padding:"8px 12px"}} onClick={()=>setShowAttendanceExport(true)}>{t.attendanceSheets}</button>
          <button style={{...s.btn,background:HW.surface2,color:HW.muted,fontSize:12,padding:"8px 12px"}} onClick={()=>setShowMgrGreeting(true)}>{t.greeting}</button>
          <button style={{...s.btn,background:`${HW.red}20`,color:HW.red,fontSize:12,padding:"8px 12px",marginLeft:"auto"}} onClick={logout}>{t.signOut}</button>
        </div>
        <div style={{height:1,background:HW.border,marginBottom:16}}/>

        {mgmtTab==="trainees"&&!selected&&(
          <div>
            <h3 style={{marginBottom:12,fontSize:16,color:HW.text}}>{t.allTrainees}</h3>
            <div style={{position:"relative",marginBottom:12}}>
              <input style={{...s.input}} placeholder={t.searchPlaceholder} value={searchQuery} onChange={e=>setSearchQuery(e.target.value)}/>
              {searchQuery&&<button onClick={()=>setSearchQuery("")} style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:HW.muted,cursor:"pointer",fontSize:18}}>✕</button>}
            </div>
            <div style={{display:"flex",gap:8,marginBottom:14,overflowX:"auto",paddingBottom:4}}>
              {[{key:"",label:`👥 ${isAr?"الكل":"All"}`},{key:"active",label:`✅ ${t.active}`},{key:"transferred",label:`🔄 ${t.transferred}`},{key:"dropped",label:`🔴 ${t.dropped}`},{key:"paid",label:`💰 ${isAr?"مدفوع":"Paid"}`},{key:"unpaid",label:`❌ ${isAr?"غير مدفوع":"Unpaid"}`},{key:"laptop_received",label:`💻 ${isAr?"لديه حاسوب":"Has Laptop"}`}].map(f=>(
                <button key={f.key} style={{...s.btn,padding:"6px 12px",fontSize:11,whiteSpace:"nowrap",background:searchQuery===f.key?HW.red:HW.surface2,color:searchQuery===f.key?HW.white:HW.muted,border:`1px solid ${HW.border}`}} onClick={()=>setSearchQuery(searchQuery===f.key?"":f.key)}>{f.label}</button>
              ))}
            </div>
            <div style={{fontSize:12,color:HW.muted,marginBottom:10}}>{filteredTrainees.length} {t.of} {trainees.length} {isAr?"متدرب":"trainees"}</div>
            {filteredTrainees.length===0?(<div style={{...s.card,textAlign:"center",padding:32}}><p style={{color:HW.muted}}>{t.noTrainees}</p></div>):(
              filteredTrainees.map(tt=>(
                <div key={tt.id} style={{...s.card,cursor:"pointer",opacity:tt.status==="dropped"?0.6:1,display:"flex",alignItems:"center",gap:12,padding:14}} onClick={()=>openProfile(tt)}>
                  <div style={{width:44,height:44,borderRadius:12,background:`linear-gradient(135deg,${HW.red},${HW.darkRed})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,fontWeight:800,color:HW.white,flexShrink:0}}>{tt.full_name?.split(" ").map(w=>w[0]).join("").slice(0,2)}</div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontWeight:700,fontSize:15,color:HW.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{tt.full_name}</div>
                    <div style={{fontSize:12,color:HW.muted,marginTop:2}}>{tt.department}{tt.university?` · ${tt.university}`:""} · GPA {tt.gpa||"—"}</div>
                    {tt.status==="transferred"&&<div style={{fontSize:11,color:"#FFA500",marginTop:2}}>🔄 {isAr?"منقول إلى:":"Transferred to:"} <b>{tt.department}</b></div>}
                    {tt.status==="dropped"&&tt.quitting_reason&&<div style={{fontSize:11,color:"#666",marginTop:2}}>🔴 {tt.quitting_reason.substring(0,50)}</div>}
                    <div style={{display:"flex",gap:6,marginTop:4,flexWrap:"wrap"}}>
                      <span style={{fontSize:10,fontWeight:700,padding:"2px 6px",borderRadius:6,color:tt.payment_status==="paid"?"#34d399":HW.red,background:tt.payment_status==="paid"?"rgba(52,211,153,.1)":`${HW.red}15`}}>💰 {tt.payment_status||"unpaid"}</span>
                      <span style={{fontSize:10,fontWeight:700,padding:"2px 6px",borderRadius:6,color:tt.laptop_returned?"#FFA500":tt.laptop_received?"#34d399":HW.muted,background:tt.laptop_returned?"rgba(255,165,0,.1)":tt.laptop_received?"rgba(52,211,153,.1)":"rgba(136,136,136,.1)"}}>💻 {tt.laptop_returned?(isAr?"مُرجع":"returned"):tt.laptop_received?(isAr?"مُستلم":"received"):(isAr?"غير مستلم":"not received")}</span>
                      {tt.laptop_serial&&<span style={{fontSize:10,fontWeight:700,padding:"2px 6px",borderRadius:6,color:"#4f8ef7",background:"rgba(79,142,247,.1)"}}>SN: {tt.laptop_serial}</span>}
                    </div>
                  </div>
                  <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6,flexShrink:0}}>
                    <span style={{padding:"3px 8px",borderRadius:20,fontSize:10,fontWeight:700,background:statusColors[tt.status]?.bg,color:statusColors[tt.status]?.color}}>{tt.status}</span>
                    <button style={{...s.btn,background:`${HW.red}15`,color:HW.red,fontSize:11,padding:"4px 8px"}} onClick={e=>{e.stopPropagation();exportPDF(tt);}}>📄 PDF</button>
                  </div>
                </div>
              ))
            )}
          </div>
        )}

        {mgmtTab==="trainees"&&selected&&(
          <div>
            <button style={{...s.btn,background:HW.surface2,color:HW.text,marginBottom:12,border:`1px solid ${HW.border}`,fontSize:13}} onClick={()=>{setSelected(null);setMsg("");}}>{t.back}</button>
            <div style={s.card}>
              <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:12}}>
                <div style={{width:52,height:52,borderRadius:14,background:`linear-gradient(135deg,${HW.red},${HW.darkRed})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,fontWeight:800,color:HW.white,flexShrink:0}}>{selected.full_name?.split(" ").map(w=>w[0]).join("").slice(0,2)}</div>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,fontSize:17,color:HW.text}}>{selected.full_name}</div>
                  <div style={{fontSize:12,color:HW.muted}}>{selected.department}{selected.university?` · ${selected.university}`:""}</div>
                  {selected.status==="transferred"&&<div style={{fontSize:12,color:"#FFA500",marginTop:2}}>🔄 {isAr?"منقول إلى:":"Transferred to:"} <b>{selected.department}</b> {selected.assigned_mentor&&`· ${selected.assigned_mentor}`}</div>}
                  {selected.status==="dropped"&&selected.quitting_date&&<div style={{fontSize:12,color:"#666",marginTop:2}}>🔴 {selected.quitting_date}{selected.quitting_reason&&` — ${selected.quitting_reason}`}</div>}
                  <span style={{padding:"3px 8px",borderRadius:20,fontSize:11,fontWeight:700,marginTop:4,display:"inline-block",background:statusColors[selected.status]?.bg,color:statusColors[selected.status]?.color}}>{selected.status}</span>
                </div>
              </div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                <button style={{...s.btn,background:`${HW.red}15`,color:HW.red,fontSize:12,padding:"8px 12px"}} onClick={()=>exportPDF()}>📄 PDF</button>
                {selected.status!=="active"&&<button style={{...s.btn,background:`${HW.red}15`,color:HW.red,fontSize:12,padding:"8px 12px"}} onClick={()=>changeStatus("active")}>{t.reactivate}</button>}
                {selected.status!=="transferred"&&<button style={{...s.btn,background:"rgba(255,165,0,.15)",color:"#FFA500",fontSize:12,padding:"8px 12px"}} onClick={()=>changeStatus("transferred")}>{t.transfer}</button>}
                {selected.status!=="inactive"&&<button style={{...s.btn,background:"rgba(136,136,136,.15)",color:HW.muted,fontSize:12,padding:"8px 12px"}} onClick={()=>changeStatus("inactive")}>{t.inactive}</button>}
                {selected.status!=="dropped"&&<button style={{...s.btn,background:"rgba(100,100,100,.2)",color:"#666",fontSize:12,padding:"8px 12px"}} onClick={()=>changeStatus("dropped")}>{t.drop}</button>}
              </div>
            </div>

            <div style={{display:"flex",gap:0,background:HW.surface2,borderRadius:12,padding:4,marginBottom:14,overflowX:"auto"}}>
              {["timeline","edit","reports","penalties","sick","payment","laptop","goals"].map(tab=>(
                <button key={tab} onClick={()=>setProfileTab(tab)} style={{...s.btn,flex:1,padding:"10px 4px",background:profileTab===tab?HW.surface:"none",color:profileTab===tab?HW.text:HW.muted,fontSize:9,borderRadius:8,whiteSpace:"nowrap",borderBottom:profileTab===tab?`2px solid ${HW.red}`:"none"}}>
                  {tab==="timeline"?"📅":tab==="edit"?"✏️":tab==="reports"?"📋":tab==="penalties"?"⚠️":tab==="sick"?"🏥":tab==="payment"?"💰":tab==="laptop"?"💻":"🎯"}
                  {" "}{tab==="timeline"?t.timeline:tab==="edit"?t.edit:tab==="reports"?t.reports:tab==="penalties"?t.penalties:tab==="sick"?t.sick:tab==="payment"?t.paymentTab:tab==="laptop"?t.laptopTab:t.goalsTab}
                </button>
              ))}
            </div>

            {profileTab==="timeline"&&(<div style={s.card}><h3 style={{marginBottom:16,fontSize:16,color:HW.text}}>{t.activityTimeline}</h3>{logs.length===0?<p style={{color:HW.muted}}>{t.noActivity}</p>:logs.map((log,i)=>(<div key={log.id} style={{display:"flex",gap:12,marginBottom:16,position:"relative"}}>{i<logs.length-1&&<div style={{position:"absolute",left:15,top:32,width:2,height:"calc(100% + 4px)",background:HW.border}}/>}<div style={{width:32,height:32,borderRadius:"50%",background:HW.surface2,border:`2px solid ${HW.border}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,flexShrink:0,zIndex:1}}>{eventIcons[log.event_type]||"📌"}</div><div style={{flex:1,paddingTop:4}}><div style={{fontWeight:600,fontSize:13,color:HW.text}}>{log.description}</div><div style={{fontSize:11,color:HW.muted,marginTop:3}}>{new Date(log.created_at).toLocaleDateString(isAr?"ar-SA":"en-GB",{day:"numeric",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"})}</div></div></div>))}</div>)}

            {profileTab==="edit"&&(
              <div style={s.card}>
                <h3 style={{marginBottom:14,fontSize:16,color:HW.text}}>{t.editProfile}</h3>
                <label style={s.label}>{t.university}</label><input style={{...s.input,marginBottom:12}} value={selected.university||""} onChange={e=>setSelected({...selected,university:e.target.value})}/>
                <label style={s.label}>{t.department}</label><input style={{...s.input,marginBottom:12}} value={selected.department||""} onChange={e=>setSelected({...selected,department:e.target.value,_original:selected._original||{...selected}})}/>
                <label style={s.label}>{t.mentor}</label><input style={{...s.input,marginBottom:12}} value={selected.assigned_mentor||""} onChange={e=>setSelected({...selected,assigned_mentor:e.target.value,_original:selected._original||{...selected}})}/>
                <label style={s.label}>{t.gpa}</label><input style={{...s.input,marginBottom:12}} type="number" step="0.01" min="0" max="4" value={selected.gpa||""} onChange={e=>setSelected({...selected,gpa:e.target.value})}/>
                <label style={s.label}>{t.joiningDate}</label><input style={{...s.input,marginBottom:12}} type="date" value={selected.joining_date||""} onChange={e=>setSelected({...selected,joining_date:e.target.value})}/>
                <label style={s.label}>{t.ojtEndDate}</label><input style={{...s.input,marginBottom:14}} type="date" value={selected.ojt_end_date||""} onChange={e=>setSelected({...selected,ojt_end_date:e.target.value})}/>
                {selected.status==="dropped"&&(<><label style={s.label}>{t.quittingDate}</label><input style={{...s.input,marginBottom:12}} type="date" value={selected.quitting_date||""} onChange={e=>setSelected({...selected,quitting_date:e.target.value})}/><label style={s.label}>{t.quittingReason}</label><textarea style={{...s.input,height:70,resize:"vertical",marginBottom:14}} value={selected.quitting_reason||""} onChange={e=>setSelected({...selected,quitting_reason:e.target.value})}/></>)}
                <button style={{...s.btn,background:HW.red,color:HW.white,width:"100%",padding:14}} onClick={saveProfile}>{t.saveChanges}</button>
                {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,fontSize:13,marginTop:8,textAlign:"center"}}>{msg}</p>}
              </div>
            )}

            {profileTab==="reports"&&(<div>{reports.length===0?<p style={{color:HW.muted,padding:16}}>{isAr?"لا توجد تقارير بعد.":"No reports yet."}</p>:reports.map(r=>{let pie=null;try{pie=r.pie_chart_json?JSON.parse(r.pie_chart_json):null;}catch(e){}const isWeekly=!!r.weekly_tasks;return(<div key={r.id} style={{...s.card,borderLeft:`4px solid ${isWeekly?"#7c5cfc":HW.red}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}><div><span style={{padding:"2px 8px",borderRadius:8,fontSize:10,fontWeight:700,marginRight:6,background:isWeekly?"rgba(124,92,252,.15)":`${HW.red}15`,color:isWeekly?"#7c5cfc":HW.red}}>{isWeekly?"📅 Weekly":"📋 Daily"}</span><span style={{fontSize:13,fontWeight:600,color:HW.text}}>{isWeekly?`${r.week_start} → ${r.week_end||""}`:r.report_date}</span></div>{r.kpi_score&&<div style={{fontSize:20,fontWeight:800,color:kpiColor(r.kpi_score)}}>{r.kpi_score}<span style={{fontSize:10,color:HW.muted,fontWeight:400}}> KPI</span></div>}</div>{!isWeekly&&(<div style={{display:"flex",gap:12,marginBottom:10,background:HW.surface,borderRadius:8,padding:"8px 12px"}}><div style={{textAlign:"center",flex:1}}><div style={{fontSize:10,color:HW.muted,fontWeight:700,marginBottom:2}}>{isAr?"دخول":"TIME IN"}</div><div style={{fontSize:14,fontWeight:800,color:"#34d399",fontFamily:"monospace"}}>{r.signin_time||"—"}</div></div><div style={{width:1,background:HW.border}}/><div style={{textAlign:"center",flex:1}}><div style={{fontSize:10,color:HW.muted,fontWeight:700,marginBottom:2}}>{isAr?"خروج":"TIME OUT"}</div><div style={{fontSize:14,fontWeight:800,color:"#4f8ef7",fontFamily:"monospace"}}>{r.signout_time||"—"}</div></div><div style={{width:1,background:HW.border}}/><div style={{textAlign:"center",flex:1}}><div style={{fontSize:10,color:HW.muted,fontWeight:700,marginBottom:2}}>{isAr?"الحالة":"STATUS"}</div><div style={{fontSize:12,fontWeight:700,color:r.attended?"#34d399":HW.red}}>{r.attended?"● Present":"○ Absent"}</div></div></div>)}{isWeekly&&r.weekly_tasks&&<div style={{fontSize:13,marginBottom:10,borderLeft:"3px solid #7c5cfc",paddingLeft:10,lineHeight:1.6,color:HW.muted}}>{r.weekly_tasks.substring(0,200)}{r.weekly_tasks.length>200?"…":""}</div>}{r.penalty_applied&&<div style={{fontSize:12,color:HW.red,fontWeight:700,marginBottom:6}}>⚠️ Penalty: -{r.penalty_amount}%</div>}{pie&&<PieChart data={pie}/>}{r.talent_notes&&<div style={{background:`${HW.red}10`,borderRadius:8,padding:10,marginTop:10,borderLeft:`3px solid ${HW.red}`}}><div style={{fontSize:10,color:HW.red,fontWeight:700,marginBottom:4}}>🌟 AI TALENT NOTES</div><div style={{fontSize:12,lineHeight:1.6,color:HW.text}}>{r.talent_notes}</div></div>}</div>);})}</div>)}

            {profileTab==="penalties"&&(<div style={s.card}><h3 style={{marginBottom:12,fontSize:16,color:HW.text}}>⚠️ {t.penalties}</h3>{penalties.length===0?<p style={{color:HW.muted}}>{isAr?"لا توجد خصومات.":"No penalties recorded."}</p>:(<><div style={{background:`${HW.red}15`,border:`1px solid ${HW.red}30`,borderRadius:10,padding:12,marginBottom:12}}><div style={{fontSize:13,color:HW.red,fontWeight:700}}>{isAr?"الإجمالي:":"Total:"} {penalties.length} × {PENALTY_PCT}% = {(penalties.length*PENALTY_PCT).toFixed(2)}%</div></div>{penalties.map(p=>(<div key={p.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:HW.surface2,borderRadius:10,borderLeft:`3px solid ${HW.red}`,padding:12,marginBottom:8}}><div><div style={{fontSize:13,fontWeight:600,color:HW.text}}>📅 {p.report_date}</div><div style={{fontSize:12,color:HW.muted,marginTop:2}}>{p.reason}</div></div><div style={{fontSize:16,fontWeight:800,color:HW.red}}>-{p.amount}%</div></div>))}</>)}</div>)}

            {profileTab==="sick"&&(<div style={s.card}><h3 style={{marginBottom:12,fontSize:16,color:HW.text}}>🏥 {isAr?"سجل الإجازات المرضية":"Sick Leave History"}</h3><div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:14}}>{[{label:isAr?"الطلبات":"Requests",value:selectedSickLeaves.length,color:"#4f8ef7"},{label:t.freeDays,value:selectedSickLeaves.reduce((a,sl)=>a+Math.min(sl.total_days,2),0),color:"#34d399"},{label:t.penaltyDays,value:selectedSickLeaves.reduce((a,sl)=>a+(sl.penalty_days||0),0),color:HW.red}].map((stat,i)=>(<div key={i} style={{background:HW.surface2,borderRadius:10,padding:12,textAlign:"center",borderTop:`3px solid ${stat.color}`}}><div style={{fontSize:20,fontWeight:800,color:stat.color}}>{stat.value}</div><div style={{fontSize:10,color:HW.muted,marginTop:2}}>{stat.label}</div></div>))}</div>{selectedSickLeaves.length===0?<p style={{color:HW.muted}}>{t.noSickLeaves}</p>:selectedSickLeaves.map(sl=>(<div key={sl.id} style={{background:HW.surface2,borderRadius:10,padding:12,marginBottom:8,borderLeft:`3px solid ${sl.penalty_applied?HW.red:"#34d399"}`}}><div style={{display:"flex",justifyContent:"space-between"}}><div><div style={{fontSize:13,fontWeight:700,color:HW.text}}>📅 {sl.start_date} → {sl.end_date}</div><div style={{fontSize:12,color:HW.muted,marginTop:2}}>{sl.reason}</div>{sl.proof_url&&<div style={{fontSize:11,color:"#4f8ef7",marginTop:4}}>📎 {isAr?"إثبات طبي":"Medical proof"}</div>}</div><div style={{textAlign:"right"}}><div style={{fontSize:13,fontWeight:700,color:sl.penalty_applied?HW.red:"#34d399"}}>{sl.total_days} {isAr?"أيام":"day(s)"}</div><div style={{fontSize:11,color:HW.muted}}>{sl.penalty_applied?`⚠️ -${(sl.penalty_days*PENALTY_PCT).toFixed(2)}%`:t.nopenalty}</div></div></div></div>))}</div>)}

            {profileTab==="payment"&&(<div style={s.card}><h3 style={{marginBottom:14,fontSize:16,color:HW.text}}>{t.paymentStatus}</h3><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>{["paid","unpaid"].map(status=>(<button key={status} style={{...s.btn,padding:14,fontSize:15,fontWeight:800,background:selected.payment_status===status?status==="paid"?"rgba(52,211,153,.2)":`${HW.red}20`:HW.surface2,color:selected.payment_status===status?status==="paid"?"#34d399":HW.red:HW.muted,border:selected.payment_status===status?`2px solid ${status==="paid"?"#34d399":HW.red}`:`1px solid ${HW.border}`}} onClick={()=>setSelected({...selected,payment_status:status})}>{status==="paid"?t.paid:t.unpaid}</button>))}</div><label style={s.label}>{t.paymentNotes}</label><textarea style={{...s.input,height:80,resize:"vertical",marginBottom:12}} placeholder={isAr?"مثال: تم الدفع لشهر يونيو 2026":"e.g. Paid for June 2026"} value={selected.payment_notes||""} onChange={e=>setSelected({...selected,payment_notes:e.target.value})}/><button style={{...s.btn,background:HW.red,color:HW.white,width:"100%",padding:14}} onClick={()=>updatePaymentStatus(selected.payment_status,selected.payment_notes)}>{t.savePayment}</button>{msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,fontSize:13,marginTop:8,textAlign:"center"}}>{msg}</p>}</div>)}

            {profileTab==="laptop"&&(<div style={s.card}><h3 style={{marginBottom:14,fontSize:16,color:HW.text}}>{t.laptopStatus}</h3><div style={{background:selected.laptop_returned?"rgba(255,165,0,.1)":selected.laptop_received?"rgba(52,211,153,.1)":`${HW.red}10`,border:`1px solid ${selected.laptop_returned?"rgba(255,165,0,.3)":selected.laptop_received?"rgba(52,211,153,.3)":`${HW.red}30`}`,borderRadius:10,padding:14,marginBottom:14,textAlign:"center"}}><div style={{fontSize:28,marginBottom:4}}>{selected.laptop_returned?"🔄":selected.laptop_received?"✅":"❌"}</div><div style={{fontSize:15,fontWeight:800,color:selected.laptop_returned?"#FFA500":selected.laptop_received?"#34d399":HW.red}}>{selected.laptop_returned?(isAr?"مُرجع للموارد البشرية":"Returned to HR"):selected.laptop_received?(isAr?"تم الاستلام":"Laptop Received"):(isAr?"لم يُستلم":"Not Received")}</div></div><div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:14}}>{!selected.laptop_received&&<button style={{...s.btn,padding:14,fontSize:14,fontWeight:800,background:"rgba(52,211,153,.15)",color:"#34d399",border:"2px solid rgba(52,211,153,.4)"}} onClick={()=>setSelected({...selected,laptop_received:true,laptop_returned:false,laptop_received_date:selected.laptop_received_date||new Date().toISOString().split("T")[0]})}>{t.markReceived}</button>}{selected.laptop_received&&!selected.laptop_returned&&<button style={{...s.btn,padding:14,fontSize:14,fontWeight:800,background:"rgba(255,165,0,.15)",color:"#FFA500",border:"2px solid rgba(255,165,0,.4)"}} onClick={()=>setSelected({...selected,laptop_returned:true,laptop_returned_date:new Date().toISOString().split("T")[0]})}>{t.markReturned}</button>}{(selected.laptop_received||selected.laptop_returned)&&<button style={{...s.btn,padding:12,fontSize:13,background:`${HW.red}15`,color:HW.red,border:`1px solid ${HW.red}40`}} onClick={()=>setSelected({...selected,laptop_received:false,laptop_returned:false,laptop_serial:"",laptop_received_date:null,laptop_returned_date:null})}>{t.reset}</button>}</div>{selected.laptop_received&&(<><label style={s.label}>{t.laptopSerial}</label><input style={{...s.input,marginBottom:12}} placeholder="e.g. HW-2026-001" value={selected.laptop_serial||""} onChange={e=>setSelected({...selected,laptop_serial:e.target.value})}/><label style={s.label}>{t.dateReceived}</label><input style={{...s.input,marginBottom:12}} type="date" value={selected.laptop_received_date||""} onChange={e=>setSelected({...selected,laptop_received_date:e.target.value})}/></>)}{selected.laptop_returned&&(<><label style={s.label}>{t.dateReturned}</label><input style={{...s.input,marginBottom:12}} type="date" value={selected.laptop_returned_date||""} onChange={e=>setSelected({...selected,laptop_returned_date:e.target.value})}/></>)}<button style={{...s.btn,background:HW.red,color:HW.white,width:"100%",padding:14}} onClick={()=>updateLaptopStatus(selected.laptop_received,selected.laptop_serial,selected.laptop_received_date,selected.laptop_returned,selected.laptop_returned_date)}>{t.saveLaptop}</button>{msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,fontSize:13,marginTop:8,textAlign:"center"}}>{msg}</p>}</div>)}

            {profileTab==="goals"&&(<div style={s.card}><h3 style={{marginBottom:12,fontSize:16,color:HW.text}}>🎯 {t.goalsTab}</h3><div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:14}}>{KRA_CATEGORIES.map(kra=>{const kraGoals=selectedGoals.filter(g=>g.kra===kra.id);const completed=kraGoals.filter(g=>g.status==="completed").length;return(<div key={kra.id} style={{background:HW.surface2,borderRadius:10,padding:10,borderTop:`3px solid ${kra.color}`,textAlign:"center"}}><div style={{fontSize:16,marginBottom:2}}>{kra.icon}</div><div style={{fontSize:9,color:kra.color,fontWeight:700,textTransform:"uppercase",marginBottom:2}}>{(isAr?kra.labelAr:kra.label).split(" ")[0]}</div><div style={{fontSize:13,fontWeight:700,color:HW.text}}>{kraGoals.length}</div><div style={{fontSize:10,color:"#34d399"}}>{completed} {isAr?"منجز":"done"}</div></div>);})}</div>{selectedGoals.length===0?<p style={{color:HW.muted}}>{t.noGoalsSet}</p>:selectedGoals.map(goal=><GoalCard key={goal.id} goal={goal} isTrainee={false} lang={lang}/>)}</div>)}
            {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,fontSize:13,marginTop:8,textAlign:"center"}}>{msg}</p>}
          </div>
        )}

        {mgmtTab==="live"&&(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginBottom:16}}>{[{label:t.signedIn,value:liveSignins.filter(s=>s.attended).length,icon:"✅",color:"#34d399"},{label:t.absent,value:liveSignins.filter(s=>!s.attended).length,icon:"❌",color:HW.red},{label:t.onTimeLabel,value:liveSignins.filter(s=>s.signin_time&&s.signin_time<=MAX_SIGNIN).length,icon:"⏰",color:"#4f8ef7"},{label:t.late,value:liveSignins.filter(s=>s.signin_time&&s.signin_time>MAX_SIGNIN).length,icon:"⚠️",color:"#FFA500"}].map((stat,i)=>(<div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,borderRadius:12,padding:14,borderTop:`3px solid ${stat.color}`}}><div style={{fontSize:20,marginBottom:6}}>{stat.icon}</div><div style={{fontSize:24,fontWeight:800,color:stat.color}}>{stat.value}</div><div style={{fontSize:11,color:HW.muted,marginTop:2}}>{stat.label}</div></div>))}</div>
            <div style={s.card}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><h3 style={{margin:0,fontSize:16,color:HW.text}}>{t.liveFeed}</h3><div style={{display:"flex",alignItems:"center",gap:6,background:"rgba(52,211,153,.1)",border:"1px solid rgba(52,211,153,.3)",borderRadius:20,padding:"4px 10px"}}><div style={{width:7,height:7,borderRadius:"50%",background:"#34d399",animation:"pulse 1.5s infinite"}}/><span style={{fontSize:11,color:"#34d399",fontWeight:700}}>LIVE</span></div></div><style>{`@keyframes pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.5;transform:scale(1.3)}}@keyframes slideIn{from{transform:translateX(-20px);opacity:0}to{transform:translateX(0);opacity:1}}`}</style>{liveSignins.length===0?(<div style={{textAlign:"center",padding:32}}><div style={{fontSize:36,marginBottom:10}}>📡</div><p style={{color:HW.muted}}>{t.waitingSignins}</p></div>):(liveSignins.map((signin,i)=>(<div key={signin.id} style={{display:"flex",alignItems:"center",gap:12,padding:"12px 0",borderBottom:`1px solid ${HW.border}`,animation:i===0?"slideIn .4s ease":"none"}}><div style={{width:40,height:40,borderRadius:"50%",background:`linear-gradient(135deg,${HW.red},${HW.darkRed})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:800,color:HW.white,flexShrink:0}}>{signin.full_name?.split(" ").map(w=>w[0]).join("").slice(0,2)}</div><div style={{flex:1,minWidth:0}}><div style={{fontWeight:700,fontSize:14,color:HW.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{signin.full_name}</div><div style={{fontSize:11,color:HW.muted}}>{signin.department}</div></div><div style={{textAlign:"right",flexShrink:0}}><div style={{fontSize:16,fontWeight:800,fontFamily:"monospace",color:signin.signin_time>MAX_SIGNIN?HW.red:"#34d399"}}>{signin.signin_time||"—"}</div><span style={{fontSize:10,fontWeight:700,color:signin.attended?"#34d399":HW.red}}>{signin.attended?"● Present":"○ Absent"}</span></div></div>)))}</div>
          </div>
        )}

        {mgmtTab==="analytics"&&(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginBottom:16}}>{[{label:t.active,value:analytics.active.length,icon:"👥",color:HW.red},{label:"Avg KPI",value:analytics.avgKpi,icon:"📊",color:"#FFA500"},{label:isAr?"الحضور":"Attendance",value:`${analytics.attendanceRate}%`,icon:"✅",color:"#34d399"},{label:t.penalties,value:analytics.totalPenalties,icon:"⚠️",color:"#f87171"},{label:isAr?"إجازات مرضية":"Sick Leaves",value:allSickLeaves.length,icon:"🏥",color:"#4f8ef7"},{label:isAr?"مدفوع":"Paid",value:trainees.filter(tt=>tt.payment_status==="paid").length,icon:"💰",color:"#34d399"},{label:t.transferred,value:trainees.filter(tt=>tt.status==="transferred").length,icon:"🔄",color:"#FFA500"},{label:t.dropped,value:trainees.filter(tt=>tt.status==="dropped").length,icon:"🔴",color:"#666"}].map((stat,i)=>(<div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,borderRadius:12,padding:14,borderTop:`3px solid ${stat.color}`}}><div style={{fontSize:20,marginBottom:6}}>{stat.icon}</div><div style={{fontSize:24,fontWeight:800,color:stat.color}}>{stat.value}</div><div style={{fontSize:11,color:HW.muted,marginTop:2}}>{stat.label}</div></div>))}</div>
            <div style={s.card}><h3 style={{marginBottom:14,fontSize:16,color:HW.text}}>{t.kpiLeaderboard}</h3>{analytics.traineeKpi.length===0?<p style={{color:HW.muted}}>{t.noData}</p>:analytics.traineeKpi.slice(0,8).map((tt,i)=>(<div key={tt.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 0",borderBottom:`1px solid ${HW.border}`}}><div style={{width:28,height:28,borderRadius:"50%",background:i===0?"#FFD700":i===1?"#C0C0C0":i===2?"#CD7F32":HW.surface2,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:800,color:i<3?"#000":HW.muted,flexShrink:0}}>{i+1}</div><div style={{flex:1,minWidth:0}}><div style={{fontWeight:600,fontSize:14,color:HW.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{tt.full_name}</div><div style={{fontSize:11,color:HW.muted}}>{tt.department}</div></div><div style={{textAlign:"right",flexShrink:0}}><div style={{fontSize:18,fontWeight:800,color:kpiColor(parseFloat(tt.avgKpi))}}>{tt.avgKpi}</div><div style={{fontSize:10,color:HW.muted}}>KPI</div></div></div>))}</div>
            {analytics.atRisk.length>0&&(<div style={{...s.card,border:"1px solid rgba(248,113,113,.3)"}}><h3 style={{marginBottom:12,color:"#f87171",fontSize:16}}>{t.atRisk}</h3>{analytics.atRisk.map(tt=>(<div key={tt.id} style={{background:HW.surface2,borderRadius:10,padding:12,marginBottom:8,borderLeft:"3px solid #f87171"}}><div style={{fontWeight:700,fontSize:14,color:HW.text}}>{tt.full_name}</div><div style={{fontSize:12,color:HW.muted,marginBottom:6}}>{tt.department}</div><div style={{display:"flex",gap:16}}><div style={{textAlign:"center"}}><div style={{fontSize:18,fontWeight:800,color:"#f87171"}}>{tt.avgKpi}</div><div style={{fontSize:10,color:HW.muted}}>{t.avgKpi}</div></div><div style={{textAlign:"center"}}><div style={{fontSize:18,fontWeight:800,color:"#f87171"}}>{tt.penalties}</div><div style={{fontSize:10,color:HW.muted}}>{t.penalties}</div></div></div></div>))}</div>)}
          </div>
        )}

        {mgmtTab==="okr"&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><h3 style={{margin:0,fontSize:16,color:HW.text}}>{t.ocrTracking}</h3><button style={{...s.btn,background:HW.red,color:HW.white,padding:"8px 14px",fontSize:13}} onClick={()=>setShowAddOkr(!showAddOkr)}>{showAddOkr?t.cancel:t.addOkr}</button></div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginBottom:14}}>{[{label:isAr?"الإجمالي":"Total",value:okrs.length,color:HW.red},{label:t.onTrack,value:okrs.filter(o=>(o.current/o.target)>=.8).length,color:"#34d399"},{label:isAr?"قيد التنفيذ":"In Progress",value:okrs.filter(o=>(o.current/o.target)>=.5&&(o.current/o.target)<.8).length,color:"#FFA500"},{label:t.behind,value:okrs.filter(o=>(o.current/o.target)<.5).length,color:"#f87171"}].map((s2,i)=>(<div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,borderRadius:12,padding:14,textAlign:"center",borderTop:`3px solid ${s2.color}`}}><div style={{fontSize:24,fontWeight:800,color:s2.color}}>{s2.value}</div><div style={{fontSize:11,color:HW.muted,marginTop:2}}>{s2.label}</div></div>))}</div>
            {showAddOkr&&(<div style={{...s.card,border:`1px solid ${HW.red}40`,marginBottom:14}}><h4 style={{marginBottom:14,color:HW.red,fontSize:15}}>{t.addOkrTitle}</h4><label style={s.label}>{t.department}</label><input style={{...s.input,marginBottom:12}} placeholder={isAr?"مثال: الهندسة":"e.g. Engineering"} value={newOkr.department} onChange={e=>setNewOkr({...newOkr,department:e.target.value})}/><label style={s.label}>{t.objective}</label><input style={{...s.input,marginBottom:12}} placeholder={isAr?"مثال: تحسين المهارات":"e.g. Improve Skills"} value={newOkr.objective} onChange={e=>setNewOkr({...newOkr,objective:e.target.value})}/><label style={s.label}>{t.keyResult}</label><input style={{...s.input,marginBottom:12}} placeholder={isAr?"مثال: إكمال 20 جلسة":"e.g. Complete 20 sessions"} value={newOkr.key_result} onChange={e=>setNewOkr({...newOkr,key_result:e.target.value})}/><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}><div><label style={s.label}>{t.target}</label><input style={s.input} type="number" value={newOkr.target} onChange={e=>setNewOkr({...newOkr,target:parseFloat(e.target.value)})}/></div><div><label style={s.label}>{t.unit}</label><select style={s.input} value={newOkr.unit} onChange={e=>setNewOkr({...newOkr,unit:e.target.value})}><option value="%">%</option><option value="sessions">{isAr?"جلسات":"Sessions"}</option><option value="reports">{isAr?"تقارير":"Reports"}</option><option value="tasks">{isAr?"مهام":"Tasks"}</option></select></div></div><label style={s.label}>{t.dueDate}</label><input style={{...s.input,marginBottom:14}} type="date" value={newOkr.due_date} onChange={e=>setNewOkr({...newOkr,due_date:e.target.value})}/><button style={{...s.btn,background:HW.red,color:HW.white,width:"100%",padding:14}} onClick={addOkr}>{t.addOkrTitle}</button>{msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,fontSize:13,marginTop:8,textAlign:"center"}}>{msg}</p>}</div>)}
            {[...new Set(okrs.map(o=>o.department))].map(dept=>(<div key={dept} style={s.card}><h4 style={{marginBottom:12,color:HW.red,fontSize:15}}>🏢 {dept}</h4>{okrs.filter(o=>o.department===dept).map(okr=><OKRBar key={okr.id} okr={okr} onUpdate={updateOkr}/>)}</div>))}
            {okrs.length===0&&<div style={{...s.card,textAlign:"center",padding:32}}><p style={{color:HW.muted}}>{isAr?"لا توجد OKRs بعد.":"No OKRs yet."}</p></div>}
          </div>
        )}

        {mgmtTab==="access"&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
              <div><h3 style={{margin:0,fontSize:16,color:HW.text}}>{t.accessLog}</h3><p style={{fontSize:12,color:HW.muted,margin:"4px 0 0"}}>{t.loggedInAs} <b style={{color:HW.red}}>{currentManagerName||"—"}</b></p></div>
              <button style={{...s.btn,background:`${HW.red}20`,color:HW.red,fontSize:12,padding:"8px 12px"}} onClick={fetchAccessLogs}>{t.refresh}</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginBottom:16}}>
              {[{label:t.totalActions,value:accessLogs.length,color:HW.red,icon:"📊"},{label:t.logins,value:accessLogs.filter(l=>l.action_type==="login").length,color:"#34d399",icon:"🔑"},{label:t.profileViews,value:accessLogs.filter(l=>l.action_type==="profile_view").length,color:"#4f8ef7",icon:"👁"},{label:t.profileEdits,value:accessLogs.filter(l=>l.action_type==="profile_edit").length,color:"#7c5cfc",icon:"✏️"},{label:t.statusChanges,value:accessLogs.filter(l=>l.action_type==="status_change").length,color:"#FFA500",icon:"🔄"},{label:t.paymentUpdates,value:accessLogs.filter(l=>l.action_type==="payment_update").length,color:"#34d399",icon:"💰"},{label:t.laptopUpdates,value:accessLogs.filter(l=>l.action_type==="laptop_update").length,color:"#0891B2",icon:"💻"},{label:t.exports,value:accessLogs.filter(l=>["excel_export","pdf_export"].includes(l.action_type)).length,color:"#f87171",icon:"📤"}].map((stat,i)=>(<div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,borderRadius:12,padding:14,borderTop:`3px solid ${stat.color}`}}><div style={{fontSize:20,marginBottom:6}}>{stat.icon}</div><div style={{fontSize:22,fontWeight:800,color:stat.color}}>{stat.value}</div><div style={{fontSize:11,color:HW.muted,marginTop:2}}>{stat.label}</div></div>))}
            </div>
            <div style={{...s.card,marginBottom:14}}>
              <h4 style={{marginBottom:12,fontSize:15,color:HW.text}}>{t.accessByPerson}</h4>
              {[...new Set(accessLogs.filter(l=>l.manager_name&&!["Unknown","Manager"].includes(l.manager_name)).map(l=>l.manager_name))].map(name=>{
                const userLogs=accessLogs.filter(l=>l.manager_name===name);
                const lastLogin=userLogs.find(l=>l.action_type==="login");
                return(<div key={name} style={{display:"flex",alignItems:"center",gap:12,padding:"12px 0",borderBottom:`1px solid ${HW.border}`}}><div style={{width:44,height:44,borderRadius:"50%",background:`linear-gradient(135deg,${HW.red},${HW.darkRed})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,fontWeight:800,color:HW.white,flexShrink:0}}>{name[0]}</div><div style={{flex:1}}><div style={{fontWeight:700,fontSize:15,color:HW.text}}>{name}</div><div style={{fontSize:11,color:HW.muted,marginTop:2}}>{userLogs.length} {isAr?"إجراء":"actions"} · {userLogs.filter(l=>l.action_type==="login").length} {isAr?"دخول":"logins"}{lastLogin&&` · ${isAr?"آخر دخول:":"Last:"} ${lastLogin.metadata?.day||""} ${lastLogin.metadata?.date||""} ${isAr?"الساعة":"at"} ${lastLogin.metadata?.time||""}`}</div></div></div>);
              })}
              {accessLogs.filter(l=>l.manager_name&&!["Unknown","Manager"].includes(l.manager_name)).length===0&&(<div style={{textAlign:"center",padding:20}}><p style={{color:HW.muted,fontSize:13}}>{t.noAccessYet}</p><p style={{color:HW.muted,fontSize:12}}>{t.selectNameToTrack}</p></div>)}
            </div>
            <div style={s.card}>
              <h4 style={{marginBottom:12,fontSize:15,color:HW.text}}>{t.fullActivityLog} ({accessLogs.length})</h4>
              {accessLogs.length===0?(<div style={{textAlign:"center",padding:24}}><div style={{fontSize:32,marginBottom:8}}>🔐</div><p style={{color:HW.muted}}>{t.noActivityYet}</p></div>):(
                accessLogs.map(log=>{
                  const acMap={login:{color:"#34d399",icon:"🔑"},logout:{color:HW.muted,icon:"🚪"},profile_view:{color:"#4f8ef7",icon:"👁"},profile_edit:{color:"#7c5cfc",icon:"✏️"},status_change:{color:"#FFA500",icon:"🔄"},payment_update:{color:"#34d399",icon:"💰"},laptop_update:{color:"#0891B2",icon:"💻"},okr_add:{color:"#34d399",icon:"🎯"},excel_export:{color:"#f87171",icon:"📊"},pdf_export:{color:HW.red,icon:"📄"}};
                  const ac=acMap[log.action_type]||{color:HW.muted,icon:"📌"};
                  return(<div key={log.id} style={{display:"flex",gap:12,padding:"12px 0",borderBottom:`1px solid ${HW.border}`}}><div style={{width:36,height:36,borderRadius:"50%",background:`${ac.color}20`,border:`2px solid ${ac.color}40`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0}}>{ac.icon}</div><div style={{flex:1,minWidth:0}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div style={{fontWeight:700,fontSize:13,color:HW.text}}>{log.manager_name||"Unknown"}</div><div style={{fontSize:10,color:HW.muted,flexShrink:0,marginLeft:8,textAlign:"right"}}><div style={{fontWeight:600}}>{log.metadata?.day||""}</div><div>{log.metadata?.date||new Date(log.created_at).toLocaleDateString(isAr?"ar-SA":"en-GB",{day:"numeric",month:"short",year:"numeric"})}</div><div style={{color:ac.color,fontWeight:700}}>{log.metadata?.time||new Date(log.created_at).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}</div></div></div><div style={{fontSize:12,color:HW.muted,marginTop:2}}>{log.description}</div><span style={{fontSize:10,fontWeight:700,padding:"2px 6px",borderRadius:6,background:`${ac.color}15`,color:ac.color,marginTop:4,display:"inline-block"}}>{log.action_type?.replace(/_/g," ")}</span></div></div>);
                })
              )}
            </div>
          </div>
        )}
      </div>
    );
  }
}