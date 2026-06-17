/* eslint-disable */
import { useEffect, useState } from "react";
import { supabase } from "./lib/supabase";
import * as XLSX from "xlsx";
import {
  Chart as ChartJS,
  ArcElement,
  BarElement,
  LineElement,
  PointElement,
  CategoryScale,
  LinearScale,
  Tooltip,
  Legend,
} from "chart.js";
import { Doughnut, Bar, Line } from "react-chartjs-2";

ChartJS.register(ArcElement, BarElement, LineElement, PointElement, CategoryScale, LinearScale, Tooltip, Legend);

const CATS_EXP = ["식비","카페","교통","쇼핑","의료","문화/여가","구독","통신","저축","대출상환","기타"];
const CATS_INC = ["이월","월급","용돈","부업","이자","환급","기타수입"];

function won(n) {
  const currency = localStorage.getItem("currency") || "KRW";
  const option = CURRENCY_OPTIONS.find(c => c.code === currency) || CURRENCY_OPTIONS[0];

  return new Intl.NumberFormat(option.locale, {
    style: "currency",
    currency,
    maximumFractionDigits: currency === "KRW" || currency === "JPY" ? 0 : 2
  }).format(Number(n || 0));
}

function comma(n) {
  return String(n || "").replace(/\D/g, "").replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function uncomma(n) {
  return Number(String(n || "0").replaceAll(",", ""));
}

const CURRENCY_OPTIONS = [
  { code: "KRW", label: "KRW ₩", locale: "ko-KR" },
  { code: "USD", label: "USD $", locale: "en-US" },
  { code: "EUR", label: "EUR €", locale: "de-DE" },
  { code: "JPY", label: "JPY ¥", locale: "ja-JP" },
  { code: "GBP", label: "GBP £", locale: "en-GB" }
];

const DEFAULT_MONEY_TYPES = ["계좌", "카드", "저축", "투자"];


const I18N = {
  ko: {
    main: "메인",
    money: "자산관리",
    system: "시스템",
    home: "홈",
    tx: "내역",
    chart: "분석",
    budget: "예산",
    saving: "저축",
    assets: "자산",
    loan: "대출",
    subs: "구독",
    settings: "설정",
    logout: "로그아웃",

    login: "로그인",
    signup: "회원가입",
    email: "이메일",
    password: "비밀번호",
    loading: "로딩 중...",
    authSub: "내 돈을 귀엽게 관리해요 💗",
    confirm: "확인",

    language: "언어 설정",
    korean: "한국어",
    english: "English",
    currency: "통화 설정",

    remainingBudget: "이번 달 남은 예산",
    quickInput: "빠른 입력",
    editTx: "입출금 수정",
    amount: "금액",
    memo: "내용",
    category: "카테고리",
    moneyType: "출입금 유형",
    transferTo: "이체 도착 유형",
    repeat: "반복",
    noRepeat: "반복 없음",
    monthlyRepeat: "매월 반복",
    weeklyRepeat: "매주 반복",
    save: "저장",
    updateSave: "수정 저장",
    cancel: "취소",

    income: "수입",
    expense: "지출",
    balance: "잔액",
    transfer: "이체",
    transferExcluded: "합계 제외",

    expenseButton: "지출 💸",
    incomeButton: "수입 💰",

    excel: "엑셀/CSV 가져오기",
    excelHint: "날짜, 내용, 금액이 있는 엑셀 파일을 넣으면 자동으로 내역에 추가돼요.",
    txList: "거래 내역",
    typeSummary: "유형별 출입금 요약",
    typeSummaryHint: "이체는 유형별 이동에는 표시되지만 전체 수입/지출에는 계산되지 않아요.",
    noTx: "아직 내역이 없어요",

    budgetUsage: "예산 사용률",
    totalBudget: "총 예산",
    used: "사용",
    carryOver: "자동 이월",
    nextBudget: "다음달 예산",
    budgetByCategory: "카테고리별 예산",
    monthlyBudget: "월 예산",
    saveBudget: "예산 저장",

    savingProgress: "저축 목표 달성률",
    totalTarget: "목표 합계",
    totalCurrent: "현재 합계",
    leftAmount: "남은 금액",
    addSaving: "저축 목표 추가",
    editSaving: "저축 목표 수정",
    savingStatus: "저축 현황",
    savingNamePlaceholder: "목표명 예: 비상금, 여행 적금",
    savingTargetPlaceholder: "목표 금액",
    savingCurrentPlaceholder: "현재 모은 금액",
    noSaving: "저축 목표가 없어요",
    expectedDate: "예상 달성일",
    unknownEstimate: "예상 불가",

    totalAccountBalance: "총 계좌 잔액",
    accountCount: "계좌 수",
    cardCount: "카드 수",
    accountBalanceManage: "계좌별 잔액 관리",
    accountNamePlaceholder: "계좌명 예: 국민은행, 카카오뱅크",
    currentBalance: "현재 잔액",
    addAccount: "계좌 추가",
    assetTrend: "월별 자산 추이",
    assetTrendHint: "계좌 잔액과 월별 수입/지출을 기준으로 계산한 간단 추이예요.",
    expectedAssets: "예상 자산",
    cardPaymentManage: "카드 결제일 관리",
    cardNamePlaceholder: "카드명 예: 신한카드",
    paymentDayPlaceholder: "결제일 예: 25",
    addCard: "카드 추가",
    monthlyDay: "매월",

    monthlySubscriptions: "월 구독료",
    subCount: "구독 수",
    subManage: "구독 관리",
    subNamePlaceholder: "구독명 예: 넷플릭스, 유튜브",
    subAmountPlaceholder: "월 구독료",
    subDayPlaceholder: "결제일 예: 15",
    addSub: "구독 추가",

    loanManage: "대출 관리",
    editLoan: "대출 수정",
    loanNamePlaceholder: "대출명",
    principalPlaceholder: "대출 원금",
    ratePlaceholder: "연이율 %",
    termPlaceholder: "기간 개월",
    loanStartDate: "대출 시작일",
    equalPayment: "원리금균등",
    equalPrincipal: "원금균등",
    addLoan: "대출 추가",
    loanPrincipal: "대출 원금",
    paidAmount: "누적 상환",
    txApplied: "입출금 반영",
    remainingLoan: "남은 잔금",
    monthlyPayment: "월 상환액",
    startDate: "시작일",

    settingsTitle: "설정",
    moneyTypeEdit: "출입금 유형 편집",
    newTypeAdd: "새 유형 추가",
    newTypePlaceholder: "예: 비상금, CMA, 현금",
    add: "추가",
    delete: "삭제",
    edit: "수정",

    incomeCategory: "수입 카테고리",
    expenseCategory: "지출 카테고리",
    noIncomeData: "이 달의 수입 데이터가 없어요",
    noExpenseData: "이 달의 지출 데이터가 없어요",
    incomeExpenseCompare: "수입 / 지출 비교",


    transferHint: "계좌/카드/저축/투자 간 이동은 수입·지출 합계에 포함되지 않아요.",
    repeatMonthly: "매월 반복",
    repeatWeekly: "매주 반복",
    txTransferPrefix: "이체",
    savingNamePlaceholder: "목표명 예: 비상금, 여행 적금",
    savingTargetPlaceholder: "목표 금액",
    savingCurrentPlaceholder: "현재 모은 금액",
    loanTypeEqualPayment: "원리금균등",
    loanTypeEqualPrincipal: "원금균등",
    annualRateShort: "연",
    monthsUnit: "개월",
    loanPaidSentence: "총 {total}개월 중 {paid}개월치 갚았어요.",
    loanRemainingSentence: "앞으로 {months}개월 남았어요.",
    loanRemainingYearSentence: "앞으로 {months}개월 · 약 {years}년 {leftMonths}개월 남았어요.",
    notEntered: "미입력",
    moneyTypeEditorTitle: "설정 · 출입금 유형 편집",
    addTypeRequired: "추가할 유형명을 입력해주세요",
    duplicateType: "이미 있는 유형이에요",
    minTypeRequired: "유형은 최소 1개 필요해요",
    moneyTypeHint: "계좌, 카드, 저축, 투자 외에 비상금/CMA/현금 같은 유형을 추가할 수 있어요.",
    chartMonthTitle: "{year}년 {month}월",
    chartDatasetMonth: "{year}년 {month}월",
    excelNoData: "엑셀에서 데이터를 찾지 못했어요.",
    emailPasswordRequired: "이메일과 비밀번호를 입력해주세요",
    accountNameRequired: "계좌명을 입력해주세요",
    cardRequired: "카드명과 결제일을 입력해주세요",
    subRequired: "구독명과 금액을 입력해주세요",
    amountRequired: "금액을 입력해주세요",
    budgetRequired: "예산 금액을 입력해주세요",
    savingRequired: "목표명과 목표 금액을 입력해주세요",
    loanRequired: "대출 정보를 모두 입력해주세요",

    account: "계좌",
    card: "카드",
    savingsType: "저축",
    investment: "투자",
    food: "식비",
    cafe: "카페",
    transport: "교통",
    shopping: "쇼핑",
    medical: "의료",
    leisure: "문화/여가",
    subscription: "구독",
    telecom: "통신",
    loanRepay: "대출상환",
    other: "기타",
    salary: "월급",
    allowance: "용돈",
    sideJob: "부업",
    interest: "이자",
    refund: "환급",
    carryForward: "이월",
    otherIncome: "기타수입"
  },
  en: {
    main: "Main",
    money: "Money",
    system: "System",
    home: "Home",
    tx: "Transactions",
    chart: "Analytics",
    budget: "Budget",
    saving: "Savings",
    assets: "Assets",
    loan: "Loans",
    subs: "Subscriptions",
    settings: "Settings",
    logout: "Logout",

    login: "Log in",
    signup: "Sign up",
    email: "Email",
    password: "Password",
    loading: "Loading...",
    authSub: "Manage your money beautifully 💗",
    confirm: "OK",

    language: "Language",
    korean: "Korean",
    english: "English",
    currency: "Currency",

    remainingBudget: "Remaining budget this month",
    quickInput: "Quick entry",
    editTx: "Edit transaction",
    amount: "Amount",
    memo: "Description",
    category: "Category",
    moneyType: "Money type",
    transferTo: "Transfer destination",
    repeat: "Repeat",
    noRepeat: "No repeat",
    monthlyRepeat: "Monthly",
    weeklyRepeat: "Weekly",
    save: "Save",
    updateSave: "Save changes",
    cancel: "Cancel",

    income: "Income",
    expense: "Expense",
    balance: "Balance",
    transfer: "Transfer",
    transferExcluded: "Excluded",

    expenseButton: "Expense 💸",
    incomeButton: "Income 💰",

    excel: "Import Excel/CSV",
    excelHint: "Upload a file with date, description, and amount to add transactions automatically.",
    txList: "Transactions",
    typeSummary: "Summary by money type",
    typeSummaryHint: "Transfers are shown by type but excluded from total income and expenses.",
    noTx: "No transactions yet",

    budgetUsage: "Budget usage",
    totalBudget: "Total budget",
    used: "Used",
    carryOver: "Carry-over",
    nextBudget: "Next month budget",
    budgetByCategory: "Budget by category",
    monthlyBudget: "Monthly budget",
    saveBudget: "Save budget",

    savingProgress: "Savings progress",
    totalTarget: "Total target",
    totalCurrent: "Current total",
    leftAmount: "Remaining",
    addSaving: "Add savings goal",
    editSaving: "Edit savings goal",
    savingStatus: "Savings status",
    savingNamePlaceholder: "Goal name e.g. emergency fund, trip",
    savingTargetPlaceholder: "Target amount",
    savingCurrentPlaceholder: "Current saved amount",
    noSaving: "No savings goals yet",
    expectedDate: "Estimated completion",
    unknownEstimate: "Not enough data",

    totalAccountBalance: "Total account balance",
    accountCount: "Accounts",
    cardCount: "Cards",
    accountBalanceManage: "Account balances",
    accountNamePlaceholder: "Account name e.g. Chase, Kakao Bank",
    currentBalance: "Current balance",
    addAccount: "Add account",
    assetTrend: "Monthly asset trend",
    assetTrendHint: "A simple trend based on account balances and monthly income/expenses.",
    expectedAssets: "Estimated assets",
    cardPaymentManage: "Card payment dates",
    cardNamePlaceholder: "Card name e.g. Chase Sapphire",
    paymentDayPlaceholder: "Payment day e.g. 25",
    addCard: "Add card",
    monthlyDay: "Monthly day",

    monthlySubscriptions: "Monthly subscriptions",
    subCount: "Subscriptions",
    subManage: "Subscription management",
    subNamePlaceholder: "Subscription e.g. Netflix, YouTube",
    subAmountPlaceholder: "Monthly fee",
    subDayPlaceholder: "Payment day e.g. 15",
    addSub: "Add subscription",

    loanManage: "Loan management",
    editLoan: "Edit loan",
    loanNamePlaceholder: "Loan name",
    principalPlaceholder: "Loan principal",
    ratePlaceholder: "Annual interest %",
    termPlaceholder: "Term in months",
    loanStartDate: "Loan start date",
    equalPayment: "Equal payment",
    equalPrincipal: "Equal principal",
    addLoan: "Add loan",
    loanPrincipal: "Loan principal",
    paidAmount: "Paid amount",
    txApplied: "Transactions applied",
    remainingLoan: "Remaining balance",
    monthlyPayment: "Monthly payment",
    startDate: "Start date",

    settingsTitle: "Settings",
    moneyTypeEdit: "Money type editor",
    newTypeAdd: "Add new type",
    newTypePlaceholder: "e.g. Emergency fund, CMA, Cash",
    add: "Add",
    delete: "Delete",
    edit: "Edit",

    incomeCategory: "Income category",
    expenseCategory: "Expense category",
    noIncomeData: "No income data this month",
    noExpenseData: "No expense data this month",
    incomeExpenseCompare: "Income / Expense comparison",


    transferHint: "Transfers between account/card/savings/investment are excluded from income and expense totals.",
    repeatMonthly: "Monthly recurring",
    repeatWeekly: "Weekly recurring",
    txTransferPrefix: "Transfer",
    savingNamePlaceholder: "Goal name e.g. emergency fund, trip",
    savingTargetPlaceholder: "Target amount",
    savingCurrentPlaceholder: "Current saved amount",
    loanTypeEqualPayment: "Equal payment",
    loanTypeEqualPrincipal: "Equal principal",
    annualRateShort: "APR",
    monthsUnit: "months",
    loanPaidSentence: "{paid} of {total} months paid.",
    loanRemainingSentence: "{months} months remaining.",
    loanRemainingYearSentence: "{months} months remaining · about {years}y {leftMonths}m.",
    notEntered: "Not entered",
    moneyTypeEditorTitle: "Settings · Money type editor",
    addTypeRequired: "Please enter a type name.",
    duplicateType: "This type already exists.",
    minTypeRequired: "At least one type is required.",
    moneyTypeHint: "You can add types like emergency fund, CMA, or cash.",
    chartMonthTitle: "{year}-{month}",
    chartDatasetMonth: "{year}-{month}",
    excelNoData: "No data found in the Excel file.",
    emailPasswordRequired: "Please enter email and password.",
    accountNameRequired: "Please enter an account name.",
    cardRequired: "Please enter card name and payment day.",
    subRequired: "Please enter subscription name and amount.",
    amountRequired: "Please enter an amount.",
    budgetRequired: "Please enter a budget amount.",
    savingRequired: "Please enter a goal name and target amount.",
    loanRequired: "Please enter all loan information.",

    account: "Account",
    card: "Card",
    savingsType: "Savings",
    investment: "Investment",
    food: "Food",
    cafe: "Cafe",
    transport: "Transport",
    shopping: "Shopping",
    medical: "Medical",
    leisure: "Leisure",
    subscription: "Subscription",
    telecom: "Telecom",
    loanRepay: "Loan repayment",
    other: "Other",
    salary: "Salary",
    allowance: "Allowance",
    sideJob: "Side job",
    interest: "Interest",
    refund: "Refund",
    carryForward: "Carry-forward",
    otherIncome: "Other income"
  }
};


function guessCategory(text, type) {
  const t = String(text || "");

  if (type === "income") return "기타수입";
  if (t.includes("스타벅스") || t.includes("커피") || t.includes("카페")) return "카페";
  if (t.includes("쿠팡") || t.includes("쇼핑") || t.includes("무신사")) return "쇼핑";
  if (t.includes("버스") || t.includes("지하철") || t.includes("택시") || t.includes("교통")) return "교통";
  if (t.includes("병원") || t.includes("약국")) return "의료";
  if (t.includes("넷플릭스") || t.includes("유튜브") || t.includes("구독")) return "구독";
  if (t.includes("통신") || t.includes("SKT") || t.includes("KT") || t.includes("LG")) return "통신";
  if (t.includes("마트") || t.includes("식당") || t.includes("배달") || t.includes("편의점")) return "식비";

  return "기타";
}

function normalizeAmount(value) {
  if (value === undefined || value === null) return 0;
  return Number(String(value).replaceAll(",", "").replaceAll("원", "").replaceAll(" ", ""));
}

function normalizeDate(value) {
  if (!value) return new Date().toISOString().slice(0, 10);

  if (typeof value === "number") {
    const d = XLSX.SSF.parse_date_code(value);
    if (d) {
      return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
    }
  }

  const text = String(value).replaceAll(".", "-").replaceAll("/", "-").trim();
  const date = new Date(text);
  if (!isNaN(date.getTime())) return date.toISOString().slice(0, 10);

  return new Date().toISOString().slice(0, 10);
}

export default function App() {
  const [user, setUser] = useState(null);
  const [loading, setLoading] = useState(true);

  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [popup, setPopup] = useState(null);
  const [language, setLanguage] = useState(localStorage.getItem("language") || "ko");
  const [currency, setCurrency] = useState(localStorage.getItem("currency") || "KRW");
  const tr = (key) => I18N[language]?.[key] || key;
  const ui = (ko, en) => (language === "en" ? en : ko);
  const fmt = (key, vars = {}) =>
    Object.entries(vars).reduce(
      (out, [name, value]) => out.replaceAll(`{${name}}`, value),
      tr(key)
    );

  const catLabel = (value) => {
    const map = {
      "식비": "food",
      "카페": "cafe",
      "교통": "transport",
      "쇼핑": "shopping",
      "의료": "medical",
      "문화/여가": "leisure",
      "구독": "subscription",
      "통신": "telecom",
      "저축": "savingsType",
      "대출상환": "loanRepay",
      "기타": "other",
      "이월": "carryForward",
      "월급": "salary",
      "용돈": "allowance",
      "부업": "sideJob",
      "이자": "interest",
      "환급": "refund",
      "기타수입": "otherIncome",
      "이체": "transfer"
    };
    return tr(map[value]) || value;
  };

  const moneyTypeLabel = (value) => {
    const map = {
      "계좌": "account",
      "카드": "card",
      "저축": "savingsType",
      "투자": "investment"
    };
    return tr(map[value]) || value;
  };

  const [tab, setTab] = useState("home");
  const [chartYear, setChartYear] = useState(new Date().getFullYear());
  const [chartMonth, setChartMonth] = useState(new Date().getMonth());

  const thisMonthKey = new Date().toISOString().slice(0, 7);
  const [txPeriodMode, setTxPeriodMode] = useState("month");
  const [txSelectedMonth, setTxSelectedMonth] = useState(thisMonthKey);
  const [txStartDate, setTxStartDate] = useState("");
  const [txEndDate, setTxEndDate] = useState("");
  const [txSort, setTxSort] = useState("latest");
  const [txs, setTxs] = useState([]);
  const [accounts, setAccounts] = useState(() => JSON.parse(localStorage.getItem("accounts") || "[]"));
  const [cards, setCards] = useState(() => JSON.parse(localStorage.getItem("cards") || "[]"));
  const [subs, setSubs] = useState(() => JSON.parse(localStorage.getItem("subs") || "[]"));

  const [accountName, setAccountName] = useState("");
  const [accountBalance, setAccountBalance] = useState("");

  const [cardName, setCardName] = useState("");
  const [cardPayDay, setCardPayDay] = useState("");

  const [subName, setSubName] = useState("");
  const [subAmount, setSubAmount] = useState("");
  const [subDay, setSubDay] = useState("");

  const [editingTxId, setEditingTxId] = useState(null);
  const [type, setType] = useState("expense");
  const [amount, setAmount] = useState("");
  const [name, setName] = useState("");
  const [category, setCategory] = useState("식비");
  const [moneyType, setMoneyType] = useState("계좌");
  const [transferTo, setTransferTo] = useState("카드");
  const [moneyTypes, setMoneyTypes] = useState(() => {
    const saved = localStorage.getItem("moneyTypes");
    return saved ? JSON.parse(saved) : DEFAULT_MONEY_TYPES;
  });
  const [newMoneyType, setNewMoneyType] = useState("");
  const [repeat, setRepeat] = useState("none");

  const [budgets, setBudgets] = useState([]);
  const [budgetCat, setBudgetCat] = useState("식비");
  const [budgetAmount, setBudgetAmount] = useState("");

  const [savings, setSavings] = useState([]);
  const [editingSavingId, setEditingSavingId] = useState(null);
  const [savingName, setSavingName] = useState("");
  const [savingTarget, setSavingTarget] = useState("");
  const [savingCurrent, setSavingCurrent] = useState("");

  const [loans, setLoans] = useState([]);
  const [editingLoanId, setEditingLoanId] = useState(null);
  const [loanName, setLoanName] = useState("");
  const [loanPrincipal, setLoanPrincipal] = useState("");
  const [loanRate, setLoanRate] = useState("");
  const [loanTerm, setLoanTerm] = useState("");
  const [loanStartDate, setLoanStartDate] = useState(new Date().toISOString().slice(0, 10));
  const [loanType, setLoanType] = useState("equal_payment");

  useEffect(() => {
    supabase.auth.getSession().then(async ({ data }) => {
      const currentUser = data.session?.user ?? null;
      setUser(currentUser);
      if (currentUser) await loadUserData(currentUser);
      setLoading(false);
    });

    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      setUser(session?.user ?? null);
    });

    return () => listener.subscription.unsubscribe();
  }, []);

  const login = async () => {
    if (!email || !password) return alert(tr("emailPasswordRequired"));
    const { data, error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) return alert(error.message);
    if (data.user) await loadUserData(data.user);
  };

  const signup = async () => {
    if (!email || !password) return alert(tr("emailPasswordRequired"));
    const { error } = await supabase.auth.signUp({ email, password });
    if (error) return alert(error.message);
    showPopup(ui("회원가입 완료! 이메일 확인이 필요할 수 있어요.", "Sign-up complete. Email verification may be required."), ui("회원가입 완료", "Sign-up complete"), "💌");
  };

  const showPopup = (message, title = "완료", emoji = "💗") => {
    setPopup({ title, message, emoji });
  };

  const closePopup = () => {
    setPopup(null);
  };

  const logout = async () => {
    await supabase.auth.signOut();
    setUser(null);
  };

  const loadUserData = async (currentUser) => {
    if (!currentUser) return;

    const [{ data: txData }, { data: budgetData }, { data: savingData }, { data: loanData }] =
      await Promise.all([
        supabase.from("transactions").select("*").eq("user_id", currentUser.id).order("date", { ascending: false }),
        supabase.from("budgets").select("*").eq("user_id", currentUser.id),
        supabase.from("savings").select("*").eq("user_id", currentUser.id),
        supabase.from("loans").select("*").eq("user_id", currentUser.id),
      ]);

    setTxs((txData || []).map(t => ({
      id: t.id,
      type: t.type,
      amount: Number(t.amount),
      name: t.name,
      category: t.category,
      date: t.date,
      repeat: t.repeat || "none",
      moneyType: t.account || "계좌",
      transferTo: "",
      memo: t.memo || "",
    })));

    setBudgets((budgetData || []).map(b => ({
      id: b.id,
      category: b.category,
      amount: Number(b.amount),
    })));

    setSavings((savingData || []).map(sv => ({
      id: sv.id,
      name: sv.name,
      target: Number(sv.target),
      current: Number(sv.current || 0),
    })));

    setLoans((loanData || []).map(l => ({
      id: l.id,
      name: l.name,
      bank: l.bank || "",
      principal: Number(l.principal),
      rate: Number(l.rate),
      term: Number(l.term),
      startDate: l.start,
      type: l.type,
      memo: l.memo || "",
    })));
  };

  const saveLocal = (key, value) => {
    localStorage.setItem(key, JSON.stringify(value));
  };

  const addAccount = () => {
    if (!accountName) return alert(tr("accountNameRequired"));
    const next = [...accounts, { id: Date.now(), name: accountName, balance: Number(accountBalance || 0) }];
    setAccounts(next);
    saveLocal("accounts", next);
    setAccountName("");
    setAccountBalance("");
    showPopup(language === "en" ? "Account added." : "계좌가 추가됐어요.", language === "en" ? "Account added" : "계좌 추가 완료", "🏦");
  };

  const deleteAccount = (id) => {
    const next = accounts.filter(a => a.id !== id);
    setAccounts(next);
    saveLocal("accounts", next);
  };

  const addCard = () => {
    if (!cardName || !cardPayDay) return alert(tr("cardRequired"));
    const next = [...cards, { id: Date.now(), name: cardName, payDay: Number(cardPayDay) }];
    setCards(next);
    saveLocal("cards", next);
    setCardName("");
    setCardPayDay("");
    showPopup(language === "en" ? "Card added." : "카드가 추가됐어요.", language === "en" ? "Card added" : "카드 추가 완료", "💳");
  };

  const deleteCard = (id) => {
    const next = cards.filter(c => c.id !== id);
    setCards(next);
    saveLocal("cards", next);
  };

  const addSub = () => {
    if (!subName || !subAmount) return alert(tr("subRequired"));
    const next = [...subs, { id: Date.now(), name: subName, amount: Number(subAmount), day: Number(subDay || 1) }];
    setSubs(next);
    saveLocal("subs", next);
    setSubName("");
    setSubAmount("");
    setSubDay("");
    showPopup(language === "en" ? "Subscription added." : "구독이 추가됐어요.", language === "en" ? "Subscription added" : "구독 추가 완료", "📱");
  };

  const deleteSub = (id) => {
    const next = subs.filter(s => s.id !== id);
    setSubs(next);
    saveLocal("subs", next);
  };

  const today = new Date().getDate();

  const totalAccountBalance = accounts.reduce((sum, a) => sum + Number(a.balance || 0), 0);
  const totalSubMonthly = subs.reduce((sum, s) => sum + Number(s.amount || 0), 0);

  const currentMonthExpenseByBudget = budgets.reduce((sum, b) => {
    return sum + txs
      .filter(t => {
        const d = new Date(t.date);
        return t.type === "expense"
          && t.category === b.category
          && d.getFullYear() === new Date().getFullYear()
          && d.getMonth() === new Date().getMonth();
      })
      .reduce((a, t) => a + Number(t.amount), 0);
  }, 0);

  const currentBudgetTotal = budgets.reduce((sum, b) => sum + Number(b.amount), 0);
  const budgetCarryOver = Math.max(0, currentBudgetTotal - currentMonthExpenseByBudget);
  const nextMonthBudgetTotal = currentBudgetTotal + budgetCarryOver;

  const getSavingPrediction = (sv) => {
    const savingTxs = txs
      .filter(t => t.type === "expense" && (t.category === "저축" || t.category === sv.name))
      .sort((a,b) => new Date(a.date) - new Date(b.date));

    const monthlyMap = {};

    savingTxs.forEach(t => {
      const d = new Date(t.date);
      const key = `${d.getFullYear()}-${d.getMonth() + 1}`;
      monthlyMap[key] = (monthlyMap[key] || 0) + Number(t.amount);
    });

    const values = Object.values(monthlyMap);
    const avg = values.length ? values.reduce((a,b) => a + b, 0) / values.length : 0;
    const left = Math.max(0, Number(sv.target) - Number(sv.current));

    if (!avg || left <= 0) return tr("unknownEstimate");

    const months = Math.ceil(left / avg);
    const targetDate = new Date();
    targetDate.setMonth(targetDate.getMonth() + months);

    return language === "en" ? `${targetDate.getFullYear()}-${targetDate.getMonth() + 1} estimate` : `${targetDate.getFullYear()}년 ${targetDate.getMonth() + 1}월 예상`;
  };

  const assetTrendLabels = Array.from({ length: 6 }, (_, i) => {
    const d = new Date();
    d.setMonth(d.getMonth() - (5 - i));
    return language === "en" ? `${d.getMonth() + 1}M` : `${d.getMonth() + 1}월`;
  });

  const assetTrendData = Array.from({ length: 6 }, (_, i) => {
    const d = new Date();
    d.setMonth(d.getMonth() - (5 - i));

    const monthIncome = txs
      .filter(t => {
        const td = new Date(t.date);
        return t.type === "income" && td.getFullYear() === d.getFullYear() && td.getMonth() === d.getMonth();
      })
      .reduce((sum,t) => sum + Number(t.amount), 0);

    const monthExpense = txs
      .filter(t => {
        const td = new Date(t.date);
        return t.type === "expense" && td.getFullYear() === d.getFullYear() && td.getMonth() === d.getMonth();
      })
      .reduce((sum,t) => sum + Number(t.amount), 0);

    return totalAccountBalance + monthIncome - monthExpense;
  });

  const getCardDDay = (payDay) => {
    const d = Number(payDay);
    if (!d) return "";
    if (d >= today) return `D-${d - today}`;
    const last = new Date(new Date().getFullYear(), new Date().getMonth() + 1, 0).getDate();
    return `D-${last - today + d}`;
  };

  const realTxs = txs.filter(t => t.type !== "transfer");

  const income = realTxs.filter(t => t.type === "income").reduce((s,t) => s + Number(t.amount), 0);
  const expense = realTxs.filter(t => t.type === "expense").reduce((s,t) => s + Number(t.amount), 0);

  const typeSummary = moneyTypes.map(mt => ({
    type: mt,
    income: txs.filter(t => t.type === "income" && t.moneyType === mt).reduce((s,t) => s + Number(t.amount), 0),
    expense: txs.filter(t => t.type === "expense" && t.moneyType === mt).reduce((s,t) => s + Number(t.amount), 0),
    transferIn: txs.filter(t => t.type === "transfer" && t.transferTo === mt).reduce((s,t) => s + Number(t.amount), 0),
    transferOut: txs.filter(t => t.type === "transfer" && t.moneyType === mt).reduce((s,t) => s + Number(t.amount), 0),
  }));

  const moveChartMonth = (diff) => {
    const d = new Date(chartYear, chartMonth + diff, 1);
    setChartYear(d.getFullYear());
    setChartMonth(d.getMonth());
  };

  const chartMonthTxs = realTxs.filter(t => {
    const d = new Date(t.date);
    return d.getFullYear() === chartYear && d.getMonth() === chartMonth;
  });

  const chartMonthIncome = chartMonthTxs
    .filter(t => t.type === "income")
    .reduce((s,t) => s + Number(t.amount), 0);

  const chartMonthExpense = chartMonthTxs
    .filter(t => t.type === "expense")
    .reduce((s,t) => s + Number(t.amount), 0);

  const chartIncomeMap = CATS_INC.map(cat => ({
    cat,
    amount: chartMonthTxs
      .filter(t => t.type === "income" && t.category === cat)
      .reduce((s,t) => s + Number(t.amount), 0),
  })).filter(x => x.amount > 0);

  const chartExpenseMap = CATS_EXP.map(cat => ({
    cat,
    amount: chartMonthTxs
      .filter(t => t.type === "expense" && t.category === cat)
      .reduce((s,t) => s + Number(t.amount), 0),
  })).filter(x => x.amount > 0);

  const createRepeatedTxs = (baseTx) => {
    if (baseTx.repeat === "none") return [baseTx];

    const list = [baseTx];
    const start = new Date(baseTx.date);

    const count = baseTx.repeat === "monthly" ? 12 : 8;

    for (let i = 1; i < count; i++) {
      const next = new Date(start);

      if (baseTx.repeat === "monthly") {
        next.setMonth(start.getMonth() + i);
      }

      if (baseTx.repeat === "weekly") {
        next.setDate(start.getDate() + i * 7);
      }

      list.push({
        ...baseTx,
        id: Date.now() + i,
        date: next.toISOString().slice(0, 10),
        repeatedFrom: baseTx.id,
      });
    }

    return list;
  };

  const saveTx = async () => {
    if (!amount) return alert(tr("amountRequired"));

    const finalType = moneyType !== transferTo && category === "이체" ? "transfer" : type;

    const baseTx = {
      id: editingTxId || Date.now(),
      type: finalType,
      amount: uncomma(amount),
      name: name || category,
      category,
      moneyType,
      transferTo: finalType === "transfer" ? transferTo : "",
      repeat,
      date: new Date().toISOString().slice(0, 10),
    };

    if (editingTxId) {
      const { data, error } = await supabase
        .from("transactions")
        .update({
          type: baseTx.type === "transfer" ? "expense" : baseTx.type,
          amount: baseTx.amount,
          date: baseTx.date,
          name: baseTx.name,
          category: baseTx.category,
          memo: baseTx.type === "transfer" ? `이체: ${baseTx.moneyType} → ${baseTx.transferTo}` : "",
          account: baseTx.moneyType || "계좌",
          repeat: baseTx.repeat || "none",
        })
        .eq("id", editingTxId)
        .eq("user_id", user.id)
        .select()
        .single();

      if (error) return alert(error.message);

      const updatedTx = {
        ...baseTx,
        id: data.id,
      };

      setTxs(txs.map(t => t.id === editingTxId ? updatedTx : t));
      showPopup(ui("입출금 내역이 수정됐어요.", "Transaction updated."), ui("수정 완료", "Updated"), "✏️");
      cancelTxEdit();
      return;
    }

    const newTxs = createRepeatedTxs(baseTx);
    const savedTxs = [];

    for (const tx of newTxs) {
      const { data, error } = await supabase
        .from("transactions")
        .insert({
          user_id: user.id,
          type: tx.type === "transfer" ? "expense" : tx.type,
          amount: tx.amount,
          date: tx.date,
          name: tx.name,
          category: tx.category,
          memo: tx.type === "transfer" ? `이체: ${tx.moneyType} → ${tx.transferTo}` : "",
          account: tx.moneyType || "계좌",
          repeat: tx.repeat || "none",
        })
        .select()
        .single();

      if (error) {
        alert(error.message);
        return;
      }

      savedTxs.push({
        ...tx,
        id: data.id,
      });
    }

    setTxs([...savedTxs, ...txs]);
    setAmount("");
    setName("");
    setRepeat("none");
    setMoneyType("계좌");
    setTransferTo("카드");
    setEditingTxId(null);

    if (repeat === "monthly") showPopup(ui("매월 반복 내역 12개월치가 추가됐어요.", "12 monthly recurring transactions added."), ui("반복 등록 완료", "Recurring entry saved"), "🔁");
    else if (repeat === "weekly") showPopup(ui("매주 반복 내역 8주치가 추가됐어요.", "8 weekly recurring transactions added."), ui("반복 등록 완료", "Recurring entry saved"), "🔁");
    else showPopup(ui("입출금 내역이 추가됐어요.", "Transaction added."), ui("기록 완료", "Saved"), "💸");
  };

  const editTx = (tx) => {
    setEditingTxId(tx.id);
    setType(tx.type === "transfer" ? "expense" : tx.type);
    setAmount(comma(tx.amount));
    setName(tx.name || "");
    setCategory(tx.type === "transfer" ? "이체" : tx.category || "식비");
    setMoneyType(tx.moneyType || "계좌");
    setTransferTo(tx.transferTo || "카드");
    setRepeat(tx.repeat || "none");
    setTab("home");
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const cancelTxEdit = () => {
    setEditingTxId(null);
    setType("expense");
    setAmount("");
    setName("");
    setCategory("식비");
    setMoneyType("계좌");
    setTransferTo("카드");
    setRepeat("none");
  };

  const deleteTx = async (txId) => {
    if (!window.confirm(ui("이 입출금 내역을 삭제할까요?", "Delete this transaction?"))) return;

    const { error } = await supabase
      .from("transactions")
      .delete()
      .eq("id", txId)
      .eq("user_id", user.id);

    if (error) return alert(error.message);

    setTxs(txs.filter(t => t.id !== txId));
    showPopup(ui("입출금 내역이 삭제됐어요.", "Transaction deleted."), ui("삭제 완료", "Deleted"), "🗑️");
  };

  const handleExcelUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    if (!rows.length) {
      alert(tr("excelNoData"));
      return;
    }

    const pickKey = (keys, patterns, fallbackIndex = 0) => {
      const lowered = keys.map(k => String(k).toLowerCase());
      const foundIndex = lowered.findIndex(k => patterns.some(p => k.includes(p.toLowerCase())));
      return foundIndex >= 0 ? keys[foundIndex] : keys[fallbackIndex];
    };

    const imported = rows.map((row, index) => {
      const keys = Object.keys(row);

      const dateKey = pickKey(keys, ["일자", "날짜", "거래일시", "date"], 0);
      const nameKey = pickKey(keys, ["내용", "적요", "거래기록사항", "거래내용", "메모", "memo", "description"], 1);
      const amountKey = pickKey(keys, ["거래금액", "금액", "amount"], 2);
      const typeKey = pickKey(keys, ["구분", "type"], -1);
      const categoryKey = pickKey(keys, ["카테고리", "category"], -1);
      const accountKey = pickKey(keys, ["출입금유형", "계좌", "account", "moneytype", "money type"], -1);
      const toAccountKey = pickKey(keys, ["이체도착", "to_account", "to account", "transfer_to"], -1);

      const rawName = row[nameKey] || row["내용"] || row["거래기록사항"] || row["거래내용"] || "엑셀 가져오기";
      const rawAmount = normalizeAmount(row[amountKey]);
      const rowText = JSON.stringify(row);
      const rawType = String(typeKey ? row[typeKey] : "");
      const categoryValue = categoryKey ? String(row[categoryKey] || "").trim() : "";
      const accountValue = accountKey ? String(row[accountKey] || "").trim() : "";
      const toAccountValue = toAccountKey ? String(row[toAccountKey] || "").trim() : "";

      let txType = "expense";
      let finalAmount = Math.abs(rawAmount);

      if (rawType.includes("입금") || rawType.toLowerCase().includes("income") || rowText.includes("수입") || rawAmount > 0) {
        txType = "income";
      }

      if (rawType.includes("출금") || rawType.toLowerCase().includes("expense") || rowText.includes("지출") || rawAmount < 0) {
        txType = "expense";
      }

      if (rawType.includes("이체") || rawType.toLowerCase().includes("transfer") || categoryValue === "이체") {
        txType = "transfer";
      }

      const finalCategory = categoryValue || guessCategory(rawName, txType === "transfer" ? "expense" : txType);
      const finalAccount = accountValue || row["account"] || row["출입금유형"] || "계좌";

      return {
        id: Date.now() + index,
        type: txType,
        amount: finalAmount,
        name: String(rawName),
        category: finalCategory,
        moneyType: finalAccount,
        transferTo: txType === "transfer" ? (toAccountValue || "") : "",
        repeat: "none",
        date: normalizeDate(row[dateKey]),
      };
    }).filter(t => t.amount > 0);

    const existingKeys = new Set(
      txs.map(t => `${t.date}-${t.name}-${t.amount}-${t.type}-${t.moneyType}`)
    );

    const filtered = imported.filter(t => {
      const key = `${t.date}-${t.name}-${t.amount}-${t.type}-${t.moneyType}`;
      return !existingKeys.has(key);
    });

    if (!filtered.length) {
      alert(ui("새로 추가할 내역이 없어요.", "No new transactions to import."));
      event.target.value = "";
      return;
    }

    const newTypes = [...new Set(filtered.map(t => t.moneyType).filter(Boolean))].filter(mt => !moneyTypes.includes(mt));
    if (newTypes.length) {
      const nextTypes = [...moneyTypes, ...newTypes];
      setMoneyTypes(nextTypes);
      localStorage.setItem("moneyTypes", JSON.stringify(nextTypes));
    }

    const savedTxs = [];

    for (const tx of filtered) {
      const { data, error } = await supabase
        .from("transactions")
        .insert({
          user_id: user.id,
          type: tx.type === "transfer" ? "expense" : tx.type,
          amount: tx.amount,
          date: tx.date,
          name: tx.name,
          category: tx.category,
          memo: tx.type === "transfer" ? `이체: ${tx.moneyType} → ${tx.transferTo || ""}` : "",
          account: tx.moneyType || "계좌",
          repeat: "none",
        })
        .select()
        .single();

      if (error) {
        alert(error.message);
        event.target.value = "";
        return;
      }

      savedTxs.push({ ...tx, id: data.id });
    }

    setTxs([...savedTxs, ...txs]);
    showPopup(
      ui(`${savedTxs.length}개 내역을 가져왔어요.`, `${savedTxs.length} transactions imported.`),
      ui("엑셀 가져오기 완료", "Import complete"),
      "📄"
    );

    event.target.value = "";
  };

  const saveBudget = async () => {
    if (!budgetAmount) return alert(tr("budgetRequired"));

    const { data, error } = await supabase
      .from("budgets")
      .upsert({
        user_id: user.id,
        category: budgetCat,
        amount: Number(budgetAmount),
      }, { onConflict: "user_id,category" })
      .select()
      .single();

    if (error) return alert(error.message);

    const next = budgets.filter(b => b.category !== budgetCat);
    setBudgets([...next, {
      id: data.id,
      category: data.category,
      amount: Number(data.amount),
    }]);
    setBudgetAmount("");
  };

  const saveSaving = async () => {
    if (!savingName || !savingTarget) return alert(tr("savingRequired"));

    if (editingSavingId) {
      const { data, error } = await supabase
        .from("savings")
        .update({
          name: savingName,
          target: Number(savingTarget),
          current: Number(savingCurrent || 0),
        })
        .eq("id", editingSavingId)
        .eq("user_id", user.id)
        .select()
        .single();

      if (error) return alert(error.message);

      setSavings(savings.map(s => s.id === editingSavingId ? {
        id: data.id,
        name: data.name,
        target: Number(data.target),
        current: Number(data.current || 0),
      } : s));
      showPopup(ui("저축 목표가 수정됐어요.", "Savings goal updated."), ui("수정 완료", "Updated"), "🐷");
    } else {
      const { data, error } = await supabase
        .from("savings")
        .insert({
          user_id: user.id,
          name: savingName,
          target: Number(savingTarget),
          current: Number(savingCurrent || 0),
        })
        .select()
        .single();

      if (error) return alert(error.message);

      setSavings([...savings, {
        id: data.id,
        name: data.name,
        target: Number(data.target),
        current: Number(data.current || 0),
      }]);
      showPopup(ui("저축 목표가 추가됐어요.", "Savings goal added."), ui("저축 목표 추가", "Savings goal added"), "🐷");
    }

    setEditingSavingId(null);
    setSavingName("");
    setSavingTarget("");
    setSavingCurrent("");
  };

  const editSaving = (saving) => {
    setEditingSavingId(saving.id);
    setSavingName(saving.name);
    setSavingTarget(String(saving.target));
    setSavingCurrent(String(saving.current));
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const cancelSavingEdit = () => {
    setEditingSavingId(null);
    setSavingName("");
    setSavingTarget("");
    setSavingCurrent("");
  };

  const deleteSaving = async (savingId) => {
    if (!window.confirm(ui("이 저축 목표를 삭제할까요?", "Delete this savings goal?"))) return;

    const { error } = await supabase
      .from("savings")
      .delete()
      .eq("id", savingId)
      .eq("user_id", user.id);

    if (error) return alert(error.message);

    setSavings(savings.filter(s => s.id !== savingId));
  };

  const totalSavingTarget = savings.reduce((sum, s) => sum + Number(s.target), 0);
  const totalSavingCurrent = savings.reduce((sum, s) => sum + Number(s.current), 0);
  const totalSavingPct = totalSavingTarget > 0
    ? Math.min(100, Math.round((totalSavingCurrent / totalSavingTarget) * 100))
    : 0;

  const saveLoan = async () => {
    if (!loanName || !loanPrincipal || !loanRate || !loanTerm) {
      return alert(tr("loanRequired"));
    }

    const payload = {
      name: loanName,
      principal: Number(loanPrincipal),
      rate: Number(loanRate),
      term: Number(loanTerm),
      start: loanStartDate,
      type: loanType,
    };

    if (editingLoanId) {
      const { data, error } = await supabase
        .from("loans")
        .update(payload)
        .eq("id", editingLoanId)
        .eq("user_id", user.id)
        .select()
        .single();

      if (error) return alert(error.message);

      setLoans(loans.map(l => l.id === editingLoanId ? {
        id: data.id,
        name: data.name,
        principal: Number(data.principal),
        rate: Number(data.rate),
        term: Number(data.term),
        startDate: data.start,
        type: data.type,
      } : l));
      showPopup(ui("대출건이 수정됐어요.", "Loan updated."), ui("수정 완료", "Updated"), "🏦");
    } else {
      const { data, error } = await supabase
        .from("loans")
        .insert({
          user_id: user.id,
          ...payload,
        })
        .select()
        .single();

      if (error) return alert(error.message);

      setLoans([...loans, {
        id: data.id,
        name: data.name,
        principal: Number(data.principal),
        rate: Number(data.rate),
        term: Number(data.term),
        startDate: data.start,
        type: data.type,
      }]);
      showPopup(ui("대출건이 추가됐어요. 입출금 카테고리에 대출명이 자동으로 표시돼요.", "Loan added. The loan name will appear in transaction categories."), ui("대출 추가 완료", "Loan added"), "🏦");
    }

    setEditingLoanId(null);
    setLoanName("");
    setLoanPrincipal("");
    setLoanRate("");
    setLoanTerm("");
    setLoanStartDate(new Date().toISOString().slice(0, 10));
    setLoanType("equal_payment");
  };

  const editLoan = (loan) => {
    setEditingLoanId(loan.id);
    setLoanName(loan.name);
    setLoanPrincipal(String(loan.principal));
    setLoanRate(String(loan.rate));
    setLoanTerm(String(loan.term));
    setLoanStartDate(loan.startDate || new Date().toISOString().slice(0, 10));
    setLoanType(loan.type || "equal_payment");
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const cancelLoanEdit = () => {
    setEditingLoanId(null);
    setLoanName("");
    setLoanPrincipal("");
    setLoanRate("");
    setLoanTerm("");
    setLoanStartDate(new Date().toISOString().slice(0, 10));
    setLoanType("equal_payment");
  };

  const deleteLoan = async (loanId) => {
    if (!window.confirm(ui("이 대출건을 삭제할까요?", "Delete this loan?"))) return;

    const { error } = await supabase
      .from("loans")
      .delete()
      .eq("id", loanId)
      .eq("user_id", user.id);

    if (error) return alert(error.message);

    setLoans(loans.filter(l => l.id !== loanId));
  };

  const calcMonthlyPayment = (loan) => {
    const P = loan.principal;
    const r = loan.rate / 100 / 12;
    const n = loan.term;

    if (loan.type === "equal_principal") return P / n + P * r;
    if (r === 0) return P / n;

    return P * r * Math.pow(1 + r, n) / (Math.pow(1 + r, n) - 1);
  };

  const getLoanPaidMonths = (loan) => {
    if (!loan.startDate) return 0;
    const start = new Date(loan.startDate);
    const now = new Date();
    const diff = (now.getFullYear() - start.getFullYear()) * 12 + (now.getMonth() - start.getMonth());
    return Math.max(0, Math.min(Number(loan.term), diff));
  };

  const getLoanTxAmount = (loan) => {
    return txs
      .filter(t => t.category === loan.name)
      .reduce((sum, t) => {
        if (t.type === "expense") return sum + Number(t.amount);
        if (t.type === "income") return sum - Number(t.amount);
        return sum;
      }, 0);
  };

  const getLoanSummary = (loan) => {
    const paidMonths = getLoanPaidMonths(loan);
    const totalMonths = Number(loan.term);
    const remainingMonths = Math.max(0, totalMonths - paidMonths);
    const monthlyPayment = calcMonthlyPayment(loan);
    const scheduledPaidAmount = Math.min(Number(loan.principal), monthlyPayment * paidMonths);
    const txPaidAmount = getLoanTxAmount(loan);
    const paidAmount = Math.max(0, scheduledPaidAmount + txPaidAmount);
    const remainingAmount = Math.max(0, Number(loan.principal) - paidAmount);
    const progress = Number(loan.principal) > 0
      ? Math.min(100, Math.round((paidAmount / Number(loan.principal)) * 100))
      : 0;
    const actualPaidMonths = monthlyPayment > 0
      ? Math.min(totalMonths, Math.floor(paidAmount / monthlyPayment))
      : paidMonths;

    const actualRemainingMonths = Math.max(0, totalMonths - actualPaidMonths);
    const actualYearsLeft = Math.floor(actualRemainingMonths / 12);
    const actualMonthsLeft = actualRemainingMonths % 12;

    return {
      paidMonths: actualPaidMonths,
      totalMonths,
      remainingMonths: actualRemainingMonths,
      monthlyPayment,
      paidAmount,
      remainingAmount,
      progress,
      yearsLeft: actualYearsLeft,
      monthsLeft: actualMonthsLeft,
    };
  };

  const catExpenseMap = CATS_EXP.map(cat => ({
    cat,
    amount: realTxs
      .filter(t => t.type === "expense" && t.category === cat)
      .reduce((s,t) => s + Number(t.amount), 0),
  })).filter(x => x.amount > 0);

  const loanCategoryNames = loans.map(l => l.name);
  const cats = type === "expense"
    ? [...CATS_EXP, ...loanCategoryNames]
    : [...CATS_INC, ...loanCategoryNames];

  const filteredSortedTxs = txs
    .filter(t => {
      if (!t.date) return true;

      if (txPeriodMode === "month") {
        return String(t.date).slice(0, 7) === txSelectedMonth;
      }

      if (txPeriodMode === "custom") {
        if (txStartDate && t.date < txStartDate) return false;
        if (txEndDate && t.date > txEndDate) return false;
        return true;
      }

      return true;
    })
    .sort((a, b) => {
      if (txSort === "latest") return new Date(b.date) - new Date(a.date);
      if (txSort === "oldest") return new Date(a.date) - new Date(b.date);
      if (txSort === "highest") return Number(b.amount) - Number(a.amount);
      if (txSort === "lowest") return Number(a.amount) - Number(b.amount);
      if (txSort === "income") return (b.type === "income") - (a.type === "income");
      if (txSort === "expense") return (b.type === "expense") - (a.type === "expense");
      return 0;
    });

  if (loading) return <div className="auth-screen">{tr("loading")}</div>;

  if (!user) {
    return (
      <div className="auth-screen">
        <div className="auth-logo">kko<span>Money</span></div>
        <div className="auth-sub">{tr("authSub")}</div>

        <div className="auth-card">
          <div className="auth-title">{tr("login")}</div>

          <div className="form-group">
            <label className="form-label">{tr("email")}</label>
            <input className="form-input" value={email} onChange={e => setEmail(e.target.value)} />
          </div>

          <div className="form-group">
            <label className="form-label">{tr("password")}</label>
            <input className="form-input" type="password" value={password} onChange={e => setPassword(e.target.value)} />
          </div>

          <button className="auth-btn" onClick={login}>{tr("login")}</button>
          <button className="auth-btn secondary" onClick={signup}>{tr("signup")}</button>
        </div>
      </div>
    );
  }

  return (
    <>
    {popup && (
      <div className="kko-popup-overlay">
        <div className="kko-popup">
          <div className="kko-popup-emoji">{popup.emoji}</div>
          <div className="kko-popup-title">{popup.title}</div>
          <div className="kko-popup-message">{popup.message}</div>
          <button className="kko-popup-btn" onClick={closePopup}>{tr("confirm")}</button>
        </div>
      </div>
    )}

    <div className="desktop-wrap">
      <aside className="sidebar">
        <div className="logo">kko<span>Money</span></div>

        <div className="nav-section">
          <div className="nav-section-title">{tr("main")}</div>
          {[
            ["home", "🏠", "home"],
            ["tx", "🧾", "tx"],
            ["chart", "📊", "chart"],
          ].map(([key, icon, label]) => (
            <button key={key} className={`nav-btn ${tab === key ? "active" : ""}`} onClick={() => setTab(key)}>
              <span className="nav-ico">{icon}</span>
              <span>{tr(label)}</span>
            </button>
          ))}
        </div>

        <div className="nav-section">
          <div className="nav-section-title">{tr("money")}</div>
          {[
            ["budget", "💗", "budget"],
            ["saving", "🐷", "saving"],
            ["assets", "🏦", "assets"],
            ["loan", "🏠", "loan"],
            ["subs", "📱", "subs"],
          ].map(([key, icon, label]) => (
            <button key={key} className={`nav-btn ${tab === key ? "active" : ""}`} onClick={() => setTab(key)}>
              <span className="nav-ico">{icon}</span>
              <span>{tr(label)}</span>
            </button>
          ))}
        </div>

        <div className="nav-section nav-bottom">
          <div className="nav-section-title">{tr("system")}</div>
          <button className={`nav-btn ${tab === "settings" ? "active" : ""}`} onClick={() => setTab("settings")}>
            <span className="nav-ico">⚙️</span>
            <span>{tr("settings")}</span>
          </button>
          <button className="nav-btn logout-btn" onClick={logout}>
            <span className="nav-ico">🚪</span>
            <span>{tr("logout")}</span>
          </button>
        </div>
      </aside>

      <main className="main-content">
        <div className="topbar">
          <div className="topbar-spacer" />
          <div className="top-controls">
            <select
              className="top-select"
              value={language}
              onChange={(e) => {
                setLanguage(e.target.value);
                localStorage.setItem("language", e.target.value);
              }}
            >
              <option value="ko">🌐 KO</option>
              <option value="en">🌐 EN</option>
            </select>

            <select
              className="top-select"
              value={currency}
              onChange={(e) => {
                setCurrency(e.target.value);
                localStorage.setItem("currency", e.target.value);
              }}
            >
              {CURRENCY_OPTIONS.map(c => (
                <option key={c.code} value={c.code}>{c.label}</option>
              ))}
            </select>

            <div className="top-avatar">🐷</div>
          </div>
        </div>
        {tab === "home" && (
          <>
            <section className="banner hero-card">
              <div className="banner-lbl">{tr("remainingBudget")}</div>
              <div className="banner-main">{won(income - expense)}</div>
              <div className="banner-row">
                <div className="banner-mini">{tr("income")}<span>{won(income)}</span></div>
                <div className="banner-mini">{tr("expense")}<span>{won(expense)}</span></div>
                <div className="banner-mini">{tr("transfer")}<span>{tr("transferExcluded")}</span></div>
              </div>
              <div className="hero-pig" aria-hidden="true">🐷</div>
              <div className="hero-coin coin-a" aria-hidden="true">🪙</div>
              <div className="hero-coin coin-b" aria-hidden="true">🪙</div>
            </section>

            <section className="card quick-card">
              <div className="card-title">{editingTxId ? tr("editTx") : tr("quickInput")}</div>

              <div className="type-toggle">
                <button className={`type-btn ${type === "expense" ? "active-expense" : ""}`} onClick={() => {setType("expense"); setCategory("식비");}}>
                  {tr("expenseButton")}
                </button>
                <button className={`type-btn ${type === "income" ? "active-income" : ""}`} onClick={() => {setType("income"); setCategory("월급");}}>
                  {tr("incomeButton")}
                </button>
              </div>

              <div className="form-group">
                <label className="form-label">{tr("amount")}</label>
                <input className="form-input" type="text" placeholder={language === "en" ? "Ex. 50,000" : "예: 50,000"} value={amount} onChange={e => setAmount(comma(e.target.value))} />
              </div>

              <div className="form-group">
                <label className="form-label">{tr("memo")}</label>
                <input className="form-input" placeholder={language === "en" ? "Description" : "내용 입력"} value={name} onChange={e => setName(e.target.value)} />
              </div>

              <div className="form-group">
                <label className="form-label">{tr("category")}</label>
                <select className="form-select" value={category} onChange={e => setCategory(e.target.value)}>
                  {cats.map(c => <option key={c} value={c}>{catLabel(c)}</option>)}
                  <option value="이체">{tr("transfer")}</option>
                </select>
              </div>

              <div className="form-group">
                <label className="form-label">{tr("moneyType")}</label>
                <select className="form-select" value={moneyType} onChange={e => setMoneyType(e.target.value)}>
                  {moneyTypes.map(mt => <option key={mt} value={mt}>{moneyTypeLabel(mt)}</option>)}
                </select>
              </div>

              {category === "이체" && (
                <div className="form-group">
                  <label className="form-label">{tr("transferTo")}</label>
                  <select className="form-select" value={transferTo} onChange={e => setTransferTo(e.target.value)}>
                    {moneyTypes.map(mt => <option key={mt} value={mt}>{moneyTypeLabel(mt)}</option>)}
                  </select>
                  <div className="hint">{tr("transferHint")}</div>
                </div>
              )}

              <div className="form-group">
                <label className="form-label">{tr("repeat")}</label>
                <select className="form-select" value={repeat} onChange={e => setRepeat(e.target.value)}>
                  <option value="none">{tr("noRepeat")}</option>
                  <option value="monthly">{tr("monthlyRepeat")}</option>
                  <option value="weekly">{tr("weeklyRepeat")}</option>
                </select>
              </div>

              <div className="tx-button-row">
                <button className="btn btn-primary" onClick={saveTx}>
                  {editingTxId ? tr("updateSave") : tr("save")}
                </button>
                {editingTxId && (
                  <button className="btn btn-secondary" onClick={cancelTxEdit}>
                    {tr("cancel")}
                  </button>
                )}
              </div>
            </section>

            <section className="dashboard-grid">
              <div className="dash-card summary-card">
                <div className="dash-title">{language === "en" ? "This Month Summary" : "이번 달 요약"}</div>
                <div className="summary-list">
                  <div className="summary-row">
                    <div className="summary-left"><span className="round-icon income-bg">↑</span>{tr("income")}</div>
                    <strong className="pink">{won(income)}</strong>
                  </div>
                  <div className="summary-row">
                    <div className="summary-left"><span className="round-icon expense-bg">↓</span>{tr("expense")}</div>
                    <strong className="blue">{won(expense)}</strong>
                  </div>
                  <div className="summary-row">
                    <div className="summary-left"><span className="round-icon transfer-bg">↔</span>{tr("transfer")}</div>
                    <strong className="purple">{tr("transferExcluded")}</strong>
                  </div>
                </div>
              </div>

              <div className="dash-card budget-card">
                <div className="dash-title">{tr("budgetUsage")}</div>
                <div className="budget-usage-wrap">
                  <div
                    className="donut"
                    style={{
                      "--pct": `${Math.min(100, currentBudgetTotal > 0 ? Math.round((currentMonthExpenseByBudget / currentBudgetTotal) * 100) : 0)}%`
                    }}
                  >
                    <div>
                      <strong>{currentBudgetTotal > 0 ? Math.round((currentMonthExpenseByBudget / currentBudgetTotal) * 100) : 0}%</strong>
                      <span>{tr("used")}</span>
                    </div>
                  </div>

                  <div className="budget-lines">
                    <div><span>{tr("totalBudget")}</span><b>{won(currentBudgetTotal)}</b></div>
                    <div><span>{tr("used")}</span><b>{won(currentMonthExpenseByBudget)}</b></div>
                    <div><span>{tr("carryOver")}</span><b>{won(budgetCarryOver)}</b></div>
                    <hr />
                    <div><span>{tr("nextBudget")}</span><b>{won(nextMonthBudgetTotal)}</b></div>
                  </div>
                </div>
              </div>

              <div className="dash-card savings-card-home">
                <div className="dash-title">{tr("savingProgress")}</div>
                <div className="simple-metric">
                  <span>{tr("totalTarget")}</span>
                  <strong>{won(totalSavingTarget)}</strong>
                </div>
                <div className="simple-metric">
                  <span>{tr("totalCurrent")}</span>
                  <strong>{won(totalSavingCurrent)}</strong>
                </div>
                <div className="bar-bg"><div className="bar-fill green" style={{ width: totalSavingPct + "%" }} /></div>
              </div>

              <div className="dash-card total-assets-card">
                <div className="dash-title">{tr("assets")}</div>
                <div className="big-metric">{won(totalAccountBalance + totalSavingCurrent)}</div>
                <button className="floating-plus" onClick={() => setTab("assets")}>+</button>
              </div>
            </section>
          </>
        )}

        {tab === "tx" && (
          <section className="card">
            <div className="card-title">{tr("txList")}</div>

            <div className="tx-filter-card">
              <div className="tx-filter-row">
                <div className="filter-field">
                  <label>{language === "en" ? "Period" : "기간"}</label>
                  <select className="form-select" value={txPeriodMode} onChange={e => setTxPeriodMode(e.target.value)}>
                    <option value="month">{language === "en" ? "Monthly" : "월별"}</option>
                    <option value="custom">{language === "en" ? "Custom period" : "기간 설정"}</option>
                    <option value="all">{language === "en" ? "All" : "전체"}</option>
                  </select>
                </div>

                {txPeriodMode === "month" && (
                  <div className="filter-field">
                    <label>{language === "en" ? "Month" : "월 선택"}</label>
                    <input className="form-input" type="month" value={txSelectedMonth} onChange={e => setTxSelectedMonth(e.target.value)} />
                  </div>
                )}

                {txPeriodMode === "custom" && (
                  <>
                    <div className="filter-field">
                      <label>{language === "en" ? "From" : "시작일"}</label>
                      <input className="form-input" type="date" value={txStartDate} onChange={e => setTxStartDate(e.target.value)} />
                    </div>
                    <div className="filter-field">
                      <label>{language === "en" ? "To" : "종료일"}</label>
                      <input className="form-input" type="date" value={txEndDate} onChange={e => setTxEndDate(e.target.value)} />
                    </div>
                  </>
                )}

                <div className="filter-field">
                  <label>{language === "en" ? "Sort" : "정렬"}</label>
                  <select className="form-select" value={txSort} onChange={e => setTxSort(e.target.value)}>
                    <option value="latest">{language === "en" ? "Latest first" : "최신순"}</option>
                    <option value="oldest">{language === "en" ? "Oldest first" : "오래된순"}</option>
                    <option value="highest">{language === "en" ? "Highest amount" : "금액 높은순"}</option>
                    <option value="lowest">{language === "en" ? "Lowest amount" : "금액 낮은순"}</option>
                    <option value="income">{language === "en" ? "Income first" : "수입 먼저"}</option>
                    <option value="expense">{language === "en" ? "Expense first" : "지출 먼저"}</option>
                  </select>
                </div>
              </div>

              <div className="tx-filter-summary">
                {language === "en" ? "Showing" : "표시 중"} <b>{filteredSortedTxs.length}</b>{language === "en" ? " transactions" : "건"}
              </div>
            </div>

            <label className="file-upload-ui">
                <input type="file" accept=".xlsx,.xls,.csv" onChange={handleExcelUpload} />
                <span className="file-upload-btn">{language === "en" ? "Choose file" : "파일 선택"}</span>
                <span className="file-upload-text">{language === "en" ? "No file selected" : "선택된 파일 없음"}</span>
              </label>

            <div className="grid3">
              <div className="sum-card"><div className="sum-lbl">{tr("income")}</div><div className="sum-val income">{won(income)}</div></div>
              <div className="sum-card"><div className="sum-lbl">{tr("expense")}</div><div className="sum-val expense">{won(expense)}</div></div>
              <div className="sum-card"><div className="sum-lbl">{tr("balance")}</div><div className="sum-val">{won(income - expense)}</div></div>
            </div>

            <div className="tx-list">
              {filteredSortedTxs.length === 0 ? <div className="empty">{tr("noTx")}</div> : filteredSortedTxs.map(t => (
                <div className="tx-row" key={t.id}>
                  <div className="tx-info">
                    <div className="tx-name">{t.name}</div>
                    <div className="tx-meta">
                      {t.date} · {catLabel(t.category)} · {moneyTypeLabel(t.moneyType || "계좌")}
                      {t.type === "transfer" ? ` → ${moneyTypeLabel(t.transferTo)}` : ""}
                      {t.repeat !== "none" ? ` · ${t.repeat === "monthly" ? tr("repeatMonthly") : tr("repeatWeekly")}` : ""}
                    </div>
                  </div>
                  <div className={`tx-amt ${t.type}`}>{t.type === "transfer" ? `${tr("txTransferPrefix")} ` : t.type === "income" ? "+" : "-"}{won(t.amount)}</div>
                  <div className="tx-actions">
                    <button onClick={() => editTx(t)}>{tr("edit")}</button>
                    <button onClick={() => deleteTx(t.id)}>{tr("delete")}</button>
                  </div>
                </div>
              ))}
            </div>
          </section>
        )}

        {tab === "budget" && (
          <>
            <section className="banner">
              <div className="banner-lbl">{tr("budgetUsage")}</div>
              <div className="banner-main">
                {(() => {
                  const budgetTotal = budgets.reduce((sum,b) => sum + Number(b.amount), 0);
                  const budgetUsed = budgets.reduce((sum,b) => {
                    return sum + txs
                      .filter(t => t.type === "expense" && t.category === b.category)
                      .reduce((a,t) => a + Number(t.amount), 0);
                  }, 0);
                  return budgetTotal === 0 ? "0%" : Math.round((budgetUsed / budgetTotal) * 100) + "%";
                })()}
              </div>
              <div className="banner-row">
                <div className="banner-mini">{tr("totalBudget")}<span>{won(budgets.reduce((s,b) => s + b.amount, 0))}</span></div>
                <div className="banner-mini">{tr("used")}<span>{won(currentMonthExpenseByBudget)}</span></div>
                <div className="banner-mini">{tr("carryOver")}<span>{won(budgetCarryOver)}</span></div>
                <div className="banner-mini">{tr("nextBudget")}<span>{won(nextMonthBudgetTotal)}</span></div>
              </div>
            </section>

            <section className="card">
              <div className="card-title">{tr("budgetByCategory")}</div>
              <select className="form-select" value={budgetCat} onChange={e => setBudgetCat(e.target.value)}>
                {CATS_EXP.map(c => <option key={c} value={c}>{catLabel(c)}</option>)}
              </select>
              <input className="form-input" type="number" placeholder={tr("monthlyBudget")} value={budgetAmount} onChange={e => setBudgetAmount(e.target.value)} />
              <button className="btn btn-primary" onClick={saveBudget}>{tr("saveBudget")}</button>

              <div className="budget-list">
                {budgets.map(b => {
                  const used = txs.filter(t => t.type === "expense" && t.category === b.category).reduce((s,t) => s + Number(t.amount), 0);
                  const pct = Math.min(100, Math.round(used / b.amount * 100));
                  return (
                    <div className="budget-item" key={b.category}>
                      <div className="budget-top"><b>{catLabel(b.category)}</b><span>{won(used)} / {won(b.amount)}</span></div>
                      <div className="bar-bg"><div className="bar-fill" style={{ width: pct + "%" }} /></div>
                    </div>
                  );
                })}
              </div>
            </section>

          </>
        )}

        {tab === "saving" && (
          <>
            <section className="banner saving-banner">
              <div className="banner-lbl">{tr("savingProgress")}</div>
              <div className="banner-main">{totalSavingPct}%</div>
              <div className="banner-row">
                <div className="banner-mini">{tr("totalTarget")}<span>{won(totalSavingTarget)}</span></div>
                <div className="banner-mini">{tr("totalCurrent")}<span>{won(totalSavingCurrent)}</span></div>
                <div className="banner-mini">{tr("leftAmount")}<span>{won(Math.max(0, totalSavingTarget - totalSavingCurrent))}</span></div>
              </div>
            </section>

            <section className="card">
              <div className="card-title">{editingSavingId ? tr("editSaving") : tr("addSaving")}</div>

              <input
                className="form-input"
                placeholder={tr("savingNamePlaceholder")}
                value={savingName}
                onChange={e => setSavingName(e.target.value)}
              />
              <input
                className="form-input"
                placeholder={tr("savingTargetPlaceholder")}
                type="number"
                value={savingTarget}
                onChange={e => setSavingTarget(e.target.value)}
              />
              <input
                className="form-input"
                placeholder={tr("savingCurrentPlaceholder")}
                type="number"
                value={savingCurrent}
                onChange={e => setSavingCurrent(e.target.value)}
              />

              <div className="saving-button-row">
                <button className="btn btn-primary" onClick={saveSaving}>
                  {editingSavingId ? tr("updateSave") : tr("addSaving")}
                </button>
                {editingSavingId && (
                  <button className="btn btn-secondary" onClick={cancelSavingEdit}>
                    {tr("cancel")}
                  </button>
                )}
              </div>
            </section>

            <section className="card">
              <div className="card-title">{tr("savingStatus")}</div>

              {savings.length === 0 ? (
                <div className="empty">{tr("noSaving")}</div>
              ) : (
                <div className="saving-grid">
                  {savings.map(sv => {
                    const pct = Number(sv.target) > 0
                      ? Math.min(100, Math.round((Number(sv.current) / Number(sv.target)) * 100))
                      : 0;
                    const left = Math.max(0, Number(sv.target) - Number(sv.current));

                    return (
                      <div className="saving-card" key={sv.id}>
                        <div className="saving-card-top">
                          <div>
                            <h3>{sv.name}</h3>
                            <p>{won(sv.current)} / {won(sv.target)}</p>
                          </div>
                          <strong>{pct}%</strong>
                        </div>

                        <div className="bar-bg">
                          <div className="bar-fill green" style={{ width: pct + "%" }} />
                        </div>

                        <div className="saving-detail">
                          <span>{tr("leftAmount")}</span>
                          <b>{won(left)}</b>
                        </div>

                        <div className="saving-detail">
                          <span>{tr("expectedDate")}</span>
                          <b>{getSavingPrediction(sv)}</b>
                        </div>

                        <div className="saving-actions">
                          <button className="btn btn-secondary" onClick={() => editSaving(sv)}>{tr("edit")}</button>
                          <button className="btn btn-danger" onClick={() => deleteSaving(sv.id)}>{tr("delete")}</button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </section>
          </>
        )}

        {tab === "assets" && (
          <>
            <section className="banner asset-banner">
              <div className="banner-lbl">{tr("totalAccountBalance")}</div>
              <div className="banner-main">{won(totalAccountBalance)}</div>
              <div className="banner-row">
                <div className="banner-mini">{tr("accountCount")}<span>{accounts.length}</span></div>
                <div className="banner-mini">{tr("cardCount")}<span>{cards.length}</span></div>
              </div>
            </section>

            <section className="card">
              <div className="card-title">{tr("accountBalanceManage")}</div>
              <input className="form-input" placeholder={tr("accountNamePlaceholder")} value={accountName} onChange={e => setAccountName(e.target.value)} />
              <input className="form-input" placeholder={tr("currentBalance")} type="number" value={accountBalance} onChange={e => setAccountBalance(e.target.value)} />
              <button className="btn btn-primary asset-add-btn" onClick={addAccount}>{tr("addAccount")}</button>

              <div className="asset-list">
                {accounts.map(a => (
                  <div className="asset-item" key={a.id}>
                    <div>
                      <b>{a.name}</b>
                      <span>{won(a.balance)}</span>
                    </div>
                    <button onClick={() => deleteAccount(a.id)}>{tr("delete")}</button>
                  </div>
                ))}
              </div>
            </section>

            <section className="card">
              <div className="card-title">{tr("assetTrend")}</div>
              <div className="chart-box wide">
                <Line
                  data={{
                    labels: assetTrendLabels,
                    datasets: [{
                      label: tr("expectedAssets"),
                      data: assetTrendData,
                    }],
                  }}
                />
              </div>
              <div className="hint">{tr("assetTrendHint")}</div>
            </section>

            <section className="card">
              <div className="card-title">{tr("cardPaymentManage")}</div>
              <input className="form-input" placeholder={tr("cardNamePlaceholder")} value={cardName} onChange={e => setCardName(e.target.value)} />
              <input className="form-input" placeholder={tr("paymentDayPlaceholder")} type="number" value={cardPayDay} onChange={e => setCardPayDay(e.target.value)} />
              <button className="btn btn-primary asset-add-btn" onClick={addCard}>{tr("addCard")}</button>

              <div className="asset-list">
                {cards.map(c => (
                  <div className="asset-item" key={c.id}>
                    <div>
                      <b>{c.name}</b>
                      <span>{`${tr("monthlyDay")} ${c.payDay} · ${getCardDDay(c.payDay)}`}</span>
                    </div>
                    <button onClick={() => deleteCard(c.id)}>{tr("delete")}</button>
                  </div>
                ))}
              </div>
            </section>
          </>
        )}

        {tab === "subs" && (
          <>
            <section className="banner sub-banner">
              <div className="banner-lbl">{tr("monthlySubscriptions")}</div>
              <div className="banner-main">{won(totalSubMonthly)}</div>
              <div className="banner-row">
                <div className="banner-mini">{tr("subCount")}<span>{subs.length}</span></div>
              </div>
            </section>

            <section className="card">
              <div className="card-title">{tr("subManage")}</div>
              <input className="form-input" placeholder={tr("subNamePlaceholder")} value={subName} onChange={e => setSubName(e.target.value)} />
              <input className="form-input" placeholder={tr("subAmountPlaceholder")} type="number" value={subAmount} onChange={e => setSubAmount(e.target.value)} />
              <input className="form-input" placeholder={tr("subDayPlaceholder")} type="number" value={subDay} onChange={e => setSubDay(e.target.value)} />
              <button className="btn btn-primary asset-add-btn" onClick={addSub}>{tr("addSub")}</button>

              <div className="asset-list">
                {subs.map(su => (
                  <div className="asset-item" key={su.id}>
                    <div>
                      <b>{su.name}</b>
                      <span>{`${won(su.amount)} · ${tr("monthlyDay")} ${su.day}`}</span>
                    </div>
                    <button onClick={() => deleteSub(su.id)}>{tr("delete")}</button>
                  </div>
                ))}
              </div>
            </section>
          </>
        )}

        {tab === "loan" && (
          <section className="card">
            <div className="card-title">{editingLoanId ? tr("editLoan") : tr("loanManage")}</div>

            <input className="form-input" placeholder={tr("loanNamePlaceholder")} value={loanName} onChange={e => setLoanName(e.target.value)} />
            <input className="form-input" placeholder={tr("principalPlaceholder")} type="number" value={loanPrincipal} onChange={e => setLoanPrincipal(e.target.value)} />
            <input className="form-input" placeholder={tr("ratePlaceholder")} type="number" value={loanRate} onChange={e => setLoanRate(e.target.value)} />
            <input className="form-input" placeholder={tr("termPlaceholder")} type="number" value={loanTerm} onChange={e => setLoanTerm(e.target.value)} />

            <label className="form-label loan-date-label">{tr("loanStartDate")}</label>
            <input className="form-input" type="date" value={loanStartDate} onChange={e => setLoanStartDate(e.target.value)} />

            <select className="form-select" value={loanType} onChange={e => setLoanType(e.target.value)}>
              <option value="equal_payment">{tr("equalPayment")}</option>
              <option value="equal_principal">{tr("equalPrincipal")}</option>
            </select>

            <div className="loan-button-row">
              <button className="btn btn-primary" onClick={saveLoan}>
                {editingLoanId ? tr("updateSave") : tr("addLoan")}
              </button>
              {editingLoanId && (
                <button className="btn btn-secondary" onClick={cancelLoanEdit}>
                  {tr("cancel")}
                </button>
              )}
            </div>

            <div className="loan-list">
              {loans.map(l => {
                const summary = getLoanSummary(l);

                return (
                  <div className="loan-card" key={l.id}>
                    <div className="loan-head">
                      <div>
                        <h3>{l.name}</h3>
                        <p>{l.type === "equal_payment" ? tr("loanTypeEqualPayment") : tr("loanTypeEqualPrincipal")} · {tr("annualRateShort")} {l.rate}% · {l.term}{tr("monthsUnit")}</p>
                      </div>
                      <div className="loan-progress-num">{summary.progress}%</div>
                    </div>

                    <div className="loan-progress-bg">
                      <div className="loan-progress-fill" style={{ width: summary.progress + "%" }} />
                    </div>

                    <div className="loan-info-grid loan-info-grid-auto">
                      <div>
                        <span>{tr("loanPrincipal")}</span>
                        <b>{won(l.principal)}</b>
                      </div>
                      <div>
                        <span>{tr("paidAmount")}</span>
                        <b>{won(summary.paidAmount)}</b>
                      </div>
                      <div>
                        <span>{tr("txApplied")}</span>
                        <b>{won(getLoanTxAmount(l))}</b>
                      </div>
                      <div>
                        <span>{tr("remainingLoan")}</span>
                        <b>{won(summary.remainingAmount)}</b>
                      </div>
                      <div>
                        <span>{tr("monthlyPayment")}</span>
                        <b>{won(summary.monthlyPayment)}</b>
                      </div>
                    </div>

                    <div className="loan-detail">
                      <div dangerouslySetInnerHTML={{__html: fmt("loanPaidSentence", { total: summary.totalMonths, paid: `<b>${summary.paidMonths}</b>` })}} />
                      <div>{summary.yearsLeft > 0 ? fmt("loanRemainingYearSentence", { months: summary.remainingMonths, years: summary.yearsLeft, leftMonths: summary.monthsLeft }) : fmt("loanRemainingSentence", { months: summary.remainingMonths })}</div>
                      <div>{tr("startDate")}: {l.startDate || tr("notEntered")}</div>
                    </div>

                    <div className="loan-card-actions">
                      <button className="btn btn-secondary" onClick={() => editLoan(l)}>{tr("edit")}</button>
                      <button className="btn btn-danger" onClick={() => deleteLoan(l.id)}>{tr("delete")}</button>
                    </div>
                  </div>
                );
              })}
            </div>
          </section>
        )}

        {tab === "settings" && (
          <section className="card">
            <div className="language-setting-card">
              <div className="card-title">{tr("language")}</div>
              <select
                className="form-select"
                value={language}
                onChange={(e) => {
                  setLanguage(e.target.value);
                  localStorage.setItem("language", e.target.value);
                }}
              >
                <option value="ko">{tr("korean")}</option>
                <option value="en">{tr("english")}</option>
              </select>
            </div>

            <div className="currency-setting-card">
              <div className="card-title">{tr("currency")}</div>
              <select
                className="form-select"
                value={currency}
                onChange={(e) => {
                  setCurrency(e.target.value);
                  localStorage.setItem("currency", e.target.value);
                }}
              >
                {CURRENCY_OPTIONS.map(c => (
                  <option key={c.code} value={c.code}>{c.label}</option>
                ))}
              </select>
            </div>

            <div className="card-title">{tr("moneyTypeEditorTitle")}</div>

            <div className="form-group">
              <label className="form-label">{tr("newTypeAdd")}</label>
              <div className="settings-row">
                <input
                  className="form-input"
                  value={newMoneyType}
                  onChange={e => setNewMoneyType(e.target.value)}
                  placeholder={tr("newTypePlaceholder")}
                />
                <button
                  className="btn btn-primary"
                  onClick={() => {
                    const v = newMoneyType.trim();
                    if (!v) return alert(tr("addTypeRequired"));
                    if (moneyTypes.includes(v)) return alert(tr("duplicateType"));
                    const next = [...moneyTypes, v];
                    setMoneyTypes(next);
                    localStorage.setItem("moneyTypes", JSON.stringify(next));
                    setNewMoneyType("");
                  }}
                >
                  {tr("add")}
                </button>
              </div>
            </div>

            <div className="type-edit-list">
              {moneyTypes.map(mt => (
                <div className="type-edit-item" key={mt}>
                  <span>{mt}</span>
                  <button
                    className="small-delete"
                    onClick={() => {
                      if (moneyTypes.length <= 1) return alert(tr("minTypeRequired"));
                      const next = moneyTypes.filter(x => x !== mt);
                      setMoneyTypes(next);
                      localStorage.setItem("moneyTypes", JSON.stringify(next));
                      if (moneyType === mt) setMoneyType(next[0]);
                      if (transferTo === mt) setTransferTo(next[0]);
                    }}
                  >
                    {tr("delete")}
                  </button>
                </div>
              ))}
            </div>

            <div className="hint">{tr("moneyTypeHint")}</div>
          </section>
        )}

        {tab === "chart" && (
          <section className="card">
            <div className="chart-month-head">
              <button className="month-arrow" onClick={() => moveChartMonth(-1)}>‹</button>
              <div>
                <div className="card-title">{tr("chart")}</div>
                <div className="chart-month-title">{fmt("chartMonthTitle", { year: chartYear, month: chartMonth + 1 })}</div>
              </div>
              <button className="month-arrow" onClick={() => moveChartMonth(1)}>›</button>
            </div>

            <div className="grid3">
              <div className="sum-card">
                <div className="sum-lbl">{tr("income")}</div>
                <div className="sum-val income">{won(chartMonthIncome)}</div>
              </div>
              <div className="sum-card">
                <div className="sum-lbl">{tr("expense")}</div>
                <div className="sum-val expense">{won(chartMonthExpense)}</div>
              </div>
              <div className="sum-card">
                <div className="sum-lbl">{tr("balance")}</div>
                <div className="sum-val">{won(chartMonthIncome - chartMonthExpense)}</div>
              </div>
            </div>

            <div className="chart-grid">
              <div className="chart-box">
                <div className="card-title">income category</div>
                {chartIncomeMap.length === 0 ? (
                  <div className="empty">{tr("noIncomeData")}</div>
                ) : (
                  <Doughnut
                    data={{
                      labels: chartIncomeMap.map(x => x.cat),
                      datasets: [{
                        label: tr("income"),
                        data: chartIncomeMap.map(x => x.amount),
                      }],
                    }}
                  />
                )}
              </div>

              <div className="chart-box">
                <div className="card-title">expense category</div>
                {chartExpenseMap.length === 0 ? (
                  <div className="empty">{tr("noExpenseData")}</div>
                ) : (
                  <Doughnut
                    data={{
                      labels: chartExpenseMap.map(x => x.cat),
                      datasets: [{
                        label: tr("expense"),
                        data: chartExpenseMap.map(x => x.amount),
                      }],
                    }}
                  />
                )}
              </div>
            </div>

            <div className="chart-box wide">
              <div className="card-title">income / expense</div>
              <Bar
                data={{
                  labels: [tr("income"), tr("expense"), tr("balance")],
                  datasets: [{
                    label: fmt("chartDatasetMonth", { year: chartYear, month: chartMonth + 1 }),
                    data: [chartMonthIncome, chartMonthExpense, chartMonthIncome - chartMonthExpense],
                  }],
                }}
              />
            </div>
          </section>
        )}
      </main>
    </div>
    </>
  );
}
