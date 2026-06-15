import { useEffect, useState } from "react";
import { supabase } from "./lib/supabase";
import * as XLSX from "xlsx";
import {
  Chart as ChartJS,
  ArcElement,
  BarElement,
  CategoryScale,
  LinearScale,
  Tooltip,
  Legend,
} from "chart.js";
import { Doughnut, Bar } from "react-chartjs-2";

ChartJS.register(ArcElement, BarElement, CategoryScale, LinearScale, Tooltip, Legend);

const CATS_EXP = ["식비","카페","교통","쇼핑","의료","문화/여가","구독","통신","저축","대출상환","기타"];
const CATS_INC = ["이월","월급","용돈","부업","이자","환급","기타수입"];

function won(n) {
  return Math.round(Number(n || 0)).toLocaleString("ko-KR") + "원";
}

function comma(n) {
  return String(n || "").replace(/\D/g, "").replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function uncomma(n) {
  return Number(String(n || "0").replaceAll(",", ""));
}

const DEFAULT_MONEY_TYPES = ["계좌", "카드", "저축", "투자"];

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

  const [tab, setTab] = useState("home");
  const [chartYear, setChartYear] = useState(new Date().getFullYear());
  const [chartMonth, setChartMonth] = useState(new Date().getMonth());
  const [txs, setTxs] = useState([]);

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
    if (!email || !password) return alert("이메일과 비밀번호를 입력해주세요");
    const { data, error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) return alert(error.message);
    if (data.user) await loadUserData(data.user);
  };

  const signup = async () => {
    if (!email || !password) return alert("이메일과 비밀번호를 입력해주세요");
    const { error } = await supabase.auth.signUp({ email, password });
    if (error) return alert(error.message);
    showPopup("회원가입 완료! 이메일 확인이 필요할 수 있어요.", "회원가입 완료", "💌");
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
    if (!amount) return alert("금액을 입력해주세요");

    const finalType = moneyType !== transferTo && category === "이체" ? "transfer" : type;

    const baseTx = {
      id: Date.now(),
      type: finalType,
      amount: uncomma(amount),
      name: name || category,
      category,
      moneyType,
      transferTo: finalType === "transfer" ? transferTo : "",
      repeat,
      date: new Date().toISOString().slice(0, 10),
    };

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

    if (repeat === "monthly") showPopup("매월 반복 내역 12개월치가 추가됐어요.", "반복 등록 완료", "🔁");
    if (repeat === "weekly") showPopup("매주 반복 내역 8주치가 추가됐어요.", "반복 등록 완료", "🔁");
  };

  const handleExcelUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    if (!rows.length) {
      alert("엑셀에서 데이터를 찾지 못했어요.");
      return;
    }

    const imported = rows.map((row, index) => {
      const keys = Object.keys(row);

      const dateKey = keys.find(k => k.includes("일자") || k.includes("날짜") || k.toLowerCase().includes("date")) || keys[0];
      const nameKey = keys.find(k => k.includes("내용") || k.includes("적요") || k.includes("거래") || k.includes("메모")) || keys[1];
      const amountKey = keys.find(k => k.includes("금액") || k.includes("출금") || k.includes("입금") || k.includes("amount")) || keys[2];

      const rawName = row[nameKey] || "엑셀 가져오기";
      const rawAmount = normalizeAmount(row[amountKey]);

      let txType = "expense";
      let finalAmount = Math.abs(rawAmount);

      const rowText = JSON.stringify(row);

      if (rowText.includes("입금") || rowText.includes("수입") || rawAmount > 0) {
        txType = "income";
      }

      if (rowText.includes("출금") || rowText.includes("지출") || rawAmount < 0) {
        txType = "expense";
      }

      return {
        id: Date.now() + index,
        type: txType,
        amount: finalAmount,
        name: String(rawName),
        category: guessCategory(rawName, txType),
        moneyType: "계좌",
        transferTo: "",
        repeat: "none",
        date: normalizeDate(row[dateKey]),
      };
    }).filter(t => t.amount > 0);

    const existingKeys = new Set(
      txs.map(t => `${t.date}-${t.name}-${t.amount}-${t.type}`)
    );

    const filtered = imported.filter(t => {
      const key = `${t.date}-${t.name}-${t.amount}-${t.type}`;
      return !existingKeys.has(key);
    });

    setTxs([...filtered, ...txs]);
    
    event.target.value = "";
  };

  const saveBudget = async () => {
    if (!budgetAmount) return alert("예산 금액을 입력해주세요");

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
    if (!savingName || !savingTarget) return alert("목표명과 목표 금액을 입력해주세요");

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
      showPopup("저축 목표가 수정됐어요.", "수정 완료", "🐷");
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
      showPopup("저축 목표가 추가됐어요.", "저축 목표 추가", "🐷");
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
    if (!window.confirm("이 저축 목표를 삭제할까요?")) return;

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
      return alert("대출 정보를 모두 입력해주세요");
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
      showPopup("대출건이 수정됐어요.", "수정 완료", "🏦");
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
      showPopup("대출건이 추가됐어요. 입출금 카테고리에 대출명이 자동으로 표시돼요.", "대출 추가 완료", "🏦");
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
    if (!window.confirm("이 대출건을 삭제할까요?")) return;

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
    const yearsLeft = Math.floor(remainingMonths / 12);
    const monthsLeft = remainingMonths % 12;

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

  if (loading) return <div className="auth-screen">로딩 중...</div>;

  if (!user) {
    return (
      <div className="auth-screen">
        <div className="auth-logo">kko<span>money</span></div>
        <div className="auth-sub">내 돈을 귀엽게 관리해요 💗</div>

        <div className="auth-card">
          <div className="auth-title">로그인</div>

          <div className="form-group">
            <label className="form-label">이메일</label>
            <input className="form-input" value={email} onChange={e => setEmail(e.target.value)} />
          </div>

          <div className="form-group">
            <label className="form-label">비밀번호</label>
            <input className="form-input" type="password" value={password} onChange={e => setPassword(e.target.value)} />
          </div>

          <button className="auth-btn" onClick={login}>로그인</button>
          <button className="auth-btn secondary" onClick={signup}>회원가입</button>
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
          <button className="kko-popup-btn" onClick={closePopup}>확인</button>
        </div>
      </div>
    )}

    <div className="desktop-wrap">
      <aside className="sidebar">
        <div className="logo">kko<span>money</span></div>

        {[
          ["home", "홈"],
          ["tx", "내역"],
          ["budget", "예산"],
          ["saving", "저축"],
          ["loan", "대출"],
          ["chart", "분석"],
          ["settings", "설정"],
        ].map(([key, label]) => (
          <button key={key} className={`nav-btn ${tab === key ? "active" : ""}`} onClick={() => setTab(key)}>
            {label}
          </button>
        ))}

        <button className="nav-btn" onClick={logout}>로그아웃</button>
      </aside>

      <main className="main-content">
        {tab === "home" && (
          <>
            <section className="banner">
              <div className="banner-lbl">이번 달 남은 예산</div>
              <div className="banner-main">{won(income - expense)}</div>
              <div className="banner-row">
                <div className="banner-mini">수입<span>{won(income)}</span></div>
                <div className="banner-mini">지출<span>{won(expense)}</span></div>
                <div className="banner-mini">이체<span>합계 제외</span></div>
              </div>
            </section>

            <section className="card">
              <div className="card-title">빠른 입력</div>

              <div className="type-toggle">
                <button className={`type-btn ${type === "expense" ? "active-expense" : ""}`} onClick={() => {setType("expense"); setCategory("식비");}}>
                  지출 💸
                </button>
                <button className={`type-btn ${type === "income" ? "active-income" : ""}`} onClick={() => {setType("income"); setCategory("월급");}}>
                  수입 💰
                </button>
              </div>

              <div className="form-group">
                <label className="form-label">금액</label>
                <input className="form-input" type="text" value={amount} onChange={e => setAmount(comma(e.target.value))} />
              </div>

              <div className="form-group">
                <label className="form-label">내용</label>
                <input className="form-input" value={name} onChange={e => setName(e.target.value)} />
              </div>

              <div className="form-group">
                <label className="form-label">카테고리</label>
                <select className="form-select" value={category} onChange={e => setCategory(e.target.value)}>
                  {cats.map(c => <option key={c}>{c}</option>)}
                  <option value="이체">이체</option>
                </select>
              </div>

              <div className="form-group">
                <label className="form-label">출입금 유형</label>
                <select className="form-select" value={moneyType} onChange={e => setMoneyType(e.target.value)}>
                  {moneyTypes.map(mt => <option key={mt}>{mt}</option>)}
                </select>
              </div>

              {category === "이체" && (
                <div className="form-group">
                  <label className="form-label">이체 도착 유형</label>
                  <select className="form-select" value={transferTo} onChange={e => setTransferTo(e.target.value)}>
                    {moneyTypes.map(mt => <option key={mt}>{mt}</option>)}
                  </select>
                  <div className="hint">계좌/카드/저축/투자 간 이동은 수입·지출 합계에 포함되지 않아요.</div>
                </div>
              )}

              <div className="form-group">
                <label className="form-label">반복</label>
                <select className="form-select" value={repeat} onChange={e => setRepeat(e.target.value)}>
                  <option value="none">반복 없음</option>
                  <option value="monthly">매월 반복</option>
                  <option value="weekly">매주 반복</option>
                </select>
              </div>

              <button className="btn btn-primary" onClick={saveTx}>저장</button>
            </section>

            <section className="card">
              <div className="card-title">유형별 출입금 요약</div>
              <div className="grid4">
                {typeSummary.map(row => (
                  <div className="sum-card" key={row.type}>
                    <div className="sum-lbl">{row.type}</div>
                    <div className="sum-val income">+{won(row.income + row.transferIn)}</div>
                    <div className="sum-val expense">-{won(row.expense + row.transferOut)}</div>
                  </div>
                ))}
              </div>
              <div className="hint">이체는 유형별 이동에는 표시되지만 전체 수입/지출에는 계산되지 않아요.</div>
            </section>

            <section className="card">
              <div className="card-title">엑셀/CSV 가져오기</div>
              <input className="form-input" type="file" accept=".xlsx,.xls,.csv" onChange={handleExcelUpload} />
              <div className="hint">날짜, 내용, 금액이 있는 엑셀 파일을 넣으면 자동으로 내역에 추가돼요.</div>
            </section>
          </>
        )}

        {tab === "tx" && (
          <section className="card">
            <div className="card-title">거래 내역</div>

            <input className="form-input" type="file" accept=".xlsx,.xls,.csv" onChange={handleExcelUpload} />

            <div className="grid3">
              <div className="sum-card"><div className="sum-lbl">수입</div><div className="sum-val income">{won(income)}</div></div>
              <div className="sum-card"><div className="sum-lbl">지출</div><div className="sum-val expense">{won(expense)}</div></div>
              <div className="sum-card"><div className="sum-lbl">잔액</div><div className="sum-val">{won(income - expense)}</div></div>
            </div>

            <div className="tx-list">
              {txs.length === 0 ? <div className="empty">아직 내역이 없어요</div> : txs.map(t => (
                <div className="tx-row" key={t.id}>
                  <div className="tx-info">
                    <div className="tx-name">{t.name}</div>
                    <div className="tx-meta">
                      {t.date} · {t.category} · {t.moneyType || "계좌"}
                      {t.type === "transfer" ? ` → ${t.transferTo}` : ""}
                      {t.repeat !== "none" ? ` · ${t.repeat === "monthly" ? "매월 반복" : "매주 반복"}` : ""}
                    </div>
                  </div>
                  <div className={`tx-amt ${t.type}`}>{t.type === "transfer" ? "이체 " : t.type === "income" ? "+" : "-"}{won(t.amount)}</div>
                </div>
              ))}
            </div>
          </section>
        )}

        {tab === "budget" && (
          <>
            <section className="banner">
              <div className="banner-lbl">예산 사용률</div>
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
                <div className="banner-mini">총 예산<span>{won(budgets.reduce((s,b) => s + b.amount, 0))}</span></div>
                <div className="banner-mini">사용<span>{won(budgets.reduce((sum,b) => {
                  return sum + txs
                    .filter(t => t.type === "expense" && t.category === b.category)
                    .reduce((a,t) => a + Number(t.amount), 0);
                }, 0))}</span></div>
              </div>
            </section>

            <section className="card">
              <div className="card-title">카테고리별 예산</div>
              <select className="form-select" value={budgetCat} onChange={e => setBudgetCat(e.target.value)}>
                {CATS_EXP.map(c => <option key={c}>{c}</option>)}
              </select>
              <input className="form-input" type="number" placeholder="월 예산" value={budgetAmount} onChange={e => setBudgetAmount(e.target.value)} />
              <button className="btn btn-primary" onClick={saveBudget}>예산 저장</button>

              <div className="budget-list">
                {budgets.map(b => {
                  const used = txs.filter(t => t.type === "expense" && t.category === b.category).reduce((s,t) => s + Number(t.amount), 0);
                  const pct = Math.min(100, Math.round(used / b.amount * 100));
                  return (
                    <div className="budget-item" key={b.category}>
                      <div className="budget-top"><b>{b.category}</b><span>{won(used)} / {won(b.amount)}</span></div>
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
              <div className="banner-lbl">저축 목표 달성률</div>
              <div className="banner-main">{totalSavingPct}%</div>
              <div className="banner-row">
                <div className="banner-mini">목표 합계<span>{won(totalSavingTarget)}</span></div>
                <div className="banner-mini">현재 합계<span>{won(totalSavingCurrent)}</span></div>
                <div className="banner-mini">남은 금액<span>{won(Math.max(0, totalSavingTarget - totalSavingCurrent))}</span></div>
              </div>
            </section>

            <section className="card">
              <div className="card-title">{editingSavingId ? "저축 목표 수정" : "저축 목표 추가"}</div>

              <input
                className="form-input"
                placeholder="목표명 예: 비상금, 여행 적금"
                value={savingName}
                onChange={e => setSavingName(e.target.value)}
              />
              <input
                className="form-input"
                placeholder="목표 금액"
                type="number"
                value={savingTarget}
                onChange={e => setSavingTarget(e.target.value)}
              />
              <input
                className="form-input"
                placeholder="현재 모은 금액"
                type="number"
                value={savingCurrent}
                onChange={e => setSavingCurrent(e.target.value)}
              />

              <div className="saving-button-row">
                <button className="btn btn-primary" onClick={saveSaving}>
                  {editingSavingId ? "수정 저장" : "저축 목표 추가"}
                </button>
                {editingSavingId && (
                  <button className="btn btn-secondary" onClick={cancelSavingEdit}>
                    취소
                  </button>
                )}
              </div>
            </section>

            <section className="card">
              <div className="card-title">저축 현황</div>

              {savings.length === 0 ? (
                <div className="empty">저축 목표가 없어요</div>
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
                          <span>남은 금액</span>
                          <b>{won(left)}</b>
                        </div>

                        <div className="saving-actions">
                          <button className="btn btn-secondary" onClick={() => editSaving(sv)}>수정</button>
                          <button className="btn btn-danger" onClick={() => deleteSaving(sv.id)}>삭제</button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </section>
          </>
        )}

        {tab === "loan" && (
          <section className="card">
            <div className="card-title">{editingLoanId ? "대출 수정" : "대출 관리"}</div>

            <input className="form-input" placeholder="대출명" value={loanName} onChange={e => setLoanName(e.target.value)} />
            <input className="form-input" placeholder="대출 원금" type="number" value={loanPrincipal} onChange={e => setLoanPrincipal(e.target.value)} />
            <input className="form-input" placeholder="연이율 %" type="number" value={loanRate} onChange={e => setLoanRate(e.target.value)} />
            <input className="form-input" placeholder="기간 개월" type="number" value={loanTerm} onChange={e => setLoanTerm(e.target.value)} />

            <label className="form-label loan-date-label">대출 시작일</label>
            <input className="form-input" type="date" value={loanStartDate} onChange={e => setLoanStartDate(e.target.value)} />

            <select className="form-select" value={loanType} onChange={e => setLoanType(e.target.value)}>
              <option value="equal_payment">원리금균등</option>
              <option value="equal_principal">원금균등</option>
            </select>

            <div className="loan-button-row">
              <button className="btn btn-primary" onClick={saveLoan}>
                {editingLoanId ? "수정 저장" : "대출 추가"}
              </button>
              {editingLoanId && (
                <button className="btn btn-secondary" onClick={cancelLoanEdit}>
                  취소
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
                        <p>{l.type === "equal_payment" ? "원리금균등" : "원금균등"} · 연 {l.rate}% · {l.term}개월</p>
                      </div>
                      <div className="loan-progress-num">{summary.progress}%</div>
                    </div>

                    <div className="loan-progress-bg">
                      <div className="loan-progress-fill" style={{ width: summary.progress + "%" }} />
                    </div>

                    <div className="loan-info-grid loan-info-grid-auto">
                      <div>
                        <span>대출 원금</span>
                        <b>{won(l.principal)}</b>
                      </div>
                      <div>
                        <span>누적 상환</span>
                        <b>{won(summary.paidAmount)}</b>
                      </div>
                      <div>
                        <span>입출금 반영</span>
                        <b>{won(getLoanTxAmount(l))}</b>
                      </div>
                      <div>
                        <span>남은 잔금</span>
                        <b>{won(summary.remainingAmount)}</b>
                      </div>
                      <div>
                        <span>월 상환액</span>
                        <b>{won(summary.monthlyPayment)}</b>
                      </div>
                    </div>

                    <div className="loan-detail">
                      <div>총 {summary.totalMonths}개월 중 <b>{summary.paidMonths}개월치</b> 갚았어요.</div>
                      <div>
                        앞으로 <b>{summary.remainingMonths}개월</b>
                        {summary.yearsLeft > 0 && <> · 약 <b>{summary.yearsLeft}년 {summary.monthsLeft}개월</b></>}
                        남았어요.
                      </div>
                      <div>시작일: {l.startDate || "미입력"}</div>
                    </div>

                    <div className="loan-card-actions">
                      <button className="btn btn-secondary" onClick={() => editLoan(l)}>수정</button>
                      <button className="btn btn-danger" onClick={() => deleteLoan(l.id)}>삭제</button>
                    </div>
                  </div>
                );
              })}
            </div>
          </section>
        )}

        {tab === "settings" && (
          <section className="card">
            <div className="card-title">설정 · 출입금 유형 편집</div>

            <div className="form-group">
              <label className="form-label">새 유형 추가</label>
              <div className="settings-row">
                <input
                  className="form-input"
                  value={newMoneyType}
                  onChange={e => setNewMoneyType(e.target.value)}
                  placeholder="예: 비상금, CMA, 현금"
                />
                <button
                  className="btn btn-primary"
                  onClick={() => {
                    const v = newMoneyType.trim();
                    if (!v) return alert("추가할 유형명을 입력해주세요");
                    if (moneyTypes.includes(v)) return alert("이미 있는 유형이에요");
                    const next = [...moneyTypes, v];
                    setMoneyTypes(next);
                    localStorage.setItem("moneyTypes", JSON.stringify(next));
                    setNewMoneyType("");
                  }}
                >
                  추가
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
                      if (moneyTypes.length <= 1) return alert("유형은 최소 1개 필요해요");
                      const next = moneyTypes.filter(x => x !== mt);
                      setMoneyTypes(next);
                      localStorage.setItem("moneyTypes", JSON.stringify(next));
                      if (moneyType === mt) setMoneyType(next[0]);
                      if (transferTo === mt) setTransferTo(next[0]);
                    }}
                  >
                    삭제
                  </button>
                </div>
              ))}
            </div>

            <div className="hint">계좌, 카드, 저축, 투자 외에 비상금/CMA/현금 같은 유형을 추가할 수 있어요.</div>
          </section>
        )}

        {tab === "chart" && (
          <section className="card">
            <div className="chart-month-head">
              <button className="month-arrow" onClick={() => moveChartMonth(-1)}>‹</button>
              <div>
                <div className="card-title">분석 차트</div>
                <div className="chart-month-title">{chartYear}년 {chartMonth + 1}월</div>
              </div>
              <button className="month-arrow" onClick={() => moveChartMonth(1)}>›</button>
            </div>

            <div className="grid3">
              <div className="sum-card">
                <div className="sum-lbl">월 수입</div>
                <div className="sum-val income">{won(chartMonthIncome)}</div>
              </div>
              <div className="sum-card">
                <div className="sum-lbl">월 지출</div>
                <div className="sum-val expense">{won(chartMonthExpense)}</div>
              </div>
              <div className="sum-card">
                <div className="sum-lbl">월 잔액</div>
                <div className="sum-val">{won(chartMonthIncome - chartMonthExpense)}</div>
              </div>
            </div>

            <div className="chart-grid">
              <div className="chart-box">
                <div className="card-title">수입 카테고리</div>
                {chartIncomeMap.length === 0 ? (
                  <div className="empty">이 달의 수입 데이터가 없어요</div>
                ) : (
                  <Doughnut
                    data={{
                      labels: chartIncomeMap.map(x => x.cat),
                      datasets: [{
                        label: "수입",
                        data: chartIncomeMap.map(x => x.amount),
                      }],
                    }}
                  />
                )}
              </div>

              <div className="chart-box">
                <div className="card-title">지출 카테고리</div>
                {chartExpenseMap.length === 0 ? (
                  <div className="empty">이 달의 지출 데이터가 없어요</div>
                ) : (
                  <Doughnut
                    data={{
                      labels: chartExpenseMap.map(x => x.cat),
                      datasets: [{
                        label: "지출",
                        data: chartExpenseMap.map(x => x.amount),
                      }],
                    }}
                  />
                )}
              </div>
            </div>

            <div className="chart-box wide">
              <div className="card-title">수입 / 지출 비교</div>
              <Bar
                data={{
                  labels: ["수입", "지출", "잔액"],
                  datasets: [{
                    label: `${chartYear}년 ${chartMonth + 1}월`,
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
