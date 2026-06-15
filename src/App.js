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
  const [savingName, setSavingName] = useState("");
  const [savingTarget, setSavingTarget] = useState("");
  const [savingCurrent, setSavingCurrent] = useState("");

  const [loans, setLoans] = useState([]);
  const [loanName, setLoanName] = useState("");
  const [loanPrincipal, setLoanPrincipal] = useState("");
  const [loanRate, setLoanRate] = useState("");
  const [loanTerm, setLoanTerm] = useState("");
  const [loanStartDate, setLoanStartDate] = useState(new Date().toISOString().slice(0, 10));
  const [loanType, setLoanType] = useState("equal_payment");

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setUser(data.session?.user ?? null);
      setLoading(false);
    });

    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      setUser(session?.user ?? null);
    });

    return () => listener.subscription.unsubscribe();
  }, []);

  const login = async () => {
    if (!email || !password) return alert("이메일과 비밀번호를 입력해주세요");
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) alert(error.message);
  };

  const signup = async () => {
    if (!email || !password) return alert("이메일과 비밀번호를 입력해주세요");
    const { error } = await supabase.auth.signUp({ email, password });
    if (error) return alert(error.message);
    alert("회원가입 완료! 이메일 확인이 필요할 수 있어요.");
  };

  const logout = async () => {
    await supabase.auth.signOut();
    setUser(null);
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

  const saveTx = () => {
    if (!amount) return alert("금액을 입력해주세요");

    const baseTx = {
      id: Date.now(),
      type,
      amount: uncomma(amount),
      name: name || category,
      category,
      repeat,
      date: new Date().toISOString().slice(0, 10),
    };

    const newTxs = createRepeatedTxs(baseTx);

    setTxs([...newTxs, ...txs]);
    setAmount("");
    setName("");
    setRepeat("none");
    setMoneyType("계좌");
    setTransferTo("카드");

    if (repeat === "monthly") alert("매월 반복 내역 12개월치가 추가됐어요.");
    if (repeat === "weekly") alert("매주 반복 내역 8주치가 추가됐어요.");
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
    alert(`${filtered.length}개 내역을 가져왔어요. 중복 ${imported.length - filtered.length}개는 제외했어요.`);
    event.target.value = "";
  };

  const saveBudget = () => {
    if (!budgetAmount) return alert("예산 금액을 입력해주세요");
    const next = budgets.filter(b => b.category !== budgetCat);
    setBudgets([...next, { category: budgetCat, amount: Number(budgetAmount) }]);
    setBudgetAmount("");
  };

  const saveSaving = () => {
    if (!savingName || !savingTarget) return alert("목표명과 목표 금액을 입력해주세요");
    setSavings([...savings, {
      id: Date.now(),
      name: savingName,
      target: Number(savingTarget),
      current: Number(savingCurrent || 0),
    }]);
    setSavingName("");
    setSavingTarget("");
    setSavingCurrent("");
  };

  const saveLoan = () => {
    if (!loanName || !loanPrincipal || !loanRate || !loanTerm) {
      return alert("대출 정보를 모두 입력해주세요");
    }

    setLoans([...loans, {
      id: Date.now(),
      name: loanName,
      principal: Number(loanPrincipal),
      rate: Number(loanRate),
      term: Number(loanTerm),
      startDate: loanStartDate,
      type: loanType,
    }]);

    setLoanName("");
    setLoanPrincipal("");
    setLoanRate("");
    setLoanTerm("");
    setLoanStartDate(new Date().toISOString().slice(0, 10));
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

  const getLoanSummary = (loan) => {
    const paidMonths = getLoanPaidMonths(loan);
    const totalMonths = Number(loan.term);
    const remainingMonths = Math.max(0, totalMonths - paidMonths);
    const monthlyPayment = calcMonthlyPayment(loan);
    const paidAmount = Math.min(Number(loan.principal), monthlyPayment * paidMonths);
    const remainingAmount = Math.max(0, Number(loan.principal) - paidAmount);
    const progress = totalMonths > 0 ? Math.min(100, Math.round((paidMonths / totalMonths) * 100)) : 0;
    const yearsLeft = Math.floor(remainingMonths / 12);
    const monthsLeft = remainingMonths % 12;

    return {
      paidMonths,
      totalMonths,
      remainingMonths,
      monthlyPayment,
      paidAmount,
      remainingAmount,
      progress,
      yearsLeft,
      monthsLeft,
    };
  };

  const catExpenseMap = CATS_EXP.map(cat => ({
    cat,
    amount: realTxs
      .filter(t => t.type === "expense" && t.category === cat)
      .reduce((s,t) => s + Number(t.amount), 0),
  })).filter(x => x.amount > 0);

  const cats = type === "expense" ? CATS_EXP : CATS_INC;

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
    <div className="desktop-wrap">
      <aside className="sidebar">
        <div className="logo">kko<span>money</span></div>

        {[
          ["home", "홈"],
          ["tx", "내역"],
          ["budget", "예산"],
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
                {budgets.length === 0 ? "0%" : Math.round(expense / budgets.reduce((s,b) => s + b.amount, 0) * 100) + "%"}
              </div>
              <div className="banner-row">
                <div className="banner-mini">총 예산<span>{won(budgets.reduce((s,b) => s + b.amount, 0))}</span></div>
                <div className="banner-mini">사용<span>{won(expense)}</span></div>
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

            <section className="card">
              <div className="card-title">저축 목표</div>
              <input className="form-input" placeholder="목표명" value={savingName} onChange={e => setSavingName(e.target.value)} />
              <input className="form-input" placeholder="목표 금액" type="number" value={savingTarget} onChange={e => setSavingTarget(e.target.value)} />
              <input className="form-input" placeholder="현재 금액" type="number" value={savingCurrent} onChange={e => setSavingCurrent(e.target.value)} />
              <button className="btn btn-primary" onClick={saveSaving}>저축 목표 추가</button>

              {savings.map(s => (
                <div className="budget-item" key={s.id}>
                  <div className="budget-top"><b>{s.name}</b><span>{won(s.current)} / {won(s.target)}</span></div>
                  <div className="bar-bg"><div className="bar-fill green" style={{ width: Math.min(100, Math.round(s.current / s.target * 100)) + "%" }} /></div>
                </div>
              ))}
            </section>
          </>
        )}

        {tab === "loan" && (
          <section className="card">
            <div className="card-title">대출 관리</div>

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

            <button className="btn btn-primary" onClick={saveLoan}>대출 추가</button>

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

                    <div className="loan-info-grid">
                      <div>
                        <span>대출 원금</span>
                        <b>{won(l.principal)}</b>
                      </div>
                      <div>
                        <span>누적 상환</span>
                        <b>{won(summary.paidAmount)}</b>
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
  );
}
