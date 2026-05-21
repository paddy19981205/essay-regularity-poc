import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  Archive,
  CheckCircle2,
  Clock,
  Download,
  FileText,
  History,
  Loader2,
  LockKeyhole,
  LogOut,
  Upload,
  XCircle,
} from "lucide-react";
import "./styles.css";

function apiUrl(path) {
  return path;
}

async function fetchJson(path, options) {
  const response = await fetch(apiUrl(path), {
    credentials: "same-origin",
    ...(options || {}),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload.error || `HTTP ${response.status}`);
    error.status = response.status;
    throw error;
  }
  return payload;
}

function formatDate(value) {
  if (!value) return "尚未完成";
  return value.replace("T", " ");
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString("zh-TW");
}

function topItems(items, limit = 8) {
  return (items || []).slice(0, limit);
}

const categoryLabels = {
  "Word Choice": "遣詞用字",
  Grammar: "文法句構",
  Agreement: "一致性",
  Spelling: "拼字",
  Article: "冠詞",
  Tense: "時態",
  Preposition: "介系詞",
  Style: "表達方式",
  Punctuation: "標點",
  Redundancy: "冗詞",
  Capitalization: "大小寫",
  Choice: "選詞",
};

const stageLabels = {
  queued: "排隊中",
  validating_zip: "檢查檔案",
  extracting: "解壓中",
  analyzing: "分析作文",
  writing_outputs: "寫入資料",
  ai_summarizing: "整理正則",
  building_docx: "產出手冊",
  completed: "完成",
  failed: "失敗",
};

const stageMessages = {
  queued: "已收到檔案，等待分析開始。",
  validating_zip: "正在檢查上傳檔案。",
  extracting: "正在解壓批次檔案。",
  analyzing: "正在讀取作文批改內容。",
  writing_outputs: "正在整理本批資料。",
  ai_summarizing: "正在整理本批專屬英語正則。",
  building_docx: "正在產出老師版 Word 手冊。",
  completed: "分析完成。",
};

function progressPercent(progress) {
  if (!progress?.total) return 0;
  return Math.max(0, Math.min(100, Math.round((progress.current / progress.total) * 100)));
}

function progressSummary(run) {
  const progress = run?.progress;
  if (!progress) return run?.status === "queued" ? "等待分析" : "分析中";
  const label = stageLabels[progress.stage] || "分析中";
  if (!progress.total) return label;
  return `${label} · ${progress.current} / ${progress.total} 份 · ${progress.percent ?? progressPercent(progress)}%`;
}

function progressMessage(progress) {
  return stageMessages[progress?.stage] || "正在分析批次內容。";
}

function safeRunError(error) {
  if (!error) return "此批次無法完成，請確認上傳檔案後重新嘗試。";
  if (error.includes("ZIP 內沒有 PDF")) return "ZIP 內沒有可分析的 PDF 檔案。";
  if (error.includes("ZIP 檔案無法讀取")) return "ZIP 檔案無法讀取，請確認壓縮檔完整。";
  if (error.includes("Only .zip")) return "請上傳 ZIP 批次檔。";
  if (error.includes("too large")) return "上傳檔案太大，請改用較小批次或聯絡管理者。";
  return "此批次無法完成，請確認上傳的是含有作文批改 PDF 的 ZIP 檔，或重新上傳。";
}

function App() {
  const [page, setPage] = useState("upload");
  const [runs, setRuns] = useState([]);
  const [activeRunId, setActiveRunId] = useState(null);
  const [pollingId, setPollingId] = useState(null);
  const [session, setSession] = useState({ loading: true, authenticated: false, username: null });
  const activeRun = useMemo(() => runs.find((run) => run.id === activeRunId) || runs[0], [runs, activeRunId]);

  async function refreshRuns() {
    const data = await fetchJson("/api/runs").catch((error) => {
      if (error.status === 401) {
        setSession({ loading: false, authenticated: false, username: null });
      }
      throw error;
    });
    setRuns(data.runs || []);
    if (!activeRunId && data.runs?.length) {
      setActiveRunId(data.runs[0].id);
    }
    return data.runs || [];
  }

  async function refreshSession() {
    const data = await fetchJson("/api/session");
    setSession({
      loading: false,
      authenticated: Boolean(data.authenticated),
      username: data.username || null,
    });
    return data;
  }

  useEffect(() => {
    refreshSession().catch(() => {
      setSession({ loading: false, authenticated: false, username: null });
    });
  }, []);

  useEffect(() => {
    if (session.authenticated) {
      refreshRuns().catch(() => {});
    }
  }, [session.authenticated]);

  useEffect(() => {
    if (!pollingId) return undefined;
    const timer = window.setInterval(async () => {
      const run = await fetchJson(`/api/runs/${pollingId}`).catch(() => null);
      await refreshRuns().catch(() => {});
      if (run && ["completed", "failed"].includes(run.status)) {
        setPollingId(null);
        setActiveRunId(run.id);
        setPage("result");
      }
    }, 1600);
    return () => window.clearInterval(timer);
  }, [pollingId]);

  async function handleCreated(run) {
    setActiveRunId(run.id);
    setPollingId(run.id);
    setPage("result");
    await refreshRuns();
  }

  function openRun(runId) {
    setActiveRunId(runId);
    setPage("result");
  }

  function goUpload() {
    setPage("upload");
  }

  function goHistory() {
    setPage("history");
  }

  async function handleLogin(authenticatedSession) {
    setSession({
      loading: false,
      authenticated: true,
      username: authenticatedSession.username,
    });
    await refreshRuns().catch(() => {});
  }

  async function handleLogout() {
    await fetchJson("/api/session", { method: "DELETE" }).catch(() => {});
    setRuns([]);
    setActiveRunId(null);
    setPollingId(null);
    setPage("upload");
    setSession({ loading: false, authenticated: false, username: null });
  }

  if (session.loading) {
    return <LoadingScreen />;
  }

  if (!session.authenticated) {
    return <LoginPage onLogin={handleLogin} />;
  }

  return (
    <div className="app-shell">
      <aside className="nav-rail">
        <div className="brand">
          <div className="mark" aria-hidden="true">正</div>
          <div>
            <strong>英語正則</strong>
            <span>教師版</span>
          </div>
        </div>
        <nav className="nav-items" aria-label="主要功能">
          <button className={`nav-item ${page === "upload" ? "active" : ""}`} onClick={() => setPage("upload")}>
            <span className="nav-icon"><Upload size={19} /></span>
            <span>上傳</span>
          </button>
          <button className={`nav-item ${page === "result" ? "active" : ""}`} onClick={() => setPage("result")}>
            <span className="nav-icon"><FileText size={19} /></span>
            <span>結果</span>
          </button>
          <button className={`nav-item ${page === "history" ? "active" : ""}`} onClick={() => setPage("history")}>
            <span className="nav-icon"><History size={19} /></span>
            <span>歷史</span>
          </button>
        </nav>
        <div className="nav-footer">
          <div className="signed-in">
            <span>登入帳號</span>
            <strong>{session.username}</strong>
          </div>
          <button className="nav-item logout-button" onClick={handleLogout}>
            <span className="nav-icon"><LogOut size={19} /></span>
            <span>登出</span>
          </button>
        </div>
      </aside>

      <div className="content-shell">
        <header className="top-appbar">
          <div>
            <span>批次產生器</span>
            <strong>英語正則批次產生器</strong>
          </div>
        </header>
        <main className="main">
          {page === "upload" && <UploadPage onCreated={handleCreated} />}
          {page === "result" && <ResultPage run={activeRun} onRefresh={refreshRuns} onUpload={goUpload} onHistory={goHistory} />}
          {page === "history" && <HistoryPage runs={runs} onRefresh={refreshRuns} onOpenRun={openRun} />}
        </main>
      </div>
    </div>
  );
}

function LoadingScreen() {
  return (
    <div className="auth-shell">
      <div className="login-card">
        <Loader2 className="spin" size={24} />
        <strong>正在檢查登入狀態</strong>
      </div>
    </div>
  );
}

function LoginPage({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [isSubmitting, setSubmitting] = useState(false);
  const [error, setError] = useState("");

  async function submit(event) {
    event.preventDefault();
    setSubmitting(true);
    setError("");
    try {
      const data = await fetchJson("/api/session", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username, password }),
      });
      await onLogin(data);
    } catch (err) {
      setError(err.message);
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="auth-shell">
      <form className="login-card elevated-card" onSubmit={submit}>
        <div className="brand login-brand">
          <div className="mark" aria-hidden="true">正</div>
          <div>
            <strong>英語正則</strong>
            <span>教師版</span>
          </div>
        </div>
        <div className="login-title">
          <LockKeyhole size={24} />
          <div>
            <h1>登入</h1>
            <p>請使用管理者提供的帳號密碼。</p>
          </div>
        </div>
        <label className="form-field">
          <span>帳號</span>
          <input value={username} onChange={(event) => setUsername(event.target.value)} autoComplete="username" autoFocus />
        </label>
        <label className="form-field">
          <span>密碼</span>
          <input
            type="password"
            value={password}
            onChange={(event) => setPassword(event.target.value)}
            autoComplete="current-password"
          />
        </label>
        {error && <div className="error-line">{error}</div>}
        <button className="md-button filled" disabled={isSubmitting || !username || !password}>
          {isSubmitting ? <Loader2 className="spin" size={18} /> : <LockKeyhole size={18} />}
          {isSubmitting ? "登入中" : "登入"}
        </button>
      </form>
    </div>
  );
}

function UploadPage({ onCreated }) {
  const inputRef = useRef(null);
  const [file, setFile] = useState(null);
  const [batchName, setBatchName] = useState("");
  const [isDragging, setDragging] = useState(false);
  const [isSubmitting, setSubmitting] = useState(false);
  const [error, setError] = useState("");

  function chooseFile(nextFile) {
    setError("");
    if (!nextFile) return;
    if (!nextFile.name.toLowerCase().endsWith(".zip")) {
      setError("請上傳 .zip 批次檔。");
      return;
    }
    setFile(nextFile);
    if (!batchName) setBatchName(nextFile.name.replace(/\.zip$/i, ""));
  }

  async function submit() {
    if (!file) {
      setError("請先選擇 ZIP 檔。");
      return;
    }
    setSubmitting(true);
    setError("");
    try {
      const run = await fetchJson(`/api/runs?batchName=${encodeURIComponent(batchName || file.name.replace(/\.zip$/i, ""))}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/zip",
          "X-File-Name": encodeURIComponent(file.name),
        },
        body: file,
      });
      await onCreated(run);
    } catch (err) {
      setError(err.message);
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <section className="page upload-page">
      <header className="page-header">
        <span>上傳批次</span>
        <h1>上傳批改結果 ZIP，產出英語正則手冊</h1>
        <p>系統會讀取批改結果中的作文內容，完成後提供老師版 Word 手冊下載。</p>
      </header>

      <div
        className={`dropzone ${isDragging ? "dragging" : ""}`}
        onDragOver={(event) => {
          event.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={(event) => {
          event.preventDefault();
          setDragging(false);
          chooseFile(event.dataTransfer.files[0]);
        }}
        onClick={() => inputRef.current?.click()}
      >
        <span className="dropzone-icon"><Archive size={34} /></span>
        <strong>{file ? file.name : "拖放 ZIP 到這裡，或點擊選擇檔案"}</strong>
        <span>{file ? `${Math.round(file.size / 1024 / 1024)} MB` : "請將批改結果資料夾壓縮成 ZIP 後上傳"}</span>
        <input ref={inputRef} type="file" accept=".zip,application/zip" onChange={(event) => chooseFile(event.target.files[0])} />
      </div>

      <div className="action-card">
        <label className="form-field">
          <span>批次名稱</span>
          <input value={batchName} onChange={(event) => setBatchName(event.target.value)} placeholder="例如：站前模D 2026-05-05" />
        </label>

        {error && <div className="error-line">{error}</div>}

        <button className="md-button filled" disabled={isSubmitting || !file} onClick={submit}>
          {isSubmitting ? <Loader2 className="spin" size={18} /> : <Upload size={18} />}
          {isSubmitting ? "上傳中" : "開始分析"}
        </button>
      </div>
    </section>
  );
}

function ResultPage({ run, onRefresh, onUpload, onHistory }) {
  useEffect(() => {
    if (run?.status === "running" || run?.status === "queued") {
      const timer = window.setInterval(() => onRefresh().catch(() => {}), 1500);
      return () => window.clearInterval(timer);
    }
    return undefined;
  }, [run?.id, run?.status]);

  if (!run) {
    return (
      <EmptyState title="尚無分析結果" body="請先上傳 ZIP 批次檔，完成後結果會出現在這裡。" />
    );
  }

  if (run.status === "failed") {
    return (
      <section className="page">
        <header className="page-header compact">
          <span>分析失敗</span>
          <h1>{run.batchName}</h1>
        </header>
        <div className="status-box failed">
          <XCircle size={22} />
          <p>{safeRunError(run.error)}</p>
        </div>
      </section>
    );
  }

  if (run.status !== "completed") {
    const progress = run.progress || {};
    const progressValue = progress.percent ?? progressPercent(progress);
    return (
      <section className="page">
        <header className="page-header compact">
          <span>分析中</span>
          <h1>{run.batchName}</h1>
          <p>大批次可先留在此頁觀察進度；歷史頁也會同步保留目前狀態。</p>
        </header>
        <div className="progress-card">
          <div className="progress-head">
            <div>
              <span>{stageLabels[progress.stage] || (run.status === "queued" ? "排隊中" : "分析中")}</span>
              <strong>{progressMessage(progress)}</strong>
            </div>
            <span className="progress-icon"><Loader2 className="spin" size={22} /></span>
          </div>

          <div className="progress-meter" aria-label="分析進度">
            <i style={{ width: `${progressValue}%` }} />
          </div>

          <div className="progress-meta">
            <span>{progress.total ? `已處理 ${progress.current} / ${progress.total} 份` : "正在準備批次"}</span>
            <strong>{progress.total ? `${progressValue}%` : "..."}</strong>
          </div>
        </div>
      </section>
    );
  }

  const summary = run.summary;
  const effectiveEssays = Math.max(0, Number(summary?.pdfs || 0) - Number(summary?.zeroScore || 0));
  const categories = topItems(summary?.categories, 8).map((item) => ({
    ...item,
    name: categoryLabels[item.name] || item.name,
  }));
  const scoreBands = summary?.scoreBands || [];
  const maxCategory = Math.max(1, ...categories.map((item) => item.count || 0));
  const maxScoreBand = Math.max(1, ...scoreBands.map((item) => item.count || 0));
  return (
    <section className="page result-page">
      <section className="completion-card">
        <div className="completion-copy">
          <span className="completion-icon"><CheckCircle2 size={24} /></span>
          <div>
            <span className="section-label">分析結果</span>
            <h1>手冊已產出</h1>
            <p>{run.batchName}</p>
            <div className="completion-meta">完成時間：{formatDate(run.completedAt)}</div>
          </div>
        </div>
        <div className="completion-actions">
          <a className="md-button filled primary-download" href={`/api/runs/${run.id}/files/docx`}>
            <Download size={18} />
            下載老師版 Word
          </a>
          <button className="md-button tonal" onClick={onUpload}>
            <Upload size={18} />
            上傳新批次
          </button>
          <button className="md-button text-button" onClick={onHistory}>
            <History size={18} />
            查看歷史
          </button>
        </div>
      </section>

      {summary && (
        <section className="overview-section">
          <div className="section-heading">
            <h2>本批概況</h2>
            <p>數字僅供老師檢查本次分析結果是否合理。</p>
          </div>
          <div className="metrics">
            <Metric label="作文份數" value={formatNumber(summary.pdfs)} />
            <Metric label="有效作文" value={formatNumber(effectiveEssays)} />
            <Metric label="0 分 / 未作答" value={formatNumber(summary.zeroScore)} />
            <Metric label="正則條目" value={formatNumber(summary.manualEntries)} />
          </div>
          <div className="run-meta">
            <span>平均分數：{summary.averageScore ?? "-"} / 20</span>
            <span>中位數：{summary.medianScore ?? "-"}</span>
          </div>
        </section>
      )}

      <ManualPreview preview={run.manualPreview} />

      <div className="supporting-grid">
        <section className="panel">
          <div className="panel-head">
            <h2>常見錯誤類型</h2>
          </div>
          <BarList items={categories} max={maxCategory} />
        </section>

        <section className="panel">
          <div className="panel-head">
            <h2>分數分布</h2>
          </div>
          <BarList items={scoreBands.map((item) => ({ name: item.label, count: item.count }))} max={maxScoreBand} />
        </section>
      </div>
    </section>
  );
}

function ManualPreview({ preview }) {
  const units = preview?.units || [];
  return (
    <section className="manual-preview panel">
      <div className="panel-head preview-head">
        <div>
          <h2>手冊內容摘要</h2>
          <p>僅顯示單元與正則標題，完整內容請下載 Word 查看。</p>
        </div>
      </div>
      {units.length === 0 ? (
        <div className="empty-panel">此批次尚未提供手冊摘要，可直接下載 Word 查看完整內容。</div>
      ) : (
        <div className="unit-list">
          {units.map((unit) => (
            <article className="unit-item" key={unit.title}>
              <div className="unit-title-row">
                <h3>{unit.title}</h3>
                <span>{formatNumber(unit.entryCount)} 條</span>
              </div>
              {unit.sampleTitles?.length > 0 && (
                <ul>
                  {unit.sampleTitles.map((title) => (
                    <li key={title}>{title}</li>
                  ))}
                </ul>
              )}
            </article>
          ))}
        </div>
      )}
    </section>
  );
}

function HistoryPage({ runs, onRefresh, onOpenRun }) {
  useEffect(() => {
    onRefresh().catch(() => {});
  }, []);

  return (
    <section className="page">
      <header className="page-header compact">
        <span>歷史結果</span>
        <h1>過去分析批次</h1>
        <p>可回看過去完成的分析批次，並重新下載老師版 Word 手冊。</p>
      </header>

      {runs.length === 0 ? (
        <EmptyState title="尚無歷史紀錄" body="完成第一次 ZIP 分析後，這裡會出現可回看的結果。" />
      ) : (
        <div className="history-list">
          {runs.map((run) => (
            <div key={run.id} className="history-row">
              <button className="history-open" onClick={() => onOpenRun(run.id)}>
                <strong>{run.batchName}</strong>
                <span>{formatDate(run.createdAt)}</span>
              </button>
              <Status status={run.status} />
              <div className="history-summary">
                {run.status === "completed" ? (
                  <>
                    <div className="history-meta">{historySummary(run.summary)}</div>
                    <a href={`/api/runs/${run.id}/files/docx`} onClick={(event) => event.stopPropagation()}>
                      下載老師版 Word
                    </a>
                  </>
                ) : (
                  run.status === "failed" ? safeRunError(run.error) : progressSummary(run)
                )}
              </div>
            </div>
          ))}
        </div>
      )}
    </section>
  );
}

function Metric({ label, value }) {
  return (
    <div className="metric">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function BarList({ items, max }) {
  if (!items?.length) {
    return <div className="empty-panel">沒有可顯示的統計資料。</div>;
  }
  return (
    <div className="bar-list">
      {items.map((item) => (
        <div className="bar-row" key={item.name}>
          <div>
            <span>{item.name}</span>
            <strong>{formatNumber(item.count)}</strong>
          </div>
          <div className="bar-track">
            <i style={{ width: `${Math.max(3, Math.round(((item.count || 0) / max) * 100))}%` }} />
          </div>
        </div>
      ))}
    </div>
  );
}

function historySummary(summary) {
  if (!summary) return "尚無統計";
  return `${formatNumber(summary.pdfs)} 份作文 · ${formatNumber(summary.manualEntries)} 條正則`;
}

function Status({ status }) {
  const label = {
    queued: "排隊中",
    running: "分析中",
    completed: "完成",
    failed: "失敗",
  }[status] || status;
  const icon = status === "completed" ? <CheckCircle2 size={14} /> : null;
  return <span className={`status ${status}`}>{icon}{label}</span>;
}

function EmptyState({ title, body }) {
  return (
    <section className="page">
      <div className="empty">
        <Clock size={28} />
        <h1>{title}</h1>
        <p>{body}</p>
      </div>
    </section>
  );
}

createRoot(document.getElementById("root")).render(<App />);
