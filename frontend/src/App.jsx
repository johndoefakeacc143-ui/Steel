import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import axios from "axios";

const API_BASE = import.meta.env.VITE_API_URL || "";

const TABS = [
  { id: "beams", label: "Beams", key: "beams" },
  { id: "columns", label: "Columns", key: "columns" },
  { id: "bracings", label: "Bracing", key: "bracings" },
  { id: "base_plates", label: "Base Plates", key: "base_plates" },
  { id: "summary", label: "Summary", key: "summary" },
];

function formatBytes(bytes) {
  if (!bytes && bytes !== 0) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function downloadBase64Excel(base64, filename) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename || "SteelDraw_Extract.xlsx";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function BrandMark({ className = "" }) {
  return (
    <div className={`flex items-center gap-3 ${className}`}>
      <div className="relative flex h-11 w-11 items-center justify-center bg-steel-900 text-steel-50 shadow-[4px_4px_0_#c45f22]">
        <svg viewBox="0 0 32 32" className="h-6 w-6" aria-hidden>
          <path
            fill="currentColor"
            d="M6 24V8h4l5 8 5-8h4v16h-3.5V14L16.5 22h-1L9.5 14v10H6z"
          />
          <rect x="5" y="26" width="22" height="2" fill="#c45f22" />
        </svg>
      </div>
      <div className="leading-none">
        <p className="font-display text-3xl tracking-[0.04em] text-steel-900 sm:text-4xl">
          SteelDraw
        </p>
        <p className="mt-1 font-mono text-[11px] uppercase tracking-[0.28em] text-ember-500">
          AI Extractor
        </p>
      </div>
    </div>
  );
}

function ErrorBanner({ message, onDismiss }) {
  if (!message) return null;
  return (
    <div
      className="animate-fade-up border-l-4 border-ember-500 bg-white/90 px-4 py-3 text-sm text-steel-800 shadow-sm"
      role="alert"
    >
      <div className="flex items-start justify-between gap-3">
        <div>
          <p className="font-semibold text-ember-600">Something went wrong</p>
          <p className="mt-1 text-steel-700">{message}</p>
        </div>
        {onDismiss && (
          <button
            type="button"
            onClick={onDismiss}
            className="font-mono text-xs uppercase tracking-wider text-steel-500 hover:text-steel-900"
          >
            Dismiss
          </button>
        )}
      </div>
    </div>
  );
}

function UploadScreen({ onFileReady, error, setError }) {
  const inputRef = useRef(null);
  const [dragging, setDragging] = useState(false);
  const [file, setFile] = useState(null);

  const acceptFile = useCallback(
    (f) => {
      setError("");
      if (!f) return;
      if (f.type !== "application/pdf" && !f.name.toLowerCase().endsWith(".pdf")) {
        setError("Please upload a PDF drawing file.");
        return;
      }
      if (f.size > 500 * 1024 * 1024) {
        setError("File exceeds the 500 MB limit.");
        return;
      }
      setFile(f);
    },
    [setError]
  );

  const onDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    const f = e.dataTransfer.files?.[0];
    acceptFile(f);
  };

  return (
    <section className="mx-auto w-full max-w-3xl animate-fade-up">
      <BrandMark className="mb-8" />
      <h1 className="max-w-xl font-display text-5xl leading-[0.95] tracking-wide text-steel-900 sm:text-6xl">
        Read steel drawings like a detailer.
      </h1>
      <p className="mt-4 max-w-lg text-lg text-steel-600">
        Drop a structural PDF. Extract beams, columns, bracing, and base plates
        into Excel — Mark, Length/Height, Quantity.
      </p>

      <div
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={onDrop}
        onClick={() => inputRef.current?.click()}
        className={`mt-10 cursor-pointer border-2 border-dashed px-6 py-14 text-center transition-colors ${
          dragging
            ? "border-ember-500 bg-ember-500/10"
            : "border-steel-400/70 bg-white/55 hover:border-steel-600 hover:bg-white/80"
        }`}
      >
        <input
          ref={inputRef}
          type="file"
          accept="application/pdf,.pdf"
          className="hidden"
          onChange={(e) => acceptFile(e.target.files?.[0])}
        />
        <p className="font-mono text-xs uppercase tracking-[0.25em] text-steel-500">
          Drag & drop PDF
        </p>
        <p className="mt-3 text-steel-700">
          or click to browse — up to 500 MB, processed page by page
        </p>
      </div>

      {file && (
        <div className="mt-5 flex flex-col gap-4 border border-steel-200 bg-white/80 p-4 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <p className="font-semibold text-steel-900">{file.name}</p>
            <p className="font-mono text-sm text-steel-500">{formatBytes(file.size)}</p>
          </div>
          <button
            type="button"
            onClick={() => onFileReady(file)}
            className="bg-steel-900 px-6 py-3 font-mono text-xs uppercase tracking-[0.2em] text-white transition hover:bg-ember-600"
          >
            Start Extraction
          </button>
        </div>
      )}

      <div className="mt-6">
        <ErrorBanner message={error} onDismiss={() => setError("")} />
      </div>
    </section>
  );
}

function PageSelectScreen({ meta, onSubmit, onCancel, error, setError, busy }) {
  const [planPages, setPlanPages] = useState("1");
  const [elevationPages, setElevationPages] = useState(
    String(Math.min(2, meta.page_count || 2))
  );

  return (
    <section className="mx-auto w-full max-w-2xl animate-fade-up">
      <BrandMark className="mb-8" />
      <h1 className="font-display text-5xl tracking-wide text-steel-900">
        Select pages
      </h1>
      <p className="mt-3 text-steel-600">
        <span className="font-semibold text-steel-800">{meta.filename}</span> has{" "}
        <span className="font-mono">{meta.page_count}</span> pages (more than 5).
        Tell us which sheets to scan.
      </p>

      <div className="mt-8 space-y-5 border border-steel-200 bg-white/80 p-6">
        <label className="block">
          <span className="font-mono text-xs uppercase tracking-[0.2em] text-steel-500">
            Plan page(s) — beam &amp; bracing details
          </span>
          <p className="mt-1 text-sm text-steel-500">
            From which page of plan do you want the beam and bracing details?
          </p>
          <input
            value={planPages}
            onChange={(e) => setPlanPages(e.target.value)}
            placeholder="e.g. 1,2 or 1-3"
            className="mt-2 w-full border border-steel-300 bg-steel-50 px-3 py-2.5 font-mono text-sm outline-none focus:border-ember-500"
          />
        </label>
        <label className="block">
          <span className="font-mono text-xs uppercase tracking-[0.2em] text-steel-500">
            Elevation page(s) — column &amp; elevation details
          </span>
          <p className="mt-1 text-sm text-steel-500">
            From which page of elevation do you want column and elevation details?
          </p>
          <input
            value={elevationPages}
            onChange={(e) => setElevationPages(e.target.value)}
            placeholder="e.g. 4,5 or 3-6"
            className="mt-2 w-full border border-steel-300 bg-steel-50 px-3 py-2.5 font-mono text-sm outline-none focus:border-ember-500"
          />
        </label>
        <p className="text-sm text-steel-500">
          Use commas and ranges (1-based). Example: <code className="font-mono">1,3-5</code>
        </p>
      </div>

      <div className="mt-6 flex flex-wrap gap-3">
        <button
          type="button"
          disabled={busy}
          onClick={() => {
            setError("");
            if (!planPages.trim() && !elevationPages.trim()) {
              setError("Enter at least one plan or elevation page.");
              return;
            }
            onSubmit({ planPages, elevationPages });
          }}
          className="bg-steel-900 px-6 py-3 font-mono text-xs uppercase tracking-[0.2em] text-white hover:bg-ember-600 disabled:opacity-60"
        >
          {busy ? "Scanning…" : "Scan Selected Pages"}
        </button>
        <button
          type="button"
          disabled={busy}
          onClick={onCancel}
          className="border border-steel-400 bg-white/70 px-6 py-3 font-mono text-xs uppercase tracking-[0.2em] text-steel-700 hover:bg-white"
        >
          Cancel
        </button>
      </div>

      <div className="mt-6">
        <ErrorBanner message={error} onDismiss={() => setError("")} />
      </div>
    </section>
  );
}

function LoadingScreen({ progress, label }) {
  return (
    <section className="mx-auto flex min-h-[70vh] w-full max-w-xl flex-col justify-center animate-fade-up">
      <BrandMark className="mb-10" />
      <h1 className="font-display text-5xl tracking-wide text-steel-900">
        AI is reading your drawing…
      </h1>
      <p className="mt-3 text-steel-600">{label}</p>

      <div className="relative mt-10 h-40 overflow-hidden border border-steel-300 bg-steel-900/90">
        <div className="absolute inset-0 opacity-30 bg-blueprint" />
        <div className="animate-scan-line absolute left-0 right-0 h-1 bg-gradient-to-r from-transparent via-ember-400 to-transparent" />
        <div className="absolute inset-0 flex items-center justify-center">
          <p className="animate-pulseSoft font-mono text-xs uppercase tracking-[0.3em] text-steel-100">
            Extracting beams · columns · bracing · base plates
          </p>
        </div>
      </div>

      <div className="mt-6">
        <div className="mb-2 flex justify-between font-mono text-xs uppercase tracking-wider text-steel-500">
          <span>Progress</span>
          <span>{Math.round(progress)}%</span>
        </div>
        <div className="h-2 w-full bg-steel-200">
          <div
            className="h-2 bg-ember-500 transition-all duration-500 ease-out"
            style={{ width: `${Math.min(100, Math.max(4, progress))}%` }}
          />
        </div>
      </div>
    </section>
  );
}

function DataTable({ rows }) {
  const columns = useMemo(() => {
    if (!rows?.length) return [];
    return Object.keys(rows[0]);
  }, [rows]);

  if (!rows?.length) {
    return (
      <p className="border border-dashed border-steel-300 bg-white/50 px-4 py-8 text-center text-steel-500">
        No rows extracted for this sheet.
      </p>
    );
  }

  return (
    <div className="overflow-x-auto border border-steel-200 bg-white/90">
      <table className="min-w-full text-left text-sm">
        <thead className="bg-steel-900 text-steel-50">
          <tr>
            {columns.map((col) => (
              <th
                key={col}
                className="whitespace-nowrap px-3 py-2.5 font-mono text-[11px] font-medium uppercase tracking-wider"
              >
                {col}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, idx) => (
            <tr
              key={idx}
              className={idx % 2 === 0 ? "bg-white" : "bg-steel-50/80"}
            >
              {columns.map((col) => (
                <td key={col} className="whitespace-nowrap px-3 py-2 text-steel-800">
                  {row[col] === "" || row[col] == null ? "—" : String(row[col])}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function MarkQuantityList({ items, title }) {
  if (!items?.length) return null;
  return (
    <div className="mt-4">
      <p className="font-mono text-[11px] uppercase tracking-[0.2em] text-steel-500">
        {title}
      </p>
      <ul className="mt-2 space-y-1.5 border border-steel-200 bg-white/90 p-4 text-sm text-steel-800">
        {items.map((item, idx) => (
          <li key={`${item.mark}-${item.length || item.weight}-${idx}`}>
            {item.sentence}
          </li>
        ))}
      </ul>
    </div>
  );
}

function ResultsScreen({ result, onReset }) {
  const [tab, setTab] = useState("beams");
  const preview = result.preview || {};
  const counts = result.counts || {};
  const mq = preview.mark_quantity || {};

  const activeRows =
    tab === "summary"
      ? preview.summary
      : preview[tab] || [];

  return (
    <section className="mx-auto w-full max-w-6xl animate-fade-up">
      <div className="flex flex-col gap-6 sm:flex-row sm:items-end sm:justify-between">
        <div>
          <BrandMark />
          <h1 className="mt-6 font-display text-5xl tracking-wide text-steel-900">
            Extraction complete
          </h1>
          <p className="mt-2 text-steel-600">
            {result.filename} · {result.page_count} page
            {result.page_count === 1 ? "" : "s"} scanned
          </p>
        </div>
        <div className="flex flex-wrap gap-3">
          <button
            type="button"
            onClick={() =>
              downloadBase64Excel(result.excel_base64, result.download_filename)
            }
            className="bg-ember-500 px-6 py-3 font-mono text-xs uppercase tracking-[0.2em] text-white hover:bg-ember-600"
          >
            Download Excel
          </button>
          <button
            type="button"
            onClick={onReset}
            className="border border-steel-400 bg-white/70 px-6 py-3 font-mono text-xs uppercase tracking-[0.2em] text-steel-700 hover:bg-white"
          >
            New Upload
          </button>
        </div>
      </div>

      <div className="mt-8 grid gap-3 sm:grid-cols-4">
        {[
          ["Beams", counts.beams ?? 0],
          ["Columns", counts.columns ?? 0],
          ["Bracing", counts.bracings ?? 0],
          ["Base Plates", counts.base_plates ?? 0],
        ].map(([label, value]) => (
          <div key={label} className="border border-steel-200 bg-white/80 px-4 py-3">
            <p className="font-mono text-[11px] uppercase tracking-[0.2em] text-steel-500">
              {label}
            </p>
            <p className="mt-1 font-display text-4xl text-steel-900">{value}</p>
          </div>
        ))}
      </div>

      <div className="mt-8 flex flex-wrap gap-2 border-b border-steel-300 pb-0">
        {TABS.map((t) => (
          <button
            key={t.id}
            type="button"
            onClick={() => setTab(t.id)}
            className={`px-4 py-2 font-mono text-xs uppercase tracking-[0.18em] transition ${
              tab === t.id
                ? "bg-steel-900 text-white"
                : "bg-white/60 text-steel-600 hover:bg-white"
            }`}
          >
            {t.label}
          </button>
        ))}
      </div>

      <div className="mt-4">
        {tab === "summary" && (
          <>
            <MarkQuantityList items={mq.beams} title="Beams — Mark × Length × Quantity" />
            <MarkQuantityList items={mq.columns} title="Columns — Mark × Height × Quantity" />
            <MarkQuantityList items={mq.bracings} title="Bracing — Mark × Length × Quantity" />
            <MarkQuantityList items={mq.base_plates} title="Base plates — Mark × Weight × Quantity" />
            {preview.engineer_notes && (
              <pre className="mt-4 whitespace-pre-wrap border border-steel-200 bg-white/90 p-4 font-sans text-sm leading-relaxed text-steel-800">
                {preview.engineer_notes}
              </pre>
            )}
          </>
        )}
        {tab !== "summary" && <DataTable rows={activeRows} />}
        {tab === "summary" && (
          <div className="mt-4">
            <DataTable rows={activeRows} />
          </div>
        )}
      </div>
    </section>
  );
}

function aiStatusLabel(aiStatus) {
  if (!aiStatus) return null;
  const on = aiStatus.ai_configured || aiStatus.openai_configured;
  if (!on) return { on: false, text: "AI off · set GEMINI_API_KEY in .env" };
  const provider = aiStatus.ai_provider || "ai";
  const model = aiStatus.ai_model || aiStatus.openai_model || "";
  return { on: true, text: `AI on · ${provider}${model ? ` · ${model}` : ""}` };
}

export default function App() {
  // screens: upload | pages | loading | results
  const [screen, setScreen] = useState("upload");
  const [error, setError] = useState("");
  const [file, setFile] = useState(null);
  const [jobMeta, setJobMeta] = useState(null);
  const [result, setResult] = useState(null);
  const [progress, setProgress] = useState(8);
  const [loadingLabel, setLoadingLabel] = useState(
    "Uploading and detecting digital vs scanned pages…"
  );
  const [aiStatus, setAiStatus] = useState(null);
  const progressTimer = useRef(null);

  useEffect(() => {
    let cancelled = false;
    axios
      .get(`${API_BASE}/api/health`, { timeout: 5000 })
      .then(({ data }) => {
        if (!cancelled) setAiStatus(data);
      })
      .catch(() => {
        if (!cancelled) setAiStatus(null);
      });
    return () => {
      cancelled = true;
    };
  }, [screen]);

  const clearProgress = () => {
    if (progressTimer.current) {
      clearInterval(progressTimer.current);
      progressTimer.current = null;
    }
  };

  const startProgress = (label) => {
    clearProgress();
    setLoadingLabel(label);
    setProgress(8);
    setScreen("loading");
    progressTimer.current = setInterval(() => {
      setProgress((p) => {
        if (p >= 92) return p;
        const step = p < 40 ? 3.5 : p < 70 ? 2 : 0.8;
        return Math.min(92, p + step);
      });
    }, 400);
  };

  useEffect(() => () => clearProgress(), []);

  const handleApiError = (err) => {
    clearProgress();
    const detail =
      err?.response?.data?.detail ||
      err?.message ||
      "Unexpected error talking to the SteelDraw API.";
    setError(typeof detail === "string" ? detail : JSON.stringify(detail));
    setScreen(jobMeta ? "pages" : "upload");
  };

  const extractWithFile = async (pdfFile, extras = {}) => {
    const form = new FormData();
    // Re-upload only on first pass; confirmations reuse job_id on the server
    if (pdfFile && !extras.job_id) {
      form.append("file", pdfFile);
    }
    if (extras.job_id) form.append("job_id", extras.job_id);
    if (extras.plan_pages) form.append("plan_pages", extras.plan_pages);
    if (extras.elevation_pages) form.append("elevation_pages", extras.elevation_pages);
    if (extras.confirm) form.append("confirm", extras.confirm);

    startProgress(
      extras.confirm
        ? "Scanning selected plan & elevation pages…"
        : "Uploading and detecting digital vs scanned pages…"
    );

    try {
      const { data } = await axios.post(`${API_BASE}/api/upload`, form, {
        headers: { "Content-Type": "multipart/form-data" },
        timeout: 1000 * 60 * 30, // large drawings
        onUploadProgress: (evt) => {
          if (!evt.total || extras.job_id) return;
          const pct = Math.round((evt.loaded / evt.total) * 25);
          setProgress((p) => Math.max(p, Math.min(28, pct)));
        },
      });

      if (data.needs_page_selection) {
        clearProgress();
        setJobMeta({
          job_id: data.job_id,
          page_count: data.page_count,
          filename: data.filename,
        });
        setScreen("pages");
        return;
      }

      clearProgress();
      setProgress(100);
      setResult(data);
      setScreen("results");
    } catch (err) {
      handleApiError(err);
    }
  };

  const onFileReady = (pdfFile) => {
    setFile(pdfFile);
    setJobMeta(null);
    setResult(null);
    extractWithFile(pdfFile);
  };

  const onPageSubmit = ({ planPages, elevationPages }) => {
    if (!jobMeta?.job_id) {
      setError("Session expired — please upload the PDF again.");
      setScreen("upload");
      return;
    }
    extractWithFile(null, {
      job_id: jobMeta.job_id,
      plan_pages: planPages,
      elevation_pages: elevationPages,
      confirm: "true",
    });
  };

  const reset = () => {
    clearProgress();
    setScreen("upload");
    setError("");
    setFile(null);
    setJobMeta(null);
    setResult(null);
    setProgress(8);
  };

  const status = aiStatusLabel(aiStatus);

  return (
    <div className="app-shell min-h-screen">
      <header className="border-b border-steel-300/60 bg-white/40 backdrop-blur-sm">
        <div className="mx-auto flex max-w-6xl items-center justify-between px-4 py-3 sm:px-6">
          <p className="font-mono text-[11px] uppercase tracking-[0.28em] text-steel-500">
            Structural steel · PDF → Excel
          </p>
          {status ? (
            <p
              className={`font-mono text-[11px] uppercase tracking-[0.18em] ${
                status.on ? "text-emerald-700" : "text-ember-600"
              }`}
              title={
                status.on
                  ? status.text
                  : aiStatus?.gemini?.hint ||
                    aiStatus?.openai?.hint ||
                    "Set GEMINI_API_KEY in project-root .env and restart backend"
              }
            >
              {status.text}
            </p>
          ) : (
            <p className="hidden font-mono text-[11px] text-steel-400 sm:block">
              Beams · Columns · Bracing · Base Plates
            </p>
          )}
        </div>
      </header>

      <main className="mx-auto px-4 py-10 sm:px-6 sm:py-14">
        {screen === "upload" && (
          <UploadScreen onFileReady={onFileReady} error={error} setError={setError} />
        )}
        {screen === "pages" && jobMeta && (
          <PageSelectScreen
            meta={jobMeta}
            onSubmit={onPageSubmit}
            onCancel={reset}
            error={error}
            setError={setError}
            busy={false}
          />
        )}
        {screen === "loading" && (
          <LoadingScreen progress={progress} label={loadingLabel} />
        )}
        {screen === "results" && result && (
          <ResultsScreen result={result} onReset={reset} />
        )}
      </main>

      <footer className="border-t border-steel-300/50 py-6 text-center font-mono text-[11px] uppercase tracking-[0.2em] text-steel-500">
        SteelDraw AI Extractor
      </footer>
    </div>
  );
}
