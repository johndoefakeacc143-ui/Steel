import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import axios from "axios";

const API_BASE = import.meta.env.VITE_API_URL || "";
const MAX_BYTES = 500 * 1024 * 1024;

const STEPS = {
  UPLOAD: "upload",
  LOADING: "loading",
  RESULTS: "results",
};

function formatBytes(bytes) {
  if (!bytes && bytes !== 0) return "—";
  const units = ["B", "KB", "MB", "GB"];
  let n = bytes;
  let i = 0;
  while (n >= 1024 && i < units.length - 1) {
    n /= 1024;
    i += 1;
  }
  return `${n.toFixed(i === 0 ? 0 : 1)} ${units[i]}`;
}

function downloadBase64Excel(base64, filename) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
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
      <svg
        width="40"
        height="40"
        viewBox="0 0 64 64"
        fill="none"
        aria-hidden="true"
        className="shrink-0"
      >
        <rect width="64" height="64" rx="10" fill="#2e3944" />
        <path
          d="M12 40 L32 14 L52 40"
          stroke="#c96a1a"
          strokeWidth="5"
          strokeLinejoin="round"
        />
        <rect x="18" y="40" width="28" height="8" rx="1" fill="#9fb8c6" />
      </svg>
      <div className="leading-none">
        <p className="font-display text-2xl font-semibold tracking-wide text-steel-900 sm:text-3xl">
          SteelDraw
        </p>
        <p className="mt-1 text-[11px] font-semibold uppercase tracking-[0.22em] text-ember-600">
          AI Extractor
        </p>
      </div>
    </div>
  );
}

function UploadScreen({ file, setFile, onStart, error, setError }) {
  const inputRef = useRef(null);
  const [dragging, setDragging] = useState(false);

  const acceptFile = useCallback(
    (next) => {
      setError("");
      if (!next) return;
      if (next.type !== "application/pdf" && !next.name.toLowerCase().endsWith(".pdf")) {
        setError("Please upload a PDF drawing file.");
        return;
      }
      if (next.size > MAX_BYTES) {
        setError("File exceeds the 500 MB limit.");
        return;
      }
      setFile(next);
    },
    [setError, setFile]
  );

  const onDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    const dropped = e.dataTransfer.files?.[0];
    acceptFile(dropped);
  };

  return (
    <section className="mx-auto flex w-full max-w-3xl flex-col items-center px-5 pb-16 pt-10 sm:pt-16">
      <BrandMark className="animate-rise mb-10" />

      <h1 className="animate-rise font-display text-4xl font-semibold uppercase tracking-wide text-steel-900 sm:text-5xl" style={{ animationDelay: "80ms" }}>
        Read the steel. Export the BOM.
      </h1>
      <p className="animate-rise mt-4 max-w-xl text-center text-base text-steel-600 sm:text-lg" style={{ animationDelay: "140ms" }}>
        Drop a structural steel PDF. AI extracts beams, columns, and base plates into Excel.
      </p>

      <div
        role="button"
        tabIndex={0}
        onKeyDown={(e) => {
          if (e.key === "Enter" || e.key === " ") inputRef.current?.click();
        }}
        onDragEnter={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={(e) => {
          e.preventDefault();
          setDragging(false);
        }}
        onDrop={onDrop}
        onClick={() => inputRef.current?.click()}
        className={`animate-rise drop-hatch mt-10 w-full cursor-pointer border-2 border-dashed px-6 py-14 text-center transition duration-300 ${
          dragging
            ? "border-ember-500 bg-ember-500/10"
            : "border-steel-300 bg-white/50 hover:border-steel-500 hover:bg-white/80"
        }`}
        style={{ animationDelay: "200ms" }}
      >
        <input
          ref={inputRef}
          type="file"
          accept="application/pdf,.pdf"
          className="hidden"
          onChange={(e) => acceptFile(e.target.files?.[0])}
        />
        <div className="mx-auto mb-4 flex h-14 w-14 items-center justify-center border border-steel-300 bg-steel-50">
          <svg width="28" height="28" viewBox="0 0 24 24" fill="none" aria-hidden="true">
            <path
              d="M12 16V4M12 4L8 8M12 4L16 8M4 16V18C4 19.1 4.9 20 6 20H18C19.1 20 20 19.1 20 18V16"
              stroke="#527589"
              strokeWidth="1.8"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </svg>
        </div>
        <p className="font-display text-2xl font-semibold uppercase tracking-wide text-steel-800">
          Drag & drop PDF here
        </p>
        <p className="mt-2 text-sm text-steel-500">or click to browse · up to 500 MB</p>
      </div>

      {file && (
        <div className="animate-rise mt-6 w-full border border-steel-200 bg-white/80 px-5 py-4">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <p className="text-xs font-semibold uppercase tracking-[0.16em] text-steel-400">
                Selected drawing
              </p>
              <p className="mt-1 font-medium text-steel-900">{file.name}</p>
              <p className="text-sm text-steel-500">{formatBytes(file.size)}</p>
            </div>
            <button
              type="button"
              onClick={() => setFile(null)}
              className="text-sm font-semibold text-steel-500 underline-offset-2 hover:text-ember-600 hover:underline"
            >
              Remove
            </button>
          </div>
        </div>
      )}

      {error && (
        <div
          role="alert"
          className="animate-rise mt-5 w-full border border-red-300 bg-red-50 px-4 py-3 text-sm text-red-800"
        >
          {error}
        </div>
      )}

      <button
        type="button"
        disabled={!file}
        onClick={onStart}
        className="animate-rise mt-8 bg-steel-900 px-10 py-3.5 font-display text-lg font-semibold uppercase tracking-[0.14em] text-white transition hover:bg-steel-800 disabled:cursor-not-allowed disabled:bg-steel-300"
        style={{ animationDelay: "260ms" }}
      >
        Extract with AI
      </button>
    </section>
  );
}

function LoadingScreen({ progress, fileName }) {
  return (
    <section className="mx-auto flex min-h-[70vh] w-full max-w-xl flex-col items-center justify-center px-5 py-16">
      <BrandMark className="mb-12" />

      <div className="relative w-full overflow-hidden border border-steel-200 bg-white/70 px-6 py-10">
        <div className="pointer-events-none absolute inset-x-0 top-0 h-1/3 bg-gradient-to-b from-ember-500/10 to-transparent animate-scan" />
        <p className="text-center font-display text-3xl font-semibold uppercase tracking-wide text-steel-900">
          AI is reading your drawing…
        </p>
        <p className="mt-3 text-center text-sm text-steel-500">
          {fileName || "Processing PDF"} · page-by-page extraction
        </p>

        <div className="mt-8 h-2 w-full overflow-hidden bg-steel-100">
          <div
            className="h-full bg-ember-500 transition-all duration-500 ease-out animate-pulse-bar"
            style={{ width: `${Math.min(progress, 100)}%` }}
          />
        </div>
        <p className="mt-3 text-center font-display text-xl font-semibold text-steel-700">
          {Math.min(Math.round(progress), 99)}%
        </p>

        <ul className="mt-8 space-y-2 text-sm text-steel-600">
          <li className={progress > 10 ? "text-steel-900" : ""}>• Detecting digital vs scanned pages</li>
          <li className={progress > 35 ? "text-steel-900" : ""}>• OCR / table parse when needed</li>
          <li className={progress > 60 ? "text-steel-900" : ""}>• Extracting beams, columns, base plates</li>
          <li className={progress > 85 ? "text-steel-900" : ""}>• Building Excel workbook</li>
        </ul>
      </div>
    </section>
  );
}

function DataTable({ columns, rows, emptyLabel }) {
  if (!rows?.length) {
    return (
      <p className="border border-dashed border-steel-200 px-4 py-8 text-center text-sm text-steel-500">
        {emptyLabel}
      </p>
    );
  }

  return (
    <div className="overflow-x-auto border border-steel-200">
      <table className="min-w-full text-left text-sm">
        <thead className="bg-steel-900 text-white">
          <tr>
            {columns.map((col) => (
              <th
                key={col}
                className="whitespace-nowrap px-3 py-2.5 font-display text-sm font-semibold uppercase tracking-wide"
              >
                {col}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, idx) => (
            <tr
              key={`${row.Mark || "row"}-${idx}`}
              className={idx % 2 === 0 ? "bg-white" : "bg-steel-50"}
            >
              {columns.map((col) => (
                <td key={col} className="whitespace-nowrap px-3 py-2 text-steel-800">
                  {row[col] ?? "—"}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function ResultsScreen({ result, onReset, onDownload }) {
  const [tab, setTab] = useState("Beams");

  const tabs = useMemo(
    () => [
      { id: "Beams", count: result.beams?.length || 0 },
      { id: "Columns", count: result.columns?.length || 0 },
      { id: "BasePlates", count: result.base_plates?.length || 0 },
      { id: "Summary", count: result.summary?.length || 0 },
    ],
    [result]
  );

  const beamCols = ["Mark", "Section Size", "Length", "Material", "Start EL", "End EL", "Page"];
  const colCols = [
    "Mark",
    "Section Size",
    "Height",
    "Base Elevation",
    "Top Elevation",
    "Material",
    "Page",
  ];
  const plateCols = [
    "Mark",
    "Plate Size",
    "Thickness",
    "Anchor Bolt Dia",
    "Anchor Bolt Qty",
    "Top of Concrete EL",
    "Page",
  ];

  return (
    <section className="mx-auto w-full max-w-6xl px-5 pb-20 pt-8">
      <div className="flex flex-wrap items-end justify-between gap-4">
        <BrandMark />
        <div className="flex flex-wrap gap-3">
          <button
            type="button"
            onClick={onReset}
            className="border border-steel-300 bg-white px-5 py-2.5 text-sm font-semibold uppercase tracking-wide text-steel-700 hover:border-steel-500"
          >
            New upload
          </button>
          <button
            type="button"
            onClick={onDownload}
            className="bg-ember-500 px-5 py-2.5 text-sm font-semibold uppercase tracking-wide text-white hover:bg-ember-600"
          >
            Download Excel
          </button>
        </div>
      </div>

      <div className="animate-rise mt-8 grid gap-3 sm:grid-cols-4">
        {[
          { label: "Beams", value: result.beams?.length || 0 },
          { label: "Columns", value: result.columns?.length || 0 },
          { label: "Base plates", value: result.base_plates?.length || 0 },
          { label: "Pages", value: result.pages || 0 },
        ].map((stat) => (
          <div key={stat.label} className="border border-steel-200 bg-white/70 px-4 py-4">
            <p className="text-xs font-semibold uppercase tracking-[0.16em] text-steel-400">
              {stat.label}
            </p>
            <p className="mt-1 font-display text-3xl font-semibold text-steel-900">{stat.value}</p>
          </div>
        ))}
      </div>

      <p className="mt-4 text-sm text-steel-500">
        PDF type: <span className="font-semibold text-steel-700">{result.pdf_type || "unknown"}</span>
        {" · "}
        File: <span className="font-semibold text-steel-700">{result.filename}</span>
      </p>

      <div className="mt-8 flex flex-wrap gap-2 border-b border-steel-200 pb-px">
        {tabs.map((t) => (
          <button
            key={t.id}
            type="button"
            onClick={() => setTab(t.id)}
            className={`px-4 py-2 font-display text-base font-semibold uppercase tracking-wide transition ${
              tab === t.id
                ? "border-b-2 border-ember-500 text-steel-900"
                : "text-steel-400 hover:text-steel-700"
            }`}
          >
            {t.id} ({t.count})
          </button>
        ))}
      </div>

      <div className="mt-5 animate-rise">
        {tab === "Beams" && (
          <DataTable columns={beamCols} rows={result.beams} emptyLabel="No beams detected." />
        )}
        {tab === "Columns" && (
          <DataTable columns={colCols} rows={result.columns} emptyLabel="No columns detected." />
        )}
        {tab === "BasePlates" && (
          <DataTable
            columns={plateCols}
            rows={result.base_plates}
            emptyLabel="No base plates detected."
          />
        )}
        {tab === "Summary" && (
          <DataTable
            columns={["Metric", "Value"]}
            rows={result.summary}
            emptyLabel="No summary available."
          />
        )}
      </div>
    </section>
  );
}

export default function App() {
  const [step, setStep] = useState(STEPS.UPLOAD);
  const [file, setFile] = useState(null);
  const [error, setError] = useState("");
  const [progress, setProgress] = useState(0);
  const [result, setResult] = useState(null);
  const progressTimer = useRef(null);

  useEffect(() => {
    return () => {
      if (progressTimer.current) clearInterval(progressTimer.current);
    };
  }, []);

  const startFakeProgress = () => {
    setProgress(4);
    if (progressTimer.current) clearInterval(progressTimer.current);
    progressTimer.current = setInterval(() => {
      setProgress((p) => {
        if (p >= 92) return p;
        const bump = p < 40 ? 3.5 : p < 70 ? 2 : 0.8;
        return Math.min(p + bump, 92);
      });
    }, 450);
  };

  const stopFakeProgress = (finalValue = 100) => {
    if (progressTimer.current) {
      clearInterval(progressTimer.current);
      progressTimer.current = null;
    }
    setProgress(finalValue);
  };

  const handleExtract = async () => {
    if (!file) return;
    setError("");
    setStep(STEPS.LOADING);
    startFakeProgress();

    const form = new FormData();
    form.append("file", file);

    try {
      const response = await axios.post(`${API_BASE}/api/extract`, form, {
        headers: { "Content-Type": "multipart/form-data" },
        timeout: 30 * 60 * 1000, // large drawings can take a while
        onUploadProgress: (evt) => {
          if (!evt.total) return;
          const uploadPct = Math.round((evt.loaded / evt.total) * 25);
          setProgress((p) => Math.max(p, uploadPct));
        },
      });

      stopFakeProgress(100);
      setResult(response.data);
      setTimeout(() => setStep(STEPS.RESULTS), 350);
    } catch (err) {
      stopFakeProgress(0);
      const detail =
        err.response?.data?.detail ||
        err.message ||
        "Something went wrong while processing the drawing.";
      setError(typeof detail === "string" ? detail : JSON.stringify(detail));
      setStep(STEPS.UPLOAD);
    }
  };

  const handleDownload = () => {
    if (!result?.excel_base64) return;
    downloadBase64Excel(result.excel_base64, result.filename);
  };

  const handleReset = () => {
    setStep(STEPS.UPLOAD);
    setFile(null);
    setResult(null);
    setError("");
    setProgress(0);
  };

  return (
    <div className="relative min-h-screen overflow-hidden bg-forge">
      <div className="pointer-events-none absolute inset-0 bg-grid opacity-70" />
      <div className="relative z-10">
        {step === STEPS.UPLOAD && (
          <UploadScreen
            file={file}
            setFile={setFile}
            onStart={handleExtract}
            error={error}
            setError={setError}
          />
        )}
        {step === STEPS.LOADING && (
          <LoadingScreen progress={progress} fileName={file?.name} />
        )}
        {step === STEPS.RESULTS && result && (
          <ResultsScreen
            result={result}
            onReset={handleReset}
            onDownload={handleDownload}
          />
        )}
      </div>
    </div>
  );
}
