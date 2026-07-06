"use client";

import { useMemo, useState } from "react";
import { generateReport, type ReportResult } from "@/lib/api";

type FileKey = "opd" | "invmaster" | "workorders" | "schedule";

const FILE_LABELS: Record<FileKey, string> = {
  opd: "1. OPD export",
  invmaster: "2. InvMaster export",
  workorders: "3. GAB 6271 Open Work Orders",
  schedule: "4. Production schedule",
};

function todayIso(): string {
  return new Date().toISOString().slice(0, 10);
}

export default function HomePage() {
  const [files, setFiles] = useState<Partial<Record<FileKey, File>>>({});
  const [reportDate, setReportDate] = useState(todayIso);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<ReportResult | null>(null);

  const ready = useMemo(
    () => Boolean(files.opd && files.invmaster && files.workorders && files.schedule),
    [files],
  );

  function onFileChange(key: FileKey, fileList: FileList | null) {
    const file = fileList?.[0];
    setFiles((prev) => ({ ...prev, [key]: file }));
    setError(null);
    setResult(null);
  }

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!files.opd || !files.invmaster || !files.workorders || !files.schedule) return;

    setLoading(true);
    setError(null);
    setResult(null);

    try {
      const report = await generateReport(
        {
          opd: files.opd,
          invmaster: files.invmaster,
          workorders: files.workorders,
          schedule: files.schedule,
        },
        reportDate,
      );
      setResult(report);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Unknown error");
    } finally {
      setLoading(false);
    }
  }

  function downloadResult() {
    if (!result) return;
    const url = URL.createObjectURL(result.blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = result.filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  const dropTotal = result?.stats.drop_log.reduce((sum, [, n]) => sum + n, 0) ?? 0;

  return (
    <main>
      <h1>Overdue Parts Report</h1>
      <p className="subtitle">
        Upload four weekly ERP/GAB exports to generate the Seawind overdue-parts report.
        Processing may take up to 2 minutes for large GAB 6271 files.
      </p>

      <form onSubmit={onSubmit}>
        <div className="card">
          <label htmlFor="report-date">Report date</label>
          <input
            id="report-date"
            type="date"
            value={reportDate}
            onChange={(e) => setReportDate(e.target.value)}
          />
        </div>

        <div className="card grid">
          {(Object.keys(FILE_LABELS) as FileKey[]).map((key) => (
            <div key={key}>
              <label htmlFor={key}>{FILE_LABELS[key]}</label>
              <input
                id={key}
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => onFileChange(key, e.target.files)}
              />
            </div>
          ))}
        </div>

        {error && <div className="alert error">{error}</div>}

        {loading && (
          <div className="alert info">
            <span className="spinner" />
            Building report — please wait. Large files can take up to 2 minutes.
          </div>
        )}

        <button className="primary" type="submit" disabled={!ready || loading}>
          {loading ? "Generating…" : "Generate report"}
        </button>
      </form>

      {result && (
        <div className="card" style={{ marginTop: "1.5rem" }}>
          <h2 style={{ margin: "0 0 0.5rem", fontSize: "1.1rem" }}>Report ready</h2>
          <p style={{ margin: 0, color: "var(--muted)", fontSize: "0.9rem" }}>
            {result.filename}
          </p>

          <div className="metrics">
            <div className="metric">
              <div className="value">{result.stats.input_counts.opd.toLocaleString()}</div>
              <div className="label">OPD lines</div>
            </div>
            <div className="metric">
              <div className="value">{result.stats.input_counts.invmaster.toLocaleString()}</div>
              <div className="label">InvMaster parts</div>
            </div>
            <div className="metric">
              <div className="value">{result.stats.input_counts.workorders.toLocaleString()}</div>
              <div className="label">6271 WO lines</div>
            </div>
            <div className="metric">
              <div className="value">{result.stats.input_counts.schedule.toLocaleString()}</div>
              <div className="label">Schedule boats</div>
            </div>
            <div className="metric">
              <div className="value">{result.stats.row_count.toLocaleString()}</div>
              <div className="label">Final rows</div>
            </div>
            <div className="metric">
              <div className="value">{result.stats.review_count.toLocaleString()}</div>
              <div className="label">Flagged for review</div>
            </div>
          </div>

          {result.stats.drop_log.length > 0 && (
            <details style={{ marginTop: "1rem" }}>
              <summary style={{ cursor: "pointer", fontSize: "0.9rem" }}>
                Rows removed by rule ({dropTotal.toLocaleString()} total)
              </summary>
              <ul className="drop-list">
                {result.stats.drop_log.map(([reason, n]) => (
                  <li key={reason}>
                    <strong>{n.toLocaleString()}</strong> — {reason}
                  </li>
                ))}
              </ul>
            </details>
          )}

          <button type="button" className="download" onClick={downloadResult}>
            Download Excel report
          </button>
        </div>
      )}
    </main>
  );
}
