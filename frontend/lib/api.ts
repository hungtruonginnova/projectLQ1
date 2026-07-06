export type ReportStats = {
  input_counts: {
    opd: number;
    invmaster: number;
    workorders: number;
    schedule: number;
  };
  drop_log: [string, number][];
  row_count: number;
  review_count: number;
  report_date: string;
  out_path: string;
};

export type ReportFiles = {
  opd: File;
  invmaster: File;
  workorders: File;
  schedule: File;
};

export type ReportResult = {
  blob: Blob;
  filename: string;
  stats: ReportStats;
};

const API_URL =
  process.env.NEXT_PUBLIC_API_URL?.replace(/\/$/, "") ??
  "https://truonghung239810-lqvhspace.hf.space";

export async function checkHealth(): Promise<boolean> {
  const res = await fetch(`${API_URL}/health`);
  return res.ok;
}

export async function generateReport(
  files: ReportFiles,
  reportDate: string,
): Promise<ReportResult> {
  const form = new FormData();
  form.append("opd", files.opd);
  form.append("invmaster", files.invmaster);
  form.append("workorders", files.workorders);
  form.append("schedule", files.schedule);
  form.append("report_date", reportDate);

  const res = await fetch(`${API_URL}/api/v1/report`, {
    method: "POST",
    body: form,
  });

  if (!res.ok) {
    let detail = `Request failed (${res.status})`;
    try {
      const body = await res.json();
      if (body.detail) {
        detail = typeof body.detail === "string" ? body.detail : JSON.stringify(body.detail);
      } else if (body.error) {
        detail = body.error;
      }
    } catch {
      const text = await res.text();
      if (text) detail = text.slice(0, 500);
    }
    throw new Error(detail);
  }

  const statsHeader = res.headers.get("X-Report-Stats");
  const filenameHeader = res.headers.get("X-Report-Filename");
  if (!statsHeader) {
    throw new Error("Missing X-Report-Stats header from API");
  }

  const stats = JSON.parse(statsHeader) as ReportStats;
  const blob = await res.blob();
  const disposition = res.headers.get("Content-Disposition") ?? "";
  const match = disposition.match(/filename="([^"]+)"/);
  const filename = filenameHeader ?? match?.[1] ?? "OVERDUE PARTS REPORT.xlsx";

  return { blob, filename, stats };
}
