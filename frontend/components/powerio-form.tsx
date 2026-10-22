"use client";

import { useState, useRef, useCallback } from "react";
import {
  Upload,
  FileSpreadsheet,
  BarChart3,
  Plus,
  Trash2,
  Zap,
  ChevronDown,
  FileText,
  CheckCircle2,
  Loader2,
  Download,
  X,
  Copy,
  Save,
} from "lucide-react";
import { cn } from "@/lib/utils";

const BACKEND_URL = (process.env.NEXT_PUBLIC_BACKEND_URL || "http://localhost:8000").replace(/\/$/, "");

const CHART_TYPES = [
  { group: "Column", options: ["Column", "Clustered Column", "Stacked Column", "100% Stacked Column"] },
  { group: "Bar", options: ["Bar", "Clustered Bar", "Stacked Bar", "100% Stacked Bar"] },
  { group: "Line", options: ["Line", "Line with Markers", "Stacked Line", "Stacked Line with Markers", "100% Stacked Line", "100% Stacked Line with Markers"] },
  { group: "Pie / Donut", options: ["Pie", "Donut"] },
  { group: "Area", options: ["Area", "Stacked Area", "100% Stacked Area"] },
  { group: "Scatter", options: ["Scatter", "XY Scatter", "Scatter with Smooth Lines", "Scatter with Straight Lines"] },
  { group: "Radar", options: ["Radar", "Radar with Markers", "Filled Radar"] },
];

interface ChartRow {
  id: string;
  dataFile: File | null;
  slideNumber: string;
  chartType: string;
  overwrite: boolean;
}

function createRow(): ChartRow {
  return {
    id: crypto.randomUUID(),
    dataFile: null,
    slideNumber: "",
    chartType: "",
    overwrite: true,
  };
}

export default function PowerIOForm() {
  const [pptxFile, setPptxFile] = useState<File | null>(null);
  const [rows, setRows] = useState<ChartRow[]>([createRow()]);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [status, setStatus] = useState<{ type: "success" | "error"; message: string } | null>(null);
  const [dragOverPptx, setDragOverPptx] = useState(false);
  const pptxInputRef = useRef<HTMLInputElement>(null);

  const updateRow = useCallback((id: string, updates: Partial<ChartRow>) => {
    setRows((prev) => prev.map((r) => (r.id === id ? { ...r, ...updates } : r)));
  }, []);

  const removeRow = useCallback((id: string) => {
    setRows((prev) => (prev.length > 1 ? prev.filter((r) => r.id !== id) : prev));
  }, []);

  const addRow = useCallback(() => {
    setRows((prev) => [...prev, createRow()]);
  }, []);

  const handlePptxDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOverPptx(false);
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith(".pptx") || file.name.endsWith(".ppt"))) {
      setPptxFile(file);
    }
  }, []);

  const handleSubmit = async () => {
    if (!pptxFile) return;

    const invalidRows = rows.filter((r) => !r.dataFile || !r.slideNumber || !r.chartType);
    if (invalidRows.length > 0) {
      setStatus({ type: "error", message: "Please complete all chart rows before generating." });
      return;
    }

    setIsSubmitting(true);
    setStatus(null);

    try {
      const formData = new FormData();
      formData.append("pptxFile", pptxFile);

      const instructions = rows.map((r) => ({
        chart_type: r.chartType.toLowerCase(),
        slide_number: parseInt(r.slideNumber, 10),
      }));

      rows.forEach((r) => {
        formData.append("dataFiles", r.dataFile!);
      });

      formData.append("instructions", JSON.stringify(instructions));

      const res = await fetch(`${BACKEND_URL}/process`, {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: "Unknown error" }));
        throw new Error(err.error || err.detail || `Server returned ${res.status}`);
      }

      const blob = await res.blob();

      // Determine filename based on overwrite preference
      const anyOverwrite = rows.some((r) => r.overwrite);
      const filename = anyOverwrite ? pptxFile.name : `PowerIO_${pptxFile.name}`;

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      setStatus({ type: "success", message: "Charts generated and file downloaded successfully!" });
    } catch (err: unknown) {
      const message =
        err instanceof TypeError
          ? "Couldn't reach the chart server. It may be waking up, so try again in a minute."
          : err instanceof Error
            ? err.message
            : "Something went wrong";
      setStatus({ type: "error", message });
    } finally {
      setIsSubmitting(false);
    }
  };

  const isReady = pptxFile && rows.every((r) => r.dataFile && r.slideNumber && r.chartType);

  return (
    <div className="flex flex-col gap-10">
      {/* PPTX Upload */}
      <div className="animate-fade-in-up" style={{ animationDelay: "0.1s", opacity: 0 }}>
        <label className="flex items-center gap-2 text-sm font-medium text-muted-foreground mb-3">
          <FileText className="w-4 h-4 text-primary" />
          PowerPoint Presentation
        </label>
        <div
          onDragOver={(e) => {
            e.preventDefault();
            setDragOverPptx(true);
          }}
          onDragLeave={() => setDragOverPptx(false)}
          onDrop={handlePptxDrop}
          onClick={() => pptxInputRef.current?.click()}
          className={cn(
            "relative flex flex-col items-center justify-center gap-4 rounded-xl border-2 border-dashed p-10 cursor-pointer transition-all duration-300",
            dragOverPptx
              ? "border-primary bg-glow scale-[1.01]"
              : pptxFile
                ? "border-success/40 bg-success/5"
                : "border-border hover:border-primary/50 hover:bg-glow/50"
          )}
        >
          <input
            ref={pptxInputRef}
            type="file"
            accept=".pptx,.ppt"
            className="hidden"
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) setPptxFile(f);
            }}
          />
          {pptxFile ? (
            <div className="flex items-center gap-3">
              <div className="flex items-center justify-center w-12 h-12 rounded-lg bg-success/10">
                <CheckCircle2 className="w-6 h-6 text-success" />
              </div>
              <div>
                <p className="text-sm font-medium text-card-foreground">{pptxFile.name}</p>
                <p className="text-xs text-muted-foreground">
                  {(pptxFile.size / 1024).toFixed(1)} KB
                </p>
              </div>
              <button
                onClick={(e) => {
                  e.stopPropagation();
                  setPptxFile(null);
                }}
                className="ml-4 p-1.5 rounded-md hover:bg-destructive/10 text-muted-foreground hover:text-destructive transition-colors"
                aria-label="Remove file"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
          ) : (
            <>
              <div className="flex items-center justify-center w-16 h-16 rounded-2xl bg-primary/10">
                <Upload className="w-7 h-7 text-primary" />
              </div>
              <div className="text-center">
                <p className="text-sm font-medium text-card-foreground">
                  Drop your .pptx file here or click to browse
                </p>
                <p className="text-xs text-muted-foreground mt-1">
                  Supports .pptx and .ppt formats
                </p>
              </div>
            </>
          )}
        </div>
      </div>

      {/* Chart Rows */}
      <div className="animate-fade-in-up" style={{ animationDelay: "0.2s", opacity: 0 }}>
        <div className="flex items-center justify-between mb-4">
          <label className="flex items-center gap-2 text-sm font-medium text-muted-foreground">
            <BarChart3 className="w-4 h-4 text-accent" />
            Chart Configurations
          </label>
          <span className="text-xs text-muted-foreground font-mono">
            {rows.length} {rows.length === 1 ? "chart" : "charts"}
          </span>
        </div>

        <div className="flex flex-col gap-4">
          {rows.map((row, idx) => (
            <ChartRowCard
              key={row.id}
              row={row}
              index={idx}
              total={rows.length}
              onUpdate={updateRow}
              onRemove={removeRow}
            />
          ))}
        </div>

        <button
          onClick={addRow}
          className="mt-4 flex items-center gap-2 rounded-lg border border-dashed border-border px-4 py-3 text-sm text-muted-foreground hover:border-primary/50 hover:text-primary hover:bg-glow/50 transition-all w-full justify-center"
        >
          <Plus className="w-4 h-4" />
          Add another chart
        </button>
      </div>

      {/* Status Message */}
      {status && (
        <div
          className={cn(
            "rounded-lg px-4 py-3 text-sm",
            status.type === "success"
              ? "bg-success/10 text-success border border-success/20"
              : "bg-destructive/10 text-destructive border border-destructive/20"
          )}
        >
          {status.message}
        </div>
      )}

      {/* Submit */}
      <div className="animate-fade-in-up" style={{ animationDelay: "0.3s", opacity: 0 }}>
        <button
          onClick={handleSubmit}
          disabled={!isReady || isSubmitting}
          className={cn(
            "w-full flex items-center justify-center gap-3 rounded-xl px-6 py-4 text-base font-semibold transition-all duration-300",
            isReady && !isSubmitting
              ? "bg-primary text-primary-foreground hover:bg-primary-hover shadow-lg shadow-primary/20 hover:shadow-primary/40 animate-pulse-glow"
              : "bg-muted text-muted-foreground cursor-not-allowed"
          )}
        >
          {isSubmitting ? (
            <>
              <Loader2 className="w-5 h-5 animate-spin" />
              Generating Charts...
            </>
          ) : (
            <>
              <Zap className="w-5 h-5" />
              Generate and Insert Charts
            </>
          )}
        </button>
        <p className="text-center text-xs text-muted-foreground mt-3">
          The PowerPoint file must be closed before generating charts.
        </p>
      </div>
    </div>
  );
}

/* ─── Chart Row Card ────────────────────────────────────────────────────── */

function ChartRowCard({
  row,
  index,
  total,
  onUpdate,
  onRemove,
}: {
  row: ChartRow;
  index: number;
  total: number;
  onUpdate: (id: string, updates: Partial<ChartRow>) => void;
  onRemove: (id: string) => void;
}) {
  const fileRef = useRef<HTMLInputElement>(null);
  const [dropdownOpen, setDropdownOpen] = useState(false);
  const [dragOver, setDragOver] = useState(false);

  return (
    <div className="relative rounded-xl border border-border bg-card p-5 transition-all hover:border-border-hover group">
      {/* Row number badge */}
      <div className="absolute -top-2.5 left-4 px-2.5 py-0.5 rounded-md bg-primary/10 text-primary text-xs font-semibold font-mono">
        Chart {index + 1}
      </div>

      {/* Remove button */}
      {total > 1 && (
        <button
          onClick={() => onRemove(row.id)}
          className="absolute top-3 right-3 p-1.5 rounded-md text-muted-foreground hover:text-destructive hover:bg-destructive/10 opacity-0 group-hover:opacity-100 transition-all"
          aria-label="Remove chart row"
        >
          <Trash2 className="w-4 h-4" />
        </button>
      )}

      <div className="flex flex-col gap-4 mt-2">
        {/* Top row: chart type + slide number */}
        <div className="flex gap-4 flex-col sm:flex-row">
          {/* Chart Type Dropdown */}
          <div className="flex-1 relative">
            <label className="text-xs text-muted-foreground mb-1.5 block">Chart Type</label>
            <button
              type="button"
              onClick={() => setDropdownOpen(!dropdownOpen)}
              className={cn(
                "w-full flex items-center justify-between rounded-lg border border-border bg-background px-3 py-2.5 text-sm transition-all",
                dropdownOpen ? "ring-2 ring-ring/40 border-primary" : "hover:border-border-hover",
                row.chartType ? "text-card-foreground" : "text-muted-foreground"
              )}
            >
              <span>{row.chartType || "Select chart type"}</span>
              <ChevronDown className={cn("w-4 h-4 text-muted-foreground transition-transform", dropdownOpen && "rotate-180")} />
            </button>
            {dropdownOpen && (
              <>
                <div className="fixed inset-0 z-10" onClick={() => setDropdownOpen(false)} />
                <div className="absolute z-20 mt-1 w-full max-h-64 overflow-y-auto rounded-lg border border-border bg-card shadow-xl shadow-background/60">
                  {CHART_TYPES.map((group) => (
                    <div key={group.group}>
                      <div className="px-3 py-1.5 text-[11px] font-semibold text-muted-foreground uppercase tracking-wider bg-muted/50">
                        {group.group}
                      </div>
                      {group.options.map((opt) => (
                        <button
                          key={opt}
                          onClick={() => {
                            onUpdate(row.id, { chartType: opt });
                            setDropdownOpen(false);
                          }}
                          className={cn(
                            "w-full text-left px-3 py-2 text-sm hover:bg-primary/10 hover:text-primary transition-colors",
                            row.chartType === opt ? "text-primary bg-primary/5" : "text-card-foreground"
                          )}
                        >
                          {opt}
                        </button>
                      ))}
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>

          {/* Slide Number */}
          <div className="w-full sm:w-32">
            <label className="text-xs text-muted-foreground mb-1.5 block">Slide Number</label>
            <input
              type="number"
              min={1}
              placeholder="e.g. 3"
              value={row.slideNumber}
              onChange={(e) => onUpdate(row.id, { slideNumber: e.target.value })}
              className="w-full rounded-lg border border-border bg-background px-3 py-2.5 text-sm text-card-foreground placeholder:text-muted-foreground focus:outline-none focus:ring-2 focus:ring-ring/40 focus:border-primary transition-all font-mono"
            />
          </div>
        </div>

        {/* Bottom row: file upload + overwrite checkbox */}
        <div className="flex gap-4 items-end flex-col sm:flex-row">
          {/* Data File Upload */}
          <div className="flex-1 w-full">
            <label className="text-xs text-muted-foreground mb-1.5 block">Data File (Excel / CSV)</label>
            <div
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={(e) => {
                e.preventDefault();
                setDragOver(false);
                const file = e.dataTransfer.files[0];
                if (file) onUpdate(row.id, { dataFile: file });
              }}
              onClick={() => fileRef.current?.click()}
              className={cn(
                "flex items-center gap-3 rounded-lg border border-dashed px-3 py-2.5 cursor-pointer transition-all text-sm",
                dragOver
                  ? "border-accent bg-glow-accent"
                  : row.dataFile
                    ? "border-success/30 bg-success/5"
                    : "border-border hover:border-primary/40 hover:bg-glow/30"
              )}
            >
              <input
                ref={fileRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                className="hidden"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) onUpdate(row.id, { dataFile: f });
                }}
              />
              {row.dataFile ? (
                <>
                  <FileSpreadsheet className="w-4 h-4 text-success shrink-0" />
                  <span className="text-card-foreground truncate">{row.dataFile.name}</span>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      onUpdate(row.id, { dataFile: null });
                    }}
                    className="ml-auto p-1 rounded hover:bg-destructive/10 text-muted-foreground hover:text-destructive transition-colors shrink-0"
                    aria-label="Remove data file"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                </>
              ) : (
                <>
                  <FileSpreadsheet className="w-4 h-4 text-muted-foreground shrink-0" />
                  <span className="text-muted-foreground">Drop or click to upload</span>
                </>
              )}
            </div>
          </div>

          {/* Overwrite Checkbox */}
          <div className="flex items-center gap-2.5 shrink-0 pb-0.5">
            <button
              type="button"
              role="checkbox"
              aria-checked={row.overwrite}
              onClick={() => onUpdate(row.id, { overwrite: !row.overwrite })}
              className={cn(
                "w-5 h-5 rounded-[var(--radius-sm)] border-2 flex items-center justify-center transition-all",
                row.overwrite
                  ? "bg-primary border-primary"
                  : "border-border hover:border-primary/50"
              )}
            >
              {row.overwrite && (
                <svg viewBox="0 0 12 12" className="w-3 h-3 text-primary-foreground">
                  <path
                    d="M2 6l3 3 5-6"
                    stroke="currentColor"
                    strokeWidth="2"
                    fill="none"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                </svg>
              )}
            </button>
            <span className="text-xs text-muted-foreground flex items-center gap-1.5">
              {row.overwrite ? (
                <>
                  <Save className="w-3.5 h-3.5" />
                  Overwrite original
                </>
              ) : (
                <>
                  <Copy className="w-3.5 h-3.5" />
                  Save as copy
                </>
              )}
            </span>
          </div>
        </div>
      </div>
    </div>
  );
}
