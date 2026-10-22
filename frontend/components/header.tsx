import {
  Zap,
  BarChart3,
  FileSpreadsheet,
  ArrowRight,
} from "lucide-react";

export default function Header() {
  return (
    <header className="relative">
      {/* Ambient glow */}
      <div className="absolute inset-0 -z-10 overflow-hidden">
        <div className="absolute top-0 left-1/4 w-96 h-96 bg-primary/5 rounded-full blur-3xl" />
        <div className="absolute top-10 right-1/4 w-72 h-72 bg-accent/5 rounded-full blur-3xl" />
      </div>

      <div className="flex flex-col items-center text-center gap-6 pt-16 pb-10">
        {/* Logo mark */}
        <div className="flex items-center justify-center w-16 h-16 rounded-2xl bg-primary/10 border border-primary/20">
          <Zap className="w-8 h-8 text-primary" />
        </div>

        <div className="flex flex-col gap-3">
          <h1 className="text-4xl sm:text-5xl font-bold tracking-tight text-balance text-foreground">
            Power<span className="text-primary">IO</span>
          </h1>
          <p className="text-lg text-muted-foreground max-w-xl mx-auto text-pretty leading-relaxed">
            Upload your presentation, attach data files, and generate beautiful charts on any slide instantly.
          </p>
        </div>

        {/* Feature pills */}
        <div className="flex flex-wrap items-center justify-center gap-3 mt-2">
          <FeaturePill icon={<FileSpreadsheet className="w-3.5 h-3.5" />} text="Excel & CSV" />
          <FeaturePill icon={<BarChart3 className="w-3.5 h-3.5" />} text="20+ Chart Types" />
          <FeaturePill icon={<ArrowRight className="w-3.5 h-3.5" />} text="Instant Export" />
        </div>
      </div>
    </header>
  );
}

function FeaturePill({ icon, text }: { icon: React.ReactNode; text: string }) {
  return (
    <span className="flex items-center gap-1.5 rounded-full border border-border bg-card px-3 py-1.5 text-xs text-muted-foreground">
      {icon}
      {text}
    </span>
  );
}
