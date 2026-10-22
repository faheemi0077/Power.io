import Header from "@/components/header";
import PowerIOForm from "@/components/powerio-form";

export default function Home() {
  return (
    <main className="min-h-screen flex flex-col">
      <div className="flex-1 w-full max-w-3xl mx-auto px-4 sm:px-6 pb-20">
        <Header />
        <PowerIOForm />
      </div>

      {/* Footer */}
      <footer className="border-t border-border py-6">
        <p className="text-center text-xs text-muted-foreground">
          PowerIO &mdash; AI-powered chart generation for PowerPoint
        </p>
      </footer>
    </main>
  );
}
