import { Header } from '@/components/Header';
import { UploadDropzone } from '@/components/UploadDropzone';
import { WelcomeHero } from '@/components/WelcomeHero';
import { StatsOverview } from '@/components/StatsOverview';
import { FilterBar } from '@/components/FilterBar';
import { PreviewTable } from '@/components/PreviewTable';
import { ChartPanel } from '@/components/ChartPanel';
import { DownloadBar } from '@/components/DownloadBar';
import { WarningsPanel } from '@/components/WarningsPanel';
import { useSurveyStore } from '@/store/surveyStore';
import { Loader2, AlertCircle } from 'lucide-react';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';

const Index = () => {
  const { parsedSurvey, isLoading, error } = useSurveyStore();

  return (
    <div className="min-h-screen bg-gradient-to-br from-background via-background to-muted/30">
      <Header />
      
      <main className="container mx-auto px-4 py-8">
        {!parsedSurvey && !isLoading && (
          <>
            <WelcomeHero />
            <div className="max-w-2xl mx-auto">
              <UploadDropzone />
            </div>
          </>
        )}

        {isLoading && (
          <div className="flex flex-col items-center justify-center py-24 animate-fade-in">
            <Loader2 className="w-12 h-12 text-primary animate-spin mb-4" />
            <p className="text-lg font-medium text-foreground">Elaborazione in corso...</p>
            <p className="text-sm text-muted-foreground">Analisi delle domande e calcolo statistiche</p>
          </div>
        )}

        {error && (
          <div className="max-w-2xl mx-auto animate-fade-in">
            <Alert variant="destructive">
              <AlertCircle className="h-4 w-4" />
              <AlertTitle>Errore</AlertTitle>
              <AlertDescription>{error}</AlertDescription>
            </Alert>
            <div className="mt-8">
              <UploadDropzone />
            </div>
          </div>
        )}

        {parsedSurvey && !isLoading && (
          <div className="space-y-6">
            <StatsOverview />
            <WarningsPanel />
            <DownloadBar />
            <FilterBar />
            
            <div className="grid lg:grid-cols-2 gap-6">
              <PreviewTable />
              <ChartPanel />
            </div>
          </div>
        )}
      </main>

      <footer className="border-t border-border/50 mt-12 py-6 text-center text-sm text-muted-foreground">
        <p>Magnews Survey Analyzer • Client-side processing • I tuoi dati non lasciano il browser</p>
      </footer>
    </div>
  );
};

export default Index;
