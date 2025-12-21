import { useState } from 'react';
import { Download, FileSpreadsheet, Image, Loader2, Table } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { useSurveyStore } from '@/store/surveyStore';
import { generateFilePerQuestionari } from '@/utils/filePerQuestionariWriter';
import { generateTabellaGrafici } from '@/utils/tabellaGraficiWriter';
import { generateChartsZip } from '@/utils/chartExporter';
import { toast } from '@/hooks/use-toast';

type ExportType = 'file_per_questionari' | 'tabella_grafici' | 'charts' | null;

export function DownloadBar() {
  const { parsedSurvey } = useSurveyStore();
  const [isExporting, setIsExporting] = useState<ExportType>(null);
  const [progress, setProgress] = useState({ current: 0, total: 0 });

  if (!parsedSurvey) return null;

  const handleFilePerQuestionariDownload = async () => {
    setIsExporting('file_per_questionari');
    try {
      await generateFilePerQuestionari(parsedSurvey);
      toast({
        title: 'File per Questionari generato',
        description: 'Il file Excel è stato scaricato con successo.',
      });
    } catch (error) {
      console.error('Error generating file_per_questionari:', error);
      toast({
        title: 'Errore',
        description: 'Impossibile generare il file Excel.',
        variant: 'destructive',
      });
    } finally {
      setIsExporting(null);
    }
  };

  const handleTabellaGraficiDownload = async () => {
    setIsExporting('tabella_grafici');
    try {
      await generateTabellaGrafici(parsedSurvey);
      toast({
        title: 'Tabella Grafici generata',
        description: 'Il file Excel è stato scaricato con successo.',
      });
    } catch (error) {
      console.error('Error generating tabella_grafici:', error);
      toast({
        title: 'Errore',
        description: 'Impossibile generare la tabella grafici.',
        variant: 'destructive',
      });
    } finally {
      setIsExporting(null);
    }
  };

  const handleChartsDownload = async () => {
    setIsExporting('charts');
    setProgress({ current: 0, total: 0 });

    try {
      await generateChartsZip(parsedSurvey, (current, total) => {
        setProgress({ current, total });
      });
      toast({
        title: 'Grafici generati',
        description: 'Il file ZIP è stato scaricato con successo.',
      });
    } catch (error) {
      console.error('Error generating charts:', error);
      toast({
        title: 'Errore',
        description: 'Impossibile generare i grafici.',
        variant: 'destructive',
      });
    } finally {
      setIsExporting(null);
      setProgress({ current: 0, total: 0 });
    }
  };

  const scaleCount = parsedSurvey.questions.filter(q => q.type === 'scale_1_10_na').length;

  return (
    <div className="glass-card rounded-xl p-4 animate-fade-in">
      <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
        <div>
          <h3 className="font-semibold text-lg flex items-center gap-2">
            <Download className="w-5 h-5 text-primary" />
            Esporta Risultati
          </h3>
          <p className="text-sm text-muted-foreground">
            Scarica i file Excel e i grafici PNG
          </p>
        </div>

        <div className="flex flex-wrap gap-3">
          {/* File per Questionari Excel */}
          <Button
            onClick={handleFilePerQuestionariDownload}
            disabled={isExporting !== null}
            className="gap-2"
          >
            {isExporting === 'file_per_questionari' ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <FileSpreadsheet className="w-4 h-4" />
            )}
            File Questionari
          </Button>

          {/* Tabella Grafici Excel */}
          <Button
            onClick={handleTabellaGraficiDownload}
            disabled={isExporting !== null || scaleCount === 0}
            variant="secondary"
            className="gap-2"
          >
            {isExporting === 'tabella_grafici' ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <Table className="w-4 h-4" />
            )}
            Tabella Grafici
          </Button>

          {/* Charts ZIP */}
          <Button
            onClick={handleChartsDownload}
            disabled={isExporting !== null || scaleCount === 0}
            variant="outline"
            className="gap-2"
          >
            {isExporting === 'charts' ? (
              <>
                <Loader2 className="w-4 h-4 animate-spin" />
                {progress.total > 0 && (
                  <span className="text-xs">
                    {progress.current}/{progress.total}
                  </span>
                )}
              </>
            ) : (
              <>
                <Image className="w-4 h-4" />
                Grafici ZIP
              </>
            )}
          </Button>
        </div>
      </div>
    </div>
  );
}
