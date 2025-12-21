import { useState } from 'react';
import { Download, FileSpreadsheet, Image, Package, Loader2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { useSurveyStore } from '@/store/surveyStore';
import { generateExcelReport } from '@/utils/excelWriter';
import { generateChartsZip } from '@/utils/chartExporter';
import { toast } from '@/hooks/use-toast';

export function DownloadBar() {
  const { parsedSurvey } = useSurveyStore();
  const [isExporting, setIsExporting] = useState<'excel' | 'charts' | null>(null);
  const [progress, setProgress] = useState({ current: 0, total: 0 });

  if (!parsedSurvey) return null;

  const handleExcelDownload = async () => {
    setIsExporting('excel');
    try {
      await generateExcelReport(parsedSurvey);
      toast({
        title: 'Report Excel generato',
        description: 'Il file è stato scaricato con successo.',
      });
    } catch (error) {
      toast({
        title: 'Errore',
        description: 'Impossibile generare il report Excel.',
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
            Scarica il report Excel e i grafici PNG
          </p>
        </div>

        <div className="flex flex-wrap gap-3">
          <Button
            onClick={handleExcelDownload}
            disabled={isExporting !== null}
            className="gap-2"
          >
            {isExporting === 'excel' ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <FileSpreadsheet className="w-4 h-4" />
            )}
            Scarica Excel
          </Button>

          <Button
            onClick={handleChartsDownload}
            disabled={isExporting !== null || scaleCount === 0}
            variant="secondary"
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
                Scarica Grafici ZIP
              </>
            )}
          </Button>
        </div>
      </div>
    </div>
  );
}
