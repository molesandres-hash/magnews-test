import React, { useCallback, useState } from 'react';
import { Upload, FileSpreadsheet, AlertCircle } from 'lucide-react';
import { cn } from '@/lib/utils';
import { parseCSVFile } from '@/utils/csvParser';
import { useSurveyStore } from '@/store/surveyStore';

export function UploadDropzone() {
  const [isDragging, setIsDragging] = useState(false);
  const { setLoading, setError, setParsedSurvey } = useSurveyStore();

  const handleFile = useCallback(async (file: File) => {
    if (!file.name.endsWith('.csv')) {
      setError('Per favore carica un file CSV.');
      return;
    }

    setLoading(true);
    setError(null);

    try {
      const survey = await parseCSVFile(file);
      setParsedSurvey(survey);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Errore durante l\'elaborazione del file.');
    }
  }, [setLoading, setError, setParsedSurvey]);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);

    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  }, [handleFile]);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleInputChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file);
  }, [handleFile]);

  return (
    <div
      className={cn(
        'upload-zone cursor-pointer relative group',
        isDragging && 'upload-zone-active'
      )}
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onClick={() => document.getElementById('file-input')?.click()}
    >
      <input
        id="file-input"
        type="file"
        accept=".csv"
        className="hidden"
        onChange={handleInputChange}
      />
      
      <div className="flex flex-col items-center gap-6">
        <div className={cn(
          'w-20 h-20 rounded-2xl flex items-center justify-center transition-all duration-300',
          'bg-gradient-to-br from-primary/10 to-accent/10',
          'group-hover:from-primary/20 group-hover:to-accent/20',
          'group-hover:scale-110'
        )}>
          <FileSpreadsheet className="w-10 h-10 text-primary" />
        </div>
        
        <div className="text-center">
          <h3 className="text-xl font-semibold text-foreground mb-2">
            Trascina qui il file CSV
          </h3>
          <p className="text-muted-foreground">
            oppure <span className="text-primary font-medium">clicca per selezionare</span>
          </p>
        </div>

        <div className="flex items-center gap-2 text-sm text-muted-foreground">
          <Upload className="w-4 h-4" />
          <span>Formato supportato: CSV (export Magnews)</span>
        </div>
      </div>

      {isDragging && (
        <div className="absolute inset-0 bg-primary/5 rounded-2xl flex items-center justify-center backdrop-blur-sm">
          <p className="text-lg font-medium text-primary">Rilascia per caricare</p>
        </div>
      )}
    </div>
  );
}
