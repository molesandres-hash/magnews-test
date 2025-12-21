import { RotateCcw, FileSpreadsheet } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { useSurveyStore } from '@/store/surveyStore';

export function Header() {
  const { parsedSurvey, reset } = useSurveyStore();

  return (
    <header className="border-b border-border/50 bg-card/50 backdrop-blur-xl sticky top-0 z-50">
      <div className="container mx-auto px-4 py-4">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-primary to-accent flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5 text-primary-foreground" />
            </div>
            <div>
              <h1 className="font-bold text-lg">Magnews Survey Analyzer</h1>
              {parsedSurvey && (
                <p className="text-xs text-muted-foreground">
                  {parsedSurvey.metadata.fileName}
                </p>
              )}
            </div>
          </div>

          {parsedSurvey && (
            <Button
              variant="outline"
              size="sm"
              onClick={reset}
              className="gap-2"
            >
              <RotateCcw className="w-4 h-4" />
              Nuovo File
            </Button>
          )}
        </div>
      </div>
    </header>
  );
}
