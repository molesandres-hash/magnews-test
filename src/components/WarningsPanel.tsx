import { AlertTriangle, Info, X } from 'lucide-react';
import { useSurveyStore } from '@/store/surveyStore';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { Button } from '@/components/ui/button';
import { useState } from 'react';

export function WarningsPanel() {
  const { parsedSurvey } = useSurveyStore();
  const [isVisible, setIsVisible] = useState(true);

  if (!parsedSurvey || parsedSurvey.metadata.warnings.length === 0 || !isVisible) {
    return null;
  }

  return (
    <Alert variant="default" className="border-warning/50 bg-warning/5 animate-fade-in">
      <AlertTriangle className="h-4 w-4 text-warning" />
      <AlertTitle className="text-warning flex items-center justify-between">
        Avvisi ({parsedSurvey.metadata.warnings.length})
        <Button
          variant="ghost"
          size="icon"
          className="h-6 w-6 text-warning hover:text-warning/80"
          onClick={() => setIsVisible(false)}
        >
          <X className="h-4 w-4" />
        </Button>
      </AlertTitle>
      <AlertDescription className="mt-2">
        <ul className="list-disc list-inside space-y-1 text-sm text-muted-foreground">
          {parsedSurvey.metadata.warnings.map((warning, idx) => (
            <li key={idx}>{warning}</li>
          ))}
        </ul>
      </AlertDescription>
    </Alert>
  );
}
