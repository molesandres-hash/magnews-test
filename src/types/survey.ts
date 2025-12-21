export type QuestionType = 
  | 'scale_1_10_na' 
  | 'open_text' 
  | 'closed_single' 
  | 'closed_binary' 
  | 'closed_multi' 
  | 'unknown';

export interface NormalizedHeader {
  rawHeader: string;
  cleanedHeader: string;
  valueSource: 'values' | 'labels' | 'unknown';
}

export interface QuestionInfo {
  id: string;
  rawHeader: string;
  cleanedHeader: string;
  questionKey: string | null;
  questionText: string;
  blockId: number | null;
  subId: number;
  type: QuestionType;
  valueSource: 'values' | 'labels' | 'unknown';
  valuesColumnIndex: number;
  labelsColumnIndex?: number;
}

export interface Respondent {
  id: string;
  displayName: string;
  originalData: Record<string, string>;
  status: string;
  isCompleted: boolean;
  isTestSession: boolean;
}

export interface ScaleAnalytics {
  questionId: string;
  mean: number;
  counts: Record<string, number>; // "10", "9", ..., "1", "N/A"
  totalRespondents: number;
  validResponses: number;
  respondentValues: Record<string, number | null>; // respondentId -> value or null
}

export interface ClosedAnalytics {
  questionId: string;
  options: Array<{
    option: string;
    count: number;
    percent: number;
  }>;
  totalRespondents: number;
}

export interface OpenAnalytics {
  questionId: string;
  responses: Array<{
    respondentId: string;
    respondentName: string;
    answer: string;
  }>;
  filledCount: number;
  emptyCount: number;
}

export interface ParsedSurvey {
  rawData: string[][];
  headers: string[];
  respondents: Respondent[];
  questions: QuestionInfo[];
  scaleAnalytics: Map<string, ScaleAnalytics>;
  closedAnalytics: Map<string, ClosedAnalytics>;
  openAnalytics: Map<string, OpenAnalytics>;
  metadata: {
    totalRows: number;
    completedCount: number;
    excludedCount: number;
    testSessionCount: number;
    warnings: string[];
    parsedAt: Date;
    fileName: string;
  };
}

export interface FilterState {
  blocks: number[];
  questionTypes: QuestionType[];
  searchText: string;
}

export interface SurveyStore {
  parsedSurvey: ParsedSurvey | null;
  isLoading: boolean;
  error: string | null;
  filters: FilterState;
  selectedQuestionId: string | null;
  setParsedSurvey: (survey: ParsedSurvey | null) => void;
  setLoading: (loading: boolean) => void;
  setError: (error: string | null) => void;
  setFilters: (filters: Partial<FilterState>) => void;
  setSelectedQuestionId: (id: string | null) => void;
  reset: () => void;
}
