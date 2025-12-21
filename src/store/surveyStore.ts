import { create } from 'zustand';
import type { SurveyStore, FilterState } from '@/types/survey';

const initialFilters: FilterState = {
  blocks: [],
  questionTypes: [],
  searchText: '',
};

export const useSurveyStore = create<SurveyStore>((set) => ({
  parsedSurvey: null,
  isLoading: false,
  error: null,
  filters: initialFilters,
  selectedQuestionId: null,

  setParsedSurvey: (survey) => set({ parsedSurvey: survey, error: null }),
  setLoading: (loading) => set({ isLoading: loading }),
  setError: (error) => set({ error, isLoading: false }),
  setFilters: (filters) => set((state) => ({ 
    filters: { ...state.filters, ...filters } 
  })),
  setSelectedQuestionId: (id) => set({ selectedQuestionId: id }),
  reset: () => set({ 
    parsedSurvey: null, 
    isLoading: false, 
    error: null, 
    filters: initialFilters,
    selectedQuestionId: null 
  }),
}));
