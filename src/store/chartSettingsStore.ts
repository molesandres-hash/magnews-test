import { create } from 'zustand';
import type { ChartSettings } from '@/types/chartSettings';
import { DEFAULT_CHART_SETTINGS } from '@/types/chartSettings';

const STORAGE_KEY = 'magnews_chart_settings';

function load(): ChartSettings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return DEFAULT_CHART_SETTINGS;
    return { ...DEFAULT_CHART_SETTINGS, ...JSON.parse(raw) };
  } catch {
    return DEFAULT_CHART_SETTINGS;
  }
}

function save(settings: ChartSettings) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

interface ChartSettingsStore {
  settings: ChartSettings;
  updateSettings: (patch: Partial<ChartSettings>) => void;
  setQuestionOverride: (questionId: string, patch: Partial<ChartSettings>) => void;
  clearQuestionOverride: (questionId: string) => void;
  clearAllOverrides: () => void;
  getEffectiveSettings: (questionId?: string) => ChartSettings;
  resetToDefaults: () => void;
}

export const useChartSettingsStore = create<ChartSettingsStore>((set, get) => ({
  settings: load(),

  updateSettings: (patch) =>
    set((state) => {
      const updated = { ...state.settings, ...patch };
      save(updated);
      return { settings: updated };
    }),

  setQuestionOverride: (questionId, patch) =>
    set((state) => {
      const overrides = { ...state.settings.perQuestionOverrides, [questionId]: patch };
      const updated = { ...state.settings, perQuestionOverrides: overrides };
      save(updated);
      return { settings: updated };
    }),

  clearQuestionOverride: (questionId) =>
    set((state) => {
      const overrides = { ...state.settings.perQuestionOverrides };
      delete overrides[questionId];
      const updated = { ...state.settings, perQuestionOverrides: overrides };
      save(updated);
      return { settings: updated };
    }),

  clearAllOverrides: () =>
    set((state) => {
      const updated = { ...state.settings, perQuestionOverrides: {} };
      save(updated);
      return { settings: updated };
    }),

  getEffectiveSettings: (questionId?: string) => {
    const s = get().settings;
    if (!questionId || !s.perQuestionOverrides[questionId]) return s;
    return { ...s, ...s.perQuestionOverrides[questionId] };
  },

  resetToDefaults: () => {
    save(DEFAULT_CHART_SETTINGS);
    set({ settings: DEFAULT_CHART_SETTINGS });
  },
}));
