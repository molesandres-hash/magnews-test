import { create } from 'zustand';
import type { CompanyTemplate, TemplateState, TemplateActions } from '@/types/companyTemplate';

const STORAGE_KEY = 'magnews_templates';
const ACTIVE_KEY = 'magnews_active_template';

function loadTemplates(): CompanyTemplate[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch { return []; }
}

function saveTemplates(templates: CompanyTemplate[]) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(templates));
}

function loadActiveId(): string | null {
  return localStorage.getItem(ACTIVE_KEY);
}

function saveActiveId(id: string | null) {
  if (id) localStorage.setItem(ACTIVE_KEY, id);
  else localStorage.removeItem(ACTIVE_KEY);
}

export const useTemplateStore = create<TemplateState & TemplateActions>((set, get) => ({
  templates: loadTemplates(),
  activeTemplateId: loadActiveId(),

  addTemplate: (template) => set((state) => {
    const updated = [...state.templates, template];
    saveTemplates(updated);
    return { templates: updated };
  }),

  updateTemplate: (id, updates) => set((state) => {
    const updated = state.templates.map(t => t.id === id ? { ...t, ...updates } : t);
    saveTemplates(updated);
    return { templates: updated };
  }),

  deleteTemplate: (id) => set((state) => {
    const updated = state.templates.filter(t => t.id !== id);
    saveTemplates(updated);
    const newActiveId = state.activeTemplateId === id ? null : state.activeTemplateId;
    if (newActiveId !== state.activeTemplateId) saveActiveId(newActiveId);
    return { templates: updated, activeTemplateId: newActiveId };
  }),

  setActiveTemplateId: (id) => {
    saveActiveId(id);
    set({ activeTemplateId: id });
  },

  getActiveTemplate: () => {
    const state = get();
    if (!state.activeTemplateId) return null;
    return state.templates.find(t => t.id === state.activeTemplateId) || null;
  },
}));
