import { create } from 'zustand';
import { supabase } from '@/integrations/supabase/client';
import type { CompanyTemplate, TemplateState, TemplateActions } from '@/types/companyTemplate';

const ACTIVE_KEY = 'magnews_active_template';

function loadActiveId(): string | null {
  return localStorage.getItem(ACTIVE_KEY);
}

function saveActiveId(id: string | null) {
  if (id) localStorage.setItem(ACTIVE_KEY, id);
  else localStorage.removeItem(ACTIVE_KEY);
}

/** Map DB row → CompanyTemplate */
function rowToTemplate(row: any): CompanyTemplate {
  return {
    id: row.id,
    name: row.name,
    primaryColor: row.primary_color,
    secondaryColor: row.secondary_color,
    accentColor: row.accent_color,
    fontFamily: row.font_family,
    logoBase64: row.logo_base64 ?? undefined,
    createdAt: row.created_at,
  };
}

export const useTemplateStore = create<TemplateState & TemplateActions & { loading: boolean; fetchTemplates: () => Promise<void> }>((set, get) => ({
  templates: [],
  activeTemplateId: loadActiveId(),
  loading: false,

  fetchTemplates: async () => {
    set({ loading: true });
    try {
      const { data, error } = await supabase
        .from('company_templates')
        .select('*')
        .order('created_at', { ascending: true });
      if (error) throw error;
      set({ templates: (data ?? []).map(rowToTemplate) });
    } catch (err) {
      console.error('Failed to fetch templates:', err);
    } finally {
      set({ loading: false });
    }
  },

  addTemplate: async (template) => {
    // Optimistic update
    set((state) => ({ templates: [...state.templates, template] }));
    try {
      const { error } = await supabase.from('company_templates').insert({
        id: template.id,
        name: template.name,
        primary_color: template.primaryColor,
        secondary_color: template.secondaryColor,
        accent_color: template.accentColor,
        font_family: template.fontFamily,
        logo_base64: template.logoBase64 ?? null,
        created_at: template.createdAt,
      });
      if (error) throw error;
    } catch (err) {
      console.error('Failed to save template:', err);
      // Rollback
      set((state) => ({ templates: state.templates.filter(t => t.id !== template.id) }));
    }
  },

  updateTemplate: async (id, updates) => {
    const prev = get().templates;
    set((state) => ({
      templates: state.templates.map(t => t.id === id ? { ...t, ...updates } : t),
    }));
    try {
      const dbUpdates: Record<string, any> = {};
      if (updates.name !== undefined) dbUpdates.name = updates.name;
      if (updates.primaryColor !== undefined) dbUpdates.primary_color = updates.primaryColor;
      if (updates.secondaryColor !== undefined) dbUpdates.secondary_color = updates.secondaryColor;
      if (updates.accentColor !== undefined) dbUpdates.accent_color = updates.accentColor;
      if (updates.fontFamily !== undefined) dbUpdates.font_family = updates.fontFamily;
      if (updates.logoBase64 !== undefined) dbUpdates.logo_base64 = updates.logoBase64 ?? null;

      const { error } = await supabase.from('company_templates').update(dbUpdates).eq('id', id);
      if (error) throw error;
    } catch (err) {
      console.error('Failed to update template:', err);
      set({ templates: prev });
    }
  },

  deleteTemplate: async (id) => {
    const prev = get().templates;
    const prevActive = get().activeTemplateId;
    set((state) => {
      const newActiveId = state.activeTemplateId === id ? null : state.activeTemplateId;
      if (newActiveId !== state.activeTemplateId) saveActiveId(newActiveId);
      return { templates: state.templates.filter(t => t.id !== id), activeTemplateId: newActiveId };
    });
    try {
      const { error } = await supabase.from('company_templates').delete().eq('id', id);
      if (error) throw error;
    } catch (err) {
      console.error('Failed to delete template:', err);
      set({ templates: prev, activeTemplateId: prevActive });
    }
  },

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
