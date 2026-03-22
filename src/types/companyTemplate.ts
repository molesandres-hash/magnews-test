export interface CompanyTemplate {
  id: string;
  name: string;
  primaryColor: string;
  secondaryColor: string;
  accentColor: string;
  fontFamily: string;
  logoBase64?: string;
  createdAt: string;
}

export interface TemplateState {
  templates: CompanyTemplate[];
  activeTemplateId: string | null;
}

export interface TemplateActions {
  addTemplate: (template: CompanyTemplate) => void;
  updateTemplate: (id: string, updates: Partial<CompanyTemplate>) => void;
  deleteTemplate: (id: string) => void;
  setActiveTemplateId: (id: string | null) => void;
  getActiveTemplate: () => CompanyTemplate | null;
}
