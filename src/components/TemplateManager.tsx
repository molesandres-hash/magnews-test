import { useState, useEffect } from 'react';
import { Settings, Plus, Pencil, Trash2, Save, X, Loader2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Sheet, SheetContent, SheetHeader, SheetTitle, SheetTrigger } from '@/components/ui/sheet';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { useTemplateStore } from '@/store/templateStore';
import type { CompanyTemplate } from '@/types/companyTemplate';

const FONT_OPTIONS = ['Arial', 'Calibri', 'Helvetica', 'Times New Roman', 'Georgia'];

interface TemplateFormData {
  name: string;
  primaryColor: string;
  secondaryColor: string;
  accentColor: string;
  fontFamily: string;
  logoBase64?: string;
}

const defaultForm: TemplateFormData = {
  name: '',
  primaryColor: '#2563EB',
  secondaryColor: '#E0E7FF',
  accentColor: '#1E40AF',
  fontFamily: 'Arial',
};

export function TemplateManager() {
  const { templates, activeTemplateId, addTemplate, updateTemplate, deleteTemplate, setActiveTemplateId } = useTemplateStore();
  const [isEditing, setIsEditing] = useState<string | null>(null);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState<TemplateFormData>(defaultForm);

  const handleSave = () => {
    if (!form.name.trim()) return;
    if (isEditing) {
      updateTemplate(isEditing, form);
      setIsEditing(null);
    } else {
      const template: CompanyTemplate = {
        ...form,
        id: crypto.randomUUID(),
        createdAt: new Date().toISOString(),
      };
      addTemplate(template);
    }
    setForm(defaultForm);
    setShowForm(false);
  };

  const handleEdit = (t: CompanyTemplate) => {
    setForm({ name: t.name, primaryColor: t.primaryColor, secondaryColor: t.secondaryColor, accentColor: t.accentColor, fontFamily: t.fontFamily, logoBase64: t.logoBase64 });
    setIsEditing(t.id);
    setShowForm(true);
  };

  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => setForm(f => ({ ...f, logoBase64: reader.result as string }));
    reader.readAsDataURL(file);
  };

  const handleCancel = () => {
    setForm(defaultForm);
    setIsEditing(null);
    setShowForm(false);
  };

  return (
    <Sheet>
      <SheetTrigger asChild>
        <Button variant="outline" size="sm" className="gap-2">
          <Settings className="w-4 h-4" />
          Impostazioni Azienda
        </Button>
      </SheetTrigger>
      <SheetContent className="overflow-y-auto">
        <SheetHeader>
          <SheetTitle>Impostazioni Azienda</SheetTitle>
        </SheetHeader>

        <div className="mt-6 space-y-6">
          {/* Active template selector */}
          <div className="space-y-2">
            <Label>Template attivo</Label>
            <Select value={activeTemplateId || '__none__'} onValueChange={(v) => setActiveTemplateId(v === '__none__' ? null : v)}>
              <SelectTrigger>
                <SelectValue placeholder="Nessuno (default)" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="__none__">Nessuno (default)</SelectItem>
                {templates.map(t => (
                  <SelectItem key={t.id} value={t.id}>
                    <div className="flex items-center gap-2">
                      <div className="w-3 h-3 rounded-full border" style={{ backgroundColor: t.primaryColor }} />
                      {t.name}
                    </div>
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>

          {/* Template list */}
          <div className="space-y-2">
            <div className="flex items-center justify-between">
              <Label>Template salvati</Label>
              {!showForm && (
                <Button size="sm" variant="outline" onClick={() => setShowForm(true)} className="gap-1">
                  <Plus className="w-3 h-3" /> Nuovo
                </Button>
              )}
            </div>

            {templates.length === 0 && !showForm && (
              <p className="text-sm text-muted-foreground py-4 text-center">Nessun template salvato</p>
            )}

            {templates.map(t => (
              <div key={t.id} className="flex items-center gap-2 p-3 rounded-lg border bg-card">
                <div className="flex gap-1">
                  <div className="w-4 h-4 rounded" style={{ backgroundColor: t.primaryColor }} />
                  <div className="w-4 h-4 rounded" style={{ backgroundColor: t.secondaryColor }} />
                  <div className="w-4 h-4 rounded" style={{ backgroundColor: t.accentColor }} />
                </div>
                <span className="flex-1 text-sm font-medium truncate">{t.name}</span>
                <Button size="icon" variant="ghost" className="h-7 w-7" onClick={() => handleEdit(t)}>
                  <Pencil className="w-3 h-3" />
                </Button>
                <Button size="icon" variant="ghost" className="h-7 w-7 text-destructive" onClick={() => deleteTemplate(t.id)}>
                  <Trash2 className="w-3 h-3" />
                </Button>
              </div>
            ))}
          </div>

          {/* Form */}
          {showForm && (
            <div className="space-y-4 p-4 rounded-lg border bg-muted/30">
              <h4 className="font-semibold text-sm">{isEditing ? 'Modifica Template' : 'Nuovo Template'}</h4>

              <div className="space-y-1">
                <Label htmlFor="tpl-name">Nome azienda</Label>
                <Input id="tpl-name" value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))} placeholder="Es. Acme Corp" />
              </div>

              <div className="grid grid-cols-3 gap-3">
                <div className="space-y-1">
                  <Label htmlFor="tpl-primary" className="text-xs">Primario</Label>
                  <input id="tpl-primary" type="color" value={form.primaryColor} onChange={e => setForm(f => ({ ...f, primaryColor: e.target.value }))} className="w-full h-9 rounded cursor-pointer" />
                </div>
                <div className="space-y-1">
                  <Label htmlFor="tpl-secondary" className="text-xs">Secondario</Label>
                  <input id="tpl-secondary" type="color" value={form.secondaryColor} onChange={e => setForm(f => ({ ...f, secondaryColor: e.target.value }))} className="w-full h-9 rounded cursor-pointer" />
                </div>
                <div className="space-y-1">
                  <Label htmlFor="tpl-accent" className="text-xs">Accento</Label>
                  <input id="tpl-accent" type="color" value={form.accentColor} onChange={e => setForm(f => ({ ...f, accentColor: e.target.value }))} className="w-full h-9 rounded cursor-pointer" />
                </div>
              </div>

              <div className="space-y-1">
                <Label>Font</Label>
                <Select value={form.fontFamily} onValueChange={v => setForm(f => ({ ...f, fontFamily: v }))}>
                  <SelectTrigger><SelectValue /></SelectTrigger>
                  <SelectContent>
                    {FONT_OPTIONS.map(f => <SelectItem key={f} value={f}>{f}</SelectItem>)}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-1">
                <Label>Logo (PNG/JPG)</Label>
                <Input type="file" accept="image/png,image/jpeg" onChange={handleLogoUpload} />
                {form.logoBase64 && (
                  <div className="flex items-center gap-2 mt-1">
                    <img src={form.logoBase64} alt="Logo preview" className="h-8 object-contain" />
                    <Button size="sm" variant="ghost" onClick={() => setForm(f => ({ ...f, logoBase64: undefined }))}>
                      <X className="w-3 h-3" />
                    </Button>
                  </div>
                )}
              </div>

              <div className="flex gap-2">
                <Button onClick={handleSave} size="sm" className="gap-1" disabled={!form.name.trim()}>
                  <Save className="w-3 h-3" /> Salva
                </Button>
                <Button onClick={handleCancel} size="sm" variant="outline" className="gap-1">
                  <X className="w-3 h-3" /> Annulla
                </Button>
              </div>
            </div>
          )}
        </div>
      </SheetContent>
    </Sheet>
  );
}
