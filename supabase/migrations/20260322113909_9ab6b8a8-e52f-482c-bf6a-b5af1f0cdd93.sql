CREATE TABLE public.company_templates (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  primary_color text NOT NULL DEFAULT '#2563EB',
  secondary_color text NOT NULL DEFAULT '#E0E7FF',
  accent_color text NOT NULL DEFAULT '#1E40AF',
  font_family text NOT NULL DEFAULT 'Arial',
  logo_base64 text,
  created_at timestamptz NOT NULL DEFAULT now()
);

ALTER TABLE public.company_templates ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Allow public read" ON public.company_templates FOR SELECT TO anon, authenticated USING (true);
CREATE POLICY "Allow public insert" ON public.company_templates FOR INSERT TO anon, authenticated WITH CHECK (true);
CREATE POLICY "Allow public update" ON public.company_templates FOR UPDATE TO anon, authenticated USING (true) WITH CHECK (true);
CREATE POLICY "Allow public delete" ON public.company_templates FOR DELETE TO anon, authenticated USING (true);