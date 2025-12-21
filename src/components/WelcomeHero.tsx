import { FileSpreadsheet, BarChart3, Download, Zap } from 'lucide-react';

const features = [
  {
    icon: FileSpreadsheet,
    title: 'Upload CSV',
    description: 'Carica il tuo export Magnews in formato CSV. Rileva automaticamente delimitatore e codifica.',
  },
  {
    icon: BarChart3,
    title: 'Analisi Automatica',
    description: 'Identifica domande scala, aperte e chiuse. Calcola medie e distribuzioni.',
  },
  {
    icon: Download,
    title: 'Export Professionale',
    description: 'Genera report Excel strutturato e grafici PNG pronti per PowerPoint.',
  },
  {
    icon: Zap,
    title: '100% Client-Side',
    description: 'Tutto funziona nel browser. I tuoi dati non lasciano mai il tuo computer.',
  },
];

export function WelcomeHero() {
  return (
    <div className="text-center max-w-3xl mx-auto mb-12 animate-fade-in">
      <div className="inline-flex items-center gap-2 px-4 py-2 rounded-full bg-primary/10 text-primary text-sm font-medium mb-6">
        <Zap className="w-4 h-4" />
        Analisi Survey Professionale
      </div>
      
      <h1 className="text-4xl md:text-5xl font-bold mb-4 tracking-tight">
        <span className="gradient-text">Magnews</span> Survey Analyzer
      </h1>
      
      <p className="text-lg text-muted-foreground mb-12">
        Trasforma i tuoi export CSV in report Excel e grafici professionali. 
        Analisi automatica delle domande scala, aperte e chiuse.
      </p>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-left">
        {features.map((feature, idx) => (
          <div 
            key={feature.title}
            className="glass-card rounded-xl p-4 animate-slide-up"
            style={{ animationDelay: `${idx * 100}ms` }}
          >
            <feature.icon className="w-8 h-8 text-primary mb-3" />
            <h3 className="font-semibold text-sm mb-1">{feature.title}</h3>
            <p className="text-xs text-muted-foreground">{feature.description}</p>
          </div>
        ))}
      </div>
    </div>
  );
}
