import { useMemo } from 'react';
import { Search, Filter, X } from 'lucide-react';
import { Input } from '@/components/ui/input';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { useSurveyStore } from '@/store/surveyStore';
import { getUniqueBlocks } from '@/utils/analytics';
import { getQuestionTypeLabel } from '@/utils/questionClassifier';
import type { QuestionType } from '@/types/survey';

const QUESTION_TYPES: QuestionType[] = ['scale_1_10_na', 'open_text', 'closed_single', 'closed_binary', 'closed_multi'];

export function FilterBar() {
  const { parsedSurvey, filters, setFilters } = useSurveyStore();

  const blocks = useMemo(() => {
    if (!parsedSurvey) return [];
    return getUniqueBlocks(parsedSurvey.questions);
  }, [parsedSurvey]);

  const toggleBlock = (blockId: number) => {
    const current = filters.blocks;
    const updated = current.includes(blockId)
      ? current.filter(b => b !== blockId)
      : [...current, blockId];
    setFilters({ blocks: updated });
  };

  const toggleType = (type: QuestionType) => {
    const current = filters.questionTypes;
    const updated = current.includes(type)
      ? current.filter(t => t !== type)
      : [...current, type];
    setFilters({ questionTypes: updated });
  };

  const clearFilters = () => {
    setFilters({ blocks: [], questionTypes: [], searchText: '' });
  };

  const hasActiveFilters = filters.blocks.length > 0 || filters.questionTypes.length > 0 || filters.searchText;

  if (!parsedSurvey) return null;

  return (
    <div className="glass-card rounded-xl p-4 space-y-4 animate-fade-in">
      {/* Search */}
      <div className="relative">
        <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
        <Input
          placeholder="Cerca domande..."
          value={filters.searchText}
          onChange={(e) => setFilters({ searchText: e.target.value })}
          className="pl-10"
        />
      </div>

      {/* Filter chips */}
      <div className="flex flex-wrap gap-2 items-center">
        <div className="flex items-center gap-1 text-sm text-muted-foreground">
          <Filter className="w-4 h-4" />
          <span>Blocchi:</span>
        </div>
        
        {blocks.map(blockId => (
          <Badge
            key={blockId}
            variant={filters.blocks.includes(blockId) ? 'default' : 'outline'}
            className="cursor-pointer transition-all hover:scale-105"
            onClick={() => toggleBlock(blockId)}
          >
            Blocco {blockId}
          </Badge>
        ))}
      </div>

      <div className="flex flex-wrap gap-2 items-center">
        <div className="flex items-center gap-1 text-sm text-muted-foreground">
          <Filter className="w-4 h-4" />
          <span>Tipo:</span>
        </div>
        
        {QUESTION_TYPES.map(type => (
          <Badge
            key={type}
            variant={filters.questionTypes.includes(type) ? 'default' : 'outline'}
            className="cursor-pointer transition-all hover:scale-105"
            onClick={() => toggleType(type)}
          >
            {getQuestionTypeLabel(type)}
          </Badge>
        ))}
      </div>

      {hasActiveFilters && (
        <div className="flex justify-end">
          <Button
            variant="ghost"
            size="sm"
            onClick={clearFilters}
            className="text-muted-foreground hover:text-foreground"
          >
            <X className="w-4 h-4 mr-1" />
            Pulisci filtri
          </Button>
        </div>
      )}
    </div>
  );
}
