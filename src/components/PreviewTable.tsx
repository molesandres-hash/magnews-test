import { useMemo } from 'react';
import { ChevronDown, ChevronRight, BarChart3 } from 'lucide-react';
import { useSurveyStore } from '@/store/surveyStore';
import { filterQuestions, groupQuestionsByBlock, getSectionDisplayName } from '@/utils/analytics';
import { getQuestionTypeLabel, getQuestionTypeColor } from '@/utils/questionClassifier';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { cn } from '@/lib/utils';
import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
} from '@/components/ui/collapsible';
import { useState } from 'react';

export function PreviewTable() {
  const { parsedSurvey, filters, selectedQuestionId, setSelectedQuestionId } = useSurveyStore();
  const [expandedBlocks, setExpandedBlocks] = useState<Set<number | null>>(new Set());

  const filteredQuestions = useMemo(() => {
    if (!parsedSurvey) return [];
    return filterQuestions(parsedSurvey.questions, filters);
  }, [parsedSurvey, filters]);

  const groupedQuestions = useMemo(() => {
    return groupQuestionsByBlock(filteredQuestions);
  }, [filteredQuestions]);

  const toggleBlock = (blockId: number | null) => {
    const updated = new Set(expandedBlocks);
    if (updated.has(blockId)) {
      updated.delete(blockId);
    } else {
      updated.add(blockId);
    }
    setExpandedBlocks(updated);
  };

  if (!parsedSurvey) return null;

  const sortedBlocks = Array.from(groupedQuestions.keys()).sort((a, b) => {
    if (a === null) return 1;
    if (b === null) return -1;
    return a - b;
  });

  // Initialize all blocks as expanded on first render
  if (expandedBlocks.size === 0 && sortedBlocks.length > 0) {
    setExpandedBlocks(new Set(sortedBlocks));
  }

  return (
    <div className="glass-card rounded-xl overflow-hidden animate-fade-in">
      <div className="p-4 border-b border-border bg-muted/30">
        <h3 className="font-semibold text-lg">Anteprima Domande</h3>
        <p className="text-sm text-muted-foreground">
          {filteredQuestions.length} domande in {sortedBlocks.length} blocchi
        </p>
      </div>

      <div className="max-h-[600px] overflow-y-auto scrollbar-thin">
        {sortedBlocks.map(blockId => {
          const questions = groupedQuestions.get(blockId) || [];
          const isExpanded = expandedBlocks.has(blockId);

          return (
            <Collapsible
              key={blockId ?? 'null'}
              open={isExpanded}
              onOpenChange={() => toggleBlock(blockId)}
            >
              <CollapsibleTrigger asChild>
                <div className="block-header flex items-center gap-3 cursor-pointer hover:bg-primary/10 transition-colors mx-2 my-2">
                  {isExpanded ? (
                    <ChevronDown className="w-5 h-5 text-primary" />
                  ) : (
                    <ChevronRight className="w-5 h-5 text-muted-foreground" />
                  )}
                  <span className="font-semibold">{getBlockDisplayName(blockId)}</span>
                  <Badge variant="secondary" className="ml-auto">
                    {questions.length} domande
                  </Badge>
                </div>
              </CollapsibleTrigger>

              <CollapsibleContent>
                <div className="divide-y divide-border/50">
                  {questions.map(question => {
                    const analytics = parsedSurvey.scaleAnalytics.get(question.id);
                    const isSelected = selectedQuestionId === question.id;

                    return (
                      <div
                        key={question.id}
                        className={cn(
                          'p-4 cursor-pointer transition-all duration-200',
                          'hover:bg-primary/5',
                          isSelected && 'bg-primary/10 border-l-4 border-l-primary'
                        )}
                        onClick={() => setSelectedQuestionId(question.id)}
                      >
                        <div className="flex items-start gap-3">
                          <div className="flex-1 min-w-0">
                            <div className="flex items-center gap-2 mb-1">
                              {question.questionKey && (
                                <span className="font-mono text-sm font-semibold text-primary">
                                  {question.questionKey}
                                </span>
                              )}
                              <Badge 
                                variant="outline" 
                                className={cn('text-xs', getQuestionTypeColor(question.type))}
                              >
                                {getQuestionTypeLabel(question.type)}
                              </Badge>
                            </div>
                            <p className="text-sm text-foreground line-clamp-2">
                              {question.questionText}
                            </p>
                          </div>

                          {analytics && (
                            <div className="flex items-center gap-2 shrink-0">
                              <div className="text-right">
                                <p className="text-2xl font-bold text-foreground">
                                  {analytics.mean.toFixed(1)}
                                </p>
                                <p className="text-xs text-muted-foreground">media</p>
                              </div>
                              <Button
                                variant="ghost"
                                size="icon"
                                className="text-primary"
                              >
                                <BarChart3 className="w-5 h-5" />
                              </Button>
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </CollapsibleContent>
            </Collapsible>
          );
        })}

        {filteredQuestions.length === 0 && (
          <div className="p-12 text-center text-muted-foreground">
            <p>Nessuna domanda trovata con i filtri attivi.</p>
          </div>
        )}
      </div>
    </div>
  );
}
