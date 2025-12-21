import type { ParsedSurvey, QuestionInfo, ScaleAnalytics } from '@/types/survey';

/**
 * Get all unique block IDs from questions
 */
export function getUniqueBlocks(questions: QuestionInfo[]): number[] {
  const blocks = new Set<number>();
  questions.forEach(q => {
    if (q.blockId !== null) {
      blocks.add(q.blockId);
    }
  });
  return Array.from(blocks).sort((a, b) => a - b);
}

/**
 * Group questions by block ID
 */
export function groupQuestionsByBlock(questions: QuestionInfo[]): Map<number | null, QuestionInfo[]> {
  const groups = new Map<number | null, QuestionInfo[]>();
  
  questions.forEach(q => {
    const key = q.blockId;
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key)!.push(q);
  });

  return groups;
}

/**
 * Get block display name
 */
export function getBlockDisplayName(blockId: number | null): string {
  if (blockId === null) return 'Domande senza numerazione';
  return `Blocco ${blockId}`;
}

/**
 * Calculate summary statistics for a block of scale questions
 */
export function getBlockSummary(
  blockId: number | null,
  questions: QuestionInfo[],
  scaleAnalytics: Map<string, ScaleAnalytics>
): {
  questionCount: number;
  averageMean: number;
  highestMean: { questionId: string; mean: number; text: string } | null;
  lowestMean: { questionId: string; mean: number; text: string } | null;
} {
  const blockQuestions = questions.filter(
    q => q.blockId === blockId && q.type === 'scale_1_10_na'
  );

  if (blockQuestions.length === 0) {
    return {
      questionCount: 0,
      averageMean: 0,
      highestMean: null,
      lowestMean: null,
    };
  }

  let totalMean = 0;
  let highest: { questionId: string; mean: number; text: string } | null = null;
  let lowest: { questionId: string; mean: number; text: string } | null = null;

  blockQuestions.forEach(q => {
    const analytics = scaleAnalytics.get(q.id);
    if (!analytics) return;

    totalMean += analytics.mean;

    if (!highest || analytics.mean > highest.mean) {
      highest = { questionId: q.id, mean: analytics.mean, text: q.questionText };
    }
    if (!lowest || analytics.mean < lowest.mean) {
      lowest = { questionId: q.id, mean: analytics.mean, text: q.questionText };
    }
  });

  return {
    questionCount: blockQuestions.length,
    averageMean: Math.round((totalMean / blockQuestions.length) * 100) / 100,
    highestMean: highest,
    lowestMean: lowest,
  };
}

/**
 * Filter questions based on filter state
 */
export function filterQuestions(
  questions: QuestionInfo[],
  filters: { blocks: number[]; questionTypes: string[]; searchText: string }
): QuestionInfo[] {
  return questions.filter(q => {
    // Block filter
    if (filters.blocks.length > 0) {
      if (q.blockId === null && !filters.blocks.includes(-1)) return false;
      if (q.blockId !== null && !filters.blocks.includes(q.blockId)) return false;
    }

    // Type filter
    if (filters.questionTypes.length > 0 && !filters.questionTypes.includes(q.type)) {
      return false;
    }

    // Search filter
    if (filters.searchText) {
      const searchLower = filters.searchText.toLowerCase();
      const matchesText = q.questionText.toLowerCase().includes(searchLower);
      const matchesKey = q.questionKey?.toLowerCase().includes(searchLower);
      if (!matchesText && !matchesKey) return false;
    }

    return true;
  });
}

/**
 * Get overall survey statistics
 */
export function getSurveyStats(survey: ParsedSurvey): {
  totalQuestions: number;
  scaleQuestions: number;
  openQuestions: number;
  closedQuestions: number;
  blocks: number;
  averageCompletionRate: number;
} {
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na').length;
  const openQuestions = survey.questions.filter(q => q.type === 'open_text').length;
  const closedQuestions = survey.questions.filter(
    q => q.type === 'closed_single' || q.type === 'closed_binary' || q.type === 'closed_multi'
  ).length;

  const blocks = getUniqueBlocks(survey.questions).length;

  // Calculate average completion rate for scale questions
  let totalValidRatio = 0;
  let scaleCount = 0;
  survey.scaleAnalytics.forEach(analytics => {
    totalValidRatio += analytics.validResponses / analytics.totalRespondents;
    scaleCount++;
  });

  const averageCompletionRate = scaleCount > 0 
    ? Math.round((totalValidRatio / scaleCount) * 100) 
    : 0;

  return {
    totalQuestions: survey.questions.length,
    scaleQuestions,
    openQuestions,
    closedQuestions,
    blocks,
    averageCompletionRate,
  };
}
