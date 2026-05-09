import actions from '../data/actions.json' with { type: 'json' };
import endings from '../data/endings.json' with { type: 'json' };
import story from '../data/story.json' with { type: 'json' };
import visualStates from '../data/visualStates.json' with { type: 'json' };

export const MAX_ROUNDS = 5;

export const INITIAL_VARIABLES = {
  earthHealth: 45,
  temperature: 45,
  biodiversity: 45,
  communityCare: 35,
  justice: 35,
  waste: 45,
  hope: 40,
  futureMomentum: 0
};

const variableKeys = Object.keys(INITIAL_VARIABLES);

export function getChapter(round) {
  const chapterIndex = Math.min(round, story.chapters.length - 1);
  return story.chapters[chapterIndex];
}

export function getVisualState(earthState) {
  return visualStates.find((state) => state.id === earthState) ?? visualStates[0];
}

export function getPrologue() {
  return story.prologue;
}

export function getInterfaceCopy() {
  return story.archiveAndEndingCardCopy;
}

export function getSuggestedActions() {
  return ['repair', 'plant', 'share', 'listen', 'reduce', 'teach', 'protest', 'compost'];
}

export function getBranchPrompt(round) {
  return story.branchPrompts.find((prompt) => prompt.afterRound === round) ?? null;
}

export function normalizeInput(input) {
  return String(input ?? '').toLowerCase().trim().replace(/\s+/g, ' ');
}

export function findAction(input) {
  const normalized = normalizeInput(input);

  if (!normalized) {
    return null;
  }

  const exactId = actions.find((action) => action.id === normalized);
  if (exactId) return exactId;

  const labelMatch = actions.find((action) => action.label === normalized);
  if (labelMatch) return labelMatch;

  const aliasMatch = actions.find((action) =>
    action.aliases.some((alias) => normalizeInput(alias) === normalized)
  );
  if (aliasMatch) return aliasMatch;

  return actions.find((action) => {
    const candidates = [action.id, action.label, ...action.aliases].map(normalizeInput);
    return candidates.some((candidate) => normalized.includes(candidate));
  }) ?? null;
}

export function applyEffects(variables, effects = {}) {
  return variableKeys.reduce((nextVariables, key) => {
    const nextValue = variables[key] + (effects[key] ?? 0);
    nextVariables[key] = clamp(nextValue, 0, 100);
    return nextVariables;
  }, {});
}

export function determineEarthState(variables, action) {
  if (variables.earthHealth <= 20 || variables.waste >= 80) return 'collapsing';
  if (variables.temperature >= 70) return 'warming';
  if (variables.futureMomentum >= 30 && variables.earthHealth >= 60) return 'turning';
  if (variables.communityCare >= 65) return 'collective';
  if (variables.biodiversity >= 65) return 'blooming';
  if (variables.earthHealth >= 60 && variables.waste <= 40) return 'repairing';
  if (variables.biodiversity <= 35 || variables.hope <= 25) return 'silent';

  if (action?.preferredEarthState === 'warming' && (variables.temperature >= 55 || variables.waste >= 60)) {
    return 'warming';
  }
  if (action?.preferredEarthState === 'collective' && variables.communityCare >= 55) return 'collective';
  if (action?.preferredEarthState === 'blooming' && variables.biodiversity >= 55) return 'blooming';
  if (action?.preferredEarthState === 'repairing' && (variables.earthHealth >= 50 || variables.waste <= 40)) {
    return 'repairing';
  }
  if (action?.preferredEarthState === 'silent' && (variables.biodiversity <= 40 || variables.hope <= 35)) {
    return 'silent';
  }
  return 'fragile';
}

export function selectStoryResponse(action, round, fallbackIndex = 0) {
  if (!action) {
    return story.fallbackResponses[fallbackIndex % story.fallbackResponses.length];
  }

  const response = story.responses.find((item) => item.keyword === action.id);
  if (response) return response;

  return {
    response: `${action.label} changes the page, but this action has no story response yet.`,
    earthVoice: 'This action is recorded in the path.',
    systemFeedback: 'The Earth state changed from the submitted action.',
    nextPrompt: getChapter(round).prompt
  };
}

function getFallbackResponse(id, fallbackIndex = 0) {
  return (
    story.fallbackResponses.find((response) => response.id === id) ??
    story.fallbackResponses[fallbackIndex % story.fallbackResponses.length]
  );
}

export function determineEnding(variables, history) {
  const counts = countActionCategories(history);
  const selected = history.map((item) => item.keyword);
  const hasAny = (keywords) => selected.some((keyword) => keywords.includes(keyword));
  const hasTechnologyPath = hasAny(['technology', 'efficiency', 'build', 'greenwash']);
  const hasCollectivePath = hasAny(['share', 'teach', 'protest', 'organise', 'listen']);
  const hasBloomingPath = hasAny(['plant', 'restore', 'compost', 'protect']);
  const hasRepairPath = hasAny(['repair', 'reduce', 'technology', 'efficiency', 'archive']);

  if (variables.earthHealth <= 30 || variables.waste >= 75 || counts.negative >= 3) {
    return getEnding('collapsing-earth');
  }

  if (
    variables.earthHealth >= 65 &&
    variables.communityCare >= 55 &&
    variables.justice >= 50 &&
    variables.futureMomentum >= 25 &&
    variables.waste <= 45 &&
    counts.positive >= 4 &&
    counts.negative === 0
  ) {
    return getEnding('turning-earth');
  }

  if (variables.communityCare >= 65 && variables.justice >= 50 && hasCollectivePath) {
    return getEnding('collective-earth');
  }

  if (variables.biodiversity >= 65 && variables.earthHealth >= 55 && hasBloomingPath) {
    return getEnding('blooming-earth');
  }

  if (variables.temperature >= 65 || variables.waste >= 65 || counts.negative >= 2) {
    return getEnding('warming-earth');
  }

  if (
    hasRepairPath &&
    (variables.earthHealth >= 55 || variables.waste <= 35 || hasTechnologyPath) &&
    counts.negative <= 1
  ) {
    return getEnding('repairing-earth');
  }

  return getEnding('fragile-earth');
}

export function countActionCategories(history) {
  return history.reduce(
    (counts, item) => {
      if (item.category && counts[item.category] !== undefined) {
        counts[item.category] += 1;
      }
      return counts;
    },
    { positive: 0, neutral: 0, negative: 0 }
  );
}

export function createInitialAppState() {
  const chapter = getChapter(0);
  const prologue = getPrologue();

  return {
    screen: 'landing',
    chapter: chapter.chapter,
    round: 0,
    maxRounds: MAX_ROUNDS,
    earthState: 'fragile',
    variables: { ...INITIAL_VARIABLES },
    history: [],
    currentStoryText: prologue.text,
    currentEarthVoice: 'Choose carefully, not perfectly. The Earth will answer what repeats.',
    currentPrompt: prologue.firstInputPrompt ?? chapter.prompt,
    lastAction: null,
    ending: null,
    feedbackMessage: 'The fragile Earth waits inside the page.',
    systemFeedback: ''
  };
}

export function submitKeyword(currentState, rawInput) {
  const normalizedInput = normalizeInput(rawInput);
  const action = findAction(normalizedInput);
  const previousEarthState = currentState.earthState;
  const fallbackIndex = currentState.history.length;

  if (!action) {
    const response = getFallbackResponse(normalizedInput ? 'fallback-unrecognised' : 'fallback-empty', fallbackIndex);

    return {
      ...currentState,
      currentStoryText: response.response,
      currentEarthVoice: response.earthVoice,
      currentPrompt: response.nextPrompt,
      lastAction: null,
      feedbackMessage: response.systemFeedback,
      systemFeedback: response.systemFeedback
    };
  }

  const nextRound = currentState.round + 1;
  const response = selectStoryResponse(action, currentState.round, fallbackIndex);
  const nextVariables = applyEffects(currentState.variables, action.effects);
  const nextEarthState = determineEarthState(nextVariables, action);

  const historyItem = {
    round: nextRound,
    input: rawInput,
    keyword: action.id,
    label: normalizedInput || action.label,
    category: action.category,
    earthStateBefore: previousEarthState,
    earthStateAfter: nextEarthState,
    responseId: action.id,
    variablesAfter: nextVariables
  };

  const history = [...currentState.history, historyItem];
  const chapter = getChapter(Math.min(nextRound, MAX_ROUNDS - 1));
  const shouldEnd = nextRound >= MAX_ROUNDS;
  const ending = shouldEnd ? determineEnding(nextVariables, history) : null;

  return {
    ...currentState,
    screen: shouldEnd ? 'ending' : 'story',
    chapter: chapter.chapter,
    round: nextRound,
    earthState: ending?.visualState ?? nextEarthState,
    variables: nextVariables,
    history,
    currentStoryText: response.response,
    currentEarthVoice: response.earthVoice,
    currentPrompt: response.nextPrompt,
    lastAction: action,
    ending,
    feedbackMessage: response.systemFeedback,
    systemFeedback: response.systemFeedback
  };
}

function getEnding(id) {
  return endings.find((ending) => ending.id === id) ?? endings[endings.length - 1];
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}
