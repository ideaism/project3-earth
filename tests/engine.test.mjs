import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import {
  createInitialAppState,
  determineEarthState,
  submitKeyword
} from '../src/engine.js';

function storyState() {
  return {
    ...createInitialAppState(),
    screen: 'story'
  };
}

function playPath(path) {
  return path.reduce((state, keyword) => submitKeyword(state, keyword), storyState());
}

describe('story engine', () => {
  it('advances the story and records history for a valid action', () => {
    const initialState = storyState();
    const nextState = submitKeyword(initialState, 'repair');

    assert.equal(nextState.round, 1);
    assert.equal(nextState.history.length, 1);
    assert.equal(nextState.history[0].keyword, 'repair');
    assert.equal(nextState.history[0].label, 'repair');
    assert.notDeepEqual(nextState.variables, initialState.variables);
  });

  it('does not advance or record empty and unknown input', () => {
    const initialState = storyState();

    for (const input of ['', '   ', 'xyz']) {
      const nextState = submitKeyword(initialState, input);
      assert.equal(nextState.round, initialState.round);
      assert.equal(nextState.history.length, 0);
      assert.deepEqual(nextState.variables, initialState.variables);
      assert.equal(nextState.ending, null);
      assert.match(nextState.systemFeedback, /Enter|No exact keyword match/);
    }
  });

  it('returns Regenerative Earth for a strong positive path', () => {
    const finalState = playPath(['repair', 'share', 'restore', 'organise', 'justice']);

    assert.equal(finalState.screen, 'ending');
    assert.equal(finalState.ending.id, 'regenerative-earth');
    assert.equal(finalState.history.length, 5);
  });

  it('returns Collapse Story for repeated harmful choices', () => {
    const finalState = playPath(['consume', 'discard', 'ignore', 'delay', 'greenwash']);

    assert.equal(finalState.screen, 'ending');
    assert.equal(finalState.ending.id, 'collapse-story');
  });

  it('keeps the technology-heavy low justice path in Technological Green Future', () => {
    const finalState = playPath(['technology', 'efficiency', 'build', 'protect', 'adapt']);

    assert.equal(finalState.screen, 'ending');
    assert.equal(finalState.ending.id, 'technological-green-future');
  });

  it('does not let the last action unconditionally override the Earth state', () => {
    const variables = {
      earthHealth: 45,
      temperature: 45,
      biodiversity: 45,
      communityCare: 35,
      justice: 35,
      waste: 45,
      hope: 40,
      futureMomentum: 0
    };
    const state = determineEarthState(variables, { preferredEarthState: 'repairing' });

    assert.equal(state, 'fragile');
  });
});
