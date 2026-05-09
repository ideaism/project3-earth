import { useEffect, useMemo, useState } from 'react';
import {
  createInitialAppState,
  getBranchPrompt,
  getChapter,
  getInterfaceCopy,
  getPrologue,
  getSuggestedActions,
  getVisualState,
  submitKeyword
} from './engine';
import { clearArchiveEntries, loadArchive, saveArchiveEntry } from './storage';

export default function App() {
  const [appState, setAppState] = useState(() => createInitialAppState());
  const [archiveState, setArchiveState] = useState(() => ({
    available: true,
    entries: [],
    error: null
  }));
  const [savedEndingId, setSavedEndingId] = useState(null);

  useEffect(() => {
    setArchiveState(loadArchive());
  }, []);

  function beginStory() {
    setAppState((current) => ({
      ...current,
      screen: 'story'
    }));
  }

  function handleSubmitKeyword(keyword) {
    setAppState((current) => submitKeyword(current, keyword));
  }

  function restart() {
    setSavedEndingId(null);
    setAppState({
      ...createInitialAppState(),
      screen: 'story'
    });
  }

  function returnHome() {
    setSavedEndingId(null);
    setAppState(createInitialAppState());
  }

  function saveCurrentEnding() {
    if (!appState.ending) return;

    const archiveEntry = {
      id: `${Date.now()}-${appState.ending.id}`,
      savedAt: new Date().toISOString(),
      endingId: appState.ending.id,
      endingTitle: appState.ending.title,
      visualState: appState.earthState,
      keywords: appState.history.map((item) => item.label),
      variables: appState.variables
    };

    const result = saveArchiveEntry(archiveEntry);
    setArchiveState(result);
    if (result.available) {
      setSavedEndingId(archiveEntry.id);
    }
  }

  function clearArchive() {
    setArchiveState(clearArchiveEntries());
  }

  if (appState.screen === 'archive') {
    return (
      <LocalArchive
        archiveState={archiveState}
        onClear={clearArchive}
        onBack={() => setAppState((current) => ({ ...current, screen: current.ending ? 'ending' : 'story' }))}
        onHome={returnHome}
      />
    );
  }

  if (appState.screen === 'ending') {
    return (
      <EndingCard
        appState={appState}
        savedEndingId={savedEndingId}
        archiveState={archiveState}
        onRestart={restart}
        onSave={saveCurrentEnding}
        onViewArchive={() => setAppState((current) => ({ ...current, screen: 'archive' }))}
        onHome={returnHome}
      />
    );
  }

  if (appState.screen === 'story') {
    return (
      <StoryScreen
        appState={appState}
        onSubmitKeyword={handleSubmitKeyword}
        onViewArchive={() => setAppState((current) => ({ ...current, screen: 'archive' }))}
        onHome={returnHome}
      />
    );
  }

  return <LandingScreen onBegin={beginStory} onViewArchive={() => setAppState((current) => ({ ...current, screen: 'archive' }))} />;
}

function LandingScreen({ onBegin, onViewArchive }) {
  const prologue = getPrologue();

  return (
    <main className="app-frame landing-frame">
      <section className="landing-copy" aria-labelledby="project-title">
        <p className="screen-kicker">{prologue.label}</p>
        <h1 id="project-title">The Earth, Written by Us</h1>
        <p className="subtitle">An Earth Story You Can Change</p>
        <p className="prologue">{prologue.text}</p>
        <div className="button-row">
          <button className="primary-button" type="button" onClick={onBegin}>
            {prologue.beginButtonLabel}
          </button>
          <button className="secondary-button" type="button" onClick={onViewArchive}>
            Archive
          </button>
        </div>
      </section>
      <section className="landing-earth" aria-label="Fragile Earth visual">
        <EarthVisual earthState="fragile" variables={null} feedbackMessage="The fragile Earth waits inside the page." />
      </section>
    </main>
  );
}

function StoryScreen({ appState, onSubmitKeyword, onViewArchive, onHome }) {
  const visualState = getVisualState(appState.earthState);

  return (
    <main className="app-frame story-frame">
      <section className="earth-panel" aria-labelledby="earth-state-heading">
        <div className="panel-heading">
          <p className="screen-kicker">Earth status</p>
          <h2 id="earth-state-heading">{visualState.name}</h2>
        </div>
        <EarthVisual
          earthState={appState.earthState}
          variables={appState.variables}
          feedbackMessage={appState.feedbackMessage}
        />
      </section>
      <StoryPanel
        appState={appState}
        onSubmitKeyword={onSubmitKeyword}
        onViewArchive={onViewArchive}
        onHome={onHome}
      />
    </main>
  );
}

function EarthVisual({ earthState, variables, feedbackMessage }) {
  const visualState = getVisualState(earthState);
  const fallbackEarthAsset = '/assets/earth/earth-fragile.png';
  const meters = variables
    ? [
        ['Health', variables.earthHealth],
        ['Heat', variables.temperature],
        ['Life', variables.biodiversity],
        ['Care', variables.communityCare],
        ['Justice', variables.justice],
        ['Waste', variables.waste]
      ]
    : [];

  return (
    <div
      className={`earth-visual ${visualState.backgroundClass} earth-animation-${visualState.animationType}`}
      key={`${earthState}-${feedbackMessage}`}
    >
      <div className="earth-orbit" aria-hidden="true">
        <img
          className="earth-image"
          src={visualState.assetPath}
          alt=""
          onError={(event) => {
            if (!event.currentTarget.src.endsWith(fallbackEarthAsset)) {
              event.currentTarget.src = fallbackEarthAsset;
            }
          }}
        />
        {(visualState.effectAssets ?? []).map((assetPath, index) => (
          <img
            className={`earth-effect earth-effect-${index + 1}`}
            src={assetPath}
            alt=""
            key={assetPath}
            onError={(event) => {
              event.currentTarget.remove();
            }}
          />
        ))}
      </div>
      <div className="earth-caption">
        <span>Current state</span>
        <strong>{visualState.name}</strong>
        <p>{feedbackMessage || 'The Earth state is waiting for action.'}</p>
      </div>
      {meters.length > 0 && (
        <dl className="variable-grid" aria-label="Current Earth variables">
          {meters.map(([label, value]) => (
            <div className="variable-meter" key={label}>
              <dt>{label}</dt>
              <dd>
                <span style={{ width: `${value}%` }} />
                <b>{value}</b>
              </dd>
            </div>
          ))}
        </dl>
      )}
    </div>
  );
}

function StoryPanel({ appState, onSubmitKeyword, onViewArchive, onHome }) {
  const chapter = getChapter(appState.round);
  const branchPrompt = getBranchPrompt(appState.round);
  const recentHistory = appState.history.slice(-5);
  const copy = getInterfaceCopy();

  return (
    <section className="story-panel" aria-labelledby="chapter-title">
      <header className="story-header">
        <div>
          <p className="screen-kicker">Chapter {chapter.chapter} / {appState.maxRounds}</p>
          <h2 id="chapter-title">{chapter.title}</h2>
        </div>
        <div className="header-actions">
          <button className="text-button" type="button" onClick={onViewArchive}>
            Archive
          </button>
          <button className="text-button" type="button" onClick={onHome}>
            Home
          </button>
        </div>
      </header>

      <ProgressIndicator round={appState.round} maxRounds={appState.maxRounds} />

      <article className="story-copy" aria-live="polite">
        <p>{appState.currentStoryText}</p>
        <blockquote>{appState.currentEarthVoice}</blockquote>
        {appState.systemFeedback && <p className="system-feedback">{appState.systemFeedback}</p>}
      </article>

      <p className="choice-prompt">{appState.currentPrompt}</p>
      <ActionChips actions={getSuggestedActions()} onSubmitKeyword={onSubmitKeyword} />

      {branchPrompt && (
        <section className="branch-panel" aria-labelledby="branch-title">
          <h3 id="branch-title">Branching choice</h3>
          <p>{branchPrompt.question}</p>
          <div className="chip-row">
            {branchPrompt.choices.map((choice) => (
              <button
                className={`choice-chip ${isRiskChoice(choice) ? 'choice-danger' : ''}`}
                type="button"
                key={choice}
                onClick={() => onSubmitKeyword(choice)}
              >
                {choice}
              </button>
            ))}
          </div>
        </section>
      )}

      <section className="recent-path" aria-labelledby="recent-path-title">
        <h3 id="recent-path-title">{copy.chosenWords}</h3>
        {recentHistory.length === 0 ? (
          <p className="muted">No keywords selected yet.</p>
        ) : (
          <ol>
            {recentHistory.map((item) => (
              <li key={`${item.round}-${item.keyword}-${item.input}`}>
                <span>{item.round}</span>
                <strong>{item.label}</strong>
                <em>{item.earthStateAfter}</em>
              </li>
            ))}
          </ol>
        )}
      </section>
    </section>
  );
}

function isRiskChoice(choice) {
  return ['extract', 'ignore', 'consume', 'discard', 'delay', 'greenwash', 'exploit', 'waste'].includes(choice);
}

function ProgressIndicator({ round, maxRounds }) {
  const progress = Math.round((round / maxRounds) * 100);
  const visibleChapter = Math.min(round + 1, maxRounds);

  return (
    <div className="progress-block" aria-label={`Round ${round} of ${maxRounds}`}>
      <div className="progress-meta">
        <span>Chapter {visibleChapter} / {maxRounds}</span>
        <span>{maxRounds - round} actions remaining</span>
      </div>
      <div className="progress-track">
        <span style={{ width: `${progress}%` }} />
      </div>
      <div className="chapter-dots" aria-hidden="true">
        {Array.from({ length: maxRounds }, (_, index) => {
          const dotClass = index < round ? 'is-complete' : index === visibleChapter - 1 ? 'is-current' : '';
          return <span className={dotClass} key={index} />;
        })}
      </div>
    </div>
  );
}

function ActionChips({ actions, onSubmitKeyword }) {
  return (
    <section className="chip-section" aria-label="Suggested action chips">
      <div className="chip-row">
        {actions.map((action) => (
          <button className="action-chip" type="button" key={action} onClick={() => onSubmitKeyword(action)}>
            {action}
          </button>
        ))}
      </div>
    </section>
  );
}

function EndingCard({ appState, savedEndingId, archiveState, onRestart, onSave, onViewArchive, onHome }) {
  const visualState = getVisualState(appState.earthState);
  const copy = getInterfaceCopy();
  const keywordPath = appState.history.map((item) => item.label);
  const archiveMessage = useMemo(() => {
    if (!archiveState.available) return 'Archive unavailable in this browser session.';
    if (savedEndingId) return 'Saved to local archive.';
    return `${copy.saveThisEnding} to the local archive.`;
  }, [archiveState.available, savedEndingId]);

  return (
    <main className="app-frame ending-frame">
      <section className="ending-card" aria-labelledby="ending-title">
        <p className="screen-kicker">{copy.endingCardLabel}</p>
        <h1 id="ending-title">{appState.ending.title}</h1>
        <p className="ending-state">{copy.finalState}: {visualState.name}</p>
        <p className="ending-summary">{appState.ending.summary}</p>
        {appState.ending.earthVoice && <blockquote>{appState.ending.earthVoice}</blockquote>}
        <p className="reflection-prompt">{appState.ending.reflectionPrompt}</p>

        <section className="path-strip" aria-label={copy.yourStoryPath}>
          {keywordPath.map((keyword, index) => (
            <span key={`${keyword}-${index}`}>{keyword}</span>
          ))}
        </section>

        <div className="button-row">
          <button className="primary-button" type="button" onClick={onRestart}>
            {copy.restart}
          </button>
          <button className="secondary-button" type="button" onClick={onSave} disabled={Boolean(savedEndingId)}>
            {copy.saveThisEnding}
          </button>
          <button className="secondary-button" type="button" onClick={onViewArchive}>
            Archive
          </button>
          <button className="text-button" type="button" onClick={onHome}>
            Home
          </button>
        </div>
        <p className="archive-note">{archiveMessage}</p>
      </section>
      <section className="ending-visual" aria-label="Final Earth visual">
        <EarthVisual earthState={appState.earthState} variables={appState.variables} feedbackMessage={copy.pathSummary} />
      </section>
    </main>
  );
}

function LocalArchive({ archiveState, onClear, onBack, onHome }) {
  const copy = getInterfaceCopy();

  return (
    <main className="app-frame archive-frame">
      <section className="archive-panel" aria-labelledby="archive-title">
        <p className="screen-kicker">{copy.archiveOfStories}</p>
        <h1 id="archive-title">Completed Story Paths</h1>
        {!archiveState.available && (
          <p className="error-text">Archive unavailable in this browser session. The prototype still works without storage.</p>
        )}
        {archiveState.entries.length === 0 ? (
          <p className="muted">{copy.emptyArchive}</p>
        ) : (
          <ol className="archive-list">
            {archiveState.entries.map((entry) => (
              <li key={entry.id}>
                <div>
                  <strong>{entry.endingTitle}</strong>
                  <span>{new Date(entry.savedAt).toLocaleString()}</span>
                </div>
                <p>{entry.keywords.join(' -> ')}</p>
              </li>
            ))}
          </ol>
        )}
        <div className="button-row">
          <button className="primary-button" type="button" onClick={onBack}>
            Back
          </button>
          <button className="secondary-button" type="button" onClick={onClear} disabled={archiveState.entries.length === 0}>
            Clear archive
          </button>
          <button className="text-button" type="button" onClick={onHome}>
            Home
          </button>
        </div>
      </section>
    </main>
  );
}
