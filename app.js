const SPREADSHEET_ID = "19GdI5qQWc-VyLQEgRiJfaxxTLwrtxU6ofx4tZYPax0M";
const GVIZ = (sheet) =>
  `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(
    sheet
  )}`;
const GVIZ_JSON = (sheet) =>
  `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(
    sheet
  )}`;

/** Speaking: C列に値がある行だけ取得（2行目・3行目のトピックを確実に含める） */
function speakingGvizUrlWithQuery(tq) {
  return `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=Speaking&tq=${encodeURIComponent(
    tq
  )}`;
}

const SUBMIT_URL = typeof window.SUBMIT_URL !== "undefined" ? window.SUBMIT_URL : "";
const SUBMIT_SECRET = typeof window.SUBMIT_SECRET !== "undefined" ? window.SUBMIT_SECRET : "";

function fetchCsv(url) {
  return fetch(url).then((r) => {
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    return r.text();
  });
}

function parseCsv(text) {
  return Papa.parse(text, {
    header: true,
    skipEmptyLines: "greedy",
  });
}

function shuffle(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * 単語の正解判定ロジック（概要）
 * 1) NFKC 正規化・前後空白除去・連続空白を1つにまとめる（英数字は全半統一）
 * 2) 完全一致（正解文全体、または / ・ 、 などで区切った別表記のいずれか）
 * 3) 別表記セグメントや全文について、表記ゆれを許容した「含む」一致（短すぎる1文字のみは除外）
 * 4) レーベンシュタイン距離に基づく類似度が一定以上なら正解（短い語句向け）
 */
function normalizeJa(s) {
  return String(s)
    .normalize("NFKC")
    .trim()
    .replace(/\s+/g, "");
}

function stripNoise(s) {
  return normalizeJa(s).replace(/[・。．、，,.．\s（）()「」『』]/g, "");
}

function splitAnswerVariants(expectedJa) {
  return String(expectedJa)
    .split(/[/／、,，|｜\n]/)
    .map((p) => normalizeJa(p.trim()))
    .filter((p) => p.length > 0);
}

function levenshtein(a, b) {
  const m = a.length;
  const n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
    }
  }
  return dp[m][n];
}

function similarityRatio(a, b) {
  if (!a.length || !b.length) return 0;
  const d = levenshtein(a, b);
  return 1 - d / Math.max(a.length, b.length);
}

function vocabIsCorrect(userText, expectedJa) {
  const u = normalizeJa(userText);
  if (!u) return false;
  const full = normalizeJa(expectedJa);
  const uPlain = stripNoise(userText);
  const fullPlain = stripNoise(expectedJa);

  if (full && u === full) return true;
  if (fullPlain && uPlain === fullPlain) return true;

  const parts = splitAnswerVariants(expectedJa);
  if (parts.some((p) => p && u === p)) return true;
  if (parts.some((p) => p && stripNoise(p) && uPlain === stripNoise(p))) return true;

  const minSub = 2;
  if (full.length >= minSub && (u.includes(full) || full.includes(u))) return true;
  if (fullPlain.length >= minSub && (uPlain.includes(fullPlain) || fullPlain.includes(uPlain))) return true;

  for (const p of parts) {
    if (!p || p.length < minSub) continue;
    if (u.includes(p) || p.includes(u)) return true;
    const pp = stripNoise(p);
    if (pp.length >= minSub && (uPlain.includes(pp) || pp.includes(uPlain))) return true;
    if (similarityRatio(u, p) >= 0.78) return true;
    if (pp.length >= 3 && similarityRatio(uPlain, pp) >= 0.78) return true;
  }
  if (full.length >= 3 && similarityRatio(u, full) >= 0.78) return true;
  if (fullPlain.length >= 3 && similarityRatio(uPlain, fullPlain) >= 0.78) return true;

  return false;
}

let participantName = "";
let currentStep = 0;
let vocabRows = [];
let grammarRows = [];
let grammarAnswers = [];
let grammarIndex = 0;
let speakingTopicCount = 0;

const SESSION_KEY = "kepty-measurement-session-v3";
let restoringSession = false;
/** loadData 完了・restore 前に persist すると screen:intro で既存セッションを潰すため抑止する */
let allowSessionPersist = false;

function readStoredSession() {
  try {
    return sessionStorage.getItem(SESSION_KEY) || localStorage.getItem(SESSION_KEY);
  } catch {
    return null;
  }
}

function getCurrentScreen() {
  const th = document.getElementById("screen-thanks");
  const flow = document.getElementById("test-flow");
  // hidden 属性 / element.hidden が真実。class の .hidden だけだと属性と不整合で誤判定しうる
  if (th && !th.hidden) return "thanks";
  if (flow && !flow.hidden) return "test";
  return "intro";
}

function collectVocabSnapshot() {
  const arr = [];
  document.querySelectorAll("#vocab-root .vocab__input").forEach((el) => {
    arr.push(el.value || "");
  });
  return arr;
}

function persistSessionState() {
  if (restoringSession) return;
  if (!allowSessionPersist) return;
  try {
    saveCurrentGrammarSelection();
    const nameEl = document.getElementById("participant-name");
    const payload = {
      v: 1,
      screen: getCurrentScreen(),
      participantName: participantName || "",
      nameDraft: nameEl ? nameEl.value || "" : "",
      currentStep,
      grammarIndex,
      grammarAnswers: grammarAnswers.slice(),
      vocabAnswers: collectVocabSnapshot(),
    };
    const str = JSON.stringify(payload);
    sessionStorage.setItem(SESSION_KEY, str);
    try {
      localStorage.setItem(SESSION_KEY, str);
    } catch (e2) {
      console.warn(e2);
    }
  } catch (e) {
    console.warn(e);
  }
}

function restoreSessionState() {
  restoringSession = true;
  let shouldPersistAfter = false;
  try {
    const raw = readStoredSession();
    if (!raw) return;
    const s = JSON.parse(raw);
    if (s.v !== 1) return;

    if (s.screen === "thanks") {
      participantName = s.participantName || "";
      showThankYou();
      if (participantName) {
        const pd = document.getElementById("participant-display");
        if (pd) pd.textContent = `実施者：${participantName}`;
      }
      shouldPersistAfter = true;
      return;
    }

    if (s.screen === "test") {
      participantName = s.participantName || "";
      const nameInput = document.getElementById("participant-name");
      if (nameInput && typeof s.nameDraft === "string") nameInput.value = s.nameDraft;
      const pd = document.getElementById("participant-display");
      if (pd) pd.textContent = participantName ? `実施者：${participantName}` : "";

      if (grammarRows.length) {
        grammarAnswers = grammarRows.map((_, i) => {
          const a = s.grammarAnswers && s.grammarAnswers[i];
          if (a == null || String(a).trim() === "") return null;
          return String(a);
        });
        grammarIndex =
          typeof s.grammarIndex === "number"
            ? Math.max(0, Math.min(s.grammarIndex, grammarRows.length - 1))
            : 0;
        buildGrammarCard();
      }

      if (Array.isArray(s.vocabAnswers)) {
        const inputs = document.querySelectorAll("#vocab-root .vocab__input");
        const n = Math.min(s.vocabAnswers.length, inputs.length);
        for (let i = 0; i < n; i++) {
          inputs[i].value = s.vocabAnswers[i] || "";
        }
      }

      const introEl = document.getElementById("intro");
      introEl.classList.add("hidden");
      introEl.setAttribute("hidden", "");
      const flow = document.getElementById("test-flow");
      flow.classList.remove("hidden");
      flow.removeAttribute("hidden");
      flow.hidden = false;
      const step = typeof s.currentStep === "number" ? s.currentStep : 0;
      setStep(Math.max(0, Math.min(2, step)));
      shouldPersistAfter = true;
      return;
    }

    if (s.screen === "intro" && typeof s.nameDraft === "string") {
      const ni = document.getElementById("participant-name");
      if (ni) ni.value = s.nameDraft;
      shouldPersistAfter = true;
    }
  } catch (e) {
    console.warn(e);
  } finally {
    restoringSession = false;
    if (shouldPersistAfter) persistSessionState();
  }
}

function setStep(n) {
  currentStep = Math.max(0, Math.min(2, n));
  document.querySelectorAll("#test-flow .panel").forEach((el, i) => {
    const active = i === currentStep;
    el.classList.toggle("is-active", active);
    el.hidden = !active;
  });
  document.querySelectorAll("#test-flow .stepper__item").forEach((el, i) => {
    el.classList.toggle("is-current", i === currentStep);
    el.classList.toggle("is-done", i < currentStep);
  });
  document.getElementById("btn-prev").disabled = currentStep === 0;
  const isLast = currentStep === 2;
  document.getElementById("btn-next").textContent = isLast
    ? "①②の結果を送信し、スピーキングへ進む"
    : "次へ進む";
  clearFooterError();
  persistSessionState();
}

function clearFooterError() {
  const el = document.getElementById("footer-error");
  if (!el) return;
  el.textContent = "";
  el.hidden = true;
}

function getIncompleteVocabIndices() {
  const areas = document.querySelectorAll("#vocab-root .vocab__input");
  const miss = [];
  areas.forEach((el, i) => {
    if (!(el.value || "").trim()) miss.push(i + 1);
  });
  return miss;
}

function getIncompleteGrammarIndices() {
  saveCurrentGrammarSelection();
  if (!grammarRows.length) return [];
  const miss = [];
  grammarAnswers.forEach((a, i) => {
    if (a == null || String(a).trim() === "") miss.push(i + 1);
  });
  return miss;
}

function formatQuestionNums(nums) {
  return nums.map((n) => `#${n}`).join("、");
}

function showIncompleteModal(message) {
  return new Promise((resolve) => {
    const modal = document.getElementById("incomplete-modal");
    const msgEl = document.getElementById("incomplete-modal-message");
    const okBtn = document.getElementById("incomplete-modal-ok");
    const cancelBtn = document.getElementById("incomplete-modal-cancel");
    const backdrop = modal.querySelector("[data-modal-dismiss]");
    msgEl.textContent = message;
    modal.hidden = false;

    function cleanup() {
      modal.hidden = true;
      okBtn.removeEventListener("click", onOk);
      cancelBtn.removeEventListener("click", onCancel);
      backdrop.removeEventListener("click", onCancel);
      document.removeEventListener("keydown", onKey);
    }

    function onOk() {
      cleanup();
      resolve(true);
    }
    function onCancel() {
      cleanup();
      resolve(false);
    }
    function onKey(e) {
      if (e.key === "Escape") onCancel();
    }
    okBtn.addEventListener("click", onOk);
    cancelBtn.addEventListener("click", onCancel);
    backdrop.addEventListener("click", onCancel);
    document.addEventListener("keydown", onKey);
    okBtn.focus();
  });
}

async function tryGoForwardTo(targetStep) {
  if (targetStep <= currentStep) return;
  saveCurrentGrammarSelection();
  const nextBtn = document.getElementById("btn-next");
  nextBtn.disabled = true;
  try {
    if (currentStep === 0) {
      if (targetStep >= 1) {
        const missV = getIncompleteVocabIndices();
        if (missV.length) {
          const msg = `設問 ${formatQuestionNums(missV)} が未入力ですが、よろしいでしょうか。`;
          if (!(await showIncompleteModal(msg))) return;
        }
      }
      if (targetStep === 1) {
        setStep(1);
        return;
      }
      if (targetStep === 2) {
        const missG = getIncompleteGrammarIndices();
        if (missG.length) {
          const msg = `設問 ${formatQuestionNums(missG)} が未選択ですが、よろしいでしょうか。`;
          if (!(await showIncompleteModal(msg))) return;
        }
        setStep(2);
        return;
      }
    }
    if (currentStep === 1 && targetStep === 2) {
      const missG = getIncompleteGrammarIndices();
      if (missG.length) {
        const msg = `設問 ${formatQuestionNums(missG)} が未選択ですが、よろしいでしょうか。`;
        if (!(await showIncompleteModal(msg))) return;
      }
      setStep(2);
    }
  } finally {
    nextBtn.disabled = false;
  }
}

function saveCurrentGrammarSelection() {
  const root = document.getElementById("grammar-root");
  const sel = root.querySelector('input[name="gopt"]:checked');
  if (sel) grammarAnswers[grammarIndex] = sel.value;
}

function renderVocabulary(rows) {
  const root = document.getElementById("vocab-root");
  root.innerHTML = "";
  let n = 0;
  rows.forEach((row) => {
    const en = vocabEnglishForPrompt(row);
    if (!en) return;
    const num = n + 1;
    const id = `vocab-${n++}`;
    const item = document.createElement("article");
    item.className = "vocab__item";
    item.innerHTML = `
      <p class="vocab__prompt">
        <span class="vocab__index">#${num}</span><span class="vocab__word">${escapeHtml(en)}</span>
      </p>
      <label class="vocab__label" for="${id}">日本語の意味を入力</label>
      <input type="text" id="${id}" class="vocab__input" autocomplete="off" spellcheck="false" placeholder="例：サッカー" />
    `;
    root.appendChild(item);
  });
}

function buildGrammarCard() {
  const root = document.getElementById("grammar-root");
  if (!grammarRows.length) {
    root.innerHTML = '<p class="status status--error">文法の問題が見つかりませんでした。</p>';
    return;
  }
  const row = grammarRows[grammarIndex];
  const q = (row.question || "").trim();
  const correct = (row.answer || "").trim();
  const wrong = [
    (row.option2 || "").trim(),
    (row.option3 || "").trim(),
    (row.option4 || "").trim(),
  ].filter(Boolean);
  const options = shuffle([correct, ...wrong].filter((x) => x.length > 0));
  const isLastQ = grammarIndex >= grammarRows.length - 1;

  root.innerHTML = `
    <div class="grammar__toolbar">
      <span class="grammar__progress">設問 <span class="grammar__index">#${grammarIndex + 1}</span><span class="grammar__total"> / ${grammarRows.length}</span></span>
    </div>
    <div class="grammar__card">
      <p class="grammar__inline-error" id="grammar-inline-error" role="alert" hidden>選択肢を1つ選んでから次の問題へ進んでください。</p>
      <p class="grammar__question">${escapeHtml(q)}</p>
      <div class="grammar__options" role="radiogroup" aria-label="選択肢">
        ${options
          .map(
            (opt) => `
          <label class="grammar__option">
            <input type="radio" name="gopt" value="${escapeHtml(opt)}" />
            <span>${escapeHtml(opt)}</span>
          </label>`
          )
          .join("")}
      </div>
      <div class="grammar__nav ${isLastQ ? "grammar__nav--end-only" : ""}">
        <button type="button" class="btn btn--ghost" id="grammar-prev" ${grammarIndex === 0 ? "disabled" : ""}>前の問題</button>
        ${
          isLastQ
            ? ""
            : `<button type="button" class="btn btn--primary" id="grammar-next">次の問題</button>`
        }
      </div>
    </div>
  `;

  root.querySelectorAll('input[name="gopt"]').forEach((inp) => {
    inp.addEventListener("change", () => {
      grammarAnswers[grammarIndex] = inp.value;
      const err = document.getElementById("grammar-inline-error");
      if (err) err.hidden = true;
      persistSessionState();
    });
  });

  const saved = grammarAnswers[grammarIndex];
  if (saved) {
    const match = Array.from(root.querySelectorAll('input[name="gopt"]')).find((i) => i.value === saved);
    if (match) match.checked = true;
  }

  document.getElementById("grammar-prev").addEventListener("click", () => {
    saveCurrentGrammarSelection();
    grammarIndex = Math.max(0, grammarIndex - 1);
    buildGrammarCard();
  });
  const nextBtn = document.getElementById("grammar-next");
  if (nextBtn) {
    nextBtn.addEventListener("click", () => {
      const sel = root.querySelector('input[name="gopt"]:checked');
      if (!sel) {
        const err = document.getElementById("grammar-inline-error");
        if (err) err.hidden = false;
        return;
      }
      grammarAnswers[grammarIndex] = sel.value;
      if (grammarIndex < grammarRows.length - 1) {
        grammarIndex += 1;
        buildGrammarCard();
      }
    });
  }
  persistSessionState();
}

function trainingTopicFromRow(row) {
  if (!row || typeof row !== "object") return "";
  /** gviz JSON（ラベル空）や CSV の C 列 */
  const fromC = row.C ?? row.c ?? row.col2;
  if (fromC != null && String(fromC).trim()) {
    const s = String(fromC).trim();
    if (s.toLowerCase() !== "training topic") return s;
  }
  const keys = Object.keys(row);
  for (const k of keys) {
    if (k.replace(/\s+/g, " ").trim().toLowerCase() === "training topic") {
      const v = row[k];
      if (v != null && String(v).trim()) return String(v);
    }
  }
  const direct = row["training topic"] ?? row["training topic "] ?? row["Training topic"];
  if (direct != null && String(direct).trim()) return String(direct);
  return "";
}

function isSpeakingHeaderRow(row) {
  const tt = trainingTopicFromRow(row).trim().toLowerCase();
  if (tt === "training topic") return true;
  const a = String(row.A ?? row.col0 ?? "").trim().toLowerCase();
  if (a === "no." || a === "no") return true;
  return false;
}

/** 英語プロンプトのみ表示（level 列や誤結合の Stage 行は出さない） */
function vocabEnglishForPrompt(row) {
  let raw = (row.english || "").trim();
  raw = raw.replace(/^stage\s*\d+\s*[:\-–—]?\s*/i, "").trim();
  return raw;
}

/** Google Visualization API の JSON（setResponse）をパースして行オブジェクトの配列にする */
function parseGvizJson(text) {
  const marker = "setResponse(";
  const start = text.indexOf(marker);
  if (start < 0) return [];
  const jsonStart = start + marker.length;
  const end = text.lastIndexOf(");");
  if (end <= jsonStart) return [];
  let payload;
  try {
    payload = JSON.parse(text.slice(jsonStart, end));
  } catch {
    return [];
  }
  if (payload.status !== "ok" || !payload.table) return [];
  const cols = (payload.table.cols || []).map((c, i) => {
    const lab = (c.label || "").trim();
    const id = (c.id || "").trim();
    return lab || id || `col${i}`;
  });
  const rows = payload.table.rows || [];
  return rows.map((r) => {
    const o = {};
    (r.c || []).forEach((cell, i) => {
      const label = cols[i] || `col${i}`;
      let v = "";
      if (cell) {
        if (cell.v != null && cell.v !== "") v = String(cell.v);
        else if (cell.f != null && cell.f !== "") v = String(cell.f);
      }
      o[label] = v;
    });
    return o;
  });
}

function themeFromRow(row) {
  if (!row || typeof row !== "object") return "";
  const hit = Object.keys(row).find((k) => k.replace(/\s+/g, " ").trim().toLowerCase() === "theme");
  if (hit) {
    const s = String(row[hit] || "").trim();
    if (s.toLowerCase() !== "theme") return s;
  }
  const b = row.B ?? row.b ?? row.col1;
  if (b != null && String(b).trim()) {
    const s = String(b).trim();
    if (s.toLowerCase() !== "theme") return s;
  }
  return String(row.theme || "").trim();
}

function renderSpeaking(rows) {
  const root = document.getElementById("speaking-root");
  root.innerHTML = "";
  const dataRows = rows
    .filter((r) => !isSpeakingHeaderRow(r))
    .filter((r) => trainingTopicFromRow(r).trim().length > 0)
    .slice(0, 2);
  speakingTopicCount = dataRows.length;
  dataRows.forEach((row, i) => {
    const text = trainingTopicFromRow(row).replace(/\r\n/g, "\n").trim();
    const theme = themeFromRow(row);
    const block = document.createElement("div");
    block.className = "speaking__topic";
    block.innerHTML = `
      <p class="speaking__label">トピック ${i + 1}${theme ? ` · ${escapeHtml(theme)}` : ""}</p>
      <p class="speaking__text">${escapeHtml(text)}</p>
    `;
    root.appendChild(block);
  });
  if (!root.children.length) {
    root.innerHTML =
      '<p class="status status--error">Speaking シートの C 列にトピックが見つかりませんでした。</p>';
    speakingTopicCount = 0;
  }
}

function collectVocabAnswers() {
  const areas = document.querySelectorAll("#vocab-root .vocab__input");
  const items = [];
  let i = 0;
  vocabRows.forEach((row) => {
    const en = vocabEnglishForPrompt(row);
    if (!en) return;
    const ja = (row.japanese || "").trim();
    const userText = areas[i] ? areas[i].value : "";
    items.push({
      english: en,
      userAnswer: userText,
      expectedJapanese: ja,
      isCorrect: vocabIsCorrect(userText, ja),
    });
    i += 1;
  });
  return items;
}

function collectGrammarAnswers() {
  saveCurrentGrammarSelection();
  return grammarRows.map((row, idx) => {
    const correct = (row.answer || "").trim();
    const chosen = grammarAnswers[idx];
    return {
      question: (row.question || "").trim(),
      chosen: chosen == null ? "" : chosen,
      correctAnswer: correct,
      isCorrect: chosen != null && chosen === correct,
    };
  });
}

function ratePercent(correct, total) {
  if (!total) return 0;
  return Math.round((correct / total) * 10000) / 100;
}

function buildPayload() {
  const vocabItems = collectVocabAnswers();
  const grammarItems = collectGrammarAnswers();

  const vCorrect = vocabItems.filter((x) => x.isCorrect).length;
  const vTotal = vocabItems.length;
  const gCorrect = grammarItems.filter((x) => x.isCorrect).length;
  const gTotal = grammarItems.length;

  const out = {
    name: participantName,
    scores: {
      vocabulary: {
        correct: vCorrect,
        total: vTotal,
        ratePercent: ratePercent(vCorrect, vTotal),
      },
      grammar: {
        correct: gCorrect,
        total: gTotal,
        ratePercent: ratePercent(gCorrect, gTotal),
      },
    },
    raw: {
      vocabulary: vocabItems,
      grammar: grammarItems,
    },
  };
  if (SUBMIT_SECRET) out.secret = SUBMIT_SECRET;
  return out;
}

function showThankYou() {
  document.getElementById("intro").classList.add("hidden");
  document.getElementById("intro").setAttribute("hidden", "");
  const tf = document.getElementById("test-flow");
  tf.classList.add("hidden");
  tf.setAttribute("hidden", "");
  tf.hidden = true;
  const th = document.getElementById("screen-thanks");
  th.classList.remove("hidden");
  th.removeAttribute("hidden");
  th.hidden = false;
  persistSessionState();
}

function downloadPayload(obj) {
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const safe = participantName.replace(/[^\w\u3000-\u9fff\u3040-\u30ff-]+/g, "_").slice(0, 40);
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `result_${safe}_${stamp}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
}

async function submitResults() {
  const payload = buildPayload();
  const btn = document.getElementById("btn-next");
  btn.disabled = true;
  showThankYou();

  if (!SUBMIT_URL || !SUBMIT_URL.trim()) {
    downloadPayload(payload);
    return;
  }

  try {
    const res = await fetch(SUBMIT_URL.trim(), {
      method: "POST",
      mode: "cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });
    const text = await res.text();
    let data = {};
    try {
      data = JSON.parse(text);
    } catch {
      data = {};
    }
    if (!res.ok || data.ok === false) {
      throw new Error(data.error || text || `HTTP ${res.status}`);
    }
  } catch (e) {
    console.error(e);
    downloadPayload(payload);
  }
}

function speakingRowsWithTopics(rows) {
  return rows
    .filter((r) => !isSpeakingHeaderRow(r))
    .filter((r) => trainingTopicFromRow(r).trim().length > 0);
}

async function fetchSpeakingRows() {
  /** セル内改行があると CSV(Papa) が壊れるため、JSON を最優先する */
  try {
    const plain = await fetchCsv(GVIZ_JSON("Speaking"));
    const rows = parseGvizJson(plain);
    const withTopic = speakingRowsWithTopics(rows);
    if (withTopic.length >= 2) return withTopic.slice(0, 2);
    if (withTopic.length === 1) return withTopic;
  } catch (e) {
    console.warn(e);
  }

  const queries = [
    "SELECT * WHERE Col3 IS NOT NULL",
    "SELECT A, B, C WHERE Col3 IS NOT NULL",
    "SELECT * WHERE C IS NOT NULL",
  ];
  for (const tq of queries) {
    try {
      const text = await fetchCsv(speakingGvizUrlWithQuery(tq));
      const rows = parseGvizJson(text);
      const withTopic = speakingRowsWithTopics(rows);
      if (withTopic.length >= 2) return withTopic.slice(0, 2);
      if (withTopic.length === 1) return withTopic;
    } catch (e) {
      console.warn(e);
    }
  }

  try {
    const sCsv = await fetchCsv(GVIZ("Speaking"));
    const sParsed = Papa.parse(sCsv, {
      header: true,
      skipEmptyLines: "greedy",
      relaxColumnCount: true,
      relaxQuotes: true,
    });
    const fromCsv = speakingRowsWithTopics(sParsed.data || []);
    if (fromCsv.length >= 2) return fromCsv.slice(0, 2);
    if (fromCsv.length === 1) return fromCsv;
  } catch (e) {
    console.warn(e);
  }
  return [];
}

async function loadData() {
  const vocabStatus = document.getElementById("vocab-status");
  const grammarStatus = document.getElementById("grammar-status");
  const speakingStatus = document.getElementById("speaking-status");

  try {
    const [vText, gText] = await Promise.all([fetchCsv(GVIZ("Vocabulary")), fetchCsv(GVIZ("Grammar"))]);

    const vocabParsed = parseCsv(vText);
    if (vocabParsed.errors.length) console.warn(vocabParsed.errors);
    vocabRows = vocabParsed.data.filter((r) => vocabEnglishForPrompt(r));
    vocabStatus.classList.add("hidden");
    document.getElementById("vocab-root").classList.remove("hidden");
    renderVocabulary(vocabRows);

    const gParsed = parseCsv(gText);
    grammarRows = gParsed.data.filter((r) => (r.question || "").trim());
    grammarAnswers = grammarRows.map(() => null);
    grammarStatus.classList.add("hidden");
    document.getElementById("grammar-root").classList.remove("hidden");
    grammarIndex = 0;
    buildGrammarCard();

    const speakingRows = await fetchSpeakingRows();
    speakingStatus.classList.add("hidden");
    document.getElementById("speaking-root").classList.remove("hidden");
    renderSpeaking(speakingRows);
  } catch (e) {
    console.error(e);
    const msg =
      "スプレッドシートの取得に失敗しました。ネットワーク接続を確認するか、共有設定を確認してください。";
    [vocabStatus, grammarStatus, speakingStatus].forEach((el) => {
      el.textContent = msg;
      el.classList.add("status--error");
    });
  }
}

function startTest() {
  const input = document.getElementById("participant-name");
  const name = (input.value || "").trim();
  const err = document.getElementById("intro-error");
  if (!name) {
    err.hidden = false;
    input.focus();
    return;
  }
  err.hidden = true;
  participantName = name;
  document.getElementById("participant-display").textContent = `実施者：${name}`;
  document.getElementById("intro").classList.add("hidden");
  document.getElementById("intro").setAttribute("hidden", "");
  const flow = document.getElementById("test-flow");
  flow.classList.remove("hidden");
  flow.removeAttribute("hidden");
  flow.hidden = false;
  setStep(0);
}

document.getElementById("btn-start").addEventListener("click", startTest);
document.getElementById("participant-name").addEventListener("keydown", (e) => {
  if (e.key === "Enter") startTest();
});

document.getElementById("btn-prev").addEventListener("click", () => {
  if (currentStep === 1) saveCurrentGrammarSelection();
  setStep(currentStep - 1);
});
document.getElementById("btn-next").addEventListener("click", async () => {
  if (currentStep < 2) {
    await tryGoForwardTo(currentStep + 1);
    return;
  }
  submitResults();
});

document.querySelectorAll("#test-flow .stepper__btn").forEach((btn) => {
  btn.addEventListener("click", async () => {
    if (currentStep === 1) saveCurrentGrammarSelection();
    const go = Number(btn.getAttribute("data-go"));
    if (Number.isNaN(go)) return;
    if (go < currentStep) {
      setStep(go);
      return;
    }
    if (go === currentStep) return;
    await tryGoForwardTo(go);
  });
});

loadData()
  .then(() => {
    allowSessionPersist = true;
    restoreSessionState();
    document.documentElement.removeAttribute("data-shell");
    const nameEl = document.getElementById("participant-name");
    const vocabRoot = document.getElementById("vocab-root");
    if (nameEl) nameEl.addEventListener("input", persistSessionState);
    if (vocabRoot) vocabRoot.addEventListener("input", persistSessionState);
  })
  .catch((e) => console.error(e));

window.addEventListener("pagehide", persistSessionState);
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden") persistSessionState();
});
