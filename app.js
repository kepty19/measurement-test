const SPREADSHEET_ID = "19GdI5qQWc-VyLQEgRiJfaxxTLwrtxU6ofx4tZYPax0M";
const GVIZ = (sheet) =>
  `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(
    sheet
  )}`;
const GVIZ_JSON = (sheet) =>
  `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(
    sheet
  )}`;

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

const HINTS = [
  'パート 1「単語」を実施中',
  'パート 2「文法」を実施中',
  'パート 3「スピーキング」を実施中',
];

let participantName = "";
let currentStep = 0;
let vocabRows = [];
let grammarRows = [];
let grammarAnswers = [];
let grammarIndex = 0;
let speakingTopicCount = 0;

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
  document.getElementById("btn-next").textContent = isLast ? "結果を送信" : "次へ進む";
  document.getElementById("footer-hint").textContent = HINTS[currentStep];
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
    const en = (row.english || "").trim();
    const ja = (row.japanese || "").trim();
    const level = (row.level || "").trim();
    if (!en) return;
    const id = `vocab-${n++}`;
    const item = document.createElement("article");
    item.className = "vocab__item";
    item.innerHTML = `
      <p class="vocab__prompt">${escapeHtml(en)}</p>
      ${level ? `<p class="vocab__meta">${escapeHtml(level)}</p>` : ""}
      <label class="vocab__label" for="${id}">日本語の意味を入力</label>
      <textarea id="${id}" class="vocab__input" rows="2" autocomplete="off" spellcheck="false" placeholder="例：パン"></textarea>
      <div class="vocab__check">
        <button type="button" class="btn--small vocab__reveal">答え合わせ（正解の目安を表示）</button>
        <p class="vocab__answer" role="status"></p>
      </div>
    `;
    const btn = item.querySelector(".vocab__reveal");
    const ans = item.querySelector(".vocab__answer");
    btn.addEventListener("click", () => {
      ans.textContent = ja ? `目安：${ja}` : "（スプレッドシートに正解がありません）";
      ans.classList.add("is-visible");
    });
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
      <span class="grammar__progress">問題 ${grammarIndex + 1} / ${grammarRows.length}</span>
    </div>
    <div class="grammar__card">
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
      saveCurrentGrammarSelection();
      if (grammarIndex < grammarRows.length - 1) {
        grammarIndex += 1;
        buildGrammarCard();
      }
    });
  }
}

function trainingTopicFromRow(row) {
  if (!row || typeof row !== "object") return "";
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
  const cols = (payload.table.cols || []).map((c) => (c.label || "").trim());
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
  if (hit) return String(row[hit] || "").trim();
  return String(row.theme || "").trim();
}

function renderSpeaking(rows) {
  const root = document.getElementById("speaking-root");
  root.innerHTML = "";
  const dataRows = rows.filter((r) => trainingTopicFromRow(r).trim().length > 0);
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
    const en = (row.english || "").trim();
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

function showToast(message, isError) {
  const el = document.getElementById("submit-toast");
  el.textContent = message;
  el.classList.remove("hidden");
  el.removeAttribute("hidden");
  el.classList.toggle("toast--error", !!isError);
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

  if (!SUBMIT_URL || !SUBMIT_URL.trim()) {
    downloadPayload(payload);
    showToast(
      "送信先 URL が未設定のため、結果を JSON でダウンロードしました。config.js に Apps Script の URL を設定してください。",
      false
    );
    btn.disabled = false;
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
    showToast("送信しました。スプレッドシートの「Result」「raw data」をご確認ください。", false);
  } catch (e) {
    console.error(e);
    downloadPayload(payload);
    showToast(
      "サーバー送信に失敗したため、結果を JSON で保存しました。ネットワーク・CORS・Apps Script のデプロイを確認してください。",
      true
    );
  } finally {
    btn.disabled = false;
  }
}

async function loadData() {
  const vocabStatus = document.getElementById("vocab-status");
  const grammarStatus = document.getElementById("grammar-status");
  const speakingStatus = document.getElementById("speaking-status");

  try {
    const [vText, gText, sJsonText] = await Promise.all([
      fetchCsv(GVIZ("Vocabulary")),
      fetchCsv(GVIZ("Grammar")),
      fetchCsv(GVIZ_JSON("Speaking")),
    ]);

    const vocabParsed = parseCsv(vText);
    if (vocabParsed.errors.length) console.warn(vocabParsed.errors);
    vocabRows = vocabParsed.data.filter((r) => (r.english || "").trim());
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

    let speakingRows = parseGvizJson(sJsonText);
    if (!speakingRows.length) {
      const sCsv = await fetchCsv(GVIZ("Speaking"));
      const sParsed = parseCsv(sCsv);
      speakingRows = sParsed.data || [];
    }
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
document.getElementById("btn-next").addEventListener("click", () => {
  if (currentStep === 1) saveCurrentGrammarSelection();
  if (currentStep < 2) setStep(currentStep + 1);
  else submitResults();
});

document.querySelectorAll("#test-flow .stepper__btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    if (currentStep === 1) saveCurrentGrammarSelection();
    const go = Number(btn.getAttribute("data-go"));
    if (!Number.isNaN(go)) setStep(go);
  });
});

loadData();
