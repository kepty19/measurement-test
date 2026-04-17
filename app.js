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
 * 4) レーベンシュタイン類似度（厳しめ → やや緩めの二段）
 * 5) 同義・類義クラスタ（品詞ゆれ・別表現）— 正解文のいずれかと同じグループなら正解
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

/** 正解文・学習者解答のどちらかと同じグループにあれば正解（品詞ゆれ・別表現・よくある言い換え） */
const VOCAB_SYNONYM_CLUSTERS = [
  ["贈り物", "プレゼント", "ギフト"],
  ["退屈", "退屈な", "暇", "つまらない", "うんざり"],
  ["起こる", "発生", "生じる"],
  ["栄養に富む", "栄養のある", "栄養素", "栄養", "栄養満点"],
  ["変形させる", "変形する", "変える", "変わる", "変化する", "変化", "変形"],
  ["鋭い", "するどい", "鋭敏"],
  ["痛み", "痛い", "苦痛", "疼痛"],
  ["若さ", "若い", "青年", "若者"],
  ["好奇心が強い", "好奇心", "興味", "興味深い"],
  ["人気のある", "人気"],
];

function clusterSynonymMatch(userText, expectedJa) {
  const u = normalizeJa(userText);
  if (!u || u.length < 2) return false;
  const parts = splitAnswerVariants(expectedJa);
  const expChunks = [
    ...parts.map((p) => normalizeJa(String(p).trim())),
    normalizeJa(expectedJa),
  ].filter((x) => x && x.length > 0);

  for (const cluster of VOCAB_SYNONYM_CLUSTERS) {
    const norms = cluster.map((c) => normalizeJa(c)).filter((n) => n.length > 0);
    const userHits = norms.some((n) => {
      if (u === n) return true;
      if (u.length >= 2 && (u.includes(n) || n.includes(u))) return true;
      return false;
    });
    if (!userHits) continue;
    const expHits = expChunks.some((exp) =>
      norms.some((n) => {
        if (!exp || !n) return false;
        if (exp === n) return true;
        if (exp.length >= 2 && n.length >= 2 && (exp.includes(n) || n.includes(exp))) return true;
        return false;
      })
    );
    if (expHits) return true;
  }
  return false;
}

const SIM_STRICT = 0.78;
const SIM_RELAXED = 0.64;

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
    if (similarityRatio(u, p) >= SIM_STRICT) return true;
    if (pp.length >= 3 && similarityRatio(uPlain, pp) >= SIM_STRICT) return true;
  }
  if (full.length >= 3 && similarityRatio(u, full) >= SIM_STRICT) return true;
  if (fullPlain.length >= 3 && similarityRatio(uPlain, fullPlain) >= SIM_STRICT) return true;

  if (clusterSynonymMatch(userText, expectedJa)) return true;

  for (const p of parts) {
    if (!p) continue;
    const pn = normalizeJa(p);
    if (pn.length >= 2 && similarityRatio(u, pn) >= SIM_RELAXED) return true;
    const pp = stripNoise(p);
    if (pp.length >= 2 && similarityRatio(uPlain, pp) >= SIM_RELAXED) return true;
  }
  if (full.length >= 2 && similarityRatio(u, full) >= SIM_RELAXED) return true;
  if (fullPlain.length >= 2 && similarityRatio(uPlain, fullPlain) >= SIM_RELAXED) return true;

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
    sessionStorage.setItem(SESSI