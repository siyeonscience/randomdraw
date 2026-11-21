(() => {
  "use strict";

  const STORAGE_KEY = "randomdraw.words.v1";
  const MAX_VISUAL_SLIPS = 24;
  const SLIP_COLORS = ["#fffaf0", "#fff3c9", "#ffe3e0", "#e3f1ff", "#e6f7e1", "#f1e6ff"];
  const reduceMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  // ---------- 상태 ----------
  /** @type {{id:number, text:string, drawn:boolean}[]} */
  let words = [];
  let nextId = 1;
  let lastPickId = null; // 방금 뽑힌 쪽지
  let busy = false;

  // ---------- 엘리먼트 ----------
  const $ = (id) => document.getElementById(id);
  const stage = $("stage");
  const pile = $("pile");
  const stageEmpty = $("stageEmpty");
  const result = $("result");
  const paper = $("paper");
  const resultAnnounce = $("resultAnnounce");
  const statusEl = $("status");
  const mainActions = $("mainActions");
  const resultActions = $("resultActions");
  const drawBtn = $("drawBtn");
  const redrawBtn = $("redrawBtn");
  const moreBtn = $("moreBtn");
  const messageEl = $("message");
  const textInput = $("textInput");
  const addTextBtn = $("addTextBtn");
  const fileInput = $("fileInput");
  const fileDrop = $("fileDrop");
  const resetBtn = $("resetBtn");
  const wordList = $("wordList");
  const listEmpty = $("listEmpty");
  const wordCount = $("wordCount");
  const drawnBox = $("drawnBox");
  const drawnCount = $("drawnCount");
  const restoreBtn = $("restoreBtn");

  // ---------- 유틸 ----------
  const sleep = (ms) => new Promise((r) => setTimeout(r, reduceMotion ? Math.min(ms, 30) : ms));
  const rand = (min, max) => Math.random() * (max - min) + min;

  function randomInt(n) {
    if (window.crypto && crypto.getRandomValues) {
      const buf = new Uint32Array(1);
      const limit = Math.floor(0x100000000 / n) * n; // 편향 제거
      do { crypto.getRandomValues(buf); } while (buf[0] >= limit);
      return buf[0] % n;
    }
    return Math.floor(Math.random() * n);
  }

  function candidates() {
    return words.filter((w) => !w.drawn);
  }

  function showMessage(text, isError = true) {
    messageEl.textContent = text;
    messageEl.style.color = isError ? "" : "var(--ink-soft)";
  }

  // ---------- 저장 ----------
  function save() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(words.map(({ text, drawn }) => ({ text, drawn }))));
    } catch (_) { /* 저장 불가 환경은 무시 */ }
  }

  function load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const arr = JSON.parse(raw);
      if (!Array.isArray(arr)) return;
      words = arr
        .filter((w) => w && typeof w.text === "string" && w.text.trim())
        .map((w) => ({ id: nextId++, text: w.text, drawn: !!w.drawn }));
    } catch (_) { words = []; }
  }

  // ---------- 단어 등록 ----------
  function parseText(str) {
    // , 또는 / 로 구분 (줄바꿈도 구분자로 허용)
    return str.split(/[,/\n\r]+/).map((s) => s.trim()).filter(Boolean);
  }

  function addWords(list) {
    const added = list.map((text) => ({ id: nextId++, text, drawn: false }));
    words.push(...added);
    save();
    renderAll();
    return added.length;
  }

  addTextBtn.addEventListener("click", () => {
    const list = parseText(textInput.value);
    if (!list.length) {
      showMessage("추가할 단어를 입력해 주세요.");
      textInput.focus();
      return;
    }
    const n = addWords(list);
    textInput.value = "";
    showMessage(`쪽지 ${n}장을 추가했어요.`, false);
  });

  textInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      addTextBtn.click();
    }
  });

  async function readSpreadsheet(file) {
    if (typeof XLSX === "undefined") {
      throw new Error("엑셀 처리 라이브러리를 불러오지 못했어요. 인터넷 연결을 확인해 주세요.");
    }
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const out = [];
    wb.SheetNames.forEach((name) => {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: "" });
      rows.forEach((row) => {
        row.forEach((cell) => {
          const t = String(cell ?? "").trim();
          if (t) out.push(t);
        });
      });
    });
    return out;
  }

  async function handleFile(file) {
    if (!file) return;
    try {
      const list = await readSpreadsheet(file);
      if (!list.length) {
        showMessage("파일에서 단어를 찾지 못했어요.");
        return;
      }
      const n = addWords(list);
      showMessage(`'${file.name}'에서 쪽지 ${n}장을 추가했어요.`, false);
    } catch (err) {
      console.error(err);
      showMessage(err.message || "파일을 읽지 못했어요.");
    } finally {
      fileInput.value = "";
    }
  }

  fileInput.addEventListener("change", () => handleFile(fileInput.files[0]));
  ["dragenter", "dragover"].forEach((ev) =>
    fileDrop.addEventListener(ev, (e) => { e.preventDefault(); fileDrop.classList.add("dragover"); })
  );
  ["dragleave", "drop"].forEach((ev) =>
    fileDrop.addEventListener(ev, () => fileDrop.classList.remove("dragover"))
  );
  fileDrop.addEventListener("drop", (e) => {
    e.preventDefault();
    handleFile(e.dataTransfer.files[0]);
  });

  resetBtn.addEventListener("click", () => {
    if (busy) return;
    if (!words.length) return;
    if (!confirm("등록된 단어를 모두 지울까요?")) return;
    words = [];
    lastPickId = null;
    save();
    showIdle();
    renderAll();
    showMessage("단어 목록을 초기화했어요.", false);
  });

  restoreBtn.addEventListener("click", () => {
    if (busy) return;
    words.forEach((w) => { w.drawn = false; });
    save();
    renderAll();
    showMessage("제외된 쪽지를 모두 후보로 되돌렸어요.", false);
  });

  wordList.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-id]");
    if (!btn || busy) return;
    const id = Number(btn.dataset.id);
    words = words.filter((w) => w.id !== id);
    if (lastPickId === id) { lastPickId = null; showIdle(); }
    save();
    renderAll();
  });

  // ---------- 렌더 ----------
  function renderList() {
    wordList.innerHTML = "";
    const frag = document.createDocumentFragment();
    words.forEach((w) => {
      const li = document.createElement("li");
      li.className = "chip" + (w.drawn ? " drawn" : "");
      const span = document.createElement("span");
      span.textContent = w.text;
      span.title = w.text;
      const del = document.createElement("button");
      del.type = "button";
      del.dataset.id = w.id;
      del.setAttribute("aria-label", `'${w.text}' 삭제`);
      del.textContent = "×";
      li.append(span, del);
      frag.appendChild(li);
    });
    wordList.appendChild(frag);
    listEmpty.hidden = words.length > 0;
    wordCount.textContent = words.length;

    const drawnN = words.filter((w) => w.drawn).length;
    drawnBox.hidden = drawnN === 0;
    drawnCount.textContent = drawnN;
  }

  function renderStatus() {
    const left = candidates().length;
    statusEl.innerHTML = words.length
      ? `남은 쪽지 <b>${left}</b>장 / 전체 ${words.length}장`
      : "남은 쪽지 0장";
    drawBtn.disabled = busy || left === 0;
    redrawBtn.disabled = busy;
    // 방금 뽑힌 쪽지를 빼면 남는 후보가 있어야 '추가로 뽑기' 가능
    const lastIsCandidate = words.some((w) => w.id === lastPickId && !w.drawn);
    moreBtn.disabled = busy || left - (lastIsCandidate ? 1 : 0) <= 0;
    stageEmpty.hidden = words.length > 0 || !result.hidden;
  }

  function renderAll() {
    renderList();
    renderStatus();
    if (result.hidden && !busy) renderPile();
  }

  function stageSize() {
    return { w: stage.clientWidth, h: stage.clientHeight };
  }

  function placeRandom(slip) {
    const { w, h } = stageSize();
    const x = rand(16, Math.max(16, w - 88));
    const y = rand(16, Math.max(16, h - 62));
    const r = rand(-40, 40);
    slip.style.transform = `translate(${x}px, ${y}px) rotate(${r}deg)`;
    slip.style.zIndex = String(Math.floor(rand(1, 50)));
  }

  function renderPile() {
    pile.innerHTML = "";
    const n = Math.min(candidates().length, MAX_VISUAL_SLIPS);
    for (let i = 0; i < n; i++) {
      const s = document.createElement("div");
      s.className = "slip";
      s.style.setProperty("--slip-color", SLIP_COLORS[i % SLIP_COLORS.length]);
      placeRandom(s);
      pile.appendChild(s);
    }
  }

  function setWord(text) {
    paper.querySelectorAll("[data-word]").forEach((el) => { el.textContent = text; });
    // 글자 수에 따라 크기 조절
    const len = [...text].length;
    let size = "3.6rem";
    if (len > 4) size = "3rem";
    if (len > 8) size = "2.4rem";
    if (len > 14) size = "1.9rem";
    if (len > 24) size = "1.5rem";
    paper.style.setProperty("--word-size", size);
  }

  function showIdle() {
    result.hidden = true;
    pile.hidden = false;
    mainActions.hidden = false;
    resultActions.hidden = true;
  }

  // ---------- 뽑기 ----------
  /**
   * @param {"first"|"again"|"more"} mode
   *  - again: 방금 뽑힌 쪽지를 후보에 포함한 채로 다시 뽑기
   *  - more : 방금 뽑힌 쪽지를 후보에서 제외하고 뽑기
   */
  async function draw(mode) {
    if (busy) return;

    if (mode === "more" && lastPickId != null) {
      const last = words.find((w) => w.id === lastPickId);
      if (last) last.drawn = true;
      save();
    }

    const pool = candidates();
    if (!pool.length) {
      renderAll();
      showMessage(words.length ? "남은 쪽지가 없어요. 제외된 쪽지를 되돌리거나 단어를 추가해 주세요." : "먼저 단어를 등록해 주세요.");
      return;
    }

    busy = true;
    showMessage("");
    const pick = pool[randomInt(pool.length)];

    // 결과 숨기고 쪽지 더미 다시 깔기
    result.hidden = true;
    mainActions.hidden = true;
    resultActions.hidden = true;
    pile.hidden = false;
    pile.classList.remove("gather");
    renderList();
    renderStatus();
    renderPile();
    stageEmpty.hidden = true;

    const slips = [...pile.children];
    await sleep(60);

    // 1) 섞기
    const rounds = 9;
    for (let i = 0; i < rounds; i++) {
      slips.forEach(placeRandom);
      await sleep(i < rounds - 3 ? 170 : 240);
    }

    // 2) 가운데로 모으기
    const { w, h } = stageSize();
    pile.classList.add("gather");
    slips.forEach((s) => {
      const x = w / 2 - 36 + rand(-30, 30);
      const y = h / 2 - 23 + rand(-22, 22);
      s.style.transform = `translate(${x}px, ${y}px) rotate(${rand(-30, 30)}deg)`;
    });
    await sleep(520);

    // 3) 한 장이 튀어나오기
    const chosen = slips[randomInt(slips.length)];
    chosen.classList.add("chosen");
    chosen.style.zIndex = "100";
    chosen.style.transform = `translate(${w / 2 - 36}px, ${h / 2 - 23}px) rotate(0deg) scale(2.2)`;
    slips.forEach((s) => { if (s !== chosen) s.classList.add("fade"); });
    await sleep(650);

    // 4) 화면 중앙에 쪽지 등장 → 펼치기
    pile.hidden = true;
    setWord(pick.text);
    paper.classList.add("folded");
    // 팝인 애니메이션 재시작
    const wrap = paper.parentElement;
    wrap.style.animation = "none";
    void wrap.offsetWidth;
    wrap.style.animation = "";
    result.hidden = false;
    await sleep(600);

    paper.classList.remove("folded");
    await sleep(900);

    lastPickId = pick.id;
    resultAnnounce.textContent = `뽑힌 단어: ${pick.text}`;
    resultActions.hidden = false;
    busy = false;
    renderStatus();
  }

  drawBtn.addEventListener("click", () => draw("first"));
  redrawBtn.addEventListener("click", () => draw("again"));
  moreBtn.addEventListener("click", () => draw("more"));

  let resizeTimer;
  window.addEventListener("resize", () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => { if (!busy && result.hidden) renderPile(); }, 150);
  });

  // ---------- 시작 ----------
  load();
  renderAll();
})();
