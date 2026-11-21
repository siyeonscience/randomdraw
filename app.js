(() => {
  "use strict";

  const STORAGE_KEY = "randomdraw.words.v1";
  const MAX_VISUAL_SLIPS = 24;
  const SLIP_W = 60;
  const SLIP_H = 43;
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
  const wordCount = $("wordCount");
  const drawnBox = $("drawnBox");
  const drawnCount = $("drawnCount");
  const restoreBtn = $("restoreBtn");
  const setupScreen = $("setupScreen");
  const drawScreen = $("drawScreen");
  const setupMessage = $("setupMessage");
  const doneBtn = $("doneBtn");
  const editBtn = $("editBtn");
  let screen = "setup";

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
    const el = screen === "setup" ? setupMessage : messageEl;
    el.textContent = text;
    el.style.color = isError ? "" : "var(--ink-soft)";
  }

  // ---------- 화면 전환 ----------
  function showScreen(name) {
    screen = name;
    setupScreen.hidden = name !== "setup";
    drawScreen.hidden = name !== "draw";
    messageEl.textContent = "";
    setupMessage.textContent = "";
    if (name === "draw") {
      lastPickId = null;
      showIdle();
    }
    renderAll();
    window.scrollTo(0, 0);
  }

  doneBtn.addEventListener("click", () => {
    if (!words.length) {
      showMessage("먼저 단어를 하나 이상 등록해 주세요.");
      return;
    }
    showScreen("draw");
  });

  editBtn.addEventListener("click", () => {
    if (busy) return;
    showScreen("setup");
  });

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
    showMessage("");
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
      showMessage("");
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

  // 확인 창(confirm) 대신 버튼을 한 번 더 누르면 지워지는 방식
  let resetArmed = false;
  let resetTimer;
  function disarmReset() {
    resetArmed = false;
    clearTimeout(resetTimer);
    resetBtn.textContent = "단어 목록 초기화";
    resetBtn.classList.remove("armed");
  }

  resetBtn.addEventListener("click", () => {
    if (busy) return;
    if (!words.length) return;
    if (!resetArmed) {
      resetArmed = true;
      resetBtn.textContent = "한 번 더 누르면 모두 삭제";
      resetBtn.classList.add("armed");
      resetTimer = setTimeout(disarmReset, 3000);
      return;
    }
    disarmReset();
    words = [];
    lastPickId = null;
    save();
    showIdle();
    renderAll();
    showMessage("");
  });

  restoreBtn.addEventListener("click", () => {
    if (busy) return;
    words.forEach((w) => { w.drawn = false; });
    save();
    renderAll();
    showMessage("");
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
    editBtn.disabled = busy;
    doneBtn.textContent = words.length ? `등록 완료 · 쪽지 ${words.length}장으로 뽑으러 가기` : "등록 완료";
  }

  function renderAll() {
    renderList();
    renderStatus();
    if (screen === "draw" && result.hidden && !busy) renderPile();
  }

  function stageSize() {
    return { w: stage.clientWidth, h: stage.clientHeight };
  }

  function placeRandom(slip) {
    const { w, h } = stageSize();
    const x = rand(16, Math.max(16, w - SLIP_W - 22));
    const y = rand(16, Math.max(16, h - SLIP_H - 22));
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
      placeRandom(s);
      pile.appendChild(s);
    }
  }

  function getPaperWidth() {
    return Math.min(340, window.innerWidth * 0.8);
  }

  function setWord(text) {
    paper.querySelectorAll("[data-word]").forEach((el) => { el.textContent = text; });
    // 글자 수에 따라 크기 조절
    const len = [...text].length;
    let size = "4.4rem";
    if (len > 3) size = "3.6rem";
    if (len > 5) size = "2.8rem";
    if (len > 8) size = "2.2rem";
    if (len > 14) size = "1.8rem";
    if (len > 24) size = "1.4rem";
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
      showMessage(words.length ? "남은 쪽지가 없어요. '단어 수정'에서 제외된 쪽지를 되돌리거나 단어를 추가해 주세요." : "먼저 단어를 등록해 주세요.");
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
      const x = w / 2 - SLIP_W / 2 + rand(-30, 30);
      const y = h / 2 - SLIP_H / 2 + rand(-22, 22);
      s.style.transform = `translate(${x}px, ${y}px) rotate(${rand(-30, 30)}deg)`;
    });
    await sleep(520);

    // 3) 한 장이 튀어나오기
    const chosen = slips[randomInt(slips.length)];
    chosen.classList.add("chosen");
    chosen.style.zIndex = "100";
    // 결과 쪽지의 접힌 크기(종이의 1/4)와 맞춰서 키움
    const scale = (getPaperWidth() / 2) / SLIP_W;
    chosen.style.transform = `translate(${w / 2 - SLIP_W / 2}px, ${h / 2 - SLIP_H / 2}px) rotate(0deg) scale(${scale})`;
    slips.forEach((s) => { if (s !== chosen) s.classList.add("fade"); });
    await sleep(650);

    // 4) 같은 자리에서 두 번 접힌 쪽지로 바꿔치기 → 두 단계로 펼치기
    setWord(pick.text);
    paper.classList.remove("f1", "taped");
    paper.classList.add("f2");
    result.hidden = false;
    pile.hidden = true;
    await sleep(250);

    paper.classList.replace("f2", "f1"); // 왼쪽 절반 펼치기
    await sleep(850);
    paper.classList.remove("f1"); // 위쪽 절반 펼치기
    await sleep(850);
    paper.classList.add("taped");

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
    resizeTimer = setTimeout(() => { if (screen === "draw" && !busy && result.hidden) renderPile(); }, 150);
  });

  // ---------- 시작 ----------
  // 손글씨 폰트를 미리 받아 둬서 쪽지가 펼쳐질 때 글꼴이 바뀌어 보이지 않게
  if (document.fonts && document.fonts.load) {
    document.fonts.load('1em "Kyobo Handwriting 2025"', "가").catch(() => {});
  }

  load();
  // 저장된 단어가 있으면 바로 뽑기 화면, 없으면 등록 화면부터
  showScreen(words.length ? "draw" : "setup");
})();
