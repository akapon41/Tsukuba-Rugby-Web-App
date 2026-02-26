(function () {
  const STORAGE_PREFIX = "blues-player-dashboard";
  const GOAL_FIELDS = [
    { name: "dfTackles", label: "DF タックル数", getter: (g) => g.df?.tackleCount },
    { name: "dfTackleRate", label: "DF タックル成功率", getter: (g) => g.df?.tackleSuccessRate, rate: true },
    { name: "atGoodStayRate", label: "AT Good/Stay率", getter: (g) => g.at?.goodStayRate, rate: true },
    { name: "atSupportRate", label: "AT Support率", getter: (g) => g.at?.supportRate, rate: true },
    { name: "atB2Rate", label: "AT B2率", getter: (g) => g.at?.b2Rate, rate: true },
  ];
  const PLAYER_ROW_START_INDEX = 7; // Excel row 8
  // `sheet_to_json(..., header:1)` is based on the sheet's `!ref`.
  // These Individual sheets start at column B, so AoA index 0 == Excel B列, 1 == C列.
  const PLAYER_NAME_COL_INDEX = 1; // Excel column C
  const PLAYER_JERSEY_COL_INDEX = 0; // Excel column B

  const state = {
    games: [],
    players: [],
    selectedPlayerId: "",
    charts: { volume: null, rate: null },
    datasetVersion: "empty",
  };

  const el = {
    fileInput: document.getElementById("file-input"),
    loadSampleBtn: document.getElementById("load-sample-btn"),
    loadStatus: document.getElementById("load-status"),
    dataStatusPill: document.getElementById("data-status-pill"),
    playerSelect: document.getElementById("player-select"),
    showDashboardBtn: document.getElementById("show-dashboard-btn"),
    dashboard: document.getElementById("dashboard"),
    playerTitle: document.getElementById("player-title"),
    summaryCards: document.getElementById("summary-cards"),
    goalForm: document.getElementById("goal-form"),
    goalSummary: document.getElementById("goal-summary"),
    matchesTableBody: document.getElementById("matches-table-body"),
    volumeChart: document.getElementById("volume-chart"),
    rateChart: document.getElementById("rate-chart"),
    volumeChartCard: document.getElementById("volume-chart-card"),
    rateChartCard: document.getElementById("rate-chart-card"),
    playerLikeBtn: document.getElementById("player-like-btn"),
    playerLikeCount: document.getElementById("player-like-count"),
    commentForm: document.getElementById("comment-form"),
    commentInput: document.getElementById("comment-input"),
    commentsList: document.getElementById("comments-list"),
  };

  function init() {
    el.fileInput.addEventListener("change", onFileInputChange);
    el.loadSampleBtn.addEventListener("click", onLoadSampleClick);
    el.showDashboardBtn.addEventListener("click", onShowDashboard);
    el.playerSelect.addEventListener("change", () => {
      state.selectedPlayerId = el.playerSelect.value;
      el.showDashboardBtn.disabled = !state.selectedPlayerId;
      if (!el.dashboard.hidden && state.selectedPlayerId) {
        renderSelectedPlayerDashboard();
      }
    });
    el.goalForm.addEventListener("input", renderSelectedPlayerDashboard);
    el.playerLikeBtn.addEventListener("click", onPlayerLike);
    el.commentForm.addEventListener("submit", onCommentSubmit);
    el.commentsList.addEventListener("click", onCommentListClick);
  }

  async function onLoadSampleClick() {
    setLoadingState(true, "サンプルファイルを読み込み中...");
    try {
      const response = await fetch("./Blues_2025.xlsx");
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });
      const parsed = parseWorkbook(workbook, "Blues_2025.xlsx");
      applyParsedGames(parsed, ["Blues_2025.xlsx"]);
    } catch (error) {
      setStatus(
        `サンプル読込に失敗しました。ローカルサーバー経由で開いているか確認してください（例: python3 -m http.server）。詳細: ${error.message}`,
        true
      );
      setLoadingState(false);
    }
  }

  async function onFileInputChange(event) {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;

    setLoadingState(true, `${files.length}件のファイルを読み込み中...`);
    try {
      const allGames = [];
      for (const file of files) {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });
        allGames.push(...parseWorkbook(workbook, file.name));
      }
      applyParsedGames(allGames, files.map((f) => f.name));
    } catch (error) {
      setStatus(`ファイル読込に失敗しました: ${error.message}`, true);
      setLoadingState(false);
    } finally {
      event.target.value = "";
    }
  }

  function applyParsedGames(games, fileNames) {
    const cleaned = dedupeAndSortGames(games);
    state.games = cleaned;
    state.players = buildPlayersFromGames(cleaned);
    state.datasetVersion = createDatasetVersion(cleaned);

    populatePlayerSelect(state.players);
    const msg = `${fileNames.join(", ")} を読み込み: ${cleaned.length}試合 / ${state.players.length}選手`;
    setStatus(msg, false);
    setLoadingState(false);

    if (state.players.length) {
      el.playerSelect.value = state.players[0].id;
      state.selectedPlayerId = state.players[0].id;
      el.showDashboardBtn.disabled = false;
      el.playerSelect.disabled = false;
    } else {
      el.playerSelect.disabled = true;
      el.showDashboardBtn.disabled = true;
      el.dashboard.hidden = true;
    }
  }

  function setLoadingState(isLoading, text) {
    el.loadSampleBtn.disabled = isLoading;
    el.fileInput.disabled = isLoading;
    if (typeof text === "string") {
      setStatus(text, false);
    }
    el.dataStatusPill.textContent = isLoading ? "読込中..." : state.games.length ? "読込済み" : "未読込";
  }

  function setStatus(message, isError) {
    el.loadStatus.textContent = message;
    el.loadStatus.classList.toggle("ng", Boolean(isError));
  }

  function populatePlayerSelect(players) {
    el.playerSelect.innerHTML = "";
    if (!players.length) {
      el.playerSelect.innerHTML = '<option value="">選手が見つかりません</option>';
      return;
    }
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "選手を選択してください";
    el.playerSelect.appendChild(placeholder);
    for (const player of players) {
      const opt = document.createElement("option");
      opt.value = player.id;
      opt.textContent = `${player.displayName} (${player.games.length}試合)`;
      el.playerSelect.appendChild(opt);
    }
  }

  function onShowDashboard() {
    state.selectedPlayerId = el.playerSelect.value;
    if (!state.selectedPlayerId) return;
    el.dashboard.hidden = false;
    renderSelectedPlayerDashboard();
    el.dashboard.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function renderSelectedPlayerDashboard() {
    if (!state.selectedPlayerId) return;
    const player = state.players.find((p) => p.id === state.selectedPlayerId);
    if (!player) return;

    el.playerTitle.textContent = `${player.displayName} 個人成績ダッシュボード`;
    renderSummaryCards(player);
    const goalConfig = readGoalConfig();
    const goalEvaluations = evaluateGoals(player, goalConfig);
    renderGoalSummary(goalEvaluations);
    renderMatchesTable(player, goalEvaluations.byGame);
    renderCharts(player);
    renderPlayerLike(player);
    renderComments(player);
  }

  function renderSummaryCards(player) {
    const games = player.games;
    const minutesList = games.map((g) => chooseNumber(g.at?.minutes, g.df?.minutes));
    const metrics = {
      matches: games.length,
      avgMinutes: average(minutesList),
      maxMinutes: Math.max(...minutesList.filter(isFiniteNumber), 0),
      blueCount: minutesList.filter((m) => getMinuteBand(m) === "blue").length,
    };

    const cards = [
      ["試合数", `${metrics.matches}`],
      ["平均出場時間", formatNumber(metrics.avgMinutes, 1)],
      ["最大出場時間", formatNumber(metrics.maxMinutes, 0)],
      ["青（長時間出場）", `${metrics.blueCount}`],
    ];

    el.summaryCards.innerHTML = "";
    for (const [label, value] of cards) {
      const card = document.createElement("div");
      card.className = "summary-card";
      card.innerHTML = `<div class="label">${escapeHtml(label)}</div><div class="value">${escapeHtml(value)}</div>`;
      el.summaryCards.appendChild(card);
    }

    const legend = document.createElement("div");
    legend.className = "minute-legend";
    legend.innerHTML = `
      <span class="minute-pill white">白: 0-19分</span>
      <span class="minute-pill sky">水色: 20-59分</span>
      <span class="minute-pill blue">青: 60分以上</span>
    `;
    el.summaryCards.appendChild(legend);
  }

  function readGoalConfig() {
    const formData = new FormData(el.goalForm);
    const config = {};
    for (const field of GOAL_FIELDS) {
      config[field.name] = toNumber(formData.get(field.name));
    }
    return config;
  }

  function evaluateGoals(player, goalConfig) {
    const byGame = player.games.map((game) => {
      const results = GOAL_FIELDS.map((field) => {
        const actual = field.getter(game);
        const target = goalConfig[field.name];
        return {
          name: field.name,
          label: field.label,
          actual,
          target,
          met: isFiniteNumber(actual) && isFiniteNumber(target) ? actual >= target : false,
        };
      });
      return {
        gameId: game.gameId,
        label: game.label,
        results,
        totalMet: results.filter((r) => r.met).length,
      };
    });

    const totals = GOAL_FIELDS.map((field) => {
      const achieved = byGame.filter((g) => g.results.find((r) => r.name === field.name)?.met).length;
      return {
        name: field.name,
        label: field.label,
        achieved,
        total: byGame.length,
      };
    });

    return { byGame, totals };
  }

  function renderGoalSummary(goalEvaluations) {
    const frag = document.createDocumentFragment();

    const headline = document.createElement("div");
    headline.className = "muted";
    headline.textContent = `目標達成回数（試合ごと判定）`;
    frag.appendChild(headline);

    const row = document.createElement("div");
    row.className = "goal-chip-row";
    for (const item of goalEvaluations.totals) {
      const ratio = item.total ? item.achieved / item.total : 0;
      const chip = document.createElement("span");
      chip.className = `goal-chip ${ratio >= 0.6 ? "good" : "bad"}`;
      chip.textContent = `${item.label}: ${item.achieved}/${item.total}`;
      row.appendChild(chip);
    }
    frag.appendChild(row);

    el.goalSummary.innerHTML = "";
    el.goalSummary.appendChild(frag);
  }

  function renderMatchesTable(player, goalByGame) {
    if (!el.matchesTableBody) return;
    void goalByGame;
    el.matchesTableBody.innerHTML = "";

    for (const game of player.games) {
      const minutes = chooseNumber(game.at?.minutes, game.df?.minutes);
      const band = getMinuteBand(minutes);
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(game.label)}</td>
        <td>${escapeHtml(formatNumber(minutes, 0))}分</td>
        <td><span class="minute-pill ${band.className}">${escapeHtml(band.label)}</span></td>
      `;
      el.matchesTableBody.appendChild(tr);
    }
  }

  function renderCharts(player) {
    const labels = player.games.map((g) => g.shortLabel);
    const commonOptions = {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: { legend: { position: "bottom" } },
      scales: {
        x: { ticks: { maxRotation: 0, autoSkip: true } },
      },
    };

    if (state.charts.volume) state.charts.volume.destroy();
    if (state.charts.rate) state.charts.rate.destroy();

    const minutes = player.games.map((g) => chooseNumber(g.at?.minutes, g.df?.minutes));
    el.volumeChartCard.querySelector("h3").textContent = "出場時間";
    el.rateChartCard.hidden = true;
    el.volumeChartCard.parentElement.classList.add("single");

    state.charts.volume = new Chart(el.volumeChart, {
      type: "bar",
      data: {
        labels,
        datasets: [{
          label: "出場時間",
          data: minutes.map((v) => (isFiniteNumber(v) ? v : null)),
          backgroundColor: minutes.map((v) => getMinuteBand(v).color),
          borderColor: minutes.map((v) => getMinuteBand(v).borderColor),
          borderWidth: 1.5,
          borderRadius: 8,
        }],
      },
      options: {
        ...commonOptions,
        plugins: { legend: { display: false } },
        scales: {
          ...commonOptions.scales,
          y: { beginAtZero: true, suggestedMax: 80, title: { display: true, text: "分" } },
        },
      },
    });
    state.charts.rate = null;
  }

  function dataset(label, values, color) {
    return {
      label,
      data: values.map((v) => (isFiniteNumber(v) ? v : null)),
      borderColor: color,
      backgroundColor: withAlpha(color, 0.16),
      tension: 0.28,
      fill: false,
      pointRadius: 3,
      pointHoverRadius: 5,
      spanGaps: true,
    };
  }

  function getMinuteBand(value) {
    if (!isFiniteNumber(value) || value < 20) {
      return {
        label: "白",
        className: "white",
        color: "rgba(255,255,255,0.95)",
        borderColor: "rgba(37,50,72,0.25)",
      };
    }
    if (value < 60) {
      return {
        label: "水色",
        className: "sky",
        color: "rgba(223,241,255,0.95)",
        borderColor: "rgba(15,79,134,0.35)",
      };
    }
    return {
      label: "青",
      className: "blue",
      color: "rgba(31,95,154,0.92)",
      borderColor: "rgba(17,63,104,0.9)",
    };
  }

  function renderPlayerLike(player) {
    const store = readJson(`${STORAGE_PREFIX}:playerLikes`, {});
    const count = Number(store[player.id] || 0);
    el.playerLikeCount.textContent = String(count);
  }

  function onPlayerLike() {
    if (!state.selectedPlayerId) return;
    const store = readJson(`${STORAGE_PREFIX}:playerLikes`, {});
    store[state.selectedPlayerId] = Number(store[state.selectedPlayerId] || 0) + 1;
    writeJson(`${STORAGE_PREFIX}:playerLikes`, store);
    const player = state.players.find((p) => p.id === state.selectedPlayerId);
    if (player) renderPlayerLike(player);
  }

  function onCommentSubmit(event) {
    event.preventDefault();
    if (!state.selectedPlayerId) return;
    const text = (el.commentInput.value || "").trim();
    if (!text) return;

    const comments = getPlayerComments(state.selectedPlayerId);
    comments.unshift({
      id: cryptoRandomId(),
      text,
      likes: 0,
      createdAt: new Date().toISOString(),
    });
    setPlayerComments(state.selectedPlayerId, comments);
    el.commentInput.value = "";
    const player = state.players.find((p) => p.id === state.selectedPlayerId);
    if (player) renderComments(player);
  }

  function onCommentListClick(event) {
    const btn = event.target.closest("button[data-action]");
    if (!btn || !state.selectedPlayerId) return;
    const id = btn.getAttribute("data-id");
    const action = btn.getAttribute("data-action");
    const comments = getPlayerComments(state.selectedPlayerId);
    const idx = comments.findIndex((c) => c.id === id);
    if (idx < 0) return;

    if (action === "like") {
      comments[idx].likes = Number(comments[idx].likes || 0) + 1;
    } else if (action === "delete") {
      comments.splice(idx, 1);
    } else {
      return;
    }

    setPlayerComments(state.selectedPlayerId, comments);
    const player = state.players.find((p) => p.id === state.selectedPlayerId);
    if (player) renderComments(player);
  }

  function renderComments(player) {
    const comments = getPlayerComments(player.id);
    el.commentsList.innerHTML = "";

    if (!comments.length) {
      const empty = document.createElement("div");
      empty.className = "muted small";
      empty.textContent = "まだコメントはありません。";
      el.commentsList.appendChild(empty);
      return;
    }

    for (const item of comments) {
      const card = document.createElement("article");
      card.className = "comment-item";
      card.innerHTML = `
        <div class="comment-meta">
          <span>${escapeHtml(formatDateTime(item.createdAt))}</span>
          <span>いいね ${escapeHtml(String(Number(item.likes || 0)))}</span>
        </div>
        <p class="comment-text">${escapeHtml(item.text)}</p>
        <div class="comment-actions">
          <button type="button" data-action="like" data-id="${escapeHtml(item.id)}">いいね</button>
          <button type="button" data-action="delete" data-id="${escapeHtml(item.id)}">削除</button>
        </div>
      `;
      el.commentsList.appendChild(card);
    }
  }

  function getPlayerComments(playerId) {
    const store = readJson(`${STORAGE_PREFIX}:comments`, {});
    return Array.isArray(store[playerId]) ? store[playerId] : [];
  }

  function setPlayerComments(playerId, comments) {
    const store = readJson(`${STORAGE_PREFIX}:comments`, {});
    store[playerId] = comments;
    writeJson(`${STORAGE_PREFIX}:comments`, store);
  }

  function parseWorkbook(workbook, sourceName) {
    const gamesByKey = new Map();
    let localSeq = 0;

    for (let sheetIndex = 0; sheetIndex < workbook.SheetNames.length; sheetIndex += 1) {
      const sheetName = workbook.SheetNames[sheetIndex];
      const match = sheetName.match(/^Individual\s+(AT|DF)\s+(.+)$/i);
      if (!match) continue;

      const side = match[1].toUpperCase();
      const sheetSuffix = (match[2] || "").trim();
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) continue;

      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: true,
        defval: null,
      });
      // AoA index is relative to sheet start column (these sheets start at B列)
      // row[1] == Excel C列 (opponent)
      const opponent = cleanText(rows?.[1]?.[1]);
      const gameKey = `${sourceName}::${sheetSuffix || sheetName}`;
      const info = parseGameInfo(sheetSuffix, opponent);

      let game = gamesByKey.get(gameKey);
      if (!game) {
        game = {
          key: gameKey,
          sourceName,
          sheetSuffix,
          opponent: info.opponent || opponent || "不明",
          order: localSeq++,
          dateCode: info.dateCode,
          dateLabel: info.dateLabel,
          label: info.label,
          playersById: new Map(),
        };
        gamesByKey.set(gameKey, game);
      }

      const records = parseIndividualSheetRows(rows, side);
      for (const rec of records) {
        const pid = rec.playerId;
        let player = game.playersById.get(pid);
        if (!player) {
          player = {
            id: pid,
            displayName: rec.displayName,
            jersey: rec.jersey,
            at: null,
            df: null,
          };
          game.playersById.set(pid, player);
        }
        if (!player.displayName || rec.displayName.length < player.displayName.length) {
          player.displayName = rec.displayName;
        }
        if (!player.jersey && rec.jersey) player.jersey = rec.jersey;
        player[side.toLowerCase()] = rec.stats;
      }
    }

    const games = [];
    for (const game of gamesByKey.values()) {
      const players = Array.from(game.playersById.values()).filter((p) => p.displayName);
      if (!players.length) continue;
      const realDataPlayers = players.filter((p) => hasMeaningfulStats(p.at) || hasMeaningfulStats(p.df));
      if (!realDataPlayers.length) continue;

      const label = game.label || buildGameLabel(game.dateLabel, game.opponent);
      games.push({
        gameId: game.key,
        sourceName: game.sourceName,
        sheetSuffix: game.sheetSuffix,
        order: game.order,
        dateCode: game.dateCode,
        dateLabel: game.dateLabel,
        opponent: game.opponent,
        label,
        shortLabel: shortGameLabel(game.dateLabel, game.opponent),
        players: realDataPlayers,
      });
    }

    return games;
  }

  function parseIndividualSheetRows(rows, side) {
    const out = [];
    let emptyRun = 0;
    for (let r = PLAYER_ROW_START_INDEX; r < rows.length; r += 1) {
      const row = rows[r] || [];
      const rawName = cleanText(row[PLAYER_NAME_COL_INDEX]); // C列
      const jersey = cleanText(row[PLAYER_JERSEY_COL_INDEX]); // B列

      if (!rawName) {
        emptyRun += 1;
        if (emptyRun >= 6 && r > 20) break;
        continue;
      }
      emptyRun = 0;

      if (/(合計|total|平均)/i.test(rawName)) continue;
      if (rawName === "0") continue;

      const displayName = extractDisplayName(rawName);
      const playerId = normalizePlayerName(rawName);
      if (!playerId) continue;

      const stats = side === "DF" ? parseDfStats(row) : parseAtStats(row);
      out.push({ playerId, displayName, jersey, stats });
    }
    return out;
  }

  function parseDfStats(row) {
    return {
      minutes: toNumber(row[2]), // D列
      plays: toNumber(row[3]),
      playsNormalized: toNumber(row[4]),
      tackleCount: toNumber(row[5]),
      tackleSuccessRate: clampRate(toNumber(row[6])),
      firstTackleTotal: toNumber(row[7]),
      firstTackleRate: clampRate(toNumber(row[8])),
      firstDefeat: toNumber(row[9]),
      firstTakedown: toNumber(row[10]),
      firstMiss: toNumber(row[11]),
      firstReactionRate: clampRate(toNumber(row[12])),
      firstFight: toNumber(row[13]),
      firstReload: toNumber(row[14]),
      firstReactionMiss: toNumber(row[15]),
      assistTotal: toNumber(row[16]),
      assistRate: clampRate(toNumber(row[17])),
      assistDefeat: toNumber(row[18]),
      assistTakedown: toNumber(row[19]),
      assistMiss: toNumber(row[20]),
      assistReactionRate: clampRate(toNumber(row[21])),
      assistFight: toNumber(row[22]),
      assistReload: toNumber(row[23]),
      assistReactionMiss: toNumber(row[24]),
    };
  }

  function parseAtStats(row) {
    return {
      minutes: toNumber(row[2]), // D列
      plays: toNumber(row[3]),
      playsNormalized: toNumber(row[4]),
      carrierContactTotal: toNumber(row[5]),
      goodStayRate: clampRate(toNumber(row[6])),
      bbRate: clampRate(toNumber(row[7])),
      gainCount: toNumber(row[8]),
      stayCount: toNumber(row[9]),
      backCount: toNumber(row[10]),
      lostCount: toNumber(row[11]),
      supportTotal: toNumber(row[12]),
      supportRate: clampRate(toNumber(row[13])),
      supportExcellent: toNumber(row[14]),
      supportGood: toNumber(row[15]),
      supportBad: toNumber(row[16]),
      b3Total: toNumber(row[17]),
      b3Rate: clampRate(toNumber(row[18])),
      b3CarrierOk: toNumber(row[19]),
      b3CarrierNg: toNumber(row[20]),
      b3SupportOk: toNumber(row[21]),
      b3SupportNg: toNumber(row[22]),
      b2Total: toNumber(row[23]),
      b2Rate: clampRate(toNumber(row[24])),
    };
  }

  function buildPlayersFromGames(games) {
    const map = new Map();
    for (const game of games) {
      for (const rec of game.players) {
        let player = map.get(rec.id);
        if (!player) {
          player = { id: rec.id, displayName: rec.displayName, jersey: rec.jersey, games: [] };
          map.set(rec.id, player);
        }
        if (!player.displayName || rec.displayName.length < player.displayName.length) {
          player.displayName = rec.displayName;
        }
        player.games.push({
          gameId: game.gameId,
          label: game.label,
          shortLabel: game.shortLabel,
          opponent: game.opponent,
          dateCode: game.dateCode,
          order: game.order,
          at: rec.at,
          df: rec.df,
          jersey: rec.jersey,
        });
      }
    }

    const players = Array.from(map.values());
    for (const player of players) {
      player.games.sort(compareGames);
    }
    players.sort((a, b) => a.displayName.localeCompare(b.displayName, "ja"));
    return players;
  }

  function dedupeAndSortGames(games) {
    const map = new Map();
    let seq = 0;
    for (const game of games) {
      const key = game.gameId;
      if (map.has(key)) continue;
      map.set(key, { ...game, globalOrder: seq++ });
    }
    return Array.from(map.values()).sort(compareGames);
  }

  function compareGames(a, b) {
    const aCode = a.dateCode;
    const bCode = b.dateCode;
    if (aCode && bCode && aCode !== bCode) return aCode - bCode;
    if (typeof a.order === "number" && typeof b.order === "number" && a.order !== b.order) return a.order - b.order;
    if (typeof a.globalOrder === "number" && typeof b.globalOrder === "number" && a.globalOrder !== b.globalOrder) {
      return a.globalOrder - b.globalOrder;
    }
    return String(a.label || "").localeCompare(String(b.label || ""), "ja");
  }

  function parseGameInfo(sheetSuffix, fallbackOpponent) {
    const text = cleanText(sheetSuffix);
    const codeMatch = text.match(/(\d{3,4})/);
    let dateCode = null;
    let dateLabel = "";
    if (codeMatch) {
      const code = codeMatch[1].padStart(4, "0");
      const month = Number(code.slice(0, 2));
      const day = Number(code.slice(2, 4));
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        dateCode = month * 100 + day;
        dateLabel = `${month}/${day}`;
      }
    }

    let opponent = fallbackOpponent || "";
    const vsMatch = text.match(/vs[_\s]*([^\s].*)$/i);
    if (vsMatch) {
      opponent = cleanOpponent(vsMatch[1]) || opponent;
    }

    return {
      dateCode,
      dateLabel,
      opponent,
      label: buildGameLabel(dateLabel, opponent),
    };
  }

  function buildGameLabel(dateLabel, opponent) {
    const parts = [];
    if (dateLabel) parts.push(dateLabel);
    if (opponent) parts.push(`vs ${opponent}`);
    return parts.join(" ") || "不明な試合";
  }

  function shortGameLabel(dateLabel, opponent) {
    if (dateLabel) return dateLabel;
    if (opponent) return `vs${opponent.slice(0, 4)}`;
    return "-";
  }

  function cleanOpponent(value) {
    let s = cleanText(value);
    if (!s) return "";
    s = s.replace(/大学.*/u, "大学");
    s = s.replace(/[A-ZＡ-Ｚ]?$/, (m) => m); // keep team suffix like A/B
    return s;
  }

  function normalizePlayerName(value) {
    let s = cleanText(value).replace(/[ \t\r\n]+/g, "").replace(/　/g, "");
    if (!s) return "";

    const kanjiAndReading = s.match(/^(.+?[一-龯々])([\u3040-\u30FFー]{2,})$/u);
    if (kanjiAndReading) s = kanjiAndReading[1];

    return s;
  }

  function extractDisplayName(value) {
    const normalized = normalizePlayerName(value);
    if (normalized) return normalized;
    return cleanText(value);
  }

  function hasMeaningfulStats(stats) {
    if (!stats) return false;
    return Object.values(stats).some((v) => isFiniteNumber(v) && v > 0);
  }

  function chooseNumber(a, b) {
    return isFiniteNumber(a) ? a : b;
  }

  function average(values) {
    const nums = values.filter(isFiniteNumber);
    if (!nums.length) return null;
    return nums.reduce((sum, n) => sum + n, 0) / nums.length;
  }

  function toNumber(value) {
    if (value == null) return null;
    if (typeof value === "number") return Number.isFinite(value) ? value : null;
    const s = String(value).trim();
    if (!s || s === "-" || /#DIV\/0!/i.test(s)) return null;
    const num = Number(s);
    return Number.isFinite(num) ? num : null;
  }

  function clampRate(value) {
    if (!isFiniteNumber(value)) return null;
    if (value < 0) return null;
    if (value > 1 && value <= 100) return value > 1.0001 ? value / 100 : value;
    if (value > 1) return null;
    return value;
  }

  function formatNumber(value, digits) {
    if (!isFiniteNumber(value)) return "-";
    return Number(value).toFixed(digits);
  }

  function formatRate(value) {
    if (!isFiniteNumber(value)) return "-";
    return `${(value * 100).toFixed(1)}%`;
  }

  function isFiniteNumber(value) {
    return typeof value === "number" && Number.isFinite(value);
  }

  function cleanText(value) {
    if (value == null) return "";
    return String(value).replace(/\s+/g, " ").trim();
  }

  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function withAlpha(hex, alpha) {
    const h = hex.replace("#", "");
    if (h.length !== 6) return hex;
    const r = parseInt(h.slice(0, 2), 16);
    const g = parseInt(h.slice(2, 4), 16);
    const b = parseInt(h.slice(4, 6), 16);
    return `rgba(${r}, ${g}, ${b}, ${alpha})`;
  }

  function readJson(key, fallback) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch {
      return fallback;
    }
  }

  function writeJson(key, value) {
    localStorage.setItem(key, JSON.stringify(value));
  }

  function cryptoRandomId() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    return `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  }

  function formatDateTime(iso) {
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return String(iso || "");
    return d.toLocaleString("ja-JP", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
    });
  }

  function createDatasetVersion(games) {
    const seed = games.map((g) => g.gameId).join("|");
    let hash = 0;
    for (let i = 0; i < seed.length; i += 1) {
      hash = (hash * 31 + seed.charCodeAt(i)) >>> 0;
    }
    return `v${hash.toString(16)}`;
  }

  init();
})();
