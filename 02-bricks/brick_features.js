// 업무 지원 아이템과 NPC 배달 퀘스트를 기존 미로 게임에 덧붙입니다.
(() => {
  const effectText = document.querySelector("#effectText");
  const questText = document.querySelector("#questText");
  const itemTypes = ["shield", "freeze", "magnet"];
  const itemLabels = { shield: "보", freeze: "정", magnet: "흡" };
  const itemNames = { shield: "긴급 보호막", freeze: "정지 명령", magnet: "회수 자석" };
  const npcNames = ["정비 요원", "기록 요원", "통신 요원", "보급 요원", "지도 요원", "안전 요원", "관제 요원"];
  const packageNames = ["공구함", "점검 문서", "통신 부품", "에너지 상자", "구역 지도", "안전 장치", "관제 기록"];

  let item = null;
  let npc = null;
  let parcel = null;
  let questState = "waiting";
  let shield = false;
  let freezeUntil = 0;
  let magnetUntil = 0;

  const basePrepare = prepare;
  const baseDraw = draw;
  const baseMovePlayer = movePlayer;
  const baseMoveEnemies = moveEnemies;
  const baseCollision = collision;
  const baseStatus = status;

  function pickCell(list, position, blocked = []) {
    const available = list.filter((cell) => !blocked.some((other) => other && other.r === cell.r && other.c === cell.c));
    return available[Math.min(position, available.length - 1)] || { r: 1, c: 1 };
  }

  function prepareFeatures() {
    const list = cells().filter((cell) => cell.r + cell.c > 8);
    const type = itemTypes[(stage - 1) % itemTypes.length];
    npc = pickCell(list, Math.floor(list.length * 0.25));
    parcel = pickCell(list, Math.floor(list.length * 0.68), [npc]);
    item = { ...pickCell(list, Math.floor(list.length * 0.45), [npc, parcel]), type, taken: false };
    questState = "waiting";
    shield = false;
    freezeUntil = 0;
    magnetUntil = 0;
    pieces.delete(key(npc.r, npc.c));
    pieces.delete(key(parcel.r, parcel.c));
    pieces.delete(key(item.r, item.c));
  }

  function activate(type) {
    if (type === "shield") shield = true;
    if (type === "freeze") freezeUntil = performance.now() + 4000;
    if (type === "magnet") magnetUntil = performance.now() + 6000;
    tell(`${itemNames[type]} 효과가 발동했습니다.`);
    refreshSupportText();
  }

  function remainingSeconds(until) {
    return Math.max(0, Math.ceil((until - performance.now()) / 1000));
  }

  function refreshSupportText() {
    if (shield) effectText.textContent = "보호막 1회";
    else if (performance.now() < freezeUntil) effectText.textContent = `정지 명령 ${remainingSeconds(freezeUntil)}초`;
    else if (performance.now() < magnetUntil) effectText.textContent = `회수 자석 ${remainingSeconds(magnetUntil)}초`;
    else effectText.textContent = "없음";

    const name = npcNames[stage - 1];
    const packageName = packageNames[stage - 1];
    if (questState === "waiting") questText.textContent = `${name}을 만나세요`;
    if (questState === "active") questText.textContent = `${packageName} 찾기`;
    if (questState === "carrying") questText.textContent = `${packageName} 운반 중`;
    if (questState === "done") questText.textContent = "배달 완료 +250점";
  }

  function drawLabel(text, row, column, fill, stroke = "#173b73") {
    const x = column * SIZE + SIZE / 2;
    const y = row * SIZE + SIZE / 2;
    ctx.fillStyle = fill;
    ctx.strokeStyle = stroke;
    ctx.lineWidth = 3;
    ctx.fillRect(x - 13, y - 13, 26, 26);
    ctx.strokeRect(x - 13, y - 13, 26, 26);
    ctx.fillStyle = "#111";
    ctx.font = "bold 15px 'Malgun Gothic'";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(text, x, y + 1);
  }

  function drawFeatures() {
    if (npc) {
      drawLabel("요", npc.r, npc.c, questState === "done" ? "#b6f34a" : "#24a9c7");
      ctx.fillStyle = "#173b73";
      ctx.fillRect(npc.c * SIZE + 15, npc.r * SIZE + 2, 6, 5);
    }
    if (parcel && questState === "active") drawLabel("물", parcel.r, parcel.c, "#f0c84b");
    if (item && !item.taken) drawLabel(itemLabels[item.type], item.r, item.c, "#b6f34a");
    if (questState === "carrying") {
      ctx.fillStyle = "#f0c84b";
      ctx.fillRect(player.c * SIZE + 12, player.r * SIZE + 1, 12, 7);
    }
  }

  function sameCell(first, second) {
    return first && second && first.r === second.r && first.c === second.c;
  }

  function collectMagnetPieces() {
    if (performance.now() >= magnetUntil) return;
    let collected = 0;
    [...pieces].forEach((pieceKey) => {
      const [row, column] = pieceKey.split(",").map(Number);
      if (Math.abs(row - player.r) + Math.abs(column - player.c) <= 1) {
        pieces.delete(pieceKey);
        collected += 1;
      }
    });
    if (collected > 0) {
      score += collected * 10;
      baseStatus();
      if (pieces.size === 0) finishStage();
    }
  }

  function checkFeatures() {
    if (item && !item.taken && sameCell(player, item)) {
      item.taken = true;
      activate(item.type);
    }
    if (sameCell(player, npc)) {
      if (questState === "waiting") {
        questState = "active";
        tell(`${npcNames[stage - 1]}: ${packageNames[stage - 1]}을 가져다주세요!`);
      } else if (questState === "carrying") {
        questState = "done";
        score += 250;
        activate(itemTypes[(stage - 1) % itemTypes.length]);
        tell("배달 완료! 업무 점수 250점을 받았습니다.");
        baseStatus();
      }
    }
    if (questState === "active" && sameCell(player, parcel)) {
      questState = "carrying";
      tell(`${packageNames[stage - 1]}을 운반합니다. NPC에게 돌아가세요.`);
    }
    collectMagnetPieces();
    refreshSupportText();
  }

  prepare = function enhancedPrepare() {
    basePrepare();
    prepareFeatures();
    baseStatus();
    baseDraw();
    drawFeatures();
    refreshSupportText();
  };

  draw = function enhancedDraw() {
    baseDraw();
    drawFeatures();
  };

  status = function enhancedStatus() {
    baseStatus();
    refreshSupportText();
  };

  movePlayer = function enhancedMovePlayer(time) {
    const before = `${player.r},${player.c}`;
    baseMovePlayer(time);
    if (state === "playing" && before !== `${player.r},${player.c}`) checkFeatures();
  };

  moveEnemies = function enhancedMoveEnemies(time) {
    if (time < freezeUntil) return;
    baseMoveEnemies(time);
  };

  collision = function enhancedCollision(time) {
    const hit = time >= protectedUntil && enemies.some((enemy) => sameCell(enemy, player));
    if (hit && shield) {
      shield = false;
      protectedUntil = time + 1500;
      reset();
      tell("긴급 보호막이 충돌을 막았습니다.");
      refreshSupportText();
      return;
    }
    baseCollision(time);
  };

  window.setInterval(refreshSupportText, 250);
  refreshSupportText();
})();
