"use strict";

const canvas = document.getElementById("gameCanvas");
const context = canvas.getContext("2d");
const scoreElement = document.getElementById("score");
const remainingBricksElement = document.getElementById("remainingBricks");
const remainingChancesElement = document.getElementById("remainingChances");
const statusElement = document.getElementById("status");
const messageElement = document.getElementById("message");

// 기획서에 확정된 게임 크기와 개수입니다.
const GAME = {
  width: canvas.width,
  height: canvas.height,
  scorePerBrick: 10,
  totalChances: 3,
};

const BRICK = {
  rows: 7,
  columns: 9,
  width: 100,
  height: 30,
  gap: 5,
  top: 30,
  left: 10,
  color: "#e53935",
};

const PADDLE = {
  width: 150,
  height: 15,
  speed: 9,
  bottomGap: 24,
  color: "#000000",
};

const BALL = {
  radius: 8,
  speedX: 4,
  speedY: -4,
  color: "#1565c0",
};

const STATE_TEXT = {
  waiting: "대기",
  playing: "진행 중",
  chanceEnded: "기회 종료",
  won: "승리",
  lost: "게임 종료",
};

let bricks = [];
let paddle = {};
let ball = {};
let score = 0;
let remainingBricks = BRICK.rows * BRICK.columns;
let remainingChances = GAME.totalChances;
let gameState = "waiting";
let leftPressed = false;
let rightPressed = false;
let animationId = null;

function createBricks() {
  bricks = [];

  for (let row = 0; row < BRICK.rows; row += 1) {
    for (let column = 0; column < BRICK.columns; column += 1) {
      bricks.push({
        x: BRICK.left + column * (BRICK.width + BRICK.gap),
        y: BRICK.top + row * (BRICK.height + BRICK.gap),
        width: BRICK.width,
        height: BRICK.height,
        visible: true,
      });
    }
  }
}

function placeBallOnPaddle() {
  ball.x = paddle.x + paddle.width / 2;
  ball.y = paddle.y - BALL.radius;
}

function resetRound() {
  createBricks();

  paddle = {
    x: (GAME.width - PADDLE.width) / 2,
    y: GAME.height - PADDLE.bottomGap - PADDLE.height,
    width: PADDLE.width,
    height: PADDLE.height,
  };

  ball = {
    x: 0,
    y: 0,
    dx: BALL.speedX,
    dy: BALL.speedY,
    radius: BALL.radius,
  };

  score = 0;
  remainingBricks = BRICK.rows * BRICK.columns;
  gameState = "waiting";
  leftPressed = false;
  rightPressed = false;
  placeBallOnPaddle();
  updateInformation();
  draw();
}

function resetGame() {
  remainingChances = GAME.totalChances;
  resetRound();
  showMessage("스페이스키를 눌러 게임을 시작하세요.");
}

function updateInformation() {
  scoreElement.textContent = String(score);
  remainingBricksElement.textContent = String(remainingBricks);
  remainingChancesElement.textContent = String(remainingChances);
  statusElement.textContent = STATE_TEXT[gameState];
}

function showMessage(message) {
  messageElement.textContent = message;
  messageElement.classList.remove("hidden");
}

function hideMessage() {
  messageElement.classList.add("hidden");
}

function drawBricks() {
  context.fillStyle = BRICK.color;

  bricks.forEach((brick) => {
    if (brick.visible) {
      context.fillRect(brick.x, brick.y, brick.width, brick.height);
    }
  });
}

function drawPaddle() {
  context.fillStyle = PADDLE.color;
  context.fillRect(paddle.x, paddle.y, paddle.width, paddle.height);
}

function drawBall() {
  context.beginPath();
  context.arc(ball.x, ball.y, ball.radius, 0, Math.PI * 2);
  context.fillStyle = BALL.color;
  context.fill();
  context.closePath();
}

function draw() {
  context.clearRect(0, 0, GAME.width, GAME.height);
  drawBricks();
  drawPaddle();
  drawBall();
}

function movePaddle() {
  if (leftPressed) {
    paddle.x = Math.max(0, paddle.x - PADDLE.speed);
  }

  if (rightPressed) {
    paddle.x = Math.min(GAME.width - paddle.width, paddle.x + PADDLE.speed);
  }
}

function circleHitsRectangle(rectangle) {
  const closestX = Math.max(rectangle.x, Math.min(ball.x, rectangle.x + rectangle.width));
  const closestY = Math.max(rectangle.y, Math.min(ball.y, rectangle.y + rectangle.height));
  const distanceX = ball.x - closestX;
  const distanceY = ball.y - closestY;

  return distanceX * distanceX + distanceY * distanceY <= ball.radius * ball.radius;
}

function handleWallCollision() {
  if (ball.x - ball.radius <= 0 && ball.dx < 0) {
    ball.x = ball.radius;
    ball.dx *= -1;
  } else if (ball.x + ball.radius >= GAME.width && ball.dx > 0) {
    ball.x = GAME.width - ball.radius;
    ball.dx *= -1;
  }

  if (ball.y - ball.radius <= 0 && ball.dy < 0) {
    ball.y = ball.radius;
    ball.dy *= -1;
  }
}

function handlePaddleCollision() {
  if (ball.dy > 0 && circleHitsRectangle(paddle)) {
    ball.y = paddle.y - ball.radius;

    // 바의 어느 지점에 맞았는지에 따라 가로 방향을 바꿉니다.
    const hitPosition = (ball.x - (paddle.x + paddle.width / 2)) / (paddle.width / 2);
    ball.dx = hitPosition * 5;
    ball.dy = -Math.abs(ball.dy);

    if (Math.abs(ball.dx) < 1.5) {
      ball.dx = ball.dx < 0 ? -1.5 : 1.5;
    }
  }
}

function handleBrickCollision() {
  for (const brick of bricks) {
    if (!brick.visible || !circleHitsRectangle(brick)) {
      continue;
    }

    const previousX = ball.x - ball.dx;
    const cameFromSide =
      previousX + ball.radius <= brick.x || previousX - ball.radius >= brick.x + brick.width;

    brick.visible = false;
    score += GAME.scorePerBrick;
    remainingBricks -= 1;

    if (cameFromSide) {
      ball.dx *= -1;
    } else {
      ball.dy *= -1;
    }

    updateInformation();

    if (remainingBricks === 0) {
      finishGame("won");
    }

    break;
  }
}

function finishGame(result) {
  gameState = result;
  updateInformation();

  if (result === "won") {
    showMessage(`승리했습니다! 최종 점수: ${score}점\n스페이스키를 눌러 다시 시작하세요.`);
  } else {
    showMessage(`게임이 종료되었습니다. 최종 점수: ${score}점\n스페이스키를 눌러 다시 시작하세요.`);
  }
}

function finishChance() {
  remainingChances -= 1;

  if (remainingChances === 0) {
    finishGame("lost");
    return;
  }

  gameState = "chanceEnded";
  updateInformation();
  showMessage(
    `이번 게임이 끝났습니다. 최종 점수: ${score}점\n남은 기회: ${remainingChances}회\n스페이스키를 눌러 다음 게임을 시작하세요.`,
  );
}

function update() {
  movePaddle();
  ball.x += ball.dx;
  ball.y += ball.dy;

  handleWallCollision();
  handlePaddleCollision();
  handleBrickCollision();

  if (gameState === "playing" && ball.y - ball.radius > GAME.height) {
    finishChance();
  }
}

function gameLoop() {
  if (gameState !== "playing") {
    animationId = null;
    draw();
    return;
  }

  update();
  draw();

  if (gameState === "playing") {
    animationId = requestAnimationFrame(gameLoop);
  } else {
    animationId = null;
  }
}

function startGame() {
  if (gameState === "won" || gameState === "lost") {
    resetGame();
  } else if (gameState === "chanceEnded") {
    resetRound();
  }

  if (gameState !== "waiting") {
    return;
  }

  gameState = "playing";
  updateInformation();
  hideMessage();

  if (animationId === null) {
    animationId = requestAnimationFrame(gameLoop);
  }
}

function handleKeyDown(event) {
  if (["ArrowLeft", "ArrowRight", "Space"].includes(event.code)) {
    event.preventDefault();
  }

  if (event.code === "ArrowLeft") {
    leftPressed = true;
  } else if (event.code === "ArrowRight") {
    rightPressed = true;
  } else if (event.code === "Space" && !event.repeat) {
    startGame();
  }
}

function handleKeyUp(event) {
  if (event.code === "ArrowLeft") {
    leftPressed = false;
  } else if (event.code === "ArrowRight") {
    rightPressed = false;
  }
}

window.addEventListener("keydown", handleKeyDown);
window.addEventListener("keyup", handleKeyUp);
window.addEventListener("blur", () => {
  leftPressed = false;
  rightPressed = false;
});

resetGame();
