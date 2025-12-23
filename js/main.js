let idleSeconds = 20 * 60;
let idleInterval;
function updateIdleDisplay() {
  let min = String(Math.floor(idleSeconds / 60)).padStart(2, "0");
  let sec = String(idleSeconds % 60).padStart(2, "0");
  $("#idleTime").text(`${min}:${sec}`);
}

function resetIdleTimer() {
  idleSeconds = 20 * 60;
  updateIdleDisplay();
}

function startIdleCountdown() {
  idleInterval = setInterval(() => {
    idleSeconds--;
    updateIdleDisplay();
    if (idleSeconds <= 0) {
      clearInterval(idleInterval);
      window.location.href = "/index.php";
    }
  }, 1000);
}

// 偵測使用者操作（點擊、按鍵）
["keydown", "click"].forEach((evt) => {
  document.addEventListener(evt, resetIdleTimer);
});

// 取得Cookie
function getCookie(name) {
  let match = document.cookie.match(new RegExp("(^| )" + name + "=([^;]+)"));
  return match ? decodeURIComponent(match[2]) : null;
}

// 取得網址參數的函式
function getQueryParam(key) {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(key);
}


// 字體大小控制
function cssControl() {
  let currentSize = localStorage.getItem("fontSize")
    ? parseInt(localStorage.getItem("fontSize"))
    : 16;
  $("body").css("font-size", currentSize + "px");

  $("#font-smaller").on("click", () => {
    if (currentSize > 12) currentSize--;
    $("body").css("font-size", currentSize + "px");
    localStorage.setItem("fontSize", currentSize);
  });

  $("#font-larger").on("click", () => {
    if (currentSize < 22) currentSize++;
    $("body").css("font-size", currentSize + "px");
    localStorage.setItem("fontSize", currentSize);
  });

  $("#font-default").on("click", () => {
    currentSize = 16;
    $("body").css("font-size", currentSize + "px");
    localStorage.setItem("fontSize", currentSize);
  });
}

// 圖片放大鏡
function createMagnifierFollowMouse(
  imgSelector,
  magnifierSelector,
  zoom = 2,
  offsetX = 0,
  offsetY = 0
) {
  const img = document.querySelector(imgSelector);
  const magnifier = document.querySelector(magnifierSelector);
  if (!img || !magnifier) return;

  img.addEventListener("mousemove", (e) => {
    magnifier.style.display = "block";

    const rect = img.getBoundingClientRect();
    const x = e.clientX - rect.left; // 滑鼠在圖片內的 X
    const y = e.clientY - rect.top; // 滑鼠在圖片內的 Y

    // 放大鏡背景
    magnifier.style.backgroundImage = `url(${img.src})`;
    magnifier.style.backgroundSize = `${img.width * zoom}px ${
      img.height * zoom
    }px`;
    magnifier.style.backgroundPosition = `
      ${-(x * zoom - magnifier.offsetWidth / 2)}px 
      ${-(y * zoom - magnifier.offsetHeight / 2)}px
    `;

    // 放大鏡位置：滑鼠右上方
    let left = e.clientX + offsetX;
    let top = e.clientY - offsetY - magnifier.offsetHeight;

    // 避免超出視窗右邊界
    if (left + magnifier.offsetWidth > window.innerWidth) {
      left = e.clientX - offsetX - magnifier.offsetWidth;
    }
    // 避免超出視窗頂部
    if (top < 0) {
      top = e.clientY + offsetY;
    }

    magnifier.style.left = `${left}px`;
    magnifier.style.top = `${top}px`;
  });

  img.addEventListener("mouseleave", () => {
    magnifier.style.display = "none";
  });
}

// Custom Toast
/**
 * 使用範例：
 * showToast("送件完成", "success");
 * showToast("資料有誤", "error");
 * showToast("提示訊息", "info"); *
 */
function showToast(message, type = "info") {
  const toast = document.createElement("div");
  toast.className = `custom-toast toast-${type}`;
  toast.innerHTML = `
    <div class="d-flex align-items-center left">
      <img src="./assets/svg/${
        type === "success"
          ? "SuccessOutlined"
          : type === "error"
          ? "ErrorOutlined"
          : "InfoOutlined"
      }.svg" width="20" height="20" alt=">" class="mr-1" />     
      ${message}
    </div>

    <div class="right">
      <img class="material-symbols-rounded close-btn" src="./assets/svg/CloseAffordance.svg" width="20" height="20" alt=">" />
    </div>
  `;

  document.getElementById("toast-container").appendChild(toast);

  // 滑入效果
  setTimeout(() => toast.classList.add("show"), 10);

  // 點 x 關閉
  toast.querySelector(".close-btn").addEventListener("click", () => {
    hideToast(toast);
  });

  // 自動 4 秒後消失
  setTimeout(() => hideToast(toast), 4000);
}

function hideToast(toast) {
  toast.classList.remove("show");
  setTimeout(() => toast.remove(), 300);
}
