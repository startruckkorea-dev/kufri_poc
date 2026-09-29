import { SHAREPOINT } from "./config.js";
import { initAuth, login, logout, getToken } from "./auth.js";
import { listFolder } from "./sharepoint.js";

const $ = (id) => document.getElementById(id);

function setStatus(text, isError = false) {
  $("status").textContent = text;
  $("status").classList.toggle("error", isError);
}

function formatSize(bytes) {
  if (bytes == null) return "";
  const units = ["B", "KB", "MB", "GB"];
  let i = 0;
  while (bytes >= 1024 && i < units.length - 1) {
    bytes /= 1024;
    i++;
  }
  return `${bytes.toFixed(i ? 1 : 0)} ${units[i]}`;
}

function renderItems(items) {
  const tbody = $("files").querySelector("tbody");
  tbody.replaceChildren(
    ...items.map((item) => {
      const tr = document.createElement("tr");
      const name = document.createElement("a");
      name.href = item.webUrl;
      name.target = "_blank";
      name.rel = "noopener";
      name.textContent = item.folder ? `${item.name}/` : item.name;
      const cells = [
        name,
        item.folder ? `항목 ${item.folder.childCount}개` : formatSize(item.size),
        new Date(item.lastModifiedDateTime).toLocaleString("ko-KR"),
      ];
      for (const cell of cells) {
        const td = document.createElement("td");
        td.append(cell);
        tr.append(td);
      }
      return tr;
    })
  );
  $("files").hidden = false;
}

async function showFolder() {
  setStatus(`${SHAREPOINT.folderPath} 폴더를 불러오는 중...`);
  try {
    const { scopes } = await getToken();
    $("scopes").textContent = `부여된 권한: ${scopes.join(", ")}`;
    const items = await listFolder();
    renderItems(items);
    setStatus(`${SHAREPOINT.sitePath}/${SHAREPOINT.folderPath} — 항목 ${items.length}개`);
  } catch (err) {
    console.error(err);
    setStatus(`폴더 조회 실패: ${err.errorCode ?? ""} ${err.message}`, true);
  }
}

async function main() {
  $("login").addEventListener("click", login);
  $("logout").addEventListener("click", logout);

  let account;
  try {
    account = await initAuth();
  } catch (err) {
    console.error(err);
    setStatus(`로그인 실패: ${err.errorCode ?? ""} ${err.message}`, true);
    $("login").hidden = false;
    return;
  }

  if (!account) {
    setStatus("Microsoft 계정으로 로그인하세요.");
    $("login").hidden = false;
    return;
  }

  $("user").textContent = `${account.name ?? ""} (${account.username})`;
  $("logout").hidden = false;
  await showFolder();
}

main();
