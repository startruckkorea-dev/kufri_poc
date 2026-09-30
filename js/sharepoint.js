import { SHAREPOINT } from "./config.js";
import { getToken } from "./auth.js";

const GRAPH = "https://graph.microsoft.com/v1.0";

export class GraphError extends Error {
  constructor(status, code, message) {
    super(`${status} ${code}: ${message}`);
    this.status = status;
    this.code = code;
  }
}

async function graphGet(url) {
  const { accessToken } = await getToken();
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new GraphError(
      res.status,
      body.error?.code ?? res.statusText,
      body.error?.message ?? ""
    );
  }
  return res.json();
}

// Graph는 사이트 경로와 항목 경로를 한 URL에 같이 쓰는 것을 허용하지 않아
// 사이트의 기본 문서 라이브러리(drive) ID를 먼저 조회해 둔다.
let driveIdPromise;
function getDriveId() {
  driveIdPromise ??= graphGet(
    `${GRAPH}/sites/${SHAREPOINT.hostname}:${SHAREPOINT.sitePath}:/drive?$select=id`
  ).then((drive) => drive.id);
  return driveIdPromise;
}

async function folderUrl(folderPath) {
  const driveId = await getDriveId();
  const path = folderPath.split("/").map(encodeURIComponent).join("/");
  return path
    ? `${GRAPH}/drives/${driveId}/root:/${path}:`
    : `${GRAPH}/drives/${driveId}/root`;
}

// 폴더 바로 아래 항목(파일/하위 폴더)을 모두 가져온다.
export async function listFolder(folderPath = SHAREPOINT.folderPath) {
  const items = [];
  let url =
    `${await folderUrl(folderPath)}/children` +
    "?$select=id,name,size,lastModifiedDateTime,webUrl,file,folder&$top=200";
  while (url) {
    const page = await graphGet(url);
    items.push(...page.value);
    url = page["@odata.nextLink"];
  }
  return items;
}
