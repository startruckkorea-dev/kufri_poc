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

function folderUrl(folderPath) {
  const path = folderPath.split("/").map(encodeURIComponent).join("/");
  const site = `${GRAPH}/sites/${SHAREPOINT.hostname}:${SHAREPOINT.sitePath}:`;
  return path ? `${site}/drive/root:/${path}:` : `${site}/drive/root`;
}

// 폴더 바로 아래 항목(파일/하위 폴더)을 모두 가져온다.
export async function listFolder(folderPath = SHAREPOINT.folderPath) {
  const items = [];
  let url =
    `${folderUrl(folderPath)}/children` +
    "?$select=id,name,size,lastModifiedDateTime,webUrl,file,folder&$top=200";
  while (url) {
    const page = await graphGet(url);
    items.push(...page.value);
    url = page["@odata.nextLink"];
  }
  return items;
}
