import { AUTH, GRAPH_SCOPES } from "./config.js";

// msal-browser는 index.html에서 <script>로 먼저 로드되어 전역 msal로 노출된다.
const { createStandardPublicClientApplication, InteractionRequiredAuthError } =
  window.msal;

let pca;

// 페이지 로드 시 한 번 호출. 로그인 리디렉션에서 돌아온 경우 응답을 처리하고
// 로그인된 계정(없으면 null)을 돌려준다.
export async function initAuth() {
  pca = await createStandardPublicClientApplication({
    auth: {
      clientId: AUTH.clientId,
      authority: `https://login.microsoftonline.com/${AUTH.tenantId}`,
      redirectUri: AUTH.redirectUri,
      postLogoutRedirectUri: AUTH.redirectUri,
    },
    cache: { cacheLocation: "sessionStorage" },
  });

  const result = await pca.handleRedirectPromise();
  const account = result?.account ?? pca.getAllAccounts()[0] ?? null;
  if (account) pca.setActiveAccount(account);
  return account;
}

export function login() {
  return pca.loginRedirect({ scopes: GRAPH_SCOPES });
}

export function logout() {
  return pca.logoutRedirect({ account: pca.getActiveAccount() });
}

// 캐시/리프레시 토큰으로 조용히 발급하고, 재로그인이 필요하면 리디렉션한다.
export async function getToken(scopes = GRAPH_SCOPES) {
  const request = { scopes, account: pca.getActiveAccount() };
  try {
    return await pca.acquireTokenSilent(request);
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      await pca.acquireTokenRedirect(request);
    }
    throw err;
  }
}
