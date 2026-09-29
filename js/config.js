// Entra ID 앱 등록 정보. SPA(public client)라 client secret 없이 PKCE로 인증한다.
export const AUTH = {
  clientId: "9b247088-5afb-4622-9c5e-b5f27142761d",
  tenantId: "19cab1f5-21f4-44df-8ac6-96d6ca595203",
  // 앱 등록의 SPA 리디렉션 URI와 정확히 일치해야 한다.
  redirectUri: `${window.location.origin}/`,
};

// .default = 앱 등록에 이미 부여된 위임 권한만 요청 (추가 동의 팝업 없음)
export const GRAPH_SCOPES = ["https://graph.microsoft.com/.default"];

export const SHAREPOINT = {
  hostname: "startruckkorea.sharepoint.com",
  sitePath: "/sites/STK-Kufri",
  // 기본 문서 라이브러리(Shared Documents) 기준 경로
  folderPath: "STK-Kufri_Data/write",
};
