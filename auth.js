// =============================
// AUTH MICROSOFT (MSAL)
// =============================

const msalConfig = {
  auth: {
    clientId: "c067304f-7176-4e3f-a5fb-202bbc3a2ec7",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin + "/efetivo-bovinos/index.html"
  },
  cache: {
    cacheLocation: "localStorage"
  }
};

const loginRequest = {
  scopes: ["User.Read", "Sites.ReadWrite.All"]
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// =============================
// LOGIN
// =============================

async function loginMicrosoft(){
  try {
    const loginResponse = await msalInstance.loginPopup(loginRequest);
    msalInstance.setActiveAccount(loginResponse.account);
    console.log("✅ Login OK:", loginResponse.account.username);
    return loginResponse.account;
  } catch (err) {
    console.error("❌ Erro login:", err);
  }
}

// =============================
// TOKEN
// =============================

async function getAccessToken(){
  const account = msalInstance.getActiveAccount();

  if(!account){
    await loginMicrosoft();
  }

  try {
    const response = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account: msalInstance.getActiveAccount()
    });

    return response.accessToken;

  } catch (err) {
    console.warn("⚠️ Silent falhou, a pedir popup...");

    const response = await msalInstance.acquireTokenPopup(loginRequest);
    return response.accessToken;
  }
}

// =============================
// EXPORT GLOBAL
// =============================

window.Auth = {
  loginMicrosoft,
  getAccessToken
};
