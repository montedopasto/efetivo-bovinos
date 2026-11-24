(function(){
  'use strict';
  const tenant='ee417351-ea90-41e0-9147-5ea6ab38ea49';
  const clientId='c067304f-7176-4e3f-a5fb-202bbc3a2ec7';
  const instance=new msal.PublicClientApplication({
    auth:{clientId,authority:'https://login.microsoftonline.com/'+tenant,
      redirectUri:window.location.origin+'/efetivo-bovinos/index.html'},
    cache:{cacheLocation:'sessionStorage',storeAuthStateInCookie:false}
  });
  const scopes=['openid','profile'];
  function allowed(account){
    return account && account.tenantId===tenant && /^[^@\s]+@montedopasto\.pt$/i.test(account.username || '');
  }
  function account(){
    const active=instance.getActiveAccount();
    return allowed(active) ? active : instance.getAllAccounts().find(allowed) || null;
  }
  async function token(interactive){
    let current=account(), response;
    if(current){
      try{response=await instance.acquireTokenSilent({scopes,account:current});}
      catch(error){if(!interactive)throw new Error('Inicia sessão com a tua conta Microsoft da Monte do Pasto.');}
    }
    if(!response){
      if(!interactive)throw new Error('Inicia sessão com a tua conta Microsoft da Monte do Pasto.');
      response=await instance.loginPopup({scopes,prompt:'select_account'});
    }
    if(!allowed(response.account) || !response.idToken)throw new Error('Utiliza uma conta @montedopasto.pt da organização.');
    instance.setActiveAccount(response.account);
    return response.idToken;
  }
  async function signOut(){
    if(window.Core)Core.clearState();
    const current=account();
    await instance.logoutPopup({account:current,postLogoutRedirectUri:window.location.origin+'/efetivo-bovinos/index.html',mainWindowRedirectUri:window.location.href});
  }
  window.BoviMicrosoft={token,account,signOut};
})();
