/*!
 * Timeføring – Advokat Geir Frøholm AS
 * © 2026 Mathias Frøholm. Alle rettigheter forbeholdt.
 * Kopiering, endring eller gjenbruk uten skriftlig tillatelse er ikke tillatt.
 */
/*
 * Innlogging med Microsoft Entra ID (MSAL.js v5, auth code + PKCE, redirect-flyt).
 * Global: window.Auth
 */
(function () {
  "use strict";

  var CFG = window.TIMER_CONFIG || {};
  var SCOPES = ["User.Read", "Files.ReadWrite", "Mail.Send"];
  var MELDING_NOKKEL = "timer.melding"; // kun systemmeldinger (ingen klientdata)
  var INAKTIV_MS = 30 * 60 * 1000;

  var pca = null;
  var konto = null;

  function konfigurert() {
    return !!(CFG.clientId && CFG.tenantId);
  }

  function settMelding(tekst) {
    try { sessionStorage.setItem(MELDING_NOKKEL, tekst); } catch (e) { /* ignorer */ }
  }

  function hentMelding() {
    try {
      var m = sessionStorage.getItem(MELDING_NOKKEL);
      sessionStorage.removeItem(MELDING_NOKKEL);
      return m;
    } catch (e) { return null; }
  }

  function tillatt(account) {
    var navn = String((account && account.username) || "").toLowerCase();
    var domene = String(CFG.tillattDomene || "").toLowerCase();
    return !!domene && navn.slice(-(domene.length + 1)) === "@" + domene;
  }

  /**
   * Initialiserer MSAL og behandler eventuelt svar fra innloggingen.
   * Returnerer kontoen hvis brukeren er innlogget med tillatt domene, ellers null.
   */
  async function init() {
    if (!konfigurert()) return null;
    if (!window.msal || !window.msal.PublicClientApplication) {
      throw new Error("Innloggingsbiblioteket kunne ikke lastes.");
    }
    pca = new window.msal.PublicClientApplication({
      auth: {
        clientId: CFG.clientId,
        authority: "https://login.microsoftonline.com/" + CFG.tenantId,
        redirectUri: CFG.redirectUri,
        postLogoutRedirectUri: CFG.redirectUri
      },
      cache: {
        cacheLocation: "sessionStorage"
      }
    });
    await pca.initialize();

    var svar = await pca.handleRedirectPromise({ navigateToLoginRequestUrl: false });
    if (svar && svar.account) {
      pca.setActiveAccount(svar.account);
    }
    var aktiv = pca.getActiveAccount();
    if (!aktiv) {
      var alle = pca.getAllAccounts();
      if (alle.length === 1) aktiv = alle[0];
    }
    if (!aktiv) return null;

    if (!tillatt(aktiv)) {
      settMelding("Kontoen " + aktiv.username + " har ikke tilgang. Logg inn med en konto på @" + CFG.tillattDomene + ".");
      await pca.logoutRedirect({ account: aktiv });
      return null;
    }
    pca.setActiveAccount(aktiv);
    konto = aktiv;
    return konto;
  }

  function login() {
    if (!pca) return Promise.reject(new Error("Innlogging er ikke konfigurert."));
    return pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
  }

  /** Henter tilgangstoken for Graph. Faller tilbake til redirect ved behov for interaksjon. */
  async function token() {
    if (!pca || !konto) throw new Error("Ikke innlogget.");
    try {
      var r = await pca.acquireTokenSilent({ scopes: SCOPES, account: konto });
      return r.accessToken;
    } catch (e) {
      if (e instanceof window.msal.InteractionRequiredAuthError) {
        await pca.acquireTokenRedirect({ scopes: SCOPES, account: konto });
        // Siden navigerer bort; returnerer et løfte som aldri løses.
        return new Promise(function () {});
      }
      throw e;
    }
  }

  function logout(melding) {
    if (melding) settMelding(melding);
    stoppInaktivitet();
    if (!pca) return Promise.resolve();
    var k = konto;
    konto = null;
    return pca.logoutRedirect({ account: k || undefined });
  }

  // ---------- Automatisk utlogging etter inaktivitet ----------
  var sisteAktivitet = Date.now();
  var inaktivTimer = null;
  var inaktivCallback = null;
  var HENDELSER = ["pointerdown", "keydown", "wheel", "touchstart", "scroll"];

  function registrerAktivitet() { sisteAktivitet = Date.now(); }

  function startInaktivitet(cb) {
    stoppInaktivitet();
    inaktivCallback = cb;
    sisteAktivitet = Date.now();
    HENDELSER.forEach(function (h) { window.addEventListener(h, registrerAktivitet, { passive: true, capture: true }); });
    inaktivTimer = window.setInterval(function () {
      if (Date.now() - sisteAktivitet >= INAKTIV_MS) {
        var c = inaktivCallback;
        stoppInaktivitet();
        if (c) c();
      }
    }, 15000);
  }

  function stoppInaktivitet() {
    if (inaktivTimer) window.clearInterval(inaktivTimer);
    inaktivTimer = null;
    HENDELSER.forEach(function (h) { window.removeEventListener(h, registrerAktivitet, { capture: true }); });
  }

  window.Auth = {
    SCOPES: SCOPES,
    INAKTIV_MS: INAKTIV_MS,
    konfigurert: konfigurert,
    init: init,
    login: login,
    logout: logout,
    token: token,
    konto: function () { return konto; },
    settMelding: settMelding,
    hentMelding: hentMelding,
    startInaktivitet: startInaktivitet,
    stoppInaktivitet: stoppInaktivitet
  };
})();
