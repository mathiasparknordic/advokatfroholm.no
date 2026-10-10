/*
 * Konfigurasjon for timeføringen.
 * clientId/tenantId er ikke hemmeligheter, men står tomme til app-registreringen i Entra ID er gjort.
 * Når clientId er tom, tilbys kun demo-modus på innloggingssiden.
 */
window.TIMER_CONFIG = {
  clientId: "",
  tenantId: "",
  redirectUri: "https://advokatfroholm.no/timer/",
  tillattDomene: "advokatfroholm.no"
};
