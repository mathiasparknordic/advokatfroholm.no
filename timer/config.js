/*
 * Konfigurasjon for timeføringen.
 * clientId/tenantId er ikke hemmeligheter, men står tomme til app-registreringen i Entra ID er gjort.
 * Når clientId er tom, tilbys kun demo-modus på innloggingssiden.
 */
window.TIMER_CONFIG = {
  clientId: "7a247c96-ad6a-4285-906c-8661a292a282",
  tenantId: "a50a33a8-d98b-425e-a2e7-2f84e0b6b6de",
  redirectUri: "https://advokatfroholm.no/timer/",
  tillattDomene: "advokatfroholm.no"
};
