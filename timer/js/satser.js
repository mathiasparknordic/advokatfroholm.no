/*!
 * Timeføring – Advokat Geir Frøholm AS
 * © 2026 Mathias Frøholm. Alle rettigheter forbeholdt.
 * Kopiering, endring eller gjenbruk uten skriftlig tillatelse er ikke tillatt.
 */
/*
 * Statens satser for kilometergodtgjørelse (egen bil).
 * Kilde: Særavtale om dekning av utgifter til reise og kost innenlands (staten), § 6 nr. 2.
 * Ny sats legges inn som ny rad med «fra»-dato. Eldre rader beholdes, slik at
 * kjøring føres med satsen som gjaldt på kjøredatoen.
 */
(function () {
  "use strict";

  var KM = [
    {
      fra: "2026-01-01",
      til: "2027-12-31",
      sats: 5.30,
      kilde: "Særavtale om reiser innenlands for statens regning, § 6 nr. 2 (gjelder 1.1.2026–31.12.2027)",
      url: "https://www.regjeringen.no/contentassets/2baeb58f246849a1ba137bbeb7f20732/no/pdfs/saeravtale-reise-i-norge.pdf"
    }
  ];

  /** Satsen som gjelder på en dato (ÅÅÅÅ-MM-DD). */
  function kmSats(dato) {
    var d = String(dato || "").slice(0, 10);
    var liste = KM.slice().sort(function (a, b) { return a.fra < b.fra ? -1 : 1; });
    var treff = null;
    liste.forEach(function (r) { if (r.fra <= d) treff = r; });
    if (!treff) return null; // før første kjente sats – må føres manuelt
    return {
      sats: treff.sats,
      kilde: treff.kilde,
      url: treff.url,
      fra: treff.fra,
      til: treff.til,
      // Etter avtaleperioden er satsen ikke bekreftet – vis varsel.
      ubekreftet: !!(treff.til && d > treff.til)
    };
  }

  window.Satser = { KM: KM, kmSats: kmSats };
})();
