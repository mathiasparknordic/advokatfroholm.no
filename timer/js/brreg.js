/*!
 * Timeføring – Advokat Geir Frøholm AS
 * © 2026 Mathias Frøholm. Alle rettigheter forbeholdt.
 * Kopiering, endring eller gjenbruk uten skriftlig tillatelse er ikke tillatt.
 */
/*
 * Oppslag i Enhetsregisteret (Brønnøysundregistrene). Åpent API, ingen nøkkel.
 * Kun org.nr. eller firmanavn sendes. Global: window.Brreg
 */
(function () {
  "use strict";

  var BASE = "https://data.brreg.no/enhetsregisteret/api";
  var TIDSAVBRUDD_MS = 8000;

  function BrregFeil(melding, kode) {
    var e = new Error(melding);
    e.name = "BrregFeil";
    e.kode = kode || "";
    return e;
  }

  function vask(orgnr) { return String(orgnr || "").replace(/\D/g, ""); }

  /** Validerer org.nr. med modulus 11. */
  function gyldigOrgnr(orgnr) {
    var s = vask(orgnr);
    if (!/^\d{9}$/.test(s)) return false;
    var vekter = [3, 2, 7, 6, 5, 4, 3, 2];
    var sum = 0;
    for (var i = 0; i < 8; i++) sum += parseInt(s[i], 10) * vekter[i];
    var rest = sum % 11;
    var kontroll = rest === 0 ? 0 : 11 - rest;
    if (kontroll === 10) return false;
    return kontroll === parseInt(s[8], 10);
  }

  function formaterOrgnr(orgnr) {
    var s = vask(orgnr);
    return s.length === 9 ? s.slice(0, 3) + " " + s.slice(3, 6) + " " + s.slice(6) : String(orgnr || "");
  }

  async function hent(url) {
    var ctrl = typeof AbortController !== "undefined" ? new AbortController() : null;
    var t = ctrl ? setTimeout(function () { ctrl.abort(); }, TIDSAVBRUDD_MS) : null;
    try {
      var res = await fetch(url, {
        headers: { Accept: "application/json" },
        signal: ctrl ? ctrl.signal : undefined,
        credentials: "omit",
        referrerPolicy: "no-referrer"
      });
      return res;
    } catch (e) {
      throw BrregFeil("Fikk ikke kontakt med Brønnøysundregistrene akkurat nå. Fyll inn manuelt, eller prøv igjen senere.", "nettverk");
    } finally {
      if (t) clearTimeout(t);
    }
  }

  function adresseTekst(a) {
    if (!a) return "";
    var linjer = (a.adresse || []).filter(Boolean).join(", ");
    var sted = String(a.poststed || "").toLowerCase().replace(/(^|[\s-])(\S)/g, function (m, f, b) { return f + b.toUpperCase(); });
    var post = [a.postnummer, sted].filter(Boolean).join(" ");
    var deler = [linjer, post].filter(Boolean);
    if (a.landkode && a.landkode !== "NO" && a.land) deler.push(a.land);
    return deler.join(", ");
  }

  function normaliser(e, kilde) {
    var fa = e.forretningsadresse || e.beliggenhetsadresse || null;
    var pa = e.postadresse || null;
    var advarsler = [];
    if (e.konkurs) advarsler.push("Enheten er registrert konkurs.");
    if (e.underAvvikling) advarsler.push("Enheten er under avvikling.");
    if (e.underTvangsavviklingEllerTvangsopplosning) advarsler.push("Enheten er under tvangsavvikling eller tvangsoppløsning.");
    if (e.slettedato) advarsler.push("Enheten ble slettet " + e.slettedato + ".");
    return {
      orgnr: e.organisasjonsnummer || "",
      navn: e.navn || "",
      orgform: (e.organisasjonsform && e.organisasjonsform.kode) || "",
      orgformNavn: (e.organisasjonsform && e.organisasjonsform.beskrivelse) || "",
      forretningsadresse: adresseTekst(fa),
      postadresse: adresseTekst(pa),
      fakturaadresse: adresseTekst(pa) || adresseTekst(fa),
      konkurs: !!e.konkurs,
      underAvvikling: !!e.underAvvikling,
      slettet: !!e.slettedato,
      kilde: kilde,
      advarsler: advarsler
    };
  }

  /** Slår opp et org.nr. Prøver enheter, deretter underenheter. Returnerer null om ikke funnet. */
  async function oppslag(orgnr) {
    var s = vask(orgnr);
    if (!gyldigOrgnr(s)) throw BrregFeil("Org.nr. " + (orgnr || "") + " er ikke gyldig (9 siffer med riktig kontrollsiffer).", "ugyldig");
    var typer = [["enheter", "enhet"], ["underenheter", "underenhet"]];
    for (var i = 0; i < typer.length; i++) {
      var res = await hent(BASE + "/" + typer[i][0] + "/" + s);
      if (res.status === 404) continue;
      if (res.status === 410) {
        return {
          orgnr: s, navn: "", orgform: "", orgformNavn: "", forretningsadresse: "", postadresse: "",
          fakturaadresse: "", konkurs: false, underAvvikling: false, slettet: true, kilde: typer[i][1],
          advarsler: ["Enheten er slettet fra Enhetsregisteret."]
        };
      }
      if (!res.ok) throw BrregFeil("Brønnøysundregistrene svarte med feil (" + res.status + "). Fyll inn manuelt, eller prøv igjen senere.", "http");
      var j;
      try { j = await res.json(); } catch (e) { throw BrregFeil("Uventet svar fra Brønnøysundregistrene.", "format"); }
      return normaliser(j, typer[i][1]);
    }
    return null;
  }

  /** Navnesøk. Returnerer inntil `antall` treff. */
  async function sok(navn, antall) {
    var q = String(navn || "").trim();
    if (q.length < 2) return [];
    var res = await hent(BASE + "/enheter?navn=" + encodeURIComponent(q) + "&size=" + (antall || 5));
    if (!res.ok) throw BrregFeil("Søket i Brønnøysundregistrene feilet (" + res.status + "). Prøv igjen senere.", "http");
    var j;
    try { j = await res.json(); } catch (e) { throw BrregFeil("Uventet svar fra Brønnøysundregistrene.", "format"); }
    var liste = (j && j._embedded && j._embedded.enheter) || [];
    return liste.map(function (e) { return normaliser(e, "enhet"); });
  }

  window.Brreg = {
    gyldigOrgnr: gyldigOrgnr,
    formaterOrgnr: formaterOrgnr,
    vask: vask,
    oppslag: oppslag,
    sok: sok
  };
})();
