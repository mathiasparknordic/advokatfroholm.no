/*!
 * Timeføring – Advokat Geir Frøholm AS
 * © 2026 Mathias Frøholm. Alle rettigheter forbeholdt.
 * Kopiering, endring eller gjenbruk uten skriftlig tillatelse er ikke tillatt.
 */
/*
 * Timeføring – brukergrensesnitt.
 * Avhenger av: TIMER_CONFIG, Auth, Graph, Store, Brreg, Oppdrag (parser.js), Faktura (excel.js).
 * Ingen inline-skript, ingen alert/confirm/prompt.
 */
(function () {
  "use strict";

  var CFG = window.TIMER_CONFIG || {};
  var rot = document.getElementById("app");

  // =====================================================================
  // Hjelpere: DOM
  // =====================================================================
  function el(tag, attrs) {
    var n = document.createElement(tag);
    if (attrs) {
      Object.keys(attrs).forEach(function (k) {
        var v = attrs[k];
        if (v === null || v === undefined || v === false) return;
        if (k === "class") n.className = v;
        else if (k === "text") n.textContent = v;
        else if (k === "html") n.innerHTML = v; // kun statisk markup (ikoner)
        else if (k.slice(0, 2) === "on" && typeof v === "function") n.addEventListener(k.slice(2), v);
        else if (k === "value" || k === "checked" || k === "disabled" || k === "selected") n[k] = v;
        else if (k === "dataset") Object.keys(v).forEach(function (d) { n.dataset[d] = v[d]; });
        else n.setAttribute(k, v === true ? "" : v);
      });
    }
    for (var i = 2; i < arguments.length; i++) leggTil(n, arguments[i]);
    return n;
  }
  function leggTil(n, c) {
    if (c === null || c === undefined || c === false) return;
    if (Array.isArray(c)) { c.forEach(function (x) { leggTil(n, x); }); return; }
    n.appendChild(c.nodeType ? c : document.createTextNode(String(c)));
  }
  function tom(n) { while (n.firstChild) n.removeChild(n.firstChild); return n; }
  function erstatt(n) { tom(n); for (var i = 1; i < arguments.length; i++) leggTil(n, arguments[i]); return n; }

  var unik = 0;
  function uid(p) { unik++; return (p || "f") + "-" + unik; }

  // =====================================================================
  // Ikoner (statisk SVG)
  // =====================================================================
  var IKON = {
    oversikt: '<rect x="4" y="4" width="7" height="7"/><rect x="13" y="4" width="7" height="4"/><rect x="13" y="10" width="7" height="10"/><rect x="4" y="13" width="7" height="7"/>',
    timer: '<circle cx="12" cy="12" r="8"/><path d="M12 7.5V12l3 2"/>',
    saker: '<path d="M3 6h6l2 2h10v11H3z"/>',
    oppdrag: '<path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4"/><path d="M12 11v6M9 14h6"/>',
    faktura: '<rect x="4" y="4" width="16" height="16"/><path d="M4 10h16M10 10v10"/>',
    innst: '<path d="M4 7h9M17 7h3M4 17h3M11 17h9"/><rect x="13" y="5" width="4" height="4"/><rect x="7" y="15" width="4" height="4"/>',
    pluss: '<path d="M12 5v14M5 12h14"/>',
    mer: '<path d="M5 12h.01M12 12h.01M19 12h.01" stroke-width="3"/>',
    ut: '<path d="M14 4h5v16h-5"/><path d="M10 8l-4 4 4 4M6 12h10"/>',
    las: '<rect x="5" y="11" width="14" height="9"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/>',
    last: '<path d="M12 4v11M7 10l5 5 5-5M5 20h14"/>',
    send: '<path d="M4 12l16-8-6 16-2-6z"/>',
    sok: '<circle cx="11" cy="11" r="6"/><path d="M20 20l-4.5-4.5"/>',
    lukk: '<path d="M6 6l12 12M18 6L6 18"/>',
    klokke: '<circle cx="12" cy="13" r="7"/><path d="M12 9.5V13M10 3h4"/>',
    advarsel: '<path d="M12 4l9 16H3z"/><path d="M12 10v4M12 17h.01"/>',
    rediger: '<path d="M4 20h4L19 9l-4-4L4 16z"/>'
  };
  function ikon(navn, klasse) {
    return el("span", { class: "ikon" + (klasse ? " " + klasse : ""), "aria-hidden": "true",
      html: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="square" stroke-linejoin="miter">' + (IKON[navn] || "") + "</svg>" });
  }

  // =====================================================================
  // Hjelpere: format
  // =====================================================================
  var nf0 = new Intl.NumberFormat("nb-NO", { maximumFractionDigits: 0 });
  var nf2 = new Intl.NumberFormat("nb-NO", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  var nfT = new Intl.NumberFormat("nb-NO", { maximumFractionDigits: 2 });
  function kr(n) {
    n = Number(n) || 0;
    var hel = Math.abs(n - Math.round(n)) < 0.005;
    return "kr " + (hel ? nf0.format(Math.round(n)) : nf2.format(n));
  }
  function krHel(n) { return "kr " + nf0.format(Math.round(Number(n) || 0)); }
  function t(n) { return nfT.format(Number(n) || 0) + " t"; }
  function tTall(n) { return nfT.format(Number(n) || 0); }
  function dato(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || "");
    if (!m) return "";
    if (iso.length > 10) { var d = new Date(iso); return pad(d.getDate()) + "." + pad(d.getMonth() + 1) + "." + d.getFullYear(); }
    return m[3] + "." + m[2] + "." + m[1];
  }
  function klokkeslett(iso) { var d = new Date(iso); return pad(d.getHours()) + ":" + pad(d.getMinutes()); }
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function isoTilDate(iso) { var p = String(iso).slice(0, 10).split("-").map(Number); return new Date(p[0], p[1] - 1, p[2]); }
  function dagerMellom(fra, til) { return Math.round((isoTilDate(til) - isoTilDate(fra)) / 864e5); }
  var UKEDAG = ["søndag", "mandag", "tirsdag", "onsdag", "torsdag", "fredag", "lørdag"];
  var MND = ["jan.", "feb.", "mars", "april", "mai", "juni", "juli", "aug.", "sep.", "okt.", "nov.", "des."];
  function datoLang(iso) {
    var d = isoTilDate(iso);
    var diff = dagerMellom(iso, Store.idag());
    var basis = d.getDate() + ". " + MND[d.getMonth()];
    if (diff === 0) return "I dag, " + basis;
    if (diff === 1) return "I går, " + basis;
    return UKEDAG[d.getDay()].charAt(0).toUpperCase() + UKEDAG[d.getDay()].slice(1) + " " + basis;
  }
  function nedtelling(iso) {
    var n = dagerMellom(Store.idag(), iso);
    if (n === 0) return "i dag";
    if (n === 1) return "i morgen";
    if (n > 1) return "om " + n + " dager";
    if (n === -1) return "skulle gått i går";
    return "skulle gått for " + (-n) + " dager siden";
  }

  /** Tolker timer: «1,5», «1.5», «1:30», «90m», «1t 30m», «2 t». Returnerer tall eller NaN. */
  function tolkTimer(s) {
    s = String(s || "").trim().toLowerCase().replace(/\s+/g, " ");
    if (!s) return NaN;
    var m;
    if ((m = /^(\d{1,2}):(\d{1,2})$/.exec(s))) { var min = +m[2]; return min < 60 ? +m[1] + min / 60 : NaN; }
    if ((m = /^(\d+)\s*m(in)?$/.exec(s))) return +m[1] / 60;
    if ((m = /^(\d+(?:[.,]\d+)?)\s*t(imer?)?\s*(?:(\d+)\s*m(in)?)?$/.exec(s))) return parseFloat(m[1].replace(",", ".")) + (m[3] ? +m[3] / 60 : 0);
    if (/^\d+(?:[.,]\d+)?$/.test(s)) return parseFloat(s.replace(",", "."));
    return NaN;
  }
  function tolkBelop(s) {
    s = String(s || "").replace(/kr/i, "").replace(/[\s ]/g, "").replace(",", ".");
    return /^\d+(\.\d+)?$/.test(s) ? parseFloat(s) : NaN;
  }

  function erDemo() { return Store.modus() === "demo"; }

  var FIRMA = { navn: "Advokat Geir Frøholm AS", orgnr: "929 196 325" };
  var OPPHAV = "© 2026 Mathias Frøholm · Alle rettigheter forbeholdt";
  function prisGrenser() {
    var O = window.Oppdrag || {};
    return { min: Number(O.TIMEPRIS_MIN) || 2500, max: Number(O.TIMEPRIS_MAX) || 3000 };
  }
  function prisUvanlig(p) { var g = prisGrenser(); return isFinite(p) && p > 0 && (p < g.min || p > g.max); }
  function prisVarselTekst() { var g = prisGrenser(); return "Timeprisen ser uvanlig ut – vanlige satser er " + krHel(g.min) + "–" + nf0.format(g.max) + " eks. mva."; }
  function sendeTidTekst(iso) { var d = isoTilDate(iso); return UKEDAG[d.getDay()] + " " + dato(iso) + " kl. 07:00"; }

  function feilTekst(e) {
    if (!e) return "Noe gikk galt.";
    if (e.name === "Valideringsfeil" || e.name === "BrregFeil") return e.message;
    if (e.name === "GraphFeil") {
      if (e.status === 0) return e.message;
      if (e.status === 401 || e.status === 403) return "Mangler tilgang i Microsoft 365 (" + e.status + "). Logg inn på nytt.";
      if (e.status === 507) return "OneDrive er full.";
      return "Microsoft 365 svarte med feil " + e.status + ". " + (e.message || "");
    }
    return e.message || String(e);
  }

  function kundeNavn(id) { var k = Store.kunde(id); return k ? k.navn : "(ukjent kunde)"; }
  function sakEtikett(s) { return s ? s.saksnr + " " + s.navn : "(ukjent sak)"; }

  // =====================================================================
  // Varsler (toast) og dialoger
  // =====================================================================
  var toastOmr;
  function toast(tekst, type) {
    if (!toastOmr) return;
    var n = el("div", { class: "toast toast-" + (type || "info"), role: type === "feil" ? "alert" : "status" }, tekst);
    toastOmr.appendChild(n);
    setTimeout(function () { n.classList.add("ut"); setTimeout(function () { n.remove(); }, 300); }, type === "feil" ? 9000 : 4500);
  }

  /** Åpner en modal dialog. Returnerer { dialog, lukk }. */
  function aapneDialog(opts) {
    var tittelId = uid("dlg");
    var d = el("dialog", { class: "dialog" + (opts.klasse ? " " + opts.klasse : ""), "aria-labelledby": tittelId });
    var lukkKnapp = el("button", { type: "button", class: "ikonknapp dialog-lukk", "aria-label": "Lukk", onclick: function () { lukk(); } }, ikon("lukk"));
    var hode = el("div", { class: "dialog-hode" }, el("h2", { id: tittelId }, opts.tittel), lukkKnapp);
    var kropp = el("div", { class: "dialog-kropp" }, opts.innhold);
    d.appendChild(hode);
    d.appendChild(kropp);
    if (opts.fot) d.appendChild(el("div", { class: "dialog-fot" }, opts.fot));
    document.body.appendChild(d);
    var lukket = false;
    function lukk(verdi) {
      if (lukket) return;
      lukket = true;
      if (d.open) d.close();
      d.remove();
      if (opts.vedLukk) opts.vedLukk(verdi);
    }
    d.addEventListener("cancel", function (e) { e.preventDefault(); lukk(); });
    d.addEventListener("click", function (e) { if (e.target === d) lukk(); });
    d.showModal();
    var forste = opts.fokus ? d.querySelector(opts.fokus) : null;
    if (forste) forste.focus();
    return { dialog: d, lukk: lukk };
  }

  /** Bekreftelsesdialog. Promise<boolean>. */
  function bekreft(o) {
    return new Promise(function (ok) {
      var svar = false;
      var okKnapp = el("button", { type: "button", class: "knapp" + (o.farlig ? " fare" : ""), onclick: function () { svar = true; h.lukk(); } }, o.ok || "OK");
      var avbryt = el("button", { type: "button", class: "knapp sekundaer", onclick: function () { h.lukk(); } }, o.avbryt || "Avbryt");
      var h = aapneDialog({
        tittel: o.tittel, klasse: "smal",
        innhold: typeof o.tekst === "string" ? el("p", null, o.tekst) : o.tekst,
        fot: [avbryt, okKnapp],
        vedLukk: function () { ok(svar); }
      });
      okKnapp.focus();
    });
  }

  // =====================================================================
  // Skjemafelt
  // =====================================================================
  function felt(o) {
    var id = o.id || uid("f");
    var input;
    if (o.type === "textarea") {
      input = el("textarea", { id: id, name: o.navn, rows: o.rader || 3, placeholder: o.placeholder, required: o.paakrevd, disabled: o.disabled });
      input.value = o.verdi == null ? "" : o.verdi;
    } else if (o.type === "select") {
      input = el("select", { id: id, name: o.navn, required: o.paakrevd, disabled: o.disabled },
        (o.valg || []).map(function (v) { return el("option", { value: v.verdi, selected: String(v.verdi) === String(o.verdi) }, v.tekst); }));
    } else {
      input = el("input", {
        id: id, name: o.navn, type: o.type || "text", value: o.verdi == null ? "" : String(o.verdi),
        inputmode: o.inputmode, autocomplete: o.autocomplete || "off", placeholder: o.placeholder,
        required: o.paakrevd, min: o.min, max: o.max, step: o.step, readonly: o.readonly, disabled: o.disabled,
        spellcheck: o.spellcheck === undefined ? null : String(o.spellcheck), enterkeyhint: o.enterkeyhint
      });
    }
    var hintId = o.hint ? uid("h") : null;
    var feilId = uid("e");
    if (hintId) input.setAttribute("aria-describedby", hintId + " " + feilId);
    else input.setAttribute("aria-describedby", feilId);
    var feilEl = el("p", { class: "felt-feil", id: feilId, "aria-live": "polite" });
    var wrap = el("div", { class: "felt" + (o.klasse ? " " + o.klasse : "") },
      el("label", { for: id }, o.etikett, o.paakrevd ? el("span", { class: "paakrevd", "aria-hidden": "true" }, " *") : null),
      o.foran ? el("div", { class: "felt-rad" }, input, o.foran) : input,
      o.hint ? el("p", { class: "felt-hint", id: hintId }, o.hint) : null,
      feilEl);
    wrap.input = input;
    wrap.settFeil = function (tekst) {
      feilEl.textContent = tekst || "";
      if (tekst) input.setAttribute("aria-invalid", "true"); else input.removeAttribute("aria-invalid");
    };
    return wrap;
  }

  function bryter(o) {
    var id = o.id || uid("b");
    var input = el("input", { type: "checkbox", id: id, class: "bryter-input", checked: !!o.verdi, role: "switch" });
    var w = el("div", { class: "bryter" + (o.klasse ? " " + o.klasse : "") }, input,
      el("label", { for: id }, el("span", { class: "bryter-spor", "aria-hidden": "true" }), el("span", null, o.etikett)));
    w.input = input;
    return w;
  }

  function chip(tekst, type, tittel) {
    return el("span", { class: "chip chip-" + (type || "noytral"), title: tittel || null }, tekst);
  }
  function statusChip(linje) {
    var s = Store.linjeStatus(linje);
    if (s.type === "sendt") return chip("Sendt " + dato(s.dato), "sendt");
    if (s.type === "ikke") return chip("Ikke fakturerbar", "noytral");
    return chip("Ufakturert", "ufakturert");
  }

  function etikett(tekst, tag) { return el(tag || "p", { class: "etikett" }, tekst); }

  // =====================================================================
  // Sak-velger (combobox med søk)
  // =====================================================================
  function sakVelger(o) {
    var id = o.id || uid("sak");
    var listeId = id + "-liste";
    var valgt = o.verdi || "";
    var aktivIndeks = -1;
    var treff = [];
    var input = el("input", {
      id: id, type: "text", role: "combobox", "aria-autocomplete": "list", "aria-expanded": "false",
      "aria-controls": listeId, autocomplete: "off", spellcheck: "false",
      placeholder: "Søk på saksnr., klient eller sak"
    });
    var liste = el("ul", { id: listeId, role: "listbox", class: "kombo-liste", hidden: true });
    var feilEl = el("p", { class: "felt-feil", "aria-live": "polite" });
    var wrap = el("div", { class: "felt kombo" + (o.klasse ? " " + o.klasse : "") },
      el("label", { for: id }, o.etikett || "Sak"), el("div", { class: "kombo-boks" }, ikon("sok", "kombo-ikon"), input, liste), feilEl);

    function kandidater() {
      return Store.data().saker.filter(function (s) { return s.status !== "avsluttet" || s.id === valgt; })
        .sort(function (a, b) { return String(b.opprettet).localeCompare(String(a.opprettet)); });
    }
    function tekstFor(s) { return s.saksnr + " · " + kundeNavn(s.kundeId) + " – " + s.navn; }
    function visValgt() {
      var s = Store.sak(valgt);
      input.value = s ? tekstFor(s) : "";
    }
    function apne() {
      var q = input.value.trim().toLowerCase();
      var s = Store.sak(valgt);
      if (s && input.value === tekstFor(s)) q = "";
      treff = kandidater().filter(function (x) {
        if (!q) return true;
        return (x.saksnr + " " + x.navn + " " + kundeNavn(x.kundeId)).toLowerCase().indexOf(q) >= 0;
      }).slice(0, 8);
      tom(liste);
      if (!treff.length) {
        liste.appendChild(el("li", { class: "kombo-tom", role: "presentation" }, q ? "Ingen saker passer søket." : "Ingen aktive saker ennå."));
      }
      treff.forEach(function (x, i) {
        var li = el("li", { role: "option", id: listeId + "-" + i, "aria-selected": x.id === valgt ? "true" : "false", class: "kombo-valg" },
          el("span", { class: "kombo-nr" }, x.saksnr),
          el("span", { class: "kombo-tekst" }, el("strong", null, x.navn), el("small", null, kundeNavn(x.kundeId) + " · " + krHel(x.timepris) + "/t")));
        li.addEventListener("mousedown", function (e) { e.preventDefault(); velg(x.id); });
        liste.appendChild(li);
      });
      aktivIndeks = treff.findIndex(function (x) { return x.id === valgt; });
      markerAktiv();
      liste.hidden = false;
      input.setAttribute("aria-expanded", "true");
    }
    function lukk() {
      liste.hidden = true;
      input.setAttribute("aria-expanded", "false");
      input.removeAttribute("aria-activedescendant");
    }
    function markerAktiv() {
      Array.prototype.forEach.call(liste.querySelectorAll("[role=option]"), function (li, i) {
        li.classList.toggle("aktiv", i === aktivIndeks);
        if (i === aktivIndeks) { input.setAttribute("aria-activedescendant", li.id); li.scrollIntoView({ block: "nearest" }); }
      });
    }
    function velg(sid) {
      valgt = sid;
      visValgt();
      lukk();
      wrap.settFeil("");
      if (o.vedEndring) o.vedEndring(sid);
    }
    input.addEventListener("focus", function () { input.select(); apne(); });
    input.addEventListener("input", function () { apne(); });
    input.addEventListener("blur", function () { setTimeout(function () { lukk(); visValgt(); }, 120); });
    input.addEventListener("keydown", function (e) {
      if (e.key === "ArrowDown") { e.preventDefault(); if (liste.hidden) apne(); aktivIndeks = Math.min(treff.length - 1, aktivIndeks + 1); markerAktiv(); }
      else if (e.key === "ArrowUp") { e.preventDefault(); aktivIndeks = Math.max(0, aktivIndeks - 1); markerAktiv(); }
      else if (e.key === "Enter") {
        if (!liste.hidden && treff[aktivIndeks]) { e.preventDefault(); velg(treff[aktivIndeks].id); }
        else if (!liste.hidden && treff.length === 1) { e.preventDefault(); velg(treff[0].id); }
      } else if (e.key === "Escape") { if (!liste.hidden) { e.stopPropagation(); e.preventDefault(); lukk(); visValgt(); } }
    });
    visValgt();
    wrap.input = input;
    wrap.verdi = function () { return valgt; };
    wrap.sett = function (sid) { valgt = sid; visValgt(); };
    wrap.settFeil = function (tekst) {
      feilEl.textContent = tekst || "";
      if (tekst) input.setAttribute("aria-invalid", "true"); else input.removeAttribute("aria-invalid");
    };
    return wrap;
  }

  // =====================================================================
  // Føringsskjema (time eller utlegg) – brukes inline og i dialog
  // =====================================================================
  var sistBrukteSak = null;
  try { sistBrukteSak = sessionStorage.getItem("timer.sistSak"); } catch (e) { /* */ }
  function huskSak(id) { sistBrukteSak = id; try { sessionStorage.setItem("timer.sistSak", id); } catch (e) { /* */ } }

  function foringsskjema(o) {
    o = o || {};
    var linje = o.linje || null;
    var type = o.type || (linje && linje.belop !== undefined ? "utlegg" : "time");
    var standardSak = (linje && linje.sakId) || o.sakId || (Store.sak(sistBrukteSak) && Store.sak(sistBrukteSak).status !== "avsluttet" ? sistBrukteSak : "");

    var form = el("form", { class: "foring", novalidate: true });
    var typeVelger = null;
    if (!linje) {
      var tId = uid("type");
      typeVelger = el("div", { class: "segment", role: "radiogroup", "aria-label": "Type føring" },
        el("input", { type: "radio", name: tId, id: tId + "-t", value: "time", checked: type === "time" }),
        el("label", { for: tId + "-t" }, "Tid"),
        el("input", { type: "radio", name: tId, id: tId + "-u", value: "utlegg", checked: type === "utlegg" }),
        el("label", { for: tId + "-u" }, "Utlegg"));
      typeVelger.addEventListener("change", function (e) { type = e.target.value; byggFelter(); });
    }

    var sakF = sakVelger({ etikett: "Sak", verdi: standardSak, vedEndring: function () { oppdaterSats(); oppdaterSum(); } });
    var datoF = felt({ etikett: "Dato", type: "date", verdi: (linje && linje.dato) || o.dato || Store.idag(), paakrevd: true });
    var timerF = felt({ etikett: "Timer", verdi: linje && linje.timer != null ? tTall(linje.timer) : (o.timer != null ? tTall(o.timer) : ""), inputmode: "decimal", placeholder: "1,5 eller 1:30", paakrevd: true, enterkeyhint: "next", klasse: "felt-timer" });
    var hurtig = el("div", { class: "hurtigtimer", role: "group", "aria-label": "Hurtigvalg timer" },
      [0.25, 0.5, 1, 1.5, 2, 3].map(function (h) {
        return el("button", { type: "button", class: "hurtig", onclick: function () { timerF.input.value = tTall(h); timerF.settFeil(""); oppdaterSum(); } }, tTall(h));
      }));
    var beskrF = felt({ etikett: "Beskrivelse", type: "textarea", rader: 2, verdi: linje ? linje.beskrivelse : "", placeholder: "Hva ble gjort?" });
    var fakturerbarF = bryter({ etikett: "Fakturerbar", verdi: linje ? linje.fakturerbar !== false : true });
    var satsF = felt({ etikett: "Timesats", verdi: linje && linje.sats != null ? String(linje.sats) : "", inputmode: "decimal", hint: "Hentes fra saken. Endre bare ved avvikende sats.", klasse: "felt-sats" });
    var belopF = felt({ etikett: "Beløp (kr)", verdi: linje && linje.belop != null ? String(linje.belop).replace(".", ",") : "", inputmode: "decimal", paakrevd: true });
    var mvaF = bryter({ etikett: "MVA-pliktig utlegg", verdi: linje ? !!linje.mvaPliktig : false });
    var sumEl = el("p", { class: "foring-sum", "aria-live": "polite" });
    var satsDetaljer = el("details", { class: "foring-mer" }, el("summary", null, "Avvikende timesats"), satsF);
    if (linje && linje.sats != null && Store.sak(linje.sakId) && Number(linje.sats) !== Number(Store.sak(linje.sakId).timepris)) satsDetaljer.open = true;

    var felterBoks = el("div", { class: "foring-felter" });
    var lagre = el("button", { type: "submit", class: "knapp" }, linje ? "Lagre endringer" : (type === "utlegg" ? "Lagre utlegg" : "Lagre føring"));
    var knapper = el("div", { class: "knapperad" }, lagre, o.ekstraKnapper || null);

    function oppdaterSats() {
      var s = Store.sak(sakF.verdi());
      if (!satsDetaljer.open || !satsF.input.value) satsF.input.value = s ? String(s.timepris) : "";
    }
    function oppdaterSum() {
      if (type !== "time") { sumEl.textContent = ""; return; }
      var h = tolkTimer(timerF.input.value);
      var sats = tolkBelop(satsF.input.value);
      if (!isFinite(sats)) { var s = Store.sak(sakF.verdi()); sats = s ? s.timepris : NaN; }
      if (!fakturerbarF.input.checked) { sumEl.textContent = isFinite(h) ? t(h) + " · ikke fakturerbart" : ""; return; }
      sumEl.textContent = isFinite(h) && isFinite(sats) ? t(h) + " × " + krHel(sats) + " = " + kr(h * sats) : "";
    }
    function byggFelter() {
      tom(felterBoks);
      if (type === "time") {
        felterBoks.appendChild(el("div", { class: "foring-grid" }, sakF, datoF, el("div", { class: "felt-gruppe-timer" }, timerF, hurtig), beskrF));
        felterBoks.appendChild(el("div", { class: "foring-bunn" }, fakturerbarF, satsDetaljer, sumEl));
        lagre.textContent = linje ? "Lagre endringer" : "Lagre føring";
      } else {
        felterBoks.appendChild(el("div", { class: "foring-grid" }, sakF, datoF, belopF, beskrF));
        felterBoks.appendChild(el("div", { class: "foring-bunn" }, mvaF));
        lagre.textContent = linje ? "Lagre endringer" : "Lagre utlegg";
      }
      oppdaterSum();
    }
    var satsVarsel = el("p", { class: "felt-varsel", "aria-live": "polite" });
    satsF.appendChild(satsVarsel);
    function sjekkSats() { satsVarsel.textContent = satsF.input.value.trim() && prisUvanlig(tolkBelop(satsF.input.value)) ? prisVarselTekst() : ""; }
    timerF.input.addEventListener("input", function () { timerF.settFeil(""); oppdaterSum(); });
    satsF.input.addEventListener("input", function () { sjekkSats(); oppdaterSum(); });
    fakturerbarF.input.addEventListener("change", oppdaterSum);

    form.addEventListener("submit", async function (e) {
      e.preventDefault();
      var ok = true;
      var sid = sakF.verdi();
      if (!sid) { sakF.settFeil("Velg en sak."); ok = false; }
      if (!datoF.input.value) { datoF.settFeil("Velg dato."); ok = false; } else datoF.settFeil("");
      var data = { sakId: sid, dato: datoF.input.value, beskrivelse: beskrF.input.value };
      if (type === "time") {
        var h = tolkTimer(timerF.input.value);
        if (!isFinite(h) || h <= 0 || h > 24) { timerF.settFeil("Skriv timer som 1,5 eller 1:30 (maks 24)."); ok = false; }
        else timerF.settFeil("");
        var sats = satsF.input.value.trim() ? tolkBelop(satsF.input.value) : (Store.sak(sid) ? Store.sak(sid).timepris : 0);
        if (!isFinite(sats)) { satsF.settFeil("Ugyldig sats."); satsDetaljer.open = true; ok = false; }
        data.timer = h; data.sats = sats; data.fakturerbar = fakturerbarF.input.checked;
      } else {
        var b = tolkBelop(belopF.input.value);
        if (!isFinite(b) || b <= 0) { belopF.settFeil("Skriv et beløp, f.eks. 350 eller 1 243,50."); ok = false; } else belopF.settFeil("");
        if (!beskrF.input.value.trim()) { beskrF.settFeil("Skriv hva utlegget gjelder."); ok = false; } else beskrF.settFeil("");
        data.belop = b; data.mvaPliktig = mvaF.input.checked;
      }
      if (!ok) { var f = form.querySelector("[aria-invalid=true]"); if (f) f.focus(); return; }
      lagre.disabled = true;
      try {
        var res;
        if (linje) res = type === "time" ? await Store.endreTime(linje.id, data) : await Store.endreUtlegg(linje.id, data);
        else res = type === "time" ? await Store.leggTilTime(data) : await Store.leggTilUtlegg(data);
        huskSak(sid);
        if (o.vedLagret) o.vedLagret(res, type);
        if (!linje && o.nullstill !== false) {
          timerF.input.value = ""; beskrF.input.value = ""; belopF.input.value = "";
          oppdaterSum();
        }
      } catch (err) {
        toast(feilTekst(err), "feil");
      } finally {
        lagre.disabled = false;
      }
    });

    if (typeVelger) form.appendChild(typeVelger);
    form.appendChild(felterBoks);
    form.appendChild(knapper);
    oppdaterSats();
    sjekkSats();
    byggFelter();
    form.sakFelt = sakF;
    form.timerFelt = timerF;
    form.beskrFelt = beskrF;
    form.sett = function (v) {
      if (v.sakId) { sakF.sett(v.sakId); oppdaterSats(); }
      if (v.timer != null) timerF.input.value = tTall(v.timer);
      if (v.beskrivelse != null) beskrF.input.value = v.beskrivelse;
      oppdaterSum();
    };
    form.fokuser = function () {
      if (!sakF.verdi()) sakF.input.focus();
      else if (type === "time") timerF.input.focus();
      else belopF.input.focus();
    };
    return form;
  }

  function hurtigforing(forhånd) {
    if (!Store.data().saker.some(function (s) { return s.status !== "avsluttet"; })) {
      toast("Opprett en sak først – under Saker og kunder eller Ny oppdragsbekreftelse.", "info");
      gaaTil("saker");
      return;
    }
    var h;
    var skjema = foringsskjema(Object.assign({}, forhånd || {}, {
      vedLagret: function (res, type) {
        h.lukk();
        toast(type === "utlegg" ? "Utlegget er lagret." : "Føringen er lagret: " + t(res.timer) + " på " + (Store.sak(res.sakId) || {}).saksnr + ".", "ok");
      }
    }));
    h = aapneDialog({ tittel: "Før timer", innhold: skjema, klasse: "dialog-foring" });
    skjema.fokuser();
  }

  function redigerLinje(linje) {
    var erUtlegg = linje.belop !== undefined;
    if (Store.erLast(linje)) {
      toast("Linjen er sendt til regnskap og er låst.", "info");
      return;
    }
    var h;
    var slett = el("button", { type: "button", class: "knapp tekst fare", onclick: async function () {
      var ok = await bekreft({ tittel: erUtlegg ? "Slette utlegget?" : "Slette føringen?", tekst: (Store.sak(linje.sakId) || {}).saksnr + " · " + dato(linje.dato) + " · " + (erUtlegg ? kr(linje.belop) : t(linje.timer)) + (linje.beskrivelse ? " – " + linje.beskrivelse : ""), ok: "Slett", farlig: true });
      if (!ok) return;
      try {
        if (erUtlegg) await Store.slettUtlegg(linje.id); else await Store.slettTime(linje.id);
        h.lukk();
        toast(erUtlegg ? "Utlegget er slettet." : "Føringen er slettet.", "ok");
      } catch (e) { toast(feilTekst(e), "feil"); }
    } }, "Slett");
    var skjema = foringsskjema({ linje: linje, type: erUtlegg ? "utlegg" : "time", ekstraKnapper: slett, vedLagret: function () { h.lukk(); toast("Endringene er lagret.", "ok"); } });
    h = aapneDialog({ tittel: erUtlegg ? "Rediger utlegg" : "Rediger føring", innhold: skjema, klasse: "dialog-foring" });
  }

  // =====================================================================
  // Stoppeklokke (én om gangen, i sessionStorage/demo: localStorage – kun sak-ID og starttid)
  // =====================================================================
  var KLOKKE = "timer.klokke";
  function klokkeLager() { return erDemo() ? window.localStorage : window.sessionStorage; }
  function hentKlokke() { try { var v = klokkeLager().getItem(KLOKKE); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
  function settKlokke(v) { try { if (v) klokkeLager().setItem(KLOKKE, JSON.stringify(v)); else klokkeLager().removeItem(KLOKKE); } catch (e) { /* */ } oppdaterKlokkeVisning(); }
  function klokkeTekst(ms) {
    var s = Math.floor(ms / 1000);
    return pad(Math.floor(s / 3600)) + ":" + pad(Math.floor(s / 60) % 60) + ":" + pad(s % 60);
  }
  var klokkeIntervall = null;
  var klokkeLyttere = [];
  function oppdaterKlokkeVisning() {
    var k = hentKlokke();
    klokkeLyttere.forEach(function (f) { f(k); });
    if (k && !klokkeIntervall) klokkeIntervall = setInterval(function () { var kk = hentKlokke(); klokkeLyttere.forEach(function (f) { f(kk); }); }, 1000);
    if (!k && klokkeIntervall) { clearInterval(klokkeIntervall); klokkeIntervall = null; }
  }

  // =====================================================================
  // Applikasjonsskall
  // =====================================================================
  var SEKSJONER = [
    { id: "oversikt", tittel: "Oversikt", ikon: "oversikt" },
    { id: "timer", tittel: "Før timer", kort: "Timer", ikon: "timer" },
    { id: "saker", tittel: "Saker og kunder", kort: "Saker", ikon: "saker" },
    { id: "oppdrag", tittel: "Ny oppdragsbekreftelse", kort: "Oppdrag", ikon: "oppdrag" },
    { id: "faktura", tittel: "Fakturering", ikon: "faktura" },
    { id: "innst", tittel: "Innstillinger", ikon: "innst" }
  ];
  var aktiv = "oversikt";
  var visning = null;      // { node, oppdater }
  var hovedEl, tittelEl, navKnapper = {}, varselOmr, klokkeChip;

  function gaaTil(id, opts) {
    aktiv = id;
    var s = SEKSJONER.find(function (x) { return x.id === id; });
    tittelEl.textContent = s.tittel;
    document.title = s.tittel + " – Timeføring";
    Object.keys(navKnapper).forEach(function (k) {
      navKnapper[k].forEach(function (b) { if (k === id) b.setAttribute("aria-current", "page"); else b.removeAttribute("aria-current"); });
    });
    var fabrikk = { oversikt: visOversikt, timer: visTimer, saker: visSaker, oppdrag: visOppdrag, faktura: visFaktura, innst: visInnst }[id];
    visning = fabrikk(opts || {});
    erstatt(hovedEl, visning.node);
    window.scrollTo(0, 0);
    if (!opts || !opts.ikkeFokus) hovedEl.focus({ preventScroll: true });
  }

  function oppdaterAlt() {
    oppdaterVarsler();
    if (visning && visning.oppdater) visning.oppdater();
  }

  function oppdaterVarsler() {
    var i = Store.innst();
    tom(varselOmr);
    if (!i.mottakerEpost) {
      varselOmr.appendChild(el("button", { type: "button", class: "chip chip-advarsel chip-knapp", onclick: function () { gaaTil("innst"); } }, ikon("advarsel"), "Mangler e-post til View"));
    }
    var fs = Store.filStatus();
    if (fs.feil) varselOmr.appendChild(chip("Grunnlaget ble ikke oppdatert", "advarsel", fs.feil));
  }

  function byggSkall(modus) {
    var konto = modus === "demo" ? { username: "demo@eksempel.invalid", name: "Demo" } : Auth.konto();
    tom(rot);
    rot.className = "app" + (modus === "demo" ? " er-demo" : "");

    var banner = modus === "demo" ? el("div", { class: "demo-banner", role: "note" },
      el("strong", null, "Demo – eksempeldata, ikke ekte klienter."),
      el("span", { class: "demo-banner-mer" }, " Endringer lagres bare i denne nettleseren. Ingen e-post sendes."),
      el("button", { type: "button", class: "demo-avslutt", onclick: function () { avslutt(null); } }, "Avslutt demo")) : null;

    // Sidenav (desktop)
    navKnapper = {};
    function navKnapp(s, klasse) {
      var b = el("button", { type: "button", class: klasse, onclick: function () { gaaTil(s.id); } }, ikon(s.ikon), el("span", null, s.kort && klasse === "bunn-knapp" ? s.kort : s.tittel));
      (navKnapper[s.id] = navKnapper[s.id] || []).push(b);
      return b;
    }
    var sidenav = el("nav", { class: "sidenav", "aria-label": "Hovedmeny" },
      el("div", { class: "sidenav-logo" }, el("img", { src: "/assets/logo-negativ.svg", alt: "Advokat Geir Frøholm AS", width: "500", height: "148" })),
      el("p", { class: "sidenav-app" }, "Timeføring"),
      el("button", { type: "button", class: "knapp messing sidenav-for", onclick: function () { hurtigforing(); } }, ikon("pluss"), "Før timer", el("kbd", { "aria-hidden": "true" }, "T")),
      el("ul", { class: "sidenav-liste" }, SEKSJONER.map(function (s) { return el("li", null, navKnapp(s, "nav-knapp")); })),
      el("div", { class: "sidenav-bunn" },
        el("p", { class: "sidenav-konto" }, el("small", null, modus === "demo" ? "Demo-modus" : "Innlogget som"), el("span", null, konto ? konto.username : "")),
        el("button", { type: "button", class: "nav-knapp", onclick: function () { avslutt(null); } }, ikon("ut"), el("span", null, modus === "demo" ? "Avslutt demo" : "Logg ut")),
        el("p", { class: "opphav opphav-sidenav" }, OPPHAV)));

    // Toppfelt
    tittelEl = el("h1", { class: "topp-tittel" });
    varselOmr = el("div", { class: "topp-varsler" });
    klokkeChip = el("button", { type: "button", class: "chip chip-klokke chip-knapp", hidden: true, onclick: function () { gaaTil("timer"); } });
    var topp = el("header", { class: "topp" },
      el("img", { class: "topp-symbol", src: "/assets/symbol-negativ.svg", alt: "", width: "116", height: "116" }),
      tittelEl,
      el("div", { class: "topp-hoyre" }, klokkeChip, varselOmr,
        el("button", { type: "button", class: "knapp messing topp-for", onclick: function () { hurtigforing(); } }, ikon("pluss"), "Før timer")));

    hovedEl = el("main", { class: "hoved", id: "hoved", tabindex: "-1" });

    // Bunnmeny (mobil)
    var merKnapp = el("button", { type: "button", class: "bunn-knapp", "aria-haspopup": "dialog", onclick: visMerMeny }, ikon("mer"), el("span", null, "Mer"));
    ["oppdrag", "faktura", "innst"].forEach(function (id) { (navKnapper[id] = navKnapper[id] || []).push(merKnapp); });
    var bunnmeny = el("nav", { class: "bunnmeny", "aria-label": "Hovedmeny (mobil)" },
      navKnapp(SEKSJONER[0], "bunn-knapp"),
      navKnapp(SEKSJONER[1], "bunn-knapp"),
      el("button", { type: "button", class: "bunn-for", "aria-label": "Ny føring", onclick: function () { hurtigforing(); } }, ikon("pluss")),
      navKnapp(SEKSJONER[2], "bunn-knapp"),
      merKnapp);

    toastOmr = el("div", { class: "toast-omr", "aria-live": "polite" });
    rot.appendChild(el("a", { class: "hopp", href: "#hoved" }, "Hopp til innhold"));
    if (banner) rot.appendChild(banner);
    rot.appendChild(el("div", { class: "ramme" }, sidenav, el("div", { class: "innhold" }, topp, hovedEl)));
    rot.appendChild(bunnmeny);
    rot.appendChild(toastOmr);

    klokkeLyttere = [function (k) {
      if (!k) { klokkeChip.hidden = true; return; }
      var s = Store.sak(k.sakId);
      klokkeChip.hidden = false;
      erstatt(klokkeChip, ikon("klokke"), klokkeTekst(Date.now() - k.start), s ? " · " + s.saksnr : "");
      klokkeChip.setAttribute("aria-label", "Stoppeklokke går: " + klokkeTekst(Date.now() - k.start) + (s ? " på sak " + s.saksnr : "") + ". Gå til Før timer.");
    }];
    oppdaterKlokkeVisning();
  }

  function visMerMeny() {
    var h;
    function valg(id) {
      var s = SEKSJONER.find(function (x) { return x.id === id; });
      return el("button", { type: "button", class: "mer-valg", onclick: function () { h.lukk(); gaaTil(id); } }, ikon(s.ikon), s.tittel);
    }
    h = aapneDialog({
      tittel: "Mer", klasse: "dialog-ark",
      innhold: el("div", { class: "mer-liste" }, valg("oppdrag"), valg("faktura"), valg("innst"),
        el("button", { type: "button", class: "mer-valg", onclick: function () { h.lukk(); avslutt(null); } }, ikon("ut"), erDemo() ? "Avslutt demo" : "Logg ut"),
        el("p", { class: "opphav opphav-mer" }, OPPHAV))
    });
  }

  // =====================================================================
  // Seksjon: Oversikt
  // =====================================================================
  function panel(tittel, innhold, ekstra) {
    return el("section", { class: "panel" + (ekstra && ekstra.klasse ? " " + ekstra.klasse : "") },
      el("div", { class: "panel-hode" }, el("h2", null, tittel), ekstra && ekstra.handling ? ekstra.handling : null),
      innhold);
  }

  function sumUfakturert() {
    var g = null;
    try { g = Store.grunnlag(); } catch (e) { console.error(e); }
    return g;
  }

  function sistSendtDato() {
    var b = Store.sendteBatcher()[0];
    return b ? b.sendt.slice(0, 10) : null;
  }

  function visOversikt() {
    var node = el("div", { class: "seksjon oversikt" });
    function oppdater() {
      var i = Store.innst();
      var g = sumUfakturert();
      var u = Store.ufakturerte();
      var alleTimer = u.timer.reduce(function (s, x) { return s + Number(x.timer || 0); }, 0);
      var ikkeFakt = u.timer.filter(function (x) { return x.fakturerbar === false; }).reduce(function (s, x) { return s + Number(x.timer || 0); }, 0);
      var aktiveSaker = Store.data().saker.filter(function (s) { return s.status !== "avsluttet"; }).length;

      // Periodelinje: fra forrige sending til neste
      var neste = Store.nesteSending();
      var periodeStart = Store.leggTilDager(neste, -(i.intervallDager || 14));
      var sist = sistSendtDato();
      var start = sist && sist < neste && sist >= Store.leggTilDager(neste, -2 * (i.intervallDager || 14)) ? sist : periodeStart;
      var lengde = Math.max(1, dagerMellom(start, neste));
      var gaatt = Math.min(lengde, Math.max(0, dagerMellom(start, Store.idag())));
      var andel = gaatt / lengde;

      var sendeTekst = i.sendingAktiv
        ? [el("span", { class: "hero-dato" }, dato(neste)), el("span", { class: "hero-ned" }, nedtelling(neste) + " · kl. 07:00")]
        : [el("span", { class: "hero-dato" }, "Av"), el("span", { class: "hero-ned" }, "Automatisk sending er slått av")];

      var mottaker = i.mottakerEpost
        ? el("p", { class: "hero-mottaker" }, "Til " + i.mottakerNavn + ", View")
        : el("button", { type: "button", class: "hero-mangler", onclick: function () { gaaTil("innst"); } }, ikon("advarsel"), "Mangler e-post til View – legg inn under Innstillinger");

      var hero = el("section", { class: "hero", "aria-label": "Sammendrag" },
        el("div", { class: "hero-belop" },
          etikett("Ufakturert"),
          el("p", { class: "hero-tall" }, g ? krHel(g.sum.grunnlag) : "–"),
          el("p", { class: "hero-under" }, g ? "eks. mva · " + kr(g.sum.total) + " inkl. mva" : "Faktureringsmodulen er ikke lastet")),
        el("div", { class: "hero-sending" },
          etikett("Neste sending til regnskap"),
          el("p", { class: "hero-neste" }, sendeTekst),
          el("div", { class: "periode", role: "img", "aria-label": "Perioden " + dato(start) + " til " + dato(neste) + ": " + gaatt + " av " + lengde + " dager er gått." },
            el("div", { class: "periode-spor" }, el("div", { class: "periode-fyll", style: "width:" + (andel * 100).toFixed(1) + "%" }),
              el("div", { class: "periode-idag", style: "left:" + (andel * 100).toFixed(1) + "%" })),
            el("div", { class: "periode-akse" }, el("span", null, dato(start)), el("span", null, dato(neste)))),
          mottaker));

      var nokkel = el("div", { class: "nokkeltall" },
        nokkeltall("Timer siden sist sending", t(alleTimer), ikkeFakt ? "herav " + t(ikkeFakt) + " ikke fakturerbart" : (u.timer.length + " føringer")),
        nokkeltall("Fakturerbart salær", g ? krHel(g.sum.salaer) : "–", g && g.sum.utlegg ? "+ " + kr(g.sum.utlegg) + " utlegg" : "eks. mva"),
        nokkeltall("Aktive saker", String(aktiveSaker), Store.data().kunder.length + " kunder"));

      // Timer per sak denne perioden
      var perSak = {};
      u.timer.forEach(function (x) { perSak[x.sakId] = perSak[x.sakId] || { timer: 0, belop: 0 }; perSak[x.sakId].timer += Number(x.timer) || 0; if (x.fakturerbar !== false) perSak[x.sakId].belop += (Number(x.timer) || 0) * (Number(x.sats) || 0); });
      var rader = Object.keys(perSak).map(function (k) { return { sak: Store.sak(k), v: perSak[k] }; }).sort(function (a, b) { return b.v.timer - a.v.timer; });
      var maks = rader.reduce(function (m, r) { return Math.max(m, r.v.timer); }, 0) || 1;
      var stolper = rader.length ? el("ul", { class: "stolper" }, rader.map(function (r) {
        return el("li", { class: "stolpe" },
          el("div", { class: "stolpe-tekst" }, el("span", { class: "stolpe-nr" }, r.sak ? r.sak.saksnr : ""), el("span", { class: "stolpe-navn" }, r.sak ? r.sak.navn : "(ukjent sak)"), el("small", null, r.sak ? kundeNavn(r.sak.kundeId) : "")),
          el("div", { class: "stolpe-spor", "aria-hidden": "true" }, el("div", { class: "stolpe-fyll", style: "width:" + (r.v.timer / maks * 100).toFixed(1) + "%" })),
          el("div", { class: "stolpe-tall" }, el("strong", null, t(r.v.timer)), el("small", null, krHel(r.v.belop))));
      })) : tomTilstand("Ingen timer ført siden forrige sending.", "Før timer", function () { hurtigforing(); });

      // Siste føringer
      var siste = Store.data().timer.slice().sort(function (a, b) { return (b.dato + b.opprettet).localeCompare(a.dato + a.opprettet); }).slice(0, 6);
      var sisteListe = siste.length ? el("ul", { class: "linjer" }, siste.map(linjeRad)) : tomTilstand("Ingen føringer ennå.", "Før timer", function () { hurtigforing(); });

      erstatt(node,
        hero,
        nokkel,
        el("div", { class: "to-kol" },
          panel("Timer per sak denne perioden", stolper),
          panel("Siste føringer", sisteListe, { handling: el("button", { type: "button", class: "knapp tekst", onclick: function () { gaaTil("timer"); } }, "Alle føringer") })));
    }
    oppdater();
    return { node: node, oppdater: oppdater };
  }

  function nokkeltall(etikettTekst, verdi, under) {
    return el("div", { class: "nokkel" }, el("p", { class: "nokkel-etikett" }, etikettTekst), el("p", { class: "nokkel-verdi" }, verdi), el("p", { class: "nokkel-under" }, under));
  }

  function tomTilstand(tekst, knappTekst, handling) {
    return el("div", { class: "tom-tilstand" }, el("p", null, tekst),
      knappTekst ? el("button", { type: "button", class: "knapp sekundaer", onclick: handling }, knappTekst) : null);
  }

  /** En rad for en time- eller utleggslinje. Klikk for å redigere (hvis ikke låst). */
  function linjeRad(l) {
    var erUtlegg = l.belop !== undefined;
    var s = Store.sak(l.sakId);
    var last = Store.erLast(l);
    var innhold = [
      el("span", { class: "linje-hoved" },
        el("span", { class: "linje-sak" }, s ? s.saksnr : "", el("span", { class: "linje-saknavn" }, s ? " " + s.navn : "")),
        el("span", { class: "linje-beskr" }, (erUtlegg ? "Utlegg: " : "") + (l.beskrivelse || "Uten beskrivelse"))),
      el("span", { class: "linje-meta" },
        el("span", { class: "linje-tall" }, erUtlegg ? kr(l.belop) : t(l.timer)),
        el("span", { class: "linje-status" }, last ? ikon("las", "ikon-liten") : null, statusChip(l)))
    ];
    if (last) return el("li", { class: "linje er-last", title: "Sendt til regnskap – låst" }, innhold);
    return el("li", { class: "linje" }, el("button", { type: "button", class: "linje-knapp", onclick: function () { redigerLinje(l); },
      "aria-label": "Rediger " + (erUtlegg ? "utlegg " + kr(l.belop) : "føring " + t(l.timer)) + ", " + (s ? s.saksnr : "") + ", " + dato(l.dato) }, innhold));
  }

  // =====================================================================
  // Seksjon: Før timer
  // =====================================================================
  var timerFilter = { status: "ufakturert", sok: "" };

  function visTimer() {
    var node = el("div", { class: "seksjon timer" });
    var listeOmr = el("div", { class: "foringsliste" });

    var skjema = foringsskjema({ vedLagret: function (res, type) {
      toast(type === "utlegg" ? "Utlegget er lagret." : "Føringen er lagret: " + t(res.timer) + ".", "ok");
      skjema.fokuser();
    } });

    // Stoppeklokke
    var klokkeVelger = sakVelger({ etikett: "Sak for stoppeklokken", verdi: sistBrukteSak || "" });
    var visTid = el("p", { class: "klokke-tid", "aria-live": "off" }, "00:00:00");
    var klokkeKnapp = el("button", { type: "button", class: "knapp" });
    var forkast = el("button", { type: "button", class: "knapp tekst", hidden: true }, "Forkast");
    var klokkeSak = el("p", { class: "klokke-sak" });
    function tegnKlokke(k) {
      if (k) {
        visTid.textContent = klokkeTekst(Date.now() - k.start);
        klokkeKnapp.textContent = "Stopp og før";
        klokkeVelger.hidden = true;
        forkast.hidden = false;
        var s = Store.sak(k.sakId);
        klokkeSak.textContent = s ? s.saksnr + " · " + s.navn : "";
        klokkeSak.hidden = false;
      } else {
        visTid.textContent = "00:00:00";
        klokkeKnapp.textContent = "Start";
        klokkeVelger.hidden = false;
        forkast.hidden = true;
        klokkeSak.hidden = true;
      }
    }
    klokkeKnapp.addEventListener("click", function () {
      var k = hentKlokke();
      if (!k) {
        var sid = klokkeVelger.verdi();
        if (!sid) { klokkeVelger.settFeil("Velg sak før du starter."); klokkeVelger.input.focus(); return; }
        settKlokke({ sakId: sid, start: Date.now() });
        toast("Stoppeklokken er startet.", "info");
      } else {
        var timer = Math.max(0.25, Math.ceil((Date.now() - k.start) / (15 * 60 * 1000)) / 4);
        settKlokke(null);
        skjema.sett({ sakId: k.sakId, timer: timer });
        skjema.beskrFelt.input.focus();
        toast("Tiden er fylt inn: " + t(timer) + " (rundet opp til nærmeste kvarter). Skriv beskrivelse og lagre.", "info");
      }
    });
    forkast.addEventListener("click", async function () {
      if (await bekreft({ tittel: "Forkaste stoppeklokken?", tekst: "Tiden som er målt blir ikke ført.", ok: "Forkast", farlig: true })) settKlokke(null);
    });
    var klokkeLytter = tegnKlokke;
    klokkeLyttere.push(klokkeLytter);

    var klokkePanel = panel("Stoppeklokke", el("div", { class: "klokke" }, klokkeVelger, klokkeSak, visTid, el("div", { class: "knapperad" }, klokkeKnapp, forkast),
      el("p", { class: "felt-hint" }, "Tiden rundes opp til nærmeste kvarter når du stopper.")), { klasse: "panel-klokke" });

    // Filter
    var fId = uid("filt");
    var statusFilter = el("div", { class: "segment liten", role: "radiogroup", "aria-label": "Vis føringer" },
      [["ufakturert", "Ufakturerte"], ["sendt", "Sendte"], ["alle", "Alle"]].map(function (v) {
        return [el("input", { type: "radio", name: fId, id: fId + v[0], value: v[0], checked: timerFilter.status === v[0] }), el("label", { for: fId + v[0] }, v[1])];
      }));
    statusFilter.addEventListener("change", function (e) { timerFilter.status = e.target.value; tegnListe(); });
    var sokF = felt({ etikett: "Søk i føringer", type: "search", verdi: timerFilter.sok, placeholder: "Sak, klient eller beskrivelse", klasse: "felt-sok" });
    sokF.input.addEventListener("input", function () { timerFilter.sok = sokF.input.value; tegnListe(); });

    function tegnListe() {
      var d = Store.data();
      var q = timerFilter.sok.trim().toLowerCase();
      var linjer = d.timer.concat(d.utlegg).filter(function (l) {
        var last = Store.erLast(l);
        if (timerFilter.status === "ufakturert" && last) return false;
        if (timerFilter.status === "sendt" && !last) return false;
        if (q) {
          var s = Store.sak(l.sakId);
          var hay = ((s ? s.saksnr + " " + s.navn + " " + kundeNavn(s.kundeId) : "") + " " + (l.beskrivelse || "")).toLowerCase();
          if (hay.indexOf(q) < 0) return false;
        }
        return true;
      }).sort(function (a, b) { return (b.dato + (b.opprettet || "")).localeCompare(a.dato + (a.opprettet || "")); });

      if (!linjer.length) {
        erstatt(listeOmr, tomTilstand(q ? "Ingen føringer passer søket." : (timerFilter.status === "sendt" ? "Ingen sendte føringer ennå." : "Ingen ufakturerte føringer. Før den første over.")));
        return;
      }
      var grupper = [];
      linjer.forEach(function (l) {
        var g = grupper[grupper.length - 1];
        if (!g || g.dato !== l.dato) { g = { dato: l.dato, linjer: [] }; grupper.push(g); }
        g.linjer.push(l);
      });
      erstatt(listeOmr, grupper.slice(0, 60).map(function (g) {
        var sumT = g.linjer.reduce(function (s, l) { return s + (l.timer ? Number(l.timer) : 0); }, 0);
        return el("section", { class: "dag" },
          el("h3", { class: "dag-hode" }, el("span", null, datoLang(g.dato)), el("span", { class: "dag-sum" }, sumT ? t(sumT) : "")),
          el("ul", { class: "linjer" }, g.linjer.map(linjeRad)));
      }));
    }

    node.appendChild(el("div", { class: "timer-topp" },
      panel("Ny føring", skjema, { klasse: "panel-foring" }),
      klokkePanel));
    node.appendChild(el("section", { class: "panel panel-liste" },
      el("div", { class: "panel-hode" }, el("h2", null, "Føringer"), statusFilter),
      sokF, listeOmr));

    tegnKlokke(hentKlokke());
    tegnListe();
    return {
      node: node,
      oppdater: function () { tegnListe(); },
      fjern: function () { klokkeLyttere = klokkeLyttere.filter(function (f) { return f !== klokkeLytter; }); }
    };
  }

  // =====================================================================
  // Seksjon: Saker og kunder
  // =====================================================================
  var sakerFilter = { sok: "", visAvsluttede: false };

  function visSaker() {
    var node = el("div", { class: "seksjon saker" });
    var listeOmr = el("div", { class: "kundeliste" });
    var sokF = felt({ etikett: "Søk", type: "search", verdi: sakerFilter.sok, placeholder: "Klient, saksnr., org.nr. eller sak", klasse: "felt-sok" });
    sokF.input.addEventListener("input", function () { sakerFilter.sok = sokF.input.value; tegn(); });
    var avsl = bryter({ etikett: "Vis avsluttede saker", verdi: sakerFilter.visAvsluttede });
    avsl.input.addEventListener("change", function () { sakerFilter.visAvsluttede = avsl.input.checked; tegn(); });

    function ufaktPerSak() {
      var m = {};
      var u = Store.ufakturerte();
      u.timer.forEach(function (x) { m[x.sakId] = m[x.sakId] || { t: 0, b: 0 }; m[x.sakId].t += Number(x.timer) || 0; if (x.fakturerbar !== false) m[x.sakId].b += (Number(x.timer) || 0) * (Number(x.sats) || 0); });
      u.utlegg.forEach(function (x) { m[x.sakId] = m[x.sakId] || { t: 0, b: 0 }; m[x.sakId].b += Number(x.belop) || 0; });
      return m;
    }

    function tegn() {
      var d = Store.data();
      var q = sakerFilter.sok.trim().toLowerCase();
      var uf = ufaktPerSak();
      var kunder = d.kunder.slice().sort(function (a, b) { return String(a.navn).localeCompare(String(b.navn), "nb"); });
      var blokker = [];
      kunder.forEach(function (k) {
        var saker = d.saker.filter(function (s) { return s.kundeId === k.id && (sakerFilter.visAvsluttede || s.status !== "avsluttet"); })
          .sort(function (a, b) { return String(a.saksnr).localeCompare(String(b.saksnr), "nb", { numeric: true }); });
        var kundeTreff = !q || (k.navn + " " + (k.orgnr || "") + " " + k.kundenr + " " + (k.kontaktperson || "")).toLowerCase().indexOf(q) >= 0;
        var sakTreff = saker.filter(function (s) { return (s.saksnr + " " + s.navn + " " + (s.motpart || "")).toLowerCase().indexOf(q) >= 0; });
        if (!kundeTreff && !sakTreff.length) return;
        var vis = kundeTreff ? saker : sakTreff;
        blokker.push(el("section", { class: "kunde" },
          el("div", { class: "kunde-hode" },
            el("div", { class: "kunde-id" },
              el("p", { class: "kunde-nr" }, "Kundenr. " + k.kundenr),
              el("h3", null, k.navn),
              el("p", { class: "kunde-info" }, [k.orgnr ? "Org.nr. " + Brreg.formaterOrgnr(k.orgnr) : "Uten org.nr.", k.kontaktperson, k.fakturaEpost || k.epost].filter(Boolean).join(" · "))),
            el("div", { class: "kunde-handling" },
              el("button", { type: "button", class: "knapp tekst", "aria-label": "Ny sak for " + k.navn, onclick: function () { sakDialog(null, k.id); } }, ikon("pluss", "ikon-liten"), "Ny sak"),
              el("button", { type: "button", class: "knapp tekst", "aria-label": "Rediger kunde " + k.navn, onclick: function () { kundeDialog(k); } }, ikon("rediger", "ikon-liten"), "Rediger kunde"))),
          vis.length ? el("ul", { class: "sakliste" }, vis.map(function (s) {
            var u = uf[s.id];
            return el("li", { class: "sak" + (s.status === "avsluttet" ? " avsluttet" : "") },
              el("button", { type: "button", class: "sak-knapp", onclick: function () { sakDialog(s); }, "aria-label": "Rediger sak " + s.saksnr + " " + s.navn },
                el("span", { class: "sak-nr" }, s.saksnr),
                el("span", { class: "sak-navn" }, s.navn, s.motpart ? el("small", null, "Motpart: " + s.motpart) : null),
                el("span", { class: "sak-pris" }, krHel(s.timepris) + "/t"),
                el("span", { class: "sak-uf" }, u ? el("span", null, t(u.t), el("small", null, krHel(u.b))) : el("small", null, "Ingenting ufakturert")),
                el("span", { class: "sak-status" }, s.status === "avsluttet" ? chip("Avsluttet", "noytral") : null)),
              s.status !== "avsluttet" ? el("button", { type: "button", class: "knapp sekundaer liten sak-for", onclick: function () { hurtigforing({ sakId: s.id }); }, "aria-label": "Før timer på " + s.saksnr }, ikon("pluss", "ikon-liten"), "Før") : null);
          })) : el("p", { class: "kunde-tom" }, "Ingen aktive saker.")));
      });
      if (!blokker.length) {
        erstatt(listeOmr, q ? tomTilstand("Ingen treff på «" + sakerFilter.sok + "».") :
          tomTilstand("Ingen kunder ennå. Start med en oppdragsbekreftelse – da opprettes kunde og sak samtidig.", "Ny oppdragsbekreftelse", function () { gaaTil("oppdrag"); }));
        return;
      }
      erstatt(listeOmr, blokker);
    }

    node.appendChild(el("div", { class: "verktoy" },
      sokF, avsl,
      el("div", { class: "knapperad" },
        el("button", { type: "button", class: "knapp sekundaer", onclick: function () { kundeDialog(null); } }, "Ny kunde"),
        el("button", { type: "button", class: "knapp", onclick: function () { sakDialog(null); } }, "Ny sak"))));
    node.appendChild(listeOmr);
    tegn();
    return { node: node, oppdater: tegn };
  }

  // ---------- Brønnøysund-boks (gjenbrukes i kunde-dialog og oppdrag) ----------
  /**
   * Lager en statusboks for Brreg. o.hentFelt() -> { orgnr, navn }, o.fyll(enhet, fraForslag).
   * Returnerer { node, slaaOpp(auto) }.
   */
  function brregBoks(o) {
    var node = el("div", { class: "brreg", "aria-live": "polite" });
    var teller = 0;
    function vis() { erstatt.apply(null, [node].concat(Array.prototype.slice.call(arguments))); node.hidden = false; }
    function enhetVisning(e) {
      return el("div", { class: "brreg-enhet" + (e.advarsler.length ? " har-advarsel" : "") },
        el("p", { class: "brreg-tittel" }, ikon(e.advarsler.length ? "advarsel" : "sok", "ikon-liten"), "Brønnøysundregistrene"),
        el("p", { class: "brreg-navn" }, e.navn || "(uten navn)", e.orgform ? el("span", { class: "chip chip-noytral" }, e.orgform) : null, e.kilde === "underenhet" ? el("span", { class: "chip chip-noytral" }, "Underenhet") : null),
        el("p", { class: "brreg-detalj" }, "Org.nr. " + Brreg.formaterOrgnr(e.orgnr) + (e.fakturaadresse ? " · " + e.fakturaadresse : "")),
        e.advarsler.length ? el("ul", { class: "brreg-advarsler" }, e.advarsler.map(function (a) { return el("li", null, a); })) : null);
    }
    async function slaaOpp(auto) {
      var f = o.hentFelt();
      var nr = Brreg.vask(f.orgnr);
      var mitt = ++teller;
      if (nr) {
        if (!Brreg.gyldigOrgnr(nr)) {
          vis(el("p", { class: "brreg-feil" }, ikon("advarsel", "ikon-liten"), "Org.nr. " + f.orgnr + " er ikke gyldig – sjekk sifrene."));
          return;
        }
        vis(el("p", { class: "brreg-laster" }, "Slår opp org.nr. " + Brreg.formaterOrgnr(nr) + " i Brønnøysundregistrene …"));
        try {
          var e = await Brreg.oppslag(nr);
          if (mitt !== teller) return;
          if (!e) { vis(el("p", { class: "brreg-feil" }, "Fant ikke org.nr. " + Brreg.formaterOrgnr(nr) + " i Enhetsregisteret.")); return; }
          vis(enhetVisning(e));
          o.fyll(e, false);
        } catch (err) {
          if (mitt !== teller) return;
          vis(el("p", { class: "brreg-feil" }, ikon("advarsel", "ikon-liten"), feilTekst(err)));
        }
        return;
      }
      var navn = String(f.navn || "").replace(/\s*\((demo|eksempel)\)\s*$/i, "").trim();
      if (navn.length < 2) { if (!auto) vis(el("p", { class: "brreg-feil" }, "Skriv org.nr. eller firmanavn først.")); return; }
      vis(el("p", { class: "brreg-laster" }, "Søker etter «" + navn + "» i Brønnøysundregistrene …"));
      try {
        var treff = await Brreg.sok(navn, 5);
        if (mitt !== teller) return;
        if (!treff.length) { vis(el("p", { class: "brreg-feil" }, "Ingen treff på «" + navn + "». Fyll inn org.nr. manuelt om kunden er en virksomhet.")); return; }
        vis(el("p", { class: "brreg-tittel" }, ikon("sok", "ikon-liten"), "Velg riktig enhet fra Brønnøysundregistrene"),
          el("ul", { class: "brreg-forslag" }, treff.map(function (e) {
            return el("li", null, el("button", { type: "button", class: "brreg-valg", onclick: function () { vis(enhetVisning(e)); o.fyll(e, true); } },
              el("strong", null, e.navn), el("span", null, Brreg.formaterOrgnr(e.orgnr) + (e.orgform ? " · " + e.orgform : "") + (e.fakturaadresse ? " · " + e.fakturaadresse : "")),
              e.advarsler.length ? el("span", { class: "brreg-varsel" }, e.advarsler.join(" ")) : null));
          })));
      } catch (err2) {
        if (mitt !== teller) return;
        vis(el("p", { class: "brreg-feil" }, ikon("advarsel", "ikon-liten"), feilTekst(err2)));
      }
    }
    node.hidden = true;
    return { node: node, slaaOpp: slaaOpp, nullstill: function () { teller++; tom(node); node.hidden = true; } };
  }

  function penNavn(n) {
    if (window.Oppdrag && typeof window.Oppdrag.normaliserNavn === "function") return window.Oppdrag.normaliserNavn(n);
    return n;
  }

  // ---------- Kunde-dialog ----------
  function kundeDialog(k) {
    var ny = !k;
    k = k || { kundenr: Store.nesteKundenr() };
    var f = {
      navn: felt({ etikett: "Navn", verdi: k.navn, paakrevd: true }),
      orgnr: felt({ etikett: "Org.nr.", verdi: k.orgnr, inputmode: "numeric", placeholder: "9 siffer" }),
      kundenr: felt({ etikett: "Kundenr.", verdi: k.kundenr, inputmode: "numeric", paakrevd: true }),
      kontaktperson: felt({ etikett: "Kontaktperson", verdi: k.kontaktperson }),
      epost: felt({ etikett: "E-post", type: "email", verdi: k.epost }),
      telefon: felt({ etikett: "Telefon", type: "tel", verdi: k.telefon }),
      fakturaadresse: felt({ etikett: "Fakturaadresse", type: "textarea", rader: 2, verdi: k.fakturaadresse }),
      fakturaEpost: felt({ etikett: "Faktura e-post", type: "email", verdi: k.fakturaEpost }),
      fakturaMerkes: felt({ etikett: "Faktura merkes", verdi: k.fakturaMerkes })
    };
    var bb = brregBoks({
      hentFelt: function () { return { orgnr: f.orgnr.input.value, navn: f.navn.input.value }; },
      fyll: function (e, valgt) {
        if (e.orgnr) f.orgnr.input.value = e.orgnr;
        if (e.navn && (valgt || !f.navn.input.value.trim() || true)) f.navn.input.value = penNavn(e.navn);
        if (e.fakturaadresse && (valgt || !f.fakturaadresse.input.value.trim())) f.fakturaadresse.input.value = e.fakturaadresse;
      }
    });
    var hent = el("button", { type: "button", class: "knapp sekundaer liten", onclick: function () { bb.slaaOpp(false); } }, "Hent fra Brønnøysund");
    var form = el("form", { class: "skjema", novalidate: true },
      el("div", { class: "grid-2" }, f.navn, el("div", { class: "felt-med-knapp" }, f.orgnr, hent)),
      bb.node,
      el("div", { class: "grid-3" }, f.kundenr, f.kontaktperson, f.telefon),
      el("div", { class: "grid-2" }, f.epost, f.fakturaEpost),
      el("div", { class: "grid-2" }, f.fakturaadresse, f.fakturaMerkes));
    var lagre = el("button", { type: "submit", class: "knapp" }, ny ? "Opprett kunde" : "Lagre kunde");
    form.appendChild(el("div", { class: "knapperad" }, lagre));
    var h = aapneDialog({ tittel: ny ? "Ny kunde" : "Rediger kunde", innhold: form, klasse: "bred", fokus: "input" });
    form.addEventListener("submit", async function (e) {
      e.preventDefault();
      Object.keys(f).forEach(function (n) { f[n].settFeil(""); });
      var v = {};
      Object.keys(f).forEach(function (n) { v[n] = f[n].input.value.trim(); });
      if (!v.navn) { f.navn.settFeil("Skriv navnet på kunden."); f.navn.input.focus(); return; }
      if (v.orgnr && !Brreg.gyldigOrgnr(v.orgnr)) { f.orgnr.settFeil("Org.nr. er ikke gyldig."); f.orgnr.input.focus(); return; }
      if (Store.kundenrFinnes(v.kundenr, ny ? null : k.id)) { f.kundenr.settFeil("Kundenr. er allerede i bruk."); f.kundenr.input.focus(); return; }
      if (ny) {
        var finnes = Store.finnKunde({ orgnr: v.orgnr, navn: v.navn });
        if (finnes) { f[finnes.grunn === "orgnr" ? "orgnr" : "navn"].settFeil("Kunden finnes allerede (kundenr. " + finnes.kunde.kundenr + ")."); return; }
      }
      lagre.disabled = true;
      try {
        if (ny) await Store.opprettKunde(v); else await Store.endreKunde(k.id, v);
        h.lukk();
        toast(ny ? "Kunden er opprettet." : "Kunden er lagret.", "ok");
      } catch (err) { toast(feilTekst(err), "feil"); }
      finally { lagre.disabled = false; }
    });
  }

  // ---------- Sak-dialog ----------
  function sakDialog(s, kundeId) {
    var ny = !s;
    var kunder = Store.data().kunder.slice().sort(function (a, b) { return String(a.navn).localeCompare(String(b.navn), "nb"); });
    if (!kunder.length) { toast("Opprett en kunde først.", "info"); kundeDialog(null); return; }
    s = s || { saksnr: Store.nesteSaksnr(), kundeId: kundeId || "", ansvarligAdvokat: "Geir Frøholm", saksbehandler: "Geir Frøholm", timepris: "", status: "aktiv" };
    var f = {
      kundeId: felt({ etikett: "Kunde", type: "select", verdi: s.kundeId, paakrevd: true, valg: [{ verdi: "", tekst: "Velg kunde" }].concat(kunder.map(function (k) { return { verdi: k.id, tekst: k.kundenr + " · " + k.navn }; })) }),
      saksnr: felt({ etikett: "Saksnr.", verdi: s.saksnr, paakrevd: true, hint: ny ? "Neste ledige er foreslått." : null }),
      navn: felt({ etikett: "Sakens navn", verdi: s.navn, paakrevd: true }),
      timepris: felt({ etikett: "Timepris (kr eks. mva)", verdi: s.timepris, inputmode: "decimal", paakrevd: true }),
      motpart: felt({ etikett: "Motpart", verdi: s.motpart }),
      ansvarligAdvokat: felt({ etikett: "Ansvarlig advokat", verdi: s.ansvarligAdvokat }),
      saksbehandler: felt({ etikett: "Saksbehandler", verdi: s.saksbehandler }),
      faktureringsrutiner: felt({ etikett: "Faktureringsrutiner", verdi: s.faktureringsrutiner }),
      betalingsbetingelser: felt({ etikett: "Betalingsbetingelser", verdi: s.betalingsbetingelser }),
      salaerberegning: felt({ etikett: "Salærberegning", type: "textarea", rader: 3, verdi: s.salaerberegning }),
      status: felt({ etikett: "Status", type: "select", verdi: s.status || "aktiv", valg: [{ verdi: "aktiv", tekst: "Aktiv" }, { verdi: "avsluttet", tekst: "Avsluttet" }] })
    };
    var sakPrisVarsel = el("p", { class: "felt-varsel", "aria-live": "polite" });
    f.timepris.appendChild(sakPrisVarsel);
    function sjekkSakPris() { sakPrisVarsel.textContent = prisUvanlig(tolkBelop(f.timepris.input.value)) ? prisVarselTekst() : ""; }
    f.timepris.input.addEventListener("input", sjekkSakPris);
    sjekkSakPris();
    var form = el("form", { class: "skjema", novalidate: true },
      el("div", { class: "grid-2" }, f.kundeId, f.saksnr),
      f.navn,
      el("div", { class: "grid-2" }, f.timepris, f.motpart),
      el("div", { class: "grid-2" }, f.ansvarligAdvokat, f.saksbehandler),
      el("div", { class: "grid-2" }, f.faktureringsrutiner, f.betalingsbetingelser),
      f.salaerberegning,
      ny ? null : f.status);
    var lagre = el("button", { type: "submit", class: "knapp" }, ny ? "Opprett sak" : "Lagre sak");
    form.appendChild(el("div", { class: "knapperad" }, lagre));
    var h = aapneDialog({ tittel: ny ? "Ny sak" : "Rediger sak " + s.saksnr, innhold: form, klasse: "bred" });
    (ny && !kundeId ? f.kundeId.input : f.navn.input).focus();
    form.addEventListener("submit", async function (e) {
      e.preventDefault();
      Object.keys(f).forEach(function (n) { f[n].settFeil(""); });
      var v = {};
      Object.keys(f).forEach(function (n) { v[n] = f[n].input.value.trim(); });
      var feil = false;
      if (!v.kundeId) { f.kundeId.settFeil("Velg kunde."); feil = true; }
      if (!/^\d+\/\d{4}$/.test(v.saksnr.replace(/\s/g, ""))) { f.saksnr.settFeil("Skriv saksnr. som N/ÅÅÅÅ, f.eks. " + Store.nesteSaksnr() + "."); feil = true; }
      else if (Store.saksnrFinnes(v.saksnr, ny ? null : s.id)) { f.saksnr.settFeil("Saksnr. finnes allerede. Neste ledige er " + Store.nesteSaksnr() + "."); feil = true; }
      if (!v.navn) { f.navn.settFeil("Skriv sakens navn."); feil = true; }
      var pris = tolkBelop(v.timepris);
      if (!isFinite(pris) || pris <= 0) { f.timepris.settFeil("Fyll inn timepris, f.eks. 2 690."); feil = true; }
      if (feil) { var x = form.querySelector("[aria-invalid=true]"); if (x) x.focus(); return; }
      if (prisUvanlig(pris) && (ny || Number(s.timepris) !== pris)) {
        var ok = await bekreft({ tittel: "Uvanlig timepris", tekst: "Timeprisen er " + krHel(pris) + ". " + prisVarselTekst() + " Er det riktig?", ok: "Ja, lagre", avbryt: "Rett timeprisen" });
        if (!ok) { f.timepris.input.focus(); f.timepris.input.select(); return; }
      }
      v.timepris = pris;
      lagre.disabled = true;
      try {
        if (ny) { v.kilde = "manuell"; await Store.opprettSak(v); } else await Store.endreSak(s.id, v);
        h.lukk();
        toast(ny ? "Sak " + v.saksnr + " er opprettet." : "Saken er lagret.", "ok");
      } catch (err) { toast(feilTekst(err), "feil"); }
      finally { lagre.disabled = false; }
    });
  }

  // =====================================================================
  // Seksjon: Ny oppdragsbekreftelse
  // =====================================================================
  var oppdragTekst = "";

  function visOppdrag() {
    var node = el("div", { class: "seksjon oppdrag" });
    var skjemaOmr = el("div", { class: "oppdrag-skjema" });

    var tekstF = felt({ etikett: "Lim inn teksten fra oppdragsbekreftelsen", type: "textarea", rader: 9, verdi: oppdragTekst, placeholder: "Merk alt i Word-dokumentet (Ctrl/Cmd + A), kopier og lim inn her." });
    tekstF.input.spellcheck = false;
    tekstF.input.addEventListener("input", function () { oppdragTekst = tekstF.input.value; });
    tekstF.input.addEventListener("paste", function () { setTimeout(function () { oppdragTekst = tekstF.input.value; tolk(); }, 0); });

    var filInput = el("input", { type: "file", id: uid("fil"), accept: ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document", class: "skjult-fil" });
    var slipp = el("label", { class: "slipp", for: filInput.id },
      ikon("oppdrag", "slipp-ikon"),
      el("span", { class: "slipp-hoved" }, "Slipp Word-fila her"),
      el("span", { class: "slipp-under" }, "eller klikk for å velge en .docx"));
    filInput.addEventListener("change", function () { if (filInput.files[0]) lesFil(filInput.files[0]); filInput.value = ""; });
    ["dragenter", "dragover"].forEach(function (h) { slipp.addEventListener(h, function (e) { e.preventDefault(); slipp.classList.add("over"); }); });
    ["dragleave", "drop"].forEach(function (h) { slipp.addEventListener(h, function (e) { e.preventDefault(); slipp.classList.remove("over"); }); });
    slipp.addEventListener("drop", function (e) { var fl = e.dataTransfer && e.dataTransfer.files[0]; if (fl) lesFil(fl); });

    var tolkKnapp = el("button", { type: "button", class: "knapp", onclick: function () { tolk(); } }, "Tolk teksten");

    async function lesFil(fil) {
      if (!/\.docx$/i.test(fil.name)) { toast("Velg en Word-fil (.docx). Eldre .doc-filer må lagres som .docx først.", "feil"); return; }
      if (fil.size > 15 * 1024 * 1024) { toast("Fila er for stor (maks 15 MB).", "feil"); return; }
      erstatt(skjemaOmr, el("p", { class: "laster" }, "Leser " + fil.name + " …"));
      try {
        var buf = await fil.arrayBuffer();
        var res = await window.Oppdrag.fraDocx(buf);
        byggSkjema(res, fil.name);
      } catch (e) {
        console.error(e);
        erstatt(skjemaOmr, el("p", { class: "melding feil" }, "Klarte ikke å lese fila: " + feilTekst(e) + " Prøv å lime inn teksten i stedet."));
      }
    }

    function tolk() {
      var tekst = tekstF.input.value;
      if (!tekst.trim()) { tekstF.settFeil("Lim inn teksten først."); tekstF.input.focus(); return; }
      tekstF.settFeil("");
      if (!window.Oppdrag) { toast("Tolkemodulen er ikke lastet.", "feil"); return; }
      try { byggSkjema(window.Oppdrag.fraTekst(tekst), null); }
      catch (e) { console.error(e); toast("Kunne ikke tolke teksten: " + feilTekst(e), "feil"); }
    }

    function byggSkjema(res, filnavn) {
      var k = res.kunde || {}, s = res.sak || {};
      var advarsler = (res.advarsler || []).slice();
      var forslag = Array.isArray(res.forslagTimepris) ? res.forslagTimepris.filter(function (x) { return isFinite(Number(x)); }) : [];
      var match = Store.finnKunde({ orgnr: k.orgnr, navn: k.navn });
      var brukEksisterende = !!match;

      // Saksnummer: bruk fra dokumentet om ledig, ellers neste ledige
      var aar = String(new Date().getFullYear());
      var saksnr = (s.saksnr || "").replace(/\s/g, "");
      var saksnrMerknad = "Neste ledige saksnr. er foreslått.";
      if (saksnr && /^\d+\/\d{4}$/.test(saksnr) && !Store.saksnrFinnes(saksnr)) saksnrMerknad = "Hentet fra oppdragsbekreftelsen.";
      else {
        if (saksnr && Store.saksnrFinnes(saksnr)) advarsler.push("Saksnr. " + saksnr + " i dokumentet er allerede i bruk – " + Store.nesteSaksnr(aar) + " er foreslått i stedet.");
        saksnr = Store.nesteSaksnr(aar);
      }

      var f = {
        navn: felt({ etikett: "Klient", verdi: k.navn, paakrevd: true }),
        orgnr: felt({ etikett: "Org.nr.", verdi: k.orgnr, inputmode: "numeric", placeholder: "9 siffer" }),
        kundenr: felt({ etikett: "Kundenr.", verdi: Store.nesteKundenr(), inputmode: "numeric", hint: "Tildeles automatisk." }),
        kontaktperson: felt({ etikett: "Kontaktperson", verdi: k.kontaktperson }),
        epost: felt({ etikett: "E-post", type: "email", verdi: k.epost }),
        telefon: felt({ etikett: "Telefon", type: "tel", verdi: k.telefon }),
        fakturaadresse: felt({ etikett: "Fakturaadresse", type: "textarea", rader: 2, verdi: k.fakturaadresse }),
        fakturaEpost: felt({ etikett: "Faktura e-post", type: "email", verdi: k.fakturaEpost }),
        fakturaMerkes: felt({ etikett: "Faktura merkes", verdi: k.fakturaMerkes })
      };
      var sf = {
        saksnr: felt({ etikett: "Saksnr.", verdi: saksnr, paakrevd: true, hint: saksnrMerknad }),
        navn: felt({ etikett: "Sakens navn", verdi: s.navn, paakrevd: true }),
        timepris: felt({ etikett: "Timepris (kr eks. mva)", verdi: s.timepris != null ? String(s.timepris) : "", inputmode: "decimal", paakrevd: true }),
        motpart: felt({ etikett: "Motpart", verdi: s.motpart }),
        ansvarligAdvokat: felt({ etikett: "Ansvarlig advokat", verdi: s.ansvarligAdvokat || "Geir Frøholm" }),
        saksbehandler: felt({ etikett: "Saksbehandler", verdi: s.saksbehandler || "Geir Frøholm" }),
        faktureringsrutiner: felt({ etikett: "Faktureringsrutiner", verdi: s.faktureringsrutiner }),
        betalingsbetingelser: felt({ etikett: "Betalingsbetingelser", verdi: s.betalingsbetingelser }),
        salaerberegning: felt({ etikett: "Salærberegning", type: "textarea", rader: 3, verdi: s.salaerberegning })
      };

      // Timepris-sjekk live
      var prisVarsel = el("p", { class: "felt-varsel", "aria-live": "polite" });
      function sjekkPris() {
        var p = tolkBelop(sf.timepris.input.value);
        prisVarsel.textContent = prisUvanlig(p) ? prisVarselTekst() : "";
      }
      sf.timepris.input.addEventListener("input", sjekkPris);
      sf.timepris.appendChild(prisVarsel);
      sjekkPris();

      // Eksisterende kunde
      var kundeValg = el("div", { class: "kundevalg" });
      var kundeFelter = el("div", { class: "kunde-felter" });
      function tegnKundevalg() {
        var m = Store.finnKunde({ orgnr: f.orgnr.input.value, navn: f.navn.input.value });
        if (m && (!match || match.kunde.id !== m.kunde.id)) { match = m; brukEksisterende = true; }
        if (!m && match && !brukEksisterende) match = null;
        tom(kundeValg);
        if (!match) { kundeFelter.hidden = false; kundeValg.hidden = true; return; }
        kundeValg.hidden = false;
        var rid = uid("kv");
        kundeValg.appendChild(el("p", { class: "kundevalg-tittel" }, ikon("saker", "ikon-liten"),
          "Kunden finnes allerede: ", el("strong", null, match.kunde.navn), " (kundenr. " + match.kunde.kundenr + ", gjenkjent på " + (match.grunn === "orgnr" ? "org.nr." : "navn") + ")"));
        var r1 = el("input", { type: "radio", name: rid, id: rid + "a", checked: brukEksisterende });
        var r2 = el("input", { type: "radio", name: rid, id: rid + "b", checked: !brukEksisterende });
        r1.addEventListener("change", function () { brukEksisterende = true; oppdaterKunde(); });
        r2.addEventListener("change", function () { brukEksisterende = false; oppdaterKunde(); });
        kundeValg.appendChild(el("div", { class: "radio-gruppe", role: "radiogroup", "aria-label": "Kunde for saken" },
          el("div", { class: "radio" }, r1, el("label", { for: rid + "a" }, "Knytt saken til eksisterende kunde")),
          el("div", { class: "radio" }, r2, el("label", { for: rid + "b" }, "Opprett ny kunde likevel"))));
        oppdaterKunde();
      }
      function oppdaterKunde() {
        kundeFelter.hidden = !!(match && brukEksisterende);
        lagreKnapp.textContent = match && brukEksisterende ? "Opprett sak hos " + match.kunde.navn : "Opprett kunde og sak";
      }

      var bb = brregBoks({
        hentFelt: function () { return { orgnr: f.orgnr.input.value, navn: f.navn.input.value }; },
        fyll: function (e, valgt) {
          if (e.orgnr) f.orgnr.input.value = e.orgnr;
          if (e.navn) f.navn.input.value = penNavn(e.navn);
          if (e.fakturaadresse && (valgt || !f.fakturaadresse.input.value.trim())) f.fakturaadresse.input.value = e.fakturaadresse;
          tegnKundevalg();
        }
      });
      var hent = el("button", { type: "button", class: "knapp sekundaer liten", onclick: function () { bb.slaaOpp(false); } }, "Hent fra Brønnøysund");
      f.orgnr.input.addEventListener("change", function () { tegnKundevalg(); if (Brreg.gyldigOrgnr(f.orgnr.input.value)) bb.slaaOpp(true); });
      f.navn.input.addEventListener("change", tegnKundevalg);

      kundeFelter.appendChild(el("div", { class: "grid-3" }, f.kundenr, f.kontaktperson, f.telefon));
      kundeFelter.appendChild(el("div", { class: "grid-2" }, f.epost, f.fakturaEpost));
      kundeFelter.appendChild(el("div", { class: "grid-2" }, f.fakturaadresse, f.fakturaMerkes));

      var lagreKnapp = el("button", { type: "submit", class: "knapp" }, "Opprett kunde og sak");
      var avbryt = el("button", { type: "button", class: "knapp tekst", onclick: function () { tom(skjemaOmr); tekstF.input.focus(); } }, "Forkast");

      var form = el("form", { class: "skjema oppdrag-form", novalidate: true },
        el("div", { class: "oppdrag-kilde" }, etikett("Tolket fra " + (filnavn ? "«" + filnavn + "»" : "innlimt tekst"))),
        advarsler.length ? el("div", { class: "melding advarsel", role: "status" },
          el("p", { class: "melding-tittel" }, ikon("advarsel", "ikon-liten"), advarsler.length === 1 ? "Sjekk dette før du lagrer" : "Sjekk disse " + advarsler.length + " punktene før du lagrer"),
          el("ul", null, advarsler.map(function (a) { return el("li", null, a); })),
          forslag.length ? el("div", { class: "forslag" }, el("span", null, "Mente du:"),
            forslag.map(function (x) {
              return el("button", { type: "button", class: "knapp sekundaer liten", onclick: function () {
                sf.timepris.input.value = String(x); sjekkPris(); sf.timepris.input.focus();
                toast("Timepris satt til " + krHel(x) + ".", "ok");
              } }, "Bruk " + krHel(x));
            })) : null) : null,
        el("fieldset", { class: "gruppe" }, el("legend", null, "Kunde"),
          el("div", { class: "grid-2" }, f.navn, el("div", { class: "felt-med-knapp" }, f.orgnr, hent)),
          bb.node, kundeValg, kundeFelter),
        el("fieldset", { class: "gruppe" }, el("legend", null, "Sak"),
          el("div", { class: "grid-3" }, sf.saksnr, sf.timepris, sf.motpart),
          sf.navn,
          el("div", { class: "grid-2" }, sf.ansvarligAdvokat, sf.saksbehandler),
          el("div", { class: "grid-2" }, sf.faktureringsrutiner, sf.betalingsbetingelser),
          sf.salaerberegning),
        el("div", { class: "knapperad" }, lagreKnapp, avbryt));

      form.addEventListener("submit", async function (e) {
        e.preventDefault();
        [f, sf].forEach(function (g) { Object.keys(g).forEach(function (n) { g[n].settFeil(""); }); });
        var kv = {}, sv = {};
        Object.keys(f).forEach(function (n) { kv[n] = f[n].input.value.trim(); });
        Object.keys(sf).forEach(function (n) { sv[n] = sf[n].input.value.trim(); });
        var feil = [];
        var eks = match && brukEksisterende ? match.kunde.id : null;
        if (!eks) {
          if (!kv.navn) { f.navn.settFeil("Skriv klientens navn."); feil.push(f.navn); }
          if (kv.orgnr && !Brreg.gyldigOrgnr(kv.orgnr)) { f.orgnr.settFeil("Org.nr. er ikke gyldig."); feil.push(f.orgnr); }
          if (!(parseInt(kv.kundenr, 10) > 0)) { f.kundenr.settFeil("Skriv et kundenr."); feil.push(f.kundenr); }
          else if (Store.kundenrFinnes(kv.kundenr)) { f.kundenr.settFeil("Kundenr. er i bruk. Neste ledige er " + Store.nesteKundenr() + "."); feil.push(f.kundenr); }
        }
        if (!/^\d+\/\d{4}$/.test(sv.saksnr.replace(/\s/g, ""))) { sf.saksnr.settFeil("Skriv saksnr. som N/ÅÅÅÅ."); feil.push(sf.saksnr); }
        else if (Store.saksnrFinnes(sv.saksnr)) { sf.saksnr.settFeil("Saksnr. finnes allerede. Neste ledige er " + Store.nesteSaksnr() + "."); feil.push(sf.saksnr); }
        if (!sv.navn) { sf.navn.settFeil("Skriv sakens navn."); feil.push(sf.navn); }
        var pris = tolkBelop(sv.timepris);
        if (!isFinite(pris)) { sf.timepris.settFeil("Skriv timepris, f.eks. 2 690."); feil.push(sf.timepris); }
        if (feil.length) { feil[0].input.focus(); return; }
        if (prisUvanlig(pris)) {
          var fortsett = await bekreft({ tittel: "Uvanlig timepris", tekst: "Timeprisen er " + krHel(pris) + ". " + prisVarselTekst() + " Er det riktig?", ok: "Ja, lagre", avbryt: "Rett timeprisen" });
          if (!fortsett) { sf.timepris.input.focus(); sf.timepris.input.select(); return; }
        }
        sv.timepris = pris;
        sv.kilde = "oppdragsbekreftelse";
        lagreKnapp.disabled = true;
        try {
          var r = await Store.opprettKundeOgSak(kv, sv, eks);
          oppdragTekst = "";
          tekstF.input.value = "";
          erstatt(skjemaOmr, el("div", { class: "melding ok kvittering", role: "status" },
            el("p", { class: "melding-tittel" }, eks ? "Saken er opprettet hos eksisterende kunde." : "Kunden og saken er opprettet."),
            el("p", null, el("strong", null, r.sak.saksnr + " " + r.sak.navn), " · " + r.kunde.navn + " (kundenr. " + r.kunde.kundenr + ") · " + krHel(r.sak.timepris) + "/t"),
            el("div", { class: "knapperad" },
              el("button", { type: "button", class: "knapp", onclick: function () { hurtigforing({ sakId: r.sak.id }); } }, "Før timer på saken"),
              el("button", { type: "button", class: "knapp tekst", onclick: function () { gaaTil("saker"); } }, "Se saker og kunder"))));
          toast("Sak " + r.sak.saksnr + " er opprettet.", "ok");
        } catch (err) {
          toast(feilTekst(err), "feil");
        } finally { lagreKnapp.disabled = false; }
      });

      erstatt(skjemaOmr, form);
      tegnKundevalg();
      // Automatisk oppslag i Brreg
      if ((k.orgnr || k.navn) && !(match && brukEksisterende && match.kunde.orgnr)) bb.slaaOpp(true);
      var forsteTomme = [f.navn, sf.navn, sf.timepris].find(function (x) { return !x.input.value.trim(); });
      form.scrollIntoView({ behavior: "smooth", block: "start" });
      (forsteTomme || sf.navn).input.focus({ preventScroll: true });
    }

    node.appendChild(el("section", { class: "panel oppdrag-inn" },
      el("div", { class: "panel-hode" }, el("h2", null, "Hent opplysninger fra oppdragsbekreftelsen")),
      el("p", { class: "ingress-liten" }, "Feltene tolkes automatisk og kan rettes før kunden og saken opprettes. Org.nr. og firmanavn slås opp i Brønnøysundregistrene."),
      el("div", { class: "oppdrag-inndata" },
        el("div", { class: "oppdrag-tekst" }, tekstF, el("div", { class: "knapperad" }, tolkKnapp)),
        el("div", { class: "oppdrag-fil" }, filInput, slipp))));
    node.appendChild(skjemaOmr);
    return { node: node, oppdater: function () { /* skjemaet holdes urørt mens det redigeres */ } };
  }

  // =====================================================================
  // Seksjon: Fakturering
  // =====================================================================
  function lastNed(buffer, navn) {
    var blob = new Blob([buffer], { type: Store.XLSX_TYPE });
    var url = URL.createObjectURL(blob);
    var a = el("a", { href: url, download: navn, class: "skjult-fil" });
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1000);
  }

  function grunnlagTabell(g) {
    var hode = ["Klient", "Sak", "Timer", "Timepris", "Salær", "Utlegg", "Grunnlag", "MVA", "Totalt"];
    return el("div", { class: "tabell-boks" }, el("table", { class: "tabell" },
      el("thead", null, el("tr", null, hode.map(function (h, i) { return el("th", { scope: "col", class: i >= 2 ? "tall" : null }, h); }))),
      el("tbody", null, g.perSak.map(function (r) {
        return el("tr", null,
          el("td", { "data-etikett": "Klient" }, r.kunde ? r.kunde.navn : ""),
          el("td", { "data-etikett": "Sak" }, el("span", { class: "td-nr" }, r.sak ? r.sak.saksnr : ""), " ", r.sak ? r.sak.navn : ""),
          el("td", { class: "tall", "data-etikett": "Timer" }, tTall(r.timer)),
          el("td", { class: "tall", "data-etikett": "Timepris" }, krHel(r.timepris) + (r.blandetSats ? " (snitt)" : "")),
          el("td", { class: "tall", "data-etikett": "Salær" }, kr(r.salaer)),
          el("td", { class: "tall", "data-etikett": "Utlegg" }, r.utlegg ? kr(r.utlegg) : "–"),
          el("td", { class: "tall", "data-etikett": "Grunnlag" }, kr(r.grunnlag)),
          el("td", { class: "tall", "data-etikett": "MVA" }, kr(r.mva)),
          el("td", { class: "tall sterk", "data-etikett": "Totalt" }, kr(r.total)));
      })),
      el("tfoot", null, el("tr", null,
        el("th", { scope: "row", colspan: "2" }, "Sum"),
        el("td", { class: "tall", "data-etikett": "Timer" }, tTall(g.sum.timer)),
        el("td", { class: "tall tom-celle" }, ""),
        el("td", { class: "tall", "data-etikett": "Salær" }, kr(g.sum.salaer)),
        el("td", { class: "tall", "data-etikett": "Utlegg" }, g.sum.utlegg ? kr(g.sum.utlegg) : "–"),
        el("td", { class: "tall", "data-etikett": "Grunnlag" }, kr(g.sum.grunnlag)),
        el("td", { class: "tall", "data-etikett": "MVA" }, kr(g.sum.mva)),
        el("td", { class: "tall sterk", "data-etikett": "Totalt" }, kr(g.sum.total))))));
  }

  function visFaktura() {
    var node = el("div", { class: "seksjon faktura" });
    var arbeider = false;

    async function lastNedLopende() {
      try {
        var buf = await Store.lagXlsx();
        lastNed(buf, "Faktureringsgrunnlag " + Store.idag() + ".xlsx");
      } catch (e) { toast("Kunne ikke lage Excel-fila: " + feilTekst(e), "feil"); }
    }

    async function sendNaa(g) {
      var i = Store.innst();
      if (!i.mottakerEpost) { toast("Legg inn e-postadressen til View under Innstillinger først.", "feil"); gaaTil("innst"); return; }
      var ok = await bekreft({
        tittel: "Sende faktureringsgrunnlaget nå?",
        ok: "Send nå",
        tekst: el("div", { class: "send-oppsummering" },
          el("dl", { class: "dl" },
            el("dt", null, "Til"), el("dd", null, i.mottakerNavn + " <" + i.mottakerEpost + ">"),
            i.kopiEpost ? [el("dt", null, "Kopi"), el("dd", null, i.kopiEpost)] : null,
            el("dt", null, "Emne"), el("dd", null, i.emne),
            el("dt", null, "Vedlegg"), el("dd", null, "Faktureringsgrunnlag " + Store.idag() + ".xlsx"),
            el("dt", null, "Innhold"), el("dd", null, g.antallLinjer + " linjer · " + g.perSak.length + " saker · " + kr(g.sum.total) + " inkl. mva")),
          el("p", null, "Linjene låses etter sending. Den faste automatiske sendingen " + sendeTidTekst(Store.nesteSending()) + " går som før, med linjer som føres etter dette."),
          erDemo() ? el("p", { class: "melding advarsel" }, "Demo: ingen e-post sendes, sendingen simuleres.") : null)
      });
      if (!ok) return;
      arbeider = true; oppdater();
      try {
        var r = await Store.sendNaa();
        toast((erDemo() ? "Demo: sending simulert. " : "Sendt til " + i.mottakerEpost + ". ") + "Arkivert som «" + r.navn + "».", "ok");
      } catch (e) {
        toast("Sendingen feilet: " + feilTekst(e), "feil");
      } finally { arbeider = false; oppdater(); }
    }

    async function lastNedBatch(b) {
      try {
        var buf = await Store.lagXlsx(Store.linjerForBatch(b.id));
        var navn = (b.fil || "").split("/").pop() || ("Faktureringsgrunnlag " + b.sendt.slice(0, 10) + ".xlsx");
        lastNed(buf, navn);
      } catch (e) { toast("Kunne ikke lage Excel-fila: " + feilTekst(e), "feil"); }
    }

    function oppdater() {
      var i = Store.innst();
      var g = null, gFeil = null;
      try { g = Store.grunnlag(); } catch (e) { gFeil = e; }
      var fs = Store.filStatus();
      var harLinjer = g && g.antallLinjer > 0;

      var filTekst = erDemo()
        ? (fs.finnes ? "Løpende fil (simulert i demo) oppdatert " + (fs.oppdatert ? dato(fs.oppdatert) + " kl. " + klokkeslett(fs.oppdatert) : "") : "Ingen løpende fil – ingenting å sende.")
        : (fs.arbeider ? "Oppdaterer fila i OneDrive …" : fs.feil ? "Fila i OneDrive ble ikke oppdatert: " + fs.feil :
          fs.finnes ? "Timeføring/Til regnskap/Faktureringsgrunnlag.xlsx · oppdatert " + (fs.oppdatert ? dato(fs.oppdatert) + " kl. " + klokkeslett(fs.oppdatert) : "") : "Ingen løpende fil i OneDrive – ingenting å sende.");

      var sendKnapp = el("button", { type: "button", class: "knapp", disabled: !harLinjer || arbeider || !i.mottakerEpost, onclick: function () { sendNaa(g); } }, ikon("send"), arbeider ? "Sender …" : "Send nå");
      var lastKnapp = el("button", { type: "button", class: "knapp sekundaer", disabled: !harLinjer, onclick: lastNedLopende }, ikon("last"), "Last ned Excel");

      var topp = el("section", { class: "faktura-topp" },
        el("div", { class: "faktura-sum" },
          etikett("Løpende grunnlag"),
          el("p", { class: "faktura-tall" }, g ? kr(g.sum.total) : "–"),
          el("p", { class: "faktura-under" }, g ? "inkl. " + kr(g.sum.mva) + " mva (" + tTall(g.mvaSats != null ? g.mvaSats : i.mvaSats) + " %) · " + kr(g.sum.grunnlag) + " eks. mva" : ""),
          g && harLinjer ? el("p", { class: "faktura-periode" }, g.antallLinjer + " linjer · " + t(g.sum.timer) + " fakturerbart · periode " + dato(g.periode.fra) + "–" + dato(g.periode.til)) : null),
        el("div", { class: "faktura-handling" },
          el("div", { class: "knapperad" }, lastKnapp, sendKnapp),
          i.mottakerEpost ? el("p", { class: "faktura-mottaker" }, "Sendes til " + i.mottakerNavn + " <" + i.mottakerEpost + ">" + (i.kopiEpost ? ", kopi " + i.kopiEpost : ""))
            : el("button", { type: "button", class: "chip chip-advarsel chip-knapp", onclick: function () { gaaTil("innst"); } }, ikon("advarsel"), "Mangler e-post til View – legg inn under Innstillinger"),
          el("p", { class: "faktura-fil" }, filTekst),
          i.sendingAktiv ? el("p", { class: "faktura-auto" }, "Neste automatiske sending: " + sendeTidTekst(Store.nesteSending()) + " (" + nedtelling(Store.nesteSending()) + ")") : el("p", { class: "faktura-auto" }, "Automatisk sending er slått av.")));

      var tabell = gFeil ? el("p", { class: "melding feil" }, "Kunne ikke beregne grunnlaget: " + feilTekst(gFeil))
        : (harLinjer ? (g.perSak.length ? grunnlagTabell(g) : tomTilstand("Kun ikke-fakturerbare timer siden forrige sending.")) : tomTilstand("Ingen ufakturerte linjer. Nye føringer havner her automatisk.", "Før timer", function () { hurtigforing(); }));

      var batcher = Store.sendteBatcher();
      var historikk = batcher.length ? el("ul", { class: "historikk" }, batcher.map(function (b) {
        var lb = Store.linjerForBatch(b.id);
        var gb = null;
        try { gb = Store.grunnlag(lb); } catch (e) { /* */ }
        return el("li", { class: "hist" },
          el("div", { class: "hist-dato" }, chip("Sendt " + dato(b.sendt), "sendt"), el("small", null, "kl. " + klokkeslett(b.sendt))),
          el("div", { class: "hist-info" }, el("strong", null, (b.fil || "").split("/").pop() || "Faktureringsgrunnlag"),
            el("small", null, (lb.timer.length + lb.utlegg.length) + " linjer" + (gb && gb.periode.fra ? " · " + dato(gb.periode.fra) + "–" + dato(gb.periode.til) : ""))),
          el("div", { class: "hist-sum" }, gb ? kr(gb.sum.total) : ""),
          el("button", { type: "button", class: "knapp tekst liten", onclick: function () { lastNedBatch(b); }, "aria-label": "Last ned grunnlaget sendt " + dato(b.sendt) }, ikon("last", "ikon-liten"), "Last ned"));
      })) : tomTilstand("Ingen grunnlag er sendt ennå.");

      erstatt(node, topp,
        panel("Per sak", tabell, { klasse: "panel-tabell" }),
        panel("Sendt til regnskap", historikk));
    }
    oppdater();
    return { node: node, oppdater: function () { if (!arbeider) oppdater(); } };
  }

  // =====================================================================
  // Seksjon: Innstillinger
  // =====================================================================
  function visInnst() {
    var node = el("div", { class: "seksjon innst" });
    var i = Store.innst();
    var f = {
      mottakerNavn: felt({ etikett: "Mottaker hos View", verdi: i.mottakerNavn, autocomplete: "off" }),
      mottakerEpost: felt({ etikett: "Mottakerens e-post", type: "email", verdi: i.mottakerEpost, placeholder: "navn@view.no", hint: "Faktureringsgrunnlaget sendes hit. Uten adresse sendes ingenting." }),
      kopiEpost: felt({ etikett: "Kopi til", type: "text", verdi: i.kopiEpost, placeholder: "Valgfritt – flere adresser skilles med komma", inputmode: "email" }),
      emne: felt({ etikett: "Emne", verdi: i.emne }),
      melding: felt({ etikett: "Melding", type: "textarea", rader: 6, verdi: i.melding }),
      intervallDager: felt({ etikett: "Intervall (dager)", verdi: String(i.intervallDager), readonly: true }),
      nesteSending: felt({ etikett: "Første sending (startdato)", verdi: dato(i.nesteSending), readonly: true }),
      mvaSats: felt({ etikett: "MVA-sats (%)", type: "number", verdi: i.mvaSats, min: "0", max: "100", step: "0.1", inputmode: "decimal" })
    };
    var aktivF = bryter({ etikett: "Automatisk sending er på", verdi: i.sendingAktiv });
    aktivF.input.addEventListener("change", function () { aktivF.querySelector("label span:last-child").textContent = aktivF.input.checked ? "Automatisk sending er på" : "Automatisk sending er av"; });
    if (!i.sendingAktiv) aktivF.querySelector("label span:last-child").textContent = "Automatisk sending er av";
    var mangler = el("div", { class: "melding advarsel", role: "status", hidden: !!i.mottakerEpost },
      el("p", { class: "melding-tittel" }, ikon("advarsel", "ikon-liten"), "Mangler e-post til View"),
      el("p", null, "Legg inn e-postadressen til " + (i.mottakerNavn || "mottakeren") + " for at faktureringsgrunnlaget skal kunne sendes."));
    var lagre = el("button", { type: "submit", class: "knapp" }, "Lagre innstillinger");
    var form = el("form", { class: "skjema", novalidate: true },
      mangler,
      el("fieldset", { class: "gruppe" }, el("legend", null, "Mottaker"),
        el("div", { class: "grid-2" }, f.mottakerNavn, f.mottakerEpost), f.kopiEpost),
      el("fieldset", { class: "gruppe" }, el("legend", null, "E-posten"), f.emne, f.melding),
      el("fieldset", { class: "gruppe" }, el("legend", null, "Sending og avgift"),
        aktivF,
        el("div", { class: "rytme" },
          el("p", { class: "rytme-neste" }, el("span", null, "Neste automatiske sending"), el("strong", null, sendeTidTekst(Store.nesteSending()))),
          el("div", { class: "grid-2" }, f.intervallDager, f.nesteSending),
          el("p", { class: "felt-hint" }, "Styres av den automatiske flyten i Microsoft 365 – endres der. Flyten sender annenhver mandag kl. 07:00.")),
        el("div", { class: "grid-3" }, f.mvaSats),
        el("p", { class: "felt-hint" }, erDemo() ? "I demo lagres innstillingene bare i denne nettleseren." : "Lagres i OneDrive (Timeføring/innstillinger.json). Power Automate-flyten leser den samme fila.")),
      el("div", { class: "knapperad" }, lagre));

    form.addEventListener("submit", async function (e) {
      e.preventDefault();
      Object.keys(f).forEach(function (n) { f[n].settFeil(""); });
      var v = {
        mottakerNavn: f.mottakerNavn.input.value.trim(),
        mottakerEpost: f.mottakerEpost.input.value.trim(),
        kopiEpost: f.kopiEpost.input.value.trim(),
        emne: f.emne.input.value.trim(),
        melding: f.melding.input.value,
        mvaSats: Number(String(f.mvaSats.input.value).replace(",", ".")),
        sendingAktiv: aktivF.input.checked
      };
      var feil = [];
      if (v.mottakerEpost && !Store.gyldigEpost(v.mottakerEpost)) { f.mottakerEpost.settFeil("Skriv en gyldig e-postadresse, f.eks. navn@view.no."); feil.push(f.mottakerEpost); }
      if (v.kopiEpost && !v.kopiEpost.split(/[;,\s]+/).filter(Boolean).every(Store.gyldigEpost)) { f.kopiEpost.settFeil("En av adressene er ikke gyldig."); feil.push(f.kopiEpost); }
      if (!v.emne) { f.emne.settFeil("Skriv et emne."); feil.push(f.emne); }
      if (!(v.mvaSats >= 0 && v.mvaSats <= 100)) { f.mvaSats.settFeil("Mellom 0 og 100."); feil.push(f.mvaSats); }
      if (feil.length) { feil[0].input.focus(); return; }
      lagre.disabled = true;
      try {
        await Store.lagreInnstillinger(v);
        mangler.hidden = !!v.mottakerEpost;
        toast("Innstillingene er lagret.", "ok");
      } catch (err) { toast(feilTekst(err), "feil"); }
      finally { lagre.disabled = false; }
    });

    var konto = erDemo() ? null : Auth.konto();
    var kontoPanel = panel(erDemo() ? "Demo" : "Konto og lagring", el("div", { class: "konto" },
      erDemo()
        ? [el("p", null, "Du ser eksempeldata. Ingenting her er ekte klienter, og ingenting sendes."),
          el("button", { type: "button", class: "knapp sekundaer", onclick: async function () {
            if (await bekreft({ tittel: "Tilbakestille demodata?", tekst: "Alle endringer du har gjort i demoen blir borte.", ok: "Tilbakestill", farlig: true })) {
              Store.nullstillDemo(); await Store.start("demo"); gaaTil("innst"); toast("Demodata er tilbakestilt.", "ok");
            }
          } }, "Tilbakestill demodata")]
        : [el("dl", { class: "dl" },
            el("dt", null, "Innlogget"), el("dd", null, konto ? konto.username : ""),
            el("dt", null, "Lagring"), el("dd", null, "OneDrive › Timeføring"),
            el("dt", null, "Utlogging"), el("dd", null, "Automatisk etter 30 minutter uten aktivitet")),
          el("button", { type: "button", class: "knapp sekundaer", onclick: async function () {
            try { await Store.lastPaaNytt(); toast("Data er lastet på nytt fra OneDrive.", "ok"); } catch (e) { toast(feilTekst(e), "feil"); }
          } }, "Last inn på nytt fra OneDrive")]));

    var hjelpPanel = panel("Om timeføringen", el("div", { class: "konto" },
      el("p", null, "Timer føres her, samles i en løpende fil i OneDrive og sendes automatisk til View annenhver mandag."),
      el("button", { type: "button", class: "knapp sekundaer", onclick: visHjelp }, "Slik fungerer det"),
      el("p", { class: "firma" }, FIRMA.navn, el("br"), "Org.nr. " + FIRMA.orgnr)));
    node.appendChild(el("div", { class: "innst-grid" }, el("section", { class: "panel" }, form), el("div", { class: "innst-side" }, hjelpPanel, kontoPanel)));
    return { node: node, oppdater: function () { mangler.hidden = !!Store.innst().mottakerEpost; } };
  }

  function visHjelp() {
    var i = Store.innst();
    var mottaker = i.mottakerNavn || "mottakeren";
    var steg = [
      ["Før timer og utlegg", "På PC eller mobil. Alt lagres i OneDrive under «Timeføring» – ikke i nettleseren."],
      ["Løpende fil i OneDrive", "Timeføring › Til regnskap › Faktureringsgrunnlag.xlsx oppdateres automatisk med alle ufakturerte linjer."],
      ["Sendes automatisk til View", "Annenhver mandag kl. 07:00 sender flyten i Microsoft 365 fila til " + mottaker + " hos View og flytter den til Timeføring › Sendt. Neste gang: " + sendeTidTekst(Store.nesteSending()) + "."],
      ["Linjene låses", "Linjene som var med i fila merkes «Sendt» og kan ikke endres. Nye føringer havner i neste periode."]
    ];
    aapneDialog({
      tittel: "Slik fungerer timeføringen", klasse: "smal",
      innhold: el("div", { class: "hjelp" },
        el("ol", { class: "hjelp-steg" }, steg.map(function (x) { return el("li", null, el("strong", null, x[0]), el("span", null, x[1])); })),
        el("p", { class: "felt-hint" }, "«Send nå» under Fakturering sender med en gang. Den faste rytmen endres ikke. Uten ufakturerte linjer finnes ingen fil, og ingenting sendes."),
        el("p", { class: "firma" }, FIRMA.navn + " · Org.nr. " + FIRMA.orgnr))
    });
  }

  // =====================================================================
  // Innlogging, oppstart og utlogging
  // =====================================================================
  function demoTillatt() { return !Auth.konfigurert() || location.hash === "#demo"; }

  function visLogin(melding) {
    rot.className = "login-side";
    var loginKnapp = el("button", { type: "button", class: "knapp messing stor", disabled: !Auth.konfigurert(), onclick: function () {
      loginKnapp.disabled = true;
      loginKnapp.textContent = "Sender deg til Microsoft …";
      Auth.login().catch(function (e) { loginKnapp.disabled = false; loginKnapp.textContent = "Logg inn med Microsoft"; vis(feilTekst(e)); });
    } }, el("span", { class: "ms-logo", "aria-hidden": "true" }, el("i"), el("i"), el("i"), el("i")), "Logg inn med Microsoft");
    var meldEl = el("p", { class: "login-melding", role: "status", hidden: !melding }, melding || "");
    function vis(t) { meldEl.textContent = t; meldEl.hidden = !t; }
    erstatt(rot, el("main", { class: "login" },
      el("div", { class: "login-kort" },
        el("img", { class: "login-logo", src: "/assets/logo-negativ.svg", alt: "Advokat Geir Frøholm AS", width: "500", height: "148" }),
        el("div", { class: "login-tekst" },
          el("h1", null, "Timeføring"),
          el("p", null, "Logg inn med Microsoft-kontoen din på advokatfroholm.no.")),
        meldEl,
        el("div", { class: "login-knapper" },
          loginKnapp,
          !Auth.konfigurert() ? el("p", { class: "login-hint" }, "Innlogging er ikke satt opp ennå.") : null,
          demoTillatt() ? el("button", { type: "button", class: "knapp lys", onclick: function () { startApp("demo"); } }, "Åpne demo") : null),
        demoTillatt() ? el("p", { class: "login-hint" }, "Demoen bruker eksempeldata og lagrer bare i denne nettleseren.") : null),
      el("p", { class: "login-fot" }, "Kun for " + FIRMA.navn + " · Org.nr. " + FIRMA.orgnr),
      el("p", { class: "opphav opphav-login" }, OPPHAV)));
    var f = rot.querySelector(".login .knapp:not([disabled])");
    if (f) f.focus();
  }

  var avlyttStore = null;
  var synkSist = 0;
  function vedSynlig() {
    if (document.visibilityState !== "visible" || !Store.modus()) return;
    if (Date.now() - synkSist < 60000) return;
    synkSist = Date.now();
    Store.synk().catch(function (e) { console.warn(e); });
  }

  async function startApp(modus) {
    rot.className = "laster-side";
    erstatt(rot, el("div", { class: "laster-skjerm", role: "status" },
      el("img", { src: "/assets/symbol-negativ.svg", alt: "", width: "116", height: "116" }),
      el("p", null, modus === "demo" ? "Åpner demo …" : "Henter data fra OneDrive …")));
    try {
      await Store.start(modus);
    } catch (e) {
      console.error(e);
      erstatt(rot, el("main", { class: "login" }, el("div", { class: "login-kort" },
        el("img", { class: "login-logo", src: "/assets/logo-negativ.svg", alt: "Advokat Geir Frøholm AS", width: "500", height: "148" }),
        el("h1", null, "Kunne ikke hente data"),
        el("p", { class: "login-melding" }, feilTekst(e)),
        el("div", { class: "login-knapper" },
          el("button", { type: "button", class: "knapp messing", onclick: function () { startApp(modus); } }, "Prøv igjen"),
          el("button", { type: "button", class: "knapp lys", onclick: function () { avslutt(null); } }, modus === "demo" ? "Tilbake" : "Logg ut")))));
      rot.className = "login-side";
      return;
    }
    byggSkall(modus);
    avlyttStore = Store.lytt(oppdaterAlt);
    Store.paaMelding(function (tekst, type) { toast(tekst, type); });
    oppdaterVarsler();
    gaaTil("oversikt", { ikkeFokus: true });
    Auth.startInaktivitet(function () { avslutt("Du ble logget ut etter 30 minutter uten aktivitet."); });
    document.addEventListener("visibilitychange", vedSynlig);
  }

  function avslutt(melding) {
    var varDemo = erDemo();
    Auth.stoppInaktivitet();
    document.removeEventListener("visibilitychange", vedSynlig);
    if (avlyttStore) avlyttStore();
    document.querySelectorAll("dialog").forEach(function (d) { d.remove(); });
    Store.stopp();
    visning = null;
    if (varDemo || !Auth.konto()) { visLogin(melding); return; }
    tom(rot);
    Auth.logout(melding);
  }

  // Hurtigtast «T» for ny føring
  document.addEventListener("keydown", function (e) {
    if (!Store.modus() || e.ctrlKey || e.metaKey || e.altKey) return;
    var tag = (e.target && e.target.tagName) || "";
    if (/^(INPUT|TEXTAREA|SELECT)$/.test(tag) || (e.target && e.target.isContentEditable)) return;
    if (document.querySelector("dialog[open]")) return;
    if (e.key === "t" || e.key === "T") { e.preventDefault(); hurtigforing(); }
  });

  // Rydd opp lyttere for stoppeklokke når visningen byttes
  var opprinneligGaaTil = gaaTil;
  gaaTil = function (id, opts) {
    if (visning && visning.fjern) visning.fjern();
    opprinneligGaaTil(id, opts);
  };

  async function boot() {
    var melding = Auth.hentMelding();
    if (Auth.konfigurert() && location.hash !== "#demo") {
      try {
        var konto = await Auth.init();
        if (konto) { await startApp("onedrive"); return; }
      } catch (e) {
        console.error(e);
        melding = "Innloggingen feilet: " + feilTekst(e);
      }
    }
    visLogin(melding);
  }

  window.addEventListener("hashchange", function () {
    if (!Store.modus() && rot.classList.contains("login-side")) visLogin(null);
  });

  boot();
})();
