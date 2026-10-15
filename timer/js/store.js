/*
 * Datalag for timeføringen. Global: window.Store
 *
 * To lagringsbackends med samme grensesnitt:
 *  - OneDrive (produksjon): Microsoft Graph, mappen /Timeføring i Geirs OneDrive.
 *  - Demo: localStorage med forhåndsutfylte eksempeldata (aldri ekte klienter).
 *
 * Alle endringer av data går gjennom endre(op): op kjøres på en kopi, lagres med
 * If-Match, og ved konflikt (412) lastes ferske data og op kjøres på nytt.
 */
(function () {
  "use strict";

  var MAPPE = "Timeføring";
  var DATA_STI = MAPPE + "/data.json";
  var INNST_STI = MAPPE + "/innstillinger.json";
  var GRUNNLAG_MAPPE = MAPPE + "/Til regnskap";
  var GRUNNLAG_NAVN = "Faktureringsgrunnlag.xlsx";
  var GRUNNLAG_STI = GRUNNLAG_MAPPE + "/" + GRUNNLAG_NAVN;
  var SENDT_MAPPE = MAPPE + "/Sendt";
  var KOPI_MAPPE = MAPPE + "/Sikkerhetskopi";
  var ANKER = "2026-10-26";
  var SENDETIME = 7;
  var XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  // ---------------- Hjelpere ----------------
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function datoStr(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
  function idag() { return datoStr(new Date()); }
  function leggTilDager(iso, n) {
    var p = String(iso).split("-").map(Number);
    var d = new Date(p[0], p[1] - 1, p[2]);
    d.setDate(d.getDate() + n);
    return datoStr(d);
  }
  function naa() { return new Date().toISOString(); }
  function nyId(prefiks) {
    var r = "";
    if (window.crypto && crypto.getRandomValues) {
      var a = new Uint8Array(8);
      crypto.getRandomValues(a);
      for (var i = 0; i < a.length; i++) r += (a[i] % 36).toString(36);
    } else {
      r = Math.random().toString(36).slice(2, 10);
    }
    return prefiks + "_" + Date.now().toString(36) + r;
  }
  function klon(o) { return JSON.parse(JSON.stringify(o)); }
  function rund2(x) { return Math.round(Number(x) * 100) / 100; }
  function normNavn(s) {
    return String(s || "").toLowerCase().replace(/[^a-z0-9æøåäöéü]/g, "");
  }

  function Valideringsfeil(melding) {
    var e = new Error(melding);
    e.name = "Valideringsfeil";
    return e;
  }

  function tomData() {
    return { versjon: 1, kunder: [], saker: [], timer: [], utlegg: [], batcher: [] };
  }

  function standardInnstillinger() {
    return {
      mottakerNavn: "",
      mottakerEpost: "",
      kopiEpost: "",
      emne: "Faktureringsgrunnlag – Advokat Geir Frøholm AS",
      melding: "Hei,\n\nVedlagt følger faktureringsgrunnlag for perioden.\n\nMed vennlig hilsen\nGeir Frøholm",
      intervallDager: 14,
      nesteSending: ANKER, // anker for flytens faste rytme (annenhver mandag kl. 07:00)
      mvaSats: 25,
      sendingAktiv: true
    };
  }

  function sikreSkjema(d) {
    d = d && typeof d === "object" ? d : tomData();
    ["kunder", "saker", "timer", "utlegg", "batcher"].forEach(function (k) {
      if (!Array.isArray(d[k])) d[k] = [];
    });
    if (!d.versjon) d.versjon = 1;
    return d;
  }

  function sikreInnst(i) {
    var s = standardInnstillinger();
    var ut = Object.assign({}, s, i || {});
    ut.intervallDager = parseInt(ut.intervallDager, 10) || 14;
    ut.mvaSats = Number(ut.mvaSats);
    if (!isFinite(ut.mvaSats)) ut.mvaSats = 25;
    ut.sendingAktiv = ut.sendingAktiv !== false;
    return ut;
  }

  // ---------------- OneDrive-backend ----------------
  function OneDriveBackend() {
    var kopiTatt = null; // dato for sikkerhetskopi i denne økten
    var G = window.Graph;
    return {
      navn: "onedrive",
      lesData: function () {
        // Mappene må finnes slik at Power Automate-flyten kan peke på dem.
        Promise.all([G.sikreMappe(GRUNNLAG_MAPPE), G.sikreMappe(SENDT_MAPPE)])
          .catch(function (e) { console.warn("Kunne ikke opprette mapper", e); });
        return G.lesJson(DATA_STI);
      },
      skrivData: async function (data, eTag) {
        var item = await G.skrivJson(DATA_STI, data, eTag ? { ifMatch: eTag } : { nyFil: true });
        return item.eTag;
      },
      lesInnst: function () { return G.lesJson(INNST_STI); },
      skrivInnst: async function (obj, eTag) {
        var item = await G.skrivJson(INNST_STI, obj, eTag ? { ifMatch: eTag } : { nyFil: true });
        return item.eTag;
      },
      grunnlagMeta: async function () {
        var m = await G.metadata(GRUNNLAG_STI);
        return m ? { eTag: m.eTag, lastModified: m.lastModifiedDateTime } : null;
      },
      skrivGrunnlag: async function (buffer, eTag) {
        var item = await G.skriv(GRUNNLAG_STI, buffer, eTag ? { ifMatch: eTag, contentType: XLSX_TYPE } : { nyFil: true, contentType: XLSX_TYPE });
        return { eTag: item.eTag, lastModified: item.lastModifiedDateTime };
      },
      slettGrunnlag: function (eTag) { return G.slett(GRUNNLAG_STI, eTag); },
      sisteSendt: async function () {
        var liste = await G.listMappe(SENDT_MAPPE);
        liste = liste.filter(function (f) { return f.file; });
        liste.sort(function (a, b) { return a.lastModifiedDateTime < b.lastModifiedDateTime ? 1 : -1; });
        return liste.length ? { navn: liste[0].name, lastModified: liste[0].lastModifiedDateTime } : null;
      },
      lesSendt: function (navn) { return G.lesFil(SENDT_MAPPE + "/" + navn); },
      flyttGrunnlagTilSendt: async function (navn) {
        var item = await G.flytt(GRUNNLAG_STI, SENDT_MAPPE, navn);
        return item.name;
      },
      sendEpost: function (m) { return G.sendEpost(m); },
      /** Daglig sikkerhetskopi av data slik den lå før første lagring i dag. Beholder siste 30. */
      sikkerhetskopi: async function (data) {
        var dag = idag();
        if (kopiTatt === dag) return;
        var sti = KOPI_MAPPE + "/data-" + dag + ".json";
        var finnes = await G.metadata(sti);
        if (!finnes) {
          try { await G.skrivJson(sti, data, { nyFil: true }); } catch (e) { if (e.status !== 409) throw e; }
          var filer = (await G.listMappe(KOPI_MAPPE)).filter(function (f) { return /^data-\d{4}-\d{2}-\d{2}\.json$/.test(f.name); });
          filer.sort(function (a, b) { return a.name < b.name ? 1 : -1; });
          for (var i = 30; i < filer.length; i++) {
            try { await G.slett(KOPI_MAPPE + "/" + filer[i].name); } catch (e) { /* ikke kritisk */ }
          }
        }
        kopiTatt = dag;
      }
    };
  }

  // ---------------- Demo-backend (localStorage) ----------------
  var DEMO = "timer-demo:";
  function lsLes(k) {
    try { var v = localStorage.getItem(DEMO + k); return v ? JSON.parse(v) : null; } catch (e) { return null; }
  }
  function lsSkriv(k, v) {
    try { localStorage.setItem(DEMO + k, JSON.stringify(v)); } catch (e) { /* privat modus e.l. */ }
  }
  function forsink(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  function DemoBackend() {
    var minne = {}; // reserve hvis localStorage ikke er tilgjengelig
    function les(k) { var v = lsLes(k); return v !== null ? v : (minne[k] || null); }
    function skriv(k, v) { minne[k] = v; lsSkriv(k, v); }
    function konflikt() { var e = new Error("Konflikt"); e.status = 412; return e; }
    return {
      navn: "demo",
      lesData: async function () {
        if (!les("data")) { var d = demoData(); skriv("data", { data: d.data, eTag: "1" }); skriv("innst", { data: d.innst, eTag: "1" }); skriv("filer", d.filer); }
        var v = les("data");
        return { data: v.data, eTag: v.eTag };
      },
      skrivData: async function (data, eTag) {
        var v = les("data");
        if (v && eTag && v.eTag !== eTag) throw konflikt();
        var ny = String((parseInt(v && v.eTag, 10) || 0) + 1);
        skriv("data", { data: data, eTag: ny });
        return ny;
      },
      lesInnst: async function () { var v = les("innst"); return v ? { data: v.data, eTag: v.eTag } : null; },
      skrivInnst: async function (obj, eTag) {
        var v = les("innst");
        if (v && eTag && v.eTag !== eTag) throw konflikt();
        var ny = String((parseInt(v && v.eTag, 10) || 0) + 1);
        skriv("innst", { data: obj, eTag: ny });
        return ny;
      },
      grunnlagMeta: async function () { var f = les("filer") || {}; return f.grunnlag || null; },
      skrivGrunnlag: async function (buffer, eTag) {
        var f = les("filer") || { grunnlag: null, sendt: [] };
        if (f.grunnlag && eTag && f.grunnlag.eTag !== eTag) throw konflikt();
        f.grunnlag = { eTag: "x" + Date.now(), lastModified: naa() };
        skriv("filer", f);
        return f.grunnlag;
      },
      slettGrunnlag: async function () { var f = les("filer") || { sendt: [] }; f.grunnlag = null; skriv("filer", f); return true; },
      sisteSendt: async function () {
        var f = les("filer") || { sendt: [] };
        var s = (f.sendt || []).slice().sort(function (a, b) { return a.lastModified < b.lastModified ? 1 : -1; });
        return s[0] || null;
      },
      lesSendt: async function () { throw new Error("Arkiverte filer finnes ikke i demo."); },
      flyttGrunnlagTilSendt: async function (navn) {
        var f = les("filer") || { sendt: [] };
        f.sendt = f.sendt || [];
        var faktisk = navn, n = 2;
        while (f.sendt.some(function (s) { return s.navn === faktisk; })) faktisk = navn.replace(/\.xlsx$/, " (" + (n++) + ").xlsx");
        f.sendt.push({ navn: faktisk, lastModified: naa() });
        f.grunnlag = null;
        skriv("filer", f);
        return faktisk;
      },
      sendEpost: async function () { await forsink(600); return true; },
      sikkerhetskopi: async function () { /* ikke i demo */ },
      nullstill: function () {
        ["data", "innst", "filer"].forEach(function (k) { delete minne[k]; try { localStorage.removeItem(DEMO + k); } catch (e) { /* */ } });
      }
    };
  }

  /** Eksempeldata for demo. Fiktive firmaer – ikke ekte klienter. */
  function demoData() {
    var t0 = idag();
    var aar = t0.slice(0, 4);
    var opprettet = new Date(Date.now() - 40 * 864e5).toISOString();
    var kunder = [
      { id: "k_demo1", kundenr: 1001, navn: "Nordvik Bygg AS (demo)", orgnr: "", kontaktperson: "Kari Eksempel", epost: "kari@eksempel.invalid", telefon: "400 00 001", fakturaadresse: "Postboks 1, 0001 Eksempelby", fakturaEpost: "faktura@eksempel.invalid", fakturaMerkes: "Prosjekt 2207", opprettet: opprettet },
      { id: "k_demo2", kundenr: 1002, navn: "Lysaker Eiendomsutvikling AS (demo)", orgnr: "", kontaktperson: "Ola Eksempel", epost: "ola@eksempel.invalid", telefon: "400 00 002", fakturaadresse: "Eksempelveien 2, 0002 Eksempelby", fakturaEpost: "", fakturaMerkes: "", opprettet: opprettet },
      { id: "k_demo3", kundenr: 1003, navn: "Holmen Arkitekter AS (demo)", orgnr: "", kontaktperson: "Nina Eksempel", epost: "nina@eksempel.invalid", telefon: "", fakturaadresse: "Eksempelgata 3, 0003 Eksempelby", fakturaEpost: "", fakturaMerkes: "Ref. NE", opprettet: opprettet },
      { id: "k_demo4", kundenr: 1004, navn: "Brattvoll Borettslag (demo)", orgnr: "", kontaktperson: "Per Eksempel", epost: "styret@eksempel.invalid", telefon: "", fakturaadresse: "c/o Forvalter, 0004 Eksempelby", fakturaEpost: "", fakturaMerkes: "", opprettet: opprettet }
    ];
    function sak(id, nr, kundeId, navn, pris, motpart) {
      return { id: id, saksnr: nr + "/" + aar, kundeId: kundeId, navn: navn, ansvarligAdvokat: "Geir Frøholm", saksbehandler: "Geir Frøholm", motpart: motpart || "", timepris: pris, faktureringsrutiner: "Normalt innen den 1. hver måned for foregående måned", betalingsbetingelser: "30 dager fra fakturaens utstedelse", salaerberegning: "Medgått tid", status: "aktiv", kilde: "demo", opprettet: opprettet };
    }
    var saker = [
      sak("s_demo1", 1, "k_demo1", "Tvist om sluttoppgjør NS 8407", 2690, "Eksempel Entreprenør AS"),
      sak("s_demo2", 2, "k_demo2", "Utbyggingsavtale og rekkefølgekrav", 2900, "Eksempel kommune"),
      sak("s_demo3", 3, "k_demo1", "Gjennomgang av totalentreprisekontrakt", 2690, ""),
      sak("s_demo4", 4, "k_demo3", "Prosjekteringsansvar – reklamasjon", 2700, "Eksempel Byggherre AS"),
      sak("s_demo5", 5, "k_demo4", "Reklamasjon på fasadearbeider", 2500, "Eksempel Fasade AS")
    ];
    var linjer = [
      [20, "s_demo1", 2.5, "Gjennomgang av sluttoppgjør og tilleggskrav"],
      [19, "s_demo2", 1.5, "Møte med klient om rekkefølgekrav"],
      [18, "s_demo1", 3, "Utkast til tilsvar på sluttoppgjørskrav"],
      [17, "s_demo4", 1, "Telefonmøte med klient"],
      [16, "s_demo3", 2, "Gjennomgang av kontraktsbilag og kravspesifikasjon"],
      [13, "s_demo1", 1.25, "Korrespondanse med motpartens advokat"],
      [12, "s_demo5", 2, "Befaring og gjennomgang av fotodokumentasjon", true],
      [11, "s_demo2", 0.5, "Intern avklaring – kort notat", false],
      [10, "s_demo4", 1.75, "Vurdering av ansvarsgrunnlag"],
      [6, "s_demo1", 2, "Forberedelse til forhandlingsmøte"],
      [5, "s_demo1", 3.5, "Forhandlingsmøte med motpart"],
      [5, "s_demo3", 1, "Kommentarer til revidert kontraktsutkast"],
      [4, "s_demo2", 2.25, "Utkast til utbyggingsavtale"],
      [3, "s_demo5", 1.5, "Reklamasjonsbrev til entreprenør"],
      [2, "s_demo4", 0.75, "Oppfølging av forsikringsselskap"],
      [1, "s_demo2", 1, "Telefonmøte med kommunen", true],
      [0, "s_demo1", 0.5, "Kort statusoppdatering til klient"]
    ];
    var sendtDag = 8; // linjer eldre enn dette ligger i sendt batch
    var sendtTid = new Date(Date.now() - sendtDag * 864e5);
    sendtTid.setHours(8, 2, 0, 0);
    var bSendt = { id: "b_demo1", opprettet: new Date(Date.now() - 22 * 864e5).toISOString(), sendt: sendtTid.toISOString(), fil: "Sendt/Faktureringsgrunnlag " + leggTilDager(t0, -sendtDag) + ".xlsx", timeIder: [], utleggIder: [] };
    var bApen = { id: "b_demo2", opprettet: sendtTid.toISOString(), sendt: null, fil: GRUNNLAG_STI, filETag: "x1", timeIder: [], utleggIder: [] };
    var satser = {}; saker.forEach(function (s) { satser[s.id] = s.timepris; });
    // Forskyv eksempeldatoer til virkedager (lør/søn flyttes til fredagen før)
    function virkedag(iso) {
      var p = iso.split("-").map(Number), d = new Date(p[0], p[1] - 1, p[2]);
      while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() - 1);
      return datoStr(d);
    }
    var timer = linjer.map(function (l, i) {
      var dato = virkedag(leggTilDager(t0, -l[0]));
      var b = l[0] > sendtDag ? bSendt : (l[0] > 0 ? bApen : null);
      var t = { id: "t_demo" + (i + 1), sakId: l[1], dato: dato, timer: l[2], beskrivelse: l[3], fakturerbar: l[4] !== false, sats: satser[l[1]], opprettet: dato + "T16:00:00.000Z", endret: dato + "T16:00:00.000Z", batchId: b ? b.id : null };
      if (b) b.timeIder.push(t.id);
      return t;
    });
    var utlegg = [
      { id: "u_demo1", sakId: "s_demo1", dato: virkedag(leggTilDager(t0, -15)), beskrivelse: "Rettsgebyr, forliksrådet", belop: 1243, mvaPliktig: false, batchId: "b_demo1" },
      { id: "u_demo2", sakId: "s_demo5", dato: virkedag(leggTilDager(t0, -3)), beskrivelse: "Kjøregodtgjørelse befaring", belop: 412, mvaPliktig: true, batchId: "b_demo2" }
    ];
    bSendt.utleggIder.push("u_demo1");
    bApen.utleggIder.push("u_demo2");
    var innst = standardInnstillinger();
    return {
      data: { versjon: 1, kunder: kunder, saker: saker, timer: timer, utlegg: utlegg, batcher: [bSendt, bApen] },
      innst: innst,
      filer: { grunnlag: { eTag: "x1", lastModified: naa() }, sendt: [{ navn: "Faktureringsgrunnlag " + leggTilDager(t0, -sendtDag) + ".xlsx", lastModified: sendtTid.toISOString() }] }
    };
  }

  // ---------------- Store ----------------
  var backend = null;
  var tilstand = { data: tomData(), eTag: null, innst: standardInnstillinger(), innstETag: null };
  var lyttere = [];
  var meldingsLyttere = [];
  var ko = Promise.resolve();          // serialiserer lagring av data
  var regenKo = Promise.resolve();     // serialiserer regenerering av xlsx
  var regenTimer = null;
  var filStatus = { finnes: false, oppdatert: null, feil: null, arbeider: false };

  function varsle() { lyttere.forEach(function (f) { try { f(); } catch (e) { console.error(e); } }); }
  function melding(tekst, type) { meldingsLyttere.forEach(function (f) { try { f(tekst, type || "info"); } catch (e) { console.error(e); } }); }

  function iKo(fn) {
    var p = ko.then(fn, fn);
    ko = p.catch(function () {});
    return p;
  }

  async function start(modus) {
    backend = modus === "demo" ? DemoBackend() : OneDriveBackend();
    var d = await backend.lesData();
    if (!d) {
      var tom = tomData();
      var e = await backend.skrivData(tom, null);
      d = { data: tom, eTag: e };
    }
    tilstand.data = sikreSkjema(d.data);
    tilstand.eTag = d.eTag;
    var i = await backend.lesInnst();
    if (!i) {
      var s = standardInnstillinger();
      var e2 = await backend.skrivInnst(s, null);
      i = { data: s, eTag: e2 };
    }
    tilstand.innst = sikreInnst(i.data);
    tilstand.innstETag = i.eTag;
    try {
      var m = await backend.grunnlagMeta();
      filStatus.finnes = !!m;
      filStatus.oppdatert = m ? m.lastModified : null;
    } catch (e3) { /* ikke kritisk */ }
    await synk();
    varsle();
  }

  function stopp() {
    if (regenTimer) clearTimeout(regenTimer);
    regenTimer = null;
    backend = null;
    tilstand = { data: tomData(), eTag: null, innst: standardInnstillinger(), innstETag: null };
    filStatus = { finnes: false, oppdatert: null, feil: null, arbeider: false };
    lyttere = [];
  }

  async function lastPaaNytt() {
    var d = await backend.lesData();
    if (d) { tilstand.data = sikreSkjema(d.data); tilstand.eTag = d.eTag; }
    var i = await backend.lesInnst();
    if (i) { tilstand.innst = sikreInnst(i.data); tilstand.innstETag = i.eTag; }
    await synk();
    varsle();
  }

  /** Kjører op(data) på en kopi og lagrer. Ved konflikt lastes ferske data og op kjøres på nytt. */
  function endre(op, opts) {
    opts = opts || {};
    return iKo(async function () {
      var resultat;
      for (var forsok = 0; forsok < 4; forsok++) {
        var kopi = klon(tilstand.data);
        resultat = op(kopi);
        try {
          await backend.sikkerhetskopi(tilstand.data);
        } catch (e) { console.warn("Sikkerhetskopi feilet", e); }
        try {
          var eTag = await backend.skrivData(kopi, tilstand.eTag);
          tilstand.data = kopi;
          tilstand.eTag = eTag;
          break;
        } catch (e) {
          if ((e.status === 412 || e.status === 409) && forsok < 3) {
            var fersk = await backend.lesData();
            tilstand.data = sikreSkjema(fersk && fersk.data);
            tilstand.eTag = fersk ? fersk.eTag : null;
            melding("Dataene var endret fra en annen enhet. De er lastet på nytt, og endringen din er lagt inn.", "info");
            continue;
          }
          throw e;
        }
      }
      varsle();
      if (opts.linjer !== false) planleggRegenerering();
      return resultat;
    });
  }

  // ---------------- Spørringer ----------------
  function data() { return tilstand.data; }
  function innst() { return tilstand.innst; }
  function kunde(id) { return tilstand.data.kunder.find(function (k) { return k.id === id; }) || null; }
  function sak(id) { return tilstand.data.saker.find(function (s) { return s.id === id; }) || null; }
  function batch(id, d) { return ((d || tilstand.data).batcher || []).find(function (b) { return b.id === id; }) || null; }
  function apenBatch(d) { return ((d || tilstand.data).batcher || []).find(function (b) { return !b.sendt; }) || null; }
  function sendteBatcher() {
    return tilstand.data.batcher.filter(function (b) { return !!b.sendt; }).sort(function (a, b) { return a.sendt < b.sendt ? 1 : -1; });
  }

  function erUfakturert(linje, d) {
    if (!linje.batchId) return true;
    var b = batch(linje.batchId, d);
    return !b || !b.sendt;
  }
  function erLast(linje) {
    if (!linje || !linje.batchId) return false;
    var b = batch(linje.batchId);
    return !!(b && b.sendt);
  }
  function linjeStatus(linje) {
    if (erLast(linje)) return { type: "sendt", dato: batch(linje.batchId).sendt };
    if (linje.fakturerbar === false) return { type: "ikke" };
    return { type: "ufakturert" };
  }
  function ufakturerte(d) {
    d = d || tilstand.data;
    return {
      timer: d.timer.filter(function (t) { return erUfakturert(t, d); }),
      utlegg: d.utlegg.filter(function (u) { return erUfakturert(u, d); })
    };
  }

  /** Data-utsnitt for Faktura-modulen: kun gitte linjer, batchId nullstilt. */
  function utsnitt(linjer, d) {
    d = d || tilstand.data;
    return {
      versjon: d.versjon,
      kunder: d.kunder,
      saker: d.saker,
      timer: linjer.timer.map(function (t) { return Object.assign({}, t, { batchId: null }); }),
      utlegg: linjer.utlegg.map(function (u) { return Object.assign({}, u, { batchId: null }); }),
      batcher: []
    };
  }

  function linjerForBatch(id) {
    return {
      timer: tilstand.data.timer.filter(function (t) { return t.batchId === id; }),
      utlegg: tilstand.data.utlegg.filter(function (u) { return u.batchId === id; })
    };
  }

  function grunnlag(linjer) {
    if (!window.Faktura || typeof window.Faktura.grunnlag !== "function") {
      throw new Error("Faktureringsmodulen er ikke lastet.");
    }
    return window.Faktura.grunnlag(utsnitt(linjer || ufakturerte()), tilstand.innst);
  }

  async function lagXlsx(linjer) {
    if (!window.Faktura || typeof window.Faktura.lagXlsx !== "function") {
      throw new Error("Faktureringsmodulen er ikke lastet.");
    }
    return await window.Faktura.lagXlsx(utsnitt(linjer || ufakturerte()), tilstand.innst);
  }

  function nesteSaksnr(aar) {
    aar = String(aar || new Date().getFullYear());
    var maks = 0;
    tilstand.data.saker.forEach(function (s) {
      var m = /^(\d+)\/(\d{4})$/.exec(String(s.saksnr || "").trim());
      if (m && m[2] === aar) maks = Math.max(maks, parseInt(m[1], 10));
    });
    return (maks + 1) + "/" + aar;
  }
  function saksnrFinnes(nr, unntakId) {
    var n = String(nr || "").replace(/\s/g, "");
    return tilstand.data.saker.some(function (s) { return s.id !== unntakId && String(s.saksnr || "").replace(/\s/g, "") === n; });
  }
  function nesteKundenr() {
    var maks = 1000;
    tilstand.data.kunder.forEach(function (k) { var n = parseInt(k.kundenr, 10); if (n > maks) maks = n; });
    return maks + 1;
  }
  function kundenrFinnes(nr, unntakId) {
    var n = parseInt(nr, 10);
    return tilstand.data.kunder.some(function (k) { return k.id !== unntakId && parseInt(k.kundenr, 10) === n; });
  }

  /** Finn eksisterende kunde på org.nr. (prioritert) eller navn. */
  function finnKunde(q) {
    var orgnr = window.Brreg ? window.Brreg.vask(q.orgnr) : String(q.orgnr || "").replace(/\D/g, "");
    if (orgnr.length === 9) {
      var k = tilstand.data.kunder.find(function (x) { return String(x.orgnr || "").replace(/\D/g, "") === orgnr; });
      if (k) return { kunde: k, grunn: "orgnr" };
    }
    var n = normNavn(q.navn);
    if (n) {
      var k2 = tilstand.data.kunder.find(function (x) { return normNavn(x.navn) === n; });
      if (k2) return { kunde: k2, grunn: "navn" };
    }
    return null;
  }

  // ---------------- Sende-rytme ----------------
  function dagerMellom(fra, til) {
    var a = String(fra).split("-").map(Number), b = String(til).split("-").map(Number);
    return Math.round((Date.UTC(b[0], b[1] - 1, b[2]) - Date.UTC(a[0], a[1] - 1, a[2])) / 864e5);
  }
  /**
   * Neste automatiske sending: første dato i rekka anker + n × intervallDager (n ≥ 0) som er ≥ i dag.
   * Er det sendedag og klokka har passert 07:00, gjelder neste dato i rekka.
   */
  function nesteSending(i, naaTid) {
    i = i || tilstand.innst;
    naaTid = naaTid || new Date();
    var iv = parseInt(i.intervallDager, 10) || 14;
    var anker = /^\d{4}-\d{2}-\d{2}$/.test(i.nesteSending || "") ? i.nesteSending : ANKER;
    var dag = datoStr(naaTid);
    var d = anker;
    if (d < dag) d = leggTilDager(anker, Math.ceil(dagerMellom(anker, dag) / iv) * iv);
    if (d === dag && naaTid.getHours() >= SENDETIME) d = leggTilDager(d, iv);
    return d;
  }

  // ---------------- Validering ----------------
  function sjekkTime(t) {
    if (!sak(t.sakId)) throw Valideringsfeil("Velg en sak.");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(t.dato || "")) throw Valideringsfeil("Velg en gyldig dato.");
    var h = Number(t.timer);
    if (!isFinite(h) || h <= 0 || h > 24) throw Valideringsfeil("Timer må være mellom 0 og 24.");
    var s = Number(t.sats);
    if (!isFinite(s) || s < 0) throw Valideringsfeil("Timesatsen er ugyldig.");
  }
  function sjekkUtlegg(u) {
    if (!sak(u.sakId)) throw Valideringsfeil("Velg en sak.");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(u.dato || "")) throw Valideringsfeil("Velg en gyldig dato.");
    var b = Number(u.belop);
    if (!isFinite(b) || b <= 0) throw Valideringsfeil("Beløpet må være større enn 0.");
    if (!String(u.beskrivelse || "").trim()) throw Valideringsfeil("Skriv en beskrivelse av utlegget.");
  }

  // ---------------- Mutasjoner ----------------
  async function sikreUlast(samling, id) {
    await synk();
    var l = tilstand.data[samling].find(function (x) { return x.id === id; });
    if (!l) throw Valideringsfeil("Linjen finnes ikke lenger.");
    if (erLast(l)) throw Valideringsfeil("Linjen er sendt til regnskap og kan ikke endres.");
    return l;
  }

  function leggTilTime(t) {
    var s = sak(t.sakId);
    var ny = {
      id: nyId("t"), sakId: t.sakId, dato: t.dato, timer: rund2(t.timer),
      beskrivelse: String(t.beskrivelse || "").trim(), fakturerbar: t.fakturerbar !== false,
      sats: t.sats != null && t.sats !== "" ? Number(t.sats) : (s ? Number(s.timepris) || 0 : 0),
      opprettet: naa(), endret: naa(), batchId: null
    };
    sjekkTime(ny);
    return endre(function (d) { d.timer.push(ny); return ny; });
  }

  async function endreTime(id, endr) {
    await sikreUlast("timer", id);
    return endre(function (d) {
      var t = d.timer.find(function (x) { return x.id === id; });
      if (!t) throw Valideringsfeil("Linjen finnes ikke lenger.");
      if (erUfakturert(t, d) === false) throw Valideringsfeil("Linjen er sendt til regnskap og kan ikke endres.");
      var kandidat = Object.assign({}, t, endr, { endret: naa() });
      kandidat.timer = rund2(kandidat.timer);
      kandidat.sats = Number(kandidat.sats);
      sjekkTime(kandidat);
      Object.assign(t, kandidat);
      return t;
    });
  }

  async function slettTime(id) {
    await sikreUlast("timer", id);
    return endre(function (d) {
      var t = d.timer.find(function (x) { return x.id === id; });
      if (t && !erUfakturert(t, d)) throw Valideringsfeil("Linjen er sendt til regnskap og kan ikke slettes.");
      d.timer = d.timer.filter(function (x) { return x.id !== id; });
      d.batcher.forEach(function (b) { if (!b.sendt) b.timeIder = (b.timeIder || []).filter(function (x) { return x !== id; }); });
    });
  }

  function leggTilUtlegg(u) {
    var ny = { id: nyId("u"), sakId: u.sakId, dato: u.dato, beskrivelse: String(u.beskrivelse || "").trim(), belop: rund2(u.belop), mvaPliktig: !!u.mvaPliktig, batchId: null };
    sjekkUtlegg(ny);
    return endre(function (d) { d.utlegg.push(ny); return ny; });
  }

  async function endreUtlegg(id, endr) {
    await sikreUlast("utlegg", id);
    return endre(function (d) {
      var u = d.utlegg.find(function (x) { return x.id === id; });
      if (!u) throw Valideringsfeil("Linjen finnes ikke lenger.");
      if (!erUfakturert(u, d)) throw Valideringsfeil("Utlegget er sendt til regnskap og kan ikke endres.");
      var k = Object.assign({}, u, endr);
      k.belop = rund2(k.belop);
      sjekkUtlegg(k);
      Object.assign(u, k);
      return u;
    });
  }

  async function slettUtlegg(id) {
    await sikreUlast("utlegg", id);
    return endre(function (d) {
      var u = d.utlegg.find(function (x) { return x.id === id; });
      if (u && !erUfakturert(u, d)) throw Valideringsfeil("Utlegget er sendt til regnskap og kan ikke slettes.");
      d.utlegg = d.utlegg.filter(function (x) { return x.id !== id; });
      d.batcher.forEach(function (b) { if (!b.sendt) b.utleggIder = (b.utleggIder || []).filter(function (x) { return x !== id; }); });
    });
  }

  var KUNDEFELT = ["navn", "orgnr", "kontaktperson", "epost", "telefon", "fakturaadresse", "fakturaEpost", "fakturaMerkes"];
  var SAKSFELT = ["saksnr", "kundeId", "navn", "ansvarligAdvokat", "saksbehandler", "motpart", "timepris", "faktureringsrutiner", "betalingsbetingelser", "salaerberegning", "status", "kilde"];

  function plukk(o, felt) {
    var ut = {};
    felt.forEach(function (f) { if (o[f] !== undefined) ut[f] = typeof o[f] === "string" ? o[f].trim() : o[f]; });
    return ut;
  }

  function sjekkKunde(k, unntakId) {
    if (!String(k.navn || "").trim()) throw Valideringsfeil("Kunden må ha et navn.");
    if (k.orgnr && window.Brreg && !window.Brreg.gyldigOrgnr(k.orgnr)) throw Valideringsfeil("Org.nr. er ikke gyldig.");
    if (!(parseInt(k.kundenr, 10) > 0)) throw Valideringsfeil("Kundenr. må være et positivt tall.");
    if (kundenrFinnes(k.kundenr, unntakId)) throw Valideringsfeil("Kundenr. " + k.kundenr + " er allerede i bruk.");
  }
  function sjekkSak(s, unntakId) {
    if (!/^\d+\/\d{4}$/.test(String(s.saksnr || "").replace(/\s/g, ""))) throw Valideringsfeil("Saksnr. må ha formen N/ÅÅÅÅ, f.eks. 2/2026.");
    if (saksnrFinnes(s.saksnr, unntakId)) throw Valideringsfeil("Saksnr. " + s.saksnr + " finnes allerede.");
    if (!String(s.navn || "").trim()) throw Valideringsfeil("Saken må ha et navn.");
    var p = Number(s.timepris);
    if (s.timepris === "" || s.timepris === null || !isFinite(p) || p <= 0) throw Valideringsfeil("Fyll inn timepris for saken.");
  }

  function lagKunde(k) {
    var ny = Object.assign({ id: nyId("k"), kundenr: k.kundenr || nesteKundenr(), navn: "", orgnr: "", kontaktperson: "", epost: "", telefon: "", fakturaadresse: "", fakturaEpost: "", fakturaMerkes: "", opprettet: naa() }, plukk(k, KUNDEFELT));
    ny.kundenr = parseInt(k.kundenr || ny.kundenr, 10);
    if (ny.orgnr && window.Brreg) ny.orgnr = window.Brreg.vask(ny.orgnr);
    return ny;
  }
  function lagSak(s) {
    var ny = Object.assign({ id: nyId("s"), saksnr: "", kundeId: "", navn: "", ansvarligAdvokat: "Geir Frøholm", saksbehandler: "Geir Frøholm", motpart: "", timepris: 0, faktureringsrutiner: "", betalingsbetingelser: "", salaerberegning: "", status: "aktiv", kilde: "manuell", opprettet: naa() }, plukk(s, SAKSFELT));
    ny.saksnr = String(ny.saksnr).replace(/\s/g, "");
    ny.timepris = ny.timepris === "" || ny.timepris == null ? "" : Number(ny.timepris);
    return ny;
  }

  function opprettKunde(k) {
    var ny = lagKunde(k);
    sjekkKunde(ny);
    return endre(function (d) { d.kunder.push(ny); return ny; }, { linjer: false });
  }

  function endreKunde(id, endr) {
    var gammel = kunde(id);
    if (!gammel) return Promise.reject(Valideringsfeil("Kunden finnes ikke."));
    var k = Object.assign({}, gammel, plukk(endr, KUNDEFELT));
    if (endr.kundenr !== undefined) k.kundenr = parseInt(endr.kundenr, 10);
    if (k.orgnr && window.Brreg) k.orgnr = window.Brreg.vask(k.orgnr);
    sjekkKunde(k, id);
    return endre(function (d) {
      var x = d.kunder.find(function (y) { return y.id === id; });
      if (x) Object.assign(x, k);
      return x;
    });
  }

  function opprettSak(s) {
    var ny = lagSak(s);
    if (!kunde(ny.kundeId)) return Promise.reject(Valideringsfeil("Velg kunde for saken."));
    sjekkSak(ny);
    return endre(function (d) { d.saker.push(ny); return ny; }, { linjer: false });
  }

  function endreSak(id, endr) {
    var gammel = sak(id);
    if (!gammel) return Promise.reject(Valideringsfeil("Saken finnes ikke."));
    var s = Object.assign({}, gammel, plukk(endr, SAKSFELT));
    s.saksnr = String(s.saksnr).replace(/\s/g, "");
    s.timepris = Number(s.timepris);
    sjekkSak(s, id);
    return endre(function (d) {
      var x = d.saker.find(function (y) { return y.id === id; });
      if (x) Object.assign(x, s);
      return x;
    });
  }

  /** Oppretter (eller gjenbruker) kunde og oppretter sak i én lagring. */
  function opprettKundeOgSak(k, s, eksisterendeKundeId) {
    var nyKunde = null;
    if (!eksisterendeKundeId) { nyKunde = lagKunde(k); sjekkKunde(nyKunde); }
    else if (!kunde(eksisterendeKundeId)) return Promise.reject(Valideringsfeil("Kunden finnes ikke lenger."));
    var nySak = lagSak(Object.assign({}, s, { kundeId: eksisterendeKundeId || nyKunde.id }));
    sjekkSak(nySak);
    return endre(function (d) {
      if (nyKunde) {
        if (d.kunder.some(function (x) { return parseInt(x.kundenr, 10) === nyKunde.kundenr; })) {
          var maks = 1000; d.kunder.forEach(function (x) { maks = Math.max(maks, parseInt(x.kundenr, 10) || 0); });
          nyKunde.kundenr = maks + 1;
        }
        d.kunder.push(nyKunde);
      }
      if (d.saker.some(function (x) { return x.saksnr === nySak.saksnr; })) throw Valideringsfeil("Saksnr. " + nySak.saksnr + " ble nettopp tatt i bruk. Velg et annet.");
      d.saker.push(nySak);
      return { kunde: nyKunde || d.kunder.find(function (x) { return x.id === eksisterendeKundeId; }), sak: nySak };
    }, { linjer: false });
  }

  // ---------------- Innstillinger ----------------
  function gyldigEpost(s) { return /^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]{2,}$/.test(String(s || "").trim()); }
  function gyldigeEposter(s) {
    var liste = String(s || "").split(/[;,\s]+/).filter(Boolean);
    return liste.every(gyldigEpost);
  }

  async function lagreInnstillinger(ny) {
    var kandidat = sikreInnst(Object.assign({}, tilstand.innst, ny));
    kandidat.mottakerEpost = String(kandidat.mottakerEpost || "").trim();
    kandidat.kopiEpost = String(kandidat.kopiEpost || "").trim();
    if (kandidat.mottakerEpost && !gyldigEpost(kandidat.mottakerEpost)) throw Valideringsfeil("Mottakerens e-postadresse er ikke gyldig.");
    if (kandidat.kopiEpost && !gyldigeEposter(kandidat.kopiEpost)) throw Valideringsfeil("Kopi-adressen er ikke gyldig.");
    if (!(kandidat.intervallDager >= 1 && kandidat.intervallDager <= 90)) throw Valideringsfeil("Intervallet må være mellom 1 og 90 dager.");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(kandidat.nesteSending || "")) kandidat.nesteSending = ANKER;
    if (!(kandidat.mvaSats >= 0 && kandidat.mvaSats <= 100)) throw Valideringsfeil("MVA-satsen må være mellom 0 og 100.");
    if (!String(kandidat.emne || "").trim()) throw Valideringsfeil("Emnefeltet kan ikke være tomt.");
    for (var forsok = 0; forsok < 3; forsok++) {
      try {
        tilstand.innstETag = await backend.skrivInnst(kandidat, tilstand.innstETag);
        tilstand.innst = kandidat;
        break;
      } catch (e) {
        if ((e.status === 412 || e.status === 409) && forsok < 2) {
          var fersk = await backend.lesInnst();
          tilstand.innstETag = fersk ? fersk.eTag : null;
          continue;
        }
        throw e;
      }
    }
    varsle();
    planleggRegenerering();
    return tilstand.innst;
  }

  // ---------------- Batch og løpende fil ----------------
  /** Leser alle Linje-ID-er fra arkene «Timelinjer» og «Utlegg». Kaster ved feil/ukjent format. */
  async function linjeIderIXlsx(buffer) {
    if (!window.ExcelJS) throw new Error("ExcelJS er ikke lastet.");
    var wb = new window.ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    var ider = new Set();
    var funnetTimelinjer = false;
    ["Timelinjer", "Utlegg"].forEach(function (navn) {
      var ws = wb.getWorksheet(navn);
      if (!ws) return;
      var hode = null;
      ws.eachRow(function (rad, r) {
        if (hode) {
          var v = rad.getCell(hode.c).value;
          if (v && typeof v === "object" && v.richText) v = v.richText.map(function (x) { return x.text; }).join("");
          if (v && typeof v === "object" && v.result !== undefined) v = v.result;
          v = String(v == null ? "" : v).trim();
          if (/^[tu]_/.test(v)) ider.add(v);
          return;
        }
        rad.eachCell(function (celle, c) {
          if (!hode && String(celle.value == null ? "" : celle.value).trim() === "Linje-ID") hode = { r: r, c: c };
        });
      });
      if (navn === "Timelinjer" && hode) funnetTimelinjer = true;
    });
    if (!funnetTimelinjer) throw new Error("Fant ikke kolonnen «Linje-ID» i arket «Timelinjer».");
    return ider;
  }

  /**
   * Steg 2 i sendelogikken: Har åpen batch en registrert fil som er borte,
   * har flyten sendt den. Marker batchen sendt og lås linjene.
   */
  async function synk() {
    if (!backend) return false;
    var b = apenBatch();
    if (!b || !b.fil) return false;
    var meta = await backend.grunnlagMeta();
    if (meta) { filStatus.finnes = true; filStatus.oppdatert = meta.lastModified; return false; }
    var sist = null;
    try { sist = await backend.sisteSendt(); } catch (e) { /* bruk nå */ }
    var bId = b.id;
    var forrigeSendt = tilstand.data.batcher.filter(function (y) { return y.sendt && y.id !== bId; })
      .map(function (y) { return y.sendt; }).sort().pop() || "";
    // Race-sikring: les Linje-ID-ene i den arkiverte fila og lås bare dem.
    var iFila = null;
    if (sist && sist.lastModified > forrigeSendt) {
      try {
        iFila = await linjeIderIXlsx(await backend.lesSendt(sist.navn));
      } catch (e) {
        console.warn("Kunne ikke lese arkivert fil – låser alle linjer i batchen.", e);
        iFila = null;
      }
    }
    var tilbake = 0;
    await endre(function (d) {
      var x = batch(bId, d);
      if (!x || x.sendt) return;
      // Nyeste fil i Sendt hører til denne batchen hvis den er nyere enn forrige registrerte sending.
      var forrige = d.batcher.filter(function (y) { return y.sendt && y.id !== bId; })
        .map(function (y) { return y.sendt; }).sort().pop() || "";
      var brukSendt = !!(sist && sist.lastModified > forrige);
      if (!brukSendt) iFila = null;
      x.sendt = brukSendt ? sist.lastModified : naa();
      if (brukSendt) x.fil = "Sendt/" + sist.navn;
      x.filETag = null;
      tilbake = 0;
      if (iFila) {
        d.timer.concat(d.utlegg).forEach(function (l) {
          var iSendtFil = iFila.has(l.id);
          if (l.batchId === bId && !iSendtFil) { l.batchId = null; tilbake++; }
          // Linjer som står i den sendte fila, men ikke ble knyttet til batchen (f.eks. fordi
          // lagringen etter skriving feilet), må også låses – ellers faktureres de på nytt.
          else if (iSendtFil && l.batchId !== bId && erUfakturert(l, d)) { l.batchId = bId; }
        });
      }
      x.timeIder = d.timer.filter(function (t) { return t.batchId === bId; }).map(function (t) { return t.id; });
      x.utleggIder = d.utlegg.filter(function (u) { return u.batchId === bId; }).map(function (u) { return u.id; });
    }, { linjer: false });
    if (tilbake) melding(tilbake + (tilbake === 1 ? " linje" : " linjer") + " kom ikke med i sendingen og er lagt tilbake som ufakturert.", "info");
    filStatus.finnes = false;
    filStatus.oppdatert = null;
    try {
      var i = await backend.lesInnst();
      if (i) { tilstand.innst = sikreInnst(i.data); tilstand.innstETag = i.eTag; }
    } catch (e) { /* ikke kritisk */ }
    melding("Faktureringsgrunnlaget er sendt til regnskap. Linjene i det er nå låst.", "ok");
    varsle();
    // Nye linjer etter sending skal inn i ny fil.
    planleggRegenerering();
    return true;
  }

  function planleggRegenerering() {
    if (!backend) return;
    if (regenTimer) clearTimeout(regenTimer);
    regenTimer = setTimeout(function () {
      regenTimer = null;
      regenerer().catch(function (e) {
        console.error(e);
        filStatus.feil = e.message || String(e);
        melding("Kunne ikke oppdatere faktureringsgrunnlaget i OneDrive: " + (e.message || e), "feil");
        varsle();
      });
    }, 1200);
  }

  /** Skriver (eller sletter) den løpende fila. Returnerer ArrayBuffer for innholdet, eller null. */
  function regenerer() {
    var p = regenKo.then(regenererNaa, regenererNaa);
    regenKo = p.catch(function () {});
    return p;
  }

  async function regenererNaa() {
    if (!backend) return null;
    filStatus.arbeider = true; filStatus.feil = null; varsle();
    try {
      for (var forsok = 0; forsok < 3; forsok++) {
        await synk();
        var u = ufakturerte();
        var b = apenBatch();
        if (u.timer.length + u.utlegg.length === 0) {
          if (b && b.fil) {
            var bid0 = b.id;
            await endre(function (d) { var x = batch(bid0, d); if (x) { x.fil = null; x.filETag = null; x.timeIder = []; x.utleggIder = []; } }, { linjer: false });
          }
          await backend.slettGrunnlag();
          filStatus.finnes = false; filStatus.oppdatert = null;
          return null;
        }
        var buffer = await lagXlsx(u);
        var meta = await backend.grunnlagMeta();
        if (!meta) {
          // Fila forsvant mens vi bygde den nye (flyten har trolig nettopp sendt den).
          // Ikke opprett en ny fil med de samme linjene – kjør synk først.
          var bNaa = apenBatch();
          if (bNaa && bNaa.fil) {
            if (forsok < 2) continue;
            throw new Error("Faktureringsgrunnlaget ble sendt mens det ble oppdatert. Last siden på nytt.");
          }
        }
        var skrevet;
        try {
          skrevet = await backend.skrivGrunnlag(buffer, meta ? meta.eTag : null);
        } catch (e) {
          if ((e.status === 412 || e.status === 409 || e.status === 404) && forsok < 2) continue;
          throw e;
        }
        var tIder = u.timer.map(function (t) { return t.id; });
        var uIder = u.utlegg.map(function (x) { return x.id; });
        var utdatert = false;
        await endre(function (d) {
          // endre() kan ha lastet ferske data (412). Da kan fila vi skrev bygge på utdatert
          // tilstand – f.eks. inneholde linjer som i mellomtiden er sendt fra en annen enhet.
          // Knytt bare linjer som fortsatt er ufakturerte, og bygg fila på nytt ved avvik.
          var fersk = ufakturerte(d);
          var fT = fersk.timer.map(function (t) { return t.id; }).sort().join("|");
          var fU = fersk.utlegg.map(function (y) { return y.id; }).sort().join("|");
          utdatert = fT !== tIder.slice().sort().join("|") || fU !== uIder.slice().sort().join("|");
          var x = apenBatch(d);
          if (!x) { x = { id: nyId("b"), opprettet: naa(), sendt: null, fil: null, timeIder: [], utleggIder: [] }; d.batcher.push(x); }
          x.fil = GRUNNLAG_STI;
          x.filETag = skrevet.eTag;
          d.timer.forEach(function (t) { if (tIder.indexOf(t.id) >= 0 && erUfakturert(t, d)) t.batchId = x.id; });
          d.utlegg.forEach(function (y) { if (uIder.indexOf(y.id) >= 0 && erUfakturert(y, d)) y.batchId = x.id; });
          x.timeIder = d.timer.filter(function (t) { return t.batchId === x.id; }).map(function (t) { return t.id; });
          x.utleggIder = d.utlegg.filter(function (y) { return y.batchId === x.id; }).map(function (y) { return y.id; });
        }, { linjer: false });
        filStatus.finnes = true; filStatus.oppdatert = skrevet.lastModified;
        if (utdatert) {
          if (forsok < 2) continue; // skriv fila på nytt med ferske data
          // Siste utvei: fjern fila heller enn å la en mulig feil fil bli sendt.
          try {
            await endre(function (d) { var x = apenBatch(d); if (x) { x.fil = null; x.filETag = null; } }, { linjer: false });
            await backend.slettGrunnlag();
            filStatus.finnes = false; filStatus.oppdatert = null;
          } catch (e2) { console.error(e2); }
          throw new Error("Dataene endret seg mens faktureringsgrunnlaget ble laget. Last siden på nytt før neste sending.");
        }
        return buffer;
      }
      throw new Error("Fila ble endret samtidig flere ganger. Prøv igjen.");
    } finally {
      filStatus.arbeider = false;
      varsle();
    }
  }

  /** «Send nå»: e-post med vedlegg, flytt fila til Sendt/, marker batch sendt, oppdater neste sending. */
  function sendNaa() {
    // Kjøres i samme kø som regenerering, så fila ikke skrives om mellom e-post og flytting.
    var p = regenKo.then(sendNaaIntern, sendNaaIntern);
    regenKo = p.catch(function () {});
    return p;
  }

  async function sendNaaIntern() {
    var i = tilstand.innst;
    if (!i.mottakerEpost || !gyldigEpost(i.mottakerEpost)) throw Valideringsfeil("Legg inn e-postadressen til mottakeren hos View under Innstillinger først.");
    if (regenTimer) { clearTimeout(regenTimer); regenTimer = null; }
    var buffer = await regenererNaa();
    if (!buffer) throw Valideringsfeil("Det er ingen ufakturerte linjer å sende.");
    // Lås nøyaktig linjene som ligger i vedlegget som sendes.
    var iVedlegg = await linjeIderIXlsx(buffer);
    var dato = idag();
    var navn = "Faktureringsgrunnlag " + dato + ".xlsx";
    await backend.sendEpost({
      til: i.mottakerEpost,
      kopi: i.kopiEpost,
      emne: i.emne,
      tekst: i.melding,
      vedlegg: { navn: navn, innhold: buffer, type: XLSX_TYPE }
    });
    var faktiskNavn = navn;
    try {
      faktiskNavn = await backend.flyttGrunnlagTilSendt(navn);
    } catch (e) {
      // E-posten er sendt. Fila må likevel ut av «Til regnskap» så flyten ikke sender igjen.
      console.error(e);
      var fjernet = false;
      try { await backend.slettGrunnlag(); fjernet = true; } catch (e2) { console.error(e2); }
      melding(fjernet
        ? "E-posten er sendt, men fila kunne ikke flyttes til Sendt. Den er fjernet fra «Til regnskap»."
        : "E-posten er sendt, men fila ligger fortsatt i «Til regnskap». Slett den i OneDrive, ellers sendes den på nytt ved neste automatiske sending.", "feil");
    }
    var b = apenBatch();
    var bId = b && b.id;
    var tilbake = 0;
    await endre(function (d) {
      var x = batch(bId, d);
      if (!x) return;
      x.sendt = naa();
      x.fil = "Sendt/" + faktiskNavn;
      x.filETag = null;
      tilbake = 0;
      d.timer.concat(d.utlegg).forEach(function (l) {
        var iFil = iVedlegg.has(l.id);
        if (l.batchId === bId && !iFil) { l.batchId = null; tilbake++; }
        else if (iFil && l.batchId !== bId && erUfakturert(l, d)) l.batchId = bId;
      });
      x.timeIder = d.timer.filter(function (t) { return t.batchId === bId; }).map(function (t) { return t.id; });
      x.utleggIder = d.utlegg.filter(function (u) { return u.batchId === bId; }).map(function (u) { return u.id; });
    }, { linjer: false });
    filStatus.finnes = false; filStatus.oppdatert = null;
    if (tilbake) melding(tilbake + (tilbake === 1 ? " linje" : " linjer") + " kom ikke med i e-posten og ligger fortsatt som ufakturert.", "info");
    varsle();
    // Rydd «Til regnskap»: ny fil for linjer ført under sendingen, eller fjern en fil som ble liggende igjen.
    planleggRegenerering();
    return { navn: faktiskNavn };
  }

  window.Store = {
    GRUNNLAG_STI: GRUNNLAG_STI,
    XLSX_TYPE: XLSX_TYPE,
    start: start,
    stopp: stopp,
    lastPaaNytt: lastPaaNytt,
    modus: function () { return backend ? backend.navn : null; },
    lytt: function (f) { lyttere.push(f); return function () { lyttere = lyttere.filter(function (x) { return x !== f; }); }; },
    paaMelding: function (f) { meldingsLyttere.push(f); },
    data: data,
    innst: innst,
    kunde: kunde,
    sak: sak,
    batch: batch,
    apenBatch: function () { return apenBatch(); },
    sendteBatcher: sendteBatcher,
    erLast: erLast,
    linjeStatus: linjeStatus,
    ufakturerte: function () { return ufakturerte(); },
    linjerForBatch: linjerForBatch,
    grunnlag: grunnlag,
    lagXlsx: lagXlsx,
    nesteSaksnr: nesteSaksnr,
    saksnrFinnes: saksnrFinnes,
    nesteKundenr: nesteKundenr,
    kundenrFinnes: kundenrFinnes,
    finnKunde: finnKunde,
    leggTilTime: leggTilTime,
    endreTime: endreTime,
    slettTime: slettTime,
    leggTilUtlegg: leggTilUtlegg,
    endreUtlegg: endreUtlegg,
    slettUtlegg: slettUtlegg,
    opprettKunde: opprettKunde,
    endreKunde: endreKunde,
    opprettSak: opprettSak,
    endreSak: endreSak,
    opprettKundeOgSak: opprettKundeOgSak,
    lagreInnstillinger: lagreInnstillinger,
    synk: synk,
    regenerer: regenerer,
    sendNaa: sendNaa,
    filStatus: function () { return filStatus; },
    gyldigEpost: gyldigEpost,
    nullstillDemo: function () { if (backend && backend.nullstill) backend.nullstill(); },
    idag: idag,
    nesteSending: function (i, t) { return nesteSending(i, t); },
    ANKER: ANKER,
    leggTilDager: leggTilDager
  };
})();
