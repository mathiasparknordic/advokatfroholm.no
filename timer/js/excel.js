/*!
 * Timeføring – Advokat Geir Frøholm AS
 * © 2026 Mathias Frøholm. Alle rettigheter forbeholdt.
 * Kopiering, endring eller gjenbruk uten skriftlig tillatelse er ikke tillatt.
 */
/*
 * Faktura – faktureringsgrunnlag og Excel-fil til regnskapsbyrået (View).
 * Advokat Geir Frøholm AS – timeføring.
 *
 *   Faktura.grunnlag(data, innstillinger) -> { perSak, sum, periode, antallLinjer, timeIder, utleggIder, linjer }
 *   Faktura.lagXlsx(data, innstillinger)  -> Promise<ArrayBuffer>   (krever global ExcelJS)
 *   Faktura.filnavn(dato)                 -> «Faktureringsgrunnlag YYYY-MM-DD.xlsx»
 *
 * Beløp regnes i hele øre (heltall) for å unngå avrundingsfeil.
 */
(function (root) {
  "use strict";

  var FIRMA = "Advokat Geir Frøholm AS";
  var TITTEL = "Faktureringsgrunnlag – " + FIRMA;
  var AVSENDER = FIRMA + "  ·  Strandveien 55, 1366 Lysaker  ·  +47 907 24 655  ·  geir@advokatfroholm.no";

  var FARGE = {
    gronn: "FF1F3D34", gronn2: "FF2A5145", messing: "FFB08D5B", lysMessing: "FFC29F6C",
    grafitt: "FF333333", lys: "FFF5F4F0", linje: "FFE3E1DA", hvit: "FFFFFFFF", graa: "FF8A8A85"
  };
  var FONT = "Arial";
  var FMT_BELOP = "#,##0.00";        // vises «1 234,56» med norske regionsinnstillinger
  var FMT_TIMER = "0.00";
  var FMT_DATO = "dd.mm.yyyy";

  /* ------------------------------------------------------------------ */
  /* Hjelpere                                                            */
  /* ------------------------------------------------------------------ */

  function tilOre(kr) { return Math.round((Number(kr) || 0) * 100 + (kr < 0 ? -1e-7 : 1e-7)); }
  function fraOre(ore) { return ore / 100; }
  function rund2(x) { return fraOre(tilOre(x)); }
  // Timer × sats i øre, avrundet til nærmeste øre
  function belopOre(timer, sats) { var x = (Number(timer) || 0) * (Number(sats) || 0) * 100; return Math.round(x + (x < 0 ? -1e-7 : 1e-7)); }

  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function isoDato(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
  function norskDato(iso) {
    if (!iso) return "";
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
    return m ? m[3] + "." + m[2] + "." + m[1] : String(iso);
  }
  // Excel-dato: ExcelJS tolker Date som UTC – lag datoen i UTC så den ikke forskyves.
  function excelDato(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || "");
    return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null;
  }

  function filnavn(dato) {
    var d;
    if (dato == null) d = new Date();
    else if (dato instanceof Date) d = dato;
    else if (/^\d{4}-\d{2}-\d{2}/.test(String(dato))) return "Faktureringsgrunnlag " + String(dato).slice(0, 10) + ".xlsx";
    else d = new Date(dato);
    if (isNaN(d.getTime())) d = new Date();
    return "Faktureringsgrunnlag " + isoDato(d) + ".xlsx";
  }

  function saksnrNokkel(s) {
    var m = /^(\d+)\s*\/\s*(\d{4})$/.exec(String(s || "").trim());
    return m ? [+m[2], +m[1]] : [9999, 0];
  }
  function sammenlignTekst(a, b) { return String(a || "").localeCompare(String(b || ""), "nb", { sensitivity: "base" }); }

  function mvaSatsAv(innst) {
    var s = innst && innst.mvaSats;
    s = (s === "" || s == null) ? 25 : Number(s);
    return isFinite(s) ? s : 25;
  }

  /* ------------------------------------------------------------------ */
  /* Grunnlag                                                            */
  /* ------------------------------------------------------------------ */

  function grunnlag(data, innstillinger) {
    data = data || {};
    var mvaSats = mvaSatsAv(innstillinger);
    var kunder = {}, saker = {};
    (data.kunder || []).forEach(function (k) { kunder[k.id] = k; });
    (data.saker || []).forEach(function (s) { saker[s.id] = s; });

    var apne = {};
    (data.batcher || []).forEach(function (b) { if (b && b.sendt == null) apne[b.id] = true; });
    function med(l) { return l && (l.batchId == null || apne[l.batchId] === true); }

    var timer = (data.timer || []).filter(med);
    var utlegg = (data.utlegg || []).filter(med);

    var ukjentKunde = { id: "", navn: "(ukjent kunde)", orgnr: "", fakturaadresse: "", fakturaEpost: "", fakturaMerkes: "" };
    function sakFor(id) { return saker[id] || { id: id || "", saksnr: "", navn: "(ukjent sak)", kundeId: "" }; }
    function kundeFor(sak) { return kunder[sak.kundeId] || ukjentKunde; }

    var grupper = {}, rekkefolge = [];
    function gruppe(sakId) {
      var key = sakId || "";
      if (!grupper[key]) {
        var sak = sakFor(sakId);
        grupper[key] = { kunde: kundeFor(sak), sak: sak, timerH: 0, salaerOre: 0, utleggOre: 0, utleggMvaOre: 0, satser: {}, harFakturerbart: false };
        rekkefolge.push(key);
      }
      return grupper[key];
    }

    var datoer = [];
    var linjerTimer = timer.map(function (t) {
      var g = gruppe(t.sakId);
      var fakt = t.fakturerbar !== false;
      var sats = Number(t.sats);
      if (!isFinite(sats)) sats = Number(g.sak.timepris) || 0;
      var ore = fakt ? belopOre(t.timer, sats) : 0;
      if (fakt) {
        g.timerH += Number(t.timer) || 0;
        g.salaerOre += ore;
        g.satser[sats] = true;
        g.harFakturerbart = true;
      }
      if (t.dato) datoer.push(String(t.dato).slice(0, 10));
      return { linje: t, kunde: g.kunde, sak: g.sak, fakturerbar: fakt, sats: sats, belop: fraOre(ore) };
    });

    var linjerUtlegg = utlegg.map(function (u) {
      var g = gruppe(u.sakId);
      var ore = tilOre(u.belop);
      g.utleggOre += ore;
      if (u.mvaPliktig) g.utleggMvaOre += ore;
      g.harFakturerbart = true;
      if (u.dato) datoer.push(String(u.dato).slice(0, 10));
      return { linje: u, kunde: g.kunde, sak: g.sak, belop: fraOre(ore) };
    });

    function kundenrNokkel(k) { var n = Number(k && k.kundenr); return (k && k.kundenr != null && k.kundenr !== "" && isFinite(n)) ? n : Infinity; }
    function sorter(a, b) {
      var ka = kundenrNokkel(a.kunde), kb = kundenrNokkel(b.kunde);
      return (ka !== kb ? (ka < kb ? -1 : 1) : 0) ||
        sammenlignTekst(a.kunde.navn, b.kunde.navn) ||
        (saksnrNokkel(a.sak.saksnr)[0] - saksnrNokkel(b.sak.saksnr)[0]) ||
        (saksnrNokkel(a.sak.saksnr)[1] - saksnrNokkel(b.sak.saksnr)[1]) ||
        sammenlignTekst(a.sak.saksnr, b.sak.saksnr);
    }

    var sumOre = { timer: 0, salaer: 0, utlegg: 0, grunnlag: 0, mva: 0, total: 0 };
    var perSak = rekkefolge.map(function (k) { return grupper[k]; })
      .filter(function (g) { return g.harFakturerbart; })
      .sort(sorter)
      .map(function (g) {
        var satser = Object.keys(g.satser).map(Number);
        var timepris = satser.length === 1 ? satser[0] : (g.timerH ? rund2(fraOre(g.salaerOre) / g.timerH) : (Number(g.sak.timepris) || 0));
        var grunnlagOre = g.salaerOre + g.utleggOre;
        var mvaOre = Math.round((g.salaerOre + g.utleggMvaOre) * mvaSats / 100);
        var timerRund = Math.round(g.timerH * 10000) / 10000;
        sumOre.timer += timerRund;
        sumOre.salaer += g.salaerOre; sumOre.utlegg += g.utleggOre; sumOre.grunnlag += grunnlagOre;
        sumOre.mva += mvaOre; sumOre.total += grunnlagOre + mvaOre;
        return {
          kunde: g.kunde, sak: g.sak,
          timer: timerRund,
          timepris: timepris,
          blandetSats: satser.length > 1,
          salaer: fraOre(g.salaerOre),
          utlegg: fraOre(g.utleggOre),
          utleggMva: fraOre(g.utleggMvaOre),
          grunnlag: fraOre(grunnlagOre),
          mva: fraOre(mvaOre),
          total: fraOre(grunnlagOre + mvaOre)
        };
      });

    function linjeSort(a, b) {
      return sorter(a, b) || sammenlignTekst(a.linje.dato, b.linje.dato) || sammenlignTekst(a.linje.opprettet, b.linje.opprettet);
    }
    linjerTimer.sort(linjeSort);
    linjerUtlegg.sort(linjeSort);

    datoer.sort();
    return {
      perSak: perSak,
      sum: {
        timer: Math.round(sumOre.timer * 10000) / 10000,
        salaer: fraOre(sumOre.salaer), utlegg: fraOre(sumOre.utlegg), grunnlag: fraOre(sumOre.grunnlag),
        mva: fraOre(sumOre.mva), total: fraOre(sumOre.total)
      },
      periode: { fra: datoer[0] || null, til: datoer[datoer.length - 1] || null },
      mvaSats: mvaSats,
      antallLinjer: timer.length + utlegg.length,
      timeIder: timer.map(function (t) { return t.id; }),
      utleggIder: utlegg.map(function (u) { return u.id; }),
      linjer: { timer: linjerTimer, utlegg: linjerUtlegg }
    };
  }

  /* ------------------------------------------------------------------ */
  /* Excel                                                               */
  /* ------------------------------------------------------------------ */

  function kol(n) { var s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

  var TYNN_LINJE = { style: "thin", color: { argb: FARGE.linje } };

  function stilOverskrift(rad, antall) {
    rad.height = 30;
    for (var c = 1; c <= antall; c++) {
      var celle = rad.getCell(c);
      celle.font = { name: FONT, size: 10, bold: true, color: { argb: FARGE.hvit } };
      celle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: FARGE.gronn } };
      var just = celle.alignment && celle.alignment.horizontal || "left";
      celle.alignment = { vertical: "middle", horizontal: just, wrapText: true, indent: celle.alignment && celle.alignment.indent || 0 };
      celle.border = { bottom: { style: "thin", color: { argb: FARGE.messing } } };
    }
  }

  // Skriver tittelblokk (rad 1–3) og returnerer overskriftsradnummer (5).
  function tittelblokk(ws, antallKol, undertittel, info) {
    ws.getRow(1).height = 26;
    var t = ws.getCell("A1");
    t.value = undertittel ? TITTEL + " – " + undertittel : TITTEL;
    t.font = { name: FONT, size: 15, bold: true, color: { argb: FARGE.gronn } };
    t.alignment = { vertical: "middle" };
    for (var c = 1; c <= antallKol; c++) {
      ws.getRow(1).getCell(c).border = { bottom: { style: "thin", color: { argb: FARGE.messing } } };
    }
    var r2 = ws.getCell("A2");
    r2.value = info;
    r2.font = { name: FONT, size: 10, color: { argb: FARGE.grafitt } };
    ws.getRow(2).height = 18;
    r2.alignment = { vertical: "bottom" };
    var r3 = ws.getCell("A3");
    r3.value = AVSENDER;
    r3.font = { name: FONT, size: 9, color: { argb: FARGE.graa } };
    return 5;
  }

  function settKolonner(ws, defs) {
    ws.columns = defs.map(function (d) { return { key: d.key, width: d.bredde }; });
  }

  function skrivRad(ws, radNr, defs, verdier, opts) {
    opts = opts || {};
    var rad = ws.getRow(radNr);
    defs.forEach(function (d, i) {
      var celle = rad.getCell(i + 1);
      var v = verdier[i];
      celle.value = (v === undefined || v === "") ? null : v;
      celle.font = { name: FONT, size: 10, color: { argb: opts.farge || FARGE.grafitt }, bold: !!opts.fet, italic: !!opts.kursiv };
      if (d.fmt) celle.numFmt = d.fmt;
      celle.alignment = { vertical: "top", horizontal: d.just || "left", wrapText: !!d.bryt, indent: d.innrykk ? 1 : 0 };
      celle.border = opts.border || { bottom: TYNN_LINJE };
      if (opts.fyll) celle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: opts.fyll } };
    });
    return rad;
  }

  function sumRad(ws, radNr, defs, forsteData, sisteData, sumKolonner, cache, tekst, funksjon) {
    var verdier = defs.map(function () { return null; });
    verdier[0] = tekst || "Sum";
    var rad = skrivRad(ws, radNr, defs, verdier, {
      fet: true, fyll: FARGE.lys,
      border: { top: { style: "thin", color: { argb: FARGE.messing } }, bottom: { style: "double", color: { argb: FARGE.messing } } }
    });
    rad.height = 20;
    sumKolonner.forEach(function (key) {
      var idx = -1;
      defs.forEach(function (d, i) { if (d.key === key) idx = i + 1; });
      if (idx < 0) return;
      var k = kol(idx);
      var f = sisteData >= forsteData
        ? (funksjon === "SUBTOTAL" ? "SUBTOTAL(109," + k + forsteData + ":" + k + sisteData + ")" : "SUM(" + k + forsteData + ":" + k + sisteData + ")")
        : "0";
      rad.getCell(idx).value = { formula: f, result: cache[key] || 0 };
    });
    for (var c = 1; c <= defs.length; c++) rad.getCell(c).alignment = { vertical: "middle", horizontal: defs[c - 1].just || "left", indent: defs[c - 1].innrykk ? 1 : 0 };
    return rad;
  }

  function oppsettArk(ws, defs, overskriftRad, sisteDataRad, liggende) {
    ws.views = [{ state: "frozen", xSplit: 0, ySplit: overskriftRad, topLeftCell: "A" + (overskriftRad + 1), showGridLines: false, zoomScale: 100 }];
    if (sisteDataRad > overskriftRad) {
      ws.autoFilter = { from: { row: overskriftRad, column: 1 }, to: { row: sisteDataRad, column: defs.length } };
    }
    ws.pageSetup = {
      paperSize: 9, orientation: liggende ? "landscape" : "portrait",
      fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.6, header: 0.3, footer: 0.3 },
      printTitlesRow: overskriftRad + ":" + overskriftRad
    };
    ws.headerFooter = { oddFooter: "&L&8" + FIRMA + "&C&8&A&R&8Side &P av &N" };
  }

  function lagXlsx(data, innstillinger) {
    var ExcelJS = root.ExcelJS;
    if (!ExcelJS) return Promise.reject(new Error("ExcelJS er ikke lastet."));
    var g;
    try { g = grunnlag(data, innstillinger); } catch (e) { return Promise.reject(e); }
    var mvaSats = g.mvaSats;
    var naa = (innstillinger && Object.prototype.toString.call(innstillinger._naa) === "[object Date]") ? innstillinger._naa : new Date();
    var generert = norskDato(isoDato(naa)) + " kl. " + pad(naa.getHours()) + ":" + pad(naa.getMinutes());
    var periodeTekst = g.periode.fra ? norskDato(g.periode.fra) + " – " + norskDato(g.periode.til) : "ingen linjer";

    var wb = new ExcelJS.Workbook();
    wb.creator = FIRMA;
    wb.lastModifiedBy = FIRMA;
    wb.company = FIRMA;
    wb.title = TITTEL;
    wb.created = naa;
    wb.modified = naa;
    wb.calcProperties = { fullCalcOnLoad: true };

    var info = "Periode: " + periodeTekst + "     Generert: " + generert + "     MVA-sats: " + String(mvaSats).replace(".", ",") + " %";

    /* ---------- Ark 1: Sammendrag ---------- */
    var sDefs = [
      { key: "kundenr", navn: "Kundenr.", bredde: 10, just: "left" },
      { key: "kunde", navn: "Kunde", bredde: 28, bryt: true },
      { key: "orgnr", navn: "Org.nr.", bredde: 12 },
      { key: "adresse", navn: "Fakturaadresse", bredde: 28, bryt: true },
      { key: "fepost", navn: "Faktura e-post", bredde: 26, bryt: true },
      { key: "merkes", navn: "Faktura merkes", bredde: 18, bryt: true },
      { key: "saksnr", navn: "Saksnr.", bredde: 10 },
      { key: "sak", navn: "Sak", bredde: 34, bryt: true },
      { key: "advokat", navn: "Ansvarlig advokat", bredde: 16, bryt: true },
      { key: "timer", navn: "Timer", bredde: 9, fmt: FMT_TIMER, just: "right" },
      { key: "timepris", navn: "Timepris", bredde: 11, fmt: FMT_BELOP, just: "right" },
      { key: "salaer", navn: "Salær eks. mva", bredde: 14, fmt: FMT_BELOP, just: "right" },
      { key: "utlegg", navn: "Utlegg", bredde: 12, fmt: FMT_BELOP, just: "right" },
      { key: "grunnlag", navn: "Grunnlag eks. mva", bredde: 15, fmt: FMT_BELOP, just: "right" },
      { key: "mva", navn: "MVA " + String(mvaSats).replace(".", ",") + " %", bredde: 13, fmt: FMT_BELOP, just: "right" },
      { key: "total", navn: "Totalt inkl. mva", bredde: 15, fmt: FMT_BELOP, just: "right" },
      { key: "betaling", navn: "Betalingsbetingelser", bredde: 26, bryt: true, innrykk: true }
    ];
    var ws1 = wb.addWorksheet("Sammendrag", { properties: { tabColor: { argb: FARGE.gronn } } });
    settKolonner(ws1, sDefs);
    var hr = tittelblokk(ws1, sDefs.length, null, info);
    var hRad = ws1.getRow(hr);
    sDefs.forEach(function (d, i) { hRad.getCell(i + 1).value = d.navn; hRad.getCell(i + 1).alignment = { horizontal: d.just || "left", indent: d.innrykk ? 1 : 0 }; });
    stilOverskrift(hRad, sDefs.length);

    function idx(defs, key) { for (var i = 0; i < defs.length; i++) if (defs[i].key === key) return i + 1; return -1; }
    var K = {}; sDefs.forEach(function (d, i) { K[d.key] = kol(i + 1); });
    var mvaFaktor = mvaSats / 100;

    var r = hr + 1;
    g.perSak.forEach(function (p) {
      var k = p.kunde || {}, s = p.sak || {};
      var rad = skrivRad(ws1, r, sDefs, [
        k.kundenr != null && k.kundenr !== "" ? Number(k.kundenr) || String(k.kundenr) : "",
        k.navn || "", k.orgnr || "", k.fakturaadresse || "", k.fakturaEpost || "", k.fakturaMerkes || "",
        s.saksnr || "", s.navn || "", s.ansvarligAdvokat || "",
        p.timer, p.timepris, p.salaer, p.utlegg, null, null, null, s.betalingsbetingelser || ""
      ]);
      rad.getCell(idx(sDefs, "grunnlag")).value = { formula: K.salaer + r + "+" + K.utlegg + r, result: p.grunnlag };
      var mvaGrunn = p.utleggMva ? "(" + K.salaer + r + "+" + p.utleggMva.toFixed(2) + ")" : K.salaer + r;
      rad.getCell(idx(sDefs, "mva")).value = { formula: "ROUND(" + mvaGrunn + "*" + mvaFaktor + ",2)", result: p.mva };
      rad.getCell(idx(sDefs, "total")).value = { formula: K.grunnlag + r + "+" + K.mva + r, result: p.total };
      if (p.utleggMva) rad.getCell(idx(sDefs, "mva")).note = "MVA av salær + mva-pliktige utlegg (kr " + p.utleggMva.toFixed(2).replace(".", ",") + ").";
      if (p.blandetSats) rad.getCell(idx(sDefs, "timepris")).note = "Snittpris – saken har timer med ulike satser. Se arket Timelinjer.";
      r++;
    });
    if (!g.perSak.length) {
      skrivRad(ws1, r, sDefs, ["", "Ingen ufakturerte linjer"], { kursiv: true, farge: FARGE.graa });
      r++;
    }
    var sisteS = r - 1;
    sumRad(ws1, r, sDefs, hr + 1, g.perSak.length ? sisteS : hr, ["timer", "salaer", "utlegg", "grunnlag", "mva", "total"], g.sum, "Sum", "SUM");
    oppsettArk(ws1, sDefs, hr, g.perSak.length ? sisteS : hr, true);
    ws1.views[0].xSplit = 2;
    ws1.views[0].topLeftCell = "C" + (hr + 1);

    /* ---------- Ark 2: Timelinjer ---------- */
    var tDefs = [
      { key: "dato", navn: "Dato", bredde: 11, fmt: FMT_DATO },
      { key: "kunde", navn: "Kunde", bredde: 26, bryt: true },
      { key: "saksnr", navn: "Saksnr.", bredde: 10 },
      { key: "sak", navn: "Sak", bredde: 30, bryt: true },
      { key: "beskrivelse", navn: "Beskrivelse", bredde: 50, bryt: true },
      { key: "timer", navn: "Timer", bredde: 9, fmt: FMT_TIMER, just: "right" },
      { key: "sats", navn: "Sats", bredde: 11, fmt: FMT_BELOP, just: "right" },
      { key: "belop", navn: "Beløp eks. mva", bredde: 14, fmt: FMT_BELOP, just: "right" },
      { key: "id", navn: "Linje-ID", bredde: 16, innrykk: true }
    ];
    var ws2 = wb.addWorksheet("Timelinjer", { properties: { tabColor: { argb: FARGE.gronn2 } } });
    settKolonner(ws2, tDefs);
    hr = tittelblokk(ws2, tDefs.length, "timelinjer", info);
    hRad = ws2.getRow(hr);
    tDefs.forEach(function (d, i) { hRad.getCell(i + 1).value = d.navn; hRad.getCell(i + 1).alignment = { horizontal: d.just || "left", indent: d.innrykk ? 1 : 0 }; });
    stilOverskrift(hRad, tDefs.length);
    r = hr + 1;
    var tK = {}; tDefs.forEach(function (d, i) { tK[d.key] = kol(i + 1); });
    var sumTimelinjer = { timer: 0, belop: 0 };
    g.linjer.timer.forEach(function (l) {
      var t = l.linje;
      var beskr = t.beskrivelse || "";
      if (!l.fakturerbar) beskr = "Ikke fakturerbar" + (beskr ? " – " + beskr : "");
      var rad = skrivRad(ws2, r, tDefs, [
        excelDato(t.dato), l.kunde.navn || "", l.sak.saksnr || "", l.sak.navn || "", beskr,
        Number(t.timer) || 0, l.sats, null, t.id || ""
      ], l.fakturerbar ? null : { kursiv: true, farge: FARGE.graa });
      rad.getCell(idx(tDefs, "belop")).value = l.fakturerbar
        ? { formula: "ROUND(" + tK.timer + r + "*" + tK.sats + r + ",2)", result: l.belop }
        : 0;
      sumTimelinjer.timer += Number(t.timer) || 0;
      sumTimelinjer.belop += tilOre(l.belop);
      r++;
    });
    if (!g.linjer.timer.length) { skrivRad(ws2, r, tDefs, [null, "Ingen timelinjer"], { kursiv: true, farge: FARGE.graa }); r++; }
    var sisteT = r - 1;
    sumTimelinjer.timer = Math.round(sumTimelinjer.timer * 10000) / 10000;
    sumTimelinjer.belop = fraOre(sumTimelinjer.belop);
    sumRad(ws2, r, tDefs, hr + 1, g.linjer.timer.length ? sisteT : hr, ["timer", "belop"], sumTimelinjer, "Sum", "SUBTOTAL");
    var ikkeFakt = g.linjer.timer.filter(function (l) { return !l.fakturerbar; });
    if (ikkeFakt.length) {
      var hRadNr = r + 1;
      var herav = skrivRad(ws2, hRadNr, tDefs, ["Herav fakturerbare timer"], { kursiv: true, border: {} });
      herav.getCell(idx(tDefs, "timer")).value = {
        formula: "SUMIFS(" + tK.timer + (hr + 1) + ":" + tK.timer + sisteT + "," + tK.belop + (hr + 1) + ":" + tK.belop + sisteT + ",\"<>0\")",
        result: g.sum.timer
      };
    }
    oppsettArk(ws2, tDefs, hr, g.linjer.timer.length ? sisteT : hr, true);

    /* ---------- Ark 3: Utlegg ---------- */
    var uDefs = [
      { key: "dato", navn: "Dato", bredde: 11, fmt: FMT_DATO },
      { key: "kunde", navn: "Kunde", bredde: 26, bryt: true },
      { key: "saksnr", navn: "Saksnr.", bredde: 10 },
      { key: "beskrivelse", navn: "Beskrivelse", bredde: 50, bryt: true },
      { key: "belop", navn: "Beløp", bredde: 14, fmt: FMT_BELOP, just: "right" },
      { key: "mva", navn: "MVA-pliktig", bredde: 12, just: "center" },
      { key: "id", navn: "Linje-ID", bredde: 16, innrykk: true }
    ];
    var ws3 = wb.addWorksheet("Utlegg", { properties: { tabColor: { argb: FARGE.messing } } });
    settKolonner(ws3, uDefs);
    hr = tittelblokk(ws3, uDefs.length, "utlegg", info);
    hRad = ws3.getRow(hr);
    uDefs.forEach(function (d, i) { hRad.getCell(i + 1).value = d.navn; hRad.getCell(i + 1).alignment = { horizontal: d.just || "left", indent: d.innrykk ? 1 : 0 }; });
    stilOverskrift(hRad, uDefs.length);
    r = hr + 1;
    g.linjer.utlegg.forEach(function (l) {
      var u = l.linje;
      skrivRad(ws3, r, uDefs, [excelDato(u.dato), l.kunde.navn || "", l.sak.saksnr || "", u.beskrivelse || "", l.belop, u.mvaPliktig ? "Ja" : "Nei", u.id || ""]);
      r++;
    });
    if (!g.linjer.utlegg.length) { skrivRad(ws3, r, uDefs, [null, "Ingen utlegg"], { kursiv: true, farge: FARGE.graa }); r++; }
    var sisteU = r - 1;
    sumRad(ws3, r, uDefs, hr + 1, g.linjer.utlegg.length ? sisteU : hr, ["belop"], { belop: g.sum.utlegg }, "Sum", "SUBTOTAL");
    oppsettArk(ws3, uDefs, hr, g.linjer.utlegg.length ? sisteU : hr, false);

    /* ---------- Ark 4: Kunder ---------- */
    var kDefs = [
      { key: "kundenr", navn: "Kundenr.", bredde: 10 },
      { key: "navn", navn: "Kunde", bredde: 28, bryt: true },
      { key: "orgnr", navn: "Org.nr.", bredde: 12 },
      { key: "adresse", navn: "Fakturaadresse", bredde: 32, bryt: true },
      { key: "fepost", navn: "Faktura e-post", bredde: 28, bryt: true },
      { key: "merkes", navn: "Faktura merkes", bredde: 20, bryt: true },
      { key: "kontakt", navn: "Kontaktperson", bredde: 22, bryt: true },
      { key: "epost", navn: "E-post", bredde: 28, bryt: true },
      { key: "telefon", navn: "Telefon", bredde: 14 }
    ];
    var ws4 = wb.addWorksheet("Kunder", { properties: { tabColor: { argb: FARGE.lysMessing } } });
    settKolonner(ws4, kDefs);
    hr = tittelblokk(ws4, kDefs.length, "kunder", info);
    hRad = ws4.getRow(hr);
    kDefs.forEach(function (d, i) { hRad.getCell(i + 1).value = d.navn; });
    stilOverskrift(hRad, kDefs.length);
    r = hr + 1;
    var settKunder = {}, kundeliste = [];
    g.linjer.timer.concat(g.linjer.utlegg).forEach(function (l) {
      var key = l.kunde.id || l.kunde.navn;
      if (!settKunder[key]) { settKunder[key] = true; kundeliste.push(l.kunde); }
    });
    kundeliste.sort(function (a, b) {
      var an = Number(a.kundenr), bn = Number(b.kundenr);
      if (isFinite(an) && isFinite(bn) && an !== bn && a.kundenr != null && b.kundenr != null) return an - bn;
      return sammenlignTekst(a.navn, b.navn);
    });
    kundeliste.forEach(function (k) {
      skrivRad(ws4, r, kDefs, [
        k.kundenr != null && k.kundenr !== "" ? Number(k.kundenr) || String(k.kundenr) : "",
        k.navn || "", k.orgnr || "", k.fakturaadresse || "", k.fakturaEpost || "", k.fakturaMerkes || "",
        k.kontaktperson || "", k.epost || "", k.telefon || ""
      ]);
      r++;
    });
    if (!kundeliste.length) { skrivRad(ws4, r, kDefs, [null, "Ingen kunder"], { kursiv: true, farge: FARGE.graa }); r++; }
    oppsettArk(ws4, kDefs, hr, kundeliste.length ? r - 1 : hr, true);

    // Kundenr. venstrejustert som tekstlignende id; org.nr. som tekst.
    [ws1, ws4].forEach(function (ws) {
      ws.getColumn(1).eachCell(function (c, rn) { if (rn > 5 && typeof c.value === "number") c.numFmt = "0"; });
      ws.getColumn(3).eachCell(function (c, rn) { if (rn > 5 && typeof c.value === "string") c.numFmt = "@"; });
    });

    wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 0, visibility: "visible" }];

    return wb.xlsx.writeBuffer().then(function (buf) {
      if (buf instanceof ArrayBuffer) return buf;
      return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    });
  }

  root.Faktura = {
    grunnlag: grunnlag,
    lagXlsx: lagXlsx,
    filnavn: filnavn
  };
})(typeof window !== "undefined" ? window : this);
