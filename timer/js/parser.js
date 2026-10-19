/*!
 * Timeføring – Advokat Geir Frøholm AS
 * © 2026 Mathias Frøholm. Alle rettigheter forbeholdt.
 * Kopiering, endring eller gjenbruk uten skriftlig tillatelse er ikke tillatt.
 */
/*
 * Oppdrag – tolker oppdragsbekreftelser (limt inn tekst eller .docx).
 * Advokat Geir Frøholm AS – timeføring.
 *
 *   Oppdrag.fraTekst(tekst)        -> { kunde, sak, advarsler, ukjenteFelt }
 *   Oppdrag.fraDocx(arrayBuffer)   -> Promise<samme resultat>   (krever global JSZip + DOMParser)
 *   Oppdrag.gyldigOrgnr(s)         -> boolean (9 siffer + mod11)
 *   Oppdrag.normaliserNavn(s)      -> «NORDVIK BYGG AS» -> «Nordvik Bygg AS»
 *   Oppdrag.docxTilTekst(arrayBuffer) -> Promise<string> (mellomformatet fraDocx tolker)
 *
 * Ingen avhengigheter utover JSZip (kun fraDocx). Ingen eval.
 */
(function (root) {
  "use strict";

  /* ------------------------------------------------------------------ */
  /* Etiketter                                                           */
  /* ------------------------------------------------------------------ */

  // Hjelpere for regex-kilder: fleksibel bindestrek/mellomrom og punktum.
  var SP = "[\\s\\u00a0]*";           // valgfritt mellomrom
  var EPOST = "e" + SP + "-?" + SP + "(?:post|mail)"; // E-post / Epost / E-mail / E post
  var ORGNR = "(?:org(?:anisasjons)?" + SP + "\\.?" + SP + "-?" + SP + "n(?:umme)?r\\.?|organisasjonsnummer)";

  // [mål, [regex-kilder for etiketten]]. Lengste treff vinner, så rekkefølgen er uviktig.
  var ETIKETTER = [
    ["kunde.navn", ["klient", "klientnavn", "klientens" + SP + "navn", "kundenavn", "oppdragsgiver"]],
    ["kunde.orgnr", [ORGNR, "klientens" + SP + ORGNR, ORGNR + SP + "klient", "foretaksnummer"]],
    ["kunde.kontaktperson", ["kontaktperson" + SP + "(?:hos" + SP + ")?klient(?:en)?", "kontaktperson", "klientens" + SP + "kontaktperson", "kontakt" + SP + "klient"]],
    ["kunde.epost", [EPOST + SP + "(?:til" + SP + ")?klient(?:en)?", "klientens" + SP + EPOST, EPOST]],
    ["kunde.telefon", ["telefon" + SP + "(?:til" + SP + ")?klient(?:en)?", "tlf\\.?" + SP + "klient(?:en)?", "klientens" + SP + "telefon", "telefon(?:nummer)?", "tlf\\.?"]],
    ["kunde.fakturaMerkes", ["faktura" + SP + "merkes", "fakturamerking", "faktura" + SP + "merking", "merking" + SP + "faktura", "deres" + SP + "ref(?:eranse|\\.)?"]],
    ["kunde.fakturaadresse", ["faktura" + SP + "-?" + SP + "adresse", "adresse" + SP + "faktura"]],
    ["kunde.fakturaEpost", ["faktura" + SP + "-?" + SP + EPOST, EPOST + SP + "(?:for" + SP + ")?faktura", "ehf" + SP + "/" + SP + EPOST]],
    ["sak.saksnr", ["saksnummer", "saksnr\\.?", "sak" + SP + "nr\\.?", "prosjektnummer", "prosjektnr\\.?", "referansenummer"]],
    ["sak.navn", ["sakens" + SP + "navn", "saksnavn", "saken" + SP + "gjelder", "oppdrag(?:et)?" + SP + "gjelder", "oppdragets" + SP + "navn"]],
    ["sak.ansvarligAdvokat", ["ansvarlig" + SP + "advokat", "ansvarlig" + SP + "partner"]],
    ["sak.saksbehandler", ["saksbehandler(?:e)?"]],
    ["sak.sekretaer", ["sekret(?:æ|ae|e)r" + SP + "-?" + SP + "kontakt", "sekret(?:æ|ae|e)r"]],
    ["sak.motpart", ["motpart(?:er|en)?", "motparter"]],
    ["sak.faktureringsrutiner", ["faktureringsrutine(?:r)?", "faktureringsrutine"]],
    ["sak.betalingsbetingelser", ["betalingsbetingelse(?:r)?", "betalingsfrist"]],
    ["sak.timepris", ["timepris(?:en)?", "timesats(?:en)?", "timerate", "timepris" + SP + "eks\\.?" + SP + "mva\\.?"]],
    ["sak.salaerberegning", ["sal(?:æ|ae|a)rberegning", "sal(?:æ|ae|a)r", "honorar", "honorarberegning", "vederlag"]]
  ];

  // Felt som kan gå over flere linjer og skal beholde linjeskift (adresse).
  var FLERLINJE_KOMMA = { "kunde.fakturaadresse": true };

  // Bygg regex per etikett: ^ etikett (valgfri kolon) og deretter: kolon+mellomrom, tab, 2+ mellomrom eller linjeslutt.
  var MATCHERE = [];
  ETIKETTER.forEach(function (def) {
    def[1].forEach(function (src) {
      MATCHERE.push({
        mal: def[0],
        re: new RegExp("^[\\s\\u00a0]*(" + src + ")[\\s\\u00a0]*(?::[\\s\\u00a0]*([\\s\\S]*)$|\\t[\\s\\u00a0]*([\\s\\S]*)$| {2,}([\\s\\S]*)$|$)", "i")
      });
    });
  });

  function finnEtikett(linje) {
    var best = null;
    for (var i = 0; i < MATCHERE.length; i++) {
      var m = MATCHERE[i].re.exec(linje);
      if (m) {
        var lengde = m[1].length;
        if (!best || lengde > best.lengde) {
          var verdi = m[2] != null ? m[2] : (m[3] != null ? m[3] : (m[4] != null ? m[4] : ""));
          best = { mal: MATCHERE[i].mal, lengde: lengde, verdi: verdi, harKolon: m[2] != null };
        }
      }
    }
    return best;
  }

  // «Noe:\tVerdi» eller «Noe:   Verdi» som ikke er en kjent etikett.
  var UKJENT_RE = /^[\s\u00a0]*([^\t:]{1,40}?)[\s\u00a0]*:(?:\t|[\s\u00a0]{2,}|[\s\u00a0]*$)[\s\u00a0]*(.*)$/;

  var TITTEL_RE = /^[\s\u00a0]*oppdragsbekreftelse(?:[\s\u00a0]*[-–—:][\s\u00a0]*|[\s\u00a0]+)(?:(?:for|til)[\s\u00a0]+)?(.*)$/i;

  /* ------------------------------------------------------------------ */
  /* Hjelpere                                                            */
  /* ------------------------------------------------------------------ */

  function rydd(s) {
    return String(s == null ? "" : s)
      .replace(/[\u00a0\u2007\u202f]/g, " ")
      .replace(/[\u200b\u200c\u200d\ufeff\u00ad]/g, "")
      .replace(/[ \t]+/g, " ")
      .trim();
  }

  var TOM_RE = /^(?:n\s*\/\s*a|n\.\s*a\.?|na|-+|–|—|\.|ikke\s+(?:aktuelt|relevant|oppgitt)|ingen\s+oppgitt|tbd|x)$/i;
  // Plassholdere i hakeparentes fra Word-malen: «[Navn]», «[fra dashboard, f.eks. 2/2026]», «kr [timepris]».
  var PLASSHOLDER_RE = /(?:\b(?:kr\.?|nok)[\s\u00a0]*)?\[[^\[\]]*\]/gi;
  function fjernPlassholdere(s) {
    s = String(s == null ? "" : s);
    if (!/\[[^\[\]]*\]/.test(s)) return s;
    s = s.replace(PLASSHOLDER_RE, " ");
    s = s.replace(/[ \t\u00a0]+/g, " ").replace(/ ([,.;:!?)])/g, "$1").replace(/\( /g, "(").trim();
    // Bare tegnsetting/«kr» igjen -> tomt
    if (/^(?:kr\.?|nok)?[\s,.;:\-–—/()]*$/i.test(s)) return "";
    return s;
  }

  function tomVerdi(s) {
    s = rydd(s);
    return s === "" || TOM_RE.test(s);
  }

  var SELSKAPSFORMER = ["AS", "ASA", "ANS", "DA", "SA", "BA", "KS", "ENK", "NUF", "IKS", "KF", "SE", "SF", "HF", "BBL", "FLI", "AB", "ApS", "GmbH", "Ltd", "LLC", "Inc", "BV", "NV", "Oy", "AG"];
  var SMAAORD = { og: 1, i: 1, av: 1, for: 1, "på": 1, til: 1, med: 1, "&": 1 };

  function storForbokstav(ord) {
    // Håndterer bindestrek: «ARE-MAGNE» -> «Are-Magne»
    return ord.split("-").map(function (d) {
      if (!d) return d;
      return d.charAt(0).toLocaleUpperCase("nb-NO") + d.slice(1).toLocaleLowerCase("nb-NO");
    }).join("-");
  }

  function normaliserNavn(navn) {
    navn = rydd(navn).replace(/^[«"'“]+|[»"'”]+$/g, "").replace(/[.,;:]+$/, "").trim();
    if (!navn) return "";
    var harSmaa = /[a-zæøåäöü]/.test(navn);
    var former = {};
    SELSKAPSFORMER.forEach(function (f) { former[f.toUpperCase()] = f; });
    var ord = navn.split(" ");
    return ord.map(function (o, i) {
      var ren = o.replace(/[.,]/g, "").toUpperCase();
      if (former[ren] && (i > 0 || ord.length === 1)) return former[ren] + o.replace(/^[^.,]*/, "").replace(/\./g, "");
      if (harSmaa) return o;                     // allerede blandet skrift – behold
      if (i > 0 && SMAAORD[o.toLowerCase()]) return o.toLowerCase();
      if (/\d/.test(o) && !/^[A-ZÆØÅ]/.test(o)) return o;
      if (o.length <= 1) return o;
      return storForbokstav(o);
    }).join(" ");
  }

  function gyldigOrgnr(s) {
    var d = String(s == null ? "" : s).replace(/\D/g, "");
    if (!/^\d{9}$/.test(d)) return false;
    var vekter = [3, 2, 7, 6, 5, 4, 3, 2];
    var sum = 0;
    for (var i = 0; i < 8; i++) sum += vekter[i] * +d.charAt(i);
    var k = 11 - (sum % 11);
    if (k === 11) k = 0;
    if (k === 10) return false;
    return k === +d.charAt(8);
  }

  function renOrgnr(s) {
    s = rydd(s);
    if (!s) return "";
    var u = s.replace(/^NO[\s-]*/i, "").replace(/[\s-]*MVA$/i, "");
    var d = u.replace(/[\s.\-\u00a0]/g, "");
    return /^\d+$/.test(d) ? d : s;           // ukjent format: behold slik det står
  }

  // Tall som «2 690», «2.690», «2 690,-», «2690,50» -> number
  function tilTall(s) {
    s = String(s).replace(/[\s\u00a0\u202f.]/g, "").replace(/,-+$/, "").replace(/,$/, "");
    if (/^\d+,\d{1,2}$/.test(s)) s = s.replace(",", ".");
    var n = parseFloat(s);
    return isFinite(n) ? n : null;
  }

  var TALL_RE = /(\d[\d \u00a0\u202f.]*\d|\d)(?:,(\d{1,2})(?!\d)|,-+)?/;
  var KR_PREFIKS_RE = /(?:\bkr\.?|\bnok)[\s\u00a0]*(\d[\d \u00a0\u202f.]*\d|\d)(?:,(\d{1,2})(?!\d)|,-+)?/i;

  function timeprisFraFelt(s) {
    s = rydd(s);
    if (!s) return null;
    var m = KR_PREFIKS_RE.exec(s) || TALL_RE.exec(s);
    if (!m) return null;
    var tall = tilTall(m[1] + (m[2] ? "," + m[2] : ""));
    return tall;
  }

  function timeprisFraSalaer(s) {
    s = rydd(s);
    if (!s) return null;
    // Foretrekk «timepris ... kr N»
    var m = /time(?:pris|sats)[^.]*?(?:\bkr\.?|\bnok)[\s\u00a0]*(\d[\d \u00a0\u202f.]*\d|\d)(?:,(\d{1,2})(?!\d))?/i.exec(s) || KR_PREFIKS_RE.exec(s);
    if (!m) return null;
    return tilTall(m[1] + (m[2] ? "," + m[2] : ""));
  }

  var TIMEPRIS_MIN = 2500, TIMEPRIS_MAX = 3000;

  function kr(n) {
    var deler = String(n).split(".");
    return deler[0].replace(/\B(?=(\d{3})+(?!\d))/g, " ") + (deler[1] ? "," + (deler[1] + "0").slice(0, 2) : "");
  }

  // Skrivefeil med ett ekstra siffer: alle tall som fås ved å fjerne ett siffer og havner i intervallet.
  function forslagTimepris(tp) {
    if (!(tp > TIMEPRIS_MAX) || Math.floor(tp) !== tp) return [];
    var s = String(tp), ut = [];
    for (var i = 0; i < s.length; i++) {
      var k = s.slice(0, i) + s.slice(i + 1);
      if (!k || k.charAt(0) === "0") continue;
      var n = +k;
      if (n >= TIMEPRIS_MIN && n <= TIMEPRIS_MAX && ut.indexOf(n) < 0) ut.push(n);
    }
    return ut.sort(function (a, b) { return a - b; });
  }

  function settFelt(res, mal, verdi) {
    var deler = mal.split(".");
    res[deler[0]][deler[1]] = verdi;
  }

  function tomtResultat() {
    return {
      kunde: { navn: "", orgnr: "", kontaktperson: "", epost: "", telefon: "", fakturaadresse: "", fakturaEpost: "", fakturaMerkes: "" },
      sak: { saksnr: "", navn: "", ansvarligAdvokat: "", saksbehandler: "", sekretaer: "", motpart: "", timepris: null, salaerberegning: "", faktureringsrutiner: "", betalingsbetingelser: "" },
      advarsler: [],
      ukjenteFelt: {}
    };
  }

  /* ------------------------------------------------------------------ */
  /* Tekst-parser                                                        */
  /* ------------------------------------------------------------------ */

  function fraTekst(tekst) {
    var res = tomtResultat();
    var raa = {};            // mal -> [linjer]
    var tittelNavn = "";
    var sett = { etikett: false };

    var linjer = String(tekst == null ? "" : tekst)
      .replace(/\r\n?/g, "\n")
      .replace(/\u2028|\u2029/g, "\n")
      .split("\n")
      .map(function (l) {
        // Pandoc/markdown tabellkanter og rammer -> tom linje
        if (/^[\s\u00a0]*[-=+|_*]{3,}[-=+|_*\s]*$/.test(l)) return "";
        // Markdown-tabell «| a | b |»
        if (/^\s*\|.*\|\s*$/.test(l)) l = l.replace(/^\s*\|/, "").replace(/\|\s*$/, "").replace(/\s*\|\s*/g, "\t");
        // **fet** / __fet__
        return l.replace(/\*\*|__/g, "").replace(/[\s\u00a0]+$/, "");
      });

    var aktiv = null;        // { mal, linjer:[], ventet: bool (etikett uten verdi ennå), ukjent: navn }
    function avslutt() {
      if (!aktiv) return;
      if (aktiv.ukjent) {
        var v = rydd(aktiv.linjer.join(" "));
        if (v && !tomVerdi(v) && !(aktiv.ukjent in res.ukjenteFelt)) res.ukjenteFelt[aktiv.ukjent] = v;
      } else if (!(aktiv.mal in raa) || raa[aktiv.mal].join("").trim() === "") {
        raa[aktiv.mal] = aktiv.linjer.slice();
      }
      aktiv = null;
    }

    for (var i = 0; i < linjer.length; i++) {
      var linje = linjer[i];
      var tom = rydd(linje) === "";

      // Gruppeoverskrift i malen: «— Klient», «— Saken», «— Fakturering» -> avslutter feltet, er aldri etikett
      if (/^[\s\u00a0]*[—–][\s\u00a0]*\S/.test(linje) && linje.indexOf(":") < 0) {
        avslutt();
        continue;
      }

      if (!tittelNavn && !sett.etikett && !aktiv) {
        var t = TITTEL_RE.exec(linje);
        if (t) { tittelNavn = rydd(t[1]); continue; }
        if (/^[\s\u00a0]*oppdragsbekreftelse[\s\u00a0]*$/i.test(linje)) continue;
      }

      var e = tom ? null : finnEtikett(linje);
      if (e) {
        sett.etikett = true;
        avslutt();
        var v = rydd(e.verdi);
        aktiv = { mal: e.mal, linjer: v ? [v] : [], ventet: !v };
        continue;
      }

      if (!tom) {
        var u = UKJENT_RE.exec(linje);
        if (u && /[A-Za-zÆØÅæøå]/.test(u[1]) && !/^https?$/i.test(u[1]) && rydd(u[1]).split(" ").length <= 5 && (/\t/.test(linje) || /:\s*$/.test(linje) || /: {2,}/.test(linje))) {
          avslutt();
          var uv = rydd(u[2]);
          aktiv = { mal: null, ukjent: rydd(u[1]), linjer: uv ? [uv] : [], ventet: !uv };
          continue;
        }
      }

      if (!aktiv) continue;

      if (tom) {
        // Etikett uten verdi: hopp over tomme linjer (Word-copy/paste gir ofte etikett / tom / verdi).
        if (aktiv.ventet) continue;
        avslutt();
        continue;
      }

      aktiv.linjer.push(rydd(linje));
      aktiv.ventet = false;
    }
    avslutt();

    // Fyll resultat
    Object.keys(raa).forEach(function (mal) {
      var l = raa[mal].map(rydd).filter(function (x) { return x !== ""; });
      var verdi = FLERLINJE_KOMMA[mal] ? l.join(", ") : l.join(" ");
      verdi = rydd(fjernPlassholdere(verdi).replace(/\s+,/g, ","));
      if (tomVerdi(verdi)) verdi = "";
      if (mal === "sak.timepris") { res._timeprisTekst = verdi; return; }
      settFelt(res, mal, verdi);
    });

    // Kundenavn: eget felt (nytt format) ellers tittel (gammelt format)
    tittelNavn = rydd(fjernPlassholdere(tittelNavn));
    if (!res.kunde.navn && tittelNavn && !tomVerdi(tittelNavn)) res.kunde.navn = tittelNavn;
    res.kunde.navn = normaliserNavn(res.kunde.navn);

    // Org.nr.
    if (res.kunde.orgnr) {
      res.kunde.orgnr = renOrgnr(res.kunde.orgnr);
      if (!gyldigOrgnr(res.kunde.orgnr)) {
        res.advarsler.push("Org.nr. «" + res.kunde.orgnr + "» ser ikke gyldig ut (9 siffer med kontrollsiffer) – sjekk nummeret.");
      }
    }

    // E-post: plukk ut adressen hvis feltet har ekstra tekst
    ["epost", "fakturaEpost"].forEach(function (f) {
      var m = /[^\s<>()«»"',;:]+@[^\s<>()«»"',;:]+\.[a-z]{2,}/i.exec(res.kunde[f]);
      if (m) res.kunde[f] = m[0].replace(/\.$/, "");
    });

    // Timepris
    var tp = null;
    if (res._timeprisTekst) tp = timeprisFraFelt(res._timeprisTekst);
    if (tp == null && res.sak.salaerberegning) tp = timeprisFraSalaer(res.sak.salaerberegning);
    delete res._timeprisTekst;
    res.sak.timepris = tp;
    res.forslagTimepris = [];
    if (tp == null) {
      res.advarsler.push("Fant ingen timepris – fyll inn manuelt.");
    } else if (tp < TIMEPRIS_MIN || tp > TIMEPRIS_MAX) {
      res.forslagTimepris = forslagTimepris(tp);
      var tekst = "Timeprisen ser uvanlig ut – sjekk før lagring (kr " + kr(tp) + ", forventet kr " + kr(TIMEPRIS_MIN) + "–" + kr(TIMEPRIS_MAX) + ").";
      if (res.forslagTimepris.length) {
        var f = res.forslagTimepris.map(function (x) { return "kr " + kr(x); });
        tekst += " Mente du " + (f.length > 1 ? f.slice(0, -1).join(", ") + " eller " + f[f.length - 1] : f[0]) + "?";
      }
      res.advarsler.push(tekst);
    }

    if (!res.kunde.navn) res.advarsler.push("Fant ikke klientnavn – fyll inn manuelt.");
    if (!res.sak.navn) res.advarsler.push("Fant ikke sakens navn.");
    return res;
  }

  /* ------------------------------------------------------------------ */
  /* DOCX                                                                */
  /* ------------------------------------------------------------------ */

  var W_NS = {
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main": 1,
    "http://purl.oclc.org/ooxml/wordprocessingml/main": 1
  };
  function erW(node, navn) {
    return node && node.nodeType === 1 && W_NS[node.namespaceURI] === 1 && node.localName === navn;
  }
  function barn(node) {
    var out = [];
    for (var c = node.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) out.push(c);
    return out;
  }

  // Tekst fra et avsnitt (w:p): w:t over flere w:r slås sammen; w:tab -> \t; w:br -> \n.
  function avsnittTekst(p) {
    var ut = "";
    (function gaa(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (!W_NS[c.namespaceURI]) {
          // Tekstbokser o.l. ligger under andre navnerom (mc:, wps:, v:) – let etter w:txbxContent
          if (c.localName === "Fallback") continue; // unngå dobbel tekst fra mc:AlternateContent
          gaa(c);
          continue;
        }
        switch (c.localName) {
          case "t": ut += c.textContent; break;
          case "tab": ut += "\t"; break;
          case "br": case "cr": ut += "\n"; break;
          case "noBreakHyphen": ut += "-"; break;
          case "softHyphen": break;
          case "delText": case "instrText": case "rPr": case "pPr": case "fldData": break;
          case "del": case "moveFrom": break;          // slettet tekst (sporede endringer)
          case "sdt":
            if (!visesPlassholder(c)) gaa(c);
            break;
          case "sdtPr": case "sdtEndPr": break;
          case "p": ut += "\n" + avsnittTekst(c); break; // nestet (tekstboks)
          default: gaa(c);
        }
      }
    })(p);
    return ut;
  }

  function visesPlassholder(sdt) {
    var pr = barn(sdt).filter(function (c) { return erW(c, "sdtPr"); })[0];
    if (!pr) return false;
    return barn(pr).some(function (c) {
      if (!erW(c, "showingPlcHdr")) return false;
      var v = c.getAttribute("w:val") || "";
      return v !== "0" && v !== "false";
    });
  }

  function sdtInnhold(sdt) {
    if (visesPlassholder(sdt)) return [];
    var inn = barn(sdt).filter(function (c) { return erW(c, "sdtContent"); })[0];
    return inn ? barn(inn) : [];
  }

  // Flater ut blokker (avsnitt, tabeller, sdt) til linjer.
  function blokker(noder, linjer) {
    noder.forEach(function (n) {
      if (erW(n, "p")) linjer.push(avsnittTekst(n));
      else if (erW(n, "tbl")) tabell(n, linjer);
      else if (erW(n, "sdt")) blokker(sdtInnhold(n), linjer);
      else if (erW(n, "customXml") || erW(n, "smartTag") || erW(n, "ins") || erW(n, "moveTo")) blokker(barn(n), linjer);
    });
  }

  function rader(tbl) {
    var ut = [];
    barn(tbl).forEach(function (c) {
      if (erW(c, "tr")) ut.push(c);
      else if (erW(c, "sdt")) sdtInnhold(c).forEach(function (x) { if (erW(x, "tr")) ut.push(x); });
      else if (erW(c, "customXml")) barn(c).forEach(function (x) { if (erW(x, "tr")) ut.push(x); });
    });
    return ut;
  }

  function celler(tr) {
    var ut = [];
    barn(tr).forEach(function (c) {
      if (erW(c, "tc")) ut.push(c);
      else if (erW(c, "sdt")) sdtInnhold(c).forEach(function (x) { if (erW(x, "tc")) ut.push(x); });
      else if (erW(c, "customXml")) barn(c).forEach(function (x) { if (erW(x, "tc")) ut.push(x); });
    });
    return ut;
  }

  function celleTekst(tc) {
    var l = [];
    blokker(barn(tc), l);
    return l.join("\n").split("\n").map(function (x) { return x.replace(/\t/g, " ").replace(/[ \u00a0]+/g, " ").trim(); })
      .filter(function (x) { return x !== ""; });
  }

  function tabell(tbl, linjer) {
    rader(tbl).forEach(function (tr) {
      var cs = celler(tr).map(celleTekst);
      // fjern helt tomme celler i slutten
      while (cs.length && cs[cs.length - 1].length === 0) cs.pop();
      if (!cs.length) { linjer.push(""); return; }
      if (cs.length === 1) {
        linjer.push.apply(linjer, cs[0]); linjer.push(""); return;
      }
      var etikett = cs[0].join(" ").replace(/[\s:]+$/, "");
      var verdi = [];
      cs.slice(1).forEach(function (c) { verdi = verdi.concat(c); });
      if (!etikett) {
        linjer.push.apply(linjer, verdi);
      } else {
        // Tom celle -> «-» (tolkes som tom) så parseren ikke tar teksten etter tabellen som verdi
        linjer.push(etikett + ":\t" + (verdi[0] || "-"));
        for (var i = 1; i < verdi.length; i++) linjer.push("\t" + verdi[i]);
      }
      linjer.push("");
    });
  }

  function docxTilTekst(arrayBuffer) {
    var JSZipRef = root.JSZip;
    var DOMParserRef = root.DOMParser;
    if (!JSZipRef) return Promise.reject(new Error("JSZip er ikke lastet."));
    if (!DOMParserRef) return Promise.reject(new Error("DOMParser er ikke tilgjengelig."));
    return Promise.resolve()
      .then(function () { return JSZipRef.loadAsync(arrayBuffer); })
      .catch(function () {
        throw new Error("Kunne ikke lese fila. Er det et Word-dokument (.docx)? Eldre .doc-filer må lagres som .docx først.");
      })
      .then(function (zip) {
        var f = zip.file("word/document.xml");
        if (!f) throw new Error("Fila ser ikke ut som et Word-dokument (.docx) – fant ikke dokumentinnhold.");
        return f.async("string");
      })
      .then(function (xml) {
        var doc = new DOMParserRef().parseFromString(xml, "application/xml");
        if (!doc || !doc.documentElement || doc.getElementsByTagName("parsererror").length) {
          throw new Error("Klarte ikke å tolke innholdet i Word-dokumentet.");
        }
        var body = null;
        barn(doc.documentElement).forEach(function (c) { if (erW(c, "body")) body = c; });
        if (!body) throw new Error("Fant ikke dokumentinnhold i Word-fila.");
        var linjer = [];
        blokker(barn(body), linjer);
        return linjer.join("\n");
      });
  }

  function fraDocx(arrayBuffer) {
    return docxTilTekst(arrayBuffer).then(fraTekst);
  }

  root.Oppdrag = {
    fraTekst: fraTekst,
    fraDocx: fraDocx,
    docxTilTekst: docxTilTekst,
    gyldigOrgnr: gyldigOrgnr,
    TIMEPRIS_MIN: TIMEPRIS_MIN,
    TIMEPRIS_MAX: TIMEPRIS_MAX,
    forslagTimepris: forslagTimepris,
    normaliserNavn: normaliserNavn
  };
})(typeof window !== "undefined" ? window : this);
