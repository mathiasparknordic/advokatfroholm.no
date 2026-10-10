/*
 * Microsoft Graph: OneDrive-filer og e-post.
 * Global: window.Graph
 *
 * Alle stier er relative til OneDrive-roten, f.eks. "Timeføring/data.json".
 */
(function () {
  "use strict";

  var BASE = "https://graph.microsoft.com/v1.0";

  function GraphFeil(status, kode, melding) {
    var e = new Error(melding || ("Graph-feil " + status));
    e.name = "GraphFeil";
    e.status = status;
    e.kode = kode || "";
    return e;
  }

  function stiUrl(sti) {
    return "/me/drive/root:/" + String(sti).split("/").filter(Boolean).map(encodeURIComponent).join("/");
  }

  function vent(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  /** Kjerneforespørsel med token og enkel retry ved 429/503/504. */
  async function kall(metode, url, opts) {
    opts = opts || {};
    var forsok = 0;
    while (true) {
      var token = await window.Auth.token();
      var headers = Object.assign({ Authorization: "Bearer " + token }, opts.headers || {});
      var body = opts.body;
      if (opts.json !== undefined) {
        headers["Content-Type"] = "application/json";
        body = JSON.stringify(opts.json);
      }
      var res;
      try {
        res = await fetch(url.indexOf("https://") === 0 ? url : BASE + url, {
          method: metode, headers: headers, body: body, cache: "no-store"
        });
      } catch (nett) {
        throw GraphFeil(0, "nettverk", "Fikk ikke kontakt med Microsoft 365. Sjekk nettforbindelsen.");
      }
      if ((res.status === 429 || res.status === 503 || res.status === 504) && forsok < 3) {
        forsok++;
        var ra = parseInt(res.headers.get("Retry-After") || "0", 10);
        await vent(Math.min(Math.max(ra * 1000, 1000 * forsok), 15000));
        continue;
      }
      if (!res.ok) {
        var kode = "", melding = "";
        try {
          var j = await res.json();
          kode = (j && j.error && j.error.code) || "";
          melding = (j && j.error && j.error.message) || "";
        } catch (e) { /* ingen JSON */ }
        throw GraphFeil(res.status, kode, melding || ("Microsoft 365 svarte med feil " + res.status + "."));
      }
      return res;
    }
  }

  /** Metadata for en fil/mappe, eller null om den ikke finnes. */
  async function metadata(sti) {
    try {
      var res = await kall("GET", stiUrl(sti));
      return await res.json();
    } catch (e) {
      if (e.status === 404) return null;
      throw e;
    }
  }

  /** Leser en JSON-fil. Returnerer { data, eTag, lastModified } eller null om fila ikke finnes. */
  async function lesJson(sti) {
    var meta = await metadata(sti);
    if (!meta) return null;
    var url = meta["@microsoft.graph.downloadUrl"];
    var tekst;
    if (url) {
      // Forhåndsautentisert, kortlevd URL – sendes uten Authorization-header.
      var r = await fetch(url, { cache: "no-store" });
      if (!r.ok) throw GraphFeil(r.status, "", "Kunne ikke laste ned " + sti + ".");
      tekst = await r.text();
    } else {
      var r2 = await kall("GET", stiUrl(sti) + ":/content");
      tekst = await r2.text();
    }
    var data;
    try { data = JSON.parse(tekst.replace(/^﻿/, "")); }
    catch (e) { throw GraphFeil(422, "ugyldigJson", "Fila " + sti + " er ikke gyldig JSON."); }
    return { data: data, eTag: meta.eTag, lastModified: meta.lastModifiedDateTime };
  }

  /** Laster ned en fil som ArrayBuffer. Kaster 404 om den ikke finnes. */
  async function lesFil(sti) {
    var meta = await metadata(sti);
    if (!meta) throw GraphFeil(404, "itemNotFound", "Fila " + sti + " finnes ikke.");
    var url = meta["@microsoft.graph.downloadUrl"];
    if (url) {
      var r = await fetch(url, { cache: "no-store" });
      if (!r.ok) throw GraphFeil(r.status, "", "Kunne ikke laste ned " + sti + ".");
      return await r.arrayBuffer();
    }
    var r2 = await kall("GET", stiUrl(sti) + ":/content");
    return await r2.arrayBuffer();
  }

  /**
   * Laster opp innhold (≤ 250 MB) til en sti. Mapper opprettes automatisk.
   * opts.ifMatch: eTag som må stemme (412 ellers). opts.nyFil: feiler med 409 om fila finnes.
   * Returnerer driveItem (med eTag).
   */
  async function skriv(sti, innhold, opts) {
    opts = opts || {};
    var headers = { "Content-Type": opts.contentType || "application/octet-stream" };
    if (opts.ifMatch) headers["If-Match"] = opts.ifMatch;
    var url = stiUrl(sti) + ":/content";
    if (opts.nyFil) url += "?@microsoft.graph.conflictBehavior=fail";
    var res = await kall("PUT", url, { headers: headers, body: innhold });
    return await res.json();
  }

  function skrivJson(sti, data, opts) {
    var o = Object.assign({ contentType: "application/json" }, opts || {});
    return skriv(sti, JSON.stringify(data, null, 2), o);
  }

  /** Sletter en fil. 404 regnes som allerede slettet. */
  async function slett(sti, ifMatch) {
    try {
      await kall("DELETE", stiUrl(sti), { headers: ifMatch ? { "If-Match": ifMatch } : {} });
      return true;
    } catch (e) {
      if (e.status === 404) return false;
      throw e;
    }
  }

  /** Sørger for at en mappe finnes (oppretter hele stien ved behov). Returnerer driveItem. */
  async function sikreMappe(sti) {
    var meta = await metadata(sti);
    if (meta) return meta;
    var deler = String(sti).split("/").filter(Boolean);
    var navn = deler.pop();
    var forelder = deler.length ? await sikreMappe(deler.join("/")) : null;
    var url = forelder ? "/me/drive/items/" + encodeURIComponent(forelder.id) + "/children" : "/me/drive/root/children";
    try {
      var res = await kall("POST", url, {
        json: { name: navn, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }
      });
      return await res.json();
    } catch (e) {
      if (e.status === 409) return await metadata(sti);
      throw e;
    }
  }

  /** Lister innholdet i en mappe (alle sider). Tom liste om mappen ikke finnes. */
  async function listMappe(sti) {
    var url = stiUrl(sti) + ":/children?$select=id,name,eTag,lastModifiedDateTime,size,file,folder&$top=200";
    var ut = [];
    try {
      while (url) {
        var res = await kall("GET", url);
        var j = await res.json();
        ut = ut.concat(j.value || []);
        url = j["@odata.nextLink"] || null;
      }
    } catch (e) {
      if (e.status === 404) return [];
      throw e;
    }
    return ut;
  }

  /**
   * Flytter en fil til en annen mappe med nytt navn. Finnes navnet fra før,
   * legges « (2)», « (3)» … til. Returnerer driveItem.
   */
  async function flytt(sti, tilMappe, nyttNavn, ifMatch) {
    var kilde = await metadata(sti);
    if (!kilde) throw GraphFeil(404, "itemNotFound", "Fila " + sti + " finnes ikke.");
    var mappe = await sikreMappe(tilMappe);
    var punkt = nyttNavn.lastIndexOf(".");
    var stamme = punkt > 0 ? nyttNavn.slice(0, punkt) : nyttNavn;
    var ext = punkt > 0 ? nyttNavn.slice(punkt) : "";
    for (var n = 1; n <= 20; n++) {
      var navn = n === 1 ? nyttNavn : stamme + " (" + n + ")" + ext;
      try {
        var res = await kall("PATCH", "/me/drive/items/" + encodeURIComponent(kilde.id), {
          headers: ifMatch ? { "If-Match": ifMatch } : {},
          json: { parentReference: { id: mappe.id }, name: navn }
        });
        return await res.json();
      } catch (e) {
        if (e.status === 409) continue; // navnet er opptatt
        throw e;
      }
    }
    throw GraphFeil(409, "nameAlreadyExists", "Fant ikke ledig filnavn i " + tilMappe + ".");
  }

  function tilBase64(buffer) {
    var bytes = new Uint8Array(buffer);
    var bin = "";
    var bit = 0x8000;
    for (var i = 0; i < bytes.length; i += bit) {
      bin += String.fromCharCode.apply(null, bytes.subarray(i, i + bit));
    }
    return btoa(bin);
  }

  /**
   * Sender e-post fra innlogget bruker.
   * { til, kopi, emne, tekst, vedlegg: { navn, innhold (ArrayBuffer), type } }
   */
  async function sendEpost(m) {
    function mottakere(s) {
      return String(s || "").split(/[;,\s]+/).filter(Boolean).map(function (a) {
        return { emailAddress: { address: a } };
      });
    }
    var melding = {
      subject: m.emne,
      body: { contentType: "Text", content: m.tekst || "" },
      toRecipients: mottakere(m.til),
      ccRecipients: mottakere(m.kopi)
    };
    if (m.vedlegg) {
      melding.attachments = [{
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: m.vedlegg.navn,
        contentType: m.vedlegg.type || "application/octet-stream",
        contentBytes: tilBase64(m.vedlegg.innhold)
      }];
    }
    await kall("POST", "/me/sendMail", { json: { message: melding, saveToSentItems: true } });
    return true;
  }

  async function meg() {
    var res = await kall("GET", "/me?$select=displayName,mail,userPrincipalName");
    return await res.json();
  }

  window.Graph = {
    metadata: metadata,
    lesJson: lesJson,
    lesFil: lesFil,
    skriv: skriv,
    skrivJson: skrivJson,
    slett: slett,
    sikreMappe: sikreMappe,
    listMappe: listMappe,
    flytt: flytt,
    sendEpost: sendEpost,
    meg: meg,
    tilBase64: tilBase64
  };
})();
