/*
 * RIF podpis - Outlook add-in (event-based activation)
 *
 * - OnNewMessageCompose: nova zprava / odpoved / preposlani -> vlozi RIF podpis,
 *   pokud je odesilatel (From) adresa @rifas.cz ze seznamu RIF_USERS.
 * - OnMessageFromChanged: uzivatel prepne From -> podpis se vymeni nebo odebere.
 * - Nastavi hlavicku X-RIF-Signature: addin, aby transport rule podpis nepridala podruhe.
 *
 * DULEZITE: klasicky Outlook pro Windows nacita JEN tento soubor (bez HTML),
 * proto jsou data i kod v jednom souboru a nic se neimportuje.
 */

// USERS-BEGIN  (generuje export-users.ps1 - rucne needitovat)
var RIF_USERS = {
  "hrebicek.pavel@rifas.cz": {
    "firstName": "Pavel",
    "lastName": "H\u0159eb\u00ed\u010dek",
    "title": "region\u00e1ln\u00ed \u0159editel Morava",
    "phone": "+420 605 252 405",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": ""
  }
};
var RIF_PEOPLE = {};
// USERS-END

/*
 * Co vlozit pri psani ze SDILENE schranky @rifas.cz:
 *   "person" = podpis toho, kdo pise (jmeno, funkce, telefon, fotka), ale s adresou sdilene schranky
 *   "none"   = zadny podpis
 * Kdo neni v RIF_PEOPLE (napr. clovek z jine firmy), nedostane nic.
 */
var SHARED_MODE = "person";

var SIG_HEADER = "X-RIF-Signature";

function esc(v) {
  return String(v == null ? "" : v)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function line(style, content) {
  return '<div style="font-family:tahoma,sans-serif;font-size:14px;' + style + '">' + content + "</div>";
}

function buildSignature(u, email) {
  var name = esc((u.firstName || "") + " " + (u.lastName || "")).replace(/^\s+|\s+$/g, "");
  var photo = u.photo
    ? '<img src="' + esc(u.photo) + '" width="160" height="160" alt="Fotka uzivatele"><br>'
    : "";
  var h = "";
  h += "<strong>S pozdravem</strong>";
  h += '<table style="padding:0;margin:10px 0;border:none;max-width:560px;"><tbody><tr>';
  h += '<td style="vertical-align:middle;padding:0 16px 0 0;width:160px;">';
  h += photo;
  h += '<a href="https://www.rifas.cz" target="_blank" style="text-decoration:none;"><img src="https://www.allrisk.cz/wp-content/uploads/2023/03/rifas_logo_2023.png" width="145" height="37" alt="Rifas" style="padding-top:8px;border:0;"></a>';
  h += "</td>";
  h += '<td style="border-left:2px solid #e8e8e8;padding:0 0 0 16px;max-width:480px;">';
  h += line("line-height:16px;font-weight:bold;color:black;", name);
  if (u.title) h += line("line-height:16px;color:black;margin-bottom:10px;", esc(u.title));
  if (u.phone) h += line("line-height:14px;color:#134865;margin-bottom:4px;", "<strong>" + esc(u.phone) + "</strong>");
  h += line(
    "line-height:14px;margin-bottom:12px;",
    '<strong><a href="mailto:' + esc(email) + '" target="_blank" style="text-decoration:none;color:#134865;">' + esc(email) + "</a></strong>"
  );
  h += line("line-height:16px;font-weight:bold;color:black;", "Realitn&iacute; &amp; investorsk&yacute; fond, a.s.");
  h += line(
    "line-height:16px;color:black;",
    esc(u.street) + "<br>" + esc(u.postalCode) + " " + esc(u.city) + "<br><br>" +
      '<a href="https://www.rifas.cz" target="_blank" style="text-decoration:none;color:#134865;">www.rifas.cz</a>'
  );
  h += '<div style="margin-top:12px;">';
  h += '<a href="https://www.facebook.com/realitniainvestorskyfondcz/" target="_blank" style="text-decoration:none;"><img src="https://www.allrisk.cz/wp-content/uploads/2022/08/facebook_podpis.png" width="24" height="24" alt="Facebook" style="border:0;"></a>&nbsp;&nbsp;';
  h += '<a href="https://www.linkedin.com/company/realitn%C3%AD-investi%C4%8Dn%C3%AD-fond-a-s" target="_blank" style="text-decoration:none;"><img src="https://www.allrisk.cz/wp-content/uploads/2022/08/linkedin_podpis.png" width="24" height="24" alt="LinkedIn" style="border:0;"></a>&nbsp;&nbsp;';
  h += '<a href="https://www.youtube.com/channel/UCdi6wcix_AxllcRbfs27VMA" target="_blank" style="text-decoration:none;"><img src="https://www.allrisk.cz/wp-content/uploads/2022/08/youtube_podpis.png" width="24" height="24" alt="YouTube" style="border:0;"></a>';
  h += "</div>";
  h += "</td></tr></tbody></table>";
  return h;
}

/* Zjisti adresu v poli From (u aliasu / sdilene schranky se lisi od prihlaseneho uzivatele). */
function getFromAddress(cb) {
  var item = Office.context.mailbox.item;
  if (item.from && item.from.getAsync) {
    item.from.getAsync(function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded && r.value && r.value.emailAddress) {
        cb(r.value.emailAddress.toLowerCase());
      } else {
        cb(Office.context.mailbox.userProfile.emailAddress.toLowerCase());
      }
    });
  } else {
    cb(Office.context.mailbox.userProfile.emailAddress.toLowerCase());
  }
}

/* Najde udaje do podpisu pro adresu v From; user = null -> podpis nevkladat, reason = proc */
function resolveSigner(fromEmail) {
  var entry = RIF_USERS[fromEmail];
  if (!entry) return { user: null, reason: fromEmail + " neni v seznamu RIF (spust export / je ve vyjimkach)." };
  // Zaznam bez jmena = sdilena / spolecna schranka (i kdyz je vedena jako UserMailbox)
  var isShared = entry.shared || (!entry.firstName && !entry.lastName);
  if (!isShared) return { user: entry, email: fromEmail };
  if (SHARED_MODE !== "person") return { user: null, reason: "podpis pro sdilene schranky je vypnuty (SHARED_MODE)." };
  var me = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
  var person = RIF_PEOPLE[me] || RIF_USERS[me];
  // Ze sdilene schranky: udaje i e-mail cloveka, ktery pise
  if (person && !person.shared && (person.firstName || person.lastName)) return { user: person, email: me };
  return { user: null, reason: fromEmail + " je sdilena; prihlaseny " + me + " neni mezi lidmi RIF." };
}

function applyRif(user, email, done) {
  var item = Office.context.mailbox.item;
  var setSig = function () {
    item.body.setSignatureAsync(buildSignature(user, email), { coercionType: Office.CoercionType.Html }, function () {
      var hdr = {};
      hdr[SIG_HEADER] = "addin";
      if (item.internetHeaders && item.internetHeaders.setAsync) {
        item.internetHeaders.setAsync(hdr, function () { done(); });
      } else {
        done();
      }
    });
  };
  // POZOR: disableClientSignatureAsync se NEPOUZIVA. V klasickem Outlooku a na Macu
  // trvale nastavi vychozi podpis celeho uctu na (none) a rozbije tak podpisy
  // Allrisk Reality lidem, kteri pisou z obou adres. setSignatureAsync podpis
  // v oblasti podpisu sam nahradi.
  setSig();
}

/* Odebere podpis jen tehdy, kdyz ho do zpravy vlozil tento doplnek (pozna se podle hlavicky).
   Cizi podpis (napr. Allrisk Reality z Outlooku) nechava byt. */
function removeRif(done) {
  var item = Office.context.mailbox.item;
  var clear = function () {
    item.body.setSignatureAsync("", { coercionType: Office.CoercionType.Html }, function () {
      item.internetHeaders.removeAsync([SIG_HEADER], function () { done(); });
    });
  };
  if (!item.internetHeaders || !item.internetHeaders.getAsync) { done(); return; }
  item.internetHeaders.getAsync([SIG_HEADER], function (r) {
    var ours = r.status === Office.AsyncResultStatus.Succeeded && r.value && r.value[SIG_HEADER];
    if (ours) { clear(); } else { done(); }
  });
}

/* Informacni lista nad zpravou (jen pro rucni tlacitko, at je videt proc se nic nestalo) */
function notify(text, done) {
  var item = Office.context.mailbox.item;
  try {
    item.notificationMessages.replaceAsync("rifInfo", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: String(text).substring(0, 150),
      icon: "Icon16",
      persistent: false
    }, function () { done(); });
  } catch (e) {
    done();
  }
}

/* Spolecna logika; verbose = ukazat duvod, kdyz se podpis nevlozi */
function run(event, verbose) {
  var finished = false;
  var done = function () {
    if (!finished) { finished = true; event.completed(); }
  };
  try {
    getFromAddress(function (email) {
      try {
        var r = resolveSigner(email);
        var user = r.user;
        if (!user) {
          // Neni RIF identita / vyjimka / sdilena schranka bez podpisu -> nesahame na nic (podpis resi transport rule / Outlook).
          if (verbose) { notify("RIF podpis: " + r.reason, done); }
          else { done(); }
          return;
        }
        applyRif(user, r.email, done);
      } catch (e) {
        if (verbose) { notify("RIF podpis - chyba: " + e.message, done); } else { done(); }
      }
    });
  } catch (e) {
    if (verbose) { notify("RIF podpis - chyba: " + e.message, done); } else { done(); }
  }
}

/* Nova zprava, odpoved, preposlani */
function onNewMessageCompose(event) {
  run(event, false);
}

/* Uzivatel prepnul From */
function onMessageFromChanged(event) {
  var finished = false;
  var done = function () {
    if (!finished) { finished = true; event.completed(); }
  };
  try {
    getFromAddress(function (email) {
      try {
        var r = resolveSigner(email);
        if (r.user) { applyRif(r.user, r.email, done); } else { removeRif(done); }
      } catch (e) { done(); }
    });
  } catch (e) { done(); }
}

/* Tlacitko na pase karet - rucni vlozeni (napr. kdyz uzivatel podpis smaze) */
function insertRifSignature(event) {
  run(event, true);
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
  Office.actions.associate("onMessageFromChanged", onMessageFromChanged);
  Office.actions.associate("insertRifSignature", insertRifSignature);
}

if (typeof module !== "undefined") {
  module.exports = { buildSignature: buildSignature, RIF_USERS: RIF_USERS, RIF_PEOPLE: RIF_PEOPLE };
}
