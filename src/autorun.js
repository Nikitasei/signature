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
  "kedzior.tomas@rifas.cz": {
    "firstName": "Tom\u00e1\u0161",
    "lastName": "Kedzior",
    "title": "\u010dlen spr\u00e1vn\u00ed rady",
    "phone": "+420 773 745 845",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/06/kedzior.tomas_.orez_-1.png"
  },
  "plch.adam@rifas.cz": {
    "firstName": "Adam",
    "lastName": "Plch",
    "title": "risk mana\u017eer",
    "phone": "+420 773 745 592",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/05/plch.adam_.png"
  },
  "kresta.michal@rifas.cz": {
    "firstName": "Michal",
    "lastName": "Kresta",
    "title": "p\u0159edseda spr\u00e1vn\u00ed rady",
    "phone": "+420 770 175 948",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/05/kresta.michal.png"
  },
  "dufek.pavel@rifas.cz": {
    "firstName": "Pavel",
    "lastName": "Dufek",
    "title": "region\u00e1ln\u00ed mana\u017eer",
    "phone": "+420 776 152 877",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2025/01/pan_dufek_kruh.png"
  },
  "hrebicek.pavel@rifas.cz": {
    "firstName": "Pavel",
    "lastName": "H\u0159eb\u00ed\u010dek",
    "title": "region\u00e1ln\u00ed \u0159editel Morava",
    "phone": "+420 605 252 405",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/10/hrebicek.pavel_.ctverec.png"
  },
  "pindej.miroslav@rifas.cz": {
    "firstName": "Miroslav",
    "lastName": "Pindej",
    "title": "region\u00e1ln\u00ed mana\u017eer",
    "phone": "+421 944 707 214",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2024/12/kulicka_fotka_email.png"
  },
  "kovalcikova.eliska@rifas.cz": {
    "firstName": "Eli\u0161ka",
    "lastName": "Koval\u010d\u00edkov\u00e1",
    "title": "oblastn\u00ed z\u00e1stupce",
    "phone": "+420 775 370 680",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2023/11/Kovalcikova_2023.png"
  },
  "stepanek.frantisek@rifas.cz": {
    "firstName": "Franti\u0161ek",
    "lastName": "\u0160t\u011bp\u00e1nek",
    "title": "oblastn\u00ed z\u00e1stupce",
    "phone": "+420 725 353 099",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2025/02/stepanek_frantisek.png"
  },
  "jancovicova.kamila@rifas.cz": {
    "firstName": "Kamila",
    "lastName": "Jan\u010dovi\u010dov\u00e1",
    "title": "realitn\u00ed specialista",
    "phone": "+420 777 132 013",
    "street": "Na St\u0159elnici 336/6",
    "postalCode": "779 00",
    "city": "Olomouc",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2025/01/pani_jancovicova_kruh.png"
  },
  "krajcova.jolana@rifas.cz": {
    "firstName": "Jolana",
    "lastName": "Kraj\u010dov\u00e1",
    "title": "realitn\u00ed mana\u017eer",
    "phone": "+420 608 713 758",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2024/05/Krajcova_RM.png"
  },
  "navrat.petr@rifas.cz": {
    "firstName": "Petr",
    "lastName": "N\u00e1vrat",
    "title": "region\u00e1ln\u00ed \u0159editel \u010cechy",
    "phone": "+420 725 293 467",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha ",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2023/11/petr_navrat_2023.png"
  },
  "pradeniak.andrej@rifas.cz": {
    "firstName": "Andrej",
    "lastName": "Pradeniak",
    "title": "oblastn\u00ed z\u00e1stupce",
    "phone": "+420 733 747 571",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2023/10/pradeniak_2023.png"
  },
  "volfova.klara@rifas.cz": {
    "firstName": "Kl\u00e1ra",
    "lastName": "Volfov\u00e1",
    "title": "region\u00e1ln\u00ed mana\u017eer",
    "phone": "+420 602 320 528",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/06/volfova.klara_.orez_-1.png"
  },
  "zajicek.radim@rifas.cz": {
    "firstName": "Radim",
    "lastName": "Zaj\u00ed\u010dek",
    "title": "oblastn\u00ed z\u00e1stupce",
    "phone": "+420 725 425 738",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/06/zajicek.radim_-2.jpg"
  },
  "dluhopisy@rifas.cz": {
    "shared": true
  },
  "jarosik.jan@rifas.cz": {
    "firstName": "Jan",
    "lastName": "Jaro\u0161\u00edk",
    "title": "oblastn\u00ed z\u00e1stupce",
    "phone": "+420 774 189 974\u202c",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/09/jarosik.jan_.orez_.png"
  },
  "kroupa.oldrich@rifas.cz": {
    "firstName": "Old\u0159ich",
    "lastName": "Kroupa",
    "title": "\u0159editel komer\u010dn\u00edho odd\u011blen\u00ed",
    "phone": "+420 733 551 960",
    "street": "Kom\u00e1rovsk\u00e1 263/20a",
    "postalCode": "617 00",
    "city": "Brno",
    "photo": ""
  },
  "holecek.martin@rifas.cz": {
    "firstName": "Martin",
    "lastName": "Hole\u010dek",
    "title": "oblastn\u00ed z\u00e1stupce",
    "phone": "+420 704 773 268",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2025/08/Holecek_emailovy_podpis.png"
  },
  "grep.test@rifas.cz": {
    "firstName": "test",
    "lastName": "grep",
    "title": "testing user",
    "phone": "+420 123 456 789",
    "street": "Kom\u00e1rovsk\u00e1 263/20a",
    "postalCode": "617 00",
    "city": "Brno",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/10/Screenshot-2026-10-23-at-14-55-01.png"
  },
  "pospisil.pavel@rifas.cz": {
    "firstName": "Pavel",
    "lastName": "Posp\u00ed\u0161il",
    "title": "oblastn\u00ed z\u00e1stupce RIF",
    "phone": "+420 607 612 471",
    "street": "Kom\u00e1rovsk\u00e1 263/20a",
    "postalCode": "617 00",
    "city": "Brno",
    "photo": ""
  },
  "rifas@rifas.cz": {
    "shared": true
  },
  "info@rifas.cz": {
    "shared": true
  },
  "drabkova.jana-backup@rifas.cz": {
    "shared": true
  },
  "dufek.pavel-backup@rifas.cz": {
    "shared": true
  },
  "francova.ludmila-backup@rifas.cz": {
    "shared": true
  },
  "holecek.martin-backup@rifas.cz": {
    "shared": true
  },
  "hrebicek.pavel-backup@rifas.cz": {
    "shared": true
  },
  "info-backup@rifas.cz": {
    "shared": true
  },
  "jancovicova.kamila-backup@rifas.cz": {
    "shared": true
  },
  "kedzior.tomas-backup@rifas.cz": {
    "shared": true
  },
  "kovalcikova.eliska-backup@rifas.cz": {
    "shared": true
  },
  "krajcova.jolana-backup@rifas.cz": {
    "shared": true
  },
  "kresta.michal-backup@rifas.cz": {
    "shared": true
  },
  "navrat.petr-backup@rifas.cz": {
    "shared": true
  },
  "pindej.miroslav-backup@rifas.cz": {
    "shared": true
  },
  "plch.adam-backup@rifas.cz": {
    "shared": true
  },
  "pradeniak.andrej-backup@rifas.cz": {
    "shared": true
  },
  "rifas-backup@rifas.cz": {
    "shared": true
  },
  "volfova.klara-backup@rifas.cz": {
    "shared": true
  },
  "vrzal.david-backup@rifas.cz": {
    "shared": true
  },
  "stepanek.frantisek-backup@rifas.cz": {
    "shared": true
  }
};
var RIF_PEOPLE = {
  "kedzior.tomas@rifas.cz": {
    "firstName": "Tom\u00e1\u0161",
    "lastName": "Kedzior",
    "title": "\u010dlen spr\u00e1vn\u00ed rady",
    "phone": "+420 773 745 845",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/06/kedzior.tomas_.orez_-1.png"
  },
  "plch.adam@rifas.cz": {
    "firstName": "Adam",
    "lastName": "Plch",
    "title": "risk mana\u017eer",
    "phone": "+420 773 745 592",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/05/plch.adam_.png"
  },
  "kresta.michal@rifas.cz": {
    "firstName": "Michal",
    "lastName": "Kresta",
    "title": "p\u0159edseda spr\u00e1vn\u00ed rady",
    "phone": "+420 770 175 948",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/05/kresta.michal.png"
  },
  "dufek.pavel@rifas.cz": {
    "firstName": "Pavel",
    "lastName": "Dufek",
    "title": "region\u00e1ln\u00ed mana\u017eer",
    "phone": "+420 776 152 877",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2025/01/pan_dufek_kruh.png"
  },
  "hrebicek.pavel@rifas.cz": {
    "firstName": "Pavel",
    "lastName": "H\u0159eb\u00ed\u010dek",
    "title": "region\u00e1ln\u00ed \u0159editel Morava",
    "phone": "+420 605 252 405",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/10/hrebicek.pavel_.ctverec.png"
  },
  "pindej.miroslav@rifas.cz": {
    "firstName": "Miroslav",
    "lastName": "Pindej",
    "title": "region\u00e1ln\u00ed mana\u017eer",
    "phone": "+421 944 707 214",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2024/12/kulicka_fotka_email.png"
  },
  "kovalcikova.eliska@rifas.cz": {
    "firstName": "Eli\u0161ka",
    "lastName": "Koval\u010d\u00edkov\u00e1",
    "title": "oblastn\u00ed z\u00e1stupce",
    "phone": "+420 775 370 680",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2023/11/Kovalcikova_2023.png"
  },
  "stepanek.frantisek@rifas.cz": {
    "firstName": "Franti\u0161ek",
    "lastName": "\u0160t\u011bp\u00e1nek",
    "title": "oblastn\u00ed z\u00e1stupce",
    "phone": "+420 725 353 099",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2025/02/stepanek_frantisek.png"
  },
  "jancovicova.kamila@rifas.cz": {
    "firstName": "Kamila",
    "lastName": "Jan\u010dovi\u010dov\u00e1",
    "title": "realitn\u00ed specialista",
    "phone": "+420 777 132 013",
    "street": "Na St\u0159elnici 336/6",
    "postalCode": "779 00",
    "city": "Olomouc",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2025/01/pani_jancovicova_kruh.png"
  },
  "krajcova.jolana@rifas.cz": {
    "firstName": "Jolana",
    "lastName": "Kraj\u010dov\u00e1",
    "title": "realitn\u00ed mana\u017eer",
    "phone": "+420 608 713 758",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2024/05/Krajcova_RM.png"
  },
  "navrat.petr@rifas.cz": {
    "firstName": "Petr",
    "lastName": "N\u00e1vrat",
    "title": "region\u00e1ln\u00ed \u0159editel \u010cechy",
    "phone": "+420 725 293 467",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha ",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2023/11/petr_navrat_2023.png"
  },
  "pradeniak.andrej@rifas.cz": {
    "firstName": "Andrej",
    "lastName": "Pradeniak",
    "title": "oblastn\u00ed z\u00e1stupce",
    "phone": "+420 733 747 571",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2023/10/pradeniak_2023.png"
  },
  "volfova.klara@rifas.cz": {
    "firstName": "Kl\u00e1ra",
    "lastName": "Volfov\u00e1",
    "title": "region\u00e1ln\u00ed mana\u017eer",
    "phone": "+420 602 320 528",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/06/volfova.klara_.orez_-1.png"
  },
  "zajicek.radim@rifas.cz": {
    "firstName": "Radim",
    "lastName": "Zaj\u00ed\u010dek",
    "title": "oblastn\u00ed z\u00e1stupce",
    "phone": "+420 725 425 738",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/06/zajicek.radim_-2.jpg"
  },
  "jarosik.jan@rifas.cz": {
    "firstName": "Jan",
    "lastName": "Jaro\u0161\u00edk",
    "title": "oblastn\u00ed z\u00e1stupce",
    "phone": "+420 774 189 974\u202c",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/09/jarosik.jan_.orez_.png"
  },
  "kroupa.oldrich@rifas.cz": {
    "firstName": "Old\u0159ich",
    "lastName": "Kroupa",
    "title": "\u0159editel komer\u010dn\u00edho odd\u011blen\u00ed",
    "phone": "+420 733 551 960",
    "street": "Kom\u00e1rovsk\u00e1 263/20a",
    "postalCode": "617 00",
    "city": "Brno",
    "photo": ""
  },
  "holecek.martin@rifas.cz": {
    "firstName": "Martin",
    "lastName": "Hole\u010dek",
    "title": "oblastn\u00ed z\u00e1stupce",
    "phone": "+420 704 773 268",
    "street": "Vinohradsk\u00e1 2516/28",
    "postalCode": "120 00",
    "city": "Praha",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2025/08/Holecek_emailovy_podpis.png"
  },
  "grep.test@rifas.cz": {
    "firstName": "test",
    "lastName": "grep",
    "title": "testing user",
    "phone": "+420 123 456 789",
    "street": "Kom\u00e1rovsk\u00e1 263/20a",
    "postalCode": "617 00",
    "city": "Brno",
    "photo": "https://www.allrisk.cz/wp-content/uploads/2026/10/Screenshot-2026-10-23-at-14-55-01.png"
  },
  "pospisil.pavel@rifas.cz": {
    "firstName": "Pavel",
    "lastName": "Posp\u00ed\u0161il",
    "title": "oblastn\u00ed z\u00e1stupce RIF",
    "phone": "+420 607 612 471",
    "street": "Kom\u00e1rovsk\u00e1 263/20a",
    "postalCode": "617 00",
    "city": "Brno",
    "photo": ""
  }
};
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
  // Vypne osobni podpis z Outlooku, aby nebyly dva.
  if (item.disableClientSignatureAsync) {
    item.disableClientSignatureAsync(function () { setSig(); });
  } else {
    setSig();
  }
}

function removeRif(done) {
  var item = Office.context.mailbox.item;
  item.body.setSignatureAsync("", { coercionType: Office.CoercionType.Html }, function () {
    if (item.internetHeaders && item.internetHeaders.removeAsync) {
      item.internetHeaders.removeAsync([SIG_HEADER], function () { done(); });
    } else {
      done();
    }
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
