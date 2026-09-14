/* global Word, Office */

// ---------------------------------------------------------------------------
// Shared action implementations for the Gemeente Gooise Meren add-in.
// Used by both the ribbon commands (src/commands/commands.js) and the
// task pane (src/taskpane/taskpane.js), so the logic exists in one place.
//
// v4 (14 sep) - deliberately small rebuild. Scope is now just:
//   Titel, Standaardtekst, Koptekst 1-6 (zonder nummering),
//   Koptekst 1-4 (met nummering).
// Everything else (Klembord, Tabelstijlen, Taal, Bijhouden, Bijlage) has
// been removed, and so has all reliance on Word's *built-in* styles
// (Heading 1-6, Normal). Every style here is a plain custom paragraph
// style, created once via context.document.addStyle(...) and (re)applied
// every time its button is used. That sidesteps the two bug classes that
// caused real trouble in the previous, bigger build:
//   - resolving a built-in style's *localized* name before it can be
//     looked up and restyled (Word Online may report "Kop 1" rather than
//     "Heading 1", depending on display language) - not needed anymore,
//     since every style here has one name we chose ourselves;
//   - Word.Style not actually exposing a settable `paragraphFormat` in
//     Word Online's runtime (confirmed live: `style.paragraphFormat.set()`
//     throws "Cannot read properties of undefined (reading 'set')") - not
//     relevant either, since none of these styles need spacing set on the
//     style itself; only font properties (name/size/bold/italic/color) are
//     used, and `Word.Style.font` is reliably supported.
//
// v4.1 (14 sep) - added the gemeente brand color (#441D42) on Titel and
// every Koptekst (genummerd and ongenummerd); Standaardtekst stays black.
// Still only `.font` being set, so this doesn't reintroduce the
// `paragraphFormat` problem above.
//
// v4.2 (14 sep) - Mark asked for Word's own built-in styles (Standaard/
// Normaal, Kop 1-9, ...) to no longer be available/relevant, so he's
// always typing in Corbel and never picks one of those by accident. There
// is no supported Word JS API to actually delete or hide a built-in
// style, and no way from Office.js to turn on Word's own "Opmaak
// beperken tot geselecteerde stijlen" document protection (which is the
// Word feature that would do this properly) - so this is a best-effort
// mitigation, not a hard guarantee, see neutralizeBuiltInStyle() below
// and the README.
//
// v4.3 (14 sep) - Mark sent screenshots of Word Online's actual style
// gallery. Two corrections/additions from that: (1) the built-in Normal
// style is called "Normaal" there, not "Standaard" as assumed in v4.2 -
// both names are now tried; (2) added every other style shown in the
// gallery (Geen afstand, Ondertitel, Nadruk, Sterk, Citaat, Subtiele
// verwijzing, Intensieve verwijzing, Titel van boek, Lijstalinea). Mark
// only asked for the font on these to be Corbel, not a different size/
// color/emphasis, so only `font.name` is overridden for this group - see
// MISC_BUILTIN_STYLE_NAMES below.
//
// v4.4 (14 sep) - Koptekst 1-4 genummerd switched from a computed text
// prefix to Word's own native numbering (Word.List), after Mark asked to
// use Word's built-in numbering and, once told what that actually means
// (see below), explicitly chose this trade-off. Numbering is now "live":
// Word renumbers itself when headings are added, removed, reordered or
// copy/pasted - no more "Vernummeren" needed for that.
// The catch, confirmed while implementing this rather than assumed: the
// Word JavaScript API has no supported way to make one list level's
// number include its parent levels' numbers (the "1.1.1" legal/outline
// style) - `Word.List.setLevelNumbering()` only controls a level's own,
// independent counter format (arabic/roman/letter). So each of Koptekst
// 1-4 genummerd now gets its own separate, independent counter (its own
// "1.", "2.", "3.", ...) rather than the originally-specified "1." / "1.1"
// / "1.1.1" / "1.1.1.1" notation - Mark accepted this in exchange for
// live numbering. "Vernummeren" is repurposed as a one-time cleanup: it
// strips any leftover text-prefix numbering a document may still have
// from before this change (see GGM.vernummeren below); it's not needed
// for numbering created after this version, Word keeps that current on
// its own.
// ---------------------------------------------------------------------------

const GGM = {};

const FONT_NAME = "Corbel";

// -- Style definitions -------------------------------------------------------
// One font spec per style. Nothing here touches paragraph spacing - just
// what was asked for: name, size, bold, italic, color.

// Gemeente-huisstijlkleur - op Titel en alle Kopteksten (genummerd en
// ongenummerd). Standaardtekst blijft zwart.
const BRAND_COLOR = "#441D42";
const BLACK = "#000000";

const TITEL_STYLE = "Titel";
const TITEL_SPEC = { name: FONT_NAME, size: 18, bold: true, italic: false, color: BRAND_COLOR };

const STANDAARD_STYLE = "Standaardtekst";
const STANDAARD_SPEC = { name: FONT_NAME, size: 11, bold: false, italic: false, color: BLACK };

const KOPTEKST_NAMES = {
  1: "Koptekst 1",
  2: "Koptekst 2",
  3: "Koptekst 3",
  4: "Koptekst 4",
  5: "Koptekst 5",
  6: "Koptekst 6",
};

const KOPTEKST_SPECS = {
  1: { name: FONT_NAME, size: 16, bold: true, italic: false, color: BRAND_COLOR },
  2: { name: FONT_NAME, size: 14, bold: true, italic: false, color: BRAND_COLOR },
  3: { name: FONT_NAME, size: 12, bold: true, italic: false, color: BRAND_COLOR },
  4: { name: FONT_NAME, size: 11, bold: true, italic: false, color: BRAND_COLOR },
  5: { name: FONT_NAME, size: 11, bold: false, italic: true, color: BRAND_COLOR },
  6: { name: FONT_NAME, size: 11, bold: false, italic: false, color: BRAND_COLOR },
};

const KOPTEKST_GENUMMERD_NAMES = {
  1: "Koptekst 1 genummerd",
  2: "Koptekst 2 genummerd",
  3: "Koptekst 3 genummerd",
  4: "Koptekst 4 genummerd",
};

// Matches the plain-text numbering prefix this add-in inserted before
// v4.4 (levels 1-4: "1.", "1.1", "1.1.1", "1.1.1.1"), followed by a tab -
// used only to clean up documents from that older approach; numbering
// created by v4.4+ is native Word numbering, not text, so there is
// nothing for this pattern to match there.
const NUMBERED_PREFIX_RE = /^\d+(\.\d+){0,3}\.?\t/;

// -- Style creation / (re)application ----------------------------------------

/** Create a custom style if it doesn't exist yet, and always (re)apply its
 *  font - every time, not just on first creation, so a style that's
 *  already present but wrongly formatted still gets corrected. Wrapped in
 *  try/catch so a problem updating the shared style definition can never
 *  block the direct, guaranteed-visible formatting applied afterwards
 *  (see applyDirectFont). */
async function ensureStyle(context, name, font) {
  try {
    const styles = context.document.getStyles();
    let style = styles.getByNameOrNullObject(name);
    style.load("isNullObject");
    await context.sync();
    if (style.isNullObject) {
      style = context.document.addStyle(name, Word.StyleType.paragraph);
    }
    style.font.set(font);
    await context.sync();
  } catch (e) {
    // eslint-disable-next-line no-console
    console.error("Kon de gedeelde stijl '" + name + "' niet bijwerken:", e);
  }
}

GGM.ensureAllStyles = async function (context) {
  await ensureStyle(context, TITEL_STYLE, TITEL_SPEC);
  await ensureStyle(context, STANDAARD_STYLE, STANDAARD_SPEC);
  for (const level of [1, 2, 3, 4, 5, 6]) {
    await ensureStyle(context, KOPTEKST_NAMES[level], KOPTEKST_SPECS[level]);
  }
  for (const level of [1, 2, 3, 4]) {
    await ensureStyle(context, KOPTEKST_GENUMMERD_NAMES[level], KOPTEKST_SPECS[level]);
  }
};

// -- Neutraliseren van Word's eigen ingebouwde stijlen ------------------------
// Word's ingebouwde stijlen (Standaard, Kop 1-9, ...) kunnen niet via de
// Word JS API worden verwijderd of via documentbeveiliging worden
// geblokkeerd - dat kan alleen "echt" via Word's eigen "Opmaak beperken tot
// geselecteerde stijlen", wat niet in de add-in-API zit. In plaats daarvan
// wordt hier hetzelfde lettertype/formaat/kleur als de bijbehorende
// Koptekst/Standaardtekst-stijl over de ingebouwde stijl heen gezet, zodat
// een per ongeluk gekozen ingebouwde stijl toch als Corbel wordt getoond.
// Deze namen gaan uit van een Nederlandstalige Word-interface (net als de
// rest van dit document) - staat Word in een andere taal, dan wordt de
// stijl simpelweg niet gevonden (isNullObject) en gebeurt er niets, nooit
// een harde fout.

// "Normal" heeft in het Nederlands twee namen gehad afhankelijk van de
// Word-versie ("Normaal" in recente Word/Word Online - bevestigd via
// Mark's screenshot van de stijlengalerij - "Standaard" in oudere
// versies). Beide proberen is onschadelijk: welke niet bestaat, wordt
// simpelweg overgeslagen (isNullObject).
const BUILTIN_NORMAL_NAMES = ["Normaal", "Standaard"];
const BUILTIN_HEADING_NAMES = {
  1: "Kop 1", 2: "Kop 2", 3: "Kop 3", 4: "Kop 4", 5: "Kop 5",
  6: "Kop 6", 7: "Kop 7", 8: "Kop 8", 9: "Kop 9",
};
// Kop 7-9 hebben geen eigen Koptekst-knop/spec; die krijgen voor de
// zekerheid dezelfde opmaak als Koptekst 6 (Corbel 11, huisstijlkleur,
// geen extra nadruk) - beter dan een ander lettertype laten staan.
const HEADING_FALLBACK_SPEC = KOPTEKST_SPECS[6];

// De rest van de stijlen die Word Online standaard in zijn stijlengalerij
// toont (bevestigd via Mark's screenshots): een mix van alinea- en
// tekenstijlen. Mark vroeg hier specifiek alleen om het lettertype Corbel
// - niet om een andere maat/kleur/vet/cursief - dus deze krijgen alleen
// hun lettertypenaam overschreven; hun eigen overige opmaak (grootte,
// cursief, kleur, ...) blijft ongemoeid. `Word.Font.set()` past alleen de
// eigenschappen aan die je meegeeft, dus { name: "Corbel" } laat de rest
// van elke stijl met rust.
const MISC_BUILTIN_STYLE_NAMES = [
  "Geen afstand", // No Spacing
  "Ondertitel", // Subtitle
  "Nadruk", // Emphasis (tekenstijl)
  "Sterk", // Strong (tekenstijl)
  "Citaat", // Quote
  "Subtiele verwijzing", // Subtle Reference (tekenstijl)
  "Intensieve verwijzing", // Intense Reference (tekenstijl)
  "Titel van boek", // Book Title (tekenstijl)
  "Lijstalinea", // List Paragraph
];
const FONT_ONLY_PATCH = { name: FONT_NAME };

/** (Re)apply font to one of Word's own built-in styles, but only if it
 *  already exists in the document - never create it under that name (that
 *  would risk creating a wrong *custom* style if the lookup ever missed
 *  for some reason, e.g. a non-Dutch display language). Also makes a
 *  best-effort, separately try/caught attempt to drop it from Word's
 *  Quick Styles gallery in the ribbon (`quickStyle`) - this property isn't
 *  guaranteed supported in every Word Online build (this project has hit
 *  an unsupported Word.Style property before, see the paragraphFormat
 *  note above), so it must never be allowed to block the font fix, and a
 *  failure here is not a real problem: the style still renders correctly,
 *  it just may still be visible in the gallery. Even when it works, this
 *  only affects the Quick Styles gallery in the ribbon, not the full
 *  "Stijlen"-deelvenster (Ctrl+Alt+Shift+S), where every built-in style
 *  remains selectable by name regardless. */
async function neutralizeBuiltInStyle(context, name, font) {
  try {
    const styles = context.document.getStyles();
    const style = styles.getByNameOrNullObject(name);
    style.load("isNullObject");
    await context.sync();
    if (style.isNullObject) return;
    style.font.set(font);
    await context.sync();
  } catch (e) {
    // eslint-disable-next-line no-console
    console.error("Kon ingebouwde stijl '" + name + "' niet overschrijven:", e);
  }

  try {
    const styles = context.document.getStyles();
    const style = styles.getByNameOrNullObject(name);
    style.load("isNullObject");
    await context.sync();
    if (style.isNullObject) return;
    style.quickStyle = false;
    await context.sync();
  } catch (e) {
    // eslint-disable-next-line no-console
    console.error("Kon ingebouwde stijl '" + name + "' niet uit de snelstijlengalerie halen (niet kritiek):", e);
  }
}

GGM.neutralizeBuiltInStyles = async function (context) {
  for (const name of BUILTIN_NORMAL_NAMES) {
    await neutralizeBuiltInStyle(context, name, STANDAARD_SPEC);
  }
  for (const level of [1, 2, 3, 4, 5, 6]) {
    await neutralizeBuiltInStyle(context, BUILTIN_HEADING_NAMES[level], KOPTEKST_SPECS[level]);
  }
  for (const level of [7, 8, 9]) {
    await neutralizeBuiltInStyle(context, BUILTIN_HEADING_NAMES[level], HEADING_FALLBACK_SPEC);
  }
  for (const name of MISC_BUILTIN_STYLE_NAMES) {
    await neutralizeBuiltInStyle(context, name, FONT_ONLY_PATCH);
  }
};

GGM.stijlenInstalleren = async function () {
  await Word.run(async (context) => {
    await GGM.ensureAllStyles(context);
    await GGM.neutralizeBuiltInStyles(context);
  });
};

/** Strip a leftover pre-v4.4 text numbering prefix from the current
 *  selection's text, if any (used when switching a paragraph to a style
 *  that isn't numbered, or re-applying a numbered one). */
async function stripLeftoverPrefixText(context) {
  const selection = context.document.getSelection();
  selection.load("text");
  await context.sync();
  const bareText = selection.text.replace(NUMBERED_PREFIX_RE, "");
  if (bareText !== selection.text) {
    selection.insertText(bareText, Word.InsertLocation.replace);
    await context.sync();
  }
}

/** Detach the current selection's first paragraph from any native Word
 *  list it's part of (see GGM.applyKoptekstGenummerd below), so switching
 *  a numbered heading to a non-numbered style doesn't leave a dangling
 *  number in front of it. Safe to call on a paragraph that isn't in a
 *  list. Try/caught defensively, same reasoning as neutralizeBuiltInStyle
 *  above: this must never block the style/font change around it. */
async function detachSelectionFromList(context) {
  try {
    context.document.getSelection().paragraphs.getFirst().detachFromList();
    await context.sync();
  } catch (e) {
    // eslint-disable-next-line no-console
    console.error("Kon nummering niet loskoppelen van de selectie:", e);
  }
}

/** Apply the font directly to the current selection's characters, on top
 *  of whatever style was just applied. Restyling the shared style
 *  (ensureStyle above) is what keeps the look consistent everywhere that
 *  style is used, but applying the same font directly to the selected
 *  text as well means the button always visibly does the right thing,
 *  even if the style lookup were to silently miss for some reason. */
async function applyDirectFont(context, font) {
  context.document.getSelection().font.set(font);
  await context.sync();
}

// -- Titel / Standaardtekst ---------------------------------------------------

GGM.titel = async function () {
  await Word.run(async (context) => {
    await ensureStyle(context, TITEL_STYLE, TITEL_SPEC);
    await stripLeftoverPrefixText(context);
    await detachSelectionFromList(context);
    context.document.getSelection().style = TITEL_STYLE;
    await context.sync();
    await applyDirectFont(context, TITEL_SPEC);
  });
};

GGM.standaardtekst = async function () {
  await Word.run(async (context) => {
    await ensureStyle(context, STANDAARD_STYLE, STANDAARD_SPEC);
    await stripLeftoverPrefixText(context);
    await detachSelectionFromList(context);
    context.document.getSelection().style = STANDAARD_STYLE;
    await context.sync();
    await applyDirectFont(context, STANDAARD_SPEC);
  });
};

// -- Kopteksten zonder nummering ----------------------------------------------

GGM.applyKoptekst = async function (level) {
  const name = KOPTEKST_NAMES[level];
  const font = KOPTEKST_SPECS[level];
  await Word.run(async (context) => {
    await ensureStyle(context, name, font);
    await stripLeftoverPrefixText(context);
    await detachSelectionFromList(context);
    context.document.getSelection().style = name;
    await context.sync();
    await applyDirectFont(context, font);
  });
};

GGM.koptekst1 = async function () { await GGM.applyKoptekst(1); };
GGM.koptekst2 = async function () { await GGM.applyKoptekst(2); };
GGM.koptekst3 = async function () { await GGM.applyKoptekst(3); };
GGM.koptekst4 = async function () { await GGM.applyKoptekst(4); };
GGM.koptekst5 = async function () { await GGM.applyKoptekst(5); };
GGM.koptekst6 = async function () { await GGM.applyKoptekst(6); };

// -- Kopteksten met nummering --------------------------------------------------
// Elk niveau (Koptekst 1-4 genummerd) heeft zijn eigen, onafhankelijke
// Word.List: alle paragrafen met dat niveau's stijl delen die ene lijst,
// zodat de teller binnen dat niveau doorloopt door het hele document.
// Zie de v4.4-toelichting bovenaan: dit is Word's eigen, "levende"
// nummering, geen door de add-in berekende tekst meer.

/** Find the id of an existing Word.List already used elsewhere in the
 *  document for this heading style, if any - so a newly (re)styled
 *  paragraph joins the same running count instead of starting its own.
 *  Returns null if this style doesn't have a list yet anywhere. */
async function findExistingListId(context, styleName) {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("style,listOrNullObject/id,listOrNullObject/isNullObject");
  await context.sync();

  for (let i = 0; i < paragraphs.items.length; i++) {
    const p = paragraphs.items[i];
    if (p.style === styleName && !p.listOrNullObject.isNullObject) {
      return p.listOrNullObject.id;
    }
  }
  return null;
}

/** Attach the current selection's first paragraph to this heading level's
 *  native numbered list - reusing the one other paragraphs of the same
 *  style already use, or starting a new one (arabic numbering, e.g.
 *  "1.", "2.", "3.", ...) if this is the first paragraph anywhere in the
 *  document with this style. Try/caught, same reasoning as
 *  neutralizeBuiltInStyle above: this is new, not-yet-live-tested API
 *  usage (Word.List), so a failure here must never block applyDirectFont
 *  afterwards - the heading should still visibly get the right style and
 *  font even if the native numbering itself doesn't take. */
async function attachSelectionToNumberedList(context, styleName) {
  try {
    const existingListId = await findExistingListId(context, styleName);
    const target = context.document.getSelection().paragraphs.getFirst();
    if (existingListId !== null) {
      target.attachToList(existingListId, 0);
    } else {
      const list = target.startNewList();
      list.setLevelNumbering(0, Word.ListNumbering.arabic);
    }
    await context.sync();
  } catch (e) {
    // eslint-disable-next-line no-console
    console.error("Kon '" + styleName + "' niet aan de automatische nummering koppelen:", e);
  }
}

GGM.applyKoptekstGenummerd = async function (level) {
  const name = KOPTEKST_GENUMMERD_NAMES[level];
  const font = KOPTEKST_SPECS[level];
  await Word.run(async (context) => {
    await ensureStyle(context, name, font);
    await stripLeftoverPrefixText(context);
    context.document.getSelection().style = name;
    await context.sync();
    await attachSelectionToNumberedList(context, name);
    await applyDirectFont(context, font);
  });
};

GGM.koptekst1Genummerd = async function () { await GGM.applyKoptekstGenummerd(1); };
GGM.koptekst2Genummerd = async function () { await GGM.applyKoptekstGenummerd(2); };
GGM.koptekst3Genummerd = async function () { await GGM.applyKoptekstGenummerd(3); };
GGM.koptekst4Genummerd = async function () { await GGM.applyKoptekstGenummerd(4); };

/** One-time cleanup for documents that still carry the old, pre-v4.4 text
 *  numbering (a literal "1." / "1.1" / ... prefix, see NUMBERED_PREFIX_RE)
 *  - strips it from every paragraph in the document. Not needed for
 *  numbering created by v4.4+: that's native Word numbering and Word
 *  keeps it current on its own, without any button. */
GGM.vernummeren = async function () {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();

    paragraphs.items.forEach((p) => {
      const bareText = p.text.replace(NUMBERED_PREFIX_RE, "");
      if (bareText !== p.text) {
        p.insertText(bareText, Word.InsertLocation.replace);
      }
    });

    await context.sync();
  });
};

// Expose to global scope for plain <script> consumption (no bundler/module
// system in this project).
if (typeof window !== "undefined") {
  window.GGM = GGM;
}
