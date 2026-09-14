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
// The numbered look for Koptekst 1-4 (1. / 1.1 / 1.1.1 / 1.1.1.1) still
// has no equivalent in the Word JS API (no supported way to link a style
// to a multilevel-list definition), so it's computed by this add-in and
// written as literal prefix text at the start of the paragraph,
// recalculated for the whole document whenever a numbered style is
// applied. That means moving, deleting or copy/pasting numbered headings
// doesn't renumber the rest on its own - use "Vernummeren" after that.
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

// Same font per level as the un-numbered Koptekst - only the numbering is
// different, per Mark's spec: level 1 gets a trailing period ("1."),
// levels 2-4 don't ("1.1", "1.1.1", "1.1.1.1").
function numberedPrefix(counters, level) {
  if (level === 1) {
    return `${counters[0]}.`;
  }
  return counters.slice(0, level).join(".");
}

// Matches a numbering prefix this add-in previously inserted (levels 1-4:
// "1.", "1.1", "1.1.1", "1.1.1.1"), followed by a tab, so it can be
// stripped and recomputed cleanly.
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

GGM.stijlenInstalleren = async function () {
  await Word.run(async (context) => {
    await GGM.ensureAllStyles(context);
  });
};

/** Strip a leftover numbering prefix from the current selection's text, if
 *  any (used when switching a paragraph to a style that isn't numbered). */
async function stripPrefixFromSelection(context) {
  const selection = context.document.getSelection();
  selection.load("text");
  await context.sync();
  const bareText = selection.text.replace(NUMBERED_PREFIX_RE, "");
  if (bareText !== selection.text) {
    selection.insertText(bareText, Word.InsertLocation.replace);
    await context.sync();
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
    await stripPrefixFromSelection(context);
    context.document.getSelection().style = TITEL_STYLE;
    await context.sync();
    await applyDirectFont(context, TITEL_SPEC);
  });
};

GGM.standaardtekst = async function () {
  await Word.run(async (context) => {
    await ensureStyle(context, STANDAARD_STYLE, STANDAARD_SPEC);
    await stripPrefixFromSelection(context);
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
    await stripPrefixFromSelection(context);
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

GGM.applyKoptekstGenummerd = async function (level) {
  const name = KOPTEKST_GENUMMERD_NAMES[level];
  const font = KOPTEKST_SPECS[level];
  await Word.run(async (context) => {
    await ensureStyle(context, name, font);
    context.document.getSelection().style = name;
    await context.sync();
    await applyDirectFont(context, font);
  });
  await GGM.renumber();
};

GGM.koptekst1Genummerd = async function () { await GGM.applyKoptekstGenummerd(1); };
GGM.koptekst2Genummerd = async function () { await GGM.applyKoptekstGenummerd(2); };
GGM.koptekst3Genummerd = async function () { await GGM.applyKoptekstGenummerd(3); };
GGM.koptekst4Genummerd = async function () { await GGM.applyKoptekstGenummerd(4); };

/** Recompute and rewrite the 1. / 1.1 / 1.1.1 / 1.1.1.1 prefixes for every
 *  paragraph in the document, in document order. Paragraphs that no longer
 *  carry a numbered Koptekst style have any leftover prefix removed. */
GGM.renumber = async function () {
  await Word.run(async (context) => {
    const levelByStyleName = {};
    for (const level of [1, 2, 3, 4]) {
      levelByStyleName[KOPTEKST_GENUMMERD_NAMES[level]] = level;
    }

    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("style,text");
    await context.sync();

    const counters = [0, 0, 0, 0];

    paragraphs.items.forEach((p) => {
      const bareText = p.text.replace(NUMBERED_PREFIX_RE, "");
      const level = levelByStyleName[p.style];

      if (!level) {
        if (bareText !== p.text) {
          p.insertText(bareText, Word.InsertLocation.replace);
        }
        return;
      }

      counters[level - 1] += 1;
      for (let i = level; i < counters.length; i++) counters[i] = 0;
      const prefix = numberedPrefix(counters, level);

      const newText = `${prefix}\t${bareText}`;
      if (newText !== p.text) {
        p.insertText(newText, Word.InsertLocation.replace);
      }
    });

    await context.sync();
  });
};

GGM.vernummeren = async function () {
  await GGM.renumber();
};

// Expose to global scope for plain <script> consumption (no bundler/module
// system in this project).
if (typeof window !== "undefined") {
  window.GGM = GGM;
}
