/* global Office, GGM */

// ---------------------------------------------------------------------------
// Gemeente Gooise Meren - Word ribbon commands.
//
// Thin wrappers: each ribbon button calls into the shared implementation in
// src/shared/ggm-actions.js and then signals completion, as required for
// Office Add-in ExecuteFunction commands.
// ---------------------------------------------------------------------------

Office.onReady();

function wrap(fn) {
  return async function (event) {
    try {
      await fn();
    } catch (e) {
      // Ribbon commands run headless (no UI); errors are swallowed after
      // being logged so a mistaken click never leaves the ribbon looking
      // stuck.
      // eslint-disable-next-line no-console
      console.error(e);
    }
    event.completed();
  };
}

Office.actions.associate("actionTitel", wrap(GGM.titel));
Office.actions.associate("actionStandaardtekst", wrap(GGM.standaardtekst));
Office.actions.associate("actionKoptekst1", wrap(GGM.koptekst1));
Office.actions.associate("actionKoptekst2", wrap(GGM.koptekst2));
Office.actions.associate("actionKoptekst3", wrap(GGM.koptekst3));
Office.actions.associate("actionKoptekst4", wrap(GGM.koptekst4));
Office.actions.associate("actionKoptekst5", wrap(GGM.koptekst5));
Office.actions.associate("actionKoptekst6", wrap(GGM.koptekst6));
Office.actions.associate("actionKoptekst1Genummerd", wrap(GGM.koptekst1Genummerd));
Office.actions.associate("actionKoptekst2Genummerd", wrap(GGM.koptekst2Genummerd));
Office.actions.associate("actionKoptekst3Genummerd", wrap(GGM.koptekst3Genummerd));
Office.actions.associate("actionKoptekst4Genummerd", wrap(GGM.koptekst4Genummerd));
Office.actions.associate("actionStijlenInstalleren", wrap(GGM.stijlenInstalleren));
Office.actions.associate("actionVernummeren", wrap(GGM.vernummeren));
