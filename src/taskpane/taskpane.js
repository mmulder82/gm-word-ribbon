/* global Office, GGM */

Office.onReady(() => {
  const statusEl = document.getElementById("statusMessage");

  function setStatus(message, isError) {
    statusEl.textContent = message || "";
    statusEl.classList.toggle("error", Boolean(isError));
  }

  async function run(action, busyLabel) {
    setStatus(busyLabel ? busyLabel + "..." : "Bezig...", false);
    try {
      await action();
      setStatus("Klaar.", false);
    } catch (e) {
      setStatus(e && e.message ? e.message : "Er ging iets mis.", true);
    }
  }

  document.getElementById("btnStijlenInstalleren").addEventListener("click", () => {
    run(async () => {
      await GGM.stijlenInstalleren();
      setStatus("Gemeentelijke stijlen (opnieuw) toegepast in dit document.", false);
    });
  });

  const actionMap = {
    titel: GGM.titel,
    standaardtekst: GGM.standaardtekst,
    koptekst1: GGM.koptekst1,
    koptekst2: GGM.koptekst2,
    koptekst3: GGM.koptekst3,
    koptekst4: GGM.koptekst4,
    koptekst5: GGM.koptekst5,
    koptekst6: GGM.koptekst6,
    koptekst1Genummerd: GGM.koptekst1Genummerd,
    koptekst2Genummerd: GGM.koptekst2Genummerd,
    koptekst3Genummerd: GGM.koptekst3Genummerd,
    koptekst4Genummerd: GGM.koptekst4Genummerd,
    vernummeren: GGM.vernummeren,
  };

  document.querySelectorAll("[data-action]").forEach((btn) => {
    const key = btn.getAttribute("data-action");
    const fn = actionMap[key];
    if (!fn) return;
    btn.addEventListener("click", () => run(fn));
  });
});
