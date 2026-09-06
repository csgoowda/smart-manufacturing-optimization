"use strict";
/* Multilingual UI layer.

   Dictionaries live in /static/locales/<code>.json — one flat JSON object per
   language. To add a language: create the JSON file (same keys as en.json)
   and append an entry (code, native name, flag SVG) to LANGUAGES below. */

const LANGUAGES = [
  {
    code: "en",
    name: "English",
  },
];

const I18N = {};

async function loadLocales() {
  await Promise.all(LANGUAGES.map(async lang => {
    const r = await fetch(`/static/locales/${lang.code}.json`);
    if (!r.ok) throw new Error(`could not load locale ${lang.code}`);
    I18N[lang.code] = await r.json();
  }));
}

let LANG = "en";

function t(key, vars) {
  const dict = I18N[LANG] || {};
  let s = dict[key] || (I18N.en || {})[key] || key;
  if (vars) for (const [k, v] of Object.entries(vars)) s = s.split("{" + k + "}").join(v);
  return s;
}

function applyI18n() {
  document.documentElement.lang = LANG;
  document.querySelectorAll("[data-i18n]").forEach(el => {
    el.textContent = t(el.dataset.i18n);
  });
  document.querySelectorAll("[data-i18n-title]").forEach(el => {
    const v = t(el.dataset.i18nTitle);
    el.title = v;
    el.setAttribute("aria-label", v);
  });
  document.querySelectorAll("[data-i18n-placeholder]").forEach(el => {
    el.placeholder = t(el.dataset.i18nPlaceholder);
  });
}

function setLang(lang) {
  if (!LANGUAGES.some(l => l.code === lang)) return;
  LANG = lang;
  localStorage.setItem("optimal-task-planner.lang", lang);
  applyI18n();
  if (typeof renderAll === "function" && window.__appReady) renderAll();
}
