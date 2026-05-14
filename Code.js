// RealTimeTranslate for Google Workspace
// Google Apps Script to translate selected text in real-time using LibreTranslate.

// LibreTranslate API endpoint (public instance)
const LIBRE_TRANSLATE_API = "https://libretranslate.de/translate";

/**
 * Translates text using LibreTranslate API.
 * @param {string} text - Text to translate.
 * @param {string} sourceLang - Source language code (e.g., "en").
 * @param {string} targetLang - Target language code (e.g., "es").
 * @return {string} Translated text.
 */
function translateText(text, sourceLang, targetLang) {
  const payload = {
    q: text,
    source: sourceLang,
    target: targetLang,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(LIBRE_TRANSLATE_API, options);
    const result = JSON.parse(response.getContentText());
    return result.translatedText || "Translation failed.";
  } catch (e) {
    console.error("Translation error:", e);
    return "Translation error.";
  }
}

/**
 * Creates a custom menu in Google Docs.
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu("RealTimeTranslate")
    .addItem("Translate Selection", "showTranslationSidebar")
    .addToUi();
}

/**
 * Shows a sidebar for language selection and translation.
 */
function showTranslationSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("RealTimeTranslate")
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Translates selected text and returns the result.
 * @param {string} sourceLang - Source language code.
 * @param {string} targetLang - Target language code.
 * @return {string} Translated text.
 */
function translateSelection(sourceLang, targetLang) {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) return "No text selected.";

  const text = selection.getSelectedElements()
    .map(element => element.getElement().asText().getText())
    .join(" ");

  return translateText(text, sourceLang, targetLang);
}