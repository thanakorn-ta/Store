/**
 * Ui.js
 * Serves the same reorder-comparison page as GitHub Pages (web/) from this
 * Apps Script Web App. The ui_*.html and Index.html files are GENERATED from
 * web/ by tools/build-appsscript.ps1 — edit web/, then rebuild; never edit
 * the generated files by hand.
 *
 * Routing lives in WebApp.js doGet(): links with ?pc=... (from the reorder
 * emails) open the confirm page, everything else opens this UI.
 */

function renderReorderUi_() {
  // The copy-paste bundle (build/apps-script/Code.gs) embeds the page as INDEX_HTML_,
  // so pasting that one file is enough; clasp deployments use the Index.html file instead.
  var page = typeof INDEX_HTML_ === 'string'
    ? HtmlService.createHtmlOutput(INDEX_HTML_)
    : HtmlService.createTemplateFromFile('Index').evaluate();
  return page
    .setTitle('Store Reorder AI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Used by Index.html as <?!= include('ui_app'); ?> */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/**
 * Called from the page's "ผู้ช่วย AI" tab. Key stays in Script Properties
 * (ANTHROPIC_API_KEY), so viewers of the Web App never see it.
 */
function askClaudeFromUi(system, userContent) {
  requireActive_();
  if (String(userContent).length > 200000) throw new Error('ข้อมูลที่ส่งให้ AI ยาวเกินไป');
  return callClaude_(system, userContent, 4000);
}

