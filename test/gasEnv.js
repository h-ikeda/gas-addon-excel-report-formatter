/**
 * Test harness for the Google Apps Script source.
 *
 * Code.gs relies on Apps Script global services (PropertiesService,
 * UrlFetchApp, ScriptApp, ...) that do not exist under Node. We load the
 * untouched source into a fresh `vm` context with those globals replaced by
 * mocks, then return the context so individual functions can be invoked and
 * asserted on.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const CODE_PATH = path.join(__dirname, '..', 'gas', 'Code.gs');
const source = fs.readFileSync(CODE_PATH, 'utf8');

/**
 * Evaluate Code.gs against the provided sandbox of mocked globals and return
 * the resulting context (the declared functions become properties on it).
 */
function loadGas(sandbox = {}) {
  const context = vm.createContext(sandbox);
  vm.runInContext(source, context);
  return context;
}

module.exports = { loadGas, CODE_PATH };
