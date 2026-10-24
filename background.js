'use strict';

// Extensions DocLens can render. Legacy binary .doc/.ppt are intentionally absent.
const WORD_EXTS = ['docx', 'docm', 'dotx', 'dotm'];
const EXCEL_EXTS = ['xlsx', 'xlsm', 'xlsb', 'xltx', 'xltm', 'xls', 'ods'];
const POWERPOINT_EXTS = ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm'];
const OFFICE_EXTS = [...WORD_EXTS, ...EXCEL_EXTS, ...POWERPOINT_EXTS];

const OFFICE_MIME_RE = new RegExp(
  '^application/(' +
    [
      'vnd\\.openxmlformats-officedocument\\.(wordprocessingml|spreadsheetml|presentationml)\\.',
      'vnd\\.ms-word\\.(document|template)\\.macroenabled',
      'vnd\\.ms-excel',
      'vnd\\.ms-powerpoint\\.(presentation|slideshow|template)\\.macroenabled',
      'vnd\\.oasis\\.opendocument\\.spreadsheet',
    ].join('|') +
    ')',
  'i'
);
const OFFICE_EXT_RE = new RegExp(`\\.(${OFFICE_EXTS.join('|')})$`, 'i');

const REDIRECT_RULE_ID = 1;
const CONTEXT_MENU_ID = 'doclens-open-link';
const DEFAULTS = { interceptLinks: true, interceptDownloads: true };

const VIEWER_URL = chrome.runtime.getURL('viewer.html');

let settings = { ...DEFAULTS };
const settingsReady = chrome.storage.sync.get(DEFAULTS).then((s) => {
  settings = s;
});

const viewerUrlFor = (url, name) =>
  `${VIEWER_URL}#${name ? `name=${encodeURIComponent(name)}&` : ''}url=${encodeURIComponent(url)}`;

// Navigations to http(s) URLs whose path ends in an Office extension are
// redirected to the viewer. The raw URL goes into the hash, which the viewer
// reads verbatim (regexSubstitution cannot percent-encode).
async function syncRedirectRule() {
  await settingsReady;
  const addRules = settings.interceptLinks
    ? [
        {
          id: REDIRECT_RULE_ID,
          priority: 1,
          action: {
            type: 'redirect',
            redirect: { regexSubstitution: `${VIEWER_URL}#url=\\0` },
          },
          condition: {
            regexFilter: `^https?://[^?#]*\\.(?:${OFFICE_EXTS.join('|')})(?:[?#].*)?$`,
            isUrlFilterCaseSensitive: false,
            resourceTypes: ['main_frame'],
          },
        },
      ]
    : [];
  await chrome.declarativeNetRequest.updateDynamicRules({
    removeRuleIds: [REDIRECT_RULE_ID],
    addRules,
  });
}

function createContextMenu() {
  const targetUrlPatterns = OFFICE_EXTS.flatMap((ext) => [
    `*://*/*.${ext}`,
    `*://*/*.${ext}?*`,
    `file:///*.${ext}`,
  ]);
  chrome.contextMenus.removeAll(() => {
    chrome.contextMenus.create({
      id: CONTEXT_MENU_ID,
      title: 'Open in DocLens',
      contexts: ['link'],
      targetUrlPatterns,
    });
  });
}

// Without "Allow access to file URLs" Chrome hides file:// navigations and
// downloads from the extension entirely, so flag it on the toolbar icon.
// (Toggling that setting reloads the extension, which re-runs this.)
async function updateFileAccessBadge() {
  const allowed = await chrome.extension.isAllowedFileSchemeAccess();
  await chrome.action.setBadgeText({ text: allowed ? '' : '!' });
  await chrome.action.setBadgeBackgroundColor({ color: '#d93025' });
  await chrome.action.setTitle({
    title: allowed ? 'DocLens' : 'DocLens – turn on "Allow access to file URLs" to open files from your computer',
  });
}
updateFileAccessBadge();

chrome.runtime.onInstalled.addListener(() => {
  createContextMenu();
  syncRedirectRule();
});
chrome.runtime.onStartup.addListener(syncRedirectRule);

chrome.storage.onChanged.addListener((changes, area) => {
  if (area !== 'sync') return;
  for (const [key, { newValue }] of Object.entries(changes)) {
    if (key in DEFAULTS) settings[key] = newValue ?? DEFAULTS[key];
  }
  if ('interceptLinks' in changes) syncRedirectRule();
});

// Local files (file:///…/report.docx) bypass declarativeNetRequest, and Chrome
// may try to show them itself ("Couldn't load plug-in"). Send the tab to the
// viewer before that happens.
chrome.webNavigation.onBeforeNavigate.addListener(
  async ({ tabId, frameId, url }) => {
    if (frameId !== 0 || tabId < 0) return;
    await settingsReady;
    if (!settings.interceptLinks) return;
    chrome.tabs.update(tabId, { url: viewerUrlFor(url) });
  },
  { url: [{ schemes: ['file'], urlMatches: `(?i)\\.(${OFFICE_EXTS.join('|')})$` }] }
);

chrome.contextMenus.onClicked.addListener((info, tab) => {
  if (info.menuItemId !== CONTEXT_MENU_ID || !info.linkUrl) return;
  chrome.tabs.create({
    url: viewerUrlFor(info.linkUrl),
    index: tab ? tab.index + 1 : undefined,
    openerTabId: tab?.id,
  });
});

const urlPathname = (url) => {
  try {
    return decodeURIComponent(new URL(url).pathname);
  } catch {
    return '';
  }
};

function isOfficeDownload(item) {
  if (item.mime && OFFICE_MIME_RE.test(item.mime)) return true;
  return [item.filename, urlPathname(item.finalUrl || item.url)].some((n) => n && OFFICE_EXT_RE.test(n));
}

// Downloads that never pass through a navigation (Content-Disposition:
// attachment, <a download>, file:// URLs) are cancelled and re-opened in the viewer.
chrome.downloads.onCreated.addListener(async (item) => {
  if (item.byExtensionId === chrome.runtime.id) return; // "Download" clicked inside DocLens
  if (item.state !== 'in_progress') return;
  const url = item.finalUrl || item.url;
  if (!/^(https?|file):/i.test(url)) return; // blob:/data: URLs can't be re-fetched
  await settingsReady;
  if (!settings.interceptDownloads || !isOfficeDownload(item)) return;

  try {
    await chrome.downloads.cancel(item.id);
    await chrome.downloads.erase({ id: item.id });
  } catch {
    // Download may already have finished; still show it in the viewer.
  }
  const name = item.filename ? item.filename.split(/[\\/]/).pop() : '';
  chrome.tabs.create({ url: viewerUrlFor(url, name) });
});
