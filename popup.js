'use strict';

const DEFAULTS = { interceptLinks: true, interceptDownloads: true };

document.getElementById('openViewer').onclick = () => {
  chrome.tabs.create({ url: chrome.runtime.getURL('viewer.html') });
  window.close();
};

document.getElementById('openSettings').onclick = () => {
  chrome.tabs.create({ url: `chrome://extensions/?id=${chrome.runtime.id}` });
};

chrome.storage.sync.get(DEFAULTS).then((settings) => {
  for (const key of Object.keys(DEFAULTS)) {
    const input = document.getElementById(key);
    input.checked = settings[key];
    input.onchange = () => chrome.storage.sync.set({ [key]: input.checked });
  }
});

chrome.extension.isAllowedFileSchemeAccess().then((allowed) => {
  document.getElementById('fileAccess').hidden = allowed;
});
