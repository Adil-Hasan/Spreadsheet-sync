import { formatDataFillEmpty, formatMsToTime } from './functions';
import { logger } from './logger';
import { ChangeTypes, LoggerType, RequestMessage, RequestType, SelectedNodesInfo, SelectedNodeType, SpreadsheetData, SpreadsheetDataSheetsInfo, SpreadsheetSheetsData, SpreadsheetSingleSheetData, TooltipTheme, WindowSize } from './models';
import { isStringType } from './regex-types';
import { appWindowMinSize, appWindowMinSizeAdvancedMode, layerNamePrefix } from './variables';

// @ts-ignore
import tippy from 'tippy.js';

import 'tippy.js/dist/tippy.css';
import 'tippy.js/animations/shift-away-subtle.css';
import './ui.scss';

// sheet is public
let sheetIsPublic = false;

// layer selection count
let layerSelectionCount: number = 0;
let layerSelectionNames: string[];

// spreadsheet data
let spreadsheetData: SpreadsheetData;
let uploadedCsvTable: SpreadsheetSingleSheetData | null = null;
let uploadedCsvTitle: string | null = null;

// if on advanced mode window
let isPreviewOpen: boolean = false;

// elems (some may not exist in simplified UI)
const buttonReCheck = document.getElementById('btn-re-check-link') as HTMLButtonElement | null;
const cornerResize = document.getElementById('corner-resize') as HTMLElement | null;

const nodesSelectedInfo = document.getElementById('nodes-selected-info') as HTMLElement | null;
const nodesSelectedCount = nodesSelectedInfo?.querySelector('.nodes-count') as HTMLElement | null;
const nodesSelectedLayersToChange = nodesSelectedInfo?.querySelector('.layers-to-change') as HTMLElement | null;

const totalLayerChanges = document.querySelector('#title .total-layer-changes') as HTMLElement | null;

// const sheetsSelector = document.getElementById('sheets-selector');

// const modalLayersLoading = document.getElementById('loading-layers-modal') as HTMLElement;
// const nodesCount = document.getElementById('nodes-count') as HTMLElement;
// const totalNodes = document.getElementById('total-nodes') as HTMLElement;
// const loading = document.getElementById('loading') as HTMLElement;



// message --------------------

window.onmessage = async (event) => {
  if (event.data?.pluginMessage?.type) {
    if (event.data.pluginMessage.type === 'init') {
      generateTooltips();
    }

    if (event.data.pluginMessage.type === 'process-image') {
      let imageUrl = event.data.pluginMessage.imageUrl;

      await fetch(imageUrl)
        .then((r) => r.arrayBuffer())
        .then((imageData) => {
          parent.postMessage(
            {
              pluginMessage: {
                type: 'image-ready',
                image: new Uint8Array(imageData),
                nodeId: event.data.pluginMessage.nodeId,
                isPlaceholder: event.data.pluginMessage.isPlaceholder
              },
            },
            '*'
          );
        });
    }

    if (event.data.pluginMessage.type === 'sync-dev') {
      setTimeout(() => {
        document.getElementById('sync').click();
      }, 100);
    }

    if (event.data.pluginMessage.type === 'set-input-url') {
      const inputUrl = document.getElementById('api-url') as HTMLInputElement;
      inputUrl.value = event.data.pluginMessage.value;

      updateInputUrl();
      checkUrlPublic();
      updatePreviewButton(event.data.pluginMessage.value);
    }

    if (event.data.pluginMessage.type === 'layers') {
      const sectionMessage = document.getElementById('section-message');
      const sectionSync = document.getElementById('section-sync');
      const titleElem = document.querySelector('#title') as HTMLElement;
      const countElem = titleElem?.querySelector('.count') as HTMLElement;
      const messageElem = titleElem?.querySelector('.message') as HTMLElement;
      const subtitleElem = document.querySelector('#subtitle') as HTMLElement;
      const buttonSync = document.querySelector('#sync') as HTMLButtonElement;

      layerSelectionNames = event.data.pluginMessage.layerNames;
      layerSelectionCount = event.data.pluginMessage.layerCount;

      updateTableSelection();

      updateTooltipSelections();

      removeClsStartingWith([sectionMessage, countElem, sectionSync].filter(Boolean), 'bg-');

      if (titleElem) { removeCls(titleElem, ['text-success', 'text-warning', 'text-error']); }
      addCls([subtitleElem, countElem].filter(Boolean), 'hidden');

      if (event.data.pluginMessage.color === 'success' && buttonSync) {
        buttonSync.disabled = false;
        addCls([sectionMessage, sectionSync].filter(Boolean), 'bg-success');
        if (countElem) addCls(countElem, 'bg-success-full');
        if (titleElem) addCls(titleElem, 'text-success');
        removeCls([subtitleElem, countElem].filter(Boolean), 'hidden');
      }

      if (event.data.pluginMessage.color === 'warning' && buttonSync) {
        buttonSync.disabled = true;
        addCls([sectionMessage, sectionSync].filter(Boolean), 'bg-warning');
        if (countElem) addCls(countElem, 'bg-warning-full');
        if (titleElem) addCls(titleElem, 'text-warning');
        removeCls([subtitleElem, countElem].filter(Boolean), 'hidden');
      }

      if (event.data.pluginMessage.color === 'error' && buttonSync) {
        buttonSync.disabled = true;
        addCls([sectionMessage, sectionSync].filter(Boolean), 'bg-error');
        if (countElem) addCls(countElem, 'bg-error-full');
        if (titleElem) addCls(titleElem, 'text-error');
        removeCls([subtitleElem, countElem].filter(Boolean), 'hidden');
      }
      if (countElem) countElem.querySelector('span').innerText = event.data.pluginMessage.layerCount;
      if (messageElem) messageElem.innerText = event.data.pluginMessage.message;
      if (event.data.pluginMessage.description && subtitleElem) {
        subtitleElem.innerText = event.data.pluginMessage.description;
      }
    }

    if (event.data.pluginMessage.type === 'update-table-selected') {
      layerSelectionNames = event.data.pluginMessage.layerNames;
      layerSelectionCount = event.data.pluginMessage.layerCount;
      updateTableSelection();
    }

    if (event.data.pluginMessage.type === 'populate-json-preview') {
      event.data.pluginMessage.data.map((sheet, index) => {
        document.getElementById(`code-preview-${index + 1}`).innerHTML = sheet.json;
        (document.getElementById(`code-to-copy-${index + 1}`) as HTMLTextAreaElement).value = JSON.stringify(sheet.data, undefined, 2);
        imageTooltipListeners(`code-preview-${index + 1}`);

        // copy json to clipboard
        const btnCopy = document.getElementById(`btn-copy-to-clipboard-${index + 1}`) as HTMLElement;
        btnCopy.addEventListener('click', () => {
          const textarea = document.getElementById(`code-to-copy-${index + 1}`) as HTMLTextAreaElement;

          textarea.select();
          document.execCommand('copy');

          const originalTooltipName = btnCopy.getAttribute('tooltip');
          tooltipUpdateTheme(btnCopy, TooltipTheme.SUCCESS);
          tooltipUpdateText(btnCopy, 'Copied!');
          setTimeout(() => {
            tooltipUpdateTheme(btnCopy);
            tooltipUpdateText(btnCopy, originalTooltipName);
          }, 1500);
        });
      });

      // for each tab sheet/json when tab clicked keep sheet index visible
      const tabsMain = document.querySelector('#tabs-main > .tab-content-holder') as HTMLElement;
      const tabs = tabsMain.querySelectorAll('input[type="radio"].tab');
      tabs.forEach((elem) => {
        elem.addEventListener('click', (event) => {
          const tabIndex = (event.target as HTMLInputElement).getAttribute('id').split('-').pop();
          tabs.forEach((tab: HTMLInputElement) => {
            const tabClick = tab.classList.contains(`tab-${tabIndex}`);
            if (tabClick) {
              tab.click();
            }
          });
        });
      });
    }

    if (event.data.pluginMessage.type === 'update-nodes-count' && nodesSelectedInfo && nodesSelectedCount && nodesSelectedLayersToChange && totalLayerChanges) {
      const selectedNodesInfo: SelectedNodesInfo = event?.data?.pluginMessage?.selectedNodesInfo;
      const totalLayerChangesCount: number = selectedNodesInfo.totalLayerChanges;
      const nodesLength = selectedNodesInfo.nodes.length;
      const nodesSelectedToChangeArray = selectedNodesInfo.typesQuantity.map((node: SelectedNodeType) => `<p class="space-x-xs"><span class="font-bold">${node.type}:</span><span>${node.count}</span></p>`).join('');

      if (nodesLength >= 0) {
        // 462        -> 00:41s = 41000
        // nodesCount -> x

        const estimatedTime = formatMsToTime((nodesLength * 41000) / 462);

        // nodesSelectedCount.innerText = `${nodesLength}`;
        // nodesSelectedEstimate.innerText = `${nodesLength > 150 ? `~= ${estimatedTime}` : ''}`;
        nodesSelectedCount.innerText = `${nodesLength !== 0 ? nodesLength > 150 ? `${nodesLength} ~= ${estimatedTime}` : `${nodesLength}` : ''}`;

        if (totalLayerChangesCount !== 0) {
          removeCls(totalLayerChanges, 'hidden');
        } else {
          addCls(totalLayerChanges, 'hidden');
        }

        if (nodesLength !== 0) {
          // removeCls(nodesSelectedInfo, 'hidden'); // enable these for extra layers info
        } else {
          // addCls(nodesSelectedInfo, 'hidden'); // enable these for extra layers info
        }
        nodesSelectedLayersToChange.innerHTML = nodesSelectedToChangeArray;
        totalLayerChanges.innerHTML = `${totalLayerChangesCount}`;
      }
    }

    // if (event.data.pluginMessage.type === 'open-preview') {
    //   togglePreview(false);
    // }

    // if (event.data.pluginMessage.type === 'close-preview') {
    //   togglePreview(false);
    // }

    // if (event.data.pluginMessage.type === 'start-loading') {
    //   if (event.data.pluginMessage.totalNodes) {
    //     totalNodes.innerText = event.data.pluginMessage.totalNodes;
    //   }
    //   loadingStart();
    // }

    // if (event.data.pluginMessage.type === 'end-loading') {
    //   loadingEnd();
    // }

    // if (event.data.pluginMessage.type === 'update-node-loading') {
    //   console.log('event.data.pluginMessage.nodesCount:', event.data.pluginMessage.nodesCount);
    //   if (event.data.pluginMessage.nodesCount) {
    //     nodesCount.innerText = event.data.pluginMessage.nodesCount;
    //     // loading.style.width = `calc(${} - 2px)`;
    //   }
    // }

    if (event.data.pluginMessage.type === 'get-api-data-sheet') {
      logger(LoggerType.UI, 'get-api-data-sheet');

      fetchData(event.data)
        .then((data) => {
          const parsedData = JSON.parse(data);
          const spreadsheetTitle = parsedData.properties.title;
          const spreadsheetDataSheetsInfo: SpreadsheetDataSheetsInfo[] = parsedData?.sheets?.map(sheet => {
            return {
              title: sheet.properties.title,
              rowCount: sheet.properties.gridProperties.rowCount,
              columnCount: sheet.properties.gridProperties.columnCount,
            }
          });
          spreadsheetData = {
            properties: {
              title: spreadsheetTitle
            },
            sheets: spreadsheetDataSheetsInfo
          };

          // updateSheetToggleButtons();
        })
        .catch((error) => {
          console.error('ERROR:', error);
        })
    }

    if (event.data.pluginMessage.type === 'get-api-data-values') {
      logger(LoggerType.UI, 'get-api-data-values');

      fetchData(event.data)
        .then(data => {
          logger(LoggerType.UI, 'get-api-data-values - data success');
          if (!event.data.pluginMessage.isCheckUrl) {
            let tableValues: SpreadsheetSheetsData;
            let normalizedForCode: any;

            try {
              const dataParsed = JSON.parse(data);
              if (dataParsed?.values) { // single sheet
                tableValues = [dataParsed.values];
              } else if (dataParsed?.valueRanges) { // multiple sheets
                tableValues = dataParsed.valueRanges.map(sheet => sheet.values);
              }
            } catch (e) {
              const csvTable = parseCsvToTable(data);
              tableValues = [csvTable];

              const url = event.data.pluginMessage.url as string;
              const title = getCsvTitleFromUrl(url);
              spreadsheetData = {
                properties: { title },
                sheets: [{ title, rowCount: csvTable.length, columnCount: csvTable[0]?.length || 0, selected: true }]
              } as any;
            }

            crateTableAndJsonTabsEachSheet(tableValues);

            if (tableValues.length === 1) {
              normalizedForCode = { values: tableValues[0] };
            } else {
              normalizedForCode = { valueRanges: tableValues.map(v => ({ values: v })) };
            }

            window.parent.postMessage(
              {
                pluginMessage: {
                  type: 'spreadsheet-data',
                  data: JSON.stringify(normalizedForCode),
                  isPreview: event.data.pluginMessage.isPreview,
                  isCheckUrl: event.data.pluginMessage.isCheckUrl,
                  spreadsheetData: event.data.pluginMessage.spreadsheetData
                },
              },
              '*'
            );
          }
        })
        .catch((error) => {
          console.error('ERROR:', error);
        })
    }
  }
}



// listeners --------------------

// listeners: add input url listeners (not present in simplified UI)
const apiUrlInput = document.getElementById('api-url') as HTMLInputElement | null;
if (apiUrlInput) {
  apiUrlInput.addEventListener('input', () => { debounce(checkUrlPublic()); updateInputUrl(); }, false);
  apiUrlInput.addEventListener('focus', (event) => { (event.target as HTMLInputElement).select(); updateInputUrl(); });
}

// listeners: CSV file upload
const csvInput = document.getElementById('csv-file') as HTMLInputElement;
if (csvInput) {
  csvInput.addEventListener('change', async () => {
    const file = csvInput.files && csvInput.files[0];
    if (!file) {
      uploadedCsvTable = null;
      uploadedCsvTitle = null;
      return;
    }
    const text = await file.text();
    uploadedCsvTable = parseCsvToTable(text);
    uploadedCsvTitle = (file.name || 'CSV').replace(/\.csv$/i, '') || 'CSV';

    // Build spreadsheet meta and send to plugin immediately
    spreadsheetData = {
      properties: { title: uploadedCsvTitle },
      sheets: [{ title: uploadedCsvTitle, rowCount: uploadedCsvTable.length, columnCount: uploadedCsvTable[0]?.length || 0, selected: true }] as any
    } as any;

    const normalizedForCode = { values: uploadedCsvTable };
    const renameNumbers = (document.getElementById('rename-numbers') as HTMLInputElement | null)?.checked || false;
    window.parent.postMessage(
      {
        pluginMessage: {
          type: 'spreadsheet-data',
          data: JSON.stringify(normalizedForCode),
          isPreview: false,
          isCheckUrl: false,
          spreadsheetData,
          renameNumbers
        }
      },
      '*'
    );

    // Allow selecting the same file again without needing to restart
    try { (csvInput as HTMLInputElement).value = ''; } catch {}
  });
}

// listeners: upload button to trigger hidden file input
const btnUploadCsv = document.getElementById('btn-upload-csv') as HTMLButtonElement | null;
if (btnUploadCsv && csvInput) {
  btnUploadCsv.addEventListener('click', () => csvInput.click());
}

// listeners: sync (not present in simplified UI)
const syncBtn = document.getElementById('sync') as HTMLButtonElement | null;
if (syncBtn) syncBtn.addEventListener('click', () => {
  logger(LoggerType.UI, 'Click: sync');
  if (uploadedCsvTable) {
    // Use uploaded CSV data directly
    const normalizedForCode = { values: uploadedCsvTable };
    const renameNumbers = (document.getElementById('rename-numbers') as HTMLInputElement | null)?.checked || false;
    window.parent.postMessage(
      {
        pluginMessage: {
          type: 'spreadsheet-data',
          data: JSON.stringify(normalizedForCode),
          isPreview: false,
          isCheckUrl: false,
          spreadsheetData,
          renameNumbers
        }
      },
      '*'
    );
  } else {
    const inputValue = getInputUrlValue();
    const renameNumbers = (document.getElementById('rename-numbers') as HTMLInputElement | null)?.checked || false;
    window.parent.postMessage(
      { pluginMessage: { type: 'get-data', url: inputValue, isPreview: false, isCheckUrl: false, spreadsheetData, renameNumbers } },
      '*'
    );
  }
});

// listeners: preview-data
const previewDataBtn = document.getElementById('preview-data') as HTMLButtonElement | null;
if (previewDataBtn) previewDataBtn.onclick = () => {
  logger(LoggerType.UI, 'Click: preview-data');
  togglePreview(isPreviewOpen ? false : true);
  if (uploadedCsvTable) {
    const normalizedForCode = { values: uploadedCsvTable };
    const renameNumbers = (document.getElementById('rename-numbers') as HTMLInputElement | null)?.checked || false;
    // Build spreadsheetData meta if missing
    if (!spreadsheetData) {
      const title = uploadedCsvTitle || 'CSV';
      spreadsheetData = {
        properties: { title },
        sheets: [{ title, rowCount: uploadedCsvTable.length, columnCount: uploadedCsvTable[0]?.length || 0, selected: true }] as any
      } as any;
    }
    crateTableAndJsonTabsEachSheet([uploadedCsvTable]);
    window.parent.postMessage(
      {
        pluginMessage: {
          type: 'spreadsheet-data',
          data: JSON.stringify(normalizedForCode),
          isPreview: true,
          isCheckUrl: false,
          spreadsheetData,
          renameNumbers
        }
      },
      '*'
    );
  } else {
    const inputValue = getInputUrlValue();
    const renameNumbers = (document.getElementById('rename-numbers') as HTMLInputElement | null)?.checked || false;
    window.parent.postMessage(
      { pluginMessage: { type: 'get-data', url: inputValue, isPreview: true, isCheckUrl: false, spreadsheetData, renameNumbers } },
      '*'
    );
  }
};

// listeners: open more info modal
const modalInfo = document.getElementById('more-info-modal') as HTMLElement | null;
const moreInfoBtn = document.getElementById('more-info') as HTMLElement | null;
if (moreInfoBtn && modalInfo) {
  moreInfoBtn.onclick = () => {
    modalInfo.style.display = 'block';
    removeCls(modalInfo, 'out');
  };
}

// listeners: close more info modal
if (modalInfo) {
  [modalInfo.querySelector('.btn-close-modal'), modalInfo.querySelector('.modal-overlay')].forEach(closeTriggers => {
    if (closeTriggers) closeTriggers.addEventListener('click', () => {
      addCls(modalInfo, 'out');

      const modalOverlay = modalInfo.querySelector('.modal-overlay');
      const style = getComputedStyle(modalOverlay as Element, 'animation');
      const styleAnimationDuration = parseFloat((style as any).animationDuration || '0');
      const styleAnimationDurationNumber = styleAnimationDuration * 1000;

      setTimeout(() => {
        modalInfo.style.display = 'none';
        removeCls(modalInfo, 'out');
      }, styleAnimationDurationNumber);
    })
  });
}

// listeners: re-check link
if (buttonReCheck) {
  buttonReCheck.onclick = () => {
    checkUrlPublic();
    toggleReCheckButton(true);
  };
}

// listeners: resize window (only when on advanced mode)
function resizeWindow(event) {
  const size: WindowSize = {
    w: Math.max(isPreviewOpen ? appWindowMinSizeAdvancedMode.w : appWindowMinSize.w, Math.floor(event.clientX + 5)),
    h: Math.max(isPreviewOpen ? appWindowMinSizeAdvancedMode.h : appWindowMinSize.h, Math.floor(event.clientY + 5))
  };
  parent.postMessage( { pluginMessage: { type: 'window-resize', size: size }}, '*');
}
if (cornerResize) {
  cornerResize.onpointerdown = (event) => {
    cornerResize.onpointermove = resizeWindow as any;
    (cornerResize as any).setPointerCapture((event as any).pointerId);
  };
  cornerResize.onpointerup = (event) => {
    cornerResize.onpointermove = null;
    (cornerResize as any).releasePointerCapture((event as any).pointerId);
  };
}

// listeners: add body class when SHIFT key is down
const body = document.querySelector('body');
window.addEventListener('keydown', (event) => {
  if (event.key.toLowerCase() === 'shift') {
    addCls(body, 'shift-pressed');
  }
});
window.addEventListener('keyup', (event) => {
  if (event.key.toLowerCase() === 'shift') {
    removeCls(body, 'shift-pressed');
  }
});

// listeners: when mouse enters the plugin window gove it focus
const wrapper = document.querySelector('.wrapper');
if (wrapper) {
  wrapper.addEventListener('mouseenter', () => {
    window.focus();
  });
}



// input --------------------

// input: get clean input value
function getInputUrlValue() {
  const textbox = document.getElementById('api-url') as HTMLInputElement | null;
  return (textbox?.value || '').trim();
}

// input: set input value local storage
function updateInputUrl() {
  const url = getInputUrlValue();
  const sectionSheetsUrl = document.getElementById('section-sheets-url') as HTMLElement | null;
  const urlValidStatus = document.getElementById('sheet-url-valid') as HTMLElement | null;

  if (sectionSheetsUrl && urlValidStatus) {
    removeClsStartingWith(sectionSheetsUrl, 'bg-');
    removeCls(urlValidStatus, ['text-success', 'text-warning', 'text-error']);
    addCls(urlValidStatus, 'hidden');

    if (url !== '') {
      const urlValid = checkUrlIsValid(url);

      if (urlValid) {
        sectionSheetsUrl.classList.add('bg-success');
        urlValidStatus.innerText = '✅ Url valid';
        addCls(urlValidStatus, 'text-success');
        removeCls(urlValidStatus, 'hidden');
      } else {
        sectionSheetsUrl.classList.add('bg-error');
        urlValidStatus.innerText = '⛔️ Url invalid';
        addCls(urlValidStatus, 'text-error');
        removeCls(urlValidStatus, 'hidden');
      }
    } else {
      sectionSheetsUrl.classList.add('bg-warning');
      urlValidStatus.innerText = '⚠️ Url not entered';
      addCls(urlValidStatus, 'text-warning');
      removeCls(urlValidStatus, 'hidden');
    }
  }

  window.parent.postMessage(
    { pluginMessage: { type: 'input-set', value: url } },
    '*'
  );
}



// helpers --------------------

function togglePreview(expand: boolean): void {
  logger(LoggerType.UI, 'togglePreview()');
  const sectionHowTo = document.getElementById('section-how-to') as HTMLElement | null;
  const previewSection = document.getElementById('preview-section') as HTMLElement | null;
  const previewDataBtn = document.getElementById('preview-data') as HTMLButtonElement | null;

  if (previewSection) {
    previewSection.classList.toggle('accordion-collapsed');
    previewSection.classList.toggle('accordion-expanded');
  }

  if (previewDataBtn) {
    previewDataBtn.classList.toggle('invisible');
  }

  if (cornerResize) cornerResize.classList.toggle('hidden');

  if (sectionHowTo) {
    sectionHowTo.classList.toggle('accordion-collapsed');
    sectionHowTo.classList.toggle('accordion-expanded');
  }

  // if is opened add tooltip listeners to images
  if (expand) {
    isPreviewOpen = true;
    window.parent.postMessage({ pluginMessage: { type: 'open-preview' } }, '*');
  } else {
    isPreviewOpen = false;
    window.parent.postMessage({ pluginMessage: { type: 'close-preview' } }, '*');
  }
}

function imageTooltipListeners(parentElem: HTMLElement | string): void {
  let elem: HTMLElement;

  if (typeof parentElem === 'string') {
    elem = document.getElementById(parentElem) as HTMLElement;
  } else {
    elem = parentElem;
  }

  if (elem) {
    const imageLinks = elem.querySelectorAll('.preview-link');
    const imageTooltip = document.getElementById('img-tooltip') as HTMLElement;

    Array.from(imageLinks).map(link => {
      link.addEventListener('mousemove', (event: MouseEvent) => {
        const imgUrl = (event.target as HTMLImageElement).getAttribute('data-image');

        imageTooltip.innerHTML = `<img src="${imgUrl}">`;
        imageTooltip.style.left = `${event.pageX + 10}px`;
        imageTooltip.style.top = `${event.pageY + 10}px`;
      });

      link.addEventListener('mouseover', (e) => {
        imageTooltip.style.opacity = '1';
      });
      link.addEventListener('mouseout', (e) => {
        imageTooltip.style.opacity = '0';
      });
    });
  }
}



// url --------------------

// url: check if url is a google spreadsheet valid url
function checkUrlIsValid(url: string): boolean {
  var validLink = new RegExp(/^(ftp|http|https):\/\/[^ "]+$/);
  const urlValid = validLink.test(url.trim());
  const urlPrefix = 'https://docs.google.com/spreadsheets';
  const isCsv = isCsvUrl(url);

  return urlValid && (url.startsWith(urlPrefix) || isCsv);
}

// url: check if google spreadsheet is public
function checkUrlPublic() {
  const url = getInputUrlValue();

  updatePreviewButton(url);

  if (!checkUrlIsValid(url)) { setUrlRequestStatus(RequestType.RESET); return; }

  // CSV doesn't require Google Sheets public check
  if (isCsvUrl(url)) {
    sheetIsPublic = true;
    setUrlRequestStatus(RequestType.SUCCESS, RequestMessage.SHEET_PUBLIC);
    updatePreviewButton(url);
    toggleReCheckButton(false);
    return;
  }

  window.parent.postMessage(
    { pluginMessage: { type: 'get-data', url, isPreview: false, isCheckUrl: true } },
    '*'
  );
}



// data --------------------

function fetchData(data: any): Promise<any> {
  setUrlRequestStatus(RequestType.RESET);

  const dataFetch = new Promise((resolve, reject) => {
    const request = new XMLHttpRequest();
    request.open('GET', data.pluginMessage.url);
    request.responseType = 'text';
    request.onerror = () => {
      setUrlRequestStatus(RequestType.ERROR, RequestMessage.INVALID_LINK);
      reject(RequestMessage.ERROR_GENERIC);
    }
    request.onload = () => {
      // Try JSON (Google Sheets). If parse fails, treat as CSV/text and resolve
      try {
        const jsonParsed = JSON.parse(request.response);
        if (jsonParsed) {
          if (jsonParsed.error) {
            sheetIsPublic = false;
            updatePreviewButton(getInputUrlValue());
            setUrlRequestStatus(RequestType.ERROR, RequestMessage.SHEET_NOT_PUBLIC);
            reject(RequestMessage.ERROR_GENERIC);
          } else {
            sheetIsPublic = true;
            updatePreviewButton(getInputUrlValue());
            setUrlRequestStatus(RequestType.SUCCESS, RequestMessage.SHEET_PUBLIC);
            resolve(request.response);
          }
        }
      } catch (e) {
        resolve(request.response);
      }

      toggleReCheckButton(false);
      // updateSheetToggleButtons();
    };
    request.send();
  });

  return dataFetch;
}

// function updateSheetToggleButtons(): void {
//   console.log('spreadsheetData:', spreadsheetData);
//   let sheetsSelectorContent: string = '';

//   if (spreadsheetData) {
//     spreadsheetData.sheets.map((sheet: SpreadsheetDataSheetsInfo, index: number) => {
//       sheetsSelectorContent = `
//         <div class="inline-flex items-center">
//           <input type="checkbox" id="check-sheet-${index + 1}" class="check-sheet m-0" name="check-sheet" value="check-sheet-${index + 1}" checked>
//           <label for="check-sheet-${index + 1}">${sheet.title}}</label>
//         </div>
//       `;
//     });
//   }
// }

function setUrlRequestStatus(requestType: RequestType, message: RequestMessage = null): void {
  const sectionHowTo = document.getElementById('section-how-to') as HTMLElement | null;
  const publishedStatus = document.getElementById('publish-status') as HTMLElement | null;

  if (requestType === RequestType.RESET) {
    if (sectionHowTo) removeClsStartingWith(sectionHowTo, 'bg-');
    if (publishedStatus) { removeCls(publishedStatus, ['text-success', 'text-error']); addCls(publishedStatus, 'hidden'); }
  }

  if (requestType === RequestType.ERROR) {
    if (sectionHowTo) { removeClsStartingWith(sectionHowTo, 'bg-'); addCls(sectionHowTo, 'bg-error'); }
    if (publishedStatus) { addCls(publishedStatus, 'text-error'); removeCls(publishedStatus, 'hidden'); publishedStatus.innerText = message; }
    toggleReCheckButton(false);
  }

  if (requestType === RequestType.SUCCESS) {
    if (sectionHowTo) { removeClsStartingWith(sectionHowTo, 'bg-'); addCls(sectionHowTo, 'bg-success'); }
    if (publishedStatus) { addCls(publishedStatus, 'text-success'); removeCls(publishedStatus, 'hidden'); publishedStatus.innerText = message; }
  }
}

// csv --------------------
function isCsvUrl(url: string): boolean {
  const lower = url.toLowerCase();
  if (lower.endsWith('.csv')) return true;
  try {
    const u = new URL(url);
    const pathnameCsv = u.pathname.toLowerCase().endsWith('.csv');
    const contentParamCsv = (u.searchParams.get('format') || u.searchParams.get('alt') || '').toLowerCase() === 'csv';
    return pathnameCsv || contentParamCsv;
  } catch {
    return false;
  }
}

function getCsvTitleFromUrl(url: string): string {
  try {
    const u = new URL(url);
    const pathParts = u.pathname.split('/').filter(Boolean);
    const filePart = pathParts[pathParts.length - 1] || 'CSV';
    return decodeURIComponent(filePart.replace(/\.csv$/i, '')) || 'CSV';
  } catch {
    return 'CSV';
  }
}

function parseCsvToTable(csvText: string): SpreadsheetSingleSheetData {
  // Simple RFC4180-ish parser supporting quoted fields with commas and quotes
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < csvText.length; i++) {
    const char = csvText[i];
    if (inQuotes) {
      if (char === '"') {
        if (csvText[i + 1] === '"') { field += '"'; i++; }
        else { inQuotes = false; }
      } else { field += char; }
    } else {
      if (char === '"') { inQuotes = true; }
      else if (char === ',') { row.push(field); field = ''; }
      else if (char === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
      else if (char === '\r') { /* ignore */ }
      else { field += char; }
    }
  }
  // push last field/row
  row.push(field);
  rows.push(row);

  // Trim possible empty trailing row
  if (rows.length && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') {
    rows.pop();
  }
  return rows as SpreadsheetSingleSheetData;
}




// table --------------------

function crateTableAndJsonTabsEachSheet(data: SpreadsheetSheetsData): void {
  data = formatDataFillEmpty(data);
  const tableHolder = document.getElementById('data-table-holder') as HTMLElement;
  const jsonHolder = document.getElementById('code-preview-holder') as HTMLElement;

  let spreadsheetSheetTabs = `<div class="tabs tabs-reversed tabs-sheets">`;
  let spreadsheetJsonTabs = `<div class="tabs tabs-reversed tabs-json">`;

  data.map((_, index) => {
    spreadsheetSheetTabs += `
      <input type="radio" class="tab tab-${index + 1}" name="tabs-sheets" id="tabs-sheets-${index + 1}" ${index === 0 ? 'checked' : ''}/>
      <label for="tabs-sheets-${index + 1}">${spreadsheetData.sheets[index].title}</label>
    `;
    spreadsheetJsonTabs += `
      <input type="radio" class="tab tab-${index + 1}" name="tabs-json" id="tabs-json-${index + 1}" ${index === 0 ? 'checked' : ''}/>
      <label for="tabs-json-${index + 1}">${spreadsheetData.sheets[index].title}</label>
    `;
  });

  spreadsheetSheetTabs += `<div class="tab-content-holder">`;
  spreadsheetJsonTabs += `<div class="tab-content-holder">`;

  data.map((_, index) => {
    spreadsheetSheetTabs += `
      <div class="tab-content tab-content-${index + 1}">
        <table id="data-table-${index + 1}"></table>
      </div>
    `;
    spreadsheetJsonTabs += `
      <div class="tab-content tab-content-${index + 1}">
        <pre id="code-preview-${index + 1}"></pre>
        <span id="btn-copy-to-clipboard-${index + 1}" class="btn-copy-to-clipboard" tooltip="Copy to Clipboard">
          <svg width="40" height="40" viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M14.9004 0H37.1272C38.7113 0 39.9999 1.282 40 2.85774V24.9683C40 26.5442 38.7113 27.8262 37.1272 27.8262H30.5949V25.2175H37.1272C37.1936 25.2174 37.2573 25.1911 37.3042 25.1444C37.3511 25.0977 37.3775 25.0344 37.3776 24.9683V2.85774C37.3775 2.79172 37.3511 2.72843 37.3042 2.68174C37.2572 2.63506 37.1936 2.60879 37.1272 2.6087H14.9004C14.8341 2.60879 14.7704 2.63506 14.7235 2.68174C14.6766 2.72843 14.6502 2.79172 14.6501 2.85774V9.56504H12.0277V2.85774C12.0277 1.282 13.3163 0 14.9004 0Z" fill="currentColor"/>
            <path fill-rule="evenodd" clip-rule="evenodd" d="M2.87285 12.1737H25.0996C26.6837 12.1737 27.9726 13.4557 27.9725 15.0316V37.1423C27.9725 38.718 26.6838 40 25.0997 40H2.87285C1.28874 40 8.74176e-05 38.718 0 37.1422V15.0316C0 13.4557 1.28874 12.1737 2.87285 12.1737ZM2.87285 37.3913H25.0997C25.1661 37.3912 25.2298 37.365 25.2767 37.3183C25.3237 37.2716 25.3501 37.2083 25.3502 37.1423H25.3501V15.0316C25.35 14.9655 25.3236 14.9022 25.2767 14.8555C25.2297 14.8088 25.1661 14.7825 25.0997 14.7824H2.87285C2.80645 14.7825 2.74279 14.8088 2.69583 14.8555C2.64888 14.9022 2.62248 14.9655 2.62241 15.0316V37.1423C2.6225 37.2083 2.64892 37.2716 2.69587 37.3183C2.74281 37.365 2.80646 37.3912 2.87285 37.3913Z" fill="currentColor"/>
          </svg>
        </span>
      </div>
    `;
  });

  spreadsheetSheetTabs += `
      </div>
    </div>
  `;
  spreadsheetJsonTabs += `
      </div>
    </div>
  `;

  data.map((_, index) => {
    spreadsheetJsonTabs += `
      <textarea id="code-to-copy-${index + 1}" class="code-to-copy" cols="30" rows="10"></textarea>
    `;
  });

  tableHolder.innerHTML = spreadsheetSheetTabs;
  jsonHolder.innerHTML = spreadsheetJsonTabs;

  data.map((sheet, index) => {
    createTable(`data-table-${index + 1}`, sheet);
  });
}

function createTable(tableElemId: string, data: SpreadsheetSingleSheetData): void {
  const table = document.getElementById(tableElemId) as HTMLTableElement;
  const extraActions = (layerName) => `
    <span class="cell-actions">
      <div class="button-holder" tooltip="${layerNamePrefix}${layerName}.rand" tooltipTable>
        <button data-type="rand" class="icon-only random"></button>
      </div>
      <div class="button-holder" tooltip="${layerNamePrefix}${layerName}.randsave" tooltipTable>
        <button data-type="randsave" class="icon-only random-save"></button>
      </div>
    </span>
  `;

  let tableHTML: string = '';
  data.map((elem, index) => {
    if (index === 0) {
      tableHTML += `
        <thead>
          <tr>
            <th class="cell-count"></th>
      `;

      elem.map((th, idx) => tableHTML += `
        <th>
          <div class="cell-inner">
            <div class="cell-name-holder" tooltip="${layerNamePrefix}${th}" tooltipTable>
              <span class="cell-name">${th}</span>
            </div>
            ${extraActions(th)}
          </div>
        </th>`);
      tableHTML += '</tr></thead>';
    } else {
      tableHTML += `
        <tbody>
          <tr>
            <td class="cell-count">${index}</td>
      `;

      elem.map((td, idx) => tableHTML += `
        <td class="${td === '' ? 'cell-empty' : ''}">
          <div class="cell-inner">
            <div class="cell-name-holder" tooltip="${layerNamePrefix}${data[0][idx]}.${index}" tooltipTable>
              <span class="cell-name">${isStringImage(td) ? `<div class="preview-link preview-link-alt" data-image="${td}">${td}</div>` : td === '' ? 'Empty' : td}</span>
            </div>
          </div>
        </td>`);
      tableHTML += '</tr>';

      if (index === data.length - 1) {
        tableHTML += '</tbody>';
      }
    }
  });
  table.innerHTML = tableHTML;

  // create tooltip for cells with images
  imageTooltipListeners(table);

  // create cell tooltips
  generateTooltips();
  updateTooltipSelections();

  // add table cell listeners
  createTableCellListeners(table, data);

  updateTableSelection();
}

function createTableCellListeners(table: HTMLTableElement, data: SpreadsheetSingleSheetData): void {
  const rows = table.querySelectorAll('tr');
  const rowsArray = Array.from(rows);

  table.addEventListener('click', (event) => {
    let target = event.target as HTMLElement;
    let labelNameSuffix = '';
    if (target.tagName.toLowerCase() === 'button') {
      labelNameSuffix = `.${target.getAttribute('data-type')}`;
      target = target.closest('th').querySelector('.cell-name');
    }

    const rowIndex = rowsArray.findIndex(row => row.contains(target));
    if (rowsArray[rowIndex]) {
      const columns = Array.from(rowsArray[rowIndex].querySelectorAll('th, td'));
      const columnIndex = columns.findIndex(column => {
        const colName = column.closest('th, td').querySelector('.cell-name');
        const targetName = target.parentElement.closest('th, td').querySelector('.cell-name');

        return colName === targetName;
      });
      const columnTite = data[0][columnIndex - 1];
      let value = '';

      if (rowIndex === 0) {
        value = `${layerNamePrefix}${columnTite}${labelNameSuffix}`;
      } else {
        value = `${layerNamePrefix}${columnTite}.${rowIndex}`;
      }
      // console.log(columnIndex, rowIndex, value);

      window.parent.postMessage(
        { pluginMessage: { type: 'table-elem-click', rowIndex, columnIndex, value, layerSelectionCount, isShift: event.shiftKey } },
        '*'
      );
    }
  });
}

function updateTableSelection(): void {
  const tableInfo = document.getElementById('table-info');

  if (spreadsheetData?.sheets) {
    const spreadsheetTitle =  tableInfo.querySelector('#spreadsheet-title');
    spreadsheetTitle.innerHTML = spreadsheetData.properties.title;

    spreadsheetData?.sheets.map((sheet, index) => {
      const table = document.getElementById(`data-table-${index + 1}`) as HTMLTableElement;

      if (table) {
        const tableToggleSelected = (onlyRemove: boolean): void => {
          if (!layerSelectionNames.some(l => l === undefined)) {
            const layerNamesArray = layerSelectionNames.map(name => name.split(layerNamePrefix).filter(str => str !== '').map(str => `#${str}`));
            let selectedArrayCount = 0;

            table.querySelectorAll('[tooltip]').forEach(elem => {
              const cell = elem.closest('th, td');
              const cellButtons = cell.querySelectorAll('.button-holder');
              if (cellButtons.length > 0) {
                cellButtons.forEach(btn => removeCls(btn, 'selected'));
              }

              if (cell.classList.contains('selected')) {
                removeCls(cell, 'selected');
              }

              if (!onlyRemove) {
                setTimeout(() => {
                  const cellName = elem.getAttribute('tooltip');
                  layerNamesArray.map(name => {
                    if (name.includes(cellName)) {
                      addCls(elem.closest('th, td'), 'selected')

                      if (elem.classList.contains('button-holder')) {
                        addCls(elem, 'selected')
                      }

                      selectedArrayCount++;
                    }
                  });

                  if (selectedArrayCount === 0) {
                    removeCls(table, 'selected-multiple');
                    removeCls(table, 'selected-single');
                  } else if (selectedArrayCount === 1) {
                    removeCls(table, 'selected-multiple');
                    addCls(table, 'selected-single');
                  } else if (selectedArrayCount > 1) {
                    addCls(table, 'selected-multiple');
                    removeCls(table, 'selected-single');
                  }
                });
              }
            });
          }
        }

        if (layerSelectionCount === 0) {
          removeCls(tableInfo, 'has-selection');
          tableToggleSelected(true);
        } else if (layerSelectionCount > 0) {
          removeCls(tableInfo, 'has-selection');
          addCls(tableInfo, 'has-selection');

          if (table.children.length > 0) {
            removeCls(table, 'selected-multiple');
            addCls(table, 'selected-multiple');

            tableToggleSelected(false);
          }
        }
      }
    });
  }
}



// buttons --------------------

// buttons: preview btn disabled state
function updatePreviewButton(url: string) {
  const buttonPreviewData = document.querySelector('#preview-data') as HTMLButtonElement | null;
  if (!buttonPreviewData) return;
  if ((uploadedCsvTable && uploadedCsvTable.length > 0) || (checkUrlIsValid(url) && sheetIsPublic)) {
    buttonPreviewData.disabled = false;
  } else {
    buttonPreviewData.disabled = true;
  }
}

function toggleReCheckButton(check: boolean): void {
  if (!buttonReCheck) return;
  if (check) {
    buttonReCheck.disabled = check;
    addCls(buttonReCheck, 'checking');
  } else {
    buttonReCheck.disabled = check;
    removeCls(buttonReCheck, 'checking');
  }
}



// tooltip --------------------

function generateTooltips(): void {
  const tooltips = document.querySelectorAll('[tooltip]:not([tooltip-gen])');

  Array.from(tooltips).map(tooltip => {
    tooltip.setAttribute('tooltip-gen', '');
    let title = tooltip.getAttribute('tooltip');

    if (tooltip.hasAttribute('tooltipTable')) {
      title = getTableTooltipStructure(title);
    }

    tippy(tooltip, {
      content: title,
      allowHTML: true,
      animation: 'shift-away-subtle',
      hideOnClick: false,
      theme: TooltipTheme.DEFAULT,
    });
  });
}

function updateTooltipSelections(): void {
  const tables = document.querySelectorAll('table');
  Array.from(tables).map(table => {
    const tooltips = table.querySelectorAll('[tooltip]');
    Array.from(tooltips).map((tooltip: HTMLElement) => {
      const title = tooltip.getAttribute('tooltip');

      if (tippy) {
        let rename = '';
        if (layerSelectionCount === 0) { rename = 'Select at least 1 layer'; tooltipUpdateTheme(tooltip, TooltipTheme.ERROR); }
        if (layerSelectionCount === 1) { rename = 'Rename 1 Layer'; tooltipUpdateTheme(tooltip); }
        if (layerSelectionCount > 1) { rename = `Rename ${layerSelectionCount} layers`; tooltipUpdateTheme(tooltip); }

        tooltipUpdateText(tooltip, getTableTooltipStructure(title, rename));
      }
    });
  });
}

function getTableTooltipStructure(title: string, rename: string = 'Rename'): string {
  return `<span class="tooltip-rename">${rename}</span><br><span class="tooltip-title">${title}</span>`;
}

function tooltipUpdateText(tooltip: HTMLElement, text: string): void {
  const tippy = tooltip['_tippy'];
  tippy.setContent(text);
}

function tooltipUpdateTheme(tooltip: HTMLElement, theme: TooltipTheme = TooltipTheme.DEFAULT): void {
  const tippy = tooltip['_tippy'];

  tippy.setProps({
    theme: theme
  });
}



// loading --------------------

// function loadingStart(): void {
//   modalLayersLoading.style.display = 'block';
//   removeCls(modalLayersLoading, 'out');
// }

// function loadingEnd(): void {
//   addCls(modalLayersLoading, 'out');

//   const modalOverlay = modalLayersLoading.querySelector('.modal-overlay');
//   const style = getComputedStyle(modalOverlay, 'animation');
//   const styleAnimationDuration = parseFloat(style.animationDuration);
//   const styleAnimationDurationNumber = styleAnimationDuration * 1000;

//   setTimeout(() => {
//     modalLayersLoading.style.display = 'none';
//     removeCls(modalLayersLoading, 'out');
//   }, styleAnimationDurationNumber);
// }



// utils --------------------

function debounce(func, wait = 1000, immediate = false) {
  var timeout;
  return function() {
      var context = this, args = arguments;
      var later = function() {
          timeout = null;
          if (!immediate) { func.apply(context, args); }
      };
      var callNow = immediate && !timeout;
      clearTimeout(timeout);
      timeout = setTimeout(later, wait);
      if (callNow) { func.apply(context, args); }
  };
}

function isStringImage(str: string): boolean {
  return str.toLowerCase().match(new RegExp(`^http(s)?://`, 'g')) ? true : false;
}

function removeClsStartingWith(elem: any | any[], strCls: string = 'bg-'): void {
  if (elem) {
    if (Array.isArray(elem)) {
      elem.map((item) => {
        Array.from(item.classList).map((cls: string) => {
          if (cls.startsWith(strCls)) {
            item.classList.remove(cls);
          }
        });
      });
    } else {
      Array.from(elem.classList).map((cls: string) => {
        if (cls.startsWith(strCls)) {
          elem.classList.remove(cls);
        }
      });
    }
  } else {
    console.error('ERROR (removeClsStartingWith): element dows not exist.', elem);
  }
}



// slider --------------------

// let isSliderMouseDown = false;
// let startX;
// let scrollLeft;
// const dragSpeed = 1;

// sheetsSelector.addEventListener('mousedown', (e) => {
//   isSliderMouseDown = true;
//   sheetsSelector.classList.add('active');
//   startX = e.pageX - sheetsSelector.offsetLeft;
//   scrollLeft = sheetsSelector.scrollLeft;
// });
// sheetsSelector.addEventListener('mouseleave', () => {
//   isSliderMouseDown = false;
//   sheetsSelector.classList.remove('active');
// });
// sheetsSelector.addEventListener('mouseup', () => {
//   isSliderMouseDown = false;
//   sheetsSelector.classList.remove('active');
// });
// sheetsSelector.addEventListener('mousemove', (e) => {
//   if(!isSliderMouseDown) return;
//   e.preventDefault();
//   const x = e.pageX - sheetsSelector.offsetLeft;
//   const walk = (x - startX) * dragSpeed; // scroll-fast
//   sheetsSelector.scrollLeft = scrollLeft - walk;
// });



// class --------------------

function addCls(elem: any | any[], cls: string | string[]): void {
  if (elem) {
    if (Array.isArray(elem)) {
      elem.map((item) => {
        if (Array.isArray(cls)) {
          cls.map((strCls) => {
            item.classList.add(strCls);
          });
        } else {
          item.classList.add(cls);
        }
      });
    } else {
      if (Array.isArray(cls)) {
        cls.map((strCls) => {
          elem.classList.add(strCls);
        });
      } else {
        elem.classList.add(cls);
      }
    }
  } else {
    console.error('ERROR (addCls): element dows not exist.', elem);
  }
}

function removeCls(elem: any | any[], cls: string | string[]): void {
  if (elem) {
    if (Array.isArray(elem)) {
      elem.map((item) => {
        if (Array.isArray(cls)) {
          cls.map((strCls) => {
            item.classList.remove(strCls);
          });
        } else {
          item.classList.remove(cls);
        }
      });
    } else {
      if (Array.isArray(cls)) {
        cls.map((strCls) => {
          elem.classList.remove(strCls);
        });
      } else {
        elem.classList.remove(cls);
      }
    }
  } else {
    console.error('ERROR (removeCls): element dows not exist.', elem);
  }
}
