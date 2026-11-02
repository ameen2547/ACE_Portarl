const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const projectRoot = path.resolve(__dirname, '..');
const backendSource = fs.readFileSync(path.join(projectRoot, 'code.gs'), 'utf8');
const frontendSource = fs.readFileSync(path.join(projectRoot, 'index.html'), 'utf8');
const TODAY = '2026-10-30';

function makeSheet(rows) {
  return {
    getDataRange() {
      return { getValues: () => rows };
    },
    getLastRow() {
      return rows.length;
    }
  };
}

function createHarness(liveRows, historyRows) {
  const properties = new Map();
  let archiveCallCount = 0;
  const context = {
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: () => null })
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: key => properties.get(key) || null,
        setProperty: (key, value) => properties.set(key, value)
      })
    },
    LockService: {
      getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} })
    },
    Utilities: {
      formatDate: value => value instanceof Date ? value.toISOString() : String(value)
    }
  };

  vm.createContext(context);
  vm.runInContext(backendSource, context, { filename: 'code.gs', timeout: 2000 });

  const liveSheet = makeSheet(liveRows);
  const historySheet = makeSheet(historyRows);
  context.requireSession_ = token => ({
    success: true,
    session: token === 'admin-token'
      ? { userName: 'Admin', userType: 'ADMIN' }
      : { userName: 'Kim', userType: 'VJ' }
  });
  context.getUserMappingDatabase_ = () => ({});
  context.getOrCreateUserMappingSheet_ = () => liveSheet;
  context.getOrCreateUserMappingHistorySheet_ = () => historySheet;
  context.userMappingSnapshotDate_ = () => TODAY;
  context.archiveUserMappingDailySnapshot_ = (_database, _rows, existingHistory) => {
    archiveCallCount++;
    return existingHistory;
  };

  return {
    context,
    get archiveCallCount() { return archiveCallCount; }
  };
}

const liveHeader = ['ID TikTok', 'User Name', 'TikTok URL', 'Assigned VJ', 'Team', 'Coins Today', 'Coins Month', 'Last Updated', 'Uploaded By', 'Total Days'];
const historyHeader = ['Snapshot Date', 'ID TikTok', 'User Name', 'TikTok URL', 'Assigned VJ', 'Team', 'Coins Today', 'Coins Month', 'Last Updated', 'Uploaded By', 'Total Days'];

test('VJ response includes only that VJ records in live and history data', () => {
  const live = [
    liveHeader,
    ['@kim-user', 'Kim User', '', 'Kim', 'Team A', 10, 10, TODAY + ' 09:00:00', 'Admin', 1],
    ['@lee-user', 'Lee User', '', 'Lee', 'Team B', 20, 20, TODAY + ' 09:00:00', 'Admin', 1]
  ];
  const history = [
    historyHeader,
    [TODAY, '@kim-user', 'Kim User', '', 'Kim', 'Team A', 10, 10, TODAY + ' 09:00:00', 'Admin', 1],
    [TODAY, '@lee-user', 'Lee User', '', 'Lee', 'Team B', 20, 20, TODAY + ' 09:00:00', 'Admin', 1]
  ];
  const harness = createHarness(live, history);
  const result = harness.context.getUserMappingData('ignored', 'ADMIN', 'vj-token', '127.0.0.1');

  assert.equal(result.success, true);
  assert.deepEqual(Array.from(result.list, row => row.assignedVJ), ['Kim']);
  assert.deepEqual(Array.from(result.historyList, row => row.assignedVJ), ['Kim']);
});

test('ADMIN response includes all live records and first history page', () => {
  const live = [liveHeader];
  const history = [historyHeader];
  for (let index = 1; index <= 3; index++) {
    live.push(['@user' + index, 'User ' + index, '', 'VJ ' + index, 'Team ' + index, index, index, TODAY, 'Admin', 1]);
    history.push([TODAY, '@user' + index, 'User ' + index, '', 'VJ ' + index, 'Team ' + index, index, index, TODAY, 'Admin', 1]);
  }
  const harness = createHarness(live, history);
  const result = harness.context.getUserMappingData('', 'VJ', 'admin-token', '127.0.0.1');

  assert.equal(result.success, true);
  assert.equal(result.userType, 'ADMIN');
  assert.equal(result.list.length, 3);
  assert.equal(result.historyList.length, 3);
  assert.equal(result.historyHasMore, false);
});

test('history is returned newest-first in pages of 100 and exposes older rows', () => {
  const history = [historyHeader];
  for (let index = 1; index <= 235; index++) {
    history.push([TODAY, '@user' + index, 'User ' + index, '', 'Kim', 'Team A', index, index, TODAY, 'Admin', index]);
  }
  const harness = createHarness([liveHeader], history);
  const first = harness.context.getUserMappingData('Kim', 'VJ', 'vj-token', '127.0.0.1');
  const second = harness.context.getUserMappingData('Kim', 'VJ', 'vj-token', '127.0.0.1', first.historyNextOffset);
  const third = harness.context.getUserMappingData('Kim', 'VJ', 'vj-token', '127.0.0.1', second.historyNextOffset);

  assert.equal(first.historyList.length, 100);
  assert.equal(first.historyHasMore, true);
  assert.equal(second.historyList.length, 100);
  assert.equal(second.historyHasMore, true);
  assert.equal(third.historyList.length, 35);
  assert.equal(third.historyHasMore, false);
  const allIds = [...first.historyList, ...second.historyList, ...third.historyList].map(row => row.tiktokId);
  assert.equal(new Set(allIds).size, 235);
  assert.equal(first.historyList[0].tiktokId, '@user235');
});

test('daily snapshot fallback is attempted at most once for the same day', () => {
  const harness = createHarness([liveHeader, ['@kim-user', 'Kim User', '', 'Kim', 'Team A', 1, 1, TODAY, 'Admin', 1]], [historyHeader]);
  harness.context.getUserMappingData('Kim', 'VJ', 'vj-token', '127.0.0.1');
  harness.context.getUserMappingData('Kim', 'VJ', 'vj-token', '127.0.0.1');

  assert.equal(harness.archiveCallCount, 1);
});

test('frontend includes transient retry, bridge guard, and history continuation control', () => {
  assert.ok(/HTTP\s*-1/.test(frontendSource));
  assert.ok(frontendSource.includes('if (!hasAppsScriptBridge())'));
  assert.ok(frontendSource.includes("loadOlderUserMapHistory('VJ')"));
  assert.ok(frontendSource.includes("loadOlderUserMapHistory('ADMIN')"));
});

test('inline frontend JavaScript parses without syntax errors', () => {
  const scriptTag = /<script\b([^>]*)>([\s\S]*?)<\/script>/gi;
  let match;
  let parsedCount = 0;
  while ((match = scriptTag.exec(frontendSource))) {
    if (/\bsrc\s*=/.test(match[1]) || !match[2].trim()) continue;
    parsedCount++;
    assert.doesNotThrow(() => new vm.Script(match[2], { filename: 'index-inline-' + parsedCount + '.js' }));
  }
  assert.ok(parsedCount > 0, 'expected at least one inline application script');
});

test('both User Mapping views provide report filters, summary cards, and history tabs', () => {
  for (const id of [
    'vjUserMapTeamFilter', 'vjUserMapDateFrom', 'vjUserMapDateTo', 'vjUmMetricUsers', 'vjUserMapTabHistory',
    'admUserMapTeamFilter', 'admUserMapDateFrom', 'admUserMapDateTo', 'admUmMetricUsers', 'admUserMapTabHistory'
  ]) {
    assert.ok(frontendSource.includes('id="' + id + '"'), 'missing report UI element: ' + id);
  }
  assert.ok(frontendSource.includes("setUserMapQuickRange('VJ','week')"));
  assert.ok(frontendSource.includes("setUserMapQuickRange('ADMIN','month')"));
  assert.ok(frontendSource.includes('function renderUserMapTeamSummary'));
  assert.ok(frontendSource.includes('var USER_MAPPING_PAGE_SIZE = 50;'));
  assert.ok(frontendSource.includes('id="vjUserMapPageText"'));
  assert.ok(frontendSource.includes('id="admUserMapPageText"'));
  assert.ok(frontendSource.includes('function changeUserMapCurrentPage'));
});

function createGiftReportHarness(giftRows, mappingRows, sessionType, sessionName) {
  const reads = { databaseOpen: 0, giftRead: 0, mappingRead: 0 };
  function sheet(rows, type) {
    return {
      getLastRow: () => rows.length,
      getLastColumn: () => rows[0] ? rows[0].length : 0,
      getRange: (row, column, rowCount, columnCount) => ({
        getValues: () => {
          reads[type]++;
          return rows.slice(row - 1, row - 1 + rowCount).map(values => values.slice(column - 1, column - 1 + columnCount));
        }
      })
    };
  }
  const giftSheet = sheet(giftRows, 'giftRead');
  const mappingSheet = sheet(mappingRows, 'mappingRead');
  const database = {
    getSheetByName: name => name === 'Gift_Details' ? giftSheet : (name === 'User_Mapping' ? mappingSheet : null)
  };
  const context = { console, Logger: { log: () => {} } };
  vm.createContext(context);
  vm.runInContext(backendSource, context, { filename: 'code.gs', timeout: 2000 });
  context.requireSession_ = () => ({ success: true, session: { userName: sessionName, userType: sessionType } });
  context.getUserMappingDatabase_ = () => { reads.databaseOpen++; return database; };
  return { context, reads };
}

const giftHeader = ['Record Key', 'Event Time', 'Round', 'VJ', 'Gifter Name', 'TikTok ID', 'Bound', 'Gift Name', 'Quantity', 'Unit Price', 'Total Coins', 'Collection ID', 'Uploaded By'];

test('Gift_Details report joins TikTok ID to User_Mapping without changing the sheet schema', () => {
  const gifts = [
    giftHeader,
    ['r1', '2026-10-31 10:00:00', 'R1', 'Kim', 'Gifter A', '@giver-a', 'Bound', 'Rose', 2, 1, 2, 'c1', 'uploader'],
    ['r2', '2026-10-31 11:00:00', 'R1', 'Lee', 'Gifter B', '@giver-b', 'Bound', 'Lion', 1, 10, 10, 'c2', 'uploader'],
    ['r3', '2026-10-30 09:00:00', 'R2', 'Kim', 'Gifter C', 'giver-c', 'Bound', 'Heart', 4, 1, 4, 'c3', 'uploader']
  ];
  const mappings = [
    liveHeader,
    ['@giver-a', 'Mapped A', '', 'Kim', 'Team A', 0, 0, TODAY, 'Admin', 0],
    ['@giver-c', 'Mapped C', '', 'Kim', 'Team A', 0, 0, TODAY, 'Admin', 0]
  ];
  const harness = createGiftReportHarness(gifts, mappings, 'ADMIN', 'Admin');
  const result = harness.context.getGiftDetailsData('admin-token', '127.0.0.1', { team: 'Team A' });

  assert.equal(result.success, true);
  assert.equal(result.totalRecords, 2);
  assert.equal(result.summary.totalCoins, 6);
  assert.equal(result.summary.mappedRecords, 2);
  assert.equal(result.list.length, 2);
  assert.equal(result.list[0].mappedUserName, 'Mapped A');
  assert.equal(result.summaryByMapping[0].tiktokId, 'giver-c');
  assert.equal(harness.reads.databaseOpen, 1);
  assert.equal(harness.reads.giftRead, 1);
  assert.equal(harness.reads.mappingRead, 1);
});

test('VJ Gift_Details report only returns own VJ rows and own mapped rows with blank VJ', () => {
  const gifts = [
    giftHeader,
    ['r1', '2026-10-31 10:00:00', 'R1', 'Kim', 'Gifter A', '@giver-a', 'Bound', 'Rose', 2, 1, 2, 'c1', 'uploader'],
    ['r2', '2026-10-31 11:00:00', 'R1', 'Lee', 'Gifter B', '@giver-b', 'Bound', 'Lion', 1, 10, 10, 'c2', 'uploader'],
    ['r3', '2026-10-30 09:00:00', 'R2', '', 'Gifter C', '@giver-c', 'Bound', 'Heart', 4, 1, 4, 'c3', 'uploader']
  ];
  const mappings = [
    liveHeader,
    ['@giver-a', 'Mapped A', '', 'Kim', 'Team A', 0, 0, TODAY, 'Admin', 0],
    ['@giver-b', 'Mapped B', '', 'Lee', 'Team B', 0, 0, TODAY, 'Admin', 0],
    ['@giver-c', 'Mapped C', '', 'Kim', 'Team A', 0, 0, TODAY, 'Admin', 0]
  ];
  const harness = createGiftReportHarness(gifts, mappings, 'VJ', 'Kim');
  const result = harness.context.getGiftDetailsData('vj-token', '127.0.0.1', {});

  assert.equal(result.success, true);
  assert.equal(result.totalRecords, 2);
  assert.deepEqual(Array.from(result.list, row => row.tiktokId), ['@giver-a', '@giver-c']);
  assert.equal(result.summary.totalCoins, 6);
  assert.equal(result.list.some(row => row.tiktokId === '@giver-b'), false);
});

test('VJ Gift_Details rows do not reveal User Mapping details owned by a different VJ', () => {
  const gifts = [
    giftHeader,
    ['r1', '2026-10-31 10:00:00', 'R1', 'Kim', 'Gifter B', '@giver-b', 'Bound', 'Lion', 1, 10, 10, 'c2', 'uploader']
  ];
  const mappings = [
    liveHeader,
    ['@giver-b', 'Mapped B', '', 'Lee', 'Team B', 0, 0, TODAY, 'Admin', 0]
  ];
  const harness = createGiftReportHarness(gifts, mappings, 'VJ', 'Kim');
  const result = harness.context.getGiftDetailsData('vj-token', '127.0.0.1', {});

  assert.equal(result.totalRecords, 1);
  assert.equal(result.list[0].isMapped, false);
  assert.equal(result.list[0].mappedUserName, '');
  assert.equal(result.list[0].team, '');
  assert.equal(result.summaryByMapping.length, 0);
});

test('Gift_Details report pages rows and keeps whole-filter totals', () => {
  const gifts = [giftHeader];
  for (let index = 1; index <= 51; index++) {
    gifts.push(['r' + index, '2026-10-31 10:00:00', 'R1', 'Kim', 'Gifter ' + index, 'gifter-' + index, 'Bound', 'Rose', 1, 1, 1, 'c' + index, 'uploader']);
  }
  const harness = createGiftReportHarness(gifts, [liveHeader], 'ADMIN', 'Admin');
  const firstPage = harness.context.getGiftDetailsData('admin-token', '127.0.0.1', { page: 1 });
  const secondPage = harness.context.getGiftDetailsData('admin-token', '127.0.0.1', { page: 2 });

  assert.equal(firstPage.list.length, 50);
  assert.equal(firstPage.totalPages, 2);
  assert.equal(firstPage.totalRecords, 51);
  assert.equal(firstPage.summary.totalCoins, 51);
  assert.equal(secondPage.list.length, 1);
  assert.equal(secondPage.page, 2);
  assert.equal(Object.hasOwn(firstPage.list[0], 'sourceRow'), false);
});

test('Gift_Details report resolves gift art by name, validates CDN host, and caches the lookup', () => {
  const gifts = [
    giftHeader,
    ['r1', '2026-10-31 10:00:00', 'R1', 'Kim', 'Gifter A', '@giver-a', 'Bound', 'Rose', 1, 1, 1, 'c1', 'uploader'],
    ['r2', '2026-10-31 11:00:00', 'R1', 'Kim', 'Gifter B', '@giver-b', 'Bound', 'Unsafe Gift', 1, 1, 1, 'c2', 'uploader']
  ];
  const harness = createGiftReportHarness(gifts, [liveHeader], 'ADMIN', 'Admin');
  const cacheValues = new Map();
  let fetchCount = 0;
  harness.context.CacheService = {
    getScriptCache: () => ({
      get: key => cacheValues.get(key) || null,
      put: (key, value) => cacheValues.set(key, value)
    })
  };
  harness.context.UrlFetchApp = {
    fetch: () => {
      fetchCount++;
      return {
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify([
          { name: 'Rose', image: { url_list: ['https://p16-webcast.tiktokcdn.com/img/rose.webp'] } },
          { name: 'Unsafe Gift', image: { url_list: ['https://images.attacker.example/unsafe.png'] } }
        ])
      };
    }
  };

  const first = harness.context.getGiftDetailsData('admin-token', '127.0.0.1', {});
  const second = harness.context.getGiftDetailsData('admin-token', '127.0.0.1', {});
  const byName = Object.fromEntries(Array.from(first.list, item => [item.giftName, item.giftImageUrl]));

  assert.equal(byName.Rose, 'https://p16-webcast.tiktokcdn.com/img/rose.webp');
  assert.equal(byName['Unsafe Gift'], '');
  assert.equal(second.list.find(item => item.giftName === 'Rose').giftImageUrl, byName.Rose);
  assert.equal(fetchCount, 1);
});

test('both User Mapping views expose Gift_Details report tabs and filtered pages', () => {
  for (const id of [
    'vjUserMapTabGifts', 'vjUserMapGiftsPanel', 'vjGiftDateFrom', 'vjGiftTeamFilter', 'vjGiftReportTbody',
    'admUserMapTabGifts', 'admUserMapGiftsPanel', 'admGiftDateFrom', 'admGiftTeamFilter', 'admGiftReportTbody'
  ]) {
    assert.ok(frontendSource.includes('id="' + id + '"'), 'missing Gift_Details UI element: ' + id);
  }
  assert.ok(frontendSource.includes('.getGiftDetailsData(sessionToken, clientIpAddress, options)'));
  assert.ok(frontendSource.includes('function renderGiftDetailsReport'));
  assert.ok(frontendSource.includes('function giftReportProfileCell'));
  assert.ok(frontendSource.includes('function giftReportGiftCell'));
  assert.ok(frontendSource.includes('var cachedMappingAvatarUrls'));
});
