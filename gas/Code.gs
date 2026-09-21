/**
 * Google Apps Script - 伝助→Googleカレンダー自動同期
 *
 * セットアップ手順:
 * 1. Google Apps Script (https://script.google.com) で新規プロジェクト作成
 * 2. このファイルの内容をコピー
 * 3. setupProperties() を実行（初期プロパティ作成）
 * 4. プロジェクトの設定 > スクリプト プロパティ で値を設定
 *    - ENABLED: true
 *    - DENSUKE_URL: 伝助URL
 *    - CALENDAR_ID: カレンダーID（primaryまたはカレンダーID）
 *    - TARGET_MEMBER: 対象メンバー名（空にするとチームモード: 個人の回答を表示しない）
 * 5. testScheduledSync() を実行してテスト
 * 6. setupHourlyTrigger() を実行してトリガー設定
 */

// ================================
// ★ スケジュール同期設定 ★
// ================================

/**
 * スクリプトプロパティから設定を取得
 * プロジェクトの設定 > スクリプト プロパティ で値を設定
 */
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  const targetMember = (props.getProperty('TARGET_MEMBER') || '').trim();
  return {
    enabled: props.getProperty('ENABLED') !== 'false',
    densukeUrl: props.getProperty('DENSUKE_URL') || '',
    calendarId: props.getProperty('CALENDAR_ID') || 'primary',
    targetMember: targetMember,
    // チームモード: TARGET_MEMBER が空のとき。タイトルの回答記号や「あなたの回答」を付けない
    teamMode: targetMember === ''
  };
}

/**
 * 初期設定用（一度だけ実行してテンプレートを作成）
 * 実行後、プロジェクトの設定 > スクリプト プロパティ から値を変更
 */
function setupProperties() {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'ENABLED': 'true',
    'DENSUKE_URL': 'https://densuke.biz/list?cd=XXXXXXXX',
    'CALENDAR_ID': 'primary',
    'TARGET_MEMBER': 'メンバー名'   // チーム共有カレンダー向けに使うときは空にする
  });
  console.log('✅ プロパティを設定しました');
  console.log('プロジェクトの設定 > スクリプト プロパティ から値を変更してください');
  console.log('設定項目: ENABLED, DENSUKE_URL, CALENDAR_ID, TARGET_MEMBER');
  console.log('TARGET_MEMBER を空にするとチームモード（個人の回答を表示しない）になります');
}

// ================================
// ★ トリガー設定 ★
// ================================

/**
 * 毎時00分にトリガーを設定
 * 初回のみ手動で実行してください
 */
function setupHourlyTrigger() {
  // 既存の doScheduledSync トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'doScheduledSync') {
      ScriptApp.deleteTrigger(trigger);
      console.log('Deleted existing trigger');
    }
  }
  
  // 毎時00分付近に実行するトリガーを設定
  ScriptApp.newTrigger('doScheduledSync')
    .timeBased()
    .everyHours(1)
    .nearMinute(0)  // 毎時00分付近で実行（±15分のズレあり）
    .create();
    
  console.log('✅ Hourly trigger set! Will run near minute 0 every hour.');
}

/**
 * トリガーを削除
 */
function removeAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'doScheduledSync') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  console.log('All doScheduledSync triggers removed');
}

// ================================
// ★ 動的トリガー管理 ★
// ================================

const DYNAMIC_TRIGGER_KEY = 'dynamicTriggerId';

/**
 * 動的トリガーを作成（X分後に実行）
 */
function createDynamicTrigger(minutes) {
  const triggerId = ScriptApp.newTrigger('doScheduledSync')
    .timeBased()
    .after(minutes * 60 * 1000)
    .create()
    .getUniqueId();
  
  PropertiesService.getScriptProperties().setProperty(DYNAMIC_TRIGGER_KEY, triggerId);
  console.log('Created dynamic trigger: ' + triggerId + ' (' + minutes + ' minutes)');
}

/**
 * 動的トリガーを削除
 */
function deleteDynamicTriggers() {
  const props = PropertiesService.getScriptProperties();
  const triggerId = props.getProperty(DYNAMIC_TRIGGER_KEY);
  
  if (triggerId) {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
        console.log('Deleted dynamic trigger: ' + triggerId);
        break;
      }
    }
    props.deleteProperty(DYNAMIC_TRIGGER_KEY);
  }
}

/**
 * 次回の同期タイミングをスケジュール
 */
function scheduleNextSync(events) {
  const now = new Date();
  let needsFrequentSync = false;
  let nextEvent = null;
  let nextEventTime = null;
  
  // イベントをチェック：開始1時間前～開始後20分の範囲
  for (const event of events) {
    for (const dateInfo of event.dates) {
      if (!dateInfo.startTime) continue; // 終日イベントはスキップ
      
      const eventTime = parseDateTime(dateInfo.date, dateInfo.startTime);
      const timeDiff = eventTime - now;
      
      // -20分 ～ +60分の範囲（開始1時間前～開始後20分）
      if (timeDiff >= -20 * 60 * 1000 && timeDiff <= 60 * 60 * 1000) {
        needsFrequentSync = true;
        console.log('Event near: ' + event.title + ' at ' + dateInfo.date + ' ' + dateInfo.startTime);
        break;
      }
      
      // 次の最も近いイベントを記録（まだ開始していない）
      if (timeDiff > 0 && (!nextEventTime || eventTime < nextEventTime)) {
        nextEvent = event;
        nextEventTime = eventTime;
      }
    }
    if (needsFrequentSync) break;
  }
  
  // 既存の動的トリガーを削除
  deleteDynamicTriggers();
  
  if (needsFrequentSync) {
    // 5分後のトリガーを作成
    createDynamicTrigger(5);
    console.log('Scheduled frequent sync (5 min)');
  } else if (nextEventTime) {
    // 次のイベントの1時間前にトリガーを設定
    const oneHourBefore = nextEventTime - 60 * 60 * 1000;
    const minutesUntilTrigger = Math.round((oneHourBefore - now) / (60 * 1000));
    
    if (minutesUntilTrigger > 0 && minutesUntilTrigger <= 24 * 60) {
      // 1分～24時間以内ならトリガー設定
      createDynamicTrigger(minutesUntilTrigger);
      console.log('Scheduled trigger at 1 hour before event: ' + nextEvent.title + ' (in ' + minutesUntilTrigger + ' min)');
    } else {
      console.log('Next event too far: ' + nextEvent.title + ' (in ' + Math.round(minutesUntilTrigger / 60) + ' hours)');
    }
  } else {
    console.log('No upcoming events, staying on hourly sync');
  }
}

/**
 * イベントを同期（追加・更新）
 */
function syncEvents(calendarId, events, config) {
  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    throw new Error('Calendar not found: ' + calendarId);
  }
  
  let added = 0;
  let updated = 0;
  let deleted = 0;
  
  const currentEventIds = events.map(e => e.id);
  
  // カレンダーの既存DENSUKEイベントを取得（今日から1年後まで）
  const now = new Date();
  const oneYearLater = new Date(now.getFullYear() + 1, now.getMonth(), now.getDate());
  const existingCalendarEvents = calendar.getEvents(now, oneYearLater);
  
  // 削除されるイベントを日付ごとに記録
  const deletedEventsByDate = {}; // { 'YYYY-MM-DD': [{ title, date }] }
  
  for (const calEvent of existingCalendarEvents) {
    const description = calEvent.getDescription() || '';
    const densukeIdMatch = description.match(/\[DENSUKE:([^\]]+)\]/);
    
    if (densukeIdMatch) {
      const densukeId = densukeIdMatch[1];
      if (!currentEventIds.includes(densukeId)) {
        const eventTitle = calEvent.getTitle();
        const eventDate = calEvent.getStartTime();
        const dateStr = formatDateStringFromDate(eventDate);
        
        console.log('Deleting event: ' + eventTitle + ' (ID: ' + densukeId + ')');
        calEvent.deleteEvent();
        
        // 日付ごとに削除イベントを記録
        if (!deletedEventsByDate[dateStr]) {
          deletedEventsByDate[dateStr] = [];
        }
        deletedEventsByDate[dateStr].push({ title: eventTitle, date: eventDate });
        
        deleted++;
      }
    }
  }
  
  // 追加されるイベントの日付を記録
  const addedEventsByDate = {}; // { 'YYYY-MM-DD': [eventData] }
  
  // イベントを追加/更新
  for (const event of events) {
    const existingEvents = findEventByDensukeId(calendar, event.id);
    
    if (existingEvents.length > 0) {
      // 既存イベントを更新
      updateEvent(existingEvents[0], event, config);
      updated++;
    } else {
      // 新規イベントを追加（通知は後でまとめて送信）
      const dateStr = event.date;
      
      if (!addedEventsByDate[dateStr]) {
        addedEventsByDate[dateStr] = [];
      }
      addedEventsByDate[dateStr].push(event);
      
      createEventWithoutNotification(calendar, event);
      added++;
    }
  }
  
  // 通知メールの送信
  sendNotifications(deletedEventsByDate, addedEventsByDate, config);

  return { added, updated, deleted };
}

/**
 * 日付をYYYY-MM-DD形式の文字列に変換（Dateオブジェクトから）
 */
function formatDateStringFromDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return year + '-' + month + '-' + day;
}

/**
 * 削除と追加の通知をまとめて送信
 */
function sendNotifications(deletedEventsByDate, addedEventsByDate, config) {
  try {
    // すべての日付を取得
    const allDates = new Set([
      ...Object.keys(deletedEventsByDate),
      ...Object.keys(addedEventsByDate)
    ]);

    for (const dateStr of allDates) {
      const deletedEvents = deletedEventsByDate[dateStr] || [];
      const addedEvents = addedEventsByDate[dateStr] || [];

      // 同一日に削除と追加の両方がある場合 → 更新通知
      if (deletedEvents.length > 0 && addedEvents.length > 0) {
        for (const event of addedEvents) {
          sendUpdateNotification(event, config);
        }
        // 削除メールは送らない
      } else {
        // 削除のみ → 削除通知
        if (deletedEvents.length > 0) {
          sendDeleteNotificationBatch(deletedEvents);
        }
        // 追加のみ → 追加通知
        for (const event of addedEvents) {
          sendAddNotification(event, config);
        }
      }
    }
  } catch (e) {
    console.log('Notification error: ' + e.message);
  }
}

/**
 * 更新通知メールを送信
 */
function sendUpdateNotification(eventData, config) {
  try {
    const emailBody =
      'イベントが更新されました。\n\n' +
      'タイトル: ' + eventData.title + '\n' +
      '日時: ' + eventData.date + (eventData.startTime ? ' ' + eventData.startTime : '') + '\n\n' +
      '伝助: ' + config.densukeUrl + '\n\n' +
      '詳細:\n' + eventData.description;

    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: '📝 イベント更新: ' + eventData.title,
      body: emailBody
    });
  } catch (e) {
    console.log('Email notification error: ' + e.message);
  }
}

/**
 * 追加通知メールを送信
 */
function sendAddNotification(eventData, config) {
  try {
    const emailBody =
      '新しいイベントがGoogleカレンダーに追加されました。\n\n' +
      'タイトル: ' + eventData.title + '\n' +
      '日時: ' + eventData.date + (eventData.startTime ? ' ' + eventData.startTime : '') + '\n\n' +
      '伝助: ' + config.densukeUrl + '\n\n' +
      '詳細:\n' + eventData.description;

    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: '📅 新規イベント追加: ' + eventData.title,
      body: emailBody
    });
  } catch (e) {
    console.log('Email notification error: ' + e.message);
  }
}

/**
 * 削除通知メールをまとめて送信
 */
function sendDeleteNotificationBatch(deletedEvents) {
  try {
    let emailBody = '以下のイベントが伝助から削除されたため、Googleカレンダーからも削除しました。\n\n';
    for (const evt of deletedEvents) {
      emailBody += '・' + evt.title + ' (' + evt.date.toLocaleDateString('ja-JP') + ')\n';
    }
    
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: '🗑️ イベント削除: ' + deletedEvents.length + '件',
      body: emailBody
    });
  } catch (e) {
    console.log('Email notification error: ' + e.message);
  }
}

/**
 * 通知なしで新規イベントを作成（リマインダー付き）
 */
function createEventWithoutNotification(calendar, eventData) {
  const description = eventData.description + '\n\n[DENSUKE:' + eventData.id + ']';
  
  let event;
  
  if (eventData.allDay) {
    // 終日イベント
    const date = new Date(eventData.date);
    event = calendar.createAllDayEvent(eventData.title, date, {
      description: description
    });
  } else {
    // 時間指定イベント
    const startDate = parseDateTime(eventData.date, eventData.startTime);
    const endDate = parseDateTime(eventData.date, eventData.endTime);
    
    event = calendar.createEvent(eventData.title, startDate, endDate, {
      description: description
    });
  }
  
  // リマインダーを設定（×欠席の場合は設定しない）
  if (event && !eventData.title.startsWith('×')) {
    try {
      console.log('Setting reminders for: ' + eventData.title);
      event.addPopupReminder(60);     // 1時間前にポップアップ
      event.addPopupReminder(1440);   // 1日前（24時間前）にポップアップ
      console.log('Reminders set successfully for: ' + eventData.title);
    } catch (e) {
      // リマインダー追加エラー
      console.log('Reminder error for ' + eventData.title + ': ' + e.message);
    }
  } else {
    console.log('Skipped reminders for: ' + eventData.title + ' (欠席または無効)');
  }
}

/**
 * 伝助IDでイベントを検索
 */
function findEventByDensukeId(calendar, densukeId) {
  // 過去1年から未来1年の範囲で検索
  const now = new Date();
  const startDate = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate());
  const endDate = new Date(now.getFullYear() + 1, now.getMonth(), now.getDate());
  
  const events = calendar.getEvents(startDate, endDate);
  const tag = '[DENSUKE:' + densukeId + ']';
  
  return events.filter(function(event) {
    const description = event.getDescription() || '';
    return description.includes(tag);
  });
}

/**
 * 既存イベントを更新
 */
function updateEvent(existingEvent, eventData, config) {
  const description = eventData.description + '\n\n[DENSUKE:' + eventData.id + ']';

  // タイトルから出欠回答記号を除去した原文を取得
  const oldTitle = existingEvent.getTitle();
  const oldOriginalTitle = removeResponseSymbol(oldTitle);
  const newOriginalTitle = removeResponseSymbol(eventData.title);

  // 伝助の原文が変更されたかチェック（出欠回答記号は無視）
  const titleChanged = (oldOriginalTitle !== newOriginalTitle);

  existingEvent.setTitle(eventData.title);
  existingEvent.setDescription(description);

  if (eventData.allDay) {
    existingEvent.setAllDayDate(new Date(eventData.date));
  } else {
    const startDate = parseDateTime(eventData.date, eventData.startTime);
    const endDate = parseDateTime(eventData.date, eventData.endTime);
    existingEvent.setTime(startDate, endDate);
  }

  // 伝助の原文が変更された場合に通知メールを送信
  if (titleChanged) {
    try {
      const emailBody =
        'イベントのタイトルが変更されました。\n\n' +
        '旧タイトル: ' + oldOriginalTitle + '\n' +
        '新タイトル: ' + newOriginalTitle + '\n\n' +
        '日時: ' + eventData.date + (eventData.startTime ? ' ' + eventData.startTime : '') + '\n\n' +
        '伝助: ' + config.densukeUrl + '\n\n' +
        '詳細:\n' + eventData.description;

      MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: '📝 イベント変更: ' + newOriginalTitle,
        body: emailBody
      });
    } catch (e) {
      console.log('Email notification error: ' + e.message);
    }
  }
}

/**
 * タイトルから出欠回答記号を除去
 */
function removeResponseSymbol(title) {
  // 先頭の出欠回答記号（◎○△×-）とスペースを除去
  return title.replace(/^[◎○△×\-]\s*/, '');
}

/**
 * 日付と時刻からDateオブジェクトを作成
 */
function parseDateTime(dateStr, timeStr) {
  const date = new Date(dateStr);
  if (timeStr) {
    const [hours, minutes] = timeStr.split(':').map(Number);
    date.setHours(hours, minutes, 0, 0);
  }
  return date;
}

// ============================================================
// スケジュール同期機能（トリガー用）
// ============================================================

/**
 * スケジュール同期のメイン関数
 * GASのトリガーから呼び出される
 */
function doScheduledSync() {
  const config = getConfig();

  if (!config.enabled) {
    console.log('Scheduled sync is disabled');
    return;
  }

  console.log('Starting scheduled sync at ' + new Date().toISOString());

  try {
    // 1. 伝助からHTMLを取得
    const html = fetchDensukeHtml(config.densukeUrl);
    if (!html) {
      throw new Error('Failed to fetch Densuke HTML');
    }

    // 2. HTMLを解析してイベントとコメントを抽出
    const { events, comments } = parseDensukeHtml(html, config.targetMember);
    console.log('Parsed ' + events.length + ' events, ' + comments.length + ' comments');

    // 3. カレンダー形式に変換（コメントも渡す）
    const calendarEvents = convertToCalendarEvents(events, comments, config);

    // 4. Googleカレンダーに同期
    const result = syncEvents(config.calendarId, calendarEvents, config);

    console.log('Scheduled sync completed: added=' + result.added + ', updated=' + result.updated + ', deleted=' + result.deleted);

    // 5. 次回の同期タイミングをスケジュール
    scheduleNextSync(events);

  } catch (error) {
    console.error('Scheduled sync error: ' + error.message);

    // エラー通知メール
    try {
      MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: '⚠️ 自動同期エラー',
        body: '伝助→Googleカレンダー同期でエラーが発生しました。\n\nエラー: ' + error.message
      });
    } catch (e) {
      console.error('Failed to send error email: ' + e.message);
    }
  }
}

/**
 * 伝助からHTMLを取得
 */
function fetchDensukeHtml(url) {
  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error('HTTP error: ' + response.getResponseCode());
    }
    
    return response.getContentText();
  } catch (error) {
    console.error('Fetch Densuke error: ' + error.message);
    return null;
  }
}

/**
 * HTMLエンティティをデコード
 */
function decodeHtmlEntities(text) {
  return text
    .replace(/&nbsp;/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .trim();
}

/**
 * 伝助HTMLを解析してイベントを抽出
 */
function parseDensukeHtml(html, targetMember) {
  const events = [];
  
  // テーブルの行を抽出
  const rowMatches = html.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];
  
  if (rowMatches.length < 2) {
    console.log('WARNING: Not enough rows in table');
    return events;
  }
  
  // メンバー名と集計列の位置を含むヘッダー行を探す
  let memberIndex = -1;
  let headerRowIndex = -1;
  let columnIndices = {
    excellent: -1,  // ◎
    good: -1,       // ○
    maybe: -1,      // △
    absent: -1      // ×
  };
  let firstMemberColumn = 1; // デフォルトはcells[1]から（集計列がない場合）
  
  const memberNamePart = targetMember.split('.').pop(); // "12.やまだ" -> "やまだ"
  
  for (let rowIdx = 0; rowIdx < Math.min(rowMatches.length, 5); rowIdx++) {
    const row = rowMatches[rowIdx];
    const cells = row.match(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi) || [];
    
    for (let i = 0; i < cells.length; i++) {
      const cellText = decodeHtmlEntities(cells[i].replace(/<[^>]+>/g, ''));
      
      // 集計列の位置を検出
      if (cellText === '◎') columnIndices.excellent = i;
      else if (cellText === '○' || cellText === '〇') columnIndices.good = i;
      else if (cellText === '△') columnIndices.maybe = i;
      else if (cellText === '×') columnIndices.absent = i;
      
      // メンバー名を検索（チームモード＝空文字のときは検索しない）
      if (memberNamePart !== '' && cellText.includes(memberNamePart)) {
        memberIndex = i;
        headerRowIndex = rowIdx;
        console.log('Found target member "' + targetMember + '" at index: ' + memberIndex + ' in row ' + rowIdx + ' (cell: ' + cellText + ')');
      }
    }
    
    // ヘッダー行は集計列かメンバー名が見つかった行
    if (memberIndex >= 0 || columnIndices.excellent >= 0) {
      headerRowIndex = rowIdx;
      break;
    }
  }
  
  // メンバー列の開始位置を特定（集計列の最後の次）
  const maxSummaryCol = Math.max(columnIndices.excellent, columnIndices.good, columnIndices.maybe, columnIndices.absent);
  if (maxSummaryCol >= 0) {
    firstMemberColumn = maxSummaryCol + 1;
    console.log('Summary columns detected: ◎=' + columnIndices.excellent + ', ○=' + columnIndices.good + ', △=' + columnIndices.maybe + ', ×=' + columnIndices.absent);
    console.log('First member column: ' + firstMemberColumn);
  } else {
    console.log('No summary columns detected, assuming members start at column 1');
  }
  
  if (memberNamePart === '') {
    console.log('Team mode: TARGET_MEMBER is empty, skipping member lookup');
  } else if (memberIndex < 0) {
    console.log('WARNING: Target member not found in header. Looking for: ' + targetMember);
  }
  
  // 前の行の時間と場所を保持（「上記、」パターン用）
  let previousTimeAndPlace = '';
  
  // データ行を処理（ヘッダー行の次から）
  let eventId = 0;
  const startRowIdx = headerRowIndex >= 0 ? headerRowIndex + 1 : 1;
  
  for (let rowIdx = startRowIdx; rowIdx < rowMatches.length; rowIdx++) {
    const row = rowMatches[rowIdx];
    
    // 集計行はスキップ
    if (row.includes('集計')) continue;
    
    const cells = row.match(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi) || [];
    if (cells.length < 6) continue;
    
    // 1列目: イベント情報
    let eventCell = decodeHtmlEntities(cells[0].replace(/<[^>]+>/g, ''));
    if (!eventCell || eventCell.length < 3) continue;
    
    // 日付パターンで始まるか確認
    if (!/^\d{1,2}\//.test(eventCell)) continue;
    
    // 元のタイトル（変換前）を保存
    const originalRawTitle = eventCell;
    
    // 「上記、」パターンの処理：前の行の時間と場所で置換
    if (eventCell.includes('上記、') && previousTimeAndPlace) {
      eventCell = eventCell.replace('上記、', previousTimeAndPlace + ' ');
    }
    
    // 時間と場所のパターンを抽出して保存（次の行で使用するため）
    const timeAndPlaceMatch = eventCell.match(/([^\d\/]*\d{1,2}:\d{2}-\d{1,2}:\d{2}[^◎〇△×]*)/);
    if (timeAndPlaceMatch) {
      previousTimeAndPlace = timeAndPlaceMatch[1]
        .replace(/\([日月火水木金土]\)/g, '')
        .trim();
    }
    
    // 出欠集計を取得（検出した列位置を使用、なければ0）
    let excellent = 0, good = 0, maybe = 0, absent = 0;
    
    if (columnIndices.excellent >= 0 && columnIndices.excellent < cells.length) {
      excellent = parseInt(decodeHtmlEntities(cells[columnIndices.excellent].replace(/<[^>]+>/g, ''))) || 0;
    }
    if (columnIndices.good >= 0 && columnIndices.good < cells.length) {
      good = parseInt(decodeHtmlEntities(cells[columnIndices.good].replace(/<[^>]+>/g, ''))) || 0;
    }
    if (columnIndices.maybe >= 0 && columnIndices.maybe < cells.length) {
      maybe = parseInt(decodeHtmlEntities(cells[columnIndices.maybe].replace(/<[^>]+>/g, ''))) || 0;
    }
    if (columnIndices.absent >= 0 && columnIndices.absent < cells.length) {
      absent = parseInt(decodeHtmlEntities(cells[columnIndices.absent].replace(/<[^>]+>/g, ''))) || 0;
    }
    
    // 各メンバーの回答を収集（メンバー名リスト用）
    const membersByResponse = {
      excellent: [],  // ◎
      good: [],       // ○
      maybe: [],      // △
      absent: [],     // ×
      unknown: []     // -
    };
    
    // ヘッダー行からメンバー名を取得し、対応する回答を収集
    if (headerRowIndex >= 0) {
      const headerRow = rowMatches[headerRowIndex];
      const headerCells = headerRow.match(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi) || [];
      
      // メンバー列から回答を収集（firstMemberColumnから開始）
      for (let i = firstMemberColumn; i < cells.length && i < headerCells.length; i++) {
        const memberName = decodeHtmlEntities(headerCells[i].replace(/<[^>]+>/g, ''));
        const response = decodeHtmlEntities(cells[i].replace(/<[^>]+>/g, ''));
        
        if (response === '◎') membersByResponse.excellent.push(memberName);
        else if (response === '○' || response === '〇') membersByResponse.good.push(memberName);
        else if (response === '△') membersByResponse.maybe.push(memberName);
        else if (response === '×') membersByResponse.absent.push(memberName);
        else membersByResponse.unknown.push(memberName);
      }
    }
    
    // 集計列がない場合は、メンバー回答から集計（フォールバック）
    if (columnIndices.excellent < 0) {
      excellent = membersByResponse.excellent.length;
      good = membersByResponse.good.length;
      maybe = membersByResponse.maybe.length;
      absent = membersByResponse.absent.length;
    }
    
    // メンバー総数から未回答者数を計算
    const totalMembers = cells.length - firstMemberColumn;
    const unknown = Math.max(0, totalMembers - (excellent + good + maybe + absent));
    
    const responseCounts = {
      excellent: excellent,
      good: good,
      maybe: maybe,
      absent: absent,
      unknown: unknown
    };
    
    // メンバーの回答を取得
    let responseSymbol = '-';
    if (memberIndex >= 0 && memberIndex < cells.length) {
      const responseText = decodeHtmlEntities(cells[memberIndex].replace(/<[^>]+>/g, ''));
      if (responseText && responseText !== '') {
        responseSymbol = responseText;
      }
    }
    
    console.log('Event: ' + eventCell + ', Response: ' + responseSymbol + ', Counts: ◎' + responseCounts.excellent + ' ○' + responseCounts.good + ' △' + responseCounts.maybe + ' ×' + responseCounts.absent);
    
    // イベント情報を解析
    const parsed = parseEventText(eventCell);
    if (parsed) {
      eventId++;
      events.push({
        id: 'event_' + eventId,
        rawTitle: eventCell,
        originalRawTitle: originalRawTitle,  // 変換前のオリジナルタイトル
        title: parsed.title,
        dates: parsed.dates,
        memberResponse: responseSymbol,
        responseCounts: responseCounts,
        membersByResponse: membersByResponse
      });
    }
  }
  
  // コメントを解析
  const comments = parseComments(html);
  
  return { events, comments };
}

/**
 * コメント欄を解析
 */
function parseComments(html) {
  const comments = [];
  
  // commentboxを抽出（改行を含む）
  const commentBoxMatch = html.match(/<div class="commentbox">([\s\S]*?)<\/div>/i);
  if (!commentBoxMatch) {
    console.log('No commentbox found');
    return comments;
  }
  
  console.log('Found commentbox: ' + commentBoxMatch[0].substring(0, 100) + '...');
  
  // commentbox全体を取得
  const commentBoxHtml = commentBoxMatch[0];
  
  // li要素を抽出（様々な形式に対応: </li>, <br></li>, <br /></li>, 改行など）
  // 方法1: </li>で終わるパターン
  let liMatches = commentBoxHtml.match(/<li>[\s\S]*?<\/li>/gi) || [];
  
  // 方法2: </li>がない場合、<li>から次の<li>または</ul>まで
  if (liMatches.length === 0) {
    console.log('Trying alternative li extraction...');
    // <li>タグで分割してから処理
    const liParts = commentBoxHtml.split(/<li>/i);
    liMatches = [];
    for (let i = 1; i < liParts.length; i++) {
      // </li>または</ul>または次の要素まで
      const endMatch = liParts[i].match(/^([\s\S]*?)(?:<\/li>|<\/ul>|$)/i);
      if (endMatch && endMatch[1].trim()) {
        liMatches.push('<li>' + endMatch[1] + '</li>');
      }
    }
  }
  
  console.log('Found ' + liMatches.length + ' li elements');
  
  for (const li of liMatches) {
    // HTMLタグを除去した全テキストを取得
    const plainText = li.replace(/<[^>]+>/g, '');
    console.log('Processing comment: ' + plainText);
    
    // 名前を抽出（<a>タグ内）
    const nameMatch = li.match(/<a[^>]*>([^<]+)<\/a>/);
    
    // コメント本文を抽出（</a>) の後から [ の前まで）
    const textMatch = plainText.match(/\)\s*(.+?)\s*\[/);
    
    // タイムスタンプを抽出
    const timestampMatch = plainText.match(/\[([^\]]+)\]/);
    
    if (nameMatch && textMatch) {
      comments.push({
        name: decodeHtmlEntities(nameMatch[1].trim()),
        text: decodeHtmlEntities(textMatch[1].trim()),
        timestamp: timestampMatch ? timestampMatch[1].trim() : ''
      });
      console.log('Added comment from: ' + nameMatch[1]);
    } else {
      console.log('Failed to parse comment. nameMatch: ' + !!nameMatch + ', textMatch: ' + !!textMatch);
    }
  }
  
  console.log('Parsed ' + comments.length + ' comments');
  return comments;
}

/**
 * イベントテキストを解析
 */
function parseEventText(text) {
  // まず日付情報を抽出
  // パターン1: 1/20(月)15:00-17:00 練習@○○小
  // パターン2: 3/20～22 合宿
  // パターン3: 3/20～21または21～22 学年合宿
  // パターン4: 1/20(月) 練習@○○小 (終日)
  
  const now = new Date();
  let year = now.getFullYear();
  
  // 「または」を含む日付範囲パターン: 3/20～21または21～22 学年合宿
  // 終了日の後の曜日も除去: 3/20(祝)～21(土)または21(土)～22(日) 学年合宿
  const orPattern = /^(\d{1,2})\/(\d{1,2})(?:\([^)]*\))?\s*[～~ー-]\s*(\d{1,2})(?:\([^)]*\))?\s*または\s*(\d{1,2})(?:\([^)]*\))?\s*[～~ー-]\s*(\d{1,2})(?:\([^)]*\))?\s*(.*)/;
  const orMatch = text.match(orPattern);
  
  if (orMatch) {
    const month = parseInt(orMatch[1]);
    const startDay1 = parseInt(orMatch[2]);
    const endDay1 = parseInt(orMatch[3]);
    const startDay2 = parseInt(orMatch[4]);
    const endDay2 = parseInt(orMatch[5]);
    const title = orMatch[6] || text;
    
    // 年を推定
    if (month < now.getMonth() + 1 - 6) {
      year++;
    }
    
    // 両方の範囲を含めて最小から最大までを展開
    const minDay = Math.min(startDay1, startDay2);
    const maxDay = Math.max(endDay1, endDay2);
    
    const dates = [];
    for (let day = minDay; day <= maxDay; day++) {
      dates.push({
        date: new Date(year, month - 1, day),
        startTime: null,
        endTime: null,
        allDay: true
      });
    }
    
    console.log('OR range event: ' + title + ', days ' + minDay + '-' + maxDay + ' (' + dates.length + ' days)');
    
    return {
      title: title.trim(),
      dates: dates
    };
  }
  
  // 日付範囲パターン: 3/20～22 合宿
  // 終了日の後の曜日も除去: 3/20(祝)～21(土) 学年合宿
  const rangePattern = /^(\d{1,2})\/(\d{1,2})(?:\([^)]*\))?\s*[～~ー-]\s*(\d{1,2})(?:\([^)]*\))?\s*(.*)/;
  const rangeMatch = text.match(rangePattern);
  
  if (rangeMatch) {
    const month = parseInt(rangeMatch[1]);
    const startDay = parseInt(rangeMatch[2]);
    const endDay = parseInt(rangeMatch[3]);
    const title = rangeMatch[4] || text;
    
    // 年を推定
    if (month < now.getMonth() + 1 - 6) {
      year++;
    }
    
    // 複数日に展開
    const dates = [];
    for (let day = startDay; day <= endDay; day++) {
      dates.push({
        date: new Date(year, month - 1, day),
        startTime: null,
        endTime: null,
        allDay: true
      });
    }
    
    console.log('Range event: ' + title + ', ' + dates.length + ' days');
    
    return {
      title: title.trim(),
      dates: dates
    };
  }
  
  // 時刻パターンを別途抽出 (HH:MM-HH:MM) - 絵文字があっても対応
  const timeMatch = text.match(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/);
  let startTime = null;
  let endTime = null;
  
  if (timeMatch) {
    startTime = timeMatch[1].padStart(2, '0') + ':' + timeMatch[2];
    endTime = timeMatch[3].padStart(2, '0') + ':' + timeMatch[4];
  }
  
  // 通常パターン: 1/20(月) 練習@○○小 または 1/20(月)🟥🟨15:00-17:00 練習@○○小
  const datePattern = /^(\d{1,2})\/(\d{1,2})(?:\([^)]*\))?/;
  const match = text.match(datePattern);
  
  if (!match) return null;
  
  const month = parseInt(match[1]);
  const day = parseInt(match[2]);
  
  // タイトルを抽出（日付と時間を除去、絵文字は保持）
  let title = text
    .replace(/^\d{1,2}\/\d{1,2}(?:\([^)]*\))?/, '') // 日付を除去
    .replace(/\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}/, '') // 時間を除去
    .trim();
  
  // タイトルが空の場合は元のテキストを使用
  if (!title) title = text;
  
  // 年を推定
  if (month < now.getMonth() + 1 - 6) {
    year++;
  }
  
  const date = new Date(year, month - 1, day);
  
  return {
    title: title,
    dates: [{
      date: date,
      startTime: startTime,
      endTime: endTime,
      allDay: startTime === null
    }]
  };
}

/**
 * イベントをカレンダー形式に変換
 */
function convertToCalendarEvents(events, comments, config) {
  const calendarEvents = [];
  const now = new Date();
  const oneWeekLater = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

  // コメントが渡されなかった場合は空配列
  comments = comments || [];

  for (const event of events) {
    const responseSymbol = getResponseSymbol(event.memberResponse);
    const counts = event.responseCounts || { excellent: 0, good: 0, maybe: 0, absent: 0, unknown: 0 };
    const membersByResponse = event.membersByResponse || { excellent: [], good: [], maybe: [], absent: [], unknown: [] };

    for (let i = 0; i < event.dates.length; i++) {
      const dateInfo = event.dates[i];
      const dateStr = formatDateString(dateInfo.date);
      const eventDate = dateInfo.date;

      // 日付とタイトル（出欠記号なし）からIDを生成して一意性を保証
      const originalTitle = event.title;
      const eventId = generateEventId(dateStr, originalTitle, i);

      // 1週間以内かどうかをチェック
      const isWithinOneWeek = eventDate <= oneWeekLater;

      // 詳細情報を構築
      let description = '日程: ' + dateStr + (dateInfo.startTime ? ' ' + dateInfo.startTime : '') + '\n';
      if (!config.teamMode) {
        description += 'あなたの回答: ' + responseSymbol + '\n';
        description += 'メンバー: ' + config.targetMember + '\n';
      }
      description += '\n【出欠状況】\n';

      // ◎
      description += '◎ ' + counts.excellent + '人\n';
      if (isWithinOneWeek && membersByResponse.excellent.length > 0) {
        description += membersByResponse.excellent.join('\n') + '\n\n';
      }

      // ○
      description += '○ ' + counts.good + '人\n';
      if (isWithinOneWeek && membersByResponse.good.length > 0) {
        description += membersByResponse.good.join('\n') + '\n\n';
      }

      // △
      description += '△ ' + counts.maybe + '人\n';
      if (isWithinOneWeek && membersByResponse.maybe.length > 0) {
        description += membersByResponse.maybe.join('\n') + '\n\n';
      }

      // ×
      description += '× ' + counts.absent + '人\n';
      if (isWithinOneWeek && membersByResponse.absent.length > 0) {
        description += membersByResponse.absent.join('\n') + '\n\n';
      }

      // -
      description += '- ' + counts.unknown + '人';
      if (isWithinOneWeek && membersByResponse.unknown.length > 0) {
        description += '\n' + membersByResponse.unknown.join('\n');
      }

      // コメント欄を追加（1週間以内のイベントのみ）
      if (isWithinOneWeek && comments.length > 0) {
        description += '\n\n【コメント】\n';
        for (const comment of comments) {
          description += '・' + comment.name + ': ' + comment.text;
          if (comment.timestamp) {
            description += ' [' + comment.timestamp + ']';
          }
          description += '\n';
        }
      }

      // 伝助タイトル（オリジナル）と同期時刻を最後に追加
      const syncTime = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
      description += '\n伝助URL: ' + config.densukeUrl;
      description += '\n伝助タイトル: ' + (event.originalRawTitle || event.rawTitle);
      description += '\n同期時刻: ' + syncTime;

      calendarEvents.push({
        id: eventId,
        densukeId: eventId,
        title: config.teamMode ? originalTitle : responseSymbol + ' ' + originalTitle,
        date: dateStr,
        startTime: dateInfo.startTime,
        endTime: dateInfo.endTime,
        allDay: dateInfo.allDay,
        description: description
      });
    }
  }

  return calendarEvents;
}

/**
 * イベントIDを生成
 * 日付とタイトルのハッシュから一意のIDを生成
 * これにより、同一日の複数イベントが安定して識別される
 */
function generateEventId(dateStr, title, index) {
  // タイトルのハッシュで安定したIDを生成
  const hash = simpleHash(title);
  return 'densuke_' + dateStr.replace(/-/g, '') + '_' + hash;
}

/**
 * 簡易ハッシュ関数
 */
function simpleHash(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  return Math.abs(hash).toString(16).slice(0, 6);
}

/**
 * 回答記号を取得
 */
function getResponseSymbol(response) {
  const symbol = response.trim();
  if (symbol === '◎' || symbol === '○' || symbol === '〇' || symbol === '△' || symbol === '×' || symbol === '-') {
    return symbol;
  }
  return '-';
}

/**
 * 日付をYYYY-MM-DD形式に変換
 */
function formatDateString(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return year + '-' + month + '-' + day;
}

/**
 * 手動テスト用関数
 * GASエディタで実行してテスト
 */
function testScheduledSync() {
  console.log('=== Test Scheduled Sync ===');
  console.log('Config:', JSON.stringify(getConfig()));
  doScheduledSync();
}

