// ============================================================
//  予約自動管理スクリプト v5.0（整理版）
//  対応: Klook / Bokun / GetYourGuide / 自社サイト(EN・Square)
//
//  列: A体験日 B開始時間 Cコース D人数 E代表者名 F予約番号 Gステータス
//      H受信日時 I経路 J変更前日時 K支払い済 L担当者 Mレビュー Nメモ
//      O来店済 P電話済 Q電話 R国 Sメール T店舗
//
//  ★ 初回だけ setupTriggers() を実行してください
// ============================================================

const CONFIG = {
  SPREADSHEET_ID:     '1233R0p0EdKEFKQLaODXKuy0KneXAHmT0EFIlpYO6aS4',
  BOOKING_SHEET:      '予約一覧',
  TIMELINE_SHEET:     '本日のタイムライン',
  STAFF_SHEET:        '担当者一覧',
  GMAIL_QUERY:        '(from:operator@klook.com OR from:no-reply@bokun.io OR from:notification.getyourguide.com OR from:fivebet.23@gmail.com OR from:onboarding@resend.dev OR subject:"Japan ticket")',
  PROCESSED_LABEL:    'Klook/処理済み_COPY',
  MODIFICATION_LABEL: '予約変更',
  TZ:                 'Asia/Tokyo',
  FILTER_DAYS:        7,    // シートに表示する範囲（前後○日）
  BATCH_SIZE:         20,   // 5分ごとの処理件数
  SYNC_DAYS:          3,    // 安全網チェックで見る日数
};

const SHOP_KYOTO = '京都';
const SHOP_OSAKA = '大阪';
const SHOP_SUMO  = '相撲';
const COL_SHOP   = 20;   // T列

const CAPACITY_LIMIT = 25;
const ALERT_EMAILS = ['yuxinyuey@gmail.com', 'haruchan.20032003@gmail.com'];


// ============================================================
//  トリガー設定（最初に1回だけ実行）
// ============================================================
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });

  // 旧インポート・バックフィルの残りフラグを掃除（残っていると処理が止まる）
  const props = PropertiesService.getScriptProperties();
  ['IMPORT_RUNNING', 'IMPORT_OFFSET', 'IMPORT_ROWS', 'IMPORT_CANCELS',
   'BACKFILL_OFFSET', 'BACKFILL_KLOOK_OFFSET', 'BACKFILL_BOKUN_OFFSET',
   'BACKFILL_MISSING_ROW', 'BACKFILL_CONTACT_ROW'].forEach(function(k) { props.deleteProperty(k); });

  ScriptApp.newTrigger('processKlookEmails').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('syncRecentEmails').timeBased().everyHours(2).create();
  ScriptApp.newTrigger('checkHourlyCapacity').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('generateDailyTimeline').timeBased().everyDays(1).atHour(7).inTimezone(CONFIG.TZ).create();
  ScriptApp.newTrigger('updateAnalytics').timeBased().everyDays(1).atHour(7).inTimezone(CONFIG.TZ).create();

  Logger.log('✅ トリガー設定完了');
  Logger.log('  5分おき : processKlookEmails（新着処理）');
  Logger.log('  2時間おき: syncRecentEmails（取りこぼし補完）');
  Logger.log('  1時間おき: checkHourlyCapacity（人数超過アラート）');
  Logger.log('  毎朝7時 : generateDailyTimeline / updateAnalytics');
}


// ============================================================
//  ① 新着メール処理（5分おき）
// ============================================================
function processKlookEmails() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) { Logger.log('他の処理が実行中のためスキップ'); return; }
  try {
    const label = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
    const base  = CONFIG.GMAIL_QUERY + ' -label:' + CONFIG.PROCESSED_LABEL;

    // 直近3日を優先 → 空きがあれば古いもの
    let threads = GmailApp.search(base + ' newer_than:3d', 0, CONFIG.BATCH_SIZE);
    if (threads.length < CONFIG.BATCH_SIZE) {
      threads = threads.concat(GmailApp.search(base + ' older_than:3d', 0, CONFIG.BATCH_SIZE - threads.length));
    }
    if (threads.length === 0) { Logger.log('新着メールなし'); return; }

    const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
    ensureBookingHeader(sheet);

    const parsed   = collectParsed_(threads, null);
    const failed   = {};
    const modified = {};

    parsed.forEach(function(item) {
      try {
        if (applyBooking_(sheet, item.data) === 'modified') modified[item.ti] = true;
      } catch (e) {
        Logger.log('反映エラー: ' + item.data.bookingId + ' → ' + e.message);
        failed[item.ti] = true;
      }
    });

    threads.forEach(function(thread, ti) {
      if (failed[ti]) return;   // 失敗したスレッドは次回再試行
      thread.addLabel(label);
      if (modified[ti]) thread.addLabel(getOrCreateLabel(CONFIG.MODIFICATION_LABEL));
    });

    applyDateFilter(sheet);
    Logger.log(threads.length + 'スレッドを処理しました');
  } finally {
    lock.releaseLock();
  }
}


// ============================================================
//  ② 取りこぼし補完（2時間おき）
//  処理済みラベルに関係なく直近メールを見直し、シートに無い予約・
//  反映されていない変更/キャンセルを補完する。
//  （同じスレッドに後から届いたキャンセル・変更メールもここで拾う）
// ============================================================
function syncRecentEmails() {
  syncEmails_(CONFIG.SYNC_DAYS);
}

// 手動用：過去30日分をまとめて見直す（取りこぼしが疑われる時に実行）
function 取りこぼしを30日分チェック() {
  syncEmails_(30);
}

function syncEmails_(days) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('他の処理が実行中のためスキップ'); return; }
  try {
    const t0    = Date.now();
    const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
    ensureBookingHeader(sheet);

    // シート上の 予約番号 → 受信日時
    const known   = {};
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 6, lastRow - 1, 3).getValues().forEach(function(r) {
        const id = String(r[0]).trim();
        if (id) known[id] = String(r[2] || '');
      });
    }

    const since   = new Date(Date.now() - days * 86400000);
    const threads = GmailApp.search(CONFIG.GMAIL_QUERY + ' newer_than:' + days + 'd', 0, 500);
    const parsed  = collectParsed_(threads, since);

    let fixed = 0;
    for (let i = 0; i < parsed.length; i++) {
      if (Date.now() - t0 > 5 * 60 * 1000) { Logger.log('時間切れ。もう一度実行してください'); break; }
      const d = parsed[i].data;
      if (known[d.bookingId] !== undefined && d.receivedAt <= known[d.bookingId]) continue;  // 反映済み
      try {
        const result = applyBooking_(sheet, d);
        known[d.bookingId] = d.receivedAt;
        if (result !== 'skip') {
          fixed++;
          Logger.log('補完: ' + d.bookingId + ' ' + d.tourDate + ' ' + d.courseName + ' [' + result + ']');
        }
      } catch (e) {
        Logger.log('補完エラー: ' + d.bookingId + ' → ' + e.message);
      }
    }
    if (fixed > 0) applyDateFilter(sheet);
    Logger.log('✅ 取りこぼしチェック完了（' + days + '日分）: ' + fixed + '件を補完');
  } finally {
    lock.releaseLock();
  }
}


// ── 共通：スレッド群を解析して受信順に並べる ──
function collectParsed_(threads, since) {
  const parsed = [];
  threads.forEach(function(thread, ti) {
    thread.getMessages().forEach(function(message) {
      try {
        const msgDate = message.getDate();
        if (since && msgDate < since) return;
        const data = parseEmail(message);
        if (!data || data.courseName === 'その他') return;
        parsed.push({ data: data, msgDate: msgDate, ti: ti });
      } catch (e) {
        Logger.log('解析エラー: ' + e.message);
      }
    });
  });
  parsed.sort(function(a, b) { return a.msgDate - b.msgDate; });
  return parsed;
}

// ── 共通：1件をシートに反映（'new' | 'modified' | 'canceled' | 'skip'）──
function applyBooking_(sheet, data) {
  if (data.isCanceled) return markAsCanceled(sheet, data.bookingId, data);
  return insertBookingInOrder(sheet, data);
}


// ============================================================
//  メール振り分け
// ============================================================
function parseEmail(message) {
  if (!message) return null;
  const from    = message.getFrom() || '';
  const subject = message.getSubject() || '';
  const raw     = message.getBody() || '';

  if (subject.indexOf('渋谷') !== -1 || raw.indexOf('渋谷') !== -1) return null;

  if (from.indexOf('resend.dev') !== -1) {
    return subject.indexOf('New booking') !== -1 ? parseENBookingEmail(message, subject, raw) : null;
  }
  if (subject.indexOf('Japan ticket') !== -1 || raw.indexOf('Japan ticket') !== -1) {
    return parseJapanTicketEmail(message, subject, raw);
  }
  if (raw.indexOf('messaging.squareup.com') !== -1 || subject.indexOf('様との新規予約') !== -1) {
    return parseSquareEmail(message, subject, raw);
  }
  if (from.indexOf('bokun.io') !== -1 || raw.indexOf('商品の予約参照番号') !== -1) {
    return parseBokunEmail(message, subject, raw);
  }
  if (from.indexOf('getyourguide.com') !== -1
    || raw.indexOf('getyourguide.com') !== -1
    || raw.indexOf('GetYourGuide') !== -1
    || (raw.indexOf('GYG') !== -1 && (raw.indexOf('参照番号') !== -1 || raw.indexOf('予約番号') !== -1))) {
    return parseGYGEmail(message, subject, raw);
  }
  return parseKlookEmail(message, subject, raw);
}

// HTML → 1行のプレーンテキスト
function cleanBody_(raw) {
  return String(raw || '')
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&#x27;|&#39;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/&[a-zA-Z0-9#]+;/g, ' ')
    .replace(/\s+/g, ' ');
}

function pad2_(n) { return String(parseInt(n, 10)).padStart(2, '0'); }


// ── Klook ──
function parseKlookEmail(message, subject, raw) {
  const body = cleanBody_(raw);

  const idMatch = body.match(/予約番号[：:\s]+([A-Z0-9]{6,12})/);
  if (!idMatch) return null;
  const bookingId = idMatch[1].trim();

  let isCanceled       = subject.includes('キャンセル') || body.includes('キャンセルされました');
  const isModification = subject.includes('予約内容に変更') || body.includes('予約内容を変更いたしました');

  // 一部キャンセル：「残りの予約: 3 x」があれば全キャンセルではなく人数変更
  const remainMatch = body.match(/残りの予約[：:\s]*(\d+)\s*[x×]/);
  const remaining   = remainMatch ? parseInt(remainMatch[1], 10) : null;
  let isPartialCancel = false;
  if (isCanceled && (subject.includes('一部キャンセル') || (remaining !== null && remaining > 0))) {
    isCanceled = false;
    isPartialCancel = true;
  }

  const dateMatch = body.match(/ご希望日[：:\s]+(\d{4}-\d{2}-\d{2})/);
  const tourDate  = dateMatch ? dateMatch[1].replace(/-/g, '/') : '';
  const timeMatch = body.match(/ご希望時間帯[：:\s]+(\d{1,2}:\d{2})/);   // 変更メールは最初＝新時間
  const tourTime  = timeMatch ? timeMatch[1] : '--:--';

  // 人数
  let participants = 0;
  if (isPartialCancel) {
    participants = remaining;
  } else {
    const paxSec = body.match(/利用者[：:\s]+([\s\S]{0,150})/);
    if (paxSec) {
      let sec = paxSec[1];
      const cut = sec.search(/予約代表者|オプション|合計|単価|ご希望|キャンセル|残り/);
      if (cut > 0) sec = sec.substring(0, cut);
      (sec.match(/(\d+)\s*[x×]/g) || []).forEach(function(m) { participants += parseInt(m, 10); });
    }
  }

  // 名前：次の項目名の手前で止める（本文が1行に圧縮されているため）
  const nameMatch = body.match(/予約代表者様?の氏名[：:\s]*(.+?)\s*(?:パスポート国|予約代表者様?の|利用者|ご希望|$)/);
  const guestName = nameMatch
    ? nameMatch[1].trim().replace(/^\((?:MR|MRS|MS|MISS|DR)\)\s*/i, '').substring(0, 60)
    : '';

  const countryMatch = body.match(/パスポート国[^A-Za-z]{0,15}([A-Za-z][A-Za-z .\-']{0,28}[A-Za-z])/);
  let country = countryMatch ? normalizeCountry(countryMatch[1].trim()) : '';

  const emailMatch = body.match(/予約代表者様のメールアドレス[：:\s]*([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})/);
  const email = emailMatch ? emailMatch[1].trim() : '';

  let phone = '';
  const phoneMatch = body.match(/予約代表者様?の(?:携帯|電話)番号[：:\s]*\+?\s*(\d{1,4})[\-\s]+([\d][\d\s\-\(\)]{4,18})/);
  if (phoneMatch) {
    phone = '+' + phoneMatch[1] + ' ' + phoneMatch[2].trim();
    if (!country) country = phoneToCountry(phone);
  }

  const courseName = detectCourse(body, subject);
  const receivedAt = Utilities.formatDate(message.getDate(), CONFIG.TZ, 'yyyy/MM/dd HH:mm');
  const isToday    = receivedAt.substring(0, 10) === tourDate;

  return {
    receivedAt: receivedAt, bookingId: bookingId, tourDate: tourDate, tourTime: tourTime,
    participants: participants, guestName: guestName, courseName: courseName,
    isCanceled: isCanceled, isModification: isModification || isPartialCancel, isToday: isToday,
    source: 'Klook', phone: phone, country: country, email: email,
    shop: detectShop(body, subject, courseName),
  };
}


// ── Bokun ──
function parseBokunEmail(message, subject, raw) {
  const body = cleanBody_(raw);

  const isCanceled     = subject.includes('キャンセル') || body.includes('キャンセルされましたのでお知らせします');
  const isModification = subject.includes('予約変更')   || body.includes('変更されましたのでお知らせします');

  const idMatch = body.match(/商品の予約参照番号\s+(KAK-[A-Z0-9]+)/i) || subject.match(/(KAK-[A-Z0-9]+)/i);
  if (!idMatch) return null;
  const bookingId = idMatch[1].trim();

  const dm = body.match(/日付\s+(?:月|火|水|木|金|土|日)\s+(\d{1,2})\.(\d{1,2})月\s+'?(\d{2})\s+@\s+(\d{1,2}:\d{2})/);
  let tourDate = '', tourTime = '--:--';
  if (dm) {
    tourDate = '20' + dm[3] + '/' + pad2_(dm[2]) + '/' + pad2_(dm[1]);
    tourTime = dm[4];
  }

  let participants = 0;
  const paxSec = body.match(/PAX\s+([\s\S]{0,120})/i);
  if (paxSec) {
    let sec = paxSec[1];
    const cut = sec.search(/ガイド言語|追加オプション|予約作成日|お客様|日付|備考/);
    if (cut > 0) sec = sec.substring(0, cut);
    (sec.match(/(\d+)\s*(?:x\s*)?(?:Adult|Child|Infant|Youth|Senior|大人|子供|こども|幼児)/gi) || [])
      .forEach(function(m) { participants += parseInt(m, 10); });
  }

  const nameMatch = body.match(/お客様\s+([^\s][^お]+?)(?:\s+お客様のメール|\s+お客様の電話|\s+日付)/);
  const guestName = nameMatch ? nameMatch[1].trim() : '';

  // 電話欄：「US+1 8059907117」のようにISO国コードが付く
  let phone = '', country = '';
  const pm = body.match(/電話番号[：:\s]*([A-Z]{2})?\s*(\+\d{1,3})[\s\-]*([\(\d][\d\s\-\(\)]{3,18})?/);
  if (pm) {
    phone = (pm[2] + ' ' + (pm[3] || '')).trim().replace(/\s+/g, ' ');
    country = (pm[1] ? iso2ToCountry(pm[1]) : '') || phoneToCountry(pm[2]);
  }

  const emailMatch = body.match(/メールアドレス[：:\s]*([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})/);
  const email = emailMatch ? emailMatch[1].trim() : '';

  const courseName = detectCourse(body, subject);
  const receivedAt = Utilities.formatDate(message.getDate(), CONFIG.TZ, 'yyyy/MM/dd HH:mm');

  return {
    receivedAt: receivedAt, bookingId: bookingId, tourDate: tourDate, tourTime: tourTime,
    participants: participants, guestName: guestName, courseName: courseName,
    isCanceled: isCanceled, isModification: isModification, isToday: receivedAt.substring(0, 10) === tourDate,
    source: 'Bokun', phone: phone, country: country, email: email,
    shop: detectShop(body, subject, courseName),
  };
}


// ── GetYourGuide ──
function parseGYGEmail(message, subject, raw) {
  const body = cleanBody_(raw);
  const subLower = subject.toLowerCase();

  const isCanceled = subLower.includes('canceled') || subLower.includes('cancelled')
    || subject.includes('キャンセル') || body.includes('キャンセルされました') || body.includes('キャンセル済み');
  const isModification = subject.includes('予約詳細の変更') || body.includes('変更がありましたのでお知らせ');

  const idMatch = subject.match(/(GYG[A-Z0-9]{6,})/i)
    || body.match(/(?:予約番号|参照番号)[：:\s]+(GYG[A-Z0-9]{6,})/i)
    || body.match(/Reference\s+[Nn]umber[:\s]+(GYG[A-Z0-9]{6,})/i);
  if (!idMatch) return null;
  const bookingId = idMatch[1].trim().toUpperCase();

  let tourDate = '', tourTime = '--:--', prevDateTime = '';
  if (isModification) {
    const m = body.match(/日付\s+新着\s+(\d{4})年(\d{1,2})月(\d{1,2})日\s+(\d{1,2}):(\d{2})\s+(\d{4})年(\d{1,2})月(\d{1,2})日\s+(\d{1,2}):(\d{2})/);
    if (m) {
      tourDate = m[1] + '/' + pad2_(m[2]) + '/' + pad2_(m[3]);
      tourTime = pad2_(m[4]) + ':' + m[5];
      prevDateTime = m[6] + '/' + pad2_(m[7]) + '/' + pad2_(m[8]) + ' ' + pad2_(m[9]) + ':' + m[10];
    }
  }
  if (!tourDate) {
    const jp = body.match(/(\d{4})年(\d{1,2})月(\d{1,2})日\s+(\d{1,2}):(\d{2})/);
    if (jp) {
      tourDate = jp[1] + '/' + pad2_(jp[2]) + '/' + pad2_(jp[3]);
      tourTime = pad2_(jp[4]) + ':' + jp[5];
    }
  }
  if (!tourDate) {
    const MONTHS = { January:1, February:2, March:3, April:4, May:5, June:6,
                     July:7, August:8, September:9, October:10, November:11, December:12 };
    const en = body.match(/(?:日付|Date)[:\s]+([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})\s+(\d{1,2}):(\d{2})\s*(AM|PM)/i);
    if (en) {
      let hour = parseInt(en[4], 10);
      const ampm = en[6].toUpperCase();
      if (ampm === 'PM' && hour < 12) hour += 12;
      if (ampm === 'AM' && hour === 12) hour = 0;
      tourDate = en[3] + '/' + pad2_(MONTHS[en[1]] || 1) + '/' + pad2_(en[2]);
      tourTime = pad2_(hour) + ':' + en[5];
    }
  }

  let participants = 0;
  const paxSec = body.match(/(?:参加者の人数|Number of participants)[:\s]*([\s\S]{0,300})/i);
  if (paxSec) {
    let sec = paxSec[1];
    const cut = sec.search(/予約者|Main customer|Booker|電話番号|よくある|オプション/i);
    if (cut > 0) sec = sec.substring(0, cut);
    const bd = sec.match(/(\d+)\s*[x×]/gi);
    if (bd) bd.forEach(function(m) { participants += parseInt(m, 10); });
    else {
      const plain = sec.match(/^\s*(\d+)/);
      if (plain) participants = parseInt(plain[1], 10);
    }
  }
  if (participants === 0) {
    const fb = body.match(/(\d+)\s*[x×]\s*Adults?/i);
    if (fb) participants = parseInt(fb[1], 10);
  }

  let guestName = '';
  const n1 = body.match(/(?:予約者|Booker|Main customer)[：:\s]*([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\s\-,\.]+?)(?:\s+customer-|\s+電話|\s+Phone:|\s+言語:|言語:)/i);
  const n2 = body.match(/名前[：:\s]+([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\s\-,\.]+?)(?:\s*(?:日付|ツアー|この顧客|サプライ|よく|電話))/i);
  const n3 = body.match(/\bName[：:\s]+([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\s\-,\.]+?)(?:\s+customer-|\s+Email|\s+Phone)/i);
  if (n1) guestName = n1[1].trim();
  else if (n2) guestName = n2[1].trim();
  else if (n3) guestName = n3[1].trim();

  let phone = '';
  const p = body.match(/電話番号[：:\s]*[\s\S]{0,5}?(\+\d{1,3}[\s\-]?\d{4,15})/)
    || body.match(/Phone[：:\s]*[\s\S]{0,40}?(\+\d{1,3}[\s\-]?\d{4,15})/i)
    || body.match(/(\+\d{1,3}[\d\s\-]{8,20})/);
  if (p) phone = p[1].trim();
  const country = phone ? phoneToCountry(phone.replace(/\s/g, '')) : '';

  const emailMatch = body.match(/([A-Za-z0-9._%+\-]+@reply\.getyourguide\.com)/);
  const email = emailMatch ? emailMatch[1].trim() : '';

  const courseName = detectCourse(body, subject);
  const receivedAt = Utilities.formatDate(message.getDate(), CONFIG.TZ, 'yyyy/MM/dd HH:mm');

  return {
    receivedAt: receivedAt, bookingId: bookingId, tourDate: tourDate, tourTime: tourTime,
    participants: participants, guestName: guestName, courseName: courseName,
    isCanceled: isCanceled, isModification: isModification, isToday: receivedAt.substring(0, 10) === tourDate,
    source: 'GetYourGuide', prevDateTime: prevDateTime, phone: phone, country: country, email: email,
    shop: detectShop(body, subject, courseName),
  };
}


// ── 自社サイト（EN Booking / resend.dev）──
function parseENBookingEmail(message, subject, raw) {
  const body = cleanBody_(raw);

  const dm = body.match(/Date\s+(\d{4})-(\d{2})-(\d{2})/);
  if (!dm) return null;
  const tourDate = dm[1] + '/' + dm[2] + '/' + dm[3];

  const tm = body.match(/Time\s+(\d{1,2}:\d{2})/);
  const tourTime = tm ? tm[1] : '--:--';

  const gm = body.match(/Guests\s+(\d+)\s*名/);
  const participants = gm ? parseInt(gm[1], 10) : 1;

  const nm = body.match(/Name\s+([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\s\-\.]+?)\s+Email/);
  const guestName = nm ? nm[1].trim() : '';

  const em = body.match(/Email\s+([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})/);
  const email = em ? em[1].trim() : '';

  const pm = body.match(/Phone\s+(\+\d[\d\s\-]{5,20})/);
  const phone = pm ? pm[1].trim() : '';

  const courseName = detectCourse(body, subject);
  const receivedAt = Utilities.formatDate(message.getDate(), CONFIG.TZ, 'yyyy/MM/dd HH:mm');

  return {
    receivedAt: receivedAt,
    bookingId: 'EN-' + tourDate.replace(/\//g, '') + tourTime.replace(':', '') + '-' + guestName.replace(/\s/g, '').substring(0, 10),
    tourDate: tourDate, tourTime: tourTime, participants: participants,
    guestName: guestName, courseName: courseName,
    isCanceled: subject.toLowerCase().includes('cancel') || subject.includes('キャンセル'),
    isModification: false, isToday: receivedAt.substring(0, 10) === tourDate,
    source: '自社サイト', phone: phone, country: phone ? phoneToCountry(phone.replace(/\s/g, '')) : '',
    email: email, shop: SHOP_KYOTO,
  };
}


// ── Square（自社サイト・転送）──
function parseSquareEmail(message, subject, raw) {
  const body = cleanBody_(raw);

  const subjName = subject.match(/Fwd:\s*(.+?)様との新規予約/);
  const bodyName = body.match(/新規予約があります。\s*([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\s\-,\.]+?)\s+[a-zA-Z0-9._%+\-]+@/);
  const guestName = bodyName ? bodyName[1].trim() : (subjName ? subjName[1].trim() : '');

  const dt = body.match(/(\d{4})年(\d{1,2})月(\d{1,2})日[^\d]*?(\d{1,2}):(\d{2})\s*-\s*\d{1,2}:\d{2}/);
  if (!dt) return null;
  const tourDate = dt[1] + '/' + pad2_(dt[2]) + '/' + pad2_(dt[3]);
  const tourTime = pad2_(dt[4]) + ':' + dt[5];

  const em = body.match(/([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})/);
  const pm = body.match(/(\+\d{1,3}[\s\d]{6,20})/);
  const phone = pm ? pm[1].trim() : '';

  const courseName = detectCourse(body, subject);
  const receivedAt = Utilities.formatDate(message.getDate(), CONFIG.TZ, 'yyyy/MM/dd HH:mm');

  return {
    receivedAt: receivedAt,
    bookingId: 'SQ-' + tourDate.replace(/\//g, '') + tourTime.replace(':', '') + '-' + guestName.replace(/\s/g, '').substring(0, 10),
    tourDate: tourDate, tourTime: tourTime,
    participants: 1,   // 人数情報なし
    guestName: guestName, courseName: courseName,
    isCanceled: subject.includes('キャンセル') || body.includes('キャンセルされました'),
    isModification: false, isToday: receivedAt.substring(0, 10) === tourDate,
    source: '自社サイト', phone: phone, country: phone ? phoneToCountry(phone.replace(/\s/g, '')) : '',
    email: em ? em[1].trim() : '', shop: SHOP_KYOTO,
  };
}


// ── コース・店舗判定 ──
function detectCourse(body, subject) {
  const t = body + subject;
  if (t.includes('金継') || t.includes('金繕'))                 return '金継ぎ';
  if (t.includes('包丁'))                                      return '包丁作り';
  if (t.includes('相撲'))                                      return '相撲ショー';
  if (t.includes('箸'))                                        return '箸作り';
  if (t.includes('リング') || t.includes('指輪'))              return 'リング作り';
  if (t.includes('Kintsugi'))                                  return '金継ぎ';
  if (/knife/i.test(t))                                        return '包丁作り';
  if (t.includes('Chopstick') || t.includes('Hashi'))          return '箸作り';
  if (/\bSumo\b/i.test(t))                                     return '相撲ショー';
  if (/\bRing Making\b|\bCoin Ring\b|\bRing Workshop\b/i.test(t)) return 'リング作り';
  return 'その他';
}

function detectShop(body, subject, courseName) {
  if (courseName === '相撲ショー') return SHOP_SUMO;
  const t = String(body || '') + String(subject || '');
  if (t.indexOf('大阪') !== -1 || /\bOsaka\b/i.test(t)) return SHOP_OSAKA;
  return SHOP_KYOTO;
}


// ============================================================
//  シート操作
// ============================================================
function ensureBookingHeader(sheet) {
  if (sheet.getLastRow() > 0) return;
  const h = ['体験日','開始時間','コース','人数','代表者名','予約番号','ステータス','受信日時','経路','変更前日時'];
  sheet.appendRow(h);
  sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#2b6cde').setFontColor('#ffffff').setFontSize(11);
  sheet.setFrozenRows(1);
  sheet.getRange('B:B').setNumberFormat('@');
}

// 戻り値: 'new' | 'modified' | 'skip'
function insertBookingInOrder(sheet, data) {
  const lastRow = sheet.getLastRow();

  // ── 既存行があれば更新 ──
  if (lastRow > 1) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    for (let i = 0; i < existing.length; i++) {
      const ex = existing[i];
      if (String(ex[5]).trim() !== data.bookingId) continue;

      const existingRecv = String(ex[7] || '');
      if (existingRecv && data.receivedAt && data.receivedAt < existingRecv) return 'skip';  // 古いメール

      // メールに無い項目は既存値を維持
      const tourDate = data.tourDate || String(ex[0]);
      const tourTime = (data.tourTime && data.tourTime !== '--:--') ? data.tourTime : String(ex[1]);
      const pax      = data.participants > 0 ? String(data.participants) : String(ex[3]);
      const changed  = tourDate !== String(ex[0]) || tourTime !== String(ex[1]) || pax !== String(ex[3]);
      if (!changed && !data.isModification) return 'skip';

      const row     = i + 2;
      const prevStr = data.prevDateTime || (ex[0] + ' ' + ex[1]);
      sheet.getRange(row, 1, 1, 10).setNumberFormat('@').setValues([[
        tourDate, tourTime, data.courseName || ex[2], pax,
        data.guestName || ex[4], data.bookingId, '🔄 予約変更', data.receivedAt,
        data.source || ex[8], prevStr,
      ]]);
      writeContactCols_(sheet, row, data);
      reapplyRowColors(sheet);
      Logger.log('変更で更新: ' + data.bookingId + ' ' + prevStr + ' → ' + tourDate + ' ' + tourTime + ' ' + pax + '名');
      return 'modified';
    }
  }

  // ── 新規挿入（体験日の降順）──
  const newKey = data.tourDate + ' ' + data.tourTime;
  let insertRow = lastRow + 1;
  if (lastRow > 1) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let j = 0; j < keys.length; j++) {
      if (newKey > (keys[j][0] + ' ' + keys[j][1])) { insertRow = j + 2; break; }
    }
  }
  if (insertRow <= lastRow) sheet.insertRowBefore(insertRow);

  sheet.getRange(insertRow, 1, 1, 10).setNumberFormat('@').setValues([[
    data.tourDate, data.tourTime, data.courseName, String(data.participants || 1),
    data.guestName, data.bookingId, data.isToday ? '★ 当日予約' : '予約済み',
    data.receivedAt, data.source || '', '',
  ]]);
  writeContactCols_(sheet, insertRow, data);
  reapplyRowColors(sheet);
  Logger.log('新規挿入: ' + data.bookingId + ' ' + newKey + ' ' + data.courseName + ' '
    + data.participants + '名 [' + data.source + '/' + (data.shop || '') + ']');
  return 'new';
}

// Q電話 R国 Sメール T店舗
function writeContactCols_(sheet, row, data) {
  if (data.phone)   sheet.getRange(row, 17).setNumberFormat('@').setValue(data.phone);
  if (data.country) sheet.getRange(row, 18).setNumberFormat('@').setValue(data.country);
  if (data.email)   sheet.getRange(row, 19).setNumberFormat('@').setValue(data.email);
  if (data.shop)    sheet.getRange(row, COL_SHOP).setNumberFormat('@').setValue(data.shop);
}

// 戻り値: 'canceled' | 'skip'
function markAsCanceled(sheet, bookingId, data) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][5]).trim() !== bookingId) continue;
      const existingRecv = String(rows[i][7] || '');
      if (existingRecv && data && data.receivedAt && data.receivedAt <= existingRecv) return 'skip';
      if (String(rows[i][6]).includes('キャンセル')) return 'skip';
      sheet.getRange(i + 2, 7).setValue('❌ キャンセル');
      if (data && data.receivedAt) sheet.getRange(i + 2, 8).setNumberFormat('@').setValue(data.receivedAt);
      sheet.getRange(i + 2, 1, 1, 10).setFontLine('line-through').setFontColor('#999999')
        .setFontWeight('normal').setBackground('#eeeeee');
      Logger.log('キャンセル: ' + bookingId);
      return 'canceled';
    }
  }
  if (!data) return 'skip';
  // シートに無い予約のキャンセル → 記録として追記
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, 10).setNumberFormat('@').setValues([[
    data.tourDate || '', data.tourTime || '--:--', data.courseName || '',
    String(data.participants || 0), data.guestName || '', bookingId,
    '❌ キャンセル', data.receivedAt || '', data.source || '', '',
  ]]).setFontLine('line-through').setFontColor('#999999').setFontWeight('normal').setBackground('#eeeeee');
  if (data.shop) sheet.getRange(newRow, COL_SHOP).setNumberFormat('@').setValue(data.shop);
  return 'canceled';
}

// 色付けを一括で行う（1行ずつより大幅に速い）
function reapplyRowColors(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const n = lastRow - 1;
  const statuses = sheet.getRange(2, 7, n, 1).getValues();
  const bg = [], fc = [], fw = [], fl = [];
  for (let i = 0; i < n; i++) {
    const s = String(statuses[i][0]);
    let b, c, w, l = 'none';
    if      (s.includes('キャンセル')) { b = '#eeeeee'; c = '#999999'; w = 'normal'; l = 'line-through'; }
    else if (s.includes('直接来店'))   { b = '#e0f7f2'; c = '#00796b'; w = 'bold'; }
    else if (s.includes('予約変更'))   { b = '#fff8f0'; c = '#e65c00'; w = 'bold'; }
    else if (s.includes('当日'))       { b = '#fff0f0'; c = '#cc0000'; w = 'bold'; }
    else                               { b = i % 2 === 0 ? '#f0f4ff' : '#ffffff'; c = '#000000'; w = 'normal'; }
    bg.push(new Array(10).fill(b));
    fc.push(new Array(10).fill(c));
    fw.push(new Array(10).fill(w));
    fl.push(new Array(10).fill(l));
  }
  sheet.getRange(2, 1, n, 10).setBackgrounds(bg).setFontColors(fc).setFontWeights(fw).setFontLines(fl).setFontSize(11);
}

// 前後 FILTER_DAYS 日以外の行を非表示（連続行はまとめて処理）
function applyDateFilter(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const now = Date.now();
  const min = new Date(now - CONFIG.FILTER_DAYS * 86400000);
  const max = new Date(now + CONFIG.FILTER_DAYS * 86400000);
  sheet.showRows(2, lastRow - 1);
  const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  let start = -1;
  for (let i = 0; i <= dates.length; i++) {
    let hide = false;
    if (i < dates.length) {
      const p = String(dates[i][0]).split('/');
      if (p.length === 3) {
        const d = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
        hide = d < min || d > max;
      }
    }
    if (hide && start === -1) start = i;
    if (!hide && start !== -1) { sheet.hideRows(start + 2, i - start); start = -1; }
  }
}


// ============================================================
//  毎朝7時：本日のタイムライン
// ============================================================
function generateDailyTimeline() {
  const ss     = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const bSheet = ss.getSheetByName(CONFIG.BOOKING_SHEET);
  if (bSheet) applyDateFilter(bSheet);

  const today      = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy/MM/dd');
  const todayLabel = Utilities.formatDate(new Date(), CONFIG.TZ, 'M月d日(EEE)');

  let tl = ss.getSheetByName(CONFIG.TIMELINE_SHEET);
  if (tl) tl.clear(); else tl = ss.insertSheet(CONFIG.TIMELINE_SHEET, 0);

  const rows = [];
  if (bSheet && bSheet.getLastRow() > 1) {
    bSheet.getDataRange().getValues().slice(1).forEach(function(r) {
      const status = String(r[6]);
      if (String(r[0]) === today && !status.includes('キャンセル')) {
        rows.push({ time: r[1], course: r[2], pax: r[3], name: r[4], id: r[5],
                    isToday: status.includes('当日'), source: r[8] || '' });
      }
    });
    rows.sort(function(a, b) { return String(a.time).localeCompare(String(b.time)); });
  }

  tl.getRange(1, 1).setValue('📅 ' + todayLabel + ' 予約タイムライン')
    .setFontSize(18).setFontWeight('bold').setFontColor('#1a237e');
  tl.getRange(2, 1).setValue('生成: ' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy/MM/dd HH:mm'))
    .setFontColor('#888888').setFontSize(10);

  if (rows.length === 0) {
    tl.getRange(4, 1).setValue('本日の予約はありません').setFontColor('#999999').setFontSize(14);
    return;
  }

  const hRow = 4;
  tl.getRange(hRow, 1, 1, 6).setValues([['開始時間','コース','人数','代表者名','予約番号','経路']])
    .setFontWeight('bold').setBackground('#2b6cde').setFontColor('#ffffff').setFontSize(12);
  tl.getRange('A:A').setNumberFormat('@');

  let totalPax = 0;
  rows.forEach(function(b, idx) {
    const rng = tl.getRange(hRow + 1 + idx, 1, 1, 6)
      .setValues([[b.time, b.course, b.pax, b.name, b.id, b.source]]).setFontSize(12);
    if (b.isToday) rng.setFontColor('#cc0000').setFontWeight('bold').setBackground('#fff0f0');
    else           rng.setBackground(idx % 2 === 0 ? '#f0f4ff' : '#ffffff');
    totalPax += Number(b.pax) || 0;
  });

  const sumRow = hRow + rows.length + 2;
  tl.getRange(sumRow, 1, 1, 6).setValues([['合計', '', totalPax + ' 名', '', '', '']])
    .setFontWeight('bold').setBackground('#e8f0fe').setFontSize(13);
  tl.getRange(sumRow + 2, 1).setValue('🔴 赤文字 = 当日受付の予約').setFontColor('#cc0000').setFontSize(10);
  [100, 180, 70, 200, 120, 100].forEach(function(w, i) { tl.setColumnWidth(i + 1, w); });
}


// ============================================================
//  1時間枠の人数超過アラート（京都・大阪のみ）
// ============================================================
function checkHourlyCapacity() {
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const rows  = sheet.getRange(2, 1, lastRow - 1, 20).getValues();
  const today = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy/MM/dd');
  const buckets = {};

  rows.forEach(function(r) {
    const date = String(r[0]).trim();
    const shop = String(r[19] || '').trim() || SHOP_KYOTO;
    if (String(r[6]).includes('キャンセル')) return;
    if (shop !== SHOP_KYOTO && shop !== SHOP_OSAKA) return;
    if (date < today) return;
    const hm = String(r[1]).trim().match(/^(\d{1,2}):/);
    if (!hm) return;
    const key = shop + '|' + date + '|' + hm[1].padStart(2, '0') + ':00';
    buckets[key] = (buckets[key] || 0) + (parseInt(r[3], 10) || 0);
  });

  const props    = PropertiesService.getScriptProperties();
  const notified = JSON.parse(props.getProperty('CAPACITY_NOTIFIED') || '{}');
  const over = [], stillOver = {};

  Object.keys(buckets).forEach(function(key) {
    if (buckets[key] <= CAPACITY_LIMIT) return;
    stillOver[key] = true;
    if (notified[key] === buckets[key]) return;   // 同じ人数で通知済み
    const p = key.split('|');
    over.push({ key: key, shop: p[0], date: p[1], hour: p[2], pax: buckets[key] });
  });
  Object.keys(notified).forEach(function(k) { if (!stillOver[k]) delete notified[k]; });

  if (over.length > 0) {
    over.sort(function(a, b) { return (a.date + a.hour).localeCompare(b.date + b.hour); });
    let body = '以下の時間枠で予約人数が上限（' + CAPACITY_LIMIT + '名）を超えています。\n\n';
    over.forEach(function(o) {
      body += '・' + o.shop + '店 ' + o.date + ' ' + o.hour + '台 → ' + o.pax + '名\n';
      notified[o.key] = o.pax;
    });
    ALERT_EMAILS.forEach(function(addr) {
      GmailApp.sendEmail(addr, '【予約警告】1時間枠の人数超過（' + over.length + '件）', body);
    });
    Logger.log('警告メール送信: ' + over.length + '件');
  }
  props.setProperty('CAPACITY_NOTIFIED', JSON.stringify(notified));
}


// ============================================================
//  分析タブ（毎朝7時）
// ============================================================
function updateAnalytics() {
  const ss  = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const src = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const lastRow = src.getLastRow();
  if (lastRow <= 1) return;
  const rows = src.getRange(2, 1, lastRow - 1, 18).getValues();

  const byCountry = {}, ltBuckets = {}, bySource = {};
  const ltAll = [0, 0, 0, 0, 0];
  let total = 0, canceled = 0, modified = 0;

  function parseD(s) {
    const m = String(s).match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
    return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
  }

  rows.forEach(function(r) {
    const status = String(r[6]);
    const source = String(r[8] || '').trim() || '不明';
    total++;
    if (!bySource[source]) bySource[source] = { total: 0, canceled: 0, modified: 0 };
    bySource[source].total++;
    if (status.includes('キャンセル')) { canceled++; bySource[source].canceled++; return; }
    if (status.includes('予約変更'))   { modified++; bySource[source].modified++; }

    const country = String(r[17] || '').trim() || '不明';
    if (!byCountry[country]) byCountry[country] = { pax: 0, cnt: 0, Klook: 0, GetYourGuide: 0, Bokun: 0, '直接来店': 0 };
    byCountry[country].pax += parseInt(r[3], 10) || 0;
    byCountry[country].cnt++;
    if (byCountry[country][source] !== undefined) byCountry[country][source]++;

    const tour = parseD(r[0]), recv = parseD(r[7]);
    if (tour && recv) {
      const days = Math.floor((tour - recv) / 86400000);
      const b = days <= 0 ? 0 : days <= 3 ? 1 : days <= 7 ? 2 : days <= 30 ? 3 : 4;
      const course = String(r[2] || 'その他');
      if (!ltBuckets[course]) ltBuckets[course] = [0, 0, 0, 0, 0];
      ltBuckets[course][b]++;
      ltAll[b]++;
    }
  });

  const out = [];
  out.push(['予約データ分析', '', '', '', '', '', '最終更新: ' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy/MM/dd HH:mm')]);
  out.push([]);
  out.push(['■ 国別の来店客割合（キャンセル除く・人数ベース）']);
  out.push(['国', '人数', '割合', '件数', 'Klook', 'GYG', 'Bokun', '直接来店']);
  let totalPax = 0;
  Object.keys(byCountry).forEach(function(c) { totalPax += byCountry[c].pax; });
  Object.keys(byCountry).sort(function(a, b) { return byCountry[b].pax - byCountry[a].pax; }).forEach(function(c) {
    const d = byCountry[c];
    const pct = totalPax > 0 ? Math.round(d.pax / totalPax * 1000) / 10 : 0;
    out.push([c, d.pax, pct + '%', d.cnt, d.Klook, d.GetYourGuide, d.Bokun, d['直接来店']]);
  });
  out.push([]);

  out.push(['■ 予約リードタイム分布（コース別・キャンセル除く）']);
  out.push(['コース', '当日', '1〜3日前', '4〜7日前', '8〜30日前', '31日以上前', '件数計']);
  function ltRow(label, arr) {
    const sum = arr.reduce(function(a, b) { return a + b; }, 0);
    if (sum === 0) return [label, '-', '-', '-', '-', '-', 0];
    function p(n) { return Math.round(n / sum * 1000) / 10 + '%'; }
    return [label, p(arr[0]), p(arr[1]), p(arr[2]), p(arr[3]), p(arr[4]), sum];
  }
  out.push(ltRow('全コース計', ltAll));
  Object.keys(ltBuckets).sort().forEach(function(c) { out.push(ltRow(c, ltBuckets[c])); });
  out.push([]);

  out.push(['■ キャンセル率・変更率']);
  out.push(['経路', '予約総数', 'キャンセル数', 'キャンセル率', '変更数', '変更率']);
  function rateRow(label, d) {
    const cr = d.total > 0 ? Math.round(d.canceled / d.total * 1000) / 10 + '%' : '-';
    const mr = d.total > 0 ? Math.round(d.modified / d.total * 1000) / 10 + '%' : '-';
    return [label, d.total, d.canceled, cr, d.modified, mr];
  }
  out.push(rateRow('全体', { total: total, canceled: canceled, modified: modified }));
  Object.keys(bySource).sort().forEach(function(s) { out.push(rateRow(s, bySource[s])); });

  let sheet = ss.getSheetByName('分析');
  if (!sheet) sheet = ss.insertSheet('分析');
  sheet.clear();
  const COLS = 8;
  const padded = out.map(function(r) { while (r.length < COLS) r.push(''); return r.slice(0, COLS); });
  sheet.getRange(1, 1, padded.length, COLS).setValues(padded);
  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  padded.forEach(function(r, i) {
    if (String(r[0]).indexOf('■') === 0) {
      sheet.getRange(i + 1, 1, 1, COLS).setBackground('#2b6cde').setFontColor('#ffffff').setFontWeight('bold');
    }
  });
  sheet.setColumnWidth(1, 180);
}


// ============================================================
//  ダッシュボード（dashboard.html から呼ばれる関数）
// ============================================================
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'dashboard';
  if (page === 'analytics') {
    try {
      return HtmlService.createHtmlOutputFromFile('analytics')
        .setTitle('予約分析').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (err) { /* analytics.html が無ければダッシュボードを表示 */ }
  }
  return HtmlService.createHtmlOutputFromFile('dashboard')
    .setTitle('予約管理').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function checkPassword_(password) {
  const correct = PropertiesService.getScriptProperties().getProperty('DASHBOARD_PASSWORD');
  if (!correct || password !== correct) throw new Error('WRONG_PASSWORD');
}

function checkDashboardPassword(password) {
  const correct = PropertiesService.getScriptProperties().getProperty('DASHBOARD_PASSWORD');
  return !!correct && password === correct;
}

// パスワード変更したい時だけ：NEW_PASSWORD に入力 → 実行 → すぐ空に戻して保存
function setDashboardPassword() {
  const NEW_PASSWORD = '';
  if (!NEW_PASSWORD) { Logger.log('NEW_PASSWORD が空です（現在のパスワードはそのまま）'); return; }
  PropertiesService.getScriptProperties().setProperty('DASHBOARD_PASSWORD', NEW_PASSWORD);
  Logger.log('パスワードを設定しました。コード内の NEW_PASSWORD を空に戻してください');
}

function getAnalyticsData() {
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, 20).getValues()
    .filter(function(r) { return String(r[5]).trim(); })
    .map(function(r) {
      return {
        date: String(r[0]).trim(), time: String(r[1]).trim(), course: String(r[2]).trim(),
        pax: parseInt(r[3], 10) || 0, status: dashStatus_(r[6]), recv: String(r[7]).trim(),
        source: String(r[8] || '').trim() || '不明', country: String(r[17] || '').trim() || '不明',
        shop: String(r[19] || '').trim() || SHOP_KYOTO,
      };
    });
}

function dashStatus_(status) {
  const s = String(status).trim();
  if (s.includes('キャンセル')) return 'canceled';
  if (s.includes('直接来店'))   return 'walkin';
  if (s.includes('予約変更'))   return 'modified';
  return 'booked';
}

function isTrue_(v) { return v === true || String(v || '').trim() === 'TRUE'; }

function getBookingsForDashboard(password) {
  if (!PropertiesService.getScriptProperties().getProperty('DASHBOARD_PASSWORD')) {
    throw new Error('パスワードが設定されていません。setDashboardPassword()を実行してください。');
  }
  checkPassword_(password);

  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  if (sheet.getLastRow() <= 1) return [];
  const today = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy/MM/dd');
  return sheet.getDataRange().getValues().slice(1)
  
    .filter(function(r) { return r[0] && r[5]; })
    .map(function(r) {
      const recv = String(r[7]).trim();
      return {
        date: String(r[0]).trim(), time: String(r[1]).trim(), course: String(r[2]).trim(),
        pax: parseInt(r[3], 10) || 0, name: String(r[4]).trim(), id: String(r[5]).trim(),
        status: dashStatus_(r[6]), sameDay: recv.substring(0, 10) === String(r[0]).trim(),
        receivedAt: recv, source: String(r[8] || '').trim(), prevDate: String(r[9] || '').trim(),
        paid: isTrue_(r[10]), staff: String(r[11] || '').trim(), reviewed: isTrue_(r[12]),
        memo: String(r[13] || '').trim(), arrived: isTrue_(r[14]), called: isTrue_(r[15]),
        phone: String(r[16] || '').trim(), email: String(r[18] || '').trim(),
                country: String(r[0]).trim() >= today
          ? (String(r[17] || '').trim() || phoneToCountry(String(r[16] || '')))
          : '',
        shop: String(r[19] || '').trim() || SHOP_KYOTO,
      };
    });
}

function findRowById_(sheet, bookingId) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;
  const ids = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(bookingId).trim()) return i + 2;
  }
  return -1;
}

function validateWalkin_(p, needShop) {
  if (!p) return 'データがありません';
  if (!p.staff) return '登録担当者は必須です';
  if (needShop && !p.shop) return '店舗は必須です';
  if (!p.date) return '体験日は必須です';
  if (!p.time) return '開始時間は必須です';
  if (!p.course) return 'コースは必須です';
  if (!p.pax || parseInt(p.pax, 10) < 1) return '人数は1以上で入力してください';
  return '';
}

// 直接来店の登録  payload: { date, time, course, pax, name, phone, staff, shop }
function createWalkinBooking(password, payload) {
  checkPassword_(password);
  const err = validateWalkin_(payload, true);
  if (err) return { ok: false, msg: err };

  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  ensureBookingHeader(sheet);

  const now = new Date();
  const bookingId  = 'WALK-' + Utilities.formatDate(now, CONFIG.TZ, 'yyyyMMddHHmmss');
  const receivedAt = Utilities.formatDate(now, CONFIG.TZ, 'yyyy/MM/dd HH:mm');

  const newKey  = payload.date + ' ' + payload.time;
  const lastRow = sheet.getLastRow();
  let insertRow = lastRow + 1;
  if (lastRow > 1) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let j = 0; j < keys.length; j++) {
      if (newKey > (keys[j][0] + ' ' + keys[j][1])) { insertRow = j + 2; break; }
    }
  }
  if (insertRow <= lastRow) sheet.insertRowBefore(insertRow);

  sheet.getRange(insertRow, 1, 1, 17).setNumberFormat('@').setValues([[
    payload.date, payload.time, payload.course, String(parseInt(payload.pax, 10)),
    String(payload.name || ''), bookingId, '🚶 直接来店', receivedAt, '直接来店', '',
    '', String(payload.staff), '', '', '', '', String(payload.phone || ''),
  ]]).setFontColor('#00796b').setFontWeight('bold').setBackground('#e0f7f2').setFontSize(11);
  sheet.getRange(insertRow, COL_SHOP).setNumberFormat('@').setValue(String(payload.shop));
  return { ok: true, id: bookingId };
}

function updateWalkinBooking(password, bookingId, payload) {
  checkPassword_(password);
  if (!String(bookingId).startsWith('WALK-')) return { ok: false, msg: 'この予約は予約サイト経由のため手動修正できません' };
  const err = validateWalkin_(payload, false);
  if (err) return { ok: false, msg: err };

  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const row = findRowById_(sheet, bookingId);
  if (row === -1) return { ok: false, msg: '予約が見つかりません: ' + bookingId };

  sheet.getRange(row, 1, 1, 5).setNumberFormat('@').setValues([[
    payload.date, payload.time, payload.course, String(parseInt(payload.pax, 10)), String(payload.name || ''),
  ]]);
  sheet.getRange(row, 12).setNumberFormat('@').setValue(String(payload.staff));
  sheet.getRange(row, 17).setNumberFormat('@').setValue(String(payload.phone || ''));
  if (payload.shop) sheet.getRange(row, COL_SHOP).setNumberFormat('@').setValue(String(payload.shop));
  return { ok: true };
}

function deleteWalkinBooking(password, bookingId) {
  checkPassword_(password);
  if (!String(bookingId).startsWith('WALK-')) return { ok: false, msg: 'この予約は予約サイト経由のため削除できません' };
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const row = findRowById_(sheet, bookingId);
  if (row === -1) return { ok: false, msg: '予約が見つかりません: ' + bookingId };
  sheet.deleteRow(row);
  return { ok: true };
}

function getStaffList(password) {
  checkPassword_(password);
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.STAFF_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(function(r) { return String(r[0]).trim(); })
    .filter(function(v) { return v; });
}

// field: 'paid' | 'staff' | 'reviewed' | 'memo' | 'arrived' | 'called'
function updateBookingDetail(password, bookingId, field, value) {
  checkPassword_(password);
  const colMap = { paid: 11, staff: 12, reviewed: 13, memo: 14, arrived: 15, called: 16 };
  const col = colMap[field];
  if (!col) return { ok: false, msg: '不明なフィールド: ' + field };

  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const row = findRowById_(sheet, bookingId);
  if (row === -1) return { ok: false, msg: '予約番号が見つかりません: ' + bookingId };

  const isBool = field === 'paid' || field === 'reviewed' || field === 'arrived' || field === 'called';
  sheet.getRange(row, col).setNumberFormat('@').setValue(isBool ? (value ? 'TRUE' : 'FALSE') : String(value || ''));
  return { ok: true };
}


// ============================================================
//  メンテナンス（手動実行用）
// ============================================================
function 診断() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  Logger.log('シート: ' + ss.getName() + ' / 行数: ' + ss.getSheetByName(CONFIG.BOOKING_SHEET).getLastRow());
  Logger.log('未処理メール（直近7日）: '
    + GmailApp.search(CONFIG.GMAIL_QUERY + ' -label:' + CONFIG.PROCESSED_LABEL + ' newer_than:7d', 0, 500).length);
  Logger.log('トリガー:');
  ScriptApp.getProjectTriggers().forEach(function(t) { Logger.log('  ' + t.getHandlerFunction()); });
}

// 同じ予約番号の重複行を削除（来店済み優先 → 受信日時が新しい方を残す）
function 重複を完全に削除() {
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const vals = sheet.getRange(2, 1, lastRow - 1, 20).getValues();
  const groups = {};
  vals.forEach(function(v, i) {
    const id = String(v[5]).trim();
    if (!id) return;
    (groups[id] = groups[id] || []).push({ row: i + 2, recv: String(v[7] || ''), arrived: isTrue_(v[14]) });
  });
  const del = [];
  Object.keys(groups).forEach(function(id) {
    const list = groups[id];
    if (list.length < 2) return;
    list.sort(function(a, b) {
      if (a.arrived !== b.arrived) return a.arrived ? -1 : 1;
      if (a.recv !== b.recv) return a.recv > b.recv ? -1 : 1;
      return b.row - a.row;
    });
    list.slice(1).forEach(function(x) { del.push(x.row); });
  });
  del.sort(function(a, b) { return b - a; }).forEach(function(r) { sheet.deleteRow(r); });
  Logger.log('重複 ' + del.length + '行を削除');
}


// ============================================================
//  ユーティリティ
// ============================================================
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getOrCreateLabel(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}


// ============================================================
//  国判定ヘルパー
// ============================================================
const PHONE_COUNTRY_MAP = {
  '+852':'香港','+853':'マカオ','+855':'カンボジア','+856':'ラオス','+880':'バングラデシュ','+886':'台湾',
  '+960':'モルディブ','+961':'レバノン','+962':'ヨルダン','+965':'クウェート','+966':'サウジアラビア',
  '+968':'オマーン','+971':'UAE','+972':'イスラエル','+973':'バーレーン','+974':'カタール',
  '+975':'ブータン','+976':'モンゴル','+977':'ネパール','+994':'アゼルバイジャン','+995':'ジョージア',
  '+351':'ポルトガル','+352':'ルクセンブルク','+353':'アイルランド','+358':'フィンランド','+359':'ブルガリア',
  '+370':'リトアニア','+371':'ラトビア','+372':'エストニア','+380':'ウクライナ','+385':'クロアチア',
  '+386':'スロベニア','+420':'チェコ','+421':'スロバキア',
  '+20':'エジプト','+27':'南アフリカ','+30':'ギリシャ','+31':'オランダ','+32':'ベルギー','+33':'フランス',
  '+34':'スペイン','+36':'ハンガリー','+39':'イタリア','+40':'ルーマニア','+41':'スイス','+43':'オーストリア',
  '+44':'イギリス','+45':'デンマーク','+46':'スウェーデン','+47':'ノルウェー','+48':'ポーランド','+49':'ドイツ',
  '+51':'ペルー','+52':'メキシコ','+54':'アルゼンチン','+55':'ブラジル','+56':'チリ',
  '+60':'マレーシア','+61':'オーストラリア','+62':'インドネシア','+63':'フィリピン','+64':'ニュージーランド',
  '+65':'シンガポール','+66':'タイ','+81':'日本','+82':'韓国','+84':'ベトナム','+86':'中国',
  '+90':'トルコ','+91':'インド','+92':'パキスタン',
  '+1':'アメリカ/カナダ','+7':'ロシア',
};
const PHONE_PREFIXES = Object.keys(PHONE_COUNTRY_MAP).sort(function(a, b) { return b.length - a.length; });

const CANADA_AREA_CODES = {
  '204':1,'226':1,'236':1,'249':1,'250':1,'263':1,'289':1,'306':1,'343':1,'354':1,
  '365':1,'367':1,'368':1,'382':1,'387':1,'403':1,'416':1,'428':1,'431':1,'437':1,
  '438':1,'450':1,'468':1,'474':1,'506':1,'514':1,'519':1,'548':1,'579':1,'581':1,
  '584':1,'587':1,'604':1,'613':1,'639':1,'647':1,'672':1,'683':1,'705':1,'709':1,
  '742':1,'753':1,'778':1,'780':1,'782':1,'807':1,'819':1,'825':1,'867':1,'873':1,
  '879':1,'902':1,'905':1,
};

function phoneToCountry(phone) {
  if (!phone) return '';
  const p = String(phone).replace(/[\s\-\(\)]/g, '');
  if (p.charAt(0) !== '+') return '';
  if (/^\+1\d{3}/.test(p)) return CANADA_AREA_CODES[p.substring(2, 5)] ? 'カナダ' : 'アメリカ';
  if (p.indexOf('+1') === 0) return 'アメリカ/カナダ';
  for (let i = 0; i < PHONE_PREFIXES.length; i++) {
    if (p.indexOf(PHONE_PREFIXES[i]) === 0) return PHONE_COUNTRY_MAP[PHONE_PREFIXES[i]];
  }
  return '';
}

const ISO2_COUNTRY_MAP = {
  'US':'アメリカ','CA':'カナダ','AU':'オーストラリア','GB':'イギリス','UK':'イギリス','NZ':'ニュージーランド',
  'DE':'ドイツ','FR':'フランス','IT':'イタリア','ES':'スペイン','NL':'オランダ','CH':'スイス','AT':'オーストリア',
  'BE':'ベルギー','SE':'スウェーデン','NO':'ノルウェー','DK':'デンマーク','FI':'フィンランド','IE':'アイルランド',
  'PT':'ポルトガル','PL':'ポーランド','CZ':'チェコ','GR':'ギリシャ','HU':'ハンガリー',
  'HK':'香港','TW':'台湾','CN':'中国','KR':'韓国','JP':'日本','SG':'シンガポール','MY':'マレーシア',
  'TH':'タイ','PH':'フィリピン','ID':'インドネシア','VN':'ベトナム','IN':'インド',
  'IL':'イスラエル','AE':'UAE','SA':'サウジアラビア','TR':'トルコ',
  'MX':'メキシコ','BR':'ブラジル','AR':'アルゼンチン','CL':'チリ','RU':'ロシア','ZA':'南アフリカ','EG':'エジプト',
};
function iso2ToCountry(code) {
  return ISO2_COUNTRY_MAP[String(code).toUpperCase()] || '';
}

const COUNTRY_JP_MAP = {
  'United States':'アメリカ','USA':'アメリカ','Canada':'カナダ','Australia':'オーストラリア',
  'United Kingdom':'イギリス','UK':'イギリス','Hong Kong':'香港','Taiwan':'台湾','China':'中国',
  'South Korea':'韓国','Korea':'韓国','Japan':'日本','Singapore':'シンガポール','Malaysia':'マレーシア',
  'Thailand':'タイ','Philippines':'フィリピン','Indonesia':'インドネシア','Vietnam':'ベトナム','India':'インド',
  'Israel':'イスラエル','Germany':'ドイツ','France':'フランス','Italy':'イタリア','Spain':'スペイン',
  'Netherlands':'オランダ','Switzerland':'スイス','Austria':'オーストリア','Belgium':'ベルギー',
  'Sweden':'スウェーデン','Norway':'ノルウェー','Denmark':'デンマーク','Finland':'フィンランド',
  'Ireland':'アイルランド','Portugal':'ポルトガル','Poland':'ポーランド','New Zealand':'ニュージーランド',
  'Mexico':'メキシコ','Brazil':'ブラジル','United Arab Emirates':'UAE','Saudi Arabia':'サウジアラビア',
  'Macau':'マカオ','Macao':'マカオ','Russia':'ロシア','Turkey':'トルコ',
};
function normalizeCountry(raw) {
  if (!raw) return '';
  const t = String(raw).trim();
  return COUNTRY_JP_MAP[t] || t;
}
function 予約番号を調べる() {
  const ids = ['GYM137800'];
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const sheetIds = sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues()
    .map(function(r) { return String(r[0]).trim(); });
  const label = GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL);

  ids.forEach(function(id) {
    Logger.log('════ ' + id + ' ════');
    Logger.log('シートにある？ ' + (sheetIds.indexOf(id) !== -1 ? 'はい' : 'いいえ'));

    const threads = GmailApp.search('"' + id + '"', 0, 5);
    if (threads.length === 0) { Logger.log('❌ Gmailにメールが見つからない'); return; }

    threads.forEach(function(th) {
      const hasLabel = label ? th.getLabels().some(function(l) { return l.getName() === label.getName(); }) : false;
      th.getMessages().forEach(function(msg) {
        Logger.log('From: ' + msg.getFrom());
        Logger.log('件名: ' + msg.getSubject());
        Logger.log('処理済みラベル: ' + hasLabel);
        const d = parseEmail(msg);
        if (!d) { Logger.log('→ 解析失敗（null）'); return; }
        Logger.log('→ 解析: ' + d.bookingId + ' / ' + d.tourDate + ' ' + d.tourTime
          + ' / コース: ' + d.courseName + ' / ' + d.participants + '名 / キャンセル: ' + d.isCanceled);
      });
    });
  });
}
function シートの行を確認() {
  const ids = ['GYGN6BWRNY7W', 'GYGFWVY87NXV'];
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 20).getValues();
  rows.forEach(function(r, i) {
    const id = String(r[5]).trim();
    if (ids.indexOf(id) === -1) return;
    Logger.log('行' + (i + 2) + ' ｜ ' + id
      + ' ｜ 日付[' + r[0] + '] 型:' + (r[0] instanceof Date ? 'Date' : typeof r[0])
      + ' ｜ 時間[' + r[1] + '] ｜ ' + r[2] + ' ｜ ' + r[3] + '名'
      + ' ｜ ステータス[' + r[6] + '] ｜ 店舗[' + r[19] + ']'
      + ' ｜ 非表示:' + sheet.isRowHiddenByUser(i + 2));
  });
}
function シートを並び替えて壊れた行を探す() {
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return;

  // ① 日付の形がおかしい行を一覧（削除はしない）
  const vals = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  let bad = 0;
  vals.forEach(function(r, i) {
    const a = r[0] instanceof Date ? 'Date型' : String(r[0]);
    if (!/^\d{4}\/\d{2}\/\d{2}$/.test(a)) {
      bad++;
      Logger.log('⚠️ 行' + (i + 2) + ' ｜ A列[' + a + '] ｜ B列[' + r[1] + '] ｜ F列[' + r[5] + ']');
    }
  });
  Logger.log('日付が不正な行: ' + bad + '件');

  // ② 全列まとめて並び替え（体験日・時間の新しい順）
  sheet.showRows(2, lastRow - 1);
  sheet.getRange(2, 1, lastRow - 1, lastCol).sort([
    { column: 1, ascending: false },
    { column: 2, ascending: false },
  ]);

  reapplyRowColors(sheet);
  applyDateFilter(sheet);
  Logger.log('✅ 並び替え完了');
}
function 壊れた行を修復する() {
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const lastRow = sheet.getLastRow();
  const vals = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  // 正常な行にある予約番号
  const goodIds = {};
  const badRows = [];
  vals.forEach(function(r, i) {
    const a = r[0] instanceof Date ? '' : String(r[0]);
    if (/^\d{4}\/\d{2}\/\d{2}$/.test(a)) {
      const id = String(r[5]).trim();
      if (id) goodIds[id] = true;
    } else {
      // 予約番号の位置：旧手動関数はB列、それ以外はF列
      const b = String(r[1]).trim(), f = String(r[5]).trim();
      const id = /^[A-Z0-9\-]{6,}$/.test(b) && !/:/.test(b) ? b : f;
      badRows.push({ row: i + 2, id: id });
    }
  });
  Logger.log('壊れた行: ' + badRows.length + '件');

  // ① 各行の正しいデータをGmailから取り直す（まだ削除しない）
  const toDelete = [], toInsert = [], keep = [];
  badRows.forEach(function(b) {
    if (!b.id) { keep.push(b); return; }
    if (goodIds[b.id]) { toDelete.push(b.row); return; }   // 正しい行が既にある → 削除だけ
    const threads = GmailApp.search('"' + b.id + '"', 0, 5);
    let latest = null;
    threads.forEach(function(th) {
      th.getMessages().forEach(function(m) {
        const d = parseEmail(m);
        if (d && d.bookingId === b.id && d.courseName !== 'その他'
            && (!latest || d.receivedAt > latest.receivedAt)) latest = d;
      });
    });
    if (latest) { toInsert.push(latest); toDelete.push(b.row); goodIds[b.id] = true; }
    else keep.push(b);
  });

  // ② 壊れた行を下から削除
  toDelete.sort(function(a, b) { return b - a; }).forEach(function(r) { sheet.deleteRow(r); });

  // ③ 正しい形で入れ直す
  toInsert.forEach(function(d) {
    if (d.isCanceled) markAsCanceled(sheet, d.bookingId, d);
    else insertBookingInOrder(sheet, d);
    Logger.log('✅ 入れ直し: ' + d.bookingId + ' ' + d.tourDate + ' ' + d.tourTime + ' ' + d.courseName + ' ' + d.participants + '名');
  });

  keep.forEach(function(b) { Logger.log('⚠️ 自動修復できず（残しました）: 行' + b.row + ' ID[' + b.id + ']'); });
  Logger.log('削除 ' + toDelete.length + ' ／ 入れ直し ' + toInsert.length + ' ／ 残し ' + keep.length);
}
function 残った壊れた行を調べる() {
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const lastRow = sheet.getLastRow();
  const vals = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  vals.forEach(function(r, i) {
    const a = r[0] instanceof Date ? 'Date型' : String(r[0]);
    if (/^\d{4}\/\d{2}\/\d{2}$/.test(a)) return;

    Logger.log('──── 行' + (i + 2) + ' ────');
    Logger.log('A〜J: ' + r.map(function(v) {
      return v instanceof Date ? 'Date型' : String(v).substring(0, 25);
    }).join(' ｜ '));

    const b = String(r[1]).trim(), f = String(r[5]).trim();
    const id = /^[A-Z0-9\-]{6,}$/.test(b) && !/:/.test(b) ? b : f;
    if (!id) { Logger.log('→ ID なし'); return; }

    const threads = GmailApp.search('"' + id + '"', 0, 3);
    Logger.log('→ Gmail検索: ' + threads.length + 'スレッド');
    threads.forEach(function(th) {
      th.getMessages().forEach(function(m) {
        const d = parseEmail(m);
        Logger.log('   件名[' + m.getSubject().substring(0, 40) + '] → '
          + (d ? 'ID:' + d.bookingId + ' 日付:' + d.tourDate + ' コース:' + d.courseName + ' キャンセル:' + d.isCanceled
               : '解析失敗(null)'));
      });
    });
  });
}
function 残りの壊れた行を直す() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const lastRow = sheet.getLastRow();
  const vals = sheet.getRange(2, 1, lastRow - 1, 20).getValues();
  const fmt = function(v, p) { return v instanceof Date ? Utilities.formatDate(v, tz, p) : String(v || '').trim(); };
  const isId = function(s) { return /^[A-Z0-9\-]{6,}$/.test(s) && s.indexOf(':') === -1; };
  const isGood = function(a) { return !(a instanceof Date) && /^\d{4}\/\d{2}\/\d{2}$/.test(String(a)); };

  const goodIds = {};
  vals.forEach(function(r) { if (isGood(r[0]) && String(r[5]).trim()) goodIds[String(r[5]).trim()] = true; });

  const del = [];
  let fixed = 0;
  const monthAgo = Utilities.formatDate(new Date(Date.now() - 30 * 86400000), tz, 'yyyy/MM/dd');

  vals.forEach(function(r, i) {
    if (isGood(r[0])) return;
    const row = i + 2;
    const b = String(r[1]).trim(), f = String(r[5]).trim();

    // 旧手動レイアウト（B列に予約番号）
    if (r[0] instanceof Date && isId(b)) {
      if (goodIds[b]) { del.push(row); return; }
      let date = '', time = '--:--';
      if (r[3] instanceof Date) { date = fmt(r[3], 'yyyy/MM/dd'); time = fmt(r[3], 'HH:mm'); }
      else {
        const m = String(r[3]).match(/(\d{4})[-\/](\d{2})[-\/](\d{2})\s+(\d{1,2}:\d{2})/);
        if (m) { date = m[1] + '/' + m[2] + '/' + m[3]; time = m[4]; }
      }
      if (!date) { Logger.log('⚠️ 日付が読めず残し: 行' + row + ' ' + b); return; }
      const course = detectCourse(String(r[7]), '');
      sheet.getRange(row, 1, 1, 10).setNumberFormat('@').setValues([[
        date, time, course, String(parseInt(r[6], 10) || 1),
        String(r[4]).replace(/^\((?:MR|MRS|MS|MISS|DR)\)\s*/i, ''), b, '予約済み',
        fmt(r[0], 'yyyy/MM/dd HH:mm'), String(r[2] || 'Klook'), '',
      ]]);
      const pm = String(r[5]).trim().match(/^(\d{1,4})-(\d+)$/);
      if (pm) sheet.getRange(row, 17).setNumberFormat('@').setValue('+' + pm[1] + ' ' + pm[2]);
      if (String(r[8]).trim()) sheet.getRange(row, 18).setNumberFormat('@').setValue(normalizeCountry(String(r[8]).trim()));
      sheet.getRange(row, 20).setNumberFormat('@').setValue(detectShop('', '', course));
      goodIds[b] = true; fixed++;
      Logger.log('✅ 旧レイアウト修正: ' + b + ' ' + date + ' ' + time + ' ' + course);
      return;
    }

    // ID なし → ゴミ行
    if (!isId(f)) { del.push(row); return; }

    // 既に正しい行がある → 重複
    if (goodIds[f]) { del.push(row); return; }

    // 日付・時間がDate型になっただけ → 文字に戻す
    if (r[0] instanceof Date) {
      sheet.getRange(row, 1, 1, 2).setNumberFormat('@')
        .setValues([[fmt(r[0], 'yyyy/MM/dd'), r[1] instanceof Date ? fmt(r[1], 'HH:mm') : String(r[1])]]);
      if (r[7] instanceof Date) sheet.getRange(row, 8).setNumberFormat('@').setValue(fmt(r[7], 'yyyy/MM/dd HH:mm'));
      goodIds[f] = true; fixed++;
      Logger.log('✅ 型を修正: ' + f + ' ' + fmt(r[0], 'yyyy/MM/dd'));
      return;
    }

    // 日付なし：1か月以上前の受信なら削除、新しいものは残す
    if (fmt(r[7], 'yyyy/MM/dd HH:mm').substring(0, 10) < monthAgo) del.push(row);
    else Logger.log('⚠️ 日付なし・最近の予約のため残し: 行' + row + ' ' + f);
  });

  del.sort(function(a, b) { return b - a; }).forEach(function(r) { sheet.deleteRow(r); });

  // 今後Date型に化けないよう A・B・H列を文字列に固定
  sheet.getRange('A:B').setNumberFormat('@');
  sheet.getRange('H:H').setNumberFormat('@');

  // 並び替え・色
  const lr = sheet.getLastRow();
  sheet.showRows(2, lr - 1);
  sheet.getRange(2, 1, lr - 1, sheet.getLastColumn()).sort([{ column: 1, ascending: false }, { column: 2, ascending: false }]);
  reapplyRowColors(sheet);
  applyDateFilter(sheet);
  Logger.log('修正 ' + fixed + ' ／ 削除 ' + del.length);

  // ── Klook「承認されました」メールが読めない原因調査 ──
  const th = GmailApp.search('"FRT865012"', 0, 1);
  if (th.length) {
    const body = cleanBody_(th[0].getMessages()[0].getBody());
    const k = body.indexOf('予約');
    Logger.log('【Klook本文】' + body.substring(Math.max(0, k - 50), k + 400));
  }
}
function 今日以降の電話番号を補完する() {
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
  const today = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy/MM/dd');
  const t0 = Date.now();
  let filled = 0, notFound = 0;

  for (let i = 0; i < rows.length; i++) {
    if (Date.now() - t0 > 5 * 60 * 1000) { Logger.log('時間切れ。もう一度実行してください'); break; }
    const r = rows[i];
    const id = String(r[5]).trim();
    if (String(r[0]) < today) continue;                    // 過去は対象外
    if (!id || id.indexOf('WALK-') === 0) continue;        // 直接来店は対象外
    if (String(r[6]).includes('キャンセル')) continue;
    if (String(r[16]).trim() && String(r[18]).trim()) continue;  // 電話・メール両方あり

    const threads = GmailApp.search('"' + id + '"', 0, 3);
    let best = null;
    threads.forEach(function(th) {
      th.getMessages().forEach(function(m) {
        const d = parseEmail(m);
        if (d && d.bookingId === id && (d.phone || d.email) && (!best || d.receivedAt > best.receivedAt)) best = d;
        // 取れなかった時の原因調査用
        if (d && d.bookingId === id && !d.phone && d.source === 'Klook') {
          const body = cleanBody_(m.getBody());
          const k = body.indexOf('携帯');
          Logger.log('⚠️ ' + id + ' 電話欄: ' + (k >= 0 ? body.substring(k - 10, k + 60) : '「携帯」の記載なし'));
        }
      });
    });

    if (!best) { notFound++; continue; }
    const row = i + 2;
    if (best.phone && !String(r[16]).trim())   sheet.getRange(row, 17).setNumberFormat('@').setValue(best.phone);
    if (best.country && !String(r[17]).trim()) sheet.getRange(row, 18).setNumberFormat('@').setValue(best.country);
    if (best.email && !String(r[18]).trim())   sheet.getRange(row, 19).setNumberFormat('@').setValue(best.email);
    filled++;
    Logger.log('✅ ' + id + ' 電話:' + (best.phone || '-') + ' 国:' + (best.country || '-'));
  }
  Logger.log('補完 ' + filled + '件 ／ 取得できず ' + notFound + '件');
}
function Klookの誤キャンセルを直す() {
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const today = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy/MM/dd');

  // キャンセル扱いになっている、今日以降のKlook予約
  const ids = [];
  rows.forEach(function(r) {
    if (String(r[8]).trim() !== 'Klook') return;
    if (!String(r[6]).includes('キャンセル')) return;
    if (String(r[0]) < today) return;
    ids.push(String(r[5]).trim());
  });
  Logger.log('対象: ' + ids.length + '件');

  ids.forEach(function(id) {
    const msgs = [];
    GmailApp.search('"' + id + '"', 0, 10).forEach(function(th) {
      th.getMessages().forEach(function(m) { msgs.push(m); });
    });
    msgs.sort(function(a, b) { return a.getDate() - b.getDate(); });
    const results = [];
    msgs.forEach(function(m) {
      const d = parseEmail(m);
      if (!d || d.bookingId !== id) return;
      results.push(applyBooking_(sheet, d));
    });
    const row = findRowById_(sheet, id);
    const now = row === -1 ? 'なし' : sheet.getRange(row, 1, 1, 7).getDisplayValues()[0];
    Logger.log(id + ' → ' + results.join(',') + ' ｜ 現在: ' + now[0] + ' ' + now[1] + ' ' + now[3] + '名 ' + now[6]);
  });
}
function 予約番号を再処理する() {
  const ids = ['UCV002133'];
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);

  ids.forEach(function(id) {
    const before = findRowById_(sheet, id);
    Logger.log('════ ' + id + ' ════');
    Logger.log('処理前: ' + (before === -1 ? 'シートになし'
      : '行' + before + ' → ' + sheet.getRange(before, 1, 1, 8).getDisplayValues()[0].join(' | ').substring(0, 150)));

    const msgs = [];
    GmailApp.search('"' + id + '"', 0, 10).forEach(function(th) {
      th.getMessages().forEach(function(m) { msgs.push(m); });
    });
    msgs.sort(function(a, b) { return a.getDate() - b.getDate(); });

    msgs.forEach(function(m) {
      const d = parseEmail(m);
      const subj = m.getSubject().substring(0, 30);
      if (!d || d.bookingId !== id) { Logger.log('件名[' + subj + '] → 解析対象外'); return; }
      const res = applyBooking_(sheet, d);
      Logger.log('件名[' + subj + '] → ' + d.tourDate + ' ' + d.tourTime + ' ' + d.participants
        + '名 キャンセル:' + d.isCanceled + ' 受信:' + d.receivedAt + ' → 結果: ' + res);
    });

    const after = findRowById_(sheet, id);
    Logger.log('処理後: ' + (after === -1 ? 'シートになし'
      : '行' + after + ' → ' + sheet.getRange(after, 1, 1, 8).getDisplayValues()[0].join(' | ').substring(0, 150)));
  });
}
function 人数0の予約を直す() {
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const today = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy/MM/dd');
  let fixed = 0;
  const manual = [];

  rows.forEach(function(r, i) {
    if (String(r[0]) < today) return;
    if (String(r[6]).includes('キャンセル')) return;
    if (parseInt(r[3], 10) > 0) return;
    const id = String(r[5]).trim();
    if (!id || id.indexOf('WALK-') === 0) return;

    const msgs = [];
    GmailApp.search('"' + id + '"', 0, 10).forEach(function(th) {
      th.getMessages().forEach(function(m) { msgs.push(m); });
    });
    msgs.sort(function(a, b) { return a.getDate() - b.getDate(); });
    let pax = 0;
    msgs.forEach(function(m) {
      const d = parseEmail(m);
      if (!d || d.bookingId !== id || d.isCanceled) return;
      if (d.participants > 0) pax = d.participants;
    });

    if (pax > 0) {
      sheet.getRange(i + 2, 4).setNumberFormat('@').setValue(String(pax));
      fixed++;
      Logger.log('✅ ' + id + ' ' + r[0] + ' → ' + pax + '名');
    } else {
      manual.push(id + '（' + r[0] + '）');
    }
  });

  Logger.log('修正 ' + fixed + '件');
  if (manual.length) Logger.log('⚠️ 手動確認が必要: ' + manual.join('、'));
}
// ── Japan ticket（KKday などの経由予約）──
function parseJapanTicketEmail(message, subject, raw) {
  const body = cleanBody_(raw);

  const dm = body.match(/予約日[：:\s]*(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!dm) return null;
  const tourDate = dm[1] + '/' + pad2_(dm[2]) + '/' + pad2_(dm[3]);
  const tm = body.match(/予約時間[：:\s]*(\d{1,2}):(\d{2})/);
  const tourTime = tm ? pad2_(tm[1]) + ':' + tm[2] : '--:--';

  // 人数：「(1人)」を合計。無ければチケット枚数
  let participants = 0;
  (body.match(/\((\d+)人\)/g) || []).forEach(function(m) { participants += parseInt(m.replace(/\D/g, ''), 10); });
  if (!participants) {
    const t = body.match(/チケット購入枚数[：:\s]*(\d+)/);
    participants = t ? parseInt(t[1], 10) : 1;
  }

  const ch   = body.match(/流入経路[：:\s]*(.+?)\s+-\s/);
  const plan = body.match(/プラン名[：:\s]*(.+?)\s+-\s/);
  const nm   = body.match(/(?:予約者|代表者|お名前)[^：:]{0,6}[：:\s]*(.+?)\s+-\s/);
  const channel = ch ? ch[1].trim() : '';
  const CH = { 'KLOOK': 'Klook', 'KKDAY': 'KKday' };
  const chName = channel ? (CH[channel.toUpperCase()] || channel) : '';

  // 予約番号：「予約詳細」リンクに番号があれば使う。無ければ受信時刻で作る
  const link = raw.match(/\/(?:reservations?|bookings?|orders?)\/([A-Za-z0-9\-]{4,})/i);
  const bookingId = 'JT-' + (link ? link[1]
    : Utilities.formatDate(message.getDate(), CONFIG.TZ, 'yyMMddHHmmss'));

  const courseName = detectCourse(body + (plan ? plan[1] : ''), subject);
  const receivedAt = Utilities.formatDate(message.getDate(), CONFIG.TZ, 'yyyy/MM/dd HH:mm');

  return {
    receivedAt: receivedAt, bookingId: bookingId, tourDate: tourDate, tourTime: tourTime,
    participants: participants, guestName: nm ? nm[1].trim() : '', courseName: courseName,
    isCanceled: subject.includes('キャンセル') || body.includes('キャンセルを受け付け'),
    isModification: subject.includes('変更'), isToday: receivedAt.substring(0, 10) === tourDate,
    source: chName ? chName + '（Japan ticket）' : 'Japan ticket', phone: '', country: '', email: '',
    shop: detectShop(body, subject, courseName),
  }
}

function JapanTicketをテスト() {
  GmailApp.search('subject:"Japan ticket"', 0, 3).forEach(function(th) {
    const m = th.getMessages()[0];
    const d = parseEmail(m);
    Logger.log(m.getSubject());
    Logger.log(d ? d.bookingId + ' / ' + d.tourDate + ' ' + d.tourTime + ' / ' + d.courseName
      + ' / ' + d.participants + '名 / 経路:' + d.source + ' / 店舗:' + d.shop : '→ 解析失敗');
  });
}
function JTの経路を直す() {
  const sheet = getOrCreateSheet(CONFIG.BOOKING_SHEET);
  const rows = sheet.getRange(2, 6, sheet.getLastRow() - 1, 4).getValues();  // F〜I列
  const CH = { 'KLOOK': 'Klook', 'KKDAY': 'KKday' };
  let n = 0;
  rows.forEach(function(r, i) {
    if (String(r[0]).indexOf('JT-') !== 0) return;
    const src = String(r[3]).trim();
    if (src.indexOf('Japan ticket') !== -1) return;
    const name = CH[src.toUpperCase()] || src;
    sheet.getRange(i + 2, 9).setValue(name ? name + '（Japan ticket）' : 'Japan ticket');
    n++;
  });
  Logger.log(n + '件を更新');
}
