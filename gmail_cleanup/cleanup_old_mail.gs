/**
 * 古いメールの一括ゴミ箱移動（管理者実行用）
 *
 * 対象  : TARGET_USER のメールボックス
 * 条件  : 2023年12月31日以前（= 2024/01/01 00:00 JST より前）に受信・送信したメール
 *         ただし「重要」マーク付きは除外
 * 処理  : ゴミ箱へ移動（30日以内なら Gmail のゴミ箱から復元可能）
 *
 * 仕組み: ドメイン全体の委任を設定したサービスアカウントで TARGET_USER になりすまし、
 *         Gmail API を呼び出す。GAS の 6 分制限に掛からないよう、約 5 分で処理を区切り、
 *         時間主導トリガーで自動的に続きを実行する。
 *
 * ── 事前準備（管理者） ─────────────────────────────────────────
 * 1. Google Cloud コンソールでプロジェクトを作成（既存でも可）し、「Gmail API」を有効化。
 * 2. サービスアカウントを作成し、JSON キーをダウンロード。
 *    ※ 組織ポリシーでキー作成が禁止されている場合は一時的に許可が必要。
 * 3. 管理コンソール > セキュリティ > アクセスとデータ管理 > API の制御
 *    > ドメイン全体の委任 > 「新しく追加」
 *      クライアント ID : サービスアカウントの「一意の ID（数字）」
 *      OAuth スコープ  : https://www.googleapis.com/auth/gmail.modify
 * 4. 新しい Apps Script プロジェクトにこのファイルと appsscript.json を貼り付け。
 * 5. プロジェクトの設定 > スクリプト プロパティ に以下を追加
 *      SA_KEY_JSON : ダウンロードした JSON キーの中身をそのまま貼り付け
 *    ※ キーはコードやリポジトリに書かないこと。作業完了後はキーと委任設定を削除推奨。
 *
 * ── 実行手順 ─────────────────────────────────────────────────
 * 1. dryRun()  … 対象件数の目安と先頭 20 件の日付・件名をログに出す（削除しない）。
 * 2. 内容を確認したら start() を実行。以降は自動でトリガーが続きを処理する。
 * 3. 進捗は showProgress()、途中で止めたいときは stop()。
 *    完了するとトリガーは自動で削除され、ログに「完了」と出る。
 */

const TARGET_USER = 'nakashima@nlg.co.jp';

// 2024-01-01 00:00:00 JST（= 2023-12-31 15:00:00 UTC）の UNIX 秒。
// Gmail の before:YYYY/MM/DD は太平洋時間で解釈されるため、秒指定で JST の境界を厳密にする。
const CUTOFF_EPOCH_SEC = 1704034800;

// 対象: カットオフより前 かつ 重要マークなし（ゴミ箱・迷惑メールは Gmail 側で自動的に除外）
const QUERY = 'before:' + CUTOFF_EPOCH_SEC + ' -is:important';

const SCOPE = 'https://www.googleapis.com/auth/gmail.modify';
const API_BASE = 'https://gmail.googleapis.com/gmail/v1/users/' + encodeURIComponent(TARGET_USER);
const TIME_BUDGET_MS = 5 * 60 * 1000;   // 1 回の実行で使う時間（6 分制限に余裕を残す）
const PARALLEL = 40;                    // 同時に送る trash リクエスト数（5 単位×40 = 200 単位/秒 < 上限 250）
const HANDLER = 'continueCleanup';
const PROP_TOTAL = 'TRASHED_TOTAL';

/** 削除せずに対象を確認する */
function dryRun() {
  const res = gmailGet_('/messages', { q: QUERY, maxResults: 20 });
  const ids = (res.messages || []).map(function (m) { return m.id; });
  console.log('検索条件: ' + QUERY);
  console.log('対象件数の目安（resultSizeEstimate、概算）: ' + (res.resultSizeEstimate || 0));
  ids.forEach(function (id) {
    const m = gmailGet_('/messages/' + id, {
      format: 'metadata', metadataHeaders: ['Date', 'Subject', 'From'],
    });
    const h = {};
    (m.payload.headers || []).forEach(function (x) { h[x.name] = x.value; });
    const date = Utilities.formatDate(new Date(Number(m.internalDate)), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    console.log(date + ' | ' + (h.From || '') + ' | ' + (h.Subject || '(件名なし)') +
      ' | labels=' + (m.labelIds || []).join(','));
  });
}

/** 一括移動を開始する */
function start() {
  PropertiesService.getScriptProperties().setProperty(PROP_TOTAL, '0');
  deleteTriggers_();
  continueCleanup();
}

/** トリガーから呼ばれる本体。時間切れなら次回分のトリガーを仕掛けて終了する */
function continueCleanup() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // 前回分がまだ動いている
  try {
    deleteTriggers_();
    const startedAt = Date.now();
    const props = PropertiesService.getScriptProperties();
    let total = Number(props.getProperty(PROP_TOTAL) || 0);

    while (Date.now() - startedAt < TIME_BUDGET_MS) {
      // ゴミ箱に移したものは検索結果から消えるので、毎回先頭から取り直せばよい
      const res = gmailGet_('/messages', { q: QUERY, maxResults: 500 });
      const ids = (res.messages || []).map(function (m) { return m.id; });
      if (ids.length === 0) {
        props.setProperty(PROP_TOTAL, String(total));
        console.log('完了: 合計 ' + total + ' 件をゴミ箱へ移動しました。');
        return;
      }
      for (let i = 0; i < ids.length && Date.now() - startedAt < TIME_BUDGET_MS; i += PARALLEL) {
        total += trashBatch_(ids.slice(i, i + PARALLEL));
        props.setProperty(PROP_TOTAL, String(total));
        Utilities.sleep(1000); // 1 ユーザーあたりの秒間クォータを超えないように
      }
      console.log('進捗: 累計 ' + total + ' 件');
    }

    // 時間切れ → 1 分後に続きを実行
    ScriptApp.newTrigger(HANDLER).timeBased().after(60 * 1000).create();
    console.log('時間切れのため中断。1 分後に再開します（累計 ' + total + ' 件）。');
  } finally {
    lock.releaseLock();
  }
}

/** 進捗表示 */
function showProgress() {
  const total = PropertiesService.getScriptProperties().getProperty(PROP_TOTAL) || '0';
  const res = gmailGet_('/messages', { q: QUERY, maxResults: 1 });
  console.log('移動済み: ' + total + ' 件 / 残り目安: ' + (res.resultSizeEstimate || 0) + ' 件');
}

/** 途中で止める（再開は continueCleanup() を手動実行） */
function stop() {
  deleteTriggers_();
  console.log('トリガーを削除しました。処理は停止します。');
}

// ── 内部処理 ─────────────────────────────────────────────────

/** 渡された ID を並列でゴミ箱へ移動。成功件数を返す。429/5xx は 1 回だけ再試行 */
function trashBatch_(ids) {
  const token = getAccessToken_();
  const makeReq = function (id) {
    return {
      url: API_BASE + '/messages/' + id + '/trash',
      method: 'post',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
    };
  };
  let pending = ids;
  let ok = 0;
  for (let attempt = 0; attempt < 2 && pending.length; attempt++) {
    const responses = UrlFetchApp.fetchAll(pending.map(makeReq));
    const retry = [];
    responses.forEach(function (r, i) {
      const code = r.getResponseCode();
      if (code === 200) ok++;
      else if (code === 404) { /* 既に削除済み */ }
      else if (code === 429 || code >= 500) retry.push(pending[i]);
      else console.warn('trash 失敗 id=' + pending[i] + ' code=' + code + ' ' + r.getContentText());
    });
    pending = retry;
    if (pending.length) Utilities.sleep(5000);
  }
  if (pending.length) console.warn(pending.length + ' 件は再試行でも失敗（次回の検索で再び対象になります）');
  return ok;
}

function gmailGet_(path, params) {
  const qs = Object.keys(params || {}).map(function (k) {
    const v = params[k];
    return (Array.isArray(v) ? v : [v]).map(function (x) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(x);
    }).join('&');
  }).join('&');
  const url = API_BASE + path + (qs ? '?' + qs : '');
  for (let attempt = 0; ; attempt++) {
    const r = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + getAccessToken_() },
      muteHttpExceptions: true,
    });
    const code = r.getResponseCode();
    if (code === 200) return JSON.parse(r.getContentText());
    if ((code === 429 || code >= 500) && attempt < 3) { Utilities.sleep(5000 * (attempt + 1)); continue; }
    throw new Error('Gmail API エラー ' + code + ': ' + r.getContentText());
  }
}

/** サービスアカウント（ドメイン全体の委任）で TARGET_USER のアクセストークンを取得 */
function getAccessToken_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('token');
  if (cached) return cached;

  const key = loadServiceAccountKey_();

  const now = Math.floor(Date.now() / 1000);
  const b64 = function (obj) {
    return Utilities.base64EncodeWebSafe(JSON.stringify(obj)).replace(/=+$/, '');
  };
  const unsigned = b64({ alg: 'RS256', typ: 'JWT' }) + '.' + b64({
    iss: key.client_email,
    sub: TARGET_USER,
    scope: SCOPE,
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600,
  });
  const sig = Utilities.base64EncodeWebSafe(
    Utilities.computeRsaSha256Signature(unsigned, key.private_key)).replace(/=+$/, '');

  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: unsigned + '.' + sig,
    },
    muteHttpExceptions: true,
  });
  if (r.getResponseCode() !== 200) {
    throw new Error('トークン取得失敗（委任設定・スコープを確認）: ' + r.getContentText());
  }
  const token = JSON.parse(r.getContentText()).access_token;
  cache.put('token', token, 50 * 60);
  return token;
}

/**
 * スクリプト プロパティ SA_KEY_JSON を読み込み、形式を検査して返す。
 * 貼り付け時に改行が「\n」の文字列のまま残ったケースなどは自動で補正する。
 */
function loadServiceAccountKey_() {
  const raw = PropertiesService.getScriptProperties().getProperty('SA_KEY_JSON');
  if (!raw) throw new Error('スクリプト プロパティ SA_KEY_JSON が未設定です。');
  let key;
  try {
    key = JSON.parse(raw.trim());
  } catch (e) {
    throw new Error('SA_KEY_JSON が JSON として読めません。JSON キーファイルの中身を { から } まで丸ごと貼り直してください。');
  }
  if (key.type !== 'service_account' || !key.client_email || !key.private_key) {
    throw new Error('SA_KEY_JSON がサービスアカウントの JSON キーではありません' +
      '（type=' + key.type + ', client_email ' + (key.client_email ? 'あり' : 'なし') +
      ', private_key ' + (key.private_key ? 'あり' : 'なし') + '）。');
  }
  // 改行が「\n」という 2 文字のまま入っている場合は本物の改行に戻す
  let pk = String(key.private_key).replace(/\\n/g, '\n').replace(/\r/g, '').trim();
  if (pk.indexOf('-----BEGIN PRIVATE KEY-----') !== 0 || pk.indexOf('-----END PRIVATE KEY-----') < 0) {
    throw new Error('private_key の形式が不正です（BEGIN/END PRIVATE KEY 行が見つかりません）。' +
      'JSON キーを貼り直すか、鍵を作り直してください。');
  }
  key.private_key = pk + '\n';
  return key;
}

/** 鍵の設定を確認する（秘密鍵そのものはログに出さない） */
function checkKey() {
  const key = loadServiceAccountKey_();
  console.log('client_email : ' + key.client_email);
  console.log('client_id    : ' + key.client_id + '（管理コンソールの委任に登録した ID と一致するか確認）');
  console.log('private_key  : ' + key.private_key.split('\n').length + ' 行');
  Utilities.computeRsaSha256Signature('test', key.private_key);
  console.log('署名テスト OK: 鍵の形式は正しいです。');
}

function deleteTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === HANDLER) ScriptApp.deleteTrigger(t);
  });
}
