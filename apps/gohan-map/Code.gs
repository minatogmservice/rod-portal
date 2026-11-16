/**
 * ごはんマップ（梶原さん専用）
 *  - 行ったお店を自分の地図にピンで残す（推し＝ハート）
 *  - 写真＋ひとことメモから、Claude が「いつもの口コミの書き方」で口コミ文を書く
 *  - 「お昼どこ行こ？」と聞くと、自分の記録からおすすめを返す
 *
 * 置き方は README.md を見てください。
 * 必要なもの：スクリプト プロパティ ANTHROPIC_API_KEY（Claude の API キー）
 * スプレッドシートと写真フォルダは、最初に開いたときに自動で作られます。
 */

// 使う Claude。安い順に 'claude-haiku-5-5'（1回1円しない）→ 'claude-sonnet-5-5' → 'claude-opus-5-5'（一番うまい・1回10〜20円）
const MODEL = 'claude-haiku-5-5';
const API_URL = 'https://api.anthropic.com/v1/messages';
const API_VERSION = '2026-10-21';
const COLS = ['id', '店名', '住所', '緯度', '経度', 'ジャンル', '価格帯', '評価', '推し', '雰囲気', '店員',
  '一言メモ', '口コミ', '写真ID', '訪問日', 'GoogleマップURL', '出どころ', '作成', '更新'];
const KEY = {
  id: 'id', name: '店名', address: '住所', lat: '緯度', lng: '経度', genre: 'ジャンル', price: '価格帯',
  rating: '評価', fav: '推し', vibe: '雰囲気', staff: '店員', memo: '一言メモ', review: '口コミ',
  photos: '写真ID', visited: '訪問日', url: 'GoogleマップURL', source: '出どころ', created: '作成', updated: '更新'
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ごはんマップ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ───────── 保存先（初回に自動作成） ───────── */
function props_() { return PropertiesService.getScriptProperties() }

function book_() {
  const p = props_();
  let id = p.getProperty('SHEET_ID'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id) } catch (e) { ss = null } }
  if (!ss) {
    ss = SpreadsheetApp.create('ごはんマップ（データ）');
    p.setProperty('SHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName('お店');
  if (!sh) {
    sh = ss.getSheets()[0].getName() === 'シート1' || ss.getSheets()[0].getName() === 'Sheet1'
      ? ss.getSheets()[0].setName('お店') : ss.insertSheet('お店');
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  let ex = ss.getSheetByName('お手本');
  if (!ex) {
    ex = ss.insertSheet('お手本');
    ex.getRange(1, 1, 1, 2).setValues([['口コミ（自分が書いた文をそのまま貼る）', 'メモ']]).setFontWeight('bold');
    ex.setColumnWidth(1, 600);
  }
  return { ss, sh, ex };
}

function folder_() {
  const p = props_();
  const id = p.getProperty('PHOTO_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id) } catch (e) {} }
  const f = DriveApp.createFolder('ごはんマップ写真');
  p.setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}

function rows_() {
  const { sh } = book_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  const vals = sh.getRange(2, 1, n - 1, COLS.length).getValues();
  return vals.map((r, i) => {
    const o = { _row: i + 2 };
    Object.keys(KEY).forEach(k => { o[k] = r[COLS.indexOf(KEY[k])] });
    o.lat = Number(o.lat) || null; o.lng = Number(o.lng) || null;
    o.rating = Number(o.rating) || 0;
    o.fav = o.fav === true || o.fav === '◎' || o.fav === 'TRUE';
    o.photos = String(o.photos || '').split(',').filter(Boolean);
    ['visited', 'created', 'updated'].forEach(k => { if (o[k] instanceof Date) o[k] = Utilities.formatDate(o[k], 'Asia/Tokyo', k === 'visited' ? 'yyyy-MM-dd' : 'yyyy-MM-dd HH:mm') });
    o.id = String(o.id);
    return o;
  }).filter(o => o.id);
}

function toRow_(o) {
  return COLS.map(c => {
    const k = Object.keys(KEY).find(x => KEY[x] === c);
    let v = o[k];
    if (k === 'fav') v = o.fav ? '◎' : '';
    if (k === 'photos') v = (o.photos || []).join(',');
    return v == null ? '' : v;
  });
}

/* ───────── 画面から呼ぶ関数 ───────── */
function getInit() {
  const { ss } = book_();
  return {
    spots: rows_().map(o => { delete o._row; return o }),
    hasKey: !!props_().getProperty('ANTHROPIC_API_KEY'),
    sheetUrl: ss.getUrl(),
    exampleCount: examples_().length
  };
}

function saveSpot(o) {
  const { sh } = book_();
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    // 新しい写真（dataURL）を Drive に保存
    const f = (o.newPhotos || []).length ? folder_() : null;
    const ids = (o.photos || []).slice();
    (o.newPhotos || []).forEach((d, i) => {
      const m = String(d).match(/^data:(image\/\w+);base64,(.+)$/);
      if (!m) return;
      const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (o.name || '店') + '_' + now.replace(/\D/g, '') + '_' + i + '.jpg');
      ids.push(f.createFile(blob).getId());
    });
    o.photos = ids; delete o.newPhotos;
    o.updated = now;
    const list = rows_();
    const cur = o.id && list.find(x => x.id === o.id);
    if (cur) {
      o.created = cur.created; o.source = cur.source || '手入力';
      sh.getRange(cur._row, 1, 1, COLS.length).setValues([toRow_(o)]);
    } else {
      o.id = o.id || 's' + Utilities.getUuid().slice(0, 7);
      o.created = now; o.source = o.source || '手入力';
      sh.appendRow(toRow_(o));
    }
    return o;
  } finally { lock.releaseLock() }
}

function deleteSpot(id) {
  const { sh } = book_();
  const cur = rows_().find(x => x.id === id);
  if (cur) sh.deleteRow(cur._row);
  return true;
}

function getPhoto(fileId) {
  const b = DriveApp.getFileById(fileId).getBlob();
  return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
}

/** Googleマップの共有リンク（maps.app.goo.gl など）から店名・場所を読む */
function resolveMapsLink(url) {
  url = String(url || '').trim().match(/https?:\/\/\S+/) ? String(url).match(/https?:\/\/\S+/)[0] : '';
  if (!url) throw new Error('GoogleマップのURLが見つかりませんでした');
  let final = url;
  for (let i = 0; i < 5; i++) {
    const r = UrlFetchApp.fetch(final, { followRedirects: false, muteHttpExceptions: true });
    const loc = r.getHeaders()['Location'] || r.getHeaders()['location'];
    if (!loc) break;
    final = loc;
  }
  const dec = decodeURIComponent(final);
  const out = { url: url };
  let m = dec.match(/\/place\/([^/@?]+)/);
  if (m) out.name = m[1].replace(/\+/g, ' ');
  m = dec.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) || dec.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/) || dec.match(/[?&](?:q|ll)=(-?\d+\.\d+),(-?\d+\.\d+)/);
  if (m) { out.lat = Number(m[1]); out.lng = Number(m[2]) }
  if (!out.name) { m = dec.match(/[?&]q=([^&]+)/); if (m && !/^-?\d/.test(m[1])) out.name = m[1].replace(/\+/g, ' ') }
  if (out.lat) {
    try {
      const g = Maps.newGeocoder().setLanguage('ja').reverseGeocode(out.lat, out.lng);
      if (g.results && g.results[0]) out.address = shortAddr_(g.results[0].formatted_address);
    } catch (e) {}
  }
  return out;
}

/** 店名や住所・エリア名から場所を探す */
function geocode(q) {
  const g = Maps.newGeocoder().setLanguage('ja').setRegion('jp').geocode(q);
  return (g.results || []).slice(0, 5).map(r => ({
    address: shortAddr_(r.formatted_address), lat: r.geometry.location.lat, lng: r.geometry.location.lng
  }));
}

function shortAddr_(a) { return String(a || '').replace(/^日本、?\s*/, '').replace(/〒\d{3}-\d{4}\s*/, '') }

/** Google Takeout の「Reviews.json」（自分の口コミ）を取り込む */
function importReviews(items) {
  const { sh } = book_();
  const list = rows_();
  const have = {};
  list.forEach(x => { have[x.url || ''] = 1; have[x.name + '|' + x.lat] = 1 });
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  const add = [];
  items.forEach(it => {
    if (!it.name) return;
    if ((it.url && have[it.url]) || have[it.name + '|' + it.lat]) return;
    have[it.url || ''] = 1; have[it.name + '|' + it.lat] = 1;
    add.push(toRow_({
      id: 's' + Utilities.getUuid().slice(0, 7), name: it.name, address: shortAddr_(it.address), lat: it.lat, lng: it.lng,
      genre: '', price: '', rating: it.rating || '', fav: Number(it.rating) >= 5, review: it.text || '',
      visited: it.date || '', url: it.url || '', source: 'Googleの口コミ', created: now, updated: now, photos: []
    }));
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, COLS.length).setValues(add);
  return { added: add.length, skipped: items.length - add.length };
}

/* ───────── Claude ───────── */
function claude_(system, messages, opt) {
  const key = props_().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('Claude の API キーが未設定です（README の手順3）');
  const body = {
    model: MODEL, max_tokens: 8000, system: system, messages: messages,
    output_config: Object.assign({ effort: (opt && opt.effort) || 'low' }, opt && opt.format ? { format: opt.format } : {})
  };
  // Opus / Sonnet は、断られたときに別のモデルで書き直す設定（Haiku には無い）
  const fb = !/haiku/.test(MODEL);
  if (fb) body.fallbacks = 'default';
  const r = UrlFetchApp.fetch(API_URL, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: Object.assign({ 'x-api-key': key, 'anthropic-version': API_VERSION }, fb ? { 'anthropic-beta': 'server-side-fallback-2026-07-01' } : {}),
    payload: JSON.stringify(body)
  });
  const code = r.getResponseCode(), res = JSON.parse(r.getContentText() || '{}');
  if (code !== 200) throw new Error('Claude エラー（' + code + '）：' + (res.error && res.error.message || ''));
  if (res.stop_reason === 'refusal') throw new Error('Claude が書けませんでした。メモを少し変えてもう一度試してください');
  return (res.content || []).filter(b => b.type === 'text').map(b => b.text).join('').trim();
}

/** お手本＝「お手本」シートに貼った文＋取り込んだ／仕上げた口コミ（長めで新しいもの） */
function examples_() {
  const { ex } = book_();
  const n = ex.getLastRow();
  const pasted = n > 1 ? ex.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0]).trim()).filter(Boolean) : [];
  const mine = rows_().filter(o => String(o.review).length >= 60)
    .sort((a, b) => String(b.updated).localeCompare(String(a.updated)))
    .map(o => '【' + o.name + (o.rating ? ' ★' + o.rating : '') + '】\n' + o.review);
  return pasted.concat(mine).slice(0, 25);
}

/**
 * 口コミ文を書く
 * d = { name, genre, price, rating, fav, vibe, staff, memo, visited, address, photos:[dataURL...], draft, ask }
 */
function writeReview(d) {
  const ex = examples_();
  const system = [
    'あなたは、Googleマップのローカルガイド（レベル8・口コミ348件）である「俺」の代わりに口コミを書く代筆係です。',
    '「俺」のプロフィール文：「ジャンル問わず来店したことのある場所を偏見無しで投稿していこうと思います (^-^)」',
    '',
    '# 書き方',
    '- 下の「お手本」は全部「俺」が実際に書いた口コミです。言い回し・文の長さ・句読点（「、、、」「！！」「ふつーーの」など伸ばし方）・顔文字の使い方・改行の入れ方・丁寧語の崩し方をよく真似してください。',
    '- お手本に無いキャッチーな宣伝文句、ハッシュタグ、絵文字の連発、「いかがでしたか」のようなAIっぽい締めは使わない。',
    '- 事実はメモと写真から分かることだけ。写真に写っていない料理名・値段・店員の言動を作らない。分からないことは書かない。',
    '- 店員の印象・雰囲気・駐車場や行きやすさなど、メモにあることは必ず入れる。',
    '- 長さはお手本の平均くらい（目安 80〜250文字）。',
    '- 出力は口コミ本文だけ。前置きや説明、かぎかっこで囲むことはしない。',
    '',
    '# お手本（俺が書いた口コミ）',
    ex.length ? ex.map((t, i) => '--- ' + (i + 1) + '\n' + t).join('\n') :
      '（まだ無し。丁寧語ベースで、ちょっとくだけた関西の30代男性っぽい自然な文で。例：「前々からずっと行きたかったお店にようやく行けました！！」「駐車場が少し停めにくいのでやや行きづらい印象ですが…」）'
  ].join('\n');

  const info = [
    '店名：' + (d.name || ''),
    d.genre && 'ジャンル：' + d.genre,
    d.price && '価格帯：' + d.price,
    d.rating && '自分の評価：★' + d.rating,
    d.fav && 'かなり気に入った（推しの店）',
    d.visited && '行った日：' + d.visited,
    d.address && '場所：' + d.address,
    d.vibe && '雰囲気メモ：' + d.vibe,
    d.staff && '店員の印象メモ：' + d.staff,
    d.memo && 'ひとこと：' + d.memo
  ].filter(Boolean).join('\n');

  const content = [];
  (d.photos || []).slice(0, 6).forEach(u => {
    const m = String(u).match(/^data:(image\/\w+);base64,(.+)$/);
    if (m) content.push({ type: 'image', source: { type: 'base64', media_type: m[1], data: m[2] } });
  });
  let text = (content.length ? '写真' + content.length + '枚と、' : '') + '下のメモから口コミを書いてください。\n\n' + info;
  if (d.draft && d.ask) text += '\n\n# 前に書いてもらった文\n' + d.draft + '\n\n# 直してほしいこと\n' + d.ask;
  content.push({ type: 'text', text: text });
  return claude_(system, [{ role: 'user', content: content }], { effort: 'medium' });
}

/**
 * お昼の相談
 * q = { text, history:[{role,text}], lat, lng, area }
 */
function askLunch(q) {
  let spots = rows_();
  let here = '';
  if (q.lat && q.lng) here = '今いる場所：緯度' + q.lat.toFixed(4) + '・経度' + q.lng.toFixed(4);
  else if (q.area) {
    here = '今いるあたり：' + q.area;
    try { const g = geocode(q.area)[0]; if (g) { q.lat = g.lat; q.lng = g.lng } } catch (e) {}
  }
  if (q.lat && q.lng) {
    spots.forEach(s => { s.km = s.lat ? dist_(q.lat, q.lng, s.lat, s.lng) : 999 });
    spots.sort((a, b) => a.km - b.km);
    spots = spots.slice(0, 120);
  }
  const lines = spots.map(s => [
    s.id, s.name, s.genre || '', s.price || '', s.rating ? '★' + s.rating : '', s.fav ? '推し' : '',
    s.km != null && s.km < 999 ? s.km.toFixed(1) + 'km' : '', (s.address || '').slice(0, 20),
    s.visited ? '最終' + s.visited : '', String(s.memo || s.review || '').replace(/\s+/g, ' ').slice(0, 90)
  ].join(' | '));
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd（E）HH:mm');
  const system = [
    'あなたは「俺」専用のごはん相談相手です。関西弁まじりで気楽に、短めに答えてください（3〜6行くらい）。',
    '基本は下の「俺の記録」の店から選ぶ。推し・評価が高い店、俺のメモの内容（雰囲気・店員・駐車場など）を理由にして2〜3軒すすめる。',
    '最近行った店ばかりにならないよう、しばらく行ってない店も混ぜる。記録に無い店を作り話でおすすめしない。',
    '距離が分かるときは近さも考える（お昼休みなら車で15分以内くらいが目安）。',
    '現在：' + now + (here ? '／' + here : ''),
    '',
    '# 俺の記録（id | 店名 | ジャンル | 価格帯 | 評価 | 推し | 距離 | 住所 | 最終訪問 | メモ）',
    lines.join('\n') || '（まだ記録なし）'
  ].join('\n');
  const msgs = (q.history || []).slice(-8).map(h => ({ role: h.role, content: h.text }));
  msgs.push({ role: 'user', content: q.text });
  const out = claude_(system, msgs, {
    format: {
      type: 'json_schema',
      schema: {
        type: 'object', additionalProperties: false, required: ['reply', 'picks'],
        properties: {
          reply: { type: 'string', description: '俺への返事' },
          picks: { type: 'array', items: { type: 'string' }, description: 'すすめた店の id（記録にある id だけ）' }
        }
      }
    }
  });
  const j = JSON.parse(out);
  const ids = {}; spots.forEach(s => ids[s.id] = 1);
  j.picks = (j.picks || []).filter(id => ids[id]);
  return j;
}

function dist_(a, b, c, d) {
  const R = 6371, t = Math.PI / 180, x = (d - b) * t * Math.cos((a + c) / 2 * t), y = (c - a) * t;
  return Math.sqrt(x * x + y * y) * R;
}

/** 動作確認用：エディタで実行すると、保存先とAPIキーの状態がログに出ます */
function setupCheck() {
  const { ss } = book_(); folder_();
  Logger.log('データ：' + ss.getUrl());
  Logger.log('APIキー：' + (props_().getProperty('ANTHROPIC_API_KEY') ? 'OK' : '未設定'));
}
