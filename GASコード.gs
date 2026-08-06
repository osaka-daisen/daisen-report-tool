/**
 * 大泉衛生株式会社 顧客状況報告書ツール — サーバー側（Google Apps Script）
 *
 * 役割
 *   1) 営業がスマホ／PCのHTMLフォームから送った報告書を、スプレッドシートに1行ずつ蓄積する
 *   2) 総務の宛先へ、書面そのままの見た目でメール通知する（PDF添付も可）
 *
 * 使い方は「セットアップ手順.md」を参照。
 */

/* ============================================================
   ▼▼▼ ここだけ設定してください ▼▼▼
   ============================================================ */

// 合言葉。index.html の SEND_KEY と同じ文字列にすること。
// 一致しない送信は記録もメールもせず弾く（公開URLへのいたずら書き込み防止）。
var SEND_KEY = 'dsn-gvfp7tylco';

// 蓄積先スプレッドシートのID
// （URL https://docs.google.com/spreadsheets/d/★この部分★/edit）
var SHEET_ID = '';

// シート名（無ければ自動で作成されます）
var SHEET_NAME = '顧客状況報告';

// 通知メールの宛先。複数はカンマ区切り
var NOTIFY_TO = '';

// 通知メールにPDFを添付するか（日本語フォントが崩れる場合は false に）
var PDF_ATTACH = true;

// PDFを保存するGoogleドライブのフォルダID（空ならドライブ保存はしない）
var DRIVE_FOLDER_ID = '';

/* ============================================================
   ▲▲▲ 設定はここまで ▲▲▲
   ============================================================ */

var HEADERS = [
  '受信日時', '報告日', '報告者', '報告区分', '相手先種別',
  '受付日', '受付時刻', '連絡の方法', '場所・補足',
  '氏名・団体名', '代表者', '担当者名', '連絡先', '現場名', '住所', '同業者名',
  '月額金額(税抜)', '契約内容',
  '広聴内容',
  '営業担当', '営業の対応', '業務担当', '業務部長', '業務の対応',
  '車番', '運転手', '同乗者1', '同乗者2', '対応日', '対応開始', '対応終了',
  '解除申請', '区', '適用理由', '該当住所', '名称', '契約中止日', '見込み', '届出登録日',
  '備考', '契約書添付', 'PDF'
];

/** 動作確認用（ブラウザでウェブアプリURLを開くと表示される） */
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: '顧客状況報告書 受付エンドポイント 稼働中' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** フォームからのPOSTを受ける */
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);

    // 合言葉チェック（不一致なら何もせず終了）
    if (SEND_KEY && d.key !== SEND_KEY) {
      return json({ ok: false, error: '合言葉が違います' });
    }

    var now = new Date();

    // --- PDF生成（失敗しても記録は続行する） ---
    var pdf = null, pdfUrl = '';
    if (PDF_ATTACH || DRIVE_FOLDER_ID) {
      try {
        pdf = makePdf(d);
        if (pdf && DRIVE_FOLDER_ID) {
          var file = DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(pdf);
          pdfUrl = file.getUrl();
        }
      } catch (pe) {
        pdf = null;
        pdfUrl = 'PDF生成失敗: ' + pe.message;
      }
    }

    // --- スプレッドシートへ追記 ---
    appendRow(d, now, pdfUrl);

    // --- 総務へ通知 ---
    if (NOTIFY_TO) sendMail(d, pdf);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function appendRow(d, now, pdfUrl) {
  if (!SHEET_ID) throw new Error('SHEET_ID が未設定です');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold').setBackground('#e8f0ee').setVerticalAlignment('middle');
    sh.setFrozenRows(1);
  }

  var yn = function (v) { return v ? '○' : ''; };
  sh.appendRow([
    now, d.hokokubi || '', d.hokokusha || '', d.kubun_text || '', d.aite_type || '',
    d.uke_date || '', d.uke_time || '', d.renraku_hoho || '', d.renraku_memo || '',
    d.aite_name || '', d.aite_daihyo || '', d.aite_tanto || '', d.renrakusaki || '',
    d.genba || '', d.jusho || '', d.dogyosha || '',
    d.getsugaku === '' || d.getsugaku == null ? '' : Number(d.getsugaku),
    d.keiyaku_naiyo || '',
    d.naiyo || '',
    d.eigyo_tanto || '', d.eigyo_taio || '', d.gyomu_tanto || '', d.gyomu_bucho || '', d.gyomu_taio || '',
    d.shaban || '', d.driver || '', d.dojo1 || '', d.dojo2 || '',
    d.taio_date || '', d.taio_start || '', d.taio_end || '',
    yn(d.kaijo_use), d.ku || '', d.tekiyo_riyu || '',
    d.kaijo_addr || d.jusho || '', d.kaijo_name || d.aite_name || '',
    d.chushi_date || '', yn(d.chushi_mikomi), d.todokede_date || '',
    d.biko || '', yn(d.keiyakusho_tenpu), pdfUrl
  ]);
}

/* 契約解除登録申請書はチェックが入っているときだけ（大阪市の顧客のみ） */
function needKaijoServer(d) {
  return !!d.kaijo_use;
}

function yenText(d) {
  if (d.getsugaku === '' || d.getsugaku == null) return '—';
  var n = Number(d.getsugaku);
  if (isNaN(n)) return String(d.getsugaku);
  return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '円（税抜）／月';
}

function docTitle(d) {
  var name = d.aite_name || '相手先未入力';
  var kb = d.kubun_text || '報告';
  return '顧客状況報告書_' + name + '_' + kb;
}

function makePdf(d) {
  var html = '<html><head><meta charset="UTF-8"><style>' + DOC_CSS + '</style></head><body>'
    + (d.doc_html || '') + '</body></html>';
  return Utilities.newBlob(html, MimeType.HTML, docTitle(d) + '.html')
    .getAs('application/pdf')
    .setName(docTitle(d) + '.pdf');
}

function sendMail(d, pdf) {
  var subj = '【顧客状況報告】' + (d.kubun_text || '') + '／' + (d.aite_name || '') + '（報告者：' + (d.hokokusha || '') + '）';

  var lines = [
    '営業から顧客状況報告書が届きました。',
    '',
    '報告日　：' + (d.hokokubi || ''),
    '報告者　：' + (d.hokokusha || ''),
    '報告区分：' + (d.kubun_text || ''),
    '相手先　：' + (d.aite_name || '') + '　' + (d.genba || ''),
    '住所　　：' + (d.jusho || ''),
    '連絡先　：' + (d.renrakusaki || ''),
    '月額金額：' + yenText(d)
  ];
  if (needKaijoServer(d)) {
    lines.push('');
    lines.push('★ 契約解除登録申請書（' + (d.ku || '—') + '区／' + (d.tekiyo_riyu || '—') + '）の提出対象です。');
  }
  lines.push('');
  lines.push('── 広聴内容 ──');
  lines.push(d.naiyo || '');

  var body = '<div style="font-family:sans-serif;font-size:14px;line-height:1.7">'
    + lines.join('<br>').replace(/\n/g, '<br>')
    + '</div><hr style="margin:20px 0"><style>' + DOC_CSS + '</style>'
    + (d.doc_html || '');

  var opts = { htmlBody: body, name: '大泉衛生 報告書システム' };
  if (pdf) opts.attachments = [pdf];
  MailApp.sendEmail(NOTIFY_TO, subj, lines.join('\n'), opts);
}

/* 書面の見た目を保つための最小CSS（HTML側の .doc スタイルと対応） */
var DOC_CSS = [
  'body{font-family:"Yu Gothic","Hiragino Kaku Gothic ProN","Meiryo",sans-serif;}',
  '.doc{width:180mm;margin:0 auto;color:#111;font-size:11.5px;line-height:1.5;}',
  '.doc .page + .page{page-break-before:always;margin-top:14mm;}',
  '.doc .topdate{text-align:right;font-size:11px;margin-bottom:3mm;}',
  '.doc .dh{text-align:center;margin-bottom:3mm;}',
  '.doc .dh h2{font-size:18px;margin:0;letter-spacing:.12em;display:inline-block;border-bottom:2px solid #111;padding:0 6mm 1mm;}',
  '.doc .reporter{text-align:right;font-size:11.5px;margin:3mm 0;}',
  '.doc .reporter span{border-bottom:1px solid #111;display:inline-block;min-width:38mm;padding:0 4px;}',
  '.doc table{width:100%;border-collapse:collapse;}',
  '.doc th,.doc td{border:1px solid #111;padding:4px 7px;vertical-align:top;font-size:11.5px;}',
  '.doc th{width:26mm;background:#f4f6f7;text-align:center;font-weight:700;white-space:nowrap;letter-spacing:.25em;}',
  '.doc th.n{letter-spacing:0;}',
  '.doc td.big{height:26mm;} .doc td.big2{height:22mm;}',
  '.doc .sub-t{width:100%;border-collapse:collapse;}',
  '.doc .sub-t td{border:none;border-bottom:1px dotted #777;padding:3px 5px;}',
  '.doc .sub-t tr:last-child td{border-bottom:none;}',
  '.doc .sub-t td.k{width:20mm;text-align:center;border-right:1px solid #111;letter-spacing:.4em;white-space:nowrap;}',
  '.doc .note{font-size:10.5px;color:#333;}',
  '.doc .kubun-line{font-size:12px;font-weight:700;}',
  '.doc .yen{font-size:14px;font-weight:700;}',
  '.doc .stamp{margin-top:5mm;display:flex;justify-content:space-between;align-items:flex-end;gap:6mm;}',
  '.doc .stamp .memo{font-size:11px;text-decoration:underline;}',
  '.doc .stamp table{width:auto;}',
  '.doc .stamp th,.doc .stamp td{width:17mm;text-align:center;letter-spacing:0;font-size:10.5px;padding:3px 2px;}',
  '.doc .stamp td{height:14mm;}',
  '.doc .cap{font-size:11.5px;font-weight:700;margin:3mm 0 1mm;}',
  '.doc .juryo{border:1px solid #111;width:34mm;font-size:11px;margin-bottom:4mm;}',
  '.doc .juryo .h{border-bottom:1px solid #111;text-align:center;padding:2px;}',
  '.doc .juryo .b{height:14mm;}',
  '.doc .kutop{text-align:right;font-size:12px;margin-bottom:3mm;}',
  '.doc .kutop span{border-bottom:1px solid #111;display:inline-block;min-width:28mm;text-align:center;}',
  '@page{size:A4;margin:12mm;}'
].join('\n');

/* ------------------------------------------------------------
   セットアップ確認用。エディタでこの関数を選んで実行すると、
   ダミーデータで1行追加＋通知メールが飛ぶかを確認できます。
   ------------------------------------------------------------ */
function testPost() {
  var dummy = {
    key: SEND_KEY,
    hokokubi: '2026-08-06', hokokusha: 'テスト太郎',
    kubun: ['閉店・廃業'], kubun_text: '閉店・廃業', aite_type: '得意先関係',
    uke_date: '2026-08-06', uke_time: '10:30', renraku_hoho: '電話', renraku_memo: '',
    aite_name: 'テスト商事株式会社', aite_tanto: '山田様',
    genba: 'テストビル1F', jusho: '大阪市東住吉区東田辺1-1-1', renrakusaki: '06-0000-0000',
    getsugaku: 22000, keiyaku_naiyo: '可燃 週2回（火・金）',
    naiyo: '8月末で閉店するため収集を止めてほしいとの連絡あり。',
    kaijo_use: true, ku: '東住吉', tekiyo_riyu: '廃業（閉店）', chushi_date: '2026-08-31',
    doc_html: '<div class="doc"><div class="page"><div class="dh"><h2>テスト出力</h2></div></div></div>'
  };
  var res = doPost({ postData: { contents: JSON.stringify(dummy) } });
  Logger.log(res.getContent());
}
