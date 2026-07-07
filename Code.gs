/**
 * 冰島停車場資訊系統 後端 v2
 * - Reports/Log：第1列英文鍵、第2列中文標題、資料自第3列起
 * - 所有欄位皆按第1列（主表為第2列）標題動態定位，欄位搬動不影響程式
 * - 無 email 通知（改由前端顯示待審卡）
 * 初次/改版使用：執行 setupV2() 一次（重建 Reports/Log、建「使用說明」、掛觸發器）
 * 部署：新增部署作業 → 網頁應用程式 → 執行身分:我 / 存取:所有人 → exec URL 填入 index.html 的 GAS_URL
 */
var SID = '10W3xfo4RWzDL2iBK8rDEQL-4u1BYc_w0vMvst8Bcwtw';
function ss_(){ return SpreadsheetApp.openById(SID); }
var MAIN_SHEET = '停車收費狀況';
var REP_SHEET = 'Reports';
var LOG_SHEET = 'Log';
var HELP_SHEET = '使用說明';
var TZ = 'Asia/Taipei';
var DATA_ROW = 3; // Reports/Log 資料起始列

/* Reports 欄定義：[英文鍵, 中文標題] */
var REP_COLS = [
  ['Timestamp','回報時間'], ['Type','類型'], ['Spot','地點(英文名)'], ['Fee','是否收費'],
  ['Amount','金額ISK'], ['App','支援APP'], ['WC','廁所'], ['Note','補充說明'],
  ['NewName','新地點名稱'], ['GmapLink','地圖連結'], ['Idea','系統建議'], ['Nick','回報者'],
  ['Action','★審核(選核可/拒絕)'], ['Status','狀態'], ['DoneAt','處理時間']
];
var LOG_COLS = [
  ['Time','時間'], ['Action','動作'], ['Type','類型'], ['Spot','地點'],
  ['Detail','變更內容'], ['ReportRow','回報列號']
];

/* ===== 標題定位工具 ===== */
function colMap_(sh, headerRow){
  var vals = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0];
  var m = {};
  for(var i = 0; i < vals.length; i++){ var k = String(vals[i]).trim(); if(k) m[k] = i + 1; }
  return m;
}
/* 主表：第2列標題，用關鍵字包含比對（標題含換行/註解） */
function mainCols_(sh){
  var vals = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];
  function find(kw){
    for(var i = 0; i < vals.length; i++){ if(String(vals[i]).indexOf(kw) > -1) return i + 1; }
    return -1;
  }
  return {
    region: find('方位'), en: find('英文'), zh: find('中文'), alias: find('別名'),
    fee: find('是否需繳費'), parka: find('Parka'), easy: find('Easy'),
    how: find('其他繳費'), amount: find('金額'), date: find('更新日期'), by: find('更新者'),
    wc: find('有廁所'), wcfee: find('廁所收費'), note: find('其他資訊'),
    lat: find('Latitude'), lon: find('Longitude')
  };
}

/* ===== 初始化 / 改版遷移 ===== */
function setupV2(){
  var ss = ss_();
  rebuild_(ss, REP_SHEET, REP_COLS);
  rebuild_(ss, LOG_SHEET, LOG_COLS);
  var rep = ss.getSheetByName(REP_SHEET);
  var cm = colMap_(rep, 1);
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['核可','拒絕'], true).setAllowInvalid(false).build();
  rep.getRange(DATA_ROW, cm['Action'], 997).setDataValidation(rule);
  buildHelp_(ss);
  var has = ScriptApp.getProjectTriggers().some(function(t){ return t.getHandlerFunction() === 'onEditReview'; });
  if(!has) ScriptApp.newTrigger('onEditReview').forSpreadsheet(ss_()).onEdit().create();
  return 'setupV2 ok';
}
function rebuild_(ss, name, cols){
  var old = ss.getSheetByName(name);
  if(old) ss.deleteSheet(old);
  var sh = ss.insertSheet(name);
  sh.appendRow(cols.map(function(c){ return c[0]; }));
  sh.appendRow(cols.map(function(c){ return c[1]; }));
  sh.getRange(1, 1, 1, cols.length).setFontColor('#999999').setFontSize(9);
  sh.getRange(2, 1, 1, cols.length).setFontWeight('bold').setBackground('#dce9f1');
  sh.setFrozenRows(2);
}

/* ===== 前端回報入口 ===== */
function doPost(e){
  var out;
  try{
    var p = JSON.parse(e.postData.contents || '{}');
    if(p.website) return json({ok:true}); // honeypot：假裝成功
    if(!p.nick) return json({ok:false, error:'nick required'});
    var typeZh = p.type === 'new' ? '新增地點' : (p.type === 'idea' ? '系統建議' : '更正資訊');
    var lock = LockService.getScriptLock(); lock.waitLock(10000);
    var ss = ss_();
    var rep = ss.getSheetByName(REP_SHEET);
    if(!rep){ setupV2(); rep = ss.getSheetByName(REP_SHEET); }
    var cm = colMap_(rep, 1);
    var row = new Array(rep.getLastColumn()).fill('');
    function set(key, val){ if(cm[key]) row[cm[key]-1] = val || ''; }
    set('Timestamp', Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss'));
    set('Type', typeZh); set('Spot', p.spot); set('Fee', p.fee); set('Amount', p.amount);
    set('App', p.app); set('WC', p.wc); set('Note', p.note); set('NewName', p.newName);
    set('GmapLink', p.gmap); set('Idea', p.idea); set('Nick', p.nick); set('Status', '待審');
    rep.appendRow(row);
    lock.releaseLock();
    out = {ok:true};
  }catch(err){ out = {ok:false, error:String(err)}; }
  return json(out);
}
function json(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ===== 審核（Reports「★審核」欄下拉觸發）===== */
function onEditReview(e){
  var sh = e.range.getSheet();
  if(sh.getName() !== REP_SHEET || e.range.getNumRows() !== 1) return;
  var cm = colMap_(sh, 1);
  if(e.range.getColumn() !== cm['Action']) return;
  var row = e.range.getRow(); if(row < DATA_ROW) return;
  var action = String(e.value || '');
  if(action !== '核可' && action !== '拒絕') return;
  var status = sh.getRange(row, cm['Status']).getValue();
  if(status === '已核可' || status === '已拒絕') return; // 防重複執行

  var vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  function get(key){ return cm[key] ? String(vals[cm[key]-1] || '').trim() : ''; }
  var r = { type:get('Type'), spot:get('Spot'), fee:get('Fee'), amount:get('Amount'),
            app:get('App'), wc:get('WC'), note:get('Note'), newName:get('NewName'),
            gmap:get('GmapLink'), idea:get('Idea'), nick:get('Nick') };
  var ss = ss_();
  var now = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm');
  var detail = '';

  if(action === '核可'){
    try{
      if(r.type.indexOf('更正') > -1)      detail = applyUpdate(ss, r);
      else if(r.type.indexOf('新增') > -1) detail = applyNew(ss, r);
      else                                 detail = '建議已閱';
    }catch(err){
      sh.getRange(row, cm['Status']).setValue('錯誤:' + String(err).slice(0,80));
      return;
    }
  }
  sh.getRange(row, cm['Status']).setValue(action === '核可' ? '已核可' : '已拒絕');
  sh.getRange(row, cm['DoneAt']).setValue(now);
  writeLog_(ss, [now, action, r.type, r.spot || r.newName, detail || (r.idea||'').slice(0,120), row]);
}
function writeLog_(ss, arr){
  var log = ss.getSheetByName(LOG_SHEET);
  if(!log) return;
  var cm = colMap_(log, 1);
  var keys = ['Time','Action','Type','Spot','Detail','ReportRow'];
  var row = new Array(log.getLastColumn()).fill('');
  for(var i = 0; i < keys.length; i++){ if(cm[keys[i]]) row[cm[keys[i]]-1] = arr[i]; }
  log.appendRow(row);
}

/* 更正既有地點：只覆寫回報有填的欄位 */
function applyUpdate(ss, r){
  var sh = ss.getSheetByName(MAIN_SHEET);
  var C = mainCols_(sh);
  var names = sh.getRange(1, C.en, sh.getLastRow()).getValues();
  var target = -1;
  for(var i = 0; i < names.length; i++){
    if(String(names[i][0]).trim() === String(r.spot).trim()){ target = i + 1; break; }
  }
  if(target < 0) throw new Error('主表找不到地點: ' + r.spot);
  var changed = [];
  function setIf(col, val, label){ if(col > 0 && val !== '' && val != null){ sh.getRange(target, col).setValue(val); changed.push(label + '=' + val); } }
  setIf(C.fee, r.fee, '收費');
  setIf(C.amount, r.amount, '金額');
  if(r.app){
    var pk = (r.app === 'Parka' || r.app === '兩者皆可') ? 'Y' : (r.app === '無' ? '' : null);
    var ez = (r.app === 'EasyPark' || r.app === '兩者皆可') ? 'Y' : (r.app === '無' ? '' : null);
    if(pk !== null){ sh.getRange(target, C.parka).setValue(pk); changed.push('Parka=' + pk); }
    if(ez !== null){ sh.getRange(target, C.easy).setValue(ez); changed.push('EasyPark=' + ez); }
  }
  if(r.wc){
    var wcv = r.wc.indexOf('Y') === 0 ? 'Y' : 'N';
    sh.getRange(target, C.wc).setValue(wcv); changed.push('廁所=' + wcv);
    if(r.wc.indexOf('收費') > -1){ sh.getRange(target, C.wcfee).setValue('Y'); }
    if(wcv === 'Y' && r.wc.indexOf('免費') > -1){ sh.getRange(target, C.wcfee).setValue('N'); }
  }
  if(r.note){
    var old = sh.getRange(target, C.note).getValue();
    var d = Utilities.formatDate(new Date(), TZ, 'MM/dd');
    sh.getRange(target, C.note).setValue((old ? old + ' ' : '') + '更新' + d + '：' + r.note);
    changed.push('補充(附加)');
  }
  sh.getRange(target, C.date).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd'));
  sh.getRange(target, C.by).setValue(r.nick);
  return '主表第' + target + '列 | ' + changed.join('、');
}

/* 新增地點：從 Google Maps 連結抽座標 */
function applyNew(ss, r){
  if(!r.newName) throw new Error('缺新地點名稱');
  var co = coordsFromLink(String(r.gmap || ''));
  var sh = ss.getSheetByName(MAIN_SHEET);
  var C = mainCols_(sh);
  var rowVals = new Array(sh.getLastColumn()).fill('');
  function put(col, val){ if(col > 0) rowVals[col-1] = val; }
  put(C.region, '待分區');
  put(C.en, r.newName);
  put(C.zh, r.newName);
  put(C.note, (r.note || '') + (co ? '' : '（座標待補：' + (r.gmap || '無連結') + '）'));
  put(C.date, Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd'));
  put(C.by, r.nick);
  if(co){ put(C.lat, co[0]); put(C.lon, co[1]); }
  sh.appendRow(rowVals);
  return '主表新增列 | ' + r.newName + (co ? ' @' + co.join(',') : '（無座標）');
}
function coordsFromLink(u){
  if(!u) return null;
  if(/goo\.gl|maps\.app/.test(u)){
    try{
      var resp = UrlFetchApp.fetch(u, {followRedirects:false, muteHttpExceptions:true});
      var loc = resp.getHeaders()['Location'] || resp.getHeaders()['location'];
      if(loc) u = loc;
    }catch(e){}
  }
  var m = u.match(/!3d(-?[\d.]+)!4d(-?[\d.]+)/) || u.match(/[?&]q=(-?[\d.]+),(-?[\d.]+)/) || u.match(/@(-?[\d.]+),(-?[\d.]+)/);
  return m ? [Number(Number(m[1]).toFixed(5)), Number(Number(m[2]).toFixed(5))] : null;
}

/* ===== 使用說明分頁（內容由 setupV2 產生，詳試算表「使用說明」）===== */
function buildHelp_(ss){
  var old = ss.getSheetByName(HELP_SHEET);
  if(old) ss.deleteSheet(old);
  var sh = ss.insertSheet(HELP_SHEET, 0);
  sh.setColumnWidth(1, 130); sh.setColumnWidth(2, 760);
  var L = [
    ['■ 冰島停車場資訊系統｜管理者使用說明', ''],
    ['', ''],
    ['網站', 'https://iceland-parking.vercel.app（資料即時同步本試算表「停車收費狀況」分頁）'],
    ['資料流', '群友在網站按「回報」→ 進 Reports 分頁排隊 → 管理員審核 → 核可後自動更新主表 → 網站即時反映'],
    ['誰是管理員', '擁有本試算表「編輯」權限的人就是管理員（用共用功能加人/移除，不需要密碼）'],
    ['', ''],
    ['■ 審核步驟（30 秒）', ''],
    ['1', '打開 Reports 分頁，看「狀態」欄為「待審」的列'],
    ['2', '看回報內容是否合理（誰報的見「回報者」欄）'],
    ['3', '想修改內容：直接改該列的儲存格（例如金額打錯，改成正確值），改完再進行下一步'],
    ['4', '在「★審核」欄下拉選「核可」或「拒絕」→ 系統立刻自動處理，狀態變「已核可/已拒絕」'],
    ['注意', '選核可的瞬間就會套用，套用的是當下列上的值（所以要改先改再核）。已核可的列再改動不會重複套用'],
    ['', ''],
    ['■ 核可後資料怎麼變（重要）', ''],
    ['類型：更正資訊', '主表對應地點（按「地點英文名」比對）：回報「有填」的欄位才會被覆寫（收費/金額/APP/廁所），沒填的完全不動。「補充說明」是附加到原文後面（格式：更新MM/DD：內容），不會覆蓋歷史。「更新日期」自動填當天、「更新者」填回報者暱稱。座標、名稱、方位永遠不會被回報改動'],
    ['類型：新增地點', '主表最底部新增一列：英文名=回報名稱、座標=自動從 Google Maps 連結抽出（抽不到會在補充欄註明待補）。「方位」會是「待分區」、中文名暫用英文名 → 請管理員之後手動補上方位與中文名（不補也不影響網站顯示）'],
    ['類型：系統建議', '主表完全不動，核可=已閱，內容記錄在 Log 分頁'],
    ['拒絕', '主表完全不動，該列只標記「已拒絕」'],
    ['', ''],
    ['■ 反悔 / 出錯了怎麼辦', ''],
    ['查記錄', 'Log 分頁有每次核可的完整記錄：時間、地點、具體改了哪些欄位'],
    ['回復資料', '檔案 → 版本記錄 → 查看版本記錄，可回復到任何時間點'],
    ['狀態欄顯示「錯誤:」', '通常是主表找不到該地點英文名（可能被改名），手動處理該筆即可'],
    ['', ''],
    ['■ 請勿踩雷', ''],
    ['1', 'Reports / Log 的前兩列是標題（第1列英文供程式用、第2列中文供人看），請勿刪除或修改第1列'],
    ['2', '主表「景點(英文、冰島文)」欄是系統比對的鍵值，改名前先確認沒有待審回報指向舊名'],
    ['3', '「狀態」「處理時間」欄由系統自動填寫，請勿手動修改'],
    ['4', '欄位順序可以搬動（程式按標題找欄位），但標題文字不要改'],
    ['', ''],
    ['■ 其他', ''],
    ['通知', '本系統不寄 email。網站頂部會顯示「N 筆回報待審核」，管理員看到就來處理即可'],
    ['My Maps 舊地圖', '保留作備援，不會自動更新；網站資料永遠最新'],
    ['原始碼備份', 'GitHub：BryanH68400/iceland-parking（前端 index.html + 本程式 Code.gs）']
  ];
  for(var i = 0; i < L.length; i++){
    sh.getRange(i+1, 1).setValue(L[i][0]);
    sh.getRange(i+1, 2).setValue(L[i][1]).setWrap(true);
    if(String(L[i][0]).indexOf('■') === 0){
      sh.getRange(i+1, 1, 1, 2).setFontWeight('bold').setBackground('#dce9f1').setFontSize(11);
    }
  }
  sh.getRange(1, 1, 1, 2).setFontSize(13);
}

/* 健康檢查 */
function doGet(){ return json({ok:true, service:'iceland-parking', ver:2, time:new Date().toISOString()}); }
