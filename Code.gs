/**
 * 冰島停車場資訊系統 後端（Apps Script，資料指向試算表副本）
 * 功能：doPost 收前端回報 → Reports 分頁；管理員在 Reports 的「審核」欄選核可/拒絕 → 自動更新主表 + Log + email 通知
 * 初次使用：執行 setup() 一次（建分頁、下拉、觸發器）
 * 部署：新增部署作業 → 網頁應用程式 → 執行身分:我 / 存取:所有人 → exec URL 填入 index.html 的 GAS_URL
 */
var SID = '10W3xfo4RWzDL2iBK8rDEQL-4u1BYc_w0vMvst8Bcwtw';
function ss_(){ return SpreadsheetApp.openById(SID); }
var MAIN_SHEET = '停車收費狀況';
var REP_SHEET = 'Reports';
var LOG_SHEET = 'Log';
var ADMIN_EMAIL = 'b0908568759@gmail.com';
var TZ = 'Asia/Taipei';

/* 主表欄位（1-based）*/
var C = { region:1, en:2, zh:3, alias:4, fee:5, parka:6, easy:7, how:8, amount:9, date:10, by:11, wc:12, wcfee:13, note:14, lat:15, lon:16 };
var REP_HEAD = ['Timestamp','Type','Spot','Fee','Amount','App','WC','Note','NewName','GmapLink','Idea','Nick','審核','狀態','處理時間'];
var COL_ACTION = 13, COL_STATUS = 14, COL_DONE = 15; // Reports 1-based

function setup(){
  var ss = ss_();
  var rep = ss.getSheetByName(REP_SHEET) || ss.insertSheet(REP_SHEET);
  if(rep.getLastRow() === 0){
    rep.appendRow(REP_HEAD);
    rep.getRange(1,1,1,REP_HEAD.length).setFontWeight('bold').setBackground('#dce9f1');
    rep.setFrozenRows(1);
  }
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['核可','拒絕'], true).setAllowInvalid(false).build();
  rep.getRange(2, COL_ACTION, 999).setDataValidation(rule);
  var log = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
  if(log.getLastRow() === 0){
    log.appendRow(['Time','Action','Type','Spot','Detail','ReportRow']);
    log.setFrozenRows(1);
  }
  // installable onEdit（simple trigger 無法寄信）
  var has = ScriptApp.getProjectTriggers().some(function(t){ return t.getHandlerFunction() === 'onEditReview'; });
  if(!has) ScriptApp.newTrigger('onEditReview').forSpreadsheet(ss_()).onEdit().create();
  return 'setup ok';
}

/* ===== 前端回報入口 ===== */
function doPost(e){
  var out;
  try{
    var p = JSON.parse(e.postData.contents || '{}');
    if(p.website) return json({ok:true}); // honeypot：假裝成功
    if(!p.nick) return json({ok:false, error:'nick required'});
    var lock = LockService.getScriptLock(); lock.waitLock(10000);
    var ss = ss_();
    var rep = ss.getSheetByName(REP_SHEET) || (setup(), ss.getSheetByName(REP_SHEET));
    var ts = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss');
    rep.appendRow([ts, p.type||'', p.spot||'', p.fee||'', p.amount||'', p.app||'', p.wc||'', p.note||'', p.newName||'', p.gmap||'', p.idea||'', p.nick||'', '', '待審', '']);
    lock.releaseLock();
    try{
      MailApp.sendEmail(ADMIN_EMAIL, '[冰島停車場] 新回報：' + (p.type==='new'?('新增 '+(p.newName||'')):(p.type==='idea'?'系統建議':('更正 '+(p.spot||'')))),
        '回報者：' + p.nick + '\n類型：' + p.type + '\n\n請開啟試算表 Reports 分頁審核（在「審核」欄選 核可/拒絕）：\n' + ss.getUrl());
    }catch(mailErr){}
    out = {ok:true};
  }catch(err){ out = {ok:false, error:String(err)}; }
  return json(out);
}
function json(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ===== 審核（Reports 分頁「審核」欄下拉觸發）===== */
function onEditReview(e){
  var sh = e.range.getSheet();
  if(sh.getName() !== REP_SHEET || e.range.getColumn() !== COL_ACTION || e.range.getNumRows() !== 1) return;
  var row = e.range.getRow(); if(row < 2) return;
  var action = String(e.value || '');
  if(action !== '核可' && action !== '拒絕') return;
  var status = sh.getRange(row, COL_STATUS).getValue();
  if(status === '已核可' || status === '已拒絕') return; // 防重複執行

  var vals = sh.getRange(row, 1, 1, REP_HEAD.length).getValues()[0];
  var r = { ts:vals[0], type:vals[1], spot:vals[2], fee:vals[3], amount:vals[4], app:vals[5], wc:vals[6], note:vals[7], newName:vals[8], gmap:vals[9], idea:vals[10], nick:vals[11] };
  var ss = ss_();
  var now = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm');
  var detail = '';

  if(action === '核可'){
    try{
      if(r.type === 'update')      detail = applyUpdate(ss, r);
      else if(r.type === 'new')    detail = applyNew(ss, r);
      else                         detail = 'idea noted';
    }catch(err){
      sh.getRange(row, COL_STATUS).setValue('錯誤:' + String(err).slice(0,80));
      return;
    }
  }
  sh.getRange(row, COL_STATUS).setValue(action === '核可' ? '已核可' : '已拒絕');
  sh.getRange(row, COL_DONE).setValue(now);
  var log = ss.getSheetByName(LOG_SHEET);
  if(log) log.appendRow([now, action, r.type, r.spot || r.newName, detail || (r.idea||'').slice(0,120), row]);
}

/* 更正既有地點：只覆寫回報有填的欄位 */
function applyUpdate(ss, r){
  var sh = ss.getSheetByName(MAIN_SHEET);
  var names = sh.getRange(1, C.en, sh.getLastRow()).getValues();
  var target = -1;
  for(var i = 0; i < names.length; i++){
    if(String(names[i][0]).trim() === String(r.spot).trim()){ target = i + 1; break; }
  }
  if(target < 0) throw new Error('main row not found: ' + r.spot);
  var changed = [];
  function setIf(col, val){ if(val !== '' && val != null){ sh.getRange(target, col).setValue(val); changed.push(col + '=' + val); } }
  setIf(C.fee, r.fee);
  setIf(C.amount, r.amount);
  if(r.app){
    var pk = (r.app === 'Parka' || r.app === '兩者皆可') ? 'Y' : (r.app === '無' ? '' : null);
    var ez = (r.app === 'EasyPark' || r.app === '兩者皆可') ? 'Y' : (r.app === '無' ? '' : null);
    if(pk !== null){ sh.getRange(target, C.parka).setValue(pk); changed.push('parka=' + pk); }
    if(ez !== null){ sh.getRange(target, C.easy).setValue(ez); changed.push('easy=' + ez); }
  }
  if(r.wc){
    var wcv = r.wc.indexOf('Y') === 0 ? 'Y' : 'N';
    sh.getRange(target, C.wc).setValue(wcv); changed.push('wc=' + wcv);
    if(r.wc.indexOf('收費') > -1){ sh.getRange(target, C.wcfee).setValue('Y'); }
    if(wcv === 'Y' && r.wc.indexOf('免費') > -1){ sh.getRange(target, C.wcfee).setValue('N'); }
  }
  if(r.note){
    var old = sh.getRange(target, C.note).getValue();
    var d = Utilities.formatDate(new Date(), TZ, 'MM/dd');
    sh.getRange(target, C.note).setValue((old ? old + ' ' : '') + '更新' + d + '：' + r.note);
    changed.push('note+');
  }
  sh.getRange(target, C.date).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd'));
  sh.getRange(target, C.by).setValue(r.nick);
  return 'row ' + target + ' | ' + changed.join(', ');
}

/* 新增地點：從 Google Maps 連結抽座標 */
function applyNew(ss, r){
  if(!r.newName) throw new Error('newName required');
  var co = coordsFromLink(String(r.gmap || ''));
  var sh = ss.getSheetByName(MAIN_SHEET);
  var rowVals = new Array(16).fill('');
  rowVals[C.region-1] = '待分區';
  rowVals[C.en-1] = r.newName;
  rowVals[C.zh-1] = r.newName;
  rowVals[C.note-1] = (r.note || r.newNote || '') + (co ? '' : '（座標待補：' + (r.gmap||'無連結') + '）');
  rowVals[C.date-1] = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd');
  rowVals[C.by-1] = r.nick;
  if(co){ rowVals[C.lat-1] = co[0]; rowVals[C.lon-1] = co[1]; }
  sh.appendRow(rowVals);
  return 'new row | ' + r.newName + (co ? ' @' + co.join(',') : ' (no coords)');
}
function coordsFromLink(u){
  if(!u) return null;
  // 短網址展開
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

/* 健康檢查 */
function doGet(){ return json({ok:true, service:'iceland-parking', time:new Date().toISOString()}); }
