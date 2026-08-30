// ============================================================
// 알림톡 템플릿 관리 - 구글 시트 백엔드 (Apps Script)
// 이 파일은 백업용입니다. 실제 동작하는 코드는
// 구글 스프레드시트 > 확장 프로그램 > Apps Script 안에 있습니다.
// 여기 내용과 실제 배포 코드가 다르다면, 실제 배포 코드가 최신본입니다.
// ============================================================

const SHEET_NAME = 'templates';
const IMAGE_FOLDER_ID = '여기에_구글드라이브_이미지_폴더_ID_입력';
const HEADERS = ['id','code','name','status','messageType','content','buttons','sendTiming','sendTarget','note','hasImage','imageUrl','updatedAt','createdAt','varExample'];

// 저장/삭제/이미지 업로드는 메인(쓰기 허용) 배포에서만 허용하고,
// 그 외 배포(예: 외부 공유용 읽기 전용 배포)에서 호출되면 차단한다.
// Session.getActiveUser().getEmail()은 웹앱 실행 계정 설정에 따라 정상 사용자에게도
// 빈 값을 반환하는 경우가 있어(실제로 겪은 문제 - 메인 배포에서 정상 저장이 막힘) 신뢰할 수 없다.
// 대신 ScriptApp.getService().getUrl()로 "지금 요청이 실제로 어느 배포 URL로 들어왔는지"를 직접 비교한다.
// 아래 값은 메인(쓰기 허용) 배포의 웹 앱 URL로, index.html의 API_BASE_URL과 동일해야 한다.
const MAIN_DEPLOYMENT_URL = 'https://script.google.com/a/macros/meatbox.co.kr/s/AKfycby4KZG_n1Pv_xOjdjLFKKnjqb5Va3ajI0zKaZnkwvmwkq_nxH-HZfGZUMIj5g9zK6Y6TA/exec';

function assertWriteAllowed(){
  let currentUrl = '';
  try{ currentUrl = ScriptApp.getService().getUrl(); }catch(e){ currentUrl = ''; }
  if(currentUrl === MAIN_DEPLOYMENT_URL) return; // 메인 배포에서 온 요청은 항상 허용
  throw new Error('권한이 없습니다 (읽기 전용 접근입니다).');
}

function getSheet(){
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function ensureHeaders(sheet){
  const firstRow = sheet.getRange(1,1,1,HEADERS.length).getValues()[0];
  if(firstRow.join('') === ''){
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  }
}

function doGet(e){
  const action = (e.parameter.action || 'list');
  if(action === 'list') return jsonResponse(listTemplates());
  return jsonResponse({error:'unknown action'});
}

function doPost(e){
  const body = JSON.parse(e.postData.contents);
  try{
    if(body.action === 'save'){ assertWriteAllowed(); return jsonResponse(saveTemplate(body.template)); }
    if(body.action === 'delete'){ assertWriteAllowed(); return jsonResponse(deleteTemplate(body.id)); }
    if(body.action === 'uploadImage'){ assertWriteAllowed(); return jsonResponse(uploadImage(body.id, body.base64, body.mimeType)); }
    if(body.action === 'deleteImage'){ assertWriteAllowed(); return jsonResponse(deleteImageFile(body.id)); }
    return jsonResponse({error:'unknown action'});
  }catch(err){
    return jsonResponse({error:String(err)});
  }
}

function jsonResponse(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function listTemplates(){
  const sheet = getSheet();
  ensureHeaders(sheet);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1).filter(r=>r[0]).map(r=>{
    const obj = {};
    headers.forEach((h,i)=> obj[h] = r[i]);
    obj.buttons = obj.buttons ? JSON.parse(obj.buttons) : [];
    obj.hasImage = obj.hasImage === true || obj.hasImage === 'TRUE';
    return obj;
  });
}

function findRowIndexById(sheet, id){
  const last = sheet.getLastRow();
  if(last < 2) return -1;
  const ids = sheet.getRange(2,1,last-1,1).getValues().flat();
  const idx = ids.indexOf(id);
  return idx === -1 ? -1 : idx + 2;
}

function saveTemplate(t){
  const sheet = getSheet();
  ensureHeaders(sheet);
  if(!t.id) t.id = Utilities.getUuid();
  // 클라이언트(브라우저)가 이미 한국 시간으로 만들어 보낸 값을 우선 사용한다.
  // 혹시 안 보내온 경우에만 서버에서 KST로 채운다 (구버전 클라이언트 호환용).
  if(!t.updatedAt) t.updatedAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  if(!t.createdAt) t.createdAt = t.updatedAt;
  const rowValues = HEADERS.map(h=>{
    if(h === 'buttons') return JSON.stringify(t.buttons||[]);
    if(h === 'hasImage') return t.hasImage ? true : ''; // 이미지 없으면 빈 칸으로 표시 (FALSE 대신)
    return t[h] !== undefined ? t[h] : '';
  });
  const rowIndex = findRowIndexById(sheet, t.id);
  if(rowIndex === -1) sheet.appendRow(rowValues);
  else sheet.getRange(rowIndex,1,1,HEADERS.length).setValues([rowValues]);
  return {ok:true, id:t.id};
}

function deleteTemplate(id){
  const sheet = getSheet();
  const rowIndex = findRowIndexById(sheet, id);
  if(rowIndex !== -1) sheet.deleteRow(rowIndex);
  deleteImageFile(id);
  return {ok:true};
}

function uploadImage(id, base64, mimeType){
  const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
  deleteImageFile(id);
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType || 'image/png', id + '.png');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // uc?export=view 형식은 브라우저에서 <img> 태그로 직접 열면 엑박(깨진 이미지)이 뜨는 경우가 많아
  // 썸네일 엔드포인트로 변경 (핫링크 임베드에 더 안정적으로 동작함)
  const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
  const sheet = getSheet();
  const rowIndex = findRowIndexById(sheet, id);
  if(rowIndex !== -1){
    sheet.getRange(rowIndex, HEADERS.indexOf('imageUrl')+1).setValue(url);
    sheet.getRange(rowIndex, HEADERS.indexOf('hasImage')+1).setValue(true);
  }
  return {ok:true, url:url};
}

function deleteImageFile(id){
  const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
  const files = folder.getFilesByName(id + '.png');
  while(files.hasNext()) files.next().setTrashed(true);
  const sheet = getSheet();
  const rowIndex = findRowIndexById(sheet, id);
  if(rowIndex !== -1){
    sheet.getRange(rowIndex, HEADERS.indexOf('imageUrl')+1).setValue('');
    sheet.getRange(rowIndex, HEADERS.indexOf('hasImage')+1).setValue('');
  }
  return {ok:true};
}
