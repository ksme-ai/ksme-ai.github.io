/**
 * KSME industry collaboration inquiry service.
 * Script Properties: NOTIFY_EMAIL, SPREADSHEET_ID (set by setupInquiry).
 * No email addresses, spreadsheet IDs or secrets belong in public frontend files.
 */
const LEGACY_HEADERS = ['접수일시','접수번호','회사·기관명','담당자 이름','회신 이메일','전화번호','협력 방식','AI 적용 연구·업무 내용','필요한 자문·협력 내용','개인정보 이용 동의','검토 상태','담당자 메모','알림 메일 상태'];
const HEADERS = [...LEGACY_HEADERS.slice(0,6),'문의·신청 유형','문의·신청 내용','추가 요청사항',...LEGACY_HEADERS.slice(9),'희망 일정','희망 장소','참석 대상·예상 인원','희망 교육 시간'];
const VISIT_TYPES = ['방문회의','방문교육','방문 세미나'];
const TAB_NAME = '산학협력 문의';
const TYPES = ['방문회의','방문교육','연구 주제 탐색','기술 자문','방문 세미나','공동연구','기타'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('산학협력 문의·방문 신청')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Run once after setting NOTIFY_EMAIL in Project Settings / Script Properties. */
function setupInquiry() {
  const properties=PropertiesService.getScriptProperties();
  const recipient=properties.getProperty('NOTIFY_EMAIL') || '';
  if(!validEmail_(recipient)) throw new Error('스크립트 속성 NOTIFY_EMAIL에 알림받을 이메일 주소 한 개를 먼저 설정해주세요.');
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try {
    let spreadsheetId=properties.getProperty('SPREADSHEET_ID');
    const book=spreadsheetId?SpreadsheetApp.openById(spreadsheetId):SpreadsheetApp.create('산학협력 탐색그룹 · 기업 문의 접수');
    properties.setProperty('SPREADSHEET_ID',book.getId());
    book.setSpreadsheetTimeZone('Asia/Seoul');
    let sheet=book.getSheetByName(TAB_NAME);
    if(!sheet){
      const existing=book.getSheets();
      sheet=existing.length===1&&existing[0].getLastRow()===0?existing[0].setName(TAB_NAME):book.insertSheet(TAB_NAME);
    }
    if(sheet.getLastRow()===0){
      sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#edf2e7').setFontColor('#233d36');
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1,HEADERS.length,160);
      sheet.setColumnWidths(8,2,360);
      sheet.setColumnWidth(2,300);
      sheet.getRange(1,1,sheet.getMaxRows(),HEADERS.length).setVerticalAlignment('top').setWrap(true);
      sheet.getRange(1,1,sheet.getMaxRows(),HEADERS.length).createFilter();
      const validation=SpreadsheetApp.newDataValidation().requireValueInList(['신규','검토 중','연락 완료','협의 중','종료'],true).setAllowInvalid(false).build();
      sheet.getRange(2,11,sheet.getMaxRows()-1,1).setDataValidation(validation);
    } else if(!headersMatch_(sheet.getRange(1,1,1,HEADERS.length).getValues()[0])) {
      throw new Error('기존 시트의 열 제목이 예상과 다릅니다. 원본을 보존하고 SPREADSHEET_ID 설정을 확인해주세요.');
    }
    MailApp.getRemainingDailyQuota(); // Request only outgoing mail permission, not Gmail inbox access.
    console.log('접수 목록: '+book.getUrl());
    return book.getUrl();
  } finally {lock.releaseLock();}
}

/** Called by google.script.run inside this web app's HTML, never by a public no-cors fetch. */
function submitInquiry(raw) {
  let data;
  try{data=validate_(raw);}catch(error){return {ok:false,message:error.message};}
  const properties=PropertiesService.getScriptProperties();
  const recipient=properties.getProperty('NOTIFY_EMAIL')||'';
  const spreadsheetId=properties.getProperty('SPREADSHEET_ID');
  if(!validEmail_(recipient)||!spreadsheetId)return {ok:false,message:'접수 창구를 준비하고 있습니다. 잠시 후 다시 이용해주세요.'};
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(15000))return {ok:false,message:'다른 문의를 접수하고 있습니다. 잠시 후 다시 시도해주세요.'};
  let stored=false;
  try{
    const book=SpreadsheetApp.openById(spreadsheetId);
    const sheet=book.getSheetByName(TAB_NAME);
    if(!sheet || !headersMatch_(sheet.getRange(1,1,1,HEADERS.length).getValues()[0]))throw new Error('접수 시트 설정 확인 필요');
    const lastRow=sheet.getLastRow();
    const found=lastRow>1?sheet.getRange(2,2,lastRow-1,1).createTextFinder(data.submissionId).matchEntireCell(true).findNext():null;
    if(found)return {ok:true,receipt:data.submissionId};
    const cache=CacheService.getScriptCache();
    const hash=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,data.email.toLowerCase()).map(v=>(v+256).toString(16).slice(-2)).join('');
    const throttle='sender:'+hash;
    if(Number(cache.get(throttle)||0)>=4)return {ok:false,message:'같은 이메일로 문의가 여러 건 접수되었습니다. 잠시 후 다시 이용해주세요.'};
    const row=[new Date(),data.submissionId,data.company,data.name,data.email,data.phone,data.types.join(', '),data.research,data.advice,'동의','신규','','대기',data.preferredSchedule,data.visitLocation,data.audience,data.trainingDuration];
    sheet.appendRow(row.map(safeCell_));
    SpreadsheetApp.flush();stored=true;
    const rowNumber=sheet.getLastRow();
    cache.put(throttle,String(Number(cache.get(throttle)||0)+1),3600);
    try{
      if(MailApp.getRemainingDailyQuota()<1)throw new Error('메일 일일 한도 도달');
      const visitLabel=data.types.filter(t=>VISIT_TYPES.includes(t)).join('·');
      MailApp.sendEmail({
        to:recipient,
        replyTo:data.email,
        name:'산학협력 탐색그룹',
        subject:'['+(visitLabel?visitLabel+' 신청':'산학협력 문의')+'] '+data.company+' · '+data.name,
        body:[
          '새로운 문의·신청이 접수되었습니다.',
          '', '회사·기관: '+data.company,'담당자: '+data.name,'회신 이메일: '+data.email,'전화번호: '+data.phone,
          '문의·신청 유형: '+(data.types.join(', ')||'미선택'),'',
          '[문의·신청 내용]',data.research,'','[추가 요청사항]',data.advice||'미기재','',
          ...(visitLabel?['[방문 신청 정보]','희망 일정: '+(data.preferredSchedule||'미기재'),'희망 장소: '+(data.visitLocation||'미기재'),'참석 대상·예상 인원: '+(data.audience||'미기재'),...(data.types.includes('방문교육')?['희망 교육 시간: '+(data.trainingDuration||'미기재')]:[]),'']:[]),
          '접수번호: '+data.submissionId,
          '접수 목록: '+book.getUrl()+'#gid='+sheet.getSheetId()+'&range=A'+rowNumber,
          '이 메일에 답장하면 문의자의 이메일로 연결됩니다.'
        ].join('\n')
      });
      sheet.getRange(rowNumber,13).setValue('발송 완료');
    }catch(mailError){
      // The inquiry remains accepted even when mail quota or delivery fails.
      try{sheet.getRange(rowNumber,13).setValue('발송 실패 · 시트에서 확인');}catch(ignored){}
      console.error('접수는 저장됐으나 알림 메일 발송 실패');
    }
    return {ok:true,receipt:data.submissionId};
  }catch(error){
    console.error('문의 저장 처리 확인 필요');
    return stored?{ok:true,receipt:data.submissionId}:{ok:false,message:'접수 결과를 확인하지 못했습니다. 같은 내용으로 다시 시도해주세요.'};
  }finally{lock.releaseLock();}
}

// Accept earlier header labels during deployment, without moving existing columns.
function headersMatch_(values){return values.length===HEADERS.length&&values.every((v,i)=>v===HEADERS[i]||(i<LEGACY_HEADERS.length&&v===LEGACY_HEADERS[i]));}

function validEmail_(value){return typeof value==='string'&&value.length<=254&&/^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(value);}
function safeCell_(value){return typeof value==='string'&&/^[=+\-@\t\r\n]/.test(value)?"'"+value:value;}
function validate_(raw){
  if(!raw||typeof raw!=='object'||JSON.stringify(raw).length>16000)throw new Error('입력 내용을 확인해주세요.');
  if(raw.website)throw new Error('접수 요청을 확인하지 못했습니다.');
  const data={};
  [['company',120,true],['name',60,true],['email',254,true],['phone',30,true],['research',3000,true],['advice',3000,false]].forEach(([key,max,required])=>{
    if(typeof raw[key]!=='string')throw new Error('입력 내용을 확인해주세요.');
    data[key]=raw[key].trim();
    if((required&&!data[key])||data[key].length>max||/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(data[key]))throw new Error('필수 항목과 입력 길이를 확인해주세요.');
  });
  if(/[\r\n]/.test(data.company+data.name))throw new Error('회사명과 이름은 한 줄로 입력해주세요.');
  if(!validEmail_(data.email))throw new Error('회신 이메일 주소를 확인해주세요.');
  if(!/^\+?[0-9\s().-]+$/.test(data.phone)||data.phone.replace(/\D/g,'').length<7||data.phone.replace(/\D/g,'').length>15)throw new Error('전화번호를 확인해주세요.');
  if(raw.consent!==true)throw new Error('문의 검토와 회신을 위한 개인정보 이용에 동의해주세요.');
  if(!Array.isArray(raw.types)||raw.types.length>TYPES.length||raw.types.some(t=>!TYPES.includes(t)))throw new Error('문의·신청 유형을 확인해주세요.');
  data.types=[...new Set(raw.types)];
  // New optional fields default to blank for forms opened before this update.
  [['preferredSchedule',200],['visitLocation',200],['audience',200],['trainingDuration',100]].forEach(([key,max])=>{
    const value=raw[key]===undefined?'':raw[key];
    if(typeof value!=='string'||value.length>max||/[\u0000-\u001F]/.test(value))throw new Error('방문 신청 정보의 입력 길이와 내용을 확인해주세요.');
    data[key]=value.trim();
  });
  if(!data.types.some(t=>VISIT_TYPES.includes(t)))data.preferredSchedule=data.visitLocation=data.audience='';
  if(!data.types.includes('방문교육'))data.trainingDuration='';
  if(typeof raw.submissionId!=='string'||!/^[a-f0-9]{8}(-[a-f0-9]{4}){3}-[a-f0-9]{12}$/i.test(raw.submissionId))throw new Error('페이지를 새로 열고 다시 시도해주세요.');
  data.submissionId=raw.submissionId;
  const elapsed=Date.now()-Number(raw.startedAt);
  if(!Number.isFinite(elapsed)||elapsed<2500||elapsed>86400000)throw new Error('입력 내용을 확인한 뒤 잠시 후 다시 제출해주세요.');
  return data;
}
