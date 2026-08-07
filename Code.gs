/**
 * Event Ticketing Platform V5 - Google Apps Script backend
 * --------------------------------------------------------
 * One backend serves:
 *   /event-ticketing-v5/   configuration, admin, vouchers
 *   /event-gate-v5/        gate operations
 *
 * Run in a NEW Google Sheet:
 *   1) setupV5System()
 *   2) bootstrapOwner()
 *   3) createClientConfigurationForm()
 *   4) Deploy as Web App (Execute as Me, Access Anyone)
 */

const TICKETING_SITE = 'https://greenprofessionals.github.io/event-ticketing-v5';
const GATE_SITE = 'https://greenprofessionals.github.io/event-gate-v5';
const CONFIG_URL = TICKETING_SITE + '/config.html';
const ADMIN_URL = TICKETING_SITE + '/admin.html';
const VOUCHER_URL = TICKETING_SITE + '/v.html';
const GATE_URL = GATE_SITE + '/index.html';

const DEFAULT_PRIMARY = '#0B3D24';
const DEFAULT_ACCENT = '#C9A24B';
const DEFAULT_TIER_COLORS = ['#0B3D24','#C9A24B','#164B8C','#8A1C2D','#6B3FA0','#0E7C7B','#D97706','#243B64','#B23A62','#5F6B6D','#3C7A3C','#A66A2C'];
const LIFECYCLE = ['Draft','Client Submitted','Preview Ready','Client Approved','Active','Closed','Archived'];
const ROLES = ['SYSTEM_OWNER','EVENT_ADMIN','FINANCE','GATE_SUPERVISOR','GATE_STAFF'];

const SHEETS = {
  EVENTS:'Events', TIERS:'Tiers', GROUPS:'Groups', ACCESS:'AccessControl',
  VOUCHERS:'Vouchers', CLAIMS:'Claims', CHECKINS:'CheckIns', PAYMENTS:'Payments',
  AUDIT:'AuditLog', COUNTERS:'Counters', FORM:'EventConfigResponses', CONTACTS:'ContactLog'
};

const HEADERS = {
  Events:['EventID','ClientName','ClientEmail','ClientPhone','OrgName','ChapterName','EventTitle','Tagline','EventDate','EventTime','VenueName','VenueAddress','ContactPhone','ContactEmail','WebsiteURL','DressCode','PrimaryColor','AccentColor','LogoFileId','BackgroundFileId','GroupLabel','UseGroups','SerialPrefix','CurrencySymbol','FooterLegalText','Capacity','Status','ConfigToken','ApprovedAt','ApprovedBy','CreatedAt','UpdatedAt'],
  Tiers:['EventID','TierKey','Label','Price','Capacity','Active','Color','SortOrder'],
  Groups:['EventID','GroupName','Active','SortOrder'],
  AccessControl:['UserID','Name','Role','PasscodeHash','EventScope','Email','Phone','Active','CreatedAt'],
  Vouchers:['Timestamp','EventID','BatchID','VoucherToken','TierKey','SuggestedGroup','PrefillName','PrefillPhone','Claimed','Serial','Dispatched','RecipientName','RecipientEmail','RecipientPhone','SentAt','IssuedBy','Status'],
  Claims:['Timestamp','EventID','Serial','CheckInToken','Name','Email','Phone','GroupName','TierKey','Source','VoucherToken','Status','AmountDue','AmountPaid','PaymentStatus','PaymentMethod','PaymentNote','TransferredFrom','UpdatedAt'],
  CheckIns:['Timestamp','EventID','Serial','Name','GroupName','TierKey','Phone','PaymentStatus','PaymentMethod','AmountPaid','CheckedInBy','GateNote','Status'],
  Payments:['Timestamp','EventID','Serial','Amount','Method','Status','Note','RecordedBy'],
  AuditLog:['Timestamp','EventID','Action','EntityType','EntityID','Actor','Role','Details'],
  Counters:['EventID','CurrentNumber'],
  EventConfigResponses:[],
  ContactLog:['Timestamp','EventID','EntityType','EntityID','Channel','Recipient','Actor','Note']
};

function setupV5System() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(name=>{ if(name==='EventConfigResponses') return; ensureSheet_(ss,SHEETS[name.toUpperCase()]||name,HEADERS[name]); });
  ensureSheet_(ss,SHEETS.EVENTS,HEADERS.Events);
  ensureSheet_(ss,SHEETS.TIERS,HEADERS.Tiers);
  ensureSheet_(ss,SHEETS.GROUPS,HEADERS.Groups);
  ensureSheet_(ss,SHEETS.ACCESS,HEADERS.AccessControl);
  ensureSheet_(ss,SHEETS.VOUCHERS,HEADERS.Vouchers);
  ensureSheet_(ss,SHEETS.CLAIMS,HEADERS.Claims);
  ensureSheet_(ss,SHEETS.CHECKINS,HEADERS.CheckIns);
  ensureSheet_(ss,SHEETS.PAYMENTS,HEADERS.Payments);
  ensureSheet_(ss,SHEETS.AUDIT,HEADERS.AuditLog);
  ensureSheet_(ss,SHEETS.COUNTERS,HEADERS.Counters);
  ensureSheet_(ss,SHEETS.CONTACTS,HEADERS.ContactLog);
  SpreadsheetApp.getUi().alert('V5 sheets created. Next run bootstrapOwner(), then createClientConfigurationForm().');
}

function bootstrapOwner() {
  const ui=SpreadsheetApp.getUi();
  const nameR=ui.prompt('System Owner','Owner name:',ui.ButtonSet.OK_CANCEL); if(nameR.getSelectedButton()!==ui.Button.OK)return;
  const passR=ui.prompt('System Owner','Choose a strong passcode:',ui.ButtonSet.OK_CANCEL); if(passR.getSelectedButton()!==ui.Button.OK)return;
  const pass=passR.getResponseText().trim(); if(pass.length<8){ui.alert('Use at least 8 characters.');return;}
  const emailR=ui.prompt('System Owner','Email (optional):',ui.ButtonSet.OK_CANCEL); if(emailR.getSelectedButton()!==ui.Button.OK)return;
  const phoneR=ui.prompt('System Owner','Phone (optional):',ui.ButtonSet.OK_CANCEL); if(phoneR.getSelectedButton()!==ui.Button.OK)return;
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ACCESS);
  const id='OWNER-'+Utilities.getUuid().slice(0,8).toUpperCase();
  sh.appendRow([id,nameR.getResponseText().trim(),'SYSTEM_OWNER',hash_(pass),'*',emailR.getResponseText().trim(),phoneR.getResponseText().trim(),true,new Date()]);
  ui.alert('Owner created','Use the passcode you just entered to unlock Admin and Gate Supervisor functions.',ui.ButtonSet.OK);
}

function addAccessUser() {
  const ui=SpreadsheetApp.getUi();
  const name=ui.prompt('Add Access User','Name:',ui.ButtonSet.OK_CANCEL); if(name.getSelectedButton()!==ui.Button.OK)return;
  const role=ui.prompt('Add Access User','Role: '+ROLES.join(', '),ui.ButtonSet.OK_CANCEL); if(role.getSelectedButton()!==ui.Button.OK)return;
  const r=role.getResponseText().trim().toUpperCase(); if(!ROLES.includes(r)){ui.alert('Invalid role.');return;}
  const scope=ui.prompt('Add Access User','Event scope: * for all, or comma-separated Event IDs',ui.ButtonSet.OK_CANCEL); if(scope.getSelectedButton()!==ui.Button.OK)return;
  const pass=ui.prompt('Add Access User','Passcode (8+ characters):',ui.ButtonSet.OK_CANCEL); if(pass.getSelectedButton()!==ui.Button.OK)return;
  if(pass.getResponseText().trim().length<8){ui.alert('Use at least 8 characters.');return;}
  const email=ui.prompt('Add Access User','Email (optional):',ui.ButtonSet.OK_CANCEL); if(email.getSelectedButton()!==ui.Button.OK)return;
  const phone=ui.prompt('Add Access User','Phone (optional):',ui.ButtonSet.OK_CANCEL); if(phone.getSelectedButton()!==ui.Button.OK)return;
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ACCESS);
  sh.appendRow(['USR-'+Utilities.getUuid().slice(0,8).toUpperCase(),name.getResponseText().trim(),r,hash_(pass.getResponseText().trim()),scope.getResponseText().trim()||'*',email.getResponseText().trim(),phone.getResponseText().trim(),true,new Date()]);
  ui.alert('Access user added.');
}

function createClientConfigurationForm() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const old=ss.getSheetByName(SHEETS.FORM);
  if(old){const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'America/New_York','yyyyMMdd_HHmmss');old.setName('EventConfigResponses_Legacy_'+stamp);}
  const before=ss.getSheets().map(s=>s.getName());
  const form=FormApp.create('Event Ticket Configuration');
  form.setDescription('Complete the event information. The Event ID and configuration key are prefilled from your private client link. Do not change them.');
  form.setConfirmationMessage('Configuration submitted. Return to the client portal, load the ticket preview, revise if needed, then approve the design.');
  const text=(t,h,req)=>{const i=form.addTextItem().setTitle(t).setHelpText(h||'');if(req)i.setRequired(true);return i;};
  const para=(t,h)=>form.addParagraphTextItem().setTitle(t).setHelpText(h||'');
  text('EventID','Prefilled. Do not change.',true);
  text('ConfigToken','Prefilled security key. Do not change.',true);
  text('OrgName','Organization name displayed on ticket',true);
  text('ChapterName','Chapter/unit/organizing body displayed on ticket',true);
  text('EventTitle','Public event title',true);
  para('Tagline','Optional subtitle');
  text('EventDate','Example: Saturday, September 5, 2026',true);
  text('EventTime','Example: 8:00 PM',true);
  text('VenueName','Venue name',true);
  text('VenueAddress','Full address',true);
  text('ContactPhone','Public event contact phone');
  text('ContactEmail','Public event contact email');
  text('WebsiteURL','Public website');
  text('DressCode','Optional');
  text('PrimaryColor','Optional HEX, e.g. #0B3D24');
  text('AccentColor','Optional HEX, e.g. #C9A24B');
  text('SerialPrefix','Short prefix, e.g. NY-');
  text('CurrencySymbol','Example: $');
  text('FooterLegalText','Optional small print');
  text('Capacity','Optional total capacity');
  form.addMultipleChoiceItem().setTitle('UseGroups').setChoiceValues(['true','false']).setRequired(true);
  text('GroupLabel','Example: Chapter, Team, Table');
  para('GroupsList','One group/chapter per line');
  for(let i=1;i<=12;i++){
    form.addSectionHeaderItem().setTitle('Ticket Tier '+i);
    text('Tier'+i+'Name','Leave blank to skip');
    text('Tier'+i+'Price','Numbers only');
    text('Tier'+i+'Capacity','Optional');
    text('Tier'+i+'Color','Optional HEX; blank uses a system default');
  }
  form.setDestination(FormApp.DestinationType.SPREADSHEET,ss.getId());
  const created=ss.getSheets().filter(s=>!before.includes(s.getName())); if(created.length)created[0].setName(SHEETS.FORM);
  PropertiesService.getScriptProperties().setProperty('V5_CONFIG_FORM_ID',form.getId());
  PropertiesService.getScriptProperties().setProperty('V5_CONFIG_FORM_URL',form.getPublishedUrl());
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='handleConfigFormSubmit_').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('handleConfigFormSubmit_').forSpreadsheet(ss).onFormSubmit().create();
  SpreadsheetApp.getUi().alert('Client configuration form created','Editor:\n'+form.getEditUrl()+'\n\nPublished:\n'+form.getPublishedUrl(),SpreadsheetApp.getUi().ButtonSet.OK);
}


function handleConfigFormSubmit_(e){
  try{
    const nv=e&&e.namedValues||{};const eventId=normId_((nv.EventID||[''])[0]),key=String((nv.ConfigToken||[''])[0]||'').trim();
    const r=getEventBase_(eventId);if(!r||String(r.ConfigToken)!==key)return;
    r.Status='Client Submitted';r.ApprovedAt='';r.ApprovedBy='';r.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,r._row,r);
    audit_(eventId,'CLIENT_SUBMITTED','Event',eventId,r.ClientName||'Client','CLIENT',{source:'Google Form'});
  }catch(err){console.log(err);}
}

// ---------- Core helpers ----------
function ensureSheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.appendRow(headers);else ensureHeaders_(sh,headers);sh.setFrozenRows(1);return sh;}
function ensureHeaders_(sh,headers){const last=Math.max(sh.getLastColumn(),1);const existing=sh.getRange(1,1,1,last).getValues()[0].map(String);headers.forEach(h=>{if(!existing.includes(h)){sh.getRange(1,sh.getLastColumn()+1).setValue(h);existing.push(h);}});}
function rows_(name){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);if(!sh||sh.getLastRow()<2)return[];const data=sh.getDataRange().getValues();const h=data.shift().map(String);return data.map((r,i)=>{const o={_row:i+2};h.forEach((k,j)=>o[k]=r[j]);return o;});}
function rowBy_(name,field,value){return rows_(name).find(r=>String(r[field])===String(value))||null;}
function writeRow_(name,headers,rowNum,obj){SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getRange(rowNum,1,1,headers.length).setValues([headers.map(h=>obj[h]===undefined?'':obj[h])]);}
function append_(name,headers,obj){SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).appendRow(headers.map(h=>obj[h]===undefined?'':obj[h]));}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function body_(e){try{return JSON.parse(e&&e.postData&&e.postData.contents||'{}')}catch(_){return{}}}
function normId_(v){return String(v||'').trim().toUpperCase().replace(/[^A-Z0-9_-]/g,'');}
function bool_(v){return v===true||v===1||String(v).toLowerCase()==='true';}
function money_(v){const n=Number(v);return isFinite(n)?n:0;}
function digits_(v){return String(v||'').replace(/\D/g,'');}
function token_(){return Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');}
function hash_(s){return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s||''),Utilities.Charset.UTF_8));}
function validHex_(v){const s=String(v||'').trim();return /^#[0-9a-f]{6}$/i.test(s)?s.toUpperCase():'';}
function tierColor_(v,i){return validHex_(v)||DEFAULT_TIER_COLORS[i%DEFAULT_TIER_COLORS.length];}
function cleanGroup_(v){const s=String(v==null?'':v).trim();return !s||/^(undefined|null|nan)$/i.test(s)?'':s;}
function splitGroups_(v){if(Array.isArray(v))return v.map(cleanGroup_).filter(Boolean);const s=String(v==null?'':v).trim();if(!s)return[];return s.split(/\r?\n|\s*;\s*|\s*,\s*/).map(cleanGroup_).filter(Boolean);}
function normalizeDate_(v){if(v instanceof Date&&!isNaN(v))return Utilities.formatDate(v,Session.getScriptTimeZone()||'America/New_York','EEEE, MMMM d, yyyy');const s=String(v||'').trim();if(!s)return'';const d=new Date(s);if(!isNaN(d)&&/(GMT|T\d\d:)/.test(s))return Utilities.formatDate(d,Session.getScriptTimeZone()||'America/New_York','EEEE, MMMM d, yyyy');return s;}
function normalizeTime_(v){if(v instanceof Date&&!isNaN(v))return Utilities.formatDate(v,Session.getScriptTimeZone()||'America/New_York','h:mm a');const s=String(v||'').trim();if(!s)return'';if(/^\d{1,2}:\d{2}(\s*[AP]M)?$/i.test(s)){if(/[AP]M/i.test(s))return s.toUpperCase();const p=s.split(':');let h=Number(p[0]);const ap=h>=12?'PM':'AM';h=h%12||12;return h+':'+p[1]+' '+ap;}const d=new Date(s);return !isNaN(d)?Utilities.formatDate(d,Session.getScriptTimeZone()||'America/New_York','h:mm a'):s;}
function dateObj_(v){if(v instanceof Date&&!isNaN(v))return new Date(v.getFullYear(),v.getMonth(),v.getDate());const d=new Date(String(v||''));return isNaN(d)?null:new Date(d.getFullYear(),d.getMonth(),d.getDate());}
function audit_(eventId,action,type,id,actor,role,details){append_(SHEETS.AUDIT,HEADERS.AuditLog,{Timestamp:new Date(),EventID:eventId,Action:action,EntityType:type,EntityID:id,Actor:actor||'',Role:role||'',Details:typeof details==='string'?details:JSON.stringify(details||{})});}
function contactLog_(eventId,type,id,channel,recipient,actor,note){append_(SHEETS.CONTACTS,HEADERS.ContactLog,{Timestamp:new Date(),EventID:eventId,EntityType:type,EntityID:id,Channel:channel,Recipient:recipient,Actor:actor||'',Note:note||''});}

// ---------- Authentication / authorization ----------
function authenticate_(pass){const h=hash_(String(pass||'').trim());const u=rows_(SHEETS.ACCESS).find(r=>bool_(r.Active)&&r.PasscodeHash===h);if(!u)return null;return{userId:u.UserID,name:u.Name,role:String(u.Role),scope:String(u.EventScope||'*'),email:u.Email||'',phone:u.Phone||''};}
function scopeAllows_(u,eventId){if(!u)return false;if(u.role==='SYSTEM_OWNER'||u.scope==='*')return true;return String(u.scope||'').split(',').map(normId_).includes(normId_(eventId));}
function requireRole_(body,roles,eventId){const u=authenticate_(body.adminPasscode);if(!u)return{ok:false,error:'Incorrect passcode.'};if(!roles.includes(u.role)&&u.role!=='SYSTEM_OWNER')return{ok:false,error:'Your role does not permit this action.'};if(eventId&&!scopeAllows_(u,eventId))return{ok:false,error:'You are not assigned to this event.'};return{ok:true,user:u};}
function roleCanGate_(role){return['SYSTEM_OWNER','EVENT_ADMIN','GATE_SUPERVISOR','GATE_STAFF'].includes(role);}

// ---------- Event + form overlay ----------
function latestValidForm_(eventId,configToken){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.FORM);if(!sh||sh.getLastRow()<2)return null;const data=sh.getDataRange().getValues();const h=data[0].map(String),ei=h.indexOf('EventID'),ki=h.indexOf('ConfigToken');if(ei<0||ki<0)return null;for(let r=data.length-1;r>=1;r--){if(normId_(data[r][ei])===normId_(eventId)&&String(data[r][ki]).trim()===String(configToken||'').trim()){const o={};h.forEach((k,i)=>o[k]=data[r][i]);return o;}}return null;}
function getEventBase_(eventId){return rowBy_(SHEETS.EVENTS,'EventID',normId_(eventId));}
function getEvent_(eventId){const base=getEventBase_(eventId);if(!base)return null;const e=Object.assign({},base);const f=latestValidForm_(e.EventID,e.ConfigToken);if(f){['OrgName','ChapterName','EventTitle','Tagline','EventDate','EventTime','VenueName','VenueAddress','ContactPhone','ContactEmail','WebsiteURL','DressCode','PrimaryColor','AccentColor','GroupLabel','UseGroups','SerialPrefix','CurrencySymbol','FooterLegalText','Capacity'].forEach(k=>{if(f[k]!==undefined&&String(f[k]).trim()!=='')e[k]=f[k];});}e.PrimaryColor=validHex_(e.PrimaryColor)||DEFAULT_PRIMARY;e.AccentColor=validHex_(e.AccentColor)||DEFAULT_ACCENT;e.EventDate=normalizeDate_(e.EventDate);e.EventTime=normalizeTime_(e.EventTime);e.Capacity=Number(e.Capacity)||0;e.UseGroups=String(e.UseGroups||'true');e.LogoURL=e.LogoFileId?driveData_(e.LogoFileId):'';e.BackgroundURL=e.BackgroundFileId?driveData_(e.BackgroundFileId):'';return e;}
function publicEvent_(e){if(!e)return null;const keys=['EventID','OrgName','ChapterName','EventTitle','Tagline','EventDate','EventTime','VenueName','VenueAddress','ContactPhone','ContactEmail','WebsiteURL','DressCode','PrimaryColor','AccentColor','GroupLabel','UseGroups','SerialPrefix','CurrencySymbol','FooterLegalText','Capacity','Status','LogoURL','BackgroundURL'];const o={};keys.forEach(k=>o[k]=e[k]);return o;}
function getTiers_(eventId){const e=getEventBase_(eventId);if(!e)return{};const f=latestValidForm_(e.EventID,e.ConfigToken);if(f){const out={};let n=0;for(let i=1;i<=12;i++){const label=String(f['Tier'+i+'Name']||'').trim();if(!label)continue;const key=(label.toLowerCase().replace(/[^a-z0-9]+/g,'').slice(0,30)||('tier'+i));out[key]={label:label,price:money_(f['Tier'+i+'Price']),capacity:Number(f['Tier'+i+'Capacity'])||0,color:tierColor_(f['Tier'+i+'Color'],n++)};}if(Object.keys(out).length)return out;}
  const out={};rows_(SHEETS.TIERS).filter(r=>normId_(r.EventID)===normId_(eventId)&&String(r.Active).toLowerCase()!=='false').sort((a,b)=>(Number(a.SortOrder)||0)-(Number(b.SortOrder)||0)).forEach((r,i)=>out[r.TierKey]={label:r.Label,price:money_(r.Price),capacity:Number(r.Capacity)||0,color:tierColor_(r.Color,i)});return out;}
function getGroups_(eventId){const e=getEventBase_(eventId);if(!e)return[];const f=latestValidForm_(e.EventID,e.ConfigToken);const fg=f?splitGroups_(f.GroupsList):[];if(fg.length)return[...new Set(fg)];return[...new Set(rows_(SHEETS.GROUPS).filter(r=>normId_(r.EventID)===normId_(eventId)&&String(r.Active).toLowerCase()!=='false').sort((a,b)=>(Number(a.SortOrder)||0)-(Number(b.SortOrder)||0)).map(r=>cleanGroup_(r.GroupName)).filter(Boolean))];}
function driveData_(id){try{const b=DriveApp.getFileById(String(id)).getBlob();return'data:'+b.getContentType()+';base64,'+Utilities.base64Encode(b.getBytes())}catch(_){return''}}
function assetFolder_(){const n='Event Ticketing V5 Assets',it=DriveApp.getFoldersByName(n);return it.hasNext()?it.next():DriveApp.createFolder(n);}
function formPrefillUrl_(eventId,token){const id=PropertiesService.getScriptProperties().getProperty('V5_CONFIG_FORM_ID');if(!id)return'';const form=FormApp.openById(id);const resp=form.createResponse();form.getItems().forEach(item=>{if(item.getTitle()==='EventID')resp.withItemResponse(item.asTextItem().createResponse(eventId));if(item.getTitle()==='ConfigToken')resp.withItemResponse(item.asTextItem().createResponse(token));});return resp.toPrefilledUrl();}

function eventSummary_(e){return{EventID:e.EventID,ClientName:e.ClientName||'',ClientEmail:e.ClientEmail||'',ClientPhone:e.ClientPhone||'',OrgName:e.OrgName||'',ChapterName:e.ChapterName||'',EventTitle:e.EventTitle||e.EventID,EventDate:e.EventDate||'',VenueName:e.VenueName||'',Status:e.Status||'Draft',Capacity:Number(e.Capacity)||0,ApprovedAt:e.ApprovedAt||'',PrimaryColor:e.PrimaryColor||DEFAULT_PRIMARY,AccentColor:e.AccentColor||DEFAULT_ACCENT};}
function currentFutureEvents_(u,includePast){const today=new Date();today.setHours(0,0,0,0);return rows_(SHEETS.EVENTS).map(r=>getEvent_(r.EventID)).filter(Boolean).filter(e=>scopeAllows_(u,e.EventID)).filter(e=>{if(includePast)return true;if(String(e.Status)==='Archived')return false;const d=dateObj_(e.EventDate);return !d||d>=today;}).sort((a,b)=>{const da=dateObj_(a.EventDate),db=dateObj_(b.EventDate);if(!da&&!db)return String(a.EventTitle).localeCompare(String(b.EventTitle));if(!da)return 1;if(!db)return-1;return da-db;}).map(eventSummary_);}

// ---------- Client configuration ----------
function createEventShell_(body,u){const eventId=normId_(body.eventId);if(!eventId)return{ok:false,error:'Event ID is required.'};if(getEventBase_(eventId))return{ok:false,error:'That Event ID already exists.'};const now=new Date(),tok=token_();append_(SHEETS.EVENTS,HEADERS.Events,{EventID:eventId,ClientName:String(body.clientName||''),ClientEmail:String(body.clientEmail||''),ClientPhone:String(body.clientPhone||''),EventTitle:String(body.eventTitle||eventId),PrimaryColor:DEFAULT_PRIMARY,AccentColor:DEFAULT_ACCENT,UseGroups:'true',GroupLabel:'Chapter',CurrencySymbol:'$',Status:'Draft',ConfigToken:tok,CreatedAt:now,UpdatedAt:now});append_(SHEETS.COUNTERS,HEADERS.Counters,{EventID:eventId,CurrentNumber:0});if(u.role==='EVENT_ADMIN'&&u.scope!=='*'){const ar=rowBy_(SHEETS.ACCESS,'UserID',u.userId);if(ar){const ids=String(ar.EventScope||'').split(',').map(normId_).filter(Boolean);if(!ids.includes(eventId))ids.push(eventId);ar.EventScope=ids.join(',');writeRow_(SHEETS.ACCESS,HEADERS.AccessControl,ar._row,ar);}}audit_(eventId,'CREATE_EVENT_SHELL','Event',eventId,u.name,u.role,{clientName:body.clientName||''});return{ok:true,eventId:eventId,configUrl:CONFIG_URL+'?event='+encodeURIComponent(eventId)+'&key='+encodeURIComponent(tok),formUrl:formPrefillUrl_(eventId,tok)};}
function clientContext_(eventId,key){const b=getEventBase_(eventId);if(!b||String(b.ConfigToken)!==String(key||''))return{ok:false,error:'This configuration link is invalid or expired.'};const f=latestValidForm_(b.EventID,b.ConfigToken);if(f){const submittedAt=f.Timestamp instanceof Date?f.Timestamp:new Date(f.Timestamp||0);const approvedAt=b.ApprovedAt instanceof Date?b.ApprovedAt:new Date(b.ApprovedAt||0);if(b.Status==='Draft'||((b.Status==='Client Approved'||b.Status==='Active')&&submittedAt>approvedAt)){b.Status='Client Submitted';b.ApprovedAt='';b.ApprovedBy='';b.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,b._row,b);audit_(b.EventID,'CLIENT_SUBMITTED','Event',b.EventID,b.ClientName||'Client','CLIENT',{});}}const e=getEvent_(eventId);return{ok:true,event:publicEvent_(e),tiers:getTiers_(eventId),groups:getGroups_(eventId),formUrl:formPrefillUrl_(b.EventID,b.ConfigToken),submitted:!!f,status:b.Status,client:{name:b.ClientName||'',email:b.ClientEmail||'',phone:b.ClientPhone||''},hasLogo:!!b.LogoFileId,hasBackground:!!b.BackgroundFileId};}
function clientUpload_(body){const b=getEventBase_(body.eventId);if(!b||String(b.ConfigToken)!==String(body.configKey||''))return{ok:false,error:'Invalid configuration link.'};const type=body.assetType==='background'?'background':'logo';const m=String(body.dataUrl||'').match(/^data:(image\/(?:png|jpeg|jpg|webp));base64,(.+)$/i);if(!m)return{ok:false,error:'Use PNG, JPG, or WebP.'};const bytes=Utilities.base64Decode(m[2]);if(bytes.length>3*1024*1024)return{ok:false,error:'Image must be 3 MB or smaller.'};const ext=m[1].includes('png')?'png':m[1].includes('webp')?'webp':'jpg';const f=assetFolder_().createFile(Utilities.newBlob(bytes,m[1],b.EventID+'-'+type+'.'+ext));b[type==='logo'?'LogoFileId':'BackgroundFileId']=f.getId();b.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,b._row,b);audit_(b.EventID,'CLIENT_UPLOAD_ASSET','Event',b.EventID,b.ClientName||'Client','CLIENT',type);return{ok:true};}
function clientPreview_(body){const b=getEventBase_(body.eventId);if(!b||String(b.ConfigToken)!==String(body.configKey||''))return{ok:false,error:'Invalid configuration link.'};const f=latestValidForm_(b.EventID,b.ConfigToken);if(!f)return{ok:false,error:'Submit the Google configuration form first.'};if(b.Status==='Draft'||b.Status==='Client Submitted')b.Status='Preview Ready';b.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,b._row,b);audit_(b.EventID,'CLIENT_PREVIEW','Event',b.EventID,b.ClientName||'Client','CLIENT',{});return{ok:true,event:publicEvent_(getEvent_(b.EventID)),tiers:getTiers_(b.EventID),groups:getGroups_(b.EventID)};}
function clientApprove_(body){const b=getEventBase_(body.eventId);if(!b||String(b.ConfigToken)!==String(body.configKey||''))return{ok:false,error:'Invalid configuration link.'};if(!latestValidForm_(b.EventID,b.ConfigToken))return{ok:false,error:'Configuration form has not been submitted.'};b.Status='Client Approved';b.ApprovedAt=new Date();b.ApprovedBy=b.ClientName||'Client';b.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,b._row,b);audit_(b.EventID,'CLIENT_APPROVE','Event',b.EventID,b.ApprovedBy,'CLIENT',{});return{ok:true,status:b.Status};}

// ---------- Capacity / tickets / vouchers ----------
function allocations_(eventId){const claims=rows_(SHEETS.CLAIMS).filter(r=>normId_(r.EventID)===normId_(eventId)&&String(r.Status)!=='Revoked');const vouchers=rows_(SHEETS.VOUCHERS).filter(r=>normId_(r.EventID)===normId_(eventId)&&!bool_(r.Claimed)&&String(r.Status)!=='Cancelled');const byTier={};claims.forEach(r=>byTier[r.TierKey]=(byTier[r.TierKey]||0)+1);vouchers.forEach(r=>{if(r.TierKey)byTier[r.TierKey]=(byTier[r.TierKey]||0)+1});return{total:claims.length+vouchers.length,byTier:byTier,claims:claims.length,pendingVouchers:vouchers.length};}
function capacity_(eventId,adds,allowOverride){const e=getEvent_(eventId);if(!e)return{ok:false,error:'Event not found.'};if(e.Status!=='Active')return{ok:false,error:'Voucher generation and ticket claiming require an Active event.'};const a=allocations_(eventId),sum=Object.values(adds||{}).reduce((x,y)=>x+Number(y||0),0);if(!allowOverride&&e.Capacity&&a.total+sum>e.Capacity)return{ok:false,error:'Event capacity exceeded. '+Math.max(0,e.Capacity-a.total)+' place(s) remain.'};const ts=getTiers_(eventId);for(const k in adds){const n=Number(adds[k])||0,cap=ts[k]&&Number(ts[k].capacity)||0;if(!allowOverride&&cap&&(a.byTier[k]||0)+n>cap)return{ok:false,error:(ts[k]?.label||k)+' capacity exceeded.'};}return{ok:true};}
function nextSerial_(eventId){const lock=LockService.getScriptLock();lock.waitLock(15000);try{let r=rowBy_(SHEETS.COUNTERS,'EventID',eventId);if(!r){append_(SHEETS.COUNTERS,HEADERS.Counters,{EventID:eventId,CurrentNumber:0});r=rowBy_(SHEETS.COUNTERS,'EventID',eventId);}const n=(Number(r.CurrentNumber)||0)+1;r.CurrentNumber=n;writeRow_(SHEETS.COUNTERS,HEADERS.Counters,r._row,r);const p=getEvent_(eventId).SerialPrefix||eventId+'-';return p+String(n).padStart(3,'0');}finally{lock.releaseLock();}}
function generateVouchers_(body,u){const eventId=normId_(body.eventId);const auth=scopeAllows_(u,eventId);if(!auth)return{ok:false,error:'Not assigned to event.'};const tiers=getTiers_(eventId),counts=body.tierCounts||{};let total=Math.max(0,Number(body.openCount)||0),adds={};Object.keys(tiers).forEach(k=>{const n=Math.max(0,Number(counts[k])||0);adds[k]=n;total+=n});if(total<1||total>100)return{ok:false,error:'Generate between 1 and 100 vouchers per batch.'};const cap=capacity_(eventId,adds,bool_(body.capacityOverride)&&['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role));if(!cap.ok)return cap;const batch=token_().slice(0,32),now=new Date(),group=cleanGroup_(body.suggestedGroup),recipientName=String(body.distributorName||'').trim(),recipientEmail=String(body.distributorEmail||'').trim(),recipientPhone=String(body.distributorPhone||'').trim();const rows=[];Object.keys(adds).forEach(k=>{for(let i=0;i<adds[k];i++)rows.push(k)});for(let i=0;i<Math.max(0,Number(body.openCount)||0);i++)rows.push('');rows.forEach(k=>append_(SHEETS.VOUCHERS,HEADERS.Vouchers,{Timestamp:now,EventID:eventId,BatchID:batch,VoucherToken:token_(),TierKey:k,SuggestedGroup:group,PrefillName:'',PrefillPhone:'',Claimed:false,Dispatched:false,RecipientName:'',RecipientEmail:'',RecipientPhone:'',IssuedBy:u.name,Status:'Available'}));audit_(eventId,'GENERATE_VOUCHERS','Batch',batch,u.name,u.role,{count:total,group:group,distributorName:recipientName});return{ok:true,batchId:batch,count:total,distributorUrl:VOUCHER_URL+'?batch='+encodeURIComponent(batch),distributorName:recipientName,distributorEmail:recipientEmail,distributorPhone:recipientPhone};}
function batch_(id){const rs=rows_(SHEETS.VOUCHERS).filter(r=>r.BatchID===String(id));if(!rs.length)return{ok:false,error:'Batch not found.'};const eventId=normId_(rs[0].EventID);return{ok:true,eventId:eventId,event:publicEvent_(getEvent_(eventId)),tiers:getTiers_(eventId),vouchers:rs.map(r=>({token:r.VoucherToken,tier:r.TierKey,claimed:bool_(r.Claimed),serial:r.Serial||'',dispatched:bool_(r.Dispatched),status:r.Status||'Available',suggestedGroup:cleanGroup_(r.SuggestedGroup)}))};}
function voucher_(tok){const r=rowBy_(SHEETS.VOUCHERS,'VoucherToken',String(tok||''));if(!r)return{ok:false,error:'Voucher not found.'};const eventId=normId_(r.EventID);if(String(r.Status)==='Cancelled')return{ok:false,error:'This voucher has been cancelled.'};if(bool_(r.Claimed)){const c=rowBy_(SHEETS.CLAIMS,'Serial',r.Serial);return{ok:false,alreadyClaimed:true,eventId:eventId,serial:r.Serial,tier:c&&c.TierKey,checkInToken:c&&c.CheckInToken,event:publicEvent_(getEvent_(eventId)),tiers:getTiers_(eventId)};}return{ok:true,eventId:eventId,event:publicEvent_(getEvent_(eventId)),tiers:getTiers_(eventId),groups:getGroups_(eventId),tier:r.TierKey||'',groupName:cleanGroup_(r.SuggestedGroup),name:r.PrefillName||'',phone:r.PrefillPhone||''};}
function sendVoucher_(body){const r=rowBy_(SHEETS.VOUCHERS,'VoucherToken',String(body.token||''));if(!r)return{ok:false,error:'Voucher not found.'};if(bool_(r.Claimed)||bool_(r.Dispatched)||String(r.Status)==='Cancelled')return{ok:false,error:'Voucher is no longer available to send.'};const email=String(body.recipientEmail||'').trim(),phone=String(body.recipientPhone||'').trim(),name=String(body.recipientName||'').trim();if(!email&&!phone)return{ok:false,error:'Enter an email or phone number.'};const url=VOUCHER_URL+'?voucher='+encodeURIComponent(r.VoucherToken);if(email){try{MailApp.sendEmail(email,'Your event ticket voucher','Your voucher is ready. Claim your ticket here: '+url)}catch(e){return{ok:false,error:'Email could not be sent: '+e.message}}}r.Dispatched=true;r.RecipientName=name;r.RecipientEmail=email;r.RecipientPhone=phone;r.SentAt=new Date();r.Status='Sent';writeRow_(SHEETS.VOUCHERS,HEADERS.Vouchers,r._row,r);contactLog_(r.EventID,'Voucher',r.VoucherToken,email?'Email':'Phone',email||phone,'Distributor','Voucher delivery');return{ok:true,url:url};}
function claim_(body){const tok=String(body.voucher||'').trim();if(!tok)return{ok:false,error:'A valid voucher is required.'};const v=rowBy_(SHEETS.VOUCHERS,'VoucherToken',tok);if(!v)return{ok:false,error:'Voucher not found.'};if(String(v.Status)==='Cancelled')return{ok:false,error:'Voucher cancelled.'};if(bool_(v.Claimed)){const c=rowBy_(SHEETS.CLAIMS,'Serial',v.Serial);return{ok:false,alreadyClaimed:true,eventId:v.EventID,serial:v.Serial,tier:c&&c.TierKey,checkInToken:c&&c.CheckInToken};}const eventId=normId_(v.EventID),e=getEvent_(eventId);if(!e||e.Status!=='Active')return{ok:false,error:'This event is not currently accepting ticket claims.'};const name=String(body.fullName||'').trim();if(!name)return{ok:false,error:'Name is required.'};const tiers=getTiers_(eventId);let tier=v.TierKey||String(body.tier||'');if(!tier)tier=Object.keys(tiers)[0]||'';if(!tiers[tier])return{ok:false,error:'Invalid ticket tier.'};const groups=getGroups_(eventId),group=cleanGroup_(body.groupName);if(String(e.UseGroups).toLowerCase()!=='false'&&groups.length&&!groups.includes(group))return{ok:false,error:'Select a valid '+(e.GroupLabel||'group')+'.'};if(!v.TierKey){const cap=capacity_(eventId,{[tier]:1},false);if(!cap.ok)return cap;}const serial=nextSerial_(eventId),qr=token_().slice(0,48),price=money_(tiers[tier].price),now=new Date();append_(SHEETS.CLAIMS,HEADERS.Claims,{Timestamp:now,EventID:eventId,Serial:serial,CheckInToken:qr,Name:name,Email:String(body.email||''),Phone:String(body.phone||''),GroupName:group,TierKey:tier,Source:'Voucher',VoucherToken:tok,Status:'Active',AmountDue:price,AmountPaid:0,PaymentStatus:'Pending',PaymentMethod:'',PaymentNote:'',UpdatedAt:now});v.Claimed=true;v.Serial=serial;v.Status='Claimed';writeRow_(SHEETS.VOUCHERS,HEADERS.Vouchers,v._row,v);audit_(eventId,'CLAIM_TICKET','Claim',serial,name,'CLAIMANT',{tier:tier});return{ok:true,eventId:eventId,serial:serial,checkInToken:qr,tier:tier,price:price};}

// ---------- Ticket admin / payments ----------
function guestSearch_(eventId,q){q=String(q||'').trim().toLowerCase();if(!q)return[];const check=rows_(SHEETS.CHECKINS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Undone');return rows_(SHEETS.CLAIMS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Revoked').filter(r=>[r.Name,r.Email,r.Phone,r.Serial].some(v=>String(v||'').toLowerCase().includes(q))).slice(0,50).map(r=>({serial:r.Serial,name:r.Name,email:r.Email||'',phone:r.Phone||'',groupName:r.GroupName||'',tier:r.TierKey,paymentStatus:r.PaymentStatus||'Pending',paymentMethod:r.PaymentMethod||'',amountDue:money_(r.AmountDue),amountPaid:money_(r.AmountPaid),status:r.Status,checkedIn:check.some(c=>c.Serial===r.Serial)}));}
function updatePayment_(body,u){const eventId=normId_(body.eventId),r=rowBy_(SHEETS.CLAIMS,'Serial',String(body.serial||''));if(!r||normId_(r.EventID)!==eventId)return{ok:false,error:'Ticket not found.'};const amount=money_(body.amountPaid),status=String(body.paymentStatus||'Pending'),method=String(body.paymentMethod||''),note=String(body.paymentNote||'');r.AmountPaid=amount;r.PaymentStatus=status;r.PaymentMethod=method;r.PaymentNote=note;r.UpdatedAt=new Date();writeRow_(SHEETS.CLAIMS,HEADERS.Claims,r._row,r);append_(SHEETS.PAYMENTS,HEADERS.Payments,{Timestamp:new Date(),EventID:eventId,Serial:r.Serial,Amount:amount,Method:method,Status:status,Note:note,RecordedBy:u.name});audit_(eventId,'UPDATE_PAYMENT','Claim',r.Serial,u.name,u.role,{amount:amount,status:status,method:method});return{ok:true};}
function transferTicket_(body,u){const eventId=normId_(body.eventId),r=rowBy_(SHEETS.CLAIMS,'Serial',String(body.serial||''));if(!r||normId_(r.EventID)!==eventId)return{ok:false,error:'Ticket not found.'};const old=r.Name;r.TransferredFrom=[r.TransferredFrom,old].filter(Boolean).join(' | ');r.Name=String(body.newName||'').trim()||r.Name;r.Email=String(body.newEmail||'').trim();r.Phone=String(body.newPhone||'').trim();r.UpdatedAt=new Date();writeRow_(SHEETS.CLAIMS,HEADERS.Claims,r._row,r);audit_(eventId,'TRANSFER_TICKET','Claim',r.Serial,u.name,u.role,{from:old,to:r.Name});return{ok:true};}
function reissueQr_(body,u){const eventId=normId_(body.eventId),r=rowBy_(SHEETS.CLAIMS,'Serial',String(body.serial||''));if(!r||normId_(r.EventID)!==eventId)return{ok:false,error:'Ticket not found.'};r.CheckInToken=token_().slice(0,48);r.UpdatedAt=new Date();writeRow_(SHEETS.CLAIMS,HEADERS.Claims,r._row,r);audit_(eventId,'REISSUE_QR','Claim',r.Serial,u.name,u.role,{});return{ok:true,checkInToken:r.CheckInToken,tier:r.TierKey};}
function setTicketStatus_(body,u,status){const eventId=normId_(body.eventId),r=rowBy_(SHEETS.CLAIMS,'Serial',String(body.serial||''));if(!r||normId_(r.EventID)!==eventId)return{ok:false,error:'Ticket not found.'};r.Status=status;r.UpdatedAt=new Date();writeRow_(SHEETS.CLAIMS,HEADERS.Claims,r._row,r);audit_(eventId,status==='Revoked'?'REVOKE_TICKET':'REACTIVATE_TICKET','Claim',r.Serial,u.name,u.role,{});return{ok:true};}


function gateSummary_(eventId){const a=allocations_(eventId),claims=rows_(SHEETS.CLAIMS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Revoked'),checks=rows_(SHEETS.CHECKINS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Undone'),e=getEvent_(eventId);return{ok:true,counts:{issued:claims.length,checkedIn:checks.length,notArrived:Math.max(0,claims.length-checks.length),allocated:a.total,capacity:e&&e.Capacity||0}};}
function gateSearch_(eventId,q,u){
  const results=guestSearch_(eventId,q);const elevated=['SYSTEM_OWNER','EVENT_ADMIN','GATE_SUPERVISOR'].includes(u.role);
  return results.map(r=>({serial:r.serial,name:r.name,phone:elevated?r.phone:'',email:elevated?r.email:'',groupName:r.groupName,tier:r.tier,paymentStatus:r.paymentStatus,amountPaid:elevated?r.amountPaid:0,checkedIn:r.checkedIn,status:r.status}));
}

// ---------- Gate ----------
function lookupClaim_(eventId,value){const s=String(value||'').trim();return rows_(SHEETS.CLAIMS).find(r=>normId_(r.EventID)===eventId&&(String(r.CheckInToken)===s||String(r.Serial).toUpperCase()===s.toUpperCase()))||null;}
function checkIn_(body,u){const eventId=normId_(body.eventId);if(!scopeAllows_(u,eventId)||!roleCanGate_(u.role))return{ok:false,error:'Not permitted for this event.'};const ev=getEvent_(eventId);if(!ev||ev.Status!=='Active')return{ok:false,error:'This event is not active for gate check-in.'};const c=lookupClaim_(eventId,body.serial);if(!c)return{ok:false,error:'Ticket not found.'};if(String(c.Status)==='Revoked')return{ok:false,error:'Ticket has been revoked.',revoked:true,serial:c.Serial,name:c.Name};const existing=rows_(SHEETS.CHECKINS).find(r=>normId_(r.EventID)===eventId&&r.Serial===c.Serial&&String(r.Status)!=='Undone');if(existing)return{ok:false,error:'Already checked in.',alreadyCheckedIn:true,serial:c.Serial,name:c.Name,checkedInAt:existing.Timestamp,checkedInBy:existing.CheckedInBy};const payStatus=String(body.paymentStatus||c.PaymentStatus||'Pending'),payMethod=String(body.paymentMethod||c.PaymentMethod||''),amount=body.amountPaid!==undefined?money_(body.amountPaid):money_(c.AmountPaid);append_(SHEETS.CHECKINS,HEADERS.CheckIns,{Timestamp:new Date(),EventID:eventId,Serial:c.Serial,Name:c.Name,GroupName:c.GroupName,TierKey:c.TierKey,Phone:c.Phone,PaymentStatus:payStatus,PaymentMethod:payMethod,AmountPaid:amount,CheckedInBy:u.name,GateNote:String(body.gateNote||''),Status:'Checked In'});audit_(eventId,'CHECK_IN','Claim',c.Serial,u.name,u.role,{});return{ok:true,serial:c.Serial,name:c.Name,groupName:c.GroupName,tier:c.TierKey,paymentStatus:payStatus};}
function undoCheckIn_(body,u){if(!['SYSTEM_OWNER','EVENT_ADMIN','GATE_SUPERVISOR'].includes(u.role))return{ok:false,error:'Supervisor permission required.'};const eventId=normId_(body.eventId),serial=String(body.serial||'');const matches=rows_(SHEETS.CHECKINS).filter(r=>normId_(r.EventID)===eventId&&r.Serial===serial&&String(r.Status)!=='Undone');if(!matches.length)return{ok:false,error:'No active check-in found.'};const r=matches[matches.length-1];r.Status='Undone';r.GateNote=(r.GateNote?String(r.GateNote)+' | ':'')+'Undone by '+u.name+': '+String(body.reason||'');writeRow_(SHEETS.CHECKINS,HEADERS.CheckIns,r._row,r);audit_(eventId,'UNDO_CHECK_IN','CheckIn',serial,u.name,u.role,{reason:body.reason||''});return{ok:true};}
function walkIn_(body,u){if(!['SYSTEM_OWNER','EVENT_ADMIN','GATE_SUPERVISOR'].includes(u.role))return{ok:false,error:'Supervisor permission required for walk-ins.'};const eventId=normId_(body.eventId),e=getEvent_(eventId);if(!e||e.Status!=='Active')return{ok:false,error:'Event not active.'};const name=String(body.fullName||'').trim();if(!name)return{ok:false,error:'Name is required.'};const tier=String(body.tier||''),tiers=getTiers_(eventId);if(!tiers[tier])return{ok:false,error:'Choose a valid tier.'};const cap=capacity_(eventId,{[tier]:1},bool_(body.capacityOverride));if(!cap.ok)return cap;const serial=nextSerial_(eventId),qr=token_().slice(0,48),amount=money_(body.amountPaid),status=String(body.paymentStatus||'Paid'),method=String(body.paymentMethod||'Cash'),now=new Date();append_(SHEETS.CLAIMS,HEADERS.Claims,{Timestamp:now,EventID:eventId,Serial:serial,CheckInToken:qr,Name:name,Email:'',Phone:String(body.phone||''),GroupName:cleanGroup_(body.groupName),TierKey:tier,Source:'Walk-In',Status:'Active',AmountDue:money_(tiers[tier].price),AmountPaid:amount,PaymentStatus:status,PaymentMethod:method,PaymentNote:String(body.note||''),UpdatedAt:now});append_(SHEETS.CHECKINS,HEADERS.CheckIns,{Timestamp:now,EventID:eventId,Serial:serial,Name:name,GroupName:cleanGroup_(body.groupName),TierKey:tier,Phone:String(body.phone||''),PaymentStatus:status,PaymentMethod:method,AmountPaid:amount,CheckedInBy:u.name,GateNote:'Walk-in',Status:'Checked In'});if(amount>0)append_(SHEETS.PAYMENTS,HEADERS.Payments,{Timestamp:now,EventID:eventId,Serial:serial,Amount:amount,Method:method,Status:status,Note:'Walk-in',RecordedBy:u.name});audit_(eventId,'WALK_IN','Claim',serial,u.name,u.role,{});return{ok:true,serial:serial,name:name,checkInToken:qr,tier:tier};}

// ---------- Reporting ----------
function dashboard_(eventId){const e=getEvent_(eventId),tiers=getTiers_(eventId),alloc=allocations_(eventId),claims=rows_(SHEETS.CLAIMS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Revoked'),checkins=rows_(SHEETS.CHECKINS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Undone'),vouchers=rows_(SHEETS.VOUCHERS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Cancelled'),payments=rows_(SHEETS.PAYMENTS).filter(r=>normId_(r.EventID)===eventId);const due=claims.reduce((s,r)=>s+money_(r.AmountDue),0),paid=claims.reduce((s,r)=>s+money_(r.AmountPaid),0),byTier={};Object.keys(tiers).forEach(k=>byTier[k]={label:tiers[k].label,color:tiers[k].color,issued:0,checkedIn:0});claims.forEach(r=>{if(byTier[r.TierKey])byTier[r.TierKey].issued++});checkins.forEach(r=>{if(byTier[r.TierKey])byTier[r.TierKey].checkedIn++});return{ok:true,event:publicEvent_(e),counts:{allocated:alloc.total,claims:claims.length,pendingVouchers:vouchers.filter(v=>!bool_(v.Claimed)).length,checkedIn:checkins.length,noShows:Math.max(0,claims.length-checkins.length),capacity:e.Capacity||0},finance:{faceValue:due,amountPaid:paid,outstanding:Math.max(0,due-paid),paymentEntries:payments.length},byTier:byTier};}
function auditList_(eventId){return{ok:true,rows:rows_(SHEETS.AUDIT).filter(r=>normId_(r.EventID)===eventId).slice(-250).reverse().map(r=>({timestamp:r.Timestamp,action:r.Action,entityType:r.EntityType,entityId:r.EntityID,actor:r.Actor,role:r.Role,details:r.Details}))};}

// ---------- Communication ----------
function logCommunication_(body,u){const eventId=normId_(body.eventId);if(eventId&&!scopeAllows_(u,eventId))return{ok:false,error:'Not assigned.'};if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE','GATE_SUPERVISOR'].includes(u.role))return{ok:false,error:'Communication permission required.'};contactLog_(eventId,String(body.entityType||'Contact'),String(body.entityId||''),String(body.channel||''),String(body.recipient||''),u.name,String(body.note||'Initiated from web interface'));return{ok:true};}
function sendAdminEmail_(body,u){const eventId=normId_(body.eventId),to=String(body.to||'').trim();if(!to)return{ok:false,error:'Email address required.'};try{MailApp.sendEmail(to,String(body.subject||'Event ticketing message'),String(body.message||''));contactLog_(eventId,String(body.entityType||'Contact'),String(body.entityId||''),'Email',to,u.name,String(body.subject||''));return{ok:true};}catch(e){return{ok:false,error:e.message};}}

// ---------- Status / admin maintenance ----------
function setEventStatus_(body,u){const eventId=normId_(body.eventId),status=String(body.status||'');if(!LIFECYCLE.includes(status))return{ok:false,error:'Invalid status.'};const r=getEventBase_(eventId);if(!r)return{ok:false,error:'Event not found.'};if(status==='Active'&&r.Status!=='Client Approved'&&u.role!=='SYSTEM_OWNER')return{ok:false,error:'Client approval is required before activation.'};r.Status=status;r.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,r._row,r);audit_(eventId,'SET_EVENT_STATUS','Event',eventId,u.name,u.role,{status:status});return{ok:true,status:status};}
function configLink_(eventId,u){const r=getEventBase_(eventId);if(!r||!scopeAllows_(u,eventId))return{ok:false,error:'Event not found or not assigned.'};if(!r.ConfigToken){r.ConfigToken=token_();r.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,r._row,r);}return{ok:true,url:CONFIG_URL+'?event='+encodeURIComponent(r.EventID)+'&key='+encodeURIComponent(r.ConfigToken),formUrl:formPrefillUrl_(r.EventID,r.ConfigToken),clientName:r.ClientName||'',clientEmail:r.ClientEmail||'',clientPhone:r.ClientPhone||''};}


function accessList_(u){if(u.role!=='SYSTEM_OWNER')return{ok:false,error:'System Owner permission required.'};return{ok:true,users:rows_(SHEETS.ACCESS).map(r=>({userId:r.UserID,name:r.Name,role:r.Role,eventScope:r.EventScope,email:r.Email||'',phone:r.Phone||'',active:bool_(r.Active)}))};}
function createAccessUser_(body,u){if(u.role!=='SYSTEM_OWNER')return{ok:false,error:'System Owner permission required.'};const role=String(body.role||'').toUpperCase();if(!ROLES.includes(role))return{ok:false,error:'Invalid role.'};const pass=String(body.newPasscode||'').trim();if(pass.length<8)return{ok:false,error:'Passcode must be at least 8 characters.'};const id='USR-'+Utilities.getUuid().slice(0,8).toUpperCase();append_(SHEETS.ACCESS,HEADERS.AccessControl,{UserID:id,Name:String(body.name||'').trim(),Role:role,PasscodeHash:hash_(pass),EventScope:String(body.eventScope||'*').trim()||'*',Email:String(body.email||'').trim(),Phone:String(body.phone||'').trim(),Active:true,CreatedAt:new Date()});audit_('','CREATE_ACCESS_USER','Access',id,u.name,u.role,{role:role,scope:body.eventScope||'*'});return{ok:true,userId:id};}
function setAccessActive_(body,u){if(u.role!=='SYSTEM_OWNER')return{ok:false,error:'System Owner permission required.'};const r=rowBy_(SHEETS.ACCESS,'UserID',String(body.userId||''));if(!r)return{ok:false,error:'User not found.'};r.Active=bool_(body.active);writeRow_(SHEETS.ACCESS,HEADERS.AccessControl,r._row,r);audit_('','SET_ACCESS_ACTIVE','Access',r.UserID,u.name,u.role,{active:r.Active});return{ok:true};}
function voucherSearch_(eventId,q){q=String(q||'').toLowerCase().trim();return rows_(SHEETS.VOUCHERS).filter(r=>normId_(r.EventID)===eventId).filter(r=>!q||[r.BatchID,r.VoucherToken,r.RecipientName,r.RecipientEmail,r.RecipientPhone,r.Serial].some(v=>String(v||'').toLowerCase().includes(q))).slice(-100).reverse().map(r=>({token:r.VoucherToken,batchId:r.BatchID,tier:r.TierKey,status:r.Status||'Available',claimed:bool_(r.Claimed),dispatched:bool_(r.Dispatched),serial:r.Serial||'',recipientName:r.RecipientName||'',recipientEmail:r.RecipientEmail||'',recipientPhone:r.RecipientPhone||''}));}
function cancelVoucher_(body,u){const r=rowBy_(SHEETS.VOUCHERS,'VoucherToken',String(body.token||''));if(!r)return{ok:false,error:'Voucher not found.'};const eventId=normId_(r.EventID);if(!scopeAllows_(u,eventId)||!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return{ok:false,error:'Admin permission required.'};if(bool_(r.Claimed))return{ok:false,error:'Claimed vouchers cannot be cancelled. Revoke the issued ticket instead.'};r.Status='Cancelled';writeRow_(SHEETS.VOUCHERS,HEADERS.Vouchers,r._row,r);audit_(eventId,'CANCEL_VOUCHER','Voucher',r.VoucherToken,u.name,u.role,{});return{ok:true};}
function backupEvent_(eventId,u){if(!scopeAllows_(u,eventId)||!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return{ok:false,error:'Admin permission required.'};const payload={createdAt:new Date().toISOString(),event:getEvent_(eventId),tiers:getTiers_(eventId),groups:getGroups_(eventId),vouchers:rows_(SHEETS.VOUCHERS).filter(r=>normId_(r.EventID)===eventId),claims:rows_(SHEETS.CLAIMS).filter(r=>normId_(r.EventID)===eventId),checkIns:rows_(SHEETS.CHECKINS).filter(r=>normId_(r.EventID)===eventId),payments:rows_(SHEETS.PAYMENTS).filter(r=>normId_(r.EventID)===eventId),audit:rows_(SHEETS.AUDIT).filter(r=>normId_(r.EventID)===eventId)};const blob=Utilities.newBlob(JSON.stringify(payload,null,2),'application/json',eventId+'-backup-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'America/New_York','yyyyMMdd-HHmmss')+'.json');const f=assetFolder_().createFile(blob);audit_(eventId,'BACKUP_EVENT','Event',eventId,u.name,u.role,{fileId:f.getId()});return{ok:true,fileName:f.getName(),driveUrl:f.getUrl()};}

// ---------- Web API ----------
function doGet(e){const p=e&&e.parameter||{},action=String(p.action||'');try{
  if(action==='clientContext')return json_(clientContext_(normId_(p.event),p.key));
  if(action==='voucher')return json_(voucher_(p.token));
  if(action==='batch')return json_(batch_(p.batch));
  return json_({ok:true,service:'Event Ticketing V5'});
}catch(err){return json_({ok:false,error:err.message});}}

function doPost(e){const b=body_(e),action=String(b.action||'');try{
  // public/client/voucher actions
  if(action==='clientContext')return json_(clientContext_(normId_(b.eventId),b.configKey));
  if(action==='clientUpload')return json_(clientUpload_(b));
  if(action==='clientPreview')return json_(clientPreview_(b));
  if(action==='clientApprove')return json_(clientApprove_(b));
  if(action==='sendVoucher')return json_(sendVoucher_(b));
  if(action==='claim')return json_(claim_(b));

  const u=authenticate_(b.adminPasscode);if(!u)return json_({ok:false,error:'Incorrect passcode.'});
  if(action==='verifyPasscode')return json_({ok:true,user:u});
  if(action==='eventsAdmin'){if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE'].includes(u.role))return json_({ok:false,error:'Admin or Finance role required.'});return json_({ok:true,user:u,events:currentFutureEvents_(u,bool_(b.includePast))});}
  if(action==='createEventShell'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return json_({ok:false,error:'Admin permission required.'});return json_(createEventShell_(b,u));}
  if(action==='configLink'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return json_({ok:false,error:'Admin permission required.'});return json_(configLink_(normId_(b.eventId),u));}
  if(action==='setEventStatus'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return json_({ok:false,error:'Admin permission required.'});if(!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Not assigned.'});return json_(setEventStatus_(b,u));}
  if(action==='generateVouchers'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return json_({ok:false,error:'Admin permission required.'});return json_(generateVouchers_(b,u));}
  if(action==='adminEventContext'){if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin or Finance permission required.'});const id=normId_(b.eventId);return json_({ok:true,event:publicEvent_(getEvent_(id)),tiers:getTiers_(id),groups:getGroups_(id)});}
  if(action==='dashboard'){if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin or Finance permission required.'});return json_(dashboard_(normId_(b.eventId)));}
  if(action==='searchGuests'){if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin or Finance permission required.'});return json_({ok:true,results:guestSearch_(normId_(b.eventId),b.query)});}
  if(action==='updatePayment'){if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE','GATE_SUPERVISOR'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Payment permission required.'});return json_(updatePayment_(b,u));}
  if(action==='transferTicket'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_(transferTicket_(b,u));}
  if(action==='reissueQr'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_(reissueQr_(b,u));}
  if(action==='revokeTicket'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_(setTicketStatus_(b,u,'Revoked'));}
  if(action==='reactivateTicket'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_(setTicketStatus_(b,u,'Active'));}
  if(action==='audit'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_(auditList_(normId_(b.eventId)));}
  if(action==='sendAdminEmail'){if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE','GATE_SUPERVISOR'].includes(u.role))return json_({ok:false,error:'Communication permission required.'});return json_(sendAdminEmail_(b,u));}
  if(action==='logCommunication')return json_(logCommunication_(b,u));
  if(action==='accessList')return json_(accessList_(u));
  if(action==='createAccessUser')return json_(createAccessUser_(b,u));
  if(action==='setAccessActive')return json_(setAccessActive_(b,u));
  if(action==='voucherSearch'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_({ok:true,rows:voucherSearch_(normId_(b.eventId),b.query)});}
  if(action==='cancelVoucher')return json_(cancelVoucher_(b,u));
  if(action==='backupEvent')return json_(backupEvent_(normId_(b.eventId),u));

  // gate-specific actions
  if(action==='gateEvents'){if(!roleCanGate_(u.role))return json_({ok:false,error:'Gate role required.'});return json_({ok:true,user:u,events:currentFutureEvents_(u,false).filter(x=>x.Status==='Active')});}
  if(action==='gateContext'){if(!roleCanGate_(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Gate permission required.'});const id=normId_(b.eventId);return json_({ok:true,user:u,event:publicEvent_(getEvent_(id)),tiers:getTiers_(id),groups:getGroups_(id)});}
  if(action==='gateSearch'){if(!roleCanGate_(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Gate permission required.'});return json_({ok:true,results:gateSearch_(normId_(b.eventId),b.query,u)});}
  if(action==='gateSummary'){if(!roleCanGate_(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Gate permission required.'});return json_(gateSummary_(normId_(b.eventId)));}
  if(action==='checkIn')return json_(checkIn_(b,u));
  if(action==='undoCheckIn')return json_(undoCheckIn_(b,u));
  if(action==='walkIn')return json_(walkIn_(b,u));

  return json_({ok:false,error:'Unknown action.'});
}catch(err){return json_({ok:false,error:err.message});}}
