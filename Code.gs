/**
 * KOSOL ENERGIE PVT LTD — Sales Team Management System
 * Production Google Apps Script backend for the current Kosol_Energie_Sales_Management_Pipeline.html
 *
 * Setup:
 * 1. Create/open the Google Sheet that will be the database.
 * 2. Extensions -> Apps Script -> replace Code.gs with this file.
 * 3. Run setupKosol() once and authorize.
 * 4. Deploy -> New deployment -> Web app -> Execute as Me -> Anyone with the link.
 * 5. Put the /exec URL in the HTML SCRIPT_URL (or config used by the HTML).
 *
 * The API contract intentionally returns {success:true/false,message:...} because
 * the supplied HTML consumes that response shape.
 */

const CFG = {
  timezone: 'Asia/Kolkata',
  imageMaxBytes: 5 * 1024 * 1024,
  sessionHours: 8,
  roles: ['Master Admin','National Head','State Head','Regional Head','Senior Sales Executive','Sales Executive'],
  sheets: {
    USERS:'Users', CLIENTS:'Clients', COLD:'ColdCalls', TASKS:'Tasks', TASK_HISTORY:'TaskHistory',
    APPROVALS:'Approvals', ATTENDANCE:'Attendance', CHAT:'Chats', SALES:'SalesOrders',
    NOTIFICATIONS:'Notifications', LOCATION:'LocationMaster', AUDIT:'AuditLog', SETTINGS:'Settings'
  }
};

const HEADERS = {
  Users:['ID','Name','Email','Phone','Role','States','Regions','Districts','Tehsils','Cities','PinCodes','ManagerEmail','Status','PasswordHash','CreatedAt','UpdatedAt','LastLogin','OnlineAt'],
  Clients:['ID','OrganizationID','Organization','ClientName','Mobile','OfficeMobile','Email','GSTIN','Address','State','Region','District','Tehsil','CityVillage','PinCode','SalesOrderStatus','SalesType','LeadTemperature','VisitType','VisitCount','VisitStatus','MeetingStatus','CallStatus','Requirements','Remarks','EntryDate','ActivityDate','Latitude','Longitude','MapUrl','LocationImageUrl','CreatedBy','CreatedByName','CreatedAt','UpdatedAt'],
  ColdCalls:['ID','ClientID','Organization','ClientName','ContactName','Phone','OfficePhone','Email','State','Region','District','Tehsil','CityVillage','PinCode','Address','VisitType','VisitCount','Reason','CallStatus','Outcome','LeadTemperature','Requirements','Remarks','EntryDate','UserID','UserEmail','UserName','CreatedAt','UpdatedAt'],
  Tasks:['ID','Title','AssignedBy','AssignedTo','FollowUpDate','Status','Remarks','CreatedAt','UpdatedAt'],
  TaskHistory:['ID','TaskID','SenderEmail','SenderName','Message','Type','Timestamp'],
  Approvals:['ID','RequestedByEmail','RequestedByName','RequestedByRole','TargetApproverRole','Category','Details','Status','ApproverEmail','ApproverName','ApproverRemarks','CreatedAt','UpdatedAt'],
  Attendance:['ID','UserID','UserEmail','Name','Date','Type','Timestamp','Status','Remarks','Explanation','Latitude','Longitude','CreatedAt','UpdatedAt'],
  Chats:['ID','SenderID','SenderEmail','SenderName','ReceiverID','ReceiverEmail','Message','Timestamp','Read','CreatedAt'],
  SalesOrders:['ID','UserID','UserEmail','UserRole','ClientID','OrganizationID','Organization','ClientName','ContactNumber','PIStatus','OrderFor','SalesType','PVType','PanelQuantity','PanelWP','RatePerWP','LPD','RatePerLPD','SalesLPD','ProjectKW','RatePerKW','KW','Amount','RepeatCount','Date','Remarks','CreatedBy','CreatedAt','UpdatedAt'],
  Notifications:['ID','UserID','UserEmail','Type','Module','RecordID','Title','Message','Timestamp','Read','CreatedAt'],
  LocationMaster:['Type','State','Region','District','Tehsil','CityVillage','PIN','Office','UpdatedAt'],
  AuditLog:['ID','ActorID','ActorEmail','ActorRole','Action','Entity','RecordID','Details','Timestamp'],
  Settings:['Key','Value','UpdatedAt']
};

function doGet() {
  return json_({success:true,service:'KOSOL ENERGIE Sales Team Management System',version:'2.0.0',time:new Date().toISOString(),message:'POST JSON requests to this endpoint.'});
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(body.action || '').trim();
    const p = body.payload || {};
    let r;
    switch(action) {
      case 'login': r=login_(p); break;
      case 'selfRegister': r=selfRegister_(p); break;
      case 'getDashboardData': r=getDashboardData_(p); break;
      case 'saveClient': r=saveClient_(p); break;
      case 'deleteClient': r=deleteClient_(p); break;
      case 'saveColdCall': r=saveColdCall_(p,false); break;
      case 'updateColdCall': r=saveColdCall_(p,true); break;
      case 'deleteColdCall': r=deleteColdCall_(p); break;
      case 'saveSalesOrder': r=saveSalesOrder_(p); break;
      case 'assignTask': r=assignTask_(p); break;
      case 'updateTaskStatus': r=updateTaskStatus_(p); break;
      case 'addTaskReply': r=addTaskReply_(p); break;
      case 'getTaskHistory': r=getTaskHistory_(p); break;
      case 'getPendingUsers': r=getPendingUsers_(p); break;
      case 'processUserApproval': r=processUserApproval_(p); break;
      case 'getAllUsersDirectory': r=getAllUsersDirectory_(p); break;
      case 'updateUser': r=updateUser_(p); break;
      case 'deleteUser': r=deleteUser_(p); break;
      case 'markAttendance': r=markAttendance_(p); break;
      case 'submitLateExplanation': r=submitLateExplanation_(p); break;
      case 'getPendingAttendanceApprovals': r=getPendingAttendanceApprovals_(p); break;
      case 'processAttendanceApproval': r=processAttendanceApproval_(p); break;
      case 'submitApprovalRequest': r=submitApprovalRequest_(p); break;
      case 'processGeneralApproval': r=processGeneralApproval_(p); break;
      case 'getChatUsers': r=getChatUsers_(p); break;
      case 'getChatMessages': r=getChatMessages_(p); break;
      case 'sendChatMessage': r=sendChatMessage_(p); break;
      case 'touchUserActivity': r=touchUserActivity_(p); break;
      case 'getLocationData': r=getLocationData_(p); break;
      case 'lookupPinCode': r=lookupPinCode_(p); break;
      case 'refreshLGDLocationMaster': r=refreshLGDLocationMaster_(p); break;
      default: r={success:false,message:'Unknown action: '+action};
    }
    return json_(r);
  } catch(err) {
    console.error(err);
    return json_({success:false,message:String(err && err.message || err)});
  }
}

function setupKosol() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(name=>{
    let sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name);
    sh.clear();
    sh.getRange(1,1,1,HEADERS[name].length).setValues([HEADERS[name]]);
    sh.setFrozenRows(1); sh.getRange(1,1,1,HEADERS[name].length).setFontWeight('bold');
    try{sh.autoResizeColumns(1,Math.min(HEADERS[name].length,20));}catch(_){ }
  });
  const now=iso_();
  const users=[
    userRow_('U001','Master Admin','admin@kosolenergie.in','', 'Master Admin',['All India'],['All Regions'],[],[],[],[],'','Active','Admin@123',now),
    userRow_('U002','National Head','national@kosolenergie.in','', 'National Head',['All India'],['All Regions'],[],[],[],[],'admin@kosolenergie.in','Active','National@123',now),
    userRow_('U003','State Head','state@kosolenergie.in','', 'State Head',['Maharashtra'],['Maharashtra'],[],[],[],[],'national@kosolenergie.in','Active','State@123',now),
    userRow_('U004','Regional Head','regional@kosolenergie.in','', 'Regional Head',['Maharashtra'],['Vidarbha'],[],[],[],[],'state@kosolenergie.in','Active','Regional@123',now),
    userRow_('U005','Senior Sales Executive','senior@kosolenergie.in','', 'Senior Sales Executive',['Maharashtra'],['Vidarbha'],['Wardha'],[],[],[],'regional@kosolenergie.in','Active','Senior@123',now),
    userRow_('U006','Sales Executive','sales@kosolenergie.in','', 'Sales Executive',['Maharashtra'],['Vidarbha'],['Wardha'],[],[],[],'senior@kosolenergie.in','Active','Sales@123',now)
  ];
  const ush=ss.getSheetByName(CFG.sheets.USERS); ush.getRange(2,1,users.length,HEADERS.Users.length).setValues(users);
  const settings=ss.getSheetByName(CFG.sheets.SETTINGS);
  settings.getRange(2,1,6,2).setValues([
    ['Organization','KOSOL ENERGIE Pvt Ltd'],['Coverage','India'],['Timezone',CFG.timezone],['Version','2.0.0'],['SetupCompletedAt',now],['SessionHours',String(CFG.sessionHours)]
  ]);
  seedLocations_();
  return {success:true,message:'KOSOL ENERGIE database initialized. Change seeded passwords before production.'};
}

function login_(p) {
  const email=normEmail_(p.email), pass=String(p.password||'');
  if(!email||!pass)return fail_('Email and password are required.');
  const u=findBy_(CFG.sheets.USERS,'Email',email); if(!u)return fail_('Invalid credentials.');
  if(String(u.Status).toLowerCase()!=='active')return fail_('User account is inactive or pending approval.');
  if(String(u.PasswordHash)!==hash_(pass))return fail_('Invalid credentials.');
  u.LastLogin=iso_();u.OnlineAt=iso_();upsert_(CFG.sheets.USERS,'ID',u.ID,u);
  const safe=safeUser_(u); safe.authToken=issueToken_(u.Email); return ok_('Login successful.',{user:safe});
}

function selfRegister_(p) {
  const master=String(p.createdBy||'').trim();
  if(master){
    const actor=findBy_(CFG.sheets.USERS,'Email',normEmail_(master));
    if(!actor || String(actor.Status)!=='Active' || actor.Role!=='Master Admin') return fail_('Only an active Master Admin can create users directly.');
    return createOrUpdateUser_(p,actor,true);
  }
  const name=String(p.name||'').trim(), email=normEmail_(p.email), pass=String(p.password||''), role=String(p.role||'Sales Executive');
  if(!name||!email||!pass||!role)return fail_('Name, email, password and role are required.');
  if(CFG.roles.indexOf(role)<0)return fail_('Invalid role.');
  if(findBy_(CFG.sheets.USERS,'Email',email))return fail_('An account with this email already exists.');
  const now=iso_(), id=id_('U');
  const u={ID:id,Name:name,Email:email,Phone:String(p.phone||''),Role:role,States:join_(p.states),Regions:join_(p.regions),Districts:join_(p.districts),Tehsils:join_(p.tehsils),Cities:join_(p.cities),PinCodes:join_(p.pinCodes),ManagerEmail:'',Status:'Pending',PasswordHash:hash_(pass),CreatedAt:now,UpdatedAt:now,LastLogin:'',OnlineAt:''};
  appendObj_(CFG.sheets.USERS,u);
  return ok_('Registration submitted for Master Admin approval.',{user:safeUser_(u),userId:id});
}

function createOrUpdateUser_(p,actor,active) {
  const name=String(p.name||'').trim(), email=normEmail_(p.email), role=String(p.role||'Sales Executive');
  if(!name||!email||!role||!String(p.password||''))return fail_('Name, email, password and role are required.');
  if(CFG.roles.indexOf(role)<0)return fail_('Invalid role.');
  const existing=findBy_(CFG.sheets.USERS,'Email',email), now=iso_(), id=existing?existing.ID:id_('U');
  const u={ID:id,Name:name,Email:email,Phone:String(p.phone||''),Role:role,States:join_(p.states),Regions:join_(p.regions),Districts:join_(p.districts),Tehsils:join_(p.tehsils),Cities:join_(p.cities),PinCodes:join_(p.pinCodes),ManagerEmail:String(p.managerEmail||''),Status:active?'Active':'Pending',PasswordHash:hash_(p.password),CreatedAt:existing?existing.CreatedAt:now,UpdatedAt:now,LastLogin:existing?existing.LastLogin:'',OnlineAt:existing?existing.OnlineAt:''};
  upsert_(CFG.sheets.USERS,'ID',id,u); audit_(actor,'CREATE/UPDATE','User',id,JSON.stringify({email,role,status:u.Status}));
  return ok_('User created successfully.',{user:safeUser_(u),userId:id});
}

function getDashboardData_(p) {
  const actor=requireUser_(p); cleanOnline_();
  const clients=filterClients_(rows_(CFG.sheets.CLIENTS),actor);
  const cold=filterCold_(rows_(CFG.sheets.COLD),actor);
  const tasks=filterTasks_(rows_(CFG.sheets.TASKS),actor);
  const approvals=filterApprovals_(rows_(CFG.sheets.APPROVALS),actor);
  const sales=filterSales_(rows_(CFG.sheets.SALES),actor);
  const today=fmtDate_(new Date());
  const attendance=rows_(CFG.sheets.ATTENDANCE).filter(a=>normEmail_(a.UserEmail)===actor.Email && String(a.Date)===today).sort((a,b)=>String(b.Timestamp).localeCompare(String(a.Timestamp)))[0]||null;
  const users=safeUsersFor_(actor);
  return {success:true,message:'Dashboard data loaded.',clients,coldCalls:cold,tasks,approvals,salesOrders:sales,todayAttendance:attendance,teamUsers:users};
}

function saveClient_(p) {
  const actor=requireUser_(p);
  if(!p.orgName||!p.contactName||!p.mobile)return fail_('Organization, contact person and mobile are required.');
  const id=String(p.id||p.selectedClientId||'')||id_('C');
  const existing=findBy_(CFG.sheets.CLIENTS,'ID',id);
  if(existing && !canSeeClient_(actor,existing))return fail_('You are not authorized to edit this client.');
  const mobile=digits_(p.mobile);
  const dup=rows_(CFG.sheets.CLIENTS).find(c=>digits_(c.Mobile)===mobile && String(c.ID)!==id);
  if(dup && !p.selectedClientId)return fail_('A client with this mobile number already exists: '+dup.ID);
  let img=String(p.locationImageData||'');
  if(img.indexOf('data:image/')===0)img=saveImage_(img,id);
  const now=iso_();
  const obj={
    ID:id,OrganizationID:String(p.organizationId||p.selectedClientId||existing?.OrganizationID||id),Organization:String(p.orgName||''),ClientName:String(p.contactName||''),Mobile:String(p.mobile||''),OfficeMobile:String(p.officeMobile||''),Email:String(p.email||''),GSTIN:String(p.gstin||''),Address:String(p.address||''),State:String(p.state||''),Region:String(p.region||''),District:String(p.district||''),Tehsil:String(p.tehsil||''),CityVillage:String(p.city||''),PinCode:String(p.pinCode||''),SalesOrderStatus:String(p.salesOrderStatus||''),SalesType:String(p.salesType||''),LeadTemperature:String(p.leadTemp||'Warm'),VisitType:String(p.visitType||'Follow Up'),VisitCount:Number(p.visitCount||1),VisitStatus:String(p.visitStatus||''),MeetingStatus:String(p.meetingStatus||''),CallStatus:String(p.callStatus||''),Requirements:String(p.requirements||''),Remarks:String(p.remarks||''),EntryDate:String(p.entryDate||fmtDate_(new Date())),ActivityDate:fmtDate_(new Date()),Latitude:String(p.latitude||''),Longitude:String(p.longitude||''),MapUrl:String(p.mapUrl||''),LocationImageUrl:img||String(existing?.LocationImageUrl||''),CreatedBy:existing?.CreatedBy||actor.Email,CreatedByName:existing?.CreatedByName||actor.Name,CreatedAt:existing?.CreatedAt||now,UpdatedAt:now
  };
  if(Number(obj.VisitCount)<1)obj.VisitCount=1;
  upsert_(CFG.sheets.CLIENTS,'ID',id,obj);
  if(String(p.selectedClientId||'') && !existing) { /* ID was intentionally reused for the selected CRM record. */ }
  notify_(actor,'CLIENT','Clients',id,'Client / lead updated',obj.Organization+' — '+obj.ClientName);
  audit_(actor,'SAVE','Client',id,JSON.stringify({organization:obj.Organization,mobile:obj.Mobile,visitCount:obj.VisitCount}));
  return ok_('Client entry saved successfully.',{client:clientOut_(obj)});
}

function deleteClient_(p){const actor=requireUser_(p),id=String(p.id||'');const c=findBy_(CFG.sheets.CLIENTS,'ID',id);if(!c)return fail_('Client not found.');if(Number(actorLevel_(actor))!==0 && !sameEmail_(c.CreatedBy,actor.Email))return fail_('You are not authorized to delete this client.');deleteBy_(CFG.sheets.CLIENTS,'ID',id);audit_(actor,'DELETE','Client',id,'Client deleted');return ok_('Client deleted successfully.');}

function saveColdCall_(p,isUpdate){
  const actor=requireUser_(p); if(!p.phone)return fail_('Phone number is required.');
  const id=String(p.id||'')||id_('CC'); const existing=findBy_(CFG.sheets.COLD,'ID',id);
  if(existing && !canSeeCold_(actor,existing))return fail_('You are not authorized to edit this cold call.');
  const now=iso_(); const obj={ID:id,ClientID:String(p.clientId||''),Organization:String(p.orgName||p.clientName||''),ClientName:String(p.clientName||p.orgName||''),ContactName:String(p.contactName||''),Phone:String(p.phone||''),OfficePhone:String(p.officePhone||''),Email:String(p.email||''),State:String(p.state||''),Region:String(p.region||''),District:String(p.district||''),Tehsil:String(p.tehsil||''),CityVillage:String(p.city||''),PinCode:String(p.pinCode||''),Address:String(p.address||''),VisitType:String(p.visitType||'Follow Up'),VisitCount:Math.max(1,Number(p.visitCount||1)),Reason:String(p.reasonForNoVisit||''),CallStatus:String(p.callStatus||''),Outcome:String(p.callOutcome||''),LeadTemperature:String(p.leadTemp||'Cold'),Requirements:String(p.requirements||''),Remarks:String(p.remarks||''),EntryDate:String(p.entryDate||fmtDate_(new Date())),UserID:existing?.UserID||actor.ID,UserEmail:existing?.UserEmail||actor.Email,UserName:existing?.UserName||actor.Name,CreatedAt:existing?.CreatedAt||now,UpdatedAt:now};
  upsert_(CFG.sheets.COLD,'ID',id,obj); audit_(actor,isUpdate?'UPDATE':'SAVE','ColdCall',id,obj.Organization); return ok_('Cold call saved successfully.',{coldCall:coldOut_(obj)});
}
function deleteColdCall_(p){const actor=requireUser_(p),id=String(p.id||''),c=findBy_(CFG.sheets.COLD,'ID',id);if(!c)return fail_('Cold call not found.');if(Number(actorLevel_(actor))!==0&&!sameEmail_(c.UserEmail,actor.Email))return fail_('You are not authorized to delete this cold call.');deleteBy_(CFG.sheets.COLD,'ID',id);audit_(actor,'DELETE','ColdCall',id,'Cold call deleted');return ok_('Cold call deleted successfully.');}

function saveSalesOrder_(p){
  const actor=requireUser_(p); const id=String(p.id||'')||id_('SO'); const existing=findBy_(CFG.sheets.SALES,'ID',id); if(existing&&Number(actorLevel_(actor))!==0&&!sameEmail_(existing.UserEmail,actor.Email))return fail_('You are not authorized to edit this sales order.');
  const org=String(p.orgName||'').trim(); if(!org)return fail_('Organization is required.');
  const type=String(p.orderFor||'').trim(); if(!type)return fail_('Order type is required.');
  let kw=Number(p.kw||0), amount=Number(p.amount||0); if(!isFinite(kw))kw=0;if(!isFinite(amount))amount=0;
  const now=iso_(); const obj={ID:id,UserID:existing?.UserID||String(p.userId||actor.ID),UserEmail:existing?.UserEmail||actor.Email,UserRole:existing?.UserRole||actor.Role,ClientID:String(p.clientId||existing?.ClientID||''),OrganizationID:String(p.organizationId||p.clientId||existing?.OrganizationID||''),Organization:org,ClientName:String(p.clientName||''),ContactNumber:String(p.contactNumber||''),PIStatus:String(p.piStatus||''),OrderFor:type,SalesType:String(p.salesType||''),PVType:String(p.pvType||''),PanelQuantity:Number(p.panelQuantity||0),PanelWP:Number(p.panelWp||0),RatePerWP:Number(p.ratePerWp||0),LPD:Number(p.lpd||0),RatePerLPD:Number(p.ratePerLpd||0),SalesLPD:Number(p.salesLpd||0),ProjectKW:Number(p.projectKw||0),RatePerKW:Number(p.ratePerKw||0),KW:kw,Amount:amount,RepeatCount:Math.max(1,Number(p.repeatCount||1)),Date:String(p.date||fmtDate_(new Date())),Remarks:String(p.remarks||''),CreatedBy:existing?.CreatedBy||actor.Name,CreatedAt:existing?.CreatedAt||now,UpdatedAt:now};
  if(type==='Solar PV Modules' && !obj.KW && obj.PanelQuantity && obj.PanelWP)obj.KW=obj.PanelQuantity*obj.PanelWP/1000;
  if(type==='Solar PV Modules' && !obj.Amount)obj.Amount=obj.PanelQuantity*obj.PanelWP*obj.RatePerWP;
  if(type==='Solar Water Heater' && !obj.Amount)obj.Amount=obj.LPD*obj.RatePerLPD;
  if(type==='Solar EPC Projects - Utility/C & I/Residentials' && !obj.Amount)obj.Amount=obj.ProjectKW*obj.RatePerKW;
  upsert_(CFG.sheets.SALES,'ID',id,obj);
  notify_(actor,'SALES','Sales Orders',id,'Sales order recorded',obj.Organization+' — '+obj.OrderFor);
  audit_(actor,existing?'UPDATE':'SAVE','SalesOrder',id,JSON.stringify({organization:obj.Organization,repeatCount:obj.RepeatCount,amount:obj.Amount}));
  return ok_('Sales order saved successfully.',{salesOrder:salesOut_(obj)});
}

function assignTask_(p){const actor=requireUser_(p);if(!p.title||!p.assignedTo)return fail_('Task title and assignee are required.');const ass=findBy_(CFG.sheets.USERS,'Email',normEmail_(p.assignedTo));if(!ass||String(ass.Status)!=='Active')return fail_('Selected assignee is not active.');if(Number(actorLevel_(actor))>=4&&normEmail_(p.assignedTo)!==actor.Email)return fail_('Sales executives may assign tasks only to themselves.');const now=iso_(),id=id_('T');const obj={ID:id,Title:String(p.title),AssignedBy:actor.Email,AssignedTo:normEmail_(p.assignedTo),FollowUpDate:String(p.followUpDate||fmtDate_(new Date())),Status:'PIPELINE',Remarks:String(p.remarks||''),CreatedAt:now,UpdatedAt:now};appendObj_(CFG.sheets.TASKS,obj);appendObj_(CFG.sheets.TASK_HISTORY,{ID:id_('TH'),TaskID:id,SenderEmail:actor.Email,SenderName:actor.Name,Message:String(p.remarks||'Task assigned.'),Type:'ASSIGNED',Timestamp:now});notifyUser_(ass,'TASK','Task & Conversations',id,'New task assigned',obj.Title);audit_(actor,'ASSIGN','Task',id,obj.AssignedTo);return ok_('Task assigned successfully.',{task:taskOut_(obj)});}
function updateTaskStatus_(p){const actor=requireUser_(p),t=findBy_(CFG.sheets.TASKS,'ID',String(p.id||''));if(!t)return fail_('Task not found.');if(Number(actorLevel_(actor))!==0&&normEmail_(t.AssignedTo)!==actor.Email&&normEmail_(t.AssignedBy)!==actor.Email)return fail_('You are not authorized to update this task.');t.Status=String(p.status||t.Status);t.Remarks=String(p.remarks||t.Remarks||'');t.UpdatedAt=iso_();upsert_(CFG.sheets.TASKS,'ID',t.ID,t);appendObj_(CFG.sheets.TASK_HISTORY,{ID:id_('TH'),TaskID:t.ID,SenderEmail:actor.Email,SenderName:actor.Name,Message:t.Remarks||('Status changed to '+t.Status),Type:'STATUS',Timestamp:t.UpdatedAt});return ok_('Task status updated successfully.',{task:taskOut_(t)});}
function addTaskReply_(p){const actor=requireUser_(p),t=findBy_(CFG.sheets.TASKS,'ID',String(p.taskId||''));if(!t)return fail_('Task not found.');if(Number(actorLevel_(actor))!==0&&normEmail_(t.AssignedTo)!==actor.Email&&normEmail_(t.AssignedBy)!==actor.Email)return fail_('You are not authorized to reply to this task.');const msg=String(p.message||'').trim();if(!msg)return fail_('Message cannot be empty.');const o={ID:id_('TH'),TaskID:t.ID,SenderEmail:actor.Email,SenderName:actor.Name,Message:msg.slice(0,5000),Type:'COMMENT',Timestamp:iso_()};appendObj_(CFG.sheets.TASK_HISTORY,o);return ok_('Task reply added successfully.',{reply:o});}
function getTaskHistory_(p){const actor=requireUser_(p),t=findBy_(CFG.sheets.TASKS,'ID',String(p.taskId||''));if(!t)return fail_('Task not found.');if(Number(actorLevel_(actor))!==0&&normEmail_(t.AssignedTo)!==actor.Email&&normEmail_(t.AssignedBy)!==actor.Email)return fail_('You are not authorized to view this task.');const thread=rows_(CFG.sheets.TASK_HISTORY).filter(x=>String(x.TaskID)===String(t.ID)).map(x=>({sender:x.SenderEmail,timestamp:x.Timestamp,message:x.Message,type:x.Type}));return {success:true,message:'Task history loaded.',thread};}

function getPendingUsers_(p){const actor=requireRole_(p,0);const pending=rows_(CFG.sheets.USERS).filter(u=>String(u.Status).toLowerCase()==='pending').map(safeUser_);return {success:true,message:'Pending users loaded.',pendingUsers:pending};}
function processUserApproval_(p){const actor=requireRole_(p,0),id=String(p.userId||''),u=findBy_(CFG.sheets.USERS,'ID',id);if(!u)return fail_('User not found.');u.Status=p.approve?'Active':'Rejected';u.UpdatedAt=iso_();upsert_(CFG.sheets.USERS,'ID',id,u);audit_(actor,p.approve?'APPROVE':'REJECT','User',id,u.Status);return ok_(p.approve?'User approved successfully.':'User rejected successfully.');}
function getAllUsersDirectory_(p){const actor=requireRole_(p,0);return {success:true,message:'User directory loaded.',users:rows_(CFG.sheets.USERS).filter(u=>String(u.Status).toLowerCase()!=='rejected').map(safeUser_)};}
function updateUser_(p){const actor=requireRole_(p,0),id=String(p.id||''),u=findBy_(CFG.sheets.USERS,'ID',id);if(!u)return fail_('User not found.');if(!p.name||!p.email||!p.role)return fail_('Name, email and role are required.');const newEmail=normEmail_(p.email);if(CFG.roles.indexOf(String(p.role))<0)return fail_('Invalid role.');u.Name=String(p.name);u.Email=newEmail;u.Phone=String(p.phone||'');u.Role=String(p.role);u.States=join_(p.states);u.Regions=join_(p.regions);u.Districts=join_(p.districts);if(p.tehsils!==undefined)u.Tehsils=join_(p.tehsils);if(p.cities!==undefined)u.Cities=join_(p.cities);if(p.pinCodes!==undefined)u.PinCodes=join_(p.pinCodes);if(String(p.password||''))u.PasswordHash=hash_(p.password);u.UpdatedAt=iso_();upsert_(CFG.sheets.USERS,'ID',id,u);audit_(actor,'UPDATE','User',id,JSON.stringify({email:newEmail,role:u.Role}));return ok_('User updated successfully.',{user:safeUser_(u)});}
function deleteUser_(p){const actor=requireRole_(p,0),id=String(p.id||'');const u=findBy_(CFG.sheets.USERS,'ID',id);if(!u)return fail_('User not found.');if(normEmail_(u.Email)===actor.Email)return fail_('You cannot delete your own account.');u.Status='Inactive';u.UpdatedAt=iso_();upsert_(CFG.sheets.USERS,'ID',id,u);audit_(actor,'DEACTIVATE','User',id,u.Email);return ok_('User account deactivated successfully.');}

function markAttendance_(p){const actor=requireUser_(p);const email=normEmail_(p.email||actor.Email);if(email!==actor.Email&&Number(actorLevel_(actor))!==0)return fail_('You can only mark your own attendance.');const date=fmtDate_(new Date()),existing=rows_(CFG.sheets.ATTENDANCE).find(a=>normEmail_(a.UserEmail)===email&&String(a.Date)===date);if(existing)return fail_('Today\'s attendance is already recorded.');const now=new Date(), mins=now.getHours()*60+now.getMinutes();let type='FULL DAY',status='PENDING';if(now.getDay()===0||now.getDay()===6)type='WEEKEND';else if(mins>14*60)type='FULL-DAY LOSS';else if(mins>9*60+15)type='HALF DAY';const o={ID:id_('AT'),UserID:actor.ID,UserEmail:actor.Email,Name:actor.Name,Date:date,Type:type,Timestamp:now.toISOString(),Status:'PENDING',Remarks:'',Explanation:'',Latitude:String(p.latitude||''),Longitude:String(p.longitude||''),CreatedAt:now.toISOString(),UpdatedAt:now.toISOString()};appendObj_(CFG.sheets.ATTENDANCE,o);return ok_('Attendance marked successfully.',{attendance:o});}
function submitLateExplanation_(p){const actor=requireUser_(p),ex=String(p.explanation||'').trim();if(!ex)return fail_('Explanation is required.');const date=fmtDate_(new Date());const o={ID:id_('ATX'),UserID:actor.ID,UserEmail:actor.Email,Name:actor.Name,Date:date,Type:'LATE EXPLANATION',Timestamp:iso_(),Status:'PENDING',Remarks:'',Explanation:ex.slice(0,3000),Latitude:'',Longitude:'',CreatedAt:iso_(),UpdatedAt:iso_()};appendObj_(CFG.sheets.ATTENDANCE,o);return ok_('Late explanation submitted for approval.');}
function getPendingAttendanceApprovals_(p){const actor=requireUser_(p);if(Number(actorLevel_(actor))>3)return {success:true,message:'No approval access.',pending:[]};const pending=rows_(CFG.sheets.ATTENDANCE).filter(a=>String(a.Status).toUpperCase()==='PENDING').map(a=>({id:a.ID,email:a.UserEmail,timestamp:a.Timestamp,explanation:a.Explanation||a.Remarks||'',status:a.Status}));return {success:true,message:'Attendance approvals loaded.',pending};}
function processAttendanceApproval_(p){const actor=requireUser_(p),a=findBy_(CFG.sheets.ATTENDANCE,'ID',String(p.attendanceId||''));if(!a)return fail_('Attendance approval record not found.');if(Number(actorLevel_(actor))>3)return fail_('You are not authorized to process attendance approvals.');if(String(p.action||'')==='PASS_TO_HEADS'){a.Status='PENDING HEAD REVIEW';a.Remarks='Passed by '+actor.Name;a.UpdatedAt=iso_();upsert_(CFG.sheets.ATTENDANCE,'ID',a.ID,a);return ok_('Attendance request passed to heads.');}a.Status=p.approve?'APPROVED':'REJECTED';a.Remarks=String(p.remarks||'');a.UpdatedAt=iso_();upsert_(CFG.sheets.ATTENDANCE,'ID',a.ID,a);return ok_(p.approve?'Attendance approved successfully.':'Attendance rejected successfully.');}

function submitApprovalRequest_(p){const actor=requireUser_(p);const o={ID:id_('AR'),RequestedByEmail:actor.Email,RequestedByName:actor.Name,RequestedByRole:actor.Role,TargetApproverRole:String(p.targetApproverRole||''),Category:String(p.category||''),Details:String(p.details||''),Status:'PENDING',ApproverEmail:'',ApproverName:'',ApproverRemarks:'',CreatedAt:iso_(),UpdatedAt:iso_()};appendObj_(CFG.sheets.APPROVALS,o);return ok_('Approval request submitted successfully.',{request:o});}
function processGeneralApproval_(p){const actor=requireUser_(p),r=findBy_(CFG.sheets.APPROVALS,'ID',String(p.requestId||''));if(!r)return fail_('Approval request not found.');if(actor.Role!=='Master Admin'&&actor.Role!==r.TargetApproverRole)return fail_('You are not the designated approver.');r.Status=p.approve?'APPROVED':'REJECTED';r.ApproverEmail=actor.Email;r.ApproverName=actor.Name;r.ApproverRemarks=String(p.remarks||'');r.UpdatedAt=iso_();upsert_(CFG.sheets.APPROVALS,'ID',r.ID,r);return ok_(p.approve?'Request approved successfully.':'Request rejected successfully.');}

function getChatUsers_(p){const actor=requireUser_(p),users=safeUsersFor_(actor).filter(u=>u.email!==actor.Email).map(u=>{u.online=isOnline_(u.onlineAt);return {id:u.id,name:u.name,role:u.role,territory:uTerritory_(u),online:u.online};});return {success:true,message:'Chat users loaded.',users};}
function getChatMessages_(p){const actor=requireUser_(p),peerId=String(p.peerId||''),peer=findBy_(CFG.sheets.USERS,'ID',peerId);if(!peer)return fail_('Chat user not found.');const msgs=rows_(CFG.sheets.CHAT).filter(m=>(normEmail_(m.SenderEmail)===actor.Email&&String(m.ReceiverID)===peerId)||(normEmail_(m.ReceiverEmail)===actor.Email&&String(m.SenderID)===peerId)).sort((a,b)=>String(a.Timestamp).localeCompare(String(b.Timestamp))).slice(-200).map(m=>({senderId:m.SenderID,senderEmail:m.SenderEmail,message:m.Message,timestamp:m.Timestamp,read:String(m.Read).toLowerCase()==='true'}));rows_(CFG.sheets.CHAT).filter(m=>normEmail_(m.ReceiverEmail)===actor.Email&&String(m.SenderID)===peerId).forEach(m=>{m.Read=true;upsert_(CFG.sheets.CHAT,'ID',m.ID,m);});return {success:true,message:'Chat messages loaded.',messages:msgs};}
function sendChatMessage_(p){const actor=requireUser_(p),peer=findBy_(CFG.sheets.USERS,'ID',String(p.receiverId||'')),msg=String(p.message||'').trim();if(!peer)return fail_('Recipient not found.');if(!msg)return fail_('Message cannot be empty.');const now=iso_(),o={ID:id_('CH'),SenderID:actor.ID,SenderEmail:actor.Email,SenderName:actor.Name,ReceiverID:peer.ID,ReceiverEmail:peer.Email,Message:msg.slice(0,5000),Timestamp:now,Read:false,CreatedAt:now};appendObj_(CFG.sheets.CHAT,o);notifyUser_(peer,'CHAT','Team Chat',o.ID,'New team message','Message from '+actor.Name);return ok_('Message sent successfully.',{message:o});}
function touchUserActivity_(p){const actor=requireUser_(p),u=findBy_(CFG.sheets.USERS,'Email',actor.Email);if(u){u.OnlineAt=iso_();upsert_(CFG.sheets.USERS,'ID',u.ID,u);}return ok_('Activity updated.');}

function getLocationData_(p){const loc=rows_(CFG.sheets.LOCATION);const out={states:[],regions:{},districts:{},tehsils:{},cities:{},pincodes:{},rows:loc};loc.forEach(r=>{if(r.State&&!out.states.includes(String(r.State)))out.states.push(String(r.State));const s=String(r.State||'');if(s&&!out.regions[s])out.regions[s]=[];if(s&&r.Region&&!out.regions[s].includes(String(r.Region)))out.regions[s].push(String(r.Region));if(s&&!out.districts[s])out.districts[s]=[];if(s&&r.District&&!out.districts[s].includes(String(r.District)))out.districts[s].push(String(r.District));if(s&&!out.tehsils[s])out.tehsils[s]={};if(s&&r.District){out.tehsils[s][r.District]=out.tehsils[s][r.District]||[];if(r.Tehsil&&!out.tehsils[s][r.District].includes(String(r.Tehsil)))out.tehsils[s][r.District].push(String(r.Tehsil));}if(s&&!out.cities[s])out.cities[s]={};if(s&&r.District&&r.Tehsil){out.cities[s][r.District]=out.cities[s][r.District]||{};out.cities[s][r.District][r.Tehsil]=out.cities[s][r.District][r.Tehsil]||[];if(r.CityVillage&&!out.cities[s][r.District][r.Tehsil].includes(String(r.CityVillage)))out.cities[s][r.District][r.Tehsil].push(String(r.CityVillage));}});return {success:true,message:'Location master loaded.',locations:out};}
function lookupPinCode_(p){const pin=digits_(p.pin);if(pin.length!==6)return fail_('Enter a valid 6-digit PIN code.');try{const res=UrlFetchApp.fetch('https://api.postalpincode.in/pincode/'+encodeURIComponent(pin),{muteHttpExceptions:true});const data=JSON.parse(res.getContentText()||'[]');const po=data[0]?.PostOffice?.[0];if(!po)return fail_('PIN Code not found.');const d={pin:pin,state:po.State||'',district:po.District||'',city:po.Block||po.Division||po.Region||'',tehsil:po.Block||'',office:po.Name||''};return {success:true,message:'PIN Code found.',data:d};}catch(e){return fail_('PIN lookup service is temporarily unavailable.');}}
function refreshLGDLocationMaster_(p){requireRole_(p,0);seedLocations_();return ok_('India location master refreshed with the built-in state and regional master. Add detailed district/tehsil/city rows to LocationMaster as needed.');}

function requireUser_(p){const q=(p&&p.user&&typeof p.user==='object')?p.user:(p||{});if(!q.email||!q.authToken)throw new Error('Authentication required.');const email=normEmail_(q.email),token=verifyToken_(q.authToken);if(!token||normEmail_(token.email)!==email)throw new Error('Session expired. Please log in again.');const u=findBy_(CFG.sheets.USERS,'Email',email);if(!u||String(u.Status)!=='Active')throw new Error('Unauthorized user.');return u;}
function requireRole_(p,maxLevel){const u=requireUser_(p);if(actorLevel_(u)>maxLevel)throw new Error('Insufficient authorization.');return u;}
function actorLevel_(u){const n=CFG.roles.indexOf(String(u.Role));return n<0?99:n;}
function canSeeClient_(a,c){if(actorLevel_(a)===0)return true;if(sameEmail_(c.CreatedBy,a.Email))return true;const states=split_(a.States),regions=split_(a.Regions),districts=split_(a.Districts);if(actorLevel_(a)===1)return true;if(actorLevel_(a)===2)return states.length===0||states.indexOf(String(c.State))>=0;if(actorLevel_(a)===3)return (regions.length&&regions.indexOf(String(c.Region))>=0)||(states.length&&states.indexOf(String(c.State))>=0);if(actorLevel_(a)>=4)return sameEmail_(c.CreatedBy,a.Email)|| (districts.length&&districts.indexOf(String(c.District))>=0);return false;}
function canSeeCold_(a,c){if(actorLevel_(a)===0)return true;if(sameEmail_(c.UserEmail,a.Email))return true;if(actorLevel_(a)===1)return true;if(actorLevel_(a)===2)return split_(a.States).indexOf(String(c.State))>=0;if(actorLevel_(a)===3)return split_(a.Regions).indexOf(String(c.Region))>=0||split_(a.States).indexOf(String(c.State))>=0;return false;}
function filterClients_(arr,a){return arr.filter(x=>canSeeClient_(a,x)).map(clientOut_);}
function filterCold_(arr,a){return arr.filter(x=>canSeeCold_(a,x)).map(coldOut_);}
function filterTasks_(arr,a){if(actorLevel_(a)===0)return arr.map(taskOut_);const visible=new Set(rows_(CFG.sheets.CLIENTS).filter(x=>canSeeClient_(a,x)).map(x=>String(x.ID)));return arr.filter(t=>sameEmail_(t.AssignedTo,a.Email)||sameEmail_(t.AssignedBy,a.Email)||visible.has(String(t.LeadID||''))).map(taskOut_);}
function filterApprovals_(arr,a){return arr.filter(r=>actorLevel_(a)===0||sameEmail_(r.RequestedByEmail,a.Email)||String(r.TargetApproverRole)===String(a.Role)).map(approvalOut_);}
function filterSales_(arr,a){if(actorLevel_(a)<=3)return arr.filter(s=>canSeeSales_(a,s)).map(salesOut_);return arr.filter(s=>sameEmail_(s.UserEmail,a.Email)).map(salesOut_);}
function canSeeSales_(a,s){if(actorLevel_(a)===0)return true;if(sameEmail_(s.UserEmail,a.Email))return true;if(actorLevel_(a)===1)return true;if(actorLevel_(a)===2){const c=findBy_(CFG.sheets.CLIENTS,'ID',s.ClientID);return c?canSeeClient_(a,c):true;}if(actorLevel_(a)===3){const c=findBy_(CFG.sheets.CLIENTS,'ID',s.ClientID);return c?canSeeClient_(a,c):false;}return false;}
function safeUsersFor_(a){return rows_(CFG.sheets.USERS).filter(u=>String(u.Status)==='Active' && (actorLevel_(a)===0||actorLevel_(u)>=4||normEmail_(u.Email)===a.Email)).map(safeUser_);}

function safeUser_(u){return {id:u.ID,name:u.Name,email:u.Email,phone:u.Phone||'',role:u.Role,designation:u.Role,states:split_(u.States),regions:split_(u.Regions),districts:split_(u.Districts),tehsils:split_(u.Tehsils),cities:split_(u.Cities),pinCodes:split_(u.PinCodes),level:actorLevel_(u),canCreateUsers:u.Role==='Master Admin',onlineAt:u.OnlineAt||'',password:''};}
function clientOut_(c){return {id:c.ID,organizationId:c.OrganizationID||c.ID,orgName:c.Organization||'',contactName:c.ClientName||'',mobile:c.Mobile||'',officeMobile:c.OfficeMobile||'',email:c.Email||'',gstin:c.GSTIN||'',address:c.Address||'',state:c.State||'',region:c.Region||'',district:c.District||'',tehsil:c.Tehsil||'',city:c.CityVillage||'',pinCode:c.PinCode||'',salesOrderStatus:c.SalesOrderStatus||'',salesType:c.SalesType||'',leadTemp:c.LeadTemperature||'Warm',visitType:c.VisitType||'',visitCount:Number(c.VisitCount||0),visitStatus:c.VisitStatus||'',meetingStatus:c.MeetingStatus||'',callStatus:c.CallStatus||'',requirements:c.Requirements||'',remarks:c.Remarks||'',entryDate:c.EntryDate||'',activityDate:c.ActivityDate||'',latitude:c.Latitude||'',longitude:c.Longitude||'',mapUrl:c.MapUrl||'',locationImageUrl:c.LocationImageUrl||'',userId:c.CreatedBy||'',userEmail:c.CreatedBy||'',createdBy:c.CreatedBy||'',createdAt:c.CreatedAt||'',updatedAt:c.UpdatedAt||''};}
function coldOut_(c){return {id:c.ID,clientId:c.ClientID||'',orgName:c.Organization||'',clientName:c.ClientName||'',contactName:c.ContactName||'',phone:c.Phone||'',officePhone:c.OfficePhone||'',email:c.Email||'',state:c.State||'',region:c.Region||'',district:c.District||'',tehsil:c.Tehsil||'',city:c.CityVillage||'',pinCode:c.PinCode||'',address:c.Address||'',visitType:c.VisitType||'',visitCount:Number(c.VisitCount||0),reason:c.Reason||'',callStatus:c.CallStatus||'',outcome:c.Outcome||'',leadTemp:c.LeadTemperature||'Cold',requirements:c.Requirements||'',remarks:c.Remarks||'',entryDate:c.EntryDate||'',timestamp:c.CreatedAt||'',userId:c.UserID||'',userEmail:c.UserEmail||''};}
function salesOut_(s){return {id:s.ID,userId:s.UserID||'',userEmail:s.UserEmail||'',userRole:s.UserRole||'',clientId:s.ClientID||'',organizationId:s.OrganizationID||'',orgName:s.Organization||'',clientName:s.ClientName||'',contactNumber:s.ContactNumber||'',piStatus:s.PIStatus||'',orderFor:s.OrderFor||'',salesType:s.SalesType||'',pvType:s.PVType||'',panelQuantity:Number(s.PanelQuantity||0),panelWp:Number(s.PanelWP||0),ratePerWp:Number(s.RatePerWP||0),lpd:Number(s.LPD||0),ratePerLpd:Number(s.RatePerLPD||0),salesLpd:Number(s.SalesLPD||0),projectKw:Number(s.ProjectKW||0),ratePerKw:Number(s.RatePerKW||0),kw:Number(s.KW||0),amount:Number(s.Amount||0),repeatCount:Math.max(1,Number(s.RepeatCount||1)),date:s.Date||'',remarks:s.Remarks||'',createdBy:s.CreatedBy||'',createdAt:s.CreatedAt||'',updatedAt:s.UpdatedAt||''};}
function taskOut_(t){return {id:t.ID,title:t.Title,assignedBy:t.AssignedBy,assignedTo:t.AssignedTo,followUpDate:t.FollowUpDate,status:t.Status,remarks:t.Remarks||'',createdAt:t.CreatedAt,updatedAt:t.UpdatedAt};}
function approvalOut_(r){return {id:r.ID,requestedByEmail:r.RequestedByEmail,requestedByName:r.RequestedByName,requestedByRole:r.RequestedByRole,targetApproverRole:r.TargetApproverRole,category:r.Category,details:r.Details,status:r.Status,approverEmail:r.ApproverEmail,approverName:r.ApproverName,approverRemarks:r.ApproverRemarks,createdAt:r.CreatedAt,updatedAt:r.UpdatedAt};}

function notify_(actor,type,module,recordId,title,message){const users=rows_(CFG.sheets.USERS).filter(u=>String(u.Status)==='Active');users.forEach(u=>{if(u.ID!==actor.ID&&actorLevel_(u)<=actorLevel_(actor)+1)notifyUser_(u,type,module,recordId,title,message);});}
function notifyUser_(u,type,module,recordId,title,message){appendObj_(CFG.sheets.NOTIFICATIONS,{ID:id_('N'),UserID:u.ID,UserEmail:u.Email,Type:type,Module:module,RecordID:recordId,Title:title,Message:String(message||''),Timestamp:iso_(),Read:false,CreatedAt:iso_()});}
function audit_(actor,action,entity,id,details){appendObj_(CFG.sheets.AUDIT,{ID:id_('X'),ActorID:actor.ID,ActorEmail:actor.Email,ActorRole:actor.Role,Action:action,Entity:entity,RecordID:id,Details:String(details||'').slice(0,5000),Timestamp:iso_()});}

function rows_(name){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);if(!sh||sh.getLastRow()<2)return [];const vals=sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getValues(),h=vals[0].map(String);return vals.slice(1).filter(r=>r.some(v=>v!==''&&v!==null)).map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i] instanceof Date?r[i].toISOString():r[i]);return o;});}
function appendObj_(name,obj){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);if(!sh)throw new Error('Sheet missing: '+name);const h=HEADERS[name];sh.appendRow(h.map(k=>obj[k]===undefined?'':obj[k]));}
function upsert_(name,key,id,obj){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);if(!sh)throw new Error('Sheet missing: '+name);const h=HEADERS[name],vals=sh.getDataRange().getValues(),ki=h.indexOf(key);for(let r=1;r<vals.length;r++){if(String(vals[r][ki]).toLowerCase()===String(id).toLowerCase()){sh.getRange(r+1,1,1,h.length).setValues([h.map(k=>obj[k]===undefined?'':obj[k])]);return;}}appendObj_(name,obj);}
function findBy_(name,key,val){const v=String(val||'').toLowerCase();return rows_(name).find(x=>String(x[key]||'').toLowerCase()===v)||null;}
function deleteBy_(name,key,val){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);if(!sh)return;const h=HEADERS[name],ki=h.indexOf(key),vals=sh.getDataRange().getValues();for(let r=1;r<vals.length;r++){if(String(vals[r][ki]).toLowerCase()===String(val).toLowerCase()){sh.deleteRow(r+1);return;}}}
function userRow_(id,name,email,phone,role,states,regions,districts,tehsils,cities,pins,manager,status,password,now){return [id,name,email,phone,role,join_(states),join_(regions),join_(districts),join_(tehsils),join_(cities),join_(pins),manager,status,hash_(password),now,now,'',''];}
function issueToken_(email){const ts=Date.now(),payload=Utilities.base64EncodeWebSafe(JSON.stringify({email:normEmail_(email),ts,nonce:Utilities.getUuid()}));const sig=Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload,getSecret_()));return payload+'.'+sig;}
function verifyToken_(token){try{const a=String(token).split('.');if(a.length!==2)return null;const sig=Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(a[0],getSecret_()));if(sig!==a[1])return null;const p=JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(a[0])).getDataAsString());if(Date.now()-Number(p.ts)>CFG.sessionHours*60*60*1000)return null;return p;}catch(_){return null;}}
function getSecret_(){const ps=PropertiesService.getScriptProperties();let s=ps.getProperty('KOSOL_AUTH_SECRET');if(!s){s=Utilities.getUuid()+Utilities.getUuid();ps.setProperty('KOSOL_AUTH_SECRET',s);}return s;}
function hash_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8).map(b=>('0'+(b<0?b+256:b).toString(16)).slice(-2)).join('');}
function saveImage_(dataUrl,id){const m=String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);if(!m)return '';const bytes=Utilities.base64Decode(m[2]);if(bytes.length>CFG.imageMaxBytes)throw new Error('Image exceeds 5 MB.');const props=PropertiesService.getScriptProperties(),folderId=props.getProperty('KOSOL_IMAGE_FOLDER_ID');let folder=folderId?DriveApp.getFolderById(folderId):DriveApp.getRootFolder();const file=folder.createFile(Utilities.newBlob(bytes,m[1],id+'.jpg'));file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);return file.getUrl();}
function seedLocations_(){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.sheets.LOCATION);if(!sh)return;const existing=rows_(CFG.sheets.LOCATION);if(existing.length)return;const now=iso_();const map={Maharashtra:['Vidarbha','Marathwada','Western Maharashtra','Khandesh'],Gujarat:['Ahmedabad','Vadodara','Surat','Rajkot','Kutch'],Karnataka:['Bengaluru','Mysuru','Hubballi-Dharwad','Mangaluru','Belagavi'],'Madhya Pradesh':['Bhopal','Indore','Jabalpur','Gwalior','Ujjain'],Rajasthan:['Jaipur','Jodhpur','Udaipur','Kota','Bikaner'],Delhi:['Delhi NCR'],'Uttar Pradesh':['Lucknow','Kanpur','Noida','Varanasi','Agra'],Telangana:['Hyderabad','Warangal','Karimnagar'],'Tamil Nadu':['Chennai','Coimbatore','Madurai','Salem'],Kerala:['Thiruvananthapuram','Kochi','Kozhikode'],'West Bengal':['Kolkata','Siliguri','Durgapur'],Odisha:['Bhubaneswar','Cuttack','Rourkela'],Punjab:['Ludhiana','Amritsar','Jalandhar'],Haryana:['Gurugram','Faridabad','Hisar'],Bihar:['Patna','Gaya','Muzaffarpur'],Jharkhand:['Ranchi','Jamshedpur','Dhanbad'],Chhattisgarh:['Raipur','Bilaspur','Durg'],Assam:['Guwahati','Dibrugarh'],Uttarakhand:['Dehradun','Haridwar'],'Himachal Pradesh':['Shimla','Kangra'],Goa:['Panaji','Margao'],'Jammu and Kashmir':['Jammu','Srinagar'],'Andhra Pradesh':['Visakhapatnam','Vijayawada','Tirupati'],Chandigarh:['Chandigarh']};const out=[];Object.keys(map).forEach(s=>map[s].forEach(r=>out.push(['STATE_REGION',s,r,'','','','','',now])));if(out.length)sh.getRange(2,1,out.length,HEADERS.LocationMaster.length).setValues(out);}
function cleanOnline_(){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.sheets.USERS);if(!sh||sh.getLastRow()<2)return;const vals=sh.getDataRange().getValues(),h=vals[0],idx=h.indexOf('OnlineAt');if(idx<0)return;const cutoff=Date.now()-10*60*1000;for(let i=1;i<vals.length;i++){const t=Date.parse(vals[i][idx]);if(t&&t<cutoff)vals[i][idx]='';}sh.getRange(2,1,vals.length-1,sh.getLastColumn()).setValues(vals.slice(1));}
function isOnline_(v){const t=Date.parse(String(v||''));return !!t&&(Date.now()-t<10*60*1000);}
function uTerritory_(u){const s=split_(u.states),r=split_(u.regions);return (s.join(', ')||'All India')+' - '+(r.join(', ')||'All Regions');}
function split_(v){return String(v||'').split(/[|,]/).map(x=>x.trim()).filter(Boolean);}
function join_(v){return (Array.isArray(v)?v:[v]).flatMap(x=>String(x||'').split(/[|,]/)).map(x=>x.trim()).filter(Boolean).filter((x,i,a)=>a.indexOf(x)===i).join(' | ');}
function normEmail_(v){return String(v||'').trim().toLowerCase();}
function sameEmail_(a,b){return normEmail_(a)===normEmail_(b);}
function digits_(v){return String(v||'').replace(/\D/g,'');}
function id_(p){return p+Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase();}
function iso_(){return new Date().toISOString();}
function fmtDate_(d){return Utilities.formatDate(d,CFG.timezone,'yyyy-MM-dd');}
function ok_(message,data){return Object.assign({success:true,message:message||'Operation completed.'},data||{});}
function fail_(message){return {success:false,message:String(message||'Operation failed.')};}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
