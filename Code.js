/*** Warfarin Food-Drug Interaction API — Google Apps Script ***/
const TZ = 'Asia/Bangkok';

const SCHEMA = {
  PATIENTS:        ['AN','HN','FullName','Ward','Bed','AdmitDate','LastSeenDate','ICD10','ICD10Desc','Status'],
  MED_RECON:       ['AN','DrugName','GenericName','DrugGroup','Dose','Route','Source','StartDate'],
  MED_INWARD:      ['AN','DrugName','GenericName','DrugGroup','Dose','Route','OrderDate','Active'],
  ALLERGY_DRUG:    ['AN','DrugName','DrugGroup','Reaction','Severity','ReportDate'],
  ALLERGY_FOOD:    ['AN','FoodName','Symptom','Severity','ReportDate'],
  OPERATIONS:      ['AN','OperationCode','OperationName','OperationDate','Surgeon','Anesthesia','Note'],
  FOOD_MAP:        ['FoodKey','FoodNameTH','Synonyms','RiskLevel','Effect','VitaminK_mcg','PopupMessage'],
  DRUG_PAIR_MAP:   ['DrugA','DrugB','PopupMessage'],
  MEAL_ORDERS:     ['OrderId','AN','MealType','MealDate','MenuText','DietType','OrderBy'],
  ALERT_LOG:       ['AlertId','Timestamp','AN','PatientName','Ward','Bed','MealType','MenuText',
                    'MatchedFoods','RiskLevel','AlertMessage','ActionBy','Result'],
  SYSTEM_LOG:      ['LogId','Timestamp','User','Action','Target','Detail','Status','IP']
};

function getSS_(){ return SpreadsheetApp.getActive(); }

function sheet_(name){
  const ss = getSS_();
  let sh = ss.getSheetByName(name);
  if (!sh){
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,SCHEMA[name].length).setValues([SCHEMA[name]])
      .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll_(name){
  const sh = sheet_(name), lr = sh.getLastRow();
  if (lr < 2) return [];
  const head = SCHEMA[name];
  return sh.getRange(2,1,lr-1,head.length).getValues().map(r=>{
    const o={}; head.forEach((h,i)=> o[h] = (r[i] instanceof Date)
      ? Utilities.formatDate(r[i],TZ,'yyyy-MM-dd') : r[i]);
    return o;
  }).filter(o=> String(o[head[0]]).trim() !== '');
}

function append_(name, obj){
  const sh = sheet_(name);
  sh.appendRow(SCHEMA[name].map(h=> obj[h] !== undefined ? obj[h] : ''));
}

function now_(){ return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function uid_(p){ return p + '-' + new Date().getTime() + '-' + Math.floor(Math.random()*900+100); }

function log_(user, action, target, detail, status){
  append_('SYSTEM_LOG', {
    LogId: uid_('LOG'), Timestamp: now_(), User: user||'system', Action: action,
    Target: target||'', Detail: detail||'', Status: status||'SUCCESS', IP: ''
  });
}

/* ---------------- Router ---------------- */
function doGet(e){  return handle_(e.parameter || {}); }
function doPost(e){
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch(err){ p = e.parameter || {}; }
  return handle_(p);
}

function out_(obj, cb){
  if (cb) return ContentService.createTextOutput(cb+'('+JSON.stringify(obj)+')')
                 .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(JSON.stringify(obj))
             .setMimeType(ContentService.MimeType.JSON);
}

function handle_(p){
  const cb = p.callback, user = p.user || 'anonymous';
  try {
    let data;
    switch (p.action) {
      case 'bootstrap':
        data = {
          patients: readAll_('PATIENTS'),
          foodMap:  readAll_('FOOD_MAP'),
          drugPair: readAll_('DRUG_PAIR_MAP')
        };
        log_(user,'BOOTSTRAP','', 'load master data','SUCCESS');
        break;

      case 'patient': {           // p.an
        const an = String(p.an||'').trim();
        const pt = readAll_('PATIENTS').filter(x=> String(x.AN)===an)[0] || null;
        data = {
          patient: pt,
          medRecon:    readAll_('MED_RECON').filter(x=> String(x.AN)===an),
          medInWard:   readAll_('MED_INWARD').filter(x=> String(x.AN)===an),
          allergyDrug: readAll_('ALLERGY_DRUG').filter(x=> String(x.AN)===an),
          allergyFood: readAll_('ALLERGY_FOOD').filter(x=> String(x.AN)===an),
          operations:  readAll_('OPERATIONS').filter(x=> String(x.AN)===an),
          meals:       readAll_('MEAL_ORDERS').filter(x=> String(x.AN)===an),
          alerts:      readAll_('ALERT_LOG').filter(x=> String(x.AN)===an)
        };
        log_(user,'VIEW_PATIENT', an, 'open patient summary','SUCCESS');
        break;
      }

      case 'saveMeal':            // p.payload = MEAL_ORDERS row
        p.payload.OrderId = p.payload.OrderId || uid_('MO');
        append_('MEAL_ORDERS', p.payload);
        log_(user,'SAVE_MEAL', p.payload.AN, p.payload.MealType+' | '+p.payload.MenuText,'SUCCESS');
        data = { orderId: p.payload.OrderId };
        break;

      case 'saveAlert':           // p.payload = ALERT_LOG row
        p.payload.AlertId   = p.payload.AlertId || uid_('AL');
        p.payload.Timestamp = now_();
        append_('ALERT_LOG', p.payload);
        log_(user,'WARFARIN_ALERT', p.payload.AN,
             p.payload.RiskLevel+' | '+p.payload.MatchedFoods, 'ALERT');
        data = { alertId: p.payload.AlertId };
        break;

      case 'report':              // สรุปรายงานผู้ป่วยวาร์ฟาริน
        data = buildWarfarinReport_();
        log_(user,'REPORT','warfarin_ward','generate report','SUCCESS');
        break;

      case 'logs':
        data = readAll_('SYSTEM_LOG').slice(-300).reverse();
        break;

      case 'log':
        log_(user, p.act, p.target, p.detail, p.status);
        data = { ok:true };
        break;

      default:
        return out_({ ok:false, error:'unknown action: '+p.action }, cb);
    }
    return out_({ ok:true, data: data }, cb);
  } catch (err){
    log_(user, p.action||'UNKNOWN','', String(err), 'ERROR');
    return out_({ ok:false, error: String(err) }, cb);
  }
}

/* ------- รายงาน: ผู้ป่วยที่ได้รับยากลุ่มวาร์ฟารินและเข้าหอ ------- */
function isWarfarin_(s){ return /warfarin|วาร์?ฟาริน|orfarin|coumadin|fargem/i.test(String(s||'')); }

function buildWarfarinReport_(){
  const pts = readAll_('PATIENTS');
  const meds = readAll_('MED_INWARD').concat(readAll_('MED_RECON'));
  const alerts = readAll_('ALERT_LOG');
  const rows = [];
  pts.forEach(pt=>{
    const w = meds.filter(m=> String(m.AN)===String(pt.AN) &&
                (isWarfarin_(m.DrugName)||isWarfarin_(m.GenericName)||isWarfarin_(m.DrugGroup)));
    if (!w.length) return;
    const a = alerts.filter(x=> String(x.AN)===String(pt.AN));
    rows.push({
      AN: pt.AN, FullName: pt.FullName, Ward: pt.Ward, Bed: pt.Bed,
      AdmitDate: pt.AdmitDate, LastSeenDate: pt.LastSeenDate,
      ICD10: pt.ICD10 + ' ' + (pt.ICD10Desc||''),
      Drugs: w.map(x=> x.DrugName + (x.Dose? ' '+x.Dose:'')).join(', '),
      AlertCount: a.length,
      HighRisk: a.filter(x=> x.RiskLevel==='HIGH').length,
      LastAlert: a.length ? a[a.length-1].Timestamp : '-'
    });
  });
  return { generatedAt: now_(), total: rows.length, rows: rows };
}

/* =======================================================
   รันครั้งเดียว: สร้างชีตทั้งหมด + ใส่ข้อมูลตัวอย่าง/ตาราง Map
   ======================================================= */
function SETUP_AND_SEED(){
  Object.keys(SCHEMA).forEach(k=> sheet_(k));

  // ---- ตาราง Map คู่ยา (ตามไฟล์ popUp drug interaction) ----
  append_('DRUG_PAIR_MAP', { DrugA:'Enalapril', DrugB:'Losartan',
    PopupMessage:'ACEIs ใช้คู่กับ ARBs ทำให้เกิด Hyperkalemia' });

  // ---- ตาราง Map อาหาร x Warfarin (แก้ไข/เพิ่มได้ในชีต) ----
  const FOODS = [
    ['KALE','คะน้า','ผักคะน้า|kale|คะน้าฮ่องกง','HIGH','Vitamin K สูงมาก → ลดฤทธิ์วาร์ฟาริน (INR ลด)',817],
    ['SPINACH','ปวยเล้ง/ผักโขม','ปวยเล้ง|ผักโขม|spinach','HIGH','Vitamin K สูงมาก → ลดฤทธิ์วาร์ฟาริน',483],
    ['BROCCOLI','บรอกโคลี','บล็อคโคลี่|บรอกโคลี|broccoli','HIGH','Vitamin K สูง → ลดฤทธิ์วาร์ฟาริน',141],
    ['CABBAGE','กะหล่ำปลี','กะหล่ำ|cabbage|กะหล่ำดาว','HIGH','Vitamin K สูง → ลดฤทธิ์วาร์ฟาริน',76],
    ['TAMLUENG','ตำลึง','ใบตำลึง|ยอดตำลึง','HIGH','Vitamin K สูง → ลดฤทธิ์วาร์ฟาริน',0],
    ['MORNING_GLORY','ผักบุ้ง','ผักบุ้งจีน|ผักบุ้งไทย','MODERATE','Vitamin K ปานกลาง → เฝ้าระวัง INR',0],
    ['CHAOM','ชะอม','ยอดชะอม','MODERATE','Vitamin K ปานกลาง → เฝ้าระวัง INR',0],
    ['LETTUCE','ผักกาดหอม','สลัด|เลตทิส|lettuce|ผักสลัด','MODERATE','Vitamin K ปานกลาง',126],
    ['CELERY','ขึ้นฉ่าย','คื่นช่าย|celery','MODERATE','Vitamin K + สาร Coumarin',29],
    ['PARSLEY','ผักชีฝรั่ง','พาสลี่ย์|parsley|ผักชีลาว','HIGH','Vitamin K สูงมาก',1640],
    ['SOYBEAN','ถั่วเหลือง/น้ำมันถั่วเหลือง','เต้าหู้|นมถั่วเหลือง|น้ำเต้าหู้|soy','MODERATE','Vitamin K ในน้ำมันถั่วเหลือง',0],
    ['SEAWEED','สาหร่าย','สาหร่ายทะเล|โนริ|วากาเมะ','MODERATE','Vitamin K สูง',0],
    ['LIVER','ตับ','ตับหมู|ตับไก่|ตับบด','HIGH','Vitamin K สูง → ลดฤทธิ์วาร์ฟาริน',0],
    ['GREEN_TEA','ชาเขียว','ใบชา|matcha|มัทฉะ','HIGH','Vitamin K สูง → ลดฤทธิ์วาร์ฟาริน',0],
    ['GARLIC','กระเทียม','กระเทียมเจียว|garlic','MODERATE','ต้านเกล็ดเลือด → เพิ่มความเสี่ยงเลือดออก',0],
    ['GINGER','ขิง','น้ำขิง|ginger','MODERATE','ต้านเกล็ดเลือด → เพิ่มความเสี่ยงเลือดออก',0],
    ['TURMERIC','ขมิ้น','ขมิ้นชัน|turmeric|แกงเหลือง','MODERATE','เพิ่มฤทธิ์วาร์ฟาริน (INR เพิ่ม)',0],
    ['GINKGO','แปะก๊วย','ginkgo','HIGH','เพิ่มความเสี่ยงเลือดออก',0],
    ['CRANBERRY','แครนเบอร์รี่','น้ำแครนเบอร์รี่|cranberry','HIGH','ยับยั้ง CYP2C9 → INR เพิ่ม',0],
    ['GRAPEFRUIT','เกรปฟรุต','ส้มโอฝรั่ง|grapefruit','MODERATE','ยับยั้ง CYP → INR เพิ่ม',0],
    ['MANGO','มะม่วงสุก','มะม่วงน้ำดอกไม้','MODERATE','เพิ่มฤทธิ์วาร์ฟาริน (INR เพิ่ม)',0],
    ['DURIAN','ทุเรียน','durian','MODERATE','มีแอลกอฮอล์/ซัลเฟอร์ → รบกวนการเมตาบอลิซึม',0],
    ['ALCOHOL','แอลกอฮอล์','เหล้า|เบียร์|ไวน์|สุรา','HIGH','รบกวนการเมตาบอลิซึม → INR แกว่ง',0],
    ['FISH_OIL','น้ำมันปลา','fish oil|โอเมก้า3','MODERATE','ต้านเกล็ดเลือด → เสี่ยงเลือดออก',0],
    ['GINSENG','โสม','ginseng|โสมเกาหลี','HIGH','ลดฤทธิ์วาร์ฟาริน',0]
  ];
  FOODS.forEach(f=> append_('FOOD_MAP', {
    FoodKey:f[0], FoodNameTH:f[1], Synonyms:f[2], RiskLevel:f[3], Effect:f[4],
    VitaminK_mcg:f[5],
    PopupMessage:'คนไข้รายนี้ได้รับยากลุ่มวาร์ฟาริน — พบ "'+f[1]+'" ในเมนู ('+f[4]+')'
  }));

  // ---- ข้อมูลตัวอย่างผู้ป่วย ----
  [['AN6900125','HN0451233','นางสมศรี ใจดี','อายุรกรรมหญิง 1','12','2026-08-25','2026-09-01','I48.0','Atrial fibrillation','ADMIT'],
   ['AN6900126','HN0398112','นายวิชัย ทองสุข','ศัลยกรรมชาย 2','07','2026-08-28','2026-09-01','I26.9','Pulmonary embolism','ADMIT'],
   ['AN6900127','HN0512004','นางสาวกัญญา แสงเดือน','อายุรกรรมหญิง 1','03','2026-08-30','2026-09-01','I80.2','Deep vein thrombosis','ADMIT'],
   ['AN6900128','HN0477865','นายประเสริฐ มั่นคง','CVT Ward','01','2026-08-20','2026-09-01','Z95.2','Prosthetic heart valve','ADMIT'],
   ['AN6900129','HN0455001','นางมาลี ศรีสุข','อายุรกรรมชาย 3','15','2026-08-29','2026-09-01','E11.9','Type 2 DM','ADMIT']
  ].forEach(r=> append_('PATIENTS', {
    AN:r[0],HN:r[1],FullName:r[2],Ward:r[3],Bed:r[4],AdmitDate:r[5],
    LastSeenDate:r[6],ICD10:r[7],ICD10Desc:r[8],Status:r[9] }));

  [['AN6900125','Warfarin 3 mg','Warfarin sodium','Anticoagulant','3 mg','PO','OPD Card','2026-08-25'],
   ['AN6900125','Enalapril 5 mg','Enalapril maleate','ACEIs','5 mg','PO','OPD Card','2026-08-25'],
   ['AN6900126','Orfarin 5 mg','Warfarin sodium','Anticoagulant','5 mg','PO','Home Med','2026-08-28'],
   ['AN6900128','Warfarin 2 mg','Warfarin sodium','Anticoagulant','2 mg','PO','OPD Card','2026-08-20'],
   ['AN6900129','Metformin 500 mg','Metformin','Biguanide','500 mg','PO','OPD Card','2026-08-29']
  ].forEach(r=> append_('MED_RECON',{AN:r[0],DrugName:r[1],GenericName:r[2],DrugGroup:r[3],
      Dose:r[4],Route:r[5],Source:r[6],StartDate:r[7]}));

  [['AN6900125','Warfarin 3 mg','Warfarin sodium','Anticoagulant','3 mg OD','PO','2026-08-25','YES'],
   ['AN6900125','Losartan 50 mg','Losartan potassium','ARBs','50 mg OD','PO','2026-08-30','YES'],
   ['AN6900126','Warfarin 5 mg','Warfarin sodium','Anticoagulant','5 mg OD','PO','2026-08-28','YES'],
   ['AN6900127','Warfarin 4 mg','Warfarin sodium','Anticoagulant','4 mg OD','PO','2026-08-30','YES'],
   ['AN6900128','Warfarin 2 mg','Warfarin sodium','Anticoagulant','2 mg OD','PO','2026-08-20','YES'],
   ['AN6900129','Insulin RI','Insulin regular','Insulin','8 unit tid','SC','2026-08-29','YES']
  ].forEach(r=> append_('MED_INWARD',{AN:r[0],DrugName:r[1],GenericName:r[2],DrugGroup:r[3],
      Dose:r[4],Route:r[5],OrderDate:r[6],Active:r[7]}));

  [['AN6900125','Penicillin','Beta-lactam','ผื่นลมพิษทั่วตัว','Moderate','2024-05-11'],
   ['AN6900126','Aspirin','NSAIDs','หลอดลมตีบ หายใจไม่ออก','Severe','2023-11-02'],
   ['AN6900128','Sulfamethoxazole','Sulfonamides','Stevens-Johnson Syndrome','Life-threatening','2022-03-19']
  ].forEach(r=> append_('ALLERGY_DRUG',{AN:r[0],DrugName:r[1],DrugGroup:r[2],
      Reaction:r[3],Severity:r[4],ReportDate:r[5]}));

  [['AN6900125','กุ้ง / อาหารทะเล','ผื่นคัน บวมรอบตา','Moderate','2024-05-11'],
   ['AN6900127','ถั่วลิสง','คลื่นไส้ อาเจียน ผื่นแดง','Moderate','2025-01-08'],
   ['AN6900128','นมวัว','ท้องเสีย ปวดท้อง','Mild','2021-07-30']
  ].forEach(r=> append_('ALLERGY_FOOD',{AN:r[0],FoodName:r[1],Symptom:r[2],
      Severity:r[3],ReportDate:r[4]}));

  [['AN6900128','35.21','Mitral valve replacement','2026-08-21','นพ.ธนกร ว.','General','ใส่ลิ้นหัวใจเทียมโลหะ ต้องได้ warfarin ตลอดชีพ'],
   ['AN6900126','38.7','IVC filter insertion','2026-08-29','นพ.สุชาติ ก.','Local','ป้องกัน PE ซ้ำ'],
   ['AN6900129','','ไม่มีประวัติผ่าตัด','','','','-']
  ].forEach(r=> append_('OPERATIONS',{AN:r[0],OperationCode:r[1],OperationName:r[2],
      OperationDate:r[3],Surgeon:r[4],Anesthesia:r[5],Note:r[6]}));

  log_('system','SETUP','ALL','initial setup & seed data','SUCCESS');
}