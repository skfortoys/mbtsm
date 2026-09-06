/*******************************************************************
 *  كتالوج الطفل المبتسم — محرك قاعدة البيانات (Google Apps Script)
 *  ----------------------------------------------------------------
 *  وظيفة هذا السكربت:
 *   1) يجهّز صفحات ملف جوجل شيت (الاصناف / الاقسام / الاعدادات / تعليمات)
 *   2) يستورد الأصناف من ملف catalog_items.csv (مُصدَّر من قاعدة بيانات الخوارزمي)
 *   3) يمشّط مجلدات الصور في درايف ويربط كل صنف بصورته تلقائياً
 *   4) يجعل مجلد الصور قابلاً للعرض العام (حتى تظهر الصور في الموقع)
 *
 *  طريقة التشغيل: افتح ملف الشيت > الإضافات > Apps Script > الصق هذا الملف
 *  ثم ارجع للشيت وحدّث الصفحة، بيظهر لك منيو باسم «الكتالوج».
 *******************************************************************/

/* ================== الإعدادات الثابتة ================== */

// ملف بيانات الأصناف المُصدَّر من قاعدة بيانات الخوارزمي (CSV داخل مجلد «سكاي سوفت» في درايف)
// يُحدَّث تلقائياً كل يوم بواسطة catalog_build.py
// معرف ملف جوجل شيت (يجعل السكربت يعمل مستقلاً أو مربوطاً بالملف)
var SHEET_ID = '14uP8Lv69Gw6r8qg5CwH1Jimox-U21XuHdJ6AMwRq4to';

var CATALOG_CSV_NAME = 'catalog_items.csv';
var CATALOG_CSV_ID   = '1smkD6iZeMEVN5ZNo0JVkB02oo2wIUXIb';  // فارغ = ابحث بالاسم

// مجلد «جميع الصور» — الصور مسمّاة بكود الصنف (مثال: 800-02122.jpg)
var ALL_IMAGES_FOLDER_ID = '1rhtZ5DywWjJYJ1a1-pZIjV5Y5es2i3u3';

// مجلد «تصنيف الصور في مجلدات» — كل مجلد بداخله = قسم
var CATEGORY_ROOT_FOLDER_ID = '1vGvAFfFDVandIxkhYn9bk-hdaiATNbxP';

// مجلد «الجديد» — الأصناف اللي تتعلّم علامة «جديد»
var NEW_FOLDER_ID = '1K85FdOgqr7H4eVSicqFmzmNzi_grIP_d';

var SH_ITEMS = 'الاصناف';
var SH_CATS  = 'الاقسام';
var SH_CONF  = 'الاعدادات';
var SH_HELP  = 'تعليمات';

var ITEM_HEADERS = [
  'كود الصنف',            // A
  'اسم الصنف',            // B
  'القسم',                // C
  'وحدة البيع',           // D  (حبه / BOX … كما في البرنامج)
  'سعر الحبة',            // E  سعر1 = سعر الجملة (سعر3 قطاعي ولا يُعرض)
  'عدد الحبات بالكرتون',  // F
  'سعر الكرتون',          // G  اتركه فارغ = سعر الحبة × عدد الحبات (يُحسب دائماً)
  'المخزون',              // H  الرئيسي + فرع02
  'الرئيسي',              // I
  'فرع02',                // J
  'فرع05',                // K
  'معرف الصورة',          // L
  'عرض',                  // M
  'جديد',                 // N
  'وسوم',                 // O
  'إظهار',                // P
  'ترتيب',                // Q
  'ملاحظة'                // R
];

// أعمدة تأتي من قاعدة البيانات وتُحدَّث في كل استيراد (0-based)
var DB_COLS = [1, 3, 4, 5, 7, 8, 9, 10];


/* ================== المنيو ================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('الكتالوج')
    .addItem('1) إعداد أولي للصفحات', 'setupSheets')
    .addItem('2) استيراد الأصناف من قاعدة بيانات الخوارزمي', 'importItems')
    .addItem('3) مزامنة الصور والأقسام من درايف', 'syncImages')
    .addItem('4) فتح صلاحية عرض الصور للعملاء', 'makeImagesPublic')
    .addSeparator()
    .addItem('تشغيل الكل (2 ثم 3)', 'runAll')
    .addToUi();
}

function runAll() { importItems(); syncImages(); }

/** تشغيل الكل من محرر Apps Script مباشرة (بدون منيو) */
function setupAndRunAll() {
  setupSheets();
  importItems();
  syncImages();
  makeImagesPublic();
  Logger.log('تم كل شيء ✅');
}

/** ملف الشيت — يعمل سواء شُغّل السكربت من داخل الملف أو مستقلاً */
function ss_() {
  var a = null;
  try { a = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  return a || SpreadsheetApp.openById(SHEET_ID);
}

/** رسالة تظهر في المنيو، وتُسجَّل فقط إن شُغّلت من المحرر */
function say_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

/* ================== 1) إعداد الصفحات ================== */

function setupSheets() {
  var ss = ss_();

  // --- الاصناف ---
  var it = ss.getSheetByName(SH_ITEMS) || ss.insertSheet(SH_ITEMS);
  if (it.getLastRow() === 0 || it.getRange(1,1).getValue() !== ITEM_HEADERS[0]) {
    it.getRange(1,1,1,ITEM_HEADERS.length).setValues([ITEM_HEADERS]);
  }
  it.getRange(1,1,1,ITEM_HEADERS.length)
    .setFontWeight('bold').setBackground('#1f3a93').setFontColor('#ffffff');
  it.setFrozenRows(1);
  it.setRightToLeft(true);

  // --- الاقسام ---
  var ct = ss.getSheetByName(SH_CATS) || ss.insertSheet(SH_CATS);
  if (ct.getLastRow() === 0) {
    ct.getRange(1,1,1,4).setValues([['اسم القسم','الترتيب','إظهار','أيقونة']]);
    var seed = [
      ['دراجات وسكوترات وسيارات ركوب دف ومشايات','🛴'],
      ['سيارات وطائرات ريموت','🚗'],
      ['سيارات دف وحديد','🚙'],
      ['مواقف سيارات وقطارات','🚉'],
      ['عرائس','👧'],
      ['دمى قطنية وملابس تنكرية','🧸'],
      ['مطابخ وتسريحات','🍳'],
      ['دكتور والعاب مواليد وعربيات','👶'],
      ['بيوت والعاب بناء وتركيب','🧱'],
      ['سبورات والعاب تعليم وذكاء','📚'],
      ['سلايم وصلصال والعاب تلوين','🎨'],
      ['مكياج وخرز واكسسوارات بناتي','💄'],
      ['كاشييرات ومكرفونات والعاب بطارية','🎤'],
      ['مسدسات صوت وطلقات','🔫'],
      ['مسدسات ماء وصابون','💦'],
      ['كور والعاب جماعيه ورياضية','⚽'],
      ['خيم ومحواجز مع كوره وبدون كوره','⛺'],
      ['مسابح وعوامات سباحة','🏊'],
      ['نظارات والعاب رمل وبحر','🏖'],
      ['سمك وحيوانات والعاب شخصيات','🐠'],
      ['كروت واكياس تعليق','🏷'],
      ['العاب منوعة','🎁'],
      ['ابو 5 ريال','5'],
      ['ابو 10 ريال','10']
    ].map(function (r, k) { return [r[0], k + 1, 'نعم', r[1]]; });
    ct.getRange(2,1,seed.length,4).setValues(seed);
  }
  ct.getRange(1,1,1,4).setFontWeight('bold').setBackground('#1f3a93').setFontColor('#ffffff');
  ct.setFrozenRows(1); ct.setRightToLeft(true);

  // --- الاعدادات ---
  var cf = ss.getSheetByName(SH_CONF) || ss.insertSheet(SH_CONF);
  if (cf.getLastRow() === 0) {
    cf.getRange(1,1,1,2).setValues([['المفتاح','القيمة']]);
    cf.getRange(2,1,14,2).setValues([
      ['اسم_المتجر',                'الطفل المبتسم لألعاب الأطفال'],
      ['وصف_المتجر',               'جملة ألعاب الأطفال — جدة'],
      ['رقم_واتساب',               '9665557282942'],
      ['العملة',                   'ر.س'],
      ['رسالة_ترحيب',              'أهلاً بك 👋 تصفح الأصناف وأضف طلبك ثم أرسله لنا مباشرة على واتساب'],
      ['اظهار_المخزون',            'لا'],
      ['اظهار_الاصناف_بلا_صورة',   'لا'],
      ['الحد_الادنى_للطلب',        '0'],
      ['نص_الشحن',                 'الشحن عبر شركات النقل — يُحسب حسب الوجهة والوزن'],
      ['نص_الدفع',                 'تحويل بنكي / نقداً عند الاستلام / حسب اتفاق الحساب'],
      ['بانر_العروض',              'خصم نهاية السنة على أصناف مختارة 🎉'],
      ['تنويه_الاسعار',            'الأسعار لا تشمل الضريبة ما لم يُذكر خلاف ذلك']
    ]);
  }
  cf.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1f3a93').setFontColor('#ffffff');
  cf.setFrozenRows(1); cf.setRightToLeft(true);

  // --- تعليمات ---
  var hp = ss.getSheetByName(SH_HELP) || ss.insertSheet(SH_HELP);
  hp.clear(); hp.setRightToLeft(true);
  hp.getRange(1,1,13,1).setValues([
    ['طريقة العمل باختصار'],
    ['1) من منيو «الكتالوج» شغّل: إعداد أولي ثم استيراد الأصناف ثم مزامنة الصور.'],
    ['2) الاستيراد يجيب من قاعدة البيانات: الكود، الاسم، وحدة البيع، سعر الجملة (سعر1)،'],
    ['   عدد الحبات بالكرتون، والأرصدة. سعر3 قطاعي لفرع02 ولا يدخل الموقع إطلاقاً.'],
    ['3) المزامنة تربط كل صنف بصورته من مجلد «جميع الصور» وتحدد قسمه من مجلدات التصنيف.'],
    ['4) اللي تعبّيه أنت يدوياً: علامة «عرض»، وترتيب العرض، وأي وسوم بحث، وسعر كرتون خاص.'],
    ['5) سعر الكرتون = سعر الحبة × عدد الحبات — يُحسب في الموقع دائماً، فلا يتقادم.'],
    ['6) عمود «إظهار» = لا  يخفي الصنف من الموقع بدون حذفه.'],
    [''],
    ['مهم: لازم يكون ملف الشيت مشارك «أي شخص لديه الرابط — مُطّلع» عشان الموقع يقرأ منه.'],
    ['ومجلد الصور كذلك — استخدم خيار «فتح صلاحية عرض الصور للعملاء» من المنيو.'],
    [''],
    ['أعِد تشغيل الاستيراد والمزامنة كل ما وصلت بضاعة جديدة أو تغيرت الأسعار.']
  ]);
  hp.setColumnWidth(1, 700);

  say_('تم إعداد الصفحات ✅');
}

/* ================== 2) استيراد الأصناف ================== */

function importItems() {
  var ss = ss_();
  var sh = ss.getSheetByName(SH_ITEMS);
  if (!sh) { setupSheets(); sh = ss.getSheetByName(SH_ITEMS); }

  var rows = readCatalogCsv_();
  if (!rows.length) throw new Error('ملف ' + CATALOG_CSV_NAME + ' فاضي أو غير موجود في درايف');

  var last = sh.getLastRow();
  var existing = last > 1 ? sh.getRange(2, 1, last - 1, ITEM_HEADERS.length).getValues() : [];
  var index = {};
  for (var i = 0; i < existing.length; i++) {
    var k = String(existing[i][0]).trim();
    if (k) index[k] = i;
  }

  var added = 0, updated = 0;
  for (var r = 0; r < rows.length; r++) {
    var src = rows[r];
    var code = String(src[0]).trim();
    if (!code) continue;

    if (index.hasOwnProperty(code)) {
      var row = existing[index[code]];
      for (var c = 0; c < DB_COLS.length; c++) row[DB_COLS[c]] = src[DB_COLS[c]];
      updated++;
    } else {
      var nr = new Array(ITEM_HEADERS.length).fill('');
      for (var j = 0; j < ITEM_HEADERS.length && j < src.length; j++) nr[j] = src[j];
      nr[15] = 'نعم';
      existing.push(nr);
      index[code] = existing.length - 1;
      added++;
    }
  }

  if (existing.length) sh.getRange(2, 1, existing.length, ITEM_HEADERS.length).setValues(existing);
  say_('تم الاستيراد ✅\nجديد: ' + added + '   محدَّث: ' + updated);
}

/** يقرأ catalog_items.csv من درايف ويعيد صفوفاً مرتبة حسب ITEM_HEADERS */
function readCatalogCsv_() {
  var file = null;
  if (CATALOG_CSV_ID) {
    try { file = DriveApp.getFileById(CATALOG_CSV_ID); } catch (e) { file = null; }
  }
  if (!file) {
    var it = DriveApp.getFilesByName(CATALOG_CSV_NAME);
    var newest = null;
    while (it.hasNext()) {
      var f = it.next();
      if (!newest || f.getLastUpdated() > newest.getLastUpdated()) newest = f;
    }
    file = newest;
  }
  if (!file) throw new Error('ما لقيت ملف ' + CATALOG_CSV_NAME + ' في درايف');

  var txt = file.getBlob().getDataAsString('UTF-8').replace(/^﻿/, '');
  var all = Utilities.parseCsv(txt);
  if (all.length < 2) return [];

  // اربط أعمدة الملف بأعمدة الشيت بالاسم (حتى لو تغيّر ترتيبها)
  var head = all[0].map(function (h) { return String(h).replace(/^﻿/, '').trim(); });
  var pos = {};
  for (var i = 0; i < head.length; i++) pos[head[i]] = i;

  var out = [];
  for (var r = 1; r < all.length; r++) {
    var src = all[r];
    if (!src.length || !String(src[pos['كود الصنف']] || '').trim()) continue;
    var row = [];
    for (var c = 0; c < ITEM_HEADERS.length; c++) {
      var p = pos[ITEM_HEADERS[c]];
      var v = (p === undefined) ? '' : src[p];
      if (c === 4 || c === 5 || c === 7 || c === 8 || c === 9 || c === 10) {
        v = (v === '' || v === undefined) ? '' : Number(v);
        if (isNaN(v)) v = '';
      }
      row.push(v === undefined ? '' : v);
    }
    out.push(row);
  }
  return out;
}

/* ================== 3) مزامنة الصور والأقسام ================== */

function syncImages() {
  var ss = ss_();
  var sh = ss.getSheetByName(SH_ITEMS);
  if (!sh || sh.getLastRow() < 2) throw new Error('لا توجد أصناف — شغّل الاستيراد أولاً');

  var imgById   = {};   // كود -> معرف الصورة
  var catByCode = {};   // كود -> اسم القسم
  var isNew     = {};   // كود -> جديد
  var cats      = {};

  // أ) مجلد جميع الصور
  scanFolder_(ALL_IMAGES_FOLDER_ID, function (code, id) {
    if (!imgById[code]) imgById[code] = id;
  });

  // ب) مجلدات التصنيف: اسم المجلد = القسم
  var root = DriveApp.getFolderById(CATEGORY_ROOT_FOLDER_ID);
  var subs = root.getFolders();
  while (subs.hasNext()) {
    var f = subs.next();
    var catName = f.getName().trim();
    cats[catName] = true;
    scanFolderObj_(f, function (code, id) {
      if (!catByCode[code]) catByCode[code] = catName;
      if (!imgById[code]) imgById[code] = id;
    });
  }

  // ج) مجلد الجديد
  try {
    scanFolder_(NEW_FOLDER_ID, function (code, id) {
      isNew[code] = true;
      if (!imgById[code]) imgById[code] = id;
    });
  } catch (e) {}

  // د) الكتابة في الشيت
  var n = sh.getLastRow() - 1;
  var data = sh.getRange(2,1,n,ITEM_HEADERS.length).getValues();
  var hitImg = 0, hitCat = 0;

  for (var i = 0; i < n; i++) {
    var code = String(data[i][0]).trim();
    if (!code) continue;
    var base = code.replace(/-1$/, '');           // 800-02655-1  ->  800-02655

    var img = imgById[code] || imgById[base];
    if (img) { data[i][11] = img; hitImg++; }

    var cat = catByCode[code] || catByCode[base];
    if (cat && !data[i][2]) { data[i][2] = cat; hitCat++; }

    if (isNew[code] || isNew[base]) data[i][13] = 'نعم';
  }
  sh.getRange(2,1,n,ITEM_HEADERS.length).setValues(data);

  // هـ) تحديث صفحة الأقسام
  var ct = ss.getSheetByName(SH_CATS);
  var have = {};
  if (ct.getLastRow() > 1) {
    var cv = ct.getRange(2,1,ct.getLastRow()-1,1).getValues();
    for (var j = 0; j < cv.length; j++) have[String(cv[j][0]).trim()] = true;
  }
  var toAdd = [];
  Object.keys(cats).forEach(function (c) {
    if (!have[c]) toAdd.push([c, '', 'نعم', '']);
  });
  if (toAdd.length) ct.getRange(ct.getLastRow()+1,1,toAdd.length,4).setValues(toAdd);

  say_('تمت المزامنة ✅\nصور مرتبطة: ' + hitImg + '\nأقسام محددة: ' + hitCat + '\nأقسام جديدة: ' + toAdd.length);
}

function scanFolder_(id, cb) { scanFolderObj_(DriveApp.getFolderById(id), cb); }

function scanFolderObj_(folder, cb) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    var mt = f.getMimeType();
    if (mt.indexOf('image/') !== 0) continue;
    var code = codeFromName_(f.getName());
    if (code) cb(code, f.getId());
  }
}

/** 800-02122.jpg   |   500-00302 - سيارة كهربائي.jpg   ->  الكود */
function codeFromName_(name) {
  var base = name.replace(/\.[a-zA-Z0-9]+$/, '');
  base = base.split(' - ')[0];
  base = base.split('_')[0];
  return base.trim();
}

/* ================== 4) فتح صلاحية العرض ================== */

function makeImagesPublic() {
  [ALL_IMAGES_FOLDER_ID, CATEGORY_ROOT_FOLDER_ID, NEW_FOLDER_ID].forEach(function (id) {
    try {
      DriveApp.getFolderById(id)
        .setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}
  });
  try {
    DriveApp.getFileById(ss_().getId())
      .setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}
  say_('تم فتح صلاحية العرض للصور وملف البيانات ✅');
}
