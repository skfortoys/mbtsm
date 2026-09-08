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
  'ملاحظة',               // R
  'الفئة العمرية',        // S  يدوي: مثال 3–5 سنوات
  'الماركة',              // T  يدوي
  'النوع'                 // U  يدوي: ولد / بنت / للجميع
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
    .addItem('تقرير الأصناف بلا صورة أو قسم', 'reportSyncIssues')
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
  // نحدّث صف العناوين دائماً حتى تضاف حقول الفلاتر للملفات القائمة.
  it.getRange(1,1,1,ITEM_HEADERS.length).setValues([ITEM_HEADERS]);
  // كود الصنف معرّف نصي، لا تاريخ ولا رقم. يمنع 2264-1 → 2264/01/01.
  it.getRange(1,1,Math.max(it.getMaxRows(), 1),1).setNumberFormat('@');
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
  applyCategoryValidation_(it, ct);

  // --- الاعدادات ---
  var cf = ss.getSheetByName(SH_CONF) || ss.insertSheet(SH_CONF);
  if (cf.getLastRow() === 0) {
    cf.getRange(1,1,1,2).setValues([['المفتاح','القيمة']]);
    cf.getRange(2,1,14,2).setValues([
      ['اسم_المتجر',                'الطفل المبتسم لألعاب الأطفال'],
      ['وصف_المتجر',               'جملة ألعاب الأطفال — جدة والدمام'],
      ['رقم_واتساب',               '9665XXXXXXXX'],
      ['رقم_واتساب_فرع02',         ''],
      ['رقم_واتساب_فرع05',         ''],
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
  ensureSettings_(cf);

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

/** يضيف إعدادات الواجهة الجديدة للملف القائم من دون لمس قيمه الحالية. */
function ensureSettings_(sheet) {
  var defaults = [
    ['عن_المتجر', 'نسعد بخدمتكم وتلبية احتياجاتكم من ألعاب الأطفال بالجملة.'],
    ['السجل_التجاري', ''],
    ['عنوان_المتجر', ''],
    ['رابط_الخريطة', '']
  ];
  var have = {};
  if (sheet.getLastRow() > 1) {
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
    for (var i = 0; i < rows.length; i++) have[String(rows[i][0]).trim()] = true;
  }
  var missing = defaults.filter(function (row) { return !have[row[0]]; });
  if (missing.length) sheet.getRange(sheet.getLastRow()+1,1,missing.length,2).setValues(missing);
}

/** قائمة منسدلة في عمود «القسم» مصدرها ورقة الأقسام نفسها. */
function applyCategoryValidation_(itemsSheet, categoriesSheet) {
  var count = Math.max(categoriesSheet.getLastRow() - 1, 1);
  var source = categoriesSheet.getRange(2, 1, count, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(source, true)
    .setAllowInvalid(true) // يبقي الأقسام المتعددة المفصولة بـ | صالحة.
    .setHelpText('اختر قسماً من السهم. للأقسام المتعددة اكتب أسماءها مفصولة بالرمز |.')
    .build();
  itemsSheet.getRange(2, 3, Math.max(itemsSheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
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
  // يعيد تحويل الأكواد التي فسرها الشيت كتاريخ، ويدمج النسخ المكررة منها.
  var normalized = normalizeExistingItems_(existing);
  existing = normalized.rows;
  var index = normalized.index;

  var added = 0, updated = 0;
  for (var r = 0; r < rows.length; r++) {
    var src = rows[r];
    var code = restoreItemCode_(src[0]);
    if (!code) continue;
    var key = codeKey_(code);

    if (index.hasOwnProperty(key)) {
      var row = existing[index[key]];
      row[0] = code;
      for (var c = 0; c < DB_COLS.length; c++) row[DB_COLS[c]] = src[DB_COLS[c]];
      updated++;
    } else {
      var nr = new Array(ITEM_HEADERS.length).fill('');
      for (var j = 0; j < ITEM_HEADERS.length && j < src.length; j++) nr[j] = src[j];
      nr[0] = code;
      nr[15] = 'نعم';
      existing.push(nr);
      index[key] = existing.length - 1;
      added++;
    }
  }

  // طبّق النص قبل الكتابة؛ وإلا قد يعيد Sheets تحويل الكود إلى تاريخ.
  if (existing.length) {
    sh.getRange(2,1,existing.length,1).setNumberFormat('@');
    sh.getRange(2, 1, existing.length, ITEM_HEADERS.length).setValues(existing);
  }
  // امسح الصفوف الزائدة التي نشأت من التكرار القديم، من دون حذف الصفوف خارج النطاق.
  if (last - 1 > existing.length) sh.getRange(existing.length + 2, 1, last - 1 - existing.length, ITEM_HEADERS.length).clearContent();
  say_('تم الاستيراد ✅\nجديد: ' + added + '   محدَّث: ' + updated + '\nدُمجت تكرارات الأكواد: ' + normalized.merged);
}

/** يعيد الكود الأصلي عندما حوّله Google Sheets إلى تاريخ مثل 2264/01/01. */
function restoreItemCode_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return String(value.getFullYear()) + '-' + String(value.getMonth() + 1);
  }
  var s = String(value || '').trim();
  var m = s.match(/^(\d{4})[\/-](\d{1,2})[\/-]\d{1,2}$/);
  return m ? m[1] + '-' + String(Number(m[2])) : s;
}

/** يطبع الأكواد، ويدمج الصفوف المكررة مع الحفاظ على أي حقل يدوي غير فارغ. */
function normalizeExistingItems_(rows) {
  var out = [], index = {}, merged = 0;
  var manualCols = [2,6,11,12,13,14,15,16,17,18,19,20];
  rows.forEach(function(row) {
    var code = restoreItemCode_(row[0]);
    row[0] = code;
    var key = codeKey_(code);
    if (!key) return;
    if (!index.hasOwnProperty(key)) { index[key] = out.length; out.push(row); return; }
    var target = out[index[key]];
    manualCols.forEach(function(c) { if (!target[c] && row[c]) target[c] = row[c]; });
    merged++;
  });
  return {rows: out, index: index, merged: merged};
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
  var catByCode = {};   // كود -> قائمة الأقسام (الصنف قد يكون في أكثر من مجلد)
  var isNew     = {};   // كود -> جديد
  var cats      = {};

  // أ) مجلدات التصنيف: اسم المجلد = القسم
  var root = DriveApp.getFolderById(CATEGORY_ROOT_FOLDER_ID);
  var subs = root.getFolders();
  while (subs.hasNext()) {
    var f = subs.next();
    var catName = f.getName().trim();
    cats[catName] = true;
    scanFolderObj_(f, function (code, id) {
      addCategory_(catByCode, code, catName);
      if (!imgById[code]) imgById[code] = id;
    });
  }

  // ب) «جميع الصور»: يربط الصورة دائماً، ويستنتج القسم إذا كانت الصورة
  // في مجلد فرعي اسمه أحد أقسام التصنيف (يعالج تنظيم Drive المتداخل).
  scanFolder_(ALL_IMAGES_FOLDER_ID, function (code, id, folderName) {
    if (!imgById[code]) imgById[code] = id;
    var detectedCat = String(folderName || '').trim();
    if (cats[detectedCat]) addCategory_(catByCode, code, detectedCat);
  });

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
  var hitImg = 0, hitCat = 0, noImg = 0, noCat = 0;

  for (var i = 0; i < n; i++) {
    var code = codeKey_(data[i][0]);
    if (!code) continue;
    var base = code.replace(/-1$/, '');           // 800-02655-1  ->  800-02655

    var img = imgById[code] || imgById[base] || '';
    // لا تبقِ معرف صورة قديمة إن أزيلت من درايف.
    data[i][11] = img;
    if (img) hitImg++; else noImg++;

    var cat = (catByCode[code] || catByCode[base] || []).slice();
    applyCategoryRules_(cat, data[i][1]);
    // مجلدات Drive هي مرجع التصنيف؛ الفاصل | يسمح بظهور الصنف في أكثر من قسم.
    data[i][2] = cat.join(' | ');
    if (cat.length) hitCat++; else noCat++;

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
  applyCategoryValidation_(sh, ct);

  say_('تمت المزامنة ✅\nصور مرتبطة: ' + hitImg + '\nأقسام محددة: ' + hitCat + '\nبلا صورة: ' + noImg + '\nبلا قسم: ' + noCat + '\nأقسام جديدة: ' + toAdd.length);
}

function addCategory_(map, code, category) {
  if (!code || !category) return;
  if (!map[code]) map[code] = [];
  if (map[code].indexOf(category) === -1) map[code].push(category);
}

/**
 * قواعد عرض إضافية: يبقى الصنف في قسم السعر، ويظهر أيضاً في القسم
 * المتخصص حين يدل اسمه عليه. لا تغير هذه القاعدة تصنيف Drive الأصلي.
 */
function applyCategoryRules_(categories, itemName) {
  var isPriceSection = categories.indexOf('ابو 5 ريال') !== -1 || categories.indexOf('ابو 10 ريال') !== -1;
  if (!isPriceSection) return;
  var name = String(itemName || '');
  if (name.indexOf('سلايم') !== -1) addCategoryToList_(categories, 'سلايم وصلصال والعاب تلوين');
  // «كرتون» ليس كرت تعليق؛ لذلك نستبعده من قاعدة الكروت.
  if (/كروت|كرت(?!ون)/.test(name)) addCategoryToList_(categories, 'كروت واكياس تعليق');
}

function addCategoryToList_(categories, category) {
  if (categories.indexOf(category) === -1) categories.push(category);
}

function scanFolder_(id, cb) { scanFolderObj_(DriveApp.getFolderById(id), cb); }

function scanFolderObj_(folder, cb) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    var mt = f.getMimeType();
    if (mt.indexOf('image/') !== 0) continue;
    var code = codeFromName_(f.getName());
    if (code) cb(codeKey_(code), f.getId(), folder.getName());
  }
  // بعض الأقسام تحتوي مجلدات فرعية؛ نمسحها أيضاً بدلاً من فقد صورها.
  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) scanFolderObj_(subfolders.next(), cb);
}

/** 800-02122.jpg   |   500-00302 - سيارة كهربائي.jpg   ->  الكود */
function codeFromName_(name) {
  var base = String(name).replace(/\.[a-zA-Z0-9]+$/, '').trim();
  // يقبل الكود في أول الاسم أو وسطه، مع مسافات حول الشرطات أو نص عربي ملاصق.
  base = base.replace(/\s*-\s*/g, '-');
  var m = base.match(/(?:^|[^A-Za-z0-9])([A-Za-z0-9]+(?:-[A-Za-z0-9]+)+|\d{3,})(?=$|[^A-Za-z0-9])/);
  return m ? m[1] : '';
}

/** ينشئ ورقة مراجعة للأصناف التي لا تستطيع المزامنة ربطها. */
function reportSyncIssues() {
  var ss = ss_(), sh = ss.getSheetByName(SH_ITEMS);
  if (!sh || sh.getLastRow() < 2) throw new Error('لا توجد أصناف للمراجعة');
  var report = ss.getSheetByName('تقرير المزامنة') || ss.insertSheet('تقرير المزامنة');
  report.clear(); report.setRightToLeft(true);
  var data = sh.getRange(2,1,sh.getLastRow()-1,ITEM_HEADERS.length).getValues();
  var out = [['كود الصنف','اسم الصنف','الصورة','القسم','سبب المراجعة']];
  data.forEach(function(row) {
    var noImage = !row[11], noCategory = !row[2];
    if (noImage || noCategory) out.push([row[0],row[1],row[11] ? 'مربوطة' : '',row[2] || '', noImage && noCategory ? 'بلا صورة وبلا قسم' : noImage ? 'بلا صورة' : 'بلا قسم']);
  });
  report.getRange(1,1,out.length,out[0].length).setValues(out);
  report.getRange(1,1,1,out[0].length).setFontWeight('bold').setBackground('#1f3a93').setFontColor('#ffffff');
  report.setFrozenRows(1); report.autoResizeColumns(1,out[0].length);
  say_('تم إنشاء تقرير المراجعة: ' + (out.length - 1) + ' صنفاً');
}

/** مفتاح داخلي للمطابقة فقط؛ لا يغيّر كود الصنف المعروض. */
function codeKey_(value) { return restoreItemCode_(value).toUpperCase(); }

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
