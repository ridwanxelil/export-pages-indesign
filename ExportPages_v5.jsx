// ============================================================
// ExportPages_v5.jsx
// Profesyonel Sayfa Disa Aktarma & Isimlendirme Eklentisi
//
// YENI OZELLIKLER (v5):
//   - Sablon sistemi: {doc} {para} {date} {page} {total}
//   - PDF preset secimi (belgede tanimli presetler)
//   - Ilerleme cubugu (progress bar)
//   - Ayarlari kaydet / yukle (JSON dosyasi)
//   - Export sonrasi klasoru otomatik ac
//   - Bos sayfalari atla
//   - Master page'e gore filtrele
//   - Layer gorunurluk kontrolu
//   - Bleed / Slug dahil etme
//
// KURULUM:
//   Windows : ...\Adobe InDesign\Scripts\Scripts Panel\
//   Mac     : ~/Library/Scripts/Scripts Panel/
// ============================================================

#target indesign

// ============================================================
// JSON POLYFİLL — InDesign eski JS motorları için (satır 1'de olmalı)
// ============================================================
if (typeof JSON === "undefined") {
    JSON = {
        stringify: function(obj) {
            var t = typeof obj;
            if (t === "string")  return '"' + obj.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"';
            if (t === "number" || t === "boolean") return String(obj);
            if (obj === null || obj === undefined) return "null";
            if (obj instanceof Array) {
                var arr = [];
                for (var i = 0; i < obj.length; i++) arr.push(JSON.stringify(obj[i]));
                return "[" + arr.join(",") + "]";
            }
            if (t === "object") {
                var pairs = [];
                for (var k in obj) {
                    if (obj.hasOwnProperty(k))
                        pairs.push('"' + k + '":' + JSON.stringify(obj[k]));
                }
                return "{" + pairs.join(",") + "}";
            }
            return "null";
        },
        parse: function(str) { return eval("(" + str + ")"); }
    };
}

// ============================================================
// SABITLER
// ============================================================
var VERSION      = "5.0";
var SETTINGS_KEY = "ExportPages_v5_settings"; // app.scriptPreferences kullanimi icin
var DEFAULTS = {
    maxCharacters  : 100,
    maxAffix       : 40,
    fallbackPrefix : "Sayfa_",
    template       : "{para}_s{page}",
    prefix         : "",
    suffix         : "",
};

// ============================================================
// AYAR KAYDET / YUKLE — DEFAULTS'dan hemen sonra tanımlanmalı
// ============================================================
var SETTINGS_FILE_PATH = (function() {
    var candidates = [];

    // 1. İndirilenler
    try {
        var userHome = Folder("~").fsName;
        var dl = new Folder(userHome + "/Downloads");
        if (!dl.exists) dl = new Folder(userHome + "/\u0130ndirilenler");
        if (dl.exists) candidates.push(dl.fsName);
    } catch(e) {}

    // 2. Masaüstü
    try {
        var userHome2 = Folder("~").fsName;
        var desk = new Folder(userHome2 + "/Desktop");
        if (!desk.exists) desk = new Folder(userHome2 + "/Masa\u00fcst\u00fc");
        if (desk.exists) candidates.push(desk.fsName);
    } catch(e) {}

    // 3. Belgelerim
    try {
        var docs = Folder.myDocuments;
        if (docs && docs.exists) candidates.push(docs.fsName);
    } catch(e) {}

    // 4. Kullanıcı ana klasörü
    try {
        var home = Folder("~");
        if (home.exists) candidates.push(home.fsName);
    } catch(e) {}

    // 5. Temp
    try { candidates.push(Folder.temp.fsName); } catch(e) { candidates.push("/tmp"); }

    // Gerçekten yazabildiğini test et
    for (var i = 0; i < candidates.length; i++) {
        var tp = candidates[i] + "/_eptest_.tmp";
        var tf = new File(tp);
        try {
            tf.encoding = "UTF-8";
            tf.open("w");
            tf.write("1");
            tf.close();
            tf.remove();
            return candidates[i] + "/ExportPages_v5_settings.json";
        } catch(e) {
            try { tf.close(); } catch(e2) {}
            try { tf.remove(); } catch(e3) {}
        }
    }
    return Folder.temp.fsName + "/ExportPages_v5_settings.json";
})();

function getSettingsFile() {
    return new File(SETTINGS_FILE_PATH);
}

function saveSettings(s) {
    if (!s.autoSave) return;

    var data = {
        paraStyleName : s.paraStyleName,
        charStyleName : s.charStyleName,
        template      : s.template,
        prefix        : s.prefix,
        suffix        : s.suffix,
        asSpread      : s.asSpread,
        useRange      : s.useRange,
        customRange   : s.customRange,
        skipEmpty     : s.skipEmpty,
        format        : s.format,
        qualityIdx    : s.qualityIdx,
        resIdx        : s.resIdx,
        pdfPreset     : s.pdfPreset,
        includeBleed  : s.includeBleed,
        includeSlug   : s.includeSlug,
        masterFilter  : s.masterFilter,
        layerStates   : s.layerStates,
        lastFolder    : s.lastFolder,
        openFolder    : s.openFolder,
        autoSave      : s.autoSave,
    };

    var jsonStr = JSON.stringify(data);
    var f = new File(SETTINGS_FILE_PATH);
    try {
        f.encoding = "UTF-8";
        f.open("w");
        f.write(jsonStr);
        f.close();
    } catch(e) {
        try { f.close(); } catch(e2) {}
        alert("⚠ Ayarlar kaydedilemedi!\n\nDosya: " + SETTINGS_FILE_PATH +
              "\nHata: " + e.message);
    }
}

function loadSettings() {
    try {
        var f = new File(SETTINGS_FILE_PATH);
        if (!f.exists) return {};
        f.encoding = "UTF-8";
        f.open("r");
        var raw = f.read();
        f.close();
        if (raw && raw.length > 2) {
            return JSON.parse(raw) || {};
        }
    } catch(e) {
        try { f.close(); } catch(e2) {}
    }
    return {};
}


// ============================================================
// ANA AKIS
// ============================================================
(function () {

    if (app.documents.length === 0) {
        alert("Lutfen once bir belge acin.");
        exit();
    }

    var doc = app.activeDocument;

    // ── Stil listelerini topla ───────────────────────────────
    var paraStyleNames = [];
    for (var ps = 0; ps < doc.paragraphStyles.length; ps++) {
        var pn = doc.paragraphStyles[ps].name;
        if (pn !== "[No Paragraph Style]" && pn !== "[Basic Paragraph]")
            paraStyleNames.push(pn);
    }
    if (paraStyleNames.length === 0) {
        alert("Belgede kullanilabilir paragraf stili bulunamadi.");
        exit();
    }

    var charStyleNames = ["(Kullanma - tum paragraf)"];
    for (var cs = 0; cs < doc.characterStyles.length; cs++) {
        var cn = doc.characterStyles[cs].name;
        if (cn !== "[No character style]") charStyleNames.push(cn);
    }

    // ── PDF Presetlerini topla ───────────────────────────────
    var pdfPresetNames = [];
    try {
        for (var pp = 0; pp < app.pdfExportPresets.length; pp++) {
            pdfPresetNames.push(app.pdfExportPresets[pp].name);
        }
    } catch(e) {}
    if (pdfPresetNames.length === 0) pdfPresetNames = ["[Varsayilan]"];

    // ── Master page listesini topla ──────────────────────────
    var masterNames = ["(Filtre Yok - Hepsi)"];
    try {
        for (var mp = 0; mp < doc.masterSpreads.length; mp++) {
            masterNames.push(doc.masterSpreads[mp].name);
        }
    } catch(e) {}

    // ── Layer listesini topla ────────────────────────────────
    var layerNames = [];
    try {
        for (var ly = 0; ly < doc.layers.length; ly++) {
            layerNames.push(doc.layers[ly].name);
        }
    } catch(e) {}

    // ── Kayitli ayarlari yukle ───────────────────────────────
    var savedSettings = loadSettings();

    // ── Dialog ──────────────────────────────────────────────
    var s = showDialog(
        paraStyleNames, charStyleNames, pdfPresetNames,
        masterNames, layerNames, doc.pages.length, savedSettings
    );
    if (!s) { exit(); }

    // Ayarlari kaydet
    saveSettings(s);

    // ── Sayfa listesini olustur ──────────────────────────────
    var pageIndices = resolvePageRange(s.pageRange, doc.pages.length);
    if (pageIndices.length === 0) {
        alert("Gecerli sayfa araligi bulunamadi.");
        exit();
    }

    // Master page filtresi uygula
    if (s.masterFilter) {
        pageIndices = filterByMaster(doc, pageIndices, s.masterFilter);
        if (pageIndices.length === 0) {
            alert("Secilen master page'e ait sayfa bulunamadi.");
            exit();
        }
    }

    // ── Stilleri al ─────────────────────────────────────────
    var targetParaStyle = safeGetStyle(doc.paragraphStyles, s.paraStyleName);
    var targetCharStyle = s.charStyleName
                         ? safeGetStyle(doc.characterStyles, s.charStyleName) : null;

    // ── Layer gorunurlugunu ayarla ───────────────────────────
    var layerStates = applyLayerVisibility(doc, s.layerStates);

    // ── PDF preset ──────────────────────────────────────────
    if (s.format === "PDF" && s.pdfPreset && s.pdfPreset !== "[Varsayilan]") {
        try {
            var preset = app.pdfExportPresets.itemByName(s.pdfPreset);
            if (preset.isValid) app.activeDocument.importPDFPreset(preset);
        } catch(e) {}
    }

    // ── Export ayarlari ──────────────────────────────────────
    var qualityMap = [
        JPEGOptionsQuality.LOW, JPEGOptionsQuality.MEDIUM,
        JPEGOptionsQuality.HIGH, JPEGOptionsQuality.MAXIMUM
    ];
    var resMap = [72, 96, 150, 300];

    if (s.format === "JPG") {
        var jo = app.jpegExportPreferences;
        jo.jpegQuality      = qualityMap[s.qualityIdx];
        jo.exportResolution = resMap[s.resIdx];
        jo.antiAlias        = true;
        jo.simulateOverprint = false;
    } else if (s.format === "PNG") {
        var po = app.pngExportPreferences;
        po.exportResolution      = resMap[s.resIdx];
        po.antiAlias             = true;
        po.transparentBackground = false;
    }

    // ── Progress bar ─────────────────────────────────────────
    var prog = createProgressBar(pageIndices.length);

    // ── Sayfalari disa aktar ─────────────────────────────────
    var errors   = [];
    var exported = 0;
    var skipped  = 0;
    var today    = getDateString();
    var docName  = sanitize(doc.name.replace(/\.[^.]+$/, ""), 40);

    for (var i = 0; i < pageIndices.length; i++) {

        var pageIdx = pageIndices[i];
        var page    = doc.pages[pageIdx];
        var pageNum = pageIdx + 1;

        // Progress bar guncelle
        updateProgress(prog, i + 1, pageIndices.length,
                       "Sayfa " + pageNum + " isleniyor...");

        // Bos sayfa kontrolu
        if (s.skipEmpty && isPageEmpty(page)) {
            skipped++;
            continue;
        }

        // Paragraf metnini al
        var paraText = "";
        if (targetParaStyle) {
            paraText = getFirstParagraphText(page, targetParaStyle, targetCharStyle);
        }
        if (paraText === "") paraText = DEFAULTS.fallbackPrefix + pageNum;
        paraText = sanitize(paraText, DEFAULTS.maxCharacters);

        // Sablon ile dosya adini olustur
        var finalName = buildFileName(s.template, {
            doc   : docName,
            para  : paraText,
            date  : today,
            page  : zeroPad(pageNum, String(doc.pages.length).length),
            total : String(doc.pages.length),
        });

        var ext      = s.format.toLowerCase() === "pdf" ? "pdf"
                     : s.format.toLowerCase() === "eps" ? "eps"
                     : s.format.toLowerCase() === "png" ? "png"
                     : s.format.toLowerCase() === "tiff" ? "tif"
                     : "jpg";

        var filePath = uniquePath(s.outputFolder.fsName,
                                  (s.prefix || "") + finalName + (s.suffix || ""),
                                  ext);

        try {
            exportPage(doc, page, s, filePath, qualityMap, resMap);
            exported++;
        } catch (e) {
            errors.push("Sayfa " + pageNum + ": " + e.message);
        }
    }

    // Progress kapat
    prog.close();

    // Layer gorunurlugunu geri yukle
    restoreLayerVisibility(doc, layerStates);

    // Export sonrasi klasoru ac
    if (s.openFolder) {
        s.outputFolder.execute();
    }

    // ── Sonuc raporu ─────────────────────────────────────────
    var msg = "Tamamlandi!\n\n"
            + "Aktarilan : " + exported + " sayfa\n"
            + (skipped > 0 ? "Atlanan   : " + skipped + " bos sayfa\n" : "")
            + (errors.length > 0 ? "Hata      : " + errors.length + " sayfa\n" : "")
            + "\nKlasor: " + s.outputFolder.fsName;
    if (errors.length) msg += "\n\nHata Detaylari:\n" + errors.join("\n");
    alert(msg);

})();

// ============================================================
// EXPORT DISPATCHER
// ============================================================
function exportPage(doc, page, s, filePath, qualityMap, resMap) {

    var fmt        = s.format;
    var pageStr    = s.asSpread ? getSpreadPageString(page) : page.name;

    if (fmt === "JPG") {
        var jo = app.jpegExportPreferences;
        jo.jpegExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
        jo.pageString      = pageStr;
        doc.exportFile(ExportFormat.JPG, new File(filePath), false);

    } else if (fmt === "PNG") {
        var po = app.pngExportPreferences;
        po.pngExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
        po.pageString     = pageStr;
        doc.exportFile(ExportFormat.PNG, new File(filePath), false);

    } else if (fmt === "PDF") {
        var pdfP = app.pdfExportPreferences;
        pdfP.pageRange           = pageStr;
        pdfP.exportReaderSpreads = s.asSpread;
        // Bleed / Slug
        pdfP.includeSlugWithPDF  = s.includeSlug  || false;
        if (s.includeBleed) {
            pdfP.useDocumentBleedWithPDF = true;
        } else {
            pdfP.useDocumentBleedWithPDF = false;
            pdfP.bleedBottom = pdfP.bleedTop = pdfP.bleedInside = pdfP.bleedOutside = 0;
        }
        doc.exportFile(ExportFormat.PDF_TYPE, new File(filePath), false);

    } else if (fmt === "EPS") {
        var ep = app.epsExportPreferences;
        ep.pageRange = pageStr;
        doc.exportFile(ExportFormat.EPS_TYPE, new File(filePath), false);

    } else if (fmt === "TIFF") {
        var jo2 = app.jpegExportPreferences;
        jo2.jpegExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
        jo2.pageString      = pageStr;
        try {
            doc.exportFile(ExportFormat.TIFF, new File(filePath), false);
        } catch(e) {
            throw new Error("TIFF bu surumde desteklenmiyor: " + e.message);
        }
    }
}

function getSpreadPageString(page) {
    try {
        var spread = page.parent;
        var names  = [];
        for (var i = 0; i < spread.pages.length; i++)
            names.push(spread.pages[i].name);
        return names.join(",");
    } catch(e) { return page.name; }
}

// ============================================================
// PROGRESS BAR
// ============================================================
function createProgressBar(total) {
    var win = new Window("palette", "Disa Aktariliyor...");
    win.orientation   = "column";
    win.alignChildren = ["fill", "top"];
    win.margins       = [16, 16, 16, 16];
    win.spacing       = 10;

    var lbl = win.add("statictext", undefined, "Hazirlanıyor...");
    lbl.preferredSize.width = 320;

    var bar = win.add("progressbar", undefined, 0, total);
    bar.preferredSize = [320, 16];

    var pct = win.add("statictext", undefined, "0 / " + total);
    pct.alignment = "center";

    win.lbl = lbl;
    win.bar = bar;
    win.pct = pct;
    win.show();
    return win;
}

function updateProgress(win, current, total, msg) {
    try {
        win.lbl.text = msg;
        win.bar.value = current;
        win.pct.text  = current + " / " + total;
        win.update();
    } catch(e) {}
}

// ============================================================
// DIALOG
// ============================================================
function showDialog(paraStyleNames, charStyleNames, pdfPresetNames,
                   masterNames, layerNames, totalPages, saved) {

    var dlg = new Window("dialog", "Profesyonel Export  v" + VERSION);
    dlg.orientation   = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.spacing       = 8;
    dlg.margins       = [18, 16, 18, 16];

    // Baslik
    var hdr = dlg.add("statictext", undefined,
        "Profesyonel Sayfa Disa Aktarici  v" + VERSION);
    hdr.graphics.font = ScriptUI.newFont("dialog", "BOLD", 12);

    // ── 2 Kolonlu Ana Duzen ──────────────────────────────────
    var mainGrp = dlg.add("group");
    mainGrp.orientation   = "row";
    mainGrp.alignChildren = ["fill", "top"];
    mainGrp.spacing       = 10;

    var colLeft  = mainGrp.add("group");
    colLeft.orientation   = "column";
    colLeft.alignChildren = ["fill", "top"];
    colLeft.spacing       = 8;
    colLeft.preferredSize.width = 390;

    var colRight = mainGrp.add("group");
    colRight.orientation   = "column";
    colRight.alignChildren = ["fill", "top"];
    colRight.spacing       = 8;
    colRight.preferredSize.width = 340;

    // ===========================================================
    // SOL KOLON
    // ===========================================================

    // ── Panel: Stil Secimi ───────────────────────────────────
    var pnlStyle = colLeft.add("panel", undefined, "Stil Secimi");
    pnlStyle.orientation   = "column";
    pnlStyle.alignChildren = ["fill", "top"];
    pnlStyle.spacing = 7; pnlStyle.margins = [12, 16, 12, 12];

    var rPara = pnlStyle.add("group");
    rPara.orientation = "row"; rPara.alignChildren = ["left", "center"];
    var lPara = rPara.add("statictext", undefined, "Paragraf Stili (*)");
    lPara.preferredSize.width = 130;
    var ddPara = rPara.add("dropdownlist", undefined, paraStyleNames);
    ddPara.preferredSize.width = 230;
    ddPara.selection = findIndex(paraStyleNames, saved.paraStyleName) || 0;

    var rChar = pnlStyle.add("group");
    rChar.orientation = "row"; rChar.alignChildren = ["left", "center"];
    var lChar = rChar.add("statictext", undefined, "Karakter Stili");
    lChar.preferredSize.width = 130;
    var ddChar = rChar.add("dropdownlist", undefined, charStyleNames);
    ddChar.preferredSize.width = 230;
    ddChar.selection = findIndex(charStyleNames, saved.charStyleName) || 0;

    // ── Panel: Dosya Adi Sablonu ─────────────────────────────
    var pnlTemplate = colLeft.add("panel", undefined, "Dosya Adi Sablonu");
    pnlTemplate.orientation   = "column";
    pnlTemplate.alignChildren = ["fill", "top"];
    pnlTemplate.spacing = 7; pnlTemplate.margins = [12, 16, 12, 12];

    var rTpl = pnlTemplate.add("group");
    rTpl.orientation = "row"; rTpl.alignChildren = ["left", "center"];
    var lTpl = rTpl.add("statictext", undefined, "Sablon:");
    lTpl.preferredSize.width = 60;
    var txtTpl = rTpl.add("edittext", undefined,
        saved.template || DEFAULTS.template);
    txtTpl.preferredSize.width = 290;

    var rTplHelp = pnlTemplate.add("group");
    rTplHelp.orientation = "row"; rTplHelp.alignChildren = ["left", "center"];
    rTplHelp.spacing = 6;
    var helpTxt = rTplHelp.add("statictext", undefined,
        "{para}  {doc}  {date}  {page}  {total}");
    helpTxt.graphics.foregroundColor =
        helpTxt.graphics.newPen(helpTxt.graphics.PenType.SOLID_COLOR, [0.4,0.4,0.4], 1);

    // Hizli sablon butonlari
    var rTplBtns = pnlTemplate.add("group");
    rTplBtns.orientation = "row"; rTplBtns.spacing = 4;

    function addTplBtn(label, tpl) {
        var b = rTplBtns.add("button", undefined, label);
        b.preferredSize = [90, 22];
        b.onClick = function() {
            txtTpl.text = tpl;
            refreshPreview();
        };
    }
    addTplBtn("{para}_s{page}",     "{para}_s{page}");
    addTplBtn("{doc}_{para}",        "{doc}_{para}");
    addTplBtn("{date}_{para}",       "{date}_{para}");
    addTplBtn("{page}_{para}",       "{page}_{para}");

    // Onizleme
    var rPrev = pnlTemplate.add("group");
    rPrev.orientation = "row"; rPrev.alignChildren = ["left", "center"];
    rPrev.add("statictext", undefined, "Ornek:");
    var lblPrev = rPrev.add("statictext", undefined, "");
    lblPrev.preferredSize.width = 300;

    // ── Panel: Dosya Adi On Ek / Son Ek ─────────────────────
    var pnlAffix = colLeft.add("panel", undefined, "Dosya Adı Ön Ek / Son Ek");
    pnlAffix.orientation   = "column";
    pnlAffix.alignChildren = ["fill", "top"];
    pnlAffix.spacing = 7; pnlAffix.margins = [12, 16, 12, 12];

    var rPrefix = pnlAffix.add("group");
    rPrefix.orientation = "row"; rPrefix.alignChildren = ["left", "center"];
    var lPrefix = rPrefix.add("statictext", undefined, "Ön Ek (prefix):");
    lPrefix.preferredSize.width = 120;
    var txtPrefix = rPrefix.add("edittext", undefined, saved.prefix || DEFAULTS.prefix);
    txtPrefix.preferredSize.width = 220;
    txtPrefix.helpTip = "Dosya adının başına eklenir. Örn: 2024_";

    var rPrefixHint = pnlAffix.add("group");
    rPrefixHint.orientation = "row";
    var hintPrefix = rPrefixHint.add("statictext", undefined, "    Örnek: 2024_  →  2024_dosyaadi.jpg");
    hintPrefix.graphics.foregroundColor =
        hintPrefix.graphics.newPen(hintPrefix.graphics.PenType.SOLID_COLOR, [0.4,0.4,0.4], 1);

    var rSuffix = pnlAffix.add("group");
    rSuffix.orientation = "row"; rSuffix.alignChildren = ["left", "center"];
    var lSuffix = rSuffix.add("statictext", undefined, "Son Ek (suffix):");
    lSuffix.preferredSize.width = 120;
    var txtSuffix = rSuffix.add("edittext", undefined, saved.suffix || DEFAULTS.suffix);
    txtSuffix.preferredSize.width = 220;
    txtSuffix.helpTip = "Dosya adının sonuna eklenir (uzantıdan önce). Örn: _TR";

    var rSuffixHint = pnlAffix.add("group");
    rSuffixHint.orientation = "row";
    var hintSuffix = rSuffixHint.add("statictext", undefined, "    Örnek: _TR  →  dosyaadi_TR.jpg");
    hintSuffix.graphics.foregroundColor =
        hintSuffix.graphics.newPen(hintSuffix.graphics.PenType.SOLID_COLOR, [0.4,0.4,0.4], 1);

    txtPrefix.onChanging = refreshPreview;
    txtSuffix.onChanging = refreshPreview;

    // ── Panel: Page Export Options ───────────────────────────
    var pnlPageOpts = colLeft.add("panel", undefined, "Page Export Options");
    pnlPageOpts.orientation   = "column";
    pnlPageOpts.alignChildren = ["fill", "top"];
    pnlPageOpts.spacing = 7; pnlPageOpts.margins = [12, 16, 12, 12];

    // Spread / Page
    var rSpread = pnlPageOpts.add("group");
    rSpread.orientation = "row"; rSpread.alignChildren = ["left", "center"];
    rSpread.spacing = 20;
    var lSpr = rSpread.add("statictext", undefined, "Export Turu:");
    lSpr.preferredSize.width = 100;
    var radPage   = rSpread.add("radiobutton", undefined, "Page");
    var radSpread = rSpread.add("radiobutton", undefined, "Spread");
    radPage.value = !saved.asSpread;
    radSpread.value = saved.asSpread || false;

    // Sayfa araligi
    var rRange = pnlPageOpts.add("group");
    rRange.orientation = "row"; rRange.alignChildren = ["left", "center"];
    var lRng = rRange.add("statictext", undefined, "Sayfa Araligi:");
    lRng.preferredSize.width = 100;
    var radAll   = rRange.add("radiobutton", undefined, "Tumu");
    var radRange = rRange.add("radiobutton", undefined, "Aralik:");
    var txtRange = rRange.add("edittext", undefined,
        saved.customRange || ("1-" + totalPages));
    txtRange.preferredSize.width = 110;

    if (saved.useRange) {
        radRange.value = true; txtRange.enabled = true;
    } else {
        radAll.value   = true; txtRange.enabled = false;
    }

    radAll.onClick   = function() { txtRange.enabled = false; };
    radRange.onClick = function() { txtRange.enabled = true; txtRange.active = true; };

    // Bos sayfalari atla
    var rSkip = pnlPageOpts.add("group");
    rSkip.orientation = "row"; rSkip.alignChildren = ["left", "center"];
    rSkip.add("statictext", undefined, "").preferredSize.width = 100;
    var chkSkip = rSkip.add("checkbox", undefined, "Bos sayfalari atla");
    chkSkip.value = saved.skipEmpty || false;

    // ===========================================================
    // SAG KOLON
    // ===========================================================

    // ── Panel: Output Format ─────────────────────────────────
    var pnlFormat = colRight.add("panel", undefined, "Output Format");
    pnlFormat.orientation   = "column";
    pnlFormat.alignChildren = ["fill", "top"];
    pnlFormat.spacing = 7; pnlFormat.margins = [12, 16, 12, 12];

    var rFmt = pnlFormat.add("group");
    rFmt.orientation = "row"; rFmt.alignChildren = ["left", "center"];
    rFmt.add("statictext", undefined, "Format:").preferredSize.width = 90;
    var ddFmt = rFmt.add("dropdownlist", undefined, ["JPG", "PNG", "PDF", "EPS", "TIFF"]);
    ddFmt.selection = findIndex(["JPG","PNG","PDF","EPS","TIFF"], saved.format) || 0;
    ddFmt.preferredSize.width = 100;

    // Kalite / Cozunurluk (JPG/PNG/TIFF)
    var rQR = pnlFormat.add("group");
    rQR.orientation = "row"; rQR.spacing = 10;

    var gQ = rQR.add("group");
    gQ.orientation = "row"; gQ.alignChildren = ["left", "center"];
    gQ.add("statictext", undefined, "Kalite:").preferredSize.width = 55;
    var ddQ = gQ.add("dropdownlist", undefined, ["Dusuk","Orta","Yuksek","Maksimum"]);
    ddQ.selection = saved.qualityIdx || 2; ddQ.preferredSize.width = 95;

    var gR = rQR.add("group");
    gR.orientation = "row"; gR.alignChildren = ["left", "center"];
    gR.add("statictext", undefined, "DPI:").preferredSize.width = 30;
    var ddR = gR.add("dropdownlist", undefined, ["72","96","150","300"]);
    ddR.selection = saved.resIdx || 2; ddR.preferredSize.width = 65;

    // PDF Preset
    var rPreset = pnlFormat.add("group");
    rPreset.orientation = "row"; rPreset.alignChildren = ["left", "center"];
    rPreset.add("statictext", undefined, "PDF Preset:").preferredSize.width = 90;
    var ddPreset = rPreset.add("dropdownlist", undefined, pdfPresetNames);
    ddPreset.selection = findIndex(pdfPresetNames, saved.pdfPreset) || 0;
    ddPreset.preferredSize.width = 180;

    // Bleed / Slug
    var rBleed = pnlFormat.add("group");
    rBleed.orientation = "row"; rBleed.spacing = 16;
    rBleed.add("statictext", undefined, "").preferredSize.width = 90;
    var chkBleed = rBleed.add("checkbox", undefined, "Bleed dahil");
    var chkSlug  = rBleed.add("checkbox", undefined, "Slug dahil");
    chkBleed.value = saved.includeBleed || false;
    chkSlug.value  = saved.includeSlug  || false;

    // Format degisince goster/gizle
    function onFormatChange() {
        var fmt = ddFmt.selection ? ddFmt.selection.text : "JPG";
        var isRaster = (fmt === "JPG" || fmt === "PNG" || fmt === "TIFF");
        var isPDF    = (fmt === "PDF");
        rQR.visible     = isRaster;
        rPreset.visible = isPDF;
        rBleed.visible  = isPDF;
        ddQ.items[0].text = (fmt === "JPG" || fmt === "TIFF") ? "Dusuk" : "Dusuk";
        lblQ_ref.text     = (fmt === "JPG" || fmt === "TIFF") ? "Kalite:" : "Kalite:";
    }
    var lblQ_ref = gQ.children[0]; // "Kalite:" static text referansi
    ddFmt.onChange = function() { onFormatChange(); refreshPreview(); };
    onFormatChange();

    // ── Panel: Master Page Filtresi ──────────────────────────
    var pnlMaster = colRight.add("panel", undefined, "Master Page Filtresi");
    pnlMaster.orientation   = "column";
    pnlMaster.alignChildren = ["fill", "top"];
    pnlMaster.spacing = 7; pnlMaster.margins = [12, 16, 12, 12];

    var rMaster = pnlMaster.add("group");
    rMaster.orientation = "row"; rMaster.alignChildren = ["left", "center"];
    rMaster.add("statictext", undefined, "Master:").preferredSize.width = 70;
    var ddMaster = rMaster.add("dropdownlist", undefined, masterNames);
    ddMaster.selection = findIndex(masterNames, saved.masterFilter) || 0;
    ddMaster.preferredSize.width = 220;

    // ── Panel: Layer Gorunurlugu ─────────────────────────────
    var pnlLayers = colRight.add("panel", undefined, "Layer Gorunurlugu");
    pnlLayers.orientation   = "column";
    pnlLayers.alignChildren = ["fill", "top"];
    pnlLayers.spacing = 5; pnlLayers.margins = [12, 16, 12, 12];

    var layerChecks = [];
    if (layerNames.length === 0) {
        pnlLayers.add("statictext", undefined, "(Layer bulunamadi)");
    } else {
        for (var lc = 0; lc < layerNames.length; lc++) {
            var rLay = pnlLayers.add("group");
            rLay.orientation = "row"; rLay.alignChildren = ["left", "center"];
            var chk = rLay.add("checkbox", undefined, layerNames[lc]);
            chk.preferredSize.width = 280;
            // Varsayilan: kayitli durum varsa yukle, yoksa tumu acik
            var savedLayerStates = saved.layerStates || {};
            chk.value = (savedLayerStates[layerNames[lc]] !== undefined)
                        ? savedLayerStates[layerNames[lc]] : true;
            layerChecks.push({ name: layerNames[lc], chk: chk });
        }
        // Hepsini sec / kaldir butonlari
        var rLayBtns = pnlLayers.add("group");
        rLayBtns.orientation = "row"; rLayBtns.spacing = 6;
        var btnAllOn = rLayBtns.add("button", undefined, "Hepsini Ac");
        btnAllOn.preferredSize = [90, 22];
        btnAllOn.onClick = function() {
            for (var k = 0; k < layerChecks.length; k++)
                layerChecks[k].chk.value = true;
        };
        var btnAllOff = rLayBtns.add("button", undefined, "Hepsini Kapat");
        btnAllOff.preferredSize = [100, 22];
        btnAllOff.onClick = function() {
            for (var k = 0; k < layerChecks.length; k++)
                layerChecks[k].chk.value = false;
        };
    }

    // ── Panel: Kayit Klasoru ─────────────────────────────────
    var selectedFolder = saved.lastFolder ? new Folder(saved.lastFolder) : null;
    var pnlFolder = colRight.add("panel", undefined, "Kayit Klasoru");
    pnlFolder.orientation   = "column";
    pnlFolder.alignChildren = ["fill", "top"];
    pnlFolder.spacing = 6; pnlFolder.margins = [12, 16, 12, 12];

    var rFolder = pnlFolder.add("group");
    rFolder.orientation = "row"; rFolder.alignChildren = ["left", "center"];
    var lblFolder = rFolder.add("statictext", undefined,
        selectedFolder ? shortenPath(selectedFolder.fsName, 30) : "(Klasor secilmedi)");
    lblFolder.preferredSize.width = 210;
    var btnBrowse = rFolder.add("button", undefined, "Gozat...");
    btnBrowse.preferredSize.width = 70;
    btnBrowse.onClick = function() {
        var f = Folder.selectDialog("Kayit klasorunu secin:");
        if (f) {
            selectedFolder = f;
            lblFolder.text    = shortenPath(f.fsName, 30);
            lblFolder.helpTip = f.fsName;
        }
    };

    // Export sonrasi ac
    var rOpenFolder = pnlFolder.add("group");
    rOpenFolder.orientation = "row"; rOpenFolder.alignChildren = ["left", "center"];
    var chkOpen = rOpenFolder.add("checkbox", undefined,
        "Export sonrasi klasoru ac");
    chkOpen.value = (saved.openFolder !== undefined) ? saved.openFolder : true;

    // ===========================================================
    // ONIZLEME GUNCELLEME
    // ===========================================================
    function refreshPreview() {
        var tpl  = txtTpl.text || DEFAULTS.template;
        var para = ddPara.selection ? ddPara.selection.text : "ParagrafMetni";
        if (para.length > 20) para = para.substring(0, 20) + "...";
        var ext  = ddFmt.selection ? ddFmt.selection.text.toLowerCase() : "jpg";
        if (ext === "tiff") ext = "tif";
        var pfx = (txtPrefix && txtPrefix.text) ? sanitize(txtPrefix.text, DEFAULTS.maxAffix) : "";
        var sfx = (txtSuffix && txtSuffix.text) ? sanitize(txtSuffix.text, DEFAULTS.maxAffix) : "";
        var preview = buildFileName(tpl, {
            doc   : "BelgeAdi",
            para  : para,
            date  : getDateString(),
            page  : "001",
            total : String(totalPages),
        });
        lblPrev.text = pfx + preview + sfx + "." + ext;
    }

    txtTpl.onChanging = refreshPreview;
    ddPara.onChange   = function() { refreshPreview(); };
    refreshPreview();

    // ===========================================================
    // ALT ALAN — ayar kaydet onay + bilgi satırı
    // ===========================================================
    var gSaveRow = dlg.add("group");
    gSaveRow.orientation   = "row";
    gSaveRow.alignChildren = ["left", "center"];
    gSaveRow.alignment     = "fill";
    gSaveRow.spacing       = 10;
    gSaveRow.margins       = [4, 2, 4, 2];

    // Onay kutusu — varsayılan AÇIK (her zaman kaydeder)
    var chkSave = gSaveRow.add("checkbox", undefined, "Ayarları otomatik kaydet");
    chkSave.value = (saved.autoSave !== false); // varsayılan: true

    // Kayıt dosyası yol bilgisi
    var lblSettingsPath = gSaveRow.add("statictext", undefined, "");
    lblSettingsPath.preferredSize.width = 460;
    try {
        lblSettingsPath.text = "\u2192 " + SETTINGS_FILE_PATH;
    } catch(e) { lblSettingsPath.text = ""; }
    lblSettingsPath.graphics.foregroundColor =
        lblSettingsPath.graphics.newPen(
            lblSettingsPath.graphics.PenType.SOLID_COLOR, [0.45,0.45,0.45], 1);

    // ===========================================================
    // AYAR BUTONLARI (alt)
    // ===========================================================
    var gBtn = dlg.add("group");
    gBtn.orientation = "row";
    gBtn.alignment   = "right";
    gBtn.spacing     = 8;

    // Ayarları Sıfırla butonu
    var btnReset = gBtn.add("button", undefined, "Ayarları Sıfırla");
    btnReset.preferredSize.width = 120;
    btnReset.onClick = function() {
        if (confirm("Tüm kayıtlı ayarlar silinecek ve varsayılanlara dönülecek.\nEmin misiniz?")) {
            try {
                var rf = new File(SETTINGS_FILE_PATH);
                if (rf.exists) rf.remove();
                alert("Ayarlar sıfırlandı.\nScripti kapatıp tekrar açtığınızda varsayılan değerlerle başlayacak.");
            } catch(e) {
                alert("Sıfırlama başarısız: " + e.message);
            }
        }
    };

    var btnCancel = gBtn.add("button", undefined, "Iptal", {name:"cancel"});
    btnCancel.preferredSize.width = 80;
    btnCancel.onClick = function() { dlg.close(0); };

    var btnOK = gBtn.add("button", undefined, "Aktar  >>", {name:"ok"});
    btnOK.preferredSize.width = 110;
    btnOK.onClick = function() {
        if (!ddPara.selection) {
            alert("Lutfen bir paragraf stili secin."); return;
        }
        if (radRange.value && txtRange.text.replace(/\s/g,"") === "") {
            alert("Lutfen sayfa araligi girin (ornek: 1-5)."); return;
        }
        if (!selectedFolder) {
            alert("Lutfen bir kayit klasoru secin."); return;
        }
        dlg.close(1);
    };

    if (dlg.show() !== 1) { return null; }

    // Layer durumlarini topla
    var layerStatesOut = {};
    for (var lx = 0; lx < layerChecks.length; lx++) {
        layerStatesOut[layerChecks[lx].name] = layerChecks[lx].chk.value;
    }

    return {
        paraStyleName : ddPara.selection.text,
        charStyleName : (ddChar.selection && ddChar.selection.index > 0)
                        ? ddChar.selection.text : null,
        template      : txtTpl.text || DEFAULTS.template,
        prefix        : sanitize(txtPrefix.text || "", DEFAULTS.maxAffix),
        suffix        : sanitize(txtSuffix.text || "", DEFAULTS.maxAffix),
        asSpread      : radSpread.value,
        pageRange     : radAll.value ? ("1-" + totalPages) : txtRange.text,
        useRange      : radRange.value,
        customRange   : txtRange.text,
        skipEmpty     : chkSkip.value,
        format        : ddFmt.selection.text,
        qualityIdx    : ddQ.selection.index,
        resIdx        : ddR.selection.index,
        pdfPreset     : ddPreset.selection ? ddPreset.selection.text : null,
        includeBleed  : chkBleed.value,
        includeSlug   : chkSlug.value,
        masterFilter  : (ddMaster.selection && ddMaster.selection.index > 0)
                        ? ddMaster.selection.text : null,
        layerStates   : layerStatesOut,
        outputFolder  : selectedFolder,
        lastFolder    : selectedFolder.fsName,
        openFolder    : chkOpen.value,
        autoSave      : chkSave.value,
    };
}

// ============================================================
// AYAR KAYDET / YUKLE  (kalıcı JSON dosyası)
// ============================================================

// Ayar dosyasının yolunu döndürür — scriptle aynı klasör
// ============================================================
// AYAR KAYDET / YUKLE  (kalıcı JSON dosyası — Belgelerim klasörü)
// ============================================================

// Ayar dosyası: her zaman yazma izni olan Belgelerim'e yazar
// Ayar dosyası yolu — script yüklenince bir kez hesaplanır, hep aynı kalır

// ============================================================
// SABLON MOTORU
// ============================================================
function buildFileName(template, vars) {
    var result = template;
    for (var key in vars) {
        if (vars.hasOwnProperty(key)) {
            result = result.replace(
                new RegExp("\\{" + key + "\\}", "g"), vars[key]
            );
        }
    }
    return sanitize(result, 150);
}

// ============================================================
// SAYFA YARDIMCILARI
// ============================================================
function filterByMaster(doc, indices, masterName) {
    var filtered = [];
    for (var i = 0; i < indices.length; i++) {
        var page = doc.pages[indices[i]];
        try {
            var applied = page.appliedMaster;
            if (applied && applied.name === masterName) {
                filtered.push(indices[i]);
            }
        } catch(e) {}
    }
    return filtered;
}

function isPageEmpty(page) {
    try {
        var items = page.allPageItems;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            // TextFrame ise ve ici dolu mu?
            if (item instanceof TextFrame) {
                if (item.contents && item.contents.replace(/\s/g,"") !== "")
                    return false;
            } else if (item instanceof Rectangle ||
                       item instanceof Oval      ||
                       item instanceof Polygon   ||
                       item instanceof GraphicLine) {
                // Baska grafik nesnesi varsa dolu kabul et
                return false;
            }
        }
        return true;
    } catch(e) { return false; }
}

function applyLayerVisibility(doc, layerStates) {
    var original = {};
    if (!layerStates) return original;
    try {
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            original[layer.name] = layer.visible;
            if (layerStates[layer.name] !== undefined) {
                layer.visible = layerStates[layer.name];
            }
        }
    } catch(e) {}
    return original;
}

function restoreLayerVisibility(doc, original) {
    try {
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            if (original[layer.name] !== undefined) {
                layer.visible = original[layer.name];
            }
        }
    } catch(e) {}
}

// ============================================================
// SAYFA ARALIGI PARSER
// "1-5" -> [0..4]   "2,4,7" -> [1,3,6]   "1-3,7,9-11" -> karisik
// ============================================================
function resolvePageRange(rangeStr, totalPages) {
    var indices = [];
    var parts   = rangeStr.replace(/\s/g, "").split(",");
    for (var i = 0; i < parts.length; i++) {
        var part = parts[i];
        if (part.indexOf("-") !== -1) {
            var b    = part.split("-");
            var from = Math.max(1, parseInt(b[0], 10));
            var to   = Math.min(totalPages, parseInt(b[1], 10));
            if (!isNaN(from) && !isNaN(to))
                for (var n = from; n <= to; n++) indices.push(n - 1);
        } else {
            var pg = parseInt(part, 10);
            if (!isNaN(pg) && pg >= 1 && pg <= totalPages)
                indices.push(pg - 1);
        }
    }
    // Tekrarlari sil ve sirala
    var seen = {}, unique = [];
    for (var j = 0; j < indices.length; j++) {
        if (!seen[indices[j]]) { seen[indices[j]] = true; unique.push(indices[j]); }
    }
    unique.sort(function(a,b){ return a-b; });
    return unique;
}

// ============================================================
// METIN CEKME
// ============================================================
function getFirstParagraphText(page, paraStyle, charStyle) {
    var frames = [];
    var items  = page.allPageItems;
    for (var i = 0; i < items.length; i++) {
        if (items[i] instanceof TextFrame) frames.push(items[i]);
        else {
            var nested = collectTextFrames(items[i]);
            for (var n = 0; n < nested.length; n++) frames.push(nested[n]);
        }
    }
    frames.sort(function(a,b) {
        try {
            var ay=a.geometricBounds[0], ax=a.geometricBounds[1];
            var by=b.geometricBounds[0], bx=b.geometricBounds[1];
            return (Math.abs(ay-by)>1) ? ay-by : ax-bx;
        } catch(e){ return 0; }
    });
    for (var f = 0; f < frames.length; f++) {
        try {
            var tf = frames[f];
            for (var p = 0; p < tf.paragraphs.length; p++) {
                var para = tf.paragraphs[p];
                if (para.appliedParagraphStyle.name !== paraStyle.name) continue;
                var txt = charStyle
                          ? extractCharStyleText(para, charStyle)
                          : para.contents + "";
                if (txt.replace(/\s/g,"") !== "") return txt;
            }
        } catch(e) {}
    }
    return "";
}

function extractCharStyleText(para, charStyle) {
    var result = "";
    try {
        var ranges = para.textStyleRanges;
        for (var r = 0; r < ranges.length; r++) {
            if (ranges[r].appliedCharacterStyle.name === charStyle.name) {
                var piece = ranges[r].contents + "";
                if (piece.replace(/\s/g,"") !== "") { result = piece; break; }
            }
        }
    } catch(e) {}
    return result;
}

// ============================================================
// YARDIMCILAR
// ============================================================
function collectTextFrames(container) {
    var result = [];
    try {
        var items = container.pageItems;
        for (var i = 0; i < items.length; i++) {
            if (items[i] instanceof TextFrame) result.push(items[i]);
            else result = result.concat(collectTextFrames(items[i]));
        }
    } catch(e) {}
    return result;
}

function safeGetStyle(collection, name) {
    try {
        var s = collection.itemByName(name);
        return s.isValid ? s : null;
    } catch(e) { return null; }
}

function sanitize(text, maxLen) {
    if (typeof text !== "string") return "";
    var s = text
        .replace(/[\r\n\t]/g," ")
        .replace(/[\/\\:*?"<>|]/g,"")
        .replace(/\s+/g," ")
        .replace(/^\s+|\s+$/g,"");
    if (s.length > maxLen) s = s.substring(0, maxLen).replace(/\s+$/,"");
    return s;
}

function uniquePath(folder, baseName, ext) {
    var base = folder + "/" + baseName + "." + ext;
    if (!new File(base).exists) return base;
    for (var i = 2; i < 9999; i++) {
        var p = folder + "/" + baseName + "_" + i + "." + ext;
        if (!new File(p).exists) return p;
    }
    return base;
}

function zeroPad(n, width) {
    var s = String(n);
    while (s.length < width) s = "0" + s;
    return s;
}

function getDateString() {
    var d = new Date();
    return d.getFullYear() + "-"
         + zeroPad(d.getMonth()+1, 2) + "-"
         + zeroPad(d.getDate(), 2);
}

function shortenPath(p, max) {
    return (p.length > max) ? "..." + p.slice(-max) : p;
}

function findIndex(arr, val) {
    if (!val) return 0;
    for (var i = 0; i < arr.length; i++)
        if (arr[i] === val) return i;
    return 0;
}
