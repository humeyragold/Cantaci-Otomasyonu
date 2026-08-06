/**
 * 14 AYAR ÇANTACI OTOMASYON SİSTEMİ - MASTER MOTORU (V68 - DÖNEMSEL RAPOR REVİZYONU)
 * - RAPOR_DONEMSEL sayfasında Milyem filtresine virgülle çoklu seçim özelliği eklendi (Örn: 725, 735).
 * - RAPOR_DONEMSEL sayfasında listelenen verilerin en altına otomatik hesaplanan "GENEL TOPLAM" satırı eklendi.
 * * 🟢 ÇİZİLEN BUTONLARA ATANACAK FONKSİYON (MAKRO) İSİMLERİ:
 * 1. "KAYDET" Butonu    ->  butonKaydiOnayla
 * 2. "ARŞİVLE" Butonu   ->  butonArsiveTasi
 * 3. "TEMİZLE" Butonu   ->  butonFormuTemizle
 * 4. "FİŞ PDF" Butonu   ->  butonPdfFis
 * 5. "EKSTRE PDF" Butonu->  butonPdfEkstre
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('💎 Çantacı Otomasyonu')
    .addItem('🚀 Sistemi Güncelle (Veriler Korunur)', 'masterSifirKurulum')
    .addItem('📅 Eski Tarihleri Onar ve Sırala', 'tarihleriOnarVeSirala')
    .addSeparator()
    .addItem('🔄 Dashboardları Manuel Senkronize Et', 'guncelleDashboards')
    .addSeparator()
    .addItem('🎨 Renkleri Onar (Geçmişi Düzelt)', 'gecmisiRenklendir')
    .addSeparator()
    .addItem('🗑️ Hatalı Fiş İptal Et', 'fisIptalEt')
    .addSeparator()
    .addItem('📑 Sayfaları A\'dan Z\'ye Sırala', 'sayfalariSirala')
    .addToUi();
}

/**
 * ⚙️ TARİH ONARIM MOTORU (ESKİ KAYITLARI DÜZELTİR)
 */
function tarihleriOnarVeSirala() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ISLEMLER");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("⚠️ ISLEMLER sayfası bulunamadı.");
    return;
  }

  var lr = sheet.getLastRow();
  if (lr < 2) {
    SpreadsheetApp.getUi().alert("⚠️ Sıralanacak veri bulunamadı.");
    return;
  }

  var range = sheet.getRange(2, 2, lr - 1, 1);
  var values = range.getValues();
  var degisiklikYapildi = false;

  for (var i = 0; i < values.length; i++) {
    var val = values[i][0];
    
    if (typeof val === 'string' && val.trim() !== "") {
      var str = val.trim();
      var d = null;
      
      if (str.indexOf('.') > -1) {
        var parts = str.split('.');
        if (parts.length === 3) {
          d = new Date(parts[2], parseInt(parts[1], 10) - 1, parts[0]);
        }
      }
      else if (str.indexOf('-') > -1) {
        var parts = str.split('-');
        if (parts.length === 3) {
          d = new Date(parts[0], parseInt(parts[1], 10) - 1, parts[2]);
        }
      }

      if (d && !isNaN(d.getTime())) {
        values[i][0] = d;
        degisiklikYapildi = true;
      }
    }
  }

  if (degisiklikYapildi) {
    range.setValues(values);
  }

  range.setNumberFormat("dd.MM.yyyy");
  sheet.getRange(2, 1, lr - 1, 9).sort({column: 2, ascending: true});
  
  guncelleDashboards();
  hesaplaDonemselRapor();
  var mOzetSheet = ss.getSheetByName("MUSTERI_OZET");
  if(mOzetSheet) {
    var aktifMusteri = mOzetSheet.getRange("B4").getValue();
    if(aktifMusteri) hesaplaMusteriOzeti(aktifMusteri);
  }

  SpreadsheetApp.getUi().alert("✅ Harika!\n\nTüm eski inatçı metinler gerçek tarihe dönüştürüldü ve kayıtlar tarih sırasına dizildi.");
}

/**
 * ⚙️ AYARLAR SAYFASI
 */
function kurAyarlarSayfasi(ss) {
  var sheet = ss.getSheetByName("AYARLAR");
  if (!sheet) {
    sheet = ss.insertSheet("AYARLAR", 0); 
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 350);
    sheet.getRange("A1:B2").merge().setValue("⚙️ SİSTEM AYARLARI (V68)")
         .setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold")
         .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(14);
    sheet.getRange("A3:B3").merge().setValue("GENEL AYARLAR").setBackground("#bdc3c7").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A4").setValue("Sistem Başlığı:");
    sheet.getRange("B4").setValue("💎 14 AYAR ÇANTACI OTOMASYON SİSTEMİ");
    
    sheet.getRange("A5").setValue("Kurumsal Ana Renk (Koyu):");
    sheet.getRange("B5").setValue("Boya Kutusundan Hücreyi Boyayın ->").setBackground("#2c3e50").setFontColor("#ffffff");
    sheet.getRange("A6").setValue("Kurumsal Yazı Rengi:");
    sheet.getRange("B6").setValue("Boya Kutusundan Hücreyi Boyayın ->").setBackground("#ffffff").setFontColor("#000000");
    
    sheet.getRange("A7").setValue("Varsayılan Sistem Milyemi:");
    sheet.getRange("B7").setValue(585);
    
    sheet.getRange("A9:B9").merge().setValue("İŞLEM RENKLERİ (Hücre Arka Planlarını Boyayın)").setBackground("#bdc3c7").setFontWeight("bold").setHorizontalAlignment("center");
    var islemler = [
      ["SATIŞ", "#2980b9"], ["ALIŞ", "#27ae60"], ["TAHSİLAT", "#c0392b"], ["ÖDEME", "#c0392b"],
      ["MÜŞTERİ İADESİ", "#d35400"], ["ATÖLYE İADESİ", "#d35400"], ["DEVİR (MÜŞTERİ)", "#8e44ad"],
      ["DEVİR (ATÖLYE)", "#8e44ad"], ["SERMAYE GİRİŞİ", "#16a085"]
    ];
    for(var i=0; i<islemler.length; i++){
      sheet.getRange(10+i, 1).setValue(islemler[i][0]).setFontWeight("bold");
      sheet.getRange(10+i, 2).setValue("Renk İçin Bu Hücreyi Boyayın ->").setBackground(islemler[i][1]).setFontColor("#ffffff");
    }
    
    sheet.getRange("A4:A18").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheet.getRange("B4:B18").setHorizontalAlignment("center").setVerticalAlignment("middle");
  }
  return sheet;
}

function getAyarlar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("AYARLAR");
  var ayarlar = {
    baslik: "💎 14 AYAR ÇANTACI OTOMASYON SİSTEMİ",
    kurumsalRenk: "#2c3e50",
    yaziRengi: "#ffffff",
    varsayilanMilyem: 585,
    islemRenkleri: {
      "SATIŞ": "#2980b9", "ALIŞ": "#27ae60", "TAHSİLAT": "#c0392b",
      "ÖDEME": "#c0392b", "MÜŞTERİ İADESİ": "#d35400", "ATÖLYE İADESİ": "#d35400",
      "DEVİR (MÜŞTERİ)": "#8e44ad", "DEVİR (ATÖLYE)": "#8e44ad", "SERMAYE GİRİŞİ": "#16a085"
    }
  };
  if (sheet) {
    ayarlar.baslik = sheet.getRange("B4").getValue() || ayarlar.baslik;
    ayarlar.kurumsalRenk = sheet.getRange("B5").getBackground();
    ayarlar.yaziRengi = sheet.getRange("B6").getBackground();
    var m = parseInt(sheet.getRange("B7").getValue());
    if(!isNaN(m) && m>0) ayarlar.varsayilanMilyem = m;
    var islemler = ["SATIŞ", "ALIŞ", "TAHSİLAT", "ÖDEME", "MÜŞTERİ İADESİ", "ATÖLYE İADESİ", "DEVİR (MÜŞTERİ)", "DEVİR (ATÖLYE)", "SERMAYE GİRİŞİ"];
    for(var i=0; i<islemler.length; i++) {
      ayarlar.islemRenkleri[islemler[i]] = sheet.getRange(10+i, 2).getBackground();
    }
  }
  return ayarlar;
}

/**
 * 🚀 SİSTEM GÜNCELLEME VE KUSURSUZ MİGRASYON MOTORU
 */
function masterSifirKurulum() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  kurAyarlarSayfasi(ss);
  var ayarlar = getAyarlar(); 
  var kurumsalRenk = ayarlar.kurumsalRenk;
  var yaziRengi = ayarlar.yaziRengi;
  var takvimKurali = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();

  var cKartlarSheet = ss.getSheetByName("CARI_KARTLAR");
  if (!cKartlarSheet) {
    cKartlarSheet = ss.insertSheet("CARI_KARTLAR");
    var cHeaders = ["Kayıt Tarihi", "Cari Adı", "Cari Tipi", "Durum", "Açılış Has Bakiye", "Açılış Adet Bakiye"];
    cKartlarSheet.getRange(1, 1, 1, cHeaders.length).setValues([cHeaders])
                 .setBackground(kurumsalRenk).setFontColor(yaziRengi).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    cKartlarSheet.getRange("A:A").setNumberFormat("dd.MM.yyyy");
    cKartlarSheet.getRange("E:E").setNumberFormat("#,##0.00");
    cKartlarSheet.getRange("F:F").setNumberFormat("0");
    cKartlarSheet.setFrozenRows(1);
    
    var islemlerSheet = ss.getSheetByName("ISLEMLER");
    if (islemlerSheet && islemlerSheet.getLastRow() > 1) {
      var data = islemlerSheet.getDataRange().getValues();
      var tempCariler = {};
      for (var i = 1; i < data.length; i++) {
        var cari = data[i][2];
        var tip = data[i][3];
        if(!cari) continue;
        if (!tempCariler[cari]) tempCariler[cari] = { mPuan: 0, aPuan: 0 };
        if (tip === "SATIŞ" || tip === "DEVİR (MÜŞTERİ)") tempCariler[cari].mPuan += 100;
        if (tip === "TAHSİLAT" || tip === "MÜŞTERİ İADESİ") tempCariler[cari].mPuan += 10;
        if (tip === "ALIŞ" || tip === "DEVİR (ATÖLYE)" || tip === "SERMAYE GİRİŞİ") tempCariler[cari].aPuan += 100;
        if (tip === "ÖDEME" || tip === "ATÖLYE İADESİ") tempCariler[cari].aPuan += 10;
      }
      var exportData = [];
      var keys = Object.keys(tempCariler);
      for(var k=0; k < keys.length; k++) {
        var cType = (tempCariler[keys[k]].aPuan > tempCariler[keys[k]].mPuan) ? "ATÖLYE" : "MÜŞTERİ";
        var es = ss.getSheetByName(keys[k]);
        var durum = (es && es.getRange("Z1").getValue() === "ARŞİVLENDİ") ? "ARŞİVLENDİ" : "AKTİF";
        var txtTarih = new Date();
        exportData.push([txtTarih, keys[k], cType, durum, "", ""]);
      }
      if(exportData.length > 0) {
        cKartlarSheet.getRange(2, 1, exportData.length, 6).setValues(exportData);
      }
    }
  }
  cKartlarSheet.getRange("A:A").setNumberFormat("dd.MM.yyyy");
  cKartlarSheet.getRange("A2:F").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  cKartlarSheet.setColumnWidth(1, 130); cKartlarSheet.setColumnWidth(2, 250); cKartlarSheet.setColumnWidth(3, 150); 
  cKartlarSheet.setColumnWidth(4, 130); cKartlarSheet.setColumnWidth(5, 160); cKartlarSheet.setColumnWidth(6, 160);

  var islemlerSheet = ss.getSheetByName("ISLEMLER");
  if (!islemlerSheet) {
    islemlerSheet = ss.insertSheet("ISLEMLER");
    var islemlerHeaders = ["İşlem ID", "Tarih", "Cari Adı", "İşlem Tipi", "İşlem Gramı", "Milyem", "Has", "Adet", "Açıklama"];
    islemlerSheet.getRange(1, 1, 1, islemlerHeaders.length).setValues([islemlerHeaders])
                 .setBackground(kurumsalRenk).setFontColor(yaziRengi).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    islemlerSheet.getRange("B:B").setNumberFormat("dd.MM.yyyy");
    islemlerSheet.getRange("E:E").setNumberFormat("#,##0.00"); 
    islemlerSheet.getRange("F:F").setNumberFormat("0"); 
    islemlerSheet.getRange("G:G").setNumberFormat("#,##0.00"); 
    islemlerSheet.getRange("H:H").setNumberFormat("0"); 
    islemlerSheet.setFrozenRows(1);
  } else {
    islemlerSheet.getRange("E1").setValue("İşlem Gramı");
    islemlerSheet.getRange("B:B").setNumberFormat("dd.MM.yyyy");
  }

  var lorettaSheet = ss.getSheetByName("LORETTA_BORC");
  var oldLorettaData = [];
  if (lorettaSheet) {
    if (lorettaSheet.getLastRow() > 1) {
      var firstHead = lorettaSheet.getRange("A1").getValue();
      if (firstHead === "İşlem ID") {
         oldLorettaData = lorettaSheet.getRange(2, 1, lorettaSheet.getLastRow() - 1, Math.max(lorettaSheet.getLastColumn(), 8)).getValues();
      } else if (lorettaSheet.getLastRow() > 7) {
         oldLorettaData = lorettaSheet.getRange(8, 1, lorettaSheet.getLastRow() - 7, Math.max(lorettaSheet.getLastColumn(), 8)).getValues();
      }
    }
    lorettaSheet.setFrozenRows(0);
    lorettaSheet.clear();
    lorettaSheet.getRange("A1:Z1000").clearDataValidations();
  } else {
    lorettaSheet = ss.insertSheet("LORETTA_BORC");
  }

  lorettaSheet.getRange("A1:H2").merge().setValue("👑 LORETTA TEDARİK VE SATIŞ ÖZETİ")
              .setBackground("#8e44ad").setFontColor("#ffffff").setFontWeight("bold")
              .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(14);
  lorettaSheet.getRange("A3").setValue("Başlangıç Tarihi:").setFontWeight("bold").setHorizontalAlignment("right");
  lorettaSheet.getRange("B3:C3").merge().setBackground("#fdfecd").setHorizontalAlignment("center").setDataValidation(takvimKurali);
  lorettaSheet.getRange("A4").setValue("Bitiş Tarihi:").setFontWeight("bold").setHorizontalAlignment("right");
  lorettaSheet.getRange("B4:C4").merge().setBackground("#fdfecd").setHorizontalAlignment("center").setDataValidation(takvimKurali);

  lorettaSheet.getRange("D3:E3").merge().setValue("Seçili Dönem Satılan Gram:").setFontWeight("bold").setHorizontalAlignment("right");
  var filterFormulaGram = '=IFERROR(SUMIFS(E8:E; D8:D; "SATIŞ"; B8:B; IF(B3<>"";">="&B3;">=1.1.1900"); B8:B; IF(B4<>"";"<="&B4;"<=1.1.2100")) - SUMIFS(E8:E; D8:D; "MÜŞTERİ İADESİ"; B8:B; IF(B3<>"";">="&B3;">=1.1.1900"); B8:B; IF(B4<>"";"<="&B4;"<=1.1.2100")); 0)';
  lorettaSheet.getRange("F3:H3").merge().setFormula(filterFormulaGram)
              .setBackground("#f1c40f").setFontWeight("bold").setHorizontalAlignment("center").setNumberFormat("#,##0.00");

  lorettaSheet.getRange("D4:E4").merge().setValue("Seçili Dönem Toplam Has:").setFontWeight("bold").setHorizontalAlignment("right");
  var filterFormulaHas = '=IFERROR(SUMIFS(G8:G; D8:D; "SATIŞ"; B8:B; IF(B3<>"";">="&B3;">=1.1.1900"); B8:B; IF(B4<>"";"<="&B4;"<=1.1.2100")) - SUMIFS(G8:G; D8:D; "MÜŞTERİ İADESİ"; B8:B; IF(B3<>"";">="&B3;">=1.1.1900"); B8:B; IF(B4<>"";"<="&B4;"<=1.1.2100")); 0)';
  lorettaSheet.getRange("F4:H4").merge().setFormula(filterFormulaHas)
              .setBackground("#27ae60").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setNumberFormat("#,##0.00");

  var lHeaders = ["İşlem ID", "Tarih", "Cari Adı", "İşlem Tipi", "İşlem Gramı", "Milyem", "Has", "Açıklama"];
  lorettaSheet.getRange(7, 1, 1, 8).setValues([lHeaders])
              .setBackground(kurumsalRenk).setFontColor(yaziRengi).setFontWeight("bold").setHorizontalAlignment("center");
  lorettaSheet.getRange("B:B").setNumberFormat("dd.MM.yyyy").setHorizontalAlignment("center");
  lorettaSheet.getRange("E:E").setNumberFormat("#,##0.00").setHorizontalAlignment("center");
  lorettaSheet.getRange("F:F").setNumberFormat("0").setHorizontalAlignment("center");
  lorettaSheet.getRange("G:G").setNumberFormat("#,##0.00").setHorizontalAlignment("center");
  lorettaSheet.setFrozenRows(7);

  var exportLoretta = [];
  if (oldLorettaData.length > 0) {
    var islemlerData = ss.getSheetByName("ISLEMLER").getDataRange().getValues();
    var islemMap = {};
    for (var m = 1; m < islemlerData.length; m++) {
       islemMap[islemlerData[m][0]] = islemlerData[m]; 
    }
    
    for(var x = 0; x < oldLorettaData.length; x++) {
       var r = oldLorettaData[x];
       var id = r[0];
       if (!id) continue;
       var islemRow = islemMap[id];
       if (islemRow) {
           exportLoretta.push([islemRow[0], islemRow[1], islemRow[2], islemRow[3], islemRow[4], islemRow[5], islemRow[6], islemRow[8]]);
       } else {
           var gr = parseFloat(r[4] || r[3]) || 0;
           exportLoretta.push([id, r[1], r[2], r[3] || "SATIŞ", gr, "", r[5] || "", "Eski Kayıt"]);
       }
    }
  }
  if (exportLoretta.length > 0) {
     lorettaSheet.getRange(8, 1, exportLoretta.length, 8).setValues(exportLoretta);
  }

  var girisSheet = ss.getSheetByName("GIRIS_FORMU") || ss.insertSheet("GIRIS_FORMU");
  girisSheet.setFrozenRows(0);
  girisSheet.clear();
  girisSheet.getRange("A1:G2").merge().setValue(ayarlar.baslik)
            .setBackground(kurumsalRenk).setFontColor(yaziRengi).setFontWeight("bold")
            .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(14).setWrap(true);
  girisSheet.getRange("B3").setValue("ATÖLYE").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  girisSheet.getRange("C3").setValue("MÜŞTERİ").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  girisSheet.getRange("A4").setValue("Cari Adı:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  girisSheet.getRange("A5").setValue("İşlem Tarihi:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  girisSheet.getRange("B5").setValue("=TODAY()").setNumberFormat("dd.MM.yyyy").setHorizontalAlignment("center").setVerticalAlignment("middle").setDataValidation(takvimKurali);
  
  girisSheet.getRange("D3").setValue("BAKİYE (HAS)").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("#34495e").setFontColor("#ffffff").setWrap(true);
  var bakiyeFormulu = '=IF(C4<>""; SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; C4; ISLEMLER!D:D; "SATIŞ") + SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; C4; ISLEMLER!D:D; "DEVİR (MÜŞTERİ)") - SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; C4; ISLEMLER!D:D; "TAHSİLAT") - SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; C4; ISLEMLER!D:D; "MÜŞTERİ İADESİ"); IF(B4<>""; SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; B4; ISLEMLER!D:D; "ALIŞ") + SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; B4; ISLEMLER!D:D; "DEVİR (ATÖLYE)") + SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; B4; ISLEMLER!D:D; "SERMAYE GİRİŞİ") - SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; B4; ISLEMLER!D:D; "ÖDEME") - SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; B4; ISLEMLER!D:D; "ATÖLYE İADESİ"); ""))';
  girisSheet.getRange("D4").setFormula(bakiyeFormulu).setNumberFormat("#,##0.00").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("#f1c40f").setFontColor("#2c3e50").setFontSize(12);

  girisSheet.getRange("E3").setValue("YENİ BAKİYE").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("#2980b9").setFontColor("#ffffff").setWrap(true);
  var yeniBakiyeFormulu = '=IF(C4<>""; D4 + SUMIFS(D8:D17; A8:A17; "SATIŞ") + SUMIFS(D8:D17; A8:A17; "DEVİR (MÜŞTERİ)") - SUMIFS(D8:D17; A8:A17; "TAHSİLAT") - SUMIFS(D8:D17; A8:A17; "MÜŞTERİ İADESİ"); IF(B4<>""; D4 + SUMIFS(D8:D17; A8:A17; "ALIŞ") + SUMIFS(D8:D17; A8:A17; "DEVİR (ATÖLYE)") + SUMIFS(D8:D17; A8:A17; "SERMAYE GİRİŞİ") - SUMIFS(D8:D17; A8:A17; "ÖDEME") - SUMIFS(D8:D17; A8:A17; "ATÖLYE İADESİ"); ""))';
  girisSheet.getRange("E4").setFormula(yeniBakiyeFormulu).setNumberFormat("#,##0.00").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("#ebf5fb").setFontColor("#2c3e50").setFontSize(12);

  var formHeaders = ["İşlem Tipi", "İşlem Gramı", "Milyem (XXX)", "Has (Canlı)", "Loretta", "Adet", "Ürün / Açıklama"];
  girisSheet.getRange(7, 1, 1, formHeaders.length).setValues([formHeaders])
            .setBackground("#16a085").setFontColor("#ffffff").setFontWeight("bold")
            .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  girisSheet.getRange(7, 5).setBackground("#8e44ad"); 

  girisSheet.getRange("B8:B17").setNumberFormat("#,##0.00").setVerticalAlignment("middle"); 
  girisSheet.getRange("C8:C17").setNumberFormat("0").setVerticalAlignment("middle"); 
  girisSheet.getRange("D8:D17").setNumberFormat("#,##0.00").setVerticalAlignment("middle"); 
  girisSheet.getRange("E8:E17").insertCheckboxes().setVerticalAlignment("middle").setHorizontalAlignment("center");
  girisSheet.getRange("F8:F17").setNumberFormat("0").setVerticalAlignment("middle");
  girisSheet.getRange("G8:G17").setVerticalAlignment("middle");
  
  girisSheet.getRange("D8:D17").setFormula('=IF(AND(B8>0; C8>0); ROUND(B8*C8/1000; 2); "")')
            .setBackground("#fbfcfc").setFontColor(kurumsalRenk).setFontWeight("bold").setVerticalAlignment("middle");
  var islemlerList = ["SATIŞ", "TAHSİLAT", "ALIŞ", "ÖDEME", "MÜŞTERİ İADESİ", "ATÖLYE İADESİ", "SERMAYE GİRİŞİ"];
  var typeRule = SpreadsheetApp.newDataValidation().requireValueInList(islemlerList, true).build();
  girisSheet.getRange("A8:A17").setDataValidation(typeRule).setVerticalAlignment("middle");
  
  girisSheet.getRange("F19:G26").clearContent().clearDataValidations().setBackground(null);
  
  var mOzetSheet = ss.getSheetByName("MUSTERI_OZET") || ss.insertSheet("MUSTERI_OZET");
  mOzetSheet.setFrozenRows(0);
  mOzetSheet.clear();
  mOzetSheet.getRange("A1:Z1000").clearDataValidations(); 
  
  mOzetSheet.getRange("A1:I2").merge().setValue("📊 MÜŞTERİ / ATÖLYE ÖZET PANELİ")
            .setBackground(kurumsalRenk).setFontColor(yaziRengi).setFontWeight("bold")
            .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(14).setWrap(true);
  mOzetSheet.getRange("A4").setValue("Seçilen Cari:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  mOzetSheet.getRange("B4:D4").merge().setBackground("#f1f2f6").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle"); 
  
  mOzetSheet.getRange("G3").setValue("Başlangıç Tarihi:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  mOzetSheet.getRange("H3:I3").merge().setBackground("#fdfecd").setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("dd.MM.yyyy").setDataValidation(takvimKurali).clearContent(); 
  
  mOzetSheet.getRange("G4").setValue("Bitiş Tarihi:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  mOzetSheet.getRange("H4:I4").merge().setBackground("#fdfecd").setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("dd.MM.yyyy").setDataValidation(takvimKurali).clearContent(); 

  mOzetSheet.getRange("B5").setValue("İŞLEM GRAMI").setBackground("#95a5a6").setFontColor("#ffffff").setFontWeight("bold")
            .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  mOzetSheet.getRange("C5").setValue("HAS ALTIN").setBackground("#f1c40f").setFontWeight("bold")
            .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  mOzetSheet.getRange("D5").setValue("ADET (PARÇA)").setBackground("#3498db").setFontColor("#ffffff").setFontWeight("bold")
            .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);

  mOzetSheet.getRange("A6").setValue("Dönem Giriş (Borçlandırma):").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  mOzetSheet.getRange("A7").setValue("Dönem İade:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  mOzetSheet.getRange("A8").setValue("Dönem Tahsilat / Ödeme:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle").setWrap(true);
  mOzetSheet.getRange("A9").setValue("DÖNEM NETİ:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  mOzetSheet.getRange("A10").setValue("GÜNCEL BAKİYE (Tüm Zamanlar):").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  
  mOzetSheet.getRange("B6:C10").setNumberFormat("#,##0.00").setVerticalAlignment("middle").setHorizontalAlignment("center");
  mOzetSheet.getRange("D6:D10").setNumberFormat("0").setVerticalAlignment("middle").setHorizontalAlignment("center");
  mOzetSheet.getRange("A9:D9").setFontWeight("bold").setFontSize(11).setBackground("#ecf0f1");
  mOzetSheet.getRange("A10:D10").setFontWeight("bold").setFontSize(12);
  mOzetSheet.getRange("G6:I6").merge().setValue("📈 KÂR VE MİLYEM ANALİZİ").setBackground("#8e44ad").setFontColor("#ffffff").setFontWeight("bold")
            .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
            
  mOzetSheet.getRange("G7").setValue("Ort. Satış Milyemi:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  mOzetSheet.getRange("H7:I7").merge().setNumberFormat("0").setFontColor("#2980b9").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  mOzetSheet.getRange("G8").setValue("Net İşçilik Kârı (Has):").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  mOzetSheet.getRange("H8:I8").merge().setNumberFormat("#,##0.00").setFontColor("#27ae60").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

  mOzetSheet.getRange("G10").setValue("Son İşlem Tarihi:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  mOzetSheet.getRange("H10:I10").merge().setNumberFormat("dd.MM.yyyy").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  var txHeaders = ["Tarih", "İşlem Tipi", "İşlem Gramı", "Milyem (XXX)", "Has Altın", "Yürüyen Bakiye", "Adet", "Açıklama", "İşlem ID (Belge)"];
  mOzetSheet.getRange(12, 1, 1, txHeaders.length).setValues([txHeaders])
            .setBackground("#34495e").setFontColor("#ffffff").setFontWeight("bold")
            .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  mOzetSheet.getRange("A13:A").setNumberFormat("dd.MM.yyyy");
  mOzetSheet.getRange("B13:B").setHorizontalAlignment("center");
  mOzetSheet.getRange("C13:C").setNumberFormat("#,##0.00");
  mOzetSheet.getRange("D13:D").setNumberFormat("0");
  mOzetSheet.getRange("E13:F").setNumberFormat("#,##0.00"); 
  mOzetSheet.getRange("G13:G").setNumberFormat("0");
  mOzetSheet.getRange("H13:I").setHorizontalAlignment("center");
  mOzetSheet.setFrozenRows(12);

  var donemRaporSheet = ss.getSheetByName("RAPOR_DONEMSEL") || ss.insertSheet("RAPOR_DONEMSEL");
  donemRaporSheet.setFrozenRows(0);
  donemRaporSheet.clear();
  donemRaporSheet.getRange("A1:Z1000").clearDataValidations();
  donemRaporSheet.getRange("A1:G2").merge().setValue("📊 DÖNEMSEL PERFORMANS VE MİLYEM ANALİZ RAPORU")
                 .setBackground("#8e44ad").setFontColor("#ffffff").setFontWeight("bold")
                 .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(14).setWrap(true);
  donemRaporSheet.getRange("A3").setValue("Başlangıç Tarihi:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  donemRaporSheet.getRange("B3").setBackground("#fdfecd").setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("dd.MM.yyyy").setDataValidation(takvimKurali);
  donemRaporSheet.getRange("A4").setValue("Bitiş Tarihi:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  donemRaporSheet.getRange("B4").setBackground("#fdfecd").setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("dd.MM.yyyy").setDataValidation(takvimKurali);
  
  donemRaporSheet.getRange("D3").setValue("Cari Tipi Filtresi:").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  var rTipRule = SpreadsheetApp.newDataValidation().requireValueInList(["TÜMÜ", "MÜŞTERİ", "ATÖLYE"], true).build();
  donemRaporSheet.getRange("E3:G3").merge().setBackground("#ecf0f1").setHorizontalAlignment("center").setVerticalAlignment("middle").setDataValidation(rTipRule).setValue("TÜMÜ");
  
  // V68: Milyem başlığı çoklu seçim bilgisiyle güncellendi
  donemRaporSheet.getRange("D4").setValue("Milyem Filtresi (Çoklu Seçim İçin Virgül Kullanın):").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle").setFontColor("#c0392b");
  donemRaporSheet.getRange("E4:G4").merge().setBackground("#ecf0f1").setHorizontalAlignment("center").setVerticalAlignment("middle").setValue("TÜMÜ").setFontWeight("bold");

  var rpHeaders = ["Cari Adı", "Toplam Çıkış (Gr)", "Toplam İade (Gr)", "Net İşlem (Gr)", "Ort. Milyem", "Tahsilat/Ödeme (Has)", "Dönem Kârı (Has)"];
  donemRaporSheet.getRange(6, 1, 1, rpHeaders.length).setValues([rpHeaders])
                 .setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold")
                 .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  donemRaporSheet.getRange("B7:D").setNumberFormat("#,##0.00");
  donemRaporSheet.getRange("E7:E").setNumberFormat("0");
  donemRaporSheet.getRange("F7:G").setNumberFormat("#,##0.00");
  donemRaporSheet.setFrozenRows(6);
  
  donemRaporSheet.setColumnWidth(1, 200); donemRaporSheet.setColumnWidth(2, 140); donemRaporSheet.setColumnWidth(3, 140); 
  donemRaporSheet.setColumnWidth(4, 150); donemRaporSheet.setColumnWidth(5, 120); donemRaporSheet.setColumnWidth(6, 170); 
  donemRaporSheet.setColumnWidth(7, 160);
  
  girisSheet.getRange("I2:J2").merge().setValue("📊 DASHBOARD_STOK (Çanta Durumu)").setBackground("#7f8c8d").setFontColor("#ffffff").setFontWeight("bold")
            .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  var mDef = ayarlar.varsayilanMilyem;
  var stokKriterler = [
    ["Çantadaki Güncel Fiziksel Ürün (Gram)", '=SUMIFS(ISLEMLER!E:E; ISLEMLER!D:D; "ALIŞ") + SUMIFS(ISLEMLER!E:E; ISLEMLER!D:D; "SERMAYE GİRİŞİ") + SUMIFS(ISLEMLER!E:E; ISLEMLER!D:D; "MÜŞTERİ İADESİ") - SUMIFS(ISLEMLER!E:E; ISLEMLER!D:D; "SATIŞ") - SUMIFS(ISLEMLER!E:E; ISLEMLER!D:D; "ATÖLYE İADESİ")'],
    ["Çantadaki Güncel Has Altın (Gram)", '=SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "ALIŞ") + SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "SERMAYE GİRİŞİ") + SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "MÜŞTERİ İADESİ") - SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "SATIŞ") - SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "ATÖLYE İADESİ")'],
    ["Çantadaki Malın Ort. Giriş Milyemi", '=IFERROR(((SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "ALIŞ") + SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "SERMAYE GİRİŞİ") - SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "ATÖLYE İADESİ")) * 1000) / (SUMIFS(ISLEMLER!E:E; ISLEMLER!D:D; "ALIŞ") + SUMIFS(ISLEMLER!E:E; ISLEMLER!D:D; "SERMAYE GİRİŞİ") - SUMIFS(ISLEMLER!E:E; ISLEMLER!D:D; "ATÖLYE İADESİ")); 0)'],
    ["Ort. Giriş İşçiliği", '=IF(J5>0; J5-' + mDef + '; 0)'],
    ["Toplam Satılan Ürün (Gram)", '=SUMIFS(ISLEMLER!E:E; ISLEMLER!D:D; "SATIŞ") - SUMIFS(ISLEMLER!E:E; ISLEMLER!D:D; "MÜŞTERİ İADESİ")'],
    ["Ort. Satış Milyemi", '=IFERROR((SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "SATIŞ") - SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "MÜŞTERİ İADESİ")) * 1000 / J7; 0)'],
    ["Ort. Satış İşçiliği", '=IF(J8>0; J8-' + mDef + '; 0)'],
    ["Net (Saf) İşçilik Kârı (Milyem)", '=IF(AND(J9>0; J6>0); J9-J6; 0)'], 
    ["NET KÂR (Has Altın Gram)", '=(SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "SATIŞ") - SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "MÜŞTERİ İADESİ")) - (J7 * J5 / 1000)'],
    ["Piyasa Toplam Alacak", '=SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "SATIŞ") + SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "DEVİR (MÜŞTERİ)") - SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "TAHSİLAT") - SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "MÜŞTERİ İADESİ")'],
    ["Atölyeye Toplam Borç", '=SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "ALIŞ") + SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "DEVİR (ATÖLYE)") - SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "ÖDEME") - SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "ATÖLYE İADESİ")']
  ];
  for (var i = 0; i < stokKriterler.length; i++) {
    girisSheet.getRange(3 + i, 9).setValue(stokKriterler[i][0]).setFontWeight("bold").setBackground("#f8f9fa").setVerticalAlignment("middle").setWrap(true);
    girisSheet.getRange(3 + i, 10).setValue(stokKriterler[i][1]).setBackground("#f8f9fa").setVerticalAlignment("middle").setHorizontalAlignment("center");
    if (i === 2 || i === 3 || i === 5 || i === 6 || i === 7) {
      girisSheet.getRange(3 + i, 10).setNumberFormat("0");
    } else {
      girisSheet.getRange(3 + i, 10).setNumberFormat("#,##0.00");
    }
  }

  var dMusteri = ss.getSheetByName("DASHBOARD_MUSTERI") || ss.insertSheet("DASHBOARD_MUSTERI");
  dMusteri.clear();
  var mHeaders = ["Müşteri Adı", "Güncel Bakiye (Has Altın)", "Toplam Satılan (Gram)", "Ort. Satış Milyemi (XXX)", "Ort. Kazanılan İşçilik (Milyem)", "Son İşlem Tarihi"];
  dMusteri.getRange(1, 1, 1, mHeaders.length).setValues([mHeaders]).setBackground("#2980b9").setFontColor("#ffffff").setFontWeight("bold")
          .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  dMusteri.getRange("F:F").setNumberFormat("dd.MM.yyyy");
  dMusteri.setFrozenRows(1);
  var dAtolye = ss.getSheetByName("DASHBOARD_ATOLYE") || ss.insertSheet("DASHBOARD_ATOLYE");
  dAtolye.clear();
  var aHeaders = ["Atölye Adı", "Güncel Bakiye (Has Altın)", "Toplam Alınan (Gram)", "Ort. Giriş Milyemi (XXX)", "Ort. Ödenen İşçilik (Milyem)", "Son İşlem Tarihi"];
  dAtolye.getRange(1, 1, 1, aHeaders.length).setValues([aHeaders]).setBackground("#d35400").setFontColor("#ffffff").setFontWeight("bold")
         .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  dAtolye.getRange("F:F").setNumberFormat("dd.MM.yyyy");
  dAtolye.setFrozenRows(1);

  var dArsiv = ss.getSheetByName("CARI_ARSIV") || ss.insertSheet("CARI_ARSIV");
  dArsiv.clear();
  var arHeaders = ["Cari Adı", "Hesap Türü", "Güncel Bakiye (Has Altın)", "Toplam İşlem (Gram)", "Ort. Milyem (XXX)", "Ort. İşçilik", "Son İşlem Tarihi", "Arşivden Çıkar ♻️"];
  dArsiv.getRange(1, 1, 1, arHeaders.length).setValues([arHeaders]).setBackground("#34495e").setFontColor("#ffffff").setFontWeight("bold")
        .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  dArsiv.getRange("G:G").setNumberFormat("dd.MM.yyyy");
  dArsiv.setFrozenRows(1);

  var dAylik = ss.getSheetByName("DASHBOARD_AYLIK") || ss.insertSheet("DASHBOARD_AYLIK");
  dAylik.clear();
  var ayHeaders = ["Dönem / Ay", "Net Satılan (Gr)", "Alınan Tahsilat (Has)", "Müşteri İadesi (Gr)", "Aylık Net Kâr (Has Gr)"];
  dAylik.getRange(1, 1, 1, ayHeaders.length).setValues([ayHeaders]).setBackground("#27ae60").setFontColor("#ffffff").setFontWeight("bold")
        .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  dAylik.setFrozenRows(1);

  var sablonSheet = ss.getSheetByName("SABLON KOPYALA");
  if(sablonSheet) ss.deleteSheet(sablonSheet);
  
  islemlerSheet.setColumnWidth(1, 150); islemlerSheet.setColumnWidth(2, 110); islemlerSheet.setColumnWidth(3, 180); 
  islemlerSheet.setColumnWidth(4, 140); islemlerSheet.setColumnWidth(5, 120); islemlerSheet.setColumnWidth(6, 110); 
  islemlerSheet.setColumnWidth(7, 120); islemlerSheet.setColumnWidth(8, 100); islemlerSheet.setColumnWidth(9, 280); 

  girisSheet.setColumnWidth(1, 130);
  girisSheet.setColumnWidth(2, 110); girisSheet.setColumnWidth(3, 110); 
  girisSheet.setColumnWidth(4, 130); girisSheet.setColumnWidth(5, 70); girisSheet.setColumnWidth(6, 70); 
  girisSheet.setColumnWidth(7, 220);  girisSheet.setColumnWidth(8, 30); girisSheet.setColumnWidth(9, 260); 
  girisSheet.setColumnWidth(10, 130);
  
  mOzetSheet.setColumnWidth(1, 110); mOzetSheet.setColumnWidth(2, 140);
  mOzetSheet.setColumnWidth(3, 110); 
  mOzetSheet.setColumnWidth(4, 100); mOzetSheet.setColumnWidth(5, 110); mOzetSheet.setColumnWidth(6, 130);
  mOzetSheet.setColumnWidth(7, 90);  mOzetSheet.setColumnWidth(8, 250); mOzetSheet.setColumnWidth(9, 150);

  lorettaSheet.setColumnWidth(1, 140); lorettaSheet.setColumnWidth(2, 110); lorettaSheet.setColumnWidth(3, 160);
  lorettaSheet.setColumnWidth(4, 120); lorettaSheet.setColumnWidth(5, 110); lorettaSheet.setColumnWidth(6, 90);
  lorettaSheet.setColumnWidth(7, 110); lorettaSheet.setColumnWidth(8, 200);

  dMusteri.setColumnWidth(1, 180); dMusteri.setColumnWidth(2, 200); dMusteri.setColumnWidth(3, 180); 
  dMusteri.setColumnWidth(4, 190); dMusteri.setColumnWidth(5, 230); dMusteri.setColumnWidth(6, 140);
  
  dAtolye.setColumnWidth(1, 180); dAtolye.setColumnWidth(2, 200);
  dAtolye.setColumnWidth(3, 180); 
  dAtolye.setColumnWidth(4, 190); dAtolye.setColumnWidth(5, 230); dAtolye.setColumnWidth(6, 140);

  dArsiv.setColumnWidth(1, 190); dArsiv.setColumnWidth(2, 130); dArsiv.setColumnWidth(3, 190); 
  dArsiv.setColumnWidth(4, 180); dArsiv.setColumnWidth(5, 160); dArsiv.setColumnWidth(6, 140);
  dArsiv.setColumnWidth(7, 150); dArsiv.setColumnWidth(8, 140);
  
  dAylik.setColumnWidth(1, 130); dAylik.setColumnWidth(2, 170); dAylik.setColumnWidth(3, 180); 
  dAylik.setColumnWidth(4, 160); dAylik.setColumnWidth(5, 180);

  SpreadsheetApp.flush(); 
  guncelleCariDropdown();
  guncelleDashboards();
  hesaplaDonemselRapor();
  SpreadsheetApp.getUi().alert("💎 V68 Sistemi Güncellendi!\n\n- Rapor sayfasında çoklu Milyem seçimi ve Genel Toplam satırı aktif edildi.");
}

/**
 * 2. YALNIZCA AÇILIR LİSTE, İMLEÇ ROTASI VE RAPOR TETİKLEYİCİLERİ (HIZLANDIRILDI)
 */
function onEdit(e) {
  if (!e) return;
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var col = range.getColumn();
  if (sheet.getName() === "GIRIS_FORMU") {
    if (row >= 8 && row <= 17) {
      if (col === 1) sheet.getRange(row, 2).activate();
      else if (col === 2) sheet.getRange(row, 3).activate();
      else if (col === 3) sheet.getRange(row, 5).activate();
      else if (col === 5) sheet.getRange(row, 6).activate();
      else if (col === 6) sheet.getRange(row, 7).activate();
      else if (col === 7) { 
        if (row < 17) sheet.getRange(row + 1, 1).activate();
        else sheet.getRange("F19").activate();
      }
    }
    else if (row === 4 && col === 2 && range.getValue()) {
      sheet.getRange("C4").clearContent();
      sheet.getRange("A8:A17").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["ALIŞ", "ÖDEME", "ATÖLYE İADESİ", "SERMAYE GİRİŞİ"], true).build());
    } 
    else if (row === 4 && col === 3 && range.getValue()) {
      sheet.getRange("B4").clearContent();
      sheet.getRange("A8:A17").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["SATIŞ", "TAHSİLAT", "MÜŞTERİ İADESİ"], true).build());
    }
    else if (row === 4 && (col === 2 || col === 3) && !range.getValue()) {
      sheet.getRange("A8:A17").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["SATIŞ", "TAHSİLAT", "ALIŞ", "ÖDEME", "MÜŞTERİ İADESİ", "ATÖLYE İADESİ", "SERMAYE GİRİŞİ"], true).build());
    }
  }
  
  else if (sheet.getName() === "CARI_ARSIV" && col === 8 && row > 1 && (e.value === "TRUE" || e.value === "DOĞRU" || e.value === true)) {
    var cariAdi = sheet.getRange(row, 1).getDisplayValue().toString().trim().toUpperCase();
    if (cariAdi) {
      var ui = SpreadsheetApp.getUi();
      var onay = ui.alert("♻️ Arşivden Çıkar", cariAdi + " isimli hesabı tekrar aktif listelere almak istediğinize emin misiniz?", ui.ButtonSet.YES_NO);
      if (onay === ui.Button.YES) {
        var cKartlar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CARI_KARTLAR");
        if(cKartlar) {
          var cData = cKartlar.getDataRange().getValues();
          var islemSayisi = 0;
          for(var i=1; i<cData.length; i++){
            if(cData[i][1].toString().trim().toUpperCase() === cariAdi) {
              cKartlar.getRange(i+1, 4).setValue("AKTİF");
              islemSayisi++;
            }
          }
          if (islemSayisi === 0) {
            var cTip = sheet.getRange(row, 2).getDisplayValue().toString().trim().toUpperCase();
            cKartlar.appendRow([new Date(), cariAdi, cTip, "AKTİF", 0, ""]);
          }
        }
        SpreadsheetApp.flush();
        guncelleCariDropdown();
        guncelleDashboards();
        SpreadsheetApp.getActiveSpreadsheet().toast(cariAdi + " aktif olarak geri döndü.", "♻️ Arşivden Çıkarıldı", 3);
        return;
      }
      range.setValue(false); 
    }
  }

  else if (sheet.getName() === "MUSTERI_OZET") {
    if ((row === 4 && (col >= 2 && col <= 4)) || ((row === 3 || row === 4) && (col >= 7 && col <= 9))) {
      var secilenMusteri = sheet.getRange("B4").getValue().toString().trim().toUpperCase();
      if (secilenMusteri !== "") hesaplaMusteriOzeti(secilenMusteri);
      else {
        sheet.getRange("B6:D10").clearContent();
        sheet.getRange("H7:I8").clearContent();
        sheet.getRange("H10:I10").clearContent();
        var lr = sheet.getLastRow();
        if(lr >= 13) sheet.getRange("A13:I" + lr).clearContent();
      }
    }
  }
  
  else if (sheet.getName() === "RAPOR_DONEMSEL") {
    if ((row === 3 || row === 4) && (col === 2 || col >= 5)) {
      hesaplaDonemselRapor();
    }
  }
}

/**
 * 🟩 MAKRO BUTON FONKSİYONU 1: KAYDET
 */
function butonKaydiOnayla() {
  prosesCokluFis();
}

/**
 * 🟩 MAKRO BUTON FONKSİYONU 2: ARŞİVLE
 */
function butonArsiveTasi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GIRIS_FORMU");
  
  var musteriAdi = sheet.getRange("C4").getValue().toString().trim().toUpperCase();
  var atolyeAdi = sheet.getRange("B4").getValue().toString().trim().toUpperCase();
  var cariAdi = musteriAdi || atolyeAdi;
  var tip = musteriAdi ? "MÜŞTERİ" : (atolyeAdi ? "ATÖLYE" : "");
  if (!cariAdi) {
    SpreadsheetApp.getUi().alert("⚠️ Hata: Arşive kaldırmak için B4'ten Atölye veya C4'ten Müşteri seçmelisiniz.");
    return;
  }
  
  var ui = SpreadsheetApp.getUi();
  var onay = ui.alert("🗄️ " + tip + " Arşive Kaldırılacak", cariAdi + " isimli hesabı arşive taşımak istediğinize emin misiniz?", ui.ButtonSet.YES_NO);
  if (onay === ui.Button.YES) {
    var durumHucresi = sheet.getRange("E18:G18");
    durumHucresi.merge().setValue("⏳ TAŞINIYOR...").setBackground("#e67e22").setFontColor("#ffffff")
                .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");
    SpreadsheetApp.flush();

    var cKartlar = ss.getSheetByName("CARI_KARTLAR");
    var islemSayisi = 0;
    if(cKartlar) {
      var cData = cKartlar.getDataRange().getValues();
      for(var i=1; i<cData.length; i++){
        if(cData[i][1].toString().trim().toUpperCase() === cariAdi) {
          cKartlar.getRange(i+1, 4).setValue("ARŞİVLENDİ");
          islemSayisi++;
        }
      }
      if(islemSayisi === 0) {
        cKartlar.appendRow([new Date(), cariAdi, tip, "ARŞİVLENDİ", 0, ""]);
      }
    }
    
    if (musteriAdi) sheet.getRange("C4").clearContent(); 
    if (atolyeAdi) sheet.getRange("B4").clearContent(); 
    
    SpreadsheetApp.flush();
    guncelleCariDropdown();
    guncelleDashboards();
    ss.toast(cariAdi + " başarıyla arşive taşındı.", "🗄️ Arşivlendi", 3);

    durumHucresi.setValue("🗄️ ARŞİVLENDİ!").setBackground("#27ae60");
    SpreadsheetApp.flush();
    Utilities.sleep(1500);
    durumHucresi.clearContent().setBackground(null);
  }
}

/**
 * 🟩 MAKRO BUTON FONKSİYONU 3: TEMİZLE
 */
function butonFormuTemizle() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GIRIS_FORMU");
  
  sheet.getRange("A8:C17").clearContent();
  sheet.getRange("E8:E17").uncheck();
  sheet.getRange("F8:G17").clearContent();
  sheet.getRange("B4:C4").clearContent();
  sheet.getRange("B5").setValue("=TODAY()");
  
  var varsayilanRule = SpreadsheetApp.newDataValidation().requireValueInList(["SATIŞ", "TAHSİLAT", "ALIŞ", "ÖDEME", "MÜŞTERİ İADESİ", "ATÖLYE İADESİ", "SERMAYE GİRİŞİ"], true).build();
  sheet.getRange("A8:A17").setDataValidation(varsayilanRule);
  ss.toast("Giriş formu başarıyla sıfırlandı.", "🧹 Form Temizlendi", 2);

  var durumHucresi = sheet.getRange("E18:G18");
  durumHucresi.merge().setValue("🧹 TEMİZLENDİ!").setBackground("#34495e").setFontColor("#ffffff")
              .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");
  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  durumHucresi.clearContent().setBackground(null);
}

/**
 * 📊 B.I. MOTORU: KÂR İZOLASYONLU DÖNEMSEL PERFORMANS VE MİLYEM RAPORU (V68 GÜNCELLENDİ)
 */
function hesaplaDonemselRapor() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dSheet = ss.getSheetByName("RAPOR_DONEMSEL");
  if (!dSheet) return;

  var baslangicVal = dSheet.getRange("B3").getValue();
  var bitisVal = dSheet.getRange("B4").getValue();
  var tipFiltre = dSheet.getRange("E3").getValue().toString().trim().toUpperCase();
  
  // V68: Milyem filtresini çoklu (virgüllü) hale getiriyoruz
  var milyemFiltreRaw = dSheet.getRange("E4").getValue().toString().trim();
  var milyemFiltre = milyemFiltreRaw.toUpperCase();
  var secilenMilyemler = milyemFiltreRaw.split(',').map(function(item) { return item.trim(); });
  
  var baslangic = (baslangicVal && !isNaN(new Date(baslangicVal))) ? new Date(baslangicVal).setHours(0,0,0,0) : null;
  var bitis = (bitisVal && !isNaN(new Date(bitisVal))) ? new Date(bitisVal).setHours(23,59,59,999) : null;

  var islemlerSheet = ss.getSheetByName("ISLEMLER");
  var cKartlarSheet = ss.getSheetByName("CARI_KARTLAR");
  if (!islemlerSheet || !cKartlarSheet) return;
  var iData = islemlerSheet.getDataRange().getValues();
  var cData = cKartlarSheet.getDataRange().getValues();
  
  var cariTipleri = {};
  for(var c=1; c<cData.length; c++) {
    var ad = cData[c][1].toString().trim().toUpperCase();
    var cTip = cData[c][2].toString().trim().toUpperCase();
    if(ad) cariTipleri[ad] = cTip;
  }

  var ayarlar = getAyarlar();
  var masterData = {};
  for (var i = 1; i < iData.length; i++) {
    var tarih = iData[i][1];
    var cari = iData[i][2].toString().trim().toUpperCase();
    var tip = iData[i][3].toString().trim().toUpperCase();
    var gram = parseFloat(iData[i][4]) || 0;
    var islemMilyemi = iData[i][5].toString().trim();
    var has = parseFloat(iData[i][6]) || 0;

    if (!cari) continue;

    var cTipObj = cariTipleri[cari] || "BİLİNMİYOR";
    if (tipFiltre !== "TÜMÜ" && cTipObj !== tipFiltre) continue;

    var tTime = new Date(tarih).getTime();
    if (baslangic && tTime < baslangic) continue;
    if (bitis && tTime > bitis) continue;
    
    // V68: Milyem eşleştirme (Çoklu arama)
    if (milyemFiltre !== "TÜMÜ" && milyemFiltre !== "") {
      var isMatch = false;
      for(var z = 0; z < secilenMilyemler.length; z++) {
        if(islemMilyemi === secilenMilyemler[z]) { 
          isMatch = true; 
          break; 
        }
      }
      if(!isMatch) continue;
    }

    if (!masterData[cari]) {
      masterData[cari] = {
        satisGram: 0,
        iadeGram: 0,
        safSatisGram: 0,
        safSatisHas: 0,
        tahsilatHas: 0,
        karHas: 0
      };
    }

    if (cTipObj === "MÜŞTERİ") {
      if (tip === "SATIŞ" || tip === "DEVİR (MÜŞTERİ)") masterData[cari].satisGram += gram;
      if (tip === "MÜŞTERİ İADESİ") masterData[cari].iadeGram += gram;
      if (tip === "TAHSİLAT") masterData[cari].tahsilatHas += has;
      if (tip === "SATIŞ") {
        masterData[cari].safSatisGram += gram;
        masterData[cari].safSatisHas += has;
      } else if (tip === "MÜŞTERİ İADESİ") {
        masterData[cari].safSatisGram -= gram;
        masterData[cari].safSatisHas -= has;
      }
    } else { 
      if (tip === "ALIŞ" || tip === "DEVİR (ATÖLYE)" || tip === "SERMAYE GİRİŞİ") masterData[cari].satisGram += gram;
      if (tip === "ATÖLYE İADESİ") masterData[cari].iadeGram += gram;
      if (tip === "ÖDEME") masterData[cari].tahsilatHas += has;
      if (tip === "ALIŞ") {
        masterData[cari].safSatisGram += gram;
        masterData[cari].safSatisHas += has;
      } else if (tip === "ATÖLYE İADESİ") {
        masterData[cari].safSatisGram -= gram;
        masterData[cari].safSatisHas -= has;
      }
    }
  }

  var matris = [];
  var cariler = Object.keys(masterData);
  for (var x = 0; x < cariler.length; x++) {
    var d = masterData[cariler[x]];
    var netIslem = d.satisGram - d.iadeGram;
    
    if (d.satisGram === 0 && d.iadeGram === 0 && d.tahsilatHas === 0 && d.safSatisHas === 0) continue;
    var ortMilyem = d.safSatisGram > 0 ? (d.safSatisHas * 1000 / d.safSatisGram) : 0;
    var netKar = d.safSatisHas - (d.safSatisGram * ayarlar.varsayilanMilyem / 1000);
    matris.push([
      cariler[x],
      d.satisGram,
      d.iadeGram,
      netIslem,
      ortMilyem,
      d.tahsilatHas,
      netKar
    ]);
  }

  matris.sort(function(a, b) {
    if (b[3] !== a[3]) return b[3] - a[3]; 
    return b[6] - a[6]; 
  });
  
  var lr = dSheet.getLastRow();
  if (lr >= 7) {
    dSheet.getRange("A7:G" + lr).clearContent();
    // Eski "Genel Toplam" satırı formatlarını silmek için
    dSheet.getRange("A7:G" + lr).setBackground(null).setFontColor(null).setFontWeight("normal");
  }
  
  if (matris.length > 0) {
    // V68: Alt Kısım "Genel Toplam" Satırı Hesaplaması
    var tSatis = 0, tIade = 0, tNet = 0, tTahsilat = 0, tKar = 0;
    for(var r=0; r<matris.length; r++) {
      tSatis += parseFloat(matris[r][1]) || 0;
      tIade += parseFloat(matris[r][2]) || 0;
      tNet += parseFloat(matris[r][3]) || 0;
      tTahsilat += parseFloat(matris[r][5]) || 0;
      tKar += parseFloat(matris[r][6]) || 0;
    }
    
    // Matrise Toplam satırını ekle
    matris.push(["GENEL TOPLAM", tSatis, tIade, tNet, "---", tTahsilat, tKar]);

    dSheet.getRange(7, 1, matris.length, 7).setValues(matris)
          .setVerticalAlignment("middle")
          .setWrap(true);
          
    // V68: Toplam satırını belirginleştir (Koyu Lacivert / Sarı)
    dSheet.getRange(6 + matris.length, 1, 1, 7).setBackground("#2c3e50").setFontColor("#f1c40f").setFontWeight("bold");
  }
}

/**
 * 📊 MÜŞTERİ ÖZET MOTORU 
 */
function hesaplaMusteriOzeti(musteriAdi) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mOzetSheet = ss.getSheetByName("MUSTERI_OZET");
  var islemlerSheet = ss.getSheetByName("ISLEMLER");
  var ayarlar = getAyarlar();
  
  var tumVeriler = islemlerSheet.getDataRange().getValues();
  
  var baslangicVal = mOzetSheet.getRange("H3").getValue(); 
  var bitisVal = mOzetSheet.getRange("H4").getValue();
  var baslangic = (baslangicVal && !isNaN(new Date(baslangicVal))) ? new Date(baslangicVal).setHours(0,0,0,0) : null;
  var bitis = (bitisVal && !isNaN(new Date(bitisVal))) ? new Date(bitisVal).setHours(23,59,59,999) : null;

  var allHasSatis = 0, allHasIade = 0, allHasTahsilat = 0;
  var allAdetSatis = 0, allAdetIade = 0, allAdetTahsilat = 0;

  var pGramSatis = 0, pHasSatis = 0, pAdetSatis = 0;
  var pGramIade = 0, pHasIade = 0, pAdetIade = 0;
  var pGramTahsilat = 0, pHasTahsilat = 0, pAdetTahsilat = 0;
  var pSafGramSatisIade = 0, pSafHasSatisIade = 0;
  
  var sonIslemTarihi = null;
  var islemGecmisi = [];
  var islemRenkleriListesi = [];
  var islemKalinlikListesi = [];
  
  var yuruyenBakiye = 0;
  
  for (var i = 1; i < tumVeriler.length; i++) {
    if (tumVeriler[i][2] === musteriAdi) {
      var id = tumVeriler[i][0];
      var tarih = tumVeriler[i][1];
      var tip = tumVeriler[i][3];
      var gram = parseFloat(tumVeriler[i][4]) || 0;
      var milyem = parseFloat(tumVeriler[i][5]) || 0;
      var has = parseFloat(tumVeriler[i][6]) || 0;
      var adet = parseInt(tumVeriler[i][7]) || 0;
      var aciklama = tumVeriler[i][8];
      var tDate = new Date(tarih);
      var tTime = tDate.getTime();

      var islemSign = 0;
      if (tip === "SATIŞ" || tip === "DEVİR (MÜŞTERİ)" || tip === "ALIŞ" || tip === "DEVİR (ATÖLYE)" || tip === "SERMAYE GİRİŞİ") {
        allHasSatis += has;
        allAdetSatis += adet;
        islemSign = 1; 
      } else if (tip === "MÜŞTERİ İADESİ" || tip === "ATÖLYE İADESİ") {
        allHasIade += has;
        allAdetIade += adet;
        islemSign = -1; 
      } else if (tip === "TAHSİLAT" || tip === "ÖDEME") {
        allHasTahsilat += has;
        allAdetTahsilat += adet;
        islemSign = -1; 
      }

      yuruyenBakiye += (has * islemSign);
      var inRange = true;
      if (baslangic && tTime < baslangic) inRange = false;
      if (bitis && tTime > bitis) inRange = false;

      if (inRange) {
        if (tip === "SATIŞ" || tip === "DEVİR (MÜŞTERİ)" || tip === "ALIŞ" || tip === "DEVİR (ATÖLYE)" || tip === "SERMAYE GİRİŞİ") {
          pGramSatis += gram;
          pHasSatis += has; pAdetSatis += adet;
          if ((tip === "SATIŞ" || tip === "ALIŞ") && (!sonIslemTarihi || tDate > sonIslemTarihi)) sonIslemTarihi = tDate;
        } else if (tip === "MÜŞTERİ İADESİ" || tip === "ATÖLYE İADESİ") {
          pGramIade += gram;
          pHasIade += has; pAdetIade += adet;
        } else if (tip === "TAHSİLAT" || tip === "ÖDEME") {
          pGramTahsilat += gram;
          pHasTahsilat += has; pAdetTahsilat += adet;
        }
        
        if (tip === "SATIŞ" || tip === "ALIŞ") {
          pSafGramSatisIade += gram;
          pSafHasSatisIade += has;
        } else if (tip === "MÜŞTERİ İADESİ" || tip === "ATÖLYE İADESİ") {
          pSafGramSatisIade -= gram;
          pSafHasSatisIade -= has;
        }

        var displayTarih = (tarih instanceof Date) ? tarih : new Date(tarih);
        
        islemGecmisi.push([
          displayTarih, 
          tip, 
          gram > 0 ? gram : "", 
          milyem > 0 ? milyem : "", 
          has, 
          yuruyenBakiye, 
          adet > 0 ? adet : "", 
          aciklama, 
          id
        ]);
        var c = ayarlar.islemRenkleri[tip] || "#000000";
        islemRenkleriListesi.push([c, c, c, c, c, c, c, c, c]);
        islemKalinlikListesi.push(["bold", "bold", "bold", "bold", "bold", "bold", "bold", "bold", "bold"]);
      }
    }
  }
  
  var allHasBakiye = allHasSatis - allHasIade - allHasTahsilat;
  var allAdetBakiye = allAdetSatis - allAdetIade - allAdetTahsilat;
  var pHasNet = pHasSatis - pHasIade - pHasTahsilat;
  var pAdetNet = pAdetSatis - pAdetIade - pAdetTahsilat;
  
  var lr = mOzetSheet.getLastRow();
  if (lr >= 13) { mOzetSheet.getRange("A13:I" + lr).clearContent();
  }
  
  var kpiMatrix = [
    [pGramSatis, pHasSatis, pAdetSatis],
    [pGramIade, pHasIade, pAdetIade],
    [pGramTahsilat, pHasTahsilat, pAdetTahsilat],
    ["---", pHasNet, pAdetNet],         
    ["---", allHasBakiye, allAdetBakiye] 
  ];
  mOzetSheet.getRange(6, 2, 5, 3).setValues(kpiMatrix);
  
  var dColor = (pHasNet > 0) ? "#27ae60" : (pHasNet < 0 ? "#c0392b" : "#2c3e50");
  mOzetSheet.getRange("B9:D9").setFontColor(dColor);
  var aColor = (allHasBakiye > 0) ? "#27ae60" : (allHasBakiye < 0 ? "#c0392b" : "#2c3e50");
  mOzetSheet.getRange("B10:D10").setFontColor(aColor);
  var ortMilyem = pSafGramSatisIade > 0 ? (pSafHasSatisIade * 1000 / pSafGramSatisIade) : 0;
  var netKar = pSafHasSatisIade - (pSafGramSatisIade * ayarlar.varsayilanMilyem / 1000);
  
  mOzetSheet.getRange("H7").setValue(ortMilyem);
  mOzetSheet.getRange("H8").setValue(netKar);
  
  var strTarih = sonIslemTarihi ? sonIslemTarihi : "İşlem Yok";
  mOzetSheet.getRange("H10").setValue(strTarih);
  
  if (islemGecmisi.length > 0) {
    var range = mOzetSheet.getRange(13, 1, islemGecmisi.length, 9);
    range.setValues(islemGecmisi);
    range.setFontColors(islemRenkleriListesi);
    range.setFontWeights(islemKalinlikListesi);
    range.setVerticalAlignment("middle").setWrap(true);
  }
}

/**
 * 🛠️ KAYIT MOTORU 
 */
function prosesCokluFis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ayarlar = getAyarlar(); 
  var girisSheet = ss.getSheetByName("GIRIS_FORMU");
  
  var atolye = girisSheet.getRange("B4").getValue().toString().trim().toUpperCase();
  var musteri = girisSheet.getRange("C4").getValue().toString().trim().toUpperCase();
  var cariAdi = atolye || musteri; 
  
  var rawDate = girisSheet.getRange("B5").getValue();
  var islemTarihi = new Date(rawDate);
  
  var formVerileri = girisSheet.getRange("A8:G17").getValues(); 

  if (!cariAdi) {
    SpreadsheetApp.getUi().alert("⚠️ Hata: Lütfen geçerli bir Cari Seçimi yapın!");
    return;
  }
  
  for (var k = 0; k < formVerileri.length; k++) {
    var checkTip = formVerileri[k][0];
    var checkGram = parseFloat(formVerileri[k][1].toString().replace(',', '.'));
    
    if (checkTip === "ALIŞ" || checkTip === "ATÖLYE İADESİ" || checkTip === "SERMAYE GİRİŞİ") {
      if (isNaN(checkGram) || checkGram <= 0) {
        SpreadsheetApp.getUi().alert("⚠️ Kayıt İşlemi İptal Edildi!\n\nSeçtiğiniz satırlarda sisteme ürün giren bir komut mevcut ancak Gram hatalı.");
        return;
      }
    }
  }

  var islemlerSheet = ss.getSheetByName("ISLEMLER");
  var cKartlarSheet = ss.getSheetByName("CARI_KARTLAR");
  var lorettaSheet = ss.getSheetByName("LORETTA_BORC");
  var islemTipiDevir = (atolye !== "") ? "DEVİR (ATÖLYE)" : "DEVİR (MÜŞTERİ)";
  var cType = (atolye !== "") ? "ATÖLYE" : "MÜŞTERİ";
  
  var cariVarMi = false;
  if(cKartlarSheet) {
    var cData = cKartlarSheet.getDataRange().getValues();
    for(var c=1; c<cData.length; c++){
      if(cData[c][1] === cariAdi) { cariVarMi = true; break; }
    }
  }

  if (!cariVarMi && cKartlarSheet) {
    var promptTitle = (atolye !== "") ? "🏭 Yeni Atölye Kaydı" : "👥 Yeni Müşteri Kaydı";
    var msgHas = "Sistemde " + cariAdi + " adında bir cari bulunamadı. Yeni cari kartı açılıyor.\n\nBu cari için geçmişten devreden BAŞLANGIÇ BAKİYESİ (Has Altın) var mı?\n(Yoksa boş bırakın.)";
    var respHas = SpreadsheetApp.getUi().prompt(promptTitle, msgHas, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
    if (respHas.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return;
    
    var devirHas = 0;
    var hasInput = respHas.getResponseText().trim().replace(',', '.');
    if (hasInput !== "" && !isNaN(parseFloat(hasInput))) devirHas = parseFloat(hasInput);

    var devirAdet = "";
    if (atolye !== "") {
      var msgAdet = "Devreden KALAN YÜRÜYEN ADET var mı?\n(Yoksa boş bırakın.)";
      var respAdet = SpreadsheetApp.getUi().prompt(promptTitle, msgAdet, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
      if (respAdet.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return;
      var adetInput = respAdet.getResponseText().trim();
      if (adetInput !== "" && !isNaN(parseInt(adetInput))) devirAdet = parseInt(adetInput);
    }

    cKartlarSheet.appendRow([islemTarihi, cariAdi, cType, "AKTİF", devirHas, devirAdet]);
    cKartlarSheet.getRange(cKartlarSheet.getLastRow(), 1, 1, 6).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);

    if (devirHas !== 0 || devirAdet !== "") {
      var uniqueID_devir = "TX-" + new Date().getTime() + "-DEVIR";
      islemlerSheet.appendRow([uniqueID_devir, islemTarihi, cariAdi, islemTipiDevir, "", "", devirHas, devirAdet, "SİSTEME AÇILIŞ BAKİYESİ"]);
      var yaziRengiDevir = ayarlar.islemRenkleri[islemTipiDevir] || "#000000";
      islemlerSheet.getRange(islemlerSheet.getLastRow(), 1, 1, 9).setFontColor(yaziRengiDevir).setFontWeight("bold")
                   .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    }
  }
  
  var durumHucresi = girisSheet.getRange("E18:G18");
  durumHucresi.merge().setValue("⏳ İŞLENİYOR...").setBackground("#c0392b").setFontColor("#ffffff")
              .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");
  SpreadsheetApp.flush(); 

  var veriEklendiMi = false;
  var timestamp = "TX-" + new Date().getTime();
  
  for (var i = 0; i < formVerileri.length; i++) {
    var islemTipi = formVerileri[i][0] ? formVerileri[i][0].toString().toUpperCase() : "";
    var gram14K = parseFloat(formVerileri[i][1].toString().replace(',', '.')) || 0; 
    var milyem = parseFloat(formVerileri[i][2]) || ayarlar.varsayilanMilyem;
    var formHas = parseFloat(formVerileri[i][3].toString().replace(',', '.')) || 0;
    
    var lorettaCheck = formVerileri[i][4] === true; 
    var adet = formVerileri[i][5] !== "" ? parseInt(formVerileri[i][5]) : ""; 
    var aciklama = formVerileri[i][6] ? formVerileri[i][6].toString().toUpperCase() : ""; 
    
    if (islemTipi && (gram14K > 0 || formHas > 0 || (adet !== "" && adet > 0))) {
      var hasAltin = 0;
      if (formHas > 0) {
        hasAltin = formHas;
      } else if (gram14K > 0) {
        hasAltin = Math.round((gram14K * (milyem / 1000)) * 100) / 100;
      }
      
      var uniqueID = timestamp + "-" + (i + 1);
      islemlerSheet.appendRow([uniqueID, islemTarihi, cariAdi, islemTipi, gram14K > 0 ? gram14K : "", milyem, hasAltin, adet, aciklama]);
      
      var yaziRengi = ayarlar.islemRenkleri[islemTipi] || "#000000";
      islemlerSheet.getRange(islemlerSheet.getLastRow(), 1, 1, 9).setFontColor(yaziRengi).setFontWeight("bold")
                   .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
      
      if (lorettaCheck && lorettaSheet && (islemTipi === "SATIŞ" || islemTipi === "MÜŞTERİ İADESİ")) {
        lorettaSheet.appendRow([uniqueID, islemTarihi, cariAdi, islemTipi, gram14K > 0 ? gram14K : "", milyem, hasAltin, aciklama]);
      }
      
      veriEklendiMi = true;
    }
  }
  
  if (veriEklendiMi || !cariVarMi) { 
    ss.toast("Girdiğiniz tüm işlemler veritabanına işlendi.", "✅ Kayıt Tamamlandı", 3);

    var lrIslemler = islemlerSheet.getLastRow();
    if (lrIslemler > 1) {
      islemlerSheet.getRange(2, 1, lrIslemler - 1, 9).sort({column: 2, ascending: true});
    }

    girisSheet.getRange("A8:C17").clearContent(); 
    girisSheet.getRange("E8:E17").uncheck(); 
    girisSheet.getRange("F8:G17").clearContent(); 
    girisSheet.getRange("D8:D17").setFormula('=IF(AND(B8>0; C8>0); ROUND(B8*C8/1000; 2); "")');
    girisSheet.getRange("B4:C4").clearContent(); 
    girisSheet.getRange("B5").setValue("=TODAY()");
    
    var varsayilanRule = SpreadsheetApp.newDataValidation().requireValueInList(["SATIŞ", "TAHSİLAT", "ALIŞ", "ÖDEME", "MÜŞTERİ İADESİ", "ATÖLYE İADESİ", "SERMAYE GİRİŞİ"], true).build();
    girisSheet.getRange("A8:A17").setDataValidation(varsayilanRule);
    
    SpreadsheetApp.flush(); 
    guncelleCariDropdown(); 
    guncelleDashboards();
    hesaplaDonemselRapor();
    
    var ozetSheet = ss.getSheetByName("MUSTERI_OZET");
    if(ozetSheet) {
      var aktifMusteri = ozetSheet.getRange("B4").getValue();
      if(aktifMusteri) hesaplaMusteriOzeti(aktifMusteri);
    }

    durumHucresi.setValue("✅ BAŞARILI!").setBackground("#27ae60");
    SpreadsheetApp.flush();
    Utilities.sleep(1500); 
    durumHucresi.clearContent().setBackground(null);
  } else {
    durumHucresi.clearContent().setBackground(null);
  }
}

/**
 * 4. DASHBOARD GÜNCELLEME MOTORU
 */
function guncelleDashboards() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ayarlar = getAyarlar();
  var mDef = ayarlar.varsayilanMilyem; 
  
  var cKartlar = ss.getSheetByName("CARI_KARTLAR");
  if(!cKartlar) return;
  var cData = cKartlar.getDataRange().getValues();
  
  var cariDurum = {};
  for(var c=1; c<cData.length; c++){
    var ad = cData[c][1].toString().trim().toUpperCase();
    var tip = cData[c][2].toString().trim().toUpperCase();
    var durum = cData[c][3].toString().trim().toUpperCase();
    var safeAd = ad.replace(/"/g, '""');
    
    if(ad) {
      if(!cariDurum[ad]) {
        cariDurum[ad] = { tip: tip, safeAd: safeAd, isArchived: (durum === "ARŞİVLENDİ") };
      } else {
        if(durum === "ARŞİVLENDİ") cariDurum[ad].isArchived = true;
      }
    }
  }

  var activeMKeys = [];
  var activeAKeys = [];
  var allArchived = []; 

  for (var key in cariDurum) {
    if (cariDurum[key].isArchived) {
      allArchived.push({ type: cariDurum[key].tip, ad: key, safeAd: cariDurum[key].safeAd });
    } else {
      if (cariDurum[key].tip === "MÜŞTERİ") activeMKeys.push({ ad: key, safeAd: cariDurum[key].safeAd });
      else activeAKeys.push({ ad: key, safeAd: cariDurum[key].safeAd });
    }
  }

  activeMKeys.sort(function(a, b) { return a.ad.localeCompare(b.ad, 'tr'); });
  activeAKeys.sort(function(a, b) { return a.ad.localeCompare(b.ad, 'tr'); });

  var dMusteri = ss.getSheetByName("DASHBOARD_MUSTERI");
  if (dMusteri && dMusteri.getLastRow() > 1) { dMusteri.getRange("A2:F").clearContent(); }
  
  if (activeMKeys.length > 0 && dMusteri) {
    var mData = [];
    for (var i = 0; i < activeMKeys.length; i++) {
      var r = i + 2;
      var k = activeMKeys[i].safeAd;
      mData.push([
        activeMKeys[i].ad,
        '=SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + k + '"; ISLEMLER!D:D; "SATIŞ") + SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + k + '"; ISLEMLER!D:D; "DEVİR (MÜŞTERİ)") - SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + k + '"; ISLEMLER!D:D; "TAHSİLAT") - SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + k + '"; ISLEMLER!D:D; "MÜŞTERİ İADESİ")',
        '=SUMIFS(ISLEMLER!E:E; ISLEMLER!C:C; "' + k + '"; ISLEMLER!D:D; "SATIŞ")',
        '=IFERROR(SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + k + '"; ISLEMLER!D:D; "SATIŞ") * 1000 / C' + r + '; 0)',
        '=IF(D' + r + '>0; D' + r + '-' + mDef + '; 0)',
        '=MAXIFS(ISLEMLER!B:B; ISLEMLER!C:C; "' + k + '")'
      ]);
    }
    dMusteri.getRange(2, 1, mData.length, 6).setValues(mData).setVerticalAlignment("middle").setWrap(true).setHorizontalAlignment("center");
    dMusteri.getRange(2, 2, mData.length, 2).setNumberFormat("#,##0.00");
    dMusteri.getRange(2, 4, mData.length, 2).setNumberFormat("0");
    dMusteri.getRange(2, 6, mData.length, 1).setNumberFormat("dd.MM.yyyy");
  }

  var dAtolye = ss.getSheetByName("DASHBOARD_ATOLYE");
  if (dAtolye && dAtolye.getLastRow() > 1) { dAtolye.getRange("A2:F").clearContent(); }
  
  if (activeAKeys.length > 0 && dAtolye) {
    var aData = [];
    for (var at = 0; at < activeAKeys.length; at++) {
      var rA = at + 2;
      var kA = activeAKeys[at].safeAd;
      aData.push([
        activeAKeys[at].ad,
        '=SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + kA + '"; ISLEMLER!D:D; "ALIŞ") + SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + kA + '"; ISLEMLER!D:D; "DEVİR (ATÖLYE)") + SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + kA + '"; ISLEMLER!D:D; "SERMAYE GİRİŞİ") - SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + kA + '"; ISLEMLER!D:D; "ÖDEME") - SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + kA + '"; ISLEMLER!D:D; "ATÖLYE İADESİ")',
        '=SUMIFS(ISLEMLER!E:E; ISLEMLER!C:C; "' + kA + '"; ISLEMLER!D:D; "ALIŞ") + SUMIFS(ISLEMLER!E:E; ISLEMLER!C:C; "' + kA + '"; ISLEMLER!D:D; "SERMAYE GİRİŞİ")',
        '=IFERROR((SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + kA + '"; ISLEMLER!D:D; "ALIŞ") + SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + kA + '"; ISLEMLER!D:D; "SERMAYE GİRİŞİ")) * 1000 / C' + rA + '; 0)',
        '=IF(D' + rA + '>0; D' + rA + '-' + mDef + '; 0)',
        '=MAXIFS(ISLEMLER!B:B; ISLEMLER!C:C; "' + kA + '")'
      ]);
    }
    dAtolye.getRange(2, 1, aData.length, 6).setValues(aData).setVerticalAlignment("middle").setWrap(true).setHorizontalAlignment("center");
    dAtolye.getRange(2, 2, aData.length, 2).setNumberFormat("#,##0.00");
    dAtolye.getRange(2, 4, aData.length, 2).setNumberFormat("0");
    dAtolye.getRange(2, 6, aData.length, 1).setNumberFormat("dd.MM.yyyy");
  }

  var dArsiv = ss.getSheetByName("CARI_ARSIV");
  if (dArsiv && dArsiv.getLastRow() > 1) { dArsiv.getRange("A2:H").clearContent(); dArsiv.getRange("H2:H").removeCheckboxes(); }
  
  if (allArchived.length > 0 && dArsiv) {
    allArchived.sort(function(a, b) { return a.ad.localeCompare(b.ad, 'tr'); });
    var arData = [];
    for (var x = 0; x < allArchived.length; x++) {
      var rArsiv = x + 2;
      var cx = allArchived[x].safeAd;
      var cAdNorm = allArchived[x].ad;
      
      if (allArchived[x].type === "MÜŞTERİ") {
        arData.push([
          cAdNorm,
          "MÜŞTERİ",
          '=SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "SATIŞ") + SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "DEVİR (MÜŞTERİ)") - SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "TAHSİLAT") - SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "MÜŞTERİ İADESİ")',
          '=SUMIFS(ISLEMLER!E:E; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "SATIŞ")',
          '=IFERROR(SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "SATIŞ") * 1000 / D' + rArsiv + '; 0)',
          '=IF(E' + rArsiv + '>0; E' + rArsiv + '-' + mDef + '; 0)',
          '=MAXIFS(ISLEMLER!B:B; ISLEMLER!C:C; "' + cx + '")'
        ]);
      } else {
        arData.push([
          cAdNorm,
          "ATÖLYE",
          '=SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "ALIŞ") + SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "DEVİR (ATÖLYE)") + SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "SERMAYE GİRİŞİ") - SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "ÖDEME") - SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "ATÖLYE İADESİ")',
          '=SUMIFS(ISLEMLER!E:E; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "ALIŞ") + SUMIFS(ISLEMLER!E:E; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "SERMAYE GİRİŞİ")',
          '=IFERROR((SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "ALIŞ") + SUMIFS(ISLEMLER!G:G; ISLEMLER!C:C; "' + cx + '"; ISLEMLER!D:D; "SERMAYE GİRİŞİ")) * 1000 / D' + rArsiv + '; 0)',
          '=IF(E' + rArsiv + '>0; E' + rArsiv + '-' + mDef + '; 0)',
          '=MAXIFS(ISLEMLER!B:B; ISLEMLER!C:C; "' + cx + '")'
        ]);
      }
    }
    
    dArsiv.getRange(2, 1, arData.length, 7).setValues(arData).setVerticalAlignment("middle").setWrap(true).setHorizontalAlignment("center");
    dArsiv.getRange(2, 3, arData.length, 2).setNumberFormat("#,##0.00");
    dArsiv.getRange(2, 5, arData.length, 2).setNumberFormat("0");
    dArsiv.getRange(2, 7, arData.length, 1).setNumberFormat("dd.MM.yyyy");
    dArsiv.getRange(2, 8, arData.length, 1).insertCheckboxes().setValue(false).setVerticalAlignment("middle").setHorizontalAlignment("center");
  }

  var dAylik = ss.getSheetByName("DASHBOARD_AYLIK");
  if (dAylik && dAylik.getLastRow() > 1) { dAylik.getRange("A2:E").clearContent(); }
  
  var aylar = getAylar();
  var ayKeys = Object.keys(aylar).sort();
  if (ayKeys.length > 0 && dAylik) {
    var ayData = [];
    for (var t = 0; t < ayKeys.length; t++) {
      var rAylik = t + 2;
      var parts = ayKeys[t].split("-");
      var tarihObj2 = new Date(parts[0], parseInt(parts[1], 10) - 1, 1);
      ayData.push([
        tarihObj2,
        '=SUMIFS(ISLEMLER!E:E; ISLEMLER!D:D; "SATIŞ"; ISLEMLER!B:B; ">="&A' + rAylik + '; ISLEMLER!B:B; "<="&EOMONTH(A' + rAylik + '; 0)) - SUMIFS(ISLEMLER!E:E; ISLEMLER!D:D; "MÜŞTERİ İADESİ"; ISLEMLER!B:B; ">="&A' + rAylik + '; ISLEMLER!B:B; "<="&EOMONTH(A' + rAylik + '; 0))',
        '=SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "TAHSİLAT"; ISLEMLER!B:B; ">="&A' + rAylik + '; ISLEMLER!B:B; "<="&EOMONTH(A' + rAylik + '; 0))',
        '=SUMIFS(ISLEMLER!E:E; ISLEMLER!D:D; "MÜŞTERİ İADESİ"; ISLEMLER!B:B; ">="&A' + rAylik + '; ISLEMLER!B:B; "<="&EOMONTH(A' + rAylik + '; 0))',
        '=(SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "SATIŞ"; ISLEMLER!B:B; ">="&A' + rAylik + '; ISLEMLER!B:B; "<="&EOMONTH(A' + rAylik + '; 0)) - SUMIFS(ISLEMLER!G:G; ISLEMLER!D:D; "MÜŞTERİ İADESİ"; ISLEMLER!B:B; ">="&A' + rAylik + '; ISLEMLER!B:B; "<="&EOMONTH(A' + rAylik + '; 0))) - (B' + rAylik + ' * GIRIS_FORMU!$J$5 / 1000)'
      ]);
    }
    dAylik.getRange(2, 1, ayData.length, 5).setValues(ayData).setVerticalAlignment("middle").setWrap(true).setHorizontalAlignment("center");
    dAylik.getRange(2, 1, ayData.length, 1).setNumberFormat("MM.yyyy");
    dAylik.getRange(2, 2, ayData.length, 4).setNumberFormat("#,##0.00");
  }
}

function guncelleCariDropdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var girisSheet = ss.getSheetByName("GIRIS_FORMU");
  if (!girisSheet) return;

  var islemlerSheet = ss.getSheetByName("ISLEMLER");
  var cKartlar = ss.getSheetByName("CARI_KARTLAR");
  if(!cKartlar || !islemlerSheet) return;

  var iData = islemlerSheet.getDataRange().getValues();
  var cariPuanlar = {};
  var benzersizMilyemler = {};
  for(var i=1; i<iData.length; i++) {
    var cari = iData[i][2].toString().trim().toUpperCase();
    var tip = iData[i][3].toString().trim().toUpperCase();
    var mly = iData[i][5].toString().trim();
    if(mly !== "") benzersizMilyemler[mly] = true;
    
    if(!cari) continue;

    if(!cariPuanlar[cari]) cariPuanlar[cari] = { m: 0, a: 0 };
    if (tip === "SATIŞ" || tip === "DEVİR (MÜŞTERİ)") cariPuanlar[cari].m += 100;
    if (tip === "TAHSİLAT" || tip === "MÜŞTERİ İADESİ") cariPuanlar[cari].m += 10;
    if (tip === "ALIŞ" || tip === "DEVİR (ATÖLYE)" || tip === "SERMAYE GİRİŞİ") cariPuanlar[cari].a += 100;
    if (tip === "ÖDEME" || tip === "ATÖLYE İADESİ") cariPuanlar[cari].a += 10;
  }

  var cData = cKartlar.getDataRange().getValues();
  var cariDurum = {};
  for(var c=1; c<cData.length; c++){
    var ad = cData[c][1].toString().trim().toUpperCase();
    var durum = cData[c][3].toString().trim().toUpperCase();
    if(ad) cariDurum[ad] = durum;
  }

  var mKeysActive = [];
  var aKeysActive = [];
  var tumCariler = Object.keys(cariPuanlar);
  for(var k=0; k<tumCariler.length; k++) {
    var ad = tumCariler[k];
    var durum = cariDurum[ad] || "AKTİF";
    if(!cariDurum[ad]) {
       var cTip = (cariPuanlar[ad].a > cariPuanlar[ad].m) ? "ATÖLYE" : "MÜŞTERİ";
       var bugun = new Date();
       cKartlar.appendRow([bugun, ad, cTip, "AKTİF", 0, ""]);
    }

    if(durum !== "ARŞİVLENDİ") {
       var finalTip = (cariPuanlar[ad].a > cariPuanlar[ad].m) ? "ATÖLYE" : "MÜŞTERİ";
       if(finalTip === "MÜŞTERİ") mKeysActive.push(ad);
       else aKeysActive.push(ad);
    }
  }

  mKeysActive.sort(function(a, b) { return a.localeCompare(b, 'tr'); });
  aKeysActive.sort(function(a, b) { return a.localeCompare(b, 'tr'); });

  var atolyeCell = girisSheet.getRange("B4");
  var musteriCell = girisSheet.getRange("C4");
  if (aKeysActive.length > 0) { 
    atolyeCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(aKeysActive, true).build()); 
  } else { 
    atolyeCell.clearDataValidations();
  }

  if (mKeysActive.length > 0) { 
    var musteriListRule = SpreadsheetApp.newDataValidation().requireValueInList(mKeysActive, true).build();
    musteriCell.setDataValidation(musteriListRule);
  } else { 
    musteriCell.clearDataValidations(); 
  }

  var mOzetSheet = ss.getSheetByName("MUSTERI_OZET");
  if (mOzetSheet) {
    var tumCarilerActive = mKeysActive.concat(aKeysActive);
    tumCarilerActive.sort(function(a, b) { return a.localeCompare(b, 'tr'); });
    if (tumCarilerActive.length > 0) {
      mOzetSheet.getRange("B4").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(tumCarilerActive, true).build());
    } else {
      mOzetSheet.getRange("B4").clearDataValidations();
    }
  }
  
  var dRaporSheet = ss.getSheetByName("RAPOR_DONEMSEL");
  if (dRaporSheet) {
    var milyemListesi = Object.keys(benzersizMilyemler);
    milyemListesi.sort(); 
    milyemListesi.unshift("TÜMÜ"); 
    // V68: Milyem veri doğrulamasını "Uyarı ver ama metin girmesine izin ver" (setAllowInvalid(true)) şeklinde esnettik!
    var mlyRule = SpreadsheetApp.newDataValidation().requireValueInList(milyemListesi, true).setAllowInvalid(true).build();
    dRaporSheet.getRange("E4:G4").setDataValidation(mlyRule);
  }
}

function getAylar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var islemlerSheet = ss.getSheetByName("ISLEMLER");
  var masterData = islemlerSheet ? islemlerSheet.getDataRange().getValues() : [];
  var aylar = {};
  for (var i = 1; i < masterData.length; i++) {
    var tarihObj = new Date(masterData[i][1]);
    if (isNaN(tarihObj.getTime())) continue;
    var yyyy = tarihObj.getFullYear();
    var mm = ("0" + (tarihObj.getMonth() + 1)).slice(-2);
    aylar[yyyy + "-" + mm + "-01"] = true;
  }
  return aylar;
}

function gecmisiRenklendir() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ayarlar = getAyarlar(); 
  var sheetsToColor = ["ISLEMLER", "MUSTERI_OZET"];
  for (var i = 0; i < sheetsToColor.length; i++) {
    var sheet = ss.getSheetByName(sheetsToColor[i]);
    if(sheet && sheet.getName() === "ISLEMLER"){
      var lr = sheet.getLastRow();
      if (lr > 1) {
        var range = sheet.getRange(2, 1, lr - 1, 9);
        var types = sheet.getRange(2, 4, lr - 1, 1).getValues();
        var colors = [], weights = [];
        for (var r = 0; r < types.length; r++) {
          var c = ayarlar.islemRenkleri[types[r][0]] || "#000000";
          var rowC = [], rowW = [];
          for (var col = 0; col < 9; col++) { rowC.push(c); rowW.push("bold"); }
          colors.push(rowC); weights.push(rowW);
        }
        range.setFontColors(colors).setFontWeights(weights);
      }
    }
  }
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("🎨 Renkler Geri Geldi!");
}

function fisIptalEt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('🗑️ Fiş İptal İşlemi', 'Lütfen iptal etmek istediğiniz kaydın "İşlem ID"sini girin:\n(ISLEMLER sayfası A sütununda bulabilirsiniz, örn: TX-...-1)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var islemID = response.getResponseText().trim();
  if (!islemID) { ui.alert('⚠️ Hata', 'Geçerli bir İşlem ID girmediniz.', ui.ButtonSet.OK); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var islemlerSheet = ss.getSheetByName("ISLEMLER");
  var data = islemlerSheet.getDataRange().getValues();

  var silinecekSatirISLEMLER = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === islemID) {
      silinecekSatirISLEMLER = i + 1;
      break;
    }
  }

  if (silinecekSatirISLEMLER === -1) { ui.alert('⚠️ Bulunamadı', 'Girdiğiniz ID ile eşleşen bir kayıt bulunamadı.', ui.ButtonSet.OK); return; }

  var onay = ui.alert('⚠️ Kayıt Silinecek!', 'İşlem kalıcı olarak silinecek. Onaylıyor musunuz?', ui.ButtonSet.YES_NO);
  if (onay === ui.Button.YES) {
    islemlerSheet.deleteRow(silinecekSatirISLEMLER);
    
    var lorettaSheet = ss.getSheetByName("LORETTA_BORC");
    if (lorettaSheet) {
      var lData = lorettaSheet.getDataRange().getValues();
      for (var x = lData.length - 1; x >= 7; x--) { 
        if (lData[x][0] === islemID) { 
          lorettaSheet.deleteRow(x + 1);
          break; 
        }
      }
    }
    
    SpreadsheetApp.flush();
    guncelleCariDropdown();
    guncelleDashboards();
    hesaplaDonemselRapor(); 
    
    var ozetSheet = ss.getSheetByName("MUSTERI_OZET");
    if (ozetSheet && ozetSheet.getRange("B4").getValue()) {
      hesaplaMusteriOzeti(ozetSheet.getRange("B4").getValue().toString().trim());
    }
    ui.alert('✅ Başarılı', 'Fiş silindi ve bakiyeler onarıldı.', ui.ButtonSet.OK);
  }
}

function sayfalariSirala() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sistemSayfalari = ["AYARLAR", "CARI_KARTLAR", "GIRIS_FORMU", "MUSTERI_OZET", "RAPOR_DONEMSEL", "ISLEMLER", "DASHBOARD_MUSTERI", "DASHBOARD_ATOLYE", "CARI_ARSIV", "DASHBOARD_AYLIK", "LORETTA_BORC"];
  
  var index = 1;
  for (var j = 0; j < sistemSayfalari.length; j++) {
    var sheet = ss.getSheetByName(sistemSayfalari[j]);
    if (sheet) { ss.setActiveSheet(sheet); ss.moveActiveSheet(index); index++; }
  }

  ss.setActiveSheet(ss.getSheetByName("GIRIS_FORMU"));
  SpreadsheetApp.getUi().alert("📑 Düzen Sağlandı!");
}

/**
 * 🟩 EKLENTİ MODÜLÜ 1: PDF DİJİTAL FİŞ OLUŞTURUCU
 */
function butonPdfFis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName("GIRIS_FORMU");
  
  var musteriAdi = formSheet.getRange("C4").getValue().toString().trim() || formSheet.getRange("B4").getValue().toString().trim();
  if(!musteriAdi) {
    SpreadsheetApp.getUi().alert("⚠️ Lütfen önce bir cari seçin ve formdaki verileri doldurun.");
    return;
  }
  
  var formVerileri = formSheet.getRange("A8:G17").getValues();
  var bosMu = true;
  for(var i=0; i<formVerileri.length; i++) {
    if(formVerileri[i][0] !== "") { bosMu = false; break; }
  }
  if(bosMu) {
    SpreadsheetApp.getUi().alert("⚠️ Fişe yazdırılacak işlem bulunamadı. Lütfen önce formu doldurun.");
    return;
  }

  var fisSheet = ss.getSheetByName("FIS_SABLONU");
  if(!fisSheet) {
    fisSheet = ss.insertSheet("FIS_SABLONU");
  } else {
    fisSheet.setFrozenRows(0);
    fisSheet.clear();
  }
  
  fisSheet.setColumnWidth(1, 130); fisSheet.setColumnWidth(2, 90); fisSheet.setColumnWidth(3, 90);
  fisSheet.setColumnWidth(4, 100); fisSheet.setColumnWidth(5, 70); fisSheet.setColumnWidth(6, 220);
  
  fisSheet.getRange("A1:F2").merge().setValue("💎 SİZİN ŞİRKET LOGONUZ / İSMİNİZ")
          .setBackground("#2c3e50").setFontColor("#ffffff").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(16).setFontWeight("bold");
  fisSheet.getRange("A3").setValue("SAYIN:").setFontWeight("bold").setHorizontalAlignment("right");
  fisSheet.getRange("B3:D3").merge().setValue(musteriAdi).setFontWeight("bold").setFontSize(12);
  fisSheet.getRange("E3").setValue("TARİH:").setFontWeight("bold").setHorizontalAlignment("right");
  
  var islemTarihi = formSheet.getRange("B5").getValue();
  fisSheet.getRange("F3").setValue(islemTarihi).setNumberFormat("dd.MM.yyyy").setHorizontalAlignment("right").setFontWeight("bold");

  var headers = ["İşlem Tipi", "İşlem Gramı", "Milyem", "Has Altın", "Adet", "Açıklama"];
  fisSheet.getRange("A5:F5").setValues([headers]).setBackground("#bdc3c7").setFontWeight("bold").setHorizontalAlignment("center");
  
  var yRow = 6;
  for(var j=0; j<formVerileri.length; j++) {
    if(formVerileri[j][0] !== "") {
      var tip = formVerileri[j][0];
      var gr = formVerileri[j][1];
      var mly = formVerileri[j][2];
      var has = formSheet.getRange(8+j, 4).getValue(); 
      var adet = formVerileri[j][5]; 
      var ack = formVerileri[j][6]; 
      fisSheet.getRange(yRow, 1, 1, 6).setValues([[tip, gr, mly, has, adet, ack]]).setHorizontalAlignment("center");
      fisSheet.getRange(yRow, 2).setNumberFormat("#,##0.00");
      fisSheet.getRange(yRow, 4).setNumberFormat("#,##0.00");
      yRow++;
    }
  }
  
  var guncelBakiye = formSheet.getRange("D4").getValue();
  var yeniBakiye = formSheet.getRange("E4").getValue();
  
  yRow++;
  fisSheet.getRange(yRow, 1, 1, 6).setBackground("#ecf0f1");
  yRow++;
  fisSheet.getRange(yRow, 4, 1, 2).merge().setValue("Önceki Bakiye:").setFontWeight("bold").setHorizontalAlignment("right");
  fisSheet.getRange(yRow, 6).setValue(guncelBakiye).setNumberFormat("#,##0.00").setHorizontalAlignment("right");
  yRow++;
  fisSheet.getRange(yRow, 4, 1, 2).merge().setValue("YENİ GÜNCEL BAKİYE:").setFontWeight("bold").setHorizontalAlignment("right").setFontSize(12);
  fisSheet.getRange(yRow, 6).setValue(yeniBakiye).setNumberFormat("#,##0.00").setFontWeight("bold").setHorizontalAlignment("right").setFontSize(12).setBackground("#f1c40f").setFontColor("#2c3e50");
  fisSheet.getRange("A1:F" + yRow).setBorder(true, true, true, true, true, true);
  SpreadsheetApp.flush();
  
  var url = ss.getUrl().replace(/edit$/, '') + 'export?exportFormat=pdf&format=pdf' +
      '&size=A5' +
      '&portrait=true' +
      '&fitw=true' +
      '&sheetnames=false&printtitle=false&pagenumbers=false' +
      '&gridlines=false' +
      '&fzr=false' +
      '&gid=' + fisSheet.getSheetId();
  var html = '<html><body style="font-family:sans-serif; text-align:center; padding-top:10px;">' +
             '<h3 style="color:#2c3e50;">Fişiniz Hazırlandı!</h3>' +
             '<p style="font-size:14px; color:#7f8c8d; margin-bottom:25px;">Aşağıdaki butona tıklayarak bilgisayarınıza veya telefonunuza indirebilirsiniz.</p>' +
             '<a href="' + url + '" target="_blank" style="background-color:#2980b9; color:white; padding:12px 24px; text-decoration:none; border-radius:5px; font-weight:bold; font-size:16px;">PDF İNDİR</a>' +
             '<script>setTimeout(function(){google.script.host.close();}, 8000);</script>' +
             '</body></html>';
             
  var userInterface = HtmlService.createHtmlOutput(html).setWidth(350).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(userInterface, '📄 Fiş Çıktısı Modülü');
}

/**
 * 🟩 EKLENTİ MODÜLÜ 2: PDF CARİ EKSTRE OLUŞTURUCU (SADELEŞTİRİLMİŞ & BASE64 MOTORU)
 */
function butonPdfEkstre() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ozetSheet = ss.getSheetByName("MUSTERI_OZET");

  if (!ozetSheet) {
    SpreadsheetApp.getUi().alert("⚠️ Müşteri Özet sayfası bulunamadı.");
    return;
  }

  var cariAdi = ozetSheet.getRange("B4").getValue().toString().trim();
  if (!cariAdi) {
    SpreadsheetApp.getUi().alert("⚠️ Lütfen önce ekstresini almak istediğiniz bir Cari (Müşteri/Atölye) seçin.");
    return;
  }

  var lr = ozetSheet.getLastRow();
  if (lr < 12) lr = 12; 

  var sheetId = ozetSheet.getSheetId();

  var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?exportFormat=pdf&format=pdf" +
      "&size=A4" +
      "&portrait=true" +
      "&fitw=true" + 
      "&sheetnames=false&printtitle=false&pagenumbers=true" + 
      "&gridlines=false" + 
      "&fzr=false" +
      "&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5" +
      "&r1=11&c1=0&r2=" + lr + "&c2=6" +
      "&gid=" + sheetId;

  var baslangic = ozetSheet.getRange("H3").getDisplayValue() || "Baştan İtibaren";
  var bitis = ozetSheet.getRange("H4").getDisplayValue() || "Bugüne Kadar";
  var donem = baslangic + " ➡️ " + bitis;

  var html = '<html><body style="font-family:sans-serif; text-align:center; padding-top:10px;">' +
             '<h3 style="color:#2c3e50;">📊 Ekstre Hazırlanıyor...</h3>' +
             '<p style="font-size:14px; color:#34495e; margin-bottom:5px;"><b>Cari:</b> ' + cariAdi + '</p>' +
             '<p style="font-size:12px; color:#7f8c8d; margin-bottom:20px;"><b>Dönem:</b> ' + donem + '</p>' +
             
             '<button id="btnIndir" onclick="baslat()" style="background-color:#27ae60; color:white; padding:12px 24px; border:none; border-radius:5px; font-weight:bold; font-size:16px; cursor:pointer; width:80%;">📥 SADELEŞTİRİLMİŞ PDF İNDİR</button>' +
             '<p id="durum" style="font-size:13px; color:#e67e22; margin-top:15px; font-weight:bold; display:none;">Sistemden çekiliyor, lütfen bekleyin... ⏳</p>' +
             
             '<script>' +
             'function baslat() {' +
             '  document.getElementById("btnIndir").style.display = "none";' +
             '  document.getElementById("durum").style.display = "block";' +
             '  google.script.run.withSuccessHandler(indir).getEkstrePdfBase64("' + url + '");' +
             '}' +
             'function indir(base64Data) {' +
             '  var link = document.createElement("a");' +
             '  link.href = "data:application/pdf;base64," + base64Data;' +
             '  link.download = "' + cariAdi + '_Ekstre.pdf";' +
             '  document.body.appendChild(link);' +
             '  link.click();' +
             '  document.body.removeChild(link);' +
             '  document.getElementById("durum").style.color = "#27ae60";' +
             '  document.getElementById("durum").innerHTML = "✅ İndirme Başarılı! Pencereyi kapatabilirsiniz.";' +
             '}' +
             '</script>' +
             '</body></html>';

  var userInterface = HtmlService.createHtmlOutput(html).setWidth(350).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(userInterface, '📄 Cari Ekstre Modülü');
}

/**
 * Arka Planda PDF'i oluşturup şifreli veriyi HTML'e gönderen gizli motor
 */
function getEkstrePdfBase64(url) {
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  return Utilities.base64Encode(response.getBlob().getBytes());
}