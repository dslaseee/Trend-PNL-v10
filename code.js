/**
 * Fungsi untuk melayani antarmuka web
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard TREND P&L')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fungsi pembantu untuk mengubah objek Tanggal/Date menjadi format 'mmm-yy' (contoh: Jan-23)
 */
function formatDateMMMYY(dateVal) {
  if (!dateVal) return "";
  
  var months = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
  
  if (Object.prototype.toString.call(dateVal) === '[object Date]') {
    var monthName = months[dateVal.getMonth()];
    var yearShort = String(dateVal.getFullYear()).slice(-2);
    return monthName + "-" + yearShort;
  }
  
  var d = new Date(String(dateVal).trim());
  if (!isNaN(d.getTime())) {
    var monthName = months[d.getMonth()];
    var yearShort = String(d.getFullYear()).slice(-2);
    return monthName + "-" + yearShort;
  }
  
  return String(dateVal);
}

/**
 * Mengambil data dari Sheet TRENDPNL dengan ekstraksi bulan dan tahun yang tepat
 */
function getSheetData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRENDPNL");
  if (!sheet) {
    throw new Error("Sheet 'TRENDPNL' tidak ditemukan!");
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  var idxNamaAkun = headers.indexOf("Nama Akun");
  var idxBulan = headers.indexOf("Bulan");
  var idxDelivery = headers.indexOf("Delivery");
  var idxNonDelivery = headers.indexOf("Non Delivery");
  var idxTotal = headers.indexOf("Total");
  
  var monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

  return rows.map(function(r) {
    var rawBulan = r[idxBulan];
    var d = new Date(rawBulan);
    
    var year = 0;
    var monthIdx = -1;
    var monthName = "";

    if (Object.prototype.toString.call(rawBulan) === '[object Date]' && !isNaN(rawBulan.getTime())) {
      year = rawBulan.getFullYear();
      monthIdx = rawBulan.getMonth();
    } else if (!isNaN(d.getTime())) {
      year = d.getFullYear();
      monthIdx = d.getMonth();
    }

    if (monthIdx >= 0) {
      monthName = monthNames[monthIdx];
    }

    return {
      namaAkun: String(r[idxNamaAkun] || "").trim(),
      bulanFormatted: formatDateMMMYY(rawBulan),
      bulanNama: monthName,
      bulanIdx: monthIdx,
      tahun: year,
      delivery: parseFloat(r[idxDelivery]) || 0,
      nonDelivery: parseFloat(r[idxNonDelivery]) || 0,
      total: parseFloat(r[idxTotal]) || 0
    };
  });
}

/**
 * Mengambil daftar unik Nama Akun untuk dropdown filter
 */
function getAccountList() {
  var rawData = getSheetData();
  var accounts = {};
  rawData.forEach(function(item) {
    if (item.namaAkun) accounts[item.namaAkun] = true;
  });
  return Object.keys(accounts).sort();
}

/**
 * 1. Data Dashboard "TOTAL BIAYA PERIODIK"
 */
function getPeriodicCostDashboard() {
  var rawData = getSheetData();
  var filtered = rawData.filter(function(item) {
    return item.namaAkun.toUpperCase() === "TOTAL BIAYA PERIODIK";
  });
  
  return filtered.map(function(item) {
    return {
      label: item.bulanFormatted,
      delivery: item.delivery,
      nonDelivery: item.nonDelivery
    };
  });
}

/**
 * 2. Data Dashboard Khusus Grafik "NPB"
 */
function getNPBDashboard() {
  var rawData = getSheetData();
  var filtered = rawData.filter(function(item) {
    return item.namaAkun.toUpperCase() === "NPB";
  });
  
  return filtered.map(function(item) {
    return {
      label: item.bulanFormatted,
      total: item.total
    };
  });
}

/**
 * 3. Data Dashboard Filter per Nama Akun
 */
function getAccountDashboard(accountName) {
  var rawData = getSheetData();
  var filtered = rawData.filter(function(item) {
    return item.namaAkun === accountName;
  });
  
  return filtered.map(function(item) {
    return {
      label: item.bulanFormatted,
      total: item.total
    };
  });
}

/**
 * 4. Data Year To Date (YTD) - Mengabaikan TOTAL BIAYA PERIODIK dan NPB
 */
function getYTDData(yearRange, endMonth) {
  var rawData = getSheetData();
  var monthOrder = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var targetMonthIdx = monthOrder.indexOf(endMonth);
  
  var years = yearRange.split("-").map(function(y) { return parseInt(y.trim()); });
  var year1 = years[0];
  var year2 = years[1];
  
  var accMap = {};
  
  rawData.forEach(function(item) {
    // Exclude TOTAL BIAYA PERIODIK dan NPB
    var akunUpper = item.namaAkun.toUpperCase();
    if (akunUpper === "TOTAL BIAYA PERIODIK" || akunUpper === "NPB") {
      return;
    }

    if (item.bulanIdx >= 0 && item.bulanIdx <= targetMonthIdx) {
      if (!accMap[item.namaAkun]) {
        accMap[item.namaAkun] = { namaAkun: item.namaAkun, totalY1: 0, totalY2: 0 };
      }
      if (item.tahun === year1) {
        accMap[item.namaAkun].totalY1 += item.total;
      } else if (item.tahun === year2) {
        accMap[item.namaAkun].totalY2 += item.total;
      }
    }
  });
  
  var list = Object.values(accMap).map(function(item) {
    return {
      namaAkun: item.namaAkun,
      totalY1: item.totalY1,
      totalY2: item.totalY2,
      selisih: item.totalY2 - item.totalY1
    };
  });
  
  var top20 = list.slice().sort(function(a, b) { return b.selisih - a.selisih; }).slice(0, 20);
  var bottom20 = list.slice().sort(function(a, b) { return a.selisih - b.selisih; }).slice(0, 20);
  
  return {
    year1: year1,
    year2: year2,
    top20: top20,
    bottom20: bottom20
  };
}

/**
 * 5. Data Month To Month (MTM) - Mengabaikan TOTAL BIAYA PERIODIK dan NPB
 */
function getMTMData(startMonth, startYear, endMonth, endYear) {
  var rawData = getSheetData();
  var sYear = parseInt(startYear);
  var eYear = parseInt(endYear);
  
  var accMap = {};
  
  rawData.forEach(function(item) {
    // Exclude TOTAL BIAYA PERIODIK dan NPB
    var akunUpper = item.namaAkun.toUpperCase();
    if (akunUpper === "TOTAL BIAYA PERIODIK" || akunUpper === "NPB") {
      return;
    }

    if (!accMap[item.namaAkun]) {
      accMap[item.namaAkun] = { namaAkun: item.namaAkun, totalStart: 0, totalEnd: 0 };
    }
    
    if (item.tahun === sYear && item.bulanNama === startMonth) {
      accMap[item.namaAkun].totalStart += item.total;
    }
    if (item.tahun === eYear && item.bulanNama === endMonth) {
      accMap[item.namaAkun].totalEnd += item.total;
    }
  });
  
  var list = Object.values(accMap).map(function(item) {
    return {
      namaAkun: item.namaAkun,
      totalStart: item.totalStart,
      totalEnd: item.totalEnd,
      selisih: item.totalEnd - item.totalStart
    };
  });
  
  var top20 = list.slice().sort(function(a, b) { return b.selisih - a.selisih; }).slice(0, 20);
  var bottom20 = list.slice().sort(function(a, b) { return a.selisih - b.selisih; }).slice(0, 20);
  
  return {
    top20: top20,
    bottom20: bottom20
  };
}