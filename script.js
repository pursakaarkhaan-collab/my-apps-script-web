function doGet(e) {
    return HtmlService.createHtmlOutputFromFile('Index')
        .setTitle('Presensi HadirQ')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Auto-run saat file spreadsheet dibuka
 */
function onOpen() {
    createAdminMenu();
}

/**
 * Menambahkan Menu Admin agar mudah diakses
 */
function createAdminMenu() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('🔒 HadirQ Admin')
        .addItem('Reset Rate Limit (Hapus Blokir HP)', 'resetRateLimit')
        .addToUi();
}

/**
 * Reset rate limit for a phone number
 * Konteks: Container-Bound (Menempel di Sheet)
 */
function resetRateLimit(phoneNumber) {
    var ui = SpreadsheetApp.getUi();

    // Jika dipanggil dari Menu (tanpa parameter), minta input via popup
    if (!phoneNumber) {
        var response = ui.prompt(
            '🔒 Reset Rate Limit',
            'Masukkan Nomor HP Orang Tua (contoh: 0812xxxx):',
            ui.ButtonSet.OK_CANCEL
        );

        if (response.getSelectedButton() !== ui.Button.OK) {
            return;
        }
        phoneNumber = response.getResponseText().trim();
    }

    if (!phoneNumber) {
        ui.alert('⚠️ Error: Nomor HP tidak boleh kosong!');
        return;
    }

    // Normalisasi Nomor
    var cleanPhone = phoneNumber.replace(/[^\d]/g, '');
    if (cleanPhone.startsWith('62')) {
        cleanPhone = '0' + cleanPhone.substring(2);
    }

    // Eksekusi Logika Inti (Hapus Cache)
    try {
        var otpCache = CacheService.getScriptCache();
        var attemptsKey = 'otp_attempts_' + cleanPhone;
        otpCache.remove(attemptsKey);

        ui.alert('✅ Berhasil!\n\nRate limit untuk nomor ' + cleanPhone + ' telah dihapus. Orang tua sekarang bisa meminta OTP kembali.');
    } catch (e) {
        ui.alert('❌ Gagal: ' + e.message);
    }
}


// ========== PERFORMANCE OPTIMIZATION CONSTANTS ==========
var MAX_PRESENSI_ROWS = 5000; // Maximum rows to process for reports (prevents quota issues)
var REPORT_CACHE_DURATION = 120; // Cache report results for 2 minutes

// ========== PERFORMANCE OPTIMIZATION FUNCTIONS ==========

/**
 * Get cached student data map for fast NIS lookup
 * Cache expires after 10 minutes
 * @returns {Object} Map of NIS -> {nama, kelas, nohp}
 */
function getCachedStudentMap() {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('studentMap');

    if (cached) {
        try {
            return JSON.parse(cached);
        } catch (e) {
            // Cache corrupted, will refresh
        }
    }

    // Build fresh map from MasterSiswa
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("MasterSiswa");
    if (!sheet) return {};

    var data = sheet.getDataRange().getValues();
    var studentMap = {};

    for (var i = 1; i < data.length; i++) {
        var nis = String(data[i][0]);
        studentMap[nis] = {
            nama: String(data[i][1] || ""),
            kelas: String(data[i][2] || ""),
            nohp: String(data[i][3] || "")
        };
    }

    // Cache for 10 minutes (600 seconds)
    cache.put('studentMap', JSON.stringify(studentMap), 600);

    return studentMap;
}

/**
 * Invalidate student cache (call when MasterSiswa is updated)
 */
function invalidateStudentCache() {
    var cache = CacheService.getScriptCache();
    cache.remove('studentMap');
}

/**
 * Get today's schedule with caching (5 minutes)
 */
function getTodayScheduleCached() {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'todaySchedule_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var cached = cache.get(cacheKey);

    if (cached) {
        try {
            return JSON.parse(cached);
        } catch (e) {
            // Cache corrupted
        }
    }

    var schedule = getTodaySchedule();
    cache.put(cacheKey, JSON.stringify(schedule), 300); // 5 minutes

    return schedule;
}

/**
 * Queue WA notification to be sent asynchronously
 * This prevents blocking the main response
 */
function queueWANotification(nis, nama, tipe, status) {
    try {
        // Store notification data in Properties for async processing
        var props = PropertiesService.getScriptProperties();
        var queue = props.getProperty('waQueue');
        var notifications = queue ? JSON.parse(queue) : [];

        notifications.push({
            nis: nis,
            nama: nama,
            tipe: tipe,
            status: status,
            timestamp: new Date().getTime()
        });

        // Keep only last 50 notifications to prevent overflow
        if (notifications.length > 50) {
            notifications = notifications.slice(-50);
        }

        props.setProperty('waQueue', JSON.stringify(notifications));

        // Create a trigger to process queue in 1 second (non-blocking)
        // Only create if no pending trigger exists
        var triggers = ScriptApp.getProjectTriggers();
        var hasPendingTrigger = triggers.some(function (t) {
            return t.getHandlerFunction() === 'processWAQueue';
        });

        if (!hasPendingTrigger) {
            ScriptApp.newTrigger('processWAQueue')
                .timeBased()
                .after(10000) // 10 seconds delay
                .create();
        }
    } catch (e) {
        // Fallback: send synchronously if queue fails
        sendAttendanceNotif(nis, nama, tipe, status);
    }
}

/**
 * Process queued WA notifications (called by trigger)
 */
function processWAQueue() {
    try {
        var props = PropertiesService.getScriptProperties();
        var queue = props.getProperty('waQueue');

        if (!queue) return;

        var notifications = JSON.parse(queue);
        props.deleteProperty('waQueue'); // Clear queue immediately

        // Process each notification
        for (var i = 0; i < notifications.length; i++) {
            var n = notifications[i];
            sendAttendanceNotif(n.nis, n.nama, n.tipe, n.status, n.timestamp);
        }
    } catch (e) {
        Logger.log('Error processing WA queue: ' + e.message);
    }

    // Clean up the trigger
    var triggers = ScriptApp.getProjectTriggers();
    for (var j = 0; j < triggers.length; j++) {
        if (triggers[j].getHandlerFunction() === 'processWAQueue') {
            ScriptApp.deleteTrigger(triggers[j]);
        }
    }
}

// ========== MONTHLY AUTO-ARCHIVE FUNCTIONS ==========

/**
 * Get list of all archive sheet names that exist
 * @returns {Array} Array of archive sheet names like ["Presensi_2024_11", "Presensi_2024_10"]
 */
function getArchiveSheetNames() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var archiveSheets = [];

    for (var i = 0; i < sheets.length; i++) {
        var name = sheets[i].getName();
        // Match pattern: Presensi_YYYY_MM
        if (/^Presensi_\d{4}_\d{2}$/.test(name)) {
            archiveSheets.push(name);
        }
    }

    // Sort descending (newest first)
    archiveSheets.sort().reverse();
    return archiveSheets;
}

/**
 * Archive old attendance data to monthly archive sheets
 * Moves data from previous months to Presensi_YYYY_MM sheets
 * Called by monthly trigger on 1st of each month
 */
function archiveMonthlyData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPresensi = ss.getSheetByName("Presensi");
    if (!sheetPresensi) return { status: "error", message: "Sheet Presensi tidak ditemukan" };

    var lastRow = sheetPresensi.getLastRow();
    if (lastRow <= 1) return { status: "success", message: "Tidak ada data untuk diarsip" };

    // Get current month/year
    var now = new Date();
    var currentMonth = now.getMonth(); // 0-11
    var currentYear = now.getFullYear();

    // Read all data
    var data = sheetPresensi.getRange(2, 1, lastRow - 1, 7).getValues();

    // Group data by month
    var monthlyData = {};
    var currentMonthRows = [];

    for (var i = 0; i < data.length; i++) {
        var rowDate = data[i][0];
        var recordDate = null;

        if (rowDate instanceof Date) {
            recordDate = rowDate;
        } else {
            // Parse dd/MM/yyyy format
            var parts = String(rowDate).split("/");
            if (parts.length >= 3) {
                recordDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
            }
        }

        if (recordDate) {
            var recordMonth = recordDate.getMonth();
            var recordYear = recordDate.getFullYear();

            // If current month, keep in main sheet
            if (recordMonth === currentMonth && recordYear === currentYear) {
                currentMonthRows.push(data[i]);
            } else {
                // Archive to monthly sheet
                var archiveKey = recordYear + "_" + String(recordMonth + 1).padStart(2, "0");
                if (!monthlyData[archiveKey]) {
                    monthlyData[archiveKey] = [];
                }
                monthlyData[archiveKey].push(data[i]);
            }
        } else {
            // Can't parse date, keep in main sheet
            currentMonthRows.push(data[i]);
        }
    }

    // Create/update archive sheets
    var archivedCount = 0;
    for (var archiveKey in monthlyData) {
        var archiveSheetName = "Presensi_" + archiveKey;
        var archiveSheet = ss.getSheetByName(archiveSheetName);

        if (!archiveSheet) {
            // Create new archive sheet
            archiveSheet = ss.insertSheet(archiveSheetName);
            archiveSheet.appendRow(["Tanggal", "NIS", "Nama", "Status", "JamMasuk", "JamPulang", "Keterangan"]);
            // Format header
            archiveSheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#4361ee").setFontColor("#ffffff");
        }

        // Append data to archive sheet
        var archiveData = monthlyData[archiveKey];
        if (archiveData.length > 0) {
            var lastArchiveRow = archiveSheet.getLastRow();
            archiveSheet.getRange(lastArchiveRow + 1, 1, archiveData.length, 7).setValues(archiveData);
            archivedCount += archiveData.length;
        }
    }

    // Replace main sheet with current month only
    if (archivedCount > 0) {
        // Clear old data (keep header)
        sheetPresensi.getRange(2, 1, lastRow - 1, 7).clearContent();

        // Write current month data back
        if (currentMonthRows.length > 0) {
            sheetPresensi.getRange(2, 1, currentMonthRows.length, 7).setValues(currentMonthRows);
        }
    }

    // Log the archive
    Logger.log("Archived " + archivedCount + " rows to monthly sheets");

    return {
        status: "success",
        message: "Berhasil arsip " + archivedCount + " data ke sheet bulanan",
        archived: archivedCount,
        remaining: currentMonthRows.length
    };
}

/**
 * Setup monthly archive trigger (runs on 1st of each month at 2 AM)
 * Call this once to enable auto-archive
 */
function setupMonthlyArchiveTrigger() {
    // Remove existing triggers first
    removeArchiveTrigger();

    // Create new monthly trigger
    ScriptApp.newTrigger("archiveMonthlyData")
        .timeBased()
        .onMonthDay(1) // 1st of each month
        .atHour(2) // 2 AM
        .create();

    return { status: "success", message: "Trigger arsip bulanan aktif! Akan berjalan setiap tanggal 1 pukul 02:00" };
}

/**
 * Remove archive trigger
 */
function removeArchiveTrigger() {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === "archiveMonthlyData") {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }
    return { status: "success", message: "Trigger arsip bulanan dihapus" };
}

/**
 * Get archive status info
 */
function getArchiveStatus() {
    var archives = getArchiveSheetNames();
    var triggers = ScriptApp.getProjectTriggers();
    var hasArchiveTrigger = triggers.some(function (t) {
        return t.getHandlerFunction() === "archiveMonthlyData";
    });

    return {
        enabled: hasArchiveTrigger,
        archiveSheets: archives,
        archiveCount: archives.length
    };
}

function processPresensi(data) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPresensi = ss.getSheetByName("Presensi");

    if (!sheetPresensi) {
        setupSheet();
        sheetPresensi = ss.getSheetByName("Presensi");
    }

    var inputId = String(data.nis);
    var statusManual = data.status; // For manual attendance (Hadir/Sakit/Ijin/Alpha)
    var keterangan = data.keterangan || "-";
    var waktu = new Date();
    var todayStr = Utilities.formatDate(waktu, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var jamSekarang = Utilities.formatDate(waktu, Session.getScriptTimeZone(), "HH:mm");

    // 1. Find student using CACHED data (OPTIMIZED)
    var studentMap = getCachedStudentMap();
    var student = studentMap[inputId];

    if (!student) {
        return {
            status: 'error',
            message: 'NIS ' + inputId + ' tidak terdaftar!',
            nama: '-',
            waktu: '-'
        };
    }

    var namaSiswa = student.nama;
    var nisSiswa = inputId;

    // 2. Get schedule for today (CACHED)
    var schedule = getTodayScheduleCached();

    // Check if today is active day
    if (schedule.aktif === false) {
        return {
            status: 'error',
            message: 'Hari ini tidak ada jadwal presensi',
            nama: namaSiswa,
            waktu: '-'
        };
    }

    var masukAkhir = schedule.masukAkhir || "07:30";

    // 3. Determine if this is MASUK or PULANG based on time
    var isMasuk = (jamSekarang <= masukAkhir);

    // For manual presensi (Sakit/Ijin/Alpha), always treat as masuk
    if (keterangan === "Manual" && statusManual !== "Hadir") {
        isMasuk = true;
    }

    // 4. Check existing attendance today (OPTIMIZED - only search last 300 rows)
    var lastRow = sheetPresensi.getLastRow();
    var existingRow = -1;
    var existingJamMasuk = "";
    var existingJamPulang = "";
    var existingStatus = "";

    if (lastRow > 1) {
        // Only search the last 300 rows (sufficient for daily attendance)
        var searchRows = Math.min(300, lastRow - 1);
        var startSearchRow = lastRow - searchRows + 1;
        var presensiData = sheetPresensi.getRange(startSearchRow, 1, searchRows, 7).getValues();

        for (var j = presensiData.length - 1; j >= 0; j--) {
            var rowDate = presensiData[j][0];
            var rowNis = presensiData[j][1];

            var rowDateStr = "";
            if (rowDate instanceof Date) {
                rowDateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
            } else {
                rowDateStr = String(rowDate).substring(0, 10);
            }

            if (rowDateStr == todayStr && String(rowNis) == nisSiswa) {
                existingRow = startSearchRow + j; // Actual row in sheet
                existingJamMasuk = String(presensiData[j][4] || "");
                existingJamPulang = String(presensiData[j][5] || "");
                existingStatus = String(presensiData[j][3] || "");
                break;
            }
        }
    }

    // 5. Handle MASUK
    if (isMasuk) {
        if (existingRow > 0 && existingJamMasuk !== "") {
            return {
                status: 'duplicate',
                message: namaSiswa + ' sudah absen masuk!',
                nama: namaSiswa,
                waktu: existingJamMasuk,
                tipe: 'masuk'
            };
        }

        // Determine status: Tepat Waktu or Terlambat
        var statusKehadiran = "Hadir";
        var ketStatus = "";

        if (statusManual && statusManual !== "Hadir") {
            // Manual: Sakit/Ijin/Alpha
            statusKehadiran = statusManual;
            ketStatus = keterangan;
        } else {
            // Scan: Check if late
            if (jamSekarang > masukAkhir) {
                ketStatus = "Terlambat";
            } else {
                ketStatus = "Tepat Waktu";
            }
        }

        var formattedDate = Utilities.formatDate(waktu, Session.getScriptTimeZone(), "dd/MM/yyyy");

        if (existingRow > 0) {
            // Update existing row - BATCH WRITE (faster)
            sheetPresensi.getRange(existingRow, 4, 1, 4).setValues([[statusKehadiran, jamSekarang, "", ketStatus]]);
        } else {
            // New row using setValues (faster than appendRow)
            var newRow = sheetPresensi.getLastRow() + 1;
            sheetPresensi.getRange(newRow, 1, 1, 7).setValues([[formattedDate, nisSiswa, namaSiswa, statusKehadiran, jamSekarang, "", ketStatus]]);
        }

        // Send WA notification DIRECTLY for reliability
        if (statusManual && statusManual !== "Hadir") {
            // Manual: Sakit/Ijin/Alpha
            sendAttendanceNotif(nisSiswa, namaSiswa, "manual", statusKehadiran, waktu.getTime());
        } else {
            // Scan masuk
            sendAttendanceNotif(nisSiswa, namaSiswa, "masuk", "Hadir", waktu.getTime());
        }

        return {
            status: 'success',
            message: 'Presensi MASUK berhasil - ' + ketStatus,
            nama: namaSiswa,
            waktu: jamSekarang,
            tipe: 'masuk',
            ketStatus: ketStatus
        };
    }

    // 6. Handle PULANG
    if (!isMasuk) {
        // Check if already scanned masuk
        if (existingRow <= 0 || existingJamMasuk === "") {
            return {
                status: 'error',
                message: namaSiswa + ' belum absen masuk!',
                nama: namaSiswa,
                waktu: '-',
                tipe: 'pulang'
            };
        }

        // Check if already scanned pulang
        if (existingJamPulang !== "") {
            return {
                status: 'duplicate',
                message: namaSiswa + ' sudah absen pulang!',
                nama: namaSiswa,
                waktu: existingJamPulang,
                tipe: 'pulang'
            };
        }

        // Update pulang time
        sheetPresensi.getRange(existingRow, 6).setValue(jamSekarang);

        // Send WA notification for pulang DIRECTLY
        sendAttendanceNotif(nisSiswa, namaSiswa, "pulang", "Hadir", waktu.getTime());

        return {
            status: 'success',
            message: 'Presensi PULANG berhasil',
            nama: namaSiswa,
            waktu: jamSekarang,
            tipe: 'pulang'
        };
    }
}

// Get today's schedule from Settings
function getTodaySchedule() {
    var settings = getSettings();
    var schedule = settings.schedule || {};

    var dayNames = ['minggu', 'senin', 'selasa', 'rabu', 'kamis', 'jumat', 'sabtu'];
    var today = new Date();
    var dayName = dayNames[today.getDay()];

    if (schedule[dayName]) {
        var daySchedule = schedule[dayName];
        // Check if day is active
        if (daySchedule.aktif === false || daySchedule.aktif === "false") {
            return { aktif: false };
        }
        return daySchedule;
    }

    // Default schedule (aktif by default)
    return {
        aktif: true,
        masukAwal: "06:30",
        masukAkhir: "07:30",
        pulangAwal: "14:00",
        pulangAkhir: "15:00"
    };
}

function getReportData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Presensi");
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var rows = data.slice(1).reverse().slice(0, 20);

    var formattedRows = rows.map(function (row) {
        var dateStr = row[0];
        if (row[0] instanceof Date) {
            dateStr = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "HH:mm");
        }
        return [dateStr, row[2], row[3]];
    });

    return formattedRows;
}

function setupSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var sheetPresensi = ss.getSheetByName("Presensi");
    if (!sheetPresensi) {
        sheetPresensi = ss.insertSheet("Presensi");
        sheetPresensi.appendRow(["Tanggal", "NIS", "Nama", "Status", "JamMasuk", "JamPulang", "Keterangan"]);
    }

    var sheetMaster = ss.getSheetByName("MasterSiswa");
    if (!sheetMaster) {
        sheetMaster = ss.insertSheet("MasterSiswa");
        sheetMaster.appendRow(["NIS", "Nama", "Kelas"]);
        sheetMaster.appendRow(["1001", "Ahmad Dahlan", "XA"]);
        sheetMaster.appendRow(["1002", "Budi Santoso", "XA"]);
        sheetMaster.appendRow(["1003", "Citra Lestari", "XB"]);
        sheetMaster.appendRow(["1004", "Dewi Sartika", "XB"]);
        sheetMaster.appendRow(["1005", "Eko Purnomo", "XC"]);
        sheetMaster.appendRow(["1006", "Fajar Nugraha", "XC"]);
        sheetMaster.appendRow(["1007", "Gita Gutawa", "XC"]);
        sheetMaster.appendRow(["1008", "Hendra Setiawan", "XC"]);
    }
}

// Get list of unique classes for dropdown
function getKelasList() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("MasterSiswa");
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var kelasSet = {};
    for (var i = 1; i < data.length; i++) {
        var kelas = String(data[i][2]).trim();
        if (kelas && kelas !== "") {
            kelasSet[kelas] = true;
        }
    }
    return Object.keys(kelasSet).sort();
}

// Get student data with optional filter
function getDataSiswa(filter) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("MasterSiswa");
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var result = [];
    var filterKelas = filter && filter.kelas ? String(filter.kelas).toLowerCase() : "";
    var filterNama = filter && filter.nama ? String(filter.nama).toLowerCase() : "";

    for (var i = 1; i < data.length; i++) {
        var nis = String(data[i][0]);
        var nama = String(data[i][1]);
        var kelas = String(data[i][2]);

        // Apply filters
        if (filterKelas && kelas.toLowerCase() !== filterKelas) continue;
        if (filterNama && nama.toLowerCase().indexOf(filterNama) === -1) continue;

        result.push({
            no: result.length + 1,
            nis: nis,
            nama: nama,
            kelas: kelas
        });
    }
    return result;
}

// Get today's attendance data with optional filter (OPTIMIZED)
function getDataPresensi(filter) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPresensi = ss.getSheetByName("Presensi");
    if (!sheetPresensi) return [];

    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    // Use cached student map instead of reading sheet (OPTIMIZED)
    var studentMap = getCachedStudentMap();
    var nisKelasMap = {};
    for (var nis in studentMap) {
        nisKelasMap[nis] = studentMap[nis].kelas;
    }

    // Only read last MAX_PRESENSI_ROWS rows (OPTIMIZED)
    var lastRow = sheetPresensi.getLastRow();
    if (lastRow <= 1) return [];

    var rowsToRead = Math.min(MAX_PRESENSI_ROWS, lastRow - 1);
    var startRow = lastRow - rowsToRead + 1;
    var data = sheetPresensi.getRange(startRow, 1, rowsToRead, 7).getValues();
    var result = [];
    var filterKelas = filter && filter.kelas ? String(filter.kelas).toLowerCase() : "";
    var filterNama = filter && filter.nama ? String(filter.nama).toLowerCase() : "";

    for (var i = 0; i < data.length; i++) {
        var rowDate = data[i][0];
        var rowDateStr = "";
        if (rowDate instanceof Date) {
            rowDateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else {
            rowDateStr = String(rowDate).substring(0, 10);
        }

        // Only today's data
        if (rowDateStr !== todayStr) continue;

        var nis = String(data[i][1]);
        var nama = String(data[i][2]);
        var status = String(data[i][3]);
        var jamMasuk = String(data[i][4] || "-");
        var jamPulang = String(data[i][5] || "-");
        var keterangan = String(data[i][6] || "");
        var kelas = nisKelasMap[nis] || "-";

        // Apply filters
        if (filterKelas && kelas.toLowerCase() !== filterKelas) continue;
        if (filterNama && nama.toLowerCase().indexOf(filterNama) === -1) continue;

        result.push({
            no: result.length + 1,
            nis: nis,
            nama: nama,
            kelas: kelas,
            status: status,
            jamMasuk: jamMasuk,
            jamPulang: jamPulang,
            keterangan: keterangan
        });
    }
    return result;
}

// Get attendance recap per student with optional filter (OPTIMIZED)
function getRekapPresensi(filter) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPresensi = ss.getSheetByName("Presensi");
    if (!sheetPresensi) return [];

    // Use cached student map (OPTIMIZED)
    var studentMap = getCachedStudentMap();
    var students = {};
    var filterKelas = filter && filter.kelas ? String(filter.kelas).toLowerCase() : "";
    var filterNama = filter && filter.nama ? String(filter.nama).toLowerCase() : "";

    // Date range filter - parse YYYY-MM-DD format from HTML date input
    var dariTgl = null;
    var sampaiTgl = null;

    if (filter && filter.dariTgl && filter.dariTgl !== "") {
        var dParts = String(filter.dariTgl).split("-");
        if (dParts.length === 3) {
            dariTgl = new Date(parseInt(dParts[0]), parseInt(dParts[1]) - 1, parseInt(dParts[2]), 0, 0, 0, 0);
        }
    }

    if (filter && filter.sampaiTgl && filter.sampaiTgl !== "") {
        var sParts = String(filter.sampaiTgl).split("-");
        if (sParts.length === 3) {
            sampaiTgl = new Date(parseInt(sParts[0]), parseInt(sParts[1]) - 1, parseInt(sParts[2]), 23, 59, 59, 999);
        }
    }

    for (var nis in studentMap) {
        var s = studentMap[nis];
        var nama = s.nama;
        var kelas = s.kelas;

        // Apply filters
        if (filterKelas && kelas.toLowerCase() !== filterKelas) continue;
        if (filterNama && nama.toLowerCase().indexOf(filterNama) === -1) continue;

        students[nis] = {
            nis: nis,
            nama: nama,
            kelas: kelas,
            hadir: 0,
            sakit: 0,
            ijin: 0,
            alpha: 0
        };
    }

    // Count attendance from Presensi - limit to MAX_PRESENSI_ROWS (OPTIMIZED)
    var lastRow = sheetPresensi.getLastRow();
    if (lastRow <= 1) return Object.values ? Object.values(students) : [];

    var rowsToRead = Math.min(MAX_PRESENSI_ROWS, lastRow - 1);
    var startRow = lastRow - rowsToRead + 1;
    var presensiData = sheetPresensi.getRange(startRow, 1, rowsToRead, 7).getValues();

    for (var i = 0; i < presensiData.length; i++) {
        var rowDate = presensiData[i][0];
        var nis = String(presensiData[i][1]);
        var status = String(presensiData[i][3]).toLowerCase();

        // Convert to Date for comparison
        var recordDate = null;
        if (rowDate instanceof Date) {
            recordDate = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate(), 12, 0, 0, 0);
        } else {
            // Try to parse string date (dd/MM/yyyy HH:mm:ss format)
            var dateStr = String(rowDate);
            var parts = dateStr.split(/[\/\s]/);
            if (parts.length >= 3) {
                var day = parseInt(parts[0]);
                var month = parseInt(parts[1]) - 1;
                var year = parseInt(parts[2]);
                recordDate = new Date(year, month, day, 12, 0, 0, 0);
            }
        }

        // Apply date filter - skip if outside range
        if (dariTgl && recordDate && recordDate.getTime() < dariTgl.getTime()) continue;
        if (sampaiTgl && recordDate && recordDate.getTime() > sampaiTgl.getTime()) continue;

        if (students[nis]) {
            if (status === "hadir") students[nis].hadir++;
            else if (status === "sakit") students[nis].sakit++;
            else if (status === "ijin") students[nis].ijin++;
            else if (status === "alpha") students[nis].alpha++;
        }
    }

    // Convert to array
    var result = [];
    var keys = Object.keys(students);
    for (var k = 0; k < keys.length; k++) {
        var s = students[keys[k]];
        s.no = result.length + 1;
        result.push(s);
    }

    return result;
}

// Get students who haven't attended today (for manual attendance) - OPTIMIZED
function getAbsentStudents(filter) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPresensi = ss.getSheetByName("Presensi");

    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    // Get all students who already attended today - only check last 500 rows (OPTIMIZED)
    var attendedToday = {};
    if (sheetPresensi) {
        var lastRow = sheetPresensi.getLastRow();
        if (lastRow > 1) {
            // For today's check, we only need last 500 rows max (covers 1 day of attendance)
            var rowsToRead = Math.min(500, lastRow - 1);
            var startRow = lastRow - rowsToRead + 1;
            var presensiData = sheetPresensi.getRange(startRow, 1, rowsToRead, 7).getValues();

            for (var i = 0; i < presensiData.length; i++) {
                var rowDate = presensiData[i][0];
                var rowDateStr = "";
                if (rowDate instanceof Date) {
                    rowDateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
                } else {
                    rowDateStr = String(rowDate).substring(0, 10);
                }

                if (rowDateStr === todayStr) {
                    var nis = String(presensiData[i][1]);
                    attendedToday[nis] = true;
                }
            }
        }
    }

    // Use cached student map (OPTIMIZED)
    var studentMap = getCachedStudentMap();
    var result = [];
    var filterKelas = filter && filter.kelas ? String(filter.kelas).toLowerCase() : "";
    var filterNama = filter && filter.nama ? String(filter.nama).toLowerCase() : "";

    for (var nis in studentMap) {
        var s = studentMap[nis];
        var nama = s.nama;
        var kelas = s.kelas;

        // Skip if already attended today
        if (attendedToday[nis]) continue;

        // Apply filters
        if (filterKelas && kelas.toLowerCase() !== filterKelas) continue;
        if (filterNama && nama.toLowerCase().indexOf(filterNama) === -1 && nis.indexOf(filterNama) === -1) continue;

        result.push({
            no: result.length + 1,
            nis: nis,
            nama: nama,
            kelas: kelas
        });
    }

    return result;
}

// ==================== SETTINGS FUNCTIONS ====================

// Setup Settings sheet if not exists
function setupSettingsSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var sheetSettings = ss.getSheetByName("Settings");
    if (!sheetSettings) {
        sheetSettings = ss.insertSheet("Settings");
        sheetSettings.appendRow(["Key", "Value"]);
        sheetSettings.appendRow(["schoolName", "MA HadirQ"]);
        // Default schedule (JSON format)
        var defaultSchedule = {
            senin: { masukAwal: "06:30", masukAkhir: "07:30", pulangAwal: "14:00", pulangAkhir: "15:00" },
            selasa: { masukAwal: "06:30", masukAkhir: "07:30", pulangAwal: "14:00", pulangAkhir: "15:00" },
            rabu: { masukAwal: "06:30", masukAkhir: "07:30", pulangAwal: "14:00", pulangAkhir: "15:00" },
            kamis: { masukAwal: "06:30", masukAkhir: "07:30", pulangAwal: "14:00", pulangAkhir: "15:00" },
            jumat: { masukAwal: "06:30", masukAkhir: "07:30", pulangAwal: "11:00", pulangAkhir: "12:00" },
            sabtu: { masukAwal: "06:30", masukAkhir: "07:30", pulangAwal: "12:00", pulangAkhir: "13:00" }
        };
        sheetSettings.appendRow(["schedule", JSON.stringify(defaultSchedule)]);
        sheetSettings.appendRow(["adminPin", "1234"]); // Default PIN
        sheetSettings.appendRow(["logoUrl", ""]); // School logo URL
    }

    var sheetHolidays = ss.getSheetByName("Holidays");
    if (!sheetHolidays) {
        sheetHolidays = ss.insertSheet("Holidays");
        sheetHolidays.appendRow(["Tanggal", "Keterangan"]);
        // Add some example holidays
        sheetHolidays.appendRow(["2025-01-01", "Tahun Baru 2025"]);
        sheetHolidays.appendRow(["2025-03-29", "Hari Raya Nyepi"]);
        sheetHolidays.appendRow(["2025-03-31", "Idul Fitri 1446H"]);
        sheetHolidays.appendRow(["2025-04-01", "Idul Fitri 1446H"]);
    }

    return { settings: sheetSettings, holidays: sheetHolidays };
}

// Get all settings
function getSettings() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Settings");

    if (!sheet) {
        setupSettingsSheet();
        sheet = ss.getSheetByName("Settings");
    }

    var data = sheet.getDataRange().getValues();
    var settings = {};

    for (var i = 1; i < data.length; i++) {
        var key = String(data[i][0]);
        var value = String(data[i][1]);

        // Parse JSON if it looks like JSON
        if (value.indexOf("{") === 0 || value.indexOf("[") === 0) {
            try {
                settings[key] = JSON.parse(value);
            } catch (e) {
                settings[key] = value;
            }
        } else {
            settings[key] = value;
        }
    }

    return settings;
}

// Save a single setting
function saveSetting(key, value) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Settings");

    if (!sheet) {
        setupSettingsSheet();
        sheet = ss.getSheetByName("Settings");
    }

    var data = sheet.getDataRange().getValues();
    var found = false;

    // Convert objects to JSON string
    var valueToSave = (typeof value === "object") ? JSON.stringify(value) : value;

    for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === key) {
            sheet.getRange(i + 1, 2).setValue(valueToSave);
            found = true;
            break;
        }
    }

    if (!found) {
        sheet.appendRow([key, valueToSave]);
    }

    return { status: "success", message: "Pengaturan berhasil disimpan" };
}

// Save schedule
function saveSchedule(schedule) {
    return saveSetting("schedule", schedule);
}

// Save school name
function saveSchoolName(name) {
    return saveSetting("schoolName", name);
}

// Save logo URL
function saveLogoUrl(url) {
    return saveSetting("logoUrl", url);
}

// Save Kop Surat URL
function saveKopSuratUrl(url) {
    return saveSetting("kopSuratUrl", url);
}

// Get holidays list
function getHolidays() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Holidays");

    if (!sheet) {
        setupSettingsSheet();
        sheet = ss.getSheetByName("Holidays");
    }

    var data = sheet.getDataRange().getValues();
    var holidays = [];

    for (var i = 1; i < data.length; i++) {
        var tanggal = data[i][0];
        var keterangan = String(data[i][1]);

        // Format date to YYYY-MM-DD
        var dateStr = "";
        if (tanggal instanceof Date) {
            dateStr = Utilities.formatDate(tanggal, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
            dateStr = String(tanggal);
        }

        if (dateStr && dateStr !== "") {
            holidays.push({
                tanggal: dateStr,
                keterangan: keterangan
            });
        }
    }

    return holidays;
}

// Add a holiday
function addHoliday(tanggal, keterangan) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Holidays");

    if (!sheet) {
        setupSettingsSheet();
        sheet = ss.getSheetByName("Holidays");
    }

    sheet.appendRow([tanggal, keterangan]);
    return { status: "success", message: "Tanggal libur berhasil ditambahkan" };
}

// Delete a holiday
function deleteHoliday(tanggal) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Holidays");

    if (!sheet) return { status: "error", message: "Sheet tidak ditemukan" };

    var data = sheet.getDataRange().getValues();

    for (var i = data.length - 1; i >= 1; i--) {
        var rowDate = data[i][0];
        var dateStr = "";

        if (rowDate instanceof Date) {
            dateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
            dateStr = String(rowDate);
        }

        if (dateStr === tanggal) {
            sheet.deleteRow(i + 1);
            return { status: "success", message: "Tanggal libur berhasil dihapus" };
        }
    }

    return { status: "error", message: "Tanggal tidak ditemukan" };
}

// Verify admin PIN
function verifyPin(pin) {
    var settings = getSettings();
    var savedPin = settings.adminPin || "1234";

    if (String(pin) === String(savedPin)) {
        return { status: "success", message: "PIN benar" };
    } else {
        return { status: "error", message: "PIN salah!" };
    }
}

// Save new admin PIN
function savePin(newPin) {
    if (!newPin || String(newPin).length < 4) {
        return { status: "error", message: "PIN minimal 4 karakter" };
    }
    return saveSetting("adminPin", String(newPin));
}

// ========== LAPORAN & EXPORT FUNCTIONS ==========

// Get report data with filters (for laporan page)
function getLaporanData(filter) {
    try {
        filter = filter || {};
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheetPresensi = ss.getSheetByName("Presensi");
        if (!sheetPresensi) return { data: [], stats: { hadir: 0, sakit: 0, ijin: 0, alpha: 0, terlambat: 0 } };

        // Build student info map using cache
        var studentMapFull = getCachedStudentMap();
        var studentMap = {};
        var filterKelas = filter.kelas ? String(filter.kelas).toLowerCase() : "";
        var filterSiswa = filter.siswa ? String(filter.siswa) : "";

        for (var nis in studentMapFull) {
            var s = studentMapFull[nis];
            if (filterKelas && s.kelas.toLowerCase() !== filterKelas) continue;
            if (filterSiswa && nis !== filterSiswa) continue;

            studentMap[nis] = {
                nis: nis,
                nama: s.nama,
                kelas: s.kelas,
                hadir: 0,
                sakit: 0,
                ijin: 0,
                alpha: 0,
                terlambat: 0,
                records: []
            };
        }

        // Date range filter
        var dariTgl = null;
        var sampaiTgl = null;

        if (filter.dariTgl) {
            var dParts = String(filter.dariTgl).split("-");
            if (dParts.length === 3) {
                dariTgl = new Date(parseInt(dParts[0]), parseInt(dParts[1]) - 1, parseInt(dParts[2]), 0, 0, 0, 0);
            }
        }

        if (filter.sampaiTgl) {
            var sParts = String(filter.sampaiTgl).split("-");
            if (sParts.length === 3) {
                sampaiTgl = new Date(parseInt(sParts[0]), parseInt(sParts[1]) - 1, parseInt(sParts[2]), 23, 59, 59, 999);
            }
        }

        // Count attendance - read from main sheet and archives if needed
        var allPresensiData = [];

        // Determine which sheets to read based on date filter
        var sheetsToRead = ["Presensi"]; // Always read main sheet

        // If filtering by date, also check archive sheets
        if (dariTgl) {
            var archiveSheets = getArchiveSheetNames();
            for (var a = 0; a < archiveSheets.length; a++) {
                // Parse archive sheet name (Presensi_YYYY_MM)
                var parts = archiveSheets[a].split("_");
                if (parts.length === 3) {
                    var archiveYear = parseInt(parts[1]);
                    var archiveMonth = parseInt(parts[2]) - 1; // 0-indexed
                    var archiveStart = new Date(archiveYear, archiveMonth, 1);
                    var archiveEnd = new Date(archiveYear, archiveMonth + 1, 0, 23, 59, 59); // Last day of month

                    // Include this archive if it overlaps with filter range
                    if (sampaiTgl && archiveStart <= sampaiTgl && archiveEnd >= dariTgl) {
                        sheetsToRead.push(archiveSheets[a]);
                    } else if (!sampaiTgl && archiveEnd >= dariTgl) {
                        sheetsToRead.push(archiveSheets[a]);
                    }
                }
            }
        }

        // Read data from all relevant sheets
        for (var si = 0; si < sheetsToRead.length; si++) {
            var sheet = ss.getSheetByName(sheetsToRead[si]);
            if (sheet) {
                var lastRow = sheet.getLastRow();
                if (lastRow > 1) {
                    // Limit rows per sheet to prevent timeout
                    var rowsToRead = Math.min(MAX_PRESENSI_ROWS, lastRow - 1);
                    var startRow = lastRow - rowsToRead + 1;
                    var sheetData = sheet.getRange(startRow, 1, rowsToRead, 7).getValues();
                    allPresensiData = allPresensiData.concat(sheetData);
                }
            }
        }

        var totalStats = { hadir: 0, sakit: 0, ijin: 0, alpha: 0, terlambat: 0 };

        for (var i = 0; i < allPresensiData.length; i++) {
            var rowDate = allPresensiData[i][0];
            var nis = String(allPresensiData[i][1]);
            var status = String(allPresensiData[i][3]).toLowerCase();
            var jamMasuk = String(allPresensiData[i][4] || "");
            var jamPulang = String(allPresensiData[i][5] || "");
            var keterangan = String(allPresensiData[i][6] || "");

            if (!studentMap[nis]) continue;

            // Date filter
            var recordDate = null;
            if (rowDate instanceof Date) {
                recordDate = rowDate;
            } else {
                var parts = String(rowDate).split("/");
                if (parts.length === 3) {
                    recordDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
                }
            }

            if (recordDate) {
                if (dariTgl && recordDate < dariTgl) continue;
                if (sampaiTgl && recordDate > sampaiTgl) continue;
            } else if (dariTgl || sampaiTgl) {
                // If filtering by date but record has no date, skip it
                continue;
            }

            // Count by status
            if (status === "hadir") {
                studentMap[nis].hadir++;
                totalStats.hadir++;
                if (keterangan.toLowerCase().indexOf("terlambat") >= 0) {
                    studentMap[nis].terlambat++;
                    totalStats.terlambat++;
                }
            } else if (status === "sakit") {
                studentMap[nis].sakit++;
                totalStats.sakit++;
            } else if (status === "ijin") {
                studentMap[nis].ijin++;
                totalStats.ijin++;
            } else if (status === "alpha") {
                studentMap[nis].alpha++;
                totalStats.alpha++;
            }

            // Add record for detail view
            var dateStr = "";
            if (recordDate) {
                dateStr = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
            }
            studentMap[nis].records.push({
                tanggal: dateStr,
                status: status,
                jamMasuk: jamMasuk,
                jamPulang: jamPulang,
                keterangan: keterangan
            });
        }

        // Convert to array
        var result = [];
        var no = 1;
        for (var nisKey in studentMap) {
            var s = studentMap[nisKey];
            result.push({
                no: no++,
                nis: s.nis,
                nama: s.nama,
                kelas: s.kelas,
                hadir: s.hadir,
                sakit: s.sakit,
                ijin: s.ijin,
                alpha: s.alpha,
                terlambat: s.terlambat,
                total: s.hadir + s.sakit + s.ijin + s.alpha,
                records: s.records
            });
        }

        return {
            data: result,
            stats: totalStats
        };
    } catch (e) {
        console.error("Error in getLaporanData: " + e.toString());
        return { data: [], stats: { hadir: 0, sakit: 0, ijin: 0, alpha: 0, terlambat: 0 }, error: e.toString() };
    }
}

// Get attendance statistics for charts
function getAttendanceStats(filter) {
    var report = getLaporanData(filter);
    return report.stats;
}

// Get student list for dropdown
function getSiswaList(kelas) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("MasterSiswa");
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var result = [];

    for (var i = 1; i < data.length; i++) {
        var nis = String(data[i][0]);
        var nama = String(data[i][1]);
        var kelasRow = String(data[i][2]);

        if (kelas && kelasRow.toLowerCase() !== String(kelas).toLowerCase()) continue;

        result.push({
            nis: nis,
            nama: nama,
            kelas: kelasRow
        });
    }

    return result;
}

function formatLaporanDateDisplay(val) {
    if (!val) return "-";
    var parts = String(val).split("-");
    if (parts.length === 3) {
        return parts[2] + "/" + parts[1] + "/" + parts[0];
    }
    return String(val);
}

function buildLaporanSpreadsheet(report, filter, signatureData) {
    filter = filter || {};
    report = report || { data: [], stats: { hadir: 0, sakit: 0, ijin: 0, alpha: 0, terlambat: 0 } };
    signatureData = signatureData || {};

    var data = Array.isArray(report.data) ? report.data : [];
    var stats = report.stats || { hadir: 0, sakit: 0, ijin: 0, alpha: 0, terlambat: 0 };

    if (data.length === 0) {
        return { status: "error", message: "Tidak ada data untuk diexport" };
    }

    // Create new spreadsheet
    var settings = getSettings();
    var schoolName = settings.schoolName || "Sekolah";
    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy_HH-mm");
    var fileName = "Laporan_Presensi_" + schoolName.replace(/\s+/g, "_") + "_" + dateStr;

    var newSS = SpreadsheetApp.create(fileName);
    var sheet = newSS.getActiveSheet();
    sheet.setName("Rekap Presensi");

    // Header info (Modified for Kop Surat)
    var schoolNameUpper = schoolName.toUpperCase();
    var kopSuratUrl = settings.kopSuratUrl || "";
    
    // Convert to valid URL if it's an ID
    if (kopSuratUrl && kopSuratUrl.indexOf("http") !== 0) {
        kopSuratUrl = "https://lh3.googleusercontent.com/d/" + kopSuratUrl;
    }
    
    var headerRows = [];
    var startRow = 1;
    
    if (kopSuratUrl) {
        try {
           // Insert Image
           // We need to fetch the blob first to be safe, or direct URL
           // SpreadsheetApp.insertImage can take a URL
           var img = sheet.insertImage(kopSuratUrl, 1, 1);
           
           // Resize image to fit width (approx)
           // standard A4 width in pixels approx 600-800 depending on DPI. 
           // Let's assume we want it 600px wide
           var imgWidth = 600; 
           var originalWidth = img.getWidth();
           var originalHeight = img.getHeight();
           var ratio = originalHeight / originalWidth;
           var newHeight = imgWidth * ratio;
           
           img.setWidth(imgWidth);
           img.setHeight(newHeight);
           
           // Calculate how many rows we need to skip
           // Default row height is 21 pixels.
           var rowsToSkip = Math.ceil(newHeight / 21) + 1;
           startRow = rowsToSkip;
           
        } catch (e) {
            // If image fails, fallback to text header
            headerRows.push([schoolNameUpper]);
            headerRows.push(["LAPORAN REKAPITULASI PRESENSI SISWA"]);
        }
    } else {
        headerRows.push([schoolNameUpper]);
        headerRows.push(["LAPORAN REKAPITULASI PRESENSI SISWA"]);
    }
    
    headerRows.push([""]);
    headerRows.push(["Filter: " + (filter.kelas || "Semua Kelas") + " | " + (filter.siswa ? "NIS " + filter.siswa : "Semua Siswa")]);
    headerRows.push(["Periode: " + formatLaporanDateDisplay(filter.dariTgl) + " s/d " + formatLaporanDateDisplay(filter.sampaiTgl)]);
    headerRows.push([""]);

    // Table header
    var tableHeader = ["No", "NIS", "Nama", "Kelas", "Hadir", "Sakit", "Ijin", "Alpha", "Terlambat", "Total"];

    // Data rows
    var dataRows = data.map(function (d) {
        return [d.no, d.nis, d.nama, d.kelas, d.hadir, d.sakit, d.ijin, d.alpha, d.terlambat, d.total];
    });

    // Summary row
    var summaryRow = ["", "", "TOTAL", "", stats.hadir, stats.sakit, stats.ijin, stats.alpha, stats.terlambat, stats.hadir + stats.sakit + stats.ijin + stats.alpha];

    // Write to sheet
    // var startRow = 1; // Already defined above
    for (var h = 0; h < headerRows.length; h++) {
        sheet.getRange(startRow + h, 1).setValue(headerRows[h][0]);
    }

    var tableStartRow = startRow + headerRows.length;
    sheet.getRange(tableStartRow, 1, 1, tableHeader.length).setValues([tableHeader]);
    sheet.getRange(tableStartRow, 1, 1, tableHeader.length).setFontWeight("bold").setBackground("#4361ee").setFontColor("#ffffff");

    if (dataRows.length > 0) {
        sheet.getRange(tableStartRow + 1, 1, dataRows.length, tableHeader.length).setValues(dataRows);
    }

    sheet.getRange(tableStartRow + dataRows.length + 1, 1, 1, tableHeader.length).setValues([summaryRow]);
    sheet.getRange(tableStartRow + dataRows.length + 1, 1, 1, tableHeader.length).setFontWeight("bold").setBackground("#f0f0f0");

    // Footer (Signatures)
    var footerStartRow = tableStartRow + dataRows.length + 3; // +3 for summary row + spacing
    
    // Kota & Tanggal
    var kota = signatureData.kota || "Kota";
    var tglCetak = signatureData.tanggal || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMMM yyyy");
    
    // Kepsek & Guru
    var kepsekNama = signatureData.kepsekNama || ".........................";
    var kepsekNip = signatureData.kepsekNip ? "NIP. " + signatureData.kepsekNip : "NIP. .........................";
    var guruNama = signatureData.guruNama || ".........................";
    var guruNip = signatureData.guruNip ? "NIP. " + signatureData.guruNip : "NIP. .........................";
    
    // Write Footer
    // Kiri: Kepala Sekolah
    sheet.getRange(footerStartRow, 2).setValue("Mengetahui,");
    sheet.getRange(footerStartRow + 1, 2).setValue("Kepala Madrasah");
    sheet.getRange(footerStartRow + 5, 2).setValue(kepsekNama).setFontWeight("bold").setFontLine("underline");
    sheet.getRange(footerStartRow + 6, 2).setValue(kepsekNip);
    
    // Kanan: Guru Kelas / Wali Kelas
    var colKanan = tableHeader.length - 1; // Agak ke kanan
    sheet.getRange(footerStartRow, colKanan).setValue(kota + ", " + tglCetak);
    sheet.getRange(footerStartRow + 1, colKanan).setValue("Guru Kelas / Wali Kelas");
    sheet.getRange(footerStartRow + 5, colKanan).setValue(guruNama).setFontWeight("bold").setFontLine("underline");
    sheet.getRange(footerStartRow + 6, colKanan).setValue(guruNip);

    // Auto resize columns
    for (var c = 1; c <= tableHeader.length; c++) {
        sheet.autoResizeColumn(c);
    }

    // Format Header Title (Only if no image, or for the sub-info)
    if (!kopSuratUrl) {
        sheet.getRange(1, 1, 1, tableHeader.length).merge().setHorizontalAlignment("center").setFontSize(16).setFontWeight("bold");
        sheet.getRange(2, 1, 1, tableHeader.length).merge().setHorizontalAlignment("center").setFontSize(14).setFontWeight("bold");
    }
    // Filter info spacing
    if (headerRows.length > 0) {
        // Adjust alignment for filter info which is at the bottom of headerRows
        // We need to find where they are.
        // Simplified: Just center align the rows where we wrote headerRows
        // But be careful not to center align the "Filter:" text if we want it left.
        // Let's just keep default alignment for filter info.
    }
    
    // Force write changes
    SpreadsheetApp.flush();

    return {
        status: "success",
        message: "Export berhasil!",
        spreadsheetId: newSS.getId(),
        url: newSS.getUrl(),
        name: fileName
    };
}

// Export report to new Google Spreadsheet
function exportToSpreadsheet(filter, signatureData) {
    filter = filter || {};
    var report = getLaporanData(filter);
    return buildLaporanSpreadsheet(report, filter, signatureData);
}

// Export report based on data currently displayed in UI
function exportPreparedLaporan(report, filter, signatureData) {
    var created = buildLaporanSpreadsheet(report, filter || {}, signatureData || {});
    if (!created || created.status !== "success") {
        return created;
    }

    try {
        // Pastikan spreadsheet baru tersedia
        // Pastikan spreadsheet baru tersedia
        Utilities.sleep(3000);

        var tempSpreadsheetId = created.spreadsheetId;
        var exportUrl = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheetId + "/export?format=xlsx&exportFormat=xlsx";

        // Request menggunakan token skrip (tidak menggunakan DriveApp)
        var response = UrlFetchApp.fetch(exportUrl, {
            method: "get",
            headers: {
                Authorization: "Bearer " + ScriptApp.getOAuthToken()
            },
            muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
            return {
                status: "error",
                message: "Gagal membuat file Excel (HTTP " + response.getResponseCode() + "). Pastikan akun memiliki akses."
            };
        }

        var xlsxBlob = response.getBlob();
        var xlsxName = created.name + ".xlsx";
        xlsxBlob.setName(xlsxName);

        // Tidak lagi melakukan operasi DriveApp (menghindari kebutuhan scope tambahan)
        return {
            status: "success",
            message: "Export Excel berhasil!",
            fileName: xlsxName,
            mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            base64: Utilities.base64Encode(xlsxBlob.getBytes())
        };
    } catch (e) {
        return { status: "error", message: "Gagal export Excel: " + e.message };
    }
}

// ========== MPWA WHATSAPP NOTIFICATION FUNCTIONS ==========

// Send WhatsApp notification via MPWA API
function sendWANotif(number, message) {
    try {
        var mpwaSettings = getMpwaSettings();
        if (!mpwaSettings.sender || !mpwaSettings.apiKey) {
            return { status: "error", message: "MYWA belum dikonfigurasi" };
        }

        // Normalize phone number
        var phone = String(number).replace(/[^0-9]/g, "");
        if (phone.startsWith("0")) {
            phone = "62" + phone.substring(1);
        }
        if (!phone.startsWith("62")) {
            phone = "62" + phone;
        }

        var payload = {
            api_key: mpwaSettings.apiKey,
            sender: mpwaSettings.sender,
            number: phone,
            message: message
        };

        var options = {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        var waUrl = mpwaSettings.waApiUrl || "https://gateway.pdmhadirq.cloud/api/send-message";
        var response = UrlFetchApp.fetch(waUrl, options);
        var result = JSON.parse(response.getContentText());

        if (result.status === true) {
            return { status: "success", message: "WA terkirim" };
        } else {
            return { status: "error", message: result.msg || "Gagal kirim WA" };
        }
    } catch (e) {
        return { status: "error", message: e.message };
    }
}

// Get parent phone number from MasterSiswa (OPTIMIZED - uses cache)
function getParentPhone(nis) {
    var studentMap = getCachedStudentMap();
    var student = studentMap[String(nis)];
    return student ? student.nohp : "";
}

// Build notification message from template
function buildNotifMessage(template, data) {
    if (!template) return "";
    var msg = template;
    msg = msg.replace(/\{nama\}/gi, data.nama || "");
    msg = msg.replace(/\{nis\}/gi, data.nis || "");
    msg = msg.replace(/\{kelas\}/gi, data.kelas || "");
    msg = msg.replace(/\{waktu\}/gi, data.waktu || "");
    msg = msg.replace(/\{tanggal\}/gi, data.tanggal || "");
    msg = msg.replace(/\{status\}/gi, data.status || "");
    return msg;
}

// Get student class from MasterSiswa (OPTIMIZED - uses cache)
function getStudentClass(nis) {
    var studentMap = getCachedStudentMap();
    var student = studentMap[String(nis)];
    return student ? student.kelas : "";
}

// Send attendance notification
function sendAttendanceNotif(nis, nama, tipe, status, scanTimestamp) {
    var mpwaSettings = getMpwaSettings();
    if (!mpwaSettings.enabled) return;

    var parentPhone = getParentPhone(nis);
    if (!parentPhone) return;

    var kelas = getStudentClass(nis);
    // Use scan timestamp if provided, otherwise use current time
    var waktu = scanTimestamp ? new Date(scanTimestamp) : new Date();
    var tanggalStr = Utilities.formatDate(waktu, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var jamStr = Utilities.formatDate(waktu, Session.getScriptTimeZone(), "HH:mm");

    var template = "";
    if (tipe === "masuk") {
        template = mpwaSettings.templateMasuk || "Ananda {nama} telah hadir di sekolah pada {tanggal} pukul {waktu}.";
    } else if (tipe === "pulang") {
        template = mpwaSettings.templatePulang || "Ananda {nama} telah pulang dari sekolah pada {tanggal} pukul {waktu}.";
    } else if (status === "Sakit") {
        template = mpwaSettings.templateSakit || "Ananda {nama} tercatat Sakit pada {tanggal}.";
    } else if (status === "Ijin") {
        template = mpwaSettings.templateIjin || "Ananda {nama} tercatat Ijin pada {tanggal}.";
    } else if (status === "Alpha") {
        template = mpwaSettings.templateAlpha || "Ananda {nama} tercatat Alpha pada {tanggal}.";
    }

    if (!template) return;

    var message = buildNotifMessage(template, {
        nama: nama,
        nis: nis,
        kelas: kelas,
        waktu: jamStr,
        tanggal: tanggalStr,
        status: status
    });

    sendWANotif(parentPhone, message);
}

// Get MPWA settings
function getMpwaSettings() {
    var settings = getSettings();
    return {
        enabled: settings.mpwaEnabled === "true" || settings.mpwaEnabled === true,
        sender: settings.mpwaSender || "",
        waApiUrl: settings.waApiUrl || "https://gateway.pdmhadirq.cloud/api/send-message",
        apiKey: settings.mpwaApiKey || "",
        templateMasuk: settings.templateMasuk || "",
        templatePulang: settings.templatePulang || "",
        templateSakit: settings.templateSakit || "",
        templateIjin: settings.templateIjin || "",
        templateAlpha: settings.templateAlpha || ""
    };
}

// Save MPWA settings
function saveMpwaSettings(data) {
    var results = [];
    if (data.enabled !== undefined) results.push(saveSetting("mpwaEnabled", String(data.enabled)));
    if (data.sender !== undefined) results.push(saveSetting("mpwaSender", data.sender));
    if (data.apiKey !== undefined) results.push(saveSetting("mpwaApiKey", data.apiKey));
    if (data.waApiUrl !== undefined) results.push(saveSetting("waApiUrl", data.waApiUrl));
    if (data.templateMasuk !== undefined) results.push(saveSetting("templateMasuk", data.templateMasuk));
    if (data.templatePulang !== undefined) results.push(saveSetting("templatePulang", data.templatePulang));
    if (data.templateSakit !== undefined) results.push(saveSetting("templateSakit", data.templateSakit));
    if (data.templateIjin !== undefined) results.push(saveSetting("templateIjin", data.templateIjin));
    if (data.templateAlpha !== undefined) results.push(saveSetting("templateAlpha", data.templateAlpha));

    return { status: "success", message: "Pengaturan MYWA disimpan!" };
}

// Test send WA notification
function testSendWA(number, message) {
    return sendWANotif(number, message || "Test notifikasi dari HadirQ");
}

// ========== TTS (TEXT-TO-SPEECH) SETTINGS ==========

// Get TTS settings
function getTtsSettings() {
    var settings = getSettings();
    return {
        enabled: settings.ttsEnabled !== "false" && settings.ttsEnabled !== false // Default ON
    };
}

// Save TTS settings
function saveTtsSettings(data) {
    if (data.enabled !== undefined) {
        saveSetting("ttsEnabled", String(data.enabled));
    }
    return { status: "success", message: "Pengaturan Suara disimpan!" };
}

// ========== QR CODE GENERATOR FUNCTIONS ==========

/**
 * Create QR Code folder in Google Drive if not exists
 * @returns {DriveApp.Folder} QR Code folder
 */
function getOrCreateQRFolder() {
    var folderName = "HadirQ_QR_Codes";
    var folders = DriveApp.getFoldersByName(folderName);
    
    if (folders.hasNext()) {
        return folders.next();
    } else {
        return DriveApp.createFolder(folderName);
    }
}

/**
 * Generate and save QR Code to Google Drive
 * @param {string} nis - Student NIS
 * @param {string} nama - Student name
 * @param {string} kelas - Student class
 * @returns {Object} Result with file info
 */
function generateAndSaveQR(nis, nama, kelas) {
    try {
        // Get QR image from API
        var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=' + encodeURIComponent(nis);
        var response = UrlFetchApp.fetch(qrUrl);
        var blob = response.getBlob();
        
        // Set filename
        var fileName = 'QR_' + nama.replace(/\s+/g, '_') + '_' + nis + '.png';
        blob.setName(fileName);
        
        // Get or create QR folder
        var qrFolder = getOrCreateQRFolder();
        
        // Check if file already exists
        var existingFiles = qrFolder.getFilesByName(fileName);
        if (existingFiles.hasNext()) {
            var existingFile = existingFiles.next();
            return {
                status: "success",
                message: "QR Code sudah ada",
                fileId: existingFile.getId(),
                fileName: fileName,
                downloadUrl: existingFile.getDownloadUrl(),
                viewUrl: existingFile.getUrl()
            };
        }
        
        // Save to Drive
        var file = qrFolder.createFile(blob);
        
        // Make file publicly viewable (optional)
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        return {
            status: "success",
            message: "QR Code berhasil disimpan ke Google Drive",
            fileId: file.getId(),
            fileName: fileName,
            downloadUrl: file.getDownloadUrl(),
            viewUrl: file.getUrl()
        };
        
    } catch (error) {
        return {
            status: "error",
            message: "Error: " + error.message
        };
    }
}

/**
 * Generate QR Code for single student and save to Drive
 * @param {string} nis - Student NIS
 * @returns {Object} Result with student data and file info
 */
function generateSingleQRToDrive(nis) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("MasterSiswa");
        
        if (!sheet) {
            return { status: "error", message: "Sheet MasterSiswa tidak ditemukan!" };
        }
        
        var data = sheet.getDataRange().getValues();
        
        // Find student by NIS
        for (var i = 1; i < data.length; i++) {
            if (String(data[i][0]) === String(nis)) {
                var studentData = {
                    nis: String(data[i][0]),
                    nama: String(data[i][1] || ""),
                    kelas: String(data[i][2] || ""),
                    nohp: String(data[i][3] || "")
                };
                
                // Generate and save QR to Drive
                var qrResult = generateAndSaveQR(studentData.nis, studentData.nama, studentData.kelas);
                
                if (qrResult.status === "success") {
                    studentData.qrFileId = qrResult.fileId;
                    studentData.qrFileName = qrResult.fileName;
                    studentData.qrDownloadUrl = qrResult.downloadUrl;
                    studentData.qrViewUrl = qrResult.viewUrl;
                }
                
                return {
                    status: "success",
                    data: studentData,
                    qrResult: qrResult
                };
            }
        }
        
        return { status: "error", message: "Siswa dengan NIS " + nis + " tidak ditemukan!" };
        
    } catch (error) {
        return { status: "error", message: "Error: " + error.message };
    }
}

/**
 * Generate QR Code for all students in a class and save to Drive
 * @param {string} kelas - Class name
 * @returns {Object} Result with array of student data and file info
 */
function generateBatchQRToDrive(kelas) {
    try {
        var students = getSiswaList(kelas);
        
        if (students.length === 0) {
            return { status: "error", message: "Tidak ada siswa ditemukan di kelas " + kelas };
        }
        
        var results = [];
        var successCount = 0;
        var errorCount = 0;
        
        // Generate QR for each student
        for (var i = 0; i < students.length; i++) {
            var student = students[i];
            var qrResult = generateAndSaveQR(student.nis, student.nama, student.kelas);
            
            if (qrResult.status === "success") {
                student.qrFileId = qrResult.fileId;
                student.qrFileName = qrResult.fileName;
                student.qrDownloadUrl = qrResult.downloadUrl;
                student.qrViewUrl = qrResult.viewUrl;
                successCount++;
            } else {
                student.qrError = qrResult.message;
                errorCount++;
            }
            
            results.push(student);
        }
        
        return {
            status: "success",
            data: results,
            summary: {
                total: students.length,
                success: successCount,
                error: errorCount
            }
        };
        
    } catch (error) {
        return { status: "error", message: "Error: " + error.message };
    }
}

/**
 * Get QR Code file from Google Drive
 * @param {string} fileId - Drive file ID
 * @returns {Object} File blob for download
 */
function getQRFileFromDrive(fileId) {
    try {
        var file = DriveApp.getFileById(fileId);
        return {
            status: "success",
            blob: file.getBlob(),
            fileName: file.getName()
        };
    } catch (error) {
        return {
            status: "error",
            message: "File tidak ditemukan: " + error.message
        };
    }
}

/**
 * List all QR Code files in Drive folder
 * @returns {Array} List of QR files
 */
function listQRFilesInDrive() {
    try {
        var qrFolder = getOrCreateQRFolder();
        var files = qrFolder.getFiles();
        var fileList = [];
        
        while (files.hasNext()) {
            var file = files.next();
            fileList.push({
                id: file.getId(),
                name: file.getName(),
                url: file.getUrl(),
                downloadUrl: file.getDownloadUrl(),
                dateCreated: file.getDateCreated(),
                size: file.getSize()
            });
        }
        
        // Sort by date created (newest first)
        fileList.sort(function(a, b) {
            return b.dateCreated - a.dateCreated;
        });
        
        return {
            status: "success",
            files: fileList,
            folderUrl: qrFolder.getUrl()
        };
        
    } catch (error) {
        return {
            status: "error",
            message: "Error: " + error.message
        };
    }
}
