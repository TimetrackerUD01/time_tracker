// services/syncService.js - Sync between SQLite and Google Sheets
const moment = require('moment-timezone');
const { CONFIG } = require('../config');

class SyncService {
    constructor(sqliteService, sheetsService) {
        this.sqliteService = sqliteService;
        this.sheetsService = sheetsService;
        this.isSyncing = false;
        this.lastSyncTime = null;
        this.syncInterval = null;
    }

    // ========== Startup: Load from Sheets to SQLite ==========

    async loadFromSheets() {
        console.log('🔄 Loading data from Sheets...');

        try {
            // 1. โหลดรายชื่อพนักงาน
            await this.loadEmployeesFromSheets();

            // 2. 🔧 DISABLED: loadOnWorkFromSheets เพราะจะ overwrite on_work ด้วยข้อมูลเก่าจาก Sheets
            // on_work จะถูกจัดการโดย SQLite เป็นหลัก (clock in/out)
            // await this.loadOnWorkFromSheets();

            // 3. 🔧 DISABLED: ปิดการ import time_records จาก Sheets เพราะข้อมูลจาก Sheets ไม่มี clock_out
            // SQLite เป็นแหล่งข้อมูลหลัก ไม่ต้อง sync กลับจาก Sheets
            // await this.loadRecentRecordsFromSheets();

            // 4. 🔧 DISABLED: Auto-repair ปิดใช้งานเพราะทำให้คนที่ clock out แล้วถูกเพิ่มกลับมา
            // จะกู้คืนเฉพาะคนที่ยังไม่ลงเวลาออก (Clock Out is null) กลับมาที่ on_work
            // console.log('🔧 Auto-repairing on_work status...');
            // await this.sqliteService.repairOnWorkFromTimeRecords();

            console.log('✅ Sync completed');
            this.lastSyncTime = new Date();

            return { success: true };

        } catch (error) {
            console.error('❌ [Sync] Failed to load from Sheets:', error);
            // ไม่ throw error เพื่อให้ระบบยังทำงานต่อได้
            return { success: false, error: error.message };
        }
    }

    async loadEmployeesFromSheets() {
        try {
            const employees = await this.sheetsService.getEmployees();
            if (employees && employees.length > 0) {
                this.sqliteService.bulkInsertEmployees(employees);
            }
        } catch (error) {
            console.error('❌ [Sync] Error loading employees:', error);
        }
    }

    async loadOnWorkFromSheets() {
        try {
            const rows = await this.sheetsService.getCachedSheetData(CONFIG.SHEETS.ON_WORK);

            if (rows && rows.length > 0) {
                const records = rows.map(row => ({
                    employee_name: row.get('ชื่อพนักงาน') || row.get('ชื่อในระบบ'),
                    system_name: row.get('ชื่อในระบบ'),
                    clock_in: row.get('เวลาเข้า'),
                    status: row.get('สถานะ') || 'ทำงาน',
                    userinfo: row.get('หมายเหตุ') || '',
                    location: row.get('พิกัด') || '',
                    location_name: row.get('สถานที่') || '',
                    main_row_id: parseInt(row.get('แถวอ้างอิง') || row.get('แถวในMain')) || null,
                    line_name: row.get('Line Name') || '',
                    line_picture: row.get('Line Picture') || ''
                })).filter(r => r.employee_name && r.clock_in);

                if (records.length > 0) {
                    this.sqliteService.bulkInsertOnWork(records);
                }
            }
        } catch (error) {
            console.error('❌ [Sync] Error loading on_work:', error);
        }
    }

    async loadRecentRecordsFromSheets() {
        try {
            const rows = await this.sheetsService.getCachedSheetData(CONFIG.SHEETS.MAIN);

            if (rows && rows.length > 0) {
                // โหลดเฉพาะข้อมูล 30 วันล่าสุด
                const thirtyDaysAgo = moment().tz(CONFIG.TIMEZONE).subtract(30, 'days');

                const records = rows.map(row => {
                    const clockIn = row._rawData[3]; // column 3: เวลาเข้า

                    // ตรวจสอบว่าเป็นข้อมูล 30 วันล่าสุด
                    let recordDate;
                    if (typeof clockIn === 'string' && clockIn.match(/^\d{2}\/\d{2}\/\d{4}/)) {
                        recordDate = moment(clockIn, 'DD/MM/YYYY HH:mm:ss');
                    } else {
                        recordDate = moment(clockIn);
                    }

                    if (!recordDate.isValid() || recordDate.isBefore(thirtyDaysAgo)) {
                        return null;
                    }

                    return {
                        employee_name: row._rawData[0],
                        line_name: row._rawData[1] || '',
                        line_picture: (row._rawData[2] || '').replace('=IMAGE("', '').replace('")', ''),
                        clock_in: row._rawData[3],
                        userinfo: row._rawData[4] || '',
                        clock_out: row._rawData[5] || '',
                        location_in: row._rawData[6] || '',
                        location_in_name: row._rawData[7] || '',
                        location_out: row._rawData[8] || '',
                        location_out_name: row._rawData[9] || '',
                        working_hours: parseFloat(row._rawData[10]) || 0,
                        note: row._rawData[11] || ''
                    };
                }).filter(r => r !== null && r.employee_name && r.clock_in);

                if (records.length > 0) {
                    this.sqliteService.bulkInsertTimeRecords(records);
                    console.log(`✅ [Sync] Loaded ${records.length} recent records (last 30 days)`);
                }
            }
        } catch (error) {
            console.error('❌ [Sync] Error loading time records:', error);
        }
    }

    // 🆕 Sync เฉพาะข้อมูล "วันนี้" จาก Sheets (เพื่อ update การลบ/แก้ไข)
    async syncCurrentDayFromSheets() {
        try {
            console.log('🔄 [Sync] Syncing current day from Sheets...');

            // ✅ Fix: Clear Cache เพื่อให้มั่นใจว่าได้ข้อมูลล่าสุดที่เราเพิ่ง Sync ขึ้นไป
            // ไม่งั้นถ้า Cache เก่า (ยังไม่มี record ที่เพิ่ง sync) -> เราจะลบ record synced=1 ในเครื่องทิ้ง แล้วไม่ได้ insert กลับมา -> หาย!
            this.sheetsService.clearCache();

            const rows = await this.sheetsService.getCachedSheetData(CONFIG.SHEETS.MAIN); // ตอนนี้จะได้ข้อมูลสดใหม่จาก Google Sheets

            if (rows && rows.length > 0) {
                const today = moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY'); // Format ใน Sheets

                // กรองเฉพาะแถวของวันนี้
                const todayRecords = rows.map(row => {
                    const clockIn = row._rawData[3];

                    // เช็คว่าตรงกับวันนี้หรือไม่
                    if (typeof clockIn === 'string' && clockIn.startsWith(today)) {
                        return {
                            employee_name: row._rawData[0],
                            line_name: row._rawData[1] || '',
                            line_picture: (row._rawData[2] || '').replace('=IMAGE("', '').replace('")', ''),
                            clock_in: row._rawData[3],
                            userinfo: row._rawData[4] || '',
                            clock_out: row._rawData[5] || '',
                            location_in: row._rawData[6] || '',
                            location_in_name: row._rawData[7] || '',
                            location_out: row._rawData[8] || '',
                            location_out_name: row._rawData[9] || '',
                            working_hours: parseFloat(row._rawData[10]) || 0,
                            note: row._rawData[11] || '',
                            synced_to_sheets: 1 // มาจาก Sheets ถือว่า sync แล้ว
                        };
                    }
                    return null;
                }).filter(r => r !== null);

                // ลบข้อมูลวันนี้ใน SQLite แล้วลงใหม่
                if (todayRecords.length > 0) {
                    console.log(`📥 [Sync] Found ${todayRecords.length} records for today from Sheets`);

                    // ⚠️ Fix: ลบเฉพาะข้อมูลวันนี้ที่ Sync แล้วเท่านั้น (synced_to_sheets = 1)
                    // เพื่อป้องกันข้อมูลที่เพิ่ง Clock In (synced_to_sheets = 0) หายไป

                    const todaySlash = moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY');

                    // 🔧 FIX: ลบเฉพาะ record ที่ synced=1 และ ยังไม่มี clock_out 
                    // เพื่อป้องกันการลบข้อมูล clock_out ที่มีอยู่แล้วในเครื่อง
                    this.sqliteService.db.prepare("DELETE FROM time_records WHERE clock_in LIKE ? AND synced_to_sheets = 1 AND (clock_out IS NULL OR TRIM(clock_out) = '')").run(`${todaySlash}%`);

                    // Insert ข้อมูลใหม่จาก Sheets (ทั้งหมดถือว่าเป็น synced_to_sheets = 1)
                    // หมายเหตุ: bulkInsertTimeRecords ต้องจัดการเรื่อง Duplicate (เช่นใช้ INSERT OR IGNORE หรือ REPLACE)
                    // หรือถ้าลบแล้ว insert ก็โอเค (แต่ระวัง record ซ้ำกับ synced=0 ที่ยังไม่ลบ)
                    // ถ้า synced=0 ยังอยู่ แล้วเรา insert ตัวเดียวกันจาก Sheet เข้ามา (ถ้ามันบังเอิญมี) จะซ้ำไหม?
                    // ปกติ bulkInsert จะเป็น INSERT INTO ...
                    // เพื่อความชัวร์ เราควรเช็คว่ามี record ซ้ำไหม ก่อน insert
                    // หรือใช้ bulkUpsert? ในที่นี้ sqliteService.bulkInsertTimeRecords เดิมอาจจะ insert ดื้อๆ

                    // เพื่อความง่ายและปลอดภัย: เราจะ insert เฉพาะ record ที่ยังไม่มีใน DB (โดยเช็คจากเวลา clock_in + employee)
                    // แต่ function bulkInsertTimeRecords อาจจะไม่มีความฉลาดพอ
                    // งั้นเราแก้ function bulkInsertTimeRecords ให้เป็น UPSERT หรือ Ignore หรือเรากรอง record ที่จะ insert ตรงนี้

                    // กรอง todayRecords: เอาเฉพาะอันที่ ไม่มีใน DB หรือ DB มีแต่ synced=1 (ซึ่งลบไปแล้ว)
                    // สรุปง่ายๆ: เอา todayRecords ทั้งหมด มา Loop insert โดยใช้ try-catch เช็ค unique constraint? (แต่เราไม่มี unique constraint ชัดเจน นอกจาก id)

                    // เอาเป็นว่า: ลบเฉพาะ synced=1 ออกไปแล้ว
                    // ดังนั้นใน DB จะเหลือแต่ synced=0
                    // เราต้อง Insert todayRecords (จาก Sheet) เข้าไป
                    // แต่ถ้า Sheet มีข้อมูลที่ตรงกับ synced=0 (เช่น เพิ่ง sync ขึ้นไป) -> มันจะ insert ซ้ำ!
                    // ดังนั้นเราต้องกรอง todayRecords เอาเฉพาะอันที่ไม่ตรงกับ synced=0 ใน DB

                    const existingUnsynced = this.sqliteService.db.prepare('SELECT employee_name, clock_in FROM time_records WHERE clock_in LIKE ? AND synced_to_sheets = 0').all(`${todaySlash}%`);

                    const recordsToInsert = todayRecords.filter(sheetRec => {
                        // ถ้าไม่ตรงกับ unsynced record ใดๆ เลย -> Insert ได้
                        return !existingUnsynced.some(local =>
                            local.employee_name === sheetRec.employee_name &&
                            local.clock_in === sheetRec.clock_in // เทียบเวลาต้องระวัง format 100%
                        );
                    });

                    if (recordsToInsert.length > 0) {
                        this.sqliteService.bulkInsertTimeRecords(recordsToInsert);
                        console.log(`✅ [Sync] Updated local records for today (Inserted ${recordsToInsert.length} from Sheets, Skipped ${todayRecords.length - recordsToInsert.length} overlapping local unsynced)`);
                    } else {
                        console.log('✨ [Sync] No new records from Sheets (All overlap with local unsynced)');
                    }
                } else {
                    // ⚠️ SAFETY: ไม่ลบข้อมูลใน SQLite เมื่อ Sheet ว่างเปล่า
                    // เพราะอาจเป็น error จาก Google Sheets API (เช่น row limit exceeded)
                    // ให้ log คำเตือนแทน
                    console.log('⚠️ [Sync] No records found in Sheets for today - keeping local SQLite data (safe mode)');
                }
            }
        } catch (error) {
            console.error('❌ [Sync] Error syncing current day:', error);
        }
    }

    // ========== Periodic: Sync SQLite to Sheets ==========

    async syncToSheets() {
        if (this.isSyncing) {
            console.log('⏳ [Sync] Already syncing, skipping...');
            return { success: false, reason: 'already_syncing' };
        }

        this.isSyncing = true;
        console.log('🔄 [Sync] Syncing SQLite to Google Sheets...');

        try {
            // ดึงข้อมูลที่ยังไม่ได้ sync
            const unsyncedRecords = this.sqliteService.getUnsyncedRecords();

            if (unsyncedRecords.length === 0) {
                console.log('✅ [Sync] No new records to sync');
                this.isSyncing = false;
                this.lastSyncTime = new Date();
                return { success: true, synced: 0 };
            }

            console.log(`📤 [Sync] Found ${unsyncedRecords.length} unsynced records`);

            // Sync ไป Google Sheets
            const mainSheet = await this.sheetsService.getSheet(CONFIG.SHEETS.MAIN);

            for (const record of unsyncedRecords) {
                try {
                    // ค้นหาแถวที่ตรงกันใน Sheets (ถ้ามี)
                    // หรือเพิ่มแถวใหม่
                    await mainSheet.addRow([
                        record.employee_name,
                        record.line_name,
                        record.line_picture ? `=IMAGE("${record.line_picture}")` : '',
                        record.clock_in,
                        record.userinfo,
                        record.clock_out,
                        record.location_in,
                        record.location_in_name,
                        record.location_out,
                        record.location_out_name,
                        record.working_hours ? record.working_hours.toFixed(2) : ''
                    ]);
                } catch (error) {
                    console.error(`❌ [Sync] Error syncing record ${record.id}:`, error);
                }
            }

            // Mark as synced
            const syncedIds = unsyncedRecords.map(r => r.id);
            this.sqliteService.markAsSynced(syncedIds);

            // Clear sheets cache
            this.sheetsService.clearCache();

            console.log(`✅ [Sync] Synced ${unsyncedRecords.length} records to Sheets`);
            this.lastSyncTime = new Date();
            this.isSyncing = false;

            return { success: true, synced: unsyncedRecords.length };

        } catch (error) {
            console.error('❌ [Sync] Error syncing to Sheets:', error);
            this.isSyncing = false;
            return { success: false, error: error.message };
        }
    }

    // ========== Sync On Work Sheet ==========

    async syncOnWorkToSheets() {
        try {
            const onWorkRecords = this.sqliteService.getOnWorkEmployees();

            // 🆕 SAFETY: ถ้า SQLite on_work ว่าง ให้ข้ามการ sync
            // เพื่อป้องกันการลบข้อมูลใน Google Sheets โดยไม่จำเป็น
            if (onWorkRecords.length === 0) {
                console.log('⚠️ [Sync] SQLite on_work is empty - skipping sync to preserve Sheets data');
                return;
            }

            const onWorkSheet = await this.sheetsService.getSheet(CONFIG.SHEETS.ON_WORK);

            // อ่านข้อมูลปัจจุบัน
            const existingRows = await onWorkSheet.getRows({ offset: 1 });

            // Clear existing data (except header)
            for (const row of existingRows) {
                await row.delete();
            }

            // เพิ่มข้อมูลใหม่
            for (const record of onWorkRecords) {
                await onWorkSheet.addRow([
                    record.clock_in,          // วันที่
                    record.employee_name,     // ชื่อพนักงาน
                    record.clock_in,          // เวลาเข้า
                    record.status,            // สถานะ
                    record.userinfo,          // หมายเหตุ
                    record.location,          // พิกัด
                    record.location_name,     // สถานที่
                    record.main_row_id,       // แถวอ้างอิง
                    record.line_name,         // Line Name
                    record.line_picture,      // Line Picture
                    record.main_row_id,       // แถวในMain
                    record.system_name        // ชื่อในระบบ
                ]);
            }

            console.log(`✅ [Sync] Synced ${onWorkRecords.length} on_work records to Sheets`);

        } catch (error) {
            console.error('❌ [Sync] Error syncing on_work:', error);
        }
    }

    // ========== Start/Stop Periodic Sync ==========

    startPeriodicSync(intervalMs = null) {
        const interval = intervalMs || CONFIG.SYNC?.INTERVAL_MS || 5 * 60 * 1000;

        if (this.syncInterval) {
            clearInterval(this.syncInterval);
        }

        console.log(`🔄 [Sync] Starting periodic sync every ${interval / 1000} seconds`);

        this.syncInterval = setInterval(async () => {
            // 1. ส่งข้อมูลใหม่ขึ้น Sheets
            await this.syncToSheets();

            // 2. 🔧 DISABLED: ปิดการ sync จาก Sheets เพราะข้อมูลมี clock_out ว่าง ทำให้ on_work เพิ่มขึ้นเรื่อยๆ
            // await this.syncCurrentDayFromSheets();

            // 3. 🔧 DISABLED: Auto-repair ปิดใช้งานเพราะทำให้คนที่ clock out แล้วถูกเพิ่มกลับมา
            // await this.sqliteService.repairOnWorkFromTimeRecords();

            // 4. Update on_work status to Sheets
            await this.syncOnWorkToSheets();
        }, interval);

        return { interval };
    }

    stopPeriodicSync() {
        if (this.syncInterval) {
            clearInterval(this.syncInterval);
            this.syncInterval = null;
            console.log('🛑 [Sync] Periodic sync stopped');
        }
    }

    // ========== Force Sync ==========

    async forceSync() {
        console.log('🔄 [Sync] Force sync initiated...');

        try {
            // Sync main records
            await this.syncToSheets();

            // Sync on_work
            await this.syncOnWorkToSheets();

            this.lastSyncTime = new Date();

            return {
                success: true,
                lastSyncTime: this.lastSyncTime,
                stats: this.sqliteService.getStats()
            };

        } catch (error) {
            console.error('❌ [Sync] Force sync failed:', error);
            return { success: false, error: error.message };
        }
    }

    // ========== Get Status ==========

    getStatus() {
        return {
            isSyncing: this.isSyncing,
            lastSyncTime: this.lastSyncTime,
            periodicSyncActive: !!this.syncInterval,
            sqliteStats: this.sqliteService.getStats()
        };
    }
}

module.exports = SyncService;
