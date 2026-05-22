// services/sqliteService.js - SQLite Database Service
const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');
const moment = require('moment-timezone');
const { CONFIG } = require('../config');

class SQLiteService {
    constructor() {
        this.db = null;
        this.isInitialized = false;
        this.dbPath = CONFIG.SQLITE?.DB_PATH || './data/timetracker.db';
    }

    initialize() {
        if (this.isInitialized) return;

        try {
            // สร้าง directory ถ้ายังไม่มี
            const dbDir = path.dirname(this.dbPath);
            if (!fs.existsSync(dbDir)) {
                fs.mkdirSync(dbDir, { recursive: true });
                console.log(`📁 Created database directory: ${dbDir}`);
            }

            // เปิด database
            this.db = new Database(this.dbPath);
            this.db.pragma('journal_mode = WAL'); // เพิ่ม performance
            this.db.pragma('foreign_keys = OFF'); // ปิด FK check เพื่อให้ import ข้อมูลจาก Sheets ได้

            console.log(`✅ SQLite database opened: ${this.dbPath}`);

            // สร้าง tables
            this.createTables();

            this.isInitialized = true;
            console.log('✅ SQLite service initialized successfully');

        } catch (error) {
            console.error('❌ Failed to initialize SQLite:', error);
            throw error;
        }
    }

    createTables() {
        // ตาราง employees
        this.db.exec(`
      CREATE TABLE IF NOT EXISTS employees (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
      )
    `);

        // ตาราง time_records (MAIN sheet)
        this.db.exec(`
      CREATE TABLE IF NOT EXISTS time_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_name TEXT NOT NULL,
        line_name TEXT,
        line_picture TEXT,
        clock_in TEXT NOT NULL,
        clock_out TEXT,
        userinfo TEXT,
        location_in TEXT,
        location_in_name TEXT,
        location_out TEXT,
        location_out_name TEXT,
        working_hours REAL,
        note TEXT,
        synced_to_sheets INTEGER DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
      )
    `);

        // ตาราง on_work (ON WORK sheet)
        this.db.exec(`
      CREATE TABLE IF NOT EXISTS on_work (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_name TEXT NOT NULL,
        system_name TEXT,
        clock_in TEXT NOT NULL,
        status TEXT DEFAULT 'ทำงาน',
        userinfo TEXT,
        location TEXT,
        location_name TEXT,
        main_row_id INTEGER,
        line_name TEXT,
        line_picture TEXT
      )
    `);

        // สร้าง indexes
        this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_time_records_date ON time_records(clock_in);
      CREATE INDEX IF NOT EXISTS idx_time_records_employee ON time_records(employee_name);
      CREATE INDEX IF NOT EXISTS idx_on_work_employee ON on_work(employee_name);
      CREATE INDEX IF NOT EXISTS idx_time_records_synced ON time_records(synced_to_sheets);
    `);

        // 🆕 UNIQUE index เพื่อป้องกัน duplicate records (employee + clock_in)
        try {
            this.db.exec(`CREATE UNIQUE INDEX IF NOT EXISTS idx_time_records_unique ON time_records(employee_name, clock_in)`);
        } catch (e) {
            console.log('⚠️ Unique index may already exist or has conflicts');
        }

        // 🆕 ตาราง night_shift_employees (พนักงานกะกลางคืน - ยกเว้นจาก Auto-checkout)
        this.db.exec(`
      CREATE TABLE IF NOT EXISTS night_shift_employees (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_name TEXT UNIQUE NOT NULL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
      )
    `);

        // 🆕 เพิ่ม columns สำหรับ LINE User ID (Personal Dashboard)
        const employeeColumns = this.db.prepare('PRAGMA table_info(employees)').all();
        const hasColumn = (name) => employeeColumns.some(c => c.name === name);
        const addColumn = (name, definition) => {
            try {
                this.db.exec(`ALTER TABLE employees ADD COLUMN ${definition}`);
                console.log(`✅ Added ${name} column to employees`);
            } catch (e) {
                console.warn(`⚠️ Failed to add ${name} column (may already exist):`, e.message);
            }
        };

        if (!hasColumn('line_user_id')) addColumn('line_user_id', 'line_user_id TEXT');
        if (!hasColumn('line_name')) addColumn('line_name', 'line_name TEXT');
        if (!hasColumn('line_picture')) addColumn('line_picture', 'line_picture TEXT');
        if (!hasColumn('registered_at')) addColumn('registered_at', 'registered_at TEXT');

        // Unique index สำหรับ line_user_id (อนุญาตค่า NULL ซ้ำกันได้)
        try {
            this.db.exec('CREATE UNIQUE INDEX IF NOT EXISTS idx_employees_line_user_id ON employees(line_user_id)');
        } catch (e) {
            console.warn('⚠️ Failed to create unique index on line_user_id:', e.message);
        }

        // Import initial night shift employees from config (if table is empty)
        this.initNightShiftFromConfig();

        console.log('✅ Database tables created/verified');
    }

    // ========== Employee Functions ==========

    getEmployees() {
        const stmt = this.db.prepare('SELECT name FROM employees ORDER BY name');
        const rows = stmt.all();
        return rows.map(r => r.name);
    }

    addEmployee(name) {
        try {
            const stmt = this.db.prepare('INSERT OR IGNORE INTO employees (name) VALUES (?)');
            stmt.run(name);
            return true;
        } catch (error) {
            console.error('Error adding employee:', error);
            return false;
        }
    }

    // 🆕 ลบพนักงาน
    deleteEmployee(name) {
        try {
            // ลบจาก employees table
            this.db.prepare('DELETE FROM employees WHERE name = ?').run(name);
            // ลบจาก on_work table (ถ้ามี)
            this.db.prepare('DELETE FROM on_work WHERE employee_name = ?').run(name);
            console.log(`✅ Deleted employee: ${name}`);
            return { success: true, message: `ลบพนักงาน "${name}" สำเร็จ` };
        } catch (error) {
            console.error('Error deleting employee:', error);
            return { success: false, error: error.message };
        }
    }

    // 🆕 ดึงรายชื่อพนักงานทั้งหมดพร้อมรายละเอียด
    getAllEmployeesWithDetails() {
        const employees = this.db.prepare('SELECT id, name, created_at FROM employees ORDER BY name').all();
        const onWorkNames = this.db.prepare('SELECT DISTINCT employee_name FROM on_work').all().map(r => r.employee_name);

        return employees.map(emp => ({
            id: emp.id,
            name: emp.name,
            createdAt: emp.created_at,
            isWorking: onWorkNames.includes(emp.name)
        }));
    }

    // 🆕 Manual Clock In (Admin กำหนดเวลาเอง)
    manualClockIn(data) {
        const { employee, clockInTime, adminNote } = data;

        // 🆕 ข้อมูล Admin เป็นค่าเริ่มต้น
        // 🆕 ข้อมูล Admin เป็นค่าเริ่มต้น (Updated URL as per request)
        const ADMIN_INFO = {
            lineName: 'Got_Songphon 🎶',
            linePicture: 'https://profile.line-scdn.net/0hN8axFz7RERdcAQ-btVduQCFEH3orLxdfJGNcd3wEHSMmYgMSYGVYeHoIR3VxNwURM2QMJHsHTHV3LjYiBzwIKAd3OFMqMg8GAG8VHwsICEcGaFUYFxpZJwdED1F2QhYFPAlcCBgAOEYieAIyAGUVBx1JCGIaYRczAB0',
            locationIn: '17.0374518, 102.4191426',
            locationOut: '17.0374518, 102.4191426'
        };

        // ตรวจสอบว่าลงเวลาเข้าแล้วหรือยัง
        const status = this.getEmployeeStatus(employee);
        if (status.isOnWork) {
            return {
                success: false,
                error: `${employee} ยังลงเวลาเข้าอยู่ ต้องลงเวลาออกก่อน`
            };
        }

        try {
            // เพิ่ม time_records พร้อมข้อมูล Admin (แก้ไข Mapping ให้ถูกต้อง)
            const insertRecord = this.db.prepare(`
                INSERT INTO time_records (employee_name, line_name, line_picture, clock_in, location_in, userinfo, synced_to_sheets)
                VALUES (?, ?, ?, ?, ?, ?, 0)
            `);
            const result = insertRecord.run(
                employee,
                ADMIN_INFO.lineName,       // ลงใน line_name (Column B)
                ADMIN_INFO.linePicture,    // ลงใน line_picture (Column C)
                clockInTime,
                ADMIN_INFO.locationIn,
                adminNote || '', // ลงใน userinfo (Column E - หมายเหตุ) - Optional
            );
            const recordId = result.lastInsertRowid;

            // เพิ่ม on_work พร้อมข้อมูล Admin
            const insertOnWork = this.db.prepare(`
                INSERT INTO on_work (employee_name, system_name, clock_in, status, main_row_id, line_name, line_picture, location_name, userinfo)
                VALUES (?, ?, ?, 'ทำงาน', ?, ?, ?, ?, ?)
            `);
            insertOnWork.run(
                employee,
                employee,
                clockInTime,
                recordId,
                ADMIN_INFO.lineName,
                ADMIN_INFO.linePicture,
                ADMIN_INFO.locationIn,
                adminNote || ''
            );

            // เพิ่มพนักงานถ้ายังไม่มี
            this.addEmployee(employee);

            console.log(`✅ Manual Clock In: ${employee} at ${clockInTime} by Admin`);

            return {
                success: true,
                message: `ลงเวลาเข้าให้ ${employee} เวลา ${clockInTime} สำเร็จ`,
                recordId,
                adminInfo: ADMIN_INFO
            };
        } catch (error) {
            console.error('Error in manual clock in:', error);
            return { success: false, error: error.message };
        }
    }

    // 🆕 Manual Clock Out (Admin กำหนดเวลาเอง)
    manualClockOut(data) {
        const { employee, clockOutTime, adminNote } = data;

        // 🆕 ข้อมูล Admin เป็นค่าเริ่มต้น
        const ADMIN_INFO = {
            locationOut: '17.0374518, 102.4191426'
        };

        const status = this.getEmployeeStatus(employee);
        if (!status.isOnWork) {
            return {
                success: false,
                error: `${employee} ยังไม่ได้ลงเวลาเข้า`
            };
        }

        try {
            const workRecord = status.workRecord;
            const hoursWorked = this.calculateWorkingHours(workRecord.clockIn, clockOutTime);

            // อัพเดท time_records พร้อม location_out
            const updateRecord = this.db.prepare(`
                UPDATE time_records 
                SET clock_out = ?, working_hours = ?, location_out = ?, note = COALESCE(note, '') || ?, synced_to_sheets = 0
                WHERE id = ?
            `);
            updateRecord.run(
                clockOutTime,
                hoursWorked.toFixed(2),
                ADMIN_INFO.locationOut,
                adminNote ? ` | ${adminNote}` : '',
                workRecord.mainRowId
            );

            // ลบจาก on_work
            const deleteOnWork = this.db.prepare('DELETE FROM on_work WHERE main_row_id = ?');
            deleteOnWork.run(workRecord.mainRowId);

            console.log(`✅ Manual Clock Out: ${employee} at ${clockOutTime} (${hoursWorked.toFixed(1)}h) by Admin`);

            return {
                success: true,
                message: `ลงเวลาออกให้ ${employee} เวลา ${clockOutTime} (ทำงาน ${hoursWorked.toFixed(1)} ชม.) สำเร็จ`,
                hoursWorked: hoursWorked.toFixed(2)
            };
        } catch (error) {
            console.error('Error in manual clock out:', error);
            return { success: false, error: error.message };
        }
    }

    // 🆕 แก้ไขเวลาเข้า/ออก
    updateTimeRecord(data) {
        const { recordId, employeeName, newClockIn, newClockOut, adminNote } = data;

        try {
            // ดึงข้อมูลเดิม
            const existingRecord = this.db.prepare('SELECT * FROM time_records WHERE id = ?').get(recordId);

            if (!existingRecord) {
                return { success: false, error: 'ไม่พบรายการนี้' };
            }

            // คำนวณชั่วโมงทำงานใหม่ (ถ้ามีทั้ง clock_in และ clock_out)
            let workingHours = existingRecord.working_hours;
            const finalClockIn = newClockIn || existingRecord.clock_in;
            const finalClockOut = newClockOut || existingRecord.clock_out;

            if (finalClockIn && finalClockOut) {
                workingHours = this.calculateWorkingHours(finalClockIn, finalClockOut);
            }

            // อัปเดท time_records
            const updateStmt = this.db.prepare(`
                UPDATE time_records 
                SET clock_in = ?,
                    clock_out = ?,
                    working_hours = ?,
                    note = COALESCE(note, '') || ?,
                    synced_to_sheets = 0
                WHERE id = ?
            `);

            updateStmt.run(
                finalClockIn,
                finalClockOut,
                workingHours,
                adminNote ? ` | ${adminNote}` : '',
                recordId
            );

            // ถ้ายังไม่มี clock_out และมี on_work ให้อัปเดท on_work ด้วย
            if (newClockIn && !finalClockOut) {
                const updateOnWork = this.db.prepare(`
                    UPDATE on_work SET clock_in = ? WHERE main_row_id = ?
                `);
                updateOnWork.run(finalClockIn, recordId);
            }

            console.log(`✅ Updated time record #${recordId}: ${employeeName}`);

            return {
                success: true,
                message: `แก้ไขเวลาสำเร็จ`,
                data: {
                    recordId,
                    clockIn: finalClockIn,
                    clockOut: finalClockOut,
                    workingHours: workingHours?.toFixed(2) || null
                }
            };
        } catch (error) {
            console.error('Error updating time record:', error);
            return { success: false, error: error.message };
        }
    }

    // 🆕 ดึงรายการเวลาสำหรับแก้ไข
    getTimeRecordsForEdit(employeeName, date) {
        const records = this.db.prepare(`
            SELECT id, employee_name, clock_in, clock_out, working_hours, note
            FROM time_records
            WHERE employee_name = ? AND clock_in LIKE ?
            ORDER BY clock_in DESC
        `).all(employeeName, `${date}%`);

        return records;
    }

    // 🆕 ลบรายการเวลา
    deleteTimeRecord(recordId) {
        try {
            const record = this.db.prepare('SELECT * FROM time_records WHERE id = ?').get(recordId);

            if (!record) {
                return { success: false, error: 'ไม่พบรายการนี้' };
            }

            // ลบจาก on_work ก่อน (ถ้ามี)
            this.db.prepare('DELETE FROM on_work WHERE main_row_id = ?').run(recordId);

            // ลบจาก time_records
            this.db.prepare('DELETE FROM time_records WHERE id = ?').run(recordId);

            console.log(`✅ Deleted time record #${recordId}: ${record.employee_name}`);

            return {
                success: true,
                message: `ลบรายการเวลาของ "${record.employee_name}" สำเร็จ`
            };
        } catch (error) {
            console.error('Error deleting time record:', error);
            return { success: false, error: error.message };
        }
    }

    // ========== Clock In/Out Functions ==========

    getEmployeeStatus(employeeName) {
        const stmt = this.db.prepare(`
      SELECT * FROM on_work 
      WHERE employee_name = ? OR system_name = ?
      LIMIT 1
    `);

        const record = stmt.get(employeeName, employeeName);

        if (record) {
            return {
                isOnWork: true,
                workRecord: {
                    id: record.id,
                    mainRowId: record.main_row_id,
                    clockIn: record.clock_in,
                    systemName: record.system_name,
                    employeeName: record.employee_name
                }
            };
        }

        return { isOnWork: false, workRecord: null };
    }

    async clockIn(data) {
        const { employee, userinfo, lat, lon, line_name, line_picture, mock_time } = data;

        // ตรวจสอบว่าลงเวลาเข้าแล้วหรือยัง
        const status = this.getEmployeeStatus(employee);
        if (status.isOnWork) {
            return {
                success: false,
                message: 'คุณลงเวลาเข้างานไปแล้ว กรุณาลงเวลาออกก่อน',
                employee,
                currentStatus: 'clocked_in',
                clockInTime: status.workRecord?.clockIn
            };
        }

        const timestamp = mock_time
            ? moment(mock_time).tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss')
            : moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss');

        const location = lat && lon ? `${lat},${lon}` : '';

        try {
            // เพิ่มลง time_records
            const insertMain = this.db.prepare(`
        INSERT INTO time_records 
        (employee_name, line_name, line_picture, clock_in, userinfo, location_in, location_in_name, synced_to_sheets)
        VALUES (?, ?, ?, ?, ?, ?, ?, 0)
      `);

            const result = insertMain.run(
                employee,
                line_name || '',
                line_picture || '',
                timestamp,
                userinfo || '',
                location,
                '' // location_name จะถูก update ทีหลัง
            );

            const mainRowId = result.lastInsertRowid;

            // เพิ่มลง on_work
            const insertOnWork = this.db.prepare(`
        INSERT INTO on_work 
        (employee_name, system_name, clock_in, status, userinfo, location, location_name, main_row_id, line_name, line_picture)
        VALUES (?, ?, ?, 'ทำงาน', ?, ?, ?, ?, ?, ?)
      `);

            insertOnWork.run(
                employee,
                employee,
                timestamp,
                userinfo || '',
                location,
                '',
                mainRowId,
                line_name || '',
                line_picture || ''
            );

            // เพิ่มพนักงานถ้ายังไม่มี
            this.addEmployee(employee);

            console.log(`✅ Clock In: ${employee} at ${timestamp.split(' ')[1]}`);

            return {
                success: true,
                message: 'บันทึกเวลาเข้างานสำเร็จ',
                employee,
                time: timestamp,
                currentStatus: 'clocked_in'
            };

        } catch (error) {
            console.error('❌ [SQLite] Clock in error:', error);
            return {
                success: false,
                message: `เกิดข้อผิดพลาด: ${error.message}`,
                employee
            };
        }
    }

    async clockOut(data) {
        const { employee, lat, lon, line_name, mock_time } = data;

        const status = this.getEmployeeStatus(employee);
        if (!status.isOnWork) {
            return {
                success: false,
                message: 'คุณต้องลงเวลาเข้างานก่อน',
                employee,
                currentStatus: 'not_clocked_in'
            };
        }

        const timestamp = mock_time
            ? moment(mock_time).tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss')
            : moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss');

        const location = lat && lon ? `${lat},${lon}` : '';
        const clockInTime = status.workRecord.clockIn;

        // คำนวณชั่วโมงทำงาน
        const hoursWorked = this.calculateWorkingHours(clockInTime, timestamp);

        try {
            // อัปเดต time_records
            const updateMain = this.db.prepare(`
        UPDATE time_records 
        SET clock_out = ?, location_out = ?, location_out_name = ?, working_hours = ?, synced_to_sheets = 0
        WHERE id = ?
      `);

            updateMain.run(
                timestamp,
                location,
                '',
                hoursWorked,
                status.workRecord.mainRowId
            );

            // ลบจาก on_work
            const deleteOnWork = this.db.prepare('DELETE FROM on_work WHERE main_row_id = ?');
            deleteOnWork.run(status.workRecord.mainRowId);

            console.log(`✅ Clock Out: ${employee} at ${timestamp.split(' ')[1]} (${hoursWorked.toFixed(1)}h)`);

            return {
                success: true,
                message: 'บันทึกเวลาออกงานสำเร็จ',
                employee,
                time: timestamp,
                hoursWorked: hoursWorked.toFixed(2),
                currentStatus: 'clocked_out'
            };

        } catch (error) {
            console.error('❌ [SQLite] Clock out error:', error);
            return {
                success: false,
                message: `เกิดข้อผิดพลาด: ${error.message}`,
                employee
            };
        }
    }

    calculateWorkingHours(clockInTime, clockOutTime) {
        try {
            let clockInMoment, clockOutMoment;

            // Parse clock in time
            if (clockInTime.match(/^\d{2}\/\d{2}\/\d{4}/)) {
                clockInMoment = moment.tz(clockInTime, 'DD/MM/YYYY HH:mm:ss', CONFIG.TIMEZONE);
            } else {
                clockInMoment = moment.tz(clockInTime, CONFIG.TIMEZONE);
            }

            // Parse clock out time
            if (clockOutTime.match(/^\d{2}\/\d{2}\/\d{4}/)) {
                clockOutMoment = moment.tz(clockOutTime, 'DD/MM/YYYY HH:mm:ss', CONFIG.TIMEZONE);
            } else {
                clockOutMoment = moment.tz(clockOutTime, CONFIG.TIMEZONE);
            }

            const hours = clockOutMoment.diff(clockInMoment, 'hours', true);
            return hours >= 0 ? hours : 0;

        } catch (error) {
            console.error('Error calculating working hours:', error);
            return 0;
        }
    }

    // ========== Admin Functions ==========

    getOnWorkEmployees() {
        const stmt = this.db.prepare('SELECT * FROM on_work ORDER BY clock_in DESC');
        return stmt.all();
    }

    getAdminStats() {
        const today = moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY');
        console.log('📅 SQLite getAdminStats - Today:', today);

        // นับพนักงานทั้งหมด
        const totalEmployees = this.db.prepare('SELECT COUNT(*) as count FROM employees').get().count;

        // คนที่กำลังทำงาน - ใช้ COUNT จาก on_work table โดยตรง
        const workingNow = this.db.prepare('SELECT COUNT(*) as count FROM on_work').get().count;
        console.log('👥 SQLite workingNow (from on_work table):', workingNow);

        // คนที่มาทำงานวันนี้ - ใช้ LIKE กับ date format DD/MM/YYYY
        const presentToday = this.db.prepare(`
      SELECT COUNT(DISTINCT employee_name) as count 
      FROM time_records 
      WHERE clock_in LIKE ?
    `).get(`${today}%`)?.count || workingNow;
        console.log('📊 SQLite presentToday:', presentToday);

        const absentToday = Math.max(0, totalEmployees - presentToday);

        // รายชื่อพนักงานที่กำลังทำงาน
        const onWorkRows = this.getOnWorkEmployees();
        const workingEmployees = onWorkRows.map(row => {
            const hours = this.calculateWorkingHours(row.clock_in, moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss'));
            return {
                name: row.employee_name,
                clockIn: row.clock_in.split(' ')[1] || row.clock_in,
                workingHours: `${hours.toFixed(1)} ชม.`
            };
        });

        return {
            totalEmployees,
            presentToday,
            workingNow,
            absentToday,
            workingEmployees
        };
    }

    // 🆕 ดึงรายชื่อพนักงานตามประเภท
    getDetailedStats(type) {
        const today = moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY');
        const lateTime = '08:30:00'; // เวลามาสาย

        switch (type) {
            case 'present': {
                // คนที่มาทำงานวันนี้
                const rows = this.db.prepare(`
                    SELECT DISTINCT employee_name, MIN(clock_in) as first_clock_in
                    FROM time_records 
                    WHERE clock_in LIKE ?
                    GROUP BY employee_name
                    ORDER BY first_clock_in
                `).all(`${today}%`);

                return rows.map(row => ({
                    name: row.employee_name,
                    clockIn: row.first_clock_in.split(' ')[1] || row.first_clock_in,
                    status: this.isLate(row.first_clock_in) ? 'สาย' : 'ตรงเวลา'
                }));
            }

            case 'late': {
                // คนที่มาสาย (หลังเวลา 08:30)
                const rows = this.db.prepare(`
                    SELECT DISTINCT employee_name, MIN(clock_in) as first_clock_in
                    FROM time_records 
                    WHERE clock_in LIKE ?
                    GROUP BY employee_name
                    ORDER BY first_clock_in
                `).all(`${today}%`);

                return rows.filter(row => this.isLate(row.first_clock_in)).map(row => ({
                    name: row.employee_name,
                    clockIn: row.first_clock_in.split(' ')[1] || row.first_clock_in,
                    lateBy: this.calculateLateMinutes(row.first_clock_in)
                }));
            }

            case 'absent': {
                // คนที่ขาดงาน - พนักงานทั้งหมดที่ไม่มีใน time_records วันนี้
                const allEmployees = this.db.prepare('SELECT name FROM employees').all();
                const presentNames = this.db.prepare(`
                    SELECT DISTINCT employee_name 
                    FROM time_records 
                    WHERE clock_in LIKE ?
                `).all(`${today}%`).map(r => r.employee_name);

                return allEmployees
                    .filter(emp => !presentNames.includes(emp.name))
                    .map(emp => ({ name: emp.name }));
            }

            case 'working': {
                // คนที่กำลังทำงาน
                const onWorkRows = this.getOnWorkEmployees();
                return onWorkRows.map(row => {
                    const hours = this.calculateWorkingHours(row.clock_in, moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss'));
                    return {
                        name: row.employee_name,
                        clockIn: row.clock_in.split(' ')[1] || row.clock_in,
                        workingHours: `${hours.toFixed(1)} ชม.`
                    };
                });
            }

            default:
                return [];
        }
    }

    // Helper: ตรวจสอบว่ามาสายหรือไม่ (หลัง 08:30)
    isLate(clockIn) {
        if (!clockIn) return false;
        const timePart = clockIn.split(' ')[1];
        if (!timePart) return false;

        // แปลงเวลาเป็นนาที เพื่อเปรียบเทียบถูกต้อง
        const [h, m] = timePart.split(':').map(Number);

        // 🆕 กะกลางคืน (18:00-06:00) ไม่นับเป็นสาย
        if (h >= 18 || h < 6) {
            return false;
        }

        // กะเช้า: สายถ้ามาหลัง 08:30
        const clockInMinutes = h * 60 + m;
        const lateThreshold = 8 * 60 + 30; // 08:30 = 510 นาที

        return clockInMinutes > lateThreshold;
    }

    // Helper: คำนวณมาสายกี่นาที
    calculateLateMinutes(clockIn) {
        if (!clockIn) return '0 นาที';
        const timePart = clockIn.split(' ')[1];
        if (!timePart) return '0 นาที';

        const [h, m] = timePart.split(':').map(Number);

        // 🆕 กะกลางคืน (18:00-06:00) ไม่นับเป็นสาย
        if (h >= 18 || h < 6) {
            return '0 นาที';
        }

        const clockInMinutes = h * 60 + m;
        const lateThreshold = 8 * 60 + 30;

        if (clockInMinutes <= lateThreshold) return '0 นาที';

        const lateMinutes = clockInMinutes - lateThreshold;

        if (lateMinutes >= 60) {
            return `${Math.floor(lateMinutes / 60)} ชม. ${lateMinutes % 60} นาที`;
        }
        return `${lateMinutes} นาที`;
    }

    // ========== Report Functions ==========

    getReportDataForExport(type, params) {
        // ใช้ SQLite เป็นแหล่งข้อมูลหลักสำหรับรายงาน (ลดการพึ่งพา Google Sheets)
        if (!this.isInitialized) {
            this.initialize();
        }

        const parseClockToMoment = (value) => {
            if (!value) return null;

            if (typeof value === 'string') {
                if (value.match(/^\d{2}\/\d{2}\/\d{4}/)) {
                    return moment.tz(value, 'DD/MM/YYYY HH:mm:ss', CONFIG.TIMEZONE);
                }
                if (value.match(/^\d{4}-\d{2}-\d{2}/)) {
                    return moment.tz(value, 'YYYY-MM-DD HH:mm:ss', CONFIG.TIMEZONE);
                }
            }

            const fallback = moment(value).tz(CONFIG.TIMEZONE);
            return fallback.isValid() ? fallback : null;
        };

        const rows = this.db.prepare('SELECT * FROM time_records').all();
        let filteredRows = [];

        switch (type) {
            case 'daily': {
                const targetDate = moment(params.date).tz(CONFIG.TIMEZONE).format('YYYY-MM-DD');
                filteredRows = rows.filter(row => {
                    const clockMoment = parseClockToMoment(row.clock_in);
                    return clockMoment && clockMoment.format('YYYY-MM-DD') === targetDate;
                });
                break;
            }
            case 'monthly': {
                const month = parseInt(params.month, 10);
                let year = parseInt(params.year, 10);
                // 🆕 แปลง พ.ศ. → ค.ศ. อัตโนมัติ (ถ้า year > 2500 แสดงว่าเป็น พ.ศ.)
                if (year > 2500) {
                    console.log(`🔧 Converting Thai year ${year} to AD year ${year - 543}`);
                    year = year - 543;
                }
                console.log(`📊 [SQLite] Filtering monthly report: month=${month}, year=${year}`);
                filteredRows = rows.filter(row => {
                    const clockMoment = parseClockToMoment(row.clock_in);
                    return clockMoment &&
                        (clockMoment.month() + 1) === month &&
                        clockMoment.year() === year;
                });
                break;
            }
            case 'range': {
                const startMoment = moment(params.startDate).tz(CONFIG.TIMEZONE).startOf('day');
                const endMoment = moment(params.endDate).tz(CONFIG.TIMEZONE).endOf('day');
                filteredRows = rows.filter(row => {
                    const clockMoment = parseClockToMoment(row.clock_in);
                    return clockMoment && clockMoment.isBetween(startMoment, endMoment, null, '[]');
                });
                break;
            }
            default:
                throw new Error(`Unsupported report type for SQLite: ${type}`);
        }

        return filteredRows.map((row, index) => ({
            no: index + 1,
            employee: row.employee_name || '',
            lineName: row.line_name || '',
            clockIn: row.clock_in || '',
            clockOut: row.clock_out || '',
            note: row.note || row.userinfo || '',
            workingHours: row.working_hours || '',
            locationIn: row.location_in_name || row.location_in || '',
            locationOut: row.location_out_name || row.location_out || '',
            userInfo: row.userinfo || ''
        }));
    }

    // ========== Sync Helper Functions ==========

    getUnsyncedRecords() {
        const stmt = this.db.prepare('SELECT * FROM time_records WHERE synced_to_sheets = 0');
        return stmt.all();
    }

    markAsSynced(ids) {
        const stmt = this.db.prepare('UPDATE time_records SET synced_to_sheets = 1 WHERE id = ?');
        const transaction = this.db.transaction((ids) => {
            for (const id of ids) {
                stmt.run(id);
            }
        });
        transaction(ids);
    }

    // ========== Bulk Insert Functions (สำหรับ sync จาก Sheets) ==========

    bulkInsertEmployees(employees) {
        const stmt = this.db.prepare('INSERT OR IGNORE INTO employees (name) VALUES (?)');
        const transaction = this.db.transaction((employees) => {
            for (const emp of employees) {
                stmt.run(emp);
            }
        });
        transaction(employees);
        console.log(`✅ [SQLite] Imported ${employees.length} employees`);
    }

    bulkInsertTimeRecords(records) {
        // 🆕 ใช้ INSERT OR IGNORE เพื่อป้องกัน duplicate (employee_name + clock_in)
        const stmt = this.db.prepare(`
      INSERT OR IGNORE INTO time_records 
      (employee_name, line_name, line_picture, clock_in, clock_out, userinfo, location_in, location_in_name, location_out, location_out_name, working_hours, note, synced_to_sheets)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
    `);

        const transaction = this.db.transaction((records) => {
            for (const r of records) {
                stmt.run(
                    r.employee_name,
                    r.line_name || '',
                    r.line_picture || '',
                    r.clock_in,
                    r.clock_out || '',
                    r.userinfo || '',
                    r.location_in || '',
                    r.location_in_name || '',
                    r.location_out || '',
                    r.location_out_name || '',
                    r.working_hours || 0,
                    r.note || ''
                );
            }
        });

        transaction(records);
        console.log(`✅ [SQLite] Imported ${records.length} time records`);
    }

    bulkInsertOnWork(records) {
        // Clear existing on_work data first
        this.db.exec('DELETE FROM on_work');

        const stmt = this.db.prepare(`
      INSERT INTO on_work 
      (employee_name, system_name, clock_in, status, userinfo, location, location_name, main_row_id, line_name, line_picture)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

        const transaction = this.db.transaction((records) => {
            for (const r of records) {
                stmt.run(
                    r.employee_name,
                    r.system_name || r.employee_name,
                    r.clock_in,
                    r.status || 'ทำงาน',
                    r.userinfo || '',
                    r.location || '',
                    r.location_name || '',
                    r.main_row_id || null,
                    r.line_name || '',
                    r.line_picture || ''
                );
            }
        });

        transaction(records);
        console.log(`✅ [SQLite] Imported ${records.length} on_work records`);
    }

    // ========== Utility Functions ==========

    close() {
        if (this.db) {
            this.db.close();
            console.log('SQLite database closed');
        }
    }

    // 🆕 Repair On-Work Status: กู้คืนสถานะการทำงานจาก Time Records
    repairOnWorkFromTimeRecords() {
        try {
            const todaySlash = moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY');

            // 1. หา Time Records ของวันนี้ ที่ยังไม่มีเวลาออก (Clock Out is null or empty or whitespace)
            // 🔧 เพิ่ม TRIM() เพื่อจัดการกับ whitespace และ LENGTH() เพื่อเช็คค่าว่างจริงๆ
            const openRecords = this.db.prepare(`
                SELECT * FROM time_records 
                WHERE clock_in LIKE ? 
                AND (clock_out IS NULL OR TRIM(clock_out) = '' OR LENGTH(TRIM(clock_out)) = 0)
            `).all(`${todaySlash}%`);

            console.log(`🔧 Repair: Found ${openRecords.length} truly open records for today (no clock_out).`);

            let repairedCount = 0;
            const repairedEmployees = [];

            // 2. เช็คว่าแต่ละคนมีใน on_work หรือยัง
            const checkOnWork = this.db.prepare('SELECT id FROM on_work WHERE employee_name = ?');
            const restoreOnWork = this.db.prepare(`
                INSERT INTO on_work 
                (employee_name, system_name, clock_in, status, userinfo, location, location_name, main_row_id, line_name, line_picture)
                VALUES (?, ?, ?, 'ทำงาน', ?, ?, ?, ?, ?, ?)
            `);

            for (const record of openRecords) {
                const onWork = checkOnWork.get(record.employee_name);

                if (!onWork) {
                    // ถ้าไม่มีใน on_work -> กู้คืนกลับมา
                    console.log(`🔧 Repairing status for: ${record.employee_name}`);

                    restoreOnWork.run(
                        record.employee_name,
                        record.employee_name, // system_name (ใช้ชื่อเดียวกัน)
                        record.clock_in,
                        record.userinfo || '',
                        record.location_in || '',
                        record.location_in_name || '',
                        record.id, // main_row_id
                        record.line_name || '',
                        record.line_picture || ''
                    );

                    repairedCount++;
                    repairedEmployees.push(record.employee_name);
                }
            }

            return {
                success: true,
                totalOpen: openRecords.length,
                repairedCount,
                repairedEmployees
            };

        } catch (error) {
            console.error('❌ Repair error:', error);
            return {
                success: false,
                error: error.message
            };
        }
    }

    getStats() {
        return {
            employees: this.db.prepare('SELECT COUNT(*) as count FROM employees').get().count,
            timeRecords: this.db.prepare('SELECT COUNT(*) as count FROM time_records').get().count,
            onWork: this.db.prepare('SELECT COUNT(*) as count FROM on_work').get().count,
            unsynced: this.db.prepare('SELECT COUNT(*) as count FROM time_records WHERE synced_to_sheets = 0').get().count
        };
    }

    // ========== 🆕 Night Shift Employee Functions ==========

    // Initialize night shift employees from config (only if table is empty)
    initNightShiftFromConfig() {
        try {
            const count = this.db.prepare('SELECT COUNT(*) as count FROM night_shift_employees').get().count;
            if (count === 0 && CONFIG.AUTO_CHECKOUT?.EXEMPT_EMPLOYEES?.length > 0) {
                const stmt = this.db.prepare('INSERT OR IGNORE INTO night_shift_employees (employee_name) VALUES (?)');
                for (const name of CONFIG.AUTO_CHECKOUT.EXEMPT_EMPLOYEES) {
                    stmt.run(name);
                }
                console.log(`✅ Imported ${CONFIG.AUTO_CHECKOUT.EXEMPT_EMPLOYEES.length} night shift employees from config`);
            }
        } catch (error) {
            console.error('Error initializing night shift from config:', error);
        }
    }

    // Get all night shift employees
    getNightShiftEmployees() {
        return this.db.prepare('SELECT * FROM night_shift_employees ORDER BY employee_name').all();
    }

    // Add night shift employee
    addNightShiftEmployee(employeeName) {
        try {
            const stmt = this.db.prepare('INSERT OR IGNORE INTO night_shift_employees (employee_name) VALUES (?)');
            const result = stmt.run(employeeName);
            if (result.changes > 0) {
                console.log(`✅ Added night shift employee: ${employeeName}`);
                return { success: true, message: `เพิ่ม ${employeeName} เป็นพนักงานกะกลางคืนแล้ว` };
            }
            return { success: false, message: `${employeeName} อยู่ในรายชื่อกะกลางคืนอยู่แล้ว` };
        } catch (error) {
            console.error('Error adding night shift employee:', error);
            return { success: false, error: error.message };
        }
    }

    // Remove night shift employee
    removeNightShiftEmployee(employeeName) {
        try {
            const stmt = this.db.prepare('DELETE FROM night_shift_employees WHERE employee_name = ?');
            const result = stmt.run(employeeName);
            if (result.changes > 0) {
                console.log(`✅ Removed night shift employee: ${employeeName}`);
                return { success: true, message: `ลบ ${employeeName} ออกจากกะกลางคืนแล้ว` };
            }
            return { success: false, message: `ไม่พบ ${employeeName} ในรายชื่อกะกลางคืน` };
        } catch (error) {
            console.error('Error removing night shift employee:', error);
            return { success: false, error: error.message };
        }
    }

    // Check if employee is night shift
    isNightShiftEmployee(employeeName) {
        const result = this.db.prepare('SELECT COUNT(*) as count FROM night_shift_employees WHERE employee_name = ?').get(employeeName);
        return result.count > 0;
    }

    // ========== 🆕 Live Dashboard Functions ==========

    /**
     * ดึงรายการลงเวลาล่าสุด (สำหรับ Live Feed)
     * @param {number} limit - จำนวนรายการที่ต้องการ
     * @param {string} date - วันที่ต้องการดู (DD/MM/YYYY) หรือ null สำหรับวันนี้
     * @returns {Array} รายการลงเวลา
     */
    getRecentActivity(limit = 30, date = null) {
        const targetDate = date || moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY');

        // ดึงรายการ clock in ของวันที่กำหนด
        const clockInRecords = this.db.prepare(`
            SELECT 
                id,
                employee_name,
                clock_in as time,
                'in' as type,
                line_picture,
                created_at
            FROM time_records
            WHERE clock_in LIKE ?
            ORDER BY id DESC
        `).all(`${targetDate}%`);

        // ดึงรายการ clock out ของวันที่กำหนด
        const clockOutRecords = this.db.prepare(`
            SELECT 
                id,
                employee_name,
                clock_out as time,
                'out' as type,
                line_picture,
                created_at
            FROM time_records
            WHERE clock_out LIKE ? AND clock_out IS NOT NULL
            ORDER BY id DESC
        `).all(`${targetDate}%`);

        // รวมและ sort ตามเวลา (ใหม่สุดก่อน)
        const allRecords = [...clockInRecords, ...clockOutRecords]
            .map(record => ({
                id: record.id,
                employee: record.employee_name,
                time: record.time,
                type: record.type,
                linePicture: record.line_picture || '',
                // แยกเวลาจาก DD/MM/YYYY HH:mm:ss
                timeOnly: record.time ? record.time.split(' ')[1] || record.time : '',
                dateOnly: record.time ? record.time.split(' ')[0] || targetDate : targetDate
            }))
            .sort((a, b) => {
                // Sort by time descending (newest first)
                const timeA = a.time || '';
                const timeB = b.time || '';
                return timeB.localeCompare(timeA);
            })
            .slice(0, limit);

        return allRecords;
    }

    /**
     * ดึงสรุปสถานะวันนี้
     * @param {string} date - วันที่ต้องการดู (DD/MM/YYYY) หรือ null สำหรับวันนี้
     * @returns {Object} สรุปจำนวน
     */
    getTodaySummary(date = null) {
        const targetDate = date || moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY');

        // จำนวนพนักงานทั้งหมด
        const totalEmployees = this.db.prepare('SELECT COUNT(*) as count FROM employees').get().count;

        // จำนวนคนที่มาวันนี้ (มี clock_in)
        const presentToday = this.db.prepare(`
            SELECT COUNT(DISTINCT employee_name) as count 
            FROM time_records 
            WHERE clock_in LIKE ?
        `).get(`${targetDate}%`).count;

        // จำนวนคนที่กำลังทำงาน (ยังไม่ clock out)
        const workingNow = this.db.prepare('SELECT COUNT(*) as count FROM on_work').get().count;

        // จำนวนคนที่ clock out แล้ววันนี้
        const clockedOut = this.db.prepare(`
            SELECT COUNT(DISTINCT employee_name) as count 
            FROM time_records 
            WHERE clock_in LIKE ? AND clock_out IS NOT NULL
        `).get(`${targetDate}%`).count;

        // จำนวนคนสาย (เข้าหลัง 08:30)
        const lateCount = this.db.prepare(`
            SELECT COUNT(DISTINCT employee_name) as count 
            FROM time_records 
            WHERE clock_in LIKE ? 
            AND CAST(SUBSTR(clock_in, 12, 2) AS INTEGER) * 60 + CAST(SUBSTR(clock_in, 15, 2) AS INTEGER) > 510
        `).get(`${targetDate}%`).count; // 510 = 8*60 + 30 = 08:30

        return {
            date: targetDate,
            total: totalEmployees,
            present: presentToday,
            absent: Math.max(0, totalEmployees - presentToday),
            working: workingNow,
            clockedOut: clockedOut,
            late: lateCount
        };
    }

    /**
     * ดึงรายการลงเวลาตามชื่อพนักงาน
     * @param {string} name - ชื่อพนักงาน (หรือส่วนหนึ่งของชื่อ)
     * @param {string} date - วันที่ต้องการดู (DD/MM/YYYY) หรือ null สำหรับวันนี้
     * @param {number} limit - จำนวนรายการ
     * @returns {Array} รายการลงเวลา
     */
    getActivityByName(name, date = null, limit = 50) {
        const targetDate = date || moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY');
        const searchName = `%${name}%`;

        const records = this.db.prepare(`
            SELECT 
                id,
                employee_name,
                clock_in,
                clock_out,
                line_picture,
                working_hours,
                created_at
            FROM time_records
            WHERE employee_name LIKE ?
            AND clock_in LIKE ?
            ORDER BY id DESC
            LIMIT ?
        `).all(searchName, `${targetDate}%`, limit);

        return records.map(record => ({
            id: record.id,
            employee: record.employee_name,
            clockIn: record.clock_in,
            clockOut: record.clock_out,
            linePicture: record.line_picture || '',
            workingHours: record.working_hours,
            // แยกเวลา
            clockInTime: record.clock_in ? record.clock_in.split(' ')[1] || '' : '',
            clockOutTime: record.clock_out ? record.clock_out.split(' ')[1] || '' : ''
        }));
    }

    // ========== 🆕 Personal Dashboard Functions ==========

    // ผูก LINE User ID กับพนักงาน
    linkLineUserId(employeeName, lineUserId, lineName, linePicture) {
        try {
            // ตรวจสอบว่า line_user_id นี้ถูกใช้ไปแล้วหรือยัง
            const existingByLine = this.db.prepare(
                'SELECT name FROM employees WHERE line_user_id = ?'
            ).get(lineUserId);

            if (existingByLine) {
                return {
                    success: false,
                    error: `บัญชี LINE นี้ถูกผูกกับ "${existingByLine.name}" แล้ว`
                };
            }

            // ตรวจสอบว่าพนักงานคนนี้ผูก LINE ไปแล้วหรือยัง
            const existingByName = this.db.prepare(
                'SELECT line_user_id FROM employees WHERE name = ?'
            ).get(employeeName);

            if (existingByName && existingByName.line_user_id) {
                return {
                    success: false,
                    error: `พนักงาน "${employeeName}" ผูกบัญชี LINE ไว้แล้ว`
                };
            }

            // ผูก LINE User ID
            const registeredAt = moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss');
            this.db.prepare(`
                UPDATE employees 
                SET line_user_id = ?, line_name = ?, line_picture = ?, registered_at = ?
                WHERE name = ?
            `).run(lineUserId, lineName, linePicture, registeredAt, employeeName);

            console.log(`✅ Linked LINE to employee: ${employeeName}`);
            return { success: true, message: 'ลงทะเบียนสำเร็จ' };
        } catch (error) {
            console.error('Error linking LINE:', error);
            return { success: false, error: error.message };
        }
    }

    // ดึงพนักงานจาก LINE User ID
    getEmployeeByLineUserId(lineUserId) {
        try {
            const employee = this.db.prepare(`
                SELECT id, name, line_user_id, line_name, line_picture, registered_at, created_at
                FROM employees WHERE line_user_id = ?
            `).get(lineUserId);

            return employee || null;
        } catch (error) {
            console.error('Error getting employee by LINE ID:', error);
            return null;
        }
    }

    // ดึงรายชื่อพนักงานที่ยังไม่ผูก LINE
    getUnlinkedEmployees() {
        try {
            const employees = this.db.prepare(`
                SELECT id, name, created_at
                FROM employees 
                WHERE line_user_id IS NULL OR line_user_id = ''
                ORDER BY name
            `).all();

            return employees;
        } catch (error) {
            console.error('Error getting unlinked employees:', error);
            return [];
        }
    }

    // ดึงสถิติส่วนตัว
    getPersonalStats(employeeName, month = null) {
        try {
            // ถ้าไม่ระบุเดือน ใช้เดือนปัจจุบัน
            const targetMonth = month || moment().tz(CONFIG.TIMEZONE).format('MM/YYYY');
            const [mm, yyyy] = targetMonth.split('/');

            // ดึงข้อมูลทั้งเดือน
            const records = this.db.prepare(`
                SELECT clock_in, clock_out, working_hours
                FROM time_records
                WHERE employee_name = ?
                AND clock_in LIKE ?
                ORDER BY id DESC
            `).all(employeeName, `%/${mm}/${yyyy}%`);

            // คำนวณสถิติ
            let workDays = 0;
            let lateDays = 0;
            let totalHours = 0;

            records.forEach(record => {
                workDays++;
                totalHours += record.working_hours || 0;

                // ตรวจสอบมาสาย (หลัง 08:30)
                if (record.clock_in) {
                    const timePart = record.clock_in.split(' ')[1];
                    if (timePart) {
                        const [hh, mi] = timePart.split(':').map(Number);
                        if (hh > 8 || (hh === 8 && mi > 30)) {
                            lateDays++;
                        }
                    }
                }
            });

            // คำนวณวันทำงานในเดือน (จันทร์-ศุกร์)
            const startOfMonth = moment(`01/${mm}/${yyyy}`, 'DD/MM/YYYY').tz(CONFIG.TIMEZONE);
            const endOfMonth = startOfMonth.clone().endOf('month');
            const today = moment().tz(CONFIG.TIMEZONE);
            const checkUntil = today.isBefore(endOfMonth) ? today : endOfMonth;

            let totalWorkDaysInMonth = 0;
            const current = startOfMonth.clone();
            while (current.isSameOrBefore(checkUntil, 'day')) {
                const dayOfWeek = current.day();
                if (dayOfWeek !== 0 && dayOfWeek !== 6) { // ไม่นับ เสาร์-อาทิตย์
                    totalWorkDaysInMonth++;
                }
                current.add(1, 'day');
            }

            const absentDays = Math.max(0, totalWorkDaysInMonth - workDays);

            return {
                employeeName,
                month: targetMonth,
                workDays,
                lateDays,
                absentDays,
                totalHours: Math.round(totalHours * 100) / 100,
                totalWorkDaysInMonth
            };
        } catch (error) {
            console.error('Error getting personal stats:', error);
            return null;
        }
    }

    // ดึงประวัติส่วนตัว
    getPersonalHistory(employeeName, limit = 30, month = null) {
        try {
            let query;
            let params;

            if (month) {
                // Filter by month (format: MM/YYYY)
                const [mm, yyyy] = month.split('/');
                const monthPattern = `%/${mm}/${yyyy}%`;

                query = `
                    SELECT 
                        id, clock_in, clock_out, working_hours,
                        location_in_name, location_out_name
                    FROM time_records
                    WHERE employee_name = ?
                    AND clock_in LIKE ?
                    ORDER BY id DESC
                    LIMIT ?
                `;
                params = [employeeName, monthPattern, limit];
            } else {
                query = `
                    SELECT 
                        id, clock_in, clock_out, working_hours,
                        location_in_name, location_out_name
                    FROM time_records
                    WHERE employee_name = ?
                    ORDER BY id DESC
                    LIMIT ?
                `;
                params = [employeeName, limit];
            }

            const records = this.db.prepare(query).all(...params);

            return records.map(r => ({
                id: r.id,
                clockIn: r.clock_in,
                clockOut: r.clock_out,
                workingHours: r.working_hours,
                locationIn: r.location_in_name || '',
                locationOut: r.location_out_name || '',
                // แยกเวลา
                clockInTime: r.clock_in ? r.clock_in.split(' ')[1] || '' : '',
                clockOutTime: r.clock_out ? r.clock_out.split(' ')[1] || '' : '',
                date: r.clock_in ? r.clock_in.split(' ')[0] || '' : ''
            }));
        } catch (error) {
            console.error('Error getting personal history:', error);
            return [];
        }
    }

    // ยกเลิกการผูก LINE (Admin only)
    unlinkLineUserId(employeeName) {
        try {
            this.db.prepare(`
                UPDATE employees 
                SET line_user_id = NULL, line_name = NULL, line_picture = NULL, registered_at = NULL
                WHERE name = ?
            `).run(employeeName);

            console.log(`✅ Unlinked LINE from employee: ${employeeName}`);
            return { success: true, message: `ยกเลิกการเชื่อมต่อ LINE ของ "${employeeName}" สำเร็จ` };
        } catch (error) {
            console.error('Error unlinking LINE:', error);
            return { success: false, error: error.message };
        }
    }

    // ดึงรายชื่อพนักงานที่ผูก LINE แล้ว (Admin)
    getAllLinkedEmployees() {
        try {
            const employees = this.db.prepare(`
                SELECT id, name, line_user_id, line_name, line_picture, registered_at
                FROM employees
                ORDER BY 
                    CASE WHEN line_user_id IS NOT NULL THEN 0 ELSE 1 END,
                    name
            `).all();

            return employees.map(e => ({
                id: e.id,
                name: e.name,
                lineUserId: e.line_user_id,
                lineName: e.line_name,
                linePicture: e.line_picture,
                registeredAt: e.registered_at,
                isLinked: !!e.line_user_id
            }));
        } catch (error) {
            console.error('Error getting linked employees:', error);
            return [];
        }
    }
}

module.exports = SQLiteService;
