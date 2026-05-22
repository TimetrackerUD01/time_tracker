// services/excelExport.js - Excel Export Service
const ExcelJS = require('exceljs');
const moment = require('moment-timezone');
const { CONFIG } = require('../config');

class ExcelExportService {
  static async createWorkbook(data, type, params) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('รายงานการลงเวลา');

    // ตั้งค่าข้อมูลองค์กร
    const orgInfo = {
      name: 'องค์การบริหารส่วนตำบลข่าใหญ่',
      address: 'อำเภอเมือง จังหวัดหนองบัวลำภู',
      phone: '042-315962'
    };

    // สร้างหัวข้อรายงาน
    let reportTitle = '';
    let reportPeriod = '';

    switch (type) {
      case 'daily':
        reportTitle = 'รายงานการลงเวลาเข้า-ออกงาน รายวัน';
        reportPeriod = `วันที่ ${moment(params.date).tz(CONFIG.TIMEZONE).format('DD MMMM YYYY')}`;
        break;
      case 'monthly':
        const monthNames = [
          'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
          'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
        ];
        const isDetailed = params.format === 'detailed';
        reportTitle = isDetailed
          ? 'รายงานการลงเวลาเข้า-ออกงาน รายเดือน (แบ่งตามวันชัดเจน)'
          : 'รายงานการลงเวลาเข้า-ออกงาน รายเดือน';
        reportPeriod = `เดือน ${monthNames[params.month - 1]} ${parseInt(params.year) + 543}`;
        break;
      case 'range':
        reportTitle = 'รายงานการลงเวลาเข้า-ออกงาน ช่วงวันที่';
        const startDate = moment(params.startDate).tz(CONFIG.TIMEZONE);
        const endDate = moment(params.endDate).tz(CONFIG.TIMEZONE);
        reportPeriod = `${startDate.format('DD MMMM YYYY')} - ${endDate.format('DD MMMM YYYY')}`;
        break;
    }

    // จัดรูปแบบหัวกระดาษ
    worksheet.mergeCells('A1:J3');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = `${orgInfo.name}\n${reportTitle}\n${reportPeriod}`;
    titleCell.font = { name: 'Angsana New', size: 18, bold: true };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    // ข้อมูลองค์กร
    worksheet.getCell('A4').value = `${orgInfo.address} โทร. ${orgInfo.phone}`;
    worksheet.getCell('A4').font = { name: 'Angsana New', size: 14 };
    worksheet.getCell('A4').alignment = { horizontal: 'center' };
    worksheet.mergeCells('A4:J4');

    // สร้างหัวตาราง
    const headerRow = 6;
    const headers = [
      'ลำดับ',
      'ชื่อ-นามสกุล',
      'วันที่',
      'เวลาเข้า',
      'เวลาออก',
      'ชั่วโมงทำงาน',
      'หมายเหตุ',
      'สถานที่เข้า',
      'สถานที่ออก',
      'ชื่อไลน์'
    ];

    headers.forEach((header, index) => {
      const cell = worksheet.getCell(headerRow, index + 1);
      cell.value = header;
      cell.font = { name: 'Angsana New', size: 14, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE6E6FA' }
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    // เพิ่มข้อมูล
    if (type === 'monthly' && params.format === 'detailed') {
      // สำหรับรายงานรายเดือนแบบ detailed: จัดเรียงข้อมูลตามวันที่
      data = ExcelExportService.organizeDetailedMonthlyData(data, params);
    }

    data.forEach((record, index) => {
      const rowNumber = headerRow + 1 + index;

      // จัดการวันที่และเวลา
      let clockInDate = null;
      let clockOutDate = null;
      let dateDisplay = '';
      let clockInTime = '';
      let clockOutTime = '';

      if (record.clockIn) {
        try {
          if (typeof record.clockIn === 'string' && record.clockIn.includes(' ')) {
            // ตรวจสอบรูปแบบ DD/MM/YYYY HH:mm:ss ก่อน
            if (record.clockIn.match(/^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}$/)) {
              clockInDate = moment.tz(record.clockIn, 'DD/MM/YYYY HH:mm:ss', CONFIG.TIMEZONE);
              console.log(`📅 Parsed DD/MM/YYYY format: ${record.clockIn} -> ${clockInDate.format('YYYY-MM-DD HH:mm:ss')}`);
            }
            // รูปแบบ YYYY-MM-DD HH:mm:ss
            else if (record.clockIn.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
              clockInDate = moment.tz(record.clockIn, 'YYYY-MM-DD HH:mm:ss', CONFIG.TIMEZONE);
              console.log(`📅 Parsed YYYY-MM-DD format: ${record.clockIn} -> ${clockInDate.format('YYYY-MM-DD HH:mm:ss')}`);
            }
            else {
              // ลองให้ moment แปลงเอง
              clockInDate = moment(record.clockIn).tz(CONFIG.TIMEZONE);
              console.log(`📅 Auto-parsed format: ${record.clockIn} -> ${clockInDate.format('YYYY-MM-DD HH:mm:ss')}`);
            }
          } else {
            clockInDate = moment(record.clockIn).tz(CONFIG.TIMEZONE);
            console.log(`📅 Parsed non-string format: ${record.clockIn} -> ${clockInDate.format('YYYY-MM-DD HH:mm:ss')}`);
          }

          if (clockInDate.isValid()) {
            dateDisplay = clockInDate.format('DD/MM/YYYY');
            clockInTime = clockInDate.format('HH:mm:ss');
            console.log(`✅ Final display: Date="${dateDisplay}", Time="${clockInTime}"`);
          } else {
            console.warn(`⚠️ Invalid clockIn date: "${record.clockIn}"`);
          }
        } catch (error) {
          console.warn('Error parsing clockIn time:', record.clockIn, error);
        }
      }

      if (record.clockOut) {
        try {
          if (typeof record.clockOut === 'string' && record.clockOut.includes(' ')) {
            // ตรวจสอบรูปแบบ DD/MM/YYYY HH:mm:ss ก่อน
            if (record.clockOut.match(/^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}$/)) {
              clockOutDate = moment.tz(record.clockOut, 'DD/MM/YYYY HH:mm:ss', CONFIG.TIMEZONE);
              console.log(`📅 Parsed clockOut DD/MM/YYYY format: ${record.clockOut} -> ${clockOutDate.format('YYYY-MM-DD HH:mm:ss')}`);
            }
            // รูปแบบ YYYY-MM-DD HH:mm:ss
            else if (record.clockOut.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
              clockOutDate = moment.tz(record.clockOut, 'YYYY-MM-DD HH:mm:ss', CONFIG.TIMEZONE);
              console.log(`📅 Parsed clockOut YYYY-MM-DD format: ${record.clockOut} -> ${clockOutDate.format('YYYY-MM-DD HH:mm:ss')}`);
            }
            else {
              // ลองให้ moment แปลงเอง
              clockOutDate = moment(record.clockOut).tz(CONFIG.TIMEZONE);
              console.log(`📅 Auto-parsed clockOut format: ${record.clockOut} -> ${clockOutDate.format('YYYY-MM-DD HH:mm:ss')}`);
            }
          } else {
            clockOutDate = moment(record.clockOut).tz(CONFIG.TIMEZONE);
            console.log(`📅 Parsed clockOut non-string format: ${record.clockOut} -> ${clockOutDate.format('YYYY-MM-DD HH:mm:ss')}`);
          }

          if (clockOutDate.isValid()) {
            clockOutTime = clockOutDate.format('HH:mm:ss');
            console.log(`✅ Final clockOut time: "${clockOutTime}"`);
          } else {
            console.warn(`⚠️ Invalid clockOut date: "${record.clockOut}"`);
          }
        } catch (error) {
          console.warn('Error parsing clockOut time:', record.clockOut, error);
        }
      }

      // จัดการชั่วโมงทำงาน
      let workingHoursDisplay = '';
      if (record.workingHours) {
        const hours = parseFloat(record.workingHours);
        if (!isNaN(hours)) {
          workingHoursDisplay = `${hours.toFixed(2)} ชม.`;
        } else {
          workingHoursDisplay = record.workingHours;
        }
      }

      const rowData = [
        record.no || (index + 1),
        record.employee || '',
        dateDisplay,
        clockInTime,
        clockOutTime,
        workingHoursDisplay,
        record.note || '',
        record.locationIn || '',
        record.locationOut || '',
        record.lineName || ''
      ];

      rowData.forEach((value, colIndex) => {
        const cell = worksheet.getCell(rowNumber, colIndex + 1);
        cell.value = value;
        cell.font = { name: 'Angsana New', size: 12 };
        cell.alignment = {
          horizontal: colIndex === 0 ? 'center' : 'left',
          vertical: 'middle'
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };

        // สีพื้นหลังสำหรับแถวต่างๆ (ตรวจสอบหมายเหตุจากคอลัมน์ E)
        if (record.note && record.note.includes('ลืมลงเวลาออก')) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFCCCC' } // สีแดงอ่อน
          };
        }
      });
    });

    // ปรับขนาดคอลัมน์
    const columnWidths = [8, 25, 15, 12, 12, 15, 25, 30, 30, 20];
    columnWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width;
    });

    // สรุปข้อมูล
    const summaryRow = headerRow + data.length + 2;

    // สถิติการทำงาน
    const totalRecords = data.length;
    const normalCheckouts = data.filter(r => !r.note || !r.note.includes('ลืมลงเวลาออก')).length;
    const missedCheckouts = data.filter(r => r.note && r.note.includes('ลืมลงเวลาออก')).length;

    worksheet.getCell(summaryRow, 1).value = `สรุปข้อมูล: ทั้งหมด ${totalRecords} รายการ | ลงเวลาออกปกติ ${normalCheckouts} คน | ลืมลงเวลาออก ${missedCheckouts} คน`;
    worksheet.getCell(summaryRow, 1).font = { name: 'Angsana New', size: 12, bold: true };
    worksheet.mergeCells(`A${summaryRow}:J${summaryRow}`);

    // วันที่สร้างรายงาน
    const footerRow = summaryRow + 2;
    const currentTime = moment().tz(CONFIG.TIMEZONE);
    worksheet.getCell(footerRow, 1).value = `สร้างรายงานเมื่อ: ${currentTime.format('DD/MM/YYYY HH:mm:ss')} (เวลาไทย)`;
    worksheet.getCell(footerRow, 1).font = { name: 'Angsana New', size: 10 };
    worksheet.getCell(footerRow, 1).alignment = { horizontal: 'right' };
    worksheet.mergeCells(`A${footerRow}:J${footerRow}`);

    // เพิ่มหมายเหตุเกี่ยวกับสี
    if (data.some(r => r.note && r.note.includes('ลืมลงเวลาออก'))) {
      const noteRow = footerRow + 1;
      worksheet.getCell(noteRow, 1).value = 'หมายเหตุ: แถวที่มีพื้นหลังสีแดงอ่อน = ลืมลงเวลาออก (ระบบอัตโนมัติ)';
      worksheet.getCell(noteRow, 1).font = { name: 'Angsana New', size: 10, italic: true };
      worksheet.mergeCells(`A${noteRow}:J${noteRow}`);
    }

    return workbook;
  }

  // ฟังก์ชันสำหรับจัดเรียงข้อมูลรายเดือนแบบ detailed
  static organizeDetailedMonthlyData(data, params) {
    console.log(`📊 Organizing detailed monthly data: ${data.length} records`);

    // จัดเรียงข้อมูลตามวันที่ และ ชื่อพนักงาน
    const sortedData = data.sort((a, b) => {
      // เรียงตามวันที่ก่อน
      const dateA = moment(a.clockIn).tz(CONFIG.TIMEZONE);
      const dateB = moment(b.clockIn).tz(CONFIG.TIMEZONE);

      if (dateA.format('YYYY-MM-DD') !== dateB.format('YYYY-MM-DD')) {
        return dateA.diff(dateB);
      }

      // ถ้าวันที่เดียวกัน เรียงตามชื่อพนักงาน
      return (a.employee || '').localeCompare(b.employee || '', 'th');
    });

    console.log(`✅ Sorted detailed data: ${sortedData.length} records`);
    return sortedData;
  }

  // 🆕 ฟังก์ชันสร้างรายงานรายเดือนแบบแยกวัน + สรุปสาย/ขาด
  static async createDailySummaryWorkbook(data, params, allEmployees = []) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('รายงานรายเดือน (แยกวัน)');

    // ข้อมูลองค์กร
    const orgInfo = {
      name: 'องค์การบริหารส่วนตำบลข่าใหญ่',
      address: 'อำเภอเมือง จังหวัดหนองบัวลำภู',
      phone: '042-315962'
    };

    const monthNames = [
      'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
      'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
    ];

    const thaiDays = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
    const month = parseInt(params.month, 10);
    let year = parseInt(params.year, 10);
    // 🆕 แปลง พ.ศ. → ค.ศ. อัตโนมัติ (ถ้า year > 2500 แสดงว่าเป็น พ.ศ.)
    if (year > 2500) {
      year = year - 543;
    }
    const thaiYear = year + 543;

    // หัวรายงาน
    worksheet.mergeCells('A1:J3');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = `${orgInfo.name}\nรายงานการลงเวลาเข้า-ออกงาน รายเดือน (แยกตามวัน)\nเดือน ${monthNames[month - 1]} ${thaiYear}`;
    titleCell.font = { name: 'Angsana New', size: 18, bold: true };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    worksheet.getCell('A4').value = `${orgInfo.address} โทร. ${orgInfo.phone}`;
    worksheet.getCell('A4').font = { name: 'Angsana New', size: 14 };
    worksheet.getCell('A4').alignment = { horizontal: 'center' };
    worksheet.mergeCells('A4:J4');

    // Helper: parse clock time
    const parseClockToMoment = (value) => {
      if (!value) return null;
      if (typeof value === 'string') {
        // รองรับรูปแบบ D/MM/YYYY H:mm:ss (ไม่มี leading zero)
        if (value.match(/^\d{1,2}\/\d{2}\/\d{4} \d{1,2}:\d{2}:\d{2}$/)) {
          return moment.tz(value, 'D/MM/YYYY H:mm:ss', CONFIG.TIMEZONE);
        }
        if (value.match(/^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}$/)) {
          return moment.tz(value, 'DD/MM/YYYY HH:mm:ss', CONFIG.TIMEZONE);
        }
        if (value.match(/^\d{4}-\d{2}-\d{2}/)) {
          return moment.tz(value, 'YYYY-MM-DD HH:mm:ss', CONFIG.TIMEZONE);
        }
      }
      const fallback = moment(value).tz(CONFIG.TIMEZONE);
      return fallback.isValid() ? fallback : null;
    };

    // Helper: ตรวจสอบมาสาย (หลัง 08:30)
    const isLate = (clockInMoment) => {
      if (!clockInMoment || !clockInMoment.isValid()) return false;
      const hour = clockInMoment.hour();
      const minute = clockInMoment.minute();
      // กะกลางคืน (18:00-06:00) ไม่นับสาย
      if (hour >= 18 || hour < 6) return false;
      // สายถ้าหลัง 08:30
      return (hour > 8) || (hour === 8 && minute > 30);
    };

    // Helper: คำนวณนาทีที่สาย
    const calculateLateMinutes = (clockInMoment) => {
      if (!clockInMoment || !clockInMoment.isValid()) return 0;
      const threshold = clockInMoment.clone().hour(8).minute(30).second(0);
      const diff = clockInMoment.diff(threshold, 'minutes');
      return diff > 0 ? diff : 0;
    };

    // จัดกลุ่มข้อมูลตามวันที่ + Deduplicate (1 คน = 1 รายการ/วัน เอาเวลาเข้าเร็วสุด)
    const dataByDate = {};
    const tempByDateEmployee = {}; // เก็บ record แยกตามวัน+ชื่อ

    data.forEach(record => {
      const clockMoment = parseClockToMoment(record.clockIn);
      if (clockMoment && clockMoment.isValid()) {
        const dateKey = clockMoment.format('YYYY-MM-DD');
        const employeeName = record.employee || '';
        const uniqueKey = `${dateKey}|${employeeName}`;

        if (!tempByDateEmployee[uniqueKey]) {
          // ยังไม่มี record ของคนนี้ในวันนี้
          tempByDateEmployee[uniqueKey] = {
            ...record,
            clockInMoment: clockMoment,
            clockOutMoment: parseClockToMoment(record.clockOut)
          };
        } else {
          // มีแล้ว เปรียบเทียบว่าอันไหนเข้าเร็วกว่า
          const existingMoment = tempByDateEmployee[uniqueKey].clockInMoment;
          if (clockMoment.isBefore(existingMoment)) {
            // อันใหม่เข้าเร็วกว่า ใช้อันใหม่แทน
            tempByDateEmployee[uniqueKey] = {
              ...record,
              clockInMoment: clockMoment,
              clockOutMoment: parseClockToMoment(record.clockOut)
            };
          }
        }
      }
    });

    // จัดกลุ่มกลับเป็น dataByDate
    Object.entries(tempByDateEmployee).forEach(([key, record]) => {
      const dateKey = key.split('|')[0];
      if (!dataByDate[dateKey]) {
        dataByDate[dateKey] = [];
      }
      dataByDate[dateKey].push(record);
    });

    console.log(`🔍 Records after dedup: ${Object.values(dataByDate).flat().length} (from ${data.length} original)`);

    // หาจำนวนวันในเดือน
    const daysInMonth = moment({ year, month: month - 1 }).daysInMonth();
    let currentRow = 6;

    // สถิติรวมทั้งเดือน
    let totalPresent = 0;
    let totalLate = 0;
    let totalAbsent = 0;
    let totalMissedCheckout = 0;

    // วนลูปแต่ละวันในเดือน
    for (let day = 1; day <= daysInMonth; day++) {
      const dateKey = moment({ year, month: month - 1, day }).format('YYYY-MM-DD');
      const dateMoment = moment({ year, month: month - 1, day });
      const dayOfWeek = dateMoment.day(); // 0=Sun, 6=Sat
      const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);
      const thaiDayName = thaiDays[dayOfWeek];

      const dayRecords = dataByDate[dateKey] || [];

      // ถ้าเป็นอนาคต ไม่แสดง
      if (dateMoment.isAfter(moment().tz(CONFIG.TIMEZONE), 'day')) {
        continue;
      }

      // หัววันที่
      worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
      const dayHeaderCell = worksheet.getCell(`A${currentRow}`);
      dayHeaderCell.value = `📅 วันที่ ${day} ${monthNames[month - 1]} ${thaiYear} (วัน${thaiDayName})${isWeekend ? ' - วันหยุด' : ''}`;
      dayHeaderCell.font = { name: 'Angsana New', size: 14, bold: true };
      dayHeaderCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: isWeekend ? 'FFFFD700' : 'FF4472C4' }
      };
      dayHeaderCell.font = { name: 'Angsana New', size: 14, bold: true, color: { argb: isWeekend ? 'FF000000' : 'FFFFFFFF' } };
      dayHeaderCell.alignment = { horizontal: 'left', vertical: 'middle' };
      currentRow++;

      // ถ้าเป็นวันหยุดและไม่มีข้อมูล
      if (isWeekend && dayRecords.length === 0) {
        worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
        worksheet.getCell(`A${currentRow}`).value = '   (วันหยุด - ไม่มีข้อมูลลงเวลา)';
        worksheet.getCell(`A${currentRow}`).font = { name: 'Angsana New', size: 12, italic: true };
        currentRow += 2;
        continue;
      }

      // หัวตาราง
      const headers = ['ลำดับ', 'ชื่อ-นามสกุล', 'เวลาเข้า', 'เวลาออก', 'ชั่วโมง', 'สถานะ', 'หมายเหตุ', 'สถานที่เข้า', 'สถานที่ออก', 'ชื่อไลน์'];
      headers.forEach((header, index) => {
        const cell = worksheet.getCell(currentRow, index + 1);
        cell.value = header;
        cell.font = { name: 'Angsana New', size: 12, bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6E6FA' } };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' },
          bottom: { style: 'thin' }, right: { style: 'thin' }
        };
      });
      currentRow++;

      // ข้อมูลพนักงานในวันนั้น
      const presentEmployees = [];
      const lateEmployees = [];
      const missedCheckoutEmployees = [];

      dayRecords.sort((a, b) => (a.employee || '').localeCompare(b.employee || '', 'th'));

      dayRecords.forEach((record, index) => {
        const clockInTime = record.clockInMoment ? record.clockInMoment.format('HH:mm:ss') : '';
        const clockOutTime = record.clockOutMoment ? record.clockOutMoment.format('HH:mm:ss') : '';

        const late = isLate(record.clockInMoment);
        const lateMinutes = calculateLateMinutes(record.clockInMoment);
        const missedCheckout = record.note && record.note.includes('ลืมลงเวลาออก');

        let status = '✅ ปกติ';
        if (late) {
          status = `⏰ สาย ${lateMinutes} นาที`;
          lateEmployees.push(record.employee);
        }
        if (missedCheckout) {
          status = '📝 ลืมลงออก';
          missedCheckoutEmployees.push(record.employee);
        }

        presentEmployees.push(record.employee);

        // ชั่วโมงทำงาน
        let workingHoursDisplay = '';
        if (record.workingHours) {
          const hours = parseFloat(record.workingHours);
          if (!isNaN(hours)) {
            workingHoursDisplay = `${hours.toFixed(2)}`;
          }
        }

        const rowData = [
          index + 1,
          record.employee || '',
          clockInTime,
          clockOutTime,
          workingHoursDisplay,
          status,
          record.note || '',
          record.locationIn || '',
          record.locationOut || '',
          record.lineName || ''
        ];

        rowData.forEach((value, colIndex) => {
          const cell = worksheet.getCell(currentRow, colIndex + 1);
          cell.value = value;
          cell.font = { name: 'Angsana New', size: 11 };
          cell.alignment = { horizontal: colIndex === 0 ? 'center' : 'left', vertical: 'middle' };
          cell.border = {
            top: { style: 'thin' }, left: { style: 'thin' },
            bottom: { style: 'thin' }, right: { style: 'thin' }
          };

          // สีพื้นหลัง
          if (late) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFE0' } }; // สีเหลืองอ่อน
          }
          if (missedCheckout) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCCCC' } }; // สีแดงอ่อน
          }
        });
        currentRow++;
      });

      // หาคนขาด (ไม่นับวันหยุด)
      const absentEmployees = [];
      if (!isWeekend && allEmployees.length > 0) {
        allEmployees.forEach(emp => {
          if (!presentEmployees.includes(emp)) {
            absentEmployees.push(emp);
          }
        });
      }

      // สรุปวัน
      currentRow++;
      worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
      const summaryCell = worksheet.getCell(`A${currentRow}`);

      let summaryText = `📊 สรุปวันที่ ${day}: `;
      summaryText += `✅ มาทำงาน ${presentEmployees.length} คน`;

      if (lateEmployees.length > 0) {
        summaryText += ` | ⏰ มาสาย ${lateEmployees.length} คน`;
      }

      if (!isWeekend && absentEmployees.length > 0) {
        summaryText += ` | ❌ ขาด ${absentEmployees.length} คน`;
      }

      if (missedCheckoutEmployees.length > 0) {
        summaryText += ` | 📝 ลืมลงออก ${missedCheckoutEmployees.length} คน`;
      }

      summaryCell.value = summaryText;
      summaryCell.font = { name: 'Angsana New', size: 12, bold: true };
      summaryCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
      currentRow++;

      // แสดงรายชื่อคนสาย (ถ้ามี)
      if (lateEmployees.length > 0) {
        worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
        worksheet.getCell(`A${currentRow}`).value = `   ⏰ คนสาย: ${lateEmployees.join(', ')}`;
        worksheet.getCell(`A${currentRow}`).font = { name: 'Angsana New', size: 11, color: { argb: 'FFFF6600' } };
        currentRow++;
      }

      // แสดงรายชื่อคนขาด (ถ้ามี และไม่ใช่วันหยุด)
      if (!isWeekend && absentEmployees.length > 0 && absentEmployees.length <= 20) {
        worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
        worksheet.getCell(`A${currentRow}`).value = `   ❌ คนขาด: ${absentEmployees.join(', ')}`;
        worksheet.getCell(`A${currentRow}`).font = { name: 'Angsana New', size: 11, color: { argb: 'FFCC0000' } };
        currentRow++;
      } else if (!isWeekend && absentEmployees.length > 20) {
        worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
        worksheet.getCell(`A${currentRow}`).value = `   ❌ คนขาด: ${absentEmployees.length} คน (มากเกินกว่าจะแสดงรายชื่อ)`;
        worksheet.getCell(`A${currentRow}`).font = { name: 'Angsana New', size: 11, color: { argb: 'FFCC0000' } };
        currentRow++;
      }

      currentRow++; // เว้นบรรทัด

      // สะสมสถิติ
      totalPresent += presentEmployees.length;
      totalLate += lateEmployees.length;
      totalAbsent += absentEmployees.length;
      totalMissedCheckout += missedCheckoutEmployees.length;
    }

    // สรุปรวมทั้งเดือน
    currentRow++;
    worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
    const monthSummaryHeader = worksheet.getCell(`A${currentRow}`);
    monthSummaryHeader.value = `📈 สรุปรวมเดือน ${monthNames[month - 1]} ${thaiYear}`;
    monthSummaryHeader.font = { name: 'Angsana New', size: 16, bold: true };
    monthSummaryHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D32' } };
    monthSummaryHeader.font = { name: 'Angsana New', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    currentRow++;

    worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = `   ✅ การลงเวลาทั้งหมด: ${totalPresent} ครั้ง | ⏰ มาสายรวม: ${totalLate} ครั้ง | ❌ ขาดรวม: ${totalAbsent} ครั้ง | 📝 ลืมลงออกรวม: ${totalMissedCheckout} ครั้ง`;
    worksheet.getCell(`A${currentRow}`).font = { name: 'Angsana New', size: 14 };
    currentRow++;

    if (allEmployees.length > 0) {
      worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
      worksheet.getCell(`A${currentRow}`).value = `   👥 พนักงานทั้งหมดในระบบ: ${allEmployees.length} คน`;
      worksheet.getCell(`A${currentRow}`).font = { name: 'Angsana New', size: 14 };
      currentRow++;
    }

    // วันที่สร้างรายงาน
    currentRow++;
    worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = `สร้างรายงานเมื่อ: ${moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss')} (เวลาไทย)`;
    worksheet.getCell(`A${currentRow}`).font = { name: 'Angsana New', size: 10 };
    worksheet.getCell(`A${currentRow}`).alignment = { horizontal: 'right' };

    // ปรับขนาดคอลัมน์
    const columnWidths = [8, 25, 12, 12, 10, 18, 25, 25, 25, 15];
    columnWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width;
    });

    return workbook;
  }
}

module.exports = ExcelExportService;
