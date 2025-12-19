// Google Apps Script - Time Tracker Map Service (Simplified)
// Deploy เป็น Web App: Execute as: Me, Access: Anyone

// ========== Configuration ==========
const CONFIG = {
  TELEGRAM: {
    BOT_TOKEN: "7610983723:AAEFXDbDlq5uTHeyID8Fc5XEmIUx-LT6rJM",
    CHAT_ID: "7809169283"
  },
  MAPS: {
    SIZE: 600,
    MAP_TYPE: Maps.StaticMap.Type.HYBRID,
    LANGUAGE: 'TH'
  }
};

const THAILAND_TIMEZONE = 'Asia/Bangkok';

// ========== Main Web App Handler ==========
function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const { action, data } = requestData;
    
    if (!action || !data) {
      throw new Error('Missing action or data');
    }
    
    switch (action) {
      case 'clockin':
        return handleClockIn(data);
      case 'clockout':
        return handleClockOut(data);
      default:
        throw new Error(`Unknown action: ${action}`);
    }
    
  } catch (error) {
    console.error('❌ Error:', error);
    return jsonResponse({ success: false, error: error.message });
  }
}

function doGet(e) {
  return jsonResponse({
    service: "Time Tracker Map Service",
    status: "running",
    timestamp: new Date().toISOString()
  });
}

// ========== Clock In Handler ==========
function handleClockIn(data) {
  try {
    const { employee, lat, lon, line_name, userinfo } = data;
    
    if (!employee || !lat || !lon) {
      throw new Error("Missing required fields");
    }

    const latitude = parseFloat(lat);
    const longitude = parseFloat(lon);
    const dateObj = parseTimestamp(data);
    
    const formattedDate = Utilities.formatDate(dateObj, THAILAND_TIMEZONE, "dd/MM/yyyy");
    const formattedTime = Utilities.formatDate(dateObj, THAILAND_TIMEZONE, "HH:mm:ss") + " น.";
    const location = getLocationAddress(latitude, longitude);
    
    const message =
      `⏱ ลงเวลาเข้างาน\n` +
      `👤 ชื่อ-นามสกุล: *${employee}*\n` +
      `📅 วันที่: *${formattedDate}*\n` +
      `🕒 เวลา: *${formattedTime}*\n` +
      `💬 ชื่อไลน์: *${line_name || 'ไม่ระบุ'}*\n` +
      (userinfo ? `📝 หมายเหตุ: *${userinfo}*\n` : "") +
      `📍 พิกัด: *${location}*\n` +
      `🗺 [📍 ดูตำแหน่งบนแผนที่](https://www.google.com/maps/place/${latitude},${longitude})`;

    const mapBlob = createMapImage(latitude, longitude);
    const telegramResult = sendMapToTelegram(mapBlob, message);
    
    return jsonResponse({
      success: telegramResult.success,
      message: telegramResult.success ? "Clock In sent" : "Failed",
      employee: employee
    });
    
  } catch (error) {
    console.error('❌ Clock In Error:', error);
    return jsonResponse({ success: false, error: error.message });
  }
}

// ========== Clock Out Handler ==========
function handleClockOut(data) {
  try {
    const { employee, lat, lon, line_name, hoursWorked } = data;
    
    if (!employee || !lat || !lon) {
      throw new Error("Missing required fields");
    }

    const latitude = parseFloat(lat);
    const longitude = parseFloat(lon);
    const dateObj = parseTimestamp(data);
    
    const formattedDate = Utilities.formatDate(dateObj, THAILAND_TIMEZONE, "dd/MM/yyyy");
    const formattedTime = Utilities.formatDate(dateObj, THAILAND_TIMEZONE, "HH:mm:ss") + " น.";
    const location = getLocationAddress(latitude, longitude);
    
    const message =
      `⏱ ลงเวลาออกงาน\n` +
      `👤 ชื่อ-นามสกุล: *${employee}*\n` +
      `📅 วันที่: *${formattedDate}*\n` +
      `🕒 เวลา: *${formattedTime}*\n` +
      `💬 ชื่อไลน์: *${line_name || 'ไม่ระบุ'}*\n` +
      `🕑 ชั่วโมงทำงาน: *${parseFloat(hoursWorked || 0).toFixed(2)} ชม.*\n` +
      `📍 พิกัด: *${location}*\n` +
      `🗺 [📍 ดูตำแหน่งบนแผนที่](https://www.google.com/maps/place/${latitude},${longitude})`;

    const mapBlob = createMapImage(latitude, longitude);
    const telegramResult = sendMapToTelegram(mapBlob, message);
    
    return jsonResponse({
      success: telegramResult.success,
      message: telegramResult.success ? "Clock Out sent" : "Failed",
      employee: employee
    });
    
  } catch (error) {
    console.error('❌ Clock Out Error:', error);
    return jsonResponse({ success: false, error: error.message });
  }
}

// ========== Helper Functions ==========
function parseTimestamp(data) {
  let dateObj;
  
  if (data.timestamp && data.timestamp.includes('/')) {
    const [datePart, timePart] = data.timestamp.split(' ');
    const [day, month, year] = datePart.split('/');
    dateObj = new Date(`${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}T${timePart}`);
  } else if (data.timestampMillis) {
    dateObj = new Date(data.timestampMillis);
  } else if (data.timestampISO) {
    dateObj = new Date(data.timestampISO);
  } else {
    dateObj = new Date();
  }
  
  return isNaN(dateObj.getTime()) ? new Date() : dateObj;
}

function createMapImage(lat, lon) {
  const map = Maps.newStaticMap()
    .setSize(CONFIG.MAPS.SIZE, CONFIG.MAPS.SIZE)
    .setLanguage(CONFIG.MAPS.LANGUAGE)
    .setMobile(true)
    .setMapType(CONFIG.MAPS.MAP_TYPE)
    .addMarker(lat, lon);
  return map.getBlob();
}

function getLocationAddress(lat, lon) {
  try {
    const response = Maps.newGeocoder()
      .setRegion('th')
      .setLanguage('th-TH')
      .reverseGeocode(lat, lon);
      
    if (response.results && response.results.length > 0) {
      return response.results[0].formatted_address;
    }
    return `${lat}, ${lon}`;
  } catch (error) {
    return `${lat}, ${lon}`;
  }
}

function sendMapToTelegram(mapBlob, caption) {
  try {
    const url = `https://api.telegram.org/bot${CONFIG.TELEGRAM.BOT_TOKEN}/sendPhoto`;
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      payload: {
        chat_id: CONFIG.TELEGRAM.CHAT_ID,
        photo: mapBlob,
        caption: caption,
        parse_mode: "Markdown"
      }
    });
    
    const result = JSON.parse(response.getContentText());
    return { success: result.ok };
    
  } catch (error) {
    console.error("❌ Telegram Error:", error);
    return { success: false, error: error.message };
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== ทดสอบระบบ ==========
function testClockIn() {
  const testData = {
    employee: 'ทดสอบระบบ',
    lat: 14.0583,
    lon: 100.6014,
    line_name: 'ผู้ทดสอบ',
    userinfo: 'ทดสอบจาก GSA',
    timestamp: new Date().toISOString()
  };
  
  return handleClockIn(testData);
}