// server.js - Time Tracker with Admin Panel and Excel Export
const express = require('express');
const cors = require('cors');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const path = require('path');
const fs = require('fs');
const cron = require('node-cron');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const ExcelJS = require('exceljs');
const moment = require('moment-timezone');
const fetch = require('node-fetch');
const { CONFIG, validateConfig } = require('./config');
const ExcelExportService = require('./services/excelExport');
// ๐• SQLite + Sync Services
const SQLiteService = require('./services/sqliteService');
const SyncService = require('./services/syncService');

const app = express();
const PORT = process.env.PORT || 3001;
console.log(`๐”ง Using PORT: ${PORT}`);

// ========== Helper Functions ==========

/**
 * เธเธณเธเธงเธ“เธเธฑเนเธงเนเธกเธเธเธฒเธฃเธ—เธณเธเธฒเธเนเธเธเน€เธ”เธตเธขเธงเธเธฑเธเธเธฑเธ admin stats
 * @param {string} clockInTime - เน€เธงเธฅเธฒเน€เธเนเธฒเธเธฒเธ
 * @param {string} [clockOutTime] - เน€เธงเธฅเธฒเธญเธญเธเธเธฒเธ (เธ–เนเธฒเนเธกเนเนเธซเนเธเธฐเนเธเนเน€เธงเธฅเธฒเธเธฑเธเธเธธเธเธฑเธ)
 * @returns {number} - เธเธฑเนเธงเนเธกเธเธเธฒเธฃเธ—เธณเธเธฒเธ (เธ—เธจเธเธดเธขเธก)
 */
function calculateWorkingHours(clockInTime, clockOutTime = null) {
  if (!clockInTime) {
    console.warn('โ ๏ธ No clock in time provided for calculation');
    return 0;
  }

  try {
    // ๐”ง เนเธเนเนเธ: เธฃเธญเธเธฃเธฑเธเธ—เธฑเนเธ DD/MM/YYYY เนเธฅเธฐ YYYY-MM-DD format
    let clockInMoment;

    // เธ•เธฃเธงเธเธชเธญเธ format เธเธญเธ clockInTime
    if (typeof clockInTime === 'string') {
      if (clockInTime.match(/^\d{2}\/\d{2}\/\d{4}/)) {
        // เธฃเธนเธเนเธเธ DD/MM/YYYY HH:mm:ss
        clockInMoment = moment.tz(clockInTime, 'DD/MM/YYYY HH:mm:ss', CONFIG.TIMEZONE);
      } else if (clockInTime.match(/^\d{4}-\d{2}-\d{2}/)) {
        // เธฃเธนเธเนเธเธ YYYY-MM-DD HH:mm:ss
        clockInMoment = moment.tz(clockInTime, 'YYYY-MM-DD HH:mm:ss', CONFIG.TIMEZONE);
      } else {
        // เธฅเธญเธเนเธเน auto parse
        clockInMoment = moment.tz(clockInTime, CONFIG.TIMEZONE);
      }
    } else {
      // เธ–เนเธฒเน€เธเนเธ Date object เธซเธฃเธทเธญ timestamp
      clockInMoment = moment.tz(clockInTime, CONFIG.TIMEZONE);
    }

    // เธ—เธณเน€เธเนเธเน€เธ”เธตเธขเธงเธเธฑเธเธเธฑเธ clockOutTime
    let endTimeMoment;
    if (clockOutTime) {
      if (typeof clockOutTime === 'string') {
        if (clockOutTime.match(/^\d{2}\/\d{2}\/\d{4}/)) {
          // เธฃเธนเธเนเธเธ DD/MM/YYYY HH:mm:ss
          endTimeMoment = moment.tz(clockOutTime, 'DD/MM/YYYY HH:mm:ss', CONFIG.TIMEZONE);
        } else if (clockOutTime.match(/^\d{4}-\d{2}-\d{2}/)) {
          // เธฃเธนเธเนเธเธ YYYY-MM-DD HH:mm:ss
          endTimeMoment = moment.tz(clockOutTime, 'YYYY-MM-DD HH:mm:ss', CONFIG.TIMEZONE);
        } else {
          // เธฅเธญเธเนเธเน auto parse
          endTimeMoment = moment.tz(clockOutTime, CONFIG.TIMEZONE);
        }
      } else {
        // เธ–เนเธฒเน€เธเนเธ Date object เธซเธฃเธทเธญ timestamp
        endTimeMoment = moment.tz(clockOutTime, CONFIG.TIMEZONE);
      }
    } else {
      endTimeMoment = moment().tz(CONFIG.TIMEZONE);
    }

    if (!clockInMoment.isValid()) {
      console.error(`โ Invalid clockInTime format: "${clockInTime}"`);
      return 0;
    }

    if (clockOutTime && !endTimeMoment.isValid()) {
      console.error(`โ Invalid clockOutTime format: "${clockOutTime}"`);
      return 0;
    }

    // เธเธณเธเธงเธ“เธเธงเธฒเธกเนเธ•เธเธ•เนเธฒเธเธเธญเธเน€เธงเธฅเธฒเนเธเธซเธเนเธงเธขเธเธฑเนเธงเนเธกเธ (เน€เธซเธกเธทเธญเธ admin stats)
    const hours = endTimeMoment.diff(clockInMoment, 'hours', true);

    // Debug: เนเธชเธ”เธเธเธฒเธฃเธเธณเธเธงเธ“
    console.log(`โฐ Working hours calculation:`, {
      clockInOriginal: clockInTime,
      clockOutOriginal: clockOutTime || 'current time',
      clockIn: clockInMoment.format('YYYY-MM-DD HH:mm:ss'),
      endTime: endTimeMoment.format('YYYY-MM-DD HH:mm:ss'),
      diffHours: hours.toFixed(2),
      clockInValid: clockInMoment.isValid(),
      endTimeValid: endTimeMoment.isValid()
    });

    // เธ•เธฃเธงเธเธชเธญเธเนเธซเนเนเธเนเนเธเธงเนเธฒเนเธกเนเน€เธเนเธเธฅเธ (เธเนเธญเธเธเธฑเธเธเธฑเธเธซเธฒ timezone)
    if (hours >= 0) {
      return hours;
    } else {
      console.warn(`โ ๏ธ Negative working hours detected: ${hours.toFixed(2)}, setting to 0`);
      return 0;
    }
  } catch (error) {
    console.error('โ Error calculating working hours:', error);
    return 0;
  }
}

/**
 * ๐”ง เธเธฑเธเธเนเธเธฑเธ helper เธชเธณเธซเธฃเธฑเธ parse เธงเธฑเธเธ—เธตเนเนเธเธฃเธนเธเนเธเธเธ•เนเธฒเธเน
 * @param {string} dateTimeString - เธงเธฑเธเธ—เธตเนเน€เธงเธฅเธฒเนเธเธฃเธนเธเนเธเธ string
 * @returns {moment.Moment} - moment object
 */
function parseDateTime(dateTimeString) {
  if (!dateTimeString) {
    return moment().tz(CONFIG.TIMEZONE);
  }

  // เธ•เธฃเธงเธเธชเธญเธ format เธเธญเธ dateTimeString
  if (typeof dateTimeString === 'string') {
    if (dateTimeString.match(/^\d{2}\/\d{2}\/\d{4}/)) {
      // เธฃเธนเธเนเธเธ DD/MM/YYYY HH:mm:ss
      return moment.tz(dateTimeString, 'DD/MM/YYYY HH:mm:ss', CONFIG.TIMEZONE);
    } else if (dateTimeString.match(/^\d{4}-\d{2}-\d{2}/)) {
      // เธฃเธนเธเนเธเธ YYYY-MM-DD HH:mm:ss
      return moment.tz(dateTimeString, 'YYYY-MM-DD HH:mm:ss', CONFIG.TIMEZONE);
    } else {
      // เธฅเธญเธเนเธเน auto parse
      return moment.tz(dateTimeString, CONFIG.TIMEZONE);
    }
  } else {
    // เธ–เนเธฒเน€เธเนเธ Date object เธซเธฃเธทเธญ timestamp
    return moment.tz(dateTimeString, CONFIG.TIMEZONE);
  }
}

// เธชเธฃเนเธฒเธ hash password (เนเธเนเนเธเธเธฒเธฃเธ•เธฑเนเธเธฃเธซเธฑเธชเธเนเธฒเธเธเธฃเธฑเนเธเนเธฃเธ)
async function createPassword(plainPassword) {
  return await bcrypt.hash(plainPassword, 10);
}

// ========== Middleware ==========
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Security middleware เธชเธณเธซเธฃเธฑเธ webhook
app.use('/api/webhook', (req, res, next) => {
  const providedSecret = req.headers['x-webhook-secret'] || req.query.secret;
  if (providedSecret !== CONFIG.RENDER.GSA_WEBHOOK_SECRET) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  next();
});

// Admin Authentication Middleware
function authenticateAdmin(req, res, next) {
  const authHeader = req.headers.authorization;
  const token = authHeader && authHeader.split(' ')[1];

  if (!token) {
    return res.status(401).json({
      success: false,
      error: 'Access token required',
      errorCode: 'NO_TOKEN'
    });
  }

  try {
    const decoded = jwt.verify(token, CONFIG.ADMIN.JWT_SECRET);
    req.user = decoded;
    next();
  } catch (error) {
    console.error('JWT verification error:', error.name, error.message);

    // เธเธฑเธ”เธเธฒเธฃ error เนเธ•เนเธฅเธฐเธเธฃเธฐเน€เธ เธ—
    let errorResponse = {
      success: false,
      error: 'Authentication failed'
    };

    if (error.name === 'TokenExpiredError') {
      errorResponse.error = 'Token has expired. Please login again.';
      errorResponse.errorCode = 'TOKEN_EXPIRED';
      errorResponse.expiredAt = error.expiredAt;
    } else if (error.name === 'JsonWebTokenError') {
      errorResponse.error = 'Invalid token format';
      errorResponse.errorCode = 'INVALID_TOKEN';
    } else if (error.name === 'NotBeforeError') {
      errorResponse.error = 'Token not active yet';
      errorResponse.errorCode = 'TOKEN_NOT_ACTIVE';
    } else {
      errorResponse.errorCode = 'TOKEN_ERROR';
    }

    return res.status(401).json(errorResponse);
  }
}

// Serve static files
app.use(express.static('public'));

// Admin routes - เนเธเนเนเธเธฅเนเธเธฒเธ public folder
app.get('/admin/login', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

app.get('/admin/dashboard', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

app.get('/admin', (req, res) => {
  res.redirect('/admin/login');
});

// Serve ads.txt specifically
app.get('/ads.txt', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'ads.txt'));
});

// Serve robots.txt (optional)
app.get('/robots.txt', (req, res) => {
  res.type('text/plain');
  res.send('User-agent: *\nDisallow: /api/\nAllow: /ads.txt');
});

// ========== Keep-Alive Service ==========
class KeepAliveService {
  constructor() {
    this.isEnabled = CONFIG.RENDER.KEEP_ALIVE_ENABLED;
    this.serviceUrl = CONFIG.RENDER.SERVICE_URL;
    this.startTime = new Date();
    this.pingCount = 0;
    this.errorCount = 0;
  }

  init() {
    if (!this.isEnabled) {
      console.log('๐”ด Keep-Alive disabled');
      return;
    }

    console.log('๐ข Keep-Alive service started');
    console.log(`๐“ Service URL: ${this.serviceUrl}`);

    // เน€เธงเธฅเธฒเธ—เธณเธเธฒเธ: 05:00-10:00 เนเธฅเธฐ 15:00-20:00 (เน€เธงเธฅเธฒเนเธ—เธข)
    // เธเธดเธเธ—เธธเธ 10 เธเธฒเธ—เธต
    cron.schedule('*/10 * * * *', () => {
      this.checkAndPing();
    }, {
      scheduled: true,
      timezone: CONFIG.TIMEZONE
    });

    // ๐• เธ•เธฃเธงเธเธชเธญเธเนเธฅเธฐเธเธฑเธ”เธเธฒเธฃเธเธเธ—เธตเนเธฅเธทเธกเธฅเธเน€เธงเธฅเธฒเธญเธญเธ - เธ—เธธเธเธงเธฑเธเน€เธงเธฅเธฒ 23:59
    cron.schedule('59 23 * * *', async () => {
      console.log('๐” Starting daily missed checkout check at 23:59...');
      try {
        const result = await sheetsService.checkAndHandleMissedCheckouts();
        console.log('โ… Daily missed checkout check completed:', result);
      } catch (error) {
        console.error('โ Error in daily missed checkout check:', error);
      }
    }, {
      scheduled: true,
      timezone: CONFIG.TIMEZONE
    });

    // Ping เธ—เธฑเธเธ—เธตเน€เธกเธทเนเธญเน€เธฃเธดเนเธกเธ•เนเธ
    setTimeout(() => this.ping(), 5000);
  }

  checkAndPing() {
    const now = new Date();
    const hour = now.getHours();

    // เน€เธเนเธเธงเนเธฒเธญเธขเธนเนเนเธเน€เธงเธฅเธฒเธ—เธณเธเธฒเธเนเธซเธก
    const isWorkingHour = (hour >= 5 && hour < 10) || (hour >= 15 && hour < 20);

    if (isWorkingHour) {
      this.ping();
    } else {
      console.log(`๐ด Outside working hours (${hour}:00), skipping ping`);
    }
  }

  async ping() {
    try {
      const response = await fetch(`${this.serviceUrl}/api/ping`, {
        method: 'GET',
        headers: {
          'User-Agent': 'KeepAlive-Service/1.0'
        }
      });

      this.pingCount++;

      if (response.ok) {
        console.log(`โ… Keep-Alive ping #${this.pingCount} successful`);
        this.errorCount = 0; // Reset error count on success
      } else {
        throw new Error(`HTTP ${response.status}`);
      }

    } catch (error) {
      this.errorCount++;
      console.log(`โ Keep-Alive ping #${this.pingCount} failed:`, error.message);

      // เธซเธฒเธเธเธดเธ”เธเธฅเธฒเธ”เธ•เธดเธ”เธ•เนเธญเธเธฑเธ 5 เธเธฃเธฑเนเธ เนเธซเนเธฅเธญเธ ping เนเธซเธกเนเธซเธฅเธฑเธ 1 เธเธฒเธ—เธต
      if (this.errorCount >= 5) {
        console.log('๐” Too many errors, will retry in 1 minute');
        setTimeout(() => this.ping(), 60000);
      }
    }
  }

  getStats() {
    const uptime = Math.floor((Date.now() - this.startTime.getTime()) / 1000);
    return {
      enabled: this.isEnabled,
      uptime: uptime,
      pingCount: this.pingCount,
      errorCount: this.errorCount,
      lastPing: new Date().toISOString()
    };
  }
}

// ========== Google Sheets Service ==========
class GoogleSheetsService {
  constructor() {
    this.doc = null;
    this.isInitialized = false;
    // เน€เธเธดเนเธกเธฃเธฐเธเธ caching เน€เธเธทเนเธญเธฅเธ”เธเธฒเธฃเน€เธฃเธตเธขเธ API
    this.cache = {
      employees: { data: null, timestamp: null, ttl: 300000 }, // 5 เธเธฒเธ—เธต
      onWork: { data: null, timestamp: null, ttl: 60000 }, // 1 เธเธฒเธ—เธต  
      main: { data: null, timestamp: null, ttl: 30000 }, // 30 เธงเธดเธเธฒเธ—เธต
      stats: { data: null, timestamp: null, ttl: 120000 } // 2 เธเธฒเธ—เธต
    };
    this.emergencyMode = false; // เน€เธฃเธดเนเธกเธ•เนเธเธเธดเธ”เธฃเธฐเธเธ emergency mode
  }

  async initialize() {
    if (this.isInitialized) return;

    try {
      const serviceAccountAuth = new JWT({
        email: CONFIG.GOOGLE_SHEETS.CLIENT_EMAIL,
        key: CONFIG.GOOGLE_SHEETS.PRIVATE_KEY,
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
      });

      this.doc = new GoogleSpreadsheet(CONFIG.GOOGLE_SHEETS.SPREADSHEET_ID, serviceAccountAuth);
      await this.doc.loadInfo();

      console.log(`โ… Connected to Google Sheets: ${this.doc.title}`);
      this.isInitialized = true;

    } catch (error) {
      console.error('โ Failed to initialize Google Sheets:', error);
      throw error;
    }
  }  // เน€เธเธดเนเธกเธเธฑเธเธเนเธเธฑเธ cache helper
  isCacheValid(cacheKey) {
    const cache = this.cache[cacheKey];
    if (!cache || !cache.data || !cache.timestamp) return false;
    return (Date.now() - cache.timestamp) < cache.ttl;
  }

  setCache(cacheKey, data) {
    if (!this.cache[cacheKey]) {
      this.cache[cacheKey] = { data: null, timestamp: null, ttl: 300000 }; // default 5 min
    }
    this.cache[cacheKey] = {
      data: data,
      timestamp: Date.now(),
      ttl: this.cache[cacheKey].ttl
    };
  }

  getCache(cacheKey) {
    const cache = this.cache[cacheKey];
    return cache && cache.data ? cache.data : null;
  }

  clearCache(cacheKey = null) {
    if (cacheKey) {
      this.cache[cacheKey].data = null;
      this.cache[cacheKey].timestamp = null;
    } else {
      // Clear all cache
      Object.keys(this.cache).forEach(key => {
        this.cache[key].data = null;
        this.cache[key].timestamp = null;
      });
    }
  }

  async getSheet(sheetName) {
    if (!this.isInitialized) {
      await this.initialize();
    }

    const sheet = this.doc.sheetsByTitle[sheetName];
    if (!sheet) {
      throw new Error(`Sheet ${sheetName} not found`);
    }

    return sheet;
  }  // เน€เธเธดเนเธกเธเธฑเธเธเนเธเธฑเธเธ”เธถเธเธเนเธญเธกเธนเธฅเธเธฃเนเธญเธก cache เนเธฅเธฐ rate limiting
  async getCachedSheetData(sheetName) {
    const cacheKey = sheetName.toLowerCase().replace(/\s+/g, '');

    // เธ•เธฃเธงเธเธชเธญเธ cache เธเนเธญเธ
    if (this.isCacheValid(cacheKey)) {
      console.log(`๐“ Using cached data for ${sheetName}`);
      return this.getCache(cacheKey);
    }

    // เธ•เธฃเธงเธเธชเธญเธ rate limit เธเนเธญเธเน€เธฃเธตเธขเธ API
    if (!apiMonitor.canMakeAPICall()) {
      console.warn(`โ ๏ธ API rate limit reached, using stale cache for ${sheetName}`);
      const staleData = this.getCache(cacheKey);
      if (staleData) {
        return staleData;
      }
      throw new Error('Rate limit exceeded and no cached data available');
    }

    console.log(`๐” Fetching fresh data from ${sheetName}`);
    apiMonitor.logAPICall(`getCachedSheetData:${sheetName}`);

    try {
      const sheet = await this.getSheet(sheetName);

      let rows;
      if (sheetName === CONFIG.SHEETS.ON_WORK) {
        rows = await sheet.getRows({ offset: 1 }); // เน€เธฃเธดเนเธกเธเธฒเธเนเธ–เธง 3
      } else {
        rows = await sheet.getRows();
      }

      // เธเธฑเธเธ—เธถเธเธฅเธ cache
      this.setCache(cacheKey, rows);

      // เน€เธชเธฃเนเธเธชเธดเนเธ API call
      apiMonitor.finishCall();
      return rows;

    } catch (error) {
      // เน€เธชเธฃเนเธเธชเธดเนเธ API call เนเธกเนเธเธฐ error
      apiMonitor.finishCall();

      console.error(`โ API Error for ${sheetName}:`, error.message);

      // เธ–เนเธฒเน€เธเนเธ quota error, เนเธเน stale cache
      if (error.message.includes('quota') || error.message.includes('limit') ||
        error.message.includes('429') || error.message.includes('RATE_LIMIT')) {
        console.warn(`โ ๏ธ Quota exceeded for ${sheetName}, using stale cache`);
        const staleData = this.getCache(cacheKey);
        if (staleData) {
          console.log(`๐“ Using stale cache for ${sheetName} (${staleData.length} items)`);
          return staleData;
        }
      }

      throw error;
    }
  }

  // เธเธฑเธเธเนเธเธฑเธเธเนเธงเธขเน€เธซเธฅเธทเธญเธชเธณเธซเธฃเธฑเธเธเธฒเธฃเน€เธเธฃเธตเธขเธเน€เธ—เธตเธขเธเธเธทเนเธญ
  normalizeEmployeeName(name) {
    if (!name) return '';

    return name.toString()
      .trim()
      .replace(/\s+/g, ' ')
      .toLowerCase();
  }

  isNameMatch(inputName, compareName) {
    if (!inputName || !compareName) return false;

    const normalizedInput = this.normalizeEmployeeName(inputName);
    const normalizedCompare = this.normalizeEmployeeName(compareName);

    return normalizedInput === normalizedCompare ||
      normalizedInput.includes(normalizedCompare) ||
      normalizedCompare.includes(normalizedInput);
  }
  /**
 * เธ•เธฃเธงเธเธชเธญเธเธงเนเธฒเธเธเธฑเธเธเธฒเธเธเธเธเธตเนเนเธ”เนเธฃเธฑเธเธเธฒเธฃเธขเธเน€เธงเนเธเธเธฒเธเธเธฒเธฃเธฅเธเน€เธงเธฅเธฒเธญเธญเธเธญเธฑเธ•เนเธเธกเธฑเธ•เธดเธซเธฃเธทเธญเนเธกเน
 * @param {string} employeeName - เธเธทเนเธญเธเธเธฑเธเธเธฒเธ
 * @returns {boolean} - true เธ–เนเธฒเนเธ”เนเธฃเธฑเธเธเธฒเธฃเธขเธเน€เธงเนเธ, false เธ–เนเธฒเนเธกเนเนเธ”เน
 */
  isEmployeeExempt(employeeName) {
    if (!employeeName) {
      return false;
    }

    const normalizedInputName = this.normalizeEmployeeName(employeeName);

    // 1. ตรวจสอบจากฐานข้อมูล SQLite (ที่เพิ่มผ่าน Admin UI)
    try {
      if (typeof sqliteService !== 'undefined' && sqliteService && typeof sqliteService.getNightShiftEmployees === 'function') {
        const dbExempts = sqliteService.getNightShiftEmployees();
        for (const record of dbExempts) {
          const normalizedExemptName = this.normalizeEmployeeName(record.employee_name);
          const isExactMatch = normalizedInputName === normalizedExemptName;
          const isPartialMatch = normalizedInputName.includes(normalizedExemptName) ||
            normalizedExemptName.includes(normalizedInputName);
            
          if (isExactMatch || isPartialMatch) {
            console.log(`🛡️ Employee exempt match found (Database): "${employeeName}" ↔️ "${record.employee_name}"`);
            return true;
          }
        }
      }
    } catch (error) {
      console.error('Error checking SQLite night shift:', error);
    }

    // 2. ตรวจสอบกับรายชื่อที่ยกเว้นจาก CONFIG (Fallback)
    if (!CONFIG.AUTO_CHECKOUT.EXEMPT_EMPLOYEES) {
      return false;
    }

    // เธ•เธฃเธงเธเธชเธญเธเธเธฑเธเธฃเธฒเธขเธเธทเนเธญเธ—เธตเนเธขเธเน€เธงเนเธ
    return CONFIG.AUTO_CHECKOUT.EXEMPT_EMPLOYEES.some(exemptName => {
      const normalizedExemptName = this.normalizeEmployeeName(exemptName);

      // เธ•เธฃเธงเธเธชเธญเธเนเธเธเธซเธฅเธฒเธขเธฃเธนเธเนเธเธ
      const isExactMatch = normalizedInputName === normalizedExemptName;
      const isPartialMatch = normalizedInputName.includes(normalizedExemptName) ||
        normalizedExemptName.includes(normalizedInputName);

      if (isExactMatch || isPartialMatch) {
        console.log(`๐ก๏ธ Employee exempt match found: "${employeeName}" โ” "${exemptName}"`);
        return true;
      }

      return false;
    });
  }

  // ๐”ง เน€เธเธดเนเธกเธเธฑเธเธเนเธเธฑเธ helper เธชเธณเธซเธฃเธฑเธ parse เนเธฅเธฐ format เธงเธฑเธเธ—เธตเน
  parseDateTime(dateTimeString) {
    return parseDateTime(dateTimeString);
  }

  parseAndFormatTime(dateTimeString, format = 'HH:mm:ss') {
    try {
      const parsed = this.parseDateTime(dateTimeString);
      if (parsed.isValid()) {
        return parsed.format(format);
      }
      return dateTimeString?.toString() || '';
    } catch (error) {
      console.error('Error parsing and formatting time:', error);
      return dateTimeString?.toString() || '';
    }
  }

  async getEmployees() {
    try {
      // เนเธเน cached data เนเธ—เธเธเธฒเธฃเน€เธฃเธตเธขเธ API เนเธซเธกเน
      const rows = await this.getCachedSheetData(CONFIG.SHEETS.EMPLOYEES);

      const employees = rows.map(row => row.get('เธเธทเนเธญ-เธเธฒเธกเธชเธเธธเธฅ')).filter(name => name);
      return employees;

    } catch (error) {
      console.error('Error getting employees:', error);
      return [];
    }
  } async getEmployeeStatus(employeeName) {
    try {
      // เนเธเน safe method เนเธ—เธ
      const rows = await this.safeGetCachedSheetData(CONFIG.SHEETS.ON_WORK);

      console.log(`๐” Checking status for: "${employeeName}"`);
      console.log(`๐“ Total rows in ON_WORK (from row 3): ${rows.length}`);

      if (rows.length === 0) {
        console.log('๐“ ON_WORK sheet is empty (from row 3)');
        return { isOnWork: false, workRecord: null };
      }

      const workRecord = rows.find(row => {
        const systemName = row.get('เธเธทเนเธญเนเธเธฃเธฐเธเธ');
        const employeeName2 = row.get('เธเธทเนเธญเธเธเธฑเธเธเธฒเธ');

        const isMatch = this.isNameMatch(employeeName, systemName) ||
          this.isNameMatch(employeeName, employeeName2);

        if (isMatch) {
          console.log(`โ… Found match: "${employeeName}" โ” "${systemName || employeeName2}"`);
        }

        return isMatch;
      });

      if (workRecord) {
        let mainRowIndex = null;

        const rowRef1 = workRecord.get('เนเธ–เธงเธญเนเธฒเธเธญเธดเธ');
        const rowRef2 = workRecord.get('เนเธ–เธงเนเธMain');

        if (rowRef1 && !isNaN(parseInt(rowRef1))) {
          mainRowIndex = parseInt(rowRef1);
        } else if (rowRef2 && !isNaN(parseInt(rowRef2))) {
          mainRowIndex = parseInt(rowRef2);
        }

        console.log(`โ… Employee "${employeeName}" is currently working`);

        return {
          isOnWork: true,
          workRecord: {
            row: workRecord,
            mainRowIndex: mainRowIndex,
            clockIn: workRecord.get('เน€เธงเธฅเธฒเน€เธเนเธฒ'),
            systemName: workRecord.get('เธเธทเนเธญเนเธเธฃเธฐเธเธ'),
            employeeName: workRecord.get('เธเธทเนเธญเธเธเธฑเธเธเธฒเธ')
          }
        };
      } else {
        console.log(`โ Employee "${employeeName}" is not currently working`);
        return { isOnWork: false, workRecord: null };
      }

    } catch (error) {
      console.error('โ Error checking employee status:', error);
      return { isOnWork: false, workRecord: null };
    }
  }
  // Admin functions
  async getAdminStats() {
    try {
      // เธ•เธฃเธงเธเธชเธญเธ cache เธชเธณเธซเธฃเธฑเธ stats เธเนเธญเธ
      if (this.isCacheValid('stats')) {
        console.log('๐“ Using cached admin stats');
        return this.getCache('stats');
      }

      console.log('๐” Fetching fresh admin stats data');      // เนเธเน safe method เนเธ—เธ เธเธฒเธฃเน€เธฃเธตเธขเธ API
      const [employees, onWorkRows, mainRows] = await Promise.all([
        this.safeGetCachedSheetData(CONFIG.SHEETS.EMPLOYEES),
        this.safeGetCachedSheetData(CONFIG.SHEETS.ON_WORK),
        this.safeGetCachedSheetData(CONFIG.SHEETS.MAIN)
      ]);

      const totalEmployees = employees.length;// เธซเธฒเธเธณเธเธงเธเธเธเธ—เธตเนเธกเธฒเธ—เธณเธเธฒเธเธงเธฑเธเธเธตเน (เนเธเนเธเนเธญเธกเธนเธฅเธเธฒเธ ON_WORK sheet เธ—เธตเนเธกเธตเธงเธฑเธเธ—เธตเนเธงเธฑเธเธเธตเน)
      const today = moment().tz(CONFIG.TIMEZONE).format('YYYY-MM-DD');
      console.log(`๐“… Today date for comparison: ${today}`);
      console.log(`๐“ Total MAIN sheet records: ${mainRows.length}`);
      console.log(`๏ฟฝ Total ON_WORK sheet records: ${onWorkRows.length}`);

      // เธเธฑเธเธเธฒเธ ON_WORK sheet เธ—เธตเนเธกเธตเธงเธฑเธเธ—เธตเนเธงเธฑเธเธเธตเน
      const presentToday = onWorkRows.filter(row => {
        const clockInDate = row.get('เน€เธงเธฅเธฒเน€เธเนเธฒ');
        if (!clockInDate) return false;

        try {
          const employeeName = row.get('เธเธทเนเธญเธเธเธฑเธเธเธฒเธ') || row.get('เธเธทเนเธญเนเธเธฃเธฐเธเธ');
          let dateStr = '';

          // เธ–เนเธฒเน€เธเนเธ string format 'YYYY-MM-DD HH:mm:ss'
          if (typeof clockInDate === 'string' && clockInDate.includes(' ')) {
            dateStr = clockInDate.split(' ')[0];
            const isToday = dateStr === today;

            if (isToday) {
              console.log(`โ… Present today (ON_WORK): ${employeeName} - ${clockInDate} (date: ${dateStr})`);
            }

            return isToday;
          }

          // เธ–เนเธฒเน€เธเนเธ ISO format
          if (typeof clockInDate === 'string' && clockInDate.includes('T')) {
            dateStr = clockInDate.split('T')[0];
            const isToday = dateStr === today;

            if (isToday) {
              console.log(`โ… Present today (ON_WORK ISO): ${employeeName} - ${clockInDate} (date: ${dateStr})`);
            }

            return isToday;
          }

          return false;
        } catch (error) {
          console.warn(`โ ๏ธ Error parsing date in ON_WORK: ${clockInDate}`, error);
          return false;
        }
      }).length;

      console.log(`๐“ Present today count: ${presentToday} out of ${onWorkRows.length} ON_WORK records`);

      // workingNow เธเธงเธฃเน€เธเนเธเธเธณเธเธงเธเธเธเธ—เธตเนเธกเธฒเธ—เธณเธเธฒเธเธงเธฑเธเธเธตเน (เน€เธ”เธตเธขเธงเธเธฑเธ presentToday)
      const workingNow = presentToday;
      const absentToday = totalEmployees - presentToday;      // เธฃเธฒเธขเธเธทเนเธญเธเธเธฑเธเธเธฒเธเธ—เธตเนเธเธณเธฅเธฑเธเธ—เธณเธเธฒเธ
      const workingEmployees = onWorkRows.map(row => {
        const clockInTime = row.get('เน€เธงเธฅเธฒเน€เธเนเธฒ');
        let workingHours = '0 เธเธก.';

        if (clockInTime) {
          // ๐ฏ เนเธเนเธเธฑเธเธเนเธเธฑเธเธเธณเธเธงเธ“เน€เธงเธฅเธฒเนเธเธเน€เธ”เธตเธขเธงเธเธฑเธเธเธฑเธ clock out
          const hours = calculateWorkingHours(clockInTime);

          if (hours > 0) {
            workingHours = `${hours.toFixed(1)} เธเธก.`;
          } else {
            workingHours = '0 เธเธก.';
          }
        }

        return {
          name: row.get('เธเธทเนเธญเธเธเธฑเธเธเธฒเธ') || row.get('เธเธทเนเธญเนเธเธฃเธฐเธเธ'),
          clockIn: clockInTime ? this.parseAndFormatTime(clockInTime, 'HH:mm') : '',
          workingHours
        };
      }); const stats = {
        totalEmployees,
        presentToday,
        workingNow,
        absentToday,
        workingEmployees
      };

      console.log('๐“ Admin stats summary:', {
        totalEmployees,
        presentToday,
        workingNow,
        absentToday,
        workingEmployeesCount: workingEmployees.length
      });

      // เธเธฑเธเธ—เธถเธเธฅเธ cache
      this.setCache('stats', stats);

      return stats;

    } catch (error) {
      console.error('Error getting admin stats:', error);
      throw error;
    }
  }
  async getReportData(type, params) {
    try {
      console.log(`๐“ Getting report data for type: ${type}`, params);

      // เนเธเน safe cached data method
      const rows = await this.safeGetCachedSheetData(CONFIG.SHEETS.MAIN);

      if (!rows || rows.length === 0) {
        console.log('โ ๏ธ No data found in MAIN sheet');
        return [];
      }

      console.log(`๐“ Found ${rows.length} total records in MAIN sheet`);

      // Debug: เนเธชเธ”เธเธ•เธฑเธงเธญเธขเนเธฒเธเธเนเธญเธกเธนเธฅเนเธกเนเธเธตเนเนเธ–เธงเนเธฃเธ
      if (rows.length > 0) {
        console.log('๐“ Sample data (first 3 rows):');
        for (let i = 0; i < Math.min(3, rows.length); i++) {
          const row = rows[i];
          // เนเธเน index เนเธ—เธเน€เธเธทเนเธญเธเธเธฒเธ sheet เนเธกเนเธกเธต header
          const employee = row._rawData[0]; // column 0: เธเธทเนเธญเธเธเธฑเธเธเธฒเธ
          const clockIn = row._rawData[3];  // column 3: เน€เธงเธฅเธฒเน€เธเนเธฒ
          console.log(`   Row ${i + 1}: Employee="${employee}", ClockIn="${clockIn}" (type: ${typeof clockIn})`);
        }

        // Debug: เนเธชเธ”เธ headers เธเธญเธ sheet
        console.log('๐“ Sheet headers:', Object.keys(rows[0]._rawData));

        // Debug: เนเธชเธ”เธเธเนเธฒเธเธญเธเนเธ•เนเธฅเธฐ column เนเธเนเธ–เธงเนเธฃเธ
        const firstRow = rows[0];
        console.log('๐“ First row values by index:');
        console.log(`   Column 0: "${firstRow._rawData[0]}" (should be เธเธทเนเธญเธเธเธฑเธเธเธฒเธ)`);
        console.log(`   Column 1: "${firstRow._rawData[1]}" (should be Line name)`);
        console.log(`   Column 2: "${firstRow._rawData[2]}" (should be เธฃเธนเธเธ เธฒเธ)`);
        console.log(`   Column 3: "${firstRow._rawData[3]}" (should be เน€เธงเธฅเธฒเน€เธเนเธฒ)`);
        console.log(`   Column 4: "${firstRow._rawData[4]}" (should be userinfo/เธซเธกเธฒเธขเน€เธซเธ•เธธ)`);
        console.log(`   Column 5: "${firstRow._rawData[5]}" (should be เน€เธงเธฅเธฒเธญเธญเธ)`);
        console.log(`   Column 6: "${firstRow._rawData[6]}" (should be เธเธดเธเธฑเธ”เน€เธเนเธฒ)`);
        console.log(`   Column 7: "${firstRow._rawData[7]}" (should be เธชเธ–เธฒเธเธ—เธตเนเน€เธเนเธฒ)`);
        console.log(`   Column 8: "${firstRow._rawData[8]}" (should be เธเธดเธเธฑเธ”เธญเธญเธ)`);
        console.log(`   Column 9: "${firstRow._rawData[9]}" (should be เธ—เธตเนเธญเธขเธนเนเธญเธญเธ)`);
        console.log(`   Column 10: "${firstRow._rawData[10]}" (should be เธเธฑเนเธงเนเธกเธเธ—เธณเธเธฒเธ)`);
        console.log(`   Column 11: "${firstRow._rawData[11]}" (should be เธซเธกเธฒเธขเน€เธซเธ•เธธเน€เธ”เธดเธก - เนเธกเนเนเธเนเนเธฅเนเธง)`);
      }

      let filteredRows = [];

      switch (type) {
        case 'daily':
          const targetDate = moment(params.date).tz(CONFIG.TIMEZONE).format('YYYY-MM-DD');
          console.log(`๐“… Filtering for daily report: ${targetDate}`);

          filteredRows = rows.filter(row => {
            const clockIn = row._rawData[3]; // column 3: เน€เธงเธฅเธฒเน€เธเนเธฒ
            if (!clockIn) return false;

            try {
              let dateStr = '';
              console.log(`๐” Checking clockIn: "${clockIn}" (type: ${typeof clockIn})`);

              // เธ–เนเธฒเน€เธเนเธ string format 'DD/MM/YYYY HH:mm:ss'
              if (typeof clockIn === 'string' && clockIn.match(/^\d{2}\/\d{2}\/\d{4}/)) {
                const datePart = clockIn.split(' ')[0]; // "26/06/2025"
                const [day, month, year] = datePart.split('/');
                dateStr = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
              }
              // เธ–เนเธฒเน€เธเนเธ string format 'YYYY-MM-DD HH:mm:ss'
              else if (typeof clockIn === 'string' && clockIn.includes(' ')) {
                dateStr = clockIn.split(' ')[0];
              } else if (typeof clockIn === 'string' && clockIn.includes('T')) {
                // ISO format
                dateStr = clockIn.split('T')[0];
              } else if (typeof clockIn === 'string' && clockIn.match(/^\d{4}-\d{2}-\d{2}$/)) {
                // Already in YYYY-MM-DD format
                dateStr = clockIn;
              } else {
                // Date object เธซเธฃเธทเธญ format เธญเธทเนเธ
                const rowDate = moment(clockIn).tz(CONFIG.TIMEZONE);
                if (rowDate.isValid()) {
                  dateStr = rowDate.format('YYYY-MM-DD');
                } else {
                  console.warn(`โ ๏ธ Invalid date format: "${clockIn}"`);
                  return false;
                }
              }

              console.log(`๐“… Extracted date: "${dateStr}" vs target: "${targetDate}"`);
              const isMatch = dateStr === targetDate;
              if (isMatch) {
                console.log(`โ… Date match found: ${row._rawData[0]} - ${clockIn}`);
              } else if (clockIn && clockIn.includes('26')) {
                console.log(`โ“ Potential match (contains '26'): ${row._rawData[0]} - ${clockIn} -> ${dateStr}`);
              }

              return isMatch;
            } catch (error) {
              console.warn('โ Error parsing date for daily report:', clockIn, error);
              return false;
            }
          });

          console.log(`๐“ Daily filter result: ${filteredRows.length} records found for ${targetDate}`);
          break;

        case 'monthly':
          const month = parseInt(params.month);
          const year = parseInt(params.year);
          console.log(`๐“… Filtering for monthly report: ${month}/${year}`);

          filteredRows = rows.filter(row => {
            const clockIn = row._rawData[3]; // column 3: เน€เธงเธฅเธฒเน€เธเนเธฒ
            if (!clockIn) return false;

            try {
              let dateStr = '';

              // เนเธเนเธงเธดเธเธตเน€เธ”เธตเธขเธงเธเธฑเธเธฃเธฒเธขเธเธฒเธเธฃเธฒเธขเธงเธฑเธ เน€เธเธทเนเธญเธเธงเธฒเธกเธชเธญเธ”เธเธฅเนเธญเธ
              if (typeof clockIn === 'string' && clockIn.match(/^\d{2}\/\d{2}\/\d{4}/)) {
                const datePart = clockIn.split(' ')[0]; // "26/06/2025"
                const [day, monthPart, yearPart] = datePart.split('/');
                dateStr = `${yearPart}-${monthPart.padStart(2, '0')}-${day.padStart(2, '0')}`;
              }
              // เธ–เนเธฒเน€เธเนเธ string format 'YYYY-MM-DD HH:mm:ss'
              else if (typeof clockIn === 'string' && clockIn.includes(' ')) {
                dateStr = clockIn.split(' ')[0];
              } else if (typeof clockIn === 'string' && clockIn.includes('T')) {
                // ISO format
                dateStr = clockIn.split('T')[0];
              } else if (typeof clockIn === 'string' && clockIn.match(/^\d{4}-\d{2}-\d{2}$/)) {
                // Already in YYYY-MM-DD format
                dateStr = clockIn;
              } else {
                // Date object เธซเธฃเธทเธญ format เธญเธทเนเธ
                const rowDate = moment(clockIn).tz(CONFIG.TIMEZONE);
                if (rowDate.isValid()) {
                  dateStr = rowDate.format('YYYY-MM-DD');
                } else {
                  console.warn(`โ ๏ธ Invalid date format: "${clockIn}"`);
                  return false;
                }
              }

              // เนเธเธฅเธเน€เธเนเธ Date object เน€เธเธทเนเธญเน€เธเธฃเธตเธขเธเน€เธ—เธตเธขเธ
              const rowDate = moment(dateStr).tz(CONFIG.TIMEZONE);
              if (!rowDate.isValid()) return false;

              const isMatch = rowDate.month() + 1 === month && rowDate.year() === year;

              if (isMatch) {
                console.log(`โ… Monthly match found: ${row._rawData[0]} - ${clockIn} -> ${dateStr}`);
              }

              return isMatch;
            } catch (error) {
              console.warn('โ Error parsing date for monthly report:', clockIn, error);
              return false;
            }
          });
          break;

        case 'range':
          const startMoment = moment(params.startDate).tz(CONFIG.TIMEZONE).startOf('day');
          const endMoment = moment(params.endDate).tz(CONFIG.TIMEZONE).endOf('day');
          console.log(`๐“… Filtering for range report: ${startMoment.format('YYYY-MM-DD')} to ${endMoment.format('YYYY-MM-DD')}`);

          filteredRows = rows.filter(row => {
            const clockIn = row._rawData[3]; // column 3: เน€เธงเธฅเธฒเน€เธเนเธฒ
            if (!clockIn) return false;

            try {
              let rowMoment;

              // เธ–เนเธฒเน€เธเนเธ string format 'DD/MM/YYYY HH:mm:ss'
              if (typeof clockIn === 'string' && clockIn.match(/^\d{2}\/\d{2}\/\d{4}/)) {
                rowMoment = moment(clockIn, 'DD/MM/YYYY HH:mm:ss').tz(CONFIG.TIMEZONE);
              }
              // เธ–เนเธฒเน€เธเนเธ string format 'YYYY-MM-DD HH:mm:ss'
              else if (typeof clockIn === 'string' && clockIn.includes(' ')) {
                rowMoment = moment(clockIn, 'YYYY-MM-DD HH:mm:ss').tz(CONFIG.TIMEZONE);
              } else if (typeof clockIn === 'string' && clockIn.includes('T')) {
                // ISO format
                rowMoment = moment(clockIn).tz(CONFIG.TIMEZONE);
              } else {
                // Date object เธซเธฃเธทเธญ format เธญเธทเนเธ
                rowMoment = moment(clockIn).tz(CONFIG.TIMEZONE);
              }

              if (!rowMoment.isValid()) return false;

              return rowMoment.isBetween(startMoment, endMoment, null, '[]');
            } catch (error) {
              console.warn('Error parsing date for range report:', clockIn, error);
              return false;
            }
          });
          break;

        default:
          throw new Error(`Unsupported report type: ${type}`);
      }

      console.log(`๐“ Filtered to ${filteredRows.length} records for ${type} report`);

      // เนเธเธฅเธเธเนเธญเธกเธนเธฅเน€เธเนเธ format เธ—เธตเนเนเธเนเธเธฒเธเธเนเธฒเธข
      const reportData = filteredRows.map((row, index) => {
        // เนเธเน index เนเธ—เธเน€เธเธทเนเธญเธเธเธฒเธ sheet เนเธกเนเธกเธต header
        const employee = row._rawData[0] || '';        // column 0: เธเธทเนเธญเธเธเธฑเธเธเธฒเธ
        const lineName = row._rawData[1] || '';        // column 1: Line name
        const clockIn = row._rawData[3] || '';         // column 3: เน€เธงเธฅเธฒเน€เธเนเธฒ
        const clockOut = row._rawData[5] || '';        // column 5: เน€เธงเธฅเธฒเธญเธญเธ
        const userInfo = row._rawData[4] || '';        // column 4: userinfo/เธซเธกเธฒเธขเน€เธซเธ•เธธ (เนเธเนเนเธ—เธเธซเธกเธฒเธขเน€เธซเธ•เธธ)
        const location = row._rawData[6] || '';        // column 6: เธเธดเธเธฑเธ”
        const locationName = row._rawData[7] || '';    // column 7: เธชเธ–เธฒเธเธ—เธตเนเน€เธเนเธฒ
        const locationOutCoords = row._rawData[8] || ''; // column 8: เธเธดเธเธฑเธ”เธญเธญเธ
        const locationOut = row._rawData[9] || '';     // column 9: เธ—เธตเนเธญเธขเธนเนเธญเธญเธ
        const workingHours = row._rawData[10] || '';   // column 10: เธเธฑเนเธงเนเธกเธเธ—เธณเธเธฒเธ
        const note = row._rawData[4] || '';            // column 4: เธซเธกเธฒเธขเน€เธซเธ•เธธ (เน€เธเธฅเธตเนเธขเธเธเธฒเธ 11 เน€เธเนเธ 4)

        // Debug: เนเธชเธ”เธเธเนเธญเธกเธนเธฅเนเธ•เนเธฅเธฐ row
        if (index < 3) {
          console.log(`๐“ Row ${index + 1} data:`, {
            employee: employee,
            clockIn: clockIn,
            clockOut: clockOut,
            lineName: lineName,
            userInfo: userInfo,
            location: location,
            locationName: locationName,
            locationOut: locationOut,
            workingHours: workingHours,
            note: note,
            allData: row._rawData
          });
        }

        return {
          no: index + 1,
          employee: employee,
          lineName: lineName,
          clockIn: clockIn,
          clockOut: clockOut,
          note: note,
          workingHours: workingHours,
          locationIn: locationName,
          locationOut: locationOut,
          userInfo: userInfo
        };
      });

      console.log(`โ… Report data prepared successfully: ${reportData.length} records`);
      return reportData;

    } catch (error) {
      console.error('โ Error getting report data:', error);
      throw error;
    }
  }

  async clockIn(data) {
    try {
      const { employee, userinfo, lat, lon, line_name, line_picture, mock_time } = data;

      console.log(`โฐ Clock In request for: "${employee}"`);
      if (mock_time) {
        console.log(`๐งช Using mock time: ${mock_time}`);
      }

      const employeeStatus = await this.getEmployeeStatus(employee);

      if (employeeStatus.isOnWork) {
        console.log(`โ Employee "${employee}" is already clocked in`);
        return {
          success: false,
          message: 'เธเธธเธ“เธฅเธเน€เธงเธฅเธฒเน€เธเนเธฒเธเธฒเธเนเธเนเธฅเนเธง เธเธฃเธธเธ“เธฒเธฅเธเน€เธงเธฅเธฒเธญเธญเธเธเนเธญเธ',
          employee,
          currentStatus: 'clocked_in',
          clockInTime: employeeStatus.workRecord?.clockIn
        };
      }

      // เนเธเน mock_time เธซเธฒเธเธกเธตเธเธฒเธฃเธชเนเธเธกเธฒ เนเธกเนเน€เธเนเธเธเธฑเนเธเนเธเนเน€เธงเธฅเธฒเธเธฑเธเธเธธเธเธฑเธ
      const timestamp = mock_time
        ? moment(mock_time).tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss')
        : moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss');

      // เนเธเธฅเธเธเธดเธเธฑเธ”เน€เธเนเธเธเธทเนเธญเธชเธ–เธฒเธเธ—เธตเน
      const locationName = await this.getLocationName(lat, lon);
      console.log(`๐“ Location: ${locationName}`);

      console.log(`โ… Proceeding with clock in for "${employee}"`);

      const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);

      const newRow = await mainSheet.addRow([
        employee,
        line_name,
        `=IMAGE("${line_picture}")`,
        timestamp,
        userinfo || '',
        '',
        `${lat},${lon}`,
        locationName,
        '',
        '',
        ''
      ]);

      const mainRowIndex = newRow.rowNumber;
      console.log(`โ… Added to MAIN sheet at row: ${mainRowIndex}`);

      const onWorkSheet = await this.getSheet(CONFIG.SHEETS.ON_WORK);
      await onWorkSheet.addRow([
        timestamp,
        employee,
        timestamp,
        'เธ—เธณเธเธฒเธ',
        userinfo || '',
        `${lat},${lon}`,
        locationName,
        mainRowIndex,
        line_name,
        line_picture,
        mainRowIndex,
        employee
      ]);      // Clear cache เน€เธเธทเนเธญเธเธเธฒเธเธกเธตเธเธฒเธฃเน€เธเธดเนเธกเธเนเธญเธกเธนเธฅเนเธซเธกเน
      this.clearCache('onwork');
      this.clearCache('main');
      this.clearCache('stats');

      console.log(`โ… Clock In successful: ${employee} at ${this.formatTime(timestamp)}, Main row: ${mainRowIndex}`);

      // เธ—เธณเธเธฒเธฃ warm cache เธญเธฑเธ•เนเธเธกเธฑเธ•เธด
      setTimeout(async () => {
        try {
          await this.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
          await this.getAdminStats();
        } catch (error) {
          console.error('โ ๏ธ Auto cache warming error:', error);
        }
      }, 2000);

      this.triggerMapGeneration('clockin', {
        employee, lat, lon, line_name, userinfo, timestamp
      });

      return {
        success: true,
        message: 'เธเธฑเธเธ—เธถเธเน€เธงเธฅเธฒเน€เธเนเธฒเธเธฒเธเธชเธณเน€เธฃเนเธ',
        employee,
        time: this.formatTime(timestamp),
        currentStatus: 'clocked_in'
      };

    } catch (error) {
      console.error('โ Clock in error:', error);
      return {
        success: false,
        message: `เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”: ${error.message}`,
        employee: data.employee
      };
    }
  }

  async clockOut(data) {
    try {
      const { employee, lat, lon, line_name, mock_time } = data;

      console.log(`โฐ Clock Out request for: "${employee}"`);
      console.log(`๐“ Location: ${lat}, ${lon}`);
      if (mock_time) {
        console.log(`๐งช Using mock time: ${mock_time}`);
      }

      const employeeStatus = await this.getEmployeeStatus(employee);
      if (!employeeStatus.isOnWork) {
        console.log(`โ Employee "${employee}" is not clocked in`);

        // เนเธเน cached data เนเธ—เธเธเธฒเธฃเน€เธฃเธตเธขเธ API เนเธซเธกเน
        const rows = await this.getCachedSheetData(CONFIG.SHEETS.ON_WORK);

        const suggestions = rows
          .map(row => ({
            systemName: row.get('เธเธทเนเธญเนเธเธฃเธฐเธเธ'),
            employeeName: row.get('เธเธทเนเธญเธเธเธฑเธเธเธฒเธ')
          }))
          .filter(emp => emp.systemName || emp.employeeName)
          .filter(emp =>
            this.isNameMatch(employee, emp.systemName) ||
            this.isNameMatch(employee, emp.employeeName)
          );

        let message = 'เธเธธเธ“เธ•เนเธญเธเธฅเธเน€เธงเธฅเธฒเน€เธเนเธฒเธเธฒเธเธเนเธญเธ เธซเธฃเธทเธญเธ•เธฃเธงเธเธชเธญเธเธเธทเนเธญเธ—เธตเนเธเนเธญเธเนเธซเนเธ–เธนเธเธ•เนเธญเธ';

        if (suggestions.length > 0) {
          const suggestedNames = suggestions.map(s => s.systemName || s.employeeName);
          message = `เนเธกเนเธเธเธเนเธญเธกเธนเธฅเธเธฒเธฃเธฅเธเน€เธงเธฅเธฒเน€เธเนเธฒเธเธฒเธ เธเธทเนเธญเธ—เธตเนเนเธเธฅเนเน€เธเธตเธขเธ: ${suggestedNames.join(', ')}`;
        }

        return {
          success: false,
          message: message,
          employee,
          currentStatus: 'not_clocked_in',
          suggestions: suggestions.length > 0 ? suggestions : undefined
        };
      }

      // เนเธเน mock_time เธซเธฒเธเธกเธตเธเธฒเธฃเธชเนเธเธกเธฒ เนเธกเนเน€เธเนเธเธเธฑเนเธเนเธเนเน€เธงเธฅเธฒเธเธฑเธเธเธธเธเธฑเธ
      const timestamp = mock_time
        ? moment(mock_time).tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss')
        : moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss');
      const workRecord = employeeStatus.workRecord;
      const clockInTime = workRecord.clockIn;
      console.log(`โฐ Clock in time: ${clockInTime}`);

      // ๐ฏ เนเธเนเธเธฑเธเธเนเธเธฑเธเธเธณเธเธงเธ“เน€เธงเธฅเธฒเนเธเธเน€เธ”เธตเธขเธงเธเธฑเธเธเธฑเธ admin stats
      const hoursWorked = calculateWorkingHours(clockInTime, timestamp);
      console.log(`โ… Working hours calculated: ${hoursWorked.toFixed(2)} hours`);

      // เนเธเธฅเธเธเธดเธเธฑเธ”เน€เธเนเธเธเธทเนเธญเธชเธ–เธฒเธเธ—เธตเน
      const locationName = await this.getLocationName(lat, lon);
      console.log(`๐“ Clock out location: ${locationName}`); console.log(`โ… Proceeding with clock out for "${employee}"`);

      // เนเธเน cached data เนเธ—เธเธเธฒเธฃเน€เธฃเธตเธขเธ API เนเธซเธกเน
      const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
      const rows = await this.getCachedSheetData(CONFIG.SHEETS.MAIN);

      console.log(`๐“ Total rows in MAIN: ${rows.length}`);
      console.log(`๐ฏ Target row index: ${workRecord.mainRowIndex}`);

      let mainRow = null;

      if (workRecord.mainRowIndex && workRecord.mainRowIndex > 1) {
        const targetIndex = workRecord.mainRowIndex - 2;

        if (targetIndex >= 0 && targetIndex < rows.length) {
          const candidateRow = rows[targetIndex];
          const candidateEmployee = candidateRow.get('เธเธทเนเธญเธเธเธฑเธเธเธฒเธ');

          if (this.isNameMatch(employee, candidateEmployee)) {
            mainRow = candidateRow;
            console.log(`โ… Found main row by index: ${targetIndex} (row ${workRecord.mainRowIndex})`);
          } else {
            console.log(`โ ๏ธ Row index found but employee name mismatch: "${candidateEmployee}" vs "${employee}"`);
          }
        } else {
          console.log(`โ ๏ธ Row index out of range: ${targetIndex} (total rows: ${rows.length})`);
        }
      }

      if (!mainRow) {
        console.log('๐” Searching by employee name and conditions...');

        const candidateRows = rows.filter(row => {
          const rowEmployee = row.get('เธเธทเนเธญเธเธเธฑเธเธเธฒเธ');
          const rowClockOut = row.get('เน€เธงเธฅเธฒเธญเธญเธ');

          return this.isNameMatch(employee, rowEmployee) && !rowClockOut;
        });

        console.log(`Found ${candidateRows.length} candidate rows without clock out`);

        if (candidateRows.length === 1) {
          mainRow = candidateRows[0];
          console.log(`โ… Found unique candidate row`);
        } else if (candidateRows.length > 1) {
          let closestRow = null;
          let minTimeDiff = Infinity;

          candidateRows.forEach((row, index) => {
            const rowClockIn = row.get('เน€เธงเธฅเธฒเน€เธเนเธฒ');
            if (rowClockIn && clockInTime) {
              const timeDiff = Math.abs(new Date(rowClockIn) - new Date(clockInTime));
              console.log(`Candidate ${index}: time diff = ${timeDiff}ms`);
              if (timeDiff < minTimeDiff) {
                minTimeDiff = timeDiff;
                closestRow = row;
              }
            }
          });

          if (closestRow && minTimeDiff < 300000) {
            mainRow = closestRow;
            console.log(`โ… Found closest matching row (time diff: ${minTimeDiff}ms)`);
          } else {
            console.log(`โ No close time match found (min diff: ${minTimeDiff}ms)`);
          }
        }
      }

      if (!mainRow) {
        console.log('๐” Searching for latest row of this employee...');

        for (let i = rows.length - 1; i >= 0; i--) {
          const row = rows[i];
          const rowEmployee = row.get('เธเธทเนเธญเธเธเธฑเธเธเธฒเธ');
          const rowClockOut = row.get('เน€เธงเธฅเธฒเธญเธญเธ');

          if (this.isNameMatch(employee, rowEmployee) && !rowClockOut) {
            mainRow = row;
            console.log(`โ… Found latest uncompleted row at index: ${i}`);
            break;
          }
        }
      }

      if (!mainRow) {
        console.log('โ Cannot find main row to update');

        return {
          success: false,
          message: 'เนเธกเนเธเธเธเนเธญเธกเธนเธฅเธเธฒเธฃเธฅเธเน€เธงเธฅเธฒเน€เธเนเธฒเธเธฒเธเธ—เธตเนเธ•เธฃเธเธเธฑเธ เธเธฃเธธเธ“เธฒเธ•เธฃเธงเธเธชเธญเธเธฃเธฐเธเธ',
          employee
        };
      }

      console.log('โ… Found main row, updating...');

      try {
        // ๐”ง เนเธเนเธงเธดเธเธต batch update เน€เธเธทเนเธญเธเนเธญเธเธเธฑเธเธเธฒเธฃเน€เธเธฅเธตเนเธขเธเธฃเธนเธเนเธเธเน€เธงเธฅเธฒเน€เธเนเธฒ
        const sheet = await this.getSheet(CONFIG.SHEETS.MAIN);
        const rowNumber = mainRow.rowNumber;

        console.log(`๐“ Updating row ${rowNumber} using batch update to preserve format`);

        // เธญเธฑเธเน€เธ”เธ•เน€เธเธเธฒเธฐเน€เธเธฅเธฅเนเธ—เธตเนเธเธณเน€เธเนเธ เนเธ”เธขเนเธกเนเนเธ•เธฐเน€เธเธฅเธฅเนเน€เธงเธฅเธฒเน€เธเนเธฒ (column D)
        const updates = [];

        // Column F: เน€เธงเธฅเธฒเธญเธญเธ (index 5)
        updates.push({
          range: `F${rowNumber}`,
          values: [[timestamp]]
        });

        // Column I: เธเธดเธเธฑเธ”เธญเธญเธ (index 8) 
        updates.push({
          range: `I${rowNumber}`,
          values: [[`${lat},${lon}`]]
        });

        // Column J: เธ—เธตเนเธญเธขเธนเนเธญเธญเธ (index 9)
        updates.push({
          range: `J${rowNumber}`,
          values: [[locationName]]
        });

        // Column K: เธเธฑเนเธงเนเธกเธเธ—เธณเธเธฒเธ (index 10)
        updates.push({
          range: `K${rowNumber}`,
          values: [[hoursWorked.toFixed(2)]]
        });

        // เธ—เธณเธเธฒเธฃเธญเธฑเธเน€เธ”เธ•เธ—เธตเธฅเธฐเน€เธเธฅเธฅเน
        for (const update of updates) {
          await sheet.loadCells(update.range);
          const cell = sheet.getCellByA1(update.range);

          // เน€เธเนเธ•เธเนเธฒเน€เธเธเธฒเธฐเธเนเธญเธกเธนเธฅ เนเธกเนเธ•เธฑเนเธเธเนเธฒ format เนเธซเน Google Sheets เธเธฑเธ”เธเธฒเธฃเน€เธญเธ
          cell.value = update.values[0][0];
        }

        await sheet.saveUpdatedCells();
        console.log('โ… Main row updated successfully using batch update (clock-in format preserved)');

      } catch (updateError) {
        console.error('โ Error updating main row:', updateError);
        throw new Error('เนเธกเนเธชเธฒเธกเธฒเธฃเธ–เธญเธฑเธเน€เธ”เธ•เธเนเธญเธกเธนเธฅเนเธ”เน: ' + updateError.message);
      } try {
        await workRecord.row.delete();
        console.log('โ… Removed from ON_WORK sheet');
        // Clear cache เน€เธเธทเนเธญเธเธเธฒเธเธกเธตเธเธฒเธฃเน€เธเธฅเธตเนเธขเธเนเธเธฅเธเธเนเธญเธกเธนเธฅ
        this.clearCache('onwork');
        this.clearCache('main');
        this.clearCache('stats');

        // เธ—เธณเธเธฒเธฃ warm cache เธญเธฑเธ•เนเธเธกเธฑเธ•เธด
        setTimeout(async () => {
          try {
            await this.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
            await this.getAdminStats();
          } catch (error) {
            console.error('โ ๏ธ Auto cache warming error:', error);
          }
        }, 2000);

      } catch (deleteError) {
        console.error('โ Error deleting from ON_WORK:', deleteError);
      }

      console.log(`โ… Clock Out successful: ${employee} at ${this.formatTime(timestamp)} (${hoursWorked.toFixed(2)} hours)`);

      try {
        this.triggerMapGeneration('clockout', {
          employee, lat, lon, line_name, timestamp, hoursWorked
        });
      } catch (webhookError) {
        console.error('โ ๏ธ Webhook error (non-critical):', webhookError);
      }

      return {
        success: true,
        message: 'เธเธฑเธเธ—เธถเธเน€เธงเธฅเธฒเธญเธญเธเธเธฒเธเธชเธณเน€เธฃเนเธ',
        employee,
        time: this.formatTime(timestamp),
        hours: hoursWorked.toFixed(2),
        currentStatus: 'clocked_out'
      };

    } catch (error) {
      console.error('โ Clock out error:', error);

      return {
        success: false,
        message: `เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”: ${error.message}`,
        employee: data.employee
      };
    }
  }

  async triggerMapGeneration(action, data) {
    try {
      console.log(`๐”” triggerMapGeneration called: ${action} for ${data.employee}`);

      const gsaWebhookUrl = process.env.GSA_MAP_WEBHOOK_URL;
      console.log(`๐”— GSA URL: ${gsaWebhookUrl ? 'Configured' : 'NOT CONFIGURED'}`);

      if (!gsaWebhookUrl) {
        console.log('โ ๏ธ GSA webhook URL not configured');
        return;
      }

      const payload = {
        action,
        data,
        timestamp: moment().tz(CONFIG.TIMEZONE).toISOString()
      };

      console.log(`๐“ค Sending to GSA: ${JSON.stringify(payload).substring(0, 100)}...`);

      const response = await fetch(gsaWebhookUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
      });

      const responseText = await response.text();
      console.log(`๐“ GSA Response: ${response.status} - ${responseText.substring(0, 100)}`);

    } catch (error) {
      console.error('โ Error triggering map generation:', error.message);
    }
  } formatTime(date) {
    try {
      // เธฃเธญเธเธฃเธฑเธเธ—เธฑเนเธ Date object เนเธฅเธฐ string
      if (typeof date === 'string') {
        // เธ–เนเธฒเน€เธเนเธเธฃเธนเธเนเธเธ 'YYYY-MM-DD HH:mm:ss' เธเธฒเธ moment
        if (date.includes(' ') && date.length === 19) {
          return date.split(' ')[1]; // เนเธเนเธชเนเธงเธเน€เธงเธฅเธฒเน€เธ—เนเธฒเธเธฑเนเธ
        }
        // เธฅเธญเธเนเธเธฅเธเน€เธเนเธ Date object
        const parsedDate = moment(date).tz(CONFIG.TIMEZONE);
        if (parsedDate.isValid()) {
          return parsedDate.format('HH:mm:ss');
        }
        return date; // เธ–เนเธฒเนเธเธฅเธเนเธกเนเนเธ”เน เธชเนเธเธเธฅเธฑเธเน€เธเนเธ string เน€เธ”เธดเธก
      }

      // เธ–เนเธฒเน€เธเนเธ Date object
      if (date instanceof Date && !isNaN(date.getTime())) {
        return moment(date).tz(CONFIG.TIMEZONE).format('HH:mm:ss');
      }

      return '';
    } catch (error) {
      console.error('Error formatting time:', error);
      return date?.toString() || '';
    }
  }

  // เน€เธเธดเนเธกเธเธฑเธเธเนเธเธฑเธเนเธเธฅเธเธเธดเธเธฑเธ”เน€เธเนเธเธเธทเนเธญเธชเธ–เธฒเธเธ—เธตเน
  async getLocationName(lat, lon) {
    try {
      // เนเธเน OpenStreetMap Nominatim API (เธเธฃเธต)
      const response = await fetch(
        `https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lon}&zoom=18&addressdetails=1&accept-language=th`
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();

      if (data && data.display_name) {
        // เนเธเนเธเธทเนเธญเธชเธ–เธฒเธเธ—เธตเนเธ—เธตเนเนเธ”เนเธเธฒเธ API
        return data.display_name;
      } else {
        // เธ–เนเธฒเนเธกเนเนเธ”เนเธเนเธญเธกเธนเธฅ เนเธเนเธเธดเธเธฑเธ”เนเธ—เธ
        return `${lat}, ${lon}`;
      }
    } catch (error) {
      console.warn(`โ ๏ธ Location lookup failed for ${lat}, ${lon}:`, error.message);
      // เธ–เนเธฒเธเธดเธ”เธเธฅเธฒเธ” เนเธเนเธเธดเธเธฑเธ”เนเธ—เธ
      return `${lat}, ${lon}`;
    }
  }

  // เธเธฑเธเธเนเธเธฑเธเธชเธณเธซเธฃเธฑเธเธ•เธฃเธงเธเธชเธญเธเนเธฅเธฐเธเธฑเธ”เธเธฒเธฃเธเธฃเธ“เธตเธฅเธทเธกเธฅเธเน€เธงเธฅเธฒเธญเธญเธ

  async checkAndHandleMissedCheckouts() {
    try {
      console.log('๐” Starting automatic missed checkout check...');

      // เนเธเน SQLite เน€เธเนเธเนเธซเธฅเนเธเธเนเธญเธกเธนเธฅเธซเธฅเธฑเธเนเธซเนเธ•เธฃเธเธเธฑเธ clock in/out
      const onWorkRows = sqliteService.getOnWorkEmployees();

      if (onWorkRows.length === 0) {
        console.log('โ… No employees currently on work, no missed checkouts to handle');
        return { success: true, processedCount: 0, message: 'No employees on work' };
      }

      console.log(`๐“ Found ${onWorkRows.length} employees currently on work`);

      const today = moment().tz(CONFIG.TIMEZONE);
      const cutoffTime = today.clone().set({
        hour: CONFIG.AUTO_CHECKOUT.CUTOFF_HOUR,
        minute: CONFIG.AUTO_CHECKOUT.CUTOFF_MINUTE,
        second: 59,
        millisecond: 999
      });

      console.log(`โฐ Processing missed checkouts for cutoff time: ${cutoffTime.format('YYYY-MM-DD HH:mm:ss')}`);
      console.log(`๐ก๏ธ Exempt employees: ${CONFIG.AUTO_CHECKOUT.EXEMPT_EMPLOYEES.join(', ')}`);

      let processedCount = 0;
      let exemptedCount = 0;
      const results = [];

      // เธเธฃเธฐเธกเธงเธฅเธเธฅเนเธ•เนเธฅเธฐเธเธเธฑเธเธเธฒเธเธ—เธตเนเธขเธฑเธเนเธกเนเนเธ”เนเธฅเธเน€เธงเธฅเธฒเธญเธญเธ
      for (const workRow of onWorkRows) {
        try {
          const employeeName = workRow.employee_name || workRow.system_name;
          const clockInTime = workRow.clock_in;
          const mainRowIndex = workRow.main_row_id;

          if (!employeeName || !clockInTime) {
            console.warn(`โ ๏ธ Missing data for work record: ${employeeName || 'Unknown'}`);
            continue;
          }

          // ๐ก๏ธ เธ•เธฃเธงเธเธชเธญเธเธงเนเธฒเน€เธเนเธเธเธเธฑเธเธเธฒเธเธ—เธตเนเธขเธเน€เธงเนเธเธซเธฃเธทเธญเนเธกเน
          const isExempt = this.isEmployeeExempt(employeeName);
          if (isExempt) {
            console.log(`๐ก๏ธ EXEMPT: ${employeeName} - skipping auto checkout (night guard)`);
            exemptedCount++;
            results.push({
              employee: employeeName,
              action: 'exempted',
              reason: 'Night guard - exempt from auto checkout',
              clockIn: clockInTime
            });
            continue;
          }

          // เธ•เธฃเธงเธเธชเธญเธเธงเนเธฒเน€เธงเธฅเธฒเน€เธเนเธฒเน€เธเนเธเธงเธฑเธเธเธตเนเธซเธฃเธทเธญเนเธกเน
          const clockInMoment = this.parseDateTime(clockInTime);
          const isToday = clockInMoment.format('YYYY-MM-DD') === today.format('YYYY-MM-DD');

          if (!isToday) {
            console.log(`โญ๏ธ Skipping ${employeeName} - not clocked in today (${clockInMoment.format('YYYY-MM-DD')})`);
            continue;
          }

          console.log(`๐” Processing missed checkout for: ${employeeName}`);
          console.log(`โฐ Clock in time: ${clockInTime}`);
          console.log(`๐“ Main row index: ${mainRowIndex}`);

          // เธญเธฑเธเน€เธ”เธ• MAIN sheet เธ”เนเธงเธขเธเนเธญเธกเธนเธฅเธฅเธทเธกเธฅเธเน€เธงเธฅเธฒเธญเธญเธ
          const result = await this.processMissedCheckout({
            employeeName,
            clockInTime,
            mainRowIndex,
            cutoffTime
          });

          if (result.success) {
            processedCount++;
            results.push({
              employee: employeeName,
              action: 'missed_checkout_processed',
              clockIn: clockInTime,
              autoClockOut: result.autoClockOut
            });

            console.log(`โ… Processed missed checkout for ${employeeName}`);
          } else {
            console.error(`โ Failed to process missed checkout for ${employeeName}: ${result.error}`);
            results.push({
              employee: employeeName,
              action: 'failed',
              error: result.error
            });
          }

        } catch (error) {
          console.error(`โ Error processing missed checkout for employee:`, error);
          results.push({
            employee: workRow.employee_name || workRow.system_name || 'Unknown',
            action: 'error',
            error: error.message
          });
        }
      }

      console.log(`โ… Missed checkout check completed.`);
      console.log(`   ๐“ Total checked: ${onWorkRows.length}`);
      console.log(`   โ… Processed: ${processedCount}`);
      console.log(`   ๐ก๏ธ Exempted: ${exemptedCount}`);

      // Sync เธเธฅเธฑเธเนเธ Google Sheets เธ—เธฑเธเธ—เธตเน€เธกเธทเนเธญเธกเธตเธเธฒเธฃเธญเธฑเธเน€เธ”เธ•เธเธฒเธ auto-checkout
      if (processedCount > 0 && syncService) {
        try {
          await syncService.syncToSheets();
          await syncService.syncOnWorkToSheets();
          this.clearCache();
          console.log('โ… Synced auto-checkout updates to Google Sheets');
        } catch (syncError) {
          console.error('โ ๏ธ Failed to sync auto-checkout updates to Sheets:', syncError.message);
        }
      }

      // เธชเนเธ notification เธ–เนเธฒเธกเธตเธเธฒเธฃเธเธฃเธฐเธกเธงเธฅเธเธฅเธซเธฃเธทเธญเธเธฒเธฃเธขเธเน€เธงเนเธ (เธเธดเธ”เธเธฒเธฃเนเธเนเธเน€เธ•เธทเธญเธ - เนเธญเธ”เธกเธดเธเธ”เธนเธเธฒเธเนเธ”เธเธเธญเธฃเนเธ”เน€เธ—เนเธฒเธเธฑเนเธ)
      // if (processedCount > 0 || exemptedCount > 0) {
      //   await this.sendMissedCheckoutNotification(results, processedCount, exemptedCount);
      // }

      return {
        success: true,
        processedCount,
        exemptedCount,
        totalChecked: onWorkRows.length,
        results,
        message: `Processed ${processedCount} missed checkouts, exempted ${exemptedCount} employees`
      };

    } catch (error) {
      console.error('โ Error in checkAndHandleMissedCheckouts:', error);
      return {
        success: false,
        error: error.message,
        processedCount: 0,
        exemptedCount: 0
      };
    }
  }

  // เธเธฑเธเธเนเธเธฑเธเธชเธณเธซเธฃเธฑเธเธเธฃเธฐเธกเธงเธฅเธเธฅเธฅเธทเธกเธฅเธเน€เธงเธฅเธฒเธญเธญเธเธเธญเธเธเธเธฑเธเธเธฒเธเธเธเธซเธเธถเนเธ
  async processMissedCheckout({ employeeName, clockInTime, mainRowIndex, cutoffTime }) {
    try {
      const autoClockOutTime = cutoffTime.format('DD/MM/YYYY HH:mm:ss');
      const hoursWorked = calculateWorkingHours(clockInTime, autoClockOutTime);
      const missedCheckoutNote = 'ลืมลงเวลาออก (ระบบอัตโนมัติ)';

      console.log(`⏰ Auto clock out for ${employeeName}: ${autoClockOutTime} (${hoursWorked.toFixed(2)} hours)`);
      console.log(`📝 Note for record: "${missedCheckoutNote}"`);

      if (!mainRowIndex || isNaN(parseInt(mainRowIndex))) {
        throw new Error('Invalid main row id for SQLite time record');
      }

      const updateMain = sqliteService.db.prepare(`
        UPDATE time_records
        SET clock_out = ?, working_hours = ?, note = COALESCE(note, '') || ?, synced_to_sheets = 0
        WHERE id = ?
      `);
      const deleteOnWork = sqliteService.db.prepare('DELETE FROM on_work WHERE main_row_id = ?');
      const tx = sqliteService.db.transaction((clockOut, workingHours, note, recordId) => {
        updateMain.run(clockOut, workingHours, ` | ${note}`, recordId);
        deleteOnWork.run(recordId);
      });

      tx(autoClockOutTime, hoursWorked.toFixed(2), missedCheckoutNote, parseInt(mainRowIndex));
      console.log(`✅ Updated SQLite record #${mainRowIndex} and removed ${employeeName} from on_work`);

      return {
        success: true,
        autoClockOut: autoClockOutTime,
        hoursWorked: hoursWorked.toFixed(2),
        note: missedCheckoutNote
      };

    } catch (error) {
      console.error(`❌ Error processing missed checkout for ${employeeName}:`, error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  // ฟังก์ชันส่ง notification เมื่อมีการประมวลผลลืมลงเวลาออก
  async sendMissedCheckoutNotification(results, processedCount, exemptedCount = 0) {
    try {
      if (!CONFIG.TELEGRAM.BOT_TOKEN || !CONFIG.TELEGRAM.CHAT_ID) {
        console.log('โ ๏ธ Telegram notification not configured for missed checkout alerts');
        return;
      }

      const successfulResults = results.filter(r => r.action === 'missed_checkout_processed');
      const exemptedResults = results.filter(r => r.action === 'exempted');
      const failedResults = results.filter(r => r.action === 'failed' || r.action === 'error');

      const today = moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY');

      let message = `๐ค– *เธฃเธฒเธขเธเธฒเธเธฅเธเน€เธงเธฅเธฒเธญเธญเธเธญเธฑเธ•เนเธเธกเธฑเธ•เธด - ${today}*\n\n`;
      message += `๐“ เธชเธฃเธธเธเธเธฅ:\n`;
      message += `   โ… เธฅเธเน€เธงเธฅเธฒเธญเธญเธเธญเธฑเธ•เนเธเธกเธฑเธ•เธด: ${processedCount} เธเธ\n`;
      message += `   ๐ก๏ธ เธขเธเน€เธงเนเธ (เธขเธฒเธกเธเธฅเธฒเธเธเธทเธ): ${exemptedCount} เธเธ\n`;
      message += `   โ เนเธกเนเธชเธณเน€เธฃเนเธ: ${failedResults.length} เธเธ\n\n`;

      if (exemptedResults.length > 0) {
        message += `๐ก๏ธ *เธเธเธฑเธเธเธฒเธเธ—เธตเนเนเธ”เนเธฃเธฑเธเธเธฒเธฃเธขเธเน€เธงเนเธ:*\n`;
        exemptedResults.forEach(result => {
          const clockInTime = moment(result.clockIn).tz(CONFIG.TIMEZONE).format('HH:mm');
          message += `โ€ข ${result.employee} - เน€เธเนเธฒเธเธฒเธ ${clockInTime} (เธขเธฒเธกเธเธฅเธฒเธเธเธทเธ)\n`;
        });
        message += '\n';
      }

      if (successfulResults.length > 0) {
        message += `โ… *เธ”เธณเน€เธเธดเธเธเธฒเธฃเธชเธณเน€เธฃเนเธ:*\n`;
        successfulResults.forEach(result => {
          const clockOutTime = moment(result.autoClockOut).tz(CONFIG.TIMEZONE).format('HH:mm');
          message += `โ€ข ${result.employee} - เธฅเธเน€เธงเธฅเธฒเธญเธญเธเธญเธฑเธ•เนเธเธกเธฑเธ•เธด ${clockOutTime}\n`;
        });
        message += '\n';
      }

      if (failedResults.length > 0) {
        message += `โ *เธ”เธณเน€เธเธดเธเธเธฒเธฃเนเธกเนเธชเธณเน€เธฃเนเธ:*\n`;
        failedResults.forEach(result => {
          message += `โ€ข ${result.employee} - ${result.error}\n`;
        });
        message += '\n';
      }

      message += `โฐ เน€เธงเธฅเธฒเธเธฃเธฐเธกเธงเธฅเธเธฅ: ${moment().tz(CONFIG.TIMEZONE).format('HH:mm:ss')}\n`;
      message += `๐’ก เธเธเธฑเธเธเธฒเธเธชเธฒเธกเธฒเธฃเธ–เธฅเธเน€เธงเธฅเธฒเน€เธเนเธฒเธเธฒเธเธงเธฑเธเนเธซเธกเนเนเธ”เนเธเธเธ•เธด\n`;
      message += `๐ก๏ธ เธเธเธฑเธเธเธฒเธเธขเธเน€เธงเนเธ: ${CONFIG.AUTO_CHECKOUT.EXEMPT_EMPLOYEES.join(', ')}\n`;
      message += `๐“ เธซเธกเธฒเธขเน€เธซเธ•เธธ "เธฅเธทเธกเธฅเธเน€เธงเธฅเธฒเธญเธญเธ (เธฃเธฐเธเธเธญเธฑเธ•เนเธเธกเธฑเธ•เธด)" เธ–เธนเธเน€เธเธตเธขเธเธฅเธเธเธญเธฅเธฑเธกเธเน E เนเธ Google Sheet`;

      // เธชเนเธเธเนเธญเธเธงเธฒเธกเนเธเธขเธฑเธ Telegram
      const telegramUrl = `https://api.telegram.org/bot${CONFIG.TELEGRAM.BOT_TOKEN}/sendMessage`;

      await fetch(telegramUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          chat_id: CONFIG.TELEGRAM.CHAT_ID,
          text: message,
          parse_mode: 'Markdown'
        })
      });

      console.log('โ… Missed checkout notification sent to Telegram');

    } catch (error) {
      console.error('โ Error sending missed checkout notification:', error);
    }
  }

  // Emergency mode functions
  setEmergencyMode(enabled) {
    this.emergencyMode = enabled;
    if (enabled) {
      console.log('๐จ Emergency mode ENABLED - Using cached data only');
      // เธเธขเธฒเธข TTL เธเธญเธ cache เน€เธเนเธ 1 เธเธฑเนเธงเนเธกเธ
      Object.keys(this.cache).forEach(key => {
        this.cache[key].ttl = 3600000; // 1 hour
      });
    } else {
      console.log('โ… Emergency mode DISABLED - Normal operation resumed');
      // เธเธทเธเธเนเธฒ TTL เน€เธ”เธดเธก
      this.cache.employees.ttl = 300000; // 5 minutes
      this.cache.onwork.ttl = 60000;     // 1 minute
      this.cache.main.ttl = 30000;       // 30 seconds
      this.cache.stats.ttl = 120000;     // 2 minutes
    }
  }

  async safeGetCachedSheetData(sheetName) {
    try {
      return await this.getCachedSheetData(sheetName);
    } catch (error) {
      console.error(`โ Failed to get data for ${sheetName}:`, error.message);

      // เน€เธเนเธฒเธชเธนเน emergency mode
      if (!this.emergencyMode) {
        this.setEmergencyMode(true);
      }

      // เธเธทเธเธเนเธฒ cache เน€เธเนเธฒ (เธ–เนเธฒเธกเธต)
      const staleData = this.getCache(sheetName.toLowerCase().replace(/\s+/g, ''));
      if (staleData) {
        console.log(`๐“ Using emergency cache for ${sheetName}`);
        return staleData;
      }

      // เธ–เนเธฒเนเธกเนเธกเธต cache เน€เธฅเธข เธเธทเธเธเนเธฒ array เธงเนเธฒเธ
      console.warn(`โ ๏ธ No cache available for ${sheetName}, returning empty data`);
      return [];
    }
  }

  // ๐• เธฅเธเธฃเธฒเธขเธเธฒเธฃเน€เธงเธฅเธฒเธเธฒเธ Main Sheet
  async deleteTimeRecordFromSheet(employeeName, clockIn) {
    try {
      const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
      const rows = await mainSheet.getRows({ offset: 1 });

      // เธซเธฒเนเธ–เธงเธ—เธตเนเธ•เธฃเธเธเธฑเธ employeeName เนเธฅเธฐ clockIn
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowEmployee = row._rawData[0]; // Column A: เธเธทเนเธญเธเธเธฑเธเธเธฒเธ
        const rowClockIn = row._rawData[3];  // Column D: เน€เธงเธฅเธฒเน€เธเนเธฒ

        // เน€เธเธฃเธตเธขเธเน€เธ—เธตเธขเธเธเธทเนเธญเนเธฅเธฐเน€เธงเธฅเธฒ (เนเธเน moment เธเนเธงเธขเน€เธฃเธทเนเธญเธ format เน€เธงเธฅเธฒ 08:00 vs 8:00)
        const isNameMatch = rowEmployee === employeeName;
        let isTimeMatch = rowClockIn === clockIn;

        if (!isTimeMatch && rowClockIn && clockIn) {
          // เธฅเธญเธ parse เนเธฅเนเธงเน€เธ—เธตเธขเธ
          const mRow = moment(rowClockIn, ['DD/MM/YYYY HH:mm:ss', 'YYYY-MM-DD HH:mm:ss', 'DD/MM/YYYY H:mm:ss']);
          const mTarget = moment(clockIn, ['DD/MM/YYYY HH:mm:ss', 'YYYY-MM-DD HH:mm:ss', 'DD/MM/YYYY H:mm:ss']);

          if (mRow.isValid() && mTarget.isValid()) {
            // เน€เธ—เธตเธขเธเธฃเธฐเธ”เธฑเธเธงเธดเธเธฒเธ—เธต
            isTimeMatch = mRow.isSame(mTarget, 'second');
          }
        }

        if (isNameMatch && isTimeMatch) {
          console.log(`๐—‘๏ธ Deleting row ${i + 2} from Main Sheet: ${employeeName}`);
          await row.delete();
          this.clearCache();
          return { success: true, message: `เธฅเธเธฃเธฒเธขเธเธฒเธฃเธเธฒเธ Sheets เธชเธณเน€เธฃเนเธ` };
        }
      }
      console.log(`โ ๏ธ Record not found in Sheets: ${employeeName} - ${clockIn}`);
      return { success: false, message: 'เนเธกเนเธเธเธฃเธฒเธขเธเธฒเธฃเนเธ Sheets' };

    } catch (error) {
      console.error('โ Error deleting from Sheets:', error);
      return { success: false, error: error.message };
    }
  }

  // ๐• เธฅเธเธเธฒเธ On Work Sheet
  async deleteFromOnWorkSheet(employeeName) {
    try {
      const onWorkSheet = await this.getSheet(CONFIG.SHEETS.ON_WORK);
      const rows = await onWorkSheet.getRows({ offset: 1 });

      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowEmployee = row.get('เธเธทเนเธญเธเธเธฑเธเธเธฒเธ') || row.get('เธเธทเนเธญเนเธเธฃเธฐเธเธ');

        if (rowEmployee === employeeName) {
          console.log(`๐—‘๏ธ Deleting from On Work Sheet: ${employeeName}`);
          await row.delete();
          return { success: true };
        }
      }

      return { success: false, message: 'เนเธกเนเธเธเนเธ On Work' };
    } catch (error) {
      console.error('โ Error deleting from On Work Sheet:', error);
      return { success: false, error: error.message };
    }
  }

  // ๐• เธฅเธเธเธเธฑเธเธเธฒเธเธเธฒเธ Employees Sheet
  async deleteEmployeeFromSheet(employeeName) {
    try {
      const employeesSheet = await this.getSheet(CONFIG.SHEETS.EMPLOYEES);
      const rows = await employeesSheet.getRows({ offset: 1 });

      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowEmployee = row._rawData[0]; // Column A: เธเธทเนเธญเธเธเธฑเธเธเธฒเธ

        if (rowEmployee === employeeName) {
          console.log(`๐—‘๏ธ Deleting employee from Sheets: ${employeeName}`);
          await row.delete();
          this.clearCache();
          return { success: true, message: `เธฅเธเธเธเธฑเธเธเธฒเธเธเธฒเธ Sheets เธชเธณเน€เธฃเนเธ` };
        }
      }

      console.log(`โ ๏ธ Employee not found in Sheets: ${employeeName}`);
      return { success: false, message: 'เนเธกเนเธเธเธเธเธฑเธเธเธฒเธเนเธ Sheets' };

    } catch (error) {
      console.error('โ Error deleting employee from Sheets:', error);
      return { success: false, error: error.message };
    }
  }

  // ๐• เน€เธเธดเนเธกเธเธเธฑเธเธเธฒเธเนเธ Employees Sheet
  async addEmployeeToSheet(employeeName) {
    try {
      const employeesSheet = await this.getSheet(CONFIG.SHEETS.EMPLOYEES);
      await employeesSheet.addRow([employeeName]);
      this.clearCache();
      console.log(`โ… Added employee to Sheets: ${employeeName}`);
      return { success: true };
    } catch (error) {
      console.error('โ Error adding employee to Sheets:', error);
      return { success: false, error: error.message };
    }
  }
}

// ========== Initialize Services ==========
const sheetsService = new GoogleSheetsService();
const keepAliveService = new KeepAliveService();
// ๐• SQLite + Sync Services
const sqliteService = new SQLiteService();
let syncService = null; // เธเธฐ initialize เธซเธฅเธฑเธเธเธฒเธ sheetsService เธเธฃเนเธญเธกเนเธเนเธเธฒเธ

// ========== Admin Authentication Routes ==========

// Admin Login
app.post('/api/admin/login', async (req, res) => {
  try {
    const { username, password } = req.body;

    if (!username || !password) {
      return res.status(400).json({
        success: false,
        message: 'เธเธฃเธธเธ“เธฒเธเธฃเธญเธเธเธทเนเธญเธเธนเนเนเธเนเนเธฅเธฐเธฃเธซเธฑเธชเธเนเธฒเธ'
      });
    }

    // เธเนเธเธซเธฒเธเธนเนเนเธเน
    const user = CONFIG.ADMIN.USERS.find(u => u.username === username);
    if (!user) {
      return res.status(401).json({
        success: false,
        message: 'เธเธทเนเธญเธเธนเนเนเธเนเธซเธฃเธทเธญเธฃเธซเธฑเธชเธเนเธฒเธเนเธกเนเธ–เธนเธเธ•เนเธญเธ'
      });
    }

    // เธ•เธฃเธงเธเธชเธญเธเธฃเธซเธฑเธชเธเนเธฒเธ
    const isValidPassword = await bcrypt.compare(password, user.password);
    if (!isValidPassword) {
      return res.status(401).json({
        success: false,
        message: 'เธเธทเนเธญเธเธนเนเนเธเนเธซเธฃเธทเธญเธฃเธซเธฑเธชเธเนเธฒเธเนเธกเนเธ–เธนเธเธ•เนเธญเธ'
      });
    }

    // เธชเธฃเนเธฒเธ JWT token
    const token = jwt.sign(
      {
        id: user.id,
        username: user.username,
        role: user.role
      },
      CONFIG.ADMIN.JWT_SECRET,
      { expiresIn: CONFIG.ADMIN.JWT_EXPIRES_IN }
    );

    res.json({
      success: true,
      message: 'เน€เธเนเธฒเธชเธนเนเธฃเธฐเธเธเธชเธณเน€เธฃเนเธ',
      token,
      user: {
        id: user.id,
        username: user.username,
        name: user.name,
        role: user.role
      }
    });

  } catch (error) {
    console.error('Admin login error:', error);
    res.status(500).json({
      success: false,
      message: 'เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”เธ เธฒเธขเนเธเธฃเธฐเธเธ'
    });
  }
});

// Verify Token
app.get('/api/admin/verify-token', authenticateAdmin, (req, res) => {
  res.json({
    success: true,
    user: req.user
  });
});

// Token Refresh - เธชเธณเธซเธฃเธฑเธเธ•เนเธญเธญเธฒเธขเธธ token เธ—เธตเนเนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ
app.post('/api/admin/refresh-token', async (req, res) => {
  try {
    const authHeader = req.headers.authorization;
    const token = authHeader && authHeader.split(' ')[1];

    if (!token) {
      return res.status(401).json({
        success: false,
        error: 'Token required for refresh'
      });
    }

    // เธ•เธฃเธงเธเธชเธญเธ token เนเธกเนเธงเนเธฒเธเธฐเธซเธกเธ”เธญเธฒเธขเธธเนเธฅเนเธง
    let decoded;
    try {
      decoded = jwt.verify(token, CONFIG.ADMIN.JWT_SECRET);
    } catch (error) {
      if (error.name === 'TokenExpiredError') {
        // เธญเธเธธเธเธฒเธ•เนเธซเน refresh token เธ—เธตเนเธซเธกเธ”เธญเธฒเธขเธธเนเธกเนเน€เธเธดเธ 7 เธงเธฑเธ
        const expiredAt = new Date(error.expiredAt);
        const now = new Date();
        const daysSinceExpired = (now - expiredAt) / (1000 * 60 * 60 * 24);

        if (daysSinceExpired <= 7) {
          // เธ–เธญเธ”เธฃเธซเธฑเธช token เนเธ”เธขเนเธกเนเธ•เธฃเธงเธเธชเธญเธเธงเธฑเธเธซเธกเธ”เธญเธฒเธขเธธ
          decoded = jwt.verify(token, CONFIG.ADMIN.JWT_SECRET, { ignoreExpiration: true });
        } else {
          return res.status(401).json({
            success: false,
            error: 'Token expired too long ago. Please login again.',
            errorCode: 'TOKEN_EXPIRED_TOO_LONG'
          });
        }
      } else {
        return res.status(401).json({
          success: false,
          error: 'Invalid token',
          errorCode: 'INVALID_TOKEN'
        });
      }
    }

    // เธ•เธฃเธงเธเธชเธญเธเธงเนเธฒเธเธนเนเนเธเนเธขเธฑเธเธเธเธญเธขเธนเนเนเธเธฃเธฐเธเธ
    const user = CONFIG.ADMIN.USERS.find(u => u.id === decoded.id);
    if (!user) {
      return res.status(401).json({
        success: false,
        error: 'User no longer exists'
      });
    }

    // เธชเธฃเนเธฒเธ token เนเธซเธกเน
    const newToken = jwt.sign(
      {
        id: user.id,
        username: user.username,
        name: user.name,
        role: user.role
      },
      CONFIG.ADMIN.JWT_SECRET,
      { expiresIn: CONFIG.ADMIN.JWT_EXPIRES_IN }
    );

    console.log(`๐” Token refreshed for user: ${user.username}`);

    res.json({
      success: true,
      message: 'Token refreshed successfully',
      token: newToken,
      user: {
        id: user.id,
        username: user.username,
        name: user.name,
        role: user.role
      }
    });

  } catch (error) {
    console.error('Token refresh error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to refresh token'
    });
  }
});

// ๐• Admin Stats - เธขเนเธฒเธขเนเธเนเธเน SQLite เนเธฅเนเธง (เธ”เธน /api/admin/stats เธ”เนเธฒเธเธฅเนเธฒเธ)

// ๐• เธเนเธเธซเธฒเธฃเธฒเธขเธเธฒเธฃเน€เธงเธฅเธฒเธ•เธฒเธกเธงเธฑเธเธ—เธตเน (เธชเธณเธซเธฃเธฑเธเนเธเนเนเธเธขเนเธญเธเธซเธฅเธฑเธ)
app.get('/api/admin/time-records-by-date/:date', authenticateAdmin, (req, res) => {
  try {
    const { date } = req.params;
    const { employee } = req.query;

    console.log(`๐” Searching time records for date: ${date}, employee: ${employee || 'all'}`);

    // เธเนเธเธซเธฒเนเธเธเธฒเธเธเนเธญเธกเธนเธฅ
    let records;
    if (employee) {
      records = sqliteService.db.prepare(`
        SELECT * FROM time_records 
        WHERE clock_in LIKE ? AND employee_name = ?
        ORDER BY clock_in ASC
      `).all(`${date}%`, employee);
    } else {
      records = sqliteService.db.prepare(`
        SELECT * FROM time_records 
        WHERE clock_in LIKE ?
        ORDER BY employee_name ASC, clock_in ASC
      `).all(`${date}%`);
    }

    console.log(`๐“ Found ${records.length} records`);

    res.json({
      success: true,
      data: records
    });

  } catch (error) {
    console.error('Search records error:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Export Routes
app.get('/api/admin/export/:type', authenticateAdmin, async (req, res) => {
  try {
    const { type } = req.params;
    const params = req.query;

    // เธ•เธฃเธงเธเธชเธญเธเธเธฃเธฐเน€เธ เธ—เธฃเธฒเธขเธเธฒเธ
    if (!['daily', 'monthly', 'range'].includes(type)) {
      return res.status(400).json({
        success: false,
        error: 'Invalid report type'
      });
    }

    // Log format parameter เน€เธเธทเนเธญ debug
    console.log(`๐“ Export request: type=${type}, format=${params.format || 'default'}`);

    // เนเธเน SQLite เน€เธเนเธเนเธซเธฅเนเธเธเนเธญเธกเธนเธฅเธซเธฅเธฑเธเธชเธณเธซเธฃเธฑเธเธฃเธฒเธขเธเธฒเธ (fallback เน€เธเนเธ Google Sheets เธซเธฒเธเธกเธตเธเธฑเธเธซเธฒ)
    let reportData = [];
    try {
      reportData = sqliteService.getReportDataForExport(type, params);
      console.log(`๐“ [SQLite] Report rows: ${reportData.length}`);
    } catch (dbError) {
      console.warn('โ ๏ธ SQLite report fetch failed, fallback to Google Sheets:', dbError.message);
      reportData = await sheetsService.getReportData(type, params);
    }

    // ๐• เธ–เนเธฒเน€เธเนเธ monthly + format=dailySummary เนเธเน function เนเธซเธกเน
    let workbook;
    if (type === 'monthly' && params.format === 'dailySummary') {
      // เธ”เธถเธเธฃเธฒเธขเธเธทเนเธญเธเธเธฑเธเธเธฒเธเธ—เธฑเนเธเธซเธกเธ”เธชเธณเธซเธฃเธฑเธเธเธณเธเธงเธ“เธเธเธเธฒเธ”
      const allEmployees = sqliteService.getEmployees();
      console.log(`๐‘ฅ All employees for absent calculation: ${allEmployees.length}`);
      workbook = await ExcelExportService.createDailySummaryWorkbook(reportData, params, allEmployees);
    } else {
      // เธชเธฃเนเธฒเธเนเธเธฅเน Excel เนเธเธเธเธเธ•เธด
      workbook = await ExcelExportService.createWorkbook(reportData, type, params);
    }

    // เธ•เธฑเนเธเธเธทเนเธญเนเธเธฅเนเธ•เธฒเธก format
    let filename = 'report.xlsx';
    if (type === 'monthly' && params.format === 'dailySummary') {
      filename = 'monthly_daily_summary_report.xlsx';
    } else if (type === 'monthly' && params.format === 'detailed') {
      filename = 'monthly_detailed_report.xlsx';
    } else if (type === 'monthly') {
      filename = 'monthly_summary_report.xlsx';
    } else {
      filename = `${type}_report.xlsx`;
    }

    // เธ•เธฑเนเธเธเนเธฒ response headers
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${filename}`);

    // เธชเนเธเนเธเธฅเน
    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    console.error('Export error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to export report'
    });
  }
});

// ========== API Rate Limiting เนเธฅเธฐ Monitoring ==========
class APIMonitor {
  constructor() {
    this.apiCalls = [];
    // เธเธฃเธฑเธเน€เธเธดเนเธก rate limit เน€เธเธทเนเธญเธฃเธญเธเธฃเธฑเธ concurrent users เธกเธฒเธเธเธถเนเธ
    this.maxCallsPerMinute = 100; // เน€เธเธดเนเธกเธเธฒเธ 30 เน€เธเนเธ 100 เธเธฃเธฑเนเธเธ•เนเธญเธเธฒเธ—เธต
    this.maxCallsPerHour = 1000; // เน€เธเธดเนเธกเธเธฒเธ 300 เน€เธเนเธ 1000 เธเธฃเธฑเนเธเธ•เนเธญเธเธฑเนเธงเนเธกเธ

    // เน€เธเธดเนเธก burst allowance เธชเธณเธซเธฃเธฑเธ peak time
    this.burstLimit = 75; // เน€เธเธดเนเธกเธเธฒเธ 50 เน€เธเนเธ 75 concurrent requests
    this.currentBurst = 0;
    this.lastBurstReset = Date.now();

    // Auto-reset burst counter every 5 seconds
    setInterval(() => {
      if (this.currentBurst > 0) {
        console.log(`๐” Auto-resetting burst counter from ${this.currentBurst} to 0`);
        this.currentBurst = 0;
      }
    }, 5000); // 5 เธงเธดเธเธฒเธ—เธต
  }

  logAPICall(operation) {
    const now = new Date();
    this.apiCalls.push({
      timestamp: now,
      operation: operation
    });

    // เธฅเธ logs เธ—เธตเนเน€เธเนเธฒเน€เธเธดเธ 1 เธเธฑเนเธงเนเธกเธ
    this.apiCalls = this.apiCalls.filter(call =>
      (now - call.timestamp) < 3600000 // 1 hour
    );

    console.log(`๐“ API Call: ${operation} (Total in last hour: ${this.apiCalls.length}, Current burst: ${this.currentBurst})`);
  }

  canMakeAPICall() {
    const now = new Date();

    // เธเธฑเธเธเธณเธเธงเธ API calls เนเธเธเธฒเธ—เธตเธ—เธตเนเนเธฅเนเธง
    const callsInLastMinute = this.apiCalls.filter(call =>
      (now - call.timestamp) < 60000 // 1 minute
    ).length;

    // เธเธฑเธเธเธณเธเธงเธ API calls เนเธเธเธฑเนเธงเนเธกเธเธ—เธตเนเนเธฅเนเธง
    const callsInLastHour = this.apiCalls.length;

    // เธ•เธฃเธงเธเธชเธญเธ burst limit
    if (this.currentBurst >= this.burstLimit) {
      console.warn(`โ ๏ธ Burst limit exceeded: ${this.currentBurst}/${this.burstLimit} concurrent requests`);
      return false;
    }

    if (callsInLastMinute >= this.maxCallsPerMinute) {
      console.warn(`โ ๏ธ Rate limit exceeded: ${callsInLastMinute} calls in last minute`);
      return false;
    }

    if (callsInLastHour >= this.maxCallsPerHour) {
      console.warn(`โ ๏ธ Rate limit exceeded: ${callsInLastHour} calls in last hour`);
      return false;
    }

    // เน€เธเธดเนเธก burst counter
    this.currentBurst++;

    return true;
  }

  // เน€เธกเธทเนเธญ API call เน€เธชเธฃเนเธเนเธฅเนเธง เธฅเธ” burst counter
  finishCall() {
    if (this.currentBurst > 0) {
      this.currentBurst--;
    }
  }

  getStats() {
    const now = new Date();
    const callsInLastMinute = this.apiCalls.filter(call =>
      (now - call.timestamp) < 60000
    ).length;
    const callsInLastHour = this.apiCalls.length;

    return {
      callsInLastMinute,
      callsInLastHour,
      maxCallsPerMinute: this.maxCallsPerMinute,
      maxCallsPerHour: this.maxCallsPerHour,
      percentageUsedPerMinute: (callsInLastMinute / this.maxCallsPerMinute) * 100,
      percentageUsedPerHour: (callsInLastHour / this.maxCallsPerHour) * 100
    };
  }
}

const apiMonitor = new APIMonitor();

// ========== Original Routes (unchanged) ==========

// Home page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Health check เนเธฅเธฐ ping endpoint
app.get('/debug/sheet-info', async (req, res) => {
  try {
    console.log('๐” Debug: Getting sheet info...');

    const mainSheet = await sheetsService.getSheet(CONFIG.SHEETS.MAIN);
    const rows = await mainSheet.getRows({ limit: 5 });

    if (rows.length > 0) {
      const headers = Object.keys(rows[0]._rawData);
      const firstRowData = rows[0]._rawData;

      console.log('๐“ MAIN Sheet Headers:', headers);
      console.log('๐“ First row data:', firstRowData);

      res.json({
        sheetTitle: mainSheet.title,
        headerCount: headers.length,
        headers: headers,
        firstRowData: firstRowData,
        sampleRows: rows.map((row, index) => ({
          rowIndex: index,
          employee: row.get('เธเธทเนเธญเธเธเธฑเธเธเธฒเธ'),
          clockIn: row.get('เน€เธงเธฅเธฒเน€เธเนเธฒ'),
          clockOut: row.get('เน€เธงเธฅเธฒเธญเธญเธ'),
          rawData: row._rawData
        }))
      });
    } else {
      res.json({ error: 'No data found' });
    }

  } catch (error) {
    console.error('โ Debug sheet info error:', error);
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/health', (req, res) => {
  res.json({
    status: 'healthy',
    timestamp: moment().tz(CONFIG.TIMEZONE).toISOString(), // เนเธเนเน€เธงเธฅเธฒเนเธ—เธข
    uptime: process.uptime(),
    keepAlive: keepAliveService.getStats(),
    environment: process.env.NODE_ENV || 'development',
    config: {
      hasLiffId: !!CONFIG.LINE.LIFF_ID,
      liffIdLength: CONFIG.LINE.LIFF_ID ? CONFIG.LINE.LIFF_ID.length : 0
    }
  });
});

// Ping endpoint เธชเธณเธซเธฃเธฑเธ keep-alive
app.get('/api/ping', (req, res) => {
  res.json({
    status: 'pong',
    timestamp: moment().tz(CONFIG.TIMEZONE).toISOString(), // เนเธเนเน€เธงเธฅเธฒเนเธ—เธข
    uptime: process.uptime()
  });
});

// Webhook endpoint เธชเธณเธซเธฃเธฑเธเธฃเธฑเธ ping เธเธฒเธ GSA
app.post('/api/webhook/ping', (req, res) => {
  console.log('๐“จ Received ping from GSA'); res.json({
    status: 'received',
    timestamp: moment().tz(CONFIG.TIMEZONE).toISOString() // เนเธเนเน€เธงเธฅเธฒเนเธ—เธข
  });
});

// API เธชเธณเธซเธฃเธฑเธ Client Configuration
app.get('/api/config', (req, res) => {
  try {
    res.json({
      success: true,
      data: {
        liffId: CONFIG.LINE.LIFF_ID,
        apiUrl: CONFIG.RENDER.SERVICE_URL + '/api',
        environment: process.env.NODE_ENV || 'development',
        features: {
          keepAlive: CONFIG.RENDER.KEEP_ALIVE_ENABLED,
          liffEnabled: !!CONFIG.LINE.LIFF_ID
        }
      }
    });
  } catch (error) {
    console.error('API Error - config:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get config'
    });
  }
});

// Get employees
app.post('/api/employees', async (req, res) => {
  try {
    let employees = sqliteService.getEmployees();

    // Fallback: use Sheets only when SQLite has no employee data
    if (!employees || employees.length === 0) {
      console.warn('⚠️ SQLite employees empty, falling back to Google Sheets');
      employees = await sheetsService.getEmployees();
    }

    res.json({
      success: true,
      data: employees
    });
  } catch (error) {
    console.error('API Error - employees:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get employees'
    });
  }
});

// Clock in
app.post('/api/clockin', async (req, res) => {
  try {
    let { employee, userinfo, lat, lon, line_name, line_picture, mock_time } = req.body;
    if (employee) employee = employee.trim();

    if (!employee || !lat || !lon) {
      return res.status(400).json({
        success: false,
        error: 'Missing required fields'
      });
    }

    // เธ•เธฃเธงเธเธชเธญเธ rate limit
    if (!apiMonitor.canMakeAPICall()) {
      return res.status(429).json({
        success: false,
        error: 'Too many requests, please try again later'
      });
    }

    // ๐• เนเธเน SQLite เนเธ—เธ Sheets (เน€เธฃเนเธงเธเธงเนเธฒ เนเธกเนเธกเธต rate limit)
    const result = await sqliteService.clockIn({
      employee, userinfo, lat, lon, line_name, line_picture, mock_time
    });

    // เธชเนเธ webhook เนเธ GSA (เนเธกเนเธเธฃเธฐเธ—เธ response เธซเธฅเธฑเธ)
    if (result?.success) {
      try {
        await sheetsService.triggerMapGeneration('clockin', {
          employee,
          lat,
          lon,
          line_name,
          userinfo,
          timestamp: result.time || undefined
        });
      } catch (webhookError) {
        console.error('โ ๏ธ GSA webhook (clockin) failed:', webhookError.message);
      }
    }

    res.json(result);

  } catch (error) {
    // เธฅเธ” burst counter เธ–เธถเธเนเธกเนเธเธฐ error
    apiMonitor.finishCall();
    console.error('API Error - clockin:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to clock in'
    });
  }
});

// Clock out
app.post('/api/clockout', async (req, res) => {
  try {
    let { employee, lat, lon, line_name, mock_time } = req.body;
    if (employee) employee = employee.trim();

    // ๐”ง FIX: เธ•เธฃเธงเธเธชเธญเธ lat/lon เธ”เนเธงเธข typeof เน€เธเธฃเธฒเธฐ 0 เน€เธเนเธ valid value
    if (!employee || lat === undefined || lat === null || lon === undefined || lon === null) {
      return res.status(400).json({
        success: false,
        error: 'Missing required fields'
      });
    }

    // เธ•เธฃเธงเธเธชเธญเธ rate limit
    if (!apiMonitor.canMakeAPICall()) {
      return res.status(429).json({
        success: false,
        error: 'Too many requests, please try again later'
      });
    }

    // ๐• เนเธเน SQLite เนเธ—เธ Sheets (เน€เธฃเนเธงเธเธงเนเธฒ เนเธกเนเธกเธต rate limit)
    const result = await sqliteService.clockOut({
      employee, lat, lon, line_name, mock_time
    });

    // เธชเนเธ webhook เนเธ GSA (เนเธกเนเธเธฃเธฐเธ—เธ response เธซเธฅเธฑเธ)
    if (result?.success) {
      try {
        await sheetsService.triggerMapGeneration('clockout', {
          employee,
          lat,
          lon,
          line_name,
          timestamp: result.time || undefined,
          hoursWorked: result.hoursWorked
        });
      } catch (webhookError) {
        console.error('โ ๏ธ GSA webhook (clockout) failed:', webhookError.message);
      }
    }

    res.json(result);

  } catch (error) {
    // เธฅเธ” burst counter เธ–เธถเธเนเธกเนเธเธฐ error
    apiMonitor.finishCall();
    console.error('API Error - clockout:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to clock out'
    });
  }
});

// API เธชเธณเธซเธฃเธฑเธเธ•เธฃเธงเธเธชเธญเธเธชเธ–เธฒเธเธฐเธเธเธฑเธเธเธฒเธ
app.post('/api/check-status', async (req, res) => {
  try {
    const { employee } = req.body;

    if (!employee) {
      return res.status(400).json({
        success: false,
        error: 'Missing employee name'
      });
    }

    // ๐• เนเธเน SQLite เนเธ—เธ Sheets
    const employeeStatus = sqliteService.getEmployeeStatus(employee);
    const onWorkEmployees = sqliteService.getOnWorkEmployees();

    const currentEmployees = onWorkEmployees.map(row => ({
      systemName: row.system_name,
      employeeName: row.employee_name,
      clockIn: row.clock_in,
      mainRowIndex: row.main_row_id
    }));

    res.json({
      success: true,
      data: {
        employee: employee,
        isOnWork: employeeStatus.isOnWork,
        hasWorkRecord: !!employeeStatus.workRecord,
        workRecord: employeeStatus.workRecord ? {
          clockIn: employeeStatus.workRecord.clockIn,
          mainRowIndex: employeeStatus.workRecord.mainRowId
        } : null,
        allCurrentEmployees: currentEmployees
      }
    });

  } catch (error) {
    console.error('API Error - check-status:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to check status'
    });
  }
});

// API Monitoring endpoint
app.get('/api/admin/api-stats', authenticateAdmin, (req, res) => {
  const stats = apiMonitor.getStats();
  res.json({
    success: true,
    data: stats
  });
});

// ๐• Dashboard Stats endpoint - เนเธเน SQLite
app.get('/api/admin/stats', authenticateAdmin, (req, res) => {
  try {
    const stats = sqliteService.getAdminStats();
    res.json({
      success: true,
      data: stats
    });
  } catch (error) {
    console.error('Error getting admin stats:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// ๐• Settings API - GET settings from process.env
app.get('/api/admin/settings', authenticateAdmin, (req, res) => {
  try {
    const settings = {
      // App Config
      PORT: process.env.PORT || '3000',
      TIMEZONE: process.env.TIMEZONE || 'Asia/Bangkok',

      // Render/Service URLs
      RENDER_SERVICE_URL: process.env.RENDER_SERVICE_URL || '',
      RENDER_EXTERNAL_HOSTNAME: process.env.RENDER_EXTERNAL_HOSTNAME || '',
      RENDER_KEEP_ALIVE_ENABLED: process.env.RENDER_KEEP_ALIVE_ENABLED || 'true',

      // Google Sheets
      GOOGLE_SHEETS_SPREADSHEET_ID: process.env.GOOGLE_SHEETS_SPREADSHEET_ID || '',
      GOOGLE_SERVICE_ACCOUNT_EMAIL: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || '',
      // Don't expose private key for security
      GOOGLE_PRIVATE_KEY: process.env.GOOGLE_PRIVATE_KEY ? '***HIDDEN***' : '',

      // LINE Config
      LIFF_ID: process.env.LIFF_ID || '',
      // Telegram Config
      TELEGRAM_BOT_TOKEN: process.env.TELEGRAM_BOT_TOKEN ? '***HIDDEN***' : '',
      TELEGRAM_CHAT_ID: process.env.TELEGRAM_CHAT_ID || '',

      // GSA Webhook
      GSA_MAP_WEBHOOK_URL: process.env.GSA_MAP_WEBHOOK_URL || '',
      GSA_WEBHOOK_SECRET: process.env.GSA_WEBHOOK_SECRET || '',
      WEBHOOK_SECRET_KEY: process.env.WEBHOOK_SECRET_KEY || ''
    };

    res.json({
      success: true,
      settings
    });
  } catch (error) {
    console.error('Error getting settings:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// ๐• Settings API - POST to save settings (SQLite + .env file)
app.post('/api/admin/settings', authenticateAdmin, async (req, res) => {
  try {
    const newSettings = req.body;

    // Read current .env file
    const envPath = path.join(__dirname, '.env');
    let envContent = '';

    try {
      envContent = require('fs').readFileSync(envPath, 'utf8');
    } catch (e) {
      console.log('โ ๏ธ .env file not found, will create new one');
      envContent = '';
    }

    // Parse existing .env content
    const envLines = envContent.split('\n');
    const envVars = {};

    for (const line of envLines) {
      const trimmed = line.trim();
      if (trimmed && !trimmed.startsWith('#')) {
        const eqIndex = trimmed.indexOf('=');
        if (eqIndex > 0) {
          const key = trimmed.substring(0, eqIndex);
          const value = trimmed.substring(eqIndex + 1);
          envVars[key] = value;
        }
      }
    }

    // Update with new settings (skip hidden values)
    for (const [key, value] of Object.entries(newSettings)) {
      if (value && !value.includes('***HIDDEN***')) {
        envVars[key] = value;
        // Also update process.env in memory
        process.env[key] = value;
      }
    }

    // Write back to .env file
    let newEnvContent = '# Time Tracker Environment Variables\n';
    newEnvContent += '# Last updated: ' + new Date().toISOString() + '\n\n';

    for (const [key, value] of Object.entries(envVars)) {
      // Handle multiline values (like GOOGLE_PRIVATE_KEY)
      if (value.includes('\n')) {
        newEnvContent += `${key}="${value.replace(/"/g, '\\"')}"\n`;
      } else {
        newEnvContent += `${key}=${value}\n`;
      }
    }

    require('fs').writeFileSync(envPath, newEnvContent, 'utf8');

    console.log('โ… Settings saved to .env file');

    res.json({
      success: true,
      message: 'Settings เธเธฑเธเธ—เธถเธเน€เธฃเธตเธขเธเธฃเนเธญเธข (เธเธฒเธเธเนเธฒเธเธฐเธกเธตเธเธฅเธซเธฅเธฑเธ Restart Server)'
    });
  } catch (error) {
    console.error('Error saving settings:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// ๐• Detailed Stats endpoint - เธ”เธถเธเธฃเธฒเธขเธเธทเนเธญเธ•เธฒเธกเธเธฃเธฐเน€เธ เธ—
app.get('/api/admin/stats/:type', authenticateAdmin, (req, res) => {
  try {
    const { type } = req.params;
    const validTypes = ['present', 'late', 'absent', 'working'];

    if (!validTypes.includes(type)) {
      return res.status(400).json({
        success: false,
        error: 'Invalid type. Use: present, late, absent, working'
      });
    }

    const data = sqliteService.getDetailedStats(type);
    res.json({
      success: true,
      type,
      count: data.length,
      data
    });
  } catch (error) {
    console.error('Error getting detailed stats:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// ๐• SQLite Stats endpoint
app.get('/api/admin/sqlite-stats', authenticateAdmin, (req, res) => {
  try {
    const stats = sqliteService.getStats();
    const syncStatus = syncService ? syncService.getStatus() : { error: 'Sync service not initialized' };

    res.json({
      success: true,
      data: {
        sqlite: stats,
        sync: syncStatus
      }
    });
  } catch (error) {
    console.error('Error getting SQLite stats:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// ========== ๐• Employee Management APIs ==========

// เธ”เธถเธเธฃเธฒเธขเธเธทเนเธญเธเธเธฑเธเธเธฒเธเธ—เธฑเนเธเธซเธกเธ”
app.get('/api/admin/employees', authenticateAdmin, (req, res) => {
  try {
    const employees = sqliteService.getAllEmployeesWithDetails();
    res.json({
      success: true,
      count: employees.length,
      data: employees
    });
  } catch (error) {
    console.error('Error getting employees:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// เน€เธเธดเนเธกเธเธเธฑเธเธเธฒเธเนเธซเธกเน
app.post('/api/admin/employees', authenticateAdmin, async (req, res) => {
  try {
    const { name } = req.body;
    if (!name || !name.trim()) {
      return res.status(400).json({ success: false, error: 'เธเธฃเธธเธ“เธฒเธฃเธฐเธเธธเธเธทเนเธญเธเธเธฑเธเธเธฒเธ' });
    }

    const result = sqliteService.addEmployee(name.trim());
    if (result) {
      // Sync เนเธ Google Sheets เธ”เนเธงเธข (เธ–เนเธฒเธกเธต syncService)
      if (syncService) {
        try {
          await sheetsService.addEmployeeToSheet(name.trim());
        } catch (e) {
          console.warn('Could not sync new employee to Sheets:', e.message);
        }
      }
      res.json({ success: true, message: `เน€เธเธดเนเธกเธเธเธฑเธเธเธฒเธ "${name}" เธชเธณเน€เธฃเนเธ` });
    } else {
      res.status(400).json({ success: false, error: 'เนเธกเนเธชเธฒเธกเธฒเธฃเธ–เน€เธเธดเนเธกเธเธเธฑเธเธเธฒเธเนเธ”เน' });
    }
  } catch (error) {
    console.error('Error adding employee:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// เธฅเธเธเธเธฑเธเธเธฒเธ
app.delete('/api/admin/employees/:name', authenticateAdmin, async (req, res) => {
  try {
    const { name } = req.params;
    const result = sqliteService.deleteEmployee(decodeURIComponent(name));

    if (result.success) {
      // Sync เนเธ Google Sheets เธ”เนเธงเธข (เธฅเธเนเธ–เธง)
      if (syncService) {
        try {
          await sheetsService.deleteEmployeeFromSheet(decodeURIComponent(name));
        } catch (e) {
          console.warn('Could not sync employee deletion to Sheets:', e.message);
        }
      }
    }

    res.json(result);
  } catch (error) {
    console.error('Error deleting employee:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ========== ๐• Night Shift Employee APIs ==========

// เธ”เธถเธเธฃเธฒเธขเธเธทเนเธญเธเธเธฑเธเธเธฒเธเธเธฐเธเธฅเธฒเธเธเธทเธ
app.get('/api/admin/night-shift', authenticateAdmin, (req, res) => {
  try {
    const employees = sqliteService.getNightShiftEmployees();
    res.json({ success: true, count: employees.length, data: employees });
  } catch (error) {
    console.error('Error getting night shift employees:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// เน€เธเธดเนเธกเธเธเธฑเธเธเธฒเธเธเธฐเธเธฅเธฒเธเธเธทเธ
app.post('/api/admin/night-shift', authenticateAdmin, (req, res) => {
  try {
    const { employeeName } = req.body;
    if (!employeeName || !employeeName.trim()) {
      return res.status(400).json({ success: false, error: 'เธเธฃเธธเธ“เธฒเธฃเธฐเธเธธเธเธทเนเธญเธเธเธฑเธเธเธฒเธ' });
    }
    const result = sqliteService.addNightShiftEmployee(employeeName.trim());
    res.json(result);
  } catch (error) {
    console.error('Error adding night shift employee:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// เธฅเธเธเธเธฑเธเธเธฒเธเธเธฐเธเธฅเธฒเธเธเธทเธ
app.delete('/api/admin/night-shift/:name', authenticateAdmin, (req, res) => {
  try {
    const { name } = req.params;
    const result = sqliteService.removeNightShiftEmployee(decodeURIComponent(name));
    res.json(result);
  } catch (error) {
    console.error('Error removing night shift employee:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ========== ๐• Personal Dashboard APIs ==========

// เธ”เธถเธเธฃเธฒเธขเธเธทเนเธญเธเธเธฑเธเธเธฒเธเธ—เธตเนเธขเธฑเธเนเธกเนเธเธนเธ LINE (เธชเธณเธซเธฃเธฑเธเธซเธเนเธฒเธฅเธเธ—เธฐเน€เธเธตเธขเธ)
app.get('/api/employees/unlinked', (req, res) => {
  try {
    const employees = sqliteService.getUnlinkedEmployees();
    res.json({
      success: true,
      count: employees.length,
      data: employees
    });
  } catch (error) {
    console.error('Error getting unlinked employees:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// เธฅเธเธ—เธฐเน€เธเธตเธขเธเธเธนเธ LINE เธเธฑเธเธเธเธฑเธเธเธฒเธ
app.post('/api/register-line', (req, res) => {
  try {
    const { employee_name, line_user_id, line_name, line_picture } = req.body;

    if (!employee_name || !line_user_id) {
      return res.status(400).json({
        success: false,
        error: 'เธเธฃเธธเธ“เธฒเธฃเธฐเธเธธเธเธทเนเธญเธเธเธฑเธเธเธฒเธเนเธฅเธฐ LINE User ID'
      });
    }

    const result = sqliteService.linkLineUserId(
      employee_name,
      line_user_id,
      line_name || '',
      line_picture || ''
    );

    if (result.success) {
      res.json(result);
    } else {
      res.status(400).json(result);
    }
  } catch (error) {
    console.error('Error registering LINE:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// เธ•เธฃเธงเธเธชเธญเธเธงเนเธฒเธเธนเธ LINE เนเธฅเนเธงเธซเธฃเธทเธญเธขเธฑเธ
app.get('/api/check-registration', (req, res) => {
  try {
    const lineUserId = req.query.line_user_id || req.headers['x-line-userid'];

    if (!lineUserId) {
      return res.json({
        success: true,
        isRegistered: false,
        message: 'เนเธกเนเธเธ LINE User ID'
      });
    }

    const employee = sqliteService.getEmployeeByLineUserId(lineUserId);

    if (employee) {
      res.json({
        success: true,
        isRegistered: true,
        employee: {
          name: employee.name,
          lineName: employee.line_name,
          linePicture: employee.line_picture,
          registeredAt: employee.registered_at
        }
      });
    } else {
      res.json({
        success: true,
        isRegistered: false
      });
    }
  } catch (error) {
    console.error('Error checking registration:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// เธ”เธถเธเธเนเธญเธกเธนเธฅเนเธเธฃเนเธเธฅเนเธชเนเธงเธเธ•เธฑเธง
app.get('/api/my/profile', (req, res) => {
  try {
    const lineUserId = req.query.line_user_id || req.headers['x-line-userid'];

    if (!lineUserId) {
      return res.status(401).json({
        success: false,
        error: 'เธเธฃเธธเธ“เธฒ Login เธ”เนเธงเธข LINE'
      });
    }

    const employee = sqliteService.getEmployeeByLineUserId(lineUserId);

    if (!employee) {
      return res.status(404).json({
        success: false,
        error: 'เนเธกเนเธเธเธเนเธญเธกเธนเธฅ เธเธฃเธธเธ“เธฒเธฅเธเธ—เธฐเน€เธเธตเธขเธเธเนเธญเธ'
      });
    }

    res.json({
      success: true,
      data: {
        name: employee.name,
        lineName: employee.line_name,
        linePicture: employee.line_picture,
        registeredAt: employee.registered_at
      }
    });
  } catch (error) {
    console.error('Error getting profile:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// เธ”เธถเธเธชเธ–เธดเธ•เธดเธชเนเธงเธเธ•เธฑเธง
app.get('/api/my/stats', (req, res) => {
  try {
    const lineUserId = req.query.line_user_id || req.headers['x-line-userid'];
    const month = req.query.month; // format: MM/YYYY

    if (!lineUserId) {
      return res.status(401).json({
        success: false,
        error: 'เธเธฃเธธเธ“เธฒ Login เธ”เนเธงเธข LINE'
      });
    }

    const employee = sqliteService.getEmployeeByLineUserId(lineUserId);

    if (!employee) {
      return res.status(404).json({
        success: false,
        error: 'เนเธกเนเธเธเธเนเธญเธกเธนเธฅ เธเธฃเธธเธ“เธฒเธฅเธเธ—เธฐเน€เธเธตเธขเธเธเนเธญเธ'
      });
    }

    const stats = sqliteService.getPersonalStats(employee.name, month);

    res.json({
      success: true,
      data: {
        ...stats,
        linePicture: employee.line_picture
      }
    });
  } catch (error) {
    console.error('Error getting personal stats:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// เธ”เธถเธเธเธฃเธฐเธงเธฑเธ•เธดเธฅเธเน€เธงเธฅเธฒเธชเนเธงเธเธ•เธฑเธง
app.get('/api/my/history', (req, res) => {
  try {
    const lineUserId = req.query.line_user_id || req.headers['x-line-userid'];
    const limit = parseInt(req.query.limit) || 30;
    const month = req.query.month; // format: MM/YYYY

    if (!lineUserId) {
      return res.status(401).json({
        success: false,
        error: 'เธเธฃเธธเธ“เธฒ Login เธ”เนเธงเธข LINE'
      });
    }

    const employee = sqliteService.getEmployeeByLineUserId(lineUserId);

    if (!employee) {
      return res.status(404).json({
        success: false,
        error: 'เนเธกเนเธเธเธเนเธญเธกเธนเธฅ เธเธฃเธธเธ“เธฒเธฅเธเธ—เธฐเน€เธเธตเธขเธเธเนเธญเธ'
      });
    }

    const history = sqliteService.getPersonalHistory(employee.name, limit, month);

    res.json({
      success: true,
      employeeName: employee.name,
      count: history.length,
      data: history
    });
  } catch (error) {
    console.error('Error getting personal history:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ========== ๐• Admin LINE Management APIs ==========

// เธ”เธถเธเธฃเธฒเธขเธเธทเนเธญเธเธเธฑเธเธเธฒเธเธ—เธตเนเธเธนเธ LINE เนเธฅเนเธง (Admin)
app.get('/api/admin/linked-employees', authenticateAdmin, (req, res) => {
  try {
    const employees = sqliteService.getAllLinkedEmployees();
    res.json({
      success: true,
      count: employees.length,
      linkedCount: employees.filter(e => e.isLinked).length,
      data: employees
    });
  } catch (error) {
    console.error('Error getting linked employees:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// เธขเธเน€เธฅเธดเธเธเธฒเธฃเน€เธเธทเนเธญเธกเธ•เนเธญ LINE (Admin)
app.delete('/api/admin/unlink-line/:name', authenticateAdmin, (req, res) => {
  try {
    const { name } = req.params;
    const result = sqliteService.unlinkLineUserId(decodeURIComponent(name));
    res.json(result);
  } catch (error) {
    console.error('Error unlinking LINE:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ========== ๐• Manual Clock In/Out APIs ==========

function normalizeEmployeesFromRequest(body) {
  if (!body) return [];

  const list = Array.isArray(body.employees)
    ? body.employees
    : (body.employee ? [body.employee] : []);

  return list
    .map(name => (name || '').toString().trim())
    .filter(name => name.length > 0);
}

async function handleManualClock(req, res, type = 'in') {
  try {
    const { date, time, note } = req.body || {};
    const employees = normalizeEmployeesFromRequest(req.body);

    if (!employees.length || !date || !time) {
      return res.status(400).json({
        success: false,
        error: 'เธเธฃเธธเธ“เธฒเธฃเธฐเธเธธ เธเธเธฑเธเธเธฒเธ เธงเธฑเธเธ—เธตเน เนเธฅเธฐเน€เธงเธฅเธฒ',
        errors: ['เธเธฃเธธเธ“เธฒเน€เธฅเธทเธญเธเธเธเธฑเธเธเธฒเธเธญเธขเนเธฒเธเธเนเธญเธข 1 เธเธ']
      });
    }

    const [year, month, day] = date.split('-');
    const clockTime = `${day}/${month}/${year} ${time}:00`;
    const actionLabel = type === 'in' ? 'เธฅเธเน€เธงเธฅเธฒเน€เธเนเธฒ' : 'เธฅเธเน€เธงเธฅเธฒเธญเธญเธ';

    const results = employees.map(employeeName => {
      const payload = type === 'in'
        ? { employee: employeeName, clockInTime: clockTime, adminNote: note || '' }
        : { employee: employeeName, clockOutTime: clockTime, adminNote: note || '' };

      const outcome = type === 'in'
        ? sqliteService.manualClockIn(payload)
        : sqliteService.manualClockOut(payload);

      return { employee: employeeName, ...outcome };
    });

    const successCount = results.filter(r => r.success).length;
    const errors = results
      .filter(r => !r.success)
      .map(r => `${r.employee}: ${r.error || 'เนเธกเนเธ—เธฃเธฒเธเธชเธฒเน€เธซเธ•เธธ'}`);

    const responsePayload = {
      success: successCount > 0,
      partial: errors.length > 0,
      message: `${actionLabel} ${successCount}/${employees.length} เธเธ`,
      results,
      errors
    };

    // Trigger sync if there is at least one successful operation
    if (successCount > 0 && syncService) {
      await syncService.syncToSheets();
      await syncService.syncOnWorkToSheets();
    }

    if (successCount === 0) {
      return res.status(400).json(responsePayload);
    }

    return res.json(responsePayload);
  } catch (error) {
    console.error('Error in manual clock handler:', error);
    return res.status(500).json({ success: false, error: error.message });
  }
}

// Manual Clock In
app.post('/api/admin/manual-clockin', authenticateAdmin, (req, res) => {
  return handleManualClock(req, res, 'in');
});

// Manual Clock Out
app.post('/api/admin/manual-clockout', authenticateAdmin, (req, res) => {
  return handleManualClock(req, res, 'out');
});

// ๐• เธ”เธถเธเธฃเธฒเธขเธเธฒเธฃเน€เธงเธฅเธฒเธชเธณเธซเธฃเธฑเธเนเธเนเนเธ
app.get('/api/admin/time-records/:employee/:date', authenticateAdmin, (req, res) => {
  try {
    const { employee, date } = req.params;
    const records = sqliteService.getTimeRecordsForEdit(decodeURIComponent(employee), date);
    res.json({
      success: true,
      count: records.length,
      data: records
    });
  } catch (error) {
    console.error('Error getting time records:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ๐• เนเธเนเนเธเน€เธงเธฅเธฒเน€เธเนเธฒ/เธญเธญเธ
app.put('/api/admin/time-records/:id', authenticateAdmin, (req, res) => {
  try {
    const { id } = req.params;
    const { employeeName, clockIn, clockOut, note } = req.body;

    // เนเธเธฅเธ date/time เน€เธเนเธ format DD/MM/YYYY HH:mm:ss
    let newClockIn = null;
    let newClockOut = null;

    if (clockIn && clockIn.date && clockIn.time) {
      const [year, month, day] = clockIn.date.split('-');
      newClockIn = `${day}/${month}/${year} ${clockIn.time}:00`;
    }

    if (clockOut && clockOut.date && clockOut.time) {
      const [year, month, day] = clockOut.date.split('-');
      newClockOut = `${day}/${month}/${year} ${clockOut.time}:00`;
    }

    const result = sqliteService.updateTimeRecord({
      recordId: parseInt(id),
      employeeName,
      newClockIn,
      newClockOut,
      adminNote: note || ''
    });

    res.json(result);
  } catch (error) {
    console.error('Error updating time record:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ๐• เธฅเธเธฃเธฒเธขเธเธฒเธฃเน€เธงเธฅเธฒ (เธเธฃเนเธญเธก sync เธฅเธเธเธฒเธ Sheets)
app.delete('/api/admin/time-records/:id', authenticateAdmin, async (req, res) => {
  try {
    const { id } = req.params;

    // เธ”เธถเธเธเนเธญเธกเธนเธฅเธเนเธญเธเธฅเธ เน€เธเธทเนเธญเนเธเน sync เธเธฑเธ Sheets
    const record = sqliteService.db.prepare('SELECT * FROM time_records WHERE id = ?').get(parseInt(id));

    // เธฅเธเธเธฒเธ SQLite
    const result = sqliteService.deleteTimeRecord(parseInt(id));

    if (result.success && record) {
      // ๐• Sync เธฅเธเธเธฒเธ Google Sheets
      try {
        await sheetsService.deleteTimeRecordFromSheet(record.employee_name, record.clock_in);

        // เธฅเธเธเธฒเธ On Work เธ”เนเธงเธข (เธ–เนเธฒเน€เธเนเธเธเธเธ—เธตเนเธขเธฑเธเนเธกเน clock out)
        if (!record.clock_out) {
          await sheetsService.deleteFromOnWorkSheet(record.employee_name);
        }

        console.log(`โ… Deleted from both SQLite and Sheets: ${record.employee_name}`);
      } catch (sheetsError) {
        console.error('โ ๏ธ Sheets sync delete error (SQLite still deleted):', sheetsError);
        // เนเธกเน return error เน€เธเธฃเธฒเธฐ SQLite เธฅเธเธชเธณเน€เธฃเนเธเนเธฅเนเธง
      }
    }

    res.json(result);
  } catch (error) {
    console.error('Error deleting time record:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ๐• Force Sync endpoint
app.post('/api/admin/force-sync', authenticateAdmin, async (req, res) => {
  try {
    if (!syncService) {
      return res.status(500).json({
        success: false,
        error: 'Sync service not initialized'
      });
    }

    const result = await syncService.forceSync();

    res.json({
      success: true,
      message: 'Force sync completed',
      data: result
    });
  } catch (error) {
    console.error('Error in force sync:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// ๐• Force Reload from Sheets (Sheets โ’ SQLite) - เธ”เธถเธเธเนเธญเธกเธนเธฅเนเธซเธกเนเธเธฒเธ Sheets เธกเธฒเธ—เธฑเธ SQLite
app.post('/api/admin/force-reload-from-sheets', authenticateAdmin, async (req, res) => {
  try {
    if (!syncService) {
      return res.status(500).json({
        success: false,
        error: 'Sync service not initialized'
      });
    }

    console.log('๐” Force reloading data from Sheets to SQLite...');

    // เน€เธเธฅเธตเธขเธฃเนเธเนเธญเธกเธนเธฅเน€เธเนเธฒเนเธ SQLite เธ—เธฑเนเธเธซเธกเธ”เน€เธเธทเนเธญเนเธซเธฅเธ”เนเธซเธกเนเธเธฒเธ Sheets (Reset Data)
    sqliteService.db.exec('DELETE FROM time_records');
    sqliteService.db.exec('DELETE FROM on_work');
    sqliteService.db.exec('DELETE FROM employees'); // ๐• เธฅเธเธฃเธฒเธขเธเธทเนเธญเธเธเธฑเธเธเธฒเธเธ”เนเธงเธข เน€เธเธทเนเธญเนเธซเนเธ•เธฃเธเธเธฑเธ Sheets 100%

    // เนเธซเธฅเธ”เนเธซเธกเนเธเธฒเธ Sheets
    await syncService.loadFromSheets();

    // เน€เธเธฅเธตเธขเธฃเน cache
    sheetsService.clearCache();

    console.log('โ… Force reload from Sheets completed');

    res.json({
      success: true,
      message: 'เนเธซเธฅเธ”เธเนเธญเธกเธนเธฅเนเธซเธกเนเธเธฒเธ Sheets เธชเธณเน€เธฃเนเธ',
      stats: sqliteService.getStats()
    });
  } catch (error) {
    console.error('Error in force reload from sheets:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// ๐• Repair On-Work Status Endpoint (Admin)
app.post('/api/admin/repair-on-work', authenticateAdmin, async (req, res) => {
  try {
    const result = sqliteService.repairOnWorkFromTimeRecords();

    if (result.success) {
      if (result.repairedCount > 0 && syncService) {
        // Force sync status to sheets
        await syncService.syncOnWorkToSheets();
      }
      res.json(result);
    } else {
      res.status(500).json(result);
    }
  } catch (error) {
    console.error('Error repairing on-work status:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ๐• System Settings APIs
// GET Settings
app.get('/api/admin/settings', authenticateAdmin, (req, res) => {
  try {
    const envPath = path.join(__dirname, '.env');
    if (!fs.existsSync(envPath)) {
      return res.json({ success: true, settings: {} });
    }
    const envContent = fs.readFileSync(envPath, 'utf8');
    const settings = {};
    envContent.split('\n').forEach(line => {
      const trimmed = line.trim();
      if (trimmed && !trimmed.startsWith('#')) {
        const [key, ...values] = trimmed.split('=');
        if (key) settings[key.trim()] = values.join('=').trim();
      }
    });
    res.json({ success: true, settings });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// POST Settings
app.post('/api/admin/settings', authenticateAdmin, (req, res) => {
  try {
    const newSettings = req.body; // { key: value }
    const envPath = path.join(__dirname, '.env');
    let envContent = '';

    if (fs.existsSync(envPath)) {
      envContent = fs.readFileSync(envPath, 'utf8');
    }

    const lines = envContent.split('\n');
    const newLines = [];
    const processedKeys = new Set();

    // Update existing keys
    lines.forEach(line => {
      const trimmed = line.trim();
      if (trimmed && !trimmed.startsWith('#')) {
        const key = trimmed.split('=')[0].trim();
        if (newSettings.hasOwnProperty(key)) {
          newLines.push(`${key}=${newSettings[key]}`);
          processedKeys.add(key);
        } else {
          newLines.push(line);
        }
      } else {
        newLines.push(line);
      }
    });

    // Add new keys
    Object.keys(newSettings).forEach(key => {
      if (!processedKeys.has(key)) {
        newLines.push(`${key}=${newSettings[key]}`);
      }
    });

    fs.writeFileSync(envPath, newLines.join('\n'));
    res.json({ success: true, message: 'เธเธฑเธเธ—เธถเธเธเธฒเธฃเธ•เธฑเนเธเธเนเธฒเน€เธฃเธตเธขเธเธฃเนเธญเธข (เธเธฃเธธเธ“เธฒ Restart Server เน€เธเธทเนเธญเนเธซเนเธเนเธฒเธเธฒเธเธญเธขเนเธฒเธเธกเธตเธเธฅ)' });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// API เธชเธณเธซเธฃเธฑเธ manual cache refresh (เธชเธณเธซเธฃเธฑเธ admin เน€เธ—เนเธฒเธเธฑเนเธ)
app.post('/api/admin/refresh-cache', authenticateAdmin, async (req, res) => {
  try {
    console.log('๐” Manual cache refresh initiated by admin');

    // Clear all cache
    sheetsService.clearCache();

    // Warm critical caches
    await sheetsService.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
    await sheetsService.getCachedSheetData(CONFIG.SHEETS.EMPLOYEES);

    res.json({
      success: true,
      message: 'Cache refreshed successfully'
    });
  } catch (error) {
    console.error('Cache refresh error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to refresh cache'
    });
  }
});

// API เธชเธณเธซเธฃเธฑเธเธ•เธฃเธงเธเธชเธญเธเธชเธ–เธฒเธเธฐ API quota
app.get('/api/admin/quota-status', authenticateAdmin, async (req, res) => {
  try {
    const apiStats = apiMonitor.getStats();
    const isEmergencyMode = sheetsService.emergencyMode || false;

    // เธ—เธ”เธชเธญเธเธเธฒเธฃเน€เธเธทเนเธญเธกเธ•เนเธญ API
    let apiHealthy = true;
    let lastError = null;

    try {
      await sheetsService.getCachedSheetData(CONFIG.SHEETS.EMPLOYEES);
    } catch (error) {
      apiHealthy = false;
      lastError = error.message;
    }

    res.json({
      success: true,
      data: {
        apiHealthy,
        emergencyMode: isEmergencyMode,
        lastError,
        apiStats,
        recommendations: apiHealthy ?
          ['เธฃเธฐเธเธเธ—เธณเธเธฒเธเธเธเธ•เธด'] :
          [
            'เธฃเธญเนเธซเน quota reset (เธ เธฒเธขเนเธ 24 เธเธฑเนเธงเนเธกเธ)',
            'เนเธเน cached data เนเธเธฃเธฐเธขเธฐเธเธตเน',
            'เธฅเธ”เธเธฒเธฃเนเธเนเธเธฒเธเธเธตเน€เธเธญเธฃเนเธ—เธตเนเธ•เนเธญเธเนเธเน API'
          ]
      }
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// API เธชเธณเธซเธฃเธฑเธเน€เธเธดเธ”/เธเธดเธ” emergency mode
app.post('/api/admin/emergency-mode', authenticateAdmin, async (req, res) => {
  try {
    const { enabled } = req.body;
    sheetsService.setEmergencyMode(enabled);

    res.json({
      success: true,
      message: `Emergency mode ${enabled ? 'enabled' : 'disabled'}`,
      emergencyMode: enabled
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// ========== ๐• Live Status Dashboard APIs (Public) ==========

// GET /api/status/live - เธ”เธถเธเธฃเธฒเธขเธเธฒเธฃเธฅเธเน€เธงเธฅเธฒเธฅเนเธฒเธชเธธเธ” (Real-time Feed)
app.get('/api/status/live', (req, res) => {
  try {
    const { limit = 30, date } = req.query;

    const summary = sqliteService.getTodaySummary(date || null);
    const records = sqliteService.getRecentActivity(parseInt(limit), date || null);

    res.json({
      success: true,
      timestamp: new Date().toISOString(),
      data: {
        summary,
        records
      }
    });
  } catch (error) {
    console.error('Error getting live status:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// GET /api/status/history - เธเนเธเธซเธฒเธเธฃเธฐเธงเธฑเธ•เธดเธเธฒเธฃเธฅเธเน€เธงเธฅเธฒ
app.get('/api/status/history', (req, res) => {
  try {
    const { name, date, limit = 50 } = req.query;

    let records;
    if (name && name.trim()) {
      records = sqliteService.getActivityByName(name.trim(), date || null, parseInt(limit));
    } else {
      records = sqliteService.getRecentActivity(parseInt(limit), date || null);
    }

    const summary = sqliteService.getTodaySummary(date || null);

    res.json({
      success: true,
      data: {
        summary,
        records,
        filters: {
          name: name || null,
          date: date || summary.date,
          limit: parseInt(limit)
        }
      }
    });
  } catch (error) {
    console.error('Error getting status history:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// ========== ๐• Auto-Update System APIs ==========

const { exec } = require('child_process');
const util = require('util');
const execPromise = util.promisify(exec);

// GET /api/admin/check-update - เธ•เธฃเธงเธเธชเธญเธเน€เธงเธญเธฃเนเธเธฑเธเนเธซเธกเน
app.get('/api/admin/check-update', authenticateAdmin, async (req, res) => {
  try {
    // เธญเนเธฒเธ version เธเธฑเธเธเธธเธเธฑเธ
    const versionPath = path.join(__dirname, 'version.json');
    let currentVersion = { version: '0.0.0', changelog: '' };

    if (fs.existsSync(versionPath)) {
      currentVersion = JSON.parse(fs.readFileSync(versionPath, 'utf8'));
    }

    // เธ”เธถเธ version เธเธฒเธ GitHub
    const githubUrl = 'https://raw.githubusercontent.com/TimetrackerUD01/time_tracker/main/version.json';
    const response = await fetch(githubUrl);

    if (!response.ok) {
      return res.json({
        success: true,
        hasUpdate: false,
        currentVersion: currentVersion.version,
        latestVersion: currentVersion.version,
        message: 'เนเธกเนเธชเธฒเธกเธฒเธฃเธ–เน€เธเธทเนเธญเธกเธ•เนเธญ GitHub เนเธ”เน'
      });
    }

    const latestVersion = await response.json();
    const hasUpdate = latestVersion.version !== currentVersion.version;

    res.json({
      success: true,
      hasUpdate,
      currentVersion: currentVersion.version,
      latestVersion: latestVersion.version,
      changelog: hasUpdate ? latestVersion.changelog : '',
      buildDate: latestVersion.buildDate
    });

  } catch (error) {
    console.error('Error checking update:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// POST /api/admin/update - เธ”เธณเน€เธเธดเธเธเธฒเธฃเธญเธฑเธเน€เธ”เธ•
app.post('/api/admin/update', authenticateAdmin, async (req, res) => {
  try {
    console.log('๐” Starting system update...');

    // เธ•เธฃเธงเธเธชเธญเธเธงเนเธฒเธญเธขเธนเนเนเธ git repository เธซเธฃเธทเธญเนเธกเน
    try {
      await execPromise('git status', { cwd: __dirname });
    } catch (gitError) {
      return res.status(400).json({
        success: false,
        error: 'เนเธกเนเธเธ Git repository - เธ•เนเธญเธ clone เธเธฒเธ GitHub เธเนเธญเธ'
      });
    }

    // git fetch
    console.log('๐“ฅ Fetching from GitHub...');
    await execPromise('git fetch origin main', { cwd: __dirname });

    // git pull
    console.log('๐“ฅ Pulling latest code...');
    const { stdout: pullOutput } = await execPromise('git pull origin main', { cwd: __dirname });
    console.log('Pull output:', pullOutput);

    // npm install (เธ–เนเธฒเธกเธต package เนเธซเธกเน)
    console.log('๐“ฆ Installing dependencies...');
    await execPromise('npm install --production', { cwd: __dirname });

    // เธญเนเธฒเธ version เนเธซเธกเน
    const versionPath = path.join(__dirname, 'version.json');
    let newVersion = { version: 'unknown' };
    if (fs.existsSync(versionPath)) {
      newVersion = JSON.parse(fs.readFileSync(versionPath, 'utf8'));
    }

    // เธชเนเธ response เธเนเธญเธ restart
    res.json({
      success: true,
      message: `เธญเธฑเธเน€เธ”เธ•เน€เธเนเธเน€เธงเธญเธฃเนเธเธฑเธ ${newVersion.version} เธชเธณเน€เธฃเนเธ!`,
      version: newVersion.version,
      restartIn: 3
    });

    // Restart PM2 เธซเธฅเธฑเธเธเธฒเธเธชเนเธ response
    setTimeout(async () => {
      try {
        console.log('๐” Restarting server with PM2...');
        await execPromise('pm2 restart all');
      } catch (pmError) {
        console.log('โน๏ธ PM2 restart failed, trying nodemon reload...');
        // เธ–เนเธฒเนเธกเนเธกเธต PM2 เธเนเธเนเธฒเธกเนเธ (nodemon เธเธฐ restart เน€เธญเธ)
      }
    }, 1000);

  } catch (error) {
    console.error('โ Update failed:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// GET /api/admin/version - เธ”เธนเน€เธงเธญเธฃเนเธเธฑเธเธเธฑเธเธเธธเธเธฑเธ
app.get('/api/admin/version', (req, res) => {
  try {
    const versionPath = path.join(__dirname, 'version.json');
    if (fs.existsSync(versionPath)) {
      const version = JSON.parse(fs.readFileSync(versionPath, 'utf8'));
      res.json({ success: true, ...version });
    } else {
      res.json({ success: true, version: '1.0.0', changelog: '' });
    }
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// ========== Error Handling ==========
app.use((error, req, res, next) => {
  console.error('Global error handler:', error);
  res.status(500).json({
    success: false,
    error: 'Internal server error'
  });
});

app.use((req, res) => {
  res.status(404).json({
    success: false,
    error: 'Not found'
  });
});

// ========== Start Server ==========
async function startServer() {
  try {
    console.log('๐€ Starting Time Tracker Server with Admin Panel...');
    console.log(`๐ Environment: ${process.env.NODE_ENV || 'development'}`);

    // เธ•เธฃเธงเธเธชเธญเธ environment variables
    if (!validateConfig()) {
      console.error('โ Server startup aborted due to missing configuration');
      process.exit(1);
    }

    // เน€เธฃเธดเนเธกเธ•เนเธ Google Sheets Service
    console.log('๐“ Initializing Google Sheets Service...');
    await sheetsService.initialize();
    console.log('โ… Google Sheets Service initialized successfully');

    // ๐• เน€เธฃเธดเนเธกเธ•เนเธ SQLite Service
    console.log('๐’พ Initializing SQLite Service...');
    sqliteService.initialize();
    console.log('โ… SQLite Service initialized successfully');

    // ๐• เน€เธฃเธดเนเธกเธ•เนเธ Sync Service
    syncService = new SyncService(sqliteService, sheetsService);

    // เนเธซเธฅเธ”เธเนเธญเธกเธนเธฅเธเธฒเธ Sheets โ’ SQLite เธ•เธญเธ startup
    if (CONFIG.SYNC.SYNC_ON_STARTUP) {
      console.log('๐” Loading data from Google Sheets to SQLite...');
      const loadResult = await syncService.loadFromSheets();
      if (loadResult.success) {
        console.log('โ… Data loaded from Sheets successfully');
      } else {
        console.warn('โ ๏ธ Failed to load data from Sheets, starting with empty database');
      }
    }

    // เน€เธฃเธดเนเธก periodic sync
    if (CONFIG.SYNC.ENABLED) {
      syncService.startPeriodicSync();
      console.log(`๐” Periodic sync enabled (every ${CONFIG.SYNC.INTERVAL_MS / 1000}s)`);
    }

    console.log('๐“ SQLite Stats:', sqliteService.getStats());

    // เน€เธฃเธดเนเธกเธ•เนเธ Keep-Alive Service
    if (CONFIG.RENDER.KEEP_ALIVE_ENABLED) {
      console.log('๐” Starting Keep-Alive Service...');
      keepAliveService.init();
    } else {
      console.log('โ ๏ธ Keep-Alive Service is disabled');
    }

    // เธ•เธฑเนเธเธเนเธฒ cron job เธชเธณเธซเธฃเธฑเธเธ•เธฃเธงเธเธชเธญเธเธฅเธทเธกเธฅเธเน€เธงเธฅเธฒเธญเธญเธ (เธ—เธธเธเธงเธฑเธเน€เธงเธฅเธฒ 23:59:59)
    cron.schedule('59 59 23 * * *', async () => {
      console.log('๐• Running daily missed checkout check at 23:59:59...');
      try {
        const result = await sheetsService.checkAndHandleMissedCheckouts();
        console.log(`โ… Missed checkout check completed: ${result.processedCount} employees processed`);

        // เธชเนเธเธเธฒเธฃเนเธเนเธเน€เธ•เธทเธญเธเธ–เนเธฒเธกเธตเธเธฒเธฃเธเธฃเธฐเธกเธงเธฅเธเธฅ
        if (result.processedCount > 0) {
          console.log(`๐“ฑ Auto-processed ${result.processedCount} missed checkouts`);
        }
      } catch (error) {
        console.error('โ Error in missed checkout check:', error);
      }
    }, {
      scheduled: true,
      timezone: CONFIG.TIMEZONE
    });

    // เน€เธฃเธดเนเธกเธ•เนเธเน€เธเธดเธฃเนเธเน€เธงเธญเธฃเน
    const server = app.listen(PORT, () => {
      console.log('๐ Server Started Successfully!');
      console.log(`๐ Server running on port ${PORT}`);
      console.log(`๐“ฑ Public URL: ${CONFIG.RENDER.SERVICE_URL}`);
      console.log(`โ๏ธ Admin Panel: ${CONFIG.RENDER.SERVICE_URL}/admin`);
      console.log(`๐• Timezone: ${CONFIG.TIMEZONE}`);
      console.log(`๐”ง Keep-Alive: ${CONFIG.RENDER.KEEP_ALIVE_ENABLED ? 'Enabled' : 'Disabled'}`);
      console.log('โ”€'.repeat(50));
    });

    // Graceful shutdown
    process.on('SIGTERM', () => {
      console.log('๐‘ Received SIGTERM, shutting down gracefully...');
      server.close(() => {
        console.log('โ… Server closed');
        process.exit(0);
      });
    });

    process.on('SIGINT', () => {
      console.log('๐‘ Received SIGINT, shutting down gracefully...');
      server.close(() => {
        console.log('โ… Server closed');
        process.exit(0);
      });
    });

  } catch (error) {
    console.error('โ Failed to start server:', error);
    process.exit(1);
  }
}

// เน€เธฃเธตเธขเธเนเธเนเธเธฑเธเธเนเธเธฑเธ startServer
startServer();

