/**
 * TAM Research Tools - Calendar automation
 *
 * Menu structure:
 * 1) [All sheets] Create season games calendar events
 * 2) [All sheets] Delete all season games calendar events
 * 3) [All sheets] Fill-in ERROR events
 * 4) [Single sheet] Create season games calendar events
 * 5) [Single sheet] Create multiple games calendar events
 * 6) [Single sheet] Create single game calendar event
 * 7) [Single sheet] Delete season games calendar events
 * 8) [Single sheet] Delete single game calendar event
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TAM Research Tools')
    .addItem('[All sheets] Create season games calendar events', 'createAllSheetsCalendarEvents')
    .addItem('[All sheets] Delete all season games calendar events', 'deleteAllSheetsSeasonEvents')
    .addItem('[All sheets] Fill-in ERROR events', 'fillInErrorEventsAllSheets')
    .addItem('[Single sheet] Create season games calendar events', 'createAllCalendarEvents')
    .addItem('[Single sheet] Create multiple games calendar events', 'createMultipleGamesCalendarEvents')
    .addItem('[Single sheet] Create single game calendar event', 'createSingleCalendarEvent')
    .addItem('[Single sheet] Delete season games calendar events', 'deleteSingleSheetSeasonEvents')
    .addItem('[Single sheet] Delete single game calendar event', 'deleteSingleGameCalendarEvent')
    .addToUi();
}

/* =========================
   Feature flags / detection
   ========================= */
function HAS_ADVANCED_CALENDAR() {
  return (typeof Calendar !== 'undefined') && Calendar.Events && Calendar.Events.insert;
}

/* =========================
   Global team data & lookups (with colors)
   ========================= */
const TEAMS = [
  { city: 'Atlanta',       full_name: 'Atlanta Hawks',              abbreviation: 'ATL', nickname: 'Hawks',         color: '#0e632c' },
  { city: 'Boston',        full_name: 'Boston Celtics',             abbreviation: 'BOS', nickname: 'Celtics',       color: '#0e632c' },
  { city: 'Brooklyn',      full_name: 'Brooklyn Nets',              abbreviation: 'BKN', nickname: 'Nets',          color: '#0e632c' },
  { city: 'Charlotte',     full_name: 'Charlotte Hornets',          abbreviation: 'CHA', nickname: 'Hornets',       color: '#0e632c' },
  { city: 'Chicago',       full_name: 'Chicago Bulls',              abbreviation: 'CHI', nickname: 'Bulls',         color: '#0e632c' },
  { city: 'Cleveland',     full_name: 'Cleveland Cavaliers',        abbreviation: 'CLE', nickname: 'Cavaliers',     color: '#0e632c' },
  { city: 'Dallas',        full_name: 'Dallas Mavericks',           abbreviation: 'DAL', nickname: 'Mavericks',     color: '#002b5e' },
  { city: 'Denver',        full_name: 'Denver Nuggets',             abbreviation: 'DEN', nickname: 'Nuggets',       color: '#0e632c' },
  { city: 'Detroit',       full_name: 'Detroit Pistons',            abbreviation: 'DET', nickname: 'Pistons',       color: '#0e632c' },
  { city: 'Golden State',  full_name: 'Golden State Warriors',      abbreviation: 'GSW', nickname: 'Warriors',      color: '#1d428a' },
  { city: 'Houston',       full_name: 'Houston Rockets',            abbreviation: 'HOU', nickname: 'Rockets',       color: '#CE1141' },
  { city: 'Indiana',       full_name: 'Indiana Pacers',             abbreviation: 'IND', nickname: 'Pacers',        color: '#0e632c' },
  { city: 'Los Angeles',   full_name: 'LA Clippers',                abbreviation: 'LAC', nickname: 'Clippers',      color: '#C8102E' },
  { city: 'Los Angeles',   full_name: 'Los Angeles Lakers',         abbreviation: 'LAL', nickname: 'Lakers',        color: '#552583' },
  { city: 'Memphis',       full_name: 'Memphis Grizzlies',          abbreviation: 'MEM', nickname: 'Grizzlies',     color: '#0e632c' },
  { city: 'Miami',         full_name: 'Miami Heat',                 abbreviation: 'MIA', nickname: 'Heat',          color: '#0e632c' },
  { city: 'Milwaukee',     full_name: 'Milwaukee Bucks',            abbreviation: 'MIL', nickname: 'Bucks',         color: '#0e632c' },
  { city: 'Minnesota',     full_name: 'Minnesota Timberwolves',     abbreviation: 'MIN', nickname: 'Timberwolves',  color: '#0e632c' },
  { city: 'New Orleans',   full_name: 'New Orleans Pelicans',       abbreviation: 'NOP', nickname: 'Pelicans',      color: '#0e632c' },
  { city: 'New York',      full_name: 'New York Knicks',            abbreviation: 'NYK', nickname: 'Knicks',        color: '#0e632c' },
  { city: 'Oklahoma City', full_name: 'Oklahoma City Thunder',      abbreviation: 'OKC', nickname: 'Thunder',       color: '#0e632c' },
  { city: 'Orlando',       full_name: 'Orlando Magic',              abbreviation: 'ORL', nickname: 'Magic',         color: '#0e632c' },
  { city: 'Philadelphia',  full_name: 'Philadelphia 76ers',         abbreviation: 'PHI', nickname: '76ers',         color: '#0e632c' },
  { city: 'Phoenix',       full_name: 'Phoenix Suns',               abbreviation: 'PHX', nickname: 'Suns',          color: '#0e632c' },
  { city: 'Portland',      full_name: 'Portland Trail Blazers',     abbreviation: 'POR', nickname: 'Trail Blazers', color: '#0e632c' },
  { city: 'Sacramento',    full_name: 'Sacramento Kings',           abbreviation: 'SAC', nickname: 'Kings',         color: '#0e632c' },
  { city: 'San Antonio',   full_name: 'San Antonio Spurs',          abbreviation: 'SAS', nickname: 'Spurs',         color: '#000000' },
  { city: 'Toronto',       full_name: 'Toronto Raptors',            abbreviation: 'TOR', nickname: 'Raptors',       color: '#0e632c' },
  { city: 'Utah',          full_name: 'Utah Jazz',                  abbreviation: 'UTA', nickname: 'Jazz',          color: '#0e632c' },
  { city: 'Washington',    full_name: 'Washington Wizards',         abbreviation: 'WAS', nickname: 'Wizards',       color: '#0e632c' },
];

const TEAMS_BY_ABBR = TEAMS.reduce((acc, t) => { acc[t.abbreviation.toUpperCase()] = t; return acc; }, {});
const TEAMS_BY_CITY = TEAMS.reduce((acc, t) => {
  const key = t.city.toLowerCase();
  if (!acc[key]) acc[key] = [];
  acc[key].push(t);
  return acc;
}, {});
const CITY_DISAMBIG_PREF = { 'los angeles': 'LAL' };

function cityToAbbr(cityOrAbbrRaw) {
  if (!cityOrAbbrRaw) return '';
  let s = String(cityOrAbbrRaw).trim();

  const maybeAbbr = s.toUpperCase();
  if (/^[A-Z]{2,4}$/.test(maybeAbbr) && TEAMS_BY_ABBR[maybeAbbr]) return maybeAbbr;

  const key = s.toLowerCase();
  const teams = TEAMS_BY_CITY[key];
  if (!teams || teams.length === 0) return s;
  if (teams.length === 1) return teams[0].abbreviation;
  if (/clippers|lac/i.test(s)) return 'LAC';
  if (/lakers|lal/i.test(s)) return 'LAL';
  const pref = CITY_DISAMBIG_PREF[key];
  return pref ? pref : teams[0].abbreviation;
}

/* =========================
   Calendar color mapping
   ========================= */
const GCAL_EVENT_COLORS = [
  { id: 1,  hex: '#a4bdfc', name: 'LAVENDER',  appConst: 'LAVENDER'  },
  { id: 2,  hex: '#7ae7bf', name: 'SAGE',      appConst: 'SAGE'      },
  { id: 3,  hex: '#dbadff', name: 'GRAPE',     appConst: 'GRAPE'     },
  { id: 4,  hex: '#ff887c', name: 'FLAMINGO',  appConst: 'FLAMINGO'  },
  { id: 5,  hex: '#fbd75b', name: 'BANANA',    appConst: 'BANANA'    },
  { id: 6,  hex: '#ffb878', name: 'TANGERINE', appConst: 'TANGERINE' },
  { id: 7,  hex: '#46d6db', name: 'PEACOCK',   appConst: 'PEACOCK'   },
  { id: 8,  hex: '#e1e1e1', name: 'GRAPHITE',  appConst: 'GRAPHITE'  },
  { id: 9,  hex: '#5484ed', name: 'BLUEBERRY', appConst: 'BLUEBERRY' },
  { id: 10, hex: '#51b749', name: 'BASIL',     appConst: 'BASIL'     },
  { id: 11, hex: '#dc2127', name: 'TOMATO',    appConst: 'TOMATO'    },
];

function _hexToRgb(hex) {
  const m = String(hex || '').trim().replace('#','').match(/^([0-9a-f]{6})$/i);
  if (!m) return { r: 0, g: 0, b: 0 };
  const n = parseInt(m[1], 16);
  return { r: (n>>16)&255, g: (n>>8)&255, b: n&255 };
}
function _dist2(a, b) {
  const dr = a.r - b.r, dg = a.g - b.g, db = a.b - b.b;
  return dr*dr + dg*dg + db*db;
}
function teamHexToColorId(hex) {
  const rgb = _hexToRgb(hex);
  let best = GCAL_EVENT_COLORS[0], bestD = Infinity;
  for (const c of GCAL_EVENT_COLORS) {
    const d = _dist2(rgb, _hexToRgb(c.hex));
    if (d < bestD) { bestD = d; best = c; }
  }
  return best.id;
}
function colorIdToCalendarAppConst(id) {
  const c = GCAL_EVENT_COLORS.find(c => c.id === id);
  return c ? c.appConst : 'BASIL';
}

/* =========================
   URL building & formatting
   ========================= */

function _eventLink(calendar, event) {
  const calendarId = (typeof event.getOriginalCalendarId === 'function')
    ? event.getOriginalCalendarId()
    : calendar.getId();
  const eventIdPart = String(event.getId()).split('@')[0];
  const eid = Utilities.base64Encode(`${eventIdPart} ${calendarId}`).replace(/=+$/,'');
  return `https://calendar.google.com/calendar/event?eid=${eid}`;
}

/**
 * Style a computed column (like "event_url"):
 * - Header background #d9d9d9
 * - Data cells background #c9daf8 (UPDATED)
 * - Box borders with color #d9d9d9
 * - URL cells in #4285f4 and no underline
 */
function _styleComputedColumn(sheet, colIndex1Based, lastRow) {
  if (colIndex1Based <= 0) return;

  sheet.getRange(1, colIndex1Based, 1, 1).setBackground('#d9d9d9');

  if (lastRow >= 2) {
    const dataRange = sheet.getRange(2, colIndex1Based, lastRow - 1, 1);
    dataRange.setBackground('#c9daf8');
    dataRange.setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);
    const vals = dataRange.getValues();
    for (let r = 0; r < vals.length; r++) {
      const v = vals[r][0];
      if (v && typeof v === 'string' && /^https?:\/\//i.test(v)) {
        dataRange.getCell(r + 1, 1).setFontColor('#4285f4').setFontLine('none');
      }
    }
  }
}

/* =========================
   Parsing helpers
   ========================= */

function _parseMonthDay(dateCell) {
  if (dateCell instanceof Date && !isNaN(dateCell.getTime())) {
    return { monthIndex: dateCell.getMonth(), day: dateCell.getDate() };
  }
  const s = String(dateCell).trim();
  const monthMap = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,sept:8,oct:9,nov:10,dec:11 };
  const cleaned = s.replace(/^[A-Za-z]{3,9},?\s*/,'').replace(/,/g,'').trim();
  const m = cleaned.match(/^([A-Za-z]{3,9})\s+(\d{1,2})$/);
  if (!m) throw new Error(`Unrecognized DATE format: "${s}"`);
  const monthToken = m[1].toLowerCase();
  const day = parseInt(m[2], 10);
  const monthIndex = monthMap[monthToken];
  if (monthIndex == null || !(day >= 1 && day <= 31)) throw new Error(`Unrecognized DATE values: "${s}"`);
  return { monthIndex, day };
}

function _parseTime(timeCell) {
  if (timeCell instanceof Date && !isNaN(timeCell.getTime())) {
    return { hour24: timeCell.getHours(), minute: timeCell.getMinutes() };
  }
  if (typeof timeCell === 'number' && !isNaN(timeCell)) {
    const totalMinutes = Math.round(timeCell * 24 * 60);
    const hour24 = Math.floor(totalMinutes / 60) % 24;
    const minute = totalMinutes % 60;
    return { hour24, minute };
  }
  const raw = String(timeCell || '').trim();
  if (!raw) throw new Error('TIME is blank');
  const s = raw.replace(/\b(ET|EST|EDT|PT|PST|PDT|UTC|GMT)\b/gi, '').trim();
  const m12 = s.match(/^(\d{1,2})(?::(\d{2}))?\s*([AaPp][Mm])$/);
  if (m12) {
    let hour = parseInt(m12[1], 10);
    const minute = m12[2] != null ? parseInt(m12[2], 10) : 0;
    const ampm = m12[3].toUpperCase();
    if (hour === 12) hour = (ampm === 'AM') ? 0 : 12;
    else if (ampm === 'PM') hour += 12;
    return { hour24: hour, minute };
  }
  const m24 = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m24) {
    const hour = parseInt(m24[1], 10);
    const minute = parseInt(m24[2], 10);
    if (hour < 0 || hour > 23 || minute < 0 || minute > 59) throw new Error(`Invalid TIME values: "${raw}"`);
    return { hour24: hour, minute };
  }
  throw new Error(`Unrecognized TIME format: "${raw}"`);
}

/* =========================
   Title & duplicate logic
   ========================= */

function _buildEventTitle(sheetInfo, opponentRaw) {
  let s = String(opponentRaw || '').trim();
  let venue = 'Home';
  if (/^@/.test(s)) { venue = 'Away'; s = s.replace(/^@+/, '').trim(); }
  else if (/^(vs\.?|v\.?)\s*/i.test(s)) { venue = 'Home'; s = s.replace(/^(vs\.?|v\.?)\s*/i, '').trim(); }
  const oppAbbr = cityToAbbr(s);
  const teamUpper = String(sheetInfo.team || '').toUpperCase();
  const prefix = `${sheetInfo.emoji}:${sheetInfo.country}:${sheetInfo.league}`;
  return `${prefix} || ${teamUpper} vs ${oppAbbr} (${venue})`;
}

/* =========================
   Row parsing & event creation
   ========================= */

function _getEventDetailsFromRow(rowData, sheetInfo, headers) {
  const dateCol = headers.indexOf('DATE');
  const opponentCol = headers.indexOf('OPPONENT');
  const timeCol = headers.indexOf('TIME');
  if (dateCol === -1 || timeCol === -1 || opponentCol === -1) {
    throw new Error('Missing one or more required headers: DATE, TIME, OPPONENT');
  }

  const dateCell = rowData[dateCol];
  const timeCell = rowData[timeCol];
  const opponentStr = rowData[opponentCol];

  const { monthIndex, day } = _parseMonthDay(dateCell);
  const { hour24, minute } = _parseTime(timeCell);

  const year = (monthIndex >= 9) ? 2025 : 2026; // season split

  const startDateTime = new Date(year, monthIndex, day, hour24, minute, 0, 0);
  const title = _buildEventTitle(sheetInfo, opponentStr);

  return { title, startTime: startDateTime };
}

/**
 * Insert event via Advanced Calendar (preferred) and set colorId.
 * Description now includes explicit URL text.
 */
function _insertEventAdvanced(calendarId, eventDetails, location, sourceUrl, colorId) {
  const endDateTime = new Date(eventDetails.startTime.getTime() + 2 * 60 * 60 * 1000);
  const resource = {
    summary: eventDetails.title,
    description: `See the complete season schedule here: ${sourceUrl}`,
    location: location,
    start: { dateTime: eventDetails.startTime.toISOString() },
    end:   { dateTime: endDateTime.toISOString() },
    colorId: String(colorId)
  };
  return Calendar.Events.insert(resource, calendarId);
}

/**
 * Insert event via CalendarApp and set mapped EventColor.
 * Description now includes explicit URL text.
 */
function _insertEventFallback(calendar, eventDetails, location, colorId) {
  const endDateTime = new Date(eventDetails.startTime.getTime() + 2 * 60 * 60 * 1000);
  const ev = calendar.createEvent(eventDetails.title, eventDetails.startTime, endDateTime, {
    description: `See the complete season schedule here: ${eventDetails.sourceUrl || ''}`,
    location: location
  });
  const constName = colorIdToCalendarAppConst(colorId);
  if (CalendarApp.EventColor[constName]) ev.setColor(CalendarApp.EventColor[constName]);
  return ev;
}

/* =========================
   Calendar resolution
   ========================= */
function _resolveTargetCalendar_() {
  const useAdvanced = HAS_ADVANCED_CALENDAR();
  if (useAdvanced) {
    const list = Calendar.CalendarList.list().items || [];
    const calListEntry = list.find(c => c.summary === '🔐 sports events');
    if (!calListEntry) throw new Error(`Calendar named '🔐 sports events' was not found.`);
    return { useAdvanced, calendarId: calListEntry.id, calendar: null };
  } else {
    const cal = CalendarApp.getCalendarsByName('🔐 sports events')[0];
    if (!cal) throw new Error(`Calendar named '🔐 sports events' was not found.`);
    return { useAdvanced, calendarId: cal.getId(), calendar: cal };
  }
}

/* =========================
   Single-row create helper (usable by retry/fill-in)
   ========================= */
function _createEventForRow_(sheet, rowNum, resolver) {
  const ss = sheet.getParent();
  const parts = sheet.getName().split(':');
  const sheetInfo = {
    emoji: parts[0],
    country: parts[1],
    league: parts[2],
    team: parts[3],
    url: ss.getUrl() + '#gid=' + sheet.getSheetId()
  };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Ensure event_url column exists
  let eventUrlColIndex = headers.indexOf('event_url') + 1;
  if (eventUrlColIndex === 0) {
    eventUrlColIndex = headers.length + 1;
    sheet.getRange(1, eventUrlColIndex).setValue('event_url');
  }

  const tvCol = headers.indexOf('TV');

  // Determine team color (by sheet team)
  const team = TEAMS_BY_ABBR[String(sheetInfo.team || '').toUpperCase()];
  const teamHex = team && team.color ? team.color : '#0e632c';
  const colorId = teamHexToColorId(teamHex);

  try {
    const { title, startTime } = _getEventDetailsFromRow(rowData, sheetInfo, headers);

    // Duplicate check
    let duplicate = false, duplicateId = null;
    if (resolver.useAdvanced) {
      const startOfDay = new Date(startTime); startOfDay.setHours(0,0,0,0);
      const endOfDay = new Date(startOfDay); endOfDay.setDate(endOfDay.getDate() + 1);
      const resp = Calendar.Events.list(resolver.calendarId, {
        timeMin: startOfDay.toISOString(),
        timeMax: endOfDay.toISOString(),
        singleEvents: true,
        maxResults: 50,
        q: title
      });
      const hit = (resp.items || []).find(ev => ev.summary === title);
      if (hit) { duplicate = true; duplicateId = hit.id; }
    } else {
      const dayEvents = resolver.calendar.getEventsForDay(startTime);
      const hit = dayEvents.find(ev => ev.getTitle() === title);
      if (hit) { duplicate = true; duplicateId = hit.getId(); }
    }

    if (duplicate) {
      // overwrite on duplicate
      if (resolver.useAdvanced && duplicateId) {
        Calendar.Events.remove(resolver.calendarId, duplicateId);
      } else if (!resolver.useAdvanced && duplicateId) {
        const evs = resolver.calendar.getEventsForDay(startTime).filter(e => e.getTitle() === title);
        evs.forEach(e => e.deleteEvent());
      }
    }

    const location = tvCol >= 0 ? (rowData[tvCol] || 'Local broadcaster') : 'Local broadcaster';
    let eventUrl = '';

    if (resolver.useAdvanced) {
      const inserted = _insertEventAdvanced(resolver.calendarId, { title, startTime }, location, sheetInfo.url, colorId);
      eventUrl = inserted && inserted.htmlLink ? inserted.htmlLink : '';
    } else {
      const created = _insertEventFallback(resolver.calendar, { title, startTime, sourceUrl: sheetInfo.url }, location, colorId);
      eventUrl = _eventLink(resolver.calendar, created);
    }

    const cell = sheet.getRange(rowNum, eventUrlColIndex);
    cell.setValue(eventUrl).setFontColor('#4285f4').setFontLine('none');
    _styleComputedColumn(sheet, eventUrlColIndex, sheet.getLastRow());
    return true;
  } catch (err) {
    sheet.getRange(rowNum, eventUrlColIndex).setValue(`ERROR: ${err.message}`);
    _styleComputedColumn(sheet, eventUrlColIndex, sheet.getLastRow());
    return false;
  }
}

/* =========================
   Core processing for a single sheet
   ========================= */
function _processSingleSheet_(sheet, ui, useAdvanced, calendarId, calendar) {
  const ss = sheet.getParent();
  const sheetName = sheet.getName();

  // Parse sheet name like '🏀:us:nba:gsw'
  const parts = sheetName.split(':');
  const sheetInfo = {
    emoji: parts[0],
    country: parts[1],
    league: parts[2],
    team: parts[3],
    url: ss.getUrl() + '#gid=' + sheet.getSheetId()
  };

  // Determine color by sheet team
  const team = TEAMS_BY_ABBR[String(sheetInfo.team || '').toUpperCase()];
  const teamHex = team && team.color ? team.color : '#0e632c';
  const colorId = teamHexToColorId(teamHex);

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (values.length === 0) return { created: 0, skipped: 0 };

  const headers = values.shift(); // row 1

  // Ensure 'event_url' column exists
  let eventUrlColIndex = headers.indexOf('event_url') + 1;
  if (eventUrlColIndex === 0) {
    eventUrlColIndex = headers.length + 1;
    sheet.getRange(1, eventUrlColIndex).setValue('event_url');
  }

  let createdCount = 0, skippedCount = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    try {
      const { title, startTime } = _getEventDetailsFromRow(row, sheetInfo, headers);

      // Duplicate check
      let isDuplicate = false;
      if (useAdvanced) {
        const startOfDay = new Date(startTime); startOfDay.setHours(0,0,0,0);
        const endOfDay = new Date(startOfDay); endOfDay.setDate(endOfDay.getDate() + 1);
        const resp = Calendar.Events.list(calendarId, {
          timeMin: startOfDay.toISOString(),
          timeMax: endOfDay.toISOString(),
          singleEvents: true,
          maxResults: 50,
          q: title
        });
        isDuplicate = (resp.items || []).some(ev => ev.summary === title);
      } else {
        const dayEvents = calendar.getEventsForDay(startTime);
        isDuplicate = dayEvents.some(ev => ev.getTitle() === title);
      }
      if (isDuplicate) { skippedCount++; continue; }

      const tvCol = headers.indexOf('TV');
      const location = tvCol >= 0 ? (row[tvCol] || 'Local broadcaster') : 'Local broadcaster';
      let htmlLink = '';

      if (useAdvanced) {
        const inserted = _insertEventAdvanced(calendarId, { title, startTime }, location, sheetInfo.url, colorId);
        htmlLink = inserted && inserted.htmlLink ? inserted.htmlLink : '';
      } else {
        const created = _insertEventFallback(calendar, { title, startTime, sourceUrl: sheetInfo.url }, location, colorId);
        htmlLink = _eventLink(calendar, created);
      }

      sheet.getRange(i + 2, eventUrlColIndex).setValue(htmlLink).setFontColor('#4285f4').setFontLine('none');
      createdCount++;
    } catch (err) {
      sheet.getRange(i + 2, eventUrlColIndex).setValue(`ERROR: ${err.message}`);
    }
  }

  _styleComputedColumn(sheet, eventUrlColIndex, sheet.getLastRow());
  return { created: createdCount, skipped: skippedCount };
}

/* =========================
   Entry points: Single sheet create / single row create / multi-row create
   ========================= */

function createAllCalendarEvents() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const resolver = _resolveTargetCalendar_();

  const res = _processSingleSheet_(sheet, ui, resolver.useAdvanced, resolver.calendarId, resolver.calendar);
  ui.alert(`Finished: ${sheet.getName()}\nCreated ${res.created} events. Skipped ${res.skipped} duplicates.`);
}

function createSingleCalendarEvent() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const rowNum = sheet.getActiveRange().getRow();

  if (rowNum === 1) { ui.alert('Please select a data row, not the header row.'); return; }

  const resolver = _resolveTargetCalendar_();

  const proceed = ui.alert('Create Event?', `Create a calendar event for row ${rowNum}?`, ui.ButtonSet.YES_NO);
  if (proceed !== ui.Button.YES) return;

  const ok = _createEventForRow_(sheet, rowNum, resolver);
  ui.alert(ok ? 'Success' : 'Failure', ok ? 'Created.' : 'Failed. See event_url cell for error.', ui.ButtonSet.OK);
}

function createMultipleGamesCalendarEvents() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const selection = ss.getSelection();
  const rangeList = selection.getActiveRangeList();
  const ranges = rangeList ? rangeList.getRanges() : (selection.getActiveRange() ? [selection.getActiveRange()] : []);
  if (!ranges.length) { ui.alert('No selection', 'Select one or more rows (ranges).', ui.ButtonSet.OK); return; }

  const resolver = _resolveTargetCalendar_();

  // Collect unique row numbers excluding header row
  const rows = new Set();
  ranges.forEach(r => {
    const start = r.getRow();
    const end = r.getLastRow();
    for (let rn = start; rn <= end; rn++) {
      if (rn !== 1) rows.add(rn);
    }
  });

  if (rows.size === 0) { ui.alert('Nothing to do', 'Header row cannot be processed.', ui.ButtonSet.OK); return; }

  let created = 0, failed = 0;
  Array.from(rows).sort((a,b)=>a-b).forEach(rowNum => {
    const ok = _createEventForRow_(sheet, rowNum, resolver);
    if (ok) created++; else failed++;
  });

  ui.alert(`Done. Created ${created} event(s). Failed ${failed}.`);
}

/* =========================
   Entry points: All sheets / Multiple sheets (UI) / Delete flows / Fill-in ERRORs
   ========================= */

function createAllSheetsCalendarEvents() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const resolver = _resolveTargetCalendar_();

  let totalCreated = 0, totalSkipped = 0;
  for (let i = 0; i < sheets.length; i++) {
    const s = sheets[i];
    const res = _processSingleSheet_(s, ui, resolver.useAdvanced, resolver.calendarId, resolver.calendar);
    totalCreated += res.created;
    totalSkipped += res.skipped;

    // Quota-friendly pause between sheets if requested later; currently immediate.
    // Utilities.sleep(10000); // enable when you want the 10s pause between sheets
  }

  ui.alert(`Finished processing ${sheets.length} sheet(s).\nCreated ${totalCreated} events. Skipped ${totalSkipped} duplicates.`);
}

function deleteAllSheetsSeasonEvents() {
  const ui = SpreadsheetApp.getUi();
  let resolver;
  try { resolver = _resolveTargetCalendar_(); }
  catch (e) { ui.alert('Error', e.message, ui.ButtonSet.OK); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const confirm = ui.alert('Confirm delete', 'This will scan ALL sheets and delete matching season game events. Continue?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  let totalDeleted = 0;
  for (const sheet of sheets) {
    const res = _deleteSeasonForSheet_(sheet, resolver);
    totalDeleted += res.deleted;
  }
  ui.alert(`Done. Deleted ${totalDeleted} event(s) across ${sheets.length} sheet(s).`);
}

function createMultipleSheetsSeasonEvents() {
  _showSheetsPicker_({ mode: 'create' });
}
function deleteMultipleSheetsSeasonEvents() {
  _showSheetsPicker_({ mode: 'delete' });
}

function fillInErrorEventsAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const resolver = _resolveTargetCalendar_();

  let retried = 0, created = 0, failed = 0;
  for (const sheet of sheets) {
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) continue;

    const headers = values[0];
    let eventUrlColIndex = headers.indexOf('event_url') + 1;
    if (eventUrlColIndex === 0) continue; // nothing to retry if no column

    for (let r = 2; r <= sheet.getLastRow(); r++) {
      const v = sheet.getRange(r, eventUrlColIndex).getValue();
      if (typeof v === 'string' && /^ERROR:/i.test(v.trim())) {
        retried++;
        const ok = _createEventForRow_(sheet, r, resolver);
        if (ok) created++; else failed++;
      }
    }
  }

  ui.alert(`Fill-in complete.\nRows retried: ${retried}\nCreated: ${created}\nFailed: ${failed}`);
}

/* =========================
   Delete helpers and entry points
   ========================= */

function _deleteSeasonForSheet_(sheet, resolver) {
  const ss = sheet.getParent();
  const sheetName = sheet.getName();
  const parts = sheetName.split(':');
  const sheetInfo = {
    emoji: parts[0],
    country: parts[1],
    league: parts[2],
    team: parts[3],
    url: ss.getUrl() + '#gid=' + sheet.getSheetId()
  };

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (!values.length) return { deleted: 0 };

  const headers = values.shift();
  const eventUrlColIndex = headers.indexOf('event_url') + 1; // 0 if missing

  let deleted = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    try {
      const { title, startTime } = _getEventDetailsFromRow(row, sheetInfo, headers);

      const startOfDay = new Date(startTime); startOfDay.setHours(0,0,0,0);
      const endOfDay = new Date(startOfDay); endOfDay.setDate(endOfDay.getDate() + 1);

      let rowDeleted = 0;

      if (resolver.useAdvanced) {
        const resp = Calendar.Events.list(resolver.calendarId, {
          timeMin: startOfDay.toISOString(),
          timeMax: endOfDay.toISOString(),
          singleEvents: true,
          maxResults: 100,
          q: title
        });
        const hits = (resp.items || []).filter(ev => ev.summary === title);
        for (const ev of hits) {
          Calendar.Events.remove(resolver.calendarId, ev.id);
          rowDeleted++;
        }
      } else {
        const events = resolver.calendar.getEventsForDay(startTime)
          .filter(ev => ev.getTitle() === title);
        events.forEach(ev => { ev.deleteEvent(); rowDeleted++; });
      }

      // Clear event_url cell if we deleted any
      if (rowDeleted > 0 && eventUrlColIndex > 0) {
        sheet.getRange(i + 2, eventUrlColIndex).clearContent();
      }

      deleted += rowDeleted;
    } catch (err) {
      // ignore row-level errors
    }
  }
  return { deleted };
}

function deleteSingleSheetSeasonEvents() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  let resolver;
  try { resolver = _resolveTargetCalendar_(); }
  catch (e) { ui.alert('Error', e.message, ui.ButtonSet.OK); return; }

  const confirm = ui.alert('Confirm delete', `This will scan the active sheet ("${sheet.getName()}") and delete matching season game events. Continue?`, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const res = _deleteSeasonForSheet_(sheet, resolver);
  ui.alert(`Done. Deleted ${res.deleted} event(s) from "${sheet.getName()}".`);
}

function deleteSingleGameCalendarEvent() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const activeRange = sheet.getActiveRange();
  if (!activeRange) { ui.alert('No selection', 'Select a single data row first.', ui.ButtonSet.OK); return; }
  const rowNum = activeRange.getRow();
  if (rowNum === 1) { ui.alert('Invalid selection', 'Please select a data row, not the header.', ui.ButtonSet.OK); return; }

  let resolver;
  try { resolver = _resolveTargetCalendar_(); }
  catch (e) { ui.alert('Error', e.message, ui.ButtonSet.OK); return; }

  const parts = sheet.getName().split(':');
  const sheetInfo = {
    emoji: parts[0],
    country: parts[1],
    league: parts[2],
    team: parts[3],
    url: ss.getUrl() + '#gid=' + sheet.getSheetId()
  };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];

  const eventUrlColIndex = headers.indexOf('event_url') + 1; // 0 if missing

  let deleted = 0;
  try {
    const { title, startTime } = _getEventDetailsFromRow(rowData, sheetInfo, headers);

    const confirm = ui.alert('Confirm delete', `Delete the calendar event for row ${rowNum}?\n\nTitle:\n${title}`, ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) return;

    const startOfDay = new Date(startTime); startOfDay.setHours(0,0,0,0);
    const endOfDay = new Date(startOfDay); endOfDay.setDate(endOfDay.getDate() + 1);

    if (resolver.useAdvanced) {
      const resp = Calendar.Events.list(resolver.calendarId, {
        timeMin: startOfDay.toISOString(),
        timeMax: endOfDay.toISOString(),
        singleEvents: true,
        maxResults: 50,
        q: title
      });
      const hits = (resp.items || []).filter(ev => ev.summary === title);
      for (const ev of hits) { Calendar.Events.remove(resolver.calendarId, ev.id); deleted++; }
    } else {
      const events = resolver.calendar.getEventsForDay(startTime)
        .filter(ev => ev.getTitle() === title);
      events.forEach(ev => { ev.deleteEvent(); deleted++; });
    }

    if (deleted > 0 && eventUrlColIndex > 0) {
      sheet.getRange(rowNum, eventUrlColIndex).clearContent();
    }
  } catch (err) {
    ui.alert('Error', err.message, ui.ButtonSet.OK);
    return;
  }

  ui.alert(`Done. Deleted ${deleted} event(s) for the selected row.`);
}

/* =========================
   Multiple sheets picker (HtmlService)
   ========================= */

function _showSheetsPicker_({ mode }) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(s => ({ id: s.getSheetId(), name: s.getName() }));

  const html = HtmlService.createHtmlOutputFromFile('SheetsPickerTemplate')
    .setWidth(420)
    .setHeight(520);

  html.append(`<script>
    window.__SHEETS__ = ${JSON.stringify(sheets)};
    window.__MODE__ = ${JSON.stringify(mode)};
  </script>`);

  SpreadsheetApp.getUi().showModalDialog(html, 'Which of the following sheets will be processed?');
}

function _processMultipleSheetsCreate_(sheetIds) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const picked = sheetIds.map(id => ss.getSheets().find(s => s.getSheetId() === id)).filter(Boolean);
  if (!picked.length) { ui.alert('No sheets selected.'); return; }

  const resolver = _resolveTargetCalendar_();

  // Ensure fallback path includes explicit URL in description
  const originalInsertFallback = _insertEventFallback;
  _insertEventFallback = function (cal, eventDetails, location, colorId) {
    if (!eventDetails.sourceUrl && eventDetails.sheet) {
      eventDetails.sourceUrl = eventDetails.sheet.getParent().getUrl() + '#gid=' + eventDetails.sheet.getSheetId();
    }
    return originalInsertFallback(cal, eventDetails, location, colorId);
  };

  let totalCreated = 0, totalSkipped = 0;
  for (const sheet of picked) {
    const res = _processSingleSheet_(sheet, ui, resolver.useAdvanced, resolver.calendarId, resolver.calendar);
    totalCreated += res.created;
    totalSkipped += res.skipped;

    // Quota-friendly pause between sheets (enable when needed)
    // Utilities.sleep(10000);
  }

  _insertEventFallback = originalInsertFallback;
  ui.alert(`Multiple sheets processed: ${picked.length}\nCreated ${totalCreated} events. Skipped ${totalSkipped} duplicates.`);
}

function _processMultipleSheetsDelete_(sheetIds) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const picked = sheetIds.map(id => ss.getSheets().find(s => s.getSheetId() === id)).filter(Boolean);
  if (!picked.length) { ui.alert('No sheets selected.'); return; }

  const confirm = ui.alert('Confirm delete', `This will scan the selected ${picked.length} sheet(s) and delete matching season game events.\nContinue?`, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  let resolver;
  try { resolver = _resolveTargetCalendar_(); }
  catch (e) { ui.alert('Error', e.message, ui.ButtonSet.OK); return; }

  let totalDeleted = 0;
  for (const sheet of picked) {
    const res = _deleteSeasonForSheet_(sheet, resolver);
    totalDeleted += res.deleted;
  }

  ui.alert(`Multiple sheets processed: ${picked.length}\nDeleted ${totalDeleted} event(s).`);
}