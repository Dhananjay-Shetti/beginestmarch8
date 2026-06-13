/* Code.gs - Beginest backend (updated) */
const SPREADSHEET_ID = '1F9QVgTgkadY-9uYsn6fc04vrzAsJgPgjvAEeVg_PK5k'; // <-- Put your Google Sheet ID here
const SHEET_CUSTOMER = 'Data';
const SHEET_ROOMMAP = 'HarborMeetingRoomMap';
const SHEET_BOOKING = 'BookingLog';
const SHEET_CREDITS = 'CompanyCreditsMap';

let _spreadsheet = null;

function getSpreadsheet() {
  if (!_spreadsheet) {
    _spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return _spreadsheet;
}

function jsonResponse(obj, statusCode) {
  const out = ContentService.createTextOutput(JSON.stringify(Object.assign({ status: statusCode || 200 }, obj)));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}

function hashPassword(password) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password), Utilities.Charset.UTF_8);
  return rawHash.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function doGet(e) {
  const action = e.parameter && e.parameter.action;
  if (!action) return jsonResponse({ error: 'Missing action' }, 400);
  try {
    if (action === 'getCustomerDetails') return getCustomerDetails(e.parameter.whatsappNumber);
    if (action === 'checkPassword') return checkPassword(e.parameter.whatsappNumber, e.parameter.password);
    if (action === 'getRoomsByHarbor') return getRoomsByHarbor(e.parameter.harbor);
    if (action === 'getRoomByNameAndHarbor') return getRoomByNameAndHarbor(e.parameter.roomName, e.parameter.harbor);
    if (action === 'getBookingsForRoomDate') return getBookingsForRoomDate(e.parameter.room, e.parameter.date);
    if (action === 'getAvailableRooms') return getAvailableRooms(e.parameter.harbor, e.parameter.date, e.parameter.start, e.parameter.end);
	  if (action === 'getBookingsByWhatsapp') return getBookingsByWhatsapp(e.parameter.whatsapp);
	  if (action === 'cancelBooking') return cancelBookingById(e.parameter.id);
    if (action === 'getBookingsByWhatsappCombined') return getBookingsByWhatsappCombined(e.parameter.whatsapp, parseInt(e.parameter.pastDays || '30', 10));
    if (action === 'checkEndCandidates') return checkEndCandidates(e.parameter.harbor || '', e.parameter.date || '', e.parameter.start || '', e.parameter.ends || '');
    if (action === 'deductCompanyCredits') return deductCompanyCreditsByParams(e.parameter); // allow GET too
    if (action === 'getCompanyCredits') return getCompanyCredits(e.parameter.company, e.parameter.harbor);;
    return jsonResponse({ error: 'Invalid action' }, 400);
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

function doPost(e) {
  const action = e.parameter && e.parameter.action;
  if (!action) return jsonResponse({ error: 'Missing action' }, 400);
  try {
    if (action === 'setPassword') return setPassword(e.parameter.whatsappNumber, e.parameter.password);
    if (action === 'saveBooking') return saveBooking(e.parameter);
	  if (action === 'cancelBooking') return cancelBookingById(e.parameter.id);
    if (action === 'deductCompanyCredits') return deductCompanyCreditsByParams(e.parameter);
    return jsonResponse({ error: 'Invalid action' }, 400);
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

/* ---------- Customers ---------- */
function getCustomerDetails(whatsapp) {
  if (!whatsapp) return jsonResponse({ message: 'Missing whatsappNumber' }, 400);
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CUSTOMER);
  if (!sheet) return jsonResponse({ message: 'No Data sheet' }, 200);
  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    for (let c = 0; c < headers.length; c++) obj[headers[c]] = row[c];
    const stored = String(obj['Whatsapp Number'] || '').trim();
    if (stored.replace(/\D/g, '') === String(whatsapp).replace(/\D/g, '')) {
      obj.WhatsappNumber = stored;
      return jsonResponse({ message: 'Customer found', customer: obj }, 200);
    }
  }
  return jsonResponse({ message: 'Not a customer' }, 200);
}

function checkPassword(whatsapp, password) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CUSTOMER);
  if (!sheet) return jsonResponse({ message: 'No Data sheet' }, 200);
  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  const headers = data[0];
  let passColIndex = headers.findIndex(h => String(h).toLowerCase().includes('password'));
  if (passColIndex === -1) passColIndex = 8;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4]).replace(/\D/g, '') === String(whatsapp).replace(/\D/g, '')) {
      const storedHash = String(data[i][passColIndex] || '').trim();
      const incomingHash = hashPassword(password);
      if (storedHash && storedHash === incomingHash) return jsonResponse({ message: 'OK' }, 200);
      return jsonResponse({ message: 'Invalid' }, 200);
    }
  }
  return jsonResponse({ message: 'Not a customer' }, 200);
}

function setPassword(whatsapp, password) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CUSTOMER);
  if (!sheet) return jsonResponse({ success: false, error: 'No Data sheet' }, 200);
  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  const headers = data[0];
  let passColIndex = headers.findIndex(h => String(h).toLowerCase().includes('password'));
  if (passColIndex === -1) passColIndex = 8;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4]).replace(/\D/g, '') === String(whatsapp).replace(/\D/g, '')) {
      sheet.getRange(i + 1, passColIndex + 1).setValue(hashPassword(password));
      return jsonResponse({ success: true }, 200);
    }
  }
  return jsonResponse({ success: false, error: 'Whatsapp not found' }, 200);
}

/* ---------- Rooms ---------- */
function getRoomsByHarbor(harbor) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ROOMMAP);
  if (!sheet) return jsonResponse({ rooms: [] }, 200);
  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (harbor && String(data[i][0]) !== String(harbor)) continue;
    let details = {};
    try { details = JSON.parse(data[i][2]); } catch (e) { details = {}; }
    if (!details.isAvailable) continue;
    out.push({ Harbor: data[i][0], MeetingRoom: data[i][1], ...details });
  }
  return jsonResponse({ rooms: out }, 200);
}

function getRoomByNameAndHarbor(roomName, harbor) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ROOMMAP);
  if (!sheet) return jsonResponse({ message: 'Room not found' }, 200);
  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  for (let i = 1; i < data.length; i++) {
    if ((harbor && String(data[i][0]) !== String(harbor)) && (String(data[i][1]) !== String(roomName))) continue;
    if (String(data[i][1]) === String(roomName) && (!harbor || String(data[i][0]) === String(harbor))) {
      let details = {};
      try { details = JSON.parse(data[i][2]); } catch (e) { details = {}; }
      return jsonResponse({ room: { Harbor: data[i][0], MeetingRoom: data[i][1], ...details } }, 200);
    }
  }
  return jsonResponse({ message: 'Room not found' }, 200);
}

/* ---------- Bookings ---------- */
function getBookingsForRoomDate(room, date) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_BOOKING);
  if (!sheet) return jsonResponse({ bookings: [] }, 200);
  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  const headers = data[0];
  const timeslotCol = headers.findIndex(h => String(h).toLowerCase().includes('booked timeslot'));
  const roomCol = headers.findIndex(h => String(h).toLowerCase().includes('meeting room'));
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const rowRoom = String(data[i][roomCol] || '');
    const timeslot = String(data[i][timeslotCol] || '');
    if (rowRoom === String(room) && timeslot.indexOf(date) !== -1) {
      out.push(timeslot);
    }
  }
  return jsonResponse({ bookings: out }, 200);
}

/* Expand a booked-timeslot string into half-hour slots (HH:MM) for that date.
   Accepts a timeslot string like "2025-09-15 15:00 to 17:00" or similar.
*/
function expandBookedSlotsFromString(timeslotStr, dateStr) {
  if (!timeslotStr) return [];
  // Find times in string (HH:MM)
  const re = /(\d{1,2}:\d{2})/g;
  const matches = [];
  let m;
  while ((m = re.exec(timeslotStr)) !== null) matches.push(m[1]);
  if (matches.length < 2) return [];
  const start = matches[0], end = matches[1];
  const sParts = start.split(':').map(Number);
  const eParts = end.split(':').map(Number);
  const sDate = new Date(dateStr + 'T' + String(sParts[0]).padStart(2,'0') + ':' + String(sParts[1]).padStart(2,'0') + ':00');
  const eDate = new Date(dateStr + 'T' + String(eParts[0]).padStart(2,'0') + ':' + String(eParts[1]).padStart(2,'0') + ':00');
  if (eDate <= sDate) return [];
  const out = [];
  const cur = new Date(sDate);
  while (cur < eDate) {
    out.push(String(cur.getHours()).padStart(2,'0') + ':' + String(cur.getMinutes()).padStart(2,'0'));
    cur.setMinutes(cur.getMinutes() + 30);
  }
  return out;
}

/* Check if requested slot (start,end) overlaps any existing booking ranges for that room on date */
function doesSlotConflictWithBookings(bookingsArray, dateStr, reqStart, reqEnd) {
  // Build requested slots
  const sParts = reqStart.split(':').map(Number);
  const eParts = reqEnd.split(':').map(Number);
  const sDate = new Date(dateStr + 'T' + String(sParts[0]).padStart(2,'0') + ':' + String(sParts[1]).padStart(2,'0') + ':00');
  const eDate = new Date(dateStr + 'T' + String(eParts[0]).padStart(2,'0') + ':' + String(eParts[1]).padStart(2,'0') + ':00');
  if (eDate <= sDate) return true; // invalid, treat as conflict
  // Expand all bookings into half-hour slots and check any intersection
  const reqSlots = [];
  const cur = new Date(sDate);
  while (cur < eDate) {
    reqSlots.push(String(cur.getHours()).padStart(2,'0') + ':' + String(cur.getMinutes()).padStart(2,'0'));
    cur.setMinutes(cur.getMinutes() + 30);
  }
  for (let ts of bookingsArray) {
    const bookedSlots = expandBookedSlotsFromString(ts, dateStr);
    for (let slot of reqSlots) if (bookedSlots.indexOf(slot) !== -1) return true;
  }
  return false;
}

/* getAvailableRooms: if date/start/end provided, return only rooms free for that slot.
   harbor param optional: if empty -> search all harbors, else filter by harbor.
*/
function getAvailableRooms(harbor, date, start, end) {

  const ss = getSpreadsheet();
  const roomSheet = ss.getSheetByName(SHEET_ROOMMAP);

  if (!roomSheet) return jsonResponse({ rooms: [] }, 200);

  const roomData = roomSheet
      .getRange(1,1,roomSheet.getLastRow(),roomSheet.getLastColumn())
      .getValues();

  const bookingIndex = buildBookingIndex(); // BUILD ONCE

  const out = [];

  for (let i = 1; i < roomData.length; i++) {

    const rowHarbor = String(roomData[i][0]);

    if (harbor && rowHarbor !== String(harbor)) continue;

    let details = {};

    try {
      details = JSON.parse(roomData[i][2]);
    } catch (e) {}

    if (!details.isAvailable) continue;

    const roomName = roomData[i][1];

    if (date && start && end) {

      const bookingsForRoom = bookingIndex[roomName + "|" + date] || [];

      if (doesSlotConflictWithBookings(bookingsForRoom, date, start, end)) {
        continue;
      }

    }

    out.push({
      Harbor: rowHarbor,
      MeetingRoom: roomName,
      ...details
    });

  }

  return jsonResponse({ rooms: out }, 200);
}

// internal helper that returns array of bookings for a room/date (not as HTTP response)
function getBookingsForRoomDateInternal(room, date) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_BOOKING);
  if (!sheet) return [];
  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  const headers = data[0];
  const timeslotCol = headers.findIndex(h => String(h).toLowerCase().includes('booked timeslot'));
  const roomCol = headers.findIndex(h => String(h).toLowerCase().includes('meeting room'));
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const rowRoom = String(data[i][roomCol] || '');
    const timeslot = String(data[i][timeslotCol] || '');
    if (rowRoom === String(room) && timeslot.indexOf(date) !== -1) {
      out.push(timeslot);
    }
  }
  return out;
}

function getCachedBookings() {

  const cache = CacheService.getScriptCache();
  const cached = cache.get("bookingData");

  if (cached) {
    return JSON.parse(cached);
  }

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_BOOKING);

  if (!sheet) return [];

  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();

  // cache for 30 seconds
  cache.put("bookingData", JSON.stringify(data), 30);

  return data;
}

function buildBookingIndex() {

  const cache = CacheService.getScriptCache();

  const cached = cache.get("bookingIndex");

  if (cached) {
    return JSON.parse(cached);
  }

  const bookingData = getCachedBookings();

  const headers = bookingData[0];

  const roomCol = headers.findIndex(h =>
    String(h).toLowerCase().includes("meeting room")
  );

  const timeslotCol = headers.findIndex(h =>
    String(h).toLowerCase().includes("booked timeslot")
  );

  const index = {};

  for (let i = 1; i < bookingData.length; i++) {

    const room = String(bookingData[i][roomCol] || "");
    const timeslot = String(bookingData[i][timeslotCol] || "");

    const dateMatch = timeslot.match(/\d{4}-\d{2}-\d{2}/);

    if (!dateMatch) continue;

    const date = dateMatch[0];

    const key = room + "|" + date;

    if (!index[key]) {
      index[key] = [];
    }

    index[key].push(timeslot);
  }

  cache.put("bookingIndex", JSON.stringify(index), 30);

  return index;
}

function saveBooking(params) {

  const lock = LockService.getScriptLock();
  lock.waitLock(5000); // wait up to 5 seconds for concurrent bookings

  try {

    // params expected: WhatsappNumber, MeetingRoom, date, start, end, Harbor (optional)
    const ss = getSpreadsheet();
    let bookingSheet = ss.getSheetByName(SHEET_BOOKING);
    const dataSheet = ss.getSheetByName(SHEET_CUSTOMER);

    // Find customer in Data sheet
    let customer = { firstName: '', lastName: '', gender: '', company: '', designation: '', harbor: '' };

    if (dataSheet) {
      const all = dataSheet.getRange(1,1,dataSheet.getLastRow(),dataSheet.getLastColumn()).getValues();
      const headers = all[0];

      for (let i = 1; i < all.length; i++) {

        const phoneCell = String(all[i][4] || '').replace(/\D/g, '');

        if (phoneCell === String(params.WhatsappNumber || '').replace(/\D/g, '')) {

          customer.firstName = all[i][1] || '';
          customer.lastName  = all[i][2] || '';
          customer.gender    = all[i][3] || '';
          customer.company   = all[i][5] || '';
          customer.designation = all[i][6] || '';

          const harborCol = headers.findIndex(h =>
            String(h).toLowerCase().includes('harbor')
          );

          if (harborCol !== -1) customer.harbor = all[i][harborCol] || '';

          break;
        }
      }
    }

    // Ensure booking sheet exists
    if (!bookingSheet) {
      bookingSheet = ss.insertSheet(SHEET_BOOKING);
      bookingSheet.appendRow([
        'ID','First Name','Last Name','Gender','Whatsapp Number',
        'Company Name','Designation','Meeting Room','Booked Timeslot','Booked On','Harbor'
      ]);
    }

    const id = Utilities.getUuid();
    const dateStr = params.date;

    const bookedTimeslot = `${dateStr} ${params.start} to ${params.end}`;
    const bookedOn = new Date();

    // ---------------------------------------------------
    // CRITICAL: CHECK ROOM AVAILABILITY AGAIN (ANTI RACE)
    // ---------------------------------------------------

    const existingBookings = getBookingsForRoomDateInternal(params.MeetingRoom, dateStr);

    if (doesSlotConflictWithBookings(existingBookings, dateStr, params.start, params.end)) {

      return jsonResponse({
        success:false,
        error:'This room was just booked by another user. Please refresh and choose another slot.'
      },200);

    }

    // ---- CREDITS LOGIC ----

    const companyName = customer.company || '';
    const harbor = customer.harbor || params.Harbor || '';

    const capacity = findRoomCapacity(params.MeetingRoom, harbor);

    let durationHours = 0;

    try {

      const sDt = new Date(dateStr + 'T' + params.start + ':00');
      const eDt = new Date(dateStr + 'T' + params.end + ':00');

      durationHours = (eDt - sDt) / (1000 * 60 * 60);

      if (isNaN(durationHours) || durationHours <= 0) durationHours = 0;

    } catch (e) {
      durationHours = 0;
    }

    const creditsNeeded = (capacity && durationHours) ? (capacity * durationHours) : 0;

    if (creditsNeeded > 0) {

      const rowIdx = findCompanyCreditsRowIndex(companyName, harbor);

      if (rowIdx === -1) {
        return jsonResponse({ success:false, error:'Company credits entry not found. Cannot book.' },200);
      }

      const creditsSheet = ss.getSheetByName(SHEET_CREDITS);
      const rem = Number(creditsSheet.getRange(rowIdx,2).getValue()) || 0;

      if (rem < creditsNeeded) {

        return jsonResponse({
          success:false,
          error:'Insufficient credits. Required: '+creditsNeeded+', Remaining: '+rem
        },200);

      }
    }

    const row = [
      id,
      customer.firstName,
      customer.lastName,
      customer.gender,
      params.WhatsappNumber,
      customer.company,
      customer.designation,
      params.MeetingRoom,
      bookedTimeslot,
      bookedOn,
      (customer.harbor || params.Harbor || '')
    ];

    // SAVE BOOKING
    bookingSheet.appendRow(row);
    CacheService.getScriptCache().remove("bookingIndex");

    // Deduct credits
    if (creditsNeeded > 0) {

      const upd = updateCompanyCredits(companyName, harbor, -creditsNeeded);

      if (!upd.success) {

        return jsonResponse({
          success:true,
          id:id,
          warning:'Booking saved but failed to deduct credits: '+(upd.error || '')
        },200);

      }

      return jsonResponse({
        success:true,
        id:id,
        remainingCredits:upd.remaining
      },200);

    }

    return jsonResponse({ success:true, id:id },200);

  } finally {

    lock.releaseLock();

  }
}



/* ---------- New endpoint ---------- */
/* Return bookings for a whatsapp number. Output: { bookings: [ {ID, FirstName, LastName, WhatsappNumber, MeetingRoom, BookedTimeslot, BookedOn, Harbor} ] } */
function getBookingsByWhatsapp(whatsapp) {
  if (!whatsapp) return jsonResponse({ bookings: [] }, 200);
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_BOOKING);
  if (!sheet) return jsonResponse({ bookings: [] }, 200);
  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  const headers = data[0].map(h => String(h || '').trim());
  // identify important columns
  const idCol = headers.findIndex(h => /id/i.test(h));
  const firstCol = headers.findIndex(h => /first/i.test(h));
  const lastCol = headers.findIndex(h => /last/i.test(h));
  const whatsappCol = headers.findIndex(h => /whatsapp/i.test(h));
  const roomCol = headers.findIndex(h => /meeting room/i.test(h));
  const timeslotCol = headers.findIndex(h => /booked timeslot/i.test(h));
  const bookedOnCol = headers.findIndex(h => /booked on/i.test(h));
  const harborCol = headers.findIndex(h => /harbor/i.test(h));

  const out = [];
  const normalized = String(whatsapp).replace(/\D/g, '');
  for (let i = 1; i < data.length; i++) {
    const rowPhone = String(data[i][whatsappCol] || '').replace(/\D/g, '');
    if (rowPhone === normalized) {
      out.push({
        ID: data[i][idCol] || '',
        FirstName: data[i][firstCol] || '',
        LastName: data[i][lastCol] || '',
        WhatsappNumber: data[i][whatsappCol] || '',
        MeetingRoom: data[i][roomCol] || '',
        BookedTimeslot: data[i][timeslotCol] || '',
        BookedOn: data[i][bookedOnCol] ? (data[i][bookedOnCol]) : '',
        Harbor: data[i][harborCol] || ''
      });
    }
  }
  return jsonResponse({ bookings: out }, 200);
}

/**
 * Cancel a booking by ID.
 * Expects: e.parameter.id (or POST body 'id')
 * Removes the row from the BookingLog sheet where first column equals the given id.
 */
function cancelBookingById(id) {
  if (!id) return jsonResponse({ success: false, error: 'Missing id' }, 400);

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_BOOKING);
  if (!sheet) return jsonResponse({ success: false, error: 'Booking sheet not found' }, 200);

  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  // header is data[0]
  const headers = data[0].map(h => String(h || '').trim().toLowerCase());
  // find indices to read values reliably
  const idCol = headers.findIndex(h => /id/.test(h));
  const companyCol = headers.findIndex(h => /company/i.test(h));
  const roomCol = headers.findIndex(h => /meeting room/i.test(h));
  const timeslotCol = headers.findIndex(h => /booked timeslot/i.test(h));
  const harborCol = headers.findIndex(h => /harbor/i.test(h));

  for (let r = 1; r < data.length; r++) {
    const rowId = String(data[r][idCol] || '').trim();
    if (rowId === String(id).trim()) {
      // before deleting, compute credits to refund
      try {
        const companyName = String(data[r][companyCol] || '').trim();
        const roomName = String(data[r][roomCol] || '').trim();
        const timeslot = String(data[r][timeslotCol] || '').trim();
        const harbor = String(data[r][harborCol] || '').trim();

        // parse timeslot like "YYYY-MM-DD HH:MM to HH:MM"
        const tsMatch = timeslot.match(/(\d{4}-\d{2}-\d{2})/);
        const dateStr = tsMatch ? tsMatch[1] : null;
        const times = timeslot.match(/(\d{1,2}:\d{2})/g) || [];
        const start = times[0] || '';
        const end = times[1] || '';

        let durationHours = 0;
        if (dateStr && start && end) {
          const sDt = new Date(dateStr + 'T' + start + ':00');
          const eDt = new Date(dateStr + 'T' + end + ':00');
          durationHours = (eDt - sDt) / (1000 * 60 * 60);
          if (isNaN(durationHours) || durationHours <= 0) durationHours = 0;
        }

        const capacity = findRoomCapacity(roomName, harbor);
        const creditsToRefund = (capacity && durationHours) ? (capacity * durationHours) : 0;
        if (creditsToRefund > 0 && companyName) {
          // attempt to add back credits
          try {
            updateCompanyCredits(companyName, harbor, creditsToRefund);
          } catch (e) {
            // swallow — we will still delete row; consider logging
          }
        }
      } catch (e) {
        // ignore errors computing refund
      }

      // delete the row (sheet rows are 1-indexed)
      sheet.deleteRow(r + 1);
      return jsonResponse({ success: true }, 200);
    }
  }

  return jsonResponse({ success: false, error: 'Booking ID not found' }, 200);
}


/**
 * getBookingsByWhatsappCombined(whatsapp, pastDays)
 * Returns { upcoming: [...], past30: [...] }
 * - `pastDays` controls the lookback window for the "past" bucket (value will be reflected in the `past30` array name for backward compatibility)
 *
 * Each booking object matches the format returned by getBookingsByWhatsapp (ID, FirstName, LastName, WhatsappNumber, MeetingRoom, BookedTimeslot, BookedOn, Harbor)
 */
function getBookingsByWhatsappCombined(whatsapp, pastDays) {
  // normalize
  if (!whatsapp) return jsonResponse({ upcoming: [], past30: [], pastDaysUsed: pastDays || 30 }, 200);
  pastDays = Number(pastDays) || 30;

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_BOOKING);
  if (!sheet) return jsonResponse({ upcoming: [], past30: [], pastDaysUsed: pastDays }, 200);

  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  if (!data || data.length < 2) return jsonResponse({ upcoming: [], past30: [], pastDaysUsed: pastDays }, 200);

  const headers = data[0].map(h => String(h || '').trim());
  const idCol = headers.findIndex(h => /id/i.test(h));
  const firstCol = headers.findIndex(h => /first/i.test(h));
  const lastCol = headers.findIndex(h => /last/i.test(h));
  const whatsappCol = headers.findIndex(h => /whatsapp/i.test(h));
  const roomCol = headers.findIndex(h => /meeting room/i.test(h));
  const timeslotCol = headers.findIndex(h => /booked timeslot/i.test(h));
  const bookedOnCol = headers.findIndex(h => /booked on/i.test(h));
  const harborCol = headers.findIndex(h => /harbor/i.test(h));

  const normalized = String(whatsapp).replace(/\D/g, '');
  const upcoming = [];
  const pastArr = [];

  // helper: parse timeslot strings like "2025-09-15 15:00 to 17:00"
  function parseTimeslotString(ts) {
    if (!ts) return null;
    // find date YYYY-MM-DD at start (or any \d{4}-\d{2}-\d{2})
    const dateMatch = ts.match(/(\d{4}-\d{2}-\d{2})/);
    const dateStr = dateMatch ? dateMatch[1] : null;
    const times = ts.match(/(\d{1,2}:\d{2})/g) || [];
    const start = times[0] || '';
    const end = times[1] || '';
    return { date: dateStr, start: start, end: end };
  }

  const now = new Date();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - pastDays);

  // iterate rows
  for (let r = 1; r < data.length; r++) {
    const rowPhone = String(data[r][whatsappCol] || '').replace(/\D/g, '');
    if (rowPhone !== normalized) continue;

    const rec = {
      ID: data[r][idCol] || '',
      FirstName: data[r][firstCol] || '',
      LastName: data[r][lastCol] || '',
      WhatsappNumber: data[r][whatsappCol] || '',
      MeetingRoom: data[r][roomCol] || '',
      BookedTimeslot: data[r][timeslotCol] || '',
      BookedOn: data[r][bookedOnCol] ? data[r][bookedOnCol] : '',
      Harbor: data[r][harborCol] || ''
    };

    // classify
    const tsObj = parseTimeslotString(rec.BookedTimeslot);
    if (!tsObj || !tsObj.date) {
      // if no valid timeslot, treat as past (safe fallback)
      pastArr.push(rec);
      continue;
    }

    // build end datetime for comparison
    try {
      // If times are missing, assume end at 00:00 => push to past as safe fallback
      if (!tsObj.end) { pastArr.push(rec); continue; }
      const endIso = tsObj.date + 'T' + tsObj.end + ':00';
      const endDt = new Date(endIso);
      if (isNaN(endDt.getTime())) {
        // fallback: push to past
        pastArr.push(rec);
        continue;
      }
      if (endDt < now) {
        // past — check cutoff
        if (endDt >= cutoff) pastArr.push(rec);
      } else {
        // upcoming
        upcoming.push(rec);
      }
    } catch (e) {
      // on parse error, put into past
      pastArr.push(rec);
    }
  }

  // keep API backward-compatible: field name past30 (even when pastDays differs)
  return jsonResponse({ upcoming: upcoming, past30: pastArr, pastDaysUsed: pastDays }, 200);
}


/* ---------- Company Credits: new helpers & endpoints ---------- */

/**
 * updateCompanyCreditsInternal(harbor, company, delta)
 * Internal helper: atomically updates RemainingCredits by adding delta (can be negative).
 * Returns { success: true, remaining } or { success: false, error }
 */
function updateCompanyCreditsInternal(harbor, company, delta) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('CompanyCreditsMap');
  if (!sheet) return { success: false, error: 'CompanyCreditsMap sheet not found' };

  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  if (!data || data.length < 2) return { success: false, error: 'No data' };

  const headers = data[0].map(h => String(h || '').trim().toLowerCase());
  const compCol = headers.findIndex(h => /company/i.test(h));
  const remCol  = headers.findIndex(h => /remainingcredits/i.test(h));
  const totCol  = headers.findIndex(h => /totalcredits/i.test(h));
  const harborCol = headers.findIndex(h => /harbor/i.test(h));

  for (let r = 1; r < data.length; r++) {
    const rowCompany = String(data[r][compCol] || '').trim();
    const rowHarbor  = String(data[r][harborCol] || '').trim();
    if (rowCompany === String(company).trim() && rowHarbor === String(harbor).trim()) {
      const current = Number(data[r][remCol] || 0);
      const updated = current + Number(delta);
      // Persist updated RemainingCredits
      sheet.getRange(r + 1, remCol + 1).setValue(updated);
      return { success: true, remaining: updated };
    }
  }
  return { success: false, error: 'Company/Harbor not found' };
}

/**
 * deductCompanyCreditsByParams(params)
 * HTTP handler (GET or POST) that expects: harbor, company, amount (positive number)
 * Returns JSON with success and remaining credits or error.
 */
function deductCompanyCreditsByParams(params) {
  params = params || {};
  const harbor = params.harbor || '';
  const company = params.company || '';
  const amount = Number(params.amount || 0);
  if (!harbor || !company || isNaN(amount)) return jsonResponse({ success: false, error: 'Missing params' }, 400);

  // We will deduct amount (so delta is -amount)
  const res = updateCompanyCreditsInternal(harbor, company, -Math.abs(amount));
  if (res.success) return jsonResponse({ success: true, remaining: res.remaining }, 200);
  return jsonResponse({ success: false, error: res.error }, 200);
}

/**
 * Try to determine room capacity (number of seats) from ROOMMAP details.
 * Looks for details.people or details.capacity (tries parseInt).
 * Returns integer capacity or null if unknown.
 */
function findRoomCapacity(roomName, harbor) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ROOMMAP);
  if (!sheet) return null;
  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  for (let i = 1; i < data.length; i++) {
    const rowHarbor = String(data[i][0] || '');
    const rowRoom = String(data[i][1] || '');
    if ((harbor && rowHarbor !== String(harbor)) && (rowRoom !== String(roomName))) continue;
    if (rowRoom === String(roomName) && (!harbor || rowHarbor === String(harbor))) {
      let details = {};
      try { details = JSON.parse(data[i][2]); } catch (e) { details = {}; }
      // try multiple keys
      if (details.people) {
        const n = parseInt(String(details.people).replace(/\D/g,''), 10);
        if (!isNaN(n) && n > 0) return n;
      }
      if (details.capacity) {
        const n = parseInt(String(details.capacity).replace(/\D/g,''), 10);
        if (!isNaN(n) && n > 0) return n;
      }
      // fallback: try to parse any numeric in details object values
      for (const k in details) {
        if (typeof details[k] === 'string') {
          const n = parseInt(details[k].replace(/\D/g,''), 10);
          if (!isNaN(n) && n > 0) return n;
        }
      }
      return null;
    }
  }
  return null;
}

/**
 * Find row index (1-based) in CompanyCreditsMap for given company+harbor.
 * Returns row index (1-based) if found, or -1 if not found.
 */
function findCompanyCreditsRowIndex(company, harbor) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CREDITS);
  if (!sheet) return -1;
  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  if (!data || data.length < 1) return -1;
  // header row = 0
  for (let r = 1; r < data.length; r++) {
    const rowCompany = String(data[r][0] || '').trim();
    const rowHarbor = String(data[r][3] || '').trim(); // expecting columns: Company, RemainingCredits, TotalCredits, Harbor
    if (rowCompany === String(company).trim() && (!harbor || rowHarbor === String(harbor).trim())) {
      return r + 1; // sheet row (1-indexed)
    }
  }
  return -1;
}

/**
 * Update RemainingCredits for company + harbor by delta (positive to add, negative to deduct).
 * Returns { success: boolean, remaining: number|null, error: string|null }
 */
function updateCompanyCredits(company, harbor, delta) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CREDITS);
  if (!sheet) return { success: false, remaining: null, error: 'Credits sheet not found' };
  const rowIndex = findCompanyCreditsRowIndex(company, harbor);
  if (rowIndex === -1) return { success: false, remaining: null, error: 'Company entry not found' };

  // Assuming header mapping: Company (col A), RemainingCredits (col B), TotalCredits (col C), Harbor (col D)
  const remainingCol = 2; // B (1-indexed)
  const cell = sheet.getRange(rowIndex, remainingCol);
  const cur = Number(cell.getValue()) || 0;
  const next = cur + Number(delta || 0);
  cell.setValue(next);
  return { success: true, remaining: next, error: null };
}

function getCompanyCredits(company, harbor) {
  if (!company) return jsonResponse({ error: 'missing company' }, 400);
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CREDITS);
  if (!sheet) return jsonResponse({ company: company, remaining: null }, 200);
  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  for (let i = 1; i < data.length; i++) {
    const rowCompany = String(data[i][0] || '').trim();
    const rowHarbor = String(data[i][3] || '').trim();
    if (rowCompany === company && (!harbor || rowHarbor === String(harbor))) {
      return jsonResponse({ company: company, remaining: Number(data[i][1]) || 0, total: Number(data[i][2]) || 0, harbor: rowHarbor }, 200);
    }
  }
  return jsonResponse({ company: company, remaining: null }, 200);
}

// Internal helper: similar to getAvailableRooms but returns array of room objects (not an HTTP response)
function getAvailableRoomsInternal(harbor, date, start, end) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ROOMMAP);
  if (!sheet) return [];
  const data = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const rowHarbor = String(data[i][0]);
    if (harbor && String(rowHarbor) !== String(harbor)) continue;
    let details = {};
    try { details = JSON.parse(data[i][2]); } catch (e) { details = {}; }
    if (!details.isAvailable) continue;
    const roomName = data[i][1];

    // If date/start/end provided -> check bookings for this room on that date
    if (date && start && end) {
      const bookingsResp = getBookingsForRoomDateInternal(roomName, date); // returns array
      if (doesSlotConflictWithBookings(bookingsResp, date, start, end)) {
        // conflict -> skip
        continue;
      }
    }
    out.push({ Harbor: rowHarbor, MeetingRoom: roomName, ...details });
  }
  return out;
}

/**
 * checkEndCandidates(harbor, date, start, ends)
 * - harbor: optional (empty = search all harbors)
 * - date: YYYY-MM-DD
 * - start: HH:MM
 * - ends: comma-separated end times "HH:MM,HH:MM,..." (client should only send candidates within MAX_DURATION window)
 *
 * Returns JSON: { results: [ { end: "HH:MM", ok: true|false }, ... ] }
 */
function checkEndCandidates(harbor, date, start, endsCSV) {
  if (!date || !start || !endsCSV) return jsonResponse({ results: [] }, 200);
  const ends = String(endsCSV).split(',').map(s => s.trim()).filter(s => s);
  const out = [];
  for (let i = 0; i < ends.length; i++) {
    const end = ends[i];
    try {
      const rooms = getAvailableRoomsInternal(harbor, date, start, end);
      out.push({ end: end, ok: (rooms && rooms.length > 0) });
    } catch (e) {
      out.push({ end: end, ok: false });
    }
  }
  return jsonResponse({ results: out }, 200);
}
