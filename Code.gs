/**
 * Appointment Booking System for Google Apps Script
 * 
 * Purpose: This system creates a flexible appointment booking interface that allows clients
 * to book appointments with variable durations based on the number of people attending.
 * The system integrates with Google Calendar and provides a user-friendly interface
 * similar to Google's appointment booking but with enhanced flexibility.
 * 
 * Method of Operation:
 * 1. Displays a web interface for clients to fill in their details
 * 2. Allows selection of number of attendees (1-9 people)
 * 3. Calculates appointment duration: 30 minutes for 1 person + 15 minutes for each additional person
 * 4. Shows available time slots based on business hours and existing calendar events
 * 5. Creates calendar events automatically when appointments are booked
 * 6. Sends confirmation emails to both client and business owner
 * 7. Generates a unique management token per appointment, stored in Script Properties,
 *    which allows clients to cancel or reschedule directly via a link in their confirmation email
 * 
 * @updated 2026-03-10 - Added appointment cancellation and rescheduling: token-based links in email/ICS/calendar,
 * OTP fallback via email, and async cleanup of expired tokens using time-based triggers
 * @updated 2025-10-16 - Added WhatsApp quick contact link in business calendar events for easy client communication
 * @updated 2025-07-28 - Fixed Google Calendar link timezone issue: removed incorrect timezone conversion
 * in createCalendarLink function. Now uses times directly as they work correctly in business calendar and emails
 * @updated 2025-06-09 - Fixed People API contact management: added 'metadata' to personFields
 * to resolve invalid 'etag' field error in people.connections.list API calls
 */

// ============================================================================
// CONFIGURATION - Fill in your business details below
// ============================================================================
const CONFIG = {
  // Business hours (24-hour format)
  // Set to null for days you don't work
  BUSINESS_HOURS: {
    SUNDAY: { start: '09:00', end: '15:15' , evening: { start: '18:30', end: '20:30' } },
    MONDAY: { start: '09:00', end: '15:15' , evening: { start: '18:30', end: '20:30' } },
    TUESDAY: { start: '09:00', end: '15:15' , evening: { start: '18:30', end: '20:30' } },
    WEDNESDAY: { start: '09:00', end: '15:15' , evening: { start: '18:30', end: '20:30' } },
    THURSDAY: { start: '09:00', end: '15:00' },
    FRIDAY: null, // Not working
    SATURDAY: null // Not working
  },
  
  // Calendar IDs to check for conflicts
  // 'primary' = your main Google Calendar
  // Add additional calendar IDs from Google Calendar settings if needed
  CALENDAR_IDS: ['primary'],
  
  // Calendar ID that requires buffer time before/after events (e.g., sports calendar)
  // Leave empty string if not needed
  SPECIAL_CALENDAR_ID: '',
  
  // Appointment duration settings
  BASE_DURATION: 30, // Minutes for first person
  ADDITIONAL_PERSON_DURATION: 15, // Additional minutes for each extra person
  
  // Business information
  BUSINESS_NAME: 'YOUR_BUSINESS_NAME_HE',           // Hebrew business name, e.g., 'העסק שלי'
  BUSINESS_NAME_EN: 'YOUR_BUSINESS_NAME_EN',         // English business name, e.g., 'My Business'
  BUSINESS_EMAIL: 'YOUR_EMAIL@gmail.com',             // Business email address
  BUSINESS_EMAIL_FROM: 'YOUR_EMAIL+meeting@gmail.com', // Email alias for sending (use Gmail + alias)
  BUSINESS_EMAIL_NAME: 'YOUR_BUSINESS_NAME_HE - קביעת פגישה', // Display name for outgoing emails
  BUSINESS_PHONE: 'YOUR_PHONE_NUMBER',               // e.g., '050-1234567'
  BUSINESS_PHONE_INTL: 'YOUR_PHONE_INTERNATIONAL',   // International format for WhatsApp, e.g., '972501234567'
  
  // Address
  BUSINESS_ADDRESS_HE: 'YOUR_ADDRESS_HE',            // Hebrew address, e.g., 'הרחוב 1, קומה 2, העיר'
  BUSINESS_ADDRESS_EN: 'YOUR_ADDRESS_EN',             // English address, e.g., 'Street 1, 2nd Floor, City'
  WAZE_LINK: 'YOUR_WAZE_NAVIGATION_LINK',            // Waze share link, e.g., 'https://waze.com/ul/...'
  
  // Event titles for calendar entries and emails
  EVENT_TITLE_HE: 'YOUR_EVENT_TITLE_HE',             // Hebrew event title, e.g., 'פגישה'
  EVENT_TITLE_EN: 'YOUR_EVENT_TITLE_EN',              // English event title, e.g., 'Appointment'

  // WhatsApp pre-filled messages (the client sees these when clicking the WhatsApp link)
  WHATSAPP_MSG_EN: 'YOUR_WHATSAPP_MESSAGE_EN',       // e.g., 'Hi, I booked an appointment for'
  WHATSAPP_MSG_HE: 'YOUR_WHATSAPP_MESSAGE_HE',       // e.g., 'היי, קבעתי פגישה ל'
  
  // ICS calendar settings
  ICS_DOMAIN: 'YOUR_DOMAIN.com',                     // Domain for ICS UIDs, e.g., 'mybusiness.com'
  
  // Time slot interval (in minutes)
  TIME_SLOT_INTERVAL: 15,

  // Management page base URL (where standalone.html is hosted)
  MANAGE_BASE_URL: 'https://YOUR_USERNAME.github.io/meeting.html'
};

/**
 * Serves the main HTML page for the appointment booking system
 * Also supports API mode when called with specific parameters
 *
 * API Mode Usage:
 * - ?action=getConfig - Returns system configuration
 * - ?action=getTimeSlots&date=YYYY-MM-DD&duration=45 - Returns available time slots
 * - ?action=getAppointmentByToken&token=<uuid> - Returns appointment details for management
 * - ?action=cancelAppointment&token=<uuid> - Cancels an appointment
 * - ?action=sendOTP&email=<email>&lang=<he|en> - Sends a one-time code to email for OTP auth
 * - ?action=verifyOTPAndGetAppointments&email=<email>&code=<code> - Verifies OTP and returns upcoming appointments
 * - ?action=diagnoseDoubleBooking&date=YYYY-MM-DD - Runs diagnostic for double booking
 * - ?action=listCalendars - Lists all available calendars
 *
 * @param {Object} e - Event object containing query parameters
 * @return {HtmlOutput|TextOutput} HTML page or JSON response
 */
function doGet(e) {
  // Check if this is an API request
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e.parameter);
  }

  // Regular HTML page serving
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('קביעת פגישה')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handles POST requests for the API
 * Used for booking appointments from external websites
 * 
 * @param {Object} e - Event object containing POST data
 * @return {TextOutput} JSON response
 */
function doPost(e) {
  try {
    // Parse the POST data
    let data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    } else {
      return createJsonResponse({ success: false, message: 'No data received' });
    }
    
    const action = data.action;
    
    switch (action) {
      case 'bookAppointment':
        const result = bookAppointment(data.appointmentData);
        return createJsonResponse(result);
        
      case 'getConfig':
        return createJsonResponse({ success: true, config: getConfig() });
        
      case 'getTimeSlots':
        const slots = getAvailableTimeSlots(data.date, data.duration);
        return createJsonResponse({ success: true, timeSlots: slots });

      case 'cancelAppointment':
        return createJsonResponse(cancelAppointment(data.token));

      case 'rescheduleAppointment':
        return createJsonResponse(rescheduleAppointment(data.token, data.newDate, data.newSlot, data.newPeople || null));

      case 'sendOTP':
        return createJsonResponse(sendOTP(data.email, data.lang || 'he'));

      case 'verifyOTPAndGetAppointments':
        return createJsonResponse(verifyOTPAndGetAppointments(data.email, data.code));
        
      default:
        return createJsonResponse({ success: false, message: 'Unknown action: ' + action });
    }
    
  } catch (error) {
    console.error('doPost error:', error);
    return createJsonResponse({ success: false, message: error.message });
  }
}

/**
 * Handles API GET requests
 * Supports both simple parameters and encoded JSON data for complex requests
 * 
 * @param {Object} params - Query parameters
 * @return {TextOutput} JSON response
 */
function handleApiRequest(params) {
  try {
    const action = params.action;
    
    // Check if there's encoded data (for complex requests like booking)
    let data = null;
    if (params.data) {
      try {
        data = JSON.parse(decodeURIComponent(params.data));
      } catch (e) {
        console.error('Failed to parse data parameter:', e);
      }
    }
    
    switch (action) {
      case 'getConfig':
        return createJsonResponse({ success: true, config: getConfig() });
        
      case 'getTimeSlots':
        if (!params.date || !params.duration) {
          return createJsonResponse({ success: false, message: 'Missing date or duration parameter' });
        }
        const slots = getAvailableTimeSlots(params.date, parseInt(params.duration));
        return createJsonResponse({ success: true, timeSlots: slots });

      case 'getDailyAvailability':
        if (!params.date) {
           return createJsonResponse({ success: false, message: 'Missing date parameter' });
        }
        const availability = getDailyAvailability(params.date);
        return createJsonResponse({ success: true, ...availability });

      case 'diagnoseDoubleBooking':
        const dateToDiagnose = params.date || new Date().toISOString().split('T')[0];
        diagnoseDoubleBooking(dateToDiagnose);
        return createJsonResponse({
          success: true,
          message: `Diagnosis started for date: ${dateToDiagnose}. Check logs for results.`
        });

      case 'listCalendars':
        listAllCalendars();
        return createJsonResponse({
          success: true,
          message: 'Calendar listing started. Check logs for results.'
        });

      case 'bookAppointment':
        // Handle booking via GET (for CORS compatibility with external hosting)
        if (data && data.appointmentData) {
          const result = bookAppointment(data.appointmentData);
          return createJsonResponse(result);
        } else {
          return createJsonResponse({ success: false, message: 'Missing appointment data' });
        }

      case 'getAppointmentByToken':
        if (!params.token) {
          return createJsonResponse({ success: false, message: 'Missing token parameter' });
        }
        return createJsonResponse(getAppointmentByToken(params.token));

      case 'cancelAppointment':
        if (!params.token) {
          return createJsonResponse({ success: false, message: 'Missing token parameter' });
        }
        return createJsonResponse(cancelAppointment(params.token));

      case 'rescheduleAppointment':
        if (!params.token || !params.newDate || !params.newSlot) {
          return createJsonResponse({ success: false, message: 'Missing token, newDate or newSlot parameter' });
        }
        return createJsonResponse(rescheduleAppointment(
          params.token,
          params.newDate,
          JSON.parse(params.newSlot),
          params.newPeople ? parseInt(params.newPeople) : null
        ));

      case 'sendOTP':
        if (!params.email) {
          return createJsonResponse({ success: false, message: 'Missing email parameter' });
        }
        return createJsonResponse(sendOTP(params.email, params.lang || 'he'));

      case 'verifyOTPAndGetAppointments':
        if (!params.email || !params.code) {
          return createJsonResponse({ success: false, message: 'Missing email or code parameter' });
        }
        return createJsonResponse(verifyOTPAndGetAppointments(params.email, params.code));
        
      default:
        return createJsonResponse({ success: false, message: 'Unknown action: ' + action });
    }
    
  } catch (error) {
    console.error('handleApiRequest error:', error);
    return createJsonResponse({ success: false, message: error.message });
  }
}

/**
 * Creates a JSON response with proper CORS headers
 * This allows the API to be called from any external website
 * 
 * @param {Object} data - The data to return as JSON
 * @return {TextOutput} JSON response with CORS headers
 */
function createJsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * Includes external files (CSS, JS) in the HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Gets available time slots for a specific date
 * @param {string} dateStr - Date in YYYY-MM-DD format
 * @param {number} duration - Appointment duration in minutes
 * @return {Array} Array of available time slots
 */
function getAvailableTimeSlots(dateStr, duration) {
  try {
    // Special calendar ID that requires buffer time
    const specialCalendarId = CONFIG.SPECIAL_CALENDAR_ID;

    logger.info('Starting getAvailableTimeSlots', {
      requestedDate: dateStr,
      requestedDuration: duration,
      specialCalendarId: specialCalendarId,
      timestamp: new Date().toISOString()
    });

    // Parse date correctly for Israel timezone
    const dateParts = dateStr.split('-');
    const year = parseInt(dateParts[0]);
    const month = parseInt(dateParts[1]) - 1; // JavaScript months are 0-based
    const day = parseInt(dateParts[2]);

    // Create date in Israel timezone
    const date = new Date(year, month, day);

    console.log(`Getting slots for: ${dateStr}, parsed as: ${date.toDateString()}`);

    const dayOfWeek = getDayOfWeek(date);
    
    // Check if it's a working day
    const businessHours = CONFIG.BUSINESS_HOURS[dayOfWeek];
    console.log(`Day: ${dayOfWeek}, Business hours:`, businessHours);
    if (!businessHours) {
      console.log('Not a working day, returning empty slots');
      return []; // Not a working day
    }
    
    // Get existing calendar events for the day from all calendars
    const startOfDay = new Date(date);
    startOfDay.setHours(0, 0, 0, 0);
    const endOfDay = new Date(date);
    endOfDay.setHours(23, 59, 59, 999);
    
    console.log(`Checking events from ${startOfDay} to ${endOfDay}`);
    
    let allEvents = [];
    
    // Check events from all configured calendars
    CONFIG.CALENDAR_IDS.forEach(calendarId => {
      try {
        const calendar = CalendarApp.getCalendarById(calendarId);
        if (calendar) {
          const events = calendar.getEvents(startOfDay, endOfDay);
          console.log(`Calendar ${calendarId}: found ${events.length} events`);

          // Add calendar ID information to each event
          events.forEach(event => {
            const isAllDay = event.isAllDayEvent();
            const eventStart = event.getStartTime();
            const eventEnd = event.getEndTime();
            const eventTitle = event.getTitle();

            console.log(`Event: "${eventTitle}" from ${eventStart} to ${eventEnd} (All-day: ${isAllDay}) from calendar: ${calendarId}`);

            // Special logging for sports calendar events
            if (calendarId === specialCalendarId) {
              logger.info('Sports calendar event found', {
                title: eventTitle,
                startTime: eventStart ? eventStart.toISOString() : 'null',
                endTime: eventEnd ? eventEnd.toISOString() : 'null',
                isAllDay: isAllDay,
                calendarId: calendarId
              });
            }

            // Add the calendar ID as a property to the event object
            event._sourceCalendarId = calendarId;
          });

          allEvents = allEvents.concat(events);
        }
      } catch (error) {
        console.warn(`Could not access calendar ${calendarId}:`, error);
      }
    });
    
    const events = allEvents;
    console.log(`Total events found: ${events.length}`);
    
    // Generate all possible time slots
    const timeSlots = [];

    // Morning slots
    addTimeSlotsForPeriod(timeSlots, date, businessHours.start, businessHours.end, duration, events, specialCalendarId);

    // Evening slots (if exists)
    if (businessHours.evening) {
      addTimeSlotsForPeriod(timeSlots, date, businessHours.evening.start, businessHours.evening.end, duration, events, specialCalendarId);
    }
    
    console.log(`Found ${timeSlots.length} available slots:`, timeSlots);

    logger.info('Completed getAvailableTimeSlots', {
      requestedDate: dateStr,
      requestedDuration: duration,
      totalSlotsFound: timeSlots.length,
      availableSlots: timeSlots.map(slot => ({
        start: slot.start,
        end: slot.end,
        startDateTime: slot.startDateTime,
        endDateTime: slot.endDateTime
      })),
      totalEventsChecked: events.length,
      specialCalendarId: specialCalendarId
    });

    return timeSlots;

  } catch (error) {
    console.error('Error getting available time slots:', error);
    logger.error('Failed to get available time slots', {
      error: error.message,
      requestedDate: dateStr,
      requestedDuration: duration
    });
    return [];
  }
}

/**
 * Adds time slots for a specific time period
 */
function addTimeSlotsForPeriod(timeSlots, date, startTime, endTime, duration, events, specialCalendarId) {
  console.log(`Adding slots for period: ${startTime} - ${endTime}, duration: ${duration} minutes`);

  const [startHour, startMinute] = startTime.split(':').map(Number);
  const [endHour, endMinute] = endTime.split(':').map(Number);

  const periodStart = new Date(date);
  periodStart.setHours(startHour, startMinute, 0, 0);

  const periodEnd = new Date(date);
  periodEnd.setHours(endHour, endMinute, 0, 0);

  console.log(`Period from ${periodStart} to ${periodEnd}`);

  let currentTime = new Date(periodStart);
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const isToday = date.getTime() === today.getTime();

  console.log(`Is today: ${isToday}, Current time: ${now}`);
  
  let slotCount = 0;
  while (currentTime.getTime() + (duration * 60000) <= periodEnd.getTime()) {
    slotCount++;
    const slotEnd = new Date(currentTime.getTime() + (duration * 60000));
    
    console.log(`Checking slot ${slotCount}: ${formatTime(currentTime)} - ${formatTime(slotEnd)}`);
    
    // Skip slots that are in the past or less than 3 hours from now if this is today
    const threeHoursFromNow = new Date(now.getTime() + (3 * 60 * 60 * 1000)); // Add 3 hours to current time
    if (isToday && currentTime <= threeHoursFromNow) {
      console.log(`Skipping slot - too close to current time (less than 3 hours)`);
      currentTime = new Date(currentTime.getTime() + (CONFIG.TIME_SLOT_INTERVAL * 60000));
      continue;
    }
    
    // Check if this slot conflicts with existing events
    const hasConflict = events.some(event => {
      const eventStart = event.getStartTime();
      const eventEnd = event.getEndTime();
      const eventTitle = event.getTitle();
      const eventCalendarId = event._sourceCalendarId || 'unknown';

      // Handle all-day events - skip them unless they are "אין לקוחות" events
      const isAllDayEvent = event.isAllDayEvent();
      if (isAllDayEvent) {
        // Special case: "אין לקוחות" events should block time slots
        // Clean the title by removing extra spaces and check if it matches "אין לקוחות"
        const cleanTitle = eventTitle.trim().replace(/\s+/g, ' ');
        if (cleanTitle === 'אין לקוחות') {
          console.log(`Blocking due to "אין לקוחות" all-day event: "${eventTitle}" (cleaned: "${cleanTitle}")`);
          return true; // This blocks the time slot
        } else {
          console.log(`Skipping all-day event: "${eventTitle}" (cleaned: "${cleanTitle}")`);
          return false; // This allows the time slot
        }
      }

      // Check if this event is from the special calendar that requires buffer time
      if (eventCalendarId === specialCalendarId) {
        // For the special calendar: block 30 minutes before and 1 hour after
        const bufferBefore = 30 * 60000; // 30 minutes in milliseconds
        const bufferAfter = 60 * 60000; // 1 hour in milliseconds

        const eventStartWithBuffer = new Date(eventStart.getTime() - bufferBefore);
        const eventEndWithBuffer = new Date(eventEnd.getTime() + bufferAfter);

        const conflict = (currentTime < eventEndWithBuffer && slotEnd > eventStartWithBuffer);

        // Detailed logging for sports calendar conflicts
        logger.info('Checking sports calendar event conflict', {
          slotStart: currentTime.toISOString(),
          slotEnd: slotEnd.toISOString(),
          eventTitle: eventTitle,
          eventStart: eventStart.toISOString(),
          eventEnd: eventEnd.toISOString(),
          bufferBeforeMinutes: 30,
          bufferAfterMinutes: 60,
          eventStartWithBuffer: eventStartWithBuffer.toISOString(),
          eventEndWithBuffer: eventEndWithBuffer.toISOString(),
          hasConflict: conflict,
          conflictCheck: `${currentTime.toISOString()} < ${eventEndWithBuffer.toISOString()} && ${slotEnd.toISOString()} > ${eventStartWithBuffer.toISOString()}`
        });

        if (conflict) {
          console.log(`Conflict with special calendar event "${eventTitle}" (${eventStart} - ${eventEnd}) from calendar ${eventCalendarId} including buffers: ${formatTime(eventStartWithBuffer)} - ${formatTime(eventEndWithBuffer)}`);
        }
        return conflict;
      } else {
        // For regular calendars: only check direct overlap
        const conflict = (currentTime < eventEnd && slotEnd > eventStart);
        if (conflict) {
          console.log(`Conflict with regular calendar event "${eventTitle}" (${eventStart} - ${eventEnd}) from calendar ${eventCalendarId}`);
        }
        return conflict;
      }
    });
    
    if (!hasConflict) {
      console.log(`Slot available: ${formatTime(currentTime)} - ${formatTime(slotEnd)}`);

      // Log available slot details for debugging
      logger.info('Available slot found', {
        slotStart: currentTime.toISOString(),
        slotEnd: slotEnd.toISOString(),
        formattedStart: formatTime(currentTime),
        formattedEnd: formatTime(slotEnd),
        durationMinutes: duration,
        checkedEventsCount: events.length
      });

      timeSlots.push({
        start: formatTime(currentTime),
        end: formatTime(slotEnd),
        startDateTime: currentTime.toISOString(),
        endDateTime: slotEnd.toISOString()
      });
    } else {
      console.log(`Slot blocked by conflict`);

      // Log blocked slot details for debugging
      logger.info('Slot blocked by conflict', {
        slotStart: currentTime.toISOString(),
        slotEnd: slotEnd.toISOString(),
        formattedStart: formatTime(currentTime),
        formattedEnd: formatTime(slotEnd),
        durationMinutes: duration
      });
    }
    
    // Move to next slot
    currentTime = new Date(currentTime.getTime() + (CONFIG.TIME_SLOT_INTERVAL * 60000));
  }
  
  console.log(`Finished checking period ${startTime} - ${endTime}. Added ${timeSlots.filter(slot => slot.start >= startTime && slot.start <= endTime).length} slots.`);
}

// Gets availability data for a specific date (Business hours + Busy slots)
// This allows client-side calculation of time slots for immediate UI updates
// @param {string} dateStr - Date in YYYY-MM-DD format
// @return {Object} Availability data { businessHours, busySlots }
function getDailyAvailability(dateStr) {
  try {
    const date = new Date(dateStr);
    const dayOfWeek = getDayOfWeek(date);
    const businessHours = CONFIG.BUSINESS_HOURS[dayOfWeek];
    
    // If closed on this day
    if (!businessHours) {
      return {
        businessHours: null,
        busySlots: []
      };
    }

    const busySlots = getBusySlotsForDate(date);
    
    return {
      businessHours: businessHours,
      busySlots: busySlots
    };

  } catch (error) {
    console.error('Error getting daily availability:', error);
    logger.error('Failed to get daily availability', {
      error: error.message,
      requestedDate: dateStr
    });
    return {
      businessHours: null,
      busySlots: [],
      error: error.message
    };
  }
}

// Helper to get all busy slots for a date including rules and buffers
function getBusySlotsForDate(date) {
  const busySlots = [];
  const startOfDay = new Date(date);
  startOfDay.setHours(0, 0, 0, 0);
  const endOfDay = new Date(date);
  endOfDay.setHours(23, 59, 59, 999);

  // Buffer time for special calendar
  const specialCalendarId = CONFIG.SPECIAL_CALENDAR_ID;
  const BUFFER_BEFORE = 30; // Minutes
  const BUFFER_AFTER = 60;  // Minutes

  // Check all calendars
  CONFIG.CALENDAR_IDS.forEach(calendarId => {
    try {
      const calendar = CalendarApp.getCalendarById(calendarId);
      if (!calendar) return;

      // Fetch events for the whole day
      // For special calendar, we need to fetch a bit more to account for buffers
      let queryStart = startOfDay;
      let queryEnd = endOfDay;
      
      if (calendarId === specialCalendarId) {
        queryStart = new Date(startOfDay.getTime() - (BUFFER_AFTER * 60000));
        queryEnd = new Date(endOfDay.getTime() + (BUFFER_BEFORE * 60000));
      }

      const events = calendar.getEvents(queryStart, queryEnd);

      events.forEach(event => {
        const eventStart = event.getStartTime();
        const eventEnd = event.getEndTime();
        const eventTitle = event.getTitle();
        
        let busyStart = eventStart;
        let busyEnd = eventEnd;

        // Apply rules
        if (event.isAllDayEvent()) {
          const cleanTitle = eventTitle.trim().replace(/\s+/g, ' ');
          if (cleanTitle === 'אין לקוחות') {
             // Blocks the whole day (or the relevant part if we treated it as specific hours, but all-day usually blocks everything)
             // For simplicity in this logic, we return it as a busy slot covering the day
             // But usually all-day events are 00:00 to 00:00 next day
             busySlots.push({
               start: eventStart.toISOString(),
               end: eventEnd.toISOString(),
               title: 'אין לקוחות'
             });
          }
          return; // Skip other all-day events
        }

        // Apply buffers for special calendar
        if (calendarId === specialCalendarId) {
          busyStart = new Date(eventStart.getTime() - (BUFFER_BEFORE * 60000));
          busyEnd = new Date(eventEnd.getTime() + (BUFFER_AFTER * 60000));
        }

        busySlots.push({
          start: busyStart.toISOString(),
          end: busyEnd.toISOString()
        });
      });

    } catch (e) {
      console.warn(`Error fetching events for calendar ${calendarId}:`, e);
    }
  });

  return busySlots;
}

/**
 * Quickly checks if a specific time slot is still available
 * 
 * Purpose: This function performs a fast check to verify that a selected time slot
 * hasn't been taken by another booking between the time the client loaded available
 * slots and when they clicked "confirm booking". This prevents double bookings.
 * 
 * Method: Instead of checking all time slots for the day (which is slow), this function
 * only queries calendar events that overlap with the specific selected time range,
 * making it much faster (~0.3-0.5 seconds vs ~3-5 seconds).
 * 
 * @param {string} startDateTime - ISO string of slot start time
 * @param {string} endDateTime - ISO string of slot end time
 * @return {boolean} True if slot is still available, false if taken
 * @author Moyshi
 * @created 2026-01-09
 * @updated 2026-01-19 - Fixed buffer logic to match getAvailableTimeSlots (30m before, 60m after)
 */
function isSlotStillAvailable(startDateTime, endDateTime) {
  try {
    const slotStart = new Date(startDateTime);
    const slotEnd = new Date(endDateTime);
    
    // Buffer time for special calendar (sports activities need buffer)
    const specialCalendarId = CONFIG.SPECIAL_CALENDAR_ID;
    const BUFFER_BEFORE = 30; // Minutes
    const BUFFER_AFTER = 60;  // Minutes
    
    console.log(`Checking slot availability: ${slotStart.toISOString()} - ${slotEnd.toISOString()}`);
    
    // Check all calendars for conflicts
    for (const calendarId of CONFIG.CALENDAR_IDS) {
      try {
        const calendar = CalendarApp.getCalendarById(calendarId);
        if (!calendar) {
          console.log(`Calendar not found: ${calendarId}`);
          continue;
        }
        
        // Get events that might conflict with this slot
        // We need to check a wider range for special calendar due to buffer
        let checkStart = slotStart;
        let checkEnd = slotEnd;
        
        if (calendarId === specialCalendarId) {
          // For special calendar, we need to look back BUFFER_AFTER before our slot starts
          // because an event that ended 59 mins ago would still block this slot due to after-buffer
          checkStart = new Date(slotStart.getTime() - (BUFFER_AFTER * 60000));
          
          // And we look forward BUFFER_BEFORE after our slot ends
          // because an event starting in 29 mins would block this slot due to before-buffer
          checkEnd = new Date(slotEnd.getTime() + (BUFFER_BEFORE * 60000));
        }
        
        const events = calendar.getEvents(checkStart, checkEnd);
        
        for (const event of events) {
          const eventStart = event.getStartTime();
          const eventEnd = event.getEndTime();
          const eventTitle = event.getTitle();
          
          // Handle all-day events logic (matching addTimeSlotsForPeriod)
          if (event.isAllDayEvent()) {
             const cleanTitle = eventTitle.trim().replace(/\s+/g, ' ');
             // Only block if title is "אין לקוחות"
             if (cleanTitle === 'אין לקוחות') {
               console.log(`Blocking due to "אין לקוחות" all-day event: "${eventTitle}"`);
               return false;
             } else {
               console.log(`Skipping harmless all-day event: "${eventTitle}"`);
               continue;
             }
          }
          
          // Calculate effective event time (with buffer for special calendar)
          if (calendarId === specialCalendarId) {
             const eventStartWithBuffer = new Date(eventStart.getTime() - (BUFFER_BEFORE * 60000));
             const eventEndWithBuffer = new Date(eventEnd.getTime() + (BUFFER_AFTER * 60000));
             
             // Check overlapping: (StartA < EndB) && (EndA > StartB)
             if (slotStart < eventEndWithBuffer && slotEnd > eventStartWithBuffer) {
                console.log(`Conflict with special calendar event "${eventTitle}" (buffers: ${BUFFER_BEFORE}/${BUFFER_AFTER})`);
                logger.info('Slot availability check - conflict found (special)', {
                  calendarId: calendarId,
                  eventTitle: eventTitle,
                  eventStart: eventStart.toISOString(),
                  eventEnd: eventEnd.toISOString(),
                  eventStartWithBuffer: eventStartWithBuffer.toISOString(),
                  eventEndWithBuffer: eventEndWithBuffer.toISOString(),
                  requestedSlotStart: slotStart.toISOString(),
                  requestedSlotEnd: slotEnd.toISOString()
                });
                return false;
             }
          } else {
            // Regular calendar - direct overlap check
            if (slotStart < eventEnd && slotEnd > eventStart) {
              console.log(`Conflict with regular calendar event "${eventTitle}"`);
              logger.info('Slot availability check - conflict found', {
                calendarId: calendarId,
                eventTitle: eventTitle,
                eventStart: eventStart.toISOString(),
                eventEnd: eventEnd.toISOString(),
                requestedSlotStart: slotStart.toISOString(),
                requestedSlotEnd: slotEnd.toISOString()
              });
              return false;
            }
          }
        }
      } catch (calError) {
        console.error(`Error checking calendar ${calendarId}:`, calError);
        logger.error('Error checking calendar for slot availability', {
          calendarId: calendarId,
          error: calError.message
        });
      }
    }
    
    console.log('Slot is still available');
    return true; // No conflicts - slot is available
    
  } catch (error) {
    console.error('Error in isSlotStillAvailable:', error);
    logger.error('Failed to check slot availability', {
      error: error.message,
      startDateTime: startDateTime,
      endDateTime: endDateTime
    });
    return false; // On error, assume slot is taken (safer to prevent double booking)
  }
}

/**
 * Books an appointment
 * @param {Object} appointmentData - Appointment details including language preference
 * @return {Object} Result of booking operation
 */
function bookAppointment(appointmentData) {
  try {
    logger.info('Starting appointment booking process', appointmentData);
    
    const { name, phone, email, numberOfPeople, selectedSlot, date, language } = appointmentData;
    
    // Validate input
    if (!name || !phone || !email || !numberOfPeople || !selectedSlot || !date) {
      throw new Error('חסרים פרטים נדרשים');
    }
    
    // Default to Hebrew if no language specified
    const clientLanguage = language || 'he';
    
    // Calculate duration - SERVER SIDE CALCULATION
    // This is critical to prevent bugs where frontend sends wrong duration or slot
    const duration = CONFIG.BASE_DURATION + ((numberOfPeople - 1) * CONFIG.ADDITIONAL_PERSON_DURATION);
    
    // Parse the date properly ignoring timezone offsets
    // Use the raw string values sent from the frontend: date ("2026-03-09") and selectedSlot.start ("18:30")
    const [year, month, day] = date.split('-').map(Number);
    const [hours, minutes] = selectedSlot.start.split(':').map(Number);
    
    // Create new Date in server's timezone (Asia/Jerusalem as per appsscript.json)
    // Note: JavaScript Date months are 0-indexed (0=January, 11=December)
    const startTime = new Date(year, month - 1, day, hours, minutes);
    
    // Calculate end time based on duration (not what frontend sent)
    const endTime = new Date(startTime.getTime() + (duration * 60000));
    
    console.log(`Booking for ${numberOfPeople} people. Duration: ${duration}min. Time: ${formatTime(startTime)} - ${formatTime(endTime)}`);

    // Quick check if the selected slot is still available (prevents double booking)
    // This is a fast check that only looks at the specific time range, using the CORRECT calculated end time
    if (!isSlotStillAvailable(startTime.toISOString(), endTime.toISOString())) {
      logger.info('Booking prevented - slot was taken before confirmation', {
        clientName: name,
        requestedDate: date,
        requestedSlot: selectedSlot,
        calculatedDuration: duration,
        calculatedEndTime: endTime.toISOString()
      });
      return {
        success: false,
        slotTaken: true,
        message: 'השעה שבחרת נתפסה בינתיים. אנא בחרי שעה אחרת.'
      };
    }
    let event = null;
    
    // Create calendar event in primary calendar (always in Hebrew for business owner)
    try {
      const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_IDS[0]);
      
      const eventTitle = `${name} (${numberOfPeople === 1 ? 'משתתפת 1' : numberOfPeople + ' בנות'}) ב${formatTimeForEventTitle(startTime)}`;
      
      // Create WhatsApp link for business owner to contact client
      const cleanClientPhone = phone.replace(/[\s\-\(\)]/g, ''); // Remove spaces, dashes, parentheses
      let whatsappPhone = cleanClientPhone;
      // Convert Israeli format to international: 05X -> 9725X
      if (whatsappPhone.startsWith('05')) {
        whatsappPhone = '972' + whatsappPhone.substring(1);
      } else if (whatsappPhone.startsWith('+972')) {
        whatsappPhone = whatsappPhone.substring(1);
      }
      const clientWhatsAppLink = `https://wa.me/${whatsappPhone}`;
      
      const eventDescription = `
לקוחה: ${name}
טלפון: ${phone}
וואטסאפ: ${clientWhatsAppLink}
מייל: ${email}
מספר משתתפות: ${numberOfPeople}
משך הפגישה: ${duration} דקות
      `.trim();
      
      logger.info('Creating calendar event', {
        calendarId: CONFIG.CALENDAR_IDS[0],
        eventTitle: eventTitle,
        startTime: startTime.toISOString(),
        endTime: endTime.toISOString(),
        duration: duration
      });
      
      event = calendar.createEvent(
        eventTitle,
        startTime,
        endTime,
        {
          description: eventDescription,
          guests: '',
          sendInvites: false
        }
      );
      
      logger.success('Calendar event created successfully', {
        clientName: name,
        eventId: event.getId(),
        startTime: startTime.toLocaleString('he-IL'),
        duration: duration,
        calendarId: CONFIG.CALENDAR_IDS[0]
      });
      
    } catch (calendarError) {
      logger.error('Failed to create calendar event', {
        error: calendarError.message,
        errorType: calendarError.name,
        stack: calendarError.stack,
        calendarId: CONFIG.CALENDAR_IDS[0],
        clientName: name
      });
      throw calendarError; // Re-throw to be caught by main try-catch
    }

    // Generate a unique management token and persist it in Script Properties.
    // The token allows the client to cancel or reschedule without an account.
    const managementToken = generateManagementToken({
      eventId: event.getId(),
      calendarId: CONFIG.CALENDAR_IDS[0],
      clientEmail: email,
      clientName: name,
      clientPhone: phone,
      startTimeIso: startTime.toISOString(),
      endTimeIso: endTime.toISOString(),
      duration: duration,
      numberOfPeople: numberOfPeople,
      language: clientLanguage
    });

    // Build the management URL and update the business-owner calendar event description
    const managementUrl = `${CONFIG.MANAGE_BASE_URL}?action=manage&token=${managementToken}`;
    try {
      const updatedDescription = event.getDescription() +
        `\n\n---\nלשינוי מועד או ביטול הפגישה:\n${managementUrl}`;
      event.setDescription(updatedDescription);
    } catch (descError) {
      logger.warn('Could not update calendar event description with management link', {
        error: descError.message, eventId: event.getId()
      });
    }
    
    // Send confirmation email to client in their preferred language
    sendConfirmationEmail(appointmentData, duration, startTime, clientLanguage, managementToken);
    
    logger.success('Appointment booking completed successfully', {
      clientName: name,
      eventId: event.getId(),
      startTime: startTime.toLocaleString('he-IL'),
      duration: duration,
      language: clientLanguage
    });
    
    // Schedule contact management to run in background (after response is sent to client)
    // This prevents the contact processing from delaying the client's response
    try {
      scheduleContactManagement(appointmentData);
    } catch (scheduleError) {
      logger.error('Failed to schedule contact management but appointment was successful', {
        error: scheduleError.message,
        clientName: name,
        eventId: event.getId()
      });
    }
    
    return {
      success: true,
      message: 'הפגישה נקבעה בהצלחה!',
      eventId: event.getId(),
      duration: duration,
      calendarLink: createCalendarLink(appointmentData, startTime, endTime, clientLanguage, managementToken)
    };
    
  } catch (error) {
    console.error('Error booking appointment:', error);
    logger.error('Failed to book appointment', { 
      error: error.message,
      errorType: error.name,
      stack: error.stack,
      appointmentData: appointmentData 
    });
    
    return {
      success: false,
      message: 'שגיאה בקביעת הפגישה: ' + error.message
    };
  }
}

/**
 * Sends confirmation email to the client in their preferred language.
 * Includes an embedded HTML management link if a managementToken is provided.
 * 
 * @param {Object} appointmentData - Appointment details
 * @param {number} duration - Duration in minutes
 * @param {Date} startTime - Appointment start time
 * @param {string} language - 'he' or 'en'
 * @param {string|null} managementToken - UUID token for cancel/reschedule link (optional)
 */
function sendConfirmationEmail(appointmentData, duration, startTime, language = 'he', managementToken = null) {
  try {
    logger.info('Starting email sending process', { 
      clientName: appointmentData.name, 
      clientEmail: appointmentData.email,
      appointmentTime: startTime.toLocaleString('he-IL'),
      duration: duration,
      language: language
    });
    
    const { name, phone, email, numberOfPeople } = appointmentData;
    
    // Email content based on language
    let subject, body, businessOwnerBody;
    
    // Build WhatsApp link with pre-filled message
    const appointmentDayFormatted = language === 'en' ? formatDateEnglish(startTime) : formatDate(startTime);
    const appointmentTimeFormatted = formatTime(startTime);
    
    let whatsappMessage, whatsappLinkText;
    if (language === 'en') {
      whatsappMessage = `${CONFIG.WHATSAPP_MSG_EN} ${appointmentDayFormatted} at ${appointmentTimeFormatted}`;
      whatsappLinkText = '📞 Contact me on WhatsApp - Click here';
    } else {
      whatsappMessage = `${CONFIG.WHATSAPP_MSG_HE}${appointmentDayFormatted} בשעה ${appointmentTimeFormatted}`;
      whatsappLinkText = '📞 ליצירת קשר איתי בוואטסאפ - לחצי כאן';
    }
    const whatsappUrl = `https://wa.me/${CONFIG.BUSINESS_PHONE_INTL}?text=${encodeURIComponent(whatsappMessage)}`;

    // Build management link text (plain text for body, HTML for htmlBody)
    let manageLinkPlain = '';
    let manageLinkHtml = '';
    if (managementToken) {
      const manageUrl = `${CONFIG.MANAGE_BASE_URL}?action=manage&token=${managementToken}`;
      if (language === 'en') {
        manageLinkPlain = `\n🔧 To reschedule or cancel your appointment:\n${manageUrl}`;
        manageLinkHtml = `<br><br>🔧 <a href="${manageUrl}" style="color:#e91e8c;font-weight:bold;">To reschedule or cancel your appointment - click here</a>`;
      } else {
        manageLinkPlain = `\n🔧 לשינוי מועד או ביטול הפגישה:\n${manageUrl}`;
        manageLinkHtml = `<br><br>🔧 <a href="${manageUrl}" style="color:#e91e8c;font-weight:bold;">לשינוי מועד או ביטול הפגישה - לחצי כאן</a>`;
      }
    }
    
    if (language === 'en') {
      // English email for client
      subject = 'Appointment Confirmation for Fitting';
      
      body = `Hello ${name},

Your appointment has been successfully booked!

Appointment Details:
📅 Date: ${appointmentDayFormatted}
🕐 Time: ${appointmentTimeFormatted}
⏱️ Duration: ${duration} minutes
👥 Number of participants: ${numberOfPeople}
📞 Phone: ${phone}

📍 Appointment Address:
${CONFIG.BUSINESS_ADDRESS_EN}

🗺️ Navigation with Waze:
${CONFIG.WAZE_LINK}

${whatsappLinkText}

Please arrive on time for your appointment.${manageLinkPlain}

Best regards,
${CONFIG.BUSINESS_NAME_EN}
Phone: ${CONFIG.BUSINESS_PHONE}`;
      
    } else {
      // Hebrew email for client (default)
      subject = 'אישור פגישה';
      
      body = `שלום ${name},

הפגישה שלך נקבעה בהצלחה!

פרטי הפגישה:
📅 תאריך: ${appointmentDayFormatted}
🕐 שעה: ${appointmentTimeFormatted}
⏱️ משך הפגישה: ${duration} דקות
👥 מספר משתתפות: ${numberOfPeople}
📞 טלפון: ${phone}

📍 כתובת הפגישה:
${CONFIG.BUSINESS_ADDRESS_HE}

🗺️ לניווט ב-Waze:
${CONFIG.WAZE_LINK}

${whatsappLinkText}

אנא הגיעי בזמן לפגישה.${manageLinkPlain}

בברכה,
${CONFIG.BUSINESS_NAME}
טלפון: ${CONFIG.BUSINESS_PHONE}`;
    }
    
    // Send confirmation email to client
    try {
      // Create HTML body - replace the WhatsApp text with a clickable link
      let htmlContent = body.replace(
        whatsappLinkText,
        `<a href="${whatsappUrl}" style="color: #25D366; font-weight: bold; text-decoration: none;">${whatsappLinkText}</a>`
      );
      // Append the embedded HTML management link (replacing the plain-text version already in htmlContent)
      if (managementToken && manageLinkPlain) {
        htmlContent = htmlContent.replace(manageLinkPlain, manageLinkHtml);
      }
      
      // Generate ICS standard file for calendar clients (Gmail, Apple Mail, Outlook)
      const endTime = new Date(startTime.getTime() + (duration * 60000));
      const icsAttachment = createIcsAttachment(appointmentData, startTime, endTime, language, managementToken);
      
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body,
        name: CONFIG.BUSINESS_EMAIL_NAME,
        replyTo: CONFIG.BUSINESS_EMAIL_FROM,
        htmlBody: `
        <div dir="${language === 'he' ? 'rtl' : 'ltr'}" style="font-family: Arial, sans-serif; text-align: ${language === 'he' ? 'right' : 'left'}; direction: ${language === 'he' ? 'rtl' : 'ltr'};">
          ${htmlContent.replace(/\n/g, '<br>')}
        </div>
        `,
        attachments: [icsAttachment]
      });
      console.log(`Confirmation email sent to client: ${email} in ${language}`);
      logger.success('Client confirmation email sent successfully', { clientEmail: email, language: language });
    } catch (error) {
      console.error('Failed to send email to client:', error);
      logger.error('Failed to send client confirmation email', { 
        clientEmail: email, 
        language: language,
        error: error.message,
        errorType: error.name,
        stack: error.stack
      });
    }
    
    // Send notification email to business owner (always in Hebrew)
    try {
      const hebrewAppointmentDay = formatDate(startTime);
      const hebrewAppointmentTime = formatTime(startTime);

      let manageLinkOwner = '';
      if (managementToken) {
        const manageUrl = `${CONFIG.MANAGE_BASE_URL}?action=manage&token=${managementToken}`;
        manageLinkOwner = `\n\nלשינוי מועד או ביטול הפגישה:\n${manageUrl}`;
      }
      
      businessOwnerBody = `התקבלה פגישה חדשה:

לקוחה: ${name}
טלפון: ${phone}
מייל: ${email}
מספר משתתפות: ${numberOfPeople}
תאריך: ${hebrewAppointmentDay}
שעה: ${hebrewAppointmentTime}
משך הפגישה: ${duration} דקות
שפת הלקוחה: ${language === 'he' ? 'עברית' : 'אנגלית'}

פרטי יצירת קשר:
מייל: ${email}
טלפון: ${phone}${manageLinkOwner}`;

      MailApp.sendEmail({
        to: CONFIG.BUSINESS_EMAIL,
        subject: `פגישה חדשה נקבעה - ${name}`,
        body: businessOwnerBody,
        htmlBody: `
        <div dir="rtl" style="font-family: Arial, sans-serif; text-align: right; direction: rtl;">
          ${businessOwnerBody.replace(/\n/g, '<br>')}
        </div>
        `
      });
      console.log(`Notification email sent to business owner: ${CONFIG.BUSINESS_EMAIL}`);
      logger.success('Business owner notification email sent successfully', { 
        businessEmail: CONFIG.BUSINESS_EMAIL,
        clientLanguage: language
      });
    } catch (error) {
      console.error('Failed to send notification to business owner:', error);
      logger.error('Failed to send business owner notification email', { 
        businessEmail: CONFIG.BUSINESS_EMAIL, 
        error: error.message,
        errorType: error.name,
        stack: error.stack
      });
    }
    
  } catch (error) {
    console.error('Error sending confirmation email:', error);
    logger.error('General error in email sending process', { 
      error: error.message,
      errorType: error.name,
      stack: error.stack,
      language: language
    });
  }
}

/**
 * Creates an iCalendar (.ics) file attachment for the appointment.
 * 
 * Purpose: Generates a standard iCalendar format file (RFC 5545) attached to confirmation emails.
 * Method of operation: Formats dates to ISO 8601, builds raw iCalendar content with METHOD:REQUEST
 * so email clients show "Add to Calendar" buttons. Adds a URL field and management link in DESCRIPTION
 * so the client can reschedule/cancel directly from their calendar app.
 * 
 * @param {Object} appointmentData - The appointment information
 * @param {Date} startTime - Start time of the event
 * @param {Date} endTime - End time of the event
 * @param {string} language - Client language ('he' or 'en')
 * @param {string|null} managementToken - UUID token for cancel/reschedule link (optional)
 * @return {Blob} - A Blob object representing the .ics file attachment
 */
function createIcsAttachment(appointmentData, startTime, endTime, language = 'he', managementToken = null) {
  const { name, numberOfPeople, email } = appointmentData;
  
  const escapeIcsText = (str) => {
    return str.replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\n/g, '\\n');
  };

  // Build management link suffix for DESCRIPTION
  let manageSuffix = '';
  let manageUrlRaw = '';
  if (managementToken) {
    manageUrlRaw = `${CONFIG.MANAGE_BASE_URL}?action=manage&token=${managementToken}`;
    if (language === 'en') {
      manageSuffix = `\n\nTo reschedule or cancel:\n${manageUrlRaw}`;
    } else {
      manageSuffix = `\n\nלשינוי מועד או ביטול הפגישה:\n${manageUrlRaw}`;
    }
  }
  
  let eventTitle, eventLocation, eventDescription;
  
  if (language === 'en') {
    eventTitle = escapeIcsText(CONFIG.EVENT_TITLE_EN + ' - ' + CONFIG.BUSINESS_NAME_EN);
    eventLocation = escapeIcsText(CONFIG.BUSINESS_NAME_EN + ', ' + CONFIG.BUSINESS_ADDRESS_EN);
    const rawDesc = `${CONFIG.EVENT_TITLE_EN} for ${numberOfPeople} ${numberOfPeople === 1 ? 'participant' : 'participants'}

📍 Address: ${CONFIG.BUSINESS_ADDRESS_EN}

🗺️ Navigation with Waze:
${CONFIG.WAZE_LINK}

📞 Phone: ${CONFIG.BUSINESS_PHONE}${manageSuffix}`;
    eventDescription = escapeIcsText(rawDesc);
  } else {
    eventTitle = escapeIcsText(CONFIG.EVENT_TITLE_HE + ' - ' + CONFIG.BUSINESS_NAME);
    eventLocation = escapeIcsText(CONFIG.BUSINESS_NAME + ', ' + CONFIG.BUSINESS_ADDRESS_HE);
    const rawDesc = `${CONFIG.EVENT_TITLE_HE} עבור ${numberOfPeople === 1 ? 'משתתפת 1' : numberOfPeople + ' בנות'}

📍 כתובת: ${CONFIG.BUSINESS_ADDRESS_HE}

🗺️ לניווט ב-Waze:
${CONFIG.WAZE_LINK}

📞 טלפון: ${CONFIG.BUSINESS_PHONE}${manageSuffix}`;
    eventDescription = escapeIcsText(rawDesc);
  }
  
  // Format dates for ICS format YYYYMMDDTHHMMSS (Local Time)
  const formatForIcsLocal = (dateObj) => {
    const y = dateObj.getFullYear();
    const m = String(dateObj.getMonth() + 1).padStart(2, '0');
    const d = String(dateObj.getDate()).padStart(2, '0');
    const h = String(dateObj.getHours()).padStart(2, '0');
    const min = String(dateObj.getMinutes()).padStart(2, '0');
    const s = String(dateObj.getSeconds()).padStart(2, '0');
    return `${y}${m}${d}T${h}${min}${s}`;
  };

  // DTSTAMP requires UTC
  const formatForIcsUTC = (dateObj) => {
    const y = dateObj.getUTCFullYear();
    const m = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
    const d = String(dateObj.getUTCDate()).padStart(2, '0');
    const h = String(dateObj.getUTCHours()).padStart(2, '0');
    const min = String(dateObj.getUTCMinutes()).padStart(2, '0');
    const s = String(dateObj.getUTCSeconds()).padStart(2, '0');
    return `${y}${m}${d}T${h}${min}${s}Z`;
  };

  const uid = Utilities.getUuid() + '@' + CONFIG.ICS_DOMAIN;
  const nowStr = formatForIcsUTC(new Date());
  
  // Explicitly use local time for the appointment boundaries
  const startStr = formatForIcsLocal(startTime);
  const endStr = formatForIcsLocal(endTime);
  
  // Construct the iCalendar file content incorporating the explicit Israel time zone
  // IMPORTANT: RFC 5545 requires lines to be separated by CRLF (\r\n)
  let icsContent = 
    "BEGIN:VCALENDAR\r\n" +
    "VERSION:2.0\r\n" +
    `PRODID:-//${CONFIG.BUSINESS_NAME_EN}//Appointment Booking System//HE\r\n` +
    "CALSCALE:GREGORIAN\r\n" +
    "METHOD:REQUEST\r\n" +
    "BEGIN:VEVENT\r\n" +
    `UID:${uid}\r\n` +
    `DTSTAMP:${nowStr}\r\n` +
    `DTSTART;TZID=Asia/Jerusalem:${startStr}\r\n` +
    `DTEND;TZID=Asia/Jerusalem:${endStr}\r\n` +
    `SUMMARY:${eventTitle}\r\n` +
    `LOCATION:${eventLocation}\r\n` +
    `DESCRIPTION:${eventDescription}\r\n`;

  // Add standard URL field so calendar apps show a clickable link
  if (manageUrlRaw) {
    icsContent += `URL:${manageUrlRaw}\r\n`;
  }

  icsContent +=
    "STATUS:CONFIRMED\r\n" +
    `ORGANIZER;CN="${CONFIG.BUSINESS_NAME}":mailto:${CONFIG.BUSINESS_EMAIL_FROM}\r\n` +
    `ATTENDEE;RSVP=TRUE;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;CN="${escapeIcsText(name)}":mailto:${email}\r\n` +
    "END:VEVENT\r\n" +
    "END:VCALENDAR";

  // Use Utilities.newBlob with UTF-8 bytes to prevent encoding issues with Hebrew chars
  return Utilities.newBlob(Utilities.newBlob(icsContent).getBytes(), 'text/calendar', 'invite.ics');
}

/**
 * Creates a Google Calendar link for the client to add the event to their calendar.
 * Includes management link in the event details when a managementToken is provided.
 * 
 * @param {Object} appointmentData - Appointment details
 * @param {Date} startTime - Appointment start time
 * @param {Date} endTime - Appointment end time
 * @param {string} language - Client's preferred language
 * @param {string|null} managementToken - UUID token for cancel/reschedule link (optional)
 */
function createCalendarLink(appointmentData, startTime, endTime, language = 'he', managementToken = null) {
  const { name, numberOfPeople } = appointmentData;

  // Build management suffix for event details
  let manageSuffix = '';
  if (managementToken) {
    const manageUrl = `${CONFIG.MANAGE_BASE_URL}?action=manage&token=${managementToken}`;
    if (language === 'en') {
      manageSuffix = `\n\nTo reschedule or cancel:\n${manageUrl}`;
    } else {
      manageSuffix = `\n\nלשינוי מועד או ביטול הפגישה:\n${manageUrl}`;
    }
  }
  
  let eventTitle, eventLocation, eventDescription;
  
  if (language === 'en') {
    // English calendar event
    eventTitle = encodeURIComponent(CONFIG.EVENT_TITLE_EN + ' - ' + CONFIG.BUSINESS_NAME_EN);
    eventLocation = encodeURIComponent(CONFIG.BUSINESS_NAME_EN + ', ' + CONFIG.BUSINESS_ADDRESS_EN);
    
    eventDescription = encodeURIComponent(`${CONFIG.EVENT_TITLE_EN} for ${numberOfPeople} ${numberOfPeople === 1 ? 'participant' : 'participants'}

📍 Address: ${CONFIG.BUSINESS_ADDRESS_EN}

🗺️ Navigation with Waze:
${CONFIG.WAZE_LINK}

📞 Phone: ${CONFIG.BUSINESS_PHONE}${manageSuffix}`);
    
  } else {
    // Hebrew calendar event (default)
    eventTitle = encodeURIComponent(CONFIG.EVENT_TITLE_HE + ' - ' + CONFIG.BUSINESS_NAME);
    eventLocation = encodeURIComponent(CONFIG.BUSINESS_NAME + ', ' + CONFIG.BUSINESS_ADDRESS_HE);
    
    eventDescription = encodeURIComponent(`${CONFIG.EVENT_TITLE_HE} עבור ${numberOfPeople === 1 ? 'משתתפת 1' : numberOfPeople + ' בנות'}

📍 כתובת: ${CONFIG.BUSINESS_ADDRESS_HE}

🗺️ לניווט ב-Waze:
${CONFIG.WAZE_LINK}

📞 טלפון: ${CONFIG.BUSINESS_PHONE}${manageSuffix}`);
  }
  
  // Format dates for Google Calendar 
  // We use local format without 'Z' (YYYYMMDDTHHMMSS) and append ctz=Asia/Jerusalem
  // This explicitly forces any calendar app to treat the time as Israel time, regardless of the phone's local timezone setting.
  const formatForGoogle = (dateObj) => {
    const y = dateObj.getFullYear();
    const m = String(dateObj.getMonth() + 1).padStart(2, '0');
    const d = String(dateObj.getDate()).padStart(2, '0');
    const h = String(dateObj.getHours()).padStart(2, '0');
    const min = String(dateObj.getMinutes()).padStart(2, '0');
    return `${y}${m}${d}T${h}${min}00`;
  };
  
  const startFormatted = formatForGoogle(startTime);
  const endFormatted = formatForGoogle(endTime);
  
  const calendarUrl = `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${eventTitle}&dates=${startFormatted}/${endFormatted}&details=${eventDescription}&location=${eventLocation}&ctz=Asia/Jerusalem`;
  
  return calendarUrl;
}

/**
 * Formats date in English format
 */
function formatDateEnglish(date) {
  return date.toLocaleDateString('en-US', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });
}

/**
 * Gets the configuration for the frontend
 */
function getConfig() {
  return {
    businessName: CONFIG.BUSINESS_NAME,
    baseDuration: CONFIG.BASE_DURATION,
    additionalPersonDuration: CONFIG.ADDITIONAL_PERSON_DURATION,
    businessAddressHe: CONFIG.BUSINESS_ADDRESS_HE,
    businessAddressEn: CONFIG.BUSINESS_ADDRESS_EN
  };
}

// ===== APPOINTMENT MANAGEMENT: TOKEN, CANCEL, RESCHEDULE, OTP, CLEANUP =====

/**
 * Generates a unique management token, persists appointment data in Script Properties.
 * 
 * Purpose: Creates a secure, unguessable link that lets a client cancel or reschedule
 * their appointment without creating an account.
 * 
 * Method: Generates a UUID, stores all needed appointment data as JSON under the key
 * "CANCEL_TOKEN_<uuid>". The token expires naturally when the appointment start time passes
 * (cleaned up by cleanupExpiredTokens).
 * 
 * @param {Object} tokenData - Appointment details to persist
 * @return {string} The generated UUID token
 */
function generateManagementToken(tokenData) {
  const token = Utilities.getUuid();
  const key = 'CANCEL_TOKEN_' + token;
  const payload = Object.assign({}, tokenData, { createdAt: new Date().toISOString() });
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(payload));
  logger.info('Management token generated', { token: token, clientEmail: tokenData.clientEmail });
  return token;
}

/**
 * Retrieves appointment details for the management UI.
 * Also schedules async cleanup of expired tokens.
 * 
 * Purpose: Validates a management token from the URL and returns the stored appointment data
 * so the frontend can display the "Manage your appointment" screen.
 * 
 * Method: Looks up "CANCEL_TOKEN_<token>" in Script Properties, checks the appointment
 * has not yet started, then returns the data. Schedules background cleanup as a side-effect.
 * 
 * @param {string} token - UUID management token
 * @return {Object} { success, appointmentDetails } or { success: false, message }
 */
function getAppointmentByToken(token) {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('CANCEL_TOKEN_' + token);
    if (!raw) {
      return { success: false, message: 'הקישור אינו תקין או שפג תוקפו.' };
    }
    const data = JSON.parse(raw);

    // Check if appointment has already started
    const startTime = new Date(data.startTimeIso);
    if (new Date() >= startTime) {
      return { success: false, expired: true, message: 'לא ניתן לשנות פגישה שכבר התחילה.' };
    }

    // Schedule async cleanup (runs 5 seconds after this response is sent)
    scheduleTokenCleanup();

    return {
      success: true,
      appointmentDetails: {
        clientName: data.clientName,
        clientEmail: data.clientEmail,
        clientPhone: data.clientPhone,
        startTimeIso: data.startTimeIso,
        endTimeIso: data.endTimeIso,
        duration: data.duration,
        numberOfPeople: data.numberOfPeople,
        language: data.language || 'he'
      }
    };
  } catch (err) {
    logger.error('getAppointmentByToken error', { error: err.message, token: token });
    return { success: false, message: 'שגיאה פנימית. אנא נסי שוב.' };
  }
}

/**
 * Cancels an appointment identified by its management token.
 * 
 * Purpose: Allows a client to cancel their appointment via the unique link in their email.
 * 
 * Method: Validates the token, deletes the Google Calendar event, removes the token from
 * Script Properties, and sends cancellation confirmation emails to both the client and the
 * business owner.
 * 
 * @param {string} token - UUID management token
 * @return {Object} { success, message }
 */
function cancelAppointment(token) {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('CANCEL_TOKEN_' + token);
    if (!raw) {
      return { success: false, message: 'הקישור אינו תקין או שפג תוקפו.' };
    }
    const data = JSON.parse(raw);

    // Check if appointment has already started
    const startTime = new Date(data.startTimeIso);
    if (new Date() >= startTime) {
      return { success: false, expired: true, message: 'לא ניתן לבטל פגישה שכבר התחילה.' };
    }

    // Delete the Google Calendar event
    try {
      const calendar = CalendarApp.getCalendarById(data.calendarId || CONFIG.CALENDAR_IDS[0]);
      if (calendar) {
        // eventId from CalendarApp ends with @google.com; getEventById needs just the ID part
        const eventIdClean = data.eventId.replace(/@.*$/, '');
        const event = calendar.getEventById(data.eventId) || calendar.getEventById(eventIdClean);
        if (event) {
          event.deleteEvent();
          logger.success('Calendar event deleted on cancellation', { eventId: data.eventId });
        } else {
          logger.warn('Calendar event not found during cancellation (may already be deleted)', { eventId: data.eventId });
        }
      }
    } catch (calErr) {
      logger.error('Failed to delete calendar event on cancellation', { error: calErr.message, eventId: data.eventId });
      // Don't abort – token is still cleaned up and emails sent
    }

    // Remove token
    props.deleteProperty('CANCEL_TOKEN_' + token);

    // Send cancellation emails
    sendCancellationEmails(data);

    return { success: true, message: 'הפגישה בוטלה בהצלחה.' };
  } catch (err) {
    logger.error('cancelAppointment error', { error: err.message, token: token });
    return { success: false, message: 'שגיאה פנימית. אנא נסי שוב.' };
  }
}

/**
 * Reschedules an appointment to a new date/time, identified by its management token.
 * Optionally updates the number of people.
 * 
 * Purpose: Allows a client to move their appointment to a new available slot via their management link.
 * 
 * Method: Validates the token and new slot (must be ≥3 hours from now), deletes the old calendar
 * event, creates a new one with a fresh token, removes the old token, and sends reschedule-specific
 * confirmation emails showing both old and new details.
 * 
 * @param {string} token - UUID management token of the existing appointment
 * @param {string} newDate - New date in YYYY-MM-DD format
 * @param {Object} newSlot - { start: 'HH:MM', end: 'HH:MM' }
 * @param {number|null} newNumberOfPeople - Updated people count (null keeps original)
 * @return {Object} { success, message, calendarLink?, duration? }
 */
function rescheduleAppointment(token, newDate, newSlot, newNumberOfPeople) {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('CANCEL_TOKEN_' + token);
    if (!raw) {
      return { success: false, message: 'הקישור אינו תקין או שפג תוקפו.' };
    }
    const data = JSON.parse(raw);

    // Check if original appointment has already started
    const oldStart = new Date(data.startTimeIso);
    if (new Date() >= oldStart) {
      return { success: false, expired: true, message: 'לא ניתן לשנות פגישה שכבר התחילה.' };
    }

    // Use new people count if provided, otherwise keep original
    const effectivePeople = (newNumberOfPeople && newNumberOfPeople > 0)
      ? newNumberOfPeople
      : data.numberOfPeople;

    // Recalculate duration based on (possibly new) people count
    const duration = CONFIG.BASE_DURATION + ((effectivePeople - 1) * CONFIG.ADDITIONAL_PERSON_DURATION);

    // Parse and validate the new slot
    const [ny, nm, nd] = newDate.split('-').map(Number);
    const [nh, nmin] = newSlot.start.split(':').map(Number);
    const newStartTime = new Date(ny, nm - 1, nd, nh, nmin);
    const newEndTime = new Date(newStartTime.getTime() + duration * 60000);

    // New slot must be at least 3 hours from now (same rule as new bookings)
    const threeHoursFromNow = new Date(Date.now() + 3 * 60 * 60 * 1000);
    if (newStartTime <= threeHoursFromNow) {
      return { success: false, message: 'לא ניתן לקבוע פגישה בפחות מ-3 שעות מעכשיו.' };
    }

    // Verify the new slot is still available
    if (!isSlotStillAvailable(newStartTime.toISOString(), newEndTime.toISOString())) {
      return { success: false, slotTaken: true, message: 'השעה שבחרת נתפסה. אנא בחרי שעה אחרת.' };
    }

    const clientLanguage = data.language || 'he';
    const newAppointmentData = {
      name: data.clientName,
      phone: data.clientPhone,
      email: data.clientEmail,
      numberOfPeople: effectivePeople,
      selectedSlot: newSlot,
      date: newDate,
      language: clientLanguage
    };

    // Preserve old appointment details for the email
    const oldAppointmentInfo = {
      startTimeIso: data.startTimeIso,
      endTimeIso: data.endTimeIso,
      duration: data.duration,
      numberOfPeople: data.numberOfPeople
    };

    // Delete old calendar event
    try {
      const calendar = CalendarApp.getCalendarById(data.calendarId || CONFIG.CALENDAR_IDS[0]);
      if (calendar) {
        const event = calendar.getEventById(data.eventId) || calendar.getEventById(data.eventId.replace(/@.*$/, ''));
        if (event) event.deleteEvent();
      }
    } catch (delErr) {
      logger.warn('Could not delete old calendar event during reschedule', { error: delErr.message });
    }

    // Remove old token
    props.deleteProperty('CANCEL_TOKEN_' + token);

    // Create new calendar event
    const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_IDS[0]);
    const cleanPhone = data.clientPhone.replace(/[\s\-\(\)]/g, '');
    let wpPhone = cleanPhone;
    if (wpPhone.startsWith('05')) wpPhone = '972' + wpPhone.substring(1);
    else if (wpPhone.startsWith('+972')) wpPhone = wpPhone.substring(1);
    const wpLink = `https://wa.me/${wpPhone}`;
    const eventTitle = `${data.clientName} (${effectivePeople === 1 ? 'משתתפת 1' : effectivePeople + ' בנות'}) ב${formatTimeForEventTitle(newStartTime)}`;
    const eventDescription = `לקוחה: ${data.clientName}\nטלפון: ${data.clientPhone}\nוואטסאפ: ${wpLink}\nמייל: ${data.clientEmail}\nמספר משתתפות: ${effectivePeople}\nמשך הפגישה: ${duration} דקות`;

    const newEvent = calendar.createEvent(eventTitle, newStartTime, newEndTime, {
      description: eventDescription,
      guests: '',
      sendInvites: false
    });

    // Generate a new management token for the rescheduled appointment
    const newToken = generateManagementToken({
      eventId: newEvent.getId(),
      calendarId: CONFIG.CALENDAR_IDS[0],
      clientEmail: data.clientEmail,
      clientName: data.clientName,
      clientPhone: data.clientPhone,
      startTimeIso: newStartTime.toISOString(),
      endTimeIso: newEndTime.toISOString(),
      duration: duration,
      numberOfPeople: effectivePeople,
      language: clientLanguage
    });

    // Update event description with new management link
    try {
      const manageUrl = `${CONFIG.MANAGE_BASE_URL}?action=manage&token=${newToken}`;
      newEvent.setDescription(newEvent.getDescription() + `\n\n---\nלשינוי מועד או ביטול הפגישה:\n${manageUrl}`);
    } catch (e) { /* non-critical */ }

    // Send reschedule-specific confirmation email (shows old vs new details)
    sendRescheduleConfirmationEmail(newAppointmentData, duration, newStartTime, clientLanguage, newToken, oldAppointmentInfo);

    logger.success('Appointment rescheduled', {
      clientName: data.clientName,
      newStart: newStartTime.toISOString(),
      newEventId: newEvent.getId()
    });

    return {
      success: true,
      message: 'הפגישה שונתה בהצלחה!',
      duration: duration,
      calendarLink: createCalendarLink(newAppointmentData, newStartTime, newEndTime, clientLanguage, newToken)
    };
  } catch (err) {
    logger.error('rescheduleAppointment error', { error: err.message, token: token });
    return { success: false, message: 'שגיאה פנימית. אנא נסי שוב.' };
  }
}

/**
 * Sends cancellation confirmation emails to both the client and the business owner.
 * 
 * @param {Object} data - Stored token data with appointment and client details
 */
function sendCancellationEmails(data) {
  try {
    const lang = data.language || 'he';
    const startTime = new Date(data.startTimeIso);
    const appointmentDay = lang === 'en' ? formatDateEnglish(startTime) : formatDate(startTime);
    const appointmentTime = formatTime(startTime);

    // Email to client
    try {
      let clientSubject, clientBody;
      if (lang === 'en') {
        clientSubject = 'Appointment Cancellation Confirmation';
        clientBody = `Hello ${data.clientName},\n\nYour appointment has been successfully cancelled.\n\nCancelled appointment:\n📅 Date: ${appointmentDay}\n🕐 Time: ${appointmentTime}\n\nTo book a new appointment:\n${CONFIG.MANAGE_BASE_URL}\n\nBest regards,\n${CONFIG.BUSINESS_NAME_EN}\nPhone: ${CONFIG.BUSINESS_PHONE}`;
      } else {
        clientSubject = 'אישור ביטול פגישה';
        clientBody = `שלום ${data.clientName},\n\nהפגישה שלך בוטלה בהצלחה.\n\nפרטי הפגישה שבוטלה:\n📅 תאריך: ${appointmentDay}\n🕐 שעה: ${appointmentTime}\n\nלקביעת פגישה חדשה:\n${CONFIG.MANAGE_BASE_URL}\n\nבברכה,\n${CONFIG.BUSINESS_NAME}\nטלפון: ${CONFIG.BUSINESS_PHONE}`;
      }
      MailApp.sendEmail({
        to: data.clientEmail,
        subject: clientSubject,
        body: clientBody,
        name: CONFIG.BUSINESS_EMAIL_NAME,
        replyTo: CONFIG.BUSINESS_EMAIL_FROM,
        htmlBody: `<div dir="${lang === 'he' ? 'rtl' : 'ltr'}" style="font-family:Arial,sans-serif;direction:${lang === 'he' ? 'rtl' : 'ltr'};text-align:${lang === 'he' ? 'right' : 'left'};">${clientBody.replace(/\n/g, '<br>')}</div>`
      });
    } catch (e) {
      logger.error('Failed to send cancellation email to client', { error: e.message });
    }

    // Email to business owner (always Hebrew)
    try {
      const ownerBody = `פגישה בוטלה:\n\nלקוחה: ${data.clientName}\nטלפון: ${data.clientPhone}\nמייל: ${data.clientEmail}\nתאריך: ${formatDate(startTime)}\nשעה: ${formatTime(startTime)}\nמספר משתתפות: ${data.numberOfPeople}`;
      MailApp.sendEmail({
        to: CONFIG.BUSINESS_EMAIL,
        subject: `פגישה בוטלה - ${data.clientName}`,
        body: ownerBody,
        htmlBody: `<div dir="rtl" style="font-family:Arial,sans-serif;direction:rtl;text-align:right;">${ownerBody.replace(/\n/g, '<br>')}</div>`
      });
    } catch (e) {
      logger.error('Failed to send cancellation email to business owner', { error: e.message });
    }
  } catch (err) {
    logger.error('sendCancellationEmails error', { error: err.message });
  }
}

/**
 * Sends reschedule-specific confirmation emails to both client and business owner.
 * Shows the new appointment details prominently, then the old (cancelled) details in a
 * subdued style so the recipient can see what changed.
 * 
 * @param {Object} appointmentData - New appointment details (name, phone, email, numberOfPeople, etc.)
 * @param {number} duration - New appointment duration in minutes
 * @param {Date} newStartTime - New start time
 * @param {string} language - Client language ('he' or 'en')
 * @param {string} managementToken - New management token for the rescheduled appointment
 * @param {Object} oldInfo - { startTimeIso, endTimeIso, duration, numberOfPeople }
 */
function sendRescheduleConfirmationEmail(appointmentData, duration, newStartTime, language, managementToken, oldInfo) {
  try {
    const { name, phone, email, numberOfPeople } = appointmentData;
    const manageUrl = `${CONFIG.MANAGE_BASE_URL}?action=manage&token=${managementToken}`;

    const oldStartTime = new Date(oldInfo.startTimeIso);
    const oldEndTime = new Date(oldInfo.endTimeIso);

    // Date/time formatters
    const newDayFmt  = language === 'en' ? formatDateEnglish(newStartTime) : formatDate(newStartTime);
    const newTimeFmt = formatTime(newStartTime);
    const oldDayFmt  = language === 'en' ? formatDateEnglish(oldStartTime) : formatDate(oldStartTime);
    const oldTimeFmt = formatTime(oldStartTime);
    const oldEndFmt  = formatTime(oldEndTime);
    const newEndTime = new Date(newStartTime.getTime() + duration * 60000);
    const newEndFmt  = formatTime(newEndTime);

    const dir = language === 'he' ? 'rtl' : 'ltr';
    const align = language === 'he' ? 'right' : 'left';

    // --- CLIENT EMAIL ---
    try {
      let subject, plainBody, htmlBody;

      if (language === 'en') {
        subject = `Appointment Rescheduled – ${newDayFmt} at ${newTimeFmt}`;

        plainBody = `Hello ${name},\n\nYour appointment has been rescheduled successfully!\n\n` +
          `NEW APPOINTMENT DETAILS:\n📅 Date: ${newDayFmt}\n🕐 Time: ${newTimeFmt} - ${newEndFmt}\n⏱️ Duration: ${duration} minutes\n👥 Number of participants: ${numberOfPeople}\n\n` +
          `📍 Address: ${CONFIG.BUSINESS_ADDRESS_EN}\n🗺️ Waze: ${CONFIG.WAZE_LINK}\n\n` +
          `To reschedule or cancel:\n${manageUrl}\n\n` +
          `--- Previous appointment (cancelled) ---\n📅 Date: ${oldDayFmt}\n🕐 Time: ${oldTimeFmt} - ${oldEndFmt}\n👥 Girls: ${oldInfo.numberOfPeople}\n\n` +
          `Best regards,\n${CONFIG.BUSINESS_NAME_EN}\nPhone: ${CONFIG.BUSINESS_PHONE}`;

        htmlBody = `<div dir="${dir}" style="font-family:Arial,sans-serif;direction:${dir};text-align:${align};">
          <p>Hello ${name},</p>
          <p><strong>Your appointment has been rescheduled successfully!</strong></p>
          <div style="background:#f0faf0;border:2px solid #4caf50;border-radius:10px;padding:1rem 1.25rem;margin:1rem 0;">
            <h3 style="margin:0 0 0.5rem;color:#2e7d32;">✅ New Appointment Details</h3>
            <p style="margin:0.25rem 0;">📅 <strong>Date:</strong> ${newDayFmt}</p>
            <p style="margin:0.25rem 0;">🕐 <strong>Time:</strong> ${newTimeFmt} - ${newEndFmt}</p>
            <p style="margin:0.25rem 0;">⏱️ <strong>Duration:</strong> ${duration} minutes</p>
            <p style="margin:0.25rem 0;">👥 <strong>Number of participants:</strong> ${numberOfPeople}</p>
          </div>
          <p>📍 <strong>Address:</strong> ${CONFIG.BUSINESS_ADDRESS_EN}<br>
          🗺️ <a href="${CONFIG.WAZE_LINK}">Navigation with Waze</a></p>
          <p>🔧 <a href="${manageUrl}" style="color:#e91e8c;font-weight:bold;">To reschedule or cancel - click here</a></p>
          <div style="margin-top:1.5rem;padding:0.75rem 1rem;background:#f5f5f5;border-radius:8px;border-left:3px solid #bbb;">
            <p style="margin:0 0 0.25rem;color:#888;font-size:0.9em;"><strong>Previous appointment (cancelled):</strong></p>
            <p style="margin:0.15rem 0;color:#999;font-size:0.85em;">📅 ${oldDayFmt}</p>
            <p style="margin:0.15rem 0;color:#999;font-size:0.85em;">🕐 ${oldTimeFmt} - ${oldEndFmt}</p>
            <p style="margin:0.15rem 0;color:#999;font-size:0.85em;">👥 ${oldInfo.numberOfPeople} ${oldInfo.numberOfPeople === 1 ? 'participant' : 'participants'}</p>
          </div>
          <p style="margin-top:1rem;">Best regards,<br>${CONFIG.BUSINESS_NAME_EN}<br>Phone: ${CONFIG.BUSINESS_PHONE}</p>
        </div>`;
      } else {
        subject = `עדכון מועד פגישה – ${newDayFmt} בשעה ${newTimeFmt}`;

        plainBody = `שלום ${name},\n\nהפגישה שלך עודכנה בהצלחה!\n\n` +
          `פרטי הפגישה החדשים:\n📅 תאריך: ${newDayFmt}\n🕐 שעה: ${newTimeFmt} - ${newEndFmt}\n⏱️ משך: ${duration} דקות\n👥 מספר משתתפות: ${numberOfPeople}\n\n` +
          `📍 כתובת: ${CONFIG.BUSINESS_ADDRESS_HE}\n🗺️ Waze: ${CONFIG.WAZE_LINK}\n\n` +
          `לשינוי מועד או ביטול הפגישה:\n${manageUrl}\n\n` +
          `--- הפגישה הקודמת (בוטלה) ---\n📅 תאריך: ${oldDayFmt}\n🕐 שעה: ${oldTimeFmt} - ${oldEndFmt}\n👥 מספר משתתפות: ${oldInfo.numberOfPeople}\n\n` +
          `בברכה,\n${CONFIG.BUSINESS_NAME}\nטלפון: ${CONFIG.BUSINESS_PHONE}`;

        htmlBody = `<div dir="${dir}" style="font-family:Arial,sans-serif;direction:${dir};text-align:${align};">
          <p>שלום ${name},</p>
          <p><strong>הפגישה שלך עודכנה בהצלחה!</strong></p>
          <div style="background:#f0faf0;border:2px solid #4caf50;border-radius:10px;padding:1rem 1.25rem;margin:1rem 0;">
            <h3 style="margin:0 0 0.5rem;color:#2e7d32;">✅ פרטי הפגישה החדשים</h3>
            <p style="margin:0.25rem 0;">📅 <strong>תאריך:</strong> ${newDayFmt}</p>
            <p style="margin:0.25rem 0;">🕐 <strong>שעה:</strong> ${newTimeFmt} - ${newEndFmt}</p>
            <p style="margin:0.25rem 0;">⏱️ <strong>משך:</strong> ${duration} דקות</p>
            <p style="margin:0.25rem 0;">👥 <strong>מספר משתתפות:</strong> ${numberOfPeople}</p>
          </div>
          <p>📍 <strong>כתובת:</strong> ${CONFIG.BUSINESS_ADDRESS_HE}<br>
          🗺️ <a href="${CONFIG.WAZE_LINK}">לניווט ב-Waze</a></p>
          <p>🔧 <a href="${manageUrl}" style="color:#e91e8c;font-weight:bold;">לשינוי מועד או ביטול הפגישה - לחצי כאן</a></p>
          <div style="margin-top:1.5rem;padding:0.75rem 1rem;background:#f5f5f5;border-radius:8px;border-right:3px solid #bbb;">
            <p style="margin:0 0 0.25rem;color:#888;font-size:0.9em;"><strong>הפגישה הקודמת (בוטלה):</strong></p>
            <p style="margin:0.15rem 0;color:#999;font-size:0.85em;">📅 ${oldDayFmt}</p>
            <p style="margin:0.15rem 0;color:#999;font-size:0.85em;">🕐 ${oldTimeFmt} - ${oldEndFmt}</p>
            <p style="margin:0.15rem 0;color:#999;font-size:0.85em;">👥 ${oldInfo.numberOfPeople === 1 ? 'משתתפת 1' : oldInfo.numberOfPeople + ' בנות'}</p>
          </div>
          <p style="margin-top:1rem;">בברכה,<br>${CONFIG.BUSINESS_NAME}<br>טלפון: ${CONFIG.BUSINESS_PHONE}</p>
        </div>`;
      }

      // ICS attachment for the new appointment
      const icsAttachment = createIcsAttachment(appointmentData, newStartTime, newEndTime, language, managementToken);

      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: plainBody,
        name: CONFIG.BUSINESS_EMAIL_NAME,
        replyTo: CONFIG.BUSINESS_EMAIL_FROM,
        htmlBody: htmlBody,
        attachments: [icsAttachment]
      });
      logger.success('Reschedule confirmation email sent to client', { clientEmail: email, language: language });
    } catch (e) {
      logger.error('Failed to send reschedule email to client', { error: e.message });
    }

    // --- BUSINESS OWNER EMAIL (always Hebrew) ---
    try {
      const hebrewNewDay  = formatDate(newStartTime);
      const hebrewNewTime = formatTime(newStartTime);
      const hebrewOldDay  = formatDate(oldStartTime);
      const hebrewOldTime = formatTime(oldStartTime);

      const ownerBody = `פגישה עודכנה:

לקוחה: ${name}
טלפון: ${phone}
מייל: ${email}

פרטי הפגישה החדשים:
תאריך: ${hebrewNewDay}
שעה: ${hebrewNewTime}
מספר משתתפות: ${numberOfPeople}
משך הפגישה: ${duration} דקות
שפת הלקוחה: ${language === 'he' ? 'עברית' : 'אנגלית'}

לשינוי מועד או ביטול הפגישה:
${manageUrl}

--- הפגישה הקודמת (בוטלה) ---
תאריך: ${hebrewOldDay}
שעה: ${hebrewOldTime}
מספר משתתפות: ${oldInfo.numberOfPeople}
משך: ${oldInfo.duration} דקות`;

      MailApp.sendEmail({
        to: CONFIG.BUSINESS_EMAIL,
        subject: `פגישה עודכנה - ${name}`,
        body: ownerBody,
        htmlBody: `<div dir="rtl" style="font-family:Arial,sans-serif;direction:rtl;text-align:right;">${ownerBody.replace(/\n/g, '<br>')}</div>`
      });
      logger.success('Reschedule notification sent to business owner', { clientName: name });
    } catch (e) {
      logger.error('Failed to send reschedule notification to business owner', { error: e.message });
    }
  } catch (err) {
    logger.error('sendRescheduleConfirmationEmail error', { error: err.message });
  }
}

/**
 * Sends a one-time password (OTP) to a client's email address.
 * Stores the OTP in Script Properties for 10 minutes.
 * 
 * Purpose: Provides a fallback authentication path for clients who have lost their
 * management link (e.g. deleted the confirmation email).
 * 
 * Method: Generates a 6-digit code, stores it under "OTP_<normalizedEmail>" with a 10-minute
 * expiry timestamp, then emails it to the client.
 * 
 * @param {string} email - Client email address
 * @param {string} lang - Language for the email ('he' or 'en')
 * @return {Object} { success, message }
 */
function sendOTP(email, lang = 'he') {
  try {
    const normalizedEmail = email.toLowerCase().trim();
    const code = String(Math.floor(100000 + Math.random() * 900000)); // 6-digit code
    const expiresAt = new Date(Date.now() + 10 * 60 * 1000).toISOString(); // 10 minutes

    const otpKey = 'OTP_' + normalizedEmail;
    PropertiesService.getScriptProperties().setProperty(otpKey, JSON.stringify({
      code: code,
      expiresAt: expiresAt
    }));

    // Send OTP email
    let subject, body;
    if (lang === 'en') {
      subject = 'Your verification code';
      body = `Your one-time verification code is:\n\n${code}\n\nThis code is valid for 10 minutes.\n\n${CONFIG.BUSINESS_NAME_EN} – Appointment Booking`;
    } else {
      subject = 'קוד האימות שלך';
      body = `קוד האימות החד-פעמי שלך הוא:\n\n${code}\n\nקוד זה תקף ל-10 דקות.\n\n${CONFIG.BUSINESS_NAME} – קביעת פגישה`;
    }

    MailApp.sendEmail({
      to: normalizedEmail,
      subject: subject,
      body: body,
      name: CONFIG.BUSINESS_EMAIL_NAME,
      replyTo: CONFIG.BUSINESS_EMAIL_FROM,
      htmlBody: `<div dir="${lang === 'he' ? 'rtl' : 'ltr'}" style="font-family:Arial,sans-serif;direction:${lang === 'he' ? 'rtl' : 'ltr'};text-align:${lang === 'he' ? 'right' : 'left'};"><p style="font-size:1.1em;">${body.replace(/\n/g, '<br>').replace(code, `<strong style="font-size:1.8em;letter-spacing:4px;">${code}</strong>`)}</p></div>`
    });

    logger.info('OTP sent', { email: normalizedEmail });
    return { success: true, message: lang === 'he' ? 'קוד נשלח למייל שלך.' : 'Code sent to your email.' };
  } catch (err) {
    logger.error('sendOTP error', { error: err.message, email: email });
    return { success: false, message: lang === 'he' ? 'שגיאה בשליחת הקוד. אנא נסי שוב.' : 'Error sending code. Please try again.' };
  }
}

/**
 * Verifies an OTP code and returns the client's upcoming appointments.
 * 
 * Purpose: After a client enters their OTP, this function authenticates them and returns
 * all their upcoming appointments so they can choose which one to manage.
 * 
 * Method: Validates the OTP (6-digit match + not expired), then searches all CANCEL_TOKEN_*
 * properties for appointments matching the client's email that have not yet started.
 * Schedules async cleanup as a side-effect.
 * 
 * @param {string} email - Client email address
 * @param {string} code - 6-digit OTP entered by the client
 * @return {Object} { success, appointments: [...] } or { success: false, message }
 */
function verifyOTPAndGetAppointments(email, code) {
  try {
    const normalizedEmail = email.toLowerCase().trim();
    const otpKey = 'OTP_' + normalizedEmail;
    const props = PropertiesService.getScriptProperties();
    const otpRaw = props.getProperty(otpKey);

    if (!otpRaw) {
      return { success: false, message: 'לא נשלח קוד למייל זה. אנא בקשי קוד חדש.' };
    }

    const otpData = JSON.parse(otpRaw);
    if (new Date() > new Date(otpData.expiresAt)) {
      props.deleteProperty(otpKey);
      return { success: false, expired: true, message: 'הקוד פג תוקף. אנא בקשי קוד חדש.' };
    }
    if (String(otpData.code) !== String(code).trim()) {
      return { success: false, message: 'הקוד שגוי. אנא נסי שוב.' };
    }

    // Code is valid – delete it (single-use)
    props.deleteProperty(otpKey);

    // Find all upcoming appointments for this email
    const allProps = props.getProperties();
    const now = new Date();
    const appointments = [];

    for (const key in allProps) {
      if (!key.startsWith('CANCEL_TOKEN_')) continue;
      try {
        const appt = JSON.parse(allProps[key]);
        if (appt.clientEmail && appt.clientEmail.toLowerCase() === normalizedEmail) {
          const startTime = new Date(appt.startTimeIso);
          if (startTime > now) {
            const token = key.replace('CANCEL_TOKEN_', '');
            appointments.push({
              token: token,
              clientName: appt.clientName,
              clientEmail: appt.clientEmail,
              clientPhone: appt.clientPhone,
              startTimeIso: appt.startTimeIso,
              endTimeIso: appt.endTimeIso,
              duration: appt.duration,
              numberOfPeople: appt.numberOfPeople,
              language: appt.language || 'he'
            });
          }
        }
      } catch (e) { /* skip malformed */ }
    }

    // Sort by start time ascending
    appointments.sort((a, b) => new Date(a.startTimeIso) - new Date(b.startTimeIso));

    // Schedule async cleanup
    scheduleTokenCleanup();

    return { success: true, appointments: appointments };
  } catch (err) {
    logger.error('verifyOTPAndGetAppointments error', { error: err.message, email: email });
    return { success: false, message: 'שגיאה פנימית. אנא נסי שוב.' };
  }
}

/**
 * Deletes expired management tokens (appointments whose start time has passed).
 * Also removes its own triggers after running to prevent accumulation.
 * 
 * Purpose: Prevents Script Properties from growing indefinitely.
 * 
 * Method: Called asynchronously by a time-based trigger (scheduleTokenCleanup).
 * Iterates all CANCEL_TOKEN_* properties and deletes any whose startTimeIso < now.
 * Also deletes expired OTP_* entries.
 */
function cleanupExpiredTokens() {
  try {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    const now = new Date();
    let deleted = 0;

    for (const key in allProps) {
      if (key.startsWith('CANCEL_TOKEN_')) {
        try {
          const data = JSON.parse(allProps[key]);
          if (new Date(data.startTimeIso) <= now) {
            props.deleteProperty(key);
            deleted++;
          }
        } catch (e) {
          props.deleteProperty(key); // remove malformed entries
          deleted++;
        }
      } else if (key.startsWith('OTP_')) {
        try {
          const otpData = JSON.parse(allProps[key]);
          if (new Date(otpData.expiresAt) <= now) {
            props.deleteProperty(key);
            deleted++;
          }
        } catch (e) {
          props.deleteProperty(key);
          deleted++;
        }
      }
    }

    logger.info('cleanupExpiredTokens completed', { deletedCount: deleted });

    // Remove own triggers to prevent accumulation
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'cleanupExpiredTokens') {
        ScriptApp.deleteTrigger(trigger);
      }
    }
  } catch (err) {
    logger.error('cleanupExpiredTokens error', { error: err.message });
  }
}

/**
 * Schedules cleanupExpiredTokens to run asynchronously in 5 seconds.
 * Uses the same time-based trigger pattern as scheduleContactManagement.
 * This is called as a side-effect of getAppointmentByToken and verifyOTPAndGetAppointments
 * so cleanup happens opportunistically without delaying the user response.
 */
function scheduleTokenCleanup() {
  try {
    ScriptApp.newTrigger('cleanupExpiredTokens')
      .timeBased()
      .after(5 * 1000) // 5 seconds
      .create();
  } catch (err) {
    logger.warn('scheduleTokenCleanup: could not create trigger', { error: err.message });
  }
}

// ===== END APPOINTMENT MANAGEMENT =====

/**
 * Formats duration in a user-friendly way
 * @param {number} minutes - Duration in minutes
 * @return {string} Formatted duration string
 */
function formatDurationHebrew(minutes) {
  if (minutes < 60) {
    return `${minutes} דקות`;
  }
  
  const hours = Math.floor(minutes / 60);
  const remainingMinutes = minutes % 60;
  
  if (remainingMinutes === 0) {
    return hours === 1 ? 'שעה' : `${hours} שעות`;
  } else {
    const hoursText = hours === 1 ? 'שעה' : `${hours} שעות`;
    return `${hoursText} ו-${remainingMinutes} דקות`;
  }
}

// Utility functions

/**
 * Gets day of week in Hebrew format
 */
function getDayOfWeek(date) {
  const days = ['SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'];
  const dayIndex = date.getDay();
  console.log(`Date: ${date.toDateString()}, Day index: ${dayIndex}, Day name: ${days[dayIndex]}`);
  return days[dayIndex];
}

/**
 * Formats time to HH:MM format
 */
function formatTime(date) {
  return date.toLocaleTimeString('he-IL', { 
    hour: '2-digit', 
    minute: '2-digit',
    hour12: false 
  });
}

/**
 * Formats time for calendar event title in a user-friendly Hebrew format
 * Examples: 8:00 -> "8", 8:15 -> "8 ורבע", 8:30 -> "8 וחצי", 8:45 -> "רבע ל9"
 */
function formatTimeForEventTitle(date) {
  const hour24 = date.getHours();
  const minute = date.getMinutes();
  
  // Convert to 12-hour format
  const hour12 = hour24 === 0 ? 12 : (hour24 > 12 ? hour24 - 12 : hour24);
  
  // Format based on minutes
  if (minute === 0) {
    // Exact hour - just show the hour
    return `${hour12}`;
  } else if (minute === 15) {
    // Quarter past - "8 ורבע"
    return `${hour12} ורבע`;
  } else if (minute === 30) {
    // Half past - "8 וחצי"
    return `${hour12} וחצי`;
  } else if (minute === 45) {
    // Quarter to next hour - "רבע ל9"
    const nextHour = hour12 === 12 ? 1 : hour12 + 1;
    return `רבע ל${nextHour}`;
  } else {
    // For other minutes, just show hour (shouldn't happen with 15-minute intervals)
    return `${hour12}`;
  }
}

/**
 * Formats date to Hebrew format
 */
function formatDate(date) {
  return date.toLocaleDateString('he-IL', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });
}

/**
 * Schedules contact management to run in background after appointment booking
 * This ensures the client gets immediate response while contacts are processed separately
 * 
 * @param {Object} appointmentData - Client's appointment information
 * @created 2025-06-08
 */
function scheduleContactManagement(appointmentData) {
  try {
    logger.info('Scheduling background contact management', {
      clientName: appointmentData.name
    });
    
    // Store appointment data temporarily for the background process
    const tempData = {
      timestamp: new Date().getTime(),
      appointmentData: appointmentData
    };
    
    // Save to Script Properties with a unique key
    const dataKey = `CONTACT_PENDING_${tempData.timestamp}`;
    PropertiesService.getScriptProperties().setProperty(dataKey, JSON.stringify(tempData));
    
    // Create a time-based trigger to run contact management in 10 seconds
    // This gives enough time for the client response to complete
    ScriptApp.newTrigger('processScheduledContactManagement')
      .timeBased()
      .after(10 * 1000) // 10 seconds delay
      .create();
    
    logger.success('Contact management scheduled successfully', {
      clientName: appointmentData.name,
      dataKey: dataKey,
      scheduledFor: '10 seconds from now'
    });
    
  } catch (error) {
    logger.error('Failed to schedule contact management', {
      error: error.message,
      clientName: appointmentData.name
    });
  }
}

/**
 * Processes scheduled contact management (called by time-based trigger)
 * This function runs in background after the client has received their response
 * 
 * @created 2025-06-08
 */
function processScheduledContactManagement() {
  try {
    logger.info('Starting scheduled contact management processing');
    
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();
    
    // Find all pending contact management tasks
    const pendingTasks = [];
    for (const key in allProperties) {
      if (key.startsWith('CONTACT_PENDING_')) {
        try {
          const data = JSON.parse(allProperties[key]);
          pendingTasks.push({ key: key, data: data });
        } catch (parseError) {
          logger.warn('Failed to parse pending contact data', { key: key });
          properties.deleteProperty(key); // Clean up corrupted data
        }
      }
    }
    
    logger.info('Found pending contact management tasks', {
      taskCount: pendingTasks.length
    });
    
    // Process each pending task
    for (const task of pendingTasks) {
      try {
        logger.info('Processing contact management task', {
          clientName: task.data.appointmentData.name,
          taskKey: task.key
        });
        
        // Process the contact management
        handleContactAfterBooking(task.data.appointmentData);
        
        // Clean up the processed task
        properties.deleteProperty(task.key);
        
        logger.success('Contact management task completed', {
          clientName: task.data.appointmentData.name,
          taskKey: task.key
        });
        
      } catch (taskError) {
        logger.error('Failed to process contact management task', {
          error: taskError.message,
          taskKey: task.key,
          clientName: task.data.appointmentData?.name
        });
        
        // Clean up failed task to prevent accumulation
        properties.deleteProperty(task.key);
      }
    }
    
    // Clean up old triggers to prevent accumulation
    cleanupContactManagementTriggers();
    
    logger.success('Scheduled contact management processing completed', {
      processedTasks: pendingTasks.length
    });
    
  } catch (error) {
    logger.error('General error in scheduled contact management', {
      error: error.message,
      errorType: error.name,
      stack: error.stack
    });
  }
}

/**
 * Cleans up old contact management triggers to prevent accumulation
 * @created 2025-06-08
 */
function cleanupContactManagementTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'processScheduledContactManagement') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    }
    
    if (deletedCount > 0) {
      logger.info('Cleaned up old contact management triggers', {
        deletedCount: deletedCount
      });
    }
    
  } catch (error) {
    logger.warn('Failed to clean up old triggers', {
      error: error.message
    });
  }
}

/**
 * Handles contact management after a successful appointment booking
 */
function handleContactAfterBooking(appointmentData) {
  try {
    logger.info('👤 Starting contact management process', {
      clientName: appointmentData.name,
      clientPhone: appointmentData.phone,
      clientEmail: appointmentData.email,
      process: 'Contact Management'
    });

    const { name, phone, email } = appointmentData;
    
    // Clean phone number for comparison (remove spaces, dashes, parentheses)
    const cleanPhone = phone.replace(/[\s\-\(\)]/g, '');
    const cleanEmail = email.toLowerCase().trim();
    
    logger.info('🧹 Cleaned contact data for search', {
      originalPhone: phone,
      cleanPhone: cleanPhone,
      originalEmail: email,
      cleanEmail: cleanEmail,
      process: 'Data Cleaning'
    });

    // Search for existing contacts
    logger.info('🔍 Starting contact search process', {
      searchPhone: cleanPhone,
      searchEmail: cleanEmail,
      process: 'Contact Search'
    });
    
    const searchResult = searchInAllContacts(cleanPhone, cleanEmail);
    
    if (searchResult.foundBoth) {
      // Contact exists with both phone and email - no action needed
      const contactName = searchResult.contact.names ? 
        searchResult.contact.names[0].displayName : 'Unknown';
      
      logger.success('✅ Contact found with both phone and email - no action needed', {
        contactId: searchResult.contact.resourceName,
        contactName: contactName,
        process: 'Contact Search',
        status: 'Found Complete'
      });
      
    } else if (searchResult.foundByPhone || searchResult.foundByEmail) {
      // Found partial match - delete old contact and create new complete one
      try {
        const existingContact = searchResult.contact;
        const contactName = existingContact.names ? 
          existingContact.names[0].displayName : 'Unknown';
          
        logger.info('🔄 Found partial match - will recreate contact with complete info', {
          contactName: contactName,
          resourceName: existingContact.resourceName,
          foundByPhone: searchResult.foundByPhone,
          foundByEmail: searchResult.foundByEmail,
          process: 'Contact Recreation'
        });
        
        // Get full contact details before deleting
        const fullContactData = People.People.get(existingContact.resourceName, {
          personFields: 'names,emailAddresses,phoneNumbers,addresses,urls,organizations,birthdays,events,relations,nicknames,occupations,biographies,userDefined,clientData'
        });
        
        logger.info('📋 Retrieved full contact data', {
          resourceName: existingContact.resourceName,
          hasNames: !!fullContactData.names,
          hasEmails: !!fullContactData.emailAddresses,
          hasPhones: !!fullContactData.phoneNumbers,
          hasAddresses: !!fullContactData.addresses,
          hasUrls: !!fullContactData.urls,
          hasOrganizations: !!fullContactData.organizations,
          process: 'Data Retrieval'
        });
        
        // Delete the existing contact
        People.People.deleteContact(existingContact.resourceName);
        
        logger.info('🗑️ Successfully deleted existing contact', {
          resourceName: existingContact.resourceName,
          contactName: contactName,
          process: 'Contact Deletion'
        });
        
        // Prepare new contact data starting with existing data
        const newContactData = {};
        
        // Preserve existing names (don't change the name!) - clean metadata
        if (fullContactData.names && fullContactData.names.length > 0) {
          newContactData.names = fullContactData.names.map(name => ({
            givenName: name.givenName,
            familyName: name.familyName,
            displayName: name.displayName,
            middleName: name.middleName,
            honorificPrefix: name.honorificPrefix,
            honorificSuffix: name.honorificSuffix
          })).filter(name => name.givenName || name.familyName || name.displayName);
        } else {
          // Only if no name exists, use the form name
          newContactData.names = [{
            givenName: `${name} הגיעה לפגישה`
          }];
        }
        
        // Handle email addresses - clean metadata
        let emailAddresses = [];
        if (fullContactData.emailAddresses) {
          emailAddresses = fullContactData.emailAddresses.map(email => ({
            value: email.value,
            type: email.type || 'other',
            displayName: email.displayName
          })).filter(email => email.value);
        }
        
        // Add the new email if not found by email
        if (!searchResult.foundByEmail) {
          const emailExists = emailAddresses.some(email => 
            email.value.toLowerCase().trim() === cleanEmail
          );
          if (!emailExists) {
            emailAddresses.push({
              value: cleanEmail,
              type: 'home'
            });
          }
        }
        if (emailAddresses.length > 0) {
          newContactData.emailAddresses = emailAddresses;
        }
        
        // Handle phone numbers - clean metadata
        let phoneNumbers = [];
        if (fullContactData.phoneNumbers) {
          phoneNumbers = fullContactData.phoneNumbers.map(phone => ({
            value: phone.value,
            type: phone.type || 'other',
            canonicalForm: phone.canonicalForm
          })).filter(phone => phone.value);
        }
        
        // Add the new phone if not found by phone
        if (!searchResult.foundByPhone) {
          const normalizedPhones = normalizePhoneNumber(cleanPhone);
          const phoneExists = phoneNumbers.some(phone => {
            const existingNormalized = normalizePhoneNumber(phone.value);
            return normalizedPhones.some(newPhone => 
              existingNormalized.includes(newPhone)
            );
          });
          if (!phoneExists) {
            phoneNumbers.push({
              value: cleanPhone,
              type: 'mobile'
            });
          }
        }
        if (phoneNumbers.length > 0) {
          newContactData.phoneNumbers = phoneNumbers;
        }
        
        // Clean and preserve other existing data - remove metadata from each field
        if (fullContactData.addresses && fullContactData.addresses.length > 0) {
          newContactData.addresses = fullContactData.addresses.map(address => ({
            streetAddress: address.streetAddress,
            city: address.city,
            region: address.region,
            postalCode: address.postalCode,
            country: address.country,
            countryCode: address.countryCode,
            type: address.type || 'other'
          })).filter(address => address.streetAddress || address.city);
        }
        
        if (fullContactData.urls && fullContactData.urls.length > 0) {
          newContactData.urls = fullContactData.urls.map(url => ({
            value: url.value,
            type: url.type || 'other'
          })).filter(url => url.value);
        }
        
        if (fullContactData.organizations && fullContactData.organizations.length > 0) {
          newContactData.organizations = fullContactData.organizations.map(org => ({
            name: org.name,
            title: org.title,
            department: org.department,
            type: org.type || 'work'
          })).filter(org => org.name || org.title);
        }
        
        if (fullContactData.birthdays && fullContactData.birthdays.length > 0) {
          newContactData.birthdays = fullContactData.birthdays.map(birthday => ({
            date: birthday.date,
            text: birthday.text
          })).filter(birthday => birthday.date || birthday.text);
        }
        
        if (fullContactData.events && fullContactData.events.length > 0) {
          newContactData.events = fullContactData.events.map(event => ({
            date: event.date,
            type: event.type || 'other'
          })).filter(event => event.date);
        }
        
        if (fullContactData.relations && fullContactData.relations.length > 0) {
          newContactData.relations = fullContactData.relations.map(relation => ({
            person: relation.person,
            type: relation.type || 'other'
          })).filter(relation => relation.person);
        }
        
        if (fullContactData.nicknames && fullContactData.nicknames.length > 0) {
          newContactData.nicknames = fullContactData.nicknames.map(nickname => ({
            value: nickname.value,
            type: nickname.type || 'other'
          })).filter(nickname => nickname.value);
        }
        
        if (fullContactData.occupations && fullContactData.occupations.length > 0) {
          newContactData.occupations = fullContactData.occupations.map(occupation => ({
            value: occupation.value
          })).filter(occupation => occupation.value);
        }
        
        if (fullContactData.biographies && fullContactData.biographies.length > 0) {
          newContactData.biographies = fullContactData.biographies.map(bio => ({
            value: bio.value,
            contentType: bio.contentType || 'TEXT_PLAIN'
          })).filter(bio => bio.value);
        }
        
        // Create new contact with all the preserved and new data
        const createdPerson = People.People.createContact(newContactData);
        
        const preservedName = newContactData.names[0].displayName || 
                             newContactData.names[0].givenName || 
                             'Unknown';
        
        logger.success('✅ Successfully recreated contact with complete information', {
          oldResourceName: existingContact.resourceName,
          newResourceName: createdPerson.resourceName,
          preservedName: preservedName,
          totalPhones: phoneNumbers.length,
          totalEmails: emailAddresses.length,
          addedPhone: !searchResult.foundByPhone,
          addedEmail: !searchResult.foundByEmail,
          preservedAdditionalData: {
            addresses: !!fullContactData.addresses,
            urls: !!fullContactData.urls,
            organizations: !!fullContactData.organizations,
            other: !!(fullContactData.birthdays || fullContactData.events || fullContactData.relations)
          },
          process: 'Contact Recreation',
          status: 'Recreation Complete'
        });
        
      } catch (error) {
        logger.error('❌ Failed to recreate contact', {
          error: error.message,
          resourceName: searchResult.contact.resourceName,
          process: 'Contact Recreation',
          status: 'Recreation Failed'
        });
      }
      
    } else {
      // No contact found - create new one
      try {
        logger.info('➕ Creating new contact', {
          name: name,
          phone: cleanPhone,
          email: cleanEmail,
          process: 'Contact Creation'
        });
        
        const contactName = `${name} הגיעה לפגישה`;
        
        const person = {
          names: [
            {
              givenName: contactName
            }
          ],
          emailAddresses: [
            {
              value: cleanEmail,
              type: 'home'
            }
          ],
          phoneNumbers: [
            {
              value: cleanPhone,
              type: 'mobile'
            }
          ]
        };
        
        const createdPerson = People.People.createContact(person);
        
        logger.success('✅ Successfully created new contact', {
          resourceName: createdPerson.resourceName,
          contactName: contactName,
          phone: cleanPhone,
          email: cleanEmail,
          process: 'Contact Creation',
          status: 'Creation Complete'
        });
        
      } catch (error) {
        logger.error('❌ Failed to create new contact', {
          error: error.message,
          clientName: name,
          phone: cleanPhone,
          email: cleanEmail,
          process: 'Contact Creation',
          status: 'Creation Failed'
        });
      }
    }
    
  } catch (error) {
    logger.error('❌ General error in contact management process', {
      error: error.message,
      errorType: error.name,
      stack: error.stack,
      appointmentData: appointmentData,
      process: 'Contact Management',
      status: 'Process Failed'
    });
  }
}

/**
 * Normalizes phone number to different formats for better matching
 * @param {string} phone - Raw phone number
 * @return {Array} Array of normalized phone number formats
 */
function normalizePhoneNumber(phone) {
  // Remove all non-digits
  const digitsOnly = phone.replace(/[^\d]/g, '');
  
  const formats = [];
  
  // Add original digits
  formats.push(digitsOnly);
  
  // If starts with 05, also add +9725 version
  if (digitsOnly.startsWith('05')) {
    formats.push('+9725' + digitsOnly.substring(2));
    formats.push('9725' + digitsOnly.substring(2));
  }
  
  // If starts with +9725, also add 05 version
  if (digitsOnly.startsWith('9725')) {
    formats.push('05' + digitsOnly.substring(4));
  }
  
  // If starts with 9725, also add 05 version and +9725 version
  if (digitsOnly.startsWith('9725') && digitsOnly.length >= 8) {
    formats.push('05' + digitsOnly.substring(4));
    formats.push('+9725' + digitsOnly.substring(4));
  }
  
  return [...new Set(formats)]; // Remove duplicates
}

/**
 * Builds a more efficient contacts map with memory management
 * @return {Object} Object with phoneMap and emailMap
 */
function buildContactsMap() {
  try {
    logger.info('🔄 Starting contact map building process');
    
    const phoneMap = {};
    const emailMap = {};
    
    let pageToken = null;
    let totalContactsProcessed = 0;
    let batchNumber = 0;
    const batchSize = 500; // Reduced batch size
    let startTime = new Date();
    
    do {
      batchNumber++;
      const batchStartTime = new Date();
      
      // Stop after processing 5000 contacts to prevent memory issues
      if (totalContactsProcessed >= 5000) {
        logger.info('📊 Reached contact limit (5000) to prevent memory issues');
        break;
      }
      
      const connectionParams = {
        personFields: 'names,emailAddresses,phoneNumbers,metadata',
        pageSize: batchSize
      };
      
      if (pageToken) {
        connectionParams.pageToken = pageToken;
      }
      
      logger.info('📥 Loading contact batch', {
        batchNumber: batchNumber,
        batchSize: batchSize,
        totalProcessed: totalContactsProcessed,
        hasNextPage: !!pageToken,
        elapsedTime: `${Math.round((new Date() - startTime) / 1000)} seconds`
      });
      
      const connections = People.People.Connections.list('people/me', connectionParams);
      const contacts = connections && connections.connections ? connections.connections : [];
      
      if (contacts.length === 0) {
        logger.info('📭 No more contacts found, ending process');
        break;
      }
      
      totalContactsProcessed += contacts.length;
      
      // Process this batch of contacts
      let phonesAdded = 0;
      let emailsAdded = 0;
      
      for (let i = 0; i < contacts.length; i++) {
        const person = contacts[i];
        const resourceName = person.resourceName;
        
        // Index by phone numbers
        if (person.phoneNumbers) {
          for (let j = 0; j < person.phoneNumbers.length; j++) {
            const phoneValue = person.phoneNumbers[j].value;
            const normalizedPhones = normalizePhoneNumber(phoneValue);
            
            // Only keep the first occurrence of each phone number
            for (const normalizedPhone of normalizedPhones) {
              if (!phoneMap[normalizedPhone]) {
                phoneMap[normalizedPhone] = {
                  resourceName: resourceName,
                  person: person
                };
                phonesAdded++;
              }
            }
          }
        }
        
        // Index by email addresses
        if (person.emailAddresses) {
          for (let j = 0; j < person.emailAddresses.length; j++) {
            const emailValue = person.emailAddresses[j].value.toLowerCase().trim();
            if (!emailMap[emailValue]) {
              emailMap[emailValue] = {
                resourceName: resourceName,
                person: person
              };
              emailsAdded++;
            }
          }
        }
        
        // Clear references to help garbage collection
        if (i % 100 === 0) {
          contacts[i] = null;
        }
      }
      
      const batchEndTime = new Date();
      const batchDuration = (batchEndTime - batchStartTime) / 1000;
      
      logger.info('✅ Batch processing completed', {
        batchNumber: batchNumber,
        contactsInBatch: contacts.length,
        phonesAdded: phonesAdded,
        emailsAdded: emailsAdded,
        batchDuration: `${batchDuration.toFixed(2)} seconds`,
        totalContacts: totalContactsProcessed,
        totalPhones: Object.keys(phoneMap).length,
        totalEmails: Object.keys(emailMap).length
      });
      
      pageToken = connections.nextPageToken;
      
      // Force garbage collection between batches
      for (let i = 0; i < contacts.length; i++) {
        contacts[i] = null;
      }
      
    } while (pageToken);
    
    const totalDuration = (new Date() - startTime) / 1000;
    logger.success('🎉 Contact map building completed', {
      totalBatches: batchNumber,
      totalContacts: totalContactsProcessed,
      uniquePhones: Object.keys(phoneMap).length,
      uniqueEmails: Object.keys(emailMap).length,
      totalDuration: `${totalDuration.toFixed(2)} seconds`
    });
    
    return {
      phoneMap: phoneMap,
      emailMap: emailMap
    };
    
  } catch (error) {
    logger.error('❌ Error building contacts map', {
      error: error.message,
      errorType: error.name,
      stack: error.stack
    });
    
    return {
      phoneMap: {},
      emailMap: {}
    };
  }
}

/**
 * Searches for a contact using phone or email
 * @param {string} cleanPhone - Cleaned phone number
 * @param {string} cleanEmail - Cleaned email address
 * @return {Object} Search result with found contact information
 */
function searchInAllContacts(cleanPhone, cleanEmail) {
  try {
    logger.info('Starting contact search', {
      searchPhone: cleanPhone,
      searchEmail: cleanEmail
    });
    
    let pageToken = null;
    let foundContact = null;
    let foundByPhone = false;
    let foundByEmail = false;
    let totalProcessed = 0;
    
    do {
      const connectionParams = {
        personFields: 'names,emailAddresses,phoneNumbers,metadata',
        pageSize: 1000
      };
      
      if (pageToken) {
        connectionParams.pageToken = pageToken;
      }
      
      const connections = People.People.Connections.list('people/me', connectionParams);
      const contacts = connections && connections.connections ? connections.connections : [];
      
      if (contacts.length === 0) {
        break;
      }
      
      totalProcessed += contacts.length;
      logger.info('Processing batch of contacts', { 
        batchSize: contacts.length,
        totalProcessed: totalProcessed
      });
      
      // Search in this batch
      for (const person of contacts) {
        // Check phone numbers
        if (!foundByPhone && person.phoneNumbers) {
          for (const phone of person.phoneNumbers) {
            const normalizedPhones = normalizePhoneNumber(phone.value);
            if (normalizedPhones.includes(cleanPhone)) {
              foundByPhone = true;
              foundContact = person;
              break;
            }
          }
        }
        
        // Check email addresses
        if (!foundByEmail && person.emailAddresses) {
          for (const email of person.emailAddresses) {
            if (email.value.toLowerCase().trim() === cleanEmail) {
              foundByEmail = true;
              foundContact = foundContact || person;
              break;
            }
          }
        }
        
        // If found both, no need to continue
        if (foundByPhone && foundByEmail) {
          break;
        }
      }
      
      // If found both, no need to continue to next batch
      if (foundByPhone && foundByEmail) {
        break;
      }
      
      pageToken = connections.nextPageToken;
      
    } while (pageToken);
    
    logger.success('Contact search completed', {
      totalProcessed: totalProcessed,
      foundByPhone: foundByPhone,
      foundByEmail: foundByEmail,
      foundBoth: foundByPhone && foundByEmail
    });
    
    return {
      foundByPhone: foundByPhone,
      foundByEmail: foundByEmail,
      foundBoth: foundByPhone && foundByEmail,
      contact: foundContact
    };
    
  } catch (error) {
    logger.error('Error in contact search', {
      error: error.message,
      errorType: error.name,
      stack: error.stack
    });
    
    return {
      foundByPhone: false,
      foundByEmail: false,
      foundBoth: false,
      contact: null
    };
  }
}

/**
 * Test function to test the contact management functionality
 * @created 2025-06-08
 */
function testContactManagement() {
  console.log('🧪 Testing contact management system...');
  
  // Test data
  const testAppointmentData = {
    name: 'שרה כהן',
    phone: '050-1234567',
    email: 'sarah.test@example.com',
    numberOfPeople: 2
  };
  
  console.log('Test appointment data:', testAppointmentData);
  
  try {
    // Test the contact management function
    handleContactAfterBooking(testAppointmentData);
    console.log('✅ Contact management test completed successfully!');
  } catch (error) {
    console.error('❌ Contact management test failed:', error);
  }
}

/**
 * Test function to test the background scheduling system
 * @created 2025-06-08
 */
function testBackgroundContactScheduling() {
  console.log('🧪 Testing background contact scheduling...');
  
  // Test data
  const testAppointmentData = {
    name: 'רחל לוי',
    phone: '052-9876543',
    email: 'rachel.test@example.com',
    numberOfPeople: 1
  };
  
  console.log('Test appointment data:', testAppointmentData);
  
  try {
    // Test the scheduling function
    scheduleContactManagement(testAppointmentData);
    console.log('✅ Background contact scheduling test completed successfully!');
    console.log('📋 Contact management will run automatically in 10 seconds');
  } catch (error) {
    console.error('❌ Background contact scheduling test failed:', error);
  }
}

/**
 * Shows current pending contact management tasks and active triggers
 * @created 2025-06-08
 */
function showContactManagementStatus() {
  console.log('📊 Contact Management Status Report');
  console.log('=' .repeat(50));
  
  try {
    // Check pending tasks
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();
    
    const pendingTasks = [];
    for (const key in allProperties) {
      if (key.startsWith('CONTACT_PENDING_')) {
        try {
          const data = JSON.parse(allProperties[key]);
          pendingTasks.push({ key: key, data: data });
        } catch (parseError) {
          console.log(`⚠️ Corrupted task found: ${key}`);
        }
      }
    }
    
    console.log(`📋 Pending contact management tasks: ${pendingTasks.length}`);
    
    if (pendingTasks.length > 0) {
      console.log('\nPending tasks:');
      pendingTasks.forEach((task, index) => {
        const timestamp = new Date(task.data.timestamp);
        console.log(`  ${index + 1}. ${task.data.appointmentData.name} - ${timestamp.toLocaleString('he-IL')}`);
      });
    }
    
    // Check active triggers
    const triggers = ScriptApp.getProjectTriggers();
    const contactTriggers = triggers.filter(t => t.getHandlerFunction() === 'processScheduledContactManagement');
    
    console.log(`\n⚡ Active contact management triggers: ${contactTriggers.length}`);
    
    if (contactTriggers.length > 0) {
      console.log('Active triggers:');
      contactTriggers.forEach((trigger, index) => {
        console.log(`  ${index + 1}. Trigger ID: ${trigger.getUniqueId()}`);
      });
    }
    
    console.log('\n' + '=' .repeat(50));
    
  } catch (error) {
    console.error('❌ Error checking contact management status:', error);
  }
}

/**
 * Manually cleans up all pending contact tasks and triggers (use carefully!)
 * @created 2025-06-08
 */
function cleanupAllContactManagement() {
  console.log('🧹 Cleaning up all contact management tasks and triggers...');
  
  try {
    // Clean up pending tasks
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();
    
    let deletedTasks = 0;
    for (const key in allProperties) {
      if (key.startsWith('CONTACT_PENDING_')) {
        properties.deleteProperty(key);
        deletedTasks++;
      }
    }
    
    // Clean up triggers
    cleanupContactManagementTriggers();
    
    console.log(`✅ Cleanup completed:`);
    console.log(`   - Deleted ${deletedTasks} pending tasks`);
    console.log(`   - Cleaned up all contact management triggers`);
    
  } catch (error) {
    console.error('❌ Error during cleanup:', error);
  }
}

/**
 * Shows total number of contacts and first few contacts for verification using People API
 * @created 2025-06-08
 */
function showContactsInfo() {
  console.log('📇 Checking Google Contacts information using People API...');
  
  try {
    let totalContacts = 0;
    let batchCount = 0;
    let pageToken = null;
    const batchSize = 1000;
    
    console.log('\n🔍 Counting total contacts...');
    
    // First pass - count all contacts
    do {
      batchCount++;
      
      const connectionParams = {
        personFields: 'names,emailAddresses,phoneNumbers,metadata',
        pageSize: batchSize
      };
      
      if (pageToken) {
        connectionParams.pageToken = pageToken;
      }
      
      const connections = People.People.Connections.list('people/me', connectionParams);
      const contacts = connections && connections.connections ? connections.connections : [];
      
      totalContacts += contacts.length;
      console.log(`   Batch ${batchCount}: ${contacts.length} contacts (Total so far: ${totalContacts})`);
      
      // Show first few contacts from first batch
      if (batchCount === 1 && contacts.length > 0) {
        console.log('\n📋 First 5 contacts:');
        for (let i = 0; i < Math.min(5, contacts.length); i++) {
          const person = contacts[i];
          const name = person.names && person.names.length > 0 ? 
            person.names[0].displayName : 'No name';
          
          const phones = person.phoneNumbers ? 
            person.phoneNumbers.map(p => p.value).join(', ') : 'None';
          const emails = person.emailAddresses ? 
            person.emailAddresses.map(e => e.value).join(', ') : 'None';
          
          console.log(`${i + 1}. ${name}`);
          console.log(`   📞 Phones: ${phones}`);
          console.log(`   📧 Emails: ${emails}`);
        }
      }
      
      pageToken = connections.nextPageToken;
      
      // Limit to prevent excessive API calls during testing
      if (batchCount >= 10) {
        console.log('⚠️ Stopped at 10 batches to prevent excessive API usage');
        break;
      }
      
    } while (pageToken);
    
    if (pageToken) {
      console.log(`📊 Estimated total contacts: ${totalContacts}+ (stopped at batch ${batchCount})`);
    } else {
      console.log(`📊 Total contacts: ${totalContacts}`);
    }
    
    console.log(`\n✅ Contact information check completed successfully!`);
    
  } catch (error) {
    console.error('❌ Error checking contacts:', error);
    console.log('💡 Make sure People API is enabled in Advanced Google Services');
  }
}


/**
 * Test function to verify the system works
 */
function testSystem() {
  console.log('Testing appointment booking system...');
  
  // Test getting available slots for today
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  const dateStr = `${year}-${month}-${day}`;
  
  const slots = getAvailableTimeSlots(dateStr, 45); // 45 minutes for 2 people
  
  console.log(`Available slots for today (${dateStr}):`, slots);
  
  // Test configuration
  const config = getConfig();
  console.log('System configuration:', config);
  
  console.log('=== TEST SYSTEM COMPLETED ===');
}

/**
 * Test function to verify contacts map building after etag fix
 * @created 2025-06-08
 */
function testContactsMapBuilding() {
  console.log('🧪 Testing contacts map building after etag fix...');
  
  try {
    // Test building the contacts map
    const contactsMap = buildContactsMap();
    
    console.log('📊 Contacts map results:');
    console.log(`  - Phone map entries: ${Object.keys(contactsMap.phoneMap).length}`);
    console.log(`  - Email map entries: ${Object.keys(contactsMap.emailMap).length}`);
    
    // Show sample phone formats
    const samplePhones = Object.keys(contactsMap.phoneMap).slice(0, 10);
    console.log('\n📞 Sample phone formats in map:', samplePhones);
    
    // Show sample emails
    const sampleEmails = Object.keys(contactsMap.emailMap).slice(0, 10);
    console.log('\n📧 Sample emails in map:', sampleEmails);
    
    // Test search functionality
    if (samplePhones.length > 0) {
      console.log('\n🔍 Testing search with first phone...');
      const testPhone = samplePhones[0];
      const searchResult = searchInAllContacts(testPhone, 'test@example.com');
      console.log('Search result for test phone:', searchResult);
    }
    
    console.log('\n✅ Contacts map building test completed!');
    
  } catch (error) {
    console.error('❌ Contacts map building test failed:', error);
  }
}

/**
 * Test function to verify sports calendar access and events
 * @created 2025-12-30
 */
function testSportsCalendarAccess() {
  console.log('🏆 Testing sports calendar access...');

  try {
    const specialCalendarId = CONFIG.SPECIAL_CALENDAR_ID;

    console.log(`Attempting to access sports calendar: ${specialCalendarId}`);

    const calendar = CalendarApp.getCalendarById(specialCalendarId);

    if (!calendar) {
      console.log('❌ Could not access sports calendar - calendar is null');
      return;
    }

    console.log(`✅ Successfully accessed sports calendar: "${calendar.getName()}"`);

    // Test getting events for today
    const today = new Date();
    const startOfDay = new Date(today);
    startOfDay.setHours(0, 0, 0, 0);
    const endOfDay = new Date(today);
    endOfDay.setHours(23, 59, 59, 999);

    console.log(`Checking events from ${startOfDay} to ${endOfDay}`);

    const events = calendar.getEvents(startOfDay, endOfDay);
    console.log(`Found ${events.length} events in sports calendar today`);

    events.forEach((event, index) => {
      const startTime = event.getStartTime();
      const endTime = event.getEndTime();
      const title = event.getTitle();
      const isAllDay = event.isAllDayEvent();

      console.log(`Event ${index + 1}: "${title}"`);
      console.log(`  Start: ${startTime ? formatTime(startTime) : 'N/A'}`);
      console.log(`  End: ${endTime ? formatTime(endTime) : 'N/A'}`);
      console.log(`  All-day: ${isAllDay}`);
      console.log(`  Description: ${event.getDescription() ? event.getDescription().substring(0, 100) + '...' : 'None'}`);
    });

    console.log('✅ Sports calendar access test completed!');

  } catch (error) {
    console.error('❌ Sports calendar access test failed:', error);
    console.log('Error details:', {
      message: error.message,
      name: error.name,
      stack: error.stack
    });
  }
}

/**
 * Function to list all available calendars and check their access
 * @created 2025-12-30
 */
function listAllCalendars() {
  logger.info('Starting calendar listing process');

  try {
    const calendars = CalendarApp.getAllCalendars();

    logger.info('Calendars found', { count: calendars.length });

    calendars.forEach((calendar, index) => {
      const id = calendar.getId();
      const name = calendar.getName();
      const isOwned = calendar.isOwnedByMe();
      const isHidden = calendar.isHidden();
      const color = calendar.getColor();

      logger.info(`Calendar ${index + 1}`, {
        name: name,
        id: id,
        ownedByMe: isOwned,
        hidden: isHidden,
        color: color
      });

      // Check if this might be the sports calendar
      if (name.toLowerCase().includes('ספורט') || name.toLowerCase().includes('sport') ||
          id === CONFIG.SPECIAL_CALENDAR_ID) {
        logger.info('POSSIBLE SPORTS CALENDAR FOUND', { name: name, id: id });
      }
    });

    // Check the configured calendar IDs
    logger.info('Checking configured calendar IDs from CONFIG');
    const configuredIds = CONFIG.CALENDAR_IDS;
    configuredIds.forEach((id, index) => {
      try {
        const calendar = CalendarApp.getCalendarById(id);
        if (calendar) {
          logger.success(`Calendar ${index + 1} accessible`, {
            id: id,
            name: calendar.getName()
          });
        } else {
          logger.error(`Calendar ${index + 1} not accessible`, { id: id });
        }
      } catch (error) {
        logger.error(`Error accessing calendar ${index + 1}`, {
          id: id,
          error: error.message
        });
      }
    });

    logger.success('Calendar listing completed', { totalCalendars: calendars.length });

  } catch (error) {
    logger.error('Error listing calendars', { error: error.message });
  }
}

/**
 * Function to display all sports calendar events for a specific date
 * @param {string} dateStr - Date in YYYY-MM-DD format (optional, defaults to today)
 * @created 2025-12-30
 */
function showSportsCalendarEvents(dateStr = null) {
  console.log('🏆 Displaying sports calendar events...');

  try {
    const specialCalendarId = CONFIG.SPECIAL_CALENDAR_ID;

    if (!dateStr) {
      const today = new Date();
      dateStr = today.toISOString().split('T')[0];
    }

    console.log(`Checking sports calendar events for date: ${dateStr}`);

    // Parse date
    const dateParts = dateStr.split('-');
    const date = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));

    const startOfDay = new Date(date);
    startOfDay.setHours(0, 0, 0, 0);
    const endOfDay = new Date(date);
    endOfDay.setHours(23, 59, 59, 999);

    console.log(`Time range: ${startOfDay} to ${endOfDay}`);

    const calendar = CalendarApp.getCalendarById(specialCalendarId);

    if (!calendar) {
      console.log('❌ Could not access sports calendar');
      return;
    }

    const events = calendar.getEvents(startOfDay, endOfDay);

    console.log(`\n📅 Sports calendar events for ${dateStr}:`);
    console.log('─'.repeat(80));

    if (events.length === 0) {
      console.log('No events found for this date');
    } else {
      events.forEach((event, index) => {
        const startTime = event.getStartTime();
        const endTime = event.getEndTime();
        const title = event.getTitle();
        const isAllDay = event.isAllDayEvent();
        const description = event.getDescription();

        console.log(`\nEvent ${index + 1}:`);
        console.log(`  📝 Title: "${title}"`);
        console.log(`  🕐 Start: ${startTime ? formatTime(startTime) : 'N/A'} (${startTime ? startTime.toISOString() : 'N/A'})`);
        console.log(`  🕐 End: ${endTime ? formatTime(endTime) : 'N/A'} (${endTime ? endTime.toISOString() : 'N/A'})`);
        console.log(`  📅 All-day: ${isAllDay}`);
        console.log(`  📝 Description: ${description ? description.substring(0, 200) + (description.length > 200 ? '...' : '') : 'None'}`);

        // Calculate buffer times
        if (startTime && endTime && !isAllDay) {
          const bufferBefore = 30 * 60000; // 30 minutes
          const bufferAfter = 60 * 60000; // 1 hour

          const eventStartWithBuffer = new Date(startTime.getTime() - bufferBefore);
          const eventEndWithBuffer = new Date(endTime.getTime() + bufferAfter);

          console.log(`  ⏰ Blocked time range (including buffers):`);
          console.log(`     From: ${formatTime(eventStartWithBuffer)} (${eventStartWithBuffer.toISOString()})`);
          console.log(`     To: ${formatTime(eventEndWithBuffer)} (${eventEndWithBuffer.toISOString()})`);
        }
      });
    }

    console.log('\n' + '─'.repeat(80));
    console.log(`✅ Found ${events.length} events in sports calendar`);

  } catch (error) {
    console.error('❌ Error displaying sports calendar events:', error);
  }
}

/**
 * Comprehensive test to diagnose the double booking issue
 * Run this function to check sports calendar access, events, and time slot calculations
 * @param {string} dateStr - Date to test in YYYY-MM-DD format (optional, defaults to today)
 * @created 2025-12-30
 */
function diagnoseDoubleBooking(dateStr = null) {
  console.log('🔍 Starting comprehensive diagnosis of double booking issue...');
  console.log('=' .repeat(80));

  if (!dateStr) {
    const today = new Date();
    dateStr = today.toISOString().split('T')[0];
  }

  console.log(`Testing date: ${dateStr}`);
  console.log('');

  try {
    // Test 0: List all calendars
    console.log('0️⃣ Listing all available calendars...');
    listAllCalendars();
    console.log('');

    // Test 1: Sports calendar access
    console.log('1️⃣ Testing sports calendar access...');
    testSportsCalendarAccess();
    console.log('');

    // Test 2: Show sports calendar events for the date
    console.log('2️⃣ Checking sports calendar events for the problematic date...');
    showSportsCalendarEvents(dateStr);
    console.log('');

    // Test 3: Show available time slots with detailed logging
    console.log('3️⃣ Testing available time slots calculation...');
    const duration = 45; // Test with 45 minutes (typical appointment)
    console.log(`Getting available slots for ${dateStr}, duration: ${duration} minutes`);

    const slots = getAvailableTimeSlots(dateStr, duration);

    console.log(`\n📊 Results:`);
    console.log(`   - Found ${slots.length} available time slots`);
    if (slots.length > 0) {
      console.log('   - Available slots:');
      slots.forEach((slot, index) => {
        console.log(`     ${index + 1}. ${slot.start} - ${slot.end}`);
      });
    } else {
      console.log('   - No available slots found!');
    }
    console.log('');

    // Test 4: Time calculations
    console.log('4️⃣ Testing time calculation logic...');
    testTimeCalculations();
    console.log('');

    // Test 5: Show recent logs
    console.log('5️⃣ Checking recent logs...');
    showTodaysLogs();
    console.log('');

    console.log('=' .repeat(80));
    console.log('✅ Diagnosis completed!');
    console.log('');
    console.log('💡 If you see the issue, check:');
    console.log('   - Are there events in the sports calendar that should block slots?');
    console.log('   - Are the buffer calculations working correctly?');
    console.log('   - Are the available slots being calculated properly?');
    console.log('   - Check the logs above for any errors or unexpected behavior.');

  } catch (error) {
    console.error('❌ Diagnosis failed:', error);
    console.log('Error details:', {
      message: error.message,
      name: error.name,
      stack: error.stack
    });
  }
}

/**
 * Test function to verify time calculations for sports calendar buffers
 * @created 2025-12-30
 */
function testTimeCalculations() {
  console.log('🧪 Testing time calculations for sports calendar buffers...');

  try {
    // Test case: sports event from 10:00 to 11:00
    const eventStart = new Date();
    eventStart.setHours(10, 0, 0, 0);

    const eventEnd = new Date();
    eventEnd.setHours(11, 0, 0, 0);

    console.log('Test event: 10:00 - 11:00');

    // Calculate buffers
    const bufferBefore = 30 * 60000; // 30 minutes in milliseconds
    const bufferAfter = 60 * 60000; // 1 hour in milliseconds

    const eventStartWithBuffer = new Date(eventStart.getTime() - bufferBefore);
    const eventEndWithBuffer = new Date(eventEnd.getTime() + bufferAfter);

    console.log(`Event start: ${formatTime(eventStart)}`);
    console.log(`Event end: ${formatTime(eventEnd)}`);
    console.log(`With buffer - start: ${formatTime(eventStartWithBuffer)}`);
    console.log(`With buffer - end: ${formatTime(eventEndWithBuffer)}`);

    // Test various slot times
    const testSlots = [
      { start: '08:00', end: '09:00', description: 'Before buffer' },
      { start: '09:30', end: '10:30', description: 'Overlaps with buffer before' },
      { start: '10:00', end: '11:00', description: 'Direct overlap' },
      { start: '11:00', end: '12:00', description: 'Overlaps with buffer after' },
      { start: '12:30', end: '13:30', description: 'After buffer' }
    ];

    testSlots.forEach(slot => {
      const slotStartParts = slot.start.split(':');
      const slotEndParts = slot.end.split(':');

      const slotStart = new Date();
      slotStart.setHours(parseInt(slotStartParts[0]), parseInt(slotStartParts[1]), 0, 0);

      const slotEnd = new Date();
      slotEnd.setHours(parseInt(slotEndParts[0]), parseInt(slotEndParts[1]), 0, 0);

      const conflict = (slotStart < eventEndWithBuffer && slotEnd > eventStartWithBuffer);

      console.log(`${slot.description}: ${slot.start}-${slot.end} - Conflict: ${conflict}`);
    });

    console.log('✅ Time calculations test completed!');

  } catch (error) {
    console.error('❌ Time calculations test failed:', error);
  }
}

/**
 * Quick test function to verify the etag fix works
 * @created 2025-06-09
 */
function testEtagFix() {
  console.log('🧪 Testing etag fix for People API...');
  
  try {
    // Test a simple contact listing with the fixed personFields
    const connectionParams = {
      personFields: 'names,emailAddresses,phoneNumbers,metadata',
      pageSize: 10 // Small batch for testing
    };
    
    console.log('📞 Attempting People API call with fixed personFields...');
    const connections = People.People.Connections.list('people/me', connectionParams);
    
    if (connections && connections.connections) {
      console.log(`✅ Success! Retrieved ${connections.connections.length} contacts`);
      
      // Show first contact metadata to verify etag is available
      if (connections.connections.length > 0) {
        const firstContact = connections.connections[0];
        console.log('📊 First contact metadata:', {
          hasMetadata: !!firstContact.metadata,
          hasEtag: firstContact.metadata && firstContact.metadata.sources ? 
            firstContact.metadata.sources.some(s => s.etag) : false,
          resourceName: firstContact.resourceName
        });
      }
    } else {
      console.log('⚠️ No connections returned, but no error occurred');
    }
    
    console.log('✅ Etag fix test completed successfully!');
    
  } catch (error) {
    console.error('❌ Etag fix test failed:', error);
    console.log('💡 Error details:', {
      message: error.message,
      name: error.name
    });
  }
}

// ===== CUSTOM LOGGING SYSTEM =====

/**
 * Custom logging system that saves logs to Script Properties for 7 days
 * This ensures we can always access logs even when regular execution logs don't work
 */
class CustomLogger {
  constructor() {
    this.properties = PropertiesService.getScriptProperties();
    this.logPrefix = 'CUSTOM_LOG_';
    this.maxLogDays = 7; // Keep logs for 7 days
  }
  
  /**
   * Logs a message with timestamp and saves to Script Properties
   */
  log(level, message, data = null) {
    try {
      const now = new Date();
      const timestamp = now.toISOString();
      const dateKey = now.toISOString().split('T')[0]; // YYYY-MM-DD
      
      const logEntry = {
        timestamp: timestamp,
        level: level,
        message: message,
        data: data ? JSON.stringify(data) : null,
        readable_time: now.toLocaleString('he-IL', { timeZone: 'Asia/Jerusalem' })
      };
      
      // Get existing logs for today
      const todayLogsKey = this.logPrefix + dateKey;
      const existingLogs = this.properties.getProperty(todayLogsKey);
      const logsArray = existingLogs ? JSON.parse(existingLogs) : [];
      
      // Add new log entry
      logsArray.push(logEntry);
      
      // Save back to properties (limit to 100 entries per day to prevent size issues)
      if (logsArray.length > 100) {
        logsArray.shift(); // Remove oldest entry
      }
      
      this.properties.setProperty(todayLogsKey, JSON.stringify(logsArray));
      
      // Also log to regular console
      console.log(`[${level}] ${message}`, data || '');
      
      // Clean old logs
      this.cleanOldLogs();
      
    } catch (error) {
      console.error('CustomLogger error:', error);
    }
  }
  
  /**
   * Info level log
   */
  info(message, data = null) {
    this.log('INFO', message, data);
  }
  
  /**
   * Error level log
   */
  error(message, data = null) {
    this.log('ERROR', message, data);
  }
  
  /**
   * Warning level log
   */
  warn(message, data = null) {
    this.log('WARN', message, data);
  }
  
  /**
   * Success level log
   */
  success(message, data = null) {
    this.log('SUCCESS', message, data);
  }
  
  /**
   * Removes logs older than maxLogDays
   */
  cleanOldLogs() {
    try {
      const allProperties = this.properties.getProperties();
      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - this.maxLogDays);
      const cutoffDateStr = cutoffDate.toISOString().split('T')[0];
      
      for (const key in allProperties) {
        if (key.startsWith(this.logPrefix)) {
          const dateStr = key.replace(this.logPrefix, '');
          if (dateStr < cutoffDateStr) {
            this.properties.deleteProperty(key);
          }
        }
      }
    } catch (error) {
      console.error('Error cleaning old logs:', error);
    }
  }
  
  /**
   * Gets all logs for a specific date
   */
  getLogsForDate(dateStr) {
    try {
      const logsKey = this.logPrefix + dateStr;
      const logs = this.properties.getProperty(logsKey);
      return logs ? JSON.parse(logs) : [];
    } catch (error) {
      console.error('Error getting logs for date:', error);
      return [];
    }
  }
  
  /**
   * Gets all available log dates
   */
  getAvailableLogDates() {
    try {
      const allProperties = this.properties.getProperties();
      const dates = [];
      
      for (const key in allProperties) {
        if (key.startsWith(this.logPrefix)) {
          const dateStr = key.replace(this.logPrefix, '');
          dates.push(dateStr);
        }
      }
      
      return dates.sort().reverse(); // Most recent first
    } catch (error) {
      console.error('Error getting available log dates:', error);
      return [];
    }
  }
}

// Global logger instance
const logger = new CustomLogger();

/**
 * Displays all saved logs from the last 7 days
 * Run this function manually to see all logs when regular execution logs don't work
 */
function showAllSavedLogs() {
  console.log('==========================================');
  console.log('     📋 Displaying All Saved Logs');
  console.log('==========================================');
  
  try {
    const availableDates = logger.getAvailableLogDates();
    
    if (availableDates.length === 0) {
      console.log('❌ No saved logs found');
      return;
    }
    
    console.log(`📅 Found logs for ${availableDates.length} days:`);
    console.log('Available dates:', availableDates.join(', '));
    console.log('');
    
    availableDates.forEach(dateStr => {
      const logs = logger.getLogsForDate(dateStr);
      
      console.log(`\n📆 Date: ${dateStr} (${logs.length} entries)`);
      console.log('─'.repeat(70));
      
      logs.forEach((log, index) => {
        const emoji = log.level === 'ERROR' ? '❌' : 
                     log.level === 'WARN' ? '⚠️' : 
                     log.level === 'SUCCESS' ? '✅' : '📝';
        
        // Show both original timestamp and readable time
        const originalTime = new Date(log.timestamp).toLocaleString('he-IL', { timeZone: 'Asia/Jerusalem' });
        
        console.log(`${emoji} [${log.level}] Original Time: ${originalTime}`);
        console.log(`   📄 Message: ${log.message}`);
        
        if (log.data) {
          try {
            const parsedData = JSON.parse(log.data);
            console.log(`   📊 Data:`, parsedData);
          } catch (e) {
            console.log(`   📊 Data: ${log.data}`);
          }
        }
        
        // Show exact ISO timestamp for debugging
        console.log(`   🕐 ISO Timestamp: ${log.timestamp}`);
        
        if (index < logs.length - 1) {
          console.log('   ─'.repeat(50));
        }
      });
    });
    
    console.log('\n==========================================');
    console.log('✅ Finished displaying saved logs');
    console.log('==========================================');
    
  } catch (error) {
    console.error('❌ Error displaying logs:', error);
  }
}

/**
 * Shows logs only from today
 * Quick function to see today's activity
 */
function showTodaysLogs() {
  console.log('📅 Displaying Today\'s Logs');
  console.log('=' .repeat(40));
  
  try {
    const today = new Date().toISOString().split('T')[0];
    const todaysLogs = logger.getLogsForDate(today);
    
    if (todaysLogs.length === 0) {
      console.log('❌ No logs found for today');
      return;
    }
    
    console.log(`Found ${todaysLogs.length} entries for today:`);
    console.log('');
    
    todaysLogs.forEach((log, index) => {
      const emoji = log.level === 'ERROR' ? '❌' : 
                   log.level === 'WARN' ? '⚠️' : 
                   log.level === 'SUCCESS' ? '✅' : '📝';
      
      const originalTime = new Date(log.timestamp).toLocaleString('he-IL', { timeZone: 'Asia/Jerusalem' });
      
      console.log(`${emoji} [${originalTime}] ${log.message}`);
      
      if (log.data) {
        try {
          const parsedData = JSON.parse(log.data);
          console.log(`   📊 Data:`, parsedData);
        } catch (e) {
          console.log(`   📊 Data: ${log.data}`);
        }
      }
      
      console.log(`   🕐 ISO: ${log.timestamp}`);
      
      if (index < todaysLogs.length - 1) {
        console.log('   ─'.repeat(30));
      }
    });
    
  } catch (error) {
    console.error('❌ Error displaying today\'s logs:', error);
  }
}

/**
 * Clears all saved logs (use carefully!)
 */
function clearAllSavedLogs() {
  console.log('🗑️ Clearing all saved logs...');
  
  try {
    const allProperties = logger.properties.getProperties();
    let deletedCount = 0;
    
    for (const key in allProperties) {
      if (key.startsWith(logger.logPrefix)) {
        logger.properties.deleteProperty(key);
        deletedCount++;
      }
    }
    
    console.log(`✅ Deleted ${deletedCount} log files`);
    
  } catch (error) {
    console.error('❌ Error clearing logs:', error);
  }
} 