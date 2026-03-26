/**
 * Entry point for the Google Apps Script Web App.
 * Serves the initial HTML file.
 *
 * @param {Object} e The event parameter for the web app URL request.
 * @return {HtmlOutput} The rendered HTML interface.
 */
function doGet(e) {
  // Use HtmlService to serve the 'Dashboard' file
  // Evaluate allows processing of scriptlets (like include) if added later
  const htmlOutput = HtmlService.createTemplateFromFile('Dashboard').evaluate();

  // Add meta tag to force viewport scaling, required for responsive design on mobile devices
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');

  // Set the title for the browser tab
  htmlOutput.setTitle('Hotel Management Dashboard');

  return htmlOutput;
}

/**
 * Example function that the frontend will call via google.script.run.getDashboardData()
 *
 * @return {string} A JSON stringified object containing the dashboard data.
 */
function getDashboardData() {
  // In a real application, you would connect to the active spreadsheet:
  // const ss = SpreadsheetApp.getActiveSpreadsheet();
  // const roomsSheet = ss.getSheetByName('Rooms');
  // const bookingsSheet = ss.getSheetByName('Bookings');

  // Here we return mock data that matches what the frontend expects.
  // This simulates reading from the database.
  const mockData = {
    metrics: {
      availableRooms: 15,
      todayArrivals: 5,
      todayDepartures: 3
    },
    rooms: [
      { number: '101', type: 'Standard Queen', status: 'Available' },
      { number: '102', type: 'Standard Queen', status: 'Occupied' },
      { number: '103', type: 'Deluxe King', status: 'Available' },
      { number: '104', type: 'Deluxe King', status: 'Maintenance' },
      { number: '201', type: 'Ocean Suite', status: 'Occupied' },
      { number: '202', type: 'Ocean Suite', status: 'Available' },
      { number: '203', type: 'Standard Twin', status: 'Available' },
      { number: '204', type: 'Standard Twin', status: 'Occupied' }
    ],
    roster: [
      { id: 'BK-1002', guestName: 'Alice Johnson', room: '102', status: 'Departure' },
      { id: 'BK-1005', guestName: 'Michael Smith', room: '201', status: 'In-House' },
      { id: 'BK-1008', guestName: 'Elena Rodriguez', room: '101', status: 'Arrival' },
      { id: 'BK-1009', guestName: 'David Chen', room: 'TBD', status: 'Arrival' }
    ]
  };

  // Using JSON.stringify ensures robust data transfer back to the frontend
  return JSON.stringify(mockData);
}
