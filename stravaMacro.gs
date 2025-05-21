function getStravaAccessToken() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let accessToken = scriptProperties.getProperty(TOKEN_STORAGE_KEY);

  if (!accessToken || isTokenExpired()) {
    
    // Refresh the token
    const url = 'https://www.strava.com/oauth/token';
    const payload = {
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      grant_type: 'refresh_token',
      refresh_token: REFRESH_TOKEN
    };

    const options = {
      method: 'post',
      payload: payload
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    accessToken = data.access_token;
    const expiresAt = data.expires_at;

    // Save new token + expiry time
    scriptProperties.setProperty(TOKEN_STORAGE_KEY, accessToken);
    scriptProperties.setProperty('STRAVA_TOKEN_EXPIRES_AT', expiresAt);
  }

  return accessToken;
}

function isTokenExpired() {
  const expiresAt = PropertiesService.getScriptProperties().getProperty('STRAVA_TOKEN_EXPIRES_AT');
  const currentTime = Math.floor(Date.now() / 1000);
  return !expiresAt || currentTime >= parseInt(expiresAt);
}

function getStravaActivities() {
  const accessToken = getStravaAccessToken();

  const url = 'https://www.strava.com/api/v3/athlete/activities?per_page=100';

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + accessToken
    }
  });

  const activities = JSON.parse(response.getContentText());

  const sheet = getOrCreateSheet('Strava API Data');
  
  let existingIds = [];

  if (sheet.getLastRow() > 1) {
    const existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    existingIds = existingData.flat().map(id => id.toString());
  }

  activities.forEach(activity => {

    if (!activity.id) {
      // This is not a real activity (possibly header or bad row), skip
      return;
    }

    const activityId = activity.id.toString();

    if (existingIds.includes(activityId)) {
      // Already logged
      Logger.log("skipping ID:", activityId)
      return;
    }

    const dateRaw = activity.start_date_local || activity.start_date;
    const dateObj = new Date(dateRaw);
    const date = Utilities.formatDate(dateObj, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "MM/dd/yyyy");

    const title = activity.name;
    const distanceMeters = activity.distance;
    const distanceMiles = distanceMeters / 1609.34;
    const movingTimeSec = activity.moving_time;
    const movingTimeFormatted = formatSecondsToHMS(movingTimeSec);

    const elevationGain = activity.total_elevation_gain * 3.28084;

    // Avoid division by zero
    const avgPace = distanceMiles > 0 ? (movingTimeSec / 60) / distanceMiles : 0;
    const avgPaceMinutes = Math.floor(avgPace);
    const avgPaceSeconds = Math.round((avgPace - avgPaceMinutes) * 60);
    const avgPaceFormatted = `${avgPaceMinutes}:${avgPaceSeconds.toString().padStart(2, "0")}`;

    if (!existingIds.includes(activity.id.toString())) {
      sheet.appendRow([
        activity.id,
        date,
        title,
        distanceMiles.toFixed(2),
        movingTimeFormatted,
        avgPaceFormatted,
        elevationGain.toFixed(0)
      ]);
    }

  });

  // Sort sheet by Date (Column B, which is index 2)
  // Skip header row
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort({column: 2, ascending: false});
}

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getLastRow() < 2) {
    sheet.clear(); // Make sure it's clean
    sheet.appendRow(['Activity ID', 'Date', 'Activity Title', 'Distance (miles)', 'Moving Time', 'Average Pace (min/mile)', 'Elevation Gain (ft)', 'GAP', 'Regularized GAP']);
  }

  return sheet;
}

function formatSecondsToHMS(seconds) {
  const hrs = Math.floor(seconds / 3600);
  const mins = Math.floor((seconds % 3600) / 60);
  const secs = seconds % 60;
  return `${hrs}:${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
}

function estimateGAP(avgPace, elevationGainFeet, distanceMiles) {
  if (distanceMiles === 0) return avgPace;

  const avgGrade = elevationGainFeet / (distanceMiles * 5280);

  // Approximate pace penalty per grade % (about +20 sec per 1% grade gain)
  const pacePenaltyPerPercent = 20;

  const paceAdjustment = avgGrade * 100 * pacePenaltyPerPercent; // seconds per mile

  const gapInMinutes = avgPace + (paceAdjustment / 60);

  return gapInMinutes;
}
