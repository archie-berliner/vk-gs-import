function importVKFollowers() {
  // Read configuration from Config sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName("Config");

  if (!configSheet) {
    throw new Error("Config sheet not found! Please create a sheet named 'Config' with:\nA1: 'VK_TOKEN', B1: your token\nA2: 'VK_DOMAIN', B2: your group domain");
  }

  // Read config values
  var configData = configSheet.getRange("A:B").getValues();
  var config = {};
  for (var i = 0; i < configData.length; i++) {
    if (configData[i][0]) {
      config[configData[i][0]] = configData[i][1];
    }
  }

  var token = config["VK_TOKEN"] || "";
  var domain = config["VK_DOMAIN"] || "";

  if (!token) {
    throw new Error("VK_TOKEN not found in Config sheet!");
  }

  if (!domain) {
    throw new Error("VK_DOMAIN not found in Config sheet! Please add:\nA2: 'VK_DOMAIN', B2: your group domain (e.g., 'mygroup' or 'club12345')");
  }

  var version = "5.199";
  var count = 1000; // Max per request

  // First, resolve the domain to get group ID
  var resolveUrl = "https://api.vk.com/method/utils.resolveScreenName" +
    "?v=" + version +
    "&access_token=" + token +
    "&screen_name=" + domain;

  var resolveResponse = UrlFetchApp.fetch(resolveUrl);
  var resolveData = JSON.parse(resolveResponse.getContentText());

  if (!resolveData.response) {
    throw new Error("Could not resolve domain. Check domain name.\n" + JSON.stringify(resolveData, null, 2));
  }

  if (resolveData.response.type !== "group" && resolveData.response.type !== "page") {
    throw new Error("The domain must be a group or page, not a user. Type found: " + resolveData.response.type);
  }

  var groupId = resolveData.response.object_id;

  Logger.log("Fetching followers for group: " + domain + " (ID: " + groupId + ")");

  // Get current followers from VK API
  var allFollowers = [];
  var offset = 0;
  var totalCount = 0;

  // Fetch all followers (paginated)
  do {
    var followersUrl = "https://api.vk.com/method/groups.getMembers" +
      "?v=" + version +
      "&access_token=" + token +
      "&group_id=" + groupId +
      "&count=" + count +
      "&offset=" + offset +
      "&fields=first_name,last_name,photo_50,city,country";

    var response = UrlFetchApp.fetch(followersUrl);
    var data = JSON.parse(response.getContentText());

    if (!data.response) {
      throw new Error("Failed to fetch group members. Response:\n" + JSON.stringify(data, null, 2));
    }

    totalCount = data.response.count;
    var items = data.response.items || [];

    // Add fetched followers to our array
    allFollowers = allFollowers.concat(items);

    offset += count;

    Logger.log("Fetched " + allFollowers.length + " of " + totalCount + " members");

    // Add delay to avoid rate limiting
    if (offset < totalCount) {
      Utilities.sleep(350); // VK API allows ~3 requests per second
    }

  } while (offset < totalCount);

  Logger.log("Total group members fetched: " + allFollowers.length);

  // Get or create the Group Members sheet
  var followersSheet = ss.getSheetByName("DataVKMembers");

  if (!followersSheet) {
    followersSheet = ss.insertSheet("DataVKMembers");
    var headers = [
      "user_id",
      "first_name",
      "last_name",
      "full_name",
      "profile_url",
      "photo_url",
      "city",
      "country",
      "first_seen_date",
      "last_checked_date",
      "status"
    ];
    followersSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    followersSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    followersSheet.setFrozenRows(1);
    Logger.log("Created new DataVKMembers sheet with headers");
  }

  // Get existing members from sheet
  var lastRow = followersSheet.getLastRow();
  var existingFollowers = {};
  var existingData = [];

  if (lastRow > 1) {
    existingData = followersSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    // Build index of existing members
    for (var i = 0; i < existingData.length; i++) {
      var followerId = existingData[i][0]; // user_id is first column
      existingFollowers[followerId] = {
        row: i + 2,
        firstSeenDate: existingData[i][8],
        status: existingData[i][10]
      };
    }
  }

  Logger.log("Found " + Object.keys(existingFollowers).length + " existing members in sheet");

  // Build current followers index
  var currentFollowers = {};
  for (var i = 0; i < allFollowers.length; i++) {
    currentFollowers[allFollowers[i].id] = true;
  }

  var newFollowersCount = 0;
  var updatedFollowersCount = 0;
  var leftGroupCount = 0;
  var currentDate = Utilities.formatDate(new Date(), "Europe/Moscow", "yyyy-MM-dd HH:mm:ss");

  // Process each member from VK API
  allFollowers.forEach(function(follower) {
    var userId = follower.id;
    var firstName = follower.first_name || "";
    var lastName = follower.last_name || "";
    var fullName = (firstName + " " + lastName).trim();
    var profileUrl = "https://vk.com/id" + userId;
    var photoUrl = follower.photo_50 || "";
    var city = follower.city ? follower.city.title : "";
    var country = follower.country ? follower.country.title : "";

    if (existingFollowers[userId]) {
      // Existing member - update their info but PRESERVE first_seen_date
      var rowNumber = existingFollowers[userId].row;
      var firstSeenDate = existingFollowers[userId].firstSeenDate;

      var rowData = [
        userId,
        firstName,
        lastName,
        fullName,
        profileUrl,
        photoUrl,
        city,
        country,
        firstSeenDate, // PRESERVE original first_seen_date
        currentDate,   // Update last checked date
        "Member"       // Mark as active member
      ];

      followersSheet.getRange(rowNumber, 1, 1, rowData.length).setValues([rowData]);
      updatedFollowersCount++;

    } else {
      // New member - add to sheet with current date as first_seen_date
      var rowData = [
        userId,
        firstName,
        lastName,
        fullName,
        profileUrl,
        photoUrl,
        city,
        country,
        currentDate,   // First seen date (new member)
        currentDate,   // Last checked date
        "Member"       // Status
      ];

      followersSheet.appendRow(rowData);
      newFollowersCount++;
    }
  });

  // Mark members who left the group
  for (var existingId in existingFollowers) {
    if (!currentFollowers[existingId]) {
      var rowNumber = existingFollowers[existingId].row;
      var currentStatus = existingFollowers[existingId].status;

      // Only update if not already marked as left
      if (currentStatus !== "Left") {
        followersSheet.getRange(rowNumber, 11).setValue("Left");
        followersSheet.getRange(rowNumber, 10).setValue(currentDate); // Update last checked date
        leftGroupCount++;
      }
    }
  }

  // Sort sheet by user_id
  if (followersSheet.getLastRow() > 1) {
    var dataRange = followersSheet.getRange(2, 1, followersSheet.getLastRow() - 1, 11);
    dataRange.sort(1); // Sort by column 1 (user_id) ascending
  }

  Logger.log("Import complete!");
  Logger.log("New members: " + newFollowersCount);
  Logger.log("Updated members: " + updatedFollowersCount);
  Logger.log("Left group: " + leftGroupCount);
  Logger.log("Total active members in sheet: " + allFollowers.length);
}
