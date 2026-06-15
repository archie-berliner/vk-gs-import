function importVKWall() {
  // Read configuration from a config sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName("Config");

  if (!configSheet) {
    throw new Error("Config sheet not found! Please create a sheet named 'Config' with:\nA1: 'VK_TOKEN', B1: your token\nA2: 'VK_DOMAIN', B2: your domain");
  }

  // Read config values (assuming format: A column = key, B column = value)
  var configData = configSheet.getRange("A:B").getValues();
  var config = {};
  for (var i = 0; i < configData.length; i++) {
    if (configData[i][0]) {
      config[configData[i][0]] = configData[i][1];
    }
  }

  var token   = config["VK_TOKEN"] || "";
  var domain  = config["VK_DOMAIN"] || "staciangstsuccs";

  if (!token) {
    throw new Error("VK_TOKEN not found in Config sheet!");
  }

  var count   = 100;                          // posts per request (max 100)
  var version = "5.199";

  // First, get the group/page ID from domain
  var resolveUrl = "https://api.vk.com/method/utils.resolveScreenName" +
    "?v=" + version +
    "&access_token=" + token +
    "&screen_name=" + domain;

  var resolveResponse = UrlFetchApp.fetch(resolveUrl);
  var resolveData = JSON.parse(resolveResponse.getContentText());

  if (!resolveData.response) {
    throw new Error("Could not resolve domain. Check domain name.\n" + JSON.stringify(resolveData, null, 2));
  }

  var ownerId = resolveData.response.object_id;
  if (resolveData.response.type === "group" || resolveData.response.type === "page") {
    ownerId = -ownerId; // Groups/pages have negative IDs
  }

  // Fetch wall posts
  var wallUrl = "https://api.vk.com/method/wall.get" +
    "?v=" + version +
    "&access_token=" + token +
    "&domain=" + domain +
    "&count=" + count;

  var response = UrlFetchApp.fetch(wallUrl);
  var data = JSON.parse(response.getContentText());

  if (!data.response || !data.response.items) {
    throw new Error("No response items from VK. Check token/domain/permissions.\n" + JSON.stringify(data, null, 2));
  }

  var items = data.response.items;

  Logger.log("Fetched " + items.length + " posts from VK");

  // Get or create sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DataVKposts");

  // If sheet doesn't exist, create it with headers
  if (!sheet) {
    sheet = ss.insertSheet("DataVKposts");
    var headers = [
      "post_id",
      "group_name",
      "date",
      "time",
      "is_pinned",
      "post_text",
      "Post_URL",
      "Image_URL",
      "views",
      "reach_subscribers",
      "reach_viral",
      "likes",
      "comments",
      "reposts"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    Logger.log("Created new sheet with headers");
  }

  // Get existing data
  var lastRow = sheet.getLastRow();
  var existingData = [];
  var postIdIndex = {}; // Map of post_id to row number

  if (lastRow > 1) {
    existingData = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    // Build index of existing post IDs
    for (var i = 0; i < existingData.length; i++) {
      var postId = existingData[i][0]; // post_id is first column
      postIdIndex[postId] = i + 2; // +2 because row 1 is header and arrays are 0-indexed
    }
  }

  Logger.log("Found " + Object.keys(postIdIndex).length + " existing posts in sheet");

  var updatedCount = 0;
  var addedCount = 0;
  var reachCallsSkipped = 0;

  items.forEach(function(item) {
    var postId        = item.id;
    var timestamp     = item.date ? new Date(item.date * 1000) : null;
    var date          = timestamp ? Utilities.formatDate(timestamp, "Europe/Moscow", "yyyy-MM-dd") : "";
    var time          = timestamp ? Utilities.formatDate(timestamp, "Europe/Moscow", "HH:mm:ss") : "";
    var isPinned      = item.is_pinned === 1 ? "YES" : "NO";
    var text          = item.text || "";
    var postUrl       = "https://vk.com/wall" + ownerId + "_" + postId;

    // Extract image URL (first photo attachment if exists)
    var imageUrl = "";
    if (item.attachments && item.attachments.length > 0) {
      for (var i = 0; i < item.attachments.length; i++) {
        if (item.attachments[i].type === "photo") {
          var photo = item.attachments[i].photo;
          // Get the largest available size
          if (photo.sizes && photo.sizes.length > 0) {
            imageUrl = photo.sizes[photo.sizes.length - 1].url;
          }
          break; // Only get first image
        }
      }
    }

    var views         = item.views ? item.views.count : "";
    var likes         = item.likes ? item.likes.count : "";
    var comments      = item.comments ? item.comments.count : "";
    var reposts       = item.reposts ? item.reposts.count : "";

    // Optimize reach data fetching
    var reachSubscribers = "";
    var reachViral = "";
    var hasNewReachData = false;
    var shouldFetchReach = false;

    // Determine if we should try to fetch reach data
    // Skip only for posts older than 30 days (VK stops collecting reach data for old posts)
    if (timestamp) {
      var postAgeInDays = (new Date().getTime() - timestamp.getTime()) / (1000 * 60 * 60 * 24);
      if (postAgeInDays <= 30) {
        shouldFetchReach = true;
      }
    } else {
      // If no timestamp, try to fetch anyway
      shouldFetchReach = true;
    }

    if (shouldFetchReach) {
      try {
        var reachUrl = "https://api.vk.com/method/stats.getPostReach" +
          "?v=" + version +
          "&access_token=" + token +
          "&owner_id=" + ownerId +
          "&post_ids=" + postId;

        var reachResponse = UrlFetchApp.fetch(reachUrl);
        var reachData = JSON.parse(reachResponse.getContentText());

        if (reachData.response && reachData.response.length > 0) {
          var reachInfo = reachData.response[0];
          // If we got actual non-zero values, use them
          if (reachInfo.reach_subscribers !== undefined &&
              reachInfo.reach_subscribers !== null &&
              reachInfo.reach_subscribers > 0) {
            reachSubscribers = reachInfo.reach_subscribers;
            hasNewReachData = true;
          }
          if (reachInfo.reach_viral !== undefined &&
              reachInfo.reach_viral !== null &&
              reachInfo.reach_viral > 0) {
            reachViral = reachInfo.reach_viral;
            hasNewReachData = true;
          }
        }
      } catch (e) {
        // If stats.getPostReach fails (missing permissions), leave fields empty
        Logger.log("Could not fetch reach data for post " + postId + ": " + e.message);
      }

      // Add small delay to avoid rate limiting
      Utilities.sleep(100);
    } else {
      // Skip API call for older posts
      reachCallsSkipped++;
    }

    var rowData = [
      postId,
      domain,  // group name
      date,
      time,
      isPinned,
      text,
      postUrl,
      imageUrl,
      views,
      reachSubscribers,
      reachViral,
      likes,
      comments,
      reposts
    ];

    // Check if post ID already exists
    if (postIdIndex[postId]) {
      // Update existing row, but preserve reach data if we don't have new data
      var rowNumber = postIdIndex[postId];

      // If we didn't get new reach data from API, preserve existing values
      if (!hasNewReachData) {
        var existingRowData = sheet.getRange(rowNumber, 1, 1, 14).getValues()[0];
        // Keep existing reach_subscribers (column 10, index 9)
        if (existingRowData[9] !== "" && existingRowData[9] !== null) {
          reachSubscribers = existingRowData[9];
        }
        // Keep existing reach_viral (column 11, index 10)
        if (existingRowData[10] !== "" && existingRowData[10] !== null) {
          reachViral = existingRowData[10];
        }

        // Update rowData with preserved values
        rowData[9] = reachSubscribers;
        rowData[10] = reachViral;
      }

      sheet.getRange(rowNumber, 1, 1, rowData.length).setValues([rowData]);
      updatedCount++;
    } else {
      // Add new row
      sheet.appendRow(rowData);
      addedCount++;
    }
  });

  // Sort the entire sheet by post_id (column 1) in ascending order
  if (sheet.getLastRow() > 1) {
    var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14);
    dataRange.sort(1); // Sort by column 1 (post_id) ascending
  }

  Logger.log("Import complete!");
  Logger.log("Updated: " + updatedCount + " posts");
  Logger.log("Added: " + addedCount + " new posts");
  Logger.log("Skipped reach API calls for " + reachCallsSkipped + " older posts (>30 days)");
  Logger.log("Total posts in sheet: " + (sheet.getLastRow() - 1));
}
