function importVKProfiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // CONFIG
  const configSheet = ss.getSheetByName("Config");
  const accessToken = configSheet.getRange("B1").getValue();
  const userIdRangeA1 = configSheet.getRange("B3").getValue();

  if (!accessToken || !userIdRangeA1) {
    throw new Error("Missing API token or USER_ID range");
  }

  const userIds = ss.getRange(userIdRangeA1)
    .getValues()
    .flat()
    .filter(Boolean);

  const sheetName = "VK_Profiles";
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clear();

  const headers = [
    "id",
    "first_name",
    "last_name",
    "followers_count",
    "last_seen_time",
    "last_seen_platform",
    "last_post_date",
    "posts_checked",
    "repost_ratio",
    "bot_score",
    "bot_label",
    "bot_confidence",
    "bot_reasons"
  ];
  sheet.appendRow(headers);

  const fields = [
    "photo_50",
    "followers_count",
    "last_seen"
  ].join(",");

  const rows = [];
  const chunkSize = 100;

  for (let i = 0; i < userIds.length; i += chunkSize) {
    const chunk = userIds.slice(i, i + chunkSize);

    const usersUrl =
      "https://api.vk.com/method/users.get" +
      "?user_ids=" + encodeURIComponent(chunk.join(",")) +
      "&fields=" + encodeURIComponent(fields) +
      "&access_token=" + accessToken +
      "&v=5.199";

    const usersResp = JSON.parse(UrlFetchApp.fetch(usersUrl).getContentText());
    if (usersResp.error) throw new Error(JSON.stringify(usersResp.error));

    for (const user of usersResp.response) {
      // ---- LAST SEEN ----
      let lastSeenDate = "";
      let lastSeenPlatform = "";
      if (user.last_seen) {
        lastSeenDate = formatDate(user.last_seen.time);
        lastSeenPlatform = user.last_seen.platform || "";
      }

      // ---- WALL ANALYSIS ----
      let postsChecked = 0;
      let reposts = 0;
      let engagementSum = 0;
      let lastPostDate = "";

      try {
        const wallUrl =
          "https://api.vk.com/method/wall.get" +
          "?owner_id=" + user.id +
          "&count=5" +
          "&access_token=" + accessToken +
          "&v=5.199";

        const wallResp = JSON.parse(UrlFetchApp.fetch(wallUrl).getContentText());

        if (wallResp.response && wallResp.response.items.length) {
          const posts = wallResp.response.items;
          postsChecked = posts.length;
          lastPostDate = formatDate(posts[0].date);

          posts.forEach(p => {
            if (p.copy_history) reposts++;
            engagementSum +=
              (p.likes?.count || 0) +
              (p.comments?.count || 0) +
              (p.reposts?.count || 0);
          });
        }
      } catch (e) {}

      const repostRatio = postsChecked ? reposts / postsChecked : 0;

      // ---- BOT SCORING + EXPLANATIONS ----
      let botScore = 0;
      let reasons = [];
      let signals = 0;

      if (!user.photo_50) {
        botScore += 20;
        reasons.push("no profile photo");
        signals++;
      }

      if (postsChecked === 0) {
        botScore += 15;
        reasons.push("no wall posts");
        signals++;
      }

      if (lastSeenDate && postsChecked === 0) {
        botScore += 15;
        reasons.push("last seen but never posted");
        signals++;
      }

      if (repostRatio >= 0.8 && postsChecked >= 3) {
        botScore += 15;
        reasons.push("≥80% reposts");
        signals++;
      }

      if ((user.followers_count || 0) < 10) {
        botScore += 10;
        reasons.push("very low follower count");
        signals++;
      }

      if (engagementSum === 0 && postsChecked > 0) {
        botScore += 10;
        reasons.push("zero engagement on recent posts");
        signals++;
      }

      // ---- LABEL ----
      let botLabel = "likely_human";
      if (botScore >= 60) botLabel = "likely_bot";
      else if (botScore >= 30) botLabel = "suspicious";

      // ---- CONFIDENCE ----
      let botConfidence = "low";
      if (signals >= 4) botConfidence = "high";
      else if (signals >= 2) botConfidence = "medium";

      rows.push([
        user.id,
        user.first_name || "",
        user.last_name || "",
        user.followers_count || 0,
        lastSeenDate,
        lastSeenPlatform,
        lastPostDate,
        postsChecked,
        repostRatio.toFixed(2),
        botScore,
        botLabel,
        botConfidence,
        reasons.join("; ")
      ]);
    }
  }

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
}

// UNIX → YYYY-MM-DD
function formatDate(unix) {
  return Utilities.formatDate(
    new Date(unix * 1000),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}
