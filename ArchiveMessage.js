const CHANNNEL_LIST_BASE_URL = "https://slack.com/api/conversations.list";
const CHANNNEL_HISTORY_BASE_URL = "https://slack.com/api/conversations.history";

const USER_INFO_BASE_URL = "https://slack.com/api/users.info";

const REPLY_LIST_BASE_URL = "https://slack.com/api/conversations.replies";

const ATTACHMENTS_FOLDER = DriveApp.getFolderById("1cys1At9ByVNlx6KAOAyO8W28o9ZShC1e");

var properties = PropertiesService.getScriptProperties();
const CHANNNEL_ADMIN_AUTH_TOKEN = properties.getProperty("CHANNNEL_ADMIN_AUTH_TOKEN");
const BACKUP_SHEET_ID = properties.getProperty("BACKUP_SHEET_ID");

var usersInfo = {};

function backUp() {
  var ss = SpreadsheetApp.openById(BACKUP_SHEET_ID);
  deleteSheets(ss);
  var ss_main = ss.insertSheet("メイン");
  var ss_mainURL = getNewSheetURL(ss, ss_main);
  ss_main.getRange(1, 1, 1, 5).setValues([["パブリック", "チャンネルID", "チャンネル名", "トピック", "リンク"]]);
  var channelsList = getAllChannels();
  var channelNum = 0;
  for (var channelInfo of channelsList) {
    var channelId = channelInfo.id,
      channelName = channelInfo.name;
    var channelSheet = ss.insertSheet(channelName);
    var channelSheetURL = getNewSheetURL(ss, channelSheet);
    ss_main.getRange(ss_main.getLastRow() + 1, 1, 1, 5).setValues([[
      channelInfo.isPrivate ? "" : "〇",
      channelId,
      channelName,
      channelInfo.topic,
      `=HYPERLINK("${channelSheetURL}", "リンク＞")`
    ]]);
    channelSheet.getRange(1, 1, 2, 7).setValues([
      [`=HYPERLINK("${ss_mainURL}", "＜メインへ戻る")`, "", "", "", "", "", ""],
      ["発言者", "発言内容", "スレッドリンク", "添付ファイル", "リアクション", "ts", "userID"]]);
    getAllMessageInChannel(ss, channelId, channelSheet, channelSheetURL);
    decorateCells(channelSheet);
    cutBlankCells(channelSheet);
    console.log(`done: ${channelNum} ${channelName}`);
    channelNum++;
  }
  console.log(`everything done!`);
}

function backUpContinue() {
  var channelIndex = 7;
  var ss = SpreadsheetApp.openById(BACKUP_SHEET_ID);
  var ss_main = ss.getSheetByName("メイン");
  var ss_mainURL = getNewSheetURL(ss, ss_main);
  var channelsList = getAllChannels();
  for (var channelInfo of channelsList.slice(channelIndex)) {
    var channelId = channelInfo.id,
      channelName = channelInfo.name;
    var oldSheet = ss.getSheetByName(channelName)
    if (oldSheet) {
      ss.deleteSheet(oldSheet);
    }
    var channelSheet = ss.insertSheet(channelName);
    var channelSheetURL = getNewSheetURL(ss, channelSheet);
    ss_main.getRange(channelIndex + 2, 1, 1, 5).setValues([[ // backUp()で失敗したchannelから
      channelInfo.isPrivate ? "" : "〇",
      channelId,
      channelName,
      channelInfo.topic,
      `=HYPERLINK("${channelSheetURL}", "リンク＞")`
    ]]);
    channelSheet.getRange(1, 1, 2, 7).setValues([
      [`=HYPERLINK("${ss_mainURL}", "＜メインへ戻る")`, "", "", "", "", "", ""],
      ["発言者", "発言内容", "スレッドリンク", "添付ファイル", "リアクション", "ts", "userID"]]);
    getAllMessageInChannel(ss, channelId, channelSheet, channelSheetURL);
    decorateCells(channelSheet);
    cutBlankCells(channelSheet);
    console.log(`done: ${index} ${channelName}`);
    index++;
  }
  console.log(`everything done!`);
}

function getAllChannels() {
  var res = requestToSlackAPI(CHANNNEL_LIST_BASE_URL, { "limit": 100, "types": "public_channel, private_channel" });
  console.log(res);
  var channelsList = [];
  if (!res.ok) return;
  var channels = res.channels;
  for (var channel of channels) {
    if (channel.is_archived) continue; // アーカイブされたものの除外
    channelsList.push({ "name": channel.name, "id": channel.id, "isPrivate": channel.is_private, "topic": channel.topic.value });
  }
  channelsList.sort(function (a, b) {
    if (a.name < b.name) return -1;
    if (a.name > b.name) return 1;
    return 0;
  });
  return channelsList;
}

function getAllMessageInChannel(ss, testChannelID, channelSheet, channelSheetURL) {
  var hasMore = true;
  var option = { "channel": testChannelID, "limit": 3000 };
  while (hasMore) {
    var res = requestToSlackAPI(CHANNNEL_HISTORY_BASE_URL, option);
    if (!res.ok) return;
    hasMore = res.has_more;
    if (hasMore) {
      option["cursor"] = res.response_metadata.next_cursor;
    }
    var messages = res.messages;
    var messageList = [];
    for (var message of messages) {
      console.log(message)
      var files = message.files;
      var fileUrls = [];
      if (!(files === undefined)) {
        for (var file of files) {
          if (file.mode == 'tombstone' || file.mode == 'hidden_by_limit') {
            continue;
          }
          var fileName = `${testChannelID}_${message.ts}_${file.name}`;
          var url = file.url_private_download;
          if (url) {
            fileUrls.push(downloadData(url, fileName));
          }
        }
      }
      var text = message.text;
      var user = message.user;
      var reactions = message.reactions;
      if (usersInfo[user] == undefined) {
        var userInfo = requestToSlackAPI(USER_INFO_BASE_URL, { "user": user });
        if (userInfo.user == undefined) continue;
        usersInfo[user] = (userInfo.user.profile.display_name == "") ? userInfo.user.real_name : userInfo.user.profile.display_name;
      }
      var flag = true;
      while (flag) {
        var textPoint = text.search(/<@U.{10}>/);
        // console.log(textPoint);
        if (textPoint == -1) {
          flag = false;
          continue;
        }
        var mentionUser = text.slice(textPoint + 2, textPoint + 13);
        if (usersInfo[mentionUser] == undefined) {
          var userInfo = requestToSlackAPI(USER_INFO_BASE_URL, { "user": mentionUser });
          usersInfo[mentionUser] = (userInfo.user.profile.display_name == "") ? userInfo.user.real_name : userInfo.user.profile.display_name;
        }
        text = text.slice(0, textPoint) + `@${usersInfo[mentionUser]}` + text.slice(textPoint + 14);
      }
      messageList.push(`【${usersInfo[user]}】\n${text}`);
      var threadURL = getAllReplyInMessage(ss, testChannelID, message.ts, channelSheetURL);
      channelSheet.getRange(channelSheet.getLastRow() + 1, 1, 1, 7).setValues([[
        usersInfo[user],
        text,
        threadURL != "" ? `=HYPERLINK("${threadURL}", "リンク＞")` : "",
        fileUrls.join(", "),
        (reactions != undefined)? `{ "reactions": ${JSON.stringify(reactions)} }` : "",
        message.ts,
        user]])
    }
  }
}

function getAllReplyInMessage(ss, channelID, messageTs, channelSheetURL) {
  var hasMore = true;
  var messageList = [];
  var option = { "channel": channelID, "ts": messageTs, "limit": 100 };
  while (hasMore) {
    var res = requestToSlackAPI(REPLY_LIST_BASE_URL, option);
    if (!res.ok) {
      console.log(res.error);
      return "";
    }

    hasMore = res.has_more;
    if (hasMore) {
      option["cursor"] = res.response_metadata.next_cursor;
    }
    var messages = res.messages;
    if (messages[0].reply_count == undefined) {
      return "";
    }
    for (var message of messages) {
      var files = message.files;
      var fileUrls = [];
      if (!(files === undefined)) {
        for (var file of files) {
          if (file.mode == 'tombstone' || file.mode == 'hidden_by_limit') {
            continue;
          }
          var fileName = `${channelID}_${message.ts}_${file.name}`;
          var url = file.url_private_download;
          if (url) {
            fileUrls.push(downloadData(url, fileName));
          }
        }
      }
      var text = message.text;
      var user = message.user;
      var reactions = message.reactions;
      if (usersInfo[user] == undefined) {
        var userInfo = requestToSlackAPI(USER_INFO_BASE_URL, { "user": user });
        if (userInfo.user == undefined) continue;
        usersInfo[user] = (userInfo.user.profile.display_name == "") ? userInfo.user.real_name : userInfo.user.profile.display_name;
      }
      var flag = true;
      while (flag) {
        var textPoint = text.search(/<@U.{10}>/);
        // console.log(textPoint);
        if (textPoint == -1) {
          flag = false;
          continue;
        }
        var mentionUser = text.slice(textPoint + 2, textPoint + 13)
        if (usersInfo[mentionUser] == undefined) {
          var userInfo = requestToSlackAPI(USER_INFO_BASE_URL, { "user": mentionUser });
          usersInfo[mentionUser] = (userInfo.user.profile.display_name == "") ? userInfo.user.real_name : userInfo.user.profile.display_name;
        }
        text = text.slice(0, textPoint) + `@${usersInfo[mentionUser]}` + text.slice(textPoint + 14);
      }
      messageList.unshift([
        usersInfo[user],
        text,
        fileUrls.join(", "),
        (reactions != undefined)? `{ "reactions": ${JSON.stringify(reactions)} }` : "",
        message.ts,
        user
      ]);
    }
  }
  messageList.unshift([`=HYPERLINK("${channelSheetURL}", "＜親チャンネルへ")`, "", "", "", "", ""], ["発言者", "発言内容", "添付ファイル", "リアクション", "ts", "userID"]);
  var oldSheet = ss.getSheetByName(messageTs)
  if (oldSheet) {
    ss.deleteSheet(oldSheet);
  }
  var threadSheet = ss.insertSheet(messageTs);
  threadSheet.getRange(1, 1, messageList.length, 6).setValues(messageList);
  decorateCells(threadSheet);
  cutBlankCells(threadSheet);
  return getNewSheetURL(ss, threadSheet);
}

function deleteSheets(spreadsheet) {
  var sheetcount = spreadsheet.getNumSheets();
  for (var i = sheetcount; i > 1; i--) {
    var sh = spreadsheet.getSheets()[i - 1];
    spreadsheet.deleteSheet(sh);
  }
}

function tests() {
  var ss = SpreadsheetApp.openById(BACKUP_SHEET_ID);
  var ss_main = ss.getSheetByName("メイン");
  ss_main.moveActiveSheet(2);
}

function test() {
  console.log(getAllChannels().slice(0));
}

function requestToSlackAPI(url, parameters) {
  while (true) {
    var response = UrlFetchApp.fetch(`${url}?${hashToQuery(parameters)}`, {
      method: 'GET',
      headers: { "Content-Type": 'application/x-www-form-urlencoded', 'Authorization': 'Bearer ' + CHANNNEL_ADMIN_AUTH_TOKEN },
      muteHttpExceptions: true
    });
    var response_code = response.getResponseCode();
    if (response_code === 200) {
      return JSON.parse(response.getContentText());
    } else if (response_code === 429) {
      Utilities.sleep(10000);
    }
  }
}

function downloadData(url, fileName) {
  var options = {
    "headers": { 'Authorization': 'Bearer ' + CHANNNEL_ADMIN_AUTH_TOKEN }
  };
  var folder = ATTACHMENTS_FOLDER;
  try {
    var response = UrlFetchApp.fetch(url, options);
    var fileBlob = response.getBlob().setName(fileName);
    var itr = folder.getFilesByName(fileName);
    if (itr.hasNext()) {
      itr.next().setTrashed(true);
    }
    var file = folder.createFile(fileBlob);
    var driveFile = DriveApp.getFileById(file.getId());
    return driveFile.getUrl();
  } catch (error) {
    return error.lineNumber + error.message + error.stack;
  }
}

function hashToQuery(hashList) {
  var result = [];
  for (var key in hashList) {
    result.push(`${key}=${hashList[key]}`);
  }
  return result.join("&");
}

function getNewSheetURL(ss, newSheet) {
  return ss.getUrl() + "#gid=" + newSheet.getSheetId();
}

function cutBlankCells(sh) {
  sh.deleteRows(sh.getLastRow() + 1, sh.getMaxRows() - sh.getLastRow() - 1);
  sh.deleteColumns(sh.getLastColumn() + 1, sh.getMaxColumns() - sh.getLastColumn() - 1);
}

function decorateCells(sh) {
  var rowNum = sh.getLastRow();
  var columnNum = sh.getLastColumn();
  var allCells = sh.getRange(1, 1, rowNum, columnNum);
  allCells.setVerticalAlignments(initQuadraticArray(rowNum, columnNum, "middle"));
  allCells.setWraps(initQuadraticArray(rowNum, columnNum, false));
  sh.setColumnWidth(1, 230);
  sh.setColumnWidth(2, 1050);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 130);
  sh.setColumnWidth(5, 130);
  sh.setColumnWidth(6, 130);
  sh.setColumnWidth(7, 130);
  sh.setFrozenRows(2);
}

function initQuadraticArray(rowSize, columnSize, arg) {
  var res = [];
  for (var i = 0; i < rowSize; i++) {
    var row = [];
    for (var j = 0; j < columnSize; j++) {
      row.push(arg);
    }
    res.push(row);
  }
  return res;
}

function deleteSheet() {
  var name = "";
  var ss = SpreadsheetApp.openById(BACKUP_SHEET_ID);
  var ss_main = ss.getSheetByName(name);
  ss.deleteSheet(ss_main)
}
