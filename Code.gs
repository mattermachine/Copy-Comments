/**
 * On opening of the document, add an item "Copy" which calls func start, onClick
 */
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Copy Comments', 'start')
      .addToUi();
}

/**
 * On installation, call onOpen
 */
function onInstall(e) {
  onOpen(e);
}


/**
 * ------------------------------------------------------------------------------------- *
 */


var SECOND = 1000;
var MINUTE = 60*SECOND;
var MAX_RUNNING_TIME = 4*MINUTE;
var TIME_TO_WAIT = 1.2*MINUTE;

/**
 * Run this on a GDrive Copy Log sheet
 */
function start() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  // Cell G2 holds track of which doc comments are being copied of.
  // This allows us to re-run the scrip multiple times without redoing docs already done,
  // and hand edit the counter if necessary.
  // var activeRow = sheet.getRange("G2:G2").getCell(1, 1).getValue();
  var activeRow = sheet.getRange("G2").getValue();

  var documentProperties = PropertiesService.getDocumentProperties();
  // documentProperties.deleteAllProperties();

  var props = {};
  props["pageToken"] = "";
  //props["activeRow"] = activeRow;

  documentProperties.setProperties(props);
  Logger.log(documentProperties);

  resume();
}


function resume(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var documentProperties = PropertiesService.getDocumentProperties();
  var props = documentProperties.getProperties();
  var pageToken = props["pageToken"];
  //var activeRow = props["activeRow"];
  var activeRow = sheet.getRange("G2").getValue();
  var docRows = sheet.getRange("D5:G1000")
  var id = docRows.getCell(activeRow, 1).getValue();
  var oldId = docRows.getCell(activeRow, 2).getValue();

  if (oldId == "") {
    return;
  }

  if (DriveApp.getFileById(id).getMimeType() == "application/vnd.google-apps.folder") {
    docRows.getCell(activeRow, 4).setValue("folder");
    sheet.getRange("G2").setValue(activeRow+1);
    resume();
    return;
  }


  var startTime = (new Date()).getTime();
  var comments = [];

  if(pageToken == "undefined"){
    appendCommentsAndReplies(id, comments);
    docRows.getCell(activeRow, 4).setValue("done");
    sheet.getRange("G2").setValue(activeRow+1);
    props["pageToken"] = "";
    documentProperties.setProperties(props);
    resume();
    return;
  }

  do
  {
    var commentRes = getComments(oldId, pageToken);
    comments = comments.concat(commentRes.items);
    pageToken = commentRes.nextPageToken;
    var currTime = (new Date()).getTime();
    if(currTime - startTime > MAX_RUNNING_TIME){
      var properties = {};
      if(pageToken == undefined){
        pageToken = "undefined";
      }
      properties["pageToken"] = pageToken;
      properties["newDocId"] = id;
      documentProperties.setProperties(properties);
      appendCommentsAndReplies(id, comments);
      var endTime = (new Date()).getTime();
      ScriptApp.newTrigger("resume")
               .timeBased()
               .at(new Date(endTime+TIME_TO_WAIT))
               .create();
      return;
    }
  }
  while (pageToken != undefined);

  appendCommentsAndReplies(id, comments);
  docRows.getCell(activeRow, 4).setValue("copied " + comments.length + " comments");
  sheet.getRange("G2").setValue(activeRow+1);
  props["pageToken"] = "";
  documentProperties.setProperties(props);
  resume();
}



/**
 * Returns comments in sets of 50
 * /Takes in the documentID and a reference to which set of 50 comments to retrieve
 * -Set optional arguments of the list of comments
 * -If there is a pageToken(reference to the next set of 50 comments), then set that argument
 */
function getComments(docId, prevToken) {
  var optionalArgs = {};
  optionalArgs["maxResults"] = 50;
  optionalArgs["includeDeleted"] = false;
  if(prevToken != ""){
    optionalArgs["pageToken"] = prevToken;
  }
  var comments = Drive.Comments.list(docId, optionalArgs);
  return comments;
}


/**
 * Appends the comments
 * Then appends the replies to each comment
 * Comments made by other authors/people will be created as the user using this add-on
 *
 * Slicing the replyResource is necessary because objects are mutable
 */
function appendCommentsAndReplies(id, comments){
  var fileId = DriveApp.getFileById(id).getId();
  for(var commentRes in comments){
    var commentId = comments[commentRes].commentId;
    var replySave = comments[commentRes].replies.slice();
    comments[commentRes].replies = [];
    var newComment = Drive.Comments.insert(comments[commentRes], id);
    if(comments[commentRes].author["isAuthenticatedUser"] == false){
      var origContent = comments[commentRes].content;
      var authorName = comments[commentRes].author["displayName"];
      //var newContent = "<html>help</html>";
      var newContent = "\"" + authorName + "\"" + ": \n---------------------\n" + origContent;
      Drive.Comments.patch({'content':newContent}, id, newComment.commentId);
    }
    comments[commentRes].replies = replySave;
    if(replySave.length != 0){
      for(var reply in replySave){
        var newReply = Drive.Replies.insert(replySave[reply], fileId, newComment.commentId);
        if(replySave[reply].author["isAuthenticatedUser"] == false){
          var origContent = replySave[reply].content;
          var authorName = replySave[reply].author["displayName"];
          var newContent = "\"" + authorName + "\"" + ": \n---------------------\n" + origContent;
          Drive.Replies.patch({'content':newContent}, id, newComment.commentId, newReply.replyId)
        }
      }
    }
  }
}

