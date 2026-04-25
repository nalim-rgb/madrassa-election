// Google Apps Script Web App for Madrassa Election

const SHEET_NAME_VOTES = "Votes";

function setJsonOutput(output) {
    return ContentService.createTextOutput(JSON.stringify(output))
        .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheetSystem() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let votesSheet = ss.getSheetByName(SHEET_NAME_VOTES);
  if (!votesSheet) {
    votesSheet = ss.insertSheet(SHEET_NAME_VOTES);
    votesSheet.appendRow(["Timestamp", "Booth", "Position", "Candidate", "SessionToken", "VoterID"]);
    votesSheet.getRange("A1:F1").setFontWeight("bold");
    votesSheet.setFrozenRows(1);
  }
  return votesSheet;
}

function doGet(e) {
  let action = e.parameter.action;
  
  if (action === 'poll') {
    let booth = e.parameter.booth;
    let cache = CacheService.getScriptCache();
    let props = PropertiesService.getScriptProperties();
    let token = cache.get("activeSession_" + booth);
    let completedRaw = cache.get("completedPositions_" + booth);
    let completed = completedRaw ? JSON.parse(completedRaw) : [];
    let voterId = token ? cache.get("session_voter_" + token) : null;
    let totalCount = parseInt(props.getProperty(booth + "CompletedCount") || "0", 10);
    
    return setJsonOutput({
        status: "success", 
        activeToken: token || null,
        voterId: voterId || "UNKNOWN",
        completedPositions: completed,
        totalCount: totalCount
    });
  }
  
  if (action === 'results') {
    let sheet = getOrCreateSheetSystem();
    let data = sheet.getDataRange().getValues();
    let rows = data.slice(1);
    
    let results = {};
    let totals = { boys: 0, girls: 0, combined: 0 };
    let tokens = { boys: {}, girls: {} };
    let latestVote = null;
    let uniqueVoters = {};
    
    rows.forEach(row => {
        let ts = row[0];
        let type = row[1];
        let position = row[2];
        let candidate = row[3];
        let token = row[4];
        let voterId = row[5] || "UNKNOWN";
        
        if(ts && ts instanceof Date) {
            if(!latestVote || ts > latestVote) latestVote = ts;
        }

        if(type) {
            tokens[type][token] = true;
        }
        
        if (!results[position]) results[position] = { combined: {}, boys: {}, girls: {} };
        if (!results[position].combined[candidate]) results[position].combined[candidate] = 0;
        if (!results[position][type]) results[position][type] = {};
        if (!results[position][type][candidate]) results[position][type][candidate] = 0;
        
        results[position].combined[candidate]++;
        results[position][type][candidate]++;
        
        if (!uniqueVoters[voterId]) uniqueVoters[voterId] = { id: voterId, booth: type, ts: ts, count: 0 };
        uniqueVoters[voterId].count++;
    });
    
    totals.boys = Object.keys(tokens.boys).length;
    totals.girls = Object.keys(tokens.girls).length;
    totals.combined = totals.boys + totals.girls;
    
    // Convert unique voters object to array and sort by latest first
    let votersList = Object.values(uniqueVoters).sort((a,b) => new Date(b.ts) - new Date(a.ts));

    return setJsonOutput({
       status: "success",
       results: results,
       totals: totals,
       votersList: votersList,
       rawRows: rows.length,
       latestVote: latestVote ? latestVote.toISOString() : null
    });
  }

  return setJsonOutput({status: "error", message: "Unknown GET action or missing action parameter."});
}

function doPost(e) {
  try {
    let bodyData = JSON.parse(e.postData.contents);
    let action = bodyData.action;
    
    if (action === 'startSession') {
       let booth = bodyData.booth;
       let cache = CacheService.getScriptCache();
       let props = PropertiesService.getScriptProperties();
       let token = "TOKEN-" + Utilities.getUuid();
       
       let lock = LockService.getScriptLock();
       let voterId = "UNKNOWN";
       try {
           lock.waitLock(5000);
           let idKey = booth + "IdCount";
           let currentIdCount = parseInt(props.getProperty(idKey) || "0", 10);
           let newIdCount = currentIdCount + 1;
           props.setProperty(idKey, newIdCount.toString());
           voterId = booth.charAt(0).toUpperCase() + "-" + newIdCount.toString().padStart(3, '0');
       } catch (e) {
           voterId = booth.charAt(0).toUpperCase() + "-" + Math.floor(Math.random()*900+100);
       } finally {
           lock.releaseLock();
       }

       cache.put("activeSession_" + booth, token, 3600);
       cache.put("completedPositions_" + booth, JSON.stringify([]), 3600);
       cache.put("session_voter_" + token, voterId, 3600);
       
       return setJsonOutput({status: "success", token: token, voterId: voterId});
    }
    
    if (action === 'vote') {
        let { booth, position, candidate, sessionToken } = bodyData;
        let cache = CacheService.getScriptCache();
        let voterId = cache.get("session_voter_" + sessionToken) || "UNKNOWN";
        
        let completedRaw = cache.get("completedPositions_" + booth);
        let completed = completedRaw ? JSON.parse(completedRaw) : [];
        
        if (completed.includes(position)) {
            return setJsonOutput({status: "error", message: "Already voted for this position"});
        }
        
        let sheet = getOrCreateSheetSystem();
        sheet.appendRow([new Date(), booth, position, candidate, sessionToken, voterId]);
        
        completed.push(position);
        cache.put("completedPositions_" + booth, JSON.stringify(completed), 3600);
        
        return setJsonOutput({status: "success"});
    }
    
    if (action === 'endSession') {
        let booth = bodyData.booth;
        let props = PropertiesService.getScriptProperties();
        let lock = LockService.getScriptLock();
        if(lock.tryLock(5000)) {
            let countKey = booth + "CompletedCount";
            let c = parseInt(props.getProperty(countKey) || "0", 10);
            props.setProperty(countKey, (c+1).toString());
            lock.releaseLock();
        }

        let cache = CacheService.getScriptCache();
        cache.remove("activeSession_" + booth);
        cache.remove("completedPositions_" + booth);
        
        return setJsonOutput({status: "success"});
    }
    
    if (action === 'killToken') {
        let cache = CacheService.getScriptCache();
        cache.remove("activeSession_" + bodyData.booth);
        cache.remove("completedPositions_" + bodyData.booth);
        return setJsonOutput({status: "success"});
    }
    
    if (action === 'deleteVoter') {
        let voterId = bodyData.voterId;
        let sheet = getOrCreateSheetSystem();
        let data = sheet.getDataRange().getValues();
        let deletedRows = 0;
        
        // delete from bottom to top to avoid shifting indexes incorrectly
        for (let i = data.length - 1; i > 0; i--) {
             if (data[i][5] === voterId) { // column F is VoterID
                 sheet.deleteRow(i + 1);
                 deletedRows++;
             }
        }
        return setJsonOutput({status: "success", deletedRows: deletedRows});
    }

    if (action === 'resetVotes') {
       let sheet = getOrCreateSheetSystem();
       let lastRow = Math.max(sheet.getLastRow(), 1);
       if (lastRow > 1) {
           sheet.deleteRows(2, lastRow - 1);
       }
       let props = PropertiesService.getScriptProperties();
       props.deleteProperty("boysIdCount");
       props.deleteProperty("girlsIdCount");
       props.deleteProperty("boysCompletedCount");
       props.deleteProperty("girlsCompletedCount");
       return setJsonOutput({status: "success"});
    }

    return setJsonOutput({status: "error", message: "Unknown action"});
    
  } catch (error) {
    return setJsonOutput({status: "error", message: error.toString()});
  }
}
