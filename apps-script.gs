// Google Apps Script Web App for Madrassa Election
// VERSION: 6 - Final Definitive Fix

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

// Helper: extract voter ID from token string (format: "B-001::VOTER::uuid")
function extractVoterIdFromToken(token) {
    if (token && token.indexOf("::VOTER::") !== -1) {
        return token.split("::VOTER::")[0];
    }
    return null;
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
    let totalCount = parseInt(props.getProperty(booth + "CompletedCount") || "0", 10);
    
    // Extract voter ID from the token string itself (guaranteed, never UNKNOWN)
    let voterId = "UNKNOWN";
    if (token) {
        voterId = extractVoterIdFromToken(token) || cache.get("sessionVoter_" + booth) || "UNKNOWN";
    }
    
    return setJsonOutput({
        status: "success", 
        activeToken: token || null,
        voterId: voterId,
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
    
    rows.forEach(function(row) {
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
    
    let votersList = Object.values(uniqueVoters).sort(function(a,b) { return new Date(b.ts) - new Date(a.ts); });

    return setJsonOutput({
       status: "success",
       results: results,
       totals: totals,
       votersList: votersList,
       rawRows: rows.length,
       latestVote: latestVote ? latestVote.toISOString() : null
    });
  }

  return setJsonOutput({status: "error", message: "Unknown GET action."});
}

function doPost(e) {
  try {
    let bodyData = JSON.parse(e.postData.contents);
    let action = bodyData.action;
    
    // ── START SESSION ──────────────────────────────────────────────────────────
    if (action === 'startSession') {
       let booth = bodyData.booth;
       let cache = CacheService.getScriptCache();
       let props = PropertiesService.getScriptProperties();
       
       // USER lock: boys and girls can run their ID counters simultaneously
       let lock = LockService.getUserLock();
       let voterId = "UNKNOWN";
       try {
           lock.waitLock(8000);
           let idKey = booth + "IdCount";
           let currentIdCount = parseInt(props.getProperty(idKey) || "0", 10);
           let newIdCount = currentIdCount + 1;
           props.setProperty(idKey, newIdCount.toString());
           voterId = booth.charAt(0).toUpperCase() + "-" + newIdCount.toString().padStart(3, '0');
       } catch (lockErr) {
           voterId = booth.charAt(0).toUpperCase() + "-???";
       } finally {
           try { lock.releaseLock(); } catch(ex) {}
       }

       // CRITICAL: Voter ID is EMBEDDED in the token string.
       // This means it can NEVER be "UNKNOWN" even if cache is evicted under load.
       let token = voterId + "::VOTER::" + Utilities.getUuid();

       cache.put("activeSession_" + booth, token, 7200);
       cache.put("completedPositions_" + booth, JSON.stringify([]), 7200);
       cache.put("sessionVoter_" + booth, voterId, 7200);
       
       return setJsonOutput({status: "success", token: token, voterId: voterId});
    }
    
    // ── VOTE ───────────────────────────────────────────────────────────────────
    if (action === 'vote') {
        let booth = bodyData.booth;
        let position = bodyData.position;
        let candidate = bodyData.candidate;
        let sessionToken = bodyData.sessionToken;
        let cache = CacheService.getScriptCache();
        
        // Verify this vote belongs to the currently active session
        let activeToken = cache.get("activeSession_" + booth);
        if (!activeToken || activeToken !== sessionToken) {
            return setJsonOutput({status: "error", message: "Session token mismatch or expired."});
        }
        
        // Extract voter ID from the token (never UNKNOWN)
        let voterId = extractVoterIdFromToken(sessionToken) || cache.get("sessionVoter_" + booth) || "UNKNOWN";
        
        let completedRaw = cache.get("completedPositions_" + booth);
        let completed = completedRaw ? JSON.parse(completedRaw) : [];
        
        if (completed.indexOf(position) !== -1) {
            return setJsonOutput({status: "error", message: "Already voted for this position"});
        }
        
        let sheet = getOrCreateSheetSystem();
        sheet.appendRow([new Date(), booth, position, candidate, sessionToken, voterId]);
        
        completed.push(position);
        cache.put("completedPositions_" + booth, JSON.stringify(completed), 7200);
        
        return setJsonOutput({status: "success"});
    }
    
    // ── END SESSION ────────────────────────────────────────────────────────────
    if (action === 'endSession') {
        let booth = bodyData.booth;
        let sessionToken = bodyData.sessionToken;
        let cache = CacheService.getScriptCache();
        
        // CRITICAL RACE CONDITION FIX:
        // Only clear the session if the token being ended is STILL the active one.
        // This prevents a finishing voter from destroying a brand-new session that
        // the officer started in the same moment.
        let currentActiveToken = cache.get("activeSession_" + booth);
        
        let props = PropertiesService.getScriptProperties();
        let lock = LockService.getUserLock();
        if (lock.tryLock(5000)) {
            let countKey = booth + "CompletedCount";
            let c = parseInt(props.getProperty(countKey) || "0", 10);
            props.setProperty(countKey, (c+1).toString());
            lock.releaseLock();
        }

        if (sessionToken && currentActiveToken && currentActiveToken !== sessionToken) {
            // A NEW session already started — preserve it, only increment count
            return setJsonOutput({status: "success", note: "new_session_preserved"});
        }

        // Safe to clear — token matches (or no session token provided)
        cache.remove("activeSession_" + booth);
        cache.remove("completedPositions_" + booth);
        cache.remove("sessionVoter_" + booth);
        
        return setJsonOutput({status: "success"});
    }
    
    // ── KILL TOKEN (reload abort) ──────────────────────────────────────────────
    if (action === 'killToken') {
        let cache = CacheService.getScriptCache();
        cache.remove("activeSession_" + bodyData.booth);
        cache.remove("completedPositions_" + bodyData.booth);
        cache.remove("sessionVoter_" + bodyData.booth);
        return setJsonOutput({status: "success"});
    }
    
    // ── DELETE VOTER ───────────────────────────────────────────────────────────
    if (action === 'deleteVoter') {
        let voterId = bodyData.voterId;
        if (!voterId || voterId === "UNKNOWN") {
            return setJsonOutput({status: "success", deletedRows: 0});
        }
        let sheet = getOrCreateSheetSystem();
        let data = sheet.getDataRange().getValues();
        let deletedRows = 0;
        for (let i = data.length - 1; i > 0; i--) {
             if (data[i][5] === voterId) {
                 sheet.deleteRow(i + 1);
                 deletedRows++;
             }
        }
        return setJsonOutput({status: "success", deletedRows: deletedRows});
    }

    // ── RESET VOTES ────────────────────────────────────────────────────────────
    if (action === 'resetVotes') {
       let sheet = getOrCreateSheetSystem();
       let lastRow = Math.max(sheet.getLastRow(), 1);
       if (lastRow > 1) {
           sheet.deleteRows(2, lastRow - 1);
       }
       
       let props = PropertiesService.getScriptProperties();
       props.setProperty("boysIdCount", "0");
       props.setProperty("girlsIdCount", "0");
       props.setProperty("boysCompletedCount", "0");
       props.setProperty("girlsCompletedCount", "0");

       let cache = CacheService.getScriptCache();
       cache.remove("activeSession_boys");
       cache.remove("activeSession_girls");
       cache.remove("completedPositions_boys");
       cache.remove("completedPositions_girls");
       cache.remove("sessionVoter_boys");
       cache.remove("sessionVoter_girls");
       
       return setJsonOutput({status: "success"});
    }

    return setJsonOutput({status: "error", message: "Unknown action"});
    
  } catch (error) {
    return setJsonOutput({status: "error", message: error.toString()});
  }
}
