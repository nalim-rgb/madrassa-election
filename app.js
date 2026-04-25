// Shared application code

const bc = new BroadcastChannel('madrassa_election');
let currentToken = null;
let currentVoterId = null;

// Free celebratory sound effect
const SUCCESS_SOUND = 'https://assets.mixkit.co/active_storage/sfx/2013/2013-preview.mp3'; 

// ============================================
// OFFLINE MOCK DATABASE LOGIC
// ============================================
function getOfflineVotes() {
    return JSON.parse(localStorage.getItem('offlineVotes') || "[]");
}
function saveOfflineVotes(votes) {
    localStorage.setItem('offlineVotes', JSON.stringify(votes));
}

async function backendRequest(action, payload = {}) {
    if(!BACKEND_URL || BACKEND_URL.trim() === "") {
        console.warn("Backend URL not configured, simulating request locally.");
        
        if (action === 'startSession') {
            let booth = payload.booth;
            let token = "TOKEN-" + Math.random().toString(36).substr(2, 9);
            let voterId;
            // Offline mock ID generation since backend is not connected
            let idKey = 'offlineIdCount_' + booth;
            let currentId = parseInt(localStorage.getItem(idKey) || "0", 10) + 1;
            localStorage.setItem(idKey, currentId);
            voterId = booth.charAt(0).toUpperCase() + "-" + currentId.toString().padStart(3, '0');
            return {status: "success", token: token, voterId: voterId};
        }
        
        if (action === 'vote') {
            let votes = getOfflineVotes();
            votes.push({
                ts: new Date().toISOString(),
                booth: payload.booth,
                position: payload.position,
                candidate: payload.candidate,
                sessionToken: payload.sessionToken,
                voterId: localStorage.getItem('activeVoterId_' + payload.booth) || "UNKNOWN"
            });
            saveOfflineVotes(votes);
            return {status: "success"};
        }

        if (action === 'endSession') {
            let booth = payload.booth;
            let countKey = 'offlineCompletedCount_' + booth;
            let current = parseInt(localStorage.getItem(countKey) || "0", 10) + 1;
            localStorage.setItem(countKey, current);
            return {status: "success"};
        }

        if (action === 'deleteVoter') {
            let votes = getOfflineVotes();
            let initialLength = votes.length;
            votes = votes.filter(v => v.voterId !== payload.voterId);
            saveOfflineVotes(votes);
            return {status: "success", deletedRows: initialLength - votes.length};
        }
        
        if (action === 'killToken') {
            localStorage.removeItem('activeSession_' + payload.booth);
            return {status: "success"};
        }

        if (action === 'resetVotes') {
            localStorage.removeItem('offlineVotes');
            localStorage.removeItem('offlineCompletedCount_boys');
            localStorage.removeItem('offlineCompletedCount_girls');
            localStorage.removeItem('offlineIdCount_boys');
            localStorage.removeItem('offlineIdCount_girls');
            return {status: "success"};
        }

        return {status: "success", mock: true};
    }
    
    // Live Backend Request
    try {
        payload.action = action;
        const response = await fetch(BACKEND_URL, {
            method: 'POST',
            mode: 'cors',
            body: JSON.stringify(payload),
            headers: {'Content-Type': 'text/plain;charset=utf-8'}
        });
        return await response.json();
    } catch (e) {
        console.error("Backend Error:", e);
        return {status: "error", message: e.message};
    }
}

async function backendPollOrGet(action, params={}) {
    if(!BACKEND_URL || BACKEND_URL.trim() === "") {
        if (action === 'poll' && params.booth) {
             let countKey = 'offlineCompletedCount_' + params.booth;
             let current = parseInt(localStorage.getItem(countKey) || "0", 10);
             let votes = getOfflineVotes();
             let activeToken = localStorage.getItem('activeSession_'+params.booth);
             let completedPositions = [];
             if (activeToken) {
                 completedPositions = votes.filter(v => v.sessionToken === activeToken).map(v => v.position);
             }
             return {
                 status: "success", 
                 activeToken: activeToken,
                 totalCount: current,
                 completedPositions: completedPositions
             };
        }
        if (action === 'results') {
            let votes = getOfflineVotes();
            
            let results = {};
            let totals = { boys: 0, girls: 0, combined: 0 };
            let tokens = { boys: {}, girls: {} };
            let latestVote = null;
            let uniqueVoters = {};
            
            POSITIONS.forEach(p => {
                results[p.title] = { combined: {}, boys: {}, girls: {} };
                p.candidates.forEach(c => {
                    results[p.title].combined[c.name] = 0;
                    results[p.title].boys[c.name] = 0;
                    results[p.title].girls[c.name] = 0;
                });
            });

            votes.forEach(row => {
                let ts = row.ts;
                let type = row.booth;
                let position = row.position;
                let candidate = row.candidate;
                let token = row.sessionToken;
                let voterId = row.voterId;
                
                if(ts) {
                    let d = new Date(ts);
                    if(!latestVote || d > new Date(latestVote)) latestVote = ts;
                }

                if(type && !tokens[type][token]) tokens[type][token] = true;
                
                if (results[position]) {
                    if (!results[position].combined[candidate]) results[position].combined[candidate] = 0;
                    if (!results[position][type]) results[position][type] = {};
                    if (!results[position][type][candidate]) results[position][type][candidate] = 0;
                    
                    results[position].combined[candidate]++;
                    results[position][type][candidate]++;
                }
                
                if (!uniqueVoters[voterId]) uniqueVoters[voterId] = { id: voterId, booth: type, ts: ts, count: 0 };
                uniqueVoters[voterId].count++;
            });
            
            totals.boys = Object.keys(tokens.boys).length;
            totals.girls = Object.keys(tokens.girls).length;
            totals.combined = totals.boys + totals.girls;
            
            let votersList = Object.values(uniqueVoters).sort((a,b) => new Date(b.ts) - new Date(a.ts));

            return {
               status: "success",
               results: results,
               totals: totals,
               votersList: votersList,
               rawRows: votes.length,
               latestVote: latestVote
            };
        }
        return {status: "success", activeToken: localStorage.getItem('activeSession_'+params.booth)};
    }
    
    // Live Backend GET
    try {
        let url = BACKEND_URL + '?action=' + encodeURIComponent(action);
        for(let k in params) url += '&' + encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
        url += '&_t=' + new Date().getTime(); // Prevent aggressive browser caching
        const response = await fetch(url);
        return await response.json();
    } catch (e) {
        return {status: "error"};
    }
}

// ============================================
// VOTER LOGIC
// ============================================
async function initVoter(booth) {
    let oldToken = localStorage.getItem('activeSession_' + booth);
    let oldVid = localStorage.getItem('activeVoterId_' + booth);

    // Clear stale sessions when the voter kiosk is freshly loaded/reloaded
    localStorage.removeItem('activeSession_' + booth);
    localStorage.removeItem('activeVoterId_' + booth);
    currentToken = null;
    currentVoterId = null;

    if (oldToken) {
        // A page reload happened mid-vote! Destroy the session in the backend.
        await backendRequest('killToken', {booth: booth});
        if (oldVid) {
            await backendRequest('deleteVoter', {voterId: oldVid});
        }
    }

    let positionsForBooth = POSITIONS.map(p => ({
        ...p,
        candidates: p.candidates.filter(c => c.booth === booth || c.booth === 'both')
    })).filter(p => p.candidates.length > 0);

    let currentStep = 0;
    let selectedVotes = {};
    let isSubmitting = false;
    let isShowingSuccess = false;  // prevents polls from hijacking the success screen
    let successTimer = null;       // tracked so new voter can cancel the previous timer

    // UI Elements
    const screenLocked = document.getElementById('screen-locked');
    const screenVoting = document.getElementById('screen-voting');
    const screenSuccess = document.getElementById('screen-success');

    function showScreen(screen) {
        [screenLocked, screenVoting, screenSuccess].forEach(s => s?.classList.add('hidden'));
        screen?.classList.remove('hidden');
    }

    async function checkActiveSession() {
        if (isSubmitting) return;
        // Don't interfere while showing the success screen
        if (isShowingSuccess) return;
        
        if (BACKEND_URL && BACKEND_URL.trim() !== "") {
            backendPollOrGet('poll', {booth: booth}).then(res => {
                if(res.status === 'success') {
                    if (res.activeToken && res.activeToken !== currentToken) {
                        startVoting(res.activeToken, res.voterId);
                    } else if (!res.activeToken && currentToken) {
                        // Backend says token is dead — abort
                        currentToken = null;
                        currentVoterId = null;
                        showScreen(screenLocked);
                    }
                }
            });
        }
    }

    setInterval(checkActiveSession, 3000);

    bc.onmessage = (e) => {
        if(e.data.type === 'startSession' && e.data.booth === booth) {
            startVoting(e.data.token, e.data.voterId);
        }
    };

    // Helper: extract voter ID embedded in the token string as a reliable fallback
    function extractVoterIdFromToken(token) {
        if (token && token.includes('::VOTER::')) {
            return token.split('::VOTER::')[0];
        }
        return null;
    }

    function startVoting(token, voterId) {
        // Cancel any pending success-screen timer from the PREVIOUS voter!
        // Without this, the previous 10s timer would kick the new voter off mid-vote.
        if (successTimer) {
            clearTimeout(successTimer);
            successTimer = null;
        }
        isShowingSuccess = false;

        currentToken = token;
        // Use the voter ID from the token string if the cache-based one is missing
        currentVoterId = (voterId && voterId !== 'UNKNOWN') ? voterId : (extractVoterIdFromToken(token) || 'UNKNOWN');
        localStorage.setItem('activeSession_' + booth, token); 
        localStorage.setItem('activeVoterId_' + booth, currentVoterId);
        currentStep = 0;
        selectedVotes = {};
        isSubmitting = false;
        renderPosition();
        showScreen(screenVoting);
    }

    function getIconHTML(iconPath) {
        if (iconPath.includes('/') || iconPath.includes('.')) {
            return `<img src="${iconPath}" alt="Icon" class="candidate-icon image-icon">`;
        }
        return `<div class="candidate-icon">${iconPath}</div>`;
    }

    function renderPosition() {
        if (currentStep >= positionsForBooth.length) {
            finishVoting();
            return;
        }

        const position = positionsForBooth[currentStep];
        document.getElementById('pos-title').innerText = position.title;
        document.getElementById('progress-fill').style.width = ((currentStep / positionsForBooth.length) * 100) + '%';
        
        let candidatesHTML = '';
        let candidates = position.candidates;
        
        candidates.forEach(c => {
            const isSelected = selectedVotes[position.title] === c.name ? 'selected' : '';
            candidatesHTML += `
                <div class="candidate-card ${isSelected}" onclick="window.selectCandidate('${c.name}')">
                    ${getIconHTML(c.icon)}
                    <div class="candidate-name">${c.name}</div>
                </div>
            `;
        });
        
        document.getElementById('candidates-container').innerHTML = candidatesHTML;
        const btnSubmit = document.getElementById('btn-submit-vote');
        btnSubmit.disabled = !selectedVotes[position.title];
        btnSubmit.innerHTML = 'Submit Vote';
    }

    window.selectCandidate = function(name) {
        selectedVotes[positionsForBooth[currentStep].title] = name;
        renderPosition(); 
    }

    window.submitPositionVote = async function() {
        if(isSubmitting) return;
        isSubmitting = true;
        const btn = document.getElementById('btn-submit-vote');
        btn.innerHTML = '<span class="loader"></span> Saving...';
        btn.disabled = true;

        let positionTitle = positionsForBooth[currentStep].title;
        let candidateName = selectedVotes[positionTitle];

        let res = await backendRequest('vote', {
            booth: booth,
            position: positionTitle,
            candidate: candidateName,
            sessionToken: currentToken
        });
        
        isSubmitting = false;

        if (res && res.status === 'success') {
            bc.postMessage({type: 'voted', booth: booth, position: positionTitle});
            currentStep++;
            renderPosition();
        } else {
            alert("Network Error: Could not save your vote! Please try again or call the Officer.");
            btn.innerHTML = 'Submit Vote';
            btn.disabled = false;
        }
    }

    async function finishVoting() {
        // Snapshot the voter ID NOW before any async calls can overwrite currentVoterId
        const completedVoterId = currentVoterId;

        localStorage.removeItem('activeSession_' + booth);
        localStorage.removeItem('activeVoterId_' + booth);
        
        bc.postMessage({type: 'success', booth: booth});
        
        await backendRequest('endSession', {booth: booth, sessionToken: currentToken});
        
        let audio = new Audio(SUCCESS_SOUND);
        audio.play().catch(e => console.log("Audio play blocked by browser", e));

        // Use the snapshotted ID — guaranteed to be correct even if currentVoterId changed
        document.getElementById('success-voter-id').innerText = completedVoterId || 'N/A';
        showScreen(screenSuccess);
        
        // Clear token immediately so polls don't mistake the finished session for an abort
        currentToken = null;
        isShowingSuccess = true;

        successTimer = setTimeout(() => {
            isShowingSuccess = false;
            successTimer = null;
            bc.postMessage({type: 'reset', booth: booth});
            currentVoterId = null;
            showScreen(screenLocked);
        }, 10000);
    }
}


// ============================================
// OFFICER LOGIC
// ============================================
function initOfficer(booth) {
    const errorMsg = document.getElementById('error-msg');
    const startBtn = document.getElementById('btn-start');
    const counterDisplay = document.getElementById('voters-count');
    
    let isSessionActive = false;
    let isVotingInProgress = false;
    let currentSessionVoterId = null;
    let lastKnownCompletedCount = 0;

    let positionsForBooth = POSITIONS.map(p => ({
        ...p,
        candidates: p.candidates.filter(c => c.booth === booth || c.booth === 'both')
    })).filter(p => p.candidates.length > 0);

    const statusGrid = document.getElementById('status-grid');
    statusGrid.innerHTML = positionsForBooth.map(p => `
        <div class="status-box" id="status-${p.title.replace(/\s+/g, '-')}">
            ${p.title}
        </div>
    `).join('');

    function resetGrid() {
        document.querySelectorAll('.status-box').forEach(el => el.classList.remove('active'));
    }

    function pingGreen(positionTitle) {
        const id = 'status-' + positionTitle.replace(/\s+/g, '-');
        const box = document.getElementById(id);
        if (box && !box.classList.contains('active')) {
            box.classList.add('active');
            let audio = new Audio('https://assets.mixkit.co/active_storage/sfx/2869/2869-preview.mp3');
            audio.volume = 0.5;
            audio.play().catch(e => e);
        }
    }

    function setOfficerState(active) {
        isSessionActive = active;
        if(active) {
            startBtn.disabled = true;
            startBtn.innerText = "Voting is progressing...";
        } else {
            startBtn.disabled = false;
            startBtn.innerText = "Start New Voter";
        }
    }

    bc.onmessage = (e) => {
        if(e.data.type === 'voted' && e.data.booth === booth) {
            pingGreen(e.data.position);
        } else if (e.data.type === 'success' && e.data.booth === booth) {
            isVotingInProgress = false;
        } else if (e.data.type === 'reset' && e.data.booth === booth) {
            resetGrid();
            setOfficerState(false);
            isVotingInProgress = false;
        } else if (e.data.type === 'startSession' && e.data.booth === booth) {
            setOfficerState(true);
            isVotingInProgress = true;
            currentSessionVoterId = e.data.voterId;
        }
    };

    setInterval(async () => {
        let res = await backendPollOrGet('poll', {booth: booth});
        if(res && res.status === 'success') {
            let newCount = res.totalCount !== undefined ? res.totalCount : lastKnownCompletedCount;
            let countIncreased = newCount > lastKnownCompletedCount;

            // STEP 1: If count went UP, session completed successfully — mark done BEFORE checking token
            if (countIncreased) {
                isVotingInProgress = false;
            }
            lastKnownCompletedCount = newCount; // Always sync (catches resets too)

            if (counterDisplay) counterDisplay.innerText = newCount;

            // STEP 2: Only AFTER updating isVotingInProgress, evaluate the token state
            if (!res.activeToken) {
                if (isVotingInProgress) {
                    // TRUE abort: count did NOT go up AND session vanished
                    isVotingInProgress = false;
                    resetGrid();
                    setOfficerState(false);

                    errorMsg.style.display = 'block';
                    errorMsg.className = 'message error';
                    errorMsg.style.backgroundColor = '#ef4444';
                    errorMsg.style.color = 'white';
                    errorMsg.innerText = `Session aborted! Erasing incomplete votes for ${currentSessionVoterId}...`;

                    if (currentSessionVoterId) {
                        backendRequest('deleteVoter', {voterId: currentSessionVoterId}).then(() => {
                            errorMsg.innerText = `Session aborted. Votes for ${currentSessionVoterId} erased.`;
                            setTimeout(() => {
                                errorMsg.style.display = 'none';
                                errorMsg.style.backgroundColor = '';
                                errorMsg.style.color = '';
                            }, 6000);
                        });
                    }
                } else {
                    // Session finished cleanly or no session was running
                    resetGrid();
                    setOfficerState(false);
                }
            } else {
                setOfficerState(true);
                if (res.completedPositions) {
                    res.completedPositions.forEach(pos => pingGreen(pos));
                }
            }
        }
    }, 4000);

    window.startNewVoter = async function() {
        startBtn.disabled = true;
        startBtn.innerHTML = '<span class="loader"></span> Authorizing...';
        resetGrid();
        
        // Fetch the absolute latest baseline before authorizing
        let pollRes = await backendPollOrGet('poll', {booth: booth});
        if (pollRes && pollRes.totalCount !== undefined) {
            lastKnownCompletedCount = pollRes.totalCount;
        }

        let res = await backendRequest('startSession', {booth: booth});
        
        // Auto-retry if Google Apps Script blocks the concurrent request
        if (res.status === 'error' || !res.voterId || res.voterId === "UNKNOWN") {
            await new Promise(r => setTimeout(r, 1500)); // wait 1.5s to let the other booth finish
            res = await backendRequest('startSession', {booth: booth});
        }

        if (res.status === 'error' || !res.voterId || res.voterId === "UNKNOWN") {
            alert("Network Error: Google Servers are momentarily busy. Please click Start New Voter again.");
            startBtn.disabled = false;
            startBtn.innerHTML = 'Start New Voter';
            return;
        }

        let token = res.token;
        let voterId = res.voterId;

        localStorage.setItem('activeSession_' + booth, token);
        localStorage.setItem('activeVoterId_' + booth, voterId);
        bc.postMessage({type: 'startSession', booth: booth, token: token, voterId: voterId});

        let audio = new Audio('https://assets.mixkit.co/active_storage/sfx/2870/2870-preview.mp3'); 
        audio.play().catch(e=>e);

        setOfficerState(true);
        isVotingInProgress = true;
        currentSessionVoterId = voterId;
        
        errorMsg.style.display = 'block';
        errorMsg.className = 'message success';
        errorMsg.innerText = 'Ballot activated for ' + voterId + '!';
        
        setTimeout(() => {
            errorMsg.style.display = 'none';
        }, 6000);
    }
}

// ============================================
// ADMIN LOGIC
// ============================================
function initAdmin() {
    const loginScreen = document.getElementById('admin-login');
    const dbScreen = document.getElementById('admin-dashboard');
    let currentAdminData = null;

    window.checkPassword = function() {
        const pw = document.getElementById('admin-pw').value.trim();
        if(pw === 'madrassa2025admin') {
            loginScreen.classList.add('hidden');
            dbScreen.classList.remove('hidden');
            loadResults();
            setInterval(loadResults, 10000);
        } else {
            alert("Incorrect Password!");
        }
    }
    
    let pwInput = document.getElementById('admin-pw');
    if (pwInput) {
        pwInput.addEventListener('keypress', function (e) {
            if (e.key === 'Enter') {
                checkPassword();
            }
        });
    }

    async function loadResults() {
        const res = await backendPollOrGet('results');
        if(res && res.status === 'success') {
            
            // Ensure all positions and candidates exist (even if they have 0 votes)
            POSITIONS.forEach(p => {
                if (!res.results[p.title]) res.results[p.title] = { combined: {}, boys: {}, girls: {} };
                p.candidates.forEach(c => {
                    if (res.results[p.title].combined[c.name] === undefined) res.results[p.title].combined[c.name] = 0;
                    if (res.results[p.title].boys[c.name] === undefined) res.results[p.title].boys[c.name] = 0;
                    if (res.results[p.title].girls[c.name] === undefined) res.results[p.title].girls[c.name] = 0;
                });
            });

            currentAdminData = res;
            renderAdmin(res);
        }
    }

    function renderAdmin(data) {
        document.getElementById('total-votes').innerText = data.totals.combined;
        // Total Voters
        document.getElementById('boys-votes').innerText = data.totals.boys;
        document.getElementById('girls-votes').innerText = data.totals.girls;
        if(data.latestVote) {
            document.getElementById('last-vote-time').innerText = "Last Vote Cast: " + new Date(data.latestVote).toLocaleString();
        } else {
            document.getElementById('last-vote-time').innerText = "No votes recorded yet.";
        }
        
        // Removed dynamic chart rendering from main dashboard per user request

        // Render Voters Table
        let tbody = document.getElementById('voters-tbody');
        if (tbody && data.votersList) {
            tbody.innerHTML = data.votersList.map(v => `
                <tr>
                    <td><strong>${v.id}</strong></td>
                    <td>${v.booth.toUpperCase()}</td>
                    <td>${new Date(v.ts).toLocaleTimeString()}</td>
                    <td>
                        <button class="btn btn-danger" style="padding: 0.5rem 1rem; width: auto; font-size: 0.9rem;" onclick="deleteVoter('${v.id}')">Delete</button>
                    </td>
                </tr>
            `).join('');
        }
    }

    // Modal Counting Logic
    window.openCounter = function(booth) {
        if (!currentAdminData) return alert("Waiting for data...");
        const modal = document.getElementById('counting-modal');
        const posDiv = document.getElementById('modal-positions');
        const resDiv = document.getElementById('modal-results');
        
        document.getElementById('modal-title').innerText = "Count " + booth.toUpperCase() + (booth === 'combined' ? "" : " Booth");
        
        posDiv.innerHTML = '';
        resDiv.innerHTML = '';
        resDiv.classList.add('hidden');
        posDiv.classList.remove('hidden');

        POSITIONS.forEach(p => {
            let btn = document.createElement('button');
            btn.className = 'btn btn-secondary';
            btn.innerText = p.title;
            btn.onclick = () => showBoothResults(booth, p.title);
            posDiv.appendChild(btn);
        });

        modal.classList.remove('hidden');
    }

    window.closeCounter = function() {
        document.getElementById('counting-modal').classList.add('hidden');
    }

    window.showBoothResults = function(booth, position) {
        const posDiv = document.getElementById('modal-positions');
        const resDiv = document.getElementById('modal-results');
        posDiv.classList.add('hidden');
        resDiv.classList.remove('hidden');

        let votesForPos = currentAdminData.results[position][booth] || {};
        
        // Sort candidates by highest votes
        let sortedCandidates = Object.keys(votesForPos).sort((a, b) => votesForPos[b] - votesForPos[a]);
        
        let html = `<h3 style="margin-bottom:1.5rem; text-align:center;">${position} (${booth.toUpperCase()})</h3>`;
        html += `<div style="display:flex; flex-direction:column; gap:1rem;">`;
        
        if (sortedCandidates.length === 0) {
            html += `<p style="text-align:center;">No votes recorded yet.</p>`;
        }

        sortedCandidates.forEach((candidate, index) => {
            let v = votesForPos[candidate] || 0;
            let trophy = index === 0 && v > 0 ? '🏆 ' : '';
            let medalColor = index === 0 ? 'var(--primary-color)' : 'var(--text-secondary)';
            
            html += `
                <div style="background:#f8fafc; padding:1.5rem; border-radius:8px; display:flex; justify-content:space-between; align-items:center; border-left: 4px solid ${medalColor};">
                    <span style="font-size:1.2rem; font-weight:800;">${trophy}${candidate}</span>
                    <span style="font-size:1.5rem; font-weight:800; color:${medalColor};">${v} Votes</span>
                </div>
            `;
        });
        
        html += `</div>`;
        html += `<button class="btn btn-secondary" style="margin-top:2rem; width:100%;" onclick="openCounter('${booth}')">Back to Positions</button>`;
        resDiv.innerHTML = html;
    }

    window.deleteVoter = async function(voterId) {
        let conf = prompt(`Are you sure you want to delete ${voterId}? This will erase all their votes instantly. Type YES to confirm:`);
        if(conf && conf.trim().toUpperCase() === "YES") {
            const res = await backendRequest('deleteVoter', {voterId: voterId});
            if(res.status === 'success') {
                alert(`Successfully deleted ${res.deletedRows} records for ${voterId}.`);
                loadResults(); // reload instantly
            } else {
                alert("Error deleting voter: " + res.message);
            }
        }
    }

    window.exportCSV = function() {
        if(!BACKEND_URL || BACKEND_URL.trim() === "") {
            let votes = getOfflineVotes();
            let csv = "Timestamp,Booth,Position,Candidate,SessionToken,VoterID\n";
            votes.forEach(v => {
                csv += `${v.ts},${v.booth},${v.position},${v.candidate},${v.sessionToken},${v.voterId}\n`;
            });
            let blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
            let link = document.createElement("a");
            if (link.download !== undefined) {
                let url = URL.createObjectURL(blob);
                link.setAttribute("href", url);
                link.setAttribute("download", "madrassa_offline_votes.csv");
                link.style.visibility = 'hidden';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            }
        } else {
            alert("The easiest way to export is directly descending from Google Sheets! If needed urgently, just copy from Google Sheets File -> Download -> CSV.");
        }
    };

    window.resetAllVotes = async function() {
        let conf = prompt("Type RESET exactly to confirm deleting ALL votes:");
        if(conf && conf.trim().toUpperCase() === "RESET") {
            await backendRequest('resetVotes');
            alert("All Votes successfully erased.");
            loadResults();
        }
    };
}
