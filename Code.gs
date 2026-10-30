/**
 * CricketScorer - Google Apps Script receiver
 *
 * Every time anything changes in the app, it sends ALL of its data to this
 * script and the script rewrites one tab per kind of data:
 *
 *   Players        every saved player and the teams they are in
 *   Teams          saved teams with their players, captain and wicketkeeper
 *   Matches        every finished match: date, teams, toss, scores, result,
 *                  category (Tournament match / Regular match) and tournament
 *   Innings        one row per innings (the match in progress is included): total, wickets, overs, extras, fall of wickets
 *   Batting        every batter in every innings: runs, balls, 4s, 6s, strike rate, how out
 *   Bowling        every bowler in every innings: overs, maidens, runs, wickets, wides, no balls, economy
 *   Toss           the toss of every match: who called, the call, the bat flip outcome,
 *                  who won the toss and what they chose
 *   Lineups        every player in every match: team, captain, wicketkeeper, playing or substitute
 *   Ball by ball   every ball and event of every match, finished or in progress
 *   Live           ball by ball record of the match being scored
 *   Current match  the match in progress: teams, captains, keepers, toss, score
 *   Tournaments    every tournament with its status, current stage, its people and its score or result
 *   Tournament teams  each team in each tournament: who plays as it, its players and its record
 *   People         each person in a tournament: teams played as, match wins
 *   Fixtures       every fixture: stage, teams, who played as each, status, result, winning team, point to, score after the match
 *   Series         the running score between each pair of teams, e.g. 3-1
 *   Activity       dated log of what was done in each tournament
 *   App data       the whole app's data, saved so every device using this Sync URL shows the same thing (do not edit)
 *
 * Setup:
 *  1. Create a Google Sheet, then open Extensions > Apps Script.
 *  2. Replace the default code with this file and save.
 *  3. Deploy > New deployment > type "Web app".
 *     Execute as: Me.   Who has access: Anyone.
 *  4. Copy the web app URL (ends in /exec) and paste it into the app:
 *     menu (three dots) > Match history > Sync URL.
 *  5. If you change this code later, use Deploy > Manage deployments >
 *     Edit > New version so the same URL keeps working.
 *
 * To add every tab and its column headings straight away, choose the function
 * setupSheets in the Apps Script editor and press Run (once). The headings are
 * also added automatically whenever the app sends data, and whenever the web app
 * URL is opened in a browser.
 *
 * The Players and Teams tabs are TWO-WAY. Whatever is changed by hand in those two tabs is read
 * back into the app within a few seconds and shown on every device:
 *   - a new row adds a player or team (leave the Id blank; the app fills it in)
 *   - a changed name, squad, captain or wicketkeeper updates it
 *   - a deleted row removes it
 * Anything the sheet has that the app lacks is also brought into the app, including matches
 * listed in the Matches tab (their summary comes back; ball-by-ball detail only if the app had it),
 * and tournaments listed in the Tournaments tab, rebuilt from the Tournaments, Tournament teams, People,
 * Fixtures and Activity tabs (teams, who plays as whom, every fixture with its stage and result, the log).
 * Nothing is removed if the whole tab is emptied. The "Sync base" tab is how the script tells
 * your hand edits from the app's own writes (do not edit or delete it).
 * In the Teams tab, put the team name in column B and the player names, separated
 * by commas, in column C (captain in D and wicketkeeper in E are optional; column F is filled in by the app).
 *
 * Open the web app URL in a browser to check it is live. It shows when the
 * last sync arrived and how many rows each tab holds.
 */

/* Leave this empty if you made the script from the sheet itself (Extensions > Apps Script).
 * If the script was created on its own at script.google.com, paste the sheet's ID here
 * (the long text in the sheet's address between /d/ and /edit). If you leave it empty in
 * that case, the script creates a new sheet called "CricketScorer data" and keeps using it. */
var SPREADSHEET_ID = '';

var MAX_CELL = 45000; // Google Sheets allows 50,000 characters in a cell.

/* The column headings of every tab, so they are always written even when a tab has no rows yet. */
var HEADINGS = {
  "Players": ["Id", "Name", "Teams"],
  "Teams": ["Id", "Team", "Players", "Captain", "Wicketkeeper", "Player ids"],
  "Matches": ["Match ID", "Date", "Team A", "Team B", "Format", "Overs", "Toss", "Scores", "Result", "Category", "Tournament", "Toss called by", "Call", "Bat flip landed", "Toss won by", "Chose to"],
  "Innings": ["Match ID", "Date", "Match", "Innings", "Batting team", "Runs", "Wickets", "Overs", "Extras", "Fall of wickets"],
  "Batting": ["Match ID", "Date", "Match", "Innings", "Team", "Batter", "How out", "Runs", "Balls", "4s", "6s", "Strike rate"],
  "Bowling": ["Match ID", "Date", "Match", "Innings", "Team", "Bowler", "Overs", "Maidens", "Runs", "Wickets", "Wides", "No balls", "Economy"],
  "Toss": ["Match ID", "Date", "Match", "Format", "Toss type", "Called by", "Call", "Bat flip outcome (side facing up)", "Toss won by", "Chose to", "Toss summary", "Result"],
  "Lineups": ["Match ID", "Date", "Match", "Team", "Player", "Role", "Status"],
  "Ball by ball": ["Match ID", "Date", "Match", "Innings", "Event #", "Type", "Ball kind", "Runs", "Bowler", "Wicket", "Batter out", "Fielder"],
  "Live": ["Match", "Innings", "Event #", "Type", "Ball kind", "Runs", "Bowler", "Wicket", "Batter out", "Fielder", "Bowler id"],
  "Current match": ["Match", "Format", "Overs", "Toss", "Tournament", "Score", "Status", "Team A", "Captain A", "Wicketkeeper A", "Team B", "Captain B", "Wicketkeeper B", "Toss called by", "Call", "Bat flip landed", "Toss won by", "Chose to"],
  "Tournaments": ["Id", "Tournament", "Type", "Overs", "Status", "Teams", "Fixtures", "Played", "Created", "People", "Score or result", "Current stage"],
  "Tournament teams": ["Tournament", "Team", "Played by", "Players", "Played", "Won", "Lost", "Tied", "Drawn", "Tournament ID"],
  "People": ["Tournament", "Person", "Teams played as", "Match wins", "Status", "Tournament ID"],
  "Fixtures": ["Tournament", "#", "Stage", "Team A", "Team B", "Status", "Result", "Scores", "Date", "Match ID", "Phase", "Person A", "Person B", "Score after match", "Winning team", "Point to", "Tournament ID"],
  "Series": ["Tournament", "Team A", "Wins A", "Wins B", "Team B", "Ties", "Draws", "Played", "Total", "Status", "Tournament ID"],
  "Activity": ["Tournament", "When", "What happened", "Tournament ID"],
  "App data": ["Part", "Shared app data (do not edit or delete)"],
  "Sync base": ["Part", "What Players and Teams looked like after the last app write (do not edit or delete)"]
};
var TAB_ORDER = ["Players", "Teams", "Matches", "Innings", "Batting", "Bowling", "Toss", "Lineups", "Ball by ball", "Live", "Current match", "Tournaments", "Tournament teams", "People", "Fixtures", "Series", "Activity", "App data", "Sync base"];

/* ---- merging, so what two devices enter at the same time is never lost ---- */
function cmp_(a, b) { return a < b ? -1 : a > b ? 1 : 0; }
function canon_(v) {
  if (v === null || v === undefined) return 'null';
  if (Object.prototype.toString.call(v) === '[object Array]') return '[' + v.map(canon_).join(',') + ']';
  if (typeof v === 'object') {
    return '{' + Object.keys(v).filter(function (k) { return v[k] !== undefined; }).sort().map(function (k) { return JSON.stringify(k) + ':' + canon_(v[k]); }).join(',') + '}';
  }
  return JSON.stringify(v);
}
/* Puts the shared data into a fixed order, the same order the app uses, so two copies with the same data compare equal. */
function normCore_(s) {
  s = s || {};
  var lib = s.lib || {};
  var byNm = function (a, b) { return cmp_(String(a.name || '').toLowerCase(), String(b.name || '').toLowerCase()) || cmp_(String(a.id), String(b.id)); };
  return {
    liveAt: String(s.liveAt !== undefined ? s.liveAt : (s.at || '')),
    lib: { players: (lib.players || []).slice().sort(byNm), teams: (lib.teams || []).slice().sort(byNm) },
    hist: (s.hist || []).slice().sort(function (a, b) { return cmp_(String(b.at), String(a.at)) || cmp_(String(a.id), String(b.id)); }),
    tours: (s.tours || []).slice().sort(function (a, b) { return cmp_(String(a.id), String(b.id)); }),
    live: s.live || null,
    del: s.del || {}
  };
}
/* Every record is kept; when both sides have the same record the one updated last wins. A dated delete marker removes a record unless it was edited after the delete. */
function mergeRecs_(x, y, pre, del) {
  var map = {};
  [x || [], y || []].forEach(function (list) {
    list.forEach(function (r) {
      if (!r || !r.id) return;
      var c = map[r.id];
      if (!c || String(r.u || '') >= String(c.u || '')) map[r.id] = r;
    });
  });
  var out = [];
  for (var id in map) {
    var r = map[id], t = del[pre + id];
    if (t && String(t) >= String(r.u || '')) continue;
    out.push(r);
  }
  return out;
}
function mergeStates_(a, b) {
  a = normCore_(a); b = normCore_(b);
  var del = {}, k;
  [a.del, b.del].forEach(function (d) { for (k in d) { if (!del[k] || String(d[k]) > String(del[k])) del[k] = d[k]; } });
  var live = b.liveAt >= a.liveAt ? b : a;
  return normCore_({
    liveAt: live.liveAt, live: live.live, del: del,
    lib: {
      players: mergeRecs_(a.lib.players, b.lib.players, 'p:', del).filter(function (p) { return p.name; }),
      teams: mergeRecs_(a.lib.teams, b.lib.teams, 't:', del)
    },
    hist: mergeRecs_(a.hist, b.hist, 'm:', del).filter(function (r) { return r.t1; }),
    tours: mergeRecs_(a.tours, b.tours, 'o:', del)
  });
}
function loadChunks_(ss, name) {
  var rows = readTab_(ss, name);
  rows.sort(function (a, b) { return Number(a[0]) - Number(b[0]); });
  return rows.map(function (r) { return String(r[1]).substring(1); }).join('');
}
function chunkRows_(text) {
  var chunks = [];
  for (var c = 0; c < text.length; c += 30000) chunks.push([chunks.length + 1, '~' + text.substr(c, 30000)]);
  return chunks;
}
function hash_(s) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s)); }
function newId_(p) { return p + new Date().getTime().toString(36) + Math.random().toString(36).slice(2, 6); }
function nameKey_(s) { return String(s || '').toLowerCase(); }

/* Writes the merged shared data and tells every device that it changed. */
function saveState_(ss, props, state, by) {
  var newAt = new Date().toISOString();
  state.at = newAt;
  writeTab_(ss, 'App data', HEADINGS['App data'], chunkRows_(JSON.stringify(state)));
  props.setProperty('stateAt', newAt);
  props.setProperty('stateBy', by || '');
}

/* ---- Players and Teams are two-way: what is typed in those two tabs of the sheet is read back into the app ---- */
function rowsNorm_(ss, name, width) {
  return readTab_(ss, name).map(function (r) {
    var o = [];
    for (var i = 0; i < width; i++) o.push(String(r[i] === undefined ? '' : r[i]).replace(/\s+/g, ' ').trim());
    return o;
  });
}
function sheetLibRows_(ss) { return { Players: rowsNorm_(ss, 'Players', 3), Teams: rowsNorm_(ss, 'Teams', 6), Matches: rowsNorm_(ss, 'Matches', 16) }; }
/* Remembers what the two tabs looked like right after the app last wrote them, so a later difference can only be a change made by hand in the sheet. */
function snapshotBase_(ss, cur) {
  cur = cur || sheetLibRows_(ss);
  writeTab_(ss, 'Sync base', HEADINGS['Sync base'], chunkRows_(JSON.stringify(cur)));
  PropertiesService.getScriptProperties().setProperty('baseHash', hash_(canon_(cur)));
}
/* The Players and Teams tabs exactly as the app builds them. */
function libRows_(state) {
  var s = normCore_(state), nmOf = {};
  s.lib.players.forEach(function (p) { nmOf[p.id] = p.name; });
  return {
    players: s.lib.players.map(function (p) {
      return [p.id, p.name, s.lib.teams.filter(function (t) { return (t.playerIds || []).indexOf(p.id) >= 0; }).map(function (t) { return t.name; }).join(', ')];
    }),
    teams: s.lib.teams.map(function (t) {
      var ids = (t.playerIds || []).filter(function (i) { return nmOf[i] !== undefined; });
      return [t.id, t.name, ids.map(function (i) { return nmOf[i]; }).join(', '), nmOf[t.capId] || '', nmOf[t.wkId] || '', ids.join(',')];
    })
  };
}
function writeLibTabs_(ss, state) {
  var r = libRows_(state);
  writeTab_(ss, 'Players', HEADINGS['Players'], r.players);
  writeTab_(ss, 'Teams', HEADINGS['Teams'], r.teams);
}
function parseDate_(v) {
  var s = String(v || ''), m = /^(\d{4}-\d{2}-\d{2})/.exec(s);
  if (m) return m[1] + 'T12:00:00.000Z';
  var d = new Date(s);
  return isNaN(d.getTime()) ? '' : d.toISOString();
}
/* Rebuilds a match history record from its row in the Matches tab. The row holds the summary (teams, toss, scores, result) but not the ball-by-ball detail. */
function recoverMatch_(r, now) {
  var at = parseDate_(r[1]) || now, scores = r[7] ? r[7].split(' | ') : [];
  return {
    id: r[0], u: at, at: at, t1: r[2], t2: r[3], format: r[4] || 'Custom', overs: r[5], toss: r[6], scores: scores, result: r[8],
    tn: r[10] || (/^Tournament/i.test(r[9]) ? 'Tournament' : ''),
    tj: { mode: r[11] ? 'Bat flip' : 'Toss done off the app', caller: r[11], call: r[12], landed: r[13], winner: r[14], choice: r[15] },
    summary: scores.join('\n') + (r[8] ? '\n\n' + r[8] : '') + '\n\n(Recovered from the Google Sheet. The ball-by-ball detail was not available.)',
    card: '[]', lu: '[]', bb: '[]'
  };
}
/* Works out what was changed by hand in the sheet (base = how the tabs were left by the app, cur = how they are now) and applies it to the shared data.
   A new row adds a player or team, a changed name or squad renames or updates it, and a removed row deletes it. */
function absorbDiff_(stored, base, cur, now) {
  var s = normCore_(JSON.parse(JSON.stringify(stored))), changed = false, del = s.del;
  var players = s.lib.players.slice(), teams = s.lib.teams.slice();
  var pById = {}, pByName = {}, tById = {}, tByName = {};
  function reindex() {
    pById = {}; pByName = {}; tById = {}; tByName = {};
    players.forEach(function (p) { pById[p.id] = p; if (!pByName[nameKey_(p.name)]) pByName[nameKey_(p.name)] = p; });
    teams.forEach(function (t) { tById[t.id] = t; if (!tByName[nameKey_(t.name)]) tByName[nameKey_(t.name)] = t; });
  }
  reindex();
  var baseP = {}, curP = {}, baseT = {}, curT = {};
  (base.Players || []).forEach(function (r) { if (r[0]) baseP[r[0]] = r; });
  (cur.Players || []).forEach(function (r) { if (r[0]) curP[r[0]] = r; });
  (base.Teams || []).forEach(function (r) { if (r[0]) baseT[r[0]] = r; });
  (cur.Teams || []).forEach(function (r) { if (r[0]) curT[r[0]] = r; });

  // Players: new rows and renames
  (cur.Players || []).forEach(function (r) {
    var id = r[0], nm = r[1];
    if (!nm) return;
    if (id && pById[id]) {
      var b = baseP[id];
      if (b && b[1] !== nm && pById[id].name !== nm) { pById[id].name = nm; pById[id].u = now; changed = true; }
    } else if (!pByName[nameKey_(nm)] && !(id && del['p:' + id])) {
      players.push({ id: id || newId_('p'), name: nm, u: now });
      reindex(); changed = true;
    }
  });
  // Players: rows removed in the sheet (never when the whole tab is empty, which is more likely a mistake than a wish)
  if ((cur.Players || []).length > 0) {
    Object.keys(baseP).forEach(function (id) {
      if (curP[id] || !pById[id]) return;
      players = players.filter(function (p) { return p.id !== id; });
      del['p:' + id] = now; changed = true;
      teams.forEach(function (t) {
        if ((t.playerIds || []).indexOf(id) >= 0) {
          t.playerIds = t.playerIds.filter(function (x) { return x !== id; });
          if (t.capId === id) t.capId = '';
          if (t.wkId === id) t.wkId = '';
          t.u = now;
        }
      });
    });
    reindex();
  }

  // Teams
  function pid(name) {
    var k = nameKey_(name);
    if (!k) return '';
    if (pByName[k]) return pByName[k].id;
    players.push({ id: newId_('p'), name: name, u: now });
    reindex(); changed = true;
    return players[players.length - 1].id;
  }
  function look(name) { var p = pByName[nameKey_(name)]; return p ? p.id : ''; }
  function idsOf(text) {
    var ids = [];
    (text ? text.split(',') : []).forEach(function (n) {
      n = n.replace(/\s+/g, ' ').trim();
      var i = n ? pid(n) : '';
      if (i && ids.indexOf(i) < 0) ids.push(i);
    });
    return ids;
  }
  (cur.Teams || []).forEach(function (r) {
    var id = r[0], name = r[1];
    if (!name) return;
    if (id && tById[id]) {
      var b = baseT[id];
      if (b && (b[1] !== name || b[2] !== r[2] || b[3] !== r[3] || b[4] !== r[4])) {
        var t = tById[id], ids = idsOf(r[2]), cp = look(r[3]), wp = look(r[4]);
        t.name = name; t.playerIds = ids; t.capId = ids.indexOf(cp) >= 0 ? cp : ''; t.wkId = ids.indexOf(wp) >= 0 ? wp : ''; t.u = now;
        reindex(); changed = true;
      }
    } else if (!tByName[nameKey_(name)] && !(id && del['t:' + id])) {
      var ids2 = idsOf(r[2]), cp2 = look(r[3]), wp2 = look(r[4]);
      teams.push({ id: id || newId_('t'), name: name, playerIds: ids2, capId: ids2.indexOf(cp2) >= 0 ? cp2 : '', wkId: ids2.indexOf(wp2) >= 0 ? wp2 : '', u: now });
      reindex(); changed = true;
    }
  });
  if ((cur.Teams || []).length > 0) {
    Object.keys(baseT).forEach(function (id) {
      if (curT[id] || !tById[id]) return;
      teams = teams.filter(function (t) { return t.id !== id; });
      del['t:' + id] = now; changed = true;
    });
  }
  // Matches that are in the sheet but not in the app are brought into the app (never the other way round, so a row deleted by hand simply comes back).
  var hist = s.hist.slice(), hById = {};
  hist.forEach(function (x) { hById[x.id] = true; });
  (cur.Matches || []).forEach(function (r) {
    var id = r[0];
    if (!id || hById[id] || del['m:' + id] || !r[2] || !r[3]) return;
    hist.push(recoverMatch_(r, now)); hById[id] = true; changed = true;
  });
  return {
    changed: changed,
    state: normCore_({ liveAt: s.liveAt, live: s.live, del: del, lib: { players: players, teams: teams }, hist: hist, tours: s.tours })
  };
}
/* ---- Tournaments: rebuilt from the Tournaments, Tournament teams, People, Fixtures and Activity tabs when the app does not have them ---- */
var STAGE_KEY_ = { 'Group stage': 'group', 'Quarter-finals': 'qf', 'Semi-finals': 'sf', 'Final': 'final' };
function tourTypeOf_(label) {
  var s = String(label || '');
  if (/IPL/i.test(s)) return 'IPL';
  if (/World Test|WTC|Test/i.test(s)) return 'WTC';
  return 'Custom';
}
function isoOf_(v, now) {
  var s = String(v || '');
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return s;
  return parseDate_(s) || now;
}
function winOf_(f, nameA, nameB) {
  var w = nameKey_(f[14]), t = String(f[6] || ''), tk = nameKey_(t);
  if (w && w === nameKey_(nameA)) return 'a';
  if (w && w === nameKey_(nameB)) return 'b';
  if (/tied/i.test(t)) return 'tie';
  if (/draw/i.test(t)) return 'draw';
  if (tk.indexOf(nameKey_(nameA) + ' won') === 0) return 'a';
  if (tk.indexOf(nameKey_(nameB) + ' won') === 0) return 'b';
  return 'tie';
}
/* tabs = { Tournaments, TTeams, People, Fixtures, Activity } as rows of text. Every tournament listed in the Tournaments tab that the app lacks is rebuilt. */
function rebuildTours_(stateIn, tabs, now) {
  var s = normCore_(stateIn), have = {}, tours = s.tours.slice(), changed = false;
  tours.forEach(function (t) { have[t.id] = true; });
  var libTeamByName = {};
  s.lib.teams.forEach(function (t) { if (!libTeamByName[nameKey_(t.name)]) libTeamByName[nameKey_(t.name)] = t.id; });
  (tabs.Tournaments || []).forEach(function (r) {
    var id = r[0];
    if (!id || have[id] || s.del['o:' + id] || !r[1]) return;
    var mine = function (rows, idIdx) {
      return (rows || []).filter(function (x) { return x[idIdx] ? x[idIdx] === id : x[0] === r[1]; });
    };
    var type = tourTypeOf_(r[2]), at = parseDate_(r[8]) || now;
    var ov = parseInt(r[3], 10);
    // The updated time is the creation date, so a fuller copy that any device still holds always wins over this rebuild.
    var T = { id: id, name: r[1], type: type, overs: (type === 'Custom' && ov > 0) ? ov : null, teams: [], fixtures: [], at: at, u: at, people: [], owner: {}, log: [] };
    var tidByName = {}, pidByName = {};
    var team = function (name) {
      var k = nameKey_(name);
      if (!k) return '';
      if (tidByName[k]) return tidByName[k];
      var tid = libTeamByName[k] || newId_('t');
      tidByName[k] = tid; T.teams.push({ id: tid, name: name });
      return tid;
    };
    var person = function (name) {
      var k = nameKey_(name);
      if (!k) return '';
      if (pidByName[k]) return pidByName[k];
      var pid = newId_('p');
      pidByName[k] = pid; T.people.push({ id: pid, name: name });
      return pid;
    };
    // People, then the teams with who plays as each
    var people = mine(tabs.People, 5);
    if (people.length) people.forEach(function (p) { person(p[1]); });
    else (r[9] ? r[9].split(',') : []).forEach(function (n) { person(n.replace(/\s+/g, ' ').trim()); });
    var tt = mine(tabs.TTeams, 9);
    if (tt.length) {
      tt.forEach(function (x) {
        var tid = team(x[1]);
        if (tid && x[2]) T.owner[tid] = person(x[2]);
      });
    } else (r[5] ? r[5].split(',') : []).forEach(function (n) { team(n.replace(/\s+/g, ' ').trim()); });
    // Fixtures, in their scheduled order
    var fx = mine(tabs.Fixtures, 16).map(function (f, i) { return { f: f, i: i }; });
    fx.sort(function (a, b) { return (Number(a.f[1]) || 0) - (Number(b.f[1]) || 0) || a.i - b.i; });
    var stageCount = {};
    fx.forEach(function (x) {
      var f = x.f, a = team(f[3]), b = team(f[4]);
      if (!a || !b) return;
      var sk = STAGE_KEY_[f[10]] || '';
      var o = { id: newId_('f'), a: a, b: b, label: f[2], done: false, result: null };
      if (sk) { o.stage = sk; o.slot = stageCount[sk] || 0; stageCount[sk] = o.slot + 1; }
      o.pa = f[11] ? person(f[11]) : '';
      o.pb = f[12] ? person(f[12]) : '';
      if (f[5] === 'Played') {
        var d = parseDate_(f[8]);
        o.done = true;
        o.result = { win: winOf_(f, f[3], f[4]), text: f[6], rec: f[9], at: d, doneAt: d, toss: '', scores: f[7] ? f[7].split(' | ') : [], summary: '', score: f[13] };
      } else if (f[5] === 'In progress') o.live = true;
      T.fixtures.push(o);
    });
    // Saved activity
    mine(tabs.Activity, 3).forEach(function (x) { T.log.push({ at: isoOf_(x[1], now), text: x[2] }); });
    T.log.push({ at: now, text: 'Rebuilt from the Google Sheet' });
    // A tournament marked finished by hand stays finished; one that finished by playing its final or every fixture needs no flag.
    var auto = T.fixtures.some(function (f) { return f.stage; });
    var natural = auto ? T.fixtures.some(function (f) { return f.stage === 'final' && f.done; })
      : (T.fixtures.length > 0 && T.fixtures.every(function (f) { return f.done; }));
    if (r[4] === 'Finished' && !natural) T.finished = true;
    tours.push(T); have[id] = true; changed = true;
  });
  return { changed: changed, state: normCore_({ liveAt: s.liveAt, live: s.live, del: s.del, lib: s.lib, hist: s.hist, tours: tours }) };
}
/* Reads the five tournament tabs only when the Tournaments tab lists something the app does not have. */
function recoverTours_(ss, state, now) {
  var s = normCore_(state), have = {};
  s.tours.forEach(function (t) { have[t.id] = true; });
  var list = rowsNorm_(ss, 'Tournaments', 12);
  var missing = list.filter(function (r) { return r[0] && r[1] && !have[r[0]] && !s.del['o:' + r[0]]; });
  if (!missing.length) return null;
  var r = rebuildTours_(s, {
    Tournaments: list, TTeams: rowsNorm_(ss, 'Tournament teams', 10), People: rowsNorm_(ss, 'People', 6),
    Fixtures: rowsNorm_(ss, 'Fixtures', 17), Activity: rowsNorm_(ss, 'Activity', 4)
  }, now);
  return r.changed ? r.state : null;
}
/* Compares the sheet with the shared data. Returns the shared data with anything the app was missing added, plus any hand edits applied, or null when the app already has everything.
   fallback is the copy the app just sent, used when nothing is saved yet. */
function absorbSheet_(ss, stored, fallback) {
  var props = PropertiesService.getScriptProperties();
  var cur = sheetLibRows_(ss);
  var st = stored || fallback;
  if (!st) return null;
  var emptyBase = { Players: [], Teams: [], Matches: [] }, base = emptyBase;
  var same = hash_(canon_(cur)) === props.getProperty('baseHash');
  if (!same) {
    // The sheet was changed since the app last wrote it (or has never been compared), so edits and deletions are worked out against the snapshot.
    var baseText = loadChunks_(ss, 'Sync base');
    try { base = baseText ? JSON.parse(baseText) : emptyBase; } catch (x) { base = emptyBase; }
  }
  var nowIso = new Date().toISOString();
  var r = absorbDiff_(st, base, cur, nowIso);
  var rt = recoverTours_(ss, r.state, nowIso);
  if (rt) r = { changed: true, state: rt };
  if (r.changed) return r.state;
  if (!same) snapshotBase_(ss, cur);
  return null;
}
/* Applies hand edits from the sheet and tells every device to load them. Must be called while holding the script lock. */
function absorbCommit_(ss, props) {
  var text = loadChunks_(ss, 'App data');
  if (!text) return false;
  var stored = null;
  try { stored = JSON.parse(text); } catch (x) { return false; }
  var ab = absorbSheet_(ss, stored);
  if (!ab) return false;
  saveState_(ss, props, ab, '');
  writeLibTabs_(ss, ab);
  snapshotBase_(ss);
  return true;
}
/* Called when the app checks for news. Looks at the sheet at most once every few seconds, however many devices are open. */
function absorbThrottled_() {
  var props = PropertiesService.getScriptProperties();
  var last = Number(props.getProperty('lastAbs') || 0), now = new Date().getTime();
  if (now - last < 6000) return;
  props.setProperty('lastAbs', String(now));
  var lk = LockService.getScriptLock(), got = false;
  try {
    got = lk.tryLock(3000);
    if (got) absorbCommit_(getSS_(), props);
  } catch (e) {
  } finally {
    if (got) { try { lk.releaseLock(); } catch (x) {} }
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var data = JSON.parse(e.postData.contents);
    if (data.app !== 'CricketScorer') {
      return out_({ ok: false, error: 'Not a CricketScorer payload' });
    }
    var ss = getSS_();
    var props = PropertiesService.getScriptProperties();
    var tabs = data.tabs;

    // Older versions of the app sent fixed fields instead of a list of tabs.
    if (!tabs || !tabs.length) {
      // A copy that carries the saved data but no tabs must never blank the tabs.
      if (data.state) tabs = [];
      else tabs = [
        { name: 'Players', head: ['Id', 'Name'], rows: data.players || [] },
        { name: 'Teams', head: ['Id', 'Team', 'Players'], rows: data.teams || [] },
        { name: 'Matches', head: data.head || [], rows: data.rows || [] },
        { name: 'Live', head: data.liveHead || [], rows: data.live || [] }
      ];
    }

    // What arrives is MERGED with what is already saved, never written over it. Whatever any device enters is kept.
    var writeTabs = true, absorbed = false, merged = null;
    if (data.state) {
      var incoming = null;
      try { incoming = JSON.parse(data.state); } catch (pe) {}
      if (incoming) {
        var stored = null, text = loadChunks_(ss, 'App data');
        if (text) { try { stored = JSON.parse(text); } catch (se) {} }
        // Anything the sheet has that the app lacks (typed by hand, or left by an older version) is taken in first, so a save from the app can never overwrite it.
        var ab = absorbSheet_(ss, stored, incoming);
        if (ab) { stored = ab; absorbed = true; }
        var inN = normCore_(incoming);
        merged = stored ? mergeStates_(stored, incoming) : inN;
        var mC = canon_(merged), iC = canon_(inN), sC = stored ? canon_(normCore_(stored)) : '';
        // The readable tabs are only written from a copy that already contains everything saved. A copy that is missing something waits until that device has loaded it and sends again.
        writeTabs = (mC === iC);
        if (absorbed || mC !== sC) {
          // stateBy names the device whose copy already had everything, so that device need not download it again.
          saveState_(ss, props, merged, writeTabs ? String(data.dev || '') : '');
        } else if (!writeTabs) {
          // Nothing new was saved, but the sending device is missing something, so tell every device to reload.
          props.setProperty('stateAt', new Date().toISOString());
          props.setProperty('stateBy', '');
        }
      }
    }

    var counts = {}, wroteLib = false;
    if (writeTabs) {
      for (var i = 0; i < tabs.length; i++) {
        var t = tabs[i];
        if (!t || !t.name || t.name === 'App data' || t.name === 'Sync base') continue;
        var head = (t.head && t.head.length) ? t.head : (HEADINGS[t.name] || []);
        try {
          var wrote = writeTab_(ss, String(t.name), head, t.rows || []);
          if (wrote && (t.name === 'Players' || t.name === 'Teams' || t.name === 'Matches')) wroteLib = true;
          counts[t.name] = (t.rows || []).length;
        } catch (tabErr) {
          // One bad tab must never stop the others from being written.
          counts[t.name] = 'error: ' + String(tabErr);
        }
      }
      props.setProperty('counts', JSON.stringify(counts));
    }
    // Changes taken in from the sheet are written back as clean rows (new rows get their ids).
    if (absorbed && !writeTabs && merged) { writeLibTabs_(ss, merged); wroteLib = true; }
    // Tabs the app did not send this time still get their headings.
    ensureHeadings_(ss);
    if (wroteLib || absorbed) snapshotBase_(ss);

    var sentAt = data.sentAt || new Date().toISOString();
    var recent = [];
    try { recent = JSON.parse(props.getProperty('recent') || '[]'); } catch (rx) {}
    recent.push(sentAt);
    if (recent.length > 30) recent = recent.slice(-30);
    props.setProperty('recent', JSON.stringify(recent));
    props.setProperty('lastSync', sentAt);

    return out_({ ok: true, sentAt: sentAt, rows: counts, tabsWritten: writeTabs, absorbed: absorbed });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : '';
  var props = PropertiesService.getScriptProperties();
  // A very small reply that only says when the shared data last changed, and which device changed it. The app checks this often.
  // It also picks up anything typed by hand into the Players or Teams tab, so the app learns about it within seconds.
  if (action === 'stateat') {
    try { absorbThrottled_(); } catch (xa) {}
    return out_({ ok: true, stateAt: props.getProperty('stateAt') || '', stateBy: props.getProperty('stateBy') || '' });
  }
  // Counts of what the sheet holds, so the app can show whether both sides are in step.
  if (action === 'counts') {
    try {
      var ssc = getSS_(), cnt = { ok: true, sheet: {}, saved: {} };
      cnt.sheet.players = readTab_(ssc, 'Players').length;
      cnt.sheet.teams = readTab_(ssc, 'Teams').length;
      cnt.sheet.matches = readTab_(ssc, 'Matches').length;
      cnt.sheet.tournaments = readTab_(ssc, 'Tournaments').length;
      var stx = null, tx = loadChunks_(ssc, 'App data');
      if (tx) { try { stx = normCore_(JSON.parse(tx)); } catch (xp) {} }
      if (stx) { cnt.saved.players = stx.lib.players.length; cnt.saved.teams = stx.lib.teams.length; cnt.saved.matches = stx.hist.length; cnt.saved.tournaments = stx.tours.length; }
      return out_(cnt);
    } catch (ecn) {
      return out_({ ok: false, error: String(ecn) });
    }
  }
  // A quick reply the app uses to confirm that what it sent has been received.
  if (action === 'check') {
    var chk = { ok: true, lastSync: props.getProperty('lastSync') || 'never', stateAt: props.getProperty('stateAt') || '', stateBy: props.getProperty('stateBy') || '', recent: [] };
    try { chk.recent = JSON.parse(props.getProperty('recent') || '[]'); } catch (xr) {}
    try { getSS_(); } catch (ec) { chk.ok = false; chk.problem = 'Could not open a spreadsheet: ' + String(ec); }
    return out_(chk);
  }
  // The app asks for the shared state to load what was entered on another device or typed into the sheet.
  if (action === 'state') {
    var lk = LockService.getScriptLock(), got = false;
    try {
      got = lk.tryLock(15000);
      var ss1 = getSS_();
      if (got) { try { absorbCommit_(ss1, props); } catch (xc) {} }
      var text = loadChunks_(ss1, 'App data');
      return out_({ ok: true, state: text, stateAt: props.getProperty('stateAt') || '', stateBy: props.getProperty('stateBy') || '' });
    } catch (errS) {
      return out_({ ok: false, error: String(errS) });
    } finally {
      if (got) { try { lk.releaseLock(); } catch (xl) {} }
    }
  }
  // Every visit makes sure all tabs exist with their headings.
  try { ensureHeadings_(getSS_()); } catch (x0) {}
  // The app asks for the saved players and teams so they can be picked for a match on any device.
  if (action === 'lib') {
    try {
      var ss0 = getSS_();
      return out_({ ok: true, players: readTab_(ss0, 'Players'), teams: readTab_(ss0, 'Teams') });
    } catch (err) {
      return out_({ ok: false, error: String(err) });
    }
  }
  var info = { ok: true, app: 'CricketScorer' };
  try {
    var ss = getSS_();
    ensureHeadings_(ss);
    info.sheet = ss.getName();
    info.sheetUrl = ss.getUrl();
    info.tabs = ss.getSheets().map(function (s) { return s.getName(); });
  } catch (e2) {
    info.ok = false;
    info.problem = 'Could not open a spreadsheet: ' + String(e2);
  }
  var counts = {};
  try { counts = JSON.parse(props.getProperty('counts') || '{}'); } catch (x) {}
  info.lastSync = props.getProperty('lastSync') || 'never';
  info.stateAt = props.getProperty('stateAt') || '';
  info.stateBy = props.getProperty('stateBy') || '';
  info.rows = counts;
  return out_(info);
}

/* Reads a tab's rows without its heading row. */
function readTab_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var width = Math.max(sh.getLastColumn(), 1);
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  return vals.filter(function (r) { return String(r[1] === undefined ? '' : r[1]) !== '' || String(r[0]) !== ''; })
    .map(function (r) { return r.map(function (v) { return v === null || v === undefined ? '' : String(v); }); });
}

function writeTab_(ss, name, header, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);

  var width = header.length;
  rows.forEach(function (r) { if (r.length > width) width = r.length; });
  if (width === 0) return false;

  var values = [];
  if (header.length) values.push(pad_(header, width));
  rows.forEach(function (r) { values.push(pad_(r, width)); });

  // A tab whose content has not changed is left alone, which keeps every save fast.
  var props = PropertiesService.getScriptProperties();
  var key = 'h_' + name;
  var hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(values)));
  if (props.getProperty(key) === hash && sh.getLastRow() >= values.length && sh.getRange(1, 1).getValue() !== '') return false;
  sh.clearContents();

  sh.getRange(1, 1, values.length, width).setValues(values);
  if (header.length) styleHeading_(sh, width);
  props.setProperty(key, hash);
  return true;
}

function styleHeading_(sh, width) {
  var r = sh.getRange(1, 1, 1, width);
  r.setFontWeight('bold');
  r.setBackground('#d9efe9');
  sh.setFrozenRows(1);
}

/* Creates any missing tab and writes the headings into any tab whose first row is empty. */
function ensureHeadings_(ss) {
  TAB_ORDER.forEach(function (name) {
    var head = HEADINGS[name];
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var first = sh.getLastRow() >= 1 ? sh.getRange(1, 1, 1, head.length).getValues()[0] : [];
    var empty = true;
    for (var i = 0; i < first.length; i++) { if (String(first[i]) !== '') { empty = false; break; } }
    if (empty) {
      sh.getRange(1, 1, 1, head.length).setValues([head]);
      styleHeading_(sh, head.length);
    }
  });
  // The default empty "Sheet1" is not needed once the real tabs exist.
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) { try { ss.deleteSheet(def); } catch (e) {} }
}

/* Finds the spreadsheet to write to. */
function getSS_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var props = PropertiesService.getScriptProperties();
  var id = SPREADSHEET_ID || props.getProperty('ssId');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('CricketScorer data');
  props.setProperty('ssId', ss.getId());
  return ss;
}

/* Adds a CricketScorer menu to the sheet when the script is bound to it. */
function onOpen() {
  try { SpreadsheetApp.getUi().createMenu('CricketScorer').addItem('Add tabs and headings', 'setupSheets').addToUi(); } catch (e) {}
}

/* Run this once from the editor to create every tab with its headings. */
function setupSheets() {
  var ss = getSS_();
  ensureHeadings_(ss);
  Logger.log('Tabs and headings are ready: ' + ss.getUrl());
  return 'Tabs and headings are ready: ' + ss.getUrl();
}

function pad_(row, width) {
  var r = [];
  for (var i = 0; i < width; i++) {
    var v = i < row.length ? row[i] : '';
    if (v === null || v === undefined) v = '';
    if (typeof v === 'object') v = JSON.stringify(v);
    if (typeof v === 'string' && v.length > MAX_CELL) v = v.slice(0, MAX_CELL);
    r.push(v);
  }
  return r;
}

function out_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
