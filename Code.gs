/**
 * Fotobingo — backend för Properties & Partners kickoff, Tammsvik 2026-08-24
 * Poäng: varje bild ger 1 poäng (inget tak) + bonus för full bricka.
 *
 * Den här filen KÖRS I APPS SCRIPT, inte i repot. Den ligger här bara för
 * spårbarhet — klistra in innehållet i script.google.com.
 *
 * SETUP (nytt Apps Script-projekt — återanvänd inte bröllopens):
 * 1. Klistra in Drive-mappens ID i FOLDER_ID nedan.
 * 2. Kör rebuildIndex EN gång från editorn och godkänn Drive-behörigheten.
 *    Loggen ska säga "Index ombyggt: N bilder". Hoppas detta över kraschar
 *    uppladdningen med ett getFolderById-fel.
 * 3. Distribuera > Ny distribution > Webbapp: Kör som "Jag", Åtkomst "Alla".
 * 4. Kopiera /exec-URL:en till API_URL i BÅDE index.html och scoreboard.html.
 *
 * VID ÄNDRING AV DENNA FIL:
 *   Hantera distributioner > pennan > NY VERSION. Aldrig "ny distribution" —
 *   det ger en ny URL, och då pratar frontenden med den gamla koden.
 *
 * Lag identifieras med id t1–t8 (namnen bor i frontend, inte här).
 * Specialbilder: taskIndex 100 = lagbild, 101 = spontanbild.
 *
 * SPONTANLIGAN (nytt): spontanbilder skickas med who = personens id (t3m5),
 * som sparas i filnamnet och i indexet. Person-id:t innehåller inget namn,
 * så inga personuppgifter hamnar i filnamnen i Drive.
 */

const FOLDER_ID = "KLISTRA_IN_MAPP_ID_HAR";   // <-- Drive-mappens ID (allt efter /folders/)
const INDEX_KEY = "fotobingo_index_pp2026";    // eget index för detta event
const NUM_TASKS = 12;                          // måste matcha TASKS.length i frontend
const FULL_CARD_BONUS = 25;                    // bonus när alla 12 uppgifter är lösta
const SPONTAN_TASK = 101;                      // taskIndex för spontanbild

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json({ ok: false, error: "ingen postData mottogs" });
    }
    const data = JSON.parse(e.postData.contents);
    if (!data.image) return json({ ok: false, error: "ingen bild i anropet" });

    const folder = DriveApp.getFolderById(FOLDER_ID);
    const parts = data.image.split(",");
    if (parts.length < 2) return json({ ok: false, error: "felaktigt bildformat" });
    const bytes = Utilities.base64Decode(parts[1]);
    const mime = (parts[0].match(/data:(.*?);/) || [null, "image/jpeg"])[1];
    const ext = (mime.split("/")[1] || "jpg").replace(/[^a-z0-9]/gi, "");
    const ts = Date.now();

    // who = person-id, bara siffror/bokstäver släpps igenom in i filnamnet
    const who = data.who ? String(data.who).replace(/[^A-Za-z0-9]/g, "") : "";
    const name = "bord" + data.table + "_uppg" + data.taskIndex +
                 (who ? "_vem" + who : "") + "_" + ts + "." + ext;

    const file = folder.createFile(Utilities.newBlob(bytes, mime, name));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const item = {
      id: file.getId(),
      table: data.table,
      taskIndex: data.taskIndex,
      who: who || null,
      url: "https://lh3.googleusercontent.com/d/" + file.getId(),
      ts: ts
    };
    appendToIndex(item);
    return json({ ok: true, url: item.url, id: item.id });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action === "list")  return json({ ok: true, items: getIndex() });
    if (action === "score") return json({ ok: true, scores: computeScores(),
                                          spontan: computeSpontan(), items: getIndex() });
    return json({ ok: true, msg: "Fotobingo backend live" });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// Varje bild ger 1 poäng (inkl. lagbilder/spontanbilder) — inget tak.
// Bonus när alla NUM_TASKS unika uppgifter är lösta. Bilder utan lag tävlar inte.
function computeScores() {
  const items = getIndex();
  const byTable = {};
  items.forEach(it => {
    if (String(it.table) === "gemensam") return;
    const t = it.table;
    if (!byTable[t]) byTable[t] = { table: t, solved: {}, photos: 0 };
    byTable[t].photos++;
    if (it.taskIndex < NUM_TASKS) byTable[t].solved[it.taskIndex] = true;
  });
  return Object.keys(byTable).map(t => {
    const row = byTable[t];
    const solvedCount = Object.keys(row.solved).length;
    const full = solvedCount >= NUM_TASKS;
    return {
      table: row.table,
      solved: solvedCount,
      photos: row.photos,
      full: full,
      points: row.photos + (full ? FULL_CARD_BONUS : 0)
    };
  }).sort((a, b) => b.points - a.points);
}

// Spontanligan: antal spontanbilder per person. Bilder utan who kan inte räknas.
function computeSpontan() {
  const items = getIndex();
  const per = {};
  let utanVem = 0;
  items.forEach(it => {
    if (Number(it.taskIndex) !== SPONTAN_TASK) return;
    if (it.who) per[it.who] = (per[it.who] || 0) + 1;
    else utanVem++;
  });
  const rows = Object.keys(per).map(w => ({ who: w, count: per[w] }))
                     .sort((a, b) => b.count - a.count);
  return { rows: rows, missing: utanVem };
}

function getIndex() {
  const raw = PropertiesService.getScriptProperties().getProperty(INDEX_KEY);
  return raw ? JSON.parse(raw) : [];
}
function appendToIndex(item) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const items = getIndex();
    items.push(item);
    PropertiesService.getScriptProperties().setProperty(INDEX_KEY, JSON.stringify(items));
  } finally {
    lock.releaseLock();
  }
}

// Läser om hela Drive-mappen. Klarar både gamla filnamn (utan _vem) och nya.
function rebuildIndex() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  const items = [];
  while (files.hasNext()) {
    const f = files.next();
    const m = f.getName().match(/^bord(.+?)_uppg(\d+)(?:_vem([A-Za-z0-9]+))?_(\d+)\./);
    if (!m) continue;
    items.push({
      id: f.getId(),
      table: isNaN(+m[1]) ? m[1] : +m[1],
      taskIndex: +m[2],
      who: m[3] || null,
      url: "https://lh3.googleusercontent.com/d/" + f.getId(),
      ts: +m[4]
    });
  }
  items.sort((a, b) => a.ts - b.ts);
  PropertiesService.getScriptProperties().setProperty(INDEX_KEY, JSON.stringify(items));
  Logger.log("Index ombyggt: " + items.length + " bilder");
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
