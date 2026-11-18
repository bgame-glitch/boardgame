/**
 * 당근 모임 일정 실시간 읽기 (구글 시트 Apps Script)
 *
 * 설치: 링크 응답 시트 → 확장 프로그램 → Apps Script → 이 내용 전체를 붙여넣고 저장
 *       → 배포 → 새 배포 → 유형: 웹 앱, 실행: 나, 액세스 권한: 모든 사용자 → 배포
 *
 *   주소?            → 다가오는 일정 목록 (모임 페이지 5개 + 시트에 붙여넣은 링크)
 *   주소?action=add&link=당근링크  → 링크를 시트에 저장하고 그 일정을 바로 돌려줘요
 */
const GROUP_PAGE = "https://www.daangn.com/kr/group/xgfk9uzqfji2/";
const GROUP_DB_ID = "412730";
const KEEP_AFTER_MS = 6 * 60 * 60 * 1000; // 시작 6시간 뒤까지 목록에 남겨요
const CACHE_SEC = 300;                     // 같은 목록은 5분 동안 재사용해요

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === "add") return json_(addLink_(p.link || ""));
    return json_(list_());
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function get_(url) {
  return UrlFetchApp.fetch(url, {
    muteHttpExceptions: true, followRedirects: true,
    headers: { "Accept-Language": "ko-KR" },
  }).getContentText("UTF-8");
}

function remixCtx_(html) {
  const s = html.indexOf("window.__remixContext = ");
  if (s < 0) return null;
  return JSON.parse(html.slice(s + 24, html.indexOf(";</script>", s)));
}

function findKey_(obj, key) {
  if (!obj || typeof obj !== "object") return undefined;
  if (Object.prototype.hasOwnProperty.call(obj, key)) return obj[key];
  for (const v of Object.values(obj)) {
    const r = findKey_(v, key);
    if (r !== undefined) return r;
  }
  return undefined;
}

// "2026년 11월 26일 목요일, 오후 8:00" → ISO (한국 시간)
function parseKoreanDate_(s) {
  const m = s && s.match(/(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일.*?(오전|오후)\s*(\d{1,2}):(\d{2})/);
  if (!m) return null;
  let h = Number(m[5]) % 12; if (m[4] === "오후") h += 12;
  const p = n => String(n).padStart(2, "0");
  return new Date(`${m[1]}-${p(m[2])}-${p(m[3])}T${p(h)}:${m[6]}:00+09:00`).toISOString();
}

function groupMeetups_() {
  const ctx = remixCtx_(get_(GROUP_PAGE));
  const arr = (ctx && findKey_(ctx, "meetups")) || [];
  return arr.map(m => ({
    id: String(m.id),
    name: m.name,
    image: m.image || null,
    groupOnly: !!m.groupOnly,
    status: m.progressStatus,
    current: m.currentParticipantsNumber,
    max: m.maximumParticipantsNumber,
    meetAt: m.meetAt,
    place: m.placeTag ? {
      name: m.placeTag.name,
      address: m.placeTag.addressText || "",
      lat: m.placeTag.coordinate ? m.placeTag.coordinate.lat : null,
      lon: m.placeTag.coordinate ? m.placeTag.coordinate.lon : null,
    } : null,
  }));
}

// 공유 링크 → 일정 정보 (우리 모임 일정이 아니면 null)
function resolveLink_(link) {
  let url = String(link).trim();
  if (!/^https:\/\/(www\.)?daangn\.com\//.test(url)) return null;
  if (!/\/kr\/meetups\//.test(url)) {
    const c = get_(url).match(/https:\/\/www\.daangn\.com\/kr\/meetups\/[A-Za-z0-9]+\/?/);
    if (!c) return null;
    url = c[0];
  }
  const ctx = remixCtx_(get_(url));
  const m = ctx && findKey_(ctx, "meetup");
  if (!m || String(m.groupId) !== GROUP_DB_ID) return null;
  const meetAt = parseKoreanDate_(m.meetupDateDescription);
  if (!meetAt) return null;
  return {
    id: String(m.meetupId),
    name: m.title,
    description: m.description || "",
    image: (m.image && (m.image.small || m.image.url)) || null,
    meetAt: meetAt,
    status: "OPENED", current: null, max: null, place: null, groupOnly: false, fromLink: true,
  };
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().find(s => /응답|Responses/.test(s.getName())) || ss.getSheets()[0];
}

function sheetLinks_() {
  const values = sheet_().getDataRange().getDisplayValues();
  const out = new Set();
  values.forEach(row => row.forEach(cell => {
    (String(cell).match(/https:\/\/(?:www\.)?daangn\.com\/[^\s",]+/g) || []).forEach(l => out.add(l));
  }));
  return [...out];
}

// 링크별로 읽은 결과를 저장해 두고 다시 읽지 않아요
function cachedLink_(link) {
  const props = PropertiesService.getScriptProperties();
  const key = "L:" + link;
  const saved = props.getProperty(key);
  if (saved) return JSON.parse(saved);
  let r = null;
  try { r = resolveLink_(link); } catch (e) { return null; } // 잠깐 실패면 다음에 다시
  props.setProperty(key, JSON.stringify(r || { bad: true }));
  return r || { bad: true };
}

function merge_(linkItems, groupItems) {
  const byId = new Map();
  linkItems.forEach(m => byId.set(m.id, m));
  groupItems.forEach(m => {
    const prev = byId.get(m.id);
    byId.set(m.id, Object.assign({}, m, prev && prev.description ? { description: prev.description } : {}));
  });
  const now = Date.now();
  return [...byId.values()]
    .filter(m => new Date(m.meetAt).getTime() + KEEP_AFTER_MS > now)
    .sort((a, b) => String(a.meetAt).localeCompare(String(b.meetAt)));
}

function list_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get("list");
  if (hit) return JSON.parse(hit);
  const linkItems = sheetLinks_().map(cachedLink_).filter(m => m && !m.bad);
  let groupItems = [];
  try { groupItems = groupMeetups_(); } catch (e) {}
  const out = { ok: true, live: true, fetchedAt: new Date().toISOString(), meetups: merge_(linkItems, groupItems) };
  try { cache.put("list", JSON.stringify(out), CACHE_SEC); } catch (e) {}
  return out;
}

function addLink_(link) {
  const found = String(link).match(/https:\/\/(?:www\.)?daangn\.com\/[^\s"]+/);
  if (!found || !/\/kr\/(share|meetups)\//.test(found[0])) {
    return { ok: false, error: "당근 일정 공유 링크가 아니에요." };
  }
  const url = found[0];
  const r = resolveLink_(url);
  if (!r) return { ok: false, error: "우리 모임 일정이 아니거나 읽을 수 없는 링크예요." };
  PropertiesService.getScriptProperties().setProperty("L:" + url, JSON.stringify(r));
  if (sheetLinks_().indexOf(url) < 0) sheet_().appendRow([new Date(), url]);
  CacheService.getScriptCache().remove("list");
  // 모임 페이지에 있는 일정이면 장소·인원까지 채워서 돌려줘요
  let full = r;
  try {
    const g = groupMeetups_().find(m => m.id === r.id);
    if (g) full = Object.assign({}, g, { description: r.description });
  } catch (e) {}
  return { ok: true, meetup: full };
}
