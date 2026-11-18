// 당근 모임 웹페이지에서 다가오는 일정을 읽어 meetups.json 에 합쳐 저장해요.
// GitHub Actions 가 2일마다 실행해요. 직접 실행: node scripts/fetch-meetups.mjs
import { readFile, writeFile } from "node:fs/promises";

const GROUP_PAGE = "https://www.daangn.com/kr/group/xgfk9uzqfji2/";
const OUT = new URL("../meetups.json", import.meta.url);
const KEEP_AFTER_MS = 6 * 60 * 60 * 1000; // 시작 6시간 뒤까지는 목록에 남겨둬요

function findKey(obj, key) {
  if (!obj || typeof obj !== "object") return undefined;
  if (Object.prototype.hasOwnProperty.call(obj, key)) return obj[key];
  for (const v of Object.values(obj)) {
    const r = findKey(v, key);
    if (r !== undefined) return r;
  }
  return undefined;
}

const res = await fetch(GROUP_PAGE, {
  headers: { "user-agent": "Mozilla/5.0 (meetup-board)", "accept-language": "ko-KR" },
});
if (!res.ok) throw new Error(`당근 모임 페이지를 못 열었어요: HTTP ${res.status}`);
const html = await res.text();

const start = html.indexOf("window.__remixContext = ");
if (start < 0) throw new Error("페이지 구조가 바뀌었어요: __remixContext 없음");
const end = html.indexOf(";</script>", start);
const ctx = JSON.parse(html.slice(start + "window.__remixContext = ".length, end));

const group = findKey(ctx, "communityGroup");
const fresh = findKey(ctx, "meetups");
if (!Array.isArray(fresh)) throw new Error("페이지 구조가 바뀌었어요: meetups 없음");

let saved = { meetups: [] };
try { saved = JSON.parse(await readFile(OUT, "utf8")); } catch {}

const now = new Date();
const byId = new Map((saved.meetups || []).map(m => [m.id, m]));

// ── 회원이 붙여넣은 공유 링크 (구글 설문지 응답 시트 CSV) ──────────────────
const GROUP_DB_ID = String(group?.dbId ?? saved.group?.dbId ?? "412730");
let config = {};
try { config = JSON.parse(await readFile(new URL("../meetup-config.json", import.meta.url), "utf8")); } catch {}

// "2026년 11월 26일 목요일, 오후 8:00" → 한국 시간 ISO
function parseKoreanDate(s) {
  const m = s && s.match(/(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일.*?(오전|오후)\s*(\d{1,2}):(\d{2})/);
  if (!m) return null;
  let h = +m[5] % 12; if (m[4] === "오후") h += 12;
  const p = n => String(n).padStart(2, "0");
  return new Date(`${m[1]}-${p(m[2])}-${p(m[3])}T${p(h)}:${m[6]}:00+09:00`).toISOString();
}
const remixCtx = html => {
  const s = html.indexOf("window.__remixContext = ");
  if (s < 0) return null;
  return JSON.parse(html.slice(s + 24, html.indexOf(";</script>", s)));
};
const get = url => fetch(url, { headers: { "user-agent": "Mozilla/5.0 (meetup-board)", "accept-language": "ko-KR" } }).then(r => r.text());

async function resolveLink(link) {
  let url = link.trim();
  if (!/^https:\/\/(www\.)?daangn\.com\//.test(url)) return null;
  // 공유 링크(/kr/share/...)는 일정 페이지 주소를 안에 담고 있어요
  if (!/\/kr\/meetups\//.test(url)) {
    const h = await get(url);
    const c = h.match(/https:\/\/www\.daangn\.com\/kr\/meetups\/[A-Za-z0-9]+\/?/);
    if (!c) return null;
    url = c[0];
  }
  const ctx = remixCtx(await get(url));
  const m = ctx && findKey(ctx, "meetup");
  if (!m || String(m.groupId) !== GROUP_DB_ID) return null; // 우리 모임 일정만
  return {
    id: String(m.meetupId),
    name: m.title,
    description: m.description || "",
    image: m.image?.small || m.image?.url || null,
    meetAt: parseKoreanDate(m.meetupDateDescription),
  };
}

const linkCache = saved.links || {}; // 링크 → { id, meetAt } (이미 지난 일정은 다시 안 읽어요)
if (config.linksCsvUrl) {
  try {
    const csv = await get(config.linksCsvUrl);
    const links = [...new Set(csv.match(/https:\/\/(?:www\.)?daangn\.com\/[^\s",]+/g) || [])];
    let added = 0;
    for (const link of links) {
      const c = linkCache[link];
      if (c && (c.bad || new Date(c.meetAt).getTime() + KEEP_AFTER_MS < now.getTime())) continue;
      try {
        const r = await resolveLink(link);
        if (!r || !r.meetAt) { linkCache[link] = { bad: true }; continue; }
        linkCache[link] = { id: r.id, meetAt: r.meetAt };
        const prev = byId.get(r.id);
        byId.set(r.id, {
          status: "OPENED", current: null, max: null, place: null, groupOnly: false,
          ...prev, // 모임 페이지에서 읽은 장소·인원이 있으면 그대로 둬요
          id: r.id, name: r.name, meetAt: r.meetAt, description: r.description,
          image: prev?.image || r.image,
          fromLink: true, lastSeenAt: prev?.lastSeenAt ?? now.toISOString(),
        });
        added++;
      } catch (e) { console.warn("링크를 못 읽었어요:", link, e.message); }
    }
    console.log(`붙여넣은 링크 ${links.length}개 중 ${added}개 반영`);
  } catch (e) { console.warn("링크 시트를 못 읽었어요:", e.message); }
}

for (const m of fresh) {
  const prev = byId.get(m.id);
  byId.set(m.id, {
    ...(prev?.description ? { description: prev.description } : {}),
    id: m.id,
    name: m.name,
    image: m.image || null,
    groupOnly: !!m.groupOnly,
    status: m.progressStatus, // OPENED / FULLY_RECRUITED / CLOSED
    current: m.currentParticipantsNumber,
    max: m.maximumParticipantsNumber,
    meetAt: m.meetAt,
    place: m.placeTag ? {
      name: m.placeTag.name,
      address: m.placeTag.addressText || "",
      lat: m.placeTag.coordinate?.lat ?? null,
      lon: m.placeTag.coordinate?.lon ?? null,
    } : null,
    lastSeenAt: now.toISOString(),
  });
}

const meetups = [...byId.values()]
  .filter(m => new Date(m.meetAt).getTime() + KEEP_AFTER_MS > now.getTime())
  .sort((a, b) => a.meetAt.localeCompare(b.meetAt));

const out = {
  group: { name: group?.name ?? saved.group?.name, dbId: group?.dbId ?? saved.group?.dbId, webUrl: GROUP_PAGE },
  fetchedAt: now.toISOString(),
  meetups,
  links: linkCache,
};
await writeFile(OUT, JSON.stringify(out, null, 2) + "\n");
console.log(`일정 ${fresh.length}개 읽음, 저장 ${meetups.length}개`);
