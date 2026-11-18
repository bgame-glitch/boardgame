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
for (const m of fresh) {
  byId.set(m.id, {
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
};
await writeFile(OUT, JSON.stringify(out, null, 2) + "\n");
console.log(`일정 ${fresh.length}개 읽음, 저장 ${meetups.length}개`);
