const todayISO = () => new Date().toISOString().slice(0, 10);

const parseVisitDate = (value) => {
  if (!value) {
    return null;
  }
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }
  const text = String(value).trim();
  const direct = new Date(text);
  if (!Number.isNaN(direct.getTime())) {
    return direct;
  }
  const normalized = text.replace(/\s+/g, "");
  const m = normalized.match(/^(\d{2,4})[./-](\d{1,2})[./-](\d{1,2})$/);
  if (m) {
    const y = m[1].length === 2 ? 2000 + Number(m[1]) : Number(m[1]);
    const mo = Number(m[2]) - 1;
    const d = Number(m[3]);
    const dt = new Date(y, mo, d);
    return Number.isNaN(dt.getTime()) ? null : dt;
  }
  const parsed = new Date(
    normalized.replace(/\./g, "-").replace(/\//g, "-")
  );
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const formatDate = (value) => {
  if (!value) {
    return "";
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  const yy = String(date.getFullYear()).slice(-2);
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yy}/${mm}/${dd}`;
};

const toISODate = (value) => {
  const d = parseVisitDate(value);
  if (!d) {
    return "";
  }
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
};

const toAssignmentDateText = (value) => {
  const d = parseVisitDate(value);
  if (!d) {
    return "";
  }
  const yy = String(d.getFullYear()).slice(-2);
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yy}/${mm}/${dd}`;
};

const formatVolunteerDateText = (value) => {
  if (!value) {
    return "";
  }
  const d = parseVisitDate(value);
  if (!d) {
    return "";
  }
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${m}/${day}`;
};

const isTrueValue = (value) =>
  value === true ||
  value === "TRUE" ||
  value === "true" ||
  value === "Y" ||
  value === "1" ||
  value === 1;

const getVisitDateValue = (row) =>
  row["방문일"] ||
  row["방문날짜"] ||
  row["방문일자"] ||
  row["날짜"] ||
  row["방문일시"] ||
  row["일자"] ||
  "";

const getVisitCardNumber = (row) => row["카드번호"] || row["구역카드"] || "";

const parseCardNumber = (value) => {
  const raw = String(value || "").trim();
  const mPair = raw.match(/(\d+)\s*-\s*(\d+)/);
  if (mPair) {
    const area = Number(mPair[1]);
    const num = Number(mPair[2]);
    if (!Number.isNaN(area) && !Number.isNaN(num)) {
      return { raw, area, num };
    }
  }
  const mSingle = raw.match(/(\d+)/);
  if (mSingle) {
    const area = Number(mSingle[1]);
    if (!Number.isNaN(area)) {
      return { raw, area, num: null };
    }
  }
  return { raw, area: null, num: null };
};

const compareCardNumbers = (a, b) => {
  const pa = parseCardNumber(a);
  const pb = parseCardNumber(b);
  if (pa.area != null && pb.area != null) {
    if (pa.area !== pb.area) {
      return pa.area - pb.area;
    }
    if (pa.num != null && pb.num != null && pa.num !== pb.num) {
      return pa.num - pb.num;
    }
    if (pa.num != null && pb.num == null) {
      return 1;
    }
    if (pa.num == null && pb.num != null) {
      return -1;
    }
  }
  return pa.raw.localeCompare(pb.raw, "ko-KR");
};

const normalizeAssignmentName = (value) => {
  return String(value || "")
    .replace(/\s*\((오전|오후)\)\s*$/g, "")
    .trim();
};

const formatAssignmentDate = (value) => {
  if (!value) {
    return "";
  }
  if (typeof value === "string") {
    if (/^\d{2}\/\d{2}\/\d{2}$/.test(value)) {
      return value;
    }
    const isoMatch = value.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) {
      const yy = isoMatch[1].slice(-2);
      const mm = isoMatch[2];
      const dd = isoMatch[3];
      return yy + "/" + mm + "/" + dd;
    }
  }
  try {
    const d = new Date(value);
    if (!isNaN(d.getTime())) {
      const yy = String(d.getFullYear()).slice(-2);
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const dd = String(d.getDate()).padStart(2, "0");
      return yy + "/" + mm + "/" + dd;
    }
  } catch (e) {
  }
  return String(value);
};

const toIsoDate = (dateStr) => {
  if (!dateStr) return null;
  const s = String(dateStr).trim();
  if (!s) return null;
  if (/^\d{2}\/\d{2}\/\d{2}$/.test(s)) {
    return "20" + s.replace(/\//g, "-");
  }
  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) return isoMatch[0];
  try {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d.toISOString().split("T")[0];
  } catch (e) {
  }
  return null;
};
