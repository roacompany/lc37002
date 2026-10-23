// PDF 내장 폰트 이름(예: "ABCDEF+Arial-BoldMT")을 PowerPoint에서 쓸 수 있는 폰트 정보로 변환한다.

const KNOWN_FAMILIES = {
  arial: "Arial",
  arialmt: "Arial",
  helvetica: "Arial",
  helveticaneue: "Arial",
  arialnarrow: "Arial Narrow",
  arialblack: "Arial Black",
  timesnewroman: "Times New Roman",
  timesnewromanps: "Times New Roman",
  times: "Times New Roman",
  timesroman: "Times New Roman",
  courier: "Courier New",
  couriernew: "Courier New",
  couriernewps: "Courier New",
  calibri: "Calibri",
  calibrilight: "Calibri Light",
  cambria: "Cambria",
  cambriamath: "Cambria Math",
  candara: "Candara",
  consolas: "Consolas",
  constantia: "Constantia",
  corbel: "Corbel",
  georgia: "Georgia",
  garamond: "Garamond",
  tahoma: "Tahoma",
  trebuchetms: "Trebuchet MS",
  verdana: "Verdana",
  segoeui: "Segoe UI",
  segoeuilight: "Segoe UI Light",
  segoeuisemibold: "Segoe UI Semibold",
  bookantiqua: "Book Antiqua",
  centurygothic: "Century Gothic",
  impact: "Impact",
  symbol: "Symbol",
  wingdings: "Wingdings",
  roboto: "Roboto",
  opensans: "Open Sans",
  lato: "Lato",
  montserrat: "Montserrat",
  // 한글 폰트
  malgungothic: "맑은 고딕",
  malgun: "맑은 고딕",
  gulim: "굴림",
  gulimche: "굴림체",
  dotum: "돋움",
  dotumche: "돋움체",
  batang: "바탕",
  batangche: "바탕체",
  gungsuh: "궁서",
  gungsuhche: "궁서체",
  nanumgothic: "나눔고딕",
  nanumgothicbold: "나눔고딕",
  nanumgothicextrabold: "나눔고딕 ExtraBold",
  nanummyeongjo: "나눔명조",
  nanumbarungothic: "나눔바른고딕",
  nanumsquare: "나눔스퀘어",
  applesdgothicneo: "Apple SD Gothic Neo",
  notosanskr: "Noto Sans KR",
  notosanscjkkr: "Noto Sans CJK KR",
  notoserifkr: "Noto Serif KR",
  notoserifcjkkr: "Noto Serif CJK KR",
  pretendard: "Pretendard",
  spoqahansans: "Spoqa Han Sans",
  spoqahansansneo: "Spoqa Han Sans Neo",
  hygothic: "HY견고딕",
  hygothicmedium: "HY중고딕",
  hyheadline: "HY헤드라인M",
  // 일본어/중국어
  msgothic: "MS Gothic",
  msmincho: "MS Mincho",
  meiryo: "Meiryo",
  yugothic: "Yu Gothic",
  simsun: "SimSun",
  simhei: "SimHei",
  microsoftyahei: "Microsoft YaHei",
};

const STYLE_TOKENS = /(bold|black|heavy|semibold|demibold|extrabold|ultrabold|italic|oblique|light|thin|medium|regular|book|roman|normal|condensed)/gi;

function normalize(s) {
  return s.toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * @param {string} rawName  PDF 폰트 이름 (서브셋 접두어 포함 가능)
 * @param {{bold?: boolean, italic?: boolean, black?: boolean}} flags  pdf.js가 판별한 스타일
 * @param {string} genericFamily  pdf.js 텍스트 스타일의 fontFamily (serif / sans-serif / monospace)
 * @param {string} text  해당 텍스트 (한글 여부 판단용)
 */
export function resolveFont(rawName, flags = {}, genericFamily = "sans-serif", text = "") {
  let name = (rawName || "").replace(/^[A-Z]{6}\+/, "");
  const [base, ...styleParts] = name.split(/[-,]/);
  const styleStr = styleParts.join(" ") + " " + (base.match(STYLE_TOKENS) || []).join(" ");

  const bold = !!(flags.bold || flags.black || /bold|black|heavy/i.test(styleStr));
  const italic = !!(flags.italic || /italic|oblique/i.test(styleStr));

  let family = null;
  const key = normalize(base);
  const keyNoSuffix = key.replace(/(psmt|mt|ps)$/, "");
  const keyNoStyle = keyNoSuffix.replace(STYLE_TOKENS, "");
  for (const k of [key, keyNoSuffix, keyNoStyle]) {
    if (KNOWN_FAMILIES[k]) {
      family = KNOWN_FAMILIES[k];
      break;
    }
  }
  if (!family && base) {
    // CamelCase 이름은 공백을 넣어 일반적인 패밀리 이름으로 복원 (SourceSansPro → Source Sans Pro).
    // 원래 이름에 공백이 없는 글꼴(NanumSquareRound 등)도 있으므로, 띄운 이름이 없고 붙인 이름이 설치돼 있으면 그것을 쓴다
    const stripped = base.replace(/(PSMT|MT|PS)$/, "").replace(STYLE_TOKENS, "");
    if (stripped && /^[A-Za-z0-9 ]+$/.test(stripped)) {
      const spaced = stripped.replace(/([a-z])([A-Z])/g, "$1 $2").trim();
      family = !isFontAvailable(spaced) && isFontAvailable(stripped) ? stripped : spaced;
    } else if (stripped) {
      family = stripped;
    }
  }

  const hasCJK = /[ᄀ-ᇿ㄰-㆏가-힯぀-ヿ一-鿿]/.test(text);
  const fallback = hasCJK
    ? "맑은 고딕"
    : genericFamily === "serif"
      ? "Times New Roman"
      : genericFamily === "monospace"
        ? "Courier New"
        : "Arial";

  if (!family || !isFontAvailable(family)) {
    // 알려진 이름이면 그대로 쓰고(사용자 PC에 설치돼 있을 가능성), 아니면 비슷한 기본 폰트로 대체
    if (!family || !Object.values(KNOWN_FAMILIES).includes(family)) family = fallback;
  }
  return { family, bold, italic, fallback };
}

const availabilityCache = new Map();
let measureCtx = null;

/** 브라우저에 설치된 폰트인지 측정 방식으로 확인한다. */
export function isFontAvailable(family) {
  if (availabilityCache.has(family)) return availabilityCache.get(family);
  if (typeof document === "undefined") return true;
  measureCtx ||= document.createElement("canvas").getContext("2d");
  const sample = "mmmmmmmmmmlliWW가나다@#";
  let available = false;
  for (const generic of ["monospace", "serif", "sans-serif"]) {
    measureCtx.font = `72px ${generic}`;
    const baseW = measureCtx.measureText(sample).width;
    measureCtx.font = `72px "${family}", ${generic}`;
    if (measureCtx.measureText(sample).width !== baseW) {
      available = true;
      break;
    }
  }
  availabilityCache.set(family, available);
  return available;
}

/** 지정한 폰트로 렌더링했을 때의 텍스트 너비(pt)를 측정한다. */
export function measureTextWidth(text, font, sizePt) {
  if (typeof document === "undefined") return null;
  measureCtx ||= document.createElement("canvas").getContext("2d");
  // 두 글꼴 모두 이 PC에 없으면 브라우저 기본 글꼴(보통 명조 계열)로 재게 되므로 비슷한 일반 글꼴을 마지막에 둔다
  const generic = font.fallback === "Times New Roman" ? "serif" : font.fallback === "Courier New" ? "monospace" : "sans-serif";
  measureCtx.font = `${font.italic ? "italic " : ""}${font.bold ? "bold " : ""}100px "${font.family}", "${font.fallback}", ${generic}`;
  return (measureCtx.measureText(text).width * sizePt) / 100;
}
