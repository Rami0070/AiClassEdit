export const SPECIAL_KEYWORDS = [
  '특수', '특수반', '특수아동', '특수교육', '도움반', '장애', '특수학급'
];

/**
 * 비고란에서 '분리요청 (사람이름)'을 정밀 파싱합니다.
 * 지원 형식:
 * - 분리요청 (김철수) / 분리요청(김철수)
 * - 분리 (김철수) / 분리(김철수)
 * - 분리요청: 김철수 / 분리: 김철수 / 분리-김철수
 */
export function parseSeparationRequests(rawRemarks: string | undefined | null): string[] {
  if (!rawRemarks) return [];
  const str = String(rawRemarks).trim();
  if (!str) return [];

  const targets: string[] = [];

  // 1) 괄호 패턴: 분리요청(이름) or 분리(이름)
  const regexParen = /분리(?:요청)?\s*[\(\[\{]\s*([가-힣a-zA-Z0-9]+)\s*[\)\]\}]/g;
  let match: RegExpExecArray | null;
  while ((match = regexParen.exec(str)) !== null) {
    const name = match[1].trim();
    if (name && !targets.includes(name)) {
      targets.push(name);
    }
  }

  // 2) 콜론/하이픈 패턴: 분리요청: 이름 or 분리: 이름 or 분리-이름
  const regexColon = /분리(?:요청)?\s*[:\-\s]\s*([가-힣a-zA-Z0-9]+)/g;
  while ((match = regexColon.exec(str)) !== null) {
    const name = match[1].trim();
    if (name && !['요청', '대상', '자', '학생'].includes(name) && !targets.includes(name)) {
      targets.push(name);
    }
  }

  return targets;
}

/**
 * 비고란 파싱: 태그 목록, 특수아동 여부, 분리요청 대상 이름 목록 반환
 */
export function parseRemarks(rawRemarks: string | undefined | null): {
  tags: string[];
  isSpecialChild: boolean;
  separateFromNames: string[];
} {
  if (!rawRemarks) {
    return { tags: [], isSpecialChild: false, separateFromNames: [] };
  }

  const str = String(rawRemarks).trim();
  if (!str) {
    return { tags: [], isSpecialChild: false, separateFromNames: [] };
  }

  const separateFromNames = parseSeparationRequests(str);

  // 구분자로 분리하되, '분리요청(이름)' 형태는 하나의 태그로 온전히 유지
  const rawTokens = str
    .split(/[,;/|\n\r]+/)
    .map(t => t.trim())
    .filter(Boolean);

  const tags: string[] = [];
  let isSpecialChild = false;

  for (const token of rawTokens) {
    // 분리요청이 포함된 토큰이면 분리요청(이름)을 단독 태그로 등록
    if (token.includes('분리')) {
      if (!tags.includes(token)) {
        tags.push(token);
      }
      continue;
    }

    const subTokens = token.split(/\s+/).filter(Boolean);
    for (const sub of subTokens) {
      const cleanSub = sub.replace(/[()\[\]{}]/g, '').trim();
      if (!cleanSub) continue;

      if (SPECIAL_KEYWORDS.some(kw => cleanSub.includes(kw))) {
        isSpecialChild = true;
      }

      if (!tags.includes(cleanSub)) {
        tags.push(cleanSub);
      }
    }
  }

  return { tags, isSpecialChild, separateFromNames };
}

/**
 * 특이사항별 고유 배경/텍스트 색상 팔레트
 * 사용자 지정:
 * - 생활지도: 빨강 (Rose)
 * - 특수아동: 파랑 (Blue)
 * - 다문화: 노랑 (Amber)
 * - 학습부진: 주황 (Orange)
 * - 운동부: 초록 (Emerald)
 * - 분리요청: 보라 (Purple)
 */
export interface TagStyle {
  bg: string;
  text: string;
  border: string;
  hexBg: string;
  hexText: string;
}

export const PRESET_TAG_STYLES: Record<string, TagStyle> = {
  '생활지도': { bg: 'bg-rose-100', text: 'text-rose-800', border: 'border-rose-300', hexBg: '#FFE4E6', hexText: '#9F1239' },
  '특수': { bg: 'bg-blue-100', text: 'text-blue-800', border: 'border-blue-300', hexBg: '#DBEAFE', hexText: '#1E40AF' },
  '특수아동': { bg: 'bg-blue-100', text: 'text-blue-800', border: 'border-blue-300', hexBg: '#DBEAFE', hexText: '#1E40AF' },
  '특수반': { bg: 'bg-blue-100', text: 'text-blue-800', border: 'border-blue-300', hexBg: '#DBEAFE', hexText: '#1E40AF' },
  '다문화': { bg: 'bg-amber-100', text: 'text-amber-800', border: 'border-amber-300', hexBg: '#FEF3C7', hexText: '#92400E' },
  '학습부진': { bg: 'bg-orange-100', text: 'text-orange-800', border: 'border-orange-300', hexBg: '#FFEDD5', hexText: '#9A3412' },
  '운동부': { bg: 'bg-emerald-100', text: 'text-emerald-800', border: 'border-emerald-300', hexBg: '#D1FAE5', hexText: '#065F46' },
  '분리요청': { bg: 'bg-purple-100', text: 'text-purple-800', border: 'border-purple-300', hexBg: '#F3E8FF', hexText: '#6B21A8' }
};

const DYNAMIC_PALETTES: TagStyle[] = [
  { bg: 'bg-teal-100', text: 'text-teal-800', border: 'border-teal-300', hexBg: '#CCFBF1', hexText: '#115E59' },
  { bg: 'bg-indigo-100', text: 'text-indigo-800', border: 'border-indigo-300', hexBg: '#E0E7FF', hexText: '#3730A3' },
  { bg: 'bg-pink-100', text: 'text-pink-800', border: 'border-pink-300', hexBg: '#FCE7F3', hexText: '#9D174D' },
  { bg: 'bg-cyan-100', text: 'text-cyan-800', border: 'border-cyan-300', hexBg: '#CFFAFE', hexText: '#155E75' },
  { bg: 'bg-lime-100', text: 'text-lime-800', border: 'border-lime-300', hexBg: '#ECFCCB', hexText: '#3F6212' },
  { bg: 'bg-fuchsia-100', text: 'text-fuchsia-800', border: 'border-fuchsia-300', hexBg: '#FAE8FF', hexText: '#86198F' },
];

export function getTagColorStyle(tag: string): TagStyle {
  const clean = tag.replace(/[\s\(\)\[\]\{\}:]/g, '');

  if (clean.includes('생활지도')) return PRESET_TAG_STYLES['생활지도'];
  if (clean.includes('특수')) return PRESET_TAG_STYLES['특수아동'];
  if (clean.includes('다문화')) return PRESET_TAG_STYLES['다문화'];
  if (clean.includes('학습부진')) return PRESET_TAG_STYLES['학습부진'];
  if (clean.includes('운동부')) return PRESET_TAG_STYLES['운동부'];
  if (clean.includes('분리')) return PRESET_TAG_STYLES['분리요청'];

  // 기타 태그: 해시 기반 고유 색상 매핑
  let hash = 0;
  for (let i = 0; i < tag.length; i++) {
    hash = (hash << 5) - hash + tag.charCodeAt(i);
    hash |= 0;
  }
  const idx = Math.abs(hash) % DYNAMIC_PALETTES.length;
  return DYNAMIC_PALETTES[idx];
}
