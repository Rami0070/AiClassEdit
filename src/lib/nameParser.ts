// 한국어 2글자 복성 목록
export const COMPOUND_SURNAMES = [
  '남궁', '제갈', '황보', '독고', '사공', '선우', '서문', '동방', '망절', '무본'
];

/**
 * 한국어 이름에서 성(surname)과 이름(firstName)을 분리합니다.
 * 규칙 C: 동일 반 내에 성을 제외한 '이름' 부분이 같은 학생이 절대 함께 배정되지 않아야 함.
 * 예: '김민준'과 '이민준'은 성이 달라도 이름이 '민준'으로 동일하므로 같은 반 배정 불가.
 * 예: '남궁민' -> 성: '남궁', 이름: '민'
 * 예: '김민' -> 성: '김', 이름: '민'
 */
export function parseKoreanName(fullName: string): { surname: string; firstName: string } {
  if (!fullName) return { surname: '', firstName: '' };
  
  // 공백 및 특수문자 제거
  const clean = fullName.replace(/[\s\-_.]/g, '').trim();
  if (clean.length === 0) return { surname: '', firstName: '' };
  if (clean.length === 1) return { surname: '', firstName: clean };

  // 2글자 성씨 확인 (길이가 3글자 이상인 경우)
  if (clean.length >= 3) {
    const twoChars = clean.substring(0, 2);
    if (COMPOUND_SURNAMES.includes(twoChars)) {
      return {
        surname: twoChars,
        firstName: clean.substring(2)
      };
    }
  }

  // 일반적인 단성 (성 1글자 + 이름 나머지)
  // 2글자 이름(예: 강산 -> 성 '강', 이름 '산')
  // 3글자 이름(예: 김민준 -> 성 '김', 이름 '민준')
  // 4글자 이름(예: 박하늘별 -> 성 '박', 이름 '하늘별')
  return {
    surname: clean.substring(0, 1),
    firstName: clean.substring(1)
  };
}

/**
 * 주어진 학생 리스트 내에서 firstName이 중복되는 충돌 쌍/그룹을 검출합니다.
 */
export function findNameCollisions(students: { id: string; name: string; firstName: string }[]): {
  firstName: string;
  count: number;
  studentIds: string[];
  studentNames: string[];
}[] {
  const map = new Map<string, typeof students>();
  for (const s of students) {
    if (!s.firstName) continue;
    const list = map.get(s.firstName) || [];
    list.push(s);
    map.set(s.firstName, list);
  }

  const collisions: {
    firstName: string;
    count: number;
    studentIds: string[];
    studentNames: string[];
  }[] = [];

  for (const [firstName, list] of map.entries()) {
    if (list.length > 1) {
      collisions.push({
        firstName,
        count: list.length,
        studentIds: list.map(s => s.id),
        studentNames: list.map(s => s.name)
      });
    }
  }

  return collisions;
}
