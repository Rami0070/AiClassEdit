import { Student, AssignmentResult, ClassSummary, SeparationPair } from './types';
import { parseKoreanName, findNameCollisions } from './nameParser';
import { parseRemarks } from './remarkParser';

export interface RawStudentRecord {
  id?: string;
  prevClass: string;
  prevNumber?: number | string;
  name: string;
  gender: string;
  koreanScore?: number | string;
  mathScore?: number | string;
  totalScore?: number | string;
  rank?: number | string;
  targetClass?: string;
  remarks?: string;
}

/**
 * 원시 레코드를 전처리하여 표준 Student 객체로 변환
 */
export function preprocessStudents(
  records: RawStudentRecord[],
  scoreType: 'score' | 'rank' = 'rank'
): Student[] {
  const parsed = records.map((r, idx) => {
    const rawName = String(r.name || '').trim();
    const { surname, firstName } = parseKoreanName(rawName);

    const gStr = String(r.gender || '').trim().toLowerCase();
    const gender: '남' | '여' = (gStr === '남' || gStr === '남자' || gStr === 'm' || gStr === 'male') ? '남' : '여';

    const prevClass = String(r.prevClass || '1반').trim();
    const prevNumber = parseInt(String(r.prevNumber || idx + 1), 10) || (idx + 1);

    const kor = parseFloat(String(r.koreanScore || '0')) || 0;
    const math = parseFloat(String(r.mathScore || '0')) || 0;

    let total = parseFloat(String(r.totalScore || '0')) || 0;
    if (total === 0 && (kor > 0 || math > 0)) {
      total = kor + math;
    }

    const rawRank = parseInt(String(r.rank || '0'), 10) || 0;
    const rawRemarks = String(r.remarks || '');
    const { tags, isSpecialChild, separateFromNames } = parseRemarks(rawRemarks);

    return {
      id: r.id || `std_${idx + 1}_${Math.random().toString(36).substring(2, 7)}`,
      originalId: String(idx + 1),
      prevClass,
      prevNumber,
      name: rawName,
      surname,
      firstName,
      gender,
      koreanScore: kor,
      mathScore: math,
      totalScore: total,
      rank: rawRank,
      score: total > 0 ? total : (rawRank > 0 ? -rawRank : 0),
      rawRemarks,
      tags,
      isSpecialChild,
      separateFromNames,
      assignedClass: String(r.targetClass || '').trim() || undefined
    };
  });

  // 석차 보정
  const needComputeRank = parsed.some(s => s.rank === 0);
  if (needComputeRank) {
    const sortedByTotal = [...parsed].sort((a, b) => b.totalScore - a.totalScore);
    sortedByTotal.forEach((s, i) => {
      if (s.rank === 0) s.rank = i + 1;
    });
  }

  return parsed.map(s => ({
    ...s,
    rankScore: scoreType === 'score' ? -s.totalScore : s.rank
  }));
}

/**
 * 학생 목록에서 '분리요청' 제약 조건 쌍을 추출
 * (A가 B를 분리요청했거나, B가 A를 분리요청한 경우)
 */
export function buildSeparationPairs(students: Student[]): { idA: string; idB: string; nameA: string; nameB: string }[] {
  const pairs: { idA: string; idB: string; nameA: string; nameB: string }[] = [];
  const pairSet = new Set<string>();

  for (const s of students) {
    if (!s.separateFromNames || s.separateFromNames.length === 0) continue;

    for (const targetName of s.separateFromNames) {
      // 대상 학생 탐색 (동명이인 고려)
      const targets = students.filter(other => other.id !== s.id && other.name === targetName);
      for (const target of targets) {
        const key = [s.id, target.id].sort().join('___');
        if (!pairSet.has(key)) {
          pairSet.add(key);
          pairs.push({
            idA: s.id,
            idB: target.id,
            nameA: s.name,
            nameB: target.name
          });
        }
      }
    }
  }

  return pairs;
}

/**
 * 특정 반 내에서 분리요청 위반 검출
 */
export function findSeparationViolations(
  students: Student[],
  pairs: { idA: string; idB: string; nameA: string; nameB: string }[]
): { studentA: Student; studentB: Student }[] {
  const idMap = new Map(students.map(s => [s.id, s]));
  const violations: { studentA: Student; studentB: Student }[] = [];

  for (const pair of pairs) {
    const sA = idMap.get(pair.idA);
    const sB = idMap.get(pair.idB);
    if (sA && sB) {
      violations.push({ studentA: sA, studentB: sB });
    }
  }

  return violations;
}

/**
 * 규칙 A: 성적 기반 스네이크 드래프트 1차 배정
 */
export function initialSnakeDraft(
  students: Student[],
  newClassNames: string[],
  prevClassNames: string[]
): Map<string, Student[]> {
  const K = newClassNames.length;
  const assignment = new Map<string, Student[]>();
  for (const name of newClassNames) {
    assignment.set(name, []);
  }

  const existingClasses = prevClassNames.length > 0
    ? prevClassNames
    : Array.from(new Set(students.map(s => s.prevClass)));

  const grouped = new Map<string, Student[]>();
  for (const s of students) {
    const key = `${s.prevClass}___${s.gender}`;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key)!.push(s);
  }

  for (const [_, list] of grouped.entries()) {
    list.sort((a, b) => a.rankScore - b.rankScore);
  }

  existingClasses.forEach((prevCls, cIdx) => {
    (['남', '여'] as const).forEach(gender => {
      const key = `${prevCls}___${gender}`;
      const groupStudents = grouped.get(key) || [];
      if (groupStudents.length === 0) return;

      const baseOffset = (cIdx + (gender === '남' ? 0 : 1)) % K;

      groupStudents.forEach((student, rankIdx) => {
        const round = Math.floor(rankIdx / K);
        const step = rankIdx % K;

        let targetSlot: number;
        if (round % 2 === 0) {
          targetSlot = (baseOffset + step) % K;
        } else {
          targetSlot = (baseOffset + (K - 1 - step)) % K;
        }

        const targetClassName = newClassNames[targetSlot];
        student.assignedClass = targetClassName;
        assignment.get(targetClassName)!.push(student);
      });
    });
  });

  return assignment;
}

/**
 * 진급반 내 번호 배정: 남녀 무관 이름순(기본) 또는 남녀 분리 모드
 */
export function assignClassNumbers(
  students: Student[],
  mode: 'name_all' | 'continuous' | 'continuous_female' | 'separate' = 'name_all'
): Student[] {
  // 1. 남녀 구분 없이 전체 이름순(가나다순)으로 1번부터 배정 (기본/표준 모드)
  if (mode === 'name_all' || !mode) {
    const sorted = [...students].sort((a, b) => a.name.localeCompare(b.name, 'ko'));
    sorted.forEach((s, idx) => {
      s.assignedNumber = idx + 1;
    });
    return sorted;
  }

  // 2. 남녀 분리 연속번호 (남 1번~ / 여 21번~)
  if (mode === 'continuous') {
    const boys = students.filter(s => s.gender === '남').sort((a, b) => a.name.localeCompare(b.name, 'ko'));
    const girls = students.filter(s => s.gender === '여').sort((a, b) => a.name.localeCompare(b.name, 'ko'));

    boys.forEach((s, idx) => { s.assignedNumber = idx + 1; });
    const girlStart = Math.max(21, boys.length + 1);
    girls.forEach((s, idx) => { s.assignedNumber = girlStart + idx; });
    return [...boys, ...girls].sort((a, b) => (a.assignedNumber || 0) - (b.assignedNumber || 0));
  }

  // 2-2. 남녀 분리 연속번호 (여 1번~ / 남 21번~)
  if (mode === 'continuous_female') {
    const boys = students.filter(s => s.gender === '남').sort((a, b) => a.name.localeCompare(b.name, 'ko'));
    const girls = students.filter(s => s.gender === '여').sort((a, b) => a.name.localeCompare(b.name, 'ko'));

    girls.forEach((s, idx) => { s.assignedNumber = idx + 1; });
    const boyStart = Math.max(21, girls.length + 1);
    boys.forEach((s, idx) => { s.assignedNumber = boyStart + idx; });
    return [...girls, ...boys].sort((a, b) => (a.assignedNumber || 0) - (b.assignedNumber || 0));
  }

  // 3. 남녀 분리 각각 1번부터 (남 1번~, 여 1번~)
  const boys = students.filter(s => s.gender === '남').sort((a, b) => a.name.localeCompare(b.name, 'ko'));
  const girls = students.filter(s => s.gender === '여').sort((a, b) => a.name.localeCompare(b.name, 'ko'));

  boys.forEach((s, idx) => { s.assignedNumber = idx + 1; });
  girls.forEach((s, idx) => { s.assignedNumber = idx + 1; });
  return [...boys, ...girls].sort((a, b) => (a.assignedNumber || 0) - (b.assignedNumber || 0));
}

/**
 * 반별 요약 계산
 */
export function calculateClassSummary(
  className: string,
  students: Student[],
  separationPairs: { idA: string; idB: string; nameA: string; nameB: string }[],
  numberingMode: 'name_all' | 'continuous' | 'continuous_female' | 'separate' = 'name_all'
): ClassSummary {
  const numberedStudents = assignClassNumbers(students, numberingMode);

  let maleCount = 0;
  let femaleCount = 0;
  let specialCount = 0;
  let totalScore = 0;
  let totalRank = 0;
  const tagCounts: Record<string, number> = {};
  const prevClassCounts: Record<string, number> = {};

  for (const s of numberedStudents) {
    if (s.gender === '남') maleCount++;
    else femaleCount++;

    if (s.isSpecialChild) specialCount++;

    totalScore += s.totalScore || s.score;
    totalRank += s.rank || 0;

    prevClassCounts[s.prevClass] = (prevClassCounts[s.prevClass] || 0) + 1;

    for (const tag of s.tags) {
      tagCounts[tag] = (tagCounts[tag] || 0) + 1;
    }
  }

  const weightedCount = numberedStudents.length + specialCount;
  const avgScore = numberedStudents.length > 0 ? Number((totalScore / numberedStudents.length).toFixed(2)) : 0;
  const avgRank = numberedStudents.length > 0 ? Number((totalRank / numberedStudents.length).toFixed(1)) : 0;

  const collisionsRaw = findNameCollisions(numberedStudents);
  const nameCollisions = collisionsRaw.map(c => ({
    firstName: c.firstName,
    students: numberedStudents.filter(s => c.studentIds.includes(s.id))
  }));

  const separationViolations = findSeparationViolations(numberedStudents, separationPairs);

  return {
    className,
    students: numberedStudents,
    maleCount,
    femaleCount,
    specialCount,
    totalStudents: numberedStudents.length,
    weightedCount,
    avgScore,
    avgRank,
    tagCounts,
    prevClassCounts,
    nameCollisions,
    separationViolations
  };
}

/**
 * 목적 함수: 비용 계산
 * - 동일 이름 충돌: 50만점 페널티
 * - 분리 요청 위반: 100만점 페널티 (절대 불가 하드 제약!)
 */
function evaluateAssignmentCost(
  assignment: Map<string, Student[]>,
  targetWeightedPerClass: number,
  targetMalePerClass: number,
  targetFemalePerClass: number,
  allUniqueTags: string[],
  separationPairs: { idA: string; idB: string; nameA: string; nameB: string }[]
) {
  let nameCollisionCost = 0;
  let separationViolationCost = 0;
  let weightedCost = 0;
  let genderCost = 0;
  let tagCost = 0;
  let scoreCost = 0;
  let collisionCount = 0;
  let separationViolationCount = 0;

  const classList = Array.from(assignment.values());
  const classAvgRanks: number[] = [];

  for (const students of classList) {
    // 1. 규칙 C: 이름 충돌 (500,000점)
    const collisions = findNameCollisions(students);
    for (const c of collisions) {
      const pairs = (c.count * (c.count - 1)) / 2;
      nameCollisionCost += pairs * 500000;
      collisionCount += pairs;
    }

    // 2. 분리 요청 위반 (1,000,000점)
    const sepViolations = findSeparationViolations(students, separationPairs);
    if (sepViolations.length > 0) {
      separationViolationCost += sepViolations.length * 1000000;
      separationViolationCount += sepViolations.length;
    }

    // 3. 규칙 F: 가중 정원
    const specialCount = students.filter(s => s.isSpecialChild).length;
    const weightedCount = students.length + specialCount;
    const wDiff = weightedCount - targetWeightedPerClass;
    weightedCost += Math.pow(wDiff, 2) * 20000;

    // 4. 규칙 B: 남녀 성비
    const maleCount = students.filter(s => s.gender === '남').length;
    const femaleCount = students.length - maleCount;
    genderCost += Math.pow(maleCount - targetMalePerClass, 2) * 6000;
    genderCost += Math.pow(femaleCount - targetFemalePerClass, 2) * 6000;

    const totalR = students.reduce((acc, s) => acc + (s.rank || 0), 0);
    classAvgRanks.push(students.length > 0 ? totalR / students.length : 0);
  }

  // 5. 비고 태그 분산
  for (const tag of allUniqueTags) {
    const counts = classList.map(stds => stds.filter(s => s.tags.includes(tag)).length);
    const mean = counts.reduce((a, b) => a + b, 0) / counts.length;
    const variance = counts.reduce((acc, c) => acc + Math.pow(c - mean, 2), 0) / counts.length;
    tagCost += variance * 2500;
  }

  // 6. 석차 분산
  if (classAvgRanks.length > 1) {
    const meanRank = classAvgRanks.reduce((a, b) => a + b, 0) / classAvgRanks.length;
    const rankVar = classAvgRanks.reduce((acc, s) => acc + Math.pow(s - meanRank, 2), 0) / classAvgRanks.length;
    scoreCost = rankVar * 400;
  }

  return {
    totalCost: nameCollisionCost + separationViolationCost + weightedCost + genderCost + tagCost + scoreCost,
    collisionCount,
    separationViolationCount,
    nameCollisionCost,
    separationViolationCost,
    weightedCost
  };
}

/**
 * 분리 요청 위반 즉시 해소 스왑
 */
function resolveSeparationViolationsDirectly(
  assignment: Map<string, Student[]>,
  newClassNames: string[],
  separationPairs: { idA: string; idB: string; nameA: string; nameB: string }[]
): boolean {
  let anyImprovement = false;

  for (let i = 0; i < newClassNames.length; i++) {
    const cNameA = newClassNames[i];
    const listA = assignment.get(cNameA)!;
    const violations = findSeparationViolations(listA, separationPairs);

    for (const v of violations) {
      // studentA를 다른 반으로 보낼 수 있는지 탐색
      const targetStudent = v.studentA;
      let swapped = false;

      for (let j = 0; j < newClassNames.length; j++) {
        if (i === j) continue;
        const cNameB = newClassNames[j];
        const listB = assignment.get(cNameB)!;

        // 반 B에 targetStudent와 분리되어야 할 대상이 있는지 검사
        const hasConflictInB = separationPairs.some(p => 
          (p.idA === targetStudent.id && listB.some(s => s.id === p.idB)) ||
          (p.idB === targetStudent.id && listB.some(s => s.id === p.idA))
        );
        if (hasConflictInB) continue;

        // 반 B에서 반 A로 보낼 동일 성별 학생 찾기
        const candidateIndex = listB.findIndex(sB => {
          if (sB.gender !== targetStudent.gender) return false;
          // sB가 반 A로 갔을 때 분리 위반이 없어야 함
          const sBConflictInA = separationPairs.some(p =>
            (p.idA === sB.id && listA.some(s => s.id !== targetStudent.id && s.id === p.idB)) ||
            (p.idB === sB.id && listA.some(s => s.id !== targetStudent.id && s.id === p.idA))
          );
          if (sBConflictInA) return false;

          // 이름 충돌도 없어야 함
          const nameConflictInA = listA.some(sA => sA.id !== targetStudent.id && sA.firstName === sB.firstName);
          const nameConflictInB = listB.some(sB2 => sB2.id !== sB.id && sB2.firstName === targetStudent.firstName);
          return !nameConflictInA && !nameConflictInB;
        });

        if (candidateIndex !== -1) {
          const partnerStudent = listB[candidateIndex];
          const idxInA = listA.findIndex(s => s.id === targetStudent.id);

          listA[idxInA] = partnerStudent;
          listB[candidateIndex] = targetStudent;
          partnerStudent.assignedClass = cNameA;
          targetStudent.assignedClass = cNameB;

          swapped = true;
          anyImprovement = true;
          break;
        }
      }

      if (swapped) break;
    }
  }

  return anyImprovement;
}

function resolveNameCollisionsDirectly(
  assignment: Map<string, Student[]>,
  newClassNames: string[]
): boolean {
  let anyImprovement = false;

  for (let i = 0; i < newClassNames.length; i++) {
    const cNameA = newClassNames[i];
    const listA = assignment.get(cNameA)!;
    const collisions = findNameCollisions(listA);

    for (const collision of collisions) {
      const conflictStudent = listA.find(s => s.firstName === collision.firstName);
      if (!conflictStudent) continue;

      let swapped = false;
      for (let j = 0; j < newClassNames.length; j++) {
        if (i === j) continue;
        const cNameB = newClassNames[j];
        const listB = assignment.get(cNameB)!;

        if (listB.some(s => s.firstName === conflictStudent.firstName)) continue;

        const candidateIndex = listB.findIndex(sB => {
          if (sB.gender !== conflictStudent.gender) return false;
          return !listA.some(sA => sA.id !== conflictStudent.id && sA.firstName === sB.firstName);
        });

        if (candidateIndex !== -1) {
          const partnerStudent = listB[candidateIndex];
          const idxInA = listA.findIndex(s => s.id === conflictStudent.id);
          listA[idxInA] = partnerStudent;
          listB[candidateIndex] = conflictStudent;
          partnerStudent.assignedClass = cNameA;
          conflictStudent.assignedClass = cNameB;
          swapped = true;
          anyImprovement = true;
          break;
        }
      }
      if (swapped) break;
    }
  }

  return anyImprovement;
}

function balanceWeightedCapacities(
  assignment: Map<string, Student[]>,
  newClassNames: string[]
): boolean {
  let changed = false;

  const getWeighted = (cName: string) => {
    const list = assignment.get(cName)!;
    const special = list.filter(s => s.isSpecialChild).length;
    return list.length + special;
  };

  for (let step = 0; step < 10; step++) {
    let maxClass = newClassNames[0];
    let minClass = newClassNames[0];
    let maxW = getWeighted(maxClass);
    let minW = getWeighted(minClass);

    for (const c of newClassNames) {
      const w = getWeighted(c);
      if (w > maxW) { maxW = w; maxClass = c; }
      if (w < minW) { minW = w; minClass = c; }
    }

    if (maxW - minW <= 1) break;

    const listMax = assignment.get(maxClass)!;
    const listMin = assignment.get(minClass)!;

    const candidateIdx = listMax.findIndex(s => {
      if (s.isSpecialChild) return false;
      return !listMin.some(m => m.firstName === s.firstName);
    });

    if (candidateIdx !== -1) {
      const student = listMax.splice(candidateIdx, 1)[0];
      student.assignedClass = minClass;
      listMin.push(student);
      changed = true;
    } else {
      break;
    }
  }

  return changed;
}

/**
 * 스마트 학급편성 종합 최적화 실행
 */
export function optimizeAssignment(
  students: Student[],
  newClassNames: string[],
  prevClassNames: string[],
  options: {
    maxIterations?: number;
    initialTemp?: number;
    coolingRate?: number;
    maxRestarts?: number;
    numberingMode?: 'name_all' | 'continuous' | 'continuous_female' | 'separate';
  } = {}
): AssignmentResult {
  const startTime = performance.now();
  const K = newClassNames.length;
  const maxIterations = options.maxIterations || 2000;
  const initialTemp = options.initialTemp || 100.0;
  const coolingRate = options.coolingRate || 0.995;
  const maxRestarts = options.maxRestarts || 2;
  const numberingMode = options.numberingMode || 'name_all';

  // 분리 요청 쌍 사전 빌드
  const separationPairs = buildSeparationPairs(students);

  const totalSpecial = students.filter(s => s.isSpecialChild).length;
  const totalWeighted = students.length + totalSpecial;
  const targetWeightedPerClass = totalWeighted / K;

  const totalMale = students.filter(s => s.gender === '남').length;
  const totalFemale = students.length - totalMale;
  const targetMalePerClass = totalMale / K;
  const targetFemalePerClass = totalFemale / K;

  const allUniqueTags = Array.from(new Set(students.flatMap(s => s.tags)));

  let globalBestAssignment: Map<string, Student[]> | null = null;
  let globalLowestCost = Infinity;
  let totalIterationsDone = 0;

  for (let restart = 0; restart < maxRestarts; restart++) {
    let currentAssignment = initialSnakeDraft(students, newClassNames, prevClassNames);

    if (restart > 0) {
      for (let p = 0; p < 20; p++) {
        const c1 = newClassNames[Math.floor(Math.random() * K)];
        const c2 = newClassNames[Math.floor(Math.random() * K)];
        if (c1 === c2) continue;
        const l1 = currentAssignment.get(c1)!;
        const l2 = currentAssignment.get(c2)!;
        if (l1.length === 0 || l2.length === 0) continue;
        const s1 = l1[Math.floor(Math.random() * l1.length)];
        const s2Candidates = l2.filter(s => s.gender === s1.gender);
        if (s2Candidates.length > 0) {
          const s2 = s2Candidates[Math.floor(Math.random() * s2Candidates.length)];
          const i1 = l1.indexOf(s1);
          const i2 = l2.indexOf(s2);
          l1[i1] = s2;
          l2[i2] = s1;
          s1.assignedClass = c2;
          s2.assignedClass = c1;
        }
      }
    }

    // 1. 이름 충돌 및 분리 요청 즉시 해소
    for (let c = 0; c < 5; c++) {
      resolveNameCollisionsDirectly(currentAssignment, newClassNames);
      resolveSeparationViolationsDirectly(currentAssignment, newClassNames, separationPairs);
    }

    balanceWeightedCapacities(currentAssignment, newClassNames);

    let currentCostObj = evaluateAssignmentCost(
      currentAssignment,
      targetWeightedPerClass,
      targetMalePerClass,
      targetFemalePerClass,
      allUniqueTags,
      separationPairs
    );

    let localBestCost = currentCostObj.totalCost;
    let localBestAssignment = new Map<string, Student[]>();
    for (const [k, v] of currentAssignment.entries()) {
      localBestAssignment.set(k, [...v]);
    }

    let temp = initialTemp;

    for (let iter = 0; iter < maxIterations; iter++) {
      totalIterationsDone++;

      const c1Name = newClassNames[Math.floor(Math.random() * K)];
      let c2Name = newClassNames[Math.floor(Math.random() * K)];
      while (c1Name === c2Name) {
        c2Name = newClassNames[Math.floor(Math.random() * K)];
      }

      const list1 = currentAssignment.get(c1Name)!;
      const list2 = currentAssignment.get(c2Name)!;
      if (list1.length === 0 || list2.length === 0) continue;

      // 이름 충돌 또는 분리 위반이 있는 학생 우선 탐색
      let s1: Student;
      const sepV1 = findSeparationViolations(list1, separationPairs);
      const col1 = findNameCollisions(list1);

      if (sepV1.length > 0 && Math.random() < 0.8) {
        s1 = sepV1[0].studentA;
      } else if (col1.length > 0 && Math.random() < 0.8) {
        const conflictFirst = col1[0].firstName;
        s1 = list1.find(s => s.firstName === conflictFirst) || list1[Math.floor(Math.random() * list1.length)];
      } else {
        s1 = list1[Math.floor(Math.random() * list1.length)];
      }

      const sameGenderOnly = Math.random() < 0.9;
      let s2Candidates = sameGenderOnly ? list2.filter(s => s.gender === s1.gender) : list2;
      if (s2Candidates.length === 0) s2Candidates = list2;

      const s2 = s2Candidates[Math.floor(Math.random() * s2Candidates.length)];

      const idx1 = list1.indexOf(s1);
      const idx2 = list2.indexOf(s2);
      list1[idx1] = s2;
      list2[idx2] = s1;
      s1.assignedClass = c2Name;
      s2.assignedClass = c1Name;

      const newCostObj = evaluateAssignmentCost(
        currentAssignment,
        targetWeightedPerClass,
        targetMalePerClass,
        targetFemalePerClass,
        allUniqueTags,
        separationPairs
      );

      const delta = newCostObj.totalCost - currentCostObj.totalCost;

      let accept = false;
      if (delta <= 0) {
        accept = true;
      } else {
        const prob = Math.exp(-delta / Math.max(temp, 0.0001));
        if (Math.random() < prob) accept = true;
      }

      if (accept) {
        currentCostObj = newCostObj;
        if (currentCostObj.totalCost < localBestCost) {
          localBestCost = currentCostObj.totalCost;
          localBestAssignment.clear();
          for (const [k, v] of currentAssignment.entries()) {
            localBestAssignment.set(k, [...v]);
          }

          if (currentCostObj.collisionCount === 0 && currentCostObj.separationViolationCount === 0 && currentCostObj.weightedCost === 0) {
            temp *= 0.99;
          }
        }
      } else {
        list1[idx1] = s1;
        list2[idx2] = s2;
        s1.assignedClass = c1Name;
        s2.assignedClass = c2Name;
      }

      temp *= coolingRate;

      if (iter % 800 === 0) {
        if (currentCostObj.separationViolationCount > 0) {
          resolveSeparationViolationsDirectly(currentAssignment, newClassNames, separationPairs);
        }
        if (currentCostObj.collisionCount > 0) {
          resolveNameCollisionsDirectly(currentAssignment, newClassNames);
        }
      }
    }

    if (localBestCost < globalLowestCost) {
      globalLowestCost = localBestCost;
      globalBestAssignment = localBestAssignment;

      const checkCollisions = evaluateAssignmentCost(
        localBestAssignment,
        targetWeightedPerClass,
        targetMalePerClass,
        targetFemalePerClass,
        allUniqueTags,
        separationPairs
      );

      if (checkCollisions.collisionCount === 0 && checkCollisions.separationViolationCount === 0) {
        break;
      }
    }
  }

  const bestAssignment = globalBestAssignment || initialSnakeDraft(students, newClassNames, prevClassNames);

  const classesSummary: Record<string, ClassSummary> = {};
  let totalCollisions = 0;
  let totalSeparationViolations = 0;
  const warnings: string[] = [];

  for (const className of newClassNames) {
    const rawList = bestAssignment.get(className) || [];
    for (const s of rawList) s.assignedClass = className;

    const summary = calculateClassSummary(className, rawList, separationPairs, numberingMode);
    classesSummary[className] = summary;

    for (const col of summary.nameCollisions) {
      totalCollisions += (col.students.length * (col.students.length - 1)) / 2;
      warnings.push(`[${className}] 동일 이름 충돌: '${col.firstName}' (${col.students.map(s => s.name).join(', ')})`);
    }

    for (const sep of summary.separationViolations) {
      totalSeparationViolations++;
      warnings.push(`[${className}] 분리 요청 위반: '${sep.studentA.name}' ↔ '${sep.studentB.name}' 동반 배정`);
    }
  }

  const allSummaries = Object.values(classesSummary);
  const avgScores = allSummaries.map(c => c.avgScore);
  const meanAvgScore = avgScores.length > 0 ? avgScores.reduce((a, b) => a + b, 0) / avgScores.length : 0;
  const scoreVariance = avgScores.length > 0 ? avgScores.reduce((acc, s) => acc + Math.pow(s - meanAvgScore, 2), 0) / avgScores.length : 0;
  const scoreStdDev = Number(Math.sqrt(scoreVariance).toFixed(2));

  const elapsedMs = Number((performance.now() - startTime).toFixed(1));
  const resolvedStatus = (totalCollisions === 0 && totalSeparationViolations === 0)
    ? 'PERFECT'
    : ((totalCollisions + totalSeparationViolations) <= 2 ? 'SATISFIED_WITH_WARNING' : 'FAILED');

  if (totalSeparationViolations > 0) {
    warnings.unshift(`경고: 분리 요청된 학생 간 동반 배정이 ${totalSeparationViolations}건 발생했습니다.`);
  }

  return {
    classes: classesSummary,
    overallStats: {
      totalStudents: students.length,
      totalWeighted,
      maleCount: totalMale,
      femaleCount: totalFemale,
      specialCount: totalSpecial,
      scoreStdDev,
      avgScoreAcrossClasses: Number(meanAvgScore.toFixed(2)),
      totalCollisions,
      totalSeparationViolations,
      totalSeparationRequests: separationPairs.length,
      iterations: totalIterationsDone,
      elapsedMs,
      resolvedStatus,
      warnings
    }
  };
}
