export interface RawStudentInput {
  prevClass: string;     // 반 (기존반)
  prevNumber: number;    // 번호 (기존번호)
  name: string;          // 이름
  gender: '남' | '여' | string; // 성별
  koreanScore?: number;  // 국어
  mathScore?: number;    // 수학
  totalScore?: number;   // 합계
  rank?: number;         // 석차
  targetClass?: string;  // 진급반
  remarks?: string;      // 비고
  id?: string;
}

export interface Student {
  id: string;
  originalId: string;
  prevClass: string;     // 기존 반
  prevNumber: number;    // 기존 번호
  name: string;          // 이름
  surname: string;       // 성 (김)
  firstName: string;     // 이름 (민준)
  gender: '남' | '여';    // 성별
  koreanScore: number;   // 국어
  mathScore: number;     // 수학
  totalScore: number;    // 합계
  rank: number;          // 석차
  score: number;         // 배정 점수
  rankScore: number;     // 정렬용 점수
  rawRemarks: string;    // 비고 원문
  tags: string[];        // 파싱된 태그 목록 (생활지도, 다문화, 특수아동 등)
  isSpecialChild: boolean;// 특수아동 여부 (2명 가중치)
  separateFromNames: string[]; // 분리 요청된 대상 학생 이름 목록 (예: ['김철수'])
  assignedClass?: string;// 진급반
  assignedNumber?: number;// 진급반에서의 새 번호 (남녀 무관 이름순 또는 남녀 분리 이름순)
}

export interface SeparationPair {
  id: string;
  studentA: Student;
  studentB: Student;
}

export interface ClassConfig {
  prevClasses: string[];
  newClasses: string[];
  scoreType: 'score' | 'rank';
  numberingMode?: 'name_all' | 'continuous' | 'continuous_female' | 'separate';
}

export interface ClassSummary {
  className: string;
  students: Student[];
  maleCount: number;
  femaleCount: number;
  specialCount: number;
  totalStudents: number;
  weightedCount: number;
  avgScore: number;
  avgRank: number;
  tagCounts: Record<string, number>;
  prevClassCounts: Record<string, number>;
  nameCollisions: {
    firstName: string;
    students: Student[];
  }[];
  separationViolations: {
    studentA: Student;
    studentB: Student;
  }[];
}

export interface AssignmentResult {
  classes: Record<string, ClassSummary>;
  overallStats: {
    totalStudents: number;
    totalWeighted: number;
    maleCount: number;
    femaleCount: number;
    specialCount: number;
    scoreStdDev: number;
    avgScoreAcrossClasses: number;
    totalCollisions: number;
    totalSeparationViolations: number;
    totalSeparationRequests: number;
    iterations: number;
    elapsedMs: number;
    resolvedStatus: 'PERFECT' | 'SATISFIED_WITH_WARNING' | 'FAILED';
    warnings: string[];
  };
}
