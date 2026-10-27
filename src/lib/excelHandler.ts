import { RawStudentRecord } from './optimizer';
import { AssignmentResult } from './types';

declare const XLSX: any;

export const UPLOAD_STANDARD_HEADERS = [
  '반', '번호', '이름', '성별', '국어', '수학', '합계', '석차', '비고'
];

export const RESULT_STANDARD_HEADERS = [
  '반', '번호', '이름', '성별', '석차', '진급반', '번호', '비고'
];

export interface SheetParseDetail {
  name: string;
  count: number;
}

/**
 * 특이사항별 엑셀 RGB 헥스 컬러 정의 (6자리 Hex, # 제외)
 * - 생활지도: 빨강 (Rose/Red)
 * - 특수/특수아동: 파랑 (Blue)
 * - 다문화: 노랑 (Amber/Yellow)
 * - 학습부진: 주황 (Orange)
 * - 운동부: 초록 (Emerald/Green)
 * - 분리요청: 보라 (Purple/Violet)
 */
export function getExcelColorForRemarks(remarks: string | undefined | null): {
  bgRgb: string;
  textRgb: string;
} | null {
  if (!remarks) return null;
  const str = String(remarks).trim();
  if (!str || str === '-' || str === '없음' || str === '해당없음') return null;

  if (str.includes('생활지도')) {
    return { bgRgb: 'FFD1D5', textRgb: '9F1239' }; // 빨강 계열
  }
  if (str.includes('특수')) {
    return { bgRgb: 'BFDBFE', textRgb: '1E40AF' }; // 파랑 계열
  }
  if (str.includes('다문화')) {
    return { bgRgb: 'FEF08A', textRgb: '854D0E' }; // 노랑 계열
  }
  if (str.includes('학습부진')) {
    return { bgRgb: 'FED7AA', textRgb: '9A3412' }; // 주황 계열
  }
  if (str.includes('운동부')) {
    return { bgRgb: 'BBF7D0', textRgb: '166534' }; // 초록 계열
  }
  if (str.includes('분리')) {
    return { bgRgb: 'E9D5FF', textRgb: '6B21A8' }; // 보라 계열
  }
  if (str.includes('영재') || str.includes('우수')) {
    return { bgRgb: 'E0E7FF', textRgb: '3730A3' }; // 인디고
  }
  if (str.includes('건강') || str.includes('보건') || str.includes('질환')) {
    return { bgRgb: 'FCE7F3', textRgb: '9D174D' }; // 핑크
  }

  // 기타 특이사항은 부드러운 청록
  return { bgRgb: 'CCFBF1', textRgb: '115E59' };
}

/**
 * 워크시트의 셀들에 스타일(배경색, 폰트, 테두리, 열너비)을 적용합니다.
 */
function styleWorksheet(worksheet: any, startRowIdx: number = 0): void {
  const xlsxLib = typeof window !== 'undefined' 
    ? ((window as any).XLSXStyle || (window as any).XLSX) 
    : (globalThis as any).XLSX;

  if (!worksheet || !worksheet['!ref']) return;

  const range = xlsxLib.utils.decode_range(worksheet['!ref']);

  // 1. 헤더 행 분석: 비고 컬럼, 진급반 컬럼, 새번호 컬럼 등 동적 탐색
  let remarksColIdx = -1;
  let targetClassColIdx = -1;
  let newNumberColIdx = -1;
  let isStatsSheet = false;
  const headerNames: string[] = [];

  for (let C = range.s.c; C <= range.e.c; C++) {
    const addr = xlsxLib.utils.encode_cell({ r: startRowIdx, c: C });
    const cell = worksheet[addr];
    const val = cell ? String(cell.v || '').trim() : '';
    headerNames[C] = val;

    if (val.includes('비고')) {
      remarksColIdx = C;
    }
    if (val === '진급반') {
      targetClassColIdx = C;
    }
    if (val === '번호' && C > 2) {
      newNumberColIdx = C;
    }
    if (val.includes('실제 학생 수') || val.includes('통계') || val.includes('가중 학생 수')) {
      isStatsSheet = true;
    }
  }

  // 2. 열 너비 자동 최적화
  const cols: Array<{ wch: number }> = [];
  for (let C = range.s.c; C <= range.e.c; C++) {
    const headerName = headerNames[C] || '';
    if (isStatsSheet) {
      if (headerName.includes('가중')) cols.push({ wch: 22 });
      else if (headerName.includes('성비') || headerName.includes('충돌') || headerName.includes('위반')) cols.push({ wch: 16 });
      else cols.push({ wch: 14 });
    } else {
      if (C === remarksColIdx) cols.push({ wch: 34 });
      else if (headerName === '이름') cols.push({ wch: 14 });
      else if (headerName === '진급반') cols.push({ wch: 12 });
      else if (headerName === '번호') cols.push({ wch: 9 });
      else if (headerName === '성별' || headerName === '반') cols.push({ wch: 9 });
      else cols.push({ wch: 11 });
    }
  }
  worksheet['!cols'] = cols;

  const thinBorder = {
    top: { style: 'thin', color: { rgb: 'CBD5E1' } },
    bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
    left: { style: 'thin', color: { rgb: 'CBD5E1' } },
    right: { style: 'thin', color: { rgb: 'CBD5E1' } }
  };

  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = xlsxLib.utils.encode_cell({ r: R, c: C });
      const cell = worksheet[cellAddress];
      if (!cell) continue;

      // 타이틀 행 (헤더 이전 상단 안내 영역)
      if (R < startRowIdx) {
        if (R === 0 && cell.v) {
          cell.s = {
            font: { name: '맑은 고딕', sz: 13, bold: true, color: { rgb: '166534' } },
            alignment: { vertical: 'center', horizontal: 'left' }
          };
        } else if (R === 1 && cell.v) {
          cell.s = {
            font: { name: '맑은 고딕', sz: 9.5, italic: true, color: { rgb: '475569' } },
            alignment: { vertical: 'center', horizontal: 'left' }
          };
        }
        continue;
      }

      // 1. 헤더 행 스타일링
      if (R === startRowIdx) {
        cell.s = {
          fill: {
            patternType: 'solid',
            fgColor: { rgb: isStatsSheet ? 'E2E8F0' : 'DCFCE7' } // 통계는 슬레이트, 학생은 파스텔 민트
          },
          font: {
            name: '맑은 고딕',
            sz: 10,
            bold: true,
            color: { rgb: isStatsSheet ? '1E293B' : '166534' }
          },
          alignment: {
            vertical: 'center',
            horizontal: 'center'
          },
          border: {
            top: { style: 'medium', color: { rgb: isStatsSheet ? '94A3B8' : '86EFAC' } },
            bottom: { style: 'medium', color: { rgb: isStatsSheet ? '94A3B8' : '86EFAC' } },
            left: { style: 'thin', color: { rgb: 'CBD5E1' } },
            right: { style: 'thin', color: { rgb: 'CBD5E1' } }
          }
        };
      } 
      // 2. 데이터 행 스타일링
      else if (R > startRowIdx) {
        if (isStatsSheet) {
          const valStr = String(cell.v || '');
          const isWarning = (valStr.includes('충돌') || valStr.includes('위반')) && !valStr.startsWith('0');
          const isSafe = valStr.includes('0건') || valStr.includes('안전');

          cell.s = {
            fill: isWarning ? { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } } : undefined,
            font: {
              name: '맑은 고딕',
              sz: 10,
              bold: isWarning || isSafe || C === 0,
              color: { rgb: isWarning ? '991B1B' : (isSafe ? '166534' : '1E293B') }
            },
            alignment: {
              vertical: 'center',
              horizontal: 'center'
            },
            border: thinBorder
          };
        } else {
          const isCenterCol = (C !== remarksColIdx);

          let cellStyle: any = {
            font: {
              name: '맑은 고딕',
              sz: 10,
              color: { rgb: '1E293B' }
            },
            alignment: {
              vertical: 'center',
              horizontal: isCenterCol ? 'center' : 'left'
            },
            border: thinBorder
          };

          // 비고란에 특이사항이 있는 경우 셀 배경색 적용!
          if (C === remarksColIdx && cell.v) {
            const colorInfo = getExcelColorForRemarks(cell.v);
            if (colorInfo) {
              cellStyle.fill = {
                patternType: 'solid',
                fgColor: { rgb: colorInfo.bgRgb } // 생활지도: 빨강, 특수: 파랑, 다문화: 노랑 등
              };
              cellStyle.font = {
                name: '맑은 고딕',
                sz: 10,
                bold: true,
                color: { rgb: colorInfo.textRgb }
              };
              cellStyle.border = {
                top: { style: 'thin', color: { rgb: colorInfo.textRgb } },
                bottom: { style: 'thin', color: { rgb: colorInfo.textRgb } },
                left: { style: 'thin', color: { rgb: colorInfo.textRgb } },
                right: { style: 'thin', color: { rgb: colorInfo.textRgb } }
              };
            }
          }

          // 진급반 컬럼 강조
          if (C === targetClassColIdx) {
            cellStyle.font = {
              name: '맑은 고딕',
              sz: 10,
              bold: true,
              color: { rgb: '065F46' }
            };
            cellStyle.fill = {
              patternType: 'solid',
              fgColor: { rgb: 'F0FDF4' }
            };
          }

          // 새 번호 컬럼 강조
          if (C === newNumberColIdx) {
            cellStyle.font = {
              name: '맑은 고딕',
              sz: 10,
              bold: true,
              color: { rgb: '047857' }
            };
            cellStyle.fill = {
              patternType: 'solid',
              fgColor: { rgb: 'F0FDF4' }
            };
          }

          cell.s = cellStyle;
        }
      }
    }
  }
}

export async function parseExcelFile(file: File): Promise<{
  records: RawStudentRecord[];
  headers: string[];
  detectedClasses: string[];
  sheetDetails: SheetParseDetail[];
}> {
  const xlsxLib = typeof window !== 'undefined' 
    ? ((window as any).XLSXStyle || (window as any).XLSX) 
    : (globalThis as any).XLSX;

  if (!xlsxLib) {
    throw new Error('엑셀 파싱 라이브러리를 찾을 수 없습니다.');
  }

  const data = await file.arrayBuffer();
  const workbook = xlsxLib.read(data, { type: 'array' });

  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error('엑셀 파일에 시트가 존재하지 않습니다.');
  }

  const records: RawStudentRecord[] = [];
  const detectedClassesSet = new Set<string>();
  const sheetDetails: SheetParseDetail[] = [];
  let commonHeaders: string[] = [];

  const norm = (str: string) => String(str || '').toLowerCase().replace(/[\s\-_]/g, '');

  const findColIndex = (headers: string[], synonyms: string[]) => {
    for (const syn of synonyms) {
      const idx = headers.findIndex(h => norm(h) === norm(syn));
      if (idx !== -1) return idx;
    }
    for (const syn of synonyms) {
      const idx = headers.findIndex(h => norm(h).includes(norm(syn)));
      if (idx !== -1) return idx;
    }
    return -1;
  };

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) continue;

    const rows: any[] = xlsxLib.utils.sheet_to_json(worksheet, { header: 1 });
    if (!rows || rows.length < 2) continue;

    let headerRowIdx = -1;
    for (let r = 0; r < Math.min(rows.length, 6); r++) {
      const row = rows[r];
      if (Array.isArray(row)) {
        const rowStr = row.map(c => norm(String(c || ''))).join(' ');
        if (rowStr.includes('이름') || rowStr.includes('성명') || rowStr.includes('학생명') || (rowStr.includes('반') && rowStr.includes('성별'))) {
          headerRowIdx = r;
          break;
        }
      }
    }

    if (headerRowIdx === -1) headerRowIdx = 0;

    const headers = (rows[headerRowIdx] as any[]).map(h => String(h || '').trim());
    if (commonHeaders.length === 0) commonHeaders = headers;

    const prevClassIdx = findColIndex(headers, ['반', '기존반', '현재반', '학급', '소속반', 'class']);
    const prevNumberIdx = findColIndex(headers, ['번호', '학번', '출석번호', 'no', 'num']);
    const nameIdx = findColIndex(headers, ['이름', '성명', '학생명', 'name']);
    const genderIdx = findColIndex(headers, ['성별', '성', '남녀', 'gender', 'sex']);
    const korIdx = findColIndex(headers, ['국어', 'korean', 'kor']);
    const mathIdx = findColIndex(headers, ['수학', 'math']);
    const totalIdx = findColIndex(headers, ['합계', '총점', '총합', 'total']);
    const rankIdx = findColIndex(headers, ['석차', '등수', '전체석차', 'rank']);
    const targetClassIdx = findColIndex(headers, ['진급반', '새반', '새로운반', '배정반', 'targetclass']);
    const remarksIdx = findColIndex(headers, ['비고', '특이사항', '메모', 'remarks', 'note']);

    if (nameIdx === -1) continue;

    let countInSheet = 0;

    for (let r = headerRowIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || !Array.isArray(row) || row.length === 0) continue;

      const name = String(row[nameIdx] || '').trim();
      if (!name) continue;

      let prevClass = prevClassIdx !== -1 ? String(row[prevClassIdx] || '').trim() : '';
      if (!prevClass) prevClass = sheetName;

      const prevNumber = prevNumberIdx !== -1 ? (parseInt(String(row[prevNumberIdx]), 10) || (countInSheet + 1)) : (countInSheet + 1);
      const gender = genderIdx !== -1 ? String(row[genderIdx] || '남').trim() : '남';
      const kor = korIdx !== -1 ? (parseFloat(String(row[korIdx])) || 0) : 0;
      const math = mathIdx !== -1 ? (parseFloat(String(row[mathIdx])) || 0) : 0;
      const total = totalIdx !== -1 ? (parseFloat(String(row[totalIdx])) || (kor + math)) : (kor + math);
      const rank = rankIdx !== -1 ? (parseInt(String(row[rankIdx]), 10) || 0) : 0;
      const targetClass = targetClassIdx !== -1 ? String(row[targetClassIdx] || '').trim() : '';
      const remarks = remarksIdx !== -1 ? String(row[remarksIdx] || '').trim() : '';

      detectedClassesSet.add(prevClass);
      countInSheet++;

      records.push({
        id: `row_${sheetName}_${r}`,
        prevClass,
        prevNumber,
        name,
        gender,
        koreanScore: kor,
        mathScore: math,
        totalScore: total,
        rank,
        targetClass,
        remarks
      });
    }

    if (countInSheet > 0) {
      sheetDetails.push({ name: sheetName, count: countInSheet });
    }
  }

  if (records.length === 0) {
    throw new Error('업로드된 엑셀 파일의 시트에서 유효한 학생 데이터를 찾을 수 없습니다.');
  }

  const sortedClasses = Array.from(detectedClassesSet).sort((a, b) => {
    const numA = parseInt(a.replace(/[^0-9]/g, ''), 10) || 0;
    const numB = parseInt(b.replace(/[^0-9]/g, ''), 10) || 0;
    return numA !== numB ? numA - numB : a.localeCompare(b, 'ko');
  });

  return {
    records,
    headers: commonHeaders,
    detectedClasses: sortedClasses,
    sheetDetails
  };
}

/**
 * 배정 결과 엑셀 워크북 생성 및 다운로드 (특이사항 셀 배경색 적용!)
 */
export function exportAssignmentToExcel(result: AssignmentResult, filename?: string): void {
  const xlsxLib = typeof window !== 'undefined' 
    ? ((window as any).XLSXStyle || (window as any).XLSX) 
    : (globalThis as any).XLSX;

  if (!xlsxLib) {
    alert('엑셀 라이브러리를 불러올 수 없습니다.');
    return;
  }

  const wb = xlsxLib.utils.book_new();

  // 1. 전체 통합 시트
  const allRows: any[] = [RESULT_STANDARD_HEADERS];
  const classes = Object.values(result.classes);

  for (const cls of classes) {
    const sortedStudents = [...cls.students].sort((a, b) => (a.assignedNumber || 0) - (b.assignedNumber || 0));

    for (const s of sortedStudents) {
      allRows.push([
        s.prevClass,
        s.prevNumber,
        s.name,
        s.gender,
        s.rank || '',
        cls.className,
        s.assignedNumber,
        s.rawRemarks
      ]);
    }
  }

  const wsAll = xlsxLib.utils.aoa_to_sheet(allRows);
  // 셀 배경색 및 스타일 적용!
  styleWorksheet(wsAll, 0);
  xlsxLib.utils.book_append_sheet(wb, wsAll, '전체배정결과');

  // 2. 각 진급반별 시트
  for (const cls of classes) {
    const classRows: any[] = [
      [`[${cls.className}] 학급 편성 명단`],
      [`총원: ${cls.totalStudents}명 | 가중 정원: ${cls.weightedCount}명 | 남: ${cls.maleCount}명 | 여: ${cls.femaleCount}명 | 특수: ${cls.specialCount}명 | 평균 석차: ${cls.avgRank}등`],
      [],
      RESULT_STANDARD_HEADERS
    ];

    const sortedStudents = [...cls.students].sort((a, b) => (a.assignedNumber || 0) - (b.assignedNumber || 0));

    sortedStudents.forEach(s => {
      classRows.push([
        s.prevClass,
        s.prevNumber,
        s.name,
        s.gender,
        s.rank || '',
        cls.className,
        s.assignedNumber,
        s.rawRemarks
      ]);
    });

    classRows.push([]);
    classRows.push(['-- 비고 항목 집계 --']);
    for (const [tag, count] of Object.entries(cls.tagCounts)) {
      classRows.push([tag, `${count}명`]);
    }

    const wsClass = xlsxLib.utils.aoa_to_sheet(classRows);
    // 반별 시트의 헤더는 3번 행(0-based 인덱스 3)
    styleWorksheet(wsClass, 3);
    const safeSheetName = cls.className.replace(/[:\\/?*\[\]]/g, '_').substring(0, 30);
    xlsxLib.utils.book_append_sheet(wb, wsClass, safeSheetName);
  }

  // 3. 학급별 통계 비교 시트
  const statsRows: any[] = [
    ['진급반', '실제 학생 수', '가중 학생 수(특수2명)', '남학생 수', '여학생 수', '남녀 성비(남:여)', '특수아동 수', '평균 석차', '평균 총점', '동일이름 충돌', '분리요청 위반']
  ];

  for (const cls of classes) {
    const ratio = cls.femaleCount > 0 ? (cls.maleCount / cls.femaleCount).toFixed(2) + ' : 1' : `${cls.maleCount} : 0`;
    statsRows.push([
      cls.className,
      cls.totalStudents,
      cls.weightedCount,
      cls.maleCount,
      cls.femaleCount,
      ratio,
      cls.specialCount,
      cls.avgRank,
      cls.avgScore,
      cls.nameCollisions.length === 0 ? '0건 (안전)' : `${cls.nameCollisions.length}건 충돌`,
      cls.separationViolations.length === 0 ? '0건 (안전)' : `${cls.separationViolations.length}건 위반`
    ]);
  }

  const wsStats = xlsxLib.utils.aoa_to_sheet(statsRows);
  styleWorksheet(wsStats, 0);
  xlsxLib.utils.book_append_sheet(wb, wsStats, '학급별통계비교');

  const now = new Date();
  const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`;
  xlsxLib.writeFile(wb, filename || `학급편성_배정결과_${dateStr}.xlsx`);
}

export function downloadSampleExcel(): void {
  const xlsxLib = typeof window !== 'undefined' 
    ? ((window as any).XLSXStyle || (window as any).XLSX) 
    : (globalThis as any).XLSX;

  if (!xlsxLib) return;

  const wb = xlsxLib.utils.book_new();

  const sampleDataSheet1: any[][] = [
    UPLOAD_STANDARD_HEADERS,
    ['1반', 1, '김민준', '남', 95, 98, 193, 1, '생활지도, 분리요청 (이민준)'],
    ['1반', 2, '박서아', '여', 92, 94, 186, 3, '다문화'],
    ['1반', 3, '남궁민', '남', 96, 92, 188, 2, ''],
    ['1반', 4, '조시우', '남', 78, 80, 158, 8, ''],
    ['1반', 5, '장예은', '여', 90, 88, 178, 5, '']
  ];

  const sampleDataSheet2: any[][] = [
    UPLOAD_STANDARD_HEADERS,
    ['2반', 1, '이민준', '남', 94, 95, 189, 2, '분리요청 (김민준)'],
    ['2반', 2, '최지우', '여', 88, 90, 178, 4, '특수아동'],
    ['2반', 3, '윤서준', '남', 85, 84, 169, 6, '생활지도'],
    ['2반', 4, '한지민', '여', 88, 87, 175, 5, '다문화'],
    ['2반', 5, '배준', '남', 75, 78, 153, 9, '학습부진']
  ];

  const ws1 = xlsxLib.utils.aoa_to_sheet(sampleDataSheet1);
  const ws2 = xlsxLib.utils.aoa_to_sheet(sampleDataSheet2);

  styleWorksheet(ws1, 0);
  styleWorksheet(ws2, 0);

  xlsxLib.utils.book_append_sheet(wb, ws1, '1반');
  xlsxLib.utils.book_append_sheet(wb, ws2, '2반');
  xlsxLib.writeFile(wb, '학급편성_표준업로드양식(9개항목_시트별기존반).xlsx');
}

export function generateSampleStudents(): RawStudentRecord[] {
  const firstNamesBoys = ['민준', '서준', '도윤', '예준', '시우', '하준', '주원', '지호', '지후', '준서', '준우', '도현', '건우', '선우', '민재', '현우'];
  const firstNamesGirls = ['서아', '하윤', '서윤', '지우', '서연', '하은', '지유', '수아', '아린', '지아', '예은', '채원', '유나', '소율', '수빈', '유진'];
  const surnames = ['김', '이', '박', '최', '정', '강', '조', '윤', '장', '임', '남궁', '제갈', '황보'];

  const records: RawStudentRecord[] = [];
  let seq = 1;

  for (let c = 1; c <= 4; c++) {
    const prevClass = `${c}반`;
    for (let num = 1; num <= 20; num++) {
      const isBoy = num % 2 === 1;
      const gender = isBoy ? '남' : '여';
      const surname = surnames[Math.floor(Math.random() * surnames.length)];
      const fList = isBoy ? firstNamesBoys : firstNamesGirls;
      const firstName = fList[Math.floor(Math.random() * 8)];
      let name = `${surname}${firstName}`;

      if (seq === 1) name = '김민준';
      if (seq === 21) name = '이민준';
      if (seq === 15) name = '조시우';
      if (seq === 35) name = '강도윤';

      const kor = Math.floor(Math.random() * 45) + 55;
      const math = Math.floor(Math.random() * 45) + 55;
      const total = kor + math;

      let remark = '';
      if (seq === 1) remark = '생활지도, 분리요청 (이민준)';
      else if (seq === 21) remark = '분리요청 (김민준)';
      else if (seq === 15) remark = '운동부';
      else if (seq === 35) remark = '학습부진, 분리요청 (조시우)';
      else if (seq === 7 || seq === 47) remark = '특수아동';
      else if (seq === 11 || seq === 51) remark = '다문화';
      else if (seq === 25 || seq === 65) remark = '생활지도';
      else if (seq === 19 || seq === 39) remark = '학습부진';

      records.push({
        id: `mock_${seq++}`,
        prevClass,
        prevNumber: num,
        name,
        gender,
        koreanScore: kor,
        mathScore: math,
        totalScore: total,
        rank: 0,
        targetClass: '',
        remarks: remark
      });
    }
  }

  const sorted = [...records].sort((a, b) => Number(b.totalScore) - Number(a.totalScore));
  sorted.forEach((item, idx) => {
    item.rank = idx + 1;
  });

  return records;
}
