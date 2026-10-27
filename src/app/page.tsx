"use client";

import React, { useState, useEffect, useRef } from "react";
import {
  preprocessStudents,
  optimizeAssignment,
  assignClassNumbers,
  buildSeparationPairs,
  RawStudentRecord
} from "@/lib/optimizer";
import {
  parseExcelFile,
  exportAssignmentToExcel,
  downloadSampleExcel,
  generateSampleStudents,
  UPLOAD_STANDARD_HEADERS,
  RESULT_STANDARD_HEADERS,
  SheetParseDetail
} from "@/lib/excelHandler";
import { parseRemarks, getTagColorStyle } from "@/lib/remarkParser";
import { AssignmentResult } from "@/lib/types";

const KOREAN_ALPHABET_CLASSES = ["가반", "나반", "다반", "라반", "마반", "바반", "사반", "아반", "자반", "차반", "카반", "타반", "파반", "하반"];

export default function Home() {
  const [currentStep, setCurrentStep] = useState<number>(1);
  const [rawRecords, setRawRecords] = useState<RawStudentRecord[]>([]);
  const [fileName, setFileName] = useState<string>("");
  const [sheetSummaryList, setSheetSummaryList] = useState<SheetParseDetail[]>([]);
  const [scoreType, setScoreType] = useState<"rank" | "score">("rank");
  const [numberingMode, setNumberingMode] = useState<"name_all" | "continuous" | "continuous_female" | "separate">("name_all");
  const [isDragging, setIsDragging] = useState<boolean>(false);

  const [prevClassCount, setPrevClassCount] = useState<number>(4);
  const [prevClassNames, setPrevClassNames] = useState<string[]>(["1반", "2반", "3반", "4반"]);

  const [newClassCount, setNewClassCount] = useState<number>(4);
  const [newClassNames, setNewClassNames] = useState<string[]>(["가반", "나반", "다반", "라반"]);

  const [isOptimizing, setIsOptimizing] = useState<boolean>(false);
  const [assignmentResult, setAssignmentResult] = useState<AssignmentResult | null>(null);
  const [activeResultTab, setActiveResultTab] = useState<string>("");
  const [searchTerm, setSearchTerm] = useState<string>("");
  const [genderFilter, setGenderFilter] = useState<string>("전체");

  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const preventDefault = (e: DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
    };
    window.addEventListener("dragover", preventDefault, false);
    window.addEventListener("drop", preventDefault, false);
    return () => {
      window.removeEventListener("dragover", preventDefault);
      window.removeEventListener("drop", preventDefault);
    };
  }, []);

  const handlePrevCountChange = (valStr: number | string) => {
    const val = Math.max(1, Math.min(20, Number(valStr) || 1));
    setPrevClassCount(val);
    setPrevClassNames((prev) => {
      const arr = [...prev];
      while (arr.length < val) arr.push(`${arr.length + 1}반`);
      return arr.slice(0, val);
    });
  };

  const handleNewCountChange = (valStr: number | string) => {
    const val = Math.max(1, Math.min(20, Number(valStr) || 1));
    setNewClassCount(val);
    setNewClassNames((prev) => {
      const arr = [...prev];
      while (arr.length < val) {
        const idx = arr.length;
        const def = idx < KOREAN_ALPHABET_CLASSES.length ? KOREAN_ALPHABET_CLASSES[idx] : `${idx + 1}반`;
        arr.push(def);
      }
      return arr.slice(0, val);
    });
  };

  const processFile = async (file: File) => {
    if (!file) return;
    setFileName(file.name);
    try {
      const { records, detectedClasses, sheetDetails } = await parseExcelFile(file);
      setRawRecords(records);
      setSheetSummaryList(sheetDetails);
      if (detectedClasses.length > 0) {
        setPrevClassCount(detectedClasses.length);
        setPrevClassNames(detectedClasses);
        setNewClassCount(detectedClasses.length);
        setNewClassNames(detectedClasses.map((_, i) => i < KOREAN_ALPHABET_CLASSES.length ? KOREAN_ALPHABET_CLASSES[i] : `${i + 1}반`));
      }
      setCurrentStep(2);
    } catch (err: any) {
      alert("파일 읽기 오류: " + err.message);
    }
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };

  const handleDragEnter = (e: React.DragEvent) => { e.preventDefault(); e.stopPropagation(); setIsDragging(true); };
  const handleDragOver = (e: React.DragEvent) => { e.preventDefault(); e.stopPropagation(); setIsDragging(true); };
  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault(); e.stopPropagation();
    if (e.currentTarget.contains(e.relatedTarget as Node)) return;
    setIsDragging(false);
  };
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault(); e.stopPropagation();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) processFile(file);
  };

  const handleLoadMockData = () => {
    const mock = generateSampleStudents();
    setRawRecords(mock);
    setFileName("2026학년도_학생명단_표준양식(분리요청_특이사항포함_80명).xlsx");
    setSheetSummaryList([
      { name: "1반", count: 20 },
      { name: "2반", count: 20 },
      { name: "3반", count: 20 },
      { name: "4반", count: 20 }
    ]);
    setPrevClassCount(4);
    setPrevClassNames(["1반", "2반", "3반", "4반"]);
    setNewClassCount(4);
    setNewClassNames(["가반", "나반", "다반", "라반"]);
    setCurrentStep(2);
  };

  const handleStartOptimization = () => {
    if (rawRecords.length === 0) {
      alert("학생 데이터를 먼저 업로드해주세요.");
      return;
    }
    if (newClassNames.length === 0) {
      alert("최소 1개 이상의 새로운 반이 필요합니다.");
      return;
    }

    setIsOptimizing(true);
    setTimeout(() => {
      try {
        const stds = preprocessStudents(rawRecords, scoreType);
        const res = optimizeAssignment(stds, newClassNames, prevClassNames, { numberingMode });
        setAssignmentResult(res);
        setActiveResultTab(newClassNames[0]);
        setCurrentStep(3);
      } catch (err: any) {
        alert("배정 연산 중 오류가 발생했습니다: " + err.message);
      } finally {
        setIsOptimizing(false);
      }
    }, 100);
  };

  const getRemarksCellBg = (tags: string[]) => {
    if (!tags || tags.length === 0) return "";
    for (const t of tags) {
      if (t.includes("생활지도")) return "bg-rose-50/70";
      if (t.includes("분리")) return "bg-purple-50/70";
      if (t.includes("특수")) return "bg-blue-50/70";
      if (t.includes("다문화")) return "bg-amber-50/70";
      if (t.includes("학습부진")) return "bg-orange-50/70";
    }
    return "bg-slate-50";
  };

  return (
    <div className="min-h-screen flex flex-col bg-slate-50">
      <header className="bg-white/95 backdrop-blur-md border-b border-emerald-100 sticky top-0 z-50 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-20 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="w-12 h-12 rounded-2xl bg-gradient-to-tr from-emerald-500 to-teal-400 flex items-center justify-center text-white shadow-md shadow-emerald-200">
              <svg className="w-7 h-7" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.2} d="M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253" />
              </svg>
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <h1 className="text-2xl font-bold tracking-tight text-slate-800">
                  <span className="text-emerald-600 font-extrabold">스마트 학급편성 시스템</span>
                </h1>
                <span className="px-2.5 py-0.5 text-xs font-semibold bg-emerald-100 text-emerald-800 rounded-full border border-emerald-200">
                  특이사항 컬러 & 분리요청
                </span>
              </div>
              <p className="text-xs text-slate-500 font-medium">
                특이사항별 고유 색상 강조 & 분리요청(이름) 무조건 다른 반 배정
              </p>
            </div>
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {currentStep === 1 && (
          <div className="space-y-8">
            <div className="rounded-3xl bg-gradient-to-r from-emerald-50 via-teal-50 to-green-50 border border-emerald-200 p-8 shadow-sm">
              <h2 className="text-2xl sm:text-3xl font-extrabold text-slate-800 tracking-tight mb-2">
                비고란 특이사항 <span className="text-emerald-600 underline">색상 배정</span> & <span className="text-purple-600 underline">분리요청</span> 자동 처리
              </h2>
              <p className="text-xs text-slate-600 mb-4 leading-relaxed">
                생활지도(빨강), 특수아동(파랑), 다문화(노랑), 학습부진(주황), 분리요청(보라) 등 특이사항별 색상이 자동 매핑되며 분리요청 대상자는 서로 다른 반으로 배정됩니다.
              </p>
              <div className="flex flex-wrap items-center justify-center gap-4 pt-2">
                <button onClick={handleLoadMockData} className="btn-3d btn-3d-white px-6 py-3 rounded-xl text-sm font-bold border border-slate-300 text-slate-700 hover:text-emerald-700 shadow-sm">
                  ⚡ 예시 데이터 불러오기 (80명)
                </button>
                <button onClick={downloadSampleExcel} className="btn-3d btn-3d-mint px-6 py-3 rounded-xl text-sm font-bold shadow-md shadow-emerald-200">
                  📥 표준 엑셀 양식 다운로드 (.xlsx)
                </button>
              </div>
            </div>

            <div className="card-soft p-8">
              <div
                onDragEnter={handleDragEnter}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
                className={`flex flex-col items-center justify-center w-full h-64 border-2 border-dashed rounded-3xl cursor-pointer transition-all duration-200 ${
                  isDragging
                    ? 'bg-emerald-100/90 border-emerald-600 scale-[1.01] shadow-2xl ring-4 ring-emerald-300'
                    : 'bg-emerald-50/40 border-emerald-300 hover:bg-emerald-50/80 hover:border-emerald-400'
                }`}
              >
                <div className="flex flex-col items-center justify-center p-6 text-center pointer-events-none">
                  <p className="text-base font-bold text-slate-700 mb-1">
                    {isDragging ? "파일을 여기에 놓아주세요!" : "파일을 여기로 드래그 앤 드롭하거나 클릭하여 선택"}
                  </p>
                  <p className="text-xs text-slate-500">모든 시트 자동 추출 · 분리요청(이름) 지원</p>
                  {fileName && <div className="mt-4 px-4 py-2 bg-emerald-100 rounded-2xl text-emerald-900 text-xs font-bold">✓ {fileName} ({rawRecords.length}명)</div>}
                </div>
                <input ref={fileInputRef} type="file" className="hidden" accept=".xlsx, .xls, .csv" onChange={handleFileInputChange} />
              </div>

              {/* 배정 성적 기준 & 진급반 번호 부여 방식 설정 메뉴 */}
              <div className="mt-8 grid grid-cols-1 lg:grid-cols-2 gap-6">
                {/* 배정 성적 기준 */}
                <div className="p-6 rounded-2xl bg-white border border-slate-200 shadow-sm flex flex-col justify-between">
                  <div className="flex items-center justify-between mb-4 pb-3 border-b border-slate-100">
                    <div className="flex items-center space-x-2">
                      <span className="text-xl">🎯</span>
                      <h3 className="text-base sm:text-lg font-bold text-slate-800">
                        배정 성적 기준
                      </h3>
                    </div>
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1 h-full items-stretch">
                    <button
                      type="button"
                      onClick={() => setScoreType('rank')}
                      className={`py-4 px-4 rounded-xl text-base font-bold transition-all text-center flex flex-col items-center justify-center gap-1 ${
                        scoreType === 'rank'
                          ? 'btn-3d btn-3d-mint shadow-md ring-2 ring-emerald-300'
                          : 'btn-3d btn-3d-white border border-slate-300 text-slate-700 hover:text-emerald-700'
                      }`}
                    >
                      <span className="text-base font-extrabold">석차 기준</span>
                      <span className="text-xs font-medium opacity-85">(1등이 우수)</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setScoreType('score')}
                      className={`py-4 px-4 rounded-xl text-base font-bold transition-all text-center flex flex-col items-center justify-center gap-1 ${
                        scoreType === 'score'
                          ? 'btn-3d btn-3d-mint shadow-md ring-2 ring-emerald-300'
                          : 'btn-3d btn-3d-white border border-slate-300 text-slate-700 hover:text-emerald-700'
                      }`}
                    >
                      <span className="text-base font-extrabold">합계 점수 기준</span>
                      <span className="text-xs font-medium opacity-85">(총점 높은 순)</span>
                    </button>
                  </div>
                </div>

                {/* 진급반 번호 부여 방식 */}
                <div className="p-6 rounded-2xl bg-white border border-slate-200 shadow-sm flex flex-col justify-between">
                  <div className="flex items-center justify-between mb-4 pb-3 border-b border-slate-100">
                    <div className="flex items-center space-x-2">
                      <span className="text-xl">🏷️</span>
                      <h3 className="text-base sm:text-lg font-bold text-slate-800">
                        진급반 번호 부여 방식
                      </h3>
                    </div>
                    <span className="text-xs font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2.5 py-1 rounded-full">
                      새 학급 번호 규칙
                    </span>
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5 pt-1">
                    <button
                      type="button"
                      onClick={() => setNumberingMode('name_all')}
                      className={`py-3 px-3 rounded-xl text-xs sm:text-sm font-bold transition-all text-center flex items-center justify-center ${
                        numberingMode === 'name_all'
                          ? 'btn-3d btn-3d-mint shadow-md ring-2 ring-emerald-300'
                          : 'btn-3d btn-3d-white border border-slate-300 text-slate-700 hover:text-emerald-700'
                      }`}
                    >
                      🌿 남녀 무관 이름순 (1~N번)
                    </button>
                    <button
                      type="button"
                      onClick={() => setNumberingMode('separate')}
                      className={`py-3 px-3 rounded-xl text-xs sm:text-sm font-bold transition-all text-center flex items-center justify-center ${
                        numberingMode === 'separate'
                          ? 'btn-3d btn-3d-mint shadow-md ring-2 ring-emerald-300'
                          : 'btn-3d btn-3d-white border border-slate-300 text-slate-700 hover:text-emerald-700'
                      }`}
                    >
                      남 1번~ / 여 1번~ (각각)
                    </button>
                    <button
                      type="button"
                      onClick={() => setNumberingMode('continuous')}
                      className={`py-3 px-3 rounded-xl text-xs sm:text-sm font-bold transition-all text-center flex items-center justify-center ${
                        numberingMode === 'continuous'
                          ? 'btn-3d btn-3d-mint shadow-md ring-2 ring-emerald-300'
                          : 'btn-3d btn-3d-white border border-slate-300 text-slate-700 hover:text-emerald-700'
                      }`}
                    >
                      남 1번~ / 여 21번~
                    </button>
                    <button
                      type="button"
                      onClick={() => setNumberingMode('continuous_female')}
                      className={`py-3 px-3 rounded-xl text-xs sm:text-sm font-bold transition-all text-center flex items-center justify-center ${
                        numberingMode === 'continuous_female'
                          ? 'btn-3d btn-3d-mint shadow-md ring-2 ring-emerald-300'
                          : 'btn-3d btn-3d-white border border-slate-300 text-slate-700 hover:text-emerald-700'
                      }`}
                    >
                      여 1번~ / 남 21번~
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {currentStep === 2 && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <h2 className="text-2xl font-extrabold text-slate-800">반 편성 및 번호 배정 상세</h2>
              <button onClick={handleStartOptimization} disabled={isOptimizing} className="btn-3d btn-3d-mint px-6 py-2.5 rounded-xl text-sm font-extrabold">
                {isOptimizing ? "배정 연산 중..." : "스마트 학급 배정 실행하기"}
              </button>
            </div>
            <div className="p-6 rounded-2xl bg-white border border-slate-200 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-4">
              <div className="flex items-center space-x-2">
                <span className="text-xl">🏷️</span>
                <span className="text-base font-bold text-slate-800">진급반 번호 부여 방식</span>
              </div>
              <div className="flex flex-wrap gap-2">
                <button
                  onClick={() => setNumberingMode("name_all")}
                  className={`py-2 px-3 rounded-xl text-xs sm:text-sm font-bold transition-all ${
                    numberingMode === "name_all" ? "btn-3d btn-3d-mint ring-2 ring-emerald-300 shadow-md" : "btn-3d btn-3d-white border border-slate-300 text-slate-700"
                  }`}
                >
                  🌿 이름순 (1~N번)
                </button>
                <button
                  onClick={() => setNumberingMode("separate")}
                  className={`py-2 px-3 rounded-xl text-xs sm:text-sm font-bold transition-all ${
                    numberingMode === "separate" ? "btn-3d btn-3d-mint ring-2 ring-emerald-300 shadow-md" : "btn-3d btn-3d-white border border-slate-300 text-slate-700"
                  }`}
                >
                  남 1번~ / 여 1번~
                </button>
                <button
                  onClick={() => setNumberingMode("continuous")}
                  className={`py-2 px-3 rounded-xl text-xs sm:text-sm font-bold transition-all ${
                    numberingMode === "continuous" ? "btn-3d btn-3d-mint ring-2 ring-emerald-300 shadow-md" : "btn-3d btn-3d-white border border-slate-300 text-slate-700"
                  }`}
                >
                  남 1번~ / 여 21번~
                </button>
                <button
                  onClick={() => setNumberingMode("continuous_female")}
                  className={`py-2 px-3 rounded-xl text-xs sm:text-sm font-bold transition-all ${
                    numberingMode === "continuous_female" ? "btn-3d btn-3d-mint ring-2 ring-emerald-300 shadow-md" : "btn-3d btn-3d-white border border-slate-300 text-slate-700"
                  }`}
                >
                  여 1번~ / 남 21번~
                </button>
              </div>
            </div>
          </div>
        )}

        {currentStep === 3 && assignmentResult && (
          <div className="space-y-8">
            <div className="flex flex-wrap items-center justify-between gap-4">
              <h2 className="text-2xl font-extrabold text-slate-800">배정 완료 대시보드</h2>
              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={() => setCurrentStep(1)}
                  className="btn-3d btn-3d-white px-4 py-2 rounded-xl text-xs font-bold flex items-center border border-slate-300 text-slate-700 hover:text-emerald-700"
                >
                  <span className="mr-1">🏠</span> 첫 화면으로 돌아가기
                </button>
                <button onClick={() => exportAssignmentToExcel(assignmentResult)} className="btn-3d btn-3d-mint px-5 py-2 rounded-xl text-xs font-bold">
                  전체 통합 엑셀 다운로드 (.xlsx)
                </button>
              </div>
            </div>

            <div className="card-soft p-6">
              <div className="flex flex-wrap gap-2 border-b border-slate-200 pb-4 mb-4">
                {newClassNames.map((name) => (
                  <button
                    key={name}
                    onClick={() => setActiveResultTab(name)}
                    className={`btn-3d px-4 py-2 rounded-xl text-xs font-bold ${activeResultTab === name ? "btn-3d-mint text-white" : "btn-3d-sage text-slate-700"}`}
                  >
                    진급 {name} ({assignmentResult.classes[name]?.totalStudents || 0}명)
                  </button>
                ))}
              </div>

              {assignmentResult.classes[activeResultTab] && (
                <div className="overflow-x-auto border border-slate-200 rounded-2xl">
                  <table className="w-full text-left text-xs font-mono">
                    <thead className="bg-slate-100 text-slate-700 font-bold uppercase">
                      <tr>
                        <th className="py-2.5 px-3">반 (기존)</th>
                        <th className="py-2.5 px-3">번호 (기존)</th>
                        <th className="py-2.5 px-3">이름</th>
                        <th className="py-2.5 px-3">성별</th>
                        <th className="py-2.5 px-3 text-center">석차</th>
                        <th className="py-2.5 px-3 bg-emerald-100 text-emerald-900 font-extrabold">진급반</th>
                        <th className="py-2.5 px-3 bg-emerald-100 text-emerald-900 font-extrabold text-center">번호 (새번호)</th>
                        <th className="py-2.5 px-3">비고 (특이사항 컬러)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 bg-white">
                      {assignmentResult.classes[activeResultTab].students.map((s) => {
                        const cellBg = getRemarksCellBg(s.tags);
                        return (
                          <tr key={s.id} className="hover:bg-emerald-50/30">
                            <td className="py-2 px-3">{s.prevClass}</td>
                            <td className="py-2 px-3">{s.prevNumber}</td>
                            <td className="py-2 px-3 font-bold text-slate-800">{s.name}</td>
                            <td className="py-2 px-3">{s.gender}</td>
                            <td className="py-2 px-3 text-center font-bold">{s.rank ? `${s.rank}등` : "-"}</td>
                            <td className="py-2 px-3 font-extrabold text-emerald-800 bg-emerald-50/50">{activeResultTab}</td>
                            <td className="py-2 px-3 text-center font-extrabold text-emerald-700 bg-emerald-50/50 text-sm">{s.assignedNumber}</td>
                            <td className={`py-2 px-3 ${cellBg}`}>
                              {s.tags && s.tags.length > 0 ? (
                                <div className="flex flex-wrap gap-1">
                                  {s.tags.map((t, idx) => {
                                    const style = getTagColorStyle(t);
                                    return (
                                      <span key={idx} className={`px-2 py-0.5 rounded text-[10px] font-bold border ${style.bg} ${style.text} ${style.border}`}>
                                        {t}
                                      </span>
                                    );
                                  })}
                                </div>
                              ) : "-"}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
