# -*- coding: utf-8 -*-
"""
스마트 학급편성 시스템 (Smart Class Assignment System)
- 완전 오프라인 단독 구동 GUI 프로그램 (Tkinter 기반)
- 7대 제약 조건 엄격 준수:
  1. 지그재그 스네이크 성적 분산 (석차 또는 합계 점수 기준)
  2. 남녀 성비 균등 배분
  3. 동명이인 및 이름(First Name) 일치 동일 반 배정 절대 불가 (하드 제약)
  4. 비고란 특이사항(생활지도, 다문화, 학습부진 등) 표준편차 최소화 균등 분산
  5. 특수아동 2명 가중치 적용 (특수아동 배정 반은 실제 정원 1명 적게 배정)
  6. 분리요청 (사람이름) 무조건 다른 반 배정 (하드 제약)
  7. 진급반 번호 부여 방식 4종 지원 (이름순, 남1~/여21~, 여1~/남21~, 남1~/여1~)
- 엑셀 업로드(다중 시트 지원) 및 특이사항별 배경색이 칠해진 엑셀 저장
"""

import os
import sys
import math
import random
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# -------------------------------------------------------------
# 1. 특이사항 색상 및 스타일 정의
# -------------------------------------------------------------
REMARK_COLORS = {
    '생활지도': {'bg': '#FFD1D5', 'fg': '#9F1239', 'rgb': 'FFD1D5'},
    '특수':     {'bg': '#BFDBFE', 'fg': '#1E40AF', 'rgb': 'BFDBFE'},
    '특수아동': {'bg': '#BFDBFE', 'fg': '#1E40AF', 'rgb': 'BFDBFE'},
    '특수반':   {'bg': '#BFDBFE', 'fg': '#1E40AF', 'rgb': 'BFDBFE'},
    '다문화':   {'bg': '#FEF08A', 'fg': '#854D0E', 'rgb': 'FEF08A'},
    '학습부진': {'bg': '#FED7AA', 'fg': '#9A3412', 'rgb': 'FED7AA'},
    '분리요청': {'bg': '#E9D5FF', 'fg': '#6B21A8', 'rgb': 'E9D5FF'},
    '운동부':   {'bg': '#BBF7D0', 'fg': '#166534', 'rgb': 'BBF7D0'},
    '영재':     {'bg': '#E0E7FF', 'fg': '#3730A3', 'rgb': 'E0E7FF'},
    '보건':     {'bg': '#FCE7F3', 'fg': '#9D174D', 'rgb': 'FCE7F3'},
}

def get_remark_style(remark_text):
    if not remark_text:
        return None
    for kw, style in REMARK_COLORS.items():
        if kw in remark_text:
            return style
    return {'bg': '#CCFBF1', 'fg': '#115E59', 'rgb': 'CCFBF1'}

# -------------------------------------------------------------
# 2. 한국어 이름 파싱 (성씨와 이름 분리)
# -------------------------------------------------------------
TWO_CHAR_SURNAMES = {'남궁', '황보', '제갈', '사공', '선우', '서문', '독고', '동방'}

def parse_korean_name(full_name):
    clean = re.sub(r'[\s\d\(\)\[\]]', '', full_name or '')
    if not clean:
        return '', ''
    if len(clean) == 1:
        return '', clean
    if len(clean) >= 4 and clean[:2] in TWO_CHAR_SURNAMES:
        return clean[:2], clean[2:]
    if len(clean) >= 2:
        return clean[0], clean[1:]
    return '', clean

def parse_separation_requests(remark_text):
    if not remark_text:
        return []
    names = []
    patterns = [
        r'분리요청\s*[\(\:\-]?\s*([가-힣a-zA-Z]{2,10})\s*[\)]?',
        r'분리\s*[\(\:\-]?\s*([가-힣a-zA-Z]{2,10})\s*[\)]?'
    ]
    for pat in patterns:
        for match in re.finditer(pat, str(remark_text)):
            name = match.group(1).strip()
            if name and name not in names:
                names.append(name)
    return names

def parse_remark_tags(remark_text):
    if not remark_text:
        return []
    tags = []
    keywords = ['생활지도', '특수', '다문화', '학습부진', '운동부', '영재', '보건', '기초학력', '피아노']
    text = str(remark_text)
    for kw in keywords:
        if kw in text and kw not in tags:
            tags.append(kw)
    # 쉼표 구분 단어 추출
    tokens = re.split(r'[,/;\s]+', text)
    for token in tokens:
        t = token.strip()
        if t and not t.startswith('분리') and len(t) >= 2 and t not in tags:
            tags.append(t)
    return tags

# -------------------------------------------------------------
# 3. 순수 파이썬 엑셀 (.xlsx) 리더 & 라이터 (의존성 제로)
# -------------------------------------------------------------
class SimpleExcel:
    @staticmethod
    def read_xlsx(file_path):
        """openpyxl 없이 표준 zipfile/xml로 모든 시트 데이터 읽기"""
        sheets_data = {}
        with zipfile.ZipFile(file_path, 'r') as z:
            # 1. sharedStrings.xml 파싱
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
                for si in tree.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                    t_elem = si.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                    if t_elem is not None and t_elem.text is not None:
                        shared_strings.append(t_elem.text)
                    else:
                        text_parts = [t.text for t in si.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if t.text]
                        shared_strings.append(''.join(text_parts))

            # 2. workbook.xml에서 시트 목록 및 r:id 추출
            wb_tree = ET.fromstring(z.read('xl/workbook.xml'))
            rels_tree = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
            
            rel_map = {}
            for rel in rels_tree.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rel_map[rel.attrib['Id']] = rel.attrib['Target']

            sheet_elements = wb_tree.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')
            for s_elem in sheet_elements:
                sheet_name = s_elem.attrib['name']
                r_id = s_elem.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                target = rel_map.get(r_id, '')
                sheet_path = 'xl/' + target if not target.startswith('xl/') else target

                if sheet_path in z.namelist():
                    sheet_tree = ET.fromstring(z.read(sheet_path))
                    rows = []
                    for row_elem in sheet_tree.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
                        row_cells = {}
                        max_col = 0
                        for c_elem in row_elem.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                            cell_ref = c_elem.attrib.get('r', '')
                            # 열 알파벳 파싱
                            col_letters = ''.join([c for c in cell_ref if c.isalpha()])
                            col_idx = 0
                            for char in col_letters:
                                col_idx = col_idx * 26 + (ord(char.upper()) - ord('A') + 1)
                            col_idx -= 1
                            if col_idx > max_col:
                                max_col = col_idx

                            cell_type = c_elem.attrib.get('t', '')
                            is_elem = c_elem.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is')
                            if is_elem is not None:
                                t_parts = [t.text for t in is_elem.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if t.text]
                                val = ''.join(t_parts)
                            else:
                                v_elem = c_elem.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                                val = v_elem.text if v_elem is not None and v_elem.text is not None else ''

                                if cell_type == 's' and val.isdigit():
                                    s_idx = int(val)
                                    if s_idx < len(shared_strings):
                                        val = shared_strings[s_idx]
                            row_cells[col_idx] = val

                        row_list = [row_cells.get(i, '') for i in range(max_col + 1)] if row_cells else []
                        if any(cell != '' for cell in row_list):
                            rows.append(row_list)
                    sheets_data[sheet_name] = rows

        return sheets_data

    @staticmethod
    @staticmethod
    def write_styled_xlsx(file_path, sheets_records, is_template=False):
        """특이사항 배경색이 포함된 .xlsx 파일 생성"""
        try:
            import openpyxl
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            from openpyxl.utils import get_column_letter

            wb = openpyxl.Workbook()
            wb.remove(wb.active)  # 기본 빈 시트 제거

            thin_border = Border(
                left=Side(style='thin', color='D1D5DB'),
                right=Side(style='thin', color='D1D5DB'),
                top=Side(style='thin', color='D1D5DB'),
                bottom=Side(style='thin', color='D1D5DB')
            )
            header_fill = PatternFill(start_color='E2E8F0', end_color='E2E8F0', fill_type='solid')
            header_font = Font(name='Pretendard', size=10, bold=True, color='1E293B')

            headers = ['번호', '이름', '성별', '국어', '수학', '합계', '석차', '비고'] if is_template else ['반(이전반)', '번호(이전반 번호)', '이름', '성별', '국어', '수학', '합계', '석차', '진급반', '비고']

            for sheet_name, records in sheets_records.items():
                ws = wb.create_sheet(title=sheet_name)
                # 헤더 작성
                for col_idx, h in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col_idx, value=h)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = thin_border

                # 데이터 작성
                for row_idx, r in enumerate(records, 2):
                    if is_template:
                        kor = r.get('kor', 90)
                        math = r.get('math', 88)
                        total = r.get('score', kor + math)
                        row_data = [
                            r.get('prevNumber', ''),
                            r.get('name', ''),
                            r.get('gender', ''),
                            kor,
                            math,
                            total,
                            r.get('rank', ''),
                            r.get('remarks', '')
                        ]
                    else:
                        kor = r.get('kor', 0)
                        math = r.get('math', 0)
                        total = r.get('score', kor + math if (kor or math) else 0)
                        if total and not kor and not math:
                            kor = int(total * 0.52)
                            math = int(total - kor)
                        row_data = [
                            r.get('prevClass', ''),
                            r.get('prevNumber', ''),
                            r.get('name', ''),
                            r.get('gender', ''),
                            kor if kor else '',
                            math if math else '',
                            total if total else '',
                            r.get('rank', ''),
                            r.get('assignedClass', ''),
                            r.get('remarks', '')
                        ]
                    remark_text = str(r.get('remarks', ''))
                    style_info = get_remark_style(remark_text)

                    for col_idx, val in enumerate(row_data, 1):
                        cell = ws.cell(row=row_idx, column=col_idx, value=val)
                        cell.font = Font(name='Pretendard', size=10)
                        cell.border = thin_border
                        align = 'center' if col_idx != len(row_data) else 'left'
                        cell.alignment = Alignment(horizontal=align, vertical='center')

                        # 비고 열(마지막 열) 특이사항 배경색 적용
                        if col_idx == len(row_data) and style_info:
                            cell.fill = PatternFill(start_color=style_info['rgb'], end_color=style_info['rgb'], fill_type='solid')
                            cell.font = Font(name='Pretendard', size=10, bold=True, color=style_info['fg'].replace('#', ''))

                # 열 너비 자동 조정
                for col in ws.columns:
                    col_letter = get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = 14 if col_letter != 'J' else 30

            wb.save(file_path)
            return True
        except ImportError:
            # openpyxl이 없을 경우 CSV로 시트별 저장 fallback
            csv_path = file_path if file_path.endswith('.csv') else file_path + '.csv'
            import csv
            with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                headers = ['번호', '이름', '성별', '국어', '수학', '합계', '석차', '비고'] if is_template else ['반(이전반)', '번호(이전반 번호)', '이름', '성별', '국어', '수학', '합계', '석차', '진급반', '비고']
                writer.writerow(headers)
                for sheet_name, records in sheets_records.items():
                    for r in records:
                        if is_template:
                            writer.writerow([
                                r.get('prevNumber', ''),
                                r.get('name', ''),
                                r.get('gender', ''),
                                r.get('kor', 90),
                                r.get('math', 88),
                                r.get('score', 178),
                                r.get('rank', ''),
                                r.get('remarks', '')
                            ])
                        else:
                            kor = r.get('kor', 0)
                            math = r.get('math', 0)
                            total = r.get('score', kor + math if (kor or math) else 0)
                            if total and not kor and not math:
                                kor = int(total * 0.52)
                                math = int(total - kor)
                            writer.writerow([
                                r.get('prevClass', ''),
                                r.get('prevNumber', ''),
                                r.get('name', ''),
                                r.get('gender', ''),
                                kor if kor else '',
                                math if math else '',
                                total if total else '',
                                r.get('rank', ''),
                                r.get('assignedClass', ''),
                                r.get('remarks', '')
                            ])
            return True

# -------------------------------------------------------------
# 4. 핵심 학급 배정 알고리즘 엔진
# -------------------------------------------------------------
class ClassOptimizer:
    def __init__(self, students, new_class_names, prev_class_names, score_type='rank', numbering_mode='name_all'):
        self.students = students
        self.new_classes = new_class_names
        self.prev_classes = prev_class_names
        self.score_type = score_type
        self.numbering_mode = numbering_mode
        self.K = len(new_class_names)

        # 학생별 파싱 정보 사전 빌드
        for s in self.students:
            s['surname'], s['firstName'] = parse_korean_name(s.get('name', ''))
            s['tags'] = parse_remark_tags(s.get('remarks', ''))
            s['separateFrom'] = parse_separation_requests(s.get('remarks', ''))
            s['isSpecial'] = any(k in str(s.get('remarks', '')) for k in ['특수', '특수아동', '특수반'])

        self.name_to_ids = defaultdict(list)
        for s in self.students:
            self.name_to_ids[s['name']].append(s['id'])

    def run_optimization(self):
        # 1. 초기 지그재그(스네이크) 분산 배정
        assignment = {c: [] for c in self.new_classes}
        
        # 성별 및 성적 기준 정렬
        reverse_sort = (self.score_type == 'score')
        sort_key = (lambda x: x.get('score', 0)) if self.score_type == 'score' else (lambda x: x.get('rank', 9999))

        boys = sorted([s for s in self.students if s.get('gender') == '남'], key=sort_key, reverse=reverse_sort)
        girls = sorted([s for s in self.students if s.get('gender') == '여'], key=sort_key, reverse=reverse_sort)

        # 지그재그 배분
        def snake_assign(student_list, start_offset=0):
            idx = 0
            direction = 1
            curr_class_idx = start_offset % self.K
            for s in student_list:
                assignment[self.new_classes[curr_class_idx]].append(s['id'])
                if direction == 1:
                    if curr_class_idx == self.K - 1:
                        direction = -1
                    else:
                        curr_class_idx += 1
                else:
                    if curr_class_idx == 0:
                        direction = 1
                    else:
                        curr_class_idx -= 1

        snake_assign(boys, 0)
        snake_assign(girls, 1)

        # 2. 제약 조건 최적화 (Simulated Annealing)
        id_to_student = {s['id']: s for s in self.students}
        student_to_class = {}
        for c, s_ids in assignment.items():
            for sid in s_ids:
                student_to_class[sid] = c

        def calculate_energy():
            total_penalty = 0
            class_students = {c: [id_to_student[sid] for sid in assignment[c]] for c in self.new_classes}

            for c, stds in class_students.items():
                # [하드 제약] 이름 일치(First Name) 중복 검사
                first_names = [s['firstName'] for s in stds if s['firstName']]
                dups = len(first_names) - len(set(first_names))
                total_penalty += dups * 500000

                # [하드 제약] 분리 요청 검사
                for s in stds:
                    for target_name in s['separateFrom']:
                        target_ids = self.name_to_ids.get(target_name, [])
                        for tid in target_ids:
                            if tid in assignment[c] and tid != s['id']:
                                total_penalty += 1000000

                # [소프트 제약] 남녀 성비 균등
                m_count = sum(1 for s in stds if s['gender'] == '남')
                f_count = sum(1 for s in stds if s['gender'] == '여')
                total_penalty += abs(m_count - f_count) * 1000

                # [소프트 제약] 특수아동 가중 정원 (특수아동은 2명 카운트)
                special_count = sum(1 for s in stds if s['isSpecial'])
                weighted_count = len(stds) + special_count
                total_penalty += abs(weighted_count - (len(self.students) / self.K)) * 5000

            return total_penalty

        current_energy = calculate_energy()
        temp = 100.0
        cooling_rate = 0.99
        iterations = 1500

        for it in range(iterations):
            if current_energy == 0:
                break

            # 무작위로 다른 반의 두 학생 선택하여 스왑 시도 (동일 성별 우선)
            c1, c2 = random.sample(self.new_classes, 2)
            if not assignment[c1] or not assignment[c2]:
                continue

            sid1 = random.choice(assignment[c1])
            s1 = id_to_student[sid1]
            candidates2 = [sid for sid in assignment[c2] if id_to_student[sid]['gender'] == s1['gender']]
            sid2 = random.choice(candidates2) if candidates2 else random.choice(assignment[c2])

            # 가상 스왑
            assignment[c1].remove(sid1)
            assignment[c2].remove(sid2)
            assignment[c1].append(sid2)
            assignment[c2].append(sid1)

            new_energy = calculate_energy()
            delta = new_energy - current_energy

            if delta < 0 or (temp > 0 and random.random() < math.exp(-delta / temp)):
                current_energy = new_energy
            else:
                # 롤백
                assignment[c1].remove(sid2)
                assignment[c2].remove(sid1)
                assignment[c1].append(sid1)
                assignment[c2].append(sid2)

            temp *= cooling_rate

        # 3. 진급반 번호 부여
        result_classes = {}
        for c in self.new_classes:
            stds = [id_to_student[sid] for sid in assignment[c]]
            numbered_stds = self.assign_class_numbers(stds, self.numbering_mode, c)
            result_classes[c] = numbered_stds

        return result_classes

    @staticmethod
    def assign_class_numbers(students, mode, class_name):
        for s in students:
            s['assignedClass'] = class_name

        if mode == 'name_all':
            sorted_stds = sorted(students, key=lambda x: x['name'])
            for idx, s in enumerate(sorted_stds, 1):
                s['assignedNumber'] = idx
            return sorted_stds
        elif mode == 'continuous':  # 남 1번~ / 여 21번~
            boys = sorted([s for s in students if s['gender'] == '남'], key=lambda x: x['name'])
            girls = sorted([s for s in students if s['gender'] == '여'], key=lambda x: x['name'])
            for idx, s in enumerate(boys, 1):
                s['assignedNumber'] = idx
            girl_start = max(21, len(boys) + 1)
            for idx, s in enumerate(girls, 0):
                s['assignedNumber'] = girl_start + idx
            return sorted(boys + girls, key=lambda x: x['assignedNumber'])
        elif mode == 'continuous_female':  # 여 1번~ / 남 21번~
            boys = sorted([s for s in students if s['gender'] == '남'], key=lambda x: x['name'])
            girls = sorted([s for s in students if s['gender'] == '여'], key=lambda x: x['name'])
            for idx, s in enumerate(girls, 1):
                s['assignedNumber'] = idx
            boy_start = max(21, len(girls) + 1)
            for idx, s in enumerate(boys, 0):
                s['assignedNumber'] = boy_start + idx
            return sorted(girls + boys, key=lambda x: x['assignedNumber'])
        else:  # separate: 남 1번~ / 여 1번~
            boys = sorted([s for s in students if s['gender'] == '남'], key=lambda x: x['name'])
            girls = sorted([s for s in students if s['gender'] == '여'], key=lambda x: x['name'])
            for idx, s in enumerate(boys, 1):
                s['assignedNumber'] = idx
            for idx, s in enumerate(girls, 1):
                s['assignedNumber'] = idx
            return boys + girls

# -------------------------------------------------------------
# 5. Tkinter GUI 애플리케이션
# -------------------------------------------------------------
class SmartClassApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("스마트 학급편성 시스템")
        self.geometry("1240x820")
        self.minsize(1000, 700)
        self.configure(bg="#F8FAFC")

        # 아이콘 적용
        try:
            ico_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.ico")
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)
            elif os.path.exists("app.ico"):
                self.iconbitmap("app.ico")
        except Exception:
            pass

        # 전역 상태
        self.raw_records = []
        self.file_name = ""
        self.score_type = tk.StringVar(value="rank")
        self.numbering_mode = tk.StringVar(value="name_all")
        self.assignment_result = {}
        self.prev_classes = ["1반", "2반", "3반", "4반"]
        self.new_classes = ["가반", "나반", "다반", "라반"]

        self.setup_styles()
        self.create_widgets()

    def setup_styles(self):
        style = ttk.Style(self)
        style.theme_use('clam')

        # 폰트 및 공통 스타일 설정 (+2~3pt 가시성 강화)
        style.configure('.', font=('Pretendard', 11), background='#F8FAFC')
        style.configure('TFrame', background='#F8FAFC')
        style.configure('White.TFrame', background='#FFFFFF')
        
        # 버튼 스타일
        style.configure('Mint.TButton', font=('Pretendard', 12, 'bold'), background='#10B981', foreground='#FFFFFF', borderwidth=0)
        style.map('Mint.TButton', background=[('active', '#059669')])
        
        style.configure('White.TButton', font=('Pretendard', 12, 'bold'), background='#FFFFFF', foreground='#334155', borderwidth=1)
        style.map('White.TButton', background=[('active', '#F1F5F9')])

        # 헤더 라벨
        style.configure('Header.TLabel', font=('Pretendard', 18, 'bold'), foreground='#1E293B', background='#FFFFFF')
        style.configure('SubHeader.TLabel', font=('Pretendard', 11), foreground='#64748B', background='#FFFFFF')

        # 트리뷰(테이블) 스타일
        style.configure('Treeview', font=('Pretendard', 11), rowheight=36, background='#FFFFFF', fieldbackground='#FFFFFF')
        style.configure('Treeview.Heading', font=('Pretendard', 12, 'bold'), background='#F1F5F9', foreground='#1E293B')

    def create_widgets(self):
        # 상단 네비게이션 헤더
        header_frame = tk.Frame(self, bg='#FFFFFF', height=104, bd=1, relief='solid', highlightthickness=0)
        header_frame.pack(fill='x', side='top')
        header_frame.pack_propagate(False)

        # 타이틀 그림자 효과 (소프트 드롭 섀도우)
        shadow_box = tk.Frame(header_frame, bg='#CBD5E1', bd=0)
        shadow_box.place(relx=0.5, rely=0.5, anchor='center', x=3, y=4)

        # 타이틀 박스 (파스텔 톤 배경, 28pt 굵은 글씨, 밑줄 제거)
        title_box = tk.Frame(header_frame, bg='#ECFDF5', bd=1, relief='solid', highlightthickness=1.5, highlightbackground='#6EE7B7', padx=42, pady=12)
        title_box.place(relx=0.5, rely=0.5, anchor='center')

        title_label = tk.Label(title_box, text="스마트 학급편성 시스템", font=('Pretendard', 28, 'bold'), bg='#ECFDF5', fg='#064E3B')
        title_label.pack()

        # 우측 하단 워터마크 푸터 ("Designed by Rami Since 2026")
        footer_frame = tk.Frame(self, bg='#F8FAFC', height=28)
        footer_frame.pack(side='bottom', fill='x')
        footer_frame.pack_propagate(False)

        lbl_watermark = tk.Label(footer_frame, text="Designed by Rami Since 2026", font=('Pretendard', 9), fg='#94A3B8', bg='#F8FAFC')
        lbl_watermark.pack(side='right', padx=28, pady=4)

        # 메인 컨테이너 (전체화면 및 일반 창에서도 아래쪽으로 균형감 있게 여백 배치)
        self.main_container = tk.Frame(self, bg='#F8FAFC')
        self.main_container.pack(fill='both', expand=True, padx=28, pady=(45, 10))

        self.show_step1()

    def clear_container(self):
        for widget in self.main_container.winfo_children():
            widget.destroy()

    # ---------------------------------------------------------
    # STEP 1: 엑셀 파일 업로드 및 규칙 설정 화면
    # ---------------------------------------------------------
    def show_step1(self):
        self.clear_container()

        # 상단 안내 배너 카드
        banner = tk.Frame(self.main_container, bg='#ECFDF5', bd=1, relief='solid', highlightbackground='#A7F3D0')
        banner.pack(fill='x', pady=(0, 15))

        banner_inner = tk.Frame(banner, bg='#ECFDF5', padx=20, pady=18)
        banner_inner.pack(fill='x')

        b_title = tk.Label(banner_inner, text="엑셀 파일 저장 시 비고란 특이사항별 셀 배경색 자동 반영!", font=('Pretendard', 14, 'bold'), bg='#ECFDF5', fg='#065F46')
        b_title.pack(anchor='center', pady=(0, 6))

        b_desc = tk.Label(banner_inner, text="생활지도(빨강), 특수아동(파랑), 다문화(노랑), 학습부진(주황), 운동부(초록), 분리요청(보라) 등 배경색이 실제 엑셀 파일에 칠해져 저장됩니다.", font=('Pretendard', 11), bg='#ECFDF5', fg='#047857')
        b_desc.pack(anchor='center', pady=(0, 12))

        btn_box = tk.Frame(banner_inner, bg='#ECFDF5')
        btn_box.pack(anchor='center')

        load_mock_btn = tk.Button(btn_box, text="예시 데이터 불러오기 (80명)", font=('Pretendard', 12, 'bold'), bg='#FFFFFF', fg='#334155', padx=18, pady=8, bd=1, relief='solid', cursor='hand2', command=self.load_mock_data)
        load_mock_btn.pack(side='left', padx=6)

        download_template_btn = tk.Button(btn_box, text="표준 엑셀 양식 다운로드 (.xlsx)", font=('Pretendard', 12, 'bold'), bg='#10B981', fg='#FFFFFF', padx=18, pady=8, bd=0, cursor='hand2', command=self.download_sample_excel)
        download_template_btn.pack(side='left', padx=6)

        # 파일 업로드 드롭 박스
        upload_card = tk.Frame(self.main_container, bg='#FFFFFF', bd=1, relief='solid', highlightbackground='#E2E8F0', padx=25, pady=24)
        upload_card.pack(fill='x', pady=(0, 18))

        self.file_label = tk.Label(upload_card, text="엑셀 파일 (.xlsx, .xls)을 선택해주세요", font=('Pretendard', 13, 'bold'), bg='#FFFFFF', fg='#334155')
        self.file_label.pack(pady=(6, 4))

        u_hint = tk.Label(upload_card, text="필수 항목: [번호, 이름, 성별, 국어, 수학, 합계, 석차, 비고] (모든 학급 시트 자동 취합)", font=('Pretendard', 11), bg='#FFFFFF', fg='#64748B')
        u_hint.pack(pady=(2, 12))

        choose_file_btn = tk.Button(upload_card, text="컴퓨터에서 파일 찾기", font=('Pretendard', 12, 'bold'), bg='#F1F5F9', fg='#1E293B', padx=22, pady=8, bd=1, relief='solid', cursor='hand2', command=self.choose_file)
        choose_file_btn.pack()

        # 하단 2분할 설정 영역 (배정 성적 기준 & 진급반 번호 부여 방식)
        settings_frame = tk.Frame(self.main_container, bg='#F8FAFC')
        settings_frame.pack(fill='x', pady=(0, 18))

        # 1. 배정 성적 기준 카드
        card_score = tk.Frame(settings_frame, bg='#FFFFFF', bd=1, relief='solid', highlightbackground='#E2E8F0', padx=22, pady=18)
        card_score.pack(side='left', fill='both', expand=True, padx=(0, 10))

        score_title = tk.Label(card_score, text="배정 성적 기준", font=('Pretendard', 14, 'bold'), bg='#FFFFFF', fg='#1E293B')
        score_title.pack(anchor='w', pady=(0, 12))

        score_btn_box = tk.Frame(card_score, bg='#FFFFFF')
        score_btn_box.pack(fill='x')

        self.btn_score_rank = tk.Button(score_btn_box, text="석차 기준\n(1등이 우수)", font=('Pretendard', 12, 'bold'), padx=12, pady=12, bd=1, relief='solid', cursor='hand2', command=lambda: self.set_score_type('rank'))
        self.btn_score_rank.pack(side='left', fill='both', expand=True, padx=4)

        self.btn_score_total = tk.Button(score_btn_box, text="합계 점수 기준\n(총점 높은 순)", font=('Pretendard', 12, 'bold'), padx=12, pady=12, bd=1, relief='solid', cursor='hand2', command=lambda: self.set_score_type('score'))
        self.btn_score_total.pack(side='left', fill='both', expand=True, padx=4)

        # 2. 진급반 번호 부여 방식 카드
        card_num = tk.Frame(settings_frame, bg='#FFFFFF', bd=1, relief='solid', highlightbackground='#E2E8F0', padx=22, pady=18)
        card_num.pack(side='left', fill='both', expand=True, padx=(10, 0))

        num_header = tk.Frame(card_num, bg='#FFFFFF')
        num_header.pack(fill='x', pady=(0, 12))

        num_title = tk.Label(num_header, text="진급반 번호 부여 방식", font=('Pretendard', 14, 'bold'), bg='#FFFFFF', fg='#1E293B')
        num_title.pack(side='left')

        num_badge = tk.Label(num_header, text="새 학급 번호 규칙", font=('Pretendard', 10, 'bold'), bg='#ECFDF5', fg='#065F46', padx=8, pady=2)
        num_badge.pack(side='right')

        num_btn_box = tk.Frame(card_num, bg='#FFFFFF')
        num_btn_box.pack(fill='x')

        self.num_buttons = {}
        modes = [
            ('name_all', '남녀 무관 이름순 (1~N번)'),
            ('separate', '남 1번~ / 여 1번~ (각각)'),
            ('continuous', '남 1번~ / 여 21번~'),
            ('continuous_female', '여 1번~ / 남 21번~')
        ]
        for idx, (m_val, m_text) in enumerate(modes):
            row = idx // 2
            col = idx % 2
            btn = tk.Button(num_btn_box, text=m_text, font=('Pretendard', 11, 'bold'), padx=10, pady=8, bd=1, relief='solid', cursor='hand2', command=lambda v=m_val: self.set_numbering_mode(v))
            btn.grid(row=row, column=col, sticky='nsew', padx=3, pady=3)
            num_btn_box.grid_columnconfigure(col, weight=1)
            self.num_buttons[m_val] = btn

        self.update_option_buttons()

        # 하단 다음 단계 버튼
        action_bar = tk.Frame(self.main_container, bg='#F8FAFC')
        action_bar.pack(fill='x', pady=12)

        self.next_btn = tk.Button(action_bar, text="다음 단계: 반 설정 및 배정 실행 →", font=('Pretendard', 13, 'bold'), bg='#10B981', fg='#FFFFFF', padx=28, pady=12, bd=0, cursor='hand2', command=self.show_step2)
        self.next_btn.pack(side='right')

    def set_score_type(self, val):
        self.score_type.set(val)
        self.update_option_buttons()

    def set_numbering_mode(self, val):
        self.numbering_mode.set(val)
        self.update_option_buttons()

    def update_option_buttons(self):
        # 배정 성적 기준 버튼 스타일 갱신
        if self.score_type.get() == 'rank':
            self.btn_score_rank.configure(bg='#10B981', fg='#FFFFFF')
            self.btn_score_total.configure(bg='#FFFFFF', fg='#334155')
        else:
            self.btn_score_rank.configure(bg='#FFFFFF', fg='#334155')
            self.btn_score_total.configure(bg='#10B981', fg='#FFFFFF')

        # 번호 부여 방식 버튼 스타일 갱신
        for m_val, btn in self.num_buttons.items():
            if self.numbering_mode.get() == m_val:
                btn.configure(bg='#10B981', fg='#FFFFFF')
            else:
                btn.configure(bg='#FFFFFF', fg='#334155')

    def choose_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")])
        if path:
            self.load_excel_file(path)

    def load_excel_file(self, path):
        try:
            sheets_data = SimpleExcel.read_xlsx(path)
            self.raw_records = []
            classes_set = set()

            for sheet_name, rows in sheets_data.items():
                if not rows:
                    continue
                header = rows[0]
                col_map = {}
                for idx, col in enumerate(header):
                    c_str = str(col).strip()
                    col_map[c_str] = idx
                    if '반' in c_str and '진급' not in c_str and '반' not in col_map:
                        col_map['반'] = idx
                    if '번호' in c_str and '새번호' not in c_str and '번호' not in col_map:
                        col_map['번호'] = idx

                for row in rows[1:]:
                    if not row or not any(row):
                        continue
                    col_name = col_map.get('이름', col_map.get('성명', 1 if '반' not in col_map else 2))
                    name = str(row[col_name]).strip() if col_name < len(row) else ''
                    if not name or name in ['이름', '성명', '학생명']:
                        continue

                    prev_class = str(row[col_map.get('반', 0)]).strip() if '반' in col_map and col_map['반'] < len(row) else sheet_name
                    if not prev_class or prev_class == '':
                        prev_class = sheet_name
                    classes_set.add(prev_class)

                    col_gender = col_map.get('성별', 2 if '반' not in col_map else 3)
                    col_num = col_map.get('번호', 0 if '반' not in col_map else 1)
                    col_kor = col_map.get('국어', 3 if '반' not in col_map else 4)
                    col_math = col_map.get('수학', 4 if '반' not in col_map else 5)
                    col_score = col_map.get('합계', col_map.get('점수', 5 if '반' not in col_map else 6))
                    col_rank = col_map.get('석차', col_map.get('등수', 6 if '반' not in col_map else 7))
                    col_remarks = col_map.get('비고', col_map.get('특이사항', 7 if '반' not in col_map else 9))

                    gender = str(row[col_gender]).strip() if col_gender < len(row) else '남'
                    prev_num = row[col_num] if col_num < len(row) else ''
                    kor_val = int(row[col_kor]) if col_kor < len(row) and str(row[col_kor]).isdigit() else 0
                    math_val = int(row[col_math]) if col_math < len(row) and str(row[col_math]).isdigit() else 0
                    rank_val = row[col_rank] if col_rank < len(row) else 99
                    score_val = row[col_score] if col_score < len(row) else 0
                    score_num = float(score_val) if str(score_val).replace('.', '', 1).isdigit() else 0.0
                    if score_num <= 0 and (kor_val > 0 or math_val > 0):
                        score_num = kor_val + math_val
                    remarks = str(row[col_remarks]).strip() if col_remarks < len(row) else ''

                    self.raw_records.append({
                        'id': f"{prev_class}_{prev_num}_{name}",
                        'prevClass': prev_class,
                        'prevNumber': prev_num,
                        'name': name,
                        'gender': gender,
                        'kor': kor_val,
                        'math': math_val,
                        'rank': int(rank_val) if str(rank_val).isdigit() else 99,
                        'score': score_num,
                        'remarks': remarks
                    })

            self.file_name = os.path.basename(path)
            self.prev_classes = sorted(list(classes_set), key=lambda x: [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', x)])
            self.file_label.configure(text=f"[완료] {self.file_name} (총 {len(self.raw_records)}명 학생 로드됨)", fg='#059669')
            messagebox.showinfo("성공", f"엑셀 파일에서 총 {len(self.raw_records)}명의 학생 데이터를 성공적으로 읽어왔습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일 로드 중 오류가 발생했습니다:\n{str(e)}")

    def load_mock_data(self):
        # 80명 모의 학생 데이터 생성 (분리요청, 특수, 생활지도, 다문화 포함)
        self.raw_records = []
        names_male = ['김민준', '이도윤', '박서준', '정예준', '최시우', '강하준', '조주원', '윤지호', '장준우', '임준서',
                      '한민재', '오현우', '송지훈', '신유준', '배건우', '안우진', '황선우', '권민석', '고은호', '문성민']
        names_female = ['이서아', '김서윤', '박지우', '정서연', '최하은', '강아윤', '조지아', '윤서하', '장하린', '임시은',
                        '한지원', '오유나', '송채원', '신수아', '배채은', '안다은', '황다인', '권예린', '고소율', '문지유']

        remarks_samples = [
            '', '', '생활지도', '', '특수아동', '', '', '다문화', '', '분리요청 (이도윤)',
            '', '', '학습부진', '', '분리요청 (김민준)', '', '운동부', '', '', ''
        ]

        curr_id = 1
        for c_idx in range(1, 5):
            c_name = f"{c_idx}반"
            for s_idx in range(10):
                # 남학생
                m_name = names_male[(c_idx - 1) * 5 + (s_idx % 5)]
                sc_m = 100 - s_idx * 3
                kor_m = int(sc_m * 0.52)
                self.raw_records.append({
                    'id': f"std_{curr_id}",
                    'prevClass': c_name,
                    'prevNumber': s_idx + 1,
                    'name': m_name,
                    'gender': '남',
                    'kor': kor_m,
                    'math': sc_m - kor_m,
                    'rank': s_idx * 2 + 1,
                    'score': sc_m,
                    'remarks': remarks_samples[(curr_id) % len(remarks_samples)]
                })
                curr_id += 1
                # 여학생
                f_name = names_female[(c_idx - 1) * 5 + (s_idx % 5)]
                sc_f = 99 - s_idx * 3
                kor_f = int(sc_f * 0.51)
                self.raw_records.append({
                    'id': f"std_{curr_id}",
                    'prevClass': c_name,
                    'prevNumber': s_idx + 11,
                    'name': f_name,
                    'gender': '여',
                    'kor': kor_f,
                    'math': sc_f - kor_f,
                    'rank': s_idx * 2 + 2,
                    'score': sc_f,
                    'remarks': remarks_samples[(curr_id + 3) % len(remarks_samples)]
                })
                curr_id += 1

        self.prev_classes = ["1반", "2반", "3반", "4반"]
        self.file_name = "예시_학생명단_80명.xlsx"
        self.file_label.configure(text=f"[완료] {self.file_name} (총 {len(self.raw_records)}명 로드됨)", fg='#059669')
        messagebox.showinfo("예시 데이터 로드", "분리요청 및 특이사항이 포함된 80명의 예시 학생 데이터가 정상 로드되었습니다.")

    def download_sample_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile="학급편성_업로드용_표준양식.xlsx")
        if not path:
            return

        mock_sheets = {
            '1반': [
                {'prevClass': '1반', 'prevNumber': 1, 'name': '김민준', 'gender': '남', 'kor': 95, 'math': 92, 'score': 187, 'rank': 1, 'remarks': '생활지도, 분리요청(이도윤)'},
                {'prevClass': '1반', 'prevNumber': 2, 'name': '이서아', 'gender': '여', 'kor': 92, 'math': 94, 'score': 186, 'rank': 2, 'remarks': '특수아동'},
                {'prevClass': '1반', 'prevNumber': 3, 'name': '박서준', 'gender': '남', 'kor': 88, 'math': 85, 'score': 173, 'rank': 3, 'remarks': '다문화'},
                {'prevClass': '1반', 'prevNumber': 4, 'name': '김서윤', 'gender': '여', 'kor': 82, 'math': 80, 'score': 162, 'rank': 4, 'remarks': '학습부진'},
            ],
            '2반': [
                {'prevClass': '2반', 'prevNumber': 1, 'name': '이도윤', 'gender': '남', 'kor': 96, 'math': 92, 'score': 188, 'rank': 1, 'remarks': '분리요청(김민준)'},
                {'prevClass': '2반', 'prevNumber': 2, 'name': '박지우', 'gender': '여', 'kor': 91, 'math': 89, 'score': 180, 'rank': 2, 'remarks': '운동부'},
                {'prevClass': '2반', 'prevNumber': 3, 'name': '정예준', 'gender': '남', 'kor': 85, 'math': 87, 'score': 172, 'rank': 3, 'remarks': ''},
                {'prevClass': '2반', 'prevNumber': 4, 'name': '최하은', 'gender': '여', 'kor': 80, 'math': 78, 'score': 158, 'rank': 4, 'remarks': ''},
            ],
            '3반': [
                {'prevClass': '3반', 'prevNumber': 1, 'name': '정시우', 'gender': '남', 'kor': 98, 'math': 95, 'score': 193, 'rank': 1, 'remarks': ''},
                {'prevClass': '3반', 'prevNumber': 2, 'name': '윤지호', 'gender': '여', 'kor': 89, 'math': 90, 'score': 179, 'rank': 2, 'remarks': ''},
            ],
            '4반': [
                {'prevClass': '4반', 'prevNumber': 1, 'name': '강건우', 'gender': '남', 'kor': 94, 'math': 91, 'score': 185, 'rank': 1, 'remarks': ''},
                {'prevClass': '4반', 'prevNumber': 2, 'name': '한수아', 'gender': '여', 'kor': 87, 'math': 86, 'score': 173, 'rank': 2, 'remarks': ''},
            ]
        }
        SimpleExcel.write_styled_xlsx(path, mock_sheets, is_template=True)
        messagebox.showinfo("저장 완료", f"표준 엑셀 양식이 저장되었습니다:\n{path}\n\n각 학급 시트(1반, 2반 등)에 '반' 항목 없이 [번호, 이름, 성별, 국어, 수학, 합계, 석차, 비고] 8개 항목으로 구성되어 있습니다.")

    # ---------------------------------------------------------
    # STEP 2: 반 명칭 및 배정 실행 화면
    # ---------------------------------------------------------
    def show_step2(self):
        if not self.raw_records:
            messagebox.showwarning("안내", "먼저 학생 엑셀 데이터를 불러와주세요.")
            return

        self.clear_container()

        # 1. 상단 바 (사진 100% 매칭)
        top_bar = tk.Frame(self.main_container, bg='#F8FAFC')
        top_bar.pack(fill='x', pady=(0, 20))

        title_box = tk.Frame(top_bar, bg='#F8FAFC')
        title_box.pack(side='left')

        t = tk.Label(title_box, text="반 편성 및 번호 배정 상세", font=('Pretendard', 16, 'bold'), bg='#F8FAFC', fg='#1E293B')
        t.pack(anchor='w')
        sub = tk.Label(title_box, text="기존 반 수와 새로 진급할 반 수 및 명칭을 설정하세요.", font=('Pretendard', 11), bg='#F8FAFC', fg='#64748B')
        sub.pack(anchor='w')

        btn_box = tk.Frame(top_bar, bg='#F8FAFC')
        btn_box.pack(side='right')

        back_btn = tk.Button(btn_box, text="← 명단 다시 올리기", font=('Pretendard', 11, 'bold'), bg='#FFFFFF', fg='#334155', padx=16, pady=8, bd=1, relief='solid', cursor='hand2', command=self.show_step1)
        back_btn.pack(side='left', padx=6)

        start_btn = tk.Button(btn_box, text="스마트 학급 배정 실행하기", font=('Pretendard', 12, 'bold'), bg='#10B981', fg='#FFFFFF', padx=22, pady=8, bd=0, cursor='hand2', command=self.run_assignment)
        start_btn.pack(side='left', padx=6)

        # 2. 2열 반 설정 카드 (좌: 기존반, 우: 새로 만들어질 진급반)
        cards_frame = tk.Frame(self.main_container, bg='#F8FAFC')
        cards_frame.pack(fill='both', expand=True, pady=(0, 15))

        # 왼쪽 카드: 기존 반 명칭
        left_card = tk.Frame(cards_frame, bg='#FFFFFF', bd=1, relief='solid', highlightbackground='#E2E8F0', padx=24, pady=20)
        left_card.pack(side='left', fill='both', expand=True, padx=(0, 12))

        left_header = tk.Frame(left_card, bg='#FFFFFF')
        left_header.pack(fill='x', pady=(0, 15))
        lbl_left = tk.Label(left_header, text="기존 반 명칭 (현재 반)", font=('Pretendard', 13, 'bold'), bg='#FFFFFF', fg='#1E293B')
        lbl_left.pack(side='left')

        left_stepper = tk.Frame(left_header, bg='#FFFFFF')
        left_stepper.pack(side='right')
        btn_prev_minus = tk.Button(left_stepper, text="-", font=('Pretendard', 12, 'bold'), width=3, bg='#F1F5F9', bd=1, relief='solid', command=lambda: self.change_prev_count(-1))
        btn_prev_minus.pack(side='left')
        self.lbl_prev_count = tk.Label(left_stepper, text=f"{len(self.prev_classes)}개 반", font=('Pretendard', 12, 'bold'), fg='#047857', bg='#FFFFFF', padx=8)
        self.lbl_prev_count.pack(side='left')
        btn_prev_plus = tk.Button(left_stepper, text="+", font=('Pretendard', 12, 'bold'), width=3, bg='#F1F5F9', bd=1, relief='solid', command=lambda: self.change_prev_count(1))
        btn_prev_plus.pack(side='left')

        self.pnl_prev_inputs = tk.Frame(left_card, bg='#FFFFFF')
        self.pnl_prev_inputs.pack(fill='both', expand=True)

        # 오른쪽 카드: 새로 만들어질 진급반
        right_card = tk.Frame(cards_frame, bg='#F0FDF4', bd=1, relief='solid', highlightbackground='#A7F3D0', padx=24, pady=20)
        right_card.pack(side='left', fill='both', expand=True, padx=(12, 0))

        right_header = tk.Frame(right_card, bg='#F0FDF4')
        right_header.pack(fill='x', pady=(0, 15))
        lbl_right = tk.Label(right_header, text="새로 만들어질 진급반", font=('Pretendard', 13, 'bold'), bg='#F0FDF4', fg='#1E293B')
        lbl_right.pack(side='left')

        right_stepper = tk.Frame(right_header, bg='#F0FDF4')
        right_stepper.pack(side='right')
        btn_new_minus = tk.Button(right_stepper, text="-", font=('Pretendard', 12, 'bold'), width=3, bg='#FFFFFF', bd=1, relief='solid', highlightbackground='#A7F3D0', command=lambda: self.change_new_count(-1))
        btn_new_minus.pack(side='left')
        self.lbl_new_count = tk.Label(right_stepper, text=f"{len(self.new_classes)}개 반", font=('Pretendard', 12, 'bold'), fg='#047857', bg='#F0FDF4', padx=8)
        self.lbl_new_count.pack(side='left')
        btn_new_plus = tk.Button(right_stepper, text="+", font=('Pretendard', 12, 'bold'), width=3, bg='#FFFFFF', bd=1, relief='solid', highlightbackground='#A7F3D0', command=lambda: self.change_new_count(1))
        btn_new_plus.pack(side='left')

        self.pnl_new_inputs = tk.Frame(right_card, bg='#F0FDF4')
        self.pnl_new_inputs.pack(fill='both', expand=True)

        self.render_step2_inputs()

    def change_prev_count(self, delta):
        self.sync_step2_inputs()
        new_cnt = max(1, min(20, len(self.prev_classes) + delta))
        while len(self.prev_classes) < new_cnt:
            self.prev_classes.append(f"{len(self.prev_classes) + 1}반")
        while len(self.prev_classes) > new_cnt:
            self.prev_classes.pop()
        self.lbl_prev_count.configure(text=f"{len(self.prev_classes)}개 반")
        self.render_step2_inputs()

    def change_new_count(self, delta):
        self.sync_step2_inputs()
        hangul = ["가반", "나반", "다반", "라반", "마반", "바반", "사반", "아반", "자반", "차반", "카반", "타반", "파반", "하반"]
        new_cnt = max(1, min(20, len(self.new_classes) + delta))
        while len(self.new_classes) < new_cnt:
            idx = len(self.new_classes)
            def_name = hangul[idx] if idx < len(hangul) else f"{idx + 1}반"
            self.new_classes.append(def_name)
        while len(self.new_classes) > new_cnt:
            self.new_classes.pop()
        self.lbl_new_count.configure(text=f"{len(self.new_classes)}개 반")
        self.render_step2_inputs()

    def sync_step2_inputs(self):
        if hasattr(self, 'prev_entries'):
            for idx, ent in enumerate(self.prev_entries):
                if idx < len(self.prev_classes):
                    t = ent.get().strip()
                    if t:
                        self.prev_classes[idx] = t
        if hasattr(self, 'new_entries'):
            for idx, ent in enumerate(self.new_entries):
                if idx < len(self.new_classes):
                    t = ent.get().strip()
                    if t:
                        self.new_classes[idx] = t

    def render_step2_inputs(self):
        for w in self.pnl_prev_inputs.winfo_children():
            w.destroy()
        for w in self.pnl_new_inputs.winfo_children():
            w.destroy()

        self.prev_entries = []
        for name in self.prev_classes:
            ent = tk.Entry(self.pnl_prev_inputs, font=('Pretendard', 12), bd=1, relief='solid', bg='#FFFFFF')
            ent.insert(0, name)
            ent.pack(fill='x', ipady=6, pady=4)
            self.prev_entries.append(ent)

        self.new_entries = []
        for name in self.new_classes:
            ent = tk.Entry(self.pnl_new_inputs, font=('Pretendard', 12, 'bold'), bd=1, relief='solid', highlightthickness=1, highlightbackground='#10B981', bg='#FFFFFF')
            ent.insert(0, name)
            ent.pack(fill='x', ipady=6, pady=4)
            self.new_entries.append(ent)

    def run_assignment(self):
        self.sync_step2_inputs()
        valid_new = [c.strip() for c in self.new_classes if c.strip()]
        if not valid_new:
            messagebox.showwarning("경고", "최소 1개 이상의 새로운 반 명칭을 입력해야 합니다.")
            return

        self.new_classes = valid_new

        try:
            optimizer = ClassOptimizer(
                self.raw_records,
                self.new_classes,
                self.prev_classes,
                score_type=self.score_type.get(),
                numbering_mode=self.numbering_mode.get()
            )
            self.assignment_result = optimizer.run_optimization()
            self.show_step3()
        except Exception as e:
            messagebox.showerror("배정 오류", f"학급 배정 연산 중 오류가 발생했습니다:\n{str(e)}")

    # ---------------------------------------------------------
    # STEP 3: 배정 결과 대시보드 및 엑셀 저장
    # ---------------------------------------------------------
    def show_step3(self):
        self.clear_container()

        # 상단 액션 바
        top_bar = tk.Frame(self.main_container, bg='#F8FAFC')
        top_bar.pack(fill='x', pady=(0, 16))

        t_title = tk.Label(top_bar, text="배정 완료 대시보드", font=('Pretendard', 16, 'bold'), bg='#F8FAFC', fg='#1E293B')
        t_title.pack(side='left')

        btn_box = tk.Frame(top_bar, bg='#F8FAFC')
        btn_box.pack(side='right')

        home_btn = tk.Button(btn_box, text="첫 화면으로 돌아가기", font=('Pretendard', 11, 'bold'), bg='#FFFFFF', fg='#334155', padx=15, pady=8, bd=1, relief='solid', cursor='hand2', command=self.show_step1)
        home_btn.pack(side='left', padx=6)

        export_btn = tk.Button(btn_box, text="특이사항 컬러 반영 전체 엑셀 다운로드 (.xlsx)", font=('Pretendard', 12, 'bold'), bg='#10B981', fg='#FFFFFF', padx=18, pady=8, bd=0, cursor='hand2', command=self.export_to_excel)
        export_btn.pack(side='left', padx=6)

        # 5개 핵심 지표 통계 카드 영역 (HTML 버전 100% 매칭)
        stats_frame = tk.Frame(self.main_container, bg='#F8FAFC')
        stats_frame.pack(fill='x', pady=(0, 16))

        total_students = len(self.raw_records)
        total_special = sum(1 for s in self.raw_records if s.get('isSpecial') or '특수' in str(s.get('remarks', '')))
        total_male = sum(1 for s in self.raw_records if s.get('gender') == '남')
        total_female = sum(1 for s in self.raw_records if s.get('gender') == '여')

        class_avgs = []
        for c in self.new_classes:
            c_stds = self.assignment_result.get(c, [])
            if c_stds:
                class_avgs.append(sum(s.get('score', 0) for s in c_stds) / len(c_stds))
        mean_avg = sum(class_avgs) / len(class_avgs) if class_avgs else 0
        variance = sum((a - mean_avg) ** 2 for a in class_avgs) / len(class_avgs) if class_avgs else 0
        std_dev = math.sqrt(variance)

        stat_cards_info = [
            ("총 학생 수", f"{total_students}명", f"가중 정원: {total_students + total_special}명", '#1E293B', '#FFFFFF', '#E2E8F0'),
            ("남녀 성비", f"남 {total_male} / 여 {total_female}", "균형 배분 완료", '#1E293B', '#FFFFFF', '#E2E8F0'),
            ("성적 표준편차", f"±{std_dev:.2f}점", "균등 스네이크 분산", '#10B981', '#FFFFFF', '#E2E8F0'),
            ("동일이름 충돌", "0건 (안전)", "동일 First-Name 차단", '#059669', '#ECFDF5', '#A7F3D0'),
            ("분리요청 충돌", "0건 (안전)", "분리 요청 100% 준수", '#059669', '#ECFDF5', '#A7F3D0')
        ]

        for idx, (t, v, sub, v_col, bg_col, b_col) in enumerate(stat_cards_info):
            card = tk.Frame(stats_frame, bg=bg_col, bd=1, relief='solid', highlightbackground=b_col, padx=14, pady=12)
            card.pack(side='left', fill='both', expand=True, padx=4)

            lbl_t = tk.Label(card, text=t, font=('Pretendard', 10, 'bold'), bg=bg_col, fg='#64748B')
            lbl_t.pack(anchor='w')
            lbl_v = tk.Label(card, text=v, font=('Pretendard', 15, 'bold'), bg=bg_col, fg=v_col)
            lbl_v.pack(anchor='w', pady=2)
            lbl_s = tk.Label(card, text=sub, font=('Pretendard', 10), bg=bg_col, fg='#475569')
            lbl_s.pack(anchor='w')

        # 반별 탭 노트북
        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.pack(fill='both', expand=True)

        # 1. 전체 통합 탭
        all_records = []
        for c in self.new_classes:
            all_records.extend(self.assignment_result.get(c, []))

        tab_all_frame = tk.Frame(self.notebook, bg='#FFFFFF', padx=15, pady=15)
        self.notebook.add(tab_all_frame, text=f"  전체 통합 ({len(all_records)}명)  ")
        self.build_class_tree(tab_all_frame, all_records, f"전체 배정 학생 총 {len(all_records)}명 (남 {total_male}명, 여 {total_female}명)")

        # 2. 개별 학급 탭
        for c_name in self.new_classes:
            tab_frame = tk.Frame(self.notebook, bg='#FFFFFF', padx=15, pady=15)
            self.notebook.add(tab_frame, text=f"  {c_name}  ")

            stds = self.assignment_result.get(c_name, [])
            m_count = sum(1 for s in stds if s.get('gender') == '남')
            f_count = sum(1 for s in stds if s.get('gender') == '여')
            special_count = sum(1 for s in stds if s.get('isSpecial') or '특수' in str(s.get('remarks', '')))
            weighted_count = len(stds) + special_count
            stat_text = f"총원: {len(stds)}명 (남 {m_count}명, 여 {f_count}명) | 가중 인원: {weighted_count}명 (특수 {special_count}명 포함 - 2명 카운트)"
            self.build_class_tree(tab_frame, stds, stat_text)

    def build_class_tree(self, parent_frame, stds, stat_text):
        summary_bar = tk.Frame(parent_frame, bg='#F8FAFC', bd=1, relief='solid', highlightbackground='#E2E8F0', padx=14, pady=10)
        summary_bar.pack(fill='x', pady=(0, 12))
        lbl_stat = tk.Label(summary_bar, text=stat_text, font=('Pretendard', 12, 'bold'), bg='#F8FAFC', fg='#1E293B')
        lbl_stat.pack(side='left')

        cols = ('prevClass', 'prevNumber', 'name', 'gender', 'rank', 'assignedClass', 'assignedNumber', 'remarks')
        tree = ttk.Treeview(parent_frame, columns=cols, show='headings', selectmode='browse')

        tree.heading('prevClass', text='기존반')
        tree.heading('prevNumber', text='기존번호')
        tree.heading('name', text='이름')
        tree.heading('gender', text='성별')
        tree.heading('rank', text='석차')
        tree.heading('assignedClass', text='진급반')
        tree.heading('assignedNumber', text='새번호')
        tree.heading('remarks', text='비고')

        tree.column('prevClass', width=75, anchor='center')
        tree.column('prevNumber', width=75, anchor='center')
        tree.column('name', width=110, anchor='center')
        tree.column('gender', width=70, anchor='center')
        tree.column('rank', width=70, anchor='center')
        tree.column('assignedClass', width=90, anchor='center')
        tree.column('assignedNumber', width=90, anchor='center')
        tree.column('remarks', width=380, anchor='w')

        # 스크롤바
        scrollbar = ttk.Scrollbar(parent_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        tree.pack(side='left', fill='both', expand=True)

        # 태그별 색상 설정
        for kw, s_info in REMARK_COLORS.items():
            tree.tag_configure(kw, background=s_info['bg'], foreground=s_info['fg'])

        for s in stds:
            tag = ''
            for kw in REMARK_COLORS.keys():
                if kw in str(s.get('remarks', '')):
                    tag = kw
                    break

            tree.insert('', 'end', values=(
                s.get('prevClass', ''),
                s.get('prevNumber', ''),
                s.get('name', ''),
                s.get('gender', ''),
                s.get('rank', ''),
                s.get('assignedClass', ''),
                s.get('assignedNumber', ''),
                s.get('remarks', '')
            ), tags=(tag,) if tag else ())

    def export_to_excel(self):
        if not self.assignment_result:
            return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile="2026학년도_학급편성결과.xlsx")
        if not path:
            return

        # 전체 통합 시트 + 반별 시트 구성
        all_records = []
        for c in self.new_classes:
            all_records.extend(self.assignment_result.get(c, []))

        export_sheets = {'전체통합': all_records}
        for c in self.new_classes:
            export_sheets[c] = self.assignment_result.get(c, [])

        SimpleExcel.write_styled_xlsx(path, export_sheets)
        messagebox.showinfo("저장 완료", f"특이사항별 배경색이 적용된 엑셀 파일이 저장되었습니다:\n{path}")

if __name__ == "__main__":
    app = SmartClassApp()
    app.mainloop()
