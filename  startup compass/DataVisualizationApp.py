# 필요한 모듈 import
import os
import re
import tkinter as tk
from tkinter import ttk,scrolledtext
import webbrowser
import threading
import time


# 데이터 시각화를 위한 라이브러리
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import pandas as pd
import requests
import json
from tkinter import messagebox
from urllib.parse import quote_plus

# matplotlib 설정
import matplotlib.font_manager as fm


# 한글 폰트 설정 (macOS용)
# plt.rcParams['font.family'] = ['AppleGothic']  # macOS
# 또는 시스템에 설치된 한글 폰트 사용
#plt.rcParams['font.family'] = ['Malgun Gothic']  # Windows
# plt.rcParams['font.family'] = ['NanumGothic']    # Linux

# 마이너스 기호 깨짐 방지
# plt.rcParams['axes.unicode_minus'] = False



# 한글 폰트 설정 (macOS용)
try:
    # 시스템에 설치된 적절한 한글 폰트 이름으로 설정
    # 예: 'NanumGothic', 'Malgun Gothic' (Windows), 'AppleSDGothicNeo-Regular' (최신 macOS)
    font_name = fm.FontProperties(family='AppleGothic').get_name() # 설치된 경우 해당 폰트 사용
    plt.rcParams['font.family'] = font_name
except RuntimeError: # AppleGothic이 없을 경우 기본 폰트 사용 경고
    print("경고: AppleGothic 폰트를 찾을 수 없습니다. 그래프의 한글이 깨질 수 있습니다.")
    # 다른 사용 가능한 한글 폰트 시도 또는 시스템 기본값 사용
    # 예: plt.rcParams['font.family'] = 'NanumGothic'

plt.rcParams['axes.unicode_minus'] = False # 마이너스 기호 깨짐 방지



class DataVisualizationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("데이터 시각화 대시보드")
        self.root.geometry("1200x800")
        
        # 컨테이너 프레임 생성
        self.container = tk.Frame(root)
        self.container.pack(side="top", fill="both", expand=True)
        
        # 그리드 설정
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)
        
        # 페이지 저장을 위한 딕셔너리
        self.frames = {}
        
        # 페이지 클래스 리스트
        pages = (MainPage, AgeClosurePage, FranchiseClosurePage, RegionalClosurePage,
                BusinessSurvivalPage, NewBusinessPage, SearchStatsPage)
        
        # 각 페이지 초기화 및 프레임 딕셔너리에 추가
        for F in pages:
            frame = F(self.container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        
        # 메인 페이지 표시
        self.show_frame(MainPage)
    
    def show_frame(self, cont):
        """특정 프레임을 앞으로 가져옴"""
        frame = self.frames[cont]
        frame.tkraise()


##########################################################################################
class BasePage(tk.Frame):
    """모든 페이지의 기본 클래스"""
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg="white")
        self.controller = controller
        
        # 상단 헤더 프레임
        header_frame = tk.Frame(self, bg="white", height=60)
        header_frame.pack(fill=tk.X, pady=10)
        
        # 로고 (왼쪽)
        logo_label = tk.Label(header_frame, text="β", font=("Arial", 24, "bold"), bg="white")
        logo_label.pack(side=tk.LEFT, padx=20)
        
        # 링크들 (오른쪽)
        self.create_hyperlink(header_frame, "Link", "#", side=tk.RIGHT)
        self.create_hyperlink(header_frame, "국세청", "https://www.nts.go.kr/", side=tk.RIGHT)
        self.create_hyperlink(header_frame, "naver data lab", "https://datalab.naver.com/", side=tk.RIGHT)
        self.create_hyperlink(header_frame, "공공데이터 포털", "https://www.data.go.kr/tcs/dss/selectDataSetList.do?keyword=%ED%8F%90%EC%97%85&brm=&svcType=&recmSe=N&conditionType=init&extsn=&kwrdArray=", side=tk.RIGHT)
        self.create_hyperlink(header_frame, "KOSIS", "https://kosis.kr/search/search.do?query=%ED%8F%90%EC%97%85", side=tk.RIGHT)
        
        # 구분선
        separator = ttk.Separator(self, orient="horizontal")
        separator.pack(fill=tk.X, pady=5)
        
    def create_hyperlink(self, parent, text, url, side=tk.LEFT):
        """클릭 가능한 하이퍼링크 생성"""
        link = tk.Label(parent, text=text, fg="black", cursor="hand2", 
                        font=("Arial", 10, "underline"), bg="white")
        link.pack(side=side, padx=10)
        
        # 마우스 이벤트 바인딩
        link.bind("<Enter>", lambda e: link.config(fg="blue"))
        link.bind("<Leave>", lambda e: link.config(fg="black"))
        link.bind("<Button-1>", lambda e: webbrowser.open_new(url))
        
        return link
    

    ##########################################################################################

class MainPage(BasePage):
    """메인 페이지 클래스"""
    def __init__(self, parent, controller):
        BasePage.__init__(self, parent, controller)
        
        # 메인 제목
        title_frame = tk.Frame(self, bg="white")
        title_frame.pack(fill=tk.X, pady=10, padx=20, anchor="w")
        
        title_label = tk.Label(title_frame, text="메뉴", font=("Arial", 18, "bold"), bg="white")
        title_label.pack(side=tk.LEFT)
        
        subtitle_label = tk.Label(title_frame, text="클릭시 해당 페이지로 넘어갑니다.", bg="white")
        subtitle_label.pack(side=tk.LEFT, padx=10)
        
        # 카드 컨테이너 프레임
        cards_frame = tk.Frame(self, bg="white")
        cards_frame.pack(pady=20, fill=tk.BOTH, expand=True)
        
        # 그리드 설정
        for i in range(2):
            cards_frame.grid_rowconfigure(i, weight=1)
        for i in range(3):
            cards_frame.grid_columnconfigure(i, weight=1)
        
        # 카드 생성
        self.create_card(cards_frame, "연령별 폐업", "2013-2023", "연령,성별,업태", 
                         0, 0, AgeClosurePage)
        self.create_card(cards_frame, "프랜차이즈 폐업", "아직 안정해짐", "Description", 
                         0, 1, FranchiseClosurePage)
        self.create_card(cards_frame, "지역별 폐업사유", "2005-2023", "폐업사유,지역,업태", 
                         0, 2, RegionalClosurePage)
        self.create_card(cards_frame, "사업 유지율", "2005-2023", "존속,연수,지역,업태", 
                         1, 0, BusinessSurvivalPage)
        self.create_card(cards_frame, "신규 사업자", "2013-2023", "연령,성별,지역", 
                         1, 1, NewBusinessPage)
        self.create_card(cards_frame, "검색 통계", "네이버 검색 통계", "Naver data lab", 
                         1, 2, SearchStatsPage)
    
    def create_card(self, parent, title, date_range, description, row, col, target_page):
        """카드 위젯 생성"""
        card = tk.Frame(parent, bd=1, relief=tk.RAISED, bg="white")
        card.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
        
        # 카드 내용
        title_label = tk.Label(card, text=title, font=("Arial", 16, "bold"), bg="white")
        title_label.pack(pady=(20, 5))
        
        date_label = tk.Label(card, text=date_range, bg="white")
        date_label.pack(pady=2)
        
        desc_label = tk.Label(card, text=description, bg="white")
        desc_label.pack(pady=(2, 20))
        
        # 페이지 전환 기능 추가
        if target_page:
            card.bind("<Button-1>", lambda e, page=target_page: self.controller.show_frame(page))
            card.bind("<Enter>", lambda e: card.config(bg="#F5F5F5", cursor="hand2"))
            card.bind("<Leave>", lambda e: card.config(bg="white", cursor=""))
            
            # 모든 자식 위젯에도 동일한 이벤트 바인딩
            for child in card.winfo_children():
                child.bind("<Button-1>", lambda e, page=target_page: self.controller.show_frame(page))
                child.bind("<Enter>", lambda e: card.config(bg="#F5F5F5", cursor="hand2"))
                child.bind("<Leave>", lambda e: card.config(bg="white", cursor=""))

##############################################################################################################################                

class DetailPage(BasePage):
    """상세 페이지의 기본 클래스"""
    def __init__(self, parent, controller, title, subtitle=""):
        BasePage.__init__(self, parent, controller)
        
        # 정보 아이콘과 제목
        info_frame = tk.Frame(self, bg="white")
        info_frame.pack(fill=tk.X, pady=20, padx=20, anchor="w")
        
        info_icon = tk.Label(info_frame, text="ⓘ", font=("Arial", 24), bg="white")
        info_icon.pack(side=tk.LEFT, padx=(0, 10))
        
        title_label = tk.Label(info_frame, text=title, font=("Arial", 18, "bold"), bg="white")
        title_label.pack(side=tk.LEFT)
        
        if subtitle:
            subtitle_label = tk.Label(info_frame, text=subtitle, bg="white")
            subtitle_label.pack(side=tk.LEFT, padx=10)
        
        # 뒤로 가기 버튼
        # back_button = tk.Button(self, text="메인으로 돌아가기", 
        #                        command=lambda: controller.show_frame(MainPage))
        # back_button.pack(pady=10)
        
        # 콘텐츠 프레임
        self.content_frame = tk.Frame(self, bg="white")
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        


##############################################################################################################################
class AgeClosurePage(DetailPage):
    """연령별 폐업 페이지"""

    def go_to_main(self):
        """메인 페이지로 돌아가기"""
        self.controller.show_frame(MainPage)

    def __init__(self, parent, controller):
        DetailPage.__init__(self, parent, controller,
                          " 4년간 연령별 폐업률 통계", "연 성별 업태")

        # 한글 폰트 설정
        plt.rcParams['font.family'] = ['AppleGothic']  # macOS용
        plt.rcParams['axes.unicode_minus'] = False

        # 컨트롤 프레임
        self.control_frame = tk.Frame(self.content_frame, bg="white")
        self.control_frame.pack(fill=tk.X, pady=10)

        # 시작 년도 선택
        tk.Label(self.control_frame, text="시작 년도:", bg="white").pack(side=tk.LEFT, padx=5)
        self.start_year_var = tk.StringVar(value="2020")
        self.start_year_combo = ttk.Combobox(self.control_frame, textvariable=self.start_year_var,
                                           values=[str(year) for year in range(2019, 2023)], width=10)
        self.start_year_combo.pack(side=tk.LEFT, padx=5)

        # 종료 년도 선택
        tk.Label(self.control_frame, text="종료 년도:", bg="white").pack(side=tk.LEFT, padx=5)
        self.end_year_var = tk.StringVar(value="2023")
        self.end_year_combo = ttk.Combobox(self.control_frame, textvariable=self.end_year_var,
                                         values=[str(year) for year in range(2019, 2023)], width=10)
        self.end_year_combo.pack(side=tk.LEFT, padx=5)

        # 조회 버튼
        self.search_button = tk.Button(self.control_frame, text="조회", command=self.load_data)
        self.search_button.pack(side=tk.LEFT, padx=10)
        # 메인으로 돌아가기 버튼
        self.home_button = tk.Button(self.control_frame, text="메인으로 돌아가기",
                               command=self.go_to_main,
                               bg="#2196F3", font=("Arial", 10, "bold"),
                               relief="raised", borderwidth=2)
        self.home_button.pack(side=tk.LEFT, padx=5)

        # 테이블 프레임 (스크롤 가능)
        self.table_frame = tk.Frame(self.content_frame, bg="white")
        self.table_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 정보 표시 프레임 (한 번만 생성)
        self.info_frame = tk.Frame(self.content_frame, bg="white")
        self.info_frame.pack(fill=tk.X, pady=5)

        # 초기 데이터 로드
        # self.load_data()

    def load_data(self):
        """KOSIS API에서 데이터를 로드하여 테이블에 표시"""
        try:
            # 기존 테이블 위젯 제거
            for widget in self.table_frame.winfo_children():
                widget.destroy()

            # 기존 정보 프레임 내용 제거 (중복 방지)
            for widget in self.info_frame.winfo_children():
                widget.destroy()

            # 년도 유효성 검사 - 종료 년도 제대로 인식
            start_year = self.start_year_var.get()
            end_year = self.end_year_var.get()

            print(f"선택된 시작 년도: {start_year}")
            print(f"선택된 종료 년도: {end_year}")

            if int(start_year) > int(end_year):
                messagebox.showerror("오류", "시작 년도가 종료 년도보다 클 수 없습니다.")
                return

            # 로딩 메시지 표시
            loading_label = tk.Label(self.table_frame, text="데이터를 불러오는 중...", bg="white")
            loading_label.pack(pady=50)
            self.table_frame.update()

            # 샘플 데이터 생성 (API 대신)
            data = self.create_sample_data(start_year, end_year)

            # 로딩 메시지 제거
            loading_label.destroy()

            # 행렬 형태 테이블 생성
            self.create_matrix_table(data)

            # 정보 표시
            count_label = tk.Label(self.info_frame,
                                 text=f"총 {len(data)}개의 데이터를 행렬 형태로 표시",
                                 bg="white", font=("Arial", 9))
            count_label.pack(side=tk.LEFT)

            info_label = tk.Label(self.info_frame,
                                text="19개 업태별 폐업자 수와 폐업률 통계 (예: 12,345명 | 13.7%)",
                                bg="white", font=("Arial", 9), fg="blue")
            info_label.pack(side=tk.RIGHT)

        except Exception as e:
            # 로딩 메시지 제거
            for widget in self.table_frame.winfo_children():
                widget.destroy()
            print(f"오류 발생: {e}")

            # 오류 시에도 샘플 데이터 표시
            sample_data = self.create_sample_data(start_year, end_year)
            self.create_matrix_table(sample_data)

    def create_sample_data(self, start_year, end_year):
        """종료 년도까지 포함한 샘플 데이터 생성 - 폐업자 수와 폐업률 함께 생성"""
        sample_data = []
        years = list(range(int(start_year), int(end_year) + 1))
        regions = ['전국']
        age_groups = ['30세 미만', '30세 이상', '40세 이상']
        genders = ['합계', '남자', '여자']

        # 실제 KOSIS 업태 분류를 참고한 다양한 업태
        business_types = [
            '농업, 임업 및 어업',
            '광업',
            '제조업',
            '전기, 가스, 증기 및 공기조절 공급업',
            '수도, 하수 및 폐기물 처리, 원료 재생업',
            '건설업',
            '도매 및 소매업',
            '운수 및 창고업',
            '숙박 및 음식점업',
            '정보통신업',
            '금융 및 보험업',
            '부동산업',
            '전문, 과학 및 기술 서비스업',
            '사업시설 관리, 사업 지원 및 임대 서비스업',
            '공공행정, 국방 및 사회보장 행정',
            '교육 서비스업',
            '보건업 및 사회복지 서비스업',
            '예술, 스포츠 및 여가관련 서비스업',
            '협회 및 단체, 수리 및 기타 개인 서비스업'
        ]

        print(f"샘플 데이터 생성: {start_year}년 ~ {end_year}년")
        print(f"업태 개수: {len(business_types)}개")

        for year in years:
            for region in regions:
                for business in business_types:
                    for age in age_groups:
                        for gender in genders:
                            import random

                            # 업태별로 다른 폐업률 범위 설정 (현실적 반영)
                            if '음식점' in business or '숙박' in business:
                                # 음식점, 숙박업은 폐업률이 높음
                                base_rate = random.uniform(20.0, 35.0)
                                base_count = random.randint(15000, 45000)
                            elif '제조업' in business or '건설업' in business:
                                # 제조업, 건설업은 중간 정도
                                base_rate = random.uniform(10.0, 20.0)
                                base_count = random.randint(8000, 25000)
                            elif '정보통신' in business or '금융' in business:
                                # IT, 금융업은 상대적으로 낮음
                                base_rate = random.uniform(5.0, 15.0)
                                base_count = random.randint(3000, 12000)
                            elif '도매' in business or '소매' in business:
                                # 도소매업은 높은 편
                                base_rate = random.uniform(15.0, 25.0)
                                base_count = random.randint(25000, 55000)
                            else:
                                # 기타 업종
                                base_rate = random.uniform(8.0, 22.0)
                                base_count = random.randint(5000, 20000)

                            # 성별에 따른 미세 조정
                            if gender == '여자':
                                closure_rate = round(base_rate + random.uniform(0, 3.0), 1)
                                closure_count = int(base_count * random.uniform(0.3, 0.5))  # 여성 비율 30-50%
                            elif gender == '남자':
                                closure_rate = round(base_rate - random.uniform(0, 2.0), 1)
                                closure_count = int(base_count * random.uniform(0.4, 0.6))  # 남성 비율 40-60%
                            else:  # 합계
                                closure_rate = round(base_rate, 1)
                                closure_count = base_count

                            # 최소값 보정
                            closure_rate = max(closure_rate, 1.0)
                            closure_count = max(closure_count, 100)

                            # 폐업자 수와 폐업률을 함께 표시하는 형태로 데이터 생성
                            combined_value = f"{closure_count:,}명 | {closure_rate}%"

                            sample_data.append({
                                'PRD_DE': str(year),
                                'C1_NM': region,
                                'C2_NM': age,
                                'C3_NM': gender,
                                'C4_NM': business,
                                'ITM_NM': '폐업현황',
                                'DT': combined_value,
                                'UNIT_NM': ''  # 단위는 이미 값에 포함되어 있음
                            })

        print(f"생성된 샘플 데이터 개수: {len(sample_data)}")
        return sample_data

    def create_matrix_table(self, data):
        """행렬 형태의 테이블 생성 (KOSIS 스타일)"""
        # 데이터 구조 분석 및 행렬 형태로 변환
        matrix_data = self.prepare_matrix_data(data)

        # 메인 컨테이너 프레임
        main_container = tk.Frame(self.table_frame, bg="white")
        main_container.pack(fill=tk.BOTH, expand=True)

        # 1. 헤더를 Canvas에 올림
        header_canvas = tk.Canvas(main_container, height=70, bg="white", highlightthickness=0)
        header_canvas.pack(fill=tk.X, side=tk.TOP, anchor='n')

        header_frame = tk.Frame(header_canvas, bg="white")
        header_canvas.create_window((0, 0), window=header_frame, anchor='nw')

        # 2. 데이터 영역도 Canvas
        data_container = tk.Frame(main_container, bg="white")
        data_container.pack(fill=tk.BOTH, expand=True, side=tk.BOTTOM)

        data_canvas = tk.Canvas(data_container, bg="white", highlightthickness=0)
        scrollbar_h = ttk.Scrollbar(data_container, orient="horizontal")
        scrollbar_v = ttk.Scrollbar(data_container, orient="vertical")

        scrollable_frame = tk.Frame(data_canvas, bg="white")
        data_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        # 스크롤 동기화 설정
        data_canvas.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)
        header_canvas.configure(xscrollcommand=scrollbar_h.set)
        scrollbar_h.config(command=lambda *args: [data_canvas.xview(*args), header_canvas.xview(*args)])
        scrollbar_v.config(command=data_canvas.yview)

        # 배치 (data_container 내부는 grid만!)
        data_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar_v.grid(row=0, column=1, sticky="ns")
        scrollbar_h.grid(row=1, column=0, columnspan=2, sticky="ew")

        data_container.grid_rowconfigure(0, weight=1)
        data_container.grid_columnconfigure(0, weight=1)
        # 오른쪽 아래 빈 공간(코너)은 필요시 비워둠
        data_container.grid_rowconfigure(0, weight=1)
        data_container.grid_columnconfigure(0, weight=1)

        # 헤더/데이터 프레임 크기 동기화
        def sync_header_width(event=None):
            header_canvas.configure(scrollregion=header_canvas.bbox("all"))
            data_canvas.configure(scrollregion=data_canvas.bbox("all"))
            header_canvas.itemconfig('all', width=scrollable_frame.winfo_width())

        scrollable_frame.bind("<Configure>", sync_header_width)
        header_frame.bind("<Configure>", sync_header_width)

        # 3. 헤더/데이터 생성
        self.create_fixed_header(header_frame, matrix_data)
        self.create_data_rows(scrollable_frame, matrix_data)


    def prepare_matrix_data(self, data):
        """데이터를 행렬 형태로 변환"""
        matrix = {}
        years = sorted(list(set([item.get('PRD_DE', '') for item in data])))

        for item in data:
            business = item.get('C4_NM', '')
            age = item.get('C2_NM', '')
            gender = item.get('C3_NM', '')
            year = item.get('PRD_DE', '')
            value = item.get('DT', '')
            unit = item.get('UNIT_NM', '')

            if business not in matrix:
                matrix[business] = {}
            if age not in matrix[business]:
                matrix[business][age] = {}
            if gender not in matrix[business][age]:
                matrix[business][age][gender] = {}

            matrix[business][age][gender][year] = f"{value}{unit}"

        return matrix, years

    def create_fixed_header(self, parent, matrix_data):
        """고정 헤더 생성"""
        matrix, years = matrix_data

        # 헤더 스타일 설정
        header_style = {
            'bg': '#E6F3FF',
            'relief': 'solid',
            'borderwidth': 1,
            'font': ('Arial', 9, 'bold'),
            'anchor': 'center'
        }

        row = 0

        # 최상단 헤더 (년도별 구분)
        tk.Label(parent, text="업태별", width=25, height=3, **header_style).grid(
            row=row, column=0, rowspan=3, sticky="nsew")
        tk.Label(parent, text="연령별", width=12, height=3, **header_style).grid(
            row=row, column=1, rowspan=3, sticky="nsew")

        # 년도 헤더
        col_start = 2
        for year in years:
            tk.Label(parent, text=f"{year}년", width=60, **header_style).grid(
                row=row, column=col_start, columnspan=6, sticky="nsew")
            col_start += 6

        row += 1

        # 성별 헤더
        col_start = 2
        for year in years:
            genders = ['합계', '남자', '여자']
            for gender in genders:
                tk.Label(parent, text=gender, width=20, **header_style).grid(
                    row=row, column=col_start, columnspan=2, sticky="nsew")
                col_start += 2

        row += 1

        # 구분 헤더 (폐업자수/폐업률)
        col_start = 2
        for year in years:
            for gender in ['합계', '남자', '여자']:
                tk.Label(parent, text="폐업자수", width=10, **header_style).grid(
                    row=row, column=col_start, sticky="nsew")
                tk.Label(parent, text="폐업률", width=10, **header_style).grid(
                    row=row, column=col_start+1, sticky="nsew")
                col_start += 2

    def create_data_rows(self, parent, matrix_data):
        """데이터 행만 생성 (헤더 제외)"""
        matrix, years = matrix_data

        cell_style = {
            'bg': 'white',
            'relief': 'solid',
            'borderwidth': 1,
            'font': ('Arial', 8),
            'anchor': 'center'
        }

        row = 0

        # 데이터 행 생성
        for business_type in sorted(matrix.keys()):
            business_data = matrix[business_type]
            age_groups = sorted(business_data.keys())

            for i, age_group in enumerate(age_groups):
                age_data = business_data[age_group]

                # 첫 번째 연령대일 때만 업태명 표시
                if i == 0:
                    business_label = tk.Label(parent, text=business_type, width=25, **cell_style)
                    business_label.grid(row=row, column=0, rowspan=len(age_groups), sticky="nsew")

                # 연령대 명
                age_label = tk.Label(parent, text=age_group, width=12, **cell_style)
                age_label.grid(row=row, column=1, sticky="nsew")

                # 데이터 셀 - 폐업자 수와 폐업률을 별도 열로 분리
                col = 2
                for year in years:
                    genders = ['합계', '남자', '여자']
                    for gender in genders:
                        original_value = age_data.get(gender, {}).get(year, '-')

                        # 원본 데이터에서 폐업자 수와 폐업률 분리
                        if original_value != '-' and '|' in original_value:
                            parts = original_value.split('|')
                            count_part = parts[0].strip()  # "12,345명"
                            rate_part = parts[1].strip()   # "13.7%"
                        else:
                            count_part = '-'
                            rate_part = '-'

                        # 폐업자 수 열
                        count_label = tk.Label(parent, text=count_part, width=10, **cell_style)
                        count_label.grid(row=row, column=col, sticky="nsew")

                        # 폐업률 열
                        rate_label = tk.Label(parent, text=rate_part, width=10, **cell_style)
                        rate_label.grid(row=row, column=col+1, sticky="nsew")

                        col += 2

                row += 1

    def format_number(self, value):
        """숫자 값에 천 단위 콤마 추가"""
        if value is None or value == '':
            return ''

        try:
            num_str = str(value).replace(',', '')
            if '.' in num_str:
                return f"{float(num_str):,.1f}"
            elif num_str.replace('-', '').isdigit():
                return f"{int(num_str):,}"
        except (ValueError, AttributeError):
            pass

        return str(value)





##############################################################################################################################


class FranchiseClosurePage(DetailPage):
    """프랜차이즈 정보 페이지"""
    
    def __init__(self, parent, controller):
        DetailPage.__init__(self, parent, controller, 
                          "프랜차이즈 브랜드별 가맹점 현황", "가맹점 20개 이상 브랜드")
        
        # 인코딩된 API 키 사용
        self.api_key = "Lv%2BNxXs1IsNjEducoENMAKDqnaeyobF%2FpoARIpgmiLwpk8%2FlWClbx7HSCNx1s8%2BD4OEWwD3%2BDz6mefsI6U1OYg%3D%3D"
        self.base_url = "https://apis.data.go.kr/1130000/FftcBrandFrcsStatsService"
        
        # 컨트롤 프레임
        self.control_frame = tk.Frame(self.content_frame, bg="white")
        self.control_frame.pack(fill=tk.X, pady=10)
        
        # 년도 선택
        tk.Label(self.control_frame, text="조회 년도:", bg="white").pack(side=tk.LEFT, padx=5)
        self.year_var = tk.StringVar(value="2023")
        self.year_combo = ttk.Combobox(self.control_frame, textvariable=self.year_var, 
                                      values=[str(year) for year in range(2020, 2025)], width=10)
        self.year_combo.pack(side=tk.LEFT, padx=5)
        
        # 조회 버튼 (AgeClosurePage와 동일한 스타일)
        self.search_button = tk.Button(self.control_frame, text="조회", command=self.load_franchise_data)
        self.search_button.pack(side=tk.LEFT, padx=10)
        
        # 메인으로 돌아가기 버튼 (AgeClosurePage와 동일한 스타일)
        self.home_button = tk.Button(self.control_frame, text="메인으로 돌아가기", 
                               command=lambda: controller.show_frame(MainPage),
                               bg="#2196F3", font=("Arial", 10, "bold"),
                               relief="raised", borderwidth=2)
        self.home_button.pack(side=tk.LEFT, padx=5)
        
        # 테이블 프레임
        self.table_frame = tk.Frame(self.content_frame, bg="white")
        self.table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 정보 표시 프레임
        self.info_frame = tk.Frame(self.content_frame, bg="white")
        self.info_frame.pack(fill=tk.X, pady=5)
        
        # 초기 안내 메시지 표시 (자동 로드 제거)
        self.show_initial_message()
    
    def show_initial_message(self):
        """초기 안내 메시지 표시"""
        initial_label = tk.Label(self.table_frame, 
                               text="년도를 선택하고 '조회' 버튼을 클릭하여 프랜차이즈 데이터를 불러오세요.", 
                               bg="white", font=("Arial", 14), fg="gray")
        initial_label.pack(pady=100)
        
        info_label = tk.Label(self.info_frame, 
                            text="가맹점 20개 이상인 브랜드만 표시됩니다.", 
                            bg="white", font=("Arial", 10), fg="blue")
        info_label.pack()
    
    def load_franchise_data(self):
        """공정거래위원회 API에서 프랜차이즈 데이터 로드"""
        try:
            # 기존 위젯 제거
            for widget in self.table_frame.winfo_children():
                widget.destroy()
            for widget in self.info_frame.winfo_children():
                widget.destroy()
            
            # 로딩 메시지
            loading_label = tk.Label(self.table_frame, text="데이터를 불러오는 중...", 
                                   bg="white", font=("Arial", 14))
            loading_label.pack(pady=50)
            self.table_frame.update()
            
            year = self.year_var.get()
            
            # API 호출
            franchise_data = self.fetch_franchise_api_data(year)
            
            # 로딩 메시지 제거
            loading_label.destroy()
            
            if franchise_data and len(franchise_data) > 0:
                # 가맹점 20개 이상인 브랜드만 필터링
                filtered_data = [item for item in franchise_data if int(item.get('frcsCnt', 0)) >= 20]
                
                if len(filtered_data) > 0:
                    # 테이블 생성
                    self.create_franchise_table(filtered_data)
                    
                    # 정보 표시
                    total_count = len(filtered_data)
                    info_label = tk.Label(self.info_frame, 
                                        text=f"가맹점 20개 이상 브랜드: {total_count}개 (전체: {len(franchise_data)}개)", 
                                        bg="white", font=("Arial", 9))
                    info_label.pack(side=tk.LEFT)
                    
                    filter_info = tk.Label(self.info_frame, 
                                         text="※ 가맹점 수 기준 내림차순 정렬", 
                                         bg="white", font=("Arial", 9), fg="blue")
                    filter_info.pack(side=tk.RIGHT)
                else:
                    # 필터링된 데이터가 없는 경우
                    no_data_label = tk.Label(self.table_frame, 
                                           text=f"가맹점 20개 이상인 브랜드가 없습니다.\n(전체 {len(franchise_data)}개 브랜드 조회됨)", 
                                           bg="white", font=("Arial", 14), fg="orange")
                    no_data_label.pack(pady=50)
            else:
                # API 실패 시 오류 메시지
                error_label = tk.Label(self.table_frame, 
                                     text="API에서 데이터를 가져올 수 없습니다.\n네트워크 연결을 확인해주세요.", 
                                     bg="white", font=("Arial", 14), fg="red")
                error_label.pack(pady=50)
                
        except Exception as e:
            print(f"데이터 로드 오류: {e}")
            # 오류 시 오류 메시지 표시
            for widget in self.table_frame.winfo_children():
                widget.destroy()
            error_label = tk.Label(self.table_frame, 
                                 text=f"데이터 로드 중 오류가 발생했습니다.\n오류: {str(e)}", 
                                 bg="white", font=("Arial", 12), fg="red")
            error_label.pack(pady=50)
    
    def fetch_franchise_api_data(self, year):
        """URL을 직접 구성하여 API 호출"""
        try:
            # 인코딩된 키를 사용하여 URL 직접 구성
            url = f"{self.base_url}/getBrandFrcsStats?serviceKey={self.api_key}&pageNo=1&numOfRows=1000&resultType=json&yr={year}"
            
            print(f"요청 URL: {url}")
            
            # 직접 구성된 URL로 GET 요청
            response = requests.get(url, timeout=15)
            
            print(f"응답 상태: {response.status_code}")
            print(f"응답 내용: {response.text[:300]}...")
            
            if response.status_code == 200:
                try:
                    data = response.json()
                    return self.parse_json_response(data, year)
                except json.JSONDecodeError as e:
                    print(f"JSON 파싱 실패: {e}")
                    return None
            else:
                print(f"API 호출 실패: {response.status_code}")
                return None
                
        except Exception as e:
            print(f"API 호출 오류: {e}")
            return None
    
    def parse_json_response(self, data, year):
        """JSON 응답 파싱"""
        try:
            print("=== API 응답 전체 구조 ===")
            print(f"Keys: {list(data.keys())}")
            print(f"resultCode: {data.get('resultCode')}")
            print(f"resultMsg: {data.get('resultMsg')}")
            print(f"totalCount: {data.get('totalCount')}")
            
            # API 응답 구조에 맞게 수정
            if 'resultCode' in data and data['resultCode'] == '00':
                if 'items' in data:
                    items = data['items']
                    print(f"items 타입: {type(items)}")
                    print(f"items 길이: {len(items) if isinstance(items, list) else 'N/A'}")
                    
                    if isinstance(items, list) and len(items) > 0:
                        print(f"첫 번째 아이템 키: {list(items[0].keys())}")
                        print(f"첫 번째 아이템 샘플: {items[0]}")
                        
                        franchise_list = []
                        for item in items:
                            franchise_info = {
                                'brandNm': item.get('brandNm', ''),
                                'corpNm': item.get('corpNm', ''),
                                'frcsCnt': item.get('frcsCnt', '0'),
                                'avrgSlsAmt': item.get('avrgSlsAmt', '0'),
                                'arUnitAvrgSlsAmt': item.get('arUnitAvrgSlsAmt', '0'),
                                'ctrtEndCnt': item.get('ctrtEndCnt', '0'),
                                'ctrtCncltnCnt': item.get('ctrtCncltnCnt', '0'),
                                'newFrcsRgsCnt': item.get('newFrcsRgsCnt', '0'),
                                'nmChgCnt': item.get('nmChgCnt', '0'),
                                'indutyLclasNm': item.get('indutyLclasNm', ''),
                                'indutyMlsfcNm': item.get('indutyMlsfcNm', ''),
                                'yr': item.get('yr', year)
                            }
                            franchise_list.append(franchise_info)
                        
                        # 가맹점 수 기준 내림차순 정렬
                        franchise_list.sort(key=lambda x: int(x.get('frcsCnt', 0)), reverse=True)
                        print(f"파싱된 데이터 개수: {len(franchise_list)}")
                        return franchise_list
                    else:
                        print("items가 비어있거나 리스트가 아닙니다.")
                else:
                    print("'items' 키를 찾을 수 없습니다.")
            else:
                print(f"API 오류: resultCode = {data.get('resultCode', 'unknown')}")
            
            return None
            
        except Exception as e:
            print(f"JSON 파싱 오류: {e}")
            return None
    
    def create_franchise_table(self, data):
        """프랜차이즈 데이터 테이블 생성 - 수정된 버전"""
        # 스크롤 가능한 테이블 프레임
        canvas = tk.Canvas(self.table_frame, bg="white")
        scrollbar_v = ttk.Scrollbar(self.table_frame, orient="vertical", command=canvas.yview)
        scrollbar_h = ttk.Scrollbar(self.table_frame, orient="horizontal", command=canvas.xview)
        scrollable_frame = tk.Frame(canvas, bg="white")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)
        
        # 테이블 헤더 (평균매출, 면적당매출 제거, 폐업률 추가)
        headers = ['순위', '브랜드명', '운영회사', '가맹점수', '계약종료', '계약해지', '신규등록', '폐업률(%)', '업종대분류']
        header_widths = [5, 18, 25, 10, 8, 8, 8, 10, 15]
        
        # 헤더 스타일 (AgeClosurePage와 동일한 색상)
        header_style = {
            'bg': '#4A5C90',
            'fg': 'white',
            'relief': 'solid',
            'borderwidth': 1,
            'font': ('Arial', 9, 'bold'),
            'anchor': 'center'
        }
        
        # 헤더 생성
        for col, (header, width) in enumerate(zip(headers, header_widths)):
            header_label = tk.Label(scrollable_frame, text=header, width=width, **header_style)
            header_label.grid(row=0, column=col, sticky="nsew", padx=1, pady=1)
        
        # 데이터 행 생성
        cell_style = {
            'bg': 'white',
            'relief': 'solid',
            'borderwidth': 1,
            'font': ('Arial', 8),
            'anchor': 'center'
        }
        
        for row, item in enumerate(data, 1):
            # 폐업률 계산
            store_count = int(item.get('frcsCnt', 0))
            end_count = int(item.get('ctrtEndCnt', 0))
            cancel_count = int(item.get('ctrtCncltnCnt', 0))
            total_closure = end_count + cancel_count
            
            if store_count > 0:
                closure_rate = round((total_closure / store_count) * 100, 1)
            else:
                closure_rate = 0.0
            
            # 순위
            rank_label = tk.Label(scrollable_frame, text=str(row), width=header_widths[0], **cell_style)
            rank_label.grid(row=row, column=0, sticky="nsew", padx=1, pady=1)
            
            # 브랜드명
            brand_label = tk.Label(scrollable_frame, text=item.get('brandNm', ''), 
                                 width=header_widths[1], **cell_style)
            brand_label.grid(row=row, column=1, sticky="nsew", padx=1, pady=1)
            
            # 운영회사
            company_label = tk.Label(scrollable_frame, text=item.get('corpNm', ''), 
                                   width=header_widths[2], **cell_style)
            company_label.grid(row=row, column=2, sticky="nsew", padx=1, pady=1)
            
            # 가맹점 수
            store_count_formatted = self.format_number(item.get('frcsCnt', '0'))
            store_label = tk.Label(scrollable_frame, text=f"{store_count_formatted}개", 
                                 width=header_widths[3], **cell_style)
            store_label.grid(row=row, column=3, sticky="nsew", padx=1, pady=1)
            
            # 계약종료
            end_count_formatted = self.format_number(item.get('ctrtEndCnt', '0'))
            end_label = tk.Label(scrollable_frame, text=end_count_formatted, 
                               width=header_widths[4], **cell_style)
            end_label.grid(row=row, column=4, sticky="nsew", padx=1, pady=1)
            
            # 계약해지
            cancel_count_formatted = self.format_number(item.get('ctrtCncltnCnt', '0'))
            cancel_label = tk.Label(scrollable_frame, text=cancel_count_formatted, 
                                  width=header_widths[5], **cell_style)
            cancel_label.grid(row=row, column=5, sticky="nsew", padx=1, pady=1)
            
            # 신규등록
            new_count_formatted = self.format_number(item.get('newFrcsRgsCnt', '0'))
            new_label = tk.Label(scrollable_frame, text=new_count_formatted, 
                               width=header_widths[6], **cell_style)
            new_label.grid(row=row, column=6, sticky="nsew", padx=1, pady=1)
            
            # 폐업률 (계산된 값)
            closure_rate_label = tk.Label(scrollable_frame, text=f"{closure_rate}%", 
                                        width=header_widths[7], **cell_style)
            closure_rate_label.grid(row=row, column=7, sticky="nsew", padx=1, pady=1)
            
            # 폐업률에 따른 색상 적용
            if closure_rate >= 20:
                closure_rate_label.config(fg='red', font=('Arial', 8, 'bold'))
            elif closure_rate >= 10:
                closure_rate_label.config(fg='orange', font=('Arial', 8, 'bold'))
            else:
                closure_rate_label.config(fg='green', font=('Arial', 8, 'bold'))
            
            # 업종대분류
            industry_label = tk.Label(scrollable_frame, text=item.get('indutyLclasNm', ''), 
                                    width=header_widths[8], **cell_style)
            industry_label.grid(row=row, column=8, sticky="nsew", padx=1, pady=1)
            
            # 행 색상 교대로 적용 (AgeClosurePage와 동일)
            if row % 2 == 0:
                for col in range(len(headers)):
                    scrollable_frame.grid_slaves(row=row, column=col)[0].config(bg='#F8F9FA')
        
        # 위젯 배치
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar_v.grid(row=0, column=1, sticky="ns")
        scrollbar_h.grid(row=1, column=0, sticky="ew")
        
        # 그리드 가중치 설정
        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)
    
    def format_number(self, value):
        """숫자 값에 천 단위 콤마 추가"""
        if value is None or value == '':
            return '0'
        
        try:
            # 문자열에서 숫자만 추출
            num_str = str(value).replace(',', '').strip()
            if num_str.replace('.', '').replace('-', '').isdigit():
                if '.' in num_str:
                    return f"{float(num_str):,.0f}"
                else:
                    return f"{int(num_str):,}"
        except (ValueError, AttributeError):
            pass
        
        return str(value)


########################################################################################################################################


class RegionalClosurePage(DetailPage):
    """지역별 폐업사유 분석 페이지 """

    def __init__(self, parent, controller):
        DetailPage.__init__(self, parent, controller,
            "지역별 폐업사유 통계", "과세유형, 지역/업태, 폐업사유별 현황")

        # CSV 파일 경로 (절대경로로 처리)
        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.csv_path = os.path.join(
            current_dir,
            "9.8.13_폐업자_현황_Ⅳ__폐업사유__지역__업태_2005_20250525201244.csv"
        )

        # 컨트롤 프레임
        self.control_frame = tk.Frame(self.content_frame, bg="white")
        self.control_frame.pack(fill=tk.X, pady=10)

        tk.Label(self.control_frame, text="조회 년도:", bg="white").pack(side=tk.LEFT, padx=5)
        self.year_var = tk.StringVar(value="2023")
        self.year_combo = ttk.Combobox(
            self.control_frame, textvariable=self.year_var,
            values=[str(y) for y in range(2019, 2024)], width=8
        )
        self.year_combo.pack(side=tk.LEFT, padx=5)

        self.search_button = tk.Button(
            self.control_frame, text="데이터 불러오기",
            command=self.start_loading, bg="#4CAF50"
        )
        self.search_button.pack(side=tk.LEFT, padx=10)

        self.home_button = tk.Button(
            self.control_frame, text="메인으로 돌아가기",
            command=lambda: controller.show_frame(MainPage),
            bg="#2196F3", font=("Arial", 10, "bold")
        )
        self.home_button.pack(side=tk.LEFT, padx=5)

        # 테이블 프레임
        self.table_frame = tk.Frame(self.content_frame, bg="white")
        self.table_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 정보 표시 프레임
        self.info_frame = tk.Frame(self.content_frame, bg="white")
        self.info_frame.pack(fill=tk.X, pady=5)

        self.show_initial_message()

    def show_initial_message(self, text=None):
        """초기 메시지 표시"""
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        for widget in self.info_frame.winfo_children():
            widget.destroy()
            
        initial_text = text or "년도를 선택하고 '데이터 불러오기' 버튼을 클릭하세요.\n\n※ 데이터 로딩에 시간이 소요될 수 있습니다."
        initial_label = tk.Label(self.table_frame,
                                text=initial_text,
                                bg="white",
                                font=("Arial", 12),
                                fg="gray")
        initial_label.pack(pady=80)

    def start_loading(self):
        """별도 스레드에서 데이터 로딩 시작"""
        self.search_button.config(state=tk.DISABLED)
        self._show_loading_message()
        threading.Thread(target=self.load_data, daemon=True).start()

    def _show_loading_message(self):
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        for widget in self.info_frame.winfo_children():
            widget.destroy()
            
        self.loading_label = tk.Label(self.table_frame, 
                                     text="데이터를 불러오는 중...", 
                                     fg="blue", bg="white", font=("Arial", 14))
        self.loading_label.pack(pady=50)

    def load_data(self):
        try:
            selected_year = self.year_var.get()
            raw_data = self.load_csv_data()
            table_data = self.refine_data(raw_data, selected_year)
            self.table_frame.after(0, self.update_table, table_data, selected_year)
        except Exception as e:
            self.table_frame.after(0, self.show_error, str(e))
        finally:
            self.table_frame.after(0, self.search_button.config, {'state': tk.NORMAL})

    def load_csv_data(self):
        """CSV 파일 로드"""
        try:
            return pd.read_csv(self.csv_path, encoding='utf-8-sig')
        except Exception:
            return pd.read_csv(self.csv_path, encoding='cp949')

    def clean_code_name(self, value):
        """코드+이름 형태에서 이름만 추출"""
        if pd.isna(value) or value == '':
            return ''
        
        value_str = str(value).strip()
        
        # 패턴 1: "15133SEJ01 법인사업자" -> "법인사업자"
        # 패턴 2: "B03 서울" -> "서울"  
        # 패턴 3: "소계" -> "소계"
        match = re.search(r'^[A-Z0-9]+\s+(.+)', value_str)
        if match:
            return match.group(1).strip()
        
        return value_str

    def clean_year(self, value):
        """Y2023 2023 -> 2023"""
        if pd.isna(value) or value == '':
            return ''
        
        value_str = str(value).strip()
        # "Y2023 2023"에서 마지막 4자리 숫자 추출
        match = re.search(r'(\d{4})$', value_str)
        if match:
            return match.group(1)
        
        return value_str

    def refine_data(self, raw_data, year):
        """데이터 정제 및 코드 제거 - 지역별/업태별 구분하여 정렬"""
        # 해당 연도 데이터만 필터링
        filtered = raw_data[raw_data['TIME 시점'].astype(str).str.contains(f'Y{year}')]
        
        refined = []
        for _, row in filtered.iterrows():
            # 각 컬럼에서 코드 제거하고 깔끔하게 정리
            refined_row = {
                '과세유형': self.clean_code_name(row['15133SEJ 과세유형별(1)']),
                '지역/업태 구분': self.clean_code_name(row['13301B 시도․업태별(1)']),
                '세부지역/업태': self.clean_code_name(row['13301B 시도․업태별(2)']),
                '시점': self.clean_year(row['TIME 시점']),
                '총계': self.format_number(row.get('16133T2008_0245 총계', '-')),
                '사업부진': self.format_number(row.get('16133T2008_0418 사업부진', '-')),
                '행정처분': self.format_number(row.get('16133T2008_0419 행정처분', '-')),
                '계절사업': self.format_number(row.get('16133T2008_0420 계절사업', '-')),
                '법인전환': self.format_number(row.get('16133T2008_0421 법인전환', '-')),
                '면세포기·적용': self.format_number(row.get('16133T2008_0422 면세포기·적용', '-')),
                '양도·양수': self.format_number(row.get('T001 양도·양수', '-')),
                '해산·합병': self.format_number(row.get('16133T2008_0425 해산·합병', '-')),
                '기타': self.format_number(row.get('16133ABA7 기타', '-')),
            }
            refined.append(refined_row)
        
        # 지역별/업태별 구분을 위한 정렬
        # 1. 과세유형 -> 2. 지역/업태 구분 -> 3. 세부지역/업태 순으로 정렬
        refined.sort(key=lambda x: (x['과세유형'], x['지역/업태 구분'], x['세부지역/업태']))
        
        return refined

    def update_table(self, data, year):
        """테이블 업데이트"""
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        for widget in self.info_frame.winfo_children():
            widget.destroy()
            
        if not data:
            self.show_initial_message(f"{year}년 데이터가 없습니다")
            return
            
        self.create_grouped_table(data)
        
        # 정보 표시
        info_text = f"{year}년 데이터 총 {len(data):,}건 로드 완료"
        info_label = tk.Label(self.info_frame, 
                             text=info_text, 
                             bg="white", 
                             fg="blue", 
                             font=("Arial", 10))
        info_label.pack(side=tk.LEFT)
        
        # 추가 정보
        detail_info = tk.Label(self.info_frame,
                              text="※ 지역별/업태별 그룹으로 구분하여 표시",
                              bg="white",
                              fg="gray",
                              font=("Arial", 9))
        detail_info.pack(side=tk.RIGHT)

    def create_grouped_table(self, data):
        """지역별/업태별 그룹으로 구분된 테이블 생성"""
        # 스크롤 가능한 테이블 컨테이너
        canvas = tk.Canvas(self.table_frame, bg="white")
        scrollbar_v = ttk.Scrollbar(self.table_frame, orient="vertical", command=canvas.yview)
        scrollbar_h = ttk.Scrollbar(self.table_frame, orient="horizontal", command=canvas.xview)
        scrollable_frame = tk.Frame(canvas, bg="white")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)

        # 헤더 생성
        headers = list(data[0].keys()) if data else []
        header_widths = [15, 12, 18, 8, 12, 10, 10, 10, 10, 12, 10, 10, 10]
        
        for col, (header, width) in enumerate(zip(headers, header_widths)):
            header_label = tk.Label(scrollable_frame, 
                                   text=header, 
                                   bg="#4A5C90", 
                                   fg="white",
                                   relief="solid",
                                   borderwidth=1, 
                                   font=('Arial', 9, 'bold'),
                                   width=width,
                                   anchor='center')
            header_label.grid(row=0, column=col, sticky="nsew", padx=1, pady=1)
        
        # 데이터를 그룹별로 분류
        grouped_data = self.group_data_by_category(data)
        
        row_idx = 1
        
        # 그룹별로 데이터 표시
        for group_key, group_data in grouped_data.items():
            # 그룹 헤더 추가 (지역별 또는 업태별 구분)
            group_header_color = "#E8F4FD" if "지역별" in group_key else "#FFF2E8"
            
            # 그룹명 표시 (전체 컬럼에 걸쳐서)
            group_label = tk.Label(scrollable_frame, 
                                 text=f"▼ {group_key}",
                                 bg=group_header_color,
                                 relief="solid",
                                 borderwidth=1,
                                 font=('Arial', 9, 'bold'),
                                 anchor='w',
                                 padx=10)
            group_label.grid(row=row_idx, column=0, columnspan=len(headers), 
                           sticky="nsew", padx=1, pady=2)
            row_idx += 1
            
            # 그룹 내 데이터 표시
            for item_idx, item in enumerate(group_data):
                # 그룹 내에서 교대로 색상 적용
                if item_idx % 2 == 0:
                    bg_color = '#F8F9FA'
                else:
                    bg_color = 'white'
                
                for col, (key, width) in enumerate(zip(headers, header_widths)):
                    cell_value = item.get(key, '-')
                    
                    # 첫 번째 컬럼(과세유형)은 그룹 내에서 중복 제거
                    if col == 0 and item_idx > 0 and item.get(key) == group_data[item_idx-1].get(key):
                        cell_value = ""
                    
                    cell_label = tk.Label(scrollable_frame, 
                                         text=cell_value, 
                                         bg=bg_color,
                                         relief="solid", 
                                         borderwidth=1, 
                                         font=('Arial', 8),
                                         width=width,
                                         anchor='center')
                    cell_label.grid(row=row_idx, column=col, sticky="nsew", padx=1, pady=1)
                
                row_idx += 1
            
            # 그룹 간 구분선 추가
            separator = tk.Frame(scrollable_frame, height=3, bg="#DDDDDD")
            separator.grid(row=row_idx, column=0, columnspan=len(headers), 
                          sticky="ew", padx=1, pady=2)
            row_idx += 1

        # 위젯 배치
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar_v.grid(row=0, column=1, sticky="ns")
        scrollbar_h.grid(row=1, column=0, sticky="ew")
        
        # 그리드 가중치 설정
        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)

    def group_data_by_category(self, data):
        """데이터를 지역별/업태별로 그룹화"""
        grouped = {}
        
        for item in data:
            tax_type = item.get('과세유형', '')
            category = item.get('지역/업태 구분', '')
            
            # 그룹 키 생성
            group_key = f"{tax_type} - {category}"
            
            if group_key not in grouped:
                grouped[group_key] = []
            
            grouped[group_key].append(item)
        
        return grouped

    def show_error(self, message):
        """오류 메시지 표시"""
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        for widget in self.info_frame.winfo_children():
            widget.destroy()
            
        error_label = tk.Label(self.table_frame,
                              text=f"오류 발생: {message}",
                              fg="red", bg="white", font=("Arial", 12))
        error_label.pack(pady=20)

    def format_number(self, value):
        """숫자 포맷팅 (천 단위 콤마 추가)"""
        if pd.isna(value) or value in ['', '-']:
            return '-'
        try:
            # 숫자 변환 시도
            if isinstance(value, (int, float)):
                return f"{int(value):,}" if value != 0 else '-'
            
            # 문자열에서 숫자만 추출
            num_str = str(value).replace(',', '').strip()
            if num_str.replace('.', '').replace('-', '').isdigit():
                num_value = int(float(num_str))
                return f"{num_value:,}" if num_value != 0 else '-'
        except (ValueError, TypeError):
            pass
        
        return str(value) if value else '-'



########################################################################################################################################

class BusinessSurvivalPage(DetailPage):
    """사업 유지율 통계 페이지 - 선택 연도 필터링 및 중복 총계 제거"""

    def __init__(self, parent, controller):
        DetailPage.__init__(self, parent, controller,
                          "사업 유지율 통계", "과세유형, 사업존속연수, 지역, 업태별 가동사업자 현황")

        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.csv_path = os.path.join(
            current_dir,
            "9.8.7_가동사업자_현황Ⅱ__사업존속연수_지역_업태_2007_20250526210827.csv"
        )

        self.control_frame = tk.Frame(self.content_frame, bg="white")
        self.control_frame.pack(fill=tk.X, pady=10)

        tk.Label(self.control_frame, text="조회 년도:", bg="white").pack(side=tk.LEFT, padx=5)
        self.year_var = tk.StringVar(value="2023")
        self.year_combo = ttk.Combobox(
            self.control_frame, textvariable=self.year_var,
            values=[str(y) for y in range(2019, 2024)], width=8
        )
        self.year_combo.pack(side=tk.LEFT, padx=5)
        
        self.search_button = tk.Button(
            self.control_frame, text="데이터 불러오기",
            command=self.start_loading, bg="#4CAF50", fg="black"
        )
        self.search_button.pack(side=tk.LEFT, padx=10)

        self.home_button = tk.Button(
            self.control_frame, text="메인으로 돌아가기",
            command=lambda: controller.show_frame(MainPage),
            bg="#2196F3", font=("Arial", 10, "bold")
        )
        self.home_button.pack(side=tk.LEFT, padx=5)

        self.table_frame = tk.Frame(self.content_frame, bg="white")
        self.table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.info_frame = tk.Frame(self.content_frame, bg="white")
        self.info_frame.pack(fill=tk.X, pady=5)

        self.tree = None
        self.show_initial_message()

    def show_initial_message(self, text=None):
        for widget in self.table_frame.winfo_children(): widget.destroy()
        for widget in self.info_frame.winfo_children(): widget.destroy()
        initial_text = text or "년도를 선택하고 '데이터 불러오기' 버튼을 클릭하세요.\n\n※ 데이터 로딩에 시간이 소요될 수 있습니다."
        tk.Label(self.table_frame, text=initial_text, bg="white", font=("Arial", 12), fg="gray").pack(pady=80)

    def start_loading(self):
        self.search_button.config(state=tk.DISABLED)
        self._show_loading_message()
        threading.Thread(target=self.load_data, daemon=True).start()

    def _show_loading_message(self):
        for widget in self.table_frame.winfo_children(): widget.destroy()
        for widget in self.info_frame.winfo_children(): widget.destroy()
        tk.Label(self.table_frame, text="데이터를 불러오는 중...", fg="blue", bg="white", font=("Arial", 14)).pack(pady=50)

    def load_data(self):
        try:
            raw_data_df = self.load_csv_data()
            if raw_data_df is None:
                self.table_frame.after(0, self.show_error, "CSV 파일을 로드하지 못했습니다.")
                return

            # refine_data 호출 시 선택된 연도 전달 (또는 내부에서 self.year_var.get() 사용)
            table_data, years_for_display, periods_for_display = self.refine_data(raw_data_df)
            
            if table_data is None: return
            # years_for_display는 이제 선택된 단일 연도 리스트가 됨
            self.table_frame.after(0, self.update_table, table_data, years_for_display, periods_for_display)
        except Exception as e:
            self.table_frame.after(0, self.show_error, f"데이터 처리 중 예외 발생: {e}")
        finally:
            if hasattr(self, 'search_button') and self.search_button.winfo_exists():
                 self.table_frame.after(0, lambda: self.search_button.config(state=tk.NORMAL))

    def load_csv_data(self):
        try:
            df = pd.read_csv(self.csv_path, encoding='utf-8-sig', header=[0, 1])
            return df
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(self.csv_path, encoding='cp949', header=[0, 1])
                return df
            except Exception: return None
        except Exception: return None

    def clean_identifier_value(self, value):
        if pd.isna(value) or str(value).strip() == '': return ''
        value_str = str(value).strip()
        match_code = re.search(r'^[A-Z0-9]+\s*(.+)', value_str)
        cleaned_val = match_code.group(1).strip() if match_code else value_str
        cleaned_val = re.sub(r'\s*\(\d\)$', '', cleaned_val).strip()
        return cleaned_val
    
    def format_number(self, value):
        if pd.isna(value) or str(value).strip() == '-': return '-'
        try:
            num_str = str(value).replace(',', '').strip()
            if num_str.replace('.', '', 1).replace('-', '', 1).isdigit():
                num_val = float(num_str)
                return f"{int(num_val):,}" if num_val == int(num_val) else f"{num_val:,.0f}"
        except (ValueError, TypeError): pass
        return str(value)

    def refine_data(self, raw_data_df):
        selected_year_str = self.year_var.get() # 사용자가 선택한 연도 (예: "2023")
        refined_list = []
        try:
            id_cols_map = {}
            if len(raw_data_df.columns) >= 3:
                id_cols_map[raw_data_df.columns[0]] = '과세유형'
                id_cols_map[raw_data_df.columns[1]] = '구분1'
                id_cols_map[raw_data_df.columns[2]] = '구분2'
            else:
                self.table_frame.after(0, self.show_error, "CSV 식별자 컬럼이 3개 미만입니다.")
                return None, None, None
            
            # 데이터 컬럼 중에서 선택된 연도에 해당하는 컬럼들만 필터링
            # 멀티인덱스 컬럼의 첫 번째 레벨이 selected_year_str과 일치하는 것을 찾음
            data_cols_for_selected_year = [
                col for col in raw_data_df.columns 
                if col not in id_cols_map.keys() and 
                   isinstance(col, tuple) and len(col) > 0 and col[0] == selected_year_str
            ]

            if not data_cols_for_selected_year:
                self.table_frame.after(0, self.show_error, f"{selected_year_str}년 데이터를 찾을 수 없습니다.")
                return None, None, None

            # 이제 years_for_display는 선택된 연도 하나만 가짐
            years_for_display = [selected_year_str]

            # 선택된 연도의 사업존속기간 추출
            ordered_survival_periods = ['총계', '6월미만', '6월이상', '1년이상', '2년이상', '3년이상', '5년 이상', '10년 이상', '20년 이상', '30년 이상']
            
            # data_cols_for_selected_year에서 기간 정보(멀티인덱스의 두 번째 레벨)를 가져옴
            available_periods_in_header_raw = list(set(col[1] for col in data_cols_for_selected_year if len(col) > 1))
            
            # 중복된 "총계"를 제거하고 ordered_survival_periods 순서에 맞게 정렬
            # "총계"는 하나만 유지 (보통 처음에 나옴)
            survival_periods_to_display = []
            has_total = False
            for period in ordered_survival_periods:
                if period in available_periods_in_header_raw:
                    if period == '총계':
                        if not has_total:
                            survival_periods_to_display.append(period)
                            has_total = True
                    else:
                        survival_periods_to_display.append(period)
            
            # 만약 ordered_survival_periods에 없는 기간이 available_periods_in_header_raw에 있다면 추가 (순서 유지 어려움)
            # 여기서는 ordered_survival_periods에 있는 것만 고려
            
            if not survival_periods_to_display:
                 self.table_frame.after(0, self.show_error, f"{selected_year_str}년의 사업존속기간 정보를 찾을 수 없습니다.")
                 return None, years_for_display, []

            # 데이터 정제 (선택된 연도에 대해서만)
            for _, row in raw_data_df.iterrows():
                refined_row = {}
                for actual_col_key, display_name in id_cols_map.items():
                    refined_row[display_name] = self.clean_identifier_value(row.get(actual_col_key, ''))
                
                # 선택된 연도의 사업존속기간 데이터만 refined_row에 추가
                for period in survival_periods_to_display:
                    col_key_for_value = (selected_year_str, period) # (예: "2023", "총계")
                    value = row.get(col_key_for_value, '-')
                    refined_row[f'{selected_year_str}_{period}'] = self.format_number(value)
                refined_list.append(refined_row)
            
            refined_list.sort(key=lambda x: (x.get('과세유형',''), x.get('구분1',''), x.get('구분2','')))
            return refined_list, years_for_display, survival_periods_to_display

        except Exception as e_gen:
            import traceback; traceback.print_exc()
            self.table_frame.after(0, self.show_error, f"데이터 정제 중 오류: {e_gen}")
            return None, None, None

    def update_table(self, data, years_for_display, periods_for_display):
        for widget in self.table_frame.winfo_children(): widget.destroy()
        for widget in self.info_frame.winfo_children(): widget.destroy()
            
        if not data or not periods_for_display or not years_for_display:
            self.show_initial_message("데이터 또는 헤더 정보가 부족하여 테이블을 생성할 수 없습니다.")
            return
            
        self.create_treeview_for_survival(data, years_for_display, periods_for_display)
        
        # 정보 라인 텍스트 (years_for_display는 이제 선택된 단일 연도 리스트)
        selected_display_year = years_for_display[0] if years_for_display else self.year_var.get()
        info_text = f"선택한 연도: {selected_display_year}년, 총 {len(data):,}건 데이터 로드 완료"
        tk.Label(self.info_frame, text=info_text, bg="white", fg="blue", font=("Arial", 10)).pack(side=tk.LEFT)

    def create_treeview_for_survival(self, data, years_to_display, survival_periods):
        # years_to_display는 이제 단일 선택된 연도만 담고 있는 리스트 (예: ["2023"])
        selected_year = years_to_display[0]

        column_ids = ['col_tax_type', 'col_cat1', 'col_cat2']
        column_display_names = ['과세유형', '구분1', '구분2']
        column_widths = [150, 120, 180]

        # 데이터 컬럼 (Treeview 헤더는 사업존속기간만 표시)
        for period in survival_periods: # survival_periods는 이미 중복 "총계"가 제거된 상태여야 함
            # Treeview 컬럼 ID는 고유해야 하므로 연도와 기간을 조합 (실제로는 선택된 연도만 사용)
            col_id = f"col_{selected_year}_{period.replace(' ', '_').replace('.', '')}"
            column_ids.append(col_id)
            column_display_names.append(period) # 헤더는 기간만 표시
            column_widths.append(90)

        self.tree = ttk.Treeview(self.table_frame, columns=column_ids, show="headings")
        
        vsb = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.table_frame, orient="horizontal", command=self.tree.xview)
        
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)
        
        for i, col_id in enumerate(column_ids):
            self.tree.heading(col_id, text=column_display_names[i], anchor='center')
            self.tree.column(col_id, width=column_widths[i], anchor='center' if i >=3 else 'w', stretch=tk.NO)

        self.tree.tag_configure('group_header_main', font=('Arial', 9, 'bold'), background='#E0E0E0')
        self.tree.tag_configure('oddrow', background='#F8F9FA')
        self.tree.tag_configure('evenrow', background='white')

        current_group_main_key = None; current_group_sub_key = None
        for item_idx, item_data in enumerate(data):
            group_main_key_val = item_data.get('과세유형', ''); group_sub_key_val = item_data.get('구분1', '')
            
            if (group_main_key_val, group_sub_key_val) != (current_group_main_key, current_group_sub_key):
                current_group_main_key = group_main_key_val; current_group_sub_key = group_sub_key_val
                group_header_text = f"{group_main_key_val} ({group_sub_key_val})" if group_sub_key_val else f"{group_main_key_val}"
                group_values = [group_header_text] + [''] * (len(column_ids) - 1)
                self.tree.insert('', 'end', values=group_values, open=True, tags=('group_header_main',))

            row_values = []
            row_values.append(item_data.get('과세유형', '-'))
            row_values.append(item_data.get('구분1', '-'))
            row_values.append(item_data.get('구분2', '-'))
            
            # 데이터 컬럼 값 채우기 (선택된 연도에 대해서만)
            for period in survival_periods:
                # refined_data에서 이미 {selected_year}_{period} 형태로 키가 저장되어 있음
                row_values.append(item_data.get(f'{selected_year}_{period}', '-'))
            
            tag = 'oddrow' if item_idx % 2 == 0 else 'evenrow'
            self.tree.insert('', 'end', values=row_values, tags=(tag,))
            
    def show_error(self, message):
        for widget in self.table_frame.winfo_children(): widget.destroy()
        for widget in self.info_frame.winfo_children(): widget.destroy()
        error_label = tk.Label(self.table_frame, text=f"오류 발생: {message}", fg="red", bg="white", font=("Arial", 12),
                              wraplength=self.table_frame.winfo_width() - 20 if self.table_frame.winfo_width() > 20 else 300)
        error_label.pack(pady=20, padx=10)




########################################################################################################################################
class NewBusinessPage(DetailPage):
    """신규 사업자 현황 페이지 - 계층적 그룹 구분감 강화"""

    def __init__(self, parent, controller):
        DetailPage.__init__(self, parent, controller,
                          "신규 사업자 현황", "연령, 성별, 지역별 신규 사업자 수")

        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.csv_path = os.path.join(
            current_dir,
            "9.8.19_신규_사업자_현황_Ⅲ__연령__성__지역_2013_20250527163606.csv"
        )

        self.control_frame = tk.Frame(self.content_frame, bg="white")
        self.control_frame.pack(fill=tk.X, pady=10)

        tk.Label(self.control_frame, text="조회 년도:", bg="white").pack(side=tk.LEFT, padx=5)
        self.year_var = tk.StringVar(value="2023")
        self.year_combo = ttk.Combobox(
            self.control_frame, textvariable=self.year_var,
            values=[str(y) for y in range(2019, 2024)], width=8 # CSV 데이터 연도 범위 확인 필요
        )
        self.year_combo.pack(side=tk.LEFT, padx=5)
        
        self.search_button = tk.Button(
            self.control_frame, text="데이터 불러오기",
            command=self.start_loading, bg="#4CAF50", fg="black"
        )
        self.search_button.pack(side=tk.LEFT, padx=10)

        self.home_button = tk.Button(
            self.control_frame, text="메인으로 돌아가기",
            command=lambda: controller.show_frame(MainPage),
            bg="#2196F3", font=("Arial", 10, "bold")
        )
        self.home_button.pack(side=tk.LEFT, padx=5)

        self.table_frame = tk.Frame(self.content_frame, bg="white")
        self.table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.info_frame = tk.Frame(self.content_frame, bg="white")
        self.info_frame.pack(fill=tk.X, pady=5)

        self.tree = None
        self.show_initial_message()

    def show_initial_message(self, text=None):
        for widget in self.table_frame.winfo_children(): widget.destroy()
        for widget in self.info_frame.winfo_children(): widget.destroy()
        initial_text = text or "년도를 선택하고 '데이터 불러오기' 버튼을 클릭하세요.\n\n※ 데이터 로딩에 시간이 소요될 수 있습니다."
        tk.Label(self.table_frame, text=initial_text, bg="white", font=("Arial", 12), fg="gray").pack(pady=80)

    def start_loading(self):
        self.search_button.config(state=tk.DISABLED)
        self._show_loading_message()
        threading.Thread(target=self.load_data, daemon=True).start()

    def _show_loading_message(self):
        for widget in self.table_frame.winfo_children(): widget.destroy()
        for widget in self.info_frame.winfo_children(): widget.destroy()
        tk.Label(self.table_frame, text="데이터를 불러오는 중...", fg="blue", bg="white", font=("Arial", 14)).pack(pady=50)

    def load_data(self):
        try:
            raw_data_df = self.load_csv_data()
            if raw_data_df is None:
                self.table_frame.after(0, self.show_error, "CSV 파일을 로드하지 못했습니다.")
                return

            table_data, age_groups_for_display = self.refine_new_business_data(raw_data_df)
            
            if table_data is None: return
            self.table_frame.after(0, self.update_table, table_data, age_groups_for_display)
        except Exception as e:
            self.table_frame.after(0, self.show_error, f"데이터 처리 중 예외 발생: {e}")
        finally:
            if hasattr(self, 'search_button') and self.search_button.winfo_exists():
                 self.table_frame.after(0, lambda: self.search_button.config(state=tk.NORMAL))

    def load_csv_data(self):
        # (이전 답변과 동일)
        try:
            df = pd.read_csv(self.csv_path, encoding='utf-8-sig', header=[0, 1])
            return df
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(self.csv_path, encoding='cp949', header=[0, 1])
                return df
            except Exception: return None
        except Exception: return None


    def clean_identifier_value(self, value):
        # (이전 답변과 동일)
        if pd.isna(value) or str(value).strip() == '': return ''
        value_str = str(value).strip()
        match_code = re.search(r'^[A-Z0-9]+\s*(.+)', value_str)
        cleaned_val = match_code.group(1).strip() if match_code else value_str
        cleaned_val = re.sub(r'\s*\(\d\)$', '', cleaned_val).strip()
        return cleaned_val
    
    def format_number(self, value):
        # (이전 답변과 동일)
        if pd.isna(value) or str(value).strip() == '-': return '-'
        try:
            num_str = str(value).replace(',', '').strip()
            if num_str.replace('.', '', 1).replace('-', '', 1).isdigit():
                num_val = float(num_str)
                return f"{int(num_val):,}" if num_val == int(num_val) else f"{num_val:,.0f}"
        except (ValueError, TypeError): pass
        return str(value)

    def refine_new_business_data(self, raw_data_df):
        # (이전 답변과 동일 - 선택 연도 필터링 로직 포함)
        selected_year_str = self.year_var.get()
        refined_list = []
        try:
            id_cols_map = {}
            if len(raw_data_df.columns) >= 3:
                id_cols_map[raw_data_df.columns[0]] = '사업자별'
                id_cols_map[raw_data_df.columns[1]] = '성별'
                id_cols_map[raw_data_df.columns[2]] = '지역별'
            else:
                self.table_frame.after(0, self.show_error, "신규사업자 CSV 식별자 컬럼이 3개 미만입니다.")
                return None, None

            data_cols_for_selected_year = [
                col for col in raw_data_df.columns 
                if col not in id_cols_map.keys() and 
                   isinstance(col, tuple) and len(col) > 0 and col[0] == selected_year_str
            ]
            if not data_cols_for_selected_year:
                self.table_frame.after(0, self.show_error, f"{selected_year_str}년 데이터를 찾을 수 없습니다.")
                return None, None

            ordered_age_groups = ["합계", "30세 미만", "30세 이상", "40세 이상", "50세 이상", "60세 이상", "70세 이상"]
            available_age_groups_raw = list(set(col[1] for col in data_cols_for_selected_year if len(col) > 1))
            age_groups_to_display = [group for group in ordered_age_groups if group in available_age_groups_raw]
            if not age_groups_to_display and available_age_groups_raw:
                age_groups_to_display = available_age_groups_raw
            if not age_groups_to_display:
                 self.table_frame.after(0, self.show_error, f"{selected_year_str}년의 연령 그룹 정보를 찾을 수 없습니다.")
                 return None, None

            for _, row in raw_data_df.iterrows():
                refined_row = {}
                for actual_col_key, display_name in id_cols_map.items():
                    refined_row[display_name] = self.clean_identifier_value(row.get(actual_col_key, ''))
                for age_group in age_groups_to_display:
                    col_key_for_value = (selected_year_str, age_group)
                    value = row.get(col_key_for_value, '-')
                    refined_row[age_group.replace(' ', '_').replace('.', '')] = self.format_number(value)
                refined_list.append(refined_row)
            
            refined_list.sort(key=lambda x: (x.get('사업자별',''), x.get('성별',''), x.get('지역별','')))
            return refined_list, age_groups_to_display
        except Exception as e_gen:
            import traceback; traceback.print_exc()
            self.table_frame.after(0, self.show_error, f"신규사업자 데이터 정제 중 오류: {e_gen}")
            return None, None

    def update_table(self, data, age_groups_for_display):
        # (이전 답변과 동일)
        for widget in self.table_frame.winfo_children(): widget.destroy()
        for widget in self.info_frame.winfo_children(): widget.destroy()
            
        if not data or not age_groups_for_display:
            self.show_initial_message("데이터 또는 헤더 정보가 부족하여 테이블을 생성할 수 없습니다.")
            return
            
        self.create_new_business_treeview(data, age_groups_for_display) # 메소드명 일관성
        
        selected_year_combo = self.year_var.get()
        info_text = f"선택한 연도: {selected_year_combo}년, 총 {len(data):,}건 데이터 로드 완료"
        tk.Label(self.info_frame, text=info_text, bg="white", fg="blue", font=("Arial", 10)).pack(side=tk.LEFT)

    def create_new_business_treeview(self, data, age_groups):
        # 컬럼 정의 (이전과 동일)
        column_ids = ['col_biz_type', 'col_gender', 'col_region']
        column_display_names = ['사업자별', '성별', '지역별']
        column_widths = [120, 80, 120]

        for age_group in age_groups:
            col_id = f"col_age_{age_group.replace(' ', '_').replace('.', '')}"
            column_ids.append(col_id)
            column_display_names.append(age_group)
            column_widths.append(90)

        self.tree = ttk.Treeview(self.table_frame, columns=column_ids, show="headings")
        
        vsb = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)
        
        # 헤더 설정 (이전과 동일)
        for i, col_id in enumerate(column_ids):
            self.tree.heading(col_id, text=column_display_names[i], anchor='center')
            self.tree.column(col_id, width=column_widths[i], anchor='center' if i >=3 else 'w', stretch=tk.NO)

        # 스타일 태그 정의 (그룹별로 다른 배경색, 폰트)
        self.tree.tag_configure('group_biz_header', font=('Arial', 10, 'bold'), background='#D8E6F3') # 사업자별 그룹
        self.tree.tag_configure('group_gender_header', font=('Arial', 9, 'italic'), background='#E9D8F3') # 성별 그룹
        self.tree.tag_configure('data_row_odd', background='#F8F9FA')
        self.tree.tag_configure('data_row_even', background='white')

        current_biz_type = None
        current_gender = None
        biz_type_iid = ''  # 현재 사업자별 그룹의 아이템 ID
        gender_iid = ''    # 현재 성별 그룹의 아이템 ID
        row_counter = 0    # 데이터 행의 홀/짝 구분용

        for item_data in data:
            biz_type_val = item_data.get('사업자별', '')
            gender_val = item_data.get('성별', '')
            region_val = item_data.get('지역별', '')
            
            # "사업자별" 그룹 처리
            if biz_type_val != current_biz_type:
                current_biz_type = biz_type_val
                current_gender = None # 사업자 그룹이 바뀌면 성별 그룹도 리셋
                # 사업자별 그룹 헤더는 첫 번째 컬럼에만 텍스트 표시
                group_biz_values = [f"▼ {biz_type_val}"] + [''] * (len(column_ids) - 1)
                biz_type_iid = self.tree.insert('', 'end', values=group_biz_values, open=True, tags=('group_biz_header',))
                row_counter = 0 # 그룹 변경 시 행 카운터 리셋

            # "성별" 그룹 처리 (biz_type_iid가 있어야 그 자식으로 들어감)
            # 그리고 "합계" 성별은 별도 그룹 헤더로 만들지 않음 (캡처 이미지 스타일)
            if gender_val != current_gender and gender_val != "합계":
                current_gender = gender_val
                # 성별 그룹 헤더는 두 번째 컬럼에만 텍스트 표시
                group_gender_values = ['', f"└ {gender_val}"] + [''] * (len(column_ids) - 2)
                parent_for_gender = biz_type_iid if biz_type_iid else '' # 최상위 또는 사업자별 그룹의 자식
                gender_iid = self.tree.insert(parent_for_gender, 'end', values=group_gender_values, open=True, tags=('group_gender_header',))
                row_counter = 0 # 그룹 변경 시 행 카운터 리셋
            elif gender_val == "합계" and biz_type_val != "합계": # 사업자별 "합계"는 별도 그룹 헤더 안함
                gender_iid = biz_type_iid # 데이터가 사업자별 그룹 바로 아래로 가도록
            elif biz_type_val == "합계" and gender_val == "합계": # 전체 합계의 경우
                gender_iid = biz_type_iid

            # 실제 데이터 행 삽입
            row_values = []
            # 데이터 행에서는 식별자 컬럼을 비우거나, 원하는 대로 표시
            # 캡처 이미지 스타일: 그룹 헤더가 있으므로 데이터 행은 식별자 비우고 '지역별'부터 시작
            row_values.append('') # 사업자별 비움
            row_values.append('') # 성별 비움
            row_values.append(region_val) # 지역별은 표시
            
            for age_group in age_groups:
                row_values.append(item_data.get(age_group.replace(' ', '_').replace('.', ''), '-'))
            
            tag = 'data_row_odd' if row_counter % 2 == 0 else 'data_row_even'
            
            # 데이터 행의 부모 결정
            parent_for_data = gender_iid if gender_val != "합계" and biz_type_val != "합계" else biz_type_iid
            if not parent_for_data: parent_for_data = '' # 최상위로 (만약 biz_type_iid도 없다면)

            self.tree.insert(parent_for_data, 'end', values=row_values, tags=(tag,))
            row_counter += 1
            
    def show_error(self, message):
        # (이전 답변과 동일)
        for widget in self.table_frame.winfo_children(): widget.destroy()
        for widget in self.info_frame.winfo_children(): widget.destroy()
        error_label = tk.Label(self.table_frame, text=f"오류 발생: {message}", fg="red", bg="white", font=("Arial", 12),
                              wraplength=self.table_frame.winfo_width() - 20 if self.table_frame.winfo_width() > 20 else 300)
        error_label.pack(pady=20, padx=10)





#########################################################################################################################################################
class SearchStatsPage(DetailPage):
    """네이버 통합 검색어 트렌드 조회 페이지 - 그래프 깨짐 방지 강화"""
    def __init__(self, parent, controller):
        DetailPage.__init__(self, parent, controller,
                          "통합 검색어 트렌드", "네이버 데이터랩 API 활용")

        # 네이버 API 키 
        self.client_id = "M8tzIwynm9iGgAgerEQq" 
        self.client_secret = "vfs4UQIB65" 

        # --- UI 구성 (이전 답변과 대부분 동일) ---
        self.search_controls_frame = tk.Frame(self.content_frame, bg="white")
        self.search_controls_frame.pack(pady=10, fill=tk.X)

        tk.Label(self.search_controls_frame, text="검색어:", bg="white").pack(side=tk.LEFT, padx=5)
        self.keyword_entry = tk.Entry(self.search_controls_frame, width=25)
        self.keyword_entry.pack(side=tk.LEFT, padx=5)
        self.keyword_entry.insert(0, "카페")

        tk.Label(self.search_controls_frame, text="시작일(YYYY-MM-DD):", bg="white").pack(side=tk.LEFT, padx=5)
        self.start_date_entry = tk.Entry(self.search_controls_frame, width=12)
        self.start_date_entry.pack(side=tk.LEFT, padx=5)
        self.start_date_entry.insert(0, "2023-01-01")

        tk.Label(self.search_controls_frame, text="종료일(YYYY-MM-DD):", bg="white").pack(side=tk.LEFT, padx=5)
        self.end_date_entry = tk.Entry(self.search_controls_frame, width=12)
        self.end_date_entry.pack(side=tk.LEFT, padx=5)
        self.end_date_entry.insert(0, "2023-12-31")

        tk.Label(self.search_controls_frame, text="구간 단위:", bg="white").pack(side=tk.LEFT, padx=5)
        self.time_unit_var = tk.StringVar(value="date")
        self.time_unit_combo = ttk.Combobox(
            self.search_controls_frame, textvariable=self.time_unit_var,
            values=["date", "week", "month"], width=8, state="readonly"
        )
        self.time_unit_combo.pack(side=tk.LEFT, padx=5)

        self.fetch_button = tk.Button(self.search_controls_frame, text="검색량 조회",
                                   command=self.start_fetch_trend_data)
        self.fetch_button.pack(side=tk.LEFT, padx=10)
        
        self.home_button = tk.Button(self.search_controls_frame, text="메인으로 돌아가기",
                               command=lambda: controller.show_frame(MainPage),
                               bg="#2196F3", font=("Arial", 10, "bold"))
        self.home_button.pack(side=tk.LEFT, padx=5)

        self.result_text_area = scrolledtext.ScrolledText(self.content_frame, wrap=tk.WORD, height=8, width=80, font=("Arial", 9))
        self.result_text_area.pack(pady=5, fill=tk.X, expand=False)

        self.graph_frame = tk.Frame(self.content_frame, bg="lightgrey")
        self.graph_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        
        self.show_initial_graph_message()

    def show_initial_graph_message(self):
        for widget in self.graph_frame.winfo_children():
            widget.destroy()
        tk.Label(self.graph_frame, text="검색어를 입력하고 '검색량 조회' 버튼을 누르세요.",
                 font=("Arial", 12), bg="lightgrey", fg="gray").pack(expand=True)

    def start_fetch_trend_data(self):
        if self.client_secret == "YOUR_NAVER_CLIENT_SECRET" or not self.client_secret: # 시크릿 값 확인
            messagebox.showerror("API 키 오류", "네이버 API 클라이언트 시크릿을 코드에 정확히 입력해주세요.")
            return

        self.fetch_button.config(state=tk.DISABLED)
        self.result_text_area.delete('1.0', tk.END)
        self.result_text_area.insert(tk.END, "데이터를 가져오는 중...\n")
        
        for widget in self.graph_frame.winfo_children(): widget.destroy()
        tk.Label(self.graph_frame, text="그래프 로딩 중...", font=("Arial", 12), bg="lightgrey").pack(expand=True)

        threading.Thread(target=self.fetch_and_display_trend, daemon=True).start()

    def fetch_and_display_trend(self):
        keyword = self.keyword_entry.get().strip()
        start_date = self.start_date_entry.get().strip()
        end_date = self.end_date_entry.get().strip()
        time_unit = self.time_unit_var.get()

        if not all([keyword, start_date, end_date, time_unit]):
            self.result_text_area.after(0, lambda: self.result_text_area.insert(tk.END, "오류: 모든 필드를 입력해주세요.\n"))
            self.fetch_button.after(0, lambda: self.fetch_button.config(state=tk.NORMAL))
            self.graph_frame.after(0, self.show_initial_graph_message)
            return

        date_pattern = r"^\d{4}-\d{2}-\d{2}$"
        if not (re.match(date_pattern, start_date) and re.match(date_pattern, end_date)):
            self.result_text_area.after(0, lambda: self.result_text_area.insert(tk.END, "오류: 날짜 형식이 올바르지 않습니다 (YYYY-MM-DD).\n"))
            self.fetch_button.after(0, lambda: self.fetch_button.config(state=tk.NORMAL))
            self.graph_frame.after(0, self.show_initial_graph_message)
            return

        api_data = self._fetch_naver_datalab_api(keyword, start_date, end_date, time_unit)

        if api_data and 'results' in api_data and api_data['results'] and api_data['results'][0]['data']: # 데이터 존재 여부 확실히 체크
            self.result_text_area.after(0, lambda: self.result_text_area.insert(tk.END, f"'{keyword}' 검색 트렌드 결과 수신 완료.\n"))
            self.graph_frame.after(0, self._display_trend_graph, api_data, keyword)
        else:
            error_message = "데이터를 가져오지 못했거나 결과가 없습니다.\n"
            if api_data and 'errorMessage' in api_data:
                 error_message += f"API 오류: {api_data['errorMessage']} (코드: {api_data.get('errorCode','')})\n"
            elif not api_data:
                 error_message += "API 호출에 실패했습니다. 네트워크 연결 또는 API 키를 확인하세요.\n"

            self.result_text_area.after(0, lambda: self.result_text_area.insert(tk.END, error_message))
            self.graph_frame.after(0, self.show_initial_graph_message)

        self.fetch_button.after(0, lambda: self.fetch_button.config(state=tk.NORMAL))

    def _fetch_naver_datalab_api(self, keyword, start_date, end_date, time_unit):
        api_url = "https://openapi.naver.com/v1/datalab/search"
        headers = {
            "X-Naver-Client-Id": self.client_id,
            "X-Naver-Client-Secret": self.client_secret,
            "Content-Type": "application/json"
        }
        body = {
            "startDate": start_date, "endDate": end_date, "timeUnit": time_unit,
            "keywordGroups": [{"groupName": keyword, "keywords": [keyword]}]
        }
        try:
            response = requests.post(api_url, headers=headers, data=json.dumps(body), timeout=10)
            # 디버깅: 응답 상태 코드와 내용 일부 출력
            print(f"API 응답 상태 코드: {response.status_code}")
            try:
                print(f"API 응답 내용 (일부): {response.text[:200]}")
            except Exception:
                pass
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as errh:
            # 오류 메시지에 응답 내용 포함
            error_details = f"HTTP 오류: {errh}\n"
            try:
                error_details += f"응답 내용: {response.text}\n"
            except Exception:
                 error_details += "응답 내용을 확인할 수 없습니다.\n"
            self.result_text_area.after(0, lambda: self.result_text_area.insert(tk.END, error_details))
            try:
                return response.json() # API가 오류 정보도 JSON으로 반환할 수 있음
            except json.JSONDecodeError:
                return None # JSON 파싱도 실패하면 None 반환
        except requests.exceptions.RequestException as err: # 포괄적인 요청 관련 예외
            self.result_text_area.after(0, lambda: self.result_text_area.insert(tk.END, f"API 요청 중 오류 발생: {err}\n"))
            return None


    def _display_trend_graph(self, api_response_data, keyword):
        for widget in self.graph_frame.winfo_children():
            widget.destroy()

        try:
            result_data = api_response_data['results'][0]['data']
            if not result_data: # 데이터가 비어있는 경우
                tk.Label(self.graph_frame, text=f"'{keyword}'에 대한 트렌드 데이터가 없습니다.", bg="lightgrey", fg="orange").pack(expand=True)
                return

            periods = [item['period'] for item in result_data]
            ratios = [item['ratio'] for item in result_data]

            fig = plt.Figure(figsize=(10, 5), dpi=100)
            ax = fig.add_subplot(111)
            
            ax.plot(periods, ratios, marker='o', linestyle='-', color='dodgerblue', label=keyword)
            
            ax.set_title(f"'{keyword}' 검색량 트렌드", fontsize=14)
            ax.set_xlabel("기간", fontsize=10)
            ax.set_ylabel("상대적 검색량 비율", fontsize=10)
            
            # X축 레이블이 너무 많을 경우, 일부만 표시 (예: 10개 간격)
            if len(periods) > 20: # 레이블 개수 임계값 (조정 가능)
                tick_spacing = max(1, len(periods) // 10) # 약 10개의 레이블이 보이도록 간격 설정
                ax.set_xticks(ax.get_xticks()[::tick_spacing])

            ax.tick_params(axis='x', rotation=45, labelsize=8)
            ax.tick_params(axis='y', labelsize=8)
            ax.grid(True, linestyle='--', alpha=0.7)
            ax.legend()
            
            fig.tight_layout()

            canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
            canvas_widget = canvas.get_tk_widget()
            canvas_widget.pack(fill=tk.BOTH, expand=True)
            canvas.draw()
        except Exception as e:
            print(f"그래프 표시 중 오류: {e}")
            tk.Label(self.graph_frame, text=f"그래프를 표시하는 중 오류가 발생했습니다:\n{e}", 
                     bg="lightgrey", fg="red").pack(expand=True)

# --- DataVisualizationApp 클래스 및 if __name__ == "__main__": 블록 ---
# (첨부해주신 paste-2.txt 의 내용과 동일하게 여기에 위치해야 합니다.)
class DataVisualizationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("데이터 시각화 대시보드")
        self.root.geometry("1200x800")
        self.container = tk.Frame(root)
        self.container.pack(side="top", fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)
        self.frames = {}
        # SearchStatsPage 포함하여 모든 페이지 클래스 나열 (다른 페이지 클래스 정의 필요)
        # pages = (MainPage, AgeClosurePage, FranchiseClosurePage, RegionalClosurePage,
        #         BusinessSurvivalPage, NewBusinessPage, SearchStatsPage) 
        # 임시로 SearchStatsPage만 테스트 가능하도록 MainPage와 함께 구성
        pages = (MainPage, SearchStatsPage) # 다른 페이지는 주석 처리하거나 실제 정의 필요
        
        for F in pages:
            frame = F(self.container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        self.show_frame(MainPage) # 또는 SearchStatsPage로 바로 시작하려면 SearchStatsPage
    
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()
if __name__ == "__main__":
    root = tk.Tk()
    app = DataVisualizationApp(root)
    root.mainloop()
#########################################################################################################################################################ㅍ