import os
import re
import time
from tkinter import filedialog, messagebox
import tkinter as tk
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# 웹 크롤링 표준 라이브러리
from selenium import webdriver
from selenium.webdriver.common.by import By


class LinkExtractorApp:

    def __init__(self, root):
        self.root = root
        self.root.title("인터넷 바로가기(.url) 링크 & 가격 추출기")
        self.root.geometry("550x330")  # 라디오 버튼 배치를 위해 창 높이 최적화
        self.root.resizable(False, False)

        self.selected_path = tk.StringVar()
        self.exclude_words = tk.StringVar()
        
        # 옵션 상품 처리 방식을 선택하기 위한 변수 (기본값: 자동 모드)
        self.option_mode = tk.StringVar(value="auto")

        self.create_widgets()

    def create_widgets(self):
        # 1. 폴더 선택 섹션
        folder_frame = tk.LabelFrame(self.root, text=" 1. 대상 폴더 지정 ", padx=10, pady=10)
        folder_frame.pack(fill="x", padx=15, pady=5)

        self.entry_path = tk.Entry(
            folder_frame, textvariable=self.selected_path, width=45, state="readonly"
        )
        self.entry_path.pack(side="left", padx=(0, 5), expand=True, fill="x")

        btn_browse = tk.Button(
            folder_frame, text="폴더 찾기", command=self.browse_folder, width=10
        )
        btn_browse.pack(side="right")

        # 2. 필터링 섹션
        filter_frame = tk.LabelFrame(
            self.root, text=" 2. 제외할 하위 폴더 키워드 ", padx=10, pady=10
        )
        filter_frame.pack(fill="x", padx=15, pady=5)

        entry_filter = tk.Entry(
            filter_frame, textvariable=self.exclude_words, width=58
        )
        entry_filter.pack(fill="x", ipady=2)
        entry_filter.insert(0, "품절, 제외, 보류")

        # 3. 옵션 상품 수집 모드 선택 섹션 (5초 대기 텍스트 안내 반영)
        mode_frame = tk.LabelFrame(
            self.root, text=" 3. 옵션 상품 원가 수집 방식 설정 ", padx=10, pady=10
        )
        mode_frame.pack(fill="x", padx=15, pady=5)

        rad_auto = tk.Radiobutton(
            mode_frame, 
            text="자동 모드 (옵션 무시, 화면에 보이는 기본 최저가 즉시 수집)", 
            variable=self.option_mode, 
            value="auto"
        )
        rad_auto.pack(anchor="w")

        rad_manual = tk.Radiobutton(
            mode_frame, 
            text="반자동 모드 (상품별 5초 대기, 내가 직접 옵션을 클릭하여 수집)", 
            variable=self.option_mode, 
            value="manual"
        )
        rad_manual.pack(anchor="w")

        # 4. 실행 버튼 섹션
        btn_start = tk.Button(
            self.root,
            text="폐쇄몰 로그인 및 링크·단가 일괄 추출",
            font=("맑은 고딕", 11, "bold"),
            bg="#2ecc71",
            fg="white",
            command=self.start_extraction,
            height=2
        )
        btn_start.pack(fill="x", padx=15, pady=10)

    def browse_folder(self):
        folder = filedialog.askdirectory(title="추출할 최상위 폴더를 선택하세요")
        if folder:
            self.selected_path.set(os.path.abspath(folder))

    def extract_url_from_file(self, file_path):
        encodings = ["cp949", "utf-8", "utf-16", "utf-8-sig", "latin-1"]
        for encoding in encodings:
            try:
                with open(file_path, "r", encoding=encoding, errors="ignore") as f:
                    for line in f:
                        clean_line = line.strip()
                        if clean_line.upper().startswith("URL="):
                            parts = clean_line.split("=", 1)
                            if len(parts) > 1:
                                return parts[1].strip()
            except Exception:
                continue
        return None

    def start_extraction(self):
        base_dir = self.selected_path.get()
        current_mode = self.option_mode.get()  # 사용자가 라디오 버튼으로 고른 모드 값

        if not base_dir:
            messagebox.showwarning("경고", "먼저 링크를 추출할 폴더를 지정해 주세요.")
            return

        raw_input = self.exclude_words.get()
        exclude_list = [w.strip() for w in raw_input.split(",") if w.strip()]

        # ---- [웹 브라우저 제어 및 반자동 로그인 대기] ----
        messagebox.showinfo(
            "로그인 안내",
            "확인을 누르면 크롬 브라우저가 열립니다.\n\n"
            "사이트 가입 아이디로 로그인을 완전히 완료하신 후에,\n"
            "다시 이 프로그램 창으로 돌아와 [확인]을 눌러주세요.",
        )

        try:
            options = webdriver.ChromeOptions()
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)
            
            driver = webdriver.Chrome(options=options)
            # ★ 요청하신 대로 대기용 첫 접속 주소를 폐쇄몰 메인 도메인으로 즉시 변경했습니다.
            driver.get("https://jy45321.imweb.me") 
        except Exception as e:
            messagebox.showerror("브라우저 실행 오류", f"크롬 브라우저를 실행할 수 없습니다.\n{e}")
            return

        messagebox.showinfo("대기 중", "로그인을 마쳤다면 이 창의 [확인]을 누르세요. 추출을 시작합니다.")

        wb = Workbook()
        ws = wb.active
        ws.title = "링크 목록"
        
        headers = [
            "카테고리", "상품명", "판매가", "배송비", "공급처", 
            "원가", "나의 배송비", "포장비", "판매처", "수수료(%)", 
            "수수료", "부가세", "마진", "마진율"
        ]
        ws.append(headers)

        count = 0
        url_file_found = 0

        # 데이터 탐색 및 수집
        for root_dir, dirs, files in os.walk(base_dir):
            if any(word in root_dir for word in exclude_list):
                continue

            for file in files:
                if file.lower().endswith(".url"):
                    url_file_found += 1
                    file_path = os.path.join(root_dir, file)
                    url = self.extract_url_from_file(file_path)

                    if url and "jy45321.imweb.me" in url.lower():
                        rel_path = os.path.relpath(root_dir, base_dir)
                        category_name = "최상위 폴더" if rel_path == "." else rel_path.replace(os.sep, " > ")
                        
                        # 파일명에서 확장자를 제외한 깨끗한 문자열 추출
                        product_name = os.path.splitext(file)[0]
                        hyperlink_formula = f'=HYPERLINK("{url}", "{url}")'

                        cost_val = ""         
                        delivery_fee_val = "" 
                        
                        try:
                            driver.get(url)
                            
                            # 반자동 모드일 경우 사람이 여유롭게 옵션을 클릭할 수 있도록 5.0초 제공
                            if current_mode == "manual":
                                time.sleep(5.0)
                            else:
                                time.sleep(1.8)  # 자동 모드는 기본 로딩 대기만 수행
                            
                            # 실시간 화면 전체 텍스트 수집
                            page_text = driver.find_element(By.TAG_NAME, "body").text
                            
                            # 1) 원가 파싱: '총 상품금액' 추적 및 괄호 수량 건너뛰기 규칙 적용
                            if "총 상품금액" in page_text:
                                cost_part = page_text.split("총 상품금액", 1)[1]
                                # (1개), (2개) 뒤에 나오는 진짜 원가 숫자만 타겟팅
                                cost_match = re.search(r'\(\d+개\)\s*([\d,]+)', cost_part)
                                if cost_match:
                                    cost_val = cost_match.group(1).replace(",", "").strip()
                                else:
                                    # 만약 옵션을 안 골라서 (X개) 표시가 없을 때를 대비한 백업용 기존 파싱
                                    cost_match = re.search(r'[\d,]+', cost_part)
                                    if cost_match:
                                        cost_val = cost_match.group().replace(",", "").strip()
                            
                            # 자동 모드이거나 옵션을 선택하지 않아 '총 상품금액'이 아예 없을 때 기본가 수집
                            if not cost_val or cost_val == "0":
                                try:
                                    price_element = driver.find_element(By.CSS_SELECTOR, "span.shop_item_price")
                                    cost_val = price_element.text.replace("원", "").replace(",", "").strip()
                                except:
                                    pass

                            # 2) 배송비 파싱: '기본' 추적 및 무료 분기
                            if "기본" in page_text:
                                delivery_part = page_text.split("기본", 1)[1].strip()
                                target_area = delivery_part[:15]
                                
                                if "무료" in target_area:
                                    delivery_fee_val = 0  
                                else:
                                    deliv_match = re.search(r'[\d,]+', target_area)
                                    if deliv_match:
                                        delivery_fee_val = deliv_match.group().replace(",", "").strip()
                                        
                        except Exception:
                            cost_val = "오류(재확인)"
                            delivery_fee_val = "오류"

                        row_data = [
                            category_name, product_name, "", delivery_fee_val, hyperlink_formula,
                            cost_val, "", "", "", "", "", "", "", ""
                        ]
                        ws.append(row_data)
                        count += 1

        driver.quit()  

        if count > 0:
            try:
                header_fill = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
                header_font = Font(name="맑은 고딕", size=11, bold=True)
                header_alignment = Alignment(horizontal="center", vertical="center")

                for col_idx in range(1, len(headers) + 1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = header_alignment

                for col_idx in range(1, len(headers) + 1):
                    max_len = 0
                    col_letter = get_column_letter(col_idx)
                    for row_idx in range(1, ws.max_row + 1):
                        cell_value = ws.cell(row=row_idx, column=col_idx).value
                        if cell_value:
                            val_str = str(cell_value)
                            if val_str.startswith("="):
                                val_str = "https://example.com"
                            byte_len = len(val_str.encode("utf-8"))
                            calc_len = (byte_len - len(val_str)) / 2 + len(val_str)
                            if calc_len > max_len:
                                max_len = calc_len
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

                output_path = os.path.join(base_dir, "Link_List.xlsx")
                wb.save(output_path)
                messagebox.showinfo(
                    "성공", 
                    f"총 {count}개의 데이터 및 단가 추출 완료!\n\n확인을 누르면 마진 계산서 양식의 엑셀 파일이 열립니다."
                )
                os.startfile(output_path)
            except Exception as e:
                messagebox.showerror("엑셀 저장 오류", f"엑셀 파일을 저장하거나 여는 도중 에러가 발생했습니다:\n{e}")
        else:
            messagebox.showwarning(
                "실패", 
                f"스캔된 총 {url_file_found}개의 .url 파일 중 타겟 폐쇄몰 도메인 주소(jy45321.imweb.me)와 일치하는 유효 링크가 없거나 주소를 읽지 못했습니다."
            )


# 프로그램 정식 시작 지점 선언문
if __name__ == "__main__":
    root = tk.Tk()
    app = LinkExtractorApp(root)
    root.mainloop()
