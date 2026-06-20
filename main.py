import os
import time
from tkinter import filedialog, messagebox
import tkinter as tk
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# 웹 크롤링을 위해 새롭게 도입된 라이브러리
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service


class LinkExtractorApp:

    def __init__(self, root):
        self.root = root
        self.root.title("인터넷 바로가기(.url) 링크 & 가격 추출기")
        self.root.geometry("550x260")
        self.root.resizable(False, False)

        self.selected_path = tk.StringVar()
        self.exclude_words = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        # 1. 폴더 선택 섹션
        folder_frame = tk.LabelFrame(self.root, text=" 1. 대상 폴더 지정 ", padx=10, pady=10)
        folder_frame.pack(fill="x", padx=15, pady=10)

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

        # 3. 실행 버튼 섹션
        btn_start = tk.Button(
            self.root,
            text="폐쇄몰 로그인 및 링크·가격 일괄 추출",
            font=("맑은 고딕", 11, "bold"),
            bg="#2ecc71",
            fg="white",
            command=self.start_extraction,
            height=2,
        )
        btn_start.pack(fill="x", padx=15, pady=15)

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
                                return parts.strip()
            except Exception:
                continue
        return None

    def start_extraction(self):
        base_dir = self.selected_path.get()

        if not base_dir:
            messagebox.showwarning("경고", "먼저 링크를 추출할 폴더를 지정해 주세요.")
            return

        raw_input = self.exclude_words.get()
        exclude_list = [w.strip() for w in raw_input.split(",") if w.strip()]

        # ---- [웹 브라우저 제어 시작] ----
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
            # 수집 대상 도메인 로그인 페이지 자동 이동
            driver.get("https://imweb.me") 
        except Exception as e:
            messagebox.showerror("브라우저 실행 오류", f"크롬 브라우저를 실행할 수 없습니다.\n크롬이 설치되어 있는지 확인하세요.\n{e}")
            return

        # 사용자가 브라우저에서 직접 로그인을 진행할 때까지 멈추고 대기(대화상자 보류 상태 활용)
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

        # 데이터 탐색 및 수집
        for root_dir, dirs, files in os.walk(base_dir):
            if any(word in root_dir for word in exclude_list):
                continue

            for file in files:
                if file.lower().endswith(".url"):
                    file_path = os.path.join(root_dir, file)
                    url = self.extract_url_from_file(file_path)

                    if url and "jy45321.imweb.me" in url:  # 지정된 도메인 링크만 크롤링 시도
                        rel_path = os.path.relpath(root_dir, base_dir)
                        category_name = "최상위 폴더" if rel_path == "." else rel_path.replace(os.sep, " > ")
                        product_name = os.path.splitext(file)
                        hyperlink_formula = f'=HYPERLINK("{url}", "{url}")'

                        # 브라우저 원격 이동 및 가격 수집
                        price_val = ""
                        try:
                            driver.get(url)
                            time.sleep(1.5)  # 페이지 로딩 안전 대기 시간
                            
                            # 아임웹 솔루션의 표준 실시간 가격 태그 위치 추적 (추후 타겟팅용)
                            price_element = driver.find_element(By.CSS_Split, "span.shop_item_price")
                            # "원", "," 등을 떼어내고 순수 숫자만 필터링하여 정제
                            price_val = price_element.text.replace("원", "").replace(",", "").strip()
                        except Exception:
                            price_val = "오류(재확인)"  # 수집 실패 시 예외 처리

                        row_data = [
                            category_name, product_name, price_val, "", hyperlink_formula,
                            "", "", "", "", "", "", "", "", ""
                        ]
                        ws.append(row_data)
                        count += 1

        driver.quit()  # 작업 완료 후 가상 브라우저 종료

        if count > 0:
            try:
                # 서식 지정 및 너비 조절 코드
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
                            if val_str.startswith("=HYPERLINK"):
                                val_str = "https://example.com"
                            byte_len = len(val_str.encode("utf-8"))
                            calc_len = (byte_len - len(val_str)) / 2 + len(val_str)
                            if calc_len > max_len:
                                max_len = calc_len
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

                output_path = os.path.join(base_dir, "Link_List.xlsx")
                wb.save(output_path)
                messagebox.showinfo("성공", f"가격 정보 수집 완료!\n\n저장 경로:\n{output_path}")
                os.startfile(output_path)
            except Exception as e:
                messagebox.showerror("오류", f"에러 발생: {e}")
        else:
            messagebox.showwarning("실패", "수집된 올바른 타겟 도메인 .url 파일이 없습니다.")


if __name__ == "__main__":
    root = tk.Tk()
    app = LinkExtractorApp(root)
    root.mainloop()
