import os
from tkinter import filedialog, messagebox
import tkinter as tk
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter  # [핵심 수정] 안전한 열 문자 변환 함수 추가


class LinkExtractorApp:

    def __init__(self, root):
        self.root = root
        self.root.title("인터넷 바로가기(.url) 링크 추출기")
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

        lbl_desc = tk.Label(
            filter_frame,
            text="제외할 폴더명을 입력하세요 (여러 개는 쉼표[,]로 구분):",
            fg="gray",
        )
        lbl_desc.pack(anchor="w", pady=(0, 5))

        entry_filter = tk.Entry(
            filter_frame, textvariable=self.exclude_words, width=58
        )
        entry_filter.pack(fill="x", ipady=2)
        entry_filter.insert(0, "품절, 제외, 보류")

        # 3. 실행 버튼 섹션
        btn_start = tk.Button(
            self.root,
            text="엑셀 링크 추출 시작",
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
                                return parts[1].strip()
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

                    if url:
                        rel_path = os.path.relpath(root_dir, base_dir)
                        if rel_path == ".":
                            category_name = "최상위 폴더"
                        else:
                            category_name = rel_path.replace(os.sep, " > ")

                        # 확장자를 제외한 순수 파일명(상품명)만 문자열로 안전하게 추출
                        product_name = os.path.splitext(file)[0]
                        hyperlink_formula = f'=HYPERLINK("{url}", "{url}")'

                        row_data = [
                            category_name,     # 카테고리
                            product_name,      # 상품명
                            "",                # 판매가
                            "",                # 배송비
                            hyperlink_formula, # 공급처
                            "",                # 원가
                            "",                # 나의 배송비
                            "",                # 포장비
                            "",                # 판매처
                            "",                # 수수료(%)
                            "",                # 수수료
                            "",                # 부가세
                            "",                # 마진
                            ""                 # 마진율
                        ]
                        
                        ws.append(row_data)
                        count += 1

        messagebox.showinfo(
            "스캔 결과", 
            f"발견된 .url 파일 수: {url_file_found}개\n실제 추출 성공한 링크 수: {count}개"
        )

        if count > 0:
            try:
                # 첫 번째 행 스타일 지정
                header_fill = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
                header_font = Font(name="맑은 고딕", size=11, bold=True)
                header_alignment = Alignment(horizontal="center", vertical="center")

                for col_idx in range(1, len(headers) + 1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = header_alignment

                # [핵심 수정] 튜플 에러 근본적 방지 문법으로 전면 교체
                # 가변적인 col.column_letter 대신 인덱스 번호를 직접 문자로 변환하는 표준 방식 적용
                for col_idx in range(1, len(headers) + 1):
                    max_len = 0
                    col_letter = get_column_letter(col_idx)  # 예: 1 -> 'A', 5 -> 'E'
                    
                    # 해당 열의 모든 행 값을 정확히 조회
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

                # 파일 저장
                output_path = os.path.join(base_dir, "Link_List.xlsx")
                wb.save(output_path)

                messagebox.showinfo(
                    "성공",
                    f"저장 경로:\n{output_path}\n\n확인을 누르면 마진 계산서 양식의 엑셀 파일이 열립니다.",
                )
                
                os.startfile(output_path)
                
            except Exception as e:
                messagebox.showerror("엑셀 저장 오류", f"엑셀 파일을 저장하거나 여는 도중 에러가 발생했습니다:\n{e}")
                
        else:
            messagebox.showwarning(
                "실패", 
                "폴더 안에 바로가기 파일은 있으나, 링크 주소 추출에 실패했습니다."
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = LinkExtractorApp(root)
    root.mainloop()
