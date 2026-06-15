import os
from tkinter import filedialog, messagebox
import tkinter as tk
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


class LinkExtractorApp:

    def __init__(self, root):
        self.root = root
        self.root.title("인터넷 바로가기(.url) 링크 추출기")
        self.root.geometry("550x260")
        self.root.resizable(False, False)

        # 변수 설정
        self.selected_path = tk.StringVar()
        self.exclude_words = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        # 1. 폴더 선택 섹션
        folder_frame = tk.LabelFrame(self.root, text=" 1. 대상 폴더 지정 ", padding=10)
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
            self.root, text=" 2. 제외할 하위 폴더 키워드 ", padding=10
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
        entry_filter.insert(0, "품절, 제외, 보류")  # 기본 예시 텍스트 제공

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
        try:
            for encoding in ["utf-8", "cp949", "utf-16"]:
                try:
                    with open(file_path, "r", encoding=encoding) as f:
                        for line in f:
                            if line.strip().startswith("URL="):
                                return line.strip().split("URL=", 1)
                except (UnicodeDecodeError, PermissionError):
                    continue
        except Exception:
            pass
        return None

    def start_extraction(self):
        base_dir = self.selected_path.get()

        if not base_dir:
            messagebox.showwarning("경고", "먼저 링크를 추출할 폴더를 지정해 주세요.")
            return

        raw_input = self.exclude_words.get()
        exclude_list = [
            w.strip() for w in raw_input.split(",") if w.strip()
        ]

        # 엑셀 파일 초기화
        wb = Workbook()
        ws = wb.active
        ws.title = "링크 목록"
        
        headers = ["카테고리", "상품명", "링크"]
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

                    if url:
                        rel_path = os.path.relpath(root_dir, base_dir)
                        if rel_path == ".":
                            category_name = "최상위 폴더"
                        else:
                            category_name = rel_path.replace(os.sep, " > ")

                        product_name = os.path.splitext(file)[0]
                        hyperlink_formula = f'=HYPERLINK("{url}", "{url}")'

                        ws.append([category_name, product_name, hyperlink_formula])
                        count += 1

        # 결과 저장 및 후처리
        if count > 0:
            # 첫 번째 행(헤더) 스타일 지정
            header_fill = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
            header_font = Font(name="맑은 고딕", size=11, bold=True)
            header_alignment = Alignment(horizontal="center", vertical="center")

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_alignment

            # 열 너비 자동 조절
            for col in ws.columns:
                max_len = 0
                col_letter = col.column_letter
                for cell in col:
                    if cell.value:
                        val_str = str(cell.value)
                        if val_str.startswith("=HYPERLINK"):
                            val_str = url
                        byte_len = len(val_str.encode("utf-8"))
                        calc_len = (byte_len - len(val_str)) / 2 + len(val_str)
                        if calc_len > max_len:
                            max_len = calc_len
                ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

            output_path = os.path.join(base_dir, "Link_List.xlsx")
            wb.save(output_path)

            filter_str = f" (제외: {', '.join(exclude_list)})" if exclude_list else ""
            messagebox.showinfo(
                "성공",
                f"총 {count}개의 상품 링크 추출 완료!{filter_str}\n\n확인을 누르면 엑셀 파일이 바로 열립니다.",
            )
            
            # [핵심 수정] 저장된 엑셀 파일을 윈도우 시스템 명령어로 자동 실행
            try:
                os.startfile(output_path)
            except Exception as e:
                messagebox.showerror("오류", f"파일을 자동으로 여는 중 오류가 발생했습니다:\n{e}")
                
        else:
            messagebox.showwarning(
                "실패", "조건에 맞는 .url 파일이 해당 폴더 내에 존재하지 않습니다."
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = LinkExtractorApp(root)
    root.mainloop()
