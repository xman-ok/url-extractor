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

        self.selected_path = tk.StringVar()
        self.exclude_words = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
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
        # 다양한 윈도우/웹 인코딩 방식을 순서대로 대입하며 파일 읽기 시도
        # cp949(한국어 윈도우 기본), utf-8(기본), utf-16(MS 웹 바로가기 특수 포맷) 전부 대응
        encodings = ["cp949", "utf-8", "utf-16", "utf-8-sig", "latin-1"]
        
        for encoding in encodings:
            try:
                with open(file_path, "r", encoding=encoding, errors="ignore") as f:
                    for line in f:
                        clean_line = line.strip()
                        # 대소문자 구분 없이 URL= 패턴 검색
                        if clean_line.upper().startswith("URL="):
                            # URL= 뒷부분의 링크만 분리하여 반환
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
        
        headers = ["카테고리", "상품명", "링크"]
        ws.append(headers)

        count = 0
        url_file_found = 0

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

                        # [오류 수정 반영] 튜플 분리가 아닌 문자열 정상 추출 완료
                        product_name = os.path.splitext(file)[0]
                        hyperlink_formula = f'=HYPERLINK("{url}", "{url}")'

                        ws.append([category_name, product_name, hyperlink_formula])
                        count += 1

        # [진단 안내] 스캔 통계 확인
        messagebox.showinfo(
            "스캔 결과", 
            f"발견된 .url 파일 수: {url_file_found}개\n실제 추출 성공한 링크 수: {count}개"
        )

        if count > 0:
            try:
                # 헤더 서식 지정
                header_fill = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
                header_font = Font(name="맑은 고딕", size=11, bold=True)
                header_alignment = Alignment(horizontal="center", vertical="center")

                for cell in ws[1]:  # 첫 번째 행 스타일 지정
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
                                val_str = "https://example.com"
                            byte_len = len(val_str.encode("utf-8"))
                            calc_len = (byte_len - len(val_str)) / 2 + len(val_str)
                            if calc_len > max_len:
                                max_len = calc_len
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

                # 파일 저장
                output_path = os.path.join(base_dir, "Link_List.xlsx")
                wb.save(output_path)

                filter_str = f" (제외: {', '.join(exclude_list)})" if exclude_list else ""
                messagebox.showinfo(
                    "성공",
                    f"저장 경로:\n{output_path}\n\n확인을 누르면 엑셀 파일이 열립니다.",
                )
                
                # 파일 자동 열기
                os.startfile(output_path)
                
            except Exception as e:
                # 저장 단계에서 오류 발생 시 팝업으로 상세 내용 표시
                messagebox.showerror("엑셀 저장 오류", f"엑셀 파일을 저장하거나 여는 도중 에러가 발생했습니다:\n{e}")
                
        else:
            messagebox.showwarning(
                "실패", 
                "폴더 안에 바로가기 파일은 있으나, 파일 내부에서 링크 주소(URL=)를 읽어내지 못했습니다.\n파일 손상이나 인코딩을 확인해 주세요."
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = LinkExtractorApp(root)
    root.mainloop()
