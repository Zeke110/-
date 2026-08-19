import datetime
import urllib.parse

import openpyxl
import pandas as pd
import requests
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class LabChemicalManager:

    def __init__(self, root):
        self.root = root
        self.root.title("시약 & 용액 재고 관리 및 자동 정보 채움 프로그램")
        self.root.geometry("1250x760")

        self.df = None
        self.file_path = None

        # ------------------------------------------------------------------
        # 새 엑셀 양식(제품명 / 제조사명 / CAT No. / CAS No. / 용량 / 잔량 /
        # 담당자 / 보관위치 / 개봉 시 유통기한 / 유해·위험성 / 등록일 / 상태)
        # 에 맞춘 컬럼 목록
        # ------------------------------------------------------------------
        self.columns = [
            "제품명",
            "제조사명",
            "CAT No.",
            "CAS No.",
            "용량",
            "잔량",
            "담당자",
            "보관위치",
            "개봉 시 유통기한",
            "유해·위험성",
            "등록일",
            "상태",
        ]

        self.setup_ui()

    # ----------------------------------------------------------------------
    # UI 구성
    # ----------------------------------------------------------------------
    def setup_ui(self):
        # 1. 상단 메뉴바
        toolbar = tk.Frame(self.root, bg="#f5f5f5", pady=8, padx=10)
        toolbar.pack(fill=tk.X, side=tk.TOP)

        tk.Button(
            toolbar,
            text="📁 엑셀 열기",
            command=self.load_excel,
            bg="#ffffff",
            relief="groove",
            width=12,
            font=("맑은 고딕", 9, "bold"),
        ).pack(side=tk.LEFT, padx=4)

        tk.Button(
            toolbar,
            text="💾 엑셀 저장",
            command=self.save_excel,
            bg="#2196F3",
            fg="white",
            relief="flat",
            width=12,
            font=("맑은 고딕", 9, "bold"),
        ).pack(side=tk.LEFT, padx=4)

        tk.Button(
            toolbar,
            text="🌐 빈 정보 자동 채우기 (PubChem API)",
            command=self.auto_fill_missing_data,
            bg="#4CAF50",
            fg="white",
            relief="flat",
            font=("맑은 고딕", 9, "bold"),
            padx=10,
        ).pack(side=tk.LEFT, padx=15)

        # 2. 검색 프레임
        search_frame = tk.LabelFrame(
            self.root,
            text=" 🔍 데이터 실시간 검색 ",
            font=("맑은 고딕", 9, "bold"),
            padx=10,
            pady=5,
        )
        search_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(search_frame, text="검색어:").pack(side=tk.LEFT, padx=2)
        self.entry_search = tk.Entry(search_frame, width=30)
        self.entry_search.pack(side=tk.LEFT, padx=5)
        self.entry_search.bind("<Return>", lambda event: self.search_data())

        tk.Button(
            search_frame, text="검색", command=self.search_data, width=8
        ).pack(side=tk.LEFT, padx=2)
        tk.Button(
            search_frame, text="초기화", command=self.reset_search, width=8
        ).pack(side=tk.LEFT, padx=2)

        # 3. 데이터 입력 / 수정 프레임
        edit_frame = tk.LabelFrame(
            self.root,
            text=" 📝 항목 추가 / 수정 / 삭제 ",
            font=("맑은 고딕", 9, "bold"),
            padx=10,
            pady=5,
        )
        edit_frame.pack(fill=tk.X, padx=10, pady=5)

        self.entries = {}

        # 12개 컬럼을 6개씩 두 줄로 배치
        r1_frame = tk.Frame(edit_frame)
        r1_frame.pack(fill=tk.X, pady=2)

        for col in self.columns[:6]:
            lbl = tk.Label(r1_frame, text=col, width=13, anchor="e")
            lbl.pack(side=tk.LEFT)
            ent = tk.Entry(r1_frame, width=12)
            ent.pack(side=tk.LEFT, padx=2)
            self.entries[col] = ent

        r2_frame = tk.Frame(edit_frame)
        r2_frame.pack(fill=tk.X, pady=2)

        for col in self.columns[6:]:
            lbl = tk.Label(r2_frame, text=col, width=13, anchor="e")
            lbl.pack(side=tk.LEFT)
            ent = tk.Entry(r2_frame, width=12)
            ent.pack(side=tk.LEFT, padx=2)
            self.entries[col] = ent

        btn_box = tk.Frame(edit_frame)
        btn_box.pack(fill=tk.X, pady=6)

        tk.Button(
            btn_box,
            text="➕ 신규 항목 추가",
            command=self.add_item,
            bg="#009688",
            fg="white",
            font=("맑은 고딕", 9, "bold"),
        ).pack(side=tk.LEFT, padx=5)

        tk.Button(
            btn_box,
            text="✏️ 선택 항목 수정 적용",
            command=self.update_item,
            bg="#FF9800",
            fg="white",
            font=("맑은 고딕", 9, "bold"),
        ).pack(side=tk.LEFT, padx=5)

        tk.Button(
            btn_box,
            text="🗑️ 선택 항목 삭제",
            command=self.delete_item,
            bg="#F44336",
            fg="white",
            font=("맑은 고딕", 9, "bold"),
        ).pack(side=tk.LEFT, padx=5)

        tk.Button(
            btn_box, text="🧹 입력창 비우기", command=self.clear_entries
        ).pack(side=tk.LEFT, padx=5)

        # 4. 데이터 목록 표
        tree_frame = tk.Frame(self.root)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=self.columns,
            show="headings",
            xscrollcommand=scroll_x.set,
            yscrollcommand=scroll_y.set,
        )

        scroll_x.config(command=self.tree.xview)
        scroll_y.config(command=self.tree.yview)

        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)

        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center")

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

    # ----------------------------------------------------------------------
    # 엑셀 불러오기 (수정됨)
    #
    # 기존 코드는 pd.read_excel(file_path) 로 무조건 "1행"을 헤더로 인식했기
    # 때문에, 실제 재고현황 엑셀처럼 제목/기준일 행이 앞에 붙어있는 경우
    # (아래 예시 참고) 헤더를 전혀 못 찾고 컬럼이 다 비어버리는 문제가 있었음.
    #
    #   1행: [프로젝트 실험실습실] 재고현황
    #   2행: 기준일 : 2026.08.19
    #   3행: (빈 행)
    #   4행: 제품명 | 제조사명 | CAT No. | ... <- 진짜 헤더
    #   5행~: 실제 데이터
    #
    # 그래서 파일의 위쪽 몇 줄을 훑어서 self.columns 와 가장 많이 겹치는
    # 행을 헤더로 자동 인식하도록 변경함.
    # ----------------------------------------------------------------------
    def find_header_row(self, file_path, max_scan=20):
        """self.columns 와 가장 많이 일치하는 행을 찾아 0-based 행 번호 반환"""
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.active

        best_row = 0
        best_score = -1

        for i, row in enumerate(
            ws.iter_rows(min_row=1, max_row=max_scan, values_only=True)
        ):
            row_vals = [str(c).strip() if c is not None else "" for c in row]
            score = sum(1 for col in self.columns if col in row_vals)
            if score > best_score:
                best_score = score
                best_row = i

        wb.close()

        # 최소 절반 이상의 컬럼명이 일치해야 "헤더 행"으로 인정 (아니면 1행 사용)
        if best_score < max(2, len(self.columns) // 2):
            return 0
        return best_row

    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        try:
            header_row = self.find_header_row(file_path)
            self.df = pd.read_excel(file_path, header=header_row)

            # 컬럼명 앞뒤 공백 제거 (엑셀에 공백이 섞여있는 경우 대비)
            self.df.columns = [str(c).strip() for c in self.df.columns]

            # 완전히 빈 행 제거 (제목/기준일 아래 빈 줄 등)
            self.df = self.df.dropna(how="all")

            missing_cols = [c for c in self.columns if c not in self.df.columns]
            for col in missing_cols:
                self.df[col] = ""

            self.df = self.df[self.columns].reset_index(drop=True)

            self.file_path = file_path
            self.display_data(self.df)

            info_msg = "엑셀 파일을 성공적으로 불러왔습니다!"
            if missing_cols:
                info_msg += (
                    "\n\n※ 다음 컬럼은 파일에 없어 빈 값으로 추가되었습니다:\n"
                    + ", ".join(missing_cols)
                )
            messagebox.showinfo("성공", info_msg)
        except Exception as e:
            messagebox.showerror("오류", f"파일 읽기 실패: {e}")

    def display_data(self, df_to_show):
        self.tree.delete(*self.tree.get_children())
        if df_to_show is None or df_to_show.empty:
            return

        for idx, row in df_to_show.iterrows():
            row_vals = [
                "" if pd.isna(val) else str(val) for val in row[self.columns]
            ]
            self.tree.insert("", tk.END, iid=idx, values=row_vals)

    def on_tree_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        idx = int(selected[0])
        row = self.df.loc[idx]
        for col in self.columns:
            val = "" if pd.isna(row[col]) else str(row[col])
            self.entries[col].delete(0, tk.END)
            self.entries[col].insert(0, val)

    def clear_entries(self):
        for col in self.columns:
            self.entries[col].delete(0, tk.END)
        self.entries["등록일"].insert(
            0, datetime.date.today().strftime("%Y.%m.%d")
        )

    def add_item(self):
        if self.df is None:
            self.df = pd.DataFrame(columns=self.columns)

        new_row = {}
        for col in self.columns:
            new_row[col] = self.entries[col].get().strip()

        if not new_row["등록일"]:
            new_row["등록일"] = datetime.date.today().strftime("%Y.%m.%d")

        self.df = pd.concat(
            [self.df, pd.DataFrame([new_row])], ignore_index=True
        )
        self.display_data(self.df)
        self.clear_entries()
        messagebox.showinfo("성공", "신규 항목이 추가되었습니다.")

    def update_item(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("경고", "수정할 항목을 목록에서 선택해주세요.")
            return

        idx = int(selected[0])
        for col in self.columns:
            self.df.at[idx, col] = self.entries[col].get().strip()

        self.display_data(self.df)
        messagebox.showinfo("성공", "선택한 항목이 수정되었습니다.")

    def delete_item(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("경고", "삭제할 항목을 선택해주세요.")
            return

        idx = int(selected[0])
        if messagebox.askyesno("확인", "정말로 이 항목을 삭제하시겠습니까?"):
            self.df = self.df.drop(idx).reset_index(drop=True)
            self.display_data(self.df)
            self.clear_entries()
            messagebox.showinfo("성공", "항목이 삭제되었습니다.")

    def search_data(self):
        if self.df is None:
            return
        keyword = self.entry_search.get().strip()
        if not keyword:
            self.display_data(self.df)
            return

        mask = self.df.astype(str).apply(
            lambda x: x.str.contains(keyword, case=False).any(), axis=1
        )
        self.display_data(self.df[mask])

    def reset_search(self):
        self.entry_search.delete(0, tk.END)
        if self.df is not None:
            self.display_data(self.df)

    def auto_fill_missing_data(self):
        """진단 및 예외 처리가 강화된 자동 채우기 기능 (신규 컬럼 기준)"""
        if self.df is None or self.df.empty:
            messagebox.showwarning(
                "경고",
                "데이터가 없습니다. 엑셀을 불러오거나 항목을 먼저 추가해주세요.",
            )
            return

        today_str = datetime.date.today().strftime("%Y.%m.%d")
        updated_count = 0
        failed_list = []

        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
                " AppleWebKit/537.36"
            )
        }

        for idx, row in self.df.iterrows():
            # 등록일 채우기
            if pd.isna(row["등록일"]) or str(row["등록일"]).strip() == "":
                self.df.at[idx, "등록일"] = today_str

            # CAS No. 우선, 없을 경우 제품명 사용
            cas_val = str(row["CAS No."]).strip() if pd.notna(row["CAS No."]) else ""
            name_val = str(row["제품명"]).strip() if pd.notna(row["제품명"]) else ""

            query_id = cas_val if cas_val else name_val

            need_hazard = pd.isna(row["유해·위험성"]) or not str(
                row["유해·위험성"]
            ).strip()

            if need_hazard and query_id:
                try:
                    encoded_id = urllib.parse.quote(query_id)
                    url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{encoded_id}/property/Title,GHSClassification/JSON"

                    res = requests.get(url, headers=headers, timeout=5)

                    if res.status_code == 200:
                        data = res.json()
                        props = data["PropertyTable"]["Properties"][0]
                        ghs_list = props.get("GHSClassification", [])
                        hazard = (
                            ", ".join(ghs_list[:3])
                            if ghs_list
                            else "GHS 정보 없음"
                        )

                        if hazard:
                            self.df.at[idx, "유해·위험성"] = hazard
                            updated_count += 1
                    else:
                        failed_list.append(
                            f"{query_id} (API 응답코드: {res.status_code})"
                        )
                except Exception as e:
                    failed_list.append(f"{query_id} (오류: {e})")

        self.display_data(self.df)

        msg = f"자동 채우기 완료!\n- 업데이트된 항목 수: {updated_count}개"
        if failed_list:
            msg += "\n\n⚠️ 아래 항목은 정보를 찾지 못했거나 검색 실패했습니다:\n" + "\n".join(
                failed_list[:5]
            )
        messagebox.showinfo("결과 안내", msg)

    def save_excel(self):
        if self.df is None:
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return

        save_path = self.file_path
        if not save_path:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")]
            )

        if save_path:
            try:
                self.df.to_excel(save_path, index=False)
                self.file_path = save_path
                messagebox.showinfo(
                    "성공",
                    f"엑셀 파일이 정상적으로 저장되었습니다!\n({save_path})",
                )
            except Exception as e:
                messagebox.showerror("오류", f"저장 중 오류 발생: {e}")


if __name__ == "__main__":

    root = tk.Tk()
    app = LabChemicalManager(root)
    root.mainloop()
