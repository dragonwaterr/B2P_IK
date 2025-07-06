import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import platform
import subprocess
from core.ppt_generator import load_bible_data, parse_selection, get_verses, create_ppt, BIBLE_BOOK_ABBR
import json
import threading
import time

class Bible2PPTApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bible2PPT for IlKwang - 성경 PPT 생성기")
        self.bible_data_path = os.path.join("data", "cache", "bible_data.json")
        self.bible_data = None
        self.bg_image_path_map = self.load_bg_image_paths()
        self.config_data = self.load_config()
        self.setup_main_ui()

    def load_config(self):
        config_path = os.path.join("data", "config.json")
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {
            "template_names": {"1": "템플릿 1", "2": "템플릿 2", "3": "템플릿 3"},
            "max_chars_per_slide": {"1": 500, "2": 400, "3": 600}
        }

    def save_config(self):
        config_path = os.path.join("data", "config.json")
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(self.config_data, f, ensure_ascii=False, indent=2)

    def setup_main_ui(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        self.bible_data = load_bible_data()
        self.setup_ui()

    def load_bg_image_paths(self):
        bg_json = os.path.join("data", "bg_images.json")
        if os.path.exists(bg_json):
            with open(bg_json, "r", encoding="utf-8") as f:
                return json.load(f)
        return {"1": "", "2": "", "3": ""}

    def save_bg_image_paths(self):
        bg_json = os.path.join("data", "bg_images.json")
        with open(bg_json, "w", encoding="utf-8") as f:
            json.dump(self.bg_image_path_map, f, ensure_ascii=False, indent=2)

    def setup_ui(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        # 전체를 감싸는 메인 프레임
        main_frame = tk.Frame(self.root)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=2)
        main_frame.grid_columnconfigure(1, weight=1)

        # 좌측 입력 영역
        left_frame = tk.LabelFrame(main_frame, text="성경 구절 및 저장 위치", padx=10, pady=10)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left_frame.grid_rowconfigure(0, weight=2)  # 구절범위 크게
        left_frame.grid_rowconfigure(1, weight=1)
        left_frame.grid_rowconfigure(2, weight=1)
        left_frame.grid_columnconfigure(1, weight=1)

        # 구절 범위 입력란 (크게)
        tk.Label(left_frame, text="구절 범위:").grid(row=0, column=0, sticky="e", pady=(10, 10))
        self.selection_entry = tk.Entry(left_frame, width=40, font=("맑은 고딕", 14))
        self.selection_entry.insert(0, "창1:1-3,2:1-2; 왕상3:1-5")
        self.selection_entry.grid(row=0, column=1, sticky="ew", pady=(10, 10), columnspan=2)

        # ENTER 키 이벤트 바인딩
        self.selection_entry.bind('<Return>', lambda event: self.generate_ppt())

        # 출력 폴더 (작고 통일)
        tk.Label(left_frame, text="출력 폴더:").grid(row=1, column=0, sticky="e", pady=(5, 5))
        self.output_dir_var = tk.StringVar()
        self.output_dir_var.set(os.path.join(os.path.expanduser("~"), "Desktop"))
        self.output_dir_entry = tk.Entry(left_frame, textvariable=self.output_dir_var, width=25, font=("맑은 고딕", 11))
        self.output_dir_entry.grid(row=1, column=1, sticky="ew", pady=(5, 5))
        tk.Button(left_frame, text="폴더 선택", command=self.browse_dir, width=10, font=("맑은 고딕", 11)).grid(row=1, column=2, sticky="ew", padx=(5, 0))

        # 저장 파일명 (작고 통일)
        tk.Label(left_frame, text="저장 파일명:").grid(row=2, column=0, sticky="e", pady=(5, 10))
        self.output_entry = tk.Entry(left_frame, width=25, font=("맑은 고딕", 11))
        self.output_entry.insert(0, "output.pptx")
        self.output_entry.grid(row=2, column=1, sticky="ew", pady=(5, 10), columnspan=2)

        # 우측 템플릿/배경 영역
        right_frame = tk.LabelFrame(main_frame, text="템플릿 및 배경", padx=10, pady=10)
        right_frame.grid(row=0, column=1, sticky="nsew")
        right_frame.grid_columnconfigure(1, weight=1)

        # 템플릿 선택 (기존)
        tk.Label(right_frame, text="템플릿:", font=("맑은 고딕", 11)).grid(row=0, column=0, sticky="w", pady=(5, 10))
        self.template_var = tk.StringVar(value="1")
        
        # 템플릿 이름으로 콤보박스 생성 (번호 제거)
        template_names = [self.config_data['template_names'][str(i)] for i in range(1, 4)]
        template_combo = ttk.Combobox(right_frame, textvariable=self.template_var, values=template_names, state="readonly", width=15, font=("맑은 고딕", 11))
        template_combo.grid(row=0, column=1, sticky="ew", pady=(5, 10), padx=(5, 0))
        template_combo.bind("<<ComboboxSelected>>", self.update_template_info)
        
        # 초기값 설정
        template_combo.set(template_names[0])

        # 배경 이미지 선택/삭제 버튼 (나란히 배치)
        bg_frame = tk.Frame(right_frame)
        bg_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 2))
        bg_frame.grid_columnconfigure(0, weight=1)
        bg_frame.grid_columnconfigure(1, weight=1)
        
        bg_select_btn = tk.Button(bg_frame, text="배경 이미지\n선택", command=self.select_bg_image, 
                                 width=9, height=2, font=("맑은 고딕", 9))
        bg_select_btn.grid(row=0, column=0, sticky="ew", padx=(0, 2))
        
        bg_delete_btn = tk.Button(bg_frame, text="배경 이미지\n삭제", command=self.delete_bg_image, 
                                 width=9, height=2, font=("맑은 고딕", 9))
        bg_delete_btn.grid(row=0, column=1, sticky="ew", padx=(2, 0))

        self.bg_image_label = tk.Label(right_frame, text=self.get_current_bg_image_name(), fg="gray", font=("맑은 고딕", 9))
        self.bg_image_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 10))

        # 최대 글자 수 설정
        tk.Label(right_frame, text="최대 글자 수:", font=("맑은 고딕", 10)).grid(row=3, column=0, sticky="w", pady=(5, 5))
        self.max_chars_var = tk.StringVar(value=str(self.config_data["max_chars_per_slide"]["1"]))
        max_chars_entry = tk.Entry(right_frame, textvariable=self.max_chars_var, width=8, font=("맑은 고딕", 10))
        max_chars_entry.grid(row=3, column=1, sticky="w", pady=(5, 5), padx=(5, 0))
        max_chars_entry.bind('<KeyRelease>', self.save_max_chars)

        # 템플릿 편집 버튼
        edit_btn = tk.Button(right_frame, text="템플릿 편집", command=self.edit_template, width=18, height=2, font=("맑은 고딕", 11))
        edit_btn.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10, 5))

        # 템플릿 이름 수정 버튼 (우하단에 작게)
        name_edit_btn = tk.Button(right_frame, text="템플릿 이름 수정", command=self.edit_template_names, 
                                 width=15, height=1, font=("맑은 고딕", 8))
        name_edit_btn.grid(row=5, column=0, columnspan=2, sticky="se", pady=(5, 0))

        # 중앙 하단 PPT 생성 버튼 (width=20, 중앙 정렬, 양쪽 여백)
        ppt_btn = tk.Button(main_frame, text="PPT 생성", command=self.generate_ppt, width=20, height=2, bg="#4A90E2", fg="white", font=("맑은 고딕", 12, "bold"))
        ppt_btn.grid(row=1, column=0, columnspan=2, pady=(20, 5), sticky="")
        main_frame.grid_rowconfigure(1, minsize=60)
        main_frame.grid_columnconfigure(0, minsize=60)
        main_frame.grid_columnconfigure(1, minsize=60)

        # 상태 메시지 (하단 전체)
        self.status_var = tk.StringVar()
        tk.Label(main_frame, textvariable=self.status_var, fg="blue", font=("맑은 고딕", 11)).grid(row=3, column=0, columnspan=2, pady=(5, 10), sticky="ew")

        # 성경 66권 fullname:약어 사전 (가장 하단)
        dict_frame = tk.LabelFrame(self.root, text="성경 66권 책이름/약어 사전", padx=5, pady=5)
        dict_frame.grid(row=10, column=0, columnspan=2, sticky="nsew", padx=10, pady=(10, 10))
        self.root.grid_rowconfigure(10, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        tk.Label(dict_frame, text="책이름 검색:").pack(anchor="w")
        self.bible_search_var = tk.StringVar()
        self.bible_search_var.trace_add('write', self.update_bible_dict_list)
        search_entry = tk.Entry(dict_frame, textvariable=self.bible_search_var, width=20)
        search_entry.pack(anchor="w", fill="x")

        list_frame = tk.Frame(dict_frame)
        list_frame.pack(fill="both", expand=True)
        self.bible_dict_listbox = tk.Listbox(list_frame, height=10)
        self.bible_dict_listbox.pack(side="left", fill="both", expand=True)
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.bible_dict_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.bible_dict_listbox.config(yscrollcommand=scrollbar.set)

        self.fullname_to_abbr = {v: k for k, v in BIBLE_BOOK_ABBR.items()}
        self.bible_fullnames = [
            "창세기", "출애굽기", "레위기", "민수기", "신명기", "여호수아", "사사기", "룻기", "사무엘상", "사무엘하", "열왕기상", "열왕기하", "역대상", "역대하", "에스라", "느헤미야", "에스더", "욥기", "시편", "잠언", "전도서", "아가", "이사야", "예레미야", "예레미야애가", "에스겔", "다니엘", "호세아", "요엘", "아모스", "오바댜", "요나", "미가", "나훔", "하박국", "스바냐", "학개", "스가랴", "말라기", "마태복음", "마가복음", "누가복음", "요한복음", "사도행전", "로마서", "고린도전서", "고린도후서", "갈라디아서", "에베소서", "빌립보서", "골로새서", "데살로니가전서", "데살로니가후서", "디모데전서", "디모데후서", "디도서", "빌레몬서", "히브리서", "야고보서", "베드로전서", "베드로후서", "요한1서", "요한2서", "요한3서", "유다서", "요한계시록"
        ]
        self.update_bible_dict_list()
        self.update_template_info()

    def update_bible_dict_list(self, *args):
        search = self.bible_search_var.get().strip()
        self.bible_dict_listbox.delete(0, tk.END)
        for fullname in self.bible_fullnames:
            abbr = self.fullname_to_abbr.get(fullname, "-")
            if not search or search in fullname:
                self.bible_dict_listbox.insert(tk.END, f"{fullname} : {abbr}")

    def browse_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PPTX files", "*.pptx")])
        if file_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, file_path)

    def browse_dir(self):
        folder_selected = filedialog.askdirectory(initialdir=self.output_dir_var.get())
        if folder_selected:
            self.output_dir_var.set(folder_selected)

    def get_template_number(self):
        """템플릿 이름으로부터 번호를 찾는 헬퍼 함수"""
        selected = self.template_var.get()
        for num, name in self.config_data["template_names"].items():
            if name == selected:
                return num
        return "1"  # 기본값

    def get_current_bg_image_name(self):
        # 템플릿 번호 추출
        template_num = self.get_template_number()
        path = self.bg_image_path_map.get(template_num, "")
        return os.path.basename(path) if path else "(없음)"

    def update_bg_image_label(self):
        self.bg_image_label.config(text=self.get_current_bg_image_name())

    def select_bg_image(self):
        file_path = filedialog.askopenfilename(
            title="배경 이미지 선택",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif"), ("All Files", "*.*")]
        )
        if file_path:
            # 템플릿 번호 추출
            template_num = self.get_template_number()
            self.bg_image_path_map[template_num] = file_path
            self.save_bg_image_paths()
            self.update_bg_image_label()

    def delete_bg_image(self):
        """현재 템플릿의 배경 이미지 삭제"""
        # 템플릿 번호 추출
        template_num = self.get_template_number()
        if template_num in self.bg_image_path_map:
            del self.bg_image_path_map[template_num]
            self.save_bg_image_paths()
            self.update_bg_image_label()
            messagebox.showinfo("완료", "배경 이미지가 삭제되었습니다.")

    def save_max_chars(self, event=None):
        """최대 글자 수 설정 저장"""
        try:
            # 템플릿 번호 추출
            template_num = self.get_template_number()
            max_chars = int(self.max_chars_var.get())
            if max_chars > 0:
                self.config_data["max_chars_per_slide"][template_num] = max_chars
                self.save_config()
        except ValueError:
            pass  # 숫자가 아닌 경우 무시

    def edit_template(self):
        # 템플릿 번호 추출
        template_num = self.get_template_number()
        template_path = os.path.abspath(os.path.join("templates", f"base_template{template_num}.pptx"))
        if not os.path.exists(template_path):
            messagebox.showerror("오류", f"템플릿 파일이 없습니다: {template_path}")
            return
        try:
            if platform.system() == "Windows":
                os.startfile(template_path)
            elif platform.system() == "Darwin":
                subprocess.call(["open", template_path])
            else:
                subprocess.call(["xdg-open", template_path])
            messagebox.showinfo("안내", "템플릿을 수정한 후 저장/닫기 하세요.\n이후 모든 PPT는 이 템플릿을 따릅니다.")
        except Exception as e:
            messagebox.showerror("오류", f"템플릿 파일을 열 수 없습니다: {e}")

    def edit_template_names(self):
        """템플릿 이름 수정 다이얼로그"""
        dialog = tk.Toplevel(self.root)
        dialog.title("템플릿 이름 수정")
        dialog.geometry("450x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 중앙 정렬
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        tk.Label(dialog, text="각 템플릿의 이름을 설정하세요:", font=("맑은 고딕", 11)).pack(pady=(10, 20))
        
        entries = {}
        for i in range(1, 4):
            frame = tk.Frame(dialog)
            frame.pack(fill="x", padx=20, pady=5)
            
            tk.Label(frame, text=f"템플릿 {i}:", width=10).pack(side="left")
            entry = tk.Entry(frame, width=20)
            entry.pack(side="right", fill="x", expand=True)
            entry.insert(0, self.config_data["template_names"][str(i)])
            entries[str(i)] = entry
        
        def save_names():
            for template_num, entry in entries.items():
                self.config_data["template_names"][template_num] = entry.get()
            self.save_config()
            self.update_template_combo()  # 콤보박스 업데이트
            dialog.destroy()
            messagebox.showinfo("완료", "템플릿 이름이 저장되었습니다.")
        
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=20)
        tk.Button(button_frame, text="저장", command=save_names, width=8, height=1).pack(side="left", padx=5)
        tk.Button(button_frame, text="취소", command=dialog.destroy, width=8, height=1).pack(side="left", padx=5)

    def update_template_info(self, event=None):
        """템플릿 변경 시 관련 정보 업데이트"""
        # 템플릿 번호 추출 (템플릿 이름으로부터 번호 찾기)
        template_num = self.get_template_number()
        self.update_bg_image_label()
        self.max_chars_var.set(str(self.config_data["max_chars_per_slide"][template_num]))

    def update_template_combo(self):
        """템플릿 이름이 변경될 때 콤보박스 업데이트"""
        template_names = [self.config_data['template_names'][str(i)] for i in range(1, 4)]
        # 현재 선택된 템플릿 이름 유지
        current_selected = self.template_var.get()
        
        # 콤보박스 값 업데이트
        combo_widget = None
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, tk.LabelFrame) and child.cget("text") == "템플릿 및 배경":
                        for grandchild in child.winfo_children():
                            if isinstance(grandchild, ttk.Combobox):
                                combo_widget = grandchild
                                break
                        break
                if combo_widget:
                    break
        
        if combo_widget:
            combo_widget['values'] = template_names
            # 현재 템플릿 이름이 여전히 유효한지 확인
            if current_selected in template_names:
                combo_widget.set(current_selected)
                self.template_var.set(current_selected)
            else:
                # 유효하지 않으면 첫 번째 템플릿으로 설정
                combo_widget.set(template_names[0])
                self.template_var.set(template_names[0])

    def generate_ppt(self):
        version = "개역개정"  # 항상 개역개정만 사용
        selection_str = self.selection_entry.get()
        output_dir = self.output_dir_var.get()
        output_filename = self.output_entry.get()
        # 확장자 자동 추가
        if not output_filename.lower().endswith('.pptx'):
            output_filename += '.pptx'
        
        # 템플릿 번호 추출
        template_num = self.get_template_number()
        
        template_path = os.path.abspath(os.path.join("templates", f"base_template{template_num}.pptx"))
        bg_image_path = self.bg_image_path_map.get(template_num, "")
        max_chars = self.config_data["max_chars_per_slide"][template_num]

        if not version or not selection_str or not output_dir or not output_filename:
            messagebox.showerror("입력 오류", "모든 입력란을 채워주세요.")
            return

        output_path = os.path.join(output_dir, output_filename)

        # 파일명 중복 체크 및 경고
        if os.path.exists(output_path):
            if not messagebox.askyesno("파일 덮어쓰기 경고", f"이미 같은 이름의 파일이 존재합니다.\n덮어쓰시겠습니까?\n\n{output_path}"):
                self.status_var.set("PPT 생성이 취소되었습니다.")
                return

        try:
            selections = parse_selection(selection_str, self.bible_data, version)
            if not selections:
                raise ValueError("구절 범위 해석 실패")
            verses = get_verses(self.bible_data, version, selections)
            if not verses:
                raise ValueError("해당 구절을 찾을 수 없습니다.")
            create_ppt(verses, output_path, template_path, bg_image_path, max_chars)
            self.status_var.set(f"PPT 파일이 생성되었습니다: {output_path}")
            messagebox.showinfo("완료", f"PPT 파일이 생성되었습니다:\n{output_path}")
        except Exception as e:
            self.status_var.set(f"오류: {e}")
            messagebox.showerror("오류", str(e))

if __name__ == "__main__":
    import sys, os
    sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    root = tk.Tk()
    app = Bible2PPTApp(root)
    root.mainloop()

def main():
    root = tk.Tk()
    app = Bible2PPTApp(root)
    root.mainloop()