import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import docx
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import difflib
import os
import shutil
from datetime import datetime

class WordComparatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("워드 파일 비교 프로그램")
        self.root.geometry("1200x800")
        self.root.configure(bg="#f0f0f0")
        
        # 파일 경로 저장 변수
        self.file1_path = ""
        self.file2_path = ""
        self.file1_content = ""
        self.file2_content = ""
        self.diff_data = []  # 비교 데이터 저장
        
        self.setup_ui()
    
    def setup_ui(self):
        # 메인 프레임
        main_frame = tk.Frame(self.root, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 제목
        title_label = tk.Label(main_frame, text="📄 워드 파일 비교 프로그램", 
                              font=("맑은 고딕", 20, "bold"), 
                              bg="#f0f0f0", fg="#2c3e50")
        title_label.pack(pady=(0, 10))
        
        # 설명
        desc_label = tk.Label(main_frame, text="💡 비교 파일의 원본 서식을 그대로 유지하며, 차이점만 색상으로 표시합니다", 
                              font=("맑은 고딕", 11), 
                              bg="#f0f0f0", fg="#7f8c8d")
        desc_label.pack(pady=(0, 20))
        
        # 파일 선택 프레임
        file_frame = tk.Frame(main_frame, bg="#f0f0f0")
        file_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 원본 파일 선택
        file1_frame = tk.LabelFrame(file_frame, text="📂 원본 파일 (기준)", 
                                   font=("맑은 고딕", 12, "bold"),
                                   bg="#ffe8e8", fg="#2c3e50", padx=10, pady=10)
        file1_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        self.file1_label = tk.Label(file1_frame, text="파일이 선택되지 않았습니다", 
                                   wraplength=300, justify=tk.LEFT,
                                   bg="#ffe8e8", fg="#7f8c8d")
        self.file1_label.pack(pady=5)
        
        self.file1_button = tk.Button(file1_frame, text="원본 파일 선택", 
                                     command=self.select_file1,
                                     bg="#e74c3c", fg="white", 
                                     font=("맑은 고딕", 10, "bold"),
                                     relief=tk.FLAT, padx=20, pady=5)
        self.file1_button.pack(pady=5)
        
        # 비교 파일 선택
        file2_frame = tk.LabelFrame(file_frame, text="📂 비교 파일 (서식 유지)", 
                                   font=("맑은 고딕", 12, "bold"),
                                   bg="#e8f8e8", fg="#2c3e50", padx=10, pady=10)
        file2_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        self.file2_label = tk.Label(file2_frame, text="파일이 선택되지 않았습니다", 
                                   wraplength=300, justify=tk.LEFT,
                                   bg="#e8f8e8", fg="#7f8c8d")
        self.file2_label.pack(pady=5)
        
        self.file2_button = tk.Button(file2_frame, text="비교 파일 선택", 
                                     command=self.select_file2,
                                     bg="#27ae60", fg="white", 
                                     font=("맑은 고딕", 10, "bold"),
                                     relief=tk.FLAT, padx=20, pady=5)
        self.file2_button.pack(pady=5)
        
        # 비교 방식 선택
        method_frame = tk.LabelFrame(main_frame, text="🔍 비교 방식 선택", 
                                     font=("맑은 고딕", 12, "bold"),
                                     bg="#f8f9fa", fg="#2c3e50", padx=10, pady=10)
        method_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.comparison_method = tk.StringVar(value="paragraph")
        
        method_content = tk.Frame(method_frame, bg="#f8f9fa")
        method_content.pack(fill=tk.X)
        
        tk.Radiobutton(method_content, text="📝 문단 단위 비교 (빠름, 권장)", 
                      variable=self.comparison_method, value="paragraph",
                      bg="#f8f9fa", font=("맑은 고딕", 10)).pack(side=tk.LEFT, padx=20)
        tk.Radiobutton(method_content, text="🔤 단어 단위 비교 (정밀함)", 
                      variable=self.comparison_method, value="word",
                      bg="#f8f9fa", font=("맑은 고딕", 10)).pack(side=tk.LEFT, padx=20)
        
        # 컨트롤 버튼 프레임
        control_frame = tk.Frame(main_frame, bg="#f0f0f0")
        control_frame.pack(pady=20)
        
        self.compare_button = tk.Button(control_frame, text="🔍 파일 비교하기", 
                                       command=self.compare_files,
                                       bg="#3498db", fg="white", 
                                       font=("맑은 고딕", 14, "bold"),
                                       relief=tk.FLAT, padx=40, pady=12,
                                       state=tk.DISABLED)
        self.compare_button.pack(side=tk.LEFT, padx=10)
        
        self.save_button = tk.Button(control_frame, text="💾 비교 결과 저장", 
                                    command=self.save_result,
                                    bg="#e67e22", fg="white", 
                                    font=("맑은 고딕", 14, "bold"),
                                    relief=tk.FLAT, padx=40, pady=12,
                                    state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=10)
        
        # 범례 프레임
        self.legend_frame = tk.Frame(main_frame, bg="#f0f0f0")
        self.legend_frame.pack(pady=(0, 10))
        
        legend_title = tk.Label(self.legend_frame, text="📊 범례 (저장된 워드 파일에서 보이는 색상)", 
                               font=("맑은 고딕", 12, "bold"),
                               bg="#f0f0f0", fg="#2c3e50")
        legend_title.pack()
        
        legend_content = tk.Frame(self.legend_frame, bg="#f0f0f0")
        legend_content.pack()
        
        # 범례 항목들
        removed_frame = tk.Frame(legend_content, bg="#f0f0f0")
        removed_frame.pack(side=tk.LEFT, padx=30)
        removed_color = tk.Label(removed_frame, text="  ", bg="#ffcdd2", relief=tk.SOLID, borderwidth=1)
        removed_color.pack(side=tk.LEFT, padx=5)
        tk.Label(removed_frame, text="원본에서 삭제된 부분", bg="#f0f0f0", font=("맑은 고딕", 11, "bold")).pack(side=tk.LEFT)
        
        added_frame = tk.Frame(legend_content, bg="#f0f0f0")
        added_frame.pack(side=tk.LEFT, padx=30)
        added_color = tk.Label(added_frame, text="  ", bg="#c8e6c9", relief=tk.SOLID, borderwidth=1)
        added_color.pack(side=tk.LEFT, padx=5)
        tk.Label(added_frame, text="비교파일에 추가된 부분", bg="#f0f0f0", font=("맑은 고딕", 11, "bold")).pack(side=tk.LEFT)
        
        # 처음에는 범례 숨기기
        self.legend_frame.pack_forget()
        
        # 결과 표시 프레임
        result_frame = tk.Frame(main_frame, bg="#ffffff")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 스크롤 가능한 텍스트 위젯
        self.result_text = scrolledtext.ScrolledText(result_frame, 
                                                    wrap=tk.WORD, 
                                                    font=("맑은 고딕", 11),
                                                    bg="white", fg="#2c3e50",
                                                    height=25)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 텍스트 태그 설정 (색상 지정)
        self.result_text.tag_configure("added", background="#c8e6c9", foreground="#2e7d32")
        self.result_text.tag_configure("removed", background="#ffcdd2", foreground="#c62828", overstrike=True)
        self.result_text.tag_configure("normal", background="white", foreground="#2c3e50")
        
        # 상태바
        self.status_var = tk.StringVar()
        self.status_var.set("원본 파일과 비교 파일을 선택해주세요")
        status_bar = tk.Label(main_frame, textvariable=self.status_var, 
                             relief=tk.SUNKEN, anchor=tk.W,
                             bg="#ecf0f1", fg="#2c3e50", font=("맑은 고딕", 9))
        status_bar.pack(fill=tk.X, pady=(10, 0))
    
    def extract_text_content(self, file_path):
        """워드 문서에서 텍스트 내용을 추출합니다."""
        try:
            doc = docx.Document(file_path)
            content = []
            
            # 본문 내용 추출
            for paragraph in doc.paragraphs:
                content.append(paragraph.text)
            
            # 표 내용 추출
            for table in doc.tables:
                for row in table.rows:
                    row_text = []
                    for cell in row.cells:
                        row_text.append(cell.text.strip())
                    content.append(" | ".join(row_text))
            
            return '\n'.join(content)
            
        except Exception as e:
            raise Exception(f"워드 파일 읽기 오류: {str(e)}")
    
    def select_file1(self):
        file_path = filedialog.askopenfilename(
            title="원본 파일 선택",
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")]
        )
        if file_path:
            self.file1_path = file_path
            self.file1_label.config(text=f"선택됨: {os.path.basename(file_path)}", fg="#27ae60")
            self.status_var.set(f"원본 파일 선택됨: {os.path.basename(file_path)}")
            self.load_file_content(file_path, 1)
            self.update_compare_button()
    
    def select_file2(self):
        file_path = filedialog.askopenfilename(
            title="비교 파일 선택 (이 파일의 서식이 유지됩니다)",
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")]
        )
        if file_path:
            self.file2_path = file_path
            self.file2_label.config(text=f"선택됨: {os.path.basename(file_path)}", fg="#27ae60")
            self.status_var.set(f"비교 파일 선택됨: {os.path.basename(file_path)}")
            self.load_file_content(file_path, 2)
            self.update_compare_button()
    
    def load_file_content(self, file_path, file_num):
        try:
            content = self.extract_text_content(file_path)
            
            if file_num == 1:
                self.file1_content = content
            else:
                self.file2_content = content
                
        except Exception as e:
            messagebox.showerror("오류", f"파일을 읽는데 실패했습니다: {str(e)}")
            if file_num == 1:
                self.file1_label.config(text="파일 읽기 실패", fg="#e74c3c")
                self.file1_content = ""
            else:
                self.file2_label.config(text="파일 읽기 실패", fg="#e74c3c")
                self.file2_content = ""
    
    def update_compare_button(self):
        if self.file1_content and self.file2_content:
            self.compare_button.config(state=tk.NORMAL)
            self.status_var.set("두 파일이 준비되었습니다. 비교 버튼을 클릭하세요.")
        else:
            self.compare_button.config(state=tk.DISABLED)
    
    def compare_files(self):
        if not self.file1_content or not self.file2_content:
            messagebox.showwarning("경고", "두 파일을 모두 선택해주세요.")
            return
        
        self.status_var.set("파일을 비교하고 있습니다...")
        self.root.update()
        
        # 결과 텍스트 초기화
        self.result_text.delete(1.0, tk.END)
        self.diff_data = []  # 비교 데이터 초기화
        
        try:
            # 비교 방식에 따른 처리
            if self.comparison_method.get() == "paragraph":
                self.compare_by_paragraphs()
            else:
                self.compare_by_words()
            
            # 범례와 저장 버튼 활성화
            self.legend_frame.pack(before=self.result_text.master, pady=(0, 10))
            self.save_button.config(state=tk.NORMAL)
            self.status_var.set("✅ 비교 완료! 이제 워드 파일로 저장할 수 있습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"파일 비교 중 오류가 발생했습니다: {str(e)}")
            self.status_var.set("❌ 파일 비교 중 오류가 발생했습니다.")
    
    def compare_by_paragraphs(self):
        """문단 단위로 비교합니다."""
        lines1 = [line.strip() for line in self.file1_content.split('\n') if line.strip()]
        lines2 = [line.strip() for line in self.file2_content.split('\n') if line.strip()]
        
        matcher = difflib.SequenceMatcher(None, lines1, lines2)
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                for i in range(i1, i2):
                    text = lines1[i] + '\n'
                    self.result_text.insert(tk.END, text, "normal")
                    self.diff_data.append(('equal', text))
            elif tag == 'delete':
                for i in range(i1, i2):
                    text = lines1[i] + '\n'
                    self.result_text.insert(tk.END, text, "removed")
                    self.diff_data.append(('delete', text))
            elif tag == 'insert':
                for j in range(j1, j2):
                    text = lines2[j] + '\n'
                    self.result_text.insert(tk.END, text, "added")
                    self.diff_data.append(('insert', text))
            elif tag == 'replace':
                # 삭제된 줄들
                for i in range(i1, i2):
                    text = lines1[i] + '\n'
                    self.result_text.insert(tk.END, text, "removed")
                    self.diff_data.append(('delete', text))
                # 추가된 줄들
                for j in range(j1, j2):
                    text = lines2[j] + '\n'
                    self.result_text.insert(tk.END, text, "added")
                    self.diff_data.append(('insert', text))
    
    def compare_by_words(self):
        """단어 단위로 비교합니다."""
        import re

        # 단어와 공백을 모두 포함하도록 토큰화
        tokens1 = re.findall(r'\S+|\s+', self.file1_content)
        tokens2 = re.findall(r'\S+|\s+', self.file2_content)

        matcher = difflib.SequenceMatcher(None, tokens1, tokens2)

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                text = ''.join(tokens2[j1:j2])
                self.result_text.insert(tk.END, text, "normal")
                self.diff_data.append(('equal', text))
            elif tag == 'delete':
                text = ''.join(tokens1[i1:i2])
                self.result_text.insert(tk.END, text, "removed")
                self.diff_data.append(('delete', text))
            elif tag == 'insert':
                text = ''.join(tokens2[j1:j2])
                self.result_text.insert(tk.END, text, "added")
                self.diff_data.append(('insert', text))
            elif tag == 'replace':
                # 삭제된 부분
                text1 = ''.join(tokens1[i1:i2])
                self.result_text.insert(tk.END, text1, "removed")
                self.diff_data.append(('delete', text1))
                # 추가된 부분
                text2 = ''.join(tokens2[j1:j2])
                self.result_text.insert(tk.END, text2, "added")
                self.diff_data.append(('insert', text2))
    
    def save_result(self):
        if not self.diff_data:
            messagebox.showwarning("경고", "저장할 결과가 없습니다.")
            return
        
        # 저장할 파일 선택
        file_path = filedialog.asksaveasfilename(
            title="비교 결과를 워드 파일로 저장",
            defaultextension=".docx",
            filetypes=[
                ("워드 문서", "*.docx"),
                ("모든 파일", "*.*")
            ]
        )
        
        if file_path:
            try:
                self.status_var.set("📝 비교 파일을 복사하고 차이점을 표시하는 중...")
                self.root.update()
                
                # 1단계: 비교 파일을 그대로 복사
                shutil.copy2(self.file2_path, file_path)
                
                # 2단계: 복사된 파일을 열어서 차이점 표시
                doc = docx.Document(file_path)
                
                self.status_var.set("🎨 차이점을 색상으로 표시하는 중...")
                self.root.update()
                
                # 3단계: 차이점을 문서에 표시 (헤더 추가 없이)
                self.highlight_differences_in_document(doc)
                
                # 문서 저장
                doc.save(file_path)
                
                messagebox.showinfo("✅ 저장 완료!", 
                                  f"비교 결과가 성공적으로 저장되었습니다!\n\n"
                                  f"📁 저장 위치: {file_path}\n\n"
                                  f"💡 특징:\n"
                                  f"• 비교 파일의 원본 서식과 내용이 모두 유지됩니다\n"
                                  f"• 🟢 초록색 배경 = 비교파일에 추가된 부분\n"
                                  f"• 완전히 깔끔한 문서 형태로 저장됩니다")
                
                self.status_var.set(f"✅ 저장 완료: {os.path.basename(file_path)}")
                
            except Exception as e:
                messagebox.showerror("저장 오류", f"파일 저장 중 오류가 발생했습니다: {str(e)}")
    
    def highlight_differences_in_document(self, doc):
        """문서 내에서 차이점을 정밀하게 하이라이트합니다."""
        try:
            # 1단계: 문단별 세밀한 비교
            self.highlight_paragraph_differences(doc)
            
            # 2단계: 표 내용 비교
            self.highlight_table_differences(doc)
            
        except Exception as e:
            print(f"하이라이트 처리 중 오류: {e}")
    
    def highlight_paragraph_differences(self, doc):
        """문단 내 세밀한 차이점을 하이라이트합니다."""
        original_lines = self.file1_content.split('\n')
        current_lines = self.file2_content.split('\n')
        
        # 줄 단위로 정밀 비교
        matcher = difflib.SequenceMatcher(None, original_lines, current_lines)
        
        doc_para_index = 0
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                # 동일한 부분은 그대로 두고 인덱스만 증가
                doc_para_index += (i2 - i1)
                
            elif tag == 'insert':
                # 추가된 줄들을 찾아서 하이라이트
                for j in range(j1, j2):
                    added_line = current_lines[j].strip()
                    if added_line:
                        self.highlight_matching_paragraphs(doc, added_line, 'added')
                        
            elif tag == 'delete':
                # 삭제된 내용은 현재 문서에 없으므로 주변 문단에 표시
                for i in range(i1, i2):
                    deleted_line = original_lines[i].strip()
                    if deleted_line:
                        self.add_deletion_marker(doc, deleted_line, doc_para_index)
                        
            elif tag == 'replace':
                # 변경된 부분을 세밀하게 비교
                for j in range(j1, j2):
                    if j < len(current_lines):
                        changed_line = current_lines[j].strip()
                        if changed_line:
                            # 해당 줄이 완전히 새로운 내용인지 부분 변경인지 확인
                            original_part = original_lines[i1:i2] if i1 < len(original_lines) else []
                            is_partial_change = self.is_partial_change(changed_line, original_part)
                            
                            if is_partial_change:
                                self.highlight_word_level_changes(doc, changed_line, original_part)
                            else:
                                self.highlight_matching_paragraphs(doc, changed_line, 'added')
    
    def highlight_matching_paragraphs(self, doc, target_text, highlight_type):
        """특정 텍스트와 일치하는 문단을 하이라이트합니다."""
        for paragraph in doc.paragraphs:
            if target_text in paragraph.text or self.text_similarity(paragraph.text.strip(), target_text) > 0.8:
                for run in paragraph.runs:
                    if run.text.strip():
                        if highlight_type == 'added':
                            run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN
                            run.font.color.rgb = RGBColor(46, 125, 50)
                        elif highlight_type == 'modified':
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                            run.font.color.rgb = RGBColor(245, 124, 0)
    
    def highlight_word_level_changes(self, doc, changed_line, original_parts):
        """단어 단위로 변경사항을 하이라이트합니다."""
        # 가장 유사한 원본 줄 찾기
        best_match = ""
        max_similarity = 0
        
        for orig in original_parts:
            similarity = self.text_similarity(changed_line, orig.strip())
            if similarity > max_similarity:
                max_similarity = similarity
                best_match = orig.strip()
        
        if max_similarity > 0.5:  # 50% 이상 유사하면 부분 변경으로 간주
            # 해당 문단을 찾아서 단어 단위로 하이라이트
            for paragraph in doc.paragraphs:
                if self.text_similarity(paragraph.text.strip(), changed_line) > 0.7:
                    # 단어 단위 차이점 표시
                    self.apply_word_level_highlight(paragraph, best_match, changed_line)
    
    def apply_word_level_highlight(self, paragraph, original_text, changed_text):
        """문단 내에서 단어 단위로 하이라이트를 적용합니다."""
        # 단어 단위로 비교
        original_words = original_text.split()
        changed_words = changed_text.split()
        
        matcher = difflib.SequenceMatcher(None, original_words, changed_words)
        
        # 현재 문단의 텍스트와 변경된 텍스트가 유사한 경우에만 적용
        if self.text_similarity(paragraph.text.strip(), changed_text) > 0.6:
            for run in paragraph.runs:
                if run.text.strip():
                    # 변경된 부분이 포함된 run에 하이라이트 적용
                    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                        if tag in ['insert', 'replace']:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                            run.font.color.rgb = RGBColor(245, 124, 0)
                            break
    
    def add_deletion_marker(self, doc, deleted_text, para_index):
        """삭제된 내용을 표시합니다."""
        if para_index < len(doc.paragraphs):
            target_para = doc.paragraphs[para_index]
            # 삭제된 내용을 주석으로 추가 (선택적)
            deleted_run = target_para.add_run(f" [삭제됨: {deleted_text[:50]}...]")
            deleted_run.font.highlight_color = WD_COLOR_INDEX.PINK
            deleted_run.font.color.rgb = RGBColor(198, 40, 40)
            deleted_run.font.strike = True
            deleted_run.font.size = Pt(8)
    
    def highlight_table_differences(self, doc):
        """표 내용의 차이점을 하이라이트합니다."""
        # 원본과 현재 문서의 표 내용 추출
        original_tables = self.extract_table_content(self.file1_path)
        current_tables = self.extract_table_content(self.file2_path)
        
        # 표 개수가 다른 경우 처리
        if len(current_tables) != len(original_tables):
            # 표가 추가되거나 삭제된 경우 전체 표를 하이라이트
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN
        
        else:
            # 표별로 셀 내용 비교
            for table_idx, table in enumerate(doc.tables):
                if table_idx < len(original_tables) and table_idx < len(current_tables):
                    orig_table = original_tables[table_idx]
                    curr_table = current_tables[table_idx]
                    
                    # 셀별 비교
                    for row_idx, row in enumerate(table.rows):
                        for cell_idx, cell in enumerate(row.cells):
                            if (row_idx < len(curr_table) and cell_idx < len(curr_table[row_idx]) and
                                row_idx < len(orig_table) and cell_idx < len(orig_table[row_idx])):
                                
                                current_cell_text = curr_table[row_idx][cell_idx]
                                original_cell_text = orig_table[row_idx][cell_idx]
                                
                                # 셀 내용이 다르면 하이라이트
                                if current_cell_text != original_cell_text:
                                    for paragraph in cell.paragraphs:
                                        for run in paragraph.runs:
                                            if original_cell_text and current_cell_text:
                                                # 완전히 다르면 추가로 표시
                                                if self.text_similarity(current_cell_text, original_cell_text) < 0.5:
                                                    run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN
                                                else:
                                                    # 부분 변경이면 노란색으로 표시
                                                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    
    def extract_table_content(self, file_path):
        """워드 파일에서 표 내용을 추출합니다."""
        try:
            doc = docx.Document(file_path)
            tables_content = []
            
            for table in doc.tables:
                table_data = []
                for row in table.rows:
                    row_data = []
                    for cell in row.cells:
                        row_data.append(cell.text.strip())
                    table_data.append(row_data)
                tables_content.append(table_data)
            
            return tables_content
        except:
            return []
    
    def text_similarity(self, text1, text2):
        """두 텍스트의 유사도를 계산합니다."""
        if not text1 or not text2:
            return 0.0
        
        matcher = difflib.SequenceMatcher(None, text1.lower(), text2.lower())
        return matcher.ratio()
    
    def is_partial_change(self, changed_line, original_parts):
        """부분 변경인지 완전 새로운 내용인지 판단합니다."""
        if not original_parts:
            return False
            
        for orig in original_parts:
            if self.text_similarity(changed_line, orig.strip()) > 0.4:
                return True
        return False

def main():
    # 필요한 라이브러리 확인
    try:
        import docx
        import difflib
        import shutil
    except ImportError as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("라이브러리 오류", 
                           "필요한 라이브러리가 설치되지 않았습니다.\n"
                           "다음 명령어를 실행해주세요:\n\n"
                           "pip install python-docx")
        root.destroy()
        return
    
    root = tk.Tk()
    app = WordComparatorGUI(root)
    root.mainloop()
if __name__ == "__main__":
    main()
