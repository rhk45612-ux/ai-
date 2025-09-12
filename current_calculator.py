import tkinter as tk
from tkinter import ttk, messagebox
import math
from datetime import datetime
import json
import os


class CurrentCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("⚡ 1차 전류 계산기")
        self.root.geometry("600x700")
        self.root.resizable(True, True)

        # 계산 기록을 저장할 리스트
        self.history = []
        self.history_file = "calculation_history.json"

        # 기록 파일 로드
        self.load_history()

        # GUI 설정
        self.setup_gui()

        # 프로그램 종료 시 기록 저장
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_gui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 제목
        title_label = ttk.Label(
            main_frame,
            text="⚡ 1차 전류 계산기",
            font=("Arial", 18, "bold"),
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # 입력 필드들
        # 용량 입력
        ttk.Label(
            main_frame,
            text="용량 (kVA):",
            font=("Arial", 12),
        ).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.capacity_var = tk.StringVar(value="1000")
        self.capacity_entry = ttk.Entry(
            main_frame,
            textvariable=self.capacity_var,
            font=("Arial", 12),
            width=20,
        )
        self.capacity_entry.grid(
            row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0)
        )

        # 전압 입력
        ttk.Label(
            main_frame,
            text="전압 (kV):",
            font=("Arial", 12),
        ).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.voltage_var = tk.StringVar(value="6.6")
        self.voltage_entry = ttk.Entry(
            main_frame,
            textvariable=self.voltage_var,
            font=("Arial", 12),
            width=20,
        )
        self.voltage_entry.grid(
            row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0)
        )

        # 계산 버튼
        calculate_btn = ttk.Button(
            main_frame,
            text="전류 계산하기",
            command=self.calculate_current,
            style="Accent.TButton",
        )
        calculate_btn.grid(
            row=3, column=0, columnspan=2, pady=20, sticky=(tk.W, tk.E)
        )

        # 결과 표시 프레임
        result_frame = ttk.LabelFrame(main_frame, text="계산 결과", padding="15")
        result_frame.grid(
            row=4,
            column=0,
            columnspan=2,
            sticky=(tk.W, tk.E, tk.N, tk.S),
            pady=10,
        )

        self.result_label = ttk.Label(
            result_frame,
            text="계산 버튼을 눌러주세요",
            font=("Arial", 14, "bold"),
            foreground="blue",
        )
        self.result_label.grid(row=0, column=0, pady=10)

        self.formula_label = ttk.Label(
            result_frame,
            text="공식: I₁ = (S × 1000) / (√3 × V × 1000) × 1.25",
            font=("Arial", 10),
            foreground="gray",
        )
        self.formula_label.grid(row=1, column=0, pady=(5, 0))

        # 상세 계산 과정
        self.detail_label = ttk.Label(
            result_frame, text="", font=("Arial", 10), foreground="darkgreen"
        )
        self.detail_label.grid(row=2, column=0, pady=(10, 0))

        # 공식 설명 프레임
        formula_frame = ttk.LabelFrame(main_frame, text="📋 공식 설명", padding="15")
        formula_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        formula_text = """I₁ = (용량 × 1000) / (1.732 × 전압 × 1000) × 1.25

• I₁: 1차 전류 (A)
• 용량: kVA 단위
• 전압: kV 단위  
• 1.732: √3 (3상 계산)
• 1.25: 안전율"""

        ttk.Label(
            formula_frame,
            text=formula_text,
            font=("Courier", 9),
            justify=tk.LEFT,
        ).grid(row=0, column=0, sticky=tk.W)

        # 계산 기록 프레임
        history_frame = ttk.LabelFrame(main_frame, text="📊 계산 기록", padding="10")
        history_frame.grid(
            row=6,
            column=0,
            columnspan=2,
            sticky=(tk.W, tk.E, tk.N, tk.S),
            pady=10,
        )

        # 기록 리스트박스와 스크롤바
        history_list_frame = ttk.Frame(history_frame)
        history_list_frame.grid(
            row=0, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S)
        )

        scrollbar = ttk.Scrollbar(history_list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.history_listbox = tk.Listbox(
            history_list_frame,
            yscrollcommand=scrollbar.set,
            height=8,
            font=("Arial", 9),
        )
        self.history_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.history_listbox.yview)

        # 기록 관리 버튼들
        btn_frame = ttk.Frame(history_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))

        ttk.Button(btn_frame, text="기록 삭제", command=self.clear_history).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        ttk.Button(btn_frame, text="파일로 저장", command=self.export_history).pack(
            side=tk.LEFT
        )

        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)
        history_frame.columnconfigure(0, weight=1)
        history_frame.rowconfigure(0, weight=1)

        # Enter 키 바인딩
        self.root.bind("<Return>", lambda event: self.calculate_current())

        # 기록 업데이트
        self.update_history_display()

    def calculate_current(self):
        try:
            # 입력값 가져오기
            capacity = float(self.capacity_var.get())
            voltage = float(self.voltage_var.get())

            # 입력 검증
            if capacity <= 0 or voltage <= 0:
                raise ValueError("양수 값을 입력해주세요.")

            # 1차 전류 계산
            current = (capacity * 1000) / (1.732 * voltage * 1000) * 1.25

            # 결과 표시
            self.result_label.config(
                text=f"1차 전류: {current:.2f} A", foreground="blue"
            )

            # 상세 계산 과정
            detail_text = (
                f"계산: ({capacity} × 1000) / (1.732 × {voltage} × 1000) × 1.25 = {current:.2f} A"
            )
            self.detail_label.config(text=detail_text)

            # 기록 추가
            calculation = {
                "capacity": capacity,
                "voltage": voltage,
                "current": round(current, 2),
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

            self.history.insert(0, calculation)
            if len(self.history) > 50:  # 최대 50개 기록 유지
                self.history = self.history[:50]

            self.update_history_display()
            self.save_history()

        except ValueError as e:
            messagebox.showerror(
                "입력 오류", f"올바른 숫자를 입력해주세요.\n{str(e)}"
            )
        except Exception as e:
            messagebox.showerror(
                "계산 오류", f"계산 중 오류가 발생했습니다:\n{str(e)}"
            )

    def update_history_display(self):
        self.history_listbox.delete(0, tk.END)
        for calc in self.history:
            display_text = (
                f"{calc['timestamp']} | {calc['capacity']} kVA, {calc['voltage']} kV → {calc['current']} A"
            )
            self.history_listbox.insert(tk.END, display_text)

    def clear_history(self):
        if messagebox.askyesno("확인", "모든 계산 기록을 삭제하시겠습니까?"):
            self.history.clear()
            self.update_history_display()
            self.save_history()

    def export_history(self):
        if not self.history:
            messagebox.showinfo("알림", "저장할 기록이 없습니다.")
            return

        try:
            filename = f"전류계산기록_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(filename, "w", encoding="utf-8") as f:
                f.write("=== 1차 전류 계산 기록 ===\n\n")
                for calc in self.history:
                    f.write(f"시간: {calc['timestamp']}\n")
                    f.write(f"용량: {calc['capacity']} kVA\n")
                    f.write(f"전압: {calc['voltage']} kV\n")
                    f.write(f"1차전류: {calc['current']} A\n")
                    f.write("-" * 40 + "\n")

            messagebox.showinfo(
                "저장 완료", f"기록이 '{filename}' 파일로 저장되었습니다."
            )
        except Exception as e:
            messagebox.showerror("저장 오류", f"파일 저장 중 오류가 발생했습니다:\n{str(e)}")

    def load_history(self):
        try:
            if os.path.exists(self.history_file):
                with open(self.history_file, "r", encoding="utf-8") as f:
                    self.history = json.load(f)
        except Exception:
            self.history = []

    def save_history(self):
        try:
            with open(self.history_file, "w", encoding="utf-8") as f:
                json.dump(self.history, f, ensure_ascii=False, indent=2)
        except Exception:
            pass  # 저장 실패해도 프로그램 동작에는 문제없음

    def on_closing(self):
        self.save_history()
        self.root.destroy()


def main():
    root = tk.Tk()
    app = CurrentCalculator(root)
    root.mainloop()


if __name__ == "__main__":
    main()
