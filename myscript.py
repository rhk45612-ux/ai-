import tkinter as tk

# 창 생성
root = tk.Tk()
root.title("프로그램 창")
root.geometry("500x500")

# 라벨 추가
label = tk.Label(root, text="실행 가능", font=("Arial", 24))
label.pack(expand=True)

# 이벤트 루프 실행
root.mainloop()
