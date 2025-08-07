import tkinter as tk
from tkinter import messagebox, ttk
from tkcalendar import Calendar
from datetime import datetime, date
from openpyxl import Workbook


class WorkScheduleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("근무표 생성기")

        self.users = []
        self.selected_user_index = None

        # === 대상 월 선택 ===
        tk.Label(root, text="대상 년월").grid(row=0, column=0)
        today = date.today()
        next_month = today.month + 1
        next_year = today.year
        if next_month > 12:
            next_month = 1
            next_year += 1

        self.year_var = tk.StringVar(value=str(next_year))
        self.year_cb = ttk.Combobox(
            root,
            textvariable=self.year_var,
            values=[str(y) for y in range(2020, 2031)],
            width=6,
            state="readonly",
        )
        self.year_cb.grid(row=0, column=1)

        self.month_var = tk.StringVar(value=str(next_month).zfill(2))
        self.month_cb = ttk.Combobox(
            root,
            textvariable=self.month_var,
            values=[str(m).zfill(2) for m in range(1, 13)],
            width=4,
            state="readonly",
        )
        self.month_cb.grid(row=0, column=2)

        # === 사용자 입력 ===
        tk.Label(root, text="이름").grid(row=1, column=0)
        self.name_entry = tk.Entry(root)
        self.name_entry.grid(row=1, column=1)

        # === 역할 선택 ===
        tk.Label(root, text="역할").grid(row=2, column=0)
        self.role_var = tk.StringVar(value="운전자")
        tk.Radiobutton(
            root, text="운전자", variable=self.role_var, value="운전자"
        ).grid(row=2, column=1, sticky="w")
        tk.Radiobutton(
            root, text="보조자", variable=self.role_var, value="보조자"
        ).grid(row=2, column=1, sticky="e")

        # === 코스 선택 ===
        tk.Label(root, text="가능 코스").grid(row=3, column=0)
        self.course_vars = [tk.BooleanVar(), tk.BooleanVar()]
        tk.Checkbutton(root, text="1코스", variable=self.course_vars[0]).grid(
            row=3, column=1, sticky="w"
        )
        tk.Checkbutton(root, text="2코스", variable=self.course_vars[1]).grid(
            row=3, column=1, sticky="e"
        )

        # === 휴가 날짜 선택 ===
        tk.Label(root, text="휴가 날짜 선택").grid(row=4, column=0, columnspan=2)
        self.vacation_calendar = Calendar(
            root,
            selectmode="day",
            date_pattern="yyyy-mm-dd",
            font=("Arial", 12),
            showweeknumbers=False,
            foreground="black",
            weekendforeground="black",
            headersforeground="black",
            firstweekday="sunday",
        )
        self.vacation_calendar.grid(row=5, column=0, columnspan=2)

        self.vacation_days = []
        self.vacation_list = tk.Listbox(root, width=30, height=4)
        self.vacation_list.grid(row=6, column=0, columnspan=2)

        tk.Button(root, text="휴가 날짜 추가", command=self.add_vacation_date).grid(
            row=7, column=0
        )
        tk.Button(root, text="선택 날짜 삭제", command=self.remove_vacation_date).grid(
            row=7, column=1
        )

        # === 근무일 수 ===
        tk.Label(root, text="근무 일수").grid(row=8, column=0)
        self.workdays_spinbox = tk.Spinbox(root, from_=0, to=31, width=5)
        self.workdays_spinbox.grid(row=8, column=1)

        # === 지정 근무 요일 ===
        tk.Label(root, text="지정 근무 요일").grid(row=9, column=0, sticky="w")

        # Frame 추가로 레이아웃 정리
        weekday_frame = tk.Frame(root)
        weekday_frame.grid(row=9, column=1, columnspan=5, sticky="w")

        self.weekday_vars = [tk.BooleanVar() for _ in range(5)]  # 월~금
        weekday_names = ["월", "화", "수", "목", "금"]
        for i, name in enumerate(weekday_names):
            tk.Checkbutton(
                weekday_frame, text=name, variable=self.weekday_vars[i]
            ).grid(row=0, column=i, padx=5)

        # === 사용자 추가 / 수정 ===
        tk.Button(root, text="사용자 추가", command=self.add_user).grid(
            row=10, column=0, pady=5
        )
        tk.Button(root, text="수정 준비", command=self.prepare_edit_user).grid(
            row=10, column=1, pady=5
        )

        # === 사용자 목록 ===
        tk.Label(root, text="입력된 사용자 목록").grid(row=0, column=6)
        self.user_listbox = tk.Listbox(root, width=50, height=25)
        self.user_listbox.grid(row=1, column=6, rowspan=9, padx=10)

        # === 확인 / 저장 ===
        tk.Button(
            root, text="입력 확인 및 엑셀 생성", command=self.show_users_and_save
        ).grid(row=10, column=6, pady=10)

    def add_vacation_date(self):
        date = self.vacation_calendar.get_date()
        if date not in self.vacation_days:
            self.vacation_days.append(date)
            self.vacation_days.sort(key=lambda d: datetime.strptime(d, "%Y-%m-%d"))
            self.vacation_list.delete(0, tk.END)
            for d in self.vacation_days:
                self.vacation_list.insert(tk.END, f"{d} ({self.get_weekday(d)})")

    def get_weekday(self, date_str):
        try:
            date_obj = datetime.strptime(date_str, "%Y-%m-%d")
            return ["월", "화", "수", "목", "금", "토", "일"][date_obj.weekday()]
        except:
            return ""

    def remove_vacation_date(self):
        selection = self.vacation_list.curselection()
        if selection:
            index = selection[0]
            self.vacation_list.delete(index)
            del self.vacation_days[index]

    def add_user(self):
        name = self.name_entry.get()
        role = self.role_var.get()
        courses = [i + 1 for i, var in enumerate(self.course_vars) if var.get()]
        vacations = self.vacation_days.copy()
        workdays = int(self.workdays_spinbox.get())
        weekdays = [i for i, var in enumerate(self.weekday_vars) if var.get()]
        target_month = f"{self.year_var.get()}-{self.month_var.get()}"

        if not name:
            messagebox.showwarning("입력 오류", "이름은 필수입니다.")
            return

        user = {
            "이름": name,
            "역할": role,
            "가능 코스": courses,
            "휴가": vacations,
            "근무일수": workdays,
            "지정요일": weekdays,
            "대상 월": target_month,
        }

        if self.selected_user_index is not None:
            self.users[self.selected_user_index] = user
            self.selected_user_index = None
        else:
            self.users.append(user)

        self.refresh_user_list()
        self.reset_form()

    def prepare_edit_user(self):
        selection = self.user_listbox.curselection()
        if not selection:
            return
        index = selection[0]
        user = self.users[index]
        self.selected_user_index = index

        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, user["이름"])
        self.role_var.set(user["역할"])
        for i in range(2):
            self.course_vars[i].set(i + 1 in user["가능 코스"])

        self.vacation_days = user["휴가"]
        self.vacation_list.delete(0, tk.END)
        for v in self.vacation_days:
            self.vacation_list.insert(tk.END, f"{v} ({self.get_weekday(v)})")

        self.workdays_spinbox.delete(0, tk.END)
        self.workdays_spinbox.insert(0, str(user["근무일수"]))

        for i in range(5):
            self.weekday_vars[i].set(i in user["지정요일"])

        year, month = user["대상 월"].split("-")
        self.year_var.set(year)
        self.month_var.set(month)

    def refresh_user_list(self):
        self.user_listbox.delete(0, tk.END)
        for idx, user in enumerate(self.users, start=1):
            self.user_listbox.insert(
                tk.END, f"{idx}. {user['이름']} ({user['대상 월']})"
            )

    def reset_form(self):
        self.name_entry.delete(0, tk.END)
        self.role_var.set("운전자")
        for var in self.course_vars:
            var.set(False)
        self.vacation_days.clear()
        self.vacation_list.delete(0, tk.END)
        self.workdays_spinbox.delete(0, tk.END)
        self.workdays_spinbox.insert(0, "0")
        for var in self.weekday_vars:
            var.set(False)

    def show_users_and_save(self):
        if not self.users:
            messagebox.showinfo("알림", "입력된 사용자가 없습니다.")
            return

        popup = tk.Toplevel(self.root)
        popup.title("사용자 정보 확인")
        popup.grab_set()
        popup.attributes("-topmost", True)

        text = tk.Text(popup, width=80, height=30, state="normal")
        text.pack()

        for idx, user in enumerate(self.users, start=1):
            text.insert(tk.END, f"[{idx}] {user['이름']}\n")
            text.insert(tk.END, f"  역할: {user['역할']}\n")
            text.insert(tk.END, f"  가능 코스: {user['가능 코스']}\n")
            text.insert(tk.END, f"  휴가 날짜: {', '.join(user['휴가'])}\n")
            text.insert(tk.END, f"  근무일수: {user['근무일수']}\n")
            text.insert(
                tk.END,
                f"  지정 요일: {[['월', '화', '수', '목', '금'][i] for i in user['지정요일']]}\n",
            )
            text.insert(tk.END, f"  대상 월: {user['대상 월']}\n\n")

        text.config(state="disabled")
        tk.Button(
            popup, text="엑셀 생성", command=lambda: self.save_to_excel(popup)
        ).pack(pady=10)

    def save_to_excel(self, popup_window=None):
        wb = Workbook()
        ws = wb.active
        ws.title = "입력정보"
        ws.append(
            [
                "이름",
                "역할",
                "가능 코스",
                "휴가 날짜",
                "근무일수",
                "지정 요일",
                "대상 월",
            ]
        )

        for user in self.users:
            ws.append(
                [
                    user["이름"],
                    user["역할"],
                    ",".join(map(str, user["가능 코스"])),
                    ", ".join(user["휴가"]),
                    user["근무일수"],
                    ",".join(
                        [["월", "화", "수", "목", "금"][i] for i in user["지정요일"]]
                    ),
                    user["대상 월"],
                ]
            )

        wb.save("근무표_입력정보.xlsx")
        messagebox.showinfo("완료", "엑셀 파일이 저장되었습니다.")
        if popup_window:
            popup_window.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = WorkScheduleApp(root)
    root.mainloop()
