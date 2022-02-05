from tkinter import *
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
import json
import datetime
import os
from datetime import date
import sys
import getpass

class WorkManagement:
    def __init__(self, root):
        self.root = root
        self.root.geometry("1530x790+0+0")
        self.root.title("QUẢN LÝ CÔNG VIỆC")

        self.var_STT = StringVar()
        self.var_profile_job = ""
        self.var_content_job = ""
        self.var_time = ""
        self.var_doc_job = ""
        self.var_storage = ""
        self.var_note = ""
        self.var_state = StringVar()
        self.var_loop_time = StringVar()
        self.var_type_work = StringVar()

        # Title
        label_title = Label(
            self.root,
            text="QUẢN LÝ CÔNG VIỆC",
            font=("Arial", 37, "bold"),
            fg="darkblue",
            bg="white",
        )
        label_title.place(x=0, y=0, width=1530, height=70)

        # Main Frame
        main_frame = Frame(
            self.root, bd=2, relief=RIDGE, bg="white", width=1500, height=660
        )
        main_frame.place(x=10, y=100)

        # upper frame
        upper_frame = LabelFrame(
            main_frame,
            bd=2,
            relief=RIDGE,
            bg="white",
            text="Thông tin công việc",
            font=("Arial", 11, "bold"),
            fg="red",
        )
        upper_frame.place(x=10, y=10, width=1480, height=370)

        # uppercanvas
        upper_canvas = Canvas(upper_frame, bg="white")
        upper_canvas.pack(side=LEFT, fill=BOTH, expand=1)

        # Scroll
        upper_frame_scroll = Scrollbar(upper_frame, command=upper_canvas.yview)
        upper_frame_scroll.pack(side=RIGHT, fill=Y)

        upper_canvas.configure(yscrollcommand=upper_frame_scroll.set)
        upper_canvas.bind(
            "<Configure>",
            lambda e: upper_canvas.configure(scrollregion=upper_canvas.bbox("all")),
        )

        # scrollable_frame
        scrollable_frame = Frame(upper_canvas, bg="white")
        upper_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        # down frame
        down_frame = LabelFrame(
            main_frame,
            bd=2,
            relief=RIDGE,
            bg="white",
            text="Danh sách công việc",
            font=("Arial", 11, "bold"),
            fg="red",
        )
        down_frame.place(x=10, y=380, width=1480, height=270)

        # STT
        stt_label = Label(
            scrollable_frame, text="ID:", font=("Arial", 12, "bold"), bg="white"
        )
        stt_label.grid(row=0, column=0, sticky=W, padx=2, pady=7)

        self.stt_entry = Entry(
            scrollable_frame,
            state=DISABLED,
            textvariable=self.var_STT,
            width=33,
            font=("Arial", 12, "bold"),
            bd=2,
        )
        self.stt_entry.grid(row=0, column=1, sticky=W, padx=2, pady=7)

        # Hồ sơ công việc:
        name_label = Label(
            scrollable_frame,
            text="Hồ sơ công việc:",
            font=("Arial", 12, "bold"),
            bg="white",
        )
        name_label.grid(row=1, column=0, sticky=W, padx=2, pady=7)

        self.name_entry = Text(
            scrollable_frame, width=33, height=4, font=("Arial", 12, "bold"), bd=2
        )
        self.name_entry.grid(row=1, column=1, sticky=W, padx=2, pady=7)
        self.name_entry.bind(
            "<KeyRelease>",
            lambda event: self.update_value_profile(
                event, self.name_entry.get(1.0, END + "-1c")
            ),
        )

        # Nội dung công việc
        content_label = Label(
            scrollable_frame,
            text="Nội dung công việc:",
            font=("Arial", 12, "bold"),
            bg="white",
        )
        content_label.grid(row=2, column=0, sticky=W, padx=2, pady=7)

        self.content_entry = Text(
            scrollable_frame, width=33, height=4, font=("Arial", 12, "bold"), bd=2
        )
        self.content_entry.grid(row=2, column=1, sticky=W, padx=2, pady=7)
        self.content_entry.bind(
            "<KeyRelease>",
            lambda event: self.update_value_content(
                event, self.content_entry.get(1.0, END + "-1c")
            ),
        )

        # Thòi gian thực hiện
        time_label = Label(
            scrollable_frame,
            text="Thời gian thực hiện:",
            font=("Arial", 12, "bold"),
            bg="white",
        )
        time_label.grid(row=0, column=2, sticky=W, padx=2, pady=7)

        self.date_entry = DateEntry(
            scrollable_frame, selectmode="day", width=33, date_pattern="dd/mm/y"
        )
        self.date_entry.grid(row=0, column=3, sticky=W, padx=2, pady=7)

        # So lan lap
        loop_label = Label(
            scrollable_frame, text="Lặp lại", font=("Arial", 12, "bold"), bg="white"
        )
        loop_label.grid(row=0, column=4, sticky=W, padx=2, pady=7)
        loop_label_entry = ttk.Combobox(
            scrollable_frame,
            textvariable=self.var_loop_time,
            width=33,
            font=("Arial", 12, "bold"),
            state="readonly",
        )
        loop_label_entry["value"] = ("Một lần", "Mỗi ngày", "Mỗi tuần", "Mỗi tháng", "Mỗi quý", "Mỗi năm")
        loop_label_entry.current(0)
        loop_label_entry.grid(row=0, column=5, sticky=W, padx=2, pady=7)

        # Căn cứ thực hiện
        doc_label = Label(
            scrollable_frame,
            text="Căn cứ thực hiện:",
            font=("Arial", 12, "bold"),
            bg="white",
        )
        doc_label.grid(row=1, column=2, sticky=W, padx=2, pady=7)

        self.doc_entry = Text(
            scrollable_frame, width=33, height=4, font=("Arial", 12, "bold"), bd=2
        )
        self.doc_entry.grid(row=1, column=3, sticky=W, padx=2, pady=7)
        self.doc_entry.bind(
            "<KeyRelease>",
            lambda event: self.update_value_doc(
                event, self.doc_entry.get(1.0, END + "-1c")
            ),
        )
        doc_label.bind("<ButtonRelease>", lambda event: self.set_link_doc(event))
        self.doc_entry.bind(
            "<Double-Button-1>",
            lambda event: self.open_file(
                event, self.doc_entry.get(1.0, END).replace("\n", "")
            ),
        )

        # Lưu trữ hồ sơ
        storage_label = Label(
            scrollable_frame,
            text="Lưu trữ Hồ Sơ:",
            font=("Arial", 12, "bold"),
            bg="white",
        )
        storage_label.grid(row=2, column=2, sticky=W, padx=2, pady=7)

        self.storage_entry = Text(
            scrollable_frame, width=33, height=4, font=("Arial", 12, "bold"), bd=2
        )
        self.storage_entry.grid(row=2, column=3, sticky=W, padx=2, pady=7)
        self.storage_entry.bind(
            "<KeyRelease>",
            lambda event: self.update_value_storage(
                event, self.storage_entry.get(1.0, END + "-1c")
            ),
        )
        storage_label.bind(
            "<ButtonRelease>", lambda event: self.set_link_storage(event)
        )
        self.storage_entry.bind(
            "<Double-Button-1>",
            lambda event: self.open_folder(
                event, self.storage_entry.get(1.0, END).replace("\n", "")
            ),
        )

        # Ghi chú
        note_label = Label(
            scrollable_frame, text="Ghi chú:", font=("Arial", 12, "bold"), bg="white"
        )
        note_label.grid(row=1, column=4, sticky=W, padx=2, pady=7)

        self.note_entry = Text(
            scrollable_frame, width=28, height=4, font=("Arial", 12, "bold"), bd=2
        )
        self.note_entry.grid(row=1, column=5, sticky=W, padx=2, pady=7)
        self.note_entry.bind(
            "<KeyRelease>",
            lambda event: self.update_value_note(
                event, self.note_entry.get(1.0, END + "-1c")
            ),
        )

        # State
        state_label = Label(
            scrollable_frame, text="Trạng thái:", font=("Arial", 12, "bold"), bg="white"
        )
        state_label.grid(row=2, column=4, sticky=W, padx=2, pady=7)

        state_entry = ttk.Combobox(
            scrollable_frame,
            textvariable=self.var_state,
            width=22,
            font=("Arial", 12, "bold"),
            state="readonly",
        )
        state_entry["value"] = ("Chưa hoàn thành", "Đã hoàn thành")
        state_entry.current(0)
        state_entry.grid(row=2, column=5, sticky=W, padx=2, pady=7)

        # Bat/tat thong bao
        # nofi_label = Label(
        #     scrollable_frame,
        #     text="Bật/tắt thông báo:",
        #     font=("Arial", 12, "bold"),
        #     bg="white",
        # )
        # nofi_label.grid(row=3, column=0, sticky=W, padx=2, pady=7)

        # self.var_nofi = StringVar()
        # self.nofi_entry = ttk.Combobox(
        #     scrollable_frame,
        #     textvariable=self.var_nofi,
        #     width=33,
        #     font=("Arial", 12, "bold"),
        #     state="readonly",
        # )
        # self.nofi_entry["value"] = ("Bật", "Tắt")
        # self.nofi_entry.current(0)
        # self.nofi_entry.grid(row=3, column=1, sticky=W, padx=2, pady=7)

        # alarm_label = Label(
        #     scrollable_frame,
        #     text="Hẹn giờ thông báo:",
        #     font=("Arial", 12, "bold"),
        #     bg="white",
        # )
        # alarm_label.grid(row=3, column=2, sticky=W, padx=2, pady=7)

        # self.timer_label = Label(
        #     scrollable_frame, text="8:00 AM", font=("Arial", 12, "bold"), bg="white"
        # )
        # self.timer_label.grid(row=3, column=3, sticky=W, padx=2, pady=7)

        # self.timer_label.bind("<ButtonRelease>", lambda event: self.get_time(event))

        # Button frame
        button_frame = Frame(upper_canvas, bg="white", width=80)
        button_frame.pack(side=RIGHT)
        save_but = Button(
            button_frame,
            command=self.save_data,
            font=("Arial", 12, "bold"),
            fg="white",
            bg="blue",
            text="Lưu",
        )
        save_but.grid(row=0, column=0, sticky=NSEW, padx=2, pady=7)

        update_but = Button(
            button_frame,
            command=self.update_data,
            font=("Arial", 12, "bold"),
            fg="white",
            bg="blue",
            text="Cập nhật",
        )
        update_but.grid(row=1, column=0, sticky=NSEW, padx=2, pady=7)

        del_but = Button(
            button_frame,
            font=("Arial", 12, "bold"),
            fg="white",
            bg="blue",
            text="Xóa",
            command=self.delete_data,
        )
        del_but.grid(row=2, column=0, sticky=NSEW, padx=2, pady=7)

        clear_but = Button(
            button_frame,
            font=("Arial", 12, "bold"),
            fg="white",
            bg="blue",
            text="Làm sạch",
            command=self.clear_data,
        )

        clear_but.grid(row=3, column=0, sticky=NSEW, padx=2, pady=7)

        export_but = Button(
            button_frame,
            font=("Arial", 12, "bold"),
            fg="white",
            bg="blue",
            text="Xuất file Excel",
            command=self.export_data,
        )

        export_but.grid(row=4, column=0, sticky=NSEW, padx=2, pady=7)

        # Search frame
        search_frame = LabelFrame(
            down_frame,
            bd=2,
            relief=RIDGE,
            bg="white",
            text="Tìm kiếm công việc",
            font=("Arial", 9, "bold"),
            fg="blue",
        )
        search_frame.place(x=10, y=0, width=1460, height=60)

        # Loai cong viec
        choose_type_work_label = Label(
            search_frame,
            text="Chọn loại công việc",
            font=("Arial", 12, "bold"),
            bg="white",
        )
        choose_type_work_label.grid(row=0, column=0, sticky=W, padx=2, pady=7)

        choose_type_work = ttk.Combobox(
            search_frame,
            textvariable=self.var_type_work,
            width=33,
            font=("Arial", 12, "bold"),
            state="readonly",
        )
        choose_type_work["value"] = (
            "Chưa hoàn thành",
            "Đã hoàn thành",
            "Tất cả",
            "Trong ngày",
        )
        choose_type_work.current(0)
        choose_type_work.grid(row=0, column=1, sticky=W, padx=2, pady=7)

        show_but = Button(
            search_frame,
            font=("Arial", 12, "bold"),
            fg="white",
            bg="blue",
            text="Hiển thị",
            command=self.fetch_data,
        )
        show_but.grid(row=0, column=2, sticky=W, padx=2, pady=7)

        search_label = Label(
            search_frame,
            text="Tìm kiếm theo:",
            font=("Arial", 12, "bold"),
            bg="red",
            fg="white",
        )
        search_label.grid(row=0, column=3, padx=10, pady=7)

        self.var_type_search = StringVar()
        type_search = ttk.Combobox(
            search_frame,
            textvariable=self.var_type_search,
            width=17,
            font=("Arial", 12, "bold"),
            state="readonly",
        )
        type_search["value"] = ("Hồ sơ công việc", "Nội dung công việc")
        type_search.current(1)
        type_search.grid(row=0, column=4, padx=10, pady=7)

        self.var_search = StringVar()
        search_entry = Entry(
            search_frame,
            textvariable=self.var_search,
            width=40,
            font=("Arial", 12, "bold"),
            bd=2,
        )
        search_entry.grid(row=0, column=5, padx=10, pady=7)

        search_but = Button(
            search_frame,
            font=("Arial", 12, "bold"),
            fg="white",
            bg="blue",
            text="Tìm Kiếm",
            command=self.search_data,
        )
        search_but.grid(row=0, column=6, sticky=W, padx=10, pady=7)

        # Table frame
        table_frame = Frame(down_frame, bd=3, relief=RIDGE)
        table_frame.place(x=0, y=70, width=1470, height=170)

        scroll_x = ttk.Scrollbar(table_frame, orient=HORIZONTAL)
        scroll_y = ttk.Scrollbar(table_frame, orient=VERTICAL)

        self.work_table = ttk.Treeview(
            table_frame,
            column=(
                "id",
                "profile_job",
                "content_job",
                "time",
                "doc_job",
                "storage",
                "note",
                "state",
            ),
            xscrollcommand=scroll_x.set,
            yscrollcommand=scroll_y.set,
        )
        scroll_x.pack(side=BOTTOM, fill=X)
        scroll_y.pack(side=RIGHT, fill=Y)

        scroll_x.config(command=self.work_table.xview)
        scroll_y.config(command=self.work_table.yview)

        self.work_table.heading("id", text="ID")
        self.work_table.heading("profile_job", text="Hồ sơ công việc")
        self.work_table.heading("content_job", text="Nội dung công việc")
        self.work_table.heading("time", text="Thời gian thực hiện")
        self.work_table.heading("doc_job", text="Căn cứ thực hiện")
        self.work_table.heading("storage", text="Lưu trữ Hồ Sơ")
        self.work_table.heading("note", text="Ghi chú")
        self.work_table.heading("state", text="Trạng thái")

        self.work_table["show"] = "headings"
        self.work_table.column("id", width=50)
        self.work_table.column("profile_job", width=200)
        self.work_table.column("content_job", width=300)
        self.work_table.column("time", width=150)
        self.work_table.column("doc_job", width=200)
        self.work_table.column("storage", width=200)
        self.work_table.column("note", width=200)
        self.work_table.column("state", width=100)
        self.work_table.pack(fill=BOTH, expand=1)
        self.work_table.bind("<ButtonRelease>", self.get_cursor)

        self.fetch_data()

    def update_value_profile(self, event, value):
        self.var_profile_job = value

    def update_value_content(self, event, value):
        self.var_content_job = value

    def update_value_doc(self, event, value):
        self.var_doc_job = value

    def update_value_storage(self, event, value):
        self.var_storage = value

    def update_value_note(self, event, value):
        self.var_note = value

    def save_data(self):
        if getattr(sys, "frozen", False):
            f = open(f"{os.path.dirname(sys.executable)}/data.json")
        # or a script file (e.g. `.py` / `.pyw`)
        elif __file__:
            f = open(f"{os.path.dirname(__file__)}/data.json")
        fdata = json.load(f)
        f.close()
        if self.var_profile_job == "" or self.var_content_job == "":
            messagebox.showerror(
                "Thiếu thông tin",
                "Út phải nhập trường Hồ sơ công việc và Nội dung công việc!!!",
            )
            return
        for index, work in enumerate(fdata):
            if work["content_job"] == self.var_content_job and work["profile_job"] == self.var_profile_job:
                is_change = messagebox.askyesno(
                    "Thay đổi thông tin",
                    "Út có muốn thay đổi thông tin của công việc này không?",
                )
                if is_change:
                    work["id"] = self.var_STT.get()
                    work["profile_job"] = self.var_profile_job
                    work["content_job"] = self.var_content_job
                    work["time"] = self.date_entry.get_date().strftime("%d/%m/%Y")
                    work["doc_job"] = self.var_doc_job
                    work["storage"] = self.var_storage
                    work["note"] = self.var_note
                    work["state"] = self.var_state.get()
                    work["loop"] = self.var_loop_time.get()
                    if getattr(sys, "frozen", False):
                        with open(
                            f"{os.path.dirname(sys.executable)}/data.json", "w"
                        ) as f:
                            json.dump(fdata, f)
                    elif __file__:
                        with open(f"{os.path.dirname(__file__)}/data.json", "w") as f:
                            json.dump(fdata, f)
                    self.fetch_data()
                    self.clear_data()
                return
        else:
            try:
                save = {}
                save["id"] = len(fdata) + 1
                save["profile_job"] = self.var_profile_job
                save["content_job"] = self.var_content_job
                save["time"] = self.date_entry.get_date().strftime("%d/%m/%Y")
                save["doc_job"] = self.var_doc_job
                save["storage"] = self.var_storage
                save["note"] = self.var_note
                save["state"] = self.var_state.get()
                save["loop"] = self.var_loop_time.get()
                print(save)
                fdata.append(save)
                if getattr(sys, "frozen", False):
                    with open(f"{os.path.dirname(sys.executable)}/data.json", "w") as f:
                        json.dump(fdata, f)
                elif __file__:
                    with open(f"{os.path.dirname(__file__)}/data.json", "w") as f:
                        json.dump(fdata, f)
                self.fetch_data()
                self.clear_data()
                messagebox.showinfo("success", "Đã thêm công việc thành công!!!")
            except Exception as ex:
                messagebox.showerror("Error", f"{str(ex)}")

    def fetch_data(self):
        self.work_table.delete(*self.work_table.get_children())
        if getattr(sys, "frozen", False):
            f = open(f"{os.path.dirname(sys.executable)}/data.json")
        # or a script file (e.g. `.py` / `.pyw`)
        elif __file__:
            f = open(f"{os.path.dirname(__file__)}/data.json")
        fdata = json.load(f)
        f.close()
        if getattr(sys, "frozen", False):
            f = open(f"{os.path.dirname(sys.executable)}/alarm.json")
        # or a script file (e.g. `.py` / `.pyw`)
        elif __file__:
            f = open(f"{os.path.dirname(__file__)}/alarm.json")
        alarm = json.load(f)
        f.close()
        print(alarm)
        fdata = sorted(
            fdata, key=lambda x: datetime.datetime.strptime(x["time"], "%d/%m/%Y")
        )
        if len(fdata) != 0:
            self.work_table.delete(*self.work_table.get_children())
            if self.var_type_work.get() == "Tất cả":
                for index,work in enumerate(fdata):
                    self.work_table.insert(
                        "",
                        END,
                        value=(
                            work["id"],
                            work["profile_job"],
                            work["content_job"],
                            work["time"],
                            work["doc_job"],
                            work["storage"],
                            work["note"],
                            work["state"],
                        ),
                    )
            if alarm["alarm"] == 1 or self.var_type_work.get() == "Trong ngày":
                self.var_type_work.set("Trong ngày")
                alarm["alarm"] = 0
                if getattr(sys, "frozen", False):
                    with open(f"{os.path.dirname(sys.executable)}/alarm.json", "w") as f:
                        json.dump(alarm, f)
                elif __file__:
                    with open(f"{os.path.dirname(__file__)}/alarm.json", "w") as f:
                        json.dump(alarm, f)
                date_now = str(date.today())
                date_now_arr = date_now.split("-")
                day_now = int(date_now_arr[2])
                month_now = int(date_now_arr[1])
                year_now = int(date_now_arr[0])
                weekday_now = datetime.datetime.now().strftime("%A")
                for index,work in enumerate(fdata):
                    date_data = work["time"]
                    date_data_arr = date_data.split("/")
                    day_data = int(date_data_arr[0])
                    month_data = int(date_data_arr[1])
                    year_data = int(date_data_arr[2])
                    weekday_data = datetime.datetime(
                        year_data, month_data, day_data
                    ).strftime("%A")
                    if work["loop"] == "Một lần":
                        if (
                            day_now == day_data
                            and month_now == month_data
                            and year_now == year_data
                        ):
                            self.work_table.insert(
                                "",
                                END,
                                value=(
                                    work["id"],
                                    work["profile_job"],
                                    work["content_job"],
                                    work["time"],
                                    work["doc_job"],
                                    work["storage"],
                                    work["note"],
                                    work["state"],
                                ),
                            )
                    elif work["loop"] == "Mỗi ngày":
                        self.work_table.insert(
                                "",
                                END,
                                value=(
                                    work["id"],
                                    work["profile_job"],
                                    work["content_job"],
                                    work["time"],
                                    work["doc_job"],
                                    work["storage"],
                                    work["note"],
                                    work["state"],
                                ),
                            )
                    elif work["loop"] == "Mỗi tuần":
                        if weekday_now == weekday_data:
                            self.work_table.insert(
                                "",
                                END,
                                value=(
                                    work["id"],
                                    work["profile_job"],
                                    work["content_job"],
                                    work["time"],
                                    work["doc_job"],
                                    work["storage"],
                                    work["note"],
                                    work["state"],
                                ),
                            )
                    elif work["loop"] == "Mỗi tháng":
                        if day_now == day_data:
                            self.work_table.insert(
                                "",
                                END,
                                value=(
                                    work["id"],
                                    work["profile_job"],
                                    work["content_job"],
                                    work["time"],
                                    work["doc_job"],
                                    work["storage"],
                                    work["note"],
                                    work["state"],
                                ),
                            )
                    elif work["loop"] == "Mỗi quý":
                        if day_now == date_data and (month_now - month_now) % 3 == 0:
                            self.work_table.insert(
                                "",
                                END,
                                value=(
                                    work["id"],
                                    work["profile_job"],
                                    work["content_job"],
                                    work["time"],
                                    work["doc_job"],
                                    work["storage"],
                                    work["note"],
                                    work["state"],
                                ),
                            )
                    else:
                        if day_now == day_data and month_now == month_data:
                            self.work_table.insert(
                                "",
                                END,
                                value=(
                                    work["id"],
                                    work["profile_job"],
                                    work["content_job"],
                                    work["time"],
                                    work["doc_job"],
                                    work["storage"],
                                    work["note"],
                                    work["state"],
                                ),
                            )

            else:
                for index, work in enumerate(fdata):
                    if work["state"] == self.var_type_work.get():
                        self.work_table.insert(
                            "",
                            END,
                            value=(
                                work["id"],
                                work["profile_job"],
                                work["content_job"],
                                work["time"],
                                work["doc_job"],
                                work["storage"],
                                work["note"],
                                work["state"],
                            ),
                        )

    def set_link_doc(self, event):
        filePath = filedialog.askopenfilename()
        if filePath != "":
            self.var_doc_job = filePath
            self.doc_entry.delete(1.0, END)
            self.doc_entry.insert(END, filePath)
            if getattr(sys, "frozen", False):
                f = open(f"{os.path.dirname(sys.executable)}/data.json")
            # or a script file (e.g. `.py` / `.pyw`)
            elif __file__:
                f = open(f"{os.path.dirname(__file__)}/data.json")
            fdata = json.load(f)
            f.close()
            fdata[int(self.var_STT.get()) - 1]["doc_job"] = filePath
            with open("data.json", "w") as f:
                json.dump(fdata, f)

    def set_link_storage(self, event):
        folder_path = filedialog.askdirectory()
        if folder_path != "":
            self.var_storage = folder_path
            self.storage_entry.delete(1.0, END)
            self.storage_entry.insert(END, folder_path)
            if getattr(sys, "frozen", False):
                f = open(f"{os.path.dirname(sys.executable)}/data.json")
            # or a script file (e.g. `.py` / `.pyw`)
            elif __file__:
                f = open(f"{os.path.dirname(__file__)}/data.json")
            fdata = json.load(f)
            f.close()
            fdata[int(self.var_STT.get()) - 1]["storage"] = folder_path
            with open("data.json", "w") as f:
                json.dump(fdata, f)

    def clear_data(self):
        self.var_STT.set("")
        self.var_profile_job = ""
        self.var_content_job = ""
        self.var_time = datetime.datetime.now().strftime("%d/%m/%Y")
        self.var_doc_job = ""
        self.var_storage = ""
        self.var_note = ""
        self.var_state.set("Chưa hoàn thành")
        self.var_loop_time.set("Một lần")
        self.set_value_all_text(
            self.var_profile_job,
            self.var_content_job,
            self.var_doc_job,
            self.var_storage,
            self.var_note,
        )
        self.date_entry.set_date(self.var_time)
        self.fetch_data()

    def set_value_all_text(self, profile, content, doc, storage, note):
        self.name_entry.delete(1.0, END)
        self.name_entry.insert(END, profile)
        self.content_entry.delete(1.0, END)
        self.content_entry.insert(END, content)
        self.doc_entry.delete(1.0, END)
        self.doc_entry.insert(END, doc)
        self.storage_entry.delete(1.0, END)
        self.storage_entry.insert(END, storage)
        self.note_entry.delete(1.0, END)
        self.note_entry.insert(END, note)

    def get_cursor(self, event):
        cursor_row = self.work_table.focus()
        content = self.work_table.item(cursor_row)
        data = content["values"]
        self.var_STT.set(data[0])
        self.var_profile_job = data[1]
        self.var_content_job = data[2]
        self.var_time = data[3]
        self.var_doc_job = data[4]
        self.var_storage = data[5]
        self.var_note = data[6]
        self.var_state.set(data[7])
        self.set_value_all_text(
            self.var_profile_job,
            self.var_content_job,
            self.var_doc_job,
            self.var_storage,
            self.var_note,
        )
        self.date_entry.set_date(self.var_time)
        if getattr(sys, "frozen", False):
            f = open(f"{os.path.dirname(sys.executable)}/data.json")
        # or a script file (e.g. `.py` / `.pyw`)
        elif __file__:
            f = open(f"{os.path.dirname(__file__)}/data.json")
        fdata = json.load(f)
        f.close()
        for work in fdata:
            if work["id"] == data[0]:
                # self.var_nofi.set(work["nofi"])
                # self.timer_label.config(text=work["timer"])
                self.var_loop_time.set(work["loop"])
                break

    def delete_data(self):
        if getattr(sys, "frozen", False):
            f = open(f"{os.path.dirname(sys.executable)}/data.json")
        # or a script file (e.g. `.py` / `.pyw`)
        elif __file__:
            f = open(f"{os.path.dirname(__file__)}/data.json")
        fdata = json.load(f)
        f.close()
        fdata.pop(int(self.var_STT.get()) - 1)
        for index in range(len(fdata)):
            fdata[index]["id"] = index + 1
        if getattr(sys, "frozen", False):
            with open(f"{os.path.dirname(sys.executable)}/data.json", "w") as f:
                json.dump(fdata, f)
        elif __file__:
            with open(f"{os.path.dirname(__file__)}/data.json", "w") as f:
                json.dump(fdata, f)
        self.clear_data()
        self.fetch_data()

    def search_data(self):
        if getattr(sys, "frozen", False):
            f = open(f"{os.path.dirname(sys.executable)}/data.json")
        # or a script file (e.g. `.py` / `.pyw`)
        elif __file__:
            f = open(f"{os.path.dirname(__file__)}/data.json")
        fdata = json.load(f)
        f.close()
        self.work_table.delete(*self.work_table.get_children())
        if self.var_type_search.get() == "Hồ sơ công việc":
            for work in fdata:
                if self.var_search.get() in work["profile_job"]:
                    self.work_table.insert(
                        "",
                        END,
                        value=(
                            work["id"],
                            work["profile_job"],
                            work["content_job"],
                            work["time"],
                            work["doc_job"],
                            work["storage"],
                            work["note"],
                            work["state"],
                        ),
                    )
        else:
            for work in fdata:
                if self.var_search.get() in work["content_job"]:
                    self.work_table.insert(
                        "",
                        END,
                        value=(
                            work["id"],
                            work["profile_job"],
                            work["content_job"],
                            work["time"],
                            work["doc_job"],
                            work["storage"],
                            work["note"],
                            work["state"],
                        ),
                    )

    def open_file(self, event, file_path):
        try:
            os.system(f'start "file" "{file_path}"')
        except Exception as ex:
            messagebox.showerror("Error", f"{str(ex)}")

    def open_folder(self, event, fol_path):
        try:
            os.system(f'start "dir" "{fol_path}"')
            # os.system("start .")
        except Exception as ex:
            messagebox.showerror("Error", f"{str(ex)}")

    def updateTime(self, time):
        self.timer_label.configure(text="{}:{} {}".format(*time))

    def get_time(self, event):

        top = Toplevel(root)

        time_picker = AnalogPicker(top)
        time_picker.pack(expand=True, fill="both")

        theme = AnalogThemes(time_picker)
        theme.setDracula()
        # theme.setNavyBlue()
        # theme.setPurple()
        ok_btn = Button(
            top, text="ok", command=lambda: self.updateTime(time_picker.time())
        )
        ok_btn.pack()

    def update_data(self):
        if getattr(sys, "frozen", False):
            f = open(f"{os.path.dirname(sys.executable)}/data.json")
        # or a script file (e.g. `.py` / `.pyw`)
        elif __file__:
            f = open(f"{os.path.dirname(__file__)}/data.json")
        fdata = json.load(f)
        f.close()
        for work in fdata:
            if str(work["id"]) == self.var_STT.get():
                work["profile_job"] = self.var_profile_job
                work["content_job"] = self.var_content_job
                work["time"] = self.date_entry.get_date().strftime("%d/%m/%Y")
                work["doc_job"] = self.var_doc_job
                work["storage"] = self.var_storage
                work["note"] = self.var_note
                work["state"] = self.var_state.get()
                work["loop"] = self.var_loop_time.get()
                if getattr(sys, "frozen", False):
                    with open(f"{os.path.dirname(sys.executable)}/data.json", "w") as f:
                        json.dump(fdata, f)
                elif __file__:
                    with open(f"{os.path.dirname(__file__)}/data.json", "w") as f:
                        json.dump(fdata, f)
                self.fetch_data()
                messagebox.showinfo("success", "Đã cập nhật công việc thành công!!!")
                return
    def export_data(self):
        try:
            filename = filedialog.asksaveasfilename(defaultextension = ".xlsx",
            filetypes = [("Text File", ".xlsx")])
            if getattr(sys, "frozen", False):
                f = open(f"{os.path.dirname(sys.executable)}/data.json")
            # or a script file (e.g. `.py` / `.pyw`)
            elif __file__:
                f = open(f"{os.path.dirname(__file__)}/data.json")
                fdata = json.load(f)
            f.close()

            fdata = sorted(
                fdata, key=lambda x: x["profile_job"]
            )
            export = {}
            export["ID"] = []
            export["Hồ sơ công việc"] = []
            export["Nội dung công việc"] = []
            export["Thời gian thực hiện"] = []
            export["Căn cứ thực hiện"] = []
            export["Lưu trữ hồ sơ"] = []
            export["Ghi chú"] = []
            for index,work in enumerate(fdata):
                export["ID"].append(index + 1)
                export["Hồ sơ công việc"].append(work["profile_job"])
                export["Nội dung công việc"].append(work["content_job"])
                export["Thời gian thực hiện"].append(work["time"])
                export["Căn cứ thực hiện"].append(work["doc_job"])
                export["Lưu trữ hồ sơ"].append(work["storage"])
                export["Ghi chú"].append(work["note"])
            df = pd.DataFrame(export)
            print(df)
            df.to_excel(filename, index=False)
        except Exception as ex:
            messagebox.showerror("Error", f"{ex}")

def write_bat():
    USER = getpass.getuser()
    file_bat = open(
        f"C:\\Users\\{USER}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\\workManage.bat",
        "w",
        encoding="utf-8",
    )
    # determine if the application is a frozen `.exe` (e.g. pyinstaller --onefile)
    if getattr(sys, "frozen", False):
        file_bat.write(f'start "file" "{os.path.dirname(sys.executable)}\\alarm.exe"')
    # or a script file (e.g. `.py` / `.pyw`)
    elif __file__:
        file_bat.write(f"python {os.path.dirname(__file__)}\\alarm.py")
    file_bat.close()


def close():
    if getattr(sys, "frozen", False):
        f = open(f"{os.path.dirname(sys.executable)}/alarm.json")
        # or a script file (e.g. `.py` / `.pyw`)
    elif __file__:
        f = open(f"{os.path.dirname(__file__)}/alarm.json")
    alarm = json.load(f)
    f.close()
    alarm["alarm"] = 0
    if getattr(sys, "frozen", False):
        with open(f"{os.path.dirname(sys.executable)}/alarm.json", "w") as f:
            json.dump(alarm, f)
    elif __file__:
        with open(f"{os.path.dirname(__file__)}/alarm.json", "w") as f:
            json.dump(alarm, f)
    root.destroy()


if __name__ == "__main__":
    try:
        write_bat()
        root = Tk()
        root.protocol("WM_DELETE_WINDOW", close)
        obj = WorkManagement(root)
        root.mainloop()
    except Exception as Ex:
        print(f"{Ex}")
