import os
import sys
import math
import queue
import threading
import subprocess
import pandas as pd
from moviepy import VideoFileClip
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def parse_time(time_value):
    if pd.isna(time_value):
        return None

    time_str = str(time_value).strip()
    if not time_str:
        return None

    try:
        return float(time_str)
    except ValueError:
        pass

    parts = time_str.split(":")
    seconds = 0.0

    for part in parts:
        seconds = seconds * 60 + float(part)

    return seconds


def open_folder(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


class VideoCutterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Video Cutter")
        self.root.geometry("720x520")
        self.root.minsize(680, 500)

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Sẵn sàng.")
        self.progress_value = tk.DoubleVar(value=0)

        self.worker_thread = None
        self.message_queue = queue.Queue()
        self.is_processing = False

        self.build_ui()
        self.poll_queue()

    def build_ui(self):
        main_frame = ttk.Frame(self.root, padding=16)
        main_frame.pack(fill="both", expand=True)

        title_label = ttk.Label(
            main_frame,
            text="Video Cutter",
            font=("Segoe UI", 18, "bold")
        )
        title_label.pack(anchor="w", pady=(0, 12))

        desc_label = ttk.Label(
            main_frame,
            text="Chọn file CSV/Excel chứa danh sách đoạn cắt, chọn thư mục output, rồi bấm Bắt đầu cắt.",
            wraplength=660
        )
        desc_label.pack(anchor="w", pady=(0, 16))

        input_frame = ttk.LabelFrame(main_frame, text="File input", padding=12)
        input_frame.pack(fill="x", pady=(0, 12))

        input_entry = ttk.Entry(input_frame, textvariable=self.input_path)
        input_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        input_button = ttk.Button(
            input_frame,
            text="Chọn file input",
            command=self.choose_input_file
        )
        input_button.pack(side="left")

        output_frame = ttk.LabelFrame(main_frame, text="Thư mục output", padding=12)
        output_frame.pack(fill="x", pady=(0, 12))

        output_entry = ttk.Entry(output_frame, textvariable=self.output_path)
        output_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        output_button = ttk.Button(
            output_frame,
            text="Chọn thư mục output",
            command=self.choose_output_folder
        )
        output_button.pack(side="left")

        actions_frame = ttk.Frame(main_frame)
        actions_frame.pack(fill="x", pady=(0, 12))

        self.start_button = ttk.Button(
            actions_frame,
            text="Bắt đầu cắt",
            command=self.start_processing
        )
        self.start_button.pack(side="left")

        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_value,
            maximum=100,
            mode="determinate"
        )
        self.progress_bar.pack(fill="x", pady=(8, 8))

        self.status_label = ttk.Label(
            main_frame,
            textvariable=self.status_text
        )
        self.status_label.pack(anchor="w", pady=(0, 12))

        log_frame = ttk.LabelFrame(main_frame, text="Log", padding=8)
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(log_frame, height=16, wrap="word", state="disabled")
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def choose_input_file(self):
        file_path = filedialog.askopenfilename(
            title="Chọn file CSV hoặc Excel",
            filetypes=[
                ("CSV and Excel files", "*.csv *.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*"),
            ]
        )
        if file_path:
            self.input_path.set(file_path)
            self.log(f"Đã chọn file input: {file_path}")

    def choose_output_folder(self):
        folder_path = filedialog.askdirectory(
            title="Chọn thư mục output"
        )
        if folder_path:
            self.output_path.set(folder_path)
            self.log(f"Đã chọn thư mục output: {folder_path}")

    def set_processing_state(self, processing):
        self.is_processing = processing
        if processing:
            self.start_button.configure(state="disabled")
        else:
            self.start_button.configure(state="normal")

    def start_processing(self):
        if self.is_processing:
            return

        input_file = self.input_path.get().strip()
        output_dir = self.output_path.get().strip()

        if not input_file:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn file input.")
            return

        if not output_dir:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn thư mục output.")
            return

        self.progress_value.set(0)
        self.status_text.set("Đang xử lý...")
        self.log("=" * 60)
        self.log("Bắt đầu xử lý...")

        self.set_processing_state(True)

        self.worker_thread = threading.Thread(
            target=self.process_videos,
            args=(input_file, output_dir),
            daemon=True
        )
        self.worker_thread.start()

    def process_videos(self, input_file, output_dir):
        try:
            if input_file.lower().endswith(".csv"):
                df = pd.read_csv(input_file)
            elif input_file.lower().endswith((".xls", ".xlsx")):
                df = pd.read_excel(input_file)
            else:
                self.message_queue.put(("error", "Chỉ hỗ trợ file CSV hoặc Excel."))
                return

            required_columns = ["file_path", "source_in", "source_out"]
            for col in required_columns:
                if col not in df.columns:
                    self.message_queue.put(("error", f"Thiếu cột bắt buộc: {col}"))
                    return

            if df.empty:
                self.message_queue.put(("error", "File input không có dữ liệu."))
                return

            os.makedirs(output_dir, exist_ok=True)
            input_dir = os.path.dirname(input_file)

            total_rows = len(df)
            success_count = 0
            error_count = 0
            error_messages = []

            for index, row in df.iterrows():
                row_number = index + 2
                raw_file_path = str(row["file_path"]).strip()
                source_in_raw = row["source_in"]
                source_out_raw = row["source_out"]

                self.message_queue.put((
                    "log",
                    f"[Dòng {row_number}] Đang xử lý: {raw_file_path}"
                ))

                try:
                    file_path = raw_file_path
                    if not os.path.isabs(file_path):
                        file_path = os.path.join(input_dir, file_path)

                    if not os.path.exists(file_path):
                        raise FileNotFoundError(f"Không tìm thấy file video: {file_path}")

                    source_in = parse_time(source_in_raw)
                    source_out = parse_time(source_out_raw)

                    if source_in is None or source_out is None:
                        raise ValueError("source_in hoặc source_out không hợp lệ.")

                    if source_in >= source_out:
                        raise ValueError("source_in phải nhỏ hơn source_out.")

                    base_name, _ = os.path.splitext(os.path.basename(file_path))
                    out_filename = f"{base_name}_clip_{index + 1}.mp4"
                    out_filepath = os.path.join(output_dir, out_filename)

                    with VideoFileClip(file_path) as video:
                        duration = float(video.duration)

                        if source_in > duration:
                            raise ValueError(
                                f"source_in ({source_in}s) vượt quá độ dài video ({duration:.2f}s)."
                            )

                        if source_out > duration:
                            self.message_queue.put((
                                "log",
                                f"[Dòng {row_number}] source_out vượt độ dài video, tự động giới hạn về {duration:.2f}s."
                            ))
                            source_out = duration

                        clip = video.subclipped(source_in, source_out)
                        clip.write_videofile(
                            out_filepath,
                            codec="libx264",
                            audio_codec="aac",
                            logger=None
                        )
                        clip.close()

                    success_count += 1
                    self.message_queue.put((
                        "log",
                        f"[Dòng {row_number}] Thành công -> {out_filepath}"
                    ))

                except Exception as e:
                    error_count += 1
                    error_message = f"[Dòng {row_number}] Lỗi: {e}"
                    error_messages.append(error_message)
                    self.message_queue.put(("log", error_message))

                progress = ((index + 1) / total_rows) * 100
                self.message_queue.put(("progress", progress))
                self.message_queue.put((
                    "status",
                    f"Đã xử lý {index + 1}/{total_rows} dòng..."
                ))

            summary_lines = [
                "Hoàn tất quá trình cắt video.",
                f"Thành công: {success_count}",
                f"Lỗi: {error_count}",
                f"Thư mục output: {output_dir}"
            ]

            if error_messages:
                summary_lines.append("")
                summary_lines.append("Một số lỗi:")
                summary_lines.extend(error_messages[:10])

            self.message_queue.put(("done", "\n".join(summary_lines)))
            self.message_queue.put(("open_folder", output_dir))

        except Exception as e:
            self.message_queue.put(("error", f"Có lỗi xảy ra: {e}"))

    def poll_queue(self):
        try:
            while True:
                message_type, payload = self.message_queue.get_nowait()

                if message_type == "log":
                    self.log(payload)
                elif message_type == "progress":
                    self.progress_value.set(payload)
                elif message_type == "status":
                    self.status_text.set(payload)
                elif message_type == "done":
                    self.set_processing_state(False)
                    self.status_text.set("Hoàn tất.")
                    self.log("-" * 60)
                    self.log(payload)
                    messagebox.showinfo("Kết quả", payload)
                elif message_type == "error":
                    self.set_processing_state(False)
                    self.status_text.set("Có lỗi.")
                    self.log(f"LỖI: {payload}")
                    messagebox.showerror("Lỗi", payload)
                elif message_type == "open_folder":
                    open_folder(payload)
        except queue.Empty:
            pass

        self.root.after(100, self.poll_queue)

    def log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


def main():
    root = tk.Tk()
    try:
        root.iconbitmap(default="")
    except Exception:
        pass
    app = VideoCutterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()