import os
import argparse
import pandas as pd
from moviepy.editor import VideoFileClip

def parse_time(time_str):
    """
    Chuyển đổi thời gian dạng chuỗi 'HH:MM:SS', 'MM:SS' hoặc giây sang giây (dạng số thực).
    """
    if pd.isna(time_str):
        return None
    time_str = str(time_str).strip()
    if not time_str:
        return None

    # Thử xử lý nếu giá trị chỉ là số (float/int represent seconds)
    try:
        return float(time_str)
    except ValueError:
        pass

    # Xử lý chuỗi có định dạng HH:MM:SS hoặc MM:SS
    parts = time_str.split(':')
    seconds = 0
    for part in parts:
        seconds = seconds * 60 + float(part)
    return seconds

def main():
    parser = argparse.ArgumentParser(description="Tự động cắt video dựa trên file cấu hình CSV hoặc Excel.")
    parser.add_argument("-i", "--input", help="Đường dẫn đến file input CSV hoặc Excel", required=True)
    args = parser.parse_args()

    input_file = args.input

    print(f"[*] Đang đọc dữ liệu từ file: {input_file}")
    try:
        if input_file.lower().endswith('.csv'):
            df = pd.read_csv(input_file)
        elif input_file.lower().endswith(('.xls', '.xlsx')):
            df = pd.read_excel(input_file)
        else:
            print("[-] Lỗi: Định dạng file không được hỗ trợ. Vui lòng sử dụng file CSV hoặc Excel.")
            return
    except Exception as e:
        print(f"[-] Lỗi khi đọc file {input_file}: {e}")
        return

    # Kiểm tra xem có đủ các cột bắt buộc không
    required_columns = ['file_path', 'source_in', 'source_out']
    for col in required_columns:
        if col not in df.columns:
            print(f"[-] Lỗi: File đính kèm thiếu cột bắt buộc '{col}'.")
            return

    # Tạo thư mục output nếu chưa tồn tại
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    print(f"[*] Thư mục đầu ra './{output_dir}/' đã sẵn sàng.")

    success_count = 0
    error_count = 0

    # Lặp qua từng dòng trong bảng dữ liệu
    for index, row in df.iterrows():
        file_path = str(row['file_path']).strip()
        source_in_raw = row['source_in']
        source_out_raw = row['source_out']

        print(f"\n[+] Đang xử lý dòng {index + 1}: {file_path}")

        try:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"Không tìm thấy file video tại '{file_path}'")

            source_in = parse_time(source_in_raw)
            source_out = parse_time(source_out_raw)

            if source_in is None or source_out is None:
                raise ValueError("Giá trị 'source_in' hoặc 'source_out' không hợp lệ hoặc bị trống.")
            
            if source_in >= source_out:
                raise ValueError(f"Thời gian bắt đầu ({source_in}s) lớn hơn hoặc bằng thời gian kết thúc ({source_out}s).")

            # Tạo tên file output tránh ghi đè: Tên gốc + "_clip_" + (index + 1) + ".mp4"
            # Cố định luôn đuôi xuất là .mp4 (có thể lấy đuôi của nguồn gốc tuỳ ý)
            base_name, _ = os.path.splitext(os.path.basename(file_path))
            out_filename = f"{base_name}_clip_{index + 1}.mp4"
            out_filepath = os.path.join(output_dir, out_filename)

            # Mở video và cắt
            with VideoFileClip(file_path) as video:
                # Kiểm tra độ dài cơ bản
                if source_in > video.duration:
                    raise ValueError(f"Thời gian bắt đầu cắt ({source_in}s) vượt quá độ dài video ({video.duration}s)")
                    
                # Giới hạn source_out sao cho ko vượt quá video_duration
                if source_out > video.duration:
                    print(f"    [!] Cảnh báo: Thời gian kết thúc vượt quá độ dài video. Đã tự giới hạn lại!")
                    source_out = video.duration
                    
                clip = video.subclip(source_in, source_out)
                print(f"    -> Đang kết xuất video để lưu tại: {out_filepath}")
                clip.write_videofile(
                    out_filepath, 
                    codec="libx264", 
                    audio_codec="aac",
                    logger='bar' # Hiển thị thanh tiến trình
                )
            
            print(f"[*] Xử lý thành công dòng {index + 1}")
            success_count += 1

        except Exception as e:
            print(f"[-] BỎ QUA - Lỗi tại dòng {index + 1}: {e}")
            error_count += 1
            continue
            
    print("\n" + "="*40)
    print("HOÀN TẤT QUÁ TRÌNH CẮT VIDEO!")
    print(f"Thành công: {success_count} | Lỗi: {error_count}")
    print("="*40)

if __name__ == "__main__":
    main()
