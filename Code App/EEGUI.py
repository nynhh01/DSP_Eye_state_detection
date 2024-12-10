import tkinter as tk
from tkinter import messagebox, filedialog
import joblib
import numpy as np
import pandas as pd

# Hàm lấy giá trị từ các ô nhập liệu
def get_input_data(entries):
    input_data = []
    for entry in entries:
        try:
            value = float(entry.get())
            input_data.append(value)
        except ValueError:
            messagebox.showerror("Lỗi", f"Vui lòng nhập giá trị hợp lệ cho {entry.get()}")
            return None
    return np.array(input_data).reshape(1, -1)

# Tải mô hình SVM đã lưu
try:
    model = joblib.load('SVM_model.joblib')  # Đảm bảo đường dẫn đúng tới mô hình
except FileNotFoundError:
    messagebox.showerror("Lỗi", "Mô hình SVM không tìm thấy! Vui lòng kiểm tra lại.")
    exit()

# Hàm dự đoán
def predict():
    input_data = get_input_data(entries)
    if input_data is None:
        return
    # Dự đoán với mô hình SVM
    prediction = model.predict(input_data)
    result = "Mắt nhắm" if prediction[0] == 0 else "Mắt mở"
    result_label.config(text=f"Dự đoán: {result}", font=("Helvetica", 12, "bold"))

# Hàm đọc file Excel, gán nhãn vào cột 'eyeDetection' và lưu lại vào file Excel mới
def load_and_predict_from_excel(file_path):
    try:
        # Đọc dữ liệu từ file Excel
        df = pd.read_excel(file_path)

        # Gán nhãn vào cột 'eyeDetection' dựa trên mô hình
        features = df.iloc[:, :-1].values  # Giả sử tất cả các cột trừ cột cuối là đặc trưng
        predictions = model.predict(features)

        # Thêm kết quả dự đoán vào cột 'eyeDetection'
        df['eyeDetection'] = predictions

        # Lưu lại dữ liệu đã gán nhãn vào file Excel mới
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Thành công", f"Đã lưu kết quả dự đoán vào file: {save_path}")
        else:
            messagebox.showerror("Lỗi", "Không lưu được file. Vui lòng chọn đường dẫn lưu.")
        
    except Exception as e:
        messagebox.showerror("Lỗi", f"Đã xảy ra lỗi khi xử lý file Excel: {e}")

# Tạo cửa sổ ứng dụng Tkinter
root = tk.Tk()
root.title("EEG Detection")

# Tạo một khung chính với lề 100px
main_frame = tk.Frame(root, padx=100, pady=100)
main_frame.pack(fill="both", expand=True)

# Tạo các nhãn và ô nhập liệu cho mỗi đặc trưng
labels = ["AF3", "F7", "F3", "FC5", "T7", "P7", "O1", "O2", "P8", "T8", "FC6", "F4", "F8", "AF4"]
entries = []

for i, label in enumerate(labels):
    tk.Label(main_frame, text=f"{label}:").grid(row=i, column=0, sticky="w", padx=5, pady=5)
    entry = tk.Entry(main_frame)
    entry.grid(row=i, column=1, sticky="ew", padx=5, pady=5)
    entries.append(entry)

# Thêm một Label để hiển thị kết quả dự đoán
result_label = tk.Label(main_frame, text="", font=("Helvetica", 12))
result_label.grid(row=len(labels), column=0, columnspan=2, sticky="w", padx=5, pady=5)

# Tạo nút để dự đoán
predict_button = tk.Button(main_frame, text="Dự đoán", command=predict, bg="limegreen", fg="black", width=20)
predict_button.grid(row=len(labels) + 1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

# Tạo nút để xóa các ô nhập liệu
def clear_fields():
    for entry in entries:
        entry.delete(0, tk.END)
    result_label.config(text="")

clear_button = tk.Button(main_frame, text="Xóa", command=clear_fields, bg="red", fg="black")
clear_button.grid(row=len(labels) + 2, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

# Tạo nút để tải và dự đoán từ file Excel
def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if file_path:
        load_and_predict_from_excel(file_path)

load_button = tk.Button(main_frame, text="Tải dữ liệu cần dự đoán từ Excel", command=load_file, bg="gold", fg="black")
load_button.grid(row=len(labels) + 3, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

# Cấu hình để các hàng và cột có thể thay đổi kích thước
main_frame.grid_rowconfigure(len(labels), weight=1)
main_frame.grid_rowconfigure(len(labels) + 1, weight=1)
main_frame.grid_rowconfigure(len(labels) + 2, weight=1)
main_frame.grid_rowconfigure(len(labels) + 3, weight=1)

main_frame.grid_columnconfigure(0, weight=1)
main_frame.grid_columnconfigure(1, weight=3)

# Chạy ứng dụng Tkinter
root.mainloop()