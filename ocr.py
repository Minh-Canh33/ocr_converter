import easyocr
import cv2
import numpy as np
# image_path = r"D:\picture.png"
# reader = easyocr.Reader(['vi','en'])
# result = reader.readtext(image_path) #-> toa độ,text, độ tin cậy
# for (bbox,text,prob) in result:
#     print(text)
# print(type(result))


import tkinter as tk
from tkinter import filedialog

# Ẩn cửa sổ chính của tkinter
root = tk.Tk()
root.withdraw()

# Mở hộp thoại chọn file và lấy đường dẫn
file_path = filedialog.askopenfilename(
    title="Chọn file",
    filetypes=[("Image files", "*.jpg"),("Image files", "*.png"), ("All files", "*.*")]
)
if file_path:
    path = cv2.imdecode(np.fromfile(file_path,dtype=np.uint8),cv2.IMREAD_COLOR)
    # print(path)
    reader = easyocr.Reader(['vi','en'])
    result = reader.readtext(path)
    with open("tam.txt",'w',encoding='utf-8') as f:
        for(bbox,text,prob) in result:
            f.write(text)