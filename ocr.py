import easyocr
import cv2
import numpy as np

import tkinter as tk
from tkinter import filedialog
from docx import Document


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
    # with open("tam.txt",'w',encoding='utf-8') as f:
    #     for(bbox,text,prob) in result:
    #         f.write(text)
    s = ""
    for(bbox,text,prob) in result:
        print(text)
        print(bbox)
        s += "".join(text + " ")
    print("_"*20)
    print(s)

# def save_file(filename,content):
#     file_path = filedialog.asksaveasfilename(
#         defaultextension=""
#     )
