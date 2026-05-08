import tkinter as tk
from tkinter import filedialog

from PIL import Image, ImageTk

from ocr_engine import extract_text
from save_utils import save_docx, save_pdf


# =========================
# ROOT
# =========================
root = tk.Tk()

root.title("OCR App")

root.geometry("1000x600")


# =========================
# FUNCTIONS
# =========================

# OPEN IMAGE
def open_image():

    file_path = filedialog.askopenfilename(
        title="Chọn file",
        filetypes=[
            ("Image files", "*.jpg"),
            ("Image files", "*.png")
        ]
    )

    if file_path:

        # OCR
        text = extract_text(file_path)

        text_area.delete("1.0", tk.END)

        text_area.insert(tk.END, text)

        # SHOW IMAGE
        img = Image.open(file_path)

        img.thumbnail((350, 500))

        photo = ImageTk.PhotoImage(img)

        image_label.config(image=photo)

        image_label.image = photo


# SAVE DOCX
def on_save_docx():

    path = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word Document", "*.docx")]
    )

    if path:

        save_docx(
            path,
            text_area.get("1.0", tk.END)
        )


# SAVE PDF
def on_save_pdf():

    path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")]
    )

    if path:

        save_pdf(
            path,
            text_area.get("1.0", tk.END)
        )


# CLOSE APP
def close_app():

    root.destroy()


# =========================
# TOP FRAME
# =========================
top_frame = tk.Frame(root)

top_frame.pack(
    fill='x',
    pady=5,
    padx=10
)


# OPEN BUTTON
btn_open = tk.Button(
    top_frame,
    text="Open Image",
    command=open_image,
    width=15
)

btn_open.pack(side=tk.LEFT)


# CLOSE BUTTON
btn_close = tk.Button(
    top_frame,
    text="Close",
    command=close_app,
    width=10
)

btn_close.pack(side=tk.RIGHT, padx=5)


# SAVE PDF BUTTON
btn_pdf = tk.Button(
    top_frame,
    text="Save PDF",
    command=on_save_pdf,
    width=10
)

btn_pdf.pack(side=tk.RIGHT, padx=5)


# SAVE DOCX BUTTON
btn_docx = tk.Button(
    top_frame,
    text="Save DOCX",
    command=on_save_docx,
    width=10
)

btn_docx.pack(side=tk.RIGHT, padx=5)


# =========================
# CONTENT FRAME
# =========================
content_frame = tk.Frame(root)

content_frame.pack(
    expand=True,
    fill='both',
    padx=10,
    pady=10
)


# =========================
# LEFT FRAME (IMAGE)
# =========================
left_frame = tk.Frame(
    content_frame,
    width=400,
    bd=2,
    relief=tk.GROOVE
)

left_frame.pack(
    side=tk.LEFT,
    fill='both',
    padx=5
)

left_frame.pack_propagate(False)


# IMAGE LABEL
image_label = tk.Label(left_frame)

image_label.pack(
    expand=True
)


# =========================
# RIGHT FRAME (TEXT)
# =========================
right_frame = tk.Frame(
    content_frame,
    bd=2,
    relief=tk.GROOVE
)

right_frame.pack(
    side=tk.RIGHT,
    expand=True,
    fill='both',
    padx=5
)


# TEXT AREA
text_area = tk.Text(
    right_frame,
    wrap=tk.WORD,
    font=("Arial", 12)
)

text_area.pack(
    expand=True,
    fill='both',
    padx=5,
    pady=5
)


# =========================
# MAIN LOOP
# =========================
root.mainloop()