import easyocr
import cv2
import numpy as np


reader = easyocr.Reader(['vi', 'en'])


def extract_text(file_path):

    img = cv2.imdecode(
        np.fromfile(file_path, dtype=np.uint8),
        cv2.IMREAD_COLOR
    )

    result = reader.readtext(img)

    s = ""

    for (_, text, _) in result:

        s += text + " "

    return s