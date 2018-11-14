from pdf2image import convert_from_path, convert_from_bytes
import numpy as np
import cv2

path = ".\Question & Dataset\RRD BYOB 2018 (Problem statements _ Datasets)\DataSet_BYOB Challenge 2\A7FX7D8H.pdf"

PIL_images = convert_from_path(path)

images = []
for img in PIL_images:
    open_cv_image = np.array(img) 
    # Convert RGB to BGR 
    images.append(open_cv_image[:, :, ::-1].copy())

# Processing the images page by page
for img in images:
    img = cv2.bitwise_not(img)
    cv2.imshow('img', img)
    kernel = np.ones((3,3),np.uint8)
    erosion = cv2.erode(img, kernel, iterations = 1)
    cv2.imshow('erosion' ,erosion)
    opening = cv2.morphologyEx(img, cv2.MORPH_OPEN, kernel)
    cv2.imshow('opening', opening)
    if cv2.waitKey(0) & 0xFF == ord('q'):
        break