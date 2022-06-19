import cv2
import numpy as np
from PIL import Image

img = cv2.imread('C:/Users/Raghul Devaraj/Pictures/Saved Pictures/My photo.jpg')

imgStack = np.hstack([img, img])
cv2.imshow("My img", imgStack)
cv2.waitKey(0)
print(img.shape)