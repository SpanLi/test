#!/usr/bin/python3
# --*-- coding:utf-8 --*--



#看了一下估计区分度都会比较高，直接用otsu+binary—inv就够了

import cv2

testimg = "/home/gk/ID.jpg"
src = cv2.imread(testimg)
s_gray = cv2.cvtColor(src,cv2.COLOR_BGR2GRAY)
threshold,imgOtsu = cv2.threshold(s_gray,0,255,cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
cv2.imshow("dst3", imgOtsu)
cv2.waitKey(0)
