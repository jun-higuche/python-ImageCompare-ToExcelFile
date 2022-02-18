import cv2, os
import numpy as np
from opencv_japanese import imread, imwrite

'''
画像を比較して異なる箇所だけを炙り出した画像を生成する。
'''
def createDiffImg(imageName1, imageName2, saveImageName):
    
    image1 = imread(imageName1)
    image2 = imread(imageName2)

    height = image1.shape[0]
    width = image1.shape[1]

    img_size = (int(width), int(height))

    # 画像をリサイズする
    image1 = cv2.resize(image1, img_size)
    image2 = cv2.resize(image2, img_size)

    im_diff = image1.astype(int) - image2.astype(int)
    im_diff_center = np.floor_divide(im_diff, 2) + 128

    imwrite(saveImageName, im_diff_center)

'''
画像を比較して類似度を返す。
1に近いほど、同じ画像
'''
def imageCompareNum(imageName1, imageName2):
    
    image1 = imread(imageName1)
    image2 = imread(imageName2)

    height = image1.shape[0]
    width = image1.shape[1]

    img_size = (int(width), int(height))

    # 画像をリサイズする
    image1 = cv2.resize(image1, img_size)
    image2 = cv2.resize(image2, img_size)

    return np.count_nonzero(image1 == image2) / image1.size
