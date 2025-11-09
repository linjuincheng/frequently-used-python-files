import cv2
import numpy as np
import matplotlib.pyplot as plt

# 讀取圖片（灰階）
image = cv2.imread(r'C:\_python 2024\2024final\test.jpg', 0)

# 建立 kernel（結構元素）
kernel = np.ones((5,5), np.uint8)

# 執行開運算（先侵蝕再膨脹）
opening = cv2.morphologyEx(image, cv2.MORPH_OPEN, kernel)

# 顯示原圖與處理後的圖
plt.subplot(1, 2, 1)
plt.title('Original')
plt.imshow(image, cmap='gray')
plt.axis('off')

plt.subplot(1, 2, 2)
plt.title('After Opening')
plt.imshow(opening, cmap='gray')
plt.axis('off')

plt.show()

