import cv2
import matplotlib.pyplot as plt

# 讀取圖片（灰階）
image = cv2.imread(r'C:\_python 2024\2024final\test.jpg', cv2.IMREAD_GRAYSCALE)

# 使用 Canny 邊緣偵測（設定兩個閾值）
edges = cv2.Canny(image, threshold1=100, threshold2=200)

# 顯示結果（使用 matplotlib）
plt.subplot(1, 2, 1)
plt.title('Original Image')
plt.imshow(image, cmap='gray')
plt.axis('off')

plt.subplot(1, 2, 2)
plt.title('Canny Edges')
plt.imshow(edges, cmap='gray')
plt.axis('off')

plt.tight_layout()
plt.show()




# 🔧 說明：
# threshold1 和 threshold2 是兩個邊緣判斷用的參數：

# 邊緣梯度強度 > threshold2 → 確定是邊緣

# 邊緣梯度介於 threshold1 ~ threshold2 → 視情況判定

# < threshold1 → 被忽略
