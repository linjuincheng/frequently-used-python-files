import matplotlib.pyplot as plt

def draw_shelf_item():
    fig, ax = plt.subplots(figsize=(4, 4))
    
    # 畫正方形外框（紅色粗體）
    square = plt.Rectangle((1, 1), 2, 2, linewidth=3, edgecolor='red', facecolor='none')
    ax.add_patch(square)
    
    # 在正方形內寫入 "Wet wipe"（紅色粗體）
    ax.text(2, 2, 'Wet wipes', color='red', fontsize=16, fontweight='bold', ha='center', va='center')
    
    # 設定顯示範圍
    ax.set_xlim(0, 4)
    ax.set_ylim(0, 4)
    
    # 隱藏軸刻度
    ax.set_xticks([])
    ax.set_yticks([])
    ax.set_frame_on(False)
    
    plt.show()

# 繪製正方形表示貨架上的東西
draw_shelf_item()







# 我要畫一個正方形來顯示擺在貨架上的東西
# 1.	正方形的外框用紅色粗體
# 2.	正方形內寫入Wet wipe(字體用紅色粗體)
# Python語法
