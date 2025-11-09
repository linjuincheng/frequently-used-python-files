import matplotlib.pyplot as plt
import numpy as np

def draw_stock_level_chart(stock_level=100):
    fig, ax = plt.subplots(figsize=(2, 6))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 100)
    
    # 畫長條形
    ax.bar(0.5, stock_level, width=0.8, color='lightgray', edgecolor='black')
    
    # 畫紅色30%線
    ax.axhline(y=30, color='red', linewidth=2, linestyle='--')
    ax.text(0.5, 30, '30%', color='red', ha='center', va='bottom', fontsize=10, fontweight='bold')
    
    # 畫綠色80%線
    ax.axhline(y=80, color='green', linewidth=2, linestyle='--')
    ax.text(0.5, 80, '80%', color='green', ha='center', va='bottom', fontsize=10, fontweight='bold')
    
    # 標示不同區域
    ax.text(0.5, 10, 'Low Stock', color='red', ha='center', fontsize=12, fontweight='bold')
    ax.text(0.5, 50, 'Safety Stock', color='gold', ha='center', fontsize=12, fontweight='bold')
    ax.text(0.5, 90, 'High Stock', color='green', ha='center', fontsize=12, fontweight='bold')
    
    # 隱藏 X 軸刻度
    ax.set_xticks([])
    ax.set_yticks([])
    
    plt.show()

# 測試繪製庫存水位圖
draw_stock_level_chart(100)



# 我要畫一個長條圖來顯示東西庫存的水位
# 1.	圖表用直立的長方形表示
# 2.	長方形內畫2條橫線—依序由下而上, 第1條線用紅色(且標示30%), 第2條線畫綠色(且標示80%)
# 3.	在第1條橫線到長方形底部的空白區域內寫入Low Stock(用紅色字體)
# 4.	在第1條橫線到第2條橫線間的空白區域內寫入Safety Stock (用黃色字體)
# 5.	在第2條橫線到長方形頂部的空白區域內寫入High Stock (用綠色字體)

