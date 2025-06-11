import matplotlib.pyplot as plt
import pandas as pd
import numpy as np

COLORS = {
    'Facebook': '#0070C0',  
    'Fanpage': '#00B0F0',  
    'Diễn đàn': '#EC977C',
    'Tin tức': '#FFC000',
    'Tiktok': '#1EAB4D',
    'Youtube': '#DF5327',
    'E-commerce': '#27536B',
    'Instagram': '#646E17'
} 
FIGSIZE = (9.74, 2.18)
BAR_DISTANCE = 10
WIDTH_BAR = 3

def generate_stacked_bar_chart(data):
    channels = data.columns[1:-1].tolist()  
    topics = data['TOPIC'].unique()
    x_labels = data['Column1'].tolist()  
    x_positions = []
    for i in range(len(topics)):
        base = i * BAR_DISTANCE
        x_positions.extend([base, base + 3])
    fig, ax = plt.subplots(figsize=FIGSIZE)
    bottom = np.zeros(len(x_positions))
    max_value  = data.iloc[:,1:-1].sum().max()
    rect_distance = np.zeros(len(x_positions))
    for channel in channels:
        values = data[channel].astype(int).tolist() if channel in data.columns else np.zeros(len(df)).astype(int).tolist()
        ax.bar(x_positions, values, bottom=bottom, color=COLORS[channel], label=channel, width=WIDTH_BAR)
        for i, (x, val) in enumerate(zip(x_positions, values)):
            if val > 0:
                if val < 10:
                    width_rect = 1
                elif val < 100:
                    width_rect = 1.5
                elif val < 1000:
                    width_rect = 2
                elif val < 10000:
                    width_rect = 2.5
                elif val < 100000:
                    width_rect = 3
                else:
                    width_rect = 3.5

                height_rect = max_value / 8
                if rect_distance[i] > 0:
                    if val/2 > rect_distance[i]:
                        y = val / 2 
                        rect_distance[i] = val/2 + max_value / 8
                    else:
                        y  =  rect_distance[i]
                        rect_distance[i] += max_value / 8
                else:
                    y = 0
                    rect_distance[i] += max_value / 8
                
                rect = plt.Rectangle(
                    (x - width_rect / 2, y),
                    width_rect, height_rect,
                    color=COLORS[channel],                 
                    linewidth=0.5,
                    zorder=3
                )
                ax.add_patch(rect)
                ax.text(
                    x,  y + height_rect/2,
                    str(val),
                    ha='center', va='center',
                    fontsize=5,
                    color='white',
                    fontweight='bold',
                    zorder=4
                )
                
                
        bottom += values
    for x, total, position in zip(x_positions, bottom, rect_distance):
        ax.text(
            x, position * 1.1,                
            str(int(total)),               
            ha='center',
            va='bottom',
            fontsize=7,
            fontweight='bold',
            color='black'
        )

    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.axhline(0, color='gray', linewidth=0.5)
    ax.tick_params(axis='x', length=2, width=0.5, colors='gray')
    ax.set_xticks(x_positions)
    ax.set_xticklabels(x_labels, fontsize = 4.5, fontweight='bold', color='black')
    ax.set_xlim(-2 , max(x_positions) + 2)
    ax.set_ylim(bottom=0)


    fig.subplots_adjust(top=0.6)  
    ax.legend(loc='upper center', bbox_to_anchor=(0.77, 1.3), ncol=len(COLORS), fontsize=6, frameon=False, handlelength=0.65, handletextpad=0.5, columnspacing=0.5)

    ax.yaxis.set_visible(False)
        # Ví dụ tên hình theo topic
# Tăng vùng dưới để có chỗ đặt hình
    

    # Chuẩn bị đường dẫn ảnh theo topic
    unique_topics = data['TOPIC'].unique()

    plt.tight_layout()
    plt.savefig('./img/stacked_bar_chart.png', dpi=300, bbox_inches='tight')
    return unique_topics

# if __name__ == "__main__":
#     df = pd.read_excel('./data/stacked_chart.xlsx', sheet_name='Sheet1', nrows=24, usecols='A:J').dropna(how='all').fillna(0)
#     unique_topics = generate_stacked_bar_chart(df)
#     print(f"Unique topics: {unique_topics}")