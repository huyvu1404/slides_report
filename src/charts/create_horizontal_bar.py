import matplotlib.pyplot as plt
import numpy as np
from matplotlib.gridspec import GridSpec
from matplotlib.patches import Patch


conversations = ["BAN LÃNH ĐẠO", "CHI NHÁNH LIÊN DOANH", "CHIẾN DỊCH & SỰ KIỆN", "CHỨNG KHOÁN", "DỊCH VỤ KHÁCH HÀNG", "HOẠT ĐỘNG FANPAGE", 
          "NHÂN SỰ & TUYỂN DỤNG", "SẢN PHẨM DỊCH VỤ CÁ NHÂN", "SẢN PHẨM DỊCH VỤ DOANH NGHIỆP", "SẢN PHẨM THẺ", "THƯƠNG HIỆU CHUNG"]
topics = ["SHB", "Vietcombank",  "Techcombank", "MB Bank", "VPBank", "MSB", "Sacombank", "ACB Bank"]
sentiments = ["TÍCH CỰC", "TRUNG LẬP", "TIÊU CỰC"]


# Màu sắc cho các sentiment
colors = {
    "TÍCH CỰC": "#92D050",  # Màu xanh lá
    "TRUNG LẬP": "#A6A6A6",  # Màu xám
    "TIÊU CỰC": "#C00000"  # Màu đỏ
}
# Tạo figure với 2 axes xếp ngang
width_ratios=[2, 0.8, 1.1, 1, 1.1, 1, 1.2, 0.95]
def create_horizontal_bar_chart(data):
    n_rows, n_cols = len(conversations), len(topics)
    fig = plt.figure(figsize=(22, 8))
    gs = GridSpec(1, 8, width_ratios= width_ratios, wspace=0)

    first_ax = fig.add_subplot(gs[0])
    axs = [first_ax]
    for i in range(1, n_cols):
        ax = fig.add_subplot(gs[i], sharey=first_ax)
        axs.append(ax)

    max_value = 110
    y = np.arange(0, 1.4 * n_rows, 1.4)
    y_middle = y - 0.7
    print(y_middle)
    print(y)
    for i, (ax, topic) in enumerate(zip(axs, topics)):
        print(f"Processing topic: {topic}")
        for spine in ax.spines.values():
            spine.set_visible(False)
        
        xmin = -120 if i == 0 else 0
        ax.set_xlim(xmin, max_value)
        ax.set_ylim(-1.2, n_rows*1.3)
        ax.hlines(y=y[:-1], xmin=xmin, xmax=max_value, colors='gray', linestyles='-', linewidth=1, alpha=0.7, zorder=0)
        ax.vlines(x=0, ymin=-1.4, ymax=n_rows*1.4, colors='gray', linestyles='-', linewidth=1, alpha=0.7, zorder=1)
        ax.tick_params(axis='y', which='both', length=0)
        ax.set_yticklabels([])
        ax.set_xticks([])
        ax.tick_params(axis='x', which='both', length=0)

        bottom = np.full((n_rows,), 4.0)
        df = data[topic].T * 100
        for j, color in enumerate(colors.values()):
            values = np.round(df[j].tolist(), 1)
            ax.barh(y_middle, values, left=bottom, color=color, height=0.8, zorder=10)
            # Add text labels above each bar segment

            bottom += values
        for k, bottom in enumerate(bottom):
            if bottom > 40:
                x_pos = bottom * 1.5
            elif bottom > 20:
                x_pos = bottom * 2
            elif bottom > 10:
                x_pos = bottom * 3
            else:
                x_pos = 30
            y_pos = y_middle[k]
            if bottom - 4 > 0:
                ax.text(x_pos, y_pos, f'{round(bottom - 4, 2)}%', ha='center', va='center', 
                    fontsize=10, color='black', fontweight='bold', zorder=20)
        # Existing text for total sum


    legend_elements = [Patch(facecolor=colors[sentiment], label=sentiment) for sentiment in sentiments]

    # Create the legend
    fig.legend(
        handles=legend_elements,
        labels=sentiments,
        loc='lower center',
        bbox_to_anchor=(0.5, 0.07),
        ncol=len(sentiments),

        prop={'weight': 'bold', 'size': 10},  
        frameon=False,
        handlelength=0.65,
        handletextpad=0.5,
        columnspacing=1.5
    )

    plt.savefig('./img/horizontal_bar.png', dpi=300, bbox_inches='tight')
    return topics, width_ratios
# if __name__ == "__main__":
#     create_horizontal_bar_chart(data)
#     plt.show()
# plt.show()