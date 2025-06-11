import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

RECT_WIDTH = 0.41
RECT_HEIGHT = 0.18
TOTAL_POSITION = (0, 0.1)
BELOW_TEXT_POSITION = (0, -0.15)
COMPARE_TEXT_POSITION = (0, -0.4)
COLORS = {
    'VPBank': '#1EAB4D',
    'Vietcombank': '#80C257',
    'ACB Bank': '#00AEEF',
    'MB Bank': '#141ED2',
    'Techcombank': '#FF0000',
    'MSB': '#E86122',
    'SHB': '#FFC000',
    'Sacombank': '#004A93'
}

def draw_doughnut(ax, values, total_compare, label_title, xtitle, ytitle, ha):

    labels = df['Topic'].tolist()
    percentages = np.round(values / values.sum() * 100, 1)
    color_values = [COLORS[label] for label in labels]

    wedges, _ = ax.pie(
        percentages,
        labels=None,
        startangle=90,
        colors=color_values,
        wedgeprops=dict(width=0.35, edgecolor='white')
    )
    last_x, last_y = None, None 
    for i, wedge in enumerate(wedges):
        angle = (wedge.theta2 + wedge.theta1) / 2
        angle_rad = np.radians(angle)
        shift_count = 0
        max_shifts = 10
        rotation_step_deg = 10
        while True:
            x = np.cos(angle_rad)
            y = np.sin(angle_rad)
            if last_x is None or np.sqrt((x - last_x)**2 + (y - last_y)**2) >= 0.16 or shift_count >= max_shifts:
                break
            angle_rad += np.radians(rotation_step_deg)
            shift_count += 1
        last_x, last_y = x, y
        rect = plt.Rectangle((x - 0.15, y), RECT_WIDTH, RECT_HEIGHT, color=color_values[i], ec='white', linewidth=0.5, zorder=3)
        ax.add_patch(rect)
        ax.text(x - 0.15 + RECT_WIDTH / 2, y + RECT_HEIGHT / 2, f'{percentages[i]}%', ha='center', va='center', fontsize=5, color='white', fontweight='bold', zorder=4)

    total_value = values.sum()
    compare_value = total_value - total_compare
    percent_diff = np.round((compare_value / total_compare) * 100, 0) if total_compare != 0 else 0
    compare_text = f"{percent_diff}%" if percent_diff < 0 else f"+{percent_diff}%"
    compare_color = '#ff0000' if percent_diff < 0 else '#1EAB4D'

    ax.text(xtitle, ytitle, label_title, ha=ha, va='bottom', fontsize=8, fontweight='bold', color='#00AEEF')
    ax.text(TOTAL_POSITION[0], TOTAL_POSITION[1], f"{total_value:,.0f}", ha='center', va='center', fontsize=16, fontweight='bold', color='#00AEEF')
    ax.text(BELOW_TEXT_POSITION[0], BELOW_TEXT_POSITION[1], "THẢO LUẬN", ha='center', va='center', fontsize=8, fontweight='bold', color='#00AEEF')
    ax.text(COMPARE_TEXT_POSITION[0], COMPARE_TEXT_POSITION[1], compare_text, ha='center', va='center', fontsize=10, fontweight='bold', color=compare_color)
    

def main(df):
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(5, 2.5))
    fig.subplots_adjust(wspace=0.05)
    if not df['This Week'].empty and not df['Last Week'].empty:
        draw_doughnut(
            ax1,
            df['This Week'],
            total_compare=df['Last Week'].sum(),
            label_title="Tuần này",
            xtitle=-1.3,
            ytitle=-1.25,
            ha='left'
        )
    
        draw_doughnut(
            ax2,
            df['Last Week'],
            total_compare=df['This Week'].sum(),
            label_title="Tuần trước",
            xtitle=1.3,
            ytitle=-1.25,
            ha='right'
        )

        fig.legend(
            labels = df['Topic'].tolist(),
            loc='upper center',
            bbox_to_anchor=(0.5, 0.1),
            ncol=8,
            fontsize=4,
            frameon=False
        )

        plt.savefig("./img/doughnut_side_by_side.png", dpi=1280, bbox_inches="tight")

if __name__ == "__main__":

    df = pd.read_excel('./data/doughnut.xlsx', sheet_name='Sheet1', nrows=8).dropna(how='all').fillna(0)
    main(df)
