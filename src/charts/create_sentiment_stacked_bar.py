import matplotlib.pyplot as plt
import pandas as pd
import numpy as np

COLORS = {
    'Positive': '#1EAB4D',  
    'Neutral': '#A6A6A6',  
    'Negative': '#C00000',
   
}
FIGSIZE = (9.74, 2.18)
BAR_DISTANCE = 6
WIDTH_BAR = 0.7

HEIGHT_RECT = 10

def generate_sentiment_stacked_bar_chart(data):
 
    x_positions = []
    for i in range(len(topics)):
        base = i * 3
        x_positions.extend([base, base + 1])
    fig, ax = plt.subplots(figsize=FIGSIZE)
    bottom = np.zeros(len(x_positions))
    rect_distance = np.zeros(len(x_positions))
    y_positions = np.zeros(len(x_positions))
    for sentiment in sentiments:
        values = data[sentiment].astype(int).tolist()
        sizes = np.round(data[sentiment] / data['Total'] * 100, 1).tolist() \
            if sentiment in data.columns else np.zeros(len(df)).astype(int).tolist()

        ax.bar(x_positions, sizes, bottom=bottom, color=COLORS[sentiment], label=sentiment, width=WIDTH_BAR)

        for i, size in enumerate(sizes):
            if size < HEIGHT_RECT:
                if bottom[i] == 0:
                    y_positions[i] = HEIGHT_RECT / 2
                else:
                    y_positions[i] = rect_distance[i] + HEIGHT_RECT / 2
                rect_distance[i] += HEIGHT_RECT
            else:
                if bottom[i] == 0:
                    y_positions[i] = size / 2
                    rect_distance[i] = size
                else:
                    if size / 2 >= rect_distance[i] + HEIGHT_RECT / 2 - bottom[i]:
                        y_positions[i] = bottom[i] + size / 2
                        rect_distance[i] = bottom[i] + size
                    else:
                        y_positions[i] = rect_distance[i] + HEIGHT_RECT / 2
                        rect_distance[i] += HEIGHT_RECT
        
        for i in range(0, len(x_positions), 2):
            left_size = sizes[i]
            right_size = sizes[i + 1]
            offset = 12
            if left_size > right_size:
                y_positions[i] = max(y_positions[i + 1] + offset, y_positions[i])
                rect_distance[i] = max(rect_distance[i + 1] + offset, rect_distance[i])
            elif right_size > left_size:
                y_positions[i + 1] = max(y_positions[i] + offset, y_positions[i + 1])
                rect_distance[i + 1] = max(rect_distance[i] + offset, rect_distance[i + 1])
              
        for i in range(len(x_positions)):
            if values[i] < 100:
                width_rect = WIDTH_BAR * 1.3
            elif values[i] < 1000:
                width_rect = WIDTH_BAR * 1.5
            else:
                width_rect = WIDTH_BAR * 2
                

            rect = plt.Rectangle(
                (x_positions[i] - width_rect / 2, y_positions[i] - HEIGHT_RECT / 2),
                width_rect, HEIGHT_RECT,
                facecolor=COLORS[sentiment],
                linewidth=0.5,
                zorder=3
            )
            ax.add_patch(rect)
            label = f'{values[i]}, {sizes[i]}%'

            ax.text(
                x_positions[i], y_positions[i],
                label,
                ha='center', va='center',
                fontsize=4.5, color='white', fontweight='bold',
                zorder=4
            )
        bottom += sizes

    
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.axhline(0, color='gray', linewidth=0.5)
    ax.tick_params(axis='x', length=0, width=0)
    ax.set_xticks(x_positions)
    ax.set_xticklabels(x_labels, fontsize = 4.5, fontweight='bold', color='black', rotation=45, ha='right')
    ax.set_xlim(-0.8 , max(x_positions) + 0.8)
    ax.set_ylim(bottom=0)
    ax.yaxis.set_visible(False)

    plt.tight_layout()
    plt.savefig('./img/sentiment_stacked_bar_chart.png', dpi=1280, bbox_inches='tight')
    return topics

if __name__ == "__main__":
    df = pd.read_excel('./data/sentiment.xlsx', sheet_name='Sheet1', nrows=25, usecols='A, E:I').dropna(how='all').dropna(how='all', axis=1).fillna(0)
    df.columns = ['Period', 'Topic', 'Positive', 'Neutral', 'Negative']
    sentiments = df.columns[2:].tolist() 
    df['Total'] = df[sentiments].sum(axis=1) 
    topics = df['Topic'].unique()
    x_labels = df['Period'].tolist() 
    generate_sentiment_stacked_bar_chart(df)