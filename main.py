import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.dates import date2num


def load_roadmap_data(filename):
    tech_df = pd.read_excel(filename, sheet_name="Technology")
    cap_df = pd.read_excel(filename, sheet_name="Capabilities")

    for df in [tech_df, cap_df]:
        df['Start Date'] = pd.to_datetime(df['Start Date'])
        df['End Date'] = pd.to_datetime(df['End Date'])

    tech_df['Layer'] = 'Technology'
    cap_df['Layer'] = 'Capabilities'

    return tech_df, cap_df


def plot_layer(ax, df, title, bar_height):
    df = df.sort_values(by='Impact Rating')
    position_tracker = {}

    for idx, row in df.iterrows():
        label = f"{row['ID']} - {row['Name']}"
        if label not in position_tracker:
            position_tracker[label] = len(position_tracker)

    spacing = 0.35  # reduce to bring bars closer
    total_tracks = len(position_tracker)
    ax.set_ylim(-spacing / 2, (total_tracks - 1) * spacing + spacing / 2)

    bar_positions = {}

    for idx, row in df.iterrows():
        start = row['Start Date']
        end = row['End Date']
        label = f"{row['ID']} - {row['Name']}"
        color = row.get('Color', '#1f77b4')
        y = position_tracker[label] * spacing  # apply spacing

        bar_positions[row['ID']] = {
            'x_start': start,
            'x_center': start + (end - start) / 2,
            'y': y,
            'layer': title
        }

        ax.add_patch(patches.Rectangle(
            (start, y - bar_height / 2),
            end - start,
            bar_height,
            facecolor=color,
            edgecolor='black',
            linewidth=1.2,
            alpha=0.9
        ))

        ax.text(start + (end - start) / 2, y, label,
                ha='center', va='center', fontsize=9, color='white')

    ax.set_yticks([])
    ax.set_ylabel(title, rotation=90, fontsize=11, labelpad=10, weight='bold', va='center')
    ax.grid(True, axis='x', linestyle='--', color='gray', alpha=0.5)

    return bar_positions


def plot_roadmap(tech_df, cap_df):
    fig, (ax1, ax2) = plt.subplots(
        2, 1,
        figsize=(14, 8),
        sharex=True,
        gridspec_kw={'height_ratios': [1, 1], 'hspace': 0}
    )

    overall_start = min(tech_df['Start Date'].min(), cap_df['Start Date'].min())
    overall_end = max(tech_df['End Date'].max(), cap_df['End Date'].max())

    for ax in (ax1, ax2):
        ax.set_xlim([overall_start, overall_end])
        ax.tick_params(axis='x', which='both', bottom=True, top=False, labelbottom=False)
        ax.spines['top'].set_visible(False)
        ax.spines['bottom'].set_linewidth(0.8)
        ax.spines['left'].set_linewidth(0.8)
        ax.spines['right'].set_visible(False)

    ax2.tick_params(axis='x', labelbottom=True)

    BAR_HEIGHT = 0.2
    cap_pos = plot_layer(ax1, cap_df, "Capabilities", BAR_HEIGHT)
    tech_pos = plot_layer(ax2, tech_df, "Technology", BAR_HEIGHT)

    ax2.set_xlabel("Time", fontsize=11)
    fig.suptitle("Roadmap: Capabilities and Technologies", fontsize=16)
    plt.subplots_adjust(left=0.1, right=0.98, top=0.93, bottom=0.08)
    plt.savefig("roadmap.png", dpi=300, bbox_inches='tight')
    plt.show()


# === Main Program ===
if __name__ == "__main__":
    excel_path = "PRM-Data.xlsx"
    tech_df, cap_df = load_roadmap_data(excel_path)
    plot_roadmap(tech_df, cap_df)
