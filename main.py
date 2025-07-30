import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches

def load_roadmap_data(filename):
    tech_df = pd.read_excel(filename, sheet_name="Technology")
    cap_df = pd.read_excel(filename, sheet_name="Capabilities")

    for df in [tech_df, cap_df]:
        df['Start Date'] = pd.to_datetime(df['Start Date'])
        df['End Date'] = pd.to_datetime(df['End Date'])

    return tech_df, cap_df

def plot_layer(ax, df, title):
    df = df.sort_values(by='Impact Rating')  # Lower ratings drawn first
    position_tracker = {}

    # Assign vertical positions
    for idx, row in df.iterrows():
        label = f"{row['ID']} - {row['Name']}"
        if label not in position_tracker:
            position_tracker[label] = len(position_tracker)

    total_tracks = len(position_tracker)
    ax.set_ylim(-0.5, total_tracks - 0.2)  # Ensure all bars fit

    for idx, row in df.iterrows():
        start = row['Start Date']
        end = row['End Date']
        label = f"{row['ID']} - {row['Name']}"
        color = row.get('Color', '#1f77b4')
        y = position_tracker[label]

        ax.add_patch(patches.Rectangle(
            (start, y),
            end - start,
            0.8,
            facecolor=color,  # fixed warning
            edgecolor='black',
            linewidth=1.2,
            alpha=0.9
        ))

        ax.text(start + (end - start) / 2, y + 0.4, label,
                ha='center', va='center', fontsize=9, color='white')

    ax.set_yticks([])
    ax.set_ylabel(title)
    ax.grid(True, axis='x')

def plot_roadmap(tech_df, cap_df):
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 8), sharex=True, gridspec_kw={'height_ratios': [1, 1]})

    overall_start = min(tech_df['Start Date'].min(), cap_df['Start Date'].min())
    overall_end = max(tech_df['End Date'].max(), cap_df['End Date'].max())

    ax1.set_xlim([overall_start, overall_end])
    ax2.set_xlim([overall_start, overall_end])

    plot_layer(ax1, cap_df, "Capabilities")
    plot_layer(ax2, tech_df, "Technology")

    ax2.set_xlabel("Time")
    fig.suptitle("Roadmap: Capabilities and Technologies", fontsize=16)

    # Clean layout and save image
    plt.subplots_adjust(hspace=0.3)
    plt.savefig("roadmap.png", dpi=300, bbox_inches='tight')
    # plt.show()  # Optional: enable for GUI viewing

# === Main Program ===
if __name__ == "__main__":
    excel_path = "PRM-Data.xlsx"
    tech_df, cap_df = load_roadmap_data(excel_path)
    plot_roadmap(tech_df, cap_df)
