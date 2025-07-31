import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches

def load_roadmap_data(filename):
    tech_df = pd.read_excel(filename, sheet_name="Technology")
    cap_df = pd.read_excel(filename, sheet_name="Capabilities")
    roadmap_df = pd.read_excel(filename, sheet_name="Roadmap")
    milestone_df = pd.read_excel(filename, sheet_name="Milestones")

    for df in [tech_df, cap_df, roadmap_df, milestone_df]:
        if 'Start Date' in df.columns:
            df['Start Date'] = pd.to_datetime(df['Start Date'])
        if 'End Date' in df.columns:
            df['End Date'] = pd.to_datetime(df['End Date'])
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'])

    return tech_df, cap_df, roadmap_df, milestone_df

def add_milestone_lines(ax, milestone_df):
    for _, row in milestone_df.iterrows():
        date = row['Date']
        milestone_id = row['ID']

        # Vertical milestone line (only in top chart)
        ax.vlines(date,ymin=-4, ymax=9.5, color='darkblue', linestyle='-', linewidth=1.5, alpha=0.3)

        # Horizontal text label centered below line, in data coordinates
        ax.text(date, -5, milestone_id,  # slightly below y-axis min
                rotation=0,  # horizontal text
                ha='center', va='top',
                fontsize=8, color='darkblue', weight='bold')


def plot_risk_section(ax, roadmap_df, milestone_df):
    max_budget = roadmap_df['Exp. Budget'].max()
    risk_points = []

    for _, row in roadmap_df.iterrows():
        start_date = row['Start Date']
        risk = row['Risk Rating']
        budget = row['Exp. Budget']
        budget_risk = row['Budget Risk']
        timeline_risk = row['Timeline Risk']
        activity_id = row['ID']

        ax.plot(start_date, risk, 'o', markersize=4, color='red')
        ax.text(start_date, risk + 1.2, activity_id, ha='center', va='bottom',
                fontsize=8, style='italic', color='black')

        line_length = (budget / max_budget) * 9
        budget_bottom = risk - line_length
        ax.vlines(start_date, risk, budget_bottom, colors='red', linewidth=1)

        rect_height = line_length + budget_risk
        rect_width_days = timeline_risk * 2 * 30.44
        rect_bottom = budget_bottom - (budget_risk / 2)
        rect_left = start_date - pd.Timedelta(days=rect_width_days / 2)

        ax.add_patch(patches.Rectangle(
            (rect_left, rect_bottom),
            pd.to_timedelta(rect_width_days, unit='D'),
            rect_height,
            linewidth=0.5,
            edgecolor='black',
            facecolor='lightblue',
            alpha=0.5
        ))

        risk_points.append((start_date, risk))

    risk_points.sort(key=lambda x: x[0])
    dates, risks = zip(*risk_points)
    ax.plot(dates, risks, color='black', linewidth=1, linestyle='--', alpha=0.6)

    add_milestone_lines(ax, milestone_df)

    ax.set_ylabel("Risk", rotation=90, fontsize=10, weight='bold')
    ax.set_ylim(-5, 10)
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.set_yticks([])
    ax.grid(True, axis='x', linestyle='--', color='gray', alpha=0.5)

def plot_layer(ax, df, title, bar_height, fixed_slots=12):
    df = df.sort_values(by='Impact Rating').reset_index(drop=True)
    position_tracker = {}

    for i, (_, row) in enumerate(df.iterrows()):
        label = f"{row['ID']} - {row['Name']}"
        position_tracker[label] = i

    for _, row in df.iterrows():
        start = row['Start Date']
        end = row['End Date']
        label = f"{row['ID']} - {row['Name']}"
        color = row.get('Color', '#1f77b4')
        y = position_tracker[label]

        ax.add_patch(patches.Rectangle(
            (start, y - bar_height / 2),
            end - start,
            bar_height,
            facecolor=color,
            edgecolor='black',
            linewidth=0.5,
            alpha=0.7
        ))

        ax.text(start + (end - start) / 2, y, label,
                ha='center', va='center', fontsize=9, color='white')

    ax.set_yticks([])
    ax.set_ylabel(title, rotation=90, fontsize=11, labelpad=10, weight='bold', va='center')
    ax.set_ylim(-0.5, fixed_slots - 0.5)
    ax.grid(True, axis='x', linestyle='--', color='gray', alpha=0.5)

def plot_combined_roadmap(tech_df, cap_df, roadmap_df, milestone_df):
    fig, (ax0, ax1, ax2) = plt.subplots(
        3, 1,
        figsize=(14, 9),
        sharex=True,
        gridspec_kw={'height_ratios': [1.5, .5, .5], 'hspace': 0}
    )

    overall_start = min(
        tech_df['Start Date'].min(),
        cap_df['Start Date'].min(),
        roadmap_df['Start Date'].min()
    )
    overall_end = max(
        tech_df['End Date'].max(),
        cap_df['End Date'].max(),
        roadmap_df['End Date'].max()
    )

    for ax in (ax0, ax1, ax2):
        ax.set_xlim([overall_start, overall_end])
        ax.tick_params(axis='x', which='both', bottom=True, top=False, labelbottom=False)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

    ax2.tick_params(axis='x', labelbottom=True)

    BAR_HEIGHT = 1
    FIXED_ROWS = 6

    plot_risk_section(ax0, roadmap_df, milestone_df)
    plot_layer(ax1, cap_df, "Capabilities", BAR_HEIGHT, FIXED_ROWS)
    plot_layer(ax2, tech_df, "Technology", BAR_HEIGHT, FIXED_ROWS)

    ax2.set_xlabel("Time", fontsize=11)
    fig.suptitle("Full Product Roadmap", fontsize=16)
    plt.subplots_adjust(left=0.1, right=0.98, top=0.93, bottom=0.08)
    plt.savefig("roadmap_full.png", dpi=300, bbox_inches='tight')


# === Main Program ===
if __name__ == "__main__":
    excel_path = "PRM-Data.xlsx"
    tech_df, cap_df, roadmap_df, milestone_df = load_roadmap_data(excel_path)
    plot_combined_roadmap(tech_df, cap_df, roadmap_df, milestone_df)
    plt.show()
