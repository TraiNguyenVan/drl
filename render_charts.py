import os
import matplotlib.pyplot as plt
import numpy as np

# Set font family to ensure Vietnamese unicode support
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['axes.unicode_minus'] = False

def generate_all_charts(students, workspace_dir):
    """
    Generates DRL charts from processed student data:
    1. Score distribution histogram
    2. Rating breakdown donut chart
    3. Criteria average vs maximum score progress bars
    4. Combined dashboard
    """
    charts_dir = os.path.join(workspace_dir, "charts")
    os.makedirs(charts_dir, exist_ok=True)
    
    # 0. Extract raw data
    totals = [s["excel_total"] for s in students]
    ratings = [s["excel_rating"] for s in students]
    
    # TC1 to TC5 averages
    # Each student's excel_tc is [tc1, tc2, tc3, tc4, tc5]
    tc_lists = [s["excel_tc"] for s in students]
    num_students = len(students)
    
    if num_students == 0:
        print("No student data available to render charts.")
        return
        
    avg_tcs = []
    for col in range(5):
        scores = [tc[col] for tc in tc_lists]
        avg_tcs.append(sum(scores) / num_students)
        
    max_tcs = [20, 25, 20, 25, 10]
    tc_labels = [
        "TC1 (Max 20)",
        "TC2 (Max 25)",
        "TC3 (Max 20)",
        "TC4 (Max 25)",
        "TC5 (Max 10)"
    ]
    
    # Rating categories and colors
    rating_order = ["Xuất sắc", "Tốt", "Khá", "Trung bình", "Yếu", "Kém"]
    rating_colors = {
        "Xuất sắc": "#1b5e20",  # Deep green
        "Tốt": "#4caf50",      # Green
        "Khá": "#2196f3",      # Blue
        "Trung bình": "#ffeb3b", # Yellow
        "Yếu": "#ff9800",      # Orange
        "Kém": "#f44336"       # Red
    }
    
    # Count rating categories
    rating_counts = {r: ratings.count(r) for r in rating_order}
    
    # Class Statistics
    class_avg = sum(totals) / num_students
    highest_score = max(totals)
    lowest_score = min(totals)
    std_dev = np.std(totals)
    
    # Styling configurations
    primary_color = "#3182ce"
    bg_color = "#f7fafc"
    grid_color = "#e2e8f0"
    text_color = "#2d3748"
    
    # ----------------------------------------------------
    # CHART 1: Score Distribution
    # ----------------------------------------------------
    fig, ax = plt.subplots(figsize=(8, 5))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')
    
    # Define bins
    bins = np.arange(min(30, int(lowest_score)//10*10), 101, 5)
    counts, edges, bars = ax.hist(totals, bins=bins, color=primary_color, edgecolor='white', alpha=0.85, rwidth=0.9)
    
    # Add values on top of bars
    for bar in bars:
        height = bar.get_height()
        if height > 0:
            ax.annotate(f'{int(height)}',
                        xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 3),  # 3 points vertical offset
                        textcoords="offset points",
                        ha='center', va='bottom', fontsize=9, color=text_color, weight='bold')
            
    # Add average line
    ax.axvline(class_avg, color="#e53e3e", linestyle="--", linewidth=1.5, label=f"Trung bình: {class_avg:.1f}")
    
    ax.set_title("Phân Phối Điểm Rèn Luyện Lớp E25CQCE02-N", fontsize=14, weight='bold', pad=15, color=text_color)
    ax.set_xlabel("Điểm Rèn Luyện (DRL)", fontsize=11, labelpad=10, color=text_color)
    ax.set_ylabel("Số Lượng Sinh Viên", fontsize=11, labelpad=10, color=text_color)
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(grid_color)
    ax.spines['bottom'].set_color(grid_color)
    ax.grid(axis='y', linestyle=':', alpha=0.6, color=grid_color)
    ax.legend(loc="upper left")
    
    plt.tight_layout()
    plt.savefig(os.path.join(charts_dir, "score_distribution.png"), dpi=150, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    
    # ----------------------------------------------------
    # CHART 2: Rating Breakdown (Donut Chart)
    # ----------------------------------------------------
    fig, ax = plt.subplots(figsize=(7, 7))
    fig.patch.set_facecolor('white')
    
    # Filter ratings with count > 0 for pie wedges
    active_ratings = [r for r in rating_order if rating_counts[r] > 0]
    active_counts = [rating_counts[r] for r in active_ratings]
    active_colors = [rating_colors[r] for r in active_ratings]
    
    # Generate Donut Chart
    wedges, texts, autotexts = ax.pie(
        active_counts, 
        labels=active_ratings, 
        colors=active_colors,
        autopct=lambda pct: f"{pct:.1f}%\n({int(round(pct*num_students/100))})",
        startangle=90,
        pctdistance=0.75,
        wedgeprops=dict(width=0.4, edgecolor='white', linewidth=2),
        textprops=dict(color=text_color)
    )
    
    # Style the text inside slices
    for autotext in autotexts:
        autotext.set_fontsize(10)
        autotext.set_weight('bold')
        
    for text in texts:
        text.set_fontsize(11)
        text.set_weight('bold')
        
    ax.set_title("Xếp Loại Kết Quả Rèn Luyện SV", fontsize=14, weight='bold', pad=15, color=text_color)
    
    # Add summary in the center of donut
    ax.text(0, 0, f"Tổng số\n{num_students}\nSV", ha='center', va='center', fontsize=14, weight='bold', color=text_color)
    
    plt.tight_layout()
    plt.savefig(os.path.join(charts_dir, "rating_distribution.png"), dpi=150, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    
    # ----------------------------------------------------
    # CHART 3: Criteria Averages vs Max Score (Progress Bar Style)
    # ----------------------------------------------------
    fig, ax = plt.subplots(figsize=(9, 5))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')
    
    y_pos = np.arange(5)
    
    # Draw background max points bars
    ax.barh(y_pos, max_tcs, color="#edf2f7", edgecolor=grid_color, label="Điểm tối đa", height=0.55)
    
    # Draw average score bars
    avg_bars = ax.barh(y_pos, avg_tcs, color="#3182ce", edgecolor="#2b6cb0", label="Điểm trung bình", height=0.55)
    
    # Add text labels with values and percentage
    for i, bar in enumerate(avg_bars):
        width = bar.get_width()
        pct = (width / max_tcs[i]) * 100
        ax.text(width + 0.3, bar.get_y() + bar.get_height()/2, 
                f"{width:.1f} / {max_tcs[i]} ({pct:.1f}%)", 
                va='center', ha='left', fontsize=10, weight='bold', color=text_color)
        
    ax.set_yticks(y_pos)
    ax.set_yticklabels(tc_labels, fontsize=11, weight='bold', color=text_color)
    ax.set_title("Điểm Trung Bình Theo Từng Tiêu Chí Đánh Giá", fontsize=14, weight='bold', pad=15, color=text_color)
    ax.set_xlabel("Thang Điểm", fontsize=11, color=text_color)
    ax.set_xlim(0, 28) # Extra space for text labels
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(grid_color)
    ax.spines['bottom'].set_color(grid_color)
    ax.grid(axis='x', linestyle=':', alpha=0.6, color=grid_color)
    ax.legend(loc="lower right")
    
    plt.tight_layout()
    plt.savefig(os.path.join(charts_dir, "criteria_averages.png"), dpi=150, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    
    # ----------------------------------------------------
    # CHART 4: Dashboard (Combined 2x2 Grid)
    # ----------------------------------------------------
    fig = plt.figure(figsize=(15, 11))
    fig.patch.set_facecolor('#f8fafc') # Light grey dashboard background
    
    # Title of the dashboard
    fig.suptitle("BÁO CÁO THỐNG KÊ KẾT QUẢ RÈN LUYỆN LỚP E25CQCE02-N", fontsize=18, weight='bold', color="#1a365d", y=0.96)
    
    # Top-Left: Score Distribution (Histogram)
    ax1 = fig.add_subplot(2, 2, 1)
    ax1.set_facecolor('white')
    counts, edges, bars = ax1.hist(totals, bins=bins, color="#4299e1", edgecolor='white', alpha=0.9, rwidth=0.85)
    for bar in bars:
        height = bar.get_height()
        if height > 0:
            ax1.annotate(f'{int(height)}',
                         xy=(bar.get_x() + bar.get_width() / 2, height),
                         xytext=(0, 2),
                         textcoords="offset points",
                         ha='center', va='bottom', fontsize=8, color=text_color, weight='bold')
    ax1.axvline(class_avg, color="#e53e3e", linestyle="--", linewidth=1.5, label=f"TB: {class_avg:.1f}")
    ax1.set_title("Phân Phối Điểm Rèn Luyện (DRL)", fontsize=12, weight='bold', pad=10, color="#1a365d")
    ax1.set_xlabel("Điểm Rèn Luyện", fontsize=9, color=text_color)
    ax1.set_ylabel("Số Lượng Sinh Viên", fontsize=9, color=text_color)
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)
    ax1.spines['left'].set_color(grid_color)
    ax1.spines['bottom'].set_color(grid_color)
    ax1.grid(axis='y', linestyle=':', alpha=0.6, color=grid_color)
    ax1.legend(loc="upper left", prop={'size': 8})
    
    # Top-Right: Rating Breakdown (Donut)
    ax2 = fig.add_subplot(2, 2, 2)
    ax2.set_facecolor('white')
    wedges, texts, autotexts = ax2.pie(
        active_counts, 
        labels=active_ratings, 
        colors=active_colors,
        autopct=lambda pct: f"{pct:.1f}%\n({int(round(pct*num_students/100))})",
        startangle=90,
        pctdistance=0.75,
        wedgeprops=dict(width=0.4, edgecolor='#f8fafc', linewidth=2),
        textprops=dict(color=text_color)
    )
    for autotext in autotexts:
        autotext.set_fontsize(8)
        autotext.set_weight('bold')
    for text in texts:
        text.set_fontsize(9)
        text.set_weight('bold')
    ax2.set_title("Tỉ Lệ Xếp Loại Rèn Luyện", fontsize=12, weight='bold', pad=10, color="#1a365d")
    ax2.text(0, 0, f"Tổng số\n{num_students}\nSV", ha='center', va='center', fontsize=12, weight='bold', color=text_color)
    
    # Bottom-Left: Criteria Averages vs Max Score
    ax3 = fig.add_subplot(2, 2, 3)
    ax3.set_facecolor('white')
    ax3.barh(y_pos, max_tcs, color="#edf2f7", edgecolor=grid_color, height=0.55)
    avg_bars_db = ax3.barh(y_pos, avg_tcs, color="#3182ce", edgecolor="#2b6cb0", height=0.55)
    for i, bar in enumerate(avg_bars_db):
        width = bar.get_width()
        pct = (width / max_tcs[i]) * 100
        ax3.text(width + 0.3, bar.get_y() + bar.get_height()/2, 
                 f"{width:.1f}/{max_tcs[i]} ({pct:.1f}%)", 
                 va='center', ha='left', fontsize=8, weight='bold', color=text_color)
    ax3.set_yticks(y_pos)
    ax3.set_yticklabels(tc_labels, fontsize=9, weight='bold', color=text_color)
    ax3.set_title("Điểm Trung Bình Theo Tiêu Chí", fontsize=12, weight='bold', pad=10, color="#1a365d")
    ax3.set_xlabel("Thang Điểm", fontsize=9, color=text_color)
    ax3.set_xlim(0, 28)
    ax3.spines['top'].set_visible(False)
    ax3.spines['right'].set_visible(False)
    ax3.spines['left'].set_color(grid_color)
    ax3.spines['bottom'].set_color(grid_color)
    ax3.grid(axis='x', linestyle=':', alpha=0.6, color=grid_color)
    
    # Bottom-Right: Statistics Scoreboard (Text Box)
    ax4 = fig.add_subplot(2, 2, 4)
    ax4.axis('off') # Hide axes entirely for the text card
    
    # Create background bounding box
    bbox_props = dict(boxstyle="round,pad=1.2", facecolor="white", edgecolor="#cbd5e0", lw=1.5)
    
    # Statistics Text compilation
    stats_text = (
        f"   BẢNG SỐ LIỆU THỐNG KÊ LỚP HỌC\n"
        f"  -----------------------------------------\n\n"
        f"  • Tổng số sinh viên:      {num_students} SV\n"
        f"  • Điểm trung bình cả lớp:  {class_avg:.2f} / 100\n"
        f"  • Điểm cao nhất lớp:      {highest_score:.1f} / 100\n"
        f"  • Điểm thấp nhất lớp:     {lowest_score:.1f} / 100\n"
        f"  • Độ lệch chuẩn (Std Dev): {std_dev:.2f}\n\n"
        f"  THỐNG KÊ XẾP LOẠI CHI TIẾT:\n"
    )
    
    for r in rating_order:
        count = rating_counts[r]
        pct = (count / num_students) * 100
        stats_text += f"   - Loại {r:<10}: {count:>2} SV ({pct:>5.1f}%)\n"
        
    ax4.text(0.05, 0.05, stats_text, fontsize=10.5, color="#2d3748", 
             verticalalignment='bottom', horizontalalignment='left', bbox=bbox_props)
    
    plt.subplots_adjust(top=0.88, bottom=0.08, left=0.08, right=0.95, hspace=0.25, wspace=0.22)
    plt.savefig(os.path.join(charts_dir, "dashboard.png"), dpi=200, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    
    # Generate the detailed sub-criteria line chart
    generate_subcriteria_linechart(students, workspace_dir)
    
    print(f"Successfully generated all DRL charts in: {charts_dir}")

def generate_subcriteria_linechart(students, workspace_dir):
    """
    Generates a large line chart showing the DRL scores of every student
    across all 23 lowest-level DRL criteria. To make the chart readable,
    the scores are normalized to a percentage (0% to 100%), and each student
    is represented by a line colored according to their total DRL score.
    """
    charts_dir = os.path.join(workspace_dir, "charts")
    os.makedirs(charts_dir, exist_ok=True)
    
    # 1. Sort students by total DRL score descending
    students_sorted = sorted(students, key=lambda s: s["excel_total"], reverse=True)
    
    subcriteria_keys = [
        "1.1", "1.2", "1.3", "1.4", "1.5",
        "2.1", "2.2", "2.3",
        "3.1", "3.2", "3.3", "3.4", "3.5",
        "4.1", "4.2", "4.3", "4.4", "4.5", "4.6",
        "5.1", "5.2", "5.3"
    ]
    
    subcriteria_max = {
        "1.1": 3.0, "1.2": 10.0, "1.3": 4.0, "1.4": 2.0, "1.5": 1.0,
        "2.1": 15.0, "2.2": 5.0, "2.3": 5.0,
        "3.1": 10.0, "3.2": 4.0, "3.3": 3.0, "3.4": 3.0, "3.5": 0.0,
        "4.1": 8.0, "4.2": 5.0, "4.3": 5.0, "4.4": 5.0, "4.5": 2.0, "4.6": 0.0,
        "5.1": 4.0, "5.2": 3.0, "5.3": 3.0
    }
    
    # 2. Extract values and calculate normalized percentages
    x_indices = np.arange(len(subcriteria_keys))
    
    fig, ax = plt.subplots(figsize=(22, 10))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')
    
    import matplotlib.colors as mcolors
    import matplotlib.cm as cm
    
    min_total = min([s["excel_total"] for s in students_sorted])
    max_total = max([s["excel_total"] for s in students_sorted])
    norm = mcolors.Normalize(vmin=min_total, vmax=max_total)
    
    # Use viridis colormap to color lines by DRL total score
    colormap = plt.get_cmap('viridis')
    
    all_norms = []
    
    # Plot a line for each student
    for s in students_sorted:
        student_norms = []
        for key in subcriteria_keys:
            score = s.get(f"sub_{key}", 0.0)
            max_val = subcriteria_max[key]
            if max_val == 0.0:  # Penalty-only criteria
                pct = 100.0 if score == 0.0 else 0.0
            else:
                pct = (score / max_val) * 100.0
            student_norms.append(pct)
            
        all_norms.append(student_norms)
        
        # Color line by student's total score
        color = colormap(norm(s["excel_total"]))
        ax.plot(x_indices, student_norms, marker='o', markersize=4, 
                color=color, alpha=0.3, linewidth=1.2)
        
    all_norms = np.array(all_norms)
    avg_norms = np.mean(all_norms, axis=0)
    
    # Plot Class Average line prominently
    ax.plot(x_indices, avg_norms, marker='D', markersize=8, 
            color='#e53e3e', linewidth=3.5, label='Trung bình cả lớp', zorder=100)
    
    # Annotate Class Average line data points
    for idx, val in enumerate(avg_norms):
        ax.annotate(f"{val:.1f}%",
                    xy=(idx, val),
                    xytext=(0, 10),
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=9, color='#e53e3e', weight='bold',
                    bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="#e53e3e", alpha=0.9, lw=1))
        
    # Y-axis styling
    ax.set_ylim(-5, 108)
    ax.set_ylabel("Tỉ Lệ Đạt Điểm Tối Đa (%)", fontsize=12, labelpad=10, color="#2d3748")
    ax.set_yticks(np.arange(0, 101, 10))
    ax.set_yticklabels([f"{y}%" for y in np.arange(0, 101, 10)], fontsize=10, color="#2d3748")
    
    # X-axis styling
    ax.set_xticks(x_indices)
    ax.set_xticklabels(subcriteria_keys, fontsize=11, weight='bold', color="#2d3748")
    ax.set_xlabel("Tiêu Chí Đánh Giá Con (Từ 1.1 đến 5.3)", fontsize=12, labelpad=15, color="#2d3748")
    
    # Vertical grid category boundary lines
    category_splits = [5, 8, 13, 19]
    for split in category_splits:
        ax.axvline(split - 0.5, color='#cbd5e0', linestyle='--', alpha=0.8, linewidth=1.2)
        
    # Main criteria tags
    ax.text(2.0, 104, "TIÊU CHÍ 1", ha='center', fontsize=11, weight='bold', color="#1a365d")
    ax.text(6.5, 104, "TIÊU CHÍ 2", ha='center', fontsize=11, weight='bold', color="#1a365d")
    ax.text(10.5, 104, "TIÊU CHÍ 3", ha='center', fontsize=11, weight='bold', color="#1a365d")
    ax.text(16.0, 104, "TIÊU CHÍ 4", ha='center', fontsize=11, weight='bold', color="#1a365d")
    ax.text(20.5, 104, "TIÊU CHÍ 5", ha='center', fontsize=11, weight='bold', color="#1a365d")
    
    # Clean up spines and grids
    ax.grid(axis='y', linestyle=':', alpha=0.6, color="#cbd5e0")
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#cbd5e0')
    ax.spines['bottom'].set_color('#cbd5e0')
    
    # Chart title
    ax.set_title("BIỂU ĐỒ ĐIỂM RÈN LUYỆN CHI TIẾT CỦA TỪNG SINH VIÊN THEO TIÊU CHÍ CON", 
                 fontsize=16, weight='bold', pad=30, color="#1a365d")
    
    # Gradient scale colorbar
    sm = cm.ScalarMappable(cmap=colormap, norm=norm)
    sm.set_array([])
    cbar = fig.colorbar(sm, ax=ax, pad=0.015, aspect=30)
    cbar.set_label('Tổng điểm rèn luyện (đ)', rotation=270, labelpad=15, fontsize=11, color="#2d3748")
    cbar.ax.tick_params(labelsize=10)
    cbar.outline.set_visible(False)
    
    ax.legend(loc="lower left", prop={'size': 11, 'weight': 'bold'}, frameon=True, facecolor='white', edgecolor='#cbd5e0')
    
    plt.tight_layout()
    plt.savefig(os.path.join(charts_dir, "student_subcriteria_linechart.png"), dpi=200, bbox_inches='tight')
    plt.close()

