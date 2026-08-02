# render_html.py
import os
import json
import urllib.request

def get_script_content(url):
    try:
        print(f"Fetching local copy of {url} for offline compatibility...")
        req = urllib.request.Request(url, headers={'User-Agent': 'Antigravity-Dashboard-Compiler/1.0'})
        with urllib.request.urlopen(req, timeout=5) as response:
            return response.read().decode('utf-8')
    except Exception as e:
        print(f"Warning: Could not fetch {url} ({e}). Falling back to CDN link.")
        return None

def generate_html_dashboard(students, workspace_dir, chart_base64_images=None):
    """
    Generates an interactive, zero-dependency HTML dashboard file 
    incorporating Chart.js for data visualization and a modern, 
    glassmorphic corporate light design.
    """
    dashboard_path = os.path.join(workspace_dir, "assets", "drl_dashboard.html")
    students_json = json.dumps(students, ensure_ascii=False)
    
    if not chart_base64_images:
        chart_base64_images = {}
        
    img_score_dist = chart_base64_images.get("score_distribution", "")
    img_rating_dist = chart_base64_images.get("rating_distribution", "")
    img_criteria_avg = chart_base64_images.get("criteria_averages", "")
    img_dashboard = chart_base64_images.get("dashboard", "")
    img_linechart = chart_base64_images.get("student_subcriteria_linechart", "")
    
    img_dashboard_tag = f'<img src="{img_dashboard}" alt="Báo cáo thống kê tổng hợp">' if img_dashboard else '<p style="color: var(--text-muted); padding: 40px 0;">Không tìm thấy ảnh báo cáo tổng hợp</p>'
    img_linechart_tag = f'<img src="{img_linechart}" alt="Biểu đồ chi tiết tiêu chí con">' if img_linechart else '<p style="color: var(--text-muted); padding: 40px 0;">Không tìm thấy ảnh biểu đồ tiêu chí con</p>'
    img_score_dist_tag = f'<img src="{img_score_dist}" alt="Phân phối điểm số">' if img_score_dist else '<p style="color: var(--text-muted); padding: 20px 0;">N/A</p>'
    img_rating_dist_tag = f'<img src="{img_rating_dist}" alt="Tỉ lệ xếp loại">' if img_rating_dist else '<p style="color: var(--text-muted); padding: 20px 0;">N/A</p>'
    img_criteria_avg_tag = f'<img src="{img_criteria_avg}" alt="Trung bình tiêu chí">' if img_criteria_avg else '<p style="color: var(--text-muted); padding: 20px 0;">N/A</p>'

    # Fetch scripts for inlining (offline support)
    chartjs_url = "https://cdn.jsdelivr.net/npm/chart.js"
    hammer_url = "https://cdn.jsdelivr.net/npm/hammerjs@2.0.8/hammer.min.js"
    zoom_url = "https://cdn.jsdelivr.net/npm/chartjs-plugin-zoom@1.2.1/dist/chartjs-plugin-zoom.min.js"
    
    chartjs_content = get_script_content(chartjs_url)
    hammer_content = get_script_content(hammer_url)
    zoom_content = get_script_content(zoom_url)
    
    chartjs_tag = f"<script>{chartjs_content}</script>" if chartjs_content else f'<script src="{chartjs_url}"></script>'
    hammer_tag = f"<script>{hammer_content}</script>" if hammer_content else f'<script src="{hammer_url}"></script>'
    zoom_tag = f"<script>{zoom_content}</script>" if zoom_content else f'<script src="{zoom_url}"></script>'

    html_content = f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Báo cáo Kết quả Rèn luyện E25CQCE02-N</title>
    
    <!-- Load Chart.js, Hammer.js (for panning/pinching), Zoom plugin and Google Fonts -->
    {chartjs_tag}
    {hammer_tag}
    {zoom_tag}
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    
    <style>
        /* Modern Design System - Corporate Light with Antigravity Details */
        :root {{
            --bg-primary: #f8fafc;
            --text-main: #0f172a;
            --text-muted: #64748b;
            --accent-blue: #2563eb;
            --accent-blue-hover: #1d4ed8;
            --card-bg: rgba(255, 255, 255, 0.85);
            --card-border: rgba(255, 255, 255, 0.6);
            --shadow-soft: 0 10px 30px -10px rgba(15, 23, 42, 0.06), 0 1px 3px rgba(15, 23, 42, 0.02);
            --shadow-hover: 0 20px 40px -15px rgba(15, 23, 42, 0.12), 0 3px 6px rgba(15, 23, 42, 0.04);
            --transition-smooth: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        }}

        * {{
            box-sizing: border-box;
            font-family: 'Inter', sans-serif;
        }}

        body {{
            background-color: var(--bg-primary);
            background-image: 
                radial-gradient(at 0% 0%, rgba(219, 234, 254, 0.3) 0, transparent 50%),
                radial-gradient(at 100% 100%, rgba(224, 242, 254, 0.3) 0, transparent 50%);
            background-attachment: fixed;
            color: var(--text-main);
            margin: 0;
            padding: 30px 40px;
            min-height: 100vh;
        }}

        h1, h2, h3, h4 {{
            font-family: 'Outfit', sans-serif;
            font-weight: 700;
            margin: 0;
        }}

        /* Header Layout */
        .dashboard-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 30px;
        }}

        .header-title h1 {{
            font-size: 28px;
            color: #1e293b;
            letter-spacing: -0.5px;
        }}

        .header-title p {{
            margin: 5px 0 0 0;
            color: var(--text-muted);
            font-size: 14px;
        }}

        .class-badge {{
            background: linear-gradient(135deg, #1e3a8a, #2563eb);
            color: white;
            padding: 8px 16px;
            border-radius: 9999px;
            font-family: 'Outfit', sans-serif;
            font-weight: 600;
            font-size: 14px;
            box-shadow: 0 4px 10px rgba(37, 99, 235, 0.2);
        }}

        /* KPI Cards Grid */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}

        .kpi-card {{
            background: var(--card-bg);
            backdrop-filter: blur(16px);
            border: 1px solid var(--card-border);
            border-radius: 16px;
            padding: 20px;
            box-shadow: var(--shadow-soft);
            transition: var(--transition-smooth);
            display: flex;
            flex-direction: column;
            will-change: transform;
        }}

        .kpi-card:hover {{
            transform: translateY(-4px);
            box-shadow: var(--shadow-hover);
        }}

        .kpi-title {{
            font-size: 13px;
            color: var(--text-muted);
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        .kpi-value {{
            font-family: 'Outfit', sans-serif;
            font-size: 32px;
            font-weight: 800;
            color: #1e293b;
            margin-top: 10px;
        }}

        .kpi-subtext {{
            font-size: 12px;
            color: var(--text-muted);
            margin-top: 5px;
        }}

        /* Filters Section */
        .filter-bar {{
            background: var(--card-bg);
            backdrop-filter: blur(16px);
            border: 1px solid var(--card-border);
            border-radius: 16px;
            padding: 16px 24px;
            box-shadow: var(--shadow-soft);
            display: flex;
            gap: 20px;
            align-items: center;
            margin-bottom: 30px;
        }}

        .filter-group {{
            display: flex;
            flex-direction: column;
            gap: 6px;
            flex: 1;
        }}

        .filter-label {{
            font-size: 12px;
            font-weight: 600;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        .search-input, .select-input {{
            width: 100%;
            height: 42px;
            padding: 0 16px;
            border-radius: 10px;
            border: 1px solid #cbd5e1;
            background-color: white;
            color: var(--text-main);
            font-size: 14px;
            outline: none;
            transition: var(--transition-smooth);
        }}

        .search-input:focus, .select-input:focus {{
            border-color: var(--accent-blue);
            box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15);
        }}

        .select-input {{
            cursor: pointer;
            appearance: none;
            background-image: url("data:image/svg+xml;charset=UTF-8,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23475569' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='6 9 12 15 18 9'%3E%3C/polyline%3E%3C/svg%3E");
            background-repeat: no-repeat;
            background-position: right 16px center;
            background-size: 16px;
            padding-right: 40px;
        }}

        /* Dashboard Grid Layout */
        .chart-grid {{
            display: grid;
            grid-template-columns: repeat(12, 1fr);
            gap: 30px;
            margin-bottom: 30px;
        }}

        .chart-card {{
            background: var(--card-bg);
            backdrop-filter: blur(16px);
            border: 1px solid var(--card-border);
            border-radius: 20px;
            padding: 24px;
            box-shadow: var(--shadow-soft);
            transition: var(--transition-smooth);
            will-change: transform;
        }}

        .chart-card:hover {{
            transform: translateY(-2px);
            box-shadow: var(--shadow-hover);
        }}

        .chart-card.col-6 {{ grid-column: span 6; }}
        .chart-card.col-4 {{ grid-column: span 4; }}
        .chart-card.col-8 {{ grid-column: span 8; }}
        .chart-card.col-12 {{ grid-column: span 12; }}

        .card-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
        }}

        .card-title {{
            display: flex;
            flex-direction: column;
        }}

        .card-title h3 {{
            font-size: 18px;
            color: #1e293b;
        }}

        .card-title p {{
            margin: 3px 0 0 0;
            font-size: 12px;
            color: var(--text-muted);
        }}

        .chart-container {{
            position: relative;
            width: 100%;
            height: 320px;
        }}
        
        .chart-container-large {{
            position: relative;
            width: 100%;
            height: 480px;
        }}

        /* Student Scorecard Card Detail Section */
        .scorecard-flex {{
            display: flex;
            gap: 24px;
        }}

        .student-profile-pane {{
            flex: 1.2;
            border-right: 1px solid #e2e8f0;
            padding-right: 24px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }}

        .profile-list-select {{
            margin-bottom: 15px;
        }}

        .profile-details {{
            display: flex;
            flex-direction: column;
            gap: 8px;
        }}

        .detail-row {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10px 0;
            border-bottom: 1px dashed #e2e8f0;
        }}

        .detail-label {{
            font-size: 13px;
            color: var(--text-muted);
            font-weight: 500;
        }}

        .detail-value {{
            font-size: 14px;
            font-weight: 600;
            color: #1e293b;
        }}

        .rating-badge {{
            padding: 4px 10px;
            border-radius: 9999px;
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
        }}

        /* Rating specific colors */
        .badge-xuat-sac {{ background-color: #dcfce7; color: #166534; }}
        .badge-tot {{ background-color: #e0f2fe; color: #0369a1; }}
        .badge-kha {{ background-color: #fef9c3; color: #854d0e; }}
        .badge-trung-binh {{ background-color: #ffedd5; color: #9a3412; }}
        .badge-yeu {{ background-color: #fee2e2; color: #991b1b; }}
        .badge-kem {{ background-color: #f1f5f9; color: #475569; }}

        .student-chart-pane {{
            flex: 1.5;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 250px;
        }}

        /* Table Card Section */
        .table-card {{
            background: var(--card-bg);
            backdrop-filter: blur(16px);
            border: 1px solid var(--card-border);
            border-radius: 20px;
            padding: 24px;
            box-shadow: var(--shadow-soft);
            margin-bottom: 20px;
            overflow: hidden;
        }}

        .table-wrapper {{
            overflow-x: auto;
            max-height: 450px;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
        }}

        th {{
            background-color: #f1f5f9;
            color: var(--text-muted);
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 14px 16px;
            position: sticky;
            top: 0;
            z-index: 10;
            user-select: none;
            transition: var(--transition-smooth);
        }}

        th:hover {{
            background-color: #e2e8f0;
            color: var(--text-main);
        }}

        td {{
            padding: 12px 16px;
            font-size: 13px;
            border-bottom: 1px solid #e2e8f0;
            color: #334155;
            transition: var(--transition-smooth);
        }}

        tr:hover td {{
            background-color: rgba(241, 245, 249, 0.5);
            cursor: pointer;
        }}

        tr.selected-row td {{
            background-color: #eff6ff;
            border-bottom-color: #bfdbfe;
        }}

        .bold-cell {{
            font-weight: 600;
            color: #1e293b;
        }}

        /* Switch styles */
        .switch-container {{
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 13px;
            font-weight: 600;
            color: var(--text-muted);
            cursor: pointer;
        }}

        .switch-container input {{
            cursor: pointer;
        }}

        .btn-reset-zoom {{
            background: white;
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            padding: 6px 14px;
            font-size: 12px;
            font-weight: 600;
            color: var(--text-muted);
            cursor: pointer;
            transition: var(--transition-smooth);
            outline: none;
        }}

        .btn-reset-zoom:hover {{
            background-color: #f1f5f9;
            color: var(--text-main);
            border-color: #94a3b8;
        }}

        /* Page Footer styling */
        .dashboard-footer {{
            text-align: center;
            color: var(--text-muted);
            font-size: 12px;
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid #cbd5e1;
        }}

        /* View Selector Styles */
        .view-selector-container {{
            display: flex;
            justify-content: center;
            margin-bottom: 25px;
        }}
        
        .view-selector {{
            background: rgba(226, 232, 240, 0.6);
            backdrop-filter: blur(8px);
            border: 1px solid rgba(255, 255, 255, 0.6);
            padding: 4px;
            border-radius: 12px;
            display: inline-flex;
            gap: 4px;
            box-shadow: var(--shadow-soft);
        }}
        
        .view-btn {{
            background: transparent;
            border: none;
            padding: 8px 20px;
            font-size: 14px;
            font-weight: 600;
            color: var(--text-muted);
            border-radius: 8px;
            cursor: pointer;
            transition: var(--transition-smooth);
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        
        .view-btn:hover {{
            color: var(--text-main);
        }}
        
        .view-btn.active {{
            background: white;
            color: var(--accent-blue);
            box-shadow: 0 4px 12px rgba(15, 23, 42, 0.05);
        }}
        
        .view-section {{
            display: block;
        }}
        
        .view-section.hidden {{
            display: none;
        }}
        
        /* Static report images grid */
        .static-grid {{
            display: grid;
            grid-template-columns: repeat(12, 1fr);
            gap: 30px;
            margin-bottom: 30px;
        }}
        
        .static-card {{
            background: var(--card-bg);
            backdrop-filter: blur(16px);
            border: 1px solid var(--card-border);
            border-radius: 20px;
            padding: 24px;
            box-shadow: var(--shadow-soft);
            display: flex;
            flex-direction: column;
            align-items: center;
            transition: var(--transition-smooth);
        }}
        
        .static-card:hover {{
            transform: translateY(-2px);
            box-shadow: var(--shadow-hover);
        }}
        
        .static-card img {{
            width: 100%;
            height: auto;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        }}
        
        .static-card.col-4 {{ grid-column: span 4; }}
        .static-card.col-6 {{ grid-column: span 6; }}
        .static-card.col-12 {{ grid-column: span 12; }}

        @media (max-width: 1024px) {{
            .chart-card.col-6 {{ grid-column: span 12; }}
            .chart-card.col-4 {{ grid-column: span 12; }}
            .chart-card.col-8 {{ grid-column: span 12; }}
            .chart-card.col-12 {{ grid-column: span 12; }}
            .scorecard-flex {{ flex-direction: column; }}
            .student-profile-pane {{ border-right: none; border-bottom: 1px solid #e2e8f0; padding-right: 0; padding-bottom: 20px; }}
            .static-card.col-4 {{ grid-column: span 12; }}
            .static-card.col-6 {{ grid-column: span 12; }}
            .static-card.col-12 {{ grid-column: span 12; }}
        }}
    </style>
</head>
<body>

    <!-- Header Block -->
    <div class="dashboard-header">
        <div class="header-title">
            <h1>Báo Cáo Điểm Rèn Luyện Sinh Viên</h1>
            <p>Học kỳ II | Năm học: 2025-2026</p>
        </div>
        <div class="class-badge">LỚP: E25CQCE02-N</div>
    </div>

    <!-- KPIs Cards -->
    <div class="kpi-grid">
        <div class="kpi-card">
            <span class="kpi-title">Sĩ Số Lớp</span>
            <span class="kpi-value" id="kpi-total-students">0</span>
            <span class="kpi-subtext">Sinh viên xếp hạng</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-title">Điểm Trung Bình Lớp</span>
            <span class="kpi-value" id="kpi-average-score">0.0</span>
            <span class="kpi-subtext">Trên thang điểm 100</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-title">Điểm Cao Nhất</span>
            <span class="kpi-value" id="kpi-highest-score">0</span>
            <span class="kpi-subtext" id="kpi-highest-name">N/A</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-title">Điểm Thấp Nhất</span>
            <span class="kpi-value" id="kpi-lowest-score">0</span>
            <span class="kpi-subtext" id="kpi-lowest-name">N/A</span>
        </div>
    </div>

    <!-- Interactive Filters Block -->
    <div class="filter-bar">
        <div class="filter-group" style="flex: 2;">
            <span class="filter-label">Tìm kiếm sinh viên</span>
            <input type="text" id="search-box" class="search-input" placeholder="Nhập tên hoặc mã sinh viên (MSV)...">
        </div>
        <div class="filter-group">
            <span class="filter-label">Bộ lọc xếp loại</span>
            <select id="rating-filter" class="select-input">
                <option value="ALL">Tất cả xếp loại</option>
                <option value="Xuất sắc">Xuất sắc (90-100đ)</option>
                <option value="Tốt">Tốt (80-89đ)</option>
                <option value="Khá">Khá (65-79đ)</option>
                <option value="Trung bình">Trung bình (50-64đ)</option>
                <option value="Yếu">Yếu (35-49đ)</option>
                <option value="Kém">Kém (Dưới 35đ)</option>
            </select>
        </div>
    </div>

    <!-- View Switcher -->
    <div class="view-selector-container">
        <div class="view-selector">
            <button class="view-btn active" id="btn-interactive" onclick="switchView('interactive')">
                📊 Biểu đồ tương tác
            </button>
            <button class="view-btn" id="btn-static" onclick="switchView('static')">
                🖼️ Báo cáo ảnh tĩnh
            </button>
        </div>
    </div>

    <!-- Interactive Charts Section -->
    <div id="interactive-section" class="view-section">

    <!-- Charts Grid Row 1 -->
    <div class="chart-grid">
        <!-- Distribution Hist Chart -->
        <div class="chart-card col-6">
            <div class="card-header">
                <div class="card-title">
                    <h3>Phân phối Điểm Rèn Luyện</h3>
                    <p>Số lượng sinh viên theo từng khung điểm</p>
                </div>
            </div>
            <div class="chart-container">
                <canvas id="distChart"></canvas>
            </div>
        </div>

        <!-- Rating Pie Chart -->
        <div class="chart-card col-6">
            <div class="card-header">
                <div class="card-title">
                    <h3>Tỉ lệ Xếp loại Rèn luyện</h3>
                    <p>Phần trăm xếp loại của lớp học</p>
                </div>
            </div>
            <div class="chart-container">
                <canvas id="ratingChart"></canvas>
            </div>
        </div>
    </div>

    <!-- Charts Grid Row 2: Criteria Averages + Interactive Student Radar Scorecard -->
    <div class="chart-grid">
        <!-- Main Criteria averages -->
        <div class="chart-card col-6">
            <div class="card-header">
                <div class="card-title">
                    <h3>Điểm trung bình theo Tiêu chí</h3>
                    <p>So sánh điểm trung bình thực tế với điểm tối đa</p>
                </div>
            </div>
            <div class="chart-container">
                <canvas id="criteriaChart"></canvas>
            </div>
        </div>

        <!-- Detailed Individual Student Radar Card -->
        <div class="chart-card col-6">
            <div class="card-header" style="margin-bottom: 15px;">
                <div class="card-title">
                    <h3>Hồ sơ & Điểm Chi tiết Sinh viên</h3>
                    <p>So sánh trực quan điểm sinh viên với điểm trung bình lớp</p>
                </div>
            </div>
            
            <!-- "Chọn sinh viên" dropdown is now full-width at the top of the card -->
            <div class="profile-list-select" style="margin-bottom: 20px;">
                <span class="filter-label" style="display: block; margin-bottom: 6px;">Chọn sinh viên</span>
                <select id="student-picker" class="select-input" style="height: 42px; width: 100%;"></select>
            </div>
            
            <div class="scorecard-flex">
                <div class="student-profile-pane">
                    <div class="profile-details">
                        <div class="detail-row">
                            <span class="detail-label">Họ và tên</span>
                            <span class="detail-value" id="prof-name">N/A</span>
                        </div>
                        <div class="detail-row">
                            <span class="detail-label">Mã số sinh viên</span>
                            <span class="detail-value" id="prof-msv">N/A</span>
                        </div>
                        <div class="detail-row">
                            <span class="detail-label">Ngày sinh</span>
                            <span class="detail-value" id="prof-dob">N/A</span>
                        </div>
                        <div class="detail-row">
                            <span class="detail-label">Tổng điểm DRL</span>
                            <span class="detail-value bold-cell" id="prof-total" style="font-size: 16px; color: var(--accent-blue);">0đ</span>
                        </div>
                        <div class="detail-row" style="border-bottom: none;">
                            <span class="detail-label">Xếp loại rèn luyện</span>
                            <span class="rating-badge" id="prof-rating">N/A</span>
                        </div>
                    </div>
                </div>
                <div class="student-chart-pane">
                    <canvas id="studentRadarChart" style="max-height: 240px; max-width: 240px;"></canvas>
                </div>
            </div>
        </div>
    </div>

    <!-- Charts Grid Row 3: Big Ass Line Chart (Subcriteria lines) -->
    <div class="chart-grid">
        <div class="chart-card col-12">
            <div class="card-header">
                <div class="card-title">
                    <h3>Biểu đồ Đường Chi tiết Tiêu chí Con (1.1 - 5.3)</h3>
                    <p>Cuộn chuột để phóng to, bấm giữ kéo để di chuyển. Nhấp trực tiếp vào đường của sinh viên để chọn nhanh.</p>
                </div>
                <div style="display: flex; gap: 20px; align-items: center;">
                    <button id="reset-zoom-btn" class="btn-reset-zoom">Reset Zoom</button>
                    <label class="switch-container">
                        <input type="checkbox" id="toggle-all-lines" checked> Hiện tất cả sinh viên
                    </label>
                </div>
            </div>
            <div class="chart-container-large">
                <canvas id="subcriteriaLineChart"></canvas>
            </div>
        </div>
    </div>
    </div>

    <!-- Static Charts Section -->
    <div id="static-section" class="view-section hidden">
        <div class="static-grid">
            <!-- Combined Dashboard Card -->
            <div class="static-card col-12">
                <div class="card-header" style="width: 100%;">
                    <div class="card-title">
                        <h3>Bảng Thống Kê Tổng Hợp (Matplotlib Dashboard)</h3>
                        <p>Báo cáo tĩnh tích hợp phân phối điểm, xếp loại, trung bình tiêu chí và bảng số liệu</p>
                    </div>
                </div>
                {img_dashboard_tag}
            </div>
            
            <!-- Student Subcriteria Line Chart Card -->
            <div class="static-card col-12">
                <div class="card-header" style="width: 100%;">
                    <div class="card-title">
                        <h3>Biểu Đồ Chi Tiết Theo Tiêu Chí Con (Matplotlib Line Chart)</h3>
                        <p>Đường biểu diễn điểm của tất cả sinh viên được chuẩn hóa theo phần trăm</p>
                    </div>
                </div>
                {img_linechart_tag}
            </div>
        </div>
        
        <!-- Row of individual charts -->
        <div class="static-grid" style="margin-top: 30px;">
            <div class="static-card col-4">
                <div class="card-header" style="width: 100%;">
                    <div class="card-title">
                        <h3>Phân Phối Điểm Số</h3>
                    </div>
                </div>
                {img_score_dist_tag}
            </div>
            <div class="static-card col-4">
                <div class="card-header" style="width: 100%;">
                    <div class="card-title">
                        <h3>Tỉ Lệ Xếp Loại</h3>
                    </div>
                </div>
                {img_rating_dist_tag}
            </div>
            <div class="static-card col-4">
                <div class="card-header" style="width: 100%;">
                    <div class="card-title">
                        <h3>Trung Bình Tiêu Chí</h3>
                    </div>
                </div>
                {img_criteria_avg_tag}
            </div>
        </div>
    </div>

    <!-- Student List Table Card -->
    <div class="table-card">
        <div class="card-header" style="margin-bottom: 15px;">
            <div class="card-title">
                <h3>Danh sách Chi tiết Sinh viên</h3>
                <p>Nhấp vào tiêu đề cột để sắp xếp thứ tự danh sách (▲/▼). Nhấp vào dòng sinh viên để xem biểu đồ radar ở trên.</p>
            </div>
        </div>
        <div class="table-wrapper">
            <table id="scores-table">
                <thead>
                    <tr>
                        <th data-sort="tt" style="width: 60px;">TT ⇅</th>
                        <th data-sort="name">Họ và Tên ⇅</th>
                        <th data-sort="msv">Mã Sinh viên ⇅</th>
                        <th data-sort="dob">Ngày sinh ⇅</th>
                        <th data-sort="tc0" style="text-align: center;">TC1 (20) ⇅</th>
                        <th data-sort="tc1" style="text-align: center;">TC2 (25) ⇅</th>
                        <th data-sort="tc2" style="text-align: center;">TC3 (20) ⇅</th>
                        <th data-sort="tc3" style="text-align: center;">TC4 (25) ⇅</th>
                        <th data-sort="tc4" style="text-align: center;">TC5 (10) ⇅</th>
                        <th data-sort="total" style="text-align: center;">Tổng cộng ⇅</th>
                        <th data-sort="rating" style="text-align: center;">Xếp loại ⇅</th>
                    </tr>
                </thead>
                <tbody id="table-body">
                    <!-- Loaded dynamically -->
                </tbody>
            </table>
        </div>
    </div>

    <!-- Footer metadata -->
    <div class="dashboard-footer">
        Hệ thống Quản lý điểm rèn luyện DRL E25CQCE02-N • Khoa Đào tạo Công nghệ thông tin 2 • Học viện Công nghệ Bưu chính Viễn thông
    </div>

    <!-- JavaScript Data and Interactive Logic -->
    <script>
        const students = {students_json};
        
        // Define rating config
        const ratingConfig = {{
            "Xuất sắc": {{ bg: "badge-xuat-sac", color: "#166534", chartBg: "rgba(22, 101, 52, 0.8)", border: "#166534" }},
            "Tốt": {{ bg: "badge-tot", color: "#0369a1", chartBg: "rgba(3, 105, 161, 0.8)", border: "#0369a1" }},
            "Khá": {{ bg: "badge-kha", color: "#854d0e", chartBg: "rgba(133, 77, 14, 0.8)", border: "#854d0e" }},
            "Trung bình": {{ bg: "badge-trung-binh", color: "#9a3412", chartBg: "rgba(154, 52, 18, 0.8)", border: "#9a3412" }},
            "Yếu": {{ bg: "badge-yeu", color: "#991b1b", chartBg: "rgba(153, 27, 27, 0.8)", border: "#991b1b" }},
            "Kém": {{ bg: "badge-kem", color: "#475569", chartBg: "rgba(71, 85, 105, 0.8)", border: "#475569" }}
        }};

        // State variables
        let filteredStudents = [...students];
        let selectedStudentMsv = students.length > 0 ? students[0].msv : null;
        
        // Sorting State
        let currentSortColumn = "tt";
        let currentSortDir = "asc";
        
        // Chart.js instances
        let distChartInstance = null;
        let ratingChartInstance = null;
        let criteriaChartInstance = null;
        let radarChartInstance = null;
        let lineChartInstance = null;

        // Subcriteria metadata
        const subcriteriaKeys = [
            "1.1", "1.2", "1.3", "1.4", "1.5",
            "2.1", "2.2", "2.3",
            "3.1", "3.2", "3.3", "3.4", "3.5",
            "4.1", "4.2", "4.3", "4.4", "4.5", "4.6",
            "5.1", "5.2", "5.3"
        ];
        
        const subcriteriaMax = {{
            "1.1": 3.0, "1.2": 10.0, "1.3": 4.0, "1.4": 2.0, "1.5": 1.0,
            "2.1": 15.0, "2.2": 5.0, "2.3": 5.0,
            "3.1": 10.0, "3.2": 4.0, "3.3": 3.0, "3.4": 3.0, "3.5": 0.0,
            "4.1": 8.0, "4.2": 5.0, "4.3": 5.0, "4.4": 5.0, "4.5": 2.0, "4.6": 0.0,
            "5.1": 4.0, "5.2": 3.0, "5.3": 3.0
        }};

        // === PRE-COMPUTED CACHES (computed once at load, never recomputed) ===
        // MSV → student Map for O(1) lookups in tooltips and selections
        const studentsByMsv = new Map(students.map(s => [s.msv, s]));

        // Pre-compute percentage arrays for every student (static data)
        const studentPctCache = new Map();
        students.forEach(s => {{
            studentPctCache.set(s.msv, subcriteriaKeys.map(key => {{
                const score = s[`sub_${{key}}`] || 0.0;
                const max = subcriteriaMax[key];
                return max === 0.0 ? (score === 0.0 ? 100 : 0) : (score / max) * 100;
            }}));
        }});

        // Pre-compute class average percentages (depends only on static students array)
        const classAvgPctsCache = subcriteriaKeys.map(key => {{
            const sum = students.reduce((acc, s) => acc + (s[`sub_${{key}}`] || 0.0), 0);
            const max = subcriteriaMax[key];
            const avg = sum / students.length;
            return max === 0.0 ? (avg === 0.0 ? 100 : 0) : (avg / max) * 100;
        }});

        // Pre-compute line colors for each student (based on static excel_total)
        const studentLineColorCache = new Map();
        students.forEach(s => {{
            studentLineColorCache.set(s.msv, getStudentLineColor(s.excel_total, 0.25));
        }});

        // Cache canvas context (avoid repeated DOM query)
        let lineChartCtx = null;

        // Initialize App
        function initApp() {{
            populateStudentPicker();
            updateHeaderArrows();
            updateDashboardData();
            setupEventListeners();
            loadStudentProfile(selectedStudentMsv);
            
            // Auto-fallback if Chart.js is not loaded
            if (typeof Chart === 'undefined') {{
                console.warn("Chart.js failed to load. Falling back to static report.");
                const btnInteractive = document.getElementById("btn-interactive");
                if (btnInteractive) btnInteractive.style.display = "none";
                switchView('static');
                
                // Show offline notice banner
                const notice = document.createElement("div");
                notice.style.background = "#fffbeb";
                notice.style.border = "1px solid #fef3c7";
                notice.style.color = "#b45309";
                notice.style.padding = "12px 20px";
                notice.style.borderRadius = "12px";
                notice.style.marginBottom = "20px";
                notice.style.fontSize = "14px";
                notice.style.fontWeight = "500";
                notice.textContent = "Lưu ý: Không thể tải thư viện vẽ biểu đồ tương tác (chế độ offline). Đang hiển thị báo cáo ảnh tĩnh độ phân giải cao.";
                document.body.insertBefore(notice, document.body.firstChild);
            }}
        }}

        // Switch between interactive dashboard and static report
        function switchView(view) {{
            const interactiveSec = document.getElementById("interactive-section");
            const staticSec = document.getElementById("static-section");
            const btnInteractive = document.getElementById("btn-interactive");
            const btnStatic = document.getElementById("btn-static");
            
            if (view === 'interactive') {{
                if (interactiveSec) interactiveSec.classList.remove("hidden");
                if (staticSec) staticSec.classList.add("hidden");
                if (btnInteractive) btnInteractive.classList.add("active");
                if (btnStatic) btnStatic.classList.remove("active");
            }} else {{
                if (interactiveSec) interactiveSec.classList.add("hidden");
                if (staticSec) staticSec.classList.remove("hidden");
                if (btnInteractive) btnInteractive.classList.remove("active");
                if (btnStatic) btnStatic.classList.add("active");
            }}
        }}

        // Populate student selector list
        function populateStudentPicker() {{
            const picker = document.getElementById("student-picker");
            picker.innerHTML = "";
            students.forEach(s => {{
                const opt = document.createElement("option");
                opt.value = s.msv;
                opt.textContent = `${{s.name}} (${{s.msv}})`;
                picker.appendChild(opt);
            }});
        }}

        // Calculate and update KPIs, Table, and Charts
        function updateDashboardData() {{
            // 1. Calculate and update KPI stats
            document.getElementById("kpi-total-students").textContent = filteredStudents.length;
            
            if (filteredStudents.length > 0) {{
                const totals = filteredStudents.map(s => s.excel_total);
                const avg = totals.reduce((a, b) => a + b, 0) / filteredStudents.length;
                document.getElementById("kpi-average-score").textContent = avg.toFixed(1);
                
                // Find highest & lowest
                const sorted = [...filteredStudents].sort((a, b) => b.excel_total - a.excel_total);
                const highest = sorted[0];
                const lowest = sorted[sorted.length - 1];
                
                document.getElementById("kpi-highest-score").textContent = highest.excel_total;
                document.getElementById("kpi-highest-name").textContent = highest.name;
                document.getElementById("kpi-lowest-score").textContent = lowest.excel_total;
                document.getElementById("kpi-lowest-name").textContent = lowest.name;
            }} else {{
                document.getElementById("kpi-average-score").textContent = "0.0";
                document.getElementById("kpi-highest-score").textContent = "0";
                document.getElementById("kpi-highest-name").textContent = "N/A";
                document.getElementById("kpi-lowest-score").textContent = "0";
                document.getElementById("kpi-lowest-name").textContent = "N/A";
            }}
            
            // 2. Render student table rows
            renderTable();
            
            // 3. Render charts
            renderDistributionChart();
            renderRatingChart();
            renderCriteriaAveragesChart();
            renderSubcriteriaLineChart();
        }}

        // Setup filter handlers and table selections
        function setupEventListeners() {{
            // Search Input listener
            document.getElementById("search-box").addEventListener("input", function(e) {{
                filterData();
            }});
            
            // Rating filter listener
            document.getElementById("rating-filter").addEventListener("change", function(e) {{
                filterData();
            }});
            
            // Student picker listener
            document.getElementById("student-picker").addEventListener("change", function(e) {{
                selectedStudentMsv = e.target.value;
                loadStudentProfile(selectedStudentMsv);
                highlightSelectedRow();
                renderSubcriteriaLineChart();
            }});
            
            // Line Chart Toggle listener
            document.getElementById("toggle-all-lines").addEventListener("change", function(e) {{
                renderSubcriteriaLineChart();
            }});

            // Reset zoom listener
            document.getElementById("reset-zoom-btn").addEventListener("click", () => {{
                if (lineChartInstance) {{
                    lineChartInstance.resetZoom();
                }}
            }});

            // Table header sorting event listeners
            document.querySelectorAll("#scores-table th").forEach(th => {{
                const sortKey = th.getAttribute("data-sort");
                if (sortKey) {{
                    th.style.cursor = "pointer";
                    th.addEventListener("click", () => {{
                        handleSort(sortKey);
                    }});
                }}
            }});
        }}

        // Handle sorting logic for columns
        function handleSort(key) {{
            if (currentSortColumn === key) {{
                currentSortDir = currentSortDir === "asc" ? "desc" : "asc";
            }} else {{
                currentSortColumn = key;
                currentSortDir = "desc"; // Default to desc for stats
                if (key === "name" || key === "msv" || key === "dob" || key === "tt") {{
                    currentSortDir = "asc";
                }}
            }}
            
            // Sort filteredStudents in-place
            filteredStudents.sort((a, b) => {{
                let valA, valB;
                
                if (key === "tt") {{
                    valA = a.tt;
                    valB = b.tt;
                }} else if (key === "name") {{
                    valA = a.name;
                    valB = b.name;
                }} else if (key === "msv") {{
                    valA = a.msv;
                    valB = b.msv;
                }} else if (key === "dob") {{
                    valA = a.dob;
                    valB = b.dob;
                }} else if (key.startsWith("tc")) {{
                    const idx = parseInt(key.replace("tc", ""));
                    valA = a.excel_tc[idx];
                    valB = b.excel_tc[idx];
                }} else if (key === "total") {{
                    valA = a.excel_total;
                    valB = b.excel_total;
                }} else if (key === "rating") {{
                    const ratingWeights = {{ "Xuất sắc": 5, "Tốt": 4, "Khá": 3, "Trung bình": 2, "Yếu": 1, "Kém": 0 }};
                    valA = ratingWeights[a.excel_rating] !== undefined ? ratingWeights[a.excel_rating] : -1;
                    valB = ratingWeights[b.excel_rating] !== undefined ? ratingWeights[b.excel_rating] : -1;
                }}
                
                if (valA < valB) return currentSortDir === "asc" ? -1 : 1;
                if (valA > valB) return currentSortDir === "asc" ? 1 : -1;
                return 0;
            }});
            
            updateHeaderArrows();
            renderTable();
        }}

        // Update sorting indicators on table headers
        function updateHeaderArrows() {{
            document.querySelectorAll("#scores-table th").forEach(th => {{
                const sortKey = th.getAttribute("data-sort");
                if (sortKey) {{
                    let baseText = th.textContent.replace(" ▲", "").replace(" ▼", "").replace(" ⇅", "");
                    
                    if (sortKey === currentSortColumn) {{
                        th.textContent = baseText + (currentSortDir === "asc" ? " ▲" : " ▼");
                        th.style.color = "var(--accent-blue)";
                        th.style.backgroundColor = "#eff6ff";
                    }} else {{
                        th.textContent = baseText + " ⇅";
                        th.style.color = "var(--text-muted)";
                        th.style.backgroundColor = "";
                    }}
                }}
            }});
        }}

        // Filter student list by search input and classification dropdown
        function filterData() {{
            const query = document.getElementById("search-box").value.trim().toLowerCase();
            const ratingFilter = document.getElementById("rating-filter").value;
            
            filteredStudents = students.filter(s => {{
                const matchesQuery = s.name.toLowerCase().includes(query) || s.msv.toLowerCase().includes(query);
                const matchesRating = (ratingFilter === "ALL") || (s.excel_rating === ratingFilter);
                return matchesQuery && matchesRating;
            }});
            
            // Reapply current sorting to the filtered list
            const prevCol = currentSortColumn;
            const prevDir = currentSortDir;
            currentSortColumn = ""; // trigger reset
            currentSortDir = prevDir;
            handleSort(prevCol);
        }}

        // Render data rows in scores table
        function renderTable() {{
            const tbody = document.getElementById("table-body");
            tbody.innerHTML = "";
            
            filteredStudents.forEach((s, idx) => {{
                const tr = document.createElement("tr");
                tr.id = `row-${{s.msv}}`;
                if (s.msv === selectedStudentMsv) tr.className = "selected-row";
                
                tr.addEventListener("click", () => {{
                    selectedStudentMsv = s.msv;
                    document.getElementById("student-picker").value = s.msv;
                    loadStudentProfile(s.msv);
                    highlightSelectedRow();
                    renderSubcriteriaLineChart();
                }});

                const ratingClass = ratingConfig[s.excel_rating] ? ratingConfig[s.excel_rating].bg : "badge-kem";
                
                tr.innerHTML = `
                    <td>${{s.tt}}</td>
                    <td class="bold-cell">${{s.name}}</td>
                    <td>${{s.msv}}</td>
                    <td>${{s.dob}}</td>
                    <td style="text-align: center;">${{s.excel_tc[0]}}</td>
                    <td style="text-align: center;">${{s.excel_tc[1]}}</td>
                    <td style="text-align: center;">${{s.excel_tc[2]}}</td>
                    <td style="text-align: center;">${{s.excel_tc[3]}}</td>
                    <td style="text-align: center;">${{s.excel_tc[4]}}</td>
                    <td style="text-align: center;" class="bold-cell">${{s.excel_total}}</td>
                    <td style="text-align: center;"><span class="rating-badge ${{ratingClass}}">${{s.excel_rating}}</span></td>
                `;
                tbody.appendChild(tr);
            }});
        }}

        // Highlight selected student row in table
        function highlightSelectedRow() {{
            document.querySelectorAll("#table-body tr").forEach(tr => {{
                tr.classList.remove("selected-row");
            }});
            const selectedRow = document.getElementById(`row-${{selectedStudentMsv}}`);
            if (selectedRow) {{
                selectedRow.classList.add("selected-row");
                selectedRow.scrollIntoView({{ behavior: 'smooth', block: 'nearest' }});
            }}
        }}

        // Load details and radar scorecard of selected student
        function loadStudentProfile(msv) {{
            const s = students.find(x => x.msv === msv);
            if (!s) return;
            
            document.getElementById("prof-name").textContent = s.name;
            document.getElementById("prof-msv").textContent = s.msv;
            document.getElementById("prof-dob").textContent = s.dob || "N/A";
            document.getElementById("prof-total").textContent = `${{s.excel_total}}đ`;
            
            const badge = document.getElementById("prof-rating");
            badge.textContent = s.excel_rating;
            badge.className = "rating-badge";
            if (ratingConfig[s.excel_rating]) {{
                badge.classList.add(ratingConfig[s.excel_rating].bg);
            }} else {{
                badge.classList.add("badge-kem");
            }}
            
            renderStudentRadarChart(s);
        }}

        // Calculate DRL score range distributions
        function renderDistributionChart() {{
            if (typeof Chart === 'undefined') return;
            const ctx = document.getElementById('distChart').getContext('2d');
            
            const bins = ["Dưới 50đ", "50 - 59đ", "60 - 69đ", "70 - 79đ", "80 - 89đ", "90 - 100đ"];
            const counts = [0, 0, 0, 0, 0, 0];
            
            filteredStudents.forEach(s => {{
                const score = s.excel_total;
                if (score < 50) counts[0]++;
                else if (score < 60) counts[1]++;
                else if (score < 70) counts[2]++;
                else if (score < 80) counts[3]++;
                else if (score < 90) counts[4]++;
                else counts[5]++;
            }});
            
            if (distChartInstance) distChartInstance.destroy();
            
            distChartInstance = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: bins,
                    datasets: [{{
                        label: 'Số lượng sinh viên',
                        data: counts,
                        backgroundColor: 'rgba(37, 99, 235, 0.8)',
                        borderColor: '#2563eb',
                        borderWidth: 1.5,
                        borderRadius: 6
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{ legend: {{ display: false }} }},
                    scales: {{
                        y: {{ beginAtZero: true, grid: {{ drawTicks: false }}, ticks: {{ precision: 0 }} }},
                        x: {{ grid: {{ display: false }} }}
                    }}
                }}
            }});
        }}

        // Donut Chart for Ratings distribution
        function renderRatingChart() {{
            if (typeof Chart === 'undefined') return;
            const ctx = document.getElementById('ratingChart').getContext('2d');
            
            const ratingsList = ["Xuất sắc", "Tốt", "Khá", "Trung bình", "Yếu", "Kém"];
            const counts = ratingsList.map(r => filteredStudents.filter(s => s.excel_rating === r).length);
            
            const activeLabels = [];
            const activeCounts = [];
            const activeColors = [];
            
            ratingsList.forEach((r, idx) => {{
                if (counts[idx] > 0) {{
                    activeLabels.push(r);
                    activeCounts.push(counts[idx]);
                    activeColors.push(ratingConfig[r] ? ratingConfig[r].color : '#64748b');
                }}
            }});
            
            if (ratingChartInstance) ratingChartInstance.destroy();
            
            if (activeCounts.length === 0) {{
                activeLabels.push("Không có dữ liệu");
                activeCounts.push(1);
                activeColors.push("#cbd5e1");
            }}
            
            ratingChartInstance = new Chart(ctx, {{
                type: 'doughnut',
                data: {{
                    labels: activeLabels,
                    datasets: [{{
                        data: activeCounts,
                        backgroundColor: activeColors,
                        borderWidth: 2,
                        borderColor: '#ffffff'
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    cutout: '60%',
                    plugins: {{
                        legend: {{ position: 'right', labels: {{ boxWidth: 12, font: {{ family: 'Inter', size: 12 }} }} }}
                    }}
                }}
            }});
        }}

        // Categories Averages progress chart
        function renderCriteriaAveragesChart() {{
            if (typeof Chart === 'undefined') return;
            const ctx = document.getElementById('criteriaChart').getContext('2d');
            
            const tcLabels = [
                "TC1: Học tập (Max 20)",
                "TC2: Quy chế (Max 25)",
                "TC3: Hoạt động (Max 20)",
                "TC4: Công dân (Max 25)",
                "TC5: Cán sự/CLB (Max 10)"
            ];
            
            const maxScores = [20, 25, 20, 25, 10];
            const avgScores = [0, 0, 0, 0, 0];
            
            if (filteredStudents.length > 0) {{
                for (let tc = 0; tc < 5; tc++) {{
                    const sum = filteredStudents.map(s => s.excel_tc[tc]).reduce((a, b) => a + b, 0);
                    avgScores[tc] = sum / filteredStudents.length;
                }}
            }}
            
            if (criteriaChartInstance) criteriaChartInstance.destroy();
            
            criteriaChartInstance = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: tcLabels,
                    datasets: [
                        {{
                            label: 'Điểm trung bình lớp',
                            data: avgScores,
                            backgroundColor: 'rgba(37, 99, 235, 0.85)',
                            borderColor: '#2563eb',
                            borderWidth: 1,
                            barThickness: 16,
                            borderRadius: 4
                        }},
                        {{
                            label: 'Điểm tối đa',
                            data: maxScores,
                            backgroundColor: '#edf2f7',
                            borderColor: '#e2e8f0',
                            borderWidth: 1,
                            barThickness: 16,
                            borderRadius: 4
                        }}
                    ]
                }},
                options: {{
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {{
                        x: {{ beginAtZero: true, max: 28, grid: {{ color: '#edf2f7' }} }},
                        y: {{ stacked: false, grid: {{ display: false }} }}
                    }},
                    plugins: {{
                        legend: {{ position: 'bottom', labels: {{ boxWidth: 12 }} }}
                    }}
                }}
            }});
        }}

        // Radar chart for selected student
        function renderStudentRadarChart(student) {{
            if (typeof Chart === 'undefined') return;
            const ctx = document.getElementById('studentRadarChart').getContext('2d');
            
            const labels = ["TC1: Học tập", "TC2: Quy chế", "TC3: Hoạt động", "TC4: Công dân", "TC5: Cán sự"];
            const maxScores = [20, 25, 20, 25, 10];
            
            const studentPcts = student.excel_tc.map((score, i) => (score / maxScores[i]) * 100);
            
            const classAverages = [0, 0, 0, 0, 0];
            for (let tc = 0; tc < 5; tc++) {{
                const sum = students.map(s => s.excel_tc[tc]).reduce((a, b) => a + b, 0);
                classAverages[tc] = ((sum / students.length) / maxScores[tc]) * 100;
            }}
            
            if (radarChartInstance) radarChartInstance.destroy();
            
            radarChartInstance = new Chart(ctx, {{
                type: 'radar',
                data: {{
                    labels: labels,
                    datasets: [
                        {{
                            label: `${{student.name}} (%)`,
                            data: studentPcts,
                            fill: true,
                            backgroundColor: 'rgba(37, 99, 235, 0.2)',
                            borderColor: '#2563eb',
                            pointBackgroundColor: '#2563eb',
                            pointBorderColor: '#fff',
                            pointHoverBackgroundColor: '#fff',
                            pointHoverBorderColor: '#2563eb',
                            borderWidth: 2
                        }},
                        {{
                            label: 'Trung bình cả lớp (%)',
                            data: classAverages,
                            fill: true,
                            backgroundColor: 'rgba(229, 62, 62, 0.1)',
                            borderColor: '#e53e3e',
                            pointBackgroundColor: '#e53e3e',
                            pointBorderColor: '#fff',
                            pointHoverBackgroundColor: '#fff',
                            pointHoverBorderColor: '#e53e3e',
                            borderWidth: 2,
                            borderDash: [4, 4]
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {{
                        r: {{
                            angleLines: {{ display: true, color: '#e2e8f0' }},
                            grid: {{ color: '#e2e8f0' }},
                            suggestedMin: 0,
                            suggestedMax: 100,
                            ticks: {{ stepSize: 20, display: false }}
                        }}
                    }},
                    plugins: {{
                        legend: {{ display: false }}
                    }}
                }}
            }});
        }}

        // Helper to get color code for student line charts
        function getStudentLineColor(totalScore, alpha = 0.3) {{
            if (totalScore >= 90) return `rgba(22, 101, 52, ${{alpha}})`;      // Xuất sắc
            if (totalScore >= 80) return `rgba(3, 105, 161, ${{alpha}})`;     // Tốt
            if (totalScore >= 65) return `rgba(133, 77, 14, ${{alpha}})`;     // Khá
            if (totalScore >= 50) return `rgba(154, 52, 18, ${{alpha}})`;     // Trung bình
            return `rgba(153, 27, 27, ${{alpha}})`;                           // Yếu/Kém
        }}

        // Render the BIG line chart showing subcriteria performance
        // OPTIMIZED: uses chart.update() instead of destroy/recreate, pre-computed caches,
        // and O(1) Map lookups instead of O(n) find() scans.
        function renderSubcriteriaLineChart() {{
            if (typeof Chart === 'undefined') return;
            if (!lineChartCtx) {{
                lineChartCtx = document.getElementById('subcriteriaLineChart').getContext('2d');
            }}

            const datasets = [];

            // 1. Add all students lines if toggled (use cached percentages & colors)
            const showAll = document.getElementById("toggle-all-lines").checked;
            if (showAll) {{
                // Sort by score for draw order (low scores behind high scores)
                const sortedForDrawing = [...filteredStudents].sort((a, b) => a.excel_total - b.excel_total);
                const len = sortedForDrawing.length;
                for (let i = 0; i < len; i++) {{
                    const s = sortedForDrawing[i];
                    if (s.msv === selectedStudentMsv) continue;

                    datasets.push({{
                        label: `${{s.name}} (${{s.msv}})`,
                        data: studentPctCache.get(s.msv),
                        borderColor: studentLineColorCache.get(s.msv),
                        backgroundColor: 'transparent',
                        borderWidth: 1.5,
                        pointRadius: 0,
                        pointHitRadius: 8,
                        tension: 0.1,
                        _msv: s.msv  // store MSV directly for O(1) tooltip/click lookup
                    }});
                }}
            }}

            // 2. Add Selected Student Line (O(1) Map lookup)
            const selectedStudent = studentsByMsv.get(selectedStudentMsv);
            if (selectedStudent) {{
                datasets.push({{
                    label: `Đang chọn: ${{selectedStudent.name}}`,
                    data: studentPctCache.get(selectedStudentMsv),
                    borderColor: '#2563eb',
                    backgroundColor: 'transparent',
                    borderWidth: 4,
                    pointRadius: 5,
                    pointBackgroundColor: '#2563eb',
                    pointBorderColor: '#ffffff',
                    pointBorderWidth: 2,
                    tension: 0.1,
                    zorder: 10,
                    _msv: selectedStudentMsv
                }});
            }}

            // 3. Add Class Average Line (use pre-computed cache)
            datasets.push({{
                label: 'Trung bình cả lớp',
                data: classAvgPctsCache,
                borderColor: '#e53e3e',
                backgroundColor: 'transparent',
                borderWidth: 3.5,
                borderDash: [6, 4],
                pointRadius: 5,
                pointBackgroundColor: '#e53e3e',
                pointBorderColor: '#ffffff',
                pointBorderWidth: 2,
                tension: 0.1,
                zorder: 5,
                _msv: null
            }});

            // UPDATE existing chart instead of destroy/recreate (massive perf win)
            if (lineChartInstance) {{
                lineChartInstance.data.datasets = datasets;
                lineChartInstance.update('none');  // 'none' skips animations for instant update
                return;
            }}

            // First-time creation only
            lineChartInstance = new Chart(lineChartCtx, {{
                type: 'line',
                data: {{
                    labels: subcriteriaKeys,
                    datasets: datasets
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    animation: {{ duration: 0 }},  // disable animations for perf
                    interaction: {{
                        mode: 'nearest',
                        intersect: false,
                        axis: 'xy'
                    }},
                    onClick: (event, elements) => {{
                        if (elements && elements.length > 0) {{
                            const datasetIndex = elements[0].datasetIndex;
                            const ds = lineChartInstance.data.datasets[datasetIndex];
                            
                            if (ds._msv == null) return;  // class average
                            if (ds.label.startsWith('Đang chọn:')) return;
                            
                            // Use stored MSV directly (O(1), no regex needed)
                            const clickedMsv = ds._msv;
                            selectedStudentMsv = clickedMsv;
                            document.getElementById("student-picker").value = clickedMsv;
                            loadStudentProfile(clickedMsv);
                            highlightSelectedRow();
                            renderSubcriteriaLineChart();
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            max: 105,
                            ticks: {{
                                callback: function(value) {{ return value + '%'; }},
                                color: '#475569'
                            }},
                            grid: {{ color: '#e2e8f0' }}
                        }},
                        x: {{
                            ticks: {{ font: {{ weight: 'bold' }}, color: '#475569' }},
                            grid: {{ color: '#edf2f7' }}
                        }}
                    }},
                    plugins: {{
                        legend: {{
                            labels: {{
                                filter: function(item, chartData) {{
                                    return item.text.includes('Trung bình cả lớp') || item.text.includes('Đang chọn:');
                                }},
                                boxWidth: 15,
                                font: {{ family: 'Inter', size: 12, weight: 'bold' }}
                            }}
                        }},
                        zoom: {{
                            zoom: {{
                                wheel: {{ enabled: true }},
                                pinch: {{ enabled: true }},
                                mode: 'xy'
                            }},
                            pan: {{
                                enabled: true,
                                mode: 'xy'
                            }}
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    const ds = context.dataset;
                                    const key = subcriteriaKeys[context.dataIndex];
                                    const pct = context.parsed.y.toFixed(1);
                                    
                                    let absScore = '';
                                    // Use stored _msv for O(1) Map lookup instead of find()
                                    if (ds._msv != null) {{
                                        const stud = studentsByMsv.get(ds._msv);
                                        if (stud) {{
                                            const score = stud[`sub_${{key}}`] || 0.0;
                                            absScore = ` (${{score}}đ)`;
                                        }}
                                    }}
                                    return `${{ds.label}}: ${{pct}}%${{absScore}}`;
                                }}
                            }}
                        }}
                    }}
                }}
            }});
        }}

        // Run application on load
        window.addEventListener("DOMContentLoaded", initApp);
    </script>
</body>
</html>
"""
    with open(dashboard_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"Successfully generated interactive HTML dashboard at {dashboard_path}")
