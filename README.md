import os
import json
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta

print("🚀 Starting portfolio project builder...")

# Create target directory
project_dir = "./ecommerce_sales_analysis"
os.makedirs(project_dir, exist_ok=True)
print(f"📁 Created directory: {project_dir}")

# ==========================================
# 1. GENERATE DATASETS & EXCEL SPREADSHEET
# ==========================================
np.random.seed(42)

start_date = datetime(2025, 7, 1)
end_date = datetime(2026, 6, 30)
delta_days = (end_date - start_date).days

# Customers Dimension
num_customers = 250
customer_ids = [f"CUST-{1000 + i}" for i in range(num_customers)]
first_names = ["John", "Jane", "Michael", "Emily", "David", "Sarah", "James", "Jessica", "Robert", "Ashley", 
               "William", "Amanda", "Joseph", "Melissa", "Charles", "Stephanie", "Thomas", "Nicole", "Daniel", "Elizabeth"]
last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Miller", "Davis", "Garcia", "Rodriguez", "Wilson",
              "Martinez", "Anderson", "Taylor", "Thomas", "Hernandez", "Moore", "Martin", "Jackson", "Thompson", "White"]

genders = ["Male", "Female", "Non-binary"]
age_groups = ["18-24", "25-34", "35-44", "45-54", "55+"]
regions = ["North America", "Europe", "Asia-Pacific", "Latin America"]

customer_list = []
for cid in customer_ids:
    first = np.random.choice(first_names)
    last = np.random.choice(last_names)
    name = f"{first} {last}"
    gender = np.random.choice(genders, p=[0.48, 0.48, 0.04])
    age = np.random.choice(age_groups, p=[0.15, 0.35, 0.25, 0.15, 0.10])
    join_days_ago = np.random.randint(0, 365)
    join_date = (start_date - timedelta(days=join_days_ago)).strftime("%Y-%m-%d")
    customer_list.append([cid, name, gender, age, join_date])

df_customers = pd.DataFrame(customer_list, columns=["Customer ID", "Customer Name", "Gender", "Age Group", "Join Date"])

# Products Matrix
products = {
    "Electronics": [
        {"name": "Wireless Headphones", "price": 120.00, "cost": 45.00},
        {"name": "Mechanical Keyboard", "price": 95.00, "cost": 35.00},
        {"name": "Gaming Mouse", "price": 60.00, "cost": 20.00},
        {"name": "USB-C Hub Multiport", "price": 45.00, "cost": 15.00},
        {"name": "Smart Fitness Watch", "price": 199.00, "cost": 80.00},
    ],
    "Clothing": [
        {"name": "Premium Cotton Hoodie", "price": 55.00, "cost": 18.00},
        {"name": "Slim Fit Denim Jeans", "price": 70.00, "cost": 22.00},
        {"name": "Breathable Running Shoes", "price": 110.00, "cost": 40.00},
        {"name": "Classic Crewneck T-Shirt", "price": 25.00, "cost": 6.50},
        {"name": "Waterproof Windbreaker", "price": 85.00, "cost": 28.00},
    ],
    "Home & Kitchen": [
        {"name": "Stainless Steel Water Bottle", "price": 30.00, "cost": 9.00},
        {"name": "Double-Walled Espresso Glasses", "price": 28.00, "cost": 8.00},
        {"name": "Non-Stick Frying Pan", "price": 45.00, "cost": 15.00},
        {"name": "Ergonomic Desk Chair", "price": 249.00, "cost": 100.00},
        {"name": "Electric Gooseneck Kettle", "price": 75.00, "cost": 25.00},
    ],
    "Books": [
        {"name": "Data Analytics Handbook", "price": 40.00, "cost": 12.00},
        {"name": "Leadership Essentials Guide", "price": 22.00, "cost": 5.00},
        {"name": "Sci-Fi Best Seller Novel", "price": 18.00, "cost": 4.00},
        {"name": "Gourmet Cooking Made Easy", "price": 35.00, "cost": 10.00},
        {"name": "Financial Freedom Blueprint", "price": 28.00, "cost": 7.00},
    ],
    "Beauty & Care": [
        {"name": "Hydrating Facial Serum", "price": 38.00, "cost": 10.00},
        {"name": "Mineral Sunscreen SPF 50", "price": 24.00, "cost": 6.50},
        {"name": "Matte Liquid Lipstick", "price": 18.00, "cost": 4.50},
        {"name": "Organic Argan Hair Oil", "price": 32.00, "cost": 9.00},
        {"name": "Ultrasonic Facial Cleansing Brush", "price": 65.00, "cost": 22.00},
    ]
}

# Transactions Generation (1200 rows)
num_sales = 1200
sales_list = []
payment_methods = ["Credit Card", "PayPal", "Apple Pay", "Google Pay", "Bank Transfer"]
segments = ["Consumer", "Corporate", "Home Office"]

for i in range(num_sales):
    order_id = f"ORD-{20250000 + i}"
    rand_day = np.random.randint(0, delta_days + 1)
    order_date = start_date + timedelta(days=rand_day)
    
    month = order_date.month
    season_multiplier = 1.35 if month in [11, 12] else 1.0 # High volume during holidays
    
    cid = np.random.choice(customer_ids)
    category = np.random.choice(list(products.keys()), p=[0.25, 0.22, 0.18, 0.15, 0.20])
    prod = np.random.choice(products[category])
    
    product_name = prod["name"]
    unit_price = prod["price"]
    unit_cost = prod["cost"]
    
    quantity = int(np.random.choice([1, 2, 3, 4, 5], p=[0.55, 0.25, 0.12, 0.05, 0.03]) * season_multiplier)
    quantity = max(1, quantity)
        
    revenue = quantity * unit_price
    cost = quantity * unit_cost
    profit = revenue - cost
    
    region = np.random.choice(regions, p=[0.40, 0.25, 0.20, 0.15])
    segment = np.random.choice(segments, p=[0.60, 0.25, 0.15])
    pay_method = np.random.choice(payment_methods, p=[0.50, 0.20, 0.15, 0.10, 0.05])
    
    sales_list.append([
        order_id, order_date.strftime("%Y-%m-%d"), cid, region, category, 
        product_name, unit_price, quantity, revenue, cost, profit, segment, pay_method
    ])

df_sales = pd.DataFrame(sales_list, columns=[
    "Order ID", "Order Date", "Customer ID", "Region", "Product Category", 
    "Product Name", "Unit Price", "Quantity", "Revenue", "Cost", "Profit", "Segment", "Payment Method"
])
df_sales = df_sales.sort_values(by="Order Date").reset_index(drop=True)

# Build & Style Excel File
print("📊 Generating styled Excel spreadsheet...")
xlsx_path = os.path.join(project_dir, "ecommerce_sales_data.xlsx")
writer = pd.ExcelWriter(xlsx_path, engine='openpyxl')

df_sales.to_excel(writer, sheet_name="Sales_Data", index=False)
df_customers.to_excel(writer, sheet_name="Customers", index=False)

wb = writer.book
instructions_ws = wb.create_sheet("README_ProjectGuide", 0)
instructions_ws.views.sheetView[0].showGridLines = True

# Formatting variables
primary_header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
thin_border = Border(
    left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'),
    top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0')
)

instructions_ws["A1"] = "E-commerce Sales Analysis Project Guide"
instructions_ws["A1"].font = Font(name="Segoe UI", size=18, bold=True, color="1B365D")
instructions_ws["A2"] = "Created for Interview Portfolio Advantage - Excel & Power BI"
instructions_ws["A2"].font = Font(name="Segoe UI", size=11, italic=True, color="555555")

# Auto-format and write spreadsheets with clean styles
for sheet_name in ["Sales_Data", "Customers"]:
    ws = wb[sheet_name]
    ws.views.sheetView[0].showGridLines = True
    
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = primary_header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.row_dimensions[1].height = 26
    
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        col_name = col[0].value
        
        for cell in col:
            cell.border = thin_border
            cell.font = Font(name="Segoe UI", size=10)
            
            if cell.row > 1:
                if col_name in ["Order Date", "Join Date", "Order ID", "Customer ID"]:
                    cell.alignment = Alignment(horizontal="center")
                elif col_name in ["Unit Price", "Revenue", "Cost", "Profit"]:
                    cell.number_format = "$#,##0.00"
                    cell.alignment = Alignment(horizontal="right")
                elif col_name in ["Quantity"]:
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right")
                    
            max_len = max(max_len, len(str(cell.value or '')))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

wb.save(xlsx_path)
print(f"✅ Excel saved successfully at {xlsx_path}")

# Export Dynamic Web DB Dataset (data.js)
sales_json = df_sales.to_dict(orient="records")
customers_json = df_customers.to_dict(orient="records")

data_js_path = os.path.join(project_dir, "data.js")
with open(data_js_path, "w", encoding="utf-8") as f:
    f.write(f"const SALES_DATA = {json.dumps(sales_json, ensure_ascii=False)};\n\n")
    f.write(f"const CUSTOMERS_DATA = {json.dumps(customers_json, ensure_ascii=False)};\n")
print(f"✅ Extracted dynamic JSON database to {data_js_path}")

# ==========================================
# 2. WRITE INDEX.HTML (FRONTEND STRUCTURAL LAYER)
# ==========================================
print("📝 Writing index.html structural file...")
index_html_content = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>E-commerce Sales Analysis Portfolio Project</title>
  
  <!-- Tailwind CSS V4 -->
  <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4.3.0/dist/index.global.js" crossorigin="anonymous"></script>
  
  <!-- Chart.js for Interactive Visualizations -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  
  <!-- Lucide Icons -->
  <script src="https://unpkg.com/lucide@latest"></script>
  
  <!-- Custom Style overrides -->
  <link rel="stylesheet" href="styles.css">
</head>
<body class="min-h-screen bg-slate-50 text-slate-900 dark:bg-slate-950 dark:text-slate-100 antialiased flex flex-col transition-colors duration-200">

  <!-- Header / Navigation -->
  <header class="sticky top-0 z-40 w-full border-b border-slate-200 bg-white/80 backdrop-blur dark:border-slate-800 dark:bg-slate-900/80">
    <div class="flex h-16 items-center justify-between px-6">
      <div class="flex items-center gap-3">
        <div class="p-2 bg-indigo-600 rounded-lg text-white">
          <i data-lucide="shopping-bag" class="w-6 h-6"></i>
        </div>
        <div>
          <h1 class="text-lg font-bold tracking-tight">E-Commerce Sales Insights</h1>
          <p class="text-xs text-slate-500 dark:text-slate-400">Excel & Power BI Showcase Portfolio</p>
        </div>
      </div>
      
      <div class="flex items-center gap-4">
        <!-- Light / Dark toggle -->
        <button id="themeToggle" class="p-2 rounded-lg text-slate-500 hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800 transition-colors" title="Toggle Theme">
          <i data-lucide="sun" id="sunIcon" class="w-5 h-5 hidden"></i>
          <i data-lucide="moon" id="moonIcon" class="w-5 h-5"></i>
        </button>
      </div>
    </div>
  </header>

  <!-- Main Layout -->
  <div class="flex-1 flex flex-col lg:flex-row">
    <!-- Sidebar Navigation -->
    <aside class="w-full lg:w-64 border-r border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 lg:sticky lg:top-16 lg:h-[calc(100vh-4rem)] overflow-y-auto">
      <div class="p-4 space-y-1">
        <button data-tab="dashboard" class="tab-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-left font-medium text-indigo-600 bg-indigo-50 dark:text-indigo-400 dark:bg-indigo-950/50 transition-all">
          <i data-lucide="layout-dashboard" class="w-5 h-5"></i>
          <span>Interactive Dashboard</span>
        </button>
        <button data-tab="pipeline" class="tab-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-left font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-50 dark:text-slate-400 dark:hover:text-slate-200 dark:hover:bg-slate-800/50 transition-all">
          <i data-lucide="git-branch" class="w-5 h-5"></i>
          <span>Data Pipeline & Star Schema</span>
        </button>
        <button data-tab="code" class="tab-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-left font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-50 dark:text-slate-400 dark:hover:text-slate-200 dark:hover:bg-slate-800/50 transition-all">
          <i data-lucide="code-2" class="w-5 h-5"></i>
          <span>DAX & Excel Formulas</span>
        </button>
        <button data-tab="explorer" class="tab-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-left font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-50 dark:text-slate-400 dark:hover:text-slate-200 dark:hover:bg-slate-800/50 transition-all">
          <i data-lucide="database" class="w-5 h-5"></i>
          <span>Raw Data Explorer</span>
        </button>
        <button data-tab="interview" class="tab-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-left font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-50 dark:text-slate-400 dark:hover:text-slate-200 dark:hover:bg-slate-800/50 transition-all">
          <i data-lucide="sparkles" class="w-5 h-5"></i>
          <span>Interview Cheat Sheet</span>
        </button>
      </div>
      
      <div class="mt-auto p-4 border-t border-slate-200 dark:border-slate-800 text-xs text-slate-500 dark:text-slate-400">
        <p class="font-medium text-slate-700 dark:text-slate-300">Portfolio Focus:</p>
        <p class="mt-1">Designed for Business Intelligence roles. Highlights ETL, Star Schema Data Modeling, DAX measures, and actionable insights.</p>
      </div>
    </aside>

    <!-- Main Content Panel -->
    <main class="flex-1 p-6 overflow-y-auto max-w-7xl mx-auto w-full">
      
      <!-- ================== SECTION: DASHBOARD ================== -->
      <section id="tab-dashboard" class="tab-pane space-y-6">
        <!-- Filter Bar -->
        <div class="bg-white dark:bg-slate-900 rounded-xl p-4 border border-slate-200 dark:border-slate-800 flex flex-wrap gap-4 items-center justify-between">
          <div class="flex flex-wrap gap-3 items-center">
            <span class="text-xs font-semibold uppercase tracking-wider text-slate-400 dark:text-slate-500 flex items-center gap-1.5">
              <i data-lucide="sliders-horizontal" class="w-4 h-4"></i>
              Interactive Filters:
            </span>
            
            <select id="regionFilter" class="bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-2 focus:ring-indigo-500 focus:outline-none">
              <option value="All">All Regions</option>
              <option value="North America">North America</option>
              <option value="Europe">Europe</option>
              <option value="Asia-Pacific">Asia-Pacific</option>
              <option value="Latin America">Latin America</option>
            </select>
            
            <select id="segmentFilter" class="bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-2 focus:ring-indigo-500 focus:outline-none">
              <option value="All">All Segments</option>
              <option value="Consumer">Consumer</option>
              <option value="Corporate">Corporate</option>
              <option value="Home Office">Home Office</option>
            </select>
          </div>
          
          <div class="text-xs text-indigo-600 dark:text-indigo-400 font-semibold flex items-center gap-1">
            <span class="inline-block w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
            <span>Simulated Power BI Dashboard</span>
          </div>
        </div>

        <!-- KPI Cards Grid -->
        <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4">
          <!-- KPI 1: Revenue -->
          <div class="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm hover:shadow transition-all relative overflow-hidden group">
            <div class="absolute right-3 top-3 text-indigo-500/10 dark:text-indigo-400/10 group-hover:scale-110 transition-transform">
              <i data-lucide="dollar-sign" class="w-12 h-12"></i>
            </div>
            <p class="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Total Revenue</p>
            <h3 class="text-2xl font-bold tracking-tight mt-1" id="kpi-revenue">$0.00</h3>
            <p class="text-xs text-emerald-600 dark:text-emerald-400 mt-2 font-medium flex items-center gap-1">
              <i data-lucide="trending-up" class="w-3 h-3"></i>
              <span>+14.3% YoY Sales</span>
            </p>
          </div>

          <!-- KPI 2: Profit -->
          <div class="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm hover:shadow transition-all relative overflow-hidden group">
            <div class="absolute right-3 top-3 text-emerald-500/10 dark:text-emerald-400/10 group-hover:scale-110 transition-transform">
              <i data-lucide="coins" class="w-12 h-12"></i>
            </div>
            <p class="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Gross Profit</p>
            <h3 class="text-2xl font-bold tracking-tight mt-1" id="kpi-profit">$0.00</h3>
            <p class="text-xs text-emerald-600 dark:text-emerald-400 mt-2 font-medium flex items-center gap-1">
              <i data-lucide="sparkles" class="w-3 h-3"></i>
              <span>High margin efficiency</span>
            </p>
          </div>

          <!-- KPI 3: Profit Margin -->
          <div class="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm hover:shadow transition-all relative overflow-hidden group">
            <div class="absolute right-3 top-3 text-amber-500/10 dark:text-amber-400/10 group-hover:scale-110 transition-transform">
              <i data-lucide="percent" class="w-12 h-12"></i>
            </div>
            <p class="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Profit Margin</p>
            <h3 class="text-2xl font-bold tracking-tight mt-1" id="kpi-margin">0.0%</h3>
            <p class="text-xs text-slate-500 dark:text-slate-400 mt-2 font-medium flex items-center gap-1">
              <i data-lucide="scale" class="w-3 h-3"></i>
              <span>Target Benchmark: 65.0%</span>
            </p>
          </div>

          <!-- KPI 4: AOV -->
          <div class="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm hover:shadow transition-all relative overflow-hidden group">
            <div class="absolute right-3 top-3 text-sky-500/10 dark:text-sky-400/10 group-hover:scale-110 transition-transform">
              <i data-lucide="shopping-cart" class="w-12 h-12"></i>
            </div>
            <p class="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Average Order Value (AOV)</p>
            <h3 class="text-2xl font-bold tracking-tight mt-1" id="kpi-aov">$0.00</h3>
            <p class="text-xs text-indigo-600 dark:text-indigo-400 mt-2 font-medium flex items-center gap-1">
              <i data-lucide="arrow-up-right" class="w-3 h-3"></i>
              <span>Driven by Electronics</span>
            </p>
          </div>
        </div>

        <!-- Dashboard Charts Grid -->
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
          
          <!-- Chart 1: Revenue Trends -->
          <div class="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm flex flex-col h-[320px]">
            <div class="flex justify-between items-center mb-4">
              <div>
                <h4 class="font-bold text-sm text-slate-800 dark:text-slate-100 flex items-center gap-2">
                  <i data-lucide="line-chart" class="w-4 h-4 text-indigo-500"></i>
                  Monthly Revenue Trend
                </h4>
                <p class="text-xs text-slate-500 dark:text-slate-400">Shows Holiday Seasonality Peaks (Nov-Dec)</p>
              </div>
            </div>
            <div class="flex-1 relative min-h-0">
              <canvas id="revenueTrendChart"></canvas>
            </div>
          </div>

          <!-- Chart 2: Best-selling Categories -->
          <div class="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm flex flex-col h-[320px]">
            <div class="flex justify-between items-center mb-4">
              <div>
                <h4 class="font-bold text-sm text-slate-800 dark:text-slate-100 flex items-center gap-2">
                  <i data-lucide="bar-chart-3" class="w-4 h-4 text-indigo-500"></i>
                  Revenue & Margin by Category
                </h4>
                <p class="text-xs text-slate-500 dark:text-slate-400">Compare product segment performance</p>
              </div>
            </div>
            <div class="flex-1 relative min-h-0">
              <canvas id="categoryChart"></canvas>
            </div>
          </div>

          <!-- Chart 3: Regional Sales Comparison -->
          <div class="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm flex flex-col h-[320px]">
            <div class="flex justify-between items-center mb-4">
              <div>
                <h4 class="font-bold text-sm text-slate-800 dark:text-slate-100 flex items-center gap-2">
                  <i data-lucide="globe" class="w-4 h-4 text-indigo-500"></i>
                  Regional Sales Comparison
                </h4>
                <p class="text-xs text-slate-500 dark:text-slate-400">Sales volume and share by geography</p>
              </div>
            </div>
            <div class="flex-1 relative min-h-0 flex items-center justify-center">
              <canvas id="regionalChart"></canvas>
            </div>
          </div>

          <!-- Chart 4: Segment Share -->
          <div class="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm flex flex-col h-[320px]">
            <div class="flex justify-between items-center mb-4">
              <div>
                <h4 class="font-bold text-sm text-slate-800 dark:text-slate-100 flex items-center gap-2">
                  <i data-lucide="pie-chart" class="w-4 h-4 text-indigo-500"></i>
                  Customer Segment Purchase Patterns
                </h4>
                <p class="text-xs text-slate-500 dark:text-slate-400">Contribution of different customer segments</p>
              </div>
            </div>
            <div class="flex-1 relative min-h-0">
              <canvas id="segmentChart"></canvas>
            </div>
          </div>
        </div>

        <!-- KPI Deep-Dive Narrative Panel -->
        <div class="bg-indigo-50 dark:bg-indigo-950/20 border border-indigo-100 dark:border-indigo-900/50 rounded-xl p-6">
          <h4 class="font-bold text-indigo-900 dark:text-indigo-400 flex items-center gap-2">
            <i data-lucide="brain-circuit" class="w-5 h-5"></i>
            Key Business Insights (Portfolio Selling Point)
          </h4>
          <div class="mt-4 grid grid-cols-1 md:grid-cols-3 gap-6 text-xs text-indigo-950 dark:text-slate-300">
            <div class="bg-white/60 dark:bg-slate-900/40 p-4 rounded-lg border border-indigo-100/30">
              <p class="font-bold text-indigo-800 dark:text-indigo-300">📈 Holiday Seasonality Surge</p>
              <p class="mt-1 leading-relaxed">November experienced a massive revenue surge (~35% above the annual average), directly matching the impact of Thanksgiving/Black Friday campaign strategies.</p>
            </div>
            <div class="bg-white/60 dark:bg-slate-900/40 p-4 rounded-lg border border-indigo-100/30">
              <p class="font-bold text-indigo-800 dark:text-indigo-300">💻 Category Efficiency Leaders</p>
              <p class="mt-1 leading-relaxed">While **Electronics** drives the absolute highest total transaction volume, **Beauty & Care** and **Books** demonstrate exceptional net margins, making them highly profitable subsegments.</p>
            </div>
            <div class="bg-white/60 dark:bg-slate-900/40 p-4 rounded-lg border border-indigo-100/30">
              <p class="font-bold text-indigo-800 dark:text-indigo-300">🎯 Demographic Strategy</p>
              <p class="mt-1 leading-relaxed">Consumers represent over 60% of all revenues, heavily concentrated in North America. Expanding targeted local campaigns to Europe & APAC represents our primary untapped expansion frontier.</p>
            </div>
          </div>
        </div>
      </section>

      <!-- ================== SECTION: PIPELINE & SCHEMA ================== -->
      <section id="tab-pipeline" class="tab-pane hidden space-y-6">
        <div class="bg-white dark:bg-slate-900 rounded-xl p-6 border border-slate-200 dark:border-slate-800 space-y-6">
          <div>
            <h3 class="text-lg font-bold flex items-center gap-2">
              <i data-lucide="git-fork" class="w-5 h-5 text-indigo-500"></i>
              ETL & Data Architecture Flow
            </h3>
            <p class="text-xs text-slate-500 dark:text-slate-400 mt-1">
              How the E-commerce Sales Analysis flows from the transactional databases into clean business dashboards.
            </p>
          </div>

          <!-- Flowchart Visual -->
          <div class="grid grid-cols-1 md:grid-cols-5 gap-4 items-center">
            
            <div class="bg-slate-50 dark:bg-slate-800/50 p-4 rounded-xl border border-slate-200 dark:border-slate-700 text-center relative">
              <div class="w-8 h-8 rounded-full bg-indigo-600 text-white flex items-center justify-center font-bold text-xs mx-auto mb-2">1</div>
              <p class="font-bold text-xs">Raw Transactional Data</p>
              <p class="text-[10px] text-slate-500 dark:text-slate-400 mt-1">E-commerce backend SQL servers exporting CSV logs.</p>
            </div>
            
            <div class="text-center text-slate-400 hidden md:block">
              <i data-lucide="chevron-right" class="w-6 h-6 mx-auto"></i>
            </div>

            <div class="bg-slate-50 dark:bg-slate-800/50 p-4 rounded-xl border border-slate-200 dark:border-slate-700 text-center relative">
              <div class="w-8 h-8 rounded-full bg-indigo-600 text-white flex items-center justify-center font-bold text-xs mx-auto mb-2">2</div>
              <p class="font-bold text-xs">Excel / SQL Prep (ETL)</p>
              <p class="text-[10px] text-slate-500 dark:text-slate-400 mt-1">Removed empty cells, formatted datatypes, validated foreign keys.</p>
            </div>

            <div class="text-center text-slate-400 hidden md:block">
              <i data-lucide="chevron-right" class="w-6 h-6 mx-auto"></i>
            </div>

            <div class="bg-slate-50 dark:bg-slate-800/50 p-4 rounded-xl border border-slate-200 dark:border-slate-700 text-center relative">
              <div class="w-8 h-8 rounded-full bg-indigo-600 text-white flex items-center justify-center font-bold text-xs mx-auto mb-2">3</div>
              <p class="font-bold text-xs">Power BI Modeling</p>
              <p class="text-[10px] text-slate-500 dark:text-slate-400 mt-1">Star Schema relationship mapping using clean primary keys.</p>
            </div>
          </div>

          <!-- Data Modeling Relationship (Star Schema Visualizer) -->
          <div class="border-t border-slate-200 dark:border-slate-800 pt-6">
            <h4 class="font-bold text-sm mb-4 flex items-center gap-2">
              <i data-lucide="network" class="w-4 h-4 text-indigo-500"></i>
              Data Modeling Star Schema Schema (Best Practice)
            </h4>
            
            <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
              <!-- Dimension Table -->
              <div class="border border-indigo-200 dark:border-indigo-900 bg-indigo-50/20 dark:bg-indigo-950/10 rounded-xl p-4">
                <div class="bg-indigo-600 text-white px-3 py-1.5 rounded-lg text-xs font-bold flex items-center justify-between">
                  <span>dim_Customers</span>
                  <span class="text-[10px] opacity-80">Dimension</span>
                </div>
                <ul class="mt-3 space-y-1.5 text-xs">
                  <li class="flex justify-between font-bold text-indigo-700 dark:text-indigo-400"><span>Customer ID</span> <span class="text-[10px] opacity-60">PK / String</span></li>
                  <li class="flex justify-between"><span>Customer Name</span> <span class="text-[10px] opacity-60">String</span></li>
                  <li class="flex justify-between"><span>Gender</span> <span class="text-[10px] opacity-60">String</span></li>
                  <li class="flex justify-between"><span>Age Group</span> <span class="text-[10px] opacity-60">String</span></li>
                  <li class="flex justify-between"><span>Join Date</span> <span class="text-[10px] opacity-60">Date</span></li>
                </ul>
              </div>

              <!-- Fact Table -->
              <div class="border border-emerald-200 dark:border-emerald-900 bg-emerald-50/20 dark:bg-emerald-950/10 rounded-xl p-4 md:col-span-1 relative">
                <div class="absolute -left-3 top-1/2 -translate-y-1/2 bg-white dark:bg-slate-900 px-2 py-1 text-[10px] font-bold border border-slate-200 dark:border-slate-700 rounded shadow hidden md:block">1 : N</div>
                <div class="bg-emerald-600 text-white px-3 py-1.5 rounded-lg text-xs font-bold flex items-center justify-between">
                  <span>fact_Sales</span>
                  <span class="text-[10px] opacity-80">Fact Table</span>
                </div>
                <ul class="mt-3 space-y-1.5 text-xs">
                  <li class="flex justify-between text-slate-400"><span>Order ID</span> <span class="text-[10px] opacity-60">ID / Key</span></li>
                  <li class="flex justify-between font-bold text-emerald-700 dark:text-emerald-400"><span>Customer ID</span> <span class="text-[10px] opacity-60">FK</span></li>
                  <li class="flex justify-between"><span>Order Date</span> <span class="text-[10px] opacity-60">Date</span></li>
                  <li class="flex justify-between"><span>Product Name</span> <span class="text-[10px] opacity-60">String</span></li>
                  <li class="flex justify-between"><span>Quantity</span> <span class="text-[10px] opacity-60">Integer</span></li>
                  <li class="flex justify-between"><span>Revenue</span> <span class="text-[10px] opacity-60">Decimal</span></li>
                  <li class="flex justify-between"><span>Cost</span> <span class="text-[10px] opacity-60">Decimal</span></li>
                  <li class="flex justify-between"><span>Profit</span> <span class="text-[10px] opacity-60">Decimal</span></li>
                </ul>
              </div>

              <!-- Calendar Dimension -->
              <div class="border border-indigo-200 dark:border-indigo-900 bg-indigo-50/20 dark:bg-indigo-950/10 rounded-xl p-4">
                <div class="bg-indigo-600 text-white px-3 py-1.5 rounded-lg text-xs font-bold flex items-center justify-between">
                  <span>dim_Calendar</span>
                  <span class="text-[10px] opacity-80">Dimension</span>
                </div>
                <ul class="mt-3 space-y-1.5 text-xs">
                  <li class="flex justify-between font-bold text-indigo-700 dark:text-indigo-400"><span>Date</span> <span class="text-[10px] opacity-60">PK / Date</span></li>
                  <li class="flex justify-between"><span>Year</span> <span class="text-[10px] opacity-60">Int</span></li>
                  <li class="flex justify-between"><span>Month</span> <span class="text-[10px] opacity-60">String</span></li>
                  <li class="flex justify-between"><span>Quarter</span> <span class="text-[10px] opacity-60">String</span></li>
                  <li class="flex justify-between"><span>Day of Week</span> <span class="text-[10px] opacity-60">Int</span></li>
                </ul>
              </div>
            </div>
          </div>
        </div>
      </section>

      <!-- ================== SECTION: CODE SNIPPETS ================== -->
      <section id="tab-code" class="tab-pane hidden space-y-6">
        <div class="bg-white dark:bg-slate-900 rounded-xl p-6 border border-slate-200 dark:border-slate-800 space-y-6">
          <div>
            <h3 class="text-lg font-bold flex items-center gap-2">
              <i data-lucide="terminal" class="w-5 h-5 text-indigo-500"></i>
              Calculations Formula Repository
            </h3>
            <p class="text-xs text-slate-500 dark:text-slate-400 mt-1">
              Ready-to-use DAX measures for Power BI and standard Excel formula structures used during analysis.
            </p>
          </div>

          <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
            
            <!-- Power BI DAX Measures -->
            <div class="space-y-4">
              <h4 class="text-xs font-bold uppercase tracking-wider text-indigo-600 dark:text-indigo-400 flex items-center gap-1.5">
                <span class="inline-block w-2 h-2 rounded bg-indigo-600"></span>
                Power BI (DAX Formulas)
              </h4>
              
              <div class="space-y-4">
                <div class="border border-slate-200 dark:border-slate-800 rounded-lg overflow-hidden">
                  <div class="bg-slate-50 dark:bg-slate-800 px-4 py-2 border-b border-slate-200 dark:border-slate-800 flex justify-between items-center">
                    <span class="text-xs font-bold text-slate-700 dark:text-slate-300">Total Profit Margin %</span>
                  </div>
                  <pre class="p-4 bg-slate-950 text-slate-100 font-mono text-xs overflow-x-auto"><code>Profit Margin % = 
DIVIDE(
    SUM(fact_Sales[Profit]), 
    SUM(fact_Sales[Revenue]), 
    0
)</code></pre>
                </div>

                <div class="border border-slate-200 dark:border-slate-800 rounded-lg overflow-hidden">
                  <div class="bg-slate-50 dark:bg-slate-800 px-4 py-2 border-b border-slate-200 dark:border-slate-800 flex justify-between items-center">
                    <span class="text-xs font-bold text-slate-700 dark:text-slate-300">YTD Revenue (Year-to-Date)</span>
                  </div>
                  <pre class="p-4 bg-slate-950 text-slate-100 font-mono text-xs overflow-x-auto"><code>YTD Revenue = 
TOTALYTD(
    SUM(fact_Sales[Revenue]), 
    dim_Calendar[Date]
)</code></pre>
                </div>
              </div>
            </div>

            <!-- Excel Formulas -->
            <div class="space-y-4">
              <h4 class="text-xs font-bold uppercase tracking-wider text-emerald-600 dark:text-emerald-400 flex items-center gap-1.5">
                <span class="inline-block w-2 h-2 rounded bg-emerald-600"></span>
                Excel Functions & Formulas
              </h4>
              
              <div class="space-y-4">
                <div class="border border-slate-200 dark:border-slate-800 rounded-lg overflow-hidden">
                  <div class="bg-slate-50 dark:bg-slate-800 px-4 py-2 border-b border-slate-200 dark:border-slate-800 flex justify-between items-center">
                    <span class="text-xs font-bold text-slate-700 dark:text-slate-300">Relational Lookup (Modern XLOOKUP)</span>
                  </div>
                  <pre class="p-4 bg-slate-950 text-slate-100 font-mono text-xs overflow-x-auto"><code>=XLOOKUP(
    C2, 
    Customers!$A$2:$A$251, 
    Customers!$B$2:$B$251, 
    "Customer Not Found"
)</code></pre>
                </div>

                <div class="border border-slate-200 dark:border-slate-800 rounded-lg overflow-hidden">
                  <div class="bg-slate-50 dark:bg-slate-800 px-4 py-2 border-b border-slate-200 dark:border-slate-800 flex justify-between items-center">
                    <span class="text-xs font-bold text-slate-700 dark:text-slate-300">Cohort Revenue Sum (SUMIFS)</span>
                  </div>
                  <pre class="p-4 bg-slate-950 text-slate-100 font-mono text-xs overflow-x-auto"><code>=SUMIFS(
    Sales_Data!$I$2:$I$1201, 
    Sales_Data!$D$2:$D$1201, "North America", 
    Sales_Data!$E$2:$E$1201, "Electronics"
)</code></pre>
                </div>
              </div>
            </div>
          </div>
        </div>
      </section>

      <!-- ================== SECTION: RAW DATA EXPLORER ================== -->
      <section id="tab-explorer" class="tab-pane hidden space-y-6">
        <div class="bg-white dark:bg-slate-900 rounded-xl p-6 border border-slate-200 dark:border-slate-800 space-y-4">
          <div class="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div>
              <h3 class="text-lg font-bold flex items-center gap-2">
                <i data-lucide="search" class="w-5 h-5 text-indigo-500"></i>
                E-commerce Transaction Log
              </h3>
              <p class="text-xs text-slate-500 dark:text-slate-400 mt-1">
                Real-time interactive client-side exploration of the raw transactional dataset (1,200 total entries).
              </p>
            </div>
            
            <div class="flex items-center gap-2">
              <input type="text" id="tableSearch" placeholder="Search product, ID, or region..." class="bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-2 text-xs w-64 focus:ring-2 focus:ring-indigo-500 focus:outline-none">
            </div>
          </div>

          <!-- Raw Data Table -->
          <div class="overflow-x-auto border border-slate-200 dark:border-slate-800 rounded-lg">
            <table class="w-full text-left text-xs border-collapse">
              <thead class="bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 sticky top-0 font-bold border-b border-slate-200 dark:border-slate-800">
                <tr>
                  <th class="p-3">Order ID</th>
                  <th class="p-3">Date</th>
                  <th class="p-3">Region</th>
                  <th class="p-3">Category</th>
                  <th class="p-3">Product Name</th>
                  <th class="p-3 text-right">Unit Price</th>
                  <th class="p-3 text-right">Qty</th>
                  <th class="p-3 text-right">Revenue</th>
                  <th class="p-3 text-right">Profit</th>
                </tr>
              </thead>
              <tbody id="rawTableBody" class="divide-y divide-slate-100 dark:divide-slate-800">
                <!-- Dynamically loaded rows -->
              </tbody>
            </table>
          </div>

          <!-- Pagination Details -->
          <div class="flex items-center justify-between pt-2">
            <span class="text-xs text-slate-500" id="tableInfo">Showing 1 to 10 of 1200 rows</span>
            <div class="flex gap-2">
              <button id="prevPage" class="px-3 py-1.5 border border-slate-200 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800 rounded-lg text-xs font-semibold disabled:opacity-40 disabled:cursor-not-allowed">Previous</button>
              <button id="nextPage" class="px-3 py-1.5 border border-slate-200 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800 rounded-lg text-xs font-semibold disabled:opacity-40 disabled:cursor-not-allowed">Next</button>
            </div>
          </div>
        </div>
      </section>

      <!-- ================== SECTION: INTERVIEW CHEAT SHEET ================== -->
      <section id="tab-interview" class="tab-pane hidden space-y-6">
        <div class="bg-white dark:bg-slate-900 rounded-xl p-6 border border-slate-200 dark:border-slate-800 space-y-6">
          <div>
            <h3 class="text-lg font-bold flex items-center gap-2">
              <i data-lucide="trophy" class="w-5 h-5 text-indigo-500"></i>
              Interview Advantage: KPI tracking & Insights
            </h3>
            <p class="text-xs text-slate-500 dark:text-slate-400 mt-1">
              Ace your interviews by explaining standard analytical models, KPI strategies, and product lifecycle concepts shown in this dashboard.
            </p>
          </div>

          <div class="space-y-6">
            <div class="space-y-2 border-l-4 border-indigo-500 pl-4">
              <p class="font-bold text-sm text-slate-800 dark:text-slate-100">Q1: "How did you design the data model for your sales tracking project?"</p>
              <p class="text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
                <strong>Answer Structure:</strong> "I modeled the dataset using a clean **Star Schema** architecture to maximize performance inside the Power BI query engine. By splitting the transactional flat table into a central <code class="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded text-indigo-600 font-mono">fact_Sales</code> table and matching lookup tables (<code class="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded text-indigo-600 font-mono">dim_Customers</code> and <code class="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded text-indigo-600 font-mono">dim_Calendar</code>), I achieved a 1-to-many relationship cardinality. This structure minimizes column lookup overhead and ensures filter context behaves predictably across months and demographics."
              </p>
            </div>

            <div class="space-y-2 border-l-4 border-indigo-500 pl-4">
              <p class="font-bold text-sm text-slate-800 dark:text-slate-100">Q2: "What primary business KPIs does your analysis track, and why are they valuable?"</p>
              <p class="text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
                <strong>Answer Structure:</strong> "I track four foundational e-commerce metrics: **Total Revenue**, **Gross Profit**, **Profit Margin**, and **Average Order Value (AOV)**. Looking at revenue alone is misleading; looking at Profit Margin ensures we spot categories that drive sales volume but have high operational costs (such as Electronics). Tracking AOV by segment helps identify whether customers are bundling items or ordering single-item shipments, allowing us to build data-driven recommendations."
              </p>
            </div>

            <div class="space-y-2 border-l-4 border-indigo-500 pl-4">
              <p class="font-bold text-sm text-slate-800 dark:text-slate-100">Q3: "If revenue is falling but transaction count is growing, what would you analyze?"</p>
              <p class="text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
                <strong>Answer Structure:</strong> "This suggests that while search traffic and orders are solid, customers are purchasing cheaper items (falling AOV) or items with deep promotional discounts. I would use DAX to build a cohort analyzer comparing Average Order Value month-over-month. Additionally, I would drill down into categories to check if higher-priced inventory (like Electronics) is experiencing stockouts, causing customer trends to skew toward lower-cost categories."
              </p>
            </div>
          </div>
        </div>
      </section>

    </main>
  </div>

  <!-- Footer -->
  <footer class="border-t border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 py-4 px-6 text-center text-xs text-slate-500 dark:text-slate-400">
    <p>© 2026 E-commerce Analytics Portfolio Project. Custom built with full offline interactive dashboard capabilities.</p>
  </footer>

  <!-- Include data first, then application logic -->
  <script src="data.js"></script>
  <script src="app.js"></script>
</body>
</html>"""

index_html_path = os.path.join(project_dir, "index.html")
with open(index_html_path, "w", encoding="utf-8") as f:
    f.write(index_html_content)
print(f"✅ Structural template written to {index_html_path}")

# ==========================================
# 3. WRITE APP.JS (FRONTEND APPLICATION LOGIC)
# ==========================================
print("📝 Writing app.js application logic...")
app_js_content = """document.addEventListener('DOMContentLoaded', function () {
  
  // Initialize Lucide Icons
  lucide.createIcons();
  
  // Theme Toggle Logic
  const themeToggle = document.getElementById('themeToggle');
  const sunIcon = document.getElementById('sunIcon');
  const moonIcon = document.getElementById('moonIcon');
  
  themeToggle.addEventListener('click', () => {
    const isDark = document.documentElement.classList.toggle('dark');
    if (isDark) {
      sunIcon.classList.remove('hidden');
      moonIcon.classList.add('hidden');
    } else {
      sunIcon.classList.add('hidden');
      moonIcon.classList.remove('hidden');
    }
    renderCharts();
  });

  // State Management
  let activeTab = 'dashboard';
  let currentPage = 1;
  const rowsPerPage = 12;
  let filteredSales = [...SALES_DATA];
  
  // Filters
  const regionFilter = document.getElementById('regionFilter');
  const segmentFilter = document.getElementById('segmentFilter');
  const tableSearch = document.getElementById('tableSearch');

  // Chart instances
  let revenueChartInstance = null;
  let categoryChartInstance = null;
  let regionalChartInstance = null;
  let segmentChartInstance = null;

  // Tab switching logic
  const tabs = document.querySelectorAll('.tab-btn');
  const tabPanes = document.querySelectorAll('.tab-pane');
  
  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      tabs.forEach(t => {
        t.classList.remove('text-indigo-600', 'bg-indigo-50', 'dark:text-indigo-400', 'dark:bg-indigo-950/50');
        t.classList.add('text-slate-600', 'hover:text-slate-900', 'hover:bg-slate-50', 'dark:text-slate-400', 'dark:hover:text-slate-200', 'dark:hover:bg-slate-800/50');
      });
      tab.classList.add('text-indigo-600', 'bg-indigo-50', 'dark:text-indigo-400', 'dark:bg-indigo-950/50');
      tab.classList.remove('text-slate-600', 'hover:text-slate-900', 'hover:bg-slate-50', 'dark:text-slate-400', 'dark:hover:text-slate-200');
      
      const targetPane = tab.getAttribute('data-tab');
      tabPanes.forEach(pane => {
        if (pane.id === `tab-${targetPane}`) {
          pane.classList.remove('hidden');
        } else {
          pane.classList.add('hidden');
        }
      });
      
      activeTab = targetPane;
      if (activeTab === 'dashboard') {
        renderCharts();
      }
    });
  });

  // KPI & Analytics calculations
  function applyFilters() {
    const regionVal = regionFilter.value;
    const segmentVal = segmentFilter.value;
    
    filteredSales = SALES_DATA.filter(row => {
      const matchRegion = regionVal === 'All' || row['Region'] === regionVal;
      const matchSegment = segmentVal === 'All' || row['Segment'] === segmentVal;
      return matchRegion && matchSegment;
    });

    calculateKPIs();
    renderCharts();
  }

  function calculateKPIs() {
    let totalRev = 0;
    let totalCost = 0;
    let totalQty = 0;
    
    filteredSales.forEach(row => {
      totalRev += row['Revenue'];
      totalCost += row['Cost'];
      totalQty += row['Quantity'];
    });
    
    const profit = totalRev - totalCost;
    const margin = totalRev > 0 ? (profit / totalRev) * 100 : 0;
    const aov = filteredSales.length > 0 ? totalRev / filteredSales.length : 0;

    // Display formatted
    document.getElementById('kpi-revenue').innerText = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(totalRev);
    document.getElementById('kpi-profit').innerText = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(profit);
    document.getElementById('kpi-margin').innerText = margin.toFixed(1) + '%';
    document.getElementById('kpi-aov').innerText = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 2 }).format(aov);
  }

  // Visualizations logic
  function renderCharts() {
    const isDark = document.documentElement.classList.contains('dark');
    const textThemeColor = isDark ? '#cbd5e1' : '#475569';
    const gridThemeColor = isDark ? '#334155' : '#f1f5f9';

    // Grouping by Month for Monthly Revenue Trend
    const monthlyData = {};
    filteredSales.forEach(row => {
      const d = new Date(row['Order Date']);
      const monthKey = d.toLocaleString('default', { month: 'short', year: 'numeric' });
      monthlyData[monthKey] = (monthlyData[monthKey] || 0) + row['Revenue'];
    });

    const monthLabels = Object.keys(monthlyData);
    const monthRevenues = Object.values(monthlyData);

    // Grouping by Category for Categories Chart
    const catData = {};
    const catProfitData = {};
    filteredSales.forEach(row => {
      const cat = row['Product Category'];
      catData[cat] = (catData[cat] || 0) + row['Revenue'];
      catProfitData[cat] = (catProfitData[cat] || 0) + row['Profit'];
    });

    const catLabels = Object.keys(catData);
    const catRevenues = Object.values(catData);
    const catProfits = Object.values(catProfitData);

    // Grouping by Region for Doughnut Chart
    const regData = {};
    filteredSales.forEach(row => {
      const reg = row['Region'];
      regData[reg] = (regData[reg] || 0) + row['Revenue'];
    });
    const regLabels = Object.keys(regData);
    const regRevenues = Object.values(regData);

    // Grouping by Segment
    const segData = {};
    filteredSales.forEach(row => {
      const seg = row['Segment'];
      segData[seg] = (segData[seg] || 0) + row['Revenue'];
    });
    const segLabels = Object.keys(segData);
    const segRevenues = Object.values(segData);

    // Render / Update Chart 1: Revenue Monthly Trend
    if (revenueChartInstance) revenueChartInstance.destroy();
    revenueChartInstance = new Chart(document.getElementById('revenueTrendChart'), {
      type: 'line',
      data: {
        labels: monthLabels,
        datasets: [{
          label: 'Monthly Revenue',
          data: monthRevenues,
          borderColor: '#4f46e5',
          backgroundColor: 'rgba(79, 70, 229, 0.05)',
          fill: true,
          tension: 0.3,
          borderWidth: 2
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false }
        },
        scales: {
          x: { ticks: { color: textThemeColor }, grid: { display: false } },
          y: { 
            ticks: { 
              color: textThemeColor,
              callback: function(value) { return '$' + (value / 1000) + 'k'; }
            },
            grid: { color: gridThemeColor }
          }
        }
      }
    });

    // Chart 2: Category Breakdown
    if (categoryChartInstance) categoryChartInstance.destroy();
    categoryChartInstance = new Chart(document.getElementById('categoryChart'), {
      type: 'bar',
      data: {
        labels: catLabels,
        datasets: [
          {
            label: 'Revenue',
            data: catRevenues,
            backgroundColor: '#6366f1',
            borderRadius: 6
          },
          {
            label: 'Net Profit',
            data: catProfits,
            backgroundColor: '#10b981',
            borderRadius: 6
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'top',
            labels: { color: textThemeColor, boxWidth: 12, font: { size: 10 } }
          }
        },
        scales: {
          x: { ticks: { color: textThemeColor }, grid: { display: false } },
          y: { 
            ticks: { 
              color: textThemeColor,
              callback: function(value) { return '$' + (value / 1000) + 'k'; }
            },
            grid: { color: gridThemeColor }
          }
        }
      }
    });

    // Chart 3: Regional Comparison
    if (regionalChartInstance) regionalChartInstance.destroy();
    regionalChartInstance = new Chart(document.getElementById('regionalChart'), {
      type: 'doughnut',
      data: {
        labels: regLabels,
        datasets: [{
          data: regRevenues,
          backgroundColor: ['#4f46e5', '#3b82f6', '#10b981', '#f59e0b', '#ec4899']
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'right',
            labels: { color: textThemeColor, boxWidth: 10, font: { size: 10 } }
          }
        }
      }
    });

    // Chart 4: Segment Share
    if (segmentChartInstance) segmentChartInstance.destroy();
    segmentChartInstance = new Chart(document.getElementById('segmentChart'), {
      type: 'pie',
      data: {
        labels: segLabels,
        datasets: [{
          data: segRevenues,
          backgroundColor: ['#6366f1', '#14b8a6', '#f59e0b']
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'right',
            labels: { color: textThemeColor, boxWidth: 10, font: { size: 10 } }
          }
        }
      }
    });
  }

  // Interactive Table Exploration Logic
  function getFilteredTableData() {
    const q = tableSearch.value.toLowerCase();
    return SALES_DATA.filter(row => {
      return (
        row['Order ID'].toLowerCase().includes(q) ||
        row['Product Name'].toLowerCase().includes(q) ||
        row['Region'].toLowerCase().includes(q) ||
        row['Product Category'].toLowerCase().includes(q)
      );
    });
  }

  function renderTable() {
    const tblBody = document.getElementById('rawTableBody');
    const tableData = getFilteredTableData();
    
    const startIdx = (currentPage - 1) * rowsPerPage;
    const endIdx = startIdx + rowsPerPage;
    const paginatedData = tableData.slice(startIdx, endIdx);
    
    tblBody.innerHTML = '';
    
    if (paginatedData.length === 0) {
      tblBody.innerHTML = `<tr><td colspan="9" class="p-6 text-center text-slate-500 font-medium">No transactions match your search query.</td></tr>`;
      document.getElementById('tableInfo').innerText = 'Showing 0 to 0 of 0 entries';
      document.getElementById('prevPage').disabled = true;
      document.getElementById('nextPage').disabled = true;
      return;
    }
    
    paginatedData.forEach(row => {
      const tr = document.createElement('tr');
      tr.className = 'hover:bg-slate-50 dark:hover:bg-slate-800/30 transition-colors';
      tr.innerHTML = `
        <td class="p-3 font-semibold text-slate-800 dark:text-slate-300">\${row['Order ID']}</td>
        <td class="p-3">\${row['Order Date']}</td>
        <td class="p-3">\${row['Region']}</td>
        <td class="p-3"><span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-indigo-50 dark:bg-indigo-950 text-indigo-600 dark:text-indigo-400">\${row['Product Category']}</span></td>
        <td class="p-3 font-medium text-slate-700 dark:text-slate-300">\${row['Product Name']}</td>
        <td class="p-3 text-right font-mono">$\${row['Unit Price'].toFixed(2)}</td>
        <td class="p-3 text-right font-mono">\${row['Quantity']}</td>
        <td class="p-3 text-right font-bold text-slate-900 dark:text-slate-100 font-mono">$\${row['Revenue'].toFixed(2)}</td>
        <td class="p-3 text-right font-semibold text-emerald-600 dark:text-emerald-400 font-mono">$\${row['Profit'].toFixed(2)}</td>
      `;
      tblBody.appendChild(tr);
    });

    const totalEntries = tableData.length;
    const showingTo = Math.min(endIdx, totalEntries);
    document.getElementById('tableInfo').innerText = `Showing \${startIdx + 1} to \${showingTo} of \${totalEntries} entries`;
    
    document.getElementById('prevPage').disabled = currentPage === 1;
    document.getElementById('nextPage').disabled = endIdx >= totalEntries;
  }

  // Event Listeners for Filters & Table pagination
  regionFilter.addEventListener('change', applyFilters);
  segmentFilter.addEventListener('change', applyFilters);
  
  tableSearch.addEventListener('input', () => {
    currentPage = 1;
    renderTable();
  });

  document.getElementById('prevPage').addEventListener('click', () => {
    if (currentPage > 1) {
      currentPage--;
      renderTable();
    }
  });

  document.getElementById('nextPage').addEventListener('click', () => {
    const tableData = getFilteredTableData();
    if (currentPage * rowsPerPage < tableData.length) {
      currentPage++;
      renderTable();
    }
  });

  // Init Calls
  calculateKPIs();
  renderCharts();
  renderTable();
});
"""

app_js_path = os.path.join(project_dir, "app.js")
with open(app_js_path, "w", encoding="utf-8") as f:
    f.write(app_js_content)
print(f"✅ Application controller logic written to {app_js_path}")

# ==========================================
# 4. WRITE STYLES.CSS (CUSTOM PRESENTATION SHEETS)
# ==========================================
print("📝 Writing styles.css style sheet...")
styles_css_content = """/* Custom Global Overrides */
body {
  font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
}

/* Custom dashboard scroll bars */
::-webkit-scrollbar {
  width: 6px;
  height: 6px;
}
::-webkit-scrollbar-track {
  background: transparent;
}
::-webkit-scrollbar-thumb {
  background: #cbd5e1;
  border-radius: 10px;
}
.dark ::-webkit-scrollbar-thumb {
  background: #475569;
}

/* Safe dynamic tables overflow rules */
.table-wrapper {
  overflow-x: auto;
  -webkit-overflow-scrolling: touch;
}
"""

styles_css_path = os.path.join(project_dir, "styles.css")
with open(styles_css_path, "w", encoding="utf-8") as f:
    f.write(styles_css_content)
print(f"✅ Design stylesheet written to {styles_css_path}")

print("\n🎉 Build Complete! All structural files generated in './ecommerce_sales_analysis'.")
print("👉 Simply double-click on 'ecommerce_sales_analysis/index.html' to open your interactive portfolio.")
