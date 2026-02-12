"""
=====================================================
MANUFACTURING OEE OPTIMIZATION PROJECT 
Author: Minh Hieu
Description:
- Import production data from Excel
- Calculate OEE (Availability, Performance, Quality)
- Analyze Cost, Cycle Time, Lead Time
- Generate statistical visualizations
- Run optimization scenario (reduce downtime)

=====================================================
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# 1. IMPORT EXCEL FILE
# Đổi đường dẫn nếu cần
file_path = "Manufacturing_Production_Sample_Data.xlsx"

data = pd.read_excel(file_path)

# Kiểm tra dữ liệu
print("Preview Data:")
print(data.head())

# ==========================================
# 2. DATA PREPARATION
# ==========================================

# Đảm bảo Date là datetime
data["Date"] = pd.to_datetime(data["Date"])

# Đảm bảo các cột số là numeric
numeric_cols = data.select_dtypes(include=np.number).columns
data[numeric_cols] = data[numeric_cols].apply(pd.to_numeric)

# ==========================================
# 3. OEE CALCULATION
# ==========================================

data["Operating_Time"] = data["Planned_Production_Time"] - data["Downtime"]

data["Availability"] = data["Operating_Time"] / data["Planned_Production_Time"]

data["Performance"] = (
    data["Ideal_Cycle_Time"] * data["Total_Output"]
) / data["Operating_Time"]

data["Good_Output"] = data["Total_Output"] - data["Defect_Quantity"]

data["Quality"] = data["Good_Output"] / data["Total_Output"]

data["OEE"] = (
    data["Availability"]
    * data["Performance"]
    * data["Quality"]
)

print("Average OEE Before Optimization:",
      round(data["OEE"].mean() * 100, 2), "%")

# ==========================================
# 4. HISTOGRAM – DEFECT DISTRIBUTION
# ==========================================

plt.figure()
plt.hist(data["Defect_Quantity"], bins=15)
plt.title("Histogram - Defect Quantity")
plt.xlabel("Defect Quantity")
plt.ylabel("Frequency")
plt.show()

# ==========================================
# 5. PARETO (TỔNG LỖI THEO NGÀY)
# ==========================================

pareto_data = data[["Date", "Defect_Quantity"]].copy()

pareto_data = pareto_data.sort_values(
    by="Defect_Quantity",
    ascending=False
)

pareto_data["Cum_Percent"] = (
    pareto_data["Defect_Quantity"].cumsum()
    / pareto_data["Defect_Quantity"].sum()
    * 100
)

plt.figure()
plt.bar(
    pareto_data["Date"].astype(str),
    pareto_data["Defect_Quantity"]
)
plt.xticks(rotation=90)
plt.title("Pareto - Defect by Day")
plt.show()

# ==========================================
# 6. CORRELATION MATRIX
# ==========================================

plt.figure()
sns.heatmap(
    data[[
        "OEE",
        "Downtime",
        "Defect_Quantity",
        "Material_Cost",
        "Labor_Cost"
    ]].corr(),
    annot=True
)

plt.title("Correlation Matrix")
plt.show()

# ==========================================
# 7. CYCLE TIME & LEAD TIME
# ==========================================

data["Cycle_Time"] = (
    data["Operating_Time"] / data["Total_Output"]
)

data["Lead_Time"] = data["Cycle_Time"] * data["Total_Output"]

plt.figure()
plt.plot(data["Date"], data["Cycle_Time"])
plt.xticks(rotation=45)
plt.title("Cycle Time Trend")
plt.show()

avg_cycle_time = data["Cycle_Time"].mean()
print("Average Cycle Time:",
      round(avg_cycle_time, 4), "minutes/unit")

plt.figure()
plt.plot(data["Date"], data["Lead_Time"])
plt.xticks(rotation=45)
plt.title("Lead Time Trend")
plt.show()

avg_lead_time = data["Lead_Time"].mean()
print("Average Lead Time:",
      round(avg_lead_time, 2), "minutes per day")


# ==========================================
# 8. COST IMPACT ANALYSIS
# ==========================================

data["Total_Cost"] = (
    data["Material_Cost"]
    + data["Labor_Cost"]
    + data["Overhead_Cost"]
)

data["Cost_per_Good_Unit"] = (
    data["Total_Cost"]
    / data["Good_Output"]
)

plt.figure()
plt.scatter(
    data["OEE"],
    data["Cost_per_Good_Unit"]
)
plt.xlabel("OEE")
plt.ylabel("Cost per Good Unit")
plt.title("OEE vs Cost per Good Unit")
plt.show()

# ==========================================
# 9. OPTIMIZATION SCENARIO
# Reduce Downtime by 20%
# ==========================================

optimized = data.copy()

optimized["Downtime"] *= 0.8

optimized["Operating_Time"] = (
    optimized["Planned_Production_Time"]
    - optimized["Downtime"]
)

optimized["Availability"] = (
    optimized["Operating_Time"]
    / optimized["Planned_Production_Time"]
)

optimized["OEE"] = (
    optimized["Availability"]
    * optimized["Performance"]
    * optimized["Quality"]
)

print("OEE After Optimization:",
      round(optimized["OEE"].mean() * 100, 2), "%")

# ==========================================
# AUTOMATED INSIGHTS
# ==========================================

print("\n***** PRODUCTION INSIGHTS *****")

# OEE evaluation
if data["OEE"].mean() > 0.8:
    print("OEE Level: World Class (>80%)")
elif data["OEE"].mean() > 0.6:
    print("OEE Level: Acceptable but improvement needed")
else:
    print("OEE Level: Low efficiency - urgent improvement required")


# Downtime impact
corr = data[["OEE", "Downtime"]].corr().iloc[0,1]
print("Correlation OEE vs Downtime:", round(corr,2))

if corr < -0.5:
    print("Downtime strongly reduces OEE")

# Cost impact
corr_cost = data[["OEE", "Cost_per_Good_Unit"]].corr().iloc[0,1]
print("Correlation OEE vs Cost per Unit:", round(corr_cost,2))

if corr_cost < -0.5:
    print("Higher OEE significantly reduces production cost")
    

# ==========================================
# END OF PROJECT
# ==========================================
