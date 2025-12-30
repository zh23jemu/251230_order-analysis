import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.statespace.sarimax import SARIMAX
from statsmodels.tsa.stattools import adfuller
from statsmodels.graphics.tsaplots import plot_acf, plot_pacf
import os

# 设置中文显示
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# 1. 数据整合与预处理
def load_and_merge_data():
    """加载并合并12个月的订单数据"""
    excel_file = r'c:\Coding\!temp_project\251230_order-analysis\案例-附件1：订单数据.xlsx'
    sheet_names = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月']
    
    # 合并所有月份数据
    all_data = []
    for sheet in sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet)
        all_data.append(df)
    
    merged_df = pd.concat(all_data, ignore_index=True)
    print(f"合并后数据形状: {merged_df.shape}")
    print(f"数据列名: {merged_df.columns.tolist()}")
    
    # 数据清洗
    merged_df = merged_df.dropna()
    merged_df['订单编号'] = merged_df['订单编号'].astype(int)
    merged_df['SKU编号'] = merged_df['SKU编号'].astype(int)
    merged_df['订货量'] = merged_df['订货量'].astype(int)
    
    # 添加时间特征
    merged_df['月份'] = merged_df['时间'].dt.month
    merged_df['季度'] = merged_df['时间'].dt.quarter
    merged_df['周'] = merged_df['时间'].dt.isocalendar().week
    merged_df['年份'] = merged_df['时间'].dt.year
    merged_df['日'] = merged_df['时间'].dt.day
    merged_df['小时'] = merged_df['时间'].dt.hour
    
    return merged_df

# 2. 季节性销售特点分析
def seasonal_analysis(df):
    """分析销售的季节性特点"""
    print("\n=== 季节性销售特点分析 ===")
    
    # 按月统计销售总量
    monthly_sales = df.groupby('月份')['订货量'].sum().reset_index()
    print("按月销售总量:")
    print(monthly_sales)
    
    # 按季度统计销售总量
    quarterly_sales = df.groupby('季度')['订货量'].sum().reset_index()
    print("\n按季度销售总量:")
    print(quarterly_sales)
    
    # 可视化
    plt.figure(figsize=(12, 5))
    
    plt.subplot(1, 2, 1)
    plt.bar(monthly_sales['月份'], monthly_sales['订货量'])
    plt.title('按月销售总量')
    plt.xlabel('月份')
    plt.ylabel('销售总量')
    
    plt.subplot(1, 2, 2)
    plt.bar(quarterly_sales['季度'], quarterly_sales['订货量'])
    plt.title('按季度销售总量')
    plt.xlabel('季度')
    plt.ylabel('销售总量')
    
    plt.tight_layout()
    plt.savefig('季节性销售分析.png', dpi=300)
    plt.close()
    
    return monthly_sales, quarterly_sales

# 3. 客户下单规律分析
def customer_order_patterns(df):
    """分析客户下单规律"""
    print("\n=== 客户下单规律分析 ===")
    
    # 客户订单频率
    customer_order_count = df.groupby('客户编号')['订单编号'].nunique().reset_index()
    customer_order_count.columns = ['客户编号', '订单数']
    print(f"平均每个客户订单数: {customer_order_count['订单数'].mean():.2f}")
    print(f"客户订单数分布:")
    print(customer_order_count['订单数'].describe())
    
    # 订单时间分布（小时）
    hourly_orders = df.groupby('小时')['订单编号'].nunique().reset_index()
    print("\n按小时订单分布:")
    print(hourly_orders)
    
    # 可视化
    plt.figure(figsize=(12, 5))
    
    plt.subplot(1, 2, 1)
    plt.hist(customer_order_count['订单数'], bins=20)
    plt.title('客户订单频率分布')
    plt.xlabel('订单数')
    plt.ylabel('客户数量')
    
    plt.subplot(1, 2, 2)
    plt.plot(hourly_orders['小时'], hourly_orders['订单编号'], marker='o')
    plt.title('按小时订单分布')
    plt.xlabel('小时')
    plt.ylabel('订单数')
    plt.grid(True)
    
    plt.tight_layout()
    plt.savefig('客户下单规律分析.png', dpi=300)
    plt.close()
    
    return customer_order_count, hourly_orders

# 4. 累托法则（80/20法则）分析
def pareto_analysis(df):
    """使用累托法则进行SKU和客户分类"""
    print("\n=== 累托法则（80/20法则）分析 ===")
    
    # SKU分类
    sku_sales = df.groupby('SKU编号')['订货量'].sum().reset_index()
    sku_sales = sku_sales.sort_values(by='订货量', ascending=False).reset_index(drop=True)
    sku_sales['累计销售量'] = sku_sales['订货量'].cumsum()
    sku_sales['累计销售占比'] = sku_sales['累计销售量'] / sku_sales['订货量'].sum() * 100
    
    # 确定A、B、C类SKU
    total_sku = len(sku_sales)
    a_threshold = 0.2  # 20% SKU
    b_threshold = 0.5  # 50% SKU
    
    sku_sales['分类'] = 'C'
    sku_sales.loc[0:int(total_sku * a_threshold), '分类'] = 'A'
    sku_sales.loc[int(total_sku * a_threshold):int(total_sku * b_threshold), '分类'] = 'B'
    
    print(f"SKU总数: {total_sku}")
    print(f"A类SKU数量: {len(sku_sales[sku_sales['分类'] == 'A'])}，占比: {len(sku_sales[sku_sales['分类'] == 'A'])/total_sku*100:.1f}%")
    print(f"B类SKU数量: {len(sku_sales[sku_sales['分类'] == 'B'])}，占比: {len(sku_sales[sku_sales['分类'] == 'B'])/total_sku*100:.1f}%")
    print(f"C类SKU数量: {len(sku_sales[sku_sales['分类'] == 'C'])}，占比: {len(sku_sales[sku_sales['分类'] == 'C'])/total_sku*100:.1f}%")
    
    # 各分类销售占比
    a_sales = sku_sales[sku_sales['分类'] == 'A']['订货量'].sum()
    b_sales = sku_sales[sku_sales['分类'] == 'B']['订货量'].sum()
    c_sales = sku_sales[sku_sales['分类'] == 'C']['订货量'].sum()
    total_sales = sku_sales['订货量'].sum()
    
    print(f"A类SKU销售占比: {a_sales/total_sales*100:.1f}%")
    print(f"B类SKU销售占比: {b_sales/total_sales*100:.1f}%")
    print(f"C类SKU销售占比: {c_sales/total_sales*100:.1f}%")
    
    # 客户分类
    customer_sales = df.groupby('客户编号')['订货量'].sum().reset_index()
    customer_sales = customer_sales.sort_values(by='订货量', ascending=False).reset_index(drop=True)
    customer_sales['累计销售量'] = customer_sales['订货量'].cumsum()
    customer_sales['累计销售占比'] = customer_sales['累计销售量'] / customer_sales['订货量'].sum() * 100
    
    total_customers = len(customer_sales)
    customer_sales['分类'] = 'C'
    customer_sales.loc[0:int(total_customers * a_threshold), '分类'] = 'A'
    customer_sales.loc[int(total_customers * a_threshold):int(total_customers * b_threshold), '分类'] = 'B'
    
    print(f"\n客户总数: {total_customers}")
    print(f"A类客户数量: {len(customer_sales[customer_sales['分类'] == 'A'])}，占比: {len(customer_sales[customer_sales['分类'] == 'A'])/total_customers*100:.1f}%")
    print(f"B类客户数量: {len(customer_sales[customer_sales['分类'] == 'B'])}，占比: {len(customer_sales[customer_sales['分类'] == 'B'])/total_customers*100:.1f}%")
    print(f"C类客户数量: {len(customer_sales[customer_sales['分类'] == 'C'])}，占比: {len(customer_sales[customer_sales['分类'] == 'C'])/total_customers*100:.1f}%")
    
    # 可视化累托曲线
    plt.figure(figsize=(12, 5))
    
    plt.subplot(1, 2, 1)
    plt.plot(sku_sales.index/total_sku*100, sku_sales['累计销售占比'], marker='.', label='SKU累托曲线')
    plt.axhline(y=80, color='r', linestyle='--', label='80%销售线')
    plt.axvline(x=20, color='g', linestyle='--', label='20%SKU线')
    plt.title('SKU累托曲线')
    plt.xlabel('SKU占比 (%)')
    plt.ylabel('累计销售占比 (%)')
    plt.legend()
    
    plt.subplot(1, 2, 2)
    plt.plot(customer_sales.index/total_customers*100, customer_sales['累计销售占比'], marker='.', label='客户累托曲线')
    plt.axhline(y=80, color='r', linestyle='--', label='80%销售线')
    plt.axvline(x=20, color='g', linestyle='--', label='20%客户线')
    plt.title('客户累托曲线')
    plt.xlabel('客户占比 (%)')
    plt.ylabel('累计销售占比 (%)')
    plt.legend()
    
    plt.tight_layout()
    plt.savefig('累托法则分析.png', dpi=300)
    plt.close()
    
    return sku_sales, customer_sales

# 5. EIQ分析
def eiq_analysis(df):
    """进行EIQ分析"""
    print("\n=== EIQ分析 ===")
    
    # 订单量(I)：每个订单的订货量
    order_quantity = df.groupby('订单编号')['订货量'].sum().reset_index()
    order_quantity.columns = ['订单编号', '订单总量']
    
    # 品项数(E)：每个订单包含的SKU数量
    order_sku_count = df.groupby('订单编号')['SKU编号'].nunique().reset_index()
    order_sku_count.columns = ['订单编号', '品项数']
    
    # 订货量(Q)：每个SKU的平均订货量
    sku_avg_quantity = df.groupby('SKU编号')['订货量'].mean().reset_index()
    sku_avg_quantity.columns = ['SKU编号', '平均订货量']
    
    print(f"订单平均总量: {order_quantity['订单总量'].mean():.2f}")
    print(f"订单平均品项数: {order_sku_count['品项数'].mean():.2f}")
    print(f"SKU平均订货量: {sku_avg_quantity['平均订货量'].mean():.2f}")
    
    # 可视化
    plt.figure(figsize=(15, 5))
    
    plt.subplot(1, 3, 1)
    plt.hist(order_quantity['订单总量'], bins=20)
    plt.title('订单量(I)分布')
    plt.xlabel('订单总量')
    plt.ylabel('订单数量')
    
    plt.subplot(1, 3, 2)
    plt.hist(order_sku_count['品项数'], bins=20)
    plt.title('品项数(E)分布')
    plt.xlabel('品项数')
    plt.ylabel('订单数量')
    
    plt.subplot(1, 3, 3)
    plt.hist(sku_avg_quantity['平均订货量'], bins=20)
    plt.title('订货量(Q)分布')
    plt.xlabel('平均订货量')
    plt.ylabel('SKU数量')
    
    plt.tight_layout()
    plt.savefig('EIQ分析.png', dpi=300)
    plt.close()
    
    return order_quantity, order_sku_count, sku_avg_quantity

# 6. SARIMA销售预测
def sarima_forecast(df, sku_id, forecast_weeks=52):
    """使用SARIMA模型对指定SKU进行销售预测"""
    print(f"\n=== SARIMA预测 - SKU {sku_id} ===")
    
    # 按周聚合SKU销售数据
    sku_data = df[df['SKU编号'] == sku_id]
    weekly_sales = sku_data.groupby(['年份', '周'])['订货量'].sum().reset_index()
    weekly_sales['日期'] = pd.to_datetime(weekly_sales['年份'].astype(str) + '-W' + weekly_sales['周'].astype(str) + '-1', format='%Y-W%U-%w')
    weekly_sales = weekly_sales.sort_values('日期')
    weekly_sales = weekly_sales.set_index('日期')['订货量']
    
    # 检查平稳性
    result = adfuller(weekly_sales)
    print(f'ADF检验结果: 统计量={result[0]:.4f}, p值={result[1]:.4f}, 临界值={result[4]}')
    
    if result[1] > 0.05:
        # 非平稳，进行差分
        weekly_sales_diff = weekly_sales.diff().dropna()
        result_diff = adfuller(weekly_sales_diff)
        print(f'一阶差分后ADF检验结果: 统计量={result_diff[0]:.4f}, p值={result_diff[1]:.4f}')
    
    # 拟合SARIMA模型（简化版本，实际应用中需要优化参数）
    try:
        # 使用简化的SARIMA模型，实际应用中应通过网格搜索优化参数
        model = SARIMAX(weekly_sales, order=(1, 1, 1), seasonal_order=(1, 1, 1, 52))
        model_fit = model.fit(disp=False)
        
        # 预测未来一年
        forecast = model_fit.forecast(steps=forecast_weeks)
        
        # 可视化
        plt.figure(figsize=(12, 6))
        plt.plot(weekly_sales, label='历史销售数据')
        plt.plot(forecast, label='预测销售数据', color='red')
        plt.title(f'SKU {sku_id} 销售预测')
        plt.xlabel('日期')
        plt.ylabel('销售量')
        plt.legend()
        plt.grid(True)
        plt.savefig(f'SARIMA预测_SKU_{sku_id}.png', dpi=300)
        plt.close()
        
        print(f"预测完成，未来{forecast_weeks}周预测结果已保存")
        return forecast
    except Exception as e:
        print(f"模型拟合失败: {e}")
        return None

# 主函数
def main():
    """主函数"""
    print("=== F布行出库效率提升解决方案 ===")
    
    # 1. 数据加载与合并
    df = load_and_merge_data()
    
    # 2. 季节性销售分析
    monthly_sales, quarterly_sales = seasonal_analysis(df)
    
    # 3. 客户下单规律分析
    customer_order_count, hourly_orders = customer_order_patterns(df)
    
    # 4. 累托法则分析
    sku_sales, customer_sales = pareto_analysis(df)
    
    # 5. EIQ分析
    order_quantity, order_sku_count, sku_avg_quantity = eiq_analysis(df)
    
    # 6. SARIMA销售预测 - 选择前3个A类SKU进行预测
    a_sku_list = sku_sales[sku_sales['分类'] == 'A']['SKU编号'].head(3).tolist()
    for sku_id in a_sku_list:
        sarima_forecast(df, sku_id)
    
    print("\n=== 分析完成 ===")
    print("所有分析结果和图表已保存到当前目录")

if __name__ == "__main__":
    main()