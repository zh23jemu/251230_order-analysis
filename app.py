from flask import Flask, render_template, jsonify, request
import pandas as pd
import numpy as np
import json
from datetime import datetime, timedelta
from statsmodels.tsa.statespace.sarimax import SARIMAX
from statsmodels.tsa.stattools import adfuller

app = Flask(__name__)

# 全局变量存储数据，避免重复加载
global_data = None

def convert_to_native_types(obj):
    """将numpy类型转换为Python原生类型"""
    if isinstance(obj, np.integer):
        return int(obj)
    elif isinstance(obj, np.floating):
        return float(obj)
    elif isinstance(obj, np.ndarray):
        return obj.tolist()
    elif isinstance(obj, dict):
        return {key: convert_to_native_types(value) for key, value in obj.items()}
    elif isinstance(obj, list):
        return [convert_to_native_types(item) for item in obj]
    else:
        return obj

# 加载并预处理数据
def load_and_merge_data():
    """加载并合并12个月的订单数据"""
    excel_file = r'c:\Coding\!completed_project\251230_order-analysis\案例-附件1：订单数据.xlsx'
    sheet_names = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月']
    
    # 合并所有月份数据
    all_data = []
    for sheet in sheet_names:
        try:
            df = pd.read_excel(excel_file, sheet_name=sheet)
            all_data.append(df)
        except Exception as e:
            print(f"无法读取工作表 {sheet}: {e}")
    
    if not all_data:
        raise ValueError("没有成功读取任何数据")
    
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
    merged_df['小时'] = merged_df['时间'].dt.hour
    merged_df['年份'] = merged_df['时间'].dt.year
    merged_df['周'] = merged_df['时间'].dt.isocalendar().week
    
    # 添加订单类型识别（基于客户编号）
    # 规则：C0001-C0050为镇内客户，C0051-C0101为镇外客户
    merged_df['订单类型'] = merged_df['客户编号'].apply(lambda x: '镇内' if int(x[1:]) <= 50 else '镇外')
    
    # 添加分拣时间分析（假设：订单处理时间 = 订单量 / 处理效率）
    # 假设处理效率：每小时处理1000件商品
    merged_df['分拣时间'] = merged_df['订货量'] / 1000
    
    # 计算准时完成情况
    merged_df['准时完成'] = merged_df.apply(lambda row: row['分拣时间'] <= (1 if row['订单类型'] == '镇内' else 2), axis=1)
    
    return merged_df

# 加载基础数据
def load_data():
    """加载并分析基础数据"""
    global global_data
    
    if global_data is None:
        global_data = load_and_merge_data()
    
    df = global_data
    
    # 1. 季节性销售分析
    monthly_sales_df = df.groupby('月份')['订货量'].sum().reset_index()
    monthly_sales = {
        '月份': monthly_sales_df['月份'].tolist(),
        '订货量': monthly_sales_df['订货量'].tolist()
    }
    
    quarterly_sales_df = df.groupby('季度')['订货量'].sum().reset_index()
    quarterly_sales = {
        '季度': quarterly_sales_df['季度'].tolist(),
        '订货量': quarterly_sales_df['订货量'].tolist()
    }
    
    # 2. 客户下单规律分析
    hourly_orders_df = df.groupby('小时')['订单编号'].nunique().reset_index()
    hourly_orders = {
        '小时': hourly_orders_df['小时'].tolist(),
        '订单数': hourly_orders_df['订单编号'].tolist()
    }
    
    # 3. SKU分类（累托法则）
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
    
    # 计算各分类统计
    sku_class_stats = []
    for class_name in ['A类', 'B类', 'C类']:
        class_data = sku_sales[sku_sales['分类'] == class_name[0]]
        sku_class_stats.append({
            'class': class_name,
            'count': len(class_data),
            'percentage': len(class_data) / total_sku * 100,
            'sales_percentage': class_data['订货量'].sum() / sku_sales['订货量'].sum() * 100
        })
    
    sku_classes = {
        '分类': [stat['class'] for stat in sku_class_stats],
        '数量': [stat['count'] for stat in sku_class_stats],
        '占比': [round(stat['percentage'], 1) for stat in sku_class_stats],
        '销售占比': [round(stat['sales_percentage'], 1) for stat in sku_class_stats]
    }
    
    # 4. 客户分类
    customer_sales = df.groupby('客户编号')['订货量'].sum().reset_index()
    customer_sales = customer_sales.sort_values(by='订货量', ascending=False).reset_index(drop=True)
    customer_sales['累计销售量'] = customer_sales['订货量'].cumsum()
    customer_sales['累计销售占比'] = customer_sales['累计销售量'] / customer_sales['订货量'].sum() * 100
    
    total_customers = len(customer_sales)
    customer_sales['分类'] = 'C'
    customer_sales.loc[0:int(total_customers * a_threshold), '分类'] = 'A'
    customer_sales.loc[int(total_customers * a_threshold):int(total_customers * b_threshold), '分类'] = 'B'
    
    customer_class_stats = []
    for class_name in ['A类', 'B类', 'C类']:
        class_data = customer_sales[customer_sales['分类'] == class_name[0]]
        customer_class_stats.append({
            'class': class_name,
            'count': len(class_data),
            'percentage': len(class_data) / total_customers * 100
        })
    
    customer_classes = {
        '分类': [stat['class'] for stat in customer_class_stats],
        '数量': [stat['count'] for stat in customer_class_stats],
        '占比': [round(stat['percentage'], 1) for stat in customer_class_stats]
    }
    
    # 5. EIQ分析
    # 订单量(I)：每个订单的订货量
    order_quantity = df.groupby('订单编号')['订货量'].sum().reset_index()
    avg_order_quantity = order_quantity['订货量'].mean()
    
    # 品项数(E)：每个订单包含的SKU数量
    order_sku_count = df.groupby('订单编号')['SKU编号'].nunique().reset_index()
    avg_order_items = order_sku_count['SKU编号'].mean()
    
    # 订货量(Q)：每个SKU的平均订货量
    sku_avg_quantity = df.groupby('SKU编号')['订货量'].mean().reset_index()
    avg_sku_quantity = sku_avg_quantity['订货量'].mean()
    
    eiq_data = {
        '指标': ['订单平均总量', '订单平均品项数', 'SKU平均订货量'],
        '数值': [float(round(avg_order_quantity, 2)), float(round(avg_order_items, 2)), float(round(avg_sku_quantity, 2))]
    }
    
    # 6. 获取前20个A类SKU供选择
    top_skus = [int(x) for x in sku_sales[sku_sales['分类'] == 'A']['SKU编号'].head(20).tolist()]
    
    # 6. 分拣时间分析
    # 计算不同类型订单的分拣时间统计
    order_type_stats = df.groupby('订单类型').agg({
        '分拣时间': ['mean', 'median', 'max', 'count'],
        '准时完成': ['sum', 'mean']
    }).round(2)
    
    order_type_stats.columns = ['平均分拣时间', '中位数分拣时间', '最大分拣时间', '订单数量', '准时完成数量', '准时完成率']
    order_type_stats = order_type_stats.reset_index()
    
    # 计算总体准时完成率
    total_on_time = df['准时完成'].sum()
    total_orders = len(df)
    overall_on_time_rate = round((total_on_time / total_orders) * 100, 2)
    
    # 准备分拣时间数据
    sorting_data = {
        '订单类型': [str(x) for x in order_type_stats['订单类型'].tolist()],
        '平均分拣时间': [float(x) for x in order_type_stats['平均分拣时间'].tolist()],
        '中位数分拣时间': [float(x) for x in order_type_stats['中位数分拣时间'].tolist()],
        '最大分拣时间': [float(x) for x in order_type_stats['最大分拣时间'].tolist()],
        '订单数量': [int(x) for x in order_type_stats['订单数量'].tolist()],
        '准时完成数量': [int(x) for x in order_type_stats['准时完成数量'].tolist()],
        '准时完成率': [round(rate * 100, 2) for rate in order_type_stats['准时完成率'].tolist()]
    }
    
    result = {
        'monthly_sales': monthly_sales,
        'quarterly_sales': quarterly_sales,
        'hourly_orders': hourly_orders,
        'sku_classes': sku_classes,
        'customer_classes': customer_classes,
        'eiq_data': eiq_data,
        'top_skus': top_skus,
        'sorting_data': sorting_data,
        'overall_on_time_rate': overall_on_time_rate,
        'total_orders': int(total_orders),
        'total_on_time': int(total_on_time)
    }
    
    return convert_to_native_types(result)

# SARIMA销售预测
def sarima_forecast(sku_id, forecast_weeks=52):
    """使用SARIMA模型对指定SKU进行销售预测"""
    global global_data
    
    if global_data is None:
        global_data = load_and_merge_data()
    
    df = global_data
    
    # 按周聚合SKU销售数据
    sku_data = df[df['SKU编号'] == sku_id]
    weekly_sales = sku_data.groupby(['年份', '周'])['订货量'].sum().reset_index()
    
    if weekly_sales.empty:
        return {
            'error': f'没有找到SKU {sku_id} 的销售数据'
        }
    
    # 创建日期索引
    weekly_sales['日期'] = pd.to_datetime(weekly_sales['年份'].astype(str) + '-W' + weekly_sales['周'].astype(str) + '-1', format='%Y-W%U-%w')
    weekly_sales = weekly_sales.sort_values('日期')
    
    # 转换为时间序列
    ts = weekly_sales.set_index('日期')['订货量']
    
    try:
        # 拟合SARIMA模型
        model = SARIMAX(ts, order=(1, 1, 1), seasonal_order=(1, 1, 1, 52))
        model_fit = model.fit(disp=False)
        
        # 预测未来一年
        forecast = model_fit.forecast(steps=forecast_weeks)
        
        # 准备历史数据和预测数据
        # 确保索引是DatetimeIndex
        if hasattr(ts.index, 'strftime'):
            history_dates = ts.index.strftime('%Y-%m-%d').tolist()
        else:
            # 如果索引不是DatetimeIndex，使用原始的日期列
            history_dates = weekly_sales['日期'].dt.strftime('%Y-%m-%d').tolist()
        
        history_values = ts.values.tolist()
        
        if hasattr(forecast.index, 'strftime'):
            forecast_dates = forecast.index.strftime('%Y-%m-%d').tolist()
        else:
            # 生成未来日期
            last_date = weekly_sales['日期'].max()
            forecast_dates = [(last_date + timedelta(weeks=i+1)).strftime('%Y-%m-%d') for i in range(forecast_weeks)]
        
        forecast_values = forecast.values.tolist()
        
        # 转换为标准Python类型
        history_values = [float(val) for val in history_values]
        forecast_values = [float(val) for val in forecast_values]
        
        return convert_to_native_types({
            'sku_id': int(sku_id),
            'history': {
                'dates': history_dates,
                'values': history_values
            },
            'forecast': {
                'dates': forecast_dates,
                'values': forecast_values
            }
        })
    except Exception as e:
        return {
            'error': f'模型拟合失败: {str(e)}'
        }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/data')
def get_data():
    data = load_data()
    return jsonify(data)

@app.route('/forecast/<int:sku_id>')
def get_forecast(sku_id):
    result = sarima_forecast(sku_id)
    return jsonify(result)

if __name__ == '__main__':
    app.run(debug=True, host='localhost', port=5000)