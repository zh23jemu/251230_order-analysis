from flask import Flask, render_template, jsonify
import pandas as pd
import numpy as np
import json
from datetime import datetime, timedelta

app = Flask(__name__)

# 加载并预处理数据
# 这里简化处理，实际应该从原数据重新计算或保存之前的分析结果
def load_data():
    # 模拟数据，实际应该从原数据计算
    monthly_sales = {
        '月份': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
        '订货量': [160841, 110511, 100443, 118711, 149418, 162492, 137737, 93623, 116908, 138681, 174677, 178600]
    }
    
    quarterly_sales = {
        '季度': [1, 2, 3, 4],
        '订货量': [371795, 430621, 348268, 491958]
    }
    
    hourly_orders = {
        '小时': [10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20],
        '订单数': [10, 17, 17, 60, 88, 140, 157, 121, 119, 112, 26]
    }
    
    sku_classes = {
        '分类': ['A类', 'B类', 'C类'],
        '数量': [800, 1201, 1999],
        '占比': [20.0, 30.0, 50.0],
        '销售占比': [48.8, 24.8, 26.4]
    }
    
    customer_classes = {
        '分类': ['A类', 'B类', 'C类'],
        '数量': [20, 31, 50],
        '占比': [19.8, 30.7, 49.5]
    }
    
    eiq_data = {
        '指标': ['订单平均总量', '订单平均品项数', 'SKU平均订货量'],
        '数值': [4653.38, 1289.30, 2.51]
    }
    
    return {
        'monthly_sales': monthly_sales,
        'quarterly_sales': quarterly_sales,
        'hourly_orders': hourly_orders,
        'sku_classes': sku_classes,
        'customer_classes': customer_classes,
        'eiq_data': eiq_data
    }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/data')
def get_data():
    data = load_data()
    return jsonify(data)

if __name__ == '__main__':
    app.run(debug=True, host='localhost', port=5000)