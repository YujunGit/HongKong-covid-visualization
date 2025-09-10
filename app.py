#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
香港疫情数据可视化大屏 - Flask应用
"""

from flask import Flask, render_template, jsonify
import pandas as pd
import os
import json
from datetime import datetime

app = Flask(__name__)

def load_data():
    """
    加载疫情数据
    """
    try:
        # 获取当前脚本所在目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        excel_file = os.path.join(current_dir, "香港各区疫情数据_20250322.xlsx")
        
        # 读取Excel文件
        df = pd.read_excel(excel_file)
        
        # 确保日期列是datetime类型
        df['报告日期'] = pd.to_datetime(df['报告日期'])
        
        return df
    except Exception as e:
        print(f"加载数据时发生错误：{str(e)}")
        return None

def get_daily_statistics(df):
    """
    获取每日统计数据
    """
    if df is None:
        return None
    
    # 按日期分组，计算每日全港数据
    daily_stats = df.groupby('报告日期').agg({
        '新增确诊': 'sum',
        '累计确诊': 'sum',
        '新增康复': 'sum',
        '累计康复': 'sum',
        '新增死亡': 'sum',
        '累计死亡': 'sum'
    }).reset_index()
    
    # 计算现存确诊
    daily_stats['现存确诊'] = daily_stats['累计确诊'] - daily_stats['累计康复'] - daily_stats['累计死亡']
    
    # 按日期排序
    daily_stats = daily_stats.sort_values('报告日期')
    
    return daily_stats

def get_region_statistics(df):
    """
    获取各地区统计数据
    """
    if df is None:
        return None
    
    # 按地区分组，计算各地区最终数据
    region_stats = df.groupby('地区名称').agg({
        '累计确诊': 'max',
        '累计康复': 'max',
        '累计死亡': 'max',
        '人口': 'first'
    }).reset_index()
    
    # 计算现存确诊和发病率
    region_stats['现存确诊'] = region_stats['累计确诊'] - region_stats['累计康复'] - region_stats['累计死亡']
    region_stats['发病率'] = (region_stats['累计确诊'] / region_stats['人口'] * 100000).round(2)
    
    # 按累计确诊数排序
    region_stats = region_stats.sort_values('累计确诊', ascending=False)
    
    return region_stats

@app.route('/')
def index():
    """
    主页面
    """
    return render_template('dashboard.html')

@app.route('/api/daily_data')
def api_daily_data():
    """
    API: 获取每日数据
    """
    df = load_data()
    daily_stats = get_daily_statistics(df)
    
    if daily_stats is None:
        return jsonify({'error': '数据加载失败'})
    
    # 转换为前端需要的格式
    dates = daily_stats['报告日期'].dt.strftime('%Y-%m-%d').tolist()
    new_cases = daily_stats['新增确诊'].tolist()
    cumulative_cases = daily_stats['累计确诊'].tolist()
    new_recovered = daily_stats['新增康复'].tolist()
    cumulative_recovered = daily_stats['累计康复'].tolist()
    new_deaths = daily_stats['新增死亡'].tolist()
    cumulative_deaths = daily_stats['累计死亡'].tolist()
    active_cases = daily_stats['现存确诊'].tolist()
    
    return jsonify({
        'dates': dates,
        'new_cases': new_cases,
        'cumulative_cases': cumulative_cases,
        'new_recovered': new_recovered,
        'cumulative_recovered': cumulative_recovered,
        'new_deaths': new_deaths,
        'cumulative_deaths': cumulative_deaths,
        'active_cases': active_cases
    })

@app.route('/api/region_data')
def api_region_data():
    """
    API: 获取各地区数据
    """
    df = load_data()
    region_stats = get_region_statistics(df)
    
    if region_stats is None:
        return jsonify({'error': '数据加载失败'})
    
    # 转换为前端需要的格式
    regions = region_stats['地区名称'].tolist()
    cumulative_cases = region_stats['累计确诊'].tolist()
    active_cases = region_stats['现存确诊'].tolist()
    incidence_rates = region_stats['发病率'].tolist()
    
    return jsonify({
        'regions': regions,
        'cumulative_cases': cumulative_cases,
        'active_cases': active_cases,
        'incidence_rates': incidence_rates
    })

@app.route('/api/summary_data')
def api_summary_data():
    """
    API: 获取汇总数据
    """
    df = load_data()
    daily_stats = get_daily_statistics(df)
    
    if daily_stats is None:
        return jsonify({'error': '数据加载失败'})
    
    # 计算关键指标
    total_cases = daily_stats['累计确诊'].iloc[-1]
    total_recovered = daily_stats['累计康复'].iloc[-1]
    total_deaths = daily_stats['累计死亡'].iloc[-1]
    active_cases = daily_stats['现存确诊'].iloc[-1]
    
    # 计算比率
    recovery_rate = (total_recovered / total_cases * 100) if total_cases > 0 else 0
    death_rate = (total_deaths / total_cases * 100) if total_cases > 0 else 0
    
    # 最高单日新增
    max_daily_new = daily_stats['新增确诊'].max()
    max_daily_date = daily_stats.loc[daily_stats['新增确诊'].idxmax(), '报告日期'].strftime('%Y-%m-%d')
    
    return jsonify({
        'total_cases': int(total_cases),
        'total_recovered': int(total_recovered),
        'total_deaths': int(total_deaths),
        'active_cases': int(active_cases),
        'recovery_rate': round(recovery_rate, 2),
        'death_rate': round(death_rate, 2),
        'max_daily_new': int(max_daily_new),
        'max_daily_date': max_daily_date
    })

@app.route('/api/map_data')
def api_map_data():
    """
    API: 获取地图数据
    """
    df = load_data()
    region_stats = get_region_statistics(df)
    
    if region_stats is None:
        return jsonify({'error': '数据加载失败'})
    
    # 英文地区名称到中文的映射
    region_name_mapping = {
        'Central and Western': '中西区',
        'Wan Chai': '湾仔区',
        'Eastern': '东区',
        'Southern': '南区',
        'Yau Tsim Mong': '油尖旺区',
        'Sham Shui Po': '深水埗区',
        'Kowloon City': '九龙城区',
        'Wong Tai Sin': '黄大仙区',
        'Kwun Tong': '观塘区',
        'Kwai Tsing': '葵青区',
        'Tsuen Wan': '荃湾区',
        'Tuen Mun': '屯门区',
        'Yuen Long': '元朗区',
        'North': '北区',
        'Tai Po': '大埔区',
        'Sha Tin': '沙田区',
        'Sai Kung': '西贡区',
        'Islands': '离岛区'
    }
    
    # 转换为地图需要的格式
    map_data = []
    for idx, row in region_stats.iterrows():
        # 查找对应的英文名称
        english_name = None
        for eng_name, chinese_name in region_name_mapping.items():
            if chinese_name == row['地区名称']:
                english_name = eng_name
                break
        
        if english_name:
            map_data.append({
                'name': english_name,  # 使用英文名称匹配GeoJSON文件
                'chinese_name': row['地区名称'],  # 中文名称用于显示
                'value': int(row['累计确诊']),
                'incidence_rate': float(row['发病率']),
                'active_cases': int(row['现存确诊']),
                'population': int(row['人口'])
            })
    
    return jsonify({
        'map_data': map_data,
        'max_cases': int(region_stats['累计确诊'].max()),
        'min_cases': int(region_stats['累计确诊'].min())
    })

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
