#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
读取香港各区疫情数据Excel文件并计算确诊病例数
"""

import pandas as pd
import os
import numpy as np

def read_excel_data():
    """
    读取香港各区疫情数据Excel文件
    """
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Excel文件路径
    excel_file = os.path.join(current_dir, "香港各区疫情数据_20250322.xlsx")
    
    try:
        # 检查文件是否存在
        if not os.path.exists(excel_file):
            print(f"错误：文件 {excel_file} 不存在")
            return None
        
        # 读取Excel文件
        print("正在读取Excel文件...")
        df = pd.read_excel(excel_file)
        
        # 显示文件基本信息
        print(f"文件读取成功！")
        print(f"总行数：{len(df)}")
        print(f"总列数：{len(df.columns)}")
        print(f"列名：{list(df.columns)}")
        print("-" * 50)
        
        return df
        
    except Exception as e:
        print(f"读取文件时发生错误：{str(e)}")
        return None


def calculate_daily_statistics(df):
    """
    计算每日新增与累计确诊数据
    """
    if df is None:
        print("数据为空，无法计算")
        return None
    
    try:
        # 确保日期列是datetime类型
        df['报告日期'] = pd.to_datetime(df['报告日期'])
        
        # 按日期分组，计算每日全港新增确诊和累计确诊
        daily_stats = df.groupby('报告日期').agg({
            '新增确诊': 'sum',      # 每日新增确诊总数
            '累计确诊': 'sum',      # 每日累计确诊总数
            '新增康复': 'sum',      # 每日新增康复总数
            '累计康复': 'sum',      # 每日累计康复总数
            '新增死亡': 'sum',      # 每日新增死亡总数
            '累计死亡': 'sum'       # 每日累计死亡总数
        }).reset_index()
        
        # 计算每日现存确诊数（累计确诊 - 累计康复 - 累计死亡）
        daily_stats['现存确诊'] = daily_stats['累计确诊'] - daily_stats['累计康复'] - daily_stats['累计死亡']
        
        # 按日期排序
        daily_stats = daily_stats.sort_values('报告日期')
        
        return daily_stats
        
    except Exception as e:
        print(f"计算统计数据时发生错误：{str(e)}")
        return None


def display_daily_statistics(daily_stats):
    """
    显示每日统计数据
    """
    if daily_stats is None:
        print("统计数据为空")
        return
    
    print("\n" + "=" * 100)
    print("香港疫情每日统计数据")
    print("=" * 100)
    
    # 显示前20天的数据
    print("前20天数据：")
    print("-" * 100)
    print(daily_stats.head(20).to_string(index=False))
    
    # 显示总体统计信息
    print("\n" + "=" * 100)
    print("总体统计信息：")
    print("-" * 100)
    print(f"数据时间范围：{daily_stats['报告日期'].min().strftime('%Y-%m-%d')} 至 {daily_stats['报告日期'].max().strftime('%Y-%m-%d')}")
    print(f"总天数：{len(daily_stats)} 天")
    print(f"最高单日新增确诊：{daily_stats['新增确诊'].max():,} 例")
    print(f"最高单日新增确诊日期：{daily_stats.loc[daily_stats['新增确诊'].idxmax(), '报告日期'].strftime('%Y-%m-%d')}")
    print(f"最终累计确诊：{daily_stats['累计确诊'].iloc[-1]:,} 例")
    print(f"最终累计康复：{daily_stats['累计康复'].iloc[-1]:,} 例")
    print(f"最终累计死亡：{daily_stats['累计死亡'].iloc[-1]:,} 例")
    print(f"最终现存确诊：{daily_stats['现存确诊'].iloc[-1]:,} 例")
    
    # 计算平均每日新增确诊
    avg_daily_new = daily_stats['新增确诊'].mean()
    print(f"平均每日新增确诊：{avg_daily_new:.1f} 例")


def display_region_statistics(df):
    """
    显示各地区统计信息
    """
    if df is None:
        print("数据为空，无法计算")
        return
    
    print("\n" + "=" * 100)
    print("香港各地区疫情统计")
    print("=" * 100)
    
    # 按地区分组，计算各地区累计数据
    region_stats = df.groupby('地区名称').agg({
        '累计确诊': 'max',      # 各地区最终累计确诊数
        '累计康复': 'max',      # 各地区最终累计康复数
        '累计死亡': 'max',      # 各地区最终累计死亡数
        '人口': 'first'        # 各地区人口数
    }).reset_index()
    
    # 计算各地区现存确诊数
    region_stats['现存确诊'] = region_stats['累计确诊'] - region_stats['累计康复'] - region_stats['累计死亡']
    
    # 计算各地区发病率（每10万人）
    region_stats['发病率'] = (region_stats['累计确诊'] / region_stats['人口'] * 100000).round(2)
    
    # 按累计确诊数排序
    region_stats = region_stats.sort_values('累计确诊', ascending=False)
    
    print("各地区累计疫情数据（按累计确诊数排序）：")
    print("-" * 100)
    print(region_stats.to_string(index=False))
    
    # 显示疫情最严重的5个地区
    print("\n疫情最严重的5个地区：")
    print("-" * 50)
    top5_regions = region_stats.head(5)
    for idx, row in top5_regions.iterrows():
        print(f"{row['地区名称']}: 累计确诊 {row['累计确诊']:,} 例, 发病率 {row['发病率']:.2f}/10万")

if __name__ == "__main__":
    # 执行读取操作
    print("开始分析香港疫情数据...")
    data = read_excel_data()
    
    if data is not None:
        # 计算每日统计数据
        daily_stats = calculate_daily_statistics(data)
        
        # 显示每日统计数据
        display_daily_statistics(daily_stats)
        
        # 显示各地区统计数据
        display_region_statistics(data)
        
        print("\n" + "=" * 100)
        print("数据分析完成！")
    else:
        print("数据读取失败！")
