import streamlit as st
import pandas as pd
import re
from collections import Counter
import os
from datetime import datetime


# --------------------------
# 1. 核心功能：地址清洗函数（严格匹配您的需求）
# --------------------------
def clean_location(location):
    """
    地址清洗规则（100%匹配需求）：
    - 保留：小区/楼宇主体+分区、知名地标完整名称（如工人体育场）
    - 删除：门牌号、楼号、单元、楼层、房间号、冗余前缀（北京市朝阳区）
    - 区分：相似地点（国贸商场≠国贸写字楼）
    """
    if pd.isna(location) or str(location).strip() == '':
        return '未知地点'

    location_str = str(location).strip()

    # 保留知名地标完整名称（工人体育场、国家体育场等）
    landmarks = ['工人体育场', '国家体育场', '农业展览馆', '工业展览馆',
                 '798艺术区', '国贸商场', '国贸写字楼', '建外SOHO',
                 '朝阳公园', '红领巾公园', '今日美术馆']
    for landmark in landmarks:
        if landmark in location_str:
            # 去除前缀和后续详细信息
            location_str = re.sub(r'^北京市朝阳区|^朝阳区|^北京市', '', location_str).strip()
            location_str = re.sub(f'{landmark}\\s*[0-9号楼单元层房室门].*', landmark, location_str)
            return location_str

    # 保留小区+分区（如瑞平家园 D 区、珠江罗马嘉园 - 西区）
    district_pattern = r'([东西南北中ABCD][区]|[一二三四五六七八九十][区])'
    if re.search(district_pattern, location_str):
        match = re.match(r'(.+?' + district_pattern + r')\s*[0-9号楼单元层房室].*', location_str)
        if match:
            cleaned = re.sub(r'^北京市朝阳区|^朝阳区|^北京市', '', match.group(1)).strip()
            return cleaned

    # 保留常见后缀地点（小区、家园、大厦等）
    suffixes = ['小区', '家园', '花园', '公寓', '大厦', '写字楼', '商场', '酒店', '场馆']
    for suffix in suffixes:
        if suffix in location_str:
            match = re.match(r'(.+?' + suffix + r')\s*[0-9号楼单元层房室].*', location_str)
            if match:
                cleaned = re.sub(r'^北京市朝阳区|^朝阳区|^北京市', '', match.group(1)).strip()
                return cleaned

    # 标记无效地址
    if re.match(r'^北京市朝阳区\s*$|^朝阳区\s*$|^北京市\s*$|^[0-9号楼单元层房室门\s]+$', location_str):
        return '未知地点'

    # 最终清理
    cleaned = re.sub(r'\s*[0-9]+\s*[号楼单元层房室门号].*', '', location_str)
    cleaned = re.sub(r'^北京市朝阳区|^朝阳区|^北京市', '', cleaned).strip()
    return cleaned if cleaned and len(cleaned) >= 2 else '未知地点'


# --------------------------
# 2. 核心功能：数据处理主函数
# --------------------------
def process_complaints(original_df):
    """处理原始工单数据，生成3个Sheet所需内容"""
    # 步骤1：地址清洗
    original_df['清洗后地址'] = original_df['投诉地点'].apply(clean_location)

    # 步骤2：生成Sheet1（小区投诉统计）
    valid_df = original_df[original_df['清洗后地址'] != '未知地点']
    location_stats = valid_df['清洗后地址'].value_counts().reset_index()
    location_stats.columns = ['小区/地点', '投诉数量']
    total_valid = len(valid_df)
    location_stats['占比(%)'] = (location_stats['投诉数量'] / total_valid * 100).round(2)
    sheet1_data = location_stats.sort_values('投诉数量', ascending=False).reset_index(drop=True)

    # 步骤3：生成Sheet2（地址空白投诉明细）
    unknown_df = original_df[original_df['清洗后地址'] == '未知地点']
    sheet2_data = unknown_df.drop(columns=['清洗后地址'], errors='ignore')  # 删除临时列

    # 步骤4：生成Sheet3（数据统计摘要）
    total_all = len(original_df)
    total_valid = len(valid_df)
    total_unknown = len(unknown_df)
    total_locations = len(sheet1_data)
    avg_per_location = total_valid / total_locations if total_locations > 0 else 0
    max_complaints = sheet1_data['投诉数量'].max() if len(sheet1_data) > 0 else 0
    top10_total = sheet1_data['投诉数量'].head(10).sum()
    top10_ratio = (top10_total / total_valid * 100) if total_valid > 0 else 0

    sheet3_data = pd.DataFrame({
        '统计指标': [
            '总投诉工单数量', '有效地址投诉数量', '地址空白投诉数量',
            '涉及小区/地点总数', '平均每小区投诉次数', '最高单小区投诉次数',
            'Top10小区投诉占比(%)', '处理时间'
        ],
        '数值': [
            total_all, total_valid, total_unknown,
            total_locations, round(avg_per_location, 2), max_complaints,
            round(top10_ratio, 2), datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ]
    })

    return sheet1_data, sheet2_data, sheet3_data


# --------------------------
# 3. 可视化界面（Streamlit）
# --------------------------
def main():
    # 页面配置
    st.set_page_config(
        page_title="朝阳区网络投诉工单处理工具",
        page_icon="📊",
        layout="wide"
    )

    # 标题与说明
    st.title("📊 朝阳区网络投诉工单地址统计处理工具")
    st.markdown("""
    ### 使用说明（仅需3步）
    1. **导入原始文件**：上传包含"投诉地点"列的Excel工单文件（如"投诉工单汇总.xls"）
    2. **开始处理**：点击按钮，自动执行地址清洗和统计
    3. **导出结果**：下载包含3个Sheet的Excel结果文件
    """)
    st.divider()

    # 步骤1：上传原始文件
    st.subheader("1️⃣ 导入原始工单文件")
    uploaded_file = st.file_uploader(
        "请上传Excel格式的原始工单文件（支持.xls/.xlsx）",
        type=['xls', 'xlsx'],
        help="文件需包含'投诉地点'列，所有数据默认视为朝阳区工单"
    )

    if uploaded_file is not None:
        # 读取原始文件
        try:
            original_df = pd.read_excel(uploaded_file)

            # 验证是否包含"投诉地点"列
            if '投诉地点' not in original_df.columns:
                st.error("❌ 上传的文件中未找到'投诉地点'列，请检查文件格式！")
                return

            # 显示原始数据概览
            st.success(f"✅ 成功读取文件！原始工单总数：{len(original_df)} 条")
            with st.expander("查看原始数据预览（前5行）"):
                st.dataframe(original_df.head(), use_container_width=True)

            st.divider()

            # 步骤2：开始处理数据
            st.subheader("2️⃣ 开始处理数据")
            if st.button("🚀 点击开始处理", type="primary", use_container_width=True):
                with st.spinner("正在处理数据...（地址清洗→统计分析→生成结果）"):
                    # 执行处理逻辑
                    sheet1, sheet2, sheet3 = process_complaints(original_df)

                    # 显示处理结果概览
                    st.success("🎉 数据处理完成！")
                    st.subheader("处理结果概览")

                    # 分栏显示关键统计
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("总工单数量", sheet3.loc[0, '数值'])
                        st.metric("有效地址工单", sheet3.loc[1, '数值'])
                    with col2:
                        st.metric("未知地点工单", sheet3.loc[2, '数值'])
                        st.metric("涉及小区/地点数", sheet3.loc[3, '数值'])
                    with col3:
                        st.metric("最高单小区投诉", f"{sheet3.loc[5, '数值']} 次")
                        st.metric("Top10小区占比", f"{sheet3.loc[6, '数值']}%")

                    # 显示各Sheet预览
                    with st.expander("查看小区投诉统计（Top10）"):
                        st.dataframe(sheet1.head(10), use_container_width=True)
                    with st.expander("查看地址空白投诉明细（前5行）"):
                        if len(sheet2) > 0:
                            st.dataframe(sheet2.head(), use_container_width=True)
                        else:
                            st.info("✅ 无地址空白的投诉工单")
                    with st.expander("查看数据统计摘要"):
                        st.dataframe(sheet3, use_container_width=True)

                    st.divider()

                    # 步骤3：导出结果文件
                    st.subheader("3️⃣ 导出结果文件")

                    # 生成Excel文件（在内存中）
                    from io import BytesIO
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        sheet1.to_excel(writer, sheet_name='小区投诉统计', index=False)
                        sheet2.to_excel(writer, sheet_name='地址空白投诉明细', index=False)
                        sheet3.to_excel(writer, sheet_name='数据统计摘要', index=False)
                    output.seek(0)  # 重置文件指针
                    # 下载按钮
                    result_filename = f"朝阳区投诉工单统计结果_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
                    st.download_button(
                        label="💾 下载Excel结果文件",
                        data=output,
                        file_name=result_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

        except Exception as e:
            st.error(f"❌ 处理过程中出错：{str(e)}")
            st.info("请检查文件格式是否正确，或联系管理员排查问题")


if __name__ == "__main__":
    main()