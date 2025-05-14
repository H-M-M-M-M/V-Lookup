import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.title("📊 多Sheet按SN保留最新记录并汇总（支持变动列名）")

uploaded_file = st.file_uploader("请上传包含多个Sheet的Excel文件", type=["xlsx"])

# 智能识别 SN 列名
def find_sn_column(columns):
    for col in columns:
        col_clean = col.strip().lower().replace(" ", "")
        if col_clean in ["sn", "serialnumber", "sfc"]:
            return col
    return None

if uploaded_file:
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
    sn_dfs = []

    for sheet_name, df in all_sheets.items():
        st.write(f"📄 正在处理 Sheet：**{sheet_name}**")

        sn_col = find_sn_column(df.columns)
        if sn_col and {'Date', 'Time'}.issubset(df.columns):
            # 合并日期时间列
            df['DateTime'] = pd.to_datetime(
                df['Date'].astype(str) + ' ' + df['Time'].astype(str),
                errors='coerce'
            )
            df = df.dropna(subset=['DateTime'])

            # 标准化 SN 列名为 'SN'
            df = df.rename(columns={sn_col: 'SN'})

            # 按 SN 取最后一条记录
            df_latest = df.sort_values('DateTime').groupby('SN', as_index=False).last()

            # 给列加上 Sheet 前缀，SN 除外
            df_latest = df_latest.rename(columns=lambda col: f"{sheet_name}_{col}" if col != 'SN' else col)

            sn_dfs.append(df_latest)
        else:
            st.warning(f"⚠️ Sheet「{sheet_name}」缺少 SN/Date/Time 列，或格式不正确，已跳过。")

    if sn_dfs:
        # 合并所有 Sheet（横向合并）
        from functools import reduce
        merged_df = reduce(lambda left, right: pd.merge(left, right, on='SN', how='outer'), sn_dfs)

        st.success("✅ 处理完成！以下是结果预览：")
        st.dataframe(merged_df.head(20))

        # 导出为 Excel
        def convert_df_to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='SN汇总结果')
            return output.getvalue()

        excel_data = convert_df_to_excel(merged_df)

        st.download_button(
            label="📥 下载汇总结果 Excel",
            data=excel_data,
            file_name="SN汇总结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("❌ 没有找到包含有效 SN/Date/Time 的 Sheet，未能生成汇总结果。")
