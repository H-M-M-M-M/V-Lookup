import streamlit as st
import pandas as pd
from io import BytesIO
from functools import reduce

st.title("📁 多文件多Sheet按SN保留最新记录并合并")

uploaded_files = st.file_uploader("请上传一个或多个 Excel 文件", type=["xlsx"], accept_multiple_files=True)

def find_column(columns, keywords):
    """模糊匹配列名，返回第一个匹配的列"""
    for col in columns:
        col_clean = col.strip().lower().replace(" ", "")
        for kw in keywords:
            if kw in col_clean:
                return col
    return None

if uploaded_files:
    sn_dfs = []

    for file in uploaded_files:
        all_sheets = pd.read_excel(file, sheet_name=None)
        st.write(f"📂 正在处理文件：**{file.name}**")

        for sheet_name, df in all_sheets.items():
            st.write(f" 📄 Sheet：**{sheet_name}**")

            if df.empty:
                st.warning(f" ⚠️ Sheet「{sheet_name}」为空，跳过。")
                continue

            sn_col = find_column(df.columns, ['sn', 'serialnumber', 'sfc'])
            date_col = find_column(df.columns, ['testdate', 'date'])
            time_col = find_column(df.columns, ['testtime', 'time'])

            if sn_col and date_col and time_col:
                df['DateTime'] = pd.to_datetime(
                    df[date_col].astype(str) + ' ' + df[time_col].astype(str),
                    errors='coerce'
                )
                df = df.dropna(subset=['DateTime'])

                df = df.rename(columns={sn_col: 'SN'})

                df_latest = df.sort_values('DateTime').groupby('SN', as_index=False).last()

                prefix = f"{file.name}_{sheet_name}"
                df_latest = df_latest.rename(columns=lambda col: f"{prefix}_{col}" if col != 'SN' else col)

                sn_dfs.append(df_latest)
            else:
                st.warning(f" ⚠️ 跳过 Sheet「{sheet_name}」：未检测到 SN、Date、Time 列。")

    if sn_dfs:
        merged_df = reduce(lambda left, right: pd.merge(left, right, on='SN', how='outer'), sn_dfs)

        st.success("✅ 所有文件处理完成！以下是合并结果预览：")
        st.dataframe(merged_df.head(20))

        # 转为 Excel
        def convert_df_to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='汇总结果')
            return output.getvalue()

        excel_data = convert_df_to_excel(merged_df)

        st.download_button(
            label="📥 下载汇总结果 Excel",
            data=excel_data,
            file_name="SN_汇总结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("❌ 没有找到有效数据，未能生成结果。")
