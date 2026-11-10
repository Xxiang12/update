import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel 自动处理工具", page_icon="📊")
st.title("📊 Excel 自动处理工具")
st.write("请上传原始 Excel 表格，系统将自动筛选出【保障】记录，并生成新表供你下载。")

uploaded_file = st.file_uploader("选择你的 Excel 文件", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 读取数据
        df = pd.read_excel(uploaded_file, header=1, dtype=str)

        # 【1：数据筛选】
        # 筛选原表中 '调试保障' 列等于 '保障' 的行
        df_filtered = df[df['调试保障'] == '保障'].copy()

        st.success(f"✅ 文件读取成功！原表共 {len(df)} 行，筛选后保留【保障】记录 {len(df_filtered)} 行。")

        # 【2：定义列映射并新增列】
        column_mapping = {
            '序号': '序号',
            '日期': '日期',
            '时间': '时间',
            '会议名称调试/保障（内容）': '会议名称（内容）',
            '会议承办单位': '承办部门',
            '会议召开地点': '召开地点',
            '会议类别': '会议类别',
            '技术保障人员': '保障人员',
            '会议保障情况': '会议状态',
            '备注': '备注'
        }

        # 提取需要的列并重命名
        new_df = df_filtered[list(column_mapping.keys())].copy()
        new_df.rename(columns=column_mapping, inplace=True)

        # 【 3：新增空白列】
        # 在图二的 '会议状态' 和 '备注' 之间插入新增列
        new_df['会议调试+保障表序号'] = ''

        # 调整列的顺序，让新增列排在备注前面
        target_columns = [
            '序号', '日期', '时间', '会议名称（内容）', '承办部门','召开地点', '会议类别', '保障人员', '会议状态','会议调试+保障表序号', '备注'
        ]
        new_df = new_df[target_columns]

        st.success("✅ 数据处理完成！以下是生成的新表预览：")
        st.dataframe(new_df.head())

        # 准备文件下载
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            new_df.to_excel(writer, index=False)

        excel_data = output_buffer.getvalue()

        st.download_button(
            label="📥 点击下载筛选后的表格",
            data=excel_data,
            file_name="筛选后的保障会议表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except KeyError as e:
        st.error(f"❌ 处理失败：原表中找不到列名 {e}。请检查表头文字（包括全角/半角括号和空格）是否完全一致。")
    except Exception as e:
        st.error(f"❌ 发生未知错误：{e}")
