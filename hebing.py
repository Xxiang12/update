import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import io

st.set_page_config(page_title="多表合并", page_icon="📊")
st.title("📊 多表合并工具")
st.write("请上传多个表头一致的 Excel 文件。")

# 表头行配置
header_row = st.number_input("请输入表头所在的行号 (通常第一行是表头填1，如果有大标题填2)", min_value=1, value=2)

uploaded_files = st.file_uploader("选择你的 Excel 文件（可多选）", type=["xlsx", "xls"], accept_multiple_files=True)


def copy_style(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)


if uploaded_files:
    if st.button("开始合并表格"):
        out_wb = openpyxl.Workbook()
        out_ws = out_wb.active
        out_ws.title = "合并数据"

        is_first_file = True
        current_out_row = 1

        progress_bar = st.progress(0)

        for i, file in enumerate(uploaded_files):
            try:
                # 从内存中读取文件
                src_wb = openpyxl.load_workbook(file)
                src_ws = src_wb.active
                max_row = src_ws.max_row
                max_col = src_ws.max_column

                if is_first_file:
                    for r in range(1, header_row + 1):
                        for c in range(1, max_col + 1):
                            src_cell = src_ws.cell(row=r, column=c)
                            out_cell = out_ws.cell(row=r, column=c)
                            out_cell.value = src_cell.value
                            copy_style(src_cell, out_cell)
                    # 复制列宽
                    for col_idx in range(1, max_col + 1):
                        col_letter = get_column_letter(col_idx)
                        if col_letter in src_ws.column_dimensions:
                            out_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width

                    current_out_row = header_row + 1
                    is_first_file = False

                # 数据区截断：从表头下一行开始，到倒数第二行结束（跳过最后一行）
                for r in range(header_row + 1, max_row):
                    for c in range(1, max_col + 1):
                        src_cell = src_ws.cell(row=r, column=c)
                        out_cell = out_ws.cell(row=current_out_row, column=c)
                        out_cell.value = src_cell.value
                        copy_style(src_cell, out_cell)

                    if r in src_ws.row_dimensions:
                        out_ws.row_dimensions[current_out_row].height = src_ws.row_dimensions[r].height

                    current_out_row += 1

                progress_bar.progress((i + 1) / len(uploaded_files))

            except Exception as e:
                st.error(f"处理文件 {file.name} 时出错: {e}")

        # 保存到内存
        output_buffer = io.BytesIO()
        out_wb.save(output_buffer)
        excel_data = output_buffer.getvalue()

        st.success("✅ 合并完成！已保留原表格式并剔除了每个表的最后一行。")
        st.download_button(
            label="📥 点击下载合并后的表格",
            data=excel_data,
            file_name="合并结果_新生成.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )