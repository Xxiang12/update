import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import io

st.set_page_config(page_title="多表合并", page_icon="📊")
st.title("📊 多表合并工具")
st.write("请上传多个表头一致的 Excel 文件。")

# 表头行配置
header_row = st.number_input("请输入表头所在的行号 (第一行是表头填1，如果有大标题填2)", min_value=1, value=2)

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
                src_wb = openpyxl.load_workbook(file, data_only=True) # data_only 处理公式
                src_ws = src_wb.active
                
                # 【 1：寻找真实的最后一行】
                # 从最大行往上倒着找，找到第一个有内容的行才算数，忽略底部的空白行
                real_max_row = src_ws.max_row
                while real_max_row > header_row:
                    row_values = [src_ws.cell(row=real_max_row, column=c).value for c in range(1, src_ws.max_column + 1)]
                    # 如果这一行所有单元格都是 None 或空字符串，说明是空白行，往上继续找
                    if all(v is None or str(v).strip() == '' for v in row_values):
                        real_max_row -= 1
                    else:
                        break # 找到真实有数据的行了，跳出循环

                max_col = src_ws.max_column

                # 如果是第一个文件，复制表头和大标题
                if is_first_file:
                    for r in range(1, header_row + 1):
                        for c in range(1, max_col + 1):
                            src_cell = src_ws.cell(row=r, column=c)
                            out_cell = out_ws.cell(row=r, column=c)
                            out_cell.value = src_cell.value
                            copy_style(src_cell, out_cell)
                    
                    for col_idx in range(1, max_col + 1):
                        col_letter = get_column_letter(col_idx)
                        if col_letter in src_ws.column_dimensions:
                            out_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                    
                    current_out_row = header_row + 1
                    is_first_file = False

                # 【 2：精确截断】
                # 数据区是从 表头下一行开始，到 (真实最后一行 - 1) 结束
                # 比如真实最后一行是 20，range(表头+1, 20) 只会循环到 19，完美去掉最后一行
                for r in range(header_row + 1, real_max_row):
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

        # 保存
        output_buffer = io.BytesIO()
        out_wb.save(output_buffer)
        excel_data = output_buffer.getvalue()

        st.success("✅ 合并完成！已保留原表格式并剔除了每个表的最后一行。")
        st.download_button(
            label="📥 点击下载合并后的表格",
            data=excel_data,
            file_name="合并后表格.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
