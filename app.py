import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile
import os

st.set_page_config(page_title="Walmart WFS 货件申请生成工具 v2", layout="wide", page_icon="📦")

st.title("📦 Walmart WFS 货件申请生成工具 (模板保护版)")
st.markdown("本程序已锁定 GitHub 仓库中的 Walmart 官方 WFS 原始模板。处理时将直接在模板指定单元格内填入数据，**完美保留原表格的所有公式、样式与数据校验格式**。")

# 检查 GitHub 仓库中是否存在模板文件
TEMPLATE_FILENAME = "WFS_M.xlsx"
template_exists = os.path.exists(TEMPLATE_FILENAME)

# 侧边栏：文件上传区
with st.sidebar:
    st.header("1. 上传基础数据")
    inventory_file = st.file_uploader("上传库存信息表 (inventory.xlsx)", type=['xlsx', 'csv'])
    
    st.header("2. 上传当前装箱单")
    packing_files = st.file_uploader("上传装箱单文件 (支持多选)", type=['xlsx', 'csv'], accept_multiple_files=True)
    
    st.header("3. 模板状态检查")
    if template_exists:
        st.success(f"🟢 已在本地检测到 WFS 官方模板 ({TEMPLATE_FILENAME})")
    else:
        st.error(f"🔴 未在仓库中找到 {TEMPLATE_FILENAME}，请先将模板上传至 GitHub 根目录！")

# 核心处理模块
if inventory_file and packing_files and template_exists:
    if st.button("🚀 开始保持格式处理", type="primary"):
        with st.spinner('正在精准写入模板，请稍候...'):
            try:
                # 1. 读取用户上传的库存表
                if inventory_file.name.endswith('.csv'):
                    inventory_df = pd.read_csv(inventory_file, dtype={'GTIN': str, 'SKU': str})
                else:
                    inventory_df = pd.read_excel(inventory_file, dtype={'GTIN': str, 'SKU': str})
                
                # 清洗库存表的两列，防止前后空格导致匹配失败
                inventory_df['SKU'] = inventory_df['SKU'].astype(str).str.strip()
                
                summary_data = []
                generated_files = {}

                # 2. 逐个处理装箱单
                for pack_file in packing_files:
                    if pack_file.name.endswith('.csv'):
                        pack_df = pd.read_csv(pack_file, header=3)
                    else:
                        pack_df = pd.read_excel(pack_file, header=3)

                    # 2.1 提取核心业务字段
                    order_id = pack_df['出库单号'].dropna().iloc[0] if '出库单号' in pack_df.columns and not pack_df['出库单号'].dropna().empty else "未知单号"
                    valid_boxes = pack_df['箱号'].dropna().unique()
                    box_count = len(valid_boxes)
                    
                    weights = pd.to_numeric(pack_df['重量'], errors='coerce').dropna()
                    max_weight = weights.max() if not weights.empty else 0
                    min_weight = weights.min() if not weights.empty else 0
                    
                    # 2.2 按"平台SKU"进行数量透视汇总
                    grouped_pack = pack_df.groupby('平台SKU', as_index=False)['数量'].sum()
                    total_qty = grouped_pack['数量'].sum()
                    
                    # 记录汇总报告行
                    summary_data.append({
                        "装箱单文件名": pack_file.name,
                        "出库单号": order_id,
                        "总出运数量": total_qty,
                        "装箱数": box_count,
                        "最重重量(kg)": f"{max_weight:.2f}",
                        "最轻重量(kg)": f"{min_weight:.2f}"
                    })
                    
                    # 3. 使用 openpyxl 载入本地官方 WFS 模板文件（实现不破坏格式的写入）
                    wb = openpyxl.load_workbook(TEMPLATE_FILENAME)
                    ws = wb.active  # 默认使用第一个活动工作表
                    
                    # 【安全机制】先清除第3行往后的历史残留数据值（如果有的话），但保留单元格本身的格式和下拉菜单
                    if ws.max_row >= 3:
                        for r in range(3, ws.max_row + 1):
                            for c in range(1, 10):
                                ws.cell(row=r, column=c).value = None
                    
                    # 4. 开始精准按行写入 WFS 模板
                    start_row = 3
                    for idx, row in grouped_pack.iterrows():
                        sku = str(row['平台SKU']).strip()
                        qty = int(row['数量'])
                        
                        # 从库存表中查找对应的 GTIN 和 Item name
                        match = inventory_df[inventory_df['SKU'] == sku]
                        
                        if not match.empty:
                            gtin_val = match.iloc[0]['GTIN']
                            if pd.isna(gtin_val):
                                gtin = ""
                            else:
                                # 清除可能因浮点数产生的 .0，并强制补齐为沃尔玛标准的 14 位长字符串
                                gtin = str(gtin_val).split('.')[0].strip().zfill(14)
                            item_name = str(match.iloc[0]['Item name']).strip()
                        else:
                            gtin = "未匹配到GTIN"
                            item_name = "未匹配到产品名称"
                        
                        current_row = start_row + idx
                        
                        # 按列填充数据 (1对应A列，2对应B列...)
                        ws.cell(row=current_row, column=1, value="GTIN")                  # A: Product type ID
                        ws.cell(row=current_row, column=2, value=gtin)                    # B: Product ID
                        ws.cell(row=current_row, column=3, value=sku)                     # C: SKU
                        ws.cell(row=current_row, column=4, value=item_name)               # D: Item Description
                        ws.cell(row=current_row, column=5, value=qty)                     # E: Item Qty (Total # of Sellable Units)
                        ws.cell(row=current_row, column=6, value=1)                       # F: Vendor pack Qty (# of Cases)
                        ws.cell(row=current_row, column=7, value=qty)                     # G: Inner pack Qty (Sellable Units per Case)
                        ws.cell(row=current_row, column=8, value="No Prep Required")      # H: Labelling
                        ws.cell(row=current_row, column=9, value="No Prep Required")      # I: Poly Bags (Clear)

                    # 5. 将写好的 Excel 模板对象保存到内存流中
                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    # 记录该货件对应的出运表格
                    export_name = f"WFS_Upload_{order_id}.xlsx"
                    generated_files[export_name] = output.read()
                    wb.close()

                # 6. 生成一份独立的“处理汇总报告”
                summary_df = pd.DataFrame(summary_data)
                summary_output = io.BytesIO()
                with pd.ExcelWriter(summary_output, engine='openpyxl') as writer:
                    summary_df.to_excel(writer, index=False, sheet_name='处理汇总')
                summary_output.seek(0)
                generated_files["处理汇总报告_Summary.xlsx"] = summary_output.read()

                # 7. 整体打包打包为单个 ZIP 文件供用户一键下载
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for file_name, data in generated_files.items():
                        zip_file.writestr(file_name, data)
                
                st.success(f"✅ 格式无损处理完成！共成功填入 {len(packing_files)} 份 WFS 官方模板。")
                
                # 页面展示摘要
                st.subheader("📊 本次货件处理汇总（可复制核对）")
                st.dataframe(summary_df, use_container_width=True)
                
                # 下载按钮
                st.download_button(
                    label="📦 一键下载打包好的 WFS 官方申请表 (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Walmart_WFS_Templates_Output.zip",
                    mime="application/zip",
                    type="primary"
                )

            except Exception as e:
                st.error(f"❌ 运行中出现错误，请检查装箱单或库存表字段名是否正确。错误摘要：{str(e)}")
else:
    if not template_exists:
        st.warning(f"⚠️ 无法继续：请先在你的 GitHub 代码根目录下放置名为 '{TEMPLATE_FILENAME}' 的官方空白模板表格。")
    else:
        st.info("👈 请在左侧边栏上传必备的库存信息表以及一份或多份装箱单表格。")