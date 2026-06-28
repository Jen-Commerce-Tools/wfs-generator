import streamlit as st
import pandas as pd
import io
import zipfile
import math

st.set_page_config(page_title="Walmart WFS 货件申请生成工具", layout="wide", page_icon="📦")

st.title("📦 Walmart WFS 货件申请生成工具")
st.markdown("上传库存总表、WFS模板以及一份或多份ERP装箱单，系统将自动汇总并生成可直接上传的WFS文件。")

# 侧边栏：文件上传区
with st.sidebar:
    st.header("1. 上传基础文件")
    inventory_file = st.file_uploader("上传库存信息表 (inventory)", type=['xlsx', 'csv'])
    wfs_template_file = st.file_uploader("上传WFS空模板 (WFS template)", type=['xlsx', 'csv'])
    
    st.header("2. 上传装箱单")
    packing_files = st.file_uploader("上传装箱单文件 (支持多选)", type=['xlsx', 'csv'], accept_multiple_files=True)

# 核心处理模块
if inventory_file and wfs_template_file and packing_files:
    if st.button("🚀 开始处理生成", type="primary"):
        with st.spinner('正在处理数据，请稍候...'):
            try:
                # 1. 读取基础表
                # 读取库存表，强制GTIN作为字符串读取以保留前导0
                if inventory_file.name.endswith('.csv'):
                    inventory_df = pd.read_csv(inventory_file, dtype={'GTIN': str})
                else:
                    inventory_df = pd.read_excel(inventory_file, dtype={'GTIN': str})
                
                # 读取WFS模板（无表头模式，为了保留前两行特殊表头）
                if wfs_template_file.name.endswith('.csv'):
                    wfs_df = pd.read_csv(wfs_template_file, header=None)
                else:
                    wfs_df = pd.read_excel(wfs_template_file, header=None)
                wfs_headers = wfs_df.iloc[:2].copy() # 截取前两行双表头
                
                summary_data = []
                generated_files = {}

                # 2. 遍历处理每一个装箱单
                for pack_file in packing_files:
                    # 读取装箱单，跳过前3行，将第4行（index=3）作为表头
                    if pack_file.name.endswith('.csv'):
                        pack_df = pd.read_csv(pack_file, header=3)
                    else:
                        pack_df = pd.read_excel(pack_file, header=3)

                    # 数据清洗与提取
                    # 2.1 提取出库单号
                    order_id = pack_df['出库单号'].dropna().iloc[0] if '出库单号' in pack_df.columns and not pack_df['出库单号'].dropna().empty else "未知单号"
                    
                    # 2.2 计算装箱数 (去重后的有效箱号)
                    valid_boxes = pack_df['箱号'].dropna().unique()
                    box_count = len(valid_boxes)
                    
                    # 2.3 计算最重/最轻重量
                    weights = pd.to_numeric(pack_df['重量'], errors='coerce').dropna()
                    max_weight = weights.max() if not weights.empty else 0
                    min_weight = weights.min() if not weights.empty else 0
                    
                    # 2.4 透视汇总：按"平台SKU"汇总"数量"
                    grouped_pack = pack_df.groupby('平台SKU', as_index=False)['数量'].sum()
                    total_qty = grouped_pack['数量'].sum()
                    
                    # 汇总信息记录
                    summary_data.append({
                        "装箱单文件名": pack_file.name,
                        "出库单号": order_id,
                        "总出运数量": total_qty,
                        "装箱数": box_count,
                        "最重重量": f"{max_weight:.2f}",
                        "最轻重量": f"{min_weight:.2f}"
                    })
                    
                    # 3. 构建当前货件的 WFS Data
                    wfs_rows = []
                    for _, row in grouped_pack.iterrows():
                        sku = str(row['平台SKU']).strip()
                        qty = int(row['数量'])
                        
                        # 在库存表中匹配 SKU
                        match = inventory_df[inventory_df['SKU'].astype(str).str.strip() == sku]
                        
                        if not match.empty:
                            gtin = str(match.iloc[0]['GTIN']).replace('.0', '')
                            # 补齐14位GTIN标准
                            gtin = gtin.zfill(14) if gtin and gtin != 'nan' else ""
                            item_name = str(match.iloc[0]['Item name'])
                        else:
                            gtin = "未匹配到GTIN"
                            item_name = "未匹配到名称"
                            
                        # 按WFS模板列顺序填入数据
                        wfs_rows.append({
                            0: "GTIN",                               # Product type ID
                            1: gtin,                                 # Product ID
                            2: sku,                                  # SKU
                            3: item_name,                            # Item Description
                            4: qty,                                  # Item Qty
                            5: 1,                                    # Vendor pack Qty
                            6: qty,                                  # Inner pack Qty
                            7: "No Prep Required",                   # Labelling
                            8: "No Prep Required"                    # Poly Bags
                        })
                    
                    df_wfs_data = pd.DataFrame(wfs_rows)
                    
                    # 拼接双表头和数据
                    final_wfs_df = pd.concat([wfs_headers, df_wfs_data], ignore_index=True)
                    
                    # 将生成的 DataFrame 转存为 Excel 字节流
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_wfs_df.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
                    output.seek(0)
                    
                    # 记录文件流以备打包
                    export_name = f"WFS_Upload_{order_id}.xlsx"
                    generated_files[export_name] = output.read()

                # 4. 生成汇总报告
                summary_df = pd.DataFrame(summary_data)
                
                # 将汇总报告也存为字节流
                summary_output = io.BytesIO()
                with pd.ExcelWriter(summary_output, engine='openpyxl') as writer:
                    summary_df.to_excel(writer, index=False, sheet_name='Summary')
                summary_output.seek(0)
                generated_files["处理汇总报告_Summary.xlsx"] = summary_output.read()

                # 5. 打包为 ZIP 压缩包
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for file_name, data in generated_files.items():
                        zip_file.writestr(file_name, data)
                
                st.success(f"✅ 处理完成！共成功处理 {len(packing_files)} 份装箱单。")
                
                # 在页面上展示汇总报告
                st.subheader("📊 本次处理汇总信息")
                st.dataframe(summary_df, use_container_width=True)
                
                # 提供一键下载按钮
                st.download_button(
                    label="📦 一键下载所有 WFS 表格与汇总报告 (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="WFS_Shipments_Export.zip",
                    mime="application/zip",
                    type="primary"
                )

            except Exception as e:
                st.error(f"处理过程中出现错误，请检查表头或格式是否被更改。详细报错：{str(e)}")
else:
    st.info("👈 请在左侧边栏上传必备的库存表、WFS模板以及装箱单文件。")