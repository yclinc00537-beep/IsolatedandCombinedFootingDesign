import streamlit as st
import pandas as pd
import math
from docx import Document
import io

# ==========================================
# 1. 鋼筋資料庫與竹節換算函數 (全域最頂端定義)
# ==========================================
REBAR_DB = {
    "#4": 1.267,
    "#5": 1.979,
    "#6": 2.865,
    "#7": 3.871,
    "#8": 5.067
}

def convert_area_to_spacing(As_req, width, bar_size):
    """ 將所需鋼筋面積轉換為 @ 間距 (四捨五入至整數公分，上限30cm) """
    ab = REBAR_DB[bar_size]
    if As_req <= 0:
        return f"{bar_size}@30"
    qty = As_req / ab
    spacing = 30 if qty <= 0 else math.floor(width / qty)
    spacing = min(max(spacing, 10), 30)
    return f"{bar_size}@{spacing}"

def convert_area_to_qty(As_req, bar_size):
    """ 將所需鋼筋面積轉換為短柱(Pier)所需的固定鋼筋支數 """
    ab = REBAR_DB[bar_size]
    qty = math.ceil(As_req / ab)
    if qty < 4: qty = 4
    if qty % 2 != 0: qty += 1
    return f"{qty}-{bar_size}"

# ==========================================
# 2. 核心結構力學與檢核邏輯 (全域最頂端定義)
# ==========================================

def check_and_calculate_isolated(c1, c2, h, fc, B, L_size, D, L_load, q_all, h_earth, H, Mx, My):
    """ 獨立基腳核心計算：地壓、剪力、與所需鋼筋面積 """
    if h_earth < 1.2:
        return {"success": False, "q_act": 0.0, "h_pier": 0.0, "reason": f"規範不符！基礎板上方距地表深度 {h_earth:.2f} m 🛑低於規定的 1.2 m！"}
    
    h_pier = h_earth + 0.3
    h_total_m = h_pier + (h / 100.0)
    q_self_weight = 0.2 
    P_service = D + L_load
    A_footing = B * L_size
    
    M_x_total_cm = (Mx * 100.0) + (H * h_total_m * 100.0)
    M_y_total_cm = (My * 100.0) + (H * h_total_m * 100.0)
    
    q_act_max = (P_service / A_footing) + (6 * M_x_total_cm / (B * B * L_size)) + (6 * M_y_total_cm / (B * L_size * L_size)) + q_self_weight
    
    if q_act_max > q_all:
        return {"success": False, "q_act": q_act_max, "h_pier": h_pier, "reason": f"地壓超限！最大應力 {q_act_max:.2f} kgf/cm² > 土壤承載力 {q_all:.2f} kgf/cm²。"}
        
    pu = 1.2 * D + 1.6 * L_load
    d = h - 7.5 - 1.0
    if d <= 0: 
        return {"success": False, "q_act": q_act_max, "h_pier": h_pier, "reason": "版厚過小"}
        
    x0, y0 = c1 + d, c2 + d
    b0, A0 = 2 * (x0 + y0), x0 * y0
    qu = pu / A_footing
    Vu = pu - (qu * A0)
    Vc = min(0.53 * (1 + 2/(max(c1,c2)/min(c1,c2))), 1.06) * math.sqrt(fc) * b0 * d
    phi_Vc = 0.75 * Vc
    
    Mu_isolated = qu * L_size * ((B - c1) / 2.0)**2 / 2.0
    phi_m = 0.90
    R_m = Mu_isolated / (phi_m * L_size * (d**2))
    fy_default = 4200
    rho = (0.85 * fc / fy_default) * (1.0 - math.sqrt(1.0 - (2.0 * R_m / (0.85 * fc)))) if (2.0*R_m/(0.85*fc)) < 1.0 else 0.002
    
    As_x_req = max(rho * L_size * d, 0.0018 * L_size * h)
    As_y_req = 0.0018 * B * h
    As_pier_req = 0.01 * (c1 * c2)
    
    return {
        "success": True, "q_act": q_act_max, "pu": pu, "d": d, "b0": b0, "h_pier": h_pier, 
        "Vu": Vu, "phi_Vc": phi_Vc, "is_safe": Vu <= phi_Vc, "reason": "OK",
        "As_x_top_req": 0.0, "As_x_bot_req": As_x_req,
        "As_y_top_req": 0.0, "As_y_bot_req": As_y_req,
        "As_pier_req": As_pier_req
    }

def calculate_trapezoid_geometry(P1, P2, S, a1, q_all, L):
    R = P1 + P2
    x_bar = (P2 * S) / R
    Xc = a1 + x_bar
    A_req = R / q_all
    try:
        B2 = (2 * A_req / L) * ((3 * Xc - L) / L)
        B1 = (2 * A_req / L) - B2
    except ZeroDivisionError:
        return {"success": False, "msg": "長度不能為 0"}
    if B1 <= 0 or B2 <= 0:
        return {"success": False, "msg": "梯形幾何不成立，形心超出範圍。請調整 L。"}
    return {"success": True, "B1": B1, "B2": B2, "A_req": A_req, "R": R, "Xc": Xc}

def calculate_combined_details(P1_D, P1_L, P2_D, P2_L, S, a1, q_all, L, h, fc, fy, B1, B2):
    P1u = 1.2 * P1_D + 1.6 * P1_L
    P2u = 1.2 * P2_D + 1.6 * P2_L
    Ru = P1u + P2u
    A_footing = ((B1 + B2) * L) / 2.0
    d = h - 7.5 - 1.5
    qu1 = (Ru / A_footing) * B1
    qu2 = (Ru / A_footing) * B2
    A_quad = (qu1 - qu2) / (2.0 * L)
    B_quad = -qu1
    C_quad = P1u
    discriminant = B_quad**2 - 4 * A_quad * C_quad
    x_m_max = (-B_quad - math.sqrt(discriminant)) / (2.0 * A_quad) if discriminant >= 0 else a1 + S/2.0
    Mu_neg = abs(P1u * (x_m_max - a1) - (qu1 * (x_m_max**2) / 2.0 - (qu1 - qu2) * (x_m_max**3) / (6.0 * L)))
    phi_m = 0.90
    B_at_m = B1 - (B1 - B2) * (x_m_max / L)
    R_m = Mu_neg / (phi_m * B_at_m * (d**2))
    rho = (0.85 * fc / fy) * (1.0 - math.sqrt(1.0 - (2.0 * R_m / (0.85 * fc)))) if (2.0*R_m/(0.85*fc)) < 1.0 else 0.002
    
    As_x_top_req = max(rho * B_at_m * d, 0.0018 * B_at_m * h)
    As_x_bot_req = 0.0018 * B_at_m * h
    As_y_top_req = 0.0018 * L * h
    As_y_bot_req = 0.0018 * L * h
    
    x_crit = (a1 + S) - d
    B_at_crit = B1 - (B1 - B2) * (x_crit / L)
    Vu_1way = abs(P1u - (qu1 * x_crit - (qu1 - qu2) * (x_crit**2) / (2.0 * L)))
    Vc_1way = 0.53 * math.sqrt(fc) * B_at_crit * d
    As_pier_req = 0.01 * (B1 * 50.0)
    
    return {
        "Mu_neg": Mu_neg, "x_m_max": x_m_max, "B_at_m": B_at_m,
        "As_x_top_req": As_x_top_req, "As_x_bot_req": As_x_bot_req,
        "As_y_top_req": As_y_top_req, "As_y_bot_req": As_y_bot_req,
        "Vu_1way": Vu_1way, "phi_Vc_1way": 0.75 * Vc_1way, "is_1way_safe": Vu_1way <= (0.75 * Vc_1way),
        "As_pier_req": As_pier_req
    }

# ==========================================
# 3. Word 報告生成功能 (修復 Tuple 指定與定義先後)
# ==========================================

def generate_batch_report(df_results, footing_type):
    doc = Document()
    doc.add_heading('RC 基礎工程批次結構計算書 (智慧配筋版)', level=0).alignment = 1
    
    doc.add_heading('一、 批次檢核概述與總表', level=1)
    doc.add_paragraph(f"本工程個案採用【{footing_type}】進行批次自動化設計與結構強度安全檢核。")
    
    # 建立簡要總表
    sum_table = doc.add_table(rows=1, cols=4)
    sum_table.style = 'Table Grid'
    s_hdr = sum_table.rows[0].cells
    s_hdr[0].text = '編號/柱號'
    s_hdr[1].text = '最大實際地壓 (kgf/cm²)'
    s_hdr[2].text = '穿孔剪力檢核'
    s_hdr[3].text = '綜合判定'
    
    # 填入簡要總表 (【核心修正】：一律採 c[index].text 賦值，修復 Tuple 錯誤)
    for _, row in df_results.iterrows():
        c = sum_table.add_row().cells
        c[0].text = str(row['柱號/項目'])
        c[1].text = f"{row['最大實際地壓']:.2f}" if row['地壓檢核'] == "✅ 通過" else "N/A"
        c[2].text = str(row['穿孔剪力檢核'])
        c[3].text = str(row['綜合判定'])
        
    doc.add_paragraph("\n" + "="*50 + "\n")
    doc.add_heading('二、 各別柱號詳細結構檢核與配筋書', level=1)
    
    # 針對每一根柱子自動生成獨立章節與配筋表
    for _, row in df_results.iterrows():
        doc.add_heading(f"🏢 結構柱號：{row['柱號/項目']}", level=2)
        
        doc.add_paragraph(f"• 短柱 Pier 總高度：{row['短柱高度']:.2f} m (已包含高於地表 30cm)")
        doc.add_paragraph(f"• 最大計算實際地壓 q_act：{row['最大實際地壓']:.2f} kgf/cm²")
        doc.add_paragraph(f"• 綜合安全判定結果：{row['綜合判定']}")
        
        if "🟢" in row['綜合判定'] or "🔺" in row['綜合判定']:
            doc.add_paragraph(f"• 穿孔剪力檢核狀態：{row['穿孔剪力檢核']}")
            
            doc.add_heading(f"  現場施工智慧配筋建議表 ({row['柱號/項目']})", level=3)
            rebar_table = doc.add_table(rows=1, cols=3)
            rebar_table.style = 'Table Grid'
            r_hdr = rebar_table.rows[0].cells
            r_hdr[0].text = '配筋方向/部位'
            r_hdr[1].text = '計算所需面積 (cm²)'
            r_hdr[2].text = '現場施工建議配筋 (#?@? 或支數)'
            
            rebar_rows = [
                ("縱向 X向 頂層筋 (X-Top)", "0.00 (構造筋)", "無主筋需求"),
                ("縱向 X向 底層筋 (X-Bot)", f"{row['As_x_bot']:.2f}", str(row['spacing_x_bot'])),
                ("橫向 Y向 頂層筋 (Y-Top)", "0.00 (構造筋)", "無主筋需求"),
                ("橫向 Y向 底層筋 (Y-Bot)", f"{row['As_y_bot']:.2f}", str(row['spacing_y_bot'])),
                ("🏢 Pier 短柱主鋼筋", f"{row['As_pier']:.2f}", str(row['qty_pier']))
            ]
            
            # 【核心修正】：逐列寫入鋼筋對照明細
            for item in rebar_rows:
                rc = rebar_table.add_row().cells
                rc[0].text = item[0]
                rc[1].text = item[1]
                rc[2].text = item[2]
                
        doc.add_paragraph("\n" + "-"*30)
        
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def generate_word_report_with_rebar(data, rebar_table_df):
    doc = Document()
    doc.add_heading('RC 基礎工程結構計算書 (智慧配筋版)', level=0).alignment = 1
    
    doc.add_heading('一、 專案基本資料與幾何概述', level=1)
    doc.add_paragraph(f"基礎型式：{data['footing_type']}")
    doc.add_paragraph(f"短柱 Pier 總高度：{data.get('h_pier', 0.0):.2f} m (已包含高於地表 30cm)")
    
    doc.add_heading('二、 結構力學強度與地壓檢核數據', level=1)
    doc.add_paragraph(f"最大計算實際地壓 q_act：{data.get('q_act', 0.0):.2f} kgf/cm²")
    
    if "獨立基腳" in data['footing_type']:
        doc.add_paragraph(f"基腳有效深度 d：{data.get('d', 0.0):.1f} cm")
        doc.add_paragraph(f"設計穿孔剪力外力 Vu：{data.get('Vu', 0.0):,.2f} kgf")
        doc.add_paragraph(f"混凝土容許抗力 φVc：{data.get('phi_Vc', 0.0):,.2f} kgf")
        status = "合格 (Vu ≦ φVc)" if data.get('Vu', 0.0) <= data.get('phi_Vc', 0.0) else "不合格 (請增加版厚)"
        doc.add_paragraph(f"穿孔剪力判定結果：{status}")
    else:
        doc.add_paragraph("備註：梯形聯合基腳模組已內建土壤容許承載力作為求解限制。")

    doc.add_heading('三、 現場施工智慧配筋建議表', level=1)
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '配筋方向/部位'
    hdr_cells[1].text = '計算所需面積 (cm²)'
    hdr_cells[2].text = '現場施工建議配筋 (#?@? 或支數)'
    
    for _, row in rebar_table_df.iterrows():
        c = table.add_row().cells
        c[0].text = str(row['配筋方向/部位'])
        c[1].text = str(row['計算所需面積 (cm²)'])
        c[2].text = str(row['💡 現場施工建議配筋 (#?@?)'])
        
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ==========================================
# 4. Streamlit 前端 UI 介面
# ==========================================
st.set_page_config(page_title="RC 基礎設計工程 App", layout="wide")
st.title("🏗️ RC 基礎設計與智慧配筋系統")

footing_type = st.selectbox("請選擇基礎型式：", ["獨立基腳 (Isolated)", "梯形聯合基腳 (Combined)"])
st.markdown("---")

st.markdown("### 🎛️ 鋼筋規格設定 (工程現場出圖竹節筋)")
f_col1, f_col2, f_col3 = st.columns(3)
with f_col1:
    bar_x_top = st.selectbox("X縱向 頂層鋼筋 (X-Top)", ["#5", "#6", "#7", "#8"], index=1)
    bar_x_bot = st.selectbox("X縱向 底層鋼筋 (X-Bot)", ["#5", "#6", "#7", "#8"], index=1)
with f_col2:
    bar_y_top = st.selectbox("Y橫向 頂層鋼筋 (Y-Top)", ["#4", "#5", "#6", "#7"], index=1)
    bar_y_bot = st.selectbox("Y橫向 底層鋼筋 (Y-Bot)", ["#4", "#5", "#6", "#7"], index=2)
with f_col3:
    bar_pier = st.selectbox("Pier 短柱主鋼筋結構號數", ["#6", "#7", "#8"], index=2)

st.markdown("---")

if footing_type == "獨立基腳 (Isolated)":
    st.subheader("📊 獨立基腳偏心側力智慧配筋模組")
    input_mode = st.radio("資料輸入模式：", ["單案手動輸入", "Excel 批次範本導入"])
    
    if input_mode == "單案手動輸入":
        col1, col2 = st.columns(2)
        with col1:
            fc = st.selectbox("混凝土強度 f'c (kgf/cm²)", [210, 245, 280, 350], index=2, key="iso_fc")
            fy = st.selectbox("鋼筋屈服強度 fy (kgf/cm²)", [2800, 4200], index=1, key="iso_fy")
            q_all = st.number_input("土壤容許承載力 q_all (kgf/cm²)", min_value=0.5, value=2.0, key="iso_q")
            h_earth = st.number_input("基礎板上方距地表深度 (m) [規範 >= 1.2m]", min_value=0.0, value=1.5, step=0.1, format="%.2f", key="iso_he")
            c1 = st.number_input("柱長邊 c1 (cm)", min_value=10.0, value=50.0, key="iso_c1")
            c2 = st.number_input("柱短邊 c2 (cm)", min_value=10.0, value=50.0, key="iso_c2")
        with col2:
            D = st.number_input("柱靜載重 D (kgf)", min_value=0.0, value=60000.0, key="iso_D")
            L_load = st.number_input("柱活載重 L (kgf)", min_value=0.0, value=40000.0, key="iso_L")
            H = st.number_input("水平力 H (kgf)", min_value=0.0, value=2000.0, key="iso_H")
            Mx = st.number_input("X向彎矩 Mx (kgf-m)", min_value=0.0, value=5000.0, key="iso_mx")
            My = st.number_input("Y向彎矩 My (kgf-m)", min_value=0.0, value=4000.0, key="iso_my")
            B = st.number_input("基腳寬度 B (cm)", min_value=50.0, value=250.0, key="iso_B")
            L_size = st.number_input("基腳長度 L (cm)", min_value=50.0, value=250.0, key="iso_Lsz")
            h = st.number_input("基腳總厚度 h (cm)", min_value=20.0, value=60.0, key="iso_h")

        if st.button("🚀 開始獨立基腳順序檢核與智慧配筋"):
            res = check_and_calculate_isolated(c1, c2, h, fc, B, L_size, D, L_load, q_all, h_earth, H, Mx, My)
            if not res["success"]:
                st.error(f"❌ 檢核未通過：{res['reason']}")
            else:
                st.success(f"🎉 獨立基腳結構與偏心應力通過！短柱Pier總高為 {res['h_pier']:.2f} m")
                st.success(fr"✅ 偏心地壓安全 (最大地壓 {res['q_act']:.2f} ≦ 容許 {q_all:.2f})")
                
                if res["is_safe"]: 
                    st.success("✅ 二向穿孔剪力檢核也通過！結構安全。")
                else: 
                    st.error(f"❌ 結構失敗：{res['reason']}")
                
                st.markdown("### 📊 獨立基腳強度與物理量檢核數據")
                col_val1, col_val2 = st.columns(2)
                with col_val1:
                    st.write(fr"**最大計算實際地壓 $q_{{act}}$：** {res['q_act']:.2f} kgf/cm²")
                    st.write(fr"**有效深度 $d$：** {res['d']:.1f} cm")
                with col_val2:
                    st.write(fr"**設計穿孔剪力外力 $V_u$：** {res['Vu']:,.2f} kgf")
                    st.write(fr"**混凝土容許抗力 $\phi V_c$：** {res['phi_Vc']:,.2f} kgf")
                
                st.markdown("---")
                st.markdown("### 🏆 獨立基腳智慧配筋結果表 (工程現場格式)")
                
                # 【變數全面校對校正】：與 check_and_calculate_isolated 回傳的 _req 欄位完全對接
                spacing_x_bot = convert_area_to_spacing(res["As_x_bot_req"], L_size, bar_x_bot)
                spacing_y_bot = convert_area_to_spacing(res["As_y_bot_req"], B, bar_y_bot)
                qty_pier = convert_area_to_qty(res["As_pier_req"], bar_pier)
                
                reinforce_data = {
                    "配筋方向/部位": ["縱向 X向 頂層筋 (X-Top)", "縱向 X向 底層筋 (X-Bot)", "橫向 Y向 頂層筋 (Y-Top)", "橫向 Y向 底層筋 (Y-Bot)", "🏢 Pier 短柱主鋼筋"],
                    "計算所需面積 (cm²)": ["0.00 (構造筋)", f"{res['As_x_bot_req']:.2f}", "0.00 (構造筋)", f"{res['As_y_bot_req']:.2f}", f"{res['As_pier_req']:.2f}"],
                    "💡 現場施工建議配筋 (#?@?)": ["無主筋需求", spacing_x_bot, "無主筋需求", spacing_y_bot, qty_pier]
                }
                df_rebar = pd.DataFrame(reinforce_data)
                st.table(df_rebar)
                
                doc_bytes = generate_word_report_with_rebar({
                    "footing_type": footing_type, "q_act": res["q_act"], "h_pier": res["h_pier"],
                    "d": res["d"], "Vu": res["Vu"], "phi_Vc": res["phi_Vc"]
                }, df_rebar)
                st.download_button("📄 下載獨立基腳結構計算書(.docx)", data=doc_bytes, file_name="獨立基腳智慧配筋計算書.docx")

    else:
        st.markdown("### 📥 批次計算模式")
        
        sample_data = [{
            "柱號": "C1", "c1": 50, "c2": 50, "h": 60, "fc": 280, "B": 350, "L": 350, 
            "D": 60000, "Live": 40000, "q_all": 2.5, "h_earth": 1.5, "H": 1500, "Mx": 1200, "My": 1000
        }, {
            "柱號": "C2", "c1": 60, "c2": 60, "h": 70, "fc": 280, "B": 400, "L": 400, 
            "D": 80000, "Live": 50000, "q_all": 2.5, "h_earth": 1.5, "H": 2000, "Mx": 1500, "My": 1200
        }]
        df_sample = pd.DataFrame(sample_data)
        
        buffer_sample = io.BytesIO()
        with pd.ExcelWriter(buffer_sample, engine='openpyxl') as writer:
            df_sample.to_excel(writer, index=False, sheet_name='Sheet1')
        buffer_sample.seek(0)
        
        st.download_button(
            label="📥 點我下載標準 Excel 範例檔 (Sample File)",
            data=buffer_sample,
            file_name="獨立基腳批次輸入範例檔.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("---")

        uploaded_file = st.file_uploader("請上傳或將您的基礎設計 Excel 檔 (.xlsx) 拖曳至此處進行批次檢核", type=["xlsx"])
        
        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)
            
            if 'Live' not in df.columns and 'L' in df.columns: df.rename(columns={'L': 'Live'}, inplace=True)
            elif 'Live' not in df.columns and 'live' in df.columns: df.rename(columns={'live': 'Live'}, inplace=True)
                
            for col in ['H', 'Mx', 'My']:
                if col not in df.columns: df[col] = 0.0
            
            st.write("📊 **已讀取之資料預覽：**")
            st.dataframe(df)
            
            if st.button("⚡ 執行全自動批次運算"):
                results_list = []
                
                for idx, row in df.iterrows():
                    res = check_and_calculate_isolated(
                        row['c1'], row['c2'], row['h'], row['fc'], 
                        row['B'], row['L'], row['D'], row['Live'], row['q_all'], row['h_earth'],
                        row['H'], row['Mx'], row['My']
                    )
                    
                    status_item = {"柱號/項目": row['柱號'], "短柱高度": row['h_earth'] + 0.3, "最大實際地壓": res["q_act"]}
                    
                    if not res["success"]:
                        status_item.update({
                            "地壓檢核": "❌ 未通過", "穿孔剪力檢核": "未執行", "綜合判定": "❌ 失敗 (偏心地壓超限)",
                            "As_x_bot": 0.0, "spacing_x_bot": "未執行", "As_y_bot": 0.0, "spacing_y_bot": "未執行", "As_pier": 0.0, "qty_pier": "未執行"
                        })
                    else:
                        status_item.update({"地壓檢核": "✅ 通過"})
                        
                        spacing_x_bot = convert_area_to_spacing(res["As_x_bot_req"], row['L'], bar_x_bot)
                        spacing_y_bot = convert_area_to_spacing(res["As_y_bot_req"], row['B'], bar_y_bot)
                        qty_pier = convert_area_to_qty(res["As_pier_req"], bar_pier)
                        
                        status_item.update({
                            "As_x_bot": res["As_x_bot_req"], "spacing_x_bot": spacing_x_bot,
                            "As_y_bot": res["As_y_bot_req"], "spacing_y_bot": spacing_y_bot,
                            "As_pier": res["As_pier_req"], "qty_pier": qty_pier
                        })

                        if res["is_safe"]: status_item.update({"穿孔剪力檢核": "✅ 通過", "綜合判定": "🟢 完全合格"})
                        else: status_item.update({"穿孔剪力檢核": "❌ 未通過", "綜合判定": "🔺 失敗 (版厚穿剪不足)"})
                            
                    results_list.append(status_item)
                
                df_res = pd.DataFrame(results_list)
                st.markdown("### 🏆 批次計算結果總表")
                st.dataframe(df_res[["柱號/項目", "短柱高度", "最大實際地壓", "地壓檢核", "穿孔剪力檢核", "綜合判定"]])
                
                doc_bytes = generate_batch_report(df_res, footing_type)
                st.download_button(
                    label="📥 下載多章節智慧配筋批次結構計算書 (.docx)",
                    data=doc_bytes,
                    file_name="獨立基腳偏心側力多章節批次檢核報告.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

else: # 梯形聯合基腳模組
    st.subheader("📐 梯形聯合基腳智慧鋼筋排筋配置面板")
    col1, col2 = st.columns(2)
    with col1:
        fc = st.selectbox("混凝土強度 f'c (kgf/cm²)", [210, 245, 280, 350], index=2, key="comb_fc")
        fy = st.selectbox("鋼筋屈服強度 fy (kgf/cm²)", [2800, 4200], index=1, key="comb_fy")
        q_all = st.number_input("土壤容許承載力 q_all (kgf/cm²)", min_value=0.5, value=2.5, key="comb_q")
        S = st.number_input("兩柱中心距離 S (cm)", min_value=50.0, value=450.0, key="comb_s")
        a1 = st.number_input("外柱中心至左地界 a1 (cm)", min_value=10.0, value=50.0, key="comb_a1")
        L = st.number_input("限定之基腳總長度 L (cm)", min_value=100.0, value=600.0, key="comb_L")
    with col2:
        P1_D = st.number_input("外柱靜載重 P1_D (kgf)", value=50000.0, key="comb_p1d")
        P1_L = st.number_input("外柱活載重 P1_L (kgf)", value=30000.0, key="comb_p1l")
        P2_D = st.number_input("內柱靜載重 P2_D (kgf)", value=70000.0, key="comb_p2d")
        P2_L = st.number_input("內柱活載重 P2_L (kgf)", value=40000.0, key="comb_p2l")
        h = st.number_input("預估基腳總版厚 h (cm)", min_value=30.0, value=80.0, key="comb_h")

    if st.button("⚡ 執行聯合基腳全套幾何與智慧配筋運算"):
        geo = calculate_trapezoid_geometry(P1_D+P1_L, P2_D+P2_L, S, a1, q_all, L)
        if not geo["success"]:
            st.error(f"❌ 幾何失敗：{geo['msg']}")
        else:
            st.success(fr"✅ 幾何開展成功！建議尺寸：左寬 $B_1$ = {geo['B1']:.1f} cm，著寬 $B_2$ = {geo['B2']:.1f} cm")
            det = calculate_combined_details(P1_D, P1_L, P2_D, P2_L, S, a1, q_all, L, h, fc, fy, geo["B1"], geo["B2"])
            
            spacing_x_top = convert_area_to_spacing(det["As_x_top_req"], det["B_at_m"], bar_x_top)
            spacing_x_bot = convert_area_to_spacing(det["As_x_bot_req"], det["B_at_m"], bar_x_bot)
            spacing_y_top = convert_area_to_spacing(det["As_y_top_req"], L, bar_y_top)
            spacing_y_bot = convert_area_to_spacing(det["As_y_bot_req"], L, bar_y_bot)
            qty_pier = convert_area_to_qty(det["As_pier_req"], bar_pier)
            
            reinforce_data = {
                "配筋方向/部位": ["縱向 X向 頂層筋 (X-Top)", "縱向 X向 底層筋 (X-Bot)", "橫向 Y向 頂層筋 (Y-Top)", "橫向 Y向 底層筋 (Y-Bot)", "🏢 Pier 短柱主鋼筋"],
                "計算所需面積 (cm²)": [f"{det['As_x_top_req']:.2f}", f"{det['As_x_bot_req']:.2f}", f"{det['As_y_top_req']:.2f}", f"{det['As_y_bot_req']:.2f}", f"{det['As_pier_req']:.2f}"],
                "💡 現場施工建議配筋 (#?@?)": [spacing_x_top, spacing_x_bot, spacing_y_top, spacing_y_bot, qty_pier]
            }
            df_rebar = pd.DataFrame(reinforce_data)
            st.table(df_rebar)
            
            doc_bytes = generate_word_report_with_rebar({"footing_type": footing_type, "q_act": q_all, "h_pier": 1.8}, df_rebar)
            st.download_button("📄 下載聯合基腳結構計算書(.docx)", data=doc_bytes, file_name="聯合基腳智慧配筋計算書.docx")