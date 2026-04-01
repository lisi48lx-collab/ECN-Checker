import streamlit as st
import pandas as pd
import io

# ================= 网页全局设置 =================
st.set_page_config(
    page_title="ECN 深度审计引擎",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 隐藏默认的 Streamlit 菜单，让它看起来更像一个纯净的独立软件
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ================= 页面标题 =================
st.title("⚡ ECN 深度审计引擎")
st.markdown("**Ultimate Precision | Web 在线版 (支持双擎解析 & 智能容错)**")
st.markdown("---")

# ================= 上传区 =================
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. 导入设计变更单 (ECN)")
    ecn_file = st.file_uploader("支持格式：.xlsx, .xlsm, .csv", type=["xlsx", "xlsm", "csv"], key="ecn")

with col2:
    st.subheader("2. 导入 U9 BOM 报表")
    u9_file = st.file_uploader("支持格式：.xlsx, .xls, .csv", type=["xlsx", "xls", "csv"], key="u9")

st.markdown("---")

# ================= 核心清洗与比对算法 =================
def smart_match(val1, val2):
    """智能视觉容错算法：抹平数字前导零 (05=5)，忽略大小写 (v1=V1)"""
    if val1 == val2: return True
    if val1 and val2:
        v1, v2 = str(val1).strip(), str(val2).strip()
        if v1.upper() == v2.upper(): return True
        if v1.isdigit() and v2.isdigit() and int(v1) == int(v2): return True
    return False

def highlight_status(val):
    """网页端表格红绿灯渲染"""
    if val == "✅ 完美匹配":
        return 'color: #2ecc71; font-weight: bold'
    elif val == "待定" or val == "":
        return ''
    else:
        return 'color: #ff6b6b; font-weight: bold'

# ================= 执行按钮 =================
if st.button("🚀 启动高精度比对", type="primary", use_container_width=True):
    if not ecn_file or not u9_file:
        st.warning("⚠️ 请先在上方上传 ECN 和 U9 BOM 文件！")
    else:
        with st.spinner("引擎正在进行多维交叉核对，请稍候..."):
            try:
                # ================= 1. 智能读取 ECN (自动识别 CSV / Excel) =================
                cols_idx = [4, 5, 7, 8, 15] 
                col_names = ["子阶代码", "子阶名称", "规格", "变更前版本", "母阶代码"]

                if ecn_file.name.lower().endswith('.csv'):
                    try:
                        df_ecn = pd.read_csv(ecn_file, encoding='utf-8', skiprows=7, header=None, usecols=cols_idx, names=col_names, dtype=str)
                    except UnicodeDecodeError:
                        ecn_file.seek(0) # 重置文件指针
                        df_ecn = pd.read_csv(ecn_file, encoding='gbk', skiprows=7, header=None, usecols=cols_idx, names=col_names, dtype=str)
                else:
                    xls = pd.ExcelFile(ecn_file, engine="openpyxl")
                    target_sheet = xls.sheet_names[0]
                    for sn in xls.sheet_names:
                        if "统计" not in sn and "Option" not in sn and "option" not in sn.lower():
                            target_sheet = sn
                            break
                    df_ecn = pd.read_excel(xls, sheet_name=target_sheet, skiprows=7, header=None, usecols=cols_idx, names=col_names, dtype=str)

                # ================= 2. 智能读取 U9 BOM (自动识别 CSV / Excel) =================
                if u9_file.name.lower().endswith('.csv'):
                    try:
                        df_u9_raw = pd.read_csv(u9_file, encoding='utf-8', header=None, dtype=str)
                        u9_file.seek(0)
                    except UnicodeDecodeError:
                        u9_file.seek(0)
                        df_u9_raw = pd.read_csv(u9_file, encoding='gbk', header=None, dtype=str)
                        u9_file.seek(0)
                    
                    h_idx = next((i for i, row in df_u9_raw.iterrows() if '料品编码' in row.values or '子件编码' in row.values), 0)
                    try:
                        df_u9 = pd.read_csv(u9_file, encoding='utf-8', skiprows=h_idx, dtype=str)
                    except UnicodeDecodeError:
                        u9_file.seek(0)
                        df_u9 = pd.read_csv(u9_file, encoding='gbk', skiprows=h_idx, dtype=str)
                else:
                    df_u9_raw = pd.read_excel(u9_file, header=None, dtype=str)
                    h_idx = next((i for i, row in df_u9_raw.iterrows() if '料品编码' in row.values or '子件编码' in row.values), 0)
                    u9_file.seek(0)
                    df_u9 = pd.read_excel(u9_file, skiprows=h_idx, dtype=str)

                # ================= 3. 核心比对清洗引擎 =================
                clean = lambda v: str(v).strip().replace('.0','') if pd.notna(v) and str(v).strip() != "" and str(v).strip() != "nan" else None
                
                u9_cols = [str(c).strip() for c in df_u9.columns]
                def find_col(possible_names):
                    for n in possible_names:
                        if n in u9_cols: return n
                    return None

                u_col_no = find_col(['料品编码', '子件编码', '物料编码', '品号'])
                u_col_name = find_col(['料品名称', '子件名称', '品名'])
                u_col_spec = find_col(['规格型号', '规格', '料品规格'])
                # 💡 补充了你昨天发现的 "料品版本"
                u_col_ver = find_col(['料品版本', '版本', '版次', '子件版本', '版本号'])

                if not u_col_no:
                    st.error("❌ U9表格中未检测到【料品编码】列，请检查导出格式！")
                    st.stop()

                df_u9['IDX_NO'] = df_u9[u_col_no].apply(clean)
                df_u9['IDX_NAME'] = df_u9[u_col_name].apply(clean) if u_col_name else None
                df_u9['IDX_SPEC'] = df_u9[u_col_spec].apply(clean) if u_col_spec else None
                df_u9['IDX_VER'] = df_u9[u_col_ver].apply(clean) if u_col_ver else None

                results = []
                ecn_records = df_ecn.to_dict('records')

                for e in ecn_records:
                    e_pno = clean(e.get("子阶代码"))
                    e_spec = clean(e.get("规格"))
                    e_name = clean(e.get("子阶名称"))
                    e_old_ver = clean(e.get("变更前版本"))
                    e_parent = clean(e.get("母阶代码")) 
                    
                    if not e_pno and not e_spec and not e_name: continue

                    status, detail = "待定", ""
                    search_space = df_u9
                    
                    if e_parent and '母件编码' in df_u9.columns:
                        parent_filter = df_u9[df_u9['母件编码'].apply(clean) == e_parent]
                        if not parent_filter.empty: search_space = parent_filter

                    match_no = search_space[search_space['IDX_NO'] == e_pno] if e_pno else pd.DataFrame()

                    if not match_no.empty:
                        u = match_no.iloc[0]
                        errs = []
                        
                        # 💡 全面应用 smart_match 智能容错对比
                        if e_name:
                            if u.get('IDX_NAME') is None: errs.append("[U9无品名列]")
                            elif not smart_match(e_name, u.get('IDX_NAME')): errs.append(f"[品名] U9为:{u.get('IDX_NAME')}")
                        
                        if e_spec:
                            if u.get('IDX_SPEC') is None: errs.append("[U9无规格列]")
                            elif not smart_match(e_spec, u.get('IDX_SPEC')): errs.append(f"[规格] U9为:{u.get('IDX_SPEC')}")
                        
                        if e_old_ver:
                            if u.get('IDX_VER') is None: errs.append("[U9无版本列]")
                            elif not smart_match(e_old_ver, u.get('IDX_VER')): errs.append(f"[版本] U9实为:{u.get('IDX_VER')}")

                        status = "✅ 完美匹配" if not errs else "⚠️ 信息不符"
                        detail = " | ".join(errs)
                    else:
                        # 安全结构防断行
                        if e_spec or e_name:
                            query = search_space
                            if e_spec and 'IDX_SPEC' in query.columns: 
                                query = query[query['IDX_SPEC'] == e_spec]
                            if e_name and 'IDX_NAME' in query.columns: 
                                query = query[query['IDX_NAME'] == e_name]
                            
                            if not query.empty:
                                suggested_no = query.iloc[0].get(u_col_no)
                                status = "🚫 品号疑似错误"
                                detail = f"品号不存在，根据规格反查应为: {suggested_no}"
                            else:
                                status = "🚨 彻底未命中"
                                detail = "系统中无此品号，且根据规格无法推导。"
                        else:
                            status = "🚨 彻底未命中"
                            detail = "品号错误，且没有填写规格可供反查。"

                    results.append({
                        "填报母阶": e_parent if e_parent else "-",
                        "填报子阶": e_pno,
                        "填报名称": e_name,
                        "填报规格": e_spec,
                        "填报版本": e_old_ver,
                        "最终判定": status,
                        "系统反馈": detail
                    })

                # ================= 4. 渲染展示大屏 =================
                st.success("✅ 核对完毕！请查看下方审计报告。")
                df_results = pd.DataFrame(results)
                
                # 应用红绿灯渲染 (兼容 pandas 的旧版 applymap 和新版 map)
                try:
                    styled_df = df_results.style.map(highlight_status, subset=['最终判定'])
                except AttributeError:
                    styled_df = df_results.style.applymap(highlight_status, subset=['最终判定'])
                    
                st.dataframe(styled_df, use_container_width=True, height=450)

                # ================= 5. 提供下载按钮 =================
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_results.to_excel(writer, index=False, sheet_name='审计报告')
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 下载 Excel 完整审计报告",
                    data=excel_data,
                    file_name="ECN高精度审计报告_网页版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            except Exception as ex:
                st.error(f"❌ 运行发生底层错误，请检查表格格式！详细报错信息：\n{ex}")