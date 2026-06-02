import streamlit as st
import pandas as pd
import math
from datetime import datetime

# ===============================
# 1. 基础配置与数据加载
# ===============================
st.set_page_config(page_title="SADE 采购决策支持系统", layout="wide")

@st.cache_data
def load_all_data():
    try:
        # 1. 加载合同表 (Excel)
        df_contracts = pd.read_excel("contracts_b.xlsx")
        for col in ["DE", "PN", "Price", "MOQ 12ml"]:
            df_contracts[col] = pd.to_numeric(df_contracts[col], errors='coerce')
        
        # 2. 加载运费表 (修改这里：从 read_csv 改为 read_excel)
        # 请确保文件名正确，例如 "Transport PE.xlsx"
        df_transport = pd.read_excel("Transport PE.xlsx") 
        
        # 3. 数据清洗 (Excel 读取后列名和字符串前后常有空格)
        df_transport.columns = df_transport.columns.str.strip()
        df_transport['Dpt'] = df_transport['Dpt'].astype(str).str.strip()
        df_transport['DEPARTEMENTS'] = df_transport['DEPARTEMENTS'].astype(str).str.strip()
        df_transport['Supplier'] = df_transport['Supplier'].astype(str).str.strip()
        
        # 4. 生成省份列表 (用于下拉菜单显示)
        dept_info = df_transport[['Dpt', 'DEPARTEMENTS']].drop_duplicates().sort_values('Dpt')
        dept_list = [f"{row['Dpt']} - {row['DEPARTEMENTS']}" for _, row in dept_info.iterrows()]
        
        return df_contracts, df_transport, dept_list
    except Exception as e:
        st.error(f"加载文件失败，请检查文件是否存在且格式正确: {e}")
        return None, None, []

contracts, transport_db, dept_options_list = load_all_data()

# ===============================
# 2. 你的核心业务规则 (Rules)
# ===============================
def rule_distributor_purchase(quantity, package, DE):
    return (package == "couronne" or DE < 125 or (DE < 200 and quantity < 960) or (225 <= DE <= 355 and quantity < 360))

def rule_contract_purchase(quantity, package, DE):
    return ((package == "barre" and 125 <= DE <= 200 and 960 <= quantity <2000)
            or (package == "barre" and 225 <= DE <= 355 and 360 <= quantity < 1000))

def rule_factory_purchase(quantity, package, DE):
    return ((package == "barre" and 225 <= DE <= 355 and 1000 <= quantity)
            or (package == "barre" and 125 <= DE <= 200 and 2000 <= quantity)
            or package.lower() == "touret" or (package == "barre" and 355 < DE))

def rule_distributor_purchase_dipipe(quantity, DE):
    return (DE < 80)

def rule_contract_purchase_dipipe(quantity, DE):
    # 铸铁管(Fonte)的合同规则
    conditions = [
        (DE >= 300 and quantity <= 264), (DE >= 250 and quantity <= 396),
        (DE >= 200 and quantity <= 440), (DE >= 150 and quantity <= 594),
        (DE >= 125 and quantity <= 770), (DE >= 100 and quantity <= 891),
        (DE >= 80 and quantity <= 968)
    ]
    return any(conditions)

def rule_factory_purchase_dipipe(quantity, DE):
    return not rule_contract_purchase_dipipe(quantity, DE) and DE >= 80
def generate_email_template(supplier, material, quantity, de, pn, package):
    subject = f"Demande de prix - {material} - DE{de} PN{pn}"
    body = f"Bonjour,\n\nDans le cadre d'un nouveau projet, nous souhaiterions obtenir votre meilleure offre pour :\n- Produit : {material}\n- DE : {de} / PN : {pn}\n- Quantité : {quantity} ml\n- Conditionnement : {package}\n\nCordialement,"
    return subject, body

# ===============================
# 3. 价格计算逻辑 (MOQ + Transport)
# ===============================
def calculate_all_totals(material, de, pn, quantity, package, dept_code, today):
    pkg_str = str(package).lower() if package else ""
    mask = (
        (contracts["Material"] == material) &
        (contracts["Valid_Until"] >= today) &
        (contracts["DE"] == float(de)) &
        (contracts["PN"] == float(pn)) &
        (contracts["Package"].astype(str).str.lower() == pkg_str)
    )
    valid_matches = contracts[mask].copy()
    
    # 关键点：如果 MOQ 12ml 为空，说明不符合合同价执行条件
    valid_matches = valid_matches[valid_matches["MOQ 12ml"].notna() & (valid_matches["MOQ 12ml"] > 0)]
    if valid_matches.empty:
        return None

    # 计算车数
    valid_matches["Nb_Camions"] = valid_matches["MOQ 12ml"].apply(lambda x: math.ceil(quantity / x))

    # 获取对应省份和供应商的运费
    def get_fee(supplier):
        fee_m = (transport_db["Supplier"].str.contains(supplier, case=False, na=False)) & (transport_db["Dpt"] == str(dept_code))
        res = transport_db[fee_m]["Transport"]
        return res.iloc[0] if not res.empty else 0

    valid_matches["Transport_Unit"] = valid_matches["Supplier"].apply(get_fee)
    
    # 金额计算
    valid_matches["Material_Total"] = valid_matches["Price"] * quantity
    valid_matches["Total_Transport"] = valid_matches["Nb_Camions"] * valid_matches["Transport_Unit"]
    valid_matches["Grand_Total"] = valid_matches["Material_Total"] + valid_matches["Total_Transport"]

    display_df = valid_matches[["Supplier", "Price", "Nb_Camions", "Transport_Unit", "Total_Transport", "Grand_Total"]].copy()
    display_df.columns = ["Fournisseur", "Unit (€/ml)", "Camions", "Frais/Cam", "Total Trans", "TOTAL HT"]
    
    for col in ["Unit (€/ml)", "Frais/Cam", "Total Trans", "TOTAL HT"]:
        display_df[col] = display_df[col].map("{:,.2f} €".format)
    return display_df.sort_values("TOTAL HT")

# ===============================
# 4. Streamlit UI
# ===============================
st.title("🛡️ SADE Purchasing Decision Support")

if contracts is not None:
    with st.form("purchase_form"):
        col1, col2 = st.columns(2)
        with col1:
            material_choice = st.selectbox("Matériau:", options=[""] + sorted(contracts["Material"].dropna().unique().tolist()))
            package_choice = st.selectbox("Conditionnement:", options=["", "barre", "couronne", "touret"])
            qty_input = st.number_input("Quantité (ml):", min_value=0, step=1)
        with col2:
            de_choice = st.selectbox("DE (Diamètre):", options=[""] + sorted([int(x) for x in contracts["DE"].dropna().unique()]))
            pn_choice = st.selectbox("PN (Pression):", options=[""] + sorted([float(x) for x in contracts["PN"].dropna().unique()]))
            dept_full = st.selectbox("Département de livraison:", options=[""] + dept_options_list)
        
        submit_btn = st.form_submit_button("Run Decision", type="primary")

    if submit_btn and material_choice and package_choice and de_choice and dept_full:
        dept_code = dept_full.split(" - ")[0]
        today = datetime.today()
        
        # --- 决策逻辑判定 ---
        is_fonte = "fonte" in material_choice.lower()
        price_table = None
        decision_msg = ""
        
        # 1. 判定是否满足合同规则
        if is_fonte:
            if rule_contract_purchase_dipipe(qty_input, de_choice):
                price_table = calculate_all_totals(material_choice, de_choice, pn_choice, qty_input, package_choice, dept_code, today)
                decision_msg = "✅ Decision: Application tarif contractuel Electrosteel"
            elif rule_factory_purchase_dipipe(qty_input, de_choice):
                decision_msg = "✅ Decision: Consultation Electrosteel (Usine)"
            else:
                decision_msg = "🛒 Decision: Consultation Négoce"
        else:
            if rule_contract_purchase(qty_input, package_choice, de_choice):
                price_table = calculate_all_totals(material_choice, de_choice, pn_choice, qty_input, package_choice, dept_code, today)
                decision_msg = "✅ Decision: Application tarif contractuelle"
            elif rule_factory_purchase(qty_input, package_choice, de_choice):
                price_table = calculate_all_totals(material_choice, de_choice, pn_choice, qty_input, package_choice, dept_code, today)
                decision_msg = "✅ Decision: Consultation Fabricant (Elydan, Centraltubi) pour avoir meilleure prix que les conditions contractuels "
            else:
                decision_msg = "🛒 Decision: Consultation Négoce"

        # --- 显示结果 ---
        st.divider()
        st.subheader(decision_msg)
        
        if price_table is not None:
            st.write("### 💰 Comparatif des prix (Transport inclus)")
            st.table(price_table)
        else:
            if "Application" in decision_msg:
                st.warning("⚠️ Contrat trouvé mais MOQ 12ml non renseignée dans le fichier Excel.")
            
            # 邮件草稿
            if "Consultation" in decision_msg:
                st.info("📧 **Brouillon d'Email de consultation**")
                subject, body = generate_email_template(material_choice, qty_input, de_choice, pn_choice, package_choice)

                st.text_area("Copier :", value=body, height=120)
