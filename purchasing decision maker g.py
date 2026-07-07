import streamlit as st
import pandas as pd
import math
import urllib.parse
from datetime import datetime

# ===============================
# 1. 基础配置与数据加载
# ===============================
st.set_page_config(page_title="SADE 采购决策支持系统", layout="wide")

def _parse_ml_par_unit(x):
    """把 '6m' / '12m' / NaN 这样的字符串解析成数字长度（米）"""
    if pd.isna(x):
        return None
    s = str(x).lower().replace("m", "").strip()
    try:
        val = float(s)
        return val if val > 0 else None
    except (ValueError, TypeError):
        return None


@st.cache_data
def load_all_data():
    try:
        # 1. 加载合同表 (Excel)
        df_contracts = pd.read_excel("contracts_b.xlsx")
        for col in ["DE", "PN", "Price", "MOQ 12ml"]:
            df_contracts[col] = pd.to_numeric(df_contracts[col], errors='coerce')

        # 2. 加载运费表
        df_transport = pd.read_excel("Transport PE.xlsx")
        df_transport.columns = df_transport.columns.str.strip()
        df_transport['Dpt'] = df_transport['Dpt'].astype(str).str.strip()
        df_transport['DEPARTEMENTS'] = df_transport['DEPARTEMENTS'].astype(str).str.strip()
        df_transport['Supplier'] = df_transport['Supplier'].astype(str).str.strip()

        # 3. 生成省份列表 (用于下拉菜单显示)
        dept_info = df_transport[['Dpt', 'DEPARTEMENTS']].drop_duplicates().sort_values('Dpt')
        dept_list = [f"{row['Dpt']} - {row['DEPARTEMENTS']}" for _, row in dept_info.iterrows()]

        # 4. 加载 négoce 价格表 (contracts_negoce.xlsx)
        df_negoce = pd.read_excel("contracts_negoce.xlsx")
        # 统一去除列名前后空格，并做大小写不敏感的列名匹配，避免因为
        # "Franco"/"FRANCO"/"Franco " 这类差异导致 KeyError
        df_negoce.columns = df_negoce.columns.str.strip()

        expected_cols = ["Package", "DE", "SDR", "PN", "Price", "Supplier",
                          "Material", "Valid_Until", "ml par unit", "Franco"]
        col_lookup = {c.lower().strip(): c for c in df_negoce.columns}
        rename_map = {}
        missing_cols = []
        for expected in expected_cols:
            key = expected.lower().strip()
            if key in col_lookup:
                if col_lookup[key] != expected:
                    rename_map[col_lookup[key]] = expected
            else:
                missing_cols.append(expected)

        if rename_map:
            df_negoce = df_negoce.rename(columns=rename_map)

        if missing_cols:
            st.warning(
                f"⚠️ Colonnes manquantes dans contracts_negoce.xlsx : {missing_cols}. "
                f"Colonnes trouvées : {list(df_negoce.columns)}"
            )
            for col in missing_cols:
                df_negoce[col] = pd.NA

        df_negoce['Package'] = df_negoce['Package'].astype(str).str.strip()
        df_negoce['Supplier'] = df_negoce['Supplier'].astype(str).str.strip()
        df_negoce['Material'] = df_negoce['Material'].astype(str).str.strip()
        df_negoce['DE'] = pd.to_numeric(df_negoce['DE'], errors='coerce')
        df_negoce['PN'] = pd.to_numeric(df_negoce['PN'], errors='coerce')
        df_negoce['Price'] = pd.to_numeric(df_negoce['Price'], errors='coerce')
        df_negoce['Franco'] = pd.to_numeric(df_negoce['Franco'], errors='coerce')
        df_negoce['Valid_Until'] = pd.to_datetime(df_negoce['Valid_Until'], errors='coerce')
        df_negoce['ml_par_unit_val'] = df_negoce['ml par unit'].apply(_parse_ml_par_unit)

        return df_contracts, df_transport, dept_list, df_negoce
    except Exception as e:
        st.error(f"加载文件失败，请检查文件是否存在且格式正确: {e}")
        return None, None, [], None


contracts, transport_db, dept_options_list, negoce_db = load_all_data()

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

def generate_email_template(material, quantity, de, pn, package, dept):
    subject = f"Demande de prix - {material} - DE{de} PN{pn} - {dept}"
    body = f"Bonjour,\n\nDans le cadre d'un nouveau projet, nous souhaiterions obtenir votre meilleure offre pour :\n- Produit : {material}\n- DE : {de} / PN : {pn}\n- Quantité : {quantity} ml\n- Conditionnement : {package}\n- Departement : {dept}\n\n Merci par avance.\nCordialement,"
    return subject, body

# ===============================
# 3. 价格计算逻辑 (MOQ + Transport / Contrat)
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

    display_df = display_df.sort_values("TOTAL HT")
    for col in ["Unit (€/ml)", "Frais/Cam", "Total Trans", "TOTAL HT"]:
        display_df[col] = display_df[col].map("{:,.2f} €".format)
    return display_df

# ===============================
# 3bis. 价格计算逻辑 (Négoce + Franco)
# ===============================
def calculate_negoce_totals(material, de, pn, quantity, package, today):
    """
    在 contracts_negoce.xlsx 中按 DE + PN + Package + Material 查找négoce价格，
    换算成本次数量所需件数与总价，并对比 Franco 门槛。
    找不到匹配则返回 None（此时上层逻辑会回退到邮件询价草稿）。
    """
    if negoce_db is None:
        return None

    pkg_str = str(package).lower() if package else ""
    mat_str = str(material).lower() if material else ""

    mask = (
        (negoce_db["DE"] == float(de)) &
        (negoce_db["PN"] == float(pn)) &
        (negoce_db["Package"].str.lower() == pkg_str) &
        (negoce_db["Material"].str.lower() == mat_str) &
        (negoce_db["Valid_Until"] >= today)
    )
    matches = negoce_db[mask].copy()
    if matches.empty:
        return None

    def _compute(row):
        ml_val = row["ml_par_unit_val"]
        if ml_val and ml_val > 0:
            nb_unit = math.ceil(quantity / ml_val)
        else:
            # Pas de longueur d'unité renseignée -> on considère le prix comme un prix au ml
            nb_unit = quantity
        total = nb_unit * row["Price"]* quantity
        return pd.Series({"Nb_Unites": nb_unit, "Total_HT": total})

    calc = matches.apply(_compute, axis=1)
    matches = pd.concat([matches, calc], axis=1)

    def _franco_status(row):
        franco = row["Franco"]
        total = row["Total_HT"]
        if pd.isna(franco):
            return "—"
        if total >= franco:
            return "✅ Franco atteint"
        return f"⚠️ Reste {franco - total:,.2f} € pour atteindre le Franco ({franco:,.0f} €)"

    matches["Franco_Status"] = matches.apply(_franco_status, axis=1)

    # 排序（在格式化为字符串之前，保证数值排序正确）
    matches = matches.sort_values("Total_HT")

    display_df = matches[["Supplier", "Price", "ml par unit", "Nb_Unites", "Total_HT", "Franco_Status"]].copy()
    display_df.columns = ["Fournisseur", "Prix/Unité (€)", "Longueur/Unité", "Nb Unités", "TOTAL HT (€)", "Statut Franco"]
    display_df["Nb Unités"] = display_df["Nb Unités"].astype(int)
    display_df["Prix/Unité (€)"] = display_df["Prix/Unité (€)"].map("{:,.2f} €".format)
    display_df["TOTAL HT (€)"] = display_df["TOTAL HT (€)"].map("{:,.2f} €".format)
    
    return display_df.reset_index(drop=True)

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
        negoce_table = None
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

        # 2. 如果决策落在 "Consultation Négoce"，先尝试用 contracts_negoce.xlsx 计算价格
        if "Négoce" in decision_msg:
            negoce_table = calculate_negoce_totals(material_choice, de_choice, pn_choice, qty_input, package_choice, today)

        # --- 显示结果 ---
        st.divider()
        st.subheader(decision_msg)

        if price_table is not None:
            st.write("### 💰 Comparatif des prix (Transport inclus)")
            st.table(price_table)

        elif negoce_table is not None:
            st.write("### 💰 Comparatif des prix Négoce (avec condition Franco)")
            st.table(negoce_table)

        else:
            if "Application" in decision_msg:
                st.warning("⚠️ Contrat trouvé mais MOQ 12ml non renseignée dans le fichier Excel.")

            # 邮件草稿 -> uniquement si aucun prix négoce n'a été trouvé
            if "Consultation" in decision_msg:
                st.info("📧 **Brouillon d'Email de consultation**")
                subject, body = generate_email_template(material_choice, qty_input, de_choice, pn_choice, package_choice, dept_full)

                st.text_area("Copier :", value=body, height=350)

                mailto_link = (
                    f"mailto:?subject={urllib.parse.quote(subject)}"
                    f"&body={urllib.parse.quote(body)}"
                )

                st.markdown(
                    f"""
                    <a href="{mailto_link}" target="_blank">
                        <button style="
                            background-color: #0072C6;
                            color: white;
                            padding: 8px 16px;
                            border: none;
                            border-radius: 4px;
                            cursor: pointer;
                            font-size: 14px;
                        ">
                        📨 Ouvrir dans Outlook
                        </button>
                    </a>
                    """,
                    unsafe_allow_html=True,
                )
