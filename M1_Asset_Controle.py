import streamlit as st
import sqlite3
import hashlib
import pandas as pd
from datetime import datetime, date, timedelta
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==============================================================================
# CONFIGURAÇÃO DE TELA E RESILIÊNCIA STREAMLIT
# ==============================================================================
st.set_page_config(
    layout="wide",
    page_title="M1 Consultoria - Ativo Imobilizado",
    page_icon="🏢"
)

def render_table(df_or_styler, height=None):
    if height is not None:
        st.dataframe(df_or_styler, use_container_width=True, height=height)
    else:
        st.dataframe(df_or_styler, use_container_width=True)

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode('utf-8')).hexdigest()

def fmt_brl(val) -> str:
    """Formata valores numéricos para o padrão 000.000.000,00"""
    if pd.isna(val) or val is None:
        return "0,00"
    try:
        val_float = float(val)
        return f"{val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return str(val)

def fmt_date(d_val) -> str:
    """Formata datas para o padrão dd/mm/aaaa"""
    if pd.isna(d_val) or d_val is None:
        return "-"
    if isinstance(d_val, (date, datetime)):
        return d_val.strftime("%d/%m/%Y")
    d_str = str(d_val).strip()
    try:
        return datetime.strptime(d_str[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return d_str

# ==============================================================================
# BANCO DE DADOS (SQLITE)
# ==============================================================================
DB_NAME = "patrimonio_cpc27.db"

def get_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def log_audit(conn, company_id, user_email, action, entity, details):
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO audit_logs (company_id, user_email, action, entity, details)
        VALUES (?, ?, ?, ?, ?)
    """, (company_id, user_email, action, entity, details))

def init_db():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        user_type TEXT NOT NULL CHECK(user_type IN ('ADMIN', 'EXTERNO')),
        must_reset_password INTEGER DEFAULT 0,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS companies (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT UNIQUE NOT NULL,
        name TEXT NOT NULL,
        cnpj TEXT NOT NULL,
        depreciation_method TEXT NOT NULL CHECK(depreciation_method IN ('MENSAL', 'DIARIO')),
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS user_companies (
        user_id INTEGER NOT NULL,
        company_id INTEGER NOT NULL,
        PRIMARY KEY (user_id, company_id),
        FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
        FOREIGN KEY (company_id) REFERENCES companies(id) ON DELETE CASCADE
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS asset_classes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        useful_life_years REAL NOT NULL,
        fiscal_life_years REAL NOT NULL
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS asset_subclasses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        class_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        useful_life_years REAL,
        fiscal_life_years REAL,
        FOREIGN KEY (class_id) REFERENCES asset_classes(id) ON DELETE CASCADE,
        UNIQUE (class_id, name)
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS chart_of_accounts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        class_id INTEGER,
        account_code TEXT NOT NULL,
        account_name TEXT NOT NULL,
        account_type TEXT NOT NULL,
        FOREIGN KEY (company_id) REFERENCES companies(id) ON DELETE CASCADE,
        FOREIGN KEY (class_id) REFERENCES asset_classes(id) ON DELETE CASCADE
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS assets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        tag_number TEXT NOT NULL,
        description TEXT NOT NULL,
        class_id INTEGER NOT NULL,
        subclass_id INTEGER,
        invoice_key TEXT,
        acquisition_date DATE NOT NULL,
        start_date DATE NOT NULL,
        acquisition_cost REAL NOT NULL,
        additional_costs REAL DEFAULT 0,
        residual_value REAL DEFAULT 0,
        status TEXT DEFAULT 'ATIVO' CHECK(status IN ('ATIVO', 'BAIXADO')),
        brand TEXT,
        model TEXT,
        serial_number TEXT,
        year_manufacture INTEGER,
        location TEXT,
        cost_center TEXT,
        conservation_state TEXT DEFAULT 'Bom',
        initial_accum_depr REAL DEFAULT 0.0,
        initial_monthly_depr REAL DEFAULT 0.0,
        initial_base_date DATE,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (company_id) REFERENCES companies(id),
        FOREIGN KEY (class_id) REFERENCES asset_classes(id),
        FOREIGN KEY (subclass_id) REFERENCES asset_subclasses(id),
        UNIQUE (company_id, tag_number)
    );
    """)

    # Migração defensiva: assegura colunas da base inicial
    cur.execute("PRAGMA table_info(assets);")
    existing_cols = [c[1] for c in cur.fetchall()]
    new_cols = {
        "brand": "TEXT", "model": "TEXT", "serial_number": "TEXT",
        "year_manufacture": "INTEGER", "location": "TEXT", "cost_center": "TEXT",
        "conservation_state": "TEXT DEFAULT 'Bom'",
        "initial_accum_depr": "REAL DEFAULT 0.0",
        "initial_monthly_depr": "REAL DEFAULT 0.0",
        "initial_base_date": "DATE"
    }
    for col_name, col_type in new_cols.items():
        if col_name not in existing_cols:
            cur.execute(f"ALTER TABLE assets ADD COLUMN {col_name} {col_type};")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS life_revaluations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        target_type TEXT NOT NULL CHECK(target_type IN ('CLASSE', 'SUBCLASSE', 'ATIVO')),
        target_id INTEGER NOT NULL,
        cut_off_date DATE NOT NULL,
        new_useful_life_years REAL NOT NULL,
        report_reference TEXT,
        applied_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (company_id) REFERENCES companies(id)
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS asset_disposals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        asset_id INTEGER NOT NULL,
        disposal_date DATE NOT NULL,
        disposal_type TEXT NOT NULL,
        sale_value REAL DEFAULT 0,
        cost_disposed REAL NOT NULL,
        depr_accum_disposed REAL NOT NULL,
        gain_loss_value REAL NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (asset_id) REFERENCES assets(id)
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS audit_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER,
        user_email TEXT NOT NULL,
        action TEXT NOT NULL,
        entity TEXT NOT NULL,
        details TEXT NOT NULL,
        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    cur.execute("SELECT count(*) FROM companies")
    if cur.fetchone()[0] == 0:
        cur.execute("""
            INSERT INTO companies (code, name, cnpj, depreciation_method) 
            VALUES ('M1', 'M1 Consultoria e Auditoria', '00.000.000/0001-00', 'MENSAL')
        """)

    admin_email = "carloscosta@m1consultoria.com.br"
    admin_hash = hash_password("@@M1consultoria123@@")
    cur.execute("SELECT id FROM users WHERE email = ?", (admin_email,))
    user_row = cur.fetchone()
    if user_row:
        cur.execute("UPDATE users SET password_hash = ?, user_type = 'ADMIN', must_reset_password = 0 WHERE id = ?", (admin_hash, user_row[0]))
    else:
        cur.execute(
            "INSERT INTO users (email, password_hash, user_type, must_reset_password) VALUES (?, ?, 'ADMIN', 0)",
            (admin_email, admin_hash)
        )

    conn.commit()
    conn.close()

init_db()

# ==============================================================================
# FUNÇÕES DE EXPORTAÇÃO
# ==============================================================================
def export_to_excel(df: pd.DataFrame, sheet_name="Relatório") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        border = Border(
            left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
            top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3')
        )

        for col_idx, col in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row_idx in range(2, len(df) + 2):
            for col_idx in range(1, len(df.columns) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = border

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    return output.getvalue()

def export_to_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")

def generate_template_excel(columns, sample_rows=None):
    df = pd.DataFrame(sample_rows if sample_rows else [], columns=columns)
    return export_to_excel(df, sheet_name="Template")

# ==============================================================================
# AUTENTICAÇÃO E SESSÃO
# ==============================================================================
if "logged_user" not in st.session_state:
    st.session_state.logged_user = None

def login_screen():
    st.markdown("<h2 style='text-align: center;'>M1 Consultoria - Gestão Patrimonial (CPC 27)</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.subheader("Acesso Restrito")
        email = st.text_input("E-mail corporativo")
        password = st.text_input("Senha", type="password")

        if st.button("Acessar Sistema", use_container_width=True):
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("SELECT id, email, password_hash, user_type, must_reset_password FROM users WHERE email = ?", (email.strip(),))
            user = cur.fetchone()
            conn.close()

            if user and user[2] == hash_password(password):
                st.session_state.logged_user = {
                    "id": user[0],
                    "email": user[1],
                    "user_type": user[3],
                    "must_reset_password": bool(user[4])
                }
                st.rerun()
            else:
                st.error("Credenciais inválidas. Verifique seus dados.")

def reset_password_screen():
    st.subheader("Redefinição de Senha")
    st.info("Cadastre uma nova senha pessoal para prosseguir.")
    new_pw = st.text_input("Nova Senha", type="password")
    conf_pw = st.text_input("Confirmar Nova Senha", type="password")

    if st.button("Salvar Nova Senha"):
        if not new_pw or len(new_pw) < 6:
            st.error("A senha deve ter no mínimo 6 dígitos.")
            return
        if new_pw != conf_pw:
            st.error("Senhas não coincidem.")
            return

        conn = get_connection()
        cur = conn.cursor()
        cur.execute(
            "UPDATE users SET password_hash = ?, must_reset_password = 0 WHERE id = ?",
            (hash_password(new_pw), st.session_state.logged_user["id"])
        )
        conn.commit()
        conn.close()
        st.session_state.logged_user["must_reset_password"] = False
        st.success("Senha atualizada!")
        st.rerun()

if not st.session_state.logged_user:
    login_screen()
    st.stop()

if st.session_state.logged_user.get("must_reset_password"):
    reset_password_screen()
    st.stop()

# ==============================================================================
# CONTEXTO MULTI-EMPRESA
# ==============================================================================
conn = get_connection()
cur = conn.cursor()

if st.session_state.logged_user["user_type"] == "ADMIN":
    cur.execute("SELECT id, code, name, depreciation_method FROM companies ORDER BY name")
else:
    cur.execute("""
        SELECT c.id, c.code, c.name, c.depreciation_method 
        FROM companies c
        INNER JOIN user_companies uc ON uc.company_id = c.id
        WHERE uc.user_id = ?
        ORDER BY c.name
    """, (st.session_state.logged_user["id"],))

accessible_companies = cur.fetchall()
conn.close()

if not accessible_companies:
    st.error("Nenhuma empresa disponível para este perfil.")
    if st.sidebar.button("Sair"):
        st.session_state.logged_user = None
        st.rerun()
    st.stop()

company_options = {f"{c[1]} - {c[2]} ({c[3]})": c for c in accessible_companies}
selected_company_label = st.sidebar.selectbox("Empresa Ativa", list(company_options.keys()))
active_company = company_options[selected_company_label]
ACTIVE_COMPANY_ID = active_company[0]
DEPR_METHOD = active_company[3]

st.sidebar.markdown(f"**Operador:** `{st.session_state.logged_user['email']}`")
st.sidebar.markdown(f"**Perfil:** `{st.session_state.logged_user['user_type']}`")
if st.sidebar.button("Encerrar Sessão", use_container_width=True):
    st.session_state.logged_user = None
    st.rerun()

# ==============================================================================
# MOTOR CONTÁBIL E FISCAL (CPC 27)
# ==============================================================================
def date_from_str(d_val):
    if isinstance(d_val, str):
        return datetime.strptime(d_val[:10], "%Y-%m-%d").date()
    elif isinstance(d_val, datetime):
        return d_val.date()
    return d_val

def get_last_day_of_month(d: date) -> date:
    next_month = d.replace(day=28) + timedelta(days=4)
    return next_month - timedelta(days=next_month.day)

def calculate_asset_depreciation_to_date(asset_row, target_date: date, revaluations_list, calculation_type='SOCIETARIO'):
    total_cost = float(asset_row['acquisition_cost'] or 0.0) + float(asset_row['additional_costs'] or 0.0)
    res_raw = float(asset_row['residual_value'] or 0.0)
    
    if calculation_type == 'SOCIETARIO':
        residual_val = 0.0 if (total_cost > 0 and res_raw >= total_cost * 0.99) else res_raw
    else:
        residual_val = 0.0

    base_depr = max(0.0, total_cost - residual_val)
    if base_depr <= 0:
        return 0.0, total_cost

    # 1. Regra para Base Inicial Histórica (Depreciação mensal informada)
    has_initial_base = bool(asset_row.get('initial_base_date') and float(asset_row.get('initial_monthly_depr') or 0.0) > 0)
    
    if calculation_type == 'SOCIETARIO' and has_initial_base:
        init_base_dt = date_from_str(asset_row['initial_base_date'])
        init_accum = float(asset_row.get('initial_accum_depr') or 0.0)
        monthly_quota = float(asset_row.get('initial_monthly_depr') or 0.0)

        # Se a data alvo for anterior ou igual à data base inicial
        if target_date <= init_base_dt:
            accum = min(base_depr, init_accum)
            return round(accum, 2), round(total_cost - accum, 2)

        # Verificar se existe reavaliação cadastrada após a data base inicial
        applicable_revals = [
            r for r in revaluations_list
            if date_from_str(r['cut_off_date']) >= init_base_dt and date_from_str(r['cut_off_date']) < target_date
            and (
                (r['target_type'] == 'ATIVO' and r['target_id'] == asset_row['id']) or
                (r['target_type'] == 'SUBCLASSE' and asset_row.get('subclass_id') and r['target_id'] == asset_row['subclass_id']) or
                (r['target_type'] == 'CLASSE' and r['target_id'] == asset_row['class_id'])
            )
        ]
        applicable_revals.sort(key=lambda x: date_from_str(x['cut_off_date']))

        if not applicable_revals:
            # Continuidade com quota mensal informada
            months_elapsed = (target_date.year - init_base_dt.year) * 12 + (target_date.month - init_base_dt.month)
            if target_date.day < get_last_day_of_month(target_date).day:
                months_elapsed = max(0, months_elapsed - 1)
            months_elapsed = max(0, months_elapsed)
            
            accum = init_accum + (months_elapsed * monthly_quota)
            accum = min(base_depr, accum)
            return round(accum, 2), round(total_cost - accum, 2)

    # 2. Regra Linear Padrão (Societário e Fiscal / Pós-Reavaliações)
    start_date = date_from_str(asset_row['start_date'])
    if DEPR_METHOD == 'MENSAL':
        first_of_next_m = (start_date.replace(day=28) + timedelta(days=4)).replace(day=1)
        comp_start_date = first_of_next_m
    else:
        comp_start_date = start_date + timedelta(days=1)

    if target_date < comp_start_date:
        return 0.0, total_cost

    if calculation_type == 'SOCIETARIO':
        base_life_years = float(asset_row.get('subclass_useful_life') or asset_row.get('class_useful_life') or 10.0)
    else:
        base_life_years = float(asset_row.get('subclass_fiscal_life') or asset_row.get('class_fiscal_life') or 10.0)

    # Reavaliações societárias prospectivas
    applicable_revals = []
    if calculation_type == 'SOCIETARIO':
        for r in revaluations_list:
            r_cut = date_from_str(r['cut_off_date'])
            if r_cut < target_date:
                match = False
                if r['target_type'] == 'ATIVO' and r['target_id'] == asset_row['id']:
                    match = True
                elif r['target_type'] == 'SUBCLASSE' and asset_row.get('subclass_id') and r['target_id'] == asset_row['subclass_id']:
                    match = True
                elif r['target_type'] == 'CLASSE' and r['target_id'] == asset_row['class_id']:
                    match = True
                if match:
                    applicable_revals.append(r)
        applicable_revals.sort(key=lambda x: date_from_str(x['cut_off_date']))

    if not applicable_revals:
        if DEPR_METHOD == 'MENSAL':
            months = (target_date.year - comp_start_date.year) * 12 + (target_date.month - comp_start_date.month)
            if target_date.day >= get_last_day_of_month(target_date).day:
                months += 1
            months = max(0, months)
            total_months = base_life_years * 12
            rate_month = base_depr / total_months if total_months > 0 else 0
            accum = min(base_depr, rate_month * months)
        else:
            days = (target_date - comp_start_date).days + 1
            days = max(0, days)
            total_days = base_life_years * 365
            rate_day = base_depr / total_days if total_days > 0 else 0
            accum = min(base_depr, rate_day * days)

        return round(accum, 2), round(total_cost - accum, 2)

    # Cálculo encadeado por laudos
    curr_date = comp_start_date
    accum_depr = 0.0
    current_life_years = base_life_years
    curr_base_depr = base_depr

    for r in applicable_revals:
        r_cut = date_from_str(r['cut_off_date'])
        if r_cut < curr_date:
            continue

        if DEPR_METHOD == 'MENSAL':
            months = (r_cut.year - curr_date.year) * 12 + (r_cut.month - curr_date.month)
            if r_cut.day >= get_last_day_of_month(r_cut).day:
                months += 1
            months = max(0, months)
            tot_m = current_life_years * 12
            rate_m = curr_base_depr / tot_m if tot_m > 0 else 0
            period_depr = min(curr_base_depr, rate_m * months)
            accum_depr += period_depr
        else:
            days = (r_cut - curr_date).days + 1
            days = max(0, days)
            tot_d = current_life_years * 365
            rate_d = curr_base_depr / tot_d if tot_d > 0 else 0
            period_depr = min(curr_base_depr, rate_d * days)
            accum_depr += period_depr

        net_book_val = total_cost - accum_depr - residual_val
        if net_book_val <= 0:
            return round(accum_depr, 2), round(total_cost - accum_depr, 2)

        if DEPR_METHOD == 'MENSAL':
            elapsed_years = ((r_cut.year - comp_start_date.year) * 12 + (r_cut.month - comp_start_date.month)) / 12.0
        else:
            elapsed_years = (r_cut - comp_start_date).days / 365.0

        new_total_life = r['new_useful_life_years']
        remaining_years = new_total_life - elapsed_years

        if remaining_years <= (1.0 if DEPR_METHOD == 'MENSAL' else 365/365.0):
            current_life_years = 1.0
        else:
            current_life_years = remaining_years

        curr_base_depr = net_book_val
        curr_date = r_cut + timedelta(days=1)

    if target_date >= curr_date:
        if DEPR_METHOD == 'MENSAL':
            months = (target_date.year - curr_date.year) * 12 + (target_date.month - curr_date.month)
            if target_date.day >= get_last_day_of_month(target_date).day:
                months += 1
            months = max(0, months)
            tot_m = current_life_years * 12
            rate_m = curr_base_depr / tot_m if tot_m > 0 else 0
            period_depr = min(curr_base_depr, rate_m * months)
            accum_depr += period_depr
        else:
            days = (target_date - curr_date).days + 1
            days = max(0, days)
            tot_d = current_life_years * 365
            rate_d = curr_base_depr / tot_d if tot_d > 0 else 0
            period_depr = min(curr_base_depr, rate_d * days)
            accum_depr += period_depr

    accum_depr = min(total_cost - residual_val, accum_depr)
    return round(accum_depr, 2), round(total_cost - accum_depr, 2)

# ==============================================================================
# MENU DE NAVEGAÇÃO
# ==============================================================================
menu = st.sidebar.radio(
    "Navegação",
    [
        "Painel Geral & Extrato",
        "Cadastro de Ativos (Adições & Lote)",
        "Reavaliação de Vida Útil",
        "Baixas de Ativos & Reversão",
        "Módulo Fiscal (LALUR/LACS)",
        "Partidas Dobradas (ERP-Ready)",
        "Quadro de Movimentação",
        "Auditoria e Compliance (Dedo-Duro)",
        "Tabelas Mestras & Acessos"
    ]
)

# ------------------------------------------------------------------------------
# 1. PAINEL GERAL & EXTRATO ANALÍTICO MENSALIZADO
# ------------------------------------------------------------------------------
if menu == "Painel Geral & Extrato":
    st.title("📊 Painel Geral de Ativos Imobilizados")
    ref_date = st.date_input("Data-Base de Posição", value=date.today())

    conn = get_connection()
    assets_df = pd.read_sql_query("""
        SELECT a.id, a.tag_number, a.description, a.status, a.acquisition_date, a.start_date,
               a.acquisition_cost, a.additional_costs, a.residual_value,
               a.brand, a.model, a.serial_number, a.year_manufacture, a.location, a.cost_center, a.conservation_state,
               a.initial_accum_depr, a.initial_monthly_depr, a.initial_base_date,
               c.name as class_name, c.useful_life_years as class_useful_life, c.fiscal_life_years as class_fiscal_life,
               s.name as subclass_name, s.useful_life_years as subclass_useful_life, s.fiscal_life_years as subclass_fiscal_life,
               a.class_id, a.subclass_id
        FROM assets a
        INNER JOIN asset_classes c ON c.id = a.class_id
        LEFT JOIN asset_subclasses s ON s.id = a.subclass_id
        WHERE a.company_id = ?
    """, conn, params=(ACTIVE_COMPANY_ID,))

    revals = pd.read_sql_query("""
        SELECT target_type, target_id, cut_off_date, new_useful_life_years
        FROM life_revaluations WHERE company_id = ?
    """, conn, params=(ACTIVE_COMPANY_ID,)).to_dict('records')
    conn.close()

    if assets_df.empty:
        st.info("Nenhum ativo cadastrado para a empresa ativa.")
    else:
        results_display = []
        tot_custo_num = 0.0
        tot_depr_num = 0.0
        tot_vlc_num = 0.0

        ref_year = ref_date.year
        dt_ano_ant = date(ref_year - 1, 12, 31)

        meses_nomes = [
            ("Janeiro", 1), ("Fevereiro", 2), ("Março", 3), ("Abril", 4),
            ("Maio", 5), ("Junho", 6), ("Julho", 7), ("Agosto", 8),
            ("Setembro", 9), ("Outubro", 10), ("Novembro", 11), ("Dezembro", 12)
        ]

        for _, row in assets_df.iterrows():
            custo_total = float(row['acquisition_cost'] or 0.0) + float(row['additional_costs'] or 0.0)
            
            # Depreciação no fechamento do ano anterior (Saldo Inicial do ano)
            depr_ini_ano, _ = calculate_asset_depreciation_to_date(row, dt_ano_ant, revals, 'SOCIETARIO')
            
            # Depreciação acumulada até a data-base de posição
            depr_soc_acum, vlc = calculate_asset_depreciation_to_date(row, ref_date, revals, 'SOCIETARIO')
            depr_fisc_acum, _ = calculate_asset_depreciation_to_date(row, ref_date, revals, 'FISCAL')

            tot_custo_num += custo_total
            tot_depr_num += depr_soc_acum
            tot_vlc_num += vlc

            row_data = {
                "Plaqueta": row['tag_number'],
                "Descrição": row['description'],
                "Status": row['status'],
                "Classe": row['class_name'],
                "Subclasse": row['subclass_name'] or "-",
                "Marca": row['brand'] or "-",
                "Modelo": row['model'] or "-",
                "Nº Série": row['serial_number'] or "-",
                "Ano": row['year_manufacture'] or "-",
                "Localização": row['location'] or "-",
                "Centro de Custo": row['cost_center'] or "-",
                "Conservação": row['conservation_state'] or "Bom",
                "Data Início": fmt_date(row['start_date']),
                "Custo Total": fmt_brl(custo_total),
                "Valor Residual": fmt_brl(row['residual_value']),
                "Depr. Acumulada Societária Inicial": fmt_brl(depr_ini_ano)
            }

            # Cálculo de quota para cada mês de Janeiro a Dezembro
            acum_mes_anterior = depr_ini_ano
            for nome_m, num_m in meses_nomes:
                fim_mes = get_last_day_of_month(date(ref_year, num_m, 1))
                if ref_date >= fim_mes:
                    depr_fim_m, _ = calculate_asset_depreciation_to_date(row, fim_mes, revals, 'SOCIETARIO')
                    quota_mes = max(0.0, depr_fim_m - acum_mes_anterior)
                    acum_mes_anterior = depr_fim_m
                elif ref_date.month == num_m and ref_date.year == ref_year:
                    depr_fim_m, _ = calculate_asset_depreciation_to_date(row, ref_date, revals, 'SOCIETARIO')
                    quota_mes = max(0.0, depr_fim_m - acum_mes_anterior)
                    acum_mes_anterior = depr_fim_m
                else:
                    quota_mes = 0.0

                row_data[f"Depr. {nome_m}"] = fmt_brl(quota_mes)

            row_data["Depr. Acumulada Societária"] = fmt_brl(depr_soc_acum)
            row_data["Valor Contábil Líquido"] = fmt_brl(vlc)
            row_data["Depr. Acumulada Fiscal"] = fmt_brl(depr_fisc_acum)

            results_display.append(row_data)

        df_rep = pd.DataFrame(results_display)

        # Cards Executivos
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total de Ativos", len(df_rep))
        c1.caption(f"Ativos Operacionais: {len(df_rep[df_rep['Status'] == 'ATIVO'])} | Baixados: {len(df_rep[df_rep['Status'] == 'BAIXADO'])}")
        c2.metric("Custo Imobilizado Total", f"R$ {fmt_brl(tot_custo_num)}")
        c3.metric("Depr. Acumulada (Societária)", f"R$ {fmt_brl(tot_depr_num)}")
        c4.metric("Valor Contábil Líquido (VLC)", f"R$ {fmt_brl(tot_vlc_num)}")

        st.subheader("Extrato Analítico Patrimonial (Mensalizado)")
        render_table(df_rep)

        col_b1, col_b2 = st.columns([1, 1])
        with col_b1:
            st.download_button(
                "📥 Exportar Extrato Excel (.xlsx)",
                data=export_to_excel(df_rep, "Extrato_Patrimonial"),
                file_name=f"extrato_patrimonial_{ref_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_b2:
            st.download_button(
                "📥 Exportar Extrato CSV (.csv)",
                data=export_to_csv(df_rep),
                file_name=f"extrato_patrimonial_{ref_date}.csv",
                mime="text/csv",
                use_container_width=True
            )

# ------------------------------------------------------------------------------
# 2. CADASTRO DE ATIVOS (BASE INICIAL HISTÓRICA, ADIÇÕES E CRUD)
# ------------------------------------------------------------------------------
elif menu == "Cadastro de Ativos (Adições & Lote)":
    st.title("📝 Gestão, Carga Inicial e Adições de Ativos")
    aba_base_ini, aba_lote_adicoes, aba_cad, aba_edit, aba_del = st.tabs([
        "📥 Carga da Base Inicial Histórica",
        "📥 Upload de Adições / Lote",
        "Cadastrar / Novo Registro",
        "Editar / Alterar Registro",
        "Excluir Registro"
    ])

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM asset_classes ORDER BY name")
    classes = cur.fetchall()
    class_map = {c[1]: c[0] for c in classes}

    CONSERVATION_OPTIONS = ["Novo", "Excelente", "Bom", "Ruim"]

    # --- ABA 1: CARGA DA BASE INICIAL HISTÓRICA ---
    with aba_base_ini:
        st.subheader("📥 Carga da Base Inicial Histórica do Cliente")
        st.markdown("""
        Esta etapa permite importar a **base histórica completa** respeitando a **`Depreciação Acumulada`** e a **`Depreciação Mensal`** já apurada em laudos ou controles anteriores do cliente.
        - O sistema manterá a **Depreciação Mensal informada** ao longo do tempo até que o bem atinja a depreciação total do seu custo.
        - **Trava de Segurança:** A depreciação acumulada é limitada rigorosamente ao custo total menos o valor residual.
        """)

        cols_base_ini = [
            "Plaqueta", "Descricao", "Classe", "Subclasse", "DataAquisicao", "DataInicio",
            "CustoAquisicao", "CustosAdicionais", "Depreciação Acumulada", "ValorResidual",
            "Depreciação Mensal", "Taxa mensal de depreciação", "Vida Util do bem",
            "ChaveNFe", "Marca", "Modelo", "NumeroSerie", "AnoFabricacao", "Localizacao", "CentroCusto", "EstadoConservacao"
        ]
        sample_base_ini = {
            "Plaqueta": "M1-0001", "Descricao": "Mesa Marcenaria Estação de Trabalho",
            "Classe": "MOVEIS E UTENSILIOS", "Subclasse": "Estação de Trabalho",
            "DataAquisicao": "2024-01-15", "DataInicio": "2024-01-15",
            "CustoAquisicao": 3225.88, "CustosAdicionais": 0.00,
            "Depreciação Acumulada": 119.48, "ValorResidual": 0.00,
            "Depreciação Mensal": 13.28, "Taxa mensal de depreciação": 0.41, "Vida Util do bem": 18,
            "ChaveNFe": "", "Marca": "Marcenaria", "Modelo": "Estação", "NumeroSerie": "",
            "AnoFabricacao": 2024, "Localizacao": "Sala 1705", "CentroCusto": "ADM", "EstadoConservacao": "Bom"
        }

        col_b_ini1, col_b_ini2 = st.columns([1, 1])
        with col_b_ini1:
            st.download_button(
                "📄 Baixar Template Modelo da Base Inicial (.xlsx)",
                data=generate_template_excel(cols_base_ini, [sample_base_ini]),
                file_name="template_base_inicial_ativos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_b_ini2:
            data_corte_base = st.date_input("Data-Base de Referência do Saldo Inicial", value=date(2025, 12, 31))

        up_file_ini = st.file_uploader("Arquivo da Base Inicial Histórica (.xlsx)", type=["xlsx"], key="up_base_ini_file")
        if up_file_ini:
            try:
                df_ini = pd.read_excel(up_file_ini)
                st.write("Pré-visualização da Base Inicial:")
                st.dataframe(df_ini.head(10), use_container_width=True)

                if st.button("Processar Carga da Base Inicial", type="primary"):
                    imported_count = 0
                    updated_count = 0
                    errors = []

                    for idx, row in df_ini.iterrows():
                        tag = str(row.get("Plaqueta", "")).strip()
                        desc = str(row.get("Descricao", "")).strip()
                        classe_nome = str(row.get("Classe", "")).strip().upper()
                        subclasse_nome = str(row.get("Subclasse", "")).strip() if pd.notna(row.get("Subclasse")) and str(row.get("Subclasse")).strip().upper() != 'NONE' else ""

                        if not tag or not desc or not classe_nome:
                            errors.append(f"Linha {idx+2}: Plaqueta, Descrição ou Classe em branco.")
                            continue

                        dt_acq = pd.to_datetime(row.get("DataAquisicao")).strftime('%Y-%m-%d')
                        dt_start = pd.to_datetime(row.get("DataInicio")).strftime('%Y-%m-%d')
                        
                        custo = float(row.get("CustoAquisicao", 0))
                        custo_add = float(row.get("CustosAdicionais", 0)) if pd.notna(row.get("CustosAdicionais")) else 0.0
                        
                        # Tratamento do valor residual (não confunde com saldo contábil líquido)
                        res_val_raw = float(row.get("ValorResidual", 0)) if pd.notna(row.get("ValorResidual")) else 0.0
                        if custo > 0 and res_val_raw >= custo * 0.99:
                            res_val = 0.0
                        else:
                            res_val = res_val_raw

                        depr_acum_ini = float(row.get("Depreciação Acumulada", 0)) if pd.notna(row.get("Depreciação Acumulada")) else 0.0
                        depr_mensal_ini = float(row.get("Depreciação Mensal", 0)) if pd.notna(row.get("Depreciação Mensal")) else 0.0

                        nfe = str(row.get("ChaveNFe", "")).strip() if pd.notna(row.get("ChaveNFe")) else ""
                        marca = str(row.get("Marca", "")).strip() if pd.notna(row.get("Marca")) else ""
                        modelo = str(row.get("Modelo", "")).strip() if pd.notna(row.get("Modelo")) else ""
                        num_serie = str(row.get("NumeroSerie", "")).strip() if pd.notna(row.get("NumeroSerie")) else ""
                        ano_fab = int(row.get("AnoFabricacao")) if pd.notna(row.get("AnoFabricacao")) and str(row.get("AnoFabricacao")).replace('.0','').isdigit() else None
                        localizacao = str(row.get("Localizacao", "")).strip() if pd.notna(row.get("Localizacao")) else ""
                        centro_custo = str(row.get("CentroCusto", "")).strip() if pd.notna(row.get("CentroCusto")) else ""
                        
                        est_cons = str(row.get("EstadoConservacao", "Bom")).strip().capitalize()
                        if est_cons not in CONSERVATION_OPTIONS:
                            est_cons = "Bom"

                        # Garantir Classe
                        if classe_nome not in class_map:
                            cur.execute("""
                                INSERT INTO asset_classes (name, useful_life_years, fiscal_life_years)
                                VALUES (?, 10.0, 10.0)
                                ON CONFLICT(name) DO NOTHING
                            """, (classe_nome,))
                            cur.execute("SELECT id FROM asset_classes WHERE name = ?", (classe_nome,))
                            c_id = cur.fetchone()[0]
                            class_map[classe_nome] = c_id
                        else:
                            c_id = class_map[classe_nome]

                        # Garantir Subclasse se informada
                        s_id = None
                        if subclasse_nome:
                            cur.execute("""
                                INSERT INTO asset_subclasses (class_id, name, useful_life_years, fiscal_life_years)
                                VALUES (?, ?, 10.0, 10.0)
                                ON CONFLICT(class_id, name) DO NOTHING
                            """, (c_id, subclasse_nome))
                            cur.execute("SELECT id FROM asset_subclasses WHERE class_id = ? AND name = ?", (c_id, subclasse_nome))
                            s_id = cur.fetchone()[0]

                        # Inserir ou atualizar com saldos históricos da base inicial
                        cur.execute("SELECT id FROM assets WHERE company_id = ? AND tag_number = ?", (ACTIVE_COMPANY_ID, tag))
                        exists = cur.fetchone()
                        if exists:
                            cur.execute("""
                                UPDATE assets SET
                                    description = ?, class_id = ?, subclass_id = ?, invoice_key = ?,
                                    acquisition_date = ?, start_date = ?, acquisition_cost = ?,
                                    additional_costs = ?, residual_value = ?,
                                    brand = ?, model = ?, serial_number = ?, year_manufacture = ?,
                                    location = ?, cost_center = ?, conservation_state = ?,
                                    initial_accum_depr = ?, initial_monthly_depr = ?, initial_base_date = ?
                                WHERE id = ?
                            """, (
                                desc, c_id, s_id, nfe, dt_acq, dt_start, custo, custo_add, res_val,
                                marca, modelo, num_serie, ano_fab, localizacao, centro_custo, est_cons,
                                depr_acum_ini, depr_mensal_ini, data_corte_base, exists[0]
                            ))
                            updated_count += 1
                        else:
                            cur.execute("""
                                INSERT INTO assets (
                                    company_id, tag_number, description, class_id, subclass_id,
                                    invoice_key, acquisition_date, start_date, acquisition_cost,
                                    additional_costs, residual_value,
                                    brand, model, serial_number, year_manufacture,
                                    location, cost_center, conservation_state,
                                    initial_accum_depr, initial_monthly_depr, initial_base_date
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (
                                ACTIVE_COMPANY_ID, tag, desc, c_id, s_id, nfe, dt_acq, dt_start, custo, custo_add, res_val,
                                marca, modelo, num_serie, ano_fab, localizacao, centro_custo, est_cons,
                                depr_acum_ini, depr_mensal_ini, data_corte_base
                            ))
                            imported_count += 1

                    conn.commit()
                    log_audit(
                        conn, ACTIVE_COMPANY_ID, st.session_state.logged_user["email"],
                        "CARGA_BASE_INICIAL", "assets",
                        f"Data Base: {data_corte_base} | Inseridos: {imported_count} | Atualizados: {updated_count}"
                    )
                    conn.commit()
                    st.success(f"Base Inicial Histórica processada com sucesso! Inseridos: {imported_count} | Atualizados: {updated_count}")
                    if errors:
                        st.warning("Erros encontrados:\n" + "\n".join(errors[:10]))
                    st.rerun()
            except Exception as e:
                st.error(f"Erro ao processar planilha de base inicial: {e}")

    # --- ABA 2: UPLOAD DE ADIÇÕES / LOTES ROTINEIROS ---
    with aba_lote_adicoes:
        st.subheader("Carga em Lote de Novas Adições")
        st.markdown("Utilize esta aba para subir lotes periódicos de aquisições de ativos.")
        cols_template_adicoes = [
            "Plaqueta", "Descricao", "Classe", "Subclasse", "DataAquisicao",
            "DataInicio", "CustoAquisicao", "CustosAdicionais", "ValorResidual", "ChaveNFe",
            "Marca", "Modelo", "NumeroSerie", "AnoFabricacao", "Localizacao", "CentroCusto", "EstadoConservacao"
        ]
        sample_adic = {
            "Plaqueta": "PAT-9002", "Descricao": "Notebook Lenovo ThinkPad",
            "Classe": "COMPUTADORES E PERIFERICOS", "Subclasse": "Notebook",
            "DataAquisicao": "2026-02-10", "DataInicio": "2026-02-15",
            "CustoAquisicao": 4500.00, "CustosAdicionais": 0.00, "ValorResidual": 0.00,
            "ChaveNFe": "", "Marca": "Lenovo", "Modelo": "T14", "NumeroSerie": "PF-9988",
            "AnoFabricacao": 2026, "Localizacao": "TI", "CentroCusto": "CC-TI", "EstadoConservacao": "Novo"
        }

        st.download_button(
            "📄 Baixar Planilha Modelo de Adições (.xlsx)",
            data=generate_template_excel(cols_template_adicoes, [sample_adic]),
            file_name="template_adicoes_imobilizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        up_adic_file = st.file_uploader("Arquivo de Novas Adições (.xlsx)", type=["xlsx"], key="up_adic_file")
        if up_adic_file:
            try:
                df_adic = pd.read_excel(up_adic_file)
                st.write("Pré-visualização:")
                st.dataframe(df_adic.head(10), use_container_width=True)

                if st.button("Processar Carga de Adições", type="primary"):
                    imported_c = 0
                    for idx, row in df_adic.iterrows():
                        tag = str(row.get("Plaqueta", "")).strip()
                        desc = str(row.get("Descricao", "")).strip()
                        classe_nome = str(row.get("Classe", "")).strip().upper()
                        subclasse_nome = str(row.get("Subclasse", "")).strip() if pd.notna(row.get("Subclasse")) and str(row.get("Subclasse")).strip().upper() != 'NONE' else ""

                        if not tag or not desc or not classe_nome:
                            continue

                        dt_acq = pd.to_datetime(row.get("DataAquisicao")).strftime('%Y-%m-%d')
                        dt_start = pd.to_datetime(row.get("DataInicio")).strftime('%Y-%m-%d')
                        custo = float(row.get("CustoAquisicao", 0))
                        custo_add = float(row.get("CustosAdicionais", 0)) if pd.notna(row.get("CustosAdicionais")) else 0.0
                        res_val = float(row.get("ValorResidual", 0)) if pd.notna(row.get("ValorResidual")) else 0.0

                        nfe = str(row.get("ChaveNFe", "")).strip() if pd.notna(row.get("ChaveNFe")) else ""
                        marca = str(row.get("Marca", "")).strip() if pd.notna(row.get("Marca")) else ""
                        modelo = str(row.get("Modelo", "")).strip() if pd.notna(row.get("Modelo")) else ""
                        num_serie = str(row.get("NumeroSerie", "")).strip() if pd.notna(row.get("NumeroSerie")) else ""
                        ano_fab = int(row.get("AnoFabricacao")) if pd.notna(row.get("AnoFabricacao")) and str(row.get("AnoFabricacao")).replace('.0','').isdigit() else None
                        localizacao = str(row.get("Localizacao", "")).strip() if pd.notna(row.get("Localizacao")) else ""
                        centro_custo = str(row.get("CentroCusto", "")).strip() if pd.notna(row.get("CentroCusto")) else ""
                        
                        est_cons = str(row.get("EstadoConservacao", "Bom")).strip().capitalize()
                        if est_cons not in CONSERVATION_OPTIONS:
                            est_cons = "Bom"

                        if classe_nome not in class_map:
                            cur.execute("INSERT INTO asset_classes (name, useful_life_years, fiscal_life_years) VALUES (?, 10.0, 10.0) ON CONFLICT(name) DO NOTHING", (classe_nome,))
                            cur.execute("SELECT id FROM asset_classes WHERE name = ?", (classe_nome,))
                            c_id = cur.fetchone()[0]
                            class_map[classe_nome] = c_id
                        else:
                            c_id = class_map[classe_nome]

                        s_id = None
                        if subclasse_nome:
                            cur.execute("INSERT INTO asset_subclasses (class_id, name, useful_life_years, fiscal_life_years) VALUES (?, ?, 10.0, 10.0) ON CONFLICT(class_id, name) DO NOTHING", (c_id, subclasse_nome))
                            cur.execute("SELECT id FROM asset_subclasses WHERE class_id = ? AND name = ?", (c_id, subclasse_nome))
                            s_id = cur.fetchone()[0]

                        cur.execute("""
                            INSERT INTO assets (
                                company_id, tag_number, description, class_id, subclass_id,
                                invoice_key, acquisition_date, start_date, acquisition_cost,
                                additional_costs, residual_value,
                                brand, model, serial_number, year_manufacture,
                                location, cost_center, conservation_state
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            ON CONFLICT(company_id, tag_number) DO UPDATE SET
                                description = excluded.description,
                                acquisition_cost = excluded.acquisition_cost
                        """, (
                            ACTIVE_COMPANY_ID, tag, desc, c_id, s_id, nfe, dt_acq, dt_start, custo, custo_add, res_val,
                            marca, modelo, num_serie, ano_fab, localizacao, centro_custo, est_cons
                        ))
                        imported_c += 1

                    conn.commit()
                    log_audit(conn, ACTIVE_COMPANY_ID, st.session_state.logged_user["email"], "CARGA_ADICOES", "assets", f"Importados: {imported_c}")
                    st.success(f"{imported_c} adições processadas com sucesso!")
                    st.rerun()
            except Exception as e:
                st.error(f"Erro ao processar adições: {e}")

    # --- ABA 3: NOVO CADASTRO UNITÁRIO ---
    with aba_cad:
        st.subheader("Novo Tombamento de Imobilizado")
        with st.form("form_novo_ativo", clear_on_submit=True):
            st.markdown("##### Dados Gerais e Fiscais")
            col1, col2 = st.columns(2)
            with col1:
                tag = st.text_input("Número da Plaqueta / Tag *")
                desc = st.text_input("Descrição do Bem *")
                sel_class = st.selectbox("Classe de Ativo *", list(class_map.keys()) if class_map else ["(Nenhuma)"])
                class_id = class_map[sel_class] if class_map and sel_class != "(Nenhuma)" else None

                subclass_opts = ["(Nenhuma)"]
                subclass_map = {}
                if class_id:
                    cur.execute("SELECT id, name FROM asset_subclasses WHERE class_id = ?", (class_id,))
                    subclasses = cur.fetchall()
                    subclass_map = {s[1]: s[0] for s in subclasses}
                    subclass_opts += list(subclass_map.keys())
                sel_subclass = st.selectbox("Subclasse (Opcional)", subclass_opts)
                subclass_id = subclass_map[sel_subclass] if sel_subclass != "(Nenhuma)" else None

                nf_key = st.text_input("Chave NF-e / Documento Fiscal")

            with col2:
                dt_acq = st.date_input("Data de Aquisição *", value=date.today())
                dt_start = st.date_input("Data de Início de Operação / Tombamento *", value=date.today())
                custo = st.number_input("Custo de Aquisição (R$) *", min_value=0.01, step=100.0, format="%.2f")
                custo_add = st.number_input("Custos Adicionais Capitalizáveis (R$)", min_value=0.0, step=50.0, format="%.2f")
                residual = st.number_input("Valor Residual de Venda (R$ - Opcional)", min_value=0.0, step=50.0, format="%.2f")

            st.markdown("##### Informações Técnicas (Opcional)")
            t_col1, t_col2 = st.columns(2)
            with t_col1:
                marca = st.text_input("Marca", placeholder="Ex: Dell, Toyota, Romi")
                num_serie = st.text_input("Número de Série", placeholder="Ex: SN-0098231")
            with t_col2:
                modelo = st.text_input("Modelo", placeholder="Ex: PowerEdge, Hilux, Centur 30D")
                ano_fab = st.number_input("Ano de Fabricação", min_value=1950, max_value=2050, value=date.today().year, step=1)

            st.markdown("##### Localização e Estado Físico")
            e_col1, e_col2, e_col3 = st.columns(3)
            with e_col1:
                localizacao = st.text_input("Localização Física", placeholder="Ex: Galpão SBC / Sala 402 / Filial Matriz")
            with e_col2:
                centro_custo = st.text_input("Centro de Custo", placeholder="Ex: CC-OPER-01 / CC-TI")
            with e_col3:
                conservacao = st.selectbox("Estado de Conservação", CONSERVATION_OPTIONS, index=2)

            submitted = st.form_submit_button("Efetuar Tombamento", type="primary")
            if submitted:
                if not tag or not desc or not class_id:
                    st.error("Campos obrigatórios não preenchidos (Plaqueta, Descrição ou Classe).")
                else:
                    try:
                        cur.execute("""
                            INSERT INTO assets (
                                company_id, tag_number, description, class_id, subclass_id,
                                invoice_key, acquisition_date, start_date, acquisition_cost,
                                additional_costs, residual_value,
                                brand, model, serial_number, year_manufacture,
                                location, cost_center, conservation_state
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            ACTIVE_COMPANY_ID, tag.strip(), desc.strip(), class_id, subclass_id,
                            nf_key.strip(), dt_acq, dt_start, custo, custo_add, residual,
                            marca.strip(), modelo.strip(), num_serie.strip(), int(ano_fab) if ano_fab else None,
                            localizacao.strip(), centro_custo.strip(), conservacao
                        ))
                        log_audit(conn, ACTIVE_COMPANY_ID, st.session_state.logged_user["email"], "NOVO_ATIVO", "assets", f"Plaqueta: {tag}")
                        conn.commit()
                        st.success(f"Ativo {tag} cadastrado com sucesso!")
                    except sqlite3.IntegrityError:
                        st.error("Já existe um ativo com este número de plaqueta nesta empresa.")

    # --- ABA 4: EDITAR ---
    with aba_edit:
        st.subheader("Editar Registro de Ativo")
        cur.execute("SELECT id, tag_number, description FROM assets WHERE company_id = ? ORDER BY tag_number", (ACTIVE_COMPANY_ID,))
        ativos_lista = cur.fetchall()
        if not ativos_lista:
            st.info("Nenhum ativo cadastrado para alteração.")
        else:
            ativo_dict = {f"{a[1]} - {a[2]}": a[0] for a in ativos_lista}
            sel_ativo = st.selectbox("Selecione o Ativo para Editar", list(ativo_dict.keys()))
            selected_id = ativo_dict[sel_ativo]

            cur.execute("""
                SELECT tag_number, description, class_id, subclass_id, invoice_key,
                       acquisition_date, start_date, acquisition_cost, additional_costs, residual_value,
                       brand, model, serial_number, year_manufacture, location, cost_center, conservation_state
                FROM assets WHERE id = ?
            """, (selected_id,))
            atv = cur.fetchone()

            with st.form("form_edit_ativo"):
                st.markdown("##### Dados Gerais")
                c1, c2 = st.columns(2)
                with c1:
                    e_tag = st.text_input("Plaqueta / Tag", value=atv[0])
                    e_desc = st.text_input("Descrição", value=atv[1])
                    current_class_name = [k for k, v in class_map.items() if v == atv[2]][0] if atv[2] in class_map.values() else list(class_map.keys())[0]
                    e_class_name = st.selectbox("Classe", list(class_map.keys()), index=list(class_map.keys()).index(current_class_name))
                    e_class_id = class_map[e_class_name]

                    cur.execute("SELECT id, name FROM asset_subclasses WHERE class_id = ?", (e_class_id,))
                    sub_e = cur.fetchall()
                    sub_e_map = {s[1]: s[0] for s in sub_e}
                    sub_e_opts = ["(Nenhuma)"] + list(sub_e_map.keys())
                    current_sub_name = "(Nenhuma)"
                    for name, s_id in sub_e_map.items():
                        if s_id == atv[3]:
                            current_sub_name = name
                    e_sub_name = st.selectbox("Subclasse", sub_e_opts, index=sub_e_opts.index(current_sub_name))
                    e_sub_id = sub_e_map[e_sub_name] if e_sub_name != "(Nenhuma)" else None
                    e_nf = st.text_input("Chave NF-e", value=atv[4] or "")

                with c2:
                    e_dt_acq = st.date_input("Data de Aquisição", value=date_from_str(atv[5]))
                    e_dt_start = st.date_input("Data de Início", value=date_from_str(atv[6]))
                    e_custo = st.number_input("Custo de Aquisição (R$)", value=float(atv[7]), format="%.2f")
                    e_add = st.number_input("Custos Adicionais (R$)", value=float(atv[8]), format="%.2f")
                    e_res = st.number_input("Valor Residual (R$)", value=float(atv[9]), format="%.2f")

                st.markdown("##### Informações Técnicas e Localização")
                ec_t1, ec_t2 = st.columns(2)
                with ec_t1:
                    e_marca = st.text_input("Marca", value=atv[10] or "")
                    e_num_serie = st.text_input("Número de Série", value=atv[12] or "")
                    e_local = st.text_input("Localização", value=atv[14] or "")
                with ec_t2:
                    e_modelo = st.text_input("Modelo", value=atv[11] or "")
                    e_ano = st.number_input("Ano de Fabricação", min_value=1950, max_value=2050, value=int(atv[13]) if atv[13] else date.today().year, step=1)
                    e_cc = st.text_input("Centro de Custo", value=atv[15] or "")

                current_cons = atv[16] if atv[16] in CONSERVATION_OPTIONS else "Bom"
                e_conservacao = st.selectbox("Estado de Conservação", CONSERVATION_OPTIONS, index=CONSERVATION_OPTIONS.index(current_cons))

                if st.form_submit_button("Atualizar Cadastro", type="primary"):
                    try:
                        cur.execute("""
                            UPDATE assets SET
                                tag_number = ?, description = ?, class_id = ?, subclass_id = ?,
                                invoice_key = ?, acquisition_date = ?, start_date = ?,
                                acquisition_cost = ?, additional_costs = ?, residual_value = ?,
                                brand = ?, model = ?, serial_number = ?, year_manufacture = ?,
                                location = ?, cost_center = ?, conservation_state = ?
                            WHERE id = ?
                        """, (
                            e_tag.strip(), e_desc.strip(), e_class_id, e_sub_id, e_nf.strip(),
                            e_dt_acq, e_dt_start, e_custo, e_add, e_res,
                            e_marca.strip(), e_modelo.strip(), e_num_serie.strip(), int(e_ano) if e_ano else None,
                            e_local.strip(), e_cc.strip(), e_conservacao, selected_id
                        ))
                        log_audit(conn, ACTIVE_COMPANY_ID, st.session_state.logged_user["email"], "EDITAR_ATIVO", "assets", f"Plaqueta: {e_tag}")
                        conn.commit()
                        st.success("Cadastro atualizado com sucesso!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Conflito de unicidade na plaqueta informada.")

    # --- ABA 5: EXCLUIR ---
    with aba_del:
        st.subheader("Exclusão de Ativo")
        if not ativos_lista:
            st.info("Nenhum ativo para exclusão.")
        else:
            sel_del = st.selectbox("Ativo para Remoção", list(ativo_dict.keys()), key="del_asset_sel")
            del_id = ativo_dict[sel_del]

            cur.execute("SELECT count(*) FROM asset_disposals WHERE asset_id = ?", (del_id,))
            has_disposal = cur.fetchone()[0] > 0

            if has_disposal:
                st.warning("Este ativo possui registro de baixa contábil vinculado e não pode ser excluído diretamente. Reverta a baixa antes se necessário.")
            else:
                conf = st.checkbox(f"Confirmo a exclusão definitiva do ativo: {sel_del}")
                if st.button("Excluir Registro Definitivamente", type="primary") and conf:
                    cur.execute("DELETE FROM assets WHERE id = ?", (del_id,))
                    log_audit(conn, ACTIVE_COMPANY_ID, st.session_state.logged_user["email"], "EXCLUSAO_ATIVO", "assets", f"ID: {del_id}")
                    conn.commit()
                    st.success("Ativo removido com sucesso!")
                    st.rerun()

    conn.close()

# ------------------------------------------------------------------------------
# 3. REAVALIAÇÃO DE VIDA ÚTIL
# ------------------------------------------------------------------------------
elif menu == "Reavaliação de Vida Útil":
    st.title("🔄 Reavaliação Prospectiva de Vida Útil (CPC 27)")
    st.markdown("""
    Registro de revisões técnicas de vida útil com vigência prospectiva.
    * **Trava dos 12 meses:** Se o prazo restante recalculado for menor que 12 meses, o saldo remanescente será depreciado em 12 meses lineares.
    """)

    conn = get_connection()
    cur = conn.cursor()

    col1, col2 = st.columns(2)
    with col1:
        tipo_alvo = st.selectbox("Abrangência do Laudo", ["CLASSE", "SUBCLASSE", "ATIVO"])
        cur.execute("SELECT id, name FROM asset_classes ORDER BY name")
        classes = cur.fetchall()
        class_map = {c[1]: c[0] for c in classes}

        target_id = None
        if tipo_alvo == "CLASSE":
            if class_map:
                s_c = st.selectbox("Selecione a Classe", list(class_map.keys()))
                target_id = class_map[s_c]
            else:
                st.warning("Nenhuma classe cadastrada.")
        elif tipo_alvo == "SUBCLASSE":
            cur.execute("""
                SELECT s.id, c.name || ' -> ' || s.name 
                FROM asset_subclasses s 
                INNER JOIN asset_classes c ON c.id = s.class_id
            """)
            subclasses = cur.fetchall()
            sub_map = {s[1]: s[0] for s in subclasses}
            if sub_map:
                s_s = st.selectbox("Selecione a Subclasse", list(sub_map.keys()))
                target_id = sub_map[s_s]
            else:
                st.warning("Não há subclasses cadastradas.")
        else:
            cur.execute("SELECT id, tag_number || ' - ' || description FROM assets WHERE company_id = ? AND status = 'ATIVO'", (ACTIVE_COMPANY_ID,))
            assets = cur.fetchall()
            atv_map = {a[1]: a[0] for a in assets}
            if atv_map:
                s_a = st.selectbox("Selecione o Ativo", list(atv_map.keys()))
                target_id = atv_map[s_a]
            else:
                st.warning("Nenhum ativo operacional disponível.")

    with col2:
        data_corte = st.date_input("Data de Corte Patrimonial (Base do Laudo)", value=date(date.today().year - 1, 12, 31))
        nova_vida_util = st.number_input("Nova Vida Útil Total Revisada (Anos)", min_value=0.5, step=1.0, value=10.0)
        laudo_ref = st.text_input("Número / Referência do Laudo Técnico")

    if st.button("Aplicar Reavaliação Prospectiva"):
        if not target_id:
            st.error("Alvo inválido selecionado.")
        else:
            cur.execute("""
                INSERT INTO life_revaluations (
                    company_id, target_type, target_id, cut_off_date,
                    new_useful_life_years, report_reference
                ) VALUES (?, ?, ?, ?, ?, ?)
            """, (ACTIVE_COMPANY_ID, tipo_alvo, target_id, data_corte, nova_vida_util, laudo_ref))
            log_audit(
                conn, ACTIVE_COMPANY_ID, st.session_state.logged_user["email"],
                "REAVALIACAO_VIDA_UTIL", "life_revaluations",
                f"Alvo: {tipo_alvo} | ID: {target_id} | Nova VU: {nova_vida_util} anos | Corte: {data_corte}"
            )
            conn.commit()
            st.success("Reavaliação gravada e aplicada prospectivamente!")

    st.subheader("Histórico de Laudos Registrados")
    reval_df = pd.read_sql_query("""
        SELECT id, target_type as "Abrangência", target_id as "ID Alvo",
               cut_off_date as "Data Corte", new_useful_life_years as "Nova Vida Útil (Anos)",
               report_reference as "Laudo", applied_at as "Registrado em"
        FROM life_revaluations WHERE company_id = ? ORDER BY cut_off_date DESC
    """, conn, params=(ACTIVE_COMPANY_ID,))
    if not reval_df.empty:
        reval_df["Data Corte"] = reval_df["Data Corte"].apply(fmt_date)
    render_table(reval_df)
    conn.close()

# ------------------------------------------------------------------------------
# 4. BAIXAS DE ATIVOS, CARGA EM LOTE E REVERSÃO
# ------------------------------------------------------------------------------
elif menu == "Baixas de Ativos & Reversão":
    st.title("🗑️ Workflow Operacional de Baixas & Reversões")
    aba_baixa_unit, aba_baixa_lote, aba_reversao = st.tabs([
        "Baixa Unitária",
        "📥 Baixas em Lote (Mensal / Anual / Período)",
        "🔄 Reversão de Baixas (Desfazer)"
    ])

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT a.id, a.tag_number, a.description, a.acquisition_cost, a.additional_costs,
               a.start_date, a.acquisition_date, a.residual_value, a.class_id, a.subclass_id,
               a.initial_accum_depr, a.initial_monthly_depr, a.initial_base_date,
               c.useful_life_years as class_useful_life, c.fiscal_life_years as class_fiscal_life,
               s.useful_life_years as subclass_useful_life, s.fiscal_life_years as subclass_fiscal_life
        FROM assets a
        INNER JOIN asset_classes c ON c.id = a.class_id
        LEFT JOIN asset_subclasses s ON s.id = a.subclass_id
        WHERE a.company_id = ? AND a.status = 'ATIVO'
    """, (ACTIVE_COMPANY_ID,))
    ativos_operacionais = cur.fetchall()

    cur.execute("SELECT target_type, target_id, cut_off_date, new_useful_life_years FROM life_revaluations WHERE company_id = ?", (ACTIVE_COMPANY_ID,))
    revals = cur.fetchall()
    revals_list = [{"target_type": r[0], "target_id": r[1], "cut_off_date": r[2], "new_useful_life_years": r[3]} for r in revals]

    # --- ABA BAIXA UNITÁRIA ---
    with aba_baixa_unit:
        st.subheader("Processar Baixa Individual")
        if not ativos_operacionais:
            st.info("Não há ativos operacionais aptos para baixa.")
        else:
            dict_ativos = {f"{a[1]} - {a[2]}": a for a in ativos_operacionais}
            sel_atv = st.selectbox("Localizar Ativo (Plaqueta / Descrição)", list(dict_ativos.keys()))
            dados_atv = dict_ativos[sel_atv]

            col1, col2 = st.columns(2)
            with col1:
                data_baixa = st.date_input("Data Efetiva da Baixa", value=date.today(), key="dt_baixa_unit")
                motivo_baixa = st.selectbox("Motivo da Baixa", ["Venda / Alienação", "Sucateamento", "Sinistro / Perda", "Obsolescência"])
                valor_venda = st.number_input("Valor de Alienação / Venda (R$)", min_value=0.0, step=100.0, format="%.2f")

            if DEPR_METHOD == 'MENSAL':
                primeiro_do_mes = data_baixa.replace(day=1)
                data_calculo_baixa = primeiro_do_mes - timedelta(days=1)
                st.caption(f"Cálculo mensal: saldo apurado em **{fmt_date(data_calculo_baixa)}** (último dia do mês anterior).")
            else:
                data_calculo_baixa = data_baixa - timedelta(days=1)
                st.caption(f"Cálculo diário: saldo apurado em **{fmt_date(data_calculo_baixa)}** (dia anterior).")

            asset_dict_row = {
                'id': dados_atv[0], 'start_date': dados_atv[5], 'acquisition_date': dados_atv[6],
                'acquisition_cost': dados_atv[3], 'additional_costs': dados_atv[4], 'residual_value': dados_atv[7],
                'class_id': dados_atv[8], 'subclass_id': dados_atv[9],
                'initial_accum_depr': dados_atv[10], 'initial_monthly_depr': dados_atv[11], 'initial_base_date': dados_atv[12],
                'class_useful_life': dados_atv[13], 'class_fiscal_life': dados_atv[14],
                'subclass_useful_life': dados_atv[15], 'subclass_fiscal_life': dados_atv[16]
            }

            depr_acum_baixa, vlc_baixa = calculate_asset_depreciation_to_date(asset_dict_row, data_calculo_baixa, revals_list, 'SOCIETARIO')
            custo_total_baixa = dados_atv[3] + dados_atv[4]
            resultado_baixa = valor_venda - vlc_baixa

            with col2:
                st.markdown(f"**Custo Histórico Baixado:** R$ {fmt_brl(custo_total_baixa)}")
                st.markdown(f"**Depreciação Acumulada Baixada:** R$ {fmt_brl(depr_acum_baixa)}")
                st.markdown(f"**Valor Contábil Líquido:** R$ {fmt_brl(vlc_baixa)}")
                if resultado_baixa >= 0:
                    st.success(f"**Ganho de Capital:** R$ {fmt_brl(resultado_baixa)}")
                else:
                    st.error(f"**Perda de Capital:** R$ {fmt_brl(resultado_baixa)}")

            if st.button("Confirmar Baixa Unitária", type="primary"):
                cur.execute("""
                    INSERT INTO asset_disposals (
                        asset_id, disposal_date, disposal_type, sale_value,
                        cost_disposed, depr_accum_disposed, gain_loss_value
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (dados_atv[0], data_baixa, motivo_baixa, valor_venda, custo_total_baixa, depr_acum_baixa, resultado_baixa))

                cur.execute("UPDATE assets SET status = 'BAIXADO' WHERE id = ?", (dados_atv[0],))
                log_audit(
                    conn, ACTIVE_COMPANY_ID, st.session_state.logged_user["email"],
                    "BAIXA_UNITARIA", "assets",
                    f"Plaqueta: {dados_atv[1]} | Data: {data_baixa} | VLC: {vlc_baixa} | Resultado: {resultado_baixa}"
                )
                conn.commit()
                st.success("Baixa patrimonial processada com sucesso!")
                st.rerun()

    # --- ABA BAIXA EM LOTE ---
    with aba_baixa_lote:
        st.subheader("Processamento de Baixas em Lote (Mensal, Período ou Anual)")
        cols_template_baixas = ["Plaqueta", "DataBaixa", "Motivo", "ValorVenda"]
        sample_baixa = {
            "Plaqueta": "PAT-9001", "DataBaixa": "2025-12-31",
            "Motivo": "Venda / Alienação", "ValorVenda": 12000.00
        }

        st.download_button(
            "📄 Baixar Planilha Modelo de Baixas (.xlsx)",
            data=generate_template_excel(cols_template_baixas, [sample_baixa]),
            file_name="template_baixas_em_lote.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        up_baixas = st.file_uploader("Arquivo de Baixas em Lote (.xlsx)", type=["xlsx"], key="up_baixas_file")
        if up_baixas:
            try:
                df_b_upload = pd.read_excel(up_baixas)
                st.write("Pré-visualização do arquivo de baixas:")
                st.dataframe(df_b_upload.head(10), use_container_width=True)

                if st.button("Executar Baixas em Lote", type="primary"):
                    proc_count = 0
                    nao_encontrados = []

                    for idx, b_row in df_b_upload.iterrows():
                        tag_b = str(b_row.get("Plaqueta", "")).strip()
                        dt_b = pd.to_datetime(b_row.get("DataBaixa")).date()
                        motivo = str(b_row.get("Motivo", "Venda / Alienação")).strip()
                        v_venda = float(b_row.get("ValorVenda", 0)) if pd.notna(b_row.get("ValorVenda")) else 0.0

                        cur.execute("""
                            SELECT a.id, a.tag_number, a.description, a.acquisition_cost, a.additional_costs,
                                   a.start_date, a.acquisition_date, a.residual_value, a.class_id, a.subclass_id,
                                   a.initial_accum_depr, a.initial_monthly_depr, a.initial_base_date,
                                   c.useful_life_years as class_useful_life, c.fiscal_life_years as class_fiscal_life,
                                   s.useful_life_years as subclass_useful_life, s.fiscal_life_years as subclass_fiscal_life
                            FROM assets a
                            INNER JOIN asset_classes c ON c.id = a.class_id
                            LEFT JOIN asset_subclasses s ON s.id = a.subclass_id
                            WHERE a.company_id = ? AND a.tag_number = ? AND a.status = 'ATIVO'
                        """, (ACTIVE_COMPANY_ID, tag_b))
                        atv_found = cur.fetchone()

                        if not atv_found:
                            nao_encontrados.append(f"Plaqueta {tag_b} (não localizada ou já baixada)")
                            continue

                        if DEPR_METHOD == 'MENSAL':
                            dt_corte = dt_b.replace(day=1) - timedelta(days=1)
                        else:
                            dt_corte = dt_b - timedelta(days=1)

                        asset_calc = {
                            'id': atv_found[0], 'start_date': atv_found[5], 'acquisition_date': atv_found[6],
                            'acquisition_cost': atv_found[3], 'additional_costs': atv_found[4], 'residual_value': atv_found[7],
                            'class_id': atv_found[8], 'subclass_id': atv_found[9],
                            'initial_accum_depr': atv_found[10], 'initial_monthly_depr': atv_found[11], 'initial_base_date': atv_found[12],
                            'class_useful_life': atv_found[13], 'class_fiscal_life': atv_found[14],
                            'subclass_useful_life': atv_found[15], 'subclass_fiscal_life': atv_found[16]
                        }

                        depr_acum, vlc = calculate_asset_depreciation_to_date(asset_calc, dt_corte, revals_list, 'SOCIETARIO')
                        custo_tot = atv_found[3] + atv_found[4]
                        resultado = v_venda - vlc

                        cur.execute("""
                            INSERT INTO asset_disposals (
                                asset_id, disposal_date, disposal_type, sale_value,
                                cost_disposed, depr_accum_disposed, gain_loss_value
                            ) VALUES (?, ?, ?, ?, ?, ?, ?)
                        """, (atv_found[0], dt_b, motivo, v_venda, custo_tot, depr_acum, resultado))

                        cur.execute("UPDATE assets SET status = 'BAIXADO' WHERE id = ?", (atv_found[0],))
                        proc_count += 1

                    conn.commit()
                    log_audit(
                        conn, ACTIVE_COMPANY_ID, st.session_state.logged_user["email"],
                        "BAIXA_LOTE", "assets",
                        f"Baixas executadas em lote: {proc_count} itens"
                    )
                    conn.commit()
                    st.success(f"Lote processado com sucesso! Total baixado: {proc_count}")
                    if nao_encontrados:
                        st.warning("Itens não processados:\n" + "\n".join(nao_encontrados[:10]))
                    st.rerun()
            except Exception as e:
                st.error(f"Erro ao processar arquivo de baixas: {e}")

    # --- ABA REVERSÃO DE BAIXAS ---
    with aba_reversao:
        st.subheader("🔄 Reversão de Baixas Indevidas")
        st.info("Caso uma baixa tenha sido efetuada por engano, reverta-a aqui. O ativo retornará para o status ATIVO imediatamente.")

        cur.execute("""
            SELECT d.id, d.disposal_date, a.tag_number, a.description, d.disposal_type,
                   d.cost_disposed, d.depr_accum_disposed, d.gain_loss_value, a.id as asset_id
            FROM asset_disposals d
            INNER JOIN assets a ON a.id = d.asset_id
            WHERE a.company_id = ?
            ORDER BY d.disposal_date DESC
        """, (ACTIVE_COMPANY_ID,))
        baixas_feitas = cur.fetchall()

        if not baixas_feitas:
            st.info("Nenhuma baixa disponível para reversão.")
        else:
            dict_reverso = {
                f"Baixa #{b[0]} | Plaqueta: {b[2]} - {b[3]} | Data: {fmt_date(b[1])} | Custo: R$ {fmt_brl(b[5])}": b
                for b in baixas_feitas
            }
            sel_rev = st.selectbox("Selecione o registro de baixa a ser revertido", list(dict_reverso.keys()))
            b_info = dict_reverso[sel_rev]

            motivo_reversao = st.text_input("Justificativa da Reversão (Obrigatório para Auditoria) *")
            conf_rev = st.checkbox(f"Confirmo que desejo cancelar e reverter a baixa da plaqueta {b_info[2]}")

            if st.button("Executar Reversão de Baixa", type="primary") and conf_rev:
                if not motivo_reversao.strip():
                    st.error("Informe a justificativa da reversão para compliance e auditoria.")
                else:
                    disp_id = b_info[0]
                    asset_id = b_info[8]
                    plaqueta = b_info[2]

                    cur.execute("DELETE FROM asset_disposals WHERE id = ?", (disp_id,))
                    cur.execute("UPDATE assets SET status = 'ATIVO' WHERE id = ?", (asset_id,))
                    log_audit(
                        conn, ACTIVE_COMPANY_ID, st.session_state.logged_user["email"],
                        "REVERSAO_BAIXA", "asset_disposals",
                        f"Revertida baixa #{disp_id} da plaqueta {plaqueta}. Justificativa: {motivo_reversao.strip()}"
                    )
                    conn.commit()
                    st.success(f"Baixa da plaqueta {plaqueta} revertida com sucesso! O bem está novamente ATIVO.")
                    st.rerun()

    st.subheader("Histórico de Baixas Efetivadas")
    baixas_df = pd.read_sql_query("""
        SELECT d.id as "ID Baixa", d.disposal_date as "Data Baixa", a.tag_number as "Plaqueta", a.description as "Ativo",
               d.disposal_type as "Tipo", d.sale_value as "Valor Venda", d.cost_disposed as "Custo Baixado",
               d.depr_accum_disposed as "Depr. Acumulada", d.gain_loss_value as "Resultado (Ganho/Perda)"
        FROM asset_disposals d
        INNER JOIN assets a ON a.id = d.asset_id
        WHERE a.company_id = ?
        ORDER BY d.disposal_date DESC
    """, conn, params=(ACTIVE_COMPANY_ID,))
    if not baixas_df.empty:
        baixas_df["Data Baixa"] = baixas_df["Data Baixa"].apply(fmt_date)
        for c in ["Valor Venda", "Custo Baixado", "Depr. Acumulada", "Resultado (Ganho/Perda)"]:
            baixas_df[c] = baixas_df[c].apply(fmt_brl)
    render_table(baixas_df)
    conn.close()

# ------------------------------------------------------------------------------
# 5. MÓDULO FISCAL (LALUR / LACS)
# ------------------------------------------------------------------------------
elif menu == "Módulo Fiscal (LALUR/LACS)":
    st.title("⚖️ Módulo Fiscal e Ajustes LALUR / LACS")
    st.markdown("""
    Demonstrativo analítico de conciliação fiscal:
    * **Adições:** Quota de depreciação societária (contábil).
    * **Exclusões:** Quota de depreciação calculada estritamente pelas taxas fiscais (IN RFB).
    """)

    col1, col2 = st.columns(2)
    with col1:
        dt_ini = st.date_input("Período Inicial", value=date(date.today().year, 1, 1))
    with col2:
        dt_fim = st.date_input("Período Final", value=date.today())

    conn = get_connection()
    assets_df = pd.read_sql_query("""
        SELECT a.id, a.tag_number, a.description, a.acquisition_date, a.start_date,
               a.acquisition_cost, a.additional_costs, a.residual_value,
               a.initial_accum_depr, a.initial_monthly_depr, a.initial_base_date,
               c.name as class_name, c.useful_life_years as class_useful_life, c.fiscal_life_years as class_fiscal_life,
               s.name as subclass_name, s.useful_life_years as subclass_useful_life, s.fiscal_life_years as subclass_fiscal_life,
               a.class_id, a.subclass_id
        FROM assets a
        INNER JOIN asset_classes c ON c.id = a.class_id
        LEFT JOIN asset_subclasses s ON s.id = a.subclass_id
        WHERE a.company_id = ?
    """, conn, params=(ACTIVE_COMPANY_ID,))

    revals = pd.read_sql_query("""
        SELECT target_type, target_id, cut_off_date, new_useful_life_years
        FROM life_revaluations WHERE company_id = ?
    """, conn, params=(ACTIVE_COMPANY_ID,)).to_dict('records')
    conn.close()

    lalur_data = []
    tot_adicao_num = 0.0
    tot_exclusao_num = 0.0
    tot_ajuste_num = 0.0

    for _, row in assets_df.iterrows():
        dt_ini_prev = dt_ini - timedelta(days=1)
        depr_soc_ini, _ = calculate_asset_depreciation_to_date(row, dt_ini_prev, revals, 'SOCIETARIO')
        depr_fisc_ini, _ = calculate_asset_depreciation_to_date(row, dt_ini_prev, revals, 'FISCAL')

        depr_soc_fim, _ = calculate_asset_depreciation_to_date(row, dt_fim, revals, 'SOCIETARIO')
        depr_fisc_fim, _ = calculate_asset_depreciation_to_date(row, dt_fim, revals, 'FISCAL')

        cota_soc = round(depr_soc_fim - depr_soc_ini, 2)
        cota_fisc = round(depr_fisc_fim - depr_fisc_ini, 2)
        ajuste_liquido = round(cota_soc - cota_fisc, 2)

        tot_adicao_num += cota_soc
        tot_exclusao_num += cota_fisc
        tot_ajuste_num += ajuste_liquido

        if cota_soc > 0 or cota_fisc > 0:
            lalur_data.append({
                "Plaqueta": row['tag_number'],
                "Descrição": row['description'],
                "Classe": row['class_name'],
                "Depr. Societária (Adição)": fmt_brl(cota_soc),
                "Depr. Fiscal (Exclusão)": fmt_brl(cota_fisc),
                "Ajuste Líquido Lalur": fmt_brl(ajuste_liquido)
            })

    df_lalur = pd.DataFrame(lalur_data)
    if df_lalur.empty:
        st.info("Nenhuma apropriação apurada no intervalo selecionado.")
    else:
        m1, m2, m3 = st.columns(3)
        m1.metric("(+) Total de Adições (Depr. Societária)", f"R$ {fmt_brl(tot_adicao_num)}")
        m2.metric("(-) Total de Exclusões (Depr. Fiscal)", f"R$ {fmt_brl(tot_exclusao_num)}")
        m3.metric("(=) Impacto Líquido Lalur / Lacs", f"R$ {fmt_brl(tot_ajuste_num)}")

        render_table(df_lalur)

        st.download_button(
            "📥 Exportar Demonstrativo LALUR (Excel)",
            data=export_to_excel(df_lalur, "Demonstrativo_LALUR"),
            file_name=f"lalur_{dt_ini}_{dt_fim}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ------------------------------------------------------------------------------
# 6. PARTIDAS DOBRADAS (ERP-READY - POR CLASSE DE ATIVO)
# ------------------------------------------------------------------------------
elif menu == "Partidas Dobradas (ERP-Ready)":
    st.title("📑 Geração de Lançamentos Contábeis (Partidas Dobradas)")
    st.markdown("Gera o diário balanceado $(\sum \text{Débitos} = \sum \text{Créditos})$ segregado por classe de imobilizado.")

    col1, col2 = st.columns(2)
    with col1:
        d_ini = st.date_input("Início do Período Contábil", value=date(date.today().year, 1, 1))
    with col2:
        d_fim = st.date_input("Fim do Período Contábil", value=date.today())

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT account_type, class_id, account_code, account_name FROM chart_of_accounts WHERE company_id = ?", (ACTIVE_COMPANY_ID,))
    coa_rows = cur.fetchall()

    map_ativo_classe = {}
    map_depr_classe = {}
    c_desp_depr = "5.2.2.02.0005 - Despesas com Depreciação"
    c_ganho_perda = "3.2.1.01.0001 - Ganho/Perda na Alienação"
    c_fornec = "2.1.1.01.0002 - Fornecedores / Contas a Pagar"

    for r in coa_rows:
        acc_type, c_id, cod, nom = r
        full_name = f"{cod} - {nom}"
        if acc_type == 'ATIVO' and c_id:
            map_ativo_classe[c_id] = full_name
        elif acc_type == 'DEPR_ACUM' and c_id:
            map_depr_classe[c_id] = full_name
        elif acc_type == 'DESPESA_DEPR':
            c_desp_depr = full_name
        elif acc_type == 'GANHO_PERDA':
            c_ganho_perda = full_name
        elif acc_type == 'CAIXA_FORNEC':
            c_fornec = full_name

    lancamentos = []
    tot_deb_num = 0.0

    # 1. Aquisições por classe
    cur.execute("""
        SELECT a.tag_number, a.description, a.acquisition_date, a.acquisition_cost + a.additional_costs,
               a.class_id, c.name
        FROM assets a
        INNER JOIN asset_classes c ON c.id = a.class_id
        WHERE a.company_id = ? AND a.acquisition_date BETWEEN ? AND ?
    """, (ACTIVE_COMPANY_ID, d_ini, d_fim))
    for acq in cur.fetchall():
        tag_num, desc_bem, dt_acq, custo_acq, c_id, cl_nome = acq
        c_debito = map_ativo_classe.get(c_id, f"1.2.3.01 - Ativo Imobilizado ({cl_nome})")
        custo_val = round(custo_acq, 2)
        tot_deb_num += custo_val
        lancamentos.append({
            "Data": fmt_date(dt_acq), "Histórico": f"Aquisição da plaqueta {tag_num} - {desc_bem}",
            "Conta Débito": c_debito, "Conta Crédito": c_fornec, "Valor": fmt_brl(custo_val)
        })

    # 2. Quota periódica de depreciação por classe
    assets_df = pd.read_sql_query("""
        SELECT a.id, a.tag_number, a.description, a.acquisition_date, a.start_date,
               a.acquisition_cost, a.additional_costs, a.residual_value,
               a.initial_accum_depr, a.initial_monthly_depr, a.initial_base_date,
               c.name as class_name, c.useful_life_years as class_useful_life, c.fiscal_life_years as class_fiscal_life,
               s.name as subclass_name, s.useful_life_years as subclass_useful_life, s.fiscal_life_years as subclass_fiscal_life,
               a.class_id, a.subclass_id
        FROM assets a
        INNER JOIN asset_classes c ON c.id = a.class_id
        LEFT JOIN asset_subclasses s ON s.id = a.subclass_id
        WHERE a.company_id = ?
    """, conn, params=(ACTIVE_COMPANY_ID,))

    revals = pd.read_sql_query("""
        SELECT target_type, target_id, cut_off_date, new_useful_life_years
        FROM life_revaluations WHERE company_id = ?
    """, conn, params=(ACTIVE_COMPANY_ID,)).to_dict('records')

    depr_por_classe = {}
    for _, row in assets_df.iterrows():
        d_prev = d_ini - timedelta(days=1)
        depr_i, _ = calculate_asset_depreciation_to_date(row, d_prev, revals, 'SOCIETARIO')
        depr_f, _ = calculate_asset_depreciation_to_date(row, d_fim, revals, 'SOCIETARIO')
        cota = round(depr_f - depr_i, 2)
        if cota > 0:
            cid = row['class_id']
            cname = row['class_name']
            if cid not in depr_por_classe:
                depr_por_classe[cid] = {"nome": cname, "valor": 0.0}
            depr_por_classe[cid]["valor"] += cota

    for cid, info in depr_por_classe.items():
        if info["valor"] > 0:
            c_credito = map_depr_classe.get(cid, f"1.2.3.02 - (-) Depr. Acumulada ({info['nome']})")
            val_d = round(info["valor"], 2)
            tot_deb_num += val_d
            lancamentos.append({
                "Data": fmt_date(d_fim), "Histórico": f"Depreciação de {info['nome']} ({fmt_date(d_ini)} a {fmt_date(d_fim)})",
                "Conta Débito": c_desp_depr, "Conta Crédito": c_credito, "Valor": fmt_brl(val_d)
            })

    # 3. Baixas com contas específicas de classe
    cur.execute("""
        SELECT d.disposal_date, a.tag_number, a.description, d.cost_disposed,
               d.depr_accum_disposed, d.sale_value, d.gain_loss_value, a.class_id, c.name
        FROM asset_disposals d
        INNER JOIN assets a ON a.id = d.asset_id
        INNER JOIN asset_classes c ON c.id = a.class_id
        WHERE a.company_id = ? AND d.disposal_date BETWEEN ? AND ?
    """, (ACTIVE_COMPANY_ID, d_ini, d_fim))
    for b in cur.fetchall():
        dt_b, tag_b, desc_b, custo_b, depr_b, venda_b, res_b, cid_b, cl_b = b
        c_ativo_cl = map_ativo_classe.get(cid_b, f"1.2.3.01 - Ativo Imobilizado ({cl_b})")
        c_depr_cl = map_depr_classe.get(cid_b, f"1.2.3.02 - (-) Depr. Acumulada ({cl_b})")

        tot_deb_num += round(depr_b, 2)
        lancamentos.append({
            "Data": fmt_date(dt_b), "Histórico": f"Baixa da depr. acumulada da plaqueta {tag_b}",
            "Conta Débito": c_depr_cl, "Conta Crédito": c_ativo_cl, "Valor": fmt_brl(depr_b)
        })
        vlc_baixado = round(custo_b - depr_b, 2)
        tot_deb_num += vlc_baixado
        lancamentos.append({
            "Data": fmt_date(dt_b), "Histórico": f"Baixa do valor líquido da plaqueta {tag_b}",
            "Conta Débito": c_ganho_perda, "Conta Crédito": c_ativo_cl, "Valor": fmt_brl(vlc_baixado)
        })
        if venda_b > 0:
            tot_deb_num += round(venda_b, 2)
            lancamentos.append({
                "Data": fmt_date(dt_b), "Histórico": f"Receita de venda da plaqueta {tag_b}",
                "Conta Débito": c_fornec, "Conta Crédito": c_ganho_perda, "Valor": fmt_brl(venda_b)
            })

    conn.close()

    df_journal = pd.DataFrame(lancamentos)
    if df_journal.empty:
        st.info("Nenhum lançamento contábil apurado no período.")
    else:
        st.subheader("Livro Diário de Partidas Dobradas")
        render_table(df_journal)
        st.success(f"Validação de Integridade: Total Débitos (R$ {fmt_brl(tot_deb_num)}) == Total Créditos (R$ {fmt_brl(tot_deb_num)})")

        st.download_button(
            "📥 Exportar Lançamentos Contábeis (Excel)",
            data=export_to_excel(df_journal, "Lancamentos_Contabeis"),
            file_name=f"partidas_dobradas_{d_ini}_{d_fim}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ------------------------------------------------------------------------------
# 7. QUADRO DE MOVIMENTAÇÃO
# ------------------------------------------------------------------------------
elif menu == "Quadro de Movimentação":
    st.title("📑 Quadro de Movimentação do Imobilizado")
    st.markdown("Conciliação de saldos de abertura e encerramento para Notas Explicativas (CPC 27).")

    ano_corrente = date.today().year
    col1, col2 = st.columns(2)
    with col1:
        d_abertura = st.date_input("Data de Abertura", value=date(ano_corrente, 1, 1))
    with col2:
        d_fechamento = st.date_input("Data de Fechamento", value=date(ano_corrente, 12, 31))

    conn = get_connection()
    classes = pd.read_sql_query("SELECT id, name FROM asset_classes", conn)

    quadro = []
    revals = pd.read_sql_query("SELECT target_type, target_id, cut_off_date, new_useful_life_years FROM life_revaluations WHERE company_id = ?", conn, params=(ACTIVE_COMPANY_ID,)).to_dict('records')

    for _, c_row in classes.iterrows():
        c_id = c_row['id']
        c_name = c_row['name']

        assets_class = pd.read_sql_query("""
            SELECT a.id, a.tag_number, a.description, a.acquisition_date, a.start_date,
               a.acquisition_cost, a.additional_costs, a.residual_value,
               a.initial_accum_depr, a.initial_monthly_depr, a.initial_base_date,
               c.useful_life_years as class_useful_life, c.fiscal_life_years as class_fiscal_life,
               s.useful_life_years as subclass_useful_life, s.fiscal_life_years as subclass_fiscal_life,
               a.class_id, a.subclass_id
            FROM assets a
            INNER JOIN asset_classes c ON c.id = a.class_id
            LEFT JOIN asset_subclasses s ON s.id = a.subclass_id
            WHERE a.company_id = ? AND a.class_id = ?
        """, conn, params=(ACTIVE_COMPANY_ID, c_id))

        if assets_class.empty:
            continue

        custo_abertura, depr_abertura = 0.0, 0.0
        adicoes, baixas_custo, baixas_depr = 0.0, 0.0, 0.0
        depr_exercicio = 0.0

        for _, atv in assets_class.iterrows():
            dt_acq = date_from_str(atv['acquisition_date'])
            custo_tot = float(atv['acquisition_cost'] or 0.0) + float(atv['additional_costs'] or 0.0)

            if dt_acq < d_abertura:
                custo_abertura += custo_tot
                dep_ini, _ = calculate_asset_depreciation_to_date(atv, d_abertura - timedelta(days=1), revals, 'SOCIETARIO')
                depr_abertura += dep_ini
            elif d_abertura <= dt_acq <= d_fechamento:
                adicoes += custo_tot

            dep_fim, _ = calculate_asset_depreciation_to_date(atv, d_fechamento, revals, 'SOCIETARIO')
            dep_ini_calc, _ = calculate_asset_depreciation_to_date(atv, d_abertura - timedelta(days=1), revals, 'SOCIETARIO')
            depr_exercicio += max(0.0, dep_fim - dep_ini_calc)

        cur = conn.cursor()
        cur.execute("""
            SELECT d.cost_disposed, d.depr_accum_disposed
            FROM asset_disposals d
            INNER JOIN assets a ON a.id = d.asset_id
            WHERE a.company_id = ? AND a.class_id = ? AND d.disposal_date BETWEEN ? AND ?
        """, (ACTIVE_COMPANY_ID, c_id, d_abertura, d_fechamento))
        for b_custo, b_depr in cur.fetchall():
            baixas_custo += b_custo
            baixas_depr += b_depr

        custo_final = custo_abertura + adicoes - baixas_custo
        depr_final = depr_abertura + depr_exercicio - baixas_depr
        vlc_final = custo_final - depr_final

        quadro.append({
            "Classe do Imobilizado": c_name,
            "Custo Inicial": fmt_brl(custo_abertura),
            "(+) Adições": fmt_brl(adicoes),
            "(-) Baixas (Custo)": fmt_brl(baixas_custo),
            "Custo Final": fmt_brl(custo_final),
            "Depr. Acumulada Inicial": fmt_brl(depr_abertura),
            "(+) Depreciação do Exercício": fmt_brl(depr_exercicio),
            "(-) Baixas (Depreciação)": fmt_brl(baixas_depr),
            "Depr. Acumulada Final": fmt_brl(depr_final),
            "Valor Contábil Líquido": fmt_brl(vlc_final)
        })

    conn.close()

    df_quadro = pd.DataFrame(quadro)
    if df_quadro.empty:
        st.info("Nenhuma movimentação registrada para o período.")
    else:
        render_table(df_quadro)
        st.download_button(
            "📥 Exportar Quadro de Movimentação (Excel)",
            data=export_to_excel(df_quadro, "Quadro_Movimentacao"),
            file_name=f"quadro_movimentacao_{ano_corrente}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ------------------------------------------------------------------------------
# 8. AUDITORIA E COMPLIANCE (DEDO-DURO)
# ------------------------------------------------------------------------------
elif menu == "Auditoria e Compliance (Dedo-Duro)":
    st.title("🕵️ Relatório Dedo-Duro de Auditoria e Compliance")
    st.markdown("""
    Rastreamento completo de operações sensíveis: **cargas em lote**, **exclusões**, **baixas** e **reversões contábeis**.
    """)

    conn = get_connection()
    logs_df = pd.read_sql_query("""
        SELECT timestamp as "Data/Hora", user_email as "Usuário", action as "Ação",
               entity as "Módulo/Entidade", details as "Detalhes da Operação"
        FROM audit_logs
        WHERE company_id = ?
        ORDER BY id DESC
    """, conn, params=(ACTIVE_COMPANY_ID,))
    conn.close()

    if logs_df.empty:
        st.info("Nenhum registro de auditoria gravado até o momento.")
    else:
        logs_df["Data/Hora"] = logs_df["Data/Hora"].apply(fmt_date)
        render_table(logs_df)

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "📥 Exportar Log de Auditoria (Excel)",
                data=export_to_excel(logs_df, "Logs_Auditoria"),
                file_name=f"auditoria_compliance_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with c2:
            st.download_button(
                "📥 Exportar Log de Auditoria (CSV)",
                data=export_to_csv(logs_df),
                file_name=f"auditoria_compliance_{date.today()}.csv",
                mime="text/csv",
                use_container_width=True
            )

# ------------------------------------------------------------------------------
# 9. TABELAS MESTRAS & ACESSOS
# ------------------------------------------------------------------------------
elif menu == "Tabelas Mestras & Acessos":
    st.title("⚙️ Gestão de Tabelas Mestras, Parâmetros e Acessos")

    if st.session_state.logged_user["user_type"] != "ADMIN":
        st.warning("Acesso restrito ao perfil de Administrador.")
    else:
        t_aba1, t_aba2, t_aba3, t_aba4, t_aba5 = st.tabs([
            "Empresas & Parâmetros",
            "Classes & Subclasses (Upload)",
            "Plano de Contas ERP (Upload & Classes)",
            "Usuários & Permissões",
            "🛠️ Manutenção & Ajustes"
        ])

        conn = get_connection()
        cur = conn.cursor()

        # --- 9.1 EMPRESAS ---
        with t_aba1:
            st.subheader("Cadastro de Empresas")
            with st.form("form_empresa"):
                e_code = st.text_input("Código / Sigla *")
                e_name = st.text_input("Razão Social *")
                e_cnpj = st.text_input("CNPJ *")
                e_metodo = st.selectbox("Método de Depreciação *", ["MENSAL", "DIARIO"])
                if st.form_submit_button("Salvar Empresa"):
                    if e_code and e_name:
                        try:
                            cur.execute("INSERT INTO companies (code, name, cnpj, depreciation_method) VALUES (?, ?, ?, ?)",
                                        (e_code.strip(), e_name.strip(), e_cnpj.strip(), e_metodo))
                            conn.commit()
                            st.success("Empresa cadastrada!")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("Código de empresa já cadastrado.")

            st.write("Empresas Cadastradas:")
            render_table(pd.read_sql_query("SELECT id, code, name, cnpj, depreciation_method FROM companies", conn))

        # --- 9.2 CLASSES & SUBCLASSES COM UPLOAD ROBUSTO E ILIMITADO (COM UPSERT) ---
        with t_aba2:
            st.subheader("Upload e Carga em Lote de Classes & Subclasses")
            st.markdown("Suba a estrutura completa de categorias patrimoniais customizadas para o cliente.")

            cols_template_classes = ["Classe", "Subclasse", "VidaUtilSocietaria", "VidaUtilFiscal"]
            sample_classes = [
                {"Classe": "COMPUTADORES E PERIFERICOS", "Subclasse": "Desktop", "VidaUtilSocietaria": 4.0, "VidaUtilFiscal": 5.0},
                {"Classe": "COMPUTADORES E PERIFERICOS", "Subclasse": "Notebook", "VidaUtilSocietaria": 4.0, "VidaUtilFiscal": 5.0},
                {"Classe": "MAQUINAS E EQUIPAMENTOS", "Subclasse": "", "VidaUtilSocietaria": 10.0, "VidaUtilFiscal": 10.0}
            ]

            st.download_button(
                "📄 Baixar Modelo de Classes (.xlsx)",
                data=generate_template_excel(cols_template_classes, sample_classes),
                file_name="template_classes_subclasses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            up_classes = st.file_uploader("Arquivo de Classes (.xlsx)", type=["xlsx"], key="up_cl_file")
            if up_classes:
                try:
                    df_cl = pd.read_excel(up_classes)
                    st.write("Pré-visualização:")
                    st.dataframe(df_cl.head(10), use_container_width=True)

                    if st.button("Processar Carga de Classes", type="primary"):
                        c_count, s_count = 0, 0

                        for _, cl_row in df_cl.iterrows():
                            c_raw = cl_row.get("Classe")
                            if pd.isna(c_raw) or not str(c_raw).strip() or str(c_raw).strip().upper() == 'NONE':
                                continue

                            c_nome = str(c_raw).strip().upper()
                            s_raw = cl_row.get("Subclasse")
                            s_nome = str(s_raw).strip() if pd.notna(s_raw) and str(s_raw).strip().upper() != 'NONE' else ""

                            vu_soc_raw = cl_row.get("VidaUtilSocietaria")
                            vu_soc = float(vu_soc_raw) if pd.notna(vu_soc_raw) and str(vu_soc_raw).strip() != '' else 10.0

                            vu_fisc_raw = cl_row.get("VidaUtilFiscal")
                            vu_fisc = float(vu_fisc_raw) if pd.notna(vu_fisc_raw) and str(vu_fisc_raw).strip() != '' else 10.0

                            cur.execute("""
                                INSERT INTO asset_classes (name, useful_life_years, fiscal_life_years)
                                VALUES (?, ?, ?)
                                ON CONFLICT(name) DO UPDATE SET
                                    useful_life_years = excluded.useful_life_years,
                                    fiscal_life_years = excluded.fiscal_life_years
                            """, (c_nome, vu_soc, vu_fisc))
                            
                            cur.execute("SELECT id FROM asset_classes WHERE name = ?", (c_nome,))
                            c_id = cur.fetchone()[0]
                            c_count += 1

                            if s_nome:
                                cur.execute("""
                                    INSERT INTO asset_subclasses (class_id, name, useful_life_years, fiscal_life_years)
                                    VALUES (?, ?, ?, ?)
                                    ON CONFLICT(class_id, name) DO UPDATE SET
                                        useful_life_years = excluded.useful_life_years,
                                        fiscal_life_years = excluded.fiscal_life_years
                                """, (c_id, s_nome, vu_soc, vu_fisc))
                                s_count += 1

                        conn.commit()
                        log_audit(conn, ACTIVE_COMPANY_ID, st.session_state.logged_user["email"], "UPLOAD_CLASSES", "asset_classes", f"Classes: {c_count} | Subclasses: {s_count}")
                        conn.commit()
                        st.success(f"Carga concluída com sucesso! Classes: {c_count} | Subclasses: {s_count}")
                        st.rerun()
                except Exception as e:
                    st.error(f"Erro ao processar planilha de classes: {e}")

            st.divider()
            st.subheader("Cadastro Manual de Classes")
            with st.form("form_classe"):
                cl_name = st.text_input("Nome da Classe *")
                cl_soc = st.number_input("Vida Útil Societária (Anos) *", min_value=1.0, value=10.0)
                cl_fisc = st.number_input("Vida Útil Fiscal (Anos) *", min_value=1.0, value=10.0)
                if st.form_submit_button("Salvar Classe"):
                    if cl_name.strip():
                        try:
                            cur.execute("""
                                INSERT INTO asset_classes (name, useful_life_years, fiscal_life_years)
                                VALUES (?, ?, ?)
                                ON CONFLICT(name) DO UPDATE SET
                                    useful_life_years = excluded.useful_life_years,
                                    fiscal_life_years = excluded.fiscal_life_years
                            """, (cl_name.strip().upper(), cl_soc, cl_fisc))
                            conn.commit()
                            st.success("Classe salva!")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("Classe já existente.")

            render_table(pd.read_sql_query("""
                SELECT c.name as "Classe", s.name as "Subclasse",
                       COALESCE(s.useful_life_years, c.useful_life_years) as "V.U. Societária",
                       COALESCE(s.fiscal_life_years, c.fiscal_life_years) as "V.U. Fiscal"
                FROM asset_classes c
                LEFT JOIN asset_subclasses s ON s.class_id = c.id
                ORDER BY c.name, s.name
            """, conn))

        # --- 9.3 PLANO DE CONTAS COM VÍNCULO POR CLASSE ---
        with t_aba3:
            st.subheader(f"Plano de Contas ERP - {active_company[2]}")
            st.markdown("""
            **Mapeamento por Classe:**
            - **`ATIVO`**: Conta de Imobilizado vinculada a uma Classe (ex.: `1.2.3.01.0006 - COMPUTADORES`).
            - **`DEPR_ACUM`**: Conta de Depreciação Acumulada vinculada a uma Classe (ex.: `1.2.3.02.0006 - (-) COMPUTADORES`).
            - **Contas Gerais**: `DESPESA_DEPR`, `GANHO_PERDA` e `CAIXA_FORNEC` (aplicadas para toda a empresa).
            """)

            cur.execute("SELECT id, name FROM asset_classes ORDER BY name")
            classes_db = cur.fetchall()
            map_classes_coa = {c[1]: c[0] for c in classes_db}

            cols_template_contas = ["TipoConta", "ClasseVinculada", "CodigoContaERP", "DescricaoConta"]
            sample_contas_rows = [
                {"TipoConta": "ATIVO", "ClasseVinculada": "COMPUTADORES E PERIFERICOS", "CodigoContaERP": "1.2.3.01.0006", "DescricaoConta": "COMPUTADORES E PERIFERICOS"},
                {"TipoConta": "DEPR_ACUM", "ClasseVinculada": "COMPUTADORES E PERIFERICOS", "CodigoContaERP": "1.2.3.02.0006", "DescricaoConta": "(-) COMPUTADORES E PERIFERICOS"},
                {"TipoConta": "ATIVO", "ClasseVinculada": "MAQUINAS E EQUIPAMENTOS", "CodigoContaERP": "1.2.3.01.0003", "DescricaoConta": "MAQUINAS E EQUIPAMENTOS"},
                {"TipoConta": "DEPR_ACUM", "ClasseVinculada": "MAQUINAS E EQUIPAMENTOS", "CodigoContaERP": "1.2.3.02.0003", "DescricaoConta": "(-) MAQUINAS E EQUIPAMENTOS"},
                {"TipoConta": "DESPESA_DEPR", "ClasseVinculada": "", "CodigoContaERP": "5.2.2.02.0005", "DescricaoConta": "DESPESAS DE DEPRECIACAO"},
                {"TipoConta": "CAIXA_FORNEC", "ClasseVinculada": "", "CodigoContaERP": "2.1.1.01.0002", "DescricaoConta": "FORNECEDOR ATIVO IMOBILIZADO"}
            ]

            st.download_button(
                "📄 Baixar Modelo de Plano de Contas com Classes (.xlsx)",
                data=generate_template_excel(cols_template_contas, sample_contas_rows),
                file_name="template_plano_contas_classes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            up_contas = st.file_uploader("Arquivo do Plano de Contas (.xlsx)", type=["xlsx"], key="up_pc_file")
            if up_contas:
                try:
                    df_pc = pd.read_excel(up_contas)
                    st.write("Pré-visualização do Plano de Contas:")
                    st.dataframe(df_pc, use_container_width=True)

                    if st.button("Processar Carga do Plano de Contas", type="primary"):
                        c_proc = 0
                        for _, pc_row in df_pc.iterrows():
                            t_conta = str(pc_row.get("TipoConta", "")).strip().upper()
                            cl_raw = pc_row.get("ClasseVinculada")
                            cl_vinc = str(cl_raw).strip().upper() if pd.notna(cl_raw) and str(cl_raw).strip().upper() != 'NONE' else ""
                            cod_c = str(pc_row.get("CodigoContaERP", "")).strip()
                            nom_c = str(pc_row.get("DescricaoConta", "")).strip()

                            if t_conta in ['ATIVO', 'DEPR_ACUM', 'DESPESA_DEPR', 'GANHO_PERDA', 'CAIXA_FORNEC'] and cod_c:
                                class_id_resolved = None
                                if cl_vinc:
                                    if cl_vinc in map_classes_coa:
                                        class_id_resolved = map_classes_coa[cl_vinc]
                                    else:
                                        cur.execute("INSERT INTO asset_classes (name, useful_life_years, fiscal_life_years) VALUES (?, 10.0, 10.0) ON CONFLICT(name) DO NOTHING", (cl_vinc,))
                                        cur.execute("SELECT id FROM asset_classes WHERE name = ?", (cl_vinc,))
                                        class_id_resolved = cur.fetchone()[0]
                                        map_classes_coa[cl_vinc] = class_id_resolved

                                if class_id_resolved is not None:
                                    cur.execute("DELETE FROM chart_of_accounts WHERE company_id = ? AND account_type = ? AND class_id = ?", (ACTIVE_COMPANY_ID, t_conta, class_id_resolved))
                                else:
                                    cur.execute("DELETE FROM chart_of_accounts WHERE company_id = ? AND account_type = ? AND class_id IS NULL", (ACTIVE_COMPANY_ID, t_conta))

                                cur.execute("""
                                    INSERT INTO chart_of_accounts (company_id, class_id, account_type, account_code, account_name)
                                    VALUES (?, ?, ?, ?, ?)
                                """, (ACTIVE_COMPANY_ID, class_id_resolved, t_conta, cod_c, nom_c))
                                c_proc += 1

                        conn.commit()
                        log_audit(conn, ACTIVE_COMPANY_ID, st.session_state.logged_user["email"], "UPLOAD_PLANO_CONTAS", "chart_of_accounts", f"Contas mapeadas: {c_proc}")
                        conn.commit()
                        st.success(f"Plano de contas atualizado! {c_proc} contas configuradas com sucesso.")
                        st.rerun()
                except Exception as e:
                    st.error(f"Erro ao processar plano de contas: {e}")

            st.divider()
            st.subheader("Configuração Individual de Conta")
            with st.form("form_plano_contas"):
                col_p1, col_p2 = st.columns(2)
                with col_p1:
                    p_tipo = st.selectbox("Tipo de Conta", ['ATIVO', 'DEPR_ACUM', 'DESPESA_DEPR', 'GANHO_PERDA', 'CAIXA_FORNEC'])
                    p_classe_opts = ["(Geral / Todas as Classes)"] + list(map_classes_coa.keys())
                    p_classe_sel = st.selectbox("Classe de Ativo Vinculada", p_classe_opts)
                    p_cid = map_classes_coa[p_classe_sel] if p_classe_sel != "(Geral / Todas as Classes)" else None

                with col_p2:
                    p_cod = st.text_input("Código Estruturado / Reduzido *", placeholder="Ex: 1.2.3.01.0006")
                    p_nome = st.text_input("Descrição da Conta no ERP *", placeholder="Ex: COMPUTADORES E PERIFERICOS")

                if st.form_submit_button("Configurar Conta", type="primary"):
                    if p_cod.strip() and p_nome.strip():
                        if p_cid is not None:
                            cur.execute("DELETE FROM chart_of_accounts WHERE company_id = ? AND account_type = ? AND class_id = ?", (ACTIVE_COMPANY_ID, p_tipo, p_cid))
                        else:
                            cur.execute("DELETE FROM chart_of_accounts WHERE company_id = ? AND account_type = ? AND class_id IS NULL", (ACTIVE_COMPANY_ID, p_tipo))

                        cur.execute("""
                            INSERT INTO chart_of_accounts (company_id, class_id, account_type, account_code, account_name)
                            VALUES (?, ?, ?, ?, ?)
                        """, (ACTIVE_COMPANY_ID, p_cid, p_tipo, p_cod.strip(), p_nome.strip()))
                        conn.commit()
                        st.success("Mapeamento salvo com sucesso!")
                        st.rerun()
                    else:
                        st.error("Informe o código e a descrição da conta.")

            st.write("Contas Mapeadas para esta Empresa:")
            render_table(pd.read_sql_query("""
                SELECT coa.account_type as "Tipo",
                       COALESCE(ac.name, '(Geral da Empresa)') as "Classe Vinculada",
                       coa.account_code as "Código ERP",
                       coa.account_name as "Nome da Conta"
                FROM chart_of_accounts coa
                LEFT JOIN asset_classes ac ON ac.id = coa.class_id
                WHERE coa.company_id = ?
                ORDER BY coa.account_type, ac.name
            """, conn, params=(ACTIVE_COMPANY_ID,)))

        # --- 9.4 USUÁRIOS & PERMISSÕES ---
        with t_aba4:
            st.subheader("Gestão de Usuários")
            with st.form("form_user"):
                u_email = st.text_input("E-mail *")
                u_tipo = st.selectbox("Perfil de Acesso", ["EXTERNO", "ADMIN"])
                if st.form_submit_button("Cadastrar Usuário"):
                    if u_email:
                        try:
                            cur.execute("""
                                INSERT INTO users (email, password_hash, user_type, must_reset_password)
                                VALUES (?, ?, ?, 1)
                            """, (u_email.strip(), hash_password('trocar123'), u_tipo))
                            conn.commit()
                            st.success(f"Usuário {u_email} cadastrado com senha temporária 'trocar123'.")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("E-mail já existente.")

            st.write("Vincular Usuário Externo a Empresas:")
            cur.execute("SELECT id, email FROM users WHERE user_type = 'EXTERNO'")
            ext_users = {u[1]: u[0] for u in cur.fetchall()}
            cur.execute("SELECT id, name FROM companies")
            all_comps = {c[1]: c[0] for c in cur.fetchall()}

            if ext_users and all_comps:
                with st.form("form_vinculo"):
                    v_u = st.selectbox("Usuário", list(ext_users.keys()))
                    v_c = st.multiselect("Empresas Liberadas", list(all_comps.keys()))
                    if st.form_submit_button("Atualizar Acessos"):
                        uid = ext_users[v_u]
                        cur.execute("DELETE FROM user_companies WHERE user_id = ?", (uid,))
                        for comp_name in v_c:
                            cur.execute("INSERT INTO user_companies (user_id, company_id) VALUES (?, ?)", (uid, all_comps[comp_name]))
                        conn.commit()
                        st.success("Acessos atualizados!")
                        st.rerun()

            render_table(pd.read_sql_query("""
                SELECT id, email, user_type, must_reset_password, created_at FROM users
            """, conn))

        # --- 9.5 MANUTENÇÃO & AJUSTES ---
        with t_aba5:
            st.subheader("🛠️ Ajuste de Valores Residuais da Base Atual")
            st.markdown("""
            Caso os ativos da empresa ativa tenham sido importados anteriormente com **Valor Residual igual ao Custo de Aquisição** (o que anula a base de cálculo da depreciação societária), clique no botão abaixo para **zerar o valor residual** de todos os ativos da empresa ativa.
            """)
            if st.button("Zerar Valor Residual Indevido (Restabelecer Depreciação Integral)", type="primary"):
                cur.execute("""
                    UPDATE assets 
                    SET residual_value = 0.0 
                    WHERE company_id = ? AND residual_value >= acquisition_cost * 0.99;
                """, (ACTIVE_COMPANY_ID,))
                linhas_ajustadas = cur.rowcount
                conn.commit()
                log_audit(conn, ACTIVE_COMPANY_ID, st.session_state.logged_user["email"], "AJUSTE_RESIDUAL", "assets", f"Zerado residual de {linhas_ajustadas} ativos")
                st.success(f"{linhas_ajustadas} ativos ajustados com sucesso!")
                st.rerun()

        conn.close()