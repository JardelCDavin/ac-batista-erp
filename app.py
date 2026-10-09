import json
import os
import streamlit as st
import urllib.parse
from datetime import date, datetime as dt_mod, time as time_mod, timedelta
import time as time_lib
import io
import xml.etree.ElementTree as ET
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import unicodedata

# --- CONEXÃO COM O SUPABASE (POSTGRESQL NA NUVEM) ---
@st.cache_resource
def conectar_supabase():
    try:
        conn = st.connection("postgres", type="sql")
        return conn
    except Exception as e:
        st.error(f"❌ Erro ao conectar ao banco de dados Supabase: {e}")
        return None

conn_db = conectar_supabase()

def ler_dados_sql(query, params=None):
    if conn_db:
        try:
            return conn_db.query(query, params=params, ttl=0)
        except Exception as e:
            return pd.DataFrame()
    return pd.DataFrame()

def executar_sql(query, params=None):
    if conn_db:
        try:
            with conn_db.session as s:
                s.execute(query, params)
                s.commit()
                return True
        except Exception as e:
            st.error(f"Erro na execução SQL: {e}")
            return False
    return False

# --- BLINDAGEM DE API: MEMÓRIA DE CURTO PRAZO ---
@st.cache_data(ttl=300)
def buscar_dados_tabela_cache(nome_tabela):
    try:
        query = f"SELECT * FROM {nome_tabela}"
        df = ler_dados_sql(query)
        if not df.empty:
            df.columns = [str(c).strip().upper() for c in df.columns]
            return df.to_dict(orient="records")
        return []
    except Exception:
        return []

# Inicialização segura do estado da sessão
st.session_state.setdefault('logged_in', True)
st.session_state.setdefault('user_nivel', 'ADMIN')
if 'logado' not in st.session_state: st.session_state.logado = False
if 'nivel' not in st.session_state: st.session_state.nivel = "Comum"
if 'usuario' not in st.session_state: st.session_state.usuario = "Nenhum"
if 'filial_nome' not in st.session_state: st.session_state.filial_nome = ""
if 'libera_digitacao_semanal' not in st.session_state: st.session_state.libera_digitacao_semanal = True
if 'modulo_ativo' not in st.session_state: st.session_state.modulo_ativo = 'Lançamento'
df_cot_reais = pd.DataFrame()

@st.cache_data(ttl=600)
def carregar_dados_planilha():
    try:
        dados = buscar_dados_tabela_cache("media_venda_diaria")
        return dados
    except Exception:
        return []

@st.cache_data(ttl=300)
def carregar_fornecedores_ativos():
    try:
        dados = buscar_dados_tabela_cache("fornecedores")
        if dados:
            ativos = [str(row.get('FORNECEDOR', '')).strip() for row in dados if str(row.get('ATIVO', '')).strip().upper() == 'SIM']
            return ativos if ativos else ["FORNECEDOR PADRÃO"]
        return ["FORNECEDOR PADRÃO"]
    except Exception: return ["FORNECEDOR PADRÃO"]

@st.cache_data(ttl=300)
def carregar_todos_fornecedores_cadastrados():
    try:
        dados = buscar_dados_tabela_cache("fornecedores")
        if dados:
            todos = [str(row.get('FORNECEDOR', '')).strip() for row in dados if str(row.get('FORNECEDOR', '')).strip()]
            return sorted(list(set(todos))) if todos else ["FORNECEDOR PADRÃO"]
        return ["FORNECEDOR PADRÃO"]
    except Exception: return ["FORNECEDOR PADRÃO"]

@st.cache_data(ttl=300)
def carregar_catalogo_produtos_mapeamento():
    try:
        dados = buscar_dados_tabela_cache("produtos")
        cod_para_prod = {}
        prod_para_cod = {}
        precos_base_map = {}
        lista_nomes = []
        if dados:
            for row in dados:
                p_nome = str(row.get('PRODUTO', '')).strip().upper()
                p_cod = str(row.get('CODIGO', '')).strip()
                p_preco = row.get('PRECO_BASE', 0.0)
                
                try: p_preco_float = float(str(p_preco).replace(',', '.')) if p_preco else 0.0
                except: p_preco_float = 0.0

                if p_nome:
                    lista_nomes.append(p_nome)
                    if p_cod:
                        cod_para_prod[p_cod] = p_nome
                        prod_para_cod[p_nome] = p_cod
                    if p_preco_float > 0:
                        precos_base_map[p_nome] = p_preco_float
                        
        return cod_para_prod, prod_para_cod, sorted(list(set(lista_nomes))), precos_base_map
    except Exception: return {}, {}, ["COXA SOLTEIRA PILÃO KG", "LINGUIÇA SUÍNA CHURRASCO", "SASSAMI KG"], {}

@st.cache_data(ttl=60)
def carregar_de_para_fornecedores():
    dados = buscar_dados_tabela_cache("de_para_fornecedores")
    mapeamento = {}
    if dados:
        for row in dados:
            cnpj = str(row.get("CNPJ_FORNECEDOR", "")).strip().replace(".", "").replace("/", "").replace("-", "").replace("'", "")
            if cnpj.isdigit(): cnpj = cnpj.zfill(14)
            cod_forn = str(row.get("CODIGO_ITEM_FORNECEDOR", "")).strip().replace("'", "")
            cod_interno = str(row.get("CODIGO_INTERNO", "")).strip().replace("'", "")
            nome_interno = str(row.get("PRODUTO_INTERNO", "")).strip()
            
            if cnpj and cod_forn:
                chave = f"{cnpj}_{cod_forn}"
                mapeamento[chave] = {"CODIGO_INTERNO": cod_interno, "PRODUTO_INTERNO": nome_interno}
    return mapeamento

@st.cache_data(ttl=300)
def carregar_dados_filiais_dict():
    dados = buscar_dados_tabela_cache("filiais")
    filiais_map = {}
    if not dados:
        return {
            "TEJUCO": {
                "RAZAO": "AC BATISTA ALIMENTAÇÃO - TEJUCO",
                "CNPJ": "06.121.429/0008-90",
                "IE": "625274795.07.43",
                "ENDERECO": "AV. GENERAL OSORIO, 255 - SÃO JOÃO DEL REI/MG",
                "CEP": "36300-168",
                "EMAIL": "comprasacbatista@gmail.com"
            }
        }
    for row in dados:
        status = str(row.get("STATUS", "")).strip().upper()
        if status == "ATIVO":
            texto_empresa = str(row.get("DADOS_EMPRESA", ""))
            linhas = [l.strip() for l in texto_empresa.split("\n") if l.strip()]
            razao = linhas[0] if len(linhas) > 0 else "AC BATISTA ALIMENTAÇÃO"
            
            nome_filial = "TEJUCO"
            for f in ["TEJUCO", "CENTRO", "MATOSINHOS", "COLONIA", "BARBACENA", "LEOPOLDINA", "DONA MARIA", "RM SABOR"]:
                if f in razao.upper():
                    nome_filial = f
                    break
            
            cnpj, ie, endereco, cep, email = "", "", "", "", ""
            for linha in linhas:
                l_up = linha.upper()
                if "CNPJ:" in l_up: cnpj = linha.split(":")[-1].strip()
                elif "ESTADUAL:" in l_up or "INS. ESTADUAL:" in l_up: ie = linha.split(":")[-1].strip()
                elif "CEP:" in l_up: cep = linha.split(":")[-1].strip()
                elif "E-MAIL:" in l_up or "EMAIL:" in l_up: email = linha.split(":")[-1].strip()
                elif not any(x in l_up for x in ["CNPJ:", "ESTADUAL:", "CEP:", "E-MAIL:", "EMAIL:"]) and linha != razao:
                    if not endereco: endereco = linha
                    else: endereco += " - " + linha
                        
            filial_key = nome_filial.upper()
            filiais_map[filial_key] = {
                "RAZAO": razao, "CNPJ": cnpj if cnpj else "06.121.429/0008-90",
                "IE": ie if ie else "625274795.07.43", "ENDERECO": endereco if endereco else "SÃO JOÃO DEL REI/MG",
                "CEP": cep if cep else "36300-168", "EMAIL": email if email else "comprasacbatista@gmail.com"
            }
    return filiais_map

MOCK_FILIAIS = ["TEJUCO", "CENTRO", "MATOSINHOS", "RM SABOR", "COLONIA", "BARBACENA", "LEOPOLDINA"]

@st.cache_data(ttl=600)
def carregar_proteinas_semanal():
    try:
        dados = buscar_dados_tabela_cache("produtos")
        if dados:
            df = pd.DataFrame(dados)
            if 'PRODUTO' in df.columns:
                return sorted(df['PRODUTO'].dropna().astype(str).str.strip().unique().tolist())
    except Exception: pass
    return []

def normalizar_nome_produto(nome):
    n = str(nome).upper().strip()
    n = unicodedata.normalize('NFKD', n).encode('ASCII', 'ignore').decode('utf-8')
    for prep in [" DE ", " DA ", " DO ", " COM "]:
        n = n.replace(prep, " ")
    remover = [" CONGELADO", " CONGELADA", " RESFRIADO", " RESFRIADA", " IN NATURA", " KG", " KGS", " UNID", " UN"]
    for r in remover:
        n = n.replace(r, "")
    return " ".join(n.split()).strip()

st.set_page_config(page_title="Portal AC Batista", layout="wide")

def validar_usuario_supabase(usuario, senha):
    if conn_db is None: return None, "❌ Erro: Não foi possível conectar ao banco de dados Supabase."
    u_in = str(usuario).strip().upper()
    s_in = str(senha).strip()
    try:
        query = "SELECT * FROM bd_usuarios WHERE UPPER(TRIM(usuario)) = :user"
        df_user = ler_dados_sql(query, {"user": u_in})
        if not df_user.empty:
            linha = df_user.iloc[0]
            senha_tabela = str(linha.get('senha', '')).strip()
            if senha_tabela == s_in:
                return {
                    'USUARIO': str(linha.get('usuario', '')).strip(),
                    'FILIAL': str(linha.get('filial', '')).strip().upper(),
                    'NIVEL': str(linha.get('nivel', '')).strip().upper() 
                }, None
            else: return None, "❌ Palavra-passe incorreta. Tenta novamente."
    except Exception as e: return None, f"❌ Erro crítico ao conectar à base de dados: {e}"
    return None, "❌ Utilizador não localizado."

# =========================================================================
# 🛑 TELA DE LOGIN ISOLADA (ESTÁVEL PARA TELEMÓVEL E PC)
# =========================================================================
container_login = st.empty()

if not st.session_state.get('logado', False):
    with container_login.container():
        st.markdown("""
            <style>
                [data-test-id="stSidebar"] { display: none !important; }
                .stMainBlockContainer { 
                    max-width: 450px; 
                    margin: 0 auto; 
                    padding-top: 4rem; 
                }
                .login-title {
                    color: #004A99;
                    font-weight: 700;
                    text-align: center;
                    margin-bottom: 0.2rem;
                    font-size: 1.8rem;
                }
                .login-subtitle {
                    color: #666666;
                    text-align: center;
                    margin-bottom: 2rem;
                    font-size: 0.95rem;
                }
            </style>
        """, unsafe_allow_html=True)
        
        st.markdown("<h1 class='login-title'>🔒 AC Batista ERP</h1>", unsafe_allow_html=True)
        st.markdown("<p class='login-subtitle'>Portal de Gestão e Suprimentos</p>", unsafe_allow_html=True)
        
        with st.form("form_login_ac_batista"):
            usuario = st.text_input("Usuário", key="txt_usuario_final", placeholder="Digite o seu usuário")
            senha = st.text_input("Senha", type="password", key="txt_senha_final", placeholder="Digite a sua senha")
            
            st.markdown("<div style='margin-top: 1rem;'></div>", unsafe_allow_html=True)
            botao_submeter = st.form_submit_button("Acessar", use_container_width=True)

            if botao_submeter:
                if not usuario or not senha:
                    st.error("⚠️ Por favor, preencha o usuário e a senha.")
                else:
                    registro, erro = validar_usuario_supabase(usuario, senha)
                    if erro:
                        st.error(erro)
                    else:
                        container_login.empty()
                        st.session_state.logado = True
                        st.session_state.usuario = str(registro.get('USUARIO', '')).strip()
                        nome_empresa = str(registro.get('FILIAL', '')).strip().upper()
                        nivel_detectado = str(registro.get('NIVEL', '')).strip().upper()
                        
                        if nivel_detectado == "FORNECEDOR":
                            st.session_state.nivel = "Fornecedor"
                            st.session_state.filial_nome = nome_empresa
                        elif nome_empresa == "ADMINISTRATIVO" or nivel_detectado == "ADMIN":
                            st.session_state.nivel = "Admin"
                            st.session_state.filial_nome = "ADMINISTRATIVO"
                        else:
                            st.session_state.nivel = "Nutricionista"
                            st.session_state.filial_nome = nome_empresa
                        st.cache_data.clear()
                        st.rerun()

st.markdown("""<style>.stButton>button { background-color: #004A99; color: white; width: 100%; border-radius: 5px; height: 3em; font-weight: bold; } .stButton>button:hover { background-color: #003366; color: white; }</style>""", unsafe_allow_html=True)

@st.cache_data(ttl=60)
def ler_status_digitacao():
    try:
        df = ler_dados_sql("SELECT valor FROM configuracoes WHERE UPPER(TRIM(chave)) = 'STATUS_DIGITACAO'")
        if not df.empty:
            val = str(df.iloc[0].get('valor', '')).strip().upper()
            return True if val == "LIGADO" else False
    except Exception:
        pass
    return False

def salvar_governanca(ligado):
    novo_valor = "LIGADO" if ligado else "DESLIGADO"
    try:
        query_check = "SELECT * FROM configuracoes WHERE UPPER(TRIM(chave)) = 'STATUS_DIGITACAO'"
        df = ler_dados_sql(query_check)
        if not df.empty:
            executar_sql("UPDATE configuracoes SET valor = :val WHERE UPPER(TRIM(chave)) = 'STATUS_DIGITACAO'", {"val": novo_valor})
        else:
            executar_sql("INSERT INTO configuracoes (chave, valor) VALUES ('STATUS_DIGITACAO', :val)", {"val": novo_valor})
        st.session_state['libera_digitacao_semanal'] = ligado
    except Exception as e:
        st.error(f"Erro ao salvar governança: {e}")

def encurtar_nome_fornecedor(nome_completo):
    n = str(nome_completo).strip().upper()
    if "LIDER" in n or "RUBBO" in n: return "Líder"
    elif "OESA" in n: return "Oesa"
    elif "RIO BRANCO" in n or "PIF PAF" in n or "RIO" in n: return "Rio Branco"
    elif "SEARA" in n: return "Seara"
    elif "FRIGO" in n: return "Frigo"
    partes = n.split()
    return partes[0].title() if partes else "Fornecedor"

def tratar_preco_float(valor): 
    try: 
        if isinstance(valor, (int, float)): 
            val = float(valor)
            if val >= 100 and val % 1 == 0 and val not in [100, 200, 500, 1000]: 
                return val / 100.0
            elif val >= 40 and val < 100 and val % 1 == 0: 
                return val / 10.0
            return val
        
        v_str = str(valor).replace("R$", "").strip()
        if not v_str or v_str.lower() == 'nan': return 0.0
        
        if "," in v_str and "." in v_str:
            if v_str.rfind(",") > v_str.rfind("."): 
                v_str = v_str.replace(".", "").replace(",", ".")
            else: 
                v_str = v_str.replace(",", "")
        elif "," in v_str and "." not in v_str: 
            v_str = v_str.replace(",", ".")
            
        num = float(v_str)
        if num >= 100 and num % 1 == 0 and num not in [100, 200, 500, 1000]:
            return num / 100.0
        elif num >= 40 and num < 100 and num % 1 == 0: 
            return num / 10.0
        return num
    except: return 0.0

def tratar_qtd_float(valor):
    try:
        if isinstance(valor, (int, float)):
            return float(valor)
        v_str = str(valor).replace("R$", "").replace("kg", "").replace("KG", "").strip()
        if not v_str or v_str.lower() == 'nan': return 0.0
        if "," in v_str and "." in v_str:
            if v_str.rfind(",") > v_str.rfind("."): 
                v_str = v_str.replace(".", "").replace(",", ".")
            else: 
                v_str = v_str.replace(",", "")
        elif "," in v_str and "." not in v_str: 
            v_str = v_str.replace(",", ".")
        return float(v_str)
    except: return 0.0

def tratar_peso_float(valor):
    try:
        if isinstance(valor, (int, float)):
            return float(valor)
        v_str = str(valor).replace("R$", "").strip()
        if not v_str or v_str.lower() == 'nan': return 1.0
        if "," in v_str and "." in v_str:
            if v_str.rfind(",") > v_str.rfind("."): 
                v_str = v_str.replace(".", "").replace(",", ".")
            else: 
                v_str = v_str.replace(",", "")
        elif "," in v_str and "." not in v_str: 
            v_str = v_str.replace(",", ".")
        return float(v_str)
    except: return 1.0

def gerar_pdf_pedido(num_pedido, filial_nome, dados_filial, forn_alvo, telefone_forn, prazo_pgto, df_itens, data_entrega):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    styles = getSampleStyleSheet()
    
    estilo_titulo = ParagraphStyle('Titulo', parent=styles['Heading1'], fontSize=15, textColor=colors.HexColor('#004A99'), alignment=1, spaceAfter=4)
    estilo_sub = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor('#555555'), alignment=1, spaceAfter=12)
    estilo_corpo = ParagraphStyle('Corpo', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor('#222222'), spaceAfter=3)
    estilo_aviso_box = ParagraphStyle('AvisoBox', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#900C3F'), fontName='Helvetica-Bold', alignment=1, leading=12)
    
    story.append(Paragraph("<b>👨‍🍳 AC BATISTA ALIMENTAÇÃO</b>", estilo_titulo))
    story.append(Paragraph(f"<b>{dados_filial['RAZAO']}</b><br/>CNPJ Faturamento: {dados_filial['CNPJ']} | Inscrição Estadual: {dados_filial['IE']}<br/>{dados_filial['ENDERECO']} - CEP: {dados_filial['CEP']}<br/>E-mail: {dados_filial['EMAIL']}", estilo_sub))
    
    info_data = [
        [
            Paragraph(f"<b>Fornecedor:</b> {forn_alvo}<br/><b>Telefone Contato:</b> {telefone_forn}<br/><b>Prazo de Pagamento:</b> {prazo_pgto}", estilo_corpo),
            Paragraph(f"<b>Pedido Nº:</b> {num_pedido}<br/><b>Data e Hora:</b> {dt_mod.now().strftime('%d/%m/%Y %H:%M')}<br/><b>Previsão de Entrega:</b> {data_entrega.strftime('%d/%m/%Y')}", estilo_corpo)
        ]
    ]
    t_info = Table(info_data, colWidths=[270, 270])
    t_info.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f8f9fa')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#cccccc')),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_info)
    story.append(Spacer(1, 10))
    
    tabela_conteudo = [["Código", "Descrição do Produto", "Preço Unit.", "Quantidade", "Preço Total (R$)"]]
    for _, row in df_itens.iterrows():
        tabela_conteudo.append([
            str(row["Código"]),
            str(row["Produto"]),
            f"R$ {row['Preço Unit.']:,.2f}",
            f"{row['Qtd']:,.2f}",
            f"R$ {row['Total']:,.2f}"
        ])
    
    val_total = df_itens["Total"].sum()
    tabela_conteudo.append(["", "", "", "PREÇO TOTAL:", f"R$ {val_total:,.2f}"])
    
    t_itens = Table(tabela_conteudo, colWidths=[55, 245, 75, 70, 95])
    t_itens.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#004A99')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 6),
        ('BACKGROUND', (0,1), (-1,-2), colors.HexColor('#ffffff')),
        ('GRID', (0,0), (-1,-2), 0.5, colors.HexColor('#dddddd')),
        ('ALIGN', (2,1), (-1,-1), 'RIGHT'),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#fff3cd')),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('LINEABOVE', (0,-1), (-1,-1), 1, colors.HexColor('#004A99')),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_itens)
    story.append(Spacer(1, 12))
    
    aviso_texto = "<b>⚠️ ATENÇÃO: INFORMAÇÃO DE ENTREGA OBRIGATÓRIA ⚠️</b><br/>" \
                  "ENTREGAS DE MERCADORIAS APENAS NOS HORÁRIOS DE SEGUNDA A SEXTA<br/>" \
                  "<b>MANHÃ: 07:00 AS 10:30hrs</b> &nbsp;|&nbsp; <b>TARDE: 14:00 AS 15:00hrs</b>"
    
    t_aviso = Table([[Paragraph(aviso_texto, estilo_aviso_box)]], colWidths=[540])
    t_aviso.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#ffcccc')),
        ('BOX', (0,0), (-1,-1), 1.5, colors.HexColor('#cc0000')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 8),
        ('ALIGN', (0,0), (-1,-1), 'CENTER')
    ]))
    story.append(t_aviso)
    
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def consolidar_proteina_semanal_geral(restaurante_filtro="Todos"):
    dados = buscar_dados_tabela_cache("auditoria_consolidada")
    if not dados:
        st.warning("A tabela de auditoria está vazia!")
        return

    df_linhas = []
    for linha in dados:
        filial_reg = str(linha.get("RESTAURANTE", "")).strip()
        if restaurante_filtro != "Todos" and filial_reg != restaurante_filtro: continue
        if str(linha.get("STATUS_VALIDACAO", "")).strip().upper() == "APROVADO": continue
        produto = str(linha.get("PRODUTO", "")).strip()
        try:
            val = linha.get("PEDIDO_NUTRICIONISTA", 0)
            qtd = tratar_qtd_float(val)
        except: qtd = 0.0
        if produto and qtd > 0:
            df_linhas.append({"Filial": filial_reg, "Proteína / Item": produto, "Quantidade Lançada (KG)": qtd, "Status": linha.get("STATUS_VALIDACAO", ""), "Justificativa da Nutricionista": linha.get("JUSTIFICATIVA", "")})

    if not df_linhas:
        st.warning(f"Nenhum pedido localizado para: {restaurante_filtro}")
        return

    df_final = pd.DataFrame(df_linhas)
    df_editado = st.data_editor(df_final, use_container_width=True, hide_index=True, disabled=["Filial", "Proteína / Item", "Status", "Justificativa da Nutricionista"])

    if st.button("💾 Guardar Correção da Filial", type="primary"):
        try:
            for _, r_ed in df_editado.iterrows():
                f_v, p_v, q_v = str(r_ed["Filial"]).strip().upper(), str(r_ed["Proteína / Item"]).strip().upper(), tratar_qtd_float(r_ed["Quantidade Lançada (KG)"])
                executar_sql(
                    "UPDATE auditoria_consolidada SET pedido_nutricionista = :qtd, status_validacao = 'APROVADO' WHERE UPPER(TRIM(restaurante)) = :filial AND UPPER(TRIM(produto)) = :prod",
                    {"qtd": f"{q_v:.3f}", "filial": f_v, "prod": p_v}
                )
            st.success("✅ Guardado com sucesso!")
            st.cache_data.clear(); time_lib.sleep(1); st.rerun()
        except Exception as e: st.error(f"Erro: {e}")
        def modulo_cotacao_consolidacao(): 
         st.title("📊 Cotação & Consolidação") 
    st.markdown("## ⚙ Painel de Distribuição de Suprimentos") 
    st.info("Espaço destinado ao gerenciamento logístico de insumos e fechamento de cargas do Diretor Jardel.") 
    
    with st.expander("⏱️ Controlo de Prazo e Acompanhamento de Fornecedores", expanded=True):
        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1: data_limite = st.date_input("Data Limite de Cotação", value=date.today(), key="dt_limite_diretor")
        with col_p2: hora_limite = st.time_input("Horário Limite", value=time_mod(9, 0, 0), key="hr_limite_diretor")
        with col_p3:
            st.write("")
            st.write("")
            if st.button("⏰ Prorrogar Prazo (+2 Horas)", use_container_width=True, key="btn_prorrogar_prazo"):
                st.success("✅ Prazo prorrogado com sucesso para os fornecedores!")
        
        st.markdown("---")
        st.markdown("#### 📋 Status de Envio dos Fornecedores")
        try:
            fornecedores_cadastrados = set(carregar_fornecedores_ativos())
            dados_cot_atuais = buscar_dados_tabela_cache("bd_cotacao")
            fornecedores_que_enviaram = {str(c.get("FORNECEDOR", "")).strip().upper() for c in dados_cot_atuais if c.get("FORNECEDOR")} if dados_cot_atuais else set()
            todos_forn = sorted(list(fornecedores_cadastrados.union(fornecedores_que_enviaram))) or ["FORNECEDOR PADRÃO"]
            
            status_lista = []
            for forn in todos_forn:
                nome_curto = encurtar_nome_fornecedor(forn)
                if str(forn).strip().upper() in fornecedores_que_enviaram:
                    status_lista.append({"Fornecedor": nome_curto, "Status": "🟢 Finalizado / Enviado", "Detalhes": "Proposta registada"})
                else:
                    status_lista.append({"Fornecedor": nome_curto, "Status": "🔴 Pendente", "Detalhes": "A aguardar envio"})
            st.dataframe(pd.DataFrame(status_lista), use_container_width=True, hide_index=True)
        except Exception as e_st:
            st.caption(f"ℹ️ A aguardar dados para exibir o status: {e_st}")

    st.markdown("---")

    try: 
        dados_cotacao_brutos = buscar_dados_tabela_cache("bd_cotacao") 
        if not dados_cotacao_brutos: 
            st.info("ℹ️ Nenhuma proposta de cotação encontrada na tabela bd_cotacao para o lote atual.")
            df_bruto = pd.DataFrame()
        else:
            df_bruto = pd.DataFrame(dados_cotacao_brutos) 
            df_bruto.columns = [str(c).strip().upper() for c in df_bruto.columns] 

        if "PRECO_PACOTE" in df_bruto.columns: df_bruto["PRECO_PACOTE"] = df_bruto["PRECO_PACOTE"].apply(tratar_preco_float)
        if "PESO_EMBALAGEM" in df_bruto.columns: df_bruto["PESO_EMBALAGEM"] = df_bruto["PESO_EMBALAGEM"].apply(tratar_peso_float)

        if "PRECO_PACOTE" in df_bruto.columns and "PESO_EMBALAGEM" in df_bruto.columns:
            df_bruto["PRECO_KG_EQUIV"] = df_bruto.apply(lambda row: round(row["PRECO_PACOTE"] / row["PESO_EMBALAGEM"], 2) if row["PESO_EMBALAGEM"] > 0 else row["PRECO_PACOTE"], axis=1)

        aba_grade, aba_precos, aba_conferencia, aba_relatorio = st.tabs(["📦 1. Grade por Filial", "🏪 2. Mesa de Decisão", "📑 3. Conferência & Disparo", "📈 4. Relatório de Divergências"])

        with aba_grade: 
            st.markdown("### 📋 Volume de Proteínas Solicitado por Filial") 
            st.caption("Aqui o sistema busca o que cada nutricionista digitou e monta a grade horizontal automática.") 
            try: 
                dados_auditoria_brutos = buscar_dados_tabela_cache("auditoria_consolidada") 
                if dados_auditoria_brutos: 
                    df_auditoria_total = pd.DataFrame(dados_auditoria_brutos) 
                    df_auditoria_total.columns = [str(c).strip().upper() for c in df_auditoria_total.columns] 
                    
                    if "STATUS_VALIDACAO" in df_auditoria_total.columns:
                        df_validos = df_auditoria_total[df_auditoria_total["STATUS_VALIDACAO"].astype(str).str.upper().isin(["APROVADO", "CONCLUÍDO FILIAL", "DENTRO DO LIMITE", "⚠ EXCEÇÃO (ESTOURADO)"])]
                        if df_validos.empty:
                            df_validos = df_auditoria_total
                    else:
                        df_validos = df_auditoria_total

                    if not df_validos.empty and all(c in df_validos.columns for c in ["PRODUTO", "RESTAURANTE", "PEDIDO_NUTRICIONISTA"]): 
                        df_validos["PEDIDO_NUTRICIONISTA"] = df_validos["PEDIDO_NUTRICIONISTA"].apply(tratar_qtd_float) 
                        df_grade_filiais = df_validos.pivot_table(index="PRODUTO", columns="RESTAURANTE", values="PEDIDO_NUTRICIONISTA", aggfunc="sum").fillna(0.0).reset_index() 
                        df_grade_filiais.columns.name = None 
                        colunas_restaurantes = [col for col in df_grade_filiais.columns if col != "PRODUTO"] 
                        df_grade_filiais["Volume Total (KG)"] = df_grade_filiais[colunas_restaurantes].sum(axis=1) 
                        st.dataframe(df_grade_filiais, hide_index=True, use_container_width=True) 
                    else: st.info("ℹ️ Nenhum pedido aprovado na auditoria.") 
            except Exception as e: st.error(f"Erro: {e}")
        
        with aba_precos: 
            st.markdown("### 📊 Mesa de Decisão Comercial - Diretor Jardel")
            st.caption("O sistema calcula automaticamente o preço por KG equivalente e destaca em verde o menor preço de cada produto.")

            if "PRECO_KG_EQUIV" in df_bruto.columns and "FORNECEDOR" in df_bruto.columns:
                df_bruto["FORNECEDOR_CURTO"] = df_bruto["FORNECEDOR"].apply(encurtar_nome_fornecedor)
                df_mapa_completo = df_bruto.pivot_table(index="PRODUTO", columns="FORNECEDOR_CURTO", values="PRECO_KG_EQUIV", aggfunc="min").reset_index()
                df_mapa_completo.columns.name = None
            else: df_mapa_completo = pd.DataFrame()

            if not df_mapa_completo.empty:
                colunas_fornecedores = [col for col in df_mapa_completo.columns if col != "PRODUTO"]
                if not colunas_fornecedores: st.info("ℹ️ Nenhum fornecedor enviou propostas até o momento.")
                else:
                    df_mapa_completo["Menor R$/KG"] = df_mapa_completo[colunas_fornecedores].min(axis=1)
                    df_mapa_completo["Sugestão Sistema"] = df_mapa_completo[colunas_fornecedores].idxmin(axis=1)
                    
                    st.session_state["df_jardel_decisao_salvo"] = df_mapa_completo
                    st.session_state["colunas_fornecedores_ativos"] = colunas_fornecedores

                    cols_exibicao = ["PRODUTO"] + colunas_fornecedores + ["Menor R$/KG", "Sugestão Sistema"]
                    df_tabela_analise = df_mapa_completo[cols_exibicao].copy()

                    config_analise = {
                        "PRODUTO": st.column_config.TextColumn("Descrição do Produto", disabled=True),
                        "Menor R$/KG": st.column_config.NumberColumn("Menor Preço/KG", disabled=True, format="R$ %.2f"),
                        "Sugestão Sistema": st.column_config.TextColumn("Sugestão (Mais Barato)", disabled=True)
                    }
                    for forn in colunas_fornecedores:
                        config_analise[forn] = st.column_config.NumberColumn(f"Preço ({forn})", disabled=True, format="R$ %.2f")

                    def colorir_menor_preco(row):
                        estilos = [''] * len(row)
                        menor = row.get("Menor R$/KG", None)
                        if menor is not None:
                            for idx, col in enumerate(row.index):
                                if col in colunas_fornecedores:
                                    val = row[col]
                                    if pd.notna(val) and abs(val - menor) < 0.001:
                                        estilos[idx] = 'background-color: #d4edda; color: #155724; font-weight: bold;'
                        return estilos

                    df_estilizado_mesa = df_tabela_analise.style.apply(colorir_menor_preco, axis=1)
                    st.dataframe(df_estilizado_mesa, column_config=config_analise, hide_index=True, use_container_width=True)

                st.write("---") 
                if st.button("⚡ Fechar Cotação e Atualizar Base", type="primary", use_container_width=True): 
                    try:
                        timestamp_agora = str(dt_mod.now().strftime('%d/%m/%Y %H:%M:%S'))
                        registros_abertos = buscar_dados_tabela_cache("pedidos_abertos")
                        if registros_abertos:
                            for r_ab in registros_abertos:
                                executar_sql(
                                    "INSERT INTO historico_pedidos (timestamp_arquivamento, data_hora_pedido, filial, pedido_num, fornecedor, produto, quantidade, valor_total, status, data_previsao_entrega, numero_nf, custo_unitario_com_impostos) VALUES (:ts, :dh, :fil, :pnum, :forn, :prod, :qtd, :val, 'FINALIZADO', :dprev, :nnf, 0.0)",
                                    {
                                        "ts": timestamp_agora, "dh": str(r_ab.get("DATA_HORA", "")), "fil": str(r_ab.get("FILIAL", "")),
                                        "pnum": str(r_ab.get("PEDIDO_NUM", "")), "forn": str(r_ab.get("FORNECEDOR", "")), "prod": str(r_ab.get("PRODUTO", "")),
                                        "qtd": tratar_qtd_float(r_ab.get("QUANTIDADE", 0)), "val": tratar_preco_float(r_ab.get("VALOR_TOTAL", 0)),
                                        "dprev": str(r_ab.get("DATA_PREVISAO_ENTREGA", "")), "nnf": str(r_ab.get("NUMERO_NF", ""))
                                    }
                                )

                        executar_sql("DELETE FROM pedidos_abertos")
                        executar_sql("DELETE FROM auditoria_consolidada")
                        executar_sql("DELETE FROM bd_cotacao")

                        st.success("✅ Cotação finalizada! Os pedidos abertos foram arquivados com sucesso!") 
                        st.balloons(); st.cache_data.clear(); st.rerun()
                    except Exception as e_automacao: st.error(f"Erro ao processar o arquivamento: {e_automacao}")

        with aba_conferencia:
            st.markdown("### 📑 Espelho de Pedidos e Carrinho de Revisão")
            st.caption("Revê o pedido, altera o fornecedor de destino na tabela e gera os espelhos oficiais em PDF.")
            
            try:
                dados_aud = buscar_dados_tabela_cache("auditoria_consolidada")
                prazos_reais = {str(r.get("FORNECEDOR", "")).strip(): str(r.get("PRAZO_PAGAMENTO", "7 Dias")) for r in dados_aud if isinstance(dados_aud, list)}
                
                prazos_entrega_forn = {}
                dados_forn_cad = buscar_dados_tabela_cache("fornecedores")
                if dados_forn_cad:
                    for f_reg in dados_forn_cad:
                        nome_f = str(f_reg.get("FORNECEDOR", "")).strip()
                        tel_f = str(f_reg.get("TELEFONE", f_reg.get("CONTATO", "(32) 90000-0000"))).strip()
                        dias_e = 3 
                        for chave_col in ["PRAZO_ENTREGA", "DIAS_ENTREGA", "ENTREGA"]:
                            if chave_col in f_reg and str(f_reg[chave_col]).strip().isdigit():
                                dias_e = int(str(f_reg[chave_col]).strip())
                                break
                        if nome_f:
                            prazos_entrega_forn[nome_f.upper()] = {"telefone": tel_f, "dias_entrega": dias_e}

                if "df_jardel_decisao_salvo" in st.session_state and "colunas_fornecedores_ativos" in st.session_state:
                    df_dec = st.session_state["df_jardel_decisao_salvo"]
                    forn_lista = st.session_state["colunas_fornecedores_ativos"]
                    
                    if dados_aud:
                        df_aud_conf = pd.DataFrame(dados_aud)
                        df_aud_conf.columns = [str(c).strip().upper() for c in df_aud_conf.columns]
                        
                        if "STATUS_VALIDACAO" in df_aud_conf.columns:
                            df_aprov_conf = df_aud_conf[df_aud_conf["STATUS_VALIDACAO"].astype(str).str.upper().isin(["APROVADO", "CONCLUÍDO FILIAL", "DENTRO DO LIMITE", "⚠ EXCEÇÃO (ESTOURADO)"])].copy()
                        else:
                            df_aprov_conf = df_aud_conf.copy()

                        if not df_aprov_conf.empty and "RESTAURANTE" in df_aprov_conf.columns:
                            lista_filiais_disponiveis = sorted(df_aprov_conf["RESTAURANTE"].dropna().astype(str).str.strip().unique().tolist())
                            filial_selecionada_aba3 = st.selectbox("🏢 Escolhe a Filial para emitir os Pedidos:", lista_filiais_disponiveis, key="sb_filial_conferencia_aba3")
                            
                            if filial_selecionada_aba3:
                                df_itens_filial = df_aprov_conf[df_aprov_conf["RESTAURANTE"].astype(str).str.strip().str.upper() == filial_selecionada_aba3.upper()].copy()
                                
                                st.markdown(f"#### 🛒 Carrinho Editável - {filial_selecionada_aba3}")
                                key_carrinho_state = f"carrinho_df_{filial_selecionada_aba3}"
                                key_gerado_state = f"espelho_gerado_{filial_selecionada_aba3}"
                                
                                if key_carrinho_state not in st.session_state:
                                    lista_carrinho = []
                                    for _, row_i in df_itens_filial.iterrows():
                                        produto_nome = str(row_i.get("PRODUTO", "")).strip()
                                        qtd_nutri_original = tratar_qtd_float(row_i.get("PEDIDO_NUTRICIONISTA", 0.0))

                                        vencedor_aba2 = forn_lista[0] if forn_lista else "FORNECEDOR PADRÃO"
                                        l_dec = df_dec[df_dec["PRODUTO"] == produto_nome]
                                        
                                        preco_sugerido_unit = 0.0
                                        if not l_dec.empty:
                                            sug = str(l_dec.iloc[0].get("Sugestão Sistema", "")).strip()
                                            if sug in forn_lista: vencedor_aba2 = sug
                                            if vencedor_aba2 in l_dec.columns:
                                                val_p = l_dec.iloc[0][vencedor_aba2]
                                                preco_sugerido_unit = float(val_p) if pd.notna(val_p) else 0.0

                                        lista_carrinho.append({
                                            "Excluir?": False,
                                            "Produto": produto_nome,
                                            "Qtd Solicitada (KG)": float(qtd_nutri_original),
                                            "Fornecedor Destino": vencedor_aba2,
                                            "Preço Unit. (R$)": float(preco_sugerido_unit),
                                            "Preço Total (R$)": float(qtd_nutri_original * preco_sugerido_unit)
                                        })
                                    st.session_state[key_carrinho_state] = pd.DataFrame(lista_carrinho)

                                df_carrinho_editado = st.data_editor(
                                    st.session_state[key_carrinho_state],
                                    column_config={
                                        "Excluir?": st.column_config.CheckboxColumn("Remover", default=False),
                                        "Produto": st.column_config.TextColumn("Descrição do Produto", disabled=True),
                                        "Qtd Solicitada (KG)": st.column_config.NumberColumn("Qtd (KG)", min_value=0.0, step=0.5, format="%.2f"),
                                        "Fornecedor Destino": st.column_config.SelectboxColumn("Fornecedor Destino", options=forn_lista, required=True),
                                        "Preço Unit. (R$)": st.column_config.NumberColumn("Preço Unit. (R$)", disabled=True, format="R$ %.2f"),
                                        "Preço Total (R$)": st.column_config.NumberColumn("Preço Total (R$)", disabled=True, format="R$ %.2f")
                                    },
                                    hide_index=True,
                                    use_container_width=True,
                                    key=f"carrinho_edit_livre_{filial_selecionada_aba3}"
                                )

                                mudou = False
                                for i in range(len(df_carrinho_editado)):
                                    f_novo = df_carrinho_editado.loc[i, "Fornecedor Destino"]
                                    f_antigo = st.session_state[key_carrinho_state].loc[i, "Fornecedor Destino"]
                                    q_novo = tratar_qtd_float(df_carrinho_editado.loc[i, "Qtd Solicitada (KG)"])
                                    q_antigo = tratar_qtd_float(st.session_state[key_carrinho_state].loc[i, "Qtd Solicitada (KG)"])
                                    exc_novo = df_carrinho_editado.loc[i, "Excluir?"]
                                    exc_antigo = st.session_state[key_carrinho_state].loc[i, "Excluir?"]

                                    if f_novo != f_antigo or q_novo != q_antigo or exc_novo != exc_antigo:
                                        mudou = True
                                        p_nome = df_carrinho_editado.loc[i, "Produto"]
                                        p_unit = 0.0
                                        l_dec = df_dec[df_dec["PRODUTO"] == p_nome]
                                        if not l_dec.empty and f_novo in l_dec.columns:
                                            val_f = l_dec.iloc[0][f_novo]
                                            p_unit = float(val_f) if pd.notna(val_f) else 0.0
                                        
                                        df_carrinho_editado.loc[i, "Preço Unit. (R$)"] = p_unit
                                        df_carrinho_editado.loc[i, "Preço Total (R$)"] = round(q_novo * p_unit, 2)

                                if mudou:
                                    st.session_state[key_carrinho_state] = df_carrinho_editado
                                    st.rerun()

                                valor_total_geral_carrinho = df_carrinho_editado[df_carrinho_editado["Excluir?"] == False]["Preço Total (R$)"].sum()
                                st.markdown(f"### 💰 **Valor Total do Carrinho: R$ {valor_total_geral_carrinho:,.2f}**")
                                st.markdown("---")
                                
                                if st.button(f"🖨️ Gerar Espelhos de Pedidos Oficiais para {filial_selecionada_aba3}", type="primary", use_container_width=True):
                                    itens_validos = df_carrinho_editado[(df_carrinho_editado["Excluir?"] == False) & (df_carrinho_editado["Qtd Solicitada (KG)"] > 0)]
                                    if itens_validos.empty:
                                        st.warning("⚠️ Nenhum item no carrinho para emitir.")
                                        st.session_state[key_gerado_state] = False
                                    else:
                                        st.session_state[key_gerado_state] = True
                                        st.rerun()

                                if st.session_state.get(key_gerado_state, False):
                                    st.success("✅ Espelhos gerados com sucesso! Selecione o fornecedor abaixo para baixar o PDF ou disparar o pedido:")
                                    
                                    itens_validos_gerar = df_carrinho_editado[(df_carrinho_editado["Excluir?"] == False) & (df_carrinho_editado["Qtd Solicitada (KG)"] > 0)]
                                    forn_unicos = itens_validos_gerar["Fornecedor Destino"].unique()
                                    dicionario_filiais_sheets = carregar_dados_filiais_dict()
                                    dados_filial = dicionario_filiais_sheets.get(filial_selecionada_aba3.upper(), {
                                        "RAZAO": f"AC BATISTA - {filial_selecionada_aba3}", "CNPJ": "06.121.429/0001-18", "IE": "625274795.00.00", "ENDERECO": "SÃO JOÃO DEL REI/MG", "CEP": "36300-000", "EMAIL": "comprasacbatista@gmail.com"
                                    })

                                    abas_forn = st.tabs([f"📦 Fornecedor: {f}" for f in forn_unicos])
                                    
                                    for idx_f, forn_alvo in enumerate(forn_unicos):
                                        with abas_forn[idx_f]:
                                            df_f_pedidos = itens_validos_gerar[itens_validos_gerar["Fornecedor Destino"] == forn_alvo].copy()
                                            
                                            linhas_espelho = []
                                            for _, row_item in df_f_pedidos.iterrows():
                                                p_nome = row_item["Produto"]
                                                p_qtd = tratar_qtd_float(row_item["Qtd Solicitada (KG)"])
                                                preco_u = float(row_item["Preço Unit. (R$)"])
                                                
                                                linhas_espelho.append({
                                                    "Código": "0545",
                                                    "Produto": p_nome,
                                                    "Preço Unit.": preco_u,
                                                    "Qtd": p_qtd,
                                                    "Total": preco_u * p_qtd
                                                })
                                            
                                            df_tabela_espelho = pd.DataFrame(linhas_espelho)
                                            info_forn = prazos_entrega_forn.get(forn_alvo.upper(), {"telefone": "(32) 99999-9999", "dias_entrega": 3})
                                            tel_fornecedor = info_forn["telefone"]
                                            dias_uteis_entrega = info_forn["dias_entrega"]
                                            prazo_forn = prazos_reais.get(forn_alvo, "7/14 Dias")
                                            data_prevista_auto = date.today() + timedelta(days=dias_uteis_entrega)
                                            
                                            with st.container(border=True):
                                                st.markdown(f"### 👨‍🍳 Pedido Nº {idx_f+1:04d} para **{forn_alvo}** ({filial_selecionada_aba3})")
                                                st.markdown(f"📞 **Telefone:** {tel_fornecedor} | 💳 **Prazo:** {prazo_forn} | 🚚 **Previsão:** {data_prevista_auto.strftime('%d/%m/%Y')}")
                                                
                                                st.dataframe(df_tabela_espelho, use_container_width=True, hide_index=True, column_config={
                                                    "Preço Unit.": st.column_config.NumberColumn(format="R$ %.2f"),
                                                    "Total": st.column_config.NumberColumn(format="R$ %.2f")
                                                })
                                                val_total_pedido = df_tabela_espelho["Total"].sum()
                                                st.markdown(f"### **Valor Total do Pedido: R$ {val_total_pedido:,.2f}**")
                                                
                                                prev_entrega = st.date_input("Data de Previsão de Entrega:", value=data_prevista_auto, key=f"prev_pdf_{filial_selecionada_aba3}_{idx_f+1}_{forn_alvo}", format="DD/MM/YYYY")
                                                
                                                pdf_bytes = gerar_pdf_pedido(
                                                    num_pedido=f"{idx_f+1:04d}",
                                                    filial_nome=filial_selecionada_aba3,
                                                    dados_filial=dados_filial,
                                                    forn_alvo=forn_alvo,
                                                    telefone_forn=tel_fornecedor,
                                                    prazo_pgto=prazo_forn,
                                                    df_itens=df_tabela_espelho,
                                                    data_entrega=prev_entrega
                                                )
                                                
                                                st.download_button(
                                                    label=f"📥 Baixar Espelho em PDF (Pedido {idx_f+1:04d} - {forn_alvo})",
                                                    data=pdf_bytes,
                                                    file_name=f"Pedido_{idx_f+1:04d}_{forn_alvo}_{filial_selecionada_aba3}.pdf",
                                                    mime="application/pdf",
                                                    key=f"dl_pdf_{filial_selecionada_aba3}_{idx_f+1}_{forn_alvo}"
                                                )
                                                
                                                st.write("")
                                                if st.button(f"🚀 Disparar Pedido Oficial Nº {idx_f+1:04d} para {forn_alvo}", key=f"btn_disp_{filial_selecionada_aba3}_{idx_f+1}_{forn_alvo}", type="primary"):
                                                    try:
                                                        ts_registro = dt_mod.now().strftime('%d/%m/%Y %H:%M:%S')
                                                        str_prev_entrega = prev_entrega.strftime('%d/%m/%Y')
                                                        
                                                        for _, row_grv in df_tabela_espelho.iterrows():
                                                            executar_sql(
                                                                "INSERT INTO pedidos_abertos (data_hora, filial, pedido_num, fornecedor, produto, quantidade, valor_total, status, data_previsao_entrega, numero_nf) VALUES (:dh, :fil, :pnum, :forn, :prod, :qtd, :val, 'ATIVO', :dprev, '')",
                                                                {
                                                                    "dh": ts_registro, "fil": filial_selecionada_aba3, "pnum": f"{idx_f+1:04d}",
                                                                    "forn": forn_alvo, "prod": str(row_grv["Produto"]),
                                                                    "qtd": tratar_qtd_float(row_grv['Qtd']), "val": tratar_preco_float(row_grv['Total']),
                                                                    "dprev": str_prev_entrega
                                                                }
                                                            )
                                                        st.success(f"✅ Pedido Nº {idx_f+1:04d} disparado e registado com sucesso!")
                                                        st.cache_data.clear()
                                                    except Exception as erro_aberto:
                                                        st.error(f"Erro ao guardar na base: {erro_aberto}")
            except Exception as e_conf:
                st.error(f"Erro ao montar a conferência: {e_conf}")

        with aba_relatorio:
            st.markdown("### 📈 Relatório Gerencial de Divergências")
            try:
                dados_hist_rel = buscar_dados_tabela_cache("historico_pedidos")
                if dados_hist_rel:
                    df_h_rel = pd.DataFrame(dados_hist_rel)
                    st.dataframe(df_h_rel, use_container_width=True, hide_index=True)
                else:
                    st.info("ℹ️ Nenhum dado no histórico de pedidos.")
            except Exception as e_rel:
                st.error(f"Erro ao carregar relatório: {e_rel}")
    except Exception as erro_modulo_compras: 
        st.error(f"❌ Erro ao processar o painel: {erro_modulo_compras}")

def interface_lancamento_proteina_filial(filial_passada="CENTRO"):
    try:
        st.subheader("📋 Digitação Semanal por Filial")
        nivel_usuario = st.session_state.get('nivel', 'Comum')
        filial_do_login = st.session_state.get('filial_nome', 'TEJUCO')
        
        if nivel_usuario == "Admin":
            filial_selected = st.selectbox("🏢 Seleciona a Filial:", MOCK_FILIAIS, index=MOCK_FILIAIS.index(filial_do_login) if filial_do_login in MOCK_FILIAIS else 0)
        else:
            filial_selected = filial_do_login
            st.info(f"📍 **A tua Filial Ativa:** {filial_selected}")
            
        if not filial_selected: return
        
        tab_lancamento, tab_conferencia = st.tabs(["📌 Aba 1: Lançamento", "🔍 Aba 2: Conferência"])
        
        with tab_lancamento:
            status_liberado = ler_status_digitacao() 
            if st.session_state.get('nivel') != "Admin" and not status_liberado:
                st.warning("🛑 A digitação semanal está fechada pelo Diretor Jardel.")
            else:
                proteinas_lista = carregar_proteinas_semanal()
                if not proteinas_lista: 
                    st.warning("Nenhuma proteína localizada.")
                else:
                    produto_selecionado = st.selectbox("Seleciona a Proteína:", [""] + proteinas_lista, format_func=lambda x: "-- Seleciona --" if x == "" else x)
                    if produto_selecionado:
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            semana_selecionada = st.selectbox("Semanas:", ["1 Semana", "2 Semanas"], key="sem_sel")
                            st.session_state['semana_atual'] = "Semana 1" if semana_selecionada == "1 Semana" else "Semana 2"
                        with col2: dias_input = st.number_input("Dias", value=0.0, min_value=0.0, step=1.0)
                        with col3: stock_input = st.number_input("Estoque", value=0.0, min_value=0.0, step=0.1)
                        
                        pedido_input = st.number_input("Quantidade Pedido (KG)", value=0.0, min_value=0.0, step=1.0)
                        filial_limpa = str(filial_selected).strip().upper()
                        
                        dados_media = carregar_dados_planilha()
                        media_refeicoes_diaria = 100.0
                        for row in dados_media:
                            if str(row.get('RESTAURANTES', '')).strip().upper() == filial_limpa:
                                try: media_refeicoes_diaria = float(str(row.get('JUNHO', 100.0)).replace(",", "."))
                                except: media_refeicoes_diaria = 100.0
                                break
                                
                        dias_calculo = float(dias_input) if dias_input > 0 else 1.0
                        pedido_sugerido = max((media_refeicoes_diaria * 0.165 * dias_calculo) - stock_input, 0.0)
                        st.metric(label="Sugerido", value=f"{pedido_sugerido:.1f} KG")
                        
                        status_v, jst, block = "DENTRO DO LIMITE", "", False
                        if pedido_input <= 0.0:
                            st.error("❌ Não podes enviar pedido zerado.")
                            block = True
                        elif pedido_input > pedido_sugerido:
                            status_v = "⚠ EXCEÇÃO (ESTOURADO)"
                            jst = st.text_input("Justificativa Obrigatória:", key="just_prot")
                            if not jst.strip(): block = True
                            
                        if st.button("➕ ADICIONAR À CONFERÊNCIA", disabled=block, key="btn_grv"):
                            try:
                                executar_sql(
                                    "INSERT INTO auditoria_consolidada (data_envio, restaurante, produto, semanas_pedidas, estoque_atual, pedido_nutricionista, sugestao_sistema, status_validacao, justificativa) VALUES (:dh, :rest, :prod, :sem, :est, :ped, :sug, :stat, :jst)",
                                    {
                                        "dh": str(dt_mod.now().strftime('%d/%m/%Y %H:%M:%S')), "rest": filial_selected, "prod": produto_selecionado,
                                        "sem": st.session_state['semana_atual'], "est": f"{stock_input:.3f}", "ped": f"{pedido_input:.3f}",
                                        "sug": f"{pedido_sugerido:.3f}", "stat": status_v, "jst": jst
                                    }
                                )
                                st.success("✔️ Adicionado com sucesso!")
                                st.cache_data.clear(); time_lib.sleep(1); st.rerun()
                            except Exception as e: st.error(f"Erro: {e}")
        
        with tab_conferencia:
            st.subheader("📋 Conferência de Pedidos Pendentes")
            try:
                todos_dados = buscar_dados_tabela_cache("auditoria_consolidada")
                if todos_dados:
                    df_r = pd.DataFrame(todos_dados)
                    df_r.columns = [str(c).strip().upper() for c in df_r.columns]
                    filial_limpa = str(filial_selected).strip().upper()
                    
                    if "RESTAURANTE" in df_r.columns and "STATUS_VALIDACAO" in df_r.columns:
                        df_r["RESTAURANTE"] = df_r["RESTAURANTE"].astype(str).str.strip().str.upper()
                        df_f = df_r[(df_r["RESTAURANTE"] == filial_limpa) & (~df_r["STATUS_VALIDACAO"].astype(str).str.upper().str.contains("CONCLUÍDO|APROVADO", na=False))].copy()
                        
                        if not df_f.empty:
                            df_f["EXCLUIR"] = False
                            cols_nec = ["EXCLUIR", "PRODUTO", "SEMANAS_PEDIDAS", "PEDIDO_NUTRICIONISTA", "STATUS_VALIDACAO", "JUSTIFICATIVA"]
                            for col in cols_nec:
                                if col not in df_f.columns: df_f[col] = ""
                            df_ex = df_f[cols_nec].copy()
                            df_ex["PEDIDO_NUTRICIONISTA"] = df_ex["PEDIDO_NUTRICIONISTA"].apply(tratar_qtd_float)
                            df_ex.columns = ["Excluir?", "Produto", "Semana", "Qtd Solicitada (KG)", "Status", "Justificativa Operacional"]
                            
                            df_ed = st.data_editor(df_ex, hide_index=True, use_container_width=True, key=f"ed_lote_{filial_limpa}")
                            
                            if st.button("🚀 CONFIRMAR E ENVIAR PARA O DIRETOR", type="primary", use_container_width=True):
                                for _, item_editado in df_ed.iterrows():
                                    if item_editado["Excluir?"]:
                                        executar_sql(
                                            "DELETE FROM auditoria_consolidada WHERE UPPER(TRIM(restaurante)) = :rest AND UPPER(TRIM(produto)) = :prod AND UPPER(TRIM(semanas_pedidas)) = :sem",
                                            {"rest": filial_limpa, "prod": str(item_editado["Produto"]).strip().upper(), "sem": str(item_editado["Semana"]).strip().upper()}
                                        )
                                    else:
                                        executar_sql(
                                            "UPDATE auditoria_consolidada SET pedido_nutricionista = :qtd, justificativa = :jst, status_validacao = 'CONCLUÍDO FILIAL' WHERE UPPER(TRIM(restaurante)) = :rest AND UPPER(TRIM(produto)) = :prod AND UPPER(TRIM(semanas_pedidas)) = :sem",
                                            {"qtd": f"{float(item_editado['Qtd Solicitada (KG)']):.3f}", "jst": str(item_editado["Justificativa Operacional"]), "rest": filial_limpa, "prod": str(item_editado["Produto"]).strip().upper(), "sem": str(item_editado["Semana"]).strip().upper()}
                                        )
                                st.success("✔️ Lote processado e enviado com sucesso!")
                                st.cache_data.clear(); time_lib.sleep(1); st.rerun()
            except Exception as e_conf:
                st.error(f"Erro ao processar lote: {e_conf}")
    except Exception as erro_mod: 
        st.error(f"❌ Erro: {erro_mod}")

if conn_db is None: conn_db = conectar_supabase()
if st.session_state.get('logado', False):
    st.sidebar.write(f"👤 Utilizador: **{st.session_state.get('usuario', 'Nenhum')}**")
    st.sidebar.write(f"🔑 Acesso: **{st.session_state.get('nivel', 'Comum')}**")
    if st.sidebar.button("🚪 Sair/Logoff"): st.session_state.logado = False; st.rerun()

if not st.session_state.get('logado', False): st.stop()
st.sidebar.write("---")

nivel_atual = st.session_state.get('nivel', 'Comum')
if nivel_atual == 'Admin':
    opcoes_menu = ["🔍 Conferência e Consolidação", "🥗 Lançamento de Pedidos", "📊 Cotação & Consolidação"]
elif nivel_atual == 'Fornecedor':
    opcoes_menu = ["🤝 Portal de Cotação"]
else:
    opcoes_menu = ["🥗 Lançamento de Pedidos"]

modulo_selecionado = st.sidebar.radio("Navegação do Sistema:", opcoes_menu)

if modulo_selecionado == "🔍 Conferência e Consolidação":
    st.title("🔍 Conferência e Consolidação de Pedidos")
    with st.container(border=True):
        status_atual = ler_status_digitacao()
        col_status, col_acao = st.columns([1, 1])
        with col_status:
            if status_atual: st.markdown("🟢 **Cotação LIBERADA**")
            else: st.markdown("🔴 **Cotação BLOQUEADA**")
        with col_acao:
            if status_atual:
                if st.button("Bloquear Cotação", use_container_width=True): salvar_governanca(False); st.rerun()
            else:
                if st.button("Liberar Cotação", use_container_width=True): salvar_governanca(True); st.rerun()
    restaurante_filtrado = st.selectbox("Filtrar por Restaurante:", ["Todos"] + MOCK_FILIAIS) 
    consolidar_proteina_semanal_geral(restaurante_filtrado)

elif modulo_selecionado == "🥗 Lançamento de Pedidos":
    interface_lancamento_proteina_filial()

elif modulo_selecionado == "🤝 Portal de Cotação":
    st.title("🤝 Portal do Fornecedor")
    fornecedor_logado = st.session_state.get('filial_nome', 'FORNECEDOR PADRÃO')
    st.success(f"🏢 Empresa Logada: {fornecedor_logado}")
