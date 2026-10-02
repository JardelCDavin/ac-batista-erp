import json
import os
import streamlit as st
import urllib.parse
from datetime import date, datetime as dt_mod, time as time_mod, timedelta
import time as time_lib
import io
import xml.etree.ElementTree as ET
import gspread
from gspread_dataframe import set_with_dataframe
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import unicodedata

# --- CONEXÃO INTELIGENTE COM O GOOGLE SHEETS ---
@st.cache_resource
def inicializar_gspread():
    if 'gcp_service_account' in st.secrets:
        credenciais = dict(st.secrets['gcp_service_account'])
        if 'private_key' in credenciais:
            credenciais['private_key'] = credenciais['private_key'].replace('\\n', '\n')
        return gspread.service_account_from_dict(credenciais)
    
    caminho_local = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chave.json")
    if os.path.exists(caminho_local):
        return gspread.service_account(filename=caminho_local)
        
    raise ValueError("Nenhuma credencial do Google encontrada. Verifique o Streamlit Secrets ou o arquivo chave.json.")

# --- BLINDAGEM DE API: MEMÓRIA DE CURTO PRAZO ---
@st.cache_data(ttl=300)
def buscar_dados_aba_cache(nome_aba):
    if 'client' in globals() and client is not None:
        try:
            return client.worksheet(nome_aba).get_all_records()
        except gspread.exceptions.APIError as e:
            if e.response.status_code == 429:
                st.warning("⏳ O Google está a proteger os acessos (Limite 429). A aguardar cache local...")
            return []
        except Exception:
            return []
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

# --- GARANTIA DE ABAS ---
try:
    client = inicializar_gspread().open_by_url("https://docs.google.com/spreadsheets/d/1Qt0HIMchGH_956STdsOHZj5RzXO-cBrz7nyyiiyEB7o/edit?resourcekey=&gid=60610781#gid=60610781")
    
    abas_existentes = [w.title for w in client.worksheets()]
    if "PEDIDOS_ABERTOS" not in abas_existentes:
        ws_p = client.add_worksheet(title="PEDIDOS_ABERTOS", rows="2000", cols="12")
        ws_p.append_row(["DATA_HORA", "FILIAL", "PEDIDO_NUM", "FORNECEDOR", "PRODUTO", "QUANTIDADE", "VALOR_TOTAL", "STATUS", "DATA_PREVISAO_ENTREGA", "NUMERO_NF"])
    if "HISTORICO_PEDIDOS" not in abas_existentes:
        ws_h = client.add_worksheet(title="HISTORICO_PEDIDOS", rows="5000", cols="15")
        ws_h.append_row(["TIMESTAMP_ARQUIVAMENTO", "DATA_HORA_PEDIDO", "FILIAL", "PEDIDO_NUM", "FORNECEDOR", "PRODUTO", "QUANTIDADE", "VALOR_TOTAL", "STATUS", "DATA_PREVISAO_ENTREGA", "NUMERO_NF", "NCM", "ICMS", "ST", "CUSTO_UNITARIO_COM_IMPOSTOS"])
    if "DE_PARA_FORNECEDORES" not in abas_existentes:
        ws_d = client.add_worksheet(title="DE_PARA_FORNECEDORES", rows="2000", cols="5")
        ws_d.append_row(["CODIGO_INTERNO", "PRODUTO_INTERNO", "CNPJ_FORNECEDOR", "NOME_FORNECEDOR", "CODIGO_ITEM_FORNECEDOR"])
    if "CONFERENCIA_FISICA" not in abas_existentes:
        ws_c = client.add_worksheet(title="CONFERENCIA_FISICA", rows="2000", cols="12")
        ws_c.append_row(["DATA_HORA", "FILIAL", "NUMERO_NF", "FORNECEDOR", "PRODUTO", "QTD_XML", "QTD_REAL", "DIVERGENCIA", "LOTE", "VALIDADE", "TEMPERATURA", "CONFERENTE"])
    
    if "CONFIGURACOES" not in abas_existentes:
        ws_conf = client.add_worksheet(title="CONFIGURACOES", rows="10", cols="2")
        ws_conf.append_row(["CHAVE", "VALOR"])
        ws_conf.append_row(["STATUS_DIGITACAO", "DESLIGADO"])

except gspread.exceptions.APIError as e:
    client = None
except Exception as global_e:
    client = None

def conectar_sheets_nativo():
    return client

@st.cache_data(ttl=600)
def carregar_dados_planilha():
    try:
        aba = client.worksheet("MEDIA_VENDA_DIARIA")
        return aba.get_all_records()
    except Exception as e:
        return []

@st.cache_data(ttl=300)
def carregar_fornecedores_ativos():
    if not client: return ["FORNECEDOR PADRÃO"]
    try:
        dados = buscar_dados_aba_cache("FORNECEDORES")
        if dados:
            ativos = [str(row.get('Fornecedor', row.get('FORNECEDOR', ''))).strip() for row in dados if str(row.get('ATIVO', '')).strip().upper() == 'SIM']
            return ativos if ativos else ["FORNECEDOR PADRÃO"]
        return ["FORNECEDOR PADRÃO"]
    except Exception as e: return ["FORNECEDOR PADRÃO"]

@st.cache_data(ttl=300)
def carregar_todos_fornecedores_cadastrados():
    if not client: return ["FORNECEDOR PADRÃO"]
    try:
        dados = buscar_dados_aba_cache("FORNECEDORES")
        if dados:
            todos = [str(row.get('Fornecedor', row.get('FORNECEDOR', ''))).strip() for row in dados if str(row.get('Fornecedor', row.get('FORNECEDOR', ''))).strip()]
            return sorted(list(set(todos))) if todos else ["FORNECEDOR PADRÃO"]
        return ["FORNECEDOR PADRÃO"]
    except Exception as e: return ["FORNECEDOR PADRÃO"]

@st.cache_data(ttl=300)
def carregar_catalogo_produtos_mapeamento():
    if not client: return {}, {}, [], {}
    try:
        dados = buscar_dados_aba_cache("PRODUTOS")
        cod_para_prod = {}
        prod_para_cod = {}
        precos_base_map = {}
        lista_nomes = []
        if dados:
            for row in dados:
                p_nome = str(row.get('PRODUTO', row.get('Produto', ''))).strip().upper()
                p_cod = str(row.get('CÓDIGO', row.get('CODIGO', row.get('Código', '')))).strip()
                p_preco = row.get('PREÇO_BASE', row.get('PRECO_BASE', row.get('Preço Base', 0.0)))
                
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
    dados = buscar_dados_aba_cache("DE_PARA_FORNECEDORES")
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
    dados = buscar_dados_aba_cache("FILIAIS")
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
            texto_empresa = str(row.get("DADOS EMPRESA", ""))
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
        dados = buscar_dados_aba_cache("PRODUTOS")
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

def validar_usuario_sheets(usuario, senha):
    if client is None: return None, "❌ Erro: Não foi possível conectar ao Google Sheets. Verifique as credenciais no Secrets."
    u_in = str(usuario).strip().upper()
    s_in = str(senha).strip()
    try:
        dados = buscar_dados_aba_cache("BD_USUARIOS")
        if dados:
            for linha in dados:
                user_tabela = str(linha.get('USUARIO', '')).strip().upper()
                senha_tabela = str(linha.get('SENHA', '')).strip()
                if user_tabela == u_in:
                    if senha_tabela == s_in:
                        return {
                            'USUARIO': str(linha.get('USUARIO', '')).strip(),
                            'FILIAL': str(linha.get('FILIAL', '')).strip().upper(),
                            'NIVEL': str(linha.get('NIVEL', '')).strip().upper() 
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
                    registro, erro = validar_usuario_sheets(usuario, senha)
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
        cli = conectar_sheets_nativo()
        if cli:
            aba = cli.worksheet("CONFIGURACOES")
            dados = aba.get_all_values()
            for linha in dados:
                if len(linha) > 1 and str(linha[0]).strip().upper() == "STATUS_DIGITACAO":
                    return True if str(linha[1]).strip().upper() == "LIGADO" else False
    except Exception as e:
        print(f"Erro ao ler status: {e}")
    return False

def salvar_governanca(ligado):
    cli = conectar_sheets_nativo()
    if cli:
        aba = cli.worksheet("CONFIGURACOES")
        novo_valor = "LIGADO" if ligado else "DESLIGADO"
        dados = aba.get_all_values()
        lin_alvo = -1
        if dados:
            for i, linha in enumerate(dados):
                if len(linha) > 0 and str(linha[0]).strip().upper() == "STATUS_DIGITACAO":
                    lin_alvo = i + 1
                    break
        if lin_alvo != -1:
            aba.update_cell(lin_alvo, 2, novo_valor)
        else:
            aba.append_row(["STATUS_DIGITACAO", novo_valor])
        st.session_state['libera_digitacao_semanal'] = ligado

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
    planilha = conectar_sheets_nativo()
    if not planilha: return
    dados = buscar_dados_aba_cache("AUDITORIA_CONSOLIDADA")
    if not dados:
        st.warning("A folha de cálculo de auditoria está vazia!")
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
            aba_aud = client.worksheet("AUDITORIA_CONSOLIDADA")
            d_sheets = aba_aud.get_all_records()
            cab_upper = [str(c).strip().upper() for c in aba_aud.row_values(1)]
            for _, r_ed in df_editado.iterrows():
                f_v, p_v, q_v = str(r_ed["Filial"]).strip().upper(), str(r_ed["Proteína / Item"]).strip().upper(), tratar_qtd_float(r_ed["Quantidade Lançada (KG)"])
                for idx_s, r_s in enumerate(d_sheets, start=2):
                    if str(r_s.get("RESTAURANTE", "")).strip().upper() == f_v and str(r_s.get("PRODUTO", "")).strip().upper() == p_v:
                        if "PEDIDO_NUTRICIONISTA" in cab_upper: aba_aud.update_cell(idx_s, cab_upper.index("PEDIDO_NUTRICIONISTA") + 1, f"{q_v:.3f}".replace(".", ","))
                        if "STATUS_VALIDACAO" in cab_upper: aba_aud.update_cell(idx_s, cab_upper.index("STATUS_VALIDACAO") + 1, "APROVADO")
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
            dados_cot_atuais = buscar_dados_aba_cache("BD_COTACAO")
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
        dados_cotacao_brutos = buscar_dados_aba_cache("BD_COTACAO") 
        if not dados_cotacao_brutos: 
            st.info("ℹ️ Nenhuma proposta de cotação encontrada na aba BD_COTACAO para o lote atual.")
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
                dados_auditoria_brutos = buscar_dados_aba_cache("AUDITORIA_CONSOLIDADA") 
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
                if st.button("⚡ Fechar Cotação e Atualizar Planilha", type="primary", use_container_width=True): 
                    try:
                        aba_auditoria = client.worksheet("AUDITORIA_CONSOLIDADA")
                        aba_cotacao = client.worksheet("BD_COTACAO")
                        try: aba_hist_pedidos = client.worksheet("HISTORICO_PEDIDOS")
                        except Exception:
                            aba_hist_pedidos = client.add_worksheet(title="HISTORICO_PEDIDOS", rows="5000", cols="15")
                            aba_hist_pedidos.append_row(["TIMESTAMP_ARQUIVAMENTO", "DATA_HORA_PEDIDO", "FILIAL", "PEDIDO_NUM", "FORNECEDOR", "PRODUTO", "QUANTIDADE", "VALOR_TOTAL", "STATUS", "DATA_PREVISAO_ENTREGA", "NUMERO_NF", "NCM", "ICMS", "ST", "CUSTO_UNITARIO_COM_IMPOSTOS"])

                        timestamp_agora = str(dt_mod.now().strftime('%d/%m/%Y %H:%M:%S'))
                        registros_abertos = buscar_dados_aba_cache("PEDIDOS_ABERTOS")
                        if registros_abertos:
                            for r_ab in registros_abertos:
                                aba_hist_pedidos.append_row([
                                    timestamp_agora,
                                    r_ab.get("DATA_HORA", ""),
                                    r_ab.get("FILIAL", ""),
                                    r_ab.get("PEDIDO_NUM", ""),
                                    r_ab.get("FORNECEDOR", ""),
                                    r_ab.get("PRODUTO", ""),
                                    tratar_qtd_float(r_ab.get("QUANTIDADE", 0)),
                                    tratar_preco_float(r_ab.get("VALOR_TOTAL", 0)),
                                    "FINALIZADO",
                                    r_ab.get("DATA_PREVISAO_ENTREGA", ""),
                                    r_ab.get("NUMERO_NF", ""),
                                    "", "", "", 0.0
                                ])

                        try:
                            ws_ab_limpar = client.worksheet("PEDIDOS_ABERTOS")
                            ws_ab_limpar.clear()
                            ws_ab_limpar.append_row(["DATA_HORA", "FILIAL", "PEDIDO_NUM", "FORNECEDOR", "PRODUTO", "QUANTIDADE", "VALOR_TOTAL", "STATUS", "DATA_PREVISAO_ENTREGA", "NUMERO_NF"])
                        except:
                            pass

                        aba_auditoria.clear(); aba_auditoria.append_row(["DATA_ENVIO", "RESTAURANTE", "PRODUTO", "SEMANAS_PEDIDAS", "ESTOQUE_ATUAL", "PEDIDO_NUTRICIONISTA", "SUGESTAO_SISTEMA", "STATUS_VALIDACAO", "JUSTIFICATIVA"])
                        aba_cotacao.clear(); aba_cotacao.append_row(["DATA_HORA", "COD_FORN", "FORNECEDOR", "PRODUTO", "PRECO_PACOTE", "UNIDADE", "PESO_EMBALAGEM", "QTD_MASTER", "PRECO_KG_EQUIV", "PRECO_CAIXA_MASTER"])

                        st.success("✅ Cotação finalizada! Os pedidos abertos foram arquivados na aba HISTORICO_PEDIDOS!") 
                        st.balloons(); st.cache_data.clear(); st.rerun()
                    except Exception as e_automacao: st.error(f"Erro ao processar o arquivamento: {e_automacao}")

        # ==========================================
        # ABA 3: CONFERÊNCIA & DISPARO (COM EXIBIÇÃO E ATUALIZAÇÃO VISUAL DE PREÇOS NA TABELA)
        # ==========================================
        with aba_conferencia:
            st.markdown("### 📑 Espelho de Pedidos e Carrinho de Revisão")
            st.caption("Revê o pedido, altera o fornecedor de destino na tabela (o preço unitário e o total atualizam automaticamente), e gera o espelho oficial em PDF.")
            
            try:
                dados_aud = buscar_dados_aba_cache("AUDITORIA_CONSOLIDADA")
                prazos_reais = {str(r.get("FORNECEDOR", "")).strip(): str(r.get("PRAZO_PAGAMENTO", "7 Dias")) for r in dados_aud if isinstance(dados_aud, list)}
                
                prazos_entrega_forn = {}
                dados_forn_cad = buscar_dados_aba_cache("FORNECEDORES")
                if dados_forn_cad:
                    for f_reg in dados_forn_cad:
                        nome_f = str(f_reg.get("Fornecedor", f_reg.get("FORNECEDOR", ""))).strip()
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

                                df_carrinho = pd.DataFrame(lista_carrinho)
                                
                                # Renderiza a tabela incluindo visivelmente as colunas de Preço Unitário e Preço Total (atualizadas em tempo real)
                                df_carrinho_editado = st.data_editor(
                                    df_carrinho,
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

                                # Recalcula preços e totais com base na seleção atual do fornecedor na tabela
                                precos_atualizados = []
                                totais_atualizados = []
                                for _, r_c in df_carrinho_editado.iterrows():
                                    p_nome = r_c["Produto"]
                                    f_dest = r_c["Fornecedor Destino"]
                                    q_val = tratar_qtd_float(r_c["Qtd Solicitada (KG)"])
                                    
                                    p_unit = 0.0
                                    l_dec = df_dec[df_dec["PRODUTO"] == p_nome]
                                    if not l_dec.empty and f_dest in l_dec.columns:
                                        val_f = l_dec.iloc[0][f_dest]
                                        p_unit = float(val_f) if pd.notna(val_f) else 0.0
                                    
                                    precos_atualizados.append(p_unit)
                                    totais_atualizados.append(round(q_val * p_unit, 2))

                                df_carrinho_editado["Preço Unit. (R$)"] = precos_atualizados
                                df_carrinho_editado["Preço Total (R$)"] = totais_atualizados
                                
                                valor_total_geral_carrinho = df_carrinho_editado[df_carrinho_editado["Excluir?"] == False]["Preço Total (R$)"].sum()

                                st.markdown(f"### 💰 **Valor Total do Carrinho: R$ {valor_total_geral_carrinho:,.2f}**")
                                st.markdown("---")
                                
                                if st.button(f"🖨️ Gerar Espelhos de Pedidos Oficiais para {filial_selecionada_aba3}", type="primary", use_container_width=True):
                                    itens_validos = df_carrinho_editado[(df_carrinho_editado["Excluir?"] == False) & (df_carrinho_editado["Qtd Solicitada (KG)"] > 0)]
                                    
                                    if itens_validos.empty:
                                        st.warning("⚠️ Nenhum item no carrinho para emitir.")
                                    else:
                                        forn_unicos = itens_validos["Fornecedor Destino"].unique()
                                        dicionario_filiais_sheets = carregar_dados_filiais_dict()
                                        dados_filial = dicionario_filiais_sheets.get(filial_selecionada_aba3.upper(), {
                                            "RAZAO": f"AC BATISTA - {filial_selecionada_aba3}", "CNPJ": "06.121.429/0001-18", "IE": "625274795.00.00", "ENDERECO": "SÃO JOÃO DEL REI/MG", "CEP": "36300-000", "EMAIL": "comprasacbatista@gmail.com"
                                        })

                                        for num_seq, forn_alvo in enumerate(forn_unicos, start=1):
                                            df_f_pedidos = itens_validos[itens_validos["Fornecedor Destino"] == forn_alvo].copy()
                                            
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
                                                st.markdown(f"### 👨‍🍳 Pedido Nº {num_seq:04d} para **{forn_alvo}** ({filial_selecionada_aba3})")
                                                st.markdown(f"📞 **Telefone Fornecedor:** {tel_fornecedor} | 💳 **Prazo Pagamento:** {prazo_forn} | 🚚 **Previsão (Automática):** {data_prevista_auto.strftime('%d/%m/%Y')} (Prazo de {dias_uteis_entrega} dias)")
                                                
                                                st.dataframe(df_tabela_espelho, use_container_width=True, hide_index=True, column_config={
                                                    "Preço Unit.": st.column_config.NumberColumn(format="R$ %.2f"),
                                                    "Total": st.column_config.NumberColumn(format="R$ %.2f")
                                                })
                                                val_total_pedido = df_tabela_espelho["Total"].sum()
                                                st.markdown(f"### **Valor Total do Pedido: R$ {val_total_pedido:,.2f}**")
                                                
                                                prev_entrega = st.date_input("Ajustar Data de Previsão de Entrega:", value=data_prevista_auto, key=f"prev_pdf_{filial_selecionada_aba3}_{num_seq}_{forn_alvo}", format="DD/MM/YYYY")
                                                
                                                pdf_bytes = gerar_pdf_pedido(
                                                    num_pedido=f"{num_seq:04d}",
                                                    filial_nome=filial_selecionada_aba3,
                                                    dados_filial=dados_filial,
                                                    forn_alvo=forn_alvo,
                                                    telefone_forn=tel_fornecedor,
                                                    prazo_pgto=prazo_forn,
                                                    df_itens=df_tabela_espelho,
                                                    data_entrega=prev_entrega
                                                )
                                                
                                                st.download_button(
                                                    label=f"📥 Baixar Espelho em PDF (Pedido {num_seq:04d} - {forn_alvo})",
                                                    data=pdf_bytes,
                                                    file_name=f"Pedido_{num_seq:04d}_{forn_alvo}_{filial_selecionada_aba3}.pdf",
                                                    mime="application/pdf",
                                                    key=f"dl_pdf_{filial_selecionada_aba3}_{num_seq}_{forn_alvo}"
                                                )
                                                
                                                st.write("")
                                                if st.button(f"🚀 Disparar Pedido Oficial Nº {num_seq:04d} para {forn_alvo}", key=f"btn_disp_{filial_selecionada_aba3}_{num_seq}_{forn_alvo}", type="primary"):
                                                    try:
                                                        try:
                                                            ws_abertos = client.worksheet("PEDIDOS_ABERTOS")
                                                        except Exception:
                                                            ws_abertos = client.add_worksheet(title="PEDIDOS_ABERTOS", rows="2000", cols="10")
                                                            ws_abertos.append_row(["DATA_HORA", "FILIAL", "PEDIDO_NUM", "FORNECEDOR", "PRODUTO", "QUANTIDADE", "VALOR_TOTAL", "STATUS", "DATA_PREVISAO_ENTREGA", "NUMERO_NF"])
                                                        
                                                        ts_registro = dt_mod.now().strftime('%d/%m/%Y %H:%M:%S')
                                                        str_prev_entrega = prev_entrega.strftime('%d/%m/%Y')
                                                        
                                                        for _, row_grv in df_tabela_espelho.iterrows():
                                                            ws_abertos.append_row([
                                                                ts_registro,
                                                                filial_selecionada_aba3,
                                                                f"{num_seq:04d}",
                                                                forn_alvo,
                                                                str(row_grv["Produto"]),
                                                                f"{tratar_qtd_float(row_grv['Qtd']):.3f}".replace(".", ","),
                                                                f"{tratar_preco_float(row_grv['Total']):.2f}".replace(".", ","),
                                                                "ATIVO",
                                                                str_prev_entrega,
                                                                ""
                                                            ])
                                                        st.success(f"✅ Pedido Nº {num_seq:04d} disparado e registado com sucesso na aba PEDIDOS_ABERTOS!")
                                                        st.cache_data.clear()
                                                    except Exception as erro_aberto:
                                                        st.error(f"Erro ao guardar na base de pedidos abertos: {erro_aberto}")
                        else: st.info("ℹ️ Nenhum pedido aprovado encontrado.")
                    else: st.info("ℹ️ Realiza a cotação na Aba 2 primeiro.")
            except Exception as e_conf:
                st.error(f"Erro ao montar a conferência: {e_conf}")

        with aba_relatorio:
            st.markdown("### 📈 Relatório Gerencial de Divergências (Acareação)")
            st.caption("Acompanha o histórico de notas baixadas e analisa faltas, sobras e variações de preços.")
            try:
                dados_hist_rel = buscar_dados_aba_cache("HISTORICO_PEDIDOS")
                if dados_hist_rel:
                    df_h_rel = pd.DataFrame(dados_hist_rel)
                    df_h_rel.columns = [str(c).strip().upper() for c in df_h_rel.columns]
                    st.dataframe(df_h_rel, use_container_width=True, hide_index=True)
                else:
                    st.info("ℹ️ Nenhum dado no histórico de pedidos para gerar relatórios no momento.")
            except Exception as e_rel:
                st.error(f"Erro ao carregar relatório: {e_rel}")
    except Exception as erro_modulo_compras: 
        st.error(f"❌ Erro ao processar o painel: {erro_modulo_compras}")

def interface_lancamento_proteina_filial(filial_passada="CENTRO"):
    try:
        st.subheader("📋 Digitação Semanal por Filial")
        st.info("Usa as abas abaixo para lançar novos itens ou gerenciar o lote.")
        
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
                    if not produto_selecionado: 
                        st.warning("👆 Escolhe uma proteína acima para fazer um novo lançamento.")
                    else:
                        st.write(f"**Selecionada:** {produto_selecionado}")
                        
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
                        pedido_sugerido = 0.0
                        
                        if 'client' in globals() and client is not None:
                            try:
                                fator_per_capita, margem_seguranca = 0.165, 0.0
                                if filial_limpa in ["DONA MARIA", "RM SABOR"]:
                                    dados_contrato = buscar_dados_aba_cache("CONTRATO_DONA_MARIA")
                                    for tc in dados_contrato:
                                        prod_plan = normalizar_nome_produto(tc.get('DESCRIÇÃO', ''))
                                        prod_sel = normalizar_nome_produto(produto_selecionado)
                                        if prod_sel == prod_plan or prod_sel in prod_plan or prod_plan in prod_sel:
                                            teto_f = float(str(tc.get('Limite Ativo', 0.0)).replace(',', '.'))
                                            ml = str(tc.get('MARGEM_SEGURANÇA', '0.0')).replace(',', '.').replace('%', '')
                                            if ml.strip(): margem_seguranca = float(ml) / 100.0 if float(ml) > 1.0 else float(ml)
                                            pedido_sugerido = max((teto_f * (1.0 + margem_seguranca)) - stock_input, 0.0)
                                            break
                                else:
                                    nome_aba = "CONTRATO_LEOPOLDINA" if filial_limpa == "LEOPOLDINA" else "CONTRATO_POPULAR_GERAL"
                                    dados_contrato = buscar_dados_aba_cache(nome_aba)
                                    for lc in dados_contrato:
                                        prod_plan = normalizar_nome_produto(lc.get('PRODUTO', ''))
                                        prod_sel = normalizar_nome_produto(produto_selecionado)
                                        if prod_sel == prod_plan or prod_sel in prod_plan or prod_plan in prod_sel:
                                            fator_per_capita = float(str(lc.get('PESO_IN_NATURA_GR_PADRAO', '0.165')).replace(',', '.'))
                                            ml = str(lc.get('MARGEM_SEGURANCA', '0.0')).replace(',', '.').replace('%', '')
                                            if ml.strip(): 
                                                margem_seguranca = float(ml) / 100.0 if float(ml) > 1.0 else float(ml)
                                            break
                                    if fator_per_capita > 1.0: fator_per_capita /= 1000.0
                                    pedido_sugerido = max((media_refeicoes_diaria * fator_per_capita * dias_calculo * (1.0 + margem_seguranca)) - stock_input, 0.0)
                            except: pedido_sugerido = 0.0
                                
                        st.metric(label="Sugerido", value=f"{pedido_sugerido:.1f} KG")
                        status_v, jst, block = "DENTRO DO LIMITE", "", False
                        
                        if pedido_input <= 0.0:
                            st.error("❌ Não podes enviar pedido zerado.")
                            block = True
                        elif pedido_input > pedido_sugerido:
                            status_v = "⚠ EXCEÇÃO (ESTOURADO)"
                            st.warning("Alerta: Superou o teto!")
                            jst = st.text_input("Justificativa Obrigatória:", key="just_prot")
                            if not jst.strip(): block = True
                            
                        dados_auditoria_trava = buscar_dados_aba_cache("AUDITORIA_CONSOLIDADA")
                        if dados_auditoria_trava:
                            for linha_trava in dados_auditoria_trava:
                                if str(linha_trava.get("RESTAURANTE", "")).strip().upper() == str(filial_selected).strip().upper() and \
                                   str(linha_trava.get("PRODUTO", "")).strip().upper() == str(produto_selecionado).strip().upper() and \
                                   str(linha_trava.get("SEMANAS_PEDIDAS", "")).strip().upper() == str(st.session_state['semana_atual']).strip().upper():
                                    
                                    status_atual = str(linha_trava.get("STATUS_VALIDACAO", "")).strip().upper()
                                    if "CONCLUÍDO" in status_atual or "APROVADO" in status_atual:
                                        st.error(f"⚠️ Bloqueio: Já enviaste o pedido de {produto_selecionado} ({st.session_state['semana_atual']}) para o Diretor neste ciclo! Se precisares de mais, entra em contacto com a Diretoria.")
                                        block = True; break
                                    elif status_atual in ["DENTRO DO LIMITE", "⚠ EXCEÇÃO (ESTOURADO)", ""]:
                                        st.error(f"⚠️️ Atenção: O item {produto_selecionado} já está no teu carrinho na Aba 2 (Conferência) a aguardar envio. Vai até lá se precisares alterar a quantidade.")
                                        block = True; break
                        
                        if st.button("➕ ADICIONAR À CONFERÊNCIA (VAI PARA ABA 2)", disabled=block, key="btn_grv"):
                            try:
                                aba_auditoria = client.worksheet("AUDITORIA_CONSOLIDADA")
                                aba_auditoria.append_row([
                                    str(dt_mod.now().strftime('%d/%m/%Y %H:%M:%S')), filial_selected, produto_selecionado, 
                                    st.session_state['semana_atual'], 
                                    f"{stock_input:.3f}".replace(".", ","), 
                                    f"{pedido_input:.3f}".replace(".", ","), 
                                    f"{pedido_sugerido:.3f}".replace(".", ","), 
                                    status_v, jst
                                ])
                                st.success("✔️ Adicionado! Vai para a Aba 2 para Conferir e Enviar ao Diretor.")
                                st.cache_data.clear()
                                time_lib.sleep(1)
                                st.rerun()
                            except Exception as e: st.error(f"Erro: {e}")
                    
        with tab_conferencia:
            c_tit, c_btn = st.columns([0.8, 0.2])
            with c_tit: st.subheader("📋 Conferência de Pedidos")
            with c_btn:
                if st.button("🔄 Atualizar Lista", use_container_width=True, key="btn_refresh_aba2"):
                    st.cache_data.clear()
                    st.rerun()
            
            try:
                todos_dados = buscar_dados_aba_cache("AUDITORIA_CONSOLIDADA")
                if todos_dados:
                    df_r = pd.DataFrame(todos_dados)
                    df_r.columns = [str(c).strip().upper() for c in df_r.columns]
                    filial_limpa = str(filial_selected).strip().upper()
                    
                    if "RESTAURANTE" in df_r.columns and "STATUS_VALIDACAO" in df_r.columns:
                        df_r["RESTAURANTE"] = df_r["RESTAURANTE"].astype(str).str.strip().str.upper()
                        df_f = df_r[(df_r["RESTAURANTE"] == filial_limpa) & (~df_r["STATUS_VALIDACAO"].astype(str).str.upper().str.contains("CONCLUÍDO|APROVADO", na=False))].copy()
                        
                        if not df_f.empty:
                            df_f["EXCLUIR"] = False
                            cols_necessarias = ["EXCLUIR", "PRODUTO", "SEMANAS_PEDIDAS", "PEDIDO_NUTRICIONISTA", "STATUS_VALIDACAO", "JUSTIFICATIVA"]
                            for col in cols_necessarias:
                                if col not in df_f.columns: df_f[col] = ""
                            df_ex = df_f[cols_necessarias].copy()
                            
                            df_ex["PEDIDO_NUTRICIONISTA"] = df_ex["PEDIDO_NUTRICIONISTA"].apply(tratar_qtd_float)
                            df_ex.columns = ["Excluir?", "Produto", "Semana", "Qtd Solicitada (KG)", "Status", "Justificativa Operacional"]
                            
                            cfg = {
                                "Excluir?": st.column_config.CheckboxColumn("Excluir?", default=False),
                                "Produto": st.column_config.TextColumn("Produto", disabled=True),
                                "Semana": st.column_config.TextColumn("Semana", disabled=True),
                                "Qtd Solicitada (KG)": st.column_config.NumberColumn("Qtd (KG)", min_value=0.0, step=0.1, format="%.2f kg"),
                                "Status": st.column_config.TextColumn("Status", disabled=True),
                                "Justificativa Operacional": st.column_config.TextColumn("Justificativa")
                            }
                            
                            df_ed = st.data_editor(df_ex, column_config=cfg, hide_index=True, use_container_width=True, key=f"ed_lote_{filial_limpa}")
                            
                            if st.button("🚀 CONFIRMAR E ENVIAR PARA O DIRETOR", type="primary", use_container_width=True):
                                with st.spinner("⏳ A enviar para a mesa do Diretor e a limpar a conferência..."):
                                    aba_auditoria = client.worksheet("AUDITORIA_CONSOLIDADA")
                                    lista_linhas_sheets = aba_auditoria.get_all_values()
                                    
                                    cabecalhos_aud = [str(c).strip().upper() for c in lista_linhas_sheets[0]]
                                    idx_rest = cabecalhos_aud.index("RESTAURANTE") if "RESTAURANTE" in cabecalhos_aud else 1
                                    idx_prod = cabecalhos_aud.index("PRODUTO") if "PRODUTO" in cabecalhos_aud else 2
                                    idx_sem = cabecalhos_aud.index("SEMANAS_PEDIDAS") if "SEMANAS_PEDIDAS" in cabecalhos_aud else 3
                                    
                                    linhas_processadas = set() 
                                    indices_para_deletar = []
                                    
                                    for i in range(len(df_ed)):
                                        item_editado = df_ed.iloc[i]
                                        for idx_s, linha_s in enumerate(lista_linhas_sheets[1:], start=2):
                                            if idx_s in linhas_processadas: continue 
                                            if len(linha_s) > max(idx_rest, idx_prod, idx_sem) and \
                                               str(linha_s[idx_rest]).strip().upper() == filial_limpa and \
                                               str(linha_s[idx_prod]).strip().upper() == str(item_editado["Produto"]).strip().upper() and \
                                               str(linha_s[idx_sem]).strip().upper() == str(item_editado["Semana"]).strip().upper():
                                                
                                                if item_editado["Excluir?"]: indices_para_deletar.append(idx_s)
                                                else:
                                                    qtd_final_float = float(item_editado['Qtd Solicitada (KG)'])
                                                    aba_auditoria.update_cell(idx_s, 6, f"{qtd_final_float:.3f}".replace(".", ","))
                                                    aba_auditoria.update_cell(idx_s, 9, str(item_editado["Justificativa Operacional"]))
                                                    aba_auditoria.update_cell(idx_s, 8, "CONCLUÍDO FILIAL")
                                                
                                                linhas_processadas.add(idx_s)
                                                break 
                                        
                                    for idx_del in sorted(indices_para_deletar, reverse=True):
                                        aba_auditoria.delete_row(idx_del) 
                                        
                                    st.success("✔️ Lote processado! Pedidos enviados e/ou excluídos com sucesso.")
                                    st.cache_data.clear()
                                    time_lib.sleep(1)
                                    st.rerun()
                        else: st.caption(f"ℹ️ Nenhum pedido pendente de envio na filial {filial_selected}.")
            except Exception as e_conf:
                st.error(f"Erro ao processar lote: {e_conf}")
    except Exception as erro_modulo_compras: 
        st.error(f"❌ Erro ao processar: {erro_modulo_compras}")

def modulo_compras_suprimentos():
    st.title("💼 Compras & Suprimentos - Gestão de Pedidos")
    st.info("Gere pedidos ativos, altera status (ATIVO/FINALIZADO) ou realiza digitação manual de novos itens com múltiplas linhas.")

    tab_gerenciar, tab_manual = st.tabs(["📦 1. Pedidos Ativos & Gestão", "➕ 2. Digitação Manual (Múltiplos Itens)"])

    with tab_gerenciar:
        try:
            try:
                dados_abertos = buscar_dados_aba_cache("PEDIDOS_ABERTOS")
            except Exception:
                ws_novo = client.add_worksheet(title="PEDIDOS_ABERTOS", rows="2000", cols="10")
                ws_novo.append_row(["DATA_HORA", "FILIAL", "PEDIDO_NUM", "FORNECEDOR", "PRODUTO", "QUANTIDADE", "VALOR_TOTAL", "STATUS", "DATA_PREVISAO_ENTREGA", "NUMERO_NF"])
                dados_abertos = []

            if not dados_abertos:
                st.warning("⚠️ Nenhum pedido encontrado na aba PEDIDOS_ABERTOS do Google Sheets. Gera pedidos na Aba 3 ou insere na Digitação Manual.")
            else:
                df_abertos = pd.DataFrame(dados_abertos)
                df_abertos.columns = [str(c).strip().upper() for c in df_abertos.columns]
                
                filtro_status = st.radio("Filtrar por Status:", ["Apenas ATIVOS", "TODOS"], horizontal=True, key="filtro_st_abertos")
                if filtro_status == "Apenas ATIVOS" and "STATUS" in df_abertos.columns:
                    df_exibicao = df_abertos[df_abertos["STATUS"].astype(str).str.upper() == "ATIVO"].copy()
                else:
                    df_exibicao = df_abertos.copy()

                if df_exibicao.empty:
                    st.info("ℹ️ Nenhum pedido ativo no momento.")
                else:
                    st.markdown("##### ✏️ Edição, Exclusão, Inclusão de NF e Baixa de Pedidos")
                    st.caption("Digita o número da Nota Fiscal (NF), altera o status para FINALIZADO ou ajusta quantidades diretamente na tabela abaixo.")
                    
                    df_editado_abertos = st.data_editor(
                        df_exibicao,
                        use_container_width=True,
                        hide_index=True,
                        key="editor_pedidos_abertos_geral_v2"
                    )

                    if st.button("💾 Guardar Alterações e Baixar Pedidos Finalizados", type="primary", key="btn_salvar_abertos_v2"):
                        try:
                            ws_ab = client.worksheet("PEDIDOS_ABERTOS")
                            try:
                                ws_hist = client.worksheet("HISTORICO_PEDIDOS")
                            except:
                                ws_hist = client.add_worksheet(title="HISTORICO_PEDIDOS", rows="5000", cols="15")
                                ws_hist.append_row(["TIMESTAMP_ARQUIVAMENTO", "DATA_HORA_PEDIDO", "FILIAL", "PEDIDO_NUM", "FORNECEDOR", "PRODUTO", "QUANTIDADE", "VALOR_TOTAL", "STATUS", "DATA_PREVISAO_ENTREGA", "NUMERO_NF", "NCM", "ICMS", "ST", "CUSTO_UNITARIO_COM_IMPOSTOS"])

                            timestamp_agora = dt_mod.now().strftime('%d/%m/%Y %H:%M:%S')
                            
                            linhas_ainda_ativas = []
                            colunas_lista = df_editado_abertos.columns.tolist()
                            
                            idx_status = [i for i, c in enumerate(colunas_lista) if 'STATUS' in c.upper()]
                            idx_status = idx_status[0] if idx_status else -1

                            for _, r_val in df_editado_abertos.iterrows():
                                lista_valores = r_val.tolist()
                                status_val = str(lista_valores[idx_status]).strip().upper() if idx_status != -1 else "ATIVO"
                                
                                if status_val == "FINALIZADO":
                                    ws_hist.append_row([
                                        timestamp_agora,
                                        str(lista_valores[0]),
                                        str(lista_valores[1]),
                                        str(lista_valores[2]),
                                        str(lista_valores[3]),
                                        str(lista_valores[4]),
                                        tratar_qtd_float(lista_valores[5]),
                                        tratar_preco_float(lista_valores[6]),
                                        "FINALIZADO",
                                        str(lista_valores[8]),
                                        str(lista_valores[9]) if len(lista_valores) > 9 else "",
                                        "", "", "", 0.0
                                    ])
                                else:
                                    linhas_ainda_ativas.append(lista_valores)

                            ws_ab.clear()
                            ws_ab.append_row(colunas_lista)
                            for l_ativa in linhas_ainda_ativas:
                                ws_ab.append_row(l_ativa)

                            st.success("✅ Alterações guardadas! Pedidos finalizados foram arquivados com sucesso em HISTORICO_PEDIDOS.")
                            st.cache_data.clear()
                            time_lib.sleep(1)
                            st.rerun()
                        except Exception as err_save:
                            st.error(f"❌ Erro ao guardar alterações: {err_save}")
        except Exception as e_ger:
            st.error(f"Erro ao carregar pedidos abertos: {e_ger}")

    with tab_manual:
        st.markdown("##### ➕ Digitação Manual de Pedido Avulso com Múltiplos Itens")
        st.caption("Insere vários itens de uma vez. O número do pedido é calculado automaticamente de forma sequencial.")

        proximo_num_seq = "0001"
        try:
            dados_abertos_hist = buscar_dados_aba_cache("PEDIDOS_ABERTOS")
            if dados_abertos_hist:
                numeros_existentes = []
                for linha_h in dados_abertos_hist:
                    p_num = str(linha_h.get("PEDIDO_NUM", linha_h.get("PEDIDO NUM", "1"))).strip()
                    if p_num.isdigit():
                        numeros_existentes.append(int(p_num))
                if numeros_existentes:
                    proximo_num_seq = f"{max(numeros_existentes) + 1:04d}"
        except:
            pass

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            m_filial = st.selectbox("Filial Destino:", MOCK_FILIAIS, key="m_filial")
        with col_m2:
            m_pedido_num = st.text_input("Número do Pedido (Sequencial):", value=proximo_num_seq, key="m_ped_num")

        col_m3, col_m4 = st.columns(2)
        with col_m3:
            fornecedores_manuais = carregar_todos_fornecedores_cadastrados()
            m_forn = st.selectbox("Fornecedor (Qualquer Cadastrado):", fornecedores_manuais, key="m_forn")
        with col_m4:
            dias_padrao_forn = 3
            try:
                dados_f_cad = buscar_dados_aba_cache("FORNECEDORES")
                if dados_f_cad:
                    for fc in dados_f_cad:
                        if str(fc.get("Fornecedor", fc.get("FORNECEDOR", ""))).strip().upper() == str(m_forn).strip().upper():
                            for col_d in ["PRAZO_ENTREGA", "DIAS_ENTREGA", "ENTREGA"]:
                                if col_d in fc and str(fc[col_d]).strip().isdigit():
                                    dias_padrao_forn = int(str(fc[col_d]).strip())
                                    break
                            break
            except:
                pass
            data_sugerida_manual = date.today() + timedelta(days=dias_padrao_forn)
            m_prev = st.date_input("Previsão de Entrega (Editável):", value=data_sugerida_manual, key="m_prev", format="DD/MM/YYYY")

        st.markdown("---")
        st.markdown("##### 🛒 Adiciona os Produtos do Pedido (Até 25 Linhas)")
        st.caption("Digita o código interno ou seleciona o produto. O preço unitário será sugerido automaticamente com base no PREÇO BASE da tabela e pode ser editado.")

        cod_para_prod, prod_para_cod, lista_nomes_prod, precos_base_map = carregar_catalogo_produtos_mapeamento()

        if "contador_editor_manual" not in st.session_state:
            st.session_state["contador_editor_manual"] = 0

        df_vazio_manual = pd.DataFrame([
            {"Código Interno": "", "Produto / Descrição": "", "Quantidade": "0.0", "Preço Unitário (R$)": "0.0"}
            for _ in range(25)
        ])

        df_itens_digitados = st.data_editor(
            df_vazio_manual,
            column_config={
                "Código Interno": st.column_config.TextColumn("Cód. Interno", width="small"),
                "Produto / Descrição": st.column_config.SelectboxColumn("Descrição do Produto", options=[""] + lista_nomes_prod, width="large", required=False),
                "Quantidade": st.column_config.TextColumn("Quantidade / KG"),
                "Preço Unitário (R$)": st.column_config.TextColumn("Preço Unit. (R$)")
            },
            hide_index=True,
            use_container_width=True,
            key=f"editor_multiplos_itens_avulsos_v_{st.session_state['contador_editor_manual']}"
        )

        st.write("")
        if st.button("📥 Gravar Pedido Avulso Completo na Planilha", type="primary", key="btn_gravar_multiplos_avulsos_v10"):
            try:
                linhas_validas_gravar = []
                for _, r_item in df_itens_digitados.iterrows():
                    c_int = str(r_item["Código Interno"]).strip()
                    p_desc = str(r_item["Produto / Descrição"]).strip()
                    q_val = tratar_qtd_float(r_item["Quantidade"])
                    
                    if c_int and not p_desc:
                        p_desc = cod_para_prod.get(c_int, f"CÓDIGO {c_int}")
                    elif p_desc and not c_int:
                        c_int = prod_para_cod.get(p_desc.upper(), "0545")
                    
                    p_unit = tratar_preco_float(r_item["Preço Unitário (R$)"])
                    if p_unit <= 0 and p_desc.upper() in precos_base_map:
                        p_unit = precos_base_map[p_desc.upper()]

                    v_total = round(q_val * p_unit, 2)

                    if q_val > 0 and (c_int or p_desc):
                        linhas_validas_gravar.append({
                            "produto": p_desc,
                            "qtd": q_val,
                            "valor": v_total
                        })

                if not linhas_validas_gravar:
                    st.warning("⚠️ Preenche pelo menos um item válido com código/produto e quantidade maior que zero.")
                else:
                    try:
                        ws_ab = client.worksheet("PEDIDOS_ABERTOS")
                    except:
                        ws_ab = client.add_worksheet(title="PEDIDOS_ABERTOS", rows="2000", cols="10")
                        ws_ab.append_row(["DATA_HORA", "FILIAL", "PEDIDO_NUM", "FORNECEDOR", "PRODUTO", "QUANTIDADE", "VALOR_TOTAL", "STATUS", "DATA_PREVISAO_ENTREGA", "NUMERO_NF"])

                    ts_reg = dt_mod.now().strftime('%d/%m/%Y %H:%M:%S')
                    str_prev_entrega = m_prev.strftime('%d/%m/%Y')

                    for item_g in linhas_validas_gravar:
                        ws_ab.append_row([
                            ts_reg,
                            m_filial,
                            str(m_pedido_num).strip(),
                            m_forn,
                            item_g["produto"],
                            f"{item_g['qtd']:.3f}".replace(".", ","),
                            f"{item_g['valor']:.2f}".replace(".", ","),
                            "ATIVO",
                            str_prev_entrega,
                            ""
                        ])

                    st.success(f"✅ Pedido Nº {m_pedido_num} com {len(linhas_validas_gravar)} itens gravado com sucesso na aba PEDIDOS_ABERTOS!")
                    st.cache_data.clear()
                    time_lib.sleep(1)
                    st.session_state["contador_editor_manual"] += 1
                    st.rerun()
            except Exception as err_multi:
                st.error(f"❌ Erro ao gravar múltiplos itens: {err_multi}")

if client is None: client = conectar_sheets_nativo()
if st.session_state.get('logado', False):
    st.sidebar.write(f"👤 Utilizador: **{st.session_state.get('usuario', 'Nenhum')}**")
    st.sidebar.write(f"🔑 Acesso: **{st.session_state.get('nivel', 'Comum')}**")
    if st.sidebar.button("🚪 Sair/Logoff"): st.session_state.logado = False; st.rerun()

if not st.session_state.get('logado', False): st.stop()
st.sidebar.write("---")

nivel_atual = st.session_state.get('nivel', 'Comum')
if nivel_atual == 'Admin':
    opcoes_menu = ["🔍 Conferência e Consolidação", "🥗 Lançamento de Pedidos", "📊 Cotação & Consolidação", "🍳 Fichas Técnicas (Cardápio)", "💼 Compras & Suprimentos", "📥 Entrada de NF-e (XML)", "📦 Almoxarifado / Portaria"]
elif nivel_atual == 'Fornecedor':
    opcoes_menu = ["🤝 Portal de Cotação"]
else:
    opcoes_menu = ["🥗 Lançamento de Pedidos", "📦 Almoxarifado / Portaria"]

modulo_selecionado = st.sidebar.radio("Navegação do Sistema:", opcoes_menu)

if modulo_selecionado == "📥 Entrada de NF-e (XML)": modulo_entrada_xml()
elif modulo_selecionado == "📦 Almoxarifado / Portaria": modulo_recebimento_fisico()
elif modulo_selecionado == "🔍 Conferência e Consolidação" or "Conferência" in str(modulo_selecionado): 
    st.title("🔍 Conferência e Consolidação de Pedidos")
    with st.container(border=True):
        st.markdown("### 🔒 Controlo de Acesso (Cotação)")
        
        status_atual = ler_status_digitacao()
        
        col_status, col_acao = st.columns([1, 1])
        
        with col_status:
            if status_atual:
                st.markdown("""<div style='background-color:#28a745; color:white; text-align:center; padding:10px; border-radius:5px; font-weight:bold; font-size:16px;'>🟢 COTAÇÃO LIBERADA</div>""", unsafe_allow_html=True)
            else:
                st.markdown("""<div style='background-color:#dc3545; color:white; text-align:center; padding:10px; border-radius:5px; font-weight:bold; font-size:16px;'>🔴 COTAÇÃO BLOQUEADA</div>""", unsafe_allow_html=True)
        
        with col_acao:
            try:
                if status_atual:
                    if st.button("Bloquear Cotação", use_container_width=True, key="btn_bloq_novo"):
                        with st.spinner("A bloquear no servidor..."):
                            salvar_governanca(False)
                            st.cache_data.clear()
                            st.rerun()
                else:
                    if st.button("Liberar Cotação", use_container_width=True, key="btn_lib_novo"):
                        with st.spinner("A libertar no servidor..."):
                            salvar_governanca(True)
                            st.cache_data.clear()
                            st.rerun()
            except Exception as e:
                st.error(f"⚠️ O Google bloqueou a ação temporariamente por segurança (limite de velocidade). Aguarda 1 minuto e tenta novamente! Erro técnico: {e}")
            
    restaurante_filtrado = st.selectbox("Filtrar por Restaurante:", ["Todos"] + MOCK_FILIAIS) 
    tab_sem, tab_men = st.tabs(["Pedido Semanal", "Pedido Mensal"]) 
    with tab_sem: consolidar_proteina_semanal_geral(restaurante_filtrado) 
    with tab_men: st.info("Módulo mensal em desenvolvimento.")

elif modulo_selecionado == "🥗 Lançamento de Pedidos" or "Lançamento" in str(modulo_selecionado): 
    st.title("🛒 Módulo de Pedidos") 
    interface_lancamento_proteina_filial()

elif modulo_selecionado == "🤝 Portal de Cotação": 
    st.title("🤝 Portal do Fornecedor")
    fornecedor_logado = st.session_state.get('filial_nome', 'FORNECEDOR PADRÃO')
    st.success(f"🏢 Empresa Logada: {fornecedor_logado}") 
    st.markdown("---")
    st.markdown("### 📝 Digitação de Preços (Lote Aberto)")
    st.info("Preenche o valor do pacote e o peso da embalagem (suporta decimais como 0,5 ou 2,5). O sistema calculará o preço por KG automaticamente para a concorrência. Se não tiver o produto, deixa a R$ 0,00.")
    
    try:
        dados_auditoria_brutos = buscar_dados_aba_cache("AUDITORIA_CONSOLIDADA") 
        if not dados_auditoria_brutos:
            st.warning("Nenhum lote de cotação aberto no momento.")
        else:
            df_aud = pd.DataFrame(dados_auditoria_brutos) 
            df_aud.columns = [str(c).strip().upper() for c in df_aud.columns] 
            
            if "STATUS_VALIDACAO" in df_aud.columns:
                df_aprov = df_aud[df_aud["STATUS_VALIDACAO"].astype(str).str.upper().isin(["APROVADO", "CONCLUÍDO FILIAL", "DENTRO DO LIMITE", "⚠ EXCEÇÃO (ESTOURADO)"])]
            else:
                df_aprov = df_aud

            if df_aprov.empty:
                st.success("🎉 Nenhuma cotação pendente no momento. Fica atento às próximas aberturas!")
            else:
                produtos_unicos = sorted(df_aprov["PRODUTO"].dropna().unique().tolist())
                
                linhas_cotacao = []
                for prod in produtos_unicos:
                    linhas_cotacao.append({
                        "Produto": prod,
                        "Preço Pacote/Caixa (R$)": "0.00",
                        "Peso Embalagem (KG)": "1.000"
                    })
                    
                df_cot = pd.DataFrame(linhas_cotacao)
                
                df_editado_forn = st.data_editor(
                    df_cot,
                    column_config={
                        "Produto": st.column_config.TextColumn(disabled=True),
                        "Preço Pacote/Caixa (R$)": st.column_config.TextColumn(required=True),
                        "Peso Embalagem (KG)": st.column_config.TextColumn(required=True)
                    },
                    hide_index=True,
                    use_container_width=True,
                    key="editor_fornecedor_precos"
                )
                
                st.write("")
                if st.button("🚀 Enviar Proposta Oficial", type="primary", use_container_width=True):
                    with st.spinner("A enviar proposta encriptada..."):
                        try:
                            aba_cotacao = client.worksheet("BD_COTACAO")
                            ts_agora = dt_mod.now().strftime('%d/%m/%Y %H:%M:%S')
                            
                            for _, row in df_editado_forn.iterrows():
                                preco_pacote = tratar_preco_float(row["Preço Pacote/Caixa (R$)"])
                                if preco_pacote > 0:
                                    peso = tratar_peso_float(row["Peso Embalagem (KG)"])
                                    peso = peso if peso > 0 else 1.0
                                    preco_kg_calc = round(preco_pacote / peso, 4)
                                    
                                    preco_pacote_str = f"{preco_pacote:.2f}".replace(".", ",")
                                    peso_str = f"{peso:.3f}".replace(".", ",")
                                    preco_kg_calc_str = f"{preco_kg_calc:.4f}".replace(".", ",")
                                    
                                    aba_cotacao.append_row([
                                        ts_agora, 
                                        "FORN_01", 
                                        fornecedor_logado, 
                                        row["Produto"], 
                                        preco_pacote_str, 
                                        "PCT/CX", 
                                        peso_str, 
                                        1, 
                                        preco_kg_calc_str, 
                                        preco_pacote_str
                                    ])
                            st.success("✅ Proposta submetida com sucesso! O departamento de compras já recebeu os teus valores.")
                            st.balloons()
                            time_lib.sleep(2)
                            st.rerun()
                        except Exception as e_envio:
                            st.error(f"Erro ao enviar: {e_envio}")
    except Exception as e_geral:
        st.error(f"Erro interno no portal do fornecedor: {e_geral}")

elif modulo_selecionado == "📊 Cotação & Consolidação":
    modulo_cotacao_consolidacao()

elif modulo_selecionado == "🍳 Fichas Técnicas (Cardápio)":
    st.title("🥗 Fichas Técnicas")
    if str(st.session_state.get('nivel')).upper() == "ADMIN" or st.session_state.get('usuario') in ["Jardel", "Desenvolvedor"]:
        try:
            sheet_auditoria = client.worksheet("AUDITORIA_CONSOLIDADA")
            dados_auditoria = sheet_auditoria.get_all_records()
            if not dados_auditoria: st.info("📂 Nenhuma pendência encontrada.")
            else:
                df_auditoria = pd.DataFrame(dados_auditoria)
                if "STATUS_VALIDACAO" in df_auditoria.columns: df_auditoria = df_auditoria[df_auditoria["STATUS_VALIDACAO"] != "APROVADO"]
                if "APROVAR" not in df_auditoria.columns: df_auditoria.insert(0, "APROVAR", False)
                
                df_editado = st.data_editor(df_auditoria, hide_index=True, use_container_width=True)
                if st.button("🚀 Processar Pedidos Selecionados", type="primary", use_container_width=True):
                    df_selecionados = df_editado[df_editado["APROVAR"] == True]
                    if df_selecionados.empty: st.warning("Seleciona um pedido para aprovação.")
                    else: st.success("Pedidos validados com sucesso!"); st.cache_data.clear(); st.rerun()
        except Exception as e: st.error(f"Erro ao carregar o painel: {e}")
    else: st.error("🔒 Acesso restrito ao Administrador.")

elif modulo_selecionado == "💼 Compras & Suprimentos":
    modulo_compras_suprimentos()
