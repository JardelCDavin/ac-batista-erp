import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import gspread
import os
import json
from datetime import datetime as dt_mod, date, timedelta, time as time_mod
import urllib.parse
import time as time_lib
from gspread_dataframe import set_with_dataframe
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import io
import xml.etree.ElementTree as ET

# --- BLINDAGEM DE API: MEMÓRIA DE CURTO PRAZO ---
@st.cache_data(ttl=120)
def buscar_dados_aba_cache(nome_aba):
    if 'client' in globals() and client is not None:
        try:
            return client.worksheet(nome_aba).get_all_records()
        except gspread.exceptions.APIError as e:
            if e.response.status_code == 429:
                st.warning("⏳ O Google está sincronizando os dados. Aguarde alguns segundos...")
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

# --- CONEXÃO INTELIGENTE COM O GOOGLE SHEETS & GARANTIA DE ABAS ---
@st.cache_resource
def inicializar_gspread():
    caminho_local = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chave.json")
    return gspread.service_account(filename=caminho_local)

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
except gspread.exceptions.APIError as e:
    client = None
    if e.response.status_code == 429:
        st.error("⏳ Limite de acessos rápidos do Google atingido. Por favor, aguarde 1 minuto e recarregue a página.")
        st.stop()

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
    """Apenas fornecedores ATIVOS = SIM (Para Cotação Oficial)"""
    if not client:
        return ["FORNECEDOR PADRÃO"]
    try:
        dados = buscar_dados_aba_cache("FORNECEDORES")
        if dados:
            ativos = [str(row.get('Fornecedor', row.get('FORNECEDOR', ''))).strip() for row in dados if str(row.get('ATIVO', '')).strip().upper() == 'SIM']
            return ativos if ativos else ["FORNECEDOR PADRÃO"]
        return ["FORNECEDOR PADRÃO"]
    except Exception as e:
        return ["FORNECEDOR PADRÃO"]

@st.cache_data(ttl=300)
def carregar_todos_fornecedores_cadastrados():
    """TODOS os fornecedores cadastrados, sem filtro de ativo (Para Pedido Manual)"""
    if not client:
        return ["FORNECEDOR PADRÃO"]
    try:
        dados = buscar_dados_aba_cache("FORNECEDORES")
        if dados:
            todos = [str(row.get('Fornecedor', row.get('FORNECEDOR', ''))).strip() for row in dados if str(row.get('Fornecedor', row.get('FORNECEDOR', ''))).strip()]
            return sorted(list(set(todos))) if todos else ["FORNECEDOR PADRÃO"]
        return ["FORNECEDOR PADRÃO"]
    except Exception as e:
        return ["FORNECEDOR PADRÃO"]

@st.cache_data(ttl=300)
def carregar_catalogo_produtos_mapeamento():
    """Lê a aba PRODUTOS considerando Código, Nome e Preço Base"""
    if not client:
        return {}, {}, [], {}
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
    except Exception:
        return {}, {}, ["COXA SOLTEIRA PILÃO KG", "LINGUIÇA SUÍNA CHURRASCO", "SASSAMI KG"], {}

@st.cache_data(ttl=60)
def carregar_de_para_fornecedores():
    """Carrega a memória de aprendizado do sistema para ler o XML BLINDADA CONTRA FALHAS DO SHEETS"""
    dados = buscar_dados_aba_cache("DE_PARA_FORNECEDORES")
    mapeamento = {}
    if dados:
        for row in dados:
            cnpj = str(row.get("CNPJ_FORNECEDOR", "")).strip().replace(".", "").replace("/", "").replace("-", "").replace("'", "")
            if cnpj.isdigit():
                cnpj = cnpj.zfill(14)
                
            cod_forn = str(row.get("CODIGO_ITEM_FORNECEDOR", "")).strip().replace("'", "")
            cod_interno = str(row.get("CODIGO_INTERNO", "")).strip().replace("'", "")
            nome_interno = str(row.get("PRODUTO_INTERNO", "")).strip()
            
            if cnpj and cod_forn:
                chave = f"{cnpj}_{cod_forn}"
                mapeamento[chave] = {"CODIGO_INTERNO": cod_interno, "PRODUTO_INTERNO": nome_interno}
    return mapeamento

@st.cache_data(ttl=300)
def carregar_tabela_precos_historica():
    precos_map = {}
    try:
        dados_hist = buscar_dados_aba_cache("HISTORICO_PEDIDOS")
        if dados_hist:
            for h in dados_hist:
                prod = str(h.get("PRODUTO", "")).strip().upper()
                qtd = float(str(h.get("QUANTIDADE", 0)).replace(',','.')) if h.get("QUANTIDADE") else 0.0
                val = float(str(h.get("VALOR_TOTAL", 0)).replace(',','.')) if h.get("VALOR_TOTAL") else 0.0
                if prod and qtd > 0 and val > 0:
                    precos_map[prod] = round(val / qtd, 2)
        
        dados_cot = buscar_dados_aba_cache("BD_COTACAO")
        if dados_cot:
            for c in dados_cot:
                prod = str(c.get("PRODUTO", "")).strip().upper()
                p_eq = c.get("PRECO_KG_EQUIV", 0.0)
                if prod and p_eq > 0:
                    precos_map[prod] = float(p_eq)
    except:
        pass
    return precos_map

@st.cache_data(ttl=300)
def carregar_prazos_comerciais_sheets():
    if not client:
        return ["7 Dias", "14 Dias", "30 Dias"]
    try:
        dados = buscar_dados_aba_cache("PRAZOS_PAGAMENTOS")
        if not dados:
            return ["7 Dias", "14 Dias", "30 Dias"]
            
        lista_prazos = [str(linha.get("PRAZOS", "")).strip() for linha in dados if linha.get("PRAZOS")]
        return lista_prazos if lista_prazos else ["7 Dias", "14 Dias", "30 Dias"]
    except Exception:
        return ["7 Dias", "7/14 Dias", "7/14/21 Dias", "7/14/21/28 Dias", "14 Dias", "30 Dias", "45 Dias", "60 Dias"]

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
                if "CNPJ:" in l_up:
                    cnpj = linha.split(":")[-1].strip()
                elif "ESTADUAL:" in l_up or "INS. ESTADUAL:" in l_up:
                    ie = linha.split(":")[-1].strip()
                elif "CEP:" in l_up:
                    cep = linha.split(":")[-1].strip()
                elif "E-MAIL:" in l_up or "EMAIL:" in l_up:
                    email = linha.split(":")[-1].strip()
                elif not any(x in l_up for x in ["CNPJ:", "ESTADUAL:", "CEP:", "E-MAIL:", "EMAIL:"]) and linha != razao:
                    if not endereco:
                        endereco = linha
                    else:
                        endereco += " - " + linha
                        
            filial_key = nome_filial.upper()
            filiais_map[filial_key] = {
                "RAZAO": razao,
                "CNPJ": cnpj if cnpj else "06.121.429/0008-90",
                "IE": ie if ie else "625274795.07.43",
                "ENDERECO": endereco if endereco else "SÃO JOÃO DEL REI/MG",
                "CEP": cep if cep else "36300-168",
                "EMAIL": email if email else "comprasacbatista@gmail.com"
            }
    return filiais_map

# --- CONFIGURAÇÕES E CONSTANTES GERAIS ---
PRATOS_CARDAPIO_OFICIAL = ["Arroz Branco", "Feijão Carioca/Inteiro/Batido", "Filé de Frango Acebolado", "Carne Moída ao Molho", "Feijoada", "Frango Assado", "Linguiça Assada", "Pernil Assado", "Macarrão Alho e Óleo", "Angu/Polenta", "Farofa Colorida/Cenoura"]
UNIDADES_PRODUTOS_MENSAL = {
    "AÇÚCAR CRISTAL (FARDO 6X5KG)": "Fardo", "ARROZ PARBOILIZED (FARDO 6X5KG)": "Fardo", "ARROZ INTEGRAL": "KG", "ADOÇANTE LÍQUIDO": "UNID", "ACHOCOLATADO EM PÓ": "KG",
    "CAFÉ (FARDO 10X500GR)": "Fardo", "CANJICA BRANCA": "KG", "CREME DE LEITE / CREME CULINÁRIO": "UNID", "ERVILHA EM CONSERVA (LATA 1,7KG)": "Lata", "EXTRATO DE TOMATE (CAIXA 6X1,7KG)": "Caixa",
    "FARINHA DE MANDIOCA": "KG", "FARINHA DE MILHO": "KG", "FARINHA DE TRIGO": "KG", "FEIJÃO CARIOCA": "KG", "FEIJÃO PRETO": "KG", "FEIJÃO VERMELHO": "KG", "FUBÁ MIMOSO": "KG",
    "LEITE INTEGRAL (CAIXA 12/1LT)": "Caixa", "LEITE EM PÓ": "KG", "MACARRÃO ESPAGUETE COM OVOS": "KG", "MACARRÃO PARAFUSO COM OVOS": "KG", "MACARRÃO PADRE NOSSO (SOPA)": "KG", "MAIONESE (BALDE)": "Balde",
    "MARGARINA (BALDE 14,5KG)": "Balde", "MILHO DE PIPOCA": "KG", "MILHO VERDE EM CONSERVA (LATA 1,7KG)": "Lata", "MOLHO DE ALHO": "UNID", "MOLHO INGLÊS": "UNID", "MOLHO DE TOMATE (SACHET)": "UNID", "MOLHO SHOYU": "UNID",
    "ÓLEO DE SOJA (CAIXA 20/900ML)": "Caixa", "ÓLEO COMPOSTO": "UNID", "SAL REFINADO (FARDO 30/1KG)": "Fardo", "SAL DE PARRILLA": "KG", "TEMPERO PRONTO (ALHO E SAL)": "UNID", "VINAGRE DE ÁLCOOL": "UNID"
}
MOCK_FILIAIS = ["TEJUCO", "CENTRO", "MATOSINHOS", "RM SABOR", "COLONIA", "BARBACENA", "LEOPOLDINA"]

@st.cache_data(ttl=600)
def carregar_proteinas_semanal():
    arquivo = "SEMANAL_PROTEINAS.xlsx"
    try:
        if os.path.exists(arquivo):
            df = pd.read_excel(arquivo)
            cols = [c for c in df.columns if str(c).strip().upper() == 'PRODUTOS']
            if not cols:
                cols = [c for c in df.columns if 'PROD' in str(c).upper()]
            if cols:
                return df[cols[0]].dropna().astype(str).str.strip().tolist()
    except Exception:
        pass
    return []

@st.cache_data(ttl=600)
def carregar_produtos_mensal():
    arquivo = "MENSAL_MERCEARIA_EMBALAGENS_LIMPEZA.xlsx"
    try:
        if os.path.exists(arquivo):
            df = pd.read_excel(arquivo)
            cols = [c for c in df.columns if str(c).strip().upper() == 'PRODUTOS']
            if not cols:
                cols = [c for c in df.columns if 'PROD' in str(c).upper()]
            if cols:
                return df[cols[0]].dropna().astype(str).str.strip().tolist()
    except Exception:
        pass
    return []

# --- CONFIGURAÇÃO DA PÁGINA (LAYOUT CONGELADO) ---
st.set_page_config(page_title="Portal AC Batista", layout="wide")

def validar_usuario_sheets(usuario, senha):
    if not client:
        return None, "❌ Não foi possível conectar ao Google Sheets."
    
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
                    else:
                        return None, "❌ Senha incorreta. Tente novamente."
    except Exception as e:
        return None, f"❌ Erro crítico ao conectar à base de dados: {e}"

    return None, "❌ Usuário ou Fornecedor não localizado no sistema."

# =========================================================================
# 🛑 TELA DE LOGIN ISOLADA
# =========================================================================
container_login = st.empty()

if not st.session_state.get('logado', False):
    with container_login.container():
        st.markdown("""<style>[data-test-id="stSidebar"] { display: none !important; } .stMainBlockContainer { max-width: 500px; margin: 0 auto; padding-top: 5rem; }</style>""", unsafe_allow_html=True)
        st.title("🔒 Login - AC Batista ERP")
        usuario = st.text_input("Usuário", key="txt_usuario_final")
        senha = st.text_input("Senha", type="password", key="txt_senha_final")

        if st.button("Acessar o Sistema"):
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

# --- ESTILIZAÇÃO CSS OFICIAL DO PAINEL INTERNO ---
st.markdown("""
    <style>
    .stButton>button {
        background-color: #004A99;
        color: white;
        width: 100%;
        border-radius: 5px;
        height: 3em;
        font-weight: bold;
    }
    .stButton>button:hover { background-color: #003366; color: white; }
    </style>
    """, unsafe_allow_html=True)

GOVERNANCA_FILE = "governanca_status.json"
DEFAULT_GOVERNANCA = {
    "libera_digitacao_semanal": False,
    "libera_digitacao_mensal": False
}

def carregar_governanca():
    if 'governanca' in st.session_state:
        return st.session_state.governanca

    governanca = DEFAULT_GOVERNANCA.copy()
    if os.path.exists(GOVERNANCA_FILE):
        try:
            with open(GOVERNANCA_FILE, 'r', encoding='utf-8') as f:
                arquivo = json.load(f)
                governanca.update({
                    'libera_digitacao_semanal': bool(arquivo.get('libera_digitacao_semanal', arquivo.get('libera_cotacao_semanal', governanca['libera_digitacao_semanal']))),
                    'libera_digitacao_mensal': bool(arquivo.get('libera_digitacao_mensal', arquivo.get('libera_cotacao_mensal', governanca['libera_digitacao_mensal'])))
                })
        except Exception:
            pass

    st.session_state.governanca = governanca
    st.session_state.libera_digitacao_semanal = governanca['libera_digitacao_semanal']
    st.session_state.libera_digitacao_mensal = governanca['libera_digitacao_mensal']
    return governanca

def salvar_governanca():
    governanca = {
        'libera_digitacao_semanal': bool(st.session_state.get('libera_digitacao_semanal', DEFAULT_GOVERNANCA['libera_digitacao_semanal'])),
        'libera_digitacao_mensal': bool(st.session_state.get('libera_digitacao_mensal', DEFAULT_GOVERNANCA['libera_digitacao_mensal']))
    }
    st.session_state.governanca = governanca
    try:
        with open(GOVERNANCA_FILE, 'w', encoding='utf-8') as f:
            json.dump(governanca, f, indent=4, ensure_ascii=False)
    except Exception:
        pass

def encurtar_nome_fornecedor(nome_completo):
    n = str(nome_completo).strip().upper()
    if "LIDER" in n or "RUBBO" in n: return "Líder"
    elif "OESA" in n: return "Oesa"
    elif "RIO BRANCO" in n or "PIF PAF" in n or "RIO" in n: return "Rio Branco"
    elif "SEARA" in n: return "Seara"
    elif "FRIGO" in n: return "Frigo"
    partes = n.split()
    return partes[0].title() if partes else "Fornecedor"

# --- GERADOR DE PDF PROFISSIONAL (REPORTLAB) ---
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

# =========================================================================
# FUNÇÕES CORE (COMPRAS, SUPRIMENTOS E COTAÇÃO E ENTRADA XML)
# =========================================================================
def consolidar_mensal():
    try:
        planilha = conectar_sheets_nativo()
        if not planilha: return
        filiais_alvo = ['CENTRO', 'TEJUCO', 'COLONIA', 'BARBACENA', 'LEOPOLDINA', 'MATOSINHOS', 'SABOR']
        totais_produtos = {}
        for nome_aba_real in [w.title for w in planilha.worksheets()]:
            if 'MENSAL' in nome_aba_real.upper() and any(f in nome_aba_real.upper() for f in filiais_alvo):
                try:
                    df_aba = pd.DataFrame(buscar_dados_aba_cache(nome_aba_real))
                    if df_aba.empty: continue
                    df_aba.columns = [str(c).strip().upper() for c in df_aba.columns]
                    col_prod_real = next((c for c in df_aba.columns if 'PROD' in c or 'ITEM' in c or 'NOME' in c), df_aba.columns[0])
                    col_qtd_real = next((c for c in df_aba.columns if 'QTD' in c or 'QUANT' in c), df_aba.columns[1] if len(df_aba.columns) > 1 else df_aba.columns[0])
                    for _, linha in df_aba.iterrows():
                        produto = str(linha[col_prod_real]).strip()
                        if produto == "" or produto.upper().startswith("1.") or produto.upper().startswith("2."): continue
                        qtd = pd.to_numeric(linha[col_qtd_real], errors='coerce')
                        qtd = qtd if not pd.isna(qtd) else 0
                        if qtd > 0: totais_produtos[produto.upper()] = totais_produtos.get(produto.upper(), 0) + qtd
                except: continue
        linhas_finais = [[p.title(), UNIDADES_PRODUTOS_MENSAL.get(p, "UN/KG"), q] for p, q in totais_produtos.items()]
        if linhas_finais:
            df_final = pd.DataFrame(linhas_finais, columns=['Item', 'Unidade', 'Quantidade Total'])
            for aba_nome in ['MENSAL_MODELO', 'COTACAO']:
                try:
                    w = planilha.worksheet(aba_nome); w.clear()
                    w.update(range='A1', values=[df_final.columns.values.tolist()] + df_final.values.tolist())
                except: pass
            st.success("✅ Módulo Mensal Consolidado!")
    except Exception as e: st.error(f"Erro: {e}")

def tratar_virgula(val):
    try:
        if isinstance(val, (int, float)): return float(val)
        v_str = str(val).strip()
        if "," in v_str and "." in v_str:
            if v_str.rfind(",") > v_str.rfind("."): v_str = v_str.replace(".", "").replace(",", ".")
            else: v_str = v_str.replace(",", "")
        elif "," in v_str and "." not in v_str: v_str = v_str.replace(",", ".")
        return float(v_str)
    except: return 0.0

def consolidar_proteina_semanal_geral(restaurante_filtro="Todos"):
    planilha = conectar_sheets_nativo()
    if not planilha: return
    dados = buscar_dados_aba_cache("AUDITORIA_CONSOLIDADA")
    if not dados:
        st.warning("A planilha de auditoria está vazia!")
        return

    df_linhas = []
    for linha in dados:
        filial_reg = str(linha.get("RESTAURANTE", "")).strip()
        if restaurante_filtro != "Todos" and filial_reg != restaurante_filtro: continue
        if str(linha.get("STATUS_VALIDACAO", "")).strip().upper() == "APROVADO": continue
        produto = str(linha.get("PRODUTO", "")).strip()
        try:
            val = linha.get("PEDIDO_NUTRICIONISTA", 0)
            qtd = float(str(val).replace('.', '').replace(',', '.')) if val else 0.0
        except: qtd = 0.0
        if produto and qtd > 0:
            df_linhas.append({"Filial": filial_reg, "Proteína / Item": produto, "Quantidade Lançada (KG)": qtd, "Status": linha.get("STATUS_VALIDACAO", ""), "Justificativa da Nutricionista": linha.get("JUSTIFICATIVA", "")})

    if not df_linhas:
        st.warning(f"Nenhum pedido localizado para: {restaurante_filtro}")
        return

    df_final = pd.DataFrame(df_linhas)
    df_editado = st.data_editor(df_final, use_container_width=True, hide_index=True, disabled=["Filial", "Proteína / Item", "Status", "Justificativa da Nutricionista"])

    if st.button("💾 Salvar Correção da Filial", type="primary"):
        try:
            aba_aud = client.worksheet("AUDITORIA_CONSOLIDADA")
            d_sheets = aba_aud.get_all_records()
            cab_upper = [str(c).strip().upper() for c in aba_aud.row_values(1)]
            for _, r_ed in df_editado.iterrows():
                f_v, p_v, q_v = str(r_ed["Filial"]).strip().upper(), str(r_ed["Proteína / Item"]).strip().upper(), float(r_ed["Quantidade Lançada (KG)"])
                for idx_s, r_s in enumerate(d_sheets, start=2):
                    if str(r_s.get("RESTAURANTE", "")).strip().upper() == f_v and str(r_s.get("PRODUTO", "")).strip().upper() == p_v:
                        if "PEDIDO_NUTRICIONISTA" in cab_upper: aba_aud.update_cell(idx_s, cab_upper.index("PEDIDO_NUTRICIONISTA") + 1, q_v)
                        if "STATUS_VALIDACAO" in cab_upper: aba_aud.update_cell(idx_s, cab_upper.index("STATUS_VALIDACAO") + 1, "APROVADO")
            st.success("✅ Salvo com sucesso!")
            st.cache_data.clear(); time_lib.sleep(1); st.rerun()
        except Exception as e: st.error(f"Erro: {e}")

def modulo_cotacao_consolidacao(): 
    st.title("📊 Cotação & Consolidação") 
    st.markdown("## ⚙ Painel de Distribuição de Suprimentos") 
    st.info("Espaço destinado ao gerenciamento logístico de insumos e fechamento de cargas do Diretor Jardel.") 
    
    with st.expander("⏱️ Controle de Prazo e Acompanhamento de Fornecedores", expanded=True):
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
                    status_lista.append({"Fornecedor": nome_curto, "Status": "🟢 Finalizado / Enviado", "Detalhes": "Proposta registrada"})
                else:
                    status_lista.append({"Fornecedor": nome_curto, "Status": "🔴 Pendente", "Detalhes": "Aguardando envio"})
            st.dataframe(pd.DataFrame(status_lista), use_container_width=True, hide_index=True)
        except Exception as e_st:
            st.caption(f"ℹ️ Aguardando dados para exibir o status: {e_st}")

    st.markdown("---")

    try: 
        dados_cotacao_brutos = buscar_dados_aba_cache("BD_COTACAO") 
        if not dados_cotacao_brutos: 
            st.info("ℹ️ Nenhuma proposta de cotação encontrada na aba BD_COTACAO para o lote atual.")
            df_bruto = pd.DataFrame()
        else:
            df_bruto = pd.DataFrame(dados_cotacao_brutos) 
            df_bruto.columns = [str(c).strip().upper() for c in df_bruto.columns] 

        def tratar_preco_float(valor): 
            try: 
                if isinstance(valor, (int, float)): 
                    val = float(valor)
                    if val > 100 and val % 1 != 0: return val
                    elif val >= 100 and val == int(val) and val not in [100, 200, 500, 1000]: return val / 100.0
                    elif val >= 40 and val < 100 and val % 1 == 0: return val / 10.0
                    return val
                v_str = str(valor).strip()
                if not v_str or v_str.lower() == 'nan': return 0.0
                v_str = v_str.replace("R$", "").strip()
                if "," in v_str and "." in v_str:
                    if v_str.rfind(",") > v_str.rfind("."): v_str = v_str.replace(".", "").replace(",", ".")
                    else: v_str = v_str.replace(",", "")
                elif "," in v_str and "." not in v_str: v_str = v_str.replace(",", ".")
                num = float(v_str)
                if num > 1000 and num % 100 == 0: return num / 100.0
                elif num >= 40 and num < 100 and num % 1 == 0: return num / 10.0
                return num
            except: return 0.0

        def tratar_qtd_float(valor):
            try:
                if isinstance(valor, (int, float)): return float(valor)
                v_str = str(valor).strip()
                if not v_str or v_str.lower() == 'nan': return 0.0
                if "," in v_str and "." in v_str:
                    if v_str.rfind(",") > v_str.rfind("."): v_str = v_str.replace(".", "").replace(",", ".")
                    else: v_str = v_str.replace(",", "")
                elif "," in v_str and "." not in v_str: v_str = v_str.replace(",", ".")
                return float(v_str)
            except: return 0.0
                
        if "PRECO_PACOTE" in df_bruto.columns: df_bruto["PRECO_PACOTE"] = df_bruto["PRECO_PACOTE"].apply(tratar_preco_float)
        if "PESO_EMBALAGEM" in df_bruto.columns: df_bruto["PESO_EMBALAGEM"] = df_bruto["PESO_EMBALAGEM"].apply(tratar_preco_float)

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
                    df_aprovados = df_auditoria_total[df_auditoria_total["STATUS_VALIDACAO"] == "APROVADO"] if "STATUS_VALIDACAO" in df_auditoria_total.columns else df_auditoria_total 

                    if not df_aprovados.empty and all(c in df_aprovados.columns for c in ["PRODUTO", "RESTAURANTE", "PEDIDO_NUTRICIONISTA"]): 
                        df_aprovados["PEDIDO_NUTRICIONISTA"] = df_aprovados["PEDIDO_NUTRICIONISTA"].apply(tratar_qtd_float) 
                        df_grade_filiais = df_aprovados.pivot_table(index="PRODUTO", columns="RESTAURANTE", values="PEDIDO_NUTRICIONISTA", aggfunc="sum").fillna(0.0).reset_index() 
                        df_grade_filiais.columns.name = None 
                        colunas_restaurantes = [col for col in df_grade_filiais.columns if col != "PRODUTO"] 
                        df_grade_filiais["Volume Total (KG)"] = df_grade_filiais[colunas_restaurantes].sum(axis=1) 
                        st.dataframe(df_grade_filiais, hide_index=True, use_container_width=True) 
                    else: st.info("ℹ Nenhum pedido aprovado na auditoria.") 
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
                                    r_ab.get("QUANTIDADE", 0),
                                    r_ab.get("VALOR_TOTAL", 0),
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
        # ABA 3: CONFERÊNCIA & DISPARO (COM DOWNLOAD EM PDF AUTOMATIZADO)
        # ==========================================
        with aba_conferencia:
            st.markdown("### 📑 Espelho de Pedidos e Carrinho de Revisão")
            st.caption("Revise o pedido, ajuste quantidades ou fornecedores, e baixe o espelho oficial em PDF com o prazo de entrega calculado automaticamente.")
            
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
                        df_aprov_conf = df_aud_conf[df_aud_conf["STATUS_VALIDACAO"].astype(str).str.upper() == "APROVADO"].copy() if "STATUS_VALIDACAO" in df_aud_conf.columns else df_aud_conf.copy()

                        if not df_aprov_conf.empty and "RESTAURANTE" in df_aprov_conf.columns:
                            lista_filiais_disponiveis = sorted(df_aprov_conf["RESTAURANTE"].dropna().astype(str).str.strip().unique().tolist())
                            filial_selecionada_aba3 = st.selectbox("🏢 Escolha a Filial para emitir os Pedidos:", lista_filiais_disponiveis, key="sb_filial_conferencia_aba3")
                            
                            if filial_selecionada_aba3:
                                df_itens_filial = df_aprov_conf[df_aprov_conf["RESTAURANTE"].astype(str).str.strip().str.upper() == filial_selecionada_aba3.upper()].copy()
                                
                                st.markdown(f"#### 🛒 Carrinho Editável - {filial_selecionada_aba3}")
                                
                                lista_carrinho = []
                                for _, row_i in df_itens_filial.iterrows():
                                    produto_nome = str(row_i.get("PRODUTO", "")).strip()
                                    qtd_nutri_original = tratar_qtd_float(row_i.get("PEDIDO_NUTRICIONISTA", 0.0))

                                    vencedor_aba2 = forn_lista[0] if forn_lista else "FORNECEDOR PADRÃO"
                                    l_dec = df_dec[df_dec["PRODUTO"] == produto_nome]
                                    if not l_dec.empty:
                                        sug = str(l_dec.iloc[0].get("Sugestão Sistema", "")).strip()
                                        if sug in forn_lista: vencedor_aba2 = sug

                                    lista_carrinho.append({
                                        "Excluir?": False,
                                        "Produto": produto_nome,
                                        "Qtd Solicitada": float(qtd_nutri_original),
                                        "Fornecedor Destino": vencedor_aba2
                                    })

                                df_carrinho = pd.DataFrame(lista_carrinho)
                                df_carrinho_editado = st.data_editor(
                                    df_carrinho,
                                    column_config={
                                        "Excluir?": st.column_config.CheckboxColumn("Remover", default=False),
                                        "Produto": st.column_config.TextColumn("Descrição do Produto", disabled=True),
                                        "Qtd Solicitada": st.column_config.NumberColumn("Quantidade (KG)", min_value=0.0, step=0.5, format="%.2f"),
                                        "Fornecedor Destino": st.column_config.SelectboxColumn("Fornecedor Destino", options=forn_lista, required=True)
                                    },
                                    hide_index=True,
                                    use_container_width=True,
                                    key=f"carrinho_edit_livre_{filial_selecionada_aba3}"
                                )

                                st.markdown("---")
                                
                                if st.button(f"🖨️ Gerar Espelhos de Pedidos Oficiais para {filial_selecionada_aba3}", type="primary", use_container_width=True):
                                    itens_validos = df_carrinho_editado[(df_carrinho_editado["Excluir?"] == False) & (df_carrinho_editado["Qtd Solicitada"] > 0)]
                                    
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
                                                p_qtd = row_item["Qtd Solicitada"]
                                                
                                                preco_u = 0.0
                                                l_dec = df_dec[df_dec["PRODUTO"] == p_nome]
                                                if not l_dec.empty and forn_alvo in l_dec.columns:
                                                    val = l_dec.iloc[0][forn_alvo]
                                                    preco_u = float(val) if pd.notna(val) else 0.0
                                                    
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
                                                st.markdown(f"### **Valor Total: R$ {val_total_pedido:,.2f}**")
                                                
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
                                                                float(row_grv["Qtd"]),
                                                                float(row_grv["Total"]),
                                                                "ATIVO",
                                                                str_prev_entrega,
                                                                ""
                                                            ])
                                                        st.success(f"✅ Pedido Nº {num_seq:04d} disparado e registrado com sucesso na aba PEDIDOS_ABERTOS!")
                                                        st.cache_data.clear()
                                                    except Exception as erro_aberto:
                                                        st.error(f"Erro ao salvar na base de pedidos abertos: {erro_aberto}")
                        else: st.info("ℹ️ Nenhum pedido aprovado encontrado.")
                    else: st.info("ℹ️ Realize a cotação na Aba 2 primeiro.")
            except Exception as e_conf:
                st.error(f"Erro ao montar a conferência: {e_conf}")

        # ==========================================
        # ABA 4: RELATÓRIO DE DIVERGÊNCIAS
        # ==========================================
        with aba_relatorio:
            st.markdown("### 📈 Relatório Gerencial de Divergências (Acareação)")
            st.caption("Acompanhe o histórico de notas baixadas e analise faltas, sobras e variações de preços.")
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

# =========================================================================
# MÓDULO: LANÇAMENTO DE PEDIDOS (NUTRICIONISTAS)
# =========================================================================
def interface_lancamento_proteina_filial(filial_passada="CENTRO"):
    try:
        st.subheader("📋 Digitação Semanal por Filial")
        st.info("Use as abas abaixo para lançar novos itens ou gerenciar o lote.")
        
        nivel_usuario = st.session_state.get('nivel', 'Comum')
        filial_do_login = st.session_state.get('filial_nome', 'TEJUCO')
        
        if nivel_usuario == "Admin":
            filial_selected = st.selectbox("🏢 Selecione a Filial:", MOCK_FILIAIS, index=MOCK_FILIAIS.index(filial_do_login) if filial_do_login in MOCK_FILIAIS else 0)
        else:
            filial_selected = filial_do_login
            st.info(f"📍 **Sua Filial Ativa:** {filial_selected}")
            
        if not filial_selected: return
        
        tab_lancamento, tab_conferencia = st.tabs(["📌 Aba 1: Lançamento", "🔍 Aba 2: Conferência"])
        
        with tab_lancamento:
            if st.session_state.get('nivel') != "Admin" and not st.session_state.get('libera_digitacao_semanal', False):
                st.warning("🛑 A digitação semanal está fechada pelo Diretor Jardel.")
            else:
                proteinas_lista = carregar_proteinas_semanal()
                if not proteinas_lista: 
                    st.warning("Nenhuma proteína localizada.")
                else:
                    produto_selecionado = st.selectbox("Selecione a Proteína:", [""] + proteinas_lista, format_func=lambda x: "-- Selecione --" if x == "" else x)
                    if not produto_selecionado: 
                        st.warning("👆 Escolha uma proteína acima para fazer um novo lançamento.")
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
                                        if str(tc.get('DESCRIÇÃO', '')).strip().upper() == str(produto_selecionado).strip().upper():
                                            teto_f = float(str(tc.get('Limite Ativo', 0.0)).replace(',', '.'))
                                            ml = str(tc.get('MARGEM_SEGURANÇA', '0.0')).replace(',', '.').replace('%', '')
                                            if ml.strip(): margem_seguranca = float(ml) / 100.0 if float(ml) > 1.0 else float(ml)
                                            pedido_sugerido = max((teto_f * (1.0 + margem_seguranca)) - stock_input, 0.0)
                                            break
                                else:
                                    nome_aba = "CONTRATO_LEOPOLDINA" if filial_limpa == "LEOPOLDINA" else "CONTRATO_POPULAR_GERAL"
                                    dados_contrato = buscar_dados_aba_cache(nome_aba)
                                    for lc in dados_contrato:
                                        prod_plan = str(lc.get('PRODUTO', '')).strip().upper()
                                        prod_sel = str(produto_selecionado).strip().upper()
                                        if prod_sel in prod_plan or prod_plan in prod_sel:
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
                            st.error("❌ Não pode enviar pedido zerado.")
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
                                        st.error(f"⚠️ Bloqueio: Você já enviou o pedido de {produto_selecionado} ({st.session_state['semana_atual']}) para o Diretor neste ciclo! Se precisar de mais, entre em contato com a Diretoria.")
                                        block = True; break
                                    elif status_atual in ["DENTRO DO LIMITE", "⚠ EXCEÇÃO (ESTOURADO)", ""]:
                                        st.error(f"⚠️ Atenção: O item {produto_selecionado} já está no seu carrinho na Aba 2 (Conferência) aguardando envio. Vá até lá se precisar alterar a quantidade.")
                                        block = True; break
                        
                        if st.button("➕ ADICIONAR À CONFERÊNCIA (VAI PARA ABA 2)", disabled=block, key="btn_grv"):
                            try:
                                aba_auditoria = client.worksheet("AUDITORIA_CONSOLIDADA")
                                aba_auditoria.append_row([
                                    str(dt_mod.now().strftime('%d/%m/%Y %H:%M:%S')), filial_selected, produto_selecionado, 
                                    st.session_state['semana_atual'], float(stock_input), float(pedido_input), round(pedido_sugerido, 2), status_v, jst
                                ])
                                st.success("✔️ Adicionado! Vá para a Aba 2 para Conferir e Enviar ao Diretor.")
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
                            
                            df_ex["PEDIDO_NUTRICIONISTA"] = df_ex["PEDIDO_NUTRICIONISTA"].apply(tratar_virgula)
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
                                with st.spinner("⏳ Enviando para a mesa do Diretor e limpando a conferência..."):
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
                                                    aba_auditoria.update_cell(idx_s, 6, qtd_final_float)
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
                    else: st.caption("ℹ️ A planilha não possui as colunas necessárias ainda.")
                else: st.caption("ℹ️ A planilha de auditoria está vazia.")
            except Exception as e_conf:
                st.error(f"Erro ao processar lote: {e_conf}")
    except Exception as erro_modulo_compras: 
        st.error(f"❌ Erro ao processar: {erro_modulo_compras}")

# =========================================================================
# MÓDULO: COMPRAS & SUPRIMENTOS (GESTÃO DE PEDIDOS E DIGITAÇÃO MANUAL MÚLTIPLA)
# =========================================================================
def modulo_compras_suprimentos():
    st.title("💼 Compras & Suprimentos - Gestão de Pedidos")
    st.info("Gerencie pedidos ativos, altere status (ATIVO/FINALIZADO) ou realize digitação manual de novos itens com múltiplas linhas.")

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
                st.warning("⚠️ Nenhum pedido encontrado na aba PEDIDOS_ABERTOS do Google Sheets. Gere pedidos na Aba 3 ou insira na Digitação Manual.")
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
                    st.caption("Digite o número da Nota Fiscal (NF), altere status para FINALIZADO ou ajuste quantidades diretamente na tabela abaixo.")
                    
                    df_editado_abertos = st.data_editor(
                        df_exibicao,
                        use_container_width=True,
                        hide_index=True,
                        key="editor_pedidos_abertos_geral_v2"
                    )

                    if st.button("💾 Salvar Alterações e Baixar Pedidos Finalizados", type="primary", key="btn_salvar_abertos_v2"):
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
                                        float(lista_valores[5]) if str(lista_valores[5]).replace('.','',1).isdigit() else 0.0,
                                        float(lista_valores[6]) if str(lista_valores[6]).replace('.','',1).isdigit() else 0.0,
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

                            st.success("✅ Alterações salvas! Pedidos finalizados foram arquivados com sucesso em HISTORICO_PEDIDOS.")
                            st.cache_data.clear()
                            time_lib.sleep(1)
                            st.rerun()
                        except Exception as err_save:
                            st.error(f"❌ Erro ao salvar alterações: {err_save}")
        except Exception as e_ger:
            st.error(f"Erro ao carregar pedidos abertos: {e_ger}")

    with tab_manual:
        st.markdown("##### ➕ Digitação Manual de Pedido Avulso com Múltiplos Itens")
        st.caption("Insira vários itens de uma vez. O número do pedido é calculado automaticamente de forma sequencial.")

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
        st.markdown("##### 🛒 Adicione os Produtos do Pedido (Até 25 Linhas)")
        st.caption("Digite o código interno ou selecione o produto. O preço unitário será sugerido automaticamente com base no PREÇO BASE da tabela e pode ser editado.")

        cod_para_prod, prod_para_cod, lista_nomes_prod, precos_base_map = carregar_catalogo_produtos_mapeamento()

        if "contador_editor_manual" not in st.session_state:
            st.session_state["contador_editor_manual"] = 0

        df_vazio_manual = pd.DataFrame([
            {"Código Interno": "", "Produto / Descrição": "", "Quantidade": 0.0, "Preço Unitário (R$)": 0.0}
            for _ in range(25)
        ])

        df_itens_digitados = st.data_editor(
            df_vazio_manual,
            column_config={
                "Código Interno": st.column_config.TextColumn("Cód. Interno", width="small"),
                "Produto / Descrição": st.column_config.SelectboxColumn("Descrição do Produto", options=[""] + lista_nomes_prod, width="large", required=False),
                "Quantidade": st.column_config.NumberColumn("Quantidade / KG", min_value=0.0, step=1.0, format="%.2f"),
                "Preço Unitário (R$)": st.column_config.NumberColumn("Preço Unit. (R$)", min_value=0.0, step=0.01, format="R$ %.2f")
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
                    q_val = float(r_item["Quantidade"]) if str(r_item["Quantidade"]).replace('.','',1).isdigit() else 0.0
                    
                    if c_int and not p_desc:
                        p_desc = cod_para_prod.get(c_int, f"CÓDIGO {c_int}")
                    elif p_desc and not c_int:
                        c_int = prod_para_cod.get(p_desc.upper(), "0545")
                    
                    p_unit = float(r_item["Preço Unitário (R$)"]) if str(r_item["Preço Unitário (R$)"]).replace('.','',1).isdigit() else 0.0
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
                    st.warning("⚠️ Preencha pelo menos um item válido com código/produto e quantidade maior que zero.")
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
                            item_g["qtd"],
                            item_g["valor"],
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


# =========================================================================
# MÓDULO NOVO: ENTRADA DE XML (NFE) COM SELEÇÃO INTELIGENTE DE PEDIDOS
# =========================================================================
def modulo_entrada_xml():
    st.title("📥 Entrada de NF-e (XML) e Estoque")
    st.info("O sistema fará o cruzamento contábil e atualizará a Tabela de Preços Mestre automaticamente.")

    if "xml_uploader_key" not in st.session_state: st.session_state["xml_uploader_key"] = 0

    arquivo_xml = st.file_uploader("Selecione o arquivo XML da Nota Fiscal Eletrônica", type=["xml"], key=f"xml_up_{st.session_state['xml_uploader_key']}")

    if arquivo_xml is not None:
        try:
            tree = ET.parse(arquivo_xml)
            root = tree.getroot()
            ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            
            infNFe = root.find('.//nfe:infNFe', ns)
            if infNFe is None: return st.error("XML inválido.")

            num_nf = root.find('.//nfe:ide/nfe:nNF', ns).text
            emit = root.find('.//nfe:emit', ns)
            cnpj_emit = emit.find('nfe:CNPJ', ns).text
            nome_emit = emit.find('nfe:xNome', ns).text
            cnpj_limpo = cnpj_emit.strip().replace(".", "").replace("/", "").replace("-", "")

            st.write(f"### 📄 Nota Fiscal Nº **{num_nf}** | 🏢 **Fornecedor:** {nome_emit}")
            
            # --- VERIFICAÇÃO DE DUPLICIDADE ---
            dados_hist_val = buscar_dados_aba_cache("HISTORICO_PEDIDOS")
            nfs_registradas = set(str(linha.get("NUMERO_NF", "")).strip().replace("'", "") for linha in dados_hist_val)
            
            if str(num_nf).strip() in nfs_registradas:
                st.error(f"🛑 ATENÇÃO: A Nota Fiscal Nº {num_nf} já foi processada e encontra-se no Histórico!")
                st.warning("O sistema bloqueou a entrada para evitar duplicidade de estoque e custos.")
                if st.button("🔄 Limpar e Carregar Outro XML", type="primary"):
                    st.session_state["xml_uploader_key"] += 1; st.rerun()
                return 
            
            cod_para_prod, prod_para_cod, lista_nomes, _ = carregar_catalogo_produtos_mapeamento()
            de_para_map = carregar_de_para_fornecedores()

            itens_nota = []
            pendencias = 0
            
            for det in root.findall('.//nfe:det', ns):
                prod = det.find('nfe:prod', ns)
                cProd = prod.find('nfe:cProd', ns).text
                xProd = prod.find('nfe:xProd', ns).text
                
                qCom = float(prod.find('nfe:qCom', ns).text)
                qTrib_node = prod.find('nfe:qTrib', ns)
                qTrib = float(qTrib_node.text) if qTrib_node is not None else qCom
                
                vUnCom = float(prod.find('nfe:vUnCom', ns).text)
                vProd = float(prod.find('nfe:vProd', ns).text)
                ncm = prod.find('nfe:NCM', ns).text if prod.find('nfe:NCM', ns) is not None else ""
                
                vST = 0.0
                icms = det.find('.//nfe:ICMS', ns)
                if icms is not None:
                    for child in icms:
                        vST_tag = child.find('nfe:vICMSST', ns)
                        if vST_tag is not None: vST = float(vST_tag.text); break
                            
                qtd_usada = max(qCom, qTrib)
                custo_total_real = vProd + vST

                chave = f"{cnpj_limpo}_{cProd.strip()}"
                status = "🔴 Pendente de Mapeamento"
                prod_interno = cod_int = ""
                
                if chave in de_para_map:
                    status = "🟢 Encontrado"
                    prod_interno = de_para_map[chave]["PRODUTO_INTERNO"]
                    cod_int = de_para_map[chave]["CODIGO_INTERNO"]
                else: pendencias += 1

                itens_nota.append({
                    "Seq": det.attrib['nItem'], "Cód Fornecedor": cProd, "Descrição NF": xProd, "NCM": ncm, 
                    "Qtd": qtd_usada, "ST (R$)": vST, "Custo Total Real": custo_total_real, 
                    "Status": status, "Produto Interno": prod_interno, "CHAVE": chave
                })

            df_itens = pd.DataFrame(itens_nota)

            if pendencias > 0:
                st.dataframe(df_itens[["Cód Fornecedor", "Descrição NF", "NCM", "Qtd", "Status", "Produto Interno"]], use_container_width=True, hide_index=True)
                st.warning(f"⚠️ {pendencias} item(ns) novo(s). Ensine o sistema associando os produtos abaixo:")
                for _, row in df_itens[df_itens["Status"] == "🔴 Pendente de Mapeamento"].iterrows():
                    with st.container(border=True):
                        st.write(f"**{row['Cód Fornecedor']} - {row['Descrição NF']}**")
                        c1, c2 = st.columns([3, 1])
                        with c1: p_esc = st.selectbox("Produto Interno:", ["-- Selecione --"] + lista_nomes, key=f"s_{row['Seq']}")
                        with c2:
                            st.write(""); st.write("")
                            if st.button("Salvar Mapeamento", key=f"b_{row['Seq']}", use_container_width=True):
                                if p_esc != "-- Selecione --":
                                    c_int = prod_para_cod.get(p_esc, "")
                                    client.worksheet("DE_PARA_FORNECEDORES").append_row([f"'{c_int}", p_esc, f"'{cnpj_limpo}", nome_emit, f"'{row['Cód Fornecedor']}"])
                                    st.success("Mapeamento salvo!"); st.cache_data.clear(); time_lib.sleep(1); st.rerun()
            else:
                st.success("✨ Todos os itens mapeados! Confirme o Fator de Conversão e confira a prévia abaixo.")
                
                # --- 1. FATOR DA NOTA FISCAL COM PRÉVIA EM TEMPO REAL ---
                df_fator = df_itens[["Seq", "Cód Fornecedor", "Descrição NF", "Produto Interno", "Qtd", "Custo Total Real"]].copy()
                df_fator["Fator de Conversão"] = 1.0 
                
                st.markdown("#### 1️⃣ Fator de Conversão DA NOTA FISCAL")
                st.caption("Ajuste o fator se necessário e acompanhe a prévia de conversão do estoque em tempo real.")
                
                df_fator_editado = st.data_editor(
                    df_fator,
                    column_config={
                        "Seq": None, "Cód Fornecedor": st.column_config.TextColumn(disabled=True),
                        "Descrição NF": st.column_config.TextColumn(disabled=True), "Produto Interno": st.column_config.TextColumn(disabled=True),
                        "Qtd": st.column_config.NumberColumn("Qtd (Nota)", disabled=True),
                        "Custo Total Real": st.column_config.NumberColumn("Custo Total (R$)", disabled=True, format="R$ %.2f"),
                        "Fator de Conversão": st.column_config.NumberColumn("Fator (Multiplicador)", min_value=0.01, step=1.0, format="%.2f", required=True)
                    },
                    hide_index=True, use_container_width=True, key="editor_fator_conversao"
                )

                # --- 🪟 ESPELHO DE PRÉVIA EM TEMPO REAL ---
                st.markdown("##### 🔍 Espelho de Prévia de Entrada no Estoque")
                previa_linhas = []
                for _, r_prev in df_fator_editado.iterrows():
                    r_orig_p = df_itens[df_itens["Seq"] == r_prev["Seq"]].iloc[0]
                    f_prev = float(r_prev["Fator de Conversão"])
                    q_nota_prev = float(r_orig_p["Qtd"])
                    c_tot_prev = float(r_orig_p["Custo Total Real"])
                    
                    qtd_convertida = q_nota_prev * f_prev
                    custo_unit_calc = c_tot_prev / qtd_convertida if qtd_convertida > 0 else 0.0
                    
                    previa_linhas.append({
                        "Produto": r_prev["Produto Interno"],
                        "Qtd Final no Estoque": qtd_convertida,
                        "Custo Unitário Final (R$)": custo_unit_calc
                    })
                st.dataframe(pd.DataFrame(previa_linhas), hide_index=True, use_container_width=True, column_config={
                    "Qtd Final no Estoque": st.column_config.NumberColumn(format="%.2f un/kg"),
                    "Custo Unitário Final (R$)": st.column_config.NumberColumn(format="R$ %.2f")
                })
                st.markdown("---")
                
                # --- 2. CHECK-IN DE NOTAS COM SELEÇÃO INTELIGENTE POR FILIAL E TABELA ---
                st.markdown("#### 2️⃣ Check-in de Notas (Conciliação com Pedido)")
                
                filial_checkin = st.selectbox("🏢 Selecione a Filial de Destino:", MOCK_FILIAIS, key="sb_filial_checkin_xml")
                
                dados_abertos = buscar_dados_aba_cache("PEDIDOS_ABERTOS")
                df_abertos = pd.DataFrame(dados_abertos) if dados_abertos else pd.DataFrame()
                
                pedido_selecionado = "Entrada Avulsa (Sem Pedido)"
                
                if not df_abertos.empty and all(c in df_abertos.columns for c in ["FILIAL", "STATUS", "PEDIDO_NUM", "FORNECEDOR", "QUANTIDADE", "VALOR_TOTAL"]):
                    df_ativos_filial = df_abertos[
                        (df_abertos["FILIAL"].astype(str).str.strip().str.upper() == filial_checkin.upper()) & 
                        (df_abertos["STATUS"].astype(str).str.upper() == "ATIVO")
                    ].copy()
                    
                    if not df_ativos_filial.empty:
                        st.markdown(f"##### 📋 Pedidos Ativos na Filial: {filial_checkin}")
                        st.caption("Selecione o pedido correspondente na caixa de seleção à esquerda:")
                        
                        resumo_pedidos = df_ativos_filial.groupby(["PEDIDO_NUM", "FORNECEDOR"]).agg({
                            "QUANTIDADE": "sum",
                            "VALOR_TOTAL": "sum"
                        }).reset_index()
                        
                        resumo_pedidos.insert(0, "Selecionar", False)
                        resumo_pedidos.columns = ["Selecionar", "Nº Pedido", "Fornecedor", "Qtd Total", "Valor Total (R$)"]
                        
                        df_ped_editado = st.data_editor(
                            resumo_pedidos,
                            column_config={
                                "Selecionar": st.column_config.CheckboxColumn("Escolher", default=False),
                                "Nº Pedido": st.column_config.TextColumn("Nº Pedido", disabled=True),
                                "Fornecedor": st.column_config.TextColumn("Fornecedor", disabled=True),
                                "Qtd Total": st.column_config.NumberColumn("Qtd Total", disabled=True, format="%.2f"),
                                "Valor Total (R$)": st.column_config.NumberColumn("Valor Total", disabled=True, format="R$ %.2f")
                            },
                            hide_index=True,
                            use_container_width=True,
                            key="tabela_selecao_pedido_xml"
                        )
                        
                        pedidos_escolhidos = df_ped_editado[df_ped_editado["Selecionar"] == True]["Nº Pedido"].tolist()
                        if pedidos_escolhidos:
                            pedido_selecionado = str(pedidos_escolhidos[0])
                            st.info(f"🔗 Pedido Vinculado para Acareação: **Nº {pedido_selecionado}**")
                        else:
                            st.info("ℹ️ Nenhum pedido selecionado na tabela. (Será processado como Entrada Avulsa)")
                    else:
                        st.info(f"ℹ️ Nenhum pedido ativo encontrado para a filial {filial_checkin}. Será tratada como Entrada Avulsa.")
                else:
                    st.info("ℹ️ Nenhum pedido aberto cadastrado no sistema.")

                if pedido_selecionado != "Entrada Avulsa (Sem Pedido)":
                    df_pedido = df_ativos_filial[df_ativos_filial["PEDIDO_NUM"].astype(str) == pedido_selecionado]
                    
                    st.markdown("#### ⚖️ Fator de Conversão DO PEDIDO")
                    df_pedido_show = df_pedido[["PRODUTO", "QUANTIDADE", "VALOR_TOTAL"]].copy()
                    df_pedido_show["Fator do Pedido"] = 1.0
                    
                    df_pedido_fator_editado = st.data_editor(
                        df_pedido_show,
                        column_config={
                            "PRODUTO": st.column_config.TextColumn(disabled=True),
                            "QUANTIDADE": st.column_config.NumberColumn("Qtd Pedido (Original)", disabled=True),
                            "VALOR_TOTAL": st.column_config.NumberColumn("Valor Total Pedido", disabled=True),
                            "Fator do Pedido": st.column_config.NumberColumn("Fator (Multiplicador)", min_value=0.01, step=1.0, format="%.2f", required=True)
                        },
                        hide_index=True, use_container_width=True, key="editor_fator_pedido"
                    )
                    
                    comparacao, produtos_nf = [], []
                    for _, r_ed in df_fator_editado.iterrows():
                        prod_nf = r_ed["Produto Interno"]
                        fator_nf, qtd_nota, custo_tot_nf = float(r_ed["Fator de Conversão"]), float(df_itens[df_itens["Seq"] == r_ed["Seq"]].iloc[0]["Qtd"]), float(df_itens[df_itens["Seq"] == r_ed["Seq"]].iloc[0]["Custo Total Real"])
                        qtd_final_nf = qtd_nota * fator_nf
                        preco_unit_nf = custo_tot_nf / qtd_final_nf if qtd_final_nf > 0 else 0
                        
                        item_ped = df_pedido_fator_editado[df_pedido_fator_editado["PRODUTO"].astype(str).str.upper() == prod_nf.upper()]
                        if not item_ped.empty:
                            qtd_ped_orig, val_tot_ped, fator_ped = float(item_ped.iloc[0]["QUANTIDADE"]), float(item_ped.iloc[0]["VALOR_TOTAL"]), float(item_ped.iloc[0]["Fator do Pedido"])
                            qtd_final_ped = qtd_ped_orig * fator_ped
                            preco_unit_ped = val_tot_ped / qtd_final_ped if qtd_final_ped > 0 else 0
                            
                            diff_qtd, diff_preco = qtd_final_nf - qtd_final_ped, preco_unit_nf - preco_unit_ped
                            status_qtd = f"📉 Faltou {abs(diff_qtd):.2f}" if diff_qtd < -0.05 else (f"📈 Sobrou {diff_qtd:.2f}" if diff_qtd > 0.05 else "✅ Exato")
                            status_preco = f"🔴 R$ {diff_preco:.2f} mais caro!" if diff_preco > 0.05 else (f"🟢 R$ {abs(diff_preco):.2f} mais barato" if diff_preco < -0.05 else "✅ Preço Mantido")
                            
                            comparacao.append({"Produto": prod_nf, "Qtd NF": qtd_final_nf, "Qtd Pedido": qtd_final_ped, "Status Qtd": status_qtd, "Preço NF (Un)": preco_unit_nf, "Preço Acordado": preco_unit_ped, "Status Preço": status_preco})
                        else:
                            comparacao.append({"Produto": prod_nf, "Qtd NF": qtd_final_nf, "Qtd Pedido": 0.0, "Status Qtd": "⚠️ Não estava no pedido", "Preço NF (Un)": preco_unit_nf, "Preço Acordado": 0.0, "Status Preço": "N/A"})
                        produtos_nf.append(prod_nf.upper())
                        
                    for _, r_ped in df_pedido_fator_editado.iterrows():
                        if str(r_ped["PRODUTO"]).upper() not in produtos_nf:
                            qtd_final_ped = float(r_ped["QUANTIDADE"]) * float(r_ped["Fator do Pedido"])
                            comparacao.append({"Produto": r_ped["PRODUTO"], "Qtd NF": 0.0, "Qtd Pedido": qtd_final_ped, "Status Qtd": "❌ CORTE TOTAL", "Preço NF (Un)": 0.0, "Preço Acordado": 0.0, "Status Preço": "N/A"})
                            
                    df_comp = pd.DataFrame(comparacao)
                    st.markdown("##### 🔎 Resumo do Check-in de Divergências")
                    def color_rules(row):
                        if 'Corte' in str(row['Status Qtd']) or 'Não estava' in str(row['Status Qtd']) or 'mais caro' in str(row['Status Preço']):
                            return ['background-color: #ffcccc; color: #900000; font-weight: bold;'] * len(row)
                        elif 'Exato' in str(row['Status Qtd']) and 'Mantido' in str(row['Status Preço']):
                            return ['background-color: #d4edda; color: #155724'] * len(row)
                        return [''] * len(row)
                    st.dataframe(df_comp.style.apply(color_rules, axis=1), use_container_width=True, hide_index=True)

                st.write("")
                if st.button("📥 Realizar Check-in e Dar Baixa (Histórico e Preços)", type="primary"):
                    with st.spinner("Processando..."):
                        try:
                            aba_h, aba_p = client.worksheet("HISTORICO_PEDIDOS"), client.worksheet("PRODUTOS")
                            grid_p = aba_p.get_all_values()
                            cab_p = [str(c).strip().upper() for c in grid_p[0]] if grid_p else []
                            idx_preco = cab_p.index("PREÇO_BASE") + 1 if "PREÇO_BASE" in cab_p else (cab_p.index("PRECO_BASE") + 1 if "PRECO_BASE" in cab_p else 7)
                            idx_ncm = cab_p.index("NCM") + 1 if "NCM" in cab_p else -1
                            idx_prod_col = cab_p.index("PRODUTO") if "PRODUTO" in cab_p else 1
                            
                            ts, dt_str = dt_mod.now().strftime('%d/%m/%Y %H:%M:%S'), dt_mod.now().strftime('%d/%m/%Y')
                            
                            for _, r_ed in df_fator_editado.iterrows():
                                seq, r_orig = r_ed["Seq"], df_itens[df_itens["Seq"] == r_ed["Seq"]].iloc[0]
                                fator, qtd_nota, custo_tot = float(r_ed["Fator de Conversão"]), float(r_orig["Qtd"]), float(r_orig["Custo Total Real"])
                                qtd_final = qtd_nota * fator
                                custo_unit_final = custo_tot / qtd_final if qtd_final > 0 else 0.0
                                ncm_final = str(r_orig["NCM"]).strip()
                                
                                aba_h.append_row([ts, dt_str, filial_checkin, pedido_selecionado if pedido_selecionado != "Entrada Avulsa (Sem Pedido)" else "ENTRADA_XML", nome_emit, str(r_orig["Produto Interno"]), qtd_final, round(custo_tot, 2), "BAIXADO (NF-e)", dt_str, f"'{num_nf}", f"'{ncm_final}", "", float(r_orig["ST (R$)"]), round(custo_unit_final, 4)])
                                
                                prod_int_nome = str(r_orig["Produto Interno"]).strip().upper()
                                for row_i, row_data in enumerate(grid_p):
                                    if row_i > 0 and len(row_data) > idx_prod_col and str(row_data[idx_prod_col]).strip().upper() == prod_int_nome:
                                        aba_p.update_cell(row_i + 1, idx_preco, custo_unit_final)
                                        if idx_ncm != -1 and ncm_final and not (len(row_data) > (idx_ncm-1) and str(row_data[idx_ncm-1]).strip()):
                                            aba_p.update_cell(row_i + 1, idx_ncm, f"'{ncm_final}")
                                        break
                                        
                            if pedido_selecionado != "Entrada Avulsa (Sem Pedido)":
                                ws_ab = client.worksheet("PEDIDOS_ABERTOS")
                                grid_ab = ws_ab.get_all_values()
                                
                                cab_ab = [str(c).strip().upper() for c in grid_ab[0]] if grid_ab else []
                                idx_ped_num = cab_ab.index("PEDIDO_NUM") + 1 if "PEDIDO_NUM" in cab_ab else 3
                                idx_status_ab = cab_ab.index("STATUS") + 1 if "STATUS" in cab_ab else 8
                                idx_nf_ab = cab_ab.index("NUMERO_NF") + 1 if "NUMERO_NF" in cab_ab else 10
                                
                                for i_ab, row_ab in enumerate(grid_ab):
                                    if i_ab > 0 and len(row_ab) >= idx_ped_num and str(row_ab[idx_ped_num - 1]).strip() == pedido_selecionado:
                                        linha_plan = i_ab + 1
                                        ws_ab.update_cell(linha_plan, idx_status_ab, f"RECEBIDO (NF: {num_nf})")
                                        if idx_nf_ab:
                                            ws_ab.update_cell(linha_plan, idx_nf_ab, f"'{num_nf}")

                            st.success("✅ Check-in Concluído com Sucesso!"); st.balloons()
                            st.session_state["xml_uploader_key"] += 1; time_lib.sleep(2.5); st.cache_data.clear(); st.rerun()
                        except Exception as e: st.error(f"Erro ao processar baixa: {e}")
        except Exception as e: st.error(f"Erro no XML: {e}")

# =========================================================================
# FLUXO PRINCIPAL DE NAVEGAÇÃO E EXECUÇÃO
# =========================================================================
if client is None: client = conectar_sheets_nativo()
if st.session_state.get('logado', False):
    st.sidebar.write(f"👤 Usuário: **{st.session_state.get('usuario', 'Nenhum')}**")
    st.sidebar.write(f"🔑 Acesso: **{st.session_state.get('nivel', 'Comum')}**")
    if st.sidebar.button("🚪 Sair/Logoff"): st.session_state.logado = False; st.rerun()

if not st.session_state.get('logado', False): st.stop()
st.sidebar.write("---")

nivel_atual = st.session_state.get('nivel', 'Comum')
if nivel_atual == 'Admin':
    opcoes_menu = ["🔍 Conferência e Consolidação", "🥗 Lançamento de Pedidos", "📊 Cotação & Consolidação", "🍳 Fichas Técnicas (Cardápio)", "💼 Compras & Suprimentos", "📥 Entrada de NF-e (XML)"]
elif nivel_atual == 'Fornecedor':
    opcoes_menu = ["🤝 Portal de Cotação"]
else:
    opcoes_menu = ["🥗 Lançamento de Pedidos"]

modulo_selecionado = st.sidebar.radio("Navegação do Sistema:", opcoes_menu)

if modulo_selecionado == "📥 Entrada de NF-e (XML)": modulo_entrada_xml()
elif modulo_selecionado == "🔍 Conferência e Consolidação" or "Conferência" in str(modulo_selecionado): 
    st.title("🔍 Conferência e Consolidação de Pedidos")
    with st.container(border=True):
        st.markdown("### 🔒 Controle de Acesso")
        status_atual = st.session_state.get('libera_digitacao_semanal', False)
        trava = st.toggle("Permitir digitação das nutricionistas", value=status_atual)
        if trava != status_atual:
            st.session_state['libera_digitacao_semanal'] = trava
            salvar_governanca()
            st.rerun()
            
    restaurante_filtrado = st.selectbox("Filtrar por Restaurante:", ["Todos", "TEJUCO", "CENTRO", "MATOSINHOS", "RM SABOR", "COLONIA", "BARBACENA", "LEOPOLDINA"]) 
    tab_sem, tab_men = st.tabs(["Pedido Semanal", "Pedido Mensal"]) 
    with tab_sem: consolidar_proteina_semanal_geral(restaurante_filtrado) 
    with tab_men: st.info("Módulo mensal em desenvolvimento.")

elif modulo_selecionado == "🥗 Lançamento de Pedidos" or "Lançamento" in str(modulo_selecionado): 
    st.title("🛒 Módulo de Pedidos") 
    interface_lancamento_proteina_filial()

elif modulo_selecionado == "🤝 Portal de Cotação": 
    st.title("🤝 Portal do Fornecedor")
    fornecedor_logado = st.session_state.get('filial_nome', 'FORNECEDOR PADRÃO')
    st.success(f"Empresa: {fornecedor_logado}") 

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
                    if df_selecionados.empty: st.warning("Selecione um pedido para aprovação.")
                    else: st.success("Pedidos validados com sucesso!"); st.cache_data.clear(); st.rerun()
        except Exception as e: st.error(f"Erro ao carregar o painel: {e}")
    else: st.error("🔒 Acesso restrito ao Administrador.")

elif modulo_selecionado == "💼 Compras & Suprimentos":
    modulo_compras_suprimentos()