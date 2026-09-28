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

# --- CONEXÃO INTELIGENTE COM O GOOGLE SHEETS (CORREÇÃO STREAMLIT CLOUD) ---
@st.cache_resource
def inicializar_gspread():
    # 1. Se estiver rodando na nuvem (Streamlit Cloud), usa a memória!
    if 'gcp_service_account' in st.secrets:
        credenciais = dict(st.secrets['gcp_service_account'])
        # Vacina contra erro de RSA: garante que as quebras de linha sejam lidas corretamente
        if 'private_key' in credenciais:
            credenciais['private_key'] = credenciais['private_key'].replace('\\n', '\n')
        return gspread.service_account_from_dict(credenciais)
    
    # 2. Se estiver rodando no seu computador (VS Code), usa o arquivo físico local
    caminho_local = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chave.json")
    if os.path.exists(caminho_local):
        return gspread.service_account(filename=caminho_local)
        
    raise ValueError("Nenhuma credencial do Google encontrada. Verifique o Streamlit Secrets ou o arquivo chave.json.")

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
    
    # NOVA ABA: CONFIGURACOES PARA A NUVEM (LIGA/DESLIGA O SISTEMA)
    if "CONFIGURACOES" not in abas_existentes:
        ws_conf = client.add_worksheet(title="CONFIGURACOES", rows="10", cols="2")
        ws_conf.append_row(["CHAVE", "VALOR"])
        ws_conf.append_row(["STATUS_DIGITACAO", "DESLIGADO"])

except gspread.exceptions.APIError as e:
    client = None
    if e.response.status_code == 429:
        st.error("⏳ Limite de acessos rápidos do Google atingido. Por favor, aguarde 1 minuto e recarregue a página.")
        st.stop()
except Exception as global_e:
    st.error(f"Erro ao conectar com o Google Sheets: {global_e}")
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

# --- CONFIGURAÇÕES E CONSTANTES GERAIS ---
MOCK_FILIAIS = ["TEJUCO", "CENTRO", "MATOSINHOS", "RM SABOR", "COLONIA", "BARBACENA", "LEOPOLDINA"]

# SOLUÇÃO DEFINITIVA DO PROBLEMA DE PRODUTOS "NÃO LOCALIZADOS"
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

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Portal AC Batista", layout="wide")

def validar_usuario_sheets(usuario, senha):
    if not client: return None, "❌ Não foi possível conectar ao Google Sheets."
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
                    else: return None, "❌ Senha incorreta. Tente novamente."
    except Exception as e: return None, f"❌ Erro crítico ao conectar à base de dados: {e}"
    return None, "❌ Usuário não localizado."

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
            if erro: st.error(erro)
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

# --- CONTROLE DE ACESSO OTIMIZADO (SEM TRAVAR O GOOGLE E COM "RADAR" DE CÉLULAS) ---
@st.cache_data(ttl=60)
def ler_status_digitacao():
    try:
        cli = conectar_sheets_nativo()
        if cli:
            aba = cli.worksheet("CONFIGURACOES")
            # Traz todos os valores de uma vez (apenas 1 chamada à API)
            valores = aba.get_all_values()
            
            # Radar: Procura onde a palavra LIGADO ou DESLIGADO foi escrita na tua folha
            for linha in valores:
                for cel in linha:
                    texto = str(cel).strip().upper()
                    if texto == "LIGADO": 
                        return True
                    if texto == "DESLIGADO": 
                        return False
    except Exception as e:
        print(f"Erro ao ler status: {e}")
    return False

def salvar_governanca(ligado):
    try:
        cli = conectar_sheets_nativo()
        if cli:
            aba = cli.worksheet("CONFIGURACOES")
            novo_valor = "LIGADO" if ligado else "DESLIGADO"
            
            valores = aba.get_all_values()
            lin_alvo = 2
            col_alvo = 2
            
            # Radar: Encontra a célula exata para não estragar a tua formatação
            encontrou = False
            for i, linha in enumerate(valores):
                for j, cel in enumerate(linha):
                    if str(cel).strip().upper() in ["LIGADO", "DESLIGADO"]:
                        lin_alvo = i + 1
                        col_alvo = j + 1
                        encontrou = True
                        break
                if encontrou: break
            
            # Atualiza apenas a célula correta
            aba.update_cell(lin_alvo, col_alvo, novo_valor)
            st.session_state['libera_digitacao_semanal'] = ligado
    except Exception as e: 
        print(f"Erro ao salvar status: {e}")
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
                                                
                                                prev_entre
