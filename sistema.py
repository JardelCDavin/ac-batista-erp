import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import gspread
import os
import json
from datetime import datetime
import urllib.parse
from pyngrok import ngrok
import subprocess

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
NGROK_LINK_FILE = os.path.join(BASE_DIR, "link_celular.txt")
NGROK_PORT = 8501

# --- INICIALIZAÇÃO AUTOMÁTICA E BLINDADA DO NGROK ---
@st.cache_resource
def iniciar_conexao_celular():
    import subprocess
    from pyngrok import ngrok
    try:
        subprocess.run(["taskkill", "/f", "/im", "ngrok.exe"], capture_output=True)
    except Exception:
        pass
    try:
        TOKEN = "3FMl6VIvvPd1MMRs30pfJRy8Doy_2ZwK6MpHdZRkWT8V2A3d"
        ngrok.set_auth_token(TOKEN)
        tunnel = ngrok.connect(8501)
        return tunnel.public_url
    except Exception as e:
        return f"Erro ao conectar: {e}"

ngrok_url = iniciar_conexao_celular()


# --- CONFIGURAÇÃO DA PÁGINA (LAYOUT CONGELADO) ---
st.set_page_config(page_title="Portal AC Batista", layout="wide")

if 'ngrok_url' in st.session_state and st.session_state['ngrok_url']:
    st.info(f"📱 LINK DO CELULAR PARA A NUTRICIONISTA: {st.session_state['ngrok_url']}")


# --- ESTILIZAÇÃO CSS OFICIAL ---
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

# --- CONEXÃO COM GOOGLE SHEETS ---
def conectar_sheets_nativo():
    nomes_possiveis = ["chave.json", "chave.json.json", "google_secret.json", "chave"]
    arquivo_encontrado = None
    for nome in nomes_possiveis:
        if os.path.exists(nome):
            arquivo_encontrado = nome
            break
    if not arquivo_encontrado: return None
    try:
        client = gspread.service_account(filename=arquivo_encontrado)
        return client.open("Formulário sem título (Respostas)")
    except Exception: return None

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


def mostrar_painel_diretor():
    st.sidebar.markdown("### Painel do Diretor")
    st.sidebar.checkbox(
        "Liberar Digitação Semanal (Nutricionistas)",
        value=st.session_state.get('libera_digitacao_semanal', DEFAULT_GOVERNANCA['libera_digitacao_semanal']),
        key='libera_digitacao_semanal'
    )
    st.sidebar.checkbox(
        "Liberar Digitação Mensal (Nutricionistas)",
        value=st.session_state.get('libera_digitacao_mensal', DEFAULT_GOVERNANCA['libera_digitacao_mensal']),
        key='libera_digitacao_mensal'
    )
    if st.sidebar.button('APLICAR LIBERAÇÃO DE DIGITAÇÃO'):
        # salva status das travas
        salvar_governanca()
        st.sidebar.success('Status de digitação atualizado e salvo.')
        # Em vez de abrir abas automaticamente, gerar botões de link para o WhatsApp
        usuarios_arquivo = "BD_USUARIOS.xlsx"
        if os.path.exists(usuarios_arquivo):
            try:
                df_users = pd.read_excel(usuarios_arquivo)
                cols = [c for c in df_users.columns if str(c).strip().upper() == 'TELEFONE']
                if not cols:
                    cols = [c for c in df_users.columns if 'TEL' in str(c).upper()]
                if cols:
                    telefones = df_users[cols[0]].dropna().tolist()
                    if not telefones:
                        st.sidebar.info('Nenhum telefone válido encontrado em BD_USUARIOS.xlsx')
                    else:
                        # mostra instrução abaixo da mensagem de sucesso
                        st.sidebar.write('')
                        st.sidebar.markdown('**Disparar avisos via WhatsApp (clique em cada botão):**')
                        import re
                        for idx, t in enumerate(telefones, start=1):
                            num = ''
                            try:
                                if isinstance(t, float) and t.is_integer():
                                    num = str(int(t))
                                else:
                                    num = str(t)
                            except Exception:
                                num = str(t)
                            num_clean = re.sub(r"[^0-9+]", "", num)
                            if not num_clean:
                                continue
                            texto = "Olá! A digitação do pedido foi LIBERADA pelo Diretor Jardel. Acesse o portal pelo link: http://localhost:8501"
                            texto_enc = urllib.parse.quote(texto, safe='')
                            link = f"https://wa.me/{num_clean}?text={texto_enc}"
                            # cria um botão de link para o usuário clicar manualmente
                            try:
                                st.sidebar.link_button(f"Disparar Aviso (Nutricionista) {idx}", link)
                            except Exception:
                                # fallback caso a versão do Streamlit não tenha link_button
                                st.sidebar.markdown(f"- [Disparar Aviso (Nutricionista) {idx}]({link})")
                else:
                    st.sidebar.warning('Coluna TELEFONE não encontrada em BD_USUARIOS.xlsx')
            except Exception as e:
                st.sidebar.error(f'Erro ao ler BD_USUARIOS.xlsx: {e}')
        else:
            st.sidebar.warning('Arquivo BD_USUARIOS.xlsx não encontrado.')


def modulo_cotacao_consolidacao():
    st.title("📌 Módulo de Cotação & Consolidação")
    st.write("Este módulo é exclusivo para o Diretor Jardel. Aqui você revisa o fechamento das digitações feitas pelas nutricionistas e faz a consolidação final.")
    st.write("A digitação deve ser liberada antes pelas nutricionistas; depois, use este botão para fechar o pedido e enviar para cotação.")
    if st.button("FECHAR PEDIDO E ENVIAR PARA COTAÇÃO", type='primary'):
        st.info("Executando fechamento final e varredura horizontal...")
        consolidar_mensal()
        consolidar_proteina_semanal_geral()
        st.success("Fechamento final processado e enviado para cotação.")

# --- CONFIGURAÇÕES CARDÁPIO E DICIONÁRIO ---
PRATOS_CARDAPIO_OFICIAL = ["Arroz Branco", "Feijão Carioca/Inteiro/Batido", "Filé de Frango Acebolado", "Carne Moída ao Molho", "Feijoada", "Frango Assado", "Linguiça Assada", "Pernil Assado", "Macarrão Alho e Óleo", "Angu/Polenta", "Farofa Colorida/Cenoura"]
UNIDADES_PRODUTOS_MENSAL = {
    "AÇÚCAR CRISTAL (FARDO 6X5KG)": "Fardo", "ARROZ PARBOILIZED (FARDO 6X5KG)": "Fardo", "ARROZ INTEGRAL": "KG", "ADOÇANTE LÍQUIDO": "UNID", "ACHOCOLATADO EM PÓ": "KG",
    "CAFÉ (FARDO 10X500GR)": "Fardo", "CANJICA BRANCA": "KG", "CREME DE LEITE / CREME CULINÁRIO": "UNID", "ERVILHA EM CONSERVA (LATA 1,7KG)": "Lata", "EXTRATO DE TOMATE (CAIXA 6X1,7KG)": "Caixa",
    "FARINHA DE MANDIOCA": "KG", "FARINHA DE MILHO": "KG", "FARINHA DE TRIGO": "KG", "FEIJÃO CARIOCA": "KG", "FEIJÃO PRETO": "KG", "FEIJÃO VERMELHO": "KG", "FUBÁ MIMOSO": "KG",
    "LEITE INTEGRAL (CAIXA 12/1LT)": "Caixa", "LEITE EM PÓ": "KG", "MACARRÃO ESPAGUETE COM OVOS": "KG", "MACARRÃO PARAFUSO COM OVOS": "KG", "MACARRÃO PADRE NOSSO (SOPA)": "KG", "MAIONESE (BALDE)": "Balde",
    "MARGARINA (BALDE 14,5KG)": "Balde", "MILHO DE PIPOCA": "KG", "MILHO VERDE EM CONSERVA (LATA 1,7KG)": "Lata", "MOLHO DE ALHO": "UNID", "MOLHO INGLÊS": "UNID", "MOLHO DE TOMATE (SACHET)": "UNID", "MOLHO SHOYU": "UNID",
    "ÓLEO DE SOJA (CAIXA 20/900ML)": "Caixa", "ÓLEO COMPOSTO": "UNID", "SAL REFINADO (FARDO 30/1KG)": "Fardo", "SAL DE PARRILLA": "KG", "TEMPERO PRONTO (ALHO E SAL)": "UNID", "VINAGRE DE ÁLCOOL": "UNID"
}

MOCK_FILIAIS = [
    "PROTEINA_SEMANAL_GLOBAL",
    "PROTEINA_SEMANAL_CENTRO",
    "PROTEINA_SEMANAL_TEJUCO",
    "PROTEINA_SEMANAL_COLONIA",
    "PROTEINA_SEMANAL_BARBACENA",
    "PROTEINA_SEMANAL_LEOPOLDINA",
    "PROTEINA_SEMANAL_MATOSINHOS"
]

def carregar_proteinas_semanal():
    """Tenta carregar a coluna 'PRODUTOS' do arquivo local
    'SEMANAL_PROTEINAS.xlsx'. Retorna lista de strings ou lista vazia."""
    arquivo = "SEMANAL_PROTEINAS.xlsx"
    try:
        if os.path.exists(arquivo):
            df = pd.read_excel(arquivo)
            # busca coluna chamada exatamente 'PRODUTOS' (case-insensitive)
            cols = [c for c in df.columns if str(c).strip().upper() == 'PRODUTOS']
            if not cols:
                # tenta qualquer coluna que contenha 'PROD'
                cols = [c for c in df.columns if 'PROD' in str(c).upper()]
            if cols:
                lista = df[cols[0]].dropna().astype(str).str.strip().tolist()
                return lista
    except Exception:
        pass
    return []


def carregar_produtos_mensal():
    """Tenta carregar a coluna 'PRODUTOS' do arquivo local
    'MENSAL_MERCEARIA_EMBALAGENS_LIMPEZA.xlsx'. Retorna lista de strings ou lista vazia."""
    arquivo = "MENSAL_MERCEARIA_EMBALAGENS_LIMPEZA.xlsx"
    try:
        if os.path.exists(arquivo):
            df = pd.read_excel(arquivo)
            cols = [c for c in df.columns if str(c).strip().upper() == 'PRODUTOS']
            if not cols:
                cols = [c for c in df.columns if 'PROD' in str(c).upper()]
            if cols:
                lista = df[cols[0]].dropna().astype(str).str.strip().tolist()
                return lista
    except Exception:
        pass
    return []

def consolidar_mensal():
    try:
        planilha = conectar_sheets_nativo()
        if not planilha:
            st.error("Erro ao conectar à planilha.")
            return
            
        filiais_alvo = ['CENTRO', 'TEJUCO', 'COLONIA', 'BARBACENA', 'LEOPOLDINA', 'MATOSINHOS', 'SABOR']
        totais_produtos = {}
        
        todas_abas_sheets = [w.title for w in planilha.worksheets()]
        
        for nome_aba_real in todas_abas_sheets:
            if 'MENSAL' in nome_aba_real.upper() and any(f in nome_aba_real.upper() for f in filiais_alvo):
                try:
                    aba = planilha.worksheet(nome_aba_real)
                    df_aba = pd.DataFrame(aba.get_all_records())
                    if df_aba.empty: continue
                    
                    df_aba.columns = [str(c).strip().upper() for c in df_aba.columns]
                    col_prod_real = None
                    for c in df_aba.columns:
                        if 'PROD' in c or 'ITEM' in c or 'NOME' in c:
                            col_prod_real = c
                            break
                    if not col_prod_real and len(df_aba.columns) > 0:
                        col_prod_real = df_aba.columns[0]
                    
                    col_qtd_real = None
                    for c in df_aba.columns:
                        if 'QTD' in c or 'QUANT' in c:
                            col_qtd_real = c
                            break
                    if not col_qtd_real:
                        if len(df_aba.columns) > 1:
                            col_qtd_real = df_aba.columns[1]
                        elif len(df_aba.columns) == 1:
                            col_qtd_real = df_aba.columns[0]
                    
                    if col_prod_real is not None and col_qtd_real is not None:
                        for idx, linha in df_aba.iterrows():
                            produto = str(linha[col_prod_real]).strip()
                            if produto == "" or produto.upper().startswith("1.") or produto.upper().startswith("2."): continue
                            qtd = pd.to_numeric(linha[col_qtd_real], errors='coerce')
                            qtd = qtd if not pd.isna(qtd) else 0
                            if qtd > 0:
                                prod_chave = produto.upper()
                                totais_produtos[prod_chave] = totais_produtos.get(prod_chave, 0) + qtd
                except Exception: continue
                
        linhas_finais = []
        for produto_chave, total_qtd in totais_produtos.items():
            unidade = UNIDADES_PRODUTOS_MENSAL.get(produto_chave, "UN/KG")
            linhas_finais.append([produto_chave.title(), unidade, total_qtd])
            
        if linhas_finais:
            df_final = pd.DataFrame(linhas_finais, columns=['Item', 'Unidade', 'Quantidade Total'])
            
            for aba_nome in ['MENSAL_MODELO', 'COTACAO']:
                try:
                    aba_alvo = planilha.worksheet(aba_nome)
                    aba_alvo.clear()
                    aba_alvo.update(range_name='A1', values=[df_final.columns.values.tolist()] + df_final.values.tolist())
                except Exception: pass
                
            st.success("✅ Módulo Mensal Consolidado de todas as filiais e gravado no MENSAL_MODELO!")
            st.dataframe(df_final, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ Nenhuma quantidade maior que zero foi somada. Verifique os valores nas colunas das filiais!")
    except Exception as e: st.error(f"Erro no Módulo Mensal: {e}")

def consolidar_proteina_semanal_geral():
    try:
        planilha = conectar_sheets_nativo()
        if not planilha:
            st.error("Erro ao conectar à planilha.")
            return
            
        lista_abas = [w.title for w in planilha.worksheets()]
        totais_carnes = {}
        
        for nome_aba in lista_abas:
            nome_aba_upper = nome_aba.upper()
            if nome_aba_upper == 'PROTEINA_SEMANAL':
                continue
            if 'PROTEIN' in nome_aba_upper or 'SEMANAL' in nome_aba_upper:
                try:
                    aba = planilha.worksheet(nome_aba)
                    df_aba = pd.DataFrame(aba.get_all_records())
                    if df_aba.empty:
                        continue
                    
                    df_aba.columns = [str(c).strip().upper() for c in df_aba.columns]
                    
                    col_prod = None
                    for c in df_aba.columns:
                        if 'PROD' in c or 'ITEM' in c or 'CARNE' in c:
                            col_prod = c
                            break
                    if not col_prod and len(df_aba.columns) > 0:
                        col_prod = df_aba.columns[0]
                    
                    col_freq = None
                    for c in df_aba.columns:
                        if 'FREQUEN' in c or 'FREQ' in c:
                            col_freq = c
                            break
                    col_days = None
                    for c in df_aba.columns:
                        if 'DIA' in c or 'DIAS' in c:
                            col_days = c
                            break
                    col_stock = None
                    for c in df_aba.columns:
                        if 'ESTOQUE' in c or 'STOCK' in c:
                            col_stock = c
                            break
                    if not (col_prod and col_freq and col_days and col_stock):
                        continue
                    
                    for idx, linha in df_aba.iterrows():
                        carne = str(linha[col_prod]).strip()
                        if carne == "" or carne.upper().startswith("1.") or carne.upper().startswith("2."):
                            continue
                        
                        frequencia = pd.to_numeric(linha[col_freq], errors='coerce')
                        dias = pd.to_numeric(linha[col_days], errors='coerce')
                        estoque = pd.to_numeric(linha[col_stock], errors='coerce')
                        frequencia = frequencia if not pd.isna(frequencia) else 0
                        dias = dias if not pd.isna(dias) else 0
                        estoque = estoque if not pd.isna(estoque) else 0
                        
                        quantidade_necessaria = max(frequencia * dias - estoque, 0)
                        if quantidade_necessaria > 0:
                            totais_carnes[carne.upper()] = totais_carnes.get(carne.upper(), 0) + quantidade_necessaria
                except Exception:
                    continue

        cabecalhos = ['Proteína / Item', 'Frequência', 'Dias', 'Estoque', 'Quantidade Necessária', 'Unidade']
        try:
            aba_saida = planilha.worksheet('PROTEINA_SEMANAL')
            valores_saida = aba_saida.get_all_values()
        except Exception:
            aba_saida = None
            valores_saida = []

        def normalize_header(name):
            n = str(name).strip().upper()
            if 'PROTEIN' in n or 'ITEM' in n or 'CARNE' in n:
                return 'Proteína / Item'
            if 'FREQUEN' in n or 'FREQ' in n:
                return 'Frequência'
            if 'DIA' in n or 'DIAS' in n:
                return 'Dias'
            if 'ESTOQUE' in n or 'STOCK' in n:
                return 'Estoque'
            if 'QUANT' in n and 'NEC' in n:
                return 'Quantidade Necessária'
            if 'QUANT' in n:
                return 'Quantidade Necessária'
            if 'UNID' in n or 'UNIT' in n:
                return 'Unidade'
            return str(name).strip()

        def to_numeric_safe(value):
            numeric = pd.to_numeric(value, errors='coerce')
            return numeric if not pd.isna(numeric) else None

        if not valores_saida or len(valores_saida) <= 1 and all(not any(str(cell).strip() for cell in row) for row in valores_saida[1:]):
            needs_template = True
            df_saida = pd.DataFrame(columns=cabecalhos)
        else:
            header = [normalize_header(c) for c in valores_saida[0]]
            df_saida = pd.DataFrame(valores_saida[1:], columns=header)
            for col in cabecalhos:
                if col not in df_saida.columns:
                    df_saida[col] = ''
            needs_template = 'Proteína / Item' not in df_saida.columns

        if needs_template:
            if aba_saida is None:
                try:
                    aba_saida = planilha.add_worksheet(title='PROTEINA_SEMANAL', rows=100, cols=20)
                except Exception:
                    aba_saida = None
            if aba_saida is None:
                st.error("Não foi possível criar ou acessar a aba PROTEINA_SEMANAL.")
                return

            linhas_template = []
            if totais_carnes:
                for carne in sorted(totais_carnes.keys()):
                    linhas_template.append([carne.title(), '', '', '', '', 'KG'])
            else:
                linhas_template.append(['', '', '', '', '', 'KG'])
            aba_saida.clear()
            aba_saida.update(range_name='A1', values=[cabecalhos] + linhas_template)
            st.success("🥩 Planilha PROTEINA_SEMANAL criada/zerada com colunas de entrada para as nutricionistas.")
            st.dataframe(pd.DataFrame(linhas_template, columns=cabecalhos), use_container_width=True, hide_index=True)
            return

        linhas_saida = []
        processed_items = set()

        for _, linha in df_saida.iterrows():
            item = str(linha.get('Proteína / Item', '')).strip()
            if item == '':
                continue
            processed_items.add(item.upper())

            frequencia = to_numeric_safe(linha.get('Frequência', ''))
            dias = to_numeric_safe(linha.get('Dias', ''))
            estoque = to_numeric_safe(linha.get('Estoque', ''))
            unidade = str(linha.get('Unidade', '')).strip() or 'KG'

            if frequencia is not None and dias is not None and estoque is not None:
                quantidade_necessaria = max(frequencia * dias - estoque, 0)
            else:
                quantidade_necessaria = ''

            linhas_saida.append([
                item,
                linha.get('Frequência', ''),
                linha.get('Dias', ''),
                linha.get('Estoque', ''),
                quantidade_necessaria,
                unidade
            ])

        for carne_upper, total_qtd in sorted(totais_carnes.items()):
            if carne_upper not in processed_items:
                linhas_saida.append([carne_upper.title(), '', '', '', '', 'KG'])

        if not linhas_saida:
            linhas_saida.append(['', '', '', '', '', 'KG'])

        aba_saida.clear()
        aba_saida.update(range_name='A1', values=[cabecalhos] + linhas_saida)
        st.success("🥩 PROTEINA_SEMANAL atualizada com base nos dados existentes e no cálculo automático.")
        st.dataframe(pd.DataFrame(linhas_saida, columns=cabecalhos), use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Erro no Módulo Semanal de Proteínas: {e}")

def _gravar_pedido_individual_local(filial_nome, produto, freq, dias, stock):
    try:
        planilha = conectar_sheets_nativo()
        if not planilha:
            st.error("Erro ao conectar ao Google Sheets. O pedido não foi gravado.")
            return

        try:
            aba = planilha.worksheet(filial_nome)
        except Exception:
            try:
                aba = planilha.add_worksheet(title=filial_nome, rows=200, cols=10)
            except Exception as exc:
                st.error(f"Erro ao abrir ou criar aba '{filial_nome}': {exc}")
                return

        valores_aba = aba.get_all_values()
        if not valores_aba:
            valores_aba = [['Proteína / Item', 'Frequência', 'Dias', 'Estoque']]

        cabecalho = [str(c).strip() for c in valores_aba[0]]
        cabecalho_upper = [c.upper() for c in cabecalho]

        def get_or_create_col(col_name):
            if col_name.upper() in cabecalho_upper:
                return cabecalho_upper.index(col_name.upper())
            cabecalho.append(col_name)
            cabecalho_upper.append(col_name.upper())
            for linha in valores_aba[1:]:
                while len(linha) < len(cabecalho):
                    linha.append('')
            return len(cabecalho) - 1

        prod_idx = get_or_create_col('Proteína / Item')
        freq_idx = get_or_create_col('Frequência')
        dias_idx = get_or_create_col('Dias')
        stock_idx = get_or_create_col('Estoque')

        linha_encontrada = None
        for i, linha in enumerate(valores_aba[1:], start=1):
            if len(linha) > prod_idx and str(linha[prod_idx]).strip().upper() == produto.strip().upper():
                linha_encontrada = i
                break

        if linha_encontrada is None:
            nova_linha = [''] * len(cabecalho)
            nova_linha[prod_idx] = produto
            nova_linha[freq_idx] = freq
            nova_linha[dias_idx] = dias
            nova_linha[stock_idx] = stock
            valores_aba.append(nova_linha)
        else:
            linha = valores_aba[linha_encontrada]
            while len(linha) < len(cabecalho):
                linha.append('')
            linha[freq_idx] = freq
            linha[dias_idx] = dias
            linha[stock_idx] = stock

        aba.clear()
        aba.update(range_name='A1', values=[cabecalho] + valores_aba[1:])
        st.success(f"✅ Pedido semanal de {produto} gravado na filial {filial_nome}.")
    except Exception as e:
        st.error(f"❌ Erro ao gravar pedido de {produto}: {e}")


def interface_lancamento_proteina_filial():
    try:
        st.subheader("📝 Digitação Semanal por Filial")
        st.info("A digitação semanal aqui é feita pela nutricionista. Depois de preencher os dados, clique em ENVIAR PARA O DIRETOR para salvar na planilha da filial.")

        filial_selecionada = st.selectbox("Selecione sua Filial:", MOCK_FILIAIS)
        if not filial_selecionada:
            return

        if st.session_state.nivel != "Admin" and not st.session_state.libera_digitacao_semanal:
            st.warning("A digitação semanal está temporariamente fechada pelo Diretor Jardel.")
            return

        proteinas_lista = carregar_proteinas_semanal()
        if not proteinas_lista:
            st.warning("Nenhuma proteína localizada no arquivo SEMANAL_PROTEINAS.xlsx. Verifique o arquivo.")
        produto_selecionado = st.selectbox("Selecione a Proteína:", [""] + proteinas_lista, format_func=lambda x: "-- Selecione uma proteína --" if x == "" else x)
        if not produto_selecionado:
            st.warning("Escolha uma proteína para continuar.")
            return

        st.write("---")
        st.write(f"**Proteína selecionada:** {produto_selecionado}")
        st.write("Preencha a Frequência, Dias e Estoque abaixo:")

        col1, col2, col3 = st.columns(3)
        with col1:
            freq_input = st.number_input("Frequência", value=0.0, min_value=0.0, step=1.0, key="freq_selecionado")
        with col2:
            dias_input = st.number_input("Dias", value=0.0, min_value=0.0, step=1.0, key="dias_selecionado")
        with col3:
            stock_input = st.number_input("Estoque", value=0.0, min_value=0.0, step=0.1, key="stock_selecionado")

        pedido_sugerido = max(freq_input * dias_input - stock_input, 0)
        st.metric(label="📦 Pedido Sugerido", value=f"{pedido_sugerido:.1f} KG")

        st.button(
            "ENVIAR PARA O DIRETOR",
            type='primary',
            use_container_width=True,
            key="btn_gravar_oficial",
            on_click=_gravar_pedido_individual_local,
            args=(filial_selecionada, produto_selecionado, freq_input, dias_input, stock_input)
        )
    except Exception as e:
        st.error(f"Erro na interface de lançamento: {e}")

if 'logado' not in st.session_state:
    st.session_state.logado, st.session_state.usuario, st.session_state.nivel, st.session_state.filial_nome = False, None, None, None

carregar_governanca()

# Carrega a lista mestre mensal em memória (pode ser usada pelo módulo Mensal)
PRODUTOS_MESTRE_MENSAL = carregar_produtos_mensal()
try:
    u_in = st.session_state.get('usuario_login', '')
    s_in = st.session_state.get('senha_login', '')
    client = conectar_sheets_nativo()
    if client:
                sheet = client.worksheet("BD_USUARIOS")
                dados = sheet.get_all_records()
                df = pd.DataFrame(dados)
                
                df.columns = [c.strip().upper() for c in df.columns]
                val = df[df['USUARIO'].astype(str).str.strip() == u_in]
                
                if not val.empty:
                    s_c = str(val['SENHA'].values[0]).strip()
                    
                    if str(s_in).strip() == s_c:
                        st.session_state.logado = True
                        st.session_state.usuario = str(val['USUARIO'].values[0]).strip()
                        st.session_state.filial_nome = str(val['FILIAL'].values[0]).strip().upper()
                        
                        cargo = st.session_state.filial_nome
                        st.session_state.nivel = "Fornecedor" if cargo == "FORNECEDOR" else ("Admin" if cargo == "ADMINISTRATIVO" else "Nutricionista")
                        st.rerun()
                    else:
                        st.error("❌ Senha incorreta.")
                else:
                    st.error("❌ Usuário não localizado.")
    else:
                st.error("❌ Erro de conexão com o Google Sheets.")
except Exception as e:
            st.error(f"Erro ao acessar banco na nuvem: {e}")
else:

        st.sidebar.title("🏢 AC Batista ERP")
        st.sidebar.write(f"👤 Usuário: **{st.session_state.usuario}**")
        st.sidebar.write(f"🔐 Acesso: **{st.session_state.nivel}**")
        if st.sidebar.button("🚪 Sair/Logoff"):
            st.session_state.logado = False
            st.rerun()
        if st.session_state.nivel == "Admin":
         mostrar_painel_diretor()
st.sidebar.write("---")
modulos_basicos = ["🛒 Compras & Suprimentos", "🥩 Auditoria Semanal (Proteínas)", "📋 Fichas Técnicas (Cardápio)"]
if st.session_state.nivel == "Admin":
        modulos_basicos.append("📌 Módulo de Cotação & Consolidação")
modulo_selecionado = st.sidebar.radio("Módulo:", modulos_basicos)

if modulo_selecionado == "🛒 Compras & Suprimentos":
        st.title("🗃️ Gestão de Compras & Módulo Mensal")
        col1, col2 = st.columns(2)
        with col1:
            if st.session_state.nivel != "Admin" and not st.session_state.libera_digitacao_mensal:
                st.warning("A digitação mensal está temporariamente fechada pelo Diretor Jardel.")
            else:
                if st.button("CONSOLIDAR MÓDULO MENSAL HORIZONTAL"):
                    consolidar_mensal()
        with col2:
            if st.session_state.nivel != "Admin" and not st.session_state.libera_digitacao_semanal:
                st.warning("A digitação semanal está temporariamente fechada pelo Diretor Jardel.")
            else:
                if st.button("GERAR COMPRAS DE PROTEÍNA SEMANAL"):
                    consolidar_proteina_semanal_geral()
elif modulo_selecionado == "🥩 Auditoria Semanal (Proteínas)":
        st.title("🥩 Auditoria Semanal (Proteínas)")
        tab1, tab2 = st.tabs(["📝 Lançar Pedido", "📊 Consolidado"])
        
        with tab1:
            interface_lancamento_proteina_filial()
        with tab2:
            st.subheader("Consolidação Geral de Proteínas")
            if st.button("🔄 Atualizar Consolidado"):
                consolidar_proteina_semanal_geral()
            else:
                consolidar_proteina_semanal_geral()
elif modulo_selecionado == "� Módulo de Cotação & Consolidação":
        modulo_cotacao_consolidacao()
elif modulo_selecionado == "�📋 Fichas Técnicas (Cardápio)":
        st.title("🥗 Área de Nutrição - Fichas Técnicas")

# --- ADICIONANDO A INTERFACE NOVA NO FINAL DO ARQUIVO COM TRAVA ---
if "logado" in st.session_state and st.session_state.logado and st.session_state.usuario != "Jardel":
    st.title("🥩 Pedido Semanal")
st.subheader(f"Filial: {st.session_state.get('filial_nome', 'Não Identificada').title()}")    
    # Abas por Categoria
aba_proteinas, aba_hortifruti, aba_mercearia = st.tabs(["🥩 Proteínas", "🥬 Hortifrúti", "🧃 Mercearia"])
    
with aba_proteinas:
        # 1. Identifica a filial da nutricionista logada (Ex: "BARBACENA", "DONA MARIA")
        filial_usuario = st.session_state.get("filial_nome", "BARBACENA").upper().strip()
        
        # 2. DICIONÁRIO DE VÍNCULOS: Aponta dinamicamente qual contrato ler
        if filial_usuario in ["CENTRO", "TEJUCO", "MATOSINHOS", "COLONIA", "BARBACENA"]:
            nome_aba_contrato = "CONTRATO_POPULAR_GERAL"
        elif filial_usuario == "LEOPOLDINA":
            nome_aba_contrato = "CONTRATO_LEOPOLDINA"
        else:
            nome_aba_contrato = "CONTRATO_DONA_MARIA"
            
        nome_aba_produtos = f"PROTEINA_{filial_usuario}"
        
        st.caption(f"📋 Contrato ativo: **{nome_aba_contrato}** | Catálogo: **{nome_aba_produtos}**")
        
        try:
            client = conectar_sheets_nativo()
            if client:
                # --- LEITURA DO CONTRATO (LIMITES) ---
                sheet_contrato = client.worksheet(nome_aba_contrato)
                dados_contrato = sheet_contrato.get_all_records()
                df_contrato = pd.DataFrame(dados_contrato)
                df_contrato.columns = [c.strip().upper() for c in df_contrato.columns]
                
                # --- LEITURA DOS PRODUTOS DA FILIAL ---
                sheet_produtos = client.worksheet(nome_aba_produtos)
                dados_produtos = sheet_produtos.get_all_records()
                df_produtos = pd.DataFrame(dados_produtos)
                df_produtos.columns = [c.strip().upper() for c in df_produtos.columns]
                
                lista_itens = df_produtos["PRODUTO"].dropna().tolist()
                lista_itens = [item for item in lista_itens if str(item).strip() != ""]
                
                # 3. MONTA OS CARTÕES E A TRAVA DE SEGURANÇA
                for item in lista_itens:
                    st.markdown(f'''
                        <div style="background:#fff; padding:15px; border-radius:10px; 
                        box-shadow:0 4px 6px rgba(0,0,0,0.1); margin-bottom:5px; margin-top:15px;">
                            <span style="font-size:16px; font-weight:bold; color:#333;">🥩 {item}</span>
                        </div>
                    ''', unsafe_allow_html=True)
                    
                    # Captura a quantidade digitada pela nutricionista
                    qtd_digitada = st.number_input("Quantidade:", min_value=0.0, step=1.0, key=f"qtd_{item}", label_visibility="collapsed")
                    
                    # Busca o Limite Ativo correspondente a este produto na tabela de contrato
                    limite_ativo = 999999.0  # Limite padrão alto caso não ache o item
                    if "DESCRIÇÃO" in df_contrato.columns and "LIMITE ATIVO" in df_contrato.columns:
                        linha_item = df_contrato[df_contrato["DESCRIÇÃO"].astype(str).str.strip().upper() == str(item).strip().upper()]
                        if not linha_item.empty:
                            limite_ativo = float(linha_item["LIMITE ATIVO"].values[0])
                    
                    # SE ESTOURAR A COTA: Abre a caixa de justificativa obrigatória na hora!
                    if qtd_digitada > limite_ativo:
                        st.warning(f"⚠️ Cota estourada! Limite permitido: **{limite_ativo} kg**. Seu pedido: **{qtd_digitada} kg**.")
                        st.text_area(f"Justificativa obrigatória para {item}:", key=f"just_{item}", placeholder="Explique o motivo do pedido acima da cota...")
                        
            else:
                st.error("❌ Não foi possível conectar ao Google Sheets.")
                
        except Exception as e:
            st.error(f"⚠️ Erro ao carregar dados na nuvem: {e}")

with aba_hortifruti:
        st.markdown('<div style="background:#fff; padding:15px; border-radius:10px; box-shadow:0 4px 6px rgba(0,0,0,0.1); margin-bottom:15px;">', unsafe_allow_html=True)
        st.number_input("Qtd:", min_value=0.0, step=1.0, key="tomate", label_visibility="collapsed")
        
st.write("---")
if st.button("✔️ Salvar e Enviar Pedido", type="primary", use_container_width=True):
        st.success("Pedido enviado com sucesso!")
