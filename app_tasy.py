import pandas as pd
import numpy as np
from datetime import datetime
import warnings
import streamlit as st
from pathlib import Path
import io
import json
warnings.filterwarnings('ignore')

# ==================== CONFIGURAÇÃO DA PÁGINA ====================
st.set_page_config(
    page_title="Gerador de Planilha TASY",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #0066cc;
        color: white;
        height: 3em;
        border-radius: 10px;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #0052a3;
    }
    .success-box {
        padding: 1rem;
        border-radius: 10px;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .warning-box {
        padding: 1rem;
        border-radius: 10px;
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        color: #856404;
    }
    .error-box {
        padding: 1rem;
        border-radius: 10px;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
    }
    </style>
""", unsafe_allow_html=True)

# ==================== CARREGAR CONFIGURAÇÕES ====================
@st.cache_data
def carregar_configuracoes():
    """Carrega configurações do arquivo JSON"""
    try:
        with open('config_exames.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        st.error("❌ Arquivo config_exames.json não encontrado!")
        st.info("📄 Crie o arquivo config_exames.json no mesmo diretório do app_tasy.py")
        st.stop()
    except json.JSONDecodeError as e:
        st.error(f"❌ Erro ao ler config_exames.json: {str(e)}")
        st.info("Verifique se o JSON está com formato válido")
        st.stop()

# Carregar configurações
CONFIG = carregar_configuracoes()
ESTABELECIMENTOS = CONFIG['estabelecimentos']

# Construir mapas de exames
# Estrutura: codigo_tasy -> {'nome': nome, 'colunas_lab': [...], 'categoria': categoria}
MAPA_EXAMES_COMPLETO = {}
MAPA_EXAMES_POR_CODIGO = {}

for exame in CONFIG['exames']:
    codigo_completo = f"NR_EXAME_{exame['codigo_tasy']}"
    
    # Mapa completo para exibição na UI
    for coluna in exame['colunas_lab']:
        MAPA_EXAMES_COMPLETO[coluna] = {
            'codigo': codigo_completo,
            'nome': exame['nome']
        }
    
    # Mapa por código com todas as variações de coluna
    MAPA_EXAMES_POR_CODIGO[codigo_completo] = {
        'nome': exame['nome'],
        'colunas_lab': exame['colunas_lab'],
        'categoria': exame['categoria']
    }

# Ordem das colunas
ORDEM_COLUNAS_TASY = CONFIG.get('ordem_colunas_tasy', [
    'NM_PACIENTE', 'NR_ATENDIMENTO', 'DT_RESULTADO', 'DS_PROTOCOLO', 'CD_ESTABELECIMENTO'
])

# ==================== TÍTULO ====================
st.title("🏥 Gerador de Planilha para TASY")
st.markdown("---")

# ==================== FUNÇÕES ====================
def mapear_exames_para_tasy(df, categoria=None):
    """
    Mapeia colunas do DataFrame para códigos TASY.
    Testa todas as variações de nomenclatura até encontrar uma que existe.
    
    Args:
        df: DataFrame com os dados do laboratório
        categoria: 'basico' ou 'resultados2' (None = todos)
    
    Returns:
        dict: {codigo_tasy: nome_coluna_encontrada}
    """
    mapeamento = {}
    
    for codigo_tasy, info in MAPA_EXAMES_POR_CODIGO.items():
        # Filtrar por categoria se especificado
        if categoria and info['categoria'] != categoria:
            continue
        
        # Testar cada variação de coluna até encontrar uma que existe
        for coluna_variacao in info['colunas_lab']:
            if coluna_variacao in df.columns:
                mapeamento[codigo_tasy] = coluna_variacao
                break  # Encontrou, não precisa testar as outras variações
    
    return mapeamento

def normalizar_nome(nome):
    if pd.isna(nome):
        return ''
    return str(nome).strip().lower().replace('  ', ' ')

def converter_valor_numerico(valor):
    if pd.isna(valor) or valor == '' or valor == ' ':
        return None
    try:
        if isinstance(valor, str):
            valor = valor.strip().replace(',', '.')
        return float(valor)
    except:
        return None

def formatar_data(data):
    if pd.isna(data):
        return None
    try:
        if isinstance(data, str):
            if '/' in data and ':' in data:
                return data
            data = pd.to_datetime(data)
        return data.strftime('%d/%m/%Y %H:%M:%S')
    except:
        return str(data)

def detectar_colunas_nome(df):
    possiveis = ['nome', 'paciente', 'nm_paciente', 'Paciente', 'Nome', 'NM_PACIENTE']
    for col in df.columns:
        if col in possiveis or col.lower() in [p.lower() for p in possiveis]:
            return col
    return df.columns[0]

def detectar_colunas_atendimento(df):
    possiveis = ['atendimento', 'nr_atendimento', 'Atendimento', 'NR_ATENDIMENTO']
    for col in df.columns:
        if col in possiveis or col.lower() in [p.lower() for p in possiveis]:
            return col
    return None

def ler_excel(arquivo):
    """Lê arquivo Excel enviado"""
    try:
        return pd.read_excel(arquivo, engine='openpyxl')
    except:
        try:
            return pd.read_excel(arquivo, engine='xlrd')
        except:
            return pd.read_excel(arquivo)

# ==================== SIDEBAR ====================
with st.sidebar:
    st.header("⚙️ Configurações")
    
    # Exibir versão da config
    st.info(f"📋 Config v{CONFIG.get('versao', 'N/A')}\n\n🗓️ Atualizada em: {CONFIG.get('ultima_atualizacao', 'N/A')}")
    
    protocolo = st.selectbox(
        "Protocolo",
        ["MENSAL", "TRIMESTRAL", "SEMESTRAL", "ANUAL"],
        index=0
    )
    
    # Seletor de Estabelecimento
    st.markdown("---")
    st.subheader("🏢 Estabelecimento")
    
    estabelecimentos_opcoes = list(ESTABELECIMENTOS.keys())
    estabelecimento_selecionado = st.selectbox(
        "Selecione o estabelecimento:",
        estabelecimentos_opcoes,
        index=0,
        help="Escolha o estabelecimento para todos os registros"
    )
    
    cd_estabelecimento_fixo = ESTABELECIMENTOS[estabelecimento_selecionado]
    
    # Mostrar código selecionado
    st.info(f"**Código:** {cd_estabelecimento_fixo}")
    
    st.markdown("---")
    
    st.header("📁 Arquivos")
    
    arquivo_basicos = st.file_uploader(
        "1️⃣ Exames Básicos (Laboratório)",
        type=['xlsx', 'xls'],
        help="Planilha com os exames básicos do laboratório"
    )
    
    arquivo_resultados2 = st.file_uploader(
        "2️⃣ Resultados 2 (Opcional)",
        type=['xlsx', 'xls'],
        help="Planilha com resultados complementares"
    )
    
    arquivo_pacientes = st.file_uploader(
        "3️⃣ Pacientes do TASY",
        type=['xlsx', 'xls'],
        help="Planilha com Nome + Atendimento dos pacientes"
    )
    
    st.markdown("---")
    
    processar = st.button("🚀 Processar", type="primary", use_container_width=True)

# ==================== ÁREA PRINCIPAL ====================
if not processar:
    # Tela inicial
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.image("https://via.placeholder.com/400x200/0066cc/ffffff?text=TASY+Import", use_container_width=True)
        
        st.markdown("""
        ### 📋 Como usar:
        
        1. **Selecione o protocolo** no menu lateral
        2. **Selecione o estabelecimento** no menu lateral
        3. **Faça upload dos arquivos**:
           - Exames Básicos (obrigatório)
           - Resultados 2 (opcional)
           - Pacientes do TASY (obrigatório)
        4. **Clique em Processar**
        5. **Baixe o arquivo** gerado
        
        ---
        
        ### ✅ Formatos aceitos:
        - `.xlsx` (Excel 2007+)
        - `.xls` (Excel 97-2003)
        
        ### 🔍 O que o sistema faz:
        - Cruza dados por nome do paciente
        - Mapeia exames para códigos TASY
        - Aplica o estabelecimento selecionado para todos os registros
        - Gera relatório de inconsistências
        """)
    
    # TABELA DE REFERÊNCIA DE EXAMES
    st.markdown("---")
    st.header("📊 Tabela de Referência - Mapeamento de Exames")
    
    # Criar DataFrame para exibição
    tabela_referencia = []
    codigos_unicos = {}
    
    for key, value in MAPA_EXAMES_COMPLETO.items():
        codigo = value['codigo']
        nome = value['nome']
        if codigo not in codigos_unicos:
            codigos_unicos[codigo] = {
                'codigo': codigo.replace('NR_EXAME_', ''),
                'nome': nome,
                'colunas_lab': [key]
            }
        else:
            if key not in codigos_unicos[codigo]['colunas_lab']:
                codigos_unicos[codigo]['colunas_lab'].append(key)
    
    for codigo, info in sorted(codigos_unicos.items()):
        tabela_referencia.append({
            'Código TASY': info['codigo'],
            'Nome do Exame': info['nome'],
            'Colunas do Lab': ', '.join(info['colunas_lab'])
        })
    
    df_referencia = pd.DataFrame(tabela_referencia)
    
    # Adicionar filtro de busca
    col1, col2 = st.columns([3, 1])
    with col1:
        busca = st.text_input("🔍 Buscar exame:", placeholder="Digite o nome ou código...")
    
    # Filtrar tabela
    if busca:
        mask = (
            df_referencia['Nome do Exame'].str.contains(busca, case=False, na=False) |
            df_referencia['Código TASY'].str.contains(busca, case=False, na=False) |
            df_referencia['Colunas do Lab'].str.contains(busca, case=False, na=False)
        )
        df_filtrado = df_referencia[mask]
    else:
        df_filtrado = df_referencia
    
    # Exibir tabela
    st.dataframe(
        df_filtrado,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Código TASY": st.column_config.TextColumn("Código TASY", width="small"),
            "Nome do Exame": st.column_config.TextColumn("Nome do Exame", width="large"),
            "Colunas do Lab": st.column_config.TextColumn("Colunas do Laboratório", width="medium")
        }
    )
    
    st.info(f"📌 Total de exames mapeados: **{len(df_filtrado)}** de **{len(df_referencia)}**")
    
    # Tabela de Estabelecimentos
    st.markdown("---")
    st.header("🏢 Códigos de Estabelecimento")
    
    df_estabelecimentos = pd.DataFrame([
        {"Código": v, "Estabelecimento": k}
        for k, v in sorted(ESTABELECIMENTOS.items(), key=lambda x: x[1])
    ])
    
    st.dataframe(df_estabelecimentos, use_container_width=True, hide_index=True)

else:
    if not arquivo_basicos or not arquivo_pacientes:
        st.error("❌ Por favor, envie pelo menos os arquivos de Exames Básicos e Pacientes!")
    else:
        try:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("📂 Carregando arquivos...")
            progress_bar.progress(10)
            
            pacientes = ler_excel(arquivo_pacientes)
            col_nome_pac = detectar_colunas_nome(pacientes)
            col_atend_pac = detectar_colunas_atendimento(pacientes)
            
            if not col_atend_pac:
                st.error(f"❌ Coluna de atendimento não encontrada!\n\nColunas disponíveis: {', '.join(pacientes.columns)}")
                st.stop()
            
            basicos = ler_excel(arquivo_basicos)
            col_nome_lab = detectar_colunas_nome(basicos)
            
            if arquivo_resultados2:
                resultados2 = ler_excel(arquivo_resultados2)
            else:
                resultados2 = None
            
            progress_bar.progress(30)
            
            status_text.text("🔍 Indexando pacientes...")
            pacientes['nome_normalizado'] = pacientes[col_nome_pac].apply(normalizar_nome)
            indice_pacientes = pacientes.set_index('nome_normalizado')[col_atend_pac].to_dict()
            progress_bar.progress(40)
            
            status_text.text("⚙️ Processando exames básicos...")
            basicos['nome_normalizado'] = basicos[col_nome_lab].apply(normalizar_nome)
            
            # ===== MUDANÇA: Aplicar estabelecimento fixo selecionado pelo usuário =====
            basicos['CD_ESTABELECIMENTO'] = cd_estabelecimento_fixo
            
            # ===== NOVO: Mapear colunas usando a função que testa todas as variações =====
            mapa_basicos = mapear_exames_para_tasy(basicos, categoria='basico')
            
            for col_tasy, col_lab in mapa_basicos.items():
                basicos[col_tasy] = basicos[col_lab].apply(converter_valor_numerico)
            
            basicos['NR_ATENDIMENTO'] = basicos['nome_normalizado'].map(indice_pacientes)
            basicos['nome_original'] = basicos[col_nome_lab]
            sem_atendimento_basicos = basicos[basicos['NR_ATENDIMENTO'].isna()]
            progress_bar.progress(60)
            
            if resultados2 is not None:
                status_text.text("⚙️ Processando resultados 2...")
                col_nome_r2 = detectar_colunas_nome(resultados2)
                resultados2['nome_normalizado'] = resultados2[col_nome_r2].apply(normalizar_nome)
                
                # ===== MUDANÇA: Aplicar estabelecimento fixo selecionado pelo usuário =====
                resultados2['CD_ESTABELECIMENTO'] = cd_estabelecimento_fixo
                
                # ===== NOVO: Mapear colunas usando a função que testa todas as variações =====
                mapa_resultados2 = mapear_exames_para_tasy(resultados2, categoria='resultados2')
                
                for col_tasy, col_lab in mapa_resultados2.items():
                    resultados2[col_tasy] = resultados2[col_lab].apply(converter_valor_numerico)
                
                resultados2['NR_ATENDIMENTO'] = resultados2['nome_normalizado'].map(indice_pacientes)
                sem_atendimento_r2 = resultados2[resultados2['NR_ATENDIMENTO'].isna()]
            else:
                sem_atendimento_r2 = pd.DataFrame()
            
            progress_bar.progress(75)
            
            status_text.text("🔄 Mesclando dados...")
            colunas_necessarias_basicos = ['nome_original', 'dthr_os', 'nome_normalizado', 
                                            'NR_ATENDIMENTO', 'CD_ESTABELECIMENTO']
            colunas_necessarias_basicos.extend([col for col in mapa_basicos.keys() if col in basicos.columns])
            basicos_sel = basicos[colunas_necessarias_basicos]
            
            if resultados2 is not None:
                colunas_necessarias_r2 = ['nome_normalizado', 'dthr_os']
                colunas_necessarias_r2.extend([col for col in mapa_resultados2.keys() if col in resultados2.columns])
                resultados2_sel = resultados2[colunas_necessarias_r2]
                
                dados_mesclados = pd.merge(
                    basicos_sel,
                    resultados2_sel,
                    on=['nome_normalizado', 'dthr_os'],
                    how='outer',
                    suffixes=('', '_r2')
                )
            else:
                dados_mesclados = basicos_sel.copy()
            
            if 'nome_original' not in dados_mesclados.columns or dados_mesclados['nome_original'].isna().any():
                nome_map = pacientes.set_index('nome_normalizado')[col_nome_pac].to_dict()
                if 'nome_original' in dados_mesclados.columns:
                    dados_mesclados['nome_original'] = dados_mesclados['nome_original'].fillna(
                        dados_mesclados['nome_normalizado'].map(nome_map)
                    )
                else:
                    dados_mesclados['nome_original'] = dados_mesclados['nome_normalizado'].map(nome_map)
            
            planilha_final = pd.DataFrame()
            planilha_final['NM_PACIENTE'] = dados_mesclados['nome_original']
            planilha_final['NR_ATENDIMENTO'] = dados_mesclados['NR_ATENDIMENTO']
            planilha_final['DT_RESULTADO'] = dados_mesclados['dthr_os'].apply(formatar_data)
            planilha_final['DS_PROTOCOLO'] = protocolo
            planilha_final['CD_ESTABELECIMENTO'] = dados_mesclados['CD_ESTABELECIMENTO']
            
            for coluna in ORDEM_COLUNAS_TASY[5:]:
                if coluna in dados_mesclados.columns:
                    valor = dados_mesclados[coluna]
                    if isinstance(valor, pd.DataFrame):
                        planilha_final[coluna] = valor.iloc[:, 0]
                    else:
                        planilha_final[coluna] = valor
                else:
                    planilha_final[coluna] = None
            
            planilha_final = planilha_final[planilha_final['NR_ATENDIMENTO'].notna()]
            planilha_final['NR_ATENDIMENTO'] = planilha_final['NR_ATENDIMENTO'].astype(int)
            planilha_final['CD_ESTABELECIMENTO'] = planilha_final['CD_ESTABELECIMENTO'].astype(int)
            
            progress_bar.progress(90)
            
            status_text.text("💾 Gerando arquivo...")
            output = io.BytesIO()
            planilha_final.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
            
            progress_bar.progress(100)
            status_text.text("✅ Concluído!")
            
            st.success(f"✅ Processamento concluído com sucesso!\n\n**Estabelecimento aplicado:** {estabelecimento_selecionado} (Código: {cd_estabelecimento_fixo})")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📊 Total de registros", len(planilha_final))
            with col2:
                st.metric("📋 Protocolo", protocolo)
            with col3:
                st.metric("⚠️ Inconsistências", len(sem_atendimento_basicos) + len(sem_atendimento_r2))
            
            st.download_button(
                label="⬇️ Baixar Planilha para TASY",
                data=output,
                file_name=f"Planilha_Importacao_TASY_{estabelecimento_selecionado}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            if len(sem_atendimento_basicos) + len(sem_atendimento_r2) > 0:
                with st.expander("⚠️ Ver pacientes não encontrados"):
                    nomes_sem_atend = set()
                    if len(sem_atendimento_basicos) > 0:
                        nomes_sem_atend.update(sem_atendimento_basicos['nome_original'].unique())
                    if len(sem_atendimento_r2) > 0:
                        nomes_sem_atend.update(resultados2[resultados2['NR_ATENDIMENTO'].isna()][col_nome_r2].unique())
                    
                    for nome in sorted(nomes_sem_atend):
                        st.write(f"- {nome}")
                    
                    inconsist_output = io.BytesIO()
                    pd.DataFrame({'Paciente': list(nomes_sem_atend)}).to_excel(inconsist_output, index=False)
                    inconsist_output.seek(0)
                    
                    st.download_button(
                        label="⬇️ Baixar Relatório de Inconsistências",
                        data=inconsist_output,
                        file_name="Relatorio_Inconsistencias.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            
            with st.expander("👁️ Visualizar dados processados"):
                st.dataframe(planilha_final.head(20), use_container_width=True)
            
        except Exception as e:
            st.error(f"❌ Erro ao processar: {str(e)}")
            st.exception(e)
