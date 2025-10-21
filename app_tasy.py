import pandas as pd
import numpy as np
from datetime import datetime
import warnings
import streamlit as st
from pathlib import Path
import io
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

# ==================== TÍTULO ====================
st.title("🏥 Gerador de Planilha para TASY")
st.markdown("---")

# ==================== CONFIGURAÇÕES ====================
ESTABELECIMENTOS = {
    ' MATRIZ': 1, ' MONTE SERRAT': 3, ' CONVENIOS': 4,
    ' RIO VERMELHO': 5, ' SANTO ESTEVAO': 7,
    'MATRIZ': 1, 'MONTE SERRAT': 3, 'CONVENIOS': 4,
    'RIO VERMELHO': 5, 'SANTO ESTEVAO': 7
}

# Mapeamento completo com nomes dos exames
MAPA_EXAMES_COMPLETO = {
    'C_Hb': {'codigo': 'NR_EXAME_36486', 'nome': 'Hemoglobina (Hb)'},
    'hb ': {'codigo': 'NR_EXAME_36486', 'nome': 'Hemoglobina (Hb)'},
    'hb': {'codigo': 'NR_EXAME_36486', 'nome': 'Hemoglobina (Hb)'},
    'C_Ht': {'codigo': 'NR_EXAME_36485', 'nome': 'Hematócrito (Ht)'},
    'ht ': {'codigo': 'NR_EXAME_36485', 'nome': 'Hematócrito (Ht)'},
    'ht': {'codigo': 'NR_EXAME_36485', 'nome': 'Hematócrito (Ht)'},
    'UREI': {'codigo': 'NR_EXAME_36452', 'nome': 'Uréia Pré'},
    'UPD': {'codigo': 'NR_EXAME_36581', 'nome': 'Uréia Pós-Diálise'},
    'CREA': {'codigo': 'NR_EXAME_36434', 'nome': 'Creatinina'},
    'crea ': {'codigo': 'NR_EXAME_36434', 'nome': 'Creatinina'},
    'crea': {'codigo': 'NR_EXAME_36434', 'nome': 'Creatinina'},
    'CALCIO': {'codigo': 'NR_EXAME_36433', 'nome': 'Cálcio'},
    'cálcio': {'codigo': 'NR_EXAME_36433', 'nome': 'Cálcio'},
    'FOSFS': {'codigo': 'NR_EXAME_36435', 'nome': 'Fósforo'},
    'Na': {'codigo': 'NR_EXAME_36461', 'nome': 'Sódio'},
    'POTAS': {'codigo': 'NR_EXAME_36436', 'nome': 'Potássio (K)'},
    'TGP': {'codigo': 'NR_EXAME_36437', 'nome': 'TGP (ALT)'},
    'GLIC': {'codigo': 'NR_EXAME_36438', 'nome': 'Glicose'},
    'CTOT': {'codigo': 'NR_EXAME_36447', 'nome': 'Colesterol Total'},
    'HDL': {'codigo': 'NR_EXAME_36579', 'nome': 'HDL Colesterol'},
    'Col_LDL': {'codigo': 'NR_EXAME_36580', 'nome': 'LDL Colesterol'},
    'TRIG': {'codigo': 'NR_EXAME_36449', 'nome': 'Triglicerídeos'},
    'PLAQ': {'codigo': 'NR_EXAME_36483', 'nome': 'Plaquetas'},
    'FA': {'codigo': 'NR_EXAME_36522', 'nome': 'Fosfatase Alcalina'},
    'PT': {'codigo': 'NR_EXAME_36523', 'nome': 'Proteínas Totais'},
    'ALB': {'codigo': 'NR_EXAME_36584', 'nome': 'Albumina'},
    'GLB': {'codigo': 'NR_EXAME_36585', 'nome': 'Globulinas'},
    'Rel_Alb_Gl': {'codigo': 'NR_EXAME_36587', 'nome': 'Relação Albumina/Globulina'},
    'Hb_A1c': {'codigo': 'NR_EXAME_36453', 'nome': 'Hemoglobina Glicada (HbA1c)'},
    'Ferritina': {'codigo': 'NR_EXAME_36518', 'nome': 'Ferritina'},
    'T4': {'codigo': 'NR_EXAME_36450', 'nome': 'T4 Total'},
    'TSH': {'codigo': 'NR_EXAME_36455', 'nome': 'TSH'},
    'FER': {'codigo': 'NR_EXAME_36567', 'nome': 'Ferro Sérico'},
    'CTT': {'codigo': 'NR_EXAME_36582', 'nome': 'Capacidade Total de Fixação'},
    'IST': {'codigo': 'NR_EXAME_36520', 'nome': 'Índice Saturação Transferrina'},
    'VITD25OH': {'codigo': 'NR_EXAME_36457', 'nome': 'Vitamina D (25-OH)'},
    'PTH_DB': {'codigo': 'NR_EXAME_36502', 'nome': 'PTH (Paratormônio)'},
    'AAU': {'codigo': 'NR_EXAME_36578', 'nome': 'Anticorpo Anti-HBs'},
    'ANTI_HBS': {'codigo': 'NR_EXAME_36456', 'nome': 'Anti-HBs (Hepatite B)'},
    'HCV': {'codigo': 'NR_EXAME_36574', 'nome': 'Anti-HCV (Hepatite C)'},
    'ALU_SER': {'codigo': 'NR_EXAME_36448', 'nome': 'Alumínio Sérico'}
}

# Criar mapas simples para processamento
MAPA_BASICOS = {}
MAPA_RESULTADOS2 = {}

for key, value in MAPA_EXAMES_COMPLETO.items():
    if key in ['C_Hb', 'hb ', 'hb', 'C_Ht', 'ht ', 'ht', 'UREI', 'UPD', 'CREA', 'crea ', 'crea', 
               'CALCIO', 'cálcio', 'FOSFS', 'Na', 'POTAS', 'TGP', 'GLIC']:
        MAPA_BASICOS[key] = value['codigo']
    else:
        MAPA_RESULTADOS2[key] = value['codigo']

ORDEM_COLUNAS_TASY = [
    'NM_PACIENTE', 'NR_ATENDIMENTO', 'DT_RESULTADO', 'DS_PROTOCOLO', 'CD_ESTABELECIMENTO',
    'NR_EXAME_36433', 'NR_EXAME_36434', 'NR_EXAME_36435', 'NR_EXAME_36436', 'NR_EXAME_36437',
    'NR_EXAME_36438', 'NR_EXAME_36439', 'NR_EXAME_36447', 'NR_EXAME_36448', 'NR_EXAME_36449',
    'NR_EXAME_36450', 'NR_EXAME_36452', 'NR_EXAME_36453', 'NR_EXAME_36455', 'NR_EXAME_36456',
    'NR_EXAME_36457', 'NR_EXAME_36461', 'NR_EXAME_36483', 'NR_EXAME_36485', 'NR_EXAME_36486',
    'NR_EXAME_36501', 'NR_EXAME_36502', 'NR_EXAME_36518', 'NR_EXAME_36520', 'NR_EXAME_36522',
    'NR_EXAME_36523', 'NR_EXAME_36567', 'NR_EXAME_36574', 'NR_EXAME_36578', 'NR_EXAME_36579',
    'NR_EXAME_36580', 'NR_EXAME_36581', 'NR_EXAME_36582', 'NR_EXAME_36584', 'NR_EXAME_36585',
    'NR_EXAME_36587', 'NR_EXAME_36588'
]

# ==================== FUNÇÕES ====================
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

def obter_cd_estabelecimento(setor):
    if pd.isna(setor):
        return None
    setor_upper = str(setor).upper().strip()
    return ESTABELECIMENTOS.get(setor_upper, ESTABELECIMENTOS.get(' ' + setor_upper))

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
    
    protocolo = st.selectbox(
        "Protocolo",
        ["MENSAL", "TRIMESTRAL", "SEMESTRAL", "ANUAL"],
        index=0
    )
    
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
        2. **Faça upload dos arquivos**:
           - Exames Básicos (obrigatório)
           - Resultados 2 (opcional)
           - Pacientes do TASY (obrigatório)
        3. **Clique em Processar**
        4. **Baixe o arquivo** gerado
        
        ---
        
        ### ✅ Formatos aceitos:
        - `.xlsx` (Excel 2007+)
        - `.xls` (Excel 97-2003)
        
        ### 🔍 O que o sistema faz:
        - Cruza dados por nome do paciente
        - Mapeia exames para códigos TASY
        - Valida estabelecimentos
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
        {"Código": 1, "Estabelecimento": "MATRIZ"},
        {"Código": 3, "Estabelecimento": "MONTE SERRAT"},
        {"Código": 4, "Estabelecimento": "CONVÊNIOS"},
        {"Código": 5, "Estabelecimento": "RIO VERMELHO"},
        {"Código": 7, "Estabelecimento": "SANTO ESTEVÃO"}
    ])
    
    st.dataframe(df_estabelecimentos, use_container_width=True, hide_index=True)

else:
    # [RESTO DO CÓDIGO DE PROCESSAMENTO PERMANECE IGUAL]
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
            basicos['CD_ESTABELECIMENTO'] = basicos['setor_solic'].apply(obter_cd_estabelecimento)
            
            for col_lab, col_tasy in MAPA_BASICOS.items():
                if col_lab in basicos.columns:
                    basicos[col_tasy] = basicos[col_lab].apply(converter_valor_numerico)
            
            basicos['NR_ATENDIMENTO'] = basicos['nome_normalizado'].map(indice_pacientes)
            basicos['nome_original'] = basicos[col_nome_lab]
            sem_atendimento_basicos = basicos[basicos['NR_ATENDIMENTO'].isna()]
            progress_bar.progress(60)
            
            if resultados2 is not None:
                status_text.text("⚙️ Processando resultados 2...")
                col_nome_r2 = detectar_colunas_nome(resultados2)
                resultados2['nome_normalizado'] = resultados2[col_nome_r2].apply(normalizar_nome)
                resultados2['CD_ESTABELECIMENTO'] = resultados2['setor_solic'].apply(obter_cd_estabelecimento)
                
                for col_lab, col_tasy in MAPA_RESULTADOS2.items():
                    if col_lab in resultados2.columns:
                        resultados2[col_tasy] = resultados2[col_lab].apply(converter_valor_numerico)
                
                resultados2['NR_ATENDIMENTO'] = resultados2['nome_normalizado'].map(indice_pacientes)
                sem_atendimento_r2 = resultados2[resultados2['NR_ATENDIMENTO'].isna()]
            else:
                sem_atendimento_r2 = pd.DataFrame()
            
            progress_bar.progress(75)
            
            status_text.text("🔄 Mesclando dados...")
            colunas_necessarias_basicos = ['nome_original', 'dthr_os', 'nome_normalizado', 
                                            'NR_ATENDIMENTO', 'CD_ESTABELECIMENTO']
            colunas_necessarias_basicos.extend([col for col in MAPA_BASICOS.values() if col in basicos.columns])
            basicos_sel = basicos[colunas_necessarias_basicos]
            
            if resultados2 is not None:
                colunas_necessarias_r2 = ['nome_normalizado', 'dthr_os']
                colunas_necessarias_r2.extend([col for col in MAPA_RESULTADOS2.values() if col in resultados2.columns])
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
            
            st.success("✅ Processamento concluído com sucesso!")
            
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
                file_name=f"Planilha_Importacao_TASY_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
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
