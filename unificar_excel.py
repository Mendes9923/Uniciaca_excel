import streamlit as st
import pandas as pd
import io
import numpy as np

st.set_page_config(page_title="LM Unificador de Excel", layout="centered")

# 🏢 Marca LM micros saas no topo
st.markdown(
    """
    <div style='text-align: center; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 30px;'>
        <h1 style='color: white; margin: 0; font-size: 2.5em;'>📊 LM10 Solutions</h1>
        <p style='color: white; margin: 5px 0 0 0; font-size: 1.1em;'>Unificador de Arquivos Excel</p>
    </div>
    """,
    unsafe_allow_html=True
)

st.title("🔄 Unificador de Arquivos Excel")
st.write("Faça upload de múltiplos arquivos Excel para consolidar em um único arquivo.")

# 📋 Configurações de unificação
col1, col2 = st.columns(2)

with col1:
    # Opção de adicionar coluna com nome do arquivo
    adicionar_coluna_origem = st.checkbox("Adicionar coluna 'Arquivo Origem'", value=True)
    
    # Escolher se vai pular linhas iniciais
    pular_linhas = st.number_input("Pular linhas iniciais", min_value=0, max_value=100, value=0)

with col2:
    # Escolher se vai usar apenas colunas específicas
    usar_colunas_especificas = st.checkbox("Selecionar colunas específicas")
    colunas_para_manter = st.text_input("Colunas (separadas por vírgula)", 
                                        placeholder="Ex: Data, Valor, Descrição",
                                        disabled=not usar_colunas_especificas)

# 📝 Nome do arquivo final
nome_arquivo = st.text_input(
    "Digite o nome para o arquivo final (sem .xlsx)", 
    value="Arquivos_Consolidados"
)

# ⬆️ Upload dos arquivos
uploaded_files = st.file_uploader(
    "Escolha os arquivos Excel (múltiplos arquivos permitidos)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    dataframes = []
    arquivos_com_erro = []
    
    # Barra de progresso
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, uploaded_file in enumerate(uploaded_files):
        try:
            status_text.text(f"Processando: {uploaded_file.name}")
            
            # Ler arquivo Excel pulando linhas se necessário
            if pular_linhas > 0:
                df = pd.read_excel(uploaded_file, skiprows=pular_linhas)
            else:
                df = pd.read_excel(uploaded_file)
            
            # Adicionar coluna com nome do arquivo de origem
            if adicionar_coluna_origem:
                nome_sem_extensao = uploaded_file.name.replace(".xlsx", "").replace(".xls", "")
                df["Arquivo Origem"] = nome_sem_extensao
            
            # Selecionar apenas colunas específicas se necessário
            if usar_colunas_especificas and colunas_para_manter:
                colunas_lista = [col.strip() for col in colunas_para_manter.split(",")]
                colunas_existentes = [col for col in colunas_lista if col in df.columns]
                if colunas_existentes:
                    df = df[colunas_existentes]
                else:
                    st.warning(f"Nenhuma das colunas especificadas encontrada em {uploaded_file.name}")
            
            dataframes.append(df)
            
        except Exception as e:
            arquivos_com_erro.append(f"{uploaded_file.name}: {str(e)}")
        
        # Atualizar progresso
        progress_bar.progress((idx + 1) / len(uploaded_files))
    
    status_text.text("Processamento concluído!")
    
    # Mostrar arquivos com erro
    if arquivos_com_erro:
        with st.expander("⚠️ Arquivos com erro"):
            for erro in arquivos_com_erro:
                st.error(erro)
    
    # Unir todos os DataFrames se houver dados
    if dataframes:
        df_final = pd.concat(dataframes, ignore_index=True)
        
        st.success(f"✅ {len(dataframes)} arquivos combinados com sucesso!")
        
        # Estatísticas básicas
        st.subheader("📊 Resumo da Consolidação")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total de Linhas", len(df_final))
        with col2:
            st.metric("Total de Colunas", len(df_final.columns))
        with col3:
            st.metric("Arquivos Processados", len(dataframes))
        
        # Mostrar preview dos dados
        st.subheader("🔍 Preview dos Dados Consolidados")
        st.dataframe(df_final.head(100))
        
        # Botão para mostrar dados completos
        if st.button("Mostrar dados completos"):
            st.dataframe(df_final)
        
        # 💾 Gerar arquivo Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Escrever dados consolidados
            df_final.to_excel(writer, index=False, sheet_name='Consolidado')
            
            # Adicionar uma sheet com metadados
            metadata = pd.DataFrame({
                "Informação": ["Data de Processamento", "Total de Arquivos", "Total de Linhas", "Total de Colunas"],
                "Valor": [
                    pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                    len(dataframes),
                    len(df_final),
                    len(df_final.columns)
                ]
            })
            metadata.to_excel(writer, index=False, sheet_name='Metadados')
            
            # Ajustar largura das colunas automaticamente (VERSÃO CORRIGIDA)
            worksheet = writer.sheets['Consolidado']
            for i, col in enumerate(df_final.columns):
                try:
                    # Converter para string e lidar com valores nulos
                    series_str = df_final[col].astype(str)
                    # Substituir 'nan' por string vazia para não contar como largura
                    series_str = series_str.replace('nan', '')
                    # Calcular o comprimento máximo
                    max_len = series_str.str.len().max()
                    # Se não houver dados ou for NaN, usar o tamanho do cabeçalho
                    if pd.isna(max_len) or max_len == 0:
                        max_len = len(str(col))
                    column_width = max(max_len, len(str(col))) + 2
                    # Limitar a largura máxima
                    worksheet.set_column(i, i, min(column_width, 50))
                except Exception as e:
                    # Em caso de erro, usar largura padrão
                    worksheet.set_column(i, i, 15)
        
        # 🔽 Botão de download
        st.download_button(
            label="📥 Baixar Excel Consolidado",
            data=output.getvalue(),
            file_name=f"{nome_arquivo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    else:
        st.error("Nenhum arquivo foi processado com sucesso. Verifique os formatos dos arquivos.")

# 📖 Instruções de uso
with st.expander("📖 Como usar"):
    st.markdown("""
    ### Instruções:
    1. **Faça upload** de um ou mais arquivos Excel (.xlsx ou .xls)
    2. **Configure as opções** de unificação:
       - Adicionar coluna com nome do arquivo de origem
       - Pular linhas iniciais se necessário
       - Selecionar apenas colunas específicas
    3. **Defina o nome** do arquivo final
    4. **Clique em Baixar** para obter o arquivo consolidado
    
    ### Recursos:
    - ✅ Suporta arquivos .xlsx e .xls
    - ✅ Mantém formatação básica dos dados
    - ✅ Adiciona metadados do processamento
    - ✅ Preview dos dados antes de baixar
    - ✅ Tratamento de erros por arquivo
    """)

# 🔽 Rodapé
st.markdown(
    """
    <hr style="margin-top:50px">
    <div style='text-align: center; color: grey; padding: 20px;'>
        <strong>LM10 Solutions</strong> - Soluções em Software como Serviço<br>
        Unificador de Excel - Versão Gratuita
    </div>
    """,
    unsafe_allow_html=True
)