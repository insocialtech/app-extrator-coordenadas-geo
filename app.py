import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Extrator de Coordenadas Geográficas", layout="wide")

def extrair_coordenadas(texto):
    """
    Extrai todos os pares de coordenadas do texto.
    Retorna lista de dicionários com latitude e longitude em decimal e DMS.
    Suporta formatos comuns em português e inglês.
    """
    if pd.isna(texto) or texto.strip().upper() in ['NÃO CONSTA', 'NAO CONSTA', 'NOT INFORMED', '']:
        return []
    
    # Regex para padrões tipo 3º03'52,9838"S e 59º54'46,6013"W
    padrao = r'(\d{1,3})[º°](\d{1,2})\'([\d,\.]+)"?\s*([NS])[\s,;eE]*(\d{1,3})[º°](\d{1,2})\'([\d,\.]+)"?\s*([WO])'
    matches = re.findall(padrao, texto)
    resultados = []
    for match in matches:
        graus_lat, min_lat, seg_lat, hem_lat, graus_lon, min_lon, seg_lon, hem_lon = match
        # Corrige vírgula e ponto nos segundos
        seg_lat = seg_lat.replace(',', '.')
        seg_lon = seg_lon.replace(',', '.')
        # Converte para decimal
        lat = int(graus_lat) + int(min_lat)/60 + float(seg_lat)/3600
        if hem_lat.upper() == 'S':
            lat = -lat
        lon = int(graus_lon) + int(min_lon)/60 + float(seg_lon)/3600
        if hem_lon.upper() == 'W':
            lon = -lon
        # DMS formatado
        lat_dms = f"{graus_lat}º{min_lat}'{seg_lat}\"{hem_lat}"
        lon_dms = f"{graus_lon}º{min_lon}'{seg_lon}\"{hem_lon}"
        resultados.append({
            'LATITUDE': lat,
            'LONGITUDE': lon,
            'LATITUDE_DMS': lat_dms,
            'LONGITUDE_DMS': lon_dms
        })
    return resultados

def expandir_dataframe(df, coluna_coordenadas):
    """
    Expande o DataFrame para uma linha por ponto extraído.
    """
    linhas_expandidas = []
    for idx, row in df.iterrows():
        pontos = extrair_coordenadas(str(row[coluna_coordenadas]))
        for i, ponto in enumerate(pontos, 1):
            nova_linha = row.to_dict()
            nova_linha['PONTO'] = f"P{i:02d}"
            nova_linha.update(ponto)
            linhas_expandidas.append(nova_linha)
    return pd.DataFrame(linhas_expandidas)

def configurar_exportacao(df):
    """Padroniza os dados para exportação, garantindo compatibilidade internacional"""
    if df is None or df.empty:
        return pd.DataFrame()
    # Garante que todos os textos estejam em unicode e sem caracteres problemáticos
    df = df.applymap(lambda x: x.encode('utf-8', 'ignore').decode('utf-8') if isinstance(x, str) else x)
    # Padroniza decimais
    for col in ['LATITUDE', 'LONGITUDE']:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: f"{x:.6f}" if isinstance(x, (int, float, float)) else x)
    df.fillna('N/A', inplace=True)
    return df

def gerar_nome_arquivo(base_name, ext):
    """Gera nome de arquivo único com timestamp"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base_name}_{timestamp}.{ext}"

st.title("🌍 Extrator Universal de Coordenadas Geográficas")

arquivo = st.file_uploader("Carregue seu arquivo Excel (.xlsx)", type=["xlsx"])

if arquivo:
    try:
        xls = pd.ExcelFile(arquivo)
        aba = st.selectbox("Selecione a aba:", xls.sheet_names)
        df = pd.read_excel(arquivo, sheet_name=aba)
        coluna_coord = st.selectbox("Selecione a coluna de coordenadas:", df.columns)
        
        if st.button("Processar e Expandir"):
            with st.spinner("Processando coordenadas..."):
                df_expandido = expandir_dataframe(df, coluna_coord)
                if not df_expandido.empty:
                    df_export = configurar_exportacao(df_expandido)
                    st.success(f"Processamento concluído! {len(df_export)} pontos extraídos.")
                    st.dataframe(df_export.head(10))
                    
                    # CSV UTF-8 BOM (compatível Excel/Windows)
                    csv_utf8_bom = df_export.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                    # CSV padrão inglês (delimiter=comma, decimal=dot)
                    csv_en = df_export.to_csv(index=False, sep=',', decimal='.', encoding='utf-8').encode('utf-8')
                    # Excel
                    excel_buffer = BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                        df_export.to_excel(writer, index=False, sheet_name='Coordenadas')
                        writer.close()
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.download_button(
                            label="⬇️ Baixar CSV (UTF-8 BOM, Excel/Windows)",
                            data=csv_utf8_bom,
                            file_name=gerar_nome_arquivo("coordenadas_expandido", "csv"),
                            mime='text/csv'
                        )
                    with col2:
                        st.download_button(
                            label="⬇️ Baixar CSV (Inglês/Universal)",
                            data=csv_en,
                            file_name=gerar_nome_arquivo("coordinates_expanded", "csv"),
                            mime='text/csv'
                        )
                    with col3:
                        st.download_button(
                            label="⬇️ Baixar Excel (.xlsx)",
                            data=excel_buffer.getvalue(),
                            file_name=gerar_nome_arquivo("coordenadas_expandido", "xlsx"),
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                    st.info("Você pode baixar o resultado em qualquer formato acima. Todos suportam caracteres especiais.")
                else:
                    st.warning("Nenhum ponto de coordenada encontrado para exportação.")
    except Exception as e:
        st.error(f"Erro no processamento: {str(e)}")
        st.stop()

st.markdown("""
---
**Exportação multilíngue:**  
- CSV UTF-8 BOM: Compatível com Excel (Windows/Português)
- CSV Universal: Inglês, separador vírgula, decimal ponto
- Excel (.xlsx): Totalmente compatível com acentos, português e inglês
""")
