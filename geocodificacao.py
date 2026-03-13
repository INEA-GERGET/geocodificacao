import pandas as pd
import geopy
import time
import numpy as np

nome_arquivo = "SERVMOLA_dados.xlsx"

# 1. Configuração do Geocodificador
geolocator = geopy.geocoders.ArcGIS(timeout=10)

MUNICIPIOS_RJ = [
    "ANGRA DOS REIS", "APERIBÉ", "ARARUAMA", "AREAL", "ARMAÇÃO DOS BÚZIOS", 
    "ARRAIAL DO CABO", "BARRA DO PIRAÍ", "BARRA MANSA", "BELFORD ROXO", 
    "BOM JARDIM", "BOM JESUS DO ITABAPOANA", "CABO FRIO", "CACHOEIRAS DE MACACU", 
    "CAMBUCI", "CAMPOS DOS GOYTACAZES", "CANTAGALO", "CARAPEBUS", "CARDOSO MOREIRA", 
    "CARMO", "CASIMIRO DE ABREU", "COMENDADOR LEVY GASPARIAN", "CONCEIÇÃO DE MACABU", 
    "DUAS BARRAS", "DUQUE DE CAXIAS", "ENGENHEIRO PAULO DE FRONTINHO", "GUAPIMIRIM", 
    "IGUABA GRANDE", "ITABORAÍ", "ITAGUAÍ", "ITALVA", "ITAOCARA", "ITAPERUNA", 
    "ITATIAIA", "JAPERI", "LAJE DO MURIAÉ", "MACAÉ", "MACUCO", "MAGÉ", "MANGARATIBA", 
    "MARICÁ", "MENDES", "MESQUITA", "MIGUEL PEREIRA", "MIRACEMA", "NATIVIDADE", 
    "NILÓPOLIS", "NITERÓI", "NOVA FRIBURGO", "NOVA IGUAÇU", "PARACAMBI", 
    "PARAÍBA DO SUL", "PARATY", "PATY DO ALFERES", "PETRÓPOLIS", "PINHEIRAL", 
    "PIRAÍ", "PORCIÚNCULA", "PORTO REAL", "QUATIS", "QUEIMADOS", "QUISSAMÃ", 
    "RESENDE", "RIO BONITO", "RIO DAS FLORES", "RIO DAS OSTRAS", "RIO DE JANEIRO", 
    "SANTA MARIA MADALENA", "SANTO ANTÔNIO DE PÁDUA", "SÃO FIDÉLIS", "SÃO FRANCISCO DE ITABAPOANA", 
    "SÃO GONÇALO", "SÃO JOÃO DA BARRA", "SÃO JOÃO DE MERITI", "SÃO JOSÉ DE UBÁ", 
    "SÃO JOSÉ DO VALE DO RIO PRETO", "SÃO PEDRO DA ALDEIA", "SÃO SEBASTIÃO DO ALTO", 
    "SAPUCAIA", "SAQUAREMA", "SEROPÉDICA", "SILVA JARDIM", "SUMIDOURO", "TANGUÁ", 
    "TERESÓPOLIS", "TRAJANO DE MORAES", "TRÊS RIOS", "VALENÇA", "VARRE-SAI", 
    "VASSOURAS", "VOLTA REDONDA"
]

def processar_geocodificacao(linha):
    # --- NOVA VALIDAÇÃO: Pular se coord_origem já existir ---
    origem_atual = str(linha.get('coord_origem', '')).strip()
    lat_atual = linha.get('Latitude')
    lon_atual = linha.get('Longitude')
    
    if origem_atual not in ['', 'nan', 'None']:
        # Se já tem algo escrito em coord_origem, retorna os valores atuais sem mexer
        return pd.Series([lat_atual, lon_atual, origem_atual])
    # -------------------------------------------------------

    lat_orig = str(lat_atual).strip() if pd.notna(lat_atual) else ''
    lon_orig = str(lon_atual).strip() if pd.notna(lon_atual) else ''
    municipio = str(linha.get('Município', '')).strip().upper()
    endereco = str(linha.get('Endereço', '')).strip()

    is_lat_empty = pd.isna(lat_atual) or lat_orig in ['', '-', 'None', 'nan']
    is_lon_empty = pd.isna(lon_atual) or lon_orig in ['', '-', 'None', 'nan']

    if not (is_lat_empty or is_lon_empty):
        return pd.Series([lat_orig, lon_orig, "SERVMOLA"])
    
    else: 
        origem_automatica = "gerado automaticamente pelo geopy-arcgis"
        if municipio not in MUNICIPIOS_RJ:
            return pd.Series(["Não se aplica", "Não se aplica", origem_automatica])
            
        else:
            query = f"{endereco}, {municipio}, Brasil"
            try:
                location = geolocator.geocode(query)
                time.sleep(0.5) 
                if location:
                    return pd.Series([location.latitude, location.longitude, origem_automatica])
                else:
                    return pd.Series([None, None, origem_automatica])
            except Exception:
                return pd.Series([None, None, origem_automatica])

# 3. Execução do Loop de Guias
todas_guias = pd.read_excel(nome_arquivo, sheet_name=None)
meus_dfs = {}

for guia, df in todas_guias.items():
    if guia == "Análise":
        print(f"Pulando guia: {guia} (mantendo dados originais)")
        meus_dfs[guia] = df
        continue
        
    print(f"Processando guia: {guia}")
    
    # Garantir que a coluna coord_origem exista antes de rodar o apply para evitar erros de busca
    if 'coord_origem' not in df.columns:
        df['coord_origem'] = np.nan

    df[['Latitude', 'Longitude', 'coord_origem']] = df.apply(processar_geocodificacao, axis=1)
    meus_dfs[guia] = df
    print(f"Guia {guia} concluída.")

# 4. Salvar arquivo final

with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
    for nome_aba, df_processado in meus_dfs.items():
        df_processado.to_excel(writer, sheet_name=nome_aba, index=False)
        print(f"Aba '{nome_aba}' salva com sucesso.")

print(f"\n✅ Processo concluído! Arquivo: '{nome_arquivo}'")
