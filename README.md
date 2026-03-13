# Geocodificação - SERVMOLA
## Sumário
1. [Descrição](#Descrição)
2. [Uso](#Uso)
3. [Instalação](#Instalação)


## Descrição
Este script realiza a geocodificação com base no endereço e município dos dados SERVMOLA e gera a Latitude e Longitude, quando este dado está faltando. Também cria uma coluna de origem da coordenada (`coord_origem`) que mostra a origem da Latitude e Longitude, podendo ser própria da SERVMOLA ou ter sido gerada automaticamente por este script.


## Uso
Em uma pasta coloque a planilha de dados `SERVMOLA_dados.xlsx` e esse script. No seu compilador de preferência, execute o arquivo `geocodificacao.py`, espere um aviso de "Processo concluído!".  

## Instalação
Os arquivos .txt são para serem utilizados dentro do [ArcGIS](https://www.arcgis.com/index.html) e os .py em um compilador Python, como explicado no [mapa mental](#Uso).
É recomendado criar um ambiente virtual a partir do ArcGIS no seu compilador para fazer as intealações das bibliotecas necessárias.

Para utilizar os scripts é necessária a instalação das bibliotecas Python:

* **geopy**: Para acessar serviços de geocodificação de diversos provedores (ArcGIS, neste caso) e calcular distâncias geodésicas.
* **pandas**: Para manipulação e análise de dados tabulares.
* **time**: Para gerenciar intervalos de tempo e pausas (sleep) entre requisições de APIs, evitando bloqueios.
* **numpy**: Para suporte a arrays multidimensionais e funções matemáticas de alto nível, utilizada na manipulação de valores nulos (NaN).

Se necessário, coloque o seguinte comando no terminal:
`pip install pandas geopy time numpy`
