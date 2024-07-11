# Processamento de CNPJs com Interface Gráfica

Este projeto é uma aplicação em Python que utiliza a biblioteca Tkinter para criar uma interface gráfica de usuário (GUI) para processar CNPJs, consultando uma API pública e salvando os resultados em um arquivo Excel.

## Funcionalidades

- Leitura de CNPJs de um arquivo Excel.
- Consulta de informações dos CNPJs em uma API pública.
- Exibição de uma previsão do tempo de processamento e a hora prevista de término.
- Salvamento dos resultados da consulta em um arquivo Excel.
- Interface gráfica com Tkinter para facilitar a interação do usuário.

## Requisitos

- Python 3.x
- Bibliotecas Python:
  - `tkinter`
  - `requests`
  - `pandas`
  - `Pillow`

## Instalação

1. Clone este repositório para sua máquina local:
    ```sh
    git clone https://github.com/seu-usuario/processamento-cnpjs.git
    ```

2. Navegue até o diretório do projeto:
    ```sh
    cd processamento-cnpjs
    ```

3. Instale as dependências necessárias:
    ```sh
    pip install -r requirements.txt
    ```

4. Certifique-se de ter o arquivo `cnpjs.xlsx` no mesmo diretório do script. Este arquivo deve conter uma coluna chamada "CNPJ" com os números de CNPJs a serem processados.

5. Adicione uma imagem chamada `Toad.jpg` no mesmo diretório do script para ser exibida na interface gráfica.

## Uso

1. Execute o script principal:
    ```sh
    python script.py
    ```

2. A interface gráfica será aberta. Clique no botão "Iniciar Processamento" para começar a processar os CNPJs.

3. O script irá consultar a API para cada CNPJ, exibir uma previsão do tempo de processamento e a hora prevista de término, e salvar os resultados em um arquivo Excel chamado `resultado_cnpjs.xlsx`.

## API Utilizada

Este projeto utiliza a API pública para consulta de CNPJs. A documentação da API pode ser encontrada em [https://www.cnpj.ws/docs/api-publica/consultando-cnpj](https://www.cnpj.ws/docs/api-publica/consultando-cnpj).

## Estrutura do Código

- `previsao_tempo(total_cnpjs, taxa_consulta)`: Calcula a previsão de tempo para processar os CNPJs.
- `calcular_hora_prevista(tempo_minutos)`: Calcula a hora prevista de término do processamento.
- `processar_cnpjs()`: Função principal que lê os CNPJs de um arquivo Excel, consulta a API, e salva os resultados em um novo arquivo Excel.
- `iniciar_processamento()`: Inicia o processamento dos CNPJs em uma thread separada.
- Criação da interface gráfica com Tkinter.

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues e pull requests para melhorar este projeto.
