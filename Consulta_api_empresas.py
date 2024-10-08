import tkinter as tk
from tkinter import messagebox
import threading
import requests
import json
import pandas as pd
import time
import datetime
from PIL import Image, ImageTk

# Documentação da API
# https://www.cnpj.ws/docs/api-publica/consultando-cnpj

def previsao_tempo(total_cnpjs, taxa_consulta):
    tempo_minutos = (total_cnpjs // taxa_consulta)
    return tempo_minutos  # Retorna o tempo em minutos

def calcular_hora_prevista(tempo_minutos):
    hora_atual = datetime.datetime.now()
    hora_prevista = hora_atual + datetime.timedelta(minutes=tempo_minutos)
    return hora_prevista  # Retorna a hora prevista de fim

def processar_cnpjs():
    # Ler os CNPJs de um arquivo Excel
    cnpjs_df = pd.read_excel("cnpjs.xlsx")
    cnpjs = cnpjs_df["CNPJ"].tolist()

    resultados = []
    consulta_count = 0

    print('iniciado')

    tempo_previsao = previsao_tempo(len(cnpjs), 3)  # Calcula a previsão de tempo em minutos
    hora_prevista = calcular_hora_prevista(tempo_previsao)  # Calcula a hora prevista de fim

    previsao_label.config(text=f"Previsão de tempo para processar {len(cnpjs)} CNPJs: {tempo_previsao} minutos\nHora prevista de fim: {hora_prevista.strftime('%H:%M')}")  # Atualiza o texto do rótulo com a previsão de tempo e hora prevista de fim

    # Itera sobre cada CNPJ na lista
    for cnpj in cnpjs:
        url = f"https://publica.cnpj.ws/cnpj/{cnpj}"

        resp = requests.get(url)
        if resp.status_code == 200:
            data = json.loads(resp.text)

            razao_social = data["razao_social"]
            natureza_juridica_descricao = data["natureza_juridica"]["descricao"]
            cep = data["estabelecimento"]["cep"]
            tipo_logradouro = data["estabelecimento"]["tipo_logradouro"]
            logradouro = data["estabelecimento"]["logradouro"]
            bairro = data["estabelecimento"]["bairro"]
            numero = data["estabelecimento"]["numero"]
            atividade_principal = data["estabelecimento"]["atividade_principal"]["subclasse"]
            atividade_principal_descricao = data["estabelecimento"]["atividade_principal"]["descricao"]
            cidade = data["estabelecimento"]["cidade"]["nome"]
            ibge_id = data["estabelecimento"]['cidade']['ibge_id']
            estado = data["estabelecimento"]["estado"]["nome"]

            if "simples" in data and data["simples"] is not None:
                optante = data["simples"]["simples"]
            else:
                optante = "Não"

            tipo = data["estabelecimento"]["tipo"]

            resultados.append({"CNPJ": cnpj, "Razão Social": razao_social, "Natureza jurídica descrição":natureza_juridica_descricao,"cep": cep,
                               "tipo_logradouro": tipo_logradouro, "logradouro": logradouro,
                               "bairro": bairro, "numero": numero, "Cidade": cidade,
                               "IBGE ID": ibge_id, "Estado": estado,
                               "Optante Simples Nacional": optante, "Tipo": tipo,
                               "Atividade principal": atividade_principal,
                               "Atividade principal descricao":atividade_principal_descricao})

        consulta_count += 1

        if consulta_count % 3 == 0:
            time.sleep(60)

    # Cria um DataFrame Pandas com os resultados
    df = pd.DataFrame(resultados)

    # Salva o DataFrame em um arquivo Excel
    df.to_excel("resultado_cnpjs.xlsx", index=False)
    messagebox.showinfo("Fim", "O processamento foi concluído!")

def iniciar_processamento():
    process_thread = threading.Thread(target=processar_cnpjs)
    process_thread.start()

# Criar uma janela
root = tk.Tk()
root.title("Processamento de CNPJs")

# Definir tamanho fixo da tela
largura = 300
altura = 300
x = (root.winfo_screenwidth() // 2) - (largura // 2)
y = (root.winfo_screenheight() // 2) - (altura // 2)
root.geometry(f'{largura}x{altura}+{x}+{y}')

imagem = Image.open("Toad.jpg")
imagem = imagem.resize((100, 100))
imagem = ImageTk.PhotoImage(imagem)

# Exibir a imagem em um rótulo
label_imagem = tk.Label(root, image=imagem)
label_imagem.image = imagem  # Mantenha uma referência para evitar a coleta de lixo
label_imagem.pack(pady=10)

# Definir cor de fundo
root.configure(bg='darkblue')

# Rótulo para exibir a previsão de tempo e hora prevista de fim
previsao_label = tk.Label(root, text="", bg='darkblue', fg='white', wraplength=250, justify="left")
previsao_label.pack(pady=10)

# Adicionar um botão para iniciar o processamento dos CNPJs
process_button = tk.Button(root, text="Iniciar Processamento", command=iniciar_processamento)
process_button.pack(pady=20)

# Iniciar o loop principal da interface gráfica
root.mainloop()
