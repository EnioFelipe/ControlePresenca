import os.path
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

'''
 Código feito por:
 Enio Felipe Botelho Miguel
 Fabio Eizo Rodriguez Matsumoto
 João Victor Moreira Tamassia

 Integração Python com API GOOGLE
 Para atualização de planilhas - Sexta esportiva - UTFPR/CP - DACOMP
'''



# Se modificar esses escopos, exclua o arquivo token.json.
ESCOPOS = ["https://www.googleapis.com/auth/spreadsheets"]

# ID e intervalo de uma planilha de exemplo.
ID_PLANILHA_EXEMPLO = "18muX2uGgFQL4F-x-c"
INTERVALO_PLANILHA_EXEMPLO = "Respostas ao formulário 2!A1:D"
INTERVALO_CONTROLE_PRESENCA = "Controle de Presença!A1:C"
INTERVALO_QUANTIDADE_PRESENCA = "Quantidade Presença!A1:B"


def main():
    credenciais = None
    
    if os.path.exists("token.json"):
        credenciais = Credentials.from_authorized_user_file("token.json", ESCOPOS)
    # Se não houver credenciais (válidas), permita que o usuário faça login.
    if not credenciais or not credenciais.valid:
        if credenciais and credenciais.expired and credenciais.refresh_token:
            credenciais.refresh(Request())
        else:
            fluxo = InstalledAppFlow.from_client_secrets_file(
                "client_secret.json", ESCOPOS
            )
            credenciais = fluxo.run_local_server(port=0)
        # Salve as credenciais para a próxima execução
        with open("token.json", "w") as token:
            token.write(credenciais.to_json())

    try:
        servico = build("sheets", "v4", credentials=credenciais)

        # Chame a API do Sheets
        planilha = servico.spreadsheets()
        resultado = planilha.values().get(spreadsheetId=ID_PLANILHA_EXEMPLO, range=INTERVALO_PLANILHA_EXEMPLO).execute()
        valores = resultado.get("values", [])
        print(valores)

        # Verifique se há novas respostas do formulário
        if valores:
            for linha in valores:
                if len(linha) >= 4:
                    timestamp, email, nome, ra = linha
                    
                    # Verifique se o RA e o timestamp já existem na planilha de controle de presença
                    resultado_presenca = servico.spreadsheets().values().get(spreadsheetId=ID_PLANILHA_EXEMPLO, range=INTERVALO_CONTROLE_PRESENCA).execute()
                    valores_presenca = resultado_presenca.get("values", [])
                    if valores_presenca:
                        for linha_presenca in valores_presenca:
                            if len(linha_presenca) >= 2 and ra == linha_presenca[0] and timestamp == linha_presenca[1]:
                                print(f"Presença duplicada para RA {ra} no horário {timestamp}.")
                                break
                        else:
                            # Adicione uma nova entrada se o RA e o timestamp ainda não estiverem na planilha de controle de presença
                            valores_presenca.append([ra, timestamp, "1"])
                    else:
                        # Adicione uma nova entrada se a planilha de controle de presença estiver vazia
                        valores_presenca = [[ra, timestamp, "1"]]
                    
                    # Atualize a planilha de controle de presença
                    resultado_atualizacao = servico.spreadsheets().values().update(
                        spreadsheetId=ID_PLANILHA_EXEMPLO,
                        range=INTERVALO_CONTROLE_PRESENCA,
                        valueInputOption="USER_ENTERED",
                        body={"values": valores_presenca}
                    ).execute()
                    print("Controle de presença atualizado.")

                    # Verifique se o RA existe na planilha de quantidade de presença
                    resultado_quantidade_presenca = servico.spreadsheets().values().get(
                        spreadsheetId=ID_PLANILHA_EXEMPLO,
                        range=INTERVALO_QUANTIDADE_PRESENCA
                    ).execute()
                    valores_quantidade_presenca = resultado_quantidade_presenca.get("values", [])
                    ra_encontrado = False
                    for linha_quantidade_presenca in valores_quantidade_presenca:
                        if len(linha_quantidade_presenca) >= 1 and ra == linha_quantidade_presenca[0]:
                            ra_encontrado = True
                            # Incremente a contagem
                            contagem = int(linha_quantidade_presenca[1]) + 1
                            linha_quantidade_presenca[1] = str(contagem)  # Atualize a contagem na lista
                            break
                    if not ra_encontrado:
                        # Adicione uma nova entrada para o RA na planilha de quantidade de presença
                        valores_quantidade_presenca.append([ra, "1"])

                    # Atualize a planilha de quantidade de presença
                    resultado_atualizacao_quantidade_presenca = servico.spreadsheets().values().update(
                        spreadsheetId=ID_PLANILHA_EXEMPLO,
                        range=INTERVALO_QUANTIDADE_PRESENCA,
                        valueInputOption="USER_ENTERED",
                        body={"values": valores_quantidade_presenca}
                    ).execute()
                    print("Quantidade de presença atualizada.")
        # Limpe os dados de presença após a atualização da quantidade de presença
        resultado_limpeza = servico.spreadsheets().values().clear(
            spreadsheetId=ID_PLANILHA_EXEMPLO,
            range=INTERVALO_CONTROLE_PRESENCA
        ).execute()
        print("Dados de presença limpos.")

    except HttpError as erro:
        print(erro)


if __name__ == "__main__":
    main()
 