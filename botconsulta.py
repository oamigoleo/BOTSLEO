import openpyxl
import webbrowser
import urllib.parse
import pyautogui
from time import sleep
from datetime import datetime
import win32com.client as win32

def formatar_data(data):
    try:
        if isinstance(data, str):
            data_obj = datetime.strptime(data.split()[0], "%Y-%m-%d")
        else:
            data_obj = data
        
        return data_obj.strftime("%d/%m/%Y")
    except ValueError:
        return str(data)

def enviar_whatsapp(telefone, nome, RE, data, hora, tipo_consulta):
    data_texto = formatar_data(data)
    
    mensagem = (f"Olá, {nome}. Essa é uma mensagem automática do ambulatório.\n\n"
                f"Sua consulta está marcada para o dia {data_texto} às {hora}\n"
                f"Tipo de consulta: {tipo_consulta}\n\n"
                f"Te aguardamos lá. Até breve.")
                
    print(f"Mensagem a ser enviada: {mensagem}") 
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone=55{telefone}&text={urllib.parse.quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(12)  
    pyautogui.press('enter') 
    sleep(8)  
    fechar_aba_navegador()  
    
    enviar_email(nome, email, mensagem)

def fechar_aba_navegador():
    pyautogui.hotkey('ctrl', 'w')

def enviar_email(nome, email_destino, mensagem):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = email_destino
    mail.Subject = f"Consulta marcada - {nome}"
    mail.Body = mensagem
    mail.Send()

workbook = openpyxl.load_workbook('consulta.xlsx')
pagina = workbook['Planilha1']

for row in pagina.iter_rows(min_row=2, values_only=True): 
    try:
        RE, nome, data, hora, tipo_consulta, telefone, email = row
        if telefone and nome and RE and data and hora and tipo_consulta and email:
            enviar_whatsapp(telefone, nome, RE, data, hora, tipo_consulta)
        else:
            print("Informações incompletas para enviar mensagem.")
    except Exception as e:
        print(f"Erro ao processar a linha: {e}")

print("Processo de envio concluído!")
