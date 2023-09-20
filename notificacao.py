import time
import win32com.client
from pushbullet import Pushbullet

# Substitua 'seu_api_key' pelo seu token de API do Pushbullet
pb = Pushbullet('seu_api_key')

def check_outlook_inbox():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 representa a pasta de entrada (caixa de entrada)

    while True:
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        latest_email = messages.GetFirst()

        # Verifica se o e-mail mais recente ainda não foi processado
        if latest_email.UnRead:
            # Obtém o conteúdo do e-mail
            email_body = latest_email.Body
                
            # Obtém o remetente do e-mail
            email_sender = latest_email.SenderName

            # Envia uma notificação para o seu celular via Pushbullet com o conteúdo do e-mail e o remetente
            push_title = f"Novo E-mail de {email_sender}"
            push_body = f"Conteúdo do E-mail:\n{email_body}"
            push = pb.push_note(push_title, push_body)

            # Marca o e-mail como lido
            latest_email.UnRead = False
            latest_email.Save()

        # Aguarda 30 segundos antes de verificar novamente
        time.sleep(30)

if __name__ == "__main__":
    check_outlook_inbox()
