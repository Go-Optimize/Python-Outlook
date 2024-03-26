"""
Este módulo contém a função para automatizar o envio de e-mails utilizando Python e Outlook.

Autor: Maria Clara Gomes Loiola
Data: 25/03/2024
Pylint: 4.62 -> 7.69
"""

# Importações
import win32com.client

def enviar_email(remetente, destinatarios, cc, titulo, assunto, caminho_screenshot, caminho_assinatura, caminho_anexo):
    """
    Esta função realiza uma operação específica com os parâmetros fornecidos.

    Parâmetros:
    remetente (string): Remetente específicado pelo usuário.
    destinatarios (list): Lista de destinatários do e-mail a ser enviado.
    cc (list): Lista de cópias do e-mail a ser enviado.
    titulo (string): Texto do título do e-mail a ser enviado.
    assunto (string): Texto do assunto do e-mail a ser enviado.
    caminho_print (f-string): f-string do caminho em que a imagem do print está localizada.
    caminho_assinatura (f-string): f-string do caminho em que a imagem da assinatura está localizada.
    caminho_anexo (f-string):  f-string do caminho em que o anexo do e-mail está localizado.

    Retorno:
    none: Esta função realiza uma ação e não gera retorno.
    """
    try:
        # Criar uma instância do Outlook
        outlook = win32com.client.Dispatch('Outlook.Application')

        # Criar um e-mail
        email = outlook.CreateItem(0)  # 0 representa o tipo de item de e-mail

        # Acessar o perfil desejado pelo endereço de e-mail
        profile = None
        namespace = outlook.GetNamespace("MAPI")
        for acc in namespace.Accounts:
            if acc.SmtpAddress == remetente:
                profile = acc
                break

        if profile is not None:
            email._oleobj_.Invoke(*(64209, 0, 8, 0, profile)) # Configurar o perfil do remetente
        else:
            print("Perfil não encontrado")

        # Configurar os campos do e-mail
        email.To = ';'.join(destinatarios)
        email.CC = ';'.join(cc)

        # Conteúdo do E-Mail
        email.Subject = titulo
        email.Body = assunto

        # Anexar imagem
        screenshot_attachment = email.Attachments.Add(caminho_screenshot)
        screenshot_attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "imagem")

        # Configurar a imagem como incorporada
        email.HTMLBody += '<img src="cid:imagem">'

        # Adicionar linha "Atenciosamente,"
        email.HTMLBody += "<br><br>Atenciosamente,"

        # Anexar imagem da assinatura
        assinatura_attachment = email.Attachments.Add(caminho_assinatura)
        assinatura_attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "assinatura")

        # Configurar a imagem da assinatura como incorporada
        email.HTMLBody += '<br><img src="cid:assinatura">'

        # Anexar arquivo de planilha
        email.Attachments.Add(caminho_anexo)

        # Enviar o e-mail
        email.Send()

    except Exception as e:
        input("Erro ao enviar e-mail. Erro: " + str(e))
