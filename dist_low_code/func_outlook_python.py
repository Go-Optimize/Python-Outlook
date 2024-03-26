# Importações
import win32com.client

def enviar_email(remetente, destinatarios, cc, titulo, assunto, caminho_print, caminho_assinatura, caminho_anexo):
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
        attachment = email.Attachments.Add(caminho_print)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "imagem")
    
        # Configurar a imagem como incorporada
        email.HTMLBody = email.HTMLBody + f'<img src="cid:imagem">'
    
        # Adicionar linha "Atenciosamente,"
        email.HTMLBody += "<br><br>Atenciosamente,"
    
        # Anexar imagem da assinatura
        assinatura_attachment = email.Attachments.Add(caminho_assinatura)
        assinatura_attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "assinatura")
    
        # Configurar a imagem da assinatura como incorporada
        email.HTMLBody += f'<br><img src="cid:assinatura">'
    
        # Anexar arquivo de planilha
        planilha_attachment = email.Attachments.Add(caminho_anexo)
    
        # Enviar o e-mail
        email.Send()
    except Exception as e:
        input("Erro ao envioar e-mail. Erro: " + str(e))
