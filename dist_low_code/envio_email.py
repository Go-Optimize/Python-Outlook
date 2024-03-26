# Importações
import func_outlook_python

#* FORMULÁRIO DE ENVIO ====================================
# Lista de destinários e cópias
remetente = "seu_email@exemplo.com" # Insira o e-mail do qual você deseja enviar
destinatarios = ["dest_1@exemplo.com", "dest_2@exemplo.com", "dest_3@exemplo.com"] # Insira os e-mails dos destinatários
cc = ["cc_1@exemplo.com", "cc_2@exemplo.com"] # Insira os e-mails dos destinatários em cópia (caso não tenha nenhum deixar colchetes vazios [])

# Estrutura do e-mail
titulo = "[nome do relatório]"
corpo = "Prezados bom dia! \nSegue em anexo [nome do relatório] atualizado até o dia [data]. \nQualquer dúvida estou à disposição."

# Caminho dos anexos
print = r'C:\Users\Usuario\Entregas\Relatorio\print.png'
assinatura = r'C:\Users\Usuario\Entregas\Relatorio\assinatura.png'
anexo = r'C:\Users\Usuario\Entregas\Relatorio\anexo.xlsx'
# =========================================================

# ENVIO ===================================================
func_outlook_python.enviar_email(
    remetente = remetente,
    destinatarios = destinatarios, 
    cc = cc,
    titulo = titulo,
    assunto = corpo,
    caminho_print = print,
    caminho_assinatura = assinatura,
    caminho_anexo = anexo)
# =========================================================