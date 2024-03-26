"""
Este módulo contém um exemplo de utilização da função enviar_email do módulo func_outlook_python.
Pylint: 0 -> 10
"""

# Importações
import func_outlook_python

#* FORMULÁRIO DE ENVIO ====================================
# Lista de destinários e cópias

# Insira o e-mail do qual você deseja enviar
REMENTENTE = "seu_email@exemplo.com"
# Insira os e-mails dos destinatários
DESTINATARIOS = ["dest_1@exemplo.com", "dest_2@exemplo.com", "dest_3@exemplo.com"]
# Insira os e-mails dos destinatários em cópia (caso não tenha nenhum deixar colchetes vazios [])
CC = ["cc_1@exemplo.com", "cc_2@exemplo.com"]

# Estrutura do e-mail
TITULO = "[nome do relatório]"
CORPO = "Prezados bom dia! \nSegue em anexo [nome do relatório] atualizado até o dia [data]."

# Caminho dos anexos
SCREENSHOT = r'C:\Users\Usuario\Entregas\Relatorio\print.png'
ASSINATURA = r'C:\Users\Usuario\Entregas\Relatorio\assinatura.png'
ANEXO = r'C:\Users\Usuario\Entregas\Relatorio\anexo.xlsx'
# =========================================================

# ENVIO ===================================================
func_outlook_python.enviar_email(
    remetente = REMENTENTE,
    destinatarios = DESTINATARIOS,
    cc = CC,
    titulo = TITULO,
    assunto = CORPO,
    caminho_screenshot = SCREENSHOT,
    caminho_assinatura = ASSINATURA,
    caminho_anexo = ANEXO)
# =========================================================
