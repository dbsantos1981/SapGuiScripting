from Sap import *

# este arquivo tem exemplos de uso da classe SapGui
# this file has examples of using the SapGui class

# cria um objeto da classe SapGUi
# creates an object of class SapGUi
obj = SapGui()

# tenta iniciar uma sessão através da classe SapGui
# tries to start a session through the SapGui class
session = obj.connect_to_session()

# Inicia uma transação
# Start a transaction
obj.start_transaction(session, "iw33")

# coleta o texto da barra de título
# collects the title bar text
title_text = obj.get_title_text(session)
print(title_text)

# coleta o texto da barra de status
# collects the status bar text
status_text = obj.get_statusbar_text(session)
print(status_text)

# pressionar botão na tela
# press button on the screen
retorno = obj.press_button(session, "wnd[0]/tbar[0]/btn[3]")
print(retorno)

# seleciona item na tela
# select item on screen
retorno = obj.select(session, "wnd[0]/mbar/menu[2]/menu[2]")
print(retorno)

# envia uma tecla especial para a tela
# sends a special key to the screen
retorno = obj.send_vkey(session, 0)
print(retorno)

# Insere valores em um campo do tipo texto
# Insert values in a text field
retorno = obj.set_text(session, "wnd[0]/usr/ctxtCAUFVD-AUFNR", '12345678')
print(retorno)

# Ativa ou não um checkbox
# Activate or not a checkbox
obj.start_transaction(session, "iw38")
retorno1 = obj.set_checkbox(session, "wnd[0]/usr/chkDY_HIS", True)
retorno2 = obj.set_checkbox(session, "wnd[0]/usr/chkDY_IAR", False)
print(retorno1)
print(retorno2)