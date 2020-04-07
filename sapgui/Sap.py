import pythoncom
import win32com.client
import sys

class SapGui():

    """
    Para essa classe funcionar é necessária a instalação do pacote pywin32
    For this class to work it is necessary to install the pywin32 package
    """

    def connect_to_session(self, session_number=0):
        """
        Conecta a uma sessão SAP ativa (SAP usado nos testes: 7.40)
        :return: session (objeto do SAP que manipula as telas)
        :raises: string com mensagem de erro
        
        Connects to an active SAP session (SAP used in tests: 7.40)
        : return: session (SAP object that handles the screens)
        : raises: string with error message
        """
        try:
            session = win32com.client.GetObject("SAPGUI").GetScriptingEngine.children(0).children(session_number)
            if session.findById("wnd[0]/titl").text == "SAP":
                if self.__test_id(self, session, 'wnd[0]/usr/btnSTARTBUTTON'):
                    self.press_button(self, session, 'wnd[0]/usr/btnSTARTBUTTON')
                    return session
                else:
                    return 'SAP Logon aberto, sem login. Favor inserir usuário e senha.'
            else:
                # session é um COMObject, tipo 'win32com.client.CDispatch'
                # session is a COMObject, type 'win32com.client.CDispatch'
                return session

        except pythoncom.com_error as err:
            if err.args[0] == -2147352567:
                return 'Erro: ' + str(err.args[0]) + ' = SAP Logon aberto, mas não conectado a um servidor de Script.'
            elif err.args[0] == -2147221020:
                return 'Erro: ' + str(err.args[0]) + ' = SAP Logon fechado. Abra-o, entre no servidor e faça o login.'
            elif err.args[0] == -2147221014:
                return 'Erro: ' + str(err.args[0]) + ' = SAP Logon fechado. Abra-o, entre no servidor e faça o login.'
            else:
                return 'Erro: ' +  str(err.args[0]) + ' - Favor informar o número do erro.'
        
        except:
            return sys.exc_info()[0]
        
        finally:
            session = None

    def press_button(self, session, button_id):
        """
        Retorna True se conseguiu pressionar o botão
        Returns True if successfully pressed the button
        """
        if self.__test_id(session, button_id) == True:
            session.findById(button_id).press()
            return True
        else:
            return False
    
    def select(self, session, id):
        """
        Retorna True se conseguiu selecionar o item
        Returns True if the item was successfully selected
        """
        if self.__test_id(session, id) == True:
            session.findById(id).select()
            return True
        else:
            return False
    
    def __test_id(self, session, id):
        """
        Retorna True caso o id exista na tela
        Returns True if the id exists on the screen
        """
        if isinstance(session, str):
            return session
        else:
            try:
                session.findById(id)
                return True
            except:
                return False

    def __get_text(self, session, id):
        """
        Retorna o texto de um id passado como parametro
        Returns the text of an id passed as a parameter
        """
        if isinstance(session, str):
            return session
        else:
            if SapGui.__test_id(self, session, id):
                return session.findbyid(id).text

    def send_vkey(self, session, vkey):
        """
        Retorna True caso consiga enviar a tecla
        Returns True if it was able to send the key
        """
        if isinstance(session, str):
            return session
        else:
            try:
                session.findbyid('wnd[0]').sendvkey(vkey)
                return True
            except:
                return False
        
        # como pegar as teclas ativas em uma sessão
        # for x in range(256):
        #     print('O cod. {} equivale a {}'.format(x, session.GetVKeyDescription(x)))

    def start_transaction(self, session, transaction, hide=False):
        """
        Inicia a transação SAP recebida como parametro
        Starts the SAP transaction received as a parameter
        
        Se 'hide' for True, ele minimiza a janela do SAP
        If 'hide' is True, it minimizes the SAP window

        : return: True se deu certo, ou 'mensagem' caso tenha dado erro
        : return: True if it worked, or 'message' if it went wrong
        """
        #session = self.__connect_to_session()

        if isinstance(session, str):
            return session
        else:
            session.starttransaction(transaction)
            if self.get_statusbar_text(session) == '':
                if hide == True:
                    session.findById("wnd[0]").Iconify()
                return True
            else:
                return self.get_statusbar_text(session)

    def get_title_text(self, session):
        """
        Retorna o texto do título da tela do SAP
        Returns the title text of the SAP screen
        """
        id = 'wnd[0]/titl'
        return self.__get_text(session, id)

    def get_statusbar_text(self, session):
        """
        Retorna o texto da barra de status da tela do SAP
        Returns the text from the SAP screen status bar
        """

        id = 'wnd[0]/sbar/pane[0]' 
        return self.__get_text(session, id)

    def set_text(self, session, id, text):
        """
        Insere valroes em um campo do tipo texto
        Insert values in a text field
        """
        if self.__test_id(session, id) == True:
            session.findById(id).text = text
            return True
        else:
            return False

    def set_checkbox(self, session, id, check):
        """
        Ativa ou não um checkbox
        Activate or not a checkbox
        """
        if self.__test_id(session, id) == True:
            session.findById(id).setFocus()
            if check == True:
                session.findById(id).selected = 1
            elif check == False:
                session.findById(id).selected = 0
            return True
        else:
            return False   