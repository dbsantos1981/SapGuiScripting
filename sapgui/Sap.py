import pythoncom
import win32com.client

class SapGui():

    def __connect_to_session(self, session_number=0):
        # """
        # Conecta a uma sessão SAP ativa (SAP usado nos testes: 7.40)
        # :return: session (objeto do SAP que manipula as telas)
        # :raises: string com mensagem de erro
        # """
        try:
            session = win32com.client.GetObject("SAPGUI").GetScriptingEngine.children(0).children(session_number)
            if session.findById("wnd[0]/titl").text == "SAP":
                return 'SAP Logon aberto, sem login. Favor inserir usuário e senha.'
            else:
                # session é um COMObject, tipo 'win32com.client.CDispatch'
                return session

        except pythoncom.com_error as err:
            if err.args[0] == -2147352567:
                return 'SAP Logon aberto, mas nao conectado ao servidor de Script.'
            elif err.args[0] == -2147221020:
                return 'SAP Logon fechado. Abra-o, entre no servidor e faça o login.'
            else:
                return 'Erro desconhecido. Acionar suporte.'
        
    def __test_id(self, session, id):
        """
        Retorna True caso o id exista na tela
        Returns True if the id exists on the screen
        """
        try:
            session.findById(id)
            return True
        except:
            return False

    def __get_text(self, id):
        """
        Retorna o texto de um id passado como parametro
        Returns the text of an id passed as a parameter
        """
        session = SapGui.__connect_to_session(self)

        if isinstance(session, str):
            return session
        
        else:
            if SapGui.__test_id(self, session, id):
                return session.findbyid(id).text
                
    @staticmethod
    def get_title_text(self):
        """
        Retorna o texto da barra de título da tela do SAP
        Returns the text from the SAP screen title bar
        """

        id = 'wnd[0]/titl' 
        return self.__get_text(self, id)

    @staticmethod
    def get_statusbar_text(self):
        """
        Retorna o texto da barra de status da tela do SAP
        Returns the text from the SAP screen status bar
        """

        id = 'wnd[0]/sbar/pane[0]' 
        return self.__get_text(self, id)

# a = SapGui.get_statusbar_text(SapGui)
# b = SapGui.get_title_text(SapGui)
# print(a)
# print(b)

# Sequencia de instalacao de requisitos
# python -m pip install -U virtualenv
# criar o venv e ativar (virtualenv nome_da_virtualenv) e (.\scripts\activate)
# python -m pip install -U pywin32
#