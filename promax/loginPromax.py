import requests, re, random

class Promax:
    def __init__(self, usuario, senha, unidade = 0):
        self.usuario = usuario
        self.senha = senha
        self.unidade = "0410001" if unidade == 0  else "3630002"
        self.codunidade = unidade

    def getunicidade(self):
        return self.unidade
            
    def getunidade(self):
        if self.unidade == "0410001": 
            return "Juiz de fora" 
        else: 
            return "Barbacena"
        
    def getCodUnidade(self):
        return self.codunidade
    
    def login(self):
        url = f"http://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?ppopcao=1"

        payload = {
                "SessionID": "",
                "call": "",
                "cmdConfirma": "Confirma",
                "ppopcao": "1",
                "Usuario": self.usuario,
                "Senha": self.senha,
            }
            
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
        response = requests.post(url, data=payload)
        response.raise_for_status()

        html = response.text
        resultado = re.search(r'name=Senha\s+value="([^"]+)"', html)

        if resultado:
            # O group(1) retorna apenas o que foi capturado dentro dos parênteses (...)
            senha_extraida = resultado.group(1)

            SessionID = self.monta_session_id()

            payload = {
                    "SessionID": SessionID,
                    "call": "",
                    "cmdConfirma": "Confirma",
                    "ppopcao": "2",
                    "versaoBase": "12.22",
                    "Usuario": self.usuario,
                    "Senha": senha_extraida,
                    "unidade": f"{self.unidade}",
                    "logUsuario": "",
                    "Animate": "0",
                    "pxReqID": ""
                }
            response = requests.post(url, data=payload)
            response.raise_for_status()

            self.SessionID = SessionID
            self.url_login_promax = f"http://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=030104000000&SubSessionID=x9Kv9I12827173"
            self.logou = True
        else:
            print("Usuário ou senha inválidos!")
            self.logou = False
            return False


    def getSessionID(self):
        return self.SessionID
    
    def get_RUL_Login_Promax(self):
        return self.url_login_promax
    
    def monta_session_id(self):
        CON = list("bcdfghjklmnpqrstvyxzw")
        NUM = list("1234567890")
        VOG = list("aeiouAEIOUCKIJKL")
        def parte():
            r = random.random()
            return CON[int(r*21)] + NUM[int(r*10)] + VOG[int(r*16)]
        valor = int(67891346 * (random.random() + 0.2))
        return parte() + parte() + str(valor)