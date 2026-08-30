import requests, re
import promax.loginPromax
import json

from datetime import date, timedelta

def extrair_caminho_pdf(conteudo_html):
    """
    Localiza e retorna o caminho do PDF dentro do comando window.open
    """
    # Regex explica: 
    # window\.open\(\s*['"] -> Procura por window.open( seguido de aspas
    # (\.\./tmp/rels/.*?\.pdf) -> Captura o caminho que começa com ../tmp/rels/ e termina em .pdf
    padrao = r'"([^"]+\.csv\.inf)"'
    
    resultado = re.search(padrao, conteudo_html)
    
    if resultado:
        # Extrai o link e remove possíveis caracteres nulos (\x00) ou espaços
        cl = (resultado.group(1).replace('\x00', '').strip())
        caminho_limpo = cl[2:len(cl)]
        return caminho_limpo
    
    return False



class sitePromoax_01_05_07_04_02_GERAL:
    """
        1 - Nível de serviço
        2 - Rating
        3 - Gestão MPD
        4 - BEES
        5 - Solicitações
        6 - Volume
        7 - BEES Delivere

        Preferências: Todos 
        Cliente: Rota e AS 
        Clique em: Gerar CSV 

    """
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "01_05_07_04_02"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=0105070402&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            headers = {
                "Content-Type": "application/x-www-form-urlencoded; charset=iso-8859-1",
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "X-Requested-With": "XMLHttpRequest",
                "Referer": f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=0105070402&SubSessionID={self.SubSessionID}",
                "Accept-Language": "pt-br",
                "Accept-Encoding": "gzip, deflate",
                "User-Agent": "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Tablet PC 2.0; Zoom 3.6.0)",
                "Host": "taruma.promaxcloud.com.br",
                "Connection": "close" # Trocamos para close para evitar o RST do firewall ao manter o socket aberto
            }

            # Dados que a função JavaScript PDF() preenche antes do submit
            payload = {
            "SessionID": self.SessionID,
            "SubSessionID": self.SubSessionID,
            "opcao": "1",
            "ppopcao": "0",
            "call": "PH050094",
            "callAux": "",
            "ajaxJson": "S",
            "callRetorno": "PY050094",
            "lnkInc": "1",
            "lnkAlt": "1",
            "lnkExc": "1",
            "lnkCon": "1",
            "lnkEsp": "1",
            "cdClienteIni": "0",
            "cdClienteFim": "9.999.999",
            "cdSetorIni": "0",
            "cdSetorFim": "99.999",
            "cdCorporativoIni": "0",
            "cdCorporativoFim": "9.999.999.999.999",
            "flPreferencias": "1111111000",
            "idRota": "S",
            "idAS": "S"
        }
            
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload)
            if response.status_code == 200:
                Texto = response.text

                #Texto = '{ "downloadCSV" : [ { "dsCaminhoArquivoCSV" :"../tmp/rels/PY050094_25082026_163219.csv.inf" }  ], "mensagemRetorno" : [ { "mensagem" :"CSV gerado com Sucesso!", "opcao" :"N" }  ] }'
                std = json.loads(Texto)
                path = std["downloadCSV"][0]["dsCaminhoArquivoCSV"]

                padrao = r'(.*rels)'
                ispath = re.search(padrao, path)

                if not (ispath == None):
                    url_cmp = path[2:len(path)]
                    await self.baixar_csv(url_cmp)
                    rst = (True, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_01_05_07_04_02:
    """
        1 - Crítica de pedidos

        referências: Ativos, Em Cadastramento, Venda Temporária  
        Cliente: Rota e AS 
        Clique em: Gerar CSV  

    """
        
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "01_05_07_04_02"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=0105070402&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            headers = {
                "Content-Type": "application/x-www-form-urlencoded; charset=iso-8859-1",
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "X-Requested-With": "XMLHttpRequest",
                "Referer": f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=0105070402&SubSessionID={self.SubSessionID}",
                "Accept-Language": "pt-br",
                "Accept-Encoding": "gzip, deflate",
                "User-Agent": "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Tablet PC 2.0; Zoom 3.6.0)",
                "Host": "taruma.promaxcloud.com.br",
                "Connection": "close" # Trocamos para close para evitar o RST do firewall ao manter o socket aberto
            }

            # Dados que a função JavaScript PDF() preenche antes do submit
            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "0",
                "call": "PH050094",
                "callAux": "",
                "ajaxJson": "S",
                "callRetorno": "PY050094",
                "lnkInc": "1",
                "lnkAlt": "1",
                "lnkExc": "1",
                "lnkCon": "1",
                "lnkEsp": "1",
                "cdClienteIni": "0",
                "cdClienteFim": "9.999.999",
                "cdSetorIni": "0",
                "cdSetorFim": "99.999",
                "cdCorporativoIni": "0",
                "cdCorporativoFim": "9.999.999.999.999",
                "idRota": "S",
                "idAS": "S",
                "": "",
                "cdClienteIni": "0",
                "cdClienteFim": "9.999.999",
                "cdSetorIni": "0",
                "cdSetorFim": "99.999",
                "cdCorporativoIni": "0",
                "cdCorporativoFim": "9.999.999.999.999",
                "flPreferencias": "1100101000",
                "idRota": "S",
                "idAS": "S"
            }
            
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload)
            if response.status_code == 200:
                Texto = response.text

                #Texto = '{ "downloadCSV" : [ { "dsCaminhoArquivoCSV" :"../tmp/rels/PY050094_25082026_163219.csv.inf" }  ], "mensagemRetorno" : [ { "mensagem" :"CSV gerado com Sucesso!", "opcao" :"N" }  ] }'
                std = json.loads(Texto)
                path = std["downloadCSV"][0]["dsCaminhoArquivoCSV"]

                padrao = r'(.*rels)'
                ispath = re.search(padrao, path)

                if not (ispath == None):
                    url_cmp = path[2:len(path)]
                    self.baixar_csv(url_cmp)
                    rst = (True, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_01_11:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax, path_savar_csv):
        self.arquivo = "01.11.csv"
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.path_savar_csv = path_savar_csv

        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            with open(self.path_savar_csv, "wb") as arquivo_csv:
                arquivo_csv.write(response.content)

            print(f"Caminho: {self.path_savar_csv}")


        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        try:
            url = "http://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe"
            
            # Dados que a função JavaScript PDF() preenche antes do submit
            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": "w0Lt8K17786294",
                "opcao": "9", 
                "frame": "9", 
                "call": "PW01084C",
                "fecharProcesso": "1",
                # ... outros campos do form1 que estão no seu HTML ...
            }
            
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, data=payload)
            
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    url_cmp = extrair_caminho_pdf(Texto)
                    await self.baixar_csv(url_cmp)
                    rst = (True, False)
                else:
                    print("Erro ao realizar o download desse arquivo: 01_11_csv ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print("Ocooreu algum erro ao realizar o download desse arquivo: 01_12_csv")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_01_20_01_24:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "01_20_01_24"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)
   
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = "http://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe"
            headers = {
                "Accept": "image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*",
                "Referer": f"http://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&SubSessionID={self.SubSessionID}&opcao=0&frame=0&ppopcao=00&call=01200124000&",
                "Accept-Language": "pt-BR",
                "User-Agent": "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Tablet PC 2.0; Zoom 3.6.0)",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "opcaoRel": "2",
                "call": "PW01044R",
                "GeraExcel": "1",
            }
            
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            
            
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "Accept": "image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*",
                        "Referer": f"http://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&SubSessionID={self.SubSessionID}&opcao=0&frame=0&ppopcao=00&call=01200124000&",
                        "Accept-Language": "pt-BR",
                        "User-Agent": "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Tablet PC 2.0; Zoom 3.6.0)",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW01044R",
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "pg1": "1",
                        "pg2": "3",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>",
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (False, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print("Ocooreu algum erro ao realizar o download desse arquivo: 01_11_csv")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_03_01_47_01:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.unicidade = sessaoLogin.getunicidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "03_01_47_01"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=0105070402&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            Primeiro_Dias_Mes = Datas.primeiro_dia_do_mes()
            Ultimo_Dia_Mes = Datas.dia_anterior()

            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW02630R",
                "visaoConsolidadora": "",
                "quebra1Ini9": "",
                "quebra1Fim9": "",
                "quebra2Ini9": "",
                "quebra2Fim9": "",
                "quebra3Ini9": "",
                "quebra3Fim9": "",
                "listaFiliais": f"{self.unicidade}",
                "quebra1": "01",
                "indMotivos": "00",
                "indSubMotivos": "00",
                "indPalmtop": "",
                "indExpurgo": "01",
                "selecionaUnidade": "U",
                "pedidosExcluidos": "S",
                "tipoRoteirizacao": "T",
                "cdAlcadaInicial": "03",
                "cdAlcadaFinal": "09",
                "dataInicial": f"{Primeiro_Dias_Mes}",
                "dataFinal": f"{Ultimo_Dia_Mes}",
                "tipoMarcaInicial": "0",
                "tipoMarcaFinal": "99999999",
                "embalagemInicial": "0",
                "embalagemFinal": "99999999",
                "mercadoriaInicial": "0",
                "mercadoriaFinal": "99999999",
                "idVisaoMultiCdd": "C",
                "idSelecaoMultiCdd": "T",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1",
                "BotOk": "<LABEL><U>O</U>K</LABEL>"
            }
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW02630R",
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "pg1": "1",
                        "pg2": "1",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",     # Substituindo o caractere quebrado (í)
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>",    # Substituindo o caractere quebrado (ó)
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (False, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_03_02_37_NS:
    """
        NÍVEL DE SERVIÇO

        Quebra 1: Operação 
        Quebra 2: Vendedor 
        Quebra 3: Motorista 
        Itens: Sim 
        Data: Primeiro Dia do Mês até o Dia Anterior 

    """
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "03_02_37_NS"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=0105070402&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            Primeiro_Dias_Mes = Datas.primeiro_dia_do_mes()
            Ultimo_Dia_Mes = Datas.dia_anterior()
            
            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW02099R",
                "visaoConsolidadora": "",
                "quebra1Ini8": "",
                "quebra1Fim8": "",
                "quebra2Ini8": "",
                "quebra2Fim8": "",
                "quebra3Ini8": "",
                "quebra3Fim8": "",
                "quebra1": "14",
                "quebra2": "06",
                "quebra3": "25",
                "indPalmtop": "",
                "statusNota": "E",
                "quebraPagina": "S",
                "itens": "S",
                "notas": "NS",
                "visao": "E",
                "quebra1Inicial": "0",
                "quebra1Final": "99999999",
                "quebra2Inicial": "0",
                "quebra2Final": "99999999",
                "quebra3Inicial": "0",
                "quebra3Final": "99999999",
                "dataInicial": f"{Primeiro_Dias_Mes}",
                "dataFinal": f"{Ultimo_Dia_Mes}",
                "valorInicial": "0,00",
                "valorFinal": "999999999,99",
                "embalagemInicial": "0",
                "embalagemFinal": "99999",
                "mercadoriaInicial": "0",
                "mercadoriaFinal": "9999999",
                "idVisaoMultiCdd": "G",
                "idSelecaoMultiCdd": "T",
                "BotStatusNfe": "<LABEL>Selecionar <U>S</U>tatus da NF-e</LABEL>",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1",
                "idStTodos": "S",
                "BotOk": "<LABEL><U>O</U>K</LABEL>"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW02099R", # Código do relatório de notas
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",  # Corrigido caractere especial 'í'
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>", # Corrigido caractere especial 'ó'
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (False, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_03_02_37_RATING:
    """
        RATING
        
        Quebra 1: Motorista 
        Quebra 2: Ajudante 1 
        Quebra 3: Ajudante 2 
        Data: Início do Mês Até o Último Dia do Mês 

    """

    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "03_02_37_RATING"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=0105070402&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            
            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            Primeiro_Dias_Mes = Datas.primeiro_dia_do_mes()
            Ultimo_Dia_Mes = Datas.dia_anterior()

            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW02099R",
                "visaoConsolidadora": "",
                "quebra1Ini8": "",
                "quebra1Fim8": "",
                "quebra2Ini8": "",
                "quebra2Fim8": "",
                "quebra3Ini8": "",
                "quebra3Fim8": "",
                "quebra1": "25",
                "quebra2": "36",
                "quebra3": "37",
                "indPalmtop": "",
                "statusNota": "E",
                "quebraPagina": "S",
                "itens": "N",
                "notas": "NS",
                "visao": "E",
                "quebra1Inicial": "0",
                "quebra1Final": "99999999",
                "quebra2Inicial": "0",
                "quebra2Final": "99999999",
                "quebra3Inicial": "0",
                "quebra3Final": "99999999",
                "dataInicial": f"{Primeiro_Dias_Mes}",
                "dataFinal": f"{Ultimo_Dia_Mes}",
                "valorInicial": "0,00",
                "valorFinal": "999999999,99",
                "idVisaoMultiCdd": "G",
                "idSelecaoMultiCdd": "T",
                "BotStatusNfe": "<LABEL>Selecionar <U>S</U>tatus da NF-e</LABEL>",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1",
                "idStTodos": "S",
                "BotOk": "<LABEL><U>O</U>K</LABEL>"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW02099R", # Código do relatório de notas
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",  # Corrigido caractere especial 'í'
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>", # Corrigido caractere especial 'ó'
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (False, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_03_02_37_MPD:
    """
        GESTÃO MPD

        Quebra 1: Ajudante 1 
        Quebra 2: Ajudante 2 
        Data: Início do Mês Até o Dia Anterior  

    """    
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()

        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "03_02_37_MPD"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=0105070402&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            Primeiro_Dias_Mes = Datas.primeiro_dia_do_mes()
            Ultimo_Dia_Mes = Datas.dia_anterior()

            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW02099R",
                "visaoConsolidadora": "",
                "quebra1Ini8": "",
                "quebra1Fim8": "",
                "quebra2Ini8": "",
                "quebra2Fim8": "",
                "quebra3Ini8": "",
                "quebra3Fim8": "",
                "quebra1": "36",
                "quebra2": "37",
                "quebra3": "00",
                "indPalmtop": "",
                "statusNota": "E",
                "quebraPagina": "S",
                "itens": "N",
                "notas": "NS",
                "visao": "E",
                "quebra1Inicial": "0",
                "quebra1Final": "99999999",
                "quebra2Inicial": "0",
                "quebra2Final": "99999999",
                "dataInicial": f"{Primeiro_Dias_Mes}",
                "dataFinal": f"{Ultimo_Dia_Mes}",
                "valorInicial": "0,00",
                "valorFinal": "999999999,99",
                "idVisaoMultiCdd": "G",
                "idSelecaoMultiCdd": "T",
                "BotStatusNfe": "<LABEL>Selecionar <U>S</U>tatus da NF-e</LABEL>",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1",
                "idStTodos": "S",
                "BotOk": "<LABEL><U>O</U>K</LABEL>"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW02099R", # Código do relatório de notas
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",  # Corrigido caractere especial 'í'
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>", # Corrigido caractere especial 'ó'
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (False, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_03_02_37_AJUD:
    """
        AJUDANTE

        Quebra 1: Operação 
        Quebra 2: Ajudante 1 
        Quebra 3: Ajudante 2 
        Itens: Sim 
        Data: Início do Mês Até o Último Dia do Mês

    """    
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "03_02_37_AJUD"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=0105070402&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado

            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            Primeiro_Dias_Mes = Datas.primeiro_dia_do_mes()
            Ultimo_Dia_Mes = Datas.dia_anterior()

            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW02099R",
                "visaoConsolidadora": "",
                "quebra1Ini8": "",
                "quebra1Fim8": "",
                "quebra2Ini8": "",
                "quebra2Fim8": "",
                "quebra3Ini8": "",
                "quebra3Fim8": "",
                "quebra1": "14",
                "quebra2": "36",
                "quebra3": "37",
                "indPalmtop": "",
                "statusNota": "E",
                "quebraPagina": "S",
                "itens": "S",
                "notas": "NS",
                "visao": "E",
                "quebra1Inicial": "0",
                "quebra1Final": "99999999",
                "quebra2Inicial": "0",
                "quebra2Final": "99999999",
                "quebra3Inicial": "0",
                "quebra3Final": "99999999",
                "dataInicial": f"{Primeiro_Dias_Mes}",
                "dataFinal": f"{Ultimo_Dia_Mes}",
                "valorInicial": "0,00",
                "valorFinal": "999999999,99",
                "embalagemInicial": "0",
                "embalagemFinal": "99999",
                "mercadoriaInicial": "0",
                "mercadoriaFinal": "9999999",
                "idVisaoMultiCdd": "G",
                "idSelecaoMultiCdd": "T",
                "BotStatusNfe": "<LABEL>Selecionar <U>S</U>tatus da NF-e</LABEL>",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1",
                "idStTodos": "S",
                "BotOk": "<LABEL><U>O</U>K</LABEL>"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW02099R", # Código do relatório de notas
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",  # Corrigido caractere especial 'í'
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>", # Corrigido caractere especial 'ó'
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (False, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_03_02_37_VOLUME:
    """
        VOLUME

        Quebra 1: Operação 
        Quebra 2: Vendedor 
        Quebra 3: Motorista 
        Itens: Sim 
        Data: Primeiro Dia do Mês até o Dia Anterior 

    """

    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "03_02_37_VOLUME"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=0105070402&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            Primeiro_Dias_Mes = Datas.primeiro_dia_do_mes()
            Ultimo_Dia_Mes = Datas.dia_anterior()

            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW02099R",
                "visaoConsolidadora": "",
                "quebra1Ini8": "",
                "quebra1Fim8": "",
                "quebra2Ini8": "",
                "quebra2Fim8": "",
                "quebra3Ini8": "",
                "quebra3Fim8": "",
                "quebra1": "14",
                "quebra2": "06",
                "quebra3": "25",
                "indPalmtop": "",
                "statusNota": "E",
                "quebraPagina": "S",
                "itens": "S",
                "notas": "NS",
                "visao": "E",
                "quebra1Inicial": "0",
                "quebra1Final": "99999999",
                "quebra2Inicial": "0",
                "quebra2Final": "99999999",
                "quebra3Inicial": "0",
                "quebra3Final": "99999999",
                "dataInicial": f"{Primeiro_Dias_Mes}",
                "dataFinal": f"{Ultimo_Dia_Mes}",
                "valorInicial": "0,00",
                "valorFinal": "999999999,99",
                "embalagemInicial": "0",
                "embalagemFinal": "99999",
                "mercadoriaInicial": "0",
                "mercadoriaFinal": "9999999",
                "idVisaoMultiCdd": "G",
                "idSelecaoMultiCdd": "T",
                "BotStatusNfe": "<LABEL>Selecionar <U>S</U>tatus da NF-e</LABEL>",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1",
                "idStTodos": "S",
                "BotOk": "<LABEL><U>O</U>K</LABEL>"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW02099R", # Código do relatório de notas
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",  # Corrigido caractere especial 'í'
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>", # Corrigido caractere especial 'ó'
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (False, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_01_20_01_47:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "01_20_01_47"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=0&call=01200147000&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW01076R",
                "opcaoRel": "2",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW01076R",
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "pg1": "1",
                        "pg2": "5",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>",
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (False, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_03_11_20:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "03_11_20"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=0&call=031120000000&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado

            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            Primeiro_Dias_Mes = Datas.primeiro_dia_do_mes()
            Ultimo_Dia_Mes = Datas.dia_anterior()

            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW02227R",
                "rtInicialh": "00000000000000000000",
                "rtFinalh": "99999999999999999999",
                "opcaoRel": "1",
                "dataInicial": f"{Primeiro_Dias_Mes}",
                "dataFinal": f"{Ultimo_Dia_Mes}",
                "veiculoInicial": "0",
                "veiculoFinal": "999",
                "cdCampoInicial": "0",
                "cdCampoFinal": "999999",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW02227R",
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>",
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (True, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_03_11_49_02:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "03_11_49_02"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=03114902000&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            headers = {
                "Content-Type": "application/x-www-form-urlencoded; charset=iso-8859-1",
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "X-Requested-With": "XMLHttpRequest",
                "Referer": f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=03114902000&SubSessionID={self.SubSessionID}",
                "Accept-Language": "pt-br",
                "Accept-Encoding": "gzip, deflate",
                "User-Agent": "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Tablet PC 2.0; Zoom 3.6.0)",
                "Host": "taruma.promaxcloud.com.br",
                "Connection": "close" # Trocamos para close para evitar o RST do firewall ao manter o socket aberto
            }

            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            Primeiro_Dias_Mes = Datas.primeiro_dia_do_mes()
            Ultimo_Dia_Mes = Datas.dia_anterior()

            # Dados que a função JavaScript PDF() preenche antes do submit
            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "6",
                "ppopcao": "0",
                "call": "PH050158",
                "callAux": "",
                "ajaxJson": "S", # Indica ao servidor que o retorno deve ser JSON
                "callRetorno": "PY050158",
                "lnkInc": "1",
                "lnkAlt": "1",
                "lnkExc": "1",
                "lnkCon": "1",
                "lnkEsp": "1",
                "opClassificacao": "0",
                "checkRota": "1",
                "checkAs": "1",
                "idRoteirizado": "0",
                "mapaInicial": "0",
                "mapaFinal": "999999",
                "dataInicial": f"{Primeiro_Dias_Mes}",
                "dataFinal": f"{Ultimo_Dia_Mes}",
                "roadInicial": "0",
                "roadFinal": "99",
                "cdTransportadoraInicial": "0",
                "cdTransportadoraFinal": "999999",
                "opArmazem": "0"
            }
            
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload)
            if response.status_code == 200:
                Texto = response.text

                #Texto = '{ "downloadCSV" : [ { "dsCaminhoArquivoCSV" :"../tmp/rels/PY050094_25082026_163219.csv.inf" }  ], "mensagemRetorno" : [ { "mensagem" :"CSV gerado com Sucesso!", "opcao" :"N" }  ] }'
                std = json.loads(Texto)
                path = std["downloadCSV"][0]["dsCaminhoArquivoCSV"]

                padrao = r'(.*rels)'
                ispath = re.search(padrao, path)

                if not (ispath == None):
                    url_cmp = path[2:len(path)]
                    await self.baixar_csv(url_cmp)
                    rst = (True, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_03_02_37_MPD:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "03_02_37_MPD"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=7&call=030237000000&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            Primeiro_Dias_Mes = Datas.primeiro_dia_do_mes()
            Ultimo_Dia_Mes = Datas.dia_anterior()

            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW02099R",
                "visaoConsolidadora": "",
                "quebra1Ini8": "",
                "quebra1Fim8": "",
                "quebra2Ini8": "",
                "quebra2Fim8": "",
                "quebra3Ini8": "",
                "quebra3Fim8": "",
                "quebra1": "36",
                "quebra2": "37",
                "quebra3": "00",
                "indPalmtop": "",
                "statusNota": "E",
                "quebraPagina": "S",
                "itens": "N",
                "notas": "NS",
                "visao": "E",
                "quebra1Inicial": "0",
                "quebra1Final": "99999999",
                "quebra2Inicial": "0",
                "quebra2Final": "99999999",
                "dataInicial": f"{Primeiro_Dias_Mes}",
                "dataFinal": f"{Ultimo_Dia_Mes}",
                "valorInicial": "0,00",
                "valorFinal": "999999999,99",
                "idVisaoMultiCdd": "G",
                "idSelecaoMultiCdd": "T",
                "BotStatusNfe": "<LABEL>Selecionar <U>S</U>tatus da NF-e</LABEL>",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1",
                "idStTodos": "S",
                "BotOk": "<LABEL><U>O</U>K</LABEL>"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW02099R",
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>",
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (False, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_03_11_40:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "03_11_40"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=0&call=031140000000&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            Primeiro_Dias_Mes = Datas.primeiro_dia_do_mes()
            Ultimo_Dia_Mes = Datas.dia_anterior()

            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW02478R",
                "cdOpcaoMpd": "T",
                "dtInicial": f"{Primeiro_Dias_Mes}",
                "dtFinal": f"{Ultimo_Dia_Mes}",
                "cdVeiculoInicial": "0",
                "cdVeiculoFinal": "999",
                "nrMapaInicial": "0",
                "nrMapaFinal": "999999",
                "cdMotoristaInicial": "0",
                "cdMotoristaFinal": "99999",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW02478R",
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",     # Corrigido: 'í'
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>",    # Corrigido: 'ó'
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (False, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_01_20_01_48:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "01_20_01_48"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=0&call=01200148000&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PP01015R",
                "ordenacao": "2",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PP01015R",
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",     # Corrigido caractere especial 'í'
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>",    # Corrigido caractere especial 'ó'
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (False, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_01_09:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax, path_savar_csv):
        self.arquivo = "01.09.csv"
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.path_savar_csv = path_savar_csv

        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            print("Salvando 11.09...")

            with open(self.path_savar_csv, "wb") as arquivo_csv:
                arquivo_csv.write(response.content)

            print(f"Caminho: {self.path_savar_csv}")


        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        try:
            url = "http://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe"
            
            # Dados que a função JavaScript PDF() preenche antes do submit
            payload = {
                "frame": "9",
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "9",
                "ppopcao": "00",
                "call": "PW01082C",
                "cdProduto": "",
                "nrDigito": "",
                "nmProduto": "",
                "nmAbreviado": "",
                "cdSubgrupo": "",
                "cdFamiliaMaterial": "",
                "nmFamiliaMaterial": "",
                "cdTipoMarca": "",
                "cdGrupo": "",
                "capacidadeGarrafeira": "",
                "cdUnidadeVenda": "",
                "sgUnidadeVenda": "",
                "cdUnidadeCompra": "",
                "nrSeparadorNf": "",
                "nrTabelaIcms": "",
                "nrClassIbge": "",
                "cdIpi": "",
                "tributacaoIcms": "",
                "origemMercadoria": "",
                "empresa": "",
                "nrFatorConversao": "",
                "qtCompraUnv": "",
                "tpRoadshow": "",
                "cxPalletRoadShow": "",
                "ftConversaoRoadshow": "",
                "nrCaixasPalletEspecial": "",
                "pesoLiquido": "",
                "pesoBruto": "",
                "ordemCarga": "",
                "cdCusto": "",
                "nmCusto": "",
                "dtVigenciaInicial": "",
                "dtVigenciaFinal": "",
                "idCstCbs": "",
                "idClasTributariaCbs": ""
            }
            
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, data=payload)
            
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    url_cmp = extrair_caminho_pdf(Texto)
                    await self.baixar_csv(url_cmp)
                    rst = (True, False)
                else:
                    print("Erro ao realizar o download desse arquivo: 01_09_csv ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print("Ocooreu algum erro ao realizar o download desse arquivo: 01_09_csv")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_01_12:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax, path_savar_csv):
        self.arquivo = "01.12.csv"
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.path_savar_csv = path_savar_csv

        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            print("Salvando 11.12...")

            with open(self.path_savar_csv, "wb") as arquivo_csv:
                arquivo_csv.write(response.content)

            print(f"Caminho: {self.path_savar_csv}")


        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        try:
            url = "http://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe"
            
            # Dados que a função JavaScript PDF() preenche antes do submit
            payload = {
                "frame": "9",
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "9",
                "ppopcao": "00",
                "call": "PW01085C",
                "dtExclusaoOri": "",
                "dtValidadeOri": "",
                "cdMaterial": "",
                "nmMaterial": "",
                "nmAbreviado": "",
                "cdFamiliaMaterial": "",
                "nmFamiliaMaterial": "",
                "cdTipoMarca": "",
                "cdGrupo": "",
                "cdSubgrupo": "",
                "unidadeVenda": "",
                "cdUnidadeCompra": "",
                "dtValidade": "",
                "dtExclusao": "",
                "nrSeparadorNf": "",
                "nrTabelaIcms": "",
                "nrClassIbge": "",
                "cdIpi": "",
                "tributacaoIcms": "",
                "origemMercadoria": "",
                "palmtop": "",
                "empresa": "",
                "cdVasGar": "",
                "dsVasGar": "",
                "tpRoadshow": "",
                "cxPalletRoadShow": "",
                "ftConversaoRoadshow": "",
                "nrFatorConversao": "",
                "nrCaixasPalletEspecial": "",
                "qtCompraUnv": "",
                "ordemCarga": "",
                "pesoLiquido": "",
                "pesoBruto": "",
                "cdCusto": "",
                "nmCusto": "",
                "dtVigenciaInicial": "",
                "dtVigenciaFinal": "",
                "nrSequenciaCstCbs": "",
                "nrSequenciaTriCbs": ""
            }
            
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, data=payload)
            
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    url_cmp = extrair_caminho_pdf(Texto)
                    await self.baixar_csv(url_cmp)
                    rst = (True, False)
                else:
                    print("Erro ao realizar o download desse arquivo: 01_09_csv ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print("Ocooreu algum erro ao realizar o download desse arquivo: 01_12_csv")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_02_05_02:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()

        self.relatorio = "02_05_02"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=0&call=020502000000&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            HJ = Datas.data_hoje()
            
            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW03064R",
                "opcaoRel": "1",
                "idListarProdutos": "S",
                "tpData": "E",
                "periodoInicial": f"{HJ}",
                "periodoFinal": f"{HJ}",
                "cdMercadoriaInicial": "0",
                "cdMercadoriaFinal": "9999999",
                "cdArmazemInicial": "0",
                "cdArmazemFinal": "99",
                "cdDepositoInicial": "1",
                "cdDepositoFinal": "1",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW03064R",
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",     # Corrigido caractere especial 'í'
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>",    # Corrigido caractere especial 'ó'
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (False, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_03_01_11:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "03_01_11"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=0&call=030111000000&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW02041R",
                "opcaoRel": "6",
                "grupoPerfilVendas": "",
                "indPalmtop": "",
                "campo1Inicial": "0",
                "campo1Final": "9999",
                "campo2Inicial": "0",
                "campo2Final": "999999",
                "listaPedidos": "S",
                "resumo": "P",
                "ttv": "P",
                "idPedAgenHoje": "S",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW02041R",
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",     # Corrigido caractere especial 'í'
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>",    # Corrigido caractere especial 'ó'
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (True, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_03_01_36_04:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "03_01_36_04"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=0&call=03013604000&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            HJ = Datas.data_hoje()            
            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW02348R",
                "tpOrdenacao1": "1",
                "tpOrdenacao2": "4",
                "tpOrdenacao3": "6",
                "tpOrigem": "",
                "tpCritica": "00",
                "idNormal": "S",
                "idSistema": "S",
                "idBloqueado": "S",
                "idComandado": "S",
                "idRejeitado": "S",
                "idTransfer": "S",
                "idTodos": "S",
                "dtInicial": f"{HJ}",
                "dtFinal": f"{HJ}",
                "cdAlcadaInicial": "01",
                "cdAlcadaFinal": "90",
                "cdCampoInicial1": "0",
                "cdCampoFinal1": "999999",
                "cdCampoInicial2": "0",
                "cdCampoFinal2": "999999",
                "cdCampoInicial3": "0",
                "cdCampoFinal3": "999999",
                "cdUsuario": "",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW02348R",
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",     # Corrigido caractere especial 'í'
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>",    # Corrigido caractere especial 'ó'
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (True, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_03_02_24:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "03_02_24"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    async def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=0&call=030224000000&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado

            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            Primeiro_Dias_Mes = Datas.tres_meses_atras()
            Ultimo_Dia_Mes = Datas.primeiro_dia_do_mes()

            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW02084R",
                "visaoConsolidadora": "",
                "opcaoRel": "04",
                "setoresInativos": "S",
                "converte": "H",
                "resumoVisao": "N",
                "selecionouROTA": "S",
                "selecionouAS": "S",
                "responsabProcesso": "S",
                "resumo": "S",
                "idVisaoMultiCdd": "C",
                "idSelecaoMultiCdd": "T",
                "dataInicial": f"{Primeiro_Dias_Mes}",
                "dataFinal": f"{Ultimo_Dia_Mes}",
                "campoInicial": "0",
                "campoFinal": "999999",
                "nrRoadInicial": "0",
                "nrRoadFinal": "99",
                "classeInicial": "",
                "classeFinal": "z",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW02084R",
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",     # Corrigido caractere especial 'í'
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>",    # Corrigido caractere especial 'ó'
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                        }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (True, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_12_06_01:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "12_06_01"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=0&call=120601000000&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado
            
            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW04105R",
                "opcaoRel": "05",
                "iniCliente": "0",
                "fimCliente": "999999",
                "iniVencimento": "00/00/0000",
                "fimVencimento": "99/99/9999",
                "iniEmissao": "00/00/0000",
                "fimEmissao": "99/99/9999",
                "iniLiberacao": "00/00/0000",
                "fimLiberacao": "99/99/9999",
                "iniVendedor": "0",
                "fimVendedor": "99999",
                "iniPortador": "0",
                "fimPortador": "999",
                "iniEspecie": "2",
                "fimEspecie": "5",
                "iniVinculo": "0",
                "fimVinculo": "99",
                "iniSegmento": "0",
                "fimSegmento": "99",
                "iniValor": "0,00",
                "fimValor": "99999999,99",
                "iniCGC": "",
                "fimCGC": "ZZZZZZZZZZZZ99",
                "iniArea": "0",
                "fimArea": "99999",
                "titulo": "T",
                "idNotasTitAtu": "S",
                "tituloPdd": "P",
                "idNotasTitNaoAtu": "S",
                "idTituloRefugo": "S",
                "idData": "F",
                "idNome": "F",
                "cdVisao": "",
                "tpConsolidacao": "1",
                "idVisaoMultiCdd": "C",
                "idSelecaoMultiCdd": "T",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW04105R",
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",     # Corrigido caractere especial 'í'
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>",    # Corrigido caractere especial 'ó'
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (True, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)

class sitePromoax_05_10:
    def __init__(self, sessaoLogin:promax.loginPromax.Promax):
        self.url = sessaoLogin.get_RUL_Login_Promax()
        self.unidade = sessaoLogin.getunidade()
        self.codunidade = sessaoLogin.getCodUnidade()
        self.relatorio = "05_10"
        self.Separar_Sessao()
        

    def Separar_Sessao(self):
            from urllib.parse import urlparse, parse_qs

            """
            Extrai SessionID e SubSessionID de uma URL e retorna um dicionário.
            """
            # 1. Analisa a URL para separar o domínio dos parâmetros (query)
            parsed_url = urlparse(self.url)

            # 2. Converte a string de parâmetros em um dicionário
            # O parse_qs retorna listas como valores (ex: {'SessionID': ['valor']})
            params = parse_qs(parsed_url.query)

            self.SessionID = params.get("SessionID", [None])[0]
            self.SubSessionID = params.get("SubSessionID", [None])[0]

    async def Salvar_em(self, path):
        
        print(f"Salvando {__class__}...")
        
        with open(path, "wb") as arquivo_csv:
            arquivo_csv.write(self.conteudo)
        
        print(f"Salvo em: {path}")

    def baixar_csv(self, url):
        try:
            url_csv = f"http://taruma.promaxcloud.com.br{url}"
            response = requests.get(url_csv)
            response.raise_for_status()

            self.conteudo = response.content
            

        except Exception as e:
            print(e)
            print(__name__)

        
    
    async def solicitar_csv(self):
        print(f"Solicitando relatório: {self.relatorio} - Unidade: {self.unidade}")
        try:
            url = f"https://taruma.promaxcloud.com.br/pw/cgi-bin/PP00100.exe?SessionID={self.SessionID}&ppopcao=0&call=0510000000000&subSessionID={self.SubSessionID}"
                # 2. Cabeçalhos (Headers) idênticos ao do tráfego capturado

            import promax.bibliotecas.data_prx as dtprx
            Datas = dtprx.Datas()
            MESANO = Datas.mes_ano_atual()

            payload = {
                "SessionID": self.SessionID,
                "SubSessionID": self.SubSessionID,
                "opcao": "1",
                "ppopcao": "00",
                "call": "PW02006R",
                "dataPrjMarcasHabilit": "12/08/2004",
                "anoMesPrjMarcasHabilit": "200408",
                "opcaoRel": "1",
                "perfilVendas": "000",
                "grupoPerfilVendas": "00",
                "mesAno": f"{MESANO}",
                "totalizaTipo": "S",
                "converteHectolitros": "S",
                "codigoInicial": "00000",
                "codigoFinal": "99999",
                "tipoMarcaInicial": "00",
                "tipoMarcaFinal": "99",
                "diasUteis": "22",
                "diasAcumulados": "22",
                "BotVisualizar": "<LABEL><U>V</U>isualizar</LABEL>",
                "GeraExcel": "1"
            }

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Connection": "keep-alive"
            }
            # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
            response = requests.post(url, headers=headers, data=payload, stream=True)
            if response.status_code == 200:
                Texto = response.text

                padrao = r'(.*login)'
                isLogin = re.search(padrao, Texto)
                if isLogin == None:
                    
                    # 2. Cabeçalhos mantendo o "close" para evitar o erro 10054 do firewall
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "Connection": "keep-alive"
                    }
        
                    # 3. O formulário exato de 843 bytes capturado (com os acentos corrigidos)
                    payload = {
                        "SessionID": self.SessionID,
                        "SubSessionID": self.SubSessionID,
                        "opcao": "88",
                        "ppopcao": "00",
                        "call": "PW02006R",
                        "paginaatual": "000001",
                        "opcaorelat": "3",
                        "sizeFontx": "",
                        "backRelatorio": "0",
                        "filaimpressao": "066",
                        "imp": '<IMG src="../icons/print.gif">',
                        "configurar": "<LABEL>I<U>m</U>primir</LABEL>",
                        "salvar": "<LABEL><U>S</U>alvar<LABEL></LABEL></LABEL>",
                        "GerExecl": "<LABEL><U>C</U>SV</LABEL>",
                        "GerPDF": "<LABEL>P<U>D</U>F</LABEL>",
                        "selimp": "<LABEL><U>V</U>oltar</LABEL>",
                        "sizeFont": "8",
                        "pri": "<LABEL>I<U>n</U>ício</LABEL>",     # Corrigido caractere especial 'í'
                        "ant": "<LABEL>An<U>t</U>erior</LABEL>",
                        "pro": "<LABEL><U>P</U>róxima</LABEL>",    # Corrigido caractere especial 'ó'
                        "ult": "<LABEL><U>F</U>inal</LABEL>",
                        "irpara": "",
                        "btirpara": "<LABEL>Ir par<U>a</U></LABEL>",
                        "fecharProcesso": "1"
                    }
                    
                    # Ao fazer este POST, o servidor processará o csv no diretório /tmp/rels/
                    response2 = requests.post(url, headers=headers, data=payload)
                    padrao = r'(.*login)'
                    isLogin = re.search(padrao, Texto)
                    if isLogin == None:
                        Texto = response2.text

                        url_cmp = extrair_caminho_pdf(Texto)
                        await self.baixar_csv(url_cmp)
                        rst = (True, False)
                else:
                    print(f"Erro ao realizar o download desse arquivo: {self.relatorio} ou sessão expirada!")
                    rst = (False, False)
            else:
                rst = (False, False)


            if rst[0]:
                pass
            else:
                print(f"Ocorreu algum erro ao realizar o download desse arquivo: {self.relatorio}")
            
            return rst
        
        except Exception as e:
            print(e)
            print(__name__)
