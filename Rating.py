import xlwings as xw
import datetime as Datas
from pathlib import Path
import shutil

def LoginPromax(unidade: int = 0):
    #PROCESSO DE LOGIN NO PROMAX
    import promax.loginPromax as lpx
    
    sessao = lpx.Promax("Alexandre", "Revenda.44", unidade)
    sessao.login()

    return sessao
    #------------------------------------------------------

async def NivelServico():
    # Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Rating"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax()

    Ativo = True
    #01.05.07.04.02
    if(Ativo):
        OP = "01.05.07.04.02_Taruma_Rating"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\ATENDIMENTO\RATING\01.05.07.04.02\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[1].range("A4").value = Unidade
        wb.sheets[1].range("B4").value = Nome
        wb.sheets[1].range("C4").value = Inicio
        wb.sheets[1].range("E4").value = "01_05_07_04_02"
        wb.sheets[1].range("F4").value = str(Caminho.absolute())
        wb.sheets[1].range("H4").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("J4").value = DataCriacao

        #Status
        wb.sheets[1].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_05_07_04_02
 
        #Status
        wb.sheets[1].range("E2").value = "Baixando arquivo CSV"
        C_01_05_07_04_02 = rpx.sitePromoax_01_05_07_04_02_GERAL(Processo_Logar_Promax)
        await C_01_05_07_04_02.solicitar_csv()
        await C_01_05_07_04_02.Salvar_em(str(Caminho.absolute()))

        wb.sheets[1].range("G4").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("I4").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[1].range("D4").value = Termino 

        #Status
        wb.sheets[1].range("E2").value = ""
#--------------------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Nível de Serviço"
    Ativo = False
    #01.20.01.47
    if(Ativo):
        OP = "01.20.01.47_Taruma_Rating"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\ATENDIMENTO\RATING\01.20.01.47\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[1].range("A6").value = Unidade
        wb.sheets[1].range("B6").value = Nome
        wb.sheets[1].range("C6").value = Inicio
        wb.sheets[1].range("E6").value = "01_20_01_47"
        wb.sheets[1].range("F6").value = str(Caminho.absolute())
        wb.sheets[1].range("H6").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("J6").value = DataCriacao

        #Status
        wb.sheets[1].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_20_01_47
        #Status
        wb.sheets[1].range("E2").value = "Baixando arquivo CSV"
        C_01_20_01_47 = rpx.sitePromoax_01_20_01_47(Processo_Logar_Promax)
        await C_01_20_01_47.solicitar_csv()
        await C_01_20_01_47.Salvar_em(str(Caminho.absolute()))

        wb.sheets[1].range("G6").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("I6").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[1].range("D6").value = Termino 

        #Status
        wb.sheets[1].range("E2").value = ""
#--------------------------------------------------------
 

if __name__ == "__main__":
    import asyncio
    asyncio.run(NivelServico())