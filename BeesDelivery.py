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

async def BeesDelivery():
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Bees Delivery"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax()

#-------------------------------------------
    # Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "BeesDelivery"
    Ativo = True
    #01.05.07.04.02
    if(Ativo):
        OP = "01.05.07.04.02_BeesDelivery"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\ATENDIMENTO\BEES DELIVERY\01.05.07.04.02\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[7].range("A4").value = Unidade
        wb.sheets[7].range("B4").value = Nome
        wb.sheets[7].range("C4").value = Inicio
        wb.sheets[7].range("E4").value = "01_05_07_04_02"
        wb.sheets[7].range("F4").value = str(Caminho.absolute())
        wb.sheets[7].range("H4").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("J4").value = DataCriacao

        #Status
        wb.sheets[7].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_05_07_04_02
        #Status
        wb.sheets[7].range("E2").value = "Baixando arquivo CSV"
        C_01_05_07_04_02 = rpx.sitePromoax_01_05_07_04_02_GERAL(Processo_Logar_Promax)
        await C_01_05_07_04_02.solicitar_csv()
        await C_01_05_07_04_02.Salvar_em(str(Caminho.absolute()))

        wb.sheets[7].range("G4").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("I4").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[7].range("D4").value = Termino 

        #Status
        wb.sheets[7].range("E2").value = ""
#-------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "beesDelivery"
    Ativo = False
    #01.20.01.47
    if(Ativo):
        OP = "01.20.01.47_Taruma_beesDelivery"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\ATENDIMENTO\BEES DELIVERY\01.20.01.47\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[7].range("A5").value = Unidade
        wb.sheets[7].range("B5").value = Nome
        wb.sheets[7].range("C5").value = Inicio
        wb.sheets[7].range("E5").value = "01_20_01_47"
        wb.sheets[7].range("F5").value = str(Caminho.absolute())
        wb.sheets[7].range("H5").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("J5").value = DataCriacao

        #Status
        wb.sheets[7].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_20_01_47
        #Status
        wb.sheets[7].range("E2").value = "Baixando arquivo CSV"
        C_01_20_01_47 = rpx.sitePromoax_01_20_01_47(Processo_Logar_Promax)
        await C_01_20_01_47.solicitar_csv()
        await C_01_20_01_47.Salvar_em(str(Caminho.absolute()))

        wb.sheets[7].range("G5").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("I5").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[7].range("D5").value = Termino 

        #Status
        wb.sheets[7].range("E2").value = ""
#--------------------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "BeesDelivery"
    Ativo = False
    #03.11.20
    if(Ativo):
        OP = "03.11.20_Taruma_BeesDelivery"
        Inicio = Datas.datetime.now()
        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\ATENDIMENTO\BEES DELIVERY\03.11.20\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[7].range("A6").value = Unidade
        wb.sheets[7].range("B6").value = Nome
        wb.sheets[7].range("C6").value = Inicio
        wb.sheets[7].range("E6").value = "03_11_20"
        wb.sheets[7].range("F6").value = str(Caminho.absolute())
        wb.sheets[7].range("H6").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("J6").value = DataCriacao

        #Status
        wb.sheets[7].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 03_11_20
        #Status
        wb.sheets[7].range("E2").value = "Baixando arquivo CSV"
        C_03_11_20 = rpx.sitePromoax_03_11_20(Processo_Logar_Promax)
        await C_03_11_20.solicitar_csv()
        await C_03_11_20.Salvar_em(str(Caminho.absolute()))

        wb.sheets[7].range("G6").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("I6").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[7].range("D6").value = Termino 

        #Status
        wb.sheets[7].range("E2").value = ""
#--------------------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "BeesDelivery"
    Ativo = False
    #03.02.24
    if(Ativo):
        OP = "03.02.24_Taruma_BeesDelivery"
        Inicio = Datas.datetime.now()
        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\ATENDIMENTO\BEES DELIVERY\03.02.24\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[7].range("A7").value = Unidade
        wb.sheets[7].range("B7").value = Nome
        wb.sheets[7].range("C7").value = Inicio
        wb.sheets[7].range("E7").value = "03_02_24"
        wb.sheets[7].range("F7").value = str(Caminho.absolute())
        wb.sheets[7].range("H7").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("J7").value = DataCriacao

        #Status
        wb.sheets[7].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 03_02_24
        #Status
        wb.sheets[7].range("E2").value = "Baixando arquivo CSV"
        C_03_02_24 = rpx.sitePromoax_03_02_24(Processo_Logar_Promax)
        await C_03_02_24.solicitar_csv()
        await C_03_02_24.Salvar_em(str(Caminho.absolute()))

        wb.sheets[7].range("G7").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("I7").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[7].range("D7").value = Termino 

        #Status
        wb.sheets[7].range("E2").value = ""
#--------------------------------------------------------


if __name__ == "__main__":
    import asyncio
    asyncio.run(BeesDelivery())