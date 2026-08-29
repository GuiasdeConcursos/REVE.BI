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

    Ativo = False
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
    Nome = "Ratinga"
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
 
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Ratinga"
    Ativo = False
    #01.20.01.48
    if(Ativo):
        OP = "01.20.01.48_Taruma_Rating"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\ATENDIMENTO\RATING\01.20.01.48\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[1].range("A7").value = Unidade
        wb.sheets[1].range("B7").value = Nome
        wb.sheets[1].range("C7").value = Inicio
        wb.sheets[1].range("E7").value = "01_20_01_48"
        wb.sheets[1].range("F7").value = str(Caminho.absolute())
        wb.sheets[1].range("H7").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("J7").value = DataCriacao

        #Status
        wb.sheets[1].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_20_01_48
        #Status
        wb.sheets[1].range("E2").value = "Baixando arquivo CSV"
        C_01_20_01_48 = rpx.sitePromoax_01_20_01_48(Processo_Logar_Promax)
        await C_01_20_01_48.solicitar_csv()
        await C_01_20_01_48.Salvar_em(str(Caminho.absolute()))

        wb.sheets[1].range("G7").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("I7").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[1].range("D7").value = Termino 

        #Status
        wb.sheets[1].range("E2").value = ""
#--------------------------------------------------------
 
 # Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Rating"
    Ativo = False
    #03.02.37
    if(Ativo):
        OP = "03_02_37_Taruma_Rating"
        Inicio = Datas.datetime.now()

        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\ATENDIMENTO\RATING\03.02.37\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[1].range("A8").value = Unidade
        wb.sheets[1].range("B8").value = Nome
        wb.sheets[1].range("C8").value = Inicio
        wb.sheets[1].range("E8").value = "03_02_37"
        wb.sheets[1].range("F8").value = str(Caminho.absolute())
        wb.sheets[1].range("H8").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("J8").value = DataCriacao

        #Status
        wb.sheets[1].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 03_02_37

        #Status
        wb.sheets[1].range("E2").value = "Baixando arquivo CSV"
        C_03_02_37 = rpx.sitePromoax_03_02_37_RATING(Processo_Logar_Promax)
        await C_03_02_37.solicitar_csv()
        await C_03_02_37.Salvar_em(str(Caminho.absolute()))

        wb.sheets[1].range("G8").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("I8").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[1].range("D8").value = Termino 

        #Status
        wb.sheets[1].range("E2").value = ""
#--------------------------------------------------------

    # Barbacena--------------------------------
    Unidade = "Barbacena"
    Nome = "Rating"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax(1)

    Ativo = False
    #01.05.07.04.02
    if(Ativo):
        OP = "01.05.07.04.02_Taruma_Rating"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\ATENDIMENTO\RATING\01.05.07.04.02\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[1].range("A14").value = Unidade
        wb.sheets[1].range("B14").value = Nome
        wb.sheets[1].range("C14").value = Inicio
        wb.sheets[1].range("E14").value = "01_05_07_04_02"
        wb.sheets[1].range("F14").value = str(Caminho.absolute())
        wb.sheets[1].range("H14").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("J14").value = DataCriacao

        #Status
        wb.sheets[1].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_05_07_04_02
 
        #Status
        wb.sheets[1].range("E2").value = "Baixando arquivo CSV"
        C_01_05_07_04_02 = rpx.sitePromoax_01_05_07_04_02_GERAL(Processo_Logar_Promax)
        await C_01_05_07_04_02.solicitar_csv()
        await C_01_05_07_04_02.Salvar_em(str(Caminho.absolute()))

        wb.sheets[1].range("G14").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("I14").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[1].range("D14").value = Termino 

        #Status
        wb.sheets[1].range("E2").value = ""
#--------------------------------------------------------
# Barbacena
    Unidade = "Barbacena"
    Nome = "Rating"
    Ativo = False
    #01.20.01.47
    if(Ativo):
        OP = "01.20.01.47_Taruma_Rating"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\ATENDIMENTO\RATING\01.20.01.47\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[1].range("A16").value = Unidade
        wb.sheets[1].range("B16").value = Nome
        wb.sheets[1].range("C16").value = Inicio
        wb.sheets[1].range("E16").value = "01_20_01_47"
        wb.sheets[1].range("F16").value = str(Caminho.absolute())
        wb.sheets[1].range("H16").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("J16").value = DataCriacao

        #Status
        wb.sheets[1].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_20_01_47
        #Status
        wb.sheets[1].range("E2").value = "Baixando arquivo CSV"
        C_01_20_01_47 = rpx.sitePromoax_01_20_01_47(Processo_Logar_Promax)
        await C_01_20_01_47.solicitar_csv()
        await C_01_20_01_47.Salvar_em(str(Caminho.absolute()))

        wb.sheets[1].range("G16").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("I16").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[1].range("D16").value = Termino 

        #Status
        wb.sheets[1].range("E2").value = ""
#--------------------------------------------------------
 
# Barbacena
    Unidade = "Barbacena"
    Nome = "Rating"
    Ativo = False
    #01.20.01.48
    if(Ativo):
        OP = "01.20.01.48_Taruma_Rating"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\ATENDIMENTO\RATING\01.20.01.48\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[1].range("A17").value = Unidade
        wb.sheets[1].range("B17").value = Nome
        wb.sheets[1].range("C17").value = Inicio
        wb.sheets[1].range("E17").value = "01_20_01_48"
        wb.sheets[1].range("F17").value = str(Caminho.absolute())
        wb.sheets[1].range("H17").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("J17").value = DataCriacao

        #Status
        wb.sheets[1].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_20_01_48
        #Status
        wb.sheets[1].range("E2").value = "Baixando arquivo CSV"
        C_01_20_01_48 = rpx.sitePromoax_01_20_01_48(Processo_Logar_Promax)
        await C_01_20_01_48.solicitar_csv()
        await C_01_20_01_48.Salvar_em(str(Caminho.absolute()))

        wb.sheets[1].range("G17").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("I17").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[1].range("D17").value = Termino 

        #Status
        wb.sheets[1].range("E2").value = ""
#--------------------------------------------------------
 
# Barbacena
    Unidade = "Barbacena"
    Nome = "Rating"
    Ativo = False
    #03.02.37
    if(Ativo):
        OP = "03_02_37_Taruma_Rating"
        Inicio = Datas.datetime.now()

        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\ATENDIMENTO\RATING\03.02.37\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[1].range("A18").value = Unidade
        wb.sheets[1].range("B18").value = Nome
        wb.sheets[1].range("C18").value = Inicio
        wb.sheets[1].range("E18").value = "03_02_37"
        wb.sheets[1].range("F18").value = str(Caminho.absolute())
        wb.sheets[1].range("H18").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("J18").value = DataCriacao

        #Status
        wb.sheets[1].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 03_02_37

        #Status
        wb.sheets[1].range("E2").value = "Baixando arquivo CSV"
        C_03_02_37 = rpx.sitePromoax_03_02_37_RATING(Processo_Logar_Promax)
        await C_03_02_37.solicitar_csv()
        await C_03_02_37.Salvar_em(str(Caminho.absolute()))

        wb.sheets[1].range("G18").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[1].range("I18").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[1].range("D18").value = Termino 

        #Status
        wb.sheets[1].range("E2").value = ""
#--------------------------------------------------------



if __name__ == "__main__":
    import asyncio
    asyncio.run(NivelServico())