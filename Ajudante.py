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

async def Produtividade_Ajudante():
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Ajudante"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax()

    # Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Ajudante"
    Ativo = True
    #03.02.37
    if(Ativo):
        OP = "AJ_03.02.37"
        Inicio = Datas.datetime.now()
        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\PRODUTIVIDADE\ID AJUDANTE\03.02.37\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[3].range("A4").value = Unidade
        wb.sheets[3].range("B4").value = Nome
        wb.sheets[3].range("C4").value = Inicio
        wb.sheets[3].range("E4").value = "03_02_37"
        wb.sheets[3].range("F4").value = str(Caminho.absolute())
        wb.sheets[3].range("H4").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("J4").value = DataCriacao

        #Status
        wb.sheets[3].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 03_02_37
        #Status
        wb.sheets[3].range("E2").value = "Baixando arquivo CSV"
        C_03_02_37 = rpx.sitePromoax_03_02_37_AJUD(Processo_Logar_Promax)
        await C_03_02_37.solicitar_csv()
        await C_03_02_37.Salvar_em(str(Caminho.absolute()))

        wb.sheets[3].range("G4").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("I4").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[3].range("D4").value = Termino 

        #Status
        wb.sheets[3].range("E2").value = ""
#--------------------------------------------------------

    # Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Ajudante"
    Ativo = True
    #01.11
    if(Ativo):
        OP = "AJ_01.11"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\PRODUTIVIDADE\ID AJUDANTE\01.11\01.11.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[3].range("A5").value = Unidade
        wb.sheets[3].range("B5").value = Nome
        wb.sheets[3].range("C5").value = Inicio
        wb.sheets[3].range("E5").value = '01_11'
        wb.sheets[3].range("F5").value = str(Caminho.absolute())
        wb.sheets[3].range("H5").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("J5").value = DataCriacao

        #Status
        wb.sheets[3].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_05_07_04_02
        #Status
        wb.sheets[3].range("E2").value = "Baixando arquivo CSV"
        C01_11 = rpx.sitePromoax_01_11(Processo_Logar_Promax,r"\\Mm04\z\PRODUTIVIDADE\ID AJUDANTE\01.11\01.11.csv")
        await C01_11.solicitar_csv()

        wb.sheets[3].range("G5").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("I5").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[3].range("D5").value = Termino 

        #Status - só muda a letra
        wb.sheets[3].range("E2").value = ""
#--------------------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Ajudante"
    Ativo = True
    #01.20.01.48
    if(Ativo):
        OP = "AJ_01.20.01.48"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\PRODUTIVIDADE\ID AJUDANTE\01.20.01.48\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[3].range("A6").value = Unidade
        wb.sheets[3].range("B6").value = Nome
        wb.sheets[3].range("C6").value = Inicio
        wb.sheets[3].range("E6").value = "01_20_01_48"
        wb.sheets[3].range("F6").value = str(Caminho.absolute())
        wb.sheets[3].range("H6").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("J6").value = DataCriacao

        #Status
        wb.sheets[3].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_20_01_48
        #Status
        wb.sheets[3].range("E2").value = "Baixando arquivo CSV"
        C_01_20_01_48 = rpx.sitePromoax_01_20_01_48(Processo_Logar_Promax)
        await C_01_20_01_48.solicitar_csv()
        await C_01_20_01_48.Salvar_em(str(Caminho.absolute()))

        wb.sheets[3].range("G6").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("I6").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[3].range("D6").value = Termino 

        #Status
        wb.sheets[3].range("E2").value = ""
#-------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Ajudante"
    Ativo = True
    #03.11.20
    if(Ativo):
        OP = "AJ_03.11.20"
        Inicio = Datas.datetime.now()
        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\PRODUTIVIDADE\ID AJUDANTE\03.11.20\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[3].range("A7").value = Unidade
        wb.sheets[3].range("B7").value = Nome
        wb.sheets[3].range("C7").value = Inicio
        wb.sheets[3].range("E7").value = "03_11_20"
        wb.sheets[3].range("F7").value = str(Caminho.absolute())
        wb.sheets[3].range("H7").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("J7").value = DataCriacao

        #Status
        wb.sheets[3].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 03_11_20
        #Status
        wb.sheets[3].range("E2").value = "Baixando arquivo CSV"
        C_03_11_20 = rpx.sitePromoax_03_11_20(Processo_Logar_Promax)
        await C_03_11_20.solicitar_csv()
        await C_03_11_20.Salvar_em(str(Caminho.absolute()))

        wb.sheets[3].range("G7").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("I7").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[3].range("D7").value = Termino 

        #Status
        wb.sheets[3].range("E2").value = ""
#--------------------------------------------------------

#----------------------BARBACENA-----------------------------
# Barbacena
    Unidade = "Barbacena"
    Nome = "Ajudante"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax(1)

    # Barbacena
    Unidade = "Barbacena"
    Nome = "Ajudante"
    Ativo = True
    #03.02.37
    if(Ativo):
        OP = "AJ_BQ_03.02.37"
        Inicio = Datas.datetime.now()
        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\PRODUTIVIDADE\ID AJUDANTE\03.02.37\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[3].range("A14").value = Unidade
        wb.sheets[3].range("B14").value = Nome
        wb.sheets[3].range("C14").value = Inicio
        wb.sheets[3].range("E14").value = "03_02_37"
        wb.sheets[3].range("F14").value = str(Caminho.absolute())
        wb.sheets[3].range("H14").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("J14").value = DataCriacao

        #Status
        wb.sheets[3].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")

        #------------Início Classe 03_02_37
        #Status
        wb.sheets[3].range("E2").value = "Baixando arquivo CSV"
        C_03_02_37 = rpx.sitePromoax_03_02_37_AJUD(Processo_Logar_Promax)
        await C_03_02_37.solicitar_csv()
        await C_03_02_37.Salvar_em(str(Caminho.absolute()))

        wb.sheets[3].range("G14").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("I14").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[3].range("D14").value = Termino 

        #Status
        wb.sheets[3].range("E2").value = ""
#--------------------------------------------------------

# Barbacena
    Unidade = "Barbacena"
    Nome = "Ajudante"
    Ativo = True
    #01.20.01.48
    if(Ativo):
        OP = "AJ_BQ_01.20.01.48"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\PRODUTIVIDADE\ID AJUDANTE\01.20.01.48\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[3].range("A15").value = Unidade
        wb.sheets[3].range("B15").value = Nome
        wb.sheets[3].range("C15").value = Inicio
        wb.sheets[3].range("E15").value = "01_20_01_48"
        wb.sheets[3].range("F15").value = str(Caminho.absolute())
        wb.sheets[3].range("H15").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("J15").value = DataCriacao

        #Status
        wb.sheets[3].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")

        #------------Início Classe 01_20_01_48
        #Status
        wb.sheets[3].range("E2").value = "Baixando arquivo CSV"
        C_01_20_01_48 = rpx.sitePromoax_01_20_01_48(Processo_Logar_Promax)
        await C_01_20_01_48.solicitar_csv()
        await C_01_20_01_48.Salvar_em(str(Caminho.absolute()))

        wb.sheets[3].range("G15").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("I15").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[3].range("D15").value = Termino 

        #Status
        wb.sheets[3].range("E2").value = ""
#-------------------------------------------
# Barbacena
    Unidade = "Barbacena"
    Nome = "Ajudante"
    Ativo = True
    #03.11.20
    if(Ativo):
        OP = "AJ_BQ_03.11.20"
        Inicio = Datas.datetime.now()
        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\PRODUTIVIDADE\ID AJUDANTE\03.11.20\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[3].range("A16").value = Unidade
        wb.sheets[3].range("B16").value = Nome
        wb.sheets[3].range("C16").value = Inicio
        wb.sheets[3].range("E16").value = "03_11_20"
        wb.sheets[3].range("F16").value = str(Caminho.absolute())
        wb.sheets[3].range("H16").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("J16").value = DataCriacao

        #Status
        wb.sheets[3].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")

        #------------Início Classe 03_11_20
        #Status
        wb.sheets[3].range("E2").value = "Baixando arquivo CSV"
        C_03_11_20 = rpx.sitePromoax_03_11_20(Processo_Logar_Promax)
        await C_03_11_20.solicitar_csv()
        await C_03_11_20.Salvar_em(str(Caminho.absolute()))

        wb.sheets[3].range("G16").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[3].range("I16").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[3].range("D16").value = Termino 

        #Status
        wb.sheets[3].range("E2").value = ""
#--------------------------------------------------------


if __name__ == "__main__":
    import asyncio
    asyncio.run(Produtividade_Ajudante())