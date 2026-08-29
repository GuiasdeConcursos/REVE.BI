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

async def Produtividade_Gestao_MPD():
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Nível de Serviço"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax()

# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Produtividade Gestão MPD"
    Ativo = False
    #01.20.01.47
    if(Ativo):
        OP = "01.20.01.47_Taruma_GMPD"
        Inicio = Datas.datetime.now()
        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\PRODUTIVIDADE\GESTÃO MPD\01.20.01.47\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[2].range("A4").value = Unidade
        wb.sheets[2].range("B4").value = Nome
        wb.sheets[2].range("C4").value = Inicio
        wb.sheets[2].range("E4").value = "01_20_01_47"
        wb.sheets[2].range("F4").value = str(Caminho.absolute())
        wb.sheets[2].range("H4").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[2].range("J4").value = DataCriacao

        #Status
        wb.sheets[2].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_20_01_47
        #Status
        wb.sheets[2].range("E2").value = "Baixando arquivo CSV"
        C_01_20_01_47 = rpx.sitePromoax_01_20_01_47(Processo_Logar_Promax)
        await C_01_20_01_47.solicitar_csv()
        await C_01_20_01_47.Salvar_em(str(Caminho.absolute()))

        wb.sheets[2].range("G4").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[2].range("I4").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[2].range("D4").value = Termino 

        #Status
        wb.sheets[2].range("E2").value = ""
#--------------------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Produtividade Gestão MPD"
    Ativo = False
    #03.11.20
    if(Ativo):
        OP = "03.11.20_Taruma_GMPD"
        Inicio = Datas.datetime.now()
        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\PRODUTIVIDADE\GESTÃO MPD\03.11.20\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[2].range("A5").value = Unidade
        wb.sheets[2].range("B5").value = Nome
        wb.sheets[2].range("C5").value = Inicio
        wb.sheets[2].range("E5").value = "03_11_20"
        wb.sheets[2].range("F5").value = str(Caminho.absolute())
        wb.sheets[2].range("H5").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[2].range("J5").value = DataCriacao

        #Status
        wb.sheets[2].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 03_11_20
        #Status
        wb.sheets[2].range("E2").value = "Baixando arquivo CSV"
        C_03_11_20 = rpx.sitePromoax_03_11_20(Processo_Logar_Promax)
        await C_03_11_20.solicitar_csv()
        await C_03_11_20.Salvar_em(str(Caminho.absolute()))

        wb.sheets[2].range("G5").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[2].range("I5").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[2].range("D5").value = Termino 

        #Status
        wb.sheets[2].range("E2").value = ""
#--------------------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Produtividade Gestão MPD"
    Ativo = False
    #03.11.49.02
    if(Ativo):
        OP = "03.11.49.02_Taruma_GMPD"
        Inicio = Datas.datetime.now()
        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\PRODUTIVIDADE\GESTÃO MPD\03.11.49.02\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[2].range("A6").value = Unidade
        wb.sheets[2].range("B6").value = Nome
        wb.sheets[2].range("C6").value = Inicio
        wb.sheets[2].range("E6").value = "03_11_49_02"
        wb.sheets[2].range("F6").value = str(Caminho.absolute())
        wb.sheets[2].range("H6").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[2].range("J6").value = DataCriacao

        #Status
        wb.sheets[2].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 03_11_49_02
        #Status
        wb.sheets[2].range("E2").value = "Baixando arquivo CSV"
        C_03_11_49_02 = rpx.sitePromoax_03_11_49_02(Processo_Logar_Promax)
        await C_03_11_49_02.solicitar_csv()
        await C_03_11_49_02.Salvar_em(str(Caminho.absolute()))

        wb.sheets[2].range("G6").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[2].range("I6").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[2].range("D6").value = Termino 

        #Status
        wb.sheets[2].range("E2").value = ""
#--------------------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Produtividade Gestão MPD"
    Ativo = False
    #03.02.37
    if(Ativo):
        OP = "03.02.37_Taruma_GMPD"
        Inicio = Datas.datetime.now()
        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\PRODUTIVIDADE\GESTÃO MPD\03.02.37\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[2].range("A7").value = Unidade
        wb.sheets[2].range("B7").value = Nome
        wb.sheets[2].range("C7").value = Inicio
        wb.sheets[2].range("E7").value = "03_02_37_02"
        wb.sheets[2].range("F7").value = str(Caminho.absolute())
        wb.sheets[2].range("H7").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[2].range("J7").value = DataCriacao

        #Status
        wb.sheets[2].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 03_02_37
        #Status
        wb.sheets[2].range("E2").value = "Baixando arquivo CSV"
        C_03_02_37 = rpx.sitePromoax_03_02_37_MPD(Processo_Logar_Promax)
        await C_03_02_37.solicitar_csv()
        await C_03_02_37.Salvar_em(str(Caminho.absolute()))

        wb.sheets[2].range("G7").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[2].range("I7").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[2].range("D7").value = Termino 

        #Status
        wb.sheets[2].range("E2").value = ""
#--------------------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Produtividade Gestão MPD"
    Ativo = False
    #03.11.40
    if(Ativo):
        OP = "03.11.40_Taruma_GMPD"
        Inicio = Datas.datetime.now()
        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\PRODUTIVIDADE\GESTÃO MPD\03.11.40\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[2].range("A7").value = Unidade
        wb.sheets[2].range("B7").value = Nome
        wb.sheets[2].range("C7").value = Inicio
        wb.sheets[2].range("E7").value = "03_11_40"
        wb.sheets[2].range("F7").value = str(Caminho.absolute())
        wb.sheets[2].range("H7").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[2].range("J7").value = DataCriacao

        #Status
        wb.sheets[2].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 03_11_40
        #Status
        wb.sheets[2].range("E2").value = "Baixando arquivo CSV"
        C_03_11_40 = rpx.sitePromoax_03_11_40(Processo_Logar_Promax)
        await C_03_11_40.solicitar_csv()
        await C_03_11_40.Salvar_em(str(Caminho.absolute()))

        wb.sheets[2].range("G7").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[2].range("I7").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[2].range("D7").value = Termino 

        #Status
        wb.sheets[2].range("E2").value = ""
#--------------------------------------------------------



if __name__ == "__main__":
    import asyncio
    asyncio.run(Produtividade_Gestao_MPD())