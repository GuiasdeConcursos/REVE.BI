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
async def verificarCache(Arquivo: Path):
    import promax.bibliotecas.data_prx as dp
    #Arquivo = Path(__file__).parent / "Promax" / "cache" / "C_01_05_07_04_02" / "Taruma.csv"
    if Arquivo.is_file():
        DtHJ = dp.Datas().data_hoje()
        DataCriacao = (Datas.datetime.fromtimestamp(Arquivo.stat().st_birthtime)).strftime("%d/%m/%Y")
        if DataCriacao == DtHJ:
            return True
        else:
            return False
    else:
        return False
    
async def Volumes():
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Volumes"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax()

#-------------------------------------------
    # Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Volumes"
    Ativo = False
    #01.05.07.04.02
    if(Ativo):
        OP = "VL_01.05.07.04.02"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\COMERCIAL\VOLUME\01.05.07.04.02\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[6].range("A4").value = Unidade
        wb.sheets[6].range("B4").value = Nome
        wb.sheets[6].range("C4").value = Inicio
        wb.sheets[6].range("E4").value = "01_05_07_04_02"
        wb.sheets[6].range("F4").value = str(Caminho.absolute())
        try:
            wb.sheets[6].range("H4").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[6].range("J4").value = DataCriacao
            #Status
            wb.sheets[6].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")
        except:
            wb.sheets[6].range("H4").value = 0
            wb.sheets[6].range("J4").value = "NX"

        #------------Início Classe 01_05_07_04_02
 
        #Status
        wb.sheets[6].range("E2").value = "Baixando arquivo CSV"
        C_01_05_07_04_02 = rpx.sitePromoax_01_05_07_04_02_GERAL(Processo_Logar_Promax)
        while True:
            check_cache = await verificarCache( Path(__file__).parent / "Promax" / "cache" / "C_01_05_07_04_02" / "Taruma.csv")
            if(not check_cache):
                saida = await C_01_05_07_04_02.solicitar_csv()
                if(saida[0] == True):
                    await C_01_05_07_04_02.Salvar_em(str(Caminho.absolute()))
                    await C_01_05_07_04_02.Salvar_em( Path(__file__).parent / "Promax" / "cache" / "C_01_05_07_04_02" / "Taruma.csv")
                    break
                else:
                    await asyncio.sleep(180)
            else:
                print("Copia de cache")
                shutil.copy(str(Path(__file__).parent / "Promax" / "cache" / "C_01_05_07_04_02" / "Taruma.csv"), str(Caminho.absolute()))
                break

        wb.sheets[6].range("G4").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("I4").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[6].range("D4").value = Termino 

        #Status
        wb.sheets[6].range("E2").value = ""
#-------------------------------------------

    # Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Volumes"
    Ativo = False
    #01.11
    if(Ativo):
        OP = "VL_01.11"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\COMERCIAL\VOLUME\01.11\01.11.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[6].range("A5").value = Unidade
        wb.sheets[6].range("B5").value = Nome
        wb.sheets[6].range("C5").value = Inicio
        wb.sheets[6].range("E5").value = '01_11'
        wb.sheets[6].range("F5").value = str(Caminho.absolute())
        wb.sheets[6].range("H5").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("J5").value = DataCriacao

        #Status
        wb.sheets[6].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_11
        #Status
        wb.sheets[6].range("E2").value = "Baixando arquivo CSV"
        C01_11 = rpx.sitePromoax_01_11(Processo_Logar_Promax,r"\\Mm04\z\COMERCIAL\VOLUME\01.11\01.11.csv")
        await C01_11.solicitar_csv()

        wb.sheets[6].range("G5").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("I5").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[6].range("D5").value = Termino 

        #Status - só muda a letra
        wb.sheets[6].range("E2").value = ""
#--------------------------------------------------------

# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Volumes"
    Ativo = False
    #01.20.01.24
    if(Ativo):
        OP = "VL_01.20.01.24"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\COMERCIAL\VOLUME\01.20.01.24\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[6].range("A6").value = Unidade
        wb.sheets[6].range("B6").value = Nome
        wb.sheets[6].range("C6").value = Inicio
        wb.sheets[6].range("E6").value = "01_20_01_24"
        wb.sheets[6].range("F6").value = str(Caminho.absolute())
        wb.sheets[6].range("H6").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("J6").value = DataCriacao

        #Status
        wb.sheets[6].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_20_01_24
        #Status
        wb.sheets[6].range("E2").value = "Baixando arquivo CSV"
        C_01_20_01_24 = rpx.sitePromoax_01_20_01_24(Processo_Logar_Promax)
        await C_01_20_01_24.solicitar_csv()
        await C_01_20_01_24.Salvar_em(str(Caminho.absolute()))

        wb.sheets[6].range("G6").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("I6").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[6].range("D6").value = Termino

        #Status
        wb.sheets[6].range("E2").value = ""
#-------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Volumes"
    Ativo = False
    #03.02.37
    if(Ativo):
        OP = "VL_03_02_37"
        Inicio = Datas.datetime.now()

        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\COMERCIAL\VOLUME\03.02.37\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[6].range("A7").value = Unidade
        wb.sheets[6].range("B7").value = Nome
        wb.sheets[6].range("C7").value = Inicio
        wb.sheets[6].range("E7").value = "03_02_37"
        wb.sheets[6].range("F7").value = str(Caminho.absolute())
        wb.sheets[6].range("H7").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("J7").value = DataCriacao

        #Status
        wb.sheets[6].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 03_02_37

        #Status
        wb.sheets[6].range("E2").value = "Baixando arquivo CSV"
        C_03_02_37 = rpx.sitePromoax_03_02_37_VOLUME(Processo_Logar_Promax)
        await C_03_02_37.solicitar_csv()
        await C_03_02_37.Salvar_em(str(Caminho.absolute()))

        wb.sheets[6].range("G7").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("I7").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[6].range("D7").value = Termino 

        #Status
        wb.sheets[6].range("E2").value = ""
#--------------------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Volumes"
    Ativo = True
    #05.10
    if(Ativo):
        OP = "VL_05_10"
        Inicio = Datas.datetime.now()

        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\COMERCIAL\VOLUME\05.10\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[6].range("A8").value = Unidade
        wb.sheets[6].range("B8").value = Nome
        wb.sheets[6].range("C8").value = Inicio
        wb.sheets[6].range("E8").value = "05_10"
        wb.sheets[6].range("F8").value = str(Caminho.absolute())
        wb.sheets[6].range("H8").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("J8").value = DataCriacao

        #Status
        wb.sheets[6].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 05_10

        #Status
        wb.sheets[6].range("E2").value = "Baixando arquivo CSV"
        C_05_10 = rpx.sitePromoax_05_10(Processo_Logar_Promax)
        await C_05_10.solicitar_csv()
        await C_05_10.Salvar_em(str(Caminho.absolute()))

        wb.sheets[6].range("G8").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("I8").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[6].range("D8").value = Termino 

        #Status
        wb.sheets[6].range("E2").value = ""
#--------------------------------------------------------
#----------------BARBACENA----------------------------
# BARBACENA
    Unidade = "BARBACENA"
    Nome = "Volumes"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax(1)

#-------------------------------------------
    # BARBACENA
    Unidade = "BARBACENA"
    Nome = "Volumes"
    Ativo = False
    #01.05.07.04.02
    if(Ativo):
        OP = "VL_BQ_01.05.07.04.02"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\COMERCIAL\VOLUME\01.05.07.04.02\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[6].range("A14").value = Unidade
        wb.sheets[6].range("B14").value = Nome
        wb.sheets[6].range("C14").value = Inicio
        wb.sheets[6].range("E14").value = "01_05_07_04_02"
        wb.sheets[6].range("F14").value = str(Caminho.absolute())
        try:
            wb.sheets[6].range("H14").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[6].range("J14").value = DataCriacao
            #Status
            wb.sheets[6].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")
        except:
            wb.sheets[6].range("H14").value = 0
            wb.sheets[6].range("J14").value = "NX"

        #------------Início Classe 01_05_07_04_02
 
        #Status
        wb.sheets[6].range("E2").value = "Baixando arquivo CSV"
        C_01_05_07_04_02 = rpx.sitePromoax_01_05_07_04_02_GERAL(Processo_Logar_Promax)
        while True:
            check_cache = await verificarCache( Path(__file__).parent / "Promax" / "cache" / "C_01_05_07_04_02" / "Tarumabq.csv")
            if(not check_cache):
                saida = await C_01_05_07_04_02.solicitar_csv()
                if(saida[0] == True):
                    await C_01_05_07_04_02.Salvar_em(str(Caminho.absolute()))
                    await C_01_05_07_04_02.Salvar_em( Path(__file__).parent / "Promax" / "cache" / "C_01_05_07_04_02" / "Tarumabq.csv")
                    break
                else:
                    await asyncio.sleep(180)
            else:
                print("Copia de cache")
                shutil.copy(str(Path(__file__).parent / "Promax" / "cache" / "C_01_05_07_04_02" / "Tarumabq.csv"), str(Caminho.absolute()))
                break

        wb.sheets[6].range("G14").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("I14").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[6].range("D14").value = Termino 

        #Status
        wb.sheets[6].range("E2").value = ""
#-------------------------------------------

# BARBACENA
    Unidade = "BARBACENA"
    Nome = "Volumes"
    Ativo = False
    #01.20.01.24
    if(Ativo):
        OP = "VL_BQ_01.20.01.24"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\COMERCIAL\VOLUME\01.20.01.24\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[6].range("A16").value = Unidade
        wb.sheets[6].range("B16").value = Nome
        wb.sheets[6].range("C16").value = Inicio
        wb.sheets[6].range("E16").value = "01_20_01_24"
        wb.sheets[6].range("F16").value = str(Caminho.absolute())
        wb.sheets[6].range("H16").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("J16").value = DataCriacao

        #Status
        wb.sheets[6].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")

        #------------Início Classe 01_20_01_24
        #Status
        wb.sheets[6].range("E2").value = "Baixando arquivo CSV"
        C_01_20_01_24 = rpx.sitePromoax_01_20_01_24(Processo_Logar_Promax)
        await C_01_20_01_24.solicitar_csv()
        await C_01_20_01_24.Salvar_em(str(Caminho.absolute()))

        wb.sheets[6].range("G16").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("I16").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[6].range("D16").value = Termino

        #Status
        wb.sheets[6].range("E2").value = ""
#-------------------------------------------
# BARBACENA
    Unidade = "BARBACENA"
    Nome = "Volumes"
    Ativo = False
    #03.02.37
    if(Ativo):
        OP = "VL_BQ_03_02_37"
        Inicio = Datas.datetime.now()

        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\COMERCIAL\VOLUME\03.02.37\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[6].range("A17").value = Unidade
        wb.sheets[6].range("B17").value = Nome
        wb.sheets[6].range("C17").value = Inicio
        wb.sheets[6].range("E17").value = "03_02_37"
        wb.sheets[6].range("F17").value = str(Caminho.absolute())
        wb.sheets[6].range("H17").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("J17").value = DataCriacao

        #Status
        wb.sheets[6].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")

        #------------Início Classe 03_02_37

        #Status
        wb.sheets[6].range("E2").value = "Baixando arquivo CSV"
        C_03_02_37 = rpx.sitePromoax_03_02_37_VOLUME(Processo_Logar_Promax)
        await C_03_02_37.solicitar_csv()
        await C_03_02_37.Salvar_em(str(Caminho.absolute()))

        wb.sheets[6].range("G17").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("I17").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[6].range("D17").value = Termino 

        #Status
        wb.sheets[6].range("E2").value = ""
#--------------------------------------------------------
# BARBACENA
    Unidade = "BARBACENA"
    Nome = "Volumes"
    Ativo = True
    #05.10
    if(Ativo):
        OP = "VL_BQ_05_10"
        Inicio = Datas.datetime.now()

        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\COMERCIAL\VOLUME\05.10\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[6].range("A18").value = Unidade
        wb.sheets[6].range("B18").value = Nome
        wb.sheets[6].range("C18").value = Inicio
        wb.sheets[6].range("E18").value = "05_10"
        wb.sheets[6].range("F18").value = str(Caminho.absolute())
        wb.sheets[6].range("H18").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("J18").value = DataCriacao

        #Status
        wb.sheets[6].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")

        #------------Início Classe 05_10

        #Status
        wb.sheets[6].range("E2").value = "Baixando arquivo CSV"
        C_05_10 = rpx.sitePromoax_05_10(Processo_Logar_Promax)
        await C_05_10.solicitar_csv()
        await C_05_10.Salvar_em(str(Caminho.absolute()))

        wb.sheets[6].range("G18").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[6].range("I18").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[6].range("D18").value = Termino 

        #Status
        wb.sheets[6].range("E2").value = ""
#--------------------------------------------------------





if __name__ == "__main__":
    import asyncio
    asyncio.run(Volumes())