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
    
async def BeesDelivery():
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Bees Delivery"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax()

#-------------------------------------------
    # Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Bees Delivery"
    Ativo = True
    #01.05.07.04.02
    if(Ativo):
        OP = "BD_01.05.07.04.02"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\ATENDIMENTO\BEES DELIVERY\01.05.07.04.02\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[7].range("A4").value = Unidade
        wb.sheets[7].range("B4").value = Nome
        wb.sheets[7].range("C4").value = Inicio
        wb.sheets[7].range("E4").value = "01_05_07_04_02"
        wb.sheets[7].range("F4").value = str(Caminho.absolute())
        try:
            wb.sheets[7].range("H4").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[7].range("J4").value = DataCriacao
            #Status
            wb.sheets[7].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")
        except:
            wb.sheets[7].range("H4").value = 0
            wb.sheets[7].range("J4").value = "NX"

        #------------Início Classe 01_05_07_04_02
 
        #Status
        wb.sheets[7].range("E2").value = "Baixando arquivo CSV"
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
    Nome = "bees Delivery"
    Ativo = False
    #01.20.01.47
    if(Ativo):
        OP = "BD_01.20.01.47"
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
    Nome = "Bees Delivery"
    Ativo = False
    #03.11.20
    if(Ativo):
        OP = "BD_03.11.20"
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
    Nome = "Bees Delivery"
    Ativo = False
    #03.02.24
    if(Ativo):
        OP = "BD_03.02.24"
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

#------------BARBACENA-----------------------------
# BARBACENA
    Unidade = "Barbacena"
    Nome = "Bees Delivery"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax(1)

# BARBACENA
    Unidade = "Barbacena"
    Nome = "Bees Delivery"
    Ativo = True
    #01.05.07.04.02
    if(Ativo):
        OP = "BD_BQ_01.05.07.04.02"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\ATENDIMENTO\BEES DELIVERY\01.05.07.04.02\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[7].range("A14").value = Unidade
        wb.sheets[7].range("B14").value = Nome
        wb.sheets[7].range("C14").value = Inicio
        wb.sheets[7].range("E14").value = "01_05_07_04_02"
        wb.sheets[7].range("F14").value = str(Caminho.absolute())
        try:
            wb.sheets[7].range("H14").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[7].range("J14").value = DataCriacao
            #Status
            wb.sheets[7].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")
        except:
            wb.sheets[7].range("H14").value = 0
            wb.sheets[7].range("J14").value = "NX"

        #------------Início Classe 01_05_07_04_02
 
        #Status
        wb.sheets[7].range("E2").value = "Baixando arquivo CSV"
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

        wb.sheets[7].range("G14").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("I14").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[7].range("D14").value = Termino 

        #Status
        wb.sheets[7].range("E2").value = ""
#-------------------------------------------
# BARBACENA
    Unidade = "Barbacena"
    Unidade = "Juiz de Fora"
    Nome = "bees Delivery"
    Ativo = False
    #01.20.01.47
    if(Ativo):
        OP = "BD_BQ_01.20.01.47"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\ATENDIMENTO\BEES DELIVERY\01.20.01.47\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[7].range("A15").value = Unidade
        wb.sheets[7].range("B15").value = Nome
        wb.sheets[7].range("C15").value = Inicio
        wb.sheets[7].range("E15").value = "01_20_01_47"
        wb.sheets[7].range("F15").value = str(Caminho.absolute())
        wb.sheets[7].range("H15").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("J15").value = DataCriacao

        #Status
        wb.sheets[7].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")

        #------------Início Classe 01_20_01_47
        #Status
        wb.sheets[7].range("E2").value = "Baixando arquivo CSV"
        C_01_20_01_47 = rpx.sitePromoax_01_20_01_47(Processo_Logar_Promax)
        await C_01_20_01_47.solicitar_csv()
        await C_01_20_01_47.Salvar_em(str(Caminho.absolute()))

        wb.sheets[7].range("G15").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("I15").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[7].range("D15").value = Termino 

        #Status
        wb.sheets[7].range("E2").value = ""
#--------------------------------------------------------
# BARBACENA
    Unidade = "Barbacena"
    Nome = "Bees Delivery"
    Ativo = False
    #03.11.20
    if(Ativo):
        OP = "BD_BQ_03.11.20"
        Inicio = Datas.datetime.now()
        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\ATENDIMENTO\BEES DELIVERY\03.11.20\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[7].range("A16").value = Unidade
        wb.sheets[7].range("B16").value = Nome
        wb.sheets[7].range("C16").value = Inicio
        wb.sheets[7].range("E16").value = "03_11_20"
        wb.sheets[7].range("F16").value = str(Caminho.absolute())
        wb.sheets[7].range("H16").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("J16").value = DataCriacao

        #Status
        wb.sheets[7].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")

        #------------Início Classe 03_11_20
        #Status
        wb.sheets[7].range("E2").value = "Baixando arquivo CSV"
        C_03_11_20 = rpx.sitePromoax_03_11_20(Processo_Logar_Promax)
        await C_03_11_20.solicitar_csv()
        await C_03_11_20.Salvar_em(str(Caminho.absolute()))

        wb.sheets[7].range("G16").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("I16").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[7].range("D16").value = Termino 

        #Status
        wb.sheets[7].range("E2").value = ""
#--------------------------------------------------------
# BARBACENA
    Unidade = "Barbacena"
    Nome = "Bees Delivery"
    Ativo = False
    #03.02.24
    if(Ativo):
        OP = "BD_BQ_03.02.24"
        Inicio = Datas.datetime.now()
        import promax.bibliotecas.nome_bi as nb
        CName = nb.GeradorNomeArquivo(Processo_Logar_Promax.getCodUnidade())
        Name = CName.obter_nome_arquivo()
        Caminho = Path(fr"\\Mm04\z\ATENDIMENTO\BEES DELIVERY\03.02.24\{Name}")

        wb = xw.apps.active.books.active  
        wb.sheets[7].range("A17").value = Unidade
        wb.sheets[7].range("B17").value = Nome
        wb.sheets[7].range("C17").value = Inicio
        wb.sheets[7].range("E17").value = "03_02_24"
        wb.sheets[7].range("F17").value = str(Caminho.absolute())
        wb.sheets[7].range("H17").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("J17").value = DataCriacao

        #Status
        wb.sheets[7].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")

        #------------Início Classe 03_02_24
        #Status
        wb.sheets[7].range("E2").value = "Baixando arquivo CSV"
        C_03_02_24 = rpx.sitePromoax_03_02_24(Processo_Logar_Promax)
        await C_03_02_24.solicitar_csv()
        await C_03_02_24.Salvar_em(str(Caminho.absolute()))

        wb.sheets[7].range("G17").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[7].range("I17").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[7].range("D17").value = Termino 

        #Status
        wb.sheets[7].range("E2").value = ""
#--------------------------------------------------------


if __name__ == "__main__":
    import asyncio
    asyncio.run(BeesDelivery())