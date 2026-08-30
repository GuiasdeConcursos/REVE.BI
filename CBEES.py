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
    
async def Produtividade_CBEES():
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "CBEES"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax()

    # Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "CBEES"
    Ativo = False
    #01.05.07.04.02
    if(Ativo):
        OP = "BE_01.05.07.04.02"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\COMERCIAL\BEES\01.05.07.04.02\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[4].range("A4").value = Unidade
        wb.sheets[4].range("B4").value = Nome
        wb.sheets[4].range("C4").value = Inicio
        wb.sheets[4].range("E4").value = "01_05_07_04_02"
        wb.sheets[4].range("F4").value = str(Caminho.absolute())
        try:
            wb.sheets[4].range("H4").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[4].range("J4").value = DataCriacao
            #Status
            wb.sheets[4].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")
        except:
            wb.sheets[4].range("H4").value = 0
            wb.sheets[4].range("J4").value = "NX"

        #------------Início Classe 01_05_07_04_02
 
        #Status
        wb.sheets[4].range("E2").value = "Baixando arquivo CSV"
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

        wb.sheets[4].range("G4").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[4].range("I4").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[4].range("D4").value = Termino 

        #Status
        wb.sheets[4].range("E2").value = ""
#-------------------------------------------
    # Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "CBEES"
    Ativo = True
    #01.11
    if(Ativo):
        OP = "BE_01.11"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\COMERCIAL\BEES\01.11\01.11.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[4].range("A5").value = Unidade
        wb.sheets[4].range("B5").value = Nome
        wb.sheets[4].range("C5").value = Inicio
        wb.sheets[4].range("E5").value = '01_11'
        wb.sheets[4].range("F5").value = str(Caminho.absolute())
        wb.sheets[4].range("H5").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[4].range("J5").value = DataCriacao

        #Status
        wb.sheets[4].range("E2").value = "Movendo arquivo antigo"
        shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")

        #------------Início Classe 01_11
        #Status
        wb.sheets[4].range("E2").value = "Baixando arquivo CSV"
        C01_11 = rpx.sitePromoax_01_11(Processo_Logar_Promax,r"\\Mm04\z\COMERCIAL\BEES\01.11\01.11.csv")
        await C01_11.solicitar_csv()

        wb.sheets[4].range("G5").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[4].range("I5").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[4].range("D5").value = Termino 

        #Status - só muda a letra
        wb.sheets[4].range("E2").value = ""
#--------------------------------------------------------

#---------------BARBACENA-----------------------------
# Barbacena
    Unidade = "Barbacena"
    Nome = "CBEES"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax(1)

    # Barbacena
    Unidade = "Barbacena"
    Nome = "CBEES"
    Ativo = False
    #01.05.07.04.02
    if(Ativo):
        OP = "BE_01.05.07.04.02"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\COMERCIAL\BEES\01.05.07.04.02\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[4].range("A14").value = Unidade
        wb.sheets[4].range("B14").value = Nome
        wb.sheets[4].range("C14").value = Inicio
        wb.sheets[4].range("E14").value = "01_05_07_04_02"
        wb.sheets[4].range("F14").value = str(Caminho.absolute())
        try:
            wb.sheets[4].range("H14").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[4].range("J14").value = DataCriacao
            #Status
            wb.sheets[4].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")
        except:
            wb.sheets[4].range("H14").value = 0
            wb.sheets[4].range("J14").value = "NX"

        #------------Início Classe 01_05_07_04_02
 
        #Status
        wb.sheets[4].range("E2").value = "Baixando arquivo CSV"
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

        wb.sheets[4].range("G14").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[4].range("I14").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[4].range("D14").value = Termino 

        #Status
        wb.sheets[4].range("E2").value = ""
#-------------------------------------------

if __name__ == "__main__":
    import asyncio
    asyncio.run(Produtividade_CBEES())