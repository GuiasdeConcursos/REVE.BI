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

async def CriticaPedidos():
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "CriticaBEES"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax()

#-------------------------------------------
    # Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #01.09
    if(Ativo):
        OP = "CP_01.09"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\01.09\01.09.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A4").value = Unidade
        wb.sheets[5].range("B4").value = Nome
        wb.sheets[5].range("C4").value = Inicio
        wb.sheets[5].range("E4").value = '01_09'
        wb.sheets[5].range("F4").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H4").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J4").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")
        except:
            pass
        finally:
            pass
        #------------Início Classe 01_09
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C01_09 = rpx.sitePromoax_01_09(Processo_Logar_Promax,r"\\Mm04\z\CRITICA PEDIDOS\01.09\01.09.csv")
        await C01_09.solicitar_csv()

        wb.sheets[5].range("G4").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I4").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D4").value = Termino 

        #Status - só muda a letra
        wb.sheets[5].range("E2").value = ""
#--------------------------------------------------------
    # Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #01.11
    if(Ativo):
        OP = "CP_01.11"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\01.11\01.11.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A5").value = Unidade
        wb.sheets[5].range("B5").value = Nome
        wb.sheets[5].range("C5").value = Inicio
        wb.sheets[5].range("E5").value = '01_11'
        wb.sheets[5].range("F5").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H5").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J5").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")
        except:
            pass
        finally:
            pass
        #------------Início Classe 01_11
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C01_11 = rpx.sitePromoax_01_11(Processo_Logar_Promax,r"\\Mm04\z\CRITICA PEDIDOS\01.11\01.11.csv")
        await C01_11.solicitar_csv()

        wb.sheets[5].range("G5").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I5").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D5").value = Termino 

        #Status - só muda a letra
        wb.sheets[5].range("E2").value = ""
#--------------------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #01.12
    if(Ativo):
        OP = "CP_01.12"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\01.12\01.12.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A6").value = Unidade
        wb.sheets[5].range("B6").value = Nome
        wb.sheets[5].range("C6").value = Inicio
        wb.sheets[5].range("E6").value = '01_12'
        wb.sheets[5].range("F6").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H6").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J6").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")
        except:
            pass
        finally:
            pas
        #------------Início Classe 01_12
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C01_12 = rpx.sitePromoax_01_12(Processo_Logar_Promax,r"\\Mm04\z\CRITICA PEDIDOS\01.12\01.12.csv")
        await C01_12.solicitar_csv()

        wb.sheets[5].range("G6").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I6").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D6").value = Termino 

        #Status - só muda a letra
        wb.sheets[5].range("E2").value = ""
#--------------------------------------------------------

    # Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #01.05.07.04.02
    if(Ativo):
        OP = "CP_01.05.07.04.02"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\01.05.07.04.02\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A7").value = Unidade
        wb.sheets[5].range("B7").value = Nome
        wb.sheets[5].range("C7").value = Inicio
        wb.sheets[5].range("E7").value = "01_05_07_04_02"
        wb.sheets[5].range("F7").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H7").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J7").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")
        except:
            pass
        finally:
            pass
        #------------Início Classe 01_05_07_04_02
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C_01_05_07_04_02 = rpx.sitePromoax_01_05_07_04_02(Processo_Logar_Promax)
        await C_01_05_07_04_02.solicitar_csv()
        await C_01_05_07_04_02.Salvar_em(str(Caminho.absolute()))

        wb.sheets[5].range("G7").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I7").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D7").value = Termino 

        #Status
        wb.sheets[5].range("E2").value = ""
#-------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #02.05.02
    if(Ativo):
        OP = "CP_02.05.02"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\02.05.02\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A8").value = Unidade
        wb.sheets[5].range("B8").value = Nome
        wb.sheets[5].range("C8").value = Inicio
        wb.sheets[5].range("E8").value = "02_05_02"
        wb.sheets[5].range("F8").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H8").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J8").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")
        finally:
            pass
        #------------Início Classe 02_05_02
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C_02_05_02 = rpx.sitePromoax_02_05_02(Processo_Logar_Promax)
        await C_02_05_02.solicitar_csv()
        await C_02_05_02.Salvar_em(str(Caminho.absolute()))

        wb.sheets[5].range("G8").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I8").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D8").value = Termino 

        #Status
        wb.sheets[5].range("E2").value = ""
#-------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #03.01.11
    if(Ativo):
        OP = "CP_03.01.11"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\03.01.11\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A9").value = Unidade
        wb.sheets[5].range("B9").value = Nome
        wb.sheets[5].range("C9").value = Inicio
        wb.sheets[5].range("E9").value = "03_01_11"
        wb.sheets[5].range("F9").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H9").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J9").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")
        finally:
            pass
        #------------Início Classe 03_01_11
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C_03_01_11 = rpx.sitePromoax_03_01_11(Processo_Logar_Promax)
        await C_03_01_11.solicitar_csv()
        await C_03_01_11.Salvar_em(str(Caminho.absolute()))

        wb.sheets[5].range("G9").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I9").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D9").value = Termino

        #Status
        wb.sheets[5].range("E2").value = ""
#-------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #03.01.36.04
    if(Ativo):
        OP = "CP_03.01.36.04"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\03.01.36.04\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A10").value = Unidade
        wb.sheets[5].range("B10").value = Nome
        wb.sheets[5].range("C10").value = Inicio
        wb.sheets[5].range("E10").value = "03_01_36_04"
        wb.sheets[5].range("F10").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H10").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J10").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")
        finally:
            pass
        #------------Início Classe 03_01_36_04
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C_03_01_36_04 = rpx.sitePromoax_03_01_36_04(Processo_Logar_Promax)
        await C_03_01_36_04.solicitar_csv()
        await C_03_01_36_04.Salvar_em(str(Caminho.absolute()))

        wb.sheets[5].range("G10").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I10").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D10").value = Termino

        #Status
        wb.sheets[5].range("E2").value = ""
#-------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #03.02.24
    if(Ativo):
        OP = "CP_03.02.24"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\03.02.24\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A11").value = Unidade
        wb.sheets[5].range("B11").value = Nome
        wb.sheets[5].range("C11").value = Inicio
        wb.sheets[5].range("E11").value = "03_02_24"
        wb.sheets[5].range("F11").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H11").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J11").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")
        finally:
            pass
        #------------Início Classe 03_02_24
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C_03_02_24 = rpx.sitePromoax_03_02_24(Processo_Logar_Promax)
        await C_03_02_24.solicitar_csv()
        await C_03_02_24.Salvar_em(str(Caminho.absolute()))

        wb.sheets[5].range("G11").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I11").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D11").value = Termino

        #Status
        wb.sheets[5].range("E2").value = ""
#-------------------------------------------
# Juiz deFora
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #12.06.01
    if(Ativo):
        OP = "CP_12.06.01"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\12.06.01\Taruma.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A12").value = Unidade
        wb.sheets[5].range("B12").value = Nome
        wb.sheets[5].range("C12").value = Inicio
        wb.sheets[5].range("E12").value = "12_06_01"
        wb.sheets[5].range("F12").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H12").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J12").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}.csv")
        finally:
            pass
        #------------Início Classe 12_06_01
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C_12_06_01 = rpx.sitePromoax_12_06_01(Processo_Logar_Promax)
        await C_12_06_01.solicitar_csv()
        await C_12_06_01.Salvar_em(str(Caminho.absolute()))

        wb.sheets[5].range("G12").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I12").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D12").value = Termino

        #Status
        wb.sheets[5].range("E2").value = ""
#-------------------------------------------


#---------------BARBACENA-----------------------------
# Barbacena
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    import promax.bibliotecas.DRPRX as rpx
    Processo_Logar_Promax = LoginPromax(1)

#-------------------------------------------
# Barbacena
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #01.05.07.04.02
    if(Ativo):
        OP = "CP_BQ_01.05.07.04.02"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\01.05.07.04.02\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A17").value = Unidade
        wb.sheets[5].range("B17").value = Nome
        wb.sheets[5].range("C17").value = Inicio
        wb.sheets[5].range("E17").value = "01_05_07_04_02"
        wb.sheets[5].range("F17").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H17").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J17").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")
        finally:
            pass
        #------------Início Classe 01_05_07_04_02
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C_01_05_07_04_02 = rpx.sitePromoax_01_05_07_04_02(Processo_Logar_Promax)
        await C_01_05_07_04_02.solicitar_csv()
        await C_01_05_07_04_02.Salvar_em(str(Caminho.absolute()))

        wb.sheets[5].range("G17").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I17").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D17").value = Termino 

        #Status
        wb.sheets[5].range("E2").value = ""
#-------------------------------------------
# Barbacena
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #02.05.02
    if(Ativo):
        OP = "CP_BQ_02.05.02"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\02.05.02\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A18").value = Unidade
        wb.sheets[5].range("B18").value = Nome
        wb.sheets[5].range("C18").value = Inicio
        wb.sheets[5].range("E18").value = "02_05_02"
        wb.sheets[5].range("F18").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H18").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J18").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")
        finally:
            pass
        #------------Início Classe 02_05_02
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C_02_05_02 = rpx.sitePromoax_02_05_02(Processo_Logar_Promax)
        await C_02_05_02.solicitar_csv()
        await C_02_05_02.Salvar_em(str(Caminho.absolute()))

        wb.sheets[5].range("G18").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I18").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D18").value = Termino 

        #Status
        wb.sheets[5].range("E2").value = ""
#-------------------------------------------
# Barbacena
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #03.01.11
    if(Ativo):
        OP = "CP_BQ_03.01.11"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\03.01.11\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A19").value = Unidade
        wb.sheets[5].range("B19").value = Nome
        wb.sheets[5].range("C19").value = Inicio
        wb.sheets[5].range("E19").value = "03_01_11"
        wb.sheets[5].range("F19").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H19").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J19").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")
        finally:
            pass
        #------------Início Classe 03_01_11
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C_03_01_11 = rpx.sitePromoax_03_01_11(Processo_Logar_Promax)
        await C_03_01_11.solicitar_csv()
        await C_03_01_11.Salvar_em(str(Caminho.absolute()))

        wb.sheets[5].range("G19").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I19").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D19").value = Termino

        #Status
        wb.sheets[5].range("E2").value = ""
#-------------------------------------------
# Barbacena
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #03.01.36.04
    if(Ativo):
        OP = "CP_BQ_03.01.36.04"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\03.01.36.04\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A20").value = Unidade
        wb.sheets[5].range("B20").value = Nome
        wb.sheets[5].range("C20").value = Inicio
        wb.sheets[5].range("E20").value = "03_01_36_04"
        wb.sheets[5].range("F20").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H20").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J20").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")
        finally:
            pass
        #------------Início Classe 03_01_36_04
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C_03_01_36_04 = rpx.sitePromoax_03_01_36_04(Processo_Logar_Promax)
        await C_03_01_36_04.solicitar_csv()
        await C_03_01_36_04.Salvar_em(str(Caminho.absolute()))

        wb.sheets[5].range("G20").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I20").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D20").value = Termino

        #Status
        wb.sheets[5].range("E2").value = ""
#-------------------------------------------
# Barbacena
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #03.02.24
    if(Ativo):
        OP = "CP_BQ_03.02.24"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\03.02.24\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A21").value = Unidade
        wb.sheets[5].range("B21").value = Nome
        wb.sheets[5].range("C21").value = Inicio
        wb.sheets[5].range("E21").value = "03_02_24"
        wb.sheets[5].range("F21").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H21").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J21").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")
        finally:
            pass
        #------------Início Classe 03_02_24
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C_03_02_24 = rpx.sitePromoax_03_02_24(Processo_Logar_Promax)
        await C_03_02_24.solicitar_csv()
        await C_03_02_24.Salvar_em(str(Caminho.absolute()))

        wb.sheets[5].range("G21").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I21").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D21").value = Termino

        #Status
        wb.sheets[5].range("E2").value = ""
#-------------------------------------------
# Barbacena
    Unidade = "Juiz de Fora"
    Nome = "Crítica de Pedidos"
    Ativo = True
    #12.06.01
    if(Ativo):
        OP = "CP_BQ_12.06.01"
        Inicio = Datas.datetime.now()
        Caminho = Path(r"\\Mm04\z\CRITICA PEDIDOS\12.06.01\Tarumabq.csv")

        wb = xw.apps.active.books.active  
        wb.sheets[5].range("A22").value = Unidade
        wb.sheets[5].range("B22").value = Nome
        wb.sheets[5].range("C22").value = Inicio
        wb.sheets[5].range("E22").value = "12_06_01"
        wb.sheets[5].range("F22").value = str(Caminho.absolute())
        try:
            wb.sheets[5].range("H22").value = Caminho.stat().st_size
            DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
            wb.sheets[5].range("J22").value = DataCriacao

            #Status
            wb.sheets[5].range("E2").value = "Movendo arquivo antigo"
            shutil.move(str(Caminho), f"c:\\ArquivosAntigos\\{OP}bq.csv")
        finally:
            pass
        #------------Início Classe 12_06_01
        #Status
        wb.sheets[5].range("E2").value = "Baixando arquivo CSV"
        C_12_06_01 = rpx.sitePromoax_12_06_01(Processo_Logar_Promax)
        await C_12_06_01.solicitar_csv()
        await C_12_06_01.Salvar_em(str(Caminho.absolute()))

        wb.sheets[5].range("G22").value = Caminho.stat().st_size
        DataCriacao = Datas.datetime.fromtimestamp(Caminho.stat().st_birthtime)
        wb.sheets[5].range("I22").value = DataCriacao

        Termino = Datas.datetime.now()
        wb.sheets[5].range("D22").value = Termino

        #Status
        wb.sheets[5].range("E2").value = ""
#-------------------------------------------


if __name__ == "__main__":
    import asyncio
    asyncio.run(CriticaPedidos())