import asyncio
Dicionario_Variaveis = {

}

async def main():

    #PROCESSO DE LOGIN NO PROMAX
    import promax.loginPromax as lpx
    unidade = 0

    sessao = lpx.Promax("Alexandre", "Revenda.44", unidade)
    sessao.login()
    #------------------------------------------------------

    import promax.bibliotecas.DRPRX as rpx

    """
        NÍVEL DE SERVIÇO

        1 - 01.05.07.04.02
        2 - 01.11
        3 - 01.20.01.47
        4 - 01.20.01.24
        5 - 03.01.47.01
        6 - 03.02.37

    """
    classe = rpx.sitePromoax_01_05_07_04_02(sessao)
    await classe.solicitar_csv()
    await classe.Salvar_em("C:\\Alexandre\\Teste\\taruma.sitePromoax_01_05_07_04_02.agosto.csv")

    pass

if __name__ == "__main__":
    asyncio.run(main())