from datetime import date

class GeradorNomeArquivo:
    def __init__(self, tipo: int = 0):
        """
        Recebe o identificador do tipo (0 para taruma, diferente de 0 para tarumabq).
        """
        self.tipo = tipo

    def obter_nome_arquivo(self) -> str:
        """Retorna o nome formatado do arquivo baseado no mês/ano atual e no tipo."""
        hoje = date.today()
        mes = hoje.strftime("%m")
        ano = hoje.strftime("%Y")
        
        # Define o prefixo conforme a regra solicitada
        prefixo = "taruma" if self.tipo == 0 else "tarumabq"
        
        return f"{prefixo}.{mes}.{ano}.csv"