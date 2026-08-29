from datetime import date, timedelta
from calendar import monthrange

class Datas:
    def __init__(self):
        self.__str__ = "Classe de manipulação de datas"

    def formatar_data(self, data):
        """Método auxiliar para retornar a data no padrão brasileiro (DD/MM/AAAA)."""
        return data.strftime("%d/%m/%Y")

    def primeiro_dia_do_mes(self):
        """Retorna a data do primeiro dia do mês atual formatada."""
        hoje = date.today()
        primeiro = hoje.replace(day=1)
        return self.formatar_data(primeiro)

    def dia_anterior(self):
        """Retorna a data do dia anterior ao atual (ontem) formatada."""
        hoje = date.today()
        ontem = hoje - timedelta(days=1)
        return self.formatar_data(ontem)

    def tres_meses_atras(self):
        """Retorna a data de 3 meses anterior à data atual formatada."""
        hoje = date.today()
        mes_alvo = hoje.month - 3
        ano_alvo = hoje.year
        if mes_alvo <= 0:
            mes_alvo += 12
            ano_alvo -= 1
        
        try:
            data_calculada = hoje.replace(year=ano_alvo, month=mes_alvo)
        except ValueError:
            ultimo_dia = monthrange(ano_alvo, mes_alvo)[1]
            data_calculada = hoje.replace(year=ano_alvo, month=mes_alvo, day=ultimo_dia)
            
        return self.formatar_data(data_calculada)

    def data_hoje(self):
        """Retorna a data atual formatada."""
        return self.formatar_data(date.today())

    def mes_ano_atual(self):
        """Retorna o mês e o ano atual no formato MM/YYYY (ex: 08/2026)."""
        return date.today().strftime("%m/%Y")