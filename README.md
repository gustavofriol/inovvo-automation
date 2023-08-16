# inovvo-automation

O presente projeto tem como objetivo a obtenção e envio automátio por email de dados de potência horária de quatro PCH's da CELESC Geração. A empresa que receberá esses dados é a Inovvo, cujo objetivo é auxiliar a CELESC a cumprir uma resolução conjunta ANA/ANEEL, que visa o cálculo de defluência dessas usinas. Tais cálculos, são feitos a partir dos dados de potência horária de cada usina, os quais são obtidos diretamente de um banco de dados Oracle.
A ideia para implementação desse sistema, é que o script seja executado a cada hora com o agendador de tarefas do windows.

O script que faz isso tudo acontecer foi desenvolvido em PYTHON.
