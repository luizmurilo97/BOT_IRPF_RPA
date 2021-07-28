import rpa as r
from pandas import read_excel
from subprocess import Popen
import locale

# CONFIGURAÇÃO PARA MOEDA DO BRASIL
locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')

"""-----------------------------------------------------------------------------
    Leitura do excel em formato xls:
    - Passar o caminho da pasta onde contem a planilha xls.
    exemplo: (r'/home/murilo/rpa/Script/doacoes.xls')
   -----------------------------------------------------------------------------
"""

Cpfs = list(read_excel(r'/home/murilo/rpa/Script/doacoes.xls')['CPF:'])
Nomes = list(read_excel(r'/home/murilo/rpa/Script/doacoes.xls')['NOME:'])
Valor = list(read_excel(r'/home/murilo/rpa/Script/doacoes.xls')['VALOR:'])

"""-----------------------------------------------------------------------------
    Configurações da biblioteca.
   -----------------------------------------------------------------------------
"""
r.init(visual_automation = True, chrome_browser=False)
r.vision('Settings.MoveMouseDelay = 0')

"""-----------------------------------------------------------------------------
    O comando Popen chama o exe do irpf automaticamente quando passamos o caminho.
    - Deixarei um exemplo no windows.
    exemplo: Popen([r"C:\Arquivos de Programas RFB\IRPF2021\IRPF2021.exe"])
   -----------------------------------------------------------------------------
"""
Popen(['java', '-jar', r'/home/murilo/ProgramasRFB/IRPF2021/irpf.jar'])

"""-----------------------------------------------------------------------------
    - Esses comandos abaixo seguidos pelo While, significará o passo a passo, desde
preencher o nome do contribuinte até chegar na tela onde ele lançará os valores da planilha.
    - Será importante você preencher o nome do contribuinte a qual você vai lançar completo, pois
digitando o nome completo evitará possiveis erros, caso algum contribuinte tenha o nome e sobre
nome iguais.
"""
while not r.type(r'imagens/pesquisar_nome.png', 'FULANO TESTE'):
    print('Tentando Fazer...')

while not r.dclick(r'imagens/selecionar_nome.png'):
    print('Tentando Fazer...')

while not r.dclick(r'imagens/rend_isento_e_nao_trib.png'):
    print('Tentando Fazer...')

"""-----------------------------------------------------------------------------
    Lanço for responsavel em pegar as informações da planilha e preencher.
    -Quando passo o (For in Nomes) significa que ele vai fazer as repetições
     de acordo com a quantidade de nomes que tiver na planilha. Desse14 modo, muito
     importante que tenha nome, cpf e valor em cada uma para evitar erros, pois,
     caso uma tenha nome e não tenha valor, essa lógica resultará em erro.
   -----------------------------------------------------------------------------
"""

for count in range(0, len(Nomes)):
    
    while not r.click(r'imagens/button_novo.png'):
        print('Tentando Fazer...')

    while not r.type(r'imagens/code_rend.png', '14'):
        print('Tentando Fazer...')    

    while not r.click(r'imagens/confirmation_rend.png'):
        print('Tentando Fazer...')    

    while not r.type(r'imagens/campo_cpf.png', Cpfs[count]):
        print('Tentando Fazer...')

    while not r.type(r'imagens/name.png', Nomes[count]):
        print('Tentando Fazer...')

    while not r.type(r'imagens/campo_valor.png', locale.currency(Valor[count], grouping=False)[3:]):
        print('Tentando Fazer...')

    while not r.click(r'imagens/btn_ok.png'):
        print('Tentando Fazer...')         