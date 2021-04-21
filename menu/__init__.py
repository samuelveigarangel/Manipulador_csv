from classe_csv import *


def menu(msg):
    print('-' * 30)
    print(msg.center(30))
    print('-' * 30)


def menu_option():

    a = Xlsx()
    while True:
        menu('GERENCIADOR DE PLANILHAS XLSX')
        print('1 - Criar Planilha\n2 - Carregar Planilha\n3 - Ver Planilhas\n4 - Ver conteudo de planilha (Crie ou carregue sua '
              'planilha primeiro)\n5 - Inserir cabeçalho\n6 - Inserir dados\n7 - Sair\n\033[031mDIGITE O NÚMERO DA OPÇÃO\033[0;0m')
        try:
            option = int(input('O que você deseja fazer: '))
        except (TypeError, ValueError):
            print('\033[031mErro. Opção Invalida. Digite a opção corretamente.\033[0;0m')
        else:
            if option == 1:
                res, arq = a.criar_planilha()
            elif option == 2:
                res, arq = a.carregar_planilha()
            elif option == 3:
                a.ver_planilha(res)
            elif option == 4:
                if 'res' and 'arq' not in locals():
                    print('\033[031mCrie ou carregue um arquivo para inserir dados\033[0;0m')
                else:
                    a.ver_conteudo(res)
            elif option == 5:
                a.inserir_cabecalho(res, arq)
            elif option == 6:
                if 'res' and 'arq' not in locals():
                    print('\033[031mCrie ou carregue um arquivo para inserir dados\033[0;0m')
                else:
                    a.inserir_dados(res, arq)
            elif option == 7:
                break
            else:
                print('\033[031mErro. Opção Invalida. Digite a opção corretamente.\033[0;0m')
    print('Obrigado. Até a próxima!')
