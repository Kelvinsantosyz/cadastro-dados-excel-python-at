"""Importa a biblioteca openpyxl para manipular arquivos do Excel"""
from openpyxl import Workbook, load_workbook
"""Importa as funções do módulo validor """
from validor import verificar_cpf_repetido,verificar_existencia_dados,\
validar_dados, usuarios

workbook = load_workbook(filename='dados cadastrado.xlsx')
worksheet = workbook.active
worksheet.title = "Usuários"

def armazenar_dados(usuarios):
    """Armazena os dados dos usuários na planilha."""
    while True:
        print("Digite uma das opções: ")
        opcao_menu = input("[i]nserir  [s]air: [l]istar: ")
        
        if opcao_menu == "s":
            workbook.save(filename='dados cadastrado.xlsx')
            print("Você saiu do programa")
            return "programa encerrado"
        
        if opcao_menu =='i':
            
            nome = input("Digite o seu nome: ")
            senha = input("Digite sua senha: ")
            cpf = input("Digite o seu CPF: ")
            
            try:
                valida, erro = validar_dados(nome, senha, cpf, usuarios)
                if valida:
                    if any(usuario['cpf'] == cpf for usuario in usuarios):
                        print("CPF já cadastrado")
                    else:
                        print("Dados válidos")
                        usuario = {
                            "nome": nome,
                            "senha": senha,
                            "cpf": cpf
                        }
                        usuarios.append(usuario)
                        print("Usuário cadastrado:")
                        linha = len(usuarios) + 1
                        worksheet[f"B{linha}"] = nome
                        worksheet[f"C{linha}"] = senha
                        worksheet[f"D{linha}"] = cpf
                        workbook.save(filename='dados cadastrado.xlsx')
                else:
                    print("Dados inválidos:", erro)
            
            except IndexError:
                print("Digite apenas valores valido")
            
            except Exception:
                print("Ocorreu erro inesperado, tente novamente")
        
        if opcao_menu =='l':
            for usuario in usuarios:
                print("Lista de usuários:", usuario)
                

#Salva todas as alterações feitas na planilha "dados cadastrado.xlsx"""
workbook.save(filename='dados cadastrado.xlsx')

#Executa a função para armazenar os dados dos usuários
armazenar_dados(usuarios)


