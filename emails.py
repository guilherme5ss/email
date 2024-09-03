""" LINKS
https://stackoverflow.com/questions/77387302/send-email-with-python-in-windows-11-with-outlook-new
https://learn.microsoft.com/en-us/graph/use-the-api
https://answers.microsoft.com/en-us/outlook_com/forum/all/cant-send-mail-via-outlooknewui-using-python/fdda8d92-abcd-4b10-9eab-afba9f6b7924
"""
import psutil
import os
import win32com.client
from collections import defaultdict
import unicodedata
import re
from arvore import Arvore

# Variáveis
# Lista com nomes de pastas reservadas Outlook em inglês e português
lista_outlook = ["Calendário", "Calendar", "Conflitos", "Conflicts", "Contatos", "Contacts", "Itens Excluídos", "DeletedItems", "Rascunhos", "Drafts", "Caixa de Entrada", "Inbox", "Diário", "Journal", "Lixo Eletrônico", "Junk", "Falhas Locais", "LocalFailures", "Pastas Gerenciadas", "ManagedEmail", "Anotações", "Notes", "Caixa de Saída", "Outbox", "Itens Enviados", "SentMail", "Falhas de Servidor", "ServerFailures", "Contatos Sugeridos", "SuggestedContacts", "Problemas de Sincroniza", "SyncIssues", "Tarefas", "Tasks", "Tarefas Pendentes", "ToDo", "Pastas Públicas", "FoldersAllPublicFolders", "Alimentações de RSS", "RssFeeds"]

padrao_outlook = { # Fonte: https://learn.microsoft.com/pt-br/office/vba/api/outlook.oldefaultfolders
    9:  ["calendario", "calendar"],
    19: ["conflitos", "conflicts"],
    10: ["contatos", "contacts"],
    3:  ["itensexcluidos", "deleteditems"],
    16: ["rascunhos", "drafts"],
    6:  ["caixadeentrada", "inbox"],
    11: ["diario", "journal"],
    23: ["lixoeletronico", "junk"],
    21: ["falhaslocais", "localfailures"],
    29: ["pastasgerenciadas", "managedemail"],
    12: ["anotacoes", "notes"],
    4:  ["caixadesaida", "outbox"],
    5:  ["itensenviados", "sentmail"],
    22: ["falhasdeservidor", "serverfailures"],
    30: ["contatossugeridos", "suggestedcontacts"],
    20: ["problemasdesincroniza", "syncissues"],
    13: ["tarefas", "tasks"],
    28: ["tarefaspendentes", "todo"],
    18: ["pastaspublicas", "foldersallpublicfolders"],
    25: ["alimentacoesderss", "rssfeeds"]
}
# DICIONÁRIO
def normalizar_string(entrada): # Remover espaços, normaliza uma string removendo acentos e caracteres especiais
    if type(entrada) == str: 
        # Remove acentos
        nfkd = unicodedata.normalize('NFKD', entrada)
        texto_sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
        
        texto_limpo = re.sub(r'[^a-zA-Z0-9]', '', texto_sem_acento) # Remove caracteres especiais e espaços
        
        return texto_limpo.lower() # Converte para minúsculas
    else: # Caso "entrada" não seja uma string
        return entrada

def chave_correspondente(valor, dicionario, normalizar = True): # Encontra chave en dicionário baseado em um valor, que pode estar em uma lista  
    # Baseado em: https://stackoverflow.com/questions/66724197/
    if normalizar == True:
        valor = normalizar_string(valor)
    for k, v in dicionario.items():
        if isinstance(v, (list, tuple, str)):  # Checa se "v" é iterável
            if valor in v:
                return k
        else:
            if valor == v:  # Caso que "v" é um valor não iterável
                return k        
# E-MAIL // OUTLOOK
def status_exe_outlook(): # Função verifica estado de execução do processo "OUTLOOK.EXE"
    for p in psutil.process_iter(attrs=['pid', 'name']):
        if "OUTLOOK.EXE" in p.info['name']:
            print("Sim", p.info['name'], "executando")
            break
    else:
        print("Não, Outlook não está em execução")
        os.startfile("outlook")
        print("Outlook esta inicializando...")

def listar_inbox(check = False, count = False): # A partir de um laço lista as mensagens na "Caixa de Entrada"
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") # Inicializando a conexão com o Outlook
    if check == True: # Se a condição for atendida executa a função "status_exe_outlook()"
        status_exe_outlook()

    # Acessar a pasta de entrada. Mais informações: https://learn.microsoft.com/pt-br/office/vba/api/outlook.oldefaultfolders
    inbox = outlook.GetDefaultFolder(6)  # 6 é o código para a pasta "Inbox"

    messages = inbox.Items # Obter todos os e-mails na pasta

    for i, message in enumerate(messages):
        if count == True: # Um contador aparece antes das mensagens
            print(str(i+1)+": "+str(message))
        else: # E-mails aparecem sem um contador
            print(message)

def listar_pastas(pasta, nivel=0): # Lista todas as pastas do email
    ''' Utilização:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") # Inicializando a conexão com o Outlook
    for folder in outlook.Folders:
        listar_pastas(folder)
    '''
    # Imprime o nome da pasta com indentação de acordo com o nível
    print("    " * nivel + pasta.Name)
    
    # Tenta acessar subpastas, se houver
    try:
        subpastas = pasta.Folders
        for subpasta in subpastas:
            listar_pastas(subpasta, nivel + 1)
    except Exception as e:
        print(f"Erro ao acessar subpastas de {pasta.Name}: {e}")

def gerar_arvore_email(pasta, arvore=None): # Cria uma estrutura de árvore a partir de uma pasta do email. Utilizando "from arvore import Arvore"
    # Adiciona a pasta atual à árvore
    if arvore is None:
        arvore = Arvore(pasta.Name)  # Caso seja a raiz

    # Tenta acessar subpastas, se houver
    try:
        subpastas = pasta.Folders
        for subpasta in subpastas:
            # Adiciona a subpasta como filho na árvore
            novo_filho = arvore.add_child(subpasta.Name)
            # Recursivamente chama para adicionar as subpastas deste filho
            gerar_arvore_email(subpasta, novo_filho)
    except Exception as e:
        print(f"Erro ao acessar subpastas de {pasta.Name}: {e}")

    return arvore  # Retorna a árvore completa
# FUNÇÕES PARA ACESSO OUTLOOK
# 3 métodos que juntos percorrem uma pasta do e-mail e acessa estas subpastas
def acessar_subpasta_email(caminho_pasta):
    """
    Acessa uma pasta ou subpasta no Outlook com base no caminho fornecido, suportando múltiplos níveis de pastas.
    
    Parâmetros:
        caminho_pasta (str): Caminho da pasta, por exemplo, "Pasta/Subpasta/Subsubpasta".
        endereco_email (str): Nome da caixa de correio principal, por exemplo, "nome@example.com".
        
    Retorna:
        A pasta acessada, se encontrada. Caso contrário, retorna None.
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") # Inicializando a conexão com o Outlook
    
    enderecos_email = []
    for pastas in outlook.Folders:
        enderecos_email.append(pastas)
    endereco_email = str(enderecos_email[1])
    # Acessa a caixa de correio principal
    try:
        caixa_correio = outlook.Folders.Item(endereco_email)
    except Exception as e:
        print(f"Erro ao acessar a caixa de correio '{endereco_email}': {e}")
        return None

    pastas = caminho_pasta
    # Navega por cada pasta na hierarquia
    pasta_atual = caixa_correio
    try:
        for pasta in pastas:
            pasta_atual = pasta_atual.Folders.Item(pasta)
        print(f"Pasta '{pasta_atual.Name}' foi acessada com sucesso.")
        return pasta_atual
    except Exception as e:
        print(f"Erro ao acessar a pasta '{pasta}': {e}")
        return None

def acessar_subpasta_outlook(caminho_pasta):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") # Inicializando a conexão com o Outlook

    # Acessar a pasta de entrada. Mais informações: https://learn.microsoft.com/pt-br/office/vba/api/outlook.oldefaultfolders
    inbox = outlook.GetDefaultFolder(6)  # 6 é a pasta de entrada (inbox)
    
    del caminho_pasta[0] #"Caixa de Entrada" (ou "Inbox") é o primeiro elemento da lista, portanto é removido pois sua implementaçõa é diferente)

    pasta_atual = inbox
    try:
        for pasta in caminho_pasta:
            pasta_atual = pasta_atual.Folders.Item(pasta)
        print(f"Pasta '{pasta_atual.Name}' foi acessada com sucesso.")
        return pasta_atual
    except Exception as e:
        print(f"Erro ao acessar a pasta '{pasta}': {e}")
        return None

def acessar_subpasta(caminho_pasta):
    # Divide o caminho em pastas e subpastas
    pastas = caminho_pasta.split("/")
    #Checar se caminho pasta é uma lista/ contem "Inbox" ou "Caixa de Entrada":
    if pastas[0]==("Inbox" or "Caixa de Entrada"):
        pasta = acessar_subpasta_outlook(pastas)
        return pasta
    else:
        pasta = acessar_subpasta_email(pastas)
        return pasta      
