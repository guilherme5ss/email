from emails import *

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") # Inicializando a conexão com o Outlook

# Acessando "Inbox". Mais informações: https://learn.microsoft.com/pt-br/office/vba/api/outlook.oldefaultfolders
inbox = outlook.GetDefaultFolder(6)  # 6 é o código para a pasta "Inbox"

# Obter todos os e-mails na pasta
messages = inbox.Items

ultima_mensagem = messages.GetLast()
print(ultima_mensagem)

# Dicionário para armazenar links e as mensagens correspondentes
links_dict = defaultdict(list)

# Expressão regular para encontrar URLs
url_regex = re.compile(r'https?://\S+')

# Percorrer todas as mensagens
for message in messages:
    try:
        body = message.Body  # Conteúdo do e-mail
        links = url_regex.findall(body)  # Encontrar todos os links no corpo do e-mail
        for link in links:
            links_dict[link].append(message.Subject)  # Armazenar o assunto da mensagem com o link
    except Exception as e:
        print(f"Erro ao processar o e-mail: {e}")

# Exibindo os resultados
for link, subjects in links_dict.items():
    print(f"Link: {link}")
    print("Encontrado em:")
    for subject in subjects:
        print(f"- {subject}")
    print("\n")