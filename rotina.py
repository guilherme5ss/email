from emails import *

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") # Inicializando a conexão com o Outlook

# Acessando "Inbox". Mais informações: https://learn.microsoft.com/pt-br/office/vba/api/outlook.oldefaultfolders
inbox = outlook.GetDefaultFolder(6)  # 6 é o código para a pasta "Inbox"

# Obter todos os e-mails na pasta
messages = inbox.Items

# Dicionário para armazenar links e as mensagens correspondentes
links_dict = defaultdict(list)

# Expressão regular para encontrar URLs
url_regex = re.compile(r'https?://\S+')

# Palavras âncoras irrelevantes
irrelevant_anchors = ["aqui", "clique", "cancelar", "unsubscribe", "sair"]

# Função para verificar se um link é irrelevante baseado na âncora
def is_relevant_link(anchor_text):
    for word in irrelevant_anchors:
        if word.lower() in anchor_text.lower():
            return False
    return True

# Percorrer todas as mensagens
for message in messages:
    try:
        # Verificar se o e-mail tem HTMLBody (alguns e-mails podem ser texto simples)
        if message.BodyFormat == 2:  # 2 = HTML format
            html_body = message.HTMLBody
            soup = BeautifulSoup(html_body, "html.parser")  # Usando BeautifulSoup para parsear o HTML

            # Encontrar todos os links (<a> tags)
            for link_tag in soup.find_all('a', href=True):
                link = link_tag['href']  # Extrair o href (link)
                anchor_text = link_tag.get_text()  # Extrair o texto âncora
                
                # Filtrar links irrelevantes
                if is_relevant_link(anchor_text):
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