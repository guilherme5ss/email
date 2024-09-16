from emails import *

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") # Inicializando a conexão com o Outlook

caminho =(get_oulook_trees()[1]).find_path_values("Respostas")

# Obter todos os e-mails na pasta
messages = acessar_subpasta(caminho).Items

# Lista para armazenar os dados dos links
data_list = []

# Expressão regular para encontrar URLs
url_regex = re.compile(r'https?://\S+')

# Palavras âncoras e domínios irrelevantes
irrelevant_anchors = ["aqui", "clique", "cancelar", "unsubscribe", "sair"]
advertising_patterns = ["promo", "marketing", "campaign", "upsell", "st_appsite_flagship"]

# Domínios irrelevantes (exemplos de domínios comuns em propagandas)
irrelevant_domains = ["mailchimp.com", "ad.doubleclick.net"]

# Função para verificar se um link é irrelevante baseado na âncora
def is_relevant_link(anchor_text, link_url):
    # Verificar se o texto âncora contém palavras irrelevantes
    for word in irrelevant_anchors:
        if word.lower() in anchor_text.lower():
            return False
    
    # Verificar se o link contém padrões de propaganda
    for pattern in advertising_patterns:
        if pattern.lower() in link_url.lower():
            return False

    # Verificar se o domínio do link é irrelevante
    for domain in irrelevant_domains:
        if domain.lower() in link_url.lower():
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

                # Filtrar links irrelevantes e de propaganda
                if is_relevant_link(anchor_text, link):
                    subject = message.Subject
                    date = message.ReceivedTime.strftime("%Y-%m-%d")  # Data do e-mail
                    time = message.ReceivedTime.strftime("%H:%M:%S")  # Hora do e-mail
                    
                    # Adicionar os dados à lista
                    data_list.append([subject, link, date, time])
    except Exception as e:
        print(f"Erro ao processar o e-mail: {e}")

# Criando um DataFrame com os dados coletados
df = pd.DataFrame(data_list, columns=["Assunto", "Link", "Data", "Hora"])

# Salvando o DataFrame em um arquivo Excel
df.to_excel("links_filtrados_outlook.xlsx", index=False)

print("Arquivo Excel criado com sucesso!")
