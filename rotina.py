from emails import *

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") # Inicializando a conexão com o Outlook

caminho =(get_oulook_trees()[1]).find_path_values("Respostas")

# Obter todos os e-mails na pasta
messages = acessar_subpasta(caminho).Items

# Dicionário para armazenar links e as mensagens correspondentes
links_dict = defaultdict(list)

# Expressão regular para encontrar URLs
url_regex = re.compile(r'https?://\S+')

# Palavras âncoras e domínios irrelevantes
irrelevant_anchors = ["aqui", "clique", "cancelar", "unsubscribe", "sair", "ajuda", "saiba por que incluímos isso.", "conte como foi sua experiência"]
advertising_patterns = ["unsubscribe", "promo", "marketing", "campaign", "upsell", "st_appsite_flagship", "home_glimmer", "logoGlimmer", "profile_glimmer"]

# Domínios irrelevantes (exemplos de domínios comuns em propagandas)
irrelevant_domains = ["mailchimp.com", "ad.doubleclick.net"]

# Função para verificar se um link é irrelevante baseado na âncora e URL
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
            
            # Usar um conjunto (set) para evitar duplicação de links no mesmo e-mail
            unique_links = set()

            # Encontrar todos os links (<a> tags)
            for link_tag in soup.find_all('a', href=True):
                link = link_tag['href']  # Extrair o href (link)
                anchor_text = link_tag.get_text()  # Extrair o texto âncora

                # Filtrar links irrelevantes e de propaganda
                if is_relevant_link(anchor_text, link) and link not in unique_links:
                    unique_links.add(link)  # Adicionar o link ao conjunto para evitar duplicações
                    
                    # Armazenar as informações do e-mail
                    subject = message.Subject
                    date = message.ReceivedTime.strftime("%Y-%m-%d")  # Data do e-mail
                    time = message.ReceivedTime.strftime("%H:%M:%S")  # Hora do e-mail
                    
                    # Adicionar o link ao dicionário com as informações do e-mail
                    links_dict[link].append({"Assunto": subject, "Data": date, "Hora": time})
    except Exception as e:
        print(f"Erro ao processar o e-mail: {e}")

# Preparando os dados para exportação
data_list = []

# Para cada link, listamos os detalhes e indicamos as repetições
for link, emails in links_dict.items():
    # Concatenar os detalhes de e-mails associados ao link
    email_details = "; ".join([f"{email['Assunto']} ({email['Data']} {email['Hora']})" for email in emails])
    
    # Adicionar a entrada à lista
    data_list.append([link, email_details, len(emails)])  # Inclui a contagem de repetições

# Criando um DataFrame com os dados coletados
df = pd.DataFrame(data_list, columns=["Link", "Emails Associados", "Repetições"])

# Salvando o DataFrame em um arquivo Excel
df.to_excel("links_com_repeticoes_outlook.xlsx", index=False)

print("Arquivo Excel criado com sucesso!")
