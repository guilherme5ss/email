from emails import *

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") # Inicializando a conexão com o Outlook

pasta = "Respostas"
caminho =(get_oulook_trees()[1]).find_path_values(pasta)

# Obter todos os e-mails na pasta
messages = acessar_subpasta(caminho).Items

# Dicionário para armazenar links e as mensagens correspondentes
links_dict = defaultdict(list)

# Lista para armazenar informações de todos os e-mails lidos
emails_lidos = []

# Lista para armazenar links considerados descartáveis (irrelevantes)
links_descartados = []

# Expressão regular para encontrar URLs
url_regex = re.compile(r'https?://\S+')

# Palavras âncoras e domínios irrelevantes
irrelevant_anchors = ["aqui", "clique", "cancelar", "unsubscribe", "sair", "ajuda", "saiba por que incluímos isso.", "conte como foi sua experiência"]
advertising_patterns = ["unsubscribe", "promo", "marketing", "upsell", "st_appsite_flagship", "home_glimmer", "logoGlimmer", "profile_glimmer"]

# Domínios irrelevantes (exemplos de domínios comuns em propagandas)
irrelevant_domains = ["mailchimp.com", "ad.doubleclick.net"]

# Função para verificar se um link é relevante e identificar o motivo de descarte
def check_link_relevance(anchor_text, link_url):
    # Verificar se o texto âncora contém palavras irrelevantes
    for word in irrelevant_anchors:
        if word.lower() in anchor_text.lower():
            return False, "Texto âncora irrelevante"
    
    # Verificar se o link contém padrões de propaganda
    for pattern in advertising_patterns:
        if pattern.lower() in link_url.lower():
            return False, "Link de propaganda"
    
    # Verificar se o domínio do link é irrelevante
    for domain in irrelevant_domains:
        if domain.lower() in link_url.lower():
            return False, "Domínio irrelevante"
    
    return True, ""

# Percorrer todas as mensagens
for message in messages:
    try:
        # Coletar informações de cada e-mail
        subject = message.Subject
        date = message.ReceivedTime.strftime("%Y-%m-%d")  # Data do e-mail
        time = message.ReceivedTime.strftime("%H:%M:%S")  # Hora do e-mail
        entry_id = message.EntryID
        store_id = message.Parent.StoreID

        # Armazenar as informações do e-mail para a aba "Emails Lidos"
        emails_lidos.append([subject, date, time, entry_id, store_id])

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

                # Verificar se o link é relevante ou descartado
                is_relevant, reason = check_link_relevance(anchor_text, link)
                if is_relevant and link not in unique_links:
                    unique_links.add(link)  # Adicionar o link ao conjunto para evitar duplicações
                    
                    # Adicionar o link ao dicionário com as informações do e-mail
                    links_dict[link].append({"Assunto": subject, "Data": date, "Hora": time})
                elif not is_relevant and link not in unique_links:
                    # Adicionar os links descartados à lista com o motivo de descarte
                    links_descartados.append([link, subject, date, reason])
                    unique_links.add(link)  # Evitar duplicação
    except Exception as e:
        print(f"Erro ao processar o e-mail: {e}")

# Preparando os dados para exportação

# Lista para armazenar os dados de links com repetições e os e-mails associados
links_data_list = []

# Para cada link, listamos os detalhes e indicamos as repetições
for link, emails in links_dict.items():
    # Organizar a lista de e-mails pela data e hora para obter o primeiro e último
    emails_sorted = sorted(emails, key=lambda x: (x['Data'], x['Hora']))

    # Primeiro e-mail associado ao link
    first_email = emails_sorted[0]
    
    # Último e-mail associado (se houver mais de um)
    last_email = emails_sorted[-1] if len(emails_sorted) > 1 else {"Assunto": "", "Data": ""}
    
    # Adicionar a entrada à lista de links
    links_data_list.append([
        link, 
        first_email["Assunto"], first_email["Data"],  # Primeiro e-mail associado
        last_email["Assunto"] if last_email["Assunto"] != first_email["Assunto"] else "",  # Ultimo Email
        last_email["Data"] if last_email["Data"] != first_email["Data"] else "",  # Ultima Data
        len(emails)  # Número de repetições
    ])

# Criando DataFrames com os dados coletados
df_links = pd.DataFrame(links_data_list, columns=["Link", "Primeiro Email", "Primeira Data", "Ultimo Email", "Ultima Data", "Repetições"])
df_links_descartados = pd.DataFrame(links_descartados, columns=["Link", "Assunto", "Data", "Motivo Descarte"])
df_emails_lidos = pd.DataFrame(emails_lidos, columns=["Assunto", "Data", "Hora", "EntryID", "StoreID"])

saida = str(pasta)+"_links.xlsx"


# Função para ajustar a largura das colunas
def ajustar_colunas(worksheet, colunas_largura):
    for col, largura in colunas_largura.items():
        worksheet.set_column(col, largura)

# Definir os tamanhos desejados para cada coluna
colunas_largura = {
    'A:A': 30,   # Coluna A (Link)
    'B:B': 60,   # Coluna B (Primeiro Email)
    'C:C': 12.5, # Coluna C (Primeira Data)
    'D:D': 12.5, # Coluna D (Último Email)
    'E:E': 12.5, # Coluna E (Última Data)
    'F:F': 10.14 # Coluna F (Repetições)
}

# Criar sumarização para assuntos
with pd.ExcelWriter(saida) as writer: # Salvando os DataFrames em um arquivo Excel com três abas
    df_links.to_excel(writer, sheet_name="Links", index=False)
    df_links_descartados.to_excel(writer, sheet_name="Links Descartados", index=False)
    df_emails_lidos.to_excel(writer, sheet_name="Emails Lidos", index=False)

    # Aplicar o ajuste de colunas na aba "Links"
    worksheet_links = writer.sheets['Links']
    ajustar_colunas(worksheet_links, colunas_largura)
    
    # Aplicar o ajuste de colunas na aba "Links Descartados"
    worksheet_links_descartados = writer.sheets['Links Descartados']
    ajustar_colunas(worksheet_links_descartados, colunas_largura)

print("Arquivo Excel com abas 'Emails Lidos', 'Links' e 'Links Descartados' criado com sucesso!")
