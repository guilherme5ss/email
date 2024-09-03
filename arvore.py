class Arvore(object): # Classe "Arvore". Inspirado em: https://stackoverflow.com/questions/41760856
    ''' # Exemplo simples de implementação              
raiz    = Arvore('Raiz')            
tronco  = raiz.add_child("Tronco")  
galho   = tronco.add_child('Galho') 
ramo    = galho.add_child('Ramo')   
print(raiz.display_arvore()) 
    '''
    def __init__(self, data, children=None, parent=None): # "Método especial, invocado quando um objeto é instanciado." 
        # Trecho da explicação retidado de "Pense em Python", uma tradução do livro Think Python (2ª edição), de Allen B. Downey
        self.data = data
        self.children = children or []
        self.parent = parent

    def add_child(self, data): # Adiciona um galho a uma parte da árvore
        new_child = Arvore(data, parent=self)
        self.children.append(new_child)
        return new_child

    def is_tree(self): # Método para conferir se faz parte de uma árvore
        return self.data is not None

    def is_root(self): # Método para conferir se parte da árvore é uma raiz, parte da árvore sem nó anterior
        return self.parent is None

    def is_branch(self): # Métoto confere se é um galho, se possui um nó precedente
        # Retorna True se este nó tiver pai, indicando que é um galho
        return self.parent is not None

    def is_leaf(self): # Indica se é o último elemento de um galho
        # Retorna True se este nó não tiver filhos, indicando que é uma folha
        return not self.children
    
    def __str__(self): # "É um método especial, como __init__, usado para retornar uma representação de string de um objeto."
        # Trecho da explicação retidado de "Pense em Python", uma tradução do livro Think Python (2ª edição), de Allen B. Downey. Link: https://t.ly/JMt66
        if self.is_leaf():
            return str(self.data)
        return '{data}/{children}'.format(data=self.data, children='/'.join(map(str, self.children)))
    
    def parents(self): # A partir de um galho busca-se sua raiz
        if self.is_root(): # Se for a raiz, retorna apenas o dado da raiz
            return str(self.data)
        
        parent_chain = [] # Lista para armazenar o caminho até a raiz
        node = self  # Começa do nó atual

        while node is not None: # Itera sobre os pais até alcançar a raiz
            parent_chain.append(str(node.data))
            node = node.parent  # Move para o pai do nó atual

        return '/'.join(reversed(parent_chain)) # Inverte a lista para mostrar o caminho da raiz até o nó atual
    
    def display_arvore(self, level=0, prefix="", is_root=True, is_last=False, marcadores=bool):# Exibe a árvore de maneira similar a uma estrura de arquivos (ou ficheiros)
        if marcadores == False:
            indent = "    " * level  # Define a indentação com base no nível da árvore
            result = indent + str(self.data) + "\n"
            for child in self.children:
                result += child.display_arvore(level + 1, marcadores = False)  # Recursivamente imprime os galhos com maior indentação
            return result
        else: # Se for a raiz, não adiciona o marcador
            if is_root:
                result = str(self.data) + "\n"
            else: # Adiciona o marcador para galhos, mas não para a raiz
                result = prefix + ("└── " if is_last else "├── ") + str(self.data) + "\n"
            
            # Atualiza o prefixo para os galhos
            if not is_last and not is_root:
                child_prefix = prefix + "│   "
            else:
                child_prefix = prefix + "    "
            
            # Percorre todos os galhos
            num_children = len(self.children)
            for i, child in enumerate(self.children):
                is_last_child = (i == num_children - 1)
                result += child.display_arvore(level + 1, child_prefix, False, is_last_child, marcadores=True)
            
            return result
        
    def find_path(self, data): # Procura o caminho até o nó que contém o dado especificado.
        """
        Procura o caminho até o nó que contém o dado especificado.
        Retorna uma lista com o caminho (nós) ou None se o dado não for encontrado.
        """
        if self.data == data:
            return [self]  # Se o dado estiver no nó atual, retorna ele como lista

        # Caso contrário, percorre os filhos para tentar encontrar o dado
        for child in self.children:
            path = child.find_path(data)
            if path:  # Se o dado foi encontrado em algum filho, retorna o caminho
                return [self] + path  # Adiciona o nó atual ao caminho

        return None  # Se não encontrou em nenhum nó

    def find_path_values(self, data): # Retorna o caminho até o nó que contém o dado especificado, mas apenas os valores dos nós.
        """
        Retorna o caminho até o nó que contém o dado especificado, mas apenas os valores dos nós.
        """
        path = self.find_path(data)
        if path:
            return [node.data for node in path]  # Retorna apenas os valores dos nós
        return None