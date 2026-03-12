import customtkinter as ctk        # Biblioteca principal para criar a interface gráfica moderna
from tkinter import filedialog     # Módulo nativo do Windows para abrir as janelas de "Salvar" e "Abrir" arquivo
from PIL import Image, ImageTk     # Pillow (PIL): Usada para manipular e exibir as imagens na tela
import fitz                        # PyMuPDF: O "motor" super rápido que lê, renderiza e extrai dados do PDF
import os                          # Para lidar com caminhos de arquivos e pastas no sistema operacional
import ctypes                      # Permite conversar diretamente com o Windows (usado para o ícone da barra de tarefas)
import pandas as pd                # A poderosa biblioteca de dados, usada aqui para estruturar e salvar as tabelas em CSV
import time                        # Usado para contar o tempo de execução das tarefas e gerar o log com a hora exata
import webbrowser                  # Permite abrir links diretamente no navegador de internet padrão do usuário

# --- CONFIGURAÇÕES GERAIS DE TEMA ---
# Define que o programa vai abrir no modo escuro e usar a cor azul como padrão para os botões
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class MeuLeitorPDF(ctk.CTk):
    """
    Classe principal do aplicativo. Ela herda tudo do CustomTkinter (ctk.CTk),
    ou seja, a própria classe 'é' a janela principal do programa.
    """
    def __init__(self):
        super().__init__() # Inicializa a janela mãe

        # --- NOME E TÍTULO ---
        self.title("Leitor PDF By SrBozzo") # Nome que aparece no topo da janela
        self.geometry("1150x750")           # Tamanho inicial da janela (Largura x Altura)

        # --- TRUQUE PARA A BARRA DE TAREFAS DO WINDOWS ---
        # O Windows costuma agrupar programas Python no ícone padrão.
        # Isso cria um ID único para o nosso app, forçando o Windows a tratá-lo como um software independente.
        try:
            meu_id_app = 'srbozzo.leitor_pdf.1_0'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(meu_id_app)
        except:
            pass # Se o usuário não estiver no Windows (ex: Mac ou Linux), ignora silenciosamente

        # --- ÍCONE DEFINITIVO (.ico) ---
        # Descobre exatamente onde este arquivo Python está salvo e busca o 'logo.ico' na mesma pasta.
        try:
            pasta_do_script = os.path.dirname(os.path.abspath(__file__))
            caminho_icone = os.path.join(pasta_do_script, "logo.ico") 
            self.iconbitmap(caminho_icone) # Aplica o ícone na janela
        except Exception as e:
            print(f"Erro ao carregar o ícone: {e}")

        # --- VARIÁVEIS DE CONTROLE (O "Cérebro" do Estado do App) ---
        self.caminho_atual = ""        # Guarda o caminho do arquivo PDF aberto no momento
        self.doc_pdf = None            # Guarda o objeto do PDF lido pelo PyMuPDF (fitz)
        self.zoom_padrao = 1.5         # Nível de zoom inicial
        self.zoom_atual = self.zoom_padrao # Nível de zoom que está sendo aplicado no momento
        self.pagina_atual = 0          # Índice da página que estamos vendo (0 é a primeira página)
        self.total_paginas = 0         # Quantidade total de páginas do PDF aberto
        self.imagem_tk = None          # Guarda a imagem renderizada da página para o Tkinter exibir
        self.id_imagem_canvas = None   # Guarda o ID do objeto de imagem dentro do Canvas (tela)
        self.drag_x = 0                # Coordenada X inicial ao clicar e arrastar a página
        self.drag_y = 0                # Coordenada Y inicial ao clicar e arrastar a página
        self.termo_pesquisa = ""       # Guarda a palavra que o usuário buscou (Ctrl+F)
        self.primeiro_log = True       # Flag para saber se apagamos a introdução e começamos o log

        # --- ESTRUTURA DA TELA (Divisão de Layout) ---
        # Criamos um frame (container) à esquerda para o PDF (ocupa a maior parte da tela)
        self.frame_esq = ctk.CTkFrame(self, width=650, corner_radius=0)
        self.frame_esq.pack(side="left", fill="both", expand=True)

        # Criamos um frame à direita para os botões e ferramentas (largura fixa)
        self.frame_dir = ctk.CTkFrame(self, width=400, corner_radius=0)
        self.frame_dir.pack(side="right", fill="both", expand=False, padx=10, pady=10)

        # --- LADO ESQUERDO: CANVAS (A tela de pintura do PDF) ---
        # O Canvas é onde a imagem gerada pelo PyMuPDF será efetivamente desenhada/colada.
        self.canvas = ctk.CTkCanvas(self.frame_esq, bg="#2b2b2b", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        # "Bindings" - Ligando eventos do mouse e teclado às nossas funções
        self.canvas.bind("<Control-ButtonPress-1>", self.iniciar_arrasto) # Ctrl + Clique esquerdo
        self.canvas.bind("<Control-B1-Motion>", self.arrastar_imagem)     # Mover o mouse segurando o botão
        self.bind("<Up>", self.pagina_anterior)                           # Seta para cima
        self.bind("<Down>", self.proxima_pagina)                          # Seta para baixo
        self.bind("<Control-MouseWheel>", self.zoom_mouse)                # Ctrl + Rodinha do mouse
        self.canvas.bind("<MouseWheel>", self.scroll_mouse)               # Rodinha do mouse (scroll simples)

        # --- MENU FLUTUANTE DE CONTROLE (Abaixo do Canvas) ---
        # Cria uma "pílula" flutuante que abriga os botões de navegação e zoom
        self.menu_controles = ctk.CTkFrame(self.frame_esq, corner_radius=20, fg_color="#1f1f1f", bg_color="#2b2b2b")
        self.menu_controles.place(relx=0.5, rely=0.95, anchor="s") # Centralizado embaixo

        # Botão: Página anterior
        self.btn_ant = ctk.CTkButton(self.menu_controles, text="<", width=30, command=self.pagina_anterior)
        self.btn_ant.pack(side="left", padx=5, pady=5)

        # Texto informando a página atual (ex: 1 / 10)
        self.lbl_paginas = ctk.CTkLabel(self.menu_controles, text="0 / 0", width=60)
        self.lbl_paginas.pack(side="left", padx=5, pady=5)

        # Botão: Próxima página
        self.btn_prox = ctk.CTkButton(self.menu_controles, text=">", width=30, command=self.proxima_pagina)
        self.btn_prox.pack(side="left", padx=5, pady=5)

        # Separador visual (|)
        self.lbl_separador = ctk.CTkLabel(self.menu_controles, text="|")
        self.lbl_separador.pack(side="left", padx=10, pady=5)

        # Botão: Menos Zoom
        self.btn_menos = ctk.CTkButton(self.menu_controles, text="-", width=30, command=self.diminuir_zoom)
        self.btn_menos.pack(side="left", padx=5, pady=5)

        # Botão: Resetar Zoom (100%)
        self.btn_reset = ctk.CTkButton(self.menu_controles, text="100%", width=50, command=self.resetar_zoom)
        self.btn_reset.pack(side="left", padx=5, pady=5)

        # Botão: Mais Zoom
        self.btn_mais = ctk.CTkButton(self.menu_controles, text="+", width=30, command=self.aumentar_zoom)
        self.btn_mais.pack(side="left", padx=5, pady=5)

        # ==========================================
        # --- LADO DIREITO: BARRA DE FERRAMENTAS ---
        # ==========================================

        # Switch de Tema (Modo Claro/Escuro) posicionado no topo à direita
        self.frame_topo_dir = ctk.CTkFrame(self.frame_dir, fg_color="transparent")
        self.frame_topo_dir.pack(fill="x", pady=(0, 10))
        
        self.switch_tema = ctk.CTkSwitch(self.frame_topo_dir, text="Modo Claro", command=self.alternar_tema)
        self.switch_tema.pack(side="right")

        # --- SEÇÃO DE PESQUISA (Ctrl+F) ---
        self.frame_pesquisa = ctk.CTkFrame(self.frame_dir, fg_color="transparent")
        self.frame_pesquisa.pack(fill="x", pady=(0, 15))
        
        # Caixa para digitar a palavra
        self.entrada_pesquisa = ctk.CTkEntry(self.frame_pesquisa, placeholder_text="Pesquisar (Ctrl+F)")
        self.entrada_pesquisa.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # Botão da lupa
        self.btn_pesquisar = ctk.CTkButton(self.frame_pesquisa, text="🔍", width=40, command=self.realizar_pesquisa)
        self.btn_pesquisar.pack(side="right")
        
        # Atalhos de teclado para a pesquisa
        self.bind("<Control-f>", lambda event: self.entrada_pesquisa.focus()) # Joga o cursor piscando para dentro da caixa
        self.entrada_pesquisa.bind("<Return>", lambda event: self.realizar_pesquisa()) # Aperta Enter para pesquisar

        # Botão Principal (Abre o arquivo PDF)
        self.botao_abrir = ctk.CTkButton(self.frame_dir, text="📂 Abrir PDF", command=self.abrir_arquivo, height=40)
        self.botao_abrir.pack(pady=(5, 15), fill="x")

        # --- SEÇÃO 1: CONVERSORES DE ARQUIVO ---
        self.lbl_titulo_conversao = ctk.CTkLabel(self.frame_dir, text="🔄 Converter Arquivo", font=("Arial", 14, "bold"))
        self.lbl_titulo_conversao.pack(pady=(5, 0))

        self.frame_conversores = ctk.CTkFrame(self.frame_dir, fg_color="transparent")
        self.frame_conversores.pack(pady=0) 
        
        self.btn_txt = ctk.CTkButton(self.frame_conversores, text="📄 Para .TXT", width=160, command=self.salvar_txt)
        self.btn_txt.pack(side="left", padx=5, pady=5)
        
        self.btn_word = ctk.CTkButton(self.frame_conversores, text="📘 Para Word", width=160, command=self.salvar_word)
        self.btn_word.pack(side="right", padx=5, pady=5)

        # --- SEÇÃO 2: EXTRATORES AVANÇADOS ---
        self.lbl_titulo_extratores = ctk.CTkLabel(self.frame_dir, text="✂️ Extração Avançada", font=("Arial", 14, "bold"))
        self.lbl_titulo_extratores.pack(pady=(15, 0))

        self.frame_extratores = ctk.CTkFrame(self.frame_dir, fg_color="transparent")
        self.frame_extratores.pack(pady=0)
        
        self.btn_img = ctk.CTkButton(self.frame_extratores, text="🖼️ Imagens", width=160, command=self.extrair_imagens)
        self.btn_img.pack(side="left", padx=5, pady=5)
        
        self.btn_tab = ctk.CTkButton(self.frame_extratores, text="📊 Tabelas (CSV)", width=160, command=self.extrair_tabelas)
        self.btn_tab.pack(side="right", padx=5, pady=5)

        # --- CAIXA DE TEXTO / LOG DE SESSÃO ---
        # Esta caixa começa com a introdução e depois vira o nosso monitor de registros (Logs)
        self.caixa_texto = ctk.CTkTextbox(self.frame_dir, width=350, font=("Consolas", 12)) 
        self.caixa_texto.pack(pady=(20, 5), fill="both", expand=True) 
        
        # --- RODAPÉ FIXO COM HIPERLIGAÇÃO ---
        self.lbl_github = ctk.CTkLabel(self.frame_dir, text="Feito por SrBozzo", font=("Arial", 12, "underline"), text_color="#1f6aa5", cursor="hand2")
        self.lbl_github.pack(pady=(0, 5))
        # Abre o site do GitHub ao clicar usando a biblioteca webbrowser
        self.lbl_github.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/SrBozzo"))

        # Introdução inicial que será exibida ao abrir o programa
        introducao = (
            "==================================\n"
            "   LEITOR PDF BY SRBOZZO v1.0\n"
            "==================================\n\n"
            "📌 INSTRUÇÕES BÁSICAS:\n"
            "> Scroll do Rato: Navega na página e \n  passa automaticamente de folha.\n"
            "> Prima Ctrl+F para pesquisar no texto.\n"
            "> Arraste a página com Ctrl + Clique.\n"
            "> Faça zoom com Ctrl + Scroll.\n\n"
            "[SISTEMA PRONTO. A aguardar ação...]\n"
        )
        self.caixa_texto.insert("1.0", introducao)
        # Bloqueia a caixa de texto para que o usuário não consiga digitar nela
        self.caixa_texto.configure(state="disabled") 

    # ==========================================
    # --- SISTEMA CENTRAL DE LOGS ---
    # ==========================================
    def adicionar_log(self, mensagem):
        """
        Adiciona uma mensagem ao monitor de atividades no painel direito, 
        com a hora exata da ocorrência.
        """
        self.caixa_texto.configure(state="normal") # Destranca a caixa para o Python poder escrever
        
        # Se for o primeiro registro, limpa a introdução antes de escrever
        if self.primeiro_log:
            self.caixa_texto.delete("1.0", "end")
            self.primeiro_log = False
            
        hora_atual = time.strftime("%H:%M:%S") # Pega a hora atual (ex: 14:30:05)
        texto_log = f"[{hora_atual}] {mensagem}\n"
        
        self.caixa_texto.insert("end", texto_log) # Insere o texto no final do log
        self.caixa_texto.see("end")               # Força a barra de rolagem a descer (acompanhar o texto novo)
        
        self.caixa_texto.configure(state="disabled") # Tranca a caixa novamente
        self.update() # Força a tela a se atualizar imediatamente para mostrar o log

    # ==========================================
    # --- LÓGICA DE TEMA (CLARO/ESCURO) ---
    # ==========================================
    def alternar_tema(self):
        """Alterna as cores do sistema entre modo Claro e Escuro com base no switch."""
        modo = "Claro" if self.switch_tema.get() == 1 else "Escuro"
        
        if self.switch_tema.get() == 1: # Se o switch foi ativado (Modo Claro)
            ctk.set_appearance_mode("light")
            self.canvas.configure(bg="#e0e0e0") # Muda o fundo da tela para cinza claro
            self.menu_controles.configure(fg_color="#d1d1d1", bg_color="#e0e0e0")
        else: # Se o switch foi desativado (Modo Escuro)
            ctk.set_appearance_mode("dark")
            self.canvas.configure(bg="#2b2b2b") # Volta para o cinza escuro
            self.menu_controles.configure(fg_color="#1f1f1f", bg_color="#2b2b2b")
            
        self.adicionar_log(f"Tema alterado para Modo {modo}.")

    # ==========================================
    # --- LÓGICA DE PESQUISA (Ctrl + F) ---
    # ==========================================
    def realizar_pesquisa(self):
        """Busca o termo digitado e manda a tela redesenhar a página para aplicar os quadrados vermelhos."""
        self.termo_pesquisa = self.entrada_pesquisa.get().strip() # Pega o texto e remove espaços sobrando
        self.atualizar_imagem_tela() # Manda renderizar o PDF novamente para os quadrados vermelhos aparecerem
        
        if self.termo_pesquisa and self.doc_pdf:
            self.adicionar_log(f"Pesquisa: Realizada procura pelo termo '{self.termo_pesquisa}'.")

    # ==========================================
    # --- LÓGICA DE EXTRAÇÃO AVANÇADA ---
    # ==========================================
    def extrair_imagens(self):
        """Varre o PDF inteiro, encontra imagens e salva na pasta escolhida."""
        if not self.doc_pdf:
            self.adicionar_log("Aviso: Nenhum PDF aberto para extrair imagens.")
            return
            
        # Pede para o usuário escolher em qual pasta ele quer salvar as imagens
        pasta_destino = filedialog.askdirectory(title="Selecione a pasta")
        if not pasta_destino: return # Se o usuário cancelar, interrompe a função
        
        self.adicionar_log("Processo: A extrair imagens do PDF...")
        inicio_tempo = time.time() # Inicia o cronômetro

        contador = 0
        # Percorre página por página do PDF
        for i in range(len(self.doc_pdf)):
            # Pega a lista de todas as imagens embutidas na página atual
            for img_info in self.doc_pdf[i].get_images(full=True):
                try:
                    xref = img_info[0] # Pega a referência cruzada (o "ID" da imagem dentro do PDF)
                    img = fitz.Pixmap(self.doc_pdf, xref) # Transforma esse ID numa matriz de pixels
                    
                    # Correção de cor: Se a imagem estiver em CMYK (usado pra impressão gráfica), converte pra RGB
                    if img.n - img.alpha > 3: 
                        img = fitz.Pixmap(fitz.csRGB, img)
                        
                    # Salva a imagem gerada no HD do usuário
                    img.save(os.path.join(pasta_destino, f"imagem_pag{i+1}_{contador}.png"))
                    contador += 1
                except: pass # Se der erro em uma imagem específica (corrompida, por ex), ignora e vai pra próxima
                    
        tempo_total = time.time() - inicio_tempo # Para o cronômetro
        self.adicionar_log(f"Sucesso: {contador} imagens extraídas em {tempo_total:.2f}s.")

    def extrair_tabelas(self):
        """Varre o PDF inteiro buscando formatações de tabela e exporta para CSV via Pandas."""
        if not self.doc_pdf:
            self.adicionar_log("Aviso: Nenhum PDF aberto para extrair tabelas.")
            return

        # Pede onde salvar o arquivo CSV
        caminho_salvar = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Arquivo CSV (Excel)", "*.csv")])
        if not caminho_salvar: return

        self.adicionar_log("Processo: A procurar tabelas e a montar DataFrame...")
        inicio_tempo = time.time()

        lista_dataframes = []
        for pagina in self.doc_pdf:
            # O PyMuPDF (fitz) tenta identificar estruturas de grades e linhas que formam tabelas na página
            tabelas = pagina.find_tables()
            for tab in tabelas:
                # Converte a tabela crua num DataFrame estruturado do Pandas
                df = tab.to_pandas()
                lista_dataframes.append(df) # Adiciona na lista geral

        if lista_dataframes:
            # Junta (concatena) todas as tabelas menores num DataFrame gigante e organizado
            df_final = pd.concat(lista_dataframes, ignore_index=True)
            # Exporta para CSV (utf-8-sig garante que os acentos do português não fiquem bugados no Excel)
            df_final.to_csv(caminho_salvar, index=False, encoding='utf-8-sig')
            
            tempo_total = time.time() - inicio_tempo
            self.adicionar_log(f"Sucesso: Tabelas exportadas para CSV em {tempo_total:.2f}s.")
        else:
            self.adicionar_log("Aviso: Nenhuma tabela estruturada foi encontrada neste PDF.")

    # ==========================================
    # --- LÓGICA DE CONVERSÃO (.TXT E WORD) ---
    # ==========================================
    def salvar_txt(self):
        """Lê os textos de todas as páginas e salva num arquivo .txt."""
        if not self.caminho_atual: 
            self.adicionar_log("Aviso: Nenhum PDF aberto para conversão.")
            return
            
        caminho_salvar = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Arquivo de Texto", "*.txt")])
        if caminho_salvar:
            inicio_tempo = time.time()
            texto_full = ""
            for p in self.doc_pdf: texto_full += p.get_text() # Extrai texto de cada página e junta tudo
            
            with open(caminho_salvar, 'w', encoding='utf-8') as f: 
                f.write(texto_full) # Escreve no arquivo e fecha
                
            tempo_total = time.time() - inicio_tempo
            self.adicionar_log(f"Sucesso: Texto puro (.txt) gerado em {tempo_total:.2f}s.")

    def salvar_word(self):
        """Converte o PDF em Word mantendo layout e imagens usando pdf2docx."""
        if not self.caminho_atual: 
            self.adicionar_log("Aviso: Nenhum PDF aberto para conversão.")
            return
            
        try: from pdf2docx import Converter # Importação local para não travar o app se a lib faltar
        except ImportError: return

        caminho_salvar = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if caminho_salvar:
            self.adicionar_log("Processo: A iniciar o motor de conversão PDF -> Word...")
            inicio_tempo = time.time()
            
            # Inicializa o conversor e faz a transformação pesada (isso pode demorar dependendo do PDF)
            cv = Converter(self.caminho_atual)
            cv.convert(caminho_salvar) 
            cv.close()
            
            tempo_total = time.time() - inicio_tempo
            self.adicionar_log(f"Sucesso: Ficheiro Word gerado mantendo layout em {tempo_total:.2f}s.")

    # ==========================================
    # --- INTERAÇÃO COM O CANVAS E MOUSE ---
    # ==========================================
    def iniciar_arrasto(self, event):
        """Registra a coordenada inicial do mouse ao segurar o botão para arrastar a imagem."""
        self.drag_x = event.x
        self.drag_y = event.y

    def arrastar_imagem(self, event):
        """Move a imagem e os retângulos de pesquisa acompanhando o movimento do mouse."""
        if self.id_imagem_canvas is not None:
            # Calcula a diferença de movimento do mouse (delta X e delta Y)
            dx = event.x - self.drag_x
            dy = event.y - self.drag_y
            
            # Move fisicamente a imagem principal e os retângulos de marcação de texto vermelho
            self.canvas.move(self.id_imagem_canvas, dx, dy)
            self.canvas.move("pesquisa", dx, dy) 
            
            # Atualiza as novas posições iniciais para o próximo micro-movimento
            self.drag_x, self.drag_y = event.x, event.y

    def zoom_mouse(self, event):
        """Detecta o scroll do mouse segurando CTRL para aumentar/diminuir o zoom."""
        if event.delta > 0: self.aumentar_zoom()  # Rodou pra cima, aproxima
        elif event.delta < 0: self.diminuir_zoom() # Rodou pra baixo, afasta

    def scroll_mouse(self, event): 
        """Navega a página para cima/baixo com o mouse. Passa de página se chegar na borda."""
        if not self.doc_pdf or self.id_imagem_canvas is None:
            return

        # Pega as coordenadas exatas das bordas da imagem renderizada na tela (Caixa Delimitadora)
        bbox = self.canvas.bbox(self.id_imagem_canvas)
        if not bbox: return
        
        y1, y2 = bbox[1], bbox[3] # y1 = topo da imagem, y2 = base da imagem
        altura_canvas = self.canvas.winfo_height() # Altura da área cinza visível
        passo_scroll = 60 # Quantidade de pixels que a página desliza por 'clique' do scroll
        
        if event.delta > 0: # Scroll rodou para CIMA
            if y1 >= 0:
                # Se o topo da imagem (y1) já está visível (maior ou igual a 0), 
                # significa que chegamos ao topo. Mudamos para a página anterior.
                if self.pagina_atual > 0:
                    self.pagina_anterior()
            else:
                # Se ainda há imagem escondida em cima, apenas deslizamos a página para baixo
                self.canvas.move(self.id_imagem_canvas, 0, passo_scroll)
                self.canvas.move("pesquisa", 0, passo_scroll)
                
        elif event.delta < 0: # Scroll rodou para BAIXO
            if y2 <= altura_canvas:
                # Se a base da imagem (y2) já subiu mais que a altura visível,
                # significa que chegamos ao fim da folha. Mudamos para a próxima página.
                if self.pagina_atual < (self.total_paginas - 1):
                    self.proxima_pagina()
            else:
                # Deslizamos a página para cima para revelar o resto do conteúdo em baixo
                self.canvas.move(self.id_imagem_canvas, 0, -passo_scroll)
                self.canvas.move("pesquisa", 0, -passo_scroll)

    # ==========================================
    # --- RENDERIZAÇÃO DE PÁGINAS E ZOOM ---
    # ==========================================
    def pagina_anterior(self, event=None):
        if self.doc_pdf and self.pagina_atual > 0:
            self.pagina_atual -= 1
            self.atualizar_imagem_tela() # Redesenha a nova página

    def proxima_pagina(self, event=None):
        if self.doc_pdf and self.pagina_atual < (self.total_paginas - 1):
            self.pagina_atual += 1
            self.atualizar_imagem_tela() # Redesenha a nova página

    def atualizar_imagem_tela(self):
        """
        O coração visual do leitor: Carrega a página atual do PDF, aplica o zoom,
        converte para uma imagem aceitável pelo Tkinter e exibe no Canvas.
        Também é responsável por desenhar a busca (Ctrl+F) em cima das palavras.
        """
        if not self.doc_pdf: return
        
        # Atualiza a contagem ali no menu flutuante (ex: 3 / 10)
        self.lbl_paginas.configure(text=f"{self.pagina_atual + 1} / {self.total_paginas}")
        
        # Carrega a página atual na memória
        pagina = self.doc_pdf.load_page(self.pagina_atual)
        
        # Aplica a multiplicação/matriz para calcular o tamanho real pelo zoom escolhido
        matriz = fitz.Matrix(self.zoom_atual, self.zoom_atual)
        
        # Gera o Pixmap (um mapa de pixels cru da imagem gerada do documento)
        pix = pagina.get_pixmap(matrix=matriz)
        
        # Verifica se tem transparência ou não para montar o modo de cor correto
        modo = "RGBA" if pix.alpha else "RGB"
        
        # Pega os bytes de pixel brutos do PyMuPDF e os transforma numa imagem da biblioteca PIL
        img_pil = Image.frombytes(modo, [pix.width, pix.height], pix.samples)
        
        # Converte a imagem da PIL para um formato especial que a interface gráfica (Tkinter) entende
        self.imagem_tk = ImageTk.PhotoImage(img_pil)
        
        # Limpa o Canvas inteiro antes de colar a nova imagem (para não sobrepor nada antigo)
        self.canvas.delete("all") 
        
        # Calcula o centro exato da área cinza do app para colar o PDF no meio
        cx, cy = self.canvas.winfo_width() / 2, self.canvas.winfo_height() / 2
        # Se o canvas ainda não atualizou as medidas e devolver número quebrado/pequeno, usa um fixo padrão
        if cx < 10: cx, cy = 325, 350 
        
        # Efetivamente cola a imagem gerada no centro da tela
        self.id_imagem_canvas = self.canvas.create_image(cx, cy, anchor="center", image=self.imagem_tk)

        # --- LÓGICA DE DESENHO DA PESQUISA (QUADRADOS VERMELHOS) ---
        if self.termo_pesquisa:
            # O PyMuPDF caça a palavra no texto raiz do PDF e nos dá uma lista de retângulos matemáticos
            areas = pagina.search_for(self.termo_pesquisa)
            
            # Descobre onde é o topo esquerdo exato da imagem colada no canvas
            topo_x = cx - (img_pil.width / 2)
            topo_y = cy - (img_pil.height / 2)
            
            # Pega as coordenadas do papel original e ajusta multiplicando pelo zoom da tela,
            # desenhando os retângulos vermelhos com a tag 'pesquisa' para podermos movê-los depois.
            for retangulo in areas:
                x0 = topo_x + (retangulo.x0 * self.zoom_atual)
                y0 = topo_y + (retangulo.y0 * self.zoom_atual)
                x1 = topo_x + (retangulo.x1 * self.zoom_atual)
                y1 = topo_y + (retangulo.y1 * self.zoom_atual)
                
                self.canvas.create_rectangle(x0, y0, x1, y1, outline="red", width=2, tags="pesquisa")

    def aumentar_zoom(self, event=None):
        if self.doc_pdf: 
            self.zoom_atual += 0.3
            self.atualizar_imagem_tela()

    def diminuir_zoom(self, event=None):
        if self.doc_pdf and self.zoom_atual > 0.5: 
            self.zoom_atual -= 0.3
            self.atualizar_imagem_tela()

    def resetar_zoom(self, event=None):
        if self.doc_pdf: 
            self.zoom_atual = self.zoom_padrao
            self.atualizar_imagem_tela()

    # ==========================================
    # --- ABERTURA DE ARQUIVO ---
    # ==========================================
    def abrir_arquivo(self):
        """Abre o seletor do Windows para carregar um novo arquivo PDF na memória."""
        self.caminho_atual = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if self.caminho_atual:
            inicio_tempo = time.time()
            
            # Abre efetivamente o motor fitz no arquivo selecionado
            self.doc_pdf = fitz.open(self.caminho_atual)
            self.total_paginas = len(self.doc_pdf)
            self.pagina_atual = 0  
            self.zoom_atual = self.zoom_padrao 
            
            # Limpa qualquer pesquisa antiga
            self.termo_pesquisa = ""
            self.entrada_pesquisa.delete(0, 'end')
            
            # Força a interface a acordar/atualizar as dimensões do Canvas e renderiza a primeira página
            self.update_idletasks() 
            self.atualizar_imagem_tela()
            
            # Adiciona o registro no monitor de Logs com o nome do arquivo
            nome_arquivo = os.path.basename(self.caminho_atual)
            tempo_total = time.time() - inicio_tempo
            self.adicionar_log(f"Sistema: Ficheiro '{nome_arquivo}' carregado em {tempo_total:.3f}s. ({self.total_paginas} páginas)")

# --- INICIALIZAÇÃO DO APLICATIVO ---
# Se este script for o arquivo principal rodado pelo Python, ele instancia a classe e mantém o loop rodando.
if __name__ == "__main__":
    app = MeuLeitorPDF() # Cria a aplicação a partir do molde/classe que programamos
    app.mainloop()       # O 'mainloop' escuta os eventos do mouse/teclado para não fechar o app sozinho