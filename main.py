import os
import threading
import time
import customtkinter as ctk
from datetime import datetime, timedelta
from tkinter import messagebox, ttk
import sistemacofin as sc
from PIL import Image, ImageTk
import pandas as pd
from tkinter import filedialog
import shutil

# Configuração global do tema
ctk.set_appearance_mode("System")  # "System", "Dark" ou "Light"
ctk.set_default_color_theme("blue")

class AutocofinApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuração da janela principal
        self.title("Autocofin | Sistema de Automação COFIN")
        self.geometry("1280x720")  # Resolução mais ampla para melhor visualização
        self.minsize(1000, 700)
        
        # Cores personalizadas
        self.cor_bg = "#F0F2F5"          # Fundo geral (cinza bem claro)
        self.cor_accent = "#1976D2"      # Cor principal (azul)
        self.cor_text = "#212121"        # Texto (quase preto)
        self.cor_success = "#4CAF50"     # Verde para sucesso
        self.cor_warning = "#FFC107"     # Amarelo para avisos
        self.cor_error = "#F44336"       # Vermelho para erros
        
        # Configurar tema da janela
        self.configure(fg_color=self.cor_bg)
        
        # Layout da interface
        self.grid_columnconfigure(0, weight=2)  # Painel de controle (mais estreito)
        self.grid_columnconfigure(1, weight=5)  # Tabela de status (mais larga)
        self.grid_rowconfigure(0, weight=1)
        
        # Criar interfaces
        self.setup_painel_controle()
        self.setup_tabela_status()
        self.setup_barra_status()
        
        # Carregar os dados existentes
        self.carregar_dados_tabela()
        
        # Iniciar monitoramento
        self.update_status_periodico()
        
        # Verificar planilha
        self.verificar_planilha()
        
        # Registrar método para tratar o fechamento da janela
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_painel_controle(self):
        # Frame esquerdo com sombra e cantos arredondados
        self.frame_controles = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color="white",
            border_width=0
        )
        self.frame_controles.grid(row=0, column=0, sticky="nsew", padx=15, pady=15)
        self.frame_controles.grid_columnconfigure(0, weight=1)
        
        # Cabeçalho com logo e título
        self.frame_header = ctk.CTkFrame(
            self.frame_controles,
            fg_color="transparent",
            height=80
        )
        self.frame_header.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 0))
        self.frame_header.grid_columnconfigure(1, weight=1)
        
        # Título do aplicativo
        self.label_titulo = ctk.CTkLabel(
            self.frame_header, 
            text="AUTOCOFIN",
            font=ctk.CTkFont(size=26, weight="bold"),
            text_color=self.cor_accent
        )
        self.label_titulo.grid(row=0, column=0, sticky="w")
        
        self.label_subtitulo = ctk.CTkLabel(
            self.frame_header, 
            text="Sistema de Automação COFIN",
            font=ctk.CTkFont(size=14),
            text_color=self.cor_text
        )
        self.label_subtitulo.grid(row=1, column=0, sticky="w")
        
        # Separador com gradiente
        self.separador = ctk.CTkFrame(
            self.frame_controles, 
            height=2,
            fg_color=self.cor_accent
        )
        self.separador.grid(row=1, column=0, sticky="ew", padx=20, pady=15)
        
        # Frame principal para os controles
        self.frame_main = ctk.CTkFrame(
            self.frame_controles,
            fg_color="transparent"
        )
        self.frame_main.grid(row=2, column=0, sticky="nsew", padx=20, pady=10)
        self.frame_main.grid_columnconfigure(0, weight=1)
        
        # Status da planilha
        self.frame_status_planilha = ctk.CTkFrame(
            self.frame_main,
            fg_color="#E3F2FD",  # Azul bem claro
            corner_radius=8
        )
        self.frame_status_planilha.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        
        self.label_planilha = ctk.CTkLabel(
            self.frame_status_planilha,
            text="Verificando planilha...",
            font=ctk.CTkFont(size=14),
            padx=15, pady=10
        )
        self.label_planilha.pack(fill="both")
        
        # Opção da unidade
        self.label_opcao = ctk.CTkLabel(
            self.frame_main,
            text="Unidade:",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w"
        )
        self.label_opcao.grid(row=1, column=0, sticky="w", pady=(0, 5))
        
        self.opcoes_unidades = [
            "303017500 - CMB SEC DISTR",
            "303017400 - CMB SEC RECEB",
            "303016500 - CMB 3 - RESUMO"
        ]
        
        self.dropdown_opcao = ctk.CTkOptionMenu(
            self.frame_main,
            values=self.opcoes_unidades,
            font=ctk.CTkFont(size=14),
            dropdown_font=ctk.CTkFont(size=14),
            button_color=self.cor_accent,
            button_hover_color="#1565C0",  # Azul mais escuro no hover
            dropdown_hover_color="#E3F2FD"  # Azul claro para hover no dropdown
        )
        self.dropdown_opcao.grid(row=2, column=0, sticky="ew", pady=(0, 15))
        self.dropdown_opcao.set(self.opcoes_unidades[0])
        
        # Status atual - Card com fundo suave
        self.frame_status = ctk.CTkFrame(
            self.frame_main,
            fg_color="#FAFAFA",
            corner_radius=8,
            border_width=1,
            border_color="#E0E0E0"
        )
        self.frame_status.grid(row=3, column=0, sticky="ew", pady=(0, 15))
        self.frame_status.grid_columnconfigure(0, weight=1)
        
        self.label_status_titulo = ctk.CTkLabel(
            self.frame_status,
            text="STATUS DA OPERAÇÃO",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#757575"
        )
        self.label_status_titulo.grid(row=0, column=0, padx=15, pady=(10, 0))
        
        self.label_status = ctk.CTkLabel(
            self.frame_status,
            text="Aguardando",
            font=ctk.CTkFont(size=14),
            wraplength=280,
            justify="left"
        )
        self.label_status.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="w")
        
        # Progresso
        self.frame_progresso = ctk.CTkFrame(
            self.frame_main,
            fg_color="transparent"
        )
        self.frame_progresso.grid(row=4, column=0, sticky="ew", pady=(0, 15))
        self.frame_progresso.grid_columnconfigure(0, weight=1)
        self.frame_progresso.grid_columnconfigure(1, weight=1)
        
        self.label_progresso = ctk.CTkLabel(
            self.frame_progresso,
            text="Progresso: 0/0",
            font=ctk.CTkFont(size=14),
            anchor="w"
        )
        self.label_progresso.grid(row=0, column=0, sticky="w")
        
        self.label_tempo = ctk.CTkLabel(
            self.frame_progresso,
            text="Tempo: --:--",
            font=ctk.CTkFont(size=14),
            anchor="e"
        )
        self.label_tempo.grid(row=0, column=1, sticky="e")
        
        self.progress_bar = ctk.CTkProgressBar(
            self.frame_main,
            height=15,
            corner_radius=5,
            progress_color=self.cor_accent
        )
        self.progress_bar.grid(row=5, column=0, sticky="ew", pady=(0, 20))
        self.progress_bar.set(0)
        
        # Resultados em cards
        self.frame_cards = ctk.CTkFrame(
            self.frame_main,
            fg_color="transparent"
        )
        self.frame_cards.grid(row=6, column=0, sticky="ew", pady=(0, 15))
        self.frame_cards.grid_columnconfigure(0, weight=1)
        self.frame_cards.grid_columnconfigure(1, weight=1)
        
        # Card de sucesso
        self.card_sucesso = ctk.CTkFrame(
            self.frame_cards,
            fg_color="#E8F5E9",
            corner_radius=8,
            border_width=1,
            border_color="#C8E6C9"
        )
        self.card_sucesso.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        
        self.label_sucesso = ctk.CTkLabel(
            self.card_sucesso,
            text="✓ Sucesso: 0",
            font=ctk.CTkFont(size=14),
            text_color="#2E7D32",
            padx=10,
            pady=12
        )
        self.label_sucesso.pack()
        
        # Card de erro
        self.card_erro = ctk.CTkFrame(
            self.frame_cards,
            fg_color="#FFEBEE",
            corner_radius=8,
            border_width=1,
            border_color="#FFCDD2"
        )
        self.card_erro.grid(row=0, column=1, padx=(5, 0), sticky="ew")
        
        self.label_erro = ctk.CTkLabel(
            self.card_erro,
            text="✗ Erros: 0",
            font=ctk.CTkFont(size=14),
            text_color="#C62828",
            padx=10,
            pady=12
        )
        self.label_erro.pack()
        
        # Botões de ação
        self.frame_botoes = ctk.CTkFrame(
            self.frame_main,
            fg_color="transparent"
        )
        self.frame_botoes.grid(row=7, column=0, sticky="ew", pady=(10, 0))
        self.frame_botoes.grid_columnconfigure(0, weight=1)
        self.frame_botoes.grid_columnconfigure(1, weight=1)
        self.frame_botoes.grid_columnconfigure(2, weight=1)
        
        self.btn_iniciar = ctk.CTkButton(
            self.frame_botoes,
            text="Iniciar",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=self.cor_accent,
            hover_color="#1565C0",
            corner_radius=8,
            height=40,
            command=self.iniciar_automacao
        )
        self.btn_iniciar.grid(row=0, column=0, padx=3, pady=5, sticky="ew")
        
        self.btn_pausar = ctk.CTkButton(
            self.frame_botoes,
            text="Pausar",
            font=ctk.CTkFont(size=14),
            fg_color=self.cor_warning,
            hover_color="#FFA000",
            corner_radius=8,
            height=40,
            command=self.pausar_automacao
        )
        self.btn_pausar.grid(row=0, column=1, padx=3, pady=5, sticky="ew")
        self.btn_pausar.configure(state="disabled")
        
        self.btn_parar = ctk.CTkButton(
            self.frame_botoes,
            text="Parar",
            font=ctk.CTkFont(size=14),
            fg_color=self.cor_error,
            hover_color="#D32F2F",
            corner_radius=8,
            height=40,
            command=self.parar_automacao
        )
        self.btn_parar.grid(row=0, column=2, padx=3, pady=5, sticky="ew")
        self.btn_parar.configure(state="disabled")
        
        # Botão de carregar arquivo
        self.btn_novo_arquivo = ctk.CTkButton(
            self.frame_main,
            text="Carregar Novo Arquivo",
            font=ctk.CTkFont(size=14),
            fg_color="#757575",  # Cinza
            hover_color="#616161",  # Cinza mais escuro
            corner_radius=8,
            height=40,
            command=self.carregar_novo_arquivo
        )
        self.btn_novo_arquivo.grid(row=8, column=0, sticky="ew", pady=(10, 0))
        
        # Data e hora da última execução no rodapé
        self.label_ultima_exec = ctk.CTkLabel(
            self.frame_main,
            text="Última execução: -",
            font=ctk.CTkFont(size=12),
            text_color="#757575",
            anchor="w"
        )
        self.label_ultima_exec.grid(row=9, column=0, sticky="w", pady=(15, 0))

        # Adicione isso após o label_ultima_exec no método setup_painel_controle:

        # Frame para os botões de relatórios
        self.frame_relatorios = ctk.CTkFrame(
            self.frame_main,
            fg_color="#F0F6FF",  # Azul bem claro
            corner_radius=8,
            border_width=1,
            border_color="#BBDEFB"
        )
        self.frame_relatorios.grid(row=10, column=0, sticky="ew", pady=(15, 0))
        self.frame_relatorios.grid_columnconfigure(0, weight=1)

        # Título para a seção de relatórios
        self.label_relatorios = ctk.CTkLabel(
            self.frame_relatorios,
            text="RELATÓRIOS",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#1565C0",
            padx=15, pady=10
        )
        self.label_relatorios.grid(row=0, column=0, sticky="w")

        # Botões para gerenciar relatórios
        self.btn_abrir_pasta = ctk.CTkButton(
            self.frame_relatorios,
            text="Abrir Pasta de Relatórios",
            font=ctk.CTkFont(size=12),
            fg_color="#42A5F5",  # Azul mais claro
            hover_color="#1E88E5",
            corner_radius=6,
            height=30,
            command=self.abrir_pasta_relatorios
        )
        self.btn_abrir_pasta.grid(row=1, column=0, padx=10, pady=(0, 5), sticky="ew")

        self.btn_baixar_sucesso = ctk.CTkButton(
            self.frame_relatorios,
            text="Baixar Relatório de Sucesso",
            font=ctk.CTkFont(size=12),
            fg_color="#66BB6A",  # Verde claro
            hover_color="#4CAF50",
            corner_radius=6,
            height=30,
            command=lambda: self.baixar_relatorio("sucesso")
        )
        self.btn_baixar_sucesso.grid(row=2, column=0, padx=10, pady=5, sticky="ew")

        self.btn_baixar_erro = ctk.CTkButton(
            self.frame_relatorios,
            text="Baixar Relatório de Erros",
            font=ctk.CTkFont(size=12),
            fg_color="#EF5350",  # Vermelho claro
            hover_color="#E53935",
            corner_radius=6,
            height=30,
            command=lambda: self.baixar_relatorio("erro")
        )
        self.btn_baixar_erro.grid(row=3, column=0, padx=10, pady=(5, 10), sticky="ew")

    def setup_tabela_status(self):
        # Frame direito
        self.frame_tabela = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color="white",
            border_width=0
        )
        self.frame_tabela.grid(row=0, column=1, sticky="nsew", padx=(0, 15), pady=15)
        self.frame_tabela.grid_columnconfigure(0, weight=1)
        self.frame_tabela.grid_rowconfigure(0, weight=0)  # Cabeçalho
        self.frame_tabela.grid_rowconfigure(1, weight=1)  # Tabela
        
        # Cabeçalho da tabela
        self.frame_tabela_header = ctk.CTkFrame(
            self.frame_tabela,
            fg_color="transparent",
            height=50
        )
        self.frame_tabela_header.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 5))
        self.frame_tabela_header.grid_columnconfigure(0, weight=1)
        
        self.label_tabela_titulo = ctk.CTkLabel(
            self.frame_tabela_header,
            text="Números de Série e Status",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=self.cor_accent,
            anchor="w"
        )
        self.label_tabela_titulo.grid(row=0, column=0, sticky="w")
        
        # Configurar estilos personalizados para a tabela
        style = ttk.Style()
        style.theme_use('default')  # Reset para o tema padrão
        
        # Configurações gerais da tabela
        style.configure(
            "Treeview", 
            background="white",
            foreground="black", 
            rowheight=28, 
            fieldbackground="white",
            font=('Segoe UI', 11)
        )
        style.configure(
            "Treeview.Heading",
            background="#F5F5F5", 
            foreground="#212121",
            relief="flat",
            font=('Segoe UI', 11, 'bold')
        )
        
        # Remover bordas
        style.layout("Treeview", [
            ('Treeview.treearea', {'sticky': 'nswe'})
        ])
        
        # Estilo do cabeçalho quando selecionado
        style.map("Treeview.Heading",
            background=[('active', '#E0E0E0')]
        )
        
        # Cores para os diferentes status
        style.configure("aguardando.Treeview.Row", background="#F5F5F5")
        style.configure("processando.Treeview.Row", background="#FFF9C4")
        style.configure("sucesso.Treeview.Row", background="#C8E6C9")
        style.configure("erro.Treeview.Row", background="#FFCDD2")
        
        # Frame para conter a tabela e scrollbars
        self.frame_treeview = ttk.Frame(self.frame_tabela)
        self.frame_treeview.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 15))
        
        # Scrollbar vertical
        self.vsb = ttk.Scrollbar(self.frame_treeview, orient="vertical")
        self.vsb.pack(side='right', fill='y')
        
        # Scrollbar horizontal
        self.hsb = ttk.Scrollbar(self.frame_treeview, orient="horizontal")
        self.hsb.pack(side='bottom', fill='x')
        
        # Tabela
        self.tabela = ttk.Treeview(
            self.frame_treeview, 
            columns=("nserie", "status", "mensagem", "hora"),
            show='headings',
            yscrollcommand=self.vsb.set,
            xscrollcommand=self.hsb.set
        )
        
        # Configurar as barras de rolagem
        self.vsb.config(command=self.tabela.yview)
        self.hsb.config(command=self.tabela.xview)
        
        # Configurar as colunas da tabela
        self.tabela.heading("nserie", text="Número de Série")
        self.tabela.heading("status", text="Status")
        self.tabela.heading("mensagem", text="Mensagem")
        self.tabela.heading("hora", text="Hora")
        
        self.tabela.column("nserie", width=150, anchor="center")
        self.tabela.column("status", width=100, anchor="center")
        self.tabela.column("mensagem", width=400)
        self.tabela.column("hora", width=100, anchor="center")
        
        self.tabela.pack(expand=True, fill='both')

    def setup_barra_status(self):
        # Barra de status na parte inferior da janela
        self.barra_status = ctk.CTkFrame(
            self,
            height=25,
            fg_color="#ECEFF1"
        )
        self.barra_status.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.barra_status.grid_columnconfigure(0, weight=1)
        
        self.label_status_bar = ctk.CTkLabel(
            self.barra_status,
            text="Pronto",
            font=ctk.CTkFont(size=11),
            text_color="#616161"
        )
        self.label_status_bar.grid(row=0, column=0, padx=10, sticky="w")

    def carregar_dados_tabela(self, preserve_existing=False):
        """Carrega os dados na tabela de status a partir do arquivo Excel"""
        try:
            # Verificar se o arquivo existe
            if not os.path.exists(sc.file_path):
                return
                
            # Carregar dados do Excel
            df = pd.read_excel(sc.file_path, sheet_name=0)
            
            # Se não estamos preservando dados, limpar tabela atual
            if not preserve_existing:
                for item in self.tabela.get_children():
                    self.tabela.delete(item)
                
                # Adicionar os números de série na tabela com status "Aguardando"
                for idx, row in df.iterrows():
                    if "Nserie" in row:
                        numero_serie = str(row["Nserie"]).strip()
                        self.tabela.insert("", "end", values=(
                            numero_serie,
                            "Aguardando",
                            "Aguardando processamento",
                            "-"
                        ), tags=("aguardando",))
                
                # Aplicar estilo
                for item in self.tabela.get_children():
                    self.tabela.item(item, tags=("aguardando",))
            else:
                # Se estamos preservando, apenas sincronizamos novos itens
                self.sincronizar_tabela_com_planilha()
        except Exception as e:
            messagebox.showerror("Erro ao carregar dados", f"Não foi possível carregar os dados da planilha: {str(e)}")

    def atualizar_status_item_tabela(self, nserie, status, mensagem):
        """Atualiza o status de um item na tabela"""
        hora_atual = datetime.now().strftime("%H:%M:%S")
        
        # Procurar o item na tabela
        for item_id in self.tabela.get_children():
            item_values = self.tabela.item(item_id, "values")
            if item_values and item_values[0] == nserie:
                # Verificar se o horário deve ser atualizado
                # - Manter o horário existente se não for "-" e o status for final (Sucesso/Erro)
                # - Atualizar horário apenas para novos status de Sucesso/Erro ou se item anterior era "Aguardando"/"Processando"
                hora = hora_atual
                if item_values[3] != "-" and item_values[1] in ["Sucesso", "Erro"]:
                    # Manter o horário registrado anteriormente
                    hora = item_values[3]
                elif status.lower() in ["sucesso", "erro"]:
                    # Registrar horário atual para novos sucessos/erros
                    hora = hora_atual
                    
                # Atualizar valores
                self.tabela.item(item_id, values=(
                    nserie,
                    status,
                    mensagem,
                    hora
                ))
                
                # Atualizar tag para mudar a cor
                if status.lower() == "processando":
                    self.tabela.item(item_id, tags=("processando",))
                elif status.lower() == "sucesso":
                    self.tabela.item(item_id, tags=("sucesso",))
                elif status.lower() == "erro":
                    self.tabela.item(item_id, tags=("erro",))
                else:
                    self.tabela.item(item_id, tags=("aguardando",))
                
                # Fazer scroll para o item atualizado
                self.tabela.see(item_id)
                break
    
    def verificar_planilha(self, preserve_table_data=False):
        """Verifica o status da planilha Excel e atualiza a interface"""
        planilha_ok, msg, total = sc.verificar_planilha()
        if planilha_ok:
            self.label_planilha.configure(text=f"✅ {msg}", text_color="green")
            self.btn_iniciar.configure(state="normal")
            # Carregar números de série na tabela, preservando dados existentes se solicitado
            self.carregar_dados_tabela(preserve_existing=preserve_table_data)
        else:
            self.label_planilha.configure(text=f"❌ {msg}", text_color="red")
            self.btn_iniciar.configure(state="disabled")
    
    def update_status_periodico(self):
        """Atualiza as informações da interface periodicamente"""
        
        # Verificação de nova tentativa pendente
        if sc.RETRY_PENDING:
            # Pausar tempo para não contar no tempo de estimativa
            tempo_pausa_inicio = time.time()
            
            # Solicitar ao usuário se deseja tentar novamente
            resposta = messagebox.askyesno(
                "Erro na Movimentação", 
                f"Ocorreu um erro ao movimentar o número de série:\n\n{sc.RETRY_ITEM}\n\nDeseja tentar novamente?"
            )
            
            # Atualizar variáveis com a resposta
            sc.RETRY_REQUESTED = resposta
            sc.RETRY_PENDING = False
            
            # Registrar o tempo de pausa para não afetar a estimativa
            tempo_pausa = time.time() - tempo_pausa_inicio
            if sc.START_TIME:
                sc.START_TIME += tempo_pausa
        
        # Verificar se há um item aguardando decisão
        if sc.RETRY_PENDING and sc.RETRY_ITEM:
            self.atualizar_status_item_tabela(
                sc.RETRY_ITEM, 
                "Aguardando Decisão",
                "Aguardando decisão para tentar novamente",
                "-"
            )
        
        # Verificar planilha periodicamente quando não estiver em execução
        # e não tiver sido finalizada recentemente
        if (not sc.IS_RUNNING and not hasattr(self, 'execution_just_finished') and 
            (not hasattr(self, '_last_check') or time.time() - self._last_check > 5)):
            # Verificar planilha sem recarregar a tabela
            planilha_ok, msg, total = sc.verificar_planilha()
            if planilha_ok:
                self.label_planilha.configure(text=f"✅ {msg}", text_color="green")
                self.btn_iniciar.configure(state="normal")
            else:
                self.label_planilha.configure(text=f"❌ {msg}", text_color="red")
                self.btn_iniciar.configure(state="disabled")
            self._last_check = time.time()
            
        # Atualizar status principal
        self.label_status.configure(text=sc.STATUS)
        
        # Atualizar contadores de progresso e estado dos botões
        if sc.IS_RUNNING:
            total_itens = sc.total_itens if sc.total_itens > 0 else 1
            progresso = sc.movimentados / total_itens
            self.label_progresso.configure(text=f"Progresso: {sc.movimentados}/{total_itens}")
            self.progress_bar.set(progresso)
            
            # Calcular tempo estimado para conclusão
            if sc.times_of_processing and sc.START_TIME and sc.movimentados > 0:
                tempo_medio_por_item = sc.times_of_processing[-1] / sc.movimentados
                itens_restantes = total_itens - sc.movimentados
                tempo_restante_sec = tempo_medio_por_item * itens_restantes
                
                # Formatar em HH:MM:SS
                tempo_restante = str(timedelta(seconds=int(tempo_restante_sec)))
                self.label_tempo.configure(text=f"Tempo estimado: {tempo_restante}")
            
            # Atualizar contadores
            self.label_sucesso.configure(text=f"✅ Sucesso: {len(sc.lista_movimentados)}")
            self.label_erro.configure(text=f"❌ Erros: {len(sc.lista_erros)}")
            
            # Atualizar tabela com base nas listas de movimentados e erros
            for item in sc.lista_movimentados:
                self.atualizar_status_item_tabela(item, "Sucesso", "Item movimentado com sucesso")
                
            for item in sc.lista_erros:
                self.atualizar_status_item_tabela(item, "Erro", "Falha na movimentação do item")
            
            # Verificar item atual sendo processado
            if hasattr(sc, 'item_atual') and sc.item_atual:
                # Não atualizar itens que já têm status de Sucesso ou Erro
                ja_finalizado = False
                for item_id in self.tabela.get_children():
                    item_values = self.tabela.item(item_id, "values")
                    if item_values and item_values[0] == sc.item_atual and item_values[1] in ["Sucesso", "Erro"]:
                        ja_finalizado = True
                        break
                        
                if not ja_finalizado:
                    self.atualizar_status_item_tabela(
                        sc.item_atual, 
                        "Processando",
                        "Processando item no sistema"
                    )
            
            # Gerenciar estado dos botões
            self.btn_iniciar.configure(state="disabled")
            self.btn_parar.configure(state="normal")
            self.btn_pausar.configure(state="normal")
            self.btn_novo_arquivo.configure(state="disabled")
            
            # Atualizar texto do botão de pausa
            if sc.IS_PAUSED:
                self.btn_pausar.configure(text="Retomar", fg_color="#4CAF50", hover_color="#388E3C")
            else:
                self.btn_pausar.configure(text="Pausar", fg_color="#FFA000", hover_color="#FF8F00")
            
        elif sc.IS_FINISHED:
            # Processo finalizado
            self.progress_bar.set(1.0 if sc.movimentados > 0 else 0.0)
            self.label_sucesso.configure(text=f"✅ Sucesso: {len(sc.lista_movimentados)}")
            self.label_erro.configure(text=f"❌ Erros: {len(sc.lista_erros)}")
            self.label_tempo.configure(text="Tempo estimado: Concluído")
            
            # Registrar data/hora da execução
            agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            self.label_ultima_exec.configure(text=f"Última execução: {agora}")
            
            # Reativar botão iniciar e desativar botões de pausa/parar
            self.btn_iniciar.configure(state="normal")
            self.btn_parar.configure(state="disabled")
            self.btn_pausar.configure(state="disabled")
            self.btn_novo_arquivo.configure(state="normal")
            
            # Marcar que uma execução acabou de ser finalizada para evitar recarregar a tabela
            self.execution_just_finished = True
            
            # Resetar flag IS_FINISHED para não repetir esta lógica
            sc.IS_FINISHED = False
        
        # Atualizar estado dos botões de relatório
        self.atualizar_estado_botoes_relatorio()
        
        # Atualizar novamente após 500ms
        self.after(500, self.update_status_periodico)
    
    def iniciar_automacao(self):
        """Inicia a automação em uma thread separada"""
        # Pegar opção selecionada
        sc.OPCAO_DESEJADA = self.dropdown_opcao.get()
        # Forçar modo headless como False
        headless = False
        
        # Confirmar antes de iniciar
        if not messagebox.askyesno("Confirmação", 
                                  "Deseja iniciar a automação?\n\n" +
                                  f"- Unidade: {sc.OPCAO_DESEJADA}\n\n" +
                                  "Certifique-se que o arquivo nserie.xlsx está pronto."):
            return
        
        # Verificar se há novos itens na planilha que não estão na tabela
        self.sincronizar_tabela_com_planilha()
        
        # Remover a flag de execução finalizada se existir
        if hasattr(self, 'execution_just_finished'):
            delattr(self, 'execution_just_finished')
        
        # Iniciar thread de automação
        threading.Thread(target=lambda: sc.run_sistema_cofin(headless=False), daemon=True).start()
    
    def pausar_automacao(self):
        """Pausa ou retoma a automação em execução"""
        if sc.IS_PAUSED:
            # Retomar
            sc.resume_script()
            self.btn_pausar.configure(text="Pausar", fg_color="#FFA000", hover_color="#FF8F00")
        else:
            # Pausar
            sc.pause_script()
            self.btn_pausar.configure(text="Retomar", fg_color="#4CAF50", hover_color="#388E3C")
    
    def parar_automacao(self):
        """Para a automação em execução"""
        if messagebox.askyesno("Parar automação", "Tem certeza que deseja interromper a automação?"):
            status = sc.force_stop_script()
            messagebox.showinfo("Automação interrompida", f"A automação foi interrompida.\n\n{status}")

    def carregar_novo_arquivo(self):
        """Permite selecionar um novo arquivo nserie.xlsx"""
        # Verificar se a automação está em execução
        if sc.IS_RUNNING:
            messagebox.showwarning("Aviso", "Pare a automação antes de carregar um novo arquivo.")
            return
            
        # Abrir diálogo para selecionar arquivo
        file_path = filedialog.askopenfilename(
            title="Selecionar arquivo nserie.xlsx", 
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        
        if not file_path:
            return  # Usuário cancelou
        
        try:
            # Copiar o arquivo selecionado para o diretório do programa
            shutil.copy(file_path, sc.file_path)
            
            # Verificar o arquivo e recarregar a tabela
            self.verificar_planilha()
            
            messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível carregar o arquivo: {str(e)}")

    def sincronizar_tabela_com_planilha(self):
        """Atualiza a tabela com novos itens da planilha sem apagar dados existentes"""
        try:
            # Verificar se o arquivo existe
            if not os.path.exists(sc.file_path):
                return
                
            # Carregar dados do Excel
            df = pd.read_excel(sc.file_path, sheet_name=0)
            
            # Obter números de série já presentes na tabela
            numeros_existentes = set()
            for item_id in self.tabela.get_children():
                item_values = self.tabela.item(item_id, "values")
                if item_values:
                    numeros_existentes.add(item_values[0])
            
            # Adicionar apenas os novos números de série
            for idx, row in df.iterrows():
                if "Nserie" in row:
                    numero_serie = str(row["Nserie"]).strip()
                    if numero_serie not in numeros_existentes:
                        self.tabela.insert("", "end", values=(
                            numero_serie,
                            "Aguardando",
                            "Aguardando processamento",
                            "-"
                        ), tags=("aguardando",))
        except Exception as e:
            messagebox.showerror("Erro ao sincronizar tabela", f"Não foi possível sincronizar a tabela com a planilha: {str(e)}")

    def on_closing(self):
        """Método chamado quando a janela é fechada"""
        # Se a automação estiver em execução, manter o navegador aberto
        if sc.IS_RUNNING:
            sc.force_stop_script()
        
        # Destruir a janela (encerrar o aplicativo)
        self.destroy()

    def abrir_pasta_relatorios(self):
        """Abre a pasta de relatórios no explorador de arquivos"""
        pasta_relatorios = os.path.join(sc.base_dir, "relatorios_excel")
        
        # Criar pasta se não existir
        if not os.path.exists(pasta_relatorios):
            os.makedirs(pasta_relatorios)
            
        # Abrir pasta no explorador de arquivos
        os.startfile(pasta_relatorios)

    def baixar_relatorio(self, tipo):
        """
        Baixa o relatório mais recente do tipo especificado
        tipo: "sucesso" ou "erro"
        """
        pasta_relatorios = os.path.join(sc.base_dir, "relatorios_excel")
        
        # Criar pasta se não existir
        if not os.path.exists(pasta_relatorios):
            os.makedirs(pasta_relatorios)
            messagebox.showinfo("Informação", "Nenhum relatório encontrado.")
            return
        
        # Procurar arquivos de acordo com o tipo
        prefixo = "movimentados_" if tipo == "sucesso" else "erros_"
        
        # Listar todos os arquivos com o prefixo
        arquivos = [f for f in os.listdir(pasta_relatorios) if f.startswith(prefixo)]
        
        if not arquivos:
            messagebox.showinfo("Informação", f"Nenhum relatório de {tipo} encontrado.")
            return
        
        # Ordenar por data (mais recente primeiro)
        arquivos.sort(reverse=True)
        
        # Solicitar ao usuário onde salvar o arquivo
        destino = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx")],
            initialfile=arquivos[0]  # Usar o nome do arquivo mais recente
        )
        
        if destino:
            # Copiar arquivo para o destino escolhido
            origem = os.path.join(pasta_relatorios, arquivos[0])
            shutil.copy(origem, destino)
            messagebox.showinfo("Sucesso", f"Relatório de {tipo} salvo com sucesso!")

    def atualizar_estado_botoes_relatorio(self):
        """Atualiza o estado dos botões de relatório com base na existência dos arquivos"""
        pasta_relatorios = os.path.join(sc.base_dir, "relatorios_excel")
        
        if not os.path.exists(pasta_relatorios):
            self.btn_baixar_sucesso.configure(state="disabled")
            self.btn_baixar_erro.configure(state="disabled")
            return
        
        # Verificar relatórios de sucesso
        arquivos_sucesso = [f for f in os.listdir(pasta_relatorios) if f.startswith("movimentados_")]
        self.btn_baixar_sucesso.configure(state="normal" if arquivos_sucesso else "disabled")
        
        # Verificar relatórios de erro
        arquivos_erro = [f for f in os.listdir(pasta_relatorios) if f.startswith("erros_")]
        self.btn_baixar_erro.configure(state="normal" if arquivos_erro else "disabled")

if __name__ == "__main__":
    app = AutocofinApp()
    app.mainloop()