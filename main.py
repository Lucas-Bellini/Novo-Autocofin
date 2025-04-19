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
        self.geometry("1280x720")  # Resolução mais ampla
        self.minsize(1200, 780)    # Tamanho mínimo aumentado para melhor visualização
        
        # Novo esquema de cores - mais neutro e profissional
        self.cor_bg = "#F5F7FA"          # Fundo geral (cinza muito claro com tom azulado)
        self.cor_accent = "#3B71CA"      # Azul principal mais neutro
        self.cor_text = "#212B36"        # Texto escuro com tom azulado
        self.cor_text_secondary = "#637381" # Texto secundário (cinza médio)
        self.cor_success = "#14A44D"     # Verde sucesso mais discreto
        self.cor_warning = "#E4A11B"     # Amarelo aviso mais discreto
        self.cor_error = "#DC4C64"       # Vermelho erro mais discreto
        self.cor_card_bg = "#FFFFFF"     # Branco para cartões
        self.cor_border = "#E2E8F0"      # Borda leve
        self.cor_input_bg = "#F8FAFC"    # Fundo para inputs
        
        # Configurar tema da janela
        self.configure(fg_color=self.cor_bg)
        
        # Layout da interface
        self.grid_columnconfigure(0, weight=5)  # Painel de controle
        self.grid_columnconfigure(1, weight=7)  # Tabela de status
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
            corner_radius=12,
            fg_color=self.cor_card_bg,
            border_width=1,
            border_color=self.cor_border
        )
        self.frame_controles.grid(row=0, column=0, sticky="nsew", padx=(20, 10), pady=20)
        self.frame_controles.grid_columnconfigure(0, weight=1)
        self.frame_controles.grid_rowconfigure(0, weight=1)
        
        # Frame com scroll para conter todos os controles
        self.scroll_controles = ctk.CTkScrollableFrame(
            self.frame_controles,
            fg_color="transparent",
            corner_radius=0,
            scrollbar_button_color=self.cor_accent,
            scrollbar_button_hover_color="#2C5FAA"
        )
        self.scroll_controles.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        self.scroll_controles.grid_columnconfigure(0, weight=1)
        
        # Cabeçalho com logo e título
        self.frame_header = ctk.CTkFrame(
            self.scroll_controles,
            fg_color="transparent",
            height=90
        )
        self.frame_header.grid(row=0, column=0, sticky="ew", padx=25, pady=(25, 10))
        
        # Título do aplicativo
        self.label_titulo = ctk.CTkLabel(
            self.frame_header, 
            text="AUTOCOFIN",
            font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"),
            text_color=self.cor_accent
        )
        self.label_titulo.grid(row=0, column=0, sticky="w")
        
        self.label_subtitulo = ctk.CTkLabel(
            self.frame_header, 
            text="Sistema de Automação COFIN",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color=self.cor_text_secondary
        )
        self.label_subtitulo.grid(row=1, column=0, sticky="w", pady=(0, 5))
        
        # Separador com gradiente
        self.separador = ctk.CTkFrame(
            self.scroll_controles, 
            height=2,
            fg_color=self.cor_accent
        )
        self.separador.grid(row=1, column=0, sticky="ew", padx=25, pady=(0, 25))
        
        # Frame principal para os controles com melhor espaçamento
        self.frame_main = ctk.CTkFrame(
            self.scroll_controles,
            fg_color="transparent"
        )
        self.frame_main.grid(row=2, column=0, sticky="nsew", padx=25, pady=0)
        self.frame_main.grid_columnconfigure(0, weight=1)
        
        # ===== Seção 1: Status da Planilha =====
        self.frame_status_planilha = ctk.CTkFrame(
            self.frame_main,
            fg_color="#EBF5FF",  # Azul muito claro
            corner_radius=10,
            border_width=1,
            border_color="#BDE0FE"
        )
        self.frame_status_planilha.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        
        self.label_planilha = ctk.CTkLabel(
            self.frame_status_planilha,
            text="Verificando planilha...",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            padx=20, pady=12,
            text_color=self.cor_accent
        )
        self.label_planilha.pack(fill="both")
        
        # ===== Seção 2: Seleção de Unidade =====
        # Título da seção
        self.label_section_unidade = ctk.CTkLabel(
            self.frame_main,
            text="CONFIGURAÇÕES",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=self.cor_text_secondary,
            anchor="w"
        )
        self.label_section_unidade.grid(row=1, column=0, sticky="w", pady=(0, 10))
        
        # Opção da unidade
        self.label_opcao = ctk.CTkLabel(
            self.frame_main,
            text="Unidade:",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=self.cor_text,
            anchor="w"
        )
        self.label_opcao.grid(row=2, column=0, sticky="w", pady=(0, 8))
        
        # Lista de opções de unidade
        self.opcoes_unidades = [
            "303017500 - CMB SEC DISTR",
            "303017400 - CMB SEC RECEB",
            "303016500 - CMB 3 - RESUMO"
        ]
        
        # Frame para listbox de unidades
        self.frame_listbox = ctk.CTkFrame(
            self.frame_main,
            fg_color=self.cor_input_bg,
            corner_radius=10,
            border_width=1,
            border_color=self.cor_border
        )
        self.frame_listbox.grid(row=3, column=0, sticky="ew", pady=(0, 20))
        self.frame_listbox.grid_columnconfigure(0, weight=1)
        
        # Listbox para seleção de unidade
        self.listbox_opcao = ctk.CTkScrollableFrame(
            self.frame_listbox,
            fg_color="transparent",
            height=100,  # Aumentado para melhor visualização
            scrollbar_button_color=self.cor_accent,
            scrollbar_button_hover_color="#2C5FAA"
        )
        self.listbox_opcao.grid(row=0, column=0, padx=5, pady=8, sticky="ew")
        
        # Variável para armazenar a opção selecionada
        self.opcao_selecionada = ctk.StringVar(value=self.opcoes_unidades[0])
        
        # Radio buttons mais espaçados e com melhor estilo
        for i, opcao in enumerate(self.opcoes_unidades):
            rb = ctk.CTkRadioButton(
                self.listbox_opcao,
                text=opcao,
                variable=self.opcao_selecionada,
                value=opcao,
                font=ctk.CTkFont(family="Segoe UI", size=13),
                fg_color=self.cor_accent,
                # Remover border_width e substituir pelos corretos
                border_width_unchecked=2,  # Largura da borda quando não selecionado
                hover_color="#2C5FAA",
                text_color=self.cor_text
            )
            rb.grid(row=i, column=0, padx=10, pady=5, sticky="w")  # Mais espaço vertical
        
        # ===== Seção 3: Opções de Execução =====
        # Checkbox para ativar/desativar alertas de erro
        self.frame_opcoes_exec = ctk.CTkFrame(
            self.frame_main,
            fg_color=self.cor_input_bg,
            corner_radius=10,
            border_width=1,
            border_color=self.cor_border
        )
        self.frame_opcoes_exec.grid(row=4, column=0, sticky="ew", pady=(0, 20))
        
        self.label_opcoes_exec = ctk.CTkLabel(
            self.frame_opcoes_exec,
            text="Opções de execução:",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=self.cor_text,
            anchor="w"
        )
        self.label_opcoes_exec.grid(row=0, column=0, padx=15, pady=(12, 4), sticky="w")
        
        self.var_alertas_erro = ctk.BooleanVar(value=True)
        self.check_alertas_erro = ctk.CTkCheckBox(
            self.frame_opcoes_exec,
            text="Alertar sobre erros durante a movimentação",
            variable=self.var_alertas_erro,
            font=ctk.CTkFont(family="Segoe UI", size=13),
            checkbox_width=22,
            checkbox_height=22,
            border_width=2,
            hover_color=self.cor_accent,
            fg_color=self.cor_accent,
            text_color=self.cor_text
        )
        self.check_alertas_erro.grid(row=1, column=0, padx=15, pady=(0, 12), sticky="w")
        
        # ===== Seção 4: Status da Operação =====
        self.label_section_status = ctk.CTkLabel(
            self.frame_main,
            text="STATUS DA OPERAÇÃO",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=self.cor_text_secondary,
            anchor="w"
        )
        self.label_section_status.grid(row=5, column=0, sticky="w", pady=(0, 10))
        
        # Status atual - Card com fundo suave e scroll
        self.frame_status = ctk.CTkFrame(
            self.frame_main,
            fg_color=self.cor_card_bg,
            corner_radius=10,
            border_width=1,
            border_color=self.cor_border,
            height=110  # Altura fixa para evitar expansão indesejada
        )
        self.frame_status.grid(row=6, column=0, sticky="ew", pady=(0, 20))
        self.frame_status.grid_columnconfigure(0, weight=1)
        self.frame_status.grid_propagate(False)  # Impedir que o frame se expanda
        
        # Frame com scroll para o texto do status
        self.status_scroll_frame = ctk.CTkScrollableFrame(
            self.frame_status,
            fg_color="transparent",
            height=80,
            scrollbar_button_color=self.cor_accent,
            scrollbar_button_hover_color="#2C5FAA"
        )
        self.status_scroll_frame.grid(row=0, column=0, padx=15, pady=15, sticky="ew")
        self.status_scroll_frame.grid_columnconfigure(0, weight=1)
        
        # Label de status dentro do frame com scroll
        self.label_status = ctk.CTkLabel(
            self.status_scroll_frame,
            text="Aguardando",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            wraplength=350,
            justify="left",
            text_color=self.cor_text
        )
        self.label_status.grid(row=0, column=0, sticky="w")
        
        # ===== Seção 5: Progresso =====
        self.label_section_progresso = ctk.CTkLabel(
            self.frame_main,
            text="PROGRESSO",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=self.cor_text_secondary,
            anchor="w"
        )
        self.label_section_progresso.grid(row=7, column=0, sticky="w", pady=(0, 10))
        
        # Frame para contadores de progresso
        self.frame_progresso = ctk.CTkFrame(
            self.frame_main,
            fg_color="transparent"
        )
        self.frame_progresso.grid(row=8, column=0, sticky="ew", pady=(0, 8))
        self.frame_progresso.grid_columnconfigure(0, weight=1)
        self.frame_progresso.grid_columnconfigure(1, weight=1)
        
        self.label_progresso = ctk.CTkLabel(
            self.frame_progresso,
            text="Progresso: 0/0",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color=self.cor_text,
            anchor="w"
        )
        self.label_progresso.grid(row=0, column=0, sticky="w")
        
        self.label_tempo = ctk.CTkLabel(
            self.frame_progresso,
            text="Tempo: --:--",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color=self.cor_text,
            anchor="e"
        )
        self.label_tempo.grid(row=0, column=1, sticky="e")
        
        # Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(
            self.frame_main,
            height=16,
            corner_radius=8,
            progress_color=self.cor_accent,
            fg_color=self.cor_input_bg
        )
        self.progress_bar.grid(row=9, column=0, sticky="ew", pady=(0, 20))
        self.progress_bar.set(0)
        
        # ===== Seção 6: Resultados em Cards =====
        # Cards de sucesso e erro em um design mais moderno
        self.frame_cards = ctk.CTkFrame(
            self.frame_main,
            fg_color="transparent"
        )
        self.frame_cards.grid(row=10, column=0, sticky="ew", pady=(0, 20))
        self.frame_cards.grid_columnconfigure(0, weight=1)
        self.frame_cards.grid_columnconfigure(1, weight=1)
        
        # Card de sucesso
        self.card_sucesso = ctk.CTkFrame(
            self.frame_cards,
            fg_color="#ECFDF5",  # Verde muito claro
            corner_radius=10,
            border_width=1,
            border_color="#C6F7E2"
        )
        self.card_sucesso.grid(row=0, column=0, padx=(0, 6), sticky="ew")
        
        self.label_sucesso = ctk.CTkLabel(
            self.card_sucesso,
            text="✓ Sucesso: 0",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=self.cor_success,
            padx=15,
            pady=15
        )
        self.label_sucesso.pack()
        
        # Card de erro
        self.card_erro = ctk.CTkFrame(
            self.frame_cards,
            fg_color="#FEF2F2",  # Vermelho muito claro
            corner_radius=10,
            border_width=1,
            border_color="#FEECEC"
        )
        self.card_erro.grid(row=0, column=1, padx=(6, 0), sticky="ew")
        
        self.label_erro = ctk.CTkLabel(
            self.card_erro,
            text="✗ Erros: 0",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=self.cor_error,
            padx=15,
            pady=15
        )
        self.label_erro.pack()
        
        # ===== Seção 7: Botões de Ação =====
        self.label_section_acoes = ctk.CTkLabel(
            self.frame_main,
            text="AÇÕES",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=self.cor_text_secondary,
            anchor="w"
        )
        self.label_section_acoes.grid(row=11, column=0, sticky="w", pady=(0, 10))
        
        # Botões de ação com design mais refinado
        self.frame_botoes = ctk.CTkFrame(
            self.frame_main,
            fg_color="transparent"
        )
        self.frame_botoes.grid(row=12, column=0, sticky="ew", pady=(0, 15))
        self.frame_botoes.grid_columnconfigure(0, weight=1)
        self.frame_botoes.grid_columnconfigure(1, weight=1)
        self.frame_botoes.grid_columnconfigure(2, weight=1)
        
        self.btn_iniciar = ctk.CTkButton(
            self.frame_botoes,
            text="INICIAR",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color=self.cor_accent,
            hover_color="#2C5FAA",
            corner_radius=8,
            height=42,
            command=self.iniciar_automacao
        )
        self.btn_iniciar.grid(row=0, column=0, padx=4, pady=5, sticky="ew")
        
        self.btn_pausar = ctk.CTkButton(
            self.frame_botoes,
            text="PAUSAR",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            fg_color=self.cor_warning,
            hover_color="#C68B09",
            corner_radius=8,
            height=42,
            command=self.pausar_automacao
        )
        self.btn_pausar.grid(row=0, column=1, padx=4, pady=5, sticky="ew")
        self.btn_pausar.configure(state="disabled")
        
        self.btn_parar = ctk.CTkButton(
            self.frame_botoes,
            text="PARAR",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            fg_color=self.cor_error,
            hover_color="#C62350",
            corner_radius=8,
            height=42,
            command=self.parar_automacao
        )
        self.btn_parar.grid(row=0, column=2, padx=4, pady=5, sticky="ew")
        self.btn_parar.configure(state="disabled")
        
        # Botão de carregar arquivo
        self.btn_novo_arquivo = ctk.CTkButton(
            self.frame_main,
            text="Carregar Novo Arquivo",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            fg_color="#64748B",
            hover_color="#475569",
            corner_radius=8,
            height=42,
            command=self.carregar_novo_arquivo
        )
        self.btn_novo_arquivo.grid(row=13, column=0, sticky="ew", pady=(0, 20))
        
        # ===== Seção 8: Relatórios =====
        # Frame para os botões de relatórios - design mais moderno
        self.frame_relatorios = ctk.CTkFrame(
            self.frame_main,
            fg_color="#F1F5F9",
            corner_radius=10,
            border_width=1,
            border_color=self.cor_border
        )
        self.frame_relatorios.grid(row=14, column=0, sticky="ew", pady=(0, 20))
        self.frame_relatorios.grid_columnconfigure(0, weight=1)
        self.frame_relatorios.grid_columnconfigure(1, weight=1)
        
        # Título da seção de relatórios
        self.label_relatorios = ctk.CTkLabel(
            self.frame_relatorios,
            text="RELATÓRIOS",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=self.cor_text_secondary,
            padx=15, pady=12
        )
        self.label_relatorios.grid(row=0, column=0, columnspan=2, sticky="w")
        
        # Botões para gerenciar relatórios - layout em grade com visual consistente
        self.btn_abrir_pasta = ctk.CTkButton(
            self.frame_relatorios,
            text="Abrir Pasta",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            fg_color="#3B82F6",
            hover_color="#2563EB",
            corner_radius=6,
            height=32,
            command=self.abrir_pasta_relatorios
        )
        self.btn_abrir_pasta.grid(row=1, column=0, columnspan=2, padx=12, pady=(0, 10), sticky="ew")
        
        self.btn_baixar_sucesso = ctk.CTkButton(
            self.frame_relatorios,
            text="Baixar Sucesso",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            fg_color=self.cor_success,
            hover_color="#0D9543",
            corner_radius=6,
            height=32,
            command=lambda: self.baixar_relatorio("sucesso")
        )
        self.btn_baixar_sucesso.grid(row=2, column=0, padx=(12, 6), pady=(0, 12), sticky="ew")
        
        self.btn_baixar_erro = ctk.CTkButton(
            self.frame_relatorios,
            text="Baixar Erros",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            fg_color=self.cor_error,
            hover_color="#C62350",
            corner_radius=6,
            height=32,
            command=lambda: self.baixar_relatorio("erro")
        )
        self.btn_baixar_erro.grid(row=2, column=1, padx=(6, 12), pady=(0, 12), sticky="ew")
        
        # Informações sobre última execução
        self.label_ultima_exec = ctk.CTkLabel(
            self.frame_main,
            text="Última execução: -",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=self.cor_text_secondary,
            anchor="w"
        )
        self.label_ultima_exec.grid(row=15, column=0, sticky="w", pady=(0, 10))

    def setup_tabela_status(self):
        # Frame direito
        self.frame_tabela = ctk.CTkFrame(
            self,
            corner_radius=12,
            fg_color=self.cor_card_bg,
            border_width=1,
            border_color=self.cor_border
        )
        self.frame_tabela.grid(row=0, column=1, sticky="nsew", padx=(10, 20), pady=20)
        self.frame_tabela.grid_columnconfigure(0, weight=1)
        self.frame_tabela.grid_rowconfigure(0, weight=0)  # Cabeçalho
        self.frame_tabela.grid_rowconfigure(1, weight=1)  # Tabela
        
        # Cabeçalho da tabela
        self.frame_tabela_header = ctk.CTkFrame(
            self.frame_tabela,
            fg_color="transparent",
            height=60
        )
        self.frame_tabela_header.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        self.frame_tabela_header.grid_columnconfigure(0, weight=1)
        
        self.label_tabela_titulo = ctk.CTkLabel(
            self.frame_tabela_header,
            text="Números de Série e Status",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color=self.cor_accent,
            anchor="w"
        )
        self.label_tabela_titulo.grid(row=0, column=0, sticky="w")
        
        # Configurar estilos personalizados para a tabela
        style = ttk.Style()
        style.theme_use('default')  # Reset para o tema padrão
        
        # Configurações gerais da tabela - cores e fontes atualizadas
        style.configure(
            "Treeview", 
            background=self.cor_card_bg,
            foreground=self.cor_text, 
            rowheight=30,  # Linhas mais altas
            fieldbackground=self.cor_card_bg,
            font=('Segoe UI', 11)
        )
        style.configure(
            "Treeview.Heading",
            background="#F1F5F9", 
            foreground=self.cor_text,
            relief="flat",
            font=('Segoe UI', 11, 'bold')
        )
        
        # Remover bordas
        style.layout("Treeview", [
            ('Treeview.treearea', {'sticky': 'nswe'})
        ])
        
        # Estilo do cabeçalho quando selecionado
        style.map(
            "Treeview.Heading",
            background=[('active', '#E2E8F0')]
        )
        
        # Cores para os diferentes status - mais suaves e elegantes
        style.configure("aguardando.Treeview.Row", background="#F8FAFC")
        style.configure("processando.Treeview.Row", background="#FFFBEB")  # Amarelo muito claro
        style.configure("sucesso.Treeview.Row", background="#EBF5FF")       # Azul muito claro
        style.configure("erro.Treeview.Row", background="#FEF2F2")          # Vermelho muito claro
        
        # Frame para conter a tabela e scrollbars
        self.frame_treeview = ttk.Frame(self.frame_tabela)
        self.frame_treeview.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        
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
        
        # Colunas com larguras mais adequadas
        self.tabela.column("nserie", width=160, anchor="center")
        self.tabela.column("status", width=110, anchor="center")
        self.tabela.column("mensagem", width=450)
        self.tabela.column("hora", width=90, anchor="center")
        
        self.tabela.pack(expand=True, fill='both')

    def setup_barra_status(self):
        # Barra de status na parte inferior da janela - design mais refinado
        self.barra_status = ctk.CTkFrame(
            self,
            height=28,
            fg_color="#F1F5F9",
            border_width=1,
            border_color=self.cor_border
        )
        self.barra_status.grid(row=1, column=0, columnspan=2, sticky="ew", padx=20, pady=(0, 20))
        self.barra_status.grid_columnconfigure(0, weight=1)
        
        self.label_status_bar = ctk.CTkLabel(
            self.barra_status,
            text="Sistema pronto",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=self.cor_text_secondary
        )
        self.label_status_bar.grid(row=0, column=0, padx=15, pady=3, sticky="w")

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

    def atualizar_status_item_tabela(self, nserie, status, mensagem, hora=None):
        """Atualiza o status de um item na tabela"""
        hora_atual = datetime.now().strftime("%H:%M:%S")
        hora = hora or hora_atual
        
        # Procurar o item na tabela
        for item_id in self.tabela.get_children():
            item_values = self.tabela.item(item_id, "values")
            if item_values and item_values[0] == nserie:
                # Verificar se o horário deve ser atualizado
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
            # Verificar se os alertas de erro estão ativados
            if self.var_alertas_erro.get():
                # Comportamento atual - solicitar confirmação do usuário
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
            else:
                # Alertas desativados - continuar sem perguntar
                sc.RETRY_REQUESTED = False  # Não tenta novamente, apenas continua
                sc.RETRY_PENDING = False    # Remove o status de pendente
                # Sem pausa do tempo - execução contínua
        
        # Resto da função permanece igual
        
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
        
        # Assegurar que o texto mais recente seja visível no frame scrollável
        self.status_scroll_frame.update_idletasks()
        self.status_scroll_frame._parent_canvas.yview_moveto(1.0)
        
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
        # Pegar opção selecionada da listbox (usando a variável do RadioButton)
        sc.OPCAO_DESEJADA = self.opcao_selecionada.get()
        # Configurar modo de alertas de erro
        sc.ALERTAR_ERROS = self.var_alertas_erro.get()
        # Forçar modo headless como False
        headless = False
        
        # Confirmar antes de iniciar
        if not messagebox.askyesno("Confirmação", 
                                  "Deseja iniciar a automação?\n\n" +
                                  f"- Unidade: {sc.OPCAO_DESEJADA}\n\n" +
                                  f"- {'Com' if sc.ALERTAR_ERROS else 'Sem'} alertas de erro\n\n" +
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