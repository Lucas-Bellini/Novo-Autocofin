import os
import sys
import tkinter as tk
import customtkinter as ctk
import subprocess
import threading
import time
import shutil
import uuid
import pandas as pd
from datetime import datetime
from tkinter import filedialog, messagebox

# Configuração global do tema
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AutocofinHub(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuração da janela principal
        self.title("Autocofin Hub | Gerenciador de Instâncias")
        self.geometry("1280x720")
        self.minsize(1200, 700)
        
        # Esquema de cores - consistente com o app principal
        self.cor_bg = "#F5F7FA"          # Fundo geral
        self.cor_accent = "#3B71CA"      # Azul principal
        self.cor_text = "#212B36"        # Texto escuro
        self.cor_text_secondary = "#637381" # Texto secundário
        self.cor_success = "#14A44D"     # Verde sucesso
        self.cor_warning = "#E4A11B"     # Amarelo aviso
        self.cor_error = "#DC4C64"       # Vermelho erro
        self.cor_card_bg = "#FFFFFF"     # Branco para cartões
        self.cor_border = "#E2E8F0"      # Borda leve
        self.cor_input_bg = "#F8FAFC"    # Fundo para inputs
        
        # Configurar tema da janela
        self.configure(fg_color=self.cor_bg)
        
        # Manter registro das instâncias
        self.instancias = []
        
        # Pasta para armazenar os diretórios das instâncias
        self.instances_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "instancias")
        if not os.path.exists(self.instances_dir):
            os.makedirs(self.instances_dir)
        
        # Opções de unidade padrão (OPM)
        self.opcoes_unidades = [
            "303017500 - CMB SEC DISTR",
            "303017400 - CMB SEC RECEB",
            "303016500 - CMB 3 - RESUMO"
        ]
        
        # Configuração da interface
        self.setup_interface()
        
        # Verificação periódica das instâncias
        self.verificar_instancias_periodico()
        
        # Carregar instâncias existentes
        self.carregar_instancias_existentes()
        
        # Registrar método para tratar o fechamento da janela
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_interface(self):
        # Layout principal em grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)  # Cabeçalho
        self.grid_rowconfigure(1, weight=1)  # Conteúdo
        self.grid_rowconfigure(2, weight=0)  # Barra de status
        
        # === Cabeçalho ===
        self.frame_header = ctk.CTkFrame(
            self,
            fg_color=self.cor_card_bg,
            corner_radius=10,
            border_width=1,
            border_color=self.cor_border
        )
        self.frame_header.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        self.frame_header.grid_columnconfigure(0, weight=1)
        
        self.label_titulo = ctk.CTkLabel(
            self.frame_header,
            text="AUTOCOFIN HUB",
            font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"),
            text_color=self.cor_accent
        )
        self.label_titulo.grid(row=0, column=0, padx=20, pady=(15, 0), sticky="w")
        
        self.label_subtitulo = ctk.CTkLabel(
            self.frame_header,
            text="Gerenciador de Múltiplas Instâncias",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color=self.cor_text_secondary
        )
        self.label_subtitulo.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")
        
        # === Conteúdo Principal ===
        self.frame_content = ctk.CTkFrame(
            self,
            fg_color="transparent"
        )
        self.frame_content.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        self.frame_content.grid_columnconfigure(0, weight=1)
        self.frame_content.grid_columnconfigure(1, weight=2)
        self.frame_content.grid_rowconfigure(0, weight=1)
        
        # === Painel esquerdo - nova instância ===
        self.frame_nova_instancia = ctk.CTkFrame(
            self.frame_content,
            fg_color=self.cor_card_bg,
            corner_radius=10,
            border_width=1,
            border_color=self.cor_border
        )
        self.frame_nova_instancia.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)
        self.frame_nova_instancia.grid_columnconfigure(0, weight=1)
        
        # Título do painel
        self.label_nova_instancia = ctk.CTkLabel(
            self.frame_nova_instancia,
            text="Adicionar Nova Instância",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color=self.cor_text
        )
        self.label_nova_instancia.grid(row=0, column=0, padx=20, pady=(20, 15), sticky="w")
        
        # Formulário para nova instância
        self.frame_form = ctk.CTkScrollableFrame(
            self.frame_nova_instancia,
            fg_color="transparent",
            scrollbar_button_color=self.cor_accent,
            scrollbar_button_hover_color="#2C5FAA"
        )
        self.frame_form.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        self.frame_form.grid_columnconfigure(1, weight=1)
        
        # Nome da instância
        ctk.CTkLabel(
            self.frame_form,
            text="Nome:",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color=self.cor_text
        ).grid(row=0, column=0, padx=(0, 10), pady=10, sticky="w")
        
        self.entry_nome = ctk.CTkEntry(
            self.frame_form,
            placeholder_text="Ex: COFIN OPM1",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            height=36,
            fg_color=self.cor_input_bg,
            border_color=self.cor_border
        )
        self.entry_nome.grid(row=0, column=1, sticky="ew", pady=10)
        
        # CPF
        ctk.CTkLabel(
            self.frame_form,
            text="CPF:",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color=self.cor_text
        ).grid(row=1, column=0, padx=(0, 10), pady=10, sticky="w")
        
        self.entry_cpf = ctk.CTkEntry(
            self.frame_form,
            placeholder_text="Apenas números",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            height=36,
            fg_color=self.cor_input_bg,
            border_color=self.cor_border
        )
        self.entry_cpf.grid(row=1, column=1, sticky="ew", pady=10)
        
        # Senha
        ctk.CTkLabel(
            self.frame_form,
            text="Senha:",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color=self.cor_text
        ).grid(row=2, column=0, padx=(0, 10), pady=10, sticky="w")
        
        self.entry_senha = ctk.CTkEntry(
            self.frame_form,
            placeholder_text="Senha do COFIN",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            height=36,
            fg_color=self.cor_input_bg,
            border_color=self.cor_border,
            show="•"
        )
        self.entry_senha.grid(row=2, column=1, sticky="ew", pady=10)
        
        # Unidade (OPM)
        ctk.CTkLabel(
            self.frame_form,
            text="Unidade:",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color=self.cor_text
        ).grid(row=3, column=0, padx=(0, 10), pady=10, sticky="w")
        
        self.unidade_var = ctk.StringVar(value=self.opcoes_unidades[0])
        self.combo_unidade = ctk.CTkOptionMenu(
            self.frame_form,
            values=self.opcoes_unidades,
            variable=self.unidade_var,
            font=ctk.CTkFont(family="Segoe UI", size=14),
            fg_color=self.cor_accent,
            button_color=self.cor_accent,
            button_hover_color="#2C5FAA",
            dropdown_font=ctk.CTkFont(family="Segoe UI", size=14),
            dropdown_fg_color=self.cor_card_bg,
            dropdown_hover_color="#EBF5FF",
            height=36,
            width=300,
            dynamic_resizing=True
        )
        self.combo_unidade.grid(row=3, column=1, sticky="w", pady=10)
        
        # Configurações adicionais
        self.frame_configs = ctk.CTkFrame(
            self.frame_form,
            fg_color="#F0F6FF",
            corner_radius=8,
            border_width=1,
            border_color="#BBDEFB"
        )
        self.frame_configs.grid(row=4, column=0, columnspan=2, sticky="ew", pady=15)
        self.frame_configs.grid_columnconfigure(0, weight=1)
        
        # Opção de alertar erros
        self.var_alertar_erros = ctk.BooleanVar(value=True)
        self.check_alertar_erros = ctk.CTkCheckBox(
            self.frame_configs,
            text="Alertar sobre erros durante a movimentação",
            variable=self.var_alertar_erros,
            font=ctk.CTkFont(family="Segoe UI", size=14),
            checkbox_width=22,
            checkbox_height=22,
            border_width=2,
            hover_color=self.cor_accent,
            fg_color=self.cor_accent
        )
        self.check_alertar_erros.grid(row=0, column=0, padx=15, pady=10, sticky="w")
        
        # Arquivo Excel com números de série
        ctk.CTkLabel(
            self.frame_form,
            text="Planilha:",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color=self.cor_text
        ).grid(row=5, column=0, padx=(0, 10), pady=10, sticky="w")
        
        self.frame_planilha = ctk.CTkFrame(
            self.frame_form,
            fg_color="transparent"
        )
        self.frame_planilha.grid(row=5, column=1, sticky="ew", pady=10)
        self.frame_planilha.grid_columnconfigure(0, weight=1)
        
        self.entry_planilha = ctk.CTkEntry(
            self.frame_planilha,
            placeholder_text="Selecione o arquivo Excel",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            height=36,
            fg_color=self.cor_input_bg,
            border_color=self.cor_border
        )
        self.entry_planilha.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        
        self.btn_selecionar_planilha = ctk.CTkButton(
            self.frame_planilha,
            text="Selecionar",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            fg_color="#64748B",
            hover_color="#475569",
            height=36,
            width=120,
            command=self.selecionar_planilha
        )
        self.btn_selecionar_planilha.grid(row=0, column=1)
        
        # Informação sobre limite de instâncias
        self.label_info_limite = ctk.CTkLabel(
            self.frame_form,
            text="⚠️ Limite: máximo de 5 instâncias simultâneas",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color=self.cor_warning
        )
        self.label_info_limite.grid(row=6, column=0, columnspan=2, pady=(15, 0), sticky="w")
        
        # Botão para adicionar instância
        self.btn_adicionar = ctk.CTkButton(
            self.frame_nova_instancia,
            text="INICIAR NOVA INSTÂNCIA",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            fg_color=self.cor_accent,
            hover_color="#2C5FAA",
            corner_radius=8,
            height=50,
            command=self.adicionar_instancia
        )
        self.btn_adicionar.grid(row=2, column=0, padx=20, pady=20, sticky="ew")
        
        # === Painel direito - instâncias ativas ===
        self.frame_instancias = ctk.CTkFrame(
            self.frame_content,
            fg_color=self.cor_card_bg,
            corner_radius=10,
            border_width=1,
            border_color=self.cor_border
        )
        self.frame_instancias.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=0)
        self.frame_instancias.grid_columnconfigure(0, weight=1)
        self.frame_instancias.grid_rowconfigure(1, weight=1)
        
        # Título do painel
        self.label_instancias_ativas = ctk.CTkLabel(
            self.frame_instancias,
            text="Instâncias Ativas",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color=self.cor_text
        )
        self.label_instancias_ativas.grid(row=0, column=0, padx=20, pady=(20, 15), sticky="w")
        
        # Frame com scroll para listar instâncias
        self.scroll_instancias = ctk.CTkScrollableFrame(
            self.frame_instancias,
            fg_color="transparent",
            scrollbar_button_color=self.cor_accent,
            scrollbar_button_hover_color="#2C5FAA"
        )
        self.scroll_instancias.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        self.scroll_instancias.grid_columnconfigure(0, weight=1)
        
        # === Barra de Status ===
        self.frame_status = ctk.CTkFrame(
            self,
            height=30,
            fg_color="#F1F5F9",
            border_width=1,
            border_color=self.cor_border
        )
        self.frame_status.grid(row=2, column=0, sticky="ew", padx=20, pady=(10, 20))
        self.frame_status.grid_propagate(False)  # Manter altura fixa
        self.frame_status.grid_columnconfigure(0, weight=1)
        
        # Status atual
        self.label_status = ctk.CTkLabel(
            self.frame_status,
            text="Hub iniciado | 0 instâncias em execução",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=self.cor_text_secondary
        )
        self.label_status.grid(row=0, column=0, padx=15, pady=5, sticky="w")

    def selecionar_planilha(self):
        """Abre diálogo para selecionar a planilha Excel"""
        arquivo = filedialog.askopenfilename(
            title="Selecionar planilha de números de série",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            self.entry_planilha.delete(0, "end")
            self.entry_planilha.insert(0, arquivo)
    
    def adicionar_instancia(self):
        """Cria e inicia uma nova instância do Autocofin"""
        # Validar entradas
        nome = self.entry_nome.get().strip()
        cpf = self.entry_cpf.get().strip()
        senha = self.entry_senha.get().strip()
        unidade = self.unidade_var.get()
        planilha = self.entry_planilha.get().strip()
        alertar_erros = self.var_alertar_erros.get()
        
        # Validações
        if not nome:
            messagebox.showerror("Erro", "Informe um nome para a instância")
            return
        
        if not cpf.isdigit() or len(cpf) != 11:
            messagebox.showerror("Erro", "CPF deve conter exatamente 11 dígitos")
            return
        
        if not senha:
            messagebox.showerror("Erro", "Informe a senha")
            return
        
        if not planilha or not os.path.exists(planilha):
            messagebox.showerror("Erro", "Selecione um arquivo Excel válido")
            return
        
        # Verificar limite de instâncias (máximo 5)
        instancias_ativas = [i for i in self.instancias if i.get("status") == "Em execução"]
        if len(instancias_ativas) >= 5:
            messagebox.showerror("Erro", "Limite de 5 instâncias simultâneas atingido.\nEncerre alguma instância antes de adicionar uma nova.")
            return
        
        # Criar diretório para a instância
        instance_id = f"{nome.replace(' ', '_')}_{str(uuid.uuid4())[:8]}"
        instance_dir = os.path.join(self.instances_dir, instance_id)
        
        try:
            # Criar diretório
            os.makedirs(instance_dir)
            
            # Copiar arquivos necessários
            shutil.copy(
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "main.py"),
                os.path.join(instance_dir, "main.py")
            )
            shutil.copy(
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "sistemacofin.py"),
                os.path.join(instance_dir, "sistemacofin.py")
            )
            
            # Criar arquivo .env com credenciais
            with open(os.path.join(instance_dir, ".env"), "w") as f:
                f.write(f"CPF={cpf}\n")
                f.write(f"SENHA={senha}\n")
                f.write(f"UNIDADE={unidade}\n")
                f.write(f"ALERTAR_ERROS={'1' if alertar_erros else '0'}\n")
            
            # Copiar planilha
            shutil.copy(planilha, os.path.join(instance_dir, "nserie.xlsx"))
            
            # No método adicionar_instancia, antes de iniciar o processo, adicione este código 
            # para criar um arquivo main_wrapper.py que captura erros:

            wrapper_content = """
import os
import sys
import traceback
from datetime import datetime

# Configurar log de erros
log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error_log.txt")

try:
    # Tentar importar as dependências necessárias
    try:
        import dotenv
    except ImportError:
        with open(log_file, "a") as f:
            f.write(f"\\n{datetime.now()} - ERRO: Módulo python-dotenv não encontrado. Execute 'pip install python-dotenv'\\n")
        sys.exit(1)
        
    # Executar o script principal com captura de erros
    try:
        import main
    except Exception as e:
        with open(log_file, "a") as f:
            f.write(f"\\n{datetime.now()} - ERRO AO INICIAR:\\n")
            f.write(traceback.format_exc())
        sys.exit(1)
except Exception as e:
    with open(log_file, "a") as f:
        f.write(f"\\n{datetime.now()} - ERRO CRÍTICO:\\n")
        f.write(str(e) + "\\n")
    sys.exit(1)
"""

            # Escrever o wrapper no diretório da instância
            with open(os.path.join(instance_dir, "main_wrapper.py"), "w") as f:
                f.write(wrapper_content)

            # Modificar comando para usar o wrapper
            comando = ["python", "main_wrapper.py"]
            
            # Verificar se o main.py foi modificado corretamente para aceitar parâmetros
            import re
            main_path = os.path.join(instance_dir, "main.py")
            sistemacofin_path = os.path.join(instance_dir, "sistemacofin.py")

            # Verificar e atualizar main.py para aceitar argumentos e dotenv
            with open(main_path, "r", encoding="utf-8") as f:
                main_content = f.read()

            # Verificar se as importações necessárias estão presentes
            if "import dotenv" not in main_content:
                # Adicionar as importações no início do arquivo
                import_code = """
import sys
import dotenv
import os

# Carregar variáveis de ambiente
dotenv_path = ".env"
for arg in sys.argv[1:]:
    if arg.startswith("--env="):
        dotenv_path = arg.split("=", 1)[1]
        break

dotenv.load_dotenv(dotenv_path)
"""
    
                # Encontrar um bom ponto para inserir o código (após os imports)
                pattern = r"(import.*?\n\n)"
                match = re.search(pattern, main_content, re.DOTALL)
                if match:
                    pos = match.end()
                    main_content = main_content[:pos] + import_code + main_content[pos:]
                else:
                    # Inserir no início se não encontrar um local adequado
                    main_content = import_code + main_content
        
                # Atualizar o título da janela para aceitar parâmetros de linha de comando
                title_pattern = r"self\.title\(['\"].*?['\"]\)"
                if re.search(title_pattern, main_content):
                    modified_title = """
        # Verificar argumentos de linha de comando para título personalizado
        titulo = "Autocofin | Sistema de Automação COFIN"
        for arg in sys.argv[1:]:
            if arg.startswith("--title="):
                titulo = f"{titulo} - {arg.split('=', 1)[1]}"
                
        self.title(titulo)
        """
                    main_content = re.sub(title_pattern, modified_title, main_content)
    
                # Salvar o arquivo modificado
                with open(main_path, "w", encoding="utf-8") as f:
                    f.write(main_content)

            # No método adicionar_instancia, após copiar os arquivos e antes de executar:
            import re

            # Verificar sistemacofin.py para garantir que tem todas as importações necessárias
            sistemacofin_path = os.path.join(instance_dir, "sistemacofin.py")

            with open(sistemacofin_path, "r", encoding="utf-8") as f:
                sistemacofin_content = f.read()

            # Verificar se 'import sys' está presente
            if "import sys" not in sistemacofin_content:
                # Adicionar a importação no início do arquivo
                sistemacofin_content = "import sys\n" + sistemacofin_content
                
                # Salvar o arquivo modificado
                with open(sistemacofin_path, "w", encoding="utf-8") as f:
                    f.write(sistemacofin_content)

            # MODIFICAÇÃO 1: Executar main.py diretamente em vez de usar wrapper
            # Criar um arquivo .bat que executará o Python com main.py
            bat_content = f"""@echo off
echo Iniciando Autocofin - {nome}
cd /d "%~dp0"
python main.py --title="{nome}"
if %ERRORLEVEL% NEQ 0 (
    echo Erro ao executar. Pressione qualquer tecla para sair...
    pause >nul
)
"""
            # Escrever o arquivo .bat no diretório da instância
            with open(os.path.join(instance_dir, "iniciar.bat"), "w") as f:
                f.write(bat_content)
            
            # MODIFICAÇÃO 2: Usar shellexecute para garantir que a janela seja visível
            import subprocess
            
            # Iniciar o processo usando o arquivo .bat
            processo = subprocess.Popen(
                os.path.join(instance_dir, "iniciar.bat"),
                shell=True,  # Importante: usar shell=True para abrir uma janela visível
                creationflags=subprocess.CREATE_NEW_CONSOLE,  # Criar um novo console
                cwd=instance_dir
            )
            
            # Adicionar à lista de instâncias
            self.instancias.append({
                "id": instance_id,
                "nome": nome,
                "diretorio": instance_dir,
                "processo": processo,
                "pid": processo.pid,
                "unidade": unidade,
                "alertar_erros": alertar_erros,
                "cpf": cpf[:3] + "..." + cpf[-2:],  # CPF parcial por segurança
                "status": "Em execução",
                "iniciado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            })
            
            # Atualizar interface
            self.atualizar_lista_instancias()
            self.atualizar_status_bar()
            
            # Limpar formulário
            self.entry_nome.delete(0, "end")
            self.entry_cpf.delete(0, "end")
            self.entry_senha.delete(0, "end")
            self.entry_planilha.delete(0, "end")
            
            messagebox.showinfo("Sucesso", f"Instância '{nome}' criada e iniciada com sucesso!")
            
        except Exception as e:
            # Limpar diretório em caso de erro
            if os.path.exists(instance_dir):
                try:
                    shutil.rmtree(instance_dir)
                except:
                    pass
            messagebox.showerror("Erro", f"Falha ao criar instância: {str(e)}")
    
    def verificar_instancias_periodico(self):
        """Verifica o status das instâncias periodicamente"""
        alteracoes = False
        
        for instancia in self.instancias:
            try:
                if "processo" in instancia and instancia["processo"]:
                    # Verificar se o processo ainda está em execução
                    status_anterior = instancia.get("status")
                    exit_code = instancia["processo"].poll()
                    if exit_code is not None and status_anterior == "Em execução":
                        # Processo encerrado
                        instancia["status"] = "Encerrada"
                        instancia["encerrado_em"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        alteracoes = True
            except:
                # Em caso de erro ao verificar processo, marcar como desconhecido
                if instancia.get("status") != "Desconhecido":
                    instancia["status"] = "Desconhecido"
                    alteracoes = True
        
        # Atualizar interface apenas se houver alterações
        if alteracoes:
            self.atualizar_lista_instancias()
            self.atualizar_status_bar()
        
        # Verificar novamente após 5 segundos
        self.after(5000, self.verificar_instancias_periodico)
    
    def atualizar_lista_instancias(self):
        """Atualiza a lista de instâncias na interface"""
        # Limpar lista atual
        for widget in self.scroll_instancias.winfo_children():
            widget.destroy()
        
        # Verificar se há instâncias
        if not self.instancias:
            label_vazio = ctk.CTkLabel(
                self.scroll_instancias,
                text="Nenhuma instância em execução.\n\nUtilize o formulário ao lado para adicionar uma nova instância.",
                font=ctk.CTkFont(family="Segoe UI", size=14),
                text_color=self.cor_text_secondary
            )
            label_vazio.pack(pady=30)
            return
        
        # Adicionar cada instância à lista
        for i, instancia in enumerate(self.instancias):
            # Card para a instância
            frame_instancia = ctk.CTkFrame(
                self.scroll_instancias,
                fg_color=self.cor_input_bg,
                corner_radius=8,
                border_width=1,
                border_color=self.cor_border
            )
            frame_instancia.grid(row=i, column=0, sticky="ew", pady=8)
            frame_instancia.grid_columnconfigure(1, weight=1)
            
            # Status com cor
            status = instancia.get("status", "Desconhecido")
            status_cor = self.cor_success if status == "Em execução" else self.cor_error
            
            # Número/ID da instância
            ctk.CTkLabel(
                frame_instancia,
                text=f"#{i+1}",
                font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
                text_color=self.cor_accent,
                width=40
            ).grid(row=0, column=0, rowspan=2, padx=(15, 5), pady=(15, 0))
            
            # Nome da instância
            ctk.CTkLabel(
                frame_instancia,
                text=instancia["nome"],
                font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
                text_color=self.cor_text
            ).grid(row=0, column=1, sticky="w", padx=5, pady=(15, 0))
            
            # Unidade e CPF parcial
            info_text = f"Unidade: {instancia.get('unidade', '-')} | CPF: {instancia.get('cpf', '-')}"
            ctk.CTkLabel(
                frame_instancia,
                text=info_text,
                font=ctk.CTkFont(family="Segoe UI", size=12),
                text_color=self.cor_text_secondary
            ).grid(row=1, column=1, sticky="w", padx=5, pady=(0, 5))
            
            # Data/hora de início e fim
            hora_texto = f"Iniciado: {instancia.get('iniciado_em', '-')}"
            if "encerrado_em" in instancia:
                hora_texto += f" | Encerrado: {instancia['encerrado_em']}"
            
            ctk.CTkLabel(
                frame_instancia,
                text=hora_texto,
                font=ctk.CTkFont(family="Segoe UI", size=12),
                text_color=self.cor_text_secondary
            ).grid(row=2, column=1, sticky="w", padx=5, pady=(0, 15))
            
            # Status da instância
            ctk.CTkLabel(
                frame_instancia,
                text=status,
                font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
                text_color=status_cor
            ).grid(row=0, column=2, rowspan=2, padx=15)
            
            # Botões de ação
            frame_botoes = ctk.CTkFrame(
                frame_instancia,
                fg_color="transparent"
            )
            frame_botoes.grid(row=0, column=3, rowspan=3, padx=15, pady=15)
            
            # Botão para encerrar instância (apenas se estiver em execução)
            if status == "Em execução":
                btn_encerrar = ctk.CTkButton(
                    frame_botoes,
                    text="Encerrar",
                    font=ctk.CTkFont(family="Segoe UI", size=13),
                    fg_color=self.cor_error,
                    hover_color="#C62350",
                    corner_radius=6,
                    height=30,
                    width=90,
                    command=lambda idx=i: self.encerrar_instancia(idx)
                )
                btn_encerrar.pack(pady=5)
            
            # Botão para abrir pasta da instância
            btn_pasta = ctk.CTkButton(
                frame_botoes,
                text="Abrir Pasta",
                font=ctk.CTkFont(family="Segoe UI", size=13),
                fg_color="#64748B",
                hover_color="#475569",
                corner_radius=6,
                height=30,
                width=90,
                command=lambda dir=instancia["diretorio"]: os.startfile(dir)
            )
            btn_pasta.pack(pady=5)
            
            # Botão para excluir a instância (apenas se estiver encerrada)
            if status != "Em execução":
                btn_excluir = ctk.CTkButton(
                    frame_botoes,
                    text="Excluir",
                    font=ctk.CTkFont(family="Segoe UI", size=13),
                    fg_color="#FF6B6B",
                    hover_color="#FF5252",
                    corner_radius=6,
                    height=30,
                    width=90,
                    command=lambda idx=i: self.excluir_instancia(idx)
                )
                btn_excluir.pack(pady=5)
    
    def encerrar_instancia(self, indice):
        """Encerra uma instância específica"""
        if indice < len(self.instancias):
            instancia = self.instancias[indice]
            
            if instancia.get("status") != "Em execução":
                return
                
            try:
                # Atualizar status antes de iniciar o encerramento
                instancia["status"] = "Encerrando..."
                instancia["encerrado_em"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                
                # Atualizar interface imediatamente
                self.atualizar_lista_instancias()
                self.atualizar_status_bar()
                self.update()  # Forçar atualização da interface
                
                # 1. Matar processos associados por PID
                if "processo" in instancia and instancia["processo"] and instancia["processo"].pid:
                    try:
                        # Matar processo principal e subprocessos usando taskkill
                        subprocess.call(f'taskkill /F /T /PID {instancia["processo"].pid}', shell=True)
                    except Exception as e:
                        print(f"Erro ao encerrar processo por PID: {e}")
                
                # 2. Buscar e matar janelas do Edge associadas à instância
                # Procurar nos arquivos de log ou por título específico
                try:
                    # Encerrar qualquer Edge aberto pela instância
                    subprocess.call(
                        f'taskkill /F /FI "WINDOWTITLE eq *{instancia["nome"]}*" /IM msedge.exe', 
                        shell=True
                    )
                except Exception as e:
                    print(f"Erro ao encerrar navegador: {e}")
                
                # 3. Matar python associado a esta instância
                try:
                    # Matar processos Python executando o script da instância
                    instance_name = os.path.basename(instancia["diretorio"])
                    # Usar tasklist para encontrar processos Python e verificar se o caminho contém o nome da instância
                    subprocess.call(
                        f'wmic process where "name=\'python.exe\' and CommandLine like \'%{instance_name}%\'" call terminate',
                        shell=True
                    )
                except Exception as e:
                    print(f"Erro ao encerrar Python: {e}")
                
                # Atualizar status após tentativas de encerramento
                instancia["status"] = "Encerrada"
                
                # Atualizar interface
                self.atualizar_lista_instancias()
                self.atualizar_status_bar()
                
                messagebox.showinfo("Instância Encerrada", f"A instância '{instancia['nome']}' foi encerrada.")
                    
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível encerrar a instância: {str(e)}")
    
    def atualizar_status_bar(self):
        """Atualiza a barra de status com informações sobre as instâncias"""
        instancias_ativas = len([i for i in self.instancias if i.get("status") == "Em execução"])
        self.label_status.configure(text=f"Hub rodando | {instancias_ativas} instâncias em execução")
        
        # Atualizar estado do botão Adicionar
        if instancias_ativas >= 5:
            self.btn_adicionar.configure(
                state="disabled",
                fg_color="#A0AEC0",
                text="LIMITE DE 5 INSTÂNCIAS ATINGIDO"
            )
            self.label_info_limite.configure(
                text="⚠️ Limite atingido: máximo de 5 instâncias simultâneas",
                text_color=self.cor_error
            )
        else:
            self.btn_adicionar.configure(
                state="normal",
                fg_color=self.cor_accent,
                text="INICIAR NOVA INSTÂNCIA"
            )
            self.label_info_limite.configure(
                text=f"⚠️ Limite: máximo de 5 instâncias simultâneas ({instancias_ativas}/5 em uso)",
                text_color=self.cor_warning
            )
    
    def carregar_instancias_existentes(self):
        """Carrega informações de instâncias que podem estar em execução"""
        if not os.path.exists(self.instances_dir):
            return
            
        dirs = [d for d in os.listdir(self.instances_dir) 
               if os.path.isdir(os.path.join(self.instances_dir, d))]
               
        for dir_name in dirs:
            dir_path = os.path.join(self.instances_dir, dir_name)
            
            # Verificar se diretório contém arquivos necessários
            if not os.path.exists(os.path.join(dir_path, "main.py")) or \
               not os.path.exists(os.path.join(dir_path, ".env")):
                continue
                
            # Ler arquivo .env para extrair informações
            env_data = {}
            try:
                with open(os.path.join(dir_path, ".env"), "r") as f:
                    for line in f:
                        if "=" in line:
                            key, value = line.strip().split("=", 1)
                            env_data[key] = value
            except:
                continue
                
            # Nome da instância do nome do diretório
            nome = dir_name.split("_")[0].replace("_", " ")
            
            # Adicionar à lista
            self.instancias.append({
                "id": dir_name,
                "nome": nome,
                "diretorio": dir_path,
                "processo": None,
                "unidade": env_data.get("UNIDADE", "-"),
                "alertar_erros": env_data.get("ALERTAR_ERROS", "1") == "1",
                "cpf": env_data.get("CPF", "")[:3] + "..." + env_data.get("CPF", "")[-2:] if "CPF" in env_data else "-",
                "status": "Encerrada",
                "iniciado_em": "Desconhecido"
            })
            
        # Atualizar interface
        if self.instancias:
            self.atualizar_lista_instancias()
            self.atualizar_status_bar()
    
    def on_closing(self):
        """Método chamado quando a janela é fechada"""
        # Verificar se há instâncias ativas
        instancias_ativas = [i for i in self.instancias if i.get("status") == "Em execução"]
        
        if instancias_ativas:
            resp = messagebox.askyesnocancel(
                "Encerrar instâncias", 
                f"Existem {len(instancias_ativas)} instâncias em execução.\n\n"
                f"• Sim: Encerrar todas as instâncias e sair\n"
                f"• Não: Sair sem encerrar as instâncias\n"
                f"• Cancelar: Voltar ao Hub"
            )
            
            if resp is None:  # Cancelar
                return
                
            if resp:  # Sim, encerrar todas
                for i in range(len(self.instancias)):
                    if self.instancias[i].get("status") == "Em execução":
                        self.encerrar_instancia(i)
                
                # Aguardar um pouco para os processos serem encerrados
                self.update()
                time.sleep(1)
        
        # Destruir a janela e encerrar aplicação
        self.destroy()

    def excluir_instancia(self, indice):
        """Exclui uma instância da lista e remove sua pasta"""
        if indice < len(self.instancias):
            instancia = self.instancias[indice]
            
            # Verificar se a instância está encerrada
            if instancia.get("status") == "Em execução":
                messagebox.showerror(
                    "Erro", 
                    "Esta instância está em execução!\nEncerre a instância antes de excluí-la."
                )
                return
            
            # Confirmar exclusão
            if not messagebox.askyesno(
                "Confirmar Exclusão", 
                f"Tem certeza que deseja excluir a instância '{instancia['nome']}'?\n\n"
                f"Esta ação excluirá permanentemente a pasta com todos os arquivos da instância."
            ):
                return
            
            # Encaminhar para o tratamento de erros
            try:
                # Excluir pasta da instância
                shutil.rmtree(instancia["diretorio"], ignore_errors=True)
                
                # Remover da lista
                del self.instancias[indice]
                
                # Atualizar interface
                self.atualizar_lista_instancias()
                self.atualizar_status_bar()
                
                messagebox.showinfo("Sucesso", f"Instância '{instancia['nome']}' excluída com sucesso!")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível excluir a instância: {str(e)}\n\n"
                                          f"Tente fechar qualquer programa que possa estar usando "
                                          f"arquivos na pasta da instância e tente novamente.")

# Iniciar aplicação quando executado diretamente
if __name__ == "__main__":
    app = AutocofinHub()
    app.mainloop()