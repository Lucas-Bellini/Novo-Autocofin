import os
import time
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from webdriver_manager.microsoft import EdgeChromiumDriverManager

# ========== Variáveis globais para a GUI ==========
IS_RUNNING = False
IS_FINISHED = False
IS_PAUSED = False  # Nova variável para controlar o estado de pausa
ALERTAR_ERROS = True  # Nova variável para controlar alertas de erro
STATUS = ""
movimentados = 0
total_itens = 0
lista_movimentados = []
lista_erros = []
item_atual = None

# Variáveis para controle de nova tentativa
RETRY_REQUESTED = False  # Indica se uma nova tentativa foi solicitada
RETRY_PENDING = False    # Indica que está aguardando resposta do usuário
RETRY_ITEM = None        # Armazena o item que está esperando por nova tentativa

# Variável global para receber a opção escolhida na GUI
OPCAO_DESEJADA = None

# Para rastrear tempo
START_TIME = None
times_of_processing = []  # guardamos timestamps a cada item finalizado

# Carregar variáveis de ambiente
load_dotenv()
CPF = os.getenv("CPF")
SENHA = os.getenv("SENHA")

# Caminhos
base_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(base_dir, "nserie.xlsx")

# Verificar se o arquivo da planilha existe
def verificar_planilha():
    """
    Verifica se a planilha existe e retorna informações sobre ela.
    """
    global total_itens  # Adicionar esta linha
    
    if not os.path.exists(file_path):
        total_itens = 0
        return False, "Arquivo nserie.xlsx não encontrado", 0
    
    try:
        df = pd.read_excel(file_path, sheet_name=0)
        total_itens = len(df)
        return True, f"Planilha encontrada: {total_itens} itens para processar", total_itens
    except Exception as e:
        total_itens = 0
        return False, f"Erro ao ler a planilha: {str(e)}", 0

# Mantemos uma referência global do driver
driver_instance = None

def click_with_overlay_wait(driver, element, wait, timeout=30):
    """
    Espera o overlay desaparecer e tenta clicar no elemento.
    Se clique normal falhar, tenta via JavaScript.
    """
    wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".block-ui-wrapper.block-ui-main.active")))
    try:
        element.click()
    except:
        driver.execute_script("arguments[0].click();", element)

def create_driver(headless=False):
    """
    Cria e retorna uma instância do driver Edge com as opções configuradas.
    """
    options = webdriver.EdgeOptions()
    options.add_argument("--start-maximized")
    
    if headless:
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        # Evitar detecção de headless
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
    
    # Usar webdriver-manager para gerenciar o driver automaticamente
    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)
    
    # Tentar disfarçar automação (se não headless)
    if not headless:
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    return driver

def run_sistema_cofin(headless=False, captcha_timeout=60):
    """
    Função principal que executa todo o fluxo de automação Selenium.
    """
    global IS_RUNNING, IS_FINISHED, STATUS
    global movimentados, lista_movimentados, lista_erros
    global OPCAO_DESEJADA
    global START_TIME, times_of_processing
    global driver_instance
    global item_atual, RETRY_ITEM, RETRY_PENDING, RETRY_REQUESTED  # Declarar todas as variáveis globais aqui no início
    
    IS_RUNNING = True
    IS_FINISHED = False
    STATUS = "Iniciando..."
    movimentados = 0
    lista_movimentados.clear()
    lista_erros.clear()
    START_TIME = time.time()
    times_of_processing.clear()

    # Verificar existência da planilha
    planilha_ok, msg, total_itens = verificar_planilha()
    if not planilha_ok:
        STATUS = msg
        IS_RUNNING = False
        IS_FINISHED = True
        return False, msg

    # Ler planilha
    df = pd.read_excel(file_path, sheet_name=0)

    try:
        STATUS = "Iniciando navegador e preparando ambiente..."
        driver = create_driver(headless=headless)
        driver_instance = driver

        url = "https://cofin.sp.gov.br/#/menu/gestao-patrimonial/movimentacao-materiais/criar-movimentacao"
        driver.get(url)

        STATUS = "Preenchendo credenciais..."
        wait = WebDriverWait(driver, 15)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "txtCpf"))
        ).send_keys(CPF)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "Password"))
        ).send_keys(SENHA)

        STATUS = f"Aguardando resolução do CAPTCHA... ({captcha_timeout}s)"
        if not headless:
            print("🛑 Aguarde! Resolva o reCAPTCHA manualmente.")
        
        # Esperar login ser processado
        try:
            WebDriverWait(driver, captcha_timeout).until(lambda d: d.current_url == url)
            STATUS = "CAPTCHA resolvido! Continuando automação..."
        except Exception as e:
            STATUS = "Tempo excedido para resolver CAPTCHA ou login."
            if headless:
                driver.quit()
                IS_RUNNING = False
                IS_FINISHED = True
                return False, "Login não concluído em modo headless (timeout)"
            
            # Em modo não-headless, continuamos e damos mais 60 segundos
            try:
                WebDriverWait(driver, 60).until(lambda d: d.current_url == url)
                STATUS = "Login finalizado com sucesso!"
            except:
                STATUS = "Falha no login, encerrando."
                driver.quit()
                IS_RUNNING = False
                IS_FINISHED = True
                return False, "Login falhou mesmo após tempo adicional"

        # Abrir dropdown
        STATUS = "Abrindo o primeiro dropdown..."
        try:
            dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'p-dropdown')]")))
            click_with_overlay_wait(driver, dropdown, wait)
            time.sleep(1)
        except Exception as e:
            STATUS = f"Erro ao abrir o dropdown: {str(e)}"
            driver.quit()
            IS_RUNNING = False
            IS_FINISHED = True
            return False, STATUS

        # Identificar listbox
        STATUS = "Identificando listbox..."
        try:
            listbox = wait.until(EC.presence_of_element_located((By.XPATH, "//ul[@role='listbox']")))
        except Exception as e:
            STATUS = f"Erro ao localizar a listbox: {str(e)}"
            driver.quit()
            IS_RUNNING = False
            IS_FINISHED = True
            return False, STATUS

        # Selecionar item
        STATUS = "Procurando opção desejada..."

        # Item para forçar rolagem
        ultimo_item_texto = "303017400 - CMB SEC RECEB"

        # Usar opção padrão se nenhuma foi fornecida pela GUI
        if not OPCAO_DESEJADA:
            OPCAO_DESEJADA = "303017500 - CMB SEC DISTR"

        opcao_desejada_texto = OPCAO_DESEJADA
        STATUS = f"Procurando opção: {opcao_desejada_texto}"

        encontrado = False
        tentativas = 0
        max_tentativas = 10

        while not encontrado and tentativas < max_tentativas:
            opcoes_visiveis = driver.find_elements(By.XPATH, "//ul[@role='listbox']//li[@role='option']")
            ultimo_item = None
            opcao_final = None

            for opcao in opcoes_visiveis:
                texto_opcao = opcao.get_attribute("aria-label")
                
                if texto_opcao == ultimo_item_texto:
                    ultimo_item = opcao

                if texto_opcao == opcao_desejada_texto:
                    opcao_final = opcao
                    break

            if opcao_final:
                STATUS = f"Opção encontrada, selecionando: {opcao_desejada_texto}"
                driver.execute_script("arguments[0].scrollIntoView();", opcao_final)
                time.sleep(0.5)
                click_with_overlay_wait(driver, opcao_final, wait)
                encontrado = True
                break

            # Se encontrou o ultimo_item_texto, rola para forçar
            if ultimo_item:
                driver.execute_script("arguments[0].scrollIntoView();", ultimo_item)
                time.sleep(0.5)

            tentativas += 1
            STATUS = f"Procurando opção: tentativa {tentativas}/{max_tentativas}"

        if not encontrado:
            STATUS = f"A opção '{opcao_desejada_texto}' não foi encontrada após {tentativas} tentativas!"
            IS_RUNNING = False
            IS_FINISHED = True
            return False, STATUS

        STATUS = f"Preparação concluída. Iniciando processamento de {total_itens} itens..."
        time.sleep(1)

        # Processamento dos números de série da planilha
        motivo_inserido = False
        index = 0

        while index < len(df) and IS_RUNNING:
            # Se o script for interrompido manualmente
            if not IS_RUNNING:
                break
            
            # Verificar se está pausado
            while IS_PAUSED and IS_RUNNING:
                time.sleep(0.5)  # Aguarda enquanto estiver pausado
            
            # Se foi interrompido durante a pausa
            if not IS_RUNNING:
                break
            
            row = df.iloc[index]
            
            # Verificar se a coluna 'Nserie' existe
            if "Nserie" not in row:
                STATUS = "Formato de planilha incorreto! Verifique se há coluna 'Nserie'"
                IS_RUNNING = False
                IS_FINISHED = True
                return False, STATUS

            numero_serie = str(row["Nserie"]).strip()
            # Define o item atual para uso na interface
            item_atual = numero_serie
            STATUS = f"Processando: {numero_serie} ({index+1}/{total_itens})"
            
            try:
                serie_field = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.ID, "inputTextNSERIE"))
                )
                serie_field.clear()
                serie_field.send_keys(numero_serie)

                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "block-ui-wrapper")))
                pesquisar_button = WebDriverWait(driver, 60).until(
                    EC.element_to_be_clickable((By.ID, "btnPesquisar"))
                )
                click_with_overlay_wait(driver, pesquisar_button, wait)

                # Esperar resultados carregarem
                try:
                    WebDriverWait(driver, 60).until(
                        EC.presence_of_element_located((By.XPATH, "//p-table//table"))
                    )
                except:
                    STATUS = f"Erro: Resultados não carregaram para {numero_serie}"
                    lista_erros.append(numero_serie)
                    elapsed = time.time() - START_TIME
                    times_of_processing.append(elapsed)
                    continue

                if not motivo_inserido:
                    STATUS = "Configurando motivo da movimentação..."
                    
                    try:
                        dropdown_motivo = wait.until(
                            EC.element_to_be_clickable(
                                (By.XPATH, "//input[@id='dropdownMMOVIMENTACAO']/ancestor::div[contains(@class, 'p-dropdown')]")
                            )
                        )
                        click_with_overlay_wait(driver, dropdown_motivo, wait)
                        time.sleep(1)
                    except Exception as e:
                        STATUS = f"Erro ao abrir dropdown de motivo: {str(e)}"
                        lista_erros.append(numero_serie)
                        continue

                    try:
                        listbox_motivo = wait.until(
                            EC.presence_of_element_located(
                                (By.XPATH, "//div[contains(@class, 'p-dropdown-panel')]//ul[@role='listbox']")
                            )
                        )
                    except Exception as e:
                        STATUS = f"Erro ao localizar listbox de motivo: {str(e)}"
                        lista_erros.append(numero_serie)
                        continue

                    # Valores para localizar na listbox
                    valores_intermediarios = [
                        "COMPRA - RECURSO TESOURO", "CONVERSÃO", "DECISÃO JUDICIAL",
                        "DOAÇÃO DE PARTICULAR", "ENTRE OPM", "ERRO DE MOVIMENTACAO",
                        "EXTINÇÃO DE OPM", "FURTO", "INCORPORAÇÃO DA RECEITA FEDERAL",
                        "LOCAÇÃO", "OBSOLECÊNCIA"
                    ]
                    valor_final = "PROVIMENTO"

                    encontrado_mot = False
                    tenta_mot = 0
                    max_tentativas_mot = 30

                    time.sleep(1)

                    # Procurar o motivo na listbox
                    while not encontrado_mot and tenta_mot < max_tentativas_mot:
                        try:
                            opcoes_visiveis = driver.find_elements(
                                By.XPATH,
                                "//div[contains(@class, 'p-dropdown-panel')]//ul[@role='listbox']//li[@role='option']"
                            )
                            opcao_final_mot = None

                            for opcao in opcoes_visiveis:
                                try:
                                    texto_opcao = opcao.get_attribute("aria-label")

                                    if texto_opcao in valores_intermediarios:
                                        driver.execute_script("arguments[0].scrollIntoView();", opcao)
                                        time.sleep(0.1)

                                    if texto_opcao == valor_final:
                                        opcao_final_mot = opcao
                                        break
                                except:
                                    # Ignora elementos obsoletos
                                    pass

                            if opcao_final_mot:
                                STATUS = f"Motivo encontrado: {valor_final}"
                                driver.execute_script("arguments[0].scrollIntoView();", opcao_final_mot)
                                time.sleep(1)
                                click_with_overlay_wait(driver, opcao_final_mot, wait)
                                encontrado_mot = True
                                break

                            # Rolar para carregar mais opções
                            listbox_wrapper = driver.find_element(By.CLASS_NAME, "p-dropdown-items-wrapper")
                            driver.execute_script("arguments[0].scrollTop += 50;", listbox_wrapper)
                            time.sleep(0.1)
                            tenta_mot += 1
                            STATUS = f"Procurando motivo: tentativa {tenta_mot}/{max_tentativas_mot}"

                        except Exception as e:
                            STATUS = f"Erro ao procurar motivo: {str(e)}"
                            tenta_mot += 1
                            time.sleep(0.3)

                    if not encontrado_mot:
                        STATUS = f"O motivo '{valor_final}' não foi encontrado após {tenta_mot} tentativas!"
                        if not headless:
                            # Em modo não-headless, permitir intervenção manual
                            STATUS = "Aguardando seleção manual do motivo da movimentação..."
                            if lista_movimentados or lista_erros:  # Se já processamos algum item
                                # Assumir que agora está ok
                                motivo_inserido = True
                            else:
                                # Para o primeiro item, pausar para intervenção manual
                                time.sleep(10)  # Dar tempo para intervenção manual
                                motivo_inserido = True
                        else:
                            # Em headless, abortar
                            driver.quit()
                            IS_RUNNING = False
                            IS_FINISHED = True
                            return False, "Não foi possível selecionar o motivo em modo headless"
                    else:
                        motivo_inserido = True

                # Tentar movimentar o item
                max_tentativas_btn = 2
                sucesso_movimentacao = False

                for tentativa in range(max_tentativas_btn):
                    try:
                        STATUS = f"Executando movimentação para {numero_serie}..."
                        try:
                            acoes_button = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable(
                                    (By.XPATH, "//span[contains(@class, 'fa-arrow-circle-down')]")
                                )
                            )
                        except:
                            STATUS = f"Botão de movimentação não encontrado para {numero_serie}"
                            break

                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", acoes_button)
                        time.sleep(1)
                        click_with_overlay_wait(driver, acoes_button, wait)

                        # Aguarda operação completar
                        try:
                            # Esperar até 5 minutos pela movimentação
                            WebDriverWait(driver, 300).until(
                                EC.visibility_of_element_located((By.XPATH, "//div[@class='block-ui-spinner']"))
                            )
                            WebDriverWait(driver, 300).until(
                                EC.invisibility_of_element_located((By.XPATH, "//div[@class='block-ui-spinner']"))
                            )
                        except:
                            STATUS = f"Timeout ao aguardar movimentação de {numero_serie}"
                            # Continue mesmo assim

                        movimentados += 1
                        elapsed = time.time() - START_TIME
                        times_of_processing.append(elapsed)
                        
                        STATUS = f"✅ Movimentado: {numero_serie} ({movimentados}/{total_itens})"
                        sucesso_movimentacao = True
                        break

                    except Exception as e:
                        STATUS = f"Erro ao movimentar {numero_serie}: {str(e)}"
                        time.sleep(2)

                if not sucesso_movimentacao:
                    STATUS = f"❌ Falha ao movimentar {numero_serie}"
                    
                    # Em vez de adicionar diretamente à lista de erros, solicitar nova tentativa
                    RETRY_ITEM = numero_serie
                    RETRY_PENDING = True
                    RETRY_REQUESTED = False
                    
                    # Pausar e aguardar resposta do usuário
                    STATUS = f"Aguardando decisão do usuário para o item {numero_serie}..."
                    while RETRY_PENDING and IS_RUNNING:
                        time.sleep(0.5)
                    
                    # Verificar a resposta do usuário
                    if RETRY_REQUESTED:
                        # Continuar no mesmo índice para tentar este item novamente
                        STATUS = f"Tentando novamente o item {numero_serie}..."
                        continue
                    else:
                        # Adicionar à lista de erros e continuar para o próximo item
                        lista_erros.append(numero_serie)
                        elapsed = time.time() - START_TIME
                        times_of_processing.append(elapsed)
                        index += 1
                        continue

                lista_movimentados.append(numero_serie)
                time.sleep(2)
                index += 1

            except Exception as e:
                STATUS = f"Erro inesperado ao processar {numero_serie}: {str(e)}"
                
                # Solicitar nova tentativa para erros inesperados também
                RETRY_ITEM = numero_serie
                RETRY_PENDING = True
                RETRY_REQUESTED = False
                
                # Pausar e aguardar resposta do usuário
                STATUS = f"Aguardando decisão do usuário para o item {numero_serie}..."
                while RETRY_PENDING and IS_RUNNING:
                    time.sleep(0.5)
                
                # Verificar a resposta do usuário
                if RETRY_REQUESTED:
                    # Continuar no mesmo índice para tentar este item novamente
                    STATUS = f"Tentando novamente o item {numero_serie}..."
                    continue
                else:
                    # Adicionar à lista de erros e continuar para o próximo item
                    lista_erros.append(numero_serie)
                    elapsed = time.time() - START_TIME
                    times_of_processing.append(elapsed)
                    index += 1
                    continue
            
        # Gerar relatórios finais
        STATUS = f"Processamento concluído: {len(lista_movimentados)} sucesso, {len(lista_erros)} falhas"
        
        # Criar a pasta 'relatorios_excel' caso não exista
        pasta_relatorios = os.path.join(base_dir, "relatorios_excel")
        if not os.path.exists(pasta_relatorios):
            os.makedirs(pasta_relatorios)

        # Gerar nome com timestamp para os arquivos
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Caminhos para cada arquivo
        caminho_movimentados = os.path.join(pasta_relatorios, f"movimentados_{timestamp}.xlsx")
        caminho_erros = os.path.join(pasta_relatorios, f"erros_{timestamp}.xlsx")

        # Criar dataframes e salvar
        if lista_movimentados:
            df_movimentados = pd.DataFrame({"NSerie": lista_movimentados})
            df_movimentados.to_excel(caminho_movimentados, index=False)
            
        if lista_erros:
            df_erros = pd.DataFrame({"NSerie": lista_erros})
            df_erros.to_excel(caminho_erros, index=False)

        if not headless:
            STATUS = "Operação concluída. Navegador permanece aberto para verificação."
        else:
            driver.quit()

        return True, STATUS

    except Exception as e:
        STATUS = f"Erro inesperado: {str(e)}"
        return False, STATUS

    finally:
        IS_RUNNING = False
        IS_FINISHED = True


def force_stop_script():
    """
    Força a finalização do script sem fechar o navegador.
    """
    global IS_RUNNING, IS_FINISHED, STATUS
    IS_RUNNING = False
    IS_FINISHED = True
    STATUS = "Encerrado manualmente. Navegador mantido aberto."
    
    return STATUS

def pause_script():
    """
    Pausa a execução do script.
    """
    global IS_PAUSED, STATUS
    IS_PAUSED = True
    STATUS = "Automação pausada. Clique em Retomar para continuar."
    return STATUS

def resume_script():
    """
    Retoma a execução do script pausado.
    """
    global IS_PAUSED, STATUS
    IS_PAUSED = False
    STATUS = "Retomando automação..."
    return STATUS