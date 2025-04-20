
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
            f.write(f"\n{datetime.now()} - ERRO: Módulo python-dotenv não encontrado. Execute 'pip install python-dotenv'\n")
        sys.exit(1)
        
    # Executar o script principal com captura de erros
    try:
        import main
    except Exception as e:
        with open(log_file, "a") as f:
            f.write(f"\n{datetime.now()} - ERRO AO INICIAR:\n")
            f.write(traceback.format_exc())
        sys.exit(1)
except Exception as e:
    with open(log_file, "a") as f:
        f.write(f"\n{datetime.now()} - ERRO CRÍTICO:\n")
        f.write(str(e) + "\n")
    sys.exit(1)
