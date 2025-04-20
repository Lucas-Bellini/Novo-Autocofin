import importlib
import subprocess
import sys

def verificar_dependencias():
    """Verifica se todas as bibliotecas necessárias estão instaladas."""
    
    dependencias = [
        "python-dotenv",
        "selenium",
        "customtkinter",
        "pandas",
        "openpyxl",
        "pillow",
        "webdriver_manager",
        "psutil"
    ]
    
    print("Verificando dependências necessárias:")
    
    faltando = []
    
    for dep in dependencias:
        try:
            # Adaptar nome do módulo para o formato de importação
            if dep == "python-dotenv":
                mod_name = "dotenv"
            else:
                mod_name = dep.lower()
            
            # Tentar importar
            importlib.import_module(mod_name)
            print(f"✓ {dep} - Instalado")
        except ImportError:
            print(f"✗ {dep} - Não encontrado")
            faltando.append(dep)
    
    # Oferecer instalação das bibliotecas faltantes
    if faltando:
        print("\nAlgumas dependências estão faltando. Deseja instalá-las agora? (S/N)")
        resposta = input().strip().upper()
        
        if resposta == "S":
            for dep in faltando:
                print(f"\nInstalando {dep}...")
                subprocess.run([sys.executable, "-m", "pip", "install", dep])
        else:
            print("\nVocê precisa instalar estas bibliotecas manualmente:")
            for dep in faltando:
                print(f"- {dep}")
    else:
        print("\nTodas as dependências estão instaladas!")

if __name__ == "__main__":
    verificar_dependencias()
    print("\nPressione ENTER para sair...")
    input()