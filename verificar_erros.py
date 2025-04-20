import os
import sys
from datetime import datetime

def verificar_erros_instancias():
    """Verifica se há logs de erro em todas as instâncias."""
    
    # Diretório base
    base_dir = os.path.dirname(os.path.abspath(__file__))
    instances_dir = os.path.join(base_dir, "instancias")
    
    if not os.path.exists(instances_dir):
        print("Diretório de instâncias não encontrado.")
        return
    
    # Encontrar todas as instâncias
    instancias = [d for d in os.listdir(instances_dir) 
                 if os.path.isdir(os.path.join(instances_dir, d))]
    
    if not instancias:
        print("Nenhuma instância encontrada.")
        return
    
    print(f"Verificando {len(instancias)} instâncias...")
    
    for instancia in instancias:
        pasta = os.path.join(instances_dir, instancia)
        print(f"\n=== Instância: {instancia} ===")
        
        # Verificar error_log.txt
        error_log = os.path.join(pasta, "error_log.txt")
        if os.path.exists(error_log):
            with open(error_log, "r", encoding="utf-8") as f:
                conteudo = f.read()
            print(f"[!] Erros encontrados:\n{conteudo}")
        else:
            print("✓ Nenhum log de erro encontrado.")

if __name__ == "__main__":
    verificar_erros_instancias()
    print("\nVerificação concluída. Pressione ENTER para sair...")
    input()