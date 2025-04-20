import os
import shutil
import sys
from datetime import datetime, timedelta

def limpar_instancias(dias=7):
    """
    Remove diretórios de instâncias antigos que não foram modificados há mais de X dias.
    
    Args:
        dias: Número de dias para considerar uma instância como antiga
    """
    print(f"Procurando instâncias não utilizadas há mais de {dias} dias...")
    
    instances_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "instancias")
    
    if not os.path.exists(instances_dir):
        print("Diretório de instâncias não encontrado.")
        return
    
    dirs = [d for d in os.listdir(instances_dir) 
           if os.path.isdir(os.path.join(instances_dir, d))]
    
    if not dirs:
        print("Nenhuma instância encontrada.")
        return
    
    limite = datetime.now() - timedelta(days=dias)
    count = 0
    
    for dir_name in dirs:
        dir_path = os.path.join(instances_dir, dir_name)
        last_modified = datetime.fromtimestamp(os.path.getmtime(dir_path))
        
        if last_modified < limite:
            print(f"Removendo: {dir_name} (Última modificação: {last_modified.strftime('%d/%m/%Y %H:%M:%S')})")
            try:
                shutil.rmtree(dir_path)
                count += 1
            except Exception as e:
                print(f"  Erro ao remover: {str(e)}")
    
    print(f"\nOperação concluída. {count} instâncias antigas removidas.")

if __name__ == "__main__":
    # Verificar se há argumento para número de dias
    dias = 7  # padrão
    
    if len(sys.argv) > 1:
        try:
            dias = int(sys.argv[1])
        except:
            print("Argumento inválido. Usando valor padrão de 7 dias.")
    
    limpar_instancias(dias)
    
    # Aguardar enter para sair
    input("\nPressione ENTER para sair...")