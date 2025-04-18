# Autocofin - Sistema de Automação COFIN

![Autocofin](https://github.com/user/repo/blob/main/assets/logo.png)

Sistema de automação para o portal COFIN que permite processar números de série de forma automática.

## Funcionalidades

- Processamento automático de números de série a partir de planilha Excel
- Interface gráfica intuitiva com atualização em tempo real
- Suporte a pausa e retomada da automação
- Geração de relatórios de sucesso e erro
- Funcionalidade de tentativa automática em caso de falha

## Requisitos

- Python 3.8 ou superior
- Microsoft Edge instalado
- Acesso ao sistema COFIN

## Instalação

1. Clone este repositório:
   ```bash
   git clone https://github.com/seu-usuario/Novo-Autocofin.git
   cd Novo-Autocofin
   ```

2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

3. Configure o arquivo `.env` com suas credenciais:
   ```
   CPF=seu_cpf
   SENHA=sua_senha
   ```

## Uso

1. Coloque a planilha com os números de série no diretório raiz (`nserie.xlsx`)

2. Execute o programa principal:
   ```bash
   python main.py
   ```

3. Selecione a unidade desejada e clique em **"Iniciar"**

## Estrutura do Projeto

```
main.py           # Interface gráfica e controles principais  
sistemacofin.py   # Lógica de automação com Selenium  
nserie.xlsx       # Arquivo de entrada (exemplo)  
```

## Licença

Este projeto está licenciado sob a licença MIT - veja o arquivo [LICENSE](LICENSE) para mais detalhes.
