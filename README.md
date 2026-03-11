# 📊 Central de Controle Full-Stack Feito Por Vicenzo!

> Painel interativo para monitoramento de dados e gestão de alertas, desenvolvido de forma independente.

## 📝 Sobre o Projeto
Esta é uma **Central de Controle** completa que integra uma interface dinâmica com um back-end robusto. O foco principal foi criar um sistema capaz de ler dados de um banco de dados SQL, processá-los via Python e exibi-los em tempo real para o usuário, permitindo o acompanhamento de métricas e alertas automáticos.

## 🛠️ Tecnologias Utilizadas

* **Front-end:** HTML5, CSS3 e JavaScript (Vanilla) para manipulação do DOM e lógica de filtros.
* **Back-end:** Python com o framework **Flask**.
* **Banco de Dados:** SQL Server acessado via biblioteca **pyodbc**.
* **Automação:** Scripts de inicialização em lote (`.bat`) para facilitar o deploy local.

## ✨ Funcionalidades
* **Visualização de Dados:** Exibição dinâmica de informações em tabelas e gráficos.
* **Sistema de Alertas:** Contador automático (`count-alerta`) que identifica pendências ou notificações críticas.
* **Filtros Inteligentes:** Lógica em JS para filtrar resultados sem a necessidade de recarregar a página.
* **Integração de Dados:** Conexão direta com banco de dados SQL para garantir dados reais e atualizados.

## 🚀 Como Executar
1. Certifique-se de ter o **Python 3.x** e o driver do **SQL Server** instalados.
2. Instale as dependências necessárias:
   ```bash
   pip install flask pyodb

3. Configure a sua string de conexão no arquivo app.py.
4. Execute o sistema através do arquivo Iniciar_Sistema.bat ou via terminal:
    ```bash
    python app.py

Desenvolvido por Vicenzo Henrique Da Silva Conceição!


    

   
