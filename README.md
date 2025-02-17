# Busca CEP Automático

## Descrição
Este projeto automatiza a busca de informações de endereço a partir de uma lista de CEPs. Utiliza Selenium para acessar um site de busca de CEP, extrair os dados e salvar os resultados em um arquivo CSV e PDF. 
Também inclui funcionalidades para login automatizado no Gmail e envio de e-mails com os resultados anexados.

## Funcionalidades
- Leitura de uma lista de CEPs de um arquivo CSV
- Busca automática dos dados do CEP em um site de consulta
- Extração e formatação dos dados
- Salvamento dos resultados em um arquivo CSV
- Geração de um relatório em PDF com os resultados
- Login automático no Gmail e envio de e-mail com o arquivo CSV anexado
- Automatiza a abertura do navegador Edge e interação com o Gmail via PyAutoGUI

## Tecnologias Utilizadas
- **Python** (linguagem principal)
- **Selenium** (automatiza a busca de CEPs no navegador)
- **PyAutoGUI** (automatiza a interação com o navegador e o envio de e-mails)
- **ReportLab** (gera relatórios em PDF)
- **Pyperclip** (manipula a área de transferência para anexar arquivos)
- **Webdriver Manager** (gerencia o driver do Chrome automaticamente)
- **CSV** (manipula arquivos de entrada e saída)
- **OS e Time** (operações no sistema e controle de tempo de execução)

## Instalação
1. Clone este repositório:
   git clone git@github.com:AndreCordeiroSantos/buscaCep.git

2. Acesse a pasta do projeto:
   cd busca-cep

3. Instale as dependências:
   pip install -r requirements.txt

## Uso
1. Adicione os CEPs ao arquivo `cep-list.csv`, com um CEP por linha.
2. Execute o script principal:
   python main.py

3. O script irá gerar os arquivos `resultados.csv` e `relatorio_resultados.pdf`.
4. O sistema abrirá o navegador automaticamente, acessará o Gmail e enviará um e-mail com o arquivo CSV anexado.

## Observações
- Certifique-se de ter o Google Chrome instalado e atualizado.
- O WebDriver do Chrome será instalado automaticamente pela biblioteca `webdriver_manager`.
- O login automatizado no Gmail pode exigir permissões adicionais e ajustes na conta.
- Ajuste as coordenadas do `pyautogui` conforme necessário para diferentes resoluções de tela.

## Autor
- **André Rafael** - Criador e desenvolvedor do projeto.

## Licença
Este projeto está sob a licença MIT.

