CGAC_Presta_Contas
CGAC_Presta_Contas é um script de automação em Python que realiza web scraping no portal Transferegov, acessando informações detalhadas sobre execução e prestação de contas de convênios. Ele lê uma planilha Excel contendo os números de processos e executa uma rotina de navegação automatizada para cada um deles, realizando o download e organização de documentos em pastas específicas.

🔁 Fluxo Automatizado
Este script executa o seguinte fluxo para cada número de processo listado em uma planilha:

Lê um arquivo Excel (.xlsx) com os números de processo (um por linha).

Para cada número:

Acessa o módulo: Execução

Vai até: Consultar Instrumentos/Pré-Instrumentos

Seleciona o pré-instrumento correspondente

Abre: Plano de Trabalho

Entra na aba: Anexos

Acessa: Anexos Execução

Baixa os documentos disponíveis

Retorna para Anexos

Acessa: Anexos Prestação de Contas

Baixa os documentos disponíveis

Acessa: Prestação de Contas

Entra em: Prestar Contas

Vai até: Cumprimento do Objeto

Baixa os documentos da tabela de anexos

Cria uma pasta com o número do processo como nome

Move os arquivos baixados para a pasta correspondente

📁 Requisitos
Python 3.8 ou superior

Navegador Google Chrome

ChromeDriver compatível com a versão do Chrome instalado

Internet estável

Arquivo .xlsx com os números de processos

📚 Dependências
Instale as dependências com:

bash
Copiar
Editar
pip install -r requirements.txt
Principais bibliotecas utilizadas:

selenium

pandas

openpyxl

time

os

shutil

▶️ Como Usar
Coloque o arquivo .xlsx com os números dos processos na mesma pasta do script.

Edite o script caso seja necessário ajustar caminhos ou nomes de colunas.

Execute o script com:

bash
Copiar
Editar
python CGAC_Presta_Contas.py
🗂️ Estrutura Esperada da Planilha
O arquivo Excel deve conter os números de processo em uma coluna (por exemplo: "processo" ou "instrumento"), preferencialmente na primeira coluna:

processo
1234567/2021
7654321/2020

⚠️ Observações
A estrutura do site Transferegov pode mudar. Ajustes podem ser necessários caso o layout ou elementos da página sejam alterados.

O usuário deve ter acesso ao sistema com as permissões adequadas.

Evite interagir com o navegador durante a execução do script.

📄 Licença
Distribuído sob a licença MIT. Consulte o arquivo LICENSE para mais informações.

🤝 Contribuições
Contribuições são bem-vindas! Abra uma issue ou envie um pull request com melhorias ou correções.