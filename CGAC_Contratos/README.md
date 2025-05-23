CGAC_Contratos
CGAC_Contratos é um script de automação em Python que realiza scraping de dados e download de arquivos de contratos no portal Transferegov. Ele automatiza a navegação por múltiplas propostas definidas em uma planilha Excel e organiza os arquivos baixados em pastas específicas para cada proposta.

🔁 Fluxo Automatizado
O script executa o seguinte ciclo de forma automatizada:

Lê uma planilha Excel contendo os números das propostas.

Para cada número de proposta:

Acessa a seção: Execução

Navega para: Consultar Instrumentos/Pré-Instrumentos

Seleciona o pré-instrumento correspondente

Entra em: Execução Convenente

Acessa: Contratos/Subconvênio

Detalha cada linha da tabela de contratos

Baixa todos os documentos disponíveis na seção: Arquivos do Contrato

Retorna à tabela para processar a próxima linha

Cria uma pasta local com o número da proposta como nome

Move os arquivos baixados para a respectiva pasta

🧩 Funcionalidades
Navegação autônoma no portal Transferegov

Download completo dos arquivos de contratos e subconvênios

Organização automática dos arquivos em pastas nomeadas com base na proposta

Iteração por múltiplos registros usando um arquivo Excel (.xlsx)

📁 Requisitos
Python 3.8 ou superior

Navegador Google Chrome

ChromeDriver compatível com sua versão do Chrome

Internet estável

Arquivo .xlsx com os números de proposta (uma por linha)

📚 Dependências
Instale as bibliotecas necessárias com:

bash
Copiar
Editar
pip install -r requirements.txt
Dependências comuns incluem:

selenium

pandas

openpyxl

time

os

shutil

▶️ Como Usar
Edite o script CGAC_Contratos.py se necessário (ex: caminho do arquivo Excel).

Coloque o arquivo .xlsx com as propostas na mesma pasta do script.

Execute o script com:

bash
Copiar
Editar
python CGAC_Contratos.py
Após a execução, as pastas com os arquivos baixados serão criadas automaticamente na pasta do projeto.

🗂️ Estrutura Esperada do Excel
O arquivo Excel deve conter os números das propostas em uma coluna, preferencialmente na primeira (coluna A), com um cabeçalho ou sem.

Exemplo:

proposta
123456/2021
654321/2020

⚠️ Observações
A estrutura do portal Transferegov pode mudar; ajustes no XPATH ou lógica podem ser necessários futuramente.

O script exige que o usuário esteja autenticado ou autorizado no sistema, caso necessário.

Evite usar o navegador durante a execução para não interferir na automação.

📄 Licença
Distribuído sob a licença MIT. Consulte o arquivo LICENSE para mais detalhes.

🤝 Contribuições
Contribuições são bem-vindas! Se encontrar problemas ou quiser melhorar algo, abra uma issue ou envie um pull request.