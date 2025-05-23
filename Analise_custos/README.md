Analise_Custos.py
O script Analise_Custos.py é uma automação desenvolvida em Python que realiza web scraping no portal Transferegov, especificamente na seção de transferências discricionárias. O objetivo é consultar propostas e baixar documentos estratégicos relacionados ao acompanhamento e análise de custos.

🔁 Fluxo Automatizado
Este script itera sobre um arquivo Excel (.xlsx) contendo números de propostas e, para cada uma delas, realiza a navegação e extração de arquivos conforme o tipo de instrumento.

Caminho Principal
Acessa o portal Transferegov:

Propostas > Consultar Propostas

Seleciona proposta com base no número

Ramo condicional baseado no tipo do instrumento:

Se o instrumento for "Convênio":

Acessa: Projeto Básico/Termo de Referência

Baixa documentos

Acessa: Plano de Trabalho > Anexos

Baixa documentos de Anexos Proposta

Acessa Anexos Execução

Baixa documentos

Retorna à página de consulta

Caso contrário:

Acessa: Plano de Trabalho > Anexos

Baixa documentos de Anexos Proposta

Acessa Anexos Execução

Baixa documentos

Retorna à página de consulta

Organização dos arquivos:

Cria uma pasta nomeada com o número da proposta

Todos os documentos baixados são salvos nessa pasta

🗂️ Exemplo de Organização de Arquivos
python
Copiar
Editar
Analise_Custos/
├── 1234567_2022/
│   ├── projeto_basico.pdf
│   ├── anexos_proposta.zip
│   └── anexos_execucao.zip
├── 7654321_2021/
│   ├── anexos_proposta.pdf
│   └── anexos_execucao.pdf
📚 Requisitos
Python 3.8 ou superior

Navegador Google Chrome

ChromeDriver compatível com a versão do navegador

Conexão com a internet

Arquivo .xlsx contendo os números das propostas

🔧 Instalação
Instale os pacotes com:

bash
Copiar
Editar
pip install -r requirements.txt
Bibliotecas principais:

selenium

pandas

openpyxl

os, shutil, time

▶️ Como Executar
Certifique-se de que o ChromeDriver esteja no PATH do sistema

Garanta que o Excel de entrada tenha uma coluna com os números de proposta

Execute o script com:

bash
Copiar
Editar
python Analise_Custos.py
📝 Estrutura Esperada do Arquivo Excel
proposta
1234567/2022
7654321/2021

⚠️ Observações
O layout do site Transferegov pode sofrer alterações. Mantenha os seletores XPATH atualizados

Evite interagir com o navegador enquanto o script estiver em execução

O nome da pasta gerada segue o número da proposta

📄 Licença
Este projeto está sob a licença MIT.

🤝 Contribuições
Contribuições, melhorias e sugestões são bem-vindas! Utilize pull requests ou abra issues no repositório.

