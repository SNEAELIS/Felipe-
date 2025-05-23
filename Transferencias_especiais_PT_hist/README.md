Transferencias_especiais_PT_Sistema.py
O script Transferencias_especiais_PT_Sistema.py é uma automação desenvolvida em Python que realiza web scraping no portal Transferegov, especificamente na seção de Transferências Especiais. Seu objetivo é coletar dados selecionados do Plano de Ação com base em números de propostas fornecidos por um arquivo Excel (.xlsx), focando apenas nas ações que foram registradas pelo Sistema e que estejam concluídas.

🔁 Fluxo Automatizado
Este script itera sobre um arquivo Excel contendo os números das propostas e, para cada uma, realiza a navegação e coleta das seguintes informações:

Navegação no Site
Acesso inicial:

Menu > Plano de Ação

Filtro da proposta:

Inserção do código do plano de ação

Aplicação do filtro

Acesso à tela de detalhamento do plano de ação

Coleta de dados condicionais da seção Plano de Trabalho:

Busca pelas linhas onde o campo Responsável é "Sistema"

Em seguida, coleta a próxima linha onde a Situação é "Concluído"

Extrai os campos:

Responsável

Data

Situação

Organização do Resultado:

As informações extraídas são registradas no próprio arquivo Excel de origem

Novas colunas e linhas são criadas de forma estruturada e legível

📌 Regras de Coleta
Apenas dados oriundos do "Sistema" são considerados

A Situação "Concluído" precisa aparecer logo após uma entrada do tipo "Sistema"

Caso as condições não sejam satisfeitas, nenhuma informação é registrada

📚 Requisitos
Python 3.8 ou superior

Navegador Google Chrome

ChromeDriver compatível com sua versão do navegador

Conexão com a internet

Arquivo Excel (.xlsx) contendo os códigos dos planos de ação

🛠️ Instalação
Instale os pacotes necessários com:

bash
Copiar
Editar
pip install -r requirements.txt
Bibliotecas utilizadas:

selenium

pandas

openpyxl

os, time, re

▶️ Como Executar
Verifique se o ChromeDriver está disponível no PATH do sistema

Prepare o arquivo .xlsx com uma coluna chamada proposta

Execute o script com o comando:

bash
Copiar
Editar
python Transferencias_especiais_PT_Sistema.py
📝 Estrutura Esperada do Arquivo Excel
proposta
1100001/2023
1100002/2023

O script irá adicionar novas colunas, como por exemplo:

proposta	responsavel_sistema	data_envio	situacao_envio
1100001/2023	Sistema	12/04/2023	Concluído

⚠️ Observações
O script depende da estrutura atual do site Transferegov; seletores XPATH podem precisar de ajustes

Evite interferência no navegador durante a execução

O processo pode levar alguns segundos por proposta, dependendo da estabilidade do portal

📄 Licença
Este projeto está licenciado sob a licença MIT.

🤝 Contribuições
Sugestões, melhorias ou correções são bem-vindas! Envie um pull request ou abra uma issue para contribuir com o projeto.