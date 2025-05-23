Transferencias_especiais_PT.py
O script Transferencias_especiais_PT.py é uma automação desenvolvida em Python que realiza web scraping no portal Transferegov, especificamente na seção de Transferências Especiais. Seu objetivo é coletar dados detalhados de planos de ação com base em números de propostas fornecidos por um arquivo Excel (.xlsx), e retornar essas informações estruturadas no próprio arquivo, organizando os dados em novas linhas para facilitar a leitura e análise.

🔁 Fluxo Automatizado
Este script itera sobre um arquivo Excel contendo números de propostas e, para cada uma, executa a seguinte sequência de navegação e coleta de dados no portal Transferegov:

Caminho Navegado no Site
Acesso inicial:

Menu > Plano de Ação

Filtro da proposta:

Insere o código do plano de ação

Aplica o filtro

Detalha o plano de ação

Coleta de Dados:

🧾 Dados Básicos:

Beneficiário

UF

Banco

Agência

Conta

Situação da Conta

Emenda Parlamentar

Valor de Investimento

Finalidades

Programações Orçamentárias

💰 Dados Orçamentários:

Empenho

Valor

Ordem de Pagamento

📋 Plano de Trabalho:

Indicação Orçamento Beneficiário

Classificação Orçamentária da Despesa

Declaração de Recurso

Prazo de Execução

Período de Execução

Executor

Objeto

Meta

Unidade de Medida

Quantidade

Responsável

Data

Situação

Dados dos Conselhos Locais ou Instâncias de Controle Social

Organização do Resultado:

Os dados coletados são inseridos no próprio arquivo Excel

Novas linhas são criadas dinamicamente para estruturar os dados de forma legível e organizada

📌 Regras de Coleta
Apenas dados oriundos do "Sistema" são considerados

A Situação "Concluído" precisa aparecer logo após uma entrada do tipo "Sistema"

Caso as condições não sejam satisfeitas, nenhuma informação é registrada

📚 Requisitos
Python 3.8 ou superior

Navegador Google Chrome

ChromeDriver compatível com a versão do navegador

Conexão com a internet

Arquivo .xlsx contendo os códigos de plano de ação

🛠️ Instalação
Instale as dependências com:

bash
Copiar
Editar
pip install -r requirements.txt
Bibliotecas utilizadas:

selenium

pandas

openpyxl

time, os, re, shutil

▶️ Como Executar
Garanta que o ChromeDriver esteja no PATH do sistema

Insira no Excel uma coluna com os códigos dos planos de ação

Execute o script com:

bash
Copiar
Editar
python Transferencias_especiais_PT.py
📝 Estrutura Esperada do Arquivo Excel
proposta
1100001/2023
1100002/2023

Após a execução, o próprio arquivo será atualizado com colunas e linhas adicionais contendo as informações extraídas.

⚠️ Observações
O layout do portal Transferegov pode ser alterado; mantenha os seletores XPATH atualizados

O script pode levar alguns segundos por proposta, dependendo do desempenho da conexão e do site

Não interaja com o navegador durante a execução automatizada

📄 Licença
Este projeto está licenciado sob a licença MIT.

🤝 Contribuições
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests com melhorias, correções ou sugestões.

