CGAC_Presta_Contas
CGAC_Presta_Contas é um script de automação desenvolvido em Python com foco em web scraping. Ele automatiza a navegação e o download de documentos no portal Transferegov, acessando dados de instrumentos ou pré-instrumentos relacionados à execução e prestação de contas de convênios.

🚀 Funcionalidades
Este script automatiza o seguinte fluxo dentro do portal Transferegov:

Acessa o módulo: Execução

Navega até: Consultar Instrumentos/Pré-Instrumentos

Seleciona o pré-instrumento desejado

Acessa: Plano de Trabalho

Abre: Anexos

Entra em: Anexos Execução

Baixa todos os documentos disponíveis

Retorna para tela de Anexos

Acessa: Anexos Prestação de Contas

Baixa todos os documentos disponíveis

Acessa: Prestação de Contas

Entra em: Prestar Contas

Abre: Cumprimento do Objeto

Baixa todos os anexos da tabela exibida

📦 Requisitos
Python 3.8 ou superior

Google Chrome instalado

ChromeDriver compatível com sua versão do Chrome

Internet estável

📚 Bibliotecas Utilizadas
selenium

time

os

traceback

(Outras bibliotecas específicas que seu script usar — adicione aqui)

⚙️ Instalação
Clone o repositório:

bash
Copiar
Editar
git clone https://github.com/seu-usuario/CGAC_Presta_Contas.git
cd CGAC_Presta_Contas
Crie um ambiente virtual (opcional, mas recomendado):

bash
Copiar
Editar
python -m venv venv
source venv/bin/activate   # Linux/macOS
venv\Scripts\activate.bat  # Windows
Instale as dependências:

bash
Copiar
Editar
pip install -r requirements.txt
▶️ Como Usar
Abra o script CGAC_Presta_Contas.py

Configure suas credenciais, se necessário.

Execute o script:

bash
Copiar
Editar
python CGAC_Presta_Contas.py
O script irá iniciar um navegador, navegar pelas seções apropriadas e realizar o download automático dos documentos.

🛠️ Estrutura do Código
clicar(): Lida com cliques em elementos da página.

conta_paginas(): Verifica o número de páginas em tabelas.

webdriver_element_wait(): Aguarda elementos da interface.

(Adicione outras funções importantes conforme necessário)

📂 Saída
Todos os documentos baixados serão salvos automaticamente na pasta padrão de downloads do navegador ou em diretórios definidos no script.

⚠️ Avisos
A estrutura do portal Transferegov pode mudar com o tempo, o que pode exigir ajustes no script.

É necessário ter permissão/autorização para acessar e baixar os documentos.

O uso automatizado do site deve seguir os termos de uso do portal.

📄 Licença
Este projeto está licenciado sob a licença MIT. Veja o arquivo LICENSE para mais detalhes.

🤝 Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues, sugerir melhorias ou enviar pull requests.