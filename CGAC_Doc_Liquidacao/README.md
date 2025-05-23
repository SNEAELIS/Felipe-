CGAC_Doc_Liquidacao
CGAC_Doc_Liquidacao é um script de automação em Python que realiza web scraping no portal Transferegov. Ele automatiza a coleta de documentos relacionados aos "Documentos de Liquidação" vinculados a propostas registradas, baixando os arquivos e organizando-os localmente em pastas nomeadas com base no número da proposta.

🔁 Fluxo Automatizado
Este script executa um loop para cada proposta listada em um arquivo Excel (.xlsx), realizando o seguinte caminho no portal Transferegov:

Acessa a seção: Execução

Vai para: Consultar Instrumentos/Pré-Instrumentos

Seleciona o pré-instrumento correspondente à proposta

Acessa: Execução Convenente

Abre: Documento de Liquidação

Detalha cada linha da tabela de documentos

Baixa os arquivos disponíveis nos anexos

Retorna à tabela para continuar a iteração

Cria uma pasta local com o número da proposta como nome

Move os arquivos baixados para a pasta correspondente

📁 Organização dos Arquivos
Ao final da execução, a estrutura de pastas será semelhante a:

objectivec
Copiar
Editar
CGAC_Doc_Liquidacao/
├── 1234567_2022/
│   ├── documento1.pdf
│   └── documento2.pdf
├── 7654321_2021/
│   ├── documentoA.pdf
│   └── documentoB.pdf
📚 Requisitos
Python 3.8 ou superior

Navegador Google Chrome

ChromeDriver compatível com sua versão do Chrome

Conexão de internet estável

Arquivo .xlsx com os números de proposta

🛠️ Instalação e Dependências
Instale as bibliotecas necessárias com:

bash
Copiar
Editar
pip install -r requirements.txt
Principais bibliotecas utilizadas:

selenium

pandas

openpyxl

time, os, shutil

▶️ Como Usar
Certifique-se de que o ChromeDriver está instalado e no PATH.

Coloque o arquivo .xlsx com as propostas na mesma pasta do script.

Execute o script com:

bash
Copiar
Editar
python CGAC_Doc_Liquidacao.py
🗂️ Estrutura Esperada do Excel
O arquivo .xlsx deve conter os números das propostas em uma coluna, preferencialmente a primeira:

proposta
1234567/2022
7654321/2021

⚠️ Observações
A estrutura do portal Transferegov pode mudar. Em caso de erros, verifique os seletores utilizados no código (XPATHs).

O usuário pode precisar estar autenticado no sistema para acessar os dados.

Não utilize o navegador durante a execução do script para evitar conflitos com a automação.

📄 Licença
Distribuído sob a licença MIT. Consulte o arquivo LICENSE para mais detalhes.

🤝 Contribuições
Contribuições são bem-vindas! Abra uma issue ou envie um pull request para sugerir melhorias ou corrigir problemas.