
# Selecao_PAC_DIE

Selecao_PAC_DIE é um script de automação desenvolvido em Python que executa web scraping no portal Transferegov, navegando pelo módulo "Seleção PAC" e baixando os documentos relacionados às propostas. Além disso, o script também acessa links externos do Google Maps ou Google Earth associados a cada proposta e realiza capturas de tela, organizando todos os arquivos em diretórios nominais.

## 🔁 Fluxo Automatizado

Este script realiza um loop para cada número de proposta listado em um arquivo Excel (.xlsx) e executa as seguintes etapas:

1. Acessa o portal Transferegov:
   - Propostas > Seleção PAC
   - Localiza a proposta pelo número
   - Detalha a linha da tabela correspondente
   - Baixa os arquivos anexos da Seleção PAC
   - Volta para a tabela para continuar o loop

2. Captura de telas geográficas:
   - Para cada proposta, acessa uma ou mais URLs do Google Maps / Google Earth
   - Realiza captura de tela da visualização da área
   - Salva as imagens na mesma pasta da proposta

3. Organização:
   - Cria uma pasta local nomeada com o número da proposta
   - Move todos os documentos baixados e imagens para esta pasta

## 🗂️ Estrutura Esperada dos Arquivos

Exemplo de organização gerada:

```
Selecao_PAC_DIE/
├── 1234567_2022/
│   ├── anexo_proposta.pdf
│   ├── screenshot_maps.png
│   └── screenshot_earth.png
├── 7654321_2021/
│   ├── anexo_documento.pdf
│   └── screenshot_google.png
```

## 📚 Requisitos

- Python 3.8 ou superior  
- Navegador Google Chrome  
- ChromeDriver compatível com a versão instalada  
- Acesso à internet  
- Arquivo .xlsx contendo os números de proposta e links de mapas

## 🔧 Instalação

Instale as dependências necessárias com:

```bash
pip install -r requirements.txt
```

Principais bibliotecas utilizadas:

- selenium  
- pandas  
- openpyxl  
- time, os, shutil  
- PIL ou pyautogui (para screenshots, conforme implementado)

## ▶️ Como Usar

1. Certifique-se de que o ChromeDriver está instalado e configurado no PATH.  
2. Prepare o arquivo Excel com as colunas necessárias:
   - Número da proposta (ex: 1234567/2022)
   - Um ou mais links do Google Maps / Google Earth  
3. Execute o script:

```bash
python Selecao_PAC_DIE.py
```

## 📝 Estrutura Esperada do Arquivo Excel

| proposta     | link_maps                               | link_earth                              |
|--------------|------------------------------------------|------------------------------------------|
| 1234567/2022 | https://maps.google.com/...             | https://earth.google.com/...            |
| 7654321/2021 | https://maps.google.com/...             | https://earth.google.com/...            |

Você pode adaptar a planilha de entrada de acordo com o número de links disponíveis.

## ⚠️ Observações

- O layout do site Transferegov pode mudar. Verifique e ajuste os seletores (XPATHs) caso ocorra falha na automação.  
- Evite interagir com o navegador durante a execução do script.  
- Para capturas de tela funcionarem corretamente, o navegador deve estar visível e o foco deve estar na aba certa.  
- Para múltiplos monitores ou resoluções específicas, pode ser necessário ajustar os parâmetros de captura.

## 📄 Licença

Este projeto está licenciado sob a licença MIT.

## 🤝 Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues, reportar bugs ou enviar pull requests com melhorias.
