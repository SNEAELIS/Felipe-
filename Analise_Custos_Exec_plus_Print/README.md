🛰️ Script de Automação de Coleta de Dados no Transferegov
Este script realiza a automação da coleta de informações públicas no portal Transferegov, focando em dados descritivos de execução e acompanhamento/fiscalização de convênios ou processos.

🚀 Funcionalidades
O script executa duas etapas principais:

1. Coleta de Execução do Plano de Trabalho
Acessa a aba "Execução" do Transferegov.

Localiza e baixa anexos dos processos.

Espera automaticamente até que todos os downloads sejam concluídos.

Move os arquivos baixados para uma pasta nomeada com o número do processo.

Compacta os arquivos em um .zip com o mesmo nome da pasta.

2. Coleta de Esclarecimentos (Acompanhamento e Fiscalização)
Acessa a aba "Acompanhamento e Fiscalização".

Utiliza pyautogui para simular comandos de teclado e realizar a impressão em PDF da página com os esclarecimentos.

Salva o arquivo PDF na mesma pasta criada na primeira etapa.

🧾 Requisitos
Navegador Google Chrome instalado e visível (modo não headless).

Python 3.8 ou superior.

Acesso à internet.

Permissão para controlar o teclado e mouse (uso do pyautogui).

📦 Bibliotecas Utilizadas
python
Copiar
Editar
from selenium import webdriver
from selenium.common import WebDriverException, TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
from selenium.webdriver import ActionChains
from colorama import init, Fore, Back, Style
from pathlib import Path
import pandas as pd
import time, os, sys, shutil, json, traceback
import fitz  # PyMuPDF
import hashlib
import pyautogui
import re
import zipfile
🗂️ Entrada de Dados
Um arquivo Excel (.xlsx) contendo os números dos processos a serem pesquisados.

🗃️ Saída
Para cada processo:

Uma pasta nomeada com o número do processo.

Arquivos baixados e o PDF de fiscalização.

Um arquivo .zip com todo o conteúdo.

⚠️ Observações
O script simula ações humanas, portanto, o computador não deve ser utilizado durante a execução.

Evite mover ou editar a pasta de download enquanto o script está rodando.

Caso o Chrome esteja configurado para pedir confirmação de download, isso deve ser desativado nas configurações do navegador.

📌 Sugestões Futuras
Implementar modo headless (sem interface gráfica) com captura automática de PDF.

Adicionar suporte a múltiplas abas ou threads para melhorar desempenho.

Gerar relatórios de sucesso/falha por processo.