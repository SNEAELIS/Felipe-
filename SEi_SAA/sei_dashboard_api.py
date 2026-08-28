"""
SEI Dashboard — API Version
============================
Substitui o scraping via Playwright por chamadas aos Web Services do SEI (SOAP).

Requer:
    pip install zeep pandas openpyxl

Uso:
    1. Copie sei_credentials.json.example para sei_credentials.json
    2. Preencha os dados (URL, chave, IDs) no sei_credentials.json
    3. Execute: python sei_dashboard_api.py

Fluxo:
    1. Carrega credenciais do JSON (apenas em memória)
    2. Conecta ao Web Service SOAP do SEI
    3. Consulta bloco de assinatura → extrai dados dos processos
    4. Para cada processo, consulta detalhes e andamentos
    5. Salva em Excel (mesmo formato do script original)
"""

import json
import os
import sys
import time
import traceback
import shutil

from pathlib import Path
from datetime import datetime
from dataclasses import dataclass
from typing import Optional, Any

import pandas as pd

from zeep import Client
from zeep.settings import Settings
from zeep.wsdl import wsdl
from zeep.exceptions import Fault, TransportError, XMLSyntaxError


os.system('cls' if os.name == 'nt' else 'clear')


# =============================================================================
# Configurações
# =============================================================================

CREDENTIALS_FILE = Path(__file__).parent / "sei_credentials.json"

ARQUIVO_FONTE = r"C:\Users\felipe.rsouza\Automação SNEAELIS\Dashboard sei DB\DB_sei_se - Copia.xlsx"
ARQUIVO_DESTINO = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\Controle_SEI\DB_sei_se.xlsx"
ARQUIVO_DESTINO_2 = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Mateus - SEI\DB_sei_se.xlsx"

INTERVALO_LOOP = 300  # 5 minutos entre execuções


# =============================================================================
# Utilitários
# =============================================================================

def carregar_credenciais(caminho: Path) -> dict:
    """
    Carrega as credenciais do arquivo JSON.
    Os dados ficam em memória apenas durante a execução e são descartados
    ao final do script (variável local sobrescrita ao sair do escopo).
    """
    if not caminho.exists():
        print(f"❌ Arquivo de credenciais não encontrado: {caminho}")
        print(f"   Copie o modelo: cp {caminho}.example {caminho}")
        print(f"   E preencha com seus dados.")
        sys.exit(1)

    with open(caminho, 'r', encoding='utf-8') as f:
        credenciais = json.load(f)

    # Remove chaves de comentário (iniciadas com '_')
    credenciais = {k: v for k, v in credenciais.items() if not k.startswith('_')}

    # Valida campos obrigatórios
    campos_obrigatorios = [
        'wsdl_url', 'sigla_sistema', 'identificacao_servico',
        'id_unidade', 'chave_acesso', 'blocos_assinatura'
    ]
    campos_faltando = [c for c in campos_obrigatorios
                       if c not in credenciais or not credenciais[c]]

    if campos_faltando:
        print(f"❌ Campos obrigatórios faltando: {campos_faltando}")
        sys.exit(1)

    # Confirma que blocos_assinatura é uma lista
    if not isinstance(credenciais['blocos_assinatura'], list):
        print("❌ 'blocos_assinatura' deve ser uma lista (ex: [\"432067\", \"432068\"])")
        sys.exit(1)

    if not credenciais['blocos_assinatura']:
        print("❌ 'blocos_assinatura' não pode ser uma lista vazia")
        sys.exit(1)

    # Confirma que a chave não é o placeholder
    if credenciais['chave_acesso'] == 'SUA_CHAVE_DE_ACESSO_AQUI':
        print("❌ Chave de acesso não foi configurada!")
        print("   Edite sei_credentials.json e substitua o placeholder pela chave real.")
        sys.exit(1)

    print("✅ Credenciais carregadas (apenas em memória)")
    return credenciais


def formato_padrao(num_sei: str) -> str:
    """
    Valida/formata número de processo SEI.
    Aceita formatos como:
      - 12.1.000000077-4
      - 12345.678901/2024-00
    Retorna string vazia se inválido.
    """
    import re
    num_sei = str(num_sei).strip()

    # Já está no formato com pontos e barras
    padrao_sei = re.compile(r"^\d{1,2}\.\d{1}\.\d{6,9}-\d{1,2}$")
    padrao_mds = re.compile(r"^\d{5}\.\d{6}\/\d{4}-\d{2}$")

    if padrao_sei.match(num_sei) or padrao_mds.match(num_sei):
        return num_sei

    # Tenta limpar e converter
    num_sei_limpo = re.sub(r'[./-]', '', num_sei)
    if len(num_sei_limpo) < 10:
        return ''

    # SEI format: 12.1.000000077-4
    if len(num_sei_limpo) <= 13:
        part1 = num_sei_limpo[:2]
        part2 = num_sei_limpo[2:3]
        part3 = num_sei_limpo[3:-2]
        part4 = num_sei_limpo[-2:]
        return f"{part1}.{part2}.{part3}-{part4}"
    # MDS format: 12345.678901/2024-00
    elif len(num_sei_limpo) >= 17:
        part1 = num_sei_limpo[:5]
        part2 = num_sei_limpo[5:11]
        part3 = num_sei_limpo[11:15]
        part4 = num_sei_limpo[15:17]
        return f"{part1}.{part2}/{part3}-{part4}"

    return ''


# =============================================================================
# WorkbookState (mesmo dataclass do script original)
# =============================================================================

@dataclass
class WorkbookState:
    """Estado dos dados a serem salvos no Excel."""
    def __init__(self):
        self.bloco: Optional[pd.DataFrame] = None
        self.processos: Optional[pd.DataFrame] = None
        self.andamento: Optional[pd.DataFrame] = None
        self.controle: Optional[pd.DataFrame] = None


# =============================================================================
# Cliente SOAP — SEI Web Services
# =============================================================================

class SEIAPIClient:
    """
    Cliente para os Web Services do SEI (SOAP).
    Documentação: SEI Web Services v5.0.2

    Endpoint: http://[servidor]/sei/controlador_ws.php?servico=sei
    Autenticação: SiglaSistema + IdentificacaoServico + IdUnidade + Chave de Acesso
    """

    def __init__(self, credenciais: dict):
        """
        Inicializa o cliente SOAP com as credenciais fornecidas.

        As credenciais são mantidas em memória apenas durante a vida
        deste objeto. Ao destruir a instância, os dados são perdidos.
        """
        self.sigla_sistema = credenciais['sigla_sistema']
        self.identificacao_servico = credenciais['identificacao_servico']
        self.id_unidade = credenciais['id_unidade']
        self.chave_acesso = credenciais['chave_acesso']  # ← só em memória!
        self.blocos_assinatura = credenciais.get('blocos_assinatura', [])
        self.blocos_internos = credenciais.get('blocos_internos', [])

        # Configura cliente SOAP
        # Tenta usar arquivo WSDL local primeiro (evita timeout de schemas externos)
        wsdl_path = Path(__file__).parent / "sei.wsdl"
        if wsdl_path.exists():
            wsdl_url = str(wsdl_path)
            print(f"   📄 Usando WSDL local: {wsdl_path}")
        else:
            wsdl_url = credenciais['wsdl_url']
            print(f"   🌐 Usando WSDL remoto: {wsdl_url}")

        settings = Settings(strict=False, xml_huge_tree=True)

        try:
            self.client = Client(wsdl_url, settings=settings)
            print(f"✅ Conectado ao Web Service")
            print(f"   Sistema: {self.sigla_sistema}")
            print(f"   Unidade: {self.id_unidade}")
        except Exception as e:
            print(f"❌ Erro ao conectar ao Web Service:")
            print(f"   {type(e).__name__}: {e}")
            raise

    def _chamar(self, servico: str, **kwargs) -> Any:
        """
        Chama um serviço SOAP do SEI, injetando automaticamente os
        parâmetros de autenticação (SiglaSistema, IdentificacaoServico,
        IdUnidade).
        """
        # A autenticação do SEI pode ser por Chave de Acesso ou Endereço.
        # Se for por Chave de Acesso, o parâmetro IdentificacaoServico
        # deve receber a chave de acesso (e não o nome do serviço).
        params = {
            'SiglaSistema': self.sigla_sistema,
            'IdentificacaoServico': self.chave_acesso,
            'IdUnidade': self.id_unidade,
        }
        params.update(kwargs)

        metodo = getattr(self.client.service, servico, None)
        if metodo is None:
            raise ValueError(f"Serviço '{servico}' não encontrado no WSDL.")

        return metodo(**params)

    # ─── Serviços de Consulta ───────────────────────────────────────────

    def consultar_bloco(self, id_bloco: str,
                        retornar_protocolos: bool = True) -> Optional[Any]:
        """
        Consulta um bloco (assinatura, reunião ou interno).

        Args:
            id_bloco: Número do bloco (obrigatório)
            retornar_protocolos: Se deve incluir os protocolos do bloco

        Retorna:
            Objeto RetornoConsultaBloco ou None
        """
        if not id_bloco:
            print("❌ id_bloco é obrigatório")
            return None

        try:
            return self._chamar(
                'consultarBloco',
                IdBloco=id_bloco,
                SinRetornarProtocolos='S' if retornar_protocolos else 'N'
            )
        except Fault as e:
            print(f"❌ Erro SOAP ao consultar bloco {id_bloco}: {e.message}")
            return None
        except Exception as e:
            print(f"❌ Erro ao consultar bloco {id_bloco}: {type(e).__name__}: {e}")
            return None

    def consultar_procedimento(self, protocolo: str,
                               ultimo_andamento: bool = True,
                               andamento_geracao: bool = True,
                               andamento_conclusao: bool = True,
                               assuntos: bool = False,
                               interessados: bool = False,
                               observacoes: bool = False,
                               unidades_aberto: bool = False,
                               relacionados: bool = False,
                               anexados: bool = False) -> Optional[Any]:
        """
        Consulta detalhes de um processo/procedimento.

        Args:
            protocolo: Número do processo (ex: 12.1.000000077-4)
            ultimo_andamento: Incluir o último andamento
            andamento_geracao: Incluir andamento de geração
            andamento_conclusao: Incluir andamento de conclusão
            assuntos: Incluir assuntos
            interessados: Incluir interessados
            observacoes: Incluir observações
            unidades_aberto: Incluir unidades onde está aberto
            relacionados: Incluir processos relacionados
            anexados: Incluir processos anexados

        Retorna:
            Objeto RetornoConsultaProcedimento ou None
        """
        try:
            return self._chamar(
                'consultarProcedimento',
                ProtocoloProcedimento=protocolo,
                SinRetornarAssuntos='S' if assuntos else 'N',
                SinRetornarInteressados='S' if interessados else 'N',
                SinRetornarObservacoes='S' if observacoes else 'N',
                SinRetornarAndamentoGeracao='S' if andamento_geracao else 'N',
                SinRetornarAndamentoConclusao='S' if andamento_conclusao else 'N',
                SinRetornarUltimoAndamento='S' if ultimo_andamento else 'N',
                SinRetornarUnidadesProcedimentoAberto='S' if unidades_aberto else 'N',
                SinRetornarProcedimentosRelacionados='S' if relacionados else 'N',
                SinRetornarProcedimentosAnexados='S' if anexados else 'N'
            )
        except Fault as e:
            print(f"❌ Erro SOAP ao consultar {protocolo}: {e.message}")
            return None
        except Exception as e:
            print(f"❌ Erro ao consultar {protocolo}: {type(e).__name__}: {e}")
            return None

    def listar_andamentos(self, protocolo: str,
                          tarefas: list) -> Optional[list]:
        """
        Lista os andamentos (histórico de trâmite) de um processo.

        ATENÇÃO: É obrigatório informar IDs de tarefas como filtro.
        O valor 65 = "atualização de andamento" genérica.
        Consulte TarefaRN.php do SEI para outros IDs.

        Args:
            protocolo: Número do processo
            tarefas: Lista de IDs de tarefas (obrigatório, ex: [65])

        Retorna:
            Lista de objetos Andamento ou None
        """
        if not tarefas:
            print("⚠️ listar_andamentos requer pelo menos um ID de tarefa")
            return None

        try:
            return self._chamar(
                'listarAndamentos',
                ProtocoloProcedimento=protocolo,
                SinRetornarAtributos='N',
                Tarefas=tarefas
            )
        except Fault as e:
            print(f"⚠️ listar_andamentos para {protocolo}: {e.message[:100]}")
            return None
        except Exception as e:
            print(f"⚠️ listar_andamentos para {protocolo}: {type(e).__name__}")
            return None

    def consultar_documento(self, protocolo_documento: str,
                            assinaturas: bool = True,
                            andamento_geracao: bool = False,
                            publicacao: bool = False,
                            campos: bool = False,
                            blocos: bool = False) -> Optional[Any]:
        """
        Consulta um documento.

        Args:
            protocolo_documento: Número do documento (ex: 0003934)
            assinaturas: Incluir assinaturas
            andamento_geracao: Incluir andamento de geração
            publicacao: Incluir dados de publicação
            campos: Incluir campos do formulário
            blocos: Incluir blocos na unidade

        Retorna:
            Objeto RetornoConsultaDocumento ou None
        """
        try:
            return self._chamar(
                'consultarDocumento',
                ProtocoloDocumento=protocolo_documento,
                SinRetornarAndamentoGeracao='S' if andamento_geracao else 'N',
                SinRetornarAssinaturas='S' if assinaturas else 'N',
                SinRetornarPublicacao='S' if publicacao else 'N',
                SinRetornarCampos='S' if campos else 'N',
                SinRetornarBlocos='S' if blocos else 'N'
            )
        except Fault as e:
            print(f"⚠️ consultarDocumento {protocolo_documento}: {e.message[:100]}")
            return None
        except Exception as e:
            print(f"⚠️ consultarDocumento {protocolo_documento}: {type(e).__name__}")
            return None

    # ─── Serviços de Listagem ───────────────────────────────────────────

    def listar_unidades(self) -> Optional[list]:
        """Lista as unidades disponíveis para o serviço."""
        try:
            return self._chamar('listarUnidades')
        except Exception as e:
            print(f"⚠️ listarUnidades: {type(e).__name__}: {e}")
            return None

    def listar_tipos_procedimento(self) -> Optional[list]:
        """Lista os tipos de processo disponíveis."""
        try:
            return self._chamar('listarTiposProcedimento')
        except Exception as e:
            print(f"⚠️ listarTiposProcedimento: {type(e).__name__}: {e}")
            return None

    def listar_series(self) -> Optional[list]:
        """Lista os tipos de documento disponíveis."""
        try:
            return self._chamar('listarSeries')
        except Exception as e:
            print(f"⚠️ listarSeries: {type(e).__name__}: {e}")
            return None


# =============================================================================
# Processamento dos dados retornados pela API
# =============================================================================

def extrair_dados_bloco(bloco_obj) -> pd.DataFrame:
    """
    Converte o objeto RetornoConsultaBloco em um DataFrame
    para a planilha 'Bloco'.

    Colunas: Número, Sinalizações, Atribuição, Estado, Geradora,
             Disponibilização, Grupo, Descrição, Ações
    """
    if not bloco_obj:
        return pd.DataFrame(columns=[
            'Número', 'Sinalizações', 'Atribuição', 'Estado',
            'Geradora', 'Disponibilização', 'Grupo', 'Descrição', 'Ações'
        ])

    # Extrai metadados do bloco
    descricao = getattr(bloco_obj, 'Descrição', '') or ''
    tipo = getattr(bloco_obj, 'Tipo', '') or ''
    estado = getattr(bloco_obj, 'Estado', '') or ''
    prioridade = getattr(bloco_obj, 'SinPrioridade', '') or ''
    revisao = getattr(bloco_obj, 'SinRevisao', '') or ''

    # Usuário de atribuição
    usuario_attr = getattr(bloco_obj, 'UsuarioAtribuicao', None)
    atribuicao = ''
    if usuario_attr:
        atribuicao = f"{getattr(usuario_attr, 'Nome', '') or ''} ({getattr(usuario_attr, 'Sigla', '') or ''})"

    # Unidade geradora
    unidade = getattr(bloco_obj, 'Unidade', None)
    geradora = ''
    if unidade:
        geradora = f"{getattr(unidade, 'Descricao', '') or ''} ({getattr(unidade, 'Sigla', '') or ''})"

    # Unidades de disponibilização
    unidades_disp = getattr(bloco_obj, 'UnidadesDisponibilizacao', None) or []
    disp_str = '; '.join([
        f"{getattr(u, 'Descricao', '') or ''} ({getattr(u, 'Sigla', '') or ''})"
        for u in (unidades_disp if isinstance(unidades_disp, list) else [])
    ])

    linha = {
        'Número': getattr(bloco_obj, 'IdBloco', '') or '',
        'Sinalizações': f"Prioridade: {prioridade}, Revisão: {revisao}",
        'Atribuição': atribuicao,
        'Estado': estado,
        'Geradora': geradora,
        'Disponibilização': disp_str,
        'Grupo': tipo,
        'Descrição': descricao,
        'Ações': f"Tipo: {tipo} | Estado: {estado}"
    }

    return pd.DataFrame([linha])


def extrair_protocolos_bloco(bloco_obj) -> pd.DataFrame:
    """
    Converte os protocolos de um bloco em um DataFrame
    para a planilha 'Processos'.

    Colunas: Número, Seq., Processo, Documento, Tipo, Assinaturas, Anotações, Ações
    """
    colunas = ['Número', 'Seq.', 'Processo', 'Documento', 'Tipo', 'Assinaturas', 'Anotações', 'Ações']

    if not bloco_obj:
        return pd.DataFrame(columns=colunas)

    protocolos = getattr(bloco_obj, 'Protocolos', None) or []
    if not protocolos or not isinstance(protocolos, list):
        return pd.DataFrame(columns=colunas)

    resultados = []
    for idx, proto in enumerate(protocolos, 1):
        protocolo_formatado = getattr(proto, 'ProtocoloFormatado', '') or ''
        identificacao = getattr(proto, 'Identificacao', '') or ''

        # Assinaturas
        assinaturas = getattr(proto, 'Assinaturas', None) or []
        ass_str = ''
        if assinaturas and isinstance(assinaturas, list):
            ass_str = '; '.join([
                f"{getattr(a, 'Nome', '') or ''} ({getattr(a, 'DataHora', '') or ''})"
                for a in assinaturas
            ])

        # Determina se é processo ou documento
        # A identificação normalmente contém o tipo
        if '/' in protocolo_formatado or '-' in protocolo_formatado:
            # Parece número de processo
            tipo = 'Processo'
            processo = protocolo_formatado
            documento = ''
        else:
            # Parece número de documento
            tipo = 'Documento'
            processo = ''
            documento = protocolo_formatado

        resultados.append({
            'Número': idx,
            'Seq.': idx,
            'Processo': processo,
            'Documento': documento,
            'Tipo': identificacao or tipo,
            'Assinaturas': ass_str,
            'Anotações': '',
            'Ações': ''
        })

    return pd.DataFrame(resultados)


def extrair_andamento_procedimento(proc_obj, protocolo: str) -> pd.DataFrame:
    """
    Extrai os andamentos de um processo a partir do
    RetornoConsultaProcedimento.

    Colunas: Processo, Data/Hora, Unidade, Usuário, Descrição
    """
    colunas = ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']

    if not proc_obj:
        return pd.DataFrame(columns=colunas)

    resultados = []

    # Último andamento
    ultimo = getattr(proc_obj, 'UltimoAndamento', None)
    if ultimo:
        unidade = getattr(ultimo, 'Unidade', None)
        unidade_str = f"{getattr(unidade, 'Descricao', '') or ''} ({getattr(unidade, 'Sigla', '') or ''})" if unidade else ''
        usuario = getattr(ultimo, 'Usuario', None)
        usuario_str = f"{getattr(usuario, 'Nome', '') or ''} ({getattr(usuario, 'Sigla', '') or ''})" if usuario else ''

        resultados.append({
            'Processo': protocolo,
            'Data/Hora': getattr(ultimo, 'DataHora', '') or '',
            'Unidade': unidade_str,
            'Usuário': usuario_str,
            'Descrição': getattr(ultimo, 'Descricao', '') or ''
        })

    # Andamento de geração
    geracao = getattr(proc_obj, 'AndamentoGeracao', None)
    if geracao:
        unidade = getattr(geracao, 'Unidade', None)
        unidade_str = f"{getattr(unidade, 'Descricao', '') or ''} ({getattr(unidade, 'Sigla', '') or ''})" if unidade else ''
        usuario = getattr(geracao, 'Usuario', None)
        usuario_str = f"{getattr(usuario, 'Nome', '') or ''} ({getattr(usuario, 'Sigla', '') or ''})" if usuario else ''

        resultados.append({
            'Processo': protocolo,
            'Data/Hora': getattr(geracao, 'DataHora', '') or '',
            'Unidade': unidade_str,
            'Usuário': usuario_str,
            'Descrição': f"[Geração] {getattr(geracao, 'Descricao', '') or ''}"
        })

    # Andamento de conclusão
    conclusao = getattr(proc_obj, 'AndamentoConclusao', None)
    if conclusao:
        unidade = getattr(conclusao, 'Unidade', None)
        unidade_str = f"{getattr(unidade, 'Descricao', '') or ''} ({getattr(unidade, 'Sigla', '') or ''})" if unidade else ''
        usuario = getattr(conclusao, 'Usuario', None)
        usuario_str = f"{getattr(usuario, 'Nome', '') or ''} ({getattr(usuario, 'Sigla', '') or ''})" if usuario else ''

        resultados.append({
            'Processo': protocolo,
            'Data/Hora': getattr(conclusao, 'DataHora', '') or '',
            'Unidade': unidade_str,
            'Usuário': usuario_str,
            'Descrição': f"[Conclusão] {getattr(conclusao, 'Descricao', '') or ''}"
        })

    if not resultados:
        return pd.DataFrame(columns=colunas)

    return pd.DataFrame(resultados)


# =============================================================================
# Funções de salvamento em Excel (reaproveitadas do script original)
# =============================================================================

def salvar_excel(arquivo: str | Path, dados: WorkbookState) -> bool:
    """
    Salva dados do WorkbookState em Excel, sobrescrevendo todas as planilhas.

    Args:
        arquivo: Caminho do arquivo Excel
        dados: WorkbookState com bloco, processos, andamento, controle

    Returns:
        bool: True se sucesso
    """
    SHEET_CONFIG = {
        'Bloco': {
            'columns': ['Número', 'Sinalizações', 'Atribuição', 'Estado',
                       'Geradora', 'Disponibilização', 'Grupo', 'Descrição', 'Ações']
        },
        'Processos': {
            'columns': ['Número', 'Seq.', 'Processo', 'Documento', 'Tipo',
                       'Assinaturas', 'Anotações', 'Ações']
        },
        'Andamento': {
            'columns': ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']
        },
        'Controle de Processo': {
            'columns': ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']
        }
    }

    total_rows = 0

    try:
        arquivo = Path(arquivo)
        sheet_mapping = {
            'Bloco': getattr(dados, 'bloco', None),
            'Processos': getattr(dados, 'processos', None),
            'Andamento': getattr(dados, 'andamento', None),
            'Controle de Processo': getattr(dados, 'controle', None)
        }

        print("\n📊 Dados recebidos no salvar_excel:")
        for sheet_name, df in sheet_mapping.items():
            if df is not None and isinstance(df, pd.DataFrame) and not df.empty:
                print(f"  ✅ {sheet_name}: {len(df)} linhas, {len(df.columns)} colunas")
            elif df is not None and isinstance(df, pd.DataFrame) and df.empty:
                print(f"  ⚠️  {sheet_name}: DataFrame vazio")
            else:
                print(f"  ❌ {sheet_name}: {type(df).__name__ if df is not None else 'None'}")

        # Prepara dados finais
        final_data = {}
        for sheet_name, df in sheet_mapping.items():
            expected_cols = SHEET_CONFIG[sheet_name]['columns']

            if df is not None and isinstance(df, pd.DataFrame):
                df_clean = df.copy()
                for col in expected_cols:
                    if col not in df_clean.columns:
                        df_clean[col] = ''
                existing_cols = [col for col in expected_cols if col in df_clean.columns]
                if existing_cols:
                    df_clean = df_clean[existing_cols]
                else:
                    df_clean = pd.DataFrame(columns=expected_cols)
            else:
                df_clean = pd.DataFrame(columns=expected_cols)

            final_data[sheet_name] = df_clean
            total_rows += len(df_clean)
            print(f"✅ {len(df_clean)} linhas preparadas para '{sheet_name}'")

        # Garante que todas as planilhas existam
        for sheet_name in SHEET_CONFIG:
            if sheet_name not in final_data:
                final_data[sheet_name] = pd.DataFrame(columns=SHEET_CONFIG[sheet_name]['columns'])

        # Cria diretório se necessário
        arquivo.parent.mkdir(parents=True, exist_ok=True)

        # Salva
        with pd.ExcelWriter(arquivo, engine='openpyxl', mode='w') as writer:
            for sheet_name, df in final_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"\n✅ Dados salvos com sucesso em {arquivo}")
        print(f"   Total: {total_rows} linhas em {len(final_data)} planilhas")
        return True

    except Exception as e:
        print(f"❌ Erro ao salvar Excel: {type(e).__name__}: {str(e)[:200]}")
        traceback.print_exc()
        return False


def salvar_excel_com_data(arquivo: str | Path, dados: WorkbookState) -> bool:
    """
    Salva dados com timestamp diário no nome do arquivo.
    Se o arquivo do dia já existir, concatena novos dados com existentes.
    """
    SHEET_CONFIG = {
        'Bloco': {
            'columns': ['Número', 'Sinalizações', 'Atribuição', 'Estado',
                       'Geradora', 'Disponibilização', 'Grupo', 'Descrição', 'Ações']
        },
        'Processos': {
            'columns': ['Número', 'Seq.', 'Processo', 'Documento', 'Tipo',
                       'Assinaturas', 'Anotações', 'Ações']
        },
        'Andamento': {
            'columns': ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']
        },
        'Controle de Processo': {
            'columns': ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']
        }
    }

    ID_COLUMNS = {
        'Bloco': ['Número'],
        'Processos': ['Processo', 'Seq.'],
        'Andamento': ['Processo', 'Data/Hora', 'Descrição'],
        'Controle de Processo': ['Processo', 'Data/Hora', 'Descrição']
    }

    total_rows = 0
    new_rows = 0

    try:
        arquivo = Path(arquivo)
        date_stamp = datetime.now().strftime("%Y%m%d")
        base_name = arquivo.stem
        extension = arquivo.suffix
        daily_filename = arquivo.parent / f"{base_name}_{date_stamp}{extension}"

        sheet_mapping = {
            'Bloco': getattr(dados, 'bloco', None),
            'Processos': getattr(dados, 'processos', None),
            'Andamento': getattr(dados, 'andamento', None),
            'Controle de Processo': getattr(dados, 'controle', None)
        }

        # Carrega dados existentes se o arquivo já existir
        existing_data = {}
        if daily_filename.exists():
            try:
                with pd.ExcelFile(daily_filename) as xls:
                    for sheet_name in SHEET_CONFIG:
                        if sheet_name in xls.sheet_names:
                            existing_data[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
                        else:
                            existing_data[sheet_name] = pd.DataFrame(columns=SHEET_CONFIG[sheet_name]['columns'])
            except Exception:
                existing_data = {}

        # Prepara e concatena dados
        final_data = {}
        for sheet_name, df in sheet_mapping.items():
            expected_cols = SHEET_CONFIG[sheet_name]['columns']

            if df is not None and isinstance(df, pd.DataFrame):
                df_clean = df.copy()
                for col in expected_cols:
                    if col not in df_clean.columns:
                        df_clean[col] = ''
                existing_cols = [col for col in expected_cols if col in df_clean.columns]
                if existing_cols:
                    df_clean = df_clean[existing_cols]
                else:
                    df_clean = pd.DataFrame(columns=expected_cols)
            else:
                df_clean = pd.DataFrame(columns=expected_cols)

            df_existing = existing_data.get(sheet_name, pd.DataFrame(columns=expected_cols))
            for col in expected_cols:
                if col not in df_existing.columns:
                    df_existing[col] = ''
            df_existing = df_existing[expected_cols] if not df_existing.empty else pd.DataFrame(columns=expected_cols)

            if not df_clean.empty:
                id_columns = ID_COLUMNS.get(sheet_name, [])
                if not df_existing.empty:
                    combined = pd.concat([df_existing, df_clean], ignore_index=True)
                    if id_columns:
                        combined = combined.drop_duplicates(subset=id_columns, keep='first')
                    final_data[sheet_name] = combined
                    new_rows += len(df_clean) - (len(combined) - len(df_existing)) if id_columns else len(df_clean)
                    total_rows += len(combined)
                else:
                    final_data[sheet_name] = df_clean
                    new_rows += len(df_clean)
                    total_rows += len(df_clean)
            else:
                final_data[sheet_name] = df_existing
                total_rows += len(df_existing)

        # Garante planilhas
        for sheet_name in SHEET_CONFIG:
            if sheet_name not in final_data:
                final_data[sheet_name] = pd.DataFrame(columns=SHEET_CONFIG[sheet_name]['columns'])

        daily_filename.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(daily_filename, engine='openpyxl', mode='w') as writer:
            for sheet_name, df in final_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"✅ Dados com data salvos em {daily_filename} ({new_rows} novas linhas)")
        return True

    except Exception as e:
        print(f"❌ Erro ao salvar Excel com data: {type(e).__name__}: {str(e)[:200]}")
        traceback.print_exc()
        return False


# =============================================================================
# Função principal — coleta de dados via API
# =============================================================================

def executar_coleta_api():
    """
    Função principal que substitui o scraping por chamadas à API.
    Fluxo:
      1. Carrega credenciais
      2. Conecta ao Web Service
      3. Itera sobre TODOS os blocos de assinatura configurados
      4. Para cada processo em cada bloco, consulta detalhes
      5. Consolida dados de todos os blocos
      6. Salva em Excel
    """
    print("=" * 60)
    print("  SEI DASHBOARD — VERSÃO API (MÚLTIPLOS BLOCOS)")
    print("  SEI Web Services v5.0.2 (SOAP)")
    print("=" * 60)

    # 1. Carrega credenciais (apenas em memória)
    print("\n📂 Carregando credenciais...")
    credenciais = carregar_credenciais(CREDENTIALS_FILE)

    state = WorkbookState()

    try:
        # 2. Conecta ao Web Service
        print("\n🔌 Conectando ao Web Service...")
        api = SEIAPIClient(credenciais)

        # Limpa credenciais da memória
        import gc
        credenciais = None
        gc.collect()

        # Mostra quantos blocos estão configurados
        qtd_blocos = len(api.blocos_assinatura)
        print(f"   📦 {qtd_blocos} bloco(s) de assinatura configurado(s)")

        # Loop principal
        while True:
            try:
                now = datetime.now()
                current_hour = now.hour
                should_save = 7 <= current_hour < 21

                print(f"\n{'='*60}")
                print(f"  🕐 {now.strftime('%d/%m/%Y %H:%M:%S')}")
                print(f"{'='*60}")

                # Listas para consolidar dados de todos os blocos
                blocos_lista = []
                protocolos_lista = []
                andamentos_lista = []

                # 3. Itera sobre todos os blocos de assinatura
                for b_idx, id_bloco in enumerate(api.blocos_assinatura, 1):
                    print(f"\n📋 [{b_idx}/{qtd_blocos}] Consultando bloco {id_bloco}...")

                    bloco = api.consultar_bloco(id_bloco=id_bloco, retornar_protocolos=True)

                    if bloco is None:
                        print(f"   ⚠️ Bloco {id_bloco} não encontrado ou erro. Pulando...")
                        continue

                    # Extrai dados do bloco
                    df_bloco = extrair_dados_bloco(bloco)
                    if not df_bloco.empty:
                        blocos_lista.append(df_bloco)

                    # Extrai protocolos (processos/documentos do bloco)
                    df_protocolos = extrair_protocolos_bloco(bloco)
                    if not df_protocolos.empty:
                        protocolos_lista.append(df_protocolos)

                    # 4. Para cada processo neste bloco, consulta detalhes e andamentos
                    protocolos = getattr(bloco, 'Protocolos', None) or []

                    if protocolos and isinstance(protocolos, list):
                        processos_no_bloco = [p for p in protocolos
                                              if '/' in (getattr(p, 'ProtocoloFormatado', '') or '')
                                              or '-' in (getattr(p, 'ProtocoloFormatado', '') or '')]
                        total_procs = len(processos_no_bloco)

                        if total_procs > 0:
                            print(f"   🔍 {total_procs} processo(s) no bloco {id_bloco}...")

                            for idx, proto in enumerate(processos_no_bloco, 1):
                                protocolo_formatado = getattr(proto, 'ProtocoloFormatado', '') or ''

                                if not protocolo_formatado:
                                    continue

                                print(f"      [{idx}/{total_procs}] {protocolo_formatado}")

                                # Consulta detalhes do processo
                                proc = api.consultar_procedimento(
                                    protocolo=protocolo_formatado,
                                    ultimo_andamento=True,
                                    andamento_geracao=True,
                                    andamento_conclusao=True
                                )

                                if proc:
                                    df_andamento = extrair_andamento_procedimento(proc, protocolo_formatado)
                                    if not df_andamento.empty:
                                        andamentos_lista.append(df_andamento)

                                # Pequena pausa para não sobrecarregar o servidor
                                time.sleep(0.3)

                # 5. Consolida dados de todos os blocos
                if blocos_lista:
                    state.bloco = pd.concat(blocos_lista, ignore_index=True)
                else:
                    state.bloco = pd.DataFrame(columns=[
                        'Número', 'Sinalizações', 'Atribuição', 'Estado',
                        'Geradora', 'Disponibilização', 'Grupo', 'Descrição', 'Ações'
                    ])

                if protocolos_lista:
                    state.processos = pd.concat(protocolos_lista, ignore_index=True)
                else:
                    state.processos = pd.DataFrame(columns=[
                        'Número', 'Seq.', 'Processo', 'Documento', 'Tipo',
                        'Assinaturas', 'Anotações', 'Ações'
                    ])

                if andamentos_lista:
                    state.andamento = pd.concat(andamentos_lista, ignore_index=True)
                else:
                    state.andamento = pd.DataFrame(columns=['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição'])

                # Controle de Processo — usa os mesmos dados de andamento
                state.controle = state.andamento.copy() if not state.andamento.empty else \
                    pd.DataFrame(columns=['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição'])

                # 6. Salva em Excel
                print(f"\n💾 Salvando dados...")
                salvar_excel(arquivo=ARQUIVO_FONTE, dados=state)

                if should_save:
                    salvar_excel_com_data(arquivo=ARQUIVO_FONTE, dados=state)

                # Copia para destinos
                try:
                    shutil.copy(ARQUIVO_FONTE, ARQUIVO_DESTINO)
                    shutil.copy(ARQUIVO_FONTE, ARQUIVO_DESTINO_2)
                    print(f"\n✅ Cópias salvas em:")
                    print(f"   {ARQUIVO_DESTINO}")
                    print(f"   {ARQUIVO_DESTINO_2}")
                except Exception as e:
                    print(f"⚠️ Erro ao copiar arquivos: {e}")

                # Aguarda o próximo ciclo
                print(f"\n⏳ Aguardando {INTERVALO_LOOP}s até a próxima execução...")
                time.sleep(INTERVALO_LOOP)

            except KeyboardInterrupt:
                raise
            except Exception as e:
                print(f"❌ Erro no loop principal:")
                print(f"   {type(e).__name__}: {str(e)[:200]}")
                traceback.print_exc()
                print(f"\n⏳ Tentando novamente em 60s...")
                time.sleep(60)

    except KeyboardInterrupt:
        print("\n🛑 Script interrompido pelo usuário")
    except Exception as e:
        print(f"❌ Erro fatal: {type(e).__name__}: {str(e)[:200]}")
        traceback.print_exc()
    finally:
        # Limpeza: as credenciais estão em variáveis locais que serão
        # destruídas ao sair do escopo da função.
        print("\n🧹 Limpeza concluída — credenciais removidas da memória.")
        if 'api' in dir():
            del api
        if 'credenciais' in dir():
            del credenciais
        import gc
        gc.collect()


# =============================================================================
# Execução
# =============================================================================

if __name__ == "__main__":
    start_time = time.perf_counter()
    executar_coleta_api()
    elapsed = time.perf_counter() - start_time
    print(f"⏱️ Tempo total de execução: {elapsed:.2f} segundos")