import pyautogui
import time
import numpy as np  # Necessário pelo cv2
import pandas as pd
import logging
import os
import sys
import threading  # Para executar a automação em uma thread separada e manter a GUI responsiva
import tkinter as tk
import cv2  # Necessário pelo pyautogui para reconhecimento de imagem
import json

from queue import Queue
from pyautogui import ImageNotFoundException  # Importa a exceção específica
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from pathlib import Path

# -- DEV CONFIG -- 
DEBUG_MODE = True  # Liga/Desliga logs

# --- Constantes de Configuração ---
IMAGE_DIR = "imagens"  # Pasta onde as imagens estão localizadas
LOG_FILE = "automation_log.log" # Nome do arquivo de log
DEFAULT_CONFIDENCE = 0.8  # Nível de confiança para encontrar imagens (0.0 a 1.0)
DEFAULT_IMAGE_TIMEOUT = 5 if DEBUG_MODE else 10  # Tempo limite em segundos para encontrar uma imagem
DEFAULT_STEP_DELAY = 2  # Atraso em segundos entre as etapas principais
DEFAULT_CLICK_DELAY = 2  # Atraso em segundos após cada clique
TYPE_DELAY = 0.05  # Atraso em segundos entre a digitação de cada caractere
MAX_ATTEMPTS_PER_STEP = 3  # Número máximo de tentativas para cada sub-ação dentro de uma etapa
SAVE_BASE_PATH = r"C:\Users\mb.contabilidade\Desktop\Recibos Destda"

# Nomes das Colunas Obrigatórias na Planilha Excel
COL_EMPRESA = "EMPRESA"
COL_COMENC = "COMENC"
COL_SEMENC = "SEMENC"
COL_DIFATIVO = "DIFATIVO"  # Adicionada a coluna DIFATIVO
COL_DIFAUC = "DIFAUC"
COL_ST = "ST"
COL_STATUS = "STATUS"
COLUMNS = [COL_EMPRESA, COL_COMENC, COL_SEMENC, COL_DIFATIVO, COL_DIFAUC, COL_ST, COL_STATUS]

CONFIG_GERAL = {
    "confianca": DEFAULT_CONFIDENCE,
    "tempo_limite": DEFAULT_IMAGE_TIMEOUT,
    "atraso_etapa": DEFAULT_STEP_DELAY,
    "atraso_clique": DEFAULT_CLICK_DELAY,
    "atraso_digitacao": TYPE_DELAY,
    "max_tentativas": MAX_ATTEMPTS_PER_STEP,
    "ignorar_skipped": True,
    "image_dir": IMAGE_DIR,
    "log_file": LOG_FILE,
    "save_base_path": SAVE_BASE_PATH  # Caminho base para salvar arquivos gerados
}

CONFIG_NOME = {
    "confianca": "Nível de Confiança (0.0 a 1.0)",
    "tempo_limite": "Tempo Limite (segundos)",
    "atraso_etapa": "Atraso entre Etapas (segundos)",
    "atraso_clique": "Atraso após Clique (segundos)",
    "atraso_digitacao": "Atraso de Digitação (segundos)",
    "max_tentativas": "Máximo de Tentativas",
    "ignorar_skipped": "Ignorar Etapas Puladas",
    "image_dir": "Diretório de Imagens",
    "log_file": "Arquivo de Log",
    "save_base_path": "Caminho Base para Salvamento"
}

LIMITES_CONFIG = {
    "confianca": (0.0, 1.0),
    "tempo_limite": (0, 30),
    "atraso_etapa": (0, 10),
    "atraso_clique": (0, 10),
    "atraso_digitacao": (0, 5),
    "ignorar_skipped": (False, True),
    "max_tentativas": (1, 10)
}

# --- Exceções Customizadas ---
class SkipStepException(Exception):
    """Exceção levantada quando o usuário escolhe pular a etapa atual."""
    pass


class InterruptedAutomationException(Exception):
    """Exceção levantada quando o usuário escolhe parar a automação inteira."""
    pass


# --- Classe de Automação Principal ---
class WebAutomation:
    def __init__(self, gui_app=None):
        self.gui = gui_app
        self.config_file_path = Path("config.json")  # Caminho para o arquivo de configuração
        self.config_file_data = {}  # Dicionário para armazenar os dados do arquivo de configuração
        self.config = CONFIG_GERAL.copy()  # Defaults usados quando uma chave não estiver no JSON
        self.df = None  # DataFrame para armazenar os dados do Excel
        self.excel_path = None
        self.current_row_index = -1
        self.stop_requested = False  # Flag para solicitar parada da automação
        self.time_took = 0  # Variável para armazenar o tempo total de execução 
        self.start_time = None
        self.load_config()
        self.image_dir = Path(self.config["image_dir"])
        self.logger = self.setup_logger()

        # Configurações globais do pyautogui
        pyautogui.PAUSE = self.config["atraso_digitacao"]
        pyautogui.FAILSAFE = True  # Move o mouse para o canto superior esquerdo para parar a automação

        if not self.image_dir.is_dir():
            self.logger.error(f"Diretório de imagens '{self.image_dir}' não encontrado. Por favor, crie-o.")
            if self.gui:
                messagebox.showerror("Erro de Inicialização",
                                     f"Diretório de imagens '{self.image_dir}' não encontrado. A automação não pode continuar.")
            sys.exit(1)  # Sai do programa se o diretório de imagens não existir

    def check_tool_command(self):
        if self.stop_requested:
            raise InterruptedAutomationException("Usuário escolheu parar a automação.")

    def load_config(self):
        """Carrega as configurações do arquivo JSON."""
        if not self.config_file_path.exists():
            self.create_config_file()
        if self.config_file_path.exists():
            try:
                with open(self.config_file_path, 'r', encoding='utf-8') as f:
                    file_config = json.load(f)
                if not isinstance(file_config, dict):
                    raise ValueError("o conteúdo deve ser um objeto JSON")
                self._validate_config(file_config)
                self.config.update(file_config)
                self.config_file_data.update(file_config)
            except Exception as e:
                self.config = CONFIG_GERAL.copy()
                self.config_file_data = self.config.copy()
                print(f"Erro na validação do arquivo de configuração: {e}. Usando valores padrão.")
                if self.gui:
                    messagebox.showerror("Erro de Configuração",
                                         f"O arquivo de configuração é inválido:\n{e}\n\n"
                                         "Os valores padrão serão utilizados.")

    @staticmethod
    def _validate_config(file_config):
        """Confere chaves e tipos do JSON antes de incorporá-lo à configuração."""
        expected_keys = set(CONFIG_GERAL)
        actual_keys = set(file_config)
        missing_keys = expected_keys - actual_keys
        extra_keys = actual_keys - expected_keys
        errors = []

        if missing_keys:
            errors.append(f"chaves ausentes: {', '.join(sorted(missing_keys))}")
        if extra_keys:
            errors.append(f"chaves desconhecidas: {', '.join(sorted(extra_keys))}")

        for key in expected_keys & actual_keys:
            expected_value = CONFIG_GERAL[key]
            actual_value = file_config[key]
            if isinstance(expected_value, bool):
                valid_type = isinstance(actual_value, bool)
            else:
                valid_type = type(actual_value) is type(expected_value)
            if not valid_type:
                errors.append(f"tipo incorreto para '{key}': esperado {type(expected_value).__name__}, "
                              f"recebido {type(actual_value).__name__}")

        if errors:
            raise ValueError("; ".join(errors))

    def create_config_file(self):
        """Cria um arquivo de configuração JSON com valores padrão se não existir."""
        if not self.config_file_path.exists() or self.config_file_path.stat().st_size == 0:
            try:
                with open(self.config_file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.config, f, indent=4)
            except Exception as e:
                print(f"Erro ao criar o arquivo de configuração: {e}")
                if self.gui:
                    messagebox.showerror("Erro de Configuração",
                                         f"Não foi possível criar o arquivo de configuração:\n{e}")
        else:
            # Se o arquivo existe, mas não contém todas as chaves, atualiza com os valores padrão
            try:
                with open(self.config_file_path, 'r+', encoding='utf-8') as f:
                    file_config = json.load(f)
                    if not isinstance(file_config, dict):
                        raise ValueError("o conteúdo deve ser um objeto JSON")
                    updated = False
                    for key, value in self.config.items():
                        if key not in file_config:
                            file_config[key] = value
                            updated = True
                    if updated:
                        f.seek(0)
                        json.dump(file_config, f, indent=4)
                        f.truncate()
            except Exception as e:
                print(f"Erro ao atualizar o arquivo de configuração: {e}")
                if self.gui:
                    messagebox.showerror("Erro de Configuração",
                                         f"Não foi possível atualizar o arquivo de configuração:\n{e}")

    def save_config(self, values):
        """Atualiza e salva as configurações informadas pela interface."""
        for key, value in values.items():
            default_value = CONFIG_GERAL[key]
            try:
                if isinstance(default_value, bool):
                    value = bool(value)
                elif isinstance(default_value, int):
                    value = int(value)
                elif isinstance(default_value, float):
                    value = float(value)
                # Verifica se o valor está dentro dos limites permitidos
                if key in LIMITES_CONFIG:
                    min_val, max_val = LIMITES_CONFIG[key]
                    if not (min_val <= value <= max_val):
                        raise ValueError(f"Valor fora dos limites para '{CONFIG_NOME[key]}': {value} (permitido: {min_val} a {max_val})")
            except ValueError as e:
                raise ValueError(f"Valor inválido para '{CONFIG_NOME[key]}': {value}") from e
            self.config[key] = value
            self.config_file_data[key] = value

        with open(self.config_file_path, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4, ensure_ascii=False)

    def default_config(self):
        """Retorna a configuração padrão."""
        return CONFIG_GERAL.copy()

    def reset_counters(self):
        self.gui.reset_counters()  # Reseta os contadores na GUI
        
    # --- Configuração do timer ---

    def start_automation_timer(self):
        """Inicia o timer para medir o tempo total de execução."""
        self.start_time = time.time()
        self.time_took = 0  # Reinicia o tempo total de execução
        self.gui.after(0, self.gui.update_timer_label)  # Atualiza o label do timer na GUI

    def stop_automation_timer(self):
        """Para o timer e calcula o tempo total de execução."""
        if self.start_time is not None:
            self.time_took = time.time() - self.start_time

    # --- Configuração do Logger ---
    def setup_logger(self):
        log_file = self.config["log_file"]
        log_path = Path(log_file)
        if log_path.exists():
            try:
                os.remove(log_path)  # Limpa o log anterior ao iniciar
            except OSError as e:
                print(f"Não foi possível remover o arquivo de log anterior {log_path}: {e}")
                pass  # Continua mesmo que não consiga remover o log

        logger = logging.getLogger(__name__)
        logger.setLevel(logging.INFO)

        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%d/%m/%Y %H:%M:%S')

        if DEBUG_MODE:
            std_handler = logging.StreamHandler(sys.stdout)
            std_handler.setFormatter(formatter)
            logger.addHandler(std_handler)

        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        return logger

    def log_message(self, level, message):
        """Envia mensagens para o logger e para a interface gráfica."""
        if self.gui:
            # Garante que a atualização da GUI ocorra na thread principal
            self.gui.after(0, lambda: self.gui.update_log(message, level))
        if level == "INFO":
            self.logger.info(message)
        elif level == "WARNING":
            self.logger.warning(message)
        elif level == "ERROR":
            self.logger.error(message)

    def _check_image_exists_internal(self, image_name):
        """
        Verifica se um arquivo de imagem existe no diretório de imagens.
        Levanta FileNotFoundError se o arquivo não for encontrado.
        """
        image_path = self.image_dir / image_name
        if not image_path.is_file():
            raise FileNotFoundError(
                f"Arquivo de imagem não encontrado: '{image_path}'. Por favor, verifique se a imagem está na pasta '{self.image_dir}'.")
        return True  # Retorna True se o arquivo existe

    def _find_image_internal(self, image_name, timeout=None, confidence=None, exception=False):
        """
        Tenta encontrar uma imagem na tela.
        Retorna a localização da imagem se encontrada, caso contrário, levanta uma exceção.
        """
        timeout = self.config["tempo_limite"] if timeout is None else timeout
        confidence = self.config["confianca"] if confidence is None else confidence
        self._check_image_exists_internal(image_name)  # Isso levantará FileNotFoundError se a imagem não existir

        image_path = str(self.image_dir / image_name)
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                self.check_tool_command()  # Verifica se o usuário solicitou pular ou parar antes de iniciar a ação

                if self.stop_requested:
                    raise InterruptedAutomationException("Automação interrompida pelo usuário.")

                location = pyautogui.locateOnScreen(image_path, confidence=confidence)
                if location:
                    return location
                
            except ImageNotFoundException:  # Exceção específica do pyautogui quando a imagem não é encontrada
                pass  # Continua procurando
            except InterruptedAutomationException:
                raise
            except Exception as e:  # Captura outros erros inesperados durante a busca
                raise RuntimeError(f"Erro inesperado ao procurar imagem '{image_name}': {e}")
            time.sleep(0.5)  # Verifica a cada 0.5 segundo

        # Se o loop terminar sem encontrar a imagem
        if exception:
            return None
        raise ImageNotFoundException(f"Imagem '{image_name}' não encontrada na tela após {timeout} segundos.")

    def _click_image_internal(self, image_name, timeout=None, confidence=None,
                              offset_x=0, offset_y=0):
        """
        Encontra e clica em uma imagem na tela.
        Permite um offset opcional para clicar em um ponto relativo à imagem.
        Levanta uma exceção se a imagem não for encontrada ou o clique falhar.
        """
        location = self._find_image_internal(image_name, timeout,
                                             confidence)  # Isso levantará uma exceção se a imagem não for encontrada/arquivo ausente
        # Se a localização for encontrada, procede para o clique
        x, y = pyautogui.center(location)

        try:
            pyautogui.click(x + offset_x, y + offset_y)
            time.sleep(self.config["atraso_clique"])
            return True
        except Exception as e:
            raise RuntimeError(f"Erro ao executar o clique em '{image_name}': {e}")

    def _type_text_internal(self, text):
        """Digita um texto na tela."""
        if text is None:
            text = ""  # Garante que não digita None
        pyautogui.write(str(text))
        # Nenhuma exceção esperada para pyautogui.write, mas mantemos a estrutura para consistência
        return True

    def _perform_action_with_retry(self, action_func, action_desc, step_name, *args, **kwargs):
        """
        Executa uma função (ação) com lógica de retentativa e diálogo de erro.
        action_func: A função a ser executada (ex: self._click_image_internal)
        action_desc: Descrição da ação para log e mensagens (ex: "clicar em 'novo_documento.png'")
        step_name: Nome da etapa geral para o diálogo de erro (ex: "ETAPA 1: Abrir movimento da empresa")
        *args, **kwargs: Argumentos para a action_func
        """
        max_attempts = self.config["max_tentativas"]
        for attempt in range(1, max_attempts + 1):
            self.check_tool_command()  # Verifica se o usuário solicitou pular ou parar antes de iniciar a ação

            if self.stop_requested:
                raise InterruptedAutomationException("Automação interrompida pelo usuário.") #STOP 1: Interrompida pelo usuário

            try:
                self.log_message("INFO",
                                 f"Tentativa {attempt}/{max_attempts} para {action_desc} (Etapa: {step_name})")
                result = action_func(*args, **kwargs)
                self.log_message("INFO", f"Sucesso em {action_desc}.")
                return result  # Se bem-sucedido, retorna o resultado
            except (FileNotFoundError, ImageNotFoundException, RuntimeError) as e:
                self.log_message("ERROR",
                                 f"Falha em {action_desc} (Etapa: {step_name}, Tentativa {attempt}/{max_attempts}): {e}")

                # Exibe o diálogo imediatamente na primeira falha para dar controle ao usuário
                action = self._handle_step_error(step_name)

                self.gui.tools_window.geometry("-150+0")
                self.gui.tool_button_retry.config(state=tk.DISABLED)
                self.gui.tool_button_skip.config(state=tk.DISABLED)
                self.gui.hit_error = False
                if action == "Parar":
                    self.stop_automation()
                    if self.gui.tools_window.winfo_exists():
                        self.gui.tools_window.destroy()
                    raise InterruptedAutomationException("Automação interrompida pelo usuário após erro.")
                elif action == "Pular":
                    raise SkipStepException("Usuário optou por pular esta etapa.")
                elif action == "Tentar Novamente":
                    # Usuário escolheu tentar novamente, então o loop continua
                    # Adiciona um pequeno atraso para o usuário ter tempo de reagir/corrigir (ex: minimizar janela)
                    time.sleep(1)
                    self.gui.tools_window.geometry("-150+0")
                    continue  # Continua para a próxima tentativa no loop

        # Se o loop terminar sem retornar (todas as tentativas falharam e usuário optou por tentar novamente)
        raise RuntimeError(
            f"Ação '{action_desc}' falhou após múltiplas tentativas e intervenção do usuário na etapa '{step_name}'.")

    # Funções públicas que as etapas da automação usarão (com retry e tratamento de erro)
    def click_image(self, image_name, **kwargs):
        """Encontra e clica em uma imagem, com retentativas e diálogo de erro."""
        step_name = kwargs.pop('step_context', 'Ação de Clique')
        result = self._perform_action_with_retry(
            self._click_image_internal,
            f"clicar em '{image_name}'",
            step_name,
            image_name,
            **kwargs
        )
        if result is None:
            raise RuntimeError(f"Falha ao clicar em '{image_name}' após múltiplas tentativas na etapa '{step_name}'.")
        return result

    def find_image(self, image_name, **kwargs):
        """Encontra uma imagem, com retentativas e diálogo de erro. Retorna a localização."""
        step_name = kwargs.pop('step_context', 'Ação de Busca de Imagem')
        result = self._perform_action_with_retry(
            self._find_image_internal,
            f"encontrar '{image_name}'",
            step_name,
            image_name,
            **kwargs
        )
        if result is None and kwargs.get('exception', True):
            return None
        if result is None:
            raise RuntimeError(f"Imagem '{image_name}' não encontrada após múltiplas tentativas na etapa '{step_name}'.")
        return result

    def type_text(self, text, **kwargs):
        """Digita um texto, com retentativas e diálogo de erro."""
        step_name = kwargs.pop('step_context', 'Ação de Digitação')
        result = self._perform_action_with_retry(
            self._type_text_internal,
            f"digitar '{text}'",
            step_name,
            text,
            **kwargs
        )
        if result is None:
            raise RuntimeError(f"Falha ao digitar '{text}' após múltiplas tentativas na etapa '{step_name}'.")
        return result

    def format_number_for_typing(self, value):
        """
        Formata um valor (potencialmente numérico) para ser digitado,
        garantindo duas casas decimais e vírgula como separador decimal.
        """
        if pd.isna(value):
            return ""
        try:
            # Tenta converter para float para formatar como número
            num_val = float(value)
            # Formata para duas casas decimais e substitui ponto por vírgula
            return f"{num_val:.2f}".replace('.', ',')
        except (ValueError, TypeError):
            # Se não for um número ou já for uma string (ex: 'abc'), retorna como string original
            return str(value)

    def _handle_step_error(self, step_name):
        """
        Exibe um diálogo de erro na GUI e retorna a ação escolhida pelo usuário.
        Esta função bloqueará a thread de automação até que o usuário responda.
        """
        self.log_message("ERROR", f"Erro crítico na etapa: {step_name}. Solicitando ação do usuário...")
        # Chamada para o diálogo na thread principal da GUI
        response = self.gui.show_tools_error(step_name)
        return response

    def _update_excel_status(self, row_index, status):
        """Atualiza o status da linha na planilha Excel e salva o arquivo."""
        if self.df is not None and self.excel_path:
            # Verifica se a linha existe antes de tentar atualizar
            if row_index < 0 or row_index >= len(self.df):
                self.log_message("ERROR", f"Tentativa de atualizar status para índice de linha inválido: {row_index}")
                return

            original_status = self.df.at[row_index, COL_STATUS]
            self.df.at[row_index, COL_STATUS] = status
            self.log_message("INFO",
                             f"Status da linha {row_index + 2} (Empresa: {self.df.at[row_index, COL_EMPRESA]}) atualizado de '{original_status}' para '{status}'.")
            try:
                self.df.to_excel(self.excel_path, index=False)
                self.log_message("INFO", f"Planilha '{self.excel_path}' salva com sucesso.")
                self.gui.after(0, self.gui.update_status_column, row_index, status)  # Atualiza a tabela na GUI
            except Exception as e:
                self.log_message("ERROR",
                                 f"Erro ao salvar a planilha '{self.excel_path}': {e}. Por favor, feche o arquivo se estiver aberto.")
                if self.gui:
                    self.gui.after(0, self.gui.update_status_column, row_index, status)  # Atualiza a tabela na GUI
                    self.gui.after(0, messagebox.showerror, "Erro ao Salvar Planilha",
                                   f"Não foi possível salvar a planilha:\n{e}\nPor favor, feche o arquivo se estiver aberto.")

    def load_excel(self, file_path):
        """Carrega a planilha Excel e valida as colunas obrigatórias."""
        self.excel_path = Path(file_path)
        if not self.excel_path.is_file():
            self.log_message("ERROR", f"Arquivo Excel não encontrado: {self.excel_path}")
            if self.gui:
                messagebox.showerror("Erro ao Carregar Excel", "Arquivo Excel não encontrado.")
            return False

        try:
            self.df = pd.read_excel(self.excel_path)
            required_columns = COLUMNS
            missing_cols = [col for col in required_columns if col not in self.df.columns]
            if missing_cols:
                self.log_message("ERROR",
                                 f"Planilha Excel não contém as colunas obrigatórias: {', '.join(missing_cols)}")
                if self.gui:
                    messagebox.showerror("Erro de Colunas",
                                         f"A planilha não contém as colunas obrigatórias:\n{', '.join(missing_cols)}")
                self.df = None
                return False

            # Inicializa a coluna STATUS com string vazia para evitar NaN
            # Garante que o status seja string, mesmo se estiver vazio ou NaN
            self.df[COL_STATUS] = self.df[COL_STATUS].fillna("").astype(str).str.strip()
            self.log_message("INFO",
                             f"Planilha '{self.excel_path}' carregada com sucesso. {len(self.df)} linhas encontradas.")
            if self.gui:
                self.gui.after(0, self.gui.update_status_label, f"Planilha carregada: {self.excel_path.name}")
            return True
        except Exception as e:
            self.log_message("ERROR", f"Erro ao carregar a planilha Excel: {e}")
            if self.gui:
                messagebox.showerror("Erro ao Carregar Excel", f"Erro ao carregar a planilha:\n{e}")
            self.df = None
            return False

    def _is_sem_movimento(self, row_data):
        """Verifica se a empresa é 'sem movimento' com base nos dados da linha."""
        # Considera "sem movimento" se todas as colunas especificadas são nulas/vazias - AGORA INCLUINDO DIFATIVO
        return pd.isna(row_data[COL_COMENC]) and \
               pd.isna(row_data[COL_SEMENC]) and \
               pd.isna(row_data[COL_DIFAUC]) and \
               pd.isna(row_data[COL_ST]) and \
               pd.isna(row_data[COL_DIFATIVO])  # <-- CORREÇÃO AQUI

    def _create_folder_structure(self, base_path, folder_name):
        """Cria a estrutura de pastas para salvar arquivos gerados."""
        full_path = Path(base_path) / folder_name
        try:
            full_path.mkdir(parents=True, exist_ok=True)
            self.log_message("INFO", f"Estrutura de pastas criada/verificada: {full_path}")
            return full_path
        except Exception as e:
            self.log_message("ERROR", f"Erro ao criar estrutura de pastas '{full_path}': {e}")
            return None

    def run_automation_for_row(self, row_index, row_data):
        """Executa a automação para uma única linha da planilha."""
        if self.stop_requested:
            self.log_message("INFO", "Automação interrompida pela solicitação do usuário.")
            return

        self.current_row_index = row_index
        empresa = str(row_data[COL_EMPRESA]).strip()
        competencia = f"{self.gui.return_competencia()}"  # Obtém a competência da GUI

        self._update_excel_status(row_index, "processing")
        self.log_message("INFO", f"Iniciando automação para Empresa: {empresa}, Competência: {competencia}")
        self.gui.after(0, self.gui.update_progress, row_index + 1)
        self.gui.after(0, self.gui.update_status_label, f"Processando: {empresa} ({competencia})")

        is_sem_movimento = self._is_sem_movimento(row_data)
        movimento_type = "sem movimento" if is_sem_movimento else "com movimento"
        self.log_message("INFO", f"Tipo de movimento detectado para {empresa}: {movimento_type}")

        etapa1_success = False  # Flag para indicar sucesso da Etapa 1
        etapa2_success = False  # Flag para indicar sucesso da Etapa 2
        etapa3_success = False  # Flag para indicar sucesso da Etapa 3
        etapa4_success = False  # Flag para indicar sucesso da Etapa 4
        etapa5_success = False  # Flag para indicar sucesso da Etapa 5

        def processo_reinicio(step_name):
            if self.find_image("registro_estado_insercao.png", step_context=step_name, timeout=5, exception=True):
                self.click_image("nao.png", step_context=step_name)
            if self.find_image("fechartelamov.png", step_context=step_name, timeout=5, exception=True):
                self.click_image("fechartelamov.png", step_context=step_name)

            self.click_image("iniciar.png", step_context=step_name)
            self.click_image("fechardocumento.png", step_context=step_name)
            self.click_image("fechartelamov.png", step_context=step_name)
            if self.find_image("registro_estado_insercao.png", step_context=step_name, timeout=5, exception=True):
                self.click_image("nao.png", step_context=step_name)
            time.sleep(1)

        def goes_to_step1(step_name):
            processo_reinicio(step_name)
            self.click_image("novo_documento.png", step_context=step_name)

        def step1(step_name):
            # ETAPA 1: Abrir movimento da empresa
            goes_to_step1(step_name)
            time.sleep(4)  # Espera especificada pelo usuário
            self.click_image("seta_baixo.png", step_context=step_name)
            self.type_text(empresa, step_context=step_name)
            pyautogui.press('enter')
            time.sleep(self.config["atraso_etapa"])  # Espera para o sistema reagir
            self.type_text(competencia, step_context=step_name)
            self.click_image("carregar_periodo.png", step_context=step_name)
            pyautogui.press('down')  # Seta para baixo
            pyautogui.press('tab')
            self.type_text("0", step_context=step_name)  # Valor fixo "0"
            pyautogui.press('tab')

            # Clicar relativo 10 px abaixo da imagem conteudo.png
            conteudo_loc = self.find_image("conteudo.png", step_context=step_name)
            pyautogui.click(pyautogui.center(conteudo_loc).x, pyautogui.center(conteudo_loc).y + 10)
            self.log_message("INFO", f"Clicou 10px abaixo de 'conteudo.png'.")
            time.sleep(self.config["atraso_clique"])

            if is_sem_movimento:
                self.type_text("1", step_context=step_name)  # Sem movimento
            else:
                self.type_text("0", step_context=step_name)  # Com movimento
            pyautogui.press('tab')
            
            self.click_image("confirmar.png", step_context=step_name)
            if self.find_image("aviso_ja_existe.png", step_context=step_name, timeout=5, exception=True):
                self.log_message("WARNING", f"Empresa com movimento já cadastrado em '{step_name}'.")
                self.click_image("ok.png", step_context=step_name)
                self.click_image("cancelar.png", step_context=step_name)
                self.click_image("fechar_tela.png", step_context=step_name)
                raise Exception("Movimento já cadastrado")  # Indica que a etapa 1 falhou devido a movimento existente
              # Marca a Etapa 1 como bem-sucedida
            return True

        step_name = "ETAPA 1: Abrir movimento da empresa"
        if not hasattr(self.gui, 'tools_window') or not self.gui.tools_window.winfo_exists():
            self.gui.after(0,lambda: self.gui.open_tools_toplevel())
        try:
            while not etapa1_success:
                etapa1_success = step1(step_name)
            # Fim da etapa 1

            while not etapa2_success:
                # ETAPA 2: Condicional (Sem Movimento ou Com Movimento)
                step_name = f"ETAPA 2: {'Sem' if is_sem_movimento else 'Com'} Movimento"
                time.sleep(6)  # Espera especificada
                self.click_image("ABRIR.png", step_context=step_name)

                if not is_sem_movimento:  # Lógica APENAS para "Com Movimento"
                    self.click_image("destda.png", step_context=step_name)
                    self.click_image("alterar.png", step_context=step_name)

                    # Garante que os valores numéricos sejam formatados com vírgula para digitação
                    st_val = self.format_number_for_typing(row_data[COL_ST])
                    comenc_val = self.format_number_for_typing(row_data[COL_COMENC])
                    semenc_val = self.format_number_for_typing(row_data[COL_SEMENC])
                    difauc_val = self.format_number_for_typing(row_data[COL_DIFAUC])
                    difativo_val = self.format_number_for_typing(row_data[COL_DIFATIVO]) 
                    
                    self.type_text(st_val, step_context=step_name)  # Digitar valor ST
                    self.click_image("icms_entrada.png", step_context=step_name)  # Clicar em icms_entrada.png
                    self.type_text(comenc_val, step_context=step_name)  # Digitar valor COMENC
                    pyautogui.press('tab')
                    self.type_text(semenc_val, step_context=step_name)  # Digitar valor SEMENC
                    pyautogui.press('tab')
                    self.type_text(difativo_val, step_context=step_name)  # <-- CORREÇÃO: Digitar valor DIFATIVO
                    pyautogui.press('tab')
                    self.type_text(difauc_val, step_context=step_name)  # Digitar valor DIFAUC
                    pyautogui.press('tab')

                    self.click_image("confirmar.png", step_context=step_name)  # (Etapa 2 'Com Movimento' final)
                    self.click_image("fechar_tela.png", step_context=step_name)
                etapa2_success = True  # Marca a Etapa 2 como bem-sucedida

            while not etapa3_success:
                # ETAPA 3: Gerar Documento
                step_name = "ETAPA 3: Gerar Documento"
                self.click_image("encerrar.png", step_context=step_name)
                self.click_image("gerar_documento.png", step_context=step_name)
                self.click_image("iniciar_processamento.png", step_context=step_name)  # (Gerar Documento)

                if self.find_image("nao_encontrou_certificadodigital.png", step_context=step_name, timeout=5, exception=True):
                    self.click_image("ok.png", step_context=step_name)
                    self.click_image("cancelar_certificado.png", step_context=step_name)
                    self.click_image("ok.png", step_context=step_name)
                    self.log_message("WARNING", f"Certificado digital não encontrado em '{step_name}'.")
                    raise RuntimeError("Certificado digital não encontrado.")

                self.click_image("selecionar_certificado.png", step_context=step_name)
                self.click_image("autenticar.png", step_context=step_name)

                self.log_message("INFO", "Aguardando 'processo_finalizado.png'...")
                self.find_image("processo_finalizado.png", timeout=30, step_context=step_name)  # Aguarda até 30s
                self.click_image("ok.png", step_context=step_name)  # após processo finalizado
                etapa3_success = True  # Marca a Etapa 3 como bem-sucedida

            while not etapa4_success:
                # ETAPA 4: Transmissão
                step_name = "ETAPA 4: Transmissão"
                self.click_image("transmitir.png", step_context=step_name)
                self.click_image("iniciar_processamento.png", step_context=step_name)  # (Transmissão)

                # NOVO: Aguardar finalização da transmissão e clicar OK
                self.log_message("INFO", "Aguardando 'processo_finalizado.png' após transmissão...")
                self.find_image("processo_finalizado.png", timeout=30, step_context=step_name)  # Aguarda até 30s
                self.click_image("ok.png", step_context=step_name)  # Clica OK
                etapa4_success = True  # Marca a Etapa 4 como bem-sucedida

            while not etapa5_success:
                # ETAPA 5: Salvar Recibo
                step_name = "ETAPA 5: Salvar Recibo"
                self.log_message("INFO", "Iniciando ETAPA 5: Salvar Recibo.")

                time.sleep(3)  # Esperar por 3 segundos
                self.click_image("exportar.png", step_context=step_name)

                # Definir caminho de salvamento
                competencia_formatted = competencia.replace("/", "-")
                save_folder_path = self._create_folder_structure(self.config["save_base_path"], competencia_formatted)

                if not save_folder_path:
                    raise RuntimeError(
                        f"Não foi possível criar/acessar a pasta de destino para salvar o recibo: {self.config['save_base_path']}{competencia_formatted}")

                # Interagir com o diálogo de salvar arquivo
                # Tentar focar a barra de endereços (Alt+D) e digitar o caminho
                pyautogui.hotkey('alt', 'd')
                self.type_text(str(save_folder_path), step_context=step_name)
                pyautogui.press('enter')

                # Tentar focar o campo de nome do arquivo (Alt+N) e digitar o nome da empresa
                pyautogui.hotkey('alt', 'n')
                self.type_text(empresa, step_context=step_name)  # Digitar o nome da empresa

                self.click_image("salvar.png", step_context=step_name)

                # Clicar ENTER 3 vezes com um segundo entre cada clique
                for _ in range(3):
                    pyautogui.press('enter')
                    time.sleep(1)

                self.click_image("fechar_tela.png", step_context=step_name)
                self.click_image("iniciar.png", step_context=step_name)
                self.click_image("fechar_documento.png", step_context=step_name)
                self.click_image("fechar_tela.png", step_context=step_name)

                # Se todas as etapas concluídas com sucesso
                self._update_excel_status(row_index, "ok")
                self.log_message("INFO",
                                f"Automação concluída com sucesso para Empresa: {empresa}, Competência: {competencia}")
                self.gui.count_ok += 1  # Incrementa o contador de OKs
                time.sleep(self.config["atraso_etapa"])  # Atraso entre as linhas da planilha
                etapa5_success = True  # Marca a Etapa 5 como bem-sucedida

        except InterruptedAutomationException:
            # Se o usuário escolheu parar a automação
            self._update_excel_status(row_index, "error")  # Marca a linha atual como erro
            self.gui.count_error += 1  # Incrementa o contador de erros
            self.stop_requested = True  # Define a flag de parada
            return  # Sai da execução para esta linha
        except SkipStepException:
            # Se o usuário escolheu pular esta etapa
            self.gui.count_skip += 1  # Incrementa o contador de pulados
            self._update_excel_status(row_index, "skipped")  # Marca a linha como pulada
            return  # Sai da execução para esta linha e continua para a próxima
        except Exception as e:
            # Captura qualquer outra exceção não tratada pelas chamadas a _perform_action_with_retry
            # Isso indicaria um erro de lógica no script ou um cenário não previsto.
            self.log_message("ERROR",
                             f"Erro fatal ou não-tratado durante a automação para {empresa} - {competencia}: {e}")
            self._update_excel_status(row_index, "error")
            self.gui.count_error += 1  # Incrementa o contador de erros
            return  # Sai da execução para esta linha e continua para a próxima

    def start_automation(self):
        """Inicia o processo de automação iterando sobre as linhas da planilha."""
        if self.df is None or self.df.empty:
            self.log_message("WARNING", "Nenhum dado na planilha para processar ou planilha não carregada.")
            if self.gui:
                self.gui.after(0, messagebox.showwarning, "Automação",
                               "Por favor, carregue uma planilha Excel antes de iniciar.")
            return

        self.stop_requested = False
        self.gui.after(0, self.gui.set_automation_running, True)
        self.log_message("INFO", "Iniciando processo de automação...")
        self.start_automation_timer()  # Inicia o timer para medir o tempo total de execução
        self.reset_counters()  # Reinicia os contadores de OK, Skip e Error

        try:
            for index, row in self.df.iterrows():
                if self.stop_requested:
                    self.log_message("INFO", "Automação interrompida pelo usuário.")
                    break

                current_status = row[COL_STATUS]
                # Ignora linhas já processadas ou puladas
                if isinstance(current_status, str):
                    if current_status.lower() == "ok":
                        self.gui.count_ok += 1
                        self.log_message("INFO", f"Linha {index + 2} (Empresa: {row[COL_EMPRESA]}) já está com status '{current_status}'. Pulando para a próxima.")
                        continue
                    elif current_status.lower() == "skipped":
                        self.gui.count_skip += 1
                        print("valor do self.config['ignorar_skipped']:", self.config["ignorar_skipped"])
                        if self.config["ignorar_skipped"]:
                            self.log_message("INFO", f"Linha {index + 2} (Empresa: {row[COL_EMPRESA]}) já está com status '{current_status}'. Pulando para a próxima.")
                            continue
                    elif current_status.lower() == "error":
                        self.gui.count_error += 1
                    self.gui.after(0, self.gui.update_progress, index + 1)
                self.run_automation_for_row(index, row)

            self.log_message("INFO", "Processo de automação finalizado.")
        except Exception as e:
            self.log_message("ERROR", f"Erro inesperado durante a execução principal da automação: {e}")
            if self.gui:
                self.gui.deiconify()  # Traz a janela da GUI para frente
                self.gui.after(0, messagebox.showerror, "Erro Fatal na Automação",
                               f"Um erro inesperado ocorreu durante a automação:\n{e}\nPor favor, verifique os logs.")
            # Marca a linha atual como erro se a automação parou inesperadamente no meio
            if self.current_row_index != -1 and not self.stop_requested:
                self._update_excel_status(self.current_row_index, "error")
        finally:
            self.stop_automation_timer()  # Para o timer e calcula o tempo total de execução
            self.gui.after(0, self.gui.set_automation_running, False)
            self.gui.after(0, self.gui.update_progress, len(self.df))  # Garante que a barra de progresso atinja o fim
            self.gui.after(0, self.gui.update_timer_label)
            self.gui.after(0, self.gui._update_cards())
            self.gui.after(0, self.gui.update_status_label, "Automação finalizada.")
            if hasattr(self.gui, 'tools_window') and self.gui.tools_window.winfo_exists():
                self.gui.tools_window.destroy()  # Fecha a janela de ferramentas se estiver aberta
            self.gui.deiconify()
            self.log_message("INFO", f"Tempo total de execução: {self.gui.format_time(self.time_took)}")
            
    def stop_automation(self):
        """Define a flag para solicitar a parada da automação."""
        self.log_message("INFO", "Solicitação de parada da automação recebida.")
        self.stop_requested = True
        # A GUI será atualizada quando a thread de automação reconhecer a parada


# --- Classe da Interface Gráfica (GUI) ---
class AutomationGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automação Web - Adapta (One)")
        self.geometry("")
        self.resizable(False, False)
        self.automation = WebAutomation(self)  # Instancia a lógica de automação, passando a si mesmo (GUI)
        self.config_entries = {}
        self.excel_file_path = None
        self.running = False  # Estado da automação
        self.tools_response_queue = Queue()
        self.hit_error = False
        self._centralize_after_id = None

        # Contagem para os cards
        self.count_total = 0
        self.count_ok = 0
        self.count_skip = 0
        self.count_error = 0

        self._create_widgets()
        self._setup_layout()

        # Garante que o diretório de imagens exista ao iniciar a GUI
        image_dir = self.automation.image_dir
        if not image_dir.is_dir():
            image_dir.mkdir(exist_ok=True)
            self.update_log(f"Diretório '{image_dir}' criado (se não existia).", "INFO")

    def _create_widgets(self):
        """Cria todos os componentes da interface gráfica."""
        # CONFIGURAÇÕES
        self.control_frame = ttk.LabelFrame(self, text="Controles")
        self.control_frame.pack(padx=10, pady=10, fill="x")

        self.btn_load_excel = ttk.Button(self.control_frame, text="Carregar Planilha Excel",
                                         command=self.load_excel_file)
        self.btn_start = ttk.Button(self.control_frame, text="Iniciar Automação",
                                    command=self.start_automation_threaded, state=tk.DISABLED)
        self.btn_stop = ttk.Button(self.control_frame, text="Parar Automação", command=self.stop_automation,
                                   state=tk.DISABLED)
        self.open_tools_toplevel_window = ttk.Button(self.control_frame, text="Abrir Ferramentas", state=tk.DISABLED)

        # CONTROLES
        self.config_frame = ttk.LabelFrame(self, text="Configurações")
        self.config_frame.pack(padx=10, pady=10, fill="x")

        last_month = datetime.today().month - 1 if datetime.today().month > 1 else 12

        current_year = datetime.today().year
        self.spinbox_ano_var = tk.StringVar(value=str(current_year))
        self.spinbox_mes_var = tk.StringVar(value=f"{last_month:02d}")  # Mês anterior como padrão

        self.config_label_competencia = ttk.Label(self.config_frame, text="Competência (MM/AAAA):")
        
        self.frame_competencia = ttk.Frame(self.config_frame)

        def validate_spinbox(new_value, minimum, maximum):
            if new_value == "":
                return True
            if not new_value.isdigit():
                return False

            minimum = int(minimum)
            maximum = int(maximum)
            value = int(new_value)

            if len(new_value) == len(str(maximum)):
                return minimum <= value <= maximum
            return value <= maximum
        
        self.spinbox_mes = ttk.Spinbox(
            self.frame_competencia,
            from_=1,
            to=12,
            width=5,
            format="%02.0f",
            textvariable=self.spinbox_mes_var,
            validate="key",
            validatecommand=(self.register(validate_spinbox), "%P", 1, 12),
        )
        self.spinbox_ano = ttk.Spinbox(
            self.frame_competencia,
            from_=current_year - 100,
            to=current_year,
            width=7,
            format="%04.0f",
            textvariable=self.spinbox_ano_var,
            validate="key",
            validatecommand=(self.register(validate_spinbox), "%P", current_year - 100, current_year),
        )

        self.more_config = ttk.Button(self.config_frame, text="⚙ Config", cursor="hand2", width=10, command=self.open_config_window)

        self.status_label = ttk.Label(self.config_frame, text="Aguardando planilha...", foreground="red")
        self.progress_bar = ttk.Progressbar(self.config_frame, orient="horizontal", length=300, mode="determinate")
        self.progress_label = ttk.Label(self.config_frame, text="0/0")
        self.timer_label = ttk.Label(self.config_frame, text="Tempo decorrido: 00:00:00")

        # Cards de Resumo
        cards_frame = ttk.Frame(self.config_frame)
        cards_frame.grid(row=5, column=0, columnspan=4, pady=10, sticky="ew")
        for col in range(4):
            cards_frame.columnconfigure(col, weight=1)

        self.lbl_total = self._make_card(cards_frame, "Total",     "–", 0, "#555555")
        self.lbl_ok    = self._make_card(cards_frame, "OK",        "–", 1, "#3B6D11")
        self.lbl_skip  = self._make_card(cards_frame, "Pulado",    "–", 2, "#854F0B")
        self.lbl_err   = self._make_card(cards_frame, "Erro",      "–", 3, "#A32D2D")

        # Dropdown tabela
        self.button_hide_table = ttk.Button(self.config_frame, text="Ocultar Tabela", command=self.hide_table)
        # Dropdown logs
        self.button_hide_logs = ttk.Button(self.config_frame, text="Ocultar Logs", command=self.hide_logs)

        # Frame para os logs
        self.log_frame = ttk.LabelFrame(self, text="Logs de Operação")
        self.log_frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.log_text = tk.Text(self.log_frame, state="disabled", wrap="word", height=20, width=67, bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 10), insertbackground="white")
        self.log_text_scroll = ttk.Scrollbar(self.log_frame, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=self.log_text_scroll.set)

        self.log_text.tag_config("INFO", foreground="#9cdcfe")
        self.log_text.tag_config("WARNING", foreground="#dcdcaa")
        self.log_text.tag_config("ERROR", foreground="#f48771")
        self.log_text.tag_config("TIME", foreground="#606060")

        # Frame para a tabela de dados
        self.table_frame = ttk.LabelFrame(self, text="Dados da Planilha")
        self.table_frame.pack(padx=10, pady=10, fill="both", expand=True)
        self.table = ttk.Treeview(self.table_frame, columns=COLUMNS, show="headings", height=10)
        widths = [200, 60, 60, 60, 50, 40, 70]  # Larguras para cada coluna
        for col in COLUMNS:
            self.table.heading(col, text=col)
            self.table.column(col, width=widths[COLUMNS.index(col)], anchor="center")

        self.scroll_table = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=self.scroll_table.set)

    def open_config_window(self):
        """Abre a janela de configurações adicionais."""
        if hasattr(self, 'config_window') and self.config_window.winfo_exists():
            self.config_window.lift()  # Traz a janela existente para frente
            return

        self.config_window = tk.Toplevel(self)
        self.config_window.title("Configurações Adicionais")
        self.config_window.resizable(False, False)
        self.config_window.columnconfigure(1, weight=1)
        self.config_entries.clear()

        frame = ttk.LabelFrame(self.config_window, text="Configurações Adicionais")
        frame.grid(row=0, column=0, columnspan=3, sticky="nsew", padx=2, pady=2)

        row = 1
        for key, value in self.automation.config.items():
            ttk.Label(
                frame,
                text=f"{CONFIG_NOME[key]}:"
            ).grid(row=row, column=0, sticky="w", padx=2, pady=2)

            if key == "ignorar_skipped":
                var = tk.BooleanVar(value=value)
                check = ttk.Checkbutton(frame, variable=var)
                check.grid(row=row, column=1, sticky="w", padx=2, pady=2)
                self.config_entries[key] = var
                row += 1
                continue

            entry = ttk.Entry(frame, width=30)
            entry.insert(0, str(value))
            entry.grid(row=row, column=1, sticky="w", padx=2, pady=2)
            self.config_entries[key] = entry

            if key in ["save_base_path", "image_dir"]:
                ttk.Button(
                    frame,
                    text="📁",
                    command=lambda k=key: self.select_folder(k),
                    width=5
                ).grid(row=row, column=2, sticky="w", padx=2, pady=2)
            row += 1

        buttons_frame = ttk.Frame(self.config_window)
        buttons_frame.grid(row=row, column=0, columnspan=3, padx=2, pady=(8, 4))

        ttk.Button(
            buttons_frame,
            text="Fechar",
            command=self.config_window.destroy
        ).grid(row=0, column=0, padx=2)
        ttk.Button(
            buttons_frame,
            text="Salvar",
            command=self.save_config
        ).grid(row=0, column=1, padx=2)
        ttk.Button(
            buttons_frame,
            text="Restaurar Padrão",
            command=self.restore_default_config
        ).grid(row=0, column=2, padx=2)

    def select_folder(self, key):
        """Seleciona uma pasta e atualiza o campo correspondente."""
        folder_path = filedialog.askdirectory(
            title=f"Selecione {CONFIG_NOME[key].lower()}"
        )
        if folder_path:
            self.config_entries[key].delete(0, tk.END)
            self.config_entries[key].insert(0, folder_path)

    def restore_default_config(self):
        """Restaura as configurações para os valores padrão."""
        for key, value in self.automation.default_config().items():
            self.config_entries[key].delete(0, tk.END)
            self.config_entries[key].insert(0, str(value))
        
    def save_config(self):
        """Salva as configurações adicionais no arquivo de configuração."""
        values = {key: entry.get() for key, entry in self.config_entries.items()}
        try:
            self.automation.save_config(values)
            messagebox.showinfo("Configurações", "Configurações salvas com sucesso.")
            self.config_window.destroy()
        except (OSError, ValueError) as e:
            messagebox.showerror("Erro de Configuração", f"Não foi possível salvar as configurações:\n{e}")

    def _make_card(self, parent, label, value, col, color):
        frm = ttk.Frame(parent, relief="groove", padding=8)
        frm.grid(row=0, column=col, sticky="ew", padx=4)
        num_lbl = ttk.Label(frm, text=value,
                            font=("Segoe UI", 18, "bold"),
                            foreground=color)
        num_lbl.pack()
        ttk.Label(frm, text=label,
                  font=("Segoe UI", 9),
                  foreground="gray").pack()
        return num_lbl

    def _update_cards(self):
        self.lbl_total.config(text=str(self.count_total))
        self.lbl_ok.config(   text=str(self.count_ok))
        self.lbl_skip.config( text=str(self.count_skip))
        self.lbl_err.config(  text=str(self.count_error))

    def reset_counters(self):
        """Reseta os contadores de OK, Skip e Error."""
        self.count_ok = 0
        self.count_skip = 0
        self.count_error = 0
        self._update_cards()

    def _setup_layout(self):
        """Define o layout dos componentes na grade."""
        self.btn_load_excel.grid(row=0, column=0, padx=5, pady=5)
        self.btn_start.grid(row=0, column=1, padx=5, pady=5)
        self.btn_stop.grid(row=0, column=2, padx=5, pady=5)
        self.open_tools_toplevel_window.grid(row=0, column=3, padx=5, pady=5)
        self.config_label_competencia.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.frame_competencia.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.spinbox_mes.grid(row=0, column=0, padx=(0,5), pady=5, sticky="w")
        self.spinbox_ano.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.more_config.grid(row=1, column=3, padx=5, pady=5, sticky="e")
        self.status_label.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        self.progress_bar.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        self.progress_label.grid(row=3, column=2, padx=5, pady=5)
        self.timer_label.grid(row=3, column=3, padx=5, pady=5, sticky="e")
        self.button_hide_table.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        self.button_hide_logs.grid(row=4, column=2, columnspan=2, padx=5, pady=5, sticky="ew")
        self.log_text.pack(side="left", fill="both", expand=True)
        self.log_text_scroll.pack(side="right", fill="y")
        self.scroll_table.pack(side="right", fill="y")
        self.table.pack(fill="both", expand=True)
        self.hide_logs()  # Inicia com os logs ocultos para uma interface mais limpa
        self.hide_table()  # Inicia com a tabela oculta para uma interface mais limpa

    def return_competencia(self):
        """Retorna a competência no formato MM/AAAA."""
        mes = self.spinbox_mes_var.get()
        ano = self.spinbox_ano_var.get()
        if not mes or not ano:
            raise ValueError("Mês e ano da competência devem ser preenchidos.")
        mes_int = int(mes)
        ano_int = int(ano)
        current_year = datetime.today().year
        if not 1 <= mes_int <= 12:
            raise ValueError("O mês da competência deve estar entre 01 e 12.")
        if not current_year - 100 <= ano_int <= current_year:
            raise ValueError("O ano da competência está fora do intervalo permitido.")
        return f"{mes_int:02d}/{ano_int:04d}"

    def load_table_data(self, df):
        """Carrega os dados do DataFrame na tabela Treeview."""
        # Limpa a tabela antes de carregar novos dados
        self.count_total = 0  # Reinicia o contador total

        for item in self.table.get_children():
            self.table.delete(item)

        self.table.tag_configure('par', background="#a3a3a3")
        self.table.tag_configure('impar', background='white')
        self.table.tag_configure('error', background='#ffcccc')  # Vermelho claro para erros
        self.table.tag_configure('skipped', background='#ffffcc')  # Amarelo claro para pulados

        # Insere os dados do DataFrame na tabela
        df = df.fillna(0)  # Substitui NaN por 0
        for _, row in df.iterrows():
            status = str(row[COL_STATUS]).lower()
            if status == "error":
                cor_tag = 'error'
            elif status == "skipped":
                cor_tag = 'skipped'
            else:
                cor_tag = 'par' if _ % 2 == 0 else 'impar'

            self.count_total += 1  # Incrementa o contador total

            self.table.insert("", "end", values=(
                row[COL_EMPRESA],  row[COL_COMENC], row[COL_SEMENC], row[COL_DIFAUC],
                row[COL_ST], row[COL_DIFATIVO], row[COL_STATUS]), tags=(cor_tag,))  # Alterna cores para melhor legibilidade


    def hide_table(self):
        """Oculta o frame da tabela e ajusta a janela."""
        self.table_frame.pack_forget()
        self.update_idletasks()  # Garante que a GUI seja atualizada imediatamente
        self.geometry("")
        self.button_hide_table.config(text="Exibir Tabela", command=self.show_table, state=tk.DISABLED if self.excel_file_path is None else tk.NORMAL)

    def show_table(self):
        """Exibe o frame da tabela e ajusta a janela."""
        if self.button_hide_logs['text'] == "Ocultar Logs":
            self.hide_logs()  # Oculta os logs automaticamente ao exibir a tabela para melhor visibilidade

        self.table_frame.pack(padx=10, pady=10, fill="both", expand=True)
        self.update_idletasks()  # Garante que a GUI seja atualizada imediatamente
        self.geometry("")
        self.button_hide_table.config(text="Ocultar Tabela", command=self.hide_table)

    def hide_logs(self):
        """Oculta o frame de logs e ajusta a janela."""
        self.log_frame.pack_forget()
        self.update_idletasks()  # Garante que a GUI seja atualizada imediatamente
        self.geometry("")
        self.button_hide_logs.config(text="Exibir Logs", command=self.show_logs, state=tk.DISABLED if self.log_text.get("1.0", tk.END).strip() == "" else tk.NORMAL)

    def show_logs(self):
        """Exibe o frame de logs e ajusta a janela."""
        if self.button_hide_table['text'] == "Ocultar Tabela":
            self.hide_table()

        self.log_frame.pack(padx=10, pady=10, fill="both", expand=True)
        self.update_idletasks()  # Garante que a GUI seja atualizada imediatamente
        self.geometry("")
        self.button_hide_logs.config(text="Ocultar Logs", command=self.hide_logs)

    def update_log(self, message, level="INFO"):
        """Adiciona uma mensagem ao Text widget de logs."""
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, datetime.now().strftime('%H:%M:%S') + " - ", "TIME")
        self.log_text.insert(tk.END, message + "\n", level)
        self.log_text.see(tk.END)  # Rola para o final
        self.log_text.configure(state="disabled")
        # Não precisa de update_idletasks() aqui, pois já é chamado via after(0)

    def update_status_label(self, message):
        """Atualiza o texto do label de status."""
        self.status_label.config(text=f"{message}")
        if hasattr(self, 'status_error_label') and self.status_error_label.winfo_exists():
            self.status_error_label.config(text=f"{message}")  # Limpa a mensagem de erro se houver

    def update_progress_bar_max(self, max_value):
        """Configura o valor máximo da barra de progresso."""
        self.progress_bar["maximum"] = max_value
        self.progress_bar["value"] = 0
        self.progress_label.config(text=f"0/{max_value}")
        if hasattr(self, 'label_progress') and self.label_progress.winfo_exists():
            self.label_progress.config(text=f"0/{max_value}")

    def update_progress(self, current_value):
        """Atualiza o valor da barra de progresso."""
        self.progress_bar["value"] = current_value
        self.progress_label.config(text=f"{current_value}/{self.progress_bar['maximum']}")
        if hasattr(self, 'label_progress') and self.label_progress.winfo_exists():
            self.label_progress.config(text=f"{current_value}/{self.progress_bar['maximum']}")

    def format_time(self, seconds):
        hours, remainder = divmod(int(seconds), 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02}:{minutes:02}:{seconds:02}"

    def update_timer_label(self):
        """Atualiza o label do timer."""
        if self.automation.time_took == 0:
            self.timer_label.config(text="Tempo decorrido: 00:00:00")
        else:
            self.timer_label.config(text=f"Tempo decorrido: {self.format_time(self.automation.time_took)}")

    def set_automation_running(self, is_running):
        """Ajusta o estado dos botões e da GUI com base no status da automação."""
        self.running = is_running
        if is_running:
            self.btn_start.config(state=tk.DISABLED)
            self.btn_load_excel.config(state=tk.DISABLED)
            self.btn_stop.config(state=tk.NORMAL)
            self.spinbox_mes.config(state=tk.DISABLED)
            self.spinbox_ano.config(state=tk.DISABLED)
            self.open_tools_toplevel_window.config(state=tk.NORMAL, command=self.open_tools_toplevel)  # Permite abrir a janela de ferramentas durante a execução
        else:
            self.btn_start.config(state=tk.NORMAL if self.excel_file_path else tk.DISABLED)
            self.btn_load_excel.config(state=tk.NORMAL)
            self.btn_stop.config(state=tk.DISABLED)
            self.spinbox_mes.config(state=tk.DISABLED)
            self.spinbox_ano.config(state=tk.DISABLED)
            self.open_tools_toplevel_window.config(state=tk.DISABLED)
            if hasattr(self, 'tools_window') and self.tools_window.winfo_exists():
                self.tool_button_retry.config(state=tk.DISABLED)
                self.tool_button_skip.config(state=tk.DISABLED)

    def load_excel_file(self):
        """Abre uma janela de diálogo para selecionar o arquivo Excel."""
        file_path = filedialog.askopenfilename(
            title="Selecione a Planilha Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_file_path = file_path
            self.update_status_label(f"Carregando: {Path(file_path).name}...")
            if self.automation.load_excel(file_path):
                self.btn_start.config(state=tk.NORMAL)
                self.button_hide_table.config(state=tk.NORMAL)  # Permite que o usuário veja a tabela de dados

                self.update_log(f"Planilha {Path(file_path).name} carregada com sucesso.", "INFO")
                self.load_table_data(self.automation.df)  # Atualiza a tabela com os dados carregados
                self.after(0, self.show_table)  # Exibe a tabela automaticamente após o carregamento
                self.update_progress_bar_max(len(self.automation.df))
            else:
                self.btn_start.config(state=tk.DISABLED)
                self.excel_file_path = None  # Limpa o caminho se o carregamento falhou
                self.button_hide_table.config(state=tk.DISABLED)  # Desabilita o botão de tabela
                self.spinbox_mes.config(state=tk.DISABLED)
                self.spinbox_ano.config(state=tk.DISABLED)

        self.button_hide_logs.config(state=tk.NORMAL)  # Permite que o usuário veja os logs de erro

    def centraliza_widget(self, widget):
        def center():
            self._centralize_after_id = None
            if not widget.winfo_exists():
                return

            widget.update_idletasks()

            width = widget.winfo_width()
            height = widget.winfo_height()
            x = (widget.winfo_screenwidth() - width) // 2
            y = (widget.winfo_screenheight() - height) // 2

            widget.geometry(f"{width}x{height}+{x}+{y}")

        if self._centralize_after_id is not None:
            self.after_cancel(self._centralize_after_id)
        self._centralize_after_id = self.after(0, center)

    def show_tools_error(self, step_name):
        """Exibe a janela de ferramentas com opções de retry, skip ou stop."""
        self.hit_error = True
        if not hasattr(self, 'tools_window') or not self.tools_window.winfo_exists():
            self.after(0, self.open_tools_toplevel)

        self.tool_button_retry.config(state=tk.NORMAL)
        self.tool_button_skip.config(state=tk.NORMAL)

        if hasattr(self, 'tools_window') and self.tools_window.winfo_exists():
            self.centraliza_widget(self.tools_window)

        self.after(0, lambda: self.status_error_label.config(text=f"Erro na etapa: {step_name}"))

        command = self.tools_response_queue.get()  # Bloqueia até que o usuário escolha uma opção
        match command:
            case "retry":
                self.update_log(f"Usuário escolheu: Tentar Novamente na etapa '{step_name}'", "INFO")
                return "Tentar Novamente"
            case "skip":
                self.update_log(f"Usuário escolheu: Pular Etapa '{step_name}'", "INFO")
                return "Pular"
            case "stop":
                self.update_log(f"Usuário escolheu: Parar Automação na etapa '{step_name}'", "INFO")
                return "Parar"

    def open_tools_toplevel(self):
        if hasattr(self, 'tools_window') and self.tools_window.winfo_exists():
            self.tools_window.lift()  # Traz a janela existente para frente
            if self.hit_error:
                self.centraliza_widget(self.tools_window)
            return
        
        self.tools_window = tk.Toplevel(self)
        self.tools_window.attributes("-topmost", True)  # Mantém a janela de ferramentas sempre no topo
        if self.hit_error:
            self.centraliza_widget(self.tools_window)
        else:
            self.tools_window.geometry("-150+0")
        self.tools_window.config(borderwidth=3, relief="raised")
        self.tools_window.overrideredirect(True)
        self.tools_window.resizable(False, False)

        # STATUS SUPERIOR
        status_frame = ttk.Frame(self.tools_window)
        status_frame.pack(fill="x", expand=True)
        self.status_error_label = ttk.Label(status_frame, text=self.status_label["text"], foreground="red")
        self.label_progress = ttk.Label(status_frame, text=self.progress_label['text'], foreground="blue")

        self.status_error_label.pack(side="left")
        self.label_progress.pack(side="right")

        # BOTOES DE CONTROLE
        self.frame = ttk.Frame(self.tools_window)
        self.frame.pack(fill="x", expand=True)
        self.tool_button_retry = ttk.Button(self.frame, text="Tentar Novamente", command=lambda: self.send_tool_command("retry"), state=tk.DISABLED if not self.hit_error else tk.NORMAL)
        self.tool_button_skip = ttk.Button(self.frame, text="Pular Etapa", command=lambda: self.send_tool_command("skip"), state=tk.DISABLED if not self.hit_error else tk.NORMAL)
        self.tool_button_stop = ttk.Button(self.frame, text="Parar Automação", command=lambda: self.send_tool_command("stop"))
        self.tool_button_close = ttk.Button(self.frame, text="Fechar", command=self.fechar_tools_toplevel)
        self.tool_button_retry.pack(pady=5, side="left")
        self.tool_button_skip.pack(pady=5, side="left")
        self.tool_button_stop.pack(pady=5, side="left")
        self.tool_button_close.pack(pady=5)

    def send_tool_command(self, command):
        if command == "stop":
            self.automation.stop_automation()
        self.tools_response_queue.put(command)

    def fechar_tools_toplevel(self):
        """Fecha a janela Toplevel de ferramentas, se estiver aberta."""
        if self._centralize_after_id is not None:
            self.after_cancel(self._centralize_after_id)
            self._centralize_after_id = None
        if hasattr(self, 'tools_window') and self.tools_window.winfo_exists():
            self.tools_window.destroy()
            del self.tools_window
        self.deiconify()  # Traz a janela principal de volta à frente, DEIXAR OPCIONAL

    def start_automation_threaded(self):
        """Inicia a automação em uma nova thread para não bloquear a GUI."""
        if not self.excel_file_path:
            messagebox.showwarning("Início da Automação", "Por favor, carregue uma planilha Excel primeiro.")
            return

        # Cria e inicia uma nova thread para a automação
        self.iconify() # DEIXAR OPCIONAL
        automation_thread = threading.Thread(target=self.automation.start_automation)
        automation_thread.daemon = True  # Permite que o programa principal saia mesmo se a thread estiver rodando
        automation_thread.start()
        self.set_automation_running(True)
        self.update_log("Thread de automação iniciada.", "INFO")

    def stop_automation(self):
        """Solicita à thread de automação que pare."""
        self.automation.stop_automation()
        # O estado da GUI será atualizado quando a thread de automação reconhecer a parada

    def update_status_column(self, row_index, status):
        children = self.table.get_children()

        if 0 <= row_index < len(children):
            item_id = children[row_index]
            self.table.set(item_id, COL_STATUS, status)
            if status == "error":
                self.table.item(item_id, tags=('error',))
            elif status == "skipped":
                self.table.item(item_id, tags=('skipped',))
            elif status == "ok":
                self.table.item(item_id, tags=('par' if row_index % 2 == 0 else 'impar',))
        else:
            self.update_log(f"Índice de linha inválido para atualização de STATUS: {row_index}", "ERROR")

def main():
    # --- Ponto de Entrada Principal ---
    app = AutomationGUI()
    app.mainloop()

if __name__ == "__main__":
    main()