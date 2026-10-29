import pyautogui
import time
import cv2  # Necessário pelo pyautogui para reconhecimento de imagem
import numpy as np  # Necessário pelo cv2
from pyautogui import ImageNotFoundException  # Importa a exceção específica

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd

import logging
import os
import sys
from datetime import datetime
from pathlib import Path
import threading  # Para executar a automação em uma thread separada e manter a GUI responsiva

# --- Constantes de Configuração ---
IMAGE_DIR = "imagens"  # Pasta onde as imagens estão localizadas
LOG_FILE = "automation_log.txt"
DEFAULT_CONFIDENCE = 0.8  # Nível de confiança para encontrar imagens (0.0 a 1.0)
DEFAULT_IMAGE_TIMEOUT = 30  # Tempo limite em segundos para encontrar uma imagem
DEFAULT_STEP_DELAY = 2  # Atraso em segundos entre as etapas principais
DEFAULT_CLICK_DELAY = 2  # Atraso em segundos após cada clique
TYPE_DELAY = 0.05  # Atraso em segundos entre a digitação de cada caractere
MAX_ATTEMPTS_PER_STEP = 3  # Número máximo de tentativas para cada sub-ação dentro de uma etapa

# Nomes das Colunas Obrigatórias na Planilha Excel
COL_EMPRESA = "EMPRESA"
COL_COMPETENCIA = "COMPETENCIA"
COL_COMENC = "COMENC"
COL_SEMENC = "SEMENC"
COL_DIFAUC = "DIFAUC"
COL_ST = "ST"
COL_STATUS = "STATUS"
COL_DIFATIVO = "DIFATIVO"  # Adicionada a coluna DIFATIVO


# --- Configuração do Logger ---
def setup_logger():
    log_path = Path(LOG_FILE)
    if log_path.exists():
        try:
            os.remove(log_path)  # Limpa o log anterior ao iniciar
        except OSError as e:
            print(f"Não foi possível remover o arquivo de log anterior {log_path}: {e}")
            pass  # Continua mesmo que não consiga remover o log

    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)

    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%d/%m/%Y %H:%M:%S')

    file_handler = logging.FileHandler(LOG_FILE, encoding='utf-8')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger


logger = setup_logger()


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
        self.image_dir = Path(IMAGE_DIR)
        self.df = None  # DataFrame para armazenar os dados do Excel
        self.excel_path = None
        self.current_row_index = -1
        self.stop_requested = False  # Flag para solicitar parada da automação

        # Configurações globais do pyautogui
        pyautogui.PAUSE = TYPE_DELAY  # Atraso padrão entre as digitações
        pyautogui.FAILSAFE = True  # Move o mouse para o canto superior esquerdo para parar a automação

        if not self.image_dir.is_dir():
            logger.error(f"Diretório de imagens '{self.image_dir}' não encontrado. Por favor, crie-o.")
            if self.gui:
                messagebox.showerror("Erro de Inicialização",
                                     f"Diretório de imagens '{self.image_dir}' não encontrado. A automação não pode continuar.")
            sys.exit(1)  # Sai do programa se o diretório de imagens não existir

    def log_message(self, level, message):
        """Envia mensagens para o logger e para a interface gráfica."""
        if self.gui:
            # Garante que a atualização da GUI ocorra na thread principal
            self.gui.after(0, self.gui.update_log, message)
        if level == "INFO":
            logger.info(message)
        elif level == "WARNING":
            logger.warning(message)
        elif level == "ERROR":
            logger.error(message)

    def _check_image_exists_internal(self, image_name):
        """
        Verifica se um arquivo de imagem existe no diretório de imagens.
        Levanta FileNotFoundError se o arquivo não for encontrado.
        """
        image_path = self.image_dir / image_name
        if not image_path.is_file():
            raise FileNotFoundError(
                f"Arquivo de imagem não encontrado: '{image_path}'. Por favor, verifique se a imagem está na pasta '{IMAGE_DIR}'.")
        return True  # Retorna True se o arquivo existe

    def _find_image_internal(self, image_name, timeout=DEFAULT_IMAGE_TIMEOUT, confidence=DEFAULT_CONFIDENCE):
        """
        Tenta encontrar uma imagem na tela.
        Retorna a localização da imagem se encontrada, caso contrário, levanta uma exceção.
        """
        self._check_image_exists_internal(image_name)  # Isso levantará FileNotFoundError se a imagem não existir

        image_path = str(self.image_dir / image_name)
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                location = pyautogui.locateOnScreen(image_path, confidence=confidence)
                if location:
                    return location
            except ImageNotFoundException:  # Exceção específica do pyautogui quando a imagem não é encontrada
                pass  # Continua procurando
            except Exception as e:  # Captura outros erros inesperados durante a busca
                raise RuntimeError(f"Erro inesperado ao procurar imagem '{image_name}': {e}")
            time.sleep(0.5)  # Verifica a cada 0.5 segundo

        # Se o loop terminar sem encontrar a imagem
        raise ImageNotFoundException(f"Imagem '{image_name}' não encontrada na tela após {timeout} segundos.")

    def _click_image_internal(self, image_name, timeout=DEFAULT_IMAGE_TIMEOUT, confidence=DEFAULT_CONFIDENCE,
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
            time.sleep(DEFAULT_CLICK_DELAY)
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
        for attempt in range(1, MAX_ATTEMPTS_PER_STEP + 1):
            if self.stop_requested:
                raise InterruptedAutomationException("Automação interrompida pelo usuário.")

            try:
                self.log_message("INFO",
                                 f"Tentativa {attempt}/{MAX_ATTEMPTS_PER_STEP} para {action_desc} (Etapa: {step_name})")
                result = action_func(*args, **kwargs)
                self.log_message("INFO", f"Sucesso em {action_desc}.")
                return result  # Se bem-sucedido, retorna o resultado
            except (FileNotFoundError, ImageNotFoundException, RuntimeError) as e:
                self.log_message("ERROR",
                                 f"Falha em {action_desc} (Etapa: {step_name}, Tentativa {attempt}/{MAX_ATTEMPTS_PER_STEP}): {e}")

                # Exibe o diálogo imediatamente na primeira falha para dar controle ao usuário
                action = self._handle_step_error(step_name)

                if action == "Parar":
                    raise InterruptedAutomationException("Automação interrompida pelo usuário após erro.")
                elif action == "Pular":
                    raise SkipStepException("Usuário optou por pular esta etapa.")
                elif action == "Tentar Novamente":
                    # Usuário escolheu tentar novamente, então o loop continua
                    # Adiciona um pequeno atraso para o usuário ter tempo de reagir/corrigir (ex: minimizar janela)
                    time.sleep(1)
                    continue  # Continua para a próxima tentativa no loop

        # Se o loop terminar sem retornar (todas as tentativas falharam e usuário optou por tentar novamente)
        raise RuntimeError(
            f"Ação '{action_desc}' falhou após múltiplas tentativas e intervenção do usuário na etapa '{step_name}'.")

    # Funções públicas que as etapas da automação usarão (com retry e tratamento de erro)
    def click_image(self, image_name, **kwargs):
        """Encontra e clica em uma imagem, com retentativas e diálogo de erro."""
        step_name = kwargs.pop('step_context', 'Ação de Clique')
        return self._perform_action_with_retry(
            self._click_image_internal,
            f"clicar em '{image_name}'",
            step_name,
            image_name,
            **kwargs
        )

    def find_image(self, image_name, **kwargs):
        """Encontra uma imagem, com retentativas e diálogo de erro. Retorna a localização."""
        step_name = kwargs.pop('step_context', 'Ação de Busca de Imagem')
        return self._perform_action_with_retry(
            self._find_image_internal,
            f"encontrar '{image_name}'",
            step_name,
            image_name,
            **kwargs
        )

    def type_text(self, text, **kwargs):
        """Digita um texto, com retentativas e diálogo de erro."""
        step_name = kwargs.pop('step_context', 'Ação de Digitação')
        return self._perform_action_with_retry(
            self._type_text_internal,
            f"digitar '{text}'",
            step_name,
            text,
            **kwargs
        )

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
        response = self.gui.show_error_dialog(step_name)
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
            except Exception as e:
                self.log_message("ERROR",
                                 f"Erro ao salvar a planilha '{self.excel_path}': {e}. Por favor, feche o arquivo se estiver aberto.")
                if self.gui:
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
            # Valida as colunas obrigatórias - AGORA INCLUINDO COL_DIFATIVO
            required_columns = [COL_EMPRESA, COL_COMPETENCIA, COL_COMENC, COL_SEMENC, COL_DIFAUC, COL_ST, COL_STATUS,
                                COL_DIFATIVO]
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
            self.df[COL_STATUS] = self.df[COL_STATUS].astype(str).fillna("")

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
        competencia = str(row_data[COL_COMPETENCIA]).strip()  # Formato mm/aaaa

        self._update_excel_status(row_index, "processing")
        self.log_message("INFO", f"Iniciando automação para Empresa: {empresa}, Competência: {competencia}")
        self.gui.after(0, self.gui.update_progress, row_index + 1)
        self.gui.after(0, self.gui.update_status_label, f"Processando: {empresa} ({competencia})")

        is_sem_movimento = self._is_sem_movimento(row_data)
        movimento_type = "sem movimento" if is_sem_movimento else "com movimento"
        self.log_message("INFO", f"Tipo de movimento detectado para {empresa}: {movimento_type}")

        try:
            # ETAPA 1: Abrir movimento da empresa
            step_name = "ETAPA 1: Abrir movimento da empresa"
            self.click_image("novo_documento.png", step_context=step_name)
            time.sleep(4)  # Espera especificada pelo usuário
            self.click_image("seta_baixo.png", step_context=step_name)
            self.type_text(empresa, step_context=step_name)
            pyautogui.press('enter')
            time.sleep(DEFAULT_STEP_DELAY)  # Espera para o sistema reagir
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
            time.sleep(DEFAULT_CLICK_DELAY)

            if is_sem_movimento:
                self.type_text("1", step_context=step_name)  # Sem movimento
            else:
                self.type_text("0", step_context=step_name)  # Com movimento
            pyautogui.press('tab')

            # ETAPA 2: Condicional (Sem Movimento ou Com Movimento)
            step_name = f"ETAPA 2: {'Sem' if is_sem_movimento else 'Com'} Movimento"
            self.click_image("confirmar.png", step_context=step_name)  # (Etapa 2 inicial)
            time.sleep(6)  # Espera especificada
            self.click_image("abrir.png", step_context=step_name)

            if not is_sem_movimento:  # Lógica APENAS para "Com Movimento"
                self.click_image("destda.png", step_context=step_name)
                self.click_image("alterar.png", step_context=step_name)

                # Garante que os valores numéricos sejam formatados com vírgula para digitação
                st_val = self.format_number_for_typing(row_data[COL_ST])
                comenc_val = self.format_number_for_typing(row_data[COL_COMENC])
                semenc_val = self.format_number_for_typing(row_data[COL_SEMENC])
                difauc_val = self.format_number_for_typing(row_data[COL_DIFAUC])
                difativo_val = self.format_number_for_typing(
                    row_data[COL_DIFATIVO])  # <-- NOVO: Lê DIFATIVO da planilha

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

            # ETAPA 3: Gerar Documento
            step_name = "ETAPA 3: Gerar Documento"
            self.click_image("encerrar.png", step_context=step_name)
            self.click_image("gerar_documento.png", step_context=step_name)
            self.click_image("iniciar_processamento.png", step_context=step_name)  # (Gerar Documento)
            self.click_image("selecionar_certificado.png", step_context=step_name)
            self.click_image("autenticar.png", step_context=step_name)

            self.log_message("INFO", "Aguardando 'processo_finalizado.png'...")
            self.find_image("processo_finalizado.png", timeout=30, step_context=step_name)  # Aguarda até 30s
            self.click_image("ok.png", step_context=step_name)  # após processo finalizado

            # ETAPA 4: Transmissão
            step_name = "ETAPA 4: Transmissão"
            self.click_image("transmitir.png", step_context=step_name)
            self.click_image("iniciar_processamento.png", step_context=step_name)  # (Transmissão)

            # NOVO: Aguardar finalização da transmissão e clicar OK
            self.log_message("INFO", "Aguardando 'processo_finalizado.png' após transmissão...")
            self.find_image("processo_finalizado.png", timeout=30, step_context=step_name)  # Aguarda até 30s
            self.click_image("ok.png", step_context=step_name)  # Clica OK

            # ETAPA 5: Salvar Recibo
            step_name = "ETAPA 5: Salvar Recibo"
            self.log_message("INFO", "Iniciando ETAPA 5: Salvar Recibo.")

            time.sleep(3)  # Esperar por 3 segundos
            self.click_image("exportar.png", step_context=step_name)

            # Definir caminho de salvamento
            competencia_formatted = competencia.replace("/", "-")
            save_base_path = r"C:\Users\mb.contabilidade\Desktop\Recibos Destda"
            save_folder_path = self._create_folder_structure(save_base_path, competencia_formatted)

            if not save_folder_path:
                raise RuntimeError(
                    f"Não foi possível criar/acessar a pasta de destino para salvar o recibo: {save_base_path}{competencia_formatted}")

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
            time.sleep(DEFAULT_STEP_DELAY)  # Atraso entre as linhas da planilha

        except InterruptedAutomationException:
            # Se o usuário escolheu parar a automação
            self._update_excel_status(row_index, "error")  # Marca a linha atual como erro
            self.stop_requested = True  # Define a flag de parada
            return  # Sai da execução para esta linha
        except SkipStepException:
            # Se o usuário escolheu pular esta etapa
            self._update_excel_status(row_index, "skipped")  # Marca a linha como pulada
            return  # Sai da execução para esta linha e continua para a próxima
        except Exception as e:
            # Captura qualquer outra exceção não tratada pelas chamadas a _perform_action_with_retry
            # Isso indicaria um erro de lógica no script ou um cenário não previsto.
            self.log_message("ERROR",
                             f"Erro fatal ou não-tratado durante a automação para {empresa} - {competencia}: {e}")
            self._update_excel_status(row_index, "error")
            return  # Sai da execução para esta linha

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
        self.gui.after(0, self.gui.update_progress_bar_max, len(self.df))
        self.log_message("INFO", "Iniciando processo de automação...")

        try:
            for index, row in self.df.iterrows():
                if self.stop_requested:
                    self.log_message("INFO", "Automação interrompida pelo usuário.")
                    break

                current_status = row[COL_STATUS]
                # Ignora linhas já processadas ou puladas
                if isinstance(current_status, str) and current_status.lower() in ["ok", "skipped"]:
                    self.log_message("INFO",
                                     f"Linha {index + 2} (Empresa: {row[COL_EMPRESA]}) já está com status '{current_status}'. Pulando para a próxima.")
                    self.gui.after(0, self.gui.update_progress, index + 1)
                    continue

                self.run_automation_for_row(index, row)

            self.log_message("INFO", "Processo de automação finalizado.")
        except Exception as e:
            self.log_message("ERROR", f"Erro inesperado durante a execução principal da automação: {e}")
            if self.gui:
                self.gui.after(0, messagebox.showerror, "Erro Fatal na Automação",
                               f"Um erro inesperado ocorreu durante a automação:\n{e}\nPor favor, verifique os logs.")
            # Marca a linha atual como erro se a automação parou inesperadamente no meio
            if self.current_row_index != -1 and not self.stop_requested:
                self._update_excel_status(self.current_row_index, "error")
        finally:
            self.gui.after(0, self.gui.set_automation_running, False)
            self.gui.after(0, self.gui.update_progress, len(self.df))  # Garante que a barra de progresso atinja o fim
            self.gui.after(0, self.gui.update_status_label, "Automação finalizada.")

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
        self.geometry("800x600")
        self.automation = WebAutomation(self)  # Instancia a lógica de automação, passando a si mesmo (GUI)
        self.excel_file_path = None
        self.running = False  # Estado da automação

        self._create_widgets()
        self._setup_layout()

        # Garante que o diretório de imagens exista ao iniciar a GUI
        if not Path(IMAGE_DIR).is_dir():
            Path(IMAGE_DIR).mkdir(exist_ok=True)
            self.update_log(f"Diretório '{IMAGE_DIR}' criado (se não existia).")

    def _create_widgets(self):
        """Cria todos os componentes da interface gráfica."""
        # Frame para os controles
        self.control_frame = ttk.LabelFrame(self, text="Controles")
        self.control_frame.pack(padx=10, pady=10, fill="x")

        self.btn_load_excel = ttk.Button(self.control_frame, text="Carregar Planilha Excel",
                                         command=self.load_excel_file)
        self.btn_start = ttk.Button(self.control_frame, text="Iniciar Automação",
                                    command=self.start_automation_threaded, state=tk.DISABLED)
        self.btn_stop = ttk.Button(self.control_frame, text="Parar Automação", command=self.stop_automation,
                                   state=tk.DISABLED)

        self.status_label = ttk.Label(self.control_frame, text="Status: Aguardando planilha...")
        self.progress_bar = ttk.Progressbar(self.control_frame, orient="horizontal", length=300, mode="determinate")
        self.progress_label = ttk.Label(self.control_frame, text="0/0")

        # Frame para os logs
        self.log_frame = ttk.LabelFrame(self, text="Logs de Operação")
        self.log_frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.log_text = tk.Text(self.log_frame, state="disabled", wrap="word", height=20, width=80)
        self.log_text_scroll = ttk.Scrollbar(self.log_frame, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=self.log_text_scroll.set)

    def _setup_layout(self):
        """Define o layout dos componentes na grade."""
        self.btn_load_excel.grid(row=0, column=0, padx=5, pady=5)
        self.btn_start.grid(row=0, column=1, padx=5, pady=5)
        self.btn_stop.grid(row=0, column=2, padx=5, pady=5)
        self.status_label.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        self.progress_bar.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        self.progress_label.grid(row=2, column=2, padx=5, pady=5)

        self.log_text.pack(side="left", fill="both", expand=True)
        self.log_text_scroll.pack(side="right", fill="y")

    def update_log(self, message):
        """Adiciona uma mensagem ao Text widget de logs."""
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)  # Rola para o final
        self.log_text.configure(state="disabled")
        # Não precisa de update_idletasks() aqui, pois já é chamado via after(0)

    def update_status_label(self, message):
        """Atualiza o texto do label de status."""
        self.status_label.config(text=f"Status: {message}")

    def update_progress_bar_max(self, max_value):
        """Configura o valor máximo da barra de progresso."""
        self.progress_bar["maximum"] = max_value
        self.progress_bar["value"] = 0
        self.progress_label.config(text=f"0/{max_value}")

    def update_progress(self, current_value):
        """Atualiza o valor da barra de progresso."""
        self.progress_bar["value"] = current_value
        self.progress_label.config(text=f"{current_value}/{self.progress_bar['maximum']}")
        # Não precisa de update_idletasks() aqui, pois já é chamado via after(0)

    def set_automation_running(self, is_running):
        """Ajusta o estado dos botões e da GUI com base no status da automação."""
        self.running = is_running
        if is_running:
            self.btn_start.config(state=tk.DISABLED)
            self.btn_load_excel.config(state=tk.DISABLED)
            self.btn_stop.config(state=tk.NORMAL)
        else:
            self.btn_start.config(state=tk.NORMAL if self.excel_file_path else tk.DISABLED)
            self.btn_load_excel.config(state=tk.NORMAL)
            self.btn_stop.config(state=tk.DISABLED)

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
                self.update_log(f"Planilha {Path(file_path).name} carregada com sucesso.")
            else:
                self.btn_start.config(state=tk.DISABLED)
                self.excel_file_path = None  # Limpa o caminho se o carregamento falhou

    def start_automation_threaded(self):
        """Inicia a automação em uma nova thread para não bloquear a GUI."""
        if not self.excel_file_path:
            messagebox.showwarning("Início da Automação", "Por favor, carregue uma planilha Excel primeiro.")
            return

        # Cria e inicia uma nova thread para a automação
        automation_thread = threading.Thread(target=self.automation.start_automation)
        automation_thread.daemon = True  # Permite que o programa principal saia mesmo se a thread estiver rodando
        automation_thread.start()
        self.set_automation_running(True)
        self.update_log("Thread de automação iniciada.")

    def stop_automation(self):
        """Solicita à thread de automação que pare."""
        self.automation.stop_automation()
        # O estado da GUI será atualizado quando a thread de automação reconhecer a parada

    def show_error_dialog(self, step_name):
        """
        Exibe um diálogo de erro interativo.
        Esta função é chamada pela thread de automação, mas a messagebox é executada na thread principal.
        """
        # Cria um contêiner para a resposta e um evento para esperar
        response_container = []
        event = threading.Event()

        def _show_dialog_on_main_thread():
            response = messagebox.askyesnocancel(
                "Erro na Automação",
                f"Ocorreu um erro na etapa: {step_name}.\n\nO que deseja fazer?",
                icon="warning",
                type="yesnocancel",  # Yes=Tentar Novamente, No=Pular Etapa, Cancel=Parar Execução
                detail="Sim: Tentar Novamente\nNão: Pular Etapa\nCancelar: Parar Execução"
            )
            response_container.append(response)
            event.set()  # Sinaliza que a resposta foi recebida

        # Agenda a execução do diálogo na thread principal da GUI
        self.after(0, _show_dialog_on_main_thread)

        # Espera que o usuário interaja com o diálogo
        event.wait()

        # Processa a resposta
        response = response_container[0]
        if response is True:  # Sim (Tentar Novamente)
            self.update_log(f"Usuário escolheu: Tentar Novamente na etapa '{step_name}'")
            return "Tentar Novamente"
        elif response is False:  # Não (Pular Etapa)
            self.update_log(f"Usuário escolheu: Pular Etapa '{step_name}'")
            return "Pular"
        else:  # Cancelar (Parar Execução)
            self.update_log(f"Usuário escolheu: Parar Execução na etapa '{step_name}'")
            return "Parar"


# --- Ponto de Entrada Principal ---
if __name__ == "__main__":
    app = AutomationGUI()
    app.mainloop()



