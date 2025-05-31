import os
import re
import time
import datetime
import logging
from dataclasses import dataclass
from typing import List, Optional, Tuple
import pandas as pd
# from PIL import Image
# import requests
# from io import BytesIO
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys


from dotenv import load_dotenv
load_dotenv()

# Constants for selectors and URLs
URL_SEI_SP = "https://sei.sp.gov.br/sei/modulos/pesquisa/md_pesq_processo_pesquisar.php?acao_externa=protocolo_pesquisar&acao_origem_externa=protocolo_pesquisar&id_orgao_acesso_externo=0"
DRIVER_PATH = os.getenv("DRIVER_PATH", "")#insira o caminho do seu msedgedriver.exe aqui

CHECKBOXES_TO_MARK = [
    "chkSinProcessos",
    "chkSinDocumentosGerados",
    "chkSinDocumentosRecebidos"
]

TIPO_PROCESSO = "Processo de apropriação e utilização de crédito acumulado ou de produtor rural"
TIPO_DOCUMENTO = "Despacho"

SELECTORS = {
    "tipo_processo": "selTipoProcedimentoPesquisa",
    "tipo_documento": "selSeriePesquisa",
    "data_inicio": "txtDataInicio",
    "data_fim": "txtDataFim",
    "captcha_img": "imgCaptcha",
    "captcha_input": "txtInfraCaptcha",
    "botao_pesquisar": "sbmPesquisar"
}

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


@dataclass
class EdgeDriverManager:
    """Gerenciador do WebDriver Edge."""

    headless: bool = False
    timeout: int = 10
    driver_path: Optional[str] = None
    driver: Optional[webdriver.Edge] = None
    wait: Optional[WebDriverWait] = None
    logger: logging.Logger = logging.getLogger(__name__)

    def setup_driver(self) -> Optional[webdriver.Edge]:
        """Configura e inicializa o WebDriver Edge."""
        edge_options = Options()
        if self.headless:
            edge_options.add_argument("--headless")

        edge_options.add_argument("--no-sandbox")
        edge_options.add_argument("--disable-gpu")
        edge_options.add_argument("--disable-dev-shm-usage")
        edge_options.add_argument("--disable-extensions")
        edge_options.add_argument("--disable-blink-features=AutomationControlled")

        try:
            if self.driver_path and os.path.exists(self.driver_path):
                service = Service(self.driver_path)
                self.driver = webdriver.Edge(service=service, options=edge_options)
                self.logger.info(f"EdgeDriver Inicializado com sucesso: {self.driver_path}")
            else:
                self.driver = webdriver.Edge(options=edge_options)
                self.logger.info("EdgeDriver Inicializado usando driver do sistema")

            self.wait = WebDriverWait(self.driver, self.timeout)
            return self.driver

        except Exception as e:
            self.logger.error(f"Erro ao inicializar EdgeDriver: {e}")
            return None

    def close(self) -> None:
        """Fecha o WebDriver."""
        if self.driver:
            self.driver.quit()
            self.logger.info("WebDriver Fechado")


class CaptchaResolver:
    """Classe para resolver diferentes tipos de Captcha."""

    def __init__(self, driver: webdriver.Edge):
        self.driver = driver
        self.logger = logging.getLogger(__name__)

    def resolve_manual(self, captcha_element) -> str:
        """Resolve CAPTCHA com input manual do usuário."""
        print("\n" + "=".center(50, "="))
        print("RESOLUÇÃO MANUAL DE CAPTCHA")
        print("=" * 50)
        print("Observe a imagem do Captcha no navegador e digite o texto:")

        captcha_text = input("Digite o CAPTCHA: ").strip()
        return captcha_text


class FormHandler:
    def __init__(self, driver: webdriver.Edge, wait: WebDriverWait):
        self.driver = driver
        self.wait = wait
        self.logger = logging.getLogger(__name__)

    def fill_dropdown(self, selector: str, value: str, by: By = By.ID, selection_type: str = "text") -> bool:
        """Preenche um dropdown com base no tipo de seleção."""
        try:
            dropdown_element = self.wait.until(EC.presence_of_element_located((by, selector)))
            dropdown = Select(dropdown_element)

            if selection_type == "text":
                dropdown.select_by_visible_text(value)
            elif selection_type == "value":
                dropdown.select_by_value(value)
            elif selection_type == "index":
                dropdown.select_by_index(int(value))

            self.logger.info(f"Dropdown preenchido: {selector} = {value}")
            return True
        except Exception as e:
            self.logger.error(f"Erro ao preencher dropdown {selector}: {e}")
            return False

    def fill_input(self, selector: str, value: str, by: By = By.ID) -> bool:
        """Preenche um campo de input."""
        try:
            input_element = self.wait.until(EC.presence_of_element_located((by, selector)))
            input_element.clear()
            input_element.send_keys(value)
            self.logger.info(f"Input preenchido: {selector} = {value}")
            return True
        except Exception as e:
            self.logger.error(f"Erro ao preencher input {selector}: {e}")
            return False

    def marcar_checkboxes(self, checkboxes_ids: List[str]) -> bool:
        """Marca os checkboxes especificados."""
        try:
            for checkbox_id in checkboxes_ids:
                checkbox = self.wait.until(EC.presence_of_element_located((By.ID, checkbox_id)))
                self.driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                if not checkbox.is_selected():
                    try:
                        checkbox.click()
                    except Exception:
                        self.driver.execute_script("arguments[0].click();", checkbox)
                    self.logger.info(f"Checkbox marcado: {checkbox_id}")
            return True
        except Exception as e:
            self.logger.error(f"Erro ao marcar checkboxes: {e}")
            return False

    def click_element(self, selector: str, by: By = By.ID) -> bool:
        """Clica em um elemento especificado."""
        try:
            element = self.wait.until(EC.element_to_be_clickable((by, selector)))
            element.click()
            self.logger.info(f"Elemento clicado: {selector}")
            return True
        except Exception as e:
            self.logger.error(f"Erro ao clicar em {selector}: {e}")
            return False

    def executar_fluxo_pesquisa(
        self,
        tipo_processo: str,
        tipo_documento: str,
        checkboxes_ids: List[str],
        captcha_element_id: str,
        captcha_input_id: str,
        botao_pesquisar_id: str,
    ) -> bool:
        """Executa o fluxo completo de pesquisa no site."""
        try:
            self.marcar_checkboxes(checkboxes_ids)
            self.fill_dropdown(SELECTORS["tipo_processo"], tipo_processo)
            self.fill_dropdown(SELECTORS["tipo_documento"], tipo_documento)

            hoje = datetime.date.today()
            seis_meses_atras = hoje - datetime.timedelta(days=5)
            self.fill_input(SELECTORS["data_inicio"], seis_meses_atras.strftime("%d/%m/%Y"))
            self.fill_input(SELECTORS["data_fim"], hoje.strftime("%d/%m/%Y"))

            captcha_element = self.wait.until(EC.presence_of_element_located((By.ID, captcha_element_id)))
            captcha_text = CaptchaResolver(self.driver).resolve_manual(captcha_element)
            self.fill_input(captcha_input_id, captcha_text)

            self.click_element(botao_pesquisar_id)

            return True
        except Exception as e:
            self.logger.error(f"Erro ao executar fluxo de pesquisa: {e}")
            return False


class ResultadoExtractor:
    
    def __init__(self, driver: webdriver.Edge, wait: WebDriverWait, logger: Optional[logging.Logger] = None):
        self.driver = driver
        self.wait = wait
        self.logger = logger or logging.getLogger(__name__)

    def carregar_todos_os_resultados(self) -> None:
        """Carrega todos os resultados fazendo scroll até o fim da página."""
        self.logger.info("Iniciando carregamento de todos os resultados.")

        time.sleep(3)

        # Seletor(s) possíveis da div de total de resultados
        possible_selectors = [
            "div.total-registros-infinite",
            ".total-registros",
            ".resultado-total",
            "[class*='total']",
            "[class*='resultado']"
        ]

        total_results = None
        total_results_element = None

        # Scroll até encontrar a div com o total de resultados
        max_scrolls_to_find_total = 65
        scroll_pause_time = 1


        for attempt in range(max_scrolls_to_find_total):
            # Faz scroll usando a tecla PAGE_DOWN
            body = self.driver.find_element(By.TAG_NAME, "body")
            ActionChains(self.driver).move_to_element(body).send_keys(Keys.PAGE_DOWN).perform()
            time.sleep(scroll_pause_time)
            self.logger.info(f"Procurando total de resultados (tentativa {attempt + 1})...")

            # Verifica cada seletor possível
            for selector in possible_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for element in elements:
                        text = element.text.strip()
                        if text and re.search(r'\d+', text):
                            total_match = re.search(r'(\d+)', text)
                            if total_match:
                                total_results = int(total_match.group(1))
                                total_results_element = element
                                self.logger.info(f"Total de resultados encontrado: {total_results} (seletor: {selector})")
                                break
                    if total_results:
                        break
                except Exception as e:
                    self.logger.debug(f"Erro ao tentar seletor {selector}: {e}")
            
            if total_results:
                break

        if not total_results:
            self.logger.warning("Não foi possível determinar o total de resultados. Continuando com scroll até não haver mais conteúdo.")

        # Agora continua com scroll até carregar todos os resultados
        max_scroll_attempts = 50
        no_new_content_limit = 3
        no_new_content_count = 0
        last_height = self.driver.execute_script("return document.body.scrollHeight")
        last_links_count = 0


        for attempt in range(max_scroll_attempts):
            # Faz scroll usando a tecla PAGE_DOWN
            body = self.driver.find_element(By.TAG_NAME, "body")
            ActionChains(self.driver).move_to_element(body).send_keys(Keys.PAGE_DOWN).perform()
            time.sleep(scroll_pause_time)

            # Verifica se encontrou os links após o scroll
            current_links = self.driver.find_elements(By.CSS_SELECTOR, "a.protocoloNormal")
            current_links_count = len(current_links)
            self.logger.info(f"Tentativa {attempt + 1}: {current_links_count} links carregados")

            # Se temos o total esperado e já carregamos todos os links
            if total_results and current_links_count >= total_results:
                self.logger.info("Todos os resultados foram carregados.")
                break

            # Se encontrou algum link, pode prosseguir com a extração
            if current_links_count > 0:
                self.logger.info("Links encontrados após scroll, prosseguindo com a extração.")
                break

            new_height = self.driver.execute_script("return document.body.scrollHeight")

            if new_height == last_height:
                no_new_content_count += 1
                if no_new_content_count >= no_new_content_limit:
                    self.logger.warning("Fim da página alcançado sem carregar mais conteúdo.")
                    break
            else:
                no_new_content_count = 0

            last_height = new_height
            last_links_count = current_links_count

    
    def extrair_links(self) -> List[str]:
        """Extrai os hrefs de todos os links clicáveis dos resultados."""
        self.logger.info("Extraindo links dos resultados.")
        try:
            self.wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.protocoloNormal")))
            links = self.driver.find_elements(By.CSS_SELECTOR, "a.protocoloNormal")
            links_filtrados = [link.get_attribute("href") for link in links if link.get_attribute("href")]
            self.logger.info(f"{len(links_filtrados)} links extraídos.")
            return links_filtrados
        except Exception as e:
            self.logger.error(f"Erro ao extrair links: {e}")
            return []

    def extrair_dados_cliente(self, link: str) -> Tuple[str, str]:
        """Abre o link em nova aba, extrai nome e CNPJ de uma <td>, e fecha a aba."""
        self.logger.info(f"Abrindo link: {link}")
        try:
            self.driver.execute_script("window.open(arguments[0]);", link)
            self.driver.switch_to.window(self.driver.window_handles[-1])

            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(2)  # Aguarda carregamento completo

            nome = ""
            cnpj = ""

            tds = self.driver.find_elements(By.TAG_NAME, "td")
            for td in tds:
                texto = td.text
                match = re.search(r"(.+?)\s*\((\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\)", texto)
                if match:
                    nome = match.group(1).strip()
                    cnpj = match.group(2).strip()
                    break
            else:
                try:
                    trs_infra = self.driver.find_elements(By.CSS_SELECTOR, "tr.infraTrClara")
                    if len(trs_infra) >= 4:
                        tr_infra = trs_infra[3]
                        tds_infra = tr_infra.find_elements(By.TAG_NAME, "td")
                        if len(tds_infra) > 1:
                            nome = tds_infra[1].text.strip()
                            cnpj = ""
                    else:
                        self.logger.warning("Menos de 4 elementos tr.infraTrClara encontrados.")
                        #no futuro pensar em trazer apenas o CNPJ do cliente, e não o nome do cliente
                except Exception as e:
                    self.logger.warning(f"Não encontrou tr.infraTrClara ou erro ao extrair texto: {e}")

            return nome, cnpj
        except Exception as e:
            self.logger.error(f"Erro ao extrair dados do link {link}: {e}")
            return "", ""
        finally:
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])

def main() -> None:
    manager = EdgeDriverManager(headless=False, driver_path=DRIVER_PATH)
    driver = manager.setup_driver()

    if not driver:
        logging.error("Falha ao inicializar o WebDriver. Encerrando o programa.")
        return

    wait = WebDriverWait(driver, 15)
    form = FormHandler(driver, wait)

    try:
        driver.get(URL_SEI_SP)
        time.sleep(3)

        sucesso = form.executar_fluxo_pesquisa(
            tipo_processo=TIPO_PROCESSO,
            tipo_documento=TIPO_DOCUMENTO,
            checkboxes_ids=CHECKBOXES_TO_MARK,
            captcha_element_id=SELECTORS["captcha_img"],
            captcha_input_id=SELECTORS["captcha_input"],
            botao_pesquisar_id=SELECTORS["botao_pesquisar"],
        )

        if not sucesso:
            logging.error("Erro ao executar o fluxo de pesquisa. Encerrando o programa.")
            return

        extrator = ResultadoExtractor(driver, wait)

        extrator.carregar_todos_os_resultados()

        links = extrator.extrair_links()

        data = []
        for link in links:
            nome, cnpj = extrator.extrair_dados_cliente(link)
            logging.info(f"Nome: {nome} | CNPJ: {cnpj}")
            data.append({"Nome": nome, "CNPJ": cnpj})

        df = pd.DataFrame(data, columns=["Nome", "CNPJ"])
        logging.info(f"DataFrame criado com {len(df)} registros.")

        # Optionally, save to CSV or Excel here if needed
        df.to_excel("dados_extraidos.xlsx", index=False)

        time.sleep(15)

    except Exception as e:
        logging.error(f"Erro inesperado durante a execução: {e}")

    finally:
        manager.close()


if __name__ == "__main__":
    main()
