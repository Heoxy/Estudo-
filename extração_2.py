import os, re, time, datetime, logging
import pandas as pd
from dataclasses import dataclass
from typing import List, Optional, Tuple
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

# Constantes
URL_SEI_SP = "https://sei.sp.gov.br/sei/modulos/pesquisa/md_pesq_processo_pesquisar.php?acao_externa=protocolo_pesquisar&acao_origem_externa=protocolo_pesquisar&id_orgao_acesso_externo=0"
DRIVER_PATH = os.getenv("DRIVER_PATH", "")
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
CHECKBOXES = ["chkSinProcessos", "chkSinDocumentosGerados", "chkSinDocumentosRecebidos"]

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

@dataclass
class EdgeDriverManager:
    headless: bool = False
    timeout: int = 10
    driver_path: Optional[str] = None

    def setup(self) -> Tuple[Optional[webdriver.Edge], Optional[WebDriverWait]]:
        options = Options()
        if self.headless:
            options.add_argument("--headless")
        for arg in ["--no-sandbox", "--disable-gpu", "--disable-dev-shm-usage", "--disable-extensions", "--disable-blink-features=AutomationControlled", "--headless=new", "--use-gl=swiftshader"]:
            options.add_argument(arg)
        try:
            service = Service(self.driver_path) if self.driver_path and os.path.exists(self.driver_path) else None
            driver = webdriver.Edge(service=service, options=options)
            return driver, WebDriverWait(driver, self.timeout)
        except Exception as e:
            logging.error(f"Erro ao inicializar EdgeDriver: {e}")
            return None, None

class FormHandler:
    def __init__(self, driver, wait):
        self.d = driver; self.w = wait

    def select(self, sel, val):
        try: Select(self.w.until(EC.presence_of_element_located((By.ID, sel)))).select_by_visible_text(val)
        except Exception as e: logging.error(f"Erro ao selecionar {sel}: {e}")

    def input(self, sel, val):
        try:
            el = self.w.until(EC.presence_of_element_located((By.ID, sel)))
            el.clear(); el.send_keys(val)
        except Exception as e: logging.error(f"Erro ao preencher {sel}: {e}")

    def checkboxes(self):
        for cid in CHECKBOXES:
            try:
                cb = self.w.until(EC.presence_of_element_located((By.ID, cid)))
                if not cb.is_selected():
                    try: cb.click()
                    except: self.d.execute_script("arguments[0].click();", cb)
            except: pass

    def click(self, sel):
        try: self.w.until(EC.element_to_be_clickable((By.ID, sel))).click()
        except Exception as e: logging.error(f"Erro ao clicar {sel}: {e}")

    def executar_fluxo(self):
        self.checkboxes()
        self.select(SELECTORS["tipo_processo"], TIPO_PROCESSO)
        self.select(SELECTORS["tipo_documento"], TIPO_DOCUMENTO)
        hoje = datetime.date.today()
        self.input(SELECTORS["data_inicio"], (hoje - datetime.timedelta(days=5)).strftime("%d/%m/%Y"))
        self.input(SELECTORS["data_fim"], hoje.strftime("%d/%m/%Y"))
        captcha = input("\nDigite o CAPTCHA exibido no navegador: ").strip()
        self.input(SELECTORS["captcha_input"], captcha)
        self.click(SELECTORS["botao_pesquisar"])

class ResultadoExtractor:
    def __init__(self, driver, wait):
        self.d = driver; self.w = wait

    def carregar_resultados(self):
        time.sleep(3)
        total = None
        for _ in range(65):
            ActionChains(self.d).send_keys(Keys.PAGE_DOWN).perform()
            time.sleep(1)
            for sel in [".total-registros", ".resultado-total", "[class*='total']", "[class*='resultado']"]:
                for el in self.d.find_elements(By.CSS_SELECTOR, sel):
                    m = re.search(r'(\d+)', el.text)
                    if m: total = int(m.group(1)); break
                if total: break
            if total: break

        for _ in range(50):
            ActionChains(self.d).send_keys(Keys.PAGE_DOWN).perform()
            time.sleep(1)
            links = self.d.find_elements(By.CSS_SELECTOR, "a.protocoloNormal")
            if total and len(links) >= total: break

    def extrair_links(self) -> List[str]:
        try:
            self.w.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.protocoloNormal")))
            return [l.get_attribute("href") for l in self.d.find_elements(By.CSS_SELECTOR, "a.protocoloNormal") if l.get_attribute("href")]
        except Exception as e:
            logging.error(f"Erro ao extrair links: {e}")
            return []

    def extrair_dados(self, link: str) -> Tuple[str, str]:
        try:
            self.d.execute_script("window.open(arguments[0]);", link)
            self.d.switch_to.window(self.d.window_handles[-1])
            self.w.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(2)
            for td in self.d.find_elements(By.TAG_NAME, "td"):
                m = re.search(r"(.+?)\s*\((\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\)", td.text)
                if m: return m.group(1).strip(), m.group(2).strip()
            trs = self.d.find_elements(By.CSS_SELECTOR, "tr.infraTrClara")
            if len(trs) >= 4:
                tds = trs[3].find_elements(By.TAG_NAME, "td")
                if len(tds) > 1:
                    return tds[1].text.strip(), ""
            return "", ""
        except Exception as e:
            logging.error(f"Erro ao extrair dados: {e}")
            return "", ""
        finally:
            self.d.close()
            self.d.switch_to.window(self.d.window_handles[0])

def main():
    manager = EdgeDriverManager(driver_path=DRIVER_PATH)
    driver, wait = manager.setup()
    if not driver: return

    driver.get(URL_SEI_SP); time.sleep(3)
    FormHandler(driver, wait).executar_fluxo()
    ext = ResultadoExtractor(driver, wait)
    ext.carregar_resultados()
    links = ext.extrair_links()
    df = pd.DataFrame([dict(zip(["Nome", "CNPJ"], ext.extrair_dados(l))) for l in links])
    df.to_excel("dados_extraidos.xlsx", index=False)
    time.sleep(5)
    driver.quit()

if __name__ == "__main__":
    main()
