import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

def configurar_driver():
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    chrome_driver_path = os.path.join(diretorio_atual, "chromedriver.exe")

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless")  # Executa sem abrir a interface gráfica
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    driver_service = webdriver.chrome.service.Service(chrome_driver_path)
    driver = webdriver.Chrome(service=driver_service, options=chrome_options)
    return driver

def extrair_dados_estabelecimentos(driver, codigos):
    resultados = []
    url = "https://cnes.datasus.gov.br/pages/estabelecimentos/consulta.jsp"
    driver.get(url)

    wait = WebDriverWait(driver, 10)

    for codigo in codigos:
        try:
            print(f"Processando código: {codigo}")
            pesquisa_input = wait.until(
                EC.presence_of_element_located((By.ID, "pesquisaValue"))
            )
            pesquisa_input.clear()
            pesquisa_input.send_keys(codigo)

            botao_pesquisa = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn-primary')]")
            ))
            botao_pesquisa.click()

            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tr.ng-scope")))

            linhas = driver.find_elements(By.CSS_SELECTOR, "tr.ng-scope")

            if not linhas:
                resultados.append({
                    "Código": codigo,
                    "UF": "Sem resultado",
                    "Município": "Sem resultado",
                    "CNES": "Sem resultado",
                    "Nome Fantasia": "Sem resultado",
                    "Natureza Jurídica": "Sem resultado",
                    "Gestão": "Sem resultado",
                    "Atende SUS": "Sem resultado"
                })
                continue

            for linha in linhas:
                try:
                    dados = {
                        "Código": codigo,
                        "UF": linha.find_element(By.CSS_SELECTOR, "td[data-title=\"'UF'\"]").text,
                        "Município": linha.find_element(By.CSS_SELECTOR, "td[data-title=\"'Município'\"]").text,
                        "CNES": linha.find_element(By.CSS_SELECTOR, "td[data-title=\"'CNES'\"]").text,
                        "Nome Fantasia": linha.find_element(By.CSS_SELECTOR, "td[data-title*='Nome']").text,
                        "Natureza Jurídica": linha.find_element(By.CSS_SELECTOR, "td[data-title=\"'Natureza Jurídica(Grupo)'\"]").text,
                        "Gestão": linha.find_element(By.CSS_SELECTOR, "td[data-title=\"'Gestão'\"]").text,
                        "Atende SUS": linha.find_element(By.CSS_SELECTOR, "td[data-title=\"'Atende SUS'\"]").text
                    }
                    resultados.append(dados)
                except (NoSuchElementException, StaleElementReferenceException):
                    print(f"Erro ao extrair dados para o código {codigo}, passando para o próximo.")
                    break

        except TimeoutException:
            print(f"Tempo esgotado ao processar código {codigo}")
            resultados.append({
                "Código": codigo,
                "UF": "Sem resultado",
                "Município": "Sem resultado",
                "CNES": "Sem resultado",
                "Nome Fantasia": "Sem resultado",
                "Natureza Jurídica": "Sem resultado",
                "Gestão": "Sem resultado",
                "Atende SUS": "Sem resultado"
            })
        except Exception as e:
            print(f"Erro ao processar código {codigo}: {e}")

    return resultados

def salvar_excel(resultados, nome_arquivo='estabelecimentos_cnes.xlsx'):
    if resultados:
        df = pd.DataFrame(resultados)
        with pd.ExcelWriter(nome_arquivo, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Estabelecimentos')
            worksheet = writer.sheets['Estabelecimentos']
            for col_num, value in enumerate(df.columns.values):
                worksheet.set_column(col_num, col_num, len(value) + 5)
        print(f"Dados salvos em {nome_arquivo}")
    else:
        print("Nenhum dado encontrado para salvar.")

def main():
    with open('codigos.txt', 'r') as arquivo:
        codigos = [linha.strip() for linha in arquivo if linha.strip()]

    driver = configurar_driver()

    try:
        resultados = extrair_dados_estabelecimentos(driver, codigos)
        salvar_excel(resultados)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
