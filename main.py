import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException


def setup_driver():
    # Configura o driver do Chrome com as opções necessárias
    current_dir = os.path.dirname(os.path.abspath(__file__))
    chrome_driver_path = os.path.join(current_dir, "chromedriver.exe")

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless")  # Executa sem abrir a interface gráfica
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--memory-pressure-off")  # Reduz pressão na memória
    chrome_options.add_argument("--disk-cache-size=1")  # Minimiza cache em disco

    driver_service = webdriver.chrome.service.Service(chrome_driver_path)
    driver = webdriver.Chrome(service=driver_service, options=chrome_options)
    return driver


def extract_establishment_data(driver, codes, batch_size=50):
    # Processa os códigos em lotes para reduzir uso de memória
    results = []
    url = "https://cnes.datasus.gov.br/pages/estabelecimentos/consulta.jsp"
    wait = WebDriverWait(driver, 10)

    for i in range(0, len(codes), batch_size):
        batch_codes = codes[i:i + batch_size]
        print(f"Processando lote {i // batch_size + 1} de {len(codes) // batch_size + 1}")

        for code in batch_codes:
            try:
                driver.get(url)  # Recarrega a página para cada novo código
                print(f"Processando código: {code}")

                search_input = wait.until(
                    EC.presence_of_element_located((By.ID, "pesquisaValue"))
                )
                search_input.clear()
                search_input.send_keys(code)

                search_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn-primary')]"))
                )
                search_button.click()

                # Aguarda resultados ou timeout
                try:
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tr.ng-scope")))
                    rows = driver.find_elements(By.CSS_SELECTOR, "tr.ng-scope")
                except TimeoutException:
                    rows = []

                if not rows:
                    results.append(create_empty_result(code))
                    continue

                for row in rows:
                    try:
                        data = extract_row_data(row, code)
                        results.append(data)

                        # Salva resultados parciais a cada 100 registros
                        if len(results) % 100 == 0:
                            save_partial_results(results)

                    except (NoSuchElementException, StaleElementReferenceException):
                        print(f"Erro ao extrair dados para o código {code}")
                        break

            except Exception as e:
                print(f"Erro ao processar código {code}: {e}")
                results.append(create_empty_result(code))

        # Limpa a memória do driver periodicamente
        driver.execute_script("window.localStorage.clear();")
        driver.execute_script("window.sessionStorage.clear();")
        driver.delete_all_cookies()

    return results


def create_empty_result(code):
    # Cria um resultado vazio padronizado
    return {
        "Código": code,
        "UF": "Sem resultado",
        "Município": "Sem resultado",
        "CNES": "Sem resultado",
        "Nome Fantasia": "Sem resultado",
        "Natureza Jurídica": "Sem resultado",
        "Gestão": "Sem resultado",
        "Atende SUS": "Sem resultado"
    }


def extract_row_data(row, code):
    # Extrai dados de uma linha da tabela
    return {
        "Código": code,
        "UF": row.find_element(By.CSS_SELECTOR, "td[data-title=\"'UF'\"]").text,
        "Município": row.find_element(By.CSS_SELECTOR, "td[data-title=\"'Município'\"]").text,
        "CNES": row.find_element(By.CSS_SELECTOR, "td[data-title=\"'CNES'\"]").text,
        "Nome Fantasia": row.find_element(By.CSS_SELECTOR, "td[data-title*='Nome']").text,
        "Natureza Jurídica": row.find_element(By.CSS_SELECTOR, "td[data-title=\"'Natureza Jurídica(Grupo)'\"]").text,
        "Gestão": row.find_element(By.CSS_SELECTOR, "td[data-title=\"'Gestão'\"]").text,
        "Atende SUS": row.find_element(By.CSS_SELECTOR, "td[data-title=\"'Atende SUS'\"]").text
    }


def save_partial_results(results, filename='resultados_parciais.xlsx'):
    # Salva resultados parciais com ajuste de largura das colunas
    df = pd.DataFrame(results)
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Estabelecimentos')
        worksheet = writer.sheets['Estabelecimentos']

        for idx, col in enumerate(df.columns):
            # Calcula a largura máxima com base no conteúdo
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(col)
            ) + 2
            worksheet.set_column(idx, idx, max_length)

    print(f"Resultados parciais salvos em {filename}")


def save_excel(results, filename='estabelecimentos_cnes.xlsx'):
    # Salva resultados finais com ajuste de largura das colunas
    if results:
        df = pd.DataFrame(results)
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Estabelecimentos')
            worksheet = writer.sheets['Estabelecimentos']

            for idx, col in enumerate(df.columns):
                # Calcula a largura máxima com base no conteúdo
                max_length = max(
                    df[col].astype(str).apply(len).max(),
                    len(col)
                ) + 2
                worksheet.set_column(idx, idx, max_length)

        print(f"Dados salvos em {filename}")
    else:
        print("Nenhum dado encontrado para salvar.")


def main():
    # Lê os códigos do arquivo
    with open('codigos.txt', 'r') as file:
        codes = [line.strip() for line in file if line.strip()]

    driver = setup_driver()

    try:
        results = extract_establishment_data(driver, codes)
        save_excel(results)
    finally:
        driver.quit()


if __name__ == "__main__":
    main()