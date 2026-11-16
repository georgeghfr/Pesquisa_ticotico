from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time


def esperar_resultado(driver):
    print("\nAguardando resultado da consulta...")

    def procurar_linha(d):
        tabelas = d.find_elements(By.ID, "form:ticoTable")
        if not tabelas:
            return False

        tabela = tabelas[0]
        linhas = tabela.find_elements(By.XPATH, ".//tr")

        for linha in linhas:
            tds = linha.find_elements(By.TAG_NAME, "td")
            if len(tds) >= 9:
                return linha

        return False

    linha = WebDriverWait(driver, 30).until(procurar_linha)
    print("Linha de resultado encontrada.")
    return linha


def aguardar_mascaras(driver, tempo=30):
    print("\nAguardando processamento da página...")

    def pagina_livre(d):
        elementos = d.find_elements(
            By.CSS_SELECTOR, "div.ui-widget-overlay.ui-dialog-mask"
        )
        visiveis = [elemento for elemento in elementos if elemento.is_displayed()]
        return len(visiveis) == 0

    try:
        WebDriverWait(driver, tempo).until(pagina_livre)
        print("Página liberada.")
    except TimeoutException:
        print("Aviso: uma máscara permaneceu visível após o tempo limite.")


def clicar_consultar(driver):
    print("\nLocalizando botão Consultar...")
    aguardar_mascaras(driver)

    botao = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "form:nonAjax2"))
    )
    print("Botão encontrado.")

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao)

    def botao_livre(d):
        try:
            return (
                botao.is_displayed()
                and botao.is_enabled()
                and d.execute_script(
                    """
                    const el = arguments[0];
                    const rect = el.getBoundingClientRect();
                    const x = rect.left + rect.width / 2;
                    const y = rect.top + rect.height / 2;
                    const top = document.elementFromPoint(x, y);

                    return top === el || el.contains(top);
                """,
                    botao,
                )
            )
        except:
            return False

    WebDriverWait(driver, 30).until(botao_livre)
    print("Botão liberado.")
    print("Clicando em Consultar...")

    botao.click()
    print("Consulta acionada.")

    time.sleep(1.5)
    aguardar_mascaras(driver)


def preencher_cpf(driver, cpf):
    print("\nProcurando campo CPF...")
    aguardar_mascaras(driver)

    campos = driver.find_elements(
        By.XPATH,
        "//label[contains(normalize-space(), 'Por CPF')]/following-sibling::input",
    )

    print(f"Campos CPF encontrados: {len(campos)}")
    campo = None

    for elemento in campos:
        if elemento.is_displayed() and elemento.is_enabled():
            campo = elemento
            break

    if campo is None:
        raise Exception(
            "O campo CPF foi localizado, mas nenhum campo está visível e habilitado."
        )

    print("Campo CPF visível encontrado.")
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo)

    WebDriverWait(driver, 10).until(
        lambda d: campo.is_displayed() and campo.is_enabled()
    )

    driver.execute_script("arguments[0].focus();", campo)
    campo.clear()
    campo.send_keys(str(cpf))
    print(f"CPF preenchido: {cpf}")


def extrair_promocoes(td):
    try:
        abbr = td.find_element(By.TAG_NAME, "abbr")
    except:
        return "", ""

    numero_quadro = abbr.text.strip()
    texto = abbr.get_attribute("title") or ""
    promocoes = []

    for parte in texto.split("|"):
        parte = parte.strip()
        if " em " in parte:
            promocoes.append(parte)

    return numero_quadro, "; ".join(promocoes)


def extrair_nome(td):
    texto = td.text.split("\n")
    nome = texto[0].strip()
    try:
        guerra = td.find_element(By.TAG_NAME, "b").text.strip()
    except:
        guerra = ""
    return nome, guerra


def clicar_detalhes(driver, linha):
    print("\nLocalizando botão de detalhes...")
    botao = linha.find_element(By.CSS_SELECTOR, "button.btn-info.btn-xs")
    print("Botão de detalhes encontrado.")
    print("Abrindo informações detalhadas...")

    botao.click()
    time.sleep(2)

    tabela_detalhes = WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "div.ui-dialog-content table")
        )
    )
    print("Informações detalhadas disponíveis.")

    return tabela_detalhes


def extrair_detalhes(driver, tabela):
    print("\nExtraindo informações detalhadas...")
    detalhes = {}
    linhas = tabela.find_elements(By.XPATH, ".//tbody/tr")

    for linha in linhas:
        celulas = linha.find_elements(By.TAG_NAME, "td")
        if len(celulas) < 2:
            continue

        campo = " ".join(celulas[0].text.split()).strip()
        valor = " ".join(celulas[1].text.split()).strip()
        campo = campo.rstrip(":")

        if campo:
            detalhes[campo] = valor
            print(f"{campo}: {valor}")

    return detalhes


def consultar_cpf(driver, cpf):
    preencher_cpf(driver, cpf)
    clicar_consultar(driver)

    try:
        linha = esperar_resultado(driver)
    except TimeoutException:
        print("\nResultado não encontrado na primeira tentativa.")
        print("Realizando uma nova tentativa...")

        clicar_consultar(driver)
        try:
            linha = esperar_resultado(driver)
        except TimeoutException:
            print("\nCPF não encontrado no sistema.")
            print(f"CPF: {cpf}")
            print("Passando para o próximo CPF...")

            return {
                "CPF": cpf,
                "SARAM": "",
                "OM": "",
                "Nome": "",
                "Nome Guerra": "",
                "Quadro": "",
                "Posto": "",
                "Nº Quadro": "",
                "Primeira Promoção": "",
                "Promoções": "",
                "Situação OM": "",
                "Erro": "CPF não encontrado no sistema",
            }

    tds = linha.find_elements(By.TAG_NAME, "td")
    print(f"\nNúmero de colunas encontradas: {len(tds)}")

    if len(tds) < 9:
        return {
            "CPF": cpf,
            "SARAM": "",
            "OM": "",
            "Nome": "",
            "Nome Guerra": "",
            "Quadro": "",
            "Posto": "",
            "Nº Quadro": "",
            "Primeira Promoção": "",
            "Promoções": "",
            "Situação OM": "",
            "Erro": f"Linha de resultado possui apenas {len(tds)} colunas",
        }

    numero_quadro, promocoes = extrair_promocoes(tds[0])
    nome, guerra = extrair_nome(tds[4])

    tabela_detalhes = clicar_detalhes(driver, linha)
    detalhes = extrair_detalhes(driver, tabela_detalhes)

    dados = {
        "CPF": cpf,
        "SARAM": tds[3].text.strip(),
        "OM": tds[6].text.strip(),
        "Nome": nome,
        "Nome Guerra": guerra,
    }

    for campo, valor in detalhes.items():
        dados[campo] = valor

    dados.update(
        {
            "Quadro": tds[2].text.strip(),
            "Posto": tds[1].text.strip(),
            "Nº Quadro": numero_quadro,
            "Primeira Promoção": tds[5].text.strip(),
            "Promoções": promocoes,
            "Situação OM": tds[7].text.strip(),
            "Erro": "",
        }
    )

    print("\nDados encontrados:")
    print(dados)

    return dados
