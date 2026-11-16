import os
import time
from tkinter import Tk, filedialog
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from navegador import abrir_firefox


def escolher_planilha():
    root = Tk()
    root.withdraw()
    arquivo = filedialog.askopenfilename(
        title="Selecione a planilha", filetypes=[("Excel", "*.xlsx")]
    )
    root.destroy()
    return arquivo


def aguardar_mascaras(driver, tempo=30):
    print("\nAguardando processamento da página...")

    def pagina_livre(d):
        elementos = d.find_elements(
            By.CSS_SELECTOR, "div.ui-widget-overlay.ui-dialog-mask"
        )
        visiveis = [e for e in elementos if e.is_displayed()]
        return len(visiveis) == 0

    try:
        WebDriverWait(driver, tempo).until(pagina_livre)
        print("Página liberada.")
    except TimeoutException:
        print("Aviso: máscara permaneceu visível.")


def testar_busca_comandantes():
    print("=" * 60)
    print("TESTE ISOLADO: IDENTIFICAÇÃO DE COMANDANTES POR OM")
    print("=" * 60)

    arquivo = escolher_planilha()
    if not arquivo:
        print("Nenhuma planilha selecionada.")
        return

    df = pd.read_excel(arquivo)

    col_om = None
    for col in df.columns:
        col_limpa = str(col).strip().upper()
        if col_limpa in ["OM", "ORGANIZAÇÃO MILITAR", "ORGANIZACAO MILITAR"]:
            col_om = col
            break

    if not col_om:
        print("[ERRO]: Coluna de OM não encontrada na planilha.")
        return

    # Regra: Marinha e Exército são ignorados nesta automação de OMs
    termos_bloqueados = ["marinha", "mb", "exército", "exercito", "eb"]

    oms_brutas = df[col_om].dropna().astype(str).str.strip().unique()
    oms_fab = []

    for om in oms_brutas:
        om_lower = om.lower()
        if (
            not any(termo in om_lower for termo in termos_bloqueados)
            and om.upper() != "NAN"
        ):
            if om not in oms_fab:
                oms_fab.append(om)

    print(f"\nTotal de OMs exclusivas da FAB encontradas na planilha: {len(oms_fab)}")
    print(f"OMs: {oms_fab}")

    print("\nAbrindo Firefox para o teste...")
    driver = abrir_firefox()

    input(
        "\nFaça login no Tico-Tico, vá até a aba 'Avançada' e depois pressione ENTER aqui..."
    )

    comandantes_encontrados = []

    for om_alvo in oms_fab:
        print(f"\n--------------------------------------------------")
        print(f"Buscando Comandante da OM: {om_alvo}")
        print(f"--------------------------------------------------")

        try:
            aguardar_mascaras(driver)

            # Localiza o combo box de OMs na aba Avançada
            combo_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, "form:comboOM"))
            )
            select_om = Select(combo_element)

            try:
                select_om.select_by_visible_text(om_alvo)
                print(f"OM '{om_alvo}' selecionada no combo.")
            except Exception:
                print(f"[AVISO]: OM '{om_alvo}' não encontrada no <select>. Pulando...")
                continue

            # Localiza e clica no botão Consultar da aba Avançada (form:nonAjax3)
            print("Localizando botão Consultar (Avançada)...")
            aguardar_mascaras(driver)

            botao = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "form:nonAjax3"))
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", botao
            )

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
            driver.execute_script("arguments[0].click();", botao)
            print("Consulta acionada.")

            time.sleep(1.5)
            aguardar_mascaras(driver)

            # Aguarda carregar a tabela de resultados
            tabelas = WebDriverWait(driver, 30).until(
                EC.presence_of_all_elements_located((By.ID, "form:ticoTable"))
            )

            if not tabelas:
                print("Tabela de resultado não encontrada para esta OM.")
                continue

            tabela = tabelas[0]
            linhas = tabela.find_elements(By.XPATH, ".//tr")

            linhas_dados = []
            for linha in linhas:
                tds = linha.find_elements(By.TAG_NAME, "td")
                if len(tds) >= 9:
                    linhas_dados.append(tds)

            if not linhas_dados:
                print("Nenhum militar listado na tabela desta OM.")
                continue

            # Analisa os 3 primeiros militares para identificar o comandante/chefe/diretor
            melhor_candidato = None
            limite_analise = min(3, len(linhas_dados))

            for i in range(limite_analise):
                tds = linhas_dados[i]

                posto = tds[1].text.strip()
                quadro = tds[2].text.strip()
                nome_completo = tds[4].text.strip()

                # Tenta extrair o nome de guerra se houver tag <b> dentro da 5ª coluna
                try:
                    guerra = tds[4].find_element(By.TAG_NAME, "b").text.strip()
                except:
                    guerra = ""

                situacao_om = tds[7].text.strip() if len(tds) > 7 else ""

                print(
                    f"  -> [{i+1}º] Posto: {posto} | Quadro: {quadro} | Nome: {nome_completo} | Situação: {situacao_om}"
                )

                if i == 0:
                    melhor_candidato = {
                        "Posto": posto,
                        "Quadro": quadro,
                        "Nome": nome_completo,
                        "Nome Guerra": guerra,
                        "OM_Comandada": om_alvo,
                        "Situacao": situacao_om,
                    }

            if melhor_candidato:
                print(
                    f"✅ Comandante identificado para {om_alvo}: {melhor_candidato['Posto']} {melhor_candidato['Nome']} ({om_alvo})"
                )
                comandantes_encontrados.append(melhor_candidato)

        except Exception as e:
            print(f"[ERRO ao consultar OM {om_alvo}]: {e}")

    # Salva os comandantes encontrados na mesma pasta do arquivo base da consulta
    if comandantes_encontrados:
        df_com = pd.DataFrame(comandantes_encontrados)
        pasta_base = os.path.dirname(arquivo)
        saida_comandantes = os.path.join(pasta_base, "COMANDANTES_FAB_RESULTADO.xlsx")
        df_com.to_excel(saida_comandantes, index=False)
        print(f"\n📁 Planilha de comandantes salva com sucesso em: {saida_comandantes}")

    print("\n" + "=" * 60)
    print(
        f"RESUMO: {len(comandantes_encontrados)} comandantes identificados com sucesso."
    )
    print("=" * 60)
    for c in comandantes_encontrados:
        print(f"OM: {c['OM_Comandada']} -> {c['Posto']} {c['Nome']}")

    driver.quit()


if __name__ == "__main__":
    testar_busca_comandantes()
