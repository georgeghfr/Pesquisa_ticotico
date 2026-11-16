import os
import time
import traceback
from tkinter import Tk, filedialog
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from navegador import abrir_firefox
from ticotico import (
    consultar_cpf,
    aguardar_mascaras,
    clicar_detalhes,
    extrair_detalhes,
    extrair_promocoes,
    extrair_nome,
)
from excel import salvar_resultado_duas_abas


def escolher_planilha():
    root = Tk()
    root.withdraw()
    arquivo = filedialog.askopenfilename(
        title="Selecione a planilha", filetypes=[("Excel", "*.xlsx")]
    )
    root.destroy()
    return arquivo


def buscar_comandantes_oms(driver, oms_fab):
    print("\n" + "=" * 60)
    print("INICIANDO BUSCA DE COMANDANTES DE OM (ABA AVANÇADA COM REGRA CMT)")
    print("=" * 60)

    comandantes = []

    # Tabela de pesos hierárquicos caso nenhum tenha a flag explícita de CMT
    ordem_postos = {
        "TB": 1,
        "MB": 2,
        "BR": 3,
        "CEL": 4,
        "CL": 4,
        "TEN CEL": 5,
        "TC": 5,
        "MAJ": 6,
        "MJ": 6,
        "CAP": 7,
        "CP": 7,
        "1º TEN": 8,
        "1T": 8,
        "1º TEN (T)": 8,
        "2º TEN": 9,
        "2T": 9,
        "ASP": 10,
        "SO": 11,
        "1S": 12,
        "2S": 13,
        "3S": 14,
    }

    for om_alvo in oms_fab:
        print(f"\nBuscando Comandante da OM: {om_alvo}")
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

            # Clica no botão Consultar da aba Avançada (form:nonAjax3)
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
                    linhas_dados.append((linha, tds))

            if not linhas_dados:
                print("Nenhum militar listado na tabela desta OM.")
                continue

            # Analisa os até 3 primeiros militares aplicando a regra do "CMT"
            limite_analise = min(3, len(linhas_dados))
            melhor_indice = 0
            encontrou_cmt_explico = False

            print(
                f"Analisando os {limite_analise} primeiros militares de {om_alvo} para definir o comandante..."
            )

            for i in range(limite_analise):
                _, tds_cand = linhas_dados[i]
                posto_cand = tds_cand[1].text.strip().upper()

                # A coluna de Situação OM corresponde ao índice 7
                situacao_om_cand = (
                    tds_cand[7].text.strip().upper() if len(tds_cand) > 7 else ""
                )

                print(
                    f"  -> [{i+1}º militar] Posto: {posto_cand} | Situação OM: '{situacao_om_cand}'"
                )

                # Regra prioritária: Se tiver "CMT" na situação, ele é o comandante absoluto
                if "CMT" in situacao_om_cand:
                    melhor_indice = i
                    encontrou_cmt_explico = True
                    print(
                        f"🎯 Flag 'CMT' encontrada no {i+1}º militar! Definido como comandante."
                    )
                    break

            # Se nenhum dos 3 primeiros tiver explicitamente "CMT", recorre ao peso hierárquico
            if not encontrou_cmt_explico:
                melhor_peso = 999
                for i in range(limite_analise):
                    _, tds_cand = linhas_dados[i]
                    posto_cand = tds_cand[1].text.strip().upper()
                    peso_cand = ordem_postos.get(posto_cand, 99)

                    if peso_cand < melhor_peso:
                        melhor_peso = peso_cand
                        melhor_indice = i
                print(
                    f"🎯 Nenhum 'CMT' explícito nos primeiros. Usando hierarquia: {melhor_indice + 1}ª posição escolhida."
                )

            linha_cmd, tds = linhas_dados[melhor_indice]

            numero_quadro, promocoes = extrair_promocoes(tds[0])
            nome, guerra = extrair_nome(tds[4])

            # Clica no botão de detalhes do mais antigo para extrair todas as informações completas
            tabela_detalhes = clicar_detalhes(driver, linha_cmd)
            detalhes = extrair_detalhes(driver, tabela_detalhes)

            registro_cmd = {
                "CPF": "",
                "SARAM": tds[3].text.strip(),
                "Sexo": "",
                "Posto": tds[1].text.strip(),
                "Quadro": tds[2].text.strip(),
                "Especialidade": "",
                "Sub-Especialidade": "",
                "Nome": nome,
                "Nome Guerra": guerra,
                "OM": om_alvo,
                "Nº Quadro": numero_quadro,
                "Primeira Promoção": tds[5].text.strip(),
                "Promoções": promocoes,
                "Situação OM": tds[7].text.strip(),
                "Erro": "",
                "Força": "FAB",
            }

            for campo, valor in detalhes.items():
                registro_cmd[campo] = valor

            comandantes.append(registro_cmd)
            print(
                f"✅ Comandante detalhado definido para {om_alvo}: {registro_cmd['Posto']} {nome}"
            )

        except Exception as e:
            print(f"[ERRO ao consultar OM {om_alvo}]: {e}")

    return comandantes


def main():
    print("=" * 60)
    print("CONSULTA AUTOMÁTICA DE CPF - TICO TICO (DUAS ABAS)")
    print("=" * 60)

    arquivo = escolher_planilha()

    if not arquivo:
        print("Nenhuma planilha selecionada.")
        return

    print(f"\nPlanilha: {os.path.basename(arquivo)}")

    df_raw = pd.read_excel(arquivo, header=None)

    if len(df_raw) < 2:
        raise Exception("A planilha parece estar vazia.")

    # 1. Lê apenas o cabeçalho primeiro para descobrir o nome exato da coluna de CPF
    df_head = pd.read_excel(arquivo, nrows=0)
    col_cpf = None
    for col in df_head.columns:
        if str(col).strip().upper() == "CPF":
            col_cpf = col
            break

    if col_cpf is None:
        print("\n[ERRO CRÍTICO]: A coluna 'CPF' não foi encontrada na planilha.")
        print("O programa será encerrado por segurança.")
        input("\nPressione ENTER para fechar...")
        return

    # 2. Lê a planilha inteira forçando a coluna de CPF a ser lida como texto (str)
    df = pd.read_excel(arquivo, dtype={col_cpf: str})

    # Identificação automática da coluna "OM" pelo nome
    col_om = None
    for col in df.columns:
        col_limpa = str(col).strip().upper()
        if col_limpa in ["OM", "ORGANIZAÇÃO MILITAR", "ORGANIZACAO MILITAR"]:
            col_om = col
            break

    if col_om is None:
        print(
            "\n[ERRO CRÍTICO]: A coluna de 'OM' (Organização Militar) não foi encontrada na planilha."
        )
        print("O programa será encerrado por segurança.")
        input("\nPressione ENTER para fechar...")
        return

    # Identificação automática das colunas da base (suportando padrão Moodle e português)
    col_guerra_base = None
    col_nome_base = None
    col_posto_base = None
    col_quadro_base = None
    col_saram_base = None

    for col in df.columns:
        col_limpa = str(col).strip().lower()
        if col_limpa in ["nome de guerra", "guerra", "nome_guerra"]:
            col_guerra_base = col
        elif col_limpa in ["lastname", "nome"]:
            col_nome_base = col
        elif col_limpa in ["firstname", "posto"]:
            col_posto_base = col
        elif col_limpa in ["quadro"]:
            col_quadro_base = col
        elif col_limpa in ["saram", "matrícula", "matricula"]:
            col_saram_base = col

    idx_om_df = df.columns.get_loc(col_om)

    df["CPF_limpo"] = (
        df[col_cpf]
        .fillna("")
        .astype(str)
        .str.replace(r"\D", "", regex=True)
        .str.zfill(11)
    )

    # Remove linhas extras vazias na coluna de CPF
    df = df.dropna(subset=[col_cpf])
    df = df[df[col_cpf].str.strip() != ""]

    total = len(df)
    print(f"{total} linha(s) encontrada(s) na planilha.")
    print(f"-> Coluna CPF identificada: '{col_cpf}'")
    print(f"-> Coluna OM identificada: '{col_om}'")

    print("\nAbrindo Firefox...")
    driver = abrir_firefox()

    input("\nFaça login, entre em 'Por Documento' e pressione ENTER...")

    resultados = []

    termos_bloqueados = [
        "marinha",
        "marinha do brasil",
        "mb",
        "exército",
        "exercito",
        "exército brasileiro",
        "exercito brasileiro",
        "eb",
    ]

    # Conjunto que armazenará estritamente as OMs corretas retornadas pelo Tico-Tico
    oms_fab_detectadas = set()

    for indice, row in df.iterrows():
        row_raw = (
            df_raw.iloc[indice + 1] if indice + 1 < len(df_raw) else df_raw.iloc[indice]
        )

        cpf = str(row["CPF_limpo"]).strip()

        om_original = str(row.get(col_om, "")).strip()
        if not om_original or om_original.lower() == "nan":
            om_original = (
                str(row_raw.get(idx_om_df, "")).strip()
                if pd.notna(row_raw.get(idx_om_df))
                else ""
            )

        om_minuscula = om_original.lower()

        print("\n" + "=" * 60)
        print(f"Linha {indice + 1}/{total} | CPF: {cpf} | OM: {om_original}")
        print("=" * 60)

        pular_consulta = any(termo in om_minuscula for termo in termos_bloqueados)

        if pular_consulta:
            print(
                f"-> Pulando CPF {cpf}: OM '{om_original}' identificada como outra força (Marinha/Exército)."
            )

            nome_base = str(row.get(col_nome_base, "")).strip() if col_nome_base else ""
            guerra_base = (
                str(row.get(col_guerra_base, "")).strip() if col_guerra_base else ""
            )
            posto_base = (
                str(row.get(col_posto_base, "")).strip() if col_posto_base else ""
            )
            quadro_base = (
                str(row.get(col_quadro_base, "")).strip() if col_quadro_base else ""
            )
            saram_base = (
                str(row.get(col_saram_base, "")).strip() if col_saram_base else ""
            )

            registro = {
                "CPF": cpf,
                "SARAM": saram_base,
                "OM": om_original,
                "Nome": nome_base,
                "Nome Guerra": guerra_base,
                "Quadro": quadro_base,
                "Posto": posto_base,
                "Nº Quadro": "",
                "Primeira Promoção": "",
                "Promoções": "",
                "Situação OM": "",
                "Erro": "Pulado: Servidor de Outra Força (Marinha/Exército)",
            }
            resultados.append(registro)
            salvar_resultado_duas_abas(arquivo, df, resultados)
            continue

        try:
            dados = consultar_cpf(driver, cpf)

            if dados.get("Erro"):
                print(
                    f"-> Aviso: {dados.get('Erro')}. Preenchendo com dados da planilha (Possível baixa)."
                )
                dados["OM"] = om_original

            # CAPTURA DA OM CORRETA DIRETAMENTE DO RETORNO DO TICO-TICO
            om_ticotico = str(dados.get("OM", "")).strip()
            if om_ticotico and om_ticotico.upper() != "NAN":
                om_ticotico_lower = om_ticotico.lower()
                if not any(termo in om_ticotico_lower for termo in termos_bloqueados):
                    oms_fab_detectadas.add(om_ticotico)

            if not dados.get("Nome Guerra") and col_guerra_base:
                dados["Nome Guerra"] = str(row.get(col_guerra_base, "")).strip()

            resultados.append(dados)
            salvar_resultado_duas_abas(arquivo, df, resultados)
            print("OK")

        except Exception as erro:
            print("\nERRO:")
            print(erro)
            print("\nDETALHES DO ERRO:")
            traceback.print_exc()

            print("-> Preenchendo com dados da planilha (Possível baixa).")
            nome_base = str(row.get(col_nome_base, "")).strip() if col_nome_base else ""
            guerra_base = (
                str(row.get(col_guerra_base, "")).strip() if col_guerra_base else ""
            )
            posto_base = (
                str(row.get(col_posto_base, "")).strip() if col_posto_base else ""
            )
            quadro_base = (
                str(row.get(col_quadro_base, "")).strip() if col_quadro_base else ""
            )
            saram_base = (
                str(row.get(col_saram_base, "")).strip() if col_saram_base else ""
            )

            registro = {
                "CPF": cpf,
                "SARAM": saram_base,
                "OM": om_original,
                "Nome": nome_base,
                "Nome Guerra": guerra_base,
                "Quadro": quadro_base,
                "Posto": posto_base,
                "Nº Quadro": "",
                "Primeira Promoção": "",
                "Promoções": "",
                "Situação OM": "",
                "Erro": "Possível baixa",
            }
            resultados.append(registro)
            salvar_resultado_duas_abas(arquivo, df, resultados)

    # =========================================================================
    # FASE FINAL: BUSCA DE COMANDANTES APÓS A CONSULTA DOS ALUNOS
    # =========================================================================
    if oms_fab_detectadas:
        print(
            f"\nTotal de OMs exclusivas da FAB detectadas no Tico-Tico: {len(oms_fab_detectadas)}"
        )
        input(
            "\nAlunos consultados com sucesso! Vá manualmente até a aba 'Avançada' no Tico-Tico e pressione ENTER para iniciar a busca dos comandantes..."
        )

        comandantes = buscar_comandantes_oms(driver, list(oms_fab_detectadas))
        if comandantes:
            print(
                f"\n{len(comandantes)} comandante(s) encontrado(s). Salvando na aba 'Comandantes'..."
            )
            salvar_resultado_duas_abas(arquivo, df, resultados, comandantes)

    print("\nConsulta finalizada!")
    driver.quit()


if __name__ == "__main__":
    main()
