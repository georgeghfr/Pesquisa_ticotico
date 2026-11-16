import os
import pandas as pd
import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

# =========================================================================
# TABELAS DE HIERARQUIA E ORDENAÇÃO POR FORÇA (Do mais antigo para o mais moderno)
# =========================================================================
ordem_forcas = {"MARINHA": 1, "EXÉRCITO": 2, "FAB": 3}

ordem_marinha = {
    "CMG": 1,
    "CF": 2,
    "CC": 3,
    "CT": 4,
    "1º TEN (MB)": 5,
    "2º TEN (MB)": 6,
    "GM": 7,
    "ASP": 8,
    "SO": 9,
    "1º SG": 10,
    "2º SG": 11,
    "3º SG": 12,
    "CB": 13,
    "SD (CFN)": 14,
    "MN": 15,
}

ordem_exercito = {
    "CEL (EB)": 1,
    "TEN CEL (EB)": 2,
    "MAJ (EB)": 3,
    "CAP (EB)": 4,
    "1º TEN (EB)": 5,
    "2º TEN (EB)": 6,
    "ASP (EB)": 7,
    "CAD": 8,
    "S TEN": 9,
    "1º SGT": 10,
    "2º SGT": 11,
    "3º SGT": 12,
    "CB": 13,
    "SD": 14,
    "TM": 15,
    "T1": 16,
    "T2": 17,
}

ordem_fab_postos = {
    "CEL": 1,
    "CL": 1,
    "TEN CEL": 2,
    "TC": 2,
    "MAJ": 3,
    "MJ": 3,
    "CAP": 4,
    "CP": 4,
    "1º TEN": 5,
    "1T": 5,
    "1º TEN (T)": 5,
    "2º TEN": 6,
    "2T": 6,
    "ASP": 7,
    "SO": 8,
    "1S": 9,
    "2S": 10,
    "3S": 11,
}


def ajustar_posto_outras_forcas(posto, forca):
    posto_limpo = str(posto).replace("(MB)", "").replace("(EB)", "").strip().upper()
    if forca == "MARINHA":
        if posto_limpo in ["1º TEN", "1T", "2º TEN", "2T"]:
            return "1º TEN (MB)" if "1" in posto_limpo else "2º TEN (MB)"
        return posto_limpo
    elif forca == "EXÉRCITO":
        if posto_limpo in [
            "CEL",
            "TEN CEL",
            "TC",
            "MAJ",
            "MJ",
            "CAP",
            "CP",
            "1º TEN",
            "1T",
            "2º TEN",
            "2T",
            "ASP",
        ]:
            if posto_limpo in ["CEL"]:
                return "CEL (EB)"
            elif posto_limpo in ["TEN CEL", "TC"]:
                return "TEN CEL (EB)"
            elif posto_limpo in ["MAJ", "MJ"]:
                return "MAJ (EB)"
            elif posto_limpo in ["CAP", "CP"]:
                return "CAP (EB)"
            elif "1" in posto_limpo:
                return "1º TEN (EB)"
            elif "2" in posto_limpo:
                return "2º TEN (EB)"
            elif posto_limpo == "ASP":
                return "ASP (EB)"
        return posto_limpo
    return posto_limpo


def ordenar_por_antiguidade(df_entrada):
    """Função auxiliar para aplicar os critérios de ordenação hierárquica e formatação de dados."""
    df_aux = df_entrada.copy()

    if df_aux.empty:
        df_aux["Dados_formatados"] = []
        return df_aux

    # Atribuição de Força e ajuste de postos para ordenação
    for idx, row in df_aux.iterrows():
        om = str(row.get("OM", "")).lower()
        if "marinha" in om or "mb" in om:
            df_aux.at[idx, "_forca"] = "MARINHA"
        elif "exército" in om or "exercito" in om or "eb" in om:
            df_aux.at[idx, "_forca"] = "EXÉRCITO"
        else:
            df_aux.at[idx, "_forca"] = "FAB"

        p_atual = str(row.get("Posto", "")).strip().upper()
        if df_aux.at[idx, "_forca"] in ["MARINHA", "EXÉRCITO"]:
            df_aux.at[idx, "Posto"] = ajustar_posto_outras_forcas(
                p_atual, df_aux.at[idx, "_forca"]
            )

    # Pesos de ordenação
    df_aux["_forca_ordem"] = df_aux["_forca"].map(lambda x: ordem_forcas.get(x, 3))
    df_aux["_posto_ordem"] = df_aux.apply(
        lambda r: (
            ordem_marinha.get(str(r["Posto"]).strip().upper(), 99)
            if r["_forca"] == "MARINHA"
            else (
                ordem_exercito.get(str(r["Posto"]).strip().upper(), 99)
                if r["_forca"] == "EXÉRCITO"
                else ordem_fab_postos.get(str(r["Posto"]).strip().upper(), 99)
            )
        ),
        axis=1,
    )

    col_data_promo = (
        "Data da Promoção Atual"
        if "Data da Promoção Atual" in df_aux.columns
        else "Primeira Promoção"
    )
    df_aux["_data_promo"] = pd.to_datetime(
        df_aux.get(col_data_promo, ""), format="%d/%m/%Y", errors="coerce"
    )
    df_aux["_n_quadro_num"] = pd.to_numeric(
        df_aux.get("Nº Quadro", ""), errors="coerce"
    ).fillna(999999)

    colunas_ordenacao = ["_forca_ordem", "_posto_ordem", "_data_promo"]
    if "Quadro" in df_aux.columns:
        colunas_ordenacao.append("Quadro")
    colunas_ordenacao.append("_n_quadro_num")

    df_aux = df_aux.sort_values(
        by=colunas_ordenacao,
        ascending=[True] * len(colunas_ordenacao),
    ).drop(
        columns=[
            c
            for c in [
                "_forca",
                "_forca_ordem",
                "_posto_ordem",
                "_data_promo",
                "_n_quadro_num",
            ]
            if c in df_aux.columns
        ]
    )

    map_posto = {
        "ASP": "Asp",
        "1T": "1º Ten",
        "2T": "2º Ten",
        "CP": "Cap",
        "MJ": "Maj",
        "TC": "Ten Cel",
        "CL": "Cel",
    }

    map_quadro = {
        "QOAV": "Av",
        "QOINF": "Inf",
        "QOINT": "Int",
        "QOENG": "Eng",
        "QOMED": "Med",
        "QOARM": "Esp Arm",
        "QOEAV": "Esp Av",
        "QOECOM": "Esp Com",
        "QOECTA": "Esp CTA",
        "QOEFOT": "Esp Fot",
        "QOEMET": "Esp Met",
        "QOESUP": "Esp Sup Tec",
        "QOEA": "Esp Aer SVA",
        "QOCON": "QOCon",
    }

    dados_formatados_lista = []
    oficiais = ["ASP", "2T", "1T", "CP", "MJ", "TC", "CL"]

    for idx, row in df_aux.iterrows():
        p_raw = str(row.get("Posto", "")).strip().upper()
        q_raw = str(row.get("Quadro", "")).strip().upper()
        esp_raw = str(row.get("Especialidade", "")).strip().upper()
        nome = str(row.get("Nome", "")).strip()
        om = str(row.get("OM", "")).strip()

        om_lower = om.lower()
        is_marinha = "marinha" in om_lower or "mb" in om_lower
        is_exercito = (
            "exército" in om_lower or "exercito" in om_lower or "eb" in om_lower
        )

        if is_marinha or is_exercito:
            nome_forca = "MARINHA" if is_marinha else "EXÉRCITO"
            partes = [p_raw, q_raw, esp_raw, nome, nome_forca]
            partes_filtradas = [
                str(p).strip()
                for p in partes
                if p and str(p).strip() not in ["NAN", "NTE"]
            ]
            linha_fmt = " - ".join(partes_filtradas)
        else:
            posto_fmt = map_posto.get(p_raw, p_raw)
            quadro_fmt = map_quadro.get(q_raw, q_raw)
            saram = str(row.get("SARAM", "")).strip()

            if p_raw in oficiais:
                esp_parte = "" if esp_raw == "NTE" or not esp_raw else esp_raw
                q_str = quadro_fmt if quadro_fmt else ""

                sub_partes = [posto_fmt, q_str, esp_parte]
                bloco_inicial = " ".join(
                    [
                        str(p).strip()
                        for p in sub_partes
                        if p and str(p).strip() not in ["NAN", "NTE"]
                    ]
                )

                partes = [bloco_inicial, nome, saram, om]
                partes_filtradas = [
                    str(p).strip()
                    for p in partes
                    if p and str(p).strip() not in ["NAN", "NTE"]
                ]
                linha_fmt = " - ".join(partes_filtradas)
            else:
                esp_str = esp_raw if esp_raw and esp_raw != "NAN" else ""

                sub_partes = [posto_fmt, esp_str]
                bloco_inicial = " ".join(
                    [
                        str(p).strip()
                        for p in sub_partes
                        if p and str(p).strip() not in ["NAN", "NTE"]
                    ]
                )

                partes = [bloco_inicial, nome, saram, om]
                partes_filtradas = [
                    str(p).strip()
                    for p in partes
                    if p and str(p).strip() not in ["NAN", "NTE"]
                ]
                linha_fmt = " - ".join(partes_filtradas)

        dados_formatados_lista.append(linha_fmt)

    df_aux["Dados_formatados"] = dados_formatados_lista
    return df_aux


def salvar_resultado_duas_abas(arquivo_original, df_base, resultados, comandantes=None):
    pasta = os.path.dirname(arquivo_original)
    nome = os.path.splitext(os.path.basename(arquivo_original))[0]
    arquivo_saida = os.path.join(pasta, f"{nome}_RESULTADO.xlsx")

    df_res = pd.DataFrame(resultados)

    colunas_padrao = [
        "CPF",
        "SARAM",
        "Sexo",
        "Posto",
        "Quadro",
        "Especialidade",
        "Sub-Especialidade",
        "Nome",
        "Nome Guerra",
        "OM",
        "Nº Quadro",
        "Primeira Promoção",
        "Data da Promoção Atual",
        "Promoções",
        "Situação OM",
        "Erro",
        "Dados_formatados",
    ]

    for coluna in colunas_padrao:
        if coluna not in df_res.columns:
            df_res[coluna] = ""
        else:
            df_res[coluna] = df_res[coluna].fillna("").astype(str).str.upper()

    for idx, row in df_res.iterrows():
        full_name = str(row.get("Nome", "")).strip()
        war_name = str(row.get("Nome Guerra", "")).strip()
        if full_name and war_name:
            if full_name.endswith(" " + war_name):
                without_last = full_name[: -len(war_name)].strip()
                if war_name in without_last:
                    df_res.at[idx, "Nome"] = without_last

    # Aba 1: Cópia da Base com dados do Tico-Tico preenchidos nas colunas seguintes
    df_aba1 = df_base.copy()

    if "CPF_limpo" in df_aba1.columns:
        df_aba1 = df_aba1.drop(columns=["CPF_limpo"])

    colunas_tico_extras = [c for c in colunas_padrao if c not in df_aba1.columns]
    for c in colunas_tico_extras:
        df_aba1[c] = ""

    for idx, row in df_aba1.iterrows():
        cpf_base = str(row.get("CPF", "")).strip()
        match = df_res[df_res["CPF"].str.contains(cpf_base, na=False)]
        if not match.empty:
            for col in colunas_padrao:
                if col in match.columns and col != "CPF":
                    df_aba1.at[idx, col] = match.iloc[0][col]

    # Aba 2: Alunos Ordenados por Antiguidade e Formatados
    df_aba2 = ordenar_por_antiguidade(df_res.copy())

    # Mapeamento de colunas da base original (Moodle/português) para a aba 2
    for idx, row in df_aba2.iterrows():
        cpf_atual = str(row.get("CPF", "")).strip()
        match_base = (
            df_base[
                df_base.astype(str)
                .apply(lambda x: x.str.contains(cpf_atual, na=False))
                .any(axis=1)
            ]
            if cpf_atual
            else pd.DataFrame()
        )

        if not match_base.empty:
            for col_b in match_base.columns:
                col_b_lower = str(col_b).strip().lower()
                val_b = str(match_base.iloc[0][col_b]).strip()

                if (
                    col_b_lower in ["firstname", "posto"]
                    and not str(row.get("Posto", "")).strip()
                ):
                    df_aba2.at[idx, "Posto"] = val_b
                elif (
                    col_b_lower in ["lastname", "nome"]
                    and not str(row.get("Nome", "")).strip()
                ):
                    df_aba2.at[idx, "Nome"] = val_b
                elif col_b_lower in ["om"] and not str(row.get("OM", "")).strip():
                    df_aba2.at[idx, "OM"] = val_b

    df_aba2 = df_aba2[colunas_padrao]

    # Aba 3 (Opcional): Comandantes - Garantindo 100% de preservação dos registros
    df_comandantes_ordenado = None
    if comandantes:
        df_cmd_raw = pd.DataFrame(comandantes)
        for coluna in colunas_padrao:
            if coluna not in df_cmd_raw.columns:
                df_cmd_raw[coluna] = ""
            else:
                df_cmd_raw[coluna] = (
                    df_cmd_raw[coluna].fillna("").astype(str).str.upper()
                )

        # REVISÃO DEFINITIVA: Se por acaso algum dicionário veio vazio ou sem nome,
        # garantimos que ele não seja descartado criando um registro padrão com a OM para preservar a contagem.
        for idx, row in df_cmd_raw.iterrows():
            if not str(row.get("Nome", "")).strip():
                df_cmd_raw.at[idx, "Nome"] = "COMANDANTE NÃO LOCALIZADO"

        try:
            df_comandantes_ordenado = ordenar_por_antiguidade(df_cmd_raw)
            df_comandantes_ordenado = df_comandantes_ordenado[colunas_padrao]
        except Exception as e:
            print(
                f"[AVISO]: Erro ao ordenar comandantes: {e}. Mantendo registros completos sem ordenação estrita."
            )
            df_comandantes_ordenado = df_cmd_raw[colunas_padrao]

    # Salvando as abas no Excel
    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
        df_aba1.to_excel(writer, sheet_name="Consolidado_Base", index=False)
        df_aba2.to_excel(writer, sheet_name="Antiguidade_Formatada", index=False)

        if df_comandantes_ordenado is not None:
            df_comandantes_ordenado.to_excel(
                writer, sheet_name="Comandantes", index=False
            )

    wb = openpyxl.load_workbook(arquivo_saida)

    # Formatação de negrito no Nome de Guerra para as abas aplicáveis
    abas_para_formatar = ["Antiguidade_Formatada"]
    if "Comandantes" in wb.sheetnames:
        abas_para_formatar.append("Comandantes")

    for nome_aba in abas_para_formatar:
        if nome_aba in wb.sheetnames:
            ws = wb[nome_aba]
            header_row = [cell.value for cell in ws[1]]
            idx_nome = header_row.index("Nome") + 1 if "Nome" in header_row else None
            idx_guerra = (
                header_row.index("Nome Guerra") + 1
                if "Nome Guerra" in header_row
                else None
            )
            idx_dados = (
                header_row.index("Dados_formatados") + 1
                if "Dados_formatados" in header_row
                else None
            )

            if idx_guerra and (idx_nome or idx_dados):
                bold_font = InlineFont(b=True)
                for row in range(2, ws.max_row + 1):
                    cell_guerra = ws.cell(row=row, column=idx_guerra)
                    war_name = str(cell_guerra.value or "").strip()

                    if not war_name:
                        continue

                    if idx_nome:
                        cell_nome = ws.cell(row=row, column=idx_nome)
                        full_name = str(cell_nome.value or "").strip()
                        if full_name and war_name in full_name:
                            parts = full_name.split(war_name)
                            rich = CellRichText()
                            for i, part in enumerate(parts):
                                if part:
                                    rich.append(part)
                                if i < len(parts) - 1:
                                    rich.append(TextBlock(bold_font, war_name))
                            cell_nome.value = rich
                            cell_nome.alignment = openpyxl.styles.Alignment(
                                wrap_text=False
                            )

                    if idx_dados:
                        cell_dados = ws.cell(row=row, column=idx_dados)
                        full_dados = str(cell_dados.value or "").strip()
                        if full_dados and war_name in full_dados:
                            parts = full_dados.split(war_name)
                            rich_dados = CellRichText()
                            for i, part in enumerate(parts):
                                if part:
                                    rich_dados.append(part)
                                if i < len(parts) - 1:
                                    rich_dados.append(TextBlock(bold_font, war_name))
                            cell_dados.value = rich_dados
                            cell_dados.alignment = openpyxl.styles.Alignment(
                                wrap_text=False
                            )

    wb.save(arquivo_saida)
    return arquivo_saida
