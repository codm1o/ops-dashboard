"""
gerar_dados.py
--------------
Script de processamento de dados para o Operational Analytics Dashboard.

Uso:
    python gerar_dados.py <arquivo_excel> <mes> <label>

Argumentos:
    arquivo_excel   Caminho para o arquivo .xlsx com os dados do período.
    mes             Chave do mês em minúsculas sem acento (ex: fevereiro, marco, abril).
    label           Rótulo legível para exibição no dashboard (ex: "Fevereiro 2026").

Exemplo:
    python gerar_dados.py dados_fevereiro.xlsx fevereiro "Fevereiro 2026"

Estrutura esperada da planilha:
    - Data fatal       : data limite do prazo (formato DD/MM/AAAA)
    - Data conclusão   : data em que a atividade foi concluída
    - Prazo            : valor inteiro indicando antecedência (ex: D0, D-1, D-2...)
    - Evento           : "Principal" para atividades principais
    - Tipo             : "Prazo" ou "Tarefa"
    - Área             : Cível, Trabalhista ou Tributário
    - Responsável efetivo : nome do profissional
    - Cargo            : Advogado, Coordenador ou Estagiário
    - Workflow         : nome do tipo de atividade
    - Auditoria        : "Sem justificativa" quando aplicável

Saída:
    Atualiza o arquivo dados.json (criado se não existir), adicionando ou
    substituindo os dados do mês informado.
"""

import pandas as pd
import json
import sys
import os


def classificar_prazo(valor: str) -> str | None:
    """
    Converte o campo Prazo (ex: '0D', '-1D', '-3D') para a faixa correspondente.

    Faixas:
        D0    -> concluído no dia do prazo
        D-1   -> concluído 1 dia antes
        D-2   -> concluído 2 dias antes
        D-3+  -> concluído 3 ou mais dias antes
        D+    -> concluído após a data fatal
    """
    try:
        n = int(str(valor).replace("D", "").strip())
        if n == 0:
            return "D0"
        elif n == -1:
            return "D-1"
        elif n == -2:
            return "D-2"
        elif n <= -3:
            return "D-3+"
        elif n > 0:
            return "D+"
    except (ValueError, TypeError):
        return None


def percentual(parte: int, total: int) -> str:
    if total == 0:
        return "0%"
    return f"{round((parte / total) * 100, 1)}%".replace(".", ",")


def contar_area(frame: pd.DataFrame, area: str) -> int:
    return len(frame[frame["area"] == area])


def processar(arquivo: str, mes: str, label: str) -> dict:
    df = pd.read_excel(arquivo, sheet_name=0, engine="openpyxl")
    df.columns = df.columns.str.strip()

    # Normalização de datas
    df["Data fatal"] = pd.to_datetime(df["Data fatal"], errors="coerce", dayfirst=True)
    df["Data conclusão"] = pd.to_datetime(df["Data conclusão"], errors="coerce", dayfirst=True)

    # Apenas atividades concluídas
    df = df[df["Data conclusão"].notna()]

    # Campo de prazo como string limpa
    df["faixa"] = df["Prazo"].astype(str).str.strip()

    # Normalização da coluna Área
    df["area"] = df["Área"].fillna("").astype(str).str.lower().str.strip()
    df["area"] = df["area"].replace({
        "civel":       "Cível",
        "cível":       "Cível",
        "trabalhista": "Trabalhista",
        "tributario":  "Tributário",
        "tributário":  "Tributário",
    })

    # Separação por tipo de evento e atividade
    df_principal = df[df["Evento"] == "Principal"]
    df_prazo     = df_principal[df_principal["Tipo"] == "Prazo"].copy()
    df_tarefa    = df_principal[df_principal["Tipo"] != "Prazo"]

    df_prazo["faixa_group"] = df_prazo["faixa"].apply(classificar_prazo)

    # KPIs globais
    total    = len(df)
    prazos   = len(df_prazo)
    tarefas  = len(df_tarefa)
    etapas   = total - (tarefas + prazos)

    # Distribuição de workflows
    workflow_dist = df["Workflow"].value_counts().to_dict()

    # Breakdown de prazo
    d0  = len(df_prazo[df_prazo["faixa_group"] == "D0"])
    dm1 = len(df_prazo[df_prazo["faixa_group"] == "D-1"])
    dm2 = len(df_prazo[df_prazo["faixa_group"] == "D-2"])
    dm3 = len(df_prazo[df_prazo["faixa_group"] == "D-3+"])
    dp  = len(df_prazo[df_prazo["faixa_group"] == "D+"])

    # Sem justificativa por cargo
    sj      = df_prazo[df_prazo["Auditoria"] == "Sem justificativa"]
    sj_adv  = sj[sj["Cargo"] == "Advogado"]
    sj_est  = sj[sj["Cargo"] == "Estagiário"]

    # Distribuição por área
    area_trab = len(df_prazo[df_prazo["area"] == "Trabalhista"])
    area_civ  = len(df_prazo[df_prazo["area"] == "Cível"])
    area_trib = len(df_prazo[df_prazo["area"] == "Tributário"])

    # Dados individuais por profissional
    prof_data = []

    for nome, grupo_full in df.groupby("Responsável efetivo"):
        grupo_principal = grupo_full[grupo_full["Evento"] == "Principal"]
        grupo_prazo     = grupo_principal[grupo_principal["Tipo"] == "Prazo"].copy()
        tarefas_p       = len(grupo_principal[grupo_principal["Tipo"] != "Prazo"])
        etapas_p        = len(grupo_full[grupo_full["Evento"] != "Principal"])
        prazos_count    = len(grupo_prazo)

        grupo_prazo["faixa_group"] = grupo_prazo["faixa"].apply(classificar_prazo)

        d0_p  = len(grupo_prazo[grupo_prazo["faixa_group"] == "D0"])
        dm1_p = len(grupo_prazo[grupo_prazo["faixa_group"] == "D-1"])
        dm2_p = len(grupo_prazo[grupo_prazo["faixa_group"] == "D-2"])
        dm3_p = len(grupo_prazo[grupo_prazo["faixa_group"] == "D-3+"])
        dp_p  = len(grupo_prazo[grupo_prazo["faixa_group"] == "D+"])

        sem_just = len(grupo_prazo[grupo_prazo["Auditoria"] == "Sem justificativa"])

        cargo = grupo_full["Cargo"].dropna().iloc[0] if not grupo_full["Cargo"].dropna().empty else ""
        area  = grupo_full["area"].dropna().iloc[0]  if not grupo_full["area"].dropna().empty  else ""

        workflow_list = grupo_full["Workflow"].dropna().tolist()
        total_p       = tarefas_p + prazos_count + etapas_p

        prof_data.append({
            "name":             nome,
            "cargo":            cargo,
            "area":             area,
            "tarefas":          tarefas_p,
            "prazos":           prazos_count,
            "etapas":           etapas_p,
            "workflow_list":    workflow_list,
            "total":            total_p,
            "d0":               d0_p,
            "dm1":              dm1_p,
            "dm2":              dm2_p,
            "dm3":              dm3_p,
            "dp":               dp_p,
            "sem_justificativa": sem_just,
        })

    return {
        "label":           label,
        "available":       True,
        "workflowDist":    workflow_dist,
        "totalConcluidos": total,
        "tarefas":         tarefas,
        "prazos":          prazos,
        "etapas":          etapas,
        "d0":              d0,
        "dm1":             dm1,
        "dm2":             dm2,
        "dm3":             dm3,
        "dp":              dp,
        "areaTrabalhista": area_trab,
        "areaCivel":       area_civ,
        "areaTributario":  area_trib,
        "sjAdv": {
            "total":       len(sj_adv),
            "civel":       contar_area(sj_adv, "Cível"),
            "trabalhista": contar_area(sj_adv, "Trabalhista"),
            "tributario":  contar_area(sj_adv, "Tributário"),
        },
        "sjEst": {
            "total":       len(sj_est),
            "civel":       contar_area(sj_est, "Cível"),
            "trabalhista": contar_area(sj_est, "Trabalhista"),
            "tributario":  contar_area(sj_est, "Tributário"),
        },
        "profData": prof_data,
    }


def main():
    if len(sys.argv) != 4:
        print("Uso: python gerar_dados.py <arquivo_excel> <mes> <label>")
        print('Exemplo: python gerar_dados.py dados_abril.xlsx abril "Abril 2026"')
        sys.exit(1)

    arquivo = sys.argv[1]
    mes     = sys.argv[2]
    label   = sys.argv[3]

    if not os.path.exists(arquivo):
        print(f"Arquivo não encontrado: {arquivo}")
        sys.exit(1)

    print(f"Processando {label} ({mes})...")
    dados_mes = processar(arquivo, mes, label)

    # Carrega JSON existente ou cria novo
    caminho_json = "dados.json"
    if os.path.exists(caminho_json):
        with open(caminho_json, "r", encoding="utf-8") as f:
            dados_existentes = json.load(f)
    else:
        dados_existentes = {}

    dados_existentes[mes] = dados_mes

    with open(caminho_json, "w", encoding="utf-8") as f:
        json.dump(dados_existentes, f, ensure_ascii=False, indent=2)

    print(f"dados.json atualizado com sucesso para o periodo: {label}")
    print(f"  Total concluidos : {dados_mes['totalConcluidos']}")
    print(f"  Prazos           : {dados_mes['prazos']}")
    print(f"  Tarefas          : {dados_mes['tarefas']}")
    print(f"  Etapas           : {dados_mes['etapas']}")
    print(f"  D+               : {dados_mes['dp']}")


if __name__ == "__main__":
    main()
