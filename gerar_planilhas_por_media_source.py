import os
import re
import zipfile
import shutil
from datetime import datetime
import pandas as pd

# -----------------------
# Utilitários
# -----------------------
def detectar_coluna(colunas, nomes_possiveis):
    normalizadas = {col: col.lower().replace(" ", "").replace("_", "") for col in colunas}
    for original, normalizado in normalizadas.items():
        if normalizado in nomes_possiveis:
            return original
    raise Exception(f"❌ Não encontrei nenhuma coluna correspondente a {nomes_possiveis}. Colunas: {list(colunas)}")

def carregar_arquivo(caminho):
    ext = os.path.splitext(caminho)[1].lower()

    if ext == ".csv":
        try:
            df = pd.read_csv(caminho, sep=";", dtype=str, encoding='utf-8')
            if df.shape[1] == 1:
                df = pd.read_csv(caminho, sep=",", dtype=str, encoding='utf-8')
        except:
            with open(caminho, "rb") as f:
                raw = f.read().decode("utf-8", errors="replace")
            from io import StringIO
            df = pd.read_csv(StringIO(raw), dtype=str)
        return df

    elif ext in [".xlsx", ".xls"]:
        return pd.read_excel(caminho, dtype=str)

    else:
        raise Exception("❌ Formato não suportado. Use CSV ou XLSX.")

def limpar_nome(nome, max_len=50):
    nome = str(nome).strip()
    nome = re.sub(r"[<>:\"/\\|?*]", "_", nome)
    return nome[:max_len]

# -----------------------
# Gerar planilhas por media source
# -----------------------
def gerar_planilhas_por_media_source(df_bruta, pasta_saida):
    col_app = detectar_coluna(df_bruta.columns, {"appname", "app", "application"})
    col_media = detectar_coluna(df_bruta.columns, {"mediasource", "source", "mediasrc"})

    # 🔥 Normaliza colunas para lowercase para não perder medias!
    df_bruta[col_media] = df_bruta[col_media].astype(str).str.lower()
    df_bruta[col_app] = df_bruta[col_app].astype(str).str.lower()

    os.makedirs(pasta_saida, exist_ok=True)
    total = 0

    for app_name in df_bruta[col_app].dropna().unique():

        df_app = df_bruta[df_bruta[col_app] == app_name]

        medias = df_app[col_media].dropna().unique()

        for media in medias:
            df_media = df_app[df_app[col_media] == media]
            if df_media.empty:
                continue

            arquivo = os.path.join(pasta_saida, f"{limpar_nome(media)}.xlsx")
            df_media.to_excel(arquivo, index=False)
            total += 1

    return total

# -----------------------
# Matching AppName inteligente
# -----------------------
def match_apps(df_bruta, df_pub, col_app_bruta, col_app_pub):
    df_bruta[col_app_bruta] = df_bruta[col_app_bruta].astype(str)
    df_pub[col_app_pub] = df_pub[col_app_pub].astype(str)

    df_pub["_app_norm"] = (
        df_pub[col_app_pub]
        .str.normalize("NFKD").str.encode("ascii", "ignore").str.decode("ascii")
        .str.lower()
    )

    correspondencias = []

    for app_raw in df_bruta[col_app_bruta].dropna().unique():

        app_norm = (
            str(app_raw)
            .lower()
            .replace("(", " ")
            .replace(")", " ")
            .replace("-", " ")
        )
        app_norm = re.sub(r"[^a-z0-9 ]", " ", app_norm).strip()
        chave = app_norm.split()[0] if app_norm else ""

        if not chave:
            continue

        encontrados = df_pub[df_pub["_app_norm"].str.contains(chave, na=False)].copy()

        print(f"🔍 App bruto '{app_raw}' → chave '{chave}' → {len(encontrados)} correspondências")

        if not encontrados.empty:
            correspondencias.append((app_raw, encontrados))

    df_pub.drop(columns=["_app_norm"], inplace=True)
    return correspondencias

# -----------------------
# Agrupar por publisher e gerar ZIP
# -----------------------
def agrupar_por_publisher(planilha_mapeamento, df_bruta,
                          pasta_medias, pasta_publishers, pasta_zips):

    df_map = carregar_arquivo(planilha_mapeamento)

    col_app_bruta = detectar_coluna(df_bruta.columns, {"appname", "app", "application"})
    col_app_pub = detectar_coluna(df_map.columns, {"appname", "app", "application"})
    col_media = detectar_coluna(df_map.columns, {"mediasource", "mediasrc", "source"})
    col_pub = detectar_coluna(df_map.columns, {"publisher", "partner", "pub"})

    os.makedirs(pasta_publishers, exist_ok=True)
    os.makedirs(pasta_zips, exist_ok=True)

    # MATCH ENTRE APPS
    matches = match_apps(df_bruta, df_map, col_app_bruta, col_app_pub)

    if not matches:
        print("⚠ Nenhum app correspondente encontrado.")
        return

    # PROCESSAMENTO POR APP
    for app_bruto, df_filtrado in matches:

        print(f"\n==============================")
        print(f"📌 PROCESSANDO APP: {app_bruto}")
        print("==============================")

        publishers = df_filtrado.groupby(col_pub)

        for publisher, grupo in publishers:

            nome_pub = limpar_nome(publisher)
            pasta_pub = os.path.join(pasta_publishers, nome_pub)
            os.makedirs(pasta_pub, exist_ok=True)

            # ---------------------------------------------------------
            # NOVA LÓGICA — VERIFICA MEDIA SOURCE QUE EXISTE NA BRUTA
            # ---------------------------------------------------------
            medias_pub = grupo[col_media].dropna().unique()

            # somente medias realmente existentes
            medias_existentes = [
                m for m in medias_pub
                if os.path.exists(os.path.join(pasta_medias, f"{limpar_nome(m)}.xlsx"))
            ]

            if not medias_existentes:
                print(f"⚠ Publisher {publisher} não possui medias válidas.")
                continue

            # copia apenas medias válidas
            for media in medias_existentes:
                origem = os.path.join(pasta_medias, f"{limpar_nome(media)}.xlsx")
                shutil.copy2(origem, os.path.join(pasta_pub, f"{limpar_nome(media)}.xlsx"))

            # gera ZIP
            zip_path = os.path.join(pasta_zips, f"{nome_pub}.zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                for root, _, files in os.walk(pasta_pub):
                    for file in files:
                        full_path = os.path.join(root, file)
                        arcname = os.path.relpath(full_path, pasta_publishers)
                        zipf.write(full_path, arcname)

            print(f"📦 ZIP GERADO: {zip_path}")

    # APAGA PASTA TEMPORÁRIA
    try:
        shutil.rmtree(pasta_publishers)
        print(f"\n🧹 Pasta temporária apagada: {pasta_publishers}")
    except:
        pass

    print("\n✅ AGRUPAMENTO COMPLETO")

# -----------------------
# Fluxo principal
# -----------------------
if __name__ == "__main__":

    pasta_planilha_bruta = r"C:\Users\edutm\Downloads\caracol\planilhas_deteccao"
    planilha_publisher = r"C:\Users\edutm\Downloads\caracol\PID e Publisher - Ofertas.xlsx"
    caminho_saida = r"C:\Users\edutm\Downloads\caracol\saidas"

    data_formatada = datetime.now().strftime("%d-%m-%Y")
    output_date_folder = os.path.join(caminho_saida, data_formatada)
    os.makedirs(output_date_folder, exist_ok=True)

    pasta_media_sources = os.path.join(output_date_folder, "media_sources")
    pasta_publishers = os.path.join(output_date_folder, "publishers_temp")
    pasta_publishers_zips = os.path.join(output_date_folder, "publishers_zips")

    arquivos = [a for a in os.listdir(pasta_planilha_bruta) if a.lower().endswith((".csv", ".xlsx", ".xls"))]

    if not arquivos:
        print("❌ Nenhum arquivo encontrado.")
        raise SystemExit(1)

    df_bruta_total = []
    total = 0

    # CARREGA TODAS AS PLANILHAS BRUTAS
    for arq in arquivos:
        caminho = os.path.join(pasta_planilha_bruta, arq)

        try:
            df_tmp = carregar_arquivo(caminho)
            df_bruta_total.append(df_tmp)

            # gera planilhas por media source
            n = gerar_planilhas_por_media_source(df_tmp, pasta_media_sources)
            print(f"✔ {arq} → {n} medias geradas")
            total += n

        except Exception as e:
            print(f"❌ Erro ao processar {arq}: {e}")

    if total == 0:
        print("⚠ Nenhuma media gerada.")
        raise SystemExit(1)

    # UNE TODAS AS BRUTAS EM UMA SÓ
    df_bruta_final = pd.concat(df_bruta_total, ignore_index=True)

    # AGRUPA POR PUBLISHER E GERA ZIP
    agrupar_por_publisher(planilha_publisher, df_bruta_final,
                          pasta_media_sources, pasta_publishers, pasta_publishers_zips)

    print("\n🎯 PROCESSO FINALIZADO")
    print(f"📂 Pasta de saída: {output_date_folder}")
    print(f"📦 Zips gerados em: {pasta_publishers_zips}")
