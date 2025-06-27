import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io

st.set_page_config(page_title="Análise de Licenças Microsoft", layout="wide")
st.title("🔍 Analisador de Licenças Microsoft")

# Upload do arquivo de usuários
st.sidebar.header("📁 Upload de Arquivo")
usuarios_file = st.sidebar.file_uploader("Arquivo de Usuários (.csv ou .xlsx)", type=["csv", "xlsx"])

def carregar_arquivo(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        return pd.read_csv(uploaded_file, sep=None, engine="python")
    elif uploaded_file.name.endswith(".xlsx"):
        return pd.read_excel(uploaded_file)
    else:
        st.error("Formato de arquivo não suportado.")
        return None

# Licenças pagas
licencas_pagas = [
    "Business Basic",
    "Business Standard",
    "Business Premium",
    "Power BI Pro",
    "Exchange Online"
]

def possui_licenca_paga(lic_str):
    if pd.isna(lic_str) or lic_str.strip() == "Unlicensed":
        return False
    licencas = [l.strip() for l in lic_str.split("+")]
    return any(any(chave.lower() in l.lower() for chave in licencas_pagas) for l in licencas)

if usuarios_file:
    df_usuarios = carregar_arquivo(usuarios_file)
    df_usuarios.rename(columns={"\ufeffDisplay name": "Display name"}, inplace=True)

    required_columns = ["User principal name", "Licenses", "Block credential"]
    if not all(col in df_usuarios.columns for col in required_columns):
        st.error("Arquivo está faltando colunas essenciais.")
        st.stop()

    df_usuarios["É Externo"] = df_usuarios["User principal name"].str.contains("#EXT#", na=False)
    df_internos = df_usuarios[~df_usuarios["É Externo"]].copy()
    df_externos = df_usuarios[df_usuarios["É Externo"]].copy()

    df_internos["Empresa"] = df_internos["User principal name"].str.extract(r'@([\w.-]+)$')[0].str.lower()
    df_internos["Possui Licença Paga"] = df_internos["Licenses"].apply(possui_licenca_paga)

    empresas_disponiveis = sorted(df_internos["Empresa"].dropna().unique())
    empresas_selecionadas = st.sidebar.multiselect("Filtrar por Empresa(s)", options=empresas_disponiveis, default=empresas_disponiveis)

    df_filtrado = df_internos[df_internos["Empresa"].isin(empresas_selecionadas)]

    df_internos["Block_cred_normalizado"] = (
        df_internos["Block credential"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    valores_bloqueio = ["VERDADEIRO", "TRUE", "1"]
    df_internos["Bloqueado"] = df_internos["Block_cred_normalizado"].isin(valores_bloqueio)

    resumo_empresa = df_internos.groupby("Empresa").agg(
        Total_Usuarios=("User principal name", "count"),
        Bloqueados=("Bloqueado", "sum")
    )
    resumo_empresa["Ativos"] = resumo_empresa["Total_Usuarios"] - resumo_empresa["Bloqueados"]
    resumo_empresa = resumo_empresa[["Total_Usuarios", "Ativos", "Bloqueados"]]

    st.subheader("📊 Quadro Resumo por Empresa (Ativos x Bloqueados)")
    st.dataframe(resumo_empresa)

    df_licencas_pagas = df_filtrado[df_filtrado["Possui Licença Paga"]].copy()
    df_licencas_pagas["Block_cred_normalizado"] = (
        df_licencas_pagas["Block credential"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    df_licencas_pagas["Bloqueado"] = df_licencas_pagas["Block_cred_normalizado"].isin(valores_bloqueio)

    df_licencas_pagas["Licenças Pagas"] = df_licencas_pagas["Licenses"].str.split("+").apply(
        lambda licencas: "+".join([
            lic.strip() for lic in licencas if any(p.lower() in lic.lower() for p in licencas_pagas)
        ]) if isinstance(licencas, list) else ""
    )

    st.subheader("🧾 Usuários com Licenças Pagas")
    st.dataframe(df_licencas_pagas[["Display name", "Licenças Pagas"]].sort_values("Display name"))

    df_bloqueados_licenca = df_licencas_pagas[df_licencas_pagas["Bloqueado"] == True]

    st.subheader("🔒 Usuários Bloqueados com Licenças Pagas")
    st.dataframe(df_bloqueados_licenca[["Display name", "Licenças Pagas"]].sort_values("Display name"))

    st.subheader("📌 Gráfico de Distribuição de Licenças Pagas")
    licencas_explodidas = df_filtrado["Licenses"].dropna().str.split("+").explode().str.strip()
    licencas_filtradas = [lic for lic in licencas_explodidas if any(chave.lower() in lic.lower() for chave in licencas_pagas)]

    licencas_unicas = sorted(set(licencas_filtradas))
    licencas_selecionadas = st.multiselect("Filtrar por tipo de licença", options=licencas_unicas, default=licencas_unicas)

    licencas_filtradas_df = pd.Series(licencas_filtradas)
    licencas_filtradas_df = licencas_filtradas_df[licencas_filtradas_df.isin(licencas_selecionadas)]
    st.bar_chart(licencas_filtradas_df.value_counts())

    df_expandidas = df_internos[df_internos["Possui Licença Paga"]].copy()
    df_expandidas = df_expandidas[["Empresa", "Licenses", "User principal name"]].dropna()
    df_expandidas["Licenca Individual"] = df_expandidas["Licenses"].str.split("+")
    df_expandidas = df_expandidas.explode("Licenca Individual")
    df_expandidas["Licenca Individual"] = df_expandidas["Licenca Individual"].str.strip()

    def classificar_licenca(lic):
        lic = lic.lower()
        if "business basic" in lic:
            return "Microsoft 365 Business Basic"
        elif "business standard" in lic:
            return "Microsoft 365 Business Standard"
        elif "business premium" in lic:
            return "Microsoft 365 Business Premium"
        elif "power bi pro" in lic:
            return "Power BI Pro"
        elif "exchange online" in lic:
            return "Exchange Online"
        return None

    df_expandidas["Licenca Classificada"] = df_expandidas["Licenca Individual"].apply(classificar_licenca)
    df_lic_pagas = df_expandidas[df_expandidas["Licenca Classificada"].notna()]

    df_pivot = df_lic_pagas.pivot_table(
        index="Empresa",
        columns="Licenca Classificada",
        values="User principal name",
        aggfunc="count",
        fill_value=0
    ).reset_index()

    df_pivot["Total com Licença"] = df_pivot.drop(columns="Empresa").sum(axis=1)

    df_pivot = df_pivot.rename(columns={
        "Microsoft 365 Business Basic": "365 Basic",
        "Microsoft 365 Business Standard": "365 Standard",
        "Microsoft 365 Business Premium": "365 Premium",
        "Power BI Pro": "Power BI Pro",
        "Exchange Online": "Exchange Online",
    })

    ordem = ["Empresa", "365 Basic", "365 Standard", "365 Premium", "Power BI Pro", "Exchange Online", "Total com Licença"]
    df_pivot = df_pivot[[col for col in ordem if col in df_pivot.columns]]

    st.subheader("🏢 Licenças Pagas por Empresa (Formato por Coluna)")
    st.dataframe(df_pivot)

    st.subheader("🥧 Distribuição de Usuários com Licença Paga por Empresa")
    lic_por_empresa = df_internos[df_internos["Possui Licença Paga"]].groupby("Empresa")["User principal name"].count()
    if not lic_por_empresa.empty:
        fig, ax = plt.subplots()
        ax.pie(lic_por_empresa, labels=lic_por_empresa.index, autopct="%1.1f%%")
        ax.axis("equal")
        st.pyplot(fig)
    else:
        st.info("Nenhum usuário com licença paga encontrado.")

    st.subheader("❌ Usuários Sem Licenças Pagas")
    st.dataframe(df_filtrado[~df_filtrado["Possui Licença Paga"]])

    st.subheader("🌐 Análise de Usuários Externos")
    st.write(f"Total de usuários externos: {len(df_externos)}")
    st.dataframe(df_externos)

    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        df_licencas_pagas.to_excel(writer, sheet_name="Licencas_Pagas", index=False)
        df_bloqueados_licenca.to_excel(writer, sheet_name="Bloqueados_Licencas", index=False)
        df_pivot.to_excel(writer, sheet_name="Resumo_Empresa", index=False)
        resumo_empresa.to_excel(writer, sheet_name="Resumo_Bloqueio", index=True)
    st.download_button("⬇️ Baixar Relatório Consolidado", output_excel.getvalue(), file_name="analise_licencas.xlsx")

else:
    st.info("⚠️ Faça upload do arquivo de usuários para iniciar a análise.")
