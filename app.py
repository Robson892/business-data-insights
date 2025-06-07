import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from utils.pdf_generator import gerar_pdf
from utils.email_sender import enviar_email

st.set_page_config(page_title="Automação com Excel", layout="wide")
st.title("📉 Business Data Insights – Excel Edition")

arquivo = st.file_uploader("Faça upload de um arquivo .xlsx", type=["xlsx"])

if arquivo:
    df = pd.read_excel(arquivo)

    # Força a conversão para datetime apenas em colunas com "data" no nome
    for col in df.columns:
        if "data" in col.lower():
            try:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.date
            except Exception as e:
                st.warning(f"Erro ao converter a coluna {col} para data: {e}")

    st.subheader("📋 Dados Carregados:")
    st.dataframe(df)

    # FILTROS
    st.sidebar.header("🔍 Filtros Dinâmicos")

    # Filtro por data - separado início e fim
    col_data = next((col for col in df.columns if "data" in col.lower()), None)
    if col_data:
        if df[col_data].notna().any():
            min_data = df[col_data].min()
            max_data = df[col_data].max()

            data_inicio = st.sidebar.date_input("Data início:", min_value=min_data, max_value=max_data, value=min_data)
            data_fim = st.sidebar.date_input("Data fim:", min_value=min_data, max_value=max_data, value=max_data)

            if data_inicio and data_fim:
                if data_inicio > data_fim:
                    st.sidebar.warning("⚠️ A data início não pode ser maior que a data fim.")
                else:
                    df = df[(df[col_data] >= data_inicio) & (df[col_data] <= data_fim)]
        else:
            st.sidebar.warning("⚠️ Nenhuma data válida encontrada.")

    # Filtro por categoria - excluindo colunas que contenham 'data'
    colunas_obj = [col for col in df.select_dtypes(include='object').columns if "data" not in col.lower()]
    if colunas_obj:
        col_categorica = st.sidebar.selectbox("Coluna categórica:", colunas_obj)
        valores_cat = st.sidebar.multiselect("Valores da categoria:", df[col_categorica].unique(), default=list(df[col_categorica].unique()))
        df = df[df[col_categorica].isin(valores_cat)]

    st.subheader("📊 Dados Filtrados:")
    st.dataframe(df)

    # Seleção da coluna numérica para Resumo Estatístico e Outliers (DEFINIDA AQUI)
    col_num = None
    colunas_numericas = df.select_dtypes(include='number').columns.tolist()
    if colunas_numericas:
        col_num = st.selectbox("Selecione uma coluna numérica para análise:", colunas_numericas)
    else:
        st.warning("⚠️ Nenhuma coluna numérica disponível para análise estatística.")

    # GRÁFICO
    colunas_validas = df.select_dtypes(exclude=["datetime", "timedelta"]).columns.tolist()
    exibir_grafico = st.radio("Deseja exibir um gráfico?", ["Sim", "Não"])

    if exibir_grafico == "Sim" and colunas_validas:
        colunas_selecionadas = st.multiselect("Selecione até 2 colunas para o gráfico:", colunas_validas)
        tipo_grafico = st.selectbox("Selecione o tipo de gráfico:", [
            "Gráfico de Barras", "Gráfico de Linha", "Gráfico de Pizza",
            "Histograma", "Dispersão (Scatter)", "Boxplot"
        ])

        if 1 <= len(colunas_selecionadas) <= 2:
            st.subheader(f"📈 Visualização: {tipo_grafico}")
            fig, ax = plt.subplots(figsize=(8, 4))
            try:
                if tipo_grafico == "Gráfico de Barras":
                    if len(colunas_selecionadas) == 1:
                        df[colunas_selecionadas[0]].value_counts().plot(kind="bar", ax=ax)
                    else:
                        if pd.api.types.is_numeric_dtype(df[colunas_selecionadas[1]]):
                            df.groupby(colunas_selecionadas[0])[colunas_selecionadas[1]].sum().plot(kind="bar", ax=ax)
                        else:
                            st.error("Para barras com duas colunas, a segunda deve ser numérica.")
                            st.stop()

                elif tipo_grafico == "Gráfico de Linha":
                    if all(pd.api.types.is_numeric_dtype(df[col]) for col in colunas_selecionadas):
                        df[colunas_selecionadas].plot(kind="line", ax=ax)
                    else:
                        st.error("Para gráfico de linha, selecione apenas colunas numéricas.")
                        st.stop()

                elif tipo_grafico == "Gráfico de Pizza":
                    if len(colunas_selecionadas) == 1:
                        df[colunas_selecionadas[0]].value_counts().plot(kind="pie", ax=ax, autopct='%1.1f%%')
                        ax.set_ylabel("")
                    else:
                        st.error("Gráfico de pizza só aceita uma coluna categórica.")
                        st.stop()

                elif tipo_grafico == "Histograma":
                    if pd.api.types.is_numeric_dtype(df[colunas_selecionadas[0]]):
                        df[colunas_selecionadas[0]].plot(kind="hist", bins=20, ax=ax)
                    else:
                        st.error("Para histograma, selecione uma coluna numérica.")
                        st.stop()

                elif tipo_grafico == "Dispersão (Scatter)":
                    if len(colunas_selecionadas) == 2 and all(pd.api.types.is_numeric_dtype(df[col]) for col in colunas_selecionadas):
                        df.plot(kind="scatter", x=colunas_selecionadas[0], y=colunas_selecionadas[1], ax=ax)
                    else:
                        st.error("Para gráfico de dispersão, selecione duas colunas numéricas.")
                        st.stop()

                elif tipo_grafico == "Boxplot":
                    if pd.api.types.is_numeric_dtype(df[colunas_selecionadas[0]]):
                        df.boxplot(column=colunas_selecionadas[0], ax=ax)
                    else:
                        st.error("Para boxplot, selecione uma coluna numérica.")
                        st.stop()

                st.pyplot(fig)

                if st.button("📥 Baixar Gráfico e Tabela em PDF"):
                    total_registros = len(pd.read_excel(arquivo))
                    gerar_pdf(df, colunas_selecionadas, tipo_grafico, fig, col_num, total_registros)

            except Exception as e:
                st.error(f"Erro ao gerar o gráfico: {e}")

    # RESUMO ESTATÍSTICO
    st.subheader("📊 Resumo Estatístico")
    if col_num:
        col1, col2, col3 = st.columns(3)
        col1.metric("Média", f"{df[col_num].mean():.2f}")
        col2.metric("Mediana", f"{df[col_num].median():.2f}")
        col3.metric("Desvio Padrão", f"{df[col_num].std():.2f}")

        col4, col5 = st.columns(2)
        col4.metric("Mínimo", f"{df[col_num].min():.2f}")
        col5.metric("Máximo", f"{df[col_num].max():.2f}")

    # OUTLIERS
    st.subheader("🚨 Outliers Detectados")
    if col_num:
        q1 = df[col_num].quantile(0.25)
        q3 = df[col_num].quantile(0.75)
        iqr = q3 - q1
        lim_inf = q1 - 1.5 * iqr
        lim_sup = q3 + 1.5 * iqr
        outliers = df[(df[col_num] < lim_inf) | (df[col_num] > lim_sup)]

        if not outliers.empty:
            st.warning(f"{len(outliers)} outlier(s) detectado(s):")
            st.dataframe(outliers)
        else:
            st.info("Nenhum outlier detectado.")

# FORMULÁRIO DE E-MAIL
with st.form("form_email"):
    st.subheader("📧 Enviar Relatório por E-mail")
    destinatario = st.text_input("E-mail do destinatário")
    remetente = st.text_input("Seu e-mail (remetente)")
    senha = st.text_input("Senha do e-mail", type="password")
    enviar = st.form_submit_button("📨 Enviar relatório")

if enviar:
    if destinatario and remetente and senha:
        sucesso, mensagem = enviar_email(destinatario, "Relatório", "Segue o relatório em anexo.", "grafico.pdf", remetente, senha)
        st.success(mensagem) if sucesso else st.error(mensagem)
    else:
        st.warning("⚠️ Preencha todos os campos.")
