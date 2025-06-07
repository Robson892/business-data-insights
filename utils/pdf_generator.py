from fpdf import FPDF
import tempfile
import os
from io import BytesIO
import matplotlib.pyplot as plt
import pandas as pd

def gerar_pdf(df, colunas_selecionadas, tipo_grafico, fig, coluna_num, total_registros):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Cabeçalho
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Relatório Automático", ln=True, align="C")

    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, f"Total de registros: {total_registros}", ln=True)

    # Estatísticas coluna numérica
    if coluna_num:
        media = df[coluna_num].mean()
        mediana = df[coluna_num].median()
        std = df[coluna_num].std()
        minimo = df[coluna_num].min()
        maximo = df[coluna_num].max()
        pdf.cell(0, 8, f"Resumo estatístico da coluna {coluna_num}:", ln=True)
        pdf.cell(0, 8, f"Média: {media:.2f} | Mediana: {mediana:.2f} | Desvio Padrão: {std:.2f}", ln=True)
        pdf.cell(0, 8, f"Mínimo: {minimo:.2f} | Máximo: {maximo:.2f}", ln=True)

    # Gerar gráfico como imagem temporária
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
        fig.savefig(tmpfile.name, bbox_inches='tight')
        pdf.image(tmpfile.name, x=10, w=pdf.w - 20)
        tmpfile.close()
        os.unlink(tmpfile.name)

    # Tabela simplificada - primeiras 10 linhas das colunas selecionadas
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "Tabela de dados (amostra):", ln=True)

    # Limitar a 10 linhas e colunas selecionadas para tabela no PDF
    tabela = df[colunas_selecionadas].head(10) if colunas_selecionadas else df.head(10)
    pdf.set_font("Arial", size=10)
    col_width = (pdf.w - 20) / max(len(tabela.columns), 1)
    row_height = 7

    # Cabeçalho da tabela
    for col in tabela.columns:
        pdf.cell(col_width, row_height, str(col), border=1)
    pdf.ln(row_height)

    # Linhas da tabela
    for i, row in tabela.iterrows():
        for col in tabela.columns:
            texto = str(row[col])[:15]
            pdf.cell(col_width, row_height, texto, border=1)
        pdf.ln(row_height)

    # Salvar PDF
    pdf.output("grafico.pdf")
