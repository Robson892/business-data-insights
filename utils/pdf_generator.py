from fpdf import FPDF
import os
import tempfile
from datetime import datetime

class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 14)
        self.cell(0, 10, "Relatório Analítico de Dados", ln=True, align="C")
        self.set_font("Arial", "", 10)
        self.cell(0, 10, f"Tipo de Gráfico: {self.tipo_grafico}", ln=True, align="C")
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", align="C")

    def chapter_body(self, df, colunas, col_num, stats, total_registros, registros_filtrados):
        self.set_font("Arial", "", 10)

        # Resumo do filtro
        self.cell(0, 10, f"Total de registros: {total_registros}", ln=True)
        self.cell(0, 10, f"Registros após filtros: {registros_filtrados}", ln=True)
        self.ln(3)

        # Estatísticas
        self.set_font("Arial", "B", 11)
        self.cell(0, 10, f"Resumo Estatístico da Coluna: {col_num}", ln=True)
        self.set_font("Arial", "", 10)
        for k, v in stats.items():
            self.cell(0, 8, f"{k}: {v:.2f}", ln=True)

        self.ln(5)
        self.set_font("Arial", "B", 11)
        self.cell(0, 10, "Prévia dos dados (até 20 linhas):", ln=True)
        self.set_font("Arial", "", 10)
        self.ln(2)
        for index, row in df[colunas].head(20).iterrows():
            texto = " | ".join(f"{k}: {v}" for k, v in row.items())
            self.multi_cell(0, 8, texto)
            self.ln(1)

def gerar_pdf(df, colunas, tipo_grafico, fig, col_num, total_registros):
    pdf = PDF()
    pdf.tipo_grafico = tipo_grafico
    pdf.add_page()

    # Resumo estatístico
    stats = {
        "Média": df[col_num].mean(),
        "Mediana": df[col_num].median(),
        "Desvio Padrão": df[col_num].std(),
        "Mínimo": df[col_num].min(),
        "Máximo": df[col_num].max()
    }

    pdf.chapter_body(
        df=df,
        colunas=colunas,
        col_num=col_num,
        stats=stats,
        total_registros=total_registros,
        registros_filtrados=len(df)
    )

    # Salvar e inserir gráfico
    tmpfile = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    tmpfile.close()
    fig.savefig(tmpfile.name)

    pdf.ln(5)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Gráfico Gerado:", ln=True)
    pdf.image(tmpfile.name, x=10, w=pdf.w - 20)
    os.remove(tmpfile.name)

    pdf.output("grafico.pdf")
