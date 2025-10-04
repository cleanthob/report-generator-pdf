import pandas as pd
import matplotlib.pyplot as plt
import io
import base64
from datetime import datetime
import pdfkit


def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def gerar_grafico_horizontal(serie, titulo, mostrar_valor_moeda=False):
    plt.figure(figsize=(6, 4))
    ax = serie.plot(kind="barh", color="#1f77b4")
    ax.set_title(titulo)
    plt.tight_layout()
    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.set_xticks([])
    ax.tick_params(axis="y", length=0)
    for spine in ax.spines.values():
        spine.set_visible(False)
    for p in ax.patches:
        espaco = p.get_width() * 0.01
        valor = (
            formatar_moeda(p.get_width())
            if mostrar_valor_moeda
            else f"{int(p.get_width())}"
        )
        ax.annotate(
            valor,
            (p.get_width() + espaco, p.get_y() + p.get_height() / 2.0),
            ha="left",
            va="center",
            fontsize=12,
            color="black",
        )
    buf = io.BytesIO()
    plt.savefig(buf, format="png", bbox_inches="tight")
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    plt.close()
    return img_base64


def gerar_relatorio_html(df):
    colunas_necessarias = {"Departamento", "Cargo", "Salário"}
    if not colunas_necessarias.issubset(df.columns):
        raise ValueError(
            f"O arquivo precisa conter as colunas: {', '.join(colunas_necessarias)}"
        )

    # Estatísticas principais
    total_funcionarios = len(df)
    media_salarial = formatar_moeda(df["Salário"].mean())
    maior_salario = formatar_moeda(df["Salário"].max())
    menor_salario = formatar_moeda(df["Salário"].min())
    total_folha = formatar_moeda(df["Salário"].sum())

    # Gráficos em base64
    imgs = []

    # Média salarial por departamento
    media_salario_dept = df.groupby("Departamento")["Salário"].mean().sort_values()
    imgs.append(
        gerar_grafico_horizontal(
            media_salario_dept, "Média Salarial por Departamento", True
        )
    )

    # Funcionários por cargo
    qtd_cargo = df["Cargo"].value_counts().sort_values()
    imgs.append(gerar_grafico_horizontal(qtd_cargo, "Funcionários por Cargo"))

    # Média salarial por cargo
    media_salario_cargo = df.groupby("Cargo")["Salário"].mean().sort_values()
    imgs.append(
        gerar_grafico_horizontal(media_salario_cargo, "Média Salarial por Cargo", True)
    )

    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 30px;
            }}
            h1 {{
                text-align: center;
                color: #003366;
            }}
            .stats, .analises{{
                margin-bottom: 20px;
                background: #f2f2f2;
                padding: 10px;
                border-radius: 8px;
            }}
            .grafico {{
                text-align: center;
                margin: 25px 0;
            }}
            img {{
                width: 500px;
                height: auto;
            }}
        </style>
    </head>
    <body>
        <h1>Relatório da Empresa XYZ</h1>
        <p><b>Data de geração:</b> {datetime.now().strftime("%d/%m/%Y %H:%M")}</p>

        <div class="stats">
            <p><b>Número de funcionários:</b> {total_funcionarios}</p>
            <p><b>Média salarial:</b> {media_salarial}</p>
            <p><b>Maior salário:</b> {maior_salario}</p>
            <p><b>Menor salário:</b> {menor_salario}</p>
            <p><b>Total da folha:</b> {total_folha}</p>
        </div>

        <div class="grafico">
            <img src="data:image/png;base64,{imgs[0]}" />
        </div>

        <div class="grafico">
            <img src="data:image/png;base64,{imgs[1]}" />
        </div>
        
        <div class="grafico">
            <br>
            <img src="data:image/png;base64,{imgs[2]}" />
        </div>

        <div class="analises">
            <h2>Análise Geral</h2>
            <p>O relatório apresenta informações detalhadas sobre os funcionários da empresa XYZ, incluindo número total de colaboradores, distribuição de salários, cargos e departamentos. Observa-se a variação salarial entre os diferentes setores e cargos, permitindo identificar áreas com maior investimento ou necessidade de ajustes.</p>

            <h2>Média Salarial por Departamento</h2>
            <p>A média salarial por departamento mostra que alguns setores possuem remuneração significativamente acima ou abaixo da média da empresa. Essa análise auxilia a identificar potenciais desigualdades e possibilita decisões estratégicas de valorização ou realocação de recursos.</p>

            <h2>Distribuição de Funcionários por Cargo</h2>
            <p>O gráfico de distribuição de funcionários por cargo evidencia quais posições possuem maior concentração de colaboradores. Isso é útil para planejamento de recursos humanos, identificando cargos críticos com necessidade de expansão ou otimização.</p>

            <h2>Média Salarial por Cargo</h2>
            <p>A análise da média salarial por cargo permite comparar a remuneração de diferentes funções dentro da empresa. Cargos com salários mais altos podem refletir maior especialização ou responsabilidade, enquanto cargos com média salarial mais baixa podem indicar oportunidades de crescimento ou necessidade de revisão.</p>

            <h2>Conclusão</h2>
            <p>Essas análises oferecem uma visão completa da estrutura salarial e organizacional da empresa. O acompanhamento contínuo dessas métricas ajuda na tomada de decisões estratégicas, promovendo equidade interna e eficiência na gestão de pessoas.</p>
        </div>
    </body>
    </html>
    """

    return html


def salvar_pdf(html, config, nome_arquivo="relatorio.pdf"):
    pdfkit.from_string(html, nome_arquivo, configuration=config)
    print(f"✅ PDF salvo como {nome_arquivo}")


if __name__ == "__main__":

    # Caminho completo do executável wkhtmltopdf
    path_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

    df = pd.read_excel("empresa_dados.xlsx")

    # Gerar HTML e PDF
    html = gerar_relatorio_html(df)
    salvar_pdf(html, config, "relatorio.pdf")
