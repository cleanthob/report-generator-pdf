from faker import Faker
import pandas as pd
import random


def gerar_planilha_empresa(nome_arquivo="empresa_dados.xlsx", n=50):
    fake = Faker("pt_BR")

    # Lista de cargos e departamentos
    cargos_departamentos = [
        ("Analista Financeiro", "Financeiro", (2500, 4500)),
        ("Gerente Financeiro", "Financeiro", (7000, 12000)),
        ("Desenvolvedor", "TI", (3000, 8000)),
        ("Analista de Suporte", "TI", (2000, 4000)),
        ("Gerente de TI", "TI", (9000, 15000)),
        ("Analista de Marketing", "Marketing", (2500, 5000)),
        ("Gerente de Marketing", "Marketing", (8000, 13000)),
        ("Assistente de RH", "RH", (1800, 3500)),
        ("Coordenador de RH", "RH", (4000, 7000)),
    ]

    dados = []

    for _ in range(n):
        nome = fake.name()
        cargo, departamento, faixa_salario = random.choice(cargos_departamentos)
        salario = round(random.uniform(*faixa_salario), 2)
        email = fake.email()
        telefone = fake.phone_number()
        dados.append([nome, cargo, departamento, salario, email, telefone])

    # Criar DataFrame
    df = pd.DataFrame(
        dados, columns=["Nome", "Cargo", "Departamento", "Salário", "Email", "Telefone"]
    )

    # Salvar Excel
    df.to_excel(nome_arquivo, index=False)
    print(f"✅ Planilha '{nome_arquivo}' criada com {n} registros.")


if __name__ == "__main__":
    gerar_planilha_empresa()
