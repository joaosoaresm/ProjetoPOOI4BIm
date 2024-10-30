import pandas as pd
import os

# Diretório de saída
output_dir = "C:/Users/joaov/OneDrive/Desktop/projeto4bimestre_pooi"

# Função para validar entrada de dados
def validar_matricula():
    while True:
        matricula = input("Digite a matrícula do aluno (máximo 6 dígitos): ")
        if matricula.isdigit() and len(matricula) == 6:
            return matricula
        print("Matrícula inválida. Tente novamente.")

def validar_nome():
    while True:
        nome = input("Digite o nome do aluno: ")
        if nome.isalpha():
            return nome
        print("Nome inválido. Use apenas letras.")

def validar_nota(disciplina):
    while True:
        try:
            nota = float(input(f"Digite a nota de {disciplina} (0 a 10): "))
            if nota < 0 or nota > 10:
                print("Nota fora do intervalo permitido. Tente novamente.")
                nota = float(input(f"Digite a nota de {disciplina} (0 a 10): "))
            else:
                return nota
        except ValueError:
            print("Entrada inválida. Digite um número.")

# Coleta e valida os dados dos alunos
alunos = []
while True:
    matricula = validar_matricula()
    nome = validar_nome()
    nota_pvb = validar_nota("PVB")
    nota_paw = validar_nota("PAW")
    nota_bd = validar_nota("BD")
    nota_pooi = validar_nota("POOI")
    
    # Cálculo da média
    media = ((nota_pvb + nota_paw + nota_bd + nota_pooi) / 4)
    
    # Determinação da situação do aluno
    if media >= 6:
        situacao = "APROVADO"
    elif media < 6 and media >= 3.75:
        situacao = "EXAME FINAL"
    else:
        situacao = "RETIDO"
    
    # Adiciona os dados ao dicionário
    alunos.append({
        "Matricula": matricula,
        "Aluno": nome,
        "PVB": nota_pvb,
        "PAW": nota_paw,
        "BD": nota_bd,
        "POOI": nota_pooi,
        "Média": media,
        "Situação": situacao
    })
    
    # Pergunta se deseja adicionar outro aluno
    if input("Deseja adicionar outro aluno? (s/n): ").lower() != 's':
        break

# Criação do arquivo Excel
df = pd.DataFrame(alunos)
df.to_excel("notasfinaisalunos.xlsx", index=False)
print(f"Arquivo Excel gerado em: {"notasfinaisalunos.xlsx"}")

# Função para criar o arquivo HTML para cada aluno
def gerar_html(aluno):
    matricula = aluno["Matricula"]
    nome = aluno["Aluno"]
    pvb = aluno["PVB"]
    paw = aluno["PAW"]
    bd = aluno["BD"]
    pooi = aluno["POOI"]
    media = aluno["Média"]
    situacao = aluno["Situação"]
    
    # Define a cor de fundo com base na situação
    if situacao == "APROVADO":
        cor_fundo = "rgb(8, 147, 216)"
    elif situacao == "EXAME FINAL":
        cor_fundo = "rgb(179, 168, 4)"
    else:
        cor_fundo = "red"
    
    # Criação do conteúdo HTML
    html_content = f"""
    <html>
    <head><title>Notas do Aluno {nome}</title></head>
    <body style="text-align: center;">
        <tr><p>MATRICULA: {matricula}</p></tr>
        <tr><p>ALUNO: <b  style="color: red;">{nome}</b></p></tr>
        <table border="0.2" style="width: 50%; margin: 0 auto; background-color: {cor_fundo}; color: white;">
            <tr><th>PVB</th><th>PAW</th><th>BD</th><th>POOI</th><th>MEDIA</th></tr>
            <tr><th>{pvb}</th><th>{paw}</th><th>{bd}</th><th>{pooi}</th><th>{media}</th></tr>
        </table>
        <tr><p>Situação Final: <b style="color: {cor_fundo};">{situacao}</b></p></tr>
    </body>
    </html>
    """
    
    # Salva o HTML em arquivo
    html_path = os.path.join(output_dir, f"{matricula}.html")
    with open(html_path, "w") as file:
        file.write(html_content)
    print(f"Arquivo HTML gerado para o aluno {nome} em: {html_path}")

# Pergunta se o usuário deseja gerar arquivos HTML para cada aluno
if input("Deseja gerar arquivos HTML para cada aluno? (s/n): ").lower() == 's':
    for aluno in alunos:
        gerar_html(aluno)