from reportlab.pdfgen import canvas
import pandas as pd
from PIL import Image

# Converte mm em points (unidade de medida usado na biblioteca reportlab)
def mmToPoint(mm):
    return mm / 0.352777

# Retorna uma lista de alunos obtida de um arquivo excell
def getAlunosEDiretor(caminhoPlanilha):
    lines = pd.read_excel(caminhoPlanilha)
    listaAlunos = []

    for index, aluno in lines.iterrows():
        listaAlunos.append(aluno['nome'])

    lines['diretor'] = lines['diretor'].str.replace(' ', '_')

    return listaAlunos

# Obtem dimensões da imagem do certificado
def getTamanhoImagem(caminho):
    img = Image.open(caminho) 
    largura = img.width 
    altura = img.height
    
    return (largura, altura)

'''
    Obtem caminho da imagem padrão do certificado
    Obtem largura e altura do certificado
    Obtem alunos para preencher o certificado
    Lê nome do diretor
    Retorna um dicionário com todos os dados para a emissão dos certificados
'''
def lerDadosPdf():
    print("Digite as seguintes informações do seu certificado:")

    caminhoImg = str(input("Caminho da imagem do certificado: "))
    caminhoPlanilha = str(input("Caminho da planilha de alunos: "))
    (largura, altura) = getTamanhoImagem(caminhoImg)
    (alunos, diretor) = getAlunosEDiretor(caminhoPlanilha)

    # diretor = str(input("Nome do diretor: "))

    dicionario = {
        'size': (largura, altura),
        'alunos': alunos,
        'diretor': diretor,
        'img': caminhoImg
    }

    return dicionario

'''
    Criar os certificados com fonte padrão
    Cria um certificado para cada aluno
    Todos os certificados são acompanados do nome do diretor
    As coordenadas para preenchimento do certifidado só são válidas para a imagem que acompanha o código
'''
def gerarCertificados(alunos, size, caminhoImgCertificado, diretor):
    for aluno in alunos:
        pdf = canvas.Canvas('./' + aluno + '.pdf', pagesize=size)
        pdf.drawImage(caminhoImgCertificado, 0, 0)

        pdf.setFont('Helvetica-Oblique', 80)
        pdf.drawCentredString(mmToPoint(355), mmToPoint(300), aluno)

        pdf.setFont('Helvetica', 24)
        pdf.drawCentredString(mmToPoint(168), mmToPoint(115), aluno)
        pdf.drawCentredString(mmToPoint(538), mmToPoint(115), diretor)
        
        pdf.save()

if __name__ == "__main__":
    dados = lerDadosPdf()
    gerarCertificados(dados['alunos'], dados['size'], dados['img'], dados['diretor'])