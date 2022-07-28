from reportlab.pdfgen import canvas

# Converte mm em points (unidade de medida usado na biblioteca reportlab)
def mmToPoint(mm):
    return mm / 0.352777

'''
    Lê largura e altura do certificado
    Lê alunos para preencher o certificado
    Lê nome do diretor
    Lê caminho da imagem padrão do certificado
    Retorna um dicionário com todos os dados para a emissão dos certificados
'''
def lerDadosPdf():
    print("Digite as seguintes informações do seu certificado:")

    largura = int(input("Largura: "))
    altura = int(input("Altura: "))
    print()

    print("Digite as informações do seus alunos:")
    
    qtdAlunos = int(input("Quantidade de alunos: "))
    alunos = []
    for i in range(qtdAlunos):
        alunos.append(str(input(f"Nome do {i + 1}º aluno: ")))
    
    print()

    diretor = str(input("Nome do diretor: "))

    print()
    img = str(input("Caminho da imagem do certificado: "))

    dicionario = {
        'size': (largura, altura),
        'alunos': alunos,
        'diretor': diretor,
        'img': img
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
        pdf = canvas.Canvas('./' + aluno +'.pdf', pagesize=size)
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