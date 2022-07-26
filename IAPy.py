# Bibliotecas usadas no projeto todo
# -*- coding:utf-8 -*-
import quopri
from lib2to3.pgen2.pgen import DFAState
from ast import Break
import os
import time
from PIL import Image
from PIL import ImageFont

# Bibliotecas usadas no projeto no tratamento das imagens
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw

# Bibliotecas usadas no projeto email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import email
import os
import smtplib
import html2text
import imaplib
import time

# Bibliotecas usadas no projeto tratar dados
import os
import re

# Bibliotecas usadas no projeto tratar listas
import csv

from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> Projeto PI <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


def save_ass(imagem, nome):  # Salvando imagem no local preterido
    # Redimensionando quando necessário
    # imagem = imagem.resize((536, 227), Image.ANTIALIAS) Caso seja necessário redmimensionar
    imagem.save(
        f'{__location__}/AssinaturasProntas/{nome}.png', 'PNG', quality=95)


def abrirCaminhoImagemCriada(nome):
    return f'{__location__}\\AssinaturasProntas\\{nome}.png'


def abrirImagemModelo():
    return Image.open(f'{__location__}\\Images\\ass.png')


def abrirLista():
    return f'{__location__}\\Anexos\\ListaTeste.csv'


def definirFonte(fonte):  # Definindo fontes
    # Fonte usada mais de duas vezesF
    PoppinsLight = os.environ['LOCALAPPDATA'] + \
        "/Microsoft/Windows/Fonts/Poppins-Light.ttf"
    match fonte:
        case "font_nome":
            return ImageFont.truetype(os.environ['LOCALAPPDATA'] + "/Microsoft/Windows/Fonts/Poppins-ExtraBold.ttf", 26)
        case "font_dep":
            return ImageFont.truetype(os.environ['LOCALAPPDATA'] + "/Microsoft/Windows/Fonts/Poppins-SemiBold.ttf", 16)
        case "font_tel":
            return ImageFont.truetype(PoppinsLight, 17)
        case "font_email":
            return ImageFont.truetype(PoppinsLight, 17)
        case "font_site":
            return ImageFont.truetype(PoppinsLight, 17)
        case "font_unidade":
            return ImageFont.truetype(os.environ['LOCALAPPDATA'] + "/Microsoft/Windows/Fonts/Poppins-SemiBold.ttf", 17)
        case default:
            return


# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> Tratar Imagem <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
def escreverImagem(imagem, texto, font, x, y):  # Escreve na imagem
    draw = ImageDraw.Draw(imagem)
    # Defino a cor
    draw.text((x, y), texto, font=font, fill='#000000')
    return imagem


def openHtml(nome):
    valores = ''
    x = open(f"{__location__}\\Html\\{nome}.html", "r", encoding="utf8")
    for linha in x:
        valores = valores + linha
    return valores


def validarInsercaoUsuario(x, remetente):
    if x == 'base' or x == '' or x == ' ':
        enviarEmailErro(remetente)
        return False


def enviarEmailSucesso(cam_assinatura, email, remetente):
    # Iniciar servidor SMTP
    global host
    global port
    global login
    global senha
    strTo = email

    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] = 'Olá, aqui é a IA Py :)'
    msgRoot['From'] = login
    msgRoot['To'] = strTo
    msgRoot.preamble = 'This is a multi-part message in MIME format.'

    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)

    msgText = MIMEText('This is the alternative plain text message.')
    msgAlternative.attach(msgText)
    msgText = MIMEText(openHtml('emailSucess'), 'html')
    msgAlternative.attach(msgText)

    fp = open(cam_assinatura, 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()

    msgImage.add_header('Content-ID', '<image1>')
    msgRoot.attach(msgImage)

    smtp = smtplib.SMTP('outlook.office365.com', port)
    smtp.ehlo()
    smtp.starttls()
    smtp.login(login, senha)
    try:
        smtp.sendmail(login,
                      email, msgRoot.as_string())
    except:
        enviarEmailErro(remetente)
    # Encerrando
    smtp.quit()


def gerarImagemAssinatura(nome, departamento, ramal, celular, email, unidade, site, remetente):
    imagem = abrirImagemModelo()  # Predefinida pelo usuário
    if((email == 'base') or (email == '') or (email == '\n') or (email == ' ')):
        email = remetente
    if(celular == "base" or celular == ' '):  # Contem apenas um número
        imagem = escreverImagem(
            imagem, ramal, definirFonte("font_tel"), 379, 180)
        celular = "."
    elif(ramal == "base" or ramal == ' '):
        imagem = escreverImagem(imagem, celular, definirFonte(
            "font_tel"), 379, 180)  # Contem apenas celular
        ramal = "."
    elif(ramal != "base" and celular != "base" and ramal != ' ' and celular != ' '):  # Contem dois números
        imagem = escreverImagem(
            imagem, ramal, definirFonte("font_tel"), 379, 190)
        imagem = escreverImagem(
            imagem, celular, definirFonte("font_tel"), 379, 170)

    imagem = escreverImagem(imagem, nome, definirFonte(
        "font_nome"), 335, 95,)  # Escrevendo na imagem
    imagem = escreverImagem(imagem, departamento,
                            definirFonte("font_dep"), 335, 128)
    imagem = escreverImagem(
        imagem, email, definirFonte("font_email"), 379, 230)
    imagem = escreverImagem(imagem, site, definirFonte("font_site"), 379, 278)
    imagem = escreverImagem(imagem, unidade, definirFonte(
        "font_unidade"), 455, 349)

    # Tratando para sempre ter as informações preteridas
    if (validarInsercaoUsuario(nome, remetente) == False or validarInsercaoUsuario(unidade, remetente) == False or validarInsercaoUsuario(celular, remetente) == False or validarInsercaoUsuario(ramal, remetente) == False or validarInsercaoUsuario(departamento, remetente) == False):
        return

    if(remetente == 'lista'):
        imagem.show(imagem)  # Apresenta a imagem
        return

    save_ass(imagem, nome)  # Salva a imagem pronta
    cam_assinatura = abrirCaminhoImagemCriada(
        nome)  # Pega o caminho da imagem salva

    imagem.show(imagem)  # Apresenta a imagem

    # Envia o e-mail com a imagem em anexo
    enviarEmailSucesso(cam_assinatura, email, remetente)

# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> Projeto Email <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


def enviarEmailErro(remetente):
    # Iniciar servidor SMTP
    global host
    global port
    global login
    global senha

    host = smtplib.SMTP('outlook.office365.com', port)
    host.starttls()
    host.login(login, senha)

    email_msg = MIMEMultipart()
    email_msg['From'] = login
    email_msg['To'] = remetente
    email_msg['Subject'] = "Erro na Assinatura"
    email_msg.attach(MIMEText(openHtml('emailFailed'), 'html'))

    # Enviar e-mail
    try:
        host.sendmail(email_msg['From'],
                      email_msg['To'], email_msg.as_string())
    except:
        enviarEmailErro(remetente)

    # Encerrando
    host.quit()


def lerCorpoEmail(email_message):  # Percorre o e-mail todo
    for payload in email_message.get_payload():
        break
    return payload.get_payload()


def validaErroRemetenteEmail(x):
    global login
    if('Microsoft' in x or login in x):
        return False


def botAnalisarEmail():
    global host
    global port
    global login
    global senha
    # Defininido local para salvar a lista .csv
    detach_dir = f'{__location__}\\Anexos'
    # Selecionando o fornecedor de e-mail
    mail = imaplib.IMAP4_SSL('outlook.office365.com')
    mail.login(login, senha)  # Fazendo login
    mail.select("inbox")  # Selecionando caixa de entrada para analisar
    try:
        # Buscando e-mails não lidos
        result, data = mail.uid('search', None, '(UNSEEN)')
        # Armazenando a lista dos e-mails não lidos
        inbox_item_list = data[0].split()
        # Armazenando o primeiro e-mail não lido
        most_recent = inbox_item_list[-1]
        # Marcando como lido o primeiro e-mail
        result2, email_data = mail.uid('fetch', most_recent, '(RFC822)')
        # Decodificando o cabeçalho da mensagem para o tipo RAW (Tipo de dados sem tratamento)
        raw_email = email_data[0][1].decode("UTF-8")
        email_message = email.message_from_string(
            raw_email)  # Tratando o topo RAW para tipo String
        for part in email_message.walk():  # Percorrendo toda a mensagem atrás de anexos
            if part.get_content_maintype() == 'multipart':  # Verificando se está no cabeçalho da mensagem
                continue
            # Verificando se o corpo ainda não iniciou
            if part.get('Content-Disposition') is None:
                continue
        if part.get_filename() is not None:  # Tratando quando não tem nenhum tipo de anexo
            # Selecionando a pasta para armazenar os arquivos
            att_path = os.path.join(detach_dir, part.get_filename())
            if not os.path.isfile(att_path):  # Tratando para salvar duplicado
                fp = open(att_path, 'wb')  # Chamando a função de salvar do Py
                # Decodificando em arquivo legível
                fp.write(part.get_payload(decode=True))
                fp.close()  # Concluindo processo de salvamento
        if email_message.is_multipart():
            # Tratando mensagens da Microsoft e e-mails repetidos
            if(validaErroRemetenteEmail(email_message['From']) == False):
                return
            for payload in email_message.get_payload():  # Percorrendo o e-mail para pegar a mensagem
                for part in email_message.walk():  # Percorrendo e tratanto quando o tipo de texto for simples
                    if (part.get_content_type() == "text/plain"):  # Validando o texto simples
                        tratarCorpo(  # Chamando função para tratar os dados
                            part.get_payload(None, True), email_message['From'])
                return
        else:
            # Tratando erros
            if(validaErroRemetenteEmail(email_message['From']) == False):
                return
            html = f"{email_message.get_payload()}"
            h = html2text.HTML2Text()
            h.ignore_links = True
            output = (h.handle(f'''{html}'''))
            # Chamando função para tratar os dados
            tratarCorpo(output, email_message['From'])
    except IndexError:
        print("E-mail vazio")

# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> Tratar Dados <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


def tratarCasosMaiusculo(x):
    # Strings que são obrigatoriamente maiúsculas
    padroes_maiusculo = "ceo, cfo, pcb, pcp, dp, ti, rh, ii, i"
    # Strings que são obrigatoriamente minúsculas
    excecoes = ['da', 'de', 'do', 'a', 'e', 'o', 'em']
    x = x.split(' ')  # Separa a String pelos espaços em branco.
    titled_string = ''
    for i in x:  # Percorro a String toda.
        if i in excecoes:  # Quando for excecao ignora.
            titled_string += i
        # Quando for predefinido como maiusculo faz o tratamento.
        elif i in padroes_maiusculo:
            titled_string += i.upper()
        else:  # Deixa apena a primeira letra em maiusculo
            titled_string += i.title()
        titled_string += ' '
    return(titled_string[:len(titled_string)-1])


def tratarStringCorpo(x):  # Removendo espaços vazios
    dictionary = {
        1: ["\n", " "],
        2: [r"\xc2", ""],
        3: [r"\xa0", ""],
        4: [r"\xc3", ""],
        5: [r"\xa7", ""],
        6: [r"\x90", ""],
        7: [r"\x81", ""],
        8: [r"\xef", ""],
        9: [r"\xa3o", ""],
        10: [r"\x", ""],
        11: [r"https://mailto:", ""],
        12: ["nome", ""],
        13: ["departamento", ""],
        14: ["ramal", ""],
        15: ["email", ""],
        16: ["celular", ""],
        17: [r"\r", ""],
        18: [r"b'", ""],
        19: ["unidade", ""],
        21: ["*", ""],
        22: [",", ""],
        23: ["/", ""],
        24: ["(", ""],
        25: [")", ""],
        26: ["\r", ""]}
    for i in dictionary:  # Remove todas as ocorrências erradas
        x = x.replace(dictionary[i][0], dictionary[i][1])
    return re.sub(r"^\s+|\s+$", "", x)


def instanciarVariavel():
    return "base"


def quebrarCorpo(x, dados):
    try:
        mys = dados.split(' ')
        for y in mys:
            if x in y:
                idx = mys.index(y)
                y = re.sub(r"^\s+|\s+$", "", y)
        mys[idx] = "\n" + x
        if x == "unidade":
            mys[idx+1] = mys[idx+1] + "\n"
        return " ".join(mys)
    except:
        return dados


def tratarCorpoBruto(dados):
    try:
        temp = dados.replace(
            "\\/", "/").encode().decode('unicode_escape')
        decoded_string = quopri.decodestring(dados)
        dados = decoded_string
    except:
        pass
    try:
        temp = dados.decode("UTF-8")
        dados = temp
    except:
        try:
            temp = dados.decode('latin-1').encode("utf-8")
            temp = temp.decode("UTF-8")
            dados = temp
        except:
            pass
        pass
    dados = dados.replace("\n", " ")
    return dados.lower()  # Deixando tudo minúsculo para tratar


def tratarRemetenteBruto(remetente):
    if '<' in remetente:
        remetente = re.search('<(.+?)>', remetente).group(1)
    return remetente


def tratarCorpo(dados, remetente):
    nome = instanciarVariavel()
    departamento = instanciarVariavel()
    ramal = instanciarVariavel()
    celular = instanciarVariavel()
    email = instanciarVariavel()
    unidade = instanciarVariavel()
    dados = tratarCorpoBruto(dados)
    remetente = tratarRemetenteBruto(remetente)
    try:
        dados = quebrarCorpo("nome", dados)
        dados = quebrarCorpo("departamento", dados)
        dados = quebrarCorpo("ramal", dados)
        dados = quebrarCorpo("celular", dados)
        dados = quebrarCorpo("email", dados)
        dados = quebrarCorpo("unidade", dados)
        for x in dados.split('\n'):
            if "nome" in x:
                # Deixando o nome com as primeiras letras maiusculas e tratando casos especiais.
                nome = tratarCasosMaiusculo(tratarStringCorpo(x))
            elif "departamento" in x:
                # Deixando o departamento com as primeiras letras maiusculas e tratando casos especiais.
                departamento = tratarCasosMaiusculo(tratarStringCorpo(x))
            elif "ramal" in x:
                ramal = tratarStringCorpo(x)
            elif "celular" in x:
                celular = tratarStringCorpo(x)
            elif "email" in x:
                email = tratarStringCorpo(x)
                if '<' in email:
                    # Encontrando hiperlink html que fica no e-mail e tratando, caso não tiver não é um problema.
                    x = email.find('<')
                    # Removendo para entregar a String limpa.
                    email = (email[0:x])
            elif "unidade" in x:
                # Deixando a unidade com as primeiras letras maiusculas e tratando casos especiais.
                unidade = tratarCasosMaiusculo(tratarStringCorpo(x))
                break
            else:
                rest = x  # Atribuindo o resto da String desnecessário em uma variável para evitar erro

        # Encontrando o domínio do e-mail para gerar o site. Essa informação é mutável de acordo com a necessidade da organização.
        x = remetente.find('@')
        # Removendo os dados do usuário para ter apenas o domínio para entregar a String limpa somente com o site.
        site = (remetente[x+1::])
        if email == 'base':
            email = remetente
        gerarImagemAssinatura(nome, departamento, ramal, celular,
                              email, unidade, site, remetente)
    except:
        enviarEmailErro(remetente)

# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> Projeto Tratar Lista <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


def enviarLista():
    with open(abrirLista(), mode='r', encoding='utf-8') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=';')
        csv_reader.__next__()  # Ativar essa linha se a lista tiver cabeçalho

        for row in csv_reader:
            print(row[0] + ', ' + row[1] + ', ' +
                  row[2] + ', ' + row[3] + ', ' + row[4])
            imagem = abrirImagemModelo()
            # Pegando informações
            nome = row[0]
            departamento = row[1]
            ramal = row[2]
            celular = row[3]
            email = row[4]
            unidade = row[5]

            x = email.find('@')
            site = (email[x+1::])
            gerarImagemAssinatura(nome, departamento, ramal, celular,
                                  email, unidade, site, "lista")


# Iniciar servidor SMTP
host = "smtp-mail.outlook.com"
port = "587"
login = "email"
senha = "senhha"

# Diretorio atual
__location__ = os.path.realpath(
    os.path.join(os.getcwd(), os.path.dirname(__file__)))


def main():
    while True:
        botAnalisarEmail()


if __name__ == "__main__":
    main()
