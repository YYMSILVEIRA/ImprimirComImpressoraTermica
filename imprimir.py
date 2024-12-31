"""Biblioteca responsável pelo processo de impressão das notas
A ideia é que esse seja um script compilado a parte que será chamado via
parametros de execução com o arquivo gerado pelo sistema."""

import win32print
import time
from PIL import Image

class ImprimirDocumento:
    def __init__(self, printer_name, titulo):
        self.printer = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(self.printer, 1, (titulo, None, "RAW"))
        win32print.StartPagePrinter(self.printer)
    def conteudoImpressao(self, conteudo):
        self.conteudo=conteudo.replace("ã","a")
        
        # Codificação do texto para CP-850 (ou outra página de código adequada)
        self.conteudo_encoded = self.conteudo.encode("cp850")  # Altere para cp1252, se necessário
        
        # Enviar o conteúdo
        win32print.WritePrinter(self.printer, self.conteudo_encoded)
    def corteTotalPapel(self):
        win32print.WritePrinter(self.printer, b"\x1d\x56\x42\x00")
    def corteParcialPapel(self):
        win32print.WritePrinter(self.printer, b"\x1d\x56\x01")
    def definirFonteNormal(self):
        # Comando ESC/POS para voltar ao tamanho normal
        win32print.WritePrinter(self.printer, b"\x1B\x21\x00")
    def definirFonteDuplaAltura(self):
        # Comando ESC/POS para definir dupla altura
        win32print.WritePrinter(self.printer, b"\x1B\x21\x10")
    def definirFonteDuplaLargura(self):
        # Comando ESC/POS para definir dupla largura
        win32print.WritePrinter(self.printer, b"\x1B\x21\x20")
    def definirFonteDuplaAlturaLargura(self):
        # Comando ESC/POS para definir dupla altura e largura
        win32print.WritePrinter(self.printer, b"\x1B\x21\x30")
    def imprimirImagem(self, caminho_imagem):
        # Abrir a imagem
        imagem = Image.open(caminho_imagem)

        # Converter a imagem para o modo '1' (1-bit pixels, preto e branco)
        imagem = imagem.convert("1")

        # Obter as dimensões da imagem
        largura, altura = imagem.size

        # Comando ESC/POS para definir o modo de impressão de imagem
        comando_imagem = b"\x1D\x76\x30\x00"

        # Preparar os dados da imagem
        dados_imagem = b""
        for y in range(altura):
            for x in range(0, largura, 8):
                byte = 0
                for bit in range(8):
                    if x + bit < largura:
                        pixel = imagem.getpixel((x + bit, y))
                        if pixel == 0:
                            byte |= 1 << (7 - bit)
                dados_imagem += bytes([byte])

        # Calcular o número de bytes por linha
        bytes_por_linha = largura // 8
        if largura % 8 != 0:
            bytes_por_linha += 1

        # Montar o comando completo
        comando_completo = comando_imagem + bytes([bytes_por_linha % 256, bytes_por_linha // 256, altura % 256, altura // 256]) + dados_imagem

        # Enviar o comando para a impressora
        win32print.WritePrinter(self.printer, comando_completo)
    def redimensionar_imagem(self, caminho_imagem, largura_maxima):
        # Abrir a imagem
        imagem = Image.open(caminho_imagem)

        # Calcular a proporção de redimensionamento
        proporcao = largura_maxima / imagem.width
        nova_altura = int(imagem.height * proporcao)

        # Redimensionar a imagem
        imagem_redimensionada = imagem.resize((largura_maxima, nova_altura), Image.Resampling.LANCZOS)

        return imagem_redimensionada
    def converter_imagem_para_escpos(self, imagem):
        # Converter a imagem para preto e branco (1-bit)
        imagem = imagem.convert("1")

        # Obter as dimensões da imagem
        largura, altura = imagem.size

        # Preparar os dados da imagem
        dados_imagem = b""
        for y in range(altura):
            for x in range(0, largura, 8):
                byte = 0
                for bit in range(8):
                    if x + bit < largura:
                        pixel = imagem.getpixel((x + bit, y))
                        if pixel == 0:  # 0 = preto, 1 = branco
                            byte |= 1 << (7 - bit)
                dados_imagem += bytes([byte])

        return dados_imagem, largura, altura
    def imprimir_imagem(self, caminho_imagem, largura_maxima=384):
        # Redimensionar a imagem
        imagem_redimensionada = self.redimensionar_imagem(caminho_imagem, largura_maxima)

        # Converter a imagem para o formato ESC/POS
        dados_imagem, largura, altura = self.converter_imagem_para_escpos(imagem_redimensionada)

        # Comando ESC/POS para imprimir imagem rasterizada
        comando_imagem = b"\x1D\x76\x30\x00"

        # Calcular o número de bytes por linha
        bytes_por_linha = largura // 8
        if largura % 8 != 0:
            bytes_por_linha += 1

        # Montar o comando completo
        comando_completo = (
            comando_imagem +
            bytes([bytes_por_linha % 256, bytes_por_linha // 256, altura % 256, altura // 256]) +
            dados_imagem
        )

        # Enviar o comando para a impressora
        win32print.WritePrinter(self.printer, comando_completo)
    def fecharImpressao(self):
        # Finalizar a impressão
        win32print.EndPagePrinter(self.printer)
        win32print.EndDocPrinter(self.printer)
        win32print.ClosePrinter(self.printer)
        print("Impressão concluída!")
    def max_caracteres_por_linha(self, fonte_ampliada=False):
        return 24 if fonte_ampliada else 48
    def quebrar_texto(self, texto, max_caracteres):
        linhas = []
        palavras = texto.split(" ")
        linha_atual = ""

        for palavra in palavras:
            if len(linha_atual) + len(palavra) + 1 <= max_caracteres:
                linha_atual += palavra + " "
            else:
                linhas.append(linha_atual.strip())
                linha_atual = palavra + " "

        if linha_atual:
            linhas.append(linha_atual.strip())

        return linhas
    def imprimir_texto_dinamico(self, texto, fonte_ampliada=False):
        # Determinar o número máximo de caracteres por linha
        max_caracteres = self.max_caracteres_por_linha(fonte_ampliada)

        # Definir o tamanho da fonte
        if fonte_ampliada:
            self.definirFonteDuplaLargura()
        else:
            self.definirFonteNormal()

        # Quebrar o texto em várias linhas
        linhas = self.quebrar_texto(texto, max_caracteres)

        # Imprimir cada linha
        for linha in linhas:
            self.conteudoImpressao(linha + "\n")



# Exemplo de uso

'''conteudo = """
\n
===============================================\n
                IB2S Softwares                 \n
            Rua dos Metalurgicos-67            \n
                Bairro Santa Rosa              \n
                   Taquara - RS                \n
===============================================\n
Descrição                Qtd      Valor        \n
-----------------------------------------------\n
Pão Francês               6      R$ 3,00       \n
Croissant                 2      R$ 7,00       \n
Bolo de Fubá              1      R$12,00       \n
-----------------------------------------------\n
TOTAL:                           R$22,00       \n
===============================================\n
          Obrigado pela preferência!           \n
===============================================\n
\n
"""
imprimir = ImprimirDocumento("TM_T20X", "Nota Fiscal")
#imprimir.conteudoImpressao(conteudo)
imprimir.imprimir_imagem("F:\caixaQt-master\Frente-de-caixa\cryptasec2\IconeCryptSec.png",484)
imprimir.corteTotalPapel()
imprimir.fecharImpressao()
'''
imprimir = ImprimirDocumento("TM_T20X", "Nota Fiscal")

# Texto longo para teste
texto_longo = (
    "Este é um exemplo de texto que será impresso na TM-T20X. "
    "A ideia é que o texto seja quebrado dinamicamente em várias linhas, "
    "considerando o tamanho da fonte e a largura máxima da impressora."
)

# Imprimir no modo normal
imprimir.imprimir_texto_dinamico(texto_longo, fonte_ampliada=False)

# Imprimir no modo ampliado
imprimir.imprimir_texto_dinamico(texto_longo, fonte_ampliada=True)

imprimir.corteTotalPapel()
imprimir.fecharImpressao()