from barcode import Code128
from barcode.writer import ImageWriter

def gerar_codigo_de_barras(string, formato="png"):
  """
  Gera um código de barras Code-128 a partir de uma string e salva em PNG ou JPG.

  Args:
    string: A string a ser codificada no código de barras.
    formato: O formato da imagem (png ou jpg).
  """
  codigo_de_barras = Code128(string)
  nome_arquivo = f"codigo_de_barras.{formato}"
  codigo_de_barras.save(nome_arquivo, writer=ImageWriter())

# Exemplo de uso:
gerar_codigo_de_barras("Eliezer", formato="png")
gerar_codigo_de_barras("Moraes Silva", formato="jpg")
