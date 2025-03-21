import barcode
from barcode.writer import ImageWriter

def gerar_codigo_barras(texto, formato="png"):
    """
    Gera um código de barras Code-128 a partir de uma string e salva como imagem.

    :param texto: String que será convertida em código de barras.
    :param formato: Formato da imagem de saída (png ou jpg).
    """
    # Definir o tipo de código de barras
    codigo_barras = barcode.get("code128", texto, writer=ImageWriter())

    # Nome do arquivo de saída
    nome_arquivo = f"codigo_barras_{texto}.{formato}"

    # Salvar o código de barras como imagem
    codigo_barras.save(nome_arquivo)

    print(f"Código de barras salvo como {nome_arquivo}")

# Exemplo de uso
gerar_codigo_barras("01490201035", "png")
