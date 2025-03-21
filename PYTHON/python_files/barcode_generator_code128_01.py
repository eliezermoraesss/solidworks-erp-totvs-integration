import barcode
from barcode.writer import ImageWriter
from PIL import Image

def generate_barcode(text, output_filename, format='png'):
    """
    Gera um código de barras Code-128 a partir de uma string.
    
    Parâmetros:
    - text: String a ser convertida em código de barras
    - output_filename: Nome do arquivo de saída (sem extensão)
    - format: Formato da imagem ('png' ou 'jpg'), padrão é 'png'
    
    Retorna o caminho do arquivo gerado
    """
    try:
        # Cria o código de barras Code-128
        code128 = barcode.get_barcode_class('code128')
        
        # Configura o writer de imagem
        image_writer = ImageWriter()
        
        # Gera o código de barras
        generated_barcode = code128(text, writer=image_writer)
        
        # Salva o código de barras
        full_filename = generated_barcode.save(output_filename)
        
        print(f"Código de barras gerado com sucesso: {full_filename}")
        return full_filename
    
    except Exception as e:
        print(f"Erro ao gerar código de barras: {e}")
        return None

def main():
    # Exemplo de uso
    texto = input("Digite o texto para gerar o código de barras: ")
    formato = input("Escolha o formato (png/jpg) [padrão: png]: ").lower() or 'png'
    nome_arquivo = input("Digite o nome do arquivo de saída (sem extensão): ")
    
    generate_barcode(texto, nome_arquivo, formato)

if __name__ == "__main__":
    main()