import qrcode

def gerar_qr_code(conteudo, nome_arquivo):
    """
    Gera um QR Code a partir de um conteúdo e salva como PNG.
    """
    qr = qrcode.QRCode(
        version=1,  # Tamanho do QR Code (1 é o menor, 40 é o maior)
        error_correction=qrcode.constants.ERROR_CORRECT_L,  # Nível de correção de erro
        box_size=10,  # Tamanho de cada quadrado
        border=4,  # Tamanho da borda
    )
    qr.add_data(conteudo)
    qr.make(fit=True)  # Ajusta o QR Code ao conteúdo

    # Gera a imagem do QR Code
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(nome_arquivo)
    print(f"QR Code salvo como {nome_arquivo}")

if __name__ == "__main__":
    conteudo = input("Digite o conteúdo para o QR Code: ")
    nome_arquivo = input("Digite o nome do arquivo para salvar (ex: qrcode.png): ")
    gerar_qr_code(conteudo, nome_arquivo)
