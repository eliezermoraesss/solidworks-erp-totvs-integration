from flask import Flask, request, render_template
import qrcode
from io import BytesIO
import base64

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    qr_code_data = None
    if request.method == "POST":
        conteudo = request.form.get("conteudo", "https://example.com")
        # Gera o QR Code
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(conteudo)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")

        # Converte para Base64 para exibir no HTML
        buffer = BytesIO()
        img.save(buffer, format="PNG")
        qr_code_data = base64.b64encode(buffer.getvalue()).decode("utf-8")

    return render_template("index.html", qr_code_data=qr_code_data)

if __name__ == "__main__":
    app.run(debug=True)
