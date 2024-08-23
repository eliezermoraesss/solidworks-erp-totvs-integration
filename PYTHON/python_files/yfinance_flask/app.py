from flask import Flask, render_template, request, jsonify
import yfinance as yf
import plotly.graph_objects as go
import plotly.io as pio

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_data', methods=['POST'])
def get_data():
    ticker = request.form.get('ticker')
    if not ticker:
        return jsonify({'error': 'Ticker não fornecido.'})
    
    try:
        ativo = yf.Ticker(ticker)
        dados_historicos = ativo.history(period="1mo")
        
        if dados_historicos.empty:
            return jsonify({'error': 'Nenhum dado encontrado para o ticker fornecido.'})
        
        # Gerar gráfico com Plotly
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=dados_historicos.index, y=dados_historicos['Close'], mode='lines', name='Fechamento'))
        fig.update_layout(title=f'Gráfico de Fechamento para {ticker}', xaxis_title='Data', yaxis_title='Preço')
        
        # Converter o gráfico para HTML
        graph_html = pio.to_html(fig, full_html=False)
        
        return jsonify({'graph': graph_html})
    
    except Exception as e:
        return jsonify({'error': str(e)})

if __name__ == '__main__':
    app.run(debug=True)
