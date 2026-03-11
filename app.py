from flask import Flask, jsonify, send_from_directory, request
from flask_cors import CORS
import pyodbc
from datetime import datetime
import pandas as pd
import os

app = Flask(__name__)
CORS(app)

conn_str = "DSN DO BANCO DE DADOS;" 
ARQUIVO_EXCEL = 'CAMINHO DO EXCEL'  

def formatar(dt):
    if isinstance(dt, datetime):
        return dt.strftime('%d/%m/%Y')
    return "-" if dt is None or str(dt).strip() == "" else str(dt)


def processar_estatisticas_excel(inicio, fim):
    """Lê o Excel apenas uma vez e retorna ambos os contadores."""
    try:
        if not os.path.exists(ARQUIVO_EXCEL):
            return 0, 0
        
        df = pd.read_excel(ARQUIVO_EXCEL, sheet_name='REGISTRO DE PEDIDOS', usecols=['LIBERAÇÃO', 'STATUS'])
        
        df['LIBERAÇÃO'] = pd.to_datetime(df['LIBERAÇÃO'], errors='coerce')
        df = df.dropna(subset=['LIBERAÇÃO'])

        if inicio and fim:
            dt_inicio = pd.to_datetime(inicio)
            dt_fim = pd.to_datetime(fim)
            df = df[(df['LIBERAÇÃO'] >= dt_inicio) & (df['LIBERAÇÃO'] <= dt_fim)]

        no_prazo = int((df['STATUS'] == 'No Prazo').sum())
        atrasado = int((df['STATUS'] == 'Atrasado').sum())
        
        return no_prazo, atrasado
    except Exception as e:
        print(f"Erro ao processar Excel: {e}")
        return 0, 0

@app.route('/pedidos/status_excel', methods=['GET'])
def get_status_excel():
    data_inicio = request.args.get('inicio')
    data_fim = request.args.get('fim')
    
    no_prazo, atrasado = processar_estatisticas_excel(data_inicio, data_fim)
    
    return jsonify({
        'no_prazo': no_prazo, 
        'atrasado': atrasado
    })

@app.route('/pedidos/pendentes_coleta', methods=['GET'])
def get_pendentes_coleta():
    data_inicio = request.args.get('inicio')
    data_fim = request.args.get('fim')
    tipo_data = request.args.get('tipo_data', 'emissao') 
    
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM dbo.BANCO DAS EXP('', '', '')")
        rows = cursor.fetchall()

        pendentes = []
        for row in rows:
            try:
                
                val_12 = str(row[12]).strip() if len(row) > 12 else ""
                val_15 = str(row[15]).strip() if len(row) > 15 else ""

                protocolo = str(row[13]).strip() if len(row) > 13 else ""
                
                pedido   = str(row[5]).strip()
                vendedor = str(row[1]).strip()
                nf       = str(row[3]).strip()
                cliente  = str(row[9]).strip() if len(row) > 9 else str(row[2])
                transp = str(row[7]).strip().upper() if (len(row) > 7 and row[7] is not None) else ""
              
                
                dt_col_final = val_12 if val_12 and val_12.lower() != "none" else val_15
                data_emissao = row[2] 
                              
                alvo_filtro = ""
                if tipo_data == 'status':
                    
                    if dt_col_final and "/" in dt_col_final:
                        partes = dt_col_final.split('/')
                        alvo_filtro = f"{partes[2]}-{partes[1]}-{partes[0]}"
                else:
                    if isinstance(data_emissao, datetime):
                        alvo_filtro = data_emissao.strftime('%Y-%m-%d')

                
                if data_inicio and data_fim:
                    if not alvo_filtro or not (data_inicio <= alvo_filtro <= data_fim):
                        continue
               

                if nf != "" and nf != "None":
                    pendentes.append({
                        "pedido": pedido,
                        "nf": nf,
                        "vendedor": vendedor,
                        "cliente": cliente,
                        "transp": transp,
                        "data_col": dt_col_final,
                        "producao": "FATURADO",
                        "protoc_coleta": protocolo ,
                        "emissao": formatar(data_emissao)
                    })      
            except: 
                continue
        
        return jsonify(pendentes)
    except Exception as e:
        return jsonify({"erro": str(e)}), 500
    finally:
        if 'conn' in locals(): conn.close()

@app.route('/<path:filename>')
def custom_static(filename):
    return send_from_directory('.', filename)

@app.route('/')
def index():
    return send_from_directory('.', 'inicio.html')    

@app.route('/pedido/<id_pedido>', methods=['GET'])
def get_pedido(id_pedido):
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM dbo.BANCO PV('*') WHERE NUMERO = ?", (id_pedido,))
        row_p = cursor.fetchone()
        
        cursor.execute("SELECT * FROM dbo.BANCO EXP('', '', '') WHERE PEDIDO = ?", (id_pedido,))
        row_e = cursor.fetchone()

        if row_p:
            
            data = {
                
                "empresa":     row_p[0],  
                "pedido":      row_p[1],  
                "cliente":     row_p[3],  
                "emissao_ped": formatar(row_p[4]), 
                "vendedor":    row_p[13], 
                "producao":    row_p[16], 
                
                "data_col":    formatar(row_e[12]) if row_e else "NÃO EXPEDIDO",
                "nf":          row_e[3] if row_e else "Pendente",
                "emissao_nf":  formatar(row_e[2]) if row_e else "-",
                "transp":      row_e[7] if row_e else "N/A"
            }
            return jsonify(data)
        
        return jsonify({"erro": "Pedido não encontrado"}), 404

    except Exception as e:
        print(f"Erro: {e}")
        return jsonify({"erro": str(e)}), 500
    finally:
        if 'conn' in locals(): conn.close()
if __name__ == '__main__':
    
    app.run(host='0.0.0.0', port=5000, debug=True)