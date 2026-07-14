from flask import Flask, jsonify, send_from_directory, request , render_template
from flask_cors import CORS
import pyodbc
from datetime import datetime
import pandas as pd
import os
import qrcode
import io 
from flask import send_file
import requests
from io import StringIO
from flask import render_template_string
import time
from flask import Flask, jsonify, request
from datetime import datetime
import openpyxl
from openpyxl import Workbook
import os

app = Flask(__name__)
CORS(app)

conn_str = "DSN=SEU_DSN;UID=seu_usuario;PWD=sua_senha;" 
ARQUIVO_EXCEL = 'caminho/para/seu/Entrega_Pedido_BDados.xlsm'  

STATUS_DATA = {
    "ultima_atualizacao": ""
}

def formatar(dt):
    if isinstance(dt, datetime):
        return dt.strftime('%d/%m/%Y')
    return "-" if dt is None or str(dt).strip() == "" else str(dt)


def processar_estatisticas_excel(inicio, fim):
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
   
@app.route('/api/status-pbi', methods=['GET'])
def get_status_pbi():
    return jsonify(STATUS_DATA)

@app.route('/api/webhook-pbi', methods=['POST'])
def webhook_pbi():
    global STATUS_DATA
    try:
        agora = datetime.now()
        timestamp_formatado = agora.strftime('%d/%m/%y, %H:%M:%S')
        
        STATUS_DATA["ultima_atualizacao"] = timestamp_formatado
        
        print(f"✅ Power BI atualizado em: {timestamp_formatado}")
        return jsonify({"status": "sucesso", "data": timestamp_formatado}), 200
    except Exception as e:
        print(f"❌ Erro no webhook: {e}")
        return jsonify({"status": "erro", "mensagem": str(e)}), 500


@app.route('/pedido/qrcode/<id_pedido>', methods=['GET'])
def gerar_qrcode(id_pedido):
 
    base_url = f"http://SEU_IP_LOCAL:5000/pedido_detalhes.html?id={id_pedido}"
    
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(base_url)
    qr.make(fit=True)
    
    img = qr.make_image(fill_color="black", back_color="white")
    
    img_io = io.BytesIO()
    img.save(img_io, 'PNG')
    img_io.seek(0)
    
    return send_file(img_io, mimetype='image/png')

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
        cursor.execute("SELECT * FROM dbo.FN_0101_VEN_EXPEDICAO('', '', '')")
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
        try:
        
            SHEET_URL = "https://docs.google.com/spreadsheets/d/SEU_ID_DA_PLANILHA/export?format=csv&gid=SEU_GID"
            res_sheets = requests.get(SHEET_URL, timeout=15)
            
            if res_sheets.status_code == 200:
                df_sheets = pd.read_csv(StringIO(res_sheets.text), sep=',', engine='python')
                df_sheets.columns = [str(c).strip().upper() for c in df_sheets.columns]

                if 'PEDIDO' in df_sheets.columns:
                    col_ped = 'PEDIDO'
                    col_loc = [c for c in df_sheets.columns if 'LOCALIZ' in c][0]
                    mapa = dict(zip(
                        df_sheets[col_ped].astype(str).str.strip().str.upper(), 
                        df_sheets[col_loc].astype(str).str.strip()
                    ))
                else:
                    mapa = dict(zip(
                        df_sheets.iloc[:, 2].astype(str).str.strip().str.upper(), 
                        df_sheets.iloc[:, 3].astype(str).str.strip()
                    ))

                for p in pendentes:
                    cod_sistema = str(p.get('pedido', '')).strip().upper()
                    loc_encontrada = mapa.get(cod_sistema, "---")
                    loc_str = str(loc_encontrada).strip()
                    
                    if loc_str.lower() in ['nan', 'none', '', 'empty']:
                        p['localizacao'] = "---"
                    else:
                        p['localizacao'] = loc_str
            else:
                for p in pendentes: p['localizacao'] = "Erro Conexão"

        except Exception as e:
            print(f"DEBUG_ERRO: {e}") 
            for p in pendentes: p['localizacao'] = "OFFLINE"
        
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

@app.route('/conferir_geral')
def conferir_geral():
    try:
        dia_busca = request.args.get('dia')
        mes_busca = request.args.get('mes')
        ano_busca = request.args.get('ano', '2026')
        
        timestamp = int(time.time())
        
        url_sheets = f"https://docs.google.com/spreadsheets/d/e/SEU_LINK_PUBLICADO/pub?gid=SEU_GID&single=true&output=csv&cache_bust={timestamp}"
        res = requests.get(url_sheets)
        
        from io import StringIO 
        df_sheets = pd.read_csv(StringIO(res.text))
        
        df_sheets['dt_obj'] = pd.to_datetime(df_sheets.iloc[:, 0], dayfirst=True, errors='coerce')
        df_sheets = df_sheets.dropna(subset=['dt_obj'])
        
        if dia_busca and dia_busca.strip() != "":
            df_sheets = df_sheets[df_sheets['dt_obj'].dt.day == int(dia_busca)]
        if mes_busca and mes_busca.strip() != "":
            df_sheets = df_sheets[df_sheets['dt_obj'].dt.month == int(mes_busca)]
        if ano_busca and ano_busca.strip() != "":
            df_sheets = df_sheets[df_sheets['dt_obj'].dt.year == int(ano_busca)]
            
        if df_sheets.empty:
            return render_template('indexnf.html', tabelas=[])
            
        df_sheets['NF_PLANILHA'] = df_sheets.iloc[:, 2].astype(str).str.strip()
        df_sheets['NF_JOIN'] = pd.to_numeric(df_sheets['NF_PLANILHA'], errors='coerce')
        df_sheets['TIPO_PLANILHA'] = df_sheets.iloc[:, 5].fillna('NÃO INFORMADO').astype(str).str.upper().str.strip()
        df_sheets = df_sheets.dropna(subset=['NF_JOIN'])
        
        conn = pyodbc.connect(conn_str)
        query_unificada = """
        SELECT F1_DOC AS NOTA, F1_EMISSAO FROM SF1010 (NOLOCK) WHERE D_E_L_E_T_ = '' AND F1_DTDIGIT >= CONVERT(VARCHAR, GETDATE() - 120, 112)
        UNION ALL
        SELECT F1_DOC AS NOTA, F1_EMISSAO FROM SF1070 (NOLOCK) WHERE D_E_L_E_T_ = '' AND F1_DTDIGIT >= CONVERT(VARCHAR, GETDATE() - 120, 112)
        UNION ALL
        SELECT F1_DOC AS NOTA, F1_EMISSAO FROM SF1080 (NOLOCK) WHERE D_E_L_E_T_ = '' AND F1_DTDIGIT >= CONVERT(VARCHAR, GETDATE() - 120, 112)
        UNION ALL
        SELECT F1_DOC AS NOTA, F1_EMISSAO FROM SF1090 (NOLOCK) WHERE D_E_L_E_T_ = '' AND F1_DTDIGIT >= CONVERT(VARCHAR, GETDATE() - 120, 112)
        """
        df_sql = pd.read_sql(query_unificada, conn)
        conn.close()
        
        df_sql['NOTA'] = pd.to_numeric(df_sql['NOTA'], errors='coerce')
        comparativo = pd.merge(df_sheets, df_sql, left_on='NF_JOIN', right_on='NOTA', how='left')
        
        comparativo['STATUS_FINAL'] = comparativo['NOTA'].apply(
            lambda x: 'Lançada + Recebida' if pd.notnull(x) else 'Pendente no Totvs'
        )
        
        comparativo['DATA_REC_EXIBIR'] = comparativo['dt_obj'].dt.strftime('%d/%m/%Y')
        comparativo['EMISSAO_SQL_EXIBIR'] = pd.to_datetime(comparativo['F1_EMISSAO'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('-')
        comparativo['FORNECEDOR_PLANILHA'] = comparativo.iloc[:, 3].fillna('-')

        tabelas_limpas = comparativo.to_dict(orient='records')
        return render_template('indexnf.html', tabelas=tabelas_limpas)

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return f"Erro: {str(e)}"
    
@app.route('/pedido/detalhes_qr/<id_pedido>', methods=['GET'])
def detalhes_qr_final(id_pedido):
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        
        query_pv = """
        SELECT 
            TRIM(SC5.C5_NUM) AS PEDIDO, 
            TRIM(SA1.A1_NOME) AS CLIENTE, 
            TRIM(SA3.A3_NOME) AS VENDEDOR,
            CASE 
                WHEN SC5.C5_TRANSP = '000001' THEN 'RETIRA' 
                WHEN SC5.C5_TRANSP = '000005' THEN 'NOSSO CARRO'
                ELSE 'TRANSPORTADORA' 
            END AS TRANSPORTE
        FROM SC5010 SC5 (NOLOCK)
        INNER JOIN SA1010 SA1 (NOLOCK) ON SC5.C5_CLIENTE = SA1.A1_COD AND SC5.C5_LOJACLI = SA1.A1_LOJA
        INNER JOIN SA3010 SA3 (NOLOCK) ON SC5.C5_VEND1 = SA3.A3_COD
        WHERE SC5.C5_NUM = ? AND SC5.D_E_L_E_T_ = ''
        """
        cursor.execute(query_pv, (id_pedido,))
        row_p = cursor.fetchone()
        
        cursor.execute("SELECT * FROM dbo.FN_0101_VEN_EXPEDICAO('', '', '') WHERE PEDIDO = ?", (id_pedido,))
        row_e = cursor.fetchone()

        if row_p:
            dt_final = "AGUARDANDO"
            protocolo_final = "N/A"
            nf_exibicao = "Pendente"
            emissao_exibicao = "-"
            nome_transportadora = "--"

            if row_e:
                val_12 = str(row_e[12]).strip() if row_e[12] else ""
                val_15 = str(row_e[15]).strip() if len(row_e) > 15 and row_e[15] else ""
                dt_final = val_12 if val_12 and val_12.lower() != "none" else val_15
                
                protocolo_final = str(row_e[13]).strip() if len(row_e) > 13 else "N/A"
                nf_exibicao = str(row_e[3]).strip()
                emissao_exibicao = row_e[2].strftime('%d/%m/%y') if row_e[2] else "-"

                if len(row_e) > 7:
                    val_7 = str(row_e[7]).strip()
                    if val_7 and val_7.lower() != "none" and val_7 != "":
                        nome_transportadora = val_7

            tipo_transporte = str(row_p[3]).strip() if row_p[3] else "TRANSPORTADORA"

            data = {
                "pedido":      row_p[0],
                "cliente":     row_p[1],
                "vendedor":    row_p[2],
                "tipo_transp": tipo_transporte,
                "nome_transp": nome_transportadora,
                "nf":          nf_exibicao,
                "emissao_nf":  emissao_exibicao,
                "data_col":    dt_final if dt_final else "AGUARDANDO",
                "protocolo":   protocolo_final
            }
            return jsonify(data)
        
        return jsonify({"erro": "Pedido não encontrado"}), 404

    except Exception as e:
        print(f"Erro no servidor: {e}")
        return jsonify({"erro": str(e)}), 500
    finally:
        if 'conn' in locals():
            conn.close()

def formatar_data_protheus(data_raw):
    if not data_raw or str(data_raw).strip() in ["None", ""]:
        return "-"
    d_str = str(data_raw).strip()
    if len(d_str) == 8 and d_str.isdigit():
        return f"{d_str[6:8]}/{d_str[4:6]}/{d_str[0:4]}"
    try:
        return data_raw.strftime('%d/%m/%Y')
    except:
        return d_str            

@app.route('/pedido/<termo_busca>', methods=['GET'])
def get_pedido(termo_busca):
    try:
        import datetime
        import openpyxl
        
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        
        termo_limpo = termo_busca.strip().upper()
        termo_like = f"%{termo_limpo}%"
        
        sql_base = """
        SELECT 
            SC5.C5_FILIAL, 
            TRIM(SC5.C5_NUM) AS PEDIDO, 
            TRIM(SA1.A1_NOME) AS CLIENTE, 
            SC5.C5_EMISSAO, 
            TRIM(SA3.A3_NOME) AS VENDEDOR, 
            TRIM(SC5.C5_TRANSP) AS C5_TRANSP,
            ISNULL(TRIM(SA4.A4_NOME), '') AS NOME_TRANSP,
            EXPED.*
        FROM SC5010 SC5 (NOLOCK)
        INNER JOIN SA1010 SA1 (NOLOCK) ON SC5.C5_CLIENTE = SA1.A1_COD 
        INNER JOIN SA3010 SA3 (NOLOCK) ON SC5.C5_VEND1 = SA3.A3_COD
        LEFT JOIN SA4010 SA4 (NOLOCK) ON SC5.C5_TRANSP = SA4.A4_COD AND SA4.D_E_L_E_T_ = ''
        OUTER APPLY (
            SELECT * FROM dbo.FN_0101_VEN_EXPEDICAO('', '', '') 
            WHERE PEDIDO = SC5.C5_NUM
        ) AS EXPED
        WHERE SC5.D_E_L_E_T_ = ''
        """

        if len(termo_limpo) <= 6 and any(char.isdigit() for char in termo_limpo) and any(char.isalpha() for char in termo_limpo):
            query = sql_base + " AND SC5.C5_NUM = ? ORDER BY SC5.C5_EMISSAO DESC"
            parametros = [termo_limpo]
        else:
            data_limite = (datetime.date.today() - datetime.timedelta(days=365*2)).strftime('%Y%m%d')
            query = sql_base + """ 
                AND SC5.C5_EMISSAO >= ? 
                AND (SA1.A1_NOME LIKE ? OR SA3.A3_NOME LIKE ? OR SA4.A4_NOME LIKE ? OR SC5.C5_TRANSP LIKE ? OR SC5.C5_NUM LIKE ?)
                ORDER BY SC5.C5_EMISSAO DESC
            """
            parametros = [data_limite, termo_like, termo_like, termo_like, termo_like, termo_like]
        
        cursor.execute(query, parametros)
        rows = cursor.fetchall()
        
        if not rows:
            return jsonify({"erro": "Nenhum pedido ou cliente localizado"}), 404

        resultados = []

        for row in rows:
            dt_final = "AGUARDANDO"
            nome_transp_exp = ""
            nf_num = "Pendente"
            nf_data = "-"

            if len(row) > 7 and row[10]: 
                val_12 = str(row[19]).strip() if len(row) > 19 and row[19] else ""
                val_15 = str(row[22]).strip() if len(row) > 22 and row[22] else ""
                raw_col = val_12 if val_12 and val_12.lower() != "none" else val_15
                dt_final = formatar_data_protheus(raw_col)
                
                nome_transp_exp = str(row[14]).strip() if len(row) > 14 and row[14] else ""
                if nome_transp_exp == "--" or nome_transp_exp.lower() == "none":
                    nome_transp_exp = ""
                nf_num = str(row[10]).strip()
                nf_data = formatar_data_protheus(row[9])

            codigo_transp = row[5]
            nome_transp_protheus = row[6]
            
            nome_exibicao = nome_transp_protheus if nome_transp_protheus else nome_transp_exp
            if not nome_exibicao:
                nome_exibicao = "NOSSO CARRO" if not codigo_transp else f"TRANSPORTE {codigo_transp}"

            texto_transp_final = f"{nome_exibicao} | {codigo_transp}" if codigo_transp else f"{nome_exibicao} | --"

            pedido_data = {
                "empresa":     row[0],
                "pedido":      row[1],
                "cliente":     row[2],
                "emissao_ped": formatar_data_protheus(row[3]),
                "vendedor":    row[4],
                "transp":      texto_transp_final,
                "C5_TRANSP":   texto_transp_final,
                "data_col":    dt_final if dt_final else "AGUARDANDO",
                "nf":          nf_num,
                "emissao_nf":  nf_data,
                "producao":    "-",
                "localizacao": "---"
            }
            resultados.append(pedido_data)

     
        caminho_json = 'caminho/para/PedidosLoc.json'
        mapa_loc = {}
        
        if os.path.exists(caminho_json):
            try:
                import json
                with open(caminho_json, 'r', encoding='utf-8') as f:
                    mapa_loc = json.load(f)
            except Exception as e:
                print(f"Erro rápido ao ler JSON de Localizações: {e}")

        for p in resultados:
            cod_sistema = str(p.get('pedido', '')).strip().upper()
            
            if cod_sistema in mapa_loc:
                dados_json = mapa_loc[cod_sistema]
                loc_salva = str(dados_json.get('localizacao', '---')).strip()
                p['localizacao'] = loc_salva if loc_salva and loc_salva != '-' else '---'
                
                conf_salvo = str(dados_json.get('conferente', '---')).strip()
                p['conferente'] = conf_salvo if conf_salvo and conf_salvo != '-' else '---'
                
                data_salva = str(dados_json.get('data', '')).strip()
                hora_salva = str(dados_json.get('hora', '')).strip()
                if data_salva and data_salva != '-':
                    p['data_loc'] = f"{data_salva} às {hora_salva}"
                else:
                    p['data_loc'] = '---'
            else:
                p['localizacao'] = "---"
                p['conferente'] = "---"
                p['data_loc'] = "---"
        
        return jsonify(resultados)

    except Exception as e:
        print(f"Erro no servidor: {e}")
        return jsonify({"erro": str(e)}), 500
    finally:
        if 'conn' in locals(): conn.close()

@app.route('/api/finalizar_pedido', methods=['POST'])
def finalizar_pedido():
    try:
        import json
        import os
        
        dados = request.json
       
        caminho_json = 'caminho/para/PedidosLoc.json'
        
        pedido_num = str(dados.get('pedido', '')).strip().upper()
        
        if not pedido_num or pedido_num == '-':
            return jsonify({"erro": "Número de pedido inválido."}), 400

        banco_dados = {}
        if os.path.exists(caminho_json):
            try:
                with open(caminho_json, 'r', encoding='utf-8') as f:
                    banco_dados = json.load(f)
            except json.JSONDecodeError:
                banco_dados = {}

        banco_dados[pedido_num] = {
            "localizacao": dados.get('localizacao', '-'),
            "conferente": dados.get('conferente', '-'),
            "data": dados.get('data', '-'),
            "hora": dados.get('hora', '-'),
            "transportadora": dados.get('transportadora', '-')
        }

        with open(caminho_json, 'w', encoding='utf-8') as f:
            json.dump(banco_dados, f, ensure_ascii=False, indent=4)
            
        return jsonify({"status": "sucesso", "mensagem": "Dados salvos no JSON com sucesso!"}), 200

    except Exception as e:
        print(f"Erro ao salvar JSON: {str(e)}")
        return jsonify({"erro": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)