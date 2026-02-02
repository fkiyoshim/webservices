import pandas as pd
import os
import glob

def convert():
    # Procura na pasta target por qualquer arquivo .xls de resultado
    files = glob.glob("target/RT-WebServices*.xls")
    
    output_dir = 'public'
    if not os.path.exists(output_dir): 
        os.makedirs(output_dir)
    
    output_html = os.path.join(output_dir, 'index.html')

    if files:
        # Pega o arquivo mais recente caso haja mais de um
        input_file = max(files, key=os.path.getctime)
        print(f"Convertendo o arquivo: {input_file}")
        
        try:
            # engine='xlrd' para processar o formato .xls da ferramenta
            df_dict = pd.read_excel(input_file, sheet_name=None, engine='xlrd')
            
            with open(output_html, 'w', encoding='utf-8') as f:
                f.write("<html><head><meta charset='utf-8'><title>Resultados TIM/Vivo</title>")
                f.write("<style>body{font-family:sans-serif; padding:20px;} table{border-collapse:collapse;width:100%;margin-bottom:30px;} th,td{border:1px solid #ddd;padding:8px;text-align:left;} th{background:#004a99;color:white;}</style></head><body>")
                f.write(f"<h1>Relatório de Automação: {os.path.basename(input_file)}</h1>")
                
                for name, sheet in df_dict.items():
                    f.write(f"<h2>Fluxo: {name}</h2>")
                    f.write(sheet.to_html(index=False))
                f.write("</body></html>")
            print("Página HTML gerada com sucesso.")
        except Exception as e:
            with open(output_html, 'w', encoding='utf-8') as f:
                f.write(f"<h1>Erro na conversão: {str(e)}</h1>")
            print(f"Erro ao ler Excel: {e}")
    else:
        with open(output_html, 'w', encoding='utf-8') as f:
            f.write("<h1>Nenhum arquivo de resultado encontrado na pasta target.</h1>")
        print("Erro: Nenhum arquivo RT-WebServices*.xls encontrado.")

if __name__ == "__main__":
    convert()