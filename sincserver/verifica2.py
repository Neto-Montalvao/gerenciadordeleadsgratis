import os
import pandas as pd
from openpyxl import load_workbook
import shutil
import time


def encontrar_ultimo_arquivo(diretorio, parte):
    arquivos_xlsx = [f for f in os.listdir(diretorio) if (f.endswith('.xlsx') or f.endswith('.xls')) and parte in f]
    caminhos = [os.path.join(diretorio, basename) for basename in arquivos_xlsx]
    ultimo = max(caminhos, key=os.path.getmtime)
    print(ultimo)
    deletar_arquivos(diretorio, parte)
    return ultimo

def deletar_arquivos(diretorio, parte):
    arquivos_xlsx = [f for f in os.listdir(diretorio) if (f.endswith('.xlsx') or f.endswith('.xls')) and parte in f]
    
    caminhos = [(os.path.join(diretorio, basename), os.path.getmtime(os.path.join(diretorio, basename))) for basename in arquivos_xlsx]
    
    caminhos.sort(key=lambda x: x[1])
    
    if len(caminhos) > 2:
        for i in range(len(caminhos) - 2):
            try:
                os.remove(caminhos[i][0])
            except:
                None


def verificar():
        
    while True:
        print('verificando respostas2')

        lxlsx = [encontrar_ultimo_arquivo('c:\\sincserver\\driveautom\\', 'ColetaGeralResponses04'), encontrar_ultimo_arquivo('c:\\sincserver\\driveautom\\', 'ColetaGeralResponses05')]
        novo = pd.DataFrame()
        for i in lxlsx:
            xlsx = pd.ExcelFile(i)
            for sheet_name in xlsx.sheet_names:
                df = pd.read_excel(xlsx, sheet_name=sheet_name)
                df = df = df.dropna(how='all', axis=1)
                novo = pd.concat([novo, df], ignore_index=True)

        antigo = pd.read_excel("c:\\sincserver\\gdrive\\Leads-Info\\!Geral\\alldados2.xlsx")

        if not novo.equals(antigo):
            first = False
            antigo = antigo.sort_values(antigo.columns.tolist()).reset_index(drop=True)
            novo = novo.sort_values(novo.columns.tolist()).reset_index(drop=True)

            diferentes = pd.concat([antigo, novo]).drop_duplicates(keep=False)
            if not diferentes.empty:
                for i, r in diferentes.iterrows():
                        Corretor = r['Corretor']
                        if str(Corretor) in "['Paula', 'Rosana', 'Carol', 'Generozo', 'Joao', 'Luiza']":
                            origem = f'c:\\sincserver\\gdrive\\Leads-Info\\{Corretor}'
                            destino = f'c:\\sincserver\\gdrive\\Leads-Info\\{Corretor}\\respondido'
                            for arquivo in os.listdir(origem):
                                if str(r['Contato']).strip().split('.')[0] in arquivo:
                                    print(str(r['Contato']).strip().split('.')[0])
                                    shutil.move(os.path.join(origem, arquivo), os.path.join(destino, arquivo))
                        else:
                            print(Corretor)
                            print(r)
                novo.to_excel("c:\\sincserver\\gdrive\\Leads-Info\\!Geral\\alldados2.xlsx")

                book = load_workbook('c:\\sincserver\\gdrive\\Leads-Info\\!Geral\\alldados2.xlsx')

                sheet = book.active

                for column_cells in sheet.columns:
                    length = max(len(str(cell.value)) for cell in column_cells)
                    sheet.column_dimensions[column_cells[0].column_letter].width = length+5

                book.save('c:\\sincserver\\gdrive\\Leads-Info\\!Geral\\DadosOrganizados2.xlsx')
        time.sleep(80)


verificar()