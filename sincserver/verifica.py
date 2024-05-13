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


first = True
def verificar():

    while True:
        print('verificando respostas')
   
        


        lxlsx = [encontrar_ultimo_arquivo('c:\\sincserver\\driveautom\\', 'ColetaGeralResponses01'), encontrar_ultimo_arquivo('c:\\sincserver\\driveautom\\', 'ColetaGeralResponses02'), encontrar_ultimo_arquivo('c:\\sincserver\\driveautom\\', 'ColetaGeralResponses03')]
        novo = pd.DataFrame()
        for i in lxlsx:
            xlsx = pd.ExcelFile(i)
            for sheet_name in xlsx.sheet_names:
                df = pd.read_excel(xlsx, sheet_name=sheet_name)
                df = df = df.dropna(how='all', axis=1)
                novo = pd.concat([novo, df], ignore_index=True)

        antigo = pd.read_excel("c:\\sincserver\\gdrive\\Leads-Info\\!Geral\\alldados.xlsx")

        if not novo.equals(antigo) or first:
            first = False
            antigo = antigo.sort_values(antigo.columns.tolist()).reset_index(drop=True)
            novo = novo.sort_values(novo.columns.tolist()).reset_index(drop=True)

            diferentes = pd.concat([antigo, novo]).drop_duplicates(keep=False)
            if not diferentes.empty:
                for Corretor in diferentes['Corretor']:
                    origem = f'c:\\sincserver\\gdrive\\Leads-Info\\{Corretor}'
                    destino = f'c:\\sincserver\\gdrive\\Leads-Info\\{Corretor}\\respondido'
                    for telefone in diferentes['Contato']:
                        for arquivo in os.listdir(origem):
                            if str(telefone) in arquivo and not 'Classificação de leads' in arquivo:
                                shutil.move(os.path.join(origem, arquivo), os.path.join(destino, arquivo))

                novo.to_excel("c:\\sincserver\\gdrive\\Leads-Info\\!Geral\\alldados.xlsx")

                book = load_workbook('c:\\sincserver\\gdrive\\Leads-Info\\!Geral\\alldados.xlsx')

                sheet = book.active

                for column_cells in sheet.columns:
                    length = max(len(str(cell.value)) for cell in column_cells)
                    sheet.column_dimensions[column_cells[0].column_letter].width = length+5

                book.save('c:\\sincserver\\gdrive\\Leads-Info\\!Geral\\DadosOrganizados.xlsx')
        time.sleep(1)


verificar()