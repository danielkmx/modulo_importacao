import xlrd
import requests
import json


def imprimir_lista_formularios(u_r_l, u_s_e_r, p_a_s_s_w_o_r_d):
    r = requests.get(u_r_l, auth=(u_s_e_r, p_a_s_s_w_o_r_d))
    lista = list()
    lista_content = json.loads(r.content)
    for elemento in lista_content:
        lista.append(elemento['url'])
    return lista


def imprimir_dados_formularios(lista_urls, u_s_e_r, p_a_s_s_w_o_r_d):
    lista = list()
    for formulario_url in lista_urls:
        r = requests.get(formulario_url, auth=(u_s_e_r, p_a_s_s_w_o_r_d))
        formulario_json = json.loads(r.content)
        lista.append(formulario_json)
    return lista


def importar_xls_para_lista_filtrada(u_r_l, u_s_e_r, p_a_s_s_w_o_r_d):
    req = requests.get(u_r_l, auth=(u_s_e_r, p_a_s_s_w_o_r_d))
    with open('formulario.xls', 'wb') as output:
        output.write(req.content)
    workbook = xlrd.open_workbook('formulario.xls')

    worksheet = workbook.sheet_by_index(0)
    total_linhas = worksheet.nrows
    total_colunas = worksheet.ncols

    table = list()
    record = dict()

    for x in range(1, total_linhas):
        for y in range(total_colunas):
            if worksheet.cell(x, y).value:
                if '/' in worksheet.cell(0, y).value:
                    a = worksheet.cell(0, y).value.split('/')
                    record[a[len(a) - 1]] = worksheet.cell(x, y).value
                else:
                    record[worksheet.cell(0, y).value] = worksheet.cell(x, y).value
        if record:
            table.append(record)
    return table

print(teste)