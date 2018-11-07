import xlrd
import requests
import json
import os
import xlwt
import urllib
import pymongo

lista_content = dict()


client = pymongo.MongoClient('mongodb://tetorj:teto2015@ds141613.mlab.com:41613/heroku_twgwkvxl')
db = client['heroku_twgwkvxl']

def imprimir_lista_formularios(u_r_l, u_s_e_r, p_a_s_s_w_o_r_d):
    r = requests.get('https://' + u_r_l + '/api/v1/data', auth=(u_s_e_r, p_a_s_s_w_o_r_d))
    lista_content = json.loads(r.content)
    return lista_content

def retorna_respostas_com_labels(u_r_l, i_d, u_s_e_r, p_a_s_s_w_o_r_d):
    formulario = retorna_lista_com_labels(u_r_l, i_d, u_s_e_r, p_a_s_s_w_o_r_d)
    nome_enquete = retorna_nome_enquete(u_r_l, i_d, u_s_e_r, p_a_s_s_w_o_r_d)
    if u_r_l == 'kobocat.docker.kobo.techo.org':
        u_r_l = 'koboform.docker.kobo.techo.org'
    if u_r_l == 'kc.humanitarianresponse.info':
            u_r_l = 'kobo.humanitarianresponse.info'
    req_url = 'https://' + u_r_l + '/assets/' + formulario[1] + '.json'
    req = requests.get(req_url, auth=(u_s_e_r, p_a_s_s_w_o_r_d))
    lista_content = json.loads(req.content)
    lista_perguntas_labels = dict()
    for item in lista_content['content']['survey']:
        respostas_dict = {}
        if 'select_from_list_name' in item:
            for pergunta in lista_content['content']['choices']:
                    if pergunta['list_name'] in item['select_from_list_name']:
                        respostas_dict[pergunta['name']] = pergunta['label'][0]
            lista_perguntas_labels[item['label'][0]] = respostas_dict
    for resposta in formulario[0]:
        del resposta['_index']
        for key, value in lista_perguntas_labels.items():
            for chave_resposta, key_resposta in resposta.items():
                if isinstance(chave_resposta, list):
                    for grupamento in chave_resposta:
                        del grupamento['_parent_index']
                if key in chave_resposta:
                    opcoes = str(key_resposta).split(' ')
                    if len(opcoes) > 1:
                        resposta[chave_resposta] = ''
                        for opcao in opcoes:
                                if opcao in value:
                                    resposta[chave_resposta] = ' '.join([resposta[chave_resposta],value[opcao]])
                    else:
                        if key in chave_resposta:
                            if key_resposta in value:
                                chave_resposta.replace('.','')
                                resposta[chave_resposta] = value[key_resposta]
    insert_enquetes_no_mongo(formulario[0],nome_enquete)
    return formulario[0]

def insert_enquetes_no_mongo(colection,col_name):
    colection = troca_ponto_por_barra(colection)
    collist = db.list_collection_names()
    if col_name in collist:
        db.drop_collection(col_name)
    db[col_name].insert_many(colection)
def retorna_key_do_formulario(u_r_l, id, u_s_e_r, p_a_s_s_w_o_r_d):
    req = requests.get(u_r_l, auth=(u_s_e_r, p_a_s_s_w_o_r_d))
    lista_content_id = json.loads(req.content)
    for item in lista_content_id:
        if item['id'] == id:
            return item['id_string']

def retorna_lista_com_labels(u_r_l, i_d, u_s_e_r, p_a_s_s_w_o_r_d):
    global lista_content
    id_string = retorna_key_do_formulario('https://' + u_r_l + '/api/v1/data', i_d, u_s_e_r, p_a_s_s_w_o_r_d)

    if u_r_l == 'kobocat.docker.kobo.techo.org':
        u_r_l_novo = 'koboform.docker.kobo.techo.org'
    if u_r_l == 'kc.humanitarianresponse.info':
        u_r_l_novo= 'kobo.humanitarianresponse.info'
    req = requests.get('https://' + u_r_l_novo + '/assets/' + id_string + '.json', auth=(u_s_e_r, p_a_s_s_w_o_r_d))
    lista_content = json.loads(req.content)
    formulario = importar_xls_grupamento_para_lista('https://' + u_r_l + '/api/v1/data/' + str(i_d) + '.xls', u_s_e_r, p_a_s_s_w_o_r_d,i_d)
    labels_dict = dict()
    for item in lista_content['content']['survey']:
        if 'label' in item:
            item['$autoname'].replace('.','')
            labels_dict[item['$autoname']] = item['label']

    ret_formulario = list()
    for enquete in formulario:
        aux_enquete = {}
        aux_enquete['_index'] = enquete.pop('_index')
        for key, value in labels_dict.items():
            if key in enquete:
                if isinstance(enquete[key], list):
                    aux_group = []
                    e = enquete[key]
                    for item in e:
                        aux_group_item = dict()
                        aux_group_item['_parent_index'] = item.pop('_parent_index')
                        for k_group, v_group in labels_dict.items():
                           if v_group[0] in item:
                               aux_group_item[v_group[0]] = item[v_group[0]]
                        aux_group.append(aux_group_item)
                    aux_enquete[value[0]]= aux_group
                else:
                    aux_enquete[value[0]] = enquete.pop(key)
        ret_formulario.append(aux_enquete)
    retorno = [ret_formulario, id_string]
    return retorno
def retorna_nome_enquete(u_r_l, id, u_s_e_r, p_a_s_s_w_o_r_d):
    lista_enquetes = imprimir_lista_formularios(u_r_l, u_s_e_r, p_a_s_s_w_o_r_d)
    for enquete in lista_enquetes:
        if enquete['id'] == id:
            return enquete['title']
def filtra__labels_respostas_grupamento(lista_grupamento):
    lista_perguntas_labels = dict()
    for item in lista_content['content']['survey']:
        respostas_dict = {}
        if 'select_from_list_name' in item:
            for pergunta in lista_content['content']['choices']:
                if pergunta['list_name'] in item['select_from_list_name']:
                    respostas_dict[pergunta['name']] = pergunta['label'][0]
            lista_perguntas_labels[item['label'][0]] = respostas_dict
    for resposta in lista_grupamento:
        for key, value in lista_perguntas_labels.items():
                    for chave_resposta,key_resposta in resposta.items():
                        if key in chave_resposta:
                            opcoes = str(key_resposta).split(' ')
                            if len(opcoes) > 1:
                                resposta[chave_resposta] = ''
                                for opcao in opcoes:
                                    if opcao in value:
                                        resposta[chave_resposta] = ' '.join([resposta[chave_resposta], value[opcao]])
                            else:
                                if key_resposta in value:
                                    resposta[chave_resposta] = value[key_resposta]
    return lista_grupamento

def filtra_labels_perguntas_grupamento(lista_grupamento):
    global lista_content
    labels_dict = dict()
    for item in lista_content['content']['survey']:
        if 'label' in item:
            labels_dict[item['$autoname']] = item['label']
    for enquete in lista_grupamento:
                for key, value in labels_dict.items():
                    if key in enquete:
                        value[0] = value[0].replace("${nombre}","")
                        value[0] = value[0].replace("${nome_morador}", "")
                        value[0] = value[0].replace(".", "")
                        enquete[value[0]] = enquete.pop(key)
    return lista_grupamento


def importar_xls_grupamento_para_lista(u_r_l, u_s_e_r, p_a_s_s_w_o_r_d,i_d):
    req = requests.get(u_r_l, auth=(u_s_e_r, p_a_s_s_w_o_r_d))
    file = "formulario.xls"
    with open('formulario.xls', 'wb') as output:
        output.write(req.content)
    workbook = xlrd.open_workbook('formulario.xls')

    main_sheet = workbook.sheet_by_index(0)
    main_linhas = main_sheet.nrows
    main_colunas = main_sheet.ncols
    table = list()

    for x in range(1, main_linhas):
        record = dict()
        for y in range(main_colunas):
            if main_sheet.cell(x, y).value:
                if '/' in main_sheet.cell(0, y).value:
                    a = main_sheet.cell(0, y).value.split('/')
                    record[a[len(a) - 1]] = main_sheet.cell(x, y).value
                else:
                    record[main_sheet.cell(0, y).value] = main_sheet.cell(x, y).value
        if record:
            table.append(record)
    total = 1
    if workbook.nsheets < 2:
        try:
            os.remove(file)
        except OSError as e:
            print("Error %s - %s" % (e.filename, e.strerror))
        return table
    else:
        while total <= workbook.nsheets - 1:
            linhas = workbook.sheet_by_index(total).nrows
            colunas = workbook.sheet_by_index(total).ncols
            table_grupamento = list()


            for x in range(1, linhas):
                record = dict()
                for y in range(colunas):
                    if workbook.sheet_by_index(total).cell(x, y).value:
                        if '/' in workbook.sheet_by_index(total).cell(0, y).value:
                            a = workbook.sheet_by_index(total).cell(0, y).value.split('/')
                            record[a[len(a) - 1]] = workbook.sheet_by_index(total).cell(x, y).value
                        else:
                            record[workbook.sheet_by_index(total).cell(0, y).value] \
                                = workbook.sheet_by_index(total).cell(x, y).value
                if record:
                    table_grupamento.append(record)
            table_grupamento = filtra_labels_perguntas_grupamento(table_grupamento)
            table_grupamento = filtra__labels_respostas_grupamento(table_grupamento)
            for ele in table_grupamento:
                for element in table:
                    if element['_index'] == ele['_parent_index']:
                        key_value = dict()
                        for key, value in ele.items():
                                key_value[key] = value
                        element.setdefault(workbook.sheet_by_index(total).name, [])
                        element[workbook.sheet_by_index(total).name].append(key_value)
            total = total + 1
        return table

def retorna_lista_perguntas(u_r_l, i_d, u_s_e_r, p_a_s_s_w_o_r_d):
    enquetes_respondidas = retorna_respostas_com_labels(u_r_l, i_d, u_s_e_r, p_a_s_s_w_o_r_d)
    dicionario = dict()
    for enquete in enquetes_respondidas:
        for key, value in enquete.items():
            if key not in dicionario:
                dicionario[key] = 0

    return dicionario

def exporta_xls(u_r_l, i_d, u_s_e_r, p_a_s_s_w_o_r_d):
    enquetes = retorna_respostas_com_labels(u_r_l, i_d, u_s_e_r, p_a_s_s_w_o_r_d)
    req = requests.get('https://' + u_r_l + '/api/v1/data/' + str(i_d) + '.xls', auth=(u_s_e_r, p_a_s_s_w_o_r_d))
    file = "formulario.xls"
    with open('formulario.xls', 'wb') as output:
        output.write(req.content)
    workbook = xlrd.open_workbook('formulario.xls')
    wb = xlwt.Workbook()
    for sheet in workbook.sheet_names():
        wb.add_sheet(sheet, cell_overwrite_ok=True)

    all_columns = dict()
    all_columns[0] = list()
    for enquete in enquetes:
        for key, value in enquete.items():
            if key not in all_columns[0]:
                all_columns[0].append(key)

            if isinstance(value, list):
                if not key in all_columns:
                    all_columns[key] = list()

                for item in value:
                    for k,v in item.items():
                        if k not in all_columns[key]:
                            all_columns[key].append(k)

    linha_semgrupamento = 1
    linha_grupamento = 1
    for enquete in enquetes:
        coluna_semgrupamento = 0
        for elemento in all_columns[0]:
            index = 0
            valor_elemento = '?'
            if elemento in enquete:
                valor_elemento = enquete[elemento]

            if isinstance(valor_elemento, list):
                index = index + 1
                wb.active_sheet = index
                sheet = wb.get_sheet(index)

                for el in valor_elemento:
                    coluna_grupamento = 0
                    for col in all_columns[elemento]:
                        sheet.write(0, coluna_grupamento, col)
                        if col in el:
                                sheet.write(linha_grupamento, coluna_grupamento, el[col])
                        else:
                            sheet.write(linha_grupamento, coluna_grupamento, '?')
                        coluna_grupamento = coluna_grupamento + 1
                    linha_grupamento = linha_grupamento + 1
            else:
                sheet = wb.get_sheet(0)
                sheet.write(0, coluna_semgrupamento, elemento)
                sheet.write(linha_semgrupamento, coluna_semgrupamento, valor_elemento)
                coluna_semgrupamento = coluna_semgrupamento + 1

        linha_semgrupamento = linha_semgrupamento + 1

    respostas = retorna_lista_com_labels(u_r_l, i_d, u_s_e_r, p_a_s_s_w_o_r_d)
    lista_prioridades=gera_lista_prioridades(respostas[0])

    wb.add_sheet('lista_prioridades', cell_overwrite_ok=True)
    wb.active_sheet = 2
    sheet = wb.get_sheet(2)
    sheet.write(0, 0, 'id_enquete')
    sheet.write(0, 1, 'valor_prioridade')
    linha_prioridade = 1

    for id_enquete, valor_prioridade in lista_prioridades.items():
        sheet.write(linha_prioridade, 0, id_enquete)
        sheet.write(linha_prioridade, 1, valor_prioridade)
        linha_prioridade = linha_prioridade+1

    wb.save('teste_enquete.xls')

def gera_lista_prioridades(enquetes):

    pesos = {

        'Qual é a situação do terreno que habita?': {'alugado': 1, 'emprestado': 1},

        'Você ou algum membro do lar possui alguma outra casa ou terreno?': {'casa_vazia': -1,'espe_o_n_o_res': -1},

        'Observar e definir o tipo do material predominante do teto da casa.': {'1': 2, '2': 1},

        'Qual característica apresenta o teto da casa?': {
            'com_algumas_fi': 1,
            'com_muitas_fis': 2,
            'sob_aparente_r': 3},

        'Observar e definir o tipo do material predominante das paredes da casa.': {'1': 1,
                                                                                    '2': 1,
                                                                                    '3': 2,
                                                                                    '5': 1},

        'Qual característica apresenta as paredes da casa?': {
            'com_algumas_fi': 1,
            'com_muitas_fis': 2,
            'sob_aparente_r': 3},

        'Observar e definir o tipo do material predominante do piso da casa.': {'1': 1,
                                                                                '2': 1,
                                                                                '5': 1},

        'Qual característica apresenta o piso da casa?': {
            'com_algumas_fi': 1,
            'com_muitas_fis': 2,
            'sob_aparente_r': 3},

        'A casa apresenta algum dos seguintes problemas?': {'1': 1,
                                                            '2': 1,
                                                            '3': 2,
                                                            '4': 2,
                                                            'calor_e_ou_fri': 1,
                                                            'entrada_de_roe': 2},

        'A casa está localizada perto ou em alguma das seguintes áreas?': {'2': 1,
                                                                           '3': 2,
                                                                           '4': 1,
                                                                           '5': 1,
                                                                           '7': 1},

        'Nos últimos 12 (doze) meses, aconteceu alguma das situações na sua casa?': {'1': 1,
                                                                                     '3': 1},

        'Para que a família utilizará a casa do TETO?': {
            'substitui_o_total_da_moradia_a': 1,
            'cozinha': -1},

        '${nome_morador} se considera de qual gênero?': {'Feminino': 1, 'Outro': 1},

        '${nome_morador} tem alguma destas doenças permanentes ou de longa duração?': {'Hipertensão': 1,
                                                                                       'Diabetes': 1,
                                                                                       'Câncer': 1,
                                                                                       'Doenças nos rins': 1,
                                                                                       'Obesidade': 1,
                                                                                       'Depressão': 1,
                                                                                       'HIV/Aids': 1},

        'Nos últimos doze meses, ${nome_morador} teve algum destes problemas respiratórios?': {'Rinite alérgica': 1,
                                                                                            'Asma': 1,
                                                                                            'Bronquite': 1,
                                                                                            'Enfisema pulmonar': 1,
                                                                                            'Tuberculose': 1},

        'De qual deficiência ${nome_morador} é portador?': {'Visual': 1, 'Auditiva': 1, 'Motora': 1, 'Mental': 1},

        '${nome_morador} está grávida ou amamentando atualmente?': {'Está grávida': 2, 'Está amamentando': 1},

        '${nome_morador} está empregado, exerceu alguma atividade remunerada ou negócio próprio nos últimos 3 meses?': {
                                                                                                'Sim, trabalho informal': 1,
                                                                                                'Não': 1},
    }

    select_many = {
        '${nome_morador} tem alguma destas doenças permanentes ou de longa duração?',
        'Nos últimos doze meses, ${nome_morador} teve algum destes problemas respiratórios?',
        'De qual deficiência ${nome_morador} é portador',
        'Observar e definir o tipo do material predominante do teto da casa.',
        'Observar e definir o tipo do material predominante das paredes da casa.',
        'Observar e definir o tipo do material predominante do piso da casa.',
        'A casa apresenta algum dos seguintes problemas?',
        'A casa está localizada perto ou em alguma das seguintes áreas?',
        'Nos últimos 12 (doze) meses, aconteceu alguma das situações na sua casa?'
    }

    totais = {}
    for enquete in enquetes:
        total_enquete=0
        index=enquete['_index']
        num_moradores = 0
        for key, value in enquete.items():
            if key in pesos and value in pesos[key] and key not in select_many:
                total_enquete=total_enquete+pesos[key][value]
            elif key in select_many:
                for choice in value.split():
                    if choice in pesos[key]:
                        total_enquete = total_enquete + pesos[key][choice]

            if isinstance(value,list):
                total_grupo = 0
                for item in value:
                    num_moradores = num_moradores+1
                    for k,v in item.items():
                        if k in pesos and v in pesos[k] and k not in select_many:
                            total_grupo = total_grupo+pesos[k][v]
                        elif k in select_many:
                            for c in v.split():
                                if c in pesos[k]:
                                    total_enquete = total_enquete + pesos[k][c]
                        if k == 'Poderia me dizer qual é a sua renda mensal? (Somando todos os rendimentos incluindo benefícios sociais)':
                            if v < 101:
                                total_grupo = total_grupo + 5
                            elif v < 201:
                                total_grupo = total_grupo + 4
                            elif v < 301:
                                total_grupo = total_grupo + 3
                            elif v < 401:
                                total_grupo = total_grupo + 2
                            elif v < 501:
                                total_grupo = total_grupo + 1
                total_enquete = total_enquete+total_grupo

                if num_moradores>4 and num_moradores<7:
                    total_enquete = total_enquete +1
                elif num_moradores>6:
                    total_enquete = total_enquete+2

        totais[index] = total_enquete

    return totais

def troca_ponto_por_barra(formulario):
    for enquete in formulario:
        for key,value in enquete.items():
            if isinstance(key,str):
                if '.' in key:
                    enquete[key.replace('.', '')] = enquete[key]
                    del enquete[key]
                if '$' in key:
                    enquete[key.replace('$', '')] = enquete[key]
                    del enquete[key]

    return formulario
print(retorna_respostas_com_labels('kc.humanitarianresponse.info',274173,'riodejaneiro','teto2015'))