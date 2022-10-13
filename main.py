from PyQt5 import uic, QtWidgets
import sys
import pandas as pd
import openpyxl

app = QtWidgets.QApplication([])
planilha = pd.read_excel('planilha/Aeronaves_Drones_cadastrados.xlsx')
caminho = 'planilha/Aeronaves_Drones_cadastrados.xlsx'


primeira_tela = uic.loadUi('janelas/tela_inicial.ui')
tela_crud = uic.loadUi('janelas/tela_crud.ui')

tela_adicionar = uic.loadUi('janelas/tela_adicionar.ui')
tela_atualizar = uic.loadUi('janelas/tela_atualizar.ui')
tela_consultar = uic.loadUi('janelas/tela_consultar.ui')
tela_deletar = uic.loadUi('janelas/tela_deletar.ui')
tela_deletar_coluna = uic.loadUi('janelas/tela_deletar_coluna.ui')

tela_escolha_deletar = uic.loadUi('janelas/tela_escolha_deletar.ui')
tela_erro = uic.loadUi('janelas/tela_erro.ui')
tela_concluido = uic.loadUi('janelas/tela_concluido.ui')
tela_conf_del = uic.loadUi('janelas/tela_conf_del.ui')
tela_adicionado_sucesso = uic.loadUi('janelas/adicionado_sucesso.ui')


#ABRIR TELAS
def btn_tela_crud():
    primeira_tela.close()
    tela_crud.show()

def btn_tela_adicionar():
    tela_crud.close()
    tela_adicionar.show()

def btn_tela_consultar():
    tela_crud.close()
    tela_consultar.show()

def btn_tela_deletar():
    tela_crud.close()
    tela_escolha_deletar.close()
    tela_deletar.show()

def btn_tela_atualizar():
    tela_crud.close()
    tela_atualizar.show()

def btn_tela_escolha_deletar():
    tela_crud.close()
    tela_escolha_deletar.show()

#=============================================================

def btn_adicionar():
    codigo_aeronave = tela_adicionar.line_cod_aero.text()
    data_validade = tela_adicionar.line_data_valid.text()
    operador = tela_adicionar.line_operador.text()
    cpf_cnpj = tela_adicionar.line_doc.text()
    tipo_uso = tela_adicionar.line_tipo.text()
    fabricante = tela_adicionar.line_fabricante.text()
    modelo = tela_adicionar.line_modelo.text()
    numero_serie = tela_adicionar.line_num_serie.text()
    peso_max = tela_adicionar.line_peso_max.text()
    cidade_est = tela_adicionar.line_local.text()
    ramo_ativ = tela_adicionar.line_ramo_ativ.text()

    planilha.loc[len(planilha) + 1] = [codigo_aeronave, data_validade, operador, cpf_cnpj, tipo_uso, fabricante, modelo, numero_serie, cidade_est, ramo_ativ]
    planilha.to_excel(caminho, index = False)

    tela_adicionar.line_cod_aero.clear()
    tela_adicionar.line_data_valid.clear()
    tela_adicionar.line_operador.clear()
    tela_adicionar.line_doc.clear()
    tela_adicionar.line_tipo.clear()
    tela_adicionar.line_fabricante.clear()
    tela_adicionar.line_modelo.clear()
    tela_adicionar.line_num_serie.clear()
    tela_adicionar.line_peso_max.clear()
    tela_adicionar.line_local.clear()
    tela_adicionar.line_ramo_ativ.clear()
    tela_adicionado_sucesso.show()

def btn_consultar():
    item_selecionado = tela_consultar.comboBox_procura.currentText()
    procura = tela_consultar.line_procura.text()
    resultados = planilha.loc[planilha[item_selecionado] == procura]
    print(item_selecionado)
    print(procura)

    tela_consultar.table_resultados.setRowCount(len(resultados['Codigo Aeronave']))

    colunas = ['Codigo Aeronave', 'Data Validade', 'Operador', 'CPF/CNPJ', 'Tipo Uso', 'Fabricante', 'Modelo', 'Numero de serie', 'Cidade-Estado', 'Ramo de atividade']
    
    colun = 0
    for i in colunas:
        linha = 0
        for j in resultados[i]:
            j = str(j)
            tela_consultar.table_resultados.setItem(linha, colun, QtWidgets.QTableWidgetItem(j))
            linha = linha + 1
        colun = colun + 1

def deletar_coluna():
    tela_escolha_deletar.close()
    tela_deletar_coluna.show()

    coluna_selecionada = tela_deletar_coluna.escolha.currentText()

def btn_deletar():
    planilha = pd.read_excel('planilha/Aeronaves_Drones_cadastrados.xlsx')
    item_selecionado = tela_deletar.comboBox_procura.currentText()
    procura = tela_deletar.line_procura.text()
    resultados = planilha.loc[planilha[item_selecionado] == procura]
    index_value = planilha.index[planilha[item_selecionado] == procura].tolist()
    print(item_selecionado)
    print(procura)

    tela_deletar.table_resultados.setRowCount(len(resultados['Codigo Aeronave']))

    colunas = ['Codigo Aeronave', 'Data Validade', 'Operador', 'CPF/CNPJ', 'Tipo Uso', 'Fabricante', 'Modelo', 'Numero de serie', 'Cidade-Estado', 'Ramo de atividade']
    
    colun = 0
    for i in colunas:
        linha = 0
        for j in resultados[i]:
            j = str(j)
            tela_deletar.table_resultados.setItem(linha, colun, QtWidgets.QTableWidgetItem(j))
            linha = linha + 1

        colun = colun + 1


    planilha = planilha.drop(index_value)
    planilha.to_excel(caminho, index = False)

def btn_retornar():
    tela_adicionar.close()
    tela_atualizar.close()
    tela_consultar.close()
    tela_consultar.table_resultados.clear()
    tela_consultar.line_procura.clear()
    tela_deletar.close()
    tela_deletar_coluna.close()
    tela_crud.show()
    tela_conf_del.close()

#PROPRIEDADES TELA INICIAL
primeira_tela.Button_acesso.clicked.connect(btn_tela_crud)

#PROPRIEDADES TELA CRUD
tela_crud.Button_adicionar.clicked.connect(btn_tela_adicionar)
tela_crud.Button_consultar.clicked.connect(btn_tela_consultar)
tela_crud.Button_deletar.clicked.connect(btn_tela_escolha_deletar)
tela_crud.Button_atualizar.clicked.connect(btn_tela_atualizar)

#PROPRIEDADES TELA ADICIONAR
tela_adicionar.Button_retornar.clicked.connect(btn_retornar)
tela_adicionar.Button_adicionar.clicked.connect(btn_adicionar)

#PROPRIEDADES TELA CONSULTAR
tela_consultar.Button_retornar.clicked.connect(btn_retornar)
tela_consultar.Button_consultar.clicked.connect(btn_consultar)
tela_consultar.comboBox_procura.addItem('Codigo Aeronave')
tela_consultar.comboBox_procura.addItem('Data Validade')
tela_consultar.comboBox_procura.addItem('Operador')
tela_consultar.comboBox_procura.addItem('CPF/CNPJ')
tela_consultar.comboBox_procura.addItem('Tipo Uso')
tela_consultar.comboBox_procura.addItem('Fabricante')
tela_consultar.comboBox_procura.addItem('Modelo')
tela_consultar.comboBox_procura.addItem('Numero de serie')
tela_consultar.comboBox_procura.addItem('Cidade-Estado')
tela_consultar.comboBox_procura.addItem('Ramo de atividade')

#PROPRIEDADES TELA ESCOLHA DE DELETAR
tela_escolha_deletar.Button_coluna.clicked.connect(deletar_coluna)
tela_escolha_deletar.Button_linha.clicked.connect(btn_tela_deletar)

#PROPRIEDADES TELA DELETAR COLUNA
tela_deletar_coluna.escolha.addItem('Codigo Aeronave')
tela_deletar_coluna.escolha.addItem('Data Validade')
tela_deletar_coluna.escolha.addItem('Operador')
tela_deletar_coluna.escolha.addItem('CPF/CNPJ')
tela_deletar_coluna.escolha.addItem('Tipo Uso')
tela_deletar_coluna.escolha.addItem('Fabricante')
tela_deletar_coluna.escolha.addItem('Modelo')
tela_deletar_coluna.escolha.addItem('Numero de serie')
tela_deletar_coluna.escolha.addItem('Cidade-Estado')
tela_deletar_coluna.escolha.addItem('Ramo de atividade')

tela_deletar_coluna.Button_retornar.clicked.connect(btn_retornar)
#tela_deletar_coluna.Button_deletar.clicked.connect)

#PROPRIEDADES TELA DELETAR
tela_deletar.Button_retornar.clicked.connect(btn_retornar)
tela_deletar.Button_deletar.clicked.connect(btn_deletar)
tela_deletar.comboBox_procura.addItem('Codigo Aeronave')
tela_deletar.comboBox_procura.addItem('Data Validade')
tela_deletar.comboBox_procura.addItem('Operador')
tela_deletar.comboBox_procura.addItem('CPF/CNPJ')
tela_deletar.comboBox_procura.addItem('Tipo Uso')
tela_deletar.comboBox_procura.addItem('Fabricante')
tela_deletar.comboBox_procura.addItem('Modelo')
tela_deletar.comboBox_procura.addItem('Numero de serie')
tela_deletar.comboBox_procura.addItem('Cidade-Estado')
tela_deletar.comboBox_procura.addItem('Ramo de atividade')

#PROPRIEDADES TELA CONFIRMAR DELETAR
tela_conf_del.Button_cancelar.clicked.connect(btn_retornar)
#tela_conf_del.Button_deletar.clicked.connect()

#PROPRIEDADES TELA ATUALIZAR
tela_atualizar.Button_retornar.clicked.connect(btn_retornar)


primeira_tela.show()
app.exec()