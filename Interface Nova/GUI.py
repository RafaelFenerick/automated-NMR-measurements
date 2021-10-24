#!/usr/bin/python3
# -*- coding: utf-8 -*-

"""
Interface gráfica para execução do programa
de automatização de experimentos de NMR

"""

import sys, os
from threading import Thread
import multiprocessing as Queue
from time import sleep
from datetime import datetime as timecurrent
#import pythoncom
from VBS_Control import Control
from SerialCommunication import SerialCommunication

from PyQt5.QtCore import QDate, QTime, QDateTime, Qt
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtWidgets import QApplication, QWidget, QToolTip, QPushButton, QMessageBox, QDesktopWidget, \
    QMainWindow, QAction, qApp, QMenu, QTextEdit, QLabel, QHBoxLayout, QVBoxLayout, QGridLayout, QLineEdit, \
    QSizePolicy, QRadioButton, QCheckBox, QFileDialog, QFrame

def write_log(message):
    current_time = timecurrent.now()
    day_to_write = str(current_time.day) + "/" + str(current_time.month) + "/" + str(current_time.year)
    time_to_write = str(current_time.hour) + ":" + str(current_time.minute) + ":" + str(current_time.second) + \
    "." + str(current_time.microsecond)

    info_file = open("log.txt", "a")
    info_file.write(day_to_write + " " + time_to_write + " - " + message + "\n")
    info_file.close()

class GUI(QMainWindow):

    def __init__(self):

        write_log("Início __init__ GUI")

        # Chamar construtor da classe mãe
        super().__init__()

        # Inicialização de variáveis
        self.on_experiment = False
        self.states = []
        self.generalwidgets = []
        self.temperaturewidgets = []

        self.titlefont = QFont("Times", 12, QFont.Bold)
        self.commonfont = QFont("Times-Italic", 10)
        self.commonfont.setItalic(True)
        self.labelsfixedsize = 145

        # Criar interface
        self.initUI()

        write_log("Fim __init__ GUI")

    def check_queue(self):

        if not experiment_data.empty():
            data = experiment_data.get()
            self.on_experiment = False

        self.master.after(100, self.check_queue)

    def initUI(self):
        ''' Construção da interface'''


        write_log("Início initGUI")

        self.states.append(0)

        # --- Ações

        # Ação de novo experimento
        newexperimentAct = QAction(QIcon('icons/new.ico'), 'Novo', self) # Criação de ação
        newexperimentAct.setShortcut('Ctrl+N') # Atalho para ação
        newexperimentAct.setStatusTip('Iniciar novo experimento') # Comentário na barra de status
        newexperimentAct.triggered.connect(self.new_experiment) # Ação para fechar interface

        # Ação de calibração
        calibrationAct = QAction(QIcon('icons/calibration.ico'), 'Calibração', self)  # Criação de ação
        calibrationAct.setShortcut('Ctrl+C')  # Atalho para ação
        calibrationAct.setStatusTip('Iniciar calibração de temperatura da amostra')  # Comentário na barra de status
        calibrationAct.triggered.connect(self.close)  # Ação para fechar interface

        # Ação de adicionar aplicação
        addapplicationAct = QAction(QIcon('icons/addapp.ico'), 'Adicionar', self)  # Criação de ação
        addapplicationAct.setStatusTip('Adicionar nova aplicação')  # Comentário na barra de status
        addapplicationAct.triggered.connect(self.add_application)  # Ação para fechar interface

        # Ação de remover aplicação
        removeapplicationAct = QAction(QIcon('icons/removeapp.ico'), 'Remover', self)  # Criação de ação
        removeapplicationAct.setStatusTip('Remover aplicação')  # Comentário na barra de status
        removeapplicationAct.triggered.connect(self.close)  # Ação para fechar interface

        # Ação de parâmetro de baixa temperatura
        lowtemperatureparamAct = QAction(QIcon('icons/lowtemp.png'), 'Baixa Temperatura', self)  # Criação de ação
        lowtemperatureparamAct.setStatusTip('Configurar parâmetros de baixa temperatura')  # Comentário na barra de status
        lowtemperatureparamAct.triggered.connect(self.close)  # Ação para fechar interface

        # Ação de parâmetro de alta temperatura
        hightemperaturepararmAct = QAction(QIcon('icons/hightemp.png'), 'Alta Temperatura', self)  # Criação de ação
        hightemperaturepararmAct.setStatusTip('Configurar parâmetros de alta temperatura')  # Comentário na barra de status
        hightemperaturepararmAct.triggered.connect(self.close)  # Ação para fechar interface

        # Ação de parâmetro do equipamento
        equipparamAct = QAction(QIcon('icons/equip.ico'), 'Parâmetros do Equipamento', self)  # Criação de ação
        equipparamAct.setStatusTip('Configurar parâmetros do equipamento')  # Comentário na barra de status
        equipparamAct.triggered.connect(self.close)  # Ação para fechar interface

        # Ação de Unidade de Leitura de Arquivo (checkable)
        filereadAct = QAction(QIcon('icons/file.ico'), 'Unidade para Leitura de Arquivo', self)  # Criação de ação
        filereadAct.setStatusTip('Definir unidade de temperaturas para leitura de arquivo')  # Comentário na barra de status
        filereadAct.triggered.connect(self.close)  # Ação para fechar interface

        # Ação de saída
        exitAct = QAction(QIcon('icons/exit.png'), 'Sair', self)  # Criação de ação
        exitAct.setShortcut('Ctrl+Q')  # Atalho para ação
        exitAct.setStatusTip('Finalizar')  # Comentário na barra de status
        exitAct.triggered.connect(self.close)  # Ação para fechar interface

        ###

        self.statusbar = self.statusBar() # Criação da barra de status
        self.statusbar.showMessage('Pronto')
        menubar = self.menuBar() # Criação da barra de menus

        ###

        # --- Criação de menus

        ### File Menu
        fileMenu = menubar.addMenu('&Arquivo')
        fileMenu.addAction(newexperimentAct)
        fileMenu.addAction(calibrationAct)
        # Application SubMenu
        applicationMenu = QMenu('Aplicações', self)  # Criação de menu para adição de submenu
        applicationMenu.addAction(addapplicationAct)
        applicationMenu.addAction(removeapplicationAct)
        fileMenu.addMenu(applicationMenu)
        fileMenu.addAction(exitAct)

        ### Config Menu
        configMenu = menubar.addMenu('&Configurações')
        # Temperatures SubMenu
        temperaturesparamMenu = QMenu('Parâmetros de Temperatura', self)  # Criação de menu para adição de submenu
        temperaturesparamMenu.addAction(lowtemperatureparamAct)
        temperaturesparamMenu.addAction(hightemperaturepararmAct)
        configMenu.addMenu(temperaturesparamMenu)
        configMenu.addAction(equipparamAct)
        configMenu.addAction(filereadAct)

        ### Tool bar
        toolbar = self.addToolBar('Toolbar')
        toolbar.setStyleSheet('QToolBar{spacing:10px;}')
        toolbar.addAction(newexperimentAct)
        toolbar.addAction(calibrationAct)
        toolbar.addAction(addapplicationAct)
        toolbar.addAction(removeapplicationAct)
        toolbar.addAction(lowtemperatureparamAct)
        toolbar.addAction(hightemperaturepararmAct)
        toolbar.addAction(equipparamAct)
        toolbar.addAction(filereadAct)
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        toolbar.addWidget(spacer)
        toolbar.addAction(exitAct)

        # --- Grid Layout
        gridwidget = QWidget()
        self.grid = QGridLayout()
        self.grid.setHorizontalSpacing(50)
        self.grid.setVerticalSpacing(20)
        gridwidget.setLayout(self.grid)
        self.setCentralWidget(gridwidget)

        # --- Window
        self.resize(800, 600)
        self.center()
        self.setWindowTitle('PNMR Control')
        self.setWindowIcon(QIcon('icons/Icone.png'))

        self.show()

        write_log("Fim initGUI")

    def new_experiment(self):
        ''' Configuração de novo experimento '''

        write_log("Início newExperiment")

        if not self.newstateEvent(): return
        self.deletewidgets()
        self.states.append(1)

        # --- Informações de Experimento
        # Label
        label = QLabel("Informações do Experimento")
        label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        label.setAlignment(Qt.AlignCenter)
        label.setFont(self.titlefont)
        self.generalwidgets.append(label)
        self.grid.addWidget(label, 0, 0)

        # Amostra
        label = QLabel("Amostra: ")
        label.setToolTip('Nome da amostra utilizada no experimento')
        label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        label.setFixedWidth(self.labelsfixedsize)
        label.setFont(self.commonfont)
        sublayout = QHBoxLayout()
        sublayout.addWidget(label)
        self.amostraname = QLineEdit()
        self.amostraname.setAlignment(Qt.AlignLeft)
        self.amostraname.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.amostraname.setText("Amostra Teste")
        sublayout.addWidget(self.amostraname)
        self.generalwidgets.append(sublayout)
        self.grid.addLayout(sublayout, 1, 0)

        # Diretório
        label = QLabel("Diretório: ")
        label.setToolTip('Diretório onde salvar arquivo com detalhes do experimento')
        label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        label.setFixedWidth(self.labelsfixedsize)
        label.setFont(self.commonfont)
        sublayout = QHBoxLayout()
        sublayout.addWidget(label)
        self.amostrapath = QLineEdit()
        self.amostrapath.setAlignment(Qt.AlignLeft)
        self.amostrapath.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.amostrapath.setText("C:\\")
        sublayout.addWidget(self.amostrapath)
        chossepathbutton = QPushButton("Escolher")
        chossepathbutton.setFixedWidth(70)
        chossepathbutton.clicked.connect(self.pickdir)
        sublayout.addWidget(chossepathbutton)
        self.generalwidgets.append(sublayout)
        self.grid.addLayout(sublayout, 2, 0)

        # --- Parâmetros de Temperatura
        # Label
        label = QLabel("Parâmentros de Temperatura")
        label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        label.setAlignment(Qt.AlignCenter)
        label.setFont(self.titlefont)
        self.generalwidgets.append(label)
        self.grid.addWidget(label, 3, 0)

        # Escolha de temperatura
        sublayout = QHBoxLayout()
        self.radiobuttonone = QRadioButton("Uma Temperatura")
        self.radiobuttonone.setChecked(True)
        self.radiobuttonone.toggled.connect(lambda: self.temperaturechoose(0))
        sublayout.addWidget(self.radiobuttonone)
        self.radiobuttonmore = QRadioButton("Várias Temperaturas")
        self.radiobuttonmore.toggled.connect(lambda: self.temperaturechoose(1))
        sublayout.addWidget(self.radiobuttonmore)
        self.radiobuttonfile = QRadioButton("Por Arquivo")
        self.radiobuttonfile.toggled.connect(lambda: self.temperaturechoose(2))
        sublayout.addWidget(self.radiobuttonfile)
        self.generalwidgets.append(sublayout)
        self.grid.addLayout(sublayout, 4, 0)

        # Tempo de espera
        label = QLabel("Tempo de Espera [min]: ")
        label.setToolTip('Tempo de espera para estabilização de cada temperatura do experimento')
        label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        label.setFixedWidth(self.labelsfixedsize)
        label.setFont(self.commonfont)
        sublayout = QHBoxLayout()
        sublayout.addWidget(label)
        self.waittime = QLineEdit()
        self.waittime.setAlignment(Qt.AlignLeft)
        self.waittime.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.waittime.setText("10")
        sublayout.addWidget(self.waittime)
        self.generalwidgets.append(sublayout)
        self.grid.addLayout(sublayout, 9, 0)

        # Tempo estimado
        label = QLabel("Tempo Estimado: ")
        label.setToolTip('Tempo estimado de duração do experimento')
        label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        label.setFont(self.commonfont)
        self.generalwidgets.append(label)
        self.grid.addWidget(label, 10, 0)

        # --- Aplicações
        # Label
        label = QLabel("Aplicações")
        label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        label.setAlignment(Qt.AlignCenter)
        label.setFont(self.titlefont)
        self.generalwidgets.append(label)
        self.grid.addWidget(label, 0, 1)

        # Aplicações
        sublayout = QVBoxLayout()
        try:
            file = open("Applications\\Applications.txt", "r")
            self.applications_text = file.readlines()
            print(self.applications_text)
            file.close()
        except:
            self.applications_text = []

        if len(self.applications_text) == 0:
            QMessageBox.question(self, 'Mensagem', "Nenhuma Aplicação Adicionada!", QMessageBox.Ok, QMessageBox.Ok)
            self.deletewidgets()
            return

        self.applications = []
        for application in self.applications_text:
            check = QCheckBox(str(application.strip("\n")))
            sublayout.addWidget(check)
            self.applications.append(check)
        self.generalwidgets.append(sublayout)
        self.grid.addLayout(sublayout, 1, 1, 5, 1)

        # --- Parâmetros de equipamentos
        # Escolha de tipo de temperatura
        sublayout = QHBoxLayout()
        self.radiobuttonhigh = QRadioButton("Altas Temperaturas")
        self.radiobuttonhigh.setChecked(True)
        self.radiobuttonhigh.toggled.connect(lambda: self.temperaturetypechoose(0))
        sublayout.addWidget(self.radiobuttonhigh)
        self.radiobuttonlow = QRadioButton("Baixas Temperaturas")
        self.radiobuttonlow.toggled.connect(lambda: self.temperaturetypechoose(1))
        sublayout.addWidget(self.radiobuttonlow)
        self.generalwidgets.append(sublayout)
        self.grid.addLayout(sublayout, 6, 1)

        # Gas Flow
        self.gasflow_evaporator_label = QLabel("Gas Flow: ")
        self.gasflow_evaporator_label.setToolTip('Potência do Gas Flow / Evaporator')
        self.gasflow_evaporator_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.gasflow_evaporator_label.setFixedWidth(self.labelsfixedsize/2)
        self.gasflow_evaporator_label.setFont(self.commonfont)
        sublayout = QHBoxLayout()
        sublayout.addWidget(self.gasflow_evaporator_label)
        self.gasflow_evaporator = QLineEdit()
        self.gasflow_evaporator.setAlignment(Qt.AlignLeft)
        self.gasflow_evaporator.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.gasflow_evaporator.setText("2000")
        sublayout.addWidget(self.gasflow_evaporator)
        self.generalwidgets.append(sublayout)
        self.grid.addLayout(sublayout, 7, 1)

        # Rampa
        self.rampdo = QCheckBox("Rampa (variações maiores que)")
        sublayout = QHBoxLayout()
        sublayout.addWidget(self.rampdo)
        self.rampdiff = QLineEdit()
        self.rampdiff.setAlignment(Qt.AlignLeft)
        self.rampdiff.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.rampdiff.setText("15")
        sublayout.addWidget(self.rampdiff)
        self.generalwidgets.append(sublayout)
        self.grid.addLayout(sublayout, 8, 1)

        # Tune
        self.tune = QCheckBox("Tune")
        self.generalwidgets.append(self.tune)
        self.grid.addWidget(self.tune, 9, 1)

        # --- Running
        sublayout = QHBoxLayout()
        cancelexperimentbutton = QPushButton("Cancelar")
        cancelexperimentbutton.clicked.connect(self.deletewidgets)
        cancelexperimentbutton.setFixedWidth(120)
        sublayout.addWidget(cancelexperimentbutton)
        startexperimentbutton = QPushButton("Iniciar")
        startexperimentbutton.clicked.connect(lambda: self.start_experiment())
        startexperimentbutton.setFixedWidth(150)
        sublayout.addWidget(startexperimentbutton)
        self.generalwidgets.append(sublayout)
        self.grid.addLayout(sublayout, 10, 1)

        self.temperaturechoose(0)

        self.grid.setRowStretch(0, 1)
        self.grid.setRowStretch(1, 1)
        self.grid.setRowStretch(2, 1)
        self.grid.setRowStretch(3, 1)
        self.grid.setRowStretch(4, 1)
        self.grid.setRowStretch(5, 1)
        self.grid.setRowStretch(6, 2)
        self.grid.setRowStretch(7, 2)
        self.grid.setRowStretch(8, 2)
        self.grid.setRowStretch(9, 2)
        self.grid.setRowStretch(10, 1)

        self.grid.setColumnStretch(0, 2)
        self.grid.setColumnStretch(1, 1)

        write_log("Fim newExperiment")

    def start_experiment(self):
        ''' Definir parâmentros e iniciar experimento '''

        write_log("Início startExperiment")

        # Read Files
        if self.radiobuttonhigh.isChecked():
            file = open("Temp_params.txt", "r")
        else:
            file = open("Temp_params_low.txt", "r")
        lines = file.readlines()
        file.close()
        a, b = float(lines[0].strip("\n")), float(lines[1].strip("\n"))

        file = open("Equip_params.txt", "r")
        lines = file.readlines()
        file.close()
        exe_path, serialnumber = str(lines[0].strip("\n")), str(lines[1].strip("\n"))

        file = open("Filetemp_params.txt", "r")
        lines = file.readlines()
        file.close()
        filetemp = int(lines[0].strip("\n"))

        # Aplicações
        self.to_run_applications = []
        for application, application_text in zip(self.applications, self.applications_text):
            if application.isChecked():
                self.to_run_applications.append(application_text.strip("\n"))

        # --- Temperaturas
        temperatures_BVT = []
        temperatures_amostra = []
        temperatures_amostra_celsius = []

        wait_time = float(self.waittime.text())
        # Por arquivo
        if self.radiobuttonfile.isChecked():
            for temperature in self.file_temperatures_array:
                if filetemp == 1:
                    temperatures_BVT.append(temperature)
                    temperatures_amostra.append(round(float((temperature + b) / a), 2))
                    temperatures_amostra_celsius.append(round(float((temperature + b) / a) - 273.15, 2))
                elif filetemp == 2:
                    temperatures_BVT.append(round(float(temperature * a - b), 2))
                    temperatures_amostra.append(temperature)
                    temperatures_amostra_celsius.append(round(temperature - 273.15, 2))
                else:
                    temperatures_BVT.append(round(float((temperature + 273.15) * a - b), 2))
                    temperatures_amostra.append(round(temperature + 273.15, 2))
                    temperatures_amostra_celsius.append(temperature)

        # Várias
        elif self.radiobuttonmore.isChecked():
            init = float(self.initialtemperature1.text())
            end = float(self.endtemperature1.text())
            step = float(self.steptemperature1.text())
            amount = float(self.steptemperature2.text())
            tt = 0
            if end < init:
                while tt < amount:
                    temperatures_BVT.append(init)
                    temperatures_amostra.append(round(float((init + b) / a), 2))
                    temperatures_amostra_celsius.append(round(float((init + b) / a) - 273.15, 2))
                    init -= step
                    tt += 1
            else:
                while init <= end:
                    temperatures_BVT.append(init)
                    temperatures_amostra.append(round(float((init + b) / a), 2))
                    temperatures_amostra_celsius.append(round(float((init + b) / a) - 273.15, 2))
                    init += step
                    tt += 1
        # Uma
        else:
            if self.roomtemperature.isChecked():
                temperatures_BVT.append(0)
                temperatures_amostra.append(0)
                temperatures_amostra_celsius.append(0)
            else:
                temperatures_BVT.append(float(self.experimenttemperature1.text()))
                temperatures_amostra.append(round(float((float(self.experimenttemperature1.text()) + b) / a), 2))
                temperatures_amostra_celsius.append(
                    round(float((float(self.experimenttemperature1.text()) + b) / a) - 273.15, 2))

        # Rampas
        ramps = [self.rampdo.isChecked(), float(self.rampdiff.text())]

        # Informações do experimento
        amostra = self.amostraname.text()
        to_save_path = self.amostrapath.text()
        if to_save_path[len(to_save_path) - 1] != "\\":
            to_save_path += "\\"

        # Tune
        tune = self.tune.isChecked()

        # Gas Flow / Evaporator
        gas_flow = int(self.gasflow_evaporator.text())

        # Altas / Baixas temperaturas
        low_temperature = self.radiobuttonlow.isChecked()

        # Aplicações
        app_path = "Applications\\"
        applications = []
        for value in self.to_run_applications:
            applications.append(os.path.abspath(app_path + value + ".app"))

        # --- Escrever arquivo com informações do experimento
        current_time = timecurrent.now()
        info_file = open(to_save_path + amostra + " infos.txt", "w")
        info_file.write("Amostra: " + amostra + "\n")
        date = str(current_time.day) + "/" + str(current_time.month) + "/" + str(current_time.year)
        time = str(current_time.hour) + ":" + str(current_time.minute)
        info_file.write(date + " " + time + "\n")
        info_file.write("Aplicações: ")
        tr = 0
        for application in self.to_run_applications:
            if tr == 0:
                tr = 1
                info_file.write(application)
            else:
                info_file.write("," + application)
        info_file.write("\n")
        info_file.write("Temperaturas(BVT): ")
        tr = 0
        for temperature in temperatures_BVT:
            if tr == 0:
                tr = 1
                if temperature != 0:
                    info_file.write(str(temperature))
                else:
                    info_file.write("Ambiente")
            else:
                info_file.write("," + str(temperature))
        info_file.write("\n")
        info_file.write("Temperaturas(amostra): ")
        tr = 0
        for temperature in temperatures_amostra:
            if tr == 0:
                tr = 1
                if temperature != 0:
                    info_file.write(str(temperature))
                else:
                    info_file.write("Ambiente")
            else:
                info_file.write("," + str(temperature))
        info_file.write("\n")
        info_file.write("Temperaturas(amostra) ºC: ")
        tr = 0
        for temperature in temperatures_amostra_celsius:
            if tr == 0:
                tr = 1
                if temperature != 0:
                    info_file.write(str(temperature))
                else:
                    info_file.write("Ambiente")
            else:
                info_file.write("," + str(temperature))

        info_file.write("\n")
        info_file.write("Tune: ")
        if tune:
            info_file.write("Sim")
        else:
            info_file.write("Não")
        info_file.write("\n")
        if low_temperature:
            info_file.write("Evaporator: " + str(gas_flow))
        else:
            info_file.write("Gas Flow: " + str(gas_flow))
        info_file.write("\n")
        info_file.write("Baixas Temperaturas: ")
        if low_temperature:
            info_file.write("Sim")
        else:
            info_file.write("Não")
        info_file.write("\n")
        info_file.close()

        self.on_experiment = True

        write_log("Fim startExperiment")

        gui_data.put([gas_flow, tune, temperatures_BVT, wait_time, applications, ramps, serialnumber, exe_path, low_temperature, 0])

    def add_application(self):
        ''' Configuração de novo experimento '''

        write_log("Início addApplication")

        if not self.newstateEvent(): return
        self.deletewidgets()
        self.states.append(2)

        # --- Informações de Experimento
        # Label
        label = QLabel("Informações da Aplicação")
        label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        label.setAlignment(Qt.AlignCenter)
        label.setFont(self.titlefont)
        self.generalwidgets.append(label)
        self.grid.addWidget(label, 0, 0)

        # Amostra
        label = QLabel("Nome: ")
        label.setToolTip('Nome da aplicação adicionada')
        label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        label.setFixedWidth(self.labelsfixedsize)
        label.setFont(self.commonfont)
        sublayout = QHBoxLayout()
        sublayout.addWidget(label)
        self.applicationname = QLineEdit()
        self.applicationname.setAlignment(Qt.AlignLeft)
        self.applicationname.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.applicationname.setText("Aplicação Teste")
        sublayout.addWidget(self.applicationname)
        self.generalwidgets.append(sublayout)
        self.grid.addLayout(sublayout, 1, 0)

        # Diretório
        label = QLabel("Caminho: ")
        label.setToolTip('Caminho da aplicação adicionada')
        label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        label.setFixedWidth(self.labelsfixedsize)
        label.setFont(self.commonfont)
        sublayout = QHBoxLayout()
        sublayout.addWidget(label)
        self.applicationpath = QLineEdit()
        self.applicationpath.setAlignment(Qt.AlignLeft)
        self.applicationpath.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.applicationpath.setText("C:\\")
        sublayout.addWidget(self.applicationpath)
        chossepathbutton = QPushButton("Escolher")
        chossepathbutton.setFixedWidth(70)
        chossepathbutton.clicked.connect(self.pickfile)
        sublayout.addWidget(chossepathbutton)
        self.generalwidgets.append(sublayout)
        self.grid.addLayout(sublayout, 2, 0)


        # --- Running
        sublayout = QHBoxLayout()
        cancelexperimentbutton = QPushButton("Cancelar")
        cancelexperimentbutton.clicked.connect(self.deletewidgets)
        cancelexperimentbutton.setFixedWidth(120)
        sublayout.addWidget(cancelexperimentbutton)
        startexperimentbutton = QPushButton("Adicionar")
        startexperimentbutton.clicked.connect(self.deletewidgets)
        startexperimentbutton.setFixedWidth(150)
        sublayout.addWidget(startexperimentbutton)
        self.generalwidgets.append(sublayout)
        self.grid.addLayout(sublayout, 10, 1)

        self.grid.setRowStretch(0, 1)
        self.grid.setRowStretch(1, 1)
        self.grid.setRowStretch(2, 1)
        self.grid.setRowStretch(3, 1)
        self.grid.setRowStretch(4, 1)
        self.grid.setRowStretch(5, 1)
        self.grid.setRowStretch(6, 2)
        self.grid.setRowStretch(7, 2)
        self.grid.setRowStretch(8, 2)
        self.grid.setRowStretch(9, 2)
        self.grid.setRowStretch(10, 1)

        self.grid.setColumnStretch(0, 2)
        self.grid.setColumnStretch(1, 1)

        write_log("Fim addApplication")

    ###

    def temperaturechoose(self, value):
        ''' Escolha das temperaturas '''

        self.deletewidgets(1)

        # Uma temperatura
        if value == 0:
            label = QLabel(" ")
            label.setAlignment(Qt.AlignHCenter | Qt.AlignBottom)
            label.setFixedWidth(self.labelsfixedsize)
            label.setFont(self.commonfont)
            sublayout = QHBoxLayout()
            sublayout.setSpacing(5)
            sublayout.addWidget(label)
            label = QLabel("BVT [K]")
            label.setAlignment(Qt.AlignHCenter | Qt.AlignBottom)
            label.setFixedWidth(self.labelsfixedsize)
            label.setFont(self.commonfont)
            sublayout.addWidget(label)
            self.temperatura_amostra_label = QLabel("Amostra [ºC]")
            self.temperatura_amostra_label.setAlignment(Qt.AlignHCenter | Qt.AlignBottom)
            self.temperatura_amostra_label.setFixedWidth(self.labelsfixedsize)
            self.temperatura_amostra_label.setFont(self.commonfont)
            sublayout.addWidget(self.temperatura_amostra_label)
            self.temperaturewidgets.append(sublayout)
            self.generalwidgets.append(sublayout)
            self.grid.addLayout(sublayout, 5, 0)

            # Temperatura
            label = QLabel("Temperatura: ")
            label.setToolTip('Temperatura de realização do experimento')
            label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            label.setFixedWidth(self.labelsfixedsize)
            label.setFont(self.commonfont)
            sublayout = QHBoxLayout()
            sublayout.setSpacing(5)
            sublayout.addWidget(label)
            self.experimenttemperature1 = QLineEdit()
            self.experimenttemperature1.setAlignment(Qt.AlignLeft)
            self.experimenttemperature1.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.experimenttemperature1.setText("320")
            sublayout.addWidget(self.experimenttemperature1)
            self.experimenttemperature2 = QLineEdit()
            self.experimenttemperature2.setAlignment(Qt.AlignLeft)
            self.experimenttemperature2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.experimenttemperature2.setText("46")
            sublayout.addWidget(self.experimenttemperature2)
            self.temperaturewidgets.append(sublayout)
            self.generalwidgets.append(sublayout)
            self.grid.addLayout(sublayout, 6, 0)

            # Temperatura da sala
            label = QLabel(" ")
            label.setFixedWidth(self.labelsfixedsize)
            label.setFont(self.commonfont)
            sublayout = QHBoxLayout()
            sublayout.setSpacing(5)
            sublayout.addWidget(label)
            sublayout.addWidget(label)
            self.roomtemperature = QCheckBox("Temperatura da Sala")
            sublayout.addWidget(self.roomtemperature)
            self.temperaturewidgets.append(sublayout)
            self.generalwidgets.append(sublayout)
            self.grid.addLayout(sublayout, 7, 0)

        # Várias temperaturas
        if value == 1:
            label = QLabel(" ")
            label.setAlignment(Qt.AlignHCenter | Qt.AlignBottom)
            label.setFixedWidth(self.labelsfixedsize)
            label.setFont(self.commonfont)
            sublayout = QHBoxLayout()
            sublayout.setSpacing(5)
            sublayout.addWidget(label)
            label = QLabel("BVT [K]")
            label.setAlignment(Qt.AlignHCenter | Qt.AlignBottom)
            label.setFixedWidth(self.labelsfixedsize)
            label.setFont(self.commonfont)
            sublayout.addWidget(label)
            self.temperatura_amostra_label = QLabel("Amostra [ºC]")
            self.temperatura_amostra_label.setAlignment(Qt.AlignHCenter | Qt.AlignBottom)
            self.temperatura_amostra_label.setFixedWidth(self.labelsfixedsize)
            self.temperatura_amostra_label.setFont(self.commonfont)
            sublayout.addWidget(self.temperatura_amostra_label)
            self.temperaturewidgets.append(sublayout)
            self.generalwidgets.append(sublayout)
            self.grid.addLayout(sublayout, 5, 0)

            # Temperatura Inicial
            label = QLabel("Temperatura Inicial: ")
            label.setToolTip('Temperatura de realização do primeiro experimento')
            label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            label.setFixedWidth(self.labelsfixedsize)
            label.setFont(self.commonfont)
            sublayout = QHBoxLayout()
            sublayout.setSpacing(5)
            sublayout.addWidget(label)
            self.initialtemperature1 = QLineEdit()
            self.initialtemperature1.setAlignment(Qt.AlignLeft)
            self.initialtemperature1.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.initialtemperature1.setText("320")
            sublayout.addWidget(self.initialtemperature1)
            self.initialtemperature2 = QLineEdit()
            self.initialtemperature2.setAlignment(Qt.AlignLeft)
            self.initialtemperature2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.initialtemperature2.setText("46")
            sublayout.addWidget(self.initialtemperature2)
            self.temperaturewidgets.append(sublayout)
            self.generalwidgets.append(sublayout)
            self.grid.addLayout(sublayout, 6, 0)

            # Passo / Quantidade
            label = QLabel("Passo [K] / Quantidade: ")
            label.setToolTip('Espaçamento entre temperaturas')
            label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            label.setFixedWidth(self.labelsfixedsize)
            label.setFont(self.commonfont)
            sublayout = QHBoxLayout()
            sublayout.setSpacing(5)
            sublayout.addWidget(label)
            self.steptemperature1 = QLineEdit()
            self.steptemperature1.setAlignment(Qt.AlignLeft)
            self.steptemperature1.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.steptemperature1.setText("10")
            sublayout.addWidget(self.steptemperature1)
            self.steptemperature2 = QLineEdit()
            self.steptemperature2.setAlignment(Qt.AlignLeft)
            self.steptemperature2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.steptemperature2.setText("4")
            sublayout.addWidget(self.steptemperature2)
            self.temperaturewidgets.append(sublayout)
            self.generalwidgets.append(sublayout)
            self.grid.addLayout(sublayout, 7, 0)

            # Temperatura Final
            label = QLabel("Temperatura Final: ")
            label.setToolTip('Temperatura de realização do último experimento')
            label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            label.setFixedWidth(self.labelsfixedsize)
            label.setFont(self.commonfont)
            sublayout = QHBoxLayout()
            sublayout.setSpacing(5)
            sublayout.addWidget(label)
            self.endtemperature1 = QLineEdit()
            self.endtemperature1.setAlignment(Qt.AlignLeft)
            self.endtemperature1.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.endtemperature1.setText("360")
            sublayout.addWidget(self.endtemperature1)
            self.endtemperature2 = QLineEdit()
            self.endtemperature2.setAlignment(Qt.AlignLeft)
            self.endtemperature2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.endtemperature2.setText("98")
            sublayout.addWidget(self.endtemperature2)
            self.temperaturewidgets.append(sublayout)
            self.generalwidgets.append(sublayout)
            self.grid.addLayout(sublayout, 8, 0)

        # Temperatura por arquivo
        if value == 2:
            label = QLabel("Arquivo: ")
            label.setToolTip('Arquivo formato \'txt\' com temperaturas de execução do experimento')
            label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            label.setFixedWidth(self.labelsfixedsize)
            label.setFont(self.commonfont)
            sublayout = QHBoxLayout()
            sublayout.setSpacing(5)
            sublayout.addWidget(label)
            self.filetemperature = QLineEdit()
            self.filetemperature.setAlignment(Qt.AlignLeft)
            self.filetemperature.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            sublayout.addWidget(self.filetemperature)
            chossepathbutton = QPushButton("Escolher")
            chossepathbutton.setFixedWidth(70)
            chossepathbutton.clicked.connect(self.pickfile)
            sublayout.addWidget(chossepathbutton)
            self.temperaturewidgets.append(sublayout)
            self.generalwidgets.append(sublayout)
            self.grid.addLayout(sublayout, 6, 0)

    def temperaturetypechoose(self, value):
        '''Experimento em altas ou baixas temperaturas '''

        if value == 0:
            self.gasflow_evaporator_label.setText("Gas Flow: ")
        else:
            self.gasflow_evaporator_label.setText("Evaporator: ")

    ###

    def closeEvent(self, event):
        '''If we close a QWidget, the QCloseEvent is generated.
         To modify the widget behaviour we need to reimplement the closeEvent() event handler'''

        if self.on_experiment == False:
            if self.newstateEvent():
                event.accept()
            else:
                event.ignore()
            return

        reply = QMessageBox.question(self, 'Mensagem',"Experimento em execução! Finalizar e sair?", QMessageBox.Yes |
                                     QMessageBox.No, QMessageBox.Yes)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def newstateEvent(self):
        '''Mensagem warning'''

        if self.states[len(self.states)-1] == 0:
            return True

        reply = QMessageBox.question(self, 'Mensagem',"As modificações feitas serão perdidas. Continuar?", QMessageBox.Yes |
                                     QMessageBox.No, QMessageBox.Yes)

        return reply == QMessageBox.Yes

    def pickdir(self):
        '''Selecionar diretório'''

        dialog = QFileDialog()
        folder_path = dialog.getExistingDirectory(None, "Selecionar Pasta")
        self.amostrapath.setText(folder_path)

    def pickfile(self):
        '''Selecionar arquivo'''

        dialog = QFileDialog()
        file_path = dialog.getOpenFileName(None, "Selecionar Arquivo")
        self.filetemperature.setText(file_path[0])

    def center(self):
        '''Centrar interface na tela'''

        # QDesktopWidget class provides information about the user's desktop
        qr = self.frameGeometry() # get a rectangle specifying the geometry of the main window
        cp = QDesktopWidget().availableGeometry().center() #  figure out the screen resolution of our monitor
        qr.moveCenter(cp) # set the center of the rectangle to the center of the screen
        self.move(qr.topLeft()) # move the top-left point of the application window to the top-left point of the qr rectangle

    def deletewidgets(self, to_delete=0):
        '''Excluir todos os objetos criados para determinada função
            0 -> Todos os objetos
            1 -> Objetos de escolha de temperatura
        '''

        if to_delete == 0:
            for item in self.generalwidgets:
                if item.layout():
                    while item.count() > 0:
                        l_item = item.takeAt(0)
                        if not l_item:
                            continue
                        widget = l_item.widget()
                        if widget:
                            widget.deleteLater()
                else:
                    item.deleteLater()
            self.generalwidgets = []
            self.states.append(0)

        if to_delete == 1:
            for item in self.temperaturewidgets:
                if item.layout():
                    while item.count() > 0:
                        l_item = item.takeAt(0)
                        if not l_item:
                            continue
                        widget = l_item.widget()
                        if widget:
                            widget.deleteLater()
                else:
                    item.deleteLater()
                self.generalwidgets.remove(item)
            self.temperaturewidgets = []

class Experiment:

    def __init__(self):

        self.control = Control()
        self.COM = "COM4"
        self.serialcom = SerialCommunication(self.COM)
        self.executing = False
        self.to_break = False

    def __del__(self):
        del self.control

    def connect_serial(self):
        self.serialcom.Connect(self.COM)

    def run_all(self, data0, data1, data2, data3, data4, data5, data6, data7, data8, data9):
        #gas_flow, tune, temperatures_BVT, wait_time, applications, ramps, serialnumber, exe_path, low_temperature, 0

        pythoncom.CoInitialize()

        write_log("Experimento iniciado")

        self.executing = True

        self.control.set_parameters(data6, data7)

        if self.start(data0, data1, data8):
            first = True
            for temperature in data2:
                if not self.run(temperature, data3, data4, data5, data8, data9, init=first):
                    break
                #mainwindow.mpb["value"] += 1
                first = False
            else:
                self.end(data8, data9)

        self.executing = False

        write_log("Experimento finalizado")

    def start(self, gasflow, to_tune, low_temperature):

        write_log("Inicialização de equipamentos iniciada")

        self.to_break = False
        sleep(1)
        if not self.control.ConnectBVT():
            self.to_break = True
            self.control.Finish(ramp=False, bvt=False, minispec=False)
            write_log("Inicialização de equipamentos finalizada com falha em ConnectBVT")
            return False
        if not self.control.ConnectPNMR():
            self.to_break = True
            self.control.Finish(ramp=False, bvt=False)
            write_log("Inicialização de equipamentos finalizada com falha em ConnectPNMR")
            return False
        if not self.control.StartBVT(gasflow, low_temperature, tune=to_tune):
            self.to_break = True
            self.control.Finish(ramp=False)
            write_log("Inicialização de equipamentos finalizada com falha em StartBVT")
            return False
        write_log("Inicialização de equipamentos finalizada com sucesso")
        return True

    def run(self, temperature, wait_time, applications, do_ramp, low_temperature, calib, init=False):
        self.to_break = False
        sleep(1)

        write_log("Estabilização de nova temperatura: " + str(temperature) + " K(BVT)")

        if temperature != 0:
            current_temperature = self.control.GetTemperature()
            if current_temperature == -1:
                self.to_break = True
                self.control.AbortApplication()
                self.control.Finish(low_temperature=low_temperature)
                write_log("Término de execução em: " + str(temperature) + " K com falha em GetTemperature")
                return False

            if (abs(current_temperature - temperature) >= do_ramp[1] and do_ramp[0]) or init:
                if not self.control.DoRamp(temperature, 15, to_sleep=wait_time):
                    self.to_break = True
                    self.control.AbortApplication()
                    self.control.Finish(low_temperature=low_temperature)
                    write_log("Término de execução em: " + str(temperature) + " K com falha em GetTemperature")
                    return False
            else:
                if not self.control.SetTemperature(temperature, wait_time):
                    self.to_break = True
                    self.control.AbortApplication()
                    self.control.Finish(low_temperature=low_temperature)
                    write_log("Término de execução em: " + str(temperature) + " K com falha em SetTemperature")
                    return False
        if calib == 0:
            for app in applications:
                write_log("Início da aplicação: " + str(app))
                if not self.control.ExecuteApplication(app):
                    self.to_break = True
                    self.control.AbortApplication()
                    self.control.Finish(low_temperature=low_temperature)
                    write_log("Término de execução em: " + str(temperature) + " K com falha em ExecuteApplication")
                    return False
                write_log("Término da aplicação: " + str(app))
        '''else:
            bvt_temperature = self.control.GetTemperature()
            amostra_temperature = self.serialcom.ReadTemperature()
            amostra = main_window.calib_entry.get()
            to_save_path = main_window.path_entry.get()
            info_file = open(to_save_path + "\\" + amostra + " calibracao curva.txt", "a")
            info_file.write(str(bvt_temperature) + "," + str(amostra_temperature) + "\n")
            info_file.close()'''

        write_log("Término de execução em: " + str(temperature) + " K com sucesso")
        return True

    def end(self, low_temperature, calib):

        '''if calib != 0:
            amostra = main_window.calib_entry.get()
            to_save_path = main_window.path_entry.get()
            info_file = open(to_save_path + "\\" + amostra + " calibracao curva.txt", "r")

            # a = ym - b*xm
            # b = sum((xí -xm)(yi - ym))/sum(xi-xm)**2
            x = []
            y = []
            for line in info_file:
                sup = ""
                for c in line:
                    if c == ",":
                        x.append(float(sup))
                        sup = ""
                    elif c == "\n":
                        y.append(float(sup))
                        sup = ""
                    else:
                        sup += c

            xm = sum(x)/len(x)
            ym = sum(y)/len(y)
            b = sum([(v1-xm)*(v2-ym) for v1, v2 in zip(x, y)])/float(sum([(v1-xm)**2 for v1 in x]))
            a = ym - b*xm

            write_file = open(to_save_path + "\\" + amostra + " calibracao infos.txt", "a")
            write_file.write("a: " + str(a) + "\n")
            write_file.write("b: " + str(b) + "\n")
            write_file.close()

            #result = tkMessageBox.askquestion("Atualizar", "Atualizar parâmetros? (a=" + str(round(a, 2)) +
            #                                  ", b=" + str(round(b, 2)) + ")")

            result = "yes"

            if result == 'yes':
                filename = "Temp_params.txt" if not low_temperature else "Temp_params_low.txt"
                file = open(filename, "w")
                file.write(str(a))
                file.write("\n")
                file.write(str(b))
                file.close()'''


        self.to_break = False
        write_log("Finalização de equipamentos iniciada")
        sleep(1)
        if not self.control.Finish(low_temperature=low_temperature):
            self.to_break = True
            write_log("Finalização de equipamentos finalizada com falha em Finish")
            return False
        write_log("Finalização de equipamentos finalizada com sucesso")
        return True

def manage_window():
    global mainwindow
    pythoncom.CoInitialize()
    application = QApplication(sys.argv)
    mainwindow= GUI()
    sys.exit(application.exec_())

def manage_experiment():
    global experiment
    pythoncom.CoInitialize()
    experiment = Experiment()
    while gui_running:
        if not gui_data.empty():
            data = gui_data.get()

            #---

            execution = Thread(target=experiment.run_all, args=(list(data)))
            execution.start()
            stop = "1"
            if data[9] == 1:
                stop = "2"
            while not experiment.executing:
                pass
            while experiment.executing:
                if not gui_data.empty():
                    stop = gui_data.get()
                    if stop != "STOP1" and stop != "STOP2":
                        print("STOP ERROR")
                    else:
                        experiment.control.stop = True
                        break
            execution.join()

            #---

            experiment_data.put(stop)
    del experiment

def main():

    global gui_data, experiment_data, gui_running

    # Create Queues
    gui_data = Queue.Queue()
    experiment_data = Queue.Queue()

    # Create Threads
    gui_running = True
    gui = Thread(target=manage_window, args=())
    experiment = Thread(target=manage_experiment, args=())

    # Start Threads
    gui.start()
    experiment.start()

    # Wait Threads to finish
    gui.join()
    gui_running = False
    experiment.join()

if __name__ == "__main__":
    main()