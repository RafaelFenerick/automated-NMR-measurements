# -*- coding: utf-8 -*-

from Tkinter import Tk, Frame, Label, Entry, Menu, Button, Checkbutton, Radiobutton, Scrollbar, Listbox, Text, IntVar
from Tkinter import TOP, RIGHT, LEFT, CENTER, BOTTOM, DISABLED, NORMAL, SUNKEN, FLAT, RAISED, SOLID, GROOVE, X, Y, BOTH, W, N, S, E, END, INSERT
import tkMessageBox, tkFileDialog
import ttk
#from PIL import ImageTk, Image
from shutil import copyfile
from threading import Thread
import Queue
from time import sleep
from datetime import datetime as timecurrent
import pythoncom
from VBS_Control import Control
from SerialCommunication import SerialCommunication
import os
import sys

def float_to_time(number):
    number = int(number)
    hour = number / 60
    minute = number % 60

    str_minute = str(minute)
    if len(str_minute) == 1:
        str_minute = "0" + str_minute

    return str(hour) + ":" + str_minute + "h"

def sum_time(current, duration):
    hour = current.hour
    minute = current.minute

    hour_duration = int(duration)/60
    minute_duration = int(duration%60)

    end_hour = int(hour) + hour_duration
    end_minute = int(minute) + minute_duration

    while end_minute >= 60:
        end_minute -= 60
        end_hour += 1

    day = 0
    while end_hour >= 24:
        end_hour -= 24
        day += 1

    return str(day) + "d - " + str(end_hour) + ":" + str(end_minute) + "h"

def write_log(message):
    current_time = timecurrent.now()
    day_to_write = str(current_time.day) + "/" + str(current_time.month) + "/" + str(current_time.year)
    time_to_write = str(current_time.hour) + ":" + str(current_time.minute) + ":" + str(current_time.second) + \
    "." + str(current_time.microsecond)

    if(main_window.log_text) != None:
        main_window.log_text.config(state=NORMAL)
        main_window.log_text.insert(INSERT, day_to_write + " " + time_to_write + " - " + message + "\n")
        main_window.log_text.config(state=DISABLED)

    info_file = open("log.txt", "a")
    info_file.write(day_to_write + " " + time_to_write + " - " + message + "\n")
    info_file.close()

class GUI:

    def __init__(self):
        '''Inicializa o GUI, adicionando menus e frames.'''

        #Handlers
        self.main_objects = []
        self.temperature_objects = []
        self.cancel_objects = []
        self.time_objects = []
        self.screens = [0]
        self.experiment_infos = {}
        self.on_experiment = False
        self.file_temperatures_array = []
        self.estimated_time_value = 0
        self.log_text = None

        #Create Window
        self.master = Tk()
        self.master.title("PNMR Control")
        #self.master.iconbitmap(default='Bases\\logo.ico')
        self.master.geometry("850x670")
        self.master.resizable(width=False, height=False)

        #Menu
        self.menu = Menu(self.master)
        self.master.config(menu=self.menu)
        self.submenu = Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Arquivo", menu=self.submenu)
        self.editmenu = Menu(self.master, tearoff=0)
        self.menu.add_cascade(label="Configurações", menu=self.editmenu)

        self.submenu.add_command(label="Novo", command=self.config_experiment)
        self.submenu.add_command(label="Calibração", command=self.config_calib)
        self.submenu.add_command(label="Adicionar Aplicação", command=self.config_add_application)
        self.submenu.add_command(label="Remover Aplicação", command=self.config_remove_application)
        self.submenu.add_separator()
        self.submenu.add_command(label="Sair", command=self.master.destroy)

        self.editmenu.add_command(label="Parâmetros de Temperatura", command=self.config_temperature_calculus)
        self.editmenu.add_command(label="Parâmetros de Baixa Temperatura", command=self.config_low_temperature_calculus)
        self.editmenu.add_command(label="Parâmetros do Equipamento", command=self.config_equipment_info)
        self.editmenu.add_command(label="Parâmetros de Leitura", command=self.config_filetemp_info)

        ##ToolBar
        #self.toolbar = Frame(self.master, bd=1, relief=SUNKEN)
        #self.insertbutt = Button(self.toolbar, text="Isert Image", command=self.not_avaliable, state=NORMAL)
        ##self.photo = ImageTk.PhotoImage(Image.open("soma.png"))
        ##self.insertbutt.config(image=self.photo, width="10", height="10", state=DISABLED)
        #self.insertbutt.pack(side=LEFT, padx=2, pady=2)
        #self.insertbutt2 = Button(self.toolbar, text="Isert Image2", command=self.not_avaliable, state=NORMAL)
        ##self.photo2 = ImageTk.PhotoImage(Image.open("sub.png"))
        ##self.insertbutt2.config(image=self.photo2, width="10", height="10", state=DISABLED)
        #self.insertbutt2.pack(side=LEFT, padx=2, pady=2)
        #self.toolbar.pack(side=TOP, fill=X)


        #Status Bar
        self.status = Label(self.master, text="Esperando", bd=1, relief=SUNKEN, anchor=W)
        self.status.pack(side=BOTTOM, fill=X)

        #Frames
        self.right_frame = Frame(self.master, width=350)
        #self.right_frame.config(bg="blue")
        self.right_frame.pack_propagate(0)
        self.right_frame.pack(side=RIGHT, fill=BOTH, expand=False)

        self.frame_last = Frame(self.master, heigh=40)
        #self.frame_last.config(bg="black")
        self.frame_last.pack_propagate(0)
        self.frame_last.pack(side=BOTTOM, fill=BOTH)

        self.frame_ajustable = Frame(self.master, borderwidth=2)
        #self.frame_ajustable.config(bg="red")
        self.frame_ajustable.pack(side=BOTTOM, fill=BOTH, expand=True)

        self.frame_14 = Frame(self.master, heigh=40)
        # self.frame_14.config(bg="yellow")
        self.frame_14.pack_propagate(0)
        self.frame_14.pack(side=BOTTOM, fill=X)

        self.frame_13 = Frame(self.master, heigh=40)
        # self.frame_13.config(bg="yellow")
        self.frame_13.pack_propagate(0)
        self.frame_13.pack(side=BOTTOM, fill=X)

        self.frame_12 = Frame(self.master, heigh=40)
        # self.frame_12.config(bg="yellow")
        self.frame_12.pack_propagate(0)
        self.frame_12.pack(side=BOTTOM, fill=X)

        self.frame_11 = Frame(self.master, heigh=40)
        # self.frame_11.config(bg="yellow")
        self.frame_11.pack_propagate(0)
        self.frame_11.pack(side=BOTTOM, fill=X)

        self.frame_10 = Frame(self.master, heigh=40)
        #self.frame_10.config(bg="yellow")
        self.frame_10.pack_propagate(0)
        self.frame_10.pack(side=BOTTOM, fill=X)

        self.frame_9 = Frame(self.master, heigh=40)
        #self.frame_9.config(bg="yellow")
        self.frame_9.pack_propagate(0)
        self.frame_9.pack(side=BOTTOM, fill=X)

        self.frame_8 = Frame(self.master, heigh=40)
        #self.frame_8.config(bg="orange")
        self.frame_8.pack_propagate(0)
        self.frame_8.pack(side=BOTTOM, fill=X)

        self.frame_7 = Frame(self.master, heigh=40)
        #self.frame_7.config(bg="purple")
        self.frame_7.pack_propagate(0)
        self.frame_7.pack(side=BOTTOM, fill=X)

        self.frame_6 = Frame(self.master, heigh=40)
        #self.frame_6.config(bg="green")
        self.frame_6.pack_propagate(0)
        self.frame_6.pack(side=BOTTOM, fill=X)

        self.frame_5 = Frame(self.master, heigh=40)
        #self.frame_5.config(bg="white")
        self.frame_5.pack_propagate(0)
        self.frame_5.pack(side=BOTTOM, fill=BOTH)

        self.frame_4 = Frame(self.master, heigh=40)
        self.frame_4.pack_propagate(0)
        #self.frame_4.config(bg="pink")
        self.frame_4.pack(side=BOTTOM, fill=BOTH)

        self.frame_3 = Frame(self.master, heigh=40)
        #self.frame_3.config(bg="green")
        self.frame_3.pack_propagate(0)
        self.frame_3.pack(side=BOTTOM, fill=BOTH)

        self.frame_2 = Frame(self.master, heigh=40)
        #self.frame_2.config(bg="pink")
        self.frame_2.pack_propagate(0)
        self.frame_2.pack(side=BOTTOM, fill=BOTH)

        self.frame_1 = Frame(self.master, heigh=40)
        #self.frame_1.config(bg="green")
        self.frame_1.pack_propagate(0)
        self.frame_1.pack(side=BOTTOM, fill=BOTH)

        self.frame_0 = Frame(self.master, heigh=40)
        #self.frame_0.config(bg="pink")
        self.frame_0.pack_propagate(0)
        self.frame_0.pack(side=BOTTOM, fill=BOTH)

        #Sub_Frames
        self.sub_right_frame_down_down = Frame(self.right_frame, heigh=30)
        #self.sub_right_frame_down_down.config(bg="grey")
        self.sub_right_frame_down_down.pack_propagate(0)
        self.sub_right_frame_down_down.pack(side=BOTTOM, fill=X)
        self.sub_right_frame_down = Frame(self.right_frame, heigh=270)
        #self.sub_right_frame_down.config(bg="grey")
        self.sub_right_frame_down.pack_propagate(0)
        self.sub_right_frame_down.pack(side=BOTTOM, fill=X)
        self.sub_right_frame_up = Frame(self.right_frame, heigh=60)
        #self.sub_right_frame_up.config(bg="black")
        self.sub_right_frame_up.pack_propagate(0)
        self.sub_right_frame_up.pack(side=TOP, fill=X)
        self.sub_right_frame_center = Frame(self.right_frame)
        #self.sub_right_frame_center.config(bg="brown")
        self.sub_right_frame_center.pack_propagate(0)
        self.sub_right_frame_center.pack(side=BOTTOM, fill=BOTH, expand=True)

        self.check_queue()

    def check_queue(self):
        if not experiment_data.empty():
            data = experiment_data.get()


            self.on_experiment = False
            if data == "1" or data == "STOP1":
                self.status.config(text="Definindo experimento.", fg="green")
            else:
                self.status.config(text="Definindo calibração.", fg="green")
            self.stop_experiment_button.config(text="Cancelar", command=self.clean)
            self.stop_experiment_button.config(state=NORMAL)
            self.clean(array=4)
            self.log_text = None

            # Tempo Estimado
            self.estimated_time = Label(self.frame_10, text="Tempo Estimado: ", width=50, pady=8,
                                        anchor=CENTER)
            self.estimated_time.pack(side=BOTTOM)
            self.time_objects.append(self.estimated_time)

            if data == "1" or data == "STOP1":
                self.enable(True)
            else:
                self.enable(False)

        self.master.after(100, self.check_queue)

    #--- Experiment

    def config_experiment(self):
        '''Disponibiliza todos os campos para preenchimento
        a fim de executar um novo experimento.'''

        if self.on_experiment:
            tkMessageBox.showinfo("Aviso", "Experimento em Execução.")
            return

        write_log("Configuração de experimento iniciada")

        self.clean()
        self.screens.append(1)
        self.status.config(text="Definindo experimento.", fg="green")

        # Title
        self.defineexperiment_label = Label(self.frame_0, text="Definições do Experimento", width=50, pady=8, anchor=CENTER)
        self.defineexperiment_label.pack(side=TOP)
        self.main_objects.append(self.defineexperiment_label)

        # Amostra
        self.amostra_label = Label(self.frame_2, text="Amostra: ", width=30, pady=8, anchor=W)
        self.amostra_entry = Entry(self.frame_2, width=35, bd=3, state=NORMAL)
        self.amostra_entry.insert(0, "teste")
        self.amostra_label.pack(side=LEFT)
        self.amostra_entry.pack(side=LEFT)
        self.main_objects.append(self.amostra_label)
        self.main_objects.append(self.amostra_entry)

        # Path
        self.path_label = Label(self.frame_3, text="Destino: ", width=30, pady=8, anchor=W)
        self.path_entry = Entry(self.frame_3, width=35, bd=3, state=NORMAL)
        self.path_entry.insert(0, "C:\\Users\\BRUKER\\Desktop\\USUARIOS\\Eduardo_IFSC")
        #self.path_entry.insert(0, "C:/Users/Rafael Fenerick/Documents/Git/pnmr")
        self.pathbutton = Button(self.frame_3, text="Escolher", bd=3, command=self.choose_path, justify=CENTER, width=6)
        self.path_label.pack(side=LEFT)
        self.path_entry.pack(side=LEFT)
        self.pathbutton.pack(side=RIGHT)
        self.main_objects.append(self.path_label)
        self.main_objects.append(self.path_entry)
        self.main_objects.append(self.pathbutton)

        # Temperature
        self.temperature_control_var = IntVar()
        self.one_temperature_radiobutton = Radiobutton(self.frame_4, text="Uma temperatura", variable=self.temperature_control_var, command=self.config_temperature_type, value=0, padx=40)
        self.more_temperature_radiobutton = Radiobutton(self.frame_4, text="Várias temperaturas", variable=self.temperature_control_var, command=self.config_temperature_type, value=1, padx=40)
        self.more_temperature_radiobutton.pack(side=LEFT, anchor=W)
        self.one_temperature_radiobutton.pack(side=LEFT, anchor=W)
        self.main_objects.append(self.one_temperature_radiobutton)
        self.main_objects.append(self.more_temperature_radiobutton)
        self.one_temperature_radiobutton.select()

        #Construct temperature infos
        self.config_temperature_type()

        # Gas Flow
        self.frame_11.config(relief=GROOVE)
        self.temperature_level_var = IntVar()
        self.temperature_level_var.set(1)
        self.gasflow_var_checkbutton = Radiobutton(self.frame_11, text="Altas Temperaturas - Gas Flow:",
                                                       variable=self.temperature_level_var,
                                                       value=1, height=3, width=30)
        self.gasflow_var_checkbutton.pack(side=LEFT)
        self.main_objects.append(self.gasflow_var_checkbutton)

        self.gasflow_entry = Entry(self.frame_11, width=35, bd=3, state=NORMAL)
        self.gasflow_entry.insert(0, "2000")
        self.gasflow_var_checkbutton.pack(side=LEFT)
        self.gasflow_entry.pack(side=LEFT)
        self.main_objects.append(self.gasflow_var_checkbutton)
        self.main_objects.append(self.gasflow_entry)

        #Low Temperature
        self.low_temperature_checkbutton = Radiobutton(self.frame_12, text="Baixas Temperaturas - Potência:",
                                                       variable=self.temperature_level_var,
                                                       value=2, height=3, width=30)
        self.low_temperature_checkbutton.pack(side=LEFT)
        self.main_objects.append(self.low_temperature_checkbutton)

        self.low_temperature_entry = Entry(self.frame_12, width=35, bd=3, state=NORMAL)
        self.low_temperature_entry.insert(0, "10")
        self.low_temperature_entry.pack(side=LEFT)
        self.main_objects.append(self.low_temperature_entry)

        # Ramp
        self.rampmiddle_control_var = IntVar()
        self.rampmiddle_checkbutton = Checkbutton(self.frame_13, text="Rampa em Variações maiores que",
                                                  variable=self.rampmiddle_control_var,
                                                  onvalue=1, offvalue=0, height=3, width=30)
        self.rampmiddle_checkbutton.pack(side=LEFT)
        self.main_objects.append(self.rampmiddle_checkbutton)

        self.rampmiddle_entry = Entry(self.frame_13, width=35, bd=3, state=NORMAL)
        self.rampmiddle_entry.insert(0, "10")
        self.rampmiddle_entry.pack(side=LEFT)
        self.main_objects.append(self.rampmiddle_entry)

        # Tune
        self.tune_control_var = IntVar()
        self.tune_checkbutton = Checkbutton(self.frame_14, text="Tune", variable=self.tune_control_var,
                                            onvalue=1, offvalue=0, height=3, width=10)
        self.tune_checkbutton.pack(side=LEFT)
        self.main_objects.append(self.tune_checkbutton)

        # Application
        self.application_label = Label(self.sub_right_frame_up, text="Escolha a Aplicação", width=50, pady=8, anchor=CENTER)
        self.application_label.pack(side=TOP)
        self.main_objects.append(self.application_label)

        # Get stored application
        try:
            file = open("Applications\\Applications.txt", "r")
            self.applications_text = file.readlines()
            file.close()
        except:
            self.applications_text = []
        self.application_control_var = []
        self.applications = []

        if len(self.applications_text) == 0:
            tkMessageBox.showinfo("Aviso", "Nenhuma aplicação adicionada!")
            self.clean()
            return

        i = 1
        for application in self.applications_text:
            self.application_control_var.append(IntVar())
            application_checkbutton = Checkbutton(self.sub_right_frame_center, text=application.strip("\n"),
                                                  variable=self.application_control_var[len(self.application_control_var) - 1],
                                                  onvalue=i, offvalue=0, padx=40)
            application_checkbutton.pack(anchor=W)
            self.main_objects.append(application_checkbutton)
            self.applications.append(application_checkbutton)
            i += 1

        #Main Buttons
        self.start_experiment_button = Button(self.frame_last, text="Iniciar", bd=3, command=self.check_start_experiment, justify=CENTER, width=10)
        self.start_experiment_button.pack(side=RIGHT)
        self.main_objects.append(self.start_experiment_button)

        self.stop_experiment_button = Button(self.frame_last, text="Cancelar", bd=3, command=self.clean, justify=CENTER, width=10)
        self.stop_experiment_button.pack(side=RIGHT)
        self.cancel_objects.append(self.stop_experiment_button)

        write_log("Configuração de experimento finalizada")

    def config_temperature_type(self):

        self.clean(array=2)
        self.clean(array=4)

        # Tempo Estimado
        self.estimated_time = Label(self.frame_10, text="Tempo Estimado: ", width=50, pady=8,
                                    anchor=CENTER)
        self.estimated_time.pack(side=BOTTOM)
        self.time_objects.append(self.estimated_time)

        # Multiple temperatures
        if self.temperature_control_var.get() == 1:

            self.synchro_mult_command = (self.master.register(self.synchro_multiple), '%P', '%W')

            # Temperatura Inicial
            self.init_temperature_label = Label(self.frame_5, text="Temperatura Inicial: [K(BVT) / K / °C]", width=30, pady=8, anchor=W)
            self.init_temperature_entry = Entry(self.frame_5, width=11, bd=3, state=NORMAL)
            self.init_temperature_entry2 = Entry(self.frame_5, width=12, bd=3, state=NORMAL)
            self.init_temperature_entry3 = Entry(self.frame_5, width=11, bd=3, state=NORMAL)

            self.init_temperature_entry.insert(0, "310")
            self.init_temperature_entry2.insert(0, "306")
            self.init_temperature_entry3.insert(0, "32.85")

            self.init_temperature_entry.config(validate='key', validatecommand=self.synchro_mult_command)
            self.init_temperature_entry2.config(validate='key', validatecommand=self.synchro_mult_command)
            self.init_temperature_entry3.config(validate='key', validatecommand=self.synchro_mult_command)
            self.init_temperature_label.pack(side=LEFT)
            self.init_temperature_entry.pack(side=LEFT)
            self.init_temperature_entry2.pack(side=LEFT)
            self.init_temperature_entry3.pack(side=LEFT)
            self.temperature_objects.append(self.init_temperature_label)
            self.temperature_objects.append(self.init_temperature_entry)
            self.temperature_objects.append(self.init_temperature_entry2)
            self.temperature_objects.append(self.init_temperature_entry3)

            # Step Temperature
            self.step_temperature_label = Label(self.frame_6, text="Passo [K]/ Quantidade: ", width=30, pady=8, anchor=W)
            self.step_temperature_entry = Entry(self.frame_6, width=17, bd=3, state=NORMAL)

            self.step_temperature_entry.insert(0, "10")

            self.step_temperature_entry.config(validate='key', validatecommand=self.synchro_mult_command)
            self.step_temperature_entry2 = Entry(self.frame_6, width=18, bd=3, validate='key',
                                                 validatecommand=self.synchro_mult_command, state=NORMAL)
            self.step_temperature_label.pack(side=LEFT)
            self.step_temperature_entry.pack(side=LEFT)
            self.step_temperature_entry2.pack(side=LEFT)
            self.temperature_objects.append(self.step_temperature_label)
            self.temperature_objects.append(self.step_temperature_entry)
            self.temperature_objects.append(self.step_temperature_entry2)

            # Temperatura Final
            self.end_temperature_label = Label(self.frame_7, text="Temperatura Final: [K(BVT) / K / °C]", width=30, pady=8, anchor=W)
            self.end_temperature_entry = Entry(self.frame_7, width=11, bd=3, state=NORMAL)
            self.end_temperature_entry2 = Entry(self.frame_7, width=12, bd=3, state=NORMAL)
            self.end_temperature_entry3 = Entry(self.frame_7, width=11, bd=3, state=NORMAL)

            self.end_temperature_entry.insert(0, "340")
            self.end_temperature_entry2.insert(0, "345")
            self.end_temperature_entry3.insert(0, "71.85")

            self.end_temperature_entry.config(validate='key', validatecommand=self.synchro_mult_command)
            self.end_temperature_entry2.config(validate='key', validatecommand=self.synchro_mult_command)
            self.end_temperature_entry3.config(validate='key', validatecommand=self.synchro_mult_command)
            self.end_temperature_label.pack(side=LEFT)
            self.end_temperature_entry.pack(side=LEFT)
            self.end_temperature_entry2.pack(side=LEFT)
            self.end_temperature_entry3.pack(side=LEFT)
            self.temperature_objects.append(self.end_temperature_label)
            self.temperature_objects.append(self.end_temperature_entry)
            self.temperature_objects.append(self.end_temperature_entry2)
            self.temperature_objects.append(self.end_temperature_entry3)

            # Waiting Time
            self.wait_time_mult_label = Label(self.frame_9, text="Tempo de Espera: [min]", width=30, pady=8, anchor=W)
            self.wait_time_mult_entry = Entry(self.frame_9, width=36, bd=3, state=NORMAL)

            self.wait_time_mult_entry.insert(0, "10")

            self.wait_time_mult_entry.config(validate='key', validatecommand=self.synchro_mult_command)
            self.wait_time_mult_label.pack(side=LEFT)
            self.wait_time_mult_entry.pack(side=LEFT)
            self.temperature_objects.append(self.wait_time_mult_label)
            self.temperature_objects.append(self.wait_time_mult_entry)

            # Temperatures from file
            self.file_temperatures_entry = Entry(self.frame_8, width=36, bd=3, state=NORMAL)
            self.file_temperatures_entry.insert(0, "")
            self.file_temperatures_control_var = IntVar()
            self.file_temperatures_checkbox = Checkbutton(self.frame_8, text="Por arquivo", variable=self.file_temperatures_control_var,
                                                         onvalue=1, offvalue=0, height=5, width=27, command=self.toggle_temperatures_file)
            self.file_temperatures_button = Button(self.frame_8, text="Escolher", bd=3, command=self.choose_temperatures_file, justify=CENTER, width=6)

            self.file_temperatures_checkbox.pack(side=LEFT)
            self.file_temperatures_entry.pack(side=LEFT)
            self.file_temperatures_button.pack(side=RIGHT)
            self.temperature_objects.append(self.file_temperatures_entry)
            self.temperature_objects.append(self.file_temperatures_checkbox)
            self.temperature_objects.append(self.file_temperatures_button)

            self.step_temperature_entry2.insert(0, "5")

            self.toggle_temperatures_file()

        # One temperature
        if self.temperature_control_var.get() == 0:
            self.synchro_single_command = (self.master.register(self.synchro_single), '%P', '%W')

            self.temperature_label = Label(self.frame_5, text="Temperatura: [K(BVT) / K / °C]", width=30, pady=8, anchor=W)
            self.temperature_entry = Entry(self.frame_5, width=11, bd=3, state=NORMAL)
            self.temperature_entry2 = Entry(self.frame_5, width=12, bd=3, state=NORMAL)
            self.temperature_entry3 = Entry(self.frame_5, width=11, bd=3, state=NORMAL)
            self.temperature_entry.insert(0, "310")
            self.temperature_entry.config(validate='key', validatecommand=self.synchro_single_command)
            self.temperature_entry2.config(validate='key', validatecommand=self.synchro_single_command)
            self.temperature_entry3.config(validate='key', validatecommand=self.synchro_single_command)
            self.temperature_label.pack(side=LEFT)
            self.temperature_entry.pack(side=LEFT)
            self.temperature_entry2.pack(side=LEFT)
            self.temperature_entry3.pack(side=LEFT)
            self.temperature_objects.append(self.temperature_label)
            self.temperature_objects.append(self.temperature_entry)
            self.temperature_objects.append(self.temperature_entry2)
            self.temperature_objects.append(self.temperature_entry3)

            self.room_temperature_control_var = IntVar()
            self.room_temperature_checkbox = Checkbutton(self.frame_6, text="Temperatura da sala", variable=self.room_temperature_control_var,
                                                         onvalue=1, offvalue=0, height=5, width=30, command=self.toggle_temperature_type)
            self.room_temperature_checkbox.pack(side=RIGHT)
            self.temperature_objects.append(self.room_temperature_checkbox)

            self.wait_time_single_label = Label(self.frame_7, text="Tempo de Espera: [min]", width=30, pady=8, anchor=W)
            self.wait_time_single_entry = Entry(self.frame_7, width=35, bd=3, state=NORMAL)
            self.wait_time_single_entry.config(validate='key', validatecommand=self.synchro_single_command)
            self.wait_time_single_label.pack(side=LEFT)
            self.wait_time_single_entry.pack(side=LEFT)
            self.temperature_objects.append(self.wait_time_single_label)
            self.temperature_objects.append(self.wait_time_single_entry)

            self.wait_time_single_entry.insert(0, "1")

    def check_start_experiment(self):
        '''Executa checks para confirmar que todos os parâmetros
        estão preenchidos corretamente e continua com a
        execução do experimento.'''

        write_log("Checagem de início de experimento iniciada")

        # Define applications to run
        self.to_run_applications = []
        for value in self.application_control_var:
            if value.get() != 0:
                self.to_run_applications.append(self.applications_text[value.get() - 1].strip("\n"))

        if len(self.to_run_applications) == 0:
            tkMessageBox.showinfo("Aviso", "Nenhuma aplicação escolhida!")
            return

        if self.amostra_entry.get() == "":
            tkMessageBox.showinfo("Aviso", "Nome da amostra não definido!")
            return

        if self.path_entry.get() == "" or self.path_entry.get() == "C:\\":
            tkMessageBox.showinfo("Aviso", "Caminho não definido!")
            return

        if not os.path.isdir(self.path_entry.get()):
            tkMessageBox.showinfo("Aviso", "Caminho especificado não existe!")
            return

        if self.temperature_control_var.get() == 1:

            if self.file_temperatures_control_var.get() == 0:

                if self.init_temperature_entry.get() == "":
                    tkMessageBox.showinfo("Aviso", "Temeperatura inicial não definida!")
                    return

                if self.end_temperature_entry.get() == "":
                    tkMessageBox.showinfo("Aviso", "Temeperatura final não definida!")
                    return

                if self.step_temperature_entry.get() == "":
                    tkMessageBox.showinfo("Aviso", "Passo não definido!")
                    return

                if self.step_temperature_entry2.get() == "":
                    tkMessageBox.showinfo("Aviso", "Quantidade de experimentos não definida!")
                    return

            if self.file_temperatures_control_var.get() == 1:
                if self.file_temperatures_entry.get() == "":
                    tkMessageBox.showinfo("Aviso", "Nenhum arquivo selecionado!")
                    return

                if self.file_temperatures_entry.get()[len(self.file_temperatures_entry.get())-3:] != "txt":
                    tkMessageBox.showinfo("Aviso", "Formato de arquivo incorreto!")
                    return

            if self.wait_time_mult_entry.get() == "":
                tkMessageBox.showinfo("Aviso", "Tempo de espera não definido!")
                return

        if self.gasflow_entry.get() == "":
            tkMessageBox.showinfo("Aviso", "Gas Flow não definido!")
            return

        write_log("Checagem de início de experimento finalizada")

        self.start_experiment()

    def start_experiment(self):
        '''Executa o experimento especificado,
        dada a amostra e o range de temperatura
        escolhidos.'''

        write_log("Inicialização de experimento iniciada")

        self.disable(1)
        self.disable(2)
        self.clean(4)

        self.screens.append(2)
        self.status.config(text="Executando experimento.", fg="red")

        temperatures_BVT = []
        temperatures_amostra = []
        temperatures_amostra_celsius = []

        if self.temperature_level_var.get() == 1:
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

        ramps = [self.rampmiddle_control_var.get(), float(self.rampmiddle_entry.get())]

        if self.temperature_control_var.get() == 1:
            wait_time = float(self.wait_time_mult_entry.get())
            if self.file_temperatures_control_var.get() == 1:
                for temperature in self.file_temperatures_array:
                    if filetemp == 1:
                        temperatures_BVT.append(temperature)
                        temperatures_amostra.append(round(float((temperature + b)/a), 2))
                        temperatures_amostra_celsius.append(round(float((temperature + b)/a) - 273.15, 2))
                    elif filetemp == 2:
                        temperatures_BVT.append(round(float(temperature*a - b), 2))
                        temperatures_amostra.append(temperature)
                        temperatures_amostra_celsius.append(round(temperature - 273.15, 2))
                    else:
                        temperatures_BVT.append(round(float((temperature + 273.15)*a - b), 2))
                        temperatures_amostra.append(round(temperature + 273.15, 2))
                        temperatures_amostra_celsius.append(temperature)

            else:
                init = float(self.init_temperature_entry.get())
                end = float(self.end_temperature_entry.get())
                step = float(self.step_temperature_entry.get())
                amount = float(self.step_temperature_entry2.get())
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
        else:
            wait_time = float(self.wait_time_single_entry.get())
            if self.room_temperature_control_var.get() == 1:
                temperatures_BVT.append(0)
                temperatures_amostra.append(0)
                temperatures_amostra_celsius.append(0)
            else:
                temperatures_BVT.append(float(self.temperature_entry.get()))
                temperatures_amostra.append(round(float((float(self.temperature_entry.get()) + b) / a), 2))
                temperatures_amostra_celsius.append(round(float((float(self.temperature_entry.get()) + b) / a) - 273.15, 2))

        amostra = self.amostra_entry.get()
        to_save_path = self.path_entry.get()
        if to_save_path[len(to_save_path)-1] != "\\":
            to_save_path += "\\"
        tune = int(self.tune_control_var.get())
        gas_flow = int(self.gasflow_entry.get())
        low_temperature = [int(self.temperature_level_var.get())-1, float(self.low_temperature_entry.get())]
        low_temperature[0] = True if low_temperature[0] == 1 else False
        app_path = "Applications\\"
        applications = []
        for value in self.to_run_applications:
            applications.append(os.path.abspath(app_path + value + ".app"))

        current_time = timecurrent.now()

        # Escrever arquivo com informações do experimento
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
        if tune == 1:
            info_file.write("Sim")
        else:
            info_file.write("Não")
        info_file.write("\n")
        info_file.write("GasFlow: " + str(gas_flow))
        info_file.write("\n")
        info_file.write("Baixas Temperaturas: ")
        if low_temperature[0] == 1:
            info_file.write("Sim")
        else:
            info_file.write("Não")
        info_file.write("\n")
        info_file.close()

        # Tempo Final
        end_time = sum_time(current_time, self.estimated_time_value)
        self.estimated_time = Label(self.frame_10, text="Horário de Término Esperado: " + end_time,
                                    width=50, pady=8, anchor=CENTER)
        self.estimated_time.pack(side=BOTTOM)
        self.time_objects.append(self.estimated_time)

        self.stop_experiment_button.config(text="Parar", command=self.end_experiment)
        #self.stop_experiment_button.config(state=DISABLED)
        self.on_experiment = True

        write_log("Iniciando Experimento")

        self.log_text = Text(self.sub_right_frame_down)
        self.log_text.pack(side=BOTTOM)
        self.time_objects.append(self.log_text)

        self.log_text.config(state=DISABLED)

        self.mpb = ttk.Progressbar(self.sub_right_frame_down_down, orient="horizontal", length=300, mode="determinate")
        self.mpb.pack(side=BOTTOM)
        self.mpb["maximum"] = len(temperatures_BVT)
        self.mpb["value"] = 0
        self.time_objects.append(self.mpb)

        gui_data.put([gas_flow, tune, temperatures_BVT, wait_time, applications, ramps, serialnumber, exe_path, low_temperature, 0])

    def end_experiment(self):

        write_log("Interrupção de experimento")

        gui_data.put("STOP1")
        self.stop_experiment_button.config(state=DISABLED)

    #--- Calibração

    def config_calib(self):
        '''Disponibiliza todos os campos para preenchimento
        a fim de executar um novo experimento.'''

        experiment.connect_serial()

        if experiment.serialcom.connected == False:
            tkMessageBox.showinfo("Aviso", "Porta Serial não Conectada.")
            #return

        if self.on_experiment:
            tkMessageBox.showinfo("Aviso", "Experimento em Execução.")
            return

        write_log("Calibração iniciada")

        self.clean()
        self.screens.append(8)
        self.status.config(text="Definindo calibração.", fg="green")

        # Title
        self.definecalib_label = Label(self.frame_0, text="Definições de Calibração", width=50, pady=8, anchor=CENTER)
        self.definecalib_label.pack(side=TOP)
        self.main_objects.append(self.definecalib_label)

        # Amostra
        self.calib_label = Label(self.frame_2, text="Calibração: ", width=30, pady=8, anchor=W)
        self.calib_entry = Entry(self.frame_2, width=35, bd=3, state=NORMAL)
        self.calib_entry.insert(0, "calibracao")
        self.calib_label.pack(side=LEFT)
        self.calib_entry.pack(side=LEFT)
        self.main_objects.append(self.calib_label)
        self.main_objects.append(self.calib_entry)

        # Path
        self.path_label = Label(self.frame_3, text="Destino: ", width=30, pady=8, anchor=W)
        self.path_entry = Entry(self.frame_3, width=35, bd=3, state=NORMAL)
        self.path_entry.insert(0, "C:\\User\\RafaelFene\\git\\pnmr\\teste_calibracao")
        #self.path_entry.insert(0, "C:/Users/Rafael Fenerick/Documents/Git/pnmr")
        self.pathbutton = Button(self.frame_3, text="Escolher", bd=3, command=self.choose_path, justify=CENTER, width=6)
        self.path_label.pack(side=LEFT)
        self.path_entry.pack(side=LEFT)
        self.pathbutton.pack(side=RIGHT)
        self.main_objects.append(self.path_label)
        self.main_objects.append(self.path_entry)
        self.main_objects.append(self.pathbutton)

        # Tempo Estimado
        self.estimated_time = Label(self.frame_10, text="Tempo Estimado: ", width=50, pady=8,
                                    anchor=CENTER)
        self.estimated_time.pack(side=BOTTOM)
        self.time_objects.append(self.estimated_time)

        # Multiple temperatures
        self.synchro_mult_command = (self.master.register(self.synchro_multiple), '%P', '%W')

        # Temperatura Inicial
        self.init_temperature_label = Label(self.frame_5, text="Temperatura Inicial: [K(BVT) / K / °C]", width=30,
                                            pady=8, anchor=W)
        self.init_temperature_entry = Entry(self.frame_5, width=11, bd=3, state=NORMAL)
        self.init_temperature_entry2 = Entry(self.frame_5, width=12, bd=3, state=NORMAL)
        self.init_temperature_entry3 = Entry(self.frame_5, width=11, bd=3, state=NORMAL)

        self.init_temperature_entry.insert(0, "310")
        self.init_temperature_entry2.insert(0, "306")
        self.init_temperature_entry3.insert(0, "32.85")

        self.init_temperature_entry.config(validate='key', validatecommand=self.synchro_mult_command)
        self.init_temperature_entry2.config(validate='key', validatecommand=self.synchro_mult_command)
        self.init_temperature_entry3.config(validate='key', validatecommand=self.synchro_mult_command)
        self.init_temperature_label.pack(side=LEFT)
        self.init_temperature_entry.pack(side=LEFT)
        self.init_temperature_entry2.pack(side=LEFT)
        self.init_temperature_entry3.pack(side=LEFT)
        self.temperature_objects.append(self.init_temperature_label)
        self.temperature_objects.append(self.init_temperature_entry)
        self.temperature_objects.append(self.init_temperature_entry2)
        self.temperature_objects.append(self.init_temperature_entry3)

        # Step Temperature
        self.step_temperature_label = Label(self.frame_6, text="Passo [K]/ Quantidade: ", width=30, pady=8,
                                            anchor=W)
        self.step_temperature_entry = Entry(self.frame_6, width=17, bd=3, state=NORMAL)

        self.step_temperature_entry.insert(0, "10")

        self.step_temperature_entry.config(validate='key', validatecommand=self.synchro_mult_command)
        self.step_temperature_entry2 = Entry(self.frame_6, width=18, bd=3, validate='key',
                                             validatecommand=self.synchro_mult_command, state=NORMAL)
        self.step_temperature_label.pack(side=LEFT)
        self.step_temperature_entry.pack(side=LEFT)
        self.step_temperature_entry2.pack(side=LEFT)
        self.temperature_objects.append(self.step_temperature_label)
        self.temperature_objects.append(self.step_temperature_entry)
        self.temperature_objects.append(self.step_temperature_entry2)

        # Temperatura Final
        self.end_temperature_label = Label(self.frame_7, text="Temperatura Final: [K(BVT) / K / °C]", width=30,
                                           pady=8, anchor=W)
        self.end_temperature_entry = Entry(self.frame_7, width=11, bd=3, state=NORMAL)
        self.end_temperature_entry2 = Entry(self.frame_7, width=12, bd=3, state=NORMAL)
        self.end_temperature_entry3 = Entry(self.frame_7, width=11, bd=3, state=NORMAL)

        self.end_temperature_entry.insert(0, "340")
        self.end_temperature_entry2.insert(0, "345")
        self.end_temperature_entry3.insert(0, "71.85")

        self.end_temperature_entry.config(validate='key', validatecommand=self.synchro_mult_command)
        self.end_temperature_entry2.config(validate='key', validatecommand=self.synchro_mult_command)
        self.end_temperature_entry3.config(validate='key', validatecommand=self.synchro_mult_command)
        self.end_temperature_label.pack(side=LEFT)
        self.end_temperature_entry.pack(side=LEFT)
        self.end_temperature_entry2.pack(side=LEFT)
        self.end_temperature_entry3.pack(side=LEFT)
        self.temperature_objects.append(self.end_temperature_label)
        self.temperature_objects.append(self.end_temperature_entry)
        self.temperature_objects.append(self.end_temperature_entry2)
        self.temperature_objects.append(self.end_temperature_entry3)

        # Waiting Time
        self.wait_time_mult_label = Label(self.frame_9, text="Tempo de Espera: [min]", width=30, pady=8, anchor=W)
        self.wait_time_mult_entry = Entry(self.frame_9, width=36, bd=3, state=NORMAL)

        self.wait_time_mult_entry.insert(0, "10")

        self.wait_time_mult_entry.config(validate='key', validatecommand=self.synchro_mult_command)
        self.wait_time_mult_label.pack(side=LEFT)
        self.wait_time_mult_entry.pack(side=LEFT)
        self.temperature_objects.append(self.wait_time_mult_label)
        self.temperature_objects.append(self.wait_time_mult_entry)

        # Temperatures from file
        self.file_temperatures_entry = Entry(self.frame_8, width=36, bd=3, state=NORMAL)
        self.file_temperatures_entry.insert(0, "")
        self.file_temperatures_control_var = IntVar()
        self.file_temperatures_checkbox = Checkbutton(self.frame_8, text="Por arquivo",
                                                      variable=self.file_temperatures_control_var,
                                                      onvalue=1, offvalue=0, height=5, width=27,
                                                      command=self.toggle_temperatures_file)
        self.file_temperatures_button = Button(self.frame_8, text="Escolher", bd=3,
                                               command=self.choose_temperatures_file, justify=CENTER, width=6)

        self.file_temperatures_checkbox.pack(side=LEFT)
        self.file_temperatures_entry.pack(side=LEFT)
        self.file_temperatures_button.pack(side=RIGHT)
        self.temperature_objects.append(self.file_temperatures_entry)
        self.temperature_objects.append(self.file_temperatures_checkbox)
        self.temperature_objects.append(self.file_temperatures_button)

        self.step_temperature_entry2.insert(0, "5")

        self.toggle_temperatures_file()

        # Gas Flow
        self.frame_11.config(relief=GROOVE)
        self.temperature_level_var = IntVar()
        self.temperature_level_var.set(1)
        self.gasflow_var_checkbutton = Radiobutton(self.frame_11, text="Altas Temperaturas - Gas Flow:",
                                                       variable=self.temperature_level_var,
                                                       value=1, height=3, width=30)
        self.gasflow_var_checkbutton.pack(side=LEFT)
        self.main_objects.append(self.gasflow_var_checkbutton)

        self.gasflow_entry = Entry(self.frame_11, width=35, bd=3, state=NORMAL)
        self.gasflow_entry.insert(0, "2000")
        self.gasflow_var_checkbutton.pack(side=LEFT)
        self.gasflow_entry.pack(side=LEFT)
        self.main_objects.append(self.gasflow_var_checkbutton)
        self.main_objects.append(self.gasflow_entry)

        #Low Temperature
        self.low_temperature_checkbutton = Radiobutton(self.frame_12, text="Baixas Temperaturas - Potência:",
                                                       variable=self.temperature_level_var,
                                                       value=2, height=3, width=30)
        self.low_temperature_checkbutton.pack(side=LEFT)
        self.main_objects.append(self.low_temperature_checkbutton)

        self.low_temperature_entry = Entry(self.frame_12, width=35, bd=3, state=NORMAL)
        self.low_temperature_entry.insert(0, "10")
        self.low_temperature_entry.pack(side=LEFT)
        self.main_objects.append(self.low_temperature_entry)

        # Ramp
        self.rampmiddle_control_var = IntVar()
        self.rampmiddle_checkbutton = Checkbutton(self.frame_13, text="Rampa em Variações maiores que",
                                                  variable=self.rampmiddle_control_var,
                                                  onvalue=1, offvalue=0, height=3, width=30)
        self.rampmiddle_checkbutton.pack(side=LEFT)
        self.main_objects.append(self.rampmiddle_checkbutton)

        self.rampmiddle_entry = Entry(self.frame_13, width=35, bd=3, state=NORMAL)
        self.rampmiddle_entry.insert(0, "10")
        self.rampmiddle_entry.pack(side=LEFT)
        self.main_objects.append(self.rampmiddle_entry)

        # Tune
        self.tune_control_var = IntVar()
        self.tune_checkbutton = Checkbutton(self.frame_14, text="Tune", variable=self.tune_control_var,
                                            onvalue=1, offvalue=0, height=3, width=10)
        self.tune_checkbutton.pack(side=LEFT)
        self.main_objects.append(self.tune_checkbutton)

        # Application
        self.application_label = Label(self.sub_right_frame_up, text="Escolha a Aplicação", width=50, pady=8, anchor=CENTER)
        self.application_label.pack(side=TOP)
        self.main_objects.append(self.application_label)

        #Main Buttons
        self.start_experiment_button = Button(self.frame_last, text="Iniciar", bd=3, command=self.check_start_calib, justify=CENTER, width=10)
        self.start_experiment_button.pack(side=RIGHT)
        self.main_objects.append(self.start_experiment_button)

        self.stop_experiment_button = Button(self.frame_last, text="Cancelar", bd=3, command=self.clean, justify=CENTER, width=10)
        self.stop_experiment_button.pack(side=RIGHT)
        self.cancel_objects.append(self.stop_experiment_button)

        write_log("Configuração de experimento finalizada")

    def check_start_calib(self):
        '''Executa checks para confirmar que todos os parâmetros
        estão preenchidos corretamente e continua com a
        execução do experimento.'''

        write_log("Checagem de início de calibração iniciada")

        if self.path_entry.get() == "" or self.path_entry.get() == "C:\\":
            tkMessageBox.showinfo("Aviso", "Caminho não definido!")
            return

        if not os.path.isdir(self.path_entry.get()):
            tkMessageBox.showinfo("Aviso", "Caminho especificado não existe!")
            return

        if self.file_temperatures_control_var.get() == 0:
            if self.init_temperature_entry.get() == "":
                tkMessageBox.showinfo("Aviso", "Temeperatura inicial não definida!")
                return
            if self.end_temperature_entry.get() == "":
                tkMessageBox.showinfo("Aviso", "Temeperatura final não definida!")
                return
            if self.step_temperature_entry.get() == "":
                tkMessageBox.showinfo("Aviso", "Passo não definido!")
                return
            if self.step_temperature_entry2.get() == "":
                tkMessageBox.showinfo("Aviso", "Quantidade de experimentos não definida!")
                return
        if self.file_temperatures_control_var.get() == 1:
            if self.file_temperatures_entry.get() == "":
                tkMessageBox.showinfo("Aviso", "Nenhum arquivo selecionado!")
                return
            if self.file_temperatures_entry.get()[len(self.file_temperatures_entry.get())-3:] != "txt":
                tkMessageBox.showinfo("Aviso", "Formato de arquivo incorreto!")
                return

            if self.wait_time_mult_entry.get() == "":
                tkMessageBox.showinfo("Aviso", "Tempo de espera não definido!")
                return

        if self.gasflow_entry.get() == "":
            tkMessageBox.showinfo("Aviso", "Gas Flow não definido!")
            return

        write_log("Checagem de início de calibração finalizada")

        self.start_calib()

    def start_calib(self):
        '''Executa o experimento especificado,
        dada a amostra e o range de temperatura
        escolhidos.'''

        write_log("Inicialização de calibração iniciada")

        self.disable(1)
        self.disable(2)
        self.clean(4)

        self.screens.append(9)
        self.status.config(text="Executando calibração.", fg="red")

        temperatures_BVT = []
        temperatures_amostra = []
        temperatures_amostra_celsius = []

        if self.temperature_level_var.get() == 1:
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

        ramps = [self.rampmiddle_control_var.get(), float(self.rampmiddle_entry.get())]
        wait_time = float(self.wait_time_mult_entry.get())
        if self.file_temperatures_control_var.get() == 1:
            for temperature in self.file_temperatures_array:
                if filetemp == 1:
                    temperatures_BVT.append(temperature)
                    temperatures_amostra.append(round(float((temperature + b)/a), 2))
                    temperatures_amostra_celsius.append(round(float((temperature + b)/a) - 273.15, 2))
                elif filetemp == 2:
                    temperatures_BVT.append(round(float(temperature*a - b), 2))
                    temperatures_amostra.append(temperature)
                    temperatures_amostra_celsius.append(round(temperature - 273.15, 2))
                else:
                    temperatures_BVT.append(round(float((temperature + 273.15)*a - b), 2))
                    temperatures_amostra.append(round(temperature + 273.15, 2))
                    temperatures_amostra_celsius.append(temperature)

        else:
            init = float(self.init_temperature_entry.get())
            end = float(self.end_temperature_entry.get())
            step = float(self.step_temperature_entry.get())
            amount = float(self.step_temperature_entry2.get())
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

        amostra = self.calib_entry.get()
        to_save_path = self.path_entry.get()
        if to_save_path[len(to_save_path)-1] != "\\":
            to_save_path += "\\"
        tune = int(self.tune_control_var.get())
        gas_flow = int(self.gasflow_entry.get())
        low_temperature = [int(self.temperature_level_var.get())-1, float(self.low_temperature_entry.get())]
        low_temperature[0] = True if low_temperature[0] == 1 else False

        current_time = timecurrent.now()

        # Escrever arquivo com informações do experimento
        info_file = open(to_save_path + "\\" + amostra + " calibracao infos.txt", "w")
        info_file.write("Amostra: " + amostra + "\n")
        date = str(current_time.day) + "/" + str(current_time.month) + "/" + str(current_time.year)
        time = str(current_time.hour) + ":" + str(current_time.minute)
        info_file.write(date + " " + time + "\n")
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
        if tune == 1:
            info_file.write("Sim")
        else:
            info_file.write("Não")
        info_file.write("\n")
        info_file.write("GasFlow: " + str(gas_flow))
        info_file.write("\n")
        info_file.write("Baixas Temperaturas: ")
        if low_temperature[0] == 1:
            info_file.write("Sim")
        else:
            info_file.write("Não")
        info_file.write("\n")
        info_file.close()

        # Tempo Final
        end_time = sum_time(current_time, self.estimated_time_value)
        self.estimated_time = Label(self.frame_10, text="Horário de Término Esperado: " + end_time,
                                    width=50, pady=8, anchor=CENTER)
        self.estimated_time.pack(side=BOTTOM)
        self.time_objects.append(self.estimated_time)

        self.stop_experiment_button.config(text="Parar", command=self.end_calib)
        #self.stop_experiment_button.config(state=DISABLED)
        self.on_experiment = True

        write_log("Iniciando Calibração")

        self.log_text = Text(self.sub_right_frame_down)
        self.log_text.pack(side=BOTTOM)
        self.time_objects.append(self.log_text)

        self.log_text.config(state=DISABLED)

        self.mpb = ttk.Progressbar(self.sub_right_frame_down_down, orient="horizontal", length=300, mode="determinate")
        self.mpb.pack(side=BOTTOM)
        self.mpb["maximum"] = len(temperatures_BVT)
        self.mpb["value"] = 0
        self.time_objects.append(self.mpb)

        gui_data.put([gas_flow, tune, temperatures_BVT, wait_time, [], ramps, serialnumber, exe_path, low_temperature, 1])

    def end_calib(self):

        write_log("Interrupção de experimento")

        gui_data.put("STOP2")
        self.stop_experiment_button.config(state=DISABLED)

    #--- Application

    def config_add_application(self):
        '''Disponibiliza os parâmetros para
         adicionar uma nova aplicação à
        lista de aplicações.'''

        if self.on_experiment:
            tkMessageBox.showinfo("Aviso", "Experimento em Execução.")
            return

        self.clean()
        self.screens.append(3)
        self.status.config(text="Adicionando Aplicação.")

        # Title
        self.application_label = Label(self.frame_0, text="Nova Aplicação", width=50, pady=8, anchor=CENTER)
        self.application_label.pack(side=TOP)
        self.main_objects.append(self.application_label)

        # Amostra
        self.appname_label = Label(self.frame_2, text="Aplicação: ", width=20, pady=8, anchor=W)
        self.appname_entry = Entry(self.frame_2, width=35, bd=3, state=NORMAL)
        self.appname_entry.insert(0, "")
        self.appname_label.pack(side=LEFT)
        self.appname_entry.pack(side=LEFT)
        self.main_objects.append(self.appname_label)
        self.main_objects.append(self.appname_entry)

        # Path
        self.apppath_label = Label(self.frame_3, text="Arquivo: ", width=20, pady=8, anchor=W)
        self.apppath_entry = Entry(self.frame_3, width=35, bd=3, state=NORMAL)
        self.apppath_label.pack(side=LEFT)
        self.apppath_entry.pack(side=LEFT)
        self.main_objects.append(self.apppath_label)
        self.main_objects.append(self.apppath_entry)

        self.apppath_button = Button(self.frame_4, text="Escolher", bd=3, command=self.choose_app_file, justify=CENTER, width=6)
        self.apppath_button.pack(side=RIGHT)
        self.main_objects.append(self.apppath_button)

        # Main Buttons
        self.addapp_button = Button(self.frame_last, text="Adicionar", bd=3, command=self.add_application, justify=CENTER, width=10)
        self.addapp_button.pack(side=RIGHT)
        self.main_objects.append(self.addapp_button)

        self.cancel_addapp_button = Button(self.frame_last, text="Cancelar", bd=3, command=self.clean, justify=CENTER, width=10)
        self.cancel_addapp_button.pack(side=RIGHT)
        self.main_objects.append(self.cancel_addapp_button)

    def config_remove_application(self):
        '''Remove uma aplicação previamente adicionada.'''

        if self.on_experiment:
            tkMessageBox.showinfo("Aviso", "Experimento em Execução.")
            return

        self.clean()
        self.screens.append(4)
        self.status.config(text="Removendo Aplicação.")

        # Application
        self.remove_application_label = Label(self.sub_right_frame_up, text="Escolha as Aplicaçôes", width=50, pady=8,
                                       anchor=CENTER)
        self.remove_application_label.pack(side=TOP)
        self.main_objects.append(self.remove_application_label)

        # Get stored application
        try:
            file = open("Applications\\Applications.txt", "r")
            self.remove_applications_text = file.readlines()
            file.close()
        except:
            self.remove_applications_text = []
        self.remove_application_control_var = []
        self.remove_applications = []

        if len(self.remove_applications_text) == 0:
            tkMessageBox.showinfo("Aviso", "Nenhuma aplicação adicionada!")
            self.clean()
            return

        i = 1
        for application in self.remove_applications_text:
            self.remove_application_control_var.append(IntVar())
            application_checkbutton = Checkbutton(self.sub_right_frame_center, text=application.strip("\n"),
                                                  variable=self.remove_application_control_var[
                                                      len(self.remove_application_control_var) - 1],
                                                  onvalue=i, offvalue=0, padx=40)
            application_checkbutton.pack(anchor=W)
            self.main_objects.append(application_checkbutton)
            self.remove_applications.append(application_checkbutton)
            i += 1

        # Main Buttons
        self.remove_start_experiment_button = Button(self.frame_last, text="Remover", bd=3,
                                              command=self.remove_application, justify=CENTER, width=10)
        self.remove_start_experiment_button.pack(side=RIGHT)
        self.main_objects.append(self.remove_start_experiment_button)

        self.remove_stop_experiment_button = Button(self.frame_last, text="Cancelar", bd=3, command=self.clean, justify=CENTER,
                                             width=10)
        self.remove_stop_experiment_button.pack(side=RIGHT)
        self.cancel_objects.append(self.remove_stop_experiment_button)

    def add_application(self):
        '''Adicionar uma nova aplicação à
        lista de aplicações.'''

        if self.appname_entry.get() == "":
            tkMessageBox.showinfo("Aviso", "Nome para a aplicação não definido!")
            return

        if len(self.apppath_entry.get()) == 0:
            tkMessageBox.showinfo("Aviso", "Caminho da aplicação não definido!")
            return

        if self.apppath_entry.get()[len(self.apppath_entry.get())-3:] != "app":
            tkMessageBox.showinfo("Aviso", "Tipo de aplicação não coerente!")
            return

        # Copy the file to Application path
        file = open("Applications\\Applications.txt", "a")
        try:
            copyfile(self.apppath_entry.get(), "Applications\\" + self.appname_entry.get() + ".app")
            file.write(self.appname_entry.get() + "\n")
        except:
            tkMessageBox.showinfo("Aviso", "Não foi possível acessar o arquivo especificado!")
            return
        file.close()

        tkMessageBox.showinfo("Aplicação", "Aplicação salva com sucesso!")

    def remove_application(self):

        try:
            try:
                file = open("Applications\\Applications.txt", "r")
                remove_applications_text = file.readlines()
                file.close()
            except:
                remove_applications_text = []

            if len(remove_applications_text) == 0:
                tkMessageBox.showinfo("Aviso", "Nenhuma aplicação adicionada!")
                self.clean()
                return

            to_remove_applications = []
            for value in self.remove_application_control_var:
                if value.get() != 0:
                    to_remove_applications.append(self.remove_applications_text[value.get() - 1].strip("\n"))

            for value in to_remove_applications:
                os.remove("Applications\\" + value + ".app")

            file = open("Applications\\Applications.txt", "w")
            for value in remove_applications_text:
                if value.strip("\n") not in to_remove_applications:
                    file.write(value)

            file.close()
        except:
            tkMessageBox.showinfo("Aplicação", "Falha ao remover aplicações!")
            self.clean()
            return

        tkMessageBox.showinfo("Aplicação", "Aplicações removidas com sucesso!")
        self.clean()

    #--- Parameters

    def config_temperature_calculus(self):
        '''Disponibiliza os parâmetros para
         adicionar uma nova aplicação à
        lista de aplicações.'''

        if self.on_experiment:
            tkMessageBox.showinfo("Aviso", "Experimento em Execução.")
            return

        self.clean()
        self.screens.append(5)
        self.status.config(text="Configuração de parâmetros de Temperatura.")

        # Title
        self.temp_param_label = Label(self.frame_0, text="Parâmetros de Temperatura", width=50, pady=8, anchor=CENTER)
        self.temp_param_label.pack(side=TOP)
        self.main_objects.append(self.temp_param_label)

        file = open("Temp_params.txt", "r")
        lines = file.readlines()
        file.close()
        a, b = float(lines[0].strip("\n")), float(lines[1].strip("\n"))

        # Amostra
        self.a_temp_label = Label(self.frame_2, text="a: ", width=20, pady=8, anchor=W)
        self.a_temp_entry = Entry(self.frame_2, width=35, bd=3, state=NORMAL)
        self.a_temp_entry.insert(0, str(a))
        self.a_temp_label.pack(side=LEFT)
        self.a_temp_entry.pack(side=LEFT)
        self.main_objects.append(self.a_temp_label)
        self.main_objects.append(self.a_temp_entry)

        # Path
        self.b_temp_label = Label(self.frame_3, text="b: ", width=20, pady=8, anchor=W)
        self.b_temp_entry = Entry(self.frame_3, width=35, bd=3, state=NORMAL)
        self.b_temp_entry.insert(0, str(b))
        self.b_temp_label.pack(side=LEFT)
        self.b_temp_entry.pack(side=LEFT)
        self.main_objects.append(self.b_temp_label)
        self.main_objects.append(self.b_temp_entry)

        # Main Buttons
        self.temp_param_button = Button(self.frame_last, text="Salvar", bd=3, command=self.save_temperature_parameters, justify=CENTER, width=10)
        self.temp_param_button.pack(side=RIGHT)
        self.main_objects.append(self.temp_param_button)

        self.cancel_temp_param_button = Button(self.frame_last, text="Cancelar", bd=3, command=self.clean, justify=CENTER, width=10)
        self.cancel_temp_param_button.pack(side=RIGHT)
        self.main_objects.append(self.cancel_temp_param_button)

    def save_temperature_parameters(self):

        try:
            a = float(self.a_temp_entry.get())
            b = float(self.b_temp_entry.get())
        except:
            tkMessageBox.showinfo("Aplicação", "Não foi possível salvar os parâmetros!")
            return

        file = open("Temp_params.txt", "w")
        file.write(str(a))
        file.write("\n")
        file.write(str(b))
        file.close()
        tkMessageBox.showinfo("Aplicação", "Parâmetros salvos com sucesso!")

    def config_low_temperature_calculus(self):
        '''Disponibiliza os parâmetros para
         adicionar uma nova aplicação à
        lista de aplicações.'''

        if self.on_experiment:
            tkMessageBox.showinfo("Aviso", "Experimento em Execução.")
            return

        self.clean()
        self.screens.append(5)
        self.status.config(text="Configuração de parâmetros de Temperatura.")

        # Title
        self.temp_low_param_label = Label(self.frame_0, text="Parâmetros de Temperatura", width=50, pady=8, anchor=CENTER)
        self.temp_low_param_label.pack(side=TOP)
        self.main_objects.append(self.temp_low_param_label)

        file = open("Temp_params_low.txt", "r")
        lines = file.readlines()
        file.close()
        a, b = float(lines[0].strip("\n")), float(lines[1].strip("\n"))

        # Amostra
        self.a_temp_low_label = Label(self.frame_2, text="a: ", width=20, pady=8, anchor=W)
        self.a_temp_low_entry = Entry(self.frame_2, width=35, bd=3, state=NORMAL)
        self.a_temp_low_entry.insert(0, str(a))
        self.a_temp_low_label.pack(side=LEFT)
        self.a_temp_low_entry.pack(side=LEFT)
        self.main_objects.append(self.a_temp_low_label)
        self.main_objects.append(self.a_temp_low_entry)

        # Path
        self.b_temp_low_label = Label(self.frame_3, text="b: ", width=20, pady=8, anchor=W)
        self.b_temp_low_entry = Entry(self.frame_3, width=35, bd=3, state=NORMAL)
        self.b_temp_low_entry.insert(0, str(b))
        self.b_temp_low_label.pack(side=LEFT)
        self.b_temp_low_entry.pack(side=LEFT)
        self.main_objects.append(self.b_temp_low_label)
        self.main_objects.append(self.b_temp_low_entry)

        # Main Buttons
        self.temp_low_param_button = Button(self.frame_last, text="Salvar", bd=3, command=self.save_low_temperature_parameters, justify=CENTER, width=10)
        self.temp_low_param_button.pack(side=RIGHT)
        self.main_objects.append(self.temp_low_param_button)

        self.cancel_temp_low_param_button = Button(self.frame_last, text="Cancelar", bd=3, command=self.clean, justify=CENTER, width=10)
        self.cancel_temp_low_param_button.pack(side=RIGHT)
        self.main_objects.append(self.cancel_temp_low_param_button)

    def save_low_temperature_parameters(self):

        try:
            a = float(self.a_temp_low_entry.get())
            b = float(self.b_temp_low_entry.get())
        except:
            tkMessageBox.showinfo("Aplicação", "Não foi possível salvar os parâmetros!")
            return

        file = open("Temp_params_low.txt", "w")
        file.write(str(a))
        file.write("\n")
        file.write(str(b))
        file.close()
        tkMessageBox.showinfo("Aplicação", "Parâmetros salvos com sucesso!")

    #--- Equipment

    def config_equipment_info(self):
        '''Disponibiliza os parâmetros para
         adicionar uma nova aplicação à
        lista de aplicações.'''

        if self.on_experiment:
            tkMessageBox.showinfo("Aviso", "Experimento em Execução.")
            return

        self.clean()
        self.screens.append(6)
        self.status.config(text="Configuração de parâmetros do equipamento.")

        # Title
        self.equip_param_label = Label(self.frame_0, text="Parâmetros do Equipamento", width=50, pady=8, anchor=CENTER)
        self.equip_param_label.pack(side=TOP)
        self.main_objects.append(self.equip_param_label)

        file = open("Equip_params.txt", "r")
        lines = file.readlines()
        file.close()
        exe_path, serial = str(lines[0].strip("\n")), str(lines[1].strip("\n"))

        # Amostra
        self.equip_exe_label = Label(self.frame_2, text="Caminho do Software: ", width=20, pady=8, anchor=W)
        self.equip_exe_entry = Entry(self.frame_2, width=35, bd=3, state=NORMAL)
        self.equip_exe_entry.insert(0, exe_path)
        self.equippathbutton = Button(self.frame_2, text="Escolher", bd=3, command=self.choose_equip_path, justify=CENTER, width=6)
        self.equip_exe_label.pack(side=LEFT)
        self.equip_exe_entry.pack(side=LEFT)
        self.equippathbutton.pack(side=RIGHT)
        self.main_objects.append(self.equip_exe_label)
        self.main_objects.append(self.equip_exe_entry)
        self.main_objects.append(self.equippathbutton)

        # Path
        self.equip_serial_label = Label(self.frame_3, text="Número Serial: ", width=20, pady=8, anchor=W)
        self.equip_serial_entry = Entry(self.frame_3, width=35, bd=3, state=NORMAL)
        self.equip_serial_entry.insert(0, serial)
        self.equip_serial_label.pack(side=LEFT)
        self.equip_serial_entry.pack(side=LEFT)
        self.main_objects.append(self.equip_serial_label)
        self.main_objects.append(self.equip_serial_entry)

        # Main Buttons
        self.temp_param_button = Button(self.frame_last, text="Salvar", bd=3, command=self.save_equipment_info, justify=CENTER, width=10)
        self.temp_param_button.pack(side=RIGHT)
        self.main_objects.append(self.temp_param_button)

        self.cancel_temp_param_button = Button(self.frame_last, text="Cancelar", bd=3, command=self.clean, justify=CENTER, width=10)
        self.cancel_temp_param_button.pack(side=RIGHT)
        self.main_objects.append(self.cancel_temp_param_button)

    def save_equipment_info(self):

        if not os.path.isdir(self.equip_exe_entry.get()):
            tkMessageBox.showinfo("Aplicação", "Não foi possível salvar os parâmetros!")
            return

        file = open("Equip_params.txt", "w")
        file.write(self.equip_exe_entry.get())
        file.write("\n")
        file.write(self.equip_serial_entry.get())
        file.close()
        tkMessageBox.showinfo("Aplicação", "Parâmetros salvos com sucesso!")

    #--- file temperature

    def config_filetemp_info(self):
        '''Disponibiliza os parâmetros para
         adicionar uma nova aplicação à
        lista de aplicações.'''

        if self.on_experiment:
            tkMessageBox.showinfo("Aviso", "Experimento em Execução.")
            return

        self.clean()
        self.screens.append(7)
        self.status.config(text="Configuração de leitura de temperatura por arquivo.")

        # Title
        self.filetemp_label = Label(self.frame_0, text="Parâmetros de Leitura de Temperaturas", width=50, pady=8, anchor=CENTER)
        self.filetemp_label.pack(side=TOP)
        self.main_objects.append(self.filetemp_label)

        file = open("Filetemp_params.txt", "r")
        lines = file.readlines()
        file.close()
        filetemp = int(lines[0].strip("\n"))

        modes = [("BVT (K)", 1), ("Amostra (K)", 2), ("Amostra (°C)", 3)]

        self.filetemp_control_var = IntVar()
        self.filetemp_control_var.set(filetemp)

        for text, mode in modes:
            radiobutton = Radiobutton(self.sub_right_frame_center, text=text, variable=self.filetemp_control_var, value=mode)
            radiobutton.pack(anchor=W)
            self.main_objects.append(radiobutton)

        # Main Buttons
        self.temp_param_button = Button(self.frame_last, text="Salvar", bd=3, command=self.save_filetemp_info, justify=CENTER, width=10)
        self.temp_param_button.pack(side=RIGHT)
        self.main_objects.append(self.temp_param_button)

        self.cancel_temp_param_button = Button(self.frame_last, text="Cancelar", bd=3, command=self.clean, justify=CENTER, width=10)
        self.cancel_temp_param_button.pack(side=RIGHT)
        self.main_objects.append(self.cancel_temp_param_button)

    def save_filetemp_info(self):

        file = open("Filetemp_params.txt", "w")
        file.write(str(self.filetemp_control_var.get()) + "\n")
        file.write("\n")
        file.close()
        tkMessageBox.showinfo("Aplicação", "Parâmetro salvo com sucesso!")

    #--- Support Funcitons

    ### Synchro

    def synchro_multiple(self, new_text, entry):
        if new_text != "":
            try:
                float(new_text)
            except:
                return False

        file = open("Temp_params.txt", "r")
        lines = file.readlines()
        file.close()
        a, b = float(lines[0].strip("\n")), float(lines[1].strip("\n"))

        to_change = 0
        if entry == "":
            to_change = -1

        init = self.init_temperature_entry.get()
        if str(entry) == str(self.init_temperature_entry):
            init = new_text
        init = float(init) if init != "" else init

        init2 = self.init_temperature_entry2.get()
        if str(entry) == str(self.init_temperature_entry2):
            to_change = 3
            init2 = new_text
        init2 = float(init2) if init2 != "" else init2

        init3 = self.init_temperature_entry3.get()
        if str(entry) == str(self.init_temperature_entry3):
            to_change = 4
            init3 = new_text
        init3 = float(init3) if init3 != "" else init3

        end = self.end_temperature_entry.get()
        if str(entry) == str(self.end_temperature_entry):
            end = new_text
        end = float(end) if end != "" else end

        end2 = self.end_temperature_entry2.get()
        if str(entry) == str(self.end_temperature_entry2):
            end2 = new_text
            to_change = 3
        end2 = float(end2) if end2 != "" else end2

        end3 = self.end_temperature_entry3.get()
        if str(entry) == str(self.end_temperature_entry3):
            end3 = new_text
            to_change = 4
        end3 = float(end3) if end3 != "" else end3

        step = self.step_temperature_entry.get()
        if str(entry) == str(self.step_temperature_entry):
            step = new_text
            to_change = 1
        step = float(step) if step != "" else step

        amount = self.step_temperature_entry2.get()
        if str(entry) == str(self.step_temperature_entry2):
            amount = new_text
            to_change = 2
        amount = float(amount) if amount != "" else amount

        wait_time = self.wait_time_mult_entry.get()
        if str(entry) == str(self.wait_time_mult_entry):
            wait_time = new_text
            to_change = -1
        wait_time = float(wait_time) if wait_time != "" else wait_time

        if to_change == 3:

            if init2 != "":
                init = round(init2*a - b, 2)
                init3 = init2 - 273.15

                self.init_temperature_entry.config(validate='none')
                if len(self.init_temperature_entry.get()) != 0:
                    self.init_temperature_entry.delete(0, "end")
                self.init_temperature_entry.insert(0, str(init))
                self.init_temperature_entry.config(validate='key')

                self.init_temperature_entry3.config(validate='none')
                if len(self.init_temperature_entry3.get()) != 0:
                    self.init_temperature_entry3.delete(0, "end")
                self.init_temperature_entry3.insert(0, str(init3))
                self.init_temperature_entry3.config(validate='key')

            if end2 != "":
                end = round(end2*a - b, 2)
                end3 = end2 - 273.15

                self.end_temperature_entry.config(validate='none')
                if len(self.end_temperature_entry.get()) != 0:
                    self.end_temperature_entry.delete(0, "end")
                self.end_temperature_entry.insert(0, str(end))
                self.end_temperature_entry.config(validate='key')

                self.end_temperature_entry3.config(validate='none')
                if len(self.end_temperature_entry3.get()) != 0:
                    self.end_temperature_entry3.delete(0, "end")
                self.end_temperature_entry3.insert(0, str(end3))
                self.end_temperature_entry3.config(validate='key')

            if init != "" and end != "" and amount != "":
                if amount == 0:
                    return False
                if amount == 1:
                    return True
                step = round(float(abs(end - init)/float(amount - 1)), 2)

                self.step_temperature_entry.config(validate='none')
                if len(self.step_temperature_entry.get()) != 0:
                    self.step_temperature_entry.delete(0, "end")
                self.step_temperature_entry.insert(0, str(step))
                self.step_temperature_entry.config(validate='key')

        if to_change == 4:

            if init3 != "":
                init2 = init3 + 273.15
                init = round(init2*a - b, 2)


                self.init_temperature_entry.config(validate='none')
                if len(self.init_temperature_entry.get()) != 0:
                    self.init_temperature_entry.delete(0, "end")
                self.init_temperature_entry.insert(0, str(init))
                self.init_temperature_entry.config(validate='key')

                self.init_temperature_entry2.config(validate='none')
                if len(self.init_temperature_entry2.get()) != 0:
                    self.init_temperature_entry2.delete(0, "end")
                self.init_temperature_entry2.insert(0, str(init2))
                self.init_temperature_entry2.config(validate='key')

            if end3 != "":
                end2 = end3 + 273.15
                end = round(end2*a - b, 2)


                self.end_temperature_entry.config(validate='none')
                if len(self.end_temperature_entry.get()) != 0:
                    self.end_temperature_entry.delete(0, "end")
                self.end_temperature_entry.insert(0, str(end))
                self.end_temperature_entry.config(validate='key')

                self.end_temperature_entry2.config(validate='none')
                if len(self.end_temperature_entry2.get()) != 0:
                    self.end_temperature_entry2.delete(0, "end")
                self.end_temperature_entry2.insert(0, str(end2))
                self.end_temperature_entry2.config(validate='key')

            if init != "" and end != "" and amount != "":
                if amount == 0:
                    return False
                if amount == 1:
                    return True
                step = round(float(abs(end - init)/float(amount-1)), 2)

                self.step_temperature_entry.config(validate='none')
                if len(self.step_temperature_entry.get()) != 0:
                    self.step_temperature_entry.delete(0, "end")
                self.step_temperature_entry.insert(0, str(step))
                self.step_temperature_entry.config(validate='key')

        if to_change == 0 or to_change == 2:

            if init != "":
                init2 = round(((init + b) / a), 2)
                init3 = init2 - 273.15

                self.init_temperature_entry3.config(validate='none')
                if len(self.init_temperature_entry3.get()) != 0:
                    self.init_temperature_entry3.delete(0, "end")
                self.init_temperature_entry3.insert(0, str(init3))
                self.init_temperature_entry3.config(validate='key')

                self.init_temperature_entry2.config(validate='none')
                if len(self.init_temperature_entry2.get()) != 0:
                    self.init_temperature_entry2.delete(0, "end")
                self.init_temperature_entry2.insert(0, str(init2))
                self.init_temperature_entry2.config(validate='key')

            if end != "":
                end2 = round(((end + b) / a), 2)
                end3 = end2 - 273.15

                self.end_temperature_entry3.config(validate='none')
                if len(self.end_temperature_entry3.get()) != 0:
                    self.end_temperature_entry3.delete(0, "end")
                self.end_temperature_entry3.insert(0, str(end3))
                self.end_temperature_entry3.config(validate='key')

                self.end_temperature_entry2.config(validate='none')
                if len(self.end_temperature_entry2.get()) != 0:
                    self.end_temperature_entry2.delete(0, "end")
                self.end_temperature_entry2.insert(0, str(end2))
                self.end_temperature_entry2.config(validate='key')

            if init != "" and end != "" and amount != "":
                if amount == 0:
                    return False
                if amount == 1:
                    return True
                step = round(float(abs(end - init)/float(amount-1)), 2)

                self.step_temperature_entry.config(validate='none')
                if len(self.step_temperature_entry.get()) != 0:
                    self.step_temperature_entry.delete(0, "end")
                self.step_temperature_entry.insert(0, str(step))
                self.step_temperature_entry.config(validate='key')

        if to_change == 1:

            if init != "" and end != "" and step != "":
                if step == 0:
                    return True
                amount = int(abs(end - init) / step) + 1

                self.step_temperature_entry2.config(validate='none')
                if len(self.step_temperature_entry2.get()) != 0:
                    self.step_temperature_entry2.delete(0, "end")
                self.step_temperature_entry2.insert(0, str(amount))
                self.step_temperature_entry2.config(validate='key')

        if wait_time != "":
            if self.file_temperatures_control_var.get() == 0:
                if amount != "":
                    estimated_time = wait_time * amount
                else:
                    estimated_time = 0
            else:
                file_name = self.file_temperatures_entry.get()
                try:
                    temperatures_file = open(file_name, "r")
                    temperatures = temperatures_file.readlines()
                    temperatures_amount = len(temperatures)
                    temperatures_file.close()
                except:
                    temperatures_amount = 0

                estimated_time = wait_time * temperatures_amount

            self.estimated_time_value = estimated_time
            self.estimated_time.config(text="Tempo Estimado: " + float_to_time(estimated_time))

        return True

    def synchro_single(self, new_text, entry):
        if new_text != "":
            try:
                float(new_text)
            except:
                return False

        file = open("Temp_params.txt", "r")
        lines = file.readlines()
        file.close()
        a, b = float(lines[0].strip("\n")), float(lines[1].strip("\n"))

        to_change = 0

        init = self.temperature_entry.get()
        if str(entry) == str(self.temperature_entry):
            init = new_text
        init = float(init) if init != "" else init

        init2 = self.temperature_entry2.get()
        if str(entry) == str(self.temperature_entry2):
            to_change = 1
            init2 = new_text
        init2 = float(init2) if init2 != "" else init2

        init3 = self.temperature_entry3.get()
        if str(entry) == str(self.temperature_entry3):
            to_change = 2
            init3 = new_text
        init3 = float(init3) if init3 != "" else init3

        wait_time = self.wait_time_single_entry.get()
        if str(entry) == str(self.wait_time_single_entry):
            wait_time = new_text
            to_change = -1
        wait_time = float(wait_time) if wait_time != "" else wait_time

        if to_change == 0 or to_change == -1:
            if init != "":
                init2 = round((init + b)/a, 2)
                init3 = init2 - 273.15

                self.temperature_entry2.config(validate='none')
                if len(self.temperature_entry2.get()) != 0:
                    self.temperature_entry2.delete(0, "end")
                self.temperature_entry2.insert(0, str(init2))
                self.temperature_entry2.config(validate='key')

                self.temperature_entry3.config(validate='none')
                if len(self.temperature_entry3.get()) != 0:
                    self.temperature_entry3.delete(0, "end")
                self.temperature_entry3.insert(0, str(init3))
                self.temperature_entry3.config(validate='key')

        if to_change == 1:
            if init2 != "":
                init = round(init2 * a - b, 2)
                init3 = init2 - 273.15

                self.temperature_entry.config(validate='none')
                if len(self.temperature_entry.get()) != 0:
                    self.temperature_entry.delete(0, "end")
                self.temperature_entry.insert(0, str(init))
                self.temperature_entry.config(validate='key')

                self.temperature_entry3.config(validate='none')
                if len(self.temperature_entry3.get()) != 0:
                    self.temperature_entry3.delete(0, "end")
                self.temperature_entry3.insert(0, str(init3))
                self.temperature_entry3.config(validate='key')

        if to_change == 2:
            if init3 != "":
                init2 = init3 + 273.15
                init = round(init2*a - b, 2)

                self.temperature_entry.config(validate='none')
                if len(self.temperature_entry.get()) != 0:
                    self.temperature_entry.delete(0, "end")
                self.temperature_entry.insert(0, str(init))
                self.temperature_entry.config(validate='key')

                self.temperature_entry2.config(validate='none')
                if len(self.temperature_entry2.get()) != 0:
                    self.temperature_entry2.delete(0, "end")
                self.temperature_entry2.insert(0, str(init2))
                self.temperature_entry2.config(validate='key')

        if to_change == -1:
            if wait_time != "":
                self.estimated_time_value = wait_time
                self.estimated_time.config(text="Tempo Estimado: " + float_to_time(wait_time))

        return True

    ### Files

    def choose_path(self):
        '''Permite ao usuário escolher um caminho específico.'''

        directory = tkFileDialog.askdirectory(initialdir='.')
        if len(self.path_entry.get()) != 0:
            self.path_entry.delete(0, "end")
        self.path_entry.insert(0, directory)

    def choose_app_file(self):
        '''Permite ao usuário escolher um arquivo específico.'''

        filename = tkFileDialog.askopenfilename()
        if len(self.appname_entry.get()) != 0:
            self.apppath_entry.delete(0, "end")
        self.apppath_entry.insert(0, filename)

    def choose_equip_path(self):
        '''Permite ao usuário escolher um caminho específico.'''

        directory = tkFileDialog.askdirectory(initialdir='.')
        if len(self.equip_exe_entry.get()) != 0:
            self.equip_exe_entry.delete(0, "end")
        self.equip_exe_entry.insert(0, directory)

    def choose_temperatures_file(self):
        '''Permite ao usuário escolher um arquivo específico.'''

        filename = tkFileDialog.askopenfilename()

        if filename == "":
            tkMessageBox.showinfo("Aviso", "Nome do arquivo não definido!")
            return

        if filename[len(filename)-3:] != "txt":
            tkMessageBox.showinfo("Aviso", "Tipo de arquivo não coerente!")
            return

        try:
            file = open(filename, "r")
            lines = file.readlines()
            self.file_temperatures_array = []
            for line in lines:
                self.file_temperatures_array.append(float(line.strip("\n")))
            file.close()
        except:
            tkMessageBox.showinfo("Aviso", "Arquivo não pode ser lido!")
            return

        if len(self.file_temperatures_entry.get()) != 0:
            self.file_temperatures_entry.delete(0, "end")
        self.file_temperatures_entry.insert(0, filename)

        file_name = self.file_temperatures_entry.get()
        try:
            temperatures_file = open(file_name, "r")
            temperatures = temperatures_file.readlines()
            temperatures_amount = len(temperatures)
            temperatures_file.close()
        except:
            temperatures_amount = 0

        estimated_time = float(self.wait_time_mult_entry.get()) * temperatures_amount
        self.estimated_time_value = estimated_time
        self.estimated_time.config(text="Tempo Estimado: " + float_to_time(estimated_time))

    ### Handles

    def enable(self, temp, array=0):
        '''Torna todos os objetos possíveis de alteração.'''

        if array == 0:
            for object in self.main_objects:
                object.config(state=NORMAL)
            for object in self.temperature_objects:
                object.config(state=NORMAL)
            for object in self.cancel_objects:
                object.config(state=NORMAL)
            for object in self.time_objects:
                object.config(state=NORMAL)
        elif array == 1:
            for object in self.main_objects:
                object.config(state=NORMAL)
        elif array == 2:
            for object in self.temperature_objects:
                object.config(state=NORMAL)
        elif array == 3:
            for object in self.cancel_objects:
                object.config(state=NORMAL)
        elif array == 4:
            for object in self.time_objects:
                object.config(state=NORMAL)

        if temp:
            if self.temperature_control_var.get() == 1:
                self.toggle_temperatures_file()
            if self.temperature_control_var.get() == 0:
                self.toggle_temperature_type()
        else:
            self.toggle_temperatures_file()

    def disable(self, array=0):
        '''Torna todos os objetos impossíveis de alteração.'''

        if array == 0:
            for object in self.main_objects:
                object.config(state=DISABLED)
            for object in self.temperature_objects:
                object.config(state=DISABLED)
            for object in self.cancel_objects:
                object.config(state=DISABLED)
            for object in self.time_objects:
                object.config(state=DISABLED)
        elif array == 1:
            for object in self.main_objects:
                object.config(state=DISABLED)
        elif array == 2:
            for object in self.temperature_objects:
                object.config(state=DISABLED)
        elif array == 3:
            for object in self.cancel_objects:
                object.config(state=DISABLED)
        elif array == 4:
            for object in self.time_objects:
                object.config(state=DISABLED)

    def clean(self, array=0):
        '''Elimina todos os objetos da tela anterior.'''

        if array == 0:
            self.frame_ajustable.config(relief=FLAT)
            # Main
            to_remove = []
            for object in self.main_objects:
                to_remove.append(object)
                object.destroy()
            for object in to_remove:
                self.main_objects.remove(object)

            # Temperature
            to_remove = []
            for object in self.temperature_objects:
                to_remove.append(object)
                object.destroy()
            for object in to_remove:
                self.temperature_objects.remove(object)

            # Cancel
            to_remove = []
            for object in self.cancel_objects:
                to_remove.append(object)
                object.destroy()
            for object in to_remove:
                self.cancel_objects.remove(object)

            # Time
            to_remove = []
            for object in self.time_objects:
                to_remove.append(object)
                object.destroy()
            for object in to_remove:
                self.time_objects.remove(object)

        elif array == 1:
            # Main
            to_remove = []
            for object in self.main_objects:
                to_remove.append(object)
                object.destroy()
            for object in to_remove:
                self.main_objects.remove(object)

        elif array == 2:

            # Temperature
            to_remove = []
            for object in self.temperature_objects:
                to_remove.append(object)
                object.destroy()
            for object in to_remove:
                self.temperature_objects.remove(object)

        elif array == 3:
            # Cancel
            to_remove = []
            for object in self.cancel_objects:
                to_remove.append(object)
                object.destroy()
            for object in to_remove:
                self.cancel_objects.remove(object)

        elif array == 4:
            # Time
            to_remove = []
            for object in self.time_objects:
                to_remove.append(object)
                object.destroy()
            for object in to_remove:
                self.time_objects.remove(object)

    ### Toggles

    def toggle_temperature_type(self):

        if self.room_temperature_control_var.get() == 1:
            self.temperature_entry.config(state=DISABLED)
            self.temperature_entry2.config(state=DISABLED)
            self.temperature_entry3.config(state=DISABLED)
        else:
            self.temperature_entry.config(state=NORMAL)
            self.temperature_entry2.config(state=NORMAL)
            self.temperature_entry3.config(state=NORMAL)

    def toggle_temperatures_file(self):

        if self.file_temperatures_control_var.get() == 1:
            self.init_temperature_entry.config(state=DISABLED)
            self.init_temperature_entry2.config(state=DISABLED)
            self.init_temperature_entry3.config(state=DISABLED)
            self.step_temperature_entry.config(state=DISABLED)
            self.end_temperature_entry.config(state=DISABLED)
            self.end_temperature_entry2.config(state=DISABLED)
            self.end_temperature_entry3.config(state=DISABLED)
            self.step_temperature_entry2.config(state=DISABLED)
            self.file_temperatures_entry.config(state=NORMAL)
            self.file_temperatures_button.config(state=NORMAL)
            self.synchro_multiple("", "")
        else:
            self.init_temperature_entry.config(state=NORMAL)
            self.init_temperature_entry2.config(state=NORMAL)
            self.init_temperature_entry3.config(state=NORMAL)
            self.step_temperature_entry.config(state=NORMAL)
            self.end_temperature_entry.config(state=NORMAL)
            self.end_temperature_entry2.config(state=NORMAL)
            self.end_temperature_entry3.config(state=NORMAL)
            self.step_temperature_entry2.config(state=NORMAL)
            self.file_temperatures_entry.config(state=DISABLED)
            self.file_temperatures_button.config(state=DISABLED)
            self.synchro_multiple("", "")


    # Not Avaliable

    def not_avaliable(self):
        '''Auxiliar para funções ainda não criadas.'''

        tkMessageBox.showinfo("Aviso", "Função ainda não disponível.")

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
        pythoncom.CoInitialize()

        write_log("Experimento iniciado")

        self.executing = True

        self.control.set_parameters(data6, data7)

        if self.start(data0, data1, data8):
            first = True
            for temperature in data2:
                if not self.run(temperature, data3, data4, data5, data8[0], data9, init=first):
                    break
                main_window.mpb["value"] += 1
                first = False
            else:
                self.end(data8[0], data9)

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
        if not self.control.StartBVT(gasflow, low_temperature[0], tune=to_tune):
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
        else:
            bvt_temperature = self.control.GetTemperature()
            amostra_temperature = self.serialcom.GetTemperature()
            amostra = main_window.calib_entry.get()
            to_save_path = main_window.path_entry.get()
            info_file = open(to_save_path + "\\" + amostra + " calibracao curva.txt", "a")
            info_file.write(str(bvt_temperature) + "," + str(amostra_temperature) + "\n")
            info_file.close()

        write_log("Término de execução em: " + str(temperature) + " K com sucesso")
        return True

    def end(self, low_temperature, calib):

        if calib != 0:
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

            result = "no"

            if result == 'yes':
                filename = "Temp_params.txt" if not low_temperature else "Temp_params_low.txt"
                file = open(filename, "w")
                file.write(str(a))
                file.write("\n")
                file.write(str(b))
                file.close()


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
    global main_window
    pythoncom.CoInitialize()
    main_window = GUI()
    main_window.master.mainloop()

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

def manage_serial():

    while gui_running:
        if not experiment.serialcom.isconnecting:
            experiment.serialcom.ReadTemperature()

def main():

    global gui_data, experiment_data, gui_running

    # Create Queues
    gui_data = Queue.Queue()
    experiment_data = Queue.Queue()

    # Create Threads
    gui_running = True
    gui = Thread(target=manage_window, args=())
    experiment = Thread(target=manage_experiment, args=())
    serial = Thread(target=manage_serial, args=())

    # Start Threads
    gui.start()
    experiment.start()
    sleep(1)
    serial.start()

    # Wait Threads to finish
    gui.join()
    gui_running = False
    serial.join()
    experiment.join()

main()