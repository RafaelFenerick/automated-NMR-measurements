import win32com.client as vbs
import tkMessageBox
import traceback
from time import sleep

class Control():

    def __init__(self):
        self.serialnumber = ""
        self.exe_path = ""
        self.ErrNo = ""
        self.ErrMsg = ""
        self.stop = False

        file = open("output_control_errors.txt", "w")
        file.close()

    def __del__(self):
        pass

    def set_parameters(self, serialnumber, exe_path):
        self.serialnumber = serialnumber
        self.exe_path = exe_path

    def ConnectBVT(self):
        '''Connect BVT and check it's state (On/Off)'''

        if self.stop:
            self.stop = False
            return False

        file = open("output_control_errors.txt", "a")

        try:
            uti = vbs.Dispatch("WinAcquisit.Utilities")
            emb = vbs.Dispatch("WinAcquisit.Embedding")
            emb.ShowWindow(emb.NORMAL)
            self.bvt = vbs.Dispatch("WinAcquisit.BVT")
        except:
            file.write("Error - BVT not conneted")
            file.write("\n")
            file.close()
            return False

        # Check BVT state
        bOn = self.bvt.IsBVTOn
        if self.bvt.IsLastError:
            self.bvt.GetLastError
            file.write("Error - IsBVTOn -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            tkMessageBox.showinfo("Aviso", traceback.format_exc())
            return False

        if not bOn:
            file.write("Error - IsBVTOn - BVT off")
            file.write("\n")
            file.close()
            return False

        file.close()

        return True

    def ConnectPNMR(self):
        '''Connect PNMR and Minispec given SerialNumber'''

        if self.stop:
            self.stop = False
            return False

        file = open("output_control_errors.txt", "a")

        try:
            self.pnmr = vbs.Dispatch("theMinispecMqF.PNMR")
        except:
            file.write("Error - PNMR not connected" )
            file.write("\n")
            file.close()
            return False

        # Config WakeUp Behavior
        self.pnmr.ConfigWakeUp([0, self.pnmr.MAXIMIZED, 1, 0, 0, 0])
        if self.pnmr.IsLastError:
            self.ErrMsg = self.pnmr.GetLastError
            file.write("Error - ConfigWakeUp -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False

        self.pnmr.OpenPNMR()
        if self.pnmr.IsLastError:
            self.ErrMsg = self.pnmr.GetLastError
            file.write("Error - OpenPNMR -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False

        #--- Connect Minispec
        csCurrSerNo = self.pnmr.GetInstrumentSerialNumber
        if self.pnmr.IsLastError:
            self.ErrMsg = self.pnmr.GetLastError
            file.write("Error - GetInstrumentSerialNumber -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False

        # Check current Minispec SerialNumber connected
        # Connect if not same SerialNumber
        if csCurrSerNo != self.serialnumber:
            bConnectState = self.pnmr.ConnectInstrument(self.serialnumber)
            if self.pnmr.IsLastError:
                self.ErrMsg = self.pnmr.GetLastError
                file.write("Error - ConnectInstrument -" + str(self.ErrNo) + str(self.ErrMsg))
                file.write("\n")
                file.close()
                return False
            if not bConnectState:
                file.write("Error - ConnectInstrument - Minispec could not connect")
                file.write("\n")
                file.close()
                return False

        file.close()

        return True

    def StartBVT(self, gasflow, low_temperature, tune=False):
        '''Set Gasflow, turn gasflow and heater and PID tune'''

        if self.stop:
            self.stop = False
            return False

        file = open("output_control_errors.txt", "a")

        # Set GasFlow
        self.bvt.GasFlow(gasflow)
        if self.bvt.IsLastError:
            self.bvt.GetLastError
            file.write("Error - SetGasFlow -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False

        # Turn GasFlow On
        self.bvt.GasFlowOn(True)
        if self.bvt.IsLastError:
            self.bvt.GetLastError
            file.write("Error - GasFlowOn -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False

        if low_temperature:
            # Turn Evaporator On
            self.bvt.EvaporatorOn(True)
            if self.bvt.IsLastError:
                self.bvt.GetLastError
                file.write("Error - EvaporatorOn -" + str(self.ErrNo) + str(self.ErrMsg))
                file.write("\n")
                file.close()
                return False

            # Set Evaporator Power
            self.bvt.EvaporatorPower(gasflow)
            if self.bvt.IsLastError:
                self.bvt.GetLastError
                file.write("Error - EvaporatorPower -" + str(self.ErrNo) + str(self.ErrMsg))
                file.write("\n")
                file.close()
                return False

        # Turn Heater On
        self.bvt.HeaterOn(True)
        if self.bvt.IsLastError:
            self.bvt.GetLastError
            file.write("Error - HeaterOn -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False

        # Check if PID Tune On
        while self.bvt.IsPIDTuneOn:

            if self.stop:
                self.stop = False
                return False

            if self.bvt.IsLastError:
                self.bvt.GetLastError
                file.write("Error - PIDTuneOn -" + str(self.ErrNo) + str(self.ErrMsg))
                file.write("\n")
                file.close()
                return False

        if self.bvt.IsLastError:
            self.bvt.GetLastError
            file.write("Error - PIDTuneOn -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False

        # TunePID
        if tune:
            self.bvt.PIDTuneOn(True)
            if self.bvt.IsLastError:
                self.bvt.GetLastError
                file.write("Error - PIDTune -" + str(self.ErrNo) + str(self.ErrMsg))
                file.write("\n")
                file.close()
                return False

            while self.bvt.IsPIDTuneOn:

                if self.stop:
                    self.stop = False
                    return False

                if self.bvt.IsLastError:
                    self.bvt.GetLastError
                    file.write("Error - WaitPIDTune -" + str(self.ErrNo) + str(self.ErrMsg))
                    file.write("\n")
                    file.close()
                    return False

        file.close()

        return True

    def GetTemperature(self):
        '''Get current BVT temperature.'''

        if self.stop:
            self.stop = False
            return False

        file = open("output_control_errors.txt", "a")

        dTemp = self.bvt.GetTemperature
        if self.bvt.IsLastError:
            self.bvt.GetLastError
            file.write("Error - GetTemperature -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return -1

        file.close()
        self.stop = False
        return dTemp

    def SetTemperature(self, desired_temperature, wait_time):
        '''Set desired temperature. Ramp if difference is over
        determinated threshold.'''

        if self.stop:
            self.stop = False
            return False

        file = open("output_control_errors.txt", "a")

        self.bvt.DesiredTemperature(desired_temperature)
        if self.bvt.IsLastError:
            self.bvt.GetLastError
            file.write("Error - SetTemperature -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False

        temp = self.bvt.GetDesiredTemperature
        if self.bvt.IsLastError:
            self.bvt.GetLastError
            file.write("Error - GetDesiredTemperature -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False

        print("Sleeping")
        i = 0
        while i < wait_time*60:
            sleep(1)

            if self.stop:
                self.stop = False
                return False

            if not self.CheckGasFlow():
                return False

            i += 1
        print("Slept")

        file.close()

        return True
    
    def ExecuteApplication(self, app_file):
        '''Load, run and wait for execution of given application.'''

        file = open("output_control_errors.txt", "a")

        run = False
        time_out = 0
        total_wait_time = 60
        while self.pnmr.IsApplicationRunning:
            sleep(1)

            if self.stop:
                self.stop = False
                return False

            if not self.CheckGasFlow():
                return False

            time_out += 1
            if time_out == total_wait_time:
                break
        else:
            run = True

        if not run:
            file.write("Error - LoadApplication - Application runnig yet.")
            file.write("\n")
            file.close()
            return False

        self.pnmr.LoadApplication(app_file)
        if self.pnmr.IsLastError:
            self.ErrMsg = self.pnmr.GetLastError
            file.write("Error - LoadApplication -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False

        self.pnmr.IsApplicationLoaded
        if self.pnmr.IsLastError:
            self.ErrMsg = self.pnmr.GetLastError
            file.write("Error - IsApplicationLoaded -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False

        if self.stop:
            self.stop = False
            return False

        # Run Application
        self.pnmr.RunApplication()
        if self.pnmr.IsLastError:
            self.pnmr.GetLastError
            file.write("Error - RunningApplication -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False

        run = False
        time_out = 0
        total_wait_time = 10
        while not self.pnmr.IsApplicationRunning:
            sleep(1)
            #--- Stop Application??
            if self.stop:
                self.stop = False
                return False

            if not self.CheckGasFlow():
                return False

            time_out += 1
            if time_out == total_wait_time:
                break
        else:
            run = True

        if not run:
            file.write("Error - WaitApplicationtoStartRunning")
            file.write("\n")
            file.close()
            return False

        # Wait For Aquisition Done

        # Aquisition to start
        run = False
        time_out = 0
        total_wait_time = 60
        scansToDo = 0
        scansDone = 0
        while not self.pnmr.GetDataAcquisitionProgress:
            sleep(1)
            if self.stop:
                self.stop = False
                return False

            if not self.CheckGasFlow():
                return False

            time_out += 1
            if time_out == total_wait_time:
                break
        else:
            run = True

        if not run:
            file.write("Error - GetDataAcquisitionProgress Start")
            file.write("\n")
            file.close()
            return False

        # Aquisition to end
        run = False
        time_out = 0
        total_wait_time = 30*60
        scansToDo = 0
        scansDone = 0
        while self.pnmr.GetDataAcquisitionProgress:
            sleep(1)
            if self.stop:
                self.stop = False
                return False

            if not self.CheckGasFlow():
                return False

            time_out += 1
            if time_out == total_wait_time:
                break
        else:
            run = True

        if not run:
            file.write("Error - GetDataAcquisitionProgress end")
            file.write("\n")
            file.close()
            return False

        run = False
        time_out = 0
        total_wait_time = 30*60
        while self.pnmr.IsApplicationRunning:
            sleep(1)
            if self.stop:
                self.stop = False
                return False

            if not self.CheckGasFlow():
                return False

            time_out += 1
            if time_out == total_wait_time:
                break
        else:
            run = True

        if not run:
            file.write("Error - LoadApplication - Application runnig yet.")
            file.write("\n")
            file.close()
            return False

        file.close()
        self.stop = False
        return True

    def AbortApplication(self):

        file = open("output_control_errors.txt", "a")

        if self.pnmr.IsApplicationRunning:
            self.pnmr.StopApplication
            if self.pnmr.IsLastError:
                self.ErrMsg = self.pnmr.GetLastError
                file.write("Error - StopApplication -" + str(self.ErrNo) + str(self.ErrMsg))
                file.write("\n")
                file.close()
                return False
        if self.pnmr.IsApplicationLoaded:
            self.pnmr.ReleaseApplication
            if self.pnmr.IsLastError:
                self.ErrMsg = self.pnmr.GetLastError
                file.write("Error - ReleaseApplication -" + str(self.ErrNo) + str(self.ErrMsg))
                file.write("\n")
                file.close()
                return False

        file.close()
        return True

    def CheckGasFlow(self):

        gas_flow = self.bvt.GetGasFlow
        if self.bvt.IsLastError:
            self.bvt.GetLastError
            file.write("Error - GetGasFlow -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False
        if gas_flow == 0:
            return False
        gas_flow = self.bvt.IsGasFlowOn
        if self.bvt.IsLastError:
            self.bvt.GetLastError
            file.write("Error - GetGasFlow -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            return False
        if not gas_flow:
            return False

        return True

    def DoRamp(self, desired_temperature, rate, to_sleep=-1):

        if self.stop:
            self.stop = False
            return False

        file = open("output_control_errors.txt", "a")

        RampEnabled = self.bvt.IsRampEnabled
        if self.bvt.IsLastError:
            self.bvt.GetLastError
            file.write("Error - IsRampEnabled -" + str(self.ErrNo) + str(self.ErrMsg))
            file.write("\n")
            file.close()
            tkMessageBox.showinfo("Aviso", traceback.format_exc())
            return False

        if not RampEnabled:
            print("Ramp Not Enable!")

        else:
            while self.bvt.IsRampRunning:

                if self.stop:
                    self.stop = False
                    return False

                if not self.CheckGasFlow():
                    return False

                if self.bvt.IsLastError:
                    self.bvt.GetLastError
                    file.write("Error - IsRampEnabled -" + str(self.ErrNo) + str(self.ErrMsg))
                    file.write("\n")
                    file.close()
                    tkMessageBox.showinfo("Aviso", traceback.format_exc())
                    return False
                print("Ramp Still Running")
                sleep(1)

            # RampGo
            RampOn = self.bvt.RampGo(True, desired_temperature, rate, self.bvt.RAMPHOLD_OFF)
            if self.bvt.IsLastError:
                self.bvt.GetLastError
                file.write("Error - RampGo -" + str(self.ErrNo) + str(self.ErrMsg))
                file.write("\n")
                file.close()
                tkMessageBox.showinfo("Aviso", traceback.format_exc())
                return False

            if not RampOn:
                print("Ramp could not start.")

            else:
                wait_ramp = 0
                while not self.bvt.IsRampRunning:

                    if self.stop:
                        self.stop = False
                        return False

                    if not self.CheckGasFlow():
                        return False

                    if self.bvt.IsLastError:
                        self.bvt.GetLastError
                        file.write("Error - IsRampEnabled -" + str(self.ErrNo) + str(self.ErrMsg))
                        file.write("\n")
                        file.close()
                        tkMessageBox.showinfo("Aviso", traceback.format_exc())
                        return False
                    print("Ramp not Running")
                    wait_ramp += 1
                    if wait_ramp >= 60:
                        break
                    sleep(1)
                else:
                    print ("Ramp started")

                while self.bvt.IsRampRunning:

                    if self.stop:
                        self.stop = False

                        #Cancelar rampa

                        return False

                    if not self.CheckGasFlow():
                        return False

                    if self.bvt.IsLastError:
                        self.bvt.GetLastError
                        file.write("Error - IsRampEnabled -" + str(self.ErrNo) + str(self.ErrMsg))
                        file.write("\n")
                        file.close()
                        tkMessageBox.showinfo("Aviso", traceback.format_exc())
                        return False
                    print("Ramp Running")
                    sleep(1)

        if to_sleep > 0:
            print("Sleeping (pos rampa)")
            i = 0
            while i < to_sleep*60:
                sleep(1)

                if self.stop:
                    self.stop = False
                    return False

                if not self.CheckGasFlow():
                    return False

                i += 1
            print('Slept (pos rampa)')

        file.close()
        self.stop = False
        return True

    def Finish(self, ramp=True, bvt=True, minispec=True, low_temperature=0):
        '''Ramp to set temperature and turn heater off.'''

        file = open("output_control_errors.txt", "a")

        # Ramp
        if ramp:
            if low_temperature:
                self.DoRamp(275, 15)
            else:
                self.DoRamp(310, 15)

        # Turn Heater Off
        if bvt:

            if low_temperature:
                self.bvt.EvaporatorOn(False)
                if self.bvt.IsLastError:
                    self.bvt.GetLastError
                    file.write("Error - EvaporatorOn -" + str(self.ErrNo) + str(self.ErrMsg))
                    file.write("\n")
                    file.close()
                    return

            self.bvt.HeaterOn(False)
            if self.bvt.IsLastError:
                self.bvt.GetLastError
                file.write("Error - HeaterOn -" + str(self.ErrNo) +  str(self.ErrMsg))
                file.write("\n")
                file.close()
                return

        if minispec:
            self.pnmr.ClosePNMR(True)
            if self.pnmr.IsLastError:
                    self.ErrMsg = self.pnmr.GetLastError
                    file.write("Error - ClosePNMR -" + str(self.ErrNo) + str(self.ErrMsg))
                    file.write("\n")
                    file.close()
                    return False

        # Disconnect

        # Releases

        self.stop = False
        file.close()

#bvt = Control("", "")
#bvt.ConnectBVT()
#bvt.bvt.HeaterOn(False)