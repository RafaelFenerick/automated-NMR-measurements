import serial
import serial.tools.list_ports
from time import sleep

class SerialCommunication():

    def __init__(self, port):

        arduino_ports = [
            p.device
            for p in serial.tools.list_ports.comports()
            if 'Arduino' in p.description
        ]
        if not arduino_ports:
            print("No Arduino found")
        if len(arduino_ports) > 1:
            print('Multiple Arduinos found - using the first')

        if arduino_ports:
            port = arduino_ports[0]

        self.temperature = 0.0
        self.loop = True
        self.isconnecting = False

        self.connected = True
        try:
            self.port = serial.Serial(port=port, baudrate=9600, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE, bytesize=serial.EIGHTBITS, timeout=0)
        except:
            self.connected = False

    def __del__(self):
        if self.connected:
            self.port.close()

    def Connect(self, port):

        self.isconnecting = True
        sleep(1)

        arduino_ports = [
            p.device
            for p in serial.tools.list_ports.comports()
            if 'Arduino' in p.description
        ]
        #if not arduino_ports:
        #    print("No Arduino found")
        #if len(arduino_ports) > 1:
        #    print('Multiple Arduinos found - using the first')

        if arduino_ports:
            port = arduino_ports[0]

        if self.connected:
            try:
                self.port.close()
            except:
                pass
        self.connected = True
        try:
            self.port = serial.Serial(port=port, baudrate=9600, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                                      bytesize=serial.EIGHTBITS, timeout=0)
        except:
            self.connected = False

        self.isconnecting = False

    def ReadTemperature(self):

        if not self.connected:
            return None

        while True:
            data = self.port.read()
            if data == '\n':
                break

        line = ""
        while True:
            data = self.port.read()
            if data == '\n':
                break
            else:
                line += data

        if line == "":
            return None

        self.temperature = self.resistence2Temperature(float(line))
        return self.resistence2Temperature(float(line))

    def GetTemperature(self):
        return self.temperature

    def resistence2Temperature(self, resistence):

        convertdict = {-55:607800,-54:569604,-53:534036,-52:500901,-51:470019,-50:441224,-49:414363,-48:389296,
                       -47:365893,-46:344034,-45:323609,-44:304517,-43:286662,-42:269958,-41:254324,-40:239686,
                       -39:225976,-38:213129,-37:201087,-36:189794,-35:179200,-34:169258,-33:159925,-32:151159,
                       -31:142924,-30:135185,-29:127908,-28:121065,-27:114626,-26:108566,-25:102861,-24:97487,
                       -23:92425,-22:87653,-21:83155,-20:78913,-19:74910,-18:71133,-17:67568,-16:64201,-15:61020,
                       -14:58014,-13:55174,-12:52487,-11:49947,-10:47543,-9:45268,-8:43115,-7:41076,-6:39144,
                       -5:37313,-4:35578,-3:33934,-2:32374,-1:30894,0:29490,1:28155,2:26888,3:25685,4:24543,
                       5:23457,6:22425,7:21444,8:20512,9:19624,10:18780,11:17977,12:17212,13:16484,14:15791,
                       15:15130,16:14501,17:13901,18:13329,19:12783,20:12263,21:11766,22:11293,23:10840,
                       24:10409,25:10000,26:9602,27:9226,28:8866,29:8522,30:8194,31:7879,32:7579,33:7291,
                       34:7016,35:6752,36:6500,37:6258,38:6027,39:5805,40:5592,41:5389,42:5194,43:5007,44:4827,
                       45:4655,46:4490,47:4331,48:4179,49:4033,50:3893,51:3758,52:3629,53:3505,54:3385,55:3271,
                       56:3160,57:3054,58:2952,59:2854,60:2760,61:2669,62:2582,63:2498,64:2417,65:2339,66:2264,
                       67:2191,68:2122,69:2055,70:1990,71:1928,72:1868,73:1810,74:1754,75:1700,76:1648,77:1598,
                       78:1550,79:1503,80:1458,81:1414,82:1372,83:1332,84:1293,85:1255,86:1218,87:1183,88:1149,
                       89:1116,90:1084,91:1053,92:1023,93:994.2,94:966.3,95:939.3,96:913.2,97:887.9,98:863.4,
                       99:839.7,100:816.8,101:794.6,102:773.1,103:752.3,104:732.1,105:712.6,106:693.6,107:675.3,
                       108:657.5,109:640.3,110:623.6,111:607.4,112:591.6,113:576.4,114:561.6,115:547.3,116:533.4,
                       117:519.9,118:506.8,119:494.1,120:481.8,121:469.8,122:458.2,123:446.9,124:435.9,125:425.3,
                       126:414.9,127:404.9,128:395.1,129:385.6,130:376.4,131:367.4,132:358.7,133:350.3,134:342.0,
                       135:334.0,136:326.3,137:318.7,138:311.3,139:304.2,140:297.2,141:290.4,142:283.8,143:277.4,
                       144:271.2,145:265.1,146:259.2,147:253.4,148:247.8,149:242.3,150:237.0}

        temps = []
        resistences = []
        for key in convertdict:
            temps.append(key)
        temps = list(sorted(temps))
        for key in temps:
            resistences.append(convertdict[key])

        i = 0
        temperature = 1000
        while resistence < resistences[i]:
            i += 1
            if i == len(resistences):
                break
        else:
            temperature = temps[i]

        return temperature


if __name__ == "__main__":

    from time import sleep
    serial1 = SerialCommunication("COM2")
    while True:
        #sleep(1)
        serial1.Connect("COM2")
        #print("Connected: ", serial1.connected)
        print("Temperature: ", serial1.ReadTemperature())
