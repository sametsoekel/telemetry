"""

Oysa herkes öldürür sevdiğini
Kulak verin bu dediklerime
Kimi bir bakışıyla yapar bunu,
Kimi dalkavukça sözlerle,
Korkaklar öpücük ile öldürür,
Yürekliler kılıç darbeleriyle!

"""


import sys
from PyQt5.QtCore import pyqtSlot,QUrl,QTimer,QTime,QThread,pyqtSignal
from PyQt5.QtWidgets import QApplication,QDialog,QMessageBox
from PyQt5.uic import loadUi
from PyQt5 import QtWidgets
from PyQt5 import QtGui
import feedparser
import xlsxwriter
import time
import serial
import datetime

##############################hava durumu çekme##############################
"""def havadurumu():
    parse = feedparser.parse("http://rss.accuweather.com/rss/liveweather_rss.asp?metric=1&locCode=EUR|TR|26555|ESKISEHIR|")
    parse = parse["entries"][0]["summary"]
    parse = parse.split()
    return (parse[2]+parse[4]+parse[5])"""
##############################hava durumu çekme##############################


#########################################################################################################################################################################


##############################split fonksiyonları başlangıç##############################

#verideki ilk 3 satırı split etme fonksiyonu
def firstsplitter(veri,how):
    b=veri.split("*")
    datas=[]
    for i in range(0,3):
        for j in range(0,15):
            datas.append(str(str(str(str(b[i].split("[")).split("]")).split("'")[3]).split('"')[0]).split(",")[j])
        datas.append(str(str(str(b[i].split("[")).split("]")).split("'")[3]).split('"')[2])#yıldızlı veri
    return (datas[how])
        
        
#4,5 ve 6. satırları split etme fonksiyonu
def secondsplitter(veri,how):
    b=veri.split("*")
    datas=[]
    for i in range(0,3):
        for j in range(0,5):
            datas.append(str(str(str(str(str(b[i]).split("[")).split("]")).split("'")[3]).split('"')[0]).split(",")[j])
        datas.append(str(str(str(b[i].split("[")).split("]")).split("'")[3]).split('"')[2])#yıldızlı veri
    return (datas[how]) 

#7. satırı split etme fonksiyonu      
def thirdsplitter(veri,how):
    b=veri.split("*")
    datas=[]
    for j in range(0,7):
        datas.append(str(str(str(str(str(b[0]).split("[")).split("]")).split("'")[3]).split('"')[0]).split(",")[j])
    datas.append(str(str(str(b[0].split("[")).split("]")).split("'")[3]).split('"')[2])#yıldızlı veri
    return(datas[how])
    
#8. satırı split etme fonksiyonu
def fourthsplitter(veri,how):
    b=veri.split("*")
    datas=[]
    for j in range(0,6):
        datas.append(str(str(str(str(str(b[0]).split("[")).split("]")).split("'")[3]).split('"')[0]).split(",")[j])
    datas.append(str(str(str(b[0].split("[")).split("]")).split("'")[3]).split('"')[2])#yıldızlı veri
    return (datas[how])
    
workbook = xlsxwriter.Workbook('C:/Users/Public/LOG.xlsx')
worksheet = workbook.add_worksheet()     
    
def datalogger(i,gerilim,makim,gakim,bakim,durum,hiz,gyol,mwh,mah,gwh,gah,bwh,bah,zmn):
    if i==2:
        worksheet.write('A1', 'Gerilim') 
        worksheet.write('B1', 'Motor Akimi') 
        worksheet.write('C1', 'Gunes Akimi') 
        worksheet.write('D1', 'Batarya Akimi')
        worksheet.write('E1', 'Durum') 
        worksheet.write('F1', 'Hiz') 
        worksheet.write('G1', 'Gidilen Yol')
        worksheet.write('H1', 'Motor Wh')
        worksheet.write('I1', 'Motor Ah') 
        worksheet.write('J1', 'Gunes Wh')
        worksheet.write('K1', 'Gunes Ah')
        worksheet.write('L1', 'Batarya Wh')
        worksheet.write('M1', 'Batarya Ah')
        worksheet.write('N1', 'Saat')
        
    else:
        worksheet.write('A'+str(i), str(gerilim)) 
        worksheet.write('B'+str(i), str(makim))
        worksheet.write('C'+str(i), str(gakim))
        worksheet.write('D'+str(i), str(bakim)) 
        worksheet.write('E'+str(i), str(durum))
        worksheet.write('F'+str(i), str(hiz))
        worksheet.write('G'+str(i), str(gyol)) 
        worksheet.write('H'+str(i), str(mwh))
        worksheet.write('I'+str(i), str(mah))
        worksheet.write('J'+str(i), str(gwh)) 
        worksheet.write('K'+str(i), str(gah))
        worksheet.write('L'+str(i), str(bwh))
        worksheet.write('M'+str(i), str(bah)) 
        worksheet.write('N'+str(i), str(zmn))       
        
def stoplog():
    workbook.close()


        
     
      

    
     
    

##############################split fonksiyonları bitiş##############################


#########################################################################################################################################################################


def serial_ports():
    
    if sys.platform.startswith('win'):
        ports = ['COM%s' % (i + 1) for i in range(256)]
    elif sys.platform.startswith('linux') or sys.platform.startswith('cygwin'):
        ports = glob.glob('/dev/tty[A-Za-z]*')
    elif sys.platform.startswith('darwin'):
        ports = glob.glob('/dev/tty.*')
    else:
        raise EnvironmentError('Unsupported platform')

    result = []
    for port in ports:
        try:
            s = serial.Serial(port)
            s.close()
            result.append(port)
        except (OSError, serial.SerialException):
            pass
    return result

class Pencere(QDialog):
    
    def __init__(self):
        
        super(Pencere,self).__init__()
        loadUi("C:\\Users\\Public\\thewindow.ui",self)
        self.baglanbuton.clicked.connect(self.baglan)
        #self.havadurum.setText(str(havadurumu()))#hava durumu deaktif
        self.i=2
        self.run.clicked.connect(self.isleme)
        self.kesbuton.clicked.connect(self.kes)
        self.workbook = xlsxwriter.Workbook('C:/Users/Public/LOG.xlsx')
        self.worksheet = self.workbook.add_worksheet() 
        self.stoplogger.clicked.connect(stoplog)
        self.porttextbox.addItems(serial_ports())

        ###################
       
        
        ############
    def baglan(self):
        try:
            self.mySerial=serialThreadClass(None,str(self.porttextbox.currentText()),self.baudtextbox.text())
            self.mySerial.start()
            self.baglantidurum.setText("Baglanti kuruldu")       
            self.baglanbuton.setEnabled(False)
            self.kesbuton.setEnabled(True)
            self.mySerial.kutum.textChanged.connect(self.verial)
                
        except serial.serialutil.SerialException:
            self.baglantidurum.setText("Baglanti kurulamadi")
            
    def kes(self):
        try:
            self.mySerial.cls()
            self.baglantidurum.setText("Baglanti kesildi")  
            self.baglanbuton.setEnabled(True)
            self.kesbuton.setEnabled(False)
        except serial.serialutil.SerialException:
            self.baglantidurum.setText("Baglanti kesilemedi")
        except AttributeError:
            pass
        
    def verial(self):
        try:
            self.slaves=[self.slave1,self.slave2,self.slave3,self.slave4]
            if self.veritextbox.toPlainText!="":
                self.veritextbox.setPlainText("")
            if self.slaves[0].toPlainText()!="":
                self.slaves[0].setPlainText("")
            if self.slaves[1].toPlainText()!="":
                self.slaves[1].setPlainText("")
            if self.slaves[2].toPlainText()!="":
                self.slaves[2].setPlainText("")
            if self.slaves[3].toPlainText()!="":
                self.slaves[3].setPlainText("")
            self.veritextbox.setPlainText(str(self.mySerial.kutum.toPlainText()))
            c=int(self.veritextbox.toPlainText().count('*'))
            d=self.veritextbox.toPlainText().split('*')
            for i in range (0,c):
                if len(d[i])==76:
                    self.slaves[0].setPlainText(self.slaves[0].toPlainText()+d[i]+"*")
                elif len(d[i])==24:
                    self.slaves[1].setPlainText(self.slaves[1].toPlainText()+d[i]+"*")
                elif len(d[i])==33:
                    self.slaves[2].setPlainText(self.slaves[2].toPlainText()+d[i]+"*")
                elif len(d[i])==34:
                    self.slaves[3].setPlainText(self.slaves[3].toPlainText()+d[i]+"*")
            self.veritextboxshow.setPlainText(str(self.veritextboxshow.toPlainText())+str(self.veritextbox.toPlainText()))
            self.isleme()
            self.serialThreadClass.kutum.setPlainText("")
            self.veritextbox.setPlainText("")
        except AttributeError:
            pass
    
        except ValueError:
            pass
    def isleme(self):
        try:
            #ssaaaatttt#
            self.an=datetime.datetime.now()
            self.zaman = datetime.datetime.strftime(self.an,"%X")
            #ssaaaatttt#
            
            bmslist=[]
  
            
            bmsveri=str(self.slaves[0].toPlainText())
            sicaklikveri=str(self.slaves[1].toPlainText())
            hizvoltamperveri=str(self.slaves[2].toPlainText())
            ####volt ve amper verisini okuma#######
            volt=float(thirdsplitter(hizvoltamperveri,4))/10
            self.batteryvolt.display(volt)
            amper=float(thirdsplitter(hizvoltamperveri,5))/100
            self.batteryamper.display(amper)
            self.batterybar.setValue(volt*10)
        except IndexError:
            pass 
        except ValueError:
            pass
        try:
            
            #####bms verilerini okuma#####
            self.bms102.display(firstsplitter(bmsveri,3))
            self.bms103.display(firstsplitter(bmsveri,4))
            self.bms104.display(firstsplitter(bmsveri,5))
            self.bms105.display(firstsplitter(bmsveri,6))
            self.bms106.display(firstsplitter(bmsveri,7))
            self.bms107.display(firstsplitter(bmsveri,8))
            self.bms108.display(firstsplitter(bmsveri,9))
            self.bms109.display(firstsplitter(bmsveri,10))
            self.bms1010.display(firstsplitter(bmsveri,11))
            self.bms1011.display(firstsplitter(bmsveri,12))
            self.bms1012.display(firstsplitter(bmsveri,13))
            
            self.bms202.display(firstsplitter(bmsveri,19))
            self.bms203.display(firstsplitter(bmsveri,20))
            self.bms204.display(firstsplitter(bmsveri,21))
            self.bms205.display(firstsplitter(bmsveri,22))
            self.bms206.display(firstsplitter(bmsveri,23))
            self.bms207.display(firstsplitter(bmsveri,24))
            self.bms208.display(firstsplitter(bmsveri,25))
            self.bms209.display(firstsplitter(bmsveri,26))
            self.bms2010.display(firstsplitter(bmsveri,27))
            self.bms2011.display(firstsplitter(bmsveri,28))
            self.bms2012.display(firstsplitter(bmsveri,29))
            
            self.bms302.display(firstsplitter(bmsveri,35))
            self.bms303.display(firstsplitter(bmsveri,36))
            self.bms304.display(firstsplitter(bmsveri,37))
            self.bms305.display(firstsplitter(bmsveri,38))
            self.bms306.display(firstsplitter(bmsveri,39))
            self.bms307.display(firstsplitter(bmsveri,40))
            self.bms308.display(firstsplitter(bmsveri,41))
            self.bms309.display(firstsplitter(bmsveri,42))
            self.bms3010.display(firstsplitter(bmsveri,43))
            self.bms3011.display(firstsplitter(bmsveri,44))
            self.bms3012.display(firstsplitter(bmsveri,45))
            
            
            for i in range (3,14):
                bmslist.append(firstsplitter(bmsveri,i))
            for i in range (19,30):
                bmslist.append(firstsplitter(bmsveri,i))
            for i in range (35,46):
                bmslist.append(firstsplitter(bmsveri,i))
                
            self.bmsmin.display(min(bmslist))
            self.bmsmax.display(max(bmslist))
            self.bmsfark.setText(str(int(max(bmslist))-int(min(bmslist))))
        except IndexError:
            pass
        except ValueError:
            pass
        
        try:
            
            self.sensor1.display(float(secondsplitter(sicaklikveri,4))/10)
            self.sensor2.display(float(secondsplitter(sicaklikveri,9))/10)
            self.sensor3.display(float(secondsplitter(sicaklikveri,10))/10)
            self.sensor4.display(float(secondsplitter(sicaklikveri,15))/10)
            hiz=thirdsplitter(hizvoltamperveri,3)
            self.hizgostergedigital.display(hiz)
            self.hizgosterge.setValue(self.hizgostergedigital.value())
        except IndexError:
            pass
        except ValueError:
            pass
            
            #datalogger(self.i,volt,amper,0,0,0,hiz,0,0,0,0,0,0,0,self.zaman)
            self.i+=1
            
            

#---------------------------------------------------------------------------------------------------------
########seri bağlantı başlangıç##############################
class serialThreadClass(QThread):
    
    mySignal=pyqtSignal
    def __init__(self,parent=None,comport=0,baud=0):
        super(serialThreadClass,self).__init__(parent)
        self.seriport=serial.Serial(
        port=comport,\
        baudrate=int(baud),\
        parity=serial.PARITY_NONE,\
        stopbits=serial.STOPBITS_ONE,\
        bytesize=8,\
        timeout=0)
        self.veriler=""
        self.kutum=QtWidgets.QPlainTextEdit()
        
    def run(self):
        try:
            while True:
                self.veri=self.seriport.readline()
                self.mesaj=str(self.veri)
                if len(self.kutum.toPlainText())>=375:
                    self.veriler=""
                    if str(self.veri)!="b''":
                        self.veriler+=str(self.veri)
                        a=str(self.veriler).replace("r","")
                        b=a.replace("'","")
                        c=b.replace("\\","")
                        d=c.replace("b","")
                        self.kutum.setPlainText(d)
                           
                        
                        
                else:
                    if str(self.veri)!="b''":
                        self.veriler+=str(self.veri)
                        a=str(self.veriler).replace("r","")
                        b=a.replace("'","")
                        c=b.replace("\\","")
                        d=c.replace("b","")
                        self.kutum.setPlainText(d)
        except AttributeError:
            pass
    
    def cls(self):
        self.seriport.close()

                
##############################seri bağlantı bitiş##############################
        
if __name__ == '__main__':
    app=QApplication(sys.argv)
    widget=Pencere()
    widget.show()
    app.exit(app.exec())
