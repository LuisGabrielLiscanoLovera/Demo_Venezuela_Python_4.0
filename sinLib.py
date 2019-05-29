from datetime import (timedelta, datetime as pyDateTime, date as pyDate, time as pyTime)
import sys
import Tfhka
import serial
import os
import time
printer = Tfhka.Tfhka()
def abrir_puerto(port):
    global printer
    try:
        resp = printer.OpenFpctrl(port)        
    except serial.SerialException:print "Impresora no Conectada o Error Accediendo al Puerto"
def estado_error():
    global printer
    std = printer.ReadFpStatus()
    status=std[5:8]
    
    
    return status,std[17:]
def obtener_reporteZ():
    global printer
    reporte = printer.GetZReport()
    rZ= str(reporte._numberOfLastZReport).replace(".","").zfill(4)             #1 L4
    rZ+=str(reporte._zReportDate).replace("-","").replace("20","").zfill(6)    #2 L6
    rZ+=str(reporte._numberOfLastInvoice).replace(".","").zfill(8)             #3 L8
    rZ+=str(reporte._lastInvoiceDate).replace("20","").replace("-","").zfill(6)#4 L6
    rZ+=str(reporte._lastInvoiceTime).replace(":","").zfill(4)                 #5 L4
    rZ+=str(reporte._freeSalesTax).replace(".","").zfill(10)                   #6 L10
    rZ+=str(reporte._generalRate1Sale).replace(".","").zfill(10)               #7 L10
    rZ+=str(reporte._generalRate1Tax).replace(".","").zfill(10)                #8 L10
    rZ+=str(reporte._reducedRate2Sale).replace(".","").zfill(10)               #9 L10
    rZ+=str(reporte._reducedRate2Tax).replace(".","").zfill(10)                #10 L0
    rZ+=str(reporte._additionalRate3Sal).replace(".","").zfill(10)             #11 L10
    rZ+=str(reporte._additionalRate3Tax).replace(".","").zfill(10)             #12 L10
    rZ+=str(reporte._freeTaxDebit).replace(".","").zfill(10)                   #13 L10
    rZ+=str(reporte._generalRateDebit).replace(".","").zfill(10)               #14 L10
    rZ+=str(reporte._generalRateTaxDebit).replace(".","").zfill(10)            #15 L10
    rZ+=str(reporte._reducedRateDebit).replace(".","").zfill(10)               #16 L10
    rZ+=str(reporte._reducedRateTaxDebit).replace(".","").zfill(10)            #17 L10
    rZ+=str(reporte._additionalRateDebit).replace(".","").zfill(10)            #18 L10
    rZ+=str(reporte._additionalRateTaxDebit).replace(".","").zfill(10)         #19 L10
    rZ+=str(reporte._freeTaxDevolution).replace(".","").zfill(10)              #20 L10
   
    return str(rZ)
def obtener_reporteS1():
    global printer
    estado_s1 = printer.GetS1PrinterData()
    rS= "01"
    rS+= str(estado_s1._cashierNumber).replace(".","").zfill(2)                             #2 L2
    rS+= str(estado_s1._totalDailySales).replace(".","").zfill(17)                          #3 L17
    rS+= str(estado_s1._lastInvoiceNumber).replace(".","").zfill(8)                         #4 L8
    rS+= str(estado_s1._quantityOfInvoicesToday).replace(".","").zfill(5)                   #5 L5
    rS+= str(estado_s1._quantityNonFiscalDocuments).replace(".","").zfill(8)                #6 L8
    rS+= str(estado_s1._dailyClosureCounter).replace(".","").zfill(5)                       #7 L5
    rS+= str(estado_s1._fiscalReportsCounter).replace(".","").zfill(4)                      #8 L4
    rS+= str(estado_s1._rif).replace(".","").replace("?","").zfill(11)                      #9 L11
    rS+= str(estado_s1._registeredMachineNumber).replace("?","").replace(".","").zfill(10)  #10 L10
    rS+= str(estado_s1._currentPrinterTime).replace(":","").zfill(6)                        #11 L4
    rS+= str(estado_s1._currentPrinterDate).replace("-","").replace("20","").zfill(8)       #12 L6
    rS+= str(estado_s1._quantityOfNCToday).replace(".","").zfill(7)                         #13 L7
    rS+= str(estado_s1._numberNonFiscalDocuments).replace(".","").zfill(10)                 #14 L10
    return str(rS)
def cerrar_puerto():
    global printer
    resp = printer.CloseFpctrl()
    if not resp:pass
    else:pass

abrir_puerto("COM18")
#print estado_error()[1]
if estado_error()[0] != "128":
    reporteS1  =  obtener_reporteS1()
    reporteZ   =  obtener_reporteZ()
    print "\nS1 :"+reporteS1+"lenght"+len(reporteS1)
    print "\nZ  :" +reporteZ  +"lenght"+len(reporteS1)
    cerrar_puerto()
else:print "error de impresora"
