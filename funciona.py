
import Tfhka
import serial
###Imports
try:
    from gc import collect as libM#liberacion de memoria
    #import urllib.request
    import os.path as path
    import configparser as cp
    from traceback import format_exc as formErro
    from subprocess import call as spc
    from os import system as sys
    from os import remove as rm
    from os import mkdir as mk
    from os.path import dirname
    from os.path import realpath
    from datetime import date, timedelta
    from time import sleep as slp
    from ftplib import FTP_TLS
    from shutil import copy as copy
    import logging
    import logging.handlers
    import getpass
except Exception as e:
    print e
    exit()


try:
    printer = Tfhka.Tfhka()#SDK
    today         = date.today()#fecha
    conf          = 'conf.cfg'
    hostFtp       = 'ftp.cocaisystem.com'
    userFtp       = 'cocaisys'
    passFtp       = 'mE]hX9aW6X32b+'
    Ru0z          = 'Reporte.txt'
    Rs1           = 'Reporte_S1.txt'
    log           = 'log' #carpeta
    uoz           = 'uoz.exe'#ejecutable al llamar
    reIntento     = 30
    USER_NAME  = getpass.getuser()
    bat_pathAD = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' % USER_NAME
    bat_path   = r'C:\IntTFHKA'
    accDirect  = 'AccesoRunmvl.vbs'


    infoERR       = False #muestra Err
    ifExi  = lambda archivo:path.exists(archivo)#si exsiste los archivo dependiente
    spcall = lambda exe:spc(exe, shell=False)#ejecuta subProce
    
    #spcall("""REG ADD \"HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Run\" /V \"Rebootmv\" /t REG_SZ /F /D \"C:\\IntTFHKA\\runmvel.exe\"""")


    #Archivo registro
    if not path.exists(log):mk(log)
    logger    = logging.getLogger('Monitoreo de venta mvel')
    logger.setLevel(logging.DEBUG)
    mesAnio=str(today.month)+''+str(today.year)
    fileLog ='log/log-%s.log'% str(mesAnio)
    handler   = logging.handlers.TimedRotatingFileHandler(filename=fileLog, when="m", interval=1)
    formatter = logging.Formatter(fmt='%(asctime)s - %(name)s - %(levelname)s - %(message)s',datefmt='%d-%m-%y %H:%M:%S')
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    #archivo de configuracion
    configuracion   =   cp.ConfigParser()
    if (ifExi(conf)):
        configuracion.read(conf)
        carpetaCompartida   = configuracion['modoEjecucion']['carpetaCompartida']
        modoEjecucion       = int(configuracion['modoEjecucion']['modo'])
        rutaFtp             = configuracion['PathFTP']['rutaFile']
        portFtp             = int(configuracion['PathFTP']['portFtp'])
        codigo              = configuracion['General']['codigo']
        numpc               = configuracion['General']['numpc']
        portImpre           = configuracion['General']['portImpresora']
        activacion          = int(configuracion['General']['activacion'])
        tiempoInit          = int(configuracion['General']['tiempo_init'])#*60
        tiempoErr           = int(configuracion['General']['tiempo_err'])#*60
        #nombre salida archivoinfoERR       = True
        archivoU0Z          = codigo+"_%s_U0Z.txt"%numpc
        archivoS1           = codigo+"_%s_S1.txt"%numpc



    else:cretFileConf()
except Exception as e:
    if infoERR == True:logger.warning(formErro)
    logger.warning(str(e))
    #exit()#si no consigue archivo de configuracion cierra el programa


def startup(file_path=""):
    global bat_pathAD,USER_NAME,bat_path,accDirect
    if file_path == "":
        file_path = dirname(realpath(__file__))
    with open(bat_path + '\\' + accDirect, "w+") as bat_file:
        vbs=r'''
            ficheroAccesoDirecto = "%s\runmvel.lnk"
            ficheroEjecutar      = "C:\IntTFHKA\runmvel.exe"
            descripcion          = "runmvelLink"
            Set objShell = WScript.CreateObject("WScript.Shell")
            Set objAccesoDirecto = objShell.CreateShortcut(ficheroAccesoDirecto)
            objAccesoDirecto.TargetPath = ficheroEjecutar
            objAccesoDirecto.Description = descripcion
            objAccesoDirecto.Save
            Set objAccesoDirecto = Nothing
            '''%bat_pathAD
        bat_file.write(vbs)
    bat_file.close()
    spcall("cmd /c %s" % accDirect)
    rm(accDirect)
#startup()
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

def conexionFTP():
    print "conexionFTP"
    global reIntento, rutaFtp,tiempoErr,hostFtp,passFtp,userFtp,portFtp,infoERR
    #request = urllib.request.Request('http://'+hostFtp)
    #response = urllib.request.urlopen(request)
    estatusCftp=False
    if response.getcode() == 200:
        try:

            ftp = FTP_TLS()
            ftp.connect(hostFtp,portFtp)
            ftp.login(userFtp, passFtp)
            ftp.cwd(rutaFtp)

            estatusCftp = True

        except Exception as e:
            if infoERR == True:logger.warning(formErro)
            logger.warning('Error al conectar al servidor ftp '+str(e))
    else:
        logger.warning('Error al conectar al servidor Error de internet '+str(e))
        for i in range(reIntento):
            if conexionFTP()['ftp']:main()
            else:logger.warning("reconecion internet intento  "+str(i));slp(tiempoErr)#
    return {'ftp':ftp,'estatusCftp':estatusCftp}


def stFcfha():
    global configuracion,conf
    configuracion.read(conf)
    configuracion.set('FeEjecucion', 'fecha', str(today))
    with open(conf, "w+") as configfile:configuracion.write(configfile);configfile.close()
    return configuracion['FeEjecucion']['fecha']


def activarIntTFHKA():
    try:
        global uoz,Ru0z,portImpre,Rs1
        estado=False
        abrir_puerto(portImpre)
    

        #print estado_error()[1]
        if estado_error()[0] != "128":
            reporteS1  =  obtener_reporteS1()
            reporteZ   =  obtener_reporteZ()
            


            #print "nS1 :"+reporteS1  +"lenght: "+str(len(reporteS1))
            #print " nZ  :" +reporteZ  +"lenght: "+str(len(reporteZ))
            cerrar_puerto()
            estado = True
        else:logger.warning(str(estado_error()[1]))

    except Exception as e:
        if infoERR == True:
            logger.warning(formErro)
        logger.warning(str(e))
    return estado,reporteS1,reporteZ


def cretFileConf():
    confTxt = open(conf, "w+")
    final_de_confTxt = confTxt.tell()
    conftxt="""[General]\n
    codigo = set\n
    numpc = set\n
    tiempo_err = 1\n
    tiempo_init = 1\n
    activacion=1\n\n
    [PathFTP]\n
    portFtp=21
    rutafile = /ftpruta\n\n
    [FeEjecucion]\n
    fecha = /rutaDestino\n\n
    [modoEjecucion]\n
    modo = 1\n
    carpetacompartida = temp"""
    confTxt.writelines(conftxt)
    confTxt.seek(final_de_confTxt)
    confTxt.close()
    logger.warning("Aarchivo de conf ya encuentra disponible ")
    slp(30*60)
    main()


def cwU0Z():
    try:
        global archivoU0Z,Ru0z,reIntento,tiempoErr
        estado = False
        U0Z_ftp          = open(archivoU0Z, "a+")
        final_de_U0Z_ftp = U0Z_ftp.tell()
        if ifExi(Ru0z):
            uozNew   = open(Ru0z, "r+").read()
            listaU0Z = ['%s \n'% str(uozNew)]
            U0Z_ftp.writelines(listaU0Z)
            U0Z_ftp.seek(final_de_U0Z_ftp)
            estado   = True
        U0Z_ftp.close()

    except Exception as e:
        global infoERR
        if infoERR == True:
            logger.warning(formErro)
        logger.warning(str(e))
        estado  =  False
        U0Z_ftp.close()
    return estado


def cwS1():
    try:
        global archivoS1,Rs1
        S1_ftp  =  open(archivoS1, "a+")
        if ifExi(Rs1):
            final_de_S1_ftp = S1_ftp.tell()
            listaS1         = ['%s \n'% open(Rs1, "r+").read()]
            S1_ftp.writelines(listaS1)
            S1_ftp.seek(final_de_S1_ftp)
            estado            =  True

        S1_ftp.close()
        slp(2)
    except Exception as e:
        global infoERR
        if infoERR == True:
            logger.warning(formErro)
        estado  =  False
        logger.warning(str(e))
        S1_ftp.close()
    return estado


def pubU0Z():
    try:
        global archivoU0Z,infoERR
        fileU0Z = str(archivoU0Z)
        file    = open(fileU0Z,'rb')
        if(conexionFTP()['ftp'].storbinary('STOR %s' % fileU0Z, file)):StatusPubU0Z  =  True
        else:StatusPubU0Z  =  False
        file.close()
        logger.info("U0Z publicado al ftp")
        conexionFTP()['ftp'].quit()
    except Exception as e:
        if infoERR == True:logger.warning(formErro)
        file.close()
        logger.warning(str(e))
    return StatusPubU0Z


def pubS1():
    try:
        global archivoS1
        fileS1  =  str(archivoS1)
        file    =  open(fileS1,'rb')
        if(conexionFTP()['ftp'].storbinary('STOR %s' % fileS1, file)):StatusPubS1  =  True
        else:StatusPubS1  =  False
        file.close()
        conexionFTP()['ftp'].quit()
        logger.info("S1 publicado al ftp")
    except Exception as e:
        global infoERR
        if infoERR == True:
            logger.warning(formErro)
        logger.warning(str(e))
    return StatusPubS1


def copiaCCU0Z():
    try:
        global archivoU0Z,carpetaCompartida
        fileU0Z  =  str(archivoU0Z)
        if (path.exists(fileU0Z)):
            #shutil.copy(fileU0Z, carpetaCompartida)
            copy(fileU0Z, carpetaCompartida)
            logger.info("U0z copiado a a la carpeta compartida")
            StatusPubU0Z   =  True
        else:StatusPubU0Z  =  False
    except Exception as e:
        global infoERR
        if infoERR == True:
            logger.warning(formErro)


        logger.warning(str(e))
    return StatusPubU0Z


def copiaCCS1():
    try:
        global archivoS1,carpetaCompartida,tiempoErr
        fileS1  =  str(archivoS1)
        if (path.exists(fileS1)):
            copy(fileS1, carpetaCompartida)
            #shutil.copy(fileS1, carpetaCompartida)
            StatusPubS1  =  True
            logger.info("S1 copiado a a la carpeta compartida")
        else:
            slp(tiempoErr)
            copiaCCS1()
            StatusPubS1 =False
    except Exception as e:
        global infoERR
        if infoERR == True:
            logger.warning(formErro)

        logger.warning(str(e))
    return StatusPubS1


def pubCCU0Z():
    try:
        global archivoU0Z,infoERR,tiempoErr
        fileU0Z  =  str(archivoU0Z)
        file     =  open(fileU0Z,'rb')
        if (path.exists(fileU0Z)):
            if(conexionFTP()['ftp'].storbinary('STOR %s' % fileU0Z, file)):StatusPubU0Z  =  True
            else:StatusPubU0Z  =  True
            file.close()
            logger.info("U0Z publicado al ftp")
            conexionFTP()['ftp'].quit()
        else:
            slp(tiempoErr)
            pubCCU0Z()

    except Exception as e:
        if infoERR == True:logger.warning(formErro)
        logger.warning(str(e))
    return StatusPubU0Z


def pubCCS1():
    try:
        global archivoS1,infoERR
        fileS1   =  str(archivoS1)
        if (path.exists(fileS1)):
            file =  open(fileS1,'rb')
            if(conexionFTP()['ftp'].storbinary('STOR %s' % fileS1, file)):StatusPubS1  = True
            else:StatusPubS1  =  True
            file.close()
            conexionFTP()['ftp'].quit()
            logger.info("S1 publicado al ftp")
            StatusPubS1  =  False
    except Exception as e:
        if infoERR == True:logger.warning(formErro)

        StatusPubS1  =  False
        logger.warning(str(e))
    return StatusPubS1


def mainUno():    
    libM()
    global today,tiempoInit,tiempoErr,archivoU0Z,archivoS1,Rs1,Ru0z, reIntento
    slp(tiempoInit)
    
    try:
        enviado  =  False
        if (activarIntTFHKA()[0]):#si es true
            logger.info('impresora conectada satfactoriamente ')
            if (cwU0Z() and cwS1()):
                logger.info('Archivo U0Z y S1 creado satifactoriamente')
                if conexionFTP()['estatusCftp']:
                    logger.info('conexiona ftp con el servidor en sincronia ')
                    if (pubU0Z() and pubS1()):
                        #conexionFTP()['ftp'].delete(archivoU0Z)
                        #conexionFTP()['ftp'].delete(archivoS1)
                        #conexionFTP()['ftp'].retrlines('LIST')
                        conexionFTP()['ftp'].quit()
                        logger.info('archivo publicado al servidor ftp satfactoriamente ')
                        rm(Rs1)
                        rm(Ru0z)

                        enviado  = True

                    else:
                        enviado  = False

                else:
                    for i in range(reIntento):
                        if conexionFTP()['ftp']:mainUno()
                        else:logger.warning("error  con el servidor ftp intento "+str(i));slp(tiempoErr)#

            else:logger.warning("Error al escribir U0Z y S1")
        else:
            for i in range(reIntento):
                if activarIntTFHKA():mainUno()
                else:logger.warning("Error de impresora verifique conecion inento "+str(i))

    except Exception as e:
        global infoERR
        if infoERR == True:
            logger.warning(formErro)

        logger.warning(str(e))
    return enviado


def mainDos():
    libM()
    global today,tiempoInit,tiempoErr,archivoU0Z,archivoS1,Rs1,Ru0z
    slp(tiempoInit)
    
    try:
        enviado  =  False
        if (activarIntTFHKA()):
            logger.info('impresora conectada satfactoriamente ')
            if (cwU0Z() and cwS1()):
                logger.info('Archivo U0Z y S1 creado satifactoriamente')
                if (copiaCCS1() and copiaCCU0Z()):
                    logger.info('Archivo U0Z y S1 copiado a la carpeta compartida satifactoriamente')
                    enviado  =  True
                    rm(Rs1)
                    rm(Ru0z)
                else:
                    enviado  =  False
                    logger.warning("Error al escribir U0Z y S1")
            else:logger.warning("Error al escribir U0Z y S1")

        else:
            for i in range(reIntento):
                if activarIntTFHKA():mainDos()
                else:logger.warning("Error de impresora verifique conecion inento "+str(i));slp(tiempoErr)


    except Exception as e:
        global infoERR
        if infoERR == True:
            logger.warning(formErro)
        logger.warning(str(e))
    return enviado


def mainTres():
    global tiempoInit,tiempoErr,archivoU0Z,archivoS1,reIntento
    slp(tiempoInit)
    libM()
    try:
        enviado      =  False
        if (pubCCS1() and pubCCU0Z()):
            logger.info('Archivo U0Z y S1 publicado a la carpeta compartida satifactoriamente')
            enviado  =  True
            rm(archivoS1)
            rm(archivoU0Z)
        else:
            slp(tiempoErr)
            logger.warning("Error al copiar archivo U0Z y S1")



    except Exception as e:
        global infoERR
        if infoERR == True:
            logger.warning(formErro)
        logger.warning(str(e))
    return enviado


def nexDay():
    global tiempoInit
    libM()
    fechaEjecu  =  stFcfha()
    while True:
        if (str(fechaEjecu)  ==  str(date.today())):slp(tiempoInit+7200);libM()
        else:fechaEjecu       =  date.today();stFcfha();main()


def modoUno():
    global reIntento,tiempoErr
    if mainUno():        
        logger.info('Ejcucion exitoxa modo 1!')
        nexDay()
    else:
        slp(tiempoErr)
        logger.warning('Error en jecucion en systema ')
        nexDay()


def modoDos():
    global reIntento,tiempoErr

    if mainDos():
        logger.info('Ejcucion exitoxa! modo 2')
        nexDay()
    else:
        slp(tiempoErr)
        logger.warning('Error en jecucion en systema')
        modoDos()


def modoTres():
    global tiempoErr,reIntento

    if mainTres():
        logger.info('Ejcucion exitoxa! modo 3')
        nexDay()
    else:
        slp(tiempoErr)
        logger.warning('Error en jecucion en systema intento ')
        modoTres()


def main():
    global activacion, modoEjecucion


    if activacion ==1:
        if modoEjecucion == 1:modoUno()
        if modoEjecucion == 2:modoDos()
        if modoEjecucion == 3:modoTres()
        if (modoEjecucion > 3  or  modoEjecucion < 0):
            logger.warning('Error archivo de configuracion')
            exit()
    else:logger.warning('Error archivo de configuracion');exit()

libM()
main()









