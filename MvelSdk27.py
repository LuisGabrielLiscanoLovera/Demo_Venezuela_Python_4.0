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
        if resp:print   "Impresora Conectada Correctamente en: "
        else:print      "Impresora no Conectada o Error Accediendo al Puerto"
    except serial.SerialException:print "Impresora no Conectada o Error Accediendo al Puerto"

def obtener_reporteZ():
    global printer
    reporte = printer.GetZReport()
    salida= "Numero Ultimo Reporte Z: "				+ str(reporte._numberOfLastZReport) #1 L4
    salida+= "\nFecha Ultimo Reporte Z: "   		+ str(reporte._zReportDate)			#2 L6
    salida+= "\nNumero Ultima Factura: "			+ str(reporte._numberOfLastInvoice) #3 L8
    salida+= "\nFecha Ultima Factura: "				+ str(reporte._lastInvoiceDate)#4 L6
    salida+= "\nHora Ultima Factura: "				+ str(reporte._lastInvoiceTime)		#5 L4
    salida+= "\nVentas Exento: "					+ str(reporte._freeSalesTax)#6 L10		
    salida+= "\nBase Imponible Ventas IVA G: " 	    + str(reporte._generalRate1Sale)	#7 L10
    salida+= "\nImpuesto IVA G: "					+ str(reporte._generalRate1Tax)	#8L10	
    salida+= "\nBase Imponible Ventas IVA R: "		+ str(reporte._reducedRate2Sale)#9 L10	
    salida+= "\nImpuesto IVA R: "				    + str(reporte._reducedRate2Tax) #10 L0
    salida+= "\nBase Imponible Ventas IVA A: "	    + str(reporte._additionalRate3Sal)#11 L10
    salida+= "\nImpuesto IVA A: "				    + str(reporte._additionalRate3Tax)	#12 L10
    salida+= "\nNota de Debito Exento: "		    + str(reporte._freeTaxDebit)#13 L10
    salida+= "\nBI IVA G en Nota de Debito: "	    + str(reporte._generalRateDebit) #14 L10
    salida+= "\nImpuesto IVA G en Nota de Debito: " + str(reporte._generalRateTaxDebit)#15 L10
    salida+= "\nBI IVA R en Nota de Debito: "	    + str(reporte._reducedRateDebit)#16 L10
    salida+= "\nImpuesto IVA R en Nota de Debito: " + str(reporte._reducedRateTaxDebit)#17 L10
    salida+= "\nBI IVA A en Nota de Debito: "	    + str(reporte._additionalRateDebit)#18 L10
    salida+= "\nImpuesto IVA A en Nota de Debito: " + str(reporte._additionalRateTaxDebit)#L10
    salida+= "\nNota de Credito Exento: "		    + str(reporte._freeTaxDevolution)#L10
    print salida

def estado_error():
    global printer
    stado = printer.ReadFpStatus()
    print stado

def obtener_estado():
    global printer
    estado = "S1"
    if estado == "S1":
        estado_s1 = printer.GetS1PrinterData()
        salida= "01"																					#1 L2
        salida+= "\nNumero Cajero: "					  +	str(estado_s1._cashierNumber)				#nro_cajero 2 L2
        salida+= "\nSubtotal Ventas: " 				      +	str(estado_s1._totalDailySales) 			#sub_total_ventas 3 L17
        salida+= "\nNumero Ultima Factura: "		      +	str(estado_s1._lastInvoiceNumber)			#nro_ultima_factura 4 L8
        salida+= "\nCantidad Facturas Hoy: " 		      +	str(estado_s1._quantityOfInvoicesToday)		#cant_fact_emitidas_dia 5 L5
        salida+= "\nCantidad Notas de Credito Hoy: "	  +	str(estado_s1._quantityOfNCToday)			#cant_notas_credito 13 L7
        salida+= "\nNumero Ultimo Documento No Fiscal: "  + str(estado_s1._numberNonFiscalDocuments) 	#nro_ultimo_doc_no_fiscal
        salida+= "\nCantidad de Documentos No Fiscales: " +	str(estado_s1._quantityNonFiscalDocuments)	#6 L 5
        salida+= "\nCantidad de Reportes Fiscales: "      + str(estado_s1._fiscalReportsCounter)		#contador_reporte_memoria_fiscal 8 L4
        salida+= "\nCantidad de Reportes Z: "			  + str(estado_s1._dailyClosureCounter)			#contador_cierre_diario_z 7L4
        salida+= "\nNumero de RIF: "					  + str(estado_s1._rif)                      	#rif 9 L11
        salida+= "\nNumero de Registro: "				  + str(estado_s1._registeredMachineNumber)		#nro_registro_maqui 10 L10
        salida+= "\nHora de la Impresora: "				  + str(estado_s1._currentPrinterTime)			#hora_actual_impre 11 L4
        salida+= "\nFecha de la Impresora: " 			  + str(estado_s1._currentPrinterDate)			#fecha_actual_impre 12 L6
        print salida

def cerrar_puerto():
    global printer
    resp = printer.CloseFpctrl()
    if not resp:print "Impresora Desconectada"
    else:print "Error"


abrir_puerto("COM18")
estado_error()
obtener_reporteZ()
obtener_estado()
cerrar_puerto()
