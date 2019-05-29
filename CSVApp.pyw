import sys
import csv
import datetime
from PyQt5 import QtGui, QtWidgets
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QFileDialog, QApplication, QMainWindow, QTableWidgetItem
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from csv import Dialect
from csvmainwindow import Ui_MainWindow


class CSVTool(Ui_MainWindow):
    def __init__(self, MainWindow):
        Ui_MainWindow.__init__(self)
        self.setupUi(MainWindow)

        csv.register_dialect('shipDialect', delimiter='|')
        self.lineEdits = self.lineEditsToList()
        self.columnHeaders = ['EDI Client ID','Sender\'s Name', 'Pickup Address Line 1','Pickup Address Line 2', 'Pickup Address Line 3', 'Pickup City', 'Pickup Province', 'Pickup Postal Code', 'Recipient\'s Name', 'Delivery Address Line 1', 'Delivery Address Line 2', 'Delivery Address Line 3', 'Delivery City', 'Delivery Province', 'Delivery Postal Code', 'Delivery Country Code Abbrv.', 'Unique Barcode on Package', 'Sender\'s Reference Number', 'Shipment Refernce Number', 'Piece Number In Shipment', 'Piece Wieght', 'Unit of Measurement', 'Special Instructions', 'Service Type', 'Recipient\'s Phone Number', 'Recipient\'s Email Address', 'Pickup Country Code Abbrv.', 'Piece Description', 'Route Code', 'Sender\'s Phone', 'Reference 3', 'Reference 4', 'Labor Code', 'Job Option', 'Email Notification Option', 'Additional Barcode', 'Mass Email Notification', 'Requested Date Time', 'Special Mobile Instructions', 'Ready', 'Length', 'Width', 'Height', 'UOM Length', 'Payment Type Code', 'Overridden Cost', 'Job Description', 'Sales Order Number', 'Second Route Code', 'Account Number', 'Declared Value']
        self.jobs = []

        self.jobs_table = self.tableWidget
        
        self.Import.clicked.connect(self.handleImport)
        self.Export.clicked.connect(self.handleExport)
        self.AddRow.clicked.connect(self.handleAddRow)
        self.RemoveRow.clicked.connect(self.handleRemoveRow)
        self.Clear.clicked.connect(self.handleClear)
    
    def handleImport(self):
        print("Import")
        Tk().withdraw() 
        filename = askopenfilename()

        job = {}
        if filename:
            if filename.endswith('.csv'):
                with open(filename, newline='') as csvIN:
                    reader = csv.reader(csvIN, dialect='shipDialect')
                    for row in reader:
                        i = 0
                        while i < len(row) :
                            job.update({self.columnHeaders[i]: row[i]})
                            i+= 1
                        self.fillFields(job, self.lineEdits) 
            else:
                print('Invalid File Type')

    def lineEditsToList(self):
        l=[]
        l.append(self.clientIDLineEdit)
        l.append(self.senderSNameLineEdit)
        l.append(self.pickupAddressLine1LineEdit)
        l.append(self.pickupAddressLine2LineEdit)
        l.append(self.pickupAddressLine3LineEdit)
        l.append(self.pickupCityLineEdit)
        l.append(self.pickupProvinceLineEdit)
        l.append(self.pickupPostalCodeLineEdit)
        l.append(self.recipientSNameLineEdit)
        l.append(self.deliveryAddressLine1LineEdit)
        l.append(self.deliveryAddressLine2LineEdit)
        l.append(self.deliveryAdrressLine3LineEdit)
        l.append(self.deliveryCityLineEdit)
        l.append(self.deliveryProvinceLineEdit)
        l.append(self.deliveryPostalCodeLineEdit)
        l.append(self.deliveryCountryAbreviationLineEdit)
        l.append(self.uniqueBarcodeOnPkgLineEdit)
        l.append(self.sendersReferenceNumberLineEdit)
        l.append(self.shipmentReferenceNumberLineEdit)
        l.append(self.pieceNumberForShipmentLineEdit)
        l.append(self.pieceWeightLineEdit)
        l.append(self.unitOfMeasurementLineEdit)
        l.append(self.specialInstructionsLineEdit)
        l.append(self.serviceTypeLineEdit)
        l.append(self.recipientPhoneNumberLineEdit)
        l.append(self.recipientEmailAddressLineEdit)
        l.append(self.pickupCountryAbrvLineEdit)
        l.append(self.pieceDescriptionLineEdit)
        l.append(self.routeCodeLineEdit)
        l.append(self.senderSPhoneLineEdit)
        l.append(self.reference3LineEdit)
        l.append(self.reference4LineEdit)
        l.append(self.laborCodeLineEdit)
        l.append(self.jobOptionLineEdit)
        l.append(self.EmailNotificationOptionLineEdit)
        l.append(self.additionalBarcodeLineEdit)
        l.append(self.massEmailNotificationsLineEdit)
        l.append(self.requestedDateTimeLineEdit)
        l.append(self.specialMobileInstructionsLineEdit)
        l.append(self.readyLineEdit)
        l.append(self.lengthLineEdit)
        l.append(self.widthLineEdit)
        l.append(self.heightLineEdit)
        l.append(self.uOMLengthLineEdit)
        l.append(self.paymentTypeCodeLineEdit)
        l.append(self.overridenCostLineEdit)
        l.append(self.jobDescriptionLineEdit)
        l.append(self.salesOrderNumberLineEdit)
        l.append(self.secondRouteCodeLineEdit)
        l.append(self.accountNumberLineEdit)
        l.append(self.declaredValueLineEdit)

        return l
    
    def fillFields(self, job, lineEdits):
        i = 0
        for l in lineEdits:
            l.setText(job[self.columnHeaders[i]])
            i+=1

    def handleAddRow(self):
        job = []
        i = 0
        for l in self.lineEdits:
            job.append(l.text())
            i+=1
            
        self.jobs.append(job)
        print(self.jobs)

        x = 0
        for j in self.jobs:
            y = 0
            while y < 51:
                self.tableWidget.setItem(x,y, QTableWidgetItem(j[y]))
                self.tableWidget.item(x, y).setBackground(QtGui.QColor(255,204,153))
                y+= 1
            x += 1

    def handleClear(self):
        i = 0
        for l in self.lineEdits:
            l.setText('')
            i+=1

    def handleExport(self):
        csv.register_dialect('shipDialect2', delimiter='|', quotechar = '"', quoting = csv.QUOTE_ALL)
        print('Export')
        with open('exportfile.csv', 'w+', newline = '') as csvOUT:
            writer = csv.writer(csvOUT, dialect='shipDialect2')

            writer.writerows(self.jobs)
            print('Exported Succesfully')

    def handleRemoveRow(self):
        if len(self.jobs) > 0:
            print('Remove Row')
            del self.jobs[-1]
            self.jobs_table.removeRow(len(self.jobs))
            rowPosition = self.jobs_table.rowCount() 
            self.jobs_table.insertRow(rowPosition)
        
        
if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = CSVTool(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

