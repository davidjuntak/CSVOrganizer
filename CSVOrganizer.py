import csv
import os
import sys
import time
import logging
import traceback
import operator
from PyQt4 import QtCore, QtGui, uic
from PyQt4.QtCore import * 

APP_NAME = "CSVOrganizer"
APP_VERSION = "0.1"
APP_FULLNAME = "%s %s" % ( APP_NAME, APP_VERSION )
APP_AUTHOR = "david.simanjuntak"
AUTHOR_EMAIL = "david.juntak11@gmail.com"

uiFile = os.path.join(os.path.dirname(__file__), 'CSVOrganizer', 'CSVOrganizer.ui')
form_class, base_class = uic.loadUiType(uiFile)
logger = logging.getLogger( __name__ )
logger.setLevel( logging.DEBUG )

class CSVOrganizer(base_class, form_class):
    def __init__( self, parent=None ):
        super(base_class, self).__init__(parent)
        self.setupUi(self)
        self.setWindowTitle( APP_FULLNAME )

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile().toLocal8Bit().data()
            if os.path.isfile(path):
                self.txtFileLocation.setText( path )
                
        csvFilePath = self.txtFileLocation.text()
        inputFile = open( csvFilePath, "rb" )
        csvReader = csv.reader( inputFile, quotechar = '"', delimiter = ',' )
        csvData = [ line for line in csvReader ]
        self._removeDuplicateRow( csvData )
        
    def _getFileName( self, csvFilePath ):
        splittedCsv = csvFilePath.split( "/" )
        lastIndex = len( splittedCsv ) - 1
        csvFileName = splittedCsv[ lastIndex ]
        splittedFileName = csvFileName.split( "." )
        fileName = splittedFileName[ 0 ]
        return fileName
        
    def _writeAction( self, action ):
        self.txtActionLog.appendPlainText( action )
        self.txtActionLog.repaint()
        
    def createTable( self, tableView, tableData ):
        tableModel = TableModel( tableData, self ) 
        tableView.setModel( tableModel )
        tableView.resizeColumnsToContents()
        
        return tableView
        
    def _removeDuplicateRow( self, csvData, message = True ):
        cleanedData = []
        headerColumnLength = len( csvData[ 1 ] )
        for index in range( len( csvData ) ):
            if index > 1 :
                if csvData[ index ][:3] != csvData[ index - 1 ][:3]:
                    cleanedCsvData = self._insertAdditionalColumn( csvData[ index ], headerColumnLength )
                    cleanedData.append( cleanedCsvData )
            else:
                cleanedData.append( csvData[ index ] )
        
        self.createTable( self.tbvPreview, cleanedData )
        
        if message: self._writeAction( "Duplicate row removed" )
            
        return cleanedData
        
    def _insertAdditionalColumn( self, selectedCsvData, headerColumnLength ):
        rowColumnLength = len( selectedCsvData )
        if headerColumnLength > rowColumnLength:
            additionalColumn = headerColumnLength - rowColumnLength
            for column in range( rowColumnLength, rowColumnLength + additionalColumn ):
                selectedCsvData.insert( column, "" )
        return selectedCsvData

    @QtCore.pyqtSlot()
    def on_btnShiftColumn_clicked( self ):
        csvFilePath = self.txtFileLocation.text()
        inputFile = open( csvFilePath, "rb" )
        csvReader = csv.reader( inputFile, quotechar = '"', delimiter = ',' )
        csvData = [ line for line in csvReader ]
        cleanedData = self._removeDuplicateRow( csvData, False )
        startingColumn = int( self.spbStartingColumn.text() )
        shiftColumn = int( self.spbShiftColumn.text() )
        additionalColumn = startingColumn + shiftColumn
        headerColumnLength = len( cleanedData[ 0 ] )
        overColumn = False
        additionalColumnHeader = 0
        
        for index in range( len( cleanedData ) ):
            if index > 1 :
                if cleanedData[ index ][:3] != cleanedData[ index - 1 ][:3]:
                    if shiftColumn > 0:
                        for column in range( startingColumn, additionalColumn ):
                            cleanedData[ index ].insert( startingColumn, "" )
                            selectedRowLength = len( cleanedData[ index ] )
                            
                            if selectedRowLength > headerColumnLength:
                                overColumn = True
                                additionalColumnHeader = selectedRowLength - headerColumnLength
                    else:
                        for column in range( additionalColumn, startingColumn ):
                            cleanedData[ index ].pop( additionalColumn - 1 )
                            self._insertAdditionalColumn( cleanedData[ index ], headerColumnLength )
                            
        if overColumn:
            for index in range( headerColumnLength, headerColumnLength + additionalColumnHeader ):
                cleanedData[ 0 ].insert( headerColumnLength, "" )
                cleanedData[ 1 ].insert( headerColumnLength, "" )
        
        self.createTable( self.tbvPreview, cleanedData )
        self._writeAction( "Shifting %s column start from column %s" % ( shiftColumn, startingColumn ) )
        
    @QtCore.pyqtSlot()
    def on_btnSave_clicked( self ):
        csvFilePath = self.txtFileLocation.text()
        fileName = self._getFileName( csvFilePath ) + "_clean.csv"
        
        if( csvFilePath == "" ):
            QtGui.QMessageBox.warning(
                self,
                "CSV File Not Found",
                "Please Select CSV File First",
                QtGui.QMessageBox.Ok )
        else:
            outputFile = open( fileName, "wb" )
            csvWriter = csv.writer( outputFile, quotechar = '"', quoting = csv.QUOTE_ALL, delimiter = ',' )
            
            for data in self.cleanedData:
                csvWriter.writerow( data )
                
        self._writeAction( "Saving data into csv file (%s)" % fileName )
        
class TableModel(QAbstractTableModel): 
    def __init__( self, tableContent, parent = None, *args ): 
        QAbstractTableModel.__init__( self, parent, *args ) 
        self.tableContent = tableContent
 
    def rowCount( self, parent ): 
        return len( self.tableContent ) 
 
    def columnCount( self, parent ): 
        return len( self.tableContent[0] ) 
 
    def data( self, index, role ): 
        if not index.isValid(): 
            return QVariant() 
        elif role != Qt.DisplayRole: 
            return QVariant() 
        return QVariant( self.tableContent[index.row()][index.column()] ) 
                
if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)

    window = CSVOrganizer()
    window.show()

    sys.exit(app.exec_())