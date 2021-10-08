import CP_dataclass
import logging
from dataclasses_json import dataclass_json
import helper,config,json,sys
import pandas as pd
import numpy as np

from dataclasses import *
from PyQt5.QtWidgets import *
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtCore import Qt,QAbstractTableModel

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

WINDOW_TITLE = "View test result"
FILENAME = sys.argv[1]
FILEPATH = "./Data/Result/" + FILENAME
RESULT_START_INDEX = len(config.RESULT_TEMPLATE)

class pandasModel(QAbstractTableModel):

    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parnet=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None

class window(QMainWindow):
    def __init__(self,df):
        super().__init__()
        self.widget = QWidget()
        self.scroll = QScrollArea()
        self.grid = QGridLayout()

        self.row_counter = 0
        
        self.df = df

        self.initUI()

    def initUI(self):
        self.view_data()

        self.show_window()
        
    def view_data(self):
        model = pandasModel(self.df)
        table_data = QTableView()
        table_data.setModel(model)
        self.grid.addWidget(table_data,self.row_counter,0,10,15)
        self.row_counter += 10
        
        self.stat_data()
    
    def stat_data(self):
        label_statistics_text = "<b>Statistics:</b> (" + str(len(self.df)) + " valid results)"
        label_statistics = QLabel(label_statistics_text)
        self.grid.addWidget(label_statistics,self.row_counter,0,1,15)
        self.row_counter += 1
        
        self.result_columns = self.df.columns[8:]
        self.stat_header()
        #hard-coded
        #self.stat_mode()
        self.stat_mean()
        self.stat_median()
        self.stat_maximum()
        self.stat_minimum()
        #self.stat_upperQuartile()
        #self.stat_lowerQuartile()
        self.stat_range()
        #self.stat_interQuartileRange()
        #self.stat_standardDeviation()
        
    def stat_range(self):
        label_range = QLabel("Range:")
        self.grid.addWidget(label_range,self.row_counter,0,1,1)
        for index,column in enumerate(self.result_columns):
            self.create_range(column,index)
        self.row_counter += 1
        
    def create_range(self,column,index):
        answer = self.df[column].max() - self.df[column].min()
        label_range = QLabel(str(answer))
        self.grid.addWidget(label_range,self.row_counter,index+2,1,1)
        
    def stat_minimum(self):
        label_minimum = QLabel("Minimum:")
        self.grid.addWidget(label_minimum,self.row_counter,0,1,1)
        for index,column in enumerate(self.result_columns):
            self.create_minimum(column,index)
        self.row_counter += 1
        
    def create_minimum(self,column,index):
        answer = self.df[column].min()
        label_minimum = QLabel(str(answer))
        self.grid.addWidget(label_minimum,self.row_counter,index+2,1,1)
        
    def stat_maximum(self):
        label_maximum = QLabel("Maximum:")
        self.grid.addWidget(label_maximum,self.row_counter,0,1,1)
        for index,column in enumerate(self.result_columns):
            self.create_maximum(column,index)
        self.row_counter += 1
        
    def create_maximum(self,column,index):
        answer = self.df[column].max()
        label_maximum = QLabel(str(answer))
        self.grid.addWidget(label_maximum,self.row_counter,index+2,1,1)
        
    def stat_median(self):
        label_median = QLabel("Median:")
        self.grid.addWidget(label_median,self.row_counter,0,1,1)
        for index,column in enumerate(self.result_columns):
            self.create_median(column,index)
        self.row_counter += 1
        
    def create_median(self,column,index):
        answer = self.df[column].median()
        label_median = QLabel(str(answer))
        self.grid.addWidget(label_median,self.row_counter,index+2,1,1)
        
    def stat_header(self):
        for index,column in enumerate(self.result_columns):
            self.create_header(column,index)
        self.row_counter += 1
            
    def create_header(self,column,index):
        label_column = QLabel(column)
        self.grid.addWidget(label_column,self.row_counter,index + 2,1,1)
        
    def stat_mean(self):
        label_mean = QLabel("Mean:")
        self.grid.addWidget(label_mean,self.row_counter,0,1,1)
        for index,column in enumerate(self.result_columns):
            self.create_mean(column,index)
        self.row_counter += 1
    
    def create_mean(self,column,index):
        answer = round(self.df[column].mean(),3)
        label_mean = QLabel(str(answer))
        self.grid.addWidget(label_mean,self.row_counter,index+2,1,1)
        

    def show_window(self):
        self.widget.setLayout(self.grid)
        self.setWindowTitle(WINDOW_TITLE)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(self.widget)
        self.setCentralWidget(self.scroll)
        self.showMaximized()
        self.show()


def main():
    logging.info("Entered view_test_result.py " + FILENAME)
    app = QApplication(sys.argv)
    df = pd.read_csv(FILEPATH,index_col = False)
    ex = window(df)
    #logging.info("Leaving .py")
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
