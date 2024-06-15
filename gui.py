"""
I wanted to make an interactive GUI for this project, but I didn't have the time at the end, so this is what I have for now.
"""
from data_analysis import *
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QScrollArea, QTableView
from PyQt5.QtCore import QAbstractTableModel, Qt
import sys
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Financial Analysis Tool')
        self.setGeometry(100, 100, 800, 600)  # x, y, width, height

        #Creating the tab widget
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # List of tab names and their specific setups
        tab_info = [
            ("Data Analysis", self.setup_data_analysis_tab),
            ("Q 1", self.setup_results_tab),
            ("Q 2", self.setup_results_tab),
            ("Q 3", self.setup_results_tab),
            ("Q 4", self.setup_results_tab),
            ("Q 5", self.setup_results_tab)
        ]

        # Create tabs based on the above information
        for index, (tab_name, setup_func) in enumerate(tab_info, start=1):
            tab = QWidget() # Create the tab
            # Adding scrolling functionality 
            scroll_area = QScrollArea() 
            scroll_area.setWidgetResizable(True)
            scroll_area.setMinimumSize(800, 600)
            scroll_area.setWidget(tab)

            # Creating layout
            layout = QVBoxLayout()
            tab.setLayout(layout)
            self.tabs.addTab(scroll_area, tab_name)
            setup_func(layout)
            setattr(self, f'layout{index}', layout)

    def setup_data_analysis_tab(self, layout):
        # Button to load data
        load_button = QPushButton("Load Data")
        load_button.clicked.connect(self.load_excel_data)
        layout.addWidget(load_button)

        # Label to display file path
        data_label = QLabel("No file loaded.")
        layout.addWidget(data_label)
        self.data_label = data_label
    
    def setup_results_tab(self, layout):
        results_label = QLabel("Results will be displayed here.")
        layout.addWidget(results_label)

    def load_excel_data(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls);;All Files (*)", options=options)
        if file_name:
            self.data_label.setText(f"Loaded: {file_name}")
            self.mw, self.bw, self.info, self.price = load_and_process(file_name)
            self.update_q1_tab(self.mw)
            #self.update_q2_tab()
            #self.update_q3_tab()
            #self.update_q4_tab()
            #self.update_q5_tab()
    
    def update_q1_tab(self, mw):
        layout = self.layout2
        # Clear the existing content
        while layout.count():
            child = layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        # Add stocks per day figure
        fig = stocks_per_day(mw)
        canvas = FigureCanvas(fig)
        layout.addWidget(canvas)

        # Add weight distributions figure
        fig, df = weight_distributions(mw)

        # Plot
        canvas = FigureCanvas(fig)
        layout.addWidget(canvas)

        # Table
        model = DataFrameModel(df)
        table_view = QTableView()
        table_view.setModel(model)
        layout.addWidget(table_view)

    #def update_q2_tab():

    #def update_q3_tab():

    #def update_q4_tab():

    #def update_q5_tab():

class DataFrameModel(QAbstractTableModel):
    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return self._data.columns[section]
            if orientation == Qt.Vertical:
                return self._data.index[section]
        return None

if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()
