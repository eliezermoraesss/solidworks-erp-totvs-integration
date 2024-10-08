from PyQt5.QtCore import QAbstractTableModel, Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QPushButton, QVBoxLayout, QWidget
import pandas as pd


class DataFrameModel(QAbstractTableModel):
    def __init__(self, dataframe, parent=None):
        super().__init__(parent)
        self._dataframe = dataframe
        self._original_dataframe = dataframe.copy()  # Store the original DataFrame
        self.active_filters = {}  # To keep track of column filters

    def rowCount(self, parent=None):
        return len(self._dataframe)

    def columnCount(self, parent=None):
        return len(self._dataframe.columns)

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return str(self._dataframe.iloc[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return self._dataframe.columns[section]
            if orientation == Qt.Vertical:
                return str(self._dataframe.index[section])
        return None

    def apply_filter(self, column, filter_values):
        """Apply filter to the specified column based on the filter_values list"""
        if filter_values:
            self.active_filters[column] = filter_values
        else:
            self.active_filters.pop(column, None)
        self.update_filtered_data()

    def update_filtered_data(self):
        """Update the DataFrame to reflect the current filters"""
        filtered_df = self._original_dataframe.copy()
        for column, values in self.active_filters.items():
            if None in values:
                # Handle filtering None values (NaN)
                filtered_df = filtered_df[filtered_df[column].isna() | filtered_df[column].isin(values)]
            else:
                filtered_df = filtered_df[filtered_df[column].isin(values)]
        self._dataframe = filtered_df
        self.layoutChanged.emit()  # Notify the view that the data has changed

    def clear_filters(self):
        """Clear all filters and restore the original DataFrame"""
        self._dataframe = self._original_dataframe.copy()
        self.active_filters.clear()
        self.layoutChanged.emit()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Sample DataFrame
        data = {'Name': ['Alice', 'Bob', 'Charlie', 'David', None],
                'Age': [25, 30, 35, 40, None],
                'City': ['New York', 'Los Angeles', None, 'Chicago', 'Miami']}
        self.dataframe = pd.DataFrame(data)

        # Model and View
        self.model = DataFrameModel(self.dataframe)
        self.view = QTableView()
        self.view.setModel(self.model)

        # Layout
        layout = QVBoxLayout()

        # Add the table view
        layout.addWidget(self.view)

        # Button to clear all filters
        clear_all_button = QPushButton("Clear All Filters")
        clear_all_button.clicked.connect(self.clear_all_filters)
        layout.addWidget(clear_all_button)

        # Button to apply a sample filter on 'City' column
        apply_filter_button = QPushButton("Apply Filter (City = 'New York' or None)")
        apply_filter_button.clicked.connect(self.apply_city_filter)
        layout.addWidget(apply_filter_button)

        # Button to clear the 'City' column filter
        clear_city_filter_button = QPushButton("Clear 'City' Column Filter")
        clear_city_filter_button.clicked.connect(self.clear_city_filter)
        layout.addWidget(clear_city_filter_button)

        # Main widget
        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

    def clear_all_filters(self):
        """Clear all filters across all columns"""
        self.model.clear_filters()

    def apply_city_filter(self):
        """Apply filter to the 'City' column"""
        self.model.apply_filter('City', ['New York', None])  # Filter by 'New York' or NaN

    def clear_city_filter(self):
        """Clear filter from the 'City' column"""
        self.model.apply_filter('City', None)  # Clear filter by passing None


# Running the Application
if __name__ == '__main__':
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()
