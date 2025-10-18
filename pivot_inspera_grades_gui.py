#!/usr/bin/env python3

# This script is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This script is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU Lesser General Public License for more details.

# You should have received a copy of the GNU Lesser General Public License
# along with this script. If not, see <https://www.gnu.org/licenses/>.

# Author: Zophonías Oddur Jónsson (with assistant from Copilot)


import sys
import os
import subprocess
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog,
    QVBoxLayout, QHBoxLayout, QCheckBox, QMessageBox
)

class PivotGradesGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Pivot Inspera Grades")
        self.setMinimumWidth(600)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Grades file
        self.grades_label = QLabel("Grades CSV File:")
        self.grades_input = QLineEdit()
        self.grades_browse = QPushButton("Browse")
        self.grades_browse.clicked.connect(self.browse_grades)
        layout.addLayout(self._horizontal_layout([self.grades_label, self.grades_input, self.grades_browse]))

        # Students file
        self.students_label = QLabel("Student Info Excel File (optional):")
        self.students_input = QLineEdit()
        self.students_browse = QPushButton("Browse")
        self.students_browse.clicked.connect(self.browse_students)
        layout.addLayout(self._horizontal_layout([self.students_label, self.students_input, self.students_browse]))

        # Column order file
        self.column_order_label = QLabel("Column Order File (optional):")
        self.column_order_input = QLineEdit()
        self.column_order_browse = QPushButton("Browse")
        self.column_order_browse.clicked.connect(self.browse_column_order)
        layout.addLayout(self._horizontal_layout([self.column_order_label, self.column_order_input, self.column_order_browse]))

        # Output folder and base name
        self.output_folder_label = QLabel("Output Folder:")
        self.output_folder_input = QLineEdit()
        self.output_folder_browse = QPushButton("Browse")
        self.output_folder_browse.clicked.connect(self.browse_output_folder)
        layout.addLayout(self._horizontal_layout([self.output_folder_label, self.output_folder_input, self.output_folder_browse]))

        self.output_base_label = QLabel("Output Base Filename (.xlsx optional):")
        self.output_base_input = QLineEdit()
        layout.addLayout(self._horizontal_layout([self.output_base_label, self.output_base_input]))

        # Regex
        self.regex_label = QLabel("Regex for Column Ordering (optional):")
        self.regex_input = QLineEdit()
        layout.addLayout(self._horizontal_layout([self.regex_label, self.regex_input]))

        # Dry run checkbox
        self.dry_run_checkbox = QCheckBox("Dry Run (show column order only)")
        layout.addWidget(self.dry_run_checkbox)

        # Buttons
        self.run_button = QPushButton("Run")
        self.run_button.clicked.connect(self.run_script)

        self.help_button = QPushButton("Regex Help")
        self.help_button.clicked.connect(self.show_regex_help)

        self.quit_button = QPushButton("Quit")
        self.quit_button.clicked.connect(self.close)

        layout.addLayout(self._horizontal_layout([self.run_button, self.help_button, self.quit_button]))

        self.setLayout(layout)

    def _horizontal_layout(self, widgets):
        h_layout = QHBoxLayout()
        for widget in widgets:
            h_layout.addWidget(widget)
        return h_layout

    def browse_grades(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Select Grades CSV File", "", "CSV Files (*.csv)")
        if filename:
            self.grades_input.setText(filename)
            # Automatically set output folder to match input file
            folder = os.path.dirname(filename)
            self.output_folder_input.setText(folder)

    def browse_students(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Select Student Info Excel File", "", "Excel Files (*.xlsx)")
        if filename:
            self.students_input.setText(filename)

    def browse_column_order(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Select Column Order File", "", "Text Files (*.txt)")
        if filename:
            self.column_order_input.setText(filename)

    def browse_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.output_folder_input.setText(folder)

    def show_regex_help(self):
        help_text = (
            "Regex Sorting Help:\n"
            "--------------------\n"
            "Use a regular expression with ONE capturing group to define how question columns should be sorted.\n"
            "The captured value is used as the sort key.\n"
            "If numeric, sorting is numeric; otherwise, lexicographic.\n\n"
            "Examples:\n"
            "- Okt24-(\\d+)   # Sorts by the number after 'Okt24-'\n"
            "- Q(\\d+)        # Sorts Q1, Q2, Q10 numerically\n"
            "- ([A-Za-z]+)    # Sorts by alphabetic prefix\n"
            "- .*-(\\d+)      # Suggestion: match any prefix ending in a dash followed by digits\n\n"
            "Tips:\n"
            "- Escape backslashes properly: use \\\\d+ instead of \\d+.\n"
            "- If regex doesn't match a column, that column falls back to its original name."
        )
        QMessageBox.information(self, "Regex Help", help_text)

    def run_script(self):
        grades_file = self.grades_input.text().strip()
        students_file = self.students_input.text().strip()
        column_order_file = self.column_order_input.text().strip()
        output_folder = self.output_folder_input.text().strip()
        output_base = self.output_base_input.text().strip()
        regex = self.regex_input.text().strip()
        dry_run = self.dry_run_checkbox.isChecked()

        if not grades_file:
            QMessageBox.critical(self, "Error", "Grades CSV file is required.")
            return

        cmd = ["python3", "pivot_inspera_grades.py", grades_file]

        if students_file:
            cmd += ["-s", students_file]
        if column_order_file:
            cmd += ["-c", column_order_file]
        if regex:
            cmd += ["-r", regex]
        if dry_run:
            try:
                result = subprocess.run(cmd + ["--dry-run"], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
                QMessageBox.information(self, "Dry Run Result", result.stdout)
            except subprocess.CalledProcessError as e:
                QMessageBox.critical(self, "Error", f"Script execution failed:\n{e.stderr}")
            return

        # Determine output file path
        if not output_folder:
            QMessageBox.critical(self, "Error", "Output folder is required.")
            return

        if output_base:
            if not output_base.lower().endswith(".xlsx"):
                output_base += ".xlsx"
            output_file = os.path.join(output_folder, output_base)
        else:
            # Generate default output filename from grades file
            base_name = os.path.splitext(os.path.basename(grades_file))[0]
            output_file = os.path.join(output_folder, f"pivoted-{base_name}.xlsx")

        cmd += ["-o", output_file]

        try:
            subprocess.run(cmd, check=True)
            QMessageBox.information(self, "Success", f"Script executed successfully.\nOutput saved to:\n{output_file}")
        except subprocess.CalledProcessError as e:
            QMessageBox.critical(self, "Error", f"Script execution failed:\n{e.stderr}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    gui = PivotGradesGUI()
    gui.show()
    sys.exit(app.exec_())
