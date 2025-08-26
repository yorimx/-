import sys
import os
import datetime
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QListWidget, QLineEdit, QPushButton, 
                            QLabel, QDateEdit, QDoubleSpinBox, QTabWidget, 
                            QTableWidget, QTableWidgetItem, QMessageBox, 
                            QGroupBox, QFormLayout, QHeaderView, QDialog)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont

# 确保中文显示正常
import matplotlib
matplotlib.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]

class TutoringRecorder(QMainWindow):
    def __init__(self):
        super().__init__()
        self.students = {}  # 存储学生数据 {name: {records: [], payments: []}}
        self.log_file = "tutoring_log.txt"
        self.init_ui()
        self.load_data()

    def init_ui(self):
        # 设置窗口基本属性
        self.setWindowTitle("补课时间记录软件")
        self.setGeometry(100, 100, 900, 600)
        self.setStyleSheet("background-color: #ffffff;")
        
        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        
        # 左侧学生列表
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_panel.setMaximumWidth(200)
        
        # 添加学生区域
        add_student_group = QGroupBox("添加学生")
        add_student_layout = QHBoxLayout()
        self.student_name_input = QLineEdit()
        self.student_name_input.setPlaceholderText("输入学生姓名")
        add_student_btn = QPushButton("添加")
        add_student_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 5px 10px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        add_student_btn.clicked.connect(self.add_student)
        
        add_student_layout.addWidget(self.student_name_input)
        add_student_layout.addWidget(add_student_btn)
        add_student_group.setLayout(add_student_layout)
        
        # 学生列表
        self.student_list = QListWidget()
        self.student_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #e0e0e0;
                border-radius: 3px;
                padding: 5px;
            }
            QListWidget::item {
                padding: 5px;
                border-bottom: 1px solid #f0f0f0;
            }
            QListWidget::item:selected {
                background-color: #e3f2fd;
                color: #0d47a1;
            }
        """)
        self.student_list.itemClicked.connect(self.on_student_selected)
        
        # 启用学生列表的拖拽排序功能
        self.student_list.setDragDropMode(QListWidget.InternalMove)
        self.student_list.setDefaultDropAction(Qt.MoveAction)
        self.student_list.setSelectionMode(QListWidget.SingleSelection)
        self.student_list.model().rowsMoved.connect(self.on_students_reordered)
        
        # 导出按钮
        export_btn = QPushButton("导出Excel")
        export_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 3px;
                margin-top: 10px;
            }
            QPushButton:hover {
                background-color: #0b7dda;
            }
        """)
        export_btn.clicked.connect(self.export_to_excel)
        
        left_layout.addWidget(add_student_group)
        left_layout.addWidget(QLabel("学生列表:"))
        left_layout.addWidget(self.student_list)
        left_layout.addWidget(export_btn)
        
        # 右侧操作区域
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        # 标签页
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabBar::tab {
                padding: 8px 16px;
                background-color: #f5f5f5;
                border: none;
                border-radius: 3px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #e3f2fd;
                color: #0d47a1;
            }
            QTabWidget::pane {
                border: 1px solid #e0e0e0;
                border-radius: 3px;
                padding: 10px;
            }
        """)
        
        # 添加上课记录标签页
        self.attendance_tab = QWidget()
        self.init_attendance_tab()
        self.tabs.addTab(self.attendance_tab, "添加上课记录")
        
        # 上课记录表格标签页
        self.records_tab = QWidget()
        self.init_records_tab()
        self.tabs.addTab(self.records_tab, "上课记录")
        
        # 结算标签页
        self.payment_tab = QWidget()
        self.init_payment_tab()
        self.tabs.addTab(self.payment_tab, "课时结算")
        
        right_layout.addWidget(self.tabs)
        
        # 添加到主布局
        main_layout.addWidget(left_panel)
        main_layout.addWidget(right_panel, 1)
        
        self.show()

    def on_students_reordered(self, source_parent, source_start, source_end, dest_parent, dest_row):
        """处理学生列表重新排序后的逻辑"""
        # 创建新的学生顺序列表
        new_order = []
        for i in range(self.student_list.count()):
            new_order.append(self.student_list.item(i).text())
        
        # 创建新的学生字典，按照新顺序
        reordered_students = {}
        for student in new_order:
            reordered_students[student] = self.students[student]
        
        # 替换原来的学生字典
        self.students = reordered_students
        
        # 保存重新排序后的数据
        self.save_data()
        self.log_action("学生列表顺序已调整")

    def init_attendance_tab(self):
        """初始化上课记录标签页"""
        layout = QVBoxLayout(self.attendance_tab)
        
        # 选择日期和时长
        form_layout = QFormLayout()
        form_layout.setRowWrapPolicy(QFormLayout.DontWrapRows)
        form_layout.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        
        self.date_input = QDateEdit(QDate.currentDate())
        self.date_input.setDisplayFormat("yyyy-MM-dd")
        self.date_input.setCalendarPopup(True)
        
        self.duration_input = QDoubleSpinBox()
        self.duration_input.setRange(0.5, 10)
        self.duration_input.setSingleStep(0.5)
        self.duration_input.setValue(1.0)
        
        # 学生补习科目显示
        self.subjects_display_label = QLabel("未选择学生")
        self.subjects_display_label.setStyleSheet("color: #666666;")
        
        # 添加修改科目按钮
        self.modify_subjects_btn = QPushButton("修改")
        self.modify_subjects_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 4px 8px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #0b7dda;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.modify_subjects_btn.clicked.connect(self.modify_student_subjects)
        self.modify_subjects_btn.setEnabled(False)  # 初始禁用
        
        # 创建水平布局放置标签和按钮
        subjects_layout = QHBoxLayout()
        subjects_layout.addWidget(self.subjects_display_label)
        subjects_layout.addWidget(self.modify_subjects_btn)
        subjects_layout.addStretch()
        
        form_layout.addRow("日期:", self.date_input)
        form_layout.addRow("时长(小时):", self.duration_input)
        form_layout.addRow("补习科目:", subjects_layout)
        
        # 添加按钮
        add_btn = QPushButton("添加上课记录")
        add_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 3px;
                margin-top: 10px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        add_btn.clicked.connect(self.add_attendance)
        self.add_attendance_btn = add_btn  # 保存引用以便禁用/启用
        
        # 当前选中的学生
        self.selected_student_label = QLabel("未选择学生")
        self.selected_student_label.setStyleSheet("color: #666666; font-style: italic;")
        
        # 布局安排
        layout.addWidget(self.selected_student_label)
        layout.addLayout(form_layout)
        layout.addWidget(add_btn)
        layout.addStretch()
        
        # 初始禁用按钮
        add_btn.setEnabled(False)

    def init_records_tab(self):
        """初始化记录表格标签页"""
        layout = QVBoxLayout(self.records_tab)
        
        # 表格
        self.records_table = QTableWidget()
        self.records_table.setColumnCount(3)
        self.records_table.setHorizontalHeaderLabels(["日期", "时长(小时)", "累计时长(小时)"])
        self.records_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.records_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #e0e0e0;
                border-radius: 3px;
                gridline-color: #f0f0f0;
            }
            QHeaderView::section {
                background-color: #f5f5f5;
                padding: 5px;
                border: 1px solid #e0e0e0;
            }
        """)
        # 连接选择信号
        self.records_table.itemSelectionChanged.connect(self.on_record_selected)
        
        # 操作按钮布局
        button_layout = QHBoxLayout()
        self.modify_record_btn = QPushButton("修改记录")
        self.modify_record_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #f57c00;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.modify_record_btn.clicked.connect(self.modify_record)
        self.modify_record_btn.setEnabled(False)  # 初始禁用
        
        self.delete_record_btn = QPushButton("删除记录")
        self.delete_record_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.delete_record_btn.clicked.connect(self.delete_record)
        self.delete_record_btn.setEnabled(False)  # 初始禁用
        
        button_layout.addWidget(self.modify_record_btn)
        button_layout.addWidget(self.delete_record_btn)
        button_layout.addStretch()
        
        # 总时长标签
        self.total_duration_label = QLabel("总时长: 0 小时")
        self.total_duration_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        
        layout.addWidget(self.records_table)
        layout.addLayout(button_layout)
        layout.addWidget(self.total_duration_label)

    def init_payment_tab(self):
        """初始化结算标签页"""
        layout = QVBoxLayout(self.payment_tab)
        
        # 结算表单
        form_layout = QFormLayout()
        form_layout.setRowWrapPolicy(QFormLayout.DontWrapRows)
        form_layout.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        
        self.payment_date_input = QDateEdit(QDate.currentDate())
        self.payment_date_input.setDisplayFormat("yyyy-MM-dd")
        self.payment_date_input.setCalendarPopup(True)
        
        self.payment_hours_input = QDoubleSpinBox()
        self.payment_hours_input.setRange(0.5, 100)
        self.payment_hours_input.setSingleStep(0.5)
        self.payment_hours_input.setValue(1.0)
        
        form_layout.addRow("结算日期:", self.payment_date_input)
        form_layout.addRow("结算课时(小时):", self.payment_hours_input)
        
        # 添加结算按钮
        add_payment_btn = QPushButton("添加结算记录")
        add_payment_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 3px;
                margin-top: 10px;
            }
            QPushButton:hover {
                background-color: #0b7dda;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        add_payment_btn.clicked.connect(self.add_payment)
        self.add_payment_btn = add_payment_btn  # 保存引用以便禁用/启用
        
        # 结算记录表格
        self.payments_table = QTableWidget()
        self.payments_table.setColumnCount(2)
        self.payments_table.setHorizontalHeaderLabels(["日期", "结算课时(小时)"])
        self.payments_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.payments_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #e0e0e0;
                border-radius: 3px;
                gridline-color: #f0f0f0;
                margin-top: 20px;
            }
            QHeaderView::section {
                background-color: #f5f5f5;
                padding: 5px;
                border: 1px solid #e0e0e0;
            }
        """)
        
        # 总结算时长
        self.total_paid_label = QLabel("已结算总时长: 0 小时")
        self.total_paid_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        
        # 剩余时长
        self.remaining_label = QLabel("剩余未结算时长: 0 小时")
        self.remaining_label.setStyleSheet("font-weight: bold; color: #f44336; margin-top: 5px;")
        
        # 布局安排
        layout.addLayout(form_layout)
        layout.addWidget(add_payment_btn)
        layout.addWidget(self.payments_table)
        layout.addWidget(self.total_paid_label)
        layout.addWidget(self.remaining_label)
        layout.addStretch()
        
        # 初始禁用按钮
        add_payment_btn.setEnabled(False)

    def add_student(self):
        """添加学生"""
        name = self.student_name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "警告", "请输入学生姓名")
            return
        
        if name in self.students:
            QMessageBox.warning(self, "警告", "该学生已存在")
            return
        
        # 创建设置补习科目的对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("设置补习科目")
        dialog.resize(300, 150)
        
        layout = QVBoxLayout(dialog)
        
        # 科目输入
        layout.addWidget(QLabel("请输入学生补习的科目\n（多个科目用逗号分隔）:\n例如：数学,英语,物理"))
        subjects_input = QLineEdit()
        subjects_input.setPlaceholderText("数学,英语,物理")
        layout.addWidget(subjects_input)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        ok_btn = QPushButton("确定")
        cancel_btn = QPushButton("取消")
        button_layout.addWidget(ok_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
        
        # 信号连接
        ok_btn.clicked.connect(dialog.accept)
        cancel_btn.clicked.connect(dialog.reject)
        
        # 显示对话框
        if dialog.exec_() == QDialog.Accepted:
            # 获取科目列表并处理
            subjects_text = subjects_input.text().strip()
            subjects = [s.strip() for s in subjects_text.split(",") if s.strip()]
            if not subjects:
                subjects = ["未设置"]
            
            # 添加到学生字典
            self.students[name] = {
                "records": [],  # 格式: [(date, duration, subject), ...]
                "payments": [],  # 格式: [(date, hours), ...]
                "subjects": subjects  # 学生补习的科目列表
            }
            
            # 更新学生列表
            self.student_list.addItem(name)
            self.student_name_input.clear()
            
            # 保存数据
            self.save_data()
            
            # 记录日志
            self.log_action(f"添加了学生: {name}，补习科目: {', '.join(subjects)}")

    def on_student_selected(self, item):
        """当选择学生时更新界面"""
        student_name = item.text()
        self.selected_student_label.setText(f"当前学生: {student_name}")
        
        # 显示学生的补习科目
        if student_name in self.students and "subjects" in self.students[student_name]:
            subjects = self.students[student_name]["subjects"]
            self.subjects_display_label.setText(", ".join(subjects))
        else:
            self.subjects_display_label.setText("未设置")
        
        # 启用按钮
        self.add_attendance_btn.setEnabled(True)
        self.add_payment_btn.setEnabled(True)
        self.modify_subjects_btn.setEnabled(True)  # 启用修改科目按钮
        
        # 更新记录表格
        self.update_records_table(student_name)
        
        # 更新结算表格
        self.update_payments_table(student_name)
        
        # 重置记录操作按钮状态
        self.modify_record_btn.setEnabled(False)
        self.delete_record_btn.setEnabled(False)

    def add_attendance(self):
        """添加上课记录"""
        current_item = self.student_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "警告", "请先选择学生")
            return
        
        student_name = current_item.text()
        date = self.date_input.date().toString("yyyy-MM-dd")
        duration = self.duration_input.value()
        
        # 移除科目字段，只保存日期和时长
        self.students[student_name]["records"].append((date, duration))
        # 按日期排序
        self.students[student_name]["records"].sort(key=lambda x: x[0])
        
        # 更新表格
        self.update_records_table(student_name)
        
        # 保存数据
        self.save_data()
        
        # 记录日志
        self.log_action(f"为 {student_name} 添加了 {duration} 小时的上课记录，日期: {date}")
        
        QMessageBox.information(self, "成功", f"已添加 {duration} 小时的上课记录")

    def update_records_table(self, student_name):
        """更新上课记录表格"""
        records = self.students[student_name]["records"]
        self.records_table.setRowCount(len(records))
        
        # 确保表格有3列（移除科目列）
        if self.records_table.columnCount() != 3:
            self.records_table.setColumnCount(3)
            self.records_table.setHorizontalHeaderLabels(["日期", "时长(小时)", "累计时长(小时)"])
        
        total = 0.0
        cumulative = 0.0
        
        for row, record in enumerate(records):
            # 处理不同格式的记录，兼容旧数据
            if len(record) == 2:
                date, duration = record
            else:
                date, duration = record[0], record[1]  # 忽略科目字段
                
            total += duration
            cumulative += duration
            
            date_item = QTableWidgetItem(date)
            date_item.setTextAlignment(Qt.AlignCenter)
            
            duration_item = QTableWidgetItem(str(duration))
            duration_item.setTextAlignment(Qt.AlignCenter)
            
            cumulative_item = QTableWidgetItem(f"{cumulative:.1f}")
            cumulative_item.setTextAlignment(Qt.AlignCenter)
            
            self.records_table.setItem(row, 0, date_item)
            self.records_table.setItem(row, 1, duration_item)
            self.records_table.setItem(row, 2, cumulative_item)
        
        self.total_duration_label.setText(f"总时长: {total:.1f} 小时")
        
        # 更新剩余时长
        self.update_remaining_hours(student_name)

    def add_payment(self):
        """添加结算记录"""
        current_item = self.student_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "警告", "请先选择学生")
            return
        
        student_name = current_item.text()
        date = self.payment_date_input.date().toString("yyyy-MM-dd")
        hours = self.payment_hours_input.value()
        
        # 检查总时长
        total_duration = sum(d[1] for d in self.students[student_name]["records"])
        total_paid = sum(p[1] for p in self.students[student_name]["payments"])
        
        if total_paid + hours > total_duration:
            QMessageBox.warning(self, "警告", "结算课时不能超过总上课时长")
            return
        
        # 添加结算记录
        self.students[student_name]["payments"].append((date, hours))
        # 按日期排序
        self.students[student_name]["payments"].sort(key=lambda x: x[0])
        
        # 更新表格
        self.update_payments_table(student_name)
        
        # 保存数据
        self.save_data()
        
        # 记录日志
        self.log_action(f"为 {student_name} 添加了 {hours} 小时的结算记录，日期: {date}")
        
        QMessageBox.information(self, "成功", f"已添加 {hours} 小时的结算记录")

    def update_payments_table(self, student_name):
        """更新结算记录表格"""
        payments = self.students[student_name]["payments"]
        self.payments_table.setRowCount(len(payments))
        
        total_paid = 0.0
        
        for row, (date, hours) in enumerate(payments):
            total_paid += hours
            
            date_item = QTableWidgetItem(date)
            date_item.setTextAlignment(Qt.AlignCenter)
            
            hours_item = QTableWidgetItem(str(hours))
            hours_item.setTextAlignment(Qt.AlignCenter)
            
            self.payments_table.setItem(row, 0, date_item)
            self.payments_table.setItem(row, 1, hours_item)
        
        self.total_paid_label.setText(f"已结算总时长: {total_paid:.1f} 小时")
        
        # 更新剩余时长
        self.update_remaining_hours(student_name)

    def update_remaining_hours(self, student_name):
        """更新剩余未结算时长"""
        total_duration = sum(d[1] for d in self.students[student_name]["records"])
        total_paid = sum(p[1] for p in self.students[student_name]["payments"])
        remaining = total_duration - total_paid
        
        self.remaining_label.setText(f"剩余未结算时长: {remaining:.1f} 小时")
    
    def on_record_selected(self):
        """当表格中选择记录时启用按钮"""
        selected_rows = len(self.records_table.selectionModel().selectedRows())
        if selected_rows == 1:
            self.modify_record_btn.setEnabled(True)
            self.delete_record_btn.setEnabled(True)
        elif selected_rows > 1:
            self.modify_record_btn.setEnabled(False)  # 只能修改单行
            self.delete_record_btn.setEnabled(True)   # 可以删除多行
        else:
            self.modify_record_btn.setEnabled(False)
            self.delete_record_btn.setEnabled(False)
    
    def modify_record(self):
        """修改选中的上课记录"""
        selected_items = self.records_table.selectedItems()
        if not selected_items:
            return
            
        # 获取选中的行
        selected_row = selected_items[0].row()
        current_item = self.student_list.currentItem()
        
        if not current_item:
            return
            
        student_name = current_item.text()
        records = self.students[student_name]["records"]
        
        # 获取当前记录信息，处理不同格式的记录
        if len(records[selected_row]) == 2:
            current_date, current_duration = records[selected_row]
        else:
            current_date, current_duration = records[selected_row][0], records[selected_row][1]  # 忽略科目字段
        
        # 创建修改对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("修改上课记录")
        dialog.resize(300, 150)
        
        layout = QVBoxLayout(dialog)
        
        # 日期输入
        date_label = QLabel("日期:")
        date_input = QDateEdit()
        date_input.setDisplayFormat("yyyy-MM-dd")
        date_input.setCalendarPopup(True)
        
        # 时长输入
        duration_label = QLabel("时长(小时):")
        duration_input = QDoubleSpinBox()
        duration_input.setRange(0.5, 100)
        duration_input.setSingleStep(0.5)
        
        # 设置当前值
        date_input.setDate(QDate.fromString(current_date, "yyyy-MM-dd"))
        duration_input.setValue(current_duration)
        
        # 添加到布局
        layout.addWidget(date_label)
        layout.addWidget(date_input)
        layout.addWidget(duration_label)
        layout.addWidget(duration_input)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        ok_btn = QPushButton("确定")
        cancel_btn = QPushButton("取消")
        button_layout.addWidget(ok_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
        
        # 连接信号
        ok_btn.clicked.connect(dialog.accept)
        cancel_btn.clicked.connect(dialog.reject)
        
        # 显示对话框
        if dialog.exec_() == QDialog.Accepted:
            new_date = date_input.date().toString("yyyy-MM-dd")
            new_duration = duration_input.value()
            
            # 更新记录（只保存日期和时长）
            records[selected_row] = (new_date, new_duration)
            # 按日期排序
            records.sort(key=lambda x: x[0])
            
            # 更新表格
            self.update_records_table(student_name)
            
            # 保存数据
            self.save_data()
            
            # 记录日志
            self.log_action(f"修改了 {student_name} 的上课记录，日期: {new_date}，时长: {new_duration} 小时")
            
            QMessageBox.information(self, "成功", "已成功修改上课记录")
    
    def delete_record(self):
        """删除选中的上课记录"""
        selected_rows = sorted(set(item.row() for item in self.records_table.selectedItems()), reverse=True)
        if not selected_rows:
            return
            
        current_item = self.student_list.currentItem()
        if not current_item:
            return
            
        student_name = current_item.text()
        records = self.students[student_name]["records"]
        
        # 确认删除
        if len(selected_rows) == 1:
            confirm = QMessageBox.question(self, "确认删除", 
                                          f"确定要删除这条记录吗？\n日期: {records[selected_rows[0]][0]}\n时长: {records[selected_rows[0]][1]} 小时",
                                          QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        else:
            confirm = QMessageBox.question(self, "确认删除", 
                                          f"确定要删除选中的 {len(selected_rows)} 条记录吗？",
                                          QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if confirm == QMessageBox.Yes:
            # 删除记录
            for row in selected_rows:
                date, duration = records.pop(row)
                self.log_action(f"删除了 {student_name} 的上课记录，日期: {date}，时长: {duration} 小时")
            
            # 按日期排序
            records.sort(key=lambda x: x[0])
            
            # 更新表格
            self.update_records_table(student_name)
            
            # 保存数据
            self.save_data()
            
            QMessageBox.information(self, "成功", f"已成功删除 {len(selected_rows)} 条上课记录")

    def modify_student_subjects(self):
        """修改学生的补习科目"""
        current_item = self.student_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "警告", "请先选择学生")
            return
        
        student_name = current_item.text()
        
        # 创建设置补习科目的对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("修改补习科目")
        dialog.resize(300, 150)
        
        layout = QVBoxLayout(dialog)
        
        # 科目输入
        layout.addWidget(QLabel("请输入学生补习的科目\n（多个科目用逗号分隔）:\n例如：数学,英语,物理"))
        subjects_input = QLineEdit()
        
        # 设置当前科目
        if student_name in self.students and "subjects" in self.students[student_name]:
            subjects_input.setText(", ".join(self.students[student_name]["subjects"]))
        
        layout.addWidget(subjects_input)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        ok_btn = QPushButton("确定")
        cancel_btn = QPushButton("取消")
        button_layout.addWidget(ok_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
        
        # 信号连接
        ok_btn.clicked.connect(dialog.accept)
        cancel_btn.clicked.connect(dialog.reject)
        
        # 显示对话框
        if dialog.exec_() == QDialog.Accepted:
            # 获取科目列表并处理
            subjects_text = subjects_input.text().strip()
            subjects = [s.strip() for s in subjects_text.split(",") if s.strip()]
            if not subjects:
                subjects = ["未设置"]
            
            # 更新学生的补习科目
            self.students[student_name]["subjects"] = subjects
            
            # 更新界面显示
            self.subjects_display_label.setText(", ".join(subjects))
            
            # 保存数据
            self.save_data()
            
            # 记录日志
            self.log_action(f"修改了学生 {student_name} 的补习科目: {', '.join(subjects)}")
            
            QMessageBox.information(self, "成功", "已成功修改学生补习科目")

    def log_action(self, message):
        """记录操作日志"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(log_entry)

    def save_data(self):
        """保存数据到文件"""
        try:
            # 这里使用简单的文本格式保存，实际应用中可以考虑使用JSON或数据库
            with open("tutoring_data.txt", "w", encoding="utf-8") as f:
                for student, data in self.students.items():
                    f.write(f"STUDENT:{student}\n")
                    
                    # 保存补习科目
                    if "subjects" in data:
                        f.write(f"SUBJECTS:{','.join(data['subjects'])}\n")
                    
                    # 保存上课记录
                    for record in data["records"]:
                        # 处理不同格式的记录
                        if len(record) == 2:
                            date, duration = record
                            f.write(f"RECORD:{date},{duration}\n")
                        else:
                            date, duration, subject = record
                            f.write(f"RECORD:{date},{duration},{subject}\n")
                    
                    # 保存结算记录
                    for date, hours in data["payments"]:
                        f.write(f"PAYMENT:{date},{hours}\n")
                
                self.log_action("数据已保存")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存数据失败: {str(e)}")
            self.log_action(f"保存数据失败: {str(e)}")

    def load_data(self):
        """从文件加载数据"""
        try:
            if not os.path.exists("tutoring_data.txt"):
                return
                
            current_student = None
            
            with open("tutoring_data.txt", "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                        
                    parts = line.split(":", 1)
                    if len(parts) != 2:
                        continue
                        
                    type_, content = parts
                    
                    if type_ == "STUDENT":
                        current_student = content
                        if current_student not in self.students:
                            self.students[current_student] = {"records": [], "payments": []}
                            self.student_list.addItem(current_student)
                    elif type_ == "SUBJECTS" and current_student:
                        # 加载学生补习科目
                        subjects = [s.strip() for s in content.split(",") if s.strip()]
                        if subjects:
                            self.students[current_student]["subjects"] = subjects
                    elif type_ == "RECORD" and current_student:
                        record_parts = content.split(",")
                        if len(record_parts) == 2:
                            # 旧格式的记录 (date, duration)
                            date, duration = record_parts
                            try:
                                # 转换为新格式并添加默认科目
                                self.students[current_student]["records"].append((date, float(duration), "未指定"))
                            except ValueError:
                                pass
                        elif len(record_parts) >= 3:
                            # 新格式的记录 (date, duration, subject)
                            date, duration, subject = record_parts[:3]
                            try:
                                self.students[current_student]["records"].append((date, float(duration), subject.strip()))
                            except ValueError:
                                pass
                    elif type_ == "PAYMENT" and current_student:
                        payment_parts = content.split(",")
                        if len(payment_parts) == 2:
                            date, hours = payment_parts
                            try:
                                self.students[current_student]["payments"].append((date, float(hours)))
                            except ValueError:
                                pass
            
            # 确保所有学生都有subjects字段
            for student in self.students:
                if "subjects" not in self.students[student]:
                    self.students[student]["subjects"] = ["未设置"]
                    
                # 对记录进行排序
                self.students[student]["records"].sort(key=lambda x: x[0])
                self.students[student]["payments"].sort(key=lambda x: x[0])
                
            self.log_action("数据已加载")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载数据失败: {str(e)}")
            self.log_action(f"加载数据失败: {str(e)}")

    def export_to_excel(self):
        """导出数据到Excel文件"""
        if not self.students:
            QMessageBox.warning(self, "警告", "没有学生数据可导出")
            return
            
        try:
            # 创建一个ExcelWriter对象
            filename = f"补课时间记录{datetime.date.today().strftime('%Y%m%d')}.xlsx"
            with pd.ExcelWriter(filename, engine="openpyxl") as writer:
                # 总览表
                overview_data = []
                for student, data in self.students.items():
                    total_duration = sum(d[1] for d in data["records"])
                    total_paid = sum(p[1] for p in data["payments"])
                    remaining = total_duration - total_paid
                    
                    # 获取补习科目
                    subjects = data.get("subjects", ["未设置"])
                    subjects_text = ", ".join(subjects)
                    
                    overview_data.append({
                        "学生姓名": student,
                        "补习科目": subjects_text,
                        "总上课时长(小时)": total_duration,
                        "已结算时长(小时)": total_paid,
                        "剩余时长(小时)": remaining
                    })
                
                overview_df = pd.DataFrame(overview_data)
                # 添加导出时间信息
                export_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                overview_df.loc[len(overview_df)] = {"学生姓名": "记录时间", "补习科目": "", "总上课时长(小时)": export_time}
                overview_df.to_excel(writer, sheet_name="总览", index=False)
            
            # 每个学生的详细记录
            for student, data in self.students.items():
                    # 上课记录
                    records_data = []
                    cumulative = 0.0
                    for record in data["records"]:
                        # 处理不同格式的记录
                        if len(record) == 2:
                            date, duration = record
                        else:
                            date, duration = record[0], record[1]  # 忽略科目字段
                           
                        cumulative += duration
                        records_data.append({
                            "日期": date,
                            "时长(小时)": duration,
                            "累计时长(小时)": cumulative
                        })
                    
                    records_df = pd.DataFrame(records_data)
                    records_df.to_excel(writer, sheet_name=f"{student}_上课记录", index=False)
                    
                    # 结算记录
                    payments_data = []
                    cumulative_paid = 0.0
                    for date, hours in data["payments"]:
                        cumulative_paid += hours
                        payments_data.append({
                            "日期": date,
                            "结算时长(小时)": hours,
                            "累计结算(小时)": cumulative_paid
                        })
                    
                    payments_df = pd.DataFrame(payments_data)
                    payments_df.to_excel(writer, sheet_name=f"{student}_结算记录", index=False)
            
            self.log_action(f"数据已导出到 {filename}")
            QMessageBox.information(self, "成功", f"数据已成功导出到 {filename}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出Excel失败: {str(e)}")
            self.log_action(f"导出Excel失败: {str(e)}")

if __name__ == "__main__":
    # 确保中文显示正常
    font = QFont("SimHei")
    app = QApplication(sys.argv)
    app.setFont(font)
    window = TutoringRecorder()
    sys.exit(app.exec_())
