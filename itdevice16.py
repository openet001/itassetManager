import sys
import os
import time
import qrcode
import cv2
import numpy as np
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QFileDialog, QMessageBox, QMenuBar, QMenu, QAction, QDialog,
    QFormLayout, QDateEdit, QComboBox, QDialogButtonBox, QGroupBox,
    QCheckBox, QSizePolicy, QInputDialog
)
from PyQt5.QtCore import Qt, QDate, QTimer
from PyQt5.QtGui import QIcon, QPixmap, QImage
from datetime import datetime, timedelta
from PIL import Image
from pyzbar.pyzbar import decode


class AssetEditDialog(QDialog):
    def __init__(self, asset_data=None, parent=None):
        super().__init__(parent)
        self.setup_ui(asset_data)

    def setup_ui(self, asset_data):
        self.setWindowTitle("编辑资产信息")
        self.setWindowModality(Qt.ApplicationModal)
        self.resize(700, 500)
        
        self.asset_data = asset_data if asset_data else {}
        layout = QFormLayout()
        
        # 创建表单字段
        self.fields = {
            "资产编号": QLineEdit(),
            "资产名称": QLineEdit(),
            "设备型号": QLineEdit(),
            "设备分类": QComboBox(),
            "设备序列号": QLineEdit(),
            "IP地址": QLineEdit(),
            "采购合同号": QLineEdit(),
            "项目名称": QLineEdit(),
            "负责人": QLineEdit(),
            "使用地点": QLineEdit(),
            "机柜位置": QLineEdit(),
            "资产价格": QLineEdit(),
            "采购日期": QDateEdit(),
            "入库日期": QDateEdit(),
            "上线日期": QDateEdit(),
            "维护有效期": QDateEdit(),
            "供应商编码": QLineEdit(),
            "供应商名称": QLineEdit(),
            "供应商负责人": QLineEdit(),
            "设备当前状态": QComboBox(),
            "备注": QLineEdit()
        }
        
        # 初始化字段
        self.init_fields()
        
        # 添加字段到表单
        self.add_fields_to_layout(layout)
        
        # 添加按钮
        self.add_buttons(layout)
        
        self.setLayout(layout)
        self.populate_form()

    def init_fields(self):
        # 设置设备分类选项
        self.fields["设备分类"].addItems(["服务器", "网络设备", "存储设备", "PC", "笔记本", "打印机", "其他"])
        
        # 设置设备状态选项
        self.fields["设备当前状态"].addItems([
            "未投入使用", "使用中", "维修中", "维修结束", 
            "报废中", "报废流程结束", "更换为新设备"
        ])
        
        # 设置日期控件
        for field in ["采购日期", "入库日期", "上线日期", "维护有效期"]:
            self.fields[field].setDisplayFormat("yyyy-MM-dd")
            self.fields[field].setCalendarPopup(True)
            self.fields[field].setDate(QDate.currentDate())
        
        # 设置默认值
        if not self.asset_data:
            self.fields["设备当前状态"].setCurrentText("未投入使用")
            self.fields["备注"].setText("维修更换信息：无维修或补充")
            self.fields["使用地点"].setText("未指定")
            self.fields["机柜位置"].setText("未指定")

    def add_fields_to_layout(self, layout):
        # 基本信息
        basic_info = QGroupBox("基本信息")
        basic_layout = QFormLayout()
        
        # 资产编号带二维码生成按钮
        asset_id_layout = QHBoxLayout()
        asset_id_layout.addWidget(self.fields["资产编号"])
        qr_btn = QPushButton("生成二维码")
        qr_btn.clicked.connect(self.generate_qr_code)
        asset_id_layout.addWidget(qr_btn)
        basic_layout.addRow("资产编号", asset_id_layout)
        
        basic_layout.addRow("资产名称", self.fields["资产名称"])
        basic_layout.addRow("设备型号", self.fields["设备型号"])
        basic_layout.addRow("设备分类", self.fields["设备分类"])
        basic_layout.addRow("设备序列号", self.fields["设备序列号"])
        basic_info.setLayout(basic_layout)
        
        # 位置信息
        location_info = QGroupBox("位置信息")
        location_layout = QFormLayout()
        location_layout.addRow("使用地点", self.fields["使用地点"])
        location_layout.addRow("机柜位置", self.fields["机柜位置"])
        location_layout.addRow("IP地址", self.fields["IP地址"])
        location_info.setLayout(location_layout)
        
        # 项目信息
        project_info = QGroupBox("项目信息")
        project_layout = QFormLayout()
        project_layout.addRow("采购合同号", self.fields["采购合同号"])
        project_layout.addRow("项目名称", self.fields["项目名称"])
        project_layout.addRow("负责人", self.fields["负责人"])
        project_info.setLayout(project_layout)
        
        # 财务信息
        finance_info = QGroupBox("财务信息")
        finance_layout = QFormLayout()
        finance_layout.addRow("资产价格", self.fields["资产价格"])
        finance_layout.addRow("采购日期", self.fields["采购日期"])
        finance_layout.addRow("入库日期", self.fields["入库日期"])
        finance_layout.addRow("上线日期", self.fields["上线日期"])
        finance_layout.addRow("维护有效期", self.fields["维护有效期"])
        finance_info.setLayout(finance_layout)
        
        # 供应商信息
        vendor_info = QGroupBox("供应商信息")
        vendor_layout = QFormLayout()
        vendor_layout.addRow("供应商编码", self.fields["供应商编码"])
        vendor_layout.addRow("供应商名称", self.fields["供应商名称"])
        vendor_layout.addRow("供应商负责人", self.fields["供应商负责人"])
        vendor_info.setLayout(vendor_layout)
        
        # 状态信息
        status_info = QGroupBox("状态信息")
        status_layout = QFormLayout()
        status_layout.addRow("设备当前状态", self.fields["设备当前状态"])
        status_layout.addRow("备注", self.fields["备注"])
        status_info.setLayout(status_layout)
        
        # 添加到主布局
        layout.addRow(basic_info)
        layout.addRow(location_info)
        layout.addRow(project_info)
        layout.addRow(finance_info)
        layout.addRow(vendor_info)
        layout.addRow(status_info)

    def add_buttons(self, layout):
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        
        # 添加自定义按钮
        add_btn = QPushButton("Add")
        add_btn.clicked.connect(lambda: self.done(2))  # 自定义返回码2表示添加操作
        button_box.addButton(add_btn, QDialogButtonBox.ActionRole)
        
        del_btn = QPushButton("Delete")
        del_btn.clicked.connect(lambda: self.done(3))  # 自定义返回码3表示删除操作
        button_box.addButton(del_btn, QDialogButtonBox.ActionRole)
        
        layout.addRow(button_box)

    def generate_qr_code(self):
        """生成资产编号二维码"""
        asset_id = self.fields["资产编号"].text()
        if not asset_id:
            QMessageBox.warning(self, "警告", "请先填写资产编号")
            return
        
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(asset_id)
        qr.make(fit=True)
        
        img = qr.make_image(fill_color="black", back_color="white")
        
        # 弹出保存对话框
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存二维码", f"{asset_id}_二维码.png", "PNG图像 (*.png)"
        )
        
        if file_path:
            if not file_path.endswith('.png'):
                file_path += '.png'
            img.save(file_path)
            QMessageBox.information(self, "成功", f"二维码已保存到: {file_path}")

    def populate_form(self):
        """填充表单数据"""
        for field_name, widget in self.fields.items():
            if field_name in self.asset_data:
                value = str(self.asset_data[field_name]) if pd.notna(self.asset_data[field_name]) else ""
                if isinstance(widget, QLineEdit):
                    widget.setText(value)
                elif isinstance(widget, QDateEdit):
                    if value:
                        date = QDate.fromString(value, "yyyy-MM-dd")
                        if date.isValid():
                            widget.setDate(date)
                elif isinstance(widget, QComboBox):
                    index = widget.findText(value)
                    if index >= 0:
                        widget.setCurrentIndex(index)

    def get_data(self):
        """获取表单数据"""
        data = {}
        for field_name, widget in self.fields.items():
            if isinstance(widget, QLineEdit):
                data[field_name] = widget.text()
            elif isinstance(widget, QDateEdit):
                data[field_name] = widget.date().toString("yyyy-MM-dd")
            elif isinstance(widget, QComboBox):
                data[field_name] = widget.currentText()
        
        # 设置默认值
        if "备注" not in data or not data["备注"]:
            data["备注"] = "维修更换信息：无维修或补充"
        if "使用地点" not in data or not data["使用地点"]:
            data["使用地点"] = "未指定"
        if "机柜位置" not in data or not data["机柜位置"]:
            data["机柜位置"] = "未指定"
        if "设备当前状态" not in data or not data["设备当前状态"]:
            data["设备当前状态"] = "未投入使用"
            
        return data

class AssetManagementSystem(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.setup_timer()

    def setup_timer(self):
        """设置使用时长计时器"""
        self.start_time = time.time()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_usage_time)
        self.timer.start(60000)  # 每分钟检查一次

    def check_usage_time(self):
        """检查使用时长"""
        elapsed = time.time() - self.start_time
        if elapsed > 1800:  # 30分钟=1800秒
            self.timer.stop()
            self.show_appreciation_message()

    def show_appreciation_message(self):
        """显示赞赏信息"""
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("感谢使用，请体谅程序员的不易。")
        msg.setText("如果觉得好用，请给以作者鼓励和赞赏，您的支持是我们的更新动力！")
        msg.setInformativeText(
            "支付宝/微信打赏：\n\n"
            "联系电话：137-147-29-246\n"
            "联系QQ：847297@qq.com"
        )
        
        # 添加自定义按钮
        copy_btn = msg.addButton("复制信息", QMessageBox.ActionRole)
        copy_btn.clicked.connect(lambda: self.copy_to_clipboard(
            "支付宝/微信打赏：5元、10元均可\n"
            "联系电话：137-147-29-246\n"
            "联系QQ：847297@qq.com"
        ))
        
        msg.addButton(QMessageBox.Ok)
        msg.exec_()

    def copy_to_clipboard(self, text):
        """复制文本到剪贴板"""
        clipboard = QApplication.clipboard()
        clipboard.setText(text)
        QMessageBox.information(self, "成功", "信息已复制到剪贴板")

    def setup_ui(self):
        self.setWindowTitle("IT资产管理系统")
        self.setGeometry(100, 100, 1500, 800)
        
        # 初始化数据
        self.current_file = None
        self.assets_df = pd.DataFrame(columns=[
            "资产编号", "资产名称", "设备型号", "设备分类", "设备序列号", 
            "IP地址", "使用地点", "机柜位置", 
            "采购合同号", "项目名称", "负责人", "资产价格",
            "采购日期", "入库日期", "上线日期", "维护有效期", 
            "供应商编码", "供应商名称", "供应商负责人",
            "设备当前状态", "备注"
        ])
        
        self.create_menu_bar()
        self.setup_main_window()

    def create_menu_bar(self):
        """创建菜单栏"""
        menu_bar = self.menuBar()
        
        # 文件菜单
        file_menu = menu_bar.addMenu("文件(&F)")
        
        open_action = QAction("打开", self)
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)
        
        import_action = QAction("导入", self)
        import_action.triggered.connect(self.import_data)
        file_menu.addAction(import_action)
        
        export_action = QAction("导出", self)
        export_action.triggered.connect(self.export_data)
        file_menu.addAction(export_action)
        
        save_action = QAction("保存", self)
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)
        
        save_as_action = QAction("另存为", self)
        save_as_action.triggered.connect(self.save_file_as)
        file_menu.addAction(save_as_action)
        
        template_action = QAction("创建模板", self)
        template_action.triggered.connect(self.create_template)
        file_menu.addAction(template_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # 编辑菜单
        edit_menu = menu_bar.addMenu("编辑(&E)")
        
        select_all_action = QAction("全选", self)
        select_all_action.triggered.connect(self.select_all)
        edit_menu.addAction(select_all_action)
        
        copy_action = QAction("复制", self)
        copy_action.triggered.connect(self.copy_selected)
        edit_menu.addAction(copy_action)
        
        add_action = QAction("添加资产", self)
        add_action.triggered.connect(self.add_asset)
        edit_menu.addAction(add_action)
        
        # 帮助菜单
        help_menu = menu_bar.addMenu("帮助(&H)")
        
        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def setup_main_window(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout()
        central_widget.setLayout(main_layout)
        
        # 左侧面板
        left_panel = QWidget()
        left_panel.setMaximumWidth(450)
        left_layout = QVBoxLayout()
        left_panel.setLayout(left_layout)
        
        # 文件操作区域
        self.setup_file_operations(left_layout)
        
        # 查询区域
        self.setup_search_area(left_layout)
        
        left_layout.addStretch()
        main_layout.addWidget(left_panel)
        
        # 右侧面板
        right_panel = QWidget()
        right_layout = QVBoxLayout()
        right_panel.setLayout(right_layout)
        
        # 表格
        self.setup_table(right_layout)
        
        main_layout.addWidget(right_panel, stretch=1)

    def setup_file_operations(self, layout):
        file_group = QGroupBox("文件操作----（使用前请备份好您的EXCEL文档，以避免资产信息丢失）")
        file_layout = QVBoxLayout()
        file_group.setLayout(file_layout)
        
        self.file_label = QLabel("当前文件: 无")
        file_layout.addWidget(self.file_label)
        
        btn_layout = QHBoxLayout()
        self.open_file_btn = QPushButton("选择文件")
        self.open_file_btn.clicked.connect(self.open_file)
        btn_layout.addWidget(self.open_file_btn)
        
        self.create_template_btn = QPushButton("创建模板")
        self.create_template_btn.clicked.connect(self.create_template)
        btn_layout.addWidget(self.create_template_btn)
        
        file_layout.addLayout(btn_layout)
        layout.addWidget(file_group)

    def setup_search_area(self, layout):
        search_group = QGroupBox("资产查询")
        search_layout = QFormLayout()
        search_group.setLayout(search_layout)
        
        # 查询条件
        self.search_fields = {
            "资产编号": QLineEdit(),
            "资产名称": QLineEdit(),
            "设备型号": QLineEdit(),
            "设备分类": QComboBox(),
            "设备序列号": QLineEdit(),
            "IP地址": QLineEdit(),
            "采购合同号": QLineEdit(),
            "项目名称": QLineEdit(),
            "负责人": QLineEdit(),
            "供应商名称": QLineEdit(),
            "使用地点": QLineEdit(),
            "机柜位置": QLineEdit(),
            "设备当前状态": QComboBox(),
        }
        
        # 初始化查询字段
        self.init_search_fields()
        
        # 添加查询字段到布局
        self.add_search_fields_to_layout(search_layout)
        
        # 维护有效期查询
        self.setup_maintenance_query(search_layout)
        
        # 查询按钮
        self.setup_search_buttons(search_layout)
        
        layout.addWidget(search_group)

    def init_search_fields(self):
        self.search_fields["设备分类"].addItems([""] + ["服务器", "网络设备", "存储设备", "PC", "笔记本", "打印机", "其他"])
        self.search_fields["设备当前状态"].addItems([""] + [
            "未投入使用", "使用中", "维修中", "维修结束", 
            "报废中", "报废流程结束", "更换为新设备"
        ])

    def add_search_fields_to_layout(self, layout):
        layout.setVerticalSpacing(8)
        layout.setHorizontalSpacing(8)
        
        # 第一行 - 资产编号带二维码导入
        row1_layout = QHBoxLayout()
        row1_layout.addWidget(QLabel("资产编号"))
        
        # 资产编号输入框
        self.search_fields["资产编号"].setPlaceholderText("可扫描二维码自动填写")
        row1_layout.addWidget(self.search_fields["资产编号"])
        
        # 添加二维码导入按钮
        import_qr_btn = QPushButton("导入资产编号二维码图片")
        import_qr_btn.clicked.connect(self.import_qr_for_search)
        row1_layout.addWidget(import_qr_btn)
        
        layout.addRow(row1_layout)
        
        # 第二行
        row2_layout = QHBoxLayout()
        row2_layout.addWidget(QLabel("资产名称"))
        row2_layout.addWidget(self.search_fields["资产名称"])
        row2_layout.addWidget(QLabel("设备型号"))
        row2_layout.addWidget(self.search_fields["设备型号"])
        layout.addRow(row2_layout)
        
        # 第三行
        row3_layout = QHBoxLayout()
        row3_layout.addWidget(QLabel("设备分类"))
        row3_layout.addWidget(self.search_fields["设备分类"])
        row3_layout.addWidget(QLabel("设备序列号"))
        row3_layout.addWidget(self.search_fields["设备序列号"])
        layout.addRow(row3_layout)
        
        # 第四行
        row4_layout = QHBoxLayout()
        row4_layout.addWidget(QLabel("IP地址"))
        row4_layout.addWidget(self.search_fields["IP地址"])
        row4_layout.addWidget(QLabel("采购合同号"))
        row4_layout.addWidget(self.search_fields["采购合同号"])
        layout.addRow(row4_layout)
        
        # 第五行
        row5_layout = QHBoxLayout()
        row5_layout.addWidget(QLabel("项目名称"))
        row5_layout.addWidget(self.search_fields["项目名称"])
        row5_layout.addWidget(QLabel("负责人"))
        row5_layout.addWidget(self.search_fields["负责人"])
        layout.addRow(row5_layout)
        
        # 第六行
        row6_layout = QHBoxLayout()
        row6_layout.addWidget(QLabel("供应商名称"))
        row6_layout.addWidget(self.search_fields["供应商名称"])
        row6_layout.addWidget(QLabel("使用地点"))
        row6_layout.addWidget(self.search_fields["使用地点"])
        layout.addRow(row6_layout)
        
        # 第七行
        row7_layout = QHBoxLayout()
        row7_layout.addWidget(QLabel("机柜位置"))
        row7_layout.addWidget(self.search_fields["机柜位置"])
        row7_layout.addWidget(QLabel("设备状态"))
        row7_layout.addWidget(self.search_fields["设备当前状态"])
        layout.addRow(row7_layout)

    def import_qr_for_search(self):
        """导入二维码并自动填写资产编号（改进版）"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择二维码图片", "", 
            "图片文件 (*.png *.jpg *.jpeg *.bmp *.gif)"
        )
        
        if not file_path:
            return
            
        try:
            # 改进的图片读取方式
            img = cv2.imread(file_path, cv2.IMREAD_GRAYSCALE)
            
            if img is None:
                # 尝试用PIL读取（解决某些JPEG格式问题）
                pil_img = Image.open(file_path)
                img = np.array(pil_img.convert('L'))  # 转为灰度图
                
            # 图像预处理（提高识别率）
            img = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
            
            # 检测二维码
            decoded_objects = decode(img)
            
            if not decoded_objects:
                # 尝试调整图像大小（解决某些分辨率问题）
                height, width = img.shape[:2]
                img = cv2.resize(img, (width*2, height*2), interpolation=cv2.INTER_CUBIC)
                decoded_objects = decode(img)
                
            if not decoded_objects:
                QMessageBox.warning(self, "警告", "未检测到二维码，请尝试更清晰的图片")
                return
                
            # 获取第一个二维码内容
            qr_content = decoded_objects[0].data.decode("utf-8").strip()
            
            if not qr_content:
                QMessageBox.warning(self, "警告", "二维码内容为空")
                return
                
            # 自动填写到资产编号字段
            self.search_fields["资产编号"].setText(qr_content)
            QMessageBox.information(self, "成功", f"已导入资产编号: {qr_content}")
            
        except Exception as e:
            error_msg = f"二维码导入失败:\n{str(e)}\n\n可能原因:\n1. 图片损坏\n2. 非标准二维码\n3. 文件权限问题"
            QMessageBox.critical(self, "错误", error_msg)
            print(f"DEBUG: 二维码识别错误详情: {traceback.format_exc()}")


    def setup_maintenance_query(self, layout):
        maintenance_group = QGroupBox("维护有效期查询")
        maintenance_layout = QHBoxLayout()
        maintenance_group.setLayout(maintenance_layout)
        
        self.maintenance_check = QCheckBox("即将/已超期(60天内)")
        self.maintenance_check.stateChanged.connect(self.toggle_maintenance_query)
        
        self.maintenance_status = QComboBox()
        self.maintenance_status.addItems(["即将到期", "已超期", "全部"])
        self.maintenance_status.setEnabled(False)
        
        maintenance_layout.addWidget(self.maintenance_check)
        maintenance_layout.addWidget(self.maintenance_status)
        
        layout.addRow(maintenance_group)

    def setup_search_buttons(self, layout):
        button_layout = QHBoxLayout()
        self.search_btn = QPushButton("查询")
        self.search_btn.clicked.connect(self.search_assets)
        button_layout.addWidget(self.search_btn)
        
        self.clear_btn = QPushButton("清除条件")
        self.clear_btn.clicked.connect(self.clear_search)
        button_layout.addWidget(self.clear_btn)
        
        layout.addRow(button_layout)

    def setup_table(self, layout):
        self.table = QTableWidget()
        self.table.setColumnCount(15)
        self.table.setHorizontalHeaderLabels([
            "资产编号", "资产名称", "设备型号", "设备分类", "设备序列号", 
            "IP地址", "使用地点", "机柜位置",
            "采购合同号", "项目名称", "负责人", "供应商名称", 
            "维护有效期", "设备状态", "备注"
        ])
        
        # 设置列宽
        self.table.setColumnWidth(0, 120)  # 资产编号
        self.table.setColumnWidth(1, 120)  # 资产名称
        self.table.setColumnWidth(2, 120)  # 设备型号
        self.table.setColumnWidth(3, 100)  # 设备分类
        self.table.setColumnWidth(4, 120)  # 设备序列号
        self.table.setColumnWidth(5, 100)  # IP地址
        self.table.setColumnWidth(6, 120)  # 使用地点
        self.table.setColumnWidth(7, 100)  # 机柜位置
        self.table.setColumnWidth(8, 120)  # 采购合同号
        self.table.setColumnWidth(9, 120)  # 项目名称
        self.table.setColumnWidth(10, 80)  # 负责人
        self.table.setColumnWidth(11, 120) # 供应商名称
        self.table.setColumnWidth(12, 100) # 维护有效期
        self.table.setColumnWidth(13, 100) # 设备状态
        self.table.setColumnWidth(14, 200) # 备注
        
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.doubleClicked.connect(self.edit_asset)
        
        layout.addWidget(self.table)

    def open_file(self):
        """打开资产文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "打开资产文件", "", 
            "Excel文件 (*.xlsx *.xls);;CSV文件 (*.csv);;所有文件 (*)"
        )
        
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    self.assets_df = pd.read_csv(file_path, encoding='utf-8-sig')
                else:
                    self.assets_df = pd.read_excel(file_path, engine='openpyxl')
                
                # 检查必要列
                required_columns = ["资产编号", "资产名称", "设备型号", "设备序列号"]
                missing_cols = [col for col in required_columns if col not in self.assets_df.columns]
                
                if missing_cols:
                    QMessageBox.critical(self, "错误", f"文件缺少必要列: {', '.join(missing_cols)}")
                    return
                
                # 确保日期列是字符串格式
                date_columns = ["采购日期", "入库日期", "上线日期", "维护有效期"]
                for col in date_columns:
                    if col in self.assets_df.columns:
                        self.assets_df[col] = self.assets_df[col].astype(str)
                
                # 检查并添加新字段（如果旧文件没有）
                new_columns = {
                    "使用地点": "未指定",
                    "机柜位置": "未指定",
                    "设备当前状态": "未投入使用",
                    "备注": "维修更换信息：无维修或补充"
                }
                
                for col, default_value in new_columns.items():
                    if col not in self.assets_df.columns:
                        self.assets_df[col] = default_value
                
                self.current_file = file_path
                self.file_label.setText(f"当前文件: {os.path.basename(file_path)}")
                self.display_assets()
                QMessageBox.information(self, "成功", "文件加载成功")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法加载文件: {str(e)}")

    def import_data(self):
        """导入数据"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "导入数据", "", 
            "Excel文件 (*.xlsx *.xls);;CSV文件 (*.csv);;所有文件 (*)"
        )
        
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    new_data = pd.read_csv(file_path)
                else:
                    new_data = pd.read_excel(file_path)
                
                # 检查必要列
                required_columns = ["资产编号", "资产名称", "设备型号", "设备序列号"]
                missing_cols = [col for col in required_columns if col not in new_data.columns]
                
                if missing_cols:
                    QMessageBox.critical(self, "错误", f"导入文件缺少必要列: {', '.join(missing_cols)}")
                    return
                
                # 检查并添加新字段（如果导入文件没有）
                new_columns = {
                    "使用地点": "未指定",
                    "机柜位置": "未指定",
                    "设备当前状态": "未投入使用",
                    "备注": "维修更换信息：无维修或补充"
                }
                
                for col, default_value in new_columns.items():
                    if col not in new_data.columns:
                        new_data[col] = default_value
                
                # 合并数据
                if not self.assets_df.empty:
                    # 检查资产编号是否重复
                    duplicate_ids = set(new_data["资产编号"]).intersection(set(self.assets_df["资产编号"]))
                    if duplicate_ids:
                        reply = QMessageBox.question(
                            self, "确认", 
                            f"发现 {len(duplicate_ids)} 个重复资产编号，是否覆盖现有记录?",
                            QMessageBox.Yes | QMessageBox.No
                        )
                        if reply == QMessageBox.Yes:
                            # 删除重复记录
                            self.assets_df = self.assets_df[~self.assets_df["资产编号"].isin(duplicate_ids)]
                        else:
                            return
                    
                    self.assets_df = pd.concat([self.assets_df, new_data], ignore_index=True)
                else:
                    self.assets_df = new_data
                
                self.display_assets()
                QMessageBox.information(self, "成功", f"成功导入 {len(new_data)} 条记录")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法导入数据: {str(e)}")

    def export_data(self):
        """导出数据到文件"""
        if self.assets_df.empty:
            QMessageBox.warning(self, "警告", "没有数据可导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            "导出数据",
            "",
            "Excel文件 (*.xlsx);;CSV文件 (*.csv);;所有文件 (*)"
        )
        
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    self.assets_df.to_csv(file_path, index=False, encoding='utf-8-sig')
                else:
                    if not file_path.endswith('.xlsx'):
                        file_path += '.xlsx'
                    self.assets_df.to_excel(file_path, index=False)
                
                QMessageBox.information(
                    self,
                    "成功",
                    f"成功导出 {len(self.assets_df)} 条记录到: {file_path}"
                )
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法导出数据: {str(e)}")

    def save_file(self):
        """保存文件"""
        if self.assets_df.empty:
            QMessageBox.warning(self, "警告", "没有数据可保存")
            return
        
        if not self.current_file:
            self.save_file_as()
            return
        
        try:
            self.assets_df.to_excel(self.current_file, index=False)
            QMessageBox.information(self, "成功", "文件保存成功")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法保存文件: {str(e)}")

    def save_file_as(self):
        """另存为"""
        if self.assets_df.empty:
            QMessageBox.warning(self, "警告", "没有数据可保存")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "另存为", "", "Excel文件 (*.xlsx);;所有文件 (*)"
        )
        
        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            
            try:
                self.assets_df.to_excel(file_path, index=False)
                self.current_file = file_path
                self.file_label.setText(f"当前文件: {os.path.basename(file_path)}")
                QMessageBox.information(self, "成功", "文件保存成功")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法保存文件: {str(e)}")

    def create_template(self):
        """创建模板文件"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "创建模板文件", "IT资产模板.xlsx", "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            
            try:
                # 创建带有列名的空DataFrame
                template_df = pd.DataFrame(columns=[
                    "资产编号", "资产名称", "设备型号", "设备分类", "设备序列号", 
                    "IP地址", "使用地点", "机柜位置",
                    "采购合同号", "项目名称", "负责人", "资产价格",
                    "采购日期", "入库日期", "上线日期", "维护有效期", 
                    "供应商编码", "供应商名称", "供应商负责人",
                    "设备当前状态", "备注"
                ])
                
                # 添加示例数据
                example_data = {
                    "资产编号": ["ASSET-2023-001", "ASSET-2023-002"],
                    "资产名称": ["服务器1", "交换机1"],
                    "设备型号": ["DL380", "S5850"],
                    "设备分类": ["服务器", "网络设备"],
                    "设备序列号": ["SN123456", "SN789012"],
                    "IP地址": ["192.168.1.1", "192.168.1.2"],
                    "使用地点": ["数据中心A区", "办公室B区"],
                    "机柜位置": ["机柜A-01", "机柜B-02"],
                    "采购合同号": ["PO-2023-001", "PO-2023-002"],
                    "项目名称": ["项目A", "项目B"],
                    "负责人": ["张三", "王五"],
                    "资产价格": [25000, 18000],
                    "采购日期": ["2023-01-15", "2023-02-10"],
                    "入库日期": ["2023-01-20", "2023-02-15"],
                    "上线日期": ["2023-01-25", "2023-02-20"],
                    "维护有效期": ["2024-01-24", "2024-02-19"],
                    "供应商编码": ["SUP-001", "SUP-002"],
                    "供应商名称": ["戴尔科技", "华为"],
                    "供应商负责人": ["李四", "赵六"],
                    "设备当前状态": ["使用中", "未投入使用"],
                    "备注": ["维修更换信息：无维修或补充", "维修更换信息：2023年3月更换过主板"]
                }
                
                example_df = pd.DataFrame(example_data)
                template_df = pd.concat([template_df, example_df], ignore_index=True)
                
                template_df.to_excel(file_path, index=False)
                QMessageBox.information(self, "成功", f"模板文件已创建: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"创建模板失败: {str(e)}")

    def select_all(self):
        """全选"""
        self.table.selectAll()

    def copy_selected(self):
        """复制选中内容"""
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        
        # 获取选中的行和列
        rows = set(item.row() for item in selected_items)
        cols = set(item.column() for item in selected_items)
        
        # 获取数据
        data = []
        for row in sorted(rows):
            row_data = []
            for col in sorted(cols):
                item = self.table.item(row, col)
                row_data.append(item.text() if item else "")
            data.append("\t".join(row_data))
        
        # 复制到剪贴板
        clipboard = QApplication.clipboard()
        clipboard.setText("\n".join(data))

    def add_asset(self):
        """添加新资产"""
        dialog = AssetEditDialog(parent=self)
        result = dialog.exec_()
        
        if result == QDialog.Accepted or result == 2:  # 2是Add按钮的返回码
            new_data = dialog.get_data()
            
            # 验证资产编号
            if not self.validate_asset_id(new_data["资产编号"]):
                return
            
            # 添加到DataFrame
            new_row = pd.DataFrame([new_data])
            self.assets_df = pd.concat([self.assets_df, new_row], ignore_index=True)
            
            # 刷新显示
            self.display_assets()
            
            QMessageBox.information(self, "成功", "资产添加成功，请记得保存文件")

    def edit_asset(self):
        """编辑资产"""
        if self.assets_df.empty:
            QMessageBox.warning(self, "警告", "没有资产数据")
            return
        
        selected_row = self.table.currentRow()
        if selected_row < 0:
            QMessageBox.warning(self, "警告", "请先选择要编辑的资产")
            return
        
        # 获取资产编号
        asset_id_item = self.table.item(selected_row, 0)
        if not asset_id_item:
            QMessageBox.warning(self, "警告", "无法获取资产编号")
            return
        
        asset_id = asset_id_item.text()
        
        # 找到对应的资产数据
        matched_assets = self.assets_df[self.assets_df["资产编号"].astype(str) == asset_id]
        
        if matched_assets.empty:
            QMessageBox.warning(self, "警告", f"未找到资产编号为 {asset_id} 的记录")
            return
        
        asset_data = matched_assets.iloc[0].to_dict()
        
        # 打开编辑对话框
        dialog = AssetEditDialog(asset_data, self)
        result = dialog.exec_()
        
        if result == QDialog.Accepted:
            # 正常保存修改
            new_data = dialog.get_data()
            self.update_asset_data(asset_id, new_data)
        elif result == 2:  # Add按钮
            # 添加新资产
            new_data = dialog.get_data()
            if self.validate_asset_id(new_data["资产编号"]):
                new_row = pd.DataFrame([new_data])
                self.assets_df = pd.concat([self.assets_df, new_row], ignore_index=True)
                self.display_assets()
                QMessageBox.information(self, "成功", "新资产添加成功")
        elif result == 3:  # Delete按钮
            # 删除资产
            self.delete_asset(asset_id)

    def validate_asset_id(self, asset_id):
        """验证资产编号是否有效"""
        if not asset_id:
            QMessageBox.warning(self, "警告", "资产编号不能为空")
            return False
            
        if asset_id in self.assets_df["资产编号"].values:
            QMessageBox.warning(self, "警告", f"资产编号 {asset_id} 已存在！")
            return False
            
        return True

    def update_asset_data(self, old_id, new_data):
        """更新资产数据"""
        if not self.validate_asset_id(new_data["资产编号"]):
            return
            
        # 更新数据
        mask = self.assets_df["资产编号"] == old_id
        for col, value in new_data.items():
            self.assets_df.loc[mask, col] = value
            
        self.display_assets()
        QMessageBox.information(self, "成功", "资产信息已更新")

    def delete_asset(self, asset_id):
        """删除资产"""
        reply = QMessageBox.question(
            self, "确认删除", 
            f"确定要删除资产 {asset_id} 吗？此操作不可恢复！",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.assets_df = self.assets_df[self.assets_df["资产编号"] != asset_id]
            self.display_assets()
            QMessageBox.information(self, "成功", "资产已删除")

    def search_assets(self):
        """查询资产"""
        if self.assets_df.empty:
            QMessageBox.warning(self, "警告", "请先加载资产文件")
            return
        
        # 获取查询条件
        conditions = {}
        for field_name, widget in self.search_fields.items():
            if isinstance(widget, QLineEdit):
                text = widget.text().strip()
                if text:
                    conditions[field_name] = text
            elif isinstance(widget, QComboBox):
                text = widget.currentText()
                if text:
                    conditions[field_name] = text
        
        # 维护有效期查询
        maintenance_query = None
        if self.maintenance_check.isChecked():
            today = datetime.now().date()
            threshold_date = today + timedelta(days=60)
            
            # 转换维护有效期列为日期格式
            try:
                self.assets_df['维护有效期_date'] = pd.to_datetime(self.assets_df['维护有效期'], errors='coerce').dt.date
                
                status = self.maintenance_status.currentText()
                if status == "即将到期":
                    maintenance_query = (self.assets_df['维护有效期_date'] >= today) & (self.assets_df['维护有效期_date'] <= threshold_date)
                elif status == "已超期":
                    maintenance_query = self.assets_df['维护有效期_date'] < today
                else:  # 全部
                    maintenance_query = self.assets_df['维护有效期_date'] <= threshold_date
            except Exception as e:
                QMessageBox.warning(self, "警告", f"维护有效期格式错误: {str(e)}")
                return
        
        # 构建查询
        try:
            filtered_df = self.assets_df.copy()
            
            # 应用普通查询条件
            for field, value in conditions.items():
                if field in filtered_df.columns:
                    if field in ["设备分类", "设备当前状态"]:
                        filtered_df = filtered_df[filtered_df[field] == value]
                    else:
                        filtered_df = filtered_df[filtered_df[field].astype(str).str.contains(value, case=False, na=False)]
            
            # 应用维护有效期查询
            if maintenance_query is not None:
                filtered_df = filtered_df[maintenance_query]
            
            if filtered_df.empty:
                QMessageBox.information(self, "提示", "没有找到匹配的资产")
            else:
                self.display_assets(filtered_df)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"查询失败: {str(e)}")
        finally:
            # 清理临时列
            if '维护有效期_date' in self.assets_df.columns:
                self.assets_df.drop('维护有效期_date', axis=1, inplace=True)

    def toggle_maintenance_query(self, state):
        """切换维护有效期查询状态"""
        self.maintenance_status.setEnabled(state == Qt.Checked)

    def clear_search(self):
        """清除查询条件"""
        for widget in self.search_fields.values():
            if isinstance(widget, QLineEdit):
                widget.clear()
            elif isinstance(widget, QComboBox):
                widget.setCurrentIndex(0)
        
        self.maintenance_check.setChecked(False)
        self.maintenance_status.setCurrentIndex(0)
        self.display_assets()

    def display_assets(self, df=None):
        """显示资产列表"""
        display_df = df if df is not None else self.assets_df
        
        if display_df.empty:
            self.table.setRowCount(0)
            return
        
        # 设置行数和列数
        self.table.setRowCount(len(display_df))
        
        # 显示字段
        columns_to_display = [
            "资产编号", "资产名称", "设备型号", "设备分类", "设备序列号", 
            "IP地址", "使用地点", "机柜位置",
            "采购合同号", "项目名称", "负责人", "供应商名称", 
            "维护有效期", "设备当前状态", "备注"
        ]
        
        for i, row in display_df.iterrows():
            for j, col in enumerate(columns_to_display):
                value = str(row[col]) if col in row and pd.notna(row[col]) else ""
                item = QTableWidgetItem(value)
                
                # 标记特殊状态
                if col == "设备当前状态":
                    if value == "维修中":
                        item.setBackground(Qt.yellow)
                    elif value == "报废中":
                        item.setBackground(Qt.red)
                    elif value == "更换为新设备":
                        item.setBackground(Qt.green)
                
                # 标记即将/已过期的维护有效期
                if col == "维护有效期" and '维护有效期_date' in display_df.columns:
                    expiry_date = row['维护有效期_date']
                    if pd.notna(expiry_date):
                        today = datetime.now().date()
                        if expiry_date < today:
                            item.setBackground(Qt.red)
                        elif expiry_date <= (today + timedelta(days=60)):
                            item.setBackground(Qt.yellow)
                
                self.table.setItem(i, j, item)

    def show_about(self):
        """显示关于信息"""
        QMessageBox.about(self, "关于", 
            "中小企业IT资产管理系统\n版本 2.0\n"
          "作者：openet001 | 847297@qq.com\n\n"
            "功能特点：\n"
            "- 完整的资产管理功能\n"
            "- 支持Excel/CSV文件导入导出\n"
            "- 资产状态跟踪\n"
            "- 多条件组合查询\n\n"
            "开发时间：2025年"
        )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AssetManagementSystem()
    window.show()
    sys.exit(app.exec_())