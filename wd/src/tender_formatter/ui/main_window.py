from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from tender_formatter.config import read_template_profile
from tender_formatter.domain import DocumentAnalysis, FormatProfile
from tender_formatter.planner import default_output_path
from tender_formatter.ui.review_model import ReviewModel


class MainWindow(QMainWindow):
    def __init__(self, service, parent=None):
        super().__init__(parent)
        self.service = service
        self.profile: FormatProfile | None = None
        self.analysis: DocumentAnalysis | None = None
        self.setWindowTitle("钢结构技术标书 Word 自动格式化")
        self.resize(960, 680)

        self.stack = QStackedWidget()
        self.stack.addWidget(self._build_file_page())
        self.stack.addWidget(self._build_rules_page())
        self.stack.addWidget(self._build_review_page())
        self.stack.addWidget(self._build_result_page())

        self.back_button = QPushButton("上一步")
        self.next_button = QPushButton("下一步")
        self.back_button.clicked.connect(self._back)
        self.next_button.clicked.connect(self._next)
        self.back_button.setEnabled(False)
        self.next_button.setEnabled(False)

        controls = QHBoxLayout()
        controls.addWidget(self.back_button)
        controls.addStretch()
        controls.addWidget(self.next_button)
        layout = QVBoxLayout()
        layout.addWidget(self.stack)
        layout.addLayout(controls)
        central = QWidget()
        central.setLayout(layout)
        self.setCentralWidget(central)

    def _path_row(self, edit: QLineEdit, callback) -> QWidget:
        button = QPushButton("浏览…")
        button.clicked.connect(callback)
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(edit)
        layout.addWidget(button)
        return row

    def _build_file_page(self) -> QWidget:
        page = QWidget()
        layout = QFormLayout(page)
        layout.addRow(QLabel("步骤 1：选择待处理标书、企业样板和输出文件"))
        self.source_edit = QLineEdit()
        self.template_edit = QLineEdit()
        self.output_edit = QLineEdit()
        self.source_edit.textChanged.connect(self._paths_changed)
        self.template_edit.textChanged.connect(self._paths_changed)
        layout.addRow("待处理标书", self._path_row(self.source_edit, self._choose_source))
        layout.addRow("企业 Word 样板", self._path_row(self.template_edit, self._choose_template))
        layout.addRow("输出文件", self._path_row(self.output_edit, self._choose_output))
        return page

    def _build_rules_page(self) -> QWidget:
        page = QWidget()
        layout = QFormLayout(page)
        layout.addRow(QLabel("步骤 2：调整本项目格式规则"))
        self.body_size = QDoubleSpinBox()
        self.body_size.setRange(6, 72)
        self.body_size.setValue(12)
        self.left_margin = QDoubleSpinBox()
        self.left_margin.setRange(0, 10)
        self.left_margin.setValue(3.0)
        self.image_width = QDoubleSpinBox()
        self.image_width.setRange(1, 50)
        self.image_width.setValue(15.5)
        layout.addRow("正文字号（磅）", self.body_size)
        layout.addRow("左页边距（厘米）", self.left_margin)
        layout.addRow("图片最大宽度（厘米）", self.image_width)
        return page

    def _build_review_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.addWidget(QLabel("步骤 3：确认黄色和高风险识别项"))
        self.review_model = ReviewModel(self)
        self.review_table = QTableView()
        self.review_table.setModel(self.review_model)
        self.review_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.review_table)
        confirm = QPushButton("全部按建议确认")
        confirm.clicked.connect(self.review_model.confirm_all_as_suggested)
        layout.addWidget(confirm)
        self.generate_button = QPushButton("生成规范化文档")
        self.generate_button.setEnabled(False)
        self.generate_button.clicked.connect(self._generate)
        self.review_model.confirmation_changed.connect(
            self.generate_button.setEnabled
        )
        layout.addWidget(self.generate_button)
        return page

    def _build_result_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.addWidget(QLabel("步骤 4：生成结果"))
        self.result_label = QLabel("尚未生成")
        self.result_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(self.result_label)
        layout.addStretch()
        return page

    def _choose_source(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择标书", filter="Word (*.docx)")
        if path:
            self.source_edit.setText(path)
            self.output_edit.setText(str(default_output_path(Path(path))))

    def _choose_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择样板", filter="Word (*.docx)")
        if path:
            self.template_edit.setText(path)

    def _choose_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存结果", filter="Word (*.docx)")
        if path:
            self.output_edit.setText(path)

    def _paths_changed(self):
        valid = Path(self.source_edit.text()).is_file() and Path(
            self.template_edit.text()
        ).is_file()
        if self.stack.currentIndex() == 0:
            self.next_button.setEnabled(valid)

    def _apply_rule_edits(self) -> None:
        self.profile.body.size_pt = self.body_size.value()
        self.profile.page.left_cm = self.left_margin.value()
        self.profile.image.max_width_cm = self.image_width.value()

    def _next(self):
        index = self.stack.currentIndex()
        try:
            if index == 0:
                self.profile = read_template_profile(
                    Path(self.template_edit.text()), "当前项目"
                )
                self.body_size.setValue(self.profile.body.size_pt)
                self.left_margin.setValue(self.profile.page.left_cm)
                self.image_width.setValue(self.profile.image.max_width_cm)
                self.stack.setCurrentIndex(1)
            elif index == 1:
                self._apply_rule_edits()
                self.set_analysis(
                    self.service.analyze(Path(self.source_edit.text()), self.profile)
                )
                self.stack.setCurrentIndex(2)
        except Exception as exc:
            QMessageBox.critical(self, "处理失败", str(exc))
            return
        self._sync_navigation()

    def _back(self):
        self.stack.setCurrentIndex(max(0, self.stack.currentIndex() - 1))
        self._sync_navigation()

    def _sync_navigation(self):
        index = self.stack.currentIndex()
        self.back_button.setEnabled(index > 0 and index < 3)
        self.next_button.setVisible(index < 2)
        if index == 0:
            self._paths_changed()
        elif index == 1:
            self.next_button.setEnabled(True)

    def set_analysis(self, analysis: DocumentAnalysis) -> None:
        self.analysis = analysis
        self.review_model.set_analysis(analysis)

    def _generate(self):
        try:
            result = self.service.format(
                self.analysis,
                self.profile,
                self.review_model.overrides(),
                Path(self.output_edit.text()),
                {},
            )
        except Exception as exc:
            QMessageBox.critical(self, "生成失败", str(exc))
            return
        self.result_label.setText(
            f"输出文件：{result.output}\n检查报告：{result.report}\n"
            f"执行操作：{result.operation_count} 项"
        )
        self.stack.setCurrentIndex(3)
        self._sync_navigation()
