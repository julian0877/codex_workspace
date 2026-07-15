from pathlib import Path

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Qt, Signal, Slot
from PySide6.QtWidgets import (
    QDoubleSpinBox,
    QCheckBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QSpinBox,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from tender_formatter.config import read_template_profile
from tender_formatter.cover import find_cover_fields
from tender_formatter.domain import DocumentAnalysis, FormatProfile
from tender_formatter.planner import default_output_path
from tender_formatter.ui.review_model import ReviewModel


class WorkerSignals(QObject):
    result = Signal(object)
    error = Signal(str)
    finished = Signal()


class Worker(QRunnable):
    def __init__(self, function):
        super().__init__()
        self.function = function
        self.signals = WorkerSignals()

    @Slot()
    def run(self):
        try:
            self.signals.result.emit(self.function())
        except Exception as exc:
            self.signals.error.emit(str(exc))
        finally:
            self.signals.finished.emit()


class MainWindow(QMainWindow):
    def __init__(self, service, parent=None):
        super().__init__(parent)
        self.service = service
        self.profile: FormatProfile | None = None
        self.analysis: DocumentAnalysis | None = None
        self._paragraphs_confirmed = False
        self._structure_acknowledged = True
        self.thread_pool = QThreadPool.globalInstance()
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
        self.status_label = QLabel("")
        controls.addWidget(self.status_label)
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
        self.rules_layout = layout
        layout.addRow(QLabel("步骤 2：调整本项目格式规则"))
        self.body_size = QDoubleSpinBox()
        self.body_size.setRange(6, 72)
        self.body_size.setValue(12)
        self.body_font = QLineEdit("宋体")
        self.line_spacing = QDoubleSpinBox()
        self.line_spacing.setRange(1, 3)
        self.line_spacing.setSingleStep(0.1)
        self.line_spacing.setValue(1.5)
        self.top_margin = QDoubleSpinBox()
        self.top_margin.setRange(0, 10)
        self.top_margin.setValue(2.5)
        self.bottom_margin = QDoubleSpinBox()
        self.bottom_margin.setRange(0, 10)
        self.bottom_margin.setValue(2.5)
        self.left_margin = QDoubleSpinBox()
        self.left_margin.setRange(0, 10)
        self.left_margin.setValue(3.0)
        self.right_margin = QDoubleSpinBox()
        self.right_margin.setRange(0, 10)
        self.right_margin.setValue(2.5)
        self.body_page_start = QSpinBox()
        self.body_page_start.setRange(1, 9999)
        self.toc_levels = QSpinBox()
        self.toc_levels.setRange(1, 3)
        self.toc_levels.setValue(3)
        self.first_page_different = QCheckBox("正文节首页不同")
        self.first_page_different.setChecked(True)
        self.odd_even_pages = QCheckBox("奇偶页不同")
        self.heading_sizes: dict[int, QDoubleSpinBox] = {}
        for level, default in ((1, 16), (2, 14), (3, 12)):
            spin = QDoubleSpinBox()
            spin.setRange(6, 72)
            spin.setValue(default)
            self.heading_sizes[level] = spin
        self.table_width = QSpinBox()
        self.table_width.setRange(10, 100)
        self.table_width.setValue(100)
        self.repeat_table_header = QCheckBox("跨页重复首行")
        self.repeat_table_header.setChecked(True)
        self.image_width = QDoubleSpinBox()
        self.image_width.setRange(1, 50)
        self.image_width.setValue(15.5)
        layout.addRow("正文中文字体", self.body_font)
        layout.addRow("正文字号（磅）", self.body_size)
        layout.addRow("正文行距（倍）", self.line_spacing)
        layout.addRow("上页边距（厘米）", self.top_margin)
        layout.addRow("下页边距（厘米）", self.bottom_margin)
        layout.addRow("左页边距（厘米）", self.left_margin)
        layout.addRow("右页边距（厘米）", self.right_margin)
        layout.addRow("正文起始页码", self.body_page_start)
        layout.addRow("目录显示层级", self.toc_levels)
        layout.addRow("首页页眉页脚", self.first_page_different)
        layout.addRow("奇偶页页眉页脚", self.odd_even_pages)
        for level, spin in self.heading_sizes.items():
            layout.addRow(f"{level} 级标题字号（磅）", spin)
        layout.addRow("表格宽度（%）", self.table_width)
        layout.addRow("表格跨页设置", self.repeat_table_header)
        layout.addRow("图片最大宽度（厘米）", self.image_width)
        self.cover_group = QGroupBox("封面字段")
        self.cover_layout = QFormLayout(self.cover_group)
        self.cover_edits: dict[str, QLineEdit] = {}
        layout.addRow(self.cover_group)
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
        self.structure_label = QLabel("")
        self.structure_label.setWordWrap(True)
        layout.addWidget(self.structure_label)
        self.ack_structure_button = QPushButton("我已检查上述结构风险")
        self.ack_structure_button.clicked.connect(
            self._acknowledge_structure_warnings
        )
        self.ack_structure_button.setVisible(False)
        layout.addWidget(self.ack_structure_button)
        confirm = QPushButton("确认所选项")
        confirm.clicked.connect(self._confirm_selected_review)
        layout.addWidget(confirm)
        self.generate_button = QPushButton("生成规范化文档")
        self.generate_button.setEnabled(False)
        self.generate_button.clicked.connect(self._generate)
        self.review_model.confirmation_changed.connect(
            self._review_confirmation_changed
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
        self.profile.body.east_asia_font = self.body_font.text().strip() or "宋体"
        self.profile.body.size_pt = self.body_size.value()
        self.profile.body.line_spacing = self.line_spacing.value()
        self.profile.page.top_cm = self.top_margin.value()
        self.profile.page.bottom_cm = self.bottom_margin.value()
        self.profile.page.left_cm = self.left_margin.value()
        self.profile.page.right_cm = self.right_margin.value()
        self.profile.page.body_page_start = self.body_page_start.value()
        self.profile.page.first_page_different = self.first_page_different.isChecked()
        self.profile.page.odd_even_pages = self.odd_even_pages.isChecked()
        self.profile.toc_levels = self.toc_levels.value()
        for level, spin in self.heading_sizes.items():
            self.profile.headings[level].size_pt = spin.value()
        self.profile.table.width_percent = self.table_width.value()
        self.profile.table.repeat_header = self.repeat_table_header.isChecked()
        self.profile.image.max_width_cm = self.image_width.value()

    def _load_cover_fields(self, template: Path) -> None:
        while self.cover_layout.rowCount():
            self.cover_layout.removeRow(0)
        self.cover_edits = {}
        for field in sorted(find_cover_fields(template) - {"目录"}):
            edit = QLineEdit()
            self.cover_edits[field] = edit
            self.cover_layout.addRow(field, edit)

    def _confirm_selected_review(self) -> None:
        row = self.review_table.currentIndex().row()
        self.review_model.confirm_row(row)

    def _next(self):
        index = self.stack.currentIndex()
        try:
            if index == 0:
                self.profile = read_template_profile(
                    Path(self.template_edit.text()), "当前项目"
                )
                self._load_cover_fields(Path(self.template_edit.text()))
                self.body_font.setText(self.profile.body.east_asia_font)
                self.body_size.setValue(self.profile.body.size_pt)
                self.line_spacing.setValue(self.profile.body.line_spacing)
                self.top_margin.setValue(self.profile.page.top_cm)
                self.bottom_margin.setValue(self.profile.page.bottom_cm)
                self.left_margin.setValue(self.profile.page.left_cm)
                self.right_margin.setValue(self.profile.page.right_cm)
                self.body_page_start.setValue(self.profile.page.body_page_start)
                self.toc_levels.setValue(self.profile.toc_levels)
                self.first_page_different.setChecked(
                    self.profile.page.first_page_different
                )
                self.odd_even_pages.setChecked(
                    self.profile.page.odd_even_pages
                )
                for level, spin in self.heading_sizes.items():
                    spin.setValue(self.profile.headings[level].size_pt)
                self.table_width.setValue(self.profile.table.width_percent)
                self.repeat_table_header.setChecked(
                    self.profile.table.repeat_header
                )
                self.image_width.setValue(self.profile.image.max_width_cm)
                self.stack.setCurrentIndex(1)
            elif index == 1:
                self._apply_rule_edits()
                self._run_background(
                    lambda: self.service.analyze(
                        Path(self.source_edit.text()), self.profile
                    ),
                    self._analysis_ready,
                    "正在分析文档…",
                )
                return
        except Exception as exc:
            QMessageBox.critical(self, "处理失败", str(exc))
            return
        self._sync_navigation()

    def _run_background(self, function, on_result, message: str) -> None:
        self.back_button.setEnabled(False)
        self.next_button.setEnabled(False)
        self.generate_button.setEnabled(False)
        self.status_label.setText(message)
        worker = Worker(function)
        worker.signals.result.connect(on_result)
        worker.signals.error.connect(
            lambda error: QMessageBox.critical(self, "处理失败", error)
        )
        worker.signals.finished.connect(self._background_finished)
        self.thread_pool.start(worker)

    def _background_finished(self) -> None:
        self.status_label.setText("")
        self._sync_navigation()
        if self.stack.currentIndex() == 2:
            self._update_generate_enabled()

    def _analysis_ready(self, analysis: DocumentAnalysis) -> None:
        self.set_analysis(analysis)
        self.stack.setCurrentIndex(2)

    def _generation_ready(self, result) -> None:
        self.result_label.setText(
            f"输出文件：{result.output}\n检查报告：{result.report}\n"
            f"执行操作：{result.operation_count} 项"
        )
        self.stack.setCurrentIndex(3)

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
        self._structure_acknowledged = not analysis.structure_warnings
        self.structure_label.setText("\n".join(analysis.structure_warnings))
        self.ack_structure_button.setVisible(bool(analysis.structure_warnings))
        self.review_model.set_analysis(analysis)

    def _review_confirmation_changed(self, confirmed: bool) -> None:
        self._paragraphs_confirmed = confirmed
        self._update_generate_enabled()

    def _acknowledge_structure_warnings(self) -> None:
        self._structure_acknowledged = True
        self.ack_structure_button.setEnabled(False)
        self._update_generate_enabled()

    def _update_generate_enabled(self) -> None:
        self.generate_button.setEnabled(
            self._paragraphs_confirmed and self._structure_acknowledged
        )

    def _generate(self):
        self._run_background(
            lambda: self.service.format(
                self.analysis,
                self.profile,
                self.review_model.overrides(),
                Path(self.output_edit.text()),
                {
                    field: edit.text().strip()
                    for field, edit in self.cover_edits.items()
                },
            ),
            self._generation_ready,
            "正在调用 Word 生成文档…",
        )
