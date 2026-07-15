from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, Signal

from tender_formatter.domain import BlockDecision, DocumentAnalysis, RiskLevel


class ReviewModel(QAbstractTableModel):
    confirmation_changed = Signal(bool)
    _HEADERS = ("风险", "段落", "建议类型", "层级", "识别依据", "确认")

    def __init__(self, parent=None):
        super().__init__(parent)
        self._analysis: DocumentAnalysis | None = None
        self._items: list[BlockDecision] = []
        self._confirmed: set[int] = set()

    def set_analysis(self, analysis: DocumentAnalysis) -> None:
        self.beginResetModel()
        self._analysis = analysis
        self._items = [
            decision
            for decision in analysis.decisions
            if decision.risk in (RiskLevel.REVIEW, RiskLevel.HIGH)
        ]
        self._confirmed.clear()
        self.endResetModel()
        self.confirmation_changed.emit(not self._items)

    def rowCount(self, _parent=QModelIndex()) -> int:
        return len(self._items)

    def columnCount(self, _parent=QModelIndex()) -> int:
        return len(self._HEADERS)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return self._HEADERS[section]
        return None

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or role != Qt.DisplayRole:
            return None
        decision = self._items[index.row()]
        paragraph = (
            self._analysis.paragraphs[decision.index].text
            if self._analysis is not None
            else ""
        )
        values = (
            "高" if decision.risk == RiskLevel.HIGH else "复核",
            paragraph,
            decision.kind.value,
            decision.level or "",
            "；".join(decision.reasons),
            "已确认" if decision.index in self._confirmed else "待确认",
        )
        return values[index.column()]

    def confirm_all_as_suggested(self) -> None:
        self._confirmed = {decision.index for decision in self._items}
        if self._items:
            self.dataChanged.emit(
                self.index(0, 5), self.index(len(self._items) - 1, 5)
            )
        self.confirmation_changed.emit(True)

    def overrides(self) -> dict[int, BlockDecision]:
        return {
            decision.index: decision.model_copy(
                update={"confidence": 1.0, "risk": RiskLevel.INFO}
            )
            for decision in self._items
            if decision.index in self._confirmed
        }
