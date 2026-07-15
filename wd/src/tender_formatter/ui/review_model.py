from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, Signal

from tender_formatter.domain import (
    BlockDecision,
    BlockKind,
    DocumentAnalysis,
    RiskLevel,
)


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
        if not index.isValid() or role not in (Qt.DisplayRole, Qt.EditRole):
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

    def flags(self, index):
        flags = super().flags(index)
        if index.isValid() and index.column() in (2, 3):
            flags |= Qt.ItemIsEditable
        return flags

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid() or role != Qt.EditRole:
            return False
        decision = self._items[index.row()]
        try:
            if index.column() == 2:
                kind = BlockKind(str(value).strip().lower())
                level = decision.level if kind == BlockKind.HEADING else None
                if kind == BlockKind.HEADING and level is None:
                    level = 1
                updated = decision.model_copy(
                    update={"kind": kind, "level": level}
                )
            elif index.column() == 3 and decision.kind == BlockKind.HEADING:
                updated = decision.model_copy(update={"level": int(value)})
                BlockDecision.model_validate(updated.model_dump())
            else:
                return False
        except (ValueError, TypeError):
            return False
        self._items[index.row()] = updated
        self._confirmed.discard(updated.index)
        self.dataChanged.emit(self.index(index.row(), 2), self.index(index.row(), 5))
        self.confirmation_changed.emit(False)
        return True

    def confirm_row(self, row: int) -> None:
        if row < 0 or row >= len(self._items):
            return
        decision = self._items[row]
        self._confirmed.add(decision.index)
        self.dataChanged.emit(self.index(row, 5), self.index(row, 5))
        self.confirmation_changed.emit(len(self._confirmed) == len(self._items))

    def confirm_all_as_suggested(self) -> None:
        for row in range(len(self._items)):
            self.confirm_row(row)

    def overrides(self) -> dict[int, BlockDecision]:
        return {
            decision.index: decision.model_copy(
                update={"confidence": 1.0, "risk": RiskLevel.INFO}
            )
            for decision in self._items
            if decision.index in self._confirmed
        }
