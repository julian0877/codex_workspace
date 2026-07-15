from pathlib import Path

from tender_formatter.domain import (
    BlockDecision,
    BlockKind,
    DocumentAnalysis,
    FormatOperation,
    FormatPlan,
    FormatProfile,
    RiskLevel,
)


def default_output_path(source: Path) -> Path:
    return source.with_name(f"{source.stem}_已格式化.docx")


def build_plan(
    analysis: DocumentAnalysis,
    profile: FormatProfile,
    overrides: dict[int, BlockDecision],
    output: Path | None = None,
) -> FormatPlan:
    output = output or default_output_path(analysis.source)
    if analysis.source.resolve() == output.resolve():
        raise ValueError("输出文件不能覆盖源文件")

    unresolved = [
        decision.index
        for decision in analysis.decisions
        if decision.risk in (RiskLevel.REVIEW, RiskLevel.HIGH)
        and decision.index not in overrides
    ]
    if unresolved:
        indexes = "、".join(str(index + 1) for index in unresolved)
        raise ValueError(f"尚有未确认的风险段落：{indexes}")

    decisions = {
        decision.index: overrides.get(decision.index, decision)
        for decision in analysis.decisions
    }
    operations: list[FormatOperation] = []
    for index, decision in decisions.items():
        if decision.kind == BlockKind.EMPTY:
            continue
        operations.append(
            FormatOperation(
                kind="apply_paragraph",
                target=index,
                parameters={
                    "kind": decision.kind.value,
                    "level": decision.level,
                },
            )
        )
    operations.extend(
        [
            FormatOperation(
                kind="apply_page",
                target="all",
                parameters=profile.page.model_dump(),
            ),
            FormatOperation(
                kind="format_tables",
                target="all",
                parameters={
                    **profile.table.model_dump(),
                    "risky_indexes": analysis.risky_table_indexes,
                },
            ),
            FormatOperation(
                kind="format_images",
                target="all",
                parameters=profile.image.model_dump(),
            ),
            FormatOperation(kind="replace_cover", target="cover"),
            FormatOperation(kind="word_finalize", target="document"),
        ]
    )
    return FormatPlan(
        source=analysis.source,
        output=output,
        operations=operations,
    )
