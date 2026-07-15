# 钢结构技术标书 Word 自动格式化程序 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建一款 Windows 桌面程序，读取企业 Word 样板和项目覆盖规则，识别钢结构技术标书结构，经人工复核疑点后生成不覆盖源文件的规范化 `.docx`。

**Architecture:** 使用 `python-docx` 进行只读 OOXML 分析与可测试的格式操作，使用独立的 `pywin32` 适配器调用 Microsoft Word 更新目录、页码和字段。领域模型、识别规则、格式执行、Word 自动化和 PySide6 界面彼此隔离，先实现命令可测的应用服务，再将其接入四步桌面向导。

**Tech Stack:** Python 3.12、PySide6 6.7+、python-docx 1.1+、pywin32 306+、Pydantic 2.7+、pytest 8+、pytest-qt 4.4+、PyInstaller 6.8+

## Global Constraints

- 仅支持 Windows 10/11、Python 3.12 和已安装 Microsoft Word 的环境。
- 第一版输入仅接受未受保护的 `.docx`；不支持 `.doc`、`.docm` 和损坏文件。
- 永远不覆盖源文件，默认输出名为 `<原文件名>_已格式化.docx`。
- 文档分析阶段严格只读；格式修改只发生在独立输出副本。
- 自动识别一级至三级标题和正文；无编号标题证据不足时必须进入人工复核。
- 第一版覆盖封面、自动目录、页眉页脚、页码、表格、图片与题注。
- 公式深度排版、复杂项目符号和招标文件自然语言规则抽取不在第一版范围。
- Word 自动化只能关闭本程序创建的实例，不能结束用户已有的 Word 进程。

---

## File Structure

```text
pyproject.toml                         # 依赖、pytest、打包入口和工具配置
src/tender_formatter/
  __init__.py                         # 版本号
  domain.py                           # 文档、候选段落、风险、规则及结果模型
  config.py                           # JSON 项目预设和样板样式读取
  analyzer.py                         # DOCX 只读结构提取
  classifier.py                       # 标题/正文/题注规则和置信度
  planner.py                          # 将分析与用户复核结果转换为格式操作计划
  formatter.py                        # 在副本上执行样式、表格、图片和题注规则
  cover.py                            # 封面占位符校验与替换
  word_adapter.py                     # Word COM、分节、目录、页眉页脚及字段更新
  report.py                           # JSON/文本处理报告
  service.py                          # analyze/format 两阶段应用服务
  ui/main_window.py                   # 四步向导和后台任务编排
  ui/review_model.py                  # 风险列表 Qt 模型
  main.py                             # GUI 入口
tests/
  fixtures/                           # 程序生成的最小 DOCX 测试样本
  conftest.py                         # 跨测试共享的 profile、analysis 与 fake Word fixtures
  helpers.py                          # 构造测试 DOCX 和假 Word 适配器
  test_domain.py
  test_config.py
  test_analyzer.py
  test_classifier.py
  test_planner.py
  test_formatter.py
  test_cover.py
  test_word_adapter.py
  test_report.py
  test_service.py
  test_ui.py
scripts/build.ps1                     # PyInstaller 构建和产物检查
```

### Task 1: 项目骨架与领域模型

**Files:**
- Create: `pyproject.toml`
- Create: `src/tender_formatter/__init__.py`
- Create: `src/tender_formatter/domain.py`
- Create: `tests/test_domain.py`

**Interfaces:**
- Produces: `BlockKind`, `RiskLevel`, `ParagraphEvidence`, `BlockDecision`, `PageRules`, `TextStyleRules`, `TableRules`, `ImageRules`, `FormatProfile`, `DocumentAnalysis`, `FormatOperation`, `FormatPlan`, `ProcessingResult`。

- [ ] **Step 1: 创建项目配置并写领域模型失败测试**

```toml
# pyproject.toml
[build-system]
requires = ["setuptools>=70"]
build-backend = "setuptools.build_meta"

[project]
name = "tender-word-formatter"
version = "0.1.0"
requires-python = ">=3.12,<3.13"
dependencies = ["python-docx>=1.1,<2", "pydantic>=2.7,<3", "PySide6>=6.7,<7", "pywin32>=306"]

[project.optional-dependencies]
dev = ["pytest>=8,<9", "pytest-qt>=4.4,<5", "pyinstaller>=6.8,<7"]

[project.gui-scripts]
tender-word-formatter = "tender_formatter.main:main"

[tool.setuptools.packages.find]
where = ["src"]

[tool.pytest.ini_options]
pythonpath = ["src"]
testpaths = ["tests"]
```

```python
# tests/test_domain.py
from tender_formatter.domain import BlockDecision, BlockKind, FormatProfile

def test_profile_rejects_invalid_confidence_thresholds():
    try:
        FormatProfile(name="公司标准", high_confidence=0.6, review_confidence=0.8)
    except ValueError as exc:
        assert "high_confidence" in str(exc)
    else:
        raise AssertionError("expected validation error")

def test_decision_requires_level_only_for_heading():
    decision = BlockDecision(index=3, kind=BlockKind.BODY, confidence=0.95)
    assert decision.level is None
```

- [ ] **Step 2: 运行测试确认因模块不存在而失败**

Run: `python -m pytest tests/test_domain.py -v`
Expected: FAIL，包含 `ModuleNotFoundError: No module named 'tender_formatter'`。

- [ ] **Step 3: 实现最小领域模型**

```python
# src/tender_formatter/domain.py
from enum import StrEnum
from pathlib import Path
from pydantic import BaseModel, Field, model_validator

class BlockKind(StrEnum):
    HEADING = "heading"
    BODY = "body"
    CAPTION = "caption"
    EMPTY = "empty"

class RiskLevel(StrEnum):
    INFO = "info"
    REVIEW = "review"
    HIGH = "high"

class ParagraphEvidence(BaseModel):
    text: str
    style_name: str = ""
    outline_level: int | None = None
    numbering_level: int | None = None
    bold_ratio: float = 0.0
    alignment: str | None = None

class BlockDecision(BaseModel):
    index: int
    kind: BlockKind
    level: int | None = None
    confidence: float = Field(ge=0, le=1)
    reasons: list[str] = []
    risk: RiskLevel = RiskLevel.INFO

    @model_validator(mode="after")
    def validate_level(self):
        if self.kind == BlockKind.HEADING and self.level not in (1, 2, 3):
            raise ValueError("heading level must be 1, 2, or 3")
        if self.kind != BlockKind.HEADING and self.level is not None:
            raise ValueError("level is only valid for heading")
        return self

class TextStyleRules(BaseModel):
    east_asia_font: str = "宋体"
    latin_font: str = "Times New Roman"
    size_pt: float = 12
    bold: bool = False
    alignment: str = "justify"
    line_spacing: float = 1.5
    space_before_pt: float = 0
    space_after_pt: float = 0

class PageRules(BaseModel):
    top_cm: float = 2.5
    bottom_cm: float = 2.5
    left_cm: float = 3.0
    right_cm: float = 2.5
    gutter_cm: float = 0
    body_page_start: int = 1

class TableRules(BaseModel):
    width_percent: int = Field(default=100, ge=10, le=100)
    repeat_header: bool = True
    alignment: str = "center"

class ImageRules(BaseModel):
    max_width_cm: float = Field(default=15.5, gt=0)
    alignment: str = "center"
    caption_label: str = "图"

class FormatProfile(BaseModel):
    name: str
    template_path: Path | None = None
    high_confidence: float = Field(default=0.9, ge=0, le=1)
    review_confidence: float = Field(default=0.65, ge=0, le=1)
    body: TextStyleRules = TextStyleRules()
    headings: dict[int, TextStyleRules] = {}
    page: PageRules = PageRules()
    table: TableRules = TableRules()
    image: ImageRules = ImageRules()

    @model_validator(mode="after")
    def validate_thresholds(self):
        if self.high_confidence <= self.review_confidence:
            raise ValueError("high_confidence must exceed review_confidence")
        return self

class DocumentAnalysis(BaseModel):
    source: Path
    paragraphs: list[ParagraphEvidence]
    decisions: list[BlockDecision]
    table_count: int = 0
    image_count: int = 0
    section_count: int = 0

class FormatOperation(BaseModel):
    kind: str
    target: int | str
    parameters: dict = {}

class FormatPlan(BaseModel):
    source: Path
    output: Path
    operations: list[FormatOperation]
    warnings: list[str] = []

class ProcessingResult(BaseModel):
    output: Path
    report: Path
    operation_count: int
    warnings: list[str] = []
```

- [ ] **Step 4: 运行领域测试**

Run: `python -m pytest tests/test_domain.py -v`
Expected: 2 passed。

- [ ] **Step 5: 提交项目骨架**

```powershell
git add pyproject.toml src/tender_formatter/__init__.py src/tender_formatter/domain.py tests/test_domain.py
git commit -m "feat: add formatter domain model"
```

### Task 2: 项目预设与样板样式读取

**Files:**
- Create: `src/tender_formatter/config.py`
- Create: `tests/helpers.py`
- Create: `tests/test_config.py`

**Interfaces:**
- Consumes: `FormatProfile`, `TextStyleRules`。
- Produces: `save_profile(profile: FormatProfile, path: Path) -> None`、`load_profile(path: Path) -> FormatProfile`、`read_template_profile(path: Path, name: str) -> FormatProfile`。

- [ ] **Step 1: 写 JSON 往返和样板读取测试**

```python
# tests/test_config.py
from tender_formatter.config import load_profile, read_template_profile, save_profile
from tender_formatter.domain import FormatProfile
from tests.helpers import make_docx_with_styles

def test_profile_json_round_trip(tmp_path):
    path = tmp_path / "company.json"
    save_profile(FormatProfile(name="公司标准"), path)
    assert load_profile(path).name == "公司标准"

def test_read_template_extracts_normal_and_heading_styles(tmp_path):
    template = make_docx_with_styles(tmp_path / "template.docx")
    profile = read_template_profile(template, "样板")
    assert profile.body.east_asia_font == "宋体"
    assert profile.headings[1].bold is True
```

- [ ] **Step 2: 运行测试确认缺少配置接口**

Run: `python -m pytest tests/test_config.py -v`
Expected: FAIL，包含 `ModuleNotFoundError: tender_formatter.config`。

- [ ] **Step 3: 实现安全保存和样式读取**

```python
# src/tender_formatter/config.py
import json
from pathlib import Path
from docx import Document
from tender_formatter.domain import FormatProfile, TextStyleRules

def save_profile(profile: FormatProfile, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(profile.model_dump_json(indent=2), encoding="utf-8")
    temp.replace(path)

def load_profile(path: Path) -> FormatProfile:
    return FormatProfile.model_validate_json(path.read_text(encoding="utf-8"))

def _style_rules(style) -> TextStyleRules:
    font = style.font
    east_asia = font.name or "宋体"
    size = font.size.pt if font.size else 12
    return TextStyleRules(east_asia_font=east_asia, size_pt=size, bold=bool(font.bold))

def read_template_profile(path: Path, name: str) -> FormatProfile:
    document = Document(path)
    headings = {level: _style_rules(document.styles[f"Heading {level}"]) for level in (1, 2, 3)}
    return FormatProfile(
        name=name,
        template_path=path,
        body=_style_rules(document.styles["Normal"]),
        headings=headings,
    )
```

`tests/helpers.py` 中实现 `make_docx_with_styles(path)`：创建文档，将 `Normal` 设置为宋体，将 `Heading 1` 设置为加粗，保存并返回路径。

- [ ] **Step 4: 运行配置测试**

Run: `python -m pytest tests/test_config.py -v`
Expected: 2 passed。

- [ ] **Step 5: 提交预设能力**

```powershell
git add src/tender_formatter/config.py tests/helpers.py tests/test_config.py
git commit -m "feat: add template profile management"
```

### Task 3: DOCX 只读分析器

**Files:**
- Create: `src/tender_formatter/analyzer.py`
- Create: `tests/test_analyzer.py`

**Interfaces:**
- Consumes: `.docx` 路径。
- Produces: `analyze_docx(path: Path) -> DocumentAnalysis`，其中段落顺序与 `document.paragraphs` 一致，并包含样式、大纲级别、编号级别、加粗比例、表格数、图片数和分节数。

- [ ] **Step 1: 写结构提取和非法输入测试**

```python
from tender_formatter.analyzer import analyze_docx
from tests.helpers import make_mixed_docx

def test_analyzer_extracts_structure_without_modifying_source(tmp_path):
    source = make_mixed_docx(tmp_path / "mixed.docx")
    before = source.read_bytes()
    result = analyze_docx(source)
    assert [p.text for p in result.paragraphs[:3]] == ["第一章 总体方案", "1.1 编制依据", "正文内容"]
    assert result.table_count == 1
    assert result.image_count == 1
    assert source.read_bytes() == before

def test_analyzer_rejects_non_docx(tmp_path):
    path = tmp_path / "input.doc"
    path.write_bytes(b"not a docx")
    try:
        analyze_docx(path)
    except ValueError as exc:
        assert "仅支持 .docx" in str(exc)
```

- [ ] **Step 2: 运行测试确认失败**

Run: `python -m pytest tests/test_analyzer.py -v`
Expected: FAIL，缺少 `analyze_docx`。

- [ ] **Step 3: 实现 OOXML 属性提取**

在 `analyzer.py` 中使用 `Document(path)`；通过 `paragraph._p.pPr.numPr.ilvl` 读取编号层级，通过 `paragraph._p.pPr.outlineLvl` 读取大纲层级；按非空 run 字符数计算 `bold_ratio`；使用 `//a:blip` 统计图片引用。捕获 `PackageNotFoundError` 和 `BadZipFile` 并转换为包含文件名的 `ValueError`。不得保存文档。

```python
def analyze_docx(path: Path) -> DocumentAnalysis:
    if path.suffix.lower() != ".docx":
        raise ValueError("第一版仅支持 .docx 文件")
    document = Document(path)
    paragraphs = [_extract_paragraph(p) for p in document.paragraphs]
    image_count = len(document.part.element.xpath(".//a:blip"))
    return DocumentAnalysis(source=path, paragraphs=paragraphs, decisions=[],
                            table_count=len(document.tables), image_count=image_count,
                            section_count=len(document.sections))
```

- [ ] **Step 4: 运行分析器测试**

Run: `python -m pytest tests/test_analyzer.py -v`
Expected: 2 passed。

- [ ] **Step 5: 提交分析器**

```powershell
git add src/tender_formatter/analyzer.py tests/test_analyzer.py tests/helpers.py
git commit -m "feat: analyze docx structure read only"
```

### Task 4: 标题分类、置信度和编号风险

**Files:**
- Create: `src/tender_formatter/classifier.py`
- Create: `tests/test_classifier.py`

**Interfaces:**
- Consumes: `list[ParagraphEvidence]`, `FormatProfile`。
- Produces: `classify_paragraphs(paragraphs, profile) -> list[BlockDecision]`。

- [ ] **Step 1: 写常见编号、自动编号、无编号和跳级测试**

```python
from tender_formatter.classifier import classify_paragraphs
from tender_formatter.domain import FormatProfile, ParagraphEvidence, RiskLevel

def classify(*items):
    return classify_paragraphs(list(items), FormatProfile(name="test"))

def test_chinese_chapter_and_decimal_levels_are_recognized():
    result = classify(
        ParagraphEvidence(text="第一章 总体施工部署"),
        ParagraphEvidence(text="1.1 编制依据"),
        ParagraphEvidence(text="1.1.1 规范标准"),
    )
    assert [(x.kind.value, x.level) for x in result] == [("heading", 1), ("heading", 2), ("heading", 3)]

def test_word_numbering_and_heading_style_strengthen_evidence():
    result = classify(ParagraphEvidence(text="施工部署", style_name="Heading 1", numbering_level=0))
    assert result[0].level == 1
    assert result[0].confidence >= 0.9

def test_unnumbered_bold_short_paragraph_requires_review():
    result = classify(ParagraphEvidence(text="施工部署", bold_ratio=1.0))
    assert result[0].risk == RiskLevel.REVIEW

def test_heading_level_jump_is_high_risk():
    result = classify(ParagraphEvidence(text="1.1 内容"), ParagraphEvidence(text="1.1.1 内容"))
    assert result[0].risk == RiskLevel.HIGH
```

- [ ] **Step 2: 运行测试确认失败**

Run: `python -m pytest tests/test_classifier.py -v`
Expected: FAIL，缺少分类器。

- [ ] **Step 3: 实现确定性评分规则**

使用编译后的正则：`^第[一二三四五六七八九十百]+章\s*` 为一级；`^\d+[、.]\s*\S` 为一级候选；`^\d+\.\d+\s+\S` 为二级；`^\d+\.\d+\.\d+\s+\S` 为三级。编号命中基础分 0.85，匹配 Heading 样式或大纲级别加 0.15，Word 编号层级匹配加 0.15，短且加粗加 0.05，最高为 1.0。仅有短且加粗证据时置信度 0.7 并进入复核。其余非空段落为正文 0.95，空段落为 `EMPTY`。检查相邻标题层级，缺少父级时标记 `HIGH` 并添加中文原因。

- [ ] **Step 4: 运行分类测试**

Run: `python -m pytest tests/test_classifier.py -v`
Expected: 4 passed。

- [ ] **Step 5: 提交分类器**

```powershell
git add src/tender_formatter/classifier.py tests/test_classifier.py
git commit -m "feat: classify tender document headings"
```

### Task 5: 用户复核结果与格式操作计划

**Files:**
- Create: `src/tender_formatter/planner.py`
- Create: `tests/conftest.py`
- Create: `tests/test_planner.py`

**Interfaces:**
- Consumes: `DocumentAnalysis`, `FormatProfile`, `dict[int, BlockDecision]` 用户覆盖。
- Produces: `build_plan(analysis, profile, overrides, output) -> FormatPlan`。

- [ ] **Step 1: 写输出保护、复核门禁和操作生成测试**

```python
def test_plan_refuses_source_as_output(analysis, profile):
    with pytest.raises(ValueError, match="不能覆盖源文件"):
        build_plan(analysis, profile, {}, analysis.source)

def test_plan_requires_resolution_for_review_items(analysis_with_review, profile, tmp_path):
    with pytest.raises(ValueError, match="尚有未确认"):
        build_plan(analysis_with_review, profile, {}, tmp_path / "out.docx")

def test_plan_contains_heading_body_table_image_and_word_operations(confirmed_analysis, profile, tmp_path):
    plan = build_plan(confirmed_analysis, profile, {}, tmp_path / "out.docx")
    kinds = {operation.kind for operation in plan.operations}
    assert {"apply_paragraph", "format_tables", "format_images", "replace_cover", "word_finalize"} <= kinds
```

- [ ] **Step 2: 运行测试确认失败**

Run: `python -m pytest tests/test_planner.py -v`
Expected: FAIL，缺少 `build_plan`。

- [ ] **Step 3: 实现纯函数计划器**

合并用户覆盖后，拒绝仍为 `REVIEW/HIGH` 且未确认的决策；为每个非空段落生成 `apply_paragraph`；始终生成页面、表格、图片、封面和 `word_finalize` 操作。使用 `source.resolve() == output.resolve()` 阻止覆盖。输出不存在时采用 `<stem>_已格式化.docx`。

- [ ] **Step 4: 运行计划器测试**

Run: `python -m pytest tests/test_planner.py -v`
Expected: 3 passed。

- [ ] **Step 5: 提交操作计划器**

```powershell
git add src/tender_formatter/planner.py tests/test_planner.py tests/conftest.py
git commit -m "feat: build reviewed formatting plans"
```

### Task 6: DOCX 格式执行器（正文、标题、表格、图片、题注）

**Files:**
- Create: `src/tender_formatter/formatter.py`
- Create: `tests/test_formatter.py`

**Interfaces:**
- Consumes: `FormatPlan`, `FormatProfile`。
- Produces: `execute_docx_plan(plan, profile) -> Path`；仅执行非 COM 操作并返回输出路径。

- [ ] **Step 1: 写副本保护和格式结果测试**

```python
def test_formatter_changes_copy_and_preserves_source(sample_plan, profile):
    before = sample_plan.source.read_bytes()
    output = execute_docx_plan(sample_plan, profile)
    assert sample_plan.source.read_bytes() == before
    document = Document(output)
    assert document.paragraphs[0].style.name == "Heading 1"
    assert document.paragraphs[2].style.name == "Normal"

def test_formatter_repeats_table_header_and_scales_wide_image(sample_plan, profile):
    output = execute_docx_plan(sample_plan, profile)
    document = Document(output)
    assert document.tables[0].rows[0]._tr.get_or_add_trPr().first_child_found_in("w:tblHeader") is not None
    assert max(shape.width.cm for shape in document.inline_shapes) <= profile.image.max_width_cm
```

- [ ] **Step 2: 运行测试确认失败**

Run: `python -m pytest tests/test_formatter.py -v`
Expected: FAIL，缺少执行器。

- [ ] **Step 3: 实现最小安全格式执行**

先用 `shutil.copy2(plan.source, plan.output)` 建立副本。按段落索引应用 `Heading 1..3` 或 `Normal` 样式，并在样式层统一字体、字号、对齐、行距和段距；保留 run 的 `bold`、`subscript`、`superscript` 和超链接结构。表格设置居中、首行 `w:tblHeader`；图片仅在超过最大宽度时按比例缩小。已有以“图 + 数字”开头的段落应用 `Caption` 样式，缺失题注只写入警告，不生成图名。保存输出副本。

- [ ] **Step 4: 运行格式执行器测试**

Run: `python -m pytest tests/test_formatter.py -v`
Expected: 2 passed。

- [ ] **Step 5: 提交 DOCX 格式能力**

```powershell
git add src/tender_formatter/formatter.py tests/test_formatter.py
git commit -m "feat: format document content safely"
```

### Task 7: 封面字段处理

**Files:**
- Create: `src/tender_formatter/cover.py`
- Create: `tests/test_cover.py`

**Interfaces:**
- Produces: `find_cover_fields(template: Path) -> set[str]`、`replace_cover_fields(document, values: dict[str, str]) -> list[str]`。

- [ ] **Step 1: 写占位符检测、替换和缺值测试**

```python
def test_cover_fields_are_detected_and_replaced(tmp_path):
    path = make_cover_docx(tmp_path / "cover.docx", "{{项目名称}}\n{{投标单位}}\n{{日期}}")
    assert find_cover_fields(path) == {"项目名称", "投标单位", "日期"}
    document = Document(path)
    missing = replace_cover_fields(document, {"项目名称": "厂房项目", "投标单位": "建设公司", "日期": "2026年7月"})
    assert missing == []
    assert "{{" not in "\n".join(p.text for p in document.paragraphs)

def test_missing_cover_value_is_reported(tmp_path):
    document = Document(make_cover_docx(tmp_path / "cover.docx", "{{项目名称}}\n{{日期}}"))
    assert replace_cover_fields(document, {"项目名称": "厂房项目"}) == ["日期"]
```

- [ ] **Step 2: 运行测试确认失败**

Run: `python -m pytest tests/test_cover.py -v`
Expected: FAIL，缺少封面模块。

- [ ] **Step 3: 实现明确占位符替换**

使用正则 `\{\{([^{}]+)\}\}`。遍历正文段落及表格单元格段落；仅替换显式占位符，不猜测普通文字。若占位符跨多个 run，则在保留段落样式的前提下重建该段的文本 run。返回排序后的缺值字段。

- [ ] **Step 4: 运行封面测试**

Run: `python -m pytest tests/test_cover.py -v`
Expected: 2 passed。

- [ ] **Step 5: 提交封面处理**

```powershell
git add src/tender_formatter/cover.py tests/test_cover.py tests/helpers.py
git commit -m "feat: replace explicit cover fields"
```

### Task 8: Microsoft Word 自动化适配器

**Files:**
- Create: `src/tender_formatter/word_adapter.py`
- Create: `tests/test_word_adapter.py`

**Interfaces:**
- Produces: `WordSettings`、`WordAdapter.finalize(path: Path, settings: WordSettings) -> None`。

- [ ] **Step 1: 用假 COM 对象写生命周期和字段更新测试**

```python
def test_finalize_uses_private_word_instance_and_always_quits(tmp_path):
    app = FakeWordApplication()
    adapter = WordAdapter(dispatch_ex=lambda name: app)
    adapter.finalize(tmp_path / "out.docx", WordSettings(update_toc=True, body_page_start=1))
    assert app.visible is False
    assert app.opened_document.fields_updated
    assert app.opened_document.toc_updated
    assert app.opened_document.saved
    assert app.quit_called

def test_finalize_quits_private_instance_on_failure(tmp_path):
    app = FakeWordApplication(fail_on="Fields.Update")
    with pytest.raises(WordAutomationError):
        WordAdapter(dispatch_ex=lambda name: app).finalize(tmp_path / "out.docx", WordSettings())
    assert app.quit_called
```

- [ ] **Step 2: 运行测试确认失败**

Run: `python -m pytest tests/test_word_adapter.py -v`
Expected: FAIL，缺少适配器。

- [ ] **Step 3: 实现独立 Word 实例和 finally 清理**

使用 `win32com.client.DispatchEx("Word.Application")` 而不是 `Dispatch`；设置 `Visible=False`、`DisplayAlerts=0`。打开输出文档，按配置处理 `Sections` 的首页不同、奇偶页不同、页眉页脚继承及正文页码起始值；更新 `TablesOfContents`、`Fields` 和 `Repaginate()`，保存后关闭文档。所有路径在调用前转为绝对路径。`finally` 中先关闭本次文档，再调用本实例的 `Quit()`；异常统一转换为 `WordAutomationError`。

- [ ] **Step 4: 运行适配器单元测试**

Run: `python -m pytest tests/test_word_adapter.py -v`
Expected: 2 passed；该测试不启动真实 Word。

- [ ] **Step 5: 在安装 Word 的 Windows 机器运行标记集成测试**

Run: `python -m pytest tests/test_word_adapter.py -m word_integration -v`
Expected: 生成临时 DOCX、Word 打开并更新字段、测试 PASS；没有 Word 时 SKIP 并说明原因。

- [ ] **Step 6: 提交 Word 自动化层**

```powershell
git add src/tender_formatter/word_adapter.py tests/test_word_adapter.py
git commit -m "feat: finalize documents with Microsoft Word"
```

### Task 9: 报告与端到端应用服务

**Files:**
- Create: `src/tender_formatter/report.py`
- Create: `src/tender_formatter/service.py`
- Create: `tests/test_report.py`
- Create: `tests/test_service.py`

**Interfaces:**
- Produces: `FormatterService.analyze(source, profile) -> DocumentAnalysis`、`FormatterService.format(analysis, profile, overrides, output, cover_values) -> ProcessingResult`。

- [ ] **Step 1: 写报告脱敏和端到端编排测试**

```python
def test_report_contains_counts_but_not_body_text(tmp_path):
    report = write_report(tmp_path / "report.json", operation_count=8,
                          warnings=["编号跳级"], paragraph_count=100)
    data = json.loads(report.read_text(encoding="utf-8"))
    assert data["operation_count"] == 8
    assert "document_text" not in data

def test_service_analyze_then_format_calls_word_and_preserves_source(tmp_path, fake_word):
    source = make_mixed_docx(tmp_path / "input.docx")
    before = source.read_bytes()
    service = FormatterService(word=fake_word)
    analysis = service.analyze(source, FormatProfile(name="公司标准"))
    overrides = confirm_all_review_items(analysis)
    result = service.format(analysis, FormatProfile(name="公司标准"), overrides,
                            tmp_path / "input_已格式化.docx", {})
    assert result.output.exists() and result.report.exists()
    assert source.read_bytes() == before
    assert fake_word.finalized == [result.output]
```

- [ ] **Step 2: 运行测试确认失败**

Run: `python -m pytest tests/test_report.py tests/test_service.py -v`
Expected: FAIL，缺少报告和服务模块。

- [ ] **Step 3: 实现两阶段服务**

`analyze()` 调用分析器和分类器并返回含 decisions 的新 `DocumentAnalysis`。`format()` 校验复核项和封面字段，构建计划、执行 DOCX 格式、调用 Word 适配器、重新用 `Document(output)` 验证可打开性，最后原子写入 UTF-8 JSON 报告。失败时保留临时输出和错误日志，但不返回成功结果。

- [ ] **Step 4: 运行服务测试**

Run: `python -m pytest tests/test_report.py tests/test_service.py -v`
Expected: 2 passed。

- [ ] **Step 5: 提交应用服务**

```powershell
git add src/tender_formatter/report.py src/tender_formatter/service.py tests/test_report.py tests/test_service.py
git commit -m "feat: orchestrate document formatting workflow"
```

### Task 10: 四步桌面向导

**Files:**
- Create: `src/tender_formatter/ui/__init__.py`
- Create: `src/tender_formatter/ui/review_model.py`
- Create: `src/tender_formatter/ui/main_window.py`
- Create: `src/tender_formatter/main.py`
- Create: `tests/test_ui.py`

**Interfaces:**
- Consumes: `FormatterService`。
- Produces: `MainWindow(service)`，包含选择文件、格式规则、检查确认和生成结果四页。

- [ ] **Step 1: 写向导门禁和复核模型测试**

```python
def test_next_is_disabled_until_source_and_template_exist(qtbot, fake_service, tmp_path):
    window = MainWindow(fake_service)
    qtbot.addWidget(window)
    assert not window.next_button.isEnabled()
    window.source_edit.setText(str(tmp_path / "missing.docx"))
    assert not window.next_button.isEnabled()

def test_review_page_lists_only_review_and_high_risks(qtbot, fake_service):
    window = MainWindow(fake_service)
    qtbot.addWidget(window)
    window.set_analysis(analysis_with_three_risk_levels())
    assert window.review_model.rowCount() == 2

def test_generate_is_blocked_until_all_risks_confirmed(qtbot, configured_window):
    configured_window.go_to_review()
    assert not configured_window.generate_button.isEnabled()
    configured_window.review_model.confirm_all_as_suggested()
    assert configured_window.generate_button.isEnabled()
```

- [ ] **Step 2: 运行测试确认失败**

Run: `python -m pytest tests/test_ui.py -v`
Expected: FAIL，缺少 UI 模块。

- [ ] **Step 3: 实现最小四步向导**

使用 `QStackedWidget` 构建四页。文件页提供拖放区、源文件、样板和输出路径；规则页编辑页面、标题正文、表格和图片规则；复核页使用 `QAbstractTableModel` 显示风险等级、段落文本、上下文、建议类型、层级和原因，并允许逐项确认；结果页显示进度、输出路径、报告路径和警告。分析与生成通过 `QThreadPool` 执行，后台线程只传递不可变 Pydantic 模型，所有控件更新通过 Qt signal 回到主线程。

- [ ] **Step 4: 运行 UI 测试**

Run: `$env:QT_QPA_PLATFORM='offscreen'; python -m pytest tests/test_ui.py -v`
Expected: 3 passed。

- [ ] **Step 5: 手工启动烟雾测试**

Run: `python -m tender_formatter.main`
Expected: 窗口正常打开；四步标题可见；未选择有效源文件和样板时不能进入分析。

- [ ] **Step 6: 提交桌面界面**

```powershell
git add src/tender_formatter/ui src/tender_formatter/main.py tests/test_ui.py
git commit -m "feat: add four-step desktop workflow"
```

### Task 11: 打包、真实样本验证与交付文档

**Files:**
- Create: `scripts/build.ps1`
- Create: `README.md`
- Modify: `pyproject.toml`
- Create: `tests/test_package_smoke.py`

**Interfaces:**
- Produces: `dist/TenderWordFormatter/TenderWordFormatter.exe` 和操作说明。

- [ ] **Step 1: 写入口和版本烟雾测试**

```python
def test_package_entry_imports():
    from tender_formatter.main import main
    from tender_formatter import __version__
    assert callable(main)
    assert __version__ == "0.1.0"
```

- [ ] **Step 2: 运行完整自动化测试**

Run: `$env:QT_QPA_PLATFORM='offscreen'; python -m pytest -m 'not word_integration' -v`
Expected: 所有非 Word 集成测试 PASS，无 warning 被配置为 error 后漏出。

- [ ] **Step 3: 添加构建脚本和使用文档**

`scripts/build.ps1` 必须先运行测试，再执行：

```powershell
python -m PyInstaller --noconfirm --clean --windowed --name TenderWordFormatter --paths src --collect-all win32com src/tender_formatter/main.py
if (-not (Test-Path 'dist/TenderWordFormatter/TenderWordFormatter.exe')) { throw '打包产物不存在' }
```

`README.md` 写明 Windows/Word 前置条件、安装、启动、样板 Heading 1-3/Normal/Caption 样式要求、封面 `{{项目名称}}` 等占位符规则、四步操作流程、输出与报告位置、源文件保护策略和常见错误处理。

- [ ] **Step 4: 构建可执行程序**

Run: `powershell -ExecutionPolicy Bypass -File scripts/build.ps1`
Expected: 测试 PASS，生成 `dist/TenderWordFormatter/TenderWordFormatter.exe`。

- [ ] **Step 5: 用脱敏真实标书做验收矩阵**

依次测试至少三份样本：多来源复制标书、自动编号混用标书、含横向/跨页表格和浮动图片的标书。每份记录以下结果：源文件 SHA-256 前后相同；输出可由 Word 打开并再次保存；一级至三级标题抽查正确；目录已更新；正文页码起算正确；表格和图片数量未减少；高风险项均在生成前出现。失败项必须转为可复现自动测试后修复。

- [ ] **Step 6: Word 导出 PDF 视觉核对**

对每份验收样本的封面、目录、正文首页、横向表格页、跨页表格页和典型图片页导出 PDF。逐项确认字体、字号、行距、边距、页眉页脚、页码、表格宽度、图片比例和题注位置符合所选预设。

- [ ] **Step 7: 提交交付材料**

```powershell
git add pyproject.toml README.md scripts/build.ps1 tests/test_package_smoke.py
git commit -m "build: package tender word formatter"
```

## Final Verification

- [ ] Run: `python -m pytest -v` — Expected: 所有可用环境测试 PASS；仅未安装 Word 时的集成测试可明确 SKIP。
- [ ] Run: `git diff --check` — Expected: 无输出。
- [ ] Run: `git status --short` — Expected: 只出现用户原有且与本项目无关的改动，计划内文件均已提交。
- [ ] 启动打包程序，完整执行一次“选择—配置—复核—生成”，确认原文件未改变且输出与报告均可打开。
