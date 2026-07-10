import math
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


@dataclass
class CheckResult:
    name: str
    actual: float
    allowable: float
    unit: str
    expression: str
    substitution: str = ""

    @property
    def ratio(self) -> float:
        if self.allowable == 0:
            return float("inf")
        return self.actual / self.allowable

    @property
    def passed(self) -> bool:
        return self.actual <= self.allowable


MATERIAL_STRENGTHS = {
    "Q235": [
        (16, 215, 125, 320),
        (40, 205, 120, 320),
        (100, 200, 115, 320),
    ],
    "Q345": [
        (16, 305, 175, 400),
        (40, 295, 170, 400),
        (63, 290, 165, 400),
        (80, 280, 160, 400),
        (100, 270, 155, 400),
    ],
    "Q390": [
        (16, 345, 200, 415),
        (40, 330, 190, 415),
        (63, 310, 180, 415),
        (100, 295, 170, 415),
    ],
    "Q420": [
        (16, 375, 215, 440),
        (40, 355, 205, 440),
        (63, 320, 185, 440),
        (100, 305, 175, 440),
    ],
    "40Cr": [
        (100, 365, 210, None),
        (9999, 340, 195, None),
    ],
    "35CrMo": [
        (100, 365, 210, None),
        (9999, 340, 195, None),
    ],
}


TEMPLATES = {
    "5t": {
        "weight": 5,
        "safety_factor": 1.32,
        "lift_points": 2,
        "lug_plates_per_point": 2,
        "angle_deg": 45,
        "pull_force": 44,
        "material": "Q345",
        "plate_thickness": 10,
        "cover_plate_thickness": 0,
        "plate_width": 96,
        "root_gap": 5,
        "plate_height": 88,
        "hole_diameter": 30,
        "weld_beta": 1,
    },
    "10t": {
        "weight": 10,
        "safety_factor": 1.32,
        "lift_points": 2,
        "lug_plates_per_point": 2,
        "angle_deg": 45,
        "pull_force": 0,
        "material": "Q345",
        "plate_thickness": 14,
        "cover_plate_thickness": 0,
        "plate_width": 130,
        "root_gap": 5,
        "plate_height": 120,
        "hole_diameter": 40,
        "weld_beta": 1,
    },
    "15t": {
        "weight": 15,
        "safety_factor": 1.32,
        "lift_points": 2,
        "lug_plates_per_point": 2,
        "angle_deg": 45,
        "pull_force": 0,
        "material": "Q345",
        "plate_thickness": 18,
        "cover_plate_thickness": 0,
        "plate_width": 160,
        "root_gap": 5,
        "plate_height": 150,
        "hole_diameter": 50,
        "weld_beta": 1,
    },
    "30t": {
        "weight": 30,
        "safety_factor": 1.32,
        "lift_points": 2,
        "lug_plates_per_point": 2,
        "angle_deg": 45,
        "pull_force": 0,
        "material": "Q345",
        "plate_thickness": 28,
        "cover_plate_thickness": 0,
        "plate_width": 230,
        "root_gap": 0,
        "plate_height": 210,
        "hole_diameter": 75,
        "weld_beta": 1,
    },
    "40t": {
        "weight": 40,
        "safety_factor": 1.32,
        "lift_points": 2,
        "lug_plates_per_point": 2,
        "angle_deg": 45,
        "pull_force": 0,
        "material": "Q345",
        "plate_thickness": 34,
        "cover_plate_thickness": 0,
        "plate_width": 270,
        "root_gap": 0,
        "plate_height": 250,
        "hole_diameter": 90,
        "weld_beta": 1,
    },
    "50t": {
        "weight": 50,
        "safety_factor": 1.32,
        "lift_points": 2,
        "lug_plates_per_point": 2,
        "angle_deg": 45,
        "pull_force": 0,
        "material": "Q345",
        "plate_thickness": 40,
        "cover_plate_thickness": 0,
        "plate_width": 310,
        "root_gap": 0,
        "plate_height": 290,
        "hole_diameter": 100,
        "weld_beta": 1,
    },
}


def material_strength(material, thickness):
    if material not in MATERIAL_STRENGTHS:
        raise ValueError(f"暂不支持材料：{material}")
    for limit, tensile, shear, bearing in MATERIAL_STRENGTHS[material]:
        if thickness <= limit:
            return {
                "tensile": tensile,
                "shear": shear,
                "bearing": bearing if bearing else 385,
                "weld": 160 if material == "Q235" else 200,
            }
    raise ValueError(f"{material} 厚度超过材料强度表范围，请核对。")


def num(data, key):
    return float(data[key])


def calculate(data):
    weight = num(data, "weight")
    safety_factor = num(data, "safety_factor")
    lift_points = num(data, "lift_points")
    lug_plates_per_point = num(data, "lug_plates_per_point")
    angle_deg = num(data, "angle_deg")
    angle = angle_deg / 180 * math.pi
    t = num(data, "plate_thickness")
    t1 = num(data, "cover_plate_thickness")
    b_width = num(data, "plate_width")
    gap = num(data, "root_gap")
    height = num(data, "plate_height")
    hole_diameter = num(data, "hole_diameter")
    beta = num(data, "weld_beta")
    material = data["material"]

    if lift_points <= 0 or lug_plates_per_point <= 0:
        raise ValueError("吊点数量和每吊点耳板片数必须大于 0。")

    if angle <= 0 or math.sin(angle) <= 0:
        raise ValueError("吊装角度必须大于 0 度。")

    pull_force = num(data, "pull_force")
    if pull_force <= 0:
        pull_force = weight * 9.8 * safety_factor / lift_points / math.sin(angle)

    vertical_force = pull_force * math.sin(angle)
    horizontal_force = pull_force * math.cos(angle)
    plate_force = pull_force / lug_plates_per_point
    plate_vertical_force = vertical_force / lug_plates_per_point
    plate_horizontal_force = horizontal_force / lug_plates_per_point
    weld_size = 0.7 * t
    pin_diameter = hole_diameter - 4
    e = height - 0.5 * b_width
    a = 0.5 * b_width - 0.5 * hole_diameter
    b = a
    z = math.sqrt((a + 0.5 * hole_diameter) ** 2 - (0.5 * hole_diameter) ** 2)
    beff = 2 * t + 16
    net_width = min(beff, b - hole_diameter / 3)
    connection_width = b_width - gap
    strengths = material_strength(material, t)

    if min(pin_diameter, a, z, net_width, connection_width, weld_size) <= 0:
        raise ValueError("存在非正尺寸，请检查孔径、宽度、高度、间隙和厚度。")

    edge_pass = beff <= b and a >= 4 / 3 * beff
    checks = [
        CheckResult(
            "销轴孔净截面抗拉强度",
            plate_force * 1000 / t / net_width,
            strengths["tensile"],
            "N/mm2",
            "单片耳板力*1000/(t*净宽)",
            f"{plate_force:.3f}*1000/({t:.3f}*{net_width:.3f})",
        ),
        CheckResult(
            "耳板端部抗拉劈开强度",
            2 * plate_force * 1000 / t / a,
            strengths["tensile"],
            "N/mm2",
            "2*单片耳板力*1000/(t*a)",
            f"2*{plate_force:.3f}*1000/({t:.3f}*{a:.3f})",
        ),
        CheckResult(
            "耳板抗剪强度",
            plate_force * 1000 / t / z,
            strengths["shear"],
            "N/mm2",
            "单片耳板力*1000/(t*Z)",
            f"{plate_force:.3f}*1000/({t:.3f}*{z:.3f})",
        ),
        CheckResult(
            "耳板端部承压强度",
            plate_force * 1000 / pin_diameter / (t + 2 * t1),
            strengths["bearing"],
            "N/mm2",
            "单片耳板力*1000/(d*(t+2*t1))",
            f"{plate_force:.3f}*1000/({pin_diameter:.3f}*({t:.3f}+2*{t1:.3f}))",
        ),
    ]

    section_modulus = (t * b_width * b_width - t * gap * gap) / 6
    normal_stress = plate_vertical_force * 1000 / (t * connection_width)
    moment = plate_horizontal_force * e / 1000
    shear_stress = plate_horizontal_force * 1000 / (t * connection_width)
    bending_stress = moment * 1000000 / section_modulus
    combined_stress = ((normal_stress + bending_stress) ** 2 + 3 * shear_stress**2) ** 0.5
    checks.append(
        CheckResult(
            "耳板与构件连接处截面承载力",
            combined_stress,
            1.1 * strengths["tensile"],
            "N/mm2",
            "sqrt((sigma+sigma')^2+3*tau^2)",
            f"sigma={normal_stress:.3f}, sigma'={bending_stress:.3f}, tau={shear_stress:.3f}; "
            f"sqrt(({normal_stress:.3f}+{bending_stress:.3f})^2+3*{shear_stress:.3f}^2)",
        )
    )

    weld_throat = 0.7 * weld_size
    weld_length = connection_width - 2 * weld_size
    weld_modulus = 2 * 0.7 * weld_size * ((b_width - 2 * weld_size) ** 2 - gap**2) / 6
    if min(weld_throat, weld_length, weld_modulus) <= 0:
        raise ValueError("焊缝计算长度或截面模量为非正值，请调大宽度或减小焊脚。")

    weld_normal = plate_vertical_force * 1000 / (2 * 0.7 * weld_size * weld_length)
    weld_shear = plate_horizontal_force * 1000 / (2 * 0.7 * weld_size * weld_length)
    weld_bending = moment * 1000000 / weld_modulus
    weld_equivalent = ((weld_normal / beta + weld_bending / beta) ** 2 + weld_shear**2) ** 0.5
    checks.append(
        CheckResult(
            "双面角焊缝承载力",
            weld_equivalent,
            strengths["weld"],
            "N/mm2",
            "sqrt(((sigmaN+sigmaM)/beta)^2+tauV^2)",
            f"sigmaN={weld_normal:.3f}, sigmaM={weld_bending:.3f}, tauV={weld_shear:.3f}, beta={beta:.3f}; "
            f"sqrt((({weld_normal:.3f}+{weld_bending:.3f})/{beta:.3f})^2+{weld_shear:.3f}^2)",
        )
    )

    derived = {
        "pull_force": pull_force,
        "vertical_force": vertical_force,
        "horizontal_force": horizontal_force,
        "plate_force": plate_force,
        "plate_vertical_force": plate_vertical_force,
        "plate_horizontal_force": plate_horizontal_force,
        "lift_points": lift_points,
        "lug_plates_per_point": lug_plates_per_point,
        "weld_size": weld_size,
        "pin_diameter": pin_diameter,
        "edge_a": a,
        "edge_b": b,
        "edge_z": z,
        "beff": beff,
        "net_width": net_width,
        "connection_width": connection_width,
        "section_modulus": section_modulus,
        "moment": moment,
        "normal_stress": normal_stress,
        "shear_stress": shear_stress,
        "bending_stress": bending_stress,
        "weld_length": weld_length,
        "weld_modulus": weld_modulus,
        "weld_throat": weld_throat,
        "weld_normal": weld_normal,
        "weld_shear": weld_shear,
        "weld_bending": weld_bending,
        "strengths": strengths,
        "edge_pass": edge_pass,
    }
    all_passed = edge_pass and all(item.passed for item in checks)
    return derived, checks, all_passed


class LiftingLugApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("钢梁板式吊耳验算程序")
        self.geometry("1180x760")
        self.minsize(1080, 680)
        self.vars = {}
        self.last_result = None
        self._build_style()
        self._build_ui()
        self.apply_template("5t")

    def _build_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f5f7fb")
        style.configure("Card.TFrame", background="#ffffff", relief="solid", borderwidth=1)
        style.configure("TLabel", background="#f5f7fb", font=("Microsoft YaHei UI", 10))
        style.configure("Card.TLabel", background="#ffffff", font=("Microsoft YaHei UI", 10))
        style.configure("Title.TLabel", background="#f5f7fb", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Section.TLabel", background="#ffffff", font=("Microsoft YaHei UI", 12, "bold"))
        style.configure("Pass.TLabel", background="#ffffff", foreground="#0f8a43", font=("Microsoft YaHei UI", 16, "bold"))
        style.configure("Fail.TLabel", background="#ffffff", foreground="#b42318", font=("Microsoft YaHei UI", 16, "bold"))
        style.configure("TButton", font=("Microsoft YaHei UI", 10))

    def _build_ui(self):
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, padding=(18, 16, 18, 8))
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Label(header, text="钢梁板式吊耳验算程序", style="Title.TLabel").pack(side="left")

        left = ttk.Frame(self, style="Card.TFrame", padding=14)
        left.grid(row=1, column=0, sticky="nsw", padx=(18, 10), pady=(6, 18))

        right = ttk.Frame(self, style="Card.TFrame", padding=14)
        right.grid(row=1, column=1, sticky="nsew", padx=(0, 18), pady=(6, 18))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)

        ttk.Label(left, text="参数输入", style="Section.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
        ttk.Label(left, text="吨位模板", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=5)
        self.template_var = tk.StringVar(value="5t")
        template = ttk.Combobox(left, textvariable=self.template_var, values=list(TEMPLATES), width=16, state="readonly")
        template.grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Button(left, text="套用", command=lambda: self.apply_template(self.template_var.get())).grid(row=1, column=2, padx=(8, 0))

        fields = [
            ("weight", "钢梁自重 G", "t"),
            ("safety_factor", "安全系数", ""),
            ("lift_points", "吊点数量", "个"),
            ("lug_plates_per_point", "每吊点耳板片数", "片"),
            ("angle_deg", "钢丝绳夹角", "deg"),
            ("pull_force", "吊点拉力 T3", "kN，填 0 自动计算"),
            ("plate_thickness", "耳板厚度 t", "mm"),
            ("cover_plate_thickness", "单侧贴板厚度 t1", "mm"),
            ("plate_width", "耳板宽度 B", "mm"),
            ("root_gap", "根部间隙 A", "mm"),
            ("plate_height", "耳板高度 h", "mm"),
            ("hole_diameter", "耳板孔径 d0", "mm"),
            ("weld_beta", "焊缝增大系数 beta", ""),
        ]
        row = 2
        for key, label, unit in fields:
            ttk.Label(left, text=label, style="Card.TLabel").grid(row=row, column=0, sticky="w", pady=5)
            var = tk.StringVar()
            self.vars[key] = var
            ttk.Entry(left, textvariable=var, width=18).grid(row=row, column=1, sticky="ew", pady=5)
            ttk.Label(left, text=unit, style="Card.TLabel").grid(row=row, column=2, sticky="w", padx=(8, 0), pady=5)
            row += 1

        ttk.Label(left, text="耳板材质", style="Card.TLabel").grid(row=row, column=0, sticky="w", pady=5)
        self.vars["material"] = tk.StringVar()
        ttk.Combobox(left, textvariable=self.vars["material"], values=list(MATERIAL_STRENGTHS), width=16, state="readonly").grid(
            row=row, column=1, sticky="ew", pady=5
        )
        row += 1

        actions = ttk.Frame(left, style="Card.TFrame")
        actions.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(14, 0))
        ttk.Button(actions, text="计算", command=self.recalculate).pack(side="left", fill="x", expand=True)
        ttk.Button(actions, text="导出计算书", command=self.export_report).pack(side="left", fill="x", expand=True, padx=(8, 0))

        self.summary = ttk.Label(right, text="等待计算", style="Pass.TLabel")
        self.summary.grid(row=0, column=0, sticky="w")
        self.derived_text = tk.Text(right, height=8, wrap="word", font=("Consolas", 10), relief="flat", background="#f8fafc")
        self.derived_text.grid(row=1, column=0, sticky="ew", pady=(12, 12))

        columns = ("name", "actual", "allowable", "ratio", "result", "expression")
        self.table = ttk.Treeview(right, columns=columns, show="headings", height=14)
        headings = {
            "name": "验算项目",
            "actual": "计算值",
            "allowable": "允许值",
            "ratio": "利用率",
            "result": "结论",
            "expression": "公式",
        }
        widths = {"name": 210, "actual": 110, "allowable": 110, "ratio": 90, "result": 90, "expression": 260}
        for col in columns:
            self.table.heading(col, text=headings[col])
            self.table.column(col, width=widths[col], anchor="w")
        self.table.grid(row=2, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(right, orient="vertical", command=self.table.yview)
        scrollbar.grid(row=2, column=1, sticky="ns")
        self.table.configure(yscrollcommand=scrollbar.set)

        self.note = ttk.Label(
            right,
            text="说明：公式按现有 Excel 验算链整理。正式工程使用前，应结合采用规范版本和企业校审要求复核。",
            style="Card.TLabel",
        )
        self.note.grid(row=3, column=0, sticky="w", pady=(12, 0))

    def apply_template(self, name):
        for key, value in TEMPLATES[name].items():
            self.vars[key].set(str(value))
        self.recalculate()

    def read_input(self):
        return {key: var.get().strip() for key, var in self.vars.items()}

    def recalculate(self):
        try:
            data = self.read_input()
            derived, checks, all_passed = calculate(data)
            self.last_result = (data, derived, checks, all_passed)
            self.render_result(data, derived, checks, all_passed)
        except Exception as exc:
            messagebox.showerror("计算错误", str(exc))

    def render_result(self, data, derived, checks, all_passed):
        self.summary.configure(
            text="总体结论：满足要求" if all_passed else "总体结论：不满足，请调整参数",
            style="Pass.TLabel" if all_passed else "Fail.TLabel",
        )
        self.derived_text.configure(state="normal")
        self.derived_text.delete("1.0", "end")
        lines = [
            f"吊点数量: {derived['lift_points']:.0f} 个    每吊点耳板片数: {derived['lug_plates_per_point']:.0f} 片",
            f"吊点 T3 拉力设计值: {derived['pull_force']:.3f} kN",
            f"吊点 T1 竖向分量: {derived['vertical_force']:.3f} kN    吊点 T2 水平分量: {derived['horizontal_force']:.3f} kN",
            f"单片耳板设计力: {derived['plate_force']:.3f} kN    竖向: {derived['plate_vertical_force']:.3f} kN    水平: {derived['plate_horizontal_force']:.3f} kN",
            f"销轴直径 d: {derived['pin_diameter']:.3f} mm    焊脚高度 hf: {derived['weld_size']:.3f} mm",
            f"a: {derived['edge_a']:.3f} mm    b: {derived['edge_b']:.3f} mm    Z: {derived['edge_z']:.3f} mm",
            f"beff: {derived['beff']:.3f} mm    净宽: {derived['net_width']:.3f} mm",
            f"材料强度 f: {derived['strengths']['tensile']} N/mm2    fv: {derived['strengths']['shear']} N/mm2",
            f"边距要求: {'满足' if derived['edge_pass'] else '不满足'}",
        ]
        self.derived_text.insert("end", "\n".join(lines))
        self.derived_text.configure(state="disabled")

        self.table.delete(*self.table.get_children())
        self.table.insert(
            "",
            "end",
            values=("边距要求", f"beff={derived['beff']:.2f}, a={derived['edge_a']:.2f}", "b与4/3beff", "-", "满足" if derived["edge_pass"] else "不满足", "beff<=b 且 a>=4/3*beff"),
        )
        for item in checks:
            self.table.insert(
                "",
                "end",
                values=(
                    item.name,
                    f"{item.actual:.3f} {item.unit}",
                    f"{item.allowable:.3f} {item.unit}",
                    f"{item.ratio:.3f}",
                    "满足" if item.passed else "不满足",
                    item.expression,
                ),
            )

    def export_report(self):
        if not self.last_result:
            self.recalculate()
        if not self.last_result:
            return
        data, derived, checks, all_passed = self.last_result
        default_name = f"吊耳计算书_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        path = filedialog.asksaveasfilename(
            title="导出计算书",
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.build_word_report(path, data, derived, checks, all_passed)
        messagebox.showinfo("导出完成", f"已导出：\n{path}")

    def build_word_report(self, path, data, derived, checks, all_passed):
        document = Document()
        self.configure_word_styles(document)

        title = document.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title.add_run("钢梁板式吊耳验算计算书")
        title_run.bold = True
        title_run.font.size = Pt(18)

        subtitle = document.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle.add_run(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        conclusion = document.add_paragraph()
        conclusion.add_run("总体结论：").bold = True
        result_run = conclusion.add_run("满足要求" if all_passed else "不满足，请调整参数")
        result_run.bold = True

        document.add_heading("一、计算简图", level=1)
        self.add_lug_diagram(document)

        document.add_heading("二、输入参数", level=1)
        labels = self.report_labels()
        parameter_rows = [(index, label, data[key]) for index, (key, label) in enumerate(labels.items(), start=1)]
        self.add_parameter_table(document, parameter_rows)

        document.add_heading("三、主要中间量", level=1)
        derived_lines = [
            f"1. 吊点数量 n = {derived['lift_points']:.0f} 个；每吊点耳板片数 m = {derived['lug_plates_per_point']:.0f} 片。",
            f"2. 吊点钢丝绳拉力设计值 T3 = {derived['pull_force']:.3f} kN。",
            f"3. 吊点竖向分量 T1 = {derived['vertical_force']:.3f} kN；吊点水平分量 T2 = {derived['horizontal_force']:.3f} kN。",
            f"4. 单片耳板设计力 N = T3 / m = {derived['plate_force']:.3f} kN。",
            f"5. 单片耳板竖向分量 N1 = {derived['plate_vertical_force']:.3f} kN；单片耳板水平分量 N2 = {derived['plate_horizontal_force']:.3f} kN。",
            f"6. 销轴直径 d = {derived['pin_diameter']:.3f} mm；焊脚高度 hf = {derived['weld_size']:.3f} mm。",
            f"7. 耳板净距 a = {derived['edge_a']:.3f} mm，b = {derived['edge_b']:.3f} mm，端部宽度 Z = {derived['edge_z']:.3f} mm。",
            f"8. beff = {derived['beff']:.3f} mm；净宽 = {derived['net_width']:.3f} mm；焊缝计算长度 lw = {derived['weld_length']:.3f} mm。",
            f"9. 连接处截面模量 W = {derived['section_modulus']:.3f} mm3；偏心弯矩 M = {derived['moment']:.3f} kN*m。",
            f"10. 连接处正应力 sigma = {derived['normal_stress']:.3f} N/mm2；弯曲正应力 sigma' = {derived['bending_stress']:.3f} N/mm2；剪应力 tau = {derived['shear_stress']:.3f} N/mm2。",
            f"11. 焊缝计算厚度 he = {derived['weld_throat']:.3f} mm；焊缝截面模量 W1 = {derived['weld_modulus']:.3f} mm3。",
            f"12. 焊缝正应力 sigmaN = {derived['weld_normal']:.3f} N/mm2；焊缝弯曲正应力 sigmaM = {derived['weld_bending']:.3f} N/mm2；焊缝剪应力 tauV = {derived['weld_shear']:.3f} N/mm2。",
        ]
        for line in derived_lines:
            self.add_numbered_text(document, line)

        document.add_heading("四、验算结果", level=1)
        edge_result = "满足" if derived["edge_pass"] else "不满足"
        self.add_check_text(
            document,
            "4.1 边距要求",
            "判定式：beff <= b 且 a >= 4/3*beff。",
            f"代入数值：beff = {derived['beff']:.3f} mm，b = {derived['edge_b']:.3f} mm；"
            f"a = {derived['edge_a']:.3f} mm，4/3*beff = {4 / 3 * derived['beff']:.3f} mm。",
            f"结论：{edge_result}。",
        )
        for index, item in enumerate(checks, start=2):
            result = "满足" if item.passed else "不满足"
            self.add_check_text(
                document,
                f"4.{index} {item.name}",
                f"计算式：{item.expression}。",
                f"代入数值：{item.substitution} = {item.actual:.3f} {item.unit}。",
                f"计算值 = {item.actual:.3f} {item.unit}；允许值 = {item.allowable:.3f} {item.unit}；"
                f"利用率 = {item.ratio:.3f}，结论：{result}。",
            )

        note = document.add_paragraph()
        note.add_run("说明：").bold = True
        note.add_run("本计算书按当前程序内置公式生成，正式工程应用前应结合采用规范版本、设计条件和企业校审要求复核。")

        document.save(path)

    @staticmethod
    def add_lug_diagram(document):
        diagram_path = Path(__file__).resolve().parent / "assets" / "lifting_lug_diagram.png"
        if not diagram_path.exists():
            paragraph = document.add_paragraph()
            paragraph.add_run("计算简图：").bold = True
            paragraph.add_run("未找到内置简图文件 assets/lifting_lug_diagram.png。")
            return

        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run()
        run.add_picture(str(diagram_path), width=Inches(6.4))

        caption = document.add_paragraph()
        caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption.add_run("图 1  吊耳尺寸、受力及焊缝计算简图")

    @staticmethod
    def add_numbered_text(document, text):
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.left_indent = Pt(18)
        paragraph.paragraph_format.first_line_indent = Pt(-18)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.add_run(text)

    @staticmethod
    def add_parameter_table(document, rows):
        table = document.add_table(rows=1, cols=3)
        table.style = "Table Grid"
        table.autofit = False
        headers = ["序号", "已知参数", "数值"]
        widths = [900, 3900, 3600]
        for index, header in enumerate(headers):
            table.rows[0].cells[index].text = header
        for row in rows:
            cells = table.add_row().cells
            cells[0].text = str(row[0])
            cells[1].text = str(row[1])
            cells[2].text = str(row[2])
        for row in table.rows:
            for index, width in enumerate(widths):
                row.cells[index].width = width
                for paragraph in row.cells[index].paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Microsoft YaHei"
                        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        for cell in table.rows[0].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True
        document.add_paragraph()

    @staticmethod
    def add_check_text(document, title, formula_text, substitution_text, result_text):
        title_paragraph = document.add_paragraph()
        title_paragraph.paragraph_format.space_before = Pt(6)
        title_paragraph.paragraph_format.space_after = Pt(2)
        title_run = title_paragraph.add_run(title)
        title_run.bold = True

        formula_paragraph = document.add_paragraph()
        formula_paragraph.paragraph_format.left_indent = Pt(18)
        formula_paragraph.paragraph_format.space_after = Pt(2)
        formula_paragraph.add_run(formula_text)

        substitution_paragraph = document.add_paragraph()
        substitution_paragraph.paragraph_format.left_indent = Pt(18)
        substitution_paragraph.paragraph_format.space_after = Pt(2)
        substitution_paragraph.add_run(substitution_text)

        result_paragraph = document.add_paragraph()
        result_paragraph.paragraph_format.left_indent = Pt(18)
        result_paragraph.paragraph_format.space_after = Pt(4)
        result_paragraph.add_run(result_text)

    @staticmethod
    def configure_word_styles(document):
        for style_name in ["Normal", "Heading 1", "Heading 2"]:
            style = document.styles[style_name]
            style.font.name = "Microsoft YaHei"
            style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        document.styles["Normal"].font.size = Pt(10.5)

    @staticmethod
    def report_labels():
        return {
            "weight": "钢梁自重 G(t)",
            "safety_factor": "安全系数",
            "lift_points": "吊点数量",
            "lug_plates_per_point": "每吊点耳板片数",
            "angle_deg": "钢丝绳夹角(deg)",
            "pull_force": "吊点拉力 T3(kN，0 表示自动)",
            "material": "耳板材质",
            "plate_thickness": "耳板厚度 t(mm)",
            "cover_plate_thickness": "单侧贴板厚度 t1(mm)",
            "plate_width": "耳板宽度 B(mm)",
            "root_gap": "根部间隙 A(mm)",
            "plate_height": "耳板高度 h(mm)",
            "hole_diameter": "耳板孔径 d0(mm)",
            "weld_beta": "焊缝增大系数 beta",
        }

    def build_report(self, data, derived, checks, all_passed):
        lines = [
            "钢梁板式吊耳验算计算书",
            f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            "一、输入参数",
        ]
        labels = {
            "weight": "钢梁自重 G(t)",
            "safety_factor": "安全系数",
            "lift_points": "吊点数量",
            "lug_plates_per_point": "每吊点耳板片数",
            "angle_deg": "钢丝绳夹角(deg)",
            "pull_force": "吊点拉力 T3(kN，0 表示自动)",
            "material": "耳板材质",
            "plate_thickness": "耳板厚度 t(mm)",
            "cover_plate_thickness": "单侧贴板厚度 t1(mm)",
            "plate_width": "耳板宽度 B(mm)",
            "root_gap": "根部间隙 A(mm)",
            "plate_height": "耳板高度 h(mm)",
            "hole_diameter": "耳板孔径 d0(mm)",
            "weld_beta": "焊缝增大系数 beta",
        }
        for key, label in labels.items():
            lines.append(f"{label}：{data[key]}")
        lines.extend(
            [
                "",
                "二、主要中间量",
                f"吊点数量={derived['lift_points']:.0f}，每吊点耳板片数={derived['lug_plates_per_point']:.0f}",
                f"吊点 T3={derived['pull_force']:.3f} kN，吊点 T1={derived['vertical_force']:.3f} kN，吊点 T2={derived['horizontal_force']:.3f} kN",
                f"单片耳板设计力={derived['plate_force']:.3f} kN，竖向={derived['plate_vertical_force']:.3f} kN，水平={derived['plate_horizontal_force']:.3f} kN",
                f"d={derived['pin_diameter']:.3f} mm，hf={derived['weld_size']:.3f} mm，a={derived['edge_a']:.3f} mm，b={derived['edge_b']:.3f} mm，Z={derived['edge_z']:.3f} mm",
                f"beff={derived['beff']:.3f} mm，净宽={derived['net_width']:.3f} mm，焊缝计算长度={derived['weld_length']:.3f} mm",
                "",
                "三、验算结果",
                f"边距要求：{'满足' if derived['edge_pass'] else '不满足'}",
            ]
        )
        for item in checks:
            lines.append(
                f"{item.name}：计算值 {item.actual:.3f} {item.unit}，允许值 {item.allowable:.3f} {item.unit}，利用率 {item.ratio:.3f}，{'满足' if item.passed else '不满足'}"
            )
        lines.extend(["", f"总体结论：{'满足要求' if all_passed else '不满足，请调整参数'}"])
        return "\n".join(lines)


if __name__ == "__main__":
    app = LiftingLugApp()
    app.mainloop()
