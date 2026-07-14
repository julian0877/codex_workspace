from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable

from lifting_calculations import (
    overturning_calculation,
    outrigger_reaction_calculation,
)


InputSpec = tuple[str, str, str, float]
ResultSpec = tuple[str, str, str]


OVERTURNING_INPUTS: list[InputSpec] = [
    ("crane_weight", "汽车吊自重 G", "t", 72),
    ("lifted_weight", "起吊构件最大重量 Q", "t", 27),
    ("counterweight", "配重 G1", "t", 100),
    ("working_radius", "起吊作业半径 R", "m", 20),
    ("wind_height", "风动载合力点高度 H", "m", 22),
    ("crane_center_to_tip", "汽车吊重心至支脚倾覆支点距离 a", "m", 4.3),
    ("counterweight_to_center", "配重至回转中心距离 C", "m", 5.75),
    ("gravity", "重力加速度 g", "m/s2", 9.8),
    ("crane_weight_factor", "自重加权系数 KG", "", 1),
    ("lifted_weight_factor", "起升荷载加权系数 KQ", "", 1.15),
    ("wind_factor", "风荷载加权系数 KW", "", 1),
]

OVERTURNING_RESULTS: list[ResultSpec] = [
    ("wind_load", "风荷载 W", "kN"),
    ("crane_moment", "汽车吊总重对倾覆边的力矩 MG", "kN.m"),
    ("lifted_moment", "起升荷载对倾覆边的力矩 MQ", "kN.m"),
    ("wind_moment", "风荷载对倾覆边的力矩 MW", "kN.m"),
    ("resultant_moment", "倾覆边合力矩 M", "kN.m"),
    ("judgement", "结果判断", ""),
]

OUTRIGGER_INPUTS: list[InputSpec] = [
    ("boom_weight", "汽车吊吊臂自重 G3", "t", 10),
    ("boom_center_distance", "吊臂重心至回转重心距离 E", "m", 4),
    ("crane_weight_without_boom", "汽车吊自重(扣除吊臂重量) G0", "t", 31),
    ("lifted_weight", "起吊构件最大重量 G1", "t", 6.02),
    ("counterweight", "配重 G2", "t", 8),
    ("working_radius", "起吊作业半径 L1", "m", 10),
    ("longitudinal_distance", "支腿纵向距离 A", "m", 5.92),
    ("transverse_distance", "支腿横向距离 B", "m", 6.9),
    ("counterweight_to_center", "配重至回转中心距离 C", "m", 0),
    ("gravity", "重力加速度 g", "m/s2", 9.8),
    ("center_to_rear_outrigger", "回转中心至后支腿距离 D", "m", 2.9),
    ("ground_box_area", "路基箱面积 S", "m2", 4),
]

OUTRIGGER_RESULTS: list[ResultSpec] = [
    ("vertical_force", "总竖向力 F", "kN"),
    ("total_moment", "总弯矩 M", "kN.m"),
    ("angle_beta", "水平夹角 beta", "deg"),
    ("cos_beta", "水平夹角余弦值 cos beta", ""),
    ("sin_beta", "水平夹角正弦值 sin beta", ""),
    ("moment_x", "X轴弯矩 Mx", "kN.m"),
    ("moment_y", "Y轴弯矩 My", "kN.m"),
    ("reaction_1", "支腿1反力 N1", "kN"),
    ("reaction_2", "支腿2反力 N2", "kN"),
    ("reaction_3", "支腿3反力 N3", "kN"),
    ("reaction_4", "支腿4反力 N4", "kN"),
    ("max_reaction", "最大支腿反力", "kN"),
    ("max_ground_pressure", "最大地面压强 P", "kPa"),
]


class CalculationPage(ttk.Frame):
    def __init__(
        self,
        parent: ttk.Notebook,
        *,
        title: str,
        inputs: list[InputSpec],
        results: list[ResultSpec],
        calculator: Callable[..., dict[str, object]],
    ) -> None:
        super().__init__(parent, padding=14)
        self.inputs = inputs
        self.results = results
        self.calculator = calculator
        self.entries: dict[str, ttk.Entry] = {}
        self.result_vars: dict[str, tk.StringVar] = {}

        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self._build_inputs(title)
        self._build_results()
        self.calculate()

    def _build_inputs(self, title: str) -> None:
        input_frame = ttk.LabelFrame(self, text="输入参数", padding=10)
        input_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        input_frame.columnconfigure(1, weight=1)

        ttk.Label(input_frame, text=title, font=("", 12, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 10)
        )

        for row, (key, label, unit, default) in enumerate(self.inputs, start=1):
            ttk.Label(input_frame, text=label).grid(row=row, column=0, sticky="w", pady=3)
            entry = ttk.Entry(input_frame, width=16)
            entry.insert(0, str(default))
            entry.grid(row=row, column=1, sticky="ew", padx=8, pady=3)
            ttk.Label(input_frame, text=unit).grid(row=row, column=2, sticky="w", pady=3)
            self.entries[key] = entry

        button_row = len(self.inputs) + 1
        ttk.Button(input_frame, text="计算", command=self.calculate).grid(
            row=button_row, column=0, columnspan=3, sticky="ew", pady=(12, 0)
        )

    def _build_results(self) -> None:
        result_frame = ttk.LabelFrame(self, text="计算结果", padding=10)
        result_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        result_frame.columnconfigure(1, weight=1)

        for row, (key, label, unit) in enumerate(self.results):
            ttk.Label(result_frame, text=label).grid(row=row, column=0, sticky="w", pady=4)
            value_var = tk.StringVar(value="-")
            ttk.Label(result_frame, textvariable=value_var, anchor="e").grid(
                row=row, column=1, sticky="ew", padx=8, pady=4
            )
            ttk.Label(result_frame, text=unit).grid(row=row, column=2, sticky="w", pady=4)
            self.result_vars[key] = value_var

    def _read_inputs(self) -> dict[str, float]:
        values: dict[str, float] = {}
        for key, label, _unit, _default in self.inputs:
            raw = self.entries[key].get().strip()
            try:
                values[key] = float(raw)
            except ValueError as exc:
                raise ValueError(f"{label} 请输入有效数字") from exc
        return values

    def calculate(self) -> None:
        try:
            result = self.calculator(**self._read_inputs())
        except ValueError as exc:
            messagebox.showerror("输入错误", str(exc))
            return

        for key, _label, _unit in self.results:
            value = result[key]
            self.result_vars[key].set(self._format_value(value))

    @staticmethod
    def _format_value(value: object) -> str:
        if isinstance(value, float):
            return f"{value:.3f}"
        return str(value)


class LiftingCalculatorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("吊装相关验算计算程序")
        self.geometry("980x620")
        self.minsize(880, 540)

        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=12, pady=12)

        overturning_page = CalculationPage(
            notebook,
            title="汽车吊抗倾覆验算",
            inputs=OVERTURNING_INPUTS,
            results=OVERTURNING_RESULTS,
            calculator=overturning_calculation,
        )
        outrigger_page = CalculationPage(
            notebook,
            title="汽车吊支腿反力计算",
            inputs=OUTRIGGER_INPUTS,
            results=OUTRIGGER_RESULTS,
            calculator=outrigger_reaction_calculation,
        )

        notebook.add(overturning_page, text="汽车吊抗倾覆验算")
        notebook.add(outrigger_page, text="汽车吊支腿反力计算")


def main() -> None:
    app = LiftingCalculatorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
