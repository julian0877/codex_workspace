import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path

from generate_word_calculation_book import generate_word_from_data


def _make_label(frame, text, row, column=0, colspan=1):
    label = tk.Label(frame, text=text, anchor='w')
    label.grid(row=row, column=column, sticky='w', padx=4, pady=2, columnspan=colspan)
    return label


def _make_entry(frame, row, column=1, colspan=1, width=24, textvariable=None, state='normal'):
    entry = tk.Entry(frame, width=width, textvariable=textvariable, state=state)
    entry.grid(row=row, column=column, padx=4, pady=2, columnspan=colspan)
    return entry


def _make_text(frame, row, column=1, height=3, width=54, textvariable=None):
    text_widget = scrolledtext.ScrolledText(frame, height=height, width=width)
    text_widget.grid(row=row, column=column, padx=4, pady=2, columnspan=3)
    if textvariable is not None:
        text_widget.insert('1.0', textvariable.get())
    return text_widget


def _make_combobox(frame, row, column=1, width=20, values=None, textvariable=None, state='normal'):
    combobox = ttk.Combobox(frame, width=width, values=values or [], textvariable=textvariable, state=state)
    combobox.grid(row=row, column=column, padx=4, pady=2)
    return combobox


def get_design_strengths(material: str, thickness: float):
    material = material.strip()
    if thickness <= 0:
        raise ValueError('厚度必须大于 0')
    if material == 'Q235':
        if thickness <= 16:
            return 215.0, 125.0
        if thickness <= 40:
            return 205.0, 120.0
        if thickness <= 100:
            return 200.0, 115.0
        raise ValueError('Q235 材料厚度超出支持范围')
    # 其他材料按 Q345 处理
    if thickness <= 16:
        return 305.0, 175.0
    if thickness <= 40:
        return 295.0, 170.0
    if thickness <= 63:
        return 290.0, 165.0
    if thickness <= 80:
        return 280.0, 160.0
    if thickness <= 100:
        return 270.0, 155.0
    raise ValueError(f'{material} 材料厚度超出支持范围')


def update_design_strengths(*args):
    try:
        thickness = float(plate_thickness_var.get())
        f1, f1v = get_design_strengths(plate_material_var.get(), thickness)
        tensile_strength_var.set(f'{f1:.0f}')
        shear_strength_var.set(f'{f1v:.0f}')
    except ValueError:
        tensile_strength_var.set('')
        shear_strength_var.set('')


def calculate_results():
    try:
        beam_weight = float(beam_weight_var.get())
        safety_factor = float(safety_factor_var.get())
        rope_angle = float(rope_angle_var.get())
        plate_thickness = float(plate_thickness_var.get())
        plate_width = float(plate_width_var.get())
        pin_diameter = float(pin_diameter_var.get())
        hole_a = float(hole_a_var.get())
        plate_end_distance = float(plate_end_distance_var.get())
        tensile_strength = float(tensile_strength_var.get())
        shear_strength = float(shear_strength_var.get())
        weld_height = float(weld_height_var.get())
        weld_thickness = float(weld_thickness_var.get())
        weld_length = float(weld_length_var.get())
    except ValueError:
        messagebox.showwarning('参数错误', '请先选择材质、输入几何尺寸，并填写所有其他已知参数。')
        return

    from math import radians, sin, cos, sqrt

    if rope_angle <= 0:
        messagebox.showwarning('参数错误', '钢丝绳夹角必须大于 0°。')
        return
    if plate_width <= pin_diameter:
        messagebox.showwarning('参数错误', '耳板宽度必须大于销轴孔直径 d0。')
        return
    if plate_thickness <= 0 or pin_diameter <= 0:
        messagebox.showwarning('参数错误', '板厚和销轴孔直径必须大于 0。')
        return

    # 按 Excel 公式进行受力分解
    t3 = 0.5 * beam_weight * 9.8 * safety_factor / sin(radians(rope_angle))
    t1 = t3 * sin(radians(rope_angle))
    t2 = t3 * cos(radians(rope_angle))

    # 计算销轴孔净截面及端部、剪切、承压、截面承载力
    effective_width = min(2 * plate_thickness + 16, plate_width - pin_diameter / 3)
    net_tension_stress = t3 * 1000 / (2 * plate_thickness * effective_width)
    c30 = 0.5 * plate_width - 0.5 * pin_diameter
    end_tensile_stress = 2 * t3 * 1000 / (plate_thickness * c30)
    z_value = sqrt(max((c30 + 0.5 * pin_diameter) ** 2 - (0.5 * pin_diameter) ** 2, 0.0))
    shear_stress = t3 * 1000 / (2 * plate_thickness * z_value)
    c25 = pin_diameter - 4
    bearing_stress = t3 * 1000 / (c25 * (plate_thickness + 2 * plate_end_distance))
    l_value = plate_width - hole_a
    e_value = weld_height - 0.5 * plate_width
    w_value = (plate_thickness * plate_width ** 2 - plate_thickness * hole_a ** 2) / 6
    axial_stress = t1 * 1000 / (plate_thickness * l_value)
    shear_section_stress = t2 * 1000 / (plate_thickness * l_value)
    bending_stress = t2 * e_value * 1000000 / w_value
    combined_stress = ((axial_stress + bending_stress) ** 2 + 3 * shear_section_stress ** 2) ** 0.5

    fcb = 305.0 if plate_material_var.get().strip() == 'Q235' else 385.0
    ffw = 160.0 if plate_material_var.get().strip() == 'Q235' else 200.0
    weld_throat = weld_thickness if weld_thickness > 0 else 0.7 * weld_height
    weld_area = 2 * weld_length * weld_throat
    weld_capacity = weld_area * ffw

    net_tension_ok = net_tension_stress <= tensile_strength
    end_tensile_ok = end_tensile_stress <= tensile_strength
    shear_ok = shear_stress <= shear_strength
    bearing_ok = bearing_stress <= fcb
    combined_ok = combined_stress <= 1.1 * tensile_strength
    weld_ok = t3 * 1000 <= weld_capacity

    ear_result_value_var.set(
        f'T3={t3:.2f} kN, T1={t1:.2f} kN, T2={t2:.2f} kN'
    )

    ear_status = '通过' if net_tension_ok and end_tensile_ok and shear_ok and bearing_ok and combined_ok else '不通过'
    ear_result_status_var.set(ear_status)

    weld_result_value_var.set(
        f'焊缝喉厚={weld_throat:.2f} mm, 有效面积={weld_area:.1f} mm², 焊缝强度={ffw:.0f} N/mm², 允许承载力={weld_capacity:.1f} N'
    )
    weld_status = '通过' if weld_ok else '不通过'
    weld_result_status_var.set(weld_status)

    table_text = (
        '项次      计算内容                  结果                判断\n'
        f'4.2   销轴孔净截面抗拉        {net_tension_stress:7.2f} N/mm²   {"通过" if net_tension_ok else "不通过"}\n'
        f'4.2   耳板端部抗拉(劈开)        {end_tensile_stress:7.2f} N/mm²   {"通过" if end_tensile_ok else "不通过"}\n'
        f'4.3   耳板抗剪强度            {shear_stress:7.2f} N/mm²   {"通过" if shear_ok else "不通过"}\n'
        f'4.4   耳板端部承压强度        {bearing_stress:7.2f} N/mm²   {"通过" if bearing_ok else "不通过"}\n'
        f'4.5   耳板与构件连接处截面承载力 {combined_stress:7.2f} N/mm²   {"通过" if combined_ok else "不通过"}\n'
        f'5     角焊缝承载力            {weld_capacity:10.1f} N      {"通过" if weld_ok else "不通过"}'
    )
    results_text.config(state='normal')
    results_text.delete('1.0', 'end')
    results_text.insert('1.0', table_text)
    results_text.config(state='disabled')

    conclusion_text.delete('1.0', 'end')
    conclusion_text.insert(
        '1.0',
        f'计算结果：耳板验算[{ear_status}]，角焊缝验算[{weld_status}]。',
    )
    messagebox.showinfo('计算完成', '已完成计算，结果已更新。')


def generate_docx():
    if not ear_result_value_var.get().strip() or not weld_result_value_var.get().strip():
        if not messagebox.askyesno('未计算', '检测到未生成计算结果，是否继续保存当前内容？'):
            return

    known_conditions_str = '\n'.join([
        f'钢梁吨位={beam_weight_var.get().strip()} t',
        f'安全系数={safety_factor_var.get().strip()}',
        f'钢丝绳与水平面夹角θ={rope_angle_var.get().strip()}°',
        f'耳板厚度 t={plate_thickness_var.get().strip()} mm',
    ])
    material_data_str = '\n'.join([
        f'耳板材质={plate_material_var.get().strip()}',
        f'耳板宽度 B={plate_width_var.get().strip()} mm',
        f'销轴孔直径 d0={pin_diameter_var.get().strip()} mm',
        f'销轴孔 A={hole_a_var.get().strip()} mm',
        f'耳板端部距 t1={plate_end_distance_var.get().strip()} mm',
        f'抗拉强度 f1={tensile_strength_var.get().strip()} N/mm2',
        f'抗剪强度 f1v={shear_strength_var.get().strip()} N/mm2',
        f'焊缝高度 h={weld_height_var.get().strip()} mm',
        f'焊缝喉厚 a={weld_thickness_var.get().strip()} mm',
        f'焊缝长度 l={weld_length_var.get().strip()} mm',
    ])
    ear_results_str = '\n'.join([
        f'耳板承载力计算结果={ear_result_value_var.get().strip()}',
        f'耳板承载力判定={ear_result_status_var.get().strip()}',
    ])
    weld_results_str = '\n'.join([
        f'双面角焊缝承载力计算结果={weld_result_value_var.get().strip()}',
        f'双面角焊缝判定={weld_result_status_var.get().strip()}',
    ])

    data = {
        'project_name': project_name_var.get().strip(),
        'date': date_var.get().strip(),
        'project_info': project_info_var.get().strip(),
        'author': author_var.get().strip(),
        'reviewer': reviewer_var.get().strip(),
        'diagram_text': diagram_text.get('1.0', 'end').strip(),
        'known_conditions': known_conditions_str,
        'material_data': material_data_str,
        'ear_results': ear_results_str,
        'weld_results': weld_results_str,
        'conclusion': conclusion_text.get('1.0', 'end').strip(),
    }

    output_path = filedialog.asksaveasfilename(
        defaultextension='.docx',
        filetypes=[('Word 文档', '*.docx')],
        initialfile='吊耳计算书.docx',
        initialdir=str(Path(__file__).resolve().parent),
        title='保存计算书为',
    )
    if not output_path:
        return

    try:
        generated_path = generate_word_from_data(data, output_path)
        messagebox.showinfo('完成', f'已生成计算书：\n{generated_path}')
    except Exception as exc:
        messagebox.showerror('错误', f'生成失败：{exc}')


def load_sample_data():
    project_name_var.set('钢梁板式吊耳验算')
    date_var.set('2026-06-27')
    project_info_var.set('某工程钢梁吊装')
    author_var.set('设计人员')
    reviewer_var.set('审核人员')
    diagram_text.delete('1.0', 'end')
    diagram_text.insert('1.0', '本计算基于吊耳双点起吊形式，考虑吊装荷载、吊耳材质与焊缝强度。')
    beam_weight_var.set('40')
    safety_factor_var.set('1.32')
    rope_angle_var.set('45')
    plate_thickness_var.set('30')
    plate_material_var.set('Q345')
    plate_width_var.set('256')
    pin_diameter_var.set('30')
    hole_a_var.set('5')
    plate_end_distance_var.set('0')
    tensile_strength_var.set('295')
    shear_strength_var.set('170')
    weld_height_var.set('20')
    weld_thickness_var.set('14')
    weld_length_var.set('100')
    ear_result_value_var.set('T1=0.00 kN, T2=0.00 kN, 截面=256×30 mm')
    ear_result_status_var.set('通过')
    weld_result_value_var.set('焊缝有效截面=256×21.0 mm, 允许承载力=0.0 N')
    weld_result_status_var.set('通过')
    conclusion_text.delete('1.0', 'end')
    conclusion_text.insert('1.0', '经计算，吊耳和焊缝均满足验算要求。')


root = tk.Tk()
root.title('吊耳计算书生成器')
root.geometry('1000x860')

frame = tk.Frame(root)
frame.pack(fill='both', expand=True, padx=12, pady=12)

project_name_var = tk.StringVar()
date_var = tk.StringVar()
project_info_var = tk.StringVar()
author_var = tk.StringVar()
reviewer_var = tk.StringVar()
beam_weight_var = tk.StringVar()
safety_factor_var = tk.StringVar()
rope_angle_var = tk.StringVar()
plate_thickness_var = tk.StringVar()
plate_material_var = tk.StringVar(value='Q345')
plate_width_var = tk.StringVar()
pin_diameter_var = tk.StringVar(value='30')
hole_a_var = tk.StringVar(value='5')
plate_end_distance_var = tk.StringVar(value='0')
tensile_strength_var = tk.StringVar()
shear_strength_var = tk.StringVar()
weld_height_var = tk.StringVar()
weld_thickness_var = tk.StringVar()
weld_length_var = tk.StringVar()
ear_result_value_var = tk.StringVar()
ear_result_status_var = tk.StringVar(value='通过')
weld_result_value_var = tk.StringVar()
weld_result_status_var = tk.StringVar(value='通过')

plate_material_var.trace_add('write', update_design_strengths)
plate_thickness_var.trace_add('write', update_design_strengths)

_make_label(frame, '项目基本信息', 0, colspan=4)
_make_label(frame, '项目名称：', 1)
_make_entry(frame, 1, textvariable=project_name_var)
_make_label(frame, '日期：', 2)
_make_entry(frame, 2, textvariable=date_var)
_make_label(frame, '设计单位 / 工程：', 3)
_make_entry(frame, 3, textvariable=project_info_var)
_make_label(frame, '核算人员：', 4)
_make_entry(frame, 4, textvariable=author_var)
_make_label(frame, '审核人员：', 5)
_make_entry(frame, 5, textvariable=reviewer_var)

_make_label(frame, '一、计算简图说明：', 6, colspan=4)
diagram_text = _make_text(frame, 6, height=4, width=100)

_make_label(frame, '二、已知参数：', 7, colspan=4)
_make_label(frame, '钢梁吨位 (t)：', 8)
_make_entry(frame, 8, textvariable=beam_weight_var)
_make_label(frame, '安全系数：', 8, column=2)
_make_entry(frame, 8, column=3, textvariable=safety_factor_var)
_make_label(frame, '钢丝绳夹角 θ (°)：', 9)
_make_entry(frame, 9, textvariable=rope_angle_var)
_make_label(frame, '耳板厚度 t (mm)：', 9, column=2)
_make_entry(frame, 9, column=3, textvariable=plate_thickness_var)
_make_label(frame, '耳板材质：', 10)
_make_combobox(
    frame,
    10,
    width=18,
    values=['Q235', 'Q345', 'Q355', '16Mn', 'S355', 'A36'],
    textvariable=plate_material_var,
    state='readonly',
)
_make_label(frame, '耳板宽度 B (mm)：', 10, column=2)
_make_entry(frame, 10, column=3, textvariable=plate_width_var)
_make_label(frame, '销轴孔直径 d0 (mm)：', 11)
_make_entry(frame, 11, textvariable=pin_diameter_var)
_make_label(frame, '销轴孔 A (mm)：', 11, column=2)
_make_entry(frame, 11, column=3, textvariable=hole_a_var)
_make_label(frame, '抗拉强度 f1 (N/mm2)：', 12)
_make_entry(frame, 12, textvariable=tensile_strength_var, state='readonly')
_make_label(frame, '抗剪强度 f1v (N/mm2)：', 12, column=2)
_make_entry(frame, 12, column=3, textvariable=shear_strength_var, state='readonly')
_make_label(frame, '耳板端部距 t1 (mm)：', 13)
_make_entry(frame, 13, textvariable=plate_end_distance_var)
_make_label(frame, '焊缝高度 h (mm)：', 13, column=2)
_make_entry(frame, 13, column=3, textvariable=weld_height_var)
_make_label(frame, '焊缝喉厚 a (mm)：', 14)
_make_entry(frame, 14, textvariable=weld_thickness_var)
_make_label(frame, '焊缝长度 l (mm)：', 14, column=2)
_make_entry(frame, 14, column=3, textvariable=weld_length_var)

_make_label(frame, '三、计算结果：', 14, colspan=4)
_make_label(frame, '耳板承载力计算结果：', 15)
_make_entry(frame, 15, textvariable=ear_result_value_var, width=52, state='readonly')
_make_label(frame, '耳板承载力判定：', 16)
_make_entry(frame, 16, textvariable=ear_result_status_var, width=18, state='readonly')
_make_label(frame, '双面角焊缝承载力计算结果：', 15, column=2)
_make_entry(frame, 15, column=3, textvariable=weld_result_value_var, width=52, state='readonly')
_make_label(frame, '双面角焊缝判定：', 16, column=2)
_make_entry(frame, 16, column=3, textvariable=weld_result_status_var, width=18, state='readonly')
_make_label(frame, '计算结果表：', 17, colspan=4)
results_text = _make_text(frame, 18, height=8, width=100)
results_text.config(font=('Courier New', 10), state='disabled')

_make_label(frame, '四、结论：', 19, colspan=4)
conclusion_text = _make_text(frame, 19, height=4, width=100)

button_frame = tk.Frame(root)
button_frame.pack(fill='x', padx=12, pady=8)

calculate_button = tk.Button(button_frame, text='计算', command=calculate_results, width=16, bg='#2196F3', fg='white')
calculate_button.pack(side='left', padx=4)

save_button = tk.Button(button_frame, text='生成计算书', command=generate_docx, width=16, bg='#4CAF50', fg='white')
save_button.pack(side='left', padx=4)

sample_button = tk.Button(button_frame, text='加载示例数据', command=load_sample_data, width=16)
sample_button.pack(side='left', padx=4)

close_button = tk.Button(button_frame, text='退出', command=root.destroy, width=16)
close_button.pack(side='left', padx=4)

root.mainloop()
