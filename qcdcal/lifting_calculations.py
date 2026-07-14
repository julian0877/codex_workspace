from __future__ import annotations

import math
from typing import Any


Number = int | float


def _require_non_zero(value: Number, name: str) -> None:
    if value == 0:
        raise ValueError(f"{name}不能为0")


def overturning_calculation(
    *,
    crane_weight: Number,
    lifted_weight: Number,
    counterweight: Number,
    working_radius: Number,
    wind_height: Number,
    crane_center_to_tip: Number,
    counterweight_to_center: Number,
    gravity: Number,
    crane_weight_factor: Number,
    lifted_weight_factor: Number,
    wind_factor: Number,
) -> dict[str, Any]:
    wind_load = lifted_weight * 0.2 * gravity
    crane_moment = (
        crane_weight * gravity * crane_center_to_tip
        + counterweight * gravity * (crane_center_to_tip + counterweight_to_center)
    )
    lifted_moment = -lifted_weight * gravity * (working_radius - crane_center_to_tip)
    wind_moment = -wind_load * wind_height
    resultant_moment = (
        crane_moment * crane_weight_factor
        + lifted_moment * lifted_weight_factor
        + wind_moment * wind_factor
    )

    return {
        "wind_load": wind_load,
        "crane_moment": crane_moment,
        "lifted_moment": lifted_moment,
        "wind_moment": wind_moment,
        "resultant_moment": resultant_moment,
        "judgement": "满足要求" if resultant_moment > 0 else "不满足要求",
    }


def outrigger_reaction_calculation(
    *,
    boom_weight: Number,
    boom_center_distance: Number,
    crane_weight_without_boom: Number,
    lifted_weight: Number,
    counterweight: Number,
    working_radius: Number,
    longitudinal_distance: Number,
    transverse_distance: Number,
    counterweight_to_center: Number,
    gravity: Number,
    center_to_rear_outrigger: Number,
    ground_box_area: Number,
) -> dict[str, float]:
    _require_non_zero(longitudinal_distance, "支腿纵向距离")
    _require_non_zero(transverse_distance, "支腿横向距离")
    _require_non_zero(center_to_rear_outrigger, "回转中心至后支腿距离")
    _require_non_zero(ground_box_area, "路基箱面积")

    vertical_force = (
        crane_weight_without_boom + lifted_weight + counterweight + boom_weight
    ) * gravity
    total_moment = (
        working_radius * lifted_weight * gravity
        - counterweight * gravity * counterweight_to_center
        + boom_weight * boom_center_distance * gravity
    )
    angle_beta = math.degrees(
        math.atan(transverse_distance / 2 / center_to_rear_outrigger)
    )
    cos_beta = math.cos(math.radians(angle_beta))
    sin_beta = math.sin(math.radians(angle_beta))
    moment_x = total_moment * cos_beta
    moment_y = total_moment * sin_beta

    reaction_1 = (
        center_to_rear_outrigger * vertical_force / 2 / longitudinal_distance
        - moment_x / 2 / longitudinal_distance
        + moment_y / 2 / transverse_distance
    )
    reaction_2 = (
        (longitudinal_distance - center_to_rear_outrigger)
        * vertical_force
        / 2
        / longitudinal_distance
        + moment_x / 2 / longitudinal_distance
        + moment_y / 2 / transverse_distance
    )
    reaction_3 = (
        center_to_rear_outrigger * vertical_force / 2 / longitudinal_distance
        - moment_x / 2 / longitudinal_distance
        - moment_y / 2 / transverse_distance
    )
    reaction_4 = (
        (longitudinal_distance - center_to_rear_outrigger)
        * vertical_force
        / 2
        / longitudinal_distance
        + moment_x / 2 / longitudinal_distance
        - moment_y / 2 / transverse_distance
    )
    max_reaction = max(reaction_1, reaction_2, reaction_3, reaction_4)

    return {
        "vertical_force": vertical_force,
        "total_moment": total_moment,
        "angle_beta": angle_beta,
        "cos_beta": cos_beta,
        "sin_beta": sin_beta,
        "moment_x": moment_x,
        "moment_y": moment_y,
        "reaction_1": reaction_1,
        "reaction_2": reaction_2,
        "reaction_3": reaction_3,
        "reaction_4": reaction_4,
        "max_reaction": max_reaction,
        "max_ground_pressure": max_reaction / ground_box_area,
    }
