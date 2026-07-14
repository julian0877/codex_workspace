import unittest

from lifting_calculations import (
    overturning_calculation,
    outrigger_reaction_calculation,
)


class LiftingCalculationTests(unittest.TestCase):
    def assertAlmostMapping(self, result, expected, places=6):
        for key, value in expected.items():
            self.assertAlmostEqual(result[key], value, places=places, msg=key)

    def test_overturning_calculation_matches_excel_sample(self):
        result = overturning_calculation(
            crane_weight=72,
            lifted_weight=27,
            counterweight=100,
            working_radius=20,
            wind_height=22,
            crane_center_to_tip=4.3,
            counterweight_to_center=5.75,
            gravity=9.8,
            crane_weight_factor=1,
            lifted_weight_factor=1.15,
            wind_factor=1,
        )

        self.assertAlmostMapping(
            result,
            {
                "wind_load": 52.92,
                "crane_moment": 12883.08,
                "lifted_moment": -4154.22,
                "wind_moment": -1164.24,
                "resultant_moment": 6941.487,
            },
        )
        self.assertEqual(result["judgement"], "满足要求")

    def test_outrigger_reaction_calculation_matches_excel_sample(self):
        result = outrigger_reaction_calculation(
            boom_weight=10,
            boom_center_distance=4,
            crane_weight_without_boom=31,
            lifted_weight=6.02,
            counterweight=8,
            working_radius=10,
            longitudinal_distance=5.92,
            transverse_distance=6.9,
            counterweight_to_center=0,
            gravity=9.8,
            center_to_rear_outrigger=2.9,
            ground_box_area=4,
        )

        self.assertAlmostMapping(
            result,
            {
                "vertical_force": 539.196,
                "total_moment": 981.96,
                "angle_beta": 49.950272234291745,
                "cos_beta": 0.6434522276212659,
                "sin_beta": 0.7654862707908161,
                "moment_x": 631.8443494349782,
                "moment_y": 751.6768984657498,
                "reaction_1": 133.17069609057836,
                "reaction_2": 245.3659848464868,
                "reaction_3": 24.23201515351316,
                "reaction_4": 136.42730390942165,
                "max_reaction": 245.3659848464868,
                "max_ground_pressure": 61.3414962116217,
            },
        )

    def test_outrigger_rejects_zero_ground_box_area(self):
        with self.assertRaises(ValueError):
            outrigger_reaction_calculation(
                boom_weight=10,
                boom_center_distance=4,
                crane_weight_without_boom=31,
                lifted_weight=6.02,
                counterweight=8,
                working_radius=10,
                longitudinal_distance=5.92,
                transverse_distance=6.9,
                counterweight_to_center=0,
                gravity=9.8,
                center_to_rear_outrigger=2.9,
                ground_box_area=0,
            )


if __name__ == "__main__":
    unittest.main()
