from tender_formatter.config import load_profile, read_template_profile, save_profile
from tender_formatter.domain import FormatProfile
from tests.helpers import make_docx_with_styles


def test_profile_json_round_trip(tmp_path):
    path = tmp_path / "company.json"
    save_profile(FormatProfile(name="公司标准"), path)

    loaded = load_profile(path)

    assert loaded.name == "公司标准"
    assert loaded.body.east_asia_font == "宋体"


def test_read_template_extracts_normal_and_heading_styles(tmp_path):
    template = make_docx_with_styles(tmp_path / "template.docx")

    profile = read_template_profile(template, "样板")

    assert profile.body.east_asia_font == "宋体"
    assert profile.body.size_pt == 12
    assert profile.body.line_spacing == 1.5
    assert profile.body.space_before_pt == 6
    assert profile.body.space_after_pt == 3
    assert round(profile.page.left_cm, 1) == 3.2
    assert round(profile.page.right_cm, 1) == 2.2
    assert profile.headings[1].east_asia_font == "黑体"
    assert profile.headings[1].bold is True
