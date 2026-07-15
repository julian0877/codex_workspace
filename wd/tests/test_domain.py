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
