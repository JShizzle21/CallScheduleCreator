"""Tests for scheduler_main._compute_weighted_score.

Verifies that each component is normalized to roughly [0, 1] before
being combined with its weight, so that *_WEIGHT constants behave like
real relative weights instead of being swamped by the unbounded
fairness_gap.
"""

from __future__ import annotations

import math

import scheduler_main as sm
from scheduler_main import _compute_weighted_score


# ---------------------------------------------------------------------------
# Component normalization
# ---------------------------------------------------------------------------

def test_zero_inputs_give_zero_score():
    assert _compute_weighted_score(0, 0, 0, 0.0) == 0.0


def test_fairness_gap_zero_contributes_zero():
    score = _compute_weighted_score(
        fairness_gap=0, spacing_value=0, avoid_value=0, year_value=0.0
    )
    assert score == 0.0


def test_fairness_gap_at_max_diff_soft_contributes_full_weight():
    """A fairness_gap equal to MAX_DIFF_SOFT should contribute exactly
    FAIRNESS_GAP_WEIGHT (i.e., the normalized component is 1.0)."""
    score = _compute_weighted_score(
        fairness_gap=sm.MAX_DIFF_SOFT,
        spacing_value=0,
        avoid_value=0,
        year_value=0.0,
    )
    assert math.isclose(score, sm.FAIRNESS_GAP_WEIGHT)


def test_fairness_gap_clipped_above_max_diff_soft():
    """Above MAX_DIFF_SOFT, fairness_norm clips at 1.0 — so a gap of
    SOFT and a gap of SOFT+10 produce the same contribution."""
    at_threshold = _compute_weighted_score(sm.MAX_DIFF_SOFT, 0, 0, 0.0)
    way_above = _compute_weighted_score(sm.MAX_DIFF_SOFT + 10, 0, 0, 0.0)
    assert at_threshold == way_above


def test_fairness_gap_half_threshold_contributes_half_weight():
    if sm.MAX_DIFF_SOFT < 2:
        return  # not meaningful with tiny thresholds
    half_gap = sm.MAX_DIFF_SOFT // 2
    expected = sm.FAIRNESS_GAP_WEIGHT * (half_gap / sm.MAX_DIFF_SOFT)
    score = _compute_weighted_score(half_gap, 0, 0, 0.0)
    assert math.isclose(score, expected)


def test_spacing_tier_normalized_to_half_steps():
    """spacing_value ∈ {0, 1, 2} should normalize to {0, 0.5, 1.0}."""
    assert _compute_weighted_score(0, 0, 0, 0.0) == 0.0
    assert math.isclose(
        _compute_weighted_score(0, 1, 0, 0.0), sm.SPACING_WEIGHT * 0.5
    )
    assert math.isclose(
        _compute_weighted_score(0, 2, 0, 0.0), sm.SPACING_WEIGHT * 1.0
    )


def test_avoid_value_passthrough():
    assert _compute_weighted_score(0, 0, 1, 0.0) == sm.AVOID_WEIGHT


def test_year_value_passthrough():
    assert math.isclose(
        _compute_weighted_score(0, 0, 0, 0.5), sm.YEAR_BIAS_WEIGHT * 0.5
    )


# ---------------------------------------------------------------------------
# The actual fix: weights compose meaningfully now
# ---------------------------------------------------------------------------

def test_each_component_is_bounded_in_normal_regime():
    """In the normal operating regime (fairness_gap ≤ MAX_DIFF_SOFT),
    no single component contribution should exceed its own weight.
    This is the property that was broken before normalization."""
    max_score = _compute_weighted_score(
        fairness_gap=sm.MAX_DIFF_SOFT,
        spacing_value=2,
        avoid_value=1,
        year_value=1.0,
    )
    expected = (
        sm.FAIRNESS_GAP_WEIGHT
        + sm.SPACING_WEIGHT
        + sm.AVOID_WEIGHT
        + sm.YEAR_BIAS_WEIGHT
    )
    assert math.isclose(max_score, expected)


def test_doubling_a_weight_doubles_that_components_contribution():
    """If FAIRNESS_GAP_WEIGHT were doubled, fairness_gap=MAX_DIFF_SOFT
    would contribute exactly twice as much. We verify this without
    actually mutating module state by computing the ratio."""
    base = _compute_weighted_score(sm.MAX_DIFF_SOFT, 0, 0, 0.0)
    spacing_only = _compute_weighted_score(0, 2, 0, 0.0)

    # base / FAIRNESS_GAP_WEIGHT == 1.0 (the normalized fairness component)
    # spacing_only / SPACING_WEIGHT == 1.0 (the normalized spacing component)
    # So their ratio equals the weight ratio.
    expected_ratio = sm.FAIRNESS_GAP_WEIGHT / sm.SPACING_WEIGHT
    assert math.isclose(base / spacing_only, expected_ratio)
