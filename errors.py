"""Typed exceptions for the Call Schedule Creator.

All inherit from ValueError so pre-existing `except ValueError` sites and tests
keep working. The CLI and GUI entry points catch ScheduleError (the common base)
to show friendly prose messages instead of a raw traceback.
"""
from __future__ import annotations


class ScheduleError(ValueError):
    """Base class for all scheduler-originated errors meant for end users."""


class ConfigError(ScheduleError):
    """Invalid or inconsistent configuration (config.yaml, GUI settings)."""


class DataValidationError(ScheduleError):
    """Input data (flow / rules / no_call / holidays / completed) fails validation."""
