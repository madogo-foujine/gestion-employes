"""Tests des fonctions de temps (heures travaillées, retard, heures sup)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import employee_manager as em


def test_parse_time():
    assert em.parse_time("09:30") == 570
    assert em.parse_time("9h30") == 570
    assert em.parse_time("00:00") == 0
    assert em.parse_time("") is None
    assert em.parse_time("99:99") is None
    assert em.parse_time("abc") is None


def test_compute_day_hours_standard():
    assert em.compute_day_hours("09:00", "17:00") == (8.0, 0, 0.0)


def test_compute_day_hours_overtime_and_lateness():
    h, retard, sup = em.compute_day_hours("09:30", "19:00")
    assert h == 9.5
    assert retard == 30
    assert sup == 1.5


def test_compute_day_hours_invalid():
    assert em.compute_day_hours("", "") == (0.0, 0, 0.0)
    assert em.compute_day_hours("18:00", "09:00") == (0.0, 0, 0.0)  # sortie < entrée
