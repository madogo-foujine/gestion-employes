"""Tests unitaires du moteur de paie (fonctions pures, sans interface)."""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import employee_manager as em


def test_to_float():
    assert em.to_float("") == 0.0
    assert em.to_float(None) == 0.0
    assert em.to_float("abc") == 0.0
    assert em.to_float("1 234,5") == 1234.5
    assert em.to_float("2000") == 2000.0


def test_fmt_money():
    assert em.fmt_money(1234.5) == "1 234.50 DH"
    assert em.fmt_money(0) == "0.00 DH"


def test_initials():
    assert em.initials("Ahmed Bennani") == "AB"
    assert em.initials("Ahmed") == "AH"
    assert em.initials("") == "?"


def test_parse_date():
    assert em.parse_date("2022-03-15") == dt.date(2022, 3, 15)
    assert em.parse_date("15/03/2022") == dt.date(2022, 3, 15)
    assert em.parse_date("") is None
    assert em.parse_date(None) is None


def test_anciennete():
    assert em.compute_anciennete("") == ""
    assert "mois" in em.compute_anciennete("2020-01-01")


def test_ir_brackets_default():
    assert round(em.compute_ir_brut(6000), 2) == 366.67
    assert em.compute_ir_brut(3000) == 0.0


def test_payroll_basic():
    p = em.compute_payroll({"salaire_base": "6000"})
    assert p["brut"] == 6000.0
    assert p["cnss"] == 268.8
    assert p["amo"] == 135.6
    assert p["net"] == 5579.37


def test_cnss_plafond():
    # CNSS plafonnée à 6000 -> 268.80 même pour un salaire plus élevé
    assert em.compute_payroll({"salaire_base": "10000"})["cnss"] == 268.8


def test_advance_deduction():
    p = em.compute_payroll({"salaire_base": "6000", "retenue_avance": "500"})
    assert p["avance"] == 500.0
    assert p["net"] == 5079.37


def test_absence_deduction():
    p = em.compute_payroll({"salaire_base": "6000", "jours_absence": "2"})
    assert round(p["ded_absence"], 2) == 461.54


def test_total_retenues_consistency():
    p = em.compute_payroll({
        "salaire_base": "8000", "primes": "1000", "retenues": "200",
        "retenue_avance": "300", "jours_absence": "1", "personnes_charge": "2",
    })
    expected = round(p["cnss"] + p["amo"] + p["ir"] + p["autres"]
                     + p["avance"] + p["ded_absence"], 2)
    assert p["total_retenues"] == expected
    assert p["net"] == round(p["brut"] - p["total_retenues"], 2)


def test_salaire_net_helper():
    rec = {"salaire_base": "7000"}
    assert em.compute_salaire_net(rec) == em.compute_payroll(rec)["net"]
