"""Tests de sécurité : hachage des mots de passe et validation de la config."""
import hashlib
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import employee_manager as em


def test_password_scrypt_roundtrip():
    p = em.make_password("secret123")
    assert p["pw_algo"] == "scrypt"
    assert em.verify_pw("secret123", dict(p)) is True
    assert em.verify_pw("mauvais", dict(p)) is False


def test_password_no_password_set():
    assert em.verify_pw("peu importe", {}) is True


def test_password_legacy_sha256_still_verifies():
    salt = "abcd"
    h = hashlib.sha256((salt + "ancien").encode("utf-8")).hexdigest()
    cfg = {"pw_salt": salt, "pw_hash": h}  # pas de pw_algo -> sha256
    assert em.verify_pw("ancien", cfg) is True
    assert em.verify_pw("faux", cfg) is False


def test_password_hash_is_not_plaintext():
    p = em.make_password("monMotDePasse")
    assert "monMotDePasse" not in p["pw_hash"]
    assert len(p["pw_hash"]) == 64  # 32 octets en hex


def test_config_num_validation():
    assert em._num({"x": "5"}, "x", 26, 1, 31, int) == 5
    assert em._num({"x": "-3"}, "x", 26, 1, 31, int) == 26     # hors limites
    assert em._num({"x": "abc"}, "x", 26, 1, 31, int) == 26    # non numérique
    assert em._num({}, "x", 26, 1, 31, int) == 26              # absent
    assert em._num({"x": 0.5}, "x", 0.0448, 0, 1) == 0.5
    assert em._num({"x": 50}, "x", 0.0448, 0, 1) == 0.0448     # taux > 1


def test_bad_config_falls_back_to_defaults():
    em.apply_config_settings({"taux_cnss": -1, "jours_ouvrables": 0})
    assert em.TAUX_CNSS == 0.0448
    assert em.JOURS_OUVRABLES == 26
    # remet des valeurs valides pour ne pas impacter d'autres tests
    em.apply_config_settings({})
