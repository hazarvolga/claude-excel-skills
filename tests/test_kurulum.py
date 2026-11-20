"""kur.ps1 Windows PowerShell 5.1 uyumluluğu."""
from conftest import DEPO


def test_kur_ps1_utf8_bom_ile_kayitli():
    assert (DEPO / "kur.ps1").read_bytes()[:3] == b"\xef\xbb\xbf"
