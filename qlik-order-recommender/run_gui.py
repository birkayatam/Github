"""PyInstaller giris noktasi.

gui.py dogrudan hedeflenirse (`pyinstaller ... src/qlik_order_recommender/gui.py`)
PyInstaller onu bagimsiz bir script olarak calistirir, bu da icindeki
`from .order_pipeline import ...` gibi goreli import'lari "attempted
relative import with no known parent package" hatasiyla bozar. Bu kucuk
sarmalayici, gui.py'yi paketin (`qlik_order_recommender`) bir parcasi
olarak normal sekilde import eder, boylece goreli import'lar calisir.
PyInstaller bu dosyayi hedef almali, gui.py'yi degil (bkz. README)."""
from qlik_order_recommender.gui import main

if __name__ == "__main__":
    main()
