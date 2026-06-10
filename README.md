# FC Entlebuch Statistik App v1.3.2

Änderung in dieser Version:
- In der Datenbasis ist nun jede Kombination aus Spiel und Spieler vorhanden.
- Wenn ein Spieler in einem Spiel nicht eingesetzt wurde, steht in `SpielerSpielStatistik` der Status `Ohne Einsatz` und `Minuten = 0`.
- Das Spielerprofil zeigt dadurch auch Spiele ohne Einsatz an.
- In den Spieldetails werden eingesetzte Spieler direkt angezeigt; nicht eingesetzte Spieler sind in einem aufklappbaren Bereich sichtbar.

Start lokal:

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt
python3 -m streamlit run app.py
```


Neu in Version 1.3.3: Im Spielerprofil gibt es einen Vergleich Total Meisterschaft / Vorrunde / Rückrunde / Veränderung für Einsatzminuten, Startelf-Einsätze, Einwechslungen, Tore und Assists.
