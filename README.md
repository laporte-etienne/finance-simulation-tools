
# 📈 Price Deck Confidence – Python Edition

## 🇬🇧 English Version

**Description:**  
This project is a Python-based reimplementation of a financial model initially developed in Excel/VBA. It simulates future commodity prices (Oil, Gas, etc.) using Monte Carlo methods and computes confidence levels to support risk managers in validating price assumptions (Price Decks).

**Context:**  
Originally developed during a Master’s thesis at IAE Gustave Eiffel in collaboration with Société Générale (RISQ/CIB), the project helps assess the plausibility of price deck proposals using stochastic modeling and data from Bloomberg.

## 💡 Features
- Monte Carlo simulations for forward price estimation
- Parametric assumptions (Normal, Log-normal, Vasicek)
- Confidence level calculator (50%, 70%, custom)
- Visual dashboard (via notebook or Streamlit – planned)
- Open-source alternative to legacy Excel-VBA tools

## 🛠 Project Structure
```
price-deck-confidence-python/
├── notebooks/           # Jupyter notebooks (exploration, simulation)
├── scripts/             # Python logic for Monte Carlo, pricing
├── tests/               # Unit tests
├── data/                # Example inputs (to be added)
├── outputs/             # Simulated results (charts, CSV, etc.)
├── README.md            # You're reading it
└── requirements.txt     # Python dependencies
```

## ▶️ How to Run

```bash
# Clone the repo
git clone https://github.com/yourname/finance-simulation-tools.git
cd finance-simulation-tools/price-deck-confidence-python

# (Optional) Create and activate a virtual env
python -m venv .venv
source .venv/bin/activate  # or .venv\Scripts\activate on Windows

# Install requirements
pip install -r requirements.txt

# Open notebook
jupyter notebook notebooks/price_deck_simulation.ipynb
```

## 📦 Dependencies (to add in `requirements.txt`)
```
numpy
pandas
matplotlib
scipy
jupyter
```

## ✅ Next Steps
- Implement Monte Carlo engine in `montecarlo.py`
- Add price input parser (Bloomberg .xls or simulated)
- Create CLI and dashboard (Streamlit)
- Unit testing with `pytest`

## 📄 License & Credits
Originally based on the thesis *"How to give confidence levels to price deck proposals?"*  
Author: **Etienne Laporte**  
Université Paris-Est Créteil – IAE Gustave Eiffel  
Société Générale – RISQ/CIB

---

## 🇫🇷 Version Française

**Description :**  
Ce projet est une réécriture en Python d’un outil de simulation développé initialement sous Excel/VBA. Il permet de générer des prix futurs via des simulations Monte Carlo et de calculer des niveaux de confiance afin de valider les hypothèses de prix proposées par les traders.

**Origine :**  
Projet issu du mémoire de Master 2 en Gestion d’Actifs à l’IAE Gustave Eiffel, en alternance chez Société Générale. Il visait à automatiser la validation des hypothèses de prix pour les commodities (Oil & Gas).
