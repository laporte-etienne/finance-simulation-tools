import numpy as np
from montecarlo import simulate_price_paths, compute_confidence_levels

def test_simulation_output_length():
    S0, mu, sigma, T, n = 100, 0.05, 0.2, 1, 5000
    prices = simulate_price_paths(S0, mu, sigma, T, n)
    assert len(prices) == n, "Length of simulation output should match n_simulations"

def test_confidence_levels_exist():
    S0, mu, sigma, T, n = 100, 0.05, 0.2, 1, 5000
    prices = simulate_price_paths(S0, mu, sigma, T, n)
    levels = compute_confidence_levels(prices, [0.5, 0.7])
    assert "50%" in levels and "70%" in levels, "Confidence levels 50% and 70% must be present"

def test_confidence_logic():
    S0, mu, sigma, T, n = 100, 0.05, 0.2, 1, 10000
    prices = simulate_price_paths(S0, mu, sigma, T, n)
    levels = compute_confidence_levels(prices, [0.5, 0.7])
    assert levels["70%"] < levels["50%"], "70% confidence level should be lower than 50%"
