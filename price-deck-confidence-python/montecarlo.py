import numpy as np
import matplotlib.pyplot as plt

def simulate_price_paths(S0, mu, sigma, T, n_simulations):
    """ Simulate price paths using a geometric Brownian motion model """
    np.random.seed(42)
    Z = np.random.normal(size=n_simulations)
    ST = S0 * np.exp((mu - 0.5 * sigma ** 2) * T + sigma * np.sqrt(T) * Z)
    return ST

def compute_confidence_levels(simulated_prices, confidence_levels=[0.5, 0.7]):
    """ Compute price quantiles for given confidence levels """
    sorted_prices = np.sort(simulated_prices)
    results = {}
    for c in confidence_levels:
        index = int((1 - c) * len(sorted_prices))
        results[f"{int(c * 100)}%"] = sorted_prices[index]
    return results

if __name__ == "__main__":
    S0 = 90     # Initial price
    mu = 0.05   # Expected return
    sigma = 0.2 # Volatility
    T = 1       # 1 year
    n_simulations = 10000

    simulated = simulate_price_paths(S0, mu, sigma, T, n_simulations)
    conf_levels = compute_confidence_levels(simulated)

    print("Confidence levels (price thresholds):")
    for level, price in conf_levels.items():
        print(f"{level}: {price:.2f} USD")

    # Plot the distribution
    plt.hist(simulated, bins=50, alpha=0.7, color='skyblue', edgecolor='black')
    plt.axvline(conf_levels['50%'], color='red', linestyle='--', label='50% Confidence')
    plt.axvline(conf_levels['70%'], color='orange', linestyle='--', label='70% Confidence')
    plt.title('Simulated Price Distribution (T = 1 year)')
    plt.xlabel('Price')
    plt.ylabel('Frequency')
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    plt.show()
