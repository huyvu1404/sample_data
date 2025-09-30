import math

def get_sample_size(p=0.5, E=0.03, confidence=0.99, N=None):
    z_values = {0.9: 1.645, 0.95: 1.96, 0.99: 2.576}
    Z = z_values.get(confidence, 1.96)
    n0 = (Z**2 * p * (1 - p)) / (E**2)
    if N:
        n = (n0 * N) / (n0 + (N - 1))
        return math.ceil(n)
    return math.ceil(n0)

