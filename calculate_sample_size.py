import math
from scipy.stats import norm

def get_sample_size(p=0.5, E=0.03, confidence=0.99, N=None):
    try:
        Z = norm.ppf(1 - (1 - confidence)/2)
        n0 = (Z**2 * p * (1 - p)) / (E**2)
        if N:
            n = (n0 * N) / (n0 + (N - 1))
            return math.ceil(n)
        return math.ceil(n0)
    except:
        return 0
