# Binomial Option Pricer — Excel/VBA

A fully functional options pricing tool built in Excel with VBA, implementing the **Cox-Ross-Rubinstein (CRR) binomial tree model**.

---

## Overview

This pricer values European and American options (calls and puts) on a single underlying asset. It was built from scratch without any external library — only native Excel formulas and VBA.

The model discretizes time into N steps and uses backward induction to derive the fair value of an option under risk-neutral pricing, consistent with no-arbitrage theory.

---

## Features

- **Cox-Ross-Rubinstein binomial tree** calibrated to match Black-Scholes diffusion (u = e^σ√Δt)
- **European & American** options supported — early exercise logic implemented at every node
- **Call & Put** pricing
- **Visual binomial tree** — stock price tree and option value tree displayed step by step
- **Greeks** computed via bump & reprice: Delta, Gamma, Vega, Theta
- **Black-Scholes benchmark** — European prices compared to closed-form BS formula
- **Convergence analysis** — price vs. N chart showing convergence to Black-Scholes as N → ∞

---

## Model Parameters

| Parameter | Description |
|---|---|
| S | Current underlying price |
| K | Strike price |
| T | Time to maturity (in years) |
| r | Risk-free rate (continuous compounding) |
| σ | Implied volatility |
| N | Number of time steps |
| Type | Call (+1) or Put (-1) |
| Type_option | European (0) or American (1) |

---

## CRR Calibration

The up/down factors and risk-neutral probability are derived as:

```
u = e^(σ√Δt)
d = 1/u
p = (e^(rΔt) − d) / (u − d)
```

This ensures the binomial tree reproduces the correct drift and variance of a geometric Brownian motion, converging to Black-Scholes as N → ∞.

---

## Pricing Logic

**Forward pass** — build the stock price tree:
```
S(j, i) = S × u^(j−i) × d^i
```
where j = time step, i = number of down moves.

**Backward induction** — starting from terminal payoffs:
```
V(j, i) = DF × [p × V(j+1, i) + (1−p) × V(j+1, i+1)]
```
For American options, at each node:
```
V(j, i) = max(Continuation, Intrinsic Value)
```

---

## File Structure

```
binomial-option-pricer.xlsm
├── Input          → Parameters & CRR computed values
├── Tree           → Visual binomial tree (stock prices)
├── Output         → Option price, Greeks, Black-Scholes comparison
└── Convergence    → Price vs. N convergence chart
```

---

## How to Use

1. Open `binomial-option-pricer.xlsm` and **enable macros**
2. Fill in your parameters in the **Input** sheet
3. Click **Run** — the tree is built and the option is priced instantly
4. Check **Output** for price and Greeks
5. Check **Convergence** to see how the binomial price converges to Black-Scholes

---

## Why This Matters

Black-Scholes has no closed-form solution for American options (except European puts via put-call parity). The binomial tree handles early exercise naturally — at each node, the holder compares the continuation value against immediate exercise. This is the same logic used in practice for pricing American equity options and callable bonds.

---

## Tech Stack

- Microsoft Excel (.xlsm)
- VBA (Visual Basic for Applications)
- No external dependencies

---

## Output preview

![Input](Output_screenshots/Input.png)
![Binomial Tree](Output_screenshots/Binomial_tree.png)
![Output - price](Output_screenshots/Output_price.png)




