# monte-carlo-option-pricing-vba
Monte Carlo simulation in VBA for European option pricing

This project implements a Monte Carlo simulation in VBA to price a European call option.

The model is built in Excel and demonstrates how stochastic processes can be used for derivatives pricing in a practical environment.
Model

Underlying follows a Geometric Brownian Motion

Payoff: max(S_T - K, 0)

Discounted expectation under risk-neutral measure

Implementation
Language: VBA (Excel)
Simulation of price paths using random normal variables
Estimation of option price via averaging discounted payoffs


Adjustable parameters:
Spot price
Strike
Volatility
Interest rate
Number of simulations

Real-time pricing in Excel
Simple user interface

Results
The Monte Carlo price converges to the analytical Black-Scholes value as the number of simulations increases.

Project Structure
excel/ → Excel model (.xlsm)
src/ → VBA source code (.bas)

How to Use
Open the Excel file
Enable macros
Run the pricing function
