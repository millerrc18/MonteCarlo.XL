# Distribution Guide

MonteCarlo.XL supports 15 distributions in the task pane. Worksheet UDFs are available for all except `Discrete`, which is currently configured through the task pane.

When a workbook is not being simulated, `MC.*` formulas return an expected or representative value so the model remains readable and usable. During simulation, the add-in samples the configured inputs and recalculates the selected output cells.

## Quick Selection Guide

| Distribution | Formula | Use when | Example |
| --- | --- | --- | --- |
| Normal | `MC.Normal(mean, stdDev)` | Values vary symmetrically around a mean and extreme values are possible but rare. | Monthly demand around 10,000 units with historical standard deviation of 1,200. |
| Triangular | `MC.Triangular(min, mode, max)` | You only know a low, most likely, and high estimate. | A task takes at least 3 days, most likely 5, and at most 9. |
| PERT | `MC.PERT(min, mode, max)` | You have low/likely/high estimates but want a smoother curve with less weight on extremes than Triangular. | Project cost estimate from subject matter experts. |
| Lognormal | `MC.Lognormal(mu, sigma)` | Values are positive and right-skewed; large upside outliers are plausible. | Sales deal size, repair cost, or commodity price multiplier. |
| Uniform | `MC.Uniform(min, max)` | Any value in a bounded range is equally plausible. | Unknown launch date delay between 2 and 6 weeks with no better information. |
| Discrete | Task pane only | A small set of exact outcomes has known probabilities. | Scenario A/B/C adoption rates of 5%, 12%, and 25%. |
| Beta | `MC.Beta(alpha, beta)` | You need a bounded 0-to-1 rate or percentage with flexible shape. | Conversion rate, defect rate, or completion percentage. |
| Weibull | `MC.Weibull(shape, scale)` | Time-to-failure or life data where risk changes over time. | Equipment failure time or component replacement interval. |
| Exponential | `MC.Exponential(rate)` | Waiting time until an event with a constant arrival rate. | Time until next support ticket or failure under constant hazard. |
| Poisson | `MC.Poisson(lambda)` | Count of independent events in a fixed period. | Number of defects per batch or customer arrivals per hour. |
| Gamma | `MC.Gamma(shape, rate)` | Positive continuous totals built from many waiting-time-like effects. | Total processing time, claim severity, or rainfall amount. |
| Logistic | `MC.Logistic(mu, s)` | Symmetric uncertainty like Normal but with heavier tails. | Forecast error where large misses happen more often than Normal implies. |
| GEV | `MC.GEV(mu, sigma, xi)` | Block maxima or minima, especially tail-risk modeling. | Annual maximum flood level, peak daily demand, or worst annual loss. |
| Binomial | `MC.Binomial(n, p)` | Number of successes in a fixed number of independent trials. | How many of 20 prospects convert when each has a 35% chance. |
| Geometric | `MC.Geometric(p)` | Number of trials until the first success. | Number of sales calls until first win. |

## Parameter Notes

### Normal

Use `mean` for the center and `stdDev` for spread. Normal distributions can generate values below zero, so avoid them for quantities that cannot be negative unless the chance is negligible or your model clamps the result.

Example:

```excel
=MC.Normal(100000, 15000)
```

Good for forecast errors, mature product demand, and measurement noise.

### Triangular

Use `min`, `mode`, and `max`. This is often the most approachable distribution when the only available inputs are expert estimates.

Example:

```excel
=MC.Triangular(80, 100, 140)
```

Good for early estimates, schedule tasks, and vendor quotes.

### PERT

PERT also uses `min`, `mode`, and `max`, but it puts less weight on the extremes than Triangular. Use it when experts provide three-point estimates and the most likely value should dominate the curve.

Example:

```excel
=MC.PERT(50000, 75000, 125000)
```

Good for project management, cost estimates, and uncertain durations.

### Lognormal

Use `mu` and `sigma` for the underlying normal distribution, not the arithmetic mean and standard deviation of the final positive values. Lognormal is useful when values multiply together or grow by percentages.

Example:

```excel
=MC.Lognormal(10.5, 0.35)
```

Good for positive right-skewed quantities such as deal size, price changes, and claim severity.

### Uniform

Use `min` and `max`. Uniform is simple, but it says every value in the range is equally likely. Use it only when you truly do not have a better view of the likely center.

Example:

```excel
=MC.Uniform(0.02, 0.08)
```

Good for rough sensitivity ranges and placeholder assumptions.

### Discrete

Use the task pane to enter value/probability pairs. Probabilities should sum to 1. There is not currently an `MC.Discrete` worksheet formula.

Example outcomes:

| Value | Probability |
| --- | --- |
| 0.05 | 0.25 |
| 0.12 | 0.50 |
| 0.25 | 0.25 |

Good for scenario analysis, yes/no outcomes with custom payoffs, and categorical business cases.

### Beta

Use `alpha` and `beta`. Beta stays between 0 and 1, so it is a natural fit for rates, shares, and percentages.

Example:

```excel
=MC.Beta(8, 32)
```

Good for conversion rates, churn rates, market share, defect rates, and completion percentages.

### Weibull

Use `shape` and `scale`. Shape controls whether failure risk decreases, stays roughly constant, or increases over time.

Example:

```excel
=MC.Weibull(1.6, 1200)
```

Good for reliability modeling, component life, maintenance planning, and survival time.

### Exponential

Use `rate`, where the mean waiting time is `1 / rate`. Exponential assumes a constant hazard rate, so the probability of the event does not depend on how long you have already waited.

Example:

```excel
=MC.Exponential(0.2)
```

Good for time until the next independent event, simple queueing assumptions, and constant failure-rate models.

### Poisson

Use `lambda` as the expected event count in the chosen period.

Example:

```excel
=MC.Poisson(4.5)
```

Good for arrivals per hour, defects per unit, incidents per month, and other non-negative count data.

### Gamma

Use `shape` and `rate`. The mean is `shape / rate`. Gamma is positive and flexible, making it useful for skewed continuous amounts.

Example:

```excel
=MC.Gamma(3, 0.4)
```

Good for total processing time, aggregate waiting time, insurance claim severity, and rainfall.

### Logistic

Use `mu` for location and `s` for scale. Logistic is symmetric like Normal but has heavier tails.

Example:

```excel
=MC.Logistic(0, 1.5)
```

Good for forecast errors where misses are centered around zero but large misses happen more often than a Normal curve suggests.

### GEV

Use `mu`, `sigma`, and `xi`. `xi` controls tail behavior: near 0 is Gumbel-like, positive is heavier-tailed, and negative is bounded.

Example:

```excel
=MC.GEV(100, 12, 0.1)
```

Good for annual maximum demand, peak load, flood level, extreme wind, or worst annual loss.

### Binomial

Use `n` for the number of trials and `p` for the probability of success per trial.

Example:

```excel
=MC.Binomial(20, 0.35)
```

Good for wins out of prospects, defects in a batch, defaults in a portfolio segment, and pass/fail experiments.

### Geometric

Use `p` for the probability of success on each trial. The sampled value is the number of trials until the first success.

Example:

```excel
=MC.Geometric(0.25)
```

Good for number of calls until first sale, attempts until approval, or batches until first defect.

## Common Modeling Choices

Use `Triangular` or `PERT` when the input comes from expert judgment. Choose PERT when the most likely value should have more influence.

Use `Beta` for uncertain percentages that must stay between 0 and 1.

Use `Lognormal` or `Gamma` for positive skewed business amounts.

Use `Poisson`, `Binomial`, or `Geometric` for counts. Choose Poisson for event counts over time, Binomial for successes out of a fixed number of attempts, and Geometric for attempts until the first success.

Use `GEV` only when the model is about extremes, not ordinary values.

## Practical Advice

Match the distribution to the business process first, then tune parameters. A precise-looking distribution with weak assumptions can be less useful than a simple distribution with clear reasoning.

Document units and time periods in cell labels. For example, `Customer arrivals per hour` is much clearer than `Arrivals`.

Use a fixed random seed when comparing model changes. Remove the seed when producing final exploratory simulation results.
