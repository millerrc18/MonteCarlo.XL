/// <reference types="office-js" />
import { BRIDGE_VERSION } from './types';
import { MonteCarloBridgeClient } from './bridge';

function invalidValue(): CustomFunctions.Error {
  return new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue);
}

async function evaluate(functionName: string, args: number[]): Promise<number | CustomFunctions.Error> {
  const response = await MonteCarloBridgeClient.evaluateFormula({
    version: BRIDGE_VERSION,
    functionName,
    arguments: args,
  });

  return response.isValid && typeof response.value === 'number'
    ? response.value
    : invalidValue();
}

export async function normal(mean: number, stdDev: number): Promise<number | CustomFunctions.Error> {
  return evaluate('Normal', [mean, stdDev]);
}

export async function triangular(min: number, mode: number, max: number): Promise<number | CustomFunctions.Error> {
  return evaluate('Triangular', [min, mode, max]);
}

export async function pert(min: number, mode: number, max: number): Promise<number | CustomFunctions.Error> {
  return evaluate('PERT', [min, mode, max]);
}

export async function lognormal(mu: number, sigma: number): Promise<number | CustomFunctions.Error> {
  return evaluate('Lognormal', [mu, sigma]);
}

export async function uniform(min: number, max: number): Promise<number | CustomFunctions.Error> {
  return evaluate('Uniform', [min, max]);
}

export async function beta(alpha: number, betaValue: number): Promise<number | CustomFunctions.Error> {
  return evaluate('Beta', [alpha, betaValue]);
}

export async function weibull(shape: number, scale: number): Promise<number | CustomFunctions.Error> {
  return evaluate('Weibull', [shape, scale]);
}

export async function exponential(rate: number): Promise<number | CustomFunctions.Error> {
  return evaluate('Exponential', [rate]);
}

export async function poisson(lambda: number): Promise<number | CustomFunctions.Error> {
  return evaluate('Poisson', [lambda]);
}

export async function gamma(shape: number, rate: number): Promise<number | CustomFunctions.Error> {
  return evaluate('Gamma', [shape, rate]);
}

export async function logistic(mu: number, s: number): Promise<number | CustomFunctions.Error> {
  return evaluate('Logistic', [mu, s]);
}

export async function gev(mu: number, sigma: number, xi: number): Promise<number | CustomFunctions.Error> {
  return evaluate('GEV', [mu, sigma, xi]);
}

export async function binomial(n: number, p: number): Promise<number | CustomFunctions.Error> {
  return evaluate('Binomial', [n, p]);
}

export async function geometric(p: number): Promise<number | CustomFunctions.Error> {
  return evaluate('Geometric', [p]);
}

CustomFunctions.associate('MC.NORMAL', normal);
CustomFunctions.associate('MC.TRIANGULAR', triangular);
CustomFunctions.associate('MC.PERT', pert);
CustomFunctions.associate('MC.LOGNORMAL', lognormal);
CustomFunctions.associate('MC.UNIFORM', uniform);
CustomFunctions.associate('MC.BETA', beta);
CustomFunctions.associate('MC.WEIBULL', weibull);
CustomFunctions.associate('MC.EXPONENTIAL', exponential);
CustomFunctions.associate('MC.POISSON', poisson);
CustomFunctions.associate('MC.GAMMA', gamma);
CustomFunctions.associate('MC.LOGISTIC', logistic);
CustomFunctions.associate('MC.GEV', gev);
CustomFunctions.associate('MC.BINOMIAL', binomial);
CustomFunctions.associate('MC.GEOMETRIC', geometric);
