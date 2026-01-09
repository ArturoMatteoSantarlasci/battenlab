
import { BattenInputs, BattenResults } from '../types';

/**
 * Calculates sailboat batten characteristics.
 * 
 * Formulas:
 * 1. Net Deflection (δ): Weighted - Self Weighted
 * 2. Front Bend %: (δ_1/4 / δ_1/2) * 100
 * 3. Back Bend %: (δ_3/4 / δ_1/2) * 100
 * 4. Flexural Rigidity (EI): (F * L^3) / (48 * δ_1/2)
 * 5. Camber %: (δ_1/2 / L) * 100
 */
export const calculateBattenBehavior = (inputs: BattenInputs): BattenResults => {
  const { 
    testWeight, testLength, 
    self14, self12, self34,
    weighted14, weighted12, weighted34 
  } = inputs;

  // 1. Calculate Net Deflections (Delta)
  const d14 = Math.max(0, weighted14 - self14);
  const d12 = Math.max(0.1, weighted12 - self12); // Prevent zero to avoid division by zero
  const d34 = Math.max(0, weighted34 - self34);

  // 2. Constants and Force calculation
  const L = testLength; // mm
  const Lm = L / 1000; // m
  const F_newtons = testWeight * 9.80665;

  // 3. Camber Percentage
  const camberPercent = L > 0 ? (d12 / L) * 100 : 0;

  // 4. EI (Flexural Rigidity)
  const ei = d12 > 0 ? (F_newtons * Math.pow(Lm, 3)) / (48 * (d12 / 1000)) : 0;

  // 5. Bend Profile (User specific formulas)
  // front bend: (deflection_1/4 / deflection_1/2) * 100
  // back bend: (deflection_3/4 / deflection_1/2) * 100
  const frontPercent = (d14 / d12) * 100;
  const backPercent = (d34 / d12) * 100;
  
  // Estimate draft position based on asymmetry (for visualization indicator)
  const draftPos = 50 - ((d14 - d34) / d12 * 10);

  return {
    frontPercent: Number(frontPercent.toFixed(1)),
    backPercent: Number(backPercent.toFixed(1)),
    camberPercent: Number(camberPercent.toFixed(2)),
    averageEi: Number(ei.toFixed(3)),
    deflection: d12,
    draftPosition: Number(draftPos.toFixed(1))
  };
};

/**
 * Graph Interpolation Method: 4th-degree Lagrange Polynomial.
 * We use 5 control points: (0,0), (0.25L, d14), (0.5L, d12), (0.75L, d34), and (L,0).
 * This creates a smooth curve representing the physical beam bending profile.
 */
export const generateCurvePoints = (inputs: BattenInputs) => {
  const { testLength, self14, self12, self34, weighted14, weighted12, weighted34 } = inputs;
  
  const d14 = Math.max(0, weighted14 - self14);
  const d12 = Math.max(0, weighted12 - self12);
  const d34 = Math.max(0, weighted34 - self34);

  const points = [];
  const steps = 60;

  for (let i = 0; i <= steps; i++) {
    const t = i / steps;
    const x = t * testLength;
    
    // Basis functions for Lagrange polynomial
    const l0 = (t - 0.25) * (t - 0.5) * (t - 0.75) * (t - 1.0) / ( (0 - 0.25) * (0 - 0.5) * (0 - 0.75) * (0 - 1.0) );
    const l1 = (t - 0) * (t - 0.5) * (t - 0.75) * (t - 1.0) / ( (0.25 - 0) * (0.25 - 0.5) * (0.25 - 0.75) * (0.25 - 1.0) );
    const l2 = (t - 0) * (t - 0.25) * (t - 0.75) * (t - 1.0) / ( (0.5 - 0) * (0.5 - 0.25) * (0.5 - 0.75) * (0.5 - 1.0) );
    const l3 = (t - 0) * (t - 0.25) * (t - 0.5) * (t - 1.0) / ( (0.75 - 0) * (0.75 - 0.25) * (0.75 - 0.5) * (0.75 - 1.0) );
    const l4 = (t - 0) * (t - 0.25) * (t - 0.5) * (t - 0.75) / ( (1.0 - 0) * (1.0 - 0.25) * (1.0 - 0.5) * (1.0 - 0.75) );

    const y = (0 * l0) + (d14 * l1) + (d12 * l2) + (d34 * l3) + (0 * l4);

    points.push({
      x: Math.round(x),
      y: Number(y.toFixed(2))
    });
  }
  return points;
};
