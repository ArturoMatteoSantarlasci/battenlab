
export interface BattenInputs {
  testWeight: number;      // kg
  testLength: number;      // mm
  
  // Point measurements
  self14: number;          // mm
  self12: number;          // mm
  self34: number;          // mm
  
  weighted14: number;      // mm
  weighted12: number;      // mm
  weighted34: number;      // mm
}

export interface BattenResults {
  frontPercent: number;    // %
  backPercent: number;     // %
  camberPercent: number;   // %
  averageEi: number;       // N·m²
  deflection: number;      // Net mm (at center)
  draftPosition: number;   // % from front
}

export interface SavedProfile {
  id: string;
  name: string;
  inputs: BattenInputs;
  date: string;
}
