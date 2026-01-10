
import React, { useState } from 'react';
import { solveFEM, nelderMead } from '../services/femSolver';
import { FEMCalibrationInputs, FEMCalibrationResults } from '../types';

const FEMCalibrator: React.FC = () => {
  const [numSegments, setNumSegments] = useState<number>(4);
  const [inputs, setInputs] = useState<FEMCalibrationInputs>({
    weightKg: 5,
    lengthMm: 3000,
    deflectionsMm: [40, 55, 42],
  });

  const [results, setResults] = useState<FEMCalibrationResults | null>(null);
  const [isCalculating, setIsCalculating] = useState(false);

  const handleRunCalibration = () => {
    setIsCalculating(true);
    setTimeout(() => {
      const P_newton = inputs.weightKg * 9.81;
      const L_meters = inputs.lengthMm / 1000;
      const v_mis_meters = inputs.deflectionsMm.map(d => d / 1000);

      // Guess iniziale
      const initialGuessVal = (P_newton * Math.pow(L_meters, 3)) / (48 * (v_mis_meters[Math.floor(v_mis_meters.length/2)] || 0.01));
      const x0 = new Array(numSegments).fill(initialGuessVal);

      const objective = (eiValues: number[]) => {
        if (eiValues.some(v => v <= 0)) return 1e10;
        const v_calc = solveFEM(eiValues, P_newton, L_meters);
        let error = 0;
        // Se 4 pezzi, usiamo i 3 punti. Se 2 pezzi, usiamo 1 punto al centro.
        const pointsToCompare = Math.min(v_calc.length, v_mis_meters.length);
        for (let i = 0; i < pointsToCompare; i++) {
          error += Math.pow(v_calc[i] - v_mis_meters[i], 2);
        }
        return error;
      };

      const optimized = nelderMead(objective, x0);
      setResults({
        eiSegments: optimized.x,
        eiAverage: optimized.x.reduce((a, b) => a + b, 0) / numSegments,
        error: optimized.fx
      });
      setIsCalculating(false);
    }, 100);
  };

  return (
    <div className="bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
      <div className="bg-slate-900 p-6 text-white flex items-center justify-between">
        <div>
          <h2 className="text-xl font-black uppercase tracking-tighter">Analisi Avanzata FEM (Problema Inverso)</h2>
          <p className="text-xs text-slate-400 font-bold tracking-widest">Ottimizzazione EI da Deflessioni Misurate</p>
        </div>
        <div className="flex bg-white/10 p-1 rounded-lg">
          {[2, 3, 4].map(n => (
            <button 
              key={n}
              onClick={() => {
                setNumSegments(n);
                setResults(null);
                // Aggiorna numero di input deflessioni se necessario
                const newDefs = n === 2 ? [inputs.deflectionsMm[1]] : [40, 55, 42];
                setInputs({...inputs, deflectionsMm: newDefs});
              }}
              className={`px-3 py-1 text-[10px] font-black uppercase rounded-md transition-all ${numSegments === n ? 'bg-white text-slate-900' : 'text-white/50 hover:text-white'}`}
            >
              {n} Pezzi
            </button>
          ))}
        </div>
      </div>

      <div className="p-8 grid grid-cols-1 lg:grid-cols-2 gap-12">
        <div className="space-y-6">
          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1">
              <label className="text-[10px] font-black text-slate-400 uppercase">Peso Test (kg)</label>
              <input type="number" value={inputs.weightKg} onChange={e => setInputs({...inputs, weightKg: parseFloat(e.target.value) || 0})} className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl outline-none font-bold" />
            </div>
            <div className="space-y-1">
              <label className="text-[10px] font-black text-slate-400 uppercase">Lunghezza L (mm)</label>
              <input type="number" value={inputs.lengthMm} onChange={e => setInputs({...inputs, lengthMm: parseFloat(e.target.value) || 0})} className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl outline-none font-bold" />
            </div>
          </div>

          <div className="space-y-4">
            <h4 className="text-[10px] font-black text-slate-400 uppercase border-b border-slate-100 pb-2">Deflessioni Misurate (mm)</h4>
            <div className={`grid gap-3 ${numSegments === 2 ? 'grid-cols-1' : 'grid-cols-3'}`}>
              {numSegments === 2 ? (
                <div className="space-y-1 text-center">
                  <label className="text-[9px] font-bold text-slate-400">L/2</label>
                  <input type="number" value={inputs.deflectionsMm[0]} onChange={e => setInputs({...inputs, deflectionsMm: [parseFloat(e.target.value) || 0]})} className="w-full px-2 py-3 bg-slate-50 border border-slate-200 rounded-xl text-center font-bold" />
                </div>
              ) : (
                [1, 2, 3].map((pos, idx) => (
                  <div key={pos} className="space-y-1 text-center">
                    <label className="text-[9px] font-bold text-slate-400">{pos === 1 ? 'L/4' : pos === 2 ? 'L/2' : '3L/4'}</label>
                    <input type="number" value={inputs.deflectionsMm[idx]} onChange={e => {
                      const newDefs = [...inputs.deflectionsMm];
                      newDefs[idx] = parseFloat(e.target.value) || 0;
                      setInputs({...inputs, deflectionsMm: newDefs});
                    }} className="w-full px-2 py-3 bg-slate-50 border border-slate-200 rounded-xl text-center font-bold" />
                  </div>
                ))
              )}
            </div>
          </div>

          <button onClick={handleRunCalibration} disabled={isCalculating} className="w-full py-4 bg-[#A12B2B] text-white rounded-xl font-black uppercase tracking-widest text-sm shadow-xl shadow-red-900/20 active:scale-95 transition-all">
            {isCalculating ? 'Calibrazione...' : 'Avvia Ottimizzazione'}
          </button>
        </div>

        <div>
          {!results ? (
            <div className="h-full flex flex-col items-center justify-center text-slate-300 border-2 border-dashed border-slate-100 rounded-2xl p-8">
              <p className="text-xs font-bold uppercase tracking-widest text-center">Definisci i parametri per calcolare la rigidezza locale</p>
            </div>
          ) : (
            <div className="space-y-6">
               <div className="bg-slate-50 rounded-2xl p-6 border border-slate-200">
                  <h4 className="text-[10px] font-black text-slate-400 uppercase mb-4 tracking-widest">Rigidezza dei Pezzi (EI)</h4>
                  <div className="space-y-4">
                    {results.eiSegments.map((ei, i) => (
                      <div key={i} className="space-y-1">
                        <div className="flex justify-between text-[10px] font-bold uppercase">
                          <span className="text-slate-400">Pezzo {i+1}</span>
                          <span className="text-slate-700">{ei.toFixed(2)} Nm²</span>
                        </div>
                        <div className="h-2 bg-slate-200 rounded-full overflow-hidden">
                          <div className="h-full bg-[#A12B2B]" style={{ width: `${Math.min(100, (ei / (results.eiAverage * 1.5)) * 100)}%` }}></div>
                        </div>
                      </div>
                    ))}
                  </div>
               </div>
               <div className="bg-[#A12B2B] p-8 rounded-2xl text-white shadow-2xl flex flex-col items-center justify-center text-center">
                  <p className="text-xs font-bold uppercase opacity-60 tracking-[0.2em] mb-2">Rigidezza Media Flessionale Equivalente</p>
                  <div className="flex items-baseline gap-2">
                    <span className="text-6xl font-black">{results.eiAverage.toFixed(3)}</span>
                    <span className="text-sm font-bold opacity-60">N·m²</span>
                  </div>
               </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default FEMCalibrator;
