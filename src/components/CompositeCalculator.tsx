import React, { useState, useMemo } from 'react';
import { BattenSegment } from '../types';
import { calculateEquivalentEI } from '../services/calculator';

const CompositeCalculator: React.FC = () => {
  const [numSegments, setNumSegments] = useState<number>(3);
  const [segments, setSegments] = useState<BattenSegment[]>([
    { length: 1000, ei: 2.5 },
    { length: 1000, ei: 3.5 },
    { length: 1000, ei: 2.0 },
    { length: 0, ei: 0 },
  ]);
  const [showFormulaInfo, setShowFormulaInfo] = useState(false);

  const result = useMemo(() => calculateEquivalentEI(segments.slice(0, numSegments)), [segments, numSegments]);

  const updateSegment = (idx: number, field: keyof BattenSegment, val: number) => {
    const next = [...segments];
    next[idx] = { ...next[idx], [field]: val };
    setSegments(next);
  };

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
        <div className="bg-slate-800 p-6 text-white flex items-center justify-between">
          <div>
            <h2 className="text-xl font-black uppercase tracking-tighter">Calcolo Stecca Composta (Forward)</h2>
            <p className="text-xs text-slate-400 font-bold tracking-widest">EI Totale dai Singoli Pezzi</p>
          </div>
          <div className="flex items-center gap-4">
             <button
               onClick={() => setShowFormulaInfo(!showFormulaInfo)}
               className="text-[10px] font-bold uppercase tracking-widest text-slate-300 hover:text-white flex items-center gap-1 transition-colors"
             >
               <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor" className="w-4 h-4">
                 <path strokeLinecap="round" strokeLinejoin="round" d="M11.25 11.25l.041-.02a.75.75 0 011.063.852l-.708 2.836a.75.75 0 001.063.853l.041-.021M21 12a9 9 0 11-18 0 9 9 0 0118 0zm-9-3.75h.008v.008H12V8.25z" />
               </svg>
               Info Metodo
             </button>
             <div className="flex bg-white/10 p-1 rounded-lg">
              {[2, 3, 4].map(n => (
                <button key={n} onClick={() => setNumSegments(n)} className={`px-3 py-1 text-[10px] font-black uppercase rounded-md transition-all ${numSegments === n ? 'bg-white text-slate-900' : 'text-white/50 hover:text-white'}`}>{n} Pezzi</button>
              ))}
            </div>
          </div>
        </div>
        
        {showFormulaInfo && (
          <div className="bg-slate-50 border-b border-slate-200 p-6 animate-in fade-in slide-in-from-top-2">
            <div className="max-w-3xl mx-auto space-y-4">
               <div className="flex justify-between items-start">
                  <h3 className="text-sm font-black text-slate-800 uppercase tracking-wide">Metodologia di Calcolo</h3>
                  <button onClick={() => setShowFormulaInfo(false)} className="text-slate-400 hover:text-slate-600">
                    <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" /></svg>
                  </button>
               </div>
               <p className="text-xs text-slate-600 leading-relaxed">
                 Il calcolatore determina la <strong>rigidezza flessionale equivalente (EI<sub>eq</sub>)</strong> di una trave composta da segmenti discreti con EI variabile.
                 Il calcolo si basa sull'equazione della linea elastica, integrando la curvatura generata da un carico concentrato in mezzeria per ottenere la deflessione teorica (δ).
               </p>
               <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-2">
                  <div className="bg-white p-4 rounded-lg border border-slate-200 shadow-sm">
                    <p className="text-[10px] font-bold text-slate-400 uppercase mb-2">1. Integrazione Curvatura</p>
                    <p className="text-xs text-slate-600 font-mono bg-slate-50 p-2 rounded border border-slate-100 mb-2">
                      Integral = ∫ (x² / EI(x)) dx
                    </p>
                    <p className="text-[10px] text-slate-500">
                      Si calcola l'integrale del momento flettente normalizzato diviso per la rigidezza locale di ogni segmento lungo la semiluce (0 → L/2).
                    </p>
                  </div>
                  <div className="bg-white p-4 rounded-lg border border-slate-200 shadow-sm">
                    <p className="text-[10px] font-bold text-slate-400 uppercase mb-2">2. EI Equivalente</p>
                    <p className="text-xs text-slate-600 font-mono bg-slate-50 p-2 rounded border border-slate-100 mb-2">
                      EI<sub>eq</sub> = L<sup>3</sup> / (24 · Integral)
                    </p>
                    <p className="text-[10px] text-slate-500">
                      Si ricava l'EI costante che produrrebbe la stessa identica deflessione massima (δ) della trave segmentata sotto lo stesso carico.
                    </p>
                  </div>
               </div>
            </div>
          </div>
        )}
        
        <div className="p-8 grid grid-cols-1 lg:grid-cols-2 gap-12">
          <div className="space-y-4">
            {segments.slice(0, numSegments).map((s, idx) => (
              <div key={idx} className="p-4 bg-slate-50 border border-slate-200 rounded-xl grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-[9px] font-black text-slate-400 uppercase">Lunghezza Pezzo {idx+1} (mm)</label>
                  <input type="number" value={s.length} onChange={e => updateSegment(idx, 'length', parseFloat(e.target.value) || 0)} className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg font-bold" />
                </div>
                <div className="space-y-1">
                  <label className="text-[9px] font-black text-slate-400 uppercase">EI Pezzo {idx+1} (Nm²)</label>
                  <input type="number" value={s.ei} onChange={e => updateSegment(idx, 'ei', parseFloat(e.target.value) || 0)} className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg font-bold" />
                </div>
              </div>
            ))}
          </div>

          <div className="flex flex-col justify-center">
            <div className="bg-[#A12B2B] p-10 rounded-3xl text-white shadow-2xl text-center space-y-4">
               <div>
                  <p className="text-[10px] font-bold uppercase opacity-60 tracking-[0.3em]">Rigidezza Equivalente Totale</p>
                  <div className="flex items-baseline justify-center gap-2">
                    <span className="text-7xl font-black">{result.eqEi}</span>
                    <span className="text-sm font-bold opacity-60">N·m²</span>
                  </div>
               </div>
               <div className="pt-4 border-t border-white/10">
                  <p className="text-[10px] font-bold uppercase opacity-40">Lunghezza Totale Assemblata</p>
                  <p className="text-xl font-black">{result.totalLength} <span className="text-[10px] opacity-40">mm</span></p>
               </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default CompositeCalculator;