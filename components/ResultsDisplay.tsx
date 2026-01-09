
import React from 'react';
import { BattenResults } from '../types';

interface ResultsDisplayProps {
  results: BattenResults;
}

const ResultsDisplay: React.FC<ResultsDisplayProps> = ({ results }) => {
  const resultCards = [
    { label: 'Front Bend', value: results.frontPercent.toFixed(1), unit: '%', color: 'text-slate-700' },
    { label: 'Back Bend', value: results.backPercent.toFixed(1), unit: '%', color: 'text-slate-700' },
    { label: 'Camber', value: results.camberPercent.toFixed(2), unit: '%', color: 'text-[#A12B2B]' },
    { label: 'Avg EI', value: results.averageEi.toFixed(3), unit: 'N·m²', color: 'text-[#A12B2B]' },
  ];

  return (
    <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden ring-1 ring-slate-100">
      <div className="grid grid-cols-2 md:grid-cols-4 divide-x divide-slate-100">
        {resultCards.map((card, idx) => (
          <div key={idx} className="p-4 flex flex-col items-center justify-center text-center hover:bg-slate-50 transition-colors">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1 whitespace-nowrap">
              {card.label}
            </span>
            <div className="flex items-baseline gap-1">
              <span className={`text-xl md:text-2xl font-black ${card.color}`}>
                {card.value}
              </span>
              <span className="text-[10px] font-bold text-slate-300">
                {card.unit}
              </span>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

export default ResultsDisplay;
