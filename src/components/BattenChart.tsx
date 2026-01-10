import React from 'react';
import { AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, ReferenceLine, Label, ReferenceDot } from 'recharts';
import { BattenInputs } from '../types';
import { generateCurvePoints } from '../services/calculator';

interface BattenChartProps {
  inputs: BattenInputs;
  draftPos: number;
}

const BattenChart: React.FC<BattenChartProps> = ({ inputs, draftPos }) => {
  const data = generateCurvePoints(inputs);
  // Calcolo della posizione X assoluta in mm basata sulla percentuale del draft
  // Fallback se draftPos non è valido
  const validDraftPos = isNaN(draftPos) ? 50 : draftPos;
  const draftX = (validDraftPos / 100) * inputs.testLength;

  return (
    <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
      <h2 className="text-lg font-semibold mb-4 text-slate-800 flex items-center">
        <svg className="w-5 h-5 mr-2 text-[#A12B2B]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 12l3-3 3 3 4-4M8 21l4-4 4 4M3 4h18M4 4h16v12a1 1 0 01-1 1H5a1 1 0 01-1-1V4z" />
        </svg>
        Visualizzazione Profilo Stecca
      </h2>
      <div className="h-[220px] w-full">
        <ResponsiveContainer width="100%" height="100%">
          <AreaChart
            data={data}
            margin={{ top: 20, right: 30, left: 0, bottom: 0 }}
          >
            <defs>
              <linearGradient id="colorY" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor="#A12B2B" stopOpacity={0.2}/>
                <stop offset="95%" stopColor="#A12B2B" stopOpacity={0}/>
              </linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
            <XAxis 
              dataKey="x" 
              name="Position" 
              unit="mm" 
              stroke="#64748b"
              fontSize={10}
              tickLine={false}
              axisLine={false}
              interval="preserveStartEnd"
              type="number"
              domain={[0, 'dataMax']}
            />
            <YAxis 
              reversed 
              stroke="#64748b" 
              fontSize={10} 
              tickLine={false}
              axisLine={false}
              unit="mm"
              width={35}
            />
            <Tooltip 
              contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '11px' }}
              labelFormatter={(val) => `Position: ${val}mm`}
            />
            <Area 
              type="monotone" 
              dataKey="y" 
              stroke="#A12B2B" 
              strokeWidth={2.5}
              fillOpacity={1} 
              fill="url(#colorY)" 
              isAnimationActive={true}
              animationDuration={600}
            />
            
            {/* Linea Verticale che scende dal marker */}
            <ReferenceLine x={draftX} stroke="#dc2626" strokeDasharray="3 3" strokeOpacity={0.5} />
            
            {/* Pallino Rosso ben visibile all'origine (y=0) */}
            <ReferenceDot 
              x={draftX} 
              y={0} 
              r={6} 
              fill="#dc2626" 
              stroke="#fff" 
              strokeWidth={2} 
              isFront={true}
            />

            {/* Etichetta Testuale sotto il pallino */}
            <ReferenceDot x={draftX} y={0} r={0} isFront={true}>
              <Label 
                value={`DRAFT ${validDraftPos}%`} 
                position="bottom" 
                fill="#dc2626" 
                fontSize={11} 
                fontWeight="bold"
                offset={8}
              />
            </ReferenceDot>

          </AreaChart>
        </ResponsiveContainer>
      </div>
      <div className="mt-4 flex justify-between text-[10px] text-slate-400 font-bold uppercase tracking-widest px-2">
        <span>LUFF (0%)</span>
        <div className="text-[#A12B2B] flex items-center gap-1 opacity-80">
          <span className="w-2 h-2 rounded-full bg-[#dc2626] border border-white shadow-sm inline-block mr-1"></span>
          POSIZIONE DRAFT
        </div>
        <span>LEECH (100%)</span>
      </div>
    </div>
  );
};

export default BattenChart;