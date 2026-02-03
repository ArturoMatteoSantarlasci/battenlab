import React from 'react';
import { BattenInputs } from '../types';

interface BattenFormProps {
    inputs: BattenInputs;
    onChange: (updates: Partial<BattenInputs>) => void;
}

const InputGroup = ({
                        title,
                        fields,
                        inputs,
                        handleChange,
                        handleBlur,
                        getValue
                    }: {
    title: string,
    fields: any[],
    inputs: BattenInputs,
    handleChange: (e: React.ChangeEvent<HTMLInputElement>) => void,
    handleBlur: (e: React.FocusEvent<HTMLInputElement>) => void,
    getValue: (val: number) => string | number
}) => (
    <div className="space-y-3">
        <h3 className="text-xs font-bold text-slate-400 uppercase tracking-widest border-b border-slate-100 pb-1">
            {title}
        </h3>
        <div className="grid grid-cols-3 gap-2">
            {fields.map((field) => (
                <div key={field.name}>
                    <label className="block text-[10px] font-medium text-slate-500 mb-0.5 text-center">
                        {field.label}
                    </label>
                    <input
                        type="number"
                        name={field.name}
                        value={getValue(inputs[field.name as keyof BattenInputs])}
                        onChange={handleChange}
                        onBlur={handleBlur}
                        step="0.1"
                        className="w-full px-2 py-1.5 bg-slate-50 border border-slate-200 rounded text-sm text-center focus:ring-1 focus:ring-[#A12B2B] focus:border-[#A12B2B] transition-all outline-none"
                    />
                </div>
            ))}
        </div>
    </div>
);

const BattenForm: React.FC<BattenFormProps> = ({ inputs, onChange }) => {

    // Gestione mentre si scrive
    const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const { name, value } = e.target;

        if (value === '') {
            onChange({ [name]: NaN });
            return;
        }

        const numericValue = parseFloat(value);
        if (isNaN(numericValue)) return;

        if (numericValue !== inputs[name as keyof BattenInputs]) {
            onChange({ [name]: numericValue });
        }
    };

    const handleBlur = (e: React.FocusEvent<HTMLInputElement>) => {
        const { name } = e.target;
        const currentValue = inputs[name as keyof BattenInputs];

        if (isNaN(currentValue)) {
            onChange({ [name]: 0 });
        }
    };

    const getValue = (val: number) => isNaN(val) ? '' : val;

    return (
        <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 space-y-6">
            <h2 className="text-lg font-semibold text-slate-800 flex items-center">
                <svg className="w-5 h-5 mr-2 text-[#A12B2B]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
                </svg>
                Misurazioni
            </h2>

            <div className="grid grid-cols-2 gap-4">
                <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">Test Weight (kg)</label>
                    <input
                        type="number"
                        name="testWeight"
                        value={getValue(inputs.testWeight)}
                        onChange={handleChange}
                        onBlur={handleBlur}
                        step="0.1"
                        className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-[#A12B2B] outline-none"
                    />
                </div>
                <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">Length (mm)</label>
                    <input
                        type="number"
                        name="testLength"
                        value={getValue(inputs.testLength)}
                        onChange={handleChange}
                        onBlur={handleBlur}
                        step="1"
                        className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-[#A12B2B] outline-none"
                    />
                </div>
            </div>

            <InputGroup
                title="Self Weighted (mm)"
                fields={[
                    { name: 'self14', label: '1/4' },
                    { name: 'self12', label: '1/2' },
                    { name: 'self34', label: '3/4' },
                ]}
                inputs={inputs}
                handleChange={handleChange}
                handleBlur={handleBlur}
                getValue={getValue}
            />

            <InputGroup
                title="Weighted (mm)"
                fields={[
                    { name: 'weighted14', label: '1/4' },
                    { name: 'weighted12', label: '1/2' },
                    { name: 'weighted34', label: '3/4' },
                ]}
                inputs={inputs}
                handleChange={handleChange}
                handleBlur={handleBlur}
                getValue={getValue}
            />

            <div className="pt-2">
                <div className="bg-slate-50 p-3 rounded-lg border border-dashed border-slate-300">
                    <p className="text-[10px] text-slate-400 uppercase font-bold mb-2">Net Deflection (Δ):</p>
                    <div className="grid grid-cols-3 gap-2 text-center">
                        {/* Usiamo || 0 per i calcoli visivi in tempo reale */}
                        <div className="text-sm font-bold text-[#A12B2B]">{((inputs.weighted14 || 0) - (inputs.self14 || 0)).toFixed(1)}mm</div>
                        <div className="text-sm font-bold text-[#A12B2B]">{((inputs.weighted12 || 0) - (inputs.self12 || 0)).toFixed(1)}mm</div>
                        <div className="text-sm font-bold text-[#A12B2B]">{((inputs.weighted34 || 0) - (inputs.self34 || 0)).toFixed(1)}mm</div>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default BattenForm;