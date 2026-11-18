import React, { useState, useRef, useMemo } from 'react';

const AIRLINES = [
  { id: 'H8', name: 'SKY Airline', base: 'LIM' },
  { id: 'JZ', name: 'JetSMART', base: 'LIM' },
];

const JETSMART_COLS = ['DATE', 'Route', 'FLT', 'REG', 'DEP', 'ARR', 'STD', 'STA', 'PAX'];

function formatJetSmartText(rawText) {
  if (!rawText.trim()) return { formattedText: '', rows: [] };
  const flat = rawText.replace(/\s+/g, ' ').trim().toUpperCase();
  const regex = /(\d{2}[/-]\d{2}[/-]\d{4})\s+(\d+)\s+(\d+)\s+((?:CC|LV)-[A-Z0-9]{3})\s+([A-Z]{3})\s+([A-Z]{3})\s+(\d{1,2}:\d{2})\s+(\d{1,2}:\d{2})(?:\s+(\d+))?/g;
  
  const rows = [];
  const formattedLines = [];
  let match;

  while ((match = regex.exec(flat)) !== null) {
    const row = [
      match[1],
      match[2],
      match[3],
      match[4],
      match[5],
      match[6],
      match[7].padStart(5, '0'),
      match[8].padStart(5, '0'),
      match[9] || '',
    ];
    rows.push(row);
    formattedLines.push(row.join('\t'));
  }

  return {
    formattedText: formattedLines.length > 0 ? formattedLines.join('\n') : rawText,
    rows,
  };
}

export default function App() {
  const [selectedAirline, setSelectedAirline] = useState(AIRLINES[1]);
  const [detectedDate, setDetectedDate] = useState('');
  const [rotaciones, setRotaciones] = useState([]);
  const [loading, setLoading] = useState(false);
  const [lastPayload, setLastPayload] = useState({ type: null, data: null });
  const [emailText, setEmailText] = useState('');
  const [errorMsg, setErrorMsg] = useState('');
  const fileInputRef = useRef(null);

  const { rows: parsedInputRows } = useMemo(() => formatJetSmartText(emailText), [emailText]);

  const resetState = () => {
    setRotaciones([]);
    setDetectedDate('');
    setLastPayload({ type: null, data: null });
    setEmailText('');
    setErrorMsg('');
  };

  const handleFileUpload = async (file) => {
    if (!file) return;
    setLoading(true);
    setErrorMsg('');
    setLastPayload({ type: 'file', data: file });

    const formData = new FormData();
    formData.append('file', file);
    formData.append('airline', selectedAirline.id);
    formData.append('base', selectedAirline.base);

    try {
      const response = await fetch('/api/procesar-preview', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) throw new Error('No se pudo procesar el archivo. Revisa el formato.');

      const data = await response.json();
      setDetectedDate(data.fecha_detectada);
      setRotaciones(data.rotaciones);
    } catch (err) {
      setErrorMsg(err.message);
    } finally {
      setLoading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleTextProcess = async () => {
    if (!emailText.trim()) return;
    setLoading(true);
    setErrorMsg('');
    setLastPayload({ type: 'text', data: emailText });

    try {
      const response = await fetch('/api/procesar-preview', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          texto: emailText,
          airline: selectedAirline.id,
          base: selectedAirline.base,
        }),
      });

      if (!response.ok) throw new Error('No se pudo procesar el texto de vuelos.');

      const data = await response.json();
      setDetectedDate(data.fecha_detectada);
      setRotaciones(data.rotaciones);
    } catch (err) {
      setErrorMsg(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleDownloadExcel = async () => {
    if (!lastPayload.type) return;

    try {
      let response;
      if (lastPayload.type === 'text') {
        response = await fetch('/api/convertir-talma', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            texto: lastPayload.data,
            airline: selectedAirline.id,
            base: selectedAirline.base,
          }),
        });
      } else {
        const formData = new FormData();
        formData.append('file', lastPayload.data);
        formData.append('airline', selectedAirline.id);
        formData.append('base', selectedAirline.base);
        response = await fetch('/api/convertir-talma', {
          method: 'POST',
          body: formData,
        });
      }

      if (!response.ok) throw new Error('Error al descargar el archivo.');

      const blob = await response.blob();
      const disposition = response.headers.get('Content-Disposition');
      let filename = `TALMA_${selectedAirline.id}_OPERACIONES.xlsx`;
      if (disposition && disposition.includes('filename=')) {
        filename = disposition.split('filename=')[1].replace(/["']/g, '');
      }

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      alert(err.message);
    }
  };

  return (
    <div className="min-h-screen bg-[#f3f7fa] flex flex-col items-center py-10 px-4 text-slate-800 font-sans">
      <div className="w-full max-w-4xl">
        {/* Selector de pestañas */}
        <div className="flex justify-center mb-6">
          <div className="inline-flex bg-white p-1 rounded-xl shadow-sm border border-slate-200">
            {AIRLINES.map((al) => (
              <button
                key={al.id}
                onClick={() => {
                  setSelectedAirline(al);
                  resetState();
                }}
                className={`px-6 py-2 text-sm font-semibold rounded-lg transition-all ${
                  selectedAirline.id === al.id
                    ? 'bg-[#1a365d] text-white shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                {al.name}
              </button>
            ))}
          </div>
        </div>

        {/* Título */}
        <div className="text-center mb-6">
          <h1 className="text-3xl md:text-4xl font-extrabold text-[#0f284e] tracking-tight">
            Operaciones {selectedAirline.name} — {selectedAirline.base}
          </h1>
          <p className="mt-1 text-sm text-slate-500 font-normal">
            Procesador automatizado de itinerarios a Excel
          </p>
        </div>

        {/* Barra superior de Fecha detectada */}
        <div className="bg-white rounded-xl px-6 py-4 shadow-sm border border-slate-100 flex items-center justify-between mb-5">
          <span className="text-sm font-medium text-slate-600">Fecha detectada:</span>
          <div className="w-48 text-right">
            <span className="text-base font-bold text-[#0f284e] border-b-2 border-slate-300 pb-0.5 px-2">
              {detectedDate || '--/--/----'}
            </span>
          </div>
        </div>

        {/* JETSMART: Guía de 9 Columnas + Pegado Directo + Tabla de Validación */}
        {selectedAirline.id === 'JZ' ? (
          <div className="bg-white rounded-2xl p-6 shadow-sm border border-slate-200 mb-6">
            <div className="flex items-center justify-between mb-3">
              <label className="text-sm font-bold text-[#1e3a63]">
                Pega aquí la tabla del correo de JetSMART:
              </label>
              <button
                type="button"
                onClick={() => {
                  setEmailText('');
                  setRotaciones([]);
                  setDetectedDate('');
                }}
                className="text-xs text-slate-400 hover:text-rose-500 cursor-pointer"
              >
                Limpiar todo
              </button>
            </div>

            {/* Encabezados de referencia de las 9 columnas del correo */}
            <div className="grid grid-cols-9 bg-[#1a365d] text-white text-[11px] font-bold tracking-wider uppercase rounded-t-xl py-2 px-3 text-center">
              {JETSMART_COLS.map((col) => (
                <div key={col}>{col}</div>
              ))}
            </div>

            <textarea
              rows={6}
              value={emailText}
              onChange={(e) => {
                const val = e.target.value;
                const { formattedText } = formatJetSmartText(val);
                setEmailText(formattedText);
              }}
              placeholder="Selecciona la tabla en tu correo, copia (Ctrl+C) y pega aquí (Ctrl+V)..."
              className="w-full text-xs font-mono p-3.5 border-x border-b border-slate-200 rounded-b-xl focus:outline-none focus:border-[#1a365d] bg-slate-50 leading-relaxed"
            />

            {/* Vista previa en columnas ordenadas del texto pegado */}
            {parsedInputRows.length > 0 && (
              <div className="mt-4 border border-slate-200 rounded-xl overflow-hidden">
                <div className="bg-slate-100 px-3 py-1.5 text-[11px] font-semibold text-slate-600 flex justify-between">
                  <span>Columnas detectadas automáticamente ({parsedInputRows.length} filas leídas)</span>
                  <span className="text-emerald-700">✓ Alineación verificada</span>
                </div>
                <div className="max-h-44 overflow-y-auto">
                  <table className="w-full text-xs text-center font-mono">
                    <thead className="bg-slate-50 text-slate-500 border-b border-slate-200 text-[10px]">
                      <tr>
                        {JETSMART_COLS.map((c) => (
                          <th key={c} className="py-1.5 px-2">{c}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {parsedInputRows.map((row, rIdx) => (
                        <tr key={rIdx} className={row[4] === 'LIM' || row[5] === 'LIM' ? 'bg-white' : 'bg-amber-50/50 text-slate-400'}>
                          {JETSMART_COLS.map((_, cIdx) => (
                            <td key={cIdx} className="py-1 px-2">{row[cIdx] || '-'}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            <div className="mt-4 flex justify-end">
              <button
                onClick={handleTextProcess}
                disabled={loading || !emailText.trim()}
                className="bg-[#1a365d] hover:bg-[#0f284e] disabled:opacity-40 text-white font-bold text-sm px-6 py-2.5 rounded-xl shadow-sm transition-all cursor-pointer"
              >
                {loading ? 'Analizando vuelos...' : 'Ver Rotaciones y Tiempos'}
              </button>
            </div>
          </div>
        ) : (
          /* SKY: Arrastrar Archivo */
          <div
            onDrop={(e) => {
              e.preventDefault();
              if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                handleFileUpload(e.dataTransfer.files[0]);
              }
            }}
            onDragOver={(e) => e.preventDefault()}
            onClick={() => fileInputRef.current && fileInputRef.current.click()}
            className="bg-white border-2 border-dashed border-[#b8cde3] rounded-2xl p-10 flex flex-col items-center justify-center cursor-pointer transition-all hover:border-[#1a365d] hover:bg-slate-50/50 mb-6 group"
          >
            <div className="w-14 h-14 rounded-xl flex items-center justify-center text-[#64748b] group-hover:text-[#1a365d] transition-colors mb-3">
              <svg className="w-12 h-12 stroke-[1.6]" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                <polyline points="14 2 14 8 20 8" />
                <line x1="12" y1="18" x2="12" y2="12" />
                <polyline points="9 15 12 12 15 15" />
              </svg>
            </div>

            <h2 className="text-base md:text-lg font-bold text-[#1e3a63]">
              {loading ? 'Analizando vuelos en memoria...' : 'Arrastra o selecciona el PDF o Excel de vuelos'}
            </h2>

            <p className="mt-1 text-xs text-slate-400 font-normal">
              Formato oficial de Sky Airline (.xlsx o .pdf)
            </p>

            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.pdf"
              className="hidden"
              onChange={(e) => handleFileUpload(e.target.files[0])}
            />
          </div>
        )}

        {errorMsg && (
          <div className="mb-6 p-4 rounded-xl bg-rose-50 border border-rose-200 text-rose-700 text-sm font-medium">
            {errorMsg}
          </div>
        )}

        {/* Botón Verde Descargar Excel Oficial */}
        {rotaciones.length > 0 && (
          <div className="flex justify-end mb-4">
            <button
              onClick={handleDownloadExcel}
              className="flex items-center gap-2 bg-[#028457] hover:bg-[#026d48] text-white font-bold text-sm px-6 py-2.5 rounded-xl shadow-sm transition-all cursor-pointer"
            >
              <svg className="w-4 h-4 fill-current" viewBox="0 0 24 24">
                <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-7 14H7v-2h5v2zm5-4H7v-2h10v2zm0-4H7V7h10v2z"/>
              </svg>
              Descargar Excel Oficial
            </button>
          </div>
        )}

        {/* Tarjeta de Rotaciones y Tiempos */}
        {rotaciones.length > 0 && (
          <div className="bg-white rounded-2xl p-6 shadow-sm border border-slate-100">
            <h3 className="text-base font-bold text-[#1e3a63] mb-4">
              Vuelos Reconocidos y Emparejados ({rotaciones.length} rotaciones)
            </h3>

            <div className="max-h-[480px] overflow-y-auto space-y-3 pr-2 divide-y divide-slate-100">
              {rotaciones.map((rot, idx) => {
                const salidaVuelo = rot.VUELO_S ? `Vuelo ${rot.VUELO_S}` : '---';
                const salidaDest = rot.CIUDAD_S || 'Sin Salida';
                const llegadaVuelo = rot.VUELO_L ? `Vuelo ${rot.VUELO_L}` : '---';
                const llegadaOrig = rot.CIUDAD_L || 'Sin Llegada';
                const mat = rot.MATRICULA_S || rot.MATRICULA_L || '---';
                const acft = rot.AERONAVE_S || rot.AERONAVE_L || '---';

                return (
                  <div key={idx} className="pt-3 first:pt-0 flex items-center justify-between text-xs md:text-sm">
                    <div className="flex items-center gap-3">
                      <div className="w-7 h-7 rounded-lg bg-sky-50 flex items-center justify-center text-sky-600 font-bold text-sm">
                        ↗
                      </div>
                      <div>
                        <div className="font-bold text-slate-800">
                          {rot.VUELO_S ? `${salidaVuelo} → ${salidaDest}` : '--- Sin Salida'}
                        </div>
                        <div className="text-[11px] text-slate-400 mt-0.5">
                          {mat} ({acft}) {rot.ETD && `| ETD: ${rot.ETD}`}
                        </div>
                      </div>
                    </div>

                    <div className="flex items-center gap-3 text-right">
                      <div>
                        <div className="font-bold text-slate-800">
                          {rot.VUELO_L ? `${llegadaOrig} → ${llegadaVuelo}` : '--- Sin Llegada'}
                        </div>
                        <div className="text-[11px] text-emerald-600 font-semibold mt-0.5">
                          {rot.ETA ? `ETA: ${rot.ETA}` : ''} {rot.T_VUELO && `| Giro: ${rot.T_VUELO}`}
                        </div>
                      </div>
                      <div className="w-7 h-7 rounded-lg bg-sky-50 flex items-center justify-center text-sky-600 font-bold text-sm">
                        ↙
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}