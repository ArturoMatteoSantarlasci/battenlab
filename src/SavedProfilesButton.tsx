import React, { useEffect, useMemo, useState } from "react";

type SavedProfileMeta = {
  id: string;
  name: string;
  updatedAt: number;
};

type SavedProfile<T> = SavedProfileMeta & {
  data: T;
};

const STORAGE_KEY = "pst_saved_profiles_v1";

function safeParse<T>(raw: string | null): T | null {
  if (!raw) return null;
  try {
    return JSON.parse(raw) as T;
  } catch {
    return null;
  }
}

function readAll<T>(): SavedProfile<T>[] {
  const parsed = safeParse<SavedProfile<T>[]>(localStorage.getItem(STORAGE_KEY));
  return Array.isArray(parsed) ? parsed : [];
}

function writeAll<T>(profiles: SavedProfile<T>[]) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(profiles));
}

function uid() {
  // abbastanza per un localStorage id
  return `${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

export function saveProfile<T>(name: string, data: T): SavedProfile<T> {
  const profiles = readAll<T>();

  // se esiste già un profilo con stesso nome, lo sovrascrivo (scelta comoda)
  const existingIdx = profiles.findIndex(
    p => p.name.trim().toLowerCase() === name.trim().toLowerCase()
  );

  const now = Date.now();
  const newProfile: SavedProfile<T> = {
    id: existingIdx >= 0 ? profiles[existingIdx].id : uid(),
    name: name.trim(),
    updatedAt: now,
    data,
  };

  if (existingIdx >= 0) profiles[existingIdx] = newProfile;
  else profiles.unshift(newProfile);

  writeAll(profiles);
  return newProfile;
}

export function listProfiles<T>(): SavedProfileMeta[] {
  return readAll<T>()
    .map(({ id, name, updatedAt }) => ({ id, name, updatedAt }))
    .sort((a, b) => b.updatedAt - a.updatedAt);
}

export function loadProfile<T>(id: string): SavedProfile<T> | null {
  const profiles = readAll<T>();
  return profiles.find(p => p.id === id) ?? null;
}

export function deleteProfile<T>(id: string) {
  const profiles = readAll<T>().filter(p => p.id !== id);
  writeAll(profiles);
}

export function renameProfile<T>(id: string, newName: string) {
  const profiles = readAll<T>();
  const idx = profiles.findIndex(p => p.id === id);
  if (idx < 0) return;

  profiles[idx] = {
    ...profiles[idx],
    name: newName.trim(),
    updatedAt: Date.now(),
  };
  writeAll(profiles);
}

type Props<T> = {
  // dati correnti da salvare
  currentProfileData: T;

  // callback quando l’utente seleziona un profilo
  onOpenProfile: (profile: SavedProfile<T>) => void;

  // opzionale: se vuoi riusare il tuo input "Nome profilo..."
  getDefaultSaveName?: () => string;

  // opzionale: se vuoi intercettare "Salva" qui, oppure continui a usare il tuo tasto SALVA
  onAfterSave?: (saved: SavedProfile<T>) => void;

  // stile minimal: se hai classi già tue, sostituisci qui
  classNameButton?: string;
};

export default function SavedProfilesButton<T>({
  currentProfileData,
  onOpenProfile,
  getDefaultSaveName,
  onAfterSave,
  classNameButton,
}: Props<T>) {
  const [open, setOpen] = useState(false);
  const [metas, setMetas] = useState<SavedProfileMeta[]>([]);
  const [query, setQuery] = useState("");
  const [renameId, setRenameId] = useState<string | null>(null);
  const [renameValue, setRenameValue] = useState("");

  const refresh = () => setMetas(listProfiles<T>());

  useEffect(() => {
    if (open) refresh();
  }, [open]);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return metas;
    return metas.filter(m => m.name.toLowerCase().includes(q));
  }, [metas, query]);

  const btnStyle: React.CSSProperties = {
    padding: "10px 14px",
    borderRadius: 10,
    border: "1px solid rgba(156, 163, 175, 0.35)",
    background: "#fff",
    fontWeight: 700,
    cursor: "pointer",
  };

  const modalOverlay: React.CSSProperties = {
    position: "fixed",
    inset: 0,
    background: "rgba(0,0,0,0.35)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 9999,
    padding: 16,
  };

  const modalCard: React.CSSProperties = {
    width: "min(820px, 100%)",
    background: "#fff",
    borderRadius: 16,
    boxShadow: "0 20px 50px rgba(0,0,0,0.25)",
    overflow: "hidden",
  };

  const header: React.CSSProperties = {
    padding: 16,
    borderBottom: "1px solid rgba(0,0,0,0.08)",
    display: "flex",
    gap: 12,
    alignItems: "center",
    justifyContent: "space-between",
  };

  const body: React.CSSProperties = {
    padding: 16,
  };

  const row: React.CSSProperties = {
    display: "grid",
    gridTemplateColumns: "1fr auto",
    gap: 12,
    padding: "12px 10px",
    borderRadius: 12,
    border: "1px solid rgba(0,0,0,0.08)",
    alignItems: "center",
    marginBottom: 10,
  };

  const smallBtn: React.CSSProperties = {
    padding: "8px 10px",
    borderRadius: 10,
    border: "1px solid rgba(0,0,0,0.12)",
    background: "#fff",
    cursor: "pointer",
    fontWeight: 700,
  };

  const dangerBtn: React.CSSProperties = {
    ...smallBtn,
    border: "1px solid rgba(220,38,38,0.35)",
  };

  const primaryBtn: React.CSSProperties = {
    ...smallBtn,
    border: "1px solid rgba(153,27,27,0.45)",
  };

  return (
    <>
      <button
        type="button"
        className={classNameButton}
        style={!classNameButton ? btnStyle : undefined}
        onClick={() => setOpen(true)}
        title="Apri la lista dei profili salvati"
      >
        PROFILI
      </button>

      {open && (
        <div style={modalOverlay} onMouseDown={() => setOpen(false)}>
          <div
            style={modalCard}
            onMouseDown={(e) => e.stopPropagation()}
            role="dialog"
            aria-modal="true"
          >
            <div style={header}>
              <div style={{ display: "flex", flexDirection: "column" }}>
                <div style={{ fontSize: 18, fontWeight: 800 }}>
                  Profili salvati
                </div>
                <div style={{ fontSize: 12, opacity: 0.7 }}>
                  Local storage ({metas.length} totali)
                </div>
              </div>

              <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                <input
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                  placeholder="Cerca..."
                  style={{
                    padding: "10px 12px",
                    borderRadius: 10,
                    border: "1px solid rgba(0,0,0,0.12)",
                    width: 220,
                  }}
                />
                <button
                  type="button"
                  style={smallBtn}
                  onClick={() => setOpen(false)}
                >
                  Chiudi
                </button>
              </div>
            </div>

            <div style={body}>
              {filtered.length === 0 ? (
                <div style={{ opacity: 0.75 }}>
                  Nessun profilo trovato.
                </div>
              ) : (
                filtered.map((m) => (
                  <div key={m.id} style={row}>
                    <div>
                      <div style={{ fontWeight: 800 }}>{m.name}</div>
                      <div style={{ fontSize: 12, opacity: 0.65 }}>
                        Aggiornato: {new Date(m.updatedAt).toLocaleString()}
                      </div>

                      {renameId === m.id && (
                        <div style={{ marginTop: 10, display: "flex", gap: 8 }}>
                          <input
                            value={renameValue}
                            onChange={(e) => setRenameValue(e.target.value)}
                            placeholder="Nuovo nome..."
                            style={{
                              padding: "10px 12px",
                              borderRadius: 10,
                              border: "1px solid rgba(0,0,0,0.12)",
                              flex: 1,
                            }}
                          />
                          <button
                            type="button"
                            style={primaryBtn}
                            onClick={() => {
                              const n = renameValue.trim();
                              if (!n) return;
                              renameProfile<T>(m.id, n);
                              setRenameId(null);
                              setRenameValue("");
                              refresh();
                            }}
                          >
                            Salva nome
                          </button>
                          <button
                            type="button"
                            style={smallBtn}
                            onClick={() => {
                              setRenameId(null);
                              setRenameValue("");
                            }}
                          >
                            Annulla
                          </button>
                        </div>
                      )}
                    </div>

                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                      <button
                        type="button"
                        style={primaryBtn}
                        onClick={() => {
                          const p = loadProfile<T>(m.id);
                          if (!p) return;
                          onOpenProfile(p);
                          setOpen(false);
                        }}
                      >
                        Apri
                      </button>

                      <button
                        type="button"
                        style={smallBtn}
                        onClick={() => {
                          setRenameId(m.id);
                          setRenameValue(m.name);
                        }}
                      >
                        Rinomina
                      </button>

                      <button
                        type="button"
                        style={dangerBtn}
                        onClick={() => {
                          deleteProfile<T>(m.id);
                          refresh();
                        }}
                      >
                        Elimina
                      </button>
                    </div>
                  </div>
                ))
              )}

              <div
                style={{
                  marginTop: 14,
                  paddingTop: 14,
                  borderTop: "1px solid rgba(0,0,0,0.08)",
                  display: "flex",
                  justifyContent: "space-between",
                  gap: 10,
                  flexWrap: "wrap",
                }}
              >
                <button
                  type="button"
                  style={primaryBtn}
                  onClick={() => {
                    const defaultName = getDefaultSaveName?.() ?? "";
                    const name = window.prompt("Nome profilo:", defaultName);
                    if (!name?.trim()) return;

                    const saved = saveProfile<T>(name, currentProfileData);
                    onAfterSave?.(saved);
                    refresh();
                  }}
                  title="Salva lo stato corrente come nuovo profilo (o sovrascrive se stesso nome)"
                >
                  + Salva profilo corrente
                </button>

                <button
                  type="button"
                  style={smallBtn}
                  onClick={() => {
                    // export veloce
                    const all = readAll<T>();
                    navigator.clipboard
                      .writeText(JSON.stringify(all, null, 2))
                      .catch(() => {});
                    alert("JSON copiato negli appunti.");
                  }}
                  title="Copia tutti i profili come JSON"
                >
                  Export JSON
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
