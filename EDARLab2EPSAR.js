(function() {
    'use strict';

    // 1. LIMPIEZA DE INTERFAZ PREVIA Y EVENTOS
    const existingUi = document.getElementById('epsar-asistente-ui');
    if (existingUi) {
        document.onpaste = null;
        existingUi.remove();
    }

    // --- COLORES CORPORATIVOS GLOBAL OMNIUM ---
    const GO_AZUL_OSCURO = '#004381';
    const GO_AZUL_CLARO = '#0097D7';
    const GO_FONDO_SUAVE = '#F0F7FD';

    // Columnas permitidas para escritura y borrado
    const q = {
        E: { PHEU:'ph', TURBIDEZEU:'tur', V60EU:'v60', SSEU:'ss', DBOEU:'dbo', DQOEU:'dqo', NTEU:'nt', PTEU:'pt' },
        S: { PHS:'ph', CONDUCTIVIDAD:'con', TURBIDEZS:'tur', SSS:'ss', DBOS:'dbo', DQOS:'dqo', NTS:'nt', PTS:'pt' }
    };

    const edarSelector = document.getElementById('ctl00_ctl00_ContentPlaceHolder1_DropDownFiltroUnidadCoste');
    const edarWebOrig = edarSelector ? edarSelector.options[edarSelector.selectedIndex].text.trim() : 'NO DETECTADA';
    
    // Función de normalización compacta
    const n = e => e ? e.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^A-Z0-9]/g,"").trim() : "";
    const r = n(edarWebOrig);

    const ui = document.createElement('div');
    ui.id = 'epsar-asistente-ui';
    Object.assign(ui.style, {
        position: 'fixed', top: '25px', right: '25px', width: '340px',
        backgroundColor: '#ffffff', color: '#333', borderRadius: '12px',
        zIndex: '10000', fontFamily: "'Segoe UI', sans-serif",
        boxShadow: '0 15px 50px rgba(0,0,0,0.3)', border: `2px solid ${GO_AZUL_CLARO}`,
        overflow: 'hidden', animation: 'fadeIn 0.3s ease-out'
    });

    const styleSheet = document.createElement("style");
    styleSheet.innerText = `
        @keyframes fadeIn { from { opacity: 0; transform: translateY(-10px); } to { opacity: 1; transform: translateY(0); } }
        .btn-clear:hover { background-color: ${GO_FONDO_SUAVE} !important; color: ${GO_AZUL_OSCURO} !important; border-color: ${GO_AZUL_OSCURO} !important; }
    `;
    document.head.appendChild(styleSheet);

    ui.innerHTML = `
        <div style="background: linear-gradient(90deg, ${GO_AZUL_OSCURO} 0%, ${GO_AZUL_CLARO} 100%); color: white; padding: 15px 20px; font-weight: bold; display: flex; justify-content: space-between; align-items: center;">
            <span>EDARLab ➔ EPSAR</span>
            <span id="close-ui" style="cursor: pointer; font-size: 20px;">&times;</span>
        </div>
        <div style="padding: 20px;">
            <div style="margin-bottom: 15px; padding: 10px; background: ${GO_FONDO_SUAVE}; border-left: 4px solid ${GO_AZUL_CLARO}; border-radius: 4px;">
                <div style="font-size: 10px; font-weight: bold; color: ${GO_AZUL_OSCURO};">UNIDAD DE COSTE</div>
                <div style="font-size: 14px; font-weight: bold;">${edarWebOrig}</div>
            </div>
            <div id="st" style="font-size: 12px; margin-bottom: 15px; color: #555;">Esperando Ctrl+V...</div>
            <button id="btn-borrar" class="btn-clear" style="width: 100%; padding: 8px; background: transparent; border: 1px solid #ccc; color: #888; border-radius: 6px; cursor: pointer; font-size: 11px; font-weight: bold; transition: 0.2s;">
                🗑️ BORRAR COLUMNAS PERMITIDAS
            </button>
        </div>
        <div style="background: #f9f9fb; padding: 10px 20px; font-size: 10px; color: #999; border-top: 1px solid #eee; display: flex; justify-content: space-between;">
            <span>Listo</span><span>© Lucas B.</span>
        </div>
    `;
    document.body.appendChild(ui);

    // Cierre y parada
    document.getElementById('close-ui').onclick = () => { document.onpaste = null; ui.remove(); };

    // Borrado selectivo
    document.getElementById('btn-borrar').onclick = () => {
        if(!confirm("¿Borrar datos del mes en las columnas permitidas?")) return;
        let c = 0;
        for(let v=1; v<=31; v++) {
            ['E','S'].forEach(p => {
                Object.keys(q[p]).forEach(col => {
                    const el = document.getElementById(`ctl00_ctl00_ContentPlaceHolder1_Contenido_CELDA_MA_DIA_${v}_COLUMNA_${col}_texto`);
                    if(el) { el.value = ""; el.style.backgroundColor = "#fff"; el.dispatchEvent(new Event("change",{bubbles:!0})); c++; }
                });
            });
        }
        document.getElementById('st').innerHTML = `🗑️ ${c} celdas limpiadas`;
    };

    // Pegado con lógica de detección compacta (Inclusión)
    document.onpaste = e => {
        e.preventDefault();
        const status = document.getElementById('st');
        const text = (e.clipboardData || window.clipboardData).getData("text");
        const lines = text.split(/\r?\n/).map(l => l.split("\t"));
        if (lines.length < 2) return;

        const h = lines[0].map(l => l.trim().toLowerCase());
        const f = s => h.findIndex(x => x.includes(s.toLowerCase()));
        const s = { ph:f("ph (ud ph)"), pt:f("pt (mg/l)"), nt:f("nt (mg/l)"), ss:f("ss (mg/l)"), dqo:f("dqo (mg/l)"), dbo:f("dbo5 (mg/l)"), v60:f("v60 (ml/l)"), tur:f("turbidez"), con:f("conductividad") };

        let l = {}, d = "", u = !1;
        lines.forEach(row => {
            const rowTxt = row.join(" ").toUpperCase();
            if (rowTxt.includes("EDAR:")) {
                const edarIn = n(row.find(c => c.toUpperCase().includes("EDAR:"))?.split(":")[1]);
                // Lógica compacta: Si el nombre de la web está contenido en el del Excel o viceversa
                u = (edarIn.includes(r) || r.includes(edarIn)) && r !== "";
                return;
            }
            if (u) {
                if (rowTxt.includes("PUNTO MUESTREO:")) {
                    const p = n(row.find(c => c.toUpperCase().includes("PUNTO MUESTREO:"))?.split(":")[1]);
                    d = p.startsWith("E") ? "E" : p.startsWith("S") ? "S" : "";
                    return;
                }
                const dateIdx = row.find(c => /^\d{2}\/\d{2}\//.test(c.trim()));
                if (dateIdx && d) {
                    const day = parseInt(dateIdx.trim().split("/")[0], 10);
                    l[`${d}_${day}`] = { ph:row[s.ph], pt:row[s.pt], nt:row[s.nt], ss:row[s.ss], dqo:row[s.dqo], dbo:row[s.dbo], v60:row[s.v60], tur:row[s.tur], con:row[s.con] };
                }
            }
        });

        let count = 0;
        for (let v=1; v<=31; v++) {
            ['E','S'].forEach(p_idx => {
                if (l[`${p_idx}_${v}`]) {
                    for (const [webCol, intCol] of Object.entries(q[p_idx])) {
                        const el = document.getElementById(`ctl00_ctl00_ContentPlaceHolder1_Contenido_CELDA_MA_DIA_${v}_COLUMNA_${webCol}_texto`);
                        const val = l[`${p_idx}_${v}`][intCol];
                        if (el && val && val.trim() !== "" && val.trim() !== "---") {
                            el.value = val.trim().replace(".", ",");
                            el.style.backgroundColor = GO_FONDO_SUAVE;
                            el.dispatchEvent(new Event("change", { bubbles: true }));
                            count++;
                        }
                    }
                }
            });
        }
        status.innerHTML = count > 0 ? `✅ ${count} volcados con éxito` : "<span style='color:red'>❌ No se encontraron datos para esta EDAR</span>";
    };
})();
