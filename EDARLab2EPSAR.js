(function() {
    'use strict';

    // 1. LIMPIEZA DE INTERFAZ PREVIA Y RESET DE EVENTOS
    const existingUi = document.getElementById('epsar-asistente-ui');
    if (existingUi) {
        document.onpaste = null; // Desactiva el pegado si ya existía una instancia
        existingUi.remove();
    }

    // --- COLORES CORPORATIVOS GLOBAL OMNIUM ---
    const GO_AZUL_OSCURO = '#004381';
    const GO_AZUL_CLARO = '#0097D7';
    const GO_FONDO_SUAVE = '#F0F7FD';
    // ------------------------------------------

    // Diccionario de columnas permitidas (utilizado para carga y borrado)
    const qPermitidas = {
        E: { PHEU:'ph', TURBIDEZEU:'tur', V60EU:'v60', SSEU:'ss', DBOEU:'dbo', DQOEU:'dqo', NTEU:'nt', PTEU:'pt' },
        S: { PHS:'ph', CONDUCTIVIDAD:'con', TURBIDEZS:'tur', SSS:'ss', DBOS:'dbo', DQOS:'dqo', NTS:'nt', PTS:'pt' }
    };

    const edarSelector = document.getElementById('ctl00_ctl00_ContentPlaceHolder1_DropDownFiltroUnidadCoste');
    const edarWebOrig = edarSelector ? edarSelector.options[edarSelector.selectedIndex].text.trim() : 'NO DETECTADA';
    
    const cleanStr = (str) => str.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Z0-9]/g, '').trim();
    const edarWebClean = cleanStr(edarWebOrig);

    const ui = document.createElement('div');
    ui.id = 'epsar-asistente-ui';
    Object.assign(ui.style, {
        position: 'fixed', top: '25px', right: '25px', width: '340px',
        backgroundColor: '#ffffff', color: '#333', borderRadius: '12px',
        zIndex: '10000', fontFamily: "'Segoe UI', Roboto, sans-serif",
        boxShadow: '0 15px 50px rgba(0,0,0,0.3)', border: `1px solid ${GO_AZUL_CLARO}`,
        overflow: 'hidden', animation: 'fadeIn 0.3s ease-out',
        transition: 'opacity 0.3s ease'
    });

    const styleSheet = document.createElement("style");
    styleSheet.innerText = `
        @keyframes fadeIn { from { opacity: 0; transform: translateY(-10px); } to { opacity: 1; transform: translateY(0); } }
        .close-btn-hover:hover { background-color: rgba(255,255,255,0.2); }
        .btn-borrar:hover { background-color: #ffebee !important; color: #c62828 !important; border-color: #c62828 !important; }
    `;
    document.head.appendChild(styleSheet);

    ui.innerHTML = `
        <div style="background: linear-gradient(90deg, ${GO_AZUL_OSCURO} 0%, ${GO_AZUL_CLARO} 100%); color: white; padding: 15px 20px; font-weight: bold; font-size: 16px; display: flex; justify-content: space-between; align-items: center;">
            <span>EDARLab ➔ EPSAR</span>
            <span id="close-ui" class="close-btn-hover" style="cursor: pointer; padding: 0px 8px; border-radius: 4px; font-size: 22px; line-height: 1; transition: 0.2s;">&times;</span>
        </div>
        <div style="padding: 20px;">
            <div style="margin-bottom: 15px; padding: 12px; background-color: ${GO_FONDO_SUAVE}; border-left: 5px solid ${GO_AZUL_CLARO}; border-radius: 4px;">
                <span style="display: block; font-size: 10px; color: ${GO_AZUL_OSCURO}; text-transform: uppercase; font-weight: bold; letter-spacing: 0.5px;">Unidad de Coste</span>
                <span style="font-size: 15px; font-weight: 700; color: ${GO_AZUL_OSCURO};">${edarWebOrig}</span>
            </div>
            
            <div style="font-size: 13px; color: #444; line-height: 1.6; margin-bottom: 20px;">
                <b style="color: ${GO_AZUL_OSCURO}; display: block; margin-bottom: 8px; border-bottom: 2px solid ${GO_FONDO_SUAVE}; padding-bottom: 5px;">Instrucciones:</b>
                <div style="margin-bottom: 8px; display: flex; align-items: flex-start;">
                    <span style="font-weight: bold; color: ${GO_AZUL_OSCURO}; margin-right: 10px;">1.</span>
                    <span>Copiar datos del Excel <b>(Ctrl+C)</b>.</span>
                </div>
                <div style="margin-bottom: 8px; display: flex; align-items: flex-start;">
                    <span style="font-weight: bold; color: ${GO_AZUL_OSCURO}; margin-right: 10px;">2.</span>
                    <span>Pegar en la web <b>(Ctrl+V)</b>.</span>
                </div>
            </div>

            <button id="btn-borrar-datos" class="btn-borrar" style="width: 100%; padding: 10px; background-color: transparent; border: 1px solid #ccc; color: #888; border-radius: 8px; cursor: pointer; font-size: 12px; font-weight: bold; transition: 0.2s;">
                🗑️ BORRAR COLUMNAS PERMITIDAS
            </button>
        </div>
        <div id="epsar-footer" style="background: #f5f7f9; padding: 10px 20px; font-size: 11px; color: #607d8b; text-align: center; border-top: 1px solid #cfd8dc; display: flex; justify-content: space-between; align-items: center;">
            <span id="epsar-status">Preparado</span>
            <span style="font-size: 10px; color: #607d8b; font-weight: 700; letter-spacing: 0.3px;">© Lucas B.</span>
        </div>
    `;
    document.body.appendChild(ui);

    // FUNCIÓN PARA CERRAR Y DESACTIVAR PEGADO
    document.getElementById('close-ui').onclick = function() {
        document.onpaste = null; // Detiene la función de pegado global
        ui.style.opacity = '0';
        setTimeout(() => ui.remove(), 300);
    };

    // FUNCIÓN PARA BORRAR CELDAS PERMITIDAS
    document.getElementById('btn-borrar-datos').onclick = function() {
        if (!confirm("¿Deseas borrar todos los datos de las columnas permitidas para este mes?")) return;
        
        let borrados = 0;
        for (let d = 1; d <= 31; d++) {
            ['E', 'S'].forEach(p => {
                for (const webKey of Object.keys(qPermitidas[p])) {
                    const el = document.getElementById(`ctl00_ctl00_ContentPlaceHolder1_Contenido_CELDA_MA_DIA_${d}_COLUMNA_${webKey}_texto`);
                    if (el && el.value !== "") {
                        el.value = "";
                        el.style.backgroundColor = "#ffffff";
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        borrados++;
                    }
                }
            });
        }
        document.getElementById('epsar-status').innerText = `Limpieza completada (${borrados} celdas)`;
    };

    document.onpaste = function(event) {
        event.preventDefault();
        const status = document.getElementById('epsar-status');
        const clipboardData = (event.clipboardData || window.clipboardData).getData('text');
        
        // Ajuste para el problema de los días en blanco: quitamos el filter inicial
        const lines = clipboardData.split(/\r?\n/);
        
        if (lines.length < 2) {
            status.innerHTML = "<span style='color: #d32f2f;'>Error: Portapapeles insuficiente</span>";
            return;
        }

        const headers = lines[0].split('\t').map(h => h.trim().toLowerCase());
        const f = n => headers.indexOf(n.toLowerCase());
        const idx = {
            ph: f('pH (ud pH)'), pt: f('Pt (mg/l)'), nt: f('Nt (mg/l)'), ss: f('SS (mg/l)'), 
            dqo: f('DQO (mg/l)'), dbo: f('DBO5 (mg/l)'), v60: f('V60 (mL/L)'), 
            tur: f('Turbidez (UNT)'), con: f('Conductividad (µS/cm)')
        };

        let dataMap = {};
        let currentPunto = "";
        let isCorrectEdar = false;
        let edarFoundName = "";

        lines.forEach(line => {
            const rawLine = line.trim();
            if (rawLine === "") return;
            const upperLine = rawLine.toUpperCase();
            
            if (upperLine.includes('EDAR:')) {
                const edarPart = rawLine.split(/EDAR:/i)[1].split('\t')[0].trim();
                // Detección flexible corregida
                const cleanIn = cleanStr(edarPart);
                if (cleanIn.includes(edarWebClean) || edarWebClean.includes(cleanIn)) {
                    isCorrectEdar = true;
                    edarFoundName = edarPart;
                } else isCorrectEdar = false;
                return;
            }
            if (!isCorrectEdar) return;
            if (upperLine.includes('PUNTO MUESTREO:')) {
                const puntoPart = cleanStr(rawLine.split(':')[1].split('\t')[0]);
                currentPunto = (puntoPart === 'E' || puntoPart === 'S') ? puntoPart : "";
                return;
            }
            const parts = line.split('\t');
            for (let i = 0; i < parts.length; i++) {
                if (/^\d{2}\/\d{2}\//.test(parts[i].trim())) {
                    const dia = parseInt(parts[i].trim().split('/')[0], 10);
                    if (currentPunto && !isNaN(dia)) {
                        dataMap[`${currentPunto}_${dia}`] = {
                            ph: parts[idx.ph], pt: parts[idx.pt], nt: parts[idx.nt],
                            ss: parts[idx.ss], dqo: parts[idx.dqo], dbo: parts[idx.dbo],
                            v60: parts[idx.v60], tur: parts[idx.tur], con: parts[idx.con]
                        };
                    }
                    break;
                }
            }
        });

        if (!edarFoundName) {
            status.innerHTML = "<span style='color: #d32f2f;'>Error: Planta no coincide</span>";
            alert(`Nombre detectado: ${edarFoundName || 'Ninguno'}\nSe esperaba algo similar a: ${edarWebOrig}`);
            return;
        }

        let count = 0;
        for (let d = 1; d <= 31; d++) {
            ['E', 'S'].forEach(p => {
                for (const [webKey, intKey] of Object.entries(qPermitidas[p])) {
                    const el = document.getElementById(`ctl00_ctl00_ContentPlaceHolder1_Contenido_CELDA_MA_DIA_${d}_COLUMNA_${webKey}_texto`);
                    const val = dataMap[`${p}_${d}`] ? dataMap[`${p}_${d}`][intKey] : null;
                    if (el && val && val.trim() !== "" && val.trim() !== "---") {
                        el.value = val.trim().replace('.', ',');
                        el.style.backgroundColor = GO_FONDO_SUAVE;
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        count++;
                    }
                }
            });
        }

        status.innerHTML = `<span style="color: ${GO_AZUL_OSCURO}; font-weight: bold;">✅ ${count} celdas cargadas</span>`;
    };
})();
