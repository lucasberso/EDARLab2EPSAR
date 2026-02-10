(function() {
    "use strict";

    // 1. CONFIGURACIÓN Y COLORES
    const COLOR_OSCURO = "#004381";
    const COLOR_CLARO = "#0097D7";
    const COLOR_FONDO = "#F0F7FD";

    // 2. LIMPIEZA DE INTERFAZ PREVIA
    const existingUi = document.getElementById("epsar-asistente-ui");
    if (existingUi) {
        document.onpaste = null;
        existingUi.remove();
    }

    // 3. DETECCIÓN DE LA EDAR SELECCIONADA EN LA WEB
    const edarSelector = document.getElementById("ctl00_ctl00_ContentPlaceHolder1_DropDownFiltroUnidadCoste");
    const edarWebOrig = edarSelector ? edarSelector.options[edarSelector.selectedIndex].text.trim() : "NO DETECTADA";

    // Función de limpieza de strings (normalización)
    const cl = e => e ? e.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Z0-9]/g, "").trim() : "";
    const cWeb = cl(edarWebOrig);

    // 4. CREACIÓN DE LA INTERFAZ GRÁFICA (UI)
    const ui = document.createElement("div");
    ui.id = "epsar-asistente-ui";
    Object.assign(ui.style, {
        position: "fixed", top: "25px", right: "25px", width: "340px",
        backgroundColor: "#ffffff", color: "#333", borderRadius: "12px",
        zIndex: "10000", fontFamily: "'Segoe UI', sans-serif",
        boxShadow: "0 15px 50px rgba(0,0,0,0.2)", border: `1px solid ${COLOR_CLARO}`,
        overflow: "hidden", animation: "fadeInUI 0.3s ease-out"
    });

    const styleSheet = document.createElement("style");
    styleSheet.innerText = "@keyframes fadeInUI { from { opacity: 0; transform: translateY(-10px); } to { opacity: 1; transform: translateY(0); } }";
    document.head.appendChild(styleSheet);

    ui.innerHTML = `
        <div style="background:linear-gradient(90deg,${COLOR_OSCURO} 0%,${COLOR_CLARO} 100%);color:white;padding:15px 20px;font-weight:bold;font-size:16px;display:flex;justify-content:space-between;align-items:center">
            <span>EDARLab ➔ EPSAR</span>
            <span id="close-ui" style="cursor:pointer;font-size:22px;line-height:1;opacity:0.8;transition:0.2s" onmouseover="this.style.opacity=1" onmouseout="this.style.opacity=0.8">×</span>
        </div>
        <div style="padding:20px">
            <div style="margin-bottom:15px;padding:12px;background-color:${COLOR_FONDO};border-left:5px solid ${COLOR_CLARO};border-radius:4px">
                <span style="display:block;font-size:10px;color:${COLOR_OSCURO};text-transform:uppercase;font-weight:bold">EDAR</span>
                <span style="font-size:15px;font-weight:700;color:${COLOR_OSCURO}">${edarWebOrig}</span>
            </div>
            <div style="font-size:13px;color:#444;line-height:1.6;margin-bottom:15px">
                <b style="color:${COLOR_OSCURO};display:block;margin-bottom:8px;border-bottom:2px solid ${COLOR_FONDO}">Instrucciones:</b>
                <div style="display:flex;margin-bottom:8px"><b>1.&nbsp;</b><span>Copiar todo el archivo Excel <b>Ctrl+C</b>.</span></div>
                <div style="display:flex;margin-bottom:8px"><b>2.&nbsp;</b><span>Pegar <b>Ctrl+V</b> en cualquier lugar.</span></div>
                <div style="display:flex"><b>3.&nbsp;</b><span><b>Revisar</b> que los datos introducidos son correctos.</span></div>
            </div>
            <button id="btn-clear" style="width:100%;padding:6px;background-color:transparent;border:1px solid #ccc;color:#888;border-radius:4px;cursor:pointer;font-size:11px;transition:0.2s">Borrar columnas permitidas</button>
        </div>
        <div id="f-st" style="background:#f9f9fb;padding:10px 20px;font-size:11px;color:#607d8b;text-align:center;border-top:1px solid #eee;display:flex;justify-content:space-between">
            <span id="msg-c">Listo</span>
            <span style="font-size:10px;color:#607d8b;font-weight:700">© Lucas B.</span>
        </div>`;

    document.body.appendChild(ui);

    // 5. MAPEO DE COLUMNAS WEB
    const params = {
        E: { PHEU:"ph", TURBIDEZEU:"tur", V60EU:"v60", SSEU:"ss", DBOEU:"dbo", DQOEU:"dqo", NTEU:"nt", PTEU:"pt" },
        S: { PHS:"ph", CONDUCTIVIDAD:"con", TURBIDEZS:"tur", SSS:"ss", DBOS:"dbo", DQOS:"dqo", NTS:"nt", PTS:"pt" }
    };

    // BOTÓN DE BORRADO SELECTIVO
    const bClear = document.getElementById("btn-clear");
    bClear.onmouseover = () => { bClear.style.borderColor = COLOR_CLARO; bClear.style.color = COLOR_CLARO; bClear.style.backgroundColor = COLOR_FONDO; };
    bClear.onmouseout = () => { bClear.style.borderColor = "#ccc"; bClear.style.color = "#888"; bClear.style.backgroundColor = "transparent"; };
    
    bClear.onclick = function() {
        let count = 0;
        for (let d = 1; d <= 31; d++) {
            ["E", "S"].forEach(pKey => {
                Object.keys(params[pKey]).forEach(col => {
                    const el = document.getElementById(`ctl00_ctl00_ContentPlaceHolder1_Contenido_CELDA_MA_DIA_${d}_COLUMNA_${col}_texto`);
                    if (el) {
                        el.value = "";
                        el.style.backgroundColor = "#fff";
                        el.dispatchEvent(new Event("change", { bubbles: true }));
                        count++;
                    }
                });
            });
        }
        document.getElementById("msg-c").innerHTML = `🗑️ Tabla borrada (${count} celdas)`;
    };

    // CERRAR UI Y DESACTIVAR EVENTO
    document.getElementById("close-ui").onclick = () => {
        document.onpaste = null;
        ui.remove();
    };

    // 6. LÓGICA DE PEGADO (ONPASTE)
    document.onpaste = function(event) {
        event.preventDefault();
        const status = document.getElementById("msg-c");
        const clipboard = (event.clipboardData || window.clipboardData).getData("text");
        const rows = clipboard.split(/\r?\n/).map(row => row.split("\t"));

        if (rows.length < 2) return;

        // Detectar índices de columnas en el Excel
        const headers = rows[0].map(h => h.trim().toLowerCase());
        const getIdx = s => headers.findIndex(h => h.includes(s.toLowerCase()));
        const colMap = {
            ph: getIdx("ph (ud ph)"), pt: getIdx("pt (mg/l)"), nt: getIdx("nt (mg/l)"),
            ss: getIdx("ss (mg/l)"), dqo: getIdx("dqo (mg/l)"), dbo: getIdx("dbo5 (mg/l)"),
            v60: getIdx("v60 (ml/l)"), tur: getIdx("turbidez"), con: getIdx("conductividad")
        };

        let dataStore = {};
        let puntoActual = "";
        let isCorrectEdar = false;
        let foundEdarGlobal = false;

        rows.forEach(row => {
            const lineStr = row.join(" ").toUpperCase();

            // Identificar la EDAR (Igualdad estricta tras limpieza)
            if (lineStr.includes("EDAR:")) {
                const edarIn = cl(row.find(c => c.toUpperCase().includes("EDAR:"))?.split(":")[1]);
                isCorrectEdar = (edarIn === cWeb);
                if (isCorrectEdar) foundEdarGlobal = true;
                return;
            }

            if (isCorrectEdar) {
                // Identificar Punto de Muestreo (IDENTIFICACIÓN ESTRICTA E / S)
                if (lineStr.includes("PUNTO MUESTREO:")) {
                    const pIn = cl(row.find(c => c.toUpperCase().includes("PUNTO MUESTREO:"))?.split(":")[1]);
                    puntoActual = (pIn === "E" || pIn === "S") ? pIn : "";
                    return;
                }

                // Identificar fila de datos por fecha
                const cellFecha = row.find(c => /^\d{2}\/\d{2}\//.test(c.trim()));
                if (cellFecha && puntoActual) {
                    const dia = parseInt(cellFecha.trim().split("/")[0], 10);
                    dataStore[`${puntoActual}_${dia}`] = {
                        ph: row[colMap.ph], pt: row[colMap.pt], nt: row[colMap.nt],
                        ss: row[colMap.ss], dqo: row[colMap.dqo], dbo: row[colMap.dbo],
                        v60: row[colMap.v60], tur: row[colMap.tur], con: row[colMap.con]
                    };
                }
            }
        });

        if (!foundEdarGlobal) {
            status.innerHTML = `<span style="color:#d32f2f;font-weight:bold">❌ EDAR no encontrada en Excel</span>`;
            return;
        }

        // 7. VOLCADO A LA WEB
        let volcados = 0;
        for (let d = 1; d <= 31; d++) {
            ["E", "S"].forEach(p => {
                const diaData = dataStore[`${p}_${d}`];
                if (diaData) {
                    for (const [webCol, dataKey] of Object.entries(params[p])) {
                        const input = document.getElementById(`ctl00_ctl00_ContentPlaceHolder1_Contenido_CELDA_MA_DIA_${d}_COLUMNA_${webCol}_texto`);
                        const val = diaData[dataKey];
                        if (input && val && val.trim() !== "" && val.trim() !== "---") {
                            input.value = val.trim().replace(".", ",");
                            input.style.backgroundColor = COLOR_FONDO;
                            input.dispatchEvent(new Event("change", { bubbles: true }));
                            volcados++;
                        }
                    }
                }
            });
        }
        status.innerHTML = `✅ ${volcados} Celdas procesadas`;
    };

})();
