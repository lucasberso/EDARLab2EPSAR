// ==UserScript==
// @name         Parte Análisis en Influente - EDARLab➔EPSAR - GIT
// @version      6.7
// @description  Herramienta que automatiza la introducción individualizada de partes de analíticas en el portal de la EPSAR con arquitectura modular.
// @author       Lucas B.
// @match        https://aplica.epsar.gva.es/depuradoras/Partes/AnalisisInfluente.aspx*
// @downloadURL  https://github.com/lucasberso/EDARLab2EPSAR/raw/refs/heads/main/ParteInfluenteEPSAR.user.js
// @updateURL    https://github.com/lucasberso/EDARLab2EPSAR/raw/refs/heads/main/ParteInfluenteEPSAR.user.js
// @run-at       document-start
// ==/UserScript==

(function() {
    'use strict';

    // =========================================================================
    // --- 1. CONSTANTES, PAUSAS Y CONFIGURACIÓN ---
    // =========================================================================
    const diaDelMes = new Date().getDate();
    const FACTOR_SATURACION = (diaDelMes >= 1 && diaDelMes <= 4) ? 2.00 : 1.0;

    console.log(`[EDARLab➔EPSAR] Día del mes: ${diaDelMes}. Factor de retraso aplicado: x${FACTOR_SATURACION}`);

    const CONFIG_PAUSAS = {
        POST_CELDA: () => (2000 + Math.random() * 500) * FACTOR_SATURACION,
        ASIMILACION_DESPLEGABLE: (5000 + Math.random() * 3000) * FACTOR_SATURACION,
        RENDERIZADO_TABLA: (10000 + Math.random() * 5000) * FACTOR_SATURACION,
        PRE_RECALCULA: (20000 + Math.random() * 5000) * FACTOR_SATURACION,
        POST_RECALCULA: (15000 + Math.random() * 5000) * FACTOR_SATURACION,
        TRANSICION_PLANTA: (15000 + Math.random() * 5000) * FACTOR_SATURACION
    };

    const STYLE = {
        PRIMARIO: "#004381",
        SECUNDARIO: "#0097D7",
        GRIS: "#64748b",
        FONT: "'Segoe UI', sans-serif"
    };

    // --- MAPEO DE SELECTORES DE LA WEB DE LA EPSAR ---
    const SELECTORS = {
        FILTRO_EDAR: "ctl00_ctl00_ContentPlaceHolder1_DropDownFiltroUnidadCoste",
        FILTRO_MOSTRAR: "ctl00_ctl00_ContentPlaceHolder1_ButtonFiltroMostrar",
        BOTON_NUEVO_PARTE: "ctl00_ctl00_ContentPlaceHolder1_ButtonNuevoParte",
        BOTON_GUARDAR: "ctl00_ctl00_ContentPlaceHolder1_ButtonGuardar",
        FORM_EDAR: "ctl00_ctl00_ContentPlaceHolder1_Contenido_DropDownNuevaUnidadCoste",
        FORM_FECHA: "ctl00_ctl00_ContentPlaceHolder1_Contenido_TBFecha",
        FORM_MUESTREO: "ctl00_ctl00_ContentPlaceHolder1_Contenido_DDLTipMuestreo",
        INPUT_NNH4: "ctl00_ctl00_ContentPlaceHolder1_Contenido_TBNitrogenoAmoniacal",
        INPUT_NNO3: "ctl00_ctl00_ContentPlaceHolder1_Contenido_TBNitratos",
        INPUT_NNO2: "ctl00_ctl00_ContentPlaceHolder1_Contenido_TBNitritos",
        INPUT_NT: "ctl00_ctl00_ContentPlaceHolder1_Contenido_TBNitrogenoTotal",
        INPUT_PT: "ctl00_ctl00_ContentPlaceHolder1_Contenido_TBFosforoTotal"
    };

    const EXCEL_COLUMNS = { NNH4: 25, NNO2: 27, NNO3: 29 , NT: 19, PT: 22};
    const ORDEN_ESCRITURA = ["NNH4", "NNO2", "NNO3", "NT", "PT"];

    let uiContenedor = null;
    let botonLanzador = null;
    let escuchandoHerramienta = false;
    let wakeLock = null;

    window.EDARLab_Buffer = [];

    // =========================================================================
    // --- 2. GESTIÓN DE ALMACENAMIENTO Y ESTADO ---
    // =========================================================================
    const StateManager = {
        getAutomatorState() {
            const raw = sessionStorage.getItem("edar_automator_state");
            return raw ? JSON.parse(raw) : null;
        },
        setAutomatorState(state) {
            sessionStorage.setItem("edar_automator_state", JSON.stringify(state));
        },
        getBuffer() {
            const raw = sessionStorage.getItem("edarlab_global_buffer_storage");
            return raw ? JSON.parse(raw) : window.EDARLab_Buffer;
        },
        setBuffer(buffer) {
            window.EDARLab_Buffer = buffer;
            sessionStorage.setItem("edarlab_global_buffer_storage", JSON.stringify(buffer));
        },
        getLog() {
            return sessionStorage.getItem("edarlab_txt_log_accumulator") || "";
        },
        setLog(txt) {
            sessionStorage.setItem("edarlab_txt_log_accumulator", txt);
        },
        clearAll() {
            sessionStorage.removeItem("edar_automator_state");
            sessionStorage.removeItem("edarlab_global_buffer_storage");
            sessionStorage.removeItem("edarlab_txt_log_accumulator");
            window.EDARLab_Buffer = [];
        }
    };

    // =========================================================================
    // --- 3. UTILIDADES DOM Y EVENTOS ---
    // =========================================================================
    const DOMUtils = {
        trigger(el, eventType, options = {}) {
            if (!el) return;
            // Eliminamos "view: window" para evitar que el Sandbox de Tampermonkey rompa la conversión nativa
            const defaults = { bubbles: true, cancelable: true };
            let event;
            if (eventType.startsWith('key')) {
                event = new KeyboardEvent(eventType, { ...defaults, ...options });
            } else if (eventType.startsWith('mouse') || eventType === 'click') {
                event = new MouseEvent(eventType, { ...defaults, ...options });
            } else if (eventType === 'focus' || eventType === 'blur') {
                event = new FocusEvent(eventType, { ...defaults, ...options });
            } else {
                event = new Event(eventType, { bubbles: true });
            }
            el.dispatchEvent(event);
        },
        click(el) {
            if (!el) return;
            this.trigger(el, 'mousedown');
            this.trigger(el, 'mouseup');
            el.focus();
            el.click();
        }
    };

    // =========================================================================
    // --- 4. INTERCEPTOR DE ALERTAS DE DUPLICADO ---
    // =========================================================================
    if (typeof unsafeWindow !== 'undefined') {
        const originalAlert = unsafeWindow.alert;
        unsafeWindow.alert = function(msg) {
            if (msg) {
                const msgClean = msg.toString().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                if (msgClean.includes("YA EXISTE") || msgClean.includes("DUPLICADO") || msgClean.includes("EXISTE UN PARTE")) {
                    console.warn("[EDARLab] Alerta de duplicidad interceptada de forma directa.");
                    registrarParteExistenteYContinuar();
                    return true;
                }
            }
            return originalAlert(msg);
        };
    }

    function registrarParteExistenteYContinuar() {
        const estado = StateManager.getAutomatorState();
        if (!estado || estado.pasoActual !== "ESPERAR_GUARDADO") return;

        const buffer = StateManager.getBuffer();
        const tareaActual = estado.tareas[estado.tareaIndexActual];
        const candActual = buffer[tareaActual.edarIdx];
        const nombreWebActual = candActual ? candActual.nameWebAsociada : "Desconocida";

        let log = StateManager.getLog();
        const horaLog = new Date().toLocaleTimeString();

        log += `[${horaLog}] OMITIDO (El registro ya existía en la web) | EDAR ${candActual ? candActual.nameExcel : 'Desconocida'} (${nombreWebActual}) | Fecha: ${tareaActual.fecha}\n`;
        log += `--------------------------------------------------\n\n`;

        StateManager.setLog(log);

        estado.tareaIndexActual++;
        estado.pasoActual = "CLIC_NUEVO_PARTE";
        StateManager.setAutomatorState(estado);

        // window.location.href = "https://aplica.epsar.gva.es/depuradoras/Partes/AnalisisInfluente.aspx?res=E";
        ejecutarMaquinaEstados(estado)
    }

    // =========================================================================
    // --- 5. LÓGICA DE PROCESAMIENTO Y COMPATIBILIDAD ---
    // =========================================================================
    const cl = e => e ? e.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Z0-9]/g, "").trim() : "";
    const splitWords = s => s.split(/[\s\-_]+/).filter(w => w.length > 2);

    const getSim = (w1, w2) => {
        if (w1 === w2) return 1;
        let m = 0;
        for (let i = 0; i < w1.length; i++) {
            if (w2.includes(w1[i])) m++;
        }
        return m / Math.max(w1.length, w2.length);
    };

    const typeH = async(el, txt) => {
        if (!el) return;
        DOMUtils.trigger(el, 'mousedown');
        DOMUtils.trigger(el, 'mouseup');
        DOMUtils.trigger(el, 'focus');
        el.focus();
        DOMUtils.trigger(el, 'click');

        const primeraTecla = txt[0] || '';
        DOMUtils.trigger(el, 'keydown', { key: primeraTecla });

        el.value = txt;
        DOMUtils.trigger(el, 'input');
        DOMUtils.trigger(el, 'keyup', { key: primeraTecla });

        await new Promise(r => setTimeout(r, 100 + Math.random() * 50));

        DOMUtils.trigger(el, 'change');
        await new Promise(r => setTimeout(r, 50));

        DOMUtils.trigger(el, 'blur');
        el.blur();

        await new Promise(r => setTimeout(r, 100 + Math.random() * 50));
        await new Promise(r => setTimeout(r, CONFIG_PAUSAS.POST_CELDA()));
    };

    const delay = ms => new Promise(res => setTimeout(res, ms));

    function buscarMejorCoincidenciaWeb(nombreExcel) {
        const r = document.getElementById(SELECTORS.FILTRO_EDAR);
        if (!r) return { nombreWeb: "DESPLEGABLE NO ENCONTRADO", score: 0 };

        const cExcel = cl(nombreExcel);
        const wEx = splitWords(nombreExcel.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));

        let mejorCandidato = "NO ASOCIADA";
        let maxScore = 0;

        for (let i = 0; i < r.options.length; i++) {
            const nombreOpcionWeb = r.options[i].text.trim();
            const cWeb = cl(nombreOpcionWeb);
            const wWeb = splitWords(nombreOpcionWeb.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));

            let score = 0;

            if (cExcel.replace(/\s/g, "") === cWeb.replace(/\s/g, "")) {
                score = 1e7;
            } else {
                wEx.forEach(wx => {
                    let bWS = 0;
                    wWeb.forEach(ww => {
                        let s = getSim(wx, ww);
                        if (s > bWS) bWS = s;
                    });
                    if (bWS > 0.6) score += bWS;
                });
            }

            if (score > maxScore && score > 0) {
                maxScore = score;
                mejorCandidato = nombreOpcionWeb;
            }
        }

        return { nombreWeb: mejorCandidato, score: maxScore };
    }

    // =========================================================================
    // --- 6. INICIALIZACIÓN Y ENTORNO ---
    // =========================================================================
    function init() {
        console.log("[EDARLab➔EPSAR] Recuperando estado de ejecución...");
        inyectarEstilosGlobales();

        const buffer = StateManager.getBuffer();
        const estado = StateManager.getAutomatorState();

        if (buffer.length > 0) {
            window.EDARLab_Buffer = buffer;
            console.log("[EDARLab➔EPSAR] Variable global recuperada.");
        }

        if (estado) {
            console.log(`[EDARLab➔EPSAR] Estado de ejecución recuperado. Paso actual: "${estado.pasoActual}". Tarea actual: ${estado.tareaIndexActual + 1} de ${estado.totalTareas}`);

            if (estado.abortar === true) {
                console.log("[EDARLab➔EPSAR] Detectado final de ejecución. Reseteando variables.");
                limpiarYResetearTodoAZero();
                return;
            }

            // Chequeo de texto duplicado directo en el HTML por seguridad
            if (estado.pasoActual === "ESPERAR_GUARDADO") {
                const textoPagina = document.body.innerText.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                if (textoPagina.includes("YA EXISTE") || textoPagina.includes("EXISTE UN PARTE") || textoPagina.includes("PARTE DUPLICADO")) {
                    console.warn("[EDARLab] Advertencia HTML de duplicado detectada en pantalla.");
                    registrarParteExistenteYContinuar();
                    return;
                }
            }

            activarBloqueoSuspension();
            reconstruirVentanaProgresoFijaCentrada(estado);
        } else {
            crearBotonLanzador();
        }
    }

    function inyectarEstilosGlobales() {
        if (document.getElementById("edarlab-styles")) return;
        const style = document.createElement("style");
        style.id = "edarlab-styles";
        style.innerText = `@keyframes entradaUI { from { opacity: 0; transform: scale(0.95) translate(-50%, -50%); } to { opacity: 1; transform: scale(1) translate(-50%, -50%); } }
        .edar-tabla { width: 100%; border-collapse: collapse; font-size: 11px; margin-top: 10px; text-align: left; }
        .edar-tabla th { background: #f1f5f9; color: #475569; padding: 6px 8px; font-weight: 600; border-bottom: 2px solid #e2e8f0; }
        .edar-tabla td { padding: 6px 8px; border-bottom: 1px solid #f1f5f9; color: #1e293b; }
        .edar-tabla tr:hover { background: #f8fafc; }
        .edar-cb { cursor: pointer; width: 14px; height: 14px; }
        .edarlab-launcher { position: fixed; top: 20px; right: 20px; z-index: 10000; background: linear-gradient(135deg, ${STYLE.PRIMARIO}, ${STYLE.SECUNDARIO}); color: white; padding: 10px 16px; border-radius: 20px; font-family: ${STYLE.FONT}; font-size: 12px; font-weight: bold; cursor: pointer; border: none; box-shadow: 0 4px 12px rgba(0,0,0,0.15); transition: all 0.2s ease; }`;
        document.head.appendChild(style);
    }

    function crearBotonLanzador() {
        if (document.getElementById("edar-launcher-btn")) return;
        botonLanzador = document.createElement("button");
        botonLanzador.id = "edar-launcher-btn";
        botonLanzador.className = "edarlab-launcher";
        botonLanzador.innerText = "Análisis en Influente - EDARLab➔EPSAR";
        botonLanzador.onclick = despertarEscuchaNativa;
        document.body.appendChild(botonLanzador);
    }

    // =========================================================================
    // --- 7. ESCUCHA DE PORTAPAPELES Y RENDER ---
    // =========================================================================
    function despertarEscuchaNativa() {
        escuchandoHerramienta = true;
        botonLanzador.innerText = "Copia los datos y presiona Ctrl + V";
        botonLanzador.style.background = `linear-gradient(135deg, ${STYLE.PRIMARIO}, ${STYLE.SECUNDARIO})`;
        botonLanzador.ondblclick = dormirUI;

        document.onpaste = function(e) {
            if (!escuchandoHerramienta) return;

            e.preventDefault();
            window.EDARLab_Buffer = [];
            sessionStorage.removeItem("edarlab_global_buffer_storage");

            const text = (e.clipboardData || window.clipboardData).getData("text"),
                  rows = text.split(/\r?\n/).map(e => e.split("\t"));

            let candidates = [];

            rows.forEach((row, idx) => {
                const rs = row.join(" ").toUpperCase();
                if (rs.includes("EDAR:")) {
                    const n = row.find(c => c.toUpperCase().includes("EDAR:"))?.split(":")[1]?.trim() || "";
                    candidates.push({ name: n, index: idx, dC: 0, sD: {} });
                }
            });

            if (!candidates.length) {
                alert("Ninguna EDAR detectada / Formato incorrecto");
                return;
            }

            candidates.forEach((cand) => {
                let siguienteEdarIndex = rows.length;
                for (let next = cand.index + 1; next < rows.length; next++) {
                    if (rows[next].join(" ").toUpperCase().includes("EDAR:")) {
                        siguienteEdarIndex = next;
                        break;
                    }
                }
                for (let j = cand.index + 1; j < siguienteEdarIndex; j++) {
                    const rRow = rows[j],
                          rD = rRow.find(c => /^\d{2}\/\d{2}\//.test(c.trim()));
                    if (rD) {
                        const parts = rD.trim().split("/");
                        const day = parseInt(parts[0], 10);
                        const dayStr = parts[0].padStart(2, '0');
                        const monthStr = parts[1] ? parts[1].padStart(2, '0') : "";
                        let yearFull = parts[2] ? parts[2].split(" ")[0].trim() : "";
                        if (yearFull.length === 2) yearFull = "20" + yearFull;
                        const fechaFormateada = `${dayStr}/${monthStr}/${yearFull}`;

                        const rowData = {};

                        for (const [k, idx] of Object.entries(EXCEL_COLUMNS)) {
                            let v = rRow[idx] ? rRow[idx].trim() : "";
                            if (v && v !== "-" && v !== "---") {
                                if (v.includes(".") && v.includes(",")) v = v.replace(/\./g, "");
                                v = v.replace(".", ",");
                                rowData[k] = v;
                                rowData[k + "_activo"] = true; // NUEVO: Inicializa el parámetro como activo
                                cand.dC++;
                            }
                        }

                        if (Object.keys(rowData).length > 0) {
                            rowData.procesar = true;
                            rowData.fechaFormateada = fechaFormateada;
                            rowData.muestreo = "PUNTUAL";
                            cand.sD[day] = rowData;
                        }
                    }
                }
                const coincidencia = buscarMejorCoincidenciaWeb(cand.name);
                window.EDARLab_Buffer.push({
                    nameExcel: cand.name,
                    nameWebAsociada: coincidencia.nombreWeb,
                    dC: cand.dC,
                    sD: cand.sD,
                    procesar: true
                });
            });

            StateManager.setBuffer(window.EDARLab_Buffer);
            escuchandoHerramienta = false;
            document.onpaste = null;
            botonLanzador.style.display = "none";
            despertarUI(window.EDARLab_Buffer);
        };
    }

    function despertarUI(listaEdars) {
        if (uiContenedor) uiContenedor.remove();
        uiContenedor = document.createElement("div");
        uiContenedor.id = "edar-deck-ui";

        Object.assign(uiContenedor.style, {
            position: "fixed", top: "50%", left: "50%", transform: "translate(-50%, -50%)",
            width: "700px",
            background: "#ffffff", color: "#1e293b",
            borderRadius: "16px", zIndex: "100000", fontFamily: STYLE.FONT,
            boxShadow: "0 20px 50px rgba(0,0,0,0.3)", border: "1px solid rgba(0,151,215,0.3)", overflow: "hidden"
        });

        uiContenedor.innerHTML = `
            <div style="background:linear-gradient(135deg,${STYLE.PRIMARIO},${STYLE.SECUNDARIO}); padding:14px 16px; display:flex; justify-content:space-between; align-items:center; border-top-left-radius:16px; border-top-right-radius:16px;">
                <span style="color:white; font-size:14px; font-weight:700;">Selector de partes - Análisis en Influente</span>
                <span id="close-ui" style="cursor:pointer; color:white; font-size:22px; font-weight:bold;">&times;</span>
            </div>
            <div id="edar-panel-body" style="padding:16px; background:#ffffff;">
                <div style="max-height: 280px; overflow-y: auto; border: 1px solid #e2e8f0; border-radius: 8px; background: white;">
                    <table class="edar-tabla" id="tabla-edars-body">
                        <tbody></tbody>
                    </table>
                </div>
                <div style="margin-top: 16px; display: flex; justify-content: flex-end; gap: 10px;">
                    <button id="btn-cancelar-ui" style="padding: 8px 14px; background: #e2e8f0; color: #475569; border: none; border-radius: 8px; font-weight: 600; cursor: pointer;">Cancelar</button>
                    <button id="btn-ejecutar-ui" style="padding: 8px 16px; background: ${STYLE.SECUNDARIO}; color: white; border: none; border-radius: 8px; font-weight: 600; cursor: pointer;">Ejecutar</button>
                </div>
            </div>
            <div style="background:#f8fafc; padding:8px 16px; border-top:1px solid #f1f5f9; display:flex; justify-content:space-between;"><span id="txt-status-pie" style="font-size:10px; color:${STYLE.GRIS};">Haz clic en los botones para alternar el tipo de muestreo o apagar parámetros.</span></div>
        `;
        document.body.appendChild(uiContenedor);
        renderizarFilasTabla(window.EDARLab_Buffer);

        document.getElementById("close-ui").onclick = abortarYLimpiarTodo;
        document.getElementById("btn-cancelar-ui").onclick = abortarYLimpiarTodo;

        // Se elimina la lógica de "cb-select-all" para evitar errores de ejecución en la consola
        document.getElementById("btn-ejecutar-ui").onclick = iniciarSecuenciaPersistente;
    }

    function renderizarFilasTabla(listaEdars) {
        const tbody = document.querySelector("#tabla-edars-body tbody");
        if (!tbody) return;
        tbody.innerHTML = "";

        listaEdars.forEach((cand, edarIndex) => {
            if (!cand.sD || Object.keys(cand.sD).length === 0) return;

            // --- 1. CABECERA DE BLOQUE CON CHECKBOX PROPIO POR EDAR ---
            // Comprobamos si todos los días de esta EDAR están marcados para procesar
            const todosMarcados = Object.values(cand.sD).every(d => d.procesar);

            const filaHeader = document.createElement("tr");
            filaHeader.style.cssText = "background: #f1f5f9 !important; font-weight: bold; border-bottom: 2px solid #cbd5e1;";
            filaHeader.innerHTML = `
                <td style="text-align:center; padding: 8px 12px;">
                    <input type="checkbox" class="edar-grupo-cb" data-edar="${edarIndex}" ${todosMarcados ? "checked" : ""} style="cursor:pointer; width:14px; height:14px;">
                </td>
                <td colspan="3" style="padding: 8px 4px; font-size: 12px; color: ${STYLE.PRIMARIO}; font-weight: 700; user-select:none;">
                    ${cand.nameExcel} <span style="color: ${STYLE.GRIS}; font-weight: 500;">➔ Portal: ${cand.nameWebAsociada}</span>
                </td>
            `;
            tbody.appendChild(filaHeader);

            // Evento para marcar/desmarcar todo el bloque de esta EDAR de golpe
            filaHeader.querySelector(".edar-grupo-cb").onchange = function(e) {
                const eIdx = this.dataset.edar;
                Object.keys(window.EDARLab_Buffer[eIdx].sD).forEach(day => {
                    window.EDARLab_Buffer[eIdx].sD[day].procesar = e.target.checked;
                });
                StateManager.setBuffer(window.EDARLab_Buffer);
                renderizarFilasTabla(window.EDARLab_Buffer); // Refrescamos la UI
            };


            // --- 2. FILAS DIARIAS SIMPLIFICADAS ---
            Object.keys(cand.sD).forEach(day => {
                const dataDia = cand.sD[day];
                const fila = document.createElement("tr");
                fila.style.cssText = "background: #ffffff !important; border-bottom: 1px solid #f1f5f9;";

                if (!dataDia.muestreo) dataDia.muestreo = "PUNTUAL";

                // CONFIGURACIÓN DE PASTILLA ÚNICA (INTERRUPTOR)
                const esPuntual = dataDia.muestreo === "PUNTUAL";
                const bgMuestreo = esPuntual ? STYLE.PRIMARIO : "#ea580c"; // Azul o Naranja corporativo

                // Estilos dinámicos para los parámetros químicos
                const pillsHTML = ORDEN_ESCRITURA
                    .filter(k => dataDia[k] !== undefined && dataDia[k] !== "")
                    .map(k => {
                        const activo = dataDia[k + "_activo"] !== false;
                        const bg = activo ? STYLE.SECUNDARIO : "#cbd5e1";
                        const color = activo ? "white" : "#64748b";
                        const textDecoration = activo ? "none" : "line-through";
                        const opacity = activo ? "1" : "0.6";

                        return `<span class="edar-pill"
                                      data-edar="${edarIndex}"
                                      data-day="${day}"
                                      data-param="${k}"
                                      style="background:${bg}; color:${color}; text-decoration:${textDecoration}; opacity:${opacity}; padding:2px 8px; border-radius:10px; font-size:9px; font-weight:bold; margin-right:4px; cursor:pointer; user-select:none; display:inline-block; transition:all 0.1s ease;">
                                    ${k}: ${dataDia[k]}
                                </span>`;
                    })
                    .join("");

                fila.innerHTML = `
                    <td style="text-align:center; padding: 6px 8px;">
                        <input type="checkbox" class="edar-cb edar-fila-cb" data-edar="${edarIndex}" data-day="${day}" ${dataDia.procesar ? "checked" : ""}>
                    </td>
                    <td style="font-weight: 600; font-size: 11px; padding: 6px 8px; color: #475569;">${dataDia.fechaFormateada}</td>
                    <td style="text-align:center; padding: 6px 8px;">
                        <!-- PASTILLA INTERRUPTOR ÚNICA (Reduce el 50% del espacio visual de esta columna) -->
                        <span class="m-pill" data-edar="${edarIndex}" data-day="${day}" data-current="${dataDia.muestreo}" style="background:${bgMuestreo}; color:white; padding:3px 8px; border-radius:6px; font-size:9px; font-weight:bold; cursor:pointer; user-select:none; transition:all 0.1s; display:inline-block; width:70px; text-align:center;">
                            ${dataDia.muestreo}
                        </span>
                    </td>
                    <td style="padding: 6px 8px; text-align:right;">${pillsHTML || '<span style="color:#94a3b8; font-style:italic; font-size:10px;">Sin parámetros</span>'}</td>
                `;

                // Evento Checkbox fila individual
                const cb = fila.querySelector(".edar-fila-cb");
                if (cb) {
                    cb.onchange = function() {
                        window.EDARLab_Buffer[edarIndex].sD[day].procesar = cb.checked;
                        StateManager.setBuffer(window.EDARLab_Buffer);

                        // Opcional: Desactivamos el checkbox del grupo si el usuario desmarca un día suelto
                        renderizarFilasTabla(window.EDARLab_Buffer);
                    };
                }

                // Evento clic sobre la pastilla Interruptor de Muestreo
                fila.querySelector(".m-pill").onclick = function() {
                    const eIdx = this.dataset.edar;
                    const dIdx = this.dataset.day;
                    const nuevoMuestreo = this.dataset.current === "PUNTUAL" ? "INTEGRADO" : "PUNTUAL";

                    window.EDARLab_Buffer[eIdx].sD[dIdx].muestreo = nuevoMuestreo;
                    StateManager.setBuffer(window.EDARLab_Buffer);
                    renderizarFilasTabla(window.EDARLab_Buffer);
                };

                // Evento interactivo para los Parámetros Químicos (N)
                fila.querySelectorAll(".edar-pill").forEach(pill => {
                    pill.onclick = function() {
                        const eIdx = this.dataset.edar;
                        const dIdx = this.dataset.day;
                        const param = this.dataset.param;

                        const estadoActual = window.EDARLab_Buffer[eIdx].sD[dIdx][param + "_activo"] !== false;
                        window.EDARLab_Buffer[eIdx].sD[dIdx][param + "_activo"] = !estadoActual;

                        StateManager.setBuffer(window.EDARLab_Buffer);
                        renderizarFilasTabla(window.EDARLab_Buffer);
                    };
                });

                tbody.appendChild(fila);
            });
        });
    }

    // =========================================================================
    // --- 8. EJECUCIÓN SECUENCIAL (MAQUINA DE ESTADOS) ---
    // =========================================================================
    function iniciarSecuenciaPersistente() {
        activarBloqueoSuspension();

        const tareas = [];
        window.EDARLab_Buffer.forEach((c, edarIdx) => {
            Object.keys(c.sD).forEach(day => {
                const dataDia = c.sD[day];

                // Filtro de seguridad: Comprobamos si tiene al menos un parámetro activado por el usuario
                const tieneParametrosActivos = ORDEN_ESCRITURA.some(k => dataDia[k] !== undefined && dataDia[k] !== "" && dataDia[k + "_activo"] !== false);

                if (dataDia.procesar && tieneParametrosActivos) {
                    tareas.push({
                        edarIdx: edarIdx,
                        day: parseInt(day, 10),
                        fecha: dataDia.fechaFormateada,
                        muestreo: dataDia.muestreo || "PUNTUAL", // Almacenamos el muestreo dinámico seleccionado
                        valores: {
                            NNH4: dataDia.NNH4,
                            NNH4_activo: dataDia.NNH4_activo !== false,
                            NNO2: dataDia.NNO2,
                            NNO2_activo: dataDia.NNO2_activo !== false,
                            NNO3: dataDia.NNO3,
                            NNO3_activo: dataDia.NNO3_activo !== false,
                            NT: dataDia.NT,
                            NT_activo: dataDia.NT_activo !== false,
                            PT: dataDia.PT,
                            PT_activo: dataDia.PT_activo !== false
                        }
                    });
                }
            });
        });

        if (tareas.length === 0) {
            alert("Selecciona al menos una fecha para procesar en la tabla.");
            return;
        }

        console.log(`[EDARLab➔EPSAR] Lanzando ejecución de ${tareas.length} partes de análisis en influente.`);

        const nuevoEstado = {
            tareas: tareas,
            tareaIndexActual: 0,
            totalTareas: tareas.length,
            pasoActual: "CLIC_NUEVO_PARTE",
            abortar: false
        };

        const fechaHoy = new Date().toLocaleString();
        const versionScript = (typeof GM_info !== 'undefined' && GM_info.script) ? GM_info.script.version : "0.0";
        const headerTxt = `==================================================\nREPORTE DE VOLCADO INFLUENTE EDARLab➔EPSAR (v${versionScript})\nFecha: ${fechaHoy}\n==================================================\n\n`;

        StateManager.setLog(headerTxt);
        StateManager.setAutomatorState(nuevoEstado);
        reconstruirVentanaProgresoFijaCentrada(nuevoEstado, true);
    }

    function reconstruirVentanaProgresoFijaCentrada(estado, delayStart = false) {
        if (uiContenedor) uiContenedor.remove();

        uiContenedor = document.createElement("div");
        uiContenedor.id = "edar-deck-ui";

        Object.assign(uiContenedor.style, {
            position: "fixed",
            top: "50%",
            left: "50%",
            transform: "translate(-50%, -50%)",
            width: "440px",
            background: "#fffffff2",
            backdropFilter: "blur(10px)",
            color: "#1e293b",
            borderRadius: "16px",
            zIndex: "100000",
            fontFamily: STYLE.FONT,
            border: "1px solid rgba(0,151,215,0.3)",
            boxShadow: "0 0 0 2000px rgba(71, 85, 105, 0.4), 0 20px 50px rgba(0, 0, 0, 0.3)"
        });

        uiContenedor.innerHTML = `
        <div style="background:linear-gradient(135deg,${STYLE.PRIMARIO},${STYLE.SECUNDARIO}); padding:12px 16px; display:flex; justify-content:space-between; align-items:center; border-top-left-radius:16px; border-top-right-radius:16px;">
            <span style="color:white; font-size:12px; font-weight:700;">EDARLab➔EPSAR</span>
            <span id="close-ui-progreso" style="cursor:pointer; color:white; font-size:18px; font-weight:bold;">&times;</span>
        </div>
        <div id="edar-panel-body" style="padding:18px;">
            <div id="tx-progreso-global" style="font-size:13px; font-weight:700; color:#004381; margin-bottom:2px; text-align:center;"></div>
            <div id="tx-progreso-detalle" style="font-size:11px; color:${STYLE.GRIS}; margin-bottom:12px; text-align:center;"></div>
            <div style="background:#e2e8f0; border-radius:4px; height:12px; overflow:hidden; box-shadow:inset 0 1px 2px rgba(0,0,0,0.1);">
                <div id="rb-barra" style="width:0%; height:100%; background:linear-gradient(90deg, ${STYLE.PRIMARIO}, ${STYLE.SECUNDARIO}); transition:width 0.3s ease;"></div>
            </div>
            <div id="tx-porcentaje" style="text-align:center; font-size:11px; margin-top:5px; font-weight:700; color:${STYLE.GRIS};">0%</div>
        </div>
        `;

        document.body.appendChild(uiContenedor);
        document.getElementById("close-ui-progreso").onclick = abortarYLimpiarTodo;

        window.addEventListener('click', bloquearAccionUsuario, true);
        window.addEventListener('mousedown', bloquearAccionUsuario, true);
        window.addEventListener('keydown', bloquearAccionUsuario, true);
        window.addEventListener('keypress', bloquearAccionUsuario, true);

        if (delayStart) {
            document.getElementById("tx-progreso-detalle").innerText = "Iniciando volcado de datos...";
            setTimeout(() => ejecutarMaquinaEstados(estado), 8000);
        } else {
            ejecutarMaquinaEstados(estado);
        }
    }

    function bloquearAccionUsuario(e) {
        if (!e.isTrusted) return;
        if (e.target.id === "close-ui-progreso" || e.target.innerText === "×") return;
        e.stopPropagation();
        e.preventDefault();
        console.log(`[EDARLab] Acción física de ${e.type} bloqueada para proteger la tabla.`);
    }

    function actualizarProgresoVisual(estado, subPasoIndex) {
        const totalSubPasos = 5;
        const totalPasosGlobales = estado.totalTareas * totalSubPasos;
        const pasoGlobalActual = (estado.tareaIndexActual * totalSubPasos) + subPasoIndex;

        const porcentaje = Math.min(Math.round((pasoGlobalActual / totalPasosGlobales) * 100), 99);

        const rbBarra = document.getElementById("rb-barra");
        const txPorcentaje = document.getElementById("tx-porcentaje");

        if (rbBarra) rbBarra.style.width = `${porcentaje}%`;
        if (txPorcentaje) txPorcentaje.innerText = `${porcentaje}% completado (${estado.tareaIndexActual} de ${estado.totalTareas} partes)`;
    }

    // =========================================================================
    // --- 9. SUB-FUNCIONES MODULARES PARA LOS PASOS DEL FORMULARIO ---
    // =========================================================================
    async function seleccionarEDARFormulario(nombreWeb) {
        const dropNuevaEdar = document.getElementById(SELECTORS.FORM_EDAR);
        if (!dropNuevaEdar) return false;

        let indexObjetivo = -1;
        for (let i = 0; i < dropNuevaEdar.options.length; i++) {
            if (dropNuevaEdar.options[i].text.trim() === nombreWeb) {
                indexObjetivo = i;
                break;
            }
        }

        if (indexObjetivo !== -1) {
            dropNuevaEdar.selectedIndex = indexObjetivo;
            dropNuevaEdar.value = dropNuevaEdar.options[indexObjetivo].value;
            dropNuevaEdar.dispatchEvent(new Event("change", { bubbles: true }));
            return true;
        }
        return false;
    }

    async function escribirFechaFormulario(fecha) {
        const inputFecha = document.getElementById(SELECTORS.FORM_FECHA);
        if (inputFecha) {
            inputFecha.focus();
            await typeH(inputFecha, fecha);
            return true;
        }
        return false;
    }

    function configurarMuestreoFormulario(tipoMuestreo) {
        const dropMuestreo = document.getElementById(SELECTORS.FORM_MUESTREO);
        if (!dropMuestreo) return false;

        let idxObjetivo = -1;
        const tipoBuscado = tipoMuestreo.toUpperCase().trim();

        for (let i = 0; i < dropMuestreo.options.length; i++) {
            if (dropMuestreo.options[i].text.trim().toUpperCase() === tipoBuscado) {
                idxObjetivo = i;
                break;
            }
        }
        if (idxObjetivo !== -1) {
            dropMuestreo.selectedIndex = idxObjetivo;
            dropMuestreo.value = dropMuestreo.options[idxObjetivo].value;
            dropMuestreo.dispatchEvent(new Event("change", { bubbles: true }));
            return true;
        }
        return false;
    }

    async function escribirValoresQuimicos(valores) {
        const camposInputs = {
            NNH4: SELECTORS.INPUT_NNH4,
            NNO3: SELECTORS.INPUT_NNO3,
            NNO2: SELECTORS.INPUT_NNO2,
            NT: SELECTORS.INPUT_NT,
            PT: SELECTORS.INPUT_PT
        };

        for (const [k, idInput] of Object.entries(camposInputs)) {
            // NUEVO: Si la pastilla de este parámetro se apagó en la UI, saltamos la escritura de este campo
            if (valores[k + "_activo"] === false) {
                console.log(`[EDARLab] Parámetro ${k} desactivado para esta fecha. Saltando campo.`);
                continue;
            }

            const valorAscribir = valores[k];
            if (valorAscribir !== undefined && valorAscribir !== "") {
                const elField = document.getElementById(idInput);
                if (elField) {
                    elField.focus();
                    await typeH(elField, valorAscribir.toString());
                }
            }
        }
    }

    // =========================================================================
    // --- 10. MOTOR CENTRAL (MÁQUINA DE ESTADOS REFACTORIZADA) ---
    // =========================================================================
    async function ejecutarMaquinaEstados(estado) {
        if (estado.abortar === true) {
            console.log("[EDARLab➔EPSAR] Bandera detectada. Abortando ejecución.");
            limpiarYResetearTodoAZero();
            return;
        }

        // CONTROL DE FINAL DE PROCESADO
        if (estado.tareaIndexActual >= estado.totalTareas) {
            console.log("[EDARLab➔EPSAR] Bucle de partes finalizado con éxito.");
            sessionStorage.removeItem("edar_automator_state");
            sonarAlertaFin();
            document.getElementById("rb-barra").style.width = "100%";
            document.getElementById("tx-porcentaje").innerText = `100% Completado (${estado.totalTareas} de ${estado.totalTareas})`;
            document.getElementById("tx-progreso-global").innerText = "Ejecución finalizada";
            document.getElementById("tx-progreso-detalle").innerText = `Procesados un total de ${estado.totalTareas} partes individuales.`;
            document.getElementById("close-ui-progreso").onclick = limpiarYResetearTodoAZero;
            descargarReporteTxtFinal();
            liberarBloqueoSuspension();
            return;
        }

        const tareaActual = estado.tareas[estado.tareaIndexActual];
        const buffer = StateManager.getBuffer();
        const candActual = buffer[tareaActual.edarIdx];
        const nombreWebActual = candActual.nameWebAsociada;

        document.getElementById("tx-progreso-global").innerText = `Parte ${estado.tareaIndexActual + 1} de ${estado.totalTareas}`;

        // --- PASO 1: CLICK EN "NUEVO PARTE" ---
        if (estado.pasoActual === "CLIC_NUEVO_PARTE") {
            actualizarProgresoVisual(estado, 0);
            document.getElementById("tx-progreso-detalle").innerText = `Abriendo formulario para ${nombreWebActual}...`;

            const btnNuevoParte = document.getElementById(SELECTORS.BOTON_NUEVO_PARTE);
            if (btnNuevoParte) {
                estado.pasoActual = "RELLENAR_FORMULARIO";
                StateManager.setAutomatorState(estado);
                await delay(CONFIG_PAUSAS.RENDERIZADO_TABLA);
                DOMUtils.click(btnNuevoParte);
                return;
            } else {
                if (document.getElementById(SELECTORS.FORM_EDAR)) {
                    estado.pasoActual = "RELLENAR_FORMULARIO";
                    StateManager.setAutomatorState(estado);
                    ejecutarMaquinaEstados(estado);
                    return;
                }
                abortarPorFaltaDeElementos();
                return;
            }
        }

        // --- PASO 2: INTRODUCCIÓN DE DATOS EN EL NUEVO PARTE ---
        if (estado.pasoActual === "RELLENAR_FORMULARIO") {
            actualizarProgresoVisual(estado, 0);
            document.getElementById("tx-progreso-detalle").innerText = `Seleccionando EDAR...`;
            await delay(CONFIG_PAUSAS.RENDERIZADO_TABLA);

            const edarSeleccionada = await seleccionarEDARFormulario(nombreWebActual);
            if (!edarSeleccionada) {
                abortarPorFaltaDeElementos();
                return;
            }

            let estCheck = StateManager.getAutomatorState();
            if (estCheck?.abortar === true || estCheck?.tareaIndexActual === undefined) {
                limpiarYResetearTodoAZero();
                return;
            }

            // Subpaso 1: Introduciendo fecha
            actualizarProgresoVisual(estado, 1);
            await delay(CONFIG_PAUSAS.ASIMILACION_DESPLEGABLE);
            document.getElementById("tx-progreso-detalle").innerText = `Escribiendo fecha...`;
            await escribirFechaFormulario(tareaActual.fecha);

            // Subpaso 2: Definiendo tipo de análisis
            actualizarProgresoVisual(estado, 2);
            await delay(CONFIG_PAUSAS.ASIMILACION_DESPLEGABLE);
            document.getElementById("tx-progreso-detalle").innerText = `Configurando tipo de muestreo...`;
            configurarMuestreoFormulario(tareaActual.muestreo);

            // Subpaso 3: Introduciendo parámetros
            actualizarProgresoVisual(estado, 3);
            await delay(CONFIG_PAUSAS.ASIMILACION_DESPLEGABLE);
            document.getElementById("tx-progreso-detalle").innerText = `Escribiendo parámetros...`;
            await escribirValoresQuimicos(tareaActual.valores);

            await delay(CONFIG_PAUSAS.PRE_RECALCULA);

            estCheck = StateManager.getAutomatorState();
            if (estCheck?.abortar === true || estCheck?.tareaIndexActual === undefined) {
                limpiarYResetearTodoAZero();
                return;
            }

            // Subpaso 4: Guardando parte
            actualizarProgresoVisual(estado, 4);
            document.getElementById("tx-progreso-detalle").innerText = `Registrando datos...`;

            const btnGuardar = document.getElementById(SELECTORS.BOTON_GUARDAR);
            if (btnGuardar) {
                estado.pasoActual = "ESPERAR_GUARDADO";
                StateManager.setAutomatorState(estado);
                DOMUtils.click(btnGuardar);
                return;
            } else {
                abortarPorFaltaDeElementos();
                return;
            }
        }

        // --- PASO 3: CONTROL DE TRANSICIÓN POST-GUARDADO ---
        if (estado.pasoActual === "ESPERAR_GUARDADO") {
            actualizarProgresoVisual(estado, 4);
            document.getElementById("tx-progreso-detalle").innerText = "Confirmando guardado en el servidor...";
            await delay(CONFIG_PAUSAS.POST_RECALCULA);

            // 1. Filtro anticarreras por si el popup de duplicidad fue interceptado de forma paralela
            const estCheckFresh = StateManager.getAutomatorState();
            if (estCheckFresh && estCheckFresh.pasoActual !== "ESPERAR_GUARDADO") {
                console.log("[EDARLab] Duplicado detectado y procesado asíncronamente. Cancelando escritura fantasma.");
                return;
            }

            // 2. NUEVA RED DE SEGURIDAD: Control de texto de error inyectado por AJAX tras la espera
            const textoPagina = document.body.innerText.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            if (textoPagina.includes("YA EXISTE") || textoPagina.includes("EXISTE UN PARTE") || textoPagina.includes("PARTE DUPLICADO")) {
                console.warn("[EDARLab] Detectado texto de duplicidad inyectado por el servidor. Forzando omisión y log de OMITIDO...");
                registrarParteExistenteYContinuar(); // Esto escribirá OMITIDO en el TXT y redirigirá limpiamente
                return;
            }

            sonarAlertaPlanta();

            let log = StateManager.getLog();
            const fechaPlanta = new Date().toLocaleTimeString();
            // log += `[${fechaPlanta}] EDAR ${candActual.nameExcel} (${nombreWebActual}) | Fecha: ${tareaActual.fecha} | Tipo: ${tareaActual.muestreo} | NNH4: ${tareaActual.valores.NNH4 || '-'} | NNO2: ${tareaActual.valores.NNO2 || '-'} | NNO3: ${tareaActual.valores.NNO3 || '-'} | NT: ${tareaActual.valores.NT || '-'} | PT: ${tareaActual.valores.PT || '-'}\n`;
            // log += `--------------------------------------------------\n\n`;
            // 1. Mapeamos y filtramos solo los parámetros que están activos y tienen un valor asignado
            const parametrosEscritos = ORDEN_ESCRITURA
                .filter(k => tareaActual.valores[k + "_activo"] !== false && tareaActual.valores[k] !== undefined && tareaActual.valores[k] !== "")
                .map(k => `${k}: ${tareaActual.valores[k]}`)
                .join(" | ");
            
            // 2. Si no hubiera ninguno activo (caso extremo), indicamos que no se enviaron parámetros
            const detalleParamsTxt = parametrosEscritos ? `| ${parametrosEscritos}` : "| Sin parámetros introducidos";
            
            // 3. Construimos la línea limpia del log
            log += `[${fechaPlanta}] EDAR ${candActual.nameExcel} (${nombreWebActual}) | Fecha: ${tareaActual.fecha} | Tipo: ${tareaActual.muestreo} ${detalleParamsTxt}\n`;
            log += `--------------------------------------------------\n\n`;

            StateManager.setLog(log);

            estado.tareaIndexActual++;
            estado.pasoActual = "CLIC_NUEVO_PARTE";
            StateManager.setAutomatorState(estado);

            document.getElementById("tx-progreso-detalle").innerText = "Parte procesado con éxito...";
            await delay(CONFIG_PAUSAS.TRANSICION_PLANTA);

            // window.location.href = "https://aplica.epsar.gva.es/depuradoras/Partes/AnalisisInfluente.aspx?res=E";
            ejecutarMaquinaEstados(estado)
            return;
        }
    }

    // =========================================================================
    // --- 11. AUDIO, DESCARGA Y LIMPIEZA ---
    // =========================================================================
    function descargarReporteTxtFinal() {
        let textoFinal = StateManager.getLog();
        if (!textoFinal || textoFinal.trim().length < 150) return;
        textoFinal += `\n==================================================\nFIN DEL INFORME \n==================================================`;
        const blob = new Blob([textoFinal], { type: "text/plain;charset=utf-8" });
        const urlDescarga = URL.createObjectURL(blob);
        const fecha = new Date().toISOString().slice(0, 10);
        const horaClean = new Date().toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' }).replace(':', '-');
        const enlacelink = document.createElement("a");
        enlacelink.href = urlDescarga;
        enlacelink.download = `Reporte_AnalisisInfluente_EPSAR_${fecha}_${horaClean}.txt`;
        document.body.appendChild(enlacelink);
        enlacelink.click();
        document.body.removeChild(enlacelink);
        URL.revokeObjectURL(urlDescarga);
        console.log("[EDARLab➔EPSAR] Reporte de partes descargado con éxito.");
    }

    function abortarYLimpiarTodo() {
        descargarReporteTxtFinal();
        const estado = StateManager.getAutomatorState();
        if (estado && estado.tareaIndexActual < estado.totalTareas) {
            estado.abortar = true;
            StateManager.setAutomatorState(estado);
        }
        limpiarYResetearTodoAZero();
    }

    function limpiarYResetearTodoAZero() {
        console.log("[EDARLab➔EPSAR] Reseteando todas las variables y limpiando buffers.");
        window.removeEventListener('click', bloquearAccionUsuario, true);
        window.removeEventListener('mousedown', bloquearAccionUsuario, true);
        window.removeEventListener('keydown', bloquearAccionUsuario, true);
        window.removeEventListener('keypress', bloquearAccionUsuario, true);

        StateManager.clearAll();

        if (uiContenedor) {
            uiContenedor.remove();
            uiContenedor = null;
        }

        document.onpaste = null;
        escuchandoHerramienta = false;

        let botonExistente = document.getElementById("edar-launcher-btn");
        if (botonExistente) {
            botonExistente.innerText = "Análisis en Influente - EDARLab➔EPSAR";
            botonExistente.style.background = `linear-gradient(135deg, ${STYLE.PRIMARIO}, ${STYLE.SECUNDARIO})`;
            botonExistente.style.display = "block";
            botonExistente.ondblclick = null;
        } else {
            crearBotonLanzador();
        }
        liberarBloqueoSuspension();
    }

    function dormirUI() {
        limpiarYResetearTodoAZero();
    }

    if (document.readyState === "complete" || document.readyState === "interactive") {
        init();
    } else {
        document.addEventListener("DOMContentLoaded", init);
    }

    function sonarAlertaFin() {
        try {
            const audioCtx = new (window.AudioContext || window.webkitAudioContext)();

            const ahora1 = audioCtx.currentTime;
            let osc1 = audioCtx.createOscillator();
            let gain1 = audioCtx.createGain();
            osc1.type = 'triangle';
            osc1.frequency.setValueAtTime(523.25, ahora1);
            gain1.gain.setValueAtTime(0.35, ahora1);
            gain1.gain.exponentialRampToValueAtTime(0.001, ahora1 + 0.8);
            osc1.connect(gain1);
            gain1.connect(audioCtx.destination);
            osc1.start(ahora1);
            osc1.stop(ahora1 + 0.8);

            setTimeout(() => {
                const ahora2 = audioCtx.currentTime;
                let osc2 = audioCtx.createOscillator();
                let gain2 = audioCtx.createGain();
                osc2.type = 'triangle';
                osc2.frequency.setValueAtTime(659.25, ahora2);
                gain2.gain.setValueAtTime(0.32, ahora2);
                gain2.gain.exponentialRampToValueAtTime(0.001, ahora2 + 0.8);
                osc2.connect(gain2);
                gain2.connect(audioCtx.destination);
                osc2.start(ahora2);
                osc2.stop(ahora2 + 0.8);
            }, 120);

            setTimeout(() => {
                const ahora3 = audioCtx.currentTime;
                let osc3 = audioCtx.createOscillator();
                let gain3 = audioCtx.createGain();
                osc3.type = 'triangle';
                osc3.frequency.setValueAtTime(783.99, ahora3);
                gain3.gain.setValueAtTime(0.30, ahora3);
                gain3.gain.exponentialRampToValueAtTime(0.001, ahora3 + 1.0);
                osc3.connect(gain3);
                gain3.connect(audioCtx.destination);
                osc3.start(ahora3);
                osc3.stop(ahora3 + 1.0);
            }, 240);

        } catch (e) {
            console.log("[EDARLab] No se pudo reproducir la notificación de triple tono: ", e);
        }
    }

    function sonarAlertaPlanta() {
        try {
            const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
            const ahora = audioCtx.currentTime;

            let osc = audioCtx.createOscillator();
            let gain = audioCtx.createGain();
            osc.type = 'triangle';
            osc.frequency.setValueAtTime(659.25, ahora);

            gain.gain.setValueAtTime(0.35, ahora);
            gain.gain.exponentialRampToValueAtTime(0.001, ahora + 0.6);

            osc.connect(gain);
            gain.connect(audioCtx.destination);
            osc.start(ahora);
            osc.stop(ahora + 0.6);
        } catch (e) {
            console.log("[EDARLab] No se pudo reproducir el tono de planta: ", e);
        }
    }

    function abortarPorFaltaDeElementos() {
        console.error(`[EDARLab➔EPSAR] CRÍTICO: No se encontraron los elementos necesarios en el portal.`);
        try {
            descargarReporteTxtFinal();
        } catch (err) {
            console.error("[EDARLab] Error al descargar el reporte en el aborto: ", err);
        }
        sessionStorage.removeItem("edar_automator_state");

        if (typeof limpiarYResetearTodoAZero === "function") {
            limpiarYResetearTodoAZero();
        } else {
            window.removeEventListener('click', bloquearAccionUsuario, true);
            window.removeEventListener('mousedown', bloquearAccionUsuario, true);
            window.removeEventListener('keydown', bloquearAccionUsuario, true);
            window.removeEventListener('keypress', bloquearAccionUsuario, true);
        }
        liberarBloqueoSuspension();
        alert(`ERROR: No se han podido cargar los campos o botones de la web. Revisa el estado del portal e inténtalo de nuevo.`);
    }

    async function activarBloqueoSuspension() {
        try {
            if ('wakeLock' in navigator) {
                wakeLock = await navigator.wakeLock.request('screen');
                console.log("[EDARLab➔WakeLock] ☀️ Sistema de prevención de suspensión de pantalla activado.");

                document.addEventListener('visibilitychange', async () => {
                    if (wakeLock !== null && document.visibilityState === 'visible') {
                        wakeLock = await navigator.wakeLock.request('screen');
                    }
                });
            } else {
                console.warn("[EDARLab➔WakeLock] Tu navegador no soporta la API Wake Lock.");
            }
        } catch (err) {
            console.warn("[EDARLab➔WakeLock] Error al bloquear la suspensión: ", err.message);
        }
    }

    function liberarBloqueoSuspension() {
        if (wakeLock !== null) {
            wakeLock.release().then(() => {
                wakeLock = null;
                console.log("[EDARLab➔WakeLock] 🌙 Bloqueo liberado. El ordenador ya puede suspenderse con normalidad.");
            });
        }
    }
})();
