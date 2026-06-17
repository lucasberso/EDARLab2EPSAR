// ==UserScript==
// @name         Parte Mensual de Analítica - EDARLab➔EPSAR - GIT
// @version      3.1
// @description  Herramienta que automatiza la introducción de partes de analíticas en el portal de la EPSAR.
// @author       Lucas B.
// @match        https://aplica.epsar.gva.es/depuradoras/Partes/MensualAnalitica.aspx*
// @grant        none
// @downloadURL  https://github.com/lucasberso/EDARLab2EPSAR/raw/refs/heads/main/ParteMensualEPSAR.user.js
// @updateURL    https://github.com/lucasberso/EDARLab2EPSAR/raw/refs/heads/main/ParteMensualEPSAR.user.js

// ==/UserScript==

(function() {
    'use strict';

    // =========================================================================
    // --- TABLA DE CONFIGURACIÓN DE PAUSAS ---
    // =========================================================================
    const diaDelMes = new Date().getDate();
    const FACTOR_SATURACION = (diaDelMes >= 1 && diaDelMes <= 4) ? 1.2 : 1.0;

    console.log(`[EDARLab➔EPSAR] Día del mes: ${diaDelMes}. Factor de retraso aplicado: x${FACTOR_SATURACION}`);
    const CONFIG_PAUSAS = {
        // Pausas internas de escritura en tabla
        ENTRE_CARACTERES: () => (500 + Math.random() * 500) * FACTOR_SATURACION, // Retraso aleatorio entre caracteres
        POST_CELDA: () => (3000 + Math.random() * 4000) * FACTOR_SATURACION, // Retraso aleatorio tras desenfocar celda
        POST_DIA_FILA: (4000 + Math.random() * 3000) * FACTOR_SATURACION,  // Pausa tras rellenar una fila diaria completa

        // Pausas de navegación y comunicación con el servidor
        APERTURA_DESPLEGABLE: 5000 + Math.random() * 3000,    // Espera tras hacer clic en el desplegable (Apertura)
        ASIMILACION_DESPLEGABLE: 5000 + Math.random() * 1000, // Espera para que la web asimile de forma pasiva la planta elegida
        RENDERIZADO_TABLA: (14000 + Math.random() * 5000) * FACTOR_SATURACION,      // Espera tras pulsar 'Mostrar' para que se dibuje la nueva tabla
        PRE_RECALCULA: 6000 + Math.random() * 4000,           // Espera tras escribir el último dato y antes de pulsar 'Recalcula'
        POST_RECALCULA: (14000 + Math.random() * 4000) * FACTOR_SATURACION,         // Espera tras el refresco de página provocado por 'Recalcula'
        TRANSICION_PLANTA: (6000 + Math.random() * 4000) * FACTOR_SATURACION        // Espera informativa en ventana antes de saltar a la siguiente EDAR
    };

    // --- CONFIGURACIÓN DE COLORES E INTERFAZ ---
    const CONFIG_PRIMARIO = "#004381";
    const CONFIG_SECUNDARIO = "#0097D7";
    const CONFIG_GRIS = "#64748b";
    const CONFIG_FONT = "'Segoe UI', sans-serif";

    // --- MATRIZ DE MAPEO DE COLUMNAS PORTAPAPELES (Tabla Excel) ---
    const p = {
        E: { PHEU: 4, TURBIDEZEU: 7, V60EU: 9, SSEU: 10, DBOEU: 13, DQOEU: 16, NTEU: 19, PTEU: 22 },
        S: { PHS: 5, CONDUCTIVIDAD: 6, TURBIDEZS: 8, SSS: 11, DBOS: 14, DQOS: 17, NTS: 20, PTS: 23 },
        F: { FDPH: 31, FDMV: 32, FDMS: 33 }
    };

    const ordenEscritura = ["PHEU","PHS","CONDUCTIVIDAD","TURBIDEZEU","TURBIDEZS","V60EU","SSEU","SSS","DBOEU","DBOS","DQOEU","DQOS","NTEU","NTS","PTEU","PTS","FDPH","FDMV","FDMS"];

    let uiContenedor = null;
    let botonLanzador = null;
    let escuchandoHerramienta = false;
    let wakeLock = null;

    // --- VARIABLE GLOBAL ÚNICA Y CENTRAL DE CONTROL (Guarda los datos al recargar la página) ---
    window.EDARLab_Buffer = [];

    // --- FUNCIONES DE LIMPIEZA Y SIMILITUD EXACTAS ---
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

    // --- FUNCIÓN DE ESCRITURA POR CARACTERES EN TABLA ---
    const typeH = async(el, txt) => {
        // 1. Inyección directa del valor completo (foco-proof)
        el.value = txt;
        // 2. Avisamos que se ha introducido texto
        el.dispatchEvent(new Event("input", { bubbles: true }));
        // Pequeña pausa de seguridad antes de salir de la celda
        await new Promise(r => setTimeout(r, 100 + Math.random() * 50)); // ~100ms
        // 3. Quitamos el foco (Esto hace que la web aplique sus formatos internos)
        el.dispatchEvent(new Event("blur", { bubbles: true }));
        await new Promise(r => setTimeout(r, 100 + Math.random() * 50));
        // 4. Disparamos el 'change' justo después del blur, imitando al navegador nativo
        el.dispatchEvent(new Event("change", { bubbles: true }));
        // 5. Pausa dinámica al terminar la celda.
        await new Promise(r => setTimeout(r, CONFIG_PAUSAS.POST_CELDA()));
    };

    // --- FUNCIÓN DE ESCRITURA POR CARACTERES EN TABLA ---
    //const typeH = async(el, txt) => {
        //el.dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));
        //el.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));
        //el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        //el.dispatchEvent(new FocusEvent("focus", { bubbles: true }));
        //el.value="";
        //for(const c of txt){
            //el.dispatchEvent(new KeyboardEvent('keydown',{key:c,bubbles:true}));
            //el.value+=c;
            //el.dispatchEvent(new InputEvent('input',{inputType:'insertText',data:c,bubbles:true}));
            //el.dispatchEvent(new KeyboardEvent('keyup',{key:c,bubbles:true}));
            // Pausa dinámica entre caracteres
            //await new Promise(r=>setTimeout(r, CONFIG_PAUSAS.ENTRE_CARACTERES()));
        //}

        //el.dispatchEvent(new Event("change",{bubbles:true}));
       // await new Promise(r => setTimeout(r, 40 + Math.random() * 60));
        //el.dispatchEvent(new Event("blur", { bubbles: true }));
        // Pausa dinámica al terminar la celda
        //await new Promise(r=>setTimeout(r, CONFIG_PAUSAS.POST_CELDA()));
    //};

    const delay = ms => new Promise(res => setTimeout(res, ms));

    // --- FUNCIÓN DE EMPAREJAMIENTO CON EL DESPLEGABLE DE LA WEB ---
    function buscarMejorCoincidenciaWeb(nombreExcel) {
        const r = document.getElementById("ctl00_ctl00_ContentPlaceHolder1_DropDownFiltroUnidadCoste");
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

    // --- FUNCIÓN INICIALIZACIÓN ---
    // Cada vez que la web se refresca por completo, la memoria RAM del navegador se borra.
    // Esta función se ejecuta automáticamente al nacer la nueva página para leer el "sessionStorage" y recordar qué estaba haciendo el script antes de la recarga y retomar la ejecución.
    function init() {
        console.log("[EDARLab➔EPSAR] Recuperando estado de ejecución...");
        inyectarEstilosGlobales();

        const datosEnDisco = sessionStorage.getItem("edarlab_global_buffer_storage");
        const estadoGuardado = sessionStorage.getItem("edar_automator_state");

        if (datosEnDisco) {
            window.EDARLab_Buffer = JSON.parse(datosEnDisco);
            console.log("[EDARLab➔EPSAR] Variable global recuperada.");
        }

        if (estadoGuardado) {
            const estado = JSON.parse(estadoGuardado);
            console.log(`[EDARLab➔EPSAR] Estado de ejecución recuperado. Paso actual: "${estado.pasoActual}". Índice global: ${estado.indiceBucleActual}`);

            if (estado.abortar === true) {
                console.log("[EDARLab➔EPSAR] Detectado final de ejecución. Reseteando la variable global.");
                limpiarYResetearTodoAZero();
                return;
            }

            //if (window.EDARLab_Buffer.length > 0 && estado.indiceBucleActual < window.EDARLab_Buffer.length) {
                //const plantaActualNombreWeb = window.EDARLab_Buffer[estado.indiceBucleActual].nameWebAsociada;
                //const desplegableWeb = document.getElementById("ctl00_ctl00_ContentPlaceHolder1_DropDownFiltroUnidadCoste");
                //if (desplegableWeb) {
                    //for (let i = 0; i < desplegableWeb.options.length; i++) {
                        //if (desplegableWeb.options[i].text.trim() === plantaActualNombreWeb) {
                            //desplegableWeb.selectedIndex = i;
                            //break;
                        //}
                    //}
                //}
            //}
            activarBloqueoSuspension();
            reconstruirVentanaProgresoFijaCentrada(estado);
        } else {
            crearBotonLanzador(); // En el caso de que no haya una tarea activa, devuelve el botón de inicialización.
        }
    }

    function inyectarEstilosGlobales() {
        if (document.getElementById("edarlab-styles")) return;
        const style = document.createElement("style");
        style.id = "edarlab-styles";
        style.innerText = "@keyframes entradaUI { from { opacity: 0; transform: scale(0.95) translate(-50%, -50%); } to { opacity: 1; transform: scale(1) translate(-50%, -50%); } } .edar-tabla { width: 100%; border-collapse: collapse; font-size: 11px; margin-top: 10px; text-align: left; } .edar-tabla th { background: #f1f5f9; color: #475569; padding: 6px 8px; font-weight: 600; border-bottom: 2px solid #e2e8f0; } .edar-tabla td { padding: 6px 8px; border-bottom: 1px solid #f1f5f9; color: #1e293b; } .edar-tabla tr:hover { background: #f8fafc; } .edar-cb { pointer; width: 14px; height: 14px; } .edarlab-launcher { position: fixed; top: 20px; right: 20px; z-index: 10000; background: linear-gradient(135deg, #004381, #0097D7); color: white; padding: 10px 16px; border-radius: 20px; font-family: 'Segoe UI', sans-serif; font-size: 12px; font-weight: bold; cursor: pointer; border: none; box-shadow: 0 4px 12px rgba(0,0,0,0.15); transition: all 0.2s ease; }";
        document.head.appendChild(style);
    }

    function crearBotonLanzador() {
        // Crea el botón de inicialización de la herramienta
        if (document.getElementById("edar-launcher-btn")) return;
        botonLanzador = document.createElement("button");
        botonLanzador.id = "edar-launcher-btn";
        botonLanzador.className = "edarlab-launcher";
        botonLanzador.innerText = "Analíticas - EDARLab➔EPSAR";
        botonLanzador.onclick = function() { despertarEscuchaNativa(); };
        document.body.appendChild(botonLanzador);
    }

    function despertarEscuchaNativa() {
        escuchandoHerramienta = true;
        botonLanzador.innerText = "Copia los datos y presiona Ctrl + V";
        botonLanzador.style.background = "linear-gradient(135deg, #004381, #0097D7)";
        botonLanzador.ondblclick = function() { dormirUI(); };

        document.onpaste = function(e) {
            if (!escuchandoHerramienta) return;

            e.preventDefault();
            window.EDARLab_Buffer = [];
            sessionStorage.removeItem("edarlab_global_buffer_storage");

            const text = (e.clipboardData || window.clipboardData).getData("text"),
                  rows = text.split(/\r?\n/).map(e => e.split("\t"));

            let candidates = [];

            // BUCLE DE LECTURA DEL PORTAPAPELES
            rows.forEach((row, idx) => {
                const rs = row.join(" ").toUpperCase();
                if (rs.includes("EDAR:")) {
                    const n = row.find(c => c.toUpperCase().includes("EDAR:"))?.split(":")[1]?.trim() || ""; // Importante que el split vaya acompañado de [1]
                    candidates.push({ name: n, index: idx, dC: 0, sD: {} });
                }
            });

            if (!candidates.length) {
                alert("Ninguna EDAR detectada / Formato incorrecto");
                return;
            }

            // Procesamos el conteo de celdas para cada candidato usando tu logica original
            candidates.forEach((cand, cIdx) => {
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
                        const day = parseInt(rD.trim().split("/"), 10);
                        const rowData = {};
                        ["E", "S", "F"].forEach(side => {
                            for (const [k, idx] of Object.entries(p[side])) {
                                let v = rRow[idx] ? rRow[idx].trim() : "";
                                if (v && v !== "-" && v !== "---") {
                                    if (v.includes(".") && v.includes(",")) v = v.replace(/\./g, "");
                                    v = v.replace(".", ",");
                                    if (k.includes("TURBIDEZ")) {
                                        let u = parseFloat(v.replace(",", "."));
                                        if (!isNaN(u)) v = Math.round(u).toString();
                                    }
                                    rowData[k] = v;
                                    cand.dC++;
                                }
                            }
                        });
                        cand.sD[day] = rowData;
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

            sessionStorage.setItem("edarlab_global_buffer_storage", JSON.stringify(window.EDARLab_Buffer));
            escuchandoHerramienta = false;
            document.onpaste = null;
            botonLanzador.style.display = "none";
            despertarUI(window.EDARLab_Buffer);
        };
    }

    function despertarUI(listaEdars) {
        // Interfaz gráfica de lectura de los datos
        if (uiContenedor) uiContenedor.remove();
        uiContenedor = document.createElement("div");
        uiContenedor.id = "edar-deck-ui";

        Object.assign(uiContenedor.style, {
            position: "fixed", top: "50%", left: "50%", transform: "translate(-50%, -50%)",
            width: "560px", background: "#fffffff2", backdropFilter: "blur(10px)", color: "#1e293b",
            borderRadius: "16px", zIndex: "10000", fontFamily: CONFIG_FONT,
            boxShadow: "0 20px 50px rgba(0,0,0,0.3)", border: "1px solid rgba(0,151,215,0.3)", overflow: "hidden"
        });

        // Estilos y textos de la interfaz
        uiContenedor.innerHTML = `
            <div style="background:linear-gradient(135deg,${CONFIG_PRIMARIO},${CONFIG_SECUNDARIO}); padding:14px 16px; display:flex; justify-content:space-between; align-items:center; border-top-left-radius:16px; border-top-right-radius:16px;">
                <span style="color:white; font-size:14px; font-weight:700;">Selector de partes</span>
                <span id="close-ui" style="cursor:pointer; color:white; font-size:22px; font-weight:bold;">&times;</span>
            </div>
            <div style="background: #f8fafc; padding: 10px 16px; border-bottom: 1px solid #e2e8f0; display: flex; align-items: center; gap: 8px;">
                <input type="checkbox" id="chk-sobreescribir" style="cursor: pointer; width: 16px; height: 16px;">
                <label for="chk-sobreescribir" style="font-size: 11px; font-weight: 600; color: #1e293b; cursor: pointer; user-select: none;">Sobreescribir valores existentes</label>
            </div>
            <div id="edar-panel-body" style="padding:16px;">
                <div style="max-height: 220px; overflow-y: auto; border: 1px solid #e2e8f0; border-radius: 8px; background: white;">
                    <table class="edar-tabla" id="tabla-edars-body">
                        <thead>
                            <tr>
                                <th style="width: 40px; text-align:center;"><input type="checkbox" id="cb-select-all" class="edar-cb" checked></th>
                                <th>Nombre - Excel</th>
                                <th>Nombre - EPSAR</th>
                                <th style="text-align:center;">Nº Parámetros</th>
                            </tr>
                        </thead>
                        <tbody></tbody>
                    </table>
                </div>
                <div style="margin-top: 16px; display: flex; justify-content: flex-end; gap: 10px;">
                    <button id="btn-cancelar-ui" style="padding: 8px 14px; background: #e2e8f0; color: #475569; border: none; border-radius: 8px; font-weight: 600; cursor: pointer;">Cancelar</button>
                    <button id="btn-ejecutar-ui" style="padding: 8px 16px; background: ${CONFIG_SECUNDARIO}; color: white; border: none; border-radius: 8px; font-weight: 600; cursor: pointer;">Ejecutar</button>
                </div>
            </div>
            <div style="background:#f8fafc; padding:8px 16px; border-top:1px solid #f1f5f9; display:flex; justify-content:space-between;"><span id="txt-status-pie" style="font-size:10px; color:${CONFIG_GRIS};">Selecciona los partes a procesar y comprueba que coincidan ambos nombres.</span></div>
        `;
        document.body.appendChild(uiContenedor);
        renderizarFilasTabla(window.EDARLab_Buffer);

        document.getElementById("close-ui").onclick = abortarYLimpiarTodo;
        document.getElementById("btn-cancelar-ui").onclick = abortarYLimpiarTodo;

        document.getElementById("cb-select-all").onchange = function(e) {
            document.querySelectorAll(".edar-fila-cb").forEach(cb => {
                cb.checked = e.target.checked;
                cb.dispatchEvent(new Event("change"));
            });
        };
        // Inicializamos el modo sobreescribir en falso por defecto
        sessionStorage.setItem("edarlab_sobreescribir", "false");

        const checkboxSobre = document.getElementById("chk-sobreescribir");
        checkboxSobre.onchange = function(e) {
            sessionStorage.setItem("edarlab_sobreescribir", e.target.checked ? "true" : "false");
        };
        document.getElementById("btn-ejecutar-ui").onclick = iniciarSecuenciaPersistente;
    }

    function renderizarFilasTabla(listaEdars) {
        const tbody = document.querySelector("#tabla-edars-body tbody");
        if (!tbody) return;
        tbody.innerHTML = "";
        listaEdars.forEach((cand, index) => {
            const fila = document.createElement("tr");
            // Corregimos los textos y aseguramos la estructura estricta de 4 columnas (td)
            fila.innerHTML = `
                <td style="text-align:center;"><input type="checkbox" class="edar-cb edar-fila-cb" data-index="${index}" checked></td>
                <td style="font-weight: 600;">${cand.nameExcel}</td>
                <td style="font-weight: 600; color: #475569;">${cand.nameWebAsociada}</td>
                <td style="text-align:center; font-weight: 700; color: ${CONFIG_PRIMARIO};"><span id="param-contador-${index}">${cand.dC}</span></td>
            `;
            const cb = fila.querySelector(".edar-fila-cb");
            if (cb) {
                cb.onchange = function() {
                    window.EDARLab_Buffer[index].procesar = cb.checked;
                    sessionStorage.setItem("edarlab_global_buffer_storage", JSON.stringify(window.EDARLab_Buffer));
                };
            }
            tbody.appendChild(fila);
        });
    }

    function iniciarSecuenciaPersistente() {
        activarBloqueoSuspension();
        let totalActivas = 0;
        let totalCeldasCola = 0;

        window.EDARLab_Buffer.forEach(c => {
            if (c.procesar === true) {
                totalActivas++;
                totalCeldasCola += c.dC; // Sumamos directamente el conteo de celdas original
            }
        });

        if (totalActivas === 0) {
            alert("Selecciona al menos una EDAR en la tabla.");
            return;
        }

        console.log(`[EDARLab➔EPSAR] Lanzando ejecución. Celdas totales a rellenar: ${totalCeldasCola}`);

        const nuevoEstado = {
            indiceBucleActual: 0,
            totalInicialActivas: totalActivas,
            contadorCompletadas: 0,
            totalCeldasGlobales: totalCeldasCola,
            celdasYaTipeadasAcumuladas: 0,
            pasoActual: "PREPARAR_FILTRADO",
            abortar: false
        };
        // Inicializamos el encabezado limpio del informe de texto plano
        const fechaHoy = new Date().toLocaleString();
        const headerTxt = `==================================================\nREPORTE DE VOLCADO EDARLab➔EPSAR\nFecha: ${fechaHoy}\n==================================================\n\n`;
        sessionStorage.setItem("edarlab_txt_log_accumulator", headerTxt);
        sessionStorage.setItem("edar_automator_state", JSON.stringify(nuevoEstado));
        reconstruirVentanaProgresoFijaCentrada(nuevoEstado);
    }

    function reconstruirVentanaProgresoFijaCentrada(estado) {
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
            zIndex: "100000", // Subimos el z-index al máximo para asegurar que esté por encima de todo
            fontFamily: CONFIG_FONT,
            border: "1px solid rgba(0,151,215,0.3)",

            // --- EL TRUCO DEL ESCUDO ANTI-CLICS ---
            // Esto crea una "sombra" gigante de 2000px a la redonda que es transparente
            // pero que CAPTURA y bloquea todos los clics externos del usuario.
            boxShadow: "0 0 0 2000px rgba(71, 85, 105, 0.4), 0 20px 50px rgba(0, 0, 0, 0.3)"
        });

        // Ventana de ejecución original intacta
        uiContenedor.innerHTML = `
        <div style="background:linear-gradient(135deg,${CONFIG_PRIMARIO},${CONFIG_SECUNDARIO}); padding:12px 16px; display:flex; justify-content:space-between; align-items:center; border-top-left-radius:16px; border-top-right-radius:16px;">
            <span style="color:white; font-size:12px; font-weight:700;">EDARLab➔EPSAR</span>
            <span id="close-ui-progreso" style="cursor:pointer; color:white; font-size:18px; font-weight:bold;">&times;</span>
        </div>
        <div id="edar-panel-body" style="padding:18px;">
            <div id="tx-progreso-global" style="font-size:13px; font-weight:700; color:#004381; margin-bottom:2px; text-align:center;"></div>
            <div id="tx-progreso-detalle" style="font-size:11px; color:${CONFIG_GRIS}; margin-bottom:12px; text-align:center;"></div>
            <div style="background:#e2e8f0; border-radius:4px; height:12px; overflow:hidden; box-shadow:inset 0 1px 2px rgba(0,0,0,0.1);">
                <div id="rb-barra" style="width:0%; height:100%; background:linear-gradient(90deg, ${CONFIG_PRIMARIO}, ${CONFIG_SECUNDARIO}); transition:width 0.3s ease;"></div>
            </div>
            <div id="tx-porcentaje" style="text-align:center; font-size:11px; margin-top:5px; font-weight:700; color:${CONFIG_GRIS};">0%</div>
        </div>
    `;

        document.body.appendChild(uiContenedor);
        document.getElementById("close-ui-progreso").onclick = abortarYLimpiarTodo;

        // --- ESCUDO DE EVENTOS JAVASCRIPT ---
        // Captura el clic en la fase inicial (true) y lo destruye por completo
        window.addEventListener('click', bloquearAccionUsuario, true);
        window.addEventListener('mousedown', bloquearAccionUsuario, true);
        window.addEventListener('keydown', bloquearAccionUsuario, true);
        window.addEventListener('keypress', bloquearAccionUsuario, true);
        ejecutarMaquinaEstados(estado);
    }

    // Nueva función auxiliar que destruye el clic (colócala justo debajo de la otra función)
    function bloquearAccionUsuario(e) {
        // 1. Si la acción está generada por el código del script (!e.isTrusted), la DEJAMOS PASAR
        if (!e.isTrusted) {
            return;
        }

        // 2. Si es una acción del usuario pero es en el botón de cerrar de nuestra UI, la DEJAMOS PASAR
        if (e.target.id === "close-ui-progreso" || e.target.innerText === "×") {
            return;
        }

        // 3. Si es CUALQUIER otra acción física tuya (clic, tabulador, letras, flechas...), la DESTRUIMOS
        e.stopPropagation();
        e.preventDefault();
        console.log(`[EDARLab] Acción física de ${e.type} bloqueada para proteger la tabla.`);
    }

    async function ejecutarMaquinaEstados(estado) {
        if (estado.abortar === true) {
            console.log("[EDARLab➔EPSAR] Bandera detectada. Abortando ejecución.");
            limpiarYResetearTodoAZero();
            return;
        }

        if (estado.indiceBucleActual >= window.EDARLab_Buffer.length) {
            console.log("[EDARLab➔EPSAR] Ejecución finalizada con éxito.");
            sessionStorage.removeItem("edar_automator_state");
            sonarAlertaFin();
            document.getElementById("rb-barra").style.width = "100%";
            document.getElementById("tx-porcentaje").innerText = `100% Celdas volcadas (${estado.totalCeldasGlobales} de ${estado.totalCeldasGlobales})`;
            document.getElementById("tx-progreso-global").innerText = "Ejecución finalizada con éxito";
            document.getElementById("tx-progreso-detalle").innerText = `Procesados un total de ${estado.totalInicialActivas} partes.`;
            document.getElementById("close-ui-progreso").onclick = limpiarYResetearTodoAZero;
            descargarReporteTxtFinal();
            liberarBloqueoSuspension();
            return;
        }

        const candActual = window.EDARLab_Buffer[estado.indiceBucleActual];

        if (candActual.procesar === false) {
            estado.indiceBucleActual++;
            estado.pasoActual = "PREPARAR_FILTRADO";
            sessionStorage.setItem("edar_automator_state", JSON.stringify(estado));
            ejecutarMaquinaEstados(estado);
            return;
        }

        const nombreWebActual = candActual.nameWebAsociada;
        const porcentajeAnalitico = Math.round((estado.celdasYaTipeadasAcumuladas / estado.totalCeldasGlobales) * 100);
        document.getElementById("rb-barra").style.width = `${porcentajeAnalitico}%`;

        const plantasPendientes = estado.totalInicialActivas - estado.contadorCompletadas;
        document.getElementById("tx-porcentaje").innerText = `${porcentajeAnalitico}% celdas (${estado.celdasYaTipeadasAcumuladas} de ${estado.totalCeldasGlobales})`;

        let desplegableWeb = document.getElementById("ctl00_ctl00_ContentPlaceHolder1_DropDownFiltroUnidadCoste");
        let botonMostrar = document.getElementById("ctl00_ctl00_ContentPlaceHolder1_ButtonFiltroMostrar");

        // --- ESTADO A: INTERACCIÓN Y FILTRADO DE PÁGINA ---
        if (estado.pasoActual === "PREPARAR_FILTRADO") {
            document.getElementById("tx-progreso-global").innerText = `Procesando ${estado.contadorCompletadas + 1} de ${estado.totalInicialActivas}`;
            document.getElementById("tx-progreso-detalle").innerText = `Introduciendo ${nombreWebActual} en el desplegable...`;

            if (desplegableWeb) {
                desplegableWeb.focus();
                desplegableWeb.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, view: window }));
                desplegableWeb.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, view: window }));
                desplegableWeb.click();
            }

            // Pausa dinámica: apertura desplegable
            await delay(CONFIG_PAUSAS.APERTURA_DESPLEGABLE);

            let estCheck = JSON.parse(sessionStorage.getItem("edar_automator_state") || "{}");
            if (estCheck.abortar === true || estCheck.indiceBucleActual === undefined) { limpiarYResetearTodoAZero(); return; }

            desplegableWeb = document.getElementById("ctl00_ctl00_ContentPlaceHolder1_DropDownFiltroUnidadCoste");
            if (desplegableWeb) {
                let indexObjetivo = -1;
                for (let i = 0; i < desplegableWeb.options.length; i++) {
                    if (desplegableWeb.options[i].text.trim() === nombreWebActual) {
                        indexObjetivo = i;
                        break;
                    }
                }

                if (indexObjetivo !== -1) {
                    desplegableWeb.selectedIndex = indexObjetivo;
                    desplegableWeb.value = desplegableWeb.options[indexObjetivo].value;
                    desplegableWeb.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: nombreWebActual }));
                    desplegableWeb.dispatchEvent(new Event("change", { bubbles: true }));
                }
            }

            // Pausa dinámica: asimilación pasiva del desplegable
            await delay(CONFIG_PAUSAS.ASIMILACION_DESPLEGABLE);

            estCheck = JSON.parse(sessionStorage.getItem("edar_automator_state") || "{}");
            if (estCheck.abortar === true || estCheck.indiceBucleActual === undefined) { limpiarYResetearTodoAZero(); return; }

            botonMostrar = document.getElementById("ctl00_ctl00_ContentPlaceHolder1_ButtonFiltroMostrar");
            document.getElementById("tx-progreso-global").innerText = `Procesando ${estado.contadorCompletadas + 1} de ${estado.totalInicialActivas}`;
            document.getElementById("tx-progreso-detalle").innerText = "Ejecutando carga de la tabla...";

            if (botonMostrar) {
                estado.pasoActual = "ESCRIBIR_TABLA";
                sessionStorage.setItem("edar_automator_state", JSON.stringify(estado));

                botonMostrar.focus();
                botonMostrar.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
                await delay(100 + Math.random() * 100);
                botonMostrar.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
                botonMostrar.click();
                return;
            }
            estado.pasoActual = "ESCRIBIR_TABLA";
        }

        // --- ESTADO B: VOLCADO DE DATOS REAL ---
        if (estado.pasoActual === "ESCRIBIR_TABLA") {
            // Pausa dinámica: espera del dibujo definitivo de la tabla
            document.getElementById("tx-progreso-global").innerText = `Procesando ${estado.contadorCompletadas + 1} de ${estado.totalInicialActivas}`;
            document.getElementById("tx-progreso-detalle").innerText = "Esperando carga de la tabla..."
            await delay(CONFIG_PAUSAS.RENDERIZADO_TABLA);

            // === NUEVO: FILTRO DE SEGURIDAD ANTICAÍDA ===
            // Buscamos el elemento contenedor de la tabla de la EPSAR (usa su ID real, por ejemplo 'tabla_parametros' o el del contenedor principal)
            let contenedorTabla = document.getElementById("ctl00_ctl00_ContentPlaceHolder1_Contenido_tablaResto");

            if (!contenedorTabla) {
                // Si la tabla no existe en el HTML, disparamos el protocolo de aborto seguro
                abortarPorTablaInexistente();
                return; // Detiene por completo la ejecución de la máquina de estados
            }

            let estCheck = JSON.parse(sessionStorage.getItem("edar_automator_state") || "{}");
            if (estCheck.abortar === true || estCheck.indiceBucleActual === undefined) { limpiarYResetearTodoAZero(); return; }

            let celdasEstaPlanta = candActual.dC;

            document.getElementById("tx-progreso-global").innerText = `Procesando ${estado.contadorCompletadas + 1} de ${estado.totalInicialActivas}`;
            document.getElementById("tx-progreso-detalle").innerText = `Escribiendo ${celdasEstaPlanta} celdas en la tabla...`;
            let celdasEscritasEnEstaEdar = 0;
            for (let day = 1; day <= 31; day++) {
                const dataDia = candActual.sD[day];
                if (dataDia) {
                    for (const k of ordenEscritura) {
                        let estCheckLoop = JSON.parse(sessionStorage.getItem("edar_automator_state") || "{}");
                        if (estCheckLoop.abortar === true || estCheckLoop.indiceBucleActual === undefined) { limpiarYResetearTodoAZero(); return; }

                        const v = dataDia[k];
                        if (v !== undefined) {
                            const elCell = document.getElementById(`ctl00_ctl00_ContentPlaceHolder1_Contenido_CELDA_MA_DIA_${day}_COLUMNA_${k}_texto`);
                            if (elCell) {
                                const modoSobreescribir = sessionStorage.getItem("edarlab_sobreescribir") === "true";

                                // Si la celda ya tiene un valor y NO queremos sobreescribir, saltamos a la siguiente
                                if (elCell.value.trim() !== "" && !modoSobreescribir) {
                                    console.log(`[EDARLab] Celda día ${day} - Columna ${k} ya tiene datos. Saltando celda.`);
                                    await delay(CONFIG_PAUSAS.POST_CELDA());
                                    estado.celdasYaTipeadasAcumuladas++;
                                    const subPorcentaje = Math.round((estado.celdasYaTipeadasAcumuladas / estado.totalCeldasGlobales) * 100);
                                    document.getElementById("rb-barra").style.width = `${subPorcentaje}%`;
                                    document.getElementById("tx-porcentaje").innerText = `${subPorcentaje}% celdas (${estado.celdasYaTipeadasAcumuladas} de ${estado.totalCeldasGlobales})`;
                                    continue;
                                }

                                // Si está vacía o el modo sobreescribir es true, escribe el dato de manera normal:
                                elCell.focus();
                                await typeH(elCell, v.toString());
                                elCell.setAttribute("data-pasted", "true");

                                celdasEscritasEnEstaEdar++;
                                estado.celdasYaTipeadasAcumuladas++;
                                const subPorcentaje = Math.round((estado.celdasYaTipeadasAcumuladas / estado.totalCeldasGlobales) * 100);
                                document.getElementById("rb-barra").style.width = `${subPorcentaje}%`;
                                document.getElementById("tx-porcentaje").innerText = `${subPorcentaje}% celdas (${estado.celdasYaTipeadasAcumuladas} de ${estado.totalCeldasGlobales})`;
                            }
                        }
                    }
                    // Pausa de final de fila diaria introducida
                    await delay(CONFIG_PAUSAS.POST_DIA_FILA);
                }
            }

            // Pausa dinámica: pre-recalcula
            document.getElementById("tx-progreso-global").innerText = `Procesando ${estado.contadorCompletadas + 1} de ${estado.totalInicialActivas}`;
            document.getElementById("tx-progreso-detalle").innerText = "Asentando datos en la tabla...";
            await delay(CONFIG_PAUSAS.PRE_RECALCULA);

            estCheck = JSON.parse(sessionStorage.getItem("edar_automator_state") || "{}");
            if (estCheck.abortar === true || estCheck.indiceBucleActual === undefined) { limpiarYResetearTodoAZero(); return; }

            // let botonRecalcula = document.getElementById("ctl00_ctl00_ContentPlaceHolder1_Contenido_BRecalcula")
            let botonRecalcula = document.getElementById("ctl00_ctl00_ContentPlaceHolder1_ButtonGuardar");
            document.getElementById("tx-progreso-detalle").innerText = "Guardando los parámetros...";

            if (botonRecalcula) {
                estado.pasoActual = "POST_RECALCULA_ESPERA";
                estado.celdasRealesUltimaEdar = celdasEscritasEnEstaEdar;
                sessionStorage.setItem("edar_automator_state", JSON.stringify(estado));

                botonRecalcula.focus();
                botonRecalcula.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
                await delay(100 + Math.random() * 100); // El ratón se queda hundido unos milisegundos reales
                botonRecalcula.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
                botonRecalcula.click();
                return;
            }
            estado.pasoActual = "POST_RECALCULA_ESPERA";
        }

        // --- ESTADO C: POST-RECALCULA Y TRANSCISIÓN A SIGUIENTE PLANTA ---
        if (estado.pasoActual === "POST_RECALCULA_ESPERA") {
            // Pausa dinámica: post-recalcula

            document.getElementById("tx-progreso-global").innerText = `Procesando ${estado.contadorCompletadas + 1} de ${estado.totalInicialActivas}`;
            document.getElementById("tx-progreso-detalle").innerText = "Asentando datos guardados en el servidor...";
            await delay(CONFIG_PAUSAS.POST_RECALCULA);

            // === NUEVO: FILTRO DE SEGURIDAD POST-RECALCULA ===
            // Comprobamos si la tabla sigue existiendo tras el recálculo. Si el servidor ha fallado, no estará.
            let contenedorTablaPost = document.getElementById("ctl00_ctl00_ContentPlaceHolder1_Contenido_tablaResto");
            if (!contenedorTablaPost) {
                // Si el recálculo rompió la página, abortamos de inmediato y salvamos lo que llevamos
                abortarPorTablaInexistente();
                return; // Detiene por completo la máquina de estados
            }

            const candActualPost = window.EDARLab_Buffer[estado.indiceBucleActual];
            const nombreWebActualPost = candActualPost.nameWebAsociada;

            sonarAlertaPlanta();
            let estCheck = JSON.parse(sessionStorage.getItem("edar_automator_state") || "{}");
            if (estCheck.abortar === true || estCheck.indiceBucleActual === undefined) { limpiarYResetearTodoAZero(); return; }

            // Capturamos el bloc de notas de la sesión y añadimos la línea de auditoría
            let currentLog = sessionStorage.getItem("edarlab_txt_log_accumulator") || "";
            const fechaPlanta = new Date().toLocaleTimeString();
            const celdasReales = estado.celdasRealesUltimaEdar || 0;
            // USAMOS LAS NUEVAS VARIABLES SEGURAS (candActualPost y nombreWebActualPost)
            currentLog += `[${fechaPlanta}] Completado: EDAR ${candActualPost.nameExcel} (Asociada Web: ${nombreWebActualPost}) | Celdas rellenadas: ${celdasReales}\n`;
            currentLog += `--------------------------------------------------\n\n`;
            sessionStorage.setItem("edarlab_txt_log_accumulator", currentLog);

            // Pausa dinámica: cambio estético de ciclo
            document.getElementById("tx-progreso-detalle").innerText = "Parte procesado con éxito...";
            await delay(CONFIG_PAUSAS.TRANSICION_PLANTA);

            estado.indiceBucleActual++;
            estado.contadorCompletadas++;
            estado.pasoActual = "PREPARAR_FILTRADO";
            sessionStorage.setItem("edar_automator_state", JSON.stringify(estado));

            ejecutarMaquinaEstados(estado);
        }
    }

    // --- FUNCIÓN NATIVA DE COMPILACIÓN Y DESCARGA DEL TXT ---
    function descargarReporteTxtFinal() {
        let textoFinal = sessionStorage.getItem("edarlab_txt_log_accumulator");
        // Si no hay datos registrados, evitamos descargar un archivo vacío
        if (!textoFinal || textoFinal.trim().length < 150) return;
        textoFinal += `\n==================================================\nFIN DEL INFORME \n==================================================`;
        const blob = new Blob([textoFinal], { type: "text/plain;charset=utf-8" });
        const urlDescarga = URL.createObjectURL(blob);
        const fecha = new Date().toISOString().slice(0, 10);
        const horaClean = new Date().toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' }).replace(':', '-');
        const enlacelink = document.createElement("a");
        enlacelink.href = urlDescarga;
        enlacelink.download = `Reporte_Analiticas_EPSAR_${fecha}_${horaClean}.txt`;
        document.body.appendChild(enlacelink);
        enlacelink.click();
        document.body.removeChild(enlacelink);
        URL.revokeObjectURL(urlDescarga);
        console.log("[EDARLab➔EPSAR] Reporte descargado con éxito.");
    }

    function abortarYLimpiarTodo() {
        descargarReporteTxtFinal();
        const estadoGuardado = sessionStorage.getItem("edar_automator_state");
        if (estadoGuardado) {
            const estado = JSON.parse(estadoGuardado);
            if (estado.indiceBucleActual < window.EDARLab_Buffer.length) {
                estado.abortar = true;
                sessionStorage.setItem("edar_automator_state", JSON.stringify(estado));
            }
        }
        limpiarYResetearTodoAZero();
    }

    // --- FUNCIÓN DE BORRADO ABSOLUTO A CERO CORREGIDA ---
    function limpiarYResetearTodoAZero() {
        console.log("[EDARLab➔EPSAR] Reseteando todas las variables a 0 y limpiando memoria persistente.");
        window.removeEventListener('click', bloquearAccionUsuario, true);
        window.removeEventListener('mousedown', bloquearAccionUsuario, true);
        window.removeEventListener('keydown', bloquearAccionUsuario, true);
        window.removeEventListener('keypress', bloquearAccionUsuario, true);
        sessionStorage.removeItem("edar_automator_state");
        sessionStorage.removeItem("edarlab_global_buffer_storage");
        sessionStorage.removeItem("edarlab_txt_log_accumulator");

        window.EDARLab_Buffer = [];

        if (uiContenedor) {
            uiContenedor.remove();
            uiContenedor = null;
        }

        document.onpaste = null;
        escuchandoHerramienta = false;

        // CORRECCIÓN BLINDADA: Busca el botón en el DOM actual; si no existe por la recarga, lo vuelve a inyectar limpio
        let botonExistente = document.getElementById("edar-launcher-btn");
        if (botonExistente) {
            botonExistente.innerText = "Analíticas - EDARLab➔EPSAR";
            botonExistente.style.background = "linear-gradient(135deg, #004381, #0097D7)";
            botonExistente.style.display = "block";
            botonExistente.ondblclick = null;
        } else {
            crearBotonLanzador();
        }
        liberarBloqueoSuspension();
        console.log("[EDARLab➔EPSAR] Restituido el estado de origen. Listo para una nueva carga limpia.");
    }

    function dormirUI() {
        limpiarYResetearTodoAZero();
    }

    if (document.readyState === "complete" || document.readyState === "interactive") {
        init();
    } else {
        document.addEventListener("DOMContentLoaded", init);
    }

// --- FUNCIÓN DE ALERTA SONORA (ESTILO NOTIFICACIÓN) ---
    // Genera un tono de campana armónico con desvanecimiento (Fade-out)
// --- FUNCIÓN DE ALERTA SONORA (TRIPLE TONO DE ALTA VISIBILIDAD) ---
    function sonarAlertaFin() {
        try {
            const audioCtx = new (window.AudioContext || window.webkitAudioContext)();

            // 1. PRIMER TONO (Grave - Base) - Nota Do5 (523.25 Hz)
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

            // 2. SEGUNDO TONO (Medio - Transición) - Nota Mi5 (659.25 Hz)
            // Entra 120 milisegundos después
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

            // 3. TERCER TONO (Agudo - El "Brillo" final) - Nota Sol5 (783.99 Hz)
            // Entra 240 milisegundos después del primero
            setTimeout(() => {
                const ahora3 = audioCtx.currentTime;
                let osc3 = audioCtx.createOscillator();
                let gain3 = audioCtx.createGain();
                osc3.type = 'triangle';
                osc3.frequency.setValueAtTime(783.99, ahora3);
                gain3.gain.setValueAtTime(0.30, ahora3);
                gain3.gain.exponentialRampToValueAtTime(0.001, ahora3 + 1.0); // Eco más largo al final (1 segundo)
                osc3.connect(gain3);
                gain3.connect(audioCtx.destination);
                osc3.start(ahora3);
                osc3.stop(ahora3 + 1.0);
            }, 240);

        } catch (e) {
            console.log("[EDARLab] No se pudo reproducir la notificación de triple tono: ", e);
        }
    }

    // --- FUNCIÓN DE ALERTA SONORA (1 TONO CORTO - PROGRESO DE PLANTA) ---
    // --- FUNCIÓN DE ALERTA SONORA (1 TONO - VOLUMEN REFERENCIA DE FINALIZAR) ---
    function sonarAlertaPlanta() {
        try {
            const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
            const ahora = audioCtx.currentTime;

            let osc = audioCtx.createOscillator();
            let gain = audioCtx.createGain();
            osc.type = 'triangle';
            osc.frequency.setValueAtTime(659.25, ahora); // Nota Mi5 (la misma base que la campana final)

            // AJUSTE: Subido a 0.35 (igual que el tono final) y alargado a 0.6s para un mejor eco
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

    function abortarPorTablaInexistente() {
        console.error(`[EDARLab➔EPSAR] CRÍTICO: No se encontró la tabla. Abortando.`);

        // 1. Descargamos el reporte parcial automáticamente para salvar el trabajo hecho
        try {
            descargarReporteTxtFinal();
        } catch (err) {
            console.error("[EDARLab] Error al descargar el reporte en el aborto: ", err);
        }

        // 2. Limpiamos el sessionStorage para romper el bucle definitivamente
        sessionStorage.removeItem("edar_automator_state");

        // 3. Borramos la interfaz de progreso y liberamos ratón/teclado por completo
        if (typeof limpiarYResetearTodoAZero === "function") {
            limpiarYResetearTodoAZero();
        } else {
            window.removeEventListener('click', bloquearAccionUsuario, true);
            window.removeEventListener('mousedown', bloquearAccionUsuario, true);
            window.removeEventListener('keydown', bloquearAccionUsuario, true);
            window.removeEventListener('keypress', bloquearAccionUsuario, true);
        }
        liberarBloqueoSuspension();
        // 4. Lanzamos la alerta nativa idéntica a la de tu formato incorrecto
        alert(`ERROR: La tabla no se ha cargado correctamente. Revisa el informe de ejecución.`);
    }

    // --- FUNCIONES PARA IMPEDIR QUE EL ORDENADOR ENTRE EN SUSPENSIÓN ---
    async function activarBloqueoSuspension() {
        try {
            if ('wakeLock' in navigator) {
                wakeLock = await navigator.wakeLock.request('screen');
                console.log("[EDARLab➔WakeLock] ☀️ Sistema anticaída activado. La pantalla NO se dormirá.");

                // Si el usuario minimiza y maximiza, volvemos a solicitarlo automáticamente
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
                console.log("[EDARLab➔WakeLock] 🌙 Bloqueo liberado. El ordenador ya puede suspenderse.");
            });
        }
    }
})();
