// --- CONFIGURACIÓN DE COLUMNAS (FINANCIERO) ---
const COL_MONTO = 'MONTO';
const COL_RESTO = 'RESTO'; 
const COL_PROVEEDOR = 'PROVEEDOR';
const COL_MES = 'MES';
const COL_ANO = 'AÑO';
const COL_FECHA = 'FECHA';
const COL_ESTADO = 'ESTADO';
const COL_CONCEPTO = 'CONCEPTO';

// --- CONFIGURACIÓN DE COLUMNAS (VENTAS) EXACTAS SEGÚN TU EXCEL ---
const COL_V_TOTAL = 'Total'; 
const COL_V_RUBRO = 'Rubro';
const COL_V_SUBRUBRO = 'Sub Rubro'; // Con espacio, tal cual tu Excel
const COL_V_PRODUCTO = 'Producto';
const COL_V_CANTIDAD = 'Cantidad';
const COL_V_MES = 'MES';
const COL_V_ANO = 'año'; // En minúscula, tal cual tu Excel
const COL_V_COMPROBANTE = 'Comprobante'; // Identificador único de venta/ticket

// --- HOJAS DE FLUJO DE CAJA A INCLUIR ---
// Solo 2025 y 2026 (la hoja "Flujo de Caja" sin año, de 2024, está incompleta y se deja afuera
// por pedido del cliente. El día que la completen, se puede sumar acá: 2024: 'Flujo de Caja')
const FLUJO_SHEETS = {
    2025: 'Flujo de Caja 2025',
    2026: 'Flujo de Caja 2026'
};

let dataGlobal = { mercaderia: [], servicio: [], ingresos: [], ventas: [], flujoCaja: {} };
let chartEvolutivo = null;
let chartRubroEvolutivo = null;
let chartFlujo = null;

// Rubro actualmente seleccionado en la solapa de Ventas (para el detalle de subrubros y el evolutivo)
let rubroSeleccionado = null;
// Cache del set de ventas filtrado actualmente activo, para poder recalcular subrubros al clickear
let ventasActualCache = [];
// Criterio de orden de las tablas de Rubros / Subrubros: 'cant' o 'total'
let ordenRubros = 'total';
let ordenSubrubros = 'total';

const MESES_ORDEN = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC'];

document.getElementById('excel-upload').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    document.getElementById('file-name').textContent = file.name;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array', cellDates: true});
            procesarDatos(workbook);
        } catch (error) {
            console.error("Error leyendo el Excel:", error);
            alert("Hubo un problema procesando el archivo. Revisá la consola (F12).");
        }
    };
    reader.readAsArrayBuffer(file);
});

function procesarDatos(workbook) {
    const MAPPING = {
        mercaderia: 'SALIDAS MERCADERIA',
        servicio: 'SALIDAS SERVICIO',
        ingresos: 'INGRESOS',
        ventas: 'CONSOLIDADO VENTAS'
    };

    Object.keys(MAPPING).forEach(key => {
        const nombreHoja = MAPPING[key];
        if (workbook.Sheets[nombreHoja]) {
            dataGlobal[key] = XLSX.utils.sheet_to_json(workbook.Sheets[nombreHoja]);
        } else {
            console.warn(`⚠️ Hoja '${nombreHoja}' no encontrada.`);
            dataGlobal[key] = []; 
        }
    });

    if (dataGlobal.mercaderia.length === 0 && dataGlobal.ventas.length === 0) {
        alert("Atención: No se cargaron datos principales. Asegurate de que los nombres de las hojas sean exactos en tu Excel.");
        return;
    }

    // --- Flujo de Caja (parser genérico, hoja por hoja) ---
    dataGlobal.flujoCaja = {};
    Object.keys(FLUJO_SHEETS).forEach(anio => {
        const nombreHoja = FLUJO_SHEETS[anio];
        if (workbook.Sheets[nombreHoja]) {
            const filas = XLSX.utils.sheet_to_json(workbook.Sheets[nombreHoja], { header: 1, defval: null });
            const parsed = parsearHojaFlujoCaja(filas);
            if (parsed) dataGlobal.flujoCaja[anio] = parsed;
        } else {
            console.warn(`⚠️ Hoja '${nombreHoja}' no encontrada (Flujo de Caja ${anio}).`);
        }
    });

    rubroSeleccionado = null;
    llenarSelectoresFiltros();
    llenarSelectorFlujoAnio();

    // IMPORTANTE: el dashboard se muestra ANTES de generar los gráficos. Chart.js mide
    // el tamaño real del contenedor en el momento en que se crea el gráfico; si el
    // contenedor todavía está en display:none, lo mide como 0x0 y el gráfico queda
    // con una altura chica que después no se corrige sola.
    document.getElementById('dashboard').style.display = 'block';

    actualizarDashboard();
    actualizarFlujoCaja();

    // Red de seguridad extra: si el usuario carga el Excel estando parado en otra
    // solapa (ej. "Flujo de Caja"), los gráficos de las solapas ocultas en ese momento
    // igual se miden mal. Se corrigen solos al clickear esa solapa (ver cambiarPestana),
    // pero forzamos un resize también acá por las dudas.
    requestAnimationFrame(refrescarTamanosCharts);
}

function refrescarTamanosCharts() {
    [chartEvolutivo, chartRubroEvolutivo, chartFlujo].forEach(c => {
        if (c) c.resize();
    });
}

// Navegación de Pestañas (Tabs) expuesta globalmente
window.cambiarPestana = function(idPestana) {
    document.querySelectorAll('.tab-content').forEach(tab => tab.style.display = 'none');
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    
    document.getElementById(idPestana).style.display = 'block';
    event.currentTarget.classList.add('active');

    // Al volverse visible la solapa, los gráficos que estaban ocultos (medidos en 0x0
    // al momento de crearse) se recalculan correctamente.
    requestAnimationFrame(refrescarTamanosCharts);
}

function llenarSelectoresFiltros() {
    let anos = new Set(), meses = new Set(), proveedores = new Set();
    const todasSalidas = [...dataGlobal.mercaderia, ...dataGlobal.servicio];
    const todoIngreso = [...dataGlobal.ingresos, ...dataGlobal.ventas];

    todoIngreso.forEach(row => {
        const valAno = row[COL_ANO] !== undefined ? row[COL_ANO] : row[COL_V_ANO];
        const valMes = row[COL_MES] !== undefined ? row[COL_MES] : row[COL_V_MES];
        
        if(valAno) anos.add(String(valAno));
        if(valMes) meses.add(String(valMes).toUpperCase().trim().substring(0,3));
    });

    todasSalidas.forEach(row => {
        if(row[COL_ANO]) anos.add(String(row[COL_ANO]));
        if(row[COL_MES]) meses.add(String(row[COL_MES]).toUpperCase().trim().substring(0,3));
        if(row[COL_PROVEEDOR]) proveedores.add(String(row[COL_PROVEEDOR]).toUpperCase().trim());
    });

    const llenarSelect = (id, setValores) => {
        const select = document.getElementById(id);
        if(!select) return;
        select.innerHTML = '<option value="ALL">Todos</option>';
        [...setValores].sort().forEach(val => {
            if(val && val !== 'UND') select.innerHTML += `<option value="${val}">${val}</option>`;
        });
    };

    llenarSelect('filter-year', anos);
    llenarSelect('filter-month', meses);
    llenarSelect('filter-provider', proveedores);

    document.getElementById('filter-year').addEventListener('change', () => { actualizarDashboard(); generarGraficoEvolutivoRubro(rubroSeleccionado); });
    document.getElementById('filter-month').addEventListener('change', actualizarDashboard);
    document.getElementById('filter-provider').addEventListener('change', actualizarDashboard);
}

function formatearPlata(numero) { 
    return `$${Math.round(numero || 0).toLocaleString('es-AR')}`; 
}

function formatearPorcentaje(numero) {
    if(!isFinite(numero)) return "-";
    const signo = numero > 0 ? "+" : "";
    return `${signo}${numero.toFixed(1)}%`;
}

// Para porcentajes de "peso" (participación sobre un total), sin signo +/-
function formatearPorcentajeSimple(numero) {
    if(!isFinite(numero)) return "-";
    return `${numero.toFixed(1)}%`;
}

function obtenerColorClase(porcentaje) {
    if(!isFinite(porcentaje) || porcentaje === 0) return "text-neutral";
    return porcentaje > 0 ? "text-up" : "text-down";
}

function actualizarDashboard() {
    const selAno = document.getElementById('filter-year').value;
    const selMes = document.getElementById('filter-month').value;
    const selProv = document.getElementById('filter-provider').value;

    const aplicarFiltroGeneral = (row, usarProv) => {
        const rAno = row[COL_ANO] ? String(row[COL_ANO]) : null;
        const rMes = row[COL_MES] ? String(row[COL_MES]).toUpperCase().trim().substring(0,3) : null;
        const rProv = row[COL_PROVEEDOR] ? String(row[COL_PROVEEDOR]).toUpperCase().trim() : null;
        
        if (selAno !== 'ALL' && rAno !== selAno) return false;
        if (selMes !== 'ALL' && rMes !== selMes) return false;
        if (usarProv && selProv !== 'ALL' && rProv !== selProv) return false;
        return true;
    };

    const ingresosList = dataGlobal.ingresos.filter(r => aplicarFiltroGeneral(r, false));
    const mercaderiaList = dataGlobal.mercaderia.filter(r => aplicarFiltroGeneral(r, true));
    const servicioList = dataGlobal.servicio.filter(r => aplicarFiltroGeneral(r, true));
    const salidasTotales = [...mercaderiaList, ...servicioList];

    // KPIs Tab Financiero
    const totalIng = ingresosList.reduce((acc, row) => acc + (parseFloat(row[COL_MONTO]) || 0), 0);
    const totalSal = salidasTotales.reduce((acc, row) => acc + (parseFloat(row[COL_MONTO]) || 0), 0);
    
    document.getElementById('kpi-ingresos').textContent = formatearPlata(totalIng);
    document.getElementById('kpi-ytd').textContent = formatearPlata(totalIng - totalSal);
    document.getElementById('kpi-pendientes').textContent = formatearPlata(salidasTotales.reduce((acc, row) => acc + (parseFloat(row[COL_RESTO]) || 0), 0));

    generarResumenEstados(salidasTotales);
    llenarTablaGastos('table-mercaderia', mercaderiaList);
    llenarTablaGastos('table-servicio', servicioList);
    generarGraficoEvolutivo(dataGlobal.ingresos, [...dataGlobal.mercaderia, ...dataGlobal.servicio]); 

    // La tabla de Top Proveedores siempre se calcula sobre TODAS las salidas filtradas
    // por año/mes (no por proveedor), para mantener el ranking completo como contexto.
    const salidasSinFiltroProv = [...dataGlobal.mercaderia, ...dataGlobal.servicio].filter(r => aplicarFiltroGeneral(r, false));
    generarTablaProveedores(salidasSinFiltroProv);

    // Tab Ventas
    actualizarDashboardVentas(selAno, selMes);
}

// --- LÓGICA DE VENTAS ---
function actualizarDashboardVentas(selAno, selMes) {
    const notice = document.getElementById('comparative-notice');
    
    let paramsMesActual = null;
    let paramsMA = null;
    let paramsMMAA = null;

    if (selAno !== 'ALL' && selMes !== 'ALL') {
        if(notice) notice.style.display = 'none';
        const idxMes = MESES_ORDEN.indexOf(selMes);
        const anoNum = parseInt(selAno);
        
        if(idxMes !== -1) {
            paramsMesActual = { mes: selMes, ano: selAno };
            paramsMA = { 
                mes: idxMes === 0 ? MESES_ORDEN[11] : MESES_ORDEN[idxMes - 1], 
                ano: idxMes === 0 ? String(anoNum - 1) : selAno 
            };
            paramsMMAA = { mes: selMes, ano: String(anoNum - 1) };
        }
    } else {
        if(notice) notice.style.display = 'block';
    }

    const ventasActual = filtrarVentas(dataGlobal.ventas, paramsMesActual || {ano: selAno, mes: selMes});
    const ventasMA = paramsMA ? filtrarVentas(dataGlobal.ventas, paramsMA) : [];
    const ventasMMAA = paramsMMAA ? filtrarVentas(dataGlobal.ventas, paramsMMAA) : [];

    // KPI Facturación
    const factTotal = sumarColumnaVentas(ventasActual, COL_V_TOTAL);
    const factMA = sumarColumnaVentas(ventasMA, COL_V_TOTAL);
    const factMMAA = sumarColumnaVentas(ventasMMAA, COL_V_TOTAL);

    document.getElementById('kpi-v-total').textContent = formatearPlata(factTotal);
    
    const varMA = paramsMA && factMA > 0 ? ((factTotal - factMA) / factMA) * 100 : null;
    const varMMAA = paramsMMAA && factMMAA > 0 ? ((factTotal - factMMAA) / factMMAA) * 100 : null;

    actualizarKpiComparativo('kpi-v-ma', varMA);
    actualizarKpiComparativo('kpi-v-mmaa', varMMAA);

    // KPIs: Cantidad de Ventas (tickets únicos) y Ticket Promedio
    const cantVentas = contarVentasUnicas(ventasActual);
    const ticketPromedio = cantVentas > 0 ? factTotal / cantVentas : 0;
    document.getElementById('kpi-v-cantidad').textContent = cantVentas.toLocaleString('es-AR');
    document.getElementById('kpi-v-ticket-promedio').textContent = formatearPlata(ticketPromedio);

    generarTopRubros(ventasActual);
    generarTablaRubros(ventasActual, ventasMA, ventasMMAA, paramsMA != null);
}

function filtrarVentas(datos, params) {
    return datos.filter(row => {
        const rAno = row[COL_V_ANO] ? String(row[COL_V_ANO]).trim() : null;
        const rMes = row[COL_V_MES] ? String(row[COL_V_MES]).toUpperCase().trim().substring(0,3) : null;
        
        if (params.ano !== 'ALL' && rAno !== params.ano) return false;
        if (params.mes !== 'ALL' && rMes !== params.mes) return false;
        return true;
    });
}

function sumarColumnaVentas(datos, columna) {
    return datos.reduce((acc, row) => acc + (parseFloat(row[columna]) || 0), 0);
}

// Cuenta ventas (tickets) únicas usando el Nº de Comprobante. Si una fila no tiene
// comprobante, cada una cuenta como una venta individual (fallback razonable).
function contarVentasUnicas(datos) {
    const comprobantes = new Set();
    let sinComprobante = 0;
    datos.forEach(row => {
        const comp = row[COL_V_COMPROBANTE];
        if (comp !== undefined && comp !== null && String(comp).trim() !== '') {
            comprobantes.add(String(comp).trim());
        } else {
            sinComprobante++;
        }
    });
    return comprobantes.size + sinComprobante;
}

function actualizarKpiComparativo(idElemento, porcentaje) {
    const el = document.getElementById(idElemento);
    if(!el) return;
    if (porcentaje === null) {
        el.textContent = "-";
        el.className = "comparative-text text-neutral";
    } else {
        el.textContent = formatearPorcentaje(porcentaje);
        el.className = `comparative-text ${obtenerColorClase(porcentaje)}`;
    }
}

// --- TOGGLES DE ORDEN (Rubros / Subrubros) ---
window.cambiarOrdenRubros = function(campo) {
    ordenRubros = campo;
    document.querySelectorAll('#rubros-sort-toggle .sort-btn').forEach(b => b.classList.toggle('active', b.dataset.sort === campo));
    generarTopRubros(ventasActualCache);
}

window.cambiarOrdenSubrubros = function(campo) {
    ordenSubrubros = campo;
    document.querySelectorAll('#subrubros-sort-toggle .sort-btn').forEach(b => b.classList.toggle('active', b.dataset.sort === campo));
    renderDetalleSubrubros();
}

// --- RUBROS (clickeables, listado completo) + DETALLE DE SUBRUBROS + EVOLUTIVO ---
function generarTopRubros(datos) {
    ventasActualCache = datos;
    const rubroMap = {};
    let totalGeneral = 0;

    datos.forEach(row => {
        const rubro = row[COL_V_RUBRO] ? String(row[COL_V_RUBRO]).toUpperCase().trim() : 'OTROS';
        const cant = parseFloat(row[COL_V_CANTIDAD]) || 0;
        const total = parseFloat(row[COL_V_TOTAL]) || 0;
        if (!rubroMap[rubro]) rubroMap[rubro] = { cant: 0, total: 0 };
        rubroMap[rubro].cant += cant;
        rubroMap[rubro].total += total;
        totalGeneral += total;
    });

    const todosRubros = Object.entries(rubroMap).sort((a, b) => b[1][ordenRubros] - a[1][ordenRubros]);

    const tbRubros = document.querySelector('#table-top-rubros tbody');
    if (tbRubros) {
        tbRubros.innerHTML = todosRubros.map(([rubro, d]) => {
            const pct = totalGeneral > 0 ? (d.total / totalGeneral) * 100 : 0;
            return `
            <tr class="clickable-row ${rubro === rubroSeleccionado ? 'active-row' : ''}" onclick="seleccionarRubro('${rubro.replace(/'/g, "\\'")}')">
                <td>${rubro}</td>
                <td>${d.cant}</td>
                <td><strong>${formatearPlata(d.total)}</strong></td>
                <td>${formatearPorcentajeSimple(pct)}</td>
            </tr>
        `;
        }).join('') || '<tr><td colspan="4">Sin datos</td></tr>';
    }

    // Si no hay rubro seleccionado, o el que estaba ya no existe en los datos filtrados,
    // seleccionamos por defecto el rubro Nº1 según el orden activo.
    const rubrosDisponibles = Object.keys(rubroMap);
    if (!rubroSeleccionado || !rubrosDisponibles.includes(rubroSeleccionado)) {
        rubroSeleccionado = todosRubros.length > 0 ? todosRubros[0][0] : null;
    }

    renderDetalleSubrubros();
}

window.seleccionarRubro = function(rubro) {
    rubroSeleccionado = rubro;
    document.querySelectorAll('#table-top-rubros tbody tr').forEach(tr => tr.classList.remove('active-row'));
    if (window.event && window.event.currentTarget) {
        window.event.currentTarget.classList.add('active-row');
    }
    renderDetalleSubrubros();
}

function renderDetalleSubrubros() {
    const titulo = document.getElementById('subrubros-title');
    const tbSub = document.querySelector('#table-detalle-subrubros tbody');
    if (!tbSub) return;

    if (!rubroSeleccionado) {
        tbSub.innerHTML = '<tr><td colspan="4">Seleccioná un rubro para ver el detalle</td></tr>';
        generarGraficoEvolutivoRubro(null);
        return;
    }

    if (titulo) {
        // Actualizamos solo el texto del título, conservando el toggle de orden embebido
        const primerNodo = titulo.childNodes[0];
        if (primerNodo) primerNodo.textContent = `DETALLE SUBRUBROS · ${rubroSeleccionado} `;
    }

    const subMap = {};
    let totalRubro = 0;
    ventasActualCache
        .filter(row => {
            const rubro = row[COL_V_RUBRO] ? String(row[COL_V_RUBRO]).toUpperCase().trim() : 'OTROS';
            return rubro === rubroSeleccionado;
        })
        .forEach(row => {
            const sub = row[COL_V_SUBRUBRO] ? String(row[COL_V_SUBRUBRO]).toUpperCase().trim() : 'SIN SUBRUBRO';
            const cant = parseFloat(row[COL_V_CANTIDAD]) || 0;
            const total = parseFloat(row[COL_V_TOTAL]) || 0;
            if (!subMap[sub]) subMap[sub] = { cant: 0, total: 0 };
            subMap[sub].cant += cant;
            subMap[sub].total += total;
            totalRubro += total;
        });

    const subOrdenados = Object.entries(subMap).sort((a, b) => b[1][ordenSubrubros] - a[1][ordenSubrubros]);

    tbSub.innerHTML = subOrdenados.map(([sub, d]) => {
        const pct = totalRubro > 0 ? (d.total / totalRubro) * 100 : 0;
        return `
        <tr>
            <td>${sub}</td>
            <td>${d.cant}</td>
            <td><strong>${formatearPlata(d.total)}</strong></td>
            <td>${formatearPorcentajeSimple(pct)}</td>
        </tr>
    `;
    }).join('') || '<tr><td colspan="4">Sin datos para este rubro</td></tr>';

    generarGraficoEvolutivoRubro(rubroSeleccionado);
}

// Evolución mensual (facturación + cantidad) del rubro seleccionado. Se calcula sobre
// TODA la data de ventas (ignorando el filtro de Mes, ya que no tendría sentido un
// evolutivo de un solo mes) pero respetando el Año si el usuario eligió uno puntual.
function generarGraficoEvolutivoRubro(rubro) {
    const titulo = document.getElementById('rubro-evolutivo-title');
    const ctx = document.getElementById('rubroEvolutivoChart');
    if (!ctx) return;

    if (!rubro) {
        if (titulo) titulo.textContent = 'EVOLUTIVO MENSUAL POR RUBRO';
        if (chartRubroEvolutivo) { chartRubroEvolutivo.destroy(); chartRubroEvolutivo = null; }
        return;
    }

    const selAno = document.getElementById('filter-year').value;
    if (titulo) titulo.textContent = `EVOLUTIVO MENSUAL · ${rubro}`;

    const baseData = dataGlobal.ventas.filter(row => {
        const ano = row[COL_V_ANO] ? String(row[COL_V_ANO]).trim() : null;
        if (selAno !== 'ALL' && ano !== selAno) return false;
        const rRubro = row[COL_V_RUBRO] ? String(row[COL_V_RUBRO]).toUpperCase().trim() : 'OTROS';
        return rRubro === rubro;
    });

    const timelineMap = {};
    baseData.forEach(row => {
        const ano = row[COL_V_ANO] ? String(row[COL_V_ANO]).trim() : '????';
        const mes = row[COL_V_MES] ? String(row[COL_V_MES]).toUpperCase().trim().substring(0,3) : 'OTR';
        const key = `${ano}-${mes}`;
        if (!timelineMap[key]) timelineMap[key] = { ano, mes, cant: 0, total: 0 };
        timelineMap[key].cant += parseFloat(row[COL_V_CANTIDAD]) || 0;
        timelineMap[key].total += parseFloat(row[COL_V_TOTAL]) || 0;
    });

    const ordenados = Object.values(timelineMap)
        .filter(d => d.mes !== 'OTR')
        .sort((a, b) => {
            if (a.ano !== b.ano) return a.ano.localeCompare(b.ano);
            return MESES_ORDEN.indexOf(a.mes) - MESES_ORDEN.indexOf(b.mes);
        });

    const labels = ordenados.map(d => selAno !== 'ALL' ? d.mes : `${d.mes} ${d.ano}`);

    if (chartRubroEvolutivo) chartRubroEvolutivo.destroy();
    chartRubroEvolutivo = new Chart(ctx.getContext('2d'), {
        data: {
            labels,
            datasets: [
                { type: 'bar', label: 'Facturación', data: ordenados.map(d => d.total), backgroundColor: '#A39B8B', yAxisID: 'y' },
                { type: 'line', label: 'Cantidad Vendida', data: ordenados.map(d => d.cant), borderColor: '#7D8C7A', backgroundColor: 'transparent', yAxisID: 'y1', tension: 0.3 }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom' } },
            scales: {
                y: { beginAtZero: true, title: { display: true, text: 'Facturación' } },
                y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: 'Cantidad' } }
            }
        }
    });
}

function generarTablaRubros(datosActual, datosMA, datosMMAA, mostrarComparativas) {
    const rubrosSet = new Set();
    const dataAgrupada = {};

    const agrupar = (datos, keyObj) => {
        datos.forEach(row => {
            const rubro = row[COL_V_RUBRO] ? String(row[COL_V_RUBRO]).toUpperCase().trim() : 'OTROS';
            rubrosSet.add(rubro);
            if(!dataAgrupada[rubro]) dataAgrupada[rubro] = { actualTotal: 0, actualCant: 0, maTotal: 0, mmaaTotal: 0 };
            
            if(keyObj === 'actual') {
                dataAgrupada[rubro].actualTotal += parseFloat(row[COL_V_TOTAL]) || 0;
                dataAgrupada[rubro].actualCant += parseFloat(row[COL_V_CANTIDAD]) || 0;
            } else if (keyObj === 'ma') {
                dataAgrupada[rubro].maTotal += parseFloat(row[COL_V_TOTAL]) || 0;
            } else {
                dataAgrupada[rubro].mmaaTotal += parseFloat(row[COL_V_TOTAL]) || 0;
            }
        });
    };

    agrupar(datosActual, 'actual');
    agrupar(datosMA, 'ma');
    agrupar(datosMMAA, 'mmaa');

    const tbody = document.querySelector('#table-rubros tbody');
    if(!tbody) return;
    tbody.innerHTML = '';

    [...rubrosSet].sort().forEach(rubro => {
        const d = dataAgrupada[rubro];
        if (d.actualTotal === 0 && d.maTotal === 0 && d.mmaaTotal === 0) return;

        let htmlCrecMA = '-';
        let htmlCrecMMAA = '-';

        if (mostrarComparativas) {
            const varMA = d.maTotal ? ((d.actualTotal - d.maTotal) / d.maTotal) * 100 : null;
            const varMMAA = d.mmaaTotal ? ((d.actualTotal - d.mmaaTotal) / d.mmaaTotal) * 100 : null;
            
            htmlCrecMA = varMA !== null ? `<span class="${obtenerColorClase(varMA)}">${formatearPorcentaje(varMA)}</span>` : '-';
            htmlCrecMMAA = varMMAA !== null ? `<span class="${obtenerColorClase(varMMAA)}">${formatearPorcentaje(varMMAA)}</span>` : '-';
        }

        tbody.innerHTML += `
            <tr>
                <td><strong>${rubro}</strong></td>
                <td>${d.actualCant}</td>
                <td>${formatearPlata(d.actualTotal)}</td>
                <td>${htmlCrecMA}</td>
                <td>${htmlCrecMMAA}</td>
            </tr>
        `;
    });
}

// --- ORIGINALES (GASTOS Y GRÁFICOS) ---
function formatearFecha(fecha) {
    if(!fecha) return '-';
    if(fecha instanceof Date) return fecha.toLocaleDateString('es-AR');
    return String(fecha).substring(0, 10);
}

function generarResumenEstados(salidas) {
    const sumas = { "OK": 0, "PENDIENTE": 0, "RESERVADO": 0, "EN PROCESO": 0 };
    salidas.forEach(r => {
        const est = r[COL_ESTADO] ? String(r[COL_ESTADO]).toUpperCase().trim() : 'OTRO';
        if (sumas[est] !== undefined) sumas[est] += parseFloat(r[COL_MONTO]) || 0;
    });
    const grid = document.getElementById('status-grid');
    if(!grid) return;
    grid.innerHTML = Object.entries(sumas).map(([k, v]) => `
        <div class="status-item"><span>${k}</span><strong>${formatearPlata(v)}</strong></div>
    `).join('');
}

function llenarTablaGastos(idTabla, datos) {
    const tbody = document.querySelector(`#${idTabla} tbody`);
    if(!tbody) return;
    tbody.innerHTML = '';
    if(datos.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;">No hay datos</td></tr>`;
        return;
    }
    datos.forEach(row => {
        const est = row[COL_ESTADO] ? String(row[COL_ESTADO]).toUpperCase().trim() : '-';
        const resto = parseFloat(row[COL_RESTO]) || 0;
        const displayResto = (est === 'EN PROCESO' && resto > 0) ? `<span style="color:var(--salidas-color);">${formatearPlata(resto)}</span>` : '-';
        tbody.innerHTML += `<tr>
            <td>${formatearFecha(row[COL_FECHA])}</td>
            <td><strong>${row[COL_PROVEEDOR] || '-'}</strong></td>
            <td>${row[COL_CONCEPTO] || '-'}</td>
            <td>${est}</td>
            <td>${formatearPlata(parseFloat(row[COL_MONTO]) || 0)}</td>
            <td><strong>${displayResto}</strong></td>
        </tr>`;
    });
}

function generarGraficoEvolutivo(ingresos, salidas) {
    const mesesMap = {};
    ingresos.forEach(r => {
        let m = r[COL_MES] ? String(r[COL_MES]).toUpperCase().trim().substring(0,3) : 'OTR';
        if(!mesesMap[m]) mesesMap[m] = {ing:0, sal:0};
        mesesMap[m].ing += parseFloat(r[COL_MONTO]) || 0;
    });
    salidas.forEach(r => {
        let m = r[COL_MES] ? String(r[COL_MES]).toUpperCase().trim().substring(0,3) : 'OTR';
        if(!mesesMap[m]) mesesMap[m] = {ing:0, sal:0};
        mesesMap[m].sal += parseFloat(r[COL_MONTO]) || 0;
    });
    const labels = Object.keys(mesesMap).filter(m => m !== 'OTR' && m !== 'UND');
    const ctx = document.getElementById('evolutivoChart');
    if(!ctx) return;
    if (chartEvolutivo) chartEvolutivo.destroy();
    chartEvolutivo = new Chart(ctx.getContext('2d'), {
        type: 'bar',
        data: { labels, datasets: [
            { label: 'Ingresos', data: labels.map(m => mesesMap[m].ing), backgroundColor: '#7D8C7A' },
            { label: 'Salidas', data: labels.map(m => mesesMap[m].sal), backgroundColor: '#B28B84' }
        ]},
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
    });
}

// Tabla de ranking de proveedores: más accionable para decisiones que un gráfico de
// barras (permite ver de un vistazo quién concentra el gasto y a quién hay que pagarle).
// Siempre se calcula sobre TODOS los proveedores (año/mes aplicados, pero sin aplicar
// el filtro de Proveedor) para no perder el contexto global al filtrar.
function generarTablaProveedores(salidas) {
    const provMap = {};
    let totalGeneral = 0;

    salidas.forEach(r => {
        const p = r[COL_PROVEEDOR] ? String(r[COL_PROVEEDOR]).trim().toUpperCase() : 'OTROS';
        const monto = parseFloat(r[COL_MONTO]) || 0;
        const resto = parseFloat(r[COL_RESTO]) || 0;
        if (!provMap[p]) provMap[p] = { total: 0, pendiente: 0 };
        provMap[p].total += monto;
        provMap[p].pendiente += resto;
        totalGeneral += monto;
    });

    const ordenados = Object.entries(provMap).sort((a, b) => b[1].total - a[1].total);

    const tbody = document.querySelector('#table-proveedores tbody');
    if (!tbody) return;
    tbody.innerHTML = ordenados.map(([prov, d]) => {
        const pct = totalGeneral > 0 ? (d.total / totalGeneral) * 100 : 0;
        const pendienteHtml = d.pendiente > 0 ? `<span style="color:var(--salidas-color);">${formatearPlata(d.pendiente)}</span>` : '-';
        return `
            <tr>
                <td><strong>${prov}</strong></td>
                <td>${formatearPlata(d.total)}</td>
                <td>${formatearPorcentajeSimple(pct)}</td>
                <td>${pendienteHtml}</td>
            </tr>
        `;
    }).join('') || '<tr><td colspan="4">Sin datos</td></tr>';
}

// ==========================================================================
// FLUJO DE CAJA
// ==========================================================================

// Parser genérico: no asume una lista fija de conceptos (varía año a año en el Excel
// del cliente). Ubica la fila "MESES" para saber en qué columnas están ENE..DIC, y a
// partir de ahí recorre "DETALLE DE INGRESOS" / "DETALLE DE EGRESOS" tomando cualquier
// concepto que tenga datos, hasta llegar a "FLUJO DE CAJA ECONOMICO".
function parsearHojaFlujoCaja(filas) {
    let mesesRowIdx = -1;
    let labelCol = -1;

    for (let i = 0; i < filas.length; i++) {
        const fila = filas[i] || [];
        for (let c = 0; c < fila.length; c++) {
            if (fila[c] && String(fila[c]).toUpperCase().trim() === 'MESES') {
                mesesRowIdx = i;
                labelCol = c;
                break;
            }
        }
        if (mesesRowIdx !== -1) break;
    }
    if (mesesRowIdx === -1) return null;

    const mesesRow = filas[mesesRowIdx];
    const monthCols = [];
    for (let c = labelCol + 1; c < mesesRow.length; c++) {
        const val = mesesRow[c];
        if (!val) continue;
        const codigo = String(val).toUpperCase().trim().substring(0, 3);
        const idx = MESES_ORDEN.indexOf(codigo);
        if (idx !== -1 && !monthCols.some(m => m.mesIdx === idx)) {
            monthCols.push({ col: c, mesIdx: idx });
        }
        if (monthCols.length === 12) break;
    }
    if (monthCols.length === 0) return null;

    let saldoInicial = 0;
    const ingresos = [];
    const egresos = [];
    let seccion = null;

    for (let i = mesesRowIdx + 1; i < filas.length; i++) {
        const fila = filas[i] || [];
        const label = fila[labelCol];
        if (label === null || label === undefined || String(label).trim() === '') continue;
        const upper = String(label).toUpperCase().trim();

        if (upper.includes('RESUMEN DE EFECTIVO')) { seccion = 'resumen'; continue; }
        if (upper === 'SALDO INICIAL') {
            const valores = monthCols.map(m => parseFloat(fila[m.col]) || 0);
            saldoInicial = valores.find(v => v !== 0) || 0;
            continue;
        }
        if (upper.includes('DETALLE DE INGRESOS')) { seccion = 'ingresos'; continue; }
        if (upper.includes('DETALLE DE EGRESOS')) { seccion = 'egresos'; continue; }
        if (upper.startsWith('TOTAL INGRESOS')) { continue; }
        if (upper.startsWith('TOTAL EGRESOS')) { continue; }
        if (upper.includes('FLUJO DE CAJA ECONOMICO') || upper.includes('FLUJO DE CAJA NETO')) { break; }

        if (seccion === 'ingresos' || seccion === 'egresos') {
            const valores = monthCols.map(m => parseFloat(fila[m.col]) || 0);
            const total = valores.reduce((a, b) => a + b, 0);
            if (total === 0 && valores.every(v => v === 0)) continue; // concepto sin movimientos, se omite
            const item = { concepto: String(label).trim(), valores, total };
            if (seccion === 'ingresos') ingresos.push(item); else egresos.push(item);
        }
    }

    const totalIngresosMensual = MESES_ORDEN.map((_, idx) => ingresos.reduce((a, item) => a + item.valores[idx], 0));
    const totalEgresosMensual = MESES_ORDEN.map((_, idx) => egresos.reduce((a, item) => a + item.valores[idx], 0));
    const flujoNetoMensual = MESES_ORDEN.map((_, idx) => totalIngresosMensual[idx] - totalEgresosMensual[idx]);

    const saldoAcumulado = [];
    let acumulado = saldoInicial;
    flujoNetoMensual.forEach(f => { acumulado += f; saldoAcumulado.push(acumulado); });

    return {
        saldoInicial,
        ingresos,
        egresos,
        totalIngresosMensual,
        totalEgresosMensual,
        flujoNetoMensual,
        saldoAcumulado,
        totalIngresosAnual: totalIngresosMensual.reduce((a, b) => a + b, 0),
        totalEgresosAnual: totalEgresosMensual.reduce((a, b) => a + b, 0)
    };
}

function llenarSelectorFlujoAnio() {
    const select = document.getElementById('filter-flujo-year');
    if (!select) return;
    const anios = Object.keys(dataGlobal.flujoCaja).sort();
    select.innerHTML = anios.map(a => `<option value="${a}">${a}</option>`).join('') || '<option value="">Sin datos</option>';
    if (anios.length > 0) select.value = anios[anios.length - 1]; // año más reciente por defecto
    select.onchange = actualizarFlujoCaja;
}

function actualizarFlujoCaja() {
    const select = document.getElementById('filter-flujo-year');
    if (!select) return;
    const anio = select.value;
    const datos = dataGlobal.flujoCaja[anio];

    if (!datos) {
        ['kpi-fc-ingresos', 'kpi-fc-egresos', 'kpi-fc-neto', 'kpi-fc-saldo'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.textContent = '$0';
        });
        return;
    }

    document.getElementById('kpi-fc-ingresos').textContent = formatearPlata(datos.totalIngresosAnual);
    document.getElementById('kpi-fc-egresos').textContent = formatearPlata(datos.totalEgresosAnual);

    const neto = datos.totalIngresosAnual - datos.totalEgresosAnual;
    const kpiNeto = document.getElementById('kpi-fc-neto');
    kpiNeto.textContent = formatearPlata(neto);
    kpiNeto.className = obtenerColorClase(neto);

    document.getElementById('kpi-fc-saldo').textContent = formatearPlata(datos.saldoAcumulado[datos.saldoAcumulado.length - 1]);

    generarGraficoFlujo(datos);
    llenarTablaFlujoMensual(datos);
    llenarTablaFlujoDetalle('table-flujo-ingresos', datos.ingresos);
    llenarTablaFlujoDetalle('table-flujo-egresos', datos.egresos);
}

function generarGraficoFlujo(datos) {
    const ctx = document.getElementById('flujoChart');
    if (!ctx) return;
    if (chartFlujo) chartFlujo.destroy();
    chartFlujo = new Chart(ctx.getContext('2d'), {
        data: {
            labels: MESES_ORDEN,
            datasets: [
                { type: 'bar', label: 'Ingresos', data: datos.totalIngresosMensual, backgroundColor: '#7D8C7A', yAxisID: 'y' },
                { type: 'bar', label: 'Egresos', data: datos.totalEgresosMensual, backgroundColor: '#B28B84', yAxisID: 'y' },
                { type: 'line', label: 'Saldo Acumulado', data: datos.saldoAcumulado, borderColor: '#333333', backgroundColor: 'transparent', yAxisID: 'y1', tension: 0.3 }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            plugins: { legend: { position: 'bottom' } },
            scales: {
                y: { beginAtZero: true, title: { display: true, text: 'Ingresos / Egresos' } },
                y1: { position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'Saldo Acumulado' } }
            }
        }
    });
}

function llenarTablaFlujoMensual(datos) {
    const tbody = document.querySelector('#table-flujo-mensual tbody');
    if (!tbody) return;
    tbody.innerHTML = MESES_ORDEN.map((mes, idx) => `
        <tr>
            <td><strong>${mes}</strong></td>
            <td>${formatearPlata(datos.totalIngresosMensual[idx])}</td>
            <td>${formatearPlata(datos.totalEgresosMensual[idx])}</td>
            <td><span class="${obtenerColorClase(datos.flujoNetoMensual[idx])}">${formatearPlata(datos.flujoNetoMensual[idx])}</span></td>
            <td><strong>${formatearPlata(datos.saldoAcumulado[idx])}</strong></td>
        </tr>
    `).join('') + `
        <tr style="border-top: 2px solid rgba(0,0,0,0.15);">
            <td><strong>TOTAL</strong></td>
            <td><strong>${formatearPlata(datos.totalIngresosAnual)}</strong></td>
            <td><strong>${formatearPlata(datos.totalEgresosAnual)}</strong></td>
            <td><strong>${formatearPlata(datos.totalIngresosAnual - datos.totalEgresosAnual)}</strong></td>
            <td>-</td>
        </tr>
    `;
}

function llenarTablaFlujoDetalle(idTabla, items) {
    const tbody = document.querySelector(`#${idTabla} tbody`);
    if (!tbody) return;
    const ordenados = [...items].sort((a, b) => b.total - a.total);
    tbody.innerHTML = ordenados.map(item => `
        <tr><td>${item.concepto}</td><td><strong>${formatearPlata(item.total)}</strong></td></tr>
    `).join('') || '<tr><td colspan="2">Sin datos</td></tr>';
}
