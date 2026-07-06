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

let dataGlobal = { mercaderia: [], servicio: [], ingresos: [], ventas: [] };
let chartEvolutivo = null;
let chartProveedores = null;

// Rubro actualmente seleccionado en la solapa de Ventas (para el detalle de subrubros)
let rubroSeleccionado = null;
// Cache del set de ventas filtrado actualmente activo, para poder recalcular subrubros al clickear
let ventasActualCache = [];

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

    rubroSeleccionado = null;
    llenarSelectoresFiltros();
    actualizarDashboard(); 
    document.getElementById('dashboard').style.display = 'block';
}

// Navegación de Pestañas (Tabs) expuesta globalmente
window.cambiarPestana = function(idPestana) {
    document.querySelectorAll('.tab-content').forEach(tab => tab.style.display = 'none');
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    
    document.getElementById(idPestana).style.display = 'block';
    event.currentTarget.classList.add('active');
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

    document.getElementById('filter-year').addEventListener('change', actualizarDashboard);
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
    generarGraficoProveedores(salidasTotales);

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

    // KPIs nuevos: Cantidad de Ventas (tickets únicos) y Ticket Promedio
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

// --- TOP RUBROS (clickeable) + DETALLE DE SUBRUBROS ---
function generarTopRubros(datos) {
    ventasActualCache = datos;
    const rubroMap = {};

    datos.forEach(row => {
        const rubro = row[COL_V_RUBRO] ? String(row[COL_V_RUBRO]).toUpperCase().trim() : 'OTROS';
        const cant = parseFloat(row[COL_V_CANTIDAD]) || 0;
        const total = parseFloat(row[COL_V_TOTAL]) || 0;
        if (!rubroMap[rubro]) rubroMap[rubro] = { cant: 0, total: 0 };
        rubroMap[rubro].cant += cant;
        rubroMap[rubro].total += total;
    });

    const topRubros = Object.entries(rubroMap).sort((a, b) => b[1].cant - a[1].cant).slice(0, 3);

    const tbRubros = document.querySelector('#table-top-rubros tbody');
    if (tbRubros) {
        tbRubros.innerHTML = topRubros.map(([rubro, d]) => `
            <tr class="clickable-row ${rubro === rubroSeleccionado ? 'active-row' : ''}" onclick="seleccionarRubro('${rubro.replace(/'/g, "\\'")}')">
                <td>${rubro}</td>
                <td>${d.cant}</td>
                <td><strong>${formatearPlata(d.total)}</strong></td>
            </tr>
        `).join('') || '<tr><td colspan="3">Sin datos</td></tr>';
    }

    // Si no hay rubro seleccionado, o el que estaba ya no existe en los datos filtrados,
    // seleccionamos por defecto el rubro Nº1 del top.
    const rubrosDisponibles = Object.keys(rubroMap);
    if (!rubroSeleccionado || !rubrosDisponibles.includes(rubroSeleccionado)) {
        rubroSeleccionado = topRubros.length > 0 ? topRubros[0][0] : null;
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
        if (titulo) titulo.textContent = 'DETALLE SUBRUBROS';
        tbSub.innerHTML = '<tr><td colspan="3">Seleccioná un rubro para ver el detalle</td></tr>';
        return;
    }

    if (titulo) titulo.textContent = `DETALLE SUBRUBROS · ${rubroSeleccionado}`;

    const subMap = {};
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
        });

    const subOrdenados = Object.entries(subMap).sort((a, b) => b[1].total - a[1].total);

    tbSub.innerHTML = subOrdenados.map(([sub, d]) => `
        <tr>
            <td>${sub}</td>
            <td>${d.cant}</td>
            <td><strong>${formatearPlata(d.total)}</strong></td>
        </tr>
    `).join('') || '<tr><td colspan="3">Sin datos para este rubro</td></tr>';
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
        options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });
}

function generarGraficoProveedores(salidas) {
    const provMap = {};
    salidas.forEach(r => {
        const p = r[COL_PROVEEDOR] ? String(r[COL_PROVEEDOR]).trim().toUpperCase() : 'OTROS';
        provMap[p] = (provMap[p] || 0) + (parseFloat(r[COL_MONTO]) || 0);
    });
    const ordenados = Object.entries(provMap).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const ctx = document.getElementById('proveedoresChart');
    if(!ctx) return;
    if (chartProveedores) chartProveedores.destroy();
    chartProveedores = new Chart(ctx.getContext('2d'), {
        type: 'bar',
        data: { labels: ordenados.map(i => i[0]), datasets: [{ label: 'Total Gastado', data: ordenados.map(i => i[1]), backgroundColor: '#A39B8B' }]},
        options: { indexAxis: 'y', responsive: true, plugins: { legend: { display: false } } }
    });
}
