// --- CONFIGURACIÓN ---
const COL_MONTO = 'MONTO';
const COL_RESTO = 'RESTO';
const COL_PROVEEDOR = 'PROVEEDOR';
const COL_MES = 'MES';
const COL_ANO = 'AÑO';
const COL_FECHA = 'FECHA';
const COL_ESTADO = 'ESTADO';
const COL_CONCEPTO = 'CONCEPTO';

let dataGlobal = { mercaderia: [], servicio: [], ingresos: [] };

document.getElementById('excel-upload').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array', cellDates: true});
            
            // Debug: Ver qué hojas encontró
            console.log("Hojas encontradas:", workbook.SheetNames);
            
            procesarDatos(workbook);
        } catch (err) {
            console.error("Error crítico:", err);
            alert("Error al leer el archivo. Revisa la consola (F12).");
        }
    };
    reader.readAsArrayBuffer(file);
});

function procesarDatos(workbook) {
    // Función para buscar hojas de forma flexible
    const encontrarHoja = (nombres) => {
        const nombre = workbook.SheetNames.find(n => nombres.some(palabra => n.toLowerCase().includes(palabra.toLowerCase())));
        if (!nombre) console.warn("No se encontró hoja que contenga:", nombres);
        return nombre ? XLSX.utils.sheet_to_json(workbook.Sheets[nombre]) : [];
    };

    dataGlobal.mercaderia = encontrarHoja(['salidas mercaderia', 'mercaderia']);
    dataGlobal.servicio = encontrarHoja(['salidas servicio', 'servicio']);
    dataGlobal.ingresos = encontrarHoja(['ingresos']);

    console.log("Datos cargados - Mercadería:", dataGlobal.mercaderia.length, "Servicios:", dataGlobal.servicio.length, "Ingresos:", dataGlobal.ingresos.length);

    if (dataGlobal.mercaderia.length === 0 && dataGlobal.servicio.length === 0) {
        alert("¡Cuidado! No encontré datos en las hojas de salidas. Verifica que los nombres de las hojas contengan 'mercaderia' y 'servicio'.");
    }

    llenarSelectoresFiltros();
    actualizarDashboard();
    document.getElementById('dashboard').style.display = 'block';
}

function filtrarDatos() {
    const selAno = document.getElementById('filter-year').value;
    const selMes = document.getElementById('filter-month').value;
    const selProv = document.getElementById('filter-provider').value;

    const filtrar = (row) => {
        const rAno = row[COL_ANO] ? String(row[COL_ANO]).trim() : '';
        const rMes = row[COL_MES] ? String(row[COL_MES]).toUpperCase().trim() : '';
        const rProv = row[COL_PROVEEDOR] ? String(row[COL_PROVEEDOR]).toUpperCase().trim() : '';
        
        if (selAno !== 'ALL' && rAno !== selAno) return false;
        if (selMes !== 'ALL' && rMes !== selMes) return false;
        if (selProv !== 'ALL' && rProv !== selProv) return false;
        return true;
    };

    return {
        ingresos: dataGlobal.ingresos.filter(r => (selAno === 'ALL' || String(r[COL_ANO]) === selAno) && (selMes === 'ALL' || String(r[COL_MES]).toUpperCase().trim() === selMes)),
        mercaderia: dataGlobal.mercaderia.filter(filtrar),
        servicio: dataGlobal.servicio.filter(filtrar)
    };
}

function actualizarDashboard() {
    const filtrados = filtrarDatos();
    const todasSalidas = [...filtrados.mercaderia, ...filtrados.servicio];

    // Cálculo seguro de totales
    const totalIng = filtrados.ingresos.reduce((acc, row) => acc + (parseFloat(row[COL_MONTO]) || 0), 0);
    const totalSal = todasSalidas.reduce((acc, row) => acc + (parseFloat(row[COL_MONTO]) || 0), 0);
    const totalPendiente = todasSalidas.reduce((acc, row) => acc + (parseFloat(row[COL_RESTO]) || 0), 0);

    document.getElementById('kpi-ingresos').textContent = formatearPlata(totalIng);
    document.getElementById('kpi-ytd').textContent = formatearPlata(totalIng - totalSal);
    document.getElementById('kpi-pendientes').textContent = formatearPlata(totalPendiente);

    // Funciones de UI
    llenarTabla('table-mercaderia', filtrados.mercaderia);
    llenarTabla('table-servicio', filtrados.servicio);
    generarResumenEstados(todasSalidas);
    generarGraficoProveedores(todasSalidas);
}

function llenarTabla(id, datos) {
    const tbody = document.querySelector(`#${id} tbody`);
    if(!tbody) return;
    tbody.innerHTML = datos.map(row => `
        <tr>
            <td>${row[COL_FECHA] || '-'}</td>
            <td>${row[COL_PROVEEDOR] || '-'}</td>
            <td>${row[COL_CONCEPTO] || '-'}</td>
            <td>${row[COL_ESTADO] || '-'}</td>
            <td>${formatearPlata(row[COL_MONTO])}</td>
            <td style="color: #B28B84;">${row[COL_ESTADO] === 'EN PROCESO' ? formatearPlata(row[COL_RESTO]) : '-'}</td>
        </tr>
    `).join('');
}

// Helper funciones
function formatearPlata(n) { return `$${Math.round(n || 0).toLocaleString('es-AR')}`; }

function llenarSelectoresFiltros() {
    // (Código anterior para llenar los selects)
    let anos = new Set(), meses = new Set(), proveedores = new Set();
    [...dataGlobal.mercaderia, ...dataGlobal.servicio].forEach(r => {
        if(r[COL_ANO]) anos.add(String(r[COL_ANO]));
        if(r[COL_MES]) meses.add(String(r[COL_MES]).toUpperCase().trim());
        if(r[COL_PROVEEDOR]) proveedores.add(String(r[COL_PROVEEDOR]).toUpperCase().trim());
    });
    
    const fill = (id, vals) => {
        const s = document.getElementById(id);
        s.innerHTML = '<option value="ALL">Todos</option>' + [...vals].sort().map(v => `<option value="${v}">${v}</option>`).join('');
    };
    fill('filter-year', anos);
    fill('filter-month', meses);
    fill('filter-provider', proveedores);
}

function generarResumenEstados(salidas) {
    const sumas = { "OK": 0, "PENDIENTE": 0, "EN PROCESO": 0 };
    salidas.forEach(r => {
        const est = r[COL_ESTADO] ? String(r[COL_ESTADO]).toUpperCase().trim() : 'OTRO';
        if (sumas.hasOwnProperty(est)) sumas[est] += parseFloat(r[COL_MONTO] || 0);
    });
    const grid = document.getElementById('status-grid');
    if(grid) grid.innerHTML = Object.entries(sumas).map(([k, v]) => `<div class="status-item"><span>${k}</span><strong>${formatearPlata(v)}</strong></div>`).join('');
}

function generarGraficoProveedores(salidas) {
    const provMap = {};
    salidas.forEach(r => {
        const p = r[COL_PROVEEDOR] || 'Otros';
        provMap[p] = (provMap[p] || 0) + (parseFloat(r[COL_MONTO]) || 0);
    });
    const ordenados = Object.entries(provMap).sort((a,b) => b[1]-a[1]).slice(0, 8);
    // (Logica de Chart.js igual a la anterior...)
    console.log("Datos para gráfico proveedores:", ordenados);
}
