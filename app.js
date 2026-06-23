// --- CONFIGURACIÓN DE COLUMNAS EXACTAS DE TU EXCEL ---
const COL_MONTO = 'MONTO';
const COL_RESTO = 'RESTO'; 
const COL_PROVEEDOR = 'PROVEEDOR';
const COL_MES = 'MES';
const COL_ANO = 'AÑO';
const COL_FECHA = 'FECHA';
const COL_ESTADO = 'ESTADO';
const COL_CONCEPTO = 'CONCEPTO';

let dataGlobal = { mercaderia: [], servicio: [], ingresos: [] };
let chartEvolutivo = null;
let chartProveedores = null;

document.getElementById('excel-upload').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    document.getElementById('file-name').textContent = file.name;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array', cellDates: true});
        procesarDatos(workbook);
    };
    reader.readAsArrayBuffer(file);
});

function procesarDatos(workbook) {
    function buscarHoja(palabrasClave) {
        const nombreReal = workbook.SheetNames.find(nombre => 
            palabrasClave.every(palabra => nombre.toLowerCase().includes(palabra))
        );
        return nombreReal ? workbook.Sheets[nombreReal] : null;
    }

    const sheetMercaderia = buscarHoja(['salidas', 'mercaderia']);
    const sheetServicios = buscarHoja(['salidas', 'servicio']); 
    const sheetIngresos = buscarHoja(['ingresos']);

    if (!sheetMercaderia || !sheetServicios || !sheetIngresos) {
        alert("Faltan hojas en el Excel. Presioná Ctrl + F5 y reintentá.");
        return; 
    }

    // Guardar en la variable global para poder filtrar sin volver a leer el Excel
    dataGlobal.mercaderia = XLSX.utils.sheet_to_json(sheetMercaderia);
    dataGlobal.servicio = XLSX.utils.sheet_to_json(sheetServicios);
    dataGlobal.ingresos = XLSX.utils.sheet_to_json(sheetIngresos);

    llenarSelectoresFiltros();
    actualizarDashboard(); 
    document.getElementById('dashboard').style.display = 'block';
}

function llenarSelectoresFiltros() {
    let anos = new Set(), meses = new Set(), proveedores = new Set();
    const todasSalidas = [...dataGlobal.mercaderia, ...dataGlobal.servicio];

    dataGlobal.ingresos.forEach(row => {
        if(row[COL_ANO]) anos.add(row[COL_ANO]);
        if(row[COL_MES]) meses.add(String(row[COL_MES]).toUpperCase().trim());
    });
    todasSalidas.forEach(row => {
        if(row[COL_ANO]) anos.add(row[COL_ANO]);
        if(row[COL_MES]) meses.add(String(row[COL_MES]).toUpperCase().trim());
        if(row[COL_PROVEEDOR]) proveedores.add(String(row[COL_PROVEEDOR]).toUpperCase().trim());
    });

    const llenarSelect = (id, setValores) => {
        const select = document.getElementById(id);
        select.innerHTML = '<option value="ALL">Todos</option>';
        [...setValores].sort().forEach(val => {
            if(val && val !== 'UNDEFINED') select.innerHTML += `<option value="${val}">${val}</option>`;
        });
    };

    llenarSelect('filter-year', anos);
    llenarSelect('filter-month', meses);
    llenarSelect('filter-provider', proveedores);

    // Activar eventos de cambio para actualizar todo
    document.getElementById('filter-year').addEventListener('change', actualizarDashboard);
    document.getElementById('filter-month').addEventListener('change', actualizarDashboard);
    document.getElementById('filter-provider').addEventListener('change', actualizarDashboard);
}

function filtrarDatos() {
    const selAno = document.getElementById('filter-year').value;
    const selMes = document.getElementById('filter-month').value;
    const selProv = document.getElementById('filter-provider').value;

    const aplicarFiltro = (row, filtrarProv = true) => {
        const rowAno = row[COL_ANO] ? String(row[COL_ANO]) : null;
        const rowMes = row[COL_MES] ? String(row[COL_MES]).toUpperCase().trim() : null;
        const rowProv = row[COL_PROVEEDOR] ? String(row[COL_PROVEEDOR]).toUpperCase().trim() : null;

        if (selAno !== 'ALL' && rowAno !== selAno) return false;
        if (selMes !== 'ALL' && rowMes !== selMes) return false;
        if (filtrarProv && selProv !== 'ALL' && rowProv !== selProv) return false;
        return true;
    };

    // Para los ingresos no filtramos por proveedor porque no tienen.
    return {
        ingresos: dataGlobal.ingresos.filter(r => aplicarFiltro(r, false)),
        mercaderia: dataGlobal.mercaderia.filter(r => aplicarFiltro(r, true)),
        servicio: dataGlobal.servicio.filter(r => aplicarFiltro(r, true))
    };
}

function formatearPlata(numero) {
    return `$${Math.round(numero).toLocaleString('es-AR')}`;
}

function actualizarDashboard() {
    const filtrados = filtrarDatos();
    const salidasMercaderia = filtrados.mercaderia;
    const salidasServicio = filtrados.servicio;
    const ingresos = filtrados.ingresos;
    const salidasTotales = [...salidasMercaderia, ...salidasServicio];

    // 1. KPIs
    const totalIng = ingresos.reduce((acc, row) => acc + (parseFloat(row[COL_MONTO]) || 0), 0);
    const totalSal = salidasTotales.reduce((acc, row) => acc + (parseFloat(row[COL_MONTO]) || 0), 0);
    const ingresosNetos = totalIng - totalSal; 
    document.getElementById('kpi-ytd').textContent = formatearPlata(ingresosNetos);

    const gastosPendientes = salidasTotales.reduce((acc, row) => acc + (parseFloat(row[COL_RESTO]) || 0), 0);
    document.getElementById('kpi-pendientes').textContent = formatearPlata(gastosPendientes);

    // 2. Gráficos
    generarGraficoEvolutivo(ingresos, salidasTotales);
    generarGraficoProveedores(salidasTotales);

    // 3. Detalle Estados de Proveedores
    generarResumenEstados(salidasTotales);

    // 4. Tablas
    llenarTabla('table-mercaderia', salidasMercaderia);
    llenarTabla('table-servicio', salidasServicio);
}

function generarResumenEstados(salidas) {
    const sumasPorEstado = { "OK": 0, "PENDIENTE": 0, "RESERVADO": 0, "EN PROCESO": 0 };
    
    salidas.forEach(row => {
        const est = row[COL_ESTADO] ? String(row[COL_ESTADO]).toUpperCase().trim() : 'OTRO';
        if (sumasPorEstado[est] !== undefined) {
            sumasPorEstado[est] += parseFloat(row[COL_MONTO]) || 0;
        }
    });

    const grid = document.getElementById('status-grid');
    grid.innerHTML = '';
    
    for (const [estado, total] of Object.entries(sumasPorEstado)) {
        // Excluimos estados con valor 0 para no ensuciar visualmente, o si preferís se muestran igual
        grid.innerHTML += `
            <div class="status-item">
                <span>${estado}</span>
                <strong>${formatearPlata(total)}</strong>
            </div>
        `;
    }
}

function formatearFecha(fecha) {
    if(!fecha) return '-';
    // Si la fecha ya viene como objeto Date de Excel
    if(fecha instanceof Date) {
        return fecha.toLocaleDateString('es-AR');
    }
    return String(fecha).substring(0, 10);
}

function llenarTabla(idTabla, datos) {
    const tbody = document.querySelector(`#${idTabla} tbody`);
    tbody.innerHTML = '';
    
    if(datos.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;">No hay datos para esta selección</td></tr>`;
        return;
    }

    datos.forEach(row => {
        const fecha = formatearFecha(row[COL_FECHA]);
        const prov = row[COL_PROVEEDOR] || '-';
        const conc = row[COL_CONCEPTO] || '-';
        const est = row[COL_ESTADO] || '-';
        const mon = formatearPlata(parseFloat(row[COL_MONTO]) || 0);

        tbody.innerHTML += `
            <tr>
                <td>${fecha}</td>
                <td><strong>${prov}</strong></td>
                <td>${conc}</td>
                <td>${est}</td>
                <td><strong>${mon}</strong></td>
            </tr>
        `;
    });
}

// Lógica de Gráficos (Actualizada para repintar bien en los filtros)
function generarGraficoEvolutivo(ingresos, salidas) {
    const mesesMap = {};
    ingresos.forEach(row => {
        let mes = row[COL_MES] ? String(row[COL_MES]).toUpperCase().trim() : 'OTRO';
        if (!mesesMap[mes]) mesesMap[mes] = { ing: 0, sal: 0 };
        mesesMap[mes].ing += parseFloat(row[COL_MONTO]) || 0;
    });
    salidas.forEach(row => {
        let mes = row[COL_MES] ? String(row[COL_MES]).toUpperCase().trim() : 'OTRO';
        if (!mesesMap[mes]) mesesMap[mes] = { ing: 0, sal: 0 };
        mesesMap[mes].sal += parseFloat(row[COL_MONTO]) || 0;
    });

    const labels = Object.keys(mesesMap).filter(m => m !== 'OTRO' && m !== 'UNDEFINED');
    const dataIng = labels.map(m => mesesMap[m].ing);
    const dataSal = labels.map(m => mesesMap[m].sal);

    const ctx = document.getElementById('evolutivoChart').getContext('2d');
    if (chartEvolutivo) chartEvolutivo.destroy();
    chartEvolutivo = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                { label: 'Ingresos', data: dataIng, backgroundColor: '#7D8C7A' },
                { label: 'Salidas Totales', data: dataSal, backgroundColor: '#B28B84' }
            ]
        },
        options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });
}

function generarGraficoProveedores(salidas) {
    const provMap = {};
    salidas.forEach(row => {
        const prov = row[COL_PROVEEDOR] ? String(row[COL_PROVEEDOR]).trim().toUpperCase() : 'OTROS';
        provMap[prov] = (provMap[prov] || 0) + (parseFloat(row[COL_MONTO]) || 0);
    });
    const ordenados = Object.entries(provMap).sort((a, b) => b[1] - a[1]).slice(0, 8);
    
    const ctx = document.getElementById('proveedoresChart').getContext('2d');
    if (chartProveedores) chartProveedores.destroy();
    chartProveedores = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ordenados.length ? ordenados.map(i => i[0]) : ['Sin datos'],
            datasets: [{
                data: ordenados.length ? ordenados.map(i => i[1]) : [1],
                backgroundColor: ['#A39B8B', '#7D8C7A', '#B28B84', '#D4CFC7', '#8A847A', '#6B655B', '#4D564C', '#C9BDB0'],
                borderWidth: 0
            }]
        },
        options: { responsive: true, plugins: { legend: { position: 'right' } }, cutout: '70%' }
    });
}
