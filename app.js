// --- CONFIGURACIÓN DE COLUMNAS (DEBEN COINCIDIR CON TU EXCEL) ---
const COL_MONTO = 'MONTO';
const COL_RESTO = 'RESTO'; 
const COL_PROVEEDOR = 'PROVEEDOR';
const COL_MES = 'MES';
const COL_ANO = 'AÑO';
const COL_FECHA = 'FECHA';
const COL_ESTADO = 'ESTADO';
const COL_CONCEPTO = 'CONCEPTO';

// Variables globales
let dataGlobal = { mercaderia: [], servicio: [], ingresos: [] };
let chartEvolutivo = null;
let chartProveedores = null;

// Inicialización de lectura
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
        } catch (err) {
            console.error(err);
            alert("Error al leer el archivo Excel. Asegurate de que es un .xlsx válido.");
        }
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
        alert("No se encontraron las hojas necesarias. El Excel debe tener hojas que contengan los nombres: 'salidas mercaderia', 'salidas servicio' e 'ingresos'.");
        return; 
    }

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
        if(row[COL_ANO]) anos.add(String(row[COL_ANO]));
        if(row[COL_MES]) meses.add(String(row[COL_MES]).toUpperCase().trim());
    });
    todasSalidas.forEach(row => {
        if(row[COL_ANO]) anos.add(String(row[COL_ANO]));
        if(row[COL_MES]) meses.add(String(row[COL_MES]).toUpperCase().trim());
        if(row[COL_PROVEEDOR]) proveedores.add(String(row[COL_PROVEEDOR]).toUpperCase().trim());
    });

    const llenarSelect = (id, setValores) => {
        const select = document.getElementById(id);
        select.innerHTML = '<option value="ALL">Todos</option>';
        [...setValores].sort().forEach(val => {
            select.innerHTML += `<option value="${val}">${val}</option>`;
        });
    };

    llenarSelect('filter-year', anos);
    llenarSelect('filter-month', meses);
    llenarSelect('filter-provider', proveedores);

    document.getElementById('filter-year').onchange = actualizarDashboard;
    document.getElementById('filter-month').onchange = actualizarDashboard;
    document.getElementById('filter-provider').onchange = actualizarDashboard;
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

    return {
        ingresos: dataGlobal.ingresos.filter(r => aplicarFiltro(r, false)),
        mercaderia: dataGlobal.mercaderia.filter(r => aplicarFiltro(r, true)),
        servicio: dataGlobal.servicio.filter(r => aplicarFiltro(r, true))
    };
}

function formatearPlata(n) {
    return `$${Math.round(n || 0).toLocaleString('es-AR')}`;
}

function actualizarDashboard() {
    const filtrados = filtrarDatos();
    const salidasTotales = [...filtrados.mercaderia, ...filtrados.servicio];

    // KPIs
    const totalIng = filtrados.ingresos.reduce((acc, row) => acc + (parseFloat(row[COL_MONTO]) || 0), 0);
    const totalSal = salidasTotales.reduce((acc, row) => acc + (parseFloat(row[COL_MONTO]) || 0), 0);
    
    document.getElementById('kpi-ingresos').textContent = formatearPlata(totalIng);
    document.getElementById('kpi-ytd').textContent = formatearPlata(totalIng - totalSal);
    document.getElementById('kpi-pendientes').textContent = formatearPlata(salidasTotales.reduce((acc, row) => acc + (parseFloat(row[COL_RESTO]) || 0), 0));

    // Gráficos y Tablas
    generarGraficoEvolutivo(filtrados.ingresos, salidasTotales);
    generarGraficoProveedores(salidasTotales);
    generarResumenEstados(salidasTotales);
    llenarTabla('table-mercaderia', filtrados.mercaderia);
    llenarTabla('table-servicio', filtrados.servicio);
}

function generarResumenEstados(salidas) {
    const sumas = { "OK": 0, "PENDIENTE": 0, "RESERVADO": 0, "EN PROCESO": 0 };
    salidas.forEach(row => {
        const est = row[COL_ESTADO] ? String(row[COL_ESTADO]).toUpperCase().trim() : 'OTRO';
        if (sumas.hasOwnProperty(est)) sumas[est] += parseFloat(row[COL_MONTO]) || 0;
    });
    const grid = document.getElementById('status-grid');
    grid.innerHTML = Object.entries(sumas).map(([k, v]) => `
        <div class="status-item"><span>${k}</span><strong>${formatearPlata(v)}</strong></div>
    `).join('');
}

function llenarTabla(id, datos) {
    const tbody = document.querySelector(`#${id} tbody`);
    tbody.innerHTML = datos.map(row => `
        <tr>
            <td>${row[COL_FECHA] || '-'}</td>
            <td><strong>${row[COL_PROVEEDOR] || '-'}</strong></td>
            <td>${row[COL_CONCEPTO] || '-'}</td>
            <td>${row[COL_ESTADO] || '-'}</td>
            <td><strong>${formatearPlata(row[COL_MONTO])}</strong></td>
            <td style="color: #B28B84;">${row[COL_ESTADO] === 'EN PROCESO' ? formatearPlata(row[COL_RESTO]) : '-'}</td>
        </tr>
    `).join('');
}

function generarGraficoEvolutivo(ingresos, salidas) {
    const mesesMap = {};
    [...ingresos, ...salidas].forEach(row => {
        let mes = row[COL_MES] ? String(row[COL_MES]).toUpperCase().trim() : 'OTRO';
        if (!mesesMap[mes]) mesesMap[mes] = { ing: 0, sal: 0 };
    });
    ingresos.forEach(r => mesesMap[r[COL_MES]].ing += parseFloat(r[COL_MONTO] || 0));
    salidas.forEach(r => mesesMap[r[COL_MES]].sal += parseFloat(r[COL_MONTO] || 0));

    const labels = Object.keys(mesesMap);
    const ctx = document.getElementById('evolutivoChart').getContext('2d');
    if (chartEvolutivo) chartEvolutivo.destroy();
    chartEvolutivo = new Chart(ctx, {
        type: 'bar',
        data: { labels, datasets: [
            { label: 'Ingresos', data: labels.map(m => mesesMap[m].ing), backgroundColor: '#7D8C7A' },
            { label: 'Salidas', data: labels.map(m => mesesMap[m].sal), backgroundColor: '#B28B84' }
        ]},
        options: { responsive: true }
    });
}

function generarGraficoProveedores(salidas) {
    const provMap = {};
    salidas.forEach(r => provMap[r[COL_PROVEEDOR]] = (provMap[r[COL_PROVEEDOR]] || 0) + parseFloat(r[COL_MONTO] || 0));
    const ordenados = Object.entries(provMap).sort((a,b) => b[1] - a[1]).slice(0,8);
    const ctx = document.getElementById('proveedoresChart').getContext('2d');
    if (chartProveedores) chartProveedores.destroy();
    chartProveedores = new Chart(ctx, {
        type: 'bar',
        data: { labels: ordenados.map(i => i[0]), datasets: [{ data: ordenados.map(i => i[1]), backgroundColor: '#A39B8B' }]},
        options: { indexAxis: 'y', responsive: true, plugins: { legend: { display: false } } }
    });
}
