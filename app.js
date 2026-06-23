// ... (mantené las constantes y el inicio igual hasta llegar a actualizarDashboard) ...

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
    
    document.getElementById('kpi-ingresos').textContent = formatearPlata(totalIng);
    document.getElementById('kpi-ytd').textContent = formatearPlata(ingresosNetos);

    const gastosPendientes = salidasTotales.reduce((acc, row) => acc + (parseFloat(row[COL_RESTO]) || 0), 0);
    document.getElementById('kpi-pendientes').textContent = formatearPlata(gastosPendientes);

    // 2. Gráficos y Tablas
    generarGraficoEvolutivo(ingresos, salidasTotales);
    generarGraficoProveedores(salidasTotales);
    generarResumenEstados(salidasTotales);
    llenarTabla('table-mercaderia', salidasMercaderia);
    llenarTabla('table-servicio', salidasServicio);
}

// NUEVO: Gráfico de Barras Horizontales para Proveedores
function generarGraficoProveedores(salidas) {
    const provMap = {};
    salidas.forEach(row => {
        const prov = row[COL_PROVEEDOR] ? String(row[COL_PROVEEDOR]).trim().toUpperCase() : 'OTROS';
        provMap[prov] = (provMap[prov] || 0) + (parseFloat(row[COL_MONTO]) || 0);
    });
    
    // Ordenar de mayor a menor
    const ordenados = Object.entries(provMap).sort((a, b) => b[1] - a[1]).slice(0, 8);
    
    const ctx = document.getElementById('proveedoresChart').getContext('2d');
    if (chartProveedores) chartProveedores.destroy();
    
    chartProveedores = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ordenados.map(i => i[0]),
            datasets: [{
                label: 'Monto Salida',
                data: ordenados.map(i => i[1]),
                backgroundColor: '#A39B8B'
            }]
        },
        options: {
            indexAxis: 'y', // ESTO HACE QUE SEAN BARRAS HORIZONTALES
            responsive: true,
            plugins: { legend: { display: false } }
        }
    });
}

// NUEVO: Lógica de tabla con columna RESTO dinámica
function llenarTabla(idTabla, datos) {
    const tbody = document.querySelector(`#${idTabla} tbody`);
    tbody.innerHTML = '';
    
    datos.forEach(row => {
        const fecha = formatearFecha(row[COL_FECHA]);
        const prov = row[COL_PROVEEDOR] || '-';
        const est = row[COL_ESTADO] ? String(row[COL_ESTADO]).toUpperCase().trim() : '';
        const mon = formatearPlata(parseFloat(row[COL_MONTO]) || 0);
        
        // Lógica de RESTO: Si es "EN PROCESO", mostramos el valor, sino guión
        const resto = est === 'EN PROCESO' ? formatearPlata(parseFloat(row[COL_RESTO]) || 0) : '-';

        tbody.innerHTML += `
            <tr>
                <td>${fecha}</td>
                <td><strong>${prov}</strong></td>
                <td>${row[COL_CONCEPTO] || '-'}</td>
                <td>${est}</td>
                <td><strong>${mon}</strong></td>
                <td style="color: #B28B84;">${resto}</td>
            </tr>
        `;
    });
}
