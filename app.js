// Nombres de columnas esperados en tu Excel (Modificalos si son diferentes)
const COL_MONTO = 'Monto';
const COL_FECHA = 'Fecha';
const COL_ESTADO = 'Estado'; // Ej: "Pendiente", "Pagado"
const COL_PROVEEDOR = 'Proveedor';

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
    // 1. Leer las hojas (Verificá que los nombres de las pestañas en el Excel sean exactos)
    const sheetMercaderia = workbook.Sheets['salidas mercaderia'];
    const sheetServicios = workbook.Sheets['salida servicios'];
    const sheetIngresos = workbook.Sheets['ingresos'];

    if (!sheetMercaderia || !sheetServicios || !sheetIngresos) {
        alert("Atención: No se encontraron todas las hojas requeridas en el Excel.");
    }

    const dataMercaderia = sheetMercaderia ? XLSX.utils.sheet_to_json(sheetMercaderia) : [];
    const dataServicios = sheetServicios ? XLSX.utils.sheet_to_json(sheetServicios) : [];
    const dataIngresos = sheetIngresos ? XLSX.utils.sheet_to_json(sheetIngresos) : [];

    // Combinar todas las salidas
    const salidasTotales = [...dataMercaderia, ...dataServicios];

    // --- CÁLCULO YTD GANANCIAS ---
    const totalIngresos = dataIngresos.reduce((acc, row) => acc + (row[COL_MONTO] || 0), 0);
    const totalSalidas = salidasTotales.reduce((acc, row) => acc + (row[COL_MONTO] || 0), 0);
    const ytdGanancias = totalIngresos - totalSalidas;
    
    document.getElementById('kpi-ytd').textContent = `$${ytdGanancias.toLocaleString('es-AR')}`;

    // --- CÁLCULO GASTOS PENDIENTES ---
    const gastosPendientes = salidasTotales
        .filter(row => row[COL_ESTADO] && row[COL_ESTADO].toLowerCase() === 'pendiente')
        .reduce((acc, row) => acc + (row[COL_MONTO] || 0), 0);
        
    document.getElementById('kpi-pendientes').textContent = `$${gastosPendientes.toLocaleString('es-AR')}`;

    // --- PREPARAR DATOS PARA GRÁFICOS ---
    generarGraficoEvolutivo(dataIngresos, salidasTotales);
    generarGraficoProveedores(dataMercaderia); // Asumo que el detalle de proveedores sale de mercadería

    // Mostrar el dashboard
    document.getElementById('dashboard').style.display = 'block';
}

let chartEvolutivo = null;
let chartProveedores = null;

function generarGraficoEvolutivo(ingresos, salidas) {
    // Lógica simplificada: agrupar por mes
    // (En una versión final, aquí se extrae el mes de COL_FECHA de cada fila)
    // Para el diseño, usamos datos de prueba visuales basados en la estructura
    
    const ctx = document.getElementById('evolutivoChart').getContext('2d');
    if (chartEvolutivo) chartEvolutivo.destroy();

    chartEvolutivo = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun'], // Dinamizar según fechas reales
            datasets: [
                {
                    label: 'Ingresos',
                    data: [12000, 19000, 15000, 22000, 20000, 25000], // Reemplazar con datos procesados
                    backgroundColor: '#7D8C7A'
                },
                {
                    label: 'Salidas',
                    data: [8000, 15000, 10000, 18000, 12000, 14000], // Reemplazar con datos procesados
                    backgroundColor: '#B28B84'
                }
            ]
        },
        options: {
            responsive: true,
            plugins: { legend: { position: 'bottom' } },
            scales: {
                y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.05)' } },
                x: { grid: { display: false } }
            }
        }
    });
}

function generarGraficoProveedores(salidas) {
    // Agrupar por proveedor
    const proveedoresMap = {};
    salidas.forEach(row => {
        const prov = row[COL_PROVEEDOR] || 'Otros';
        proveedoresMap[prov] = (proveedoresMap[prov] || 0) + (row[COL_MONTO] || 0);
    });

    const labels = Object.keys(proveedoresMap);
    const data = Object.values(proveedoresMap);

    const ctx = document.getElementById('proveedoresChart').getContext('2d');
    if (chartProveedores) chartProveedores.destroy();

    chartProveedores = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: ['#A39B8B', '#7D8C7A', '#B28B84', '#D4CFC7', '#8A847A'],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            plugins: { legend: { position: 'bottom' } },
            cutout: '70%'
        }
    });
}
