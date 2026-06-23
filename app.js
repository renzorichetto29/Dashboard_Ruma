// Nombres de columnas adaptados a MAYÚSCULAS tal cual tu Excel
const COL_MONTO = 'MONTO';
const COL_FECHA = 'FECHA';
const COL_ESTADO = 'ESTADO'; 
const COL_PROVEEDOR = 'PROVEEDOR';

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
    // Función inteligente para encontrar hojas
    function buscarHoja(nombreBuscado) {
        const nombreNormalizado = nombreBuscado.toLowerCase().trim();
        const nombreReal = workbook.SheetNames.find(
            nombre => nombre.toLowerCase().trim() === nombreNormalizado
        );
        return nombreReal ? workbook.Sheets[nombreReal] : null;
    }

    // 1. Leer las hojas (corregido "salidas servicio")
    const sheetMercaderia = buscarHoja('salidas mercaderia');
    const sheetServicios = buscarHoja('salidas servicio'); 
    const sheetIngresos = buscarHoja('ingresos');

    if (!sheetMercaderia || !sheetServicios || !sheetIngresos) {
        const hojasEncontradas = workbook.SheetNames.join(" | ");
        alert(`Atención: No se encontraron todas las hojas.\nTu Excel tiene estas hojas: ${hojasEncontradas}`);
        return; 
    }

    const dataMercaderia = XLSX.utils.sheet_to_json(sheetMercaderia);
    const dataServicios = XLSX.utils.sheet_to_json(sheetServicios);
    const dataIngresos = XLSX.utils.sheet_to_json(sheetIngresos);

    // Combinar todas las salidas
    const salidasTotales = [...dataMercaderia, ...dataServicios];

    // --- CÁLCULO YTD GANANCIAS ---
    const totalIngresos = dataIngresos.reduce((acc, row) => acc + (row[COL_MONTO] || 0), 0);
    const totalSalidas = salidasTotales.reduce((acc, row) => acc + (row[COL_MONTO] || 0), 0);
    const ytdGanancias = totalIngresos - totalSalidas;
    
    // Formatear a moneda argentina
    document.getElementById('kpi-ytd').textContent = `$${ytdGanancias.toLocaleString('es-AR')}`;

    // --- CÁLCULO GASTOS PENDIENTES ---
    const gastosPendientes = salidasTotales
        .filter(row => {
            const estado = row[COL_ESTADO] ? String(row[COL_ESTADO]).toLowerCase().trim() : '';
            return estado === 'pendiente'; // Suma solo los que digan "Pendiente"
        })
        .reduce((acc, row) => acc + (row[COL_MONTO] || 0), 0);
        
    document.getElementById('kpi-pendientes').textContent = `$${gastosPendientes.toLocaleString('es-AR')}`;

    // --- PREPARAR DATOS PARA GRÁFICOS ---
    generarGraficoEvolutivo(dataIngresos, salidasTotales);
    generarGraficoProveedores(dataMercaderia); 

    // Mostrar el dashboard
    document.getElementById('dashboard').style.display = 'block';
}

let chartEvolutivo = null;
let chartProveedores = null;

function generarGraficoEvolutivo(ingresos, salidas) {
    const ctx = document.getElementById('evolutivoChart').getContext('2d');
    if (chartEvolutivo) chartEvolutivo.destroy();

    // Gráfico ilustrativo (se puede dinamizar luego para leer la columna 'MES')
    chartEvolutivo = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun'], 
            datasets: [
                {
                    label: 'Ingresos',
                    data: [12000, 19000, 15000, 22000, 20000, 25000], 
                    backgroundColor: '#7D8C7A'
                },
                {
                    label: 'Salidas',
                    data: [8000, 15000, 10000, 18000, 12000, 14000], 
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
            labels: labels.length > 0 ? labels : ['Sin datos'],
            datasets: [{
                data: data.length > 0 ? data : [1],
                backgroundColor: ['#A39B8B', '#7D8C7A', '#B28B84', '#D4CFC7', '#8A847A', '#6B655B'],
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
