// --- CONFIGURACIÓN ESTRICTA (Debe coincidir con tus cabeceras) ---
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
            
            // LOG DE DEPURACIÓN #1: Ver qué hojas tiene el Excel
            console.log("Hojas encontradas en el archivo:", workbook.SheetNames);
            
            procesarDatos(workbook);
        } catch (err) {
            console.error("Error al leer el archivo:", err);
            alert("Error al leer el archivo. Revisa la consola (F12) para detalles.");
        }
    };
    reader.readAsArrayBuffer(file);
});

function procesarDatos(workbook) {
    // Función de búsqueda flexible
    const obtenerHoja = (palabras) => {
        const nombreHoja = workbook.SheetNames.find(n => 
            palabras.some(p => n.toLowerCase().includes(p.toLowerCase()))
        );
        if (!nombreHoja) {
            console.warn("No se encontró hoja con palabras clave:", palabras);
            return [];
        }
        const json = XLSX.utils.sheet_to_json(workbook.Sheets[nombreHoja]);
        console.log(`Datos cargados de ${nombreHoja}:`, json.length, "filas.");
        if (json.length > 0) console.log("Ejemplo de columnas detectadas:", Object.keys(json[0]));
        return json;
    };

    dataGlobal.mercaderia = obtenerHoja(['mercaderia', 'salidas']);
    dataGlobal.servicio = obtenerHoja(['servicio', 'salidas']);
    dataGlobal.ingresos = obtenerHoja(['ingresos']);

    if (dataGlobal.mercaderia.length === 0 && dataGlobal.servicio.length === 0) {
        alert("¡Error! No pude extraer datos. Verifica en la Consola (F12) qué hojas encontró el sistema.");
        return;
    }

    actualizarDashboard();
    document.getElementById('dashboard').style.display = 'block';
}

function actualizarDashboard() {
    // Para simplificar, mostramos todo sin filtrar primero para verificar que carga
    const salidasTotales = [...dataGlobal.mercaderia, ...dataGlobal.servicio];
    
    // KPIs simples
    const totalIng = dataGlobal.ingresos.reduce((acc, row) => acc + (parseFloat(row[COL_MONTO]) || 0), 0);
    const totalSal = salidasTotales.reduce((acc, row) => acc + (parseFloat(row[COL_MONTO]) || 0), 0);
    
    document.getElementById('kpi-ingresos').textContent = formatearPlata(totalIng);
    document.getElementById('kpi-ytd').textContent = formatearPlata(totalIng - totalSal);
    
    // Renderizado de tablas
    llenarTabla('table-mercaderia', dataGlobal.mercaderia);
    llenarTabla('table-servicio', dataGlobal.servicio);
}

function llenarTabla(id, datos) {
    const tbody = document.querySelector(`#${id} tbody`);
    if (!tbody) return;
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

function formatearPlata(n) { return `$${Math.round(n || 0).toLocaleString('es-AR')}`; }
