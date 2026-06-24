// ======================================================
// CONFIGURACIÓN COLUMNAS EXCEL
// ======================================================

const COL_MONTO = 'MONTO';
const COL_RESTO = 'RESTO';
const COL_PROVEEDOR = 'PROVEEDOR';
const COL_MES = 'MES';
const COL_ANO = 'AÑO';
const COL_FECHA = 'FECHA';
const COL_ESTADO = 'ESTADO';
const COL_CONCEPTO = 'CONCEPTO';


let dataGlobal = {
    mercaderia: [],
    servicio: [],
    ingresos: []
};


let chartEvolutivo = null;
let chartProveedores = null;




// ======================================================
// CARGA EXCEL
// ======================================================


document.getElementById('excel-upload').addEventListener('change', function(e){

    const file = e.target.files[0];

    if(!file) return;


    document.getElementById('file-name').textContent = file.name;


    const reader = new FileReader();


    reader.onload = function(e){

        const data = new Uint8Array(e.target.result);

        const workbook = XLSX.read(data,{
            type:'array',
            cellDates:true
        });


        procesarDatos(workbook);

    };


    reader.readAsArrayBuffer(file);

});





// ======================================================
// PROCESAR EXCEL
// ======================================================


function procesarDatos(workbook){


    function buscarHoja(palabras){

        const nombre = workbook.SheetNames.find(nombre =>
            palabras.every(p =>
                nombre.toLowerCase().includes(p)
            )
        );


        return nombre ? workbook.Sheets[nombre] : null;

    }



    const sheetMercaderia = buscarHoja(['salidas','mercaderia']);
    const sheetServicios = buscarHoja(['salidas','servicio']);
    const sheetIngresos = buscarHoja(['ingresos']);



    if(!sheetMercaderia || !sheetServicios || !sheetIngresos){

        alert("Faltan hojas necesarias en el Excel.");

        return;

    }



    dataGlobal.mercaderia =
        XLSX.utils.sheet_to_json(sheetMercaderia);



    dataGlobal.servicio =
        XLSX.utils.sheet_to_json(sheetServicios);



    dataGlobal.ingresos =
        XLSX.utils.sheet_to_json(sheetIngresos);



    llenarSelectoresFiltros();


    actualizarDashboard();



    document.getElementById('dashboard').style.display='block';


}






// ======================================================
// FILTROS
// ======================================================


function llenarSelectoresFiltros(){


    let anos = new Set();
    let meses = new Set();
    let proveedores = new Set();



    const salidas = [
        ...dataGlobal.mercaderia,
        ...dataGlobal.servicio
    ];



    [...dataGlobal.ingresos,...salidas].forEach(row=>{


        if(row[COL_ANO])
            anos.add(row[COL_ANO]);



        if(row[COL_MES])
            meses.add(
                String(row[COL_MES])
                .toUpperCase()
                .trim()
            );


    });



    salidas.forEach(row=>{


        if(row[COL_PROVEEDOR])
            proveedores.add(
                String(row[COL_PROVEEDOR])
                .toUpperCase()
                .trim()
            );


    });





    function llenar(id,set){


        const select =
            document.getElementById(id);



        select.innerHTML =
        '<option value="ALL">Todos</option>';



        [...set]
        .sort()
        .forEach(valor=>{


            if(valor && valor!=="UNDEFINED"){


                select.innerHTML +=
                `<option value="${valor}">
                ${valor}
                </option>`;

            }


        });


    }




    llenar('filter-year',anos);
    llenar('filter-month',meses);
    llenar('filter-provider',proveedores);





    document.getElementById('filter-year')
    .onchange = actualizarDashboard;


    document.getElementById('filter-month')
    .onchange = actualizarDashboard;


    document.getElementById('filter-provider')
    .onchange = actualizarDashboard;



}








function filtrarDatos(){


    const ano =
    document.getElementById('filter-year').value;


    const mes =
    document.getElementById('filter-month').value;


    const prov =
    document.getElementById('filter-provider').value;




    function filtro(row,usarProveedor=true){


        const rAno =
        row[COL_ANO]
        ? String(row[COL_ANO])
        : '';



        const rMes =
        row[COL_MES]
        ? String(row[COL_MES])
        .toUpperCase()
        .trim()
        : '';



        const rProv =
        row[COL_PROVEEDOR]
        ? String(row[COL_PROVEEDOR])
        .toUpperCase()
        .trim()
        : '';





        if(ano!=='ALL' && rAno!==ano)
            return false;



        if(mes!=='ALL' && rMes!==mes)
            return false;



        if(
            usarProveedor &&
            prov!=='ALL' &&
            rProv!==prov
        )
            return false;



        return true;


    }




    return {

        ingresos:
        dataGlobal.ingresos.filter(r=>filtro(r,false)),


        mercaderia:
        dataGlobal.mercaderia.filter(r=>filtro(r,true)),


        servicio:
        dataGlobal.servicio.filter(r=>filtro(r,true))

    };


}





// ======================================================
// DASHBOARD PRINCIPAL
// ======================================================


function actualizarDashboard(){


    const datos = filtrarDatos();



    const ingresos = datos.ingresos;

    const mercaderia = datos.mercaderia;

    const servicio = datos.servicio;



    const salidas = [
        ...mercaderia,
        ...servicio
    ];



    const totalIngresos =
    ingresos.reduce(
        (a,r)=>
        a+(parseFloat(r[COL_MONTO])||0),
        0
    );



    const totalSalidas =
    salidas.reduce(
        (a,r)=>
        a+(parseFloat(r[COL_MONTO])||0),
        0
    );



    // INGRESOS NETOS
    document.getElementById('kpi-ytd')
    .textContent =
    formatearPlata(
        totalIngresos-totalSalidas
    );



    // NUEVO KPI INGRESOS BRUTOS
    document.getElementById('kpi-ingresos-brutos')
    .textContent =
    formatearPlata(totalIngresos);




    // RESTO SOLO EN PROCESO

    const pendientes =
    salidas.reduce((a,r)=>{


        const estado =
        String(r[COL_ESTADO]||'')
        .toUpperCase()
        .trim();



        if(estado==="EN PROCESO"){

            return a+
            (parseFloat(r[COL_RESTO])||0);

        }


        return a;


    },0);




    document.getElementById('kpi-pendientes')
    .textContent =
    formatearPlata(pendientes);




    generarGraficoEvolutivo(
        ingresos,
        salidas
    );


    generarGraficoProveedores(
        salidas
    );


    generarResumenEstados(
        salidas
    );



    llenarTabla(
        'table-mercaderia',
        mercaderia
    );


    llenarTabla(
        'table-servicio',
        servicio
    );


}
// ======================================================
// UTILIDADES
// ======================================================


function formatearPlata(numero){

    return '$' +
    Math.round(numero)
    .toLocaleString('es-AR');

}





function formatearFecha(fecha){

    if(!fecha) return '-';


    if(fecha instanceof Date){

        return fecha.toLocaleDateString('es-AR');

    }


    return String(fecha).substring(0,10);

}







// ======================================================
// TABLAS
// ======================================================


function llenarTabla(id,datos){


    const tbody =
    document.querySelector(`#${id} tbody`);



    tbody.innerHTML='';



    if(datos.length===0){

        tbody.innerHTML =
        `
        <tr>
            <td colspan="6">
            No hay datos para esta selección
            </td>
        </tr>
        `;

        return;

    }





    datos.forEach(row=>{


        const estado =
        String(row[COL_ESTADO]||'')
        .toUpperCase()
        .trim();




        let resto = 0;


        if(estado==="EN PROCESO"){

            resto =
            parseFloat(row[COL_RESTO])||0;

        }




        tbody.innerHTML +=
        `

        <tr>

            <td>
            ${formatearFecha(row[COL_FECHA])}
            </td>


            <td>
            <strong>
            ${row[COL_PROVEEDOR]||'-'}
            </strong>
            </td>


            <td>
            ${row[COL_CONCEPTO]||'-'}
            </td>


            <td>
            ${row[COL_ESTADO]||'-'}
            </td>


            <td>
            <strong>
            ${formatearPlata(
                parseFloat(row[COL_MONTO])||0
            )}
            </strong>
            </td>


            <td>
            <strong>
            ${formatearPlata(resto)}
            </strong>
            </td>


        </tr>

        `;


    });


}








// ======================================================
// ESTADOS PROVEEDORES
// ======================================================


function generarResumenEstados(salidas){


    const estados = {

        "OK":0,
        "PENDIENTE":0,
        "RESERVADO":0,
        "EN PROCESO":0

    };



    salidas.forEach(row=>{


        const estado =
        String(row[COL_ESTADO]||'')
        .toUpperCase()
        .trim();




        if(estados[estado]!==undefined){

            estados[estado]+=
            parseFloat(row[COL_MONTO])||0;

        }


    });





    const grid =
    document.getElementById('status-grid');



    grid.innerHTML='';




    Object.entries(estados)
    .forEach(([estado,total])=>{


        grid.innerHTML +=
        `

        <div class="status-item">

            <span>
            ${estado}
            </span>


            <strong>
            ${formatearPlata(total)}
            </strong>

        </div>


        `;


    });


}








// ======================================================
// GRAFICO EVOLUTIVO
// ======================================================


function generarGraficoEvolutivo(ingresos,salidas){


    const mapa={};



    ingresos.forEach(row=>{


        const mes =
        String(row[COL_MES]||'OTRO')
        .toUpperCase()
        .trim();



        if(!mapa[mes])
            mapa[mes]={
                ing:0,
                sal:0
            };



        mapa[mes].ing +=
        parseFloat(row[COL_MONTO])||0;


    });





    salidas.forEach(row=>{


        const mes =
        String(row[COL_MES]||'OTRO')
        .toUpperCase()
        .trim();



        if(!mapa[mes])
            mapa[mes]={
                ing:0,
                sal:0
            };



        mapa[mes].sal +=
        parseFloat(row[COL_MONTO])||0;



    });





    const labels =
    Object.keys(mapa)
    .filter(x=>x!=="OTRO");



    const ingresosData =
    labels.map(x=>mapa[x].ing);



    const salidasData =
    labels.map(x=>mapa[x].sal);




    const ctx =
    document
    .getElementById('evolutivoChart')
    .getContext('2d');




    if(chartEvolutivo)
        chartEvolutivo.destroy();





    chartEvolutivo =
    new Chart(ctx,{


        type:'bar',


        data:{


            labels:labels,


            datasets:[


                {
                    label:'Ingresos',
                    data:ingresosData,
                    backgroundColor:'#7D8C7A'
                },


                {
                    label:'Salidas',
                    data:salidasData,
                    backgroundColor:'#B28B84'
                }


            ]

        },


        options:{

            responsive:true,

            plugins:{

                legend:{
                    position:'bottom'
                }

            }


        }


    });


}









// ======================================================
// GRAFICO PROVEEDORES - BARRAS HORIZONTALES
// ======================================================


function generarGraficoProveedores(salidas){


    const proveedores={};



    salidas.forEach(row=>{


        const prov =
        String(row[COL_PROVEEDOR]||'OTROS')
        .toUpperCase()
        .trim();




        proveedores[prov] =
        (proveedores[prov]||0)
        +
        (parseFloat(row[COL_MONTO])||0);



    });





    const ordenados =
    Object.entries(proveedores)
    .sort((a,b)=>b[1]-a[1])
    .slice(0,8);





    const ctx =
    document
    .getElementById('proveedoresChart')
    .getContext('2d');




    if(chartProveedores)
        chartProveedores.destroy();






    chartProveedores =
    new Chart(ctx,{



        type:'bar',



        data:{


            labels:
            ordenados.length
            ?
            ordenados.map(x=>x[0])
            :
            ['Sin datos'],



            datasets:[

                {

                    label:'Monto',

                    data:
                    ordenados.length
                    ?
                    ordenados.map(x=>x[1])
                    :
                    [0],


                    backgroundColor:'#7D8C7A'


                }

            ]


        },





        options:{


            indexAxis:'y',


            responsive:true,


            plugins:{


                legend:{
                    display:false
                }


            },


            scales:{


                x:{


                    ticks:{


                        callback:function(value){

                            return formatearPlata(value);

                        }


                    }


                }


            }


        }



    });



}
