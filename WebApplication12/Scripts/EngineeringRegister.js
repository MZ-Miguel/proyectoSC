document.getElementById('btn_Conciliacion_Cargar').onclick = function () {
    document.getElementById('fup_EX_Archivo').value = null;
    document.getElementById('fup_EX_Archivo').click();
};
var fupArchivo = document.getElementById('fup_EX_Archivo');
fupArchivo.onchange = function () {
    var file = this.files[0];
    var oRequest = new FormData();
    oRequest.append("archivo", file);
    //$service.leerExcel(oRequest, function (d) {
    //    console.log(d.Data);
    //});
    var xhr = new XMLHttpRequest();
    xhr.open("post", "../EngineeringRegister/LeerExcel");
    xhr.onreadystatechange = function () {
        if (xhr.status === 200 && xhr.readyState === 4) {
            fnMostrarTablaExcel(xhr.responseText);
        }
    };
    xhr.send(oRequest);
};
var fnMostrarTablaExcel = function (rpta) {

    //if (rpta !== '') {
    //    var aResponse = rpta.split('»');
    //    matrizExcel = aResponse[1].split('¬');
    //    $$table('div_Conciliacion_TableConciliacion').setData(aResponse[1].split('¬'));
    //}

    if (rpta !== '') {
        var aResponse = rpta.split('»');
        if (aResponse[0] === '1') {
            var lHojas = aResponse[1] !== '' ? aResponse[1].split('¯') : [];
            var aFilas, aCeldas;
            var i, j, k, nHojas, nFilas, nCeldas;
            var aContenido = [];
            nHojas = lHojas.length;
            if (nHojas > 0) {
                for (i = 0; i < nHojas; i++) {
                    aFilas = lHojas[i].split('¬');
                    nFilas = aFilas.length;
                    aContenido.push('<table class="table w-100">');
                    aContenido.push('<caption>Hoja ');
                    aContenido.push(i + 1);
                    aContenido.push('</caption>');
                    aContenido.push('<tbody>');
                    for (j = 0; j < nFilas; j++) {
                        aCeldas = aFilas[j].split('¦');
                        nCeldas = aCeldas.length;
                        aContenido.push('<tr>');
                        for (k = 0; k < nCeldas; k++) {
                            aContenido.push('<td>');
                            aContenido.push(aCeldas[k]);
                            aContenido.push('</td>');
                        }
                        aContenido.push('</tr>');
                    }
                    aContenido.push('</tbody></table></br></br>');
                }
            }
        } else {
            alert(aResponse[1]);
        }
        var divContenedor = document.getElementById('div_Conciliacion_TableConciliacion');
        if (divContenedor !== null) divContenedor.innerHTML = aContenido.join('');

    }
};