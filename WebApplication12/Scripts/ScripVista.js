
const mostrarPassword = () => {
    var cambio = document.getElementById("txtPassword");
    if (cambio.type == "password") {
        cambio.type = "text";
        $('.icon').removeClass('fa fa-eye').addClass('fa fa-eye-slash');
    } else {
        cambio.type = "password";
        $('.icon').removeClass('fa fa-eye-slash').addClass('fa fa-eye');
    }
}

const menuNav = () => {
    const menuToggle = document.getElementById('wrapper');
    if (menuToggle.className.includes('toggled')) {
        $('#wrapper').removeClass('toggled');
        $('#menu-toggle').removeClass('fas fa-angle-right').addClass('fas fa-angle-left');
    } else {
        $('#wrapper').addClass('toggled');
        $('#menu-toggle').removeClass('fas fa-angle-left').addClass('fas fa-angle-right');
    }

};

const importIngenieria = () => {

    const realFileBtnIngenieria = document.getElementById('fup_EX_Archivo');
    const nameTextIngenieria = document.getElementById('textFileIngenieria');

    realFileBtnIngenieria.addEventListener('change', () => {

        if (realFileBtnIngenieria.value) {
            nameTextIngenieria.innerHTML = realFileBtnIngenieria.value.match(/[\/\\]([\w\d\s\.\-\(\)]+)$/)[1];
        } else {
            nameTextIngenieria.innerHTML = 'No hay archivo...'
        }

    })

}

const importManual = () => {

    const realFileBtnManual = document.getElementById('idManual');
    const nameTextManual = document.getElementById('textFileManual');

    realFileBtnManual.addEventListener('change', () => {

        if (realFileBtnManual.value) {
            nameTextManual.innerHTML = realFileBtnManual.value.match(/[\/\\]([\w\d\s\.\-\(\)]+)$/)[1];
        } else {
            nameTextManual.innerHTML = 'No hay archivo...'
        }

    })

}

const importRepuesto = () => {

    const realFileBtnRepuesto = document.getElementById('idRepuesto');
    const nameTextRepuesto = document.getElementById('textFileRepuesto');

    realFileBtnRepuesto.addEventListener('change', () => {

        if (realFileBtnRepuesto.value) {
            nameTextRepuesto.innerHTML = realFileBtnRepuesto.value.match(/[\/\\]([\w\d\s\.\-\(\)]+)$/)[1];
        } else {
            nameTextRepuesto.innerHTML = 'No hay archivo...'
        }

    })

}