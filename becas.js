// Programa para subir los datos de los beneficiarios de Educacion a la base de datos
// desde un archivo xlsx

import fetch from "node-fetch";
import XLSX from "xlsx";

// Obtencion de token, archivo xlsx y url de la API
const token = "ea8420e4f516eeb8a06837b2e395f928";

const file = "./becas_total.xlsx";
const file2 = "./becas_150725.xlsx";
const apiUrl = "https://rubjrz.com/api/index.php";

// Lectura de un archivo xlsx y convertir sus datos a JSON
const workbook = XLSX.readFile(file);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet);

const workbook2 = XLSX.readFile(file2);
const sheet2 = workbook2.Sheets[workbook2.SheetNames[0]];
const data2 = XLSX.utils.sheet_to_json(sheet2);

// Funcion buscar beneficiario de la API
async function buscarBeneficiario(curp) {
    const body = JSON.stringify({
        token,
        funcion: "buscarBeneficiario",
        datos: [{ busqueda: curp }]
    });

    const res = await fetch(apiUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: body,
        redirect: "follow"
    });

    const data = await res.json();
    // console.log(data);
    return data;
}

// Funcion agregar beneficiario de la API
async function agregarBeneficiario(beneficiario) {
    const body = JSON.stringify({
        token,
        funcion: "agregarBeneficiario",
        datos: [beneficiario]
    });

    const res = await fetch(apiUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: body,
        redirect: "follow"
    });

    const text = await res.text();
    // console.log(text);
    try {
        const data = JSON.parse(text);
        return data;
    } catch (err) {
        console.log("Error al parsear el JSON:", err);
        throw new Error("Respuesta no es JSON: " + text);
    };
};

// Funcion agregar apoyo de la API
async function agregarApoyo(beneficiario) {
    const body = JSON.stringify({
        token,
        funcion: "agregarApoyoBeneficiario",
        datos: [beneficiario]
    });

    const res = await fetch(apiUrl, {
        method: "POST",
        headers: { "Content-type": "application/json" },
        body: body,
        redirect: "follow"
    });

    const text = await res.text();
    try {
        const data = JSON.parse(text);
        return data;
    } catch (err) {
        console.log("Error al parsear el JSON:", err);
        throw new Error("Respuesta no es JSON: " + text);
    };
};

// Funcion para obtener y dividir los nombres del campo NOMBRE(S) del xlsx en nombre1 y nombre2
const getName = nombre => {
    let nombres, nombre1, nombre2;
    nombres = nombre.trim().split(" ");
    if (nombres.length !== 1) {
        nombres = nombre.trim().split(" ");
        nombre1 = nombres.slice(0, nombres.length - 1).join(" ");
        nombre2 = nombres[nombres.length - 1];
    }
    else {
        nombre1 = nombres[0];
        nombre2 = "";
    }

    return {
        nombre1,
        nombre2
    };
};

// Funcion para obtener la fecha de nacimiento a partir de la CURP
const getBdate = curp => {
    if (curp.length < 10) return "";
    let yy = +curp[4] > 3 ? 1900 + (+curp[4] * 10) + +curp[5] : 2000 + (+curp[4] * 10) + +curp[5];
    let mm = curp[6] + curp[7];
    let dd = curp[8] + curp[9];
    let bdate = `${yy}-${mm}-${dd}`;
    return bdate;
};

// Funcion para obtener el genero a parir de la CURP
const getGenre = curp => {
    if (curp.length < 11) return 1;
    let genre = 1;
    if (curp[10].toUpperCase() != "H") genre = 2;
    return genre;
};

// Funcion para normalizar el numero de telefono
const getPhoneNumber = tel => {
    let ans = "";
    tel.trim();
    for (let i = 0; i < tel.length; i++) {
        if (tel.charCodeAt(i) >= 48 && tel.charCodeAt(i) <= 57) ans += tel[i];
    }
    return +ans;
};

// Creamos un mapa para tener un trackeo de los beneficiarios que ya se agregaron
const map = new Map();

for (const row of data) {
    const curp = row["CURP"].toUpperCase();
    map.set(curp, curp);
};
let cont = 0;
for (const row of data2) {
    const curp = row["CURP"].toUpperCase();
    if (!curp) continue;
    if (map.has(curp)) continue;
    let resultado = await buscarBeneficiario(curp);
    if (resultado.beneficiarios.length > 0) {
        row["RUB ID"] = resultado.beneficiarios[0].idBeneficiario;
        console.log(`Iteracion ${cont} Rub ID actualizado ${resultado.beneficiarios[0].idBeneficiario}`);
    }
    else {
        const nombre = getName(row["NOMBRE(S)"]);
        const fnacimiento = getBdate(row["CURP"])
        const tel = getPhoneNumber(row["TELEFONO"].toString());
        const genero = getGenre(row["CURP"])
        const nuevoBeneficiario = {
            nombre1: nombre.nombre1,
            nombre2: nombre.nombre2,
            apellidoPaterno: row["APELLIDO PATERNO"] || "-",
            apellidoMaterno: row["APELLIDO MATERNO"] || "-",
            fechaNacimiento: fnacimiento,
            direccion: row["DIRECCION"] || "",
            colonia: row["COLONIA"] || "",
            telefono: tel,
            idGenero: genero,
            curp: row["CURP"] || "",
        };
        const addBeneficiario = await agregarBeneficiario(nuevoBeneficiario);
        row["RUB ID"] = addBeneficiario.idBeneficiario;
        console.log(`Iteracion ${cont} beneficiario agregado con RUB ID ${addBeneficiario.idBeneficiario}`);
        console.log(addBeneficiario);


    }
    cont++;
    resultado = await buscarBeneficiario(curp);
    const date = new Date();
    let yy = date.getFullYear();
    let mm = date.getMonth() + 1;
    let dd = date.getDate() - 1;

    const prueba = {
        idBeneficiario: resultado.beneficiarios[0].idBeneficiario,
        idTipoApoyo: 51,
        fechaEntrega: `${yy}-${mm}-${dd}`, //YY-MM-DD
        idDependencia: "3",
        comentarios: "Agregado automaticamente desde excel por bot CACP"
    };
    agregarApoyo(prueba);
    console.log(`Apoyo agregado con rub ID ${resultado.beneficiarios[0].idBeneficiario}`);
};
