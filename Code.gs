const CONFIG = {
  NOMBRE: "ACUEDUCTO ACUAVIDA EL RECREO",
  NIT: "901.888.280-5"
};

function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Acueducto Acuavida El Recreo")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* =========================================================
   INICIALIZAR BASE DE DATOS
   ========================================================= */

function inicializarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  crearHoja_(ss, "CLIENTES", [
    "ID",
    "SUSCRIPTOR",
    "NOMBRE",
    "DIRECCION",
    "TELEFONO",
    "CUOTA_ACUEDUCTO",
    "ESTADO",
    "OBSERVACIONES",
    "FECHA_REGISTRO"
  ]);

  crearHoja_(ss, "FACTURAS", [
    "ID_FACTURA",
    "N_FACTURA",
    "SUSCRIPTOR",
    "NOMBRE",
    "FECHA",
    "FACTURA_ANTERIOR",
    "CUOTA_ACUEDUCTO",
    "MORA",
    "MORA_MESES",
    "MATRICULA",
    "RECONEXION",
    "ABONO",
    "OTROS",
    "TOTAL",
    "MESES",
    "ANOS",
    "PERIODO_DESDE",
    "PERIODO_HASTA",
    "PAGAR_ANTES",
    "FECHA_SUSPENSION",
    "ESTADO",
    "FECHA_CREACION"
  ]);

  crearHoja_(ss, "PAGOS", [
    "ID_PAGO",
    "N_FACTURA",
    "SUSCRIPTOR",
    "NOMBRE",
    "FECHA",
    "VALOR",
    "CONCEPTO",
    "OBSERVACIONES",
    "USUARIO"
  ]);

  crearHoja_(ss, "PERIODOS", [
    "ID",
    "SUSCRIPTOR",
    "ANO",
    "MES",
    "ESTADO",
    "N_FACTURA",
    "FECHA_PAGO"
  ]);

  crearHoja_(ss, "CONFIGURACION", [
    "PARAMETRO",
    "VALOR"
  ]);

  const config = ss.getSheetByName("CONFIGURACION");

  if (config.getLastRow() <= 1) {
    config.getRange(2, 1, 7, 2).setValues([
      ["NOMBRE_ACUEDUCTO", CONFIG.NOMBRE],
      ["NIT", CONFIG.NIT],
      ["CUOTA_DEFAULT", 25000],
      ["DIAS_PARA_PAGO", 10],
      ["DIAS_SUSPENSION", 11],
      ["TESORERO", ""],
      ["ULTIMA_FACTURA", 0]
    ]);
  }

  return "Sistema inicializado correctamente.";
}

function crearHoja_(ss, nombre, encabezados) {
  let hoja = ss.getSheetByName(nombre);

  if (!hoja) {
    hoja = ss.insertSheet(nombre);
  }

  if (hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    hoja.setFrozenRows(1);
  }
}

/* =========================================================
   CLIENTES
   ========================================================= */

function buscarClientes(texto) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CLIENTES");

  if (!hoja || hoja.getLastRow() < 2) {
    return [];
  }

  const datos = hoja.getDataRange().getValues();
  const busqueda = String(texto || "").toLowerCase().trim();

  const resultados = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];

    const id = String(fila[0] || "");
    const suscriptor = String(fila[1] || "");
    const nombre = String(fila[2] || "");
    const direccion = String(fila[3] || "");
    const telefono = String(fila[4] || "");

    const coincide =
      !busqueda ||
      id.toLowerCase().includes(busqueda) ||
      suscriptor.toLowerCase().includes(busqueda) ||
      nombre.toLowerCase().includes(busqueda) ||
      direccion.toLowerCase().includes(busqueda) ||
      telefono.toLowerCase().includes(busqueda);

    if (coincide) {
      resultados.push({
        id: id,
        suscriptor: suscriptor,
        nombre: nombre,
        direccion: direccion,
        telefono: telefono,
        cuota: Number(fila[5]) || 0,
        estado: fila[6] || "Activo",
        observaciones: fila[7] || ""
      });
    }
  }

  return resultados.slice(0, 50);
}

function obtenerCliente(suscriptor) {
  const resultados = buscarClientes(String(suscriptor));
  return resultados.find(
    c => String(c.suscriptor) === String(suscriptor)
  ) || null;
}

function guardarCliente(cliente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("CLIENTES");

  if (!hoja) {
    throw new Error("Primero debes ejecutar inicializarSistema().");
  }

  const suscriptor = String(cliente.suscriptor || "").trim();
  const nombre = String(cliente.nombre || "").trim();

  if (!suscriptor) {
    throw new Error("El número de suscriptor es obligatorio.");
  }

  if (!nombre) {
    throw new Error("El nombre del cliente es obligatorio.");
  }

  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][1]) === suscriptor) {
      hoja.getRange(i + 1, 3, 1, 7).setValues([[
        nombre,
        cliente.direccion || "",
        cliente.telefono || "",
        Number(cliente.cuota) || 0,
        cliente.estado || "Activo",
        cliente.observaciones || "",
        datos[i][8] || new Date()
      ]]);

      return {
        ok: true,
        mensaje: "Cliente actualizado."
      };
    }
  }

  const id = Utilities.getUuid();

  hoja.appendRow([
    id,
    suscriptor,
    nombre,
    cliente.direccion || "",
    cliente.telefono || "",
    Number(cliente.cuota) || 0,
    cliente.estado || "Activo",
    cliente.observaciones || "",
    new Date()
  ]);

  return {
    ok: true,
    mensaje: "Cliente creado correctamente."
  };
}

/* =========================================================
   FACTURAS
   ========================================================= */

function obtenerNumeroFactura() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("CONFIGURACION");

  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === "ULTIMA_FACTURA") {
      return Number(datos[i][1]) || 0;
    }
  }

  return 0;
}

function siguienteFactura() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("CONFIGURACION");

  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === "ULTIMA_FACTURA") {
      const nuevo = Number(datos[i][1]) + 1;
      hoja.getRange(i + 1, 2).setValue(nuevo);
      return nuevo;
    }
  }

  throw new Error("No existe configuración de facturación.");
}

function generarFactura(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("FACTURAS");

  if (!hoja) {
    throw new Error("Ejecuta primero inicializarSistema().");
  }

  const cliente = obtenerCliente(datos.suscriptor);

  if (!cliente) {
    throw new Error("No se encontró el suscriptor.");
  }

  const numero = siguienteFactura();

  const cuota = Number(datos.cuota) || 0;
  const mora = Number(datos.mora) || 0;
  const matricula = Number(datos.matricula) || 0;
  const reconexion = Number(datos.reconexion) || 0;
  const abono = Number(datos.abono) || 0;
  const otros = Number(datos.otros) || 0;

  const meses = datos.meses || [];
  const anos = datos.anos || [];

  const total =
    (cuota * meses.length) +
    mora +
    matricula +
    reconexion +
    otros -
    abono;

  const fecha = datos.fecha
    ? new Date(datos.fecha + "T12:00:00")
    : new Date();

  const facturaAnterior = obtenerUltimaFacturaCliente(
    datos.suscriptor
  );

  hoja.appendRow([
    Utilities.getUuid(),
    numero,
    cliente.suscriptor,
    cliente.nombre,
    fecha,
    facturaAnterior,
    cuota,
    mora,
    datos.moraMeses || "",
    matricula,
    reconexion,
    abono,
    otros,
    total,
    meses.join(", "),
    anos.join(", "),
    datos.periodoDesde || "",
    datos.periodoHasta || "",
    datos.pagarAntes || "",
    datos.fechaSuspension || "",
    "PENDIENTE",
    new Date()
  ]);

  registrarPeriodos_(cliente.suscriptor, numero, meses, anos);

  return {
    ok: true,
    numero: numero,
    facturaAnterior: facturaAnterior,
    total: total,
    nombreArchivo: crearNombreArchivoFactura_(cliente.suscriptor, cliente.nombre, numero),
    fecha: Utilities.formatDate(
      fecha,
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    )
  };
}

function crearNombreArchivoFactura_(suscriptor, nombre, numero) {
  const limpiar = texto => String(texto || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9_-]+/g, "_")
    .replace(/^_+|_+$/g, "");

  return "Factura_" +
    limpiar(suscriptor) +
    "_N-" +
    limpiar(numero) +
    "_" +
    limpiar(nombre) +
    ".pdf";
}

function obtenerUltimaFacturaCliente(suscriptor) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("FACTURAS");

  if (!hoja || hoja.getLastRow() < 2) {
    return "";
  }

  const datos = hoja.getDataRange().getValues();

  for (let i = datos.length - 1; i >= 1; i--) {
    if (String(datos[i][2]) === String(suscriptor)) {
      return datos[i][1];
    }
  }

  return "";
}

function registrarPeriodos_(suscriptor, factura, meses, anos) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("PERIODOS");

  if (!hoja) return;

  meses.forEach((mes, index) => {
    hoja.appendRow([
      Utilities.getUuid(),
      suscriptor,
      anos[index] || "",
      mes,
      "PENDIENTE",
      factura,
      ""
    ]);
  });
}

/* =========================================================
   HISTORIAL
   ========================================================= */

function obtenerHistorial(suscriptor) {

  const hoja =
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("FACTURAS");

  if (!hoja || hoja.getLastRow() < 2) {
    return [];
  }

  const datos =
    hoja.getDataRange().getValues();

  const resultado = [];

  for (let i = datos.length - 1; i >= 1; i--) {

    if (
      String(datos[i][2]) ===
      String(suscriptor)
    ) {

      let fecha = "";

      if (datos[i][4]) {

        const fechaObj =
          datos[i][4] instanceof Date
            ? datos[i][4]
            : new Date(datos[i][4]);

        if (!isNaN(fechaObj.getTime())) {

          fecha =
            Utilities.formatDate(
              fechaObj,
              Session.getScriptTimeZone(),
              "dd/MM/yyyy"
            );

        }
      }

      resultado.push({

        factura:
          datos[i][1],

        fecha:
          fecha,

        total:
          Number(datos[i][13]) || 0,

        estado:
          datos[i][20] || "PENDIENTE",

        meses:
          datos[i][14] || ""

      });

    }
  }

  return resultado.slice(0, 30);
}

/* =========================================================
   PAGOS
   ========================================================= */

function registrarPago(pago) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("PAGOS");

  if (!hoja) {
    throw new Error("No existe la hoja PAGOS.");
  }

  if (!pago.suscriptor) {
    throw new Error("Falta el suscriptor.");
  }

  const valor = Number(pago.valor) || 0;

  if (valor <= 0) {
    throw new Error("El valor del pago debe ser mayor que cero.");
  }

  hoja.appendRow([
    Utilities.getUuid(),
    pago.factura || "",
    pago.suscriptor,
    pago.nombre || "",
    new Date(),
    valor,
    pago.concepto || "Pago de servicio de agua",
    pago.observaciones || "",
    Session.getActiveUser().getEmail() || ""
  ]);

  if (pago.factura) {
    marcarFacturaPagada_(pago.factura);
  }

  return {
    ok: true,
    mensaje: "Pago registrado correctamente."
  };
}

/* =========================================================
   BUSCAR CLIENTES Y FACTURAS PENDIENTES PARA PAGOS
   ========================================================= */

function buscarClientesParaPago(texto) {
  const clientes = buscarClientes(texto);
  const hoja = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("FACTURAS");

  if (!hoja || hoja.getLastRow() < 2) {
    return clientes.map(c => ({
      cliente: c,
      facturas: []
    }));
  }

  const datos = hoja.getDataRange().getValues();

  return clientes.map(cliente => {

    const facturas = [];

    for (let i = datos.length - 1; i >= 1; i--) {

      const suscriptor = String(datos[i][2] || "");
      const estado = String(datos[i][20] || "PENDIENTE").toUpperCase();

      if (
        suscriptor === String(cliente.suscriptor) &&
        estado !== "PAGADA"
      ) {

        facturas.push({
          factura: datos[i][1],
          fecha: formatearFecha_(datos[i][4]),
          total: Number(datos[i][13]) || 0,
          meses: datos[i][14] || "",
          estado: estado
        });
      }
    }

    return {
      cliente: cliente,
      facturas: facturas
    };
  });
}


/* =========================================================
   OBTENER UNA FACTURA PARA REGISTRAR PAGO
   ========================================================= */

function obtenerFacturaParaPago(numeroFactura) {

  const hoja = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("FACTURAS");

  if (!hoja || hoja.getLastRow() < 2) {
    return null;
  }

  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {

    if (String(datos[i][1]) === String(numeroFactura)) {

      return {
        factura: datos[i][1],
        suscriptor: datos[i][2],
        nombre: datos[i][3],
        fecha: formatearFecha_(datos[i][4]),
        total: Number(datos[i][13]) || 0,
        meses: datos[i][14] || "",
        estado: datos[i][20] || "PENDIENTE"
      };
    }
  }

  return null;
}

function marcarFacturaPagada_(numeroFactura) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  /* =========================
     MARCAR FACTURA PAGADA
  ========================= */

  const hojaFacturas =
    ss.getSheetByName("FACTURAS");


  if (!hojaFacturas) return;


  const datos =
    hojaFacturas
      .getDataRange()
      .getValues();


  let suscriptor = "";


  for (
    let i = 1;
    i < datos.length;
    i++
  ) {

    if (
      String(datos[i][1]) ===
      String(numeroFactura)
    ) {

      hojaFacturas
        .getRange(i + 1, 21)
        .setValue("PAGADA");


      suscriptor =
        String(datos[i][2] || "");

      break;
    }
  }


  /* =========================
     MARCAR PERIODOS PAGADOS
  ========================= */

  const hojaPeriodos =
    ss.getSheetByName("PERIODOS");


  if (
    !hojaPeriodos ||
    hojaPeriodos.getLastRow() < 2
  ) {
    return;
  }


  const periodos =
    hojaPeriodos
      .getDataRange()
      .getValues();


  const fechaPago =
    new Date();


  for (
    let i = 1;
    i < periodos.length;
    i++
  ) {

    const facturaPeriodo =
      String(periodos[i][5] || "");

    const suscriptorPeriodo =
      String(periodos[i][1] || "");


    if (
      facturaPeriodo ===
      String(numeroFactura) &&

      suscriptorPeriodo ===
      suscriptor
    ) {

      // Columna 5 = ESTADO
      hojaPeriodos
        .getRange(i + 1, 5)
        .setValue("PAGADA");


      // Columna 7 = FECHA_PAGO
      hojaPeriodos
        .getRange(i + 1, 7)
        .setValue(fechaPago);

    }

  }

}

/* =========================================================
   PAGOS - BUSCAR CLIENTE Y FACTURAS PENDIENTES
   ========================================================= */

function buscarClientesParaPago(texto) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojaClientes = ss.getSheetByName("CLIENTES");
  const hojaFacturas = ss.getSheetByName("FACTURAS");

  if (!hojaClientes) {
    throw new Error("No existe la hoja CLIENTES.");
  }

  if (!hojaFacturas) {
    throw new Error("No existe la hoja FACTURAS.");
  }

  const busqueda =
    String(texto || "").toLowerCase().trim();

  if (!busqueda) {
    return [];
  }

  /* =========================
     BUSCAR CLIENTES
     ========================= */

  const clientesData =
    hojaClientes.getDataRange().getValues();

  const clientes = [];

  for (let i = 1; i < clientesData.length; i++) {

    const fila = clientesData[i];

    const id =
      String(fila[0] || "");

    const suscriptor =
      String(fila[1] || "");

    const nombre =
      String(fila[2] || "");

    const direccion =
      String(fila[3] || "");

    const telefono =
      String(fila[4] || "");

    const coincide =
      id.toLowerCase().includes(busqueda) ||
      suscriptor.toLowerCase().includes(busqueda) ||
      nombre.toLowerCase().includes(busqueda) ||
      direccion.toLowerCase().includes(busqueda) ||
      telefono.toLowerCase().includes(busqueda);

    if (coincide) {

      clientes.push({
        id: id,
        suscriptor: suscriptor,
        nombre: nombre,
        direccion: direccion,
        telefono: telefono,
        cuota: Number(fila[5]) || 0,
        estado: fila[6] || "Activo",
        observaciones: fila[7] || ""
      });

    }

  }


  /* =========================
     BUSCAR FACTURAS
     ========================= */

  const facturasData =
    hojaFacturas.getDataRange().getValues();


  return clientes.map(cliente => {

    const facturas = [];

    for (
      let i = facturasData.length - 1;
      i >= 1;
      i--
    ) {

      const fila =
        facturasData[i];

      const suscriptorFactura =
        String(fila[2] || "");

      const estado =
        String(fila[20] || "PENDIENTE")
          .toUpperCase();

      if (
        suscriptorFactura ===
        String(cliente.suscriptor) &&

        estado !== "PAGADA"
      ) {

        let fecha = "";

        if (fila[4]) {

          const fechaObj =
            fila[4] instanceof Date
              ? fila[4]
              : new Date(fila[4]);

          if (
            !isNaN(
              fechaObj.getTime()
            )
          ) {

            fecha =
              Utilities.formatDate(
                fechaObj,
                Session.getScriptTimeZone(),
                "dd/MM/yyyy"
              );

          }

        }


        facturas.push({

          factura:
            fila[1],

          fecha:
            fecha,

          total:
            Number(fila[13]) || 0,

          meses:
            fila[14] || "",

          estado:
            estado

        });

      }

    }


    return {

      cliente:
        cliente,

      facturas:
        facturas

    };

  });

}