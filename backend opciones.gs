/**
 * GOOGLE APPS SCRIPT - BACKEND PARA DASHBOARD STM
 * Copiar este contenido en un nuevo proyecto de Google Apps Script.
 */

const SPREADSHEET_ID = '16Y723omF3l38Ntq0MUSh-ZMYnRIvYmctrscidox5ktc';

/**
 * Mapeo de pestañas
 */
const SHEETS = {
  CONVENIOS: 'tabla_convenios',
  NOVEDADES: 'novedades',
  USUARIOS: 'user',
  MENSAJES: 'tabla_mensaje_user',
  CAPACITACION: 'tabla_capacitacion'
};

function doGet(e) {
  const result = {
    CONVENIOS: getSheetData(SHEETS.CONVENIOS),
    NOVEDADES: getSheetData(SHEETS.NOVEDADES),
    USUARIOS: getSheetData(SHEETS.USUARIOS),
    MENSAJES: getSheetData(SHEETS.MENSAJES),
    CAPACITACIONES: getSheetData(SHEETS.CAPACITACION)
  };
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let requestData;
  try {
    requestData = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'ERROR', message: 'Invalid JSON' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const { action, sheet, id, data } = requestData;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  if (action === 'saveConsultation') {
    return saveConsultation(requestData);
  }

  const targetSheet = ss.getSheetByName(SHEETS[sheet]);

  if (!targetSheet) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'ERROR', message: 'Sheet not found' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    switch (action) {
      case 'CREATE':
        return createRecord(targetSheet, data);
      case 'UPDATE':
        return updateRecord(targetSheet, id, data);
      case 'DELETE':
        return deleteRecord(targetSheet, id);
      default:
        throw new Error('Action not supported');
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'ERROR', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Guarda una nueva consulta de capacitación
 */
function saveConsultation(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.CAPACITACION);
  if (!sheet) throw new Error("No se encontró la tabla de capacitaciones");
  
  const nextId = getNextId(sheet);
  const now = new Date();
  const fechaStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  
  // Obtenemos encabezados para mapear columnas dinámicamente
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().toLowerCase().trim());
  
  const newRow = headers.map(h => {
    if (h === 'id') return nextId;
    if (h === 'id_user') return data.id_user;
    if (h === 'capacitacion') return data.capacitacion;
    if (h === 'fecha_sol') return fechaStr;
    return '';
  });
  
  sheet.appendRow(newRow);
  return createJsonResponse({ status: 'SUCCESS', id: nextId });
}

/**
 * Obtiene todos los datos de una pestaña como Array de Objetos
 */
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      if (header) {
        obj[header.toLowerCase().trim()] = row[index];
      }
    });
    // Normalización de campo "activo" si no existe
    if (obj.activo === undefined) obj.activo = 'SI';
    return obj;
  });
}

function createRecord(sheet, data) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const nextId = getNextId(sheet);
  const sheetNameActual = sheet.getName().toLowerCase().trim();
  
  console.log("Creando registro en: " + sheetNameActual);
  console.log("Data recibida: " + JSON.stringify(data));

  const newRow = headers.map(header => {
    const key = header.toString().toLowerCase().trim();
    if (key === 'id' || key === 'id_mensaje') return nextId;
    return data[key] !== undefined ? data[key] : '';
  });
  
  sheet.appendRow(newRow);

  let emailStatus = "N/A";
  // AUTOMATIZACIÓN: Enviar email si es un nuevo usuario con correo (SOLO EN CREATE)
  if (sheetNameActual === SHEETS.USUARIOS && data.email) {
    console.log("Iniciando envío de email a: " + data.email);
    try {
      enviarEmailBienvenida(data);
      emailStatus = "SENT";
      console.log("Email enviado exitosamente");
    } catch (e) {
      emailStatus = "ERROR: " + e.toString();
      console.error("Fallo al enviar email: " + e.toString());
    }
  }

  return createJsonResponse({ status: 'SUCCESS', id: nextId, email_status: emailStatus });
}

// Función para PROBAR el envío de email desde el editor (Ejecutar manualmente aquí)
function testWelcomeEmail() {
  const testData = {
    nombre: "Usuario de Prueba",
    usuario: "testuser",
    password: "password123",
    email: "TU_CORREO_AQUÍ@gmail.com" // cambia esto para probar
  };
  console.log("Iniciando prueba manual...");
  enviarEmailBienvenida(testData);
  console.log("Prueba finalizada. Revisa tu bandeja de entrada.");
}

function enviarEmailBienvenida(u) {
  const nombreApp = "STM App";
  const asunto = `¡Bienvenido/a a ${nombreApp}! - Tus credenciales de acceso`;
  
  // Logo en Base64 para incrustar directamente
  const logoBase64 = "iVBORw0KGgoAAAANSUhEUgAABDgAAAVGCAYAAAB7TwuKAAAQAElEQVR4AeydB4AkVbX+v1tV3T05bA4sLDmIoKJixmdGUDE8/z5zQEUMDzMmwKcoCogSFQOCAgqIIBJMgJJzziwLy+YwOzl0qPp/53bXbM/sbGTDhK+2Tt187rm/qu7tc7q6JoA2ERABERABERABERABERABERABERCB8U5g3K9PAY5xf4q1QBEQAREQAREQAREQAREQAREQgQ0TUI+xTkABjrF+BmW/CIiACIiACIiACIiACIiACGwLAppDBEY5AQU4RvkJknkiIAIiIAIiIAIiIAIiIAJjg4CsFAER2L4EFODYvvw1uwiIgAiIgAiIgAiIgAhMFAJapwiIgAhsVQIKcGxVvFIuAiIgAiIgAiIgAiIgAhtLQP1EQAREQASeCwEFOJ4LPY0VAREQAREQAREQARHYdgQ0kwiIgAiIgAish4ACHOuBoyYREAEREAEREAERGEsEZKsIiIAIiIAITGQCCnBM5LOvtYuACIiACIjAxCKg1YqACIiACIiACIxjAgpwjOOTq6WJgAiIgAiIwKYRUG8REAEREAEREAERGLsEFOAYu+dOlouACIiACGxrAppPBERABERABERABERg1BJQgGPUnhoZJgIiIAJjj4AsFgEREAEREAEREAEREIHtRUABju1FXvOKgAhMRAJaswiIgAiIgAiIgAiIgAiIwFYioADHVgIrtSIgAptDQGNEQAREQAREQAREQAREQAREYPMIKMCxedw0SgS2DwHNKgIiIAIiIAIiIAIiIAIiIAIiMCIBBThGxKLKsUpAdouACIiACIiACIiACIiACIiACExMAgpwTKZZrdWKgAiIgAiIgAiIgAiIgAiIgAiIwLgkoADHkNOqggiIgAiIgAiIgAiIgAiIgAiIgAiIwFglILtFYGISUIBjYp53rVoEREAEREAEREAEREAEJi4BrVwERGBcElCAY1yeVi1KBERABERABERABERABDafgEaKgAiIwFgkoADHWDxrslkEREAEREAEREAERGB7EtDcIiACIiACo5CAAhyj8KTIJBEQAREQAREQAREY2wRkvQiIgAiIgAhsewIKcGx75ppRBERABERABERgohPQ+kVABERABERABLY4AQU4tjhSKRQBERABERABEXiuBDReBERABERABERABDaVgAIcm0pM/UVABERABERg+xOQBSIgAiIgAiIgAiIgAsMIKMAxDIiKIiACIiAC44GA1iACIiACIiACIiACIjDRCCjAMdHOuNYrAiIgAkZAIgIiIAIiIAIiIAIiIALjjIACHOPshGo5IiACW4aAtIiACIiACIiACIiACIiACIwtAgpwjK2zJWtFYLQQkB0iIAIiIAIiIAIiIAIiIAIiMKoIKMAxqk6HjBk/BLQSERABERABERABERABERABERCBbUlAAY5tSVtzrSGgnAiIgAiIgAiIgAiIgAiIgAiIgAhsQQIKcGxBmFtSlXSJgAiIgAiIgAiIgAiIgAiIgAiIgAhsPIGxGuDY+BWqpwiIgAiIgAiIgAiIgAiIgAiIgAiIwFglsNF2K8Cx0ajUUQREQAREQAREQAREQAREQAREQARGGwHZkxJQgCMloVQEREAEREAEREAEREAEREAERGD8EdCKJgwBBTgmzKnWQkVABERABERABERABERABERgbQLV+mEYHwQkABjvFyJrUOERABERABERABERABERCBrUFAOkVABMYIAQU4xsiJkpkiIAIiIAIiIAIiIAIiMDoJyCoREAERGB0EFOAYHedBVoiACIiACIiACIiACIxXAlqXCIiACIjANiGgAMc2waxJREAEREAEREAEREAE1kVA9SIgAiIgAiKwJQgowLElKEqHCIiACIiACIiACGw9AtIsAiIgAiIgAiKwEQQU4NgISOoiAiIgAiIgAiIwmgnINhEQAREQAREQAREAFODQVSACIiACIiAC452A1icCIiACIiACIiACE4CAAhwT4CRriSIgAiIgAusnoFYREAEREAEREAEREIGxT0ABjrF/DrUCERABEdjaBKRfBERABERABERABERABEY9AQU4Rv0pkoEiIAKjn4AsFAEREAEREAEREAEREAER2N4EFODY3mdA84vARCCgNYqACIiACIiACIiACIiACIjAViagAMdWBiz1IrAxBNRHBERABERABERABERABERABETguRFQgOO58dPobUNAs4iACIiACIiACIiACIiACIiACIjAegkowLFePGOlUXaKgAiIgAiIgAiIgAiIgAiIgAiIwMQmMDECHBP7HGv1IiACIiACIiACIiACIiACIiACIjDuCfgAx7hfpRYoAiIgAiIgAiIgAiIgAiIgAiIgAiKAcYpAAY5xemK1LBEQAREQAREQAREQAREQAREQgc0joFFjk4ACHGPzvMlqERABERABERABERABERABEdheBDSvCIxKAgpwjMrTIqNEQAREQAREQAREQAREQATGLgFZLgIisD0IKMCxPahrThEQAREQAREQAREQARGYyAS0dhEQARHYCgQU4NgKUKVSBERABERABERABERABJ4LAY0VAREQARHYdAIKcGw6M40QAREQAREQAREQARHYvgQ0uwiIgAiIgAisRUABjrWQqEIEREAEREAEREAExjoB2S8CIiACIiACE4+AAhwT75xrxSIgAiIgAiIgAiIgAiIgAiIgAiIw7ggowDHuTqkWJAIiIAIiIALPnYA0iIAIiIAIiIAIiMBYI6AAx1g7Y7JXBERABERgNBCQDSIgAiIgAiIgAiIgAqOMgAIco+yEyBwREAERGB8EtAoREAEREAEREAEREAER2LYEFODYtrw1mwiIgAiUCegoAiIgAiIgAiIgAiIgAiKwRQkowLFFcUqZCIjAliIgPSIgAiIgAiIgAiIgAiIgAiKwKQQU4NgUWuorAqOHgCwRAREQAREQAREQAREQAREQARGoIqAARxUMZccTAa1FBERABERgexLYf//9V++/HtmetmluERABERABERCB8UlAAY7xeV43vCr1EAEREAERmJAE9t5771P23XffP3/kIx9JDj/88ORzn/tccvTRRyff//73k1NOOSU566yzkl/96lfJv//97+TGG29Mbr/99uS+++5LHn300WT+/PnJokWLkqVLlyYb2u69996We9cjGxpv7cViMenp6UlWrFiRPPPMM94Gs+Wuu+5Kzj///OS8887ztp5xxhnJSSedlBx//PHJsccem3z7299ODjnkkOSNb3xj8qpXvSqxQMuEPNlatAiIgAiIgAhMMAIKcKzjhKtaBERABERABEY7gT322OMj5sC/973vTb74xS8mJ554YnLuuecmf/rTn5J77rknefjhh31gYPXq1UmpVEpsY91RDzzwwGG//e1v8ctf/hKnnXYafvjDH+Jb3/oWjjrqKBxxxBH4xCc+gde85jV45StfiZe85CXYb7/9sOeee2Lu3LmYNWsWpk+fjm2xhWGIuro6TJkyBTvuuKO3wWx50YtehPe///340Ic+5G098sgj8eUvfxnf/OY3cdxxx+F73/ee/vrXv+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630X6v63++8m/+8of//73qS/eB576Y6TfS9L/3pP+uYv+jD/NHzS+P0/+e99R3+ve66L+m79Dvt903v889YV7t87S9Pr/R/pBvS/8+1b/+9U96PfqmS/q978i+p9F+u7Xv6e8L987f2eG6D/mH+pPf6z6f9B8r79L0pn/2XfS343XnU6nX+9A36P/tP+4rv//pB/kH+pP/0H++zS88u9Hf7veB5v/u+fT1P9O/rYVfoX9Y6T3R/YfR+7pZ97T///S6e/D6U+V0idKvzO/L9+Xpffp5+m590TnvqO+d38P9Wfw77z+3itI09/p96j8D0y9R9/P6u+T/r6p/vunv097Oen+E+pXv9eeUP29+T6R567fXz2O+l6Q35PvuH/r78u3353v/6R6VpX6v/+/0/f6+7ffv/58un2fS98r/x7T/2Xp/5Nrf+L/N9Lfv379f871+2K9T7v/Tz2vz9O9p/ov6O+k/t9K/+vp/850XfG7Sf93pP9NrnYf0N/nUv9ff8X9nZPP8f+X/+Z/p87p66L+b/K/W//v/fGvI/+m0f+9un3zD/Un/X0u/bfp98O9v0Oat3X28vff6e/I9f+l976S3u9Kvw//X/P94F+v7k8//z++Yf6P+kL9Sfpf+U9Iv+M0rN9z3/vUn9Dvm37fXPeDzv/dMvS/u306X92DPkfP1e9df670u6LfS7+f/v6999L35fef68/Vv2u676L/mP/fP/A/6Xedv49S74v/Wfgv6p+reY5+D/V7pYyS9OdfD6e+V/UvVfo9rvv76ZfO/8G/l6TvmP/fE6r/v0S/D/97S3r//P0X/f2p/99S/z7531X6uyF9f/wX/U79+0o/i+7nkj/H/w79LPrf/7+u/y975v9/eL/Q59zvp/736PfNf0/8936vS39HfVpPr3M/rv75pX7/mvtO/f+S/l7Un/V/7+9Xf6be10q/m3yZ/A7U98V/N9e9P0///4B3609Y9O999AAAAABJRU5ErkJggg==";
  
  const logoBlob = Utilities.newBlob(Utilities.base64Decode(logoBase64), "image/png", "logo_stm.png");
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #334155; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden;">
      <div style="background: #4f46e5; padding: 25px; text-align: center;">
        <img src="cid:logo_stm" alt="STM Logo" style="width: 80px; height: auto; margin-bottom: 10px;">
        <h1 style="color: white; margin: 0; font-size: 24px;">¡Bienvenido/a, ${u.nombre}!</h1>
      </div>
      <div style="padding: 30px;">
        <p>Le damos la bienvenida a la aplicación oficial del <b>Sindicato de Trabajadores Municipales</b>. Desde ahora, podrá acceder a toda nuestra información y servicios desde su dispositivo.</p>
        
        <div style="background: #f8fafc; border-radius: 8px; padding: 20px; margin: 25px 0; border-left: 4px solid #4f46e5;">
          <p style="margin-top: 0; color: #4f46e5; font-weight: bold; text-transform: uppercase; font-size: 12px;">Tus datos de acceso:</p>
          <p style="margin: 5px 0;"><b>Usuario:</b> ${u.usuario || u.user}</p>
          <p style="margin: 5px 0;"><b>Contraseña temporal:</b> ${u.password}</p>
        </div>

        <p style="font-size: 14px; color: #64748b;"><b>Recomendación de seguridad:</b> Por su tranquilidad, le recomendamos cambiar esta contraseña inicial. Puede hacerlo desde el menú <b>"Mi Cuenta"</b> (icono de usuario) ubicado en la barra inferior de la aplicación.</p>
        
        <div style="text-align: center; margin-top: 35px;">
          <a href="#" style="background: #4f46e5; color: white; padding: 15px 30px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block;">Acceder a la App</a>
        </div>
      </div>
      <div style="background: #f1f5f9; padding: 20px; text-align: center; font-size: 12px; color: #94a3b8;">
        © 2026 Sindicato de Trabajadores Municipales. Todos los derechos reservados.
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: u.email,
    subject: asunto,
    htmlBody: htmlBody,
    inlineImages: {
      logo_stm: logoBlob
    }
  });
}

function updateRecord(sheet, id, data) {
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  let rowIndex = -1;
  let rowData = null;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === id.toString()) {
      rowIndex = i + 1;
      rowData = rows[i];
      break;
    }
  }

  if (rowIndex === -1) throw new Error('Record not found');

  headers.forEach((header, index) => {
    const key = header.toString().toLowerCase().trim();
    if (key !== 'id' && data[key] !== undefined) {
      sheet.getRange(rowIndex, index + 1).setValue(data[key]);
    }
  });

  // AUTOMATIZACIÓN: Enviar email y registrar mensaje si se actualiza la contraseña en tabla de usuarios
  let emailStatus = "N/A";
  let messageStatus = "N/A";
  const sheetName = sheet.getName().toLowerCase().trim();
  if (sheetName === SHEETS.USUARIOS && data.password) {
    // Extraer datos del usuario desde la fila encontrada
    const userData = {};
    headers.forEach((header, index) => {
      const key = header.toString().toLowerCase().trim();
      userData[key] = rowData[index];
    });

    // 1. Enviar email de notificación
    if (userData.email) {
      try {
        enviarEmailCambioPassword({
          nombre: userData.nombre || userData.name || 'Afiliado/a',
          usuario: userData.usuario || userData.user || '',
          email: userData.email,
          newPassword: data.password
        });
        emailStatus = "SENT";
        console.log("Email de cambio de contraseña enviado a: " + userData.email);
      } catch (e) {
        emailStatus = "ERROR: " + e.toString();
        console.error("Fallo al enviar email de cambio de contraseña: " + e.toString());
      }
    }

    // 2. Registrar mensaje en tabla_mensaje_user para que aparezca la campanita
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const msgSheet = ss.getSheetByName(SHEETS.MENSAJES);
      if (msgSheet) {
        const nextMsgId = getNextId(msgSheet);
        const now = new Date();
        const fechaStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        msgSheet.appendRow([
          nextMsgId,
          id,  // id_user del usuario que cambió la contraseña
          "🔐 Contraseña Modificada",
          "Su contraseña ha sido actualizada exitosamente el " + fechaStr + ". Si usted no realizó este cambio, comuníquese urgentemente con la Secretaría de Acción Social al 472 9815 / 472 3756 int. 137."
        ]);
        messageStatus = "CREATED";
        console.log("Mensaje de cambio de contraseña creado para usuario ID: " + id);
      }
    } catch (e) {
      messageStatus = "ERROR: " + e.toString();
      console.error("Fallo al crear mensaje de cambio de contraseña: " + e.toString());
    }
  }

  return createJsonResponse({ status: 'SUCCESS', email_status: emailStatus, message_status: messageStatus });
}

function deleteRecord(sheet, id) {
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === id.toString()) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error('Record not found');

  sheet.deleteRow(rowIndex);
  return createJsonResponse({ status: 'SUCCESS' });
}

function getNextId(sheet) {
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return 1;
  const ids = rows.slice(1)
    .map(r => parseFloat(r[0]))
    .filter(id => !isNaN(id));
  
  if (ids.length === 0) return 1;
  return Math.max(...ids) + 1;
}

function enviarEmailCambioPassword(u) {
  const nombreApp = "STM App";
  const asunto = `🔐 Contraseña Modificada - ${nombreApp}`;
  
  // Mismo logo Base64 que el email de bienvenida
  const logoBase64 = "iVBORw0KGgoAAAANSUhEUgAABDgAAAVGCAYAAAB7TwuKAAAQAElEQVR4AeydB4AkVbX+v1tV3T05bA4sLDmIoKJixmdGUDE8/z5zQEUMDzMmwKcoCogSFQOCAgqIIBJMgJJzziwLy+YwOzl0qPp/53bXbM/sbGTDhK+2Tt287rm/qu7tc7q6JoA2ERABERABERABERABERABERABERCB8U5g3K9PAY5xf4q1QBEQAREQAREQAREQAREQAREQgQ0TUI+xTkABjrF+BmW/CIiACIiACIiACIiACIiACGwLAppDBEY5AQU4RvkJknkiIAIiIAIiIAIiIAIiIAJjg4CsFAER2L4EFODYvvw1uwiIgAiIgAiIgAiIgAhMFAJapwiIgAhsVQIKcGxVvFIuAiIgAiIgAiIgAiIgAhtLQP1EQAREQASeCwEFOJ4LPY0VAREQAREQAREQAREQgc0joFFjk4ACHGPzvMlqERABERABERABERABERCBbUdAs4iACIxKAgpwjMrTIqNEQAREQAREQAREQAREQATGLgFZLgIisD0IKMCxPahrThEQAREQAREQAREQARGYyAS0dhEQARHYCgQU4NgKUKVSBERABERABERABERABJ4LAY0VAREQARHYdAIKcGw6M40QAREQAREQAREQARHYvgQ0uwiIgAiIgAisRUABjrWQqEIEREAEREAEREAExjoB2S8CIiACIiACE4+AAhwT75xrxSIgAiIgAiIgAiIgAiIgAiIgAiIw7ggowDHuTqkWJAIiIAIiIALPnYA0iIAIiIAIiIAIiMBYI6AAx1g7Y7JXBERABERgNBCQDSIgAiIgAiIgAiIgAqOMgAIco+yEyBwREAERGB8EtAoREAEREAEREAEREAER2LYEFODYtrw1mwiIgAiUCegoAiIgAiIgAiIgAiIgAiKwRQkowLFFcUqZCIjAliIgPSIgAiIgAiIgAiIgAiIgAiKwKQQU4NgUWuorAqOHgCwRAREQAREQAREQAREQAREQARGoIqAARxUMZccTAa1FBERABERABERABERABERABLYkAQU4tiRN6RIBERCB7U1A84uACIiACIiACIiACIjAsAIKcGxVvFIuAiIgAiIgAiIgAiIgAhtLQP1EQAREQASeCwEFOJ4LPY0VAREQAREQAREQARHYdgQ0kwiIgAiIgAish4ACHOuBoyYREAEREAEREAERGEsEZKsIiIAIiIAITGQCCnBM5LOvtYuACIiACIjAxCKg1YqACIiACIiACIxjAgpwjOOTq6WJgAiIgAiIwKYRUG8REAEREAEREAERGLsEFOAYu+dOlouACIiACGxrAppPBERABERABERABERg1BJQgGPUnhoZJgIiIAJjj4AsFgEREAEREAEREAEREIHtRUABju1FXvOKgAhMRAJaswiIgAiIgAiIgAiIgAiIwFYioADHVgIrtSIgAptDQGNEQAREQAREQAREQAREQAREYPMIKMCxedw0SgS2DwHNKgIiIAIiIAIiIAIiIAIiIAIiMCIBBThGxKLKsUpAdouACIiACIiACIiACIiACIiACExMAgpwTKZZrdWKgAiIgAiIgAiIgAiIgAiIgAiIwLgkoADHkNOqggiIgAiIgAiIgAiIgAiIgAiIgAiIwFglILtFYGISUIBjYp53rVoEREAEREAEREAEREAEJi4BrVwERGBcElCAY1yeVi1KBERABERABERABERABDafgEaKgAiIwFgkoADHWDxrslkEREAEREAEREAERGB7EtDcIiACIiACo5CAAhyj8KTIJBEQAREQAREQAREY2wRkvQiIgAiIgAhsewIKcGx75ppRBERABERABERgohPQ+kVABERABERABLY4AQU4tjhSKRQBERABERABEXiuBDReBERABERABERABDaVgAIcm0pM/UVABERABERg+xOQBSIgAiIgAiIgAiIgAsMIKMAxDIiKIiACIiAC44GA1iACIiACIiACIiACIjDRCCjAMdHOuNYrAiIgAkZAIgIiIAIiIAIiIAIiIALjjIACHOPshGo5IiACW4aAtIiACIiACIiACIiACIiACIwtAgpwjK2zJWtFYLQQkB0iIAIiIAIiIAIiIAIiIAIiMKoIKMAxqk6HjBk/BLQSERABERABERABERABERABERCBbUlAAY5tSVtzrSGgnAiIgAiIgAiIgAiIgAiIgAiIgAhsQQIKcGxBmFtSlXSJgAiIgAiIgAiIgAiIgAiIgAiIgAhsPIGxGuDY+BWqpwiIgAiIgAiIgAiIgAiIgAiIgAiIwFglsNF2K8Cx0ajUUQREQAREQAREQAREQAREQAREQARGGwHZkxJQgCMloVQEREAEREAEREAEREAEREAERGD8EdCKJgwBBTgmzKnWQkVABERABERABERABERABERgbQLV+mEYHwQkABjvFyJrUOERABERABERABERABERCBrUFAOkVABMYIAQU4xsiJkpkiIAIiIAIiIAIiIAIiMDoJyCoREAERGB0EFOAYHedBVoiACIiACIiACIiACIxXAlqXCIiACIjANiGgAMc2waxJREAEREAEREAEREAE1kVA9SIgAiIgAiKwJQgowLElKEqHCIiACIiACIiACGw9AtIsAiIgAiIgAiKwEQQU4NgISOoiAiIgAiIgAiIwmgnINhEQAREQAREQAREAAFCAQ1eBCIiACIiACIiACIiACIiACIjA+CegFU0YAgpwTJhTrYWKgAiIgAissZWdI0ABERCB8U5AKxQBEdggAfskvMGO6iACIiACIiACIiACo4CA5hABERABERABERCB8U5AAY7xfoa1PhEQAREQAREQAREQAREQgXUTUIsIiMB4JqAAx3g+u1qbCIiACIiACIiACIiACIiACIiACIjAIAEFOAZRKCMCIiACIiACIrBtCcgu0REBERABERCB8UxAAY7xfHa1NhEQAREQAREQAREQAREQga1BQDpFQATGCAEFOMbIiZKZIiACIiACIrBpBNRbBERABERABERABMYuAQU4xu65k+UiIAIiIALbmoDmE4EJQCA9v+vKrGuqMT5gQi1RixUBEdh+BPRJYvud/7E8u30aHMvrl+0iIAIiIAIiIAIiIAIiIAJjm4CsV4Bjm513WS8CIiACIiACIiACIiACIiACE5+A1jYhCSjAMSFPuxYtAiIgAiIgAiIgAiIgAiIgAiIgAuOLgAIc4+t8ajUiIAIiIAIiIAMEEREQAREQAREQAREYRkABjmFAVBQBERABERABERiv/7+9+w2VvKzjOP4+5+yhzYQkSbYH8sBlKAi0SnsMRQ95IBk+svxTUIJ/CipP+gBCwqKChLCwIjMoSjQq08Qs0vIvS/yTRkaiGIkVKv6JqGh/NJ3H/D57zz+/O7Oze2Pd5/3eeM+emZ0z3e/1IqJRBERERERABPKNgAOcuTr7WqVIiACIiACIiACIiACIiACIiACIiACW5eAAx5bdvVq9CIiACIiACIiACIiACIjARCLg6kVABFaRgAOcq0hWlkRABERABERg7RGQriIgAiIgAiIgApsRAQc4NqPdqrWIgAiIgAiIQCcBRSoCIiACIiACIiAClwIOcFwKyC0iIAIiIAIiIAIisHoEFF0EREAERCCPgAMc+XqrWgREQAREQATyEXAEERABERABERABEdgaCDjAsTV8SlqjCIiACIiACIiACIiACEw0An4NERER2CQJOMCxSe4WLUoEREAEREAERPIIuHoREAEREAER2PoJOMCx9X+GXqEIiIAIiIAIiIAIiIAIjBIBj56+BJa7xnqTIOAAxyaxG7QIERhTAqIhAiIgAiIgAiIgAiKQQcABjgydqpZLwHeLgAiIgAiIgAiIwOoRkHQREIFxJuAAx6a8e7U2ERABERABERCBiUzAuxUBERCBzYGAAxybw27VGkVABERABERg0yTg2EVABERABERA0WQTcIBjs9m1WqgIiIAIiIAIiIAIiIAIbB4E/BqtQgS2XAIOcGy5+14rFwEREAEREAEREAERGJ8E/BqtQgREYIwTcIBjjO8gmS8CIiACIiACIiACIiACq0jAkkVABEYzAQc4RvPukV0iIAIiIAIiIAIjJCDpIiACIiACIiACInApILsLAscKJjkHODYju1VLFQEREAEREAEREAEREAERGCUCXiB51BFYGAEHOBZGMR8VAREQAREQAREQAREQARGYmARcogiIwBgn4ADHGNc5slsEREAEREAEREAEREAEVpOAoouACIjA2CTgAMfY3H+yXgREQAREQAREQAREQAREYDQQuB4xswh4DSJgsrIEHOBYWeKqJwIiIAIiIAIiMHkJ+ioiIAIiIAIisI4EHOBYR8LKFgEREAEREAEREAEREAERGPcE3KQlsEkScIBjk9wtWpQIiIAIiIAIiIAIiIAIiMAq82W+CIjAOCHgAMc42UVahgiIgAiIgAjkEnB0ERABERABERCBzYCAAxybwS7VEkVABERABERgAhNwiSIgAiIgAiIgAiKwJRNwgGNL3vtatwiIgAiIgAhMYAIOcUkiIAIiIAIiIAIisEkQcIBjk9gNWoQIiIAIiIAIrC8CkiMCIiACIiACIiACIsD/BK4HAgpwrAcIDkJABERABERABERABERABERg9AvsbJmfHoAtZxWjmoADHKP6ALQ5IiACIiACIiACIiACIiACKyUgOqtUw4l7JaAARy8evUcEREAEREAEtiYC3o0IiIAIiIAIiIAILEJAAY5F8OhNIiACIiACIiACIrAGBGRFBERABERABERABDZLAg5wbJa7VYsWAREQAREQgQlOwCWKgAiIgAiIgAiIwJZIwAGOLXGva80iIAIiIAIiIAIiIAIiMC4IuJb1KCMCs0DAK1UhAAAAAABJRU5ErkJggg==";
  
  const logoBlob = Utilities.newBlob(Utilities.base64Decode(logoBase64), "image/png", "logo_stm.png");
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #334155; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden;">
      <div style="background: #4f46e5; padding: 25px; text-align: center;">
        <img src="cid:logo_stm" alt="STM Logo" style="width: 80px; height: auto; margin-bottom: 10px;">
        <h1 style="color: white; margin: 0; font-size: 24px;">🔐 Contraseña Modificada</h1>
      </div>
      <div style="padding: 30px;">
        <p>Hola <b>${u.nombre}</b>,</p>
        <p>Le informamos que su contraseña de acceso a la aplicación del <b>Sindicato de Trabajadores Municipales</b> ha sido modificada exitosamente.</p>
        
        <div style="background: #f8fafc; border-radius: 8px; padding: 20px; margin: 25px 0; border-left: 4px solid #10b981;">
          <p style="margin-top: 0; color: #10b981; font-weight: bold; text-transform: uppercase; font-size: 12px;">Datos actualizados:</p>
          <p style="margin: 5px 0;"><b>Usuario:</b> ${u.usuario}</p>
          <p style="margin: 5px 0;"><b>Nueva Contraseña:</b> ${u.newPassword}</p>
        </div>

        <div style="background: #fef2f2; border-radius: 8px; padding: 15px; margin: 20px 0; border-left: 4px solid #ef4444;">
          <p style="margin: 0; color: #991b1b; font-size: 13px;"><b>⚠️ Importante:</b> Si usted no realizó este cambio, por favor comuníquese de forma urgente con la Secretaría de Acción Social al <b>472 9815 / 472 3756 interno 137</b>.</p>
        </div>

        <p style="font-size: 14px; color: #64748b;"><b>Recomendación de seguridad:</b> No comparta su contraseña con terceros. Mantenga sus credenciales en un lugar seguro.</p>
        
        <div style="text-align: center; margin-top: 35px;">
          <a href="#" style="background: #4f46e5; color: white; padding: 15px 30px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block;">Acceder a la App</a>
        </div>
      </div>
      <div style="background: #f1f5f9; padding: 20px; text-align: center; font-size: 12px; color: #94a3b8;">
        © 2026 Sindicato de Trabajadores Municipales. Todos los derechos reservados.
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: u.email,
    subject: asunto,
    htmlBody: htmlBody,
    inlineImages: {
      logo_stm: logoBlob
    }
  });
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
