/**
 * MediParts Contact Form Handler v2.0
 * Google Apps Script para procesar formularios de contacto, guardar en Google Sheets,
 * almacenar adjuntos en Google Drive y enviar notificaciones por email con adjuntos.
 * Email destino: medibridgeusa@gmail.com
 */

function doPost(e) {
  try {
    // Parsear los datos del formulario
    const data = JSON.parse(e.postData.contents);
    
    // Validar que los datos existan
    if (!data.name || !data.issue) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'Por favor completa todos los campos requeridos'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Email destino
    const toEmail = 'medibridgeusa@gmail.com';
    
    // Asunto del email
    const subject = `Nueva Solicitud de Cotización: ${data.part || 'Repuesto'} - ${data.brand || ''} ${data.model || ''}`;
    
    // Procesar Archivos Adjuntos (Guardar en Google Drive y preparar para adjuntar al email)
    const attachments = [];
    const driveLinks = [];
    
    if (data.files && data.files.length > 0) {
      // Obtener o crear carpeta en Google Drive
      let folder;
      const folderName = "MediParts Adjuntos";
      const folders = DriveApp.getFoldersByName(folderName);
      
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder(folderName);
      }
      
      // Procesar cada archivo en Base64
      data.files.forEach(function(file) {
        try {
          const decoded = Utilities.base64Decode(file.base64);
          const blob = Utilities.newBlob(decoded, file.type, file.name);
          
          // Guardar en Drive
          const driveFile = folder.createFile(blob);
          driveLinks.push(driveFile.getUrl());
          
          // Guardar en array de adjuntos para Gmail
          attachments.push(blob);
        } catch (fileError) {
          Logger.log('Error procesando archivo ' + file.name + ': ' + fileError.toString());
        }
      });
    }

    // Crear el cuerpo del email en HTML
    const htmlBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
          }
          .header {
            background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 100%);
            color: white;
            padding: 30px;
            border-radius: 10px 10px 0 0;
            text-align: center;
          }
          .header h1 {
            margin: 0;
            font-size: 24px;
          }
          .content {
            background: #f8fafc;
            padding: 30px;
            border: 1px solid #e2e8f0;
          }
          .field {
            margin-bottom: 15px;
            background: white;
            padding: 12px 15px;
            border-radius: 8px;
            border-left: 4px solid #3b82f6;
          }
          .field-label {
            font-weight: bold;
            color: #64748b;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 3px;
          }
          .field-value {
            color: #0f172a;
            font-size: 15px;
            margin-top: 2px;
          }
          .issue-box {
            background: #fff;
            border: 2px solid #3b82f6;
            border-radius: 8px;
            padding: 15px;
            margin-top: 20px;
          }
          .links-box {
            background: #f1f5f9;
            border: 1px dashed #cbd5e1;
            border-radius: 8px;
            padding: 15px;
            margin-top: 15px;
          }
          .footer {
            background: #0f172a;
            color: #94a3b8;
            padding: 20px;
            text-align: center;
            font-size: 12px;
            border-radius: 0 0 10px 10px;
          }
          .badge {
            display: inline-block;
            background: #10b981;
            color: white;
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: bold;
            margin-top: 10px;
          }
          .timestamp {
            color: #64748b;
            font-size: 12px;
            margin-top: 15px;
            text-align: right;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>Nueva Solicitud de Cotización</h1>
          <div class="badge">MEDIPARTS SYSTEM</div>
        </div>
        
        <div class="content">
          <div class="field">
            <div class="field-label">Nombre del Cliente</div>
            <div class="field-value">${data.name}</div>
          </div>
          
          <div class="field">
            <div class="field-label">Clínica / Hospital / Empresa</div>
            <div class="field-value">${data.clinic || 'No especificado'}</div>
          </div>

          <div class="field">
            <div class="field-label">Email de Contacto</div>
            <div class="field-value">${data.email || 'No especificado'}</div>
          </div>

          <div class="field">
            <div class="field-label">WhatsApp / Teléfono</div>
            <div class="field-value">${data.phone || 'No especificado'}</div>
          </div>
          
          <div class="field">
            <div class="field-label">Equipo y Modelo</div>
            <div class="field-value">${data.brand || ''} ${data.model || ''}</div>
          </div>
          
          <div class="field">
            <div class="field-label">Pieza o Repuesto</div>
            <div class="field-value"><strong>${data.part || 'No especificado'}</strong> ${data.pn ? '(P/N: ' + data.pn + ')' : ''}</div>
          </div>
          
          <div class="issue-box">
            <div class="field-label">Detalles del Pedido / Falla</div>
            <div class="field-value" style="margin-top: 10px; white-space: pre-wrap;">${data.issue}</div>
          </div>

          ${driveLinks.length > 0 ? `
          <div class="links-box">
            <div class="field-label" style="color: #0f172a; font-weight: bold;">Enlaces a Fotos en Google Drive:</div>
            <ul style="margin: 5px 0 0 0; padding-left: 20px; font-size: 13px; color: #2563eb;">
              ${driveLinks.map((link, idx) => `<li><a href="${link}" target="_blank">Foto Adjunta ${idx+1}</a></li>`).join('')}
            </ul>
          </div>
          ` : ''}
          
          <div class="timestamp">
            <strong>Recibido:</strong> ${new Date().toLocaleString('es-ES', { 
              dateStyle: 'full', 
              timeStyle: 'short',
              timeZone: 'America/New_York'
            })}
          </div>
        </div>
        
        <div class="footer">
          <p><strong>MediParts</strong></p>
          <p>Notificaciones de Cotización Automáticas</p>
          <p>Este email contiene archivos adjuntos si el usuario los cargó en la web.</p>
        </div>
      </body>
      </html>
    `;
    
    // Crear versión de texto plano como fallback
    const plainBody = `
NUEVA SOLICITUD DE COTIZACIÓN - MEDIPARTS
================================================

Nombre del Cliente: ${data.name}
Clínica/Hospital: ${data.clinic || 'No especificado'}
Email: ${data.email || 'No especificado'}
WhatsApp/Teléfono: ${data.phone || 'No especificado'}
Equipo: ${data.brand || ''} ${data.model || ''}
Repuesto: ${data.part || 'No especificado'} ${data.pn ? '(P/N: ' + data.pn + ')' : ''}

DETALLE DEL PEDIDO / FALLA:
${data.issue}

------------------------------------------------
${driveLinks.length > 0 ? 'ARCHIVOS ADJUNTOS EN DRIVE:\n' + driveLinks.join('\n') : ''}
------------------------------------------------
Fecha de Solicitud: ${new Date().toLocaleString('es-ES')}
------------------------------------------------
    `;
    
    // Opciones del correo
    const emailOptions = {
      htmlBody: htmlBody,
      name: 'MediParts Web Contact'
    };

    // Agregar adjuntos si existen
    if (attachments.length > 0) {
      emailOptions.attachments = attachments;
    }

    // Enviar el email
    GmailApp.sendEmail(toEmail, subject, plainBody, emailOptions);
    
    // Guardar en la hoja de cálculo de Google Sheets
    saveToSpreadsheet(data, driveLinks.join(', '));
    
    // Retornar respuesta exitosa
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Solicitud enviada exitosamente. Nos pondremos en contacto pronto.'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    // Manejo de errores
    Logger.log('Error: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Hubo un error al enviar la solicitud. Por favor intenta nuevamente.'
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Función para guardar datos en Google Sheets
 */
function saveToSpreadsheet(data, linksString) {
  try {
    // ID de la hoja de cálculo de Google Sheets
    const SPREADSHEET_ID = '1kVhCypSQyg9A9Y_Sk0uaLeBpN_pL18L__JNgoJhpU4A';
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Contactos') || ss.insertSheet('Contactos');
    
    // Si la hoja está vacía, agregar encabezados correspondientes al nuevo formulario
    if (sheet.getLastRow() === 0) {
      const headers = ['Fecha', 'Hora', 'Nombre', 'Clínica/Hospital', 'Email', 'WhatsApp/Teléfono', 'Marca', 'Modelo', 'Pieza Solicitada', 'P/N', 'Falla/Descripción', 'Fotos Adjuntas', 'Estado'];
      sheet.appendRow(headers);
      
      // Formatear encabezados
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#0f172a');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // Ajustar ancho de columnas
      sheet.setColumnWidth(1, 100); // Fecha
      sheet.setColumnWidth(2, 90);  // Hora
      sheet.setColumnWidth(3, 160); // Nombre
      sheet.setColumnWidth(4, 180); // Clínica
      sheet.setColumnWidth(5, 160); // Email
      sheet.setColumnWidth(6, 130); // Teléfono
      sheet.setColumnWidth(7, 120); // Marca
      sheet.setColumnWidth(8, 120); // Modelo
      sheet.setColumnWidth(9, 180); // Pieza
      sheet.setColumnWidth(10, 100); // P/N
      sheet.setColumnWidth(11, 350); // Falla
      sheet.setColumnWidth(12, 250); // Fotos
      sheet.setColumnWidth(13, 100); // Estado
      
      // Congelar fila de encabezados
      sheet.setFrozenRows(1);
    }
    
    // Obtener fecha y hora actual
    const now = new Date();
    const fecha = Utilities.formatDate(now, 'America/New_York', 'dd/MM/yyyy');
    const hora = Utilities.formatDate(now, 'America/New_York', 'HH:mm:ss');
    
    // Agregar la nueva fila
    const newRow = [
      fecha,
      hora,
      data.name,
      data.clinic || 'No especificado',
      data.email || 'No especificado',
      data.phone || 'No especificado',
      data.brand || 'No especificado',
      data.model || 'No especificado',
      data.part || 'No especificado',
      data.pn || 'No especificado',
      data.issue,
      linksString || 'Ninguno',
      'Nuevo'
    ];
    
    sheet.appendRow(newRow);
    
    // Formatear la nueva fila
    const lastRow = sheet.getLastRow();
    const rowRange = sheet.getRange(lastRow, 1, 1, newRow.length);
    
    // Alternar color de fondo
    if (lastRow % 2 === 0) {
      rowRange.setBackground('#f8fafc');
    }
    
    // Formatear columna de estado
    const statusCell = sheet.getRange(lastRow, 13);
    statusCell.setBackground('#dbeafe');
    statusCell.setFontColor('#1d4ed8');
    statusCell.setFontWeight('bold');
    statusCell.setHorizontalAlignment('center');
    
  } catch (error) {
    Logger.log('Error guardando en Spreadsheet: ' + error.toString());
  }
}
