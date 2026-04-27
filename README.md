# Company-Project
This project was created in order to reduce AHT and maintain acurracy on the company auditable processes. This idea was launched due to areas of opportunities found in AHT and weekly audits, basically it is a template tool that made our cases faster maintain acurracy in order to prevent any deduction in weekly audits.
2 tools were created, 1 calles SA Tool and the second one is Record keeping tool
Bellow I am leaving the codes, please be aware that these codes werenmade with the help of IA, not made %100 of IA, I was in charge of supervising that everything is set and working great by providing comands. 
SA Tool
It is a manipulable template were notes have to be left on text boxes, on the drop down menu, according to the selection, the template freely adapts to our necesity, I google sheets data was created in order to keep and track users and cases. 
Code gs:
return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Suspension Appeals Tool')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// GUARDA EL USO NORMAL (TRACKER)
function logUsage(sfCase, source, ahtSeconds) {
  try {
    var sheetId = '1_Sk8ekFKxwjvajdksT908HMLEOTSnKvCH5Pw_fMNLpM';
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Hoja 1') || SpreadsheetApp.openById(sheetId).getSheets()[0];
    
    var email = Session.getActiveUser().getEmail();
    var timestamp = new Date(); 
    var lastRow = sheet.getLastRow();
    
    var isDuplicateByMe = false;
    var isDuplicateByOther = false;
    
    if (lastRow > 1) {
      var data = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
      var sfCaseStr = String(sfCase).trim(); 
      for (var i = 0; i < data.length; i++) {
        var rowEmail = String(data[i][0]).trim();
        var rowCase = String(data[i][1]).trim();
        if (rowCase === sfCaseStr) {
          if (rowEmail === email) {
            isDuplicateByMe = true; break; 
          } else {
            isDuplicateByOther = true; 
          }
        }
      }
    }
    
    if (isDuplicateByMe) return true; 
    
    sheet.appendRow([timestamp, email, sfCase, source, ahtSeconds]);
    
    if (isDuplicateByOther) {
      var newRow = sheet.getLastRow();
      var cell = sheet.getRange(newRow, 3);
      cell.setBackground('#ffcdd2').setFontColor('#b71c1c').setFontWeight('bold').setNote('⚠️ Duplicate Case!');
    }
    return true;
  } catch(error) {
    Logger.log(error.toString()); return false;
  }
}

// ==========================================
// NUEVO: SISTEMA DE SINCRONIZACIÓN EN LA NUBE
// ==========================================
function getUserData() {
  try {
    var sheetId = '1_Sk8ekFKxwjvajdksT908HMLEOTSnKvCH5Pw_fMNLpM';
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('UserSettings');
    if (!sheet) return null; // Aún no existe la pestaña

    var email = Session.getActiveUser().getEmail();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        return {
          width: data[i][1] || '320px',
          height: data[i][2] || '400px',
          text1: data[i][3] || '',
          text2: data[i][4] || '',
          text3: data[i][5] || '',
          text4: data[i][6] || '',
          text5: data[i][7] || ''
        };
      }
    }
    return null; 
  } catch (e) { return null; }
}

function saveUserData(userData) {
  try {
    var sheetId = '1_Sk8ekFKxwjvajdksT908HMLEOTSnKvCH5Pw_fMNLpM';
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('UserSettings');

    // Si la pestaña secreta no existe, el código la crea sola
    if (!sheet) {
      sheet = ss.insertSheet('UserSettings');
      sheet.appendRow(['Email', 'Width', 'Height', 'Note1', 'Note2', 'Note3', 'Note4', 'Note5']);
      sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#d0e0e3");
    }

    var email = Session.getActiveUser().getEmail();
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        rowIndex = i + 1; break;
      }
    }

    var rowData = [ email, userData.width, userData.height, userData.text1, userData.text2, userData.text3, userData.text4, userData.text5 ];

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, 8).setValues([rowData]); // Actualiza sus notas
    } else {
      sheet.appendRow(rowData); // Nuevo usuario
    }
    return true;
  } catch (e) { return false; }
}

HTML

<!DOCTYPE html>
<html lang="en">
  <head>
    <base target="_top">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700;900&display=swap" rel="stylesheet">
    
    <style>
      /* =========================================================
         VARIABLES DE DISEÑO
         ========================================================= */
      :root {
        --remitly-blue: #1a3b5c; 
        --card-bg: #f4f7f6;      
        --text-dark: #2c3e50;    
        --input-border: #cbd5e1; 
        --accent-green: #27ae60;
        --link-color: #2563eb;
        --danger-red: #d93025;
        --bg-color: #1a3b5c;     
        --header-text: #ffffff;
        --grid-bg: #e2e8f0;      
        --disclaimer-bg: #fef3c7;
        --disclaimer-text: #92400e;
        --focus-ring: rgba(37, 99, 235, 0.2);
        
        /* Colores Cronómetro Light */
        --timer-green: #27ae60;
        --timer-yellow: #d97706;
        --timer-orange: #ea580c;
        --timer-red: #d93025;

        /* Colores Gradiente Animado Light */
        --grad-1: #1a3b5c; 
        --grad-2: #1e4a79; 
        --grad-3: #112b47; 
        --grad-4: #25588c;
      }

      [data-theme="dark"] {
        --remitly-blue: #38bdf8; 
        --card-bg: #16203d;      
        --text-dark: #f1f5f9;    
        --input-border: #2c3a5a; 
        --accent-green: #10b981; 
        --link-color: #60a5fa;   
        --danger-red: #fb7185;   
        --bg-color: #0a1128;     
        --header-text: #ffffff;
        --grid-bg: #1e2a4f;      
        --disclaimer-bg: #422006; 
        --disclaimer-text: #fde047; 
        --focus-ring: rgba(56, 189, 248, 0.4);
        
        /* Colores Cronómetro Dark */
        --timer-green: #10b981;
        --timer-yellow: #fbbf24;
        --timer-orange: #f97316;
        --timer-red: #ef4444;

        /* Colores Gradiente Animado Dark */
        --grad-1: #0a1128; 
        --grad-2: #0e1a3a; 
        --grad-3: #050914; 
        --grad-4: #13224a;
      }

      body {
        font-family: 'Roboto', sans-serif;
        background-color: var(--bg-color);
        margin: 0;
        padding: 20px;
        display: flex;
        flex-direction: column;
        align-items: center;
        min-height: 100vh;
        color: var(--header-text);
        transition: background-color 0.4s ease, color 0.4s ease;
      }

      /* --- EFECTO DE GRADIENTE ANIMADO --- */
      body.animated-bg {
        background: linear-gradient(-45deg, var(--grad-1), var(--grad-2), var(--grad-3), var(--grad-4));
        background-size: 400% 400%;
        animation: slowGradient 15s ease infinite;
      }

      @keyframes slowGradient {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
      }

      /* --- CABECERA FLOTANTE BALANCEADA --- */
      .app-header {
        position: sticky; top: 0; z-index: 1000; 
        background-color: var(--bg-color); 
        width: 100%; max-width: 800px; 
        display: flex; justify-content: space-between; align-items: center; 
        margin-top: -20px; margin-bottom: 25px; padding: 20px 20px 15px 20px; 
        box-shadow: 0 10px 15px -10px rgba(0,0,0,0.1); 
        transition: background-color 0.4s ease, box-shadow 0.4s ease;
      }
      body.animated-bg .app-header {
        background-color: transparent; 
        backdrop-filter: blur(10px);
        border-bottom: 1px solid rgba(255,255,255,0.1);
      }
      [data-theme="dark"] .app-header { box-shadow: 0 10px 20px -10px rgba(0,0,0,0.6); }
      
      .header-left { display: flex; flex: 1; justify-content: flex-start; }
      .header-center { display: flex; align-items: center; justify-content: center; gap: 20px; }
      .header-right { display: flex; flex: 1; justify-content: flex-end; gap: 15px; font-size: 24px; align-items: center; }

      .theme-switch, .history-switch, .settings-switch, .metrics-switch, .cloud-switch { cursor: pointer; user-select: none; transition: transform 0.2s; opacity: 0.9; }
      .theme-switch:hover, .history-switch:hover, .settings-switch:hover, .metrics-switch:hover, .cloud-switch:hover { transform: scale(1.1); opacity: 1; }
      
      /* --- DISEÑO DEL CRONÓMETRO --- */
      .aht-timer-container {
        display: none; align-items: center; gap: 6px; background-color: var(--card-bg); border: 1px solid var(--input-border); 
        padding: 6px 14px; border-radius: 20px; box-shadow: 0 4px 10px rgba(0,0,0,0.15); transition: background-color 0.4s ease, border-color 0.4s ease;
      }
      [data-theme="dark"] .aht-timer-container { box-shadow: 0 4px 10px rgba(0,0,0,0.5); }
      .aht-timer-icon { font-size: 14px; }
      .aht-timer { font-family: 'Consolas', monospace; font-size: 14px; font-weight: 900; transition: color 0.3s ease; }
      .timer-green { color: var(--timer-green); } .timer-yellow { color: var(--timer-yellow); } .timer-orange { color: var(--timer-orange); } .timer-red { color: var(--timer-red); }

      .brand-block { display: flex; flex-direction: column; align-items: center; justify-content: center; }
      .company-name { font-size: 28px; font-weight: 900; line-height: 1; letter-spacing: -0.5px; }
      .company-slogan { font-size: 11px; font-weight: 400; text-transform: uppercase; letter-spacing: 1px; opacity: 0.9; margin-top: 4px; }
      .divider { font-size: 35px; font-weight: 300; color: rgba(255, 255, 255, 0.4); margin-bottom: 5px; }
      .header-title { font-size: 22px; font-weight: 500; text-transform: uppercase; letter-spacing: 1px; margin: 0; padding-top: 2px; transition: color 0.4s ease;}
      [data-theme="dark"] .header-title { color: var(--remitly-blue); }

      .container { background-color: var(--card-bg); color: var(--text-dark); width: 100%; max-width: 750px; padding: 40px; border-radius: 12px; box-shadow: 0 10px 30px rgba(0,0,0,0.3); transition: background-color 0.4s ease, color 0.4s ease, border 0.4s ease; }
      [data-theme="dark"] .container { box-shadow: 0 10px 35px rgba(0,0,0,0.6); border: 1px solid #2c3a5a; }

      .disclaimer { margin-bottom: 15px; padding: 18px; background-color: var(--disclaimer-bg); border-left: 5px solid var(--disclaimer-text); font-size: 14px; color: var(--disclaimer-text); border-radius: 6px; line-height: 1.6; transition: all 0.4s ease; }

      .knowledge-link-container { background-color: rgba(59, 130, 246, 0.08); border: 1px solid rgba(59, 130, 246, 0.3); padding: 12px 18px; border-radius: 6px; margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center; transition: all 0.4s ease; }
      [data-theme="dark"] .knowledge-link-container { background-color: rgba(56, 189, 248, 0.1); border-color: var(--remitly-blue); }
      .knowledge-title { font-weight: 700; font-size: 14px; color: var(--remitly-blue); text-transform: uppercase; transition: color 0.4s ease;}
      .knowledge-link { color: var(--link-color); font-weight: 700; text-decoration: none; transition: 0.2s; }
      .knowledge-link:hover { text-decoration: underline; color: var(--remitly-blue); }

      .section-header { border-bottom: 2px solid var(--remitly-blue); color: var(--remitly-blue); padding-bottom: 8px; margin-top: 30px; margin-bottom: 20px; font-weight: 900; text-transform: uppercase; font-size: 15px; transition: all 0.4s ease; }
      .section-header.first-header { margin-top: 0; }

      label { font-weight: 700; font-size: 12px; color: var(--text-dark); opacity: 0.8; margin-top: 15px; display: block; text-transform: uppercase; letter-spacing: 0.5px; transition: color 0.4s ease; }
      .required { color: var(--danger-red); opacity: 1; transition: color 0.4s ease;}

      select, input[type="text"], textarea { width: 100%; padding: 12px; margin-top: 5px; border: 1px solid var(--input-border); border-radius: 6px; font-family: 'Roboto', sans-serif; font-size: 14px; background-color: #ffffff; color: var(--text-dark); transition: border-color 0.3s ease, background-color 0.3s ease, color 0.3s ease; box-sizing: border-box; }
      textarea:not(.sp-textarea):not(#finalText) { resize: vertical; overflow-y: auto; min-height: 80px; }
      [data-theme="dark"] select, [data-theme="dark"] input[type="text"], [data-theme="dark"] textarea { background-color: #050a1f; color: #ffffff; border-color: #2c3a5a; }
      select:focus, input[type="text"]:focus, textarea:focus { border-color: var(--remitly-blue); outline: none; box-shadow: 0 0 0 3px var(--focus-ring); }

      /* =========================================================
         SISTEMA DE BOTONES (CHECKLIST)
         ========================================================= */
      .action-columns { display: flex; gap: 20px; margin-top: 15px; width: 100%; transition: all 0.3s ease; }
      .action-col-left, .action-col-right { display: flex; flex-direction: column; gap: 12px; flex: 1; transition: all 0.3s ease; }

      .action-btn { display: block; position: relative; cursor: pointer; user-select: none; margin: 0; }
      .action-btn input { position: absolute; opacity: 0; cursor: pointer; height: 0; width: 0; }
      
      .btn-content { display: flex; align-items: center; justify-content: center; text-align: center; padding: 14px 16px; border: 2px solid var(--input-border); border-radius: 8px; transition: all 0.2s ease; background: var(--grid-bg); font-size: 13px; font-weight: 500; color: var(--text-dark); line-height: 1.4; box-sizing: border-box; height: 100%; }
      [data-theme="dark"] .btn-content { background: #1e2a4f; border-color: #2c3a5a; }
      
      .action-btn:hover .btn-content { border-color: #94a3b8; }
      [data-theme="dark"] .action-btn:hover .btn-content { border-color: #475569; }

      .action-btn input:checked ~ .btn-content { border-color: var(--remitly-blue); background-color: #e0f2fe; font-weight: 900; color: #0369a1; box-shadow: inset 0 0 0 1px var(--remitly-blue), 0 4px 10px rgba(0,0,0,0.05); }
      [data-theme="dark"] .action-btn input:checked ~ .btn-content { background-color: rgba(56, 189, 248, 0.25); border-color: #38bdf8; color: #bae6fd; box-shadow: inset 0 0 0 1px #38bdf8; }
      .action-btn input:checked ~ .btn-content::after { content: " ✔"; margin-left: 8px; font-weight: 900; font-size: 1.2em; }

      .action-btn input:disabled ~ .btn-content { cursor: not-allowed; opacity: 0.8; }
      
      .status-menu { margin-top: 15px; padding: 18px; border-radius: 8px; border: 1px solid transparent; display: none; grid-template-columns: 1fr 1fr; gap: 12px; transition: all 0.3s ease; }
      
      .menu-reinstated { background-color: #ecfdf5; border-color: #a7f3d0; }
      .menu-pending { background-color: #fff7ed; border-color: #fed7aa; }
      .menu-suspended { background-color: #fff1f2; border-color: #fecdd3; }
      
      [data-theme="dark"] .menu-reinstated { background-color: rgba(16, 185, 129, 0.08); border-color: rgba(16, 185, 129, 0.2); }
      [data-theme="dark"] .menu-pending { background-color: rgba(249, 115, 22, 0.08); border-color: rgba(249, 115, 22, 0.2); }
      [data-theme="dark"] .menu-suspended { background-color: rgba(244, 63, 94, 0.08); border-color: rgba(244, 63, 94, 0.2); }

      .status-menu .btn-content { background: var(--card-bg); }
      [data-theme="dark"] .status-menu .btn-content { background: #16203d; }

      .menu-reinstated .action-btn input:checked ~ .btn-content { border-color: #059669; color: #064e3b; background-color: #d1fae5; box-shadow: inset 0 0 0 1px #059669; }
      [data-theme="dark"] .menu-reinstated .action-btn input:checked ~ .btn-content { color: #6ee7b7; background-color: rgba(16, 185, 129, 0.25); box-shadow: inset 0 0 0 1px #10b981; }
      
      .menu-pending .action-btn input:checked ~ .btn-content { border-color: #ea580c; color: #7c2d12; background-color: #ffedd5; box-shadow: inset 0 0 0 1px #ea580c; }
      [data-theme="dark"] .menu-pending .action-btn input:checked ~ .btn-content { color: #fdba74; background-color: rgba(249, 115, 22, 0.25); box-shadow: inset 0 0 0 1px #f97316; }
      
      .menu-suspended .action-btn input:checked ~ .btn-content { border-color: #e11d48; color: #881337; background-color: #ffe4e6; box-shadow: inset 0 0 0 1px #e11d48; }
      [data-theme="dark"] .menu-suspended .action-btn input:checked ~ .btn-content { color: #fda4af; background-color: rgba(244, 63, 94, 0.25); box-shadow: inset 0 0 0 1px #f43f5e; }

      .hidden { display: none !important; }
      .fade-in { animation: fadeIn 0.4s ease-out; }
      @keyframes fadeIn { from { opacity: 0; transform: translateY(-10px); } to { opacity: 1; transform: translateY(0); } }

      @keyframes bounce-arrow { 0%, 100% { transform: translateY(0); } 50% { transform: translateY(8px); } }
      .guide-container { text-align: center; margin-top: 25px; margin-bottom: -15px; }
      .guide-text { color: var(--danger-red); font-weight: 700; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; transition: color 0.4s ease;}
      .guide-arrow { color: var(--danger-red); font-size: 28px; animation: bounce-arrow 1.5s infinite ease-in-out; margin-top: 5px; transition: color 0.4s ease;}

      /* FORM LINKS SECTION */
      .additional-form-section { margin-top: 25px; text-align: center; padding: 25px; border-radius: 8px; border: 2px solid transparent; transition: all 0.4s ease;}
      .highlight-form { background-color: rgba(217, 48, 37, 0.05); border: 2px dashed var(--danger-red); }
      [data-theme="dark"] .highlight-form { background-color: rgba(251, 113, 133, 0.1); border-color: rgba(251, 113, 133, 0.4); }
      .additional-form-text { color: var(--text-dark); font-size: 14px; margin-bottom: 15px; font-weight: 500; transition: color 0.4s ease;}
      
      .additional-form-link { display: inline-block; background-color: var(--danger-red); color: #ffffff !important; font-size: 14px; font-weight: 700; text-decoration: none; padding: 12px 24px; border-radius: 6px; transition: all 0.3s ease; box-shadow: 0 4px 6px rgba(217, 48, 37, 0.2); text-align: center; }
      .additional-form-link:hover { background-color: #b71c1c; transform: translateY(-2px); box-shadow: 0 6px 12px rgba(217, 48, 37, 0.4); }
      [data-theme="dark"] .additional-form-link { background-color: #e11d48; box-shadow: 0 4px 10px rgba(225, 29, 72, 0.4); }
      [data-theme="dark"] .additional-form-link:hover { background-color: #be123c; }

      .concessions-form-link { display: inline-block; background-color: #6366f1; color: #ffffff !important; font-size: 14px; font-weight: 700; text-decoration: none; padding: 12px 24px; border-radius: 6px; transition: all 0.3s ease; box-shadow: 0 4px 6px rgba(99, 102, 241, 0.2); text-align: center; }
      .concessions-form-link:hover { background-color: #4f46e5; transform: translateY(-2px); box-shadow: 0 6px 12px rgba(99, 102, 241, 0.4); }
      [data-theme="dark"] .concessions-form-link { background-color: #818cf8; color: #0a1128 !important; box-shadow: 0 4px 10px rgba(129, 140, 248, 0.4); }
      [data-theme="dark"] .concessions-form-link:hover { background-color: #a5b4fc; }

      .btn-row { display: flex; gap: 15px; margin-top: 35px; }
      button { flex: 1; padding: 14px; border: none; border-radius: 6px; font-weight: 700; cursor: pointer; text-transform: uppercase; font-size: 13px; transition: all 0.3s ease; letter-spacing: 0.5px;}
      .btn-gen { background-color: var(--remitly-blue); color: white; }
      [data-theme="dark"] .btn-gen { color: #0a1128; box-shadow: 0 4px 15px rgba(56, 189, 248, 0.3); } 
      .btn-gen:hover { transform: translateY(-2px); opacity: 0.9;}
      .btn-reset { background-color: #e0e0e0; color: #333; }
      [data-theme="dark"] .btn-reset { background-color: #1e2a4f; color: #f1f5f9; border: 1px solid #2c3a5a;}
      .btn-reset:hover { background-color: #d6d6d6; }
      [data-theme="dark"] .btn-reset:hover { background-color: #2c3a5a; }
      .btn-copy { background-color: var(--accent-green); color: white; margin-top: 15px; width: 100%; display: none; }
      [data-theme="dark"] .btn-copy { color: #0a1128; box-shadow: 0 4px 15px rgba(16, 185, 129, 0.3); }
      .btn-copy:hover { transform: translateY(-2px); opacity: 0.9; }

      #outputContainer { margin-top: 30px; display: none; border-top: 2px dashed var(--input-border); padding-top: 20px; transition: border-color 0.4s ease;}
      #finalText { background: #050a1f; color: #34d399; font-family: 'Consolas', monospace; height: 500px; display: none; border: 1px solid var(--input-border); border-radius: 8px; padding: 15px; transition: all 0.4s ease; overflow-y: auto; }
      
      .footer-author { margin-top: 40px; font-size: 16px; font-weight: 700; color: #ffffff; text-align: center; letter-spacing: 0.5px; margin-bottom: 20px; text-shadow: 0 2px 4px rgba(0,0,0,0.5); }

      /* --- AUTO-SAVE TOAST INDICATOR --- */
      #autoSaveStatus {
        position: fixed;
        bottom: 20px;
        right: 20px;
        background-color: rgba(0, 0, 0, 0.7);
        color: white;
        padding: 8px 16px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: bold;
        pointer-events: none;
        opacity: 0;
        transition: opacity 0.3s ease;
        z-index: 9999;
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
      }
      [data-theme="dark"] #autoSaveStatus { background-color: rgba(255, 255, 255, 0.1); backdrop-filter: blur(5px); }

      /* =========================================================
         PANEL LATERAL Y MODALES
         ========================================================= */
      #leftScratchpad { position: fixed; top: 15vh; left: 0; transform: translateX(-100%); background-color: var(--card-bg); border: 1px solid var(--input-border); border-left: none; border-radius: 0 12px 12px 0; box-shadow: 5px 5px 25px rgba(0,0,0,0.15); z-index: 9999; transition: transform 0.4s cubic-bezier(0.4, 0, 0.2, 1); display: flex; flex-direction: column; padding: 15px 20px 20px 20px; width: max-content; }
      [data-theme="dark"] #leftScratchpad { box-shadow: 5px 5px 25px rgba(0,0,0,0.8); }
      #leftScratchpad.open { transform: translateX(0); }
      .scratchpad-tab { position: absolute; right: -45px; top: 20px; width: 45px; height: 120px; background-color: var(--card-bg); color: var(--text-dark); border: 1px solid var(--input-border); border-left: none; border-radius: 0 8px 8px 0; cursor: pointer; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 15px; writing-mode: vertical-rl; text-orientation: mixed; box-shadow: 3px 0 10px rgba(0,0,0,0.1); transition: background-color 0.3s ease, color 0.3s ease, opacity 0.3s; user-select: none; }
      .scratchpad-tab:hover { opacity: 0.9; }
      .sp-top-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; border-bottom: 1px solid var(--input-border); padding-bottom: 10px; }
      .sp-header-left { display: flex; align-items: center; gap: 8px; }
      .sp-menu-btn { background: transparent; border: none; font-size: 20px; cursor: pointer; color: var(--text-dark); padding: 0 5px; opacity: 0.7; transition: opacity 0.2s; outline: none; margin: 0; flex: none; width: auto; }
      .sp-menu-btn:hover { opacity: 1; color: var(--remitly-blue); }
      [data-theme="dark"] .sp-menu-btn:hover { color: var(--remitly-blue); }
      .scratchpad-title { font-size: 14px; font-weight: 900; color: var(--text-dark); margin: 0; text-transform: uppercase; }
      #btnSync { background: transparent; border: 1px solid var(--input-border); border-radius: 6px; font-size: 11px; font-weight: 700; cursor: pointer; padding: 5px 10px; color: var(--text-dark); opacity: 0.7; transition: all 0.3s ease; text-transform: none; letter-spacing: 0; flex: none; width: auto; margin:0; }
      #btnSync:hover { opacity: 1; border-color: var(--remitly-blue); }
      .sp-main-layout { display: flex; gap: 10px; flex: 1; align-items: stretch; }
      .sp-left-nav { display: flex; flex-direction: column; gap: 4px; width: 55px; overflow: hidden; transition: width 0.3s ease, opacity 0.3s ease, margin 0.3s ease; }
      .sp-left-nav.collapsed { width: 0px; opacity: 0; margin-right: -10px; pointer-events: none; }
      .sp-nav-btn { background-color: transparent; border: 1px solid var(--input-border); color: var(--text-dark); font-size: 11px; font-weight: 700; padding: 5px 6px; border-radius: 4px; cursor: pointer; opacity: 0.6; transition: all 0.2s ease; margin:0; text-align: center; white-space: nowrap; }
      .sp-nav-btn:hover { opacity: 0.8; }
      .sp-nav-btn.active { background-color: var(--remitly-blue); color: white; opacity: 1; border-color: var(--remitly-blue); }
      [data-theme="dark"] .sp-nav-btn.active { background-color: #38bdf8; color: #0a1128; border-color: #38bdf8; }
      .sp-content-area { flex: 1; display: flex; flex-direction: column; }
      .sp-textarea { display: none; width: 320px; min-width: 250px; max-width: 70vw; height: 400px; min-height: 200px; resize: both; font-family: 'Roboto', sans-serif; font-size: 13px; line-height: 1.5; margin: 0; }
      .sp-textarea.active { display: block; }

      .modal-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); backdrop-filter: blur(5px); display: none; justify-content: center; align-items: center; z-index: 10000; animation: fadeInModal 0.2s ease-out; }
      .modal-box { background: var(--card-bg); color: var(--text-dark); width: 90%; max-width: 400px; border-radius: 12px; padding: 30px 25px; text-align: center; border: 1px solid var(--input-border); animation: slideUpModal 0.3s ease-out;}
      @keyframes fadeInModal { from { opacity: 0; } to { opacity: 1; } }
      @keyframes slideUpModal { from { transform: translateY(20px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
      .modal-icon { font-size: 45px; margin-bottom: 15px; line-height: 1; }
      .modal-title { font-size: 20px; font-weight: 700; margin-bottom: 10px; color: var(--remitly-blue); }
      [data-theme="dark"] .modal-title { color: var(--remitly-blue); }
      .modal-text { font-size: 14px; line-height: 1.5; opacity: 0.8; margin-bottom: 25px; }
      .modal-buttons { display: flex; gap: 10px; justify-content: center; flex-direction: row; }
      .modal-buttons button { margin: 0; flex: 1; }
      .btn-danger { background: var(--danger-red); color: white; }
      .btn-primary { background: var(--remitly-blue); color: white; }
      [data-theme="dark"] .btn-primary { color: #0a1128; background: #38bdf8;}

      .history-list-container { max-height: 300px; overflow-y: auto; margin-top: 15px; text-align: left;}
      .history-item { background: var(--grid-bg); border: 1px solid var(--input-border); padding: 12px 15px; border-radius: 6px; margin-bottom: 10px; display: flex; justify-content: space-between; align-items: center; transition: all 0.2s ease; }
      [data-theme="dark"] .history-item { background: #1e2a4f; border-color: #2c3a5a; }
      .history-info { display: flex; flex-direction: column; gap: 4px; }
      .history-case { font-size: 14px; font-weight: 900; color: var(--remitly-blue); }
      [data-theme="dark"] .history-case { color: var(--remitly-blue); }
      .history-time { font-size: 11px; opacity: 0.7; font-weight: 500; }
      .history-copy-btn { background: var(--accent-green); color: white; padding: 6px 12px; border-radius: 4px; border: none; cursor: pointer; font-size: 11px; font-weight: bold; text-transform: uppercase; transition: 0.2s; width: auto; flex: none; margin: 0; }
      [data-theme="dark"] .history-copy-btn { color: #0a1128; }
      .history-copy-btn:hover { opacity: 0.9; transform: translateY(-1px); }
      .switch-container { display: flex; align-items: center; justify-content: space-between; margin-top: 15px; text-align: left; background: var(--grid-bg); padding: 12px 15px; border-radius: 6px; }
      [data-theme="dark"] .switch-container { background: #1e2a4f; }
      .switch-label { font-size: 13px; font-weight: 700; margin: 0; text-transform: none; display: flex; align-items: center; gap: 8px;}
      .switch { position: relative; display: inline-block; width: 44px; height: 24px; }
      .switch input { opacity: 0; width: 0; height: 0; }
      .slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #ccc; transition: .4s; border-radius: 24px; }
      .slider:before { position: absolute; content: ""; height: 18px; width: 18px; left: 3px; bottom: 3px; background-color: white; transition: .4s; border-radius: 50%; }
      input:checked + .slider { background-color: var(--accent-green); }
      input:checked + .slider:before { transform: translateX(20px); }
    </style>
  </head>
  <body>

    <div id="autoSaveStatus">⏳ Saving...</div>

    <div class="app-header">
      <div class="header-left">
        <div id="ahtTimerDisplay" class="aht-timer-container">
          <span class="aht-timer-icon">⏱️</span>
          <span id="ahtTimerText" class="aht-timer timer-green">00:00</span>
        </div>
      </div>
      <div class="header-center">
        <div class="brand-block">
          <span class="company-name">Remitly</span>
          <span class="company-slogan">Promises Delivered</span>
        </div>
        <div class="divider">|</div>
        <h1 class="header-title">Suspension Appeals</h1>
      </div>
      <div class="header-right">
        <div class="cloud-switch" onclick="openCloudModal()" title="Emergency Cloud Sync">☁️</div>
        <div class="metrics-switch" onclick="openMetricsModal()" title="View Daily Metrics">📊</div>
        <div class="settings-switch" onclick="openSettingsModal()" title="Settings">⚙️</div>
        <div class="history-switch" onclick="openHistoryModal()" title="View Recent Cases">🕒</div>
        <div class="theme-switch" onclick="toggleTheme()" title="Toggle Dark Mode">🌙</div>
      </div>
    </div>

    <div class="container">

      <div class="disclaimer">
        <strong>Disclaimer:</strong> This template is designed to streamline the review process. Please supplement these fields with any additional relevant information and ensure all assessments strictly adhere to current Standard Operating Procedures (SOPs).
      </div>

      <div class="knowledge-link-container">
        <span class="knowledge-title">Knowledge Articles:</span>
        <a href="https://sites.google.com/remitly.com/knowledge/home" target="_blank" class="knowledge-link">Access SOPs & Guidelines Here 📚</a>
      </div>
      
      <div class="section-header first-header">Findings</div>

      <label>WAS THE ACCOUNT SUSPENDED BY KOUNT OR MANUALLY? <span class="required">*</span></label>
      <select id="sourceSelect" onchange="handleSourceChange()">
        <option value="" selected>-- Select an option --</option>
        <option value="Kount">Kount</option>
        <option value="Manual">Manual</option>
        <option value="Account already worked by another associate">Account already worked by another associate</option>
        <option value="The account is not suspended, it is deleted.">The account is not suspended, it is deleted.</option>
        <option value="The account is not suspended, it is active.">The account is not suspended, it is active.</option>
        <option value="***Account suspended by the SMB Team***">***Account suspended by the SMB Team***</option>
      </select>

      <div id="kountOptions" class="hidden fade-in">
        <label>SUSPENSION TYPE (KOUNT) <span class="required">*</span></label>
        <select id="kountType">
          <option value="Fraud">Fraud</option>
          <option value="Policy">Policy</option>
          <option value="Compliance Rule">Suspended by a compliance rule</option>
        </select>
      </div>

      <div id="manualOptions" class="hidden fade-in">
        <label>SUSPENSION TYPE (MANUAL) <span class="required">*</span></label>
        <select id="manualType">
          <option value="Fraud">Fraud</option>
          <option value="Policy Violation">Policy violation</option>
          <option value="Suspended by Compliance">Suspended by compliance</option>
          <option value="Compliance Rule">Suspended by a compliance rule</option>
          <option value="Stolen Info">Stolen info</option>
          <option value="ATO">ATO</option>
        </select>
      </div>

      <div id="kountDetailsSection" class="hidden fade-in">
        <label style="color:var(--remitly-blue);">TITLE: KOUNT (ENTER DETAILS BELOW) <span class="required">*</span></label>
        <textarea id="kountDetailText" placeholder="Enter rule, score or manual details here..." spellcheck="true" lang="en"></textarea>
      </div>

      <div id="manualDetailsSection" class="hidden fade-in">
        <label style="color:var(--remitly-blue);">ENTER SUSPENSION REASON <span class="required">*</span></label>
        <textarea id="manualDetailText" placeholder="Enter suspension reason here..." spellcheck="true" lang="en"></textarea>
      </div>

      <div id="fullReviewContainer" class="fade-in">
        <label>PLEASE SHARE ANY INFO PROVIDED BY THE CX IF AVAILABLE</label>
        <textarea id="cxInfo" placeholder="Customer mentioned..." spellcheck="true" lang="en"></textarea>

        <div class="section-header">Account's Review</div>
        <label>SIFT SCORE</label><textarea id="siftScore" rows="1" spellcheck="true" lang="en"></textarea>
        <label>PAYMENT METHOD</label><textarea id="paymentMethod" rows="1" spellcheck="true" lang="en"></textarea>
        <label>LOGINS</label><textarea id="logins" rows="1" spellcheck="true" lang="en"></textarea>
        <label>PHONE</label><textarea id="phone" rows="1" spellcheck="true" lang="en"></textarea>
        <label>EMAIL</label><textarea id="email" rows="1" spellcheck="true" lang="en"></textarea>
        <label>ACCOUNT STATUS</label><textarea id="accountStatus" rows="1" spellcheck="true" lang="en"></textarea>
        <label>DUP ACCOUNTS</label><textarea id="dupAccounts" rows="1" spellcheck="true" lang="en"></textarea>
        <label>ADDITIONAL FINDINGS</label><textarea id="additionalFindings" spellcheck="true" lang="en"></textarea>
      </div>

      <div class="section-header">Actions</div>

      <label>ACTION TAKEN</label>
      <textarea id="actionsText" spellcheck="true" lang="en"></textarea>

      <label>CASE CHECKLIST:</label>
      
      <div class="action-columns" id="mainChecklistGrid">
        <div class="action-col-left">
          <label class="action-btn" id="reinContainer">
            <input type="checkbox" id="chkRein" onchange="handleStatusChange('chkRein')">
            <div class="btn-content">Account reinstated</div>
          </label>
          <label class="action-btn" id="pendContainer">
            <input type="checkbox" id="chkPend" onchange="handleStatusChange('chkPend')">
            <div class="btn-content">Account remains pending</div>
          </label>
          <label class="action-btn" id="suspContainer">
            <input type="checkbox" id="chkSusp" onchange="handleStatusChange('chkSusp')">
            <div class="btn-content">Account remains suspended</div>
          </label>
        </div>

        <div class="action-col-right" id="rightChecklistCol">
          <label class="action-btn" id="footContainer">
            <input type="checkbox" id="chkFoot">
            <div class="btn-content">Footer Added</div>
          </label>
        </div>
      </div>

      <div id="reinstatedMenu" class="status-menu menu-reinstated">
        <label class="action-btn"><input type="checkbox" id="chkDiscount"><div class="btn-content">$20 discount added</div></label>
        <label class="action-btn"><input type="checkbox" id="chkHighRisk"><div class="btn-content">High Risk Removed</div></label>
        <label class="action-btn"><input type="checkbox" id="chkLabel"><div class="btn-content">Transaction wrong label removed and correct label applied</div></label>
        <label class="action-btn"><input type="checkbox" id="chkFeedback"><div class="btn-content">Reinstated Account Feedback Form filled out</div></label>
        <label class="action-btn" style="grid-column: 1 / -1;"><input type="checkbox" id="chkConcessions"><div class="btn-content">Global Concessions Tracker filled out</div></label>
        <label class="action-btn" style="grid-column: 1 / -1;"><input type="checkbox" id="chkEmailAllowlist"><div class="btn-content">“Suspension Appeals — ALLOWLIST - Reactivating” email sent</div></label>
      </div>

      <div id="pendingMenu" class="status-menu menu-pending">
        <label class="action-btn" style="grid-column: 1 / -1;"><input type="checkbox" id="chkEmailOoch"><div class="btn-content">Contact unsuccessful/out of calling hours (OOCH) email sent</div></label>
      </div>

      <div id="suspendedMenu" class="status-menu menu-suspended">
        <label class="action-btn" style="grid-column: 1 / -1;"><input type="checkbox" id="chkFinalWordNew"><div class="btn-content">Final Word sent “Keeping Suspended- Do Not Reinstate”</div></label>
      </div>

      <div id="guideArrowContainer" class="guide-container hidden fade-in">
        <div class="guide-text">⚠️ REQUIRED STEP FOR REINSTATEMENTS ⚠️</div>
        <div class="guide-arrow">⬇</div>
      </div>

      <div class="additional-form-section hidden fade-in" id="feedbackFormSection">
        <div class="additional-form-text">Mandatory actions for reinstated accounts:</div>
        <div style="display: flex; gap: 15px; justify-content: center; flex-wrap: wrap;">
          <a href="https://docs.google.com/forms/d/e/1FAIpQLSdYmIMfpzNqO602q1-iRjK7JqZ5YSF74GgywvqzVQJdXqrNRg/viewform?pli=1&pli=1" target="_blank" class="additional-form-link">
            Reinstated Account Feedback Form
          </a>
          <a href="https://docs.google.com/forms/d/e/1FAIpQLSd74gRNvcu0_T7bShq0DyxiVOK2Y_NhuhUWCRn8xkTexI-dxA/viewform" target="_blank" class="concessions-form-link">
            Global Concessions Tracker
          </a>
        </div>
      </div>

      <div class="btn-row">
        <button class="btn-reset" onclick="triggerReset()">Reset</button>
        <button class="btn-gen" onclick="generateTemplate()">Generate Template</button>
      </div>

      <div id="outputContainer">
        <label style="font-size:16px;">Ready to Copy: <span style="font-weight:400; font-size:12px; opacity:0.6;">(Press Ctrl+Shift+C to copy)</span></label>
        <textarea id="finalText" readonly></textarea>
        <button class="btn-copy" id="btnCopy" onclick="copyToClipboard()">COPY TO CLIPBOARD</button>
      </div>

    </div>

    <div class="footer-author">
      Suspension Appeals Template | By Felix Martinez
    </div>

    <div id="leftScratchpad">
      <div class="scratchpad-tab" onclick="toggleScratchpad()">📝 Notes</div>
      <div class="sp-top-bar">
        <div class="sp-header-left">
          <button class="sp-menu-btn" onclick="toggleInnerMenu()" title="Toggle Menu">≡</button>
          <div class="scratchpad-title">Workspace</div>
        </div>
        <button id="btnSync" onclick="saveToCloud()">☁️ Save</button>
      </div>
      <div class="sp-main-layout">
        <div class="sp-left-nav" id="spSidebar">
          <button class="sp-nav-btn active" onclick="switchSpTab(1)">Note 1</button>
          <button class="sp-nav-btn" onclick="switchSpTab(2)">Note 2</button>
          <button class="sp-nav-btn" onclick="switchSpTab(3)">Note 3</button>
          <button class="sp-nav-btn" onclick="switchSpTab(4)">Note 4</button>
          <button class="sp-nav-btn" onclick="switchSpTab(5)">Note 5</button>
        </div>
        <div class="sp-content-area">
          <textarea id="spText1" class="sp-textarea active" placeholder="write something here..."></textarea>
          <textarea id="spText2" class="sp-textarea" placeholder="write something here..."></textarea>
          <textarea id="spText3" class="sp-textarea" placeholder="write something here..."></textarea>
          <textarea id="spText4" class="sp-textarea" placeholder="write something here..."></textarea>
          <textarea id="spText5" class="sp-textarea" placeholder="write something here..."></textarea>
        </div>
      </div>
    </div>

    <div id="cloudModal" class="modal-overlay">
      <div class="modal-box">
        <div class="modal-icon">☁️</div>
        <div class="modal-title">Emergency Cloud Sync</div>
        <div class="modal-text">Need to switch computers? Backup your entire case here, then restore it on the new machine.</div>
        <div style="display:flex; flex-direction:column; gap:10px;">
          <button class="btn-primary" id="btnBackupCloud" onclick="backupToCloud()">📤 Backup Case to Cloud</button>
          <button class="btn-primary" style="background-color: var(--accent-green);" id="btnRestoreCloud" onclick="restoreFromCloud()">📥 Restore Case from Cloud</button>
        </div>
        <button class="btn-reset" style="margin-top:15px; width:100%" onclick="closeCloudModal()">Cancel</button>
      </div>
    </div>

    <div id="metricsModal" class="modal-overlay">
      <div class="modal-box" style="max-width: 380px; padding: 25px;">
        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom: 2px solid var(--input-border); padding-bottom: 15px;">
          <h2 class="modal-title" style="margin:0; text-transform:none;">📊 Daily Metrics</h2>
          <button style="background:transparent; border:none; font-size:20px; cursor:pointer; color: var(--text-dark); opacity:0.6;" onclick="closeMetricsModal()">✖</button>
        </div>
        <div style="display:flex; justify-content: space-around; margin-top:25px; margin-bottom: 25px;">
            <div><h1 id="dashCases" style="color:var(--accent-green); margin:0; font-size:45px;">0</h1><p style="font-size:14px; font-weight:bold; margin-top:5px; opacity:0.8;">Cases Today</p></div>
            <div><h1 id="dashAht" style="color:var(--timer-orange); margin:0; font-size:45px;">00:00</h1><p style="font-size:14px; font-weight:bold; margin-top:5px; opacity:0.8;">Avg AHT</p></div>
        </div>
        <p style="font-size:15px; margin-top: 15px;">Target AHT: <b style="color:var(--remitly-blue);">12:24</b></p>
        <p style="font-size:11px; opacity:0.7; font-style:italic; margin-top:15px; line-height:1.4;">* Please note this is an estimated AHT calculated automatically by this tool from the moment you select a source until you generate a template.</p>
      </div>
    </div>

    <div id="settingsModal" class="modal-overlay">
      <div class="modal-box" style="max-width: 350px;">
        <div class="modal-title" style="text-transform:none; margin-bottom:20px;">⚙️ User Settings</div>
        
        <div class="switch-container">
          <label class="switch-label">✨ Animated Background</label>
          <label class="switch"><input type="checkbox" id="chkBgPref" onchange="toggleBgPreference()"><span class="slider"></span></label>
        </div>

        <div class="switch-container">
          <label class="switch-label">⏱️ Enable AHT Pacer</label>
          <label class="switch"><input type="checkbox" id="chkTimerPref" onchange="toggleTimerPreference()"><span class="slider"></span></label>
        </div>
        <p style="font-size:11px; opacity:0.7; text-align:left; margin-top:8px;">Displays a visual timer to help you track your case speed. Colors change at 5m, 9m, and 12m30s. Starts automatically when you select the suspension source.</p>
        <div class="modal-buttons" style="margin-top:25px;"><button class="btn-primary" onclick="closeSettingsModal()">Done</button></div>
      </div>
    </div>

    <div id="resetModal" class="modal-overlay">
      <div class="modal-box">
        <div class="modal-icon">⚠️</div>
        <div class="modal-title">Clear all data?</div>
        <div class="modal-text">Are you sure you want to clear the form? This action cannot be undone and you will lose your current progress.</div>
        <div class="modal-buttons">
          <button class="btn-reset" onclick="closeResetModal()">Cancel</button>
          <button class="btn-danger" onclick="confirmReset()">Yes, Clear Data</button>
        </div>
      </div>
    </div>

    <div id="errorModal" class="modal-overlay">
      <div class="modal-box">
        <div class="modal-icon">🛑</div>
        <div class="modal-title">Missing Information</div>
        <div class="modal-text" id="errorMessage">Please fill out the required fields.</div>
        <div class="modal-buttons"><button class="btn-primary" onclick="closeErrorModal()" id="btnErrorClose">Got it</button></div>
      </div>
    </div>

    <div id="historyModal" class="modal-overlay">
      <div class="modal-box" style="max-width: 500px; padding: 25px;">
        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom: 2px solid var(--input-border); padding-bottom: 15px;">
          <h2 class="modal-title" style="margin:0; text-transform:none;">🕒 Recent Cases Vault</h2>
          <button style="background:transparent; border:none; font-size:20px; cursor:pointer; color: var(--text-dark); opacity:0.6;" onclick="closeHistoryModal()">✖</button>
        </div>
        <p style="font-size:12px; opacity:0.8; margin-top:15px; text-align:left;">Did you lose your text? Don't worry. Here are your last 5 generated templates:</p>
        <div id="historyList" class="history-list-container"></div>
      </div>
    </div>

    <script>
      const hideFwSources = [
        "Account already worked by another associate", 
        "The account is not suspended, it is deleted.", 
        "The account is not suspended, it is active.",
        "***Account suspended by the SMB Team***"
      ];
      
      const formElementsForSave = ["sourceSelect", "kountType", "manualType", "kountDetailText", "manualDetailText", "cxInfo", "siftScore", "paymentMethod", "logins", "phone", "email", "accountStatus", "dupAccounts", "additionalFindings", "actionsText", "chkFoot", "chkRein", "chkPend", "chkSusp", "chkDiscount", "chkHighRisk", "chkLabel", "chkFeedback", "chkConcessions", "chkEmailAllowlist", "chkEmailOoch", "chkFinalWordNew"];

      let caseStartTime = null; let timerInterval = null; let timerSeconds = 0;
      let totalCasesProcessed = 0; let totalSecondsAllCases = 0;
      let processedCasesMemory = new Set();
      
      // NUEVO: Bandera para evitar el spam en el Tracker durante un mismo caso
      let interactionLogged = false;

      // ==========================================
      // --- SISTEMA DE AUTO-GUARDADO (DEBOUNCE) ---
      // ==========================================
      let autoSaveTimer;
      
      function triggerAutoSave() {
        let statusIndicator = document.getElementById('autoSaveStatus');
        statusIndicator.innerText = "⏳ Saving...";
        statusIndicator.style.opacity = "1";
        
        clearTimeout(autoSaveTimer);
        autoSaveTimer = setTimeout(() => {
            saveState(); 
            statusIndicator.innerText = "✔ Draft saved";
            setTimeout(() => {
                statusIndicator.style.opacity = "0";
            }, 2500); 
        }, 1000); 
      }

      function startAHTTimer() {
        var sourceVal = document.getElementById('sourceSelect').value;
        if (!caseStartTime && sourceVal !== "") { 
          caseStartTime = new Date(); timerSeconds = 0;
          if (localStorage.getItem('remitlyTemplate_showTimer') === 'true') { document.getElementById('ahtTimerDisplay').style.display = 'flex'; }
          if(!timerInterval) { timerInterval = setInterval(updateTimerDisplay, 1000); }
        }
      }

      function updateTimerDisplay() {
        if (!caseStartTime) return;
        let now = new Date(); timerSeconds = Math.floor((now - caseStartTime) / 1000);
        let m = Math.floor(timerSeconds / 60).toString().padStart(2, '0'); let s = (timerSeconds % 60).toString().padStart(2, '0');
        let timerText = document.getElementById('ahtTimerText');
        timerText.innerText = `${m}:${s}`;
        if (timerSeconds >= 750) { timerText.className = 'aht-timer timer-red'; } else if (timerSeconds >= 540) { timerText.className = 'aht-timer timer-orange'; } else if (timerSeconds >= 300) { timerText.className = 'aht-timer timer-yellow'; } else { timerText.className = 'aht-timer timer-green'; }
      }

      function stopTimer() { if(timerInterval) { clearInterval(timerInterval); timerInterval = null; } }
      function resetTimerVisuals() { stopTimer(); caseStartTime = null; let timerText = document.getElementById('ahtTimerText'); timerText.innerText = "00:00"; timerText.className = 'aht-timer timer-green'; document.getElementById('ahtTimerDisplay').style.display = 'none'; }

      function openCloudModal() { document.getElementById('cloudModal').style.display = 'flex'; }
      function closeCloudModal() { document.getElementById('cloudModal').style.display = 'none'; document.getElementById('btnBackupCloud').innerText = "📤 Backup Case to Cloud"; document.getElementById('btnRestoreCloud').innerText = "📥 Restore Case from Cloud"; }
      
      function openSettingsModal() { document.getElementById('settingsModal').style.display = 'flex'; }
      function closeSettingsModal() { document.getElementById('settingsModal').style.display = 'none'; }
      function openMetricsModal() { document.getElementById('metricsModal').style.display = 'flex'; }
      function closeMetricsModal() { document.getElementById('metricsModal').style.display = 'none'; }

      function toggleTimerPreference() {
        let isChecked = document.getElementById('chkTimerPref').checked; localStorage.setItem('remitlyTemplate_showTimer', isChecked);
        if(caseStartTime) { document.getElementById('ahtTimerDisplay').style.display = isChecked ? 'flex' : 'none'; }
      }

      function toggleBgPreference() {
        let isChecked = document.getElementById('chkBgPref').checked;
        localStorage.setItem('remitlyTemplate_animatedBg', isChecked);
        if (isChecked) { document.body.classList.add('animated-bg'); } 
        else { document.body.classList.remove('animated-bg'); }
      }

      window.onload = function() {
        loadTheme(); loadState(); 

        let showTimerPref = localStorage.getItem('remitlyTemplate_showTimer');
        if (showTimerPref === null) { localStorage.setItem('remitlyTemplate_showTimer', 'true'); document.getElementById('chkTimerPref').checked = true; } else { document.getElementById('chkTimerPref').checked = (showTimerPref === 'true'); }
        
        let showBgPref = localStorage.getItem('remitlyTemplate_animatedBg');
        if (showBgPref === 'true') { document.getElementById('chkBgPref').checked = true; document.body.classList.add('animated-bg'); } else { document.getElementById('chkBgPref').checked = false; document.body.classList.remove('animated-bg'); }

        if(localStorage.getItem('remitlyTemplate_spMenuHidden') === 'true') { document.getElementById('spSidebar').classList.add('collapsed'); }

        const container = document.querySelector('.container');
        container.addEventListener('input', startAHTTimer); container.addEventListener('change', startAHTTimer);
        document.getElementById('sourceSelect').addEventListener('change', startAHTTimer);

        formElementsForSave.forEach(id => { 
          const el = document.getElementById(id); 
          if (el) { el.addEventListener('input', triggerAutoSave); el.addEventListener('change', triggerAutoSave); } 
        });
        
        document.addEventListener('click', function(event) {
          const panel = document.getElementById('leftScratchpad');
          if (panel.classList.contains('open') && !event.target.closest('#leftScratchpad')) { panel.classList.remove('open'); }
        });

        var btnSync = document.getElementById('btnSync'); btnSync.innerText = "⏳ Load...";
        try {
          google.script.run.withSuccessHandler(function(cloudData) {
            btnSync.innerText = "☁️ Save";
            if(cloudData) {
              document.getElementById('spText1').value = cloudData.text1 || ''; document.getElementById('spText2').value = cloudData.text2 || ''; document.getElementById('spText3').value = cloudData.text3 || ''; document.getElementById('spText4').value = cloudData.text4 || ''; document.getElementById('spText5').value = cloudData.text5 || '';
              if(cloudData.width && cloudData.height) { document.querySelectorAll('.sp-textarea').forEach(ta => { ta.style.width = cloudData.width; ta.style.height = cloudData.height; }); }
            }
          }).getUserData();
        } catch(e) { btnSync.innerText = "☁️ Save"; } 
      };

      function toggleInnerMenu() { const sidebar = document.getElementById('spSidebar'); sidebar.classList.toggle('collapsed'); localStorage.setItem('remitlyTemplate_spMenuHidden', sidebar.classList.contains('collapsed')); }
      function toggleScratchpad() { document.getElementById('leftScratchpad').classList.toggle('open'); }
      function switchSpTab(index) { document.querySelectorAll('.sp-nav-btn').forEach((btn, i) => { i + 1 === index ? btn.classList.add('active') : btn.classList.remove('active'); }); document.querySelectorAll('.sp-textarea').forEach((ta, i) => { i + 1 === index ? ta.classList.add('active') : ta.classList.remove('active'); }); }

      function getFullPayload() {
        var activeTa = document.querySelector('.sp-textarea.active');
        let formDataObj = {};
        formElementsForSave.forEach(id => {
            let el = document.getElementById(id);
            if(el) formDataObj[id] = el.type === 'checkbox' ? el.checked : el.value;
        });
        return { 
          width: activeTa ? activeTa.style.width : '320px', 
          height: activeTa ? activeTa.style.height : '400px', 
          text1: document.getElementById('spText1').value, 
          text2: document.getElementById('spText2').value, 
          text3: document.getElementById('spText3').value, 
          text4: document.getElementById('spText4').value, 
          text5: document.getElementById('spText5').value,
          formData: JSON.stringify(formDataObj)
        };
      }

      function saveToCloud() {
        var btn = document.getElementById('btnSync'); btn.innerText = "⏳ Saving..."; btn.style.opacity = "0.7";
        try {
          google.script.run.withSuccessHandler(function(success) {
            if(success) { btn.innerText = "✅ Saved!"; btn.style.backgroundColor = "var(--accent-green)"; btn.style.color = "white"; btn.style.opacity = "1"; setTimeout(() => { btn.innerText = "☁️ Save"; btn.style.backgroundColor = "transparent"; btn.style.color = "var(--text-dark)"; btn.style.opacity = "0.7"; btn.style.borderColor = "var(--input-border)"; }, 3000); } else { btn.innerText = "❌ Error"; }
          }).saveUserData(getFullPayload());
        } catch(e) { btn.innerText = "❌ Error"; }
      }

      function backupToCloud() {
        var btn = document.getElementById('btnBackupCloud'); btn.innerText = "⏳ Backing up..."; 
        try {
          google.script.run.withSuccessHandler(function(success) {
            if(success) { btn.innerText = "✅ Backup Successful!"; setTimeout(closeCloudModal, 2000); } else { btn.innerText = "❌ Error"; }
          }).saveUserData(getFullPayload());
        } catch(e) { btn.innerText = "❌ Error"; }
      }

      function restoreFromCloud() {
        var btn = document.getElementById('btnRestoreCloud'); btn.innerText = "⏳ Restoring...";
        try {
          google.script.run.withSuccessHandler(function(cloudData) {
            if(cloudData && cloudData.formData) {
              let formDataObj = JSON.parse(cloudData.formData);
              formElementsForSave.forEach(id => { 
                  let el = document.getElementById(id); 
                  if(el && formDataObj[id] !== undefined) {
                      if (el.type === 'checkbox') el.checked = formDataObj[id];
                      else el.value = formDataObj[id];
                  }
              });
              saveState(); handleSourceChange(); handleStatusChange(null);
              
              document.getElementById('spText1').value = cloudData.text1 || ''; document.getElementById('spText2').value = cloudData.text2 || ''; document.getElementById('spText3').value = cloudData.text3 || ''; document.getElementById('spText4').value = cloudData.text4 || ''; document.getElementById('spText5').value = cloudData.text5 || '';

              btn.innerText = "✅ Restored!"; setTimeout(closeCloudModal, 2000);
            } else {
              btn.innerText = "❌ No Backup Found";
            }
          }).getUserData();
        } catch(e) { btn.innerText = "❌ Error"; }
      }

      function toggleTheme() {
        const html = document.documentElement; const newTheme = html.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
        html.setAttribute('data-theme', newTheme); localStorage.setItem('remitlyTemplate_theme', newTheme); document.querySelector('.theme-switch').innerText = newTheme === 'dark' ? '☀️' : '🌙';
      }

      function loadTheme() { if (localStorage.getItem('remitlyTemplate_theme') === 'dark') { document.documentElement.setAttribute('data-theme', 'dark'); document.querySelector('.theme-switch').innerText = '☀️'; } }

      document.addEventListener('keydown', function(event) {
        if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') generateTemplate();
        if ((event.ctrlKey || event.metaKey) && event.shiftKey && (event.key === 'c' || event.key === 'C')) { if (document.getElementById("outputContainer").style.display === "block") { copyToClipboard(); event.preventDefault(); } }
      });

      function saveState() { formElementsForSave.forEach(id => { const el = document.getElementById(id); if (el) localStorage.setItem("remitlyTemplate_" + id, el.type === 'checkbox' ? el.checked : el.value); }); }
      
      function loadState() {
        formElementsForSave.forEach(id => { const el = document.getElementById(id); if (el) { const val = localStorage.getItem("remitlyTemplate_" + id); if (val !== null) el.type === 'checkbox' ? el.checked = (val === 'true') : el.value = val; } });
        handleSourceChange(); handleStatusChange(null);
      }

      function clearState() { formElementsForSave.forEach(id => localStorage.removeItem("remitlyTemplate_" + id)); }
      function triggerReset() { document.getElementById('resetModal').style.display = 'flex'; }
      function closeResetModal() { document.getElementById('resetModal').style.display = 'none'; }
      function showError(message, focusElementId) { document.getElementById('errorMessage').innerText = message; document.getElementById('errorModal').style.display = 'flex'; document.getElementById('btnErrorClose').onclick = function() { closeErrorModal(); if(focusElementId) document.getElementById(focusElementId).focus(); }; }
      function closeErrorModal() { document.getElementById('errorModal').style.display = 'none'; }

      function confirmReset() {
        closeResetModal(); 
        document.querySelectorAll("input[type=text], textarea:not(.sp-textarea)").forEach(function(input) { input.value = ""; });
        document.querySelectorAll("select").forEach(function(select) { select.selectedIndex = 0; });
        document.querySelectorAll("input[type=checkbox]").forEach(function(chk) { chk.checked = false; chk.disabled = false; });

        document.getElementById("kountOptions").classList.add("hidden"); document.getElementById("manualOptions").classList.add("hidden"); document.getElementById("kountDetailsSection").classList.add("hidden"); document.getElementById("manualDetailsSection").classList.add("hidden"); document.getElementById("fullReviewContainer").classList.remove("hidden"); 

        let colLeft = document.querySelector('.action-col-left');
        let colRight = document.querySelector('.action-col-right');
        colLeft.style.display = 'flex';
        colRight.style.display = 'flex';
        colRight.style.flexDirection = 'column';
        colRight.style.gridTemplateColumns = '';
        
        document.getElementById("footContainer").style.display = '';
        document.getElementById("reinContainer").style.display = ''; document.getElementById("pendContainer").style.display = ''; document.getElementById("suspContainer").style.display = '';
        
        // Reiniciamos el escudo anti-spam
        interactionLogged = false;

        handleStatusChange(null);
        handleSourceChange();

        document.getElementById("outputContainer").style.display = "none"; document.getElementById("finalText").style.display = "none"; document.getElementById("btnCopy").style.display = "none";
        clearState(); 
        let sourceSelect = document.getElementById("sourceSelect"); sourceSelect.value = ""; resetTimerVisuals();
        window.scrollTo({ top: 0, behavior: 'smooth' });
      }

      function handleStatusChange(clickedId) {
        if (clickedId === 'chkRein' && document.getElementById('chkRein').checked) { document.getElementById('chkPend').checked = false; document.getElementById('chkSusp').checked = false; } 
        else if (clickedId === 'chkPend' && document.getElementById('chkPend').checked) { document.getElementById('chkRein').checked = false; document.getElementById('chkSusp').checked = false; } 
        else if (clickedId === 'chkSusp' && document.getElementById('chkSusp').checked) { document.getElementById('chkRein').checked = false; document.getElementById('chkPend').checked = false; }
        
        let isRein = document.getElementById('chkRein').checked;
        let isPend = document.getElementById('chkPend').checked;
        let isSusp = document.getElementById('chkSusp').checked;

        let reinC = document.getElementById("reinContainer");
        let pendC = document.getElementById("pendContainer");
        let suspC = document.getElementById("suspContainer");

        var source = document.getElementById("sourceSelect").value;
        let isException = hideFwSources.includes(source);

        if (!isException) {
            if (isRein) {
                pendC.style.display = 'none'; suspC.style.display = 'none'; reinC.style.display = '';
                document.getElementById("chkEmailAllowlist").checked = true;
                document.getElementById("chkEmailAllowlist").disabled = true; 
            } else if (isPend) {
                reinC.style.display = 'none'; suspC.style.display = 'none'; pendC.style.display = '';
                document.getElementById("chkEmailOoch").checked = true;
                document.getElementById("chkEmailOoch").disabled = true;
            } else if (isSusp) {
                reinC.style.display = 'none'; pendC.style.display = 'none'; suspC.style.display = '';
                document.getElementById("chkFinalWordNew").checked = true;
                document.getElementById("chkFinalWordNew").disabled = true;
            } else {
                reinC.style.display = ''; pendC.style.display = ''; suspC.style.display = '';
            }
        }

        document.getElementById('reinstatedMenu').style.display = isRein ? 'grid' : 'none';
        document.getElementById('pendingMenu').style.display = isPend ? 'grid' : 'none';
        document.getElementById('suspendedMenu').style.display = isSusp ? 'grid' : 'none';

        if(isRein) {
            document.getElementById("guideArrowContainer").classList.remove("hidden"); document.getElementById("feedbackFormSection").classList.remove("hidden");
            if(clickedId === 'chkRein') {
                setTimeout(() => { document.getElementById("feedbackFormSection").scrollIntoView({ behavior: 'smooth', block: 'center' }); }, 150);
            }
        } else {
            document.getElementById("guideArrowContainer").classList.add("hidden"); document.getElementById("feedbackFormSection").classList.add("hidden");
            document.getElementById("chkDiscount").checked = false; document.getElementById("chkHighRisk").checked = false; document.getElementById("chkLabel").checked = false; 
            document.getElementById("chkFeedback").checked = false; document.getElementById("chkConcessions").checked = false;
            document.getElementById("chkEmailAllowlist").checked = false; document.getElementById("chkEmailAllowlist").disabled = false;
        }

        if(!isPend) { document.getElementById("chkEmailOoch").checked = false; document.getElementById("chkEmailOoch").disabled = false; }
        if(!isSusp) { document.getElementById("chkFinalWordNew").checked = false; document.getElementById("chkFinalWordNew").disabled = false; }

        if (clickedId) triggerAutoSave(); 
      }

      function handleSourceChange() {
        var source = document.getElementById("sourceSelect").value;
        var kOpt = document.getElementById("kountOptions"); var mOpt = document.getElementById("manualOptions"); var kDet = document.getElementById("kountDetailsSection"); var mDet = document.getElementById("manualDetailsSection"); var reviewContainer = document.getElementById("fullReviewContainer");
        var footC = document.getElementById("footContainer");
        
        let colLeft = document.querySelector('.action-col-left');
        let colRight = document.querySelector('.action-col-right');

        colLeft.style.display = 'flex';
        colRight.style.display = 'flex';
        colRight.style.gridTemplateColumns = '';
        
        footC.style.display = ''; 
        
        kOpt.classList.add("hidden"); mOpt.classList.add("hidden"); kDet.classList.add("hidden"); mDet.classList.add("hidden");

        // ===============================================
        // SOLUCIÓN DE INTERACCIÓN INSTANTÁNEA
        // ===============================================
        if (source !== "" && !interactionLogged) {
            interactionLogged = true; // Activa el escudo para no mandar más pings en este caso
            let timeStamp = new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit', second:'2-digit'});
            let trackerId = "Interaction Started - " + timeStamp;
            
            try { 
                google.script.run.logUsage(trackerId, source, 0); // 0 segundos de AHT inicial
            } catch(error) { 
                console.log("Tracker ping error: ", error); 
            }
        }

        if (source === "") { reviewContainer.classList.remove("hidden"); handleStatusChange(null); return; }

        if (hideFwSources.includes(source)) {
          reviewContainer.classList.add("hidden");
          colLeft.style.display = 'none';
          document.getElementById("chkRein").checked = false; document.getElementById("chkPend").checked = false; document.getElementById("chkSusp").checked = false;

          if (source === "Account already worked by another associate" || source === "***Account suspended by the SMB Team***") {
             footC.style.display = 'none'; document.getElementById("chkFoot").checked = false;
          } else if (source === "The account is not suspended, it is deleted." || source === "The account is not suspended, it is active.") {
             footC.style.display = 'none'; document.getElementById("chkFoot").checked = false;
          }
        } else {
          reviewContainer.classList.remove("hidden");
          if (source === "Kount") { kOpt.classList.remove("hidden"); kDet.classList.remove("hidden"); } else if (source === "Manual") { mOpt.classList.remove("hidden"); mDet.classList.remove("hidden"); }
        }
        handleStatusChange(null);
      }

      function saveToSessionHistory(caseTitle, templateText) {
        let history = JSON.parse(localStorage.getItem('remitlyTemplate_history') || '[]');
        const now = new Date(); const timeString = now.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}); 
        const existingIndex = history.findIndex(item => item.sfCase === caseTitle); if (existingIndex !== -1) { history.splice(existingIndex, 1); }
        history.unshift({ sfCase: caseTitle, time: timeString, text: templateText });
        if(history.length > 5) { history.pop(); } localStorage.setItem('remitlyTemplate_history', JSON.stringify(history));
      }

      function openHistoryModal() {
        const listContainer = document.getElementById('historyList'); let history = JSON.parse(localStorage.getItem('remitlyTemplate_history') || '[]');
        if(history.length === 0) { listContainer.innerHTML = '<div style="text-align:center; padding:30px; opacity:0.6; font-size:14px;">No recent templates found yet.<br><br>Generate a template and it will appear here.</div>'; } else { listContainer.innerHTML = ''; history.forEach((item, index) => { listContainer.innerHTML += `<div class="history-item"><div class="history-info"><div class="history-case">${item.sfCase}</div><div class="history-time">Generated at ${item.time}</div></div><button class="history-copy-btn" id="histBtn${index}" onclick="copyHistoryItem(${index})">COPY</button></div>`; }); }
        document.getElementById('historyModal').style.display = 'flex';
      }

      function closeHistoryModal() { document.getElementById('historyModal').style.display = 'none'; }
      function copyHistoryItem(index) {
        let history = JSON.parse(localStorage.getItem('remitlyTemplate_history') || '[]');
        if(history[index]) { const tempTextArea = document.createElement("textarea"); tempTextArea.value = history[index].text; document.body.appendChild(tempTextArea); tempTextArea.select(); document.execCommand("copy"); document.body.removeChild(tempTextArea); let btn = document.getElementById('histBtn' + index); btn.innerText = "COPIED! ✅"; btn.style.backgroundColor = "#10b981"; btn.style.color = "white"; setTimeout(() => { btn.innerText = "COPY"; btn.style.backgroundColor = "var(--accent-green)"; }, 2000); }
      }

      function updateMetricsDashboard() {
        document.getElementById('dashCases').innerText = totalCasesProcessed;
        if(totalCasesProcessed > 0) { let avgSeconds = Math.floor(totalSecondsAllCases / totalCasesProcessed); let m = Math.floor(avgSeconds / 60).toString().padStart(2, '0'); let s = (avgSeconds % 60).toString().padStart(2, '0'); document.getElementById('dashAht').innerText = `${m}:${s}`; }
      }

      function generateTemplate() {
        var source = document.getElementById("sourceSelect").value;
        if(!source) { showError("Please select an option in the 'Findings' section first.", "sourceSelect"); return; }

        var subType = ""; var detailsContent = ""; 
        if (!hideFwSources.includes(source)) {
           if (source === "Kount") { subType = document.getElementById("kountType").value; var kText = document.getElementById("kountDetailText").value.trim(); detailsContent = `KOUNT\n\n${kText}`; } 
           else if (source === "Manual") { subType = document.getElementById("manualType").value; var mText = document.getElementById("manualDetailText").value.trim(); detailsContent = `SUSPENSION REASON\n\n${mText}`; }
        }

        var cxInfo = document.getElementById("cxInfo").value.trim(); var sift = document.getElementById("siftScore").value.trim(); var pm = document.getElementById("paymentMethod").value.trim(); var logins = document.getElementById("logins").value.trim(); var phone = document.getElementById("phone").value.trim(); var email = document.getElementById("email").value.trim(); var status = document.getElementById("accountStatus").value.trim(); var dups = document.getElementById("dupAccounts").value.trim(); var addFindings = document.getElementById("additionalFindings").value.trim(); var actions = document.getElementById("actionsText").value.trim();

        // LÓGICA DE MEMORIA INTELIGENTE ADAPTATIVA (Sin SF Case)
        var caseSignature = source + subType + cxInfo + addFindings;
        if (caseSignature.trim() === "") { caseSignature = new Date().getTime().toString(); }

        if (!processedCasesMemory.has(caseSignature)) {
            processedCasesMemory.add(caseSignature);
            totalCasesProcessed++;
            if(caseStartTime) { 
                let caseSeconds = Math.floor((new Date() - caseStartTime) / 1000); 
                totalSecondsAllCases += caseSeconds; 
            }
            updateMetricsDashboard();
        }

        var template = `Suspension Appeals\n\nNote: Working on CSP SA\n\nFindings\n\nWAS THE ACCOUNT SUSPENDED BY KOUNT OR MANUALLY? ${source}`;

        if (!hideFwSources.includes(source)) {
             template += `\n\nTYPE: ${subType}\n\n${detailsContent}`;
             if (cxInfo) template += `\n\nPLEASE SHARE ANY INFO PROVIDED BY THE CX IF AVAILABLE\n\n${cxInfo}`;
             
             template += `\n\nAccount's Review\n\n`;
             
             let reviewItems = [];
             if (sift) reviewItems.push(`SIFT SCORE: ${sift}`); 
             if (pm) reviewItems.push(`PAYMENT METHOD: ${pm}`); 
             if (logins) reviewItems.push(`LOGINS: ${logins}`); 
             if (phone) reviewItems.push(`PHONE: ${phone}`); 
             if (email) reviewItems.push(`EMAIL: ${email}`); 
             if (status) reviewItems.push(`ACCOUNT STATUS: ${status}`); 
             if (dups) reviewItems.push(`DUP ACCOUNTS: ${dups}`);
             
             if (reviewItems.length > 0) {
                 template += reviewItems.join('\n');
             }

             if (addFindings) template += `\n\nADDITIONAL FINDINGS:\n\n${addFindings}`;
        }

        if (actions) { 
            template += `\n\nACTION TAKEN\n\n${actions}`; 
        } else { 
            template += `\n\nACTION TAKEN\n\nN/A`; 
        }

        if (!hideFwSources.includes(source)) {
             template += `\n\n`;
             if (document.getElementById("chkRein").checked) {
                 template += `- Account Reinstated: Yes\n`;
                 if (document.getElementById("chkDiscount").checked) template += `- $20 discount added: Yes\n`; 
                 if (document.getElementById("chkHighRisk").checked) template += `- High Risk Removed: Yes\n`; 
                 if (document.getElementById("chkLabel").checked) template += `- Transaction wrong label removed and correct label applied: Yes\n`; 
                 if (document.getElementById("chkFeedback").checked) template += `- Reinstated Account Feedback Form filled out: Yes\n`;
                 if (document.getElementById("chkConcessions").checked) template += `- Global Concessions Tracker filled out: Yes\n`;
                 if (document.getElementById("chkEmailAllowlist").checked) template += `- “Suspension Appeals — ALLOWLIST - Reactivating” email sent: Yes\n`;
             } else if (document.getElementById("chkPend").checked) {
                 template += `- Account Remains Pending: Yes\n`;
                 if (document.getElementById("chkEmailOoch").checked) template += `- Contact unsuccessful/out of calling hours (OOCH) email sent: Yes\n`;
             } else if (document.getElementById("chkSusp").checked) {
                 template += `- Account Remains Suspended: Yes\n`;
                 if (document.getElementById("chkFinalWordNew").checked) template += `- Final Word sent “Keeping Suspended- Do Not Reinstate”: Yes\n`;
             }

             if (document.getElementById("chkFoot").checked) template += `- Footer Added: Yes\n`;
        }

        template = template.replace(/\n{3,}/g, '\n\n').trim();
        document.getElementById("finalText").value = template;
        
        let caseNameForHistory = source + (subType ? " - " + subType : "");
        saveToSessionHistory(caseNameForHistory, template);
        stopTimer();
        
        document.getElementById("outputContainer").style.display = "block"; document.getElementById("finalText").style.display = "block"; document.getElementById("btnCopy").style.display = "block";
        document.getElementById("outputContainer").scrollIntoView({behavior: "smooth"});

        let ahtSeconds = caseStartTime ? Math.round((new Date() - caseStartTime) / 1000) : 0;
        
        let timeStamp = new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit', second:'2-digit'});
        let trackerId = "Template Generated - " + timeStamp;
        try { google.script.run.logUsage(trackerId, source, ahtSeconds); } catch(error) {}
      }

      function copyToClipboard() {
        var copyText = document.getElementById("finalText"); copyText.select(); document.execCommand("copy");
        var btn = document.getElementById("btnCopy"); btn.innerText = "COPIED! ✅"; btn.style.backgroundColor = "#10b981"; setTimeout(() => { btn.innerText = "COPY TO CLIPBOARD"; btn.style.backgroundColor = "var(--accent-green)"; }, 2000);
      }
    </script>
  </body>
</html>
