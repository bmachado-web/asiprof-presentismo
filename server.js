/**
 * PresenQR — Backend (Node.js + Express)
 * Maneja OAuth2 con Google y escribe en Google Sheets API v4
 */

const express = require('express');
const { google } = require('googleapis');
const session = require('express-session');
const path = require('path');
require('dotenv').config();

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Sesión ──────────────────────────────────────────────────
app.use(session({
  secret: process.env.SESSION_SECRET || 'presenqr-secret-cambiame',
  resave: false,
  saveUninitialized: false,
  cookie: { secure: false, maxAge: 24 * 60 * 60 * 1000 }
}));

// ── OAuth2 client ────────────────────────────────────────────
function getOAuth2Client() {
  return new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );
}

// ── Middleware: verificar autenticación ──────────────────────
function requireAuth(req, res, next) {
  if (!req.session.tokens) {
    return res.status(401).json({ error: 'No autenticado', redirect: '/auth/login' });
  }
  next();
}

// ────────────────────────────────────────────────────────────
// AUTH ROUTES
// ────────────────────────────────────────────────────────────
app.get('/auth/login', (req, res) => {
  const oauth2Client = getOAuth2Client();
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/userinfo.email'
    ]
  });
  res.redirect(url);
});

app.get('/auth/callback', async (req, res) => {
  const { code, error } = req.query;
  if (error) return res.redirect('/?auth_error=' + encodeURIComponent(error));
  try {
    const oauth2Client = getOAuth2Client();
    const { tokens } = await oauth2Client.getToken(code);
    req.session.tokens = tokens;
    global.tokensMaestros = tokens; // Guardamos la llave maestra para los celulares

    oauth2Client.setCredentials(tokens);
    const oauth2 = google.oauth2({ version: 'v2', auth: oauth2Client });
    const { data } = await oauth2.userinfo.get();
    req.session.user = { email: data.email, name: data.name };

    res.redirect('/?auth=ok');
  } catch (err) {
    console.error('Auth error:', err);
    res.redirect('/?auth_error=callback_failed');
  }
});

app.post('/auth/logout', (req, res) => {
  req.session.destroy();
  res.json({ ok: true });
});

app.get('/auth/status', (req, res) => {
  res.json({
    authenticated: !!req.session.tokens,
    user: req.session.user || null
  });
});

// ────────────────────────────────────────────────────────────
// SHEETS ROUTES
// ────────────────────────────────────────────────────────────
function getSheetsClient(tokens) {
  const oauth2Client = getOAuth2Client();
  oauth2Client.setCredentials(tokens);
  oauth2Client.on('tokens', (newTokens) => {
    if (newTokens.refresh_token) tokens.refresh_token = newTokens.refresh_token;
    tokens.access_token = newTokens.access_token;
    tokens.expiry_date = newTokens.expiry_date;
  });
  return { auth: oauth2Client, sheets: google.sheets({ version: 'v4', auth: oauth2Client }) };
}

app.post('/api/sheets/create', requireAuth, async (req, res) => {
  try {
    const { titulo } = req.body;
    const { sheets } = getSheetsClient(req.session.tokens);

    const response = await sheets.spreadsheets.create({
      requestBody: {
        properties: { title: titulo || 'ASIPROF — Asistencia Guardias' },
        sheets: [
          { properties: { title: 'Registros', sheetId: 0 } },
          { properties: { title: 'Historial', sheetId: 1 } }
        ]
      }
    });

    const spreadsheetId = response.data.spreadsheetId;

    // Encabezados adaptados para ASIPROF SEGURIDAD
    const encabezados = [['Fecha', 'Legajo', 'Apellido y Nombre', 'Ubicación / Objetivo', 'Hora Inicio (Declarada)', 'Hora Escaneo (Real)', 'Horas a Cumplir', 'Geo OK', 'Distancia (m)', 'Coordenadas GPS']];
    
    await sheets.spreadsheets.values.update({
      spreadsheetId, range: 'Registros!A1', valueInputOption: 'USER_ENTERED', requestBody: { values: encabezados }
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId, range: 'Historial!A1', valueInputOption: 'USER_ENTERED', requestBody: { values: encabezados }
    });

    // Formato de encabezados
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [0, 1].map(sheetId => ({
          repeatCell: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 1 },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.06, green: 0.47, blue: 0.63 },
                textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true, fontSize: 11 }
              }
            },
            fields: 'userEnteredFormat(backgroundColor,textFormat)'
          }
        }))
      }
    });

    req.session.spreadsheetId = spreadsheetId;
    res.json({ ok: true, spreadsheetId, url: `https://docs.google.com/spreadsheets/d/${spreadsheetId}` });
  } catch (err) {
    console.error('Create sheet error:', err);
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/sheets/link', requireAuth, (req, res) => {
  const { spreadsheetId } = req.body;
  if (!spreadsheetId) return res.status(400).json({ error: 'spreadsheetId requerido' });
  req.session.spreadsheetId = spreadsheetId;
  res.json({ ok: true });
});

// Agregar un solo check-in (llamado en tiempo real cuando el empleado escanea)
app.post('/api/checkin', async (req, res) => {
  const { spreadsheetId, tokens, registro } = req.body;
  const sessionTokens = req.session?.tokens || tokens || global.tokensMaestros;
  if (!sessionTokens || !spreadsheetId) return res.status(400).json({ error: 'Faltan datos' });

  try {
    const { sheets } = getSheetsClient(sessionTokens);
    
    // Mapeo exacto de los datos de ASIPROF a las columnas
    const row = [
      registro.fecha, 
      registro.legajo, 
      registro.nombre, 
      registro.ubicacion, 
      registro.horaInicioDeclarada, 
      registro.horaEscaneoReal, 
      registro.horasCumplir, 
      registro.geoOk,
      registro.distancia, 
      `${registro.lat}, ${registro.lng}`
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: 'Historial!A1',
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [row] }
    });
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/sheets/status', requireAuth, (req, res) => {
  res.json({
    spreadsheetId: req.session.spreadsheetId || null,
    url: req.session.spreadsheetId
      ? `https://docs.google.com/spreadsheets/d/${req.session.spreadsheetId}`
      : null
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`PresenQR corriendo en http://localhost:${PORT}`));