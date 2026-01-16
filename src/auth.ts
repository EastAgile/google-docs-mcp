// src/auth.ts
import { google } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import { JWT } from 'google-auth-library'; // ADDED: Import for Service Account client
import * as fs from 'fs/promises';
import * as path from 'path';
import * as http from 'http';
import { fileURLToPath } from 'url';
import { exec } from 'child_process';
import { promisify } from 'util';

const execAsync = promisify(exec);

// --- Calculate paths relative to this script file (ESM way) ---
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRootDir = path.resolve(__dirname, '..');

const TOKEN_PATH = path.join(projectRootDir, 'token.json');
const CREDENTIALS_PATH = path.join(projectRootDir, 'credentials.json');
// --- End of path calculation ---

const SCOPES = [
  'https://www.googleapis.com/auth/documents',
  'https://www.googleapis.com/auth/drive', // Full Drive access for listing, searching, and document discovery
  'https://www.googleapis.com/auth/spreadsheets' // Google Sheets API access
];

// --- NEW FUNCTION: Handles Service Account Authentication ---
// This entire function is new. It is called only when the
// SERVICE_ACCOUNT_PATH environment variable is set.
// Supports domain-wide delegation via GOOGLE_IMPERSONATE_USER env var.
async function authorizeWithServiceAccount(): Promise<JWT> {
  const serviceAccountPath = process.env.SERVICE_ACCOUNT_PATH!; // We know this is set if we are in this function
  const impersonateUser = process.env.GOOGLE_IMPERSONATE_USER; // Optional: email of user to impersonate
  try {
    const keyFileContent = await fs.readFile(serviceAccountPath, 'utf8');
    const serviceAccountKey = JSON.parse(keyFileContent);

    const auth = new JWT({
      email: serviceAccountKey.client_email,
      key: serviceAccountKey.private_key,
      scopes: SCOPES,
      subject: impersonateUser, // Enables domain-wide delegation when set
    });
    await auth.authorize();
    if (impersonateUser) {
      console.error(`Service Account authentication successful, impersonating: ${impersonateUser}`);
    } else {
      console.error('Service Account authentication successful!');
    }
    return auth;
  } catch (error: any) {
    if (error.code === 'ENOENT') {
      console.error(`FATAL: Service account key file not found at path: ${serviceAccountPath}`);
      throw new Error(`Service account key file not found. Please check the path in SERVICE_ACCOUNT_PATH.`);
    }
    console.error('FATAL: Error loading or authorizing the service account key:', error.message);
    throw new Error('Failed to authorize using the service account. Ensure the key file is valid and the path is correct.');
  }
}
// --- END OF NEW FUNCTION---

async function loadSavedCredentialsIfExist(): Promise<OAuth2Client | null> {
  try {
    const content = await fs.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content.toString());
    const { client_secret, client_id, redirect_uris } = await loadClientSecrets();
    const client = new google.auth.OAuth2(client_id, client_secret, redirect_uris?.[0]);
    client.setCredentials(credentials);
    return client;
  } catch (err) {
    return null;
  }
}

async function loadClientSecrets() {
  const content = await fs.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content.toString());
  const key = keys.installed || keys.web;
   if (!key) throw new Error("Could not find client secrets in credentials.json.");
  return {
      client_id: key.client_id,
      client_secret: key.client_secret,
      redirect_uris: key.redirect_uris || ['http://localhost:3000/'], // Default for web clients
      client_type: keys.web ? 'web' : 'installed'
  };
}

async function saveCredentials(client: OAuth2Client): Promise<void> {
  const { client_secret, client_id } = await loadClientSecrets();
  const payload = JSON.stringify({
    type: 'authorized_user',
    client_id: client_id,
    client_secret: client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFile(TOKEN_PATH, payload);
  console.error('Token stored to', TOKEN_PATH);
}

async function openBrowser(url: string): Promise<void> {
  const platform = process.platform;
  try {
    if (platform === 'darwin') {
      await execAsync(`open "${url}"`);
    } else if (platform === 'win32') {
      await execAsync(`start "" "${url}"`);
    } else {
      await execAsync(`xdg-open "${url}"`);
    }
  } catch (error) {
    console.error('Could not open browser automatically. Please open this URL manually:', url);
  }
}

async function authenticate(): Promise<OAuth2Client> {
  const { client_secret, client_id } = await loadClientSecrets();

  // Use localhost redirect for OAuth (OOB flow is deprecated)
  const PORT = 3000;
  const redirectUri = `http://localhost:${PORT}`;

  console.error(`DEBUG: Using redirect URI: ${redirectUri}`);
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirectUri);

  const authorizeUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES.join(' '),
    prompt: 'consent', // Force consent to ensure we get a refresh token
  });

  return new Promise((resolve, reject) => {
    // Create a temporary HTTP server to receive the OAuth callback
    const server = http.createServer(async (req, res) => {
      try {
        const url = new URL(req.url || '', `http://localhost:${PORT}`);
        const code = url.searchParams.get('code');
        const error = url.searchParams.get('error');

        if (error) {
          res.writeHead(400, { 'Content-Type': 'text/html' });
          res.end(`<html><body><h1>Authorization Failed</h1><p>Error: ${error}</p><p>You can close this window.</p></body></html>`);
          server.close();
          reject(new Error(`Authorization failed: ${error}`));
          return;
        }

        if (code) {
          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end('<html><body><h1>Authorization Successful!</h1><p>You can close this window and return to the terminal.</p></body></html>');

          try {
            const { tokens } = await oAuth2Client.getToken(code);
            oAuth2Client.setCredentials(tokens);
            if (tokens.refresh_token) {
              await saveCredentials(oAuth2Client);
            } else {
              console.error("Did not receive refresh token. Token might expire.");
            }
            console.error('Authentication successful!');
            server.close();
            resolve(oAuth2Client);
          } catch (tokenErr) {
            console.error('Error retrieving access token', tokenErr);
            server.close();
            reject(new Error('Authentication failed'));
          }
        } else {
          res.writeHead(400, { 'Content-Type': 'text/html' });
          res.end('<html><body><h1>No authorization code received</h1><p>Please try again.</p></body></html>');
        }
      } catch (err) {
        console.error('Error handling OAuth callback:', err);
        res.writeHead(500, { 'Content-Type': 'text/html' });
        res.end('<html><body><h1>Server Error</h1></body></html>');
        server.close();
        reject(err);
      }
    });

    server.listen(PORT, async () => {
      console.error(`\nAuthorization server listening on http://localhost:${PORT}`);
      console.error('\nOpening browser for authorization...');
      console.error('If the browser does not open, visit this URL manually:');
      console.error(authorizeUrl);
      console.error('\n');

      await openBrowser(authorizeUrl);
    });

    server.on('error', (err: NodeJS.ErrnoException) => {
      if (err.code === 'EADDRINUSE') {
        console.error(`Port ${PORT} is already in use. Please close other applications using this port.`);
      }
      reject(err);
    });

    // Timeout after 5 minutes
    setTimeout(() => {
      server.close();
      reject(new Error('Authorization timed out after 5 minutes'));
    }, 5 * 60 * 1000);
  });
}

// --- MODIFIED: The Main Exported Function ---
// This function now acts as a router. It checks for the environment
// variable and decides which authentication method to use.
export async function authorize(): Promise<OAuth2Client | JWT> {
  // Check if the Service Account environment variable is set.
  if (process.env.SERVICE_ACCOUNT_PATH) {
    console.error('Service account path detected. Attempting service account authentication...');
    return authorizeWithServiceAccount();
  } else {
    // If not, execute the original OAuth 2.0 flow exactly as it was.
    console.error('No service account path detected. Falling back to standard OAuth 2.0 flow...');
    let client = await loadSavedCredentialsIfExist();
    if (client) {
      // Optional: Add token refresh logic here if needed, though library often handles it.
      console.error('Using saved credentials.');
      return client;
    }
    console.error('Starting authentication flow...');
    client = await authenticate();
    return client;
  }
}
// --- END OF MODIFIED: The Main Exported Function ---
