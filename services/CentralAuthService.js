const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');

const getCentralAuthUrl = () =>
  (process.env.CENTRAL_AUTH_URL || 'https://auth.bps.go.id').replace(/\/+$/, '');

const getClientId = () =>
  process.env.CENTRAL_AUTH_CLIENT_ID || 'sipanda';

// Inisialisasi JWKS client untuk verifikasi token RS256 dari Central Auth
let jwksInstance = null;
const getJwksClient = () => {
  if (!jwksInstance) {
    jwksInstance = jwksClient({
      jwksUri: `${getCentralAuthUrl()}/.well-known/jwks.json`,
      cache: true,
      rateLimit: true,
    });
  }
  return jwksInstance;
};

class CentralAuthError extends Error {
  constructor(message, statusCode = 400, errorCode = 'AUTH_ERROR') {
    super(message);
    this.name = 'CentralAuthError';
    this.statusCode = statusCode;
    this.errorCode = errorCode;
  }
}

/**
 * Memvalidasi kredensial email/SSO BPS ke Central Auth Gateway.
 * @param {string} username - Email atau username BPS
 * @param {string} password - Password email BPS
 * @param {string} [clientIp] - IP client untuk rate limiter di gateway
 * @returns {Promise<object>} Data respons sukses dari gateway
 */
async function loginWithBpsImap(username, password, clientIp) {
  const centralAuthUrl = getCentralAuthUrl();
  const clientId = getClientId();

  const headers = {
    'Content-Type': 'application/json',
  };
  if (clientIp) {
    headers['X-Forwarded-For'] = clientIp;
  }

  const payload = {
    client_id: clientId,
    username: username.trim(),
    password: password,
  };

  let response;
  try {
    response = await fetch(`${centralAuthUrl}/v1/authenticate`, {
      method: 'POST',
      headers,
      body: JSON.stringify(payload),
      signal: AbortSignal.timeout(12000), // 12 detik timeout sesuai panduan
    });
  } catch (err) {
    if (err.name === 'TimeoutError') {
      throw new CentralAuthError(
        'Koneksi ke Central Auth Gateway timeout.',
        504,
        'GATEWAY_TIMEOUT'
      );
    }
    throw new CentralAuthError(
      'Gagal menghubungi Central Authentication Gateway. Pastikan service auth menyala.',
      503,
      'GATEWAY_UNREACHABLE'
    );
  }

  let data;
  try {
    data = await response.json();
  } catch {
    data = {};
  }

  if (response.status === 200 && data.success) {
    return data;
  }

  if (response.status === 401) {
    throw new CentralAuthError(
      'Username atau password email BPS salah.',
      401,
      'INVALID_CREDENTIALS'
    );
  }

  if (response.status === 429) {
    throw new CentralAuthError(
      'Terlalu banyak percobaan login. Harap tunggu beberapa menit.',
      429,
      'RATE_LIMITED'
    );
  }

  if (response.status === 503) {
    throw new CentralAuthError(
      data.message || 'Authentication provider tidak tersedia.',
      503,
      'PROVIDER_UNAVAILABLE'
    );
  }

  const errorMsg =
    data.message || `Autentikasi gagal (Status ${response.status})`;
  throw new CentralAuthError(
    errorMsg,
    response.status,
    data.code || 'AUTH_FAILED'
  );
}

/**
 * Helper untuk verifikasi token RS256 via JWKS di Node.js
 * @param {string} token 
 * @returns {Promise<object>} Decoded token payload
 */
function verifyCentralAuthToken(token) {
  const client = getJwksClient();
  return new Promise((resolve, reject) => {
    jwt.verify(
      token,
      (header, callback) => {
        client.getSigningKey(header.kid, (err, key) => {
          if (err) return callback(err);
          const signingKey = key.getPublicKey();
          callback(null, signingKey);
        });
      },
      { algorithms: ['RS256'] },
      (err, decoded) => {
        if (err) return reject(err);
        resolve(decoded);
      }
    );
  });
}

module.exports = {
  loginWithBpsImap,
  verifyCentralAuthToken,
  CentralAuthError,
  getCentralAuthUrl,
  getClientId,
};

