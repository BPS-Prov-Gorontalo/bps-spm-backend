const jwt = require('jsonwebtoken');
const { PrismaClient } = require('../generated/prisma');
const prisma = new PrismaClient();
const { verifyCentralAuthToken } = require('../services/CentralAuthService');

exports.protect = async (req, res, next) => {
  let token;

  // Cek apakah header Authorization ada dan formatnya benar ('Bearer token')
  if (
    req.headers.authorization &&
    req.headers.authorization.startsWith('Bearer')
  ) {
    try {
      // Ambil token dari header (Bearer[0] token[1])
      token = req.headers.authorization.split(' ')[1];

      // Decode token header untuk mendeteksi algoritma (RS256 vs HS256)
      const decodedUnverified = jwt.decode(token, { complete: true });
      let decoded;

      if (decodedUnverified?.header?.alg === 'RS256') {
        // Verifikasi token asimetris RS256 dari Central Auth Gateway via JWKS
        decoded = await verifyCentralAuthToken(token);

        // Cari user di database berdasarkan identitas klaim token
        req.user = await prisma.user.findFirst({
          where: {
            OR: [
              { email: decoded.email },
              { email: decoded.username },
              { email: `${decoded.username}@bps.go.id` },
            ],
          },
          select: {
            id: true,
            email: true,
            name: true,
            role: true,
            satkerId: true,
          },
        });
      } else {
        // Verifikasi token simetris HS256 internal aplikasi Sipanda
        decoded = jwt.verify(token, process.env.ACCESS_TOKEN_SECRET);

        req.user = await prisma.user.findUnique({
          where: { id: decoded.id },
          select: {
            id: true,
            email: true,
            name: true,
            role: true,
            satkerId: true,
          },
        });
      }

      if (!req.user) {
        return res
          .status(401)
          .json({ error: 'Pengguna tidak ditemukan atau tidak memiliki akses.' });
      }

      return next(); // Lanjutkan ke controller selanjutnya
    } catch (error) {
      if (error.name === 'TokenExpiredError') {
        return res
          .status(401)
          .json({ error: 'Token autentikasi sudah kadaluarsa.' });
      }
      return res
        .status(401)
        .json({ error: 'Tidak terautentikasi, token tidak valid.' });
    }
  }

  if (!token) {
    return res
      .status(401)
      .json({ error: 'Tidak terautentikasi, tidak ada token.' });
  }
};
