const { PrismaClient } = require('../generated/prisma');

const prisma = new PrismaClient();
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
const crypto = require('crypto');
const {
  loginWithBpsImap,
  CentralAuthError,
} = require('../services/CentralAuthService');

// --- Helper: Generate Name from Email ---
const createNameFromEmail = (email) => {
  return email
    .split('@')[0]
    .replace('-', ' ')
    .split('.')
    .map((namePart) => namePart.charAt(0).toUpperCase() + namePart.slice(1))
    .join(' ');
};

exports.login = async (req, res) => {
  // Mendukung login dengan field 'email' atau 'username'
  const identifier = (req.body.email || req.body.username || '').trim();
  const password = req.body.password;

  if (!identifier || !password) {
    return res
      .status(400)
      .json({ error: 'Email/Username dan password harus diisi.' });
  }

  // Dapatkan IP klien sebenarnya untuk diteruskan ke gateway (X-Forwarded-For)
  const clientIp =
    req.headers['x-forwarded-for']?.split(',')[0].trim() ||
    req.socket?.remoteAddress ||
    req.ip;

  // 1. Akun lokal khusus (misalnya akun pengujian yang berakhiran .local)
  if (identifier.endsWith('.local')) {
    try {
      const user = await prisma.user.findFirst({
        where: { email: identifier },
      });
      if (!user) {
        return res.status(401).json({ error: 'Email atau password salah.' });
      }
      const isPasswordMatch = await bcrypt.compare(password, user.password);
      if (!isPasswordMatch) {
        return res.status(401).json({ error: 'Email atau password salah.' });
      }

      const payload = {
        id: user.id,
        name: user.name,
        role: user.role,
        satkerId: user.satkerId,
        email: user.email,
      };

      const token = jwt.sign(payload, process.env.ACCESS_TOKEN_SECRET, {
        expiresIn: '1d',
      });

      return res.status(200).json({
        success: true,
        token,
        access_token: token,
        user: {
          id: user.id,
          name: user.name,
          role: user.role,
          satkerId: user.satkerId,
          email: user.email,
        },
      });
    } catch (error) {
      console.error('Server error during local login:', error);
      return res.status(500).json({ error: 'Terjadi kesalahan pada server.' });
    }
  }

  // 2. Autentikasi melalui Central Authentication Gateway (BPS SSO/IMAP)
  try {
    const authData = await loginWithBpsImap(identifier, password, clientIp);
    const { identity, token: gatewayToken } = authData;

    const bpsUsername = identity?.username || identifier.split('@')[0];
    const bpsEmail = identity?.email || identifier;

    // Cari akun di database Sipanda
    const user = await prisma.user.findFirst({
      where: {
        OR: [
          { email: bpsEmail },
          { email: bpsUsername },
          { email: identifier },
          { email: `${bpsUsername}@bps.go.id` },
        ],
      },
    });

    if (!user) {
      return res.status(403).json({
        error:
          'Autentikasi berhasil, namun akun Anda belum terdaftar di aplikasi ini. Hubungi admin.',
        identity,
      });
    }

    // Buat token JWT sesi untuk aplikasi Sipanda
    const payload = {
      id: user.id,
      name: user.name,
      role: user.role,
      satkerId: user.satkerId,
      email: user.email,
      gatewayUserId: identity?.user_id,
    };

    const appToken = jwt.sign(payload, process.env.ACCESS_TOKEN_SECRET, {
      expiresIn: '1d',
    });

    return res.status(200).json({
      success: true,
      token: appToken,
      access_token: appToken,
      gateway_token: gatewayToken?.access_token,
      user: {
        id: user.id,
        name: user.name,
        role: user.role,
        satkerId: user.satkerId,
        email: user.email,
      },
    });
  } catch (error) {
    if (error instanceof CentralAuthError) {
      return res.status(error.statusCode).json({
        error: error.message,
        code: error.errorCode,
      });
    }

    console.error('Error saat login via Central Auth:', error.message);
    return res
      .status(500)
      .json({ error: 'Terjadi kesalahan pada server autentikasi.' });
  }
};

exports.getAllUsers = async (req, res) => {
  try {
    const users = await prisma.user.findMany({
      orderBy: [{ satker: { nama: 'asc' } }, { name: 'asc' }],
      include: {
        satker: {
          select: {
            nama: true,
          },
        },
      },
    });
    users.forEach((user) => delete user.password);
    res.status(200).json(users);
  } catch (error) {
    res.status(500).json({ error: 'Gagal mengambil daftar pengguna.' });
  }
};

// @desc    Create a new user (With Manual Name Support)
// @route   POST /api/users
exports.createUser = async (req, res) => {
  const { email, role, satkerId, name } = req.body;

  if (!email || !role) {
    return res
      .status(400)
      .json({ error: 'Email dan peran (role) harus diisi.' });
  }

  // --- REMOVED RESTRICTION FOR SUPERVISOR ---

  try {
    const finalName =
      name && name.trim() !== '' ? name : createNameFromEmail(email);

    const randomPassword = crypto.randomBytes(32).toString('hex');
    const passwordHash = bcrypt.hashSync(randomPassword, 10);

    const newUser = await prisma.user.create({
      data: {
        email,
        name: finalName,
        role,
        password: passwordHash,
        satkerId: satkerId ? parseInt(satkerId) : null,
      },
    });
    delete newUser.password;
    res.status(201).json(newUser);
  } catch (error) {
    console.error(error);
    if (error.code === 'P2002') {
      return res
        .status(409)
        .json({ error: `Email '${email}' sudah terdaftar.` });
    }
    res.status(500).json({ error: 'Gagal membuat pengguna baru.' });
  }
};

// @desc    Update a user
// @route   PUT /api/users/:id
exports.updateUser = async (req, res) => {
  const { id } = req.params;
  const { name, role, satkerId, password } = req.body;

  try {
    const dataToUpdate = {
      name,
      role,
      satkerId: satkerId ? parseInt(satkerId) : null,
    };

    if (password) {
      dataToUpdate.password = bcrypt.hashSync(password, 10);
    }

    const updatedUser = await prisma.user.update({
      where: { id: parseInt(id) },
      data: dataToUpdate,
    });
    delete updatedUser.password;
    res.status(200).json(updatedUser);
  } catch (error) {
    res.status(500).json({ error: 'Gagal memperbarui pengguna.' });
  }
};

// @desc    Delete a user
// @route   DELETE /api/users/:id
exports.deleteUser = async (req, res) => {
  const { id } = req.params;

  try {
    await prisma.user.delete({
      where: { id: parseInt(id) },
    });
    res.status(200).json({ message: 'Pengguna berhasil dihapus.' });
  } catch (error) {
    if (error.code === 'P2025') {
      return res.status(404).json({ error: 'Pengguna tidak ditemukan.' });
    }
    res.status(500).json({ error: 'Gagal menghapus pengguna.' });
  }
};
