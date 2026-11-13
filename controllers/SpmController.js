// controllers/SpmController.js

const { PrismaClient } = require('../generated/prisma');
const prisma = new PrismaClient();
const ExcelJS = require('exceljs');

// --- FUNGSI HELPER UNTUK KALKULASI PERSENTASE RINCIAN ---
async function calculateRincianPercentage(rincian) {
  if (!rincian.jawabanFlags) {
    rincian.jawabanFlags = await prisma.jawabanFlag.findMany({
      where: { rincianSpmId: rincian.id },
    });
  }
  const totalRequiredFlags = await prisma.flag.count({
    where: { kodeAkunId: rincian.kodeAkunId },
  });
  if (totalRequiredFlags === 0) return 100;

  // REVISI LOGIC:
  // "IYA" = Lengkap (1 poin)
  // "IYA_TIDAK" (Strip/-) = Lengkap/Valid (1 poin)
  // "TIDAK" = Tidak Lengkap (0 poin)
  // "BELUM_SELESAI" = Tidak Lengkap (0 poin)
  const totalValid = rincian.jawabanFlags.filter(
    (flag) => flag.tipe === 'IYA' || flag.tipe === 'IYA_TIDAK'
  ).length;

  const percentage = Math.round((totalValid / totalRequiredFlags) * 100);
  return Math.min(100, percentage);
}

// @desc    Download Template Excel (Format: "KODE - NAMA")
// @route   GET /api/spm/template
exports.downloadImportTemplate = async (req, res) => {
  try {
    // 1. Fetch Data
    const allFlags = await prisma.flag.findMany({
      distinct: ['nama'],
      orderBy: { nama: 'asc' },
      select: { nama: true },
    });

    const allKodeAkun = await prisma.kodeAkun.findMany({
      include: { templateFlags: true },
      orderBy: { kode: 'asc' },
    });

    const flagNames = allFlags.map((f) => f.nama);
    const workbook = new ExcelJS.Workbook();

    // --- SHEET 0: PETUNJUK (INSTRUCTIONS) ---
    const guideSheet = workbook.addWorksheet('Petunjuk');
    guideSheet.addRows([
      ['PANDUAN PENGISIAN'],
      [''],
      ['1. KODE AKUN (PENTING)'],
      ['   - Kolom Kode Akun sekarang memuat "NOMOR - JUDUL".'],
      ['   - Pastikan memilih dari Dropdown agar formatnya sesuai.'],
      [''],
      ['2. PANDUAN WARNA'],
      ['   - KOTAK PUTIH: Boleh diisi (Dokumen relevan).'],
      ['   - KOTAK GELAP: Tidak perlu diisi.'],
      [''],
      ['3. ISI DOKUMEN'],
      ['   - Ketik "IYA" atau "1" jika ada.'],
      ['   - Ketik "-" jika tidak ada tapi valid (Strip).'],
      ['   - Ketik "?" atau "B" jika belum selesai.'],
      ['   - KOSONG artinya "TIDAK" (Dokumen tidak ada).'],
    ]);
    guideSheet.getColumn(1).width = 60;

    // --- SHEET 1: INPUT DATA ---
    const worksheet = workbook.addWorksheet('Input Data');

    const staticHeaders = [
      'Nomor SPM',
      'Tanggal (YYYY-MM-DD)',
      'Link Google Drive',
      'Kode Program',
      'Kode Kegiatan',
      'Kode Akun',
      'Kode KRO',
      'Kode RO',
      'Kode Komponen',
      'Kode Subkomponen',
      'Uraian',
      'Jumlah',
    ];

    worksheet.columns = [
      ...staticHeaders.map((h) => ({
        header: h,
        key: h.replace(/\s/g, ''),
        width: 18,
      })),
      ...flagNames.map((f) => ({ header: f, key: f, width: 18 })),
    ];

    // Header Style
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF002B6A' },
    };
    headerRow.alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true,
    };
    headerRow.height = 45;

    // --- SHEET 2: REFERENCE (For Dropdown & Logic) ---
    const refSheet = workbook.addWorksheet('Reference');
    refSheet.state = 'hidden';
    refSheet.getRow(1).values = ['KodeAkun', ...flagNames];

    allKodeAkun.forEach((akun) => {
      // --- CHANGE: Combine Code + Name for Uniqueness ---
      const label = `${akun.kode} - ${akun.nama}`;
      const rowData = [label];

      flagNames.forEach((flagName) => {
        const isRequired = akun.templateFlags.some(
          (tf) => tf.nama === flagName
        );
        rowData.push(isRequired ? 'REQ' : 'NA');
      });
      refSheet.addRow(rowData);
    });

    const kodeAkunRange = `Reference!$A$2:$A$${allKodeAkun.length + 1}`;
    const totalColumns = staticHeaders.length + flagNames.length;

    // --- ROW STYLES ---
    for (let i = 2; i <= 500; i++) {
      const row = worksheet.getRow(i);

      // Borders
      for (let j = 1; j <= totalColumns; j++) {
        row.getCell(j).border = {
          top: { style: 'thin', color: { argb: 'FF999999' } },
          left: { style: 'thin', color: { argb: 'FF999999' } },
          bottom: { style: 'thin', color: { argb: 'FF999999' } },
          right: { style: 'thin', color: { argb: 'FF999999' } },
        };
      }

      // Validation Dropdown (Now shows "Code - Name")
      row.getCell(6).dataValidation = {
        type: 'list',
        allowBlank: false,
        formulae: [kodeAkunRange],
      };

      // Make Column F (Kode Akun) wider to fit the text
      worksheet.getColumn(6).width = 40;
    }

    // --- CONDITIONAL FORMATTING ---
    const startFlagColIndex = staticHeaders.length + 1;

    flagNames.forEach((flag, index) => {
      const colLetter = worksheet.getColumn(startFlagColIndex + index).letter;
      const refColLetter = refSheet.getColumn(index + 2).letter;

      // Validation for Flag Cells
      for (let r = 2; r <= 500; r++) {
        worksheet.getCell(`${colLetter}${r}`).dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: ['"IYA,TIDAK,-,?"'],
        };
        worksheet.getCell(`${colLetter}${r}`).alignment = {
          horizontal: 'center',
        };
      }

      // Blackout Logic
      worksheet.addConditionalFormatting({
        ref: `${colLetter}2:${colLetter}500`,
        rules: [
          {
            type: 'expression',
            // Matches the full "Code - Name" string against Reference sheet
            formulae: [
              `INDEX(Reference!$${refColLetter}$2:$${refColLetter}$999, MATCH($F2, Reference!$A$2:$A$999, 0))="NA"`,
            ],
            style: {
              fill: {
                type: 'pattern',
                pattern: 'lightDown',
                fgColor: { argb: 'FFCCCCCC' },
                bgColor: { argb: 'FF333333' },
              },
              font: { color: { argb: 'FF555555' } },
            },
          },
        ],
      });
    });

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename=Template_Import_SPM_v7.xlsx'
    );
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    res.status(500).json({ error: 'Gagal membuat template Excel.' });
  }
};

// @desc    Import Bulk SPM (Parses "KODE - NAMA" & New Flags)
// @route   POST /api/spm/import
exports.importSpms = async (req, res) => {
  if (!req.file)
    return res.status(400).json({ error: 'File Excel wajib diunggah.' });

  let satkerId = req.body.satkerId;
  if (req.user.role === 'op_satker') satkerId = req.user.satkerId;
  if (!satkerId)
    return res.status(400).json({ error: 'Satker ID tidak valid.' });

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);
    const worksheet =
      workbook.getWorksheet('Input Data') ||
      workbook.worksheets[1] ||
      workbook.worksheets[0];

    const spmGroups = {};
    const allKodeAkun = await prisma.kodeAkun.findMany();

    // 1. IDENTIFIKASI HEADER FLAGS (Kolom 13 ke atas)
    const flagHeaders = [];
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell, colNumber) => {
      if (colNumber >= 13) {
        flagHeaders.push({ col: colNumber, name: cell.text });
      }
    });

    // 2. PARSING ROWS
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      const nomorSpm = row.getCell(1).text?.trim();
      if (!nomorSpm) return;

      const tanggalVal = row.getCell(2).value;
      const tanggal = new Date(tanggalVal);
      const driveLink = row.getCell(3).text?.trim();

      const kodeProgram = row.getCell(4).text;
      const kodeKegiatan = row.getCell(5).text;
      const kodeAkunRaw = row.getCell(6).text;
      const kodeKRO = row.getCell(7).text;
      const kodeRO = row.getCell(8).text;
      const kodeKomponen = row.getCell(9).text;
      const kodeSubkomponen = row.getCell(10).text;
      const uraian = row.getCell(11).text;
      const jumlah = parseFloat(row.getCell(12).value) || 0;

      if (!spmGroups[nomorSpm]) {
        spmGroups[nomorSpm] = {
          nomorSpm,
          driveLink,
          tanggal: isNaN(tanggal) ? new Date() : tanggal,
          tahunAnggaran: isNaN(tanggal)
            ? new Date().getFullYear()
            : tanggal.getFullYear(),
          satkerId: parseInt(satkerId),
          rincian: [],
        };
      }

      // Logic pengambilan flag yang direvisi:
      // Loop berdasarkan header yang ditemukan, BUKAN berdasarkan cell yang ada isinya.
      // Ini memastikan cell kosong tetap terproses sebagai "TIDAK".
      const flagCells = {};
      flagHeaders.forEach(({ col, name }) => {
        const cell = row.getCell(col);
        let valStr = cell.text
          ? cell.text.toString().toUpperCase().trim()
          : 'TIDAK';
        let cleanValue = 'TIDAK'; // Default jika kosong/tidak dikenali

        // LOGIC MAPPING
        if (['YA', 'IYA', 'ADA', '1', 'TRUE', 'V'].includes(valStr)) {
          cleanValue = 'IYA';
        } else if (['-', 'STRIP', 'NA', 'N/A'].includes(valStr)) {
          cleanValue = 'IYA_TIDAK'; // Strip
        } else if (
          ['B', 'BELUM', 'BELUM SELESAI', '?', 'TUNDA'].includes(valStr)
        ) {
          cleanValue = 'BELUM_SELESAI';
        } else {
          cleanValue = 'TIDAK'; // Kosong atau nilai lain dianggap TIDAK
        }

        flagCells[name] = cleanValue;
      });

      spmGroups[nomorSpm].rincian.push({
        kodeProgram,
        kodeKegiatan,
        kodeAkunRaw,
        kodeKRO,
        kodeRO,
        kodeKomponen,
        kodeSubkomponen,
        uraian,
        jumlah,
        flags: flagCells,
      });
    });

    const spmKeys = Object.keys(spmGroups);
    if (spmKeys.length === 0) {
      return res
        .status(400)
        .json({ error: 'File Excel kosong atau tidak valid.' });
    }

    // 3. PRE-CHECK DUPLICATE
    // Cek apakah nomor SPM ini sudah ada untuk Satker & Tahun yang bersangkutan
    const distinctTahun = [
      ...new Set(Object.values(spmGroups).map((s) => s.tahunAnggaran)),
    ];

    const existingSpms = await prisma.spm.findMany({
      where: {
        satkerId: parseInt(satkerId),
        tahunAnggaran: { in: distinctTahun },
        nomorSpm: { in: spmKeys },
      },
      select: { nomorSpm: true, tahunAnggaran: true },
    });

    // Set komposit untuk cek duplikat cepat
    const existingSet = new Set(
      existingSpms.map((s) => `${s.nomorSpm}-${s.tahunAnggaran}`)
    );

    const toInsert = [];
    const skipped = [];

    for (const key of spmKeys) {
      const spm = spmGroups[key];
      const compositeKey = `${spm.nomorSpm}-${spm.tahunAnggaran}`;
      if (existingSet.has(compositeKey)) {
        skipped.push(spm.nomorSpm);
      } else {
        toInsert.push(spm);
      }
    }

    if (toInsert.length === 0) {
      return res.status(200).json({
        message: `Semua SPM (${skipped.length} data) sudah ada di sistem. Tidak ada data baru yang diimpor.`,
        skipped,
      });
    }

    // 4. TRANSACTION (Hanya untuk yang belum ada)
    await prisma.$transaction(async (tx) => {
      for (const spmData of toInsert) {
        const totalAnggaran = spmData.rincian.reduce(
          (sum, item) => sum + item.jumlah,
          0
        );

        const newSpm = await tx.spm.create({
          data: {
            nomorSpm: spmData.nomorSpm,
            tahunAnggaran: spmData.tahunAnggaran,
            tanggal: spmData.tanggal,
            satkerId: spmData.satkerId,
            totalAnggaran: totalAnggaran,
            driveLink: spmData.driveLink,
            status: 'MENUNGGU',
          },
        });

        for (const item of spmData.rincian) {
          let akunDb = null;

          if (item.kodeAkunRaw) {
            const parts = item.kodeAkunRaw.split(' - ');
            if (parts.length >= 2) {
              const codePart = parts[0].trim();
              const namePart = parts.slice(1).join(' - ').trim();
              akunDb = allKodeAkun.find(
                (k) => k.kode === codePart && k.nama === namePart
              );
            }
            if (!akunDb) {
              akunDb = allKodeAkun.find(
                (k) => k.kode.toString() === item.kodeAkunRaw.toString()
              );
            }
          }

          if (!akunDb) {
            throw new Error(
              `Kode Akun '${item.kodeAkunRaw}' pada SPM '${spmData.nomorSpm}' tidak ditemukan.`
            );
          }

          const newRincian = await tx.spmRincian.create({
            data: {
              spmId: newSpm.id,
              kodeAkunId: akunDb.id,
              kodeProgram: item.kodeProgram || '-',
              kodeKegiatan: item.kodeKegiatan || '-',
              kodeKRO: item.kodeKRO || '-',
              kodeRO: item.kodeRO || '-',
              kodeKomponen: item.kodeKomponen || '-',
              kodeSubkomponen: item.kodeSubkomponen || '-',
              uraian: item.uraian || '-',
              jumlah: item.jumlah,
            },
          });

          const flagData = Object.entries(item.flags).map(([nama, val]) => ({
            rincianSpmId: newRincian.id,
            nama: nama,
            tipe: val,
          }));

          if (flagData.length > 0) {
            await tx.jawabanFlag.createMany({ data: flagData });
          }
        }
      }
    });

    // 5. CONSTRUCT MESSAGE
    let msg = `Berhasil mengimpor ${toInsert.length} SPM.`;
    if (skipped.length > 0) {
      msg += `\n\n⚠️ ${
        skipped.length
      } SPM dilewati karena sudah ada di sistem:\n${skipped.join(', ')}`;
    }

    res.status(200).json({
      message: msg,
      importedCount: toInsert.length,
      skippedCount: skipped.length,
      skippedSpms: skipped,
    });
  } catch (error) {
    console.error('Import Error:', error);
    if (error.code === 'P2002')
      return res.status(409).json({
        error: 'Terdapat konflik data (Unique Constraint) saat menyimpan.',
      });
    res
      .status(500)
      .json({ error: error.message || 'Gagal memproses file import.' });
  }
};

// @desc    Membuat SPM baru dengan Rincian
// @route   POST /api/spm
exports.createSpmWithRincian = async (req, res) => {
  if (req.user.role === 'viewer') {
    return res
      .status(403)
      .json({ error: 'Akses ditolak. Viewer tidak dapat membuat data.' });
  }

  try {
    const { nomorSpm, tahunAnggaran, tanggal, satkerId, rincian, driveLink } =
      req.body;

    if (!rincian || rincian.length === 0) {
      return res
        .status(400)
        .json({ error: 'SPM harus memiliki setidaknya satu rincian.' });
    }

    const calculatedTotal = rincian.reduce(
      (total, item) => total + (Number(item.jumlah) || 0),
      0
    );

    const newSpm = await prisma.spm.create({
      data: {
        nomorSpm,
        tahunAnggaran: parseInt(tahunAnggaran),
        tanggal: new Date(tanggal),
        totalAnggaran: calculatedTotal,
        driveLink: driveLink,
        satker: { connect: { id: parseInt(satkerId) } },
        rincian: {
          create: rincian.map((r) => ({
            kodeProgram: r.kodeProgram,
            kodeKegiatan: r.kodeKegiatan,
            jumlah: parseInt(r.jumlah),
            catatan: r.catatan || null, // Capture catatan on create
            kodeAkun: { connect: { id: parseInt(r.kodeAkunId) } },
            jawabanFlags: {
              create: r.jawabanFlags.map(({ nama, tipe }) => ({ nama, tipe })),
            },
            kodeKRO: r.kodeKRO,
            kodeRO: r.kodeRO,
            kodeKomponen: r.kodeKomponen,
            kodeSubkomponen: r.kodeSubkomponen,
            uraian: r.uraian,
          })),
        },
      },
      include: { rincian: { include: { jawabanFlags: true } } },
    });

    res.status(201).json(newSpm);
  } catch (error) {
    console.error('--- DETAIL ERROR PEMBUATAN SPM ---', error);
    if (error.code === 'P2002') {
      return res
        .status(409)
        .json({ error: `Nomor SPM '${req.body.nomorSpm}' sudah terdaftar.` });
    }
    res.status(500).json({ error: 'Gagal membuat SPM beserta rinciannya.' });
  }
};

// @desc    Mendapatkan semua SPM (dengan filter & pagination)
// @route   GET /api/spm
exports.getAllSpms = async (req, res) => {
  try {
    const { satkerId, tahun, page, limit, search, status, sortBy, order } =
      req.query;

    const applyPagination = page && limit;
    const pageNum = applyPagination ? parseInt(page) : 1;
    const limitNum = applyPagination ? parseInt(limit) : undefined;
    const skip = applyPagination ? (pageNum - 1) * limitNum : 0;

    // 1. Build Where Clause
    const whereClause = {};

    // Filter Tahun
    if (tahun) {
      whereClause.tahunAnggaran = parseInt(tahun, 10);
    }

    // Filter Satker
    if (req.user.role === 'op_satker') {
      whereClause.satkerId = req.user.satkerId;
    } else if (satkerId) {
      whereClause.satkerId = parseInt(satkerId, 10);
    }

    // Filter Status (New)
    if (status && status !== 'ALL') {
      whereClause.status = status;
    }

    // Search Logic (New) - Case Insensitive (contains)
    if (search) {
      whereClause.OR = [
        { nomorSpm: { contains: search, mode: 'insensitive' } },
        // Uncomment below if you want to search by Uraian inside Rincian too (Heavy operation)
        // { rincian: { some: { uraian: { contains: search, mode: 'insensitive' } } } }
      ];
    }

    // 2. Build Sort Order & Query Execution
    // Default: Tanggal Descending (Newest first)
    if (sortBy === 'completenessPercentage') {
      const allSpms = await prisma.spm.findMany({
        where: whereClause,
        include: {
          satker: { select: { nama: true } },
          rincian: {
            include: {
              jawabanFlags: true,
            },
          },
        },
      });

      await Promise.all(
        allSpms.map(async (spm) => {
          spm._count = { rincian: spm.rincian.length };
          if (spm.rincian.length === 0) {
            spm.completenessPercentage = 100;
          } else {
            const percentages = await Promise.all(
              spm.rincian.map((rincian) => calculateRincianPercentage(rincian))
            );
            const totalPercentage = percentages.reduce((sum, p) => sum + p, 0);
            spm.completenessPercentage = Math.round(
              totalPercentage / spm.rincian.length
            );
          }
          delete spm.rincian;
        })
      );

      allSpms.sort((a, b) => {
        if (order === 'asc') {
          return a.completenessPercentage - b.completenessPercentage;
        } else {
          return b.completenessPercentage - a.completenessPercentage;
        }
      });

      const totalSpms = allSpms.length;
      const paginatedSpms = applyPagination
        ? allSpms.slice(skip, skip + limitNum)
        : allSpms;

      return res.status(200).json({
        spms: paginatedSpms,
        totalCount: totalSpms,
        totalPages: applyPagination ? Math.ceil(totalSpms / limitNum) : 1,
        currentPage: pageNum,
      });
    }

    let orderBy = { tanggal: 'desc' };

    if (sortBy === 'satker') {
      orderBy = { satker: { nama: order || 'asc' } };
    } else if (sortBy && order) {
      orderBy = { [sortBy]: order };
    }

    const [spms, totalSpms] = await prisma.$transaction([
      prisma.spm.findMany({
        where: whereClause,
        skip: applyPagination ? skip : undefined,
        take: limitNum,
        orderBy: orderBy,
        include: {
          satker: { select: { nama: true } },
          rincian: {
            include: {
              jawabanFlags: true,
            },
          },
        },
      }),
      prisma.spm.count({ where: whereClause }),
    ]);

    // Calculate percentages (Keep existing logic)
    await Promise.all(
      spms.map(async (spm) => {
        spm._count = { rincian: spm.rincian.length };
        if (spm.rincian.length === 0) {
          spm.completenessPercentage = 100;
        } else {
          const percentages = await Promise.all(
            spm.rincian.map((rincian) => calculateRincianPercentage(rincian))
          );
          const totalPercentage = percentages.reduce((sum, p) => sum + p, 0);
          spm.completenessPercentage = Math.round(
            totalPercentage / spm.rincian.length
          );
        }
        delete spm.rincian;
      })
    );

    res.status(200).json({
      spms,
      totalCount: totalSpms,
      totalPages: applyPagination ? Math.ceil(totalSpms / limitNum) : 1,
      currentPage: pageNum,
    });
  } catch (error) {
    console.error('--- DETAIL ERROR GET ALL SPMS ---', error);
    res.status(500).json({ error: 'Gagal mengambil daftar SPM.' });
  }
};
// @desc    Mendapatkan detail satu SPM
// @route   GET /api/spm/:id
exports.getSpmById = async (req, res) => {
  try {
    const { id } = req.params;
    const spm = await prisma.spm.findUnique({
      where: { id: parseInt(id) },
      include: {
        satker: true,
        rincian: {
          include: {
            kodeAkun: true,
            jawabanFlags: true,
          },
        },
      },
    });

    if (!spm) {
      return res.status(404).json({ error: 'SPM tidak ditemukan.' });
    }

    if (req.user.role === 'op_satker' && spm.satkerId !== req.user.satkerId) {
      return res.status(403).json({ error: 'Akses ditolak.' });
    }

    await Promise.all(
      spm.rincian.map(async (rincian) => {
        rincian.persentaseKelengkapan = await calculateRincianPercentage(
          rincian
        );
      })
    );

    if (spm.rincian.length === 0) {
      spm.completenessPercentage = 100;
    } else {
      const totalPercentage = spm.rincian.reduce(
        (sum, r) => sum + (r.persentaseKelengkapan || 0),
        0
      );
      spm.completenessPercentage = Math.round(
        totalPercentage / spm.rincian.length
      );
    }

    res.status(200).json(spm);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Gagal mengambil data SPM.' });
  }
};

// @desc    Update SPM
// @route   PUT /api/spm/:id
exports.updateSpm = async (req, res) => {
  const { id } = req.params;

  if (req.user.role === 'viewer') {
    return res
      .status(403)
      .json({ error: 'Akses ditolak. Viewer tidak dapat mengubah data.' });
  }

  const { nomorSpm, tahunAnggaran, tanggal, satkerId, rincian, driveLink } =
    req.body;

  try {
    const spmToUpdate = await prisma.spm.findUnique({
      where: { id: parseInt(id) },
    });
    if (!spmToUpdate) {
      return res.status(404).json({ error: 'SPM tidak ditemukan.' });
    }

    if (
      req.user.role === 'op_satker' &&
      spmToUpdate.satkerId !== req.user.satkerId
    ) {
      return res.status(403).json({
        error:
          'Akses ditolak. Anda tidak memiliki izin untuk mengedit SPM ini.',
      });
    }

    // Hanya Supervisor yang bisa edit meskipun DITERIMA, role lain diblokir
    if (spmToUpdate.status === 'DITERIMA' && req.user.role !== 'supervisor') {
      return res.status(403).json({
        error: 'Akses ditolak. SPM yang sudah diterima tidak dapat diubah.',
      });
    }

    if (
      req.user.role === 'op_satker' &&
      !['MENUNGGU', 'DITOLAK'].includes(spmToUpdate.status)
    ) {
      return res
        .status(403)
        .json({ error: 'Akses ditolak. SPM ini tidak dapat diubah lagi.' });
    }

    const calculatedTotal = rincian.reduce(
      (total, item) => total + (Number(item.jumlah) || 0),
      0
    );
    const existingRincian = await prisma.spmRincian.findMany({
      where: { spmId: parseInt(id) },
      select: { id: true },
    });
    const existingRincianIds = existingRincian.map((r) => r.id);
    const incomingRincianIds = rincian.filter((r) => r.id).map((r) => r.id);
    const rincianToDeleteIds = existingRincianIds.filter(
      (existingId) => !incomingRincianIds.includes(existingId)
    );

    await prisma.$transaction(async (tx) => {
      if (rincianToDeleteIds.length > 0) {
        await tx.jawabanFlag.deleteMany({
          where: { rincianSpmId: { in: rincianToDeleteIds } },
        });
        await tx.spmRincian.deleteMany({
          where: { id: { in: rincianToDeleteIds } },
        });
      }

      const dataToUpdate = {
        nomorSpm,
        tahunAnggaran: parseInt(tahunAnggaran),
        tanggal: new Date(tanggal),
        satkerId: parseInt(satkerId),
        totalAnggaran: calculatedTotal,
        driveLink: driveLink,
      };

      // Reset status jika user biasa yang edit dan sebelumnya DITOLAK
      // Supervisor bisa edit tanpa mereset status
      if (spmToUpdate.status === 'DITOLAK' && req.user.role !== 'supervisor') {
        dataToUpdate.status = 'MENUNGGU';
        dataToUpdate.rejectionComment = null;
      }

      const spm = await tx.spm.update({
        where: { id: parseInt(id) },
        data: dataToUpdate,
      });

      for (const rincianData of rincian) {
        const {
          id: rincianId,
          kodeAkunId,
          jawabanFlags,
          catatan, // [NEW] Ambil catatan dari payload
          ...restOfData
        } = rincianData;

        const cleanJawabanFlags = jawabanFlags.map(({ nama, tipe }) => ({
          nama,
          tipe,
        }));

        await tx.spmRincian.upsert({
          where: { id: rincianId || -1 },
          create: {
            ...restOfData,
            jumlah: parseInt(restOfData.jumlah) || 0,
            catatan: catatan || null, // [NEW] Save catatan
            spm: { connect: { id: spm.id } },
            kodeAkun: { connect: { id: parseInt(kodeAkunId) } },
            jawabanFlags: { create: cleanJawabanFlags },
          },
          update: {
            ...restOfData,
            jumlah: parseInt(restOfData.jumlah) || 0,
            catatan: catatan || null, // [NEW] Update catatan
            kodeAkun: { connect: { id: parseInt(kodeAkunId) } },
            jawabanFlags: { deleteMany: {}, create: cleanJawabanFlags },
          },
        });
      }
    });

    res.status(200).json({ message: 'SPM berhasil diupdate.' });
  } catch (error) {
    console.error('--- DETAIL ERROR UPDATE SPM ---', error);
    if (error.code === 'P2002') {
      return res
        .status(409)
        .json({ error: `Nomor SPM '${req.body.nomorSpm}' sudah terdaftar.` });
    }
    res.status(500).json({ error: 'Gagal mengupdate SPM.' });
  }
};

// @desc    Menghapus SPM
// @route   DELETE /api/spm/:id
exports.deleteSpm = async (req, res) => {
  try {
    const { id } = req.params;

    if (req.user.role === 'viewer') {
      return res
        .status(403)
        .json({ error: 'Akses ditolak. Viewer tidak dapat menghapus data.' });
    }

    const spmToDelete = await prisma.spm.findUnique({
      where: { id: parseInt(id) },
    });
    if (!spmToDelete) {
      return res.status(404).json({ error: 'SPM tidak ditemukan.' });
    }
    if (
      req.user.role === 'op_satker' &&
      spmToDelete.satkerId !== req.user.satkerId
    ) {
      return res.status(403).json({ error: 'Akses ditolak.' });
    }
    if (spmToDelete.status === 'DITERIMA' && req.user.role !== 'supervisor') {
      return res
        .status(403)
        .json({ error: 'SPM yang sudah diterima tidak dapat dihapus.' });
    }

    await prisma.spm.delete({
      where: { id: parseInt(id) },
    });

    res.status(200).json({ message: 'SPM berhasil dihapus.' });
  } catch (error) {
    console.error(error);
    if (error.code === 'P2025') {
      return res.status(404).json({ error: 'SPM tidak ditemukan.' });
    }
    res.status(500).json({ error: 'Gagal menghapus SPM.' });
  }
};

// @desc    Update Status SPM
// @route   PATCH /api/spm/:id/status
exports.updateSpmStatus = async (req, res) => {
  const { id } = req.params;
  const { status, comment } = req.body;

  // Hanya Supervisor yang boleh update status (Terima/Tolak)
  if (req.user.role !== 'supervisor') {
    return res.status(403).json({
      error: 'Akses ditolak. Hanya Supervisor yang dapat memvalidasi SPM.',
    });
  }

  if (!status || !['DITERIMA', 'DITOLAK'].includes(status)) {
    return res.status(400).json({ error: 'Status yang dikirim tidak valid.' });
  }

  try {
    const dataToUpdate = { status: status };

    if (status === 'DITOLAK') {
      dataToUpdate.rejectionComment = comment || null;
    } else {
      dataToUpdate.rejectionComment = null;
    }

    const updatedSpm = await prisma.spm.update({
      where: { id: parseInt(id) },
      data: dataToUpdate,
    });
    res.status(200).json(updatedSpm);
  } catch (error) {
    console.error('--- DETAIL ERROR UPDATE STATUS ---', error);
    if (error.code === 'P2025') {
      return res.status(404).json({ error: 'SPM tidak ditemukan.' });
    }
    res.status(500).json({ error: 'Gagal memperbarui status SPM.' });
  }
};

// @desc    Validasi Laporan SAKTI (WITHOUT PAGU)
// @route   POST /api/spm/validate-report
exports.validateSaktiReport = async (req, res) => {
  const { data: reportRows } = req.body;
  const { tahun, satkerId } = req.query;

  if (!reportRows || !Array.isArray(reportRows))
    return res.status(400).json({ error: 'Data laporan tidak valid.' });

  let targetSatkerId = null;
  if (req.user.role === 'op_satker') targetSatkerId = req.user.satkerId;
  else if (satkerId) targetSatkerId = parseInt(satkerId);
  else return res.status(400).json({ error: 'Harap pilih Satuan Kerja.' });

  try {
    const saktiData = {};
    let currentKodeAkun = '';

    for (const row of reportRows) {
      if (row[7] && /^\d{6}$/.test(String(row[7]).trim())) {
        currentKodeAkun = String(row[7]).trim();
      }

      if (
        currentKodeAkun &&
        row[13] &&
        /^\d{6}\./.test(String(row[13]).trim())
      ) {
        const uraian = String(row[13])
          .replace(/^\d{6}\.\s*/, '')
          .trim();
        const realisasi = parseInt(row[25], 10) || 0;

        if (!saktiData[currentKodeAkun]) saktiData[currentKodeAkun] = [];
        saktiData[currentKodeAkun].push({ uraian, realisasi });
      }
    }

    const rincianInDb = await prisma.spmRincian.findMany({
      where: {
        spm: { tahunAnggaran: parseInt(tahun), satkerId: targetSatkerId },
      },
      include: { kodeAkun: true, spm: { select: { nomorSpm: true } } },
    });

    let results = [];
    for (const rincian of rincianInDb) {
      const saktiItems = saktiData[rincian.kodeAkun.kode];
      let status = 'NOT_FOUND';
      let saktiAmount = null;
      let difference = null;

      if (saktiItems) {
        const matchedItem = saktiItems.find(
          (item) =>
            item.uraian.toLowerCase().trim() ===
            rincian.uraian.toLowerCase().trim()
        );
        if (matchedItem) {
          saktiAmount = matchedItem.realisasi;
          difference = rincian.jumlah - saktiAmount;
          status = difference === 0 ? 'MATCH' : 'MISMATCH';
        }
      }

      results.push({
        spmNomor: rincian.spm.nomorSpm,
        kodeProgram: rincian.kodeProgram,
        kodeKegiatan: rincian.kodeKegiatan,
        kodeKRO: rincian.kodeKRO,
        kodeRO: rincian.kodeRO,
        kodeKomponen: rincian.kodeKomponen,
        kodeSubkomponen: rincian.kodeSubkomponen,
        kodeAkun: rincian.kodeAkun.kode,
        kodeAkunNama: rincian.kodeAkun.nama,
        rincianUraian: rincian.uraian,
        appAmount: rincian.jumlah,
        saktiAmount,
        difference,
        status,
      });
    }
    res.status(200).json(results);
  } catch (error) {
    console.error('Validation Error:', error);
    res.status(500).json({ error: 'Gagal memvalidasi laporan SAKTI.' });
  }
};
