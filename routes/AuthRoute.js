const express = require('express');
const router = express.Router();
const authController = require('../controllers/UserController');
const { protect } = require('../middleware/AuthMiddleware');

router.post('/login', authController.login);
router.get('/me', protect, (req, res) => {
  res.status(200).json({ success: true, user: req.user });
});

module.exports = router;
