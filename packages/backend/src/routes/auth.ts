import { Router } from 'express';
import { authMiddleware } from '../middleware/auth.js';
import * as authController from '../controllers/auth.controller.js';
import { authLimiter } from '../middleware/security.js';

const router = Router();

router.post('/login', authLimiter, authController.login);
router.post('/logout', authMiddleware, authController.logout);
router.post('/change-password', authMiddleware, authController.changePass);

export default router;
