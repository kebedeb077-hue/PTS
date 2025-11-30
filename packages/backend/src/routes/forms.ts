import { Router } from 'express';
import { authMiddleware } from '../middleware/auth.js';
import * as formController from '../controllers/form.controller.js';

const router = Router();

router.get('/', formController.listForms);
router.post('/', authMiddleware, formController.create);
router.get('/:id', formController.getForm);
router.patch('/:id', authMiddleware, formController.update);
router.get('/:id/history', formController.getHistory);

export default router;
