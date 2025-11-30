import { Router } from 'express';
import * as officeController from '../controllers/office.controller.js';

const router = Router();

router.get('/', officeController.listOffices);
router.post('/', officeController.create);
router.get('/:id', officeController.getOffice);

// Deposits
router.get('/:officeId/deposits', officeController.listDeposits);
router.post('/:officeId/deposits', officeController.createDeposit);

// Shipments
router.get('/:officeId/shipments', officeController.listShipments);
router.post('/:officeId/shipments', officeController.createShipment);

export default router;
