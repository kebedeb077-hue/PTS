import { Request, Response } from 'express';
import {
  getAllOffices,
  getOfficeById,
  createOffice,
  getOfficeDeposits,
  addDeposit,
  getOfficeShipments,
  addShipment,
} from '../services/office.service.js';
import { logger } from '../config/logger.js';

export async function listOffices(req: Request, res: Response): Promise<void> {
  try {
    const offices = await getAllOffices();
    res.json(offices);
  } catch (error) {
    logger.error(error, 'List offices error');
    res.status(500).json({ error: 'Failed to list offices' });
  }
}

export async function getOffice(req: Request, res: Response): Promise<void> {
  try {
    const office = await getOfficeById(parseInt(req.params.id));
    res.json(office);
  } catch (error) {
    logger.error(error, 'Get office error');
    res.status(404).json({ error: 'Office not found' });
  }
}

export async function create(req: Request, res: Response): Promise<void> {
  try {
    const office = await createOffice(req.body);
    res.status(201).json(office);
  } catch (error) {
    logger.error(error, 'Create office error');
    res.status(400).json({ error: (error as Error).message });
  }
}

export async function listDeposits(req: Request, res: Response): Promise<void> {
  try {
    const deposits = await getOfficeDeposits(parseInt(req.params.officeId));
    res.json(deposits);
  } catch (error) {
    logger.error(error, 'List deposits error');
    res.status(500).json({ error: 'Failed to list deposits' });
  }
}

export async function createDeposit(req: Request, res: Response): Promise<void> {
  try {
    const deposit = await addDeposit(parseInt(req.params.officeId), req.body);
    res.status(201).json(deposit);
  } catch (error) {
    logger.error(error, 'Create deposit error');
    res.status(400).json({ error: (error as Error).message });
  }
}

export async function listShipments(req: Request, res: Response): Promise<void> {
  try {
    const shipments = await getOfficeShipments(parseInt(req.params.officeId));
    res.json(shipments);
  } catch (error) {
    logger.error(error, 'List shipments error');
    res.status(500).json({ error: 'Failed to list shipments' });
  }
}

export async function createShipment(req: Request, res: Response): Promise<void> {
  try {
    const shipment = await addShipment(parseInt(req.params.officeId), req.body);
    res.status(201).json(shipment);
  } catch (error) {
    logger.error(error, 'Create shipment error');
    res.status(400).json({ error: (error as Error).message });
  }
}
