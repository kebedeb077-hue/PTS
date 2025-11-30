import { Request, Response } from 'express';
import {
  getAllForms,
  getFormById,
  createForm,
  updateForm,
  getFormHistory,
} from '../services/form.service.js';
import { logger } from '../config/logger.js';

export async function listForms(req: Request, res: Response): Promise<void> {
  try {
    const limit = Math.min(Number(req.query.limit) || 50, 100);
    const offset = Number(req.query.offset) || 0;

    const result = await getAllForms(limit, offset);
    res.json(result);
  } catch (error) {
    logger.error(error, 'List forms error');
    res.status(500).json({ error: 'Failed to list forms' });
  }
}

export async function getForm(req: Request, res: Response): Promise<void> {
  try {
    const form = await getFormById(parseInt(req.params.id));
    res.json(form);
  } catch (error) {
    logger.error(error, 'Get form error');
    res.status(404).json({ error: 'Form not found' });
  }
}

export async function create(req: Request, res: Response): Promise<void> {
  if (!req.user) {
    res.status(401).json({ error: 'Unauthorized' });
    return;
  }

  try {
    const form = await createForm(req.body, req.user.userId);
    res.status(201).json(form);
  } catch (error) {
    logger.error(error, 'Create form error');
    res.status(400).json({ error: (error as Error).message });
  }
}

export async function update(req: Request, res: Response): Promise<void> {
  try {
    const form = await updateForm(parseInt(req.params.id), req.body);
    res.json(form);
  } catch (error) {
    logger.error(error, 'Update form error');
    res.status(400).json({ error: (error as Error).message });
  }
}

export async function getHistory(req: Request, res: Response): Promise<void> {
  try {
    const histories = await getFormHistory(parseInt(req.params.id));
    res.json(histories);
  } catch (error) {
    logger.error(error, 'Get form history error');
    res.status(500).json({ error: 'Failed to get history' });
  }
}
