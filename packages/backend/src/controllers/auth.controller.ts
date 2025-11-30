import { Request, Response } from 'express';
import { loginUser, changePassword } from '../services/auth.service.js';
import { logger } from '../config/logger.js';

export async function login(req: Request, res: Response): Promise<void> {
  try {
    const result = await loginUser(req.body);

    // Set refresh token as HttpOnly cookie
    res.cookie('refreshToken', result.refreshToken, {
      httpOnly: true,
      secure: process.env.NODE_ENV === 'production',
      sameSite: 'strict',
      maxAge: 7 * 24 * 60 * 60 * 1000, // 7 days
    });

    res.json({
      accessToken: result.accessToken,
      user: result.user,
    });
  } catch (error) {
    logger.error(error, 'Login error');
    res.status(401).json({ error: 'Invalid email or password' });
  }
}

export async function logout(req: Request, res: Response): Promise<void> {
  res.clearCookie('refreshToken');
  res.json({ success: true });
}

export async function changePass(req: Request, res: Response): Promise<void> {
  if (!req.user) {
    res.status(401).json({ error: 'Unauthorized' });
    return;
  }

  try {
    await changePassword(req.user.userId, req.body);
    res.json({ success: true, message: 'Password changed successfully' });
  } catch (error) {
    logger.error(error, 'Change password error');
    res.status(400).json({ error: (error as Error).message });
  }
}
