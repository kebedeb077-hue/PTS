import { PrismaClient } from '@prisma/client';
import * as argon2 from 'argon2';
import { z } from 'zod';
import { signAccessToken, signRefreshToken } from '../config/jwt.js';
import { logger } from '../config/logger.js';

const prisma = new PrismaClient();

const loginSchema = z.object({
  email: z.string().email(),
  password: z.string().min(6),
});

const changePasswordSchema = z.object({
  currentPassword: z.string().min(6),
  newPassword: z.string().min(8),
});

export async function loginUser(credentials: unknown) {
  const { email, password } = loginSchema.parse(credentials);

  const user = await prisma.user.findUnique({
    where: { email },
  });

  if (!user) {
    logger.warn({ email }, 'Login attempt for non-existent user');
    throw new Error('Invalid email or password');
  }

  const isValid = await argon2.verify(user.password, password);
  if (!isValid) {
    logger.warn({ email }, 'Failed login attempt');
    throw new Error('Invalid email or password');
  }

  logger.info({ email }, 'User logged in');

  const accessToken = signAccessToken({
    userId: user.id,
    email: user.email,
    role: user.role,
  });

  const refreshToken = signRefreshToken({
    userId: user.id,
    email: user.email,
    role: user.role,
  });

  return { accessToken, refreshToken, user: { id: user.id, email: user.email, name: user.name, role: user.role } };
}

export async function changePassword(userId: number, input: unknown) {
  const { currentPassword, newPassword } = changePasswordSchema.parse(input);

  const user = await prisma.user.findUnique({
    where: { id: userId },
  });

  if (!user) {
    throw new Error('User not found');
  }

  const isValid = await argon2.verify(user.password, currentPassword);
  if (!isValid) {
    logger.warn({ userId }, 'Failed password change attempt');
    throw new Error('Current password is incorrect');
  }

  const hashedPassword = await argon2.hash(newPassword, {
    type: argon2.argon2id,
    memoryCost: 19456,
    timeCost: 2,
    parallelism: 1,
  });

  await prisma.user.update({
    where: { id: userId },
    data: { password: hashedPassword },
  });

  logger.info({ userId }, 'Password changed');
  return { success: true };
}
