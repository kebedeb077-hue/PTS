import 'dotenv/config';
import { createApp } from './app.js';
import { getEnv } from './config/env.js';
import { logger } from './config/logger.js';
import { PrismaClient } from '@prisma/client';

const env = getEnv();
const prisma = new PrismaClient();

async function main() {
  const app = createApp();

  // Verify database connection
  try {
    await prisma.$queryRaw`SELECT 1`;
    logger.info('Database connection established');
  } catch (error) {
    logger.error(error, 'Failed to connect to database');
    process.exit(1);
  }

  app.listen(env.API_PORT, () => {
    logger.info(`Server listening on port ${env.API_PORT}`);
  });
}

main().catch((err) => {
  logger.error(err, 'Fatal error');
  process.exit(1);
});
