import express, { Express, Request, Response, NextFunction } from 'express';
import cors from 'cors';
import pinoHttp from 'pino-http';
import { getEnv } from './config/env.js';
import { logger } from './config/logger.js';
import { securityMiddleware, generalLimiter } from './middleware/security.js';
import authRoutes from './routes/auth.js';
import formsRoutes from './routes/forms.js';
import officesRoutes from './routes/offices.js';

const env = getEnv();

export function createApp(): Express {
  const app = express();

  // Trust proxy in production
  if (env.NODE_ENV === 'production') {
    app.set('trust proxy', 1);
  }

  // Logging
  app.use(pinoHttp({ logger }));

  // Security
  app.use(...securityMiddleware());

  // CORS
  app.use(
    cors({
      origin: env.FRONTEND_URL,
      credentials: true,
    })
  );

  // Body parser
  app.use(express.json({ limit: '10mb' }));
  app.use(express.urlencoded({ limit: '10mb', extended: true }));

  // Rate limiting
  app.use(generalLimiter);

  // Health check
  app.get('/health', (req: Request, res: Response) => {
    res.json({ status: 'ok' });
  });

  // API routes
  app.use('/api/auth', authRoutes);
  app.use('/api/forms', formsRoutes);
  app.use('/api/offices', officesRoutes);

  // 404
  app.use((req: Request, res: Response) => {
    res.status(404).json({ error: 'Not found' });
  });

  // Error handler
  app.use((err: any, req: Request, res: Response, next: NextFunction) => {
    logger.error(err, 'Unhandled error');
    res.status(500).json({ error: 'Internal server error' });
  });

  return app;
}
