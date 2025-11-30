# Publications Tracking System (PTS) - TypeScript Monorepo

A modern, secure full-stack rewrite of the classic ASP-based Publications Tracking System. Built with Node.js + Express + React + PostgreSQL, designed for scalability, security, and Vercel deployment.

## Quick Start

### Local Development with Docker

```bash
# Build and start all services (Postgres, Backend, Frontend)
docker-compose up --build

# Visit http://localhost:5173 for the frontend
# Backend API at http://localhost:5000
# Default login: admin@pts.local / admin123456
```

### Manual Setup

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Setup environment:**
   ```bash
   cp .env.example .env.local
   # Edit .env.local with your database URL and secrets
   ```

3. **Run migrations:**
   ```bash
   npm run db:migrate
   npm run db:seed
   ```

4. **Start development servers:**
   ```bash
   npm run dev
   ```

## Project Structure

```
/packages
  /backend       - Express API (Node.js + TypeScript)
  /frontend      - React app (Vite)
/infra
  /Dockerfile.backend
  /Dockerfile.frontend
/scripts
  /export_mdb_to_csv.sh    - Utility to convert .mdb files to CSV
  /package_website.sh       - Package the classic ASP site (legacy)
```

## Available Scripts

- `npm run dev` - Start both backend and frontend in development mode
- `npm run build` - Build both packages for production
- `npm run test` - Run all tests
- `npm run lint` - Lint all packages
- `npm run type-check` - TypeScript type checking
- `npm run db:migrate` - Run Prisma migrations
- `npm run db:seed` - Seed database with sample data

## API Endpoints

### Authentication
- `POST /api/auth/login` - Login user
- `POST /api/auth/logout` - Logout
- `POST /api/auth/change-password` - Change user password (requires auth)

### Forms
- `GET /api/forms` - List all forms
- `POST /api/forms` - Create new form (requires auth)
- `GET /api/forms/:id` - Get form details
- `PATCH /api/forms/:id` - Update form (requires auth)
- `GET /api/forms/:id/history` - Get form change history

### Offices
- `GET /api/offices` - List all offices
- `POST /api/offices` - Create office (requires auth)
- `GET /api/offices/:id` - Get office details with deposits and shipments
- `GET /api/offices/:officeId/deposits` - List office deposits
- `POST /api/offices/:officeId/deposits` - Record office deposit (requires auth)
- `GET /api/offices/:officeId/shipments` - List office shipments
- `POST /api/offices/:officeId/shipments` - Record office shipment (requires auth)

## Security Features

- ✅ JWT authentication with short-lived access tokens and refresh tokens
- ✅ Argon2 password hashing
- ✅ Rate limiting (auth endpoints: 5 req/15min, others: 100 req/15min)
- ✅ Input validation with Zod
- ✅ Helmet security headers (CSP, HSTS, X-Frame-Options)
- ✅ CORS with origin whitelist
- ✅ HttpOnly secure cookies for refresh tokens
- ✅ Input sanitization (XSS protection)
- ✅ Structured logging (Pino)
- ✅ HTTPS redirect in production
- ✅ Parameterized queries (Prisma ORM)

## Deployment

### Frontend (Vercel)

The frontend is ready for Vercel deployment. Simply connect your GitHub repo to Vercel:

1. Set environment variable: `VITE_API_URL` (pointing to your backend API)
2. Vercel will auto-deploy on push to `main`

### Backend (Railway, Render, or similar)

1. Set environment variables:
   - `DATABASE_URL` - PostgreSQL connection string
   - `JWT_SECRET` - Random 32+ char string
   - `JWT_REFRESH_SECRET` - Random 32+ char string
   - `API_PORT` - Port (default 5000)
   - `FRONTEND_URL` - Your frontend domain

2. Deploy using your hosting service's CLI or git integration

## Data Migration from Classic ASP

To migrate data from the original `.mdb` Access database:

1. Export `.mdb` tables to CSV:
   ```bash
   ./scripts/export_mdb_to_csv.sh Publications_Tracking_System/Forms.mdb ./migrations/data
   ```

2. Ensure `mdbtools` is installed:
   ```bash
   apt-get install mdbtools  # Ubuntu/Debian
   brew install mdbtools      # macOS
   ```

3. Use the CSV files to seed the new PostgreSQL database (refer to `packages/backend/src/seed.ts` for the migration script pattern)

## Environment Variables

Create `.env.local` in the root directory:

```bash
# Backend
NODE_ENV=development
DATABASE_URL=postgresql://pts_user:pts_password_dev@localhost:5432/pts_db
JWT_SECRET=dev_jwt_secret_change_in_prod_at_least_32_chars
JWT_REFRESH_SECRET=dev_refresh_secret_change_in_prod
API_PORT=5000
FRONTEND_URL=http://localhost:5173
LOG_LEVEL=info

# Frontend (in packages/frontend/.env)
VITE_API_URL=http://localhost:5000
```

## Testing

```bash
# Unit and integration tests (backend)
npm run test

# E2E tests (frontend)
npm run test -w frontend

# Watch mode
npm run test:watch
```

## Troubleshooting

**Port 5432 already in use:**
```bash
# Stop other postgres instances or use a different port in docker-compose.yml
```

**Database migrations fail:**
```bash
# Reset database (dev only!)
npm run db:migrate:reset
```

**Frontend can't connect to backend:**
- Check `VITE_API_URL` environment variable
- Ensure backend is running on the correct port
- Check CORS headers in browser console

## CI/CD

GitHub Actions workflows are configured:

- **CI (`.github/workflows/ci.yml`)**: Runs on push to `main`
  - Lint, type-check, test, build
  - Requires passing tests before merge

- **Deploy (`.github/workflows/deploy.yml`)**: Auto-deploys on successful CI
  - Deploys frontend to Vercel
  - Optional: Configure backend deployment

## Documentation

- `DEPLOY.md` - Deployment checklist for the classic ASP site
- `PRIVACY_AND_DB_NOTES.md` - Database security and migration notes
- `infra/mdb-to-prisma.md` - Schema mapping from `.mdb` to Prisma

## License

Internal use only. Built for the Publications Tracking System project.
