# Privacy & Database Notes

## Handling Access Database (.mdb) Files

The original Publications Tracking System uses Microsoft Access `.mdb` database files. When migrating to the new stack, follow these security guidelines:

### Never in Production

- **Do not** commit `.mdb` files to the repository. They are added to `.gitignore`.
- **Do not** expose `.mdb` files in web-accessible directories.
- **Do not** serve `.mdb` files directly from the web server.

### Migration Steps

1. **Export to CSV**
   ```bash
   ./scripts/export_mdb_to_csv.sh Publications_Tracking_System/Forms.mdb ./migrations/data
   ```

2. **Verify CSV content** for sensitive data:
   - Review for PII (personal identifying information)
   - Sanitize if needed before importing to production DB

3. **Import to PostgreSQL**
   - Use `packages/backend/src/migrations/import_mdb.ts` or similar script
   - Run in a controlled environment (not in production initially)

4. **Backup original `.mdb` files**
   - Store offline and encrypted
   - Never commit to public repositories

### Database Security in Production

#### PostgreSQL Best Practices

- Use strong passwords (20+ characters, mixed case, symbols)
- Enable SSL connections (`sslmode=require`)
- Restrict access by IP (VPC, security groups)
- Enable backups and point-in-time recovery
- Monitor for unusual queries and access patterns

#### Secrets Management

Store these in environment variables or secret management (never hardcoded):
- `DATABASE_URL` - Use service account (read + write only on necessary tables)
- `JWT_SECRET` - Rotate periodically (30+ chars)
- `JWT_REFRESH_SECRET` - Rotate periodically (30+ chars)

#### Access Control

- Enforce row-level security (RLS) if multi-tenant
- Use database roles with minimal permissions
- Audit all schema changes and data access
- Implement soft deletes for compliance

### Compliance Considerations

If the system handles sensitive data (user records, financial info, etc.):
- **GDPR**: Implement data export/delete endpoints
- **HIPAA**: Encrypt data at rest and in transit, enable audit logging
- **PCI-DSS**: If payment data, use tokenization (do not store card details)

### Local Development

- Use a non-production `.mdb` or sanitized copy
- Don't use real production data locally
- Add `migrations/data/*` to `.gitignore`

### Disaster Recovery

1. Backup PostgreSQL database daily
2. Test restore procedure monthly
3. Keep offline copies of critical `.mdb` files
4. Document recovery RTO (Recovery Time Objective) and RPO (Recovery Point Objective)

## Schema Mapping

See `infra/mdb-to-prisma.md` for detailed table mappings from Access to PostgreSQL.

## Questions?

Contact the system administrator or DevOps team for secure access to database credentials.
