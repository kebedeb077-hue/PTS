# Schema Mapping: `.mdb` to Prisma (PostgreSQL)

This document maps tables from the original Access database (`.mdb`) to the new PostgreSQL/Prisma schema.

## Original `.mdb` Tables → New Schema

### Users
| Access Field | Type | Prisma Model | Notes |
|---|---|---|---|
| UserID | AutoNumber | User.id | Primary key |
| Email | Text | User.email | Unique constraint |
| Password | Text | User.password | Must be re-hashed with Argon2 |
| Name | Text | User.name | Optional |
| Role | Text | User.role | "user" or "admin" |

### Forms
| Access Field | Type | Prisma Model | Notes |
|---|---|---|---|
| FormID | AutoNumber | Form.id | Primary key |
| FormTitle | Text | Form.title | |
| FormDescription | Text | Form.description | Optional |
| FormData | Memo | Form.data | Store as JSON |
| Status | Text | Form.status | "draft", "submitted", "approved" |
| CreatedBy | Number | Form.createdBy | FK to User.id |
| CreatedDate | DateTime | Form.createdAt | Indexed |

### FormHistory
| Access Field | Type | Prisma Model | Notes |
|---|---|---|---|
| HistoryID | AutoNumber | FormHistory.id | Primary key |
| FormID | Number | FormHistory.formId | FK to Form.id |
| Action | Text | FormHistory.action | "created", "updated", "submitted", "approved" |
| Changes | Memo | FormHistory.changes | Store as JSON |
| ChangedDate | DateTime | FormHistory.createdAt | Indexed |

### Offices
| Access Field | Type | Prisma Model | Notes |
|---|---|---|---|
| OfficeID | AutoNumber | Office.id | Primary key |
| OfficeName | Text | Office.name | Unique constraint |
| Location | Text | Office.location | Optional |
| Manager | Text | Office.manager | Optional |

### OfficeDeposits
| Access Field | Type | Prisma Model | Notes |
|---|---|---|---|
| DepositID | AutoNumber | OfficeDeposit.id | Primary key |
| OfficeID | Number | OfficeDeposit.officeId | FK to Office.id |
| Amount | Currency | OfficeDeposit.amount | Stored as Float |
| DepositDate | DateTime | OfficeDeposit.date | |
| Notes | Memo | OfficeDeposit.notes | Optional |

### OfficeShipments
| Access Field | Type | Prisma Model | Notes |
|---|---|---|---|
| ShipmentID | AutoNumber | OfficeShipment.id | Primary key |
| OfficeID | Number | OfficeShipment.officeId | FK to Office.id |
| ItemName | Text | OfficeShipment.itemName | |
| Quantity | Number | OfficeShipment.quantity | Int |
| ShipmentDate | DateTime | OfficeShipment.date | |
| Notes | Memo | OfficeShipment.notes | Optional |

### Shipments (General)
| Access Field | Type | Prisma Model | Notes |
|---|---|---|---|
| ShipmentID | AutoNumber | Shipment.id | Primary key |
| Item | Text | Shipment.item | |
| Quantity | Number | Shipment.quantity | Int |
| Source | Text | Shipment.source | Optional |
| Destination | Text | Shipment.destination | Optional |
| Status | Text | Shipment.status | "pending", "shipped", "delivered" |
| ShipmentDate | DateTime | Shipment.date | |

## Migration Notes

1. **Password Hashing**
   - Old: Passwords likely in plain text or simple hashing
   - New: Must re-hash all passwords with Argon2 during migration
   - Create temporary passwords for users if original hash is lost

2. **DateTime Fields**
   - Access uses local timezone; PostgreSQL uses UTC
   - Adjust migration script to convert properly

3. **Memo/Text Fields**
   - Large Memo fields in Access → JSON in Prisma for flexibility
   - Consider compression if very large

4. **Relationships**
   - Set `onDelete: Cascade` to maintain referential integrity
   - Ensure FK constraints during migration

5. **Data Validation**
   - Run data quality checks after migration
   - Validate all emails, dates, numeric values
   - Check for orphaned records

## Example CSV Import Script

```typescript
// In packages/backend/src/migrations/import_mdb.ts
import { PrismaClient } from '@prisma/client';
import fs from 'fs';
import csv from 'csv-parser';

const prisma = new PrismaClient();

async function importUsers() {
  return new Promise((resolve, reject) => {
    fs.createReadStream('migrations/data/Users.csv')
      .pipe(csv())
      .on('data', async (row) => {
        // Transform and validate row
        // await prisma.user.create({ ... })
      })
      .on('end', resolve)
      .on('error', reject);
  });
}

// Main
(async () => {
  await importUsers();
  // ... import other tables
  await prisma.$disconnect();
})();
```

## Rollback Plan

If migration fails:
1. Keep original `.mdb` files backed up offline
2. Restore PostgreSQL from pre-migration backup
3. Analyze error logs and fix migration script
4. Re-run migration in staging environment first

## Post-Migration Validation

```sql
-- Verify record counts
SELECT 'users' as table_name, COUNT(*) as count FROM users
UNION ALL
SELECT 'forms', COUNT(*) FROM forms
UNION ALL
SELECT 'offices', COUNT(*) FROM offices
-- ... etc

-- Check for orphaned records
SELECT * FROM form_histories WHERE "formId" NOT IN (SELECT id FROM forms);
SELECT * FROM office_deposits WHERE "officeId" NOT IN (SELECT id FROM offices);
```
