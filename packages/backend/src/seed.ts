import { PrismaClient } from '@prisma/client';
import * as argon2 from 'argon2';

const prisma = new PrismaClient();

async function main() {
  console.log('🌱 Seeding database...');

  // Clear existing data
  await prisma.$executeRawUnsafe('DELETE FROM "users"');
  await prisma.$executeRawUnsafe('DELETE FROM "offices"');
  await prisma.$executeRawUnsafe('DELETE FROM "forms"');

  // Create default admin user
  const adminPassword = await argon2.hash('admin123456', {
    type: argon2.argon2id,
    memoryCost: 19456,
    timeCost: 2,
    parallelism: 1,
  });

  const admin = await prisma.user.create({
    data: {
      email: 'admin@pts.local',
      password: adminPassword,
      name: 'Administrator',
      role: 'admin',
    },
  });

  console.log(`✓ Created admin user: ${admin.email}`);

  // Create sample offices
  const offices = await Promise.all([
    prisma.office.create({
      data: {
        name: 'Main Office',
        location: 'Addis Ababa',
        manager: 'John Doe',
      },
    }),
    prisma.office.create({
      data: {
        name: 'Northern Office',
        location: 'Mekelle',
        manager: 'Jane Smith',
      },
    }),
    prisma.office.create({
      data: {
        name: 'Southern Office',
        location: 'Hawassa',
        manager: 'Bob Johnson',
      },
    }),
  ]);

  console.log(`✓ Created ${offices.length} sample offices`);

  // Create sample forms
  const forms = await Promise.all(
    Array.from({ length: 5 }, (_, i) =>
      prisma.form.create({
        data: {
          title: `Sample Form ${i + 1}`,
          description: `This is a sample form for testing`,
          status: i % 2 === 0 ? 'draft' : 'submitted',
          createdBy: admin.id,
          data: {
            field1: `value${i + 1}`,
            field2: 'sample',
          },
        },
      })
    )
  );

  console.log(`✓ Created ${forms.length} sample forms`);

  // Create sample deposits
  for (const office of offices) {
    await prisma.officeDeposit.create({
      data: {
        officeId: office.id,
        amount: Math.random() * 10000,
        date: new Date(),
        notes: 'Sample deposit',
      },
    });
  }

  console.log('✓ Created sample deposits');

  console.log('✅ Database seeded successfully');
}

main()
  .catch((e) => {
    console.error(e);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
