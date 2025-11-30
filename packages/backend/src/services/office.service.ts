import { PrismaClient } from '@prisma/client';
import { z } from 'zod';

const prisma = new PrismaClient();

const createOfficeSchema = z.object({
  name: z.string().min(1),
  location: z.string().optional(),
  manager: z.string().optional(),
});

const createDepositSchema = z.object({
  amount: z.number().positive(),
  date: z.string().datetime().or(z.date()),
  notes: z.string().optional(),
});

const createShipmentSchema = z.object({
  itemName: z.string().min(1),
  quantity: z.number().int().positive(),
  date: z.string().datetime().or(z.date()),
  notes: z.string().optional(),
});

export async function getAllOffices() {
  return prisma.office.findMany({
    orderBy: { name: 'asc' },
  });
}

export async function getOfficeById(id: number) {
  const office = await prisma.office.findUnique({
    where: { id },
    include: {
      deposits: { orderBy: { date: 'desc' } },
      shipments: { orderBy: { date: 'desc' } },
    },
  });

  if (!office) {
    throw new Error('Office not found');
  }

  return office;
}

export async function createOffice(input: unknown) {
  const data = createOfficeSchema.parse(input);
  return prisma.office.create({ data });
}

export async function getOfficeDeposits(officeId: number) {
  const deposits = await prisma.officeDeposit.findMany({
    where: { officeId },
    orderBy: { date: 'desc' },
  });

  return deposits;
}

export async function addDeposit(officeId: number, input: unknown) {
  const data = createDepositSchema.parse(input);

  // Verify office exists
  const office = await prisma.office.findUnique({ where: { id: officeId } });
  if (!office) {
    throw new Error('Office not found');
  }

  return prisma.officeDeposit.create({
    data: { officeId, ...data },
  });
}

export async function getOfficeShipments(officeId: number) {
  const shipments = await prisma.officeShipment.findMany({
    where: { officeId },
    orderBy: { date: 'desc' },
  });

  return shipments;
}

export async function addShipment(officeId: number, input: unknown) {
  const data = createShipmentSchema.parse(input);

  const office = await prisma.office.findUnique({ where: { id: officeId } });
  if (!office) {
    throw new Error('Office not found');
  }

  return prisma.officeShipment.create({
    data: { officeId, ...data },
  });
}
