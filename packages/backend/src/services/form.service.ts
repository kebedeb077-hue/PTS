import { PrismaClient } from '@prisma/client';
import { z } from 'zod';

const prisma = new PrismaClient();

const createFormSchema = z.object({
  title: z.string().min(1),
  description: z.string().optional(),
  data: z.record(z.any()).optional(),
});

const updateFormSchema = z.object({
  title: z.string().min(1).optional(),
  description: z.string().optional(),
  data: z.record(z.any()).optional(),
  status: z.enum(['draft', 'submitted', 'approved']).optional(),
});

export async function getAllForms(limit: number = 50, offset: number = 0) {
  const [forms, total] = await Promise.all([
    prisma.form.findMany({
      take: limit,
      skip: offset,
      orderBy: { createdAt: 'desc' },
    }),
    prisma.form.count(),
  ]);

  return { forms, total, limit, offset };
}

export async function getFormById(id: number) {
  const form = await prisma.form.findUnique({
    where: { id },
    include: { histories: true },
  });

  if (!form) {
    throw new Error('Form not found');
  }

  return form;
}

export async function createForm(input: unknown, createdBy: number) {
  const data = createFormSchema.parse(input);

  const form = await prisma.form.create({
    data: {
      ...data,
      createdBy,
    },
  });

  // Log history
  await prisma.formHistory.create({
    data: {
      formId: form.id,
      action: 'created',
      changes: data,
    },
  });

  return form;
}

export async function updateForm(id: number, input: unknown) {
  const updates = updateFormSchema.parse(input);

  const form = await prisma.form.update({
    where: { id },
    data: updates,
  });

  // Log history
  await prisma.formHistory.create({
    data: {
      formId: form.id,
      action: 'updated',
      changes: updates,
    },
  });

  return form;
}

export async function getFormHistory(formId: number) {
  const histories = await prisma.formHistory.findMany({
    where: { formId },
    orderBy: { createdAt: 'desc' },
  });

  return histories;
}
