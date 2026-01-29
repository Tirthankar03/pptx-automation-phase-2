import { z } from "zod";

export const MAX_CHARS = {
  brief: 96,
  impact: 84,
  effort: 2,
  remarks: 60,
  eta: 10,
} as const;

export const rowSchema = z.object({
  brief: z.string().max(MAX_CHARS.brief, `Max ${MAX_CHARS.brief} characters`),
  impact: z.string().max(MAX_CHARS.impact, `Max ${MAX_CHARS.impact} characters`),
  effort: z.string().max(MAX_CHARS.effort, `Max ${MAX_CHARS.effort} characters`),
  remarks: z.string().max(MAX_CHARS.remarks, `Max ${MAX_CHARS.remarks} characters`),
  eta: z.string().max(MAX_CHARS.eta, `Max ${MAX_CHARS.eta} characters`),
  status: z.string(),
});

export type RowValidationErrors = {
  [K in keyof typeof MAX_CHARS]?: string;
} & Record<string, string | undefined>;

export const validateRow = (row: z.infer<typeof rowSchema>): RowValidationErrors => {
  const result = rowSchema.safeParse(row);
  const errors: Record<string, string | undefined> = {};

  if (!result.success) {
    result.error.errors.forEach((err) => {
      const field = err.path[0] as string;
      errors[field] = err.message;
    });
  }

  return errors as RowValidationErrors;
};

export const validateAllRows = (rows: z.infer<typeof rowSchema>[]): RowValidationErrors[] => {
  return rows.map(validateRow);
};

export const hasValidationErrors = (errors: RowValidationErrors[]): boolean => {
  return errors.some((rowErrors) => Object.keys(rowErrors).length > 0);
};
