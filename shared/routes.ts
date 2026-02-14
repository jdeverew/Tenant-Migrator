import { z } from 'zod';
import { insertProjectSchema, insertItemSchema, migrationProjects, migrationItems } from './schema';

export const errorSchemas = {
  validation: z.object({
    message: z.string(),
    field: z.string().optional(),
  }),
  notFound: z.object({
    message: z.string(),
  }),
  internal: z.object({
    message: z.string(),
  }),
};

export const api = {
  projects: {
    list: {
      method: 'GET' as const,
      path: '/api/projects' as const,
      responses: {
        200: z.array(z.custom<typeof migrationProjects.$inferSelect>()),
      },
    },
    get: {
      method: 'GET' as const,
      path: '/api/projects/:id' as const,
      responses: {
        200: z.custom<typeof migrationProjects.$inferSelect>(),
        404: errorSchemas.notFound,
      },
    },
    create: {
      method: 'POST' as const,
      path: '/api/projects' as const,
      input: insertProjectSchema,
      responses: {
        201: z.custom<typeof migrationProjects.$inferSelect>(),
        400: errorSchemas.validation,
      },
    },
    update: {
      method: 'PATCH' as const,
      path: '/api/projects/:id' as const,
      input: insertProjectSchema.partial(),
      responses: {
        200: z.custom<typeof migrationProjects.$inferSelect>(),
        404: errorSchemas.notFound,
      },
    },
    delete: {
      method: 'DELETE' as const,
      path: '/api/projects/:id' as const,
      responses: {
        204: z.void(),
        404: errorSchemas.notFound,
      },
    },
    stats: {
      method: 'GET' as const,
      path: '/api/projects/:id/stats' as const,
      responses: {
        200: z.object({
          total: z.number(),
          pending: z.number(),
          inProgress: z.number(),
          completed: z.number(),
          failed: z.number(),
        }),
      },
    },
  },
  items: {
    list: {
      method: 'GET' as const,
      path: '/api/projects/:projectId/items' as const,
      responses: {
        200: z.array(z.custom<typeof migrationItems.$inferSelect>()),
      },
    },
    create: {
      method: 'POST' as const,
      path: '/api/projects/:projectId/items' as const,
      input: insertItemSchema.omit({ projectId: true }),
      responses: {
        201: z.custom<typeof migrationItems.$inferSelect>(),
        400: errorSchemas.validation,
      },
    },
    update: {
      method: 'PATCH' as const,
      path: '/api/items/:id' as const,
      input: insertItemSchema.partial(),
      responses: {
        200: z.custom<typeof migrationItems.$inferSelect>(),
        404: errorSchemas.notFound,
      },
    },
    delete: {
      method: 'DELETE' as const,
      path: '/api/items/:id' as const,
      responses: {
        204: z.void(),
        404: errorSchemas.notFound,
      },
    },
  },
};

export function buildUrl(path: string, params?: Record<string, string | number>): string {
  let url = path;
  if (params) {
    Object.entries(params).forEach(([key, value]) => {
      url = url.replace(`:${key}`, String(value));
    });
  }
  return url;
}
