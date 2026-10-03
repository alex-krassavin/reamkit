import { defineCollection } from 'astro:content';
import { z } from 'astro/zod';
import { docsLoader } from '@astrojs/starlight/loaders';
import { docsSchema } from '@astrojs/starlight/schema';

export const collections = {
	docs: defineCollection({
		loader: docsLoader(),
		schema: docsSchema({
			extend: z.object({
				// The page draws its own <h1>; Starlight's title is left out (PageTitle.astro).
				ownHeading: z.boolean().default(false),
			}),
		}),
	}),
};
