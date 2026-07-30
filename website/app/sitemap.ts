import type { MetadataRoute } from 'next';
import { siteUrl } from '@/lib/shared';

export default function sitemap(): MetadataRoute.Sitemap {
  return [
    {
      url: new URL('/', siteUrl).toString(),
      changeFrequency: 'monthly',
      priority: 1,
    },
    {
      url: new URL('/tools/update-excel-from-csv', siteUrl).toString(),
      changeFrequency: 'monthly',
      priority: 0.9,
    },
    {
      url: new URL('/docs', siteUrl).toString(),
      changeFrequency: 'weekly',
      priority: 0.8,
    },
  ];
}
