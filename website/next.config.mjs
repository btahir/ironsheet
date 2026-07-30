import { createMDX } from 'fumadocs-mdx/next';

const withMDX = createMDX();

/** @type {import('next').NextConfig} */
const config = {
  reactStrictMode: true,
  turbopack: {
    // The demo intentionally exercises the workspace sources so Vercel deploys
    // the same browser package implementation that the repository tests.
    root: new URL('..', import.meta.url).pathname,
  },
};

export default withMDX(config);
