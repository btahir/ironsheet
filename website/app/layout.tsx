import { RootProvider } from 'fumadocs-ui/provider/next';
import './global.css';
import type { Metadata } from 'next';
import { Inter } from 'next/font/google';
import Script from 'next/script';
import { appDescription, appName, appTagline, siteUrl } from '@/lib/shared';

const inter = Inter({
  subsets: ['latin'],
});

const googleAnalyticsId = process.env.NEXT_PUBLIC_GOOGLE_ANALYTICS_ID;

export const metadata: Metadata = {
  metadataBase: new URL(siteUrl),
  title: {
    default: `${appName} — ${appTagline}`,
    template: `%s — ${appName}`,
  },
  description: appDescription,
  applicationName: appName,
  manifest: '/site.webmanifest',
  icons: {
    icon: [
      { url: '/favicon-16x16.png', sizes: '16x16', type: 'image/png' },
      { url: '/favicon-32x32.png', sizes: '32x32', type: 'image/png' },
    ],
    apple: '/apple-icon.png',
  },
  openGraph: {
    title: `${appName} — ${appTagline}`,
    description: appDescription,
    siteName: appName,
    images: ['/ironsheet-opengraph.png'],
    type: 'website',
  },
  twitter: {
    card: 'summary_large_image',
    title: `${appName} — ${appTagline}`,
    description: appDescription,
    images: ['/ironsheet-opengraph.png'],
  },
};

export default function Layout({ children }: LayoutProps<'/'>) {
  return (
    <html lang="en" className={inter.className} suppressHydrationWarning>
      <body className="flex flex-col min-h-screen">
        <RootProvider>{children}</RootProvider>
        {googleAnalyticsId ? (
          <>
            <Script id="google-analytics" strategy="afterInteractive">
              {`
                window.dataLayer = window.dataLayer || [];
                function gtag(){dataLayer.push(arguments);}
                gtag('js', new Date());
                gtag('config', ${JSON.stringify(googleAnalyticsId)});
              `}
            </Script>
            <Script
              src={`https://www.googletagmanager.com/gtag/js?id=${encodeURIComponent(googleAnalyticsId)}`}
              strategy="afterInteractive"
            />
          </>
        ) : null}
      </body>
    </html>
  );
}
