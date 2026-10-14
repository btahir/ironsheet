"use client";

import Script from "next/script";
import { useSyncExternalStore } from "react";

const subscribe = () => () => {};
const serverSnapshot = () => false;
const publicHosts = new Set("ironsheetdocs.vercel.app".split(","));
const browserSnapshot = () => publicHosts.has(window.location.hostname);

export default function Analytics() {
  const enabled = useSyncExternalStore(subscribe, browserSnapshot, serverSnapshot);
  if (!enabled) return null;
  return (
    <Script
      src="https://project-analytics.shiptiffin.app/script.js"
      data-website-id="a8a3d0a5-cc28-42e2-9969-a4858642ba2b"
      data-domains="ironsheetdocs.vercel.app"
      data-exclude-search="true"
      data-exclude-hash="true"
      strategy="afterInteractive"
    />
  );
}
