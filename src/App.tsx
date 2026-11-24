/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useEffect, useRef, useState } from 'react';
import { initReportBuilderEngine } from './reportEngine';
import { REPORT_HTML_TEMPLATE } from './reportTemplate';

export default function App() {
  const containerRef = useRef<HTMLDivElement>(null);
  const [resetKey, setResetKey] = useState(0);

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    el.innerHTML = REPORT_HTML_TEMPLATE;
    initReportBuilderEngine(() => {
      setResetKey((prev) => prev + 1);
    });
  }, [resetKey]);

  return (
    <div
      id="app-root-wrapper"
      ref={containerRef}
      className="text-slate-800 h-screen flex flex-col overflow-hidden"
    />
  );
}
