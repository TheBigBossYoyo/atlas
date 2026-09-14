import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import 'katex/dist/katex.min.css';
import './index.css';
import App from './App';
import { runMigration } from './lib/migrateLocalStorage';

runMigration();

// UX-18 — <App /> wraps itself in <ToastProvider> (see App.tsx) so it stays a
// complete, self-contained tree for every caller, this entry point included.
createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
