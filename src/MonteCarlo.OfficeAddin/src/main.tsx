import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './styles.css';

async function bootstrap() {
  await Office.onReady();
  const rootElement = document.getElementById('root');
  if (!rootElement) {
    throw new Error('Task pane root element was not found.');
  }

  ReactDOM.createRoot(rootElement).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>,
  );
}

bootstrap().catch((error) => {
  const rootElement = document.getElementById('root');
  if (rootElement) {
    rootElement.textContent = `MonteCarlo.XL failed to load: ${error instanceof Error ? error.message : String(error)}`;
  }
});
