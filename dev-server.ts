import { app } from './api/index';
import { createServer as createViteServer } from 'vite';

async function startDevServer() {
  const PORT = 3000;

  console.log('Starting development server with Vite middleware...');

  const viteDevServer = await createViteServer({
    server: { middlewareMode: true },
    appType: "spa",
  });

  // Attach Vite's dev server middleware
  app.use(viteDevServer.middlewares);

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Development server running on http://localhost:${PORT}`);
  });
}

startDevServer().catch((err) => {
  console.error('Failed to start development server:', err);
});
