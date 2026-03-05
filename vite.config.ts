import { defineConfig } from 'vite'

export default defineConfig({
	server: {
		allowedHosts: ['.loca.lt', 'localhost', '127.0.0.1'],
		proxy: {
			'/api': {
				target: 'http://localhost:3001',
				changeOrigin: true
			}
		}
	}
})
