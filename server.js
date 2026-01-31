import { serve } from 'bun';

const server = serve({
    port: 3030,
    fetch(req) {
        const url = new URL(req.url);
        let filePath = url.pathname;

        if (filePath === '/') filePath = '/index.html';
        
        // 忽略 .well-known 路径
        if (filePath.includes('.well-known')) {
            return new Response('Not Found', { status: 404 });
        }
        
        // 处理 node_modules 中的模块
        if (filePath.startsWith('/node_modules/')) {
            try {
                const file = Bun.file('.' + filePath);
                if (file.exists()) {
                    const content = file.textSync();
                    const ext = filePath.split('.').pop();
                    const contentType = {
                        'js': 'application/javascript',
                        'mjs': 'application/javascript',
                        'css': 'text/css',
                        'json': 'application/json'
                    }[ext] || 'application/javascript';
                    return new Response(content, {
                        headers: { 'Content-Type': contentType }
                    });
                }
            } catch (e) {}
        }
        
        try {
            const file = Bun.file('.' + filePath);
            if (file.exists()) {
                return new Response(file);
            }
        } catch (e) {}
        
        return new Response('Not Found', { status: 404 });
    },
});

console.log('🚀 Server running at http://localhost:3030');
