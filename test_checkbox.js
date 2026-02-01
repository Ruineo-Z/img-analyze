import { chromium } from 'playwright';

(async () => {
    const browser = await chromium.launch();
    const context = await browser.newContext();
    const page = await context.newPage();

    // 监听控制台消息
    page.on('console', msg => {
        if (msg.type() === 'error') {
            console.log('Console error:', msg.text());
        }
    });

    // 监听页面错误
    page.on('pageerror', err => {
        console.log('Page error:', err.message);
    });

    try {
        await page.goto('http://127.0.0.1:3030/');
        await page.waitForLoadState('networkidle');

        // 等待表单加载
        await page.waitForSelector('#formSection', { state: 'visible', timeout: 5000 }).catch(() => {
            console.log('Form section not visible, clicking upload area...');
            return page.click('#uploadArea');
        });

        // 等待表单可见
        await page.waitForSelector('#formSection', { state: 'visible', timeout: 10000 });

        // 点击环境接触史的"粉尘"选项
        const粉尘Checkbox = await page.locator('input[name="environment"][value="粉尘"]');
        await粉尘Checkbox.check();
        const粉尘Checked = await粉尘Checkbox.isChecked();
        console.log('粉尘 checked:', 粉尘Checked);

        // 点击"有害气体"选项
        const有害气体Checkbox = await page.locator('input[name="environment"][value="有害气体"]');
        await有害气体Checkbox.check();
        const有害气体Checked = await有害气体Checkbox.isChecked();
        console.log('有害气体 checked:', 有害气体Checked);

        // 再次检查"粉尘"是否还被选中
        const粉尘StillChecked = await粉尘Checkbox.isChecked();
        console.log('粉尘 still checked:', 粉尘StillChecked);

        // 检查所有environment checkbox的状态
        console.log('\n所有environment checkbox状态:');
        const环境Checkboxes = await page.locator('input[name="environment"]').all();
        for (const cb of 环境Checkboxes) {
            const value = await cb.getAttribute('value');
            const checked = await cb.isChecked();
            console.log(`  ${value}: ${checked}`);
        }

    } catch (error) {
        console.log('Error:', error.message);
    }

    await browser.close();
})();
