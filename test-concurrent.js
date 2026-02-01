import { writeFileSync, existsSync, mkdirSync, readFileSync } from 'fs';
import { join } from 'path';

// 复制一份 server.js 进行测试
const DATA_DIR = '/tmp/test-data';
const RECORDS_FILE = join(DATA_DIR, 'records.json');

// 初始化
if (!existsSync(DATA_DIR)) mkdirSync(DATA_DIR, { recursive: true });
if (!existsSync(RECORDS_FILE)) writeFileSync(RECORDS_FILE, JSON.stringify([], null, 2));

// 保存队列
const saveQueue = [];

function getAllRecords() {
    try {
        return JSON.parse(readFileSync(RECORDS_FILE, 'utf-8'));
    } catch (e) {
        return [];
    }
}

async function saveAllRecords(records) {
    return new Promise((resolve) => {
        saveQueue.push({ records, resolve });
        if (saveQueue.length === 1) processQueue();
    });
}

async function processQueue() {
    if (saveQueue.length === 0) return;
    const { records, resolve } = saveQueue[0];
    
    writeFileSync(RECORDS_FILE, JSON.stringify(records, null, 2));
    console.log(`保存完成，队列剩余: ${saveQueue.length - 1}`);
    resolve();
    
    saveQueue.shift();
    if (saveQueue.length > 0) processQueue();
}

// 模拟 10 个并发请求
async function testConcurrent() {
    console.log('开始并发测试...\n');
    
    const promises = [];
    for (let i = 0; i < 10; i++) {
        promises.push((async () => {
            const records = getAllRecords();
            records.push({ 
                id: `test-${i}`, 
                name: `用户${i}`,
                age: 20 + i,
                time: new Date().toISOString() 
            });
            await saveAllRecords(records);
        })());
    }
    
    await Promise.all(promises);
    
    // 验证结果
    setTimeout(() => {
        const final = getAllRecords();
        console.log(`\n最终记录数: ${final.length}`);
        console.log('所有用户数据:');
        final.forEach(r => console.log(`  - ${r.name}, 年龄: ${r.age}`));
        
        if (final.length === 10) {
            console.log('\n✅ 并发测试通过！所有数据都已保存，没有丢失。');
        } else {
            console.log(`\n❌ 失败！期望 10 条，实际 ${final.length} 条`);
        }
    }, 500);
}

testConcurrent();
