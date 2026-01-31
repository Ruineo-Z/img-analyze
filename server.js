import { serve } from 'bun';
import { writeFileSync, existsSync, mkdirSync, readFileSync, createWriteStream } from 'fs';
import { join, dirname } from 'path';
import XLSX from 'xlsx';

// 数据存储目录
const DATA_DIR = join(process.cwd(), 'data');
const RECORDS_FILE = join(DATA_DIR, 'records.json');

// 确保数据目录存在
if (!existsSync(DATA_DIR)) {
    mkdirSync(DATA_DIR, { recursive: true });
}
if (!existsSync(RECORDS_FILE)) {
    writeFileSync(RECORDS_FILE, JSON.stringify([], null, 2));
}

// 读取所有记录
function getAllRecords() {
    try {
        return JSON.parse(readFileSync(RECORDS_FILE, 'utf-8'));
    } catch (e) {
        return [];
    }
}

// 保存所有记录
function saveAllRecords(records) {
    writeFileSync(RECORDS_FILE, JSON.stringify(records, null, 2));
}

// 获取日期（YYYY-MM-DD）
function getToday() {
    return new Date().toISOString().split('T')[0];
}

// 中医问诊数据结构
const tcmData = {
    '一、寒热': ['喜凉恶热', '喜温恶凉', '经常畏冷', '容易感冒', '经常恶风', '脘腹腰背冷', '四肢凉', '下肢冷甚', '关节冷', '潮热', '手足心烧', '发热'],
    '二、汗出': ['自汗', '盗汗', '出虚汗或易出汗', '局部汗多', '但头汗出'],
    '三、疼痛部位': ['头痛', '咽喉痛', '胸痛', '胁痛', '脘腹痛', '腹痛', '关节痛', '肌肉痛', '腰痛', '背痛'],
    '四、疼痛性质': ['胀痛或窜痛', '绞痛', '刺痛', '固定痛', '游走痛', '灼痛', '冷痛', '隐痛', '痛喜按或按之舒', '痛拒按或压痛甚', '夜间痛甚', '得食痛缓', '进食痛甚', '阴雨天疼痛加重', '气行觉舒'],
    '五、头身不适': ['头晕', '眼花', '耳久鸣', '喷嚏', '鼻塞流清涕', '流浊涕', '喉痒', '咽部异物感', '鼻痒', '心悸', '胸闷', '胁胀', '脘痞胀', '胃脘嘈杂', '腹胀', '头重脚轻感', '倦怠乏力', '身体酸（困）重', '腰膝酸软', '皮肤瘙痒'],
    '六、睡眠': ['失眠', '多梦', '睡眠不实易醒', '嗜睡'],
    '七、情志': ['善悲易哭', '心烦', '胆怯易惊', '情志抑郁或忧虑、孤僻', '情绪易激动'],
    '八、声音': ['懒言', '声低', '声音洪亮', '声音重浊'],
    '九、咳痰喘': ['咳嗽', '干咳', '吐痰', '痰多质稠', '痰少质稠', '痰黏难咳', '痰多质稀', '痰少质稀', '泡沫痰多', '痰色白', '痰色黄', '腥臭痰', '痰中带血', '痰滑易咳', '气喘', '气短', '喘不能卧'],
    '十、饮食口味': ['口不渴', '口渴', '纳呆恶食', '进食无味', '饥不欲食', '食后痞胀', '多食易饥', '厌油腻', '口苦', '口臭', '口黏腻', '恶心', '呕吐', '干呕', '嗳气', '嗳气酸馊', '呃逆'],
    '十一、大便': ['经常腹泻', '经常便秘', '大便干结', '经常便溏', '大便先干后稀', '大便时结时溏', '大便有黏液', '大便腥腐臭气', '完谷不化', '排便无力', '排便不爽', '排便困难', '腹痛欲泻', '矢气多', '矢气甚臭'],
    '十二、小便': ['长期尿频', '排尿无力', '夜尿多', '尿短黄', '尿清长', '小便特多', '尿少', '排尿灼热', '排尿涩痛', '余尿不尽'],
    '十三、颈胸腹部体征': ['气息微弱', '三凹征阳性', '肺部干啰音', '肺部湿罗音', '桶状胸'],
    '十四、形体肌肤': ['身体素弱', '形体消瘦', '形体肥胖', '水肿', '肌肤甲错'],
    '十五、舌象': ['舌淡红', '舌淡胖', '舌淡紫', '舌赤', '舌绛', '舌红胖', '舌黯红', '舌尖红', '舌边红', '舌起芒刺', '舌有裂纹', '舌紫黯', '舌边有齿痕', '舌体干燥', '舌下络脉曲张', '舌苔薄白', '舌苔白', '舌苔腐垢', '舌苔黄', '舌苔灰黑', '舌苔黄白相间', '舌苔腻', '苔剥、少、无', '舌苔润滑', '舌苔干燥'],
    '十六、脉象': ['脉浮', '脉沉', '脉数', '脉洪', '脉细', '脉缓', '脉弦', '脉滑', '脉涩', '脉弱']
};

// 获取中医问诊Excel列
function getTCMExcelColumns(record) {
    const data = {};
    for (const [category, items] of Object.entries(tcmData)) {
        const categoryNum = category.split('、')[0];
        items.forEach(item => {
            const fieldName = `tcm_${categoryNum}_${item}`;
            const value = record[fieldName] || '';
            data[`TCM_${categoryNum}_${item}`] = value;
        });
    }
    return data;
}

const server = serve({
    port: process.env.PORT || 3030,
    fetch: async (req) => {
        const url = new URL(req.url);
        const path = url.pathname;

        // 忽略 .well-known 路径
        if (path.includes('.well-known')) {
            return new Response('Not Found', { status: 404 });
        }

        // API: 保存记录
        if (path === '/api/save' && req.method === 'POST') {
            try {
                const body = await req.json();
                const records = getAllRecords();
                
                // 如果有ID，更新现有记录；否则添加新记录
                const existingIndex = body.id ? records.findIndex(r => r.id === body.id && r.date === body.date) : -1;
                
                if (existingIndex >= 0) {
                    records[existingIndex] = { ...records[existingIndex], ...body, updatedAt: new Date().toISOString() };
                } else {
                    records.push({
                        ...body,
                        date: body.date || getToday(),
                        createdAt: new Date().toISOString()
                    });
                }
                
                saveAllRecords(records);
                return new Response(JSON.stringify({ success: true, message: '保存成功' }), {
                    headers: { 'Content-Type': 'application/json' }
                });
            } catch (e) {
                console.error('保存失败:', e);
                return new Response(JSON.stringify({ success: false, message: e.message }), {
                    status: 500,
                    headers: { 'Content-Type': 'application/json' }
                });
            }
        }

        // API: 获取所有记录
        if (path === '/api/records' && req.method === 'GET') {
            const records = getAllRecords();
            return new Response(JSON.stringify({ success: true, records }), {
                headers: { 'Content-Type': 'application/json' }
            });
        }

        // API: 按日期获取记录
        if (path === '/api/records' && req.method === 'POST') {
            try {
                const { date } = await req.json();
                const records = getAllRecords();
                const filtered = date ? records.filter(r => r.date === date) : records;
                return new Response(JSON.stringify({ success: true, records: filtered }), {
                    headers: { 'Content-Type': 'application/json' }
                });
            } catch (e) {
                return new Response(JSON.stringify({ success: false, message: e.message }), {
                    status: 500,
                    headers: { 'Content-Type': 'application/json' }
                });
            }
        }

        // API: 导出Excel
        if (path === '/api/export' && req.method === 'POST') {
            try {
                const { date } = await req.json();
                const records = getAllRecords();
                const filteredRecords = date ? records.filter(r => r.date === date) : records;

                if (filteredRecords.length === 0) {
                    return new Response(JSON.stringify({ success: false, message: '没有数据' }), {
                        status: 400,
                        headers: { 'Content-Type': 'application/json' }
                    });
                }

                // 构建Excel数据
                const excelData = filteredRecords.map(r => ({
                    '分组': r.group || '',
                    'ID号': r.id || '',
                    '姓名': r.name || '',
                    '性别': r.gender || '',
                    '民族': r.ethnicity || '',
                    '身高(M)': r.height || '',
                    '体重(KG)': r.weight || '',
                    'BMI': r.bmi || '',
                    '职业': r.occupation || '',
                    '婚姻': r.marriage || '',
                    '西医诊断': r.diagnosis_wm || '',
                    '中医诊断': r.diagnosis_tcm || '',
                    '既往病史用药史': r.medical_history || '',
                    '烟酒史': r.smoking_drinking || '',
                    '过敏史': r.allergy || '',
                    '手术史': r.surgery || '',
                    '环境接触史': r.environment || '',
                    '饮食习惯': r.diet || '',
                    '幽门螺旋杆菌': r.hp || '',
                    '病程(年)': r.duration || '',
                    '当前用药': r.medications || '',
                    'GERDQ评分': r.gerdq_total || '',
                    'mMRC分级': r.mmrc || '',
                    'CT报告': r.ct_report || '',
                    '肺纤维化位置': r.fibrosis_location || '',
                    '胃镜报告': r.gastroscopy || '',
                    '活检报告': r.biopsy || '',
                    '肺功能报告': r.lung_function || '',
                    '总蛋白': r.total_protein || '',
                    '白蛋白': r.albumin || '',
                    '前白蛋白': r.prealbumin || '',
                    '红细胞': r.rbc || '',
                    '血红蛋白': r.hemoglobin || '',
                    'PaO2': r.pao2 || '',
                    'PaCO2': r.paco2 || '',
                    'SaO2': r.sao2 || '',
                    '病原体检测': r.pathogen || '',
                    '体格检查': r.physical_exam || '',
                    'TCM总分': r.tcm_total || '',
                    ...getTCMExcelColumns(r),
                    '日期': r.date || '',
                    '创建时间': r.createdAt || ''
                }));

                return new Response(JSON.stringify({ success: true, data: excelData }), {
                    headers: { 'Content-Type': 'application/json' }
                });
            } catch (e) {
                console.error('导出失败:', e);
                return new Response(JSON.stringify({ success: false, message: e.message }), {
                    status: 500,
                    headers: { 'Content-Type': 'application/json' }
                });
            }
        }

        // API: 获取可用日期列表
        if (path === '/api/dates' && req.method === 'GET') {
            const records = getAllRecords();
            const dates = [...new Set(records.map(r => r.date))].sort().reverse();
            return new Response(JSON.stringify(dates), {
                headers: { 'Content-Type': 'application/json' }
            });
        }

        // 处理静态文件
        let filePath = path;
        if (filePath === '/') filePath = '/index.html';
        
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
console.log('📁 Data stored in:', DATA_DIR);
