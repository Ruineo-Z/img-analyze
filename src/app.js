import * as XLSX from '/node_modules/xlsx/xlsx.mjs';

class ImageAnalyzerApp {
    constructor() {
        this.dbName = 'ImgAnalyzeDB';
        this.storeName = 'records';
        this.apiKey = '';
        this.currentImageBase64 = '';
        
        this.initElements();
        this.bindEvents();
        
        // 初始化数据库后再加载数据
        this.initDB().then(() => {
            this.loadApiKey();
            this.loadTodayRecords();
        });
    }

    initElements() {
        this.uploadArea = document.getElementById('uploadArea');
        this.uploadSection = document.getElementById('uploadSection');
        this.fileInput = document.getElementById('fileInput');
        this.previewSection = document.getElementById('previewSection');
        this.imagePreview = document.getElementById('imagePreview');
        this.changeImageBtn = document.getElementById('changeImage');
        this.loadingSection = document.getElementById('loadingSection');
        this.formSection = document.getElementById('formSection');
        this.submitBtn = document.getElementById('submitBtn');
        this.successSection = document.getElementById('successSection');
        this.continueBtn = document.getElementById('continueBtn');
        this.recordsList = document.getElementById('recordsList');
        this.recordsSection = document.getElementById('recordsSection');
        this.totalCount = document.getElementById('totalCount');
        this.settingsBtn = document.getElementById('settingsBtn');
        this.downloadBtn = document.getElementById('downloadBtn');
        this.settingsModal = document.getElementById('settingsModal');
        this.closeSettings = document.getElementById('closeSettings');
        this.apiKeyInput = document.getElementById('apiKeyInput');
        this.saveApiKeyBtn = document.getElementById('saveApiKey');
        this.toast = document.getElementById('toast');
        
        this.fieldName = document.getElementById('field_name');
        this.fieldHeight = document.getElementById('field_height');
        this.fieldWeight = document.getElementById('field_weight');
    }

    bindEvents() {
        this.uploadArea.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        this.changeImageBtn.addEventListener('click', () => this.resetToUpload());
        this.submitBtn.addEventListener('click', () => this.saveRecord());
        this.continueBtn.addEventListener('click', () => this.resetToUpload());
        this.settingsBtn.addEventListener('click', () => this.showSettings());
        this.closeSettings.addEventListener('click', () => this.hideSettings());
        this.saveApiKeyBtn.addEventListener('click', () => this.saveApiKey());
        this.downloadBtn.addEventListener('click', () => this.downloadExcel());
    }

    async initDB() {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(this.dbName, 1);
            
            request.onerror = () => reject(request.error);
            request.onsuccess = () => {
                this.db = request.result;
                resolve();
            };
            
            request.onupgradeneeded = (event) => {
                const db = event.target.result;
                if (!db.objectStoreNames.contains(this.storeName)) {
                    const store = db.createObjectStore(this.storeName, { keyPath: 'id', autoIncrement: true });
                    store.createIndex('date', 'date', { unique: false });
                    store.createIndex('createdAt', 'createdAt', { unique: false });
                }
            };
        });
    }

    async loadApiKey() {
        const savedKey = localStorage.getItem('zhipu_api_key');
        if (savedKey) {
            this.apiKey = savedKey;
        } else {
            this.showSettings();
        }
    }

    async saveApiKey() {
        const key = this.apiKeyInput.value.trim();
        if (!key) {
            this.showToast('请输入 API Key', 'error');
            return;
        }
        
        this.apiKey = key;
        localStorage.setItem('zhipu_api_key', key);
        this.hideSettings();
        this.showToast('已保存', 'success');
    }

    showSettings() {
        this.apiKeyInput.value = this.apiKey;
        this.settingsModal.hidden = false;
    }

    hideSettings() {
        this.settingsModal.hidden = true;
    }

    handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;

        if (!file.type.startsWith('image/')) {
            this.showToast('请选择图片文件', 'error');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            this.currentImageBase64 = e.target.result;
            this.imagePreview.src = this.currentImageBase64;
            this.showPreview();
            this.analyzeImage();
        };
        reader.readAsDataURL(file);
    }

    showPreview() {
        this.uploadSection.hidden = true;
        this.previewSection.hidden = false;
        this.formSection.hidden = true;
        this.successSection.hidden = true;
    }

    resetToUpload() {
        this.uploadSection.hidden = false;
        this.previewSection.hidden = true;
        this.loadingSection.hidden = true;
        this.formSection.hidden = true;
        this.successSection.hidden = true;
        this.fileInput.value = '';
        this.currentImageBase64 = '';
        this.clearForm();
    }

    clearForm() {
        this.fieldName.value = '';
        this.fieldHeight.value = '';
        this.fieldWeight.value = '';
    }

    async analyzeImage() {
        this.previewSection.hidden = true;
        this.loadingSection.hidden = false;
        this.formSection.hidden = true;
        this.successSection.hidden = true;

        try {
            const data = await this.callGLMAPI();
            this.fillForm(data);
            
            this.loadingSection.hidden = true;
            this.formSection.hidden = false;
        } catch (error) {
            console.error('分析失败:', error);
            this.showToast(error.message || '分析失败，请重试', 'error');
            this.loadingSection.hidden = true;
            this.previewSection.hidden = false;
        }
    }

    async callGLMAPI() {
        if (!this.apiKey) {
            this.showSettings();
            throw new Error('请先配置 API Key');
        }

        const base64Data = this.currentImageBase64.split(',')[1];
        
        const response = await fetch('https://open.bigmodel.cn/api/paas/v4/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${this.apiKey}`
            },
            body: JSON.stringify({
                model: 'glm-4.5v',
                messages: [{
                    role: 'user',
                    content: [
                        {
                            type: 'image_url',
                            image_url: { url: base64Data }
                        },
                        {
                            type: 'text',
                            text: `请分析这张图片，提取以下信息：
1. 姓名（如有）
2. 身高（cm，只写数字）
3. 体重（kg，只写数字）

请用 JSON 格式输出：
{
  "name": "姓名或空字符串",
  "height": "身高数字或空字符串", 
  "weight": "体重数字或空字符串"
}

如果某个信息不存在或无法识别，设置为空字符串。只输出 JSON，不要其他文字。`
                        }
                    ]
                }],
                thinking: { type: 'enabled' }
            })
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(error.error?.message || 'API 调用失败');
        }

        const data = await response.json();
        return this.parseResponse(data.choices[0].message.content);
    }

    parseResponse(response) {
        let jsonStr = response.trim();
        
        const jsonMatch = jsonStr.match(/```json\s*([\s\S]*?)\s*```/);
        if (jsonMatch) {
            jsonStr = jsonMatch[1];
        }

        try {
            return JSON.parse(jsonStr);
        } catch (e) {
            console.warn('JSON 解析失败:', jsonStr);
            return { name: '', height: '', weight: '' };
        }
    }

    fillForm(data) {
        this.fieldName.value = data.name || '';
        this.fieldHeight.value = data.height || '';
        this.fieldWeight.value = data.weight || '';
    }

    async saveRecord() {
        const record = {
            name: this.fieldName.value.trim(),
            height: this.fieldHeight.value.trim(),
            weight: this.fieldWeight.value.trim(),
            date: new Date().toISOString().split('T')[0],
            createdAt: new Date().toISOString(),
            image: this.currentImageBase64.substring(0, 100) + '...' // 只存缩略
        };

        try {
            await this.addRecord(record);
            this.formSection.hidden = true;
            this.successSection.hidden = false;
            document.getElementById('successMsg').textContent = '今日已保存 ' + record.name + ' 的信息';
            this.loadTodayRecords();
            this.showToast('保存成功', 'success');
        } catch (error) {
            console.error('保存失败:', error);
            this.showToast('保存失败，请重试', 'error');
        }
    }

    async addRecord(record) {
        return new Promise((resolve, reject) => {
            const transaction = this.db.transaction([this.storeName], 'readwrite');
            const store = transaction.objectStore(this.storeName);
            const request = store.add(record);
            
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
    }

    async loadTodayRecords() {
        const today = new Date().toISOString().split('T')[0];
        
        const allRecords = await this.getAllRecords();
        const todayRecords = allRecords.filter(r => r.date === today);
        const sortedRecords = todayRecords.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

        this.totalCount.textContent = allRecords.length;
        this.renderRecords(sortedRecords);
    }

    async getAllRecords() {
        return new Promise((resolve, reject) => {
            const transaction = this.db.transaction([this.storeName], 'readonly');
            const store = transaction.objectStore(this.storeName);
            const request = store.getAll();
            
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
    }

    renderRecords(records) {
        if (records.length === 0) {
            this.recordsList.innerHTML = '<p class="empty-tip">暂无记录</p>';
            return;
        }

        this.recordsList.innerHTML = records.map(record => `
            <div class="record-item">
                <div class="record-info">
                    <div class="record-name">${record.name || '未填写姓名'}</div>
                    <div class="record-detail">身高: ${record.height || '-'}cm | 体重: ${record.weight || '-'}kg</div>
                </div>
                <div class="record-time">${this.formatTime(record.createdAt)}</div>
            </div>
        `).join('');
    }

    formatTime(isoString) {
        const date = new Date(isoString);
        return date.getHours().toString().padStart(2, '0') + ':' + 
               date.getMinutes().toString().padStart(2, '0');
    }

    async downloadExcel() {
        try {
            const allRecords = await this.getAllRecords();
            
            if (allRecords.length === 0) {
                this.showToast('暂无数据', 'error');
                return;
            }

            const data = allRecords.map(r => ({
                '姓名': r.name || '',
                '身高(cm)': r.height || '',
                '体重(kg)': r.weight || '',
                '日期': r.date,
                '记录时间': r.createdAt
            }));

            const ws = XLSX.utils.json_to_sheet(data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Records');
            
            const date = new Date().toISOString().split('T')[0];
            XLSX.writeFile(wb, `身高体重记录_${date}.xlsx`);
            
            this.showToast('下载成功', 'success');
        } catch (error) {
            console.error('下载失败:', error);
            this.showToast('下载失败，请重试', 'error');
        }
    }

    showToast(message, type = '') {
        this.toast.textContent = message;
        this.toast.className = 'toast show ' + type;
        
        setTimeout(() => {
            this.toast.className = 'toast';
        }, 3000);
    }
}

document.addEventListener('DOMContentLoaded', () => {
    window.app = new ImageAnalyzerApp();
});
