import * as XLSX from '/node_modules/xlsx/xlsx.mjs';

class ImageAnalyzerApp {
    constructor() {
        this.dbName = 'ImgAnalyzeDB';
        this.storeName = 'records';
        this.apiKey = '';
        this.currentImageBase64 = '';
        this.uploadedImages = [];  // 存储所有已上传图片的base64数据
        this.MAX_IMAGES = 10;  // 最大支持10张图片
        
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', () => this.init());
        } else {
            this.init();
        }
    }
    
    init() {
        console.log('=== init called ===');
        this.initElements();
        this.bindEvents();
        
        this.initDB().then(() => {
            this.loadApiKey();
            this.loadTodayRecords();
            this.initTCMQuestionnaire();
        });
    }

    // 中医问诊数据结构
    getTCMData() {
        return {
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
    }

    initTCMQuestionnaire() {
        const container = document.getElementById('tcmQuestionnaire');
        if (!container) return;
        
        const tcmData = this.getTCMData();
        this.tcmFields = {}; // 存储所有TCM字段引用
        
        for (const [category, items] of Object.entries(tcmData)) {
            const categoryDiv = document.createElement('details');
            categoryDiv.className = 'tcm-category';
            
            const summaryDiv = document.createElement('summary');
            summaryDiv.innerHTML = `
                <span class="tcm-category-summary">
                    <span>${category}</span>
                    <span class="tcm-category-score" id="score_${category.split('、')[0]}">0分</span>
                </span>
            `;
            
            const itemsDiv = document.createElement('div');
            itemsDiv.className = 'tcm-items';
            
            items.forEach(item => {
                const fieldName = `tcm_${category.split('、')[0]}_${item}`;
                const itemDiv = document.createElement('div');
                itemDiv.className = 'tcm-item';
                itemDiv.innerHTML = `
                    <label for="field_${fieldName}">${item}</label>
                    <select id="field_${fieldName}" name="${fieldName}" class="tcm-select">
                        <option value="0">0分 - 无</option>
                        <option value="1">1分 - 轻度</option>
                        <option value="2">2分 - 中度</option>
                        <option value="3">3分 - 重度</option>
                    </select>
                `;
                itemsDiv.appendChild(itemDiv);
                this.tcmFields[fieldName] = itemDiv.querySelector('select');
            });
            
            categoryDiv.appendChild(summaryDiv);
            categoryDiv.appendChild(itemsDiv);
            container.appendChild(categoryDiv);
            
            // 为该分类的所有下拉框添加事件监听
            itemsDiv.querySelectorAll('select').forEach(select => {
                select.addEventListener('change', () => this.calculateTCMScore(category));
            });
        }
    }

    calculateTCMScore(category) {
        const categoryNum = category.split('、')[0];
        const items = this.getTCMData()[category];
        let total = 0;
        
        items.forEach(item => {
            const fieldName = `tcm_${categoryNum}_${item}`;
            const select = document.getElementById(`field_${fieldName}`);
            if (select) {
                total += parseInt(select.value) || 0;
            }
        });
        
        // 更新分类分数显示
        const scoreEl = document.getElementById(`score_${categoryNum}`);
        if (scoreEl) {
            scoreEl.textContent = `${total}分`;
        }
        
        // 更新总分
        this.calculateTCMTotal();
    }

    calculateTCMTotal() {
        let total = 0;
        const tcmData = this.getTCMData();
        
        for (const [category, items] of Object.entries(tcmData)) {
            const categoryNum = category.split('、')[0];
            items.forEach(item => {
                const fieldName = `tcm_${categoryNum}_${item}`;
                const select = document.getElementById(`field_${fieldName}`);
                if (select) {
                    total += parseInt(select.value) || 0;
                }
            });
        }
        
        this.fieldTcmTotal.value = total;
    }

    collectTCMData() {
        const data = {};
        const tcmData = this.getTCMData();
        
        for (const [category, items] of Object.entries(tcmData)) {
            const categoryNum = category.split('、')[0];
            items.forEach(item => {
                const fieldName = `tcm_${categoryNum}_${item}`;
                const select = document.getElementById(`field_${fieldName}`);
                if (select) {
                    data[fieldName] = select.value;
                }
            });
        }
        
        return data;
    }

    initElements() {
        console.log('=== initElements called ===');
        console.log('document.readyState:', document.readyState);
        console.log('formSection exists:', !!document.getElementById('formSection'));
        this.uploadSection = document.getElementById('uploadSection');
        this.uploadArea = document.getElementById('uploadArea');
        this.fileInput = document.getElementById('fileInput');
        this.previewSection = document.getElementById('previewSection');
        this.previewGrid = document.getElementById('previewGrid');
        this.imageCount = document.getElementById('imageCount');
        this.clearAllImagesBtn = document.getElementById('clearAllImages');
        this.addMoreImagesBtn = document.getElementById('addMoreImages');
        this.analyzeImagesBtn = document.getElementById('analyzeImagesBtn');
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
        // 日期选择弹窗
        this.dateSelectModal = document.getElementById('dateSelectModal');
        this.closeDateSelect = document.getElementById('closeDateSelect');
        this.dateSelect = document.getElementById('dateSelect');
        this.confirmExport = document.getElementById('confirmExport');
        this.toast = document.getElementById('toast');
        
        // 第一部分字段
        this.fieldGroup = document.getElementById('field_group');
        this.fieldId = document.getElementById('field_id');
        this.fieldName = document.getElementById('field_name');
        this.fieldAge = document.getElementById('field_age');
        this.fieldGender = document.getElementById('field_gender');
        this.fieldEthnicity = document.getElementById('field_ethnicity');
        this.fieldHeight = document.getElementById('field_height');
        this.fieldWeight = document.getElementById('field_weight');
        this.fieldBmi = document.getElementById('field_bmi');
        this.fieldOccupation = document.getElementById('field_occupation');
        this.fieldMarriage = document.getElementById('field_marriage');
        this.fieldDiagnosisWm = document.getElementById('field_diagnosis_wm');
        this.fieldDiagnosisTcm = document.getElementById('field_diagnosis_tcm');
        this.fieldMedicalHistory = document.getElementById('field_medical_history');
        this.fieldSmokingDrinking = document.getElementById('field_smoking_drinking');
        this.fieldAllergy = document.getElementById('field_allergy');
        this.fieldSurgery = document.getElementById('field_surgery');
        // 环境接触史（多选）
        this.checkboxEnvironments = document.querySelectorAll('input[name="environment"]');
        // 饮食习惯（多选）
        this.checkboxDiets = document.querySelectorAll('input[name="diet"]');
        // 幽门螺旋杆菌
        this.fieldHp = document.getElementById('field_hp');
        this.fieldDuration = document.getElementById('field_duration');
        this.fieldMedications = document.getElementById('field_medications');
        
        // 第二部分：GERDQ评分字段（6问题 × 4天数区间 = 24字段） - 现在使用复选框
        // 不再需要预先存储引用，使用querySelector动态查找
        
        // 第三部分：mMRC字段
        this.fieldMmrc = document.getElementById('field_mmrc');
        this.fieldGerdqTotal = document.getElementById('field_gerdq_total');
        
        // 第四部分：检查报告字段
        this.fieldCtReport = document.getElementById('field_ct_report');
        this.fieldFibrosisLocation = document.getElementById('field_fibrosis_location');
        this.fieldGastroscopy = document.getElementById('field_gastroscopy');
        this.fieldBiopsy = document.getElementById('field_biopsy');
        this.fieldLungFunction = document.getElementById('field_lung_function');
        
        // 营养指标字段
        this.fieldTotalProtein = document.getElementById('field_total_protein');
        this.fieldAlbumin = document.getElementById('field_albumin');
        this.fieldPrealbumin = document.getElementById('field_prealbumin');
        this.fieldRbc = document.getElementById('field_rbc');
        this.fieldHemoglobin = document.getElementById('field_hemoglobin');
        
        // 血气分析字段
        this.fieldPao2 = document.getElementById('field_pao2');
        this.fieldPaco2 = document.getElementById('field_paco2');
        this.fieldSao2 = document.getElementById('field_sao2');
        
        // 氧合评估字段
        this.fieldOxygenFlow = document.getElementById('field_oxygen_flow');
        this.fieldOxygenConcentration = document.getElementById('field_oxygen_concentration');
        this.fieldPao2 = document.getElementById('field_pao2');
        this.fieldOxygenIndex = document.getElementById('field_oxygen_index');
        
        // 病原体检测字段（多选）
        this.checkboxPathogens = document.querySelectorAll('input[name="pathogen"]');
        
        // 体格检查字段
        this.fieldPhysicalExam = document.getElementById('field_physical_exam');
        
        // 第五部分：中医问诊字段
        this.fieldTcmTotal = document.getElementById('field_tcm_total');
        this.tcmSelects = [];
    }

    bindEvents() {
        this.uploadArea.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        this.clearAllImagesBtn.addEventListener('click', (e) => {
            e.preventDefault();
            this.clearAllImages();
        });
        this.addMoreImagesBtn.addEventListener('click', (e) => {
            e.preventDefault();
            this.fileInput.click();
        });
        this.analyzeImagesBtn.addEventListener('click', () => this.analyzeImages());
        this.submitBtn.addEventListener('click', () => this.saveRecord());
        this.continueBtn.addEventListener('click', () => this.resetToUpload());
        this.settingsBtn.addEventListener('click', () => this.showSettings());
        this.closeSettings.addEventListener('click', () => this.hideSettings());
        this.saveApiKeyBtn.addEventListener('click', () => this.saveApiKey());
        this.downloadBtn.addEventListener('click', () => this.showDateSelectModal());
        
        this.fieldHeight.addEventListener('input', () => this.calculateBMI());
        this.fieldWeight.addEventListener('input', () => this.calculateBMI());
        
        for (let q = 1; q <= 6; q++) {
            for (let d = 0; d <= 3; d++) {
                const checkbox = document.querySelector(`input[name="gerdq_${q}_${d}"]`);
                if (checkbox) {
                    checkbox.addEventListener('change', () => this.calculateGerdq());
                }
            }
        }
        
        this.fieldOxygenFlow.addEventListener('input', () => this.calculateOxygenConcentration());
        this.fieldPao2.addEventListener('input', () => this.calculateOxygenIndex());
        
        // 环境接触史互斥逻辑（选择"无"时取消其他，选择其他时取消"无"）
        this.checkboxEnvironments.forEach(cb => {
            cb.addEventListener('change', () => this.handleMutexCheckboxesWithClear(this.checkboxEnvironments, '无'));
        });
        
        // 饮食习惯互斥逻辑
        this.checkboxDiets.forEach(cb => {
            cb.addEventListener('change', () => this.handleMutexCheckboxesWithClear(this.checkboxDiets, '无'));
        });
        
        // 病原体检测互斥逻辑
        this.checkboxPathogens.forEach(cb => {
            cb.addEventListener('change', () => this.handleMutexCheckboxesWithClear(this.checkboxPathogens, '无'));
        });
        
        // 日期选择弹窗事件
        this.downloadBtn.addEventListener('click', () => this.showDateSelectModal());
        this.closeDateSelect.addEventListener('click', () => this.hideDateSelectModal());
        this.confirmExport.addEventListener('click', () => this.exportBySelectedDate());
    }
    
    showDateSelectModal() {
        this.showDateSelectDialog();
    }
    
    showDateSelectDialog() {
        this.getAllRecords().then(allRecords => {
            if (!allRecords || allRecords.length === 0) {
                this.showToast('暂无数据', 'error');
                return;
            }

            // 提取所有唯一日期
            const dates = [...new Set(allRecords.map(r => r.date))].sort().reverse();
            
            // 填充下拉框
            this.dateSelect.innerHTML = '<option value="">请选择日期</option>';
            dates.forEach(date => {
                const count = allRecords.filter(r => r.date === date).length;
                const option = document.createElement('option');
                option.value = date;
                option.textContent = `${date} (${count}条)`;
                this.dateSelect.appendChild(option);
            });
            
            this.dateSelectModal.hidden = false;
        });
    }
    
    hideDateSelectModal() {
        this.dateSelectModal.hidden = true;
    }
    
    async exportBySelectedDate() {
        const selectedDate = this.dateSelect.value;
        if (!selectedDate) {
            this.showToast('请选择日期', 'error');
            return;
        }
        
        this.hideDateSelectModal();
        
        this.showToast(`正在导出 ${selectedDate} 的数据...`, 'success');
        
        try {
            const response = await fetch('/api/export', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ date: selectedDate })
            });
            
            const result = await response.json();
            
            if (!result.success) {
                throw new Error(result.message || '导出失败');
            }
            
            const ws = XLSX.utils.json_to_sheet(result.data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Records');
            XLSX.writeFile(wb, `医疗信息录入_${selectedDate}.xlsx`);
            
            this.showToast('下载成功', 'success');
        } catch (error) {
            console.error('导出失败:', error);
            this.showToast('导出失败，请重试', 'error');
        }
    }

    calculateBMI() {
        const height = parseFloat(this.fieldHeight.value);
        const weight = parseFloat(this.fieldWeight.value);
        
        if (height > 0 && weight > 0) {
            const bmi = weight / (height * height);
            this.fieldBmi.value = bmi.toFixed(2);
        } else {
            this.fieldBmi.value = '';
        }
    }
    
    calculateGerdq() {
        let total = 0;
        for (let q = 1; q <= 6; q++) {
            for (let d = 0; d <= 3; d++) {
                const checkbox = document.querySelector(`input[name="gerdq_${q}_${d}"]`);
                if (checkbox && checkbox.checked) {
                    total += parseInt(checkbox.value) || 0;
                }
            }
        }
        this.fieldGerdqTotal.value = total;
    }
    
    handleGerdqCheckboxChange(row) {
        for (let d = 0; d <= 3; d++) {
            const checkbox = document.querySelector(`input[name="gerdq_${row}_${d}"]`);
            if (checkbox) {
                checkbox.addEventListener('change', (e) => {
                    if (e.target.checked) {
                        for (let od = 0; od <= 3; od++) {
                            if (od !== d) {
                                const otherCheckbox = document.querySelector(`input[name="gerdq_${row}_${od}"]`);
                                if (otherCheckbox) {
                                    otherCheckbox.checked = false;
                                }
                            }
                        }
                    }
                    this.calculateGerdq();
                });
            }
        }
    }
    
    calculateOxygenConcentration() {
        const oxygenFlow = parseFloat(this.fieldOxygenFlow.value);
        
        if (oxygenFlow > 0) {
            const concentration = 21 + 4 * oxygenFlow;
            this.fieldOxygenConcentration.value = concentration.toFixed(0);
        } else {
            this.fieldOxygenConcentration.value = '';
        }
        
        this.calculateOxygenIndex();
    }
    
    calculateOxygenIndex() {
        const pao2 = parseFloat(this.fieldPao2.value);
        const fio2 = parseFloat(this.fieldOxygenConcentration.value);
        
        if (pao2 > 0 && fio2 > 0) {
            const oxygenIndex = pao2 / (fio2 / 100);
            this.fieldOxygenIndex.value = oxygenIndex.toFixed(1);
        } else {
            this.fieldOxygenIndex.value = '';
        }
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
        const files = Array.from(event.target.files);
        if (!files.length) return;
        
        // 验证文件类型
        const invalidFiles = files.filter(file => !file.type.startsWith('image/'));
        if (invalidFiles.length > 0) {
            this.showToast('请选择图片文件', 'error');
            return;
        }
        
        // 检查是否超过最大数量
        if (this.uploadedImages.length + files.length > this.MAX_IMAGES) {
            this.showToast(`最多只能上传${this.MAX_IMAGES}张图片`, 'error');
            return;
        }
        
        // 转换所有文件为base64
        const readers = files.map(file => {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => resolve({ name: file.name, base64: e.target.result });
                reader.onerror = reject;
                reader.readAsDataURL(file);
            });
        });
        
        Promise.all(readers).then(results => {
            this.uploadedImages.push(...results);
            this.updatePreviewUI();
            this.showPreview();
        }).catch(error => {
            this.showToast('图片读取失败', 'error');
        });
    }
    
    updatePreviewUI() {
        this.previewGrid.innerHTML = '';
        this.imageCount.textContent = this.uploadedImages.length;
        
        this.uploadedImages.forEach((img, index) => {
            const item = document.createElement('div');
            item.className = 'preview-item';
            item.innerHTML = `
                <img src="${img.base64}" alt="${img.name}">
                <button class="remove-btn" data-index="${index}">×</button>
            `;
            item.querySelector('.remove-btn').addEventListener('click', (e) => {
                e.stopPropagation();
                this.removeImage(index);
            });
            this.previewGrid.appendChild(item);
        });
    }
    
    removeImage(index) {
        this.uploadedImages.splice(index, 1);
        this.updatePreviewUI();
    }
    
    clearAllImages() {
        this.uploadedImages = [];
        this.updatePreviewUI();
        this.showToast('已清空所有图片', 'success');
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
        this.uploadedImages = [];
        this.clearForm();
    }
    
    clearForm() {
        this.uploadedImages = [];  // 清空已上传图片
        this.fieldGroup.value = '';
        this.fieldId.value = '';
        this.fieldName.value = '';
        this.fieldAge.value = '';
        this.fieldGender.value = '';
        this.fieldEthnicity.value = '';
        this.fieldHeight.value = '';
        this.fieldWeight.value = '';
        this.fieldBmi.value = '';
        this.fieldOccupation.value = '';
        this.fieldMarriage.value = '';
        this.fieldDiagnosisWm.value = '';
        this.fieldDiagnosisTcm.value = '';
        this.fieldMedicalHistory.value = '';
        this.fieldSmokingDrinking.value = '';
        this.fieldAllergy.value = '';
        this.fieldSurgery.value = '';
        // 环境接触史（多选）
        this.checkboxEnvironments.forEach(cb => {
            cb.checked = false;
        });
        // 饮食习惯（多选）
        this.checkboxDiets.forEach(cb => {
            cb.checked = false;
        });
        // 幽门螺旋杆菌
        this.fieldHp.value = '';
        this.fieldDuration.value = '';
        this.fieldMedications.value = '';
        
        // GERDQ fields (6 questions × 4 day ranges = 24 fields) - 复选框模式
        for (let q = 1; q <= 6; q++) {
            for (let d = 0; d <= 3; d++) {
                const checkbox = document.querySelector(`input[name="gerdq_${q}_${d}"]`);
                if (checkbox) checkbox.checked = false;
            }
        }
        this.fieldGerdqTotal.value = '';
        
        // mMRC field
        this.fieldMmrc.value = '';
        
        // 第四部分：检查报告字段
        this.fieldCtReport.value = '';
        this.fieldFibrosisLocation.value = '';
        this.fieldGastroscopy.value = '';
        this.fieldBiopsy.value = '';
        this.fieldLungFunction.value = '';
        
        // 营养指标字段
        this.fieldTotalProtein.value = '';
        this.fieldAlbumin.value = '';
        this.fieldPrealbumin.value = '';
        this.fieldRbc.value = '';
        this.fieldHemoglobin.value = '';
        
        // 血气分析字段
        if (this.fieldBloodGas) {
            this.fieldBloodGas.value = '';
        }
        
        // 氧合评估字段
        this.fieldOxygenFlow.value = '';
        this.fieldOxygenConcentration.value = '';
        this.fieldPao2.value = '';
        this.fieldOxygenIndex.value = '';
        
        // 病原体检测字段（多选）
        if (this.checkboxPathogens) {
            this.checkboxPathogens.forEach(cb => cb.checked = false);
        }
        
        // 体格检查字段
        this.fieldPhysicalExam.value = '';
        
        // 第五部分：中医问诊字段
        this.fieldTcmTotal.value = '';
        
        // 重置所有TCM下拉框
        if (this.tcmFields) {
            Object.values(this.tcmFields).forEach(select => {
                select.value = '0';
            });
        }
        
        // 重置分类分数显示
        const tcmData = this.getTCMData();
        for (const category of Object.keys(tcmData)) {
            const categoryNum = category.split('、')[0];
            const scoreEl = document.getElementById(`score_${categoryNum}`);
            if (scoreEl) {
                scoreEl.textContent = '0分';
            }
        }
    }
    
    async saveRecord() {
        this.calculateBMI();
        this.calculateGerdq();
        this.calculateOxygenConcentration();
        this.calculateOxygenIndex();
        this.calculateTCMTotal();
        
        const record = {
            group: this.fieldGroup.value,
            id: this.fieldId.value.trim(),
            name: this.fieldName.value.trim(),
            age: this.fieldAge.value.trim(),
            gender: this.fieldGender.value,
            ethnicity: this.fieldEthnicity.value.trim(),
            height: this.fieldHeight.value.trim(),
            weight: this.fieldWeight.value.trim(),
            bmi: this.fieldBmi.value,
            occupation: this.fieldOccupation.value.trim(),
            marriage: this.fieldMarriage.value,
            diagnosis_wm: this.fieldDiagnosisWm.value.trim(),
            diagnosis_tcm: this.fieldDiagnosisTcm.value.trim(),
            medical_history: this.fieldMedicalHistory.value.trim(),
            smoking_drinking: this.fieldSmokingDrinking.value.trim(),
            allergy: this.fieldAllergy.value.trim(),
            surgery: this.fieldSurgery.value.trim(),
            environment: Array.from(this.checkboxEnvironments)
                .filter(cb => cb.checked)
                .map(cb => cb.value)
                .join(';'),
            diet: Array.from(this.checkboxDiets)
                .filter(cb => cb.checked)
                .map(cb => cb.value)
                .join(';'),
            hp: this.fieldHp.value,
            duration: this.fieldDuration.value.trim(),
            medications: this.fieldMedications.value.trim(),
            
            // GERDQ评分 (6 questions × 4 day ranges = 24 fields) - 复选框模式
            ...(() => {
                const gerdqData = {};
                const dayRanges = ['0天', '1天', '2-3天', '4-7天'];
                const gerdqQuestions = [
                    '胸骨后灼烧感', '胃内容物上返', '上腹部中央疼痛',
                    '感到恶心', '夜间睡眠困难', '额外服药'
                ];
                for (let q = 1; q <= 6; q++) {
                    for (let d = 0; d <= 3; d++) {
                        const key = `gerdq_${q}_${d}`;
                        const checkbox = document.querySelector(`input[name="${key}"]`);
                        const value = (checkbox && checkbox.checked) ? checkbox.value : '';
                        gerdqData[key] = value;
                        gerdqData[`GERDQ_${gerdqQuestions[q-1]}_${dayRanges[d]}`] = value;
                    }
                }
                gerdqData.gerdq_total = this.fieldGerdqTotal.value;
                return gerdqData;
            })(),
            
            // mMRC评分
            mmrc: this.fieldMmrc.value,
            
            // 第四部分：检查报告
            ct_report: this.fieldCtReport.value.trim(),
            fibrosis_location: this.fieldFibrosisLocation.value.trim(),
            gastroscopy: this.fieldGastroscopy.value.trim(),
            biopsy: this.fieldBiopsy.value.trim(),
            lung_function: this.fieldLungFunction.value.trim(),
            
            // 营养指标
            total_protein: this.fieldTotalProtein.value.trim(),
            albumin: this.fieldAlbumin.value.trim(),
            prealbumin: this.fieldPrealbumin.value.trim(),
            rbc: this.fieldRbc.value.trim(),
            hemoglobin: this.fieldHemoglobin.value.trim(),
            
            // 血气分析
            pao2: this.fieldPao2.value.trim(),
            paco2: this.fieldPaco2.value.trim(),
            sao2: this.fieldSao2.value.trim(),
            oxygen_index: this.fieldOxygenIndex.value.trim(),
            
            // 病原体检测（多选）
            pathogen: Array.from(this.checkboxPathogens)
                .filter(cb => cb.checked)
                .map(cb => cb.value)
                .join(';'),
            
            // 体格检查
            physical_exam: this.fieldPhysicalExam.value.trim(),
            
            // 第五部分：中医问诊
            tcm_total: this.fieldTcmTotal.value,
            ...this.collectTCMData(),
            
            // 日期信息
            date: new Date().toISOString().split('T')[0],
            createdAt: new Date().toISOString()
        };
        
        try {
            // 调用后端API保存记录
            const response = await fetch('/api/save', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(record)
            });
            const result = await response.json();
            
            if (result.success) {
                this.formSection.hidden = true;
                this.successSection.hidden = false;
                document.getElementById('successMsg').textContent = '已保存: ' + (record.name || record.id || '新记录');
                this.loadTodayRecords();
                this.showToast('保存成功', 'success');
            } else {
                throw new Error(result.message);
            }
        } catch (error) {
            console.error('保存失败:', error);
            this.showToast('保存失败，请重试', 'error');
        }
    }
    
    async loadTodayRecords() {
        const today = new Date().toISOString().split('T')[0];
        
        const allRecords = await this.getAllRecords();
        const todayRecords = allRecords.filter(r => r.date === today);
        const sortedRecords = todayRecords.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
        
        this.totalCount.textContent = allRecords.length;
        this.renderRecords(sortedRecords);
    }
    
    async getAllRecords(dateFilter = null) {
        try {
            const url = dateFilter 
                ? `/api/records?date=${encodeURIComponent(dateFilter)}`
                : '/api/records';
            const response = await fetch(url);
            const result = await response.json();
            
            if (result.success) {
                return result.records;
            } else {
                console.error('获取记录失败:', result.message);
                return [];
            }
        } catch (error) {
            console.error('获取记录失败:', error);
            return [];
        }
    }
    
    renderRecords(records) {
        if (records.length === 0) {
            this.recordsList.innerHTML = '<p class="empty-tip">暂无记录</p>';
            return;
        }
        
        this.recordsList.innerHTML = records.map(record => `
            <div class="record-item">
                <div class="record-info">
                    <div class="record-name">${record.name || record.id || '未识别'}</div>
                    <div class="record-detail">${record.group || ''} | ${record.gender || ''} | BMI: ${record.bmi || '-'} | GERDQ: ${record.gerdq_total || '-'} | mMRC: ${record.mmrc || '-'}</div>
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
    
    getTCMExcelColumns(record) {
        const data = {};
        const tcmData = this.getTCMData();
        
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
    
    showToast(message, type = '') {
        this.toast.textContent = message;
        this.toast.className = 'toast show ' + type;
        
        setTimeout(() => {
            this.toast.className = 'toast';
        }, 3000);
    }
    
    handleMutexCheckboxes(checkboxes, exclusiveValue) {
        const exclusiveCb = Array.from(checkboxes).find(cb => cb.value === exclusiveValue);
        const otherCbs = Array.from(checkboxes).filter(cb => cb.value !== exclusiveValue);
        
        if (exclusiveCb.checked) {
            otherCbs.forEach(cb => {
                cb.checked = false;
                cb.disabled = true;
            });
        } else {
            otherCbs.forEach(cb => {
                cb.disabled = false;
            });
        }
    }
    
    // 互斥逻辑：选择"无"时取消其他，选择其他时取消"无"
    handleMutexCheckboxesWithClear(checkboxes, exclusiveValue) {
        const exclusiveCb = Array.from(checkboxes).find(cb => cb.value === exclusiveValue);
        const otherCbs = Array.from(checkboxes).filter(cb => cb.value !== exclusiveValue);
        
        if (exclusiveCb.checked) {
            // 选中"无"时，取消其他所有选项
            otherCbs.forEach(cb => {
                cb.checked = false;
            });
        } else {
            // 选中其他选项时，取消"无"
            exclusiveCb.checked = false;
        }
    }
    
    async analyzeImage() {
        await this.analyzeImages();
    }
    
    async callGLMAPI() {
        if (!this.apiKey) {
            this.showSettings();
            throw new Error('请先配置 API Key');
        }
        
        if (this.uploadedImages.length === 0) {
            throw new Error('请先上传图片');
        }
        
        const imageContents = this.uploadedImages.map(img => ({
            type: 'image_url',
            image_url: { url: img.base64.split(',')[1] }
        }));
        
        const response = await fetch('https://open.bigmodel.cn/api/paas/v4/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${this.apiKey}`
            },
            body: JSON.stringify({
                model: 'glm-4.6v',
                messages: [{
                    role: 'user',
                    content: [
                        ...imageContents,
                        {
                            type: 'text',
                            text: `请从以上${this.uploadedImages.length}张医疗图片中提取患者信息，按以下JSON格式输出。务必确保输出的是有效的JSON格式。

【重要规则】
1. 务必提取完整信息，不要遗漏细节
2. 如果某字段信息不明确或未提及，填写"未提及"，不要留空
3. 所有文本字段必须是字符串类型
4. 严格按照JSON格式输出，不要有额外解释
5. 【关键】JSON字符串值内部绝对不要使用引号，所有引号必须是JSON结构的一部分
6. 如果诊断名称本身包含引号，请改用其他描述方式或去除引号
7. 请仔细识别所有图片，跨图片关联提取信息，确保信息完整准确

【注意：血气分析和肺功能是两种不同的报告】
- 血气分析报告：包含体温、pH值、氧分压、二氧化碳分压、血糖、乳酸等指标
- 肺功能报告：包含FVC、FEV1、PEF等呼吸功能指标
- 请注意区分，不要混淆

【字段提取规则】
- age: 年龄，只填数字，如 "45"
- ct_report: CT报告的完整描述
- gastroscopy: 胃镜报告的完整描述
- biopsy: 活检报告的完整描述
- lung_function: 肺功能报告的完整描述（如FVC、FEV1等），不要包含血气分析内容
- pao2: 血气分析中的氧分压数值，只填数字，如 "112.1"
- paco2: 血气分析中的二氧化碳分压数值，只填数字，如 "84.5"
- sao2: 血气分析中的血氧饱和度数值，只填数字，如 "97.9"
- total_protein: 总蛋白数值，只填数字
- albumin: 白蛋白数值，只填数字
- prealbumin: 前白蛋白数值，只填数字
- rbc: 红细胞计数数值，只填数字
- hemoglobin: 血红蛋白数值，只填数字

【输出格式】只输出JSON，不要有任何其他内容：
{"id": "", "name": "", "age": "", "gender": "", "ethnicity": "", "occupation": "", "marriage": "", "diagnosis_wm": "", "diagnosis_tcm": "", "medical_history": "", "smoking_drinking": "", "allergy": "", "surgery": "", "medications": "", "ct_report": "", "fibrosis_location": "", "gastroscopy": "", "biopsy": "", "lung_function": "", "total_protein": "", "albumin": "", "prealbumin": "", "rbc": "", "hemoglobin": "", "pao2": "", "paco2": "", "sao2": "", "physical_exam": ""}

请仔细识别所有图片，确保完整准确。`
                        }
                    ]
                }],
                temperature: 0.1
            })
        });
        
        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(error.error?.message || 'API 调用失败');
        }
        
        const data = await response.json();
        return this.parseResponse(data.choices[0].message.content);
    }
    
    async analyzeImages() {
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
    
    parseResponse(response) {
        if (!response) return {};
        
        // 打印原始响应，方便调试
        console.log('========== 模型原始响应 ==========');
        console.log(response);
        console.log('========== 响应结束 ==========');
        
        let jsonStr = response.trim();
        
        // 移除 ```json 和 ``` 标记
        jsonStr = jsonStr.replace(/^```json\s*/i, '').replace(/\s*```$/i, '');
        
        // 尝试直接解析
        try {
            return JSON.parse(jsonStr);
        } catch (e) {
            // 如果直接解析失败，尝试修复常见问题
            console.warn('JSON直接解析失败，尝试修复...', e.message);
        }
        
        // 尝试修复并重新解析
        const fixedJson = this.fixJSON(jsonStr);
        if (fixedJson) {
            try {
                return JSON.parse(fixedJson);
            } catch (e) {
                console.warn('修复后仍失败，尝试字段提取...');
            }
        }
        
        // 终极方案：逐字段提取
        return this.extractFields(jsonStr);
    }
    
    fixJSON(jsonStr) {
        try {
            // 移除控制字符
            let fixed = jsonStr.replace(/[\x00-\x1F\x7F]/g, '');
            
            // 处理被截断的情况
            if (!fixed.trim().endsWith('}')) {
                // 找到最后一个完整的 "key": "value" 对
                // 倒序查找 "..." 来定位
                let pos = fixed.length;
                let openQuotes = 0;
                let inString = false;
                
                // 从末尾向前扫描，找到字符串闭合的位置
                for (let i = fixed.length - 1; i >= 0; i--) {
                    const char = fixed[i];
                    if (char === '"' && (i === 0 || fixed[i-1] !== '\\')) {
                        if (!inString) {
                            // 进入字符串
                            inString = true;
                            openQuotes++;
                        } else {
                            // 退出字符串
                            openQuotes--;
                            if (openQuotes === 0) {
                                pos = i + 1;
                                break;
                            }
                        }
                    }
                }
                
                if (pos < fixed.length) {
                    // 截取到最后一个完整字符串结束的位置
                    fixed = fixed.substring(0, pos);
                    // 闭合JSON
                    fixed = fixed.replace(/,\s*$/, '') + '"}';
                }
            }
            
            return fixed;
        } catch (e) {
            return null;
        }
    }
    
    extractFields(jsonStr) {
        // 终极方案：逐字段正则提取
        const result = {};
        const fields = [
            'id', 'name', 'age', 'gender', 'ethnicity', 'occupation', 'marriage',
            'diagnosis_wm', 'diagnosis_tcm', 'medical_history', 'smoking_drinking',
            'allergy', 'surgery', 'medications', 'ct_report', 'fibrosis_location',
            'gastroscopy', 'biopsy', 'lung_function', 'total_protein', 'albumin',
            'prealbumin', 'rbc', 'hemoglobin', 'pao2', 'paco2', 'sao2', 'physical_exam'
        ];
        
        for (const field of fields) {
            // 匹配 "field": "value" 或 "field": "value"
            // 使用更灵活的正则，容忍引号内包含各种字符
            const patterns = [
                new RegExp(`"${field}"\\s*:\\s*"([^"]*)"`, 'i'),  // 标准JSON
                new RegExp(`"${field}"\\s*:\\s*'([^']*)'`, 'i'),  // 单引号
                new RegExp(`${field}\\s*:\\s*([^,}\\]\\s]*)`, 'i') // 无引号
            ];
            
            for (const pattern of patterns) {
                const match = jsonStr.match(pattern);
                if (match && match[1] && match[1].trim()) {
                    // 清理提取的值
                    let value = match[1].trim();
                    // 移除末尾可能的逗号或多余字符
                    value = value.replace(/[,，\s]+$/, '');
                    result[field] = value;
                    break;
                }
            }
        }
        
        console.log('字段提取结果：', Object.keys(result).length, '个字段');
        return result;
    }
    
    fillForm(data) {
        this.fieldId.value = data.id || '';
        this.fieldName.value = data.name || '';
        this.fieldAge.value = data.age || '';
        
        if (data.gender) {
            const genderLower = data.gender.toLowerCase();
            if (genderLower.includes('男')) this.fieldGender.value = '男';
            else if (genderLower.includes('女')) this.fieldGender.value = '女';
        }
        
        this.fieldEthnicity.value = data.ethnicity || '';
        this.fieldOccupation.value = data.occupation || '';
        
        if (data.marriage) {
            const marriageLower = data.marriage.toLowerCase();
            if (marriageLower.includes('已婚')) this.fieldMarriage.value = '已婚';
            else if (marriageLower.includes('未婚')) this.fieldMarriage.value = '未婚';
            else if (marriageLower.includes('离异')) this.fieldMarriage.value = '离异';
            else if (marriageLower.includes('丧偶')) this.fieldMarriage.value = '丧偶';
        }
        
        this.fieldDiagnosisWm.value = data.diagnosis_wm || data.diagnosisWm || '';
        this.fieldDiagnosisTcm.value = data.diagnosis_tcm || data.diagnosisTcm || '';
        this.fieldMedicalHistory.value = data.medical_history || '';
        this.fieldSmokingDrinking.value = data.smoking_drinking || '';
        this.fieldAllergy.value = data.allergy || '';
        this.fieldSurgery.value = data.surgery || '';
        
        // 环境接触史（多选）- 从完整文本中提取
        if (data.environment) {
            this.checkboxEnvironments.forEach(cb => {
                const envLower = data.environment.toLowerCase();
                cb.checked = envLower.includes(cb.value.toLowerCase());
            });
        }
        
        // 饮食习惯（多选）- 从完整文本中提取
        if (data.diet) {
            this.checkboxDiets.forEach(cb => {
                const dietLower = data.diet.toLowerCase();
                cb.checked = dietLower.includes(cb.value.toLowerCase());
            });
        }
        
        // 幽门螺旋杆菌
        if (data.hp) {
            const hpLower = data.hp.toLowerCase();
            if (hpLower.includes('阳性')) this.fieldHp.value = '阳性';
            else if (hpLower.includes('阴性')) this.fieldHp.value = '阴性';
        }
        this.fieldMedications.value = data.medications || '';
        
        // 检查报告
        this.fieldCtReport.value = data.ct_report || data.ctReport || '';
        this.fieldFibrosisLocation.value = data.fibrosis_location || data.fibrosisLocation || '';
        this.fieldGastroscopy.value = data.gastroscopy || '';
        this.fieldBiopsy.value = data.biopsy || '';
        this.fieldLungFunction.value = data.lung_function || data.lungFunction || '';
        
        // 营养指标
        this.fieldTotalProtein.value = data.total_protein || data.totalProtein || '';
        this.fieldAlbumin.value = data.albumin || '';
        this.fieldPrealbumin.value = data.prealbumin || '';
        this.fieldRbc.value = data.rbc || '';
        this.fieldHemoglobin.value = data.hemoglobin || '';
        
        // 血气分析
        this.fieldPao2.value = data.pao2 || data.pao_2 || '';
        this.fieldPaco2.value = data.paco2 || data.paco_2 || '';
        this.fieldSao2.value = data.sao2 || data.sao_2 || '';
        
        // 氧合评估
        this.calculateOxygenConcentration();
        this.calculateOxygenIndex();
        
        // 体格检查
        this.fieldPhysicalExam.value = data.physical_exam || data.physicalExam || '';
        
        // 中医问诊
        this.fieldTcmTotal.value = data.tcm_total || data.tcmTotal || '0';
    }
}

document.addEventListener('DOMContentLoaded', () => {
    window.app = new ImageAnalyzerApp();
});
